import streamlit as st
import os
import io
import re
import uuid
import sys
import zipfile
import subprocess

# Configure Streamlit Page
st.set_page_config(
    page_title="ABLBL JD Transformer AI Agent",
    page_icon="✨",
    layout="centered"
)

# Custom Premium Styles
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&display=swap');
        
        /* Font Styles */
        html, body, [class*="css"] {
            font-family: 'Outfit', sans-serif;
        }
        
        /* Main Container Styling */
        .main-title {
            text-align: center;
            font-size: 2.75rem;
            font-weight: 700;
            background: linear-gradient(135deg, #ef4444 0%, #f97316 50%, #4581b5 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin-bottom: 0.25rem;
        }
        .subtitle {
            text-align: center;
            font-size: 1.05rem;
            color: #94a3b8;
            margin-bottom: 2rem;
            line-height: 1.5;
        }
        .agent-badge {
            display: inline-block;
            background: rgba(34, 197, 94, 0.1);
            border: 1px solid rgba(34, 197, 94, 0.2);
            color: #4ade80;
            padding: 0.35rem 0.75rem;
            border-radius: 9999px;
            font-size: 0.85rem;
            font-weight: 600;
            margin-bottom: 1rem;
        }
        .badge-container {
            text-align: center;
        }
        
        /* Glassmorphic Cards */
        .glass-card {
            background: rgba(13, 25, 41, 0.7);
            backdrop-filter: blur(20px);
            -webkit-backdrop-filter: blur(20px);
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 20px;
            padding: 2rem;
            margin-bottom: 2rem;
            box-shadow: 0 20px 40px rgba(0,0,0,0.5);
        }
        
        /* Styled Footer */
        footer {
            text-align: center;
            padding-top: 3rem;
            color: #64748b;
            font-size: 0.85rem;
        }
    </style>
""", unsafe_allow_html=True)

# Try importing dependencies and helpers from app.py
try:
    from app import (
        build_docx, 
        call_groq, 
        clean_data_recursively, 
        extract_text_from_bytes, 
        GROQ_API_KEY,
        generate_all_icons
    )
    # Automatically initialize brand graphics
    generate_all_icons()
except Exception as e:
    st.error(f"Error loading backend logic: {e}")
    st.stop()

# Helper for cross-platform DOCX-to-PDF Conversion
def convert_docx_to_pdf(docx_path, pdf_path):
    if sys.platform.startswith("win"):
        # Local Windows (runs Flask / local dev)
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            import docx2pdf
            docx2pdf.convert(docx_path, pdf_path)
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
    else:
        # Streamlit Cloud Linux (using headless LibreOffice)
        outdir = os.path.dirname(pdf_path) or "."
        cmd = [
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", outdir,
            docx_path
        ]
        # Run conversion in subprocess
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

# UI Elements
st.markdown('<div class="badge-container"><div class="agent-badge"><span style="display:inline-block; width:8px; height:8px; border-radius:50%; background:#22c55e; margin-right:6px;"></span>Agent: ABLBL JD Transformer AI Agent</div></div>', unsafe_allow_html=True)
st.markdown('<h1 class="main-title">ABLBL JD Transformer AI Agent</h1>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Transform legacy JDs into official, full-bleed ABLBL Role Playbook PDFs instantly.</p>', unsafe_allow_html=True)

# Settings Panel (Expander)
with st.expander("⚙️ Settings (Groq API Key Override)"):
    api_key_override = st.text_input(
        "Groq API Key (Leave blank to use system default)",
        type="password",
        placeholder="Enter your gsk_... key"
    ).strip()

# Main Interaction Area
uploaded_files = st.file_uploader(
    "Drag & Drop Legacy JDs",
    type=["pdf", "doc", "docx", "txt"],
    accept_multiple_files=True
)

misc_text = ""

if st.button("Generate Playbooks 🚀", use_container_width=True):
    if not uploaded_files:
        st.warning("Please upload at least one legacy JD file first.")
    else:
        # Use appropriate API key
        use_api_key = api_key_override if api_key_override else GROQ_API_KEY
        
        # Temp directories
        os.makedirs("temp_uploads", exist_ok=True)
        os.makedirs("temp_outputs", exist_ok=True)
        
        processed_files = []
        progress_bar = st.progress(0)
        
        for idx, file in enumerate(uploaded_files):
            st.write(f"Processing **{file.name}**...")
            file_bytes = file.read()
            
            try:
                # Text Extraction
                payload = extract_text_from_bytes(file_bytes, file.name)
                if not payload.strip():
                    st.error(f"Could not extract text from {file.name}.")
                    continue
                
                # Fetch LLM response
                raw_data = call_groq(payload, api_key=use_api_key)
                data = clean_data_recursively(raw_data)
                
                # Self-healing location fallback
                loc = data.get("location", "").strip()
                if not loc or loc.lower() in ["none", "null", "tbd", "any", ""]:
                    match = re.search(r'(?:location|place of work)\s*[:\-–—]?\s*([a-zA-Z\s]+)', payload, re.IGNORECASE)
                    if match:
                        data["location"] = match.group(1).strip().split('\n')[0].strip()
                    else:
                        data["location"] = "Mumbai"
                
                # Generate unique paths
                safe_title = re.sub(r"[^\w\s-]", "", data.get("job_title", "Role")).replace(" ", "_")
                unique_suffix = uuid.uuid4().hex[:4]
                
                docx_path = os.path.join("temp_outputs", f"{safe_title}_Role_Playbook_{unique_suffix}.docx")
                pdf_path = docx_path.replace(".docx", ".pdf")
                
                # Build DOCX
                build_docx(data, docx_path)
                
                # Compile to PDF
                try:
                    convert_docx_to_pdf(docx_path, pdf_path)
                    processed_files.append((f"{safe_title}_Role_Playbook.pdf", pdf_path))
                except Exception as conv_err:
                    st.warning(f"PDF compilation failed for {file.name}, providing DOCX instead. Error: {conv_err}")
                    processed_files.append((f"{safe_title}_Role_Playbook.docx", docx_path))
                
            except Exception as e:
                st.error(f"Failed to process {file.name}: {e}")
            
            # Progress bar update
            progress_bar.progress((idx + 1) / len(uploaded_files))
            
        # Download Phase
        if processed_files:
            st.success("Transformation complete! Download your files below:")
            
            if len(processed_files) == 1:
                filename, filepath = processed_files[0]
                with open(filepath, "rb") as f:
                    st.download_button(
                        label=f"Download {filename} 📥",
                        data=f.read(),
                        file_name=filename,
                        mime="application/pdf" if filename.endswith(".pdf") else "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
            else:
                zip_path = os.path.join("temp_outputs", "ABLBL_Role_Playbooks.zip")
                with zipfile.ZipFile(zip_path, "w") as zipf:
                    for filename, filepath in processed_files:
                        zipf.write(filepath, arcname=filename)
                        
                with open(zip_path, "rb") as f:
                    st.download_button(
                        label="Download All Playbooks (.zip) 📥",
                        data=f.read(),
                        file_name="ABLBL_Role_Playbooks.zip",
                        mime="application/zip",
                        use_container_width=True
                    )

# Branded Footer
st.markdown("<footer>ABLBL JD Transformer AI Agent • Aditya Birla Lifestyle Brands Limited</footer>", unsafe_allow_html=True)

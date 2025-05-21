import streamlit as st
from metadata_utils import (
    extract_image_metadata,
    get_file_type,
    format_metadata,
    extract_pdf_metadata,
    update_pdf_metadata,
    extract_video_metadata,
    update_image_metadata,
)
from office_metadata import (
    extract_pptx_metadata,
    extract_docx_metadata,
    update_pptx_metadata,
    update_docx_metadata
)
import json
import os
import tempfile
import base64
from PIL import Image

# App Configuration
st.set_page_config(page_title="TagIt", layout="wide")
st.title("📁 TagIt")

# Sidebar with file upload and instructions
with st.sidebar:
    st.header("Upload File")
    uploaded_file = st.file_uploader(
        "Choose a file",
        type=["jpg", "jpeg", "png", "pdf", "docx", "pptx", "mp4", "mov", "avi"],
        help="Supported formats: Images, PDF, Office Documents, Videos"
    )
    
    st.markdown("""
    ### Features:
    - View metadata for various file types
    - Edit metadata for supported formats
    - Download modified files
    - Export metadata as JSON
    """)

if uploaded_file:
    file_name = uploaded_file.name.lower()
    
    # Determine file type based on extension (more reliable for office docs)
    if file_name.endswith(('.jpg', '.jpeg', '.png')):
        file_type = "image"
    elif file_name.endswith('.pdf'):
        file_type = "pdf"
    elif file_name.endswith('.docx'):
        file_type = "docx"
    elif file_name.endswith('.pptx'):
        file_type = "pptx"
    elif file_name.endswith(('.mp4', '.mov', '.avi')):
        file_type = "video"
    else:
        file_type = "unsupported"
    
    # Initialize session state
    if 'current_metadata' not in st.session_state:
        st.session_state.current_metadata = None
    if 'modified_path' not in st.session_state:
        st.session_state.modified_path = None
    
    # Main content layout
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.subheader("File Preview")
        if file_type == "image":
            st.image(uploaded_file, caption=f"Uploaded Image: {uploaded_file.name}", use_column_width=True)
        elif file_type == "pdf":
            try:
                st.markdown(f"""
                <iframe src="data:application/pdf;base64,{base64.b64encode(uploaded_file.getvalue()).decode('utf-8')}" 
                        width="100%" 
                        height="500" 
                        type="application/pdf">
                </iframe>
                """, unsafe_allow_html=True)
            except Exception as e:
                st.warning("Couldn't display PDF preview in this environment")
                st.download_button(
                    label="⬇ Download PDF to view",
                    data=uploaded_file,
                    file_name=uploaded_file.name,
                    mime="application/pdf"
                )
        elif file_type == "video":
            st.video(uploaded_file)
        elif file_type in ["docx", "pptx"]:
            icon = "📝" if file_type == "docx" else "📊"
            st.markdown(f"""
            <div style="text-align: center;">
                <span style="font-size: 100px;">{icon}</span>
                <p style="font-size: 24px;">{uploaded_file.name}</p>
            </div>
            """, unsafe_allow_html=True)
    
    with col2:
        st.subheader("Metadata Operations")
        
        # Image processing
        if file_type == "image":
            metadata = extract_image_metadata(uploaded_file)
            formatted_metadata = format_metadata(metadata)
            st.session_state.current_metadata = formatted_metadata

            with st.expander("📸 Image Metadata", expanded=True):
                st.json(formatted_metadata)

            basic_info = formatted_metadata.get("Basic Info", {})

            with st.expander("✏ Edit Image Metadata"):
                with st.form("image_edit_form"):
                    new_make = st.text_input("Make", basic_info.get("Make", ""))
                    new_model = st.text_input("Model", basic_info.get("Model", ""))
                    new_software = st.text_input("Software", basic_info.get("Software", ""))
                    new_datetime = st.text_input("DateTime", basic_info.get("DateTime", ""))
                    new_datetime_orig = st.text_input("DateTimeOriginal", basic_info.get("DateTimeOriginal", ""))

                    if st.form_submit_button("Update Metadata"):
                        changes = {
                            "Make": new_make,
                            "Model": new_model,
                            "Software": new_software,
                            "DateTime": new_datetime,
                            "DateTimeOriginal": new_datetime_orig,
                        }
                        try:
                            updated_path = update_image_metadata(uploaded_file, changes)

                            with open(updated_path, "rb") as f:
                                new_meta = extract_image_metadata(f)
                                st.session_state.current_metadata = format_metadata(new_meta)
                                st.session_state.modified_path = updated_path

                            st.success("✅ Metadata successfully updated!")
                            st.rerun()

                        except Exception as e:
                            st.error(f"❌ Update failed: {str(e)}")

        # PDF processing
        elif file_type == "pdf":
            st.session_state.current_metadata = extract_pdf_metadata(uploaded_file)

            with st.expander("📄 PDF Metadata", expanded=True):
                st.json(st.session_state.current_metadata)

            with st.expander("✏ Edit Metadata"):
                with st.form("pdf_edit_form"):
                    current_title = st.session_state.current_metadata.get("Title", "")
                    current_author = st.session_state.current_metadata.get("Author", "")

                    new_title = st.text_input("Title", current_title)
                    new_author = st.text_input("Author", current_author)

                    if st.form_submit_button("Update Metadata"):
                        try:
                            changes = {
                                "Title": new_title,
                                "Author": new_author,
                            }

                            updated_path = update_pdf_metadata(uploaded_file, changes)

                            if not os.path.exists(updated_path):
                                raise Exception("Modified file was not created")

                            st.session_state.current_metadata = extract_pdf_metadata(updated_path)
                            st.session_state.modified_path = updated_path
                            
                            st.success("✅ Metadata successfully updated!")
                            st.rerun()

                        except Exception as e:
                            st.error(f"❌ Update failed: {str(e)}")

        # Video processing
        elif file_type == "video":
            st.session_state.current_metadata = extract_video_metadata(uploaded_file)
            
            with st.expander("🎥 Video Metadata", expanded=True):
                st.json(st.session_state.current_metadata)
            
            st.info("Video metadata editing is currently not supported")

        # Word documents (DOCX)
        elif file_type == "docx":
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
                uploaded_file.seek(0)
                tmp_file.write(uploaded_file.read())
                tmp_path = tmp_file.name
            
            try:
                st.session_state.current_metadata = extract_docx_metadata(tmp_path)
                
                with st.expander("📝 Word Document Metadata", expanded=True):
                    st.json(st.session_state.current_metadata['CoreProperties'])
                
                with st.expander("✏ Edit Metadata"):
                    with st.form("docx_form"):
                        title = st.text_input("Title", st.session_state.current_metadata['CoreProperties'].get('title', ''))
                        author = st.text_input("Author", st.session_state.current_metadata['CoreProperties'].get('author', ''))
                        subject = st.text_input("Subject", st.session_state.current_metadata['CoreProperties'].get('subject', ''))
                        keywords = st.text_input("Keywords", st.session_state.current_metadata['CoreProperties'].get('keywords', ''))
                        
                        if st.form_submit_button("Update Metadata"):
                            changes = {
                                'title': title,
                                'author': author,
                                'subject': subject,
                                'keywords': keywords
                            }
                            result = update_docx_metadata(tmp_path, changes)
                            
                            if isinstance(result, str) and os.path.exists(result):
                                st.session_state.modified_path = result
                                st.session_state.current_metadata = extract_docx_metadata(result)
                                st.success("✅ Metadata successfully updated!")
                                st.rerun()
                            else:
                                st.error(f"❌ {result.get('Error', 'Update failed')}")
            
            finally:
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)

        # PowerPoint files (PPTX)
        elif file_type == "pptx":
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_file:
                uploaded_file.seek(0)
                tmp_file.write(uploaded_file.read())
                tmp_path = tmp_file.name
            
            try:
                st.session_state.current_metadata = extract_pptx_metadata(tmp_path)
                
                with st.expander("📊 PowerPoint Metadata", expanded=True):
                    st.json(st.session_state.current_metadata['CoreProperties'])
                
                with st.expander("✏ Edit Metadata"):
                    with st.form("pptx_form"):
                        title = st.text_input("Title", st.session_state.current_metadata['CoreProperties'].get('title', ''))
                        author = st.text_input("Author", st.session_state.current_metadata['CoreProperties'].get('author', ''))
                        subject = st.text_input("Subject", st.session_state.current_metadata['CoreProperties'].get('subject', ''))
                        keywords = st.text_input("Keywords", st.session_state.current_metadata['CoreProperties'].get('keywords', ''))
                        
                        if st.form_submit_button("Update Metadata"):
                            changes = {
                                'title': title,
                                'author': author,
                                'subject': subject,
                                'keywords': keywords
                            }
                            result = update_pptx_metadata(tmp_path, changes)
                            
                            if isinstance(result, str) and os.path.exists(result):
                                st.session_state.modified_path = result
                                st.session_state.current_metadata = extract_pptx_metadata(result)
                                st.success("✅ Metadata successfully updated!")
                                st.rerun()
                            else:
                                st.error(f"❌ {result.get('Error', 'Update failed')}")
            
            finally:
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)
        
        else:
            st.error("Unsupported file type!")

    # Download buttons section
    st.divider()
    st.subheader("Download Options")
    
    col_d1, col_d2 = st.columns(2)
    
    with col_d1:
        if st.session_state.current_metadata:
            st.download_button(
                label="💾 Download Metadata as JSON",
                data=json.dumps(st.session_state.current_metadata, indent=2),
                file_name=f"{uploaded_file.name}_metadata.json",
                mime="application/json",
            )
    
    with col_d2:
        if st.session_state.modified_path and os.path.exists(st.session_state.modified_path):
            with open(st.session_state.modified_path, "rb") as f:
                st.download_button(
                    label="⬇ Download Modified File",
                    data=f,
                    file_name=f"modified_{uploaded_file.name}",
                    mime="application/octet-stream"
                )

else:
    st.info("👈 Please upload a file to view and edit its metadata")
    st.markdown("""
    ### Supported File Types:
    - **Images**: JPG, PNG (view and edit metadata)
    - **Documents**: 
      - PDF (view and edit metadata)
      - Word DOCX (view and edit metadata)
      - PowerPoint PPTX (view and edit metadata)
    - **Videos**: MP4, MOV, AVI (view metadata only)
    """)

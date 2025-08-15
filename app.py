import streamlit as st
import pandas as pd
import io
import os
from pdf_processor import PDFProcessor
from excel_converter import ExcelConverter
from data_sanitizer import DataSanitizer

# Configure page
st.set_page_config(
    page_title="PDF to Excel Converter",
    page_icon="📄",
    layout="wide"
)

# Initialize processors
pdf_processor = PDFProcessor()
excel_converter = ExcelConverter()
data_sanitizer = DataSanitizer()

# Security notice
st.info("🔒 **Security Notice**: All files are processed in memory only. No data is stored on servers. Files are automatically deleted after processing.")

# Main title
st.title("📄 PDF to Excel Converter")
st.markdown("Convert PDF files containing traditional Chinese text to Excel format")

# Sidebar for instructions
with st.sidebar:
    st.header("Instructions")
    st.markdown("""
    1. Accept security agreement
    2. Upload a PDF file (max 50MB)
    3. Choose extraction method
    4. Click 'Convert to Excel' button
    5. Download the converted Excel file
    
    **Supported Features:**
    - Traditional Chinese character support
    - Text extraction from PDF pages
    - Table detection and extraction
    - Multiple output formats
    """)
    
    st.header("🔒 Security Features")
    st.markdown("""
    **Data Protection:**
    - In-memory processing only
    - No server-side file storage
    - Automatic data cleanup
    - File size validation
    - Secure file type checking
    - Memory usage optimization
    
    **Privacy Measures:**
    - No data logging
    - No metadata retention
    - Session-based processing
    - Immediate cleanup after download
    """)
    
    st.header("⚠️ Security Recommendations")
    st.markdown("""
    **For Classified Data:**
    - Use offline/local deployment
    - Review organizational policies
    - Consider air-gapped systems
    - Verify network security requirements
    """)

# File upload section
st.header("1. Upload PDF File")

# Security agreement checkbox
security_agreement = st.checkbox(
    "I understand that this file will be processed securely in memory and automatically deleted after conversion",
    value=False,
    help="Required: Confirm you understand the security measures before uploading"
)

# File size limit warning
st.warning("⚠️ **File Size Limit**: Maximum file size is 50MB for optimal performance and security.")

uploaded_file = None
if security_agreement:
    uploaded_file = st.file_uploader(
        "Choose a PDF file",
        type=['pdf'],
        help="Upload a PDF file containing traditional Chinese text"
    )
    
    # Validate file size
    if uploaded_file is not None:
        file_size_mb = len(uploaded_file.getvalue()) / (1024 * 1024)
        if file_size_mb > 50:
            st.error(f"❌ File too large: {file_size_mb:.1f}MB. Please use a file smaller than 50MB.")
            uploaded_file = None
else:
    st.info("👆 Please confirm the security agreement before uploading files")

if uploaded_file is not None:
    # Display file information
    st.success(f"✅ File uploaded: {uploaded_file.name}")
    st.info(f"File size: {len(uploaded_file.getvalue())} bytes")
    
    # Extraction method selection
    st.header("2. Choose Extraction Method")
    extraction_method = st.radio(
        "Select how to extract data from the PDF:",
        ["Text Extraction", "Table Detection", "Both Text and Tables"],
        help="Choose the appropriate method based on your PDF content"
    )
    
    # Processing options
    st.header("3. Processing Options")
    col1, col2 = st.columns(2)
    
    with col1:
        preserve_formatting = st.checkbox("Preserve text formatting", value=True)
        split_by_pages = st.checkbox("Split data by pages", value=False)
        remove_empty_rows = st.checkbox("Remove empty rows", value=True)
    
    with col2:
        merge_text_blocks = st.checkbox("Merge text blocks", value=False)
        # Security options
        sanitize_data = st.checkbox("🔒 Basic data sanitization", value=False, 
                                  help="Remove common sensitive patterns like emails, phone numbers")
        redact_numbers = st.checkbox("🔒 Redact numeric patterns", value=False,
                                   help="Replace sequences of digits with [REDACTED] for privacy")
    
    # Convert button
    st.header("4. Convert to Excel")
    if st.button("🔄 Convert to Excel", type="primary"):
        try:
            with st.spinner("Processing PDF file..."):
                # Read PDF file
                pdf_bytes = uploaded_file.getvalue()
                
                # Process based on selected method
                if extraction_method == "Text Extraction":
                    st.info("Extracting text from PDF...")
                    extracted_data = pdf_processor.extract_text(
                        pdf_bytes, 
                        preserve_formatting=preserve_formatting,
                        split_by_pages=split_by_pages,
                        merge_text_blocks=merge_text_blocks
                    )
                    
                    # Apply data sanitization if requested
                    if sanitize_data or redact_numbers:
                        st.info("🔒 Applying data sanitization...")
                        extracted_data = data_sanitizer.sanitize_extracted_data(
                            extracted_data, sanitize_data, redact_numbers
                        )
                    
                elif extraction_method == "Table Detection":
                    st.info("Detecting and extracting tables from PDF...")
                    extracted_data = pdf_processor.extract_tables(
                        pdf_bytes,
                        remove_empty_rows=remove_empty_rows
                    )
                    
                    # Apply data sanitization if requested
                    if sanitize_data or redact_numbers:
                        st.info("🔒 Applying data sanitization...")
                        extracted_data = data_sanitizer.sanitize_extracted_data(
                            extracted_data, sanitize_data, redact_numbers
                        )
                    
                else:  # Both Text and Tables
                    st.info("Extracting both text and tables from PDF...")
                    text_data = pdf_processor.extract_text(
                        pdf_bytes,
                        preserve_formatting=preserve_formatting,
                        split_by_pages=split_by_pages,
                        merge_text_blocks=merge_text_blocks
                    )
                    table_data = pdf_processor.extract_tables(
                        pdf_bytes,
                        remove_empty_rows=remove_empty_rows
                    )
                    extracted_data = {"text": text_data, "tables": table_data}
                    
                    # Apply data sanitization if requested
                    if sanitize_data or redact_numbers:
                        st.info("🔒 Applying data sanitization...")
                        extracted_data = data_sanitizer.sanitize_extracted_data(
                            extracted_data, sanitize_data, redact_numbers
                        )
                
                # Convert to Excel
                st.info("Converting to Excel format...")
                excel_file = excel_converter.convert_to_excel(
                    extracted_data, 
                    extraction_method,
                    uploaded_file.name
                )
                
                # Success message and preview
                st.success("✅ Conversion completed successfully!")
                
                # Show preview of extracted data
                st.header("5. Preview Extracted Data")
                if extraction_method == "Text Extraction":
                    if isinstance(extracted_data, list):
                        for i, page_data in enumerate(extracted_data):
                            with st.expander(f"Page {i+1} Preview"):
                                preview_text = str(page_data)
                                st.text_area("", value=preview_text[:500] + "..." if len(preview_text) > 500 else preview_text, height=100)
                    else:
                        preview_text = str(extracted_data)
                        st.text_area("Extracted Text Preview", value=preview_text[:1000] + "..." if len(preview_text) > 1000 else preview_text, height=200)
                
                elif extraction_method == "Table Detection":
                    if extracted_data and isinstance(extracted_data, list):
                        for i, table in enumerate(extracted_data):
                            with st.expander(f"Table {i+1} Preview"):
                                st.dataframe(table.head() if len(table) > 5 else table)
                    else:
                        st.warning("No tables detected in the PDF.")
                
                else:  # Both
                    if isinstance(extracted_data, dict):
                        if extracted_data.get("text"):
                            st.subheader("Text Data Preview")
                            text_preview = extracted_data["text"]
                            if isinstance(text_preview, list):
                                text_preview = "\n\n".join([str(x) for x in text_preview])
                            else:
                                text_preview = str(text_preview)
                            st.text_area("", value=text_preview[:1000] + "..." if len(text_preview) > 1000 else text_preview, height=150)
                        
                        if extracted_data.get("tables") and isinstance(extracted_data["tables"], list):
                            st.subheader("Table Data Preview")
                            for i, table in enumerate(extracted_data["tables"]):
                                with st.expander(f"Table {i+1} Preview"):
                                    st.dataframe(table.head() if len(table) > 5 else table)
                
                # Download section
                st.header("6. Download Excel File")
                filename = f"{os.path.splitext(uploaded_file.name)[0]}_converted.xlsx"
                
                st.download_button(
                    label="📥 Download Excel File",
                    data=excel_file.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.success(f"Click the button above to download {filename}")
                
                # Security information
                if sanitize_data or redact_numbers:
                    st.success("🔒 **Data Sanitization Applied**: Sensitive patterns have been redacted from the output.")
                
                st.info("🔒 **Security**: Original file data has been cleared from memory after processing.")
                
        except Exception as e:
            st.error(f"❌ Error during conversion: {str(e)}")
            st.error("Please check your PDF file and try again. Make sure the file is not corrupted and contains readable text.")
            
            # Security cleanup on error
            st.info("🔒 **Security**: File data has been cleared from memory due to processing error.")
            
            # Detailed error information in expander
            with st.expander("Technical Error Details"):
                st.code(str(e))

else:
    st.info("👆 Please upload a PDF file to begin conversion")
    
    # Show example of supported content
    st.header("Example Supported Content")
    st.markdown("""
    This converter supports:
    - **Text PDFs**: Documents with traditional Chinese text that can be extracted
    - **Table PDFs**: Documents containing tables with Chinese characters
    - **Mixed Content**: PDFs with both text and tabular data
    
    **Note**: Scanned PDFs (image-based) may require OCR preprocessing for best results.
    """)

# Footer
st.markdown("---")
col1, col2 = st.columns(2)

with col1:
    st.markdown("**Made with ❤️ using Streamlit**")
    st.markdown("Supports Traditional Chinese Characters")

with col2:
    st.markdown("**🔒 Security Features**")
    st.markdown("""
    - In-memory processing only
    - No data storage or logging
    - Optional data sanitization
    - Automatic cleanup after processing
    """)

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; font-size: 0.8em;'>
<strong>Security Notice:</strong> This application processes files entirely in memory. 
No data is stored on servers. For maximum security with classified documents, 
consider running this application locally or on air-gapped systems.
</div>
""", unsafe_allow_html=True)

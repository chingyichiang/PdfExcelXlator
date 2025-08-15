# PDF to Excel Converter

## Overview

This is a Streamlit web application that converts PDF files containing traditional Chinese text to Excel format. The application provides multiple extraction methods including text extraction and table detection, with specialized support for Chinese character encoding and formatting. Users can upload PDF files through a web interface and download the converted Excel files with proper Chinese character support.

## User Preferences

Preferred communication style: Simple, everyday language.

## System Architecture

### Frontend Architecture
- **Framework**: Streamlit-based web application providing a simple file upload interface
- **Layout**: Wide layout configuration with sidebar for instructions and main area for file processing
- **User Flow**: Single-page application with step-by-step process (upload → extract → convert → download)

### Backend Architecture
- **Modular Design**: Three main components separated into distinct modules:
  - `app.py`: Main Streamlit application and user interface
  - `pdf_processor.py`: PDF text and table extraction logic
  - `excel_converter.py`: Excel file generation and formatting
- **Processing Pipeline**: PDF upload → Text/Table extraction → Excel conversion → File download
- **Data Flow**: Bytes-based processing to handle file uploads and downloads in memory

### PDF Processing Strategy
- **Library Choice**: pdfplumber for robust PDF text and table extraction
- **Chinese Character Support**: UTF-8 encoding throughout the pipeline
- **Extraction Methods**: 
  - Text extraction with formatting preservation options
  - Table detection and extraction
  - Combined text and table extraction
- **Text Processing**: Configurable options for formatting preservation, page splitting, and text block merging

### Excel Generation Architecture
- **Library Stack**: pandas for data manipulation + openpyxl for Excel formatting
- **Font Strategy**: Microsoft YaHei font specifically chosen for Chinese character rendering
- **Styling System**: Consistent formatting with headers, borders, and cell alignment
- **Output Format**: In-memory BytesIO objects for efficient file handling

### File Handling Design
- **Memory-Based Processing**: All file operations handled in memory using BytesIO objects
- **No Persistent Storage**: Files are processed and downloaded without server-side storage
- **Upload Constraints**: PDF file type validation with file size reporting

## External Dependencies

### Core Libraries
- **streamlit**: Web application framework for the user interface
- **pandas**: Data manipulation and DataFrame operations for table processing
- **pdfplumber**: PDF text and table extraction with Chinese character support
- **openpyxl**: Excel file creation and advanced formatting capabilities
- **io**: Built-in Python module for in-memory file operations

### Font Dependencies
- **Microsoft YaHei**: System font dependency for proper Chinese character rendering in Excel files

### File Format Support
- **Input**: PDF files with traditional Chinese text
- **Output**: Excel files (.xlsx format) with preserved Chinese characters and formatting
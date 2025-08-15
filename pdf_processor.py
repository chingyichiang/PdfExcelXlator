import pdfplumber
import pandas as pd
import io
from typing import List, Dict, Any, Union
import re

class PDFProcessor:
    """Handles PDF processing and text/table extraction with Chinese character support"""
    
    def __init__(self):
        self.encoding = 'utf-8'
        self.max_file_size = 50 * 1024 * 1024  # 50MB in bytes
        self.allowed_mime_types = ['application/pdf']
    
    def _validate_pdf_security(self, pdf_bytes: bytes) -> bool:
        """
        Validate PDF file for security constraints
        
        Args:
            pdf_bytes: PDF file as bytes
            
        Returns:
            True if file passes security checks
            
        Raises:
            ValueError: If file fails security validation
        """
        # Check file size
        if len(pdf_bytes) > self.max_file_size:
            raise ValueError(f"File too large: {len(pdf_bytes)/1024/1024:.1f}MB. Maximum allowed: {self.max_file_size/1024/1024}MB")
        
        # Check if it's actually a PDF by looking for PDF header
        if not pdf_bytes.startswith(b'%PDF-'):
            raise ValueError("Invalid PDF file: Missing PDF header signature")
        
        # Basic PDF structure validation
        if b'%%EOF' not in pdf_bytes:
            raise ValueError("Invalid PDF file: Missing end-of-file marker")
        
        return True
    
    def _secure_cleanup(self, data=None):
        """
        Securely clear sensitive data from memory
        """
        if data is not None:
            try:
                if hasattr(data, 'clear'):
                    data.clear()
                del data
            except:
                pass  # Best effort cleanup
    
    def extract_text(self, pdf_bytes: bytes, preserve_formatting: bool = True, 
                    split_by_pages: bool = False, merge_text_blocks: bool = False) -> Union[str, List[str]]:
        """
        Extract text from PDF with support for traditional Chinese characters
        
        Args:
            pdf_bytes: PDF file as bytes
            preserve_formatting: Whether to preserve line breaks and spacing
            split_by_pages: Whether to return text split by pages
            merge_text_blocks: Whether to merge separate text blocks
            
        Returns:
            Extracted text as string or list of strings (if split_by_pages=True)
        """
        try:
            # Security validation
            self._validate_pdf_security(pdf_bytes)
            
            extracted_text = []
            pdf_stream = None
            
            try:
                pdf_stream = io.BytesIO(pdf_bytes)
                with pdfplumber.open(pdf_stream) as pdf:
                    for page_num, page in enumerate(pdf.pages):
                        # Extract text from page
                        page_text = page.extract_text()
                    
                    if page_text:
                        # Clean and process text
                        if not preserve_formatting:
                            # Remove excessive whitespace but preserve Chinese characters
                            page_text = re.sub(r'\s+', ' ', page_text).strip()
                        
                        if merge_text_blocks:
                            # Join lines that seem to be part of the same paragraph
                            lines = page_text.split('\n')
                            merged_lines = []
                            current_block = ""
                            
                            for line in lines:
                                line = line.strip()
                                if line:
                                    if current_block and not line[0].isupper() and len(current_block) < 100:
                                        current_block += " " + line
                                    else:
                                        if current_block:
                                            merged_lines.append(current_block)
                                        current_block = line
                                else:
                                    if current_block:
                                        merged_lines.append(current_block)
                                        current_block = ""
                            
                            if current_block:
                                merged_lines.append(current_block)
                            
                            page_text = '\n'.join(merged_lines)
                        
                        extracted_text.append(page_text)
                    else:
                        extracted_text.append("")  # Empty page
                
                if not extracted_text:
                    raise ValueError("No text could be extracted from the PDF")
                
                if split_by_pages:
                    return extracted_text
                else:
                    # Join all pages with page separators
                    return '\n\n--- Page Break ---\n\n'.join(extracted_text)
                    
            finally:
                # Secure cleanup
                self._secure_cleanup(pdf_stream)
                
        except Exception as e:
            # Secure cleanup on error
            self._secure_cleanup()
            raise Exception(f"Failed to extract text from PDF: {str(e)}")
    
    def extract_tables(self, pdf_bytes: bytes, remove_empty_rows: bool = True) -> List[pd.DataFrame]:
        """
        Extract tables from PDF with support for traditional Chinese characters
        
        Args:
            pdf_bytes: PDF file as bytes
            remove_empty_rows: Whether to remove empty rows from tables
            
        Returns:
            List of pandas DataFrames containing extracted tables
        """
        try:
            # Security validation
            self._validate_pdf_security(pdf_bytes)
            
            all_tables = []
            pdf_stream = None
            
            try:
                pdf_stream = io.BytesIO(pdf_bytes)
                with pdfplumber.open(pdf_stream) as pdf:
                    for page_num, page in enumerate(pdf.pages):
                        # Extract tables from page
                        tables = page.extract_tables()
                    
                    for table_num, table in enumerate(tables):
                        if table and len(table) > 0:
                            # Convert to DataFrame
                            header = table[0] if table[0] else [f'Column_{i+1}' for i in range(len(table[1]) if len(table) > 1 else 0)]
                            data = table[1:] if len(table) > 1 else []
                            df = pd.DataFrame(data, columns=header)
                            
                            # Clean the DataFrame
                            if remove_empty_rows:
                                # Remove completely empty rows
                                df = df.dropna(how='all')
                                # Remove rows where all values are empty strings
                                df = df[~(df.astype(str) == '').all(axis=1)]
                            
                            # Clean column names if they exist
                            if df.columns is not None:
                                df.columns = [str(col).strip() if col is not None else f'Column_{i+1}' 
                                            for i, col in enumerate(df.columns)]
                            
                            # Clean cell values
                            for col in df.columns:
                                df[col] = df[col].astype(str)
                                df[col] = df[col].str.strip()
                                # Replace 'None' strings with empty strings
                                df[col] = df[col].replace('None', '')
                            
                            # Add metadata
                            df.attrs['page'] = page_num + 1
                            df.attrs['table'] = table_num + 1
                            
                            if not df.empty:
                                all_tables.append(df)
            
                if not all_tables:
                    # If no tables found, try alternative extraction method
                    all_tables = self._extract_tables_alternative(pdf_bytes)
                
                return all_tables
                
            finally:
                # Secure cleanup
                self._secure_cleanup(pdf_stream)
                
        except Exception as e:
            # Secure cleanup on error
            self._secure_cleanup()
            raise Exception(f"Failed to extract tables from PDF: {str(e)}")
    
    def _extract_tables_alternative(self, pdf_bytes: bytes) -> List[pd.DataFrame]:
        """
        Alternative method to extract structured data when no tables are detected
        """
        try:
            # Try to find structured text that might be tabular
            text_data = self.extract_text(pdf_bytes, preserve_formatting=True, split_by_pages=True)
            tables = []
            
            for page_num, page_text in enumerate(text_data):
                if isinstance(page_text, str):
                    # Look for patterns that might indicate tabular data
                    lines = page_text.split('\n')
                    potential_table_lines = []
                    
                    for line in lines:
                        line = line.strip()
                        # Check if line has multiple columns (simple heuristic)
                        if len(line.split()) >= 2 and any(char in line for char in ['\t', '  ', '|', ',']):
                            potential_table_lines.append(line)
                    
                    if len(potential_table_lines) >= 2:  # At least header + 1 data row
                        # Try to parse as table
                        try:
                            # Split by multiple spaces or tabs
                            rows = []
                            for line in potential_table_lines:
                                if '\t' in line:
                                    row = line.split('\t')
                                elif '|' in line:
                                    row = [cell.strip() for cell in line.split('|') if cell.strip()]
                                else:
                                    # Split by multiple spaces
                                    row = re.split(r'\s{2,}', line)
                                
                                if row and len(row) > 1:
                                    rows.append(row)
                            
                            if rows and len(rows) > 1:
                                # Create DataFrame
                                max_cols = max(len(row) for row in rows)
                                # Pad rows to same length
                                padded_rows = []
                                for row in rows:
                                    padded_row = row + [''] * (max_cols - len(row))
                                    padded_rows.append(padded_row)
                                
                                df = pd.DataFrame(padded_rows[1:], columns=padded_rows[0])
                                df.attrs['page'] = page_num + 1
                                df.attrs['table'] = 1
                                tables.append(df)
                                
                        except:
                            continue  # Skip if parsing fails
            
            return tables
            
        except Exception as e:
            return []  # Return empty list if alternative method fails
    
    def get_pdf_info(self, pdf_bytes: bytes) -> Dict[str, Any]:
        """
        Get basic information about the PDF
        
        Args:
            pdf_bytes: PDF file as bytes
            
        Returns:
            Dictionary with PDF information
        """
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                info = {
                    'num_pages': len(pdf.pages),
                    'metadata': pdf.metadata,
                    'pages_with_text': 0,
                    'pages_with_tables': 0,
                    'total_characters': 0
                }
                
                for page in pdf.pages:
                    # Check for text
                    page_text = page.extract_text()
                    if page_text and page_text.strip():
                        info['pages_with_text'] += 1
                        info['total_characters'] += len(page_text)
                    
                    # Check for tables
                    tables = page.extract_tables()
                    if tables:
                        info['pages_with_tables'] += 1
                
                return info
                
        except Exception as e:
            return {'error': str(e)}

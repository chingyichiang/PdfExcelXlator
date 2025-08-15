import re
import pandas as pd
from typing import Union, List

class DataSanitizer:
    """Handles data sanitization and privacy protection for sensitive information"""
    
    def __init__(self):
        # Common patterns for sensitive data
        self.email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
        self.phone_pattern = re.compile(r'(\+?\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}')
        self.ssn_pattern = re.compile(r'\b\d{3}-\d{2}-\d{4}\b')
        self.credit_card_pattern = re.compile(r'\b\d{4}[-\s]?\d{4}[-\s]?\d{4}[-\s]?\d{4}\b')
        self.id_number_pattern = re.compile(r'\b\d{6,12}\b')  # Generic ID numbers
        self.numeric_sequences = re.compile(r'\b\d{4,}\b')  # 4+ digit sequences
        
        # Chinese ID patterns (if needed for Chinese documents)
        self.chinese_id_pattern = re.compile(r'\b\d{15}|\d{17}[\dXx]\b')
        
    def sanitize_text(self, text: str, basic_sanitization: bool = True, 
                     redact_numbers: bool = False) -> str:
        """
        Sanitize text data by removing or redacting sensitive information
        
        Args:
            text: Text to sanitize
            basic_sanitization: Apply basic sanitization (emails, phones, etc.)
            redact_numbers: Redact numeric sequences
            
        Returns:
            Sanitized text
        """
        if not text or not isinstance(text, str):
            return str(text) if text is not None else ""
        
        sanitized_text = text
        
        if basic_sanitization:
            # Replace emails
            sanitized_text = self.email_pattern.sub('[EMAIL_REDACTED]', sanitized_text)
            
            # Replace phone numbers
            sanitized_text = self.phone_pattern.sub('[PHONE_REDACTED]', sanitized_text)
            
            # Replace SSN
            sanitized_text = self.ssn_pattern.sub('[SSN_REDACTED]', sanitized_text)
            
            # Replace credit card numbers
            sanitized_text = self.credit_card_pattern.sub('[CARD_REDACTED]', sanitized_text)
            
            # Replace Chinese ID numbers
            sanitized_text = self.chinese_id_pattern.sub('[ID_REDACTED]', sanitized_text)
            
            # Replace generic ID numbers (be careful not to redact legitimate data)
            sanitized_text = self.id_number_pattern.sub('[ID_REDACTED]', sanitized_text)
        
        if redact_numbers:
            # Replace numeric sequences (4+ digits)
            sanitized_text = self.numeric_sequences.sub('[NUMBER_REDACTED]', sanitized_text)
        
        return sanitized_text
    
    def sanitize_dataframe(self, df: pd.DataFrame, basic_sanitization: bool = True,
                          redact_numbers: bool = False) -> pd.DataFrame:
        """
        Sanitize DataFrame by applying text sanitization to all string columns
        
        Args:
            df: DataFrame to sanitize
            basic_sanitization: Apply basic sanitization
            redact_numbers: Redact numeric sequences
            
        Returns:
            Sanitized DataFrame
        """
        if df is None or df.empty:
            return df
        
        # Create a copy to avoid modifying original
        sanitized_df = df.copy()
        
        # Apply sanitization to all string columns
        for column in sanitized_df.columns:
            if sanitized_df[column].dtype == 'object':  # String columns
                sanitized_df[column] = sanitized_df[column].apply(
                    lambda x: self.sanitize_text(str(x), basic_sanitization, redact_numbers)
                )
        
        return sanitized_df
    
    def sanitize_extracted_data(self, data: Union[str, List[str], List[pd.DataFrame], dict],
                               basic_sanitization: bool = True, redact_numbers: bool = False):
        """
        Sanitize various types of extracted data
        
        Args:
            data: Extracted data to sanitize
            basic_sanitization: Apply basic sanitization
            redact_numbers: Redact numeric sequences
            
        Returns:
            Sanitized data in the same format
        """
        if isinstance(data, str):
            return self.sanitize_text(data, basic_sanitization, redact_numbers)
        
        elif isinstance(data, list):
            if data and isinstance(data[0], str):
                # List of strings (text pages)
                return [self.sanitize_text(text, basic_sanitization, redact_numbers) 
                       for text in data]
            elif data and isinstance(data[0], pd.DataFrame):
                # List of DataFrames (tables)
                return [self.sanitize_dataframe(df, basic_sanitization, redact_numbers)
                       for df in data]
        
        elif isinstance(data, dict):
            # Dictionary with mixed content
            sanitized_dict = {}
            for key, value in data.items():
                sanitized_dict[key] = self.sanitize_extracted_data(
                    value, basic_sanitization, redact_numbers
                )
            return sanitized_dict
        
        return data  # Return unchanged if unknown type
    
    def get_sanitization_summary(self, original_text: str, sanitized_text: str) -> dict:
        """
        Generate a summary of what was sanitized
        
        Args:
            original_text: Original text
            sanitized_text: Sanitized text
            
        Returns:
            Summary dictionary with counts of redacted items
        """
        summary = {
            'emails_redacted': len(self.email_pattern.findall(original_text)),
            'phones_redacted': len(self.phone_pattern.findall(original_text)),
            'ssn_redacted': len(self.ssn_pattern.findall(original_text)),
            'cards_redacted': len(self.credit_card_pattern.findall(original_text)),
            'ids_redacted': len(self.chinese_id_pattern.findall(original_text)),
            'total_redactions': original_text.count('[') - sanitized_text.count('[')
        }
        
        return summary
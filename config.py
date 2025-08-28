"""
Konfigurasi untuk Sistem Excel Processing
"""

import os
from datetime import datetime

class Config:
    """Konfigurasi dasar aplikasi"""
    
    # Flask Configuration
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'excel_processing_secret_key_2024'
    DEBUG = os.environ.get('FLASK_DEBUG', 'True').lower() == 'true'
    
    # File Upload Configuration
    UPLOAD_FOLDER = 'uploads'
    OUTPUT_FOLDER = 'outputs'
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB
    ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
    
    # Excel Processing Configuration
    DEFAULT_SHEET_NAME = 'Sheet1'
    MAX_ROWS_PREVIEW = 5
    AUTO_COLUMN_WIDTH = True
    MAX_COLUMN_WIDTH = 50
    
    # Output Format Configuration
    OUTPUT_COLUMNS = [
        'PROVID',
        'PROVIDER_NAME',
        'SERVICECODE',
        'SERVICECODE DESCRIPTION',
        'KELAS',
        'RUANG BEDAH (SURGERY)/NON RUANG BEDAH (NON SURGERY)',
        'HELPER',
        'TARIFF',
        'TARIFF DESCRIPTION',
        'QUANTITY',
        'TOTAL BILLED',
        'GIVEN DATE (month, day, year)',
        'HEAMODIALISA/CHEMOTHERAPY/ODC/PHYSIOTHERAPY/RADIOTHERAPY',
        'HEAMODIALISA/CHEMOTHERAPY/ODC/PHYSIOTHERAPY/RADIOTHERAPY DESCRIPTION',
        'ICD_X_DIAGNOSIS_PRIMARY',
        'ICD_X_DESC_PRIMARY',
        'ICD_X_DIAGNOSIS_SECONDARY',
        'ICD_X_DESC_SECONDARY',
        'PHYSICIAN NAME',
        'PHYSICIAN DESCRIPTION (DPJP/IGD/POLICLINIC)',
        'CLIENT NAME',
        'CLIENTS DOB (month, day, year)',
        'CLIENTS SEX',
        'CLIENTS ADDRESS',
        'CLIENTS MEMBER ID',
        'CLIENTS MR NUMBER',
        'CLIENTS INVOICE NUMBER',
        'CLIENTSREGISTER NUMBER',
        'CLIENTS OTHER NUMBER',
        'admission',
        'discharge',
        'LoS'
    ]
    
    # Column Mapping Rules
    COLUMN_MAPPING_RULES = {
        'PROVID': ['provid', 'provider_id', 'id_provider'],
        'PROVIDER_NAME': ['provider', 'provider_name', 'nama_provider'],
        'SERVICECODE': ['servicecode', 'service_code', 'kode_layanan'],
        'SERVICECODE DESCRIPTION': ['description', 'service_description', 'deskripsi_layanan'],
        'KELAS': ['kelas', 'class', 'kelas_layanan'],
        'RUANG BEDAH (SURGERY)/NON RUANG BEDAH (NON SURGERY)': [
            'ruang', 'bedah', 'surgery', 'jenis_ruangan'
        ],
        'HELPER': ['helper', 'asisten', 'assistant'],
        'TARIFF': ['tariff', 'tarif', 'harga', 'biaya'],
        'TARIFF DESCRIPTION': ['tariff_description', 'deskripsi_tarif'],
        'QUANTITY': ['quantity', 'jumlah', 'qty'],
        'TOTAL BILLED': ['total', 'billed', 'total_billed', 'total_tagihan']
    }
    
    # Default Values
    DEFAULT_VALUES = {
        'GIVEN DATE (month, day, year)': lambda: datetime.now().strftime('%m/%d/%Y'),
        'CLIENTS SEX': 'L',
        'admission': lambda: datetime.now().strftime('%m/%d/%Y'),
        'discharge': lambda: datetime.now().strftime('%m/%d/%Y'),
        'LoS': '1',
        'PHYSICIAN DESCRIPTION (DPJP/IGD/POLICLINIC)': 'IGD',
        'KELAS': 'ER',
        'RUANG BEDAH (SURGERY)/NON RUANG BEDAH (NON SURGERY)': 'NON OK'
    }
    
    # Data Cleaning Rules
    CLEANING_RULES = {
        'currency_pattern': r'Rp\s*([\d,]+)',
        'date_pattern': r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})',
        'remove_patterns': [r'^\s+', r'\s+$'],  # Remove leading/trailing whitespace
        'replace_patterns': {
            r'\s+': ' ',  # Multiple spaces to single space
            r'[^\w\s\-\.]': ''  # Remove special characters except word chars, spaces, hyphens, dots
        }
    }
    
    # Validation Rules
    VALIDATION_RULES = {
        'required_columns': ['SERVICECODE', 'SERVICECODE DESCRIPTION'],
        'numeric_columns': ['TARIFF', 'QUANTITY', 'TOTAL BILLED'],
        'date_columns': ['GIVEN DATE (month, day, year)', 'admission', 'discharge', 'CLIENTS DOB (month, day, year)']
    }
    
    # Logging Configuration
    LOG_LEVEL = os.environ.get('LOG_LEVEL', 'INFO')
    LOG_FILE = 'excel_processing.log'
    
    @classmethod
    def get_output_filename(cls, input_filename, suffix='processed'):
        """Generate output filename"""
        name_without_ext = os.path.splitext(input_filename)[0]
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        return f"{suffix}_{name_without_ext}_{timestamp}.xlsx"
    
    @classmethod
    def get_upload_filename(cls, original_filename):
        """Generate secure upload filename"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        safe_name = original_filename.replace(' ', '_').replace('(', '').replace(')', '')
        return f"{timestamp}_{safe_name}"
    
    @classmethod
    def ensure_directories(cls):
        """Ensure required directories exist"""
        for directory in [cls.UPLOAD_FOLDER, cls.OUTPUT_FOLDER]:
            if not os.path.exists(directory):
                os.makedirs(directory)
                print(f"📁 Direktori {directory} dibuat")

# Development Configuration
class DevelopmentConfig(Config):
    DEBUG = True
    LOG_LEVEL = 'DEBUG'

# Production Configuration
class ProductionConfig(Config):
    DEBUG = False
    LOG_LEVEL = 'WARNING'
    SECRET_KEY = os.environ.get('SECRET_KEY') or os.urandom(24)

# Testing Configuration
class TestingConfig(Config):
    TESTING = True
    DEBUG = True
    UPLOAD_FOLDER = 'test_uploads'
    OUTPUT_FOLDER = 'test_outputs'

# Configuration dictionary
config = {
    'development': DevelopmentConfig,
    'production': ProductionConfig,
    'testing': TestingConfig,
    'default': DevelopmentConfig
}

def get_config():
    """Get configuration based on environment"""
    config_name = os.environ.get('FLASK_ENV', 'default')
    return config.get(config_name, config['default'])

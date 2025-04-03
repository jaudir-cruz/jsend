# config.py
from datetime import datetime

# Identificador único para a sessão
SESSION_ID = datetime.now().strftime('%Y%m%d_%H%M%S')

# Configurações do SMTP
SMTP_SERVER = ""
SMTP_PORT = 587
EMAIL_AUTH = ""
EMAIL_PASSWORD = ""

# Token da Hugging Face
HF_TOKEN = ""

# E-mail para notificações
NOTIFICATION_RECIPIENT = ""

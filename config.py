"""
Configuration constants for Tkinter Email Sender
"""

# SMTP Presets
SMTP_PRESETS = {
    "Gmail": {
        "host": "smtp.gmail.com",
        "port": 587,
        "use_tls": True,
        "help_url": "https://support.google.com/accounts/answer/185833",
        "help_text": "Gmail: Use App Password"
    },
    "Custom": {
        "host": "",
        "port": 587,
        "use_tls": True,
        "help_url": None,
        "help_text": ""
    }
}

# CSV Column Aliases
COLUMN_ALIASES = {
    "recipient_email": ["recipient_email", "email", "to"],
    "subject": ["subject"],
    "attachment_filename": ["attachment_filename", "attachment"],
    "body_content": ["body_content", "body", "message"]
}

# Default values
DEFAULT_DELAY_MS = 5000
DEFAULT_RANDOMIZE_PERCENT = 20
DEFAULT_DELAY_PRESETS = {
    "1s": 1000,
    "5s": 5000,
    "10s": 10000
}

# UI Constants
WINDOW_TITLE = "CSV Email Sender (Tkinter)"
WINDOW_WIDTH = 700
WINDOW_HEIGHT = 950
LOG_FONT = ("Courier", 9)

# Progress update interval (ms)
PROGRESS_UPDATE_INTERVAL = 100

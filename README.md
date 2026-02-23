# CSV Email Sender (Tkinter)

A simple Python desktop application for sending bulk emails via SMTP. Upload a CSV/Excel file with recipient data and send personalized emails with attachments.

**Cross-platform**: Works on Windows, macOS, and Linux

## Features

- **Simple SMTP Authentication**: No OAuth setup required - just use your email credentials
- **CSV/Excel Support**: Upload data files with recipient information
- **Auto-Delimiter Detection**: Supports comma, semicolon, or tab-separated CSV files
- **Bulk Sending**: Send to hundreds of recipients with configurable delays
- **Attachments**: Include file attachments with original filenames preserved
- **Customizable**: Set default subject, body, CC/BCC recipients
- **Progress Tracking**: Real-time progress bar and detailed log
- **Gmail Support**: Optimized for Gmail with App Password authentication

## Requirements

- Python 3.8 or higher
- Tkinter (GUI toolkit)
- openpyxl library (for Excel support)

## Installation

### Windows

1. **Install Python** from https://www.python.org/downloads/
   - ✅ Check "Add Python to PATH"
   - ✅ Check "tcl/tk and IDLE" (includes Tkinter)

2. **Run the application**:
   - Double-click `run.bat` OR
   - In Command Prompt: `python main.py`

### macOS

**Option 1: Using Homebrew (Recommended)**
```bash
# Install Python with Tkinter support
brew install python-tk@3.13
brew unlink python@3.13
brew link python-tk@3.13 --force

# Run the application
cd /Users/chimin/Documents/script/csv-email-sender-tk
./run.sh
```

**Option 2: Using Official Python**
1. Install from https://www.python.org/downloads/
2. Tkinter is included by default
3. Run: `python3 main.py`

### Linux (Ubuntu/Debian)

```bash
# Install Python and Tkinter
sudo apt-get update
sudo apt-get install python3 python3-tk python3-venv

# Run the application
cd /path/to/csv-email-sender-tk
./run.sh
```

### Linux (Fedora/RHEL)

```bash
# Install Python and Tkinter
sudo dnf install python3 python3-tkinter python3-venv

# Run the application
cd /path/to/csv-email-sender-tk
./run.sh
```

## SMTP Configuration

### Gmail (Recommended)

1. Enable 2-Factor Authentication on your Google Account
2. Generate an App Password:
   - Go to https://myaccount.google.com/apppasswords
   - Select "Mail" and your device
   - Copy the generated password (16 characters)
3. In the app:
   - Select "Gmail" preset
   - Enter your Gmail address
   - Paste the App Password (NOT your regular password)

**Why use an App Password?**
- More secure than using your regular password
- Required for email apps/SMTP access
- Works even if you change your main password

### Custom SMTP Server

You can use any other SMTP provider with the "Custom" preset:

1. Select "Custom" preset
2. Enter your SMTP server address and port
3. Enter your email and password

Popular alternatives:
- **Yahoo Mail**: `smtp.mail.yahoo.com:587` (requires App Password)
- **Fastmail**: `smtp.fastmail.com:587`
- **iCloud**: `smtp.mail.me.com:587` (requires App Password)
- **Your own domain**: Use your hosting provider's SMTP settings

## CSV File Format

The CSV file should have the following columns:

| Column | Required | Description |
|--------|----------|-------------|
| `recipient_email` | Yes | Primary recipient email address |
| `subject` | No | Email subject (overrides default) |
| `attachment_filename` | No | Path to attachment file for this recipient |
| `body_content` | No | Email body text (overrides default) |

**Accepted column aliases:**
- Recipient: `recipient_email`, `email`, `to`
- Subject: `subject`
- Attachment: `attachment_filename`, `attachment`
- Body: `body_content`, `body`, `message`

### Example CSV

```csv
recipient_email;subject;attachment_filename;body_content
john@example.com;Monthly Newsletter;/path/to/newsletter.pdf;Hello John!
jane@example.com;Special Offer;;Hi Jane, check this out!
bob@example.com;;;Hello Bob!
```

**Note**: The app automatically detects the delimiter (comma, semicolon, or tab).

**File Paths**:
- Windows: Use backslashes `C:\path\to\file.pdf` or forward slashes `C:/path/to/file.pdf`
- macOS/Linux: Use forward slashes `/home/user/file.pdf`

You can download a template from within the app by clicking "Download CSV Template".

## Excel Support

The app also supports Excel files (`.xlsx`, `.xls`) with the same column structure. The first row should contain the headers.

## Usage Guide

### Step 1: Upload Files

1. Click "Browse" next to "Data File" to select your CSV or Excel file
2. (Optional) Click "Browse" next to "Attachment Files" to select files to attach to all emails
3. (Optional) Download a template CSV to see the expected format

### Step 2: SMTP Configuration

1. Select a preset (Gmail recommended, or Custom for other providers)
2. Verify the SMTP server and port settings
3. Enter your email address and password/app password
4. **Important**: For Gmail, use an App Password, NOT your regular password

### Step 3: Configure Sending

1. **Delay**: Set delay between emails in milliseconds
   - Quick presets: 1s (1000ms), 5s (5000ms), 10s (10000ms)
   - Randomize: Adds ±20% variance to avoid detection
2. **Default Subject**: Subject to use if CSV column is empty
3. **Default Body**: Body text to use if CSV column is empty
4. **CC/BCC**: Optional carbon copy recipients

### Step 4: Send Emails

1. Click "Start Sending" to begin
2. Monitor progress with the progress bar and log
3. Click "Abort" to stop sending at any time
4. View detailed status: Sent count, Failed count, Elapsed time

## Tips

- **Use Gmail**: Recommended - most reliable with App Passwords
- **Start Small**: Test with a few recipients before sending to large lists
- **Use App Passwords**: For Gmail, always use App Passwords, not your main password
- **Check Your Limits**: Gmail has daily sending limits (often 500 emails/day)
- **Delays Matter**: Use delays to avoid being flagged as spam
- **Verify Attachments**: Make sure attachment file paths are correct

## Troubleshooting

### "No module named '_tkinter'" error

**macOS (Homebrew Python)**:
```bash
brew install python-tk@3.13
brew unlink python@3.13
brew link python-tk@3.13 --force
```

**Linux**:
```bash
sudo apt-get install python3-tk  # Ubuntu/Debian
sudo dnf install python3-tkinter  # Fedora
```

**Windows**:
Reinstall Python and make sure "tcl/tk and IDLE" is checked.

### "Connection failed" or "Authentication failed" error

**For Gmail**:
- Make sure you're using an App Password, NOT your regular password
- Verify 2FA is enabled on your Google Account
- Check that you entered the app password correctly (16 characters, no spaces)
- Make sure you selected "Mail" when generating the app password

**For Custom SMTP**:
- Verify SMTP server and port settings with your email provider
- Check that your email/password are correct
- Some providers require App Passwords (Yahoo, iCloud)
- Some providers block SMTP from unfamiliar locations

### Emails not being received

- Check your spam/junk folder
- Verify recipient email addresses are correct
- Some providers may block bulk senders
- Try sending to yourself first to test

### "File not found" error for attachments

- Verify attachment file paths are correct
- Use absolute paths
- Check that you have permission to read the files

### Rate limiting or "Too many requests" errors

- Increase the delay between emails (try 10 seconds)
- Enable randomization to vary the timing
- Gmail has limits - consider splitting into multiple sessions

### Attachment filename not preserved

The app uses proven email attachment methods that preserve filenames. However:
- Some email clients may display truncated names in the UI
- The actual downloaded file should have the correct name
- Try opening the email in different clients (Apple Mail, Outlook, Thunderbird)
- Check "Show original" in Gmail to verify the headers are correct

## Security Notes

- **Passwords are not stored**: You need to enter your password each session
- **App Passwords**: Always use App Passwords when available (Gmail, Yahoo, iCloud)
- **Local Only**: This app runs locally and doesn't send data to any external servers
- **Network Security**: Uses TLS/SSL connections for secure email transmission

## Gmail Sending Limits

Be aware of Gmail's daily sending limits:
- **Free Gmail accounts**: Typically 500 emails/day
- **Google Workspace**: May vary based on plan (often 2,000 emails/day)

If you need to send more emails, consider:
- Splitting your campaign across multiple days
- Using multiple Gmail accounts
- Using a dedicated email service provider (SendGrid, Mailgun, etc.)

## Changelog

### Version 1.0 (February 2026)

**Features:**
- SMTP email sending with Gmail and custom server support
- CSV and Excel file import with auto-delimiter detection
- Per-recipient and global attachment support
- Configurable delays with randomization
- CC/BCC recipient support
- Real-time progress tracking
- Cross-platform support (Windows, macOS, Linux)

**Technical Implementation:**
- Uses MIMEMultipart for email composition
- Attachment filenames preserved using proven email standards
- Threaded sending to prevent UI freezing
- Queue-based progress updates
- Proper error handling and reporting

## License

MIT License

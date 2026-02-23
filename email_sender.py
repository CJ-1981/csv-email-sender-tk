"""
SMTP Email Sender
"""
import smtplib
import time
import threading
from email.message import EmailMessage
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
import os
from typing import List, Optional, Callable, Dict, Tuple
from queue import Queue


class SMTPSender:
    """Handle SMTP email sending"""
    
    def __init__(self, host: str, port: str, email: str, password: str, use_tls: bool = True):
        self.host = host
        self.port = int(port) if isinstance(port, str) else port
        self.email = email
        self.password = password
        self.use_tls = use_tls
        self.server = None
        self.is_connected = False
    
    def connect(self) -> Tuple[bool, str]:
        """
        Connect to SMTP server
        
        Returns:
            Tuple of (success, message)
        """
        try:
            if self.use_tls:
                self.server = smtplib.SMTP(self.host, self.port, timeout=30)
                self.server.ehlo()
                if self.server.has_extn('STARTTLS'):
                    self.server.starttls()
                    self.server.ehlo()
            else:
                self.server = smtplib.SMTP_SSL(self.host, self.port, timeout=30)
                self.server.ehlo()
            
            self.server.login(self.email, self.password)
            self.is_connected = True
            return True, "Connected successfully"
        except Exception as e:
            self.is_connected = False
            return False, f"Connection failed: {str(e)}"
    
    def disconnect(self):
        """Disconnect from SMTP server"""
        if self.server:
            try:
                self.server.quit()
            except:
                pass
            finally:
                self.server = None
                self.is_connected = False
    
    def send_email(
        self,
        recipient: str,
        subject: str,
        body: str,
        attachments: Optional[List[str]] = None,
        cc: Optional[str] = None,
        bcc: Optional[str] = None
    ) -> Tuple[bool, str]:
        """
        Send a single email
        
        Args:
            recipient: Primary recipient email
            subject: Email subject
            body: Email body text
            attachments: List of file paths to attach
            cc: CC recipients (comma-separated)
            bcc: BCC recipients (comma-separated)
            
        Returns:
            Tuple of (success, message)
        """
        if not self.is_connected:
            return False, "Not connected to SMTP server"
        
        try:
            # Use EmailMessage for modern, reliable attachment handling
            msg = EmailMessage()
            msg.set_content(body)
            
            msg['Subject'] = subject
            msg['From'] = self.email
            msg['To'] = recipient
            
            if cc:
                msg['Cc'] = cc
            
            # Build recipient list
            recipients = [recipient]
            if cc:
                recipients.extend([c.strip() for c in cc.split(',') if c.strip()])
            if bcc:
                recipients.extend([b.strip() for b in bcc.split(',') if b.strip()])
            
            # Add attachments
            if attachments:
                for attachment_path in attachments:
                    if attachment_path and os.path.exists(attachment_path):
                        self._add_attachment(msg, attachment_path)
            
            # Send email
            self.server.send_message(msg, to_addrs=recipients)
            return True, "Sent successfully"
            
        except Exception as e:
            return False, f"Failed: {str(e)}"
    
    def _add_attachment(self, msg: EmailMessage, file_path: str):
        """
        Add attachment with proper filename preservation.
        Uses EmailMessage.add_attachment() which handles encoding properly.
        """
        try:
            # Guess the content type
            import mimetypes
            content_type, encoding = mimetypes.guess_type(file_path)
            if content_type is None or encoding is not None:
                content_type = 'application/octet-stream'
            
            main_type, sub_type = content_type.split('/', 1)
            
            # Read the file
            with open(file_path, 'rb') as fp:
                data = fp.read()
            
            # Get the original filename
            filename = os.path.basename(file_path)
            
            # Add attachment - EmailMessage.add_attachment handles encoding properly
            msg.add_attachment(
                data,
                maintype=main_type,
                subtype=sub_type,
                filename=filename
            )
            
        except Exception as e:
            print(f"Warning: Could not attach {file_path}: {e}")


class EmailBatchSender:
    """Handle batch email sending with progress tracking"""
    
    def __init__(self):
        self.sender: Optional[SMTPSender] = None
        self.is_running = False
        self.is_aborted = False
        self.thread: Optional[threading.Thread] = None
        self.progress_queue: Queue = Queue()
    
    def start_sending(
        self,
        host: str,
        port: int,
        email: str,
        password: str,
        use_tls: bool,
        data: List[Dict],
        default_subject: str,
        default_body: str,
        cc: Optional[str],
        bcc: Optional[str],
        attachment_files: List[str],
        delay_ms: int,
        randomize_percent: int
    ):
        """Start sending emails in background thread"""
        self.is_running = True
        self.is_aborted = False
        self.progress_queue = Queue()
        
        self.thread = threading.Thread(
            target=self._send_batch,
            args=(
                host, port, email, password, use_tls,
                data, default_subject, default_body, cc, bcc,
                attachment_files, delay_ms, randomize_percent
            ),
            daemon=True
        )
        self.thread.start()
    
    def _send_batch(
        self,
        host: str,
        port: int,
        email: str,
        password: str,
        use_tls: bool,
        data: List[Dict],
        default_subject: str,
        default_body: str,
        cc: Optional[str],
        bcc: Optional[str],
        attachment_files: List[str],
        delay_ms: int,
        randomize_percent: int
    ):
        """Background thread for batch sending"""
        total = len(data)
        sender = SMTPSender(host, port, email, password, use_tls)
        
        # Report connection attempt
        self._report_progress('connecting', 0, total, "Connecting to SMTP server...")
        
        # Connect
        success, msg = sender.connect()
        if not success:
            self._report_progress('error', 0, total, f"Connection failed: {msg}")
            self.is_running = False
            return
        
        self._report_progress('connected', 0, total, "Connected. Starting to send emails...")
        
        import random
        sent_count = 0
        failed_count = 0
        
        try:
            for i, row in enumerate(data):
                if self.is_aborted:
                    self._report_progress('aborted', sent_count + failed_count, total, "Sending aborted by user")
                    break
                
                # Get email details
                recipient = row.get('recipient_email', '')
                subject = row.get('subject', '') or default_subject
                body = row.get('body_content', '') or default_body
                
                # Build attachments
                attachments = []
                # Handle multiple attachments from CSV (attachment_filenames is a list)
                csv_attachments = row.get('attachment_filenames', [])
                for csv_attach in csv_attachments:
                    if os.path.exists(csv_attach):
                        attachments.append(csv_attach)
                # Add global attachments to all emails
                attachments.extend([f for f in attachment_files if os.path.exists(f)])
                
                # Send email
                self._report_progress('sending', sent_count + failed_count, total, f"Sending to {recipient}...")
                success, msg = sender.send_email(recipient, subject, body, attachments, cc, bcc)
                
                if success:
                    sent_count += 1
                    self._report_progress('sent', sent_count + failed_count, total, f"✓ Sent to {recipient}")
                else:
                    failed_count += 1
                    self._report_progress('failed', sent_count + failed_count, total, f"✗ Failed to {recipient}: {msg}")
                
                # Delay before next email (except for last)
                if i < total - 1 and not self.is_aborted:
                    if randomize_percent > 0:
                        variance = delay_ms * randomize_percent / 100
                        actual_delay = delay_ms + random.uniform(-variance, variance)
                    else:
                        actual_delay = delay_ms
                    time.sleep(actual_delay / 1000)
            
            if not self.is_aborted:
                self._report_progress(
                    'complete',
                    sent_count + failed_count,
                    total,
                    f"Complete! Sent: {sent_count}, Failed: {failed_count}"
                )
        
        except Exception as e:
            self._report_progress('error', sent_count + failed_count, total, f"Error: {str(e)}")
        
        finally:
            sender.disconnect()
            self.is_running = False
    
    def _report_progress(self, status: str, current: int, total: int, message: str):
        """Send progress update to queue"""
        self.progress_queue.put({
            'status': status,
            'current': current,
            'total': total,
            'message': message
        })
    
    def abort(self):
        """Abort sending"""
        self.is_aborted = True
    
    def get_progress(self) -> Optional[Dict]:
        """Get progress update from queue (non-blocking)"""
        try:
            return self.progress_queue.get_nowait()
        except:
            return None
    
    def is_sending(self) -> bool:
        """Check if currently sending"""
        return self.is_running

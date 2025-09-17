#!/usr/bin/env python3
"""
Gmail to GitHub Email Integration
Automatically forwards emails from Gmail to GitHub issues for job search tracking.
"""

import os
import json
import imaplib
import email
import re
import time
from datetime import datetime, timedelta
from typing import List, Dict, Optional, Tuple
import logging
from dataclasses import dataclass
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import ssl
import smtplib

# Try to import GitHub integration - fall back gracefully if not available
try:
    from github import Github
    GITHUB_AVAILABLE = True
except ImportError:
    GITHUB_AVAILABLE = False
    logging.warning("PyGithub not available. GitHub integration disabled.")

@dataclass
class EmailConfig:
    """Configuration for email integration"""
    gmail_user: str
    gmail_password: str  # App password
    github_token: str
    github_repo: str
    imap_server: str = "imap.gmail.com"
    imap_port: int = 993
    check_interval: int = 300  # 5 minutes
    processed_emails_file: str = "processed_emails.json"

@dataclass
class EmailMessage:
    """Represents an email message"""
    subject: str
    sender: str
    recipient: str
    body: str
    date: datetime
    message_id: str
    is_job_related: bool = False
    job_keywords: List[str] = None
    priority: str = "normal"  # low, normal, high

class EmailClassifier:
    """Classifies emails for job search relevance"""
    
    JOB_KEYWORDS = [
        # General job-related terms
        'job', 'position', 'opportunity', 'career', 'employment', 'hiring',
        'interview', 'application', 'resume', 'cv', 'candidate', 'recruiter',
        'recruitment', 'hr', 'human resources', 'talent acquisition',
        
        # Response types
        'application received', 'application status', 'next steps', 'interview invitation',
        'phone screen', 'screening call', 'technical interview', 'coding challenge',
        'assessment', 'offer', 'congratulations', 'unfortunately', 'not selected',
        
        # Company-related
        'company', 'team', 'department', 'role', 'responsibilities', 'requirements',
        'qualifications', 'experience', 'skills', 'salary', 'compensation', 'benefits',
        
        # Technical terms
        'developer', 'engineer', 'analyst', 'specialist', 'coordinator', 'manager',
        'director', 'senior', 'junior', 'lead', 'principal', 'architect',
        'software', 'technology', 'technical', 'programming', 'coding'
    ]
    
    HIGH_PRIORITY_KEYWORDS = [
        'interview', 'job offer', 'urgent', 'time sensitive', 'deadline',
        'final round', 'decision', 'congratulations', 'selected'
    ]
    
    LOW_PRIORITY_KEYWORDS = [
        'newsletter', 'marketing', 'promotion', 'advertisement', 'unsubscribe',
        'no-reply', 'noreply', 'automated', 'newsletter', 'deal', 'offer', 'sale'
    ]
    
    def classify_email(self, email_msg: EmailMessage) -> EmailMessage:
        """Classify email for job search relevance and priority"""
        text_to_analyze = f"{email_msg.subject} {email_msg.body}".lower()
        
        # Check for job-related keywords
        job_keywords_found = []
        for keyword in self.JOB_KEYWORDS:
            if keyword.lower() in text_to_analyze:
                job_keywords_found.append(keyword)
        
        # Determine if job-related
        email_msg.is_job_related = len(job_keywords_found) >= 2 or any(
            priority_word in text_to_analyze for priority_word in ['interview', 'application', 'recruiter']
        )
        
        email_msg.job_keywords = job_keywords_found
        
        # Determine priority
        if any(word in text_to_analyze for word in self.HIGH_PRIORITY_KEYWORDS):
            email_msg.priority = "high"
        elif any(word in text_to_analyze for word in self.LOW_PRIORITY_KEYWORDS):
            email_msg.priority = "low"
        else:
            email_msg.priority = "normal"
        
        return email_msg

class GmailConnector:
    """Handles Gmail IMAP connection and email retrieval"""
    
    def __init__(self, config: EmailConfig):
        self.config = config
        self.classifier = EmailClassifier()
        self.processed_emails = self._load_processed_emails()
    
    def _load_processed_emails(self) -> set:
        """Load list of already processed email IDs"""
        try:
            if os.path.exists(self.config.processed_emails_file):
                with open(self.config.processed_emails_file, 'r') as f:
                    data = json.load(f)
                    return set(data.get('processed_ids', []))
        except Exception as e:
            logging.warning(f"Could not load processed emails: {e}")
        return set()
    
    def _save_processed_emails(self):
        """Save list of processed email IDs"""
        try:
            data = {
                'processed_ids': list(self.processed_emails),
                'last_updated': datetime.now().isoformat()
            }
            with open(self.config.processed_emails_file, 'w') as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            logging.error(f"Could not save processed emails: {e}")
    
    def connect_to_gmail(self) -> imaplib.IMAP4_SSL:
        """Establish connection to Gmail IMAP"""
        try:
            # Create SSL context
            context = ssl.create_default_context()
            
            # Connect to Gmail IMAP
            mail = imaplib.IMAP4_SSL(self.config.imap_server, self.config.imap_port, ssl_context=context)
            mail.login(self.config.gmail_user, self.config.gmail_password)
            
            logging.info("Successfully connected to Gmail")
            return mail
            
        except Exception as e:
            logging.error(f"Failed to connect to Gmail: {e}")
            raise
    
    def fetch_new_emails(self, mailbox: str = "INBOX", days_back: int = 7) -> List[EmailMessage]:
        """Fetch new emails from Gmail"""
        emails = []
        mail = None
        
        try:
            mail = self.connect_to_gmail()
            mail.select(mailbox)
            
            # Search for emails from last N days
            date_since = (datetime.now() - timedelta(days=days_back)).strftime("%d-%b-%Y")
            search_criteria = f'(SINCE "{date_since}")'
            
            status, message_ids = mail.search(None, search_criteria)
            
            if status != 'OK':
                logging.error(f"Failed to search emails: {status}")
                return emails
            
            message_ids = message_ids[0].split()
            logging.info(f"Found {len(message_ids)} emails in last {days_back} days")
            
            # Process each email
            for msg_id in message_ids:
                try:
                    msg_id_str = msg_id.decode()
                    
                    # Skip if already processed
                    if msg_id_str in self.processed_emails:
                        continue
                    
                    # Fetch email
                    status, msg_data = mail.fetch(msg_id, '(RFC822)')
                    if status != 'OK':
                        continue
                    
                    # Parse email
                    email_msg = self._parse_email(msg_data[0][1], msg_id_str)
                    if email_msg:
                        # Classify email
                        email_msg = self.classifier.classify_email(email_msg)
                        emails.append(email_msg)
                        
                        # Mark as processed
                        self.processed_emails.add(msg_id_str)
                        
                except Exception as e:
                    logging.warning(f"Error processing email {msg_id}: {e}")
                    continue
            
            # Save processed emails
            self._save_processed_emails()
            
        except Exception as e:
            logging.error(f"Error fetching emails: {e}")
        finally:
            if mail:
                try:
                    mail.close()
                    mail.logout()
                except:
                    pass
        
        return emails
    
    def _parse_email(self, raw_email: bytes, message_id: str) -> Optional[EmailMessage]:
        """Parse raw email into EmailMessage object"""
        try:
            msg = email.message_from_bytes(raw_email)
            
            # Extract basic headers
            subject = self._decode_header(msg.get('Subject', ''))
            sender = self._decode_header(msg.get('From', ''))
            recipient = self._decode_header(msg.get('To', ''))
            date_str = msg.get('Date', '')
            
            # Parse date
            try:
                email_date = email.utils.parsedate_to_datetime(date_str)
            except:
                email_date = datetime.now()
            
            # Extract body
            body = self._extract_email_body(msg)
            
            return EmailMessage(
                subject=subject,
                sender=sender,
                recipient=recipient,
                body=body,
                date=email_date,
                message_id=message_id
            )
            
        except Exception as e:
            logging.error(f"Error parsing email: {e}")
            return None
    
    def _decode_header(self, header: str) -> str:
        """Decode email header"""
        try:
            decoded = email.header.decode_header(header)
            return ''.join([
                part.decode(encoding or 'utf-8') if isinstance(part, bytes) else part
                for part, encoding in decoded
            ])
        except:
            return str(header)
    
    def _extract_email_body(self, msg) -> str:
        """Extract email body text"""
        body = ""
        
        try:
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        body_part = part.get_payload(decode=True)
                        if body_part:
                            body += body_part.decode('utf-8', errors='ignore')
                    elif content_type == "text/html" and not body:
                        # Fallback to HTML if no plain text
                        html_part = part.get_payload(decode=True)
                        if html_part:
                            # Basic HTML to text conversion
                            html_text = html_part.decode('utf-8', errors='ignore')
                            # Remove HTML tags (basic)
                            body = re.sub(r'<[^>]+>', '', html_text)
            else:
                body = msg.get_payload(decode=True)
                if body:
                    body = body.decode('utf-8', errors='ignore')
        
        except Exception as e:
            logging.warning(f"Error extracting email body: {e}")
        
        return body.strip()

class GitHubIntegration:
    """Handles GitHub issue creation from emails"""
    
    def __init__(self, config: EmailConfig, connect=True):
        self.config = config
        self.github = None
        
        if connect and GITHUB_AVAILABLE and config.github_token:
            try:
                self.github = Github(config.github_token)
                self.repo = self.github.get_repo(config.github_repo)
                logging.info(f"Connected to GitHub repo: {config.github_repo}")
            except Exception as e:
                logging.error(f"Failed to connect to GitHub: {e}")
    
    def create_issue_from_email(self, email_msg: EmailMessage) -> Optional[str]:
        """Create a GitHub issue from an email"""
        if not self.github:
            logging.warning("GitHub not available, cannot create issue")
            return None
        
        try:
            # Create issue title
            title = self._create_issue_title(email_msg)
            
            # Create issue body
            body = self._create_issue_body(email_msg)
            
            # Determine labels
            labels = self._determine_labels(email_msg)
            
            # Create the issue
            issue = self.repo.create_issue(
                title=title,
                body=body,
                labels=labels
            )
            
            logging.info(f"Created GitHub issue #{issue.number}: {title}")
            return issue.html_url
            
        except Exception as e:
            logging.error(f"Failed to create GitHub issue: {e}")
            return None
    
    def _create_issue_title(self, email_msg: EmailMessage) -> str:
        """Create issue title from email"""
        priority_prefix = ""
        if email_msg.priority == "high":
            priority_prefix = "🔥 HIGH PRIORITY: "
        elif email_msg.priority == "low":
            priority_prefix = "📋 "
        else:
            priority_prefix = "📧 "
        
        # Clean up subject
        subject = email_msg.subject
        if len(subject) > 80:
            subject = subject[:77] + "..."
        
        return f"{priority_prefix}Email: {subject}"
    
    def _create_issue_body(self, email_msg: EmailMessage) -> str:
        """Create issue body from email"""
        body = f"""## Email Details

**From:** {email_msg.sender}
**To:** {email_msg.recipient}
**Date:** {email_msg.date.strftime('%Y-%m-%d %H:%M:%S')}
**Priority:** {email_msg.priority.upper()}

## Subject
{email_msg.subject}

## Message Body
```
{email_msg.body[:2000]}{'...' if len(email_msg.body) > 2000 else ''}
```

## Analysis
- **Job Related:** {'✅ Yes' if email_msg.is_job_related else '❌ No'}
- **Keywords Found:** {', '.join(email_msg.job_keywords) if email_msg.job_keywords else 'None'}

## Actions
- [ ] Read and respond
- [ ] Follow up required
- [ ] Archive
- [ ] Add to job tracker

---
*Auto-generated from Gmail integration*
"""
        return body
    
    def _determine_labels(self, email_msg: EmailMessage) -> List[str]:
        """Determine appropriate labels for the issue"""
        labels = ["email"]
        
        if email_msg.is_job_related:
            labels.append("job-search")
        
        if email_msg.priority == "high":
            labels.append("high-priority")
        elif email_msg.priority == "low":
            labels.append("low-priority")
        
        # Add labels based on keywords
        keyword_labels = {
            'interview': 'interview',
            'recruiter': 'recruiter',
            'application': 'application',
            'offer': 'job-offer',
            'rejection': 'rejection'
        }
        
        email_text = f"{email_msg.subject} {email_msg.body}".lower()
        for keyword, label in keyword_labels.items():
            if keyword in email_text:
                labels.append(label)
        
        return labels

class EmailForwarder:
    """Main class for email forwarding integration"""
    
    def __init__(self, config: EmailConfig):
        self.config = config
        self.gmail_connector = GmailConnector(config)
        self.github_integration = GitHubIntegration(config)
        
        # Setup logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('email_integration.log'),
                logging.StreamHandler()
            ]
        )
    
    def process_emails(self, days_back: int = 1) -> Dict[str, int]:
        """Process new emails and create GitHub issues"""
        stats = {
            'total_emails': 0,
            'job_related': 0,
            'issues_created': 0,
            'errors': 0
        }
        
        try:
            logging.info("Starting email processing...")
            
            # Fetch new emails
            emails = self.gmail_connector.fetch_new_emails(days_back=days_back)
            stats['total_emails'] = len(emails)
            
            logging.info(f"Processing {len(emails)} new emails")
            
            for email_msg in emails:
                try:
                    # Only create issues for job-related emails
                    if email_msg.is_job_related:
                        stats['job_related'] += 1
                        
                        # Create GitHub issue
                        issue_url = self.github_integration.create_issue_from_email(email_msg)
                        if issue_url:
                            stats['issues_created'] += 1
                            logging.info(f"Created issue for email: {email_msg.subject[:50]}...")
                        else:
                            stats['errors'] += 1
                    else:
                        logging.debug(f"Skipping non-job-related email: {email_msg.subject[:50]}...")
                
                except Exception as e:
                    logging.error(f"Error processing email: {e}")
                    stats['errors'] += 1
            
            logging.info(f"Email processing complete. Stats: {stats}")
            
        except Exception as e:
            logging.error(f"Fatal error in email processing: {e}")
            stats['errors'] += 1
        
        return stats
    
    def run_continuous(self):
        """Run email forwarding continuously"""
        logging.info(f"Starting continuous email forwarding (check every {self.config.check_interval} seconds)")
        
        while True:
            try:
                stats = self.process_emails()
                
                if stats['total_emails'] > 0:
                    logging.info(f"Processed {stats['total_emails']} emails, created {stats['issues_created']} issues")
                
                # Wait for next check
                time.sleep(self.config.check_interval)
                
            except KeyboardInterrupt:
                logging.info("Email forwarding stopped by user")
                break
            except Exception as e:
                logging.error(f"Error in continuous mode: {e}")
                time.sleep(60)  # Wait 1 minute before retrying

def load_config(config_file: str = "email_config.json") -> EmailConfig:
    """Load configuration from file or environment variables"""
    config_data = {}
    
    # Try to load from file
    if os.path.exists(config_file):
        try:
            with open(config_file, 'r') as f:
                config_data = json.load(f)
        except Exception as e:
            logging.warning(f"Could not load config file: {e}")
    
    # Override with environment variables
    env_mapping = {
        'gmail_user': 'GMAIL_USER',
        'gmail_password': 'GMAIL_APP_PASSWORD',
        'github_token': 'GITHUB_TOKEN',
        'github_repo': 'GITHUB_REPO'
    }
    
    for config_key, env_key in env_mapping.items():
        if env_key in os.environ:
            config_data[config_key] = os.environ[env_key]
    
    # Validate required fields
    required_fields = ['gmail_user', 'gmail_password', 'github_token', 'github_repo']
    missing_fields = [field for field in required_fields if not config_data.get(field)]
    
    if missing_fields:
        raise ValueError(f"Missing required configuration: {', '.join(missing_fields)}")
    
    return EmailConfig(**config_data)

def create_sample_config():
    """Create a sample configuration file"""
    sample_config = {
        "gmail_user": "your-email@gmail.com",
        "gmail_password": "your-app-password",
        "github_token": "your-github-token",
        "github_repo": "username/repository",
        "check_interval": 300,
        "processed_emails_file": "processed_emails.json"
    }
    
    with open("email_config.json.sample", 'w') as f:
        json.dump(sample_config, f, indent=2)
    
    print("Sample configuration created: email_config.json.sample")
    print("\nSetup instructions:")
    print("1. Copy email_config.json.sample to email_config.json")
    print("2. Update with your Gmail and GitHub credentials")
    print("3. Enable 2FA on Gmail and create an App Password")
    print("4. Create a GitHub Personal Access Token with repo permissions")

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == "setup":
        create_sample_config()
        sys.exit(0)
    
    try:
        # Load configuration
        config = load_config()
        
        # Create email forwarder
        forwarder = EmailForwarder(config)
        
        if len(sys.argv) > 1 and sys.argv[1] == "continuous":
            # Run continuously
            forwarder.run_continuous()
        else:
            # Run once
            stats = forwarder.process_emails()
            print(f"Email processing complete: {stats}")
    
    except Exception as e:
        logging.error(f"Error: {e}")
        print(f"\nError: {e}")
        print("\nTo create a sample configuration, run:")
        print("python email_integration.py setup")
        sys.exit(1)
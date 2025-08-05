# Gmail to GitHub Email Forwarding Setup

This guide will help you set up automatic forwarding of Gmail emails to GitHub issues for better job search tracking and organization.

## Overview

The email integration system will:
- Connect to your Gmail account via IMAP
- Classify emails for job search relevance
- Create GitHub issues for important emails
- Track processed emails to avoid duplicates
- Run continuously or on-demand

## Prerequisites

1. **Gmail Account** with 2-Factor Authentication enabled
2. **GitHub Account** with repository access
3. **Python 3.7+** installed

## Setup Steps

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Create Gmail App Password

1. Go to your [Google Account settings](https://myaccount.google.com/)
2. Select **Security** → **2-Step Verification** 
3. Under "Signing in to Google", select **App passwords**
4. Generate a new app password for "Mail"
5. Save this 16-character password - you'll need it for configuration

### 3. Create GitHub Personal Access Token

1. Go to [GitHub Settings → Developer settings → Personal access tokens](https://github.com/settings/tokens)
2. Click "Generate new token (classic)"
3. Give it a descriptive name like "Email Integration"
4. Select scopes:
   - `repo` (Full control of private repositories)
   - `public_repo` (Access public repositories)
5. Copy the generated token

### 4. Configure Email Integration

1. Create configuration file:
```bash
python email_integration.py setup
cp email_config.json.sample email_config.json
```

2. Edit `email_config.json` with your credentials:
```json
{
  "gmail_user": "your-email@gmail.com",
  "gmail_password": "your-16-char-app-password",
  "github_token": "ghp_your-github-token",
  "github_repo": "iltolesjr/RESUMES",
  "check_interval": 300,
  "processed_emails_file": "processed_emails.json"
}
```

### 5. Test the Integration

Test with recent emails (dry run):
```bash
python email_integration.py
```

This will:
- Connect to Gmail
- Fetch emails from the last day
- Classify job-related emails
- Create GitHub issues for relevant emails
- Show processing statistics

## Usage Options

### One-time Processing
```bash
python email_integration.py
```
Processes new emails once and exits.

### Continuous Monitoring
```bash
python email_integration.py continuous
```
Runs continuously, checking for new emails every 5 minutes (configurable).

### Background Service
For production use, run as a background service:
```bash
nohup python email_integration.py continuous > email_integration.log 2>&1 &
```

## Configuration Options

| Setting | Description | Default |
|---------|-------------|---------|
| `gmail_user` | Your Gmail address | Required |
| `gmail_password` | Gmail app password | Required |
| `github_token` | GitHub personal access token | Required |
| `github_repo` | GitHub repository (owner/repo) | Required |
| `check_interval` | Seconds between email checks | 300 (5 min) |
| `processed_emails_file` | File to track processed emails | processed_emails.json |

## Email Classification

The system automatically classifies emails based on keywords:

### Job-Related Keywords
- Interview, application, recruiter, hiring
- Position, opportunity, career, employment
- Resume, CV, candidate, talent acquisition
- Company names and job titles

### Priority Levels
- **High Priority**: Interview invitations, job offers, urgent deadlines
- **Normal Priority**: Standard job-related communications
- **Low Priority**: Newsletters, automated messages

### GitHub Labels
Created issues are automatically labeled:
- `email` - All email-generated issues
- `job-search` - Job-related emails
- `high-priority` - Urgent emails
- `interview` - Interview-related
- `job-offer` - Job offers
- `recruiter` - Recruiter communications

## Troubleshooting

### Connection Issues
```
Error: Authentication failed
```
- Verify Gmail app password is correct
- Ensure 2FA is enabled on Gmail
- Check that IMAP is enabled in Gmail settings

### GitHub Issues
```
Error: Failed to connect to GitHub
```
- Verify GitHub token has repo permissions
- Check repository name format (owner/repo)
- Ensure token hasn't expired

### Email Processing
```
Error: Could not parse email
```
- Check Gmail IMAP settings
- Verify email formatting
- Review email_integration.log for details

## Security Best Practices

1. **Use App Passwords**: Never use your main Gmail password
2. **Limit Token Scope**: Only grant necessary GitHub permissions
3. **Secure Storage**: Keep configuration files private
4. **Regular Rotation**: Rotate tokens and passwords periodically

## Integration with Existing Workflow

This email integration works alongside your existing job search automation:

1. **Job Scraping** (`job_scraper.py`) - Finds job postings
2. **Resume Tailoring** (`job_leads_agent.py`) - Creates custom resumes
3. **Email Integration** (`email_integration.py`) - Tracks email responses

All systems create GitHub issues for centralized tracking.

## Customization

### Add Custom Keywords
Edit the `JOB_KEYWORDS` list in `email_integration.py`:
```python
JOB_KEYWORDS = [
    'your-custom-keywords',
    'specific-companies',
    'job-titles'
]
```

### Modify Classification Rules
Update the `classify_email` method to change how emails are categorized.

### Custom GitHub Labels
Modify the `_determine_labels` method to use your preferred labeling system.

## Logs and Monitoring

- **Log File**: `email_integration.log`
- **Processed Emails**: `processed_emails.json`
- **GitHub Issues**: Check your repository's Issues tab

## Support

For issues or questions:
1. Check the log file for error details
2. Verify configuration settings
3. Test individual components (Gmail connection, GitHub API)
4. Review GitHub repository issues for similar problems
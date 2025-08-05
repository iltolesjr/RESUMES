#!/usr/bin/env python3
"""
Test script for Gmail to GitHub email integration
Tests the email integration components without requiring actual credentials
"""

import os
import json
import tempfile
from datetime import datetime
from email_integration import EmailClassifier, EmailMessage, EmailConfig, GitHubIntegration

def test_email_classifier():
    """Test email classification functionality"""
    print("Testing Email Classifier...")
    
    classifier = EmailClassifier()
    
    # Test job-related email
    job_email = EmailMessage(
        subject="Interview Invitation - Software Engineer Position",
        sender="recruiter@techcorp.com",
        recipient="candidate@gmail.com",
        body="We would like to invite you for an interview for the Software Engineer position. Please reply with your availability.",
        date=datetime.now(),
        message_id="test1"
    )
    
    classified = classifier.classify_email(job_email)
    
    assert classified.is_job_related == True, "Job email should be classified as job-related"
    assert classified.priority == "high", "Interview invitation should be high priority"
    assert "interview" in classified.job_keywords, "Should detect 'interview' keyword"
    
    print("✅ Job-related email classified correctly")
    
    # Test non-job email
    spam_email = EmailMessage(
        subject="Amazing Deal on Products!",
        sender="marketing@spamcorp.com",
        recipient="candidate@gmail.com",
        body="Get 50% off on all products. Limited time offer!",
        date=datetime.now(),
        message_id="test2"
    )
    
    classified_spam = classifier.classify_email(spam_email)
    
    assert classified_spam.is_job_related == False, "Spam email should not be job-related"
    assert classified_spam.priority == "low", "Marketing email should be low priority"
    
    print("✅ Non-job email classified correctly")

def test_github_integration():
    """Test GitHub integration (without actual GitHub connection)"""
    print("\nTesting GitHub Integration...")
    
    # Create test config without real credentials
    config = EmailConfig(
        gmail_user="test@example.com",
        gmail_password="fake-password",
        github_token="fake-token",
        github_repo="test/repo"
    )
    
    # Create GitHubIntegration without actually connecting
    github_integration = GitHubIntegration(config, connect=False)
    
    test_email = EmailMessage(
        subject="Application Status Update - Data Analyst Position",
        sender="hr@company.com",
        recipient="candidate@gmail.com",
        body="Thank you for your application. We are currently reviewing candidates.",
        date=datetime.now(),
        message_id="test3",
        is_job_related=True,
        job_keywords=["application", "position", "candidates"],
        priority="normal"
    )
    
    title = github_integration._create_issue_title(test_email)
    body = github_integration._create_issue_body(test_email)
    labels = github_integration._determine_labels(test_email)
    
    assert "Application Status Update" in title, "Title should include email subject"
    assert "hr@company.com" in body, "Body should include sender"
    assert "email" in labels, "Should have 'email' label"
    assert "job-search" in labels, "Should have 'job-search' label"
    
    print("✅ GitHub issue creation formatted correctly")

def test_config_loading():
    """Test configuration loading"""
    print("\nTesting Configuration Loading...")
    
    # Create temporary config file
    with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
        test_config = {
            "gmail_user": "test@example.com",
            "gmail_password": "test-password",
            "github_token": "test-token",
            "github_repo": "test/repo",
            "check_interval": 600
        }
        json.dump(test_config, f)
        temp_config_file = f.name
    
    try:
        # Test loading from file
        from email_integration import load_config
        
        # Set environment variables for testing
        os.environ['GMAIL_USER'] = 'env@example.com'
        os.environ['GITHUB_TOKEN'] = 'env-token'
        
        # Load config (should use env vars over file)
        config = load_config(temp_config_file)
        
        assert config.gmail_user == 'env@example.com', "Should use environment variable"
        assert config.github_token == 'env-token', "Should use environment variable"
        assert config.check_interval == 600, "Should use file value when env not set"
        
        print("✅ Configuration loading works correctly")
        
    finally:
        # Cleanup
        os.unlink(temp_config_file)
        if 'GMAIL_USER' in os.environ:
            del os.environ['GMAIL_USER']
        if 'GITHUB_TOKEN' in os.environ:
            del os.environ['GITHUB_TOKEN']

def test_email_parsing():
    """Test email parsing components"""
    print("\nTesting Email Parsing...")
    
    # Test keyword extraction through classification
    classifier = EmailClassifier()
    
    test_email = EmailMessage(
        subject="Senior Python Developer Interview",
        sender="recruiter@company.com",
        recipient="candidate@gmail.com",
        body="""
        Dear Candidate,
        
        Thank you for your application for the Senior Python Developer position.
        We would like to schedule a technical interview next week.
        
        The role requires 5+ years experience with Python, Django, and AWS.
        
        Best regards,
        Tech Recruiter
        """,
        date=datetime.now(),
        message_id="test_parsing"
    )
    
    classified = classifier.classify_email(test_email)
    
    expected_keywords = ['position', 'interview', 'application', 'developer', 'recruiter']
    found_keywords = [kw.lower() for kw in classified.job_keywords]
    
    for keyword in expected_keywords:
        assert keyword in found_keywords, f"Should detect '{keyword}' in keywords: {found_keywords}"
    
    print("✅ Keyword extraction works correctly")

def run_integration_test():
    """Run a full integration test with mock data"""
    print("\nRunning Integration Test...")
    
    # Test email messages
    test_emails = [
        {
            "subject": "Job Interview Invitation - Frontend Developer",
            "sender": "hiring@startup.com",
            "body": "We'd like to invite you for an interview for our Frontend Developer position using React and JavaScript.",
            "expected_job_related": True,
            "expected_priority": "high"
        },
        {
            "subject": "Weekly Newsletter - Tech Updates",
            "sender": "newsletter@techblog.com", 
            "body": "Here are this week's top tech stories and trends.",
            "expected_job_related": False,
            "expected_priority": "low"
        },
        {
            "subject": "Application Received - Software Engineer",
            "sender": "noreply@company.com",
            "body": "Thank you for your application. We will review and get back to you within 2 weeks.",
            "expected_job_related": True,
            "expected_priority": "normal"
        }
    ]
    
    classifier = EmailClassifier()
    
    for i, email_data in enumerate(test_emails):
        email_msg = EmailMessage(
            subject=email_data["subject"],
            sender=email_data["sender"],
            recipient="candidate@gmail.com",
            body=email_data["body"],
            date=datetime.now(),
            message_id=f"test_{i}"
        )
        
        classified = classifier.classify_email(email_msg)
        
        assert classified.is_job_related == email_data["expected_job_related"], \
            f"Email {i+1}: Job classification incorrect"
        assert classified.priority == email_data["expected_priority"], \
            f"Email {i+1}: Priority classification incorrect"
        
        print(f"✅ Email {i+1}: {email_data['subject'][:30]}... classified correctly")

def main():
    """Run all tests"""
    print("Gmail to GitHub Email Integration - Test Suite")
    print("=" * 50)
    
    try:
        test_email_classifier()
        test_github_integration()
        test_config_loading()
        test_email_parsing()
        run_integration_test()
        
        print("\n" + "=" * 50)
        print("🎉 All tests passed! Email integration is working correctly.")
        print("\nNext steps:")
        print("1. Set up your Gmail app password")
        print("2. Create a GitHub personal access token")
        print("3. Configure email_config.json")
        print("4. Run: python email_integration.py")
        
    except AssertionError as e:
        print(f"\n❌ Test failed: {e}")
        return 1
    except Exception as e:
        print(f"\n💥 Unexpected error: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    exit(main())
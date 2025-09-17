#!/usr/bin/env python3
"""
Integrated Job Search Workflow
Combines email monitoring, job scraping, and resume tailoring for comprehensive job search automation.
"""

import os
import sys
import time
import logging
from datetime import datetime, timedelta
from typing import Dict, List

# Import existing modules
try:
    from email_integration import EmailForwarder, EmailConfig, load_config
    EMAIL_AVAILABLE = True
except ImportError:
    EMAIL_AVAILABLE = False
    logging.warning("Email integration not available")

try:
    from job_scraper import JobScraper, EnhancedResumeProcessor
    JOB_SCRAPER_AVAILABLE = True
except ImportError:
    JOB_SCRAPER_AVAILABLE = False
    logging.warning("Job scraper not available")

try:
    from job_leads_agent import main as create_job_leads
    JOB_LEADS_AVAILABLE = True
except ImportError:
    JOB_LEADS_AVAILABLE = False
    logging.warning("Job leads agent not available")

class JobSearchWorkflow:
    """Integrated job search workflow manager"""
    
    def __init__(self, config_file: str = "email_config.json"):
        self.config_file = config_file
        self.setup_logging()
        
        # Initialize components
        self.email_forwarder = None
        self.job_scraper = None
        self.resume_processor = None
        
        if EMAIL_AVAILABLE:
            try:
                config = load_config(config_file)
                self.email_forwarder = EmailForwarder(config)
                logging.info("Email integration initialized")
            except Exception as e:
                logging.warning(f"Could not initialize email integration: {e}")
        
        if JOB_SCRAPER_AVAILABLE:
            try:
                self.job_scraper = JobScraper()
                self.resume_processor = EnhancedResumeProcessor()
                logging.info("Job scraper initialized")
            except Exception as e:
                logging.warning(f"Could not initialize job scraper: {e}")
    
    def setup_logging(self):
        """Setup logging for the workflow"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('job_workflow.log'),
                logging.StreamHandler()
            ]
        )
    
    def run_daily_workflow(self) -> Dict[str, int]:
        """Run the complete daily job search workflow"""
        stats = {
            'emails_processed': 0,
            'jobs_found': 0,
            'resumes_created': 0,
            'issues_created': 0,
            'errors': 0
        }
        
        logging.info("Starting daily job search workflow...")
        
        # Step 1: Process emails
        if self.email_forwarder:
            try:
                email_stats = self.email_forwarder.process_emails(days_back=1)
                stats['emails_processed'] = email_stats.get('total_emails', 0)
                stats['issues_created'] += email_stats.get('issues_created', 0)
                stats['errors'] += email_stats.get('errors', 0)
                logging.info(f"Email processing: {email_stats}")
            except Exception as e:
                logging.error(f"Email processing failed: {e}")
                stats['errors'] += 1
        
        # Step 2: Scrape new jobs (if configured)
        if self.job_scraper:
            try:
                job_search_terms = self._get_job_search_terms()
                location = self._get_search_location()
                
                for term in job_search_terms:
                    jobs = self.job_scraper.scrape_indeed_jobs(term, location, max_jobs=5)
                    stats['jobs_found'] += len(jobs)
                    
                    # Create GitHub issues for new jobs
                    for job in jobs:
                        try:
                            # Convert job to issue format
                            self._create_job_issue(job)
                            stats['issues_created'] += 1
                        except Exception as e:
                            logging.error(f"Failed to create job issue: {e}")
                            stats['errors'] += 1
                
                logging.info(f"Job scraping: Found {stats['jobs_found']} jobs")
            except Exception as e:
                logging.error(f"Job scraping failed: {e}")
                stats['errors'] += 1
        
        # Step 3: Generate tailored resumes for high-priority opportunities
        if self.resume_processor and stats['jobs_found'] > 0:
            try:
                # This would integrate with the user's profile data
                stats['resumes_created'] = self._create_tailored_resumes()
                logging.info(f"Resume creation: Generated {stats['resumes_created']} resumes")
            except Exception as e:
                logging.error(f"Resume creation failed: {e}")
                stats['errors'] += 1
        
        logging.info(f"Daily workflow complete: {stats}")
        return stats
    
    def run_email_monitoring(self, check_interval: int = 300):
        """Run continuous email monitoring"""
        if not self.email_forwarder:
            logging.error("Email integration not available")
            return
        
        logging.info(f"Starting email monitoring (check every {check_interval} seconds)")
        
        while True:
            try:
                stats = self.email_forwarder.process_emails(days_back=1)
                
                if stats['total_emails'] > 0:
                    logging.info(f"Email check: {stats}")
                
                time.sleep(check_interval)
                
            except KeyboardInterrupt:
                logging.info("Email monitoring stopped by user")
                break
            except Exception as e:
                logging.error(f"Error in email monitoring: {e}")
                time.sleep(60)  # Wait 1 minute before retrying
    
    def _get_job_search_terms(self) -> List[str]:
        """Get job search terms from configuration or defaults"""
        # This could be loaded from a config file or user profile
        return [
            "software engineer",
            "python developer", 
            "data analyst",
            "IT support specialist",
            "systems administrator"
        ]
    
    def _get_search_location(self) -> str:
        """Get search location from configuration"""
        # This could be loaded from user profile
        return "Minneapolis, MN"
    
    def _create_job_issue(self, job):
        """Create GitHub issue for a job opportunity"""
        if not self.email_forwarder or not self.email_forwarder.github_integration.github:
            logging.warning("GitHub integration not available for job issue creation")
            return
        
        try:
            title = f"🎯 Job Opportunity: {job.title} at {job.company}"
            
            body = f"""## Job Details

**Company:** {job.company}
**Location:** {job.location}
**URL:** {job.url}

## Job Description
{job.description[:1000]}{'...' if len(job.description) > 1000 else ''}

## Keywords Found
{', '.join(job.keywords) if job.keywords else 'None'}

## Actions
- [ ] Review job requirements
- [ ] Tailor resume
- [ ] Write cover letter
- [ ] Submit application
- [ ] Follow up

---
*Auto-generated from job scraping*
"""
            
            labels = ['job-opportunity', 'new-lead']
            if any(keyword in job.title.lower() for keyword in ['senior', 'lead', 'principal']):
                labels.append('senior-role')
            
            issue = self.email_forwarder.github_integration.repo.create_issue(
                title=title,
                body=body,
                labels=labels
            )
            
            logging.info(f"Created job issue #{issue.number}: {job.title}")
            
        except Exception as e:
            logging.error(f"Failed to create job issue: {e}")
            raise
    
    def _create_tailored_resumes(self) -> int:
        """Create tailored resumes for recent job opportunities"""
        # This would integrate with user profile and recent job postings
        # For now, return 0 as this requires user profile data
        logging.info("Resume tailoring would happen here with user profile data")
        return 0
    
    def check_configuration(self) -> Dict[str, bool]:
        """Check what components are properly configured"""
        status = {
            'email_integration': False,
            'job_scraping': False,
            'github_connection': False,
            'config_file_exists': False
        }
        
        # Check config file
        if os.path.exists(self.config_file):
            status['config_file_exists'] = True
        
        # Check email integration
        if self.email_forwarder:
            status['email_integration'] = True
            
            if (self.email_forwarder.github_integration and 
                self.email_forwarder.github_integration.github):
                status['github_connection'] = True
        
        # Check job scraping
        if self.job_scraper:
            status['job_scraping'] = True
        
        return status
    
    def print_status(self):
        """Print current workflow status"""
        status = self.check_configuration()
        
        print("\n🔍 Job Search Workflow Status")
        print("=" * 40)
        
        components = [
            ("Email Integration", status['email_integration']),
            ("Job Scraping", status['job_scraping']),
            ("GitHub Connection", status['github_connection']),
            ("Config File", status['config_file_exists'])
        ]
        
        for name, enabled in components:
            icon = "✅" if enabled else "❌"
            print(f"{icon} {name}")
        
        print("\n📋 Available Commands:")
        print("  workflow daily      - Run complete daily workflow")
        print("  workflow monitor    - Start continuous email monitoring")
        print("  workflow status     - Show this status")
        print("  workflow setup      - Create sample configuration")
        
        if not all(status.values()):
            print("\n⚠️  Setup required:")
            if not status['config_file_exists']:
                print("  - Run 'python job_email_workflow.py setup' to create config")
            if not status['email_integration']:
                print("  - Configure Gmail credentials in email_config.json")
            if not status['github_connection']:
                print("  - Add GitHub token to configuration")

def main():
    """Main entry point"""
    if len(sys.argv) < 2:
        print("Usage: python job_email_workflow.py <command>")
        print("Commands: daily, monitor, status, setup")
        sys.exit(1)
    
    command = sys.argv[1].lower()
    
    if command == "setup":
        # Create sample configuration
        if EMAIL_AVAILABLE:
            from email_integration import create_sample_config
            create_sample_config()
        else:
            print("Email integration not available. Install required dependencies.")
        return
    
    # Initialize workflow
    workflow = JobSearchWorkflow()
    
    if command == "status":
        workflow.print_status()
        
    elif command == "daily":
        stats = workflow.run_daily_workflow()
        print(f"\n📊 Daily Workflow Results:")
        print(f"  Emails processed: {stats['emails_processed']}")
        print(f"  Jobs found: {stats['jobs_found']}")
        print(f"  Issues created: {stats['issues_created']}")
        print(f"  Errors: {stats['errors']}")
        
    elif command == "monitor":
        workflow.run_email_monitoring()
        
    else:
        print(f"Unknown command: {command}")
        print("Available commands: daily, monitor, status, setup")
        sys.exit(1)

if __name__ == "__main__":
    main()
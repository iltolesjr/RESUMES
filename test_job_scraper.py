# test_job_scraper.py - Test the complete job scraping and resume tailoring workflow

import requests
import json
import time

BASE_URL = "http://localhost:5000"

def test_complete_workflow():
    """Test the complete job scraping and resume tailoring workflow"""
    
    # 1. Save enhanced user information
    print("1. Saving enhanced user information...")
    user_data = {
        "user_id": "test_user",
        "name": "John Doe",
        "email": "john.doe@email.com",
        "phone": "(555) 123-4567",
        "location": "Minneapolis, MN",
        "linkedin": "linkedin.com/in/johndoe",
        "github": "github.com/johndoe",
        "summary": "Experienced software engineer with 5+ years developing web applications",
        "skills": [
            "Python", "JavaScript", "React", "Node.js", "SQL", "AWS", 
            "Docker", "Git", "Agile", "REST APIs", "MongoDB", "PostgreSQL"
        ],
        "experience": [
            {
                "title": "Senior Software Engineer",
                "company": "Tech Corp",
                "duration": "2020 - Present",
                "description": "Led development of microservices architecture serving 1M+ users. Built React applications and Python APIs. Improved system performance by 40%."
            },
            {
                "title": "Software Developer",
                "company": "StartupXYZ",
                "duration": "2018 - 2020",
                "description": "Developed full-stack web applications using Python/Django and React. Implemented CI/CD pipelines and automated testing."
            }
        ],
        "education": "Bachelor of Science in Computer Science, University of Minnesota",
        "projects": [
            {
                "name": "E-commerce Platform",
                "description": "Built scalable e-commerce platform using React, Node.js, and MongoDB. Implemented payment processing and inventory management."
            },
            {
                "name": "Data Analytics Dashboard",
                "description": "Created real-time analytics dashboard using Python, Flask, and D3.js. Processed 100K+ daily events."
            }
        ],
        "certifications": "AWS Certified Solutions Architect",
        "target_roles": ["Software Engineer", "Full Stack Developer", "Backend Engineer"],
        "preferred_locations": ["Minneapolis", "Remote", "Chicago"]
    }
    
    response = requests.post(f"{BASE_URL}/save-enhanced-user-info", json=user_data)
    print(f"User info saved: {response.json()}")
    
    # 2. Scrape jobs
    print("\n2. Scraping jobs...")
    scrape_data = {
        "search_terms": ["software engineer", "python developer"],
        "location": "Minneapolis",
        "max_jobs": 10,
        "sources": ["indeed"]
    }
    
    response = requests.post(f"{BASE_URL}/scrape-jobs", json=scrape_data)
    jobs_result = response.json()
    print(f"Jobs scraped: {jobs_result.get('message', 'Error')}")
    
    if jobs_result.get('status') == 'success':
        jobs = jobs_result.get('jobs', [])
        print(f"Found {len(jobs)} jobs")
        
        # 3. Tailor resume for the first job
        if jobs:
            print(f"\n3. Tailoring resume for: {jobs[0]['title']} at {jobs[0]['company']}")
            
            tailor_data = {
                "job": jobs[0],
                "user_id": "test_user"
            }
            
            response = requests.post(f"{BASE_URL}/tailor-resume-for-job", json=tailor_data)
            tailor_result = response.json()
            
            if tailor_result.get('status') == 'success':
                print(f"Job match score: {tailor_result.get('job_match_score')}%")
                print(f"Matched keywords: {', '.join(tailor_result.get('matched_keywords', [])[:5])}")
                print(f"Skill demonstrations: {len(tailor_result.get('skill_demonstrations', []))} suggestions")
                print("\nTailored resume preview:")
                print(tailor_result.get('tailored_resume', '')[:500] + "...")
            else:
                print(f"Error tailoring resume: {tailor_result.get('message')}")
        
        # 4. Bulk tailor resumes for all jobs
        print(f"\n4. Bulk tailoring resumes for all {len(jobs)} jobs...")
        
        bulk_data = {
            "jobs": jobs,
            "user_id": "test_user"
        }
        
        response = requests.post(f"{BASE_URL}/bulk-tailor-resumes", json=bulk_data)
        bulk_result = response.json()
        
        if bulk_result.get('status') == 'success':
            tailored_resumes = bulk_result.get('tailored_resumes', [])
            print(f"\nGenerated {len(tailored_resumes)} tailored resumes")
            
            # Show top 3 matches
            print("\nTop 3 job matches:")
            for i, resume in enumerate(tailored_resumes[:3], 1):
                job_info = resume.get('job_info', {})
                match_score = resume.get('job_match_score', 0)
                print(f"{i}. {job_info.get('title')} at {job_info.get('company')} - {match_score}% match")
        else:
            print(f"Error bulk tailoring: {bulk_result.get('message')}")

if __name__ == "__main__":
    print("Testing Complete Job Scraping and Resume Tailoring Workflow")
    print("=" * 60)
    
    # Wait for server to be ready
    print("Waiting for server to be ready...")
    time.sleep(2)
    
    try:
        test_complete_workflow()
    except requests.exceptions.ConnectionError:
        print("Error: Could not connect to server. Make sure mcp_server.py is running.")
    except Exception as e:
        print(f"Test error: {e}")

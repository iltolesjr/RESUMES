import os
import yaml
from github import Github

def fetch_jobs():
    # Placeholder: Replace with actual API calls or scraping logic
    jobs = [
        {
            "title": "Data Analyst",
            "company": "TechCo",
            "location": "Minneapolis, MN",
            "description": "Analyze data and build dashboards."
        },
        {
            "title": "Project Manager",
            "company": "BizGroup",
            "location": "Remote",
            "description": "Manage projects and teams."
        }
    ]

    # Handshake job board example (placeholder)
    handshake_jobs = [
        {
            "title": "IT Support Specialist",
            "company": "Handshake Partner Co",
            "location": "Remote",
            "description": "Provide IT support for students and staff."
        }
    ]
    jobs.extend(handshake_jobs)
    with open('c:\\Users\\irato\\OneDrive - Minnesota State\\RESUMES\\RESUMES\\jobs.yaml', 'w') as f:
        yaml.dump(jobs, f)

def load_profile():
    profile_path = 'c:\\Users\\irato\\OneDrive - Minnesota State\\RESUMES\\RESUMES\\profile.yaml'
    if not os.path.exists(profile_path):
        print(f"ERROR: Profile file not found at {profile_path}. Please create it.")
        exit(1)
    with open(profile_path) as f:
        return yaml.safe_load(f)

def load_jobs():
    jobs_path = 'c:\\Users\\irato\\OneDrive - Minnesota State\\RESUMES\\RESUMES\\jobs.yaml'
    if not os.path.exists(jobs_path):
        print(f"ERROR: Jobs file not found at {jobs_path}. Please run fetch_jobs first.")
        exit(1)
    with open(jobs_path) as f:
        return yaml.safe_load(f)

def tailor_resume(profile, job):
    tailored_skills = [skill for skill in profile['skills'] if skill.lower() in job['description'].lower()]
    resume = f"# {profile['name']}\n\n"
    resume += f"**Email:** {profile['email']}\n"
    resume += f"**Location:** {profile['location']}\n\n"
    resume += "## Skills\n"
    resume += ", ".join(tailored_skills) + "\n\n"
    resume += "## Experience\n"
    for exp in profile['experience']:
        resume += f"**{exp['title']}** — {exp['company']} ({exp['years']} years)\n{exp['description']}\n\n"
    resume += f"## Summary\nSeeking: {job['title']} at {job['company']}\n"
    return resume

def create_issue(token, repo_name, job, resume_path):
    # Function starts here, previous broken line removed
    g = Github(token)
    repo = g.get_repo(repo_name)
    with open(resume_path, 'r') as f:
        resume_content = f.read()
    issue_title = f"Job Lead: {job['title']} at {job['company']}"
    issue_body = f"**Job Description:**\n{job['description']}\n\n**Tailored Resume:**\n```\n{resume_content}\n```"
    repo.create_issue(title=issue_title, body=issue_body)

def main():
    fetch_jobs()
    profile = load_profile()
    jobs = load_jobs()
    for job in jobs:
        resume = tailor_resume(profile, job)
        resume_path = f"c:\\Users\\irato\\OneDrive - Minnesota State\\RESUMES\\RESUMES\\documents\\{job['title'].replace(' ', '_')}_{job['company']}.md"
        os.makedirs(os.path.dirname(resume_path), exist_ok=True)
        with open(resume_path, 'w') as f:
            f.write(resume)
        token = os.getenv('GITHUB_TOKEN')
        repo_name = "iltolesjr/RESUMES"
        create_issue(token, repo_name, job, resume_path)

if __name__ == "__main__":
    main()

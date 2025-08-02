# job_scraper.py - New module for job scraping and enhanced resume tailoring

import requests
import time
import re
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse
import json
from typing import List, Dict, Optional
import logging
from dataclasses import dataclass
from datetime import datetime

@dataclass
class JobLead:
    title: str
    company: str
    location: str
    description: str
    requirements: str
    url: str
    keywords: List[str]
    salary_range: Optional[str] = None
    posted_date: Optional[str] = None
    application_deadline: Optional[str] = None

class JobScraper:
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        })
        
    def scrape_handshake_jobs(self, search_url: str, max_jobs: int = 25) -> List[JobLead]:
        """
        Scrape job listings from Handshake
        Note: This is a basic implementation - Handshake likely requires authentication
        """
        jobs = []
        try:
            response = self.session.get(search_url, timeout=10)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # These selectors would need to be updated based on actual Handshake HTML structure
            job_listings = soup.find_all('div', class_='job-listing') or soup.find_all('article')
            
            for job_elem in job_listings[:max_jobs]:
                try:
                    job = self._parse_handshake_job(job_elem, search_url)
                    if job:
                        jobs.append(job)
                except Exception as e:
                    logging.warning(f"Error parsing job listing: {e}")
                    continue
                    
        except Exception as e:
            logging.error(f"Error scraping Handshake: {e}")
            
        return jobs
    
    def _parse_handshake_job(self, job_elem, base_url: str) -> Optional[JobLead]:
        """Parse individual job listing from Handshake"""
        try:
            # These selectors need to be updated based on actual Handshake structure
            title = job_elem.find('h3') or job_elem.find('h2')
            title = title.get_text(strip=True) if title else "Unknown Title"
            
            company = job_elem.find(class_='company-name')
            company = company.get_text(strip=True) if company else "Unknown Company"
            
            location = job_elem.find(class_='location')
            location = location.get_text(strip=True) if location else "Remote"
            
            description_elem = job_elem.find(class_='description') or job_elem.find('p')
            description = description_elem.get_text(strip=True) if description_elem else ""
            
            # Extract job URL
            link_elem = job_elem.find('a', href=True)
            job_url = urljoin(base_url, link_elem['href']) if link_elem else base_url
            
            # Extract keywords from description
            keywords = self._extract_keywords(description + " " + title)
            
            return JobLead(
                title=title,
                company=company,
                location=location,
                description=description,
                requirements=description,  # Could be separate field
                url=job_url,
                keywords=keywords
            )
            
        except Exception as e:
            logging.error(f"Error parsing job element: {e}")
            return None
    
    def scrape_indeed_jobs(self, search_term: str, location: str = "", max_jobs: int = 25) -> List[JobLead]:
        """Scrape jobs from Indeed"""
        jobs = []
        base_url = "https://www.indeed.com"
        search_url = f"{base_url}/jobs?q={search_term}&l={location}"
        
        try:
            response = self.session.get(search_url, timeout=10)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Indeed job selectors (may need updates)
            job_listings = soup.find_all('div', class_='job_seen_beacon')
            
            for job_elem in job_listings[:max_jobs]:
                try:
                    job = self._parse_indeed_job(job_elem, base_url)
                    if job:
                        jobs.append(job)
                        time.sleep(1)  # Rate limiting
                except Exception as e:
                    logging.warning(f"Error parsing Indeed job: {e}")
                    continue
                    
        except Exception as e:
            logging.error(f"Error scraping Indeed: {e}")
            
        return jobs
    
    def _parse_indeed_job(self, job_elem, base_url: str) -> Optional[JobLead]:
        """Parse individual job from Indeed"""
        try:
            title_elem = job_elem.find('h2', class_='jobTitle')
            title = title_elem.find('a').get_text(strip=True) if title_elem and title_elem.find('a') else "Unknown Title"
            
            company_elem = job_elem.find('span', class_='companyName')
            company = company_elem.get_text(strip=True) if company_elem else "Unknown Company"
            
            location_elem = job_elem.find('div', class_='companyLocation')
            location = location_elem.get_text(strip=True) if location_elem else "Remote"
            
            # Get job URL
            link_elem = title_elem.find('a', href=True) if title_elem else None
            job_url = urljoin(base_url, link_elem['href']) if link_elem else base_url
            
            # Get full description by visiting job page
            description = self._get_indeed_job_description(job_url)
            
            keywords = self._extract_keywords(description + " " + title)
            
            return JobLead(
                title=title,
                company=company,
                location=location,
                description=description,
                requirements=description,
                url=job_url,
                keywords=keywords
            )
            
        except Exception as e:
            logging.error(f"Error parsing Indeed job element: {e}")
            return None
    
    def _get_indeed_job_description(self, job_url: str) -> str:
        """Get full job description from Indeed job page"""
        try:
            response = self.session.get(job_url, timeout=10)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.content, 'html.parser')
            desc_elem = soup.find('div', class_='jobsearch-jobDescriptionText')
            
            return desc_elem.get_text(strip=True) if desc_elem else ""
            
        except Exception as e:
            logging.warning(f"Could not fetch job description from {job_url}: {e}")
            return ""
    
    def _extract_keywords(self, text: str) -> List[str]:
        """Extract relevant keywords from job text"""
        # Common tech keywords and skills
        tech_keywords = [
            'python', 'javascript', 'java', 'react', 'node.js', 'sql', 'aws', 'docker',
            'kubernetes', 'git', 'agile', 'scrum', 'devops', 'ci/cd', 'rest api',
            'machine learning', 'data analysis', 'flask', 'django', 'mongodb',
            'postgresql', 'redis', 'microservices', 'cloud', 'azure', 'gcp',
            'html', 'css', 'typescript', 'angular', 'vue.js', 'bootstrap',
            'webpack', 'npm', 'yarn', 'testing', 'unit testing', 'integration testing'
        ]
        
        # Business/soft skills keywords
        soft_keywords = [
            'leadership', 'communication', 'teamwork', 'project management',
            'problem solving', 'analytical', 'creative', 'detail oriented',
            'collaboration', 'mentoring', 'stakeholder management', 'cross-functional'
        ]
        
        all_keywords = tech_keywords + soft_keywords
        text_lower = text.lower()
        
        found_keywords = []
        for keyword in all_keywords:
            if keyword.lower() in text_lower:
                found_keywords.append(keyword)
        
        # Also extract degree requirements, years of experience, etc.
        experience_match = re.search(r'(\d+)\+?\s*years?\s*(of\s+)?experience', text_lower)
        if experience_match:
            found_keywords.append(f"{experience_match.group(1)}+ years experience")
        
        degree_match = re.search(r'(bachelor|master|phd|doctorate).*degree', text_lower)
        if degree_match:
            found_keywords.append(degree_match.group(0))
        
        return list(set(found_keywords))  # Remove duplicates

class EnhancedResumeProcessor:
    """Enhanced resume processor with job matching capabilities"""
    
    def __init__(self):
        self.base_resume_template = self._load_base_resume()
        
    def _load_base_resume(self) -> str:
        """Load base resume template"""
        return """
{{ name }}
{{ email }} | {{ phone }} | {{ location }}
{{ linkedin }} | {{ github }}

PROFESSIONAL SUMMARY
{{ summary }}

TECHNICAL SKILLS
{{ technical_skills }}

PROFESSIONAL EXPERIENCE
{{ experience }}

EDUCATION
{{ education }}

PROJECTS
{{ projects }}

CERTIFICATIONS
{{ certifications }}
"""
    
    def tailor_resume_for_job(self, job: JobLead, user_info: Dict) -> Dict:
        """Create a tailored resume for a specific job"""
        try:
            # Extract user's base information
            base_skills = user_info.get('skills', [])
            base_experience = user_info.get('experience', [])
            base_projects = user_info.get('projects', [])
            
            # Match job keywords with user skills
            matched_skills = self._match_skills_to_job(base_skills, job.keywords)
            
            # Generate tailored summary
            tailored_summary = self._generate_tailored_summary(job, user_info)
            
            # Prioritize relevant experience
            prioritized_experience = self._prioritize_experience(base_experience, job.keywords)
            
            # Suggest skill demonstrations
            skill_demonstrations = self._suggest_skill_demonstrations(job.keywords, base_experience, base_projects)
            
            # Create the tailored resume
            tailored_resume = self.base_resume_template.format(
                name=user_info.get('name', '{{ name }}'),
                email=user_info.get('email', '{{ email }}'),
                phone=user_info.get('phone', '{{ phone }}'),
                location=user_info.get('location', '{{ location }}'),
                linkedin=user_info.get('linkedin', '{{ linkedin }}'),
                github=user_info.get('github', '{{ github }}'),
                summary=tailored_summary,
                technical_skills=self._format_skills(matched_skills),
                experience=self._format_experience(prioritized_experience),
                education=user_info.get('education', '{{ education }}'),
                projects=self._format_projects(base_projects, job.keywords),
                certifications=user_info.get('certifications', '{{ certifications }}')
            )
            
            return {
                'status': 'success',
                'tailored_resume': tailored_resume,
                'job_match_score': self._calculate_match_score(user_info, job),
                'matched_keywords': matched_skills,
                'skill_demonstrations': skill_demonstrations,
                'optimization_suggestions': self._get_optimization_suggestions(job, user_info)
            }
            
        except Exception as e:
            logging.error(f"Error tailoring resume: {e}")
            return {
                'status': 'error',
                'message': f'Error tailoring resume: {str(e)}',
                'error_type': 'exception'
            }
    
    def _match_skills_to_job(self, user_skills: List[str], job_keywords: List[str]) -> List[str]:
        """Match user skills to job requirements"""
        matched = []
        user_skills_lower = [skill.lower() for skill in user_skills]
        
        for keyword in job_keywords:
            for user_skill in user_skills:
                if keyword.lower() in user_skill.lower() or user_skill.lower() in keyword.lower():
                    if user_skill not in matched:
                        matched.append(user_skill)
        
        return matched
    
    def _generate_tailored_summary(self, job: JobLead, user_info: Dict) -> str:
        """Generate a tailored professional summary"""
        base_summary = user_info.get('summary', '')
        
        # Extract key requirements from job
        years_exp = next((kw for kw in job.keywords if 'years experience' in kw), '')
        key_skills = [kw for kw in job.keywords if kw in ['python', 'javascript', 'react', 'aws', 'sql']][:3]
        
        tailored_summary = f"Experienced {job.title.lower()} professional"
        if years_exp:
            tailored_summary += f" with {years_exp}"
        
        if key_skills:
            tailored_summary += f" specializing in {', '.join(key_skills)}"
        
        tailored_summary += f". Passionate about {job.company}'s mission and eager to contribute to innovative solutions."
        
        return tailored_summary
    
    def _prioritize_experience(self, experiences: List[Dict], job_keywords: List[str]) -> List[Dict]:
        """Prioritize experiences based on job keywords"""
        scored_experiences = []
        
        for exp in experiences:
            score = 0
            exp_text = f"{exp.get('title', '')} {exp.get('description', '')}"
            
            for keyword in job_keywords:
                if keyword.lower() in exp_text.lower():
                    score += 1
            
            scored_experiences.append((score, exp))
        
        # Sort by score (descending) and return experiences
        sorted_experiences = sorted(scored_experiences, key=lambda x: x[0], reverse=True)
        return [exp for score, exp in sorted_experiences]
    
    def _suggest_skill_demonstrations(self, job_keywords: List[str], experiences: List[Dict], projects: List[Dict]) -> List[str]:
        """Suggest creative ways to demonstrate skills"""
        suggestions = []
        
        skill_demonstrations = {
            'python': [
                "Automated data processing workflows reducing manual effort by 80%",
                "Built REST APIs serving 10,000+ daily requests",
                "Developed machine learning models for predictive analysis"
            ],
            'javascript': [
                "Created interactive dashboards improving user engagement by 40%",
                "Implemented real-time features using WebSocket connections",
                "Optimized frontend performance reducing load times by 60%"
            ],
            'react': [
                "Built responsive single-page applications with 99.9% uptime",
                "Implemented component libraries reducing development time by 50%",
                "Created mobile-first interfaces serving 50,000+ users"
            ],
            'aws': [
                "Architected cloud infrastructure supporting 1M+ transactions",
                "Implemented CI/CD pipelines reducing deployment time by 70%",
                "Optimized cloud costs saving $10,000+ annually"
            ],
            'leadership': [
                "Led cross-functional teams of 5+ developers to deliver projects on time",
                "Mentored junior developers improving team productivity by 30%",
                "Coordinated with stakeholders to align technical solutions with business goals"
            ]
        }
        
        for keyword in job_keywords:
            if keyword.lower() in skill_demonstrations:
                suggestions.extend(skill_demonstrations[keyword.lower()][:2])
        
        return suggestions[:5]  # Limit to top 5 suggestions
    
    def _calculate_match_score(self, user_info: Dict, job: JobLead) -> float:
        """Calculate how well user matches the job"""
        user_skills = user_info.get('skills', [])
        user_skills_lower = [skill.lower() for skill in user_skills]
        
        matched_keywords = 0
        total_keywords = len(job.keywords)
        
        for keyword in job.keywords:
            if any(keyword.lower() in skill for skill in user_skills_lower):
                matched_keywords += 1
        
        return round((matched_keywords / max(total_keywords, 1)) * 100, 2)
    
    def _get_optimization_suggestions(self, job: JobLead, user_info: Dict) -> List[str]:
        """Get suggestions for improving job match"""
        suggestions = []
        user_skills = [skill.lower() for skill in user_info.get('skills', [])]
        
        missing_skills = []
        for keyword in job.keywords:
            if not any(keyword.lower() in skill for skill in user_skills):
                missing_skills.append(keyword)
        
        if missing_skills:
            suggestions.append(f"Consider highlighting experience with: {', '.join(missing_skills[:3])}")
        
        suggestions.append("Add quantifiable achievements to your experience descriptions")
        suggestions.append("Include relevant coursework or certifications")
        suggestions.append("Customize your summary to align with company values")
        
        return suggestions
    
    def _format_skills(self, skills: List[str]) -> str:
        """Format skills section"""
        if not skills:
            return "{{ technical_skills }}"
        
        # Group skills by category
        languages = [s for s in skills if s.lower() in ['python', 'javascript', 'java', 'typescript', 'sql']]
        frameworks = [s for s in skills if s.lower() in ['react', 'django', 'flask', 'node.js', 'angular', 'vue.js']]
        tools = [s for s in skills if s.lower() in ['aws', 'docker', 'kubernetes', 'git', 'mongodb', 'postgresql']]
        
        formatted = ""
        if languages:
            formatted += f"Languages: {', '.join(languages)}\n"
        if frameworks:
            formatted += f"Frameworks: {', '.join(frameworks)}\n"
        if tools:
            formatted += f"Tools & Technologies: {', '.join(tools)}\n"
        
        return formatted
    
    def _format_experience(self, experiences: List[Dict]) -> str:
        """Format experience section"""
        if not experiences:
            return "{{ experience }}"
        
        formatted = ""
        for exp in experiences:
            formatted += f"{exp.get('title', 'Position')} | {exp.get('company', 'Company')}\n"
            formatted += f"{exp.get('duration', 'Duration')}\n"
            formatted += f"• {exp.get('description', 'Description')}\n\n"
        
        return formatted
    
    def _format_projects(self, projects: List[Dict], job_keywords: List[str]) -> str:
        """Format projects section with job-relevant projects first"""
        if not projects:
            return "{{ projects }}"
        
        # Score projects by relevance
        scored_projects = []
        for project in projects:
            score = 0
            project_text = f"{project.get('name', '')} {project.get('description', '')}"
            
            for keyword in job_keywords:
                if keyword.lower() in project_text.lower():
                    score += 1
            
            scored_projects.append((score, project))
        
        # Sort by relevance
        sorted_projects = sorted(scored_projects, key=lambda x: x[0], reverse=True)
        
        formatted = ""
        for score, project in sorted_projects:
            formatted += f"{project.get('name', 'Project')}\n"
            formatted += f"• {project.get('description', 'Description')}\n\n"
        
        return formatted

# Integration with existing MCP server
def add_job_scraping_endpoints(app, user_info_storage):
    """Add job scraping endpoints to existing Flask app"""
    
    job_scraper = JobScraper()
    resume_processor = EnhancedResumeProcessor()
    
    @app.route('/scrape-jobs', methods=['POST'])
    def scrape_jobs():
        try:
            data = request.get_json()
            search_terms = data.get('search_terms', [])
            location = data.get('location', '')
            max_jobs = min(data.get('max_jobs', 25), 50)  # Limit to 50 jobs
            sources = data.get('sources', ['indeed'])  # Default to Indeed
            
            all_jobs = []
            
            if 'indeed' in sources:
                for term in search_terms:
                    jobs = job_scraper.scrape_indeed_jobs(term, location, max_jobs // len(search_terms))
                    all_jobs.extend(jobs)
            
            if 'handshake' in sources and data.get('handshake_url'):
                handshake_jobs = job_scraper.scrape_handshake_jobs(data['handshake_url'], max_jobs)
                all_jobs.extend(handshake_jobs)
            
            # Convert JobLead objects to dictionaries
            jobs_data = []
            for job in all_jobs:
                jobs_data.append({
                    'title': job.title,
                    'company': job.company,
                    'location': job.location,
                    'description': job.description[:500] + "..." if len(job.description) > 500 else job.description,
                    'keywords': job.keywords,
                    'url': job.url,
                    'scraped_at': datetime.now().isoformat()
                })
            
            return jsonify({
                'status': 'success',
                'message': f'Found {len(jobs_data)} job opportunities',
                'jobs': jobs_data,
                'total_jobs': len(jobs_data)
            })
            
        except Exception as e:
            logging.error(f"Error in scrape_jobs: {e}")
            return jsonify({
                'status': 'error',
                'message': f'Error scraping jobs: {str(e)}',
                'error_type': 'exception'
            }), 500
    
    @app.route('/tailor-resume-for-job', methods=['POST'])
    def tailor_resume_for_job():
        try:
            data = request.get_json()
            job_data = data.get('job')
            user_id = data.get('user_id', 'default')
            
            if not job_data:
                return jsonify({
                    'status': 'error',
                    'message': 'Job data is required',
                    'error_type': 'validation'
                }), 400
            
            # Get user info
            user_info = user_info_storage.get(user_id, {})
            if not user_info:
                return jsonify({
                    'status': 'error',
                    'message': 'User info not found. Please save user info first.',
                    'error_type': 'validation'
                }), 400
            
            # Create JobLead object from data
            job = JobLead(
                title=job_data.get('title', ''),
                company=job_data.get('company', ''),
                location=job_data.get('location', ''),
                description=job_data.get('description', ''),
                requirements=job_data.get('requirements', ''),
                url=job_data.get('url', ''),
                keywords=job_data.get('keywords', [])
            )
            
            # Tailor the resume
            result = resume_processor.tailor_resume_for_job(job, user_info)
            
            return jsonify(result)
            
        except Exception as e:
            logging.error(f"Error in tailor_resume_for_job: {e}")
            return jsonify({
                'status': 'error',
                'message': f'Error tailoring resume: {str(e)}',
                'error_type': 'exception'
            }), 500
    
    @app.route('/bulk-tailor-resumes', methods=['POST'])
    def bulk_tailor_resumes():
        try:
            data = request.get_json()
            jobs_data = data.get('jobs', [])
            user_id = data.get('user_id', 'default')
            
            if not jobs_data:
                return jsonify({
                    'status': 'error',
                    'message': 'Jobs data is required',
                    'error_type': 'validation'
                }), 400
            
            # Get user info
            user_info = user_info_storage.get(user_id, {})
            if not user_info:
                return jsonify({
                    'status': 'error',
                    'message': 'User info not found. Please save user info first.',
                    'error_type': 'validation'
                }), 400
            
            tailored_resumes = []
            
            for job_data in jobs_data[:10]:  # Limit to 10 jobs for performance
                job = JobLead(
                    title=job_data.get('title', ''),
                    company=job_data.get('company', ''),
                    location=job_data.get('location', ''),
                    description=job_data.get('description', ''),
                    requirements=job_data.get('requirements', ''),
                    url=job_data.get('url', ''),
                    keywords=job_data.get('keywords', [])
                )
                
                result = resume_processor.tailor_resume_for_job(job, user_info)
                result['job_info'] = {
                    'title': job.title,
                    'company': job.company,
                    'url': job.url
                }
                tailored_resumes.append(result)
            
            # Sort by match score
            tailored_resumes.sort(key=lambda x: x.get('job_match_score', 0), reverse=True)
            
            return jsonify({
                'status': 'success',
                'message': f'Generated {len(tailored_resumes)} tailored resumes',
                'tailored_resumes': tailored_resumes
            })
            
        except Exception as e:
            logging.error(f"Error in bulk_tailor_resumes: {e}")
            return jsonify({
                'status': 'error',
                'message': f'Error bulk tailoring resumes: {str(e)}',
                'error_type': 'exception'
            }), 500

if __name__ == "__main__":
    # Test the scraper
    scraper = JobScraper()
    processor = EnhancedResumeProcessor()
    
    # Test Indeed scraping
    print("Testing Indeed job scraping...")
    jobs = scraper.scrape_indeed_jobs("software engineer", "Minneapolis", 5)
    
    for job in jobs:
        print(f"\nTitle: {job.title}")
        print(f"Company: {job.company}")
        print(f"Keywords: {job.keywords[:5]}")
        print(f"URL: {job.url}")

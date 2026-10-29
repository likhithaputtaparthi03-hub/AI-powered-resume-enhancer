from flask import Flask, render_template, request, send_file
from io import BytesIO
import PyPDF2
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from xml.sax.saxutils import escape

app = Flask(__name__)


# =========================================================
# PDF TEXT EXTRACTION
# =========================================================

def extract_text_from_pdf(file):

    reader = PyPDF2.PdfReader(file)

    text = ""

    for page in reader.pages:
        page_text = page.extract_text()

        if page_text:
            text += page_text + "\n"

    return text


# =========================================================
# KEYWORD FUNCTIONS
# =========================================================

def keyword_in_text(keyword, text):

    return bool(
        re.search(
            r"\b" + re.escape(keyword) + r"\b",
            text,
            re.IGNORECASE
        )
    )


def extract_technologies(text):

    technologies = [
        "Python",
        "Java",
        "C++",
        "C",
        "SQL",
        "HTML",
        "CSS",
        "JavaScript",
        "React",
        "MySQL",
        "Git",
        "GitHub",
        "Machine Learning",
        "Data Structures",
        "Algorithms",
        "Flask",
        "Django",
        "Power BI",
        "Excel"
    ]

    found = []

    for technology in technologies:

        if keyword_in_text(technology, text):
            found.append(technology)

    return found


def extract_job_keywords(job_description):

    if not job_description:
        return []

    common_keywords = [
        "Python",
        "Java",
        "C++",
        "C",
        "SQL",
        "HTML",
        "CSS",
        "JavaScript",
        "React",
        "MySQL",
        "Git",
        "GitHub",
        "Machine Learning",
        "Data Structures",
        "Algorithms",
        "Flask",
        "Django",
        "Power BI",
        "Excel",
        "AWS",
        "Cloud",
        "Communication",
        "Leadership",
        "Problem Solving",
        "Teamwork",
        "API",
        "REST API"
    ]

    found = []

    for keyword in common_keywords:

        if keyword_in_text(keyword, job_description):
            found.append(keyword)

    return found


# =========================================================
# RESUME ANALYSIS
# =========================================================

def analyze_resume(text, job_description=""):

    sections = {

        "Contact information": bool(
            re.search(
                r"[\w\.-]+@[\w\.-]+\.\w+|\+?\d[\d\s-]{8,}",
                text
            )
        ),

        "Career summary/objective": bool(
            re.search(
                r"\b(summary|objective|profile|career objective)\b",
                text,
                re.IGNORECASE
            )
        ),

        "Skills": bool(
            re.search(
                r"\b(skills|technical skills|technologies)\b",
                text,
                re.IGNORECASE
            )
        ),

        "Education": bool(
            re.search(
                r"\b(education|degree|b\.?tech|bachelor|college|university)\b",
                text,
                re.IGNORECASE
            )
        ),

        "Projects": bool(
            re.search(
                r"\b(project|projects)\b",
                text,
                re.IGNORECASE
            )
        ),

        "Experience": bool(
            re.search(
                r"\b(experience|internship|internships|work experience)\b",
                text,
                re.IGNORECASE
            )
        ),

        "Certifications": bool(
            re.search(
                r"\b(certification|certifications|certificate|certificates)\b",
                text,
                re.IGNORECASE
            )
        ),

        "LinkedIn profile": bool(
            re.search(
                r"linkedin\.com",
                text,
                re.IGNORECASE
            )
        ),

        "GitHub profile": bool(
            re.search(
                r"github\.com",
                text,
                re.IGNORECASE
            )
        )
    }


    found_sections = [
        name for name, found in sections.items()
        if found
    ]

    missing_sections = [
        name for name, found in sections.items()
        if not found
    ]


    # =====================================================
    # RESUME SCORE
    # =====================================================

    score = len(found_sections) * 10

    word_count = len(text.split())

    if word_count >= 150:
        score += 10

    score = min(score, 100)


    # =====================================================
    # GENERAL SUGGESTIONS
    # =====================================================

    suggestions = []

    if not sections["Contact information"]:
        suggestions.append(
            "Add your email address and phone number."
        )

    if not sections["Career summary/objective"]:
        suggestions.append(
            "Add a professional career summary."
        )

    if not sections["Skills"]:
        suggestions.append(
            "Add a dedicated technical skills section."
        )

    if not sections["Education"]:
        suggestions.append(
            "Add your educational qualifications."
        )

    if not sections["Projects"]:
        suggestions.append(
            "Add academic or personal projects."
        )

    if not sections["Experience"]:
        suggestions.append(
            "Add internships, training, or relevant practical experience if you have them."
        )

    if not sections["Certifications"]:
        suggestions.append(
            "Add relevant certifications."
        )

    if not sections["LinkedIn profile"]:
        suggestions.append(
            "Add your LinkedIn profile."
        )

    if not sections["GitHub profile"]:
        suggestions.append(
            "Add your GitHub profile."
        )


    # =====================================================
    # ATS KEYWORD ANALYSIS
    # =====================================================

    ats_keywords = [
        "Python",
        "Java",
        "C++",
        "SQL",
        "HTML",
        "CSS",
        "JavaScript",
        "React",
        "MySQL",
        "Git",
        "GitHub",
        "Machine Learning",
        "Data Structures",
        "Algorithms"
    ]

    found_keywords = []
    missing_keywords = []

    for keyword in ats_keywords:

        if keyword_in_text(keyword, text):
            found_keywords.append(keyword)

        else:
            missing_keywords.append(keyword)


    if ats_keywords:

        ats_score = int(
            len(found_keywords) /
            len(ats_keywords) * 100
        )

    else:

        ats_score = 0


    # =====================================================
    # JOB MATCH
    # =====================================================

    job_keywords = extract_job_keywords(
        job_description
    )

    matched_job_keywords = []
    missing_job_keywords = []

    for keyword in job_keywords:

        if keyword_in_text(keyword, text):
            matched_job_keywords.append(keyword)

        else:
            missing_job_keywords.append(keyword)


    if job_keywords:

        job_match = int(
            len(matched_job_keywords) /
            len(job_keywords) * 100
        )

    else:

        # If no job description was provided,
        # show a neutral value instead of pretending
        # that a match was calculated.
        job_match = 0


    job_suggestions = []

    for keyword in missing_job_keywords:

        job_suggestions.append(
            f"Consider adding '{keyword}' if it genuinely reflects your skills."
        )


    # =====================================================
    # SMART FEEDBACK
    # =====================================================

    smart_feedback = generate_smart_feedback(
        score,
        ats_score,
        job_match,
        word_count
    )


    return {

        "score": score,

        "ats_score": ats_score,

        "job_match": job_match,

        "sections": sections,

        "found_sections": found_sections,

        "missing_sections": missing_sections,

        "suggestions": suggestions,

        "found_keywords": found_keywords,

        "missing_keywords": missing_keywords,

        "matched_job_keywords": matched_job_keywords,

        "missing_job_keywords": missing_job_keywords,

        "job_suggestions": job_suggestions,

        "smart_feedback": smart_feedback,

        "resume_text": text,

        "job_description": job_description,

        "word_count": word_count
    }


# =========================================================
# SMART FEEDBACK
# =========================================================

def generate_smart_feedback(
    score,
    ats_score,
    job_match,
    word_count
):

    feedback = []

    if score >= 80:

        feedback.append(
            "Your resume contains most important sections."
        )

    elif score >= 60:

        feedback.append(
            "Your resume has a good basic structure but can be improved."
        )

    else:

        feedback.append(
            "Your resume needs additional sections and stronger content."
        )


    if ats_score >= 70:

        feedback.append(
            "Your resume contains several relevant ATS keywords."
        )

    else:

        feedback.append(
            "Add relevant technical keywords naturally to improve ATS compatibility."
        )


    if job_match >= 70:

        feedback.append(
            "Your resume contains several keywords related to the provided job description."
        )

    elif job_match > 0:

        feedback.append(
            "Consider adding relevant skills from the job description where they accurately represent your abilities."
        )


    if word_count < 150:

        feedback.append(
            "Your resume content is relatively short. Add meaningful project, education, and experience details."
        )


    return feedback


# =========================================================
# TEXT CLEANING
# =========================================================

def clean_sentence(sentence):
    """Clean text extracted from PDF files."""
    sentence = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", sentence)
    sentence = sentence.replace("\ufeff", "").replace("\u00a0", " ")
    sentence = sentence.strip()
    sentence = re.sub(r"^[•\-\*\d\.\)\s]+", "", sentence)
    sentence = re.sub(r"\s+", " ", sentence)
    return sentence.strip()


def normalize_resume_text(text):
    """Normalize PDF extraction while preserving section boundaries."""
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)
    text = text.replace("\ufeff", "").replace("\u00a0", " ")
    return text.replace("\r\n", "\n").replace("\r", "\n")


# =========================================================
# PROJECT ENHANCEMENT
# =========================================================

def enhance_project_sentence(sentence):
    sentence = clean_sentence(sentence)
    if not sentence:
        return ""
    return sentence.rstrip(".") + "."


def enhance_project(text):
    """Group project titles with all of their description lines.

    PDF extraction can split a project description into fragments such as
    ``content and presentation.``. Those fragments must stay with the
    project instead of being interpreted as a new project title.
    """
    action_words = (
        "built ", "developed ", "created ", "designed ",
        "implemented ", "developing ", "used ", "worked ",
        "managed ", "allows ", "uses ", "provides ",
        "helps ", "contains ", "includes "
    )

    projects = []
    current_title = None
    current_details = []

    for raw_line in normalize_resume_text(text).splitlines():
        line = clean_sentence(raw_line)
        if not line:
            continue

        lower = line.lower()

        # A project title in the supplied resume is a short, title-cased
        # line without sentence punctuation. Description fragments usually
        # end in punctuation or begin with a lowercase word.
        looks_like_title = (
            current_title is not None and
            len(line) <= 60 and
            line[0].isupper() and
            not re.search(r"[.!?]$", line) and
            not lower.startswith(action_words)
        )

        if current_title is None:
            current_title = line
        elif looks_like_title:
            projects.append((current_title, current_details))
            current_title = line
            current_details = []
        else:
            current_details.append(line)

    if current_title:
        projects.append((current_title, current_details))

    enhanced = []
    for title, details in projects:
        title = title.rstrip(".")
        cleaned_details = [d.rstrip(".") for d in details if d.strip()]

        if cleaned_details:
            description = " ".join(cleaned_details) + "."
            enhanced.append(f"• {title} — {description}")
        else:
            enhanced.append(f"• {title}")

    return "\n".join(enhanced)


# =========================================================
# EXPERIENCE ENHANCEMENT
# =========================================================

def enhance_experience(text):
    enhanced = []
    for line in normalize_resume_text(text).splitlines():
        line = clean_sentence(line)
        if not line:
            continue
        enhanced.append("• " + line.rstrip(".") + ".")
    return "\n".join(enhanced)


# =========================================================
# SUMMARY ENHANCEMENT
# =========================================================

def enhance_summary(text):
    text = clean_sentence(text)
    if not text:
        return ""

    technologies = extract_technologies(text)
    technology_text = ""
    if technologies:
        technology_text = " Hands-on knowledge of " + ", ".join(technologies) + "."

    if "student" in text.lower():
        return (
            "Motivated Computer Science student with a strong interest in software development and technology."
            + technology_text
            + " Passionate about problem solving, learning new technologies, and building practical solutions."
        )

    return (
        text.rstrip(".")
        + ". Demonstrates strong interest in software development, problem solving, and continuous learning."
        + technology_text
    )


# =========================================================
# SKILLS ENHANCEMENT
# =========================================================

def enhance_skills(text):
    technologies = extract_technologies(text)
    if technologies:
        return "Technical Skills\n" + ", ".join(technologies)
    return "Technical Skills\n" + clean_sentence(text)


# =========================================================
# EDUCATION ENHANCEMENT
# =========================================================

def enhance_education(text):
    lines = []
    for line in normalize_resume_text(text).splitlines():
        line = clean_sentence(line)
        if line:
            lines.append("• " + line)
    return "\n".join(lines)


# =========================================================
# CERTIFICATION / STRENGTHS / DECLARATION
# =========================================================

def enhance_certifications(text):
    lines = []
    for line in normalize_resume_text(text).splitlines():
        line = clean_sentence(line)
        if line:
            lines.append("• " + line)
    return "\n".join(lines)


def enhance_strengths(text):
    cleaned = clean_sentence(text)
    return "• " + cleaned if cleaned else ""


def enhance_declaration(text):
    cleaned = clean_sentence(text)
    return cleaned if cleaned else ""


# =========================================================
# AUTOMATIC SECTION DETECTION
# =========================================================

SECTION_ALIASES = {
    "Summary": ["summary", "professional summary", "career summary", "objective", "career objective", "profile", "professional profile"],
    "Skills": ["skills", "technical skills", "technical skill", "technologies", "technical expertise"],
    "Education": ["education", "educational qualifications", "academic qualifications", "academic background"],
    "Projects": ["projects", "project", "academic projects", "personal projects"],
    "Experience": ["experience", "work experience", "professional experience", "internship", "internships"],
    "Certifications": ["certifications", "certification", "certificates", "certificate"],
    "Strengths": ["strengths", "strength"],
    "Declaration": ["declaration"]
}


def normalize_heading(line):
    line = clean_sentence(line)
    line = re.sub(r"[:\-]+$", "", line)
    return line.lower().strip()


def detect_section_heading(line):
    normalized = normalize_heading(line)
    for section, aliases in SECTION_ALIASES.items():
        for alias in aliases:
            if normalized == alias:
                return section
    return None


def detect_resume_sections(text):
    """Detect sections while preventing later headings from being absorbed."""
    sections = {}
    current_section = None

    for raw_line in normalize_resume_text(text).splitlines():
        line = raw_line.strip()
        if not line:
            continue

        detected = detect_section_heading(line)
        if detected:
            current_section = detected
            sections.setdefault(current_section, [])
            continue

        if current_section:
            sections[current_section].append(line)

    return {
        section: "\n".join(content).strip()
        for section, content in sections.items()
        if content
    }


# =========================================================
# AUTOMATIC ENHANCEMENT
# =========================================================

def enhance_detected_sections(text):
    detected_sections = detect_resume_sections(text)
    enhanced_sections = {}

    for section, content in detected_sections.items():
        if section == "Summary":
            enhanced_sections[section] = enhance_summary(content)
        elif section == "Skills":
            enhanced_sections[section] = enhance_skills(content)
        elif section == "Education":
            enhanced_sections[section] = enhance_education(content)
        elif section == "Projects":
            enhanced_sections[section] = enhance_project(content)
        elif section == "Experience":
            enhanced_sections[section] = enhance_experience(content)
        elif section == "Certifications":
            enhanced_sections[section] = enhance_certifications(content)
        elif section == "Strengths":
            enhanced_sections[section] = enhance_strengths(content)
        elif section == "Declaration":
            enhanced_sections[section] = enhance_declaration(content)

    return {k: v for k, v in enhanced_sections.items() if v.strip()}


# =========================================================
# AI SKILL ENHANCEMENT
# =========================================================

def create_skill_roadmap(missing_skills):

    skill_details = {

        "Python": {
            "description": "Learn Python syntax, functions, object-oriented programming, file handling, modules, exceptions, and commonly used libraries.",
            "duration": "2-3 weeks",
            "project": "Build a Student Management System using Python.",
            "priority": "High Priority"
        },

        "Java": {
            "description": "Learn Java fundamentals, OOP concepts, collections, exception handling, and basic application development.",
            "duration": "3-4 weeks",
            "project": "Build a Library Management System using Java.",
            "priority": "High Priority"
        },

        "C++": {
            "description": "Learn C++ syntax, functions, OOP, STL, vectors, maps, and problem solving with C++.",
            "duration": "3-4 weeks",
            "project": "Build a Console-Based Student Record System using C++.",
            "priority": "High Priority"
        },

        "C": {
            "description": "Learn C programming fundamentals, functions, arrays, pointers, structures, and memory concepts.",
            "duration": "2-3 weeks",
            "project": "Build a Console-Based Banking System using C.",
            "priority": "Medium Priority"
        },

        "SQL": {
            "description": "Learn databases, SELECT queries, filtering, joins, grouping, subqueries, constraints, and basic database design.",
            "duration": "1-2 weeks",
            "project": "Create a Student Database and write queries to analyze student records.",
            "priority": "High Priority"
        },

        "HTML": {
            "description": "Learn HTML structure, semantic elements, forms, tables, links, images, and accessible page organization.",
            "duration": "4-7 days",
            "project": "Create a Professional Portfolio Website.",
            "priority": "Medium Priority"
        },

        "CSS": {
            "description": "Learn selectors, box model, positioning, Flexbox, Grid, responsive design, and modern page styling.",
            "duration": "1-2 weeks",
            "project": "Design a Responsive Student Dashboard.",
            "priority": "Medium Priority"
        },

        "JavaScript": {
            "description": "Learn variables, functions, arrays, objects, DOM manipulation, events, asynchronous JavaScript, and APIs.",
            "duration": "2-3 weeks",
            "project": "Build an Interactive To-Do List with local storage.",
            "priority": "High Priority"
        },

        "React": {
            "description": "Learn components, JSX, props, state, hooks, event handling, forms, routing, and API integration.",
            "duration": "3-4 weeks",
            "project": "Build a Student Management Dashboard using React.",
            "priority": "High Priority"
        },

        "MySQL": {
            "description": "Learn MySQL database creation, tables, relationships, joins, indexes, constraints, and CRUD operations.",
            "duration": "1-2 weeks",
            "project": "Build a College Student Database with MySQL.",
            "priority": "High Priority"
        },

        "Git": {
            "description": "Learn repositories, commits, branches, merging, pull requests, resolving conflicts, and basic Git workflows.",
            "duration": "3-5 days",
            "project": "Manage one of your college projects using Git with multiple branches.",
            "priority": "High Priority"
        },

        "GitHub": {
            "description": "Learn repositories, README files, commits, branches, pull requests, issues, and publishing projects on GitHub.",
            "duration": "3-5 days",
            "project": "Create a professional GitHub repository for your resume project.",
            "priority": "High Priority"
        },

        "Machine Learning": {
            "description": "Learn supervised and unsupervised learning, preprocessing, feature engineering, model training, evaluation, and basic ML algorithms.",
            "duration": "4-6 weeks",
            "project": "Build a Student Performance Prediction model.",
            "priority": "High Priority"
        },

        "Data Structures": {
            "description": "Learn arrays, linked lists, stacks, queues, hash tables, trees, heaps, graphs, and their common operations.",
            "duration": "4-6 weeks",
            "project": "Build a Data Structures Visualizer for common structures.",
            "priority": "High Priority"
        },

        "Algorithms": {
            "description": "Learn searching, sorting, recursion, greedy algorithms, dynamic programming, graph algorithms, and complexity analysis.",
            "duration": "4-6 weeks",
            "project": "Build an Algorithm Visualizer for sorting and searching algorithms.",
            "priority": "High Priority"
        },

        "Flask": {
            "description": "Learn Flask routing, templates, forms, request handling, sessions, APIs, and connecting Flask applications to databases.",
            "duration": "1-2 weeks",
            "project": "Build a Resume Analyzer web application using Flask.",
            "priority": "High Priority"
        },

        "Django": {
            "description": "Learn Django project structure, models, views, templates, forms, URLs, admin, authentication, and database integration.",
            "duration": "2-3 weeks",
            "project": "Build a College Management Portal using Django.",
            "priority": "High Priority"
        },

        "Power BI": {
            "description": "Learn data importing, Power Query, data cleaning, relationships, DAX basics, dashboards, and interactive reports.",
            "duration": "2-3 weeks",
            "project": "Create a Student Performance Analytics Dashboard.",
            "priority": "Medium Priority"
        },

        "Excel": {
            "description": "Learn formulas, functions, sorting, filtering, conditional formatting, pivot tables, charts, and basic data analysis.",
            "duration": "1-2 weeks",
            "project": "Create an Excel-based Student Performance Dashboard.",
            "priority": "Medium Priority"
        },

        "AWS": {
            "description": "Learn cloud fundamentals, EC2, S3, IAM, basic networking, deployment, and monitoring concepts.",
            "duration": "3-4 weeks",
            "project": "Deploy a small Flask application on AWS.",
            "priority": "High Priority"
        },

        "Cloud": {
            "description": "Learn cloud computing concepts, virtualization, storage, networking, security, scalability, and cloud deployment basics.",
            "duration": "2-3 weeks",
            "project": "Deploy a simple web application using a cloud platform.",
            "priority": "Medium Priority"
        },

        "API": {
            "description": "Learn how APIs work, HTTP methods, requests, responses, JSON, authentication, and consuming APIs from applications.",
            "duration": "1 week",
            "project": "Build an application that consumes a public REST API.",
            "priority": "High Priority"
        },

        "REST API": {
            "description": "Learn REST principles, endpoints, HTTP methods, status codes, JSON, CRUD operations, and API testing.",
            "duration": "1-2 weeks",
            "project": "Build a REST API for a Student Management System.",
            "priority": "High Priority"
        },

        "Communication": {
            "description": "Practice professional communication, technical explanations, presentation skills, active listening, and interview communication.",
            "duration": "2-3 weeks",
            "project": "Prepare and deliver a 10-minute presentation about one of your projects.",
            "priority": "Medium Priority"
        },

        "Leadership": {
            "description": "Develop task planning, delegation, decision-making, collaboration, responsibility, and team coordination skills.",
            "duration": "2-3 weeks",
            "project": "Lead a small college project and maintain a simple task plan for the team.",
            "priority": "Medium Priority"
        },

        "Problem Solving": {
            "description": "Practice breaking problems into smaller parts, identifying patterns, choosing algorithms, and testing solutions systematically.",
            "duration": "3-4 weeks",
            "project": "Solve a structured set of programming problems and document your approaches.",
            "priority": "High Priority"
        },

        "Teamwork": {
            "description": "Learn collaborative planning, Git-based workflows, communication, conflict resolution, and shared responsibility.",
            "duration": "2-3 weeks",
            "project": "Complete a small team project using GitHub issues and branches.",
            "priority": "Medium Priority"
        }
    }

    roadmap = []

    for skill in missing_skills:

        details = skill_details.get(
            skill,
            {
                "description": f"Learn the fundamentals of {skill}, practice the main concepts, and apply them in a small project.",
                "duration": "1-3 weeks",
                "project": f"Build a small practical project using {skill}.",
                "priority": "Medium Priority"
            }
        )

        roadmap.append({
            "skill": skill,
            "description": details["description"],
            "duration": details["duration"],
            "project": details["project"],
            "priority": details["priority"]
        })

    return roadmap


def generate_skill_analysis(resume_text, job_description):

    current_skills = extract_technologies(
        resume_text
    )

    required_skills = extract_job_keywords(
        job_description
    )

    missing_skills = [
        skill
        for skill in required_skills
        if skill not in current_skills
    ]

    roadmap = create_skill_roadmap(
        missing_skills
    )

    return {
        "current_skills": current_skills,
        "required_skills": required_skills,
        "missing_skills": missing_skills,
        "roadmap": roadmap
    }


# =========================================================
# SKILL ENHANCEMENT ROUTE
# =========================================================

@app.route("/skill-enhancement", methods=["POST"])
def skill_enhancement():

    resume_text = request.form.get(
        "resume_text",
        ""
    )

    job_description = request.form.get(
        "job_description",
        ""
    )


    if not resume_text.strip():

        return "Resume content was not provided."


    if not job_description.strip():

        return "Please provide a job description to generate the skill enhancement analysis."


    skill_analysis = generate_skill_analysis(
        resume_text,
        job_description
    )


    return render_template(
        "skill_enhancement.html",
        **skill_analysis
    )


# =========================================================
# HOME
# =========================================================

@app.route("/")
def home():

    return render_template(
        "index.html"
    )


# =========================================================
# ANALYZE
# =========================================================

@app.route("/analyze", methods=["POST"])
def analyze():

    file = request.files.get(
        "resume"
    )

    job_description = request.form.get(
        "job_description",
        ""
    )


    if not file:

        return "Please upload a resume PDF."


    try:

        text = extract_text_from_pdf(
            file
        )

    except Exception as e:

        return f"Error reading PDF: {e}"


    if not text.strip():

        return "Could not extract text from the uploaded PDF."


    analysis = analyze_resume(
        text,
        job_description
    )


    return render_template(
        "result.html",
        **analysis
    )


# =========================================================
# ENHANCE
# =========================================================

@app.route("/enhance", methods=["POST"])
def enhance():

    text = request.form.get(
        "text",
        ""
    )

    section_type = request.form.get(
        "section_type",
        "general"
    )


    if not text.strip():

        return "No resume content provided."


    # AUTOMATIC ENHANCEMENT

    if section_type == "automatic":

        enhanced_sections = enhance_detected_sections(
            text
        )

        return render_template(
            "enhanced_resume.html",
            original_text=text,
            enhanced_sections=enhanced_sections
        )


    # MANUAL ENHANCEMENT

    if section_type == "summary":

        enhanced = enhance_summary(text)

    elif section_type == "project":

        enhanced = enhance_project(text)

    elif section_type == "experience":

        enhanced = enhance_experience(text)

    elif section_type == "skills":

        enhanced = enhance_skills(text)

    elif section_type == "education":

        enhanced = enhance_education(text)

    elif section_type == "certifications":

        enhanced = enhance_certifications(text)

    else:

        enhanced = (
            text.strip()
            + "\n\n"
            + "Improved for professional presentation, "
              "clarity, and resume readability."
        )


    return render_template(
        "enhance.html",
        original=text,
        enhanced=enhanced
    )


# =========================================================
# DOWNLOAD TXT
# =========================================================

@app.route(
    "/download-enhanced",
    methods=["POST"]
)
def download_enhanced():

    content = request.form.get(
        "content",
        ""
    )


    output = BytesIO()

    output.write(
        content.encode("utf-8")
    )

    output.seek(0)


    return send_file(
        output,
        as_attachment=True,
        download_name="AI_Enhanced_Resume.txt",
        mimetype="text/plain"
    )


# =========================================================
# DOWNLOAD MANUAL WORD
# =========================================================

@app.route(
    "/download-word",
    methods=["POST"]
)
def download_word():

    content = request.form.get(
        "content",
        ""
    )


    document = Document()

    document.add_heading(
        "AI Enhanced Resume",
        0
    )


    paragraph = document.add_paragraph(
        content
    )


    for run in paragraph.runs:

        run.font.size = Pt(11)


    output = BytesIO()

    document.save(output)

    output.seek(0)


    return send_file(
        output,
        as_attachment=True,
        download_name="AI_Enhanced_Resume.docx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )


# =========================================================
# DOWNLOAD COMPLETE EDITED WORD RESUME
# =========================================================

@app.route(
    "/download-complete-word",
    methods=["POST"]
)
def download_complete_word():

    document = Document()

    # Professional page setup
    section = document.sections[0]
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    # Normal font
    normal_style = document.styles["Normal"]
    normal_style.font.name = "Aptos"
    normal_style.font.size = Pt(10.5)

    # Resume title
    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("AI Enhanced Resume")
    run.bold = True
    run.font.name = "Aptos Display"
    run.font.size = Pt(20)

    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subrun = subtitle.add_run("Professionally formatted with AI-powered resume enhancement")
    subrun.italic = True
    subrun.font.size = Pt(9)

    section_numbers = []

    for key in request.form:
        if key.startswith("section_"):
            try:
                number = int(key.replace("section_", ""))
                section_numbers.append(number)
            except ValueError:
                pass

    section_numbers = sorted(set(section_numbers))

    def add_heading(text):
        paragraph = document.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(9)
        paragraph.paragraph_format.space_after = Pt(4)
        paragraph.paragraph_format.keep_with_next = True

        run = paragraph.add_run(text.upper())
        run.bold = True
        run.font.name = "Aptos"
        run.font.size = Pt(11)

        # Add a bottom border to the heading paragraph.
        p = paragraph._p
        pPr = p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "666666")
        pBdr.append(bottom)
        pPr.append(pBdr)

    def add_content(content):
        for raw_line in content.splitlines():
            line = raw_line.strip()
            if not line:
                continue

            is_bullet = bool(re.match(r"^[•\-\*]\s+", line))
            cleaned = re.sub(r"^[•\-\*]\s+", "", line).strip()

            if not cleaned:
                continue

            paragraph = document.add_paragraph()
            paragraph.paragraph_format.space_after = Pt(3)
            paragraph.paragraph_format.line_spacing = 1.05

            if is_bullet:
                paragraph.style = document.styles["List Bullet"]
                paragraph.paragraph_format.left_indent = Inches(0.18)
                paragraph.paragraph_format.first_line_indent = Inches(-0.12)

            run = paragraph.add_run(cleaned)
            run.font.name = "Aptos"
            run.font.size = Pt(10.5)

    for number in section_numbers:
        title_text = request.form.get(f"title_{number}")
        content = request.form.get(f"section_{number}")

        if title_text and content:
            add_heading(title_text)
            add_content(content)

    if not section_numbers:
        document.add_paragraph("No enhanced resume sections were provided.")

    output = BytesIO()
    document.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="AI_Enhanced_Professional_Resume.docx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )


# =========================================================
# DOWNLOAD COMPLETE ENHANCED PDF RESUME
# =========================================================

@app.route(
    "/download-complete-pdf",
    methods=["POST"]
)
def download_complete_pdf():

    output = BytesIO()

    document = SimpleDocTemplate(
        output,
        pagesize=A4,
        rightMargin=17 * mm,
        leftMargin=17 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "ProfessionalResumeTitle",
        parent=styles["Title"],
        alignment=TA_CENTER,
        fontName="Helvetica-Bold",
        fontSize=19,
        leading=22,
        spaceAfter=3,
        textColor=colors.HexColor("222222")
    )

    subtitle_style = ParagraphStyle(
        "ProfessionalResumeSubtitle",
        parent=styles["Normal"],
        alignment=TA_CENTER,
        fontName="Helvetica-Oblique",
        fontSize=8.5,
        leading=11,
        spaceAfter=12,
        textColor=colors.HexColor("666666")
    )

    heading_style = ParagraphStyle(
        "ProfessionalResumeHeading",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=11.5,
        leading=14,
        spaceBefore=9,
        spaceAfter=5,
        textColor=colors.HexColor("222222"),
        borderPadding=(0, 0, 3, 0),
        borderColor=colors.HexColor("777777"),
        borderWidth=0.6,
        borderRadius=0
    )

    body_style = ParagraphStyle(
        "ProfessionalResumeBody",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=9.5,
        leading=13,
        spaceAfter=3.5,
        textColor=colors.HexColor("333333")
    )

    bullet_style = ParagraphStyle(
        "ProfessionalResumeBullet",
        parent=body_style,
        leftIndent=12,
        firstLineIndent=-7,
        bulletIndent=0,
        spaceAfter=3
    )

    story = []
    story.append(Paragraph("AI Enhanced Resume", title_style))
    story.append(Paragraph(
        "Professionally formatted with AI-powered resume enhancement",
        subtitle_style
    ))

    section_numbers = []

    for key in request.form:
        if key.startswith("section_"):
            try:
                number = int(key.replace("section_", ""))
                section_numbers.append(number)
            except ValueError:
                pass

    section_numbers = sorted(set(section_numbers))

    for number in section_numbers:
        title = request.form.get(f"title_{number}")
        content = request.form.get(f"section_{number}")

        if not title or not content:
            continue

        story.append(Paragraph(escape(title.upper()), heading_style))

        for raw_line in content.splitlines():
            line = raw_line.strip()
            if not line:
                continue

            is_bullet = bool(re.match(r"^[•\-\*]\s+", line))
            cleaned = re.sub(r"^[•\-\*]\s+", "", line).strip()

            if not cleaned:
                continue

            safe_line = escape(cleaned)

            if is_bullet:
                story.append(
                    Paragraph("• " + safe_line, bullet_style)
                )
            else:
                story.append(
                    Paragraph(safe_line, body_style)
                )

    if not section_numbers:
        story.append(
            Paragraph(
                "No enhanced resume sections were provided.",
                body_style
            )
        )

    document.build(story)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="AI_Enhanced_Professional_Resume.pdf",
        mimetype="application/pdf"
    )


# =========================================================
# RUN
# =========================================================

@app.route("/template-selection", methods=["GET", "POST"])
def template_selection():

    template_names = {
        "professional": "Professional",
        "modern": "Modern",
        "minimal": "Minimal"
    }

    enhanced_sections = {}

    # Preferred method: receive the complete sections as JSON.
    sections_json = request.form.get("sections_json", "").strip()
    if sections_json:
        try:
            import json
            decoded = json.loads(sections_json)
            if isinstance(decoded, dict):
                enhanced_sections = {
                    str(k): str(v)
                    for k, v in decoded.items()
                    if str(k).strip() and str(v).strip()
                }
        except (ValueError, TypeError):
            enhanced_sections = {}

    # Backward-compatible support for title_1/section_1 form fields.
    if not enhanced_sections:
        section_numbers = []

        for key in request.form:
            if key.startswith("section_"):
                try:
                    section_numbers.append(
                        int(key.replace("section_", ""))
                    )
                except ValueError:
                    pass

        for number in sorted(set(section_numbers)):
            title = request.form.get(
                f"title_{number}", ""
            ).strip()
            content = request.form.get(
                f"section_{number}", ""
            ).strip()

            if title and content:
                enhanced_sections[title] = content

    original_text = request.form.get(
        "original_text", ""
    )

    if request.method == "POST" and request.form.get("template"):

        selected_template = request.form.get("template")

        if selected_template not in template_names:
            selected_template = "professional"

        return render_template(
            "enhanced_resume.html",
            original_text=original_text,
            enhanced_sections=enhanced_sections,
            selected_template=selected_template,
            selected_name=template_names[selected_template]
        )

    return render_template(
        "template_selection.html",
        original_text=original_text,
        enhanced_sections=enhanced_sections,
        selected_template="",
        selected_name="",
        resume_info={}
    )


if __name__ == "__main__":

    app.run(
        debug=True
    )
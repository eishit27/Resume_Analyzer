"""
Resume & ATS Analyzer Application
Comprehensive Desktop GUI Application built with Python and Tkinter.
Extracts keywords, computes ATS match scores via Cosine Similarity,
identifies matching and missing skills, and provides structural formatting reviews.
"""

import re
import math
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

SKILL_CATEGORIES = {
    "Programming Languages": ["python", "java", "c++", "c#", "javascript", "typescript", "go", "rust", "ruby", "php", "swift", "kotlin", "sql", "html", "css", "r"],
    "Frameworks & Libraries": ["react", "angular", "vue", "django", "flask", "fastapi", "spring boot", "nodejs", "express", "pytorch", "tensorflow", "keras", "pandas", "numpy", "scikit-learn", "dotnet", "nextjs"],
    "Cloud & DevOps": ["aws", "azure", "gcp", "docker", "kubernetes", "terraform", "ci/cd", "jenkins", "git", "linux", "ansible", "nginx", "github actions"],
    "Databases & Storage": ["postgresql", "mysql", "mongodb", "redis", "elasticsearch", "cassandra", "sqlite", "oracle", "snowflake", "bigquery"],
    "Concepts & Tools": ["rest api", "graphql", "microservices", "agile", "scrum", "oop", "tdd", "unit testing", "system design", "data structures", "algorithms"]
}

SAMPLE_RESUMES = {
    "Full-Stack Python Engineer": {
        "resume": """John Doe
Software Engineer | john.doe@email.com | +1 (555) 019-2834 | github.com/johndoe

PROFESSIONAL SUMMARY
Results-driven Full-Stack Python Developer with 4+ years of experience designing scalable REST APIs, microservices, and modern web applications using Django, FastAPI, React, and PostgreSQL. Proficient in Docker, AWS, and CI/CD pipelines.

TECHNICAL SKILLS
Languages: Python, JavaScript, TypeScript, SQL, HTML5, CSS3
Frameworks & Libraries: Django, FastAPI, Flask, React, Redux, Node.js
Databases: PostgreSQL, Redis, MySQL, SQLite
Cloud & DevOps: AWS (EC2, S3, RDS), Docker, Git, GitHub Actions, Linux
Methodologies: Agile, Scrum, REST API design, Test-Driven Development (TDD), Microservices

PROFESSIONAL EXPERIENCE
Senior Software Engineer | TechCorp Inc. | 2022 – Present
• Architected high-throughput REST APIs using Python and FastAPI, reducing response latency by 35%.
• Developed responsive single-page applications (SPA) using React, TypeScript, and Redux.
• Containerized microservices using Docker and deployed on AWS ECS with automated CI/CD workflows.
• Optimized PostgreSQL query performance, reducing database query execution times by 40%.

Software Engineer | WebSolutions Ltd | 2020 – 2022
• Built full-stack web applications using Django and PostgreSQL.
• Implemented Redis caching layers, improving application throughput by 50%.
• Collaborated in an Agile Scrum team to deliver weekly production feature releases.

EDUCATION
Bachelor of Science in Computer Science | State University | GPA: 3.8/4.0
""",
        "jd": """Senior Python Full-Stack Developer
We are looking for a Senior Full-Stack Python Developer to build next-generation cloud platforms.

Responsibilities:
• Design and maintain high-performance microservices using Python (FastAPI / Django).
• Build clean, modular front-end interfaces using React, TypeScript, and modern CSS.
• Design schema and optimize queries for PostgreSQL and Redis caches.
• Manage containerized deployments using Docker and Kubernetes on AWS or GCP.
• Implement automated unit testing and CI/CD pipelines.

Required Qualifications:
• 3+ years experience with Python, FastAPI, Django, and React.
• Strong database proficiency with PostgreSQL and Redis.
• Hands-on experience with Docker, Kubernetes, AWS, and Git.
• Solid grasp of Microservices, REST APIs, System Design, and Agile methodologies.
• Experience with GraphQL or Kafka is a strong plus.
"""
    },
    "Data Scientist / ML Engineer": {
        "resume": """Jane Smith
Data Scientist & Machine Learning Engineer | jane.smith@email.com | (555) 012-3456

PROFESSIONAL SUMMARY
Machine Learning Specialist with expertise in predictive modeling, deep learning, and NLP using Python, PyTorch, TensorFlow, and Scikit-Learn. Experienced in building automated data pipelines with Pandas, NumPy, and SQL.

TECHNICAL SKILLS
Languages: Python, R, SQL, C++
Machine Learning: PyTorch, TensorFlow, Scikit-Learn, Keras, OpenCV, NLP, Computer Vision
Data Engineering: Pandas, NumPy, Spark, BigQuery, Tableau
Cloud & Tools: Docker, AWS (SageMaker), Git, Linux, Jupyter

EXPERIENCE
Lead Data Scientist | AI Innovations | 2022 – Present
• Developed deep learning computer vision models using PyTorch achieving 96% classification accuracy.
• Deployed real-time inference models using Docker and AWS SageMaker.
• Built data cleaning pipelines processing 50M+ rows with Pandas and SQL.

EDUCATION
Master of Science in Artificial Intelligence | Tech Institute
""",
        "jd": """Senior Machine Learning Engineer
Looking for an experienced Data Scientist / ML Engineer to develop predictive AI models.

Requirements:
• Strong coding skills in Python and SQL.
• Expertise in PyTorch, TensorFlow, Scikit-Learn, and Pandas.
• Experience deploying models with Docker on AWS or Azure.
• Knowledge of Deep Learning, Computer Vision, and Big Data tools like Spark.
"""
    }
}

def read_resume(filename):
    path = Path(filename)
    if path.stat().st_size > 10 * 1024 * 1024:
        raise ValueError("Please use a document smaller than 10 MB.")
    if path.suffix.lower() == ".docx":
        with zipfile.ZipFile(path) as archive:
            info = archive.getinfo("word/document.xml")
            if info.file_size > 20 * 1024 * 1024:
                raise ValueError("Document text is too large.")
            root = ET.fromstring(archive.read(info))
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        return "\n".join("".join(p.itertext()) for p in root.findall(".//w:p", ns))
    if path.suffix.lower() not in (".txt", ".md"):
        raise ValueError("Use TXT, Markdown, or DOCX; export older Word files as DOCX first.")
    return path.read_text(encoding="utf-8-sig")

class ResumeAnalyzer(ttk.Frame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.pack(fill=tk.BOTH, expand=True, padx=16, pady=16)

        self.sample_choice_var = tk.StringVar(value="Full-Stack Python Engineer")
        self._build_ui()
        self._load_sample()

    def _build_ui(self):
        # 1. Top Controls Bar
        top_bar = ttk.Frame(self)
        top_bar.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(top_bar, text="📁 Presets:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        sample_combo = ttk.Combobox(
            top_bar,
            textvariable=self.sample_choice_var,
            values=list(SAMPLE_RESUMES.keys()),
            state="readonly",
            width=28
        )
        sample_combo.pack(side=tk.LEFT, padx=(0, 8))
        sample_combo.bind("<<ComboboxSelected>>", lambda e: self._load_sample())

        ttk.Button(top_bar, text="📂 Load Resume File...", command=self._load_resume_file).pack(side=tk.LEFT, padx=4)
        ttk.Button(top_bar, text="🧹 Clear All", command=self._clear_all).pack(side=tk.LEFT, padx=4)

        self.analyze_btn = ttk.Button(top_bar, text="⚡ Run ATS Match Analysis", style='Action.TButton', command=self._run_analysis)
        self.analyze_btn.pack(side=tk.RIGHT, padx=4)

        # 2. Main Middle Section: Resume (Left) vs Job Description (Right)
        text_container = ttk.Frame(self)
        text_container.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Left Text: Resume
        left_frame = ttk.LabelFrame(text_container, text=" Candidate Resume Text ", padding=6)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))

        self.resume_text = tk.Text(left_frame, wrap=tk.WORD, font=("Segoe UI", 10), height=14)
        r_scroll = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.resume_text.yview)
        self.resume_text.configure(yscrollcommand=r_scroll.set)
        self.resume_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        r_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Right Text: Job Description
        right_frame = ttk.LabelFrame(text_container, text=" Target Job Description (JD) ", padding=6)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(6, 0))

        self.jd_text = tk.Text(right_frame, wrap=tk.WORD, font=("Segoe UI", 10), height=14)
        j_scroll = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.jd_text.yview)
        self.jd_text.configure(yscrollcommand=j_scroll.set)
        self.jd_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        j_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 3. Bottom Results & Scoring HUD
        results_frame = ttk.LabelFrame(self, text=" ATS Evaluation & Skill Gap Analysis ", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True)

        # Left Score Gauge Card
        score_card = tk.Frame(results_frame, bg="#2B6CB0", padx=16, pady=12, relief=tk.RAISED, borderwidth=1)
        score_card.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 12))

        ttk.Label(score_card, text="KEYWORD MATCH", font=("Segoe UI", 8, "bold"), background="#2B6CB0", foreground="#E2E8F0").pack()
        self.score_lbl = tk.Label(score_card, text="--%", font=("Segoe UI", 32, "bold"), bg="#2B6CB0", fg="#FFFFFF")
        self.score_lbl.pack()

        self.rating_lbl = tk.Label(score_card, text="Click 'Run Analysis'", font=("Segoe UI", 9), bg="#2B6CB0", fg="#FEFCBF")
        self.rating_lbl.pack()

        # Right Analysis Tabs
        self.tabs = ttk.Notebook(results_frame)
        self.tabs.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.tab_matched = ttk.Frame(self.tabs, padding=8)
        self.tab_missing = ttk.Frame(self.tabs, padding=8)
        self.tab_format = ttk.Frame(self.tabs, padding=8)

        self.tabs.add(self.tab_matched, text=" ✅ Matched Skills ")
        self.tabs.add(self.tab_missing, text=" ⚠️ Missing Keywords ")
        self.tabs.add(self.tab_format, text=" 📋 Structure & Recommendations ")

        # Matched box
        self.matched_box = tk.Text(self.tab_matched, wrap=tk.WORD, font=("Segoe UI", 9), height=5, relief=tk.FLAT, bg="#F7FAFC")
        self.matched_box.pack(fill=tk.BOTH, expand=True)
        self.matched_box.config(state=tk.DISABLED)

        # Missing box
        self.missing_box = tk.Text(self.tab_missing, wrap=tk.WORD, font=("Segoe UI", 9), height=5, relief=tk.FLAT, bg="#F7FAFC")
        self.missing_box.pack(fill=tk.BOTH, expand=True)
        self.missing_box.config(state=tk.DISABLED)

        # Format box
        self.format_box = tk.Text(self.tab_format, wrap=tk.WORD, font=("Segoe UI", 9), height=5, relief=tk.FLAT, bg="#F7FAFC")
        self.format_box.pack(fill=tk.BOTH, expand=True)
        self.format_box.config(state=tk.DISABLED)

    def _load_sample(self):
        choice = self.sample_choice_var.get()
        if choice in SAMPLE_RESUMES:
            self.resume_text.delete("1.0", tk.END)
            self.resume_text.insert("1.0", SAMPLE_RESUMES[choice]["resume"])
            self.jd_text.delete("1.0", tk.END)
            self.jd_text.insert("1.0", SAMPLE_RESUMES[choice]["jd"])
            self._run_analysis()

    def _load_resume_file(self):
        fp = filedialog.askopenfilename(filetypes=[("Supported documents", "*.txt *.md *.docx")])
        if not fp:
            return
        try:
            content = read_resume(fp)
            if not content.strip():
                raise ValueError("The document contains no readable text.")
            self._clear_all_results()
            self.resume_text.delete("1.0", tk.END)
            self.resume_text.insert("1.0", content)
        except Exception as exc:
            messagebox.showerror("File Error", str(exc))

    def _clear_all_results(self):
        self.score_lbl.config(text="--%")
        self.rating_lbl.config(text="Awaiting analysis")
        for box in (self.matched_box, self.missing_box, self.format_box):
            box.config(state=tk.NORMAL)
            box.delete("1.0", tk.END)
            box.config(state=tk.DISABLED)

    def _clear_all(self):
        self._clear_all_results()
        self.resume_text.delete("1.0", tk.END)
        self.jd_text.delete("1.0", tk.END)
        self.score_lbl.config(text="--%")
        self.rating_lbl.config(text="Awaiting inputs")

    def _extract_tokens(self, text):
        clean = re.sub(r'[^a-zA-Z0-9\+#\./]', ' ', text.lower())
        tokens = [t.strip('.') for t in clean.split() if len(t.strip('.')) > 1]
        return tokens

    def _find_known_skills(self, text):
        lower = text.lower()
        for alias, canonical in {"node.js": "nodejs", "next.js": "nextjs", ".net": "dotnet"}.items():
            lower = lower.replace(alias, canonical)
        found = set()
        for cat, skills in SKILL_CATEGORIES.items():
            for s in skills:
                # Word boundary check
                pattern = r'(?<![a-zA-Z0-9])' + re.escape(s) + r'(?![a-zA-Z0-9])'
                if re.search(pattern, lower):
                    found.add(s)
        return found

    def _run_analysis(self):
        res_str = self.resume_text.get("1.0", tk.END).strip()
        jd_str = self.jd_text.get("1.0", tk.END).strip()

        if not res_str or not jd_str:
            messagebox.showwarning("Incomplete Data", "Please supply both Resume and Job Description texts.")
            return

        res_tokens = self._extract_tokens(res_str)
        jd_tokens = self._extract_tokens(jd_str)

        # 1. Cosine Similarity on Token Frequencies
        res_counts = Counter(res_tokens)
        jd_counts = Counter(jd_tokens)
        all_vocab = set(res_counts.keys()).union(set(jd_counts.keys()))

        dot_prod = sum(res_counts[w] * jd_counts[w] for w in all_vocab)
        mag_res = math.sqrt(sum(v**2 for v in res_counts.values()))
        mag_jd = math.sqrt(sum(v**2 for v in jd_counts.values()))

        cosine_sim = dot_prod / (mag_res * mag_jd) if (mag_res * mag_jd) > 0 else 0

        # 2. Known Technical Skill Match
        res_skills = self._find_known_skills(res_str)
        jd_skills = self._find_known_skills(jd_str)

        matched_skills = res_skills.intersection(jd_skills)
        missing_skills = jd_skills - res_skills

        skill_match_ratio = len(matched_skills) / max(1, len(jd_skills))

        # Composite ATS Score (60% Skill Match + 40% Text Overlap)
        final_score = int(round((0.60 * skill_match_ratio + 0.40 * cosine_sim) * 100.0))
        final_score = min(98, max(12, final_score))

        self.score_lbl.config(text=f"{final_score}%")
        if final_score >= 80:
            rating = "🌟 Excellent Match"
        elif final_score >= 65:
            rating = "👍 Good Candidate Match"
        elif final_score >= 50:
            rating = "⚠️ Moderate Match"
        else:
            rating = "❌ Low Keyword Match"
        self.rating_lbl.config(text=rating)

        # 3. Populate Matched Skills Tab
        self.matched_box.config(state=tk.NORMAL)
        self.matched_box.delete("1.0", tk.END)
        if matched_skills:
            m_text = f"Found {len(matched_skills)} overlapping technical competencies:\n\n"
            m_text += " • " + "\n • ".join(sorted(s.title() for s in matched_skills))
        else:
            m_text = "No direct key technical skill matches identified."
        self.matched_box.insert("1.0", m_text)
        self.matched_box.config(state=tk.DISABLED)

        # 4. Populate Missing Skills Tab
        self.missing_box.config(state=tk.NORMAL)
        self.missing_box.delete("1.0", tk.END)
        if missing_skills:
            miss_text = f"Identified {len(missing_skills)} skills required by the JD but absent in your resume:\n\n"
            miss_text += " ⚠️ " + "\n ⚠️ ".join(sorted(s.title() for s in missing_skills))
            miss_text += "\n\nTip: Include these skills only when you genuinely have relevant experience. This heuristic score does not predict hiring outcomes."
        else:
            miss_text = "🎉 Awesome! All major technical keywords found in the Job Description are present in your resume!"
        self.missing_box.insert("1.0", miss_text)
        self.missing_box.config(state=tk.DISABLED)

        # 5. Populate Structural & ATS Formatting Checklist
        self.format_box.config(state=tk.NORMAL)
        self.format_box.delete("1.0", tk.END)

        lower_res = res_str.lower()
        sections = {
            "Contact Info (Email/Phone)": bool(re.search(r'[@\d-]{5,}', res_str)),
            "Summary / Objective": "summary" in lower_res or "profile" in lower_res or "objective" in lower_res,
            "Experience / Work History": "experience" in lower_res or "employment" in lower_res or "work" in lower_res,
            "Education": "education" in lower_res or "degree" in lower_res or "university" in lower_res,
            "Skills Section": "skills" in lower_res or "technical" in lower_res,
            "Projects": "project" in lower_res or "portfolio" in lower_res
        }

        f_text = "ATS Formatting & Structural Checklist:\n"
        for sec, present in sections.items():
            status = "✅ Present" if present else "❌ Missing / Unclear"
            f_text += f"• {sec}: {status}\n"

        words = len(res_tokens)
        f_text += f"\nWord Count: {words} words ("
        if 350 <= words <= 850:
            f_text += "Ideal 1-2 page density ✅)\n"
        elif words < 350:
            f_text += "Too brief; consider adding project details ⚠️)\n"
        else:
            f_text += "Lengthy; consider streamlining for ATS brevity ⚠️)\n"

        self.format_box.insert("1.0", f_text)
        self.format_box.config(state=tk.DISABLED)

class StandaloneResume(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("📄 Resume & ATS Keyword Match Analyzer - Python Tkinter")
        self.geometry("960x680")
        self.minsize(840, 580)
        try:
            ttk.Style(self).theme_use('vista')
        except Exception:
            pass
        ResumeAnalyzer(self)

if __name__ == '__main__':
    app = StandaloneResume()
    app.mainloop()

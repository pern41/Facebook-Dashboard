import csv
import html
import re
import zipfile
from collections import Counter
from functools import lru_cache
from pathlib import Path
from xml.sax.saxutils import escape

RAW_PATH = Path(r"D:\AI\comment\input.txt")
OUT_DIR = Path(r"C:D:\AI\comment\test")
CSV_PATH = OUT_DIR / "testcleaned_claude_comments_structured.csv"
XLSX_PATH = OUT_DIR / "testcleaned_claude_comments_structured.xlsx"
SUMMARY_PATH = OUT_DIR / "testcleaned_claude_comments_structured_summary.txt"


COLUMNS = [
    "Main Category",
    "Sub Category",
    "Intent",
    "Sentiment",
    "User Persona",
    "Adoption Stage",
    "Pain Point",
    "Feature Request",
    "Use Case Domain",
    "Keywords",
    "Confidence",
    "Summary",
]


BANNED_MAIN_CATEGORIES = {
    "Frustrations",
    "Feature Requests",
    "Adoption Stage",
    "Content Creator Segment",
    "Job / Career",
}


COLUMN_WIDTHS = {
    "Main Category": 28,
    "Sub Category": 24,
    "Intent": 24,
    "Sentiment": 14,
    "User Persona": 20,
    "Adoption Stage": 18,
    "Pain Point": 18,
    "Feature Request": 20,
    "Use Case Domain": 24,
    "Keywords": 42,
    "Confidence": 12,
    "Summary": 90,
}


TEXT_NORMALIZATIONS = (
    (r"\bfreelanceing\b", "freelancing"),
    (r"\bsubcription\b", "subscription"),
    (r"\brepalce\b", "replace"),
    (r"\bprompt engineer\b", "prompt engineering"),
    (r"\bhour saved\b", "hours saved"),
    (r"\boperation\b", "operations"),
    (r"แดชบอด", "แดชบอร์ด"),
    (r"ข้อมูบ", "ข้อมูล"),
    (r"เกี่บ", "เกี่ยว"),
    (r"วิเคราะห(?!์)", "วิเคราะห์"),
)

#Content Creation, Coding&Dev, Data, Research, Education, Operation, Automatation, Design, Entre&Strategy, Productivity 
CATEGORY_KEYWORDS = {
    "Content Creation": {
        "SEO": [
            "seo",
            "search engine optimization",
            "keyword research",
            "topic cluster",
            "backlink",
            "ranking",
            "google ranking",
            "geo",
            "generative engine optimization",
            "on page seo",
            "technical seo",
            "meta title",
            "meta description",
            "schema",
            "semantic seo",
            "content optimization",
            #TH
            "ทำ seo",
            "ปรับ seo",
            "ติดอันดับ google",
            "อันดับ google",
            "ค้นหาคีย์เวิร์ด",
            "keyword",
            "คีย์เวิร์ด",
            "วิเคราะห์คีย์เวิร์ด",
            "เขียนบทความ seo",
            "เพิ่ม traffic",
            "เพิ่มคนเข้าเว็บ",
            "ทำอันดับเว็บไซต์",
            "ปรับเว็บไซต์",
        ],
        "Article Writing": [
            "article",
            "blog",
            "blog post",
            "long form content",
            "content writing",
            "copywriting",
            "editorial",
            "newsletter",
            "ghostwriting",
            "web content",
            #TH
            "เขียนบทความ",
            "เขียนคอนเทนต์",
            "ทำบทความ",
            "สร้างบทความ",
            "เขียนข่าว",
            "เขียนรีวิว",
            "เขียนเนื้อหา",
            "เขียนเว็บ",
            "เขียนข้อความ",
            "เขียนสื่อการสอน",
            "สรุปบทความ",
            "เรียบเรียงเนื้อหา",
            "เขียนหนังสือ",
            "ช่วยเขียน",
        ],
        "Social Media": [
            "facebook post",
            "instagram caption",
            "linkedin post",
            "x post",
            "twitter thread",
            "social content",
            "content calendar",
            "content pillar",
            "engagement post",
            "viral content",
            "hook",
            "cta",
            "social app",
            "caption",
            #TH
            "โพสต์เฟซบุ๊ก",
            "โพสต์ facebook",
            "โพสต์เพจ",
            "โพสต์ไอจี",
            "แคปชั่น",
            "คิดคอนเทนต์",
            "คอนเทนต์โซเชียล",
            "วางแผนคอนเทนต์",
            "ปฏิทินคอนเทนต์",
            "content calendar",
            "โพสต์ขายของ",
            "เพิ่ม engagement",
            "ไวรัล",
            "ทำเพจ",
            "ดูแลเพจ",
            "social media",
        ],
        "Video Content": [
            "youtube script",
            "tiktok script",
            "reels script",
            "shorts script",
            "storyboard",
            "video outline",
            "narration",
            "voiceover",
            "video idea",
            "video content",
            #TH
            "สคริปต์วิดีโอ",
            "สคริปต์ youtube",
            "สคริปต์ tiktok",
            "สคริปต์ reels",
            "สคริปต์ shorts",
            "เขียนสคริปต์",
            "คิดคลิป",
            "คิดคอนเทนต์วิดีโอ",
            "วาง storyboard",
            "พากย์เสียง",
            "voice over",
            "เล่าเรื่อง",
            "ทำคลิป",
            "ยูทูบ",
            "ติ๊กต็อก",
            "คลิปสั้น",
        ],
        "Marketing Content": [
            "ad copy",
            "sales page",
            "landing page",
            "marketing campaign",
            "email marketing",
            "funnel",
            "lead magnet",
            "promotion",
            "product description",
            "sales copy",
            #TH
            "การตลาด",
            "แคมเปญการตลาด",
            "คิดแคมเปญ",
            "ยิงแอด",
            "โฆษณา",
            "เขียนโฆษณา",
            "ข้อความโฆษณา",
            "landing page",
            "หน้าขาย",
            "หน้าเซลล์",
            "โปรโมทสินค้า",
            "โปรโมชัน",
            "promotion",
            "เพิ่มยอดขาย",
            "ขายสินค้า",
            "คำอธิบายสินค้า",
            "sales page",
            "funnel",
        ],
        "Email Writing": [
            "email", 
            "gmail", 
            "inbox", 
            "cold email",
            #TH
            "อีเมล",
            "เขียนอีเมล",
            "ร่างอีเมล",
            "ตอบอีเมล",
            "ส่งอีเมล",
            "gmail",
            "ติดต่อทางอีเมล",
            "อีเมลธุรกิจ",
            "cold email",
            "อีเมลขาย",
            "อีเมลหาลูกค้า",
        ],
    },
    "Coding & Development": {
        "Frontend": [
            "html",
            "css",
            "javascript",
            "typescript",
            "react",
            "vue",
            "angular",
            "tailwind",
            "bootstrap",
            "nextjs",
            "next.js",
            "ui",
            "frontend",
            "swiftui",
             #TH
            "หน้าเว็บ",
            "เว็บ",
            "เว็บไซต์",
            "เว็บแอป",
            "เว็บแอพ",
            "ออกแบบหน้าเว็บ",
            "ออกแบบ ui",
            "ออกแบบ ux",
            "พัฒนาเว็บไซต์",
            "สร้างเว็บไซต์",
            "ทำเว็บ",
            "สร้างหน้าเว็บ",
        ],
        "Backend": [
            "nodejs",
            "node.js",
            "express",
            "python",
            "django",
            "flask",
            "php",
            "laravel",
            "java",
            "spring",
            "api",
            "backend",
            "swift",
            "cli",
            "shell",
            "bash",
            "zsh",
            "script",
            "containerized",
            "raspberry pi",
            "fax server",
            "google colab",
            "notebook",
            "backend",
            "api",
            "database",
            #TH
            "เขียนโปรแกรม",
            "เขียนโค้ด",
            "พัฒนาระบบ",
            "สร้างระบบ",
            "ระบบหลังบ้าน",
            "เชื่อม api",
            "เชื่อมต่อระบบ",
            "ระบบจัดการ",
            "โปรแกรม",
            "ซอฟต์แวร์",
            "เขียน script",
            "ทำระบบ",
        ],
        "Full Stack": [
            "full stack",
            "web app",
            "website",
            "saas",
            "application",
            "app",
            "deployment",
            "github pages",
            "vercel",
            "netlify",
            "ios app",
            "android app",
            "mobile app",
            "chrome extension",
            "extension",
            "built",
            "vibecoded",
            "vibe coded",
            "vibe-coding",
            "vibecoding",
            "pwa",
            "claude code",
            "github.com",
            "code is written",
            #TH
            "สร้างแอป",
            "ทำแอป",
            "แอปพลิเคชัน",
            "สร้างเว็บ",
            "สร้างเว็บแอป",
            "สร้างโปรแกรม",
            "พัฒนาซอฟต์แวร์",
            "สร้างระบบงาน",
            "ระบบออนไลน์",
            "แอปมือถือ",
            "ระบบสมาชิก",
        ],
        "Debugging": [
            "bug",
            "fix",
            "error",
            "debug",
            "debugging",
            "troubleshoot",
            "issue",
            "exception",
            "stack trace",
            "stderr",
            "failed command",
            "wtf command",
            "thefuck",
            #TH
            "แก้บั๊ก",
            "แก้ error",
            "แก้ปัญหา",
            "แก้โค้ด",
            "ตรวจสอบระบบ",
            "หาสาเหตุ",
            "ระบบมีปัญหา",
            "โค้ดพัง",
            "แก้ไขโปรแกรม",
        ],
        "Database": [
            "mysql",
            "postgresql",
            "postgres",
            "mongodb",
            "sqlite",
            "database",
            "query",
            "sql",
            "schema",
            "supabase",
            "firebase",
            "firestore",
            "db",
            #TH
            "ฐานข้อมูล",
            "จัดเก็บข้อมูล",
            "ออกแบบฐานข้อมูล",
            "ดึงข้อมูล",
            "ค้นหาข้อมูล",
            "เชื่อมฐานข้อมูล",
        ],
    },
    "Data Analysis": {
        "Excel": [
            "excel",
            "spreadsheet",
            "xlsx",
            "csv",
            "formula",
            "pivot table",
            "vlookup",
            "xlookup",
            "power query",
            "power pivot",
            "google sheet",
            "sheet",
            # TH
            "เอ็กเซล",
            "ไฟล์ excel",
            "สูตร excel",
            "จัดการข้อมูล",
            "สเปรดชีต",
            "google sheet",
            "ชีต",
            "pivot",
            "vlookup",
        ],
        "Business Intelligence": [
            "dashboard",
            "report",
            "analytics",
            "visualization",
            "kpi",
            "metrics",
            "business intelligence",
            "bi",
            "chart",
            "แดชบอร์ด",
            "รายงาน",
            "visualization",
            #TH
            "สรุปรายงาน",
            "ตัวชี้วัด",
            "วิเคราะห์ธุรกิจ",
            "กราฟ",
            "แผนภูมิ",
        ],
        "Data Cleaning": [
            "clean data",
            "preprocessing",
            "normalization",
            "deduplication",
            "data quality",
            "sorting data",
            "classification",
            "sort",
            "sorted",
            "categorize",
            "classify",
            "cleaning",
            "cleansing",
            "ทำความสะอาดข้อมูล",
            "จัดข้อมูล",
            "จัดหมวดหมู่",
            # TH
            "ล้างข้อมูล",
            "ทำความสะอาดข้อมูล",
            "จัดข้อมูล",
            "จัดหมวดหมู่",
            "รวมข้อมูล",
            "แยกข้อมูล",
            "คัดกรองข้อมูล",
            "ข้อมูลซ้ำ",
        ],
        "Statistical Analysis": [
            "regression",
            "correlation",
            "trend analysis",
            "forecasting",
            "prediction",
            "statistics",
            "tracker",
            "tracking",
            "monitor",
            "logging",
            "log",
            "visualizer",
            "retirement",
            "finance",
            "financial",
            "spx",
            "dealer positioning",
            "s curve",
            "s-curve",
        ],
        "Data Analysis": [
            "วิเคราะห์ข้อมูล",
            "data analysis",
            "จัดกลุ่มข้อมูล",
            "สรุปข้อมูล",
            "บทสรุป",
            "insight",
            # TH
            "หาข้อมูลเชิงลึก",
            "วิเคราะห์แนวโน้ม",
            "วิเคราะห์ผล",
            "วิเคราะห์สถิติ",
        ],
        "Data Consolidation": [
            "รวมข้อมูลหลายไฟล์",
            "merge files",
            "consolidate",
            "รวมไฟล์",
        ]  
    },
    "Research": {
        "Market Research": [
            "market research",
            "competitor analysis",
            "industry analysis",
            "market size",
            "customer research",
            "target audience",
            # TH
            "วิจัยตลาด",
            "ศึกษาตลาด",
            "คู่แข่ง",
            "วิเคราะห์คู่แข่ง",
            "หากลุ่มลูกค้า",
            "ลูกค้าเป้าหมาย",
            "ศึกษาธุรกิจ",
        ],
        "Academic Research": [
            "literature review",
            "citation",
            "journal",
            "research paper",
            "thesis",
            "dissertation",
            "publication",
            "bioinformatics",
            "gene",
            "scientific",
        ],
        "Information Gathering": [
            "summarize",
            "summary",
            "compare",
            "benchmark",
            "collect information",
            "insight",
            "research",
            "find",
            "search",
            "lookup",
            # TH
            "หาข้อมูล",
            "ค้นหาข้อมูล",
            "รวบรวมข้อมูล",
            "สรุปข้อมูล",
            "รีวิวข้อมูล",
            "เปรียบเทียบข้อมูล",
            "ศึกษาข้อมูล",
        ],
    },
    "Education & Learning": {
        "Programming Learning": [
            "learn coding",
            "coding tutorial",
            "programming",
            "python tutorial",
            "javascript tutorial",
            "zero swift experience",
            "learn",
        ],
        "Language Learning": [
            "english",
            "grammar",
            "vocabulary",
            "translation",
            "language learning",
            "translate",
        ],
        "Exam Preparation": [
            "exam",
            "test",
            "quiz",
            "certification",
            "study guide",
            "flashcard",
        ],
        "Knowledge Explanation": [
            "explain",
            "eli5",
            "tutorial",
            "walkthrough",
            "guide",
            "step by step",
            "how does",
            "how do",
            "why",
        ],
    },
    "Business Operations": {
        "SOP & Documentation": [
            "sop",
            "documentation",
            "standard operating procedure",
            "knowledge base",
            "wiki",
            "internal process",
            "docs",
            "paper based",
            # TH
            "คู่มือ",
            "เอกสาร",
            "ขั้นตอนการทำงาน",
            "มาตรฐานการทำงาน",
            "คลังความรู้",
            "ฐานความรู้",
            "จัดการความรู้",
        ],
        "Project Management": [
            "project management",
            "roadmap",
            "gantt chart",
            "sprint",
            "backlog",
            "task management",
            "trello",
            "notion",
            "jira",
            "todo",
            # TH
            "บริหารโครงการ",
            "จัดการโครงการ",
            "วางแผนงาน",
            "ติดตามงาน",
            "จัดการงาน",
            "มอบหมายงาน",
            "roadmap",
        ],
        "HR & Recruitment": [
            "recruitment",
            "hiring",
            "cv screening",
            "interview",
            "onboarding",
            "employee",
            "hr",
            "legal",
            "wellbeing",
        ],
        "Customer Service": [
            "support",
            "customer service",
            "ticket",
            "help desk",
            "faq",
        ],
        "CRM": [
            "crm", 
            "client", 
            "customer", 
            "sales", 
            "cold call", 
            "construction company", 
            "salvage yard"
            # TH
            "ลูกค้า",
            "ฐานลูกค้า",
            "จัดการลูกค้า",
            "ติดตามลูกค้า",
            "ฝ่ายขาย",
            "บริหารลูกค้า",
            "งานขาย",
        ],
        "AI Business Use Cases": [
            "marketing",
            "sales",
            "lead generation",
            "prospecting",
            "cold outreach",
            "crm",
            "customer support",
            "operations",
            "consulting",
            "agency",
            "freelance",
        ],
        "Sales Analysis": [
            "ยอดขาย",
            "รายงานยอดขาย",
            "วิเคราะห์ยอดขาย",
            "forecast ยอดขาย",
            "เช็ครายการขาย",
            "trend ยอดขาย",
            "sales report",
        ],
        "Financial Analysis": [
            "งบการเงิน",
            "วิเคราะห์การเงิน",
            "ความคุ้มค่า",
            "roi",
            "ภาษี",
            "ค่าธรรมเนียม",
            "ค่าใช้จ่าย",
        ],
        "Data Reconciliation": [
            "กระทบยอด",
            "หาผลต่าง",
            "เปรียบเทียบข้อมูล",
            "reconcile",
            "reconciliation",
        ],
        "Excel Audit": [
            "ตรวจสูตร",
            "ตรวจความผิดพลาด",
            "audit",
            "finding",
            "ตรวจสอบข้อมูล",
        ],
        "Forecasting": [
            "forecast",
            "พยากรณ์",
            "แนวโน้ม",
            "trend",
            "prediction",
        ],
        "Production Planning": [
            "แผนผลิต",
            "วางแผนผลิต",
            "production planning",
        ],
        "Stock Analysis": [
            "หุ้น",
            "stock",
            "trading",
            "portfolio",
        ], 
        "Payroll": [
            "เงินเดือน",
            "payroll",
            "พนักงาน",
        ],


    },
    "AI Automation": {
        "Agent Development": [
            "ai agent",
            "autonomous agent",
            "multi agent",
            "agent workflow",
            "agent",
            "agents",
            # TH
            "เอเจนต์",
            "ai agent",
            "ผู้ช่วย ai",
            "ตัวช่วยอัตโนมัติ",
            "ผู้ช่วยอัจฉริยะ",
        ],
        "MCP": [
            "mcp", 
            "model context protocol", 
            "claude mcp", 
            "mcp server", 
            "mcp tool"
        ],
        "API Integration": [
            "api",
            "webhook",
            "integration",
            "rest api",
            "graphql",
            "sdk",
            "sync",
            "polls",
            "polling",
        ],
        "Workflow Automation": [
            "automation",
            "automate",
            "workflow",
            "zapier",
            "make",
            "n8n",
            "trigger",
            "pipeline",
            "alert",
            "alerts",
            "notification",
            "notifications",
            "notifier",
            "ntfy",
            "xmpp",
            "tts",
            "transcriber",
            "voice transcriber",
            # TH
            "ระบบอัตโนมัติ",
            "งานอัตโนมัติ",
            "ลดงานซ้ำ",
            "ลดงาน",
            "ช่วยทำงาน",
            "ทำงานแทน",
            "งานซ้ำ",
            "เชื่อม workflow",
            "อัตโนมัติ",
            "บอท",
        ],
        "AI Models":[
            "claude",
            "chatgpt",
            "gpt-4",
            "gpt-5",
            "gemini",
            "deepseek",
            "grok",
            "qwen",
            "mistral",
            "llama",
        ],
        "AI Coding Tools":[
            "cursor",
            "windsurf",
            "cline",
            "roo code",
            "aider",
            "bolt",
            "lovable",
            "replit",
            "codex",
            "claude code",
            "v0",
        ],
        "Vibe Coding":[
            "vibe coding",
            "vibecoding",
            "vibe coder",
            "ship fast",
            "one prompt",
            "no code",
            "low code",
            "build in a weekend",
            "prototype",
            "mvp",
            "side project",
        ],
        "Prompt Engineering":[
            "prompt",
            "prompting",
            "system prompt",
            "instruction",
            "chain of thought",
            "few shot",
            "context engineering",
            "prompt library",
            "prompt template",
            # TH
            "พรอมต์",
            "เขียน prompt",
            "ออกแบบ prompt",
            "คำสั่ง ai",
            "system prompt",
            "prompt template",
        ],
    },
    "Design & Creative": {
        "UI/UX": [
            "ui",
            "ux",
            "wireframe",
            "user flow",
            "prototype",
            "design system",
            "beautiful",
            "screenshot",
            # TH
            "ออกแบบหน้าจอ",
            "ออกแบบแอป",
            "ออกแบบเว็บไซต์",
            "wireframe",
            "ต้นแบบ",
        ],
        "Branding": ["branding", "brand strategy", "brand identity", "logo", "visual identity"],
        "Creative Ideation": [
            "brainstorm",
            "idea generation",
            "concept",
            "campaign idea",
            "creative brief",
            "idea",
             # TH
            "ระดมความคิด",
            "คิดไอเดีย",
            "หาไอเดีย",
            "สร้างไอเดีย",
            "คิดแคมเปญ",
            "คิดโปรเจกต์",
        ],
    },
    "Entrepreneurship & Business Strategy": {
        "Startup": [
            "startup", 
            "founder", 
            "solo builder", 
            "indie hacker",
        ],
        "Monetization": [
            "monetization", 
            "pricing", 
            "revenue", 
            "subscription", 
            "$", 
            "paid", 
            "pro",
            "make money",
            "income",
            "side hustle",
            "freelancing",
            "client work",
            "agency",
            "subscription",
            "profit",
            "business",
            "start up",
            "saas",
            "micro saas",
            # TH
            "สร้างรายได้",
            "เพิ่มรายได้",
            "หารายได้",
            "กำไร",
            "ธุรกิจ",
            "หารายได้เสริม",
            "สร้างเงิน",
        ],
        "Business Model": [
            "business plan", 
            "business model", 
            "product market fit", 
            "go to market",
            # TH
            "โมเดลธุรกิจ",
            "แผนธุรกิจ",
            "วางแผนธุรกิจ",
            "ออกแบบธุรกิจ",
            "กลยุทธ์ธุรกิจ",
        ],
        "Product Strategy": [
            "product strategy", 
            "product", 
            "launch", 
            "app store", 
            "users",
        ],
        "Growth Strategy": [
            "growth", 
            "strategy",
        ],
        "Success Stories": [
            "success",
            "achievement",
            "win",
            "milestone",
            "revenue",
            "growth",
            "promotion",
            "job offer",
        ],
    },
    "Productivity": {
        "Personal Productivity": [
            "productivity",
            "task planning",
            "time management",
            "personal assistant",
            "note taking",
            "organization",
            "workflow improvement",
            "calendar",
            "schedule",
            "kitchen kiosk",
            "kiosk",
            "shopping list",
            "recipe",
            "planner",
            "save links",
            "bookmark",
            "notes",
            "grocery list",
            "time zone",
            "private cloud",
            "cloud storage",
            "jarvis",
            "motivation quotes",
            # TH
            "เพิ่มประสิทธิภาพ",
            "ช่วยงาน",
            "จัดการเวลา",
            "วางแผนงาน",
            "ผู้ช่วยส่วนตัว",
            "งานประจำ",
            "งานน่าเบื่อ",
        ],
        "Life Admin": [
            "family calendar",
            "meal planner",
            "fitness",
            "garden",
            "migraine",
            "doctor",
            "med",
            "health",
            "oura",
            "family",
        ],
        "Productivity Gains": [
            "save time",
            "time saved",
            "hours saved",
            "efficiency",
            "productivity boost",
            "workflow",
            "faster",
            "speed up",
            "replace",
            "automated",
            # TH
            "ประหยัดเวลา",
            "ลดเวลา",
            "ทำงานเร็วขึ้น",
            "เพิ่มความเร็ว",
            "ลดขั้นตอน",
            "ลดภาระงาน",
        ],
    },
}

COMMON_THAI_KEYWORDS = [
    "ช่วยงาน",
    "ช่วยคิด",
    "ช่วยวิเคราะห์",
    "ช่วยเขียน",
    "ช่วยสรุป",
    "หาข้อมูล",
    "ค้นหาข้อมูล",
    "จัดการข้อมูล",
    "ทำงานแทน",
    "ลดงาน",
    "ลดงานซ้ำ",
    "อัตโนมัติ",
    "สร้างระบบ",
    "ทำระบบ",
    "พัฒนาระบบ",
    "วางแผน",
    "บริหาร",
    "จัดการ",
    "ออกแบบ",
    "สร้าง",
    "ทำ",
]

POSITIVE_TERMS = [
    "good",
    "great",
    "love",
    "loved",
    "amazing",
    "awesome",
    "useful",
    "impressive",
    "huge",
    "best",
    "win",
    "beautiful",
    "insane",
    "worth",
    "cool",
    "nice",
    "genius",
    "helpful",
    "saved",
    "saves",
    "real product",
    "every day",
    "can't go back",
    # TH
    "ดี",
    "ดีมาก",
    "ยอดเยี่ยม",
    "เยี่ยม",
    "สุดยอด",
    "ชอบ",
    "รักเลย",
    "มีประโยชน์",
    "คุ้ม",
    "คุ้มมาก",
    "ช่วยได้มาก",
    "ช่วยงาน",
    "ประหยัดเวลา",
    "ใช้งานทุกวัน",
    "ใช้ทุกวัน",
    "เร็วขึ้น",
    "แม่นยำ",
    "เวิร์ค",
    "เจ๋ง",
    "โคตรดี",
    "ปัง",
    "เทพ",
    "โหด",
    "แรงมาก",
    "สุดจัด",
    "โคตรเจ๋ง",
]
NEGATIVE_TERMS = [
    "slow",
    "expensive",
    "wrong",
    "hallucination",
    "hallucinate",
    "inaccurate",
    "not working",
    "failed",
    "issue",
    "problem",
    "bug",
    "crash",
    "can't",
    "cannot",
    "wtf",
    "pain",
    "cumbersome",
    "regret",
    "slop",
    "no links",
    # TH
    "ช้า",
    "แพง",
    "แพงเกินไป",
    "ผิด",
    "ไม่ถูกต้อง",
    "ไม่แม่น",
    "มั่ว",
    "หลอน",
    "มีปัญหา",
    "ปัญหา",
    "บั๊ก",
    "เออเรอร์",
    "ค้าง",
    "ล่ม",
    "ใช้งานไม่ได้",
    "พัง",
    "เสียเวลา",
    "ไม่คุ้ม",
    "แย่",
    "ห่วย",
    "น่าผิดหวัง",
    "กาก",
    "พังมาก",
    "มั่วมาก",
    "โคตรช้า",
    "ไม่โอเค",
    "ไม่ไหว",
]
PROBLEM_TERMS = [
    "problem",
    "issue",
    "bug",
    "error",
    "wrong",
    "hallucination",
    "hallucinate",
    "slow",
    "expensive",
    "not working",
    "failed",
    "crash",
    "can't",
    "cannot",
    "limit reached",
    "context window",
    # TH
    "ปัญหา",
    "มีปัญหา",
    "บั๊ก",
    "เออเรอร์",
    "ผิดพลาด",
    "ช้า",
    "แพง",
    "ใช้งานไม่ได้",
    "ค้าง",
    "ล่ม",
    "พัง",
    "ไม่ทำงาน",
    "โควต้าเต็ม",
    "ลิมิต",
    "ข้อความหมด",
    "context",
]
COMPARISON_TERMS = [
    "mistral",
    "claude vs",
    "better than",
    "compared",
    "comparison",
    "same as",
    "equivalent",
    r"\bchatgpt\b.*\bclaude\b",
    r"\bclaude\b.*\bchatgpt\b",
    r"\b(vs|versus)\b",
    # TH
    "เทียบ",
    "เปรียบเทียบ",
    "ดีกว่า",
    "เหมือนกับ",
    "ต่างจาก",
    "คล้ายกับ",
    "chatgpt ดีกว่าไหม",
    "gemini ดีกว่าไหม",
    "ดีกว่า chatgpt",
    "ดีกว่า gemini",
    "cursor",
    r"แย่กว่า",
]
DESIRED_PATTERNS = [
    r"\bi want\b",
    r"\bi'd like\b",
    r"\bi would like\b",
    r"\bi need\b",
    r"\bwant to\b",
    r"\bwould be useful\b",
    r"\bshould\b",
    r"\bcould\b",
    r"\bcan you\b",
    r"\bcan claude\b",
    r"\bhow (do|are|can)\b",
    r"\blooking for\b",
    r"\bwish\b",
    r"\bneed to\b",
    r"\bplan to\b",
    r"อยาก",
    r"ต้องการ",
    r"กำลังหา",
    r"กำลังมองหา",
    r"ช่วย.*ได้ไหม",
    r"ทำ.*ได้ไหม",
    r"สร้าง.*ได้ไหม",
    r"มีวิธีไหม",
    r"วางแผนจะ",
    "want",
    "need",
    "wish",
    "looking for",
    "i want",
    "i need",
    "อยาก",
    "ต้องการ",
    "กำลังหา",
    "กำลังมองหา",
    "วางแผนจะ",
]
RECOMMENDATION_PATTERNS = [
    r"\bmind sharing\b",
    r"\bcan you share\b",
    r"\bwhat do you recommend\b",
    r"\bany advice\b",
    r"\bhow are you doing\b",
    r"\bhow do you\b",
    r"\bwhat tool\b",
    r"\bwhere can\b",
    r"\blet me know\b",
    r"แนะนำ",
    r"มีใครแนะนำ",
    r"ควรใช้",
    r"ใช้อะไรดี",
    r"ตัวไหนดี",
    r"มีเครื่องมือไหม",
    r"มีวิธีไหม",
    r"มีใครเคยใช้",
    r"ขอคำแนะนำ",
    r"ช่วยแนะนำ",
    r"อยากให้",
    "recommend",
    "advice",
    "what tool",
    "can you share",
    "แนะนำ",
    "ขอคำแนะนำ",
    "ใช้อะไรดี",
    "ตัวไหนดี",
]
CURRENT_USAGE_PATTERNS = [
    r"\bi (use|used|made|built|wrote|created|have|keep|run|save|track|made)\b",
    r"\bi'm using\b",
    r"\bi am using\b",
    r"\bi've (made|built|used|had)\b",
    r"\bwe (use|built|made|created)\b",
    r"\bbuilt it\b",
    r"\bfully vibecoded\b",
    r"\bvibecoded\b",
    r"ใช้",
    r"กำลังใช้",
    r"ใช้งาน",
    r"เอามาใช้",
    r"สร้าง",
    r"ทำ",
    r"พัฒนา",
    r"ทำระบบ",
    r"ทำงานด้วย",
    r"ใช้ทุกวัน",
    r"ใช้งานทุกวัน",
    r"เอาไปใช้",
    "i use",
    "i built",
    "i made",
    "i'm using",
    "ใช้",
    "กำลังใช้",
    "สร้าง",
    "พัฒนา",
    "ทำระบบ",
    "ใช้งาน",
]

INTENT_PATTERNS = {
    "Comparison": {
        "score": 8,
        "patterns": [
            r"\bvs\b",
            r"\bversus\b",
            r"\bcompare\b",
            r"\bcomparison\b",
            r"\bbetter than\b",

            # TH
            r"เปรียบเทียบ",
            r"เทียบ",
            r"ดีกว่าไหม",
            r"ต่างจาก",
            r"เหมือนกับ",
            r"claude\s*vs",
            r"chatgpt\s*vs",
        ],
    },

    "Problem": {
        "score": 7,
        "patterns": [
            r"\bbug\b",
            r"\berror\b",
            r"\bissue\b",
            r"\bproblem\b",
            r"\bfailed\b",
            r"\bnot working\b",
            r"\bcan't\b",
            r"\bcannot\b",
            r"\bslow\b",
            r"\bcrash\b",
            r"\bhallucinat",

            # TH
            r"มีปัญหา",
            r"ปัญหา",
            r"บั๊ก",
            r"เออเรอร์",
            r"ช้ามาก",
            r"ช้า",
            r"ค้าง",
            r"ล่ม",
            r"พัง",
            r"ใช้งานไม่ได้",
            r"ไม่ทำงาน",
            r"ไม่เวิร์ค",
            r"ไม่ถูกต้อง",
            r"มั่ว",
        ],
    },

    "Recommendation Request": {
        "score": 6,
        "patterns": [
            r"\brecommend\b",
            r"\badvice\b",
            r"\bwhat tool\b",
            r"\bany suggestion",
            r"\bcan you recommend\b",

            # TH
            r"แนะนำ",
            r"ขอคำแนะนำ",
            r"มีใครแนะนำ",
            r"ใช้อะไรดี",
            r"ตัวไหนดี",
            r"ควรใช้",
            r"เครื่องมือไหนดี",
            r"มีใครเคยใช้",
            r"แนะนำหน่อย",
        ],
    },

    "Desired Usage": {
        "score": 5,
        "patterns": [
            r"\bi want\b",
            r"\bi need\b",
            r"\bi would like\b",
            r"\blooking for\b",
            r"\bplan to\b",
            r"\bwant to\b",

            # TH
            r"อยาก",
            r"ต้องการ",
            r"กำลังหา",
            r"กำลังมองหา",
            r"วางแผนจะ",
            r"ตั้งใจจะ",
            r"กำลังจะ",
            r"ช่วย.*ได้ไหม",
            r"ทำ.*ได้ไหม",
            r"สร้าง.*ได้ไหม",
        ],
    },

    "Current Usage": {
        "score": 4,
        "patterns": [
            r"\bi use\b",
            r"\bi'm using\b",
            r"\bi am using\b",
            r"\bi built\b",
            r"\bi made\b",
            r"\bwe use\b",
            r"\bwe built\b",

            # TH
            r"ผมใช้",
            r"ฉันใช้",
            r"เราใช้",
            r"กำลังใช้",
            r"ใช้งาน",
            r"ใช้ทุกวัน",
            r"ใช้งานทุกวัน",
            r"กำลังพัฒนา",
            r"พัฒนาอยู่",
            r"สร้างอยู่",
            r"ทำระบบ",
            r"สร้าง workflow",
        ],
    },
}

USER_PERSONA_KEYWORDS = {
    "Developer": [
        "developer",
        "developers",
        "engineer",
        "software engineer",
        "programmer",
        "coding",
        "code",
        "coder",
        "software",
        "python",
        "javascript",
        "api",
        "github",
        "claude code",
        "frontend",
        "backend",
        "prompt engineering",
    ],
    "Content Creator": [
        "youtube",
        "tiktok",
        "instagram",
        "facebook",
        "creator",
        "content creator",
        "influencer",
        "creators",
        "influencers",
        "video",
        "shorts",
        "reels",
        "blog",
        "newsletter",
    ],
    "Marketer": [
        "marketer",
        "marketing",
        "seo",
        "advertising",
        "ads",
        "ad",
        "ad copy",
        "campaign",
        "funnel",
        "lead generation",
        "sales page",
        "email marketing",
    ],
    "Founder": [
        "startup",
        "founder",
        "solo builder",
        "indie hacker",
        "entrepreneur",
        "mvp",
        "launch",
        "saas",
        "micro saas",
        "agency",
        "business",
        "business owner",
    ],
    "Researcher": [
        "research",
        "researcher",
        "paper",
        "papers",
        "academic",
        "literature review",
        "citation",
        "journal",
        "thesis",
        "publication",
        "scientific",
    ],
    "Student": [
        "student",
        "students",
        "exam",
        "study",
        "course",
        "homework",
        "school",
        "university",
        "learning",
    ],
    "Business Owner": [
        "business owner",
        "my business",
        "my company",
        "company",
        "business",
        "operations",
        "crm",
        "customer",
        "client",
        "sales",
        "revenue",
        "employee",
    ],
    "Freelancer": [
        "freelancer",
        "freelance",
        "freelancing",
        "independent contractor",
        "client work",
        "contractor",
        "consultant",
        "consulting",
    ],
}


ADOPTION_STAGE_KEYWORDS = {
    "Beginner": [
        "new to claude",
        "just started",
        "beginner",
        "learning",
        "first day",
        "first time",
        "no experience",
        "zero experience",
    ],
    "Intermediate": [
        "daily use",
        "daily",
        "every day",
        "everyday",
        "workflow",
        "routine",
        "regular use",
        "regularly",
        "use it daily",
        "use it every day",
    ],
    "Advanced": [
        "mcp",
        "agent",
        "api",
        "automation",
        "production",
        "deployment",
        "deploy",
        "deployed",
        "prod",
        "claude code",
        "webhook",
        "sdk",
        "multi agent",
        "pipeline",
    ],
}


PAIN_POINT_KEYWORDS = {
    "Pricing": ["too expensive", "expensive", "pricing", "subscription", "cost", "paid", "price", "too pricey"],
    "Rate Limit": ["rate limit", "rate limited", "limit reached", "usage limit", "message limit", "quota", "usage cap"],
    "Context Limit": ["context window", "context limit", "context length", "too much context", "context too short"],
    "Hallucination": ["hallucination", "hallucinate", "made up", "fabricated"],
    "Accuracy": ["wrong", "inaccurate", "not accurate", "incorrect", "bad answer"],
    "Performance": ["slow", "performance", "lag", "latency", "takes forever", "sluggish"],
    "Missing Feature": ["wish", "missing feature", "need this", "feature request", "doesn't have", "lack", "missing"],
}


FEATURE_REQUEST_KEYWORDS = {
    "Memory": [
        "memory", 
        "remember", 
        "persistent memory", 
        "long term memory"
    ],
    "Image Generation": [
        "image generation", 
        "generate image", 
        "make image",
        "create image", 
        "picture", 
        "images"
    ],
    "Web Search": [
        "web search", 
        "search the web", 
        "internet search", 
        "browser", 
        "browse", 
        "web browsing"
    ],
    "Voice": [
        "voice", 
        "voice mode", 
        "speech", 
        "talk", 
        "transcribe", 
        "tts", 
        "dictation"
    ],
    "API": [
        "api", 
        "sdk", 
        "webhook", 
        "integration"
    ],
    "Agent": [
        "agent", 
        "ai agent", 
        "autonomous agent", 
        "multi agent"
    ],
    "Mobile App": [
        "mobile app", 
        "ios app", 
        "android app", 
        "phone app", 
        "app store"
    ],
    "UI/UX": [
        "ui", 
        "ux", 
        "interface", 
        "design", 
        "screenshot",  
        "layout"
    ],
}


FEATURE_REQUEST_CUES = [
    "wish",
    "want",
    "wanted",
    "wants",
    "need",
    "needed",
    "needs",
    "missing feature",
    "feature request",
    "would be useful",
    "should add",
    "could add",
    "better",
    "improve",
    "add",
    "support",
    "อยากให้",
    "ออกมาเป็น",
]


USE_CASE_DOMAIN_KEYWORDS = {
    "Career": [
        "resume", 
        "cv", 
        "interview", 
        "job", 
        "career", 
        "job offer", 
        "hiring", 
        "recruiting"
    ],
    "Business": [
        "marketing",
        "sales",
        "agency",
        "client",
        "revenue",
        "crm",
        "business",
        "operations",
        "invoice",
        "billing",
        "customer",
    ],
    "Education": [
        "exam", 
        "study", 
        "learn", 
        "learning", 
        "course", 
        "student", 
        "tutorial", 
        "school", 
        "homework"
    ],
    "Technical": [
        "python",
        "api", 
        "coding", 
        "developer", 
        "code", 
        "debug", 
        "github", 
        "database", 
        "deployment"
    ],
    "Content Creation": [
        "youtube", 
        "tiktok", 
        "content", 
        "creator", 
        "blog", 
        "article", 
        "newsletter", 
        "social media"
        ],
    "Research": [
        "research", 
        "paper", 
        "academic", 
        "literature review", 
        "citation", 
        "journal"
    ],
    "Personal Productivity": [
        "calendar",
        "shopping list",
        "recipe",
        "notes",
        "bookmark",
        "planner",
        "time zone",
        "fitness",
        "migraine",
        "personal",
    ],
}


SHORT_KEYWORD_ALLOWLIST = {
    "ai",
    "bi",
    "db",
    "hr",
    "ui",
    "ux",
    "ad",
    "app",
    "api",
    "css",
    "csv",
    "gpt",
    "mcp",
    "php",
    "pro",
    "seo",
    "sql",
}


def normalize_line(line: str) -> str:
    line = html.unescape(line)
    return line.replace("\ufeff", "").strip()


def is_time_line(line: str) -> bool:
    line = re.sub(r"^edited\s+", "", line.lower()).strip()
    return bool(re.match(r"^\d+\s*(?:mo|yr|[smhdwy])(?:\s+ago)?$", line)) or bool(
        re.match(
            r"^\d+\s*(?:sec|secs|second|seconds|min|mins|minute|minutes|hr|hrs|hour|hours|day|days|wk|wks|week|weeks|mon|mons|month|months|year|years)\s+ago$",
            line,
        )
    )


def is_comment_start(lines, index: int) -> bool:
    if index + 2 >= len(lines):
        return False
    name = lines[index].strip()
    separator = lines[index + 1].strip()
    if bool(name) and separator in {"•", "�"}:
        return is_time_line(lines[index + 2].strip())
    return bool(name) and lines[index + 1].strip() == "•" and is_time_line(lines[index + 2].strip())


def is_ui_noise(line: str) -> bool:
    if not line:
        return True
    checks = [
        r"^u/.+\s+avatar$",
        r"^\d+$",
        r"^\d+\s+more$",
        r"^\d+\s+more\s+repl(?:y|ies)$",
        r"^promoted$",
        r"^learn more$",
        r"^thumbnail image:",
        r"^upvote$",
        r"^downvote$",
        r"^share$",
        r"^reply$",
        r"^award$",
    ]
    low = line.lower()
    return any(re.search(pattern, low) for pattern in checks)


def strip_trailing_noise(body_lines):
    lines = [normalize_line(line) for line in body_lines]
    while lines and is_ui_noise(lines[-1]):
        lines.pop()
    return lines


def extract_comments(text: str):
    raw_lines = [normalize_line(line) for line in text.splitlines()]
    comments = []
    index = 0
    while index < len(raw_lines):
        if not is_comment_start(raw_lines, index):
            index += 1
            continue

        author = raw_lines[index]
        timestamp = raw_lines[index + 2]
        start = index + 3
        end = start
        while end < len(raw_lines) and not is_comment_start(raw_lines, end):
            end += 1

        body_lines = strip_trailing_noise(raw_lines[start:end])
        body_lines = [line for line in body_lines if not is_ui_noise(line)]
        text = re.sub(r"\s+", " ", " ".join(body_lines)).strip()
        if text:
            comments.append({"author": author, "timestamp": timestamp, "text": text})
        index = end
    return comments


def normalize_match_text(value: str) -> str:
    value = value.lower()
    value = (
        value.replace("’", "'")
        .replace("‘", "'")
        .replace("`", "'")
        .replace("�", "")
    )
    value = re.sub(r"[\u2010-\u2015_/]+", " ", value)
    for pattern, replacement in TEXT_NORMALIZATIONS:
        value = re.sub(pattern, replacement, value)
    return re.sub(r"\s+", " ", value).strip()


@lru_cache(maxsize=None)
def keyword_variants(term: str):
    term = normalize_match_text(term)
    if not term:
        return []

    variants = {term}
    words = term.split()
    last = words[-1]
    if re.fullmatch(r"[a-z][a-z0-9-]*", last):
        if last.endswith("y") and len(last) > 2 and last[-2] not in "aeiou":
            variants.add(" ".join(words[:-1] + [last[:-1] + "ies"]))
        elif last.endswith(("ch", "sh", "x", "z")) or (last.endswith("s") and not last.endswith("ss")):
            variants.add(" ".join(words[:-1] + [last + "es"]))
        elif not last.endswith("s"):
            variants.add(" ".join(words[:-1] + [last + "s"]))

        if last.endswith("ies") and len(last) > 3:
            variants.add(" ".join(words[:-1] + [last[:-3] + "y"]))
        elif last.endswith(("ches", "shes", "xes", "zes")) and len(last) > 4:
            variants.add(" ".join(words[:-1] + [last[:-2]]))
        elif last.endswith("sses") and len(last) > 5:
            variants.add(" ".join(words[:-1] + [last[:-2]]))
        elif last.endswith("s") and not last.endswith(("ss", "us", "is")) and len(last) > 3:
            variants.add(" ".join(words[:-1] + [last[:-1]]))

    return sorted(variants, key=len, reverse=True)


@lru_cache(maxsize=None)
def keyword_regex(variant: str):
    if variant == "$":
        return re.escape(variant)
    # Count ASCII alphanumeric chars; also count non-ASCII (Thai etc.) chars toward
    # the length guard so Thai keywords are not silently discarded.
    bare_ascii = re.sub(r"[^a-z0-9]", "", variant)
    non_ascii_len = len(re.sub(r"[\x00-\x7f]", "", variant))
    if (len(bare_ascii) + non_ascii_len) < 3 and variant not in SHORT_KEYWORD_ALLOWLIST:
        return None

    pattern = re.escape(variant).replace(r"\ ", r"\s+")
    if re.match(r"[a-z0-9]", variant):
        pattern = r"(?<![a-z0-9])" + pattern
    if re.search(r"[a-z0-9]$", variant):
        pattern += r"(?![a-z0-9])"
    return pattern


@lru_cache(maxsize=None)
def compiled_keyword_regex(variant: str):
    pattern = keyword_regex(variant)
    return re.compile(pattern) if pattern else None


def contains_term(text: str, term: str) -> bool:
    text = normalize_match_text(text)
    for variant in keyword_variants(term):
        pattern = compiled_keyword_regex(variant)
        if pattern and pattern.search(text):
            return True
    return False


def score_keyword_groups(text: str, keyword_groups):
    low = normalize_match_text(text)
    scored = []
    for label, terms in keyword_groups.items():
        matches = []
        score = 0
        for term in terms:
            if contains_term(low, term):
                matches.append(term)
                score += 2 if " " in term or len(term) > 7 else 1
        if score:
            scored.append((score, label, matches))
    scored.sort(key=lambda item: (-item[0], item[1]))
    return scored


def validate_category_schema():
    """Keep segment/pain/request dimensions out of Main Category."""
    forbidden = sorted(set(CATEGORY_KEYWORDS) & BANNED_MAIN_CATEGORIES)
    if forbidden:
        names = ", ".join(forbidden)
        raise ValueError(f"Forbidden main categories found in CATEGORY_KEYWORDS: {names}")


def score_categories(text: str):
    low = normalize_match_text(text)
    scored = []
    for main, subcats in CATEGORY_KEYWORDS.items():
        if main in BANNED_MAIN_CATEGORIES:
            continue
        for sub, terms in subcats.items():
            matches = []
            score = 0
            for term in terms:
                if contains_term(low, term):
                    matches.append(term)
                    score += 2 if " " in term or len(term) > 7 else 1
            if score:
                scored.append((score, main, sub, matches))
    scored.sort(key=lambda item: (-item[0], item[1], item[2]))
    return scored


def choose_category(text: str):
    scored = score_categories(text)
    if not scored:
        return "Other", "General", [], 0

    top_score, main, sub, matches = scored[0]

    # Product/app-building comments often mention generic "app"; prefer the
    # domain when the text clearly describes personal productivity/life-admin.
    domain_first = [
        item
        for item in scored
        if item[1] in {"Productivity", "Business Operations", "Data Analysis", "AI Automation"}
        and item[0] >= max(2, top_score - 1)
    ]
    if main == "Coding & Development" and sub == "Full Stack" and domain_first:
        top_score, main, sub, matches = domain_first[0]

    return main, sub, matches, top_score


def classify_intent(text: str):
    low = normalize_match_text(text)

    scores = {
        "Current Usage": 1,
        "Desired Usage": 2,
        "Recommendation Request": 3,
        "Problem": 4,
        "Comparison": 5,
    }

    for intent, config in INTENT_PATTERNS.items():
        weight = config["score"]

        for pattern in config["patterns"]:
            if re.search(pattern, low):
                scores[intent] += weight

    # --------------------------------------------------
    # Context bonuses
    # --------------------------------------------------

    # Question mark should NOT dominate
    if text.strip().endswith("?"):
        scores["Recommendation Request"] += 1

    # Strong Thai recommendation phrases
    if re.search(
        r"(ใช้อะไรดี|ตัวไหนดี|มีใครแนะนำ|ขอคำแนะนำ)",
        low,
    ):
        scores["Recommendation Request"] += 4

    # Strong desired intent
    if re.search(
        r"(อยาก|ต้องการ|กำลังหา|กำลังมองหา)",
        low,
    ):
        scores["Desired Usage"] += 4

    # Current usage booster
    if re.search(
        r"(ใช้ทุกวัน|กำลังพัฒนา|กำลังใช้|เราใช้|ผมใช้|ฉันใช้)",
        low,
    ):
        scores["Current Usage"] += 3

    # --------------------------------------------------
    # Select winner
    # --------------------------------------------------

    best_score = max(scores.values())

    if best_score == 0:
        return "Current Usage"

    winners = {
        intent
        for intent, score in scores.items()
        if score == best_score
    }

    # Priority order minimizes false positives
    priority = [
        "Comparison",
        "Problem",
        "Recommendation Request",
        "Desired Usage",
        "Current Usage",
    ]

    for intent in priority:
        if intent in winners:
            return intent

    return "Current Usage"


def classify_sentiment(text: str, intent: str):
    low = text.lower()
    pos = sum(1 for term in POSITIVE_TERMS if contains_term(low, term))
    neg = sum(1 for term in NEGATIVE_TERMS if contains_term(low, term))
    if intent == "Problem":
        neg += 1
    if pos > neg:
        return "Positive"
    if neg > pos:
        return "Negative"
    return "Neutral"


# New classifier: infers the likely audience/persona from role, channel, and
# work-context words. It is intentionally an output column only, not a category.
def classify_user_persona(text):
    scored = score_keyword_groups(text, USER_PERSONA_KEYWORDS)
    return scored[0][1] if scored else "General User"


# New classifier: labels maturity of Claude usage. With no explicit signal,
# Intermediate is the neutral dashboard-friendly default between novice and power user.
def classify_adoption_stage(text):
    scored = score_keyword_groups(text, ADOPTION_STAGE_KEYWORDS)
    if not scored:
        return "Intermediate"

    best_score = scored[0][0]
    winners = {label for score, label, _ in scored if score == best_score}
    for label in ["Advanced", "Intermediate", "Beginner"]:
        if label in winners:
            return label
    return scored[0][1]


# New classifier: extracts the primary complaint type when the comment signals
# pricing, limits, accuracy, speed, or missing capability pain.
def classify_pain_point(text):
    scored = score_keyword_groups(text, PAIN_POINT_KEYWORDS)
    return scored[0][1] if scored else "Other"


def has_feature_request_cue(text):
    low = normalize_match_text(text)
    if any(contains_term(low, cue) for cue in FEATURE_REQUEST_CUES):
        return True
    return bool(
        re.search(r"\bwish\b.+\b(had|could|would)\b", low)
        or re.search(r"\bneed(?:s|ed)?\b.+\b(better|feature|support)\b", low)
        or re.search(r"\bwant(?:s|ed)?\b.+\b(better|feature|support)\b", low)
    )


# New classifier: identifies the requested feature only when the text has a
# request cue, preventing normal API/agent usage from becoming feature requests.
def classify_feature_request(text):
    if not has_feature_request_cue(text):
        return "Unknown"
    scored = score_keyword_groups(text, FEATURE_REQUEST_KEYWORDS)
    return scored[0][1] if scored else "Unknown"


# New classifier: maps the comment into a broader dashboard domain that cuts
# across Main Category, such as career, business, technical, or productivity.
def classify_use_case_domain(text):
    scored = score_keyword_groups(text, USE_CASE_DOMAIN_KEYWORDS)
    return scored[0][1] if scored else "General"


def summarize(text: str, main: str, sub: str, intent: str):
    cleaned = re.sub(r"\s+", " ", text).strip()
    sentence = re.split(r"(?<=[.!?])\s+", cleaned)[0]
    if len(sentence) > 170:
        sentence = sentence[:167].rsplit(" ", 1)[0] + "..."
    if intent == "Current Usage":
        prefix = f"Uses Claude for {sub.lower()}"
    elif intent == "Desired Usage":
        prefix = f"Wants help with {sub.lower()}"
    elif intent == "Problem":
        prefix = f"Reports an issue around {sub.lower()}"
    elif intent == "Comparison":
        prefix = f"Compares Claude around {sub.lower()}"
    else:
        prefix = f"Asks about {sub.lower()}"
    if main == "Other":
        prefix = "General Claude-related comment"
    return f"{prefix}: {sentence}"


def confidence(score: int, intent: str, sentiment: str, text: str):
    conf = 0.68
    if score >= 6:
        conf += 0.25
    elif score >= 3:
        conf += 0.18
    elif score >= 1:
        conf += 0.10
    if intent in {"Comparison", "Problem", "Recommendation Request"}:
        conf += 0.03
    if sentiment != "Neutral":
        conf += 0.02
    if len(text.split()) < 6:
        conf -= 0.08
    return f"{min(0.99, max(0.55, conf)):.2f}"


def classify_comment(comment):
    text = comment["text"]
    main, sub, matches, score = choose_category(text)
    intent = classify_intent(text)
    sentiment = classify_sentiment(text, intent)
    user_persona = classify_user_persona(text)
    adoption_stage = classify_adoption_stage(text)
    pain_point = classify_pain_point(text)
    feature_request = classify_feature_request(text)
    use_case_domain = classify_use_case_domain(text)

    all_matches = []
    for _, _, _, terms in score_categories(text)[:3]:
        all_matches.extend(terms)
    keyword_list = []
    for term in all_matches + matches:
        label = term.strip()
        if label and label not in keyword_list:
            keyword_list.append(label)
    if not keyword_list:
        keyword_list = [sub] if sub != "General" else [main]

    return {
        "Main Category": main,
        "Sub Category": sub,
        "Intent": intent,
        "Sentiment": sentiment,
        "User Persona": user_persona,
        "Adoption Stage": adoption_stage,
        "Pain Point": pain_point,
        "Feature Request": feature_request,
        "Use Case Domain": use_case_domain,
        "Keywords": ", ".join(keyword_list[:8]),
        "Confidence": confidence(score, intent, sentiment, text),
        "Summary": summarize(text, main, sub, intent),
    }


def write_csv(rows):
    with CSV_PATH.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=COLUMNS)
        writer.writeheader()
        writer.writerows(rows)


def col_name(index: int):
    name = ""
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def cell_xml(value, row_index: int, col_index: int):
    ref = f"{col_name(col_index)}{row_index}"
    value = "" if value is None else str(value)
    if row_index > 1 and COLUMNS[col_index - 1] == "Confidence":
        return f'<c r="{ref}"><v>{escape(value)}</v></c>'
    return f'<c r="{ref}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'


def sheet_xml(rows):
    xml_rows = []
    all_rows = [dict(zip(COLUMNS, COLUMNS))] + rows
    for row_index, row in enumerate(all_rows, 1):
        cells = "".join(cell_xml(row.get(column, ""), row_index, col_index) for col_index, column in enumerate(COLUMNS, 1))
        xml_rows.append(f'<row r="{row_index}">{cells}</row>')
    last_ref = f"{col_name(len(COLUMNS))}{len(all_rows)}"
    cols = "".join(
        f'<col min="{index}" max="{index}" width="{COLUMN_WIDTHS.get(column, 18)}" customWidth="1"/>'
        for index, column in enumerate(COLUMNS, 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<dimension ref="A1:{last_ref}"/>'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        f'<cols>{cols}</cols>'
        f'<sheetData>{"".join(xml_rows)}</sheetData>'
        f'<autoFilter ref="A1:{col_name(len(COLUMNS))}1"/>'
        '</worksheet>'
    )


def write_xlsx(rows):
    files = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '</Types>'
        ),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>'
        ),
        "xl/_rels/workbook.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>'
        ),
        "xl/workbook.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Cleaned Comments" sheetId="1" r:id="rId1"/></sheets>'
            '</workbook>'
        ),
        "xl/styles.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            '</styleSheet>'
        ),
        "xl/worksheets/sheet1.xml": sheet_xml(rows),
    }

    with zipfile.ZipFile(XLSX_PATH, "w", zipfile.ZIP_DEFLATED) as workbook:
        for name, content in files.items():
            workbook.writestr(name, content)


def write_summary(comments, rows):
    main_counts = Counter(row["Main Category"] for row in rows)
    sub_counts = Counter(row["Sub Category"] for row in rows)
    intent_counts = Counter(row["Intent"] for row in rows)
    sentiment_counts = Counter(row["Sentiment"] for row in rows)
    persona_counts = Counter(row["User Persona"] for row in rows)
    adoption_counts = Counter(row["Adoption Stage"] for row in rows)
    pain_counts = Counter(row["Pain Point"] for row in rows)
    feature_counts = Counter(row["Feature Request"] for row in rows)
    domain_counts = Counter(row["Use Case Domain"] for row in rows)
    low_conf = sum(float(row["Confidence"]) < 0.70 for row in rows)
    lines = [
        f"Parsed comments: {len(comments)}",
        f"Output rows: {len(rows)}",
        f"Rows below 0.70 confidence: {low_conf}",
        "",
        "Top Main Categories:",
        *[f"- {name}: {count}" for name, count in main_counts.most_common(10)],
        "",
        "Top Sub Categories:",
        *[f"- {name}: {count}" for name, count in sub_counts.most_common(10)],
        "",
        "Intent Counts:",
        *[f"- {name}: {count}" for name, count in intent_counts.most_common()],
        "",
        "Sentiment Counts:",
        *[f"- {name}: {count}" for name, count in sentiment_counts.most_common()],
        "",
        "User Persona Counts:",
        *[f"- {name}: {count}" for name, count in persona_counts.most_common()],
        "",
        "Adoption Stage Counts:",
        *[f"- {name}: {count}" for name, count in adoption_counts.most_common()],
        "",
        "Pain Point Counts:",
        *[f"- {name}: {count}" for name, count in pain_counts.most_common()],
        "",
        "Feature Request Counts:",
        *[f"- {name}: {count}" for name, count in feature_counts.most_common()],
        "",
        "Use Case Domain Counts:",
        *[f"- {name}: {count}" for name, count in domain_counts.most_common()],
    ]
    SUMMARY_PATH.write_text("\n".join(lines), encoding="utf-8")

def preprocess(text):
    text = text.lower()

    common_hits = []

    for kw in COMMON_THAI_KEYWORDS:
        if kw in text:
            common_hits.append(kw)

    return common_hits

def main():
    raw = RAW_PATH.read_text(encoding="utf-8-sig")
    comments = extract_comments(raw)
    rows = [classify_comment(comment) for comment in comments]
    write_csv(rows)
    write_xlsx(rows)
    write_summary(comments, rows)
    print(f"Parsed comments: {len(comments)}")
    print(f"Wrote: {CSV_PATH}")
    print(f"Wrote: {XLSX_PATH}")
    print(f"Wrote: {SUMMARY_PATH}")


if __name__ == "__main__":
    main()

[![Live Demo](https://img.shields.io/badge/Live%20Demo-Streamlit-red)](https://facebook-dashboard-k7jchvbyqhjysovpmiefhq.streamlit.app/)

[![Python](https://img.shields.io/badge/Python-3.12-blue)]()

[![Streamlit](https://img.shields.io/badge/Streamlit-1.x-red)]()
# 📊 Facebook Comment Intelligence Dashboard

A Streamlit dashboard for parsing, classifying, and analyzing Facebook comments.

🌐 **Live Demo**

**Try it here:**  
👉 https://facebook-dashboard-k7jchvbyqhjysovpmiefhq.streamlit.app/

The application can:

- Parse Facebook copied comments
- Clean Facebook UI noise
- Classify comments into categories
- Analyze user intent
- Analyze sentiment
- Export results to CSV and Excel
- Visualize results with charts

---

## ✨ Features

✅ Parse Facebook copied comments

✅ Remove UI noise

- Like
- Reply
- Share
- Top Fan
- Author
- Timestamp

✅ Comment Classification

- Main Category
- Sub Category
- Intent
- Sentiment
- User Persona
- Pain Point
- Feature Request
- Adoption Stage

✅ Dashboard

- KPI Cards
- Main Category Chart
- Intent Chart
- Sentiment Chart

✅ Export

- CSV
- Excel

---

## 📁 Project Structure

```text
Facebook-Dashboard

├── app.py
├── clean_claude_comments2.py
├── clean_facebook_comments2.py
├── requirements.txt
├── README.md
└── .gitignore
```

---

## 🚀 Installation

Clone repository

```bash
git clone https://github.com/pern41/Facebook-Dashboard.git

cd Facebook-Dashboard
```

Install dependencies

```bash
pip install -r requirements.txt
```

---

## ▶️ Run Locally

```bash
streamlit run app.py
```

or

```bash
python -m streamlit run app.py
```

Open browser

```text
http://localhost:8501
```

---

## 📋 How To Use

### 1. Copy Facebook comments

Example

```text
John Doe

Top Fan

This tool is amazing!

3h

Like
Reply

Jane Smith

Can it export CSV?

5h

Like
Reply
```

---

### 2. Paste into Dashboard

Paste all copied comments into:

```text
Paste Facebook Comments
```

---

### 3. Click Analyze

The dashboard will:

- Parse comments
- Remove UI noise
- Extract author
- Extract timestamp
- Classify comments
- Generate charts

---

## 📊 Output Example

| Main Category | Intent | Sentiment |
|---------------|--------|----------|
| AI Models | Current Usage | Positive |
| Coding | Problem | Negative |
| Research | Desired Usage | Neutral |

---

## 📥 Export

Export analyzed comments as:

- CSV
- Excel (.xlsx)

---

## 🛠 Tech Stack

- Python
- Streamlit
- Pandas
- OpenPyXL

---

## 📌 Roadmap

### Dashboard

- [x] Comment Parser
- [x] Classification
- [x] Sentiment Analysis
- [x] CSV Export
- [x] Excel Export
- [x] Charts

### Future Features

- [ ] Keyword Editor UI
- [ ] Notion Integration
- [ ] Pie Charts
- [ ] Dark Mode
- [ ] Admin Panel
- [ ] Keyword CSV Management
- [ ] Multi-language Support

---

## 👨‍💻 Author

GitHub:

https://github.com/pern41

Project:

https://github.com/pern41/Facebook-Dashboard

---

## ⭐ Support

If you find this project useful, please give it a ⭐ on GitHub.
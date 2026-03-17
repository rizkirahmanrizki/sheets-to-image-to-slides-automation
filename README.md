# 📊 Sheets → Slides Automation (Google Apps Script)

Automates the process of converting Google Sheets data into presentation-ready slides.

This script exports selected ranges from Google Sheets, converts them into images, and inserts them into Google Slides — enabling automated reporting workflows.

---

# 🎯 Objective

To eliminate manual steps in reporting workflows by:

* exporting sheet data
* converting it into visual format
* updating presentation slides automatically

---

# 🔗 Workflow Overview

```text
Google Sheets Data
        ↓
Export as PDF
        ↓
Convert to Image
        ↓
Insert into Slides
        ↓
Updated Presentation
```

---

# ⚙️ Key Features

* Export specific sheet ranges
* Automatically convert data into slide-ready visuals
* Replace existing slide images
* Auto-scale and position content
* Supports multiple sections (jobs)
* Fully automated reporting pipeline

---

# 🧩 Use Cases

* weekly/monthly reporting decks
* KPI dashboards in presentations
* automated executive reports
* replacing manual copy-paste workflows
* data snapshot presentations

---

# 📁 Project Structure

```text
sheets-to-slides-automation
│
├── src
│   └── syncSheetsToSlides.js
│
├── README.md
└── appsscript.json
```

---

# ⚙️ Configuration

Update the script with your own IDs:

```javascript
SPREADSHEET_ID = "YOUR_SPREADSHEET_ID"
PRESENTATION_ID = "YOUR_PRESENTATION_ID"
```

---

# 🛠 Job Configuration

Each job defines:

* sheet (gid)
* range to export
* target slide index
* label

Example:

```javascript
{
  gid: "SHEET_GID",
  range: "A1:D20",
  slideIndex: 0,
  name: "Section Name"
}
```

---

# 🚀 Why This Project Matters

This project demonstrates:

* workflow automation using Google Apps Script
* integration between Sheets, Drive, and Slides
* dynamic content generation
* replacing manual reporting processes

---

# 🧠 Skills Demonstrated

* automation engineering
* API usage (Google services)
* data-to-visual transformation
* scalable reporting workflows

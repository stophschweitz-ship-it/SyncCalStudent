# Unified Student Calendar

A flat, minimalist student-focused HTML calendar unifying **Due Dates**, **Clubs**, **Mentorship**, **Email Dates**, and **Sports** across five live integrations, with customizable activity colors and 5 distinct calendar views (**Month**, **Week**, **To-Do**, **Orbit**, and **Stream**).

---

## Quick Start

1. Start the local sync server:
   ```bash
   python3 server.py
   ```
2. Open **`http://localhost:8080`** in your browser.
3. Enter your name, connect your logins, and customize your activity colors in the left sidebar.

---

## Service Integration Guides (README per Service)

- [**Canvas LMS Setup Guide** (`README_CANVAS.md`)](./README_CANVAS.md) — Sync via Canvas Calendar `.ics` Feed or Canvas REST API Access Token.
- [**Virtual Virginia Setup Guide** (`README_VIRTUAL_VIRGINIA.md`)](./README_VIRTUAL_VIRGINIA.md) — Sync VVA pacing guides, quizzes, and seminars via `.ics` Feed or VVA API Token.
- [**Gmail Setup Guide** (`README_GMAIL.md`)](./README_GMAIL.md) — Sync via Gmail IMAP App Password, Google Calendar `.ics` Feed, or Google OAuth Token.
- [**Microsoft Outlook Setup Guide** (`README_OUTLOOK.md`)](./README_OUTLOOK.md) — Sync via Published Outlook `.ics` URL, Microsoft Graph API Token, or Office 365 IMAP.
- [**Microsoft Teams Setup Guide** (`README_TEAMS.md`)](./README_TEAMS.md) — Sync Teams mentorship calls, club meetings, and sports schedules via Microsoft Graph or Exchange `.ics` Feed.

---

## Customizing Activity Colors

- **By Activity Category**: In the left sidebar under **Activities & Colors**, click the colored circle on the right side of **Due Dates**, **Clubs**, **Mentorship**, **Email Dates**, or **Sports** to pick any custom color.
- **By Individual Activity**: Click any date to view its schedule in the right-hand **Selected Day** panel, then click the color swatch on any individual activity (or select a color when creating a `+ New` activity).

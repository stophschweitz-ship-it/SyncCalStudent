# Unified Student Calendar

A flat, minimalist student-focused HTML calendar unifying **Due Dates**, **Clubs**, **Mentorship**, **Email Dates**, and **Sports** across your Canvas, Gmail, Outlook, and Teams calendars, with customizable activity colors and 5 distinct calendar views (**Month**, **Week**, **To-Do**, **Orbit**, and **Stream**).

---

## Quick Start

1. Start the local sync server:
   ```bash
   python3 server.py
   ```
2. Open **`http://localhost:8080`** in your browser.
3. Enter your name, connect your logins, and customize your activity colors in the left sidebar.

---

## Customizing Activity Colors

- **By Activity Category**: In the left sidebar under **Activities & Colors**, click the colored circle on the right side of **Due Dates**, **Clubs**, **Mentorship**, **Email Dates**, or **Sports** to pick any custom color.
- **By Individual Activity**: Click any date to view its schedule in the right-hand **Selected Day** panel, then click the color swatch on any individual activity (or select a color when creating a `+ New` activity).

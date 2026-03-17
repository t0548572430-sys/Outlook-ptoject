# 📧 Outlook Bridge: Local Draft Automator

A specialized Web-to-Desktop integration solution designed for recruitment agencies to streamline candidate CV distribution via Microsoft Outlook.

## 🚀 The Challenge
Recruiters need to send personalized emails with CV attachments to multiple recipients. However, modern web browsers operate in a **Sandbox environment**, which imposes several restrictions:
* **Security Constraints:** Web apps cannot directly trigger local software (like Outlook) with attachments.
* **Protocol Limitations:** The standard `mailto:` protocol does not support file attachments or multiple separate drafts.
* **Local File Access:** Browsers cannot programmatically "grab" files from a user's local disk to attach them to an external mail client.

## 💡 The Solution
This project implements a **bridge architecture** that bypasses browser limitations by using a local listener/integration component. This allows the web interface to communicate with the Windows COM API (Component Object Model) to automate Outlook.

### Key Features:
* **Mass Draft Creation:** Generates individual, separate drafts for multiple recipients simultaneously.
* **Automated Attachments:** Seamlessly attaches local files to the Outlook draft (not as links, but as actual attachments).
* **User-in-the-Loop:** Drafts are opened for final review, allowing recruiters to maintain a personal touch before manual sending.

## 🏗️ Technical Architecture
The system consists of two main layers:
1.  **Frontend (Web UI):** A clean, responsive interface built with HTML5/CSS3 and JavaScript to collect email metadata and attachments.
2.  **Local Integration Layer:** [Python Flask] that interfaces with the **Outlook Object Model**.

## 🛠️ Technologies Used
* **Frontend:** JavaScript (ES6+), HTML5, CSS3.
* **Backend/Bridge:** [Python with `pywin32`].
* **Communication:** RESTful API / Localhost Socket.

## 📋 How It Works (The Logic)
1.  **Input:** User fills the subject, body, and recipients, and selects a file.
2.  **Transmission:** The Web App sends the payload to the local bridge via an asynchronous request.
3.  **Automation:** The bridge triggers the Outlook application instance, creates a `MailItem`, populates the fields, attaches the file, and calls the `.Display()` method instead of `.Send()`.

---

> **Developer's Note:** This project demonstrates the ability to solve "Real World" constraints where web-based software must interact with legacy desktop environments securely and efficiently.

# 📧🚀 Outlook Draft Automation Bridge

A professional **Full-Stack solution** designed to bridge the gap between web applications and the local desktop environment. This project enables users to generate multiple Outlook email drafts, including attachments, directly from a web interface.

---

## 📝 The Technical Challenge
Modern web browsers operate in a **Sandbox environment**, which prevents direct interaction with local software like Microsoft Outlook or access to the local file system for security reasons. Standard solutions like `mailto:` links do not support high-quality file attachments or the automated creation of multiple separate drafts.

## 💡 The Solution: Local Agent Architecture
This project implements a **Local Agent architecture** to bypass browser limitations:
* **Frontend:** A responsive HTML5/CSS3 interface for data entry.
* **Backend (The Bridge):** A local **Flask (Python)** server that communicates with the **Outlook COM API** using the `win32com` library.
* **Focus Management:** Utilizes the **Windows API (win32gui)** to bypass "Focus Stealing Prevention" mechanisms, forcing the Outlook window to the foreground upon draft creation.

---

## 🖥️ Demo
![Project Demo](./your-animation.gif)
*Visual demonstration of the automation process from browser to Outlook.*

---

## 🛠️ Tech Stack
* **Frontend:** HTML5, CSS3 (Flexbox, Animations), JavaScript (Fetch API).
* **Backend:** Python 3.x, Flask.
* **Integration:** **PyWin32 (COM Objects)**, Windows API.

---

## 🚀 Getting Started

### Prerequisites
* **Operating System:** Windows (Required for Outlook COM integration).
* **Microsoft Outlook:** Installed and configured with an active account.
* **Python 3.x:** Installed on your local machine.

### Installation
1. **Clone the repository:**
   ```bash
   git clone [https://github.com/your-username/outlook-draft-automation.git](https://github.com/your-username/outlook-draft-automation.git)
Install required dependencies:

Bash
pip install flask flask-cors pywin32
Running the Project
Start the local backend server:

Bash
python app.py
Open index.html in your web browser.

Fill in the details, attach a file, and click "Create Outlook Drafts".

🧠 Key Features & Logic
Multi-Recipient Support: Automatically splits comma-separated email lists into individual, personalized drafts.

Attachment Handling: Temporary storage logic that uploads the file, attaches it to the Outlook object, and performs a Server Cleanup after the process is complete.

Smart Automation: The system detects if Outlook is running; if not, it automatically initializes the application before generating drafts.

Enhanced UI/UX: Clean design with real-time visual feedback and loading states during request processing.

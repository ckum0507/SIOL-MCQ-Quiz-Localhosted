# SIOL Quiz Platform

A web-based quiz management system built with Flask, designed for conducting live, time-bound quizzes over a local network with complete admin control and structured result evaluation.

---

## Features

### Admin Features
- Secure admin login
- Upload quiz rules (`.docx`)
- Upload quiz questions (`.xlsx`)
- Automatic question and answer validation
- Quiz preview before hosting
- Assign test name, quiz duration, and theme
- Host / unhost quiz dynamically
- Generate QR code for LAN access
- Export quiz data as a ZIP file
- Clear temporary cache safely

---

### Participant Features
- Participant details entry
- Rules display with original document formatting
- Time-bound quiz with auto-submit
- Optional question answering
- Responsive design for desktop and mobile
- Confirmation page after submission

---

## Quiz Configuration

- Default quiz time = number of questions (in minutes)
- Max quiz time limited to 180 minutes
- Timer displayed in MM:SS format
- Auto submission when timer expires

---

## Question File Format (`questions.xlsx`)

| Column | Description |
|------|------------|
| A | Question text |
| B–F | Options A–E |
| G | Correct option letter (A–E) |
| ~ | Use `~` row to mark end of questions |

---

## Result Export Format

The exported Excel file contains:
1. Timestamp
2. Team Leader Name
3. Team Name (optional)
4. Standard
5. Phone Number
6. School Name
7. School Address
8. Total Score / Total Questions
9. Time Taken / Time Given  
10. One column per question with selected answer

Correct answers are included in headers for manual verification.

---

## File Naming Convention

All files are saved using the following format:

<TestName><AdminUser><Timestamp><Port><FileType>

makefile
Copy code

Example:
SIOL_ROUND_1_admin_20260110_183022_5000_questions.xlsx

yaml
Copy code

---

## Supported Themes

- Default (Dark Purple)
- Light
- Dark
- Dark Blue
- Solar
- Forest
- Sunset

Themes selected during preview apply to all participant pages.

---

## Hosting the Application

### Run Locally
```bash
python app.py
Custom Host and Port
bash
Copy code
python app.py --host 0.0.0.0 --port 3000
The application will be accessible via:

http://127.0.0.1:<port>

http://<LAN-IP>:<port>

QR Code Access
Admins can generate a QR code pointing to the local network IP, allowing participants to join the quiz instantly from their devices.

Folder Structure
arduino
Copy code
SIOL/
├── app.py
├── config.py
├── uploads/
│   ├── *_rules.docx
│   └── *_questions.xlsx
├── exports/
│   └── *_results.xlsx
├── temp/
├── static/
│   ├── css/
│   ├── js/
│   └── images/
├── templates/
│   ├── admin_*.html
│   └── user_*.html
└── README.md
Technology Stack
Backend: Python, Flask

Frontend: HTML, CSS, JavaScript

Session Management: Flask-Session

File Processing: OpenPyXL, python-docx

QR Code: qrcode

Styling: Custom CSS with theme variables

Intended Use
This platform is ideal for:

School or college quiz events

LAN-based competitions

Controlled quiz environments

Offline or low-internet scenarios

License
This project is intended for internal and educational use.
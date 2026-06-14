# Flex Attendance System (QR Based) 📱

A real-time QR code based attendance system built with Python and Flask — no roll calls, no paper sheets. Students just scan a QR code and mark themselves present from any device.

---

## 📌 Project Overview

This system was built to solve a real problem — taking attendance manually in class is slow and wastes time. With this system, the teacher opens a dashboard, a QR code is generated automatically, students scan it with their phones, enter their roll number, and attendance is marked instantly.

At the end of class, the teacher clicks one button and everything is saved to an Excel sheet automatically.

---

## ✨ Features

- 📲 **QR Code Generation** — Dynamic QR code generated and displayed on teacher's dashboard
- 🎓 **Student Interface** — Clean mobile-friendly page where students enter their roll number
- 📊 **Live Teacher Dashboard** — Real-time stats showing Present, Absent, and Total students
- ✅ **Roll Number Validation** — Only registered students can mark attendance
- 💾 **Auto Excel Save** — One click saves attendance to Excel with today's date as column header
- 🌐 **ngrok Integration** — Works across devices on different networks
- 🔄 **Auto Refresh** — Dashboard updates every 5 seconds automatically
- 📅 **Date Tracking** — Each session's attendance saved as a new column in the Excel sheet

---

## ⚙️ Technologies Used

- **Python** — Core language
- **Flask** — Web framework for backend and routing
- **qrcode** — QR code generation
- **openpyxl** — Excel read/write for attendance records
- **ngrok** — Tunneling for cross-device access
- **HTML/CSS** — Frontend with glassmorphism design

---

## 🛠️ How to Run

1. Clone the repository
```bash
git clone https://github.com/kainat219/flex-attendance-system.git
```

2. Install required libraries
```bash
pip install flask qrcode openpyxl
```

3. Add your student data in `AttendanceSheet.xlsx` — roll numbers in column B

4. Start ngrok on port 5000
```bash
ngrok http 5000
```

5. Run the system
```bash
python attendance.py
```

6. Open `http://localhost:5000` on teacher's device — QR code will appear automatically!

---

## 📂 Project Structure

```
├── attendance.py              # Main Flask app — all routes and logic
├── AttendanceSheet.xlsx       # Excel file with student roll numbers
├── qr_temp.png                # Auto-generated QR code image
└── .gitignore
```

---

## 🖥️ How It Works

```
Teacher opens dashboard
        ↓
QR Code is generated with today's session link
        ↓
Students scan QR → Enter roll number → Attendance marked ✅
        ↓
Dashboard updates live (every 5 seconds)
        ↓
Teacher clicks "Finalize" → Saved to Excel automatically 💾
```

---

## 🧠 What I Learned

- Building a full web application with Flask from scratch
- Generating and serving QR codes dynamically
- Reading and writing Excel files using openpyxl
- Using ngrok to expose a local server to other devices
- Designing a clean, responsive UI with glassmorphism styling
- Handling real-time data updates with auto-refresh

---

## 👩‍💻 Project Type

- **Course:** Software Requirement Engineering 
- **Type:** Semester Project
- **Language:** Python
- **Framework:** Flask

---

## 📎 Note

This system was built and tested at FAST-NUCES Lahore. It works on any device that can scan a QR code — phones, tablets, laptops. No app installation needed for students.

---

## Author

**Kainat Afzal**
BS Software Engineering — FAST-NUCES Lahore
[LinkedIn](https://linkedin.com/in/kainat-afzal) | [GitHub](https://github.com/kainat219)

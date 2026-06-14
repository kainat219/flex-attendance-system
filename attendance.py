from flask import Flask, request, render_template_string
import qrcode, openpyxl, socket, urllib.request, json
from datetime import date

app = Flask(__name__)

EXCEL_FILE = "AttendanceSheet (8).xlsx"
TODAY = date.today().strftime("%d/%m/%Y")
PORT = 5000

present_students = set()
public_url = ""

def get_ngrok_url():
    global public_url
    try:
        res = urllib.request.urlopen("http://localhost:4040/api/tunnels")
        data = json.loads(res.read())
        public_url = data['tunnels'][0]['public_url']
    except:
        public_url = f"http://{socket.gethostbyname(socket.gethostname())}:{PORT}"

STUDENT_PAGE = """
<!DOCTYPE html>
<html>
<head>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Mark Attendance</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            min-height: 100vh;
            background: linear-gradient(135deg, #1a1a2e, #16213e, #0f3460);
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }
        .card {
            background: rgba(255,255,255,0.05);
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 24px;
            padding: 40px 30px;
            width: 100%;
            max-width: 400px;
            text-align: center;
            box-shadow: 0 25px 50px rgba(0,0,0,0.4);
        }
        .icon { font-size: 60px; margin-bottom: 15px; }
        h2 { color: #ffffff; font-size: 26px; margin-bottom: 8px; }
        p { color: #a0aec0; margin-bottom: 25px; font-size: 15px; }
        input {
            width: 100%;
            padding: 16px 20px;
            font-size: 18px;
            border-radius: 12px;
            border: 2px solid rgba(255,255,255,0.1);
            background: rgba(255,255,255,0.08);
            color: white;
            outline: none;
            margin-bottom: 15px;
            transition: border 0.3s;
        }
        input::placeholder { color: #718096; }
        input:focus { border-color: #4ade80; }
        button {
            width: 100%;
            padding: 16px;
            font-size: 18px;
            font-weight: bold;
            border: none;
            border-radius: 12px;
            background: linear-gradient(135deg, #4ade80, #22c55e);
            color: white;
            cursor: pointer;
            transition: transform 0.2s, box-shadow 0.2s;
            box-shadow: 0 4px 15px rgba(74,222,128,0.4);
        }
        button:hover { transform: translateY(-2px); box-shadow: 0 8px 25px rgba(74,222,128,0.5); }
        .success {
            background: rgba(74,222,128,0.15);
            border: 1px solid #4ade80;
            color: #4ade80;
            padding: 20px;
            border-radius: 12px;
            font-size: 20px;
            font-weight: bold;
            margin-top: 10px;
        }
        .error {
            background: rgba(239,68,68,0.15);
            border: 1px solid #ef4444;
            color: #ef4444;
            padding: 20px;
            border-radius: 12px;
            font-size: 16px;
            margin-top: 10px;
        }
        .uni { color: #4ade80; font-size: 13px; margin-top: 20px; letter-spacing: 1px; text-transform: uppercase; }
    </style>
</head>
<body>
    <div class="card">
        <div class="icon">📱</div>
        <h2>Mark Your Attendance</h2>
        <p>Enter your roll number to mark yourself present</p>
        {% if message %}
            <div class="{{ msg_class }}">{{ message }}</div>
        {% else %}
            <form method="POST" action="/attend">
                <input type="text" name="roll" placeholder="e.g. 24L-0001" required autofocus>
                <button type="submit">✅ Mark Present</button>
            </form>
        {% endif %}
        <p class="uni">FAST-NU Attendance System</p>
    </div>
</body>
</html>
"""

TEACHER_PAGE = """
<!DOCTYPE html>
<html>
<head>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Attendance Dashboard</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            background: linear-gradient(135deg, #1a1a2e, #16213e, #0f3460);
            min-height: 100vh;
            color: white;
            padding: 20px;
        }
        .header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 25px;
            padding: 20px 25px;
            background: rgba(255,255,255,0.05);
            border-radius: 16px;
            border: 1px solid rgba(255,255,255,0.1);
        }
        .header h1 { font-size: 22px; color: #4ade80; }
        .header p { color: #a0aec0; font-size: 14px; margin-top: 4px; }
        .date-badge {
            background: linear-gradient(135deg, #4ade80, #22c55e);
            color: white;
            padding: 8px 20px;
            border-radius: 20px;
            font-weight: bold;
            font-size: 15px;
        }
        .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 25px; }
        .qr-card {
            background: rgba(255,255,255,0.05);
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 20px;
            padding: 25px;
            text-align: center;
        }
        .qr-card h3 { color: #4ade80; margin-bottom: 15px; font-size: 16px; letter-spacing: 1px; text-transform: uppercase; }
        .qr-card img { width: 200px; border-radius: 12px; background: white; padding: 10px; }
        .qr-card .link {
            margin-top: 12px;
            font-size: 12px;
            color: #a0aec0;
            word-break: break-all;
        }
        .stats-card {
            background: rgba(255,255,255,0.05);
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 20px;
            padding: 25px;
            display: flex;
            flex-direction: column;
            gap: 15px;
            justify-content: center;
        }
        .stat {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 15px 20px;
            border-radius: 12px;
        }
        .stat.present { background: rgba(74,222,128,0.1); border: 1px solid rgba(74,222,128,0.3); }
        .stat.absent { background: rgba(239,68,68,0.1); border: 1px solid rgba(239,68,68,0.3); }
        .stat.total { background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.1); }
        .stat-label { font-size: 15px; color: #a0aec0; }
        .stat-number { font-size: 36px; font-weight: bold; }
        .stat.present .stat-number { color: #4ade80; }
        .stat.absent .stat-number { color: #ef4444; }
        .stat.total .stat-number { color: white; }
        .finalize-btn {
            width: 100%;
            padding: 18px;
            font-size: 18px;
            font-weight: bold;
            border: none;
            border-radius: 14px;
            background: linear-gradient(135deg, #4ade80, #22c55e);
            color: white;
            cursor: pointer;
            margin-bottom: 25px;
            box-shadow: 0 4px 20px rgba(74,222,128,0.4);
            transition: transform 0.2s;
        }
        .finalize-btn:hover { transform: translateY(-2px); }
        .table-card {
            background: rgba(255,255,255,0.05);
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 20px;
            overflow: hidden;
        }
        .table-header {
            padding: 18px 25px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
            font-size: 16px;
            color: #4ade80;
            font-weight: bold;
            letter-spacing: 1px;
        }
        table { width: 100%; border-collapse: collapse; }
        th { padding: 12px 20px; text-align: center; color: #a0aec0; font-size: 13px; text-transform: uppercase; letter-spacing: 1px; border-bottom: 1px solid rgba(255,255,255,0.05); }
        td { padding: 12px 20px; text-align: center; border-bottom: 1px solid rgba(255,255,255,0.03); font-size: 14px; }
        tr:hover td { background: rgba(255,255,255,0.03); }
        .badge-P {
            background: rgba(74,222,128,0.15);
            color: #4ade80;
            border: 1px solid rgba(74,222,128,0.4);
            padding: 4px 16px;
            border-radius: 20px;
            font-weight: bold;
            font-size: 13px;
        }
        .badge-A {
            background: rgba(239,68,68,0.15);
            color: #ef4444;
            border: 1px solid rgba(239,68,68,0.4);
            padding: 4px 16px;
            border-radius: 20px;
            font-weight: bold;
            font-size: 13px;
        }
    </style>
    <meta http-equiv="refresh" content="5">
</head>
<body>
    <div class="header">
        <div>
            <h1>🎓 FAST-NU Attendance System</h1>
            <p>Real-time attendance tracking — No roll call needed</p>
        </div>
        <div class="date-badge">📅 {{ today }}</div>
    </div>

    <div class="grid">
        <div class="qr-card">
            <h3>📲 Scan to Mark Attendance</h3>
            <img src="/qr" alt="QR Code">
            <div class="link">{{ url }}/attend</div>
        </div>
        <div class="stats-card">
            <div class="stat present">
                <span class="stat-label">✅ Present</span>
                <span class="stat-number">{{ present }}</span>
            </div>
            <div class="stat absent">
                <span class="stat-label">❌ Absent</span>
                <span class="stat-number">{{ absent }}</span>
            </div>
            <div class="stat total">
                <span class="stat-label">👥 Total Students</span>
                <span class="stat-number">{{ total }}</span>
            </div>
        </div>
    </div>

    <form method="POST" action="/finalize">
        <button class="finalize-btn" type="submit">💾 Finalize Attendance & Save to Excel</button>
    </form>

    <div class="table-card">
        <div class="table-header">📋 Student Attendance List</div>
        <table>
            <tr>
                <th>#</th>
                <th>Roll Number</th>
                <th>Status</th>
            </tr>
            {% for s in students %}
            <tr>
                <td style="color:#a0aec0">{{ loop.index }}</td>
                <td style="color:white;font-weight:500">{{ s.roll }}</td>
                <td><span class="badge-{{ s.status }}">{{ s.status }}</span></td>
            </tr>
            {% endfor %}
        </table>
    </div>
</body>
</html>
"""

def load_students():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    students = []
    for row in range(2, ws.max_row + 1):
        roll = ws.cell(row=row, column=2).value
        if roll is None:
            break
        students.append(str(roll).strip())
    return students

ALL_STUDENTS = load_students()

@app.route('/')
def dashboard():
    get_ngrok_url()
    students_status = [{'roll': r, 'status': 'P' if r in present_students else 'A'} for r in ALL_STUDENTS]
    return render_template_string(TEACHER_PAGE,
        today=TODAY, url=public_url,
        students=students_status,
        present=len(present_students),
        absent=len(ALL_STUDENTS)-len(present_students),
        total=len(ALL_STUDENTS))

@app.route('/attend', methods=['GET', 'POST'])
def attend():
    if request.method == 'GET':
        return render_template_string(STUDENT_PAGE, message=None, msg_class=None)
    roll = request.form.get('roll', '').strip()
    if roll in ALL_STUDENTS:
        present_students.add(roll)
        return render_template_string(STUDENT_PAGE,
            message=f"✅ {roll} — Attendance marked successfully!",
            msg_class="success")
    else:
        return render_template_string(STUDENT_PAGE,
            message="❌ Roll number not found. Please check and try again.",
            msg_class="error")

@app.route('/qr')
def qr_image():
    get_ngrok_url()
    img = qrcode.make(f"{public_url}/attend")
    img.save("qr_temp.png")
    return open("qr_temp.png", "rb").read(), 200, {'Content-Type': 'image/png'}

@app.route('/finalize', methods=['POST'])
def finalize():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    next_col = ws.max_column + 1
    ws.cell(row=1, column=next_col).value = TODAY
    for row in range(2, ws.max_row + 1):
        roll = ws.cell(row=row, column=2).value
        if roll is None:
            break
        ws.cell(row=row, column=next_col).value = "P" if str(roll).strip() in present_students else "A"
    wb.save(EXCEL_FILE)
    return """
    <div style='font-family:Segoe UI;min-height:100vh;background:linear-gradient(135deg,#1a1a2e,#16213e,#0f3460);
    display:flex;align-items:center;justify-content:center;'>
    <div style='text-align:center;background:rgba(255,255,255,0.05);border:1px solid rgba(74,222,128,0.3);
    border-radius:24px;padding:50px;'>
    <div style='font-size:70px'>✅</div>
    <h2 style='color:#4ade80;font-size:28px;margin:20px 0 10px'>Excel Sheet Updated!</h2>
    <p style='color:#a0aec0'>Attendance saved successfully for """ + TODAY + """</p>
    </div></div>"""

if __name__ == '__main__':
    print("=" * 45)
    print("   FAST-NU ATTENDANCE SYSTEM")
    print(f"   Date     : {TODAY}")
    print(f"   Students : {len(ALL_STUDENTS)} loaded")
    print(f"   Dashboard: http://localhost:{PORT}")
    print("=" * 45)
    app.run(host='0.0.0.0', port=PORT, debug=False)
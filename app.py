import os
import json
import cloudinary
import cloudinary.uploader

import psycopg2
from psycopg2.extras import RealDictCursor

from flask import Flask, render_template, request, jsonify, send_file

from openpyxl import Workbook
from reportlab.pdfgen import canvas

app = Flask(__name__, template_folder='templates')
import os

cloudinary.config(
    cloud_name=os.getenv("CLOUDINARY_CLOUD_NAME", "elp5alks"),
    api_key=os.getenv("CLOUDINARY_API_KEY", "255227795427121"),
    api_secret=os.getenv("CLOUDINARY_API_SECRET", "58LQwglKjQ3muJwMUJCPfN5wx7s"),
    secure=True
)
DATABASE_URL = os.getenv(
    "DATABASE_URL",
    "postgresql://postgres.txffcbpthqdfuvccbqhd:Mskbtech477%40@aws-1-ap-south-1.pooler.supabase.com:6543/postgres"
)

def get_connection():
    return psycopg2.connect(DATABASE_URL)

# SEQUENCED CORRECTIONS: Ms. Seema Jabbar is now sorted directly beneath the Floor Incharge position
BASE_ROSTER = [
  {"id":"FAC1","name":"Mr. Mujtaba Khan","grade":"Faculty & Staff","section":"Managing Director","house":"General"},
  {"id":"FAC2","name":"Ms. Veena","grade":"Faculty & Staff","section":"Academic Director","house":"General"},
  {"id":"FAC3","name":"Ms. Syeda Mahjabeen","grade":"Faculty & Staff","section":"Academic Coordinator","house":"General"},
  {"id":"FAC32","name":"Mrs. Ayesha Khan","grade":"Faculty & Staff","section":"Head Mistress (SHNR)","house":"General"},
  {"id":"FAC4","name":"Ms. Asma Sultana","grade":"Faculty & Staff","section":"Academic Incharge(FLK)","house":"General"},
  {"id":"FAC5","name":"Mr. Wahed","grade":"Faculty & Staff","section":"Deeniyath HOD","house":"General"},
  {"id":"FAC6","name":"Ms. Doris Shaik","grade":"Faculty & Staff","section":"Admin Incharge","house":"General"},
  {"id":"FAC7","name":"Ms. Mahjabeen","grade":"Faculty & Staff","section":"Floor In Charge","house":"General"},
  {"id":"FAC8","name":"Ms. Seema Jabbar","grade":"Faculty & Staff","section":"In Charge","house":"General"},
  {"id":"FAC33","name":"Mr. Azmath","grade":"Faculty & Staff","section":"Supporting Coordinator","house":"General"},
  {"id":"FAC26","name":"Mr. Salman","grade":"Faculty & Staff","section":"Accountant","house":"General"},
  {"id":"FAC9","name":"Ms. Shazia","grade":"Faculty & Staff","section":"Teacher","house":"General"},
  {"id":"FAC10","name":"Ms. Parveen","grade":"Faculty & Staff","section":"Teacher","house":"General"},
  {"id":"FAC11","name":"Ms. Rubeena","grade":"Faculty & Staff","section":"Teacher","house":"General"},
  {"id":"FAC12","name":"Ms. Yasmeen","grade":"Faculty & Staff","section":"Teacher","house":"General"},
  {"id":"FAC13","name":"Ms. Heena","grade":"Faculty & Staff","section":"Teacher","house":"General"},
  {"id":"FAC14","name":"Ms. Mariyam","grade":"Faculty & Staff","section":"Teacher","house":"General"},
  {"id":"FAC15","name":"Ms. Asfiya","grade":"Faculty & Staff","section":"Teacher","house":"General"},
  {"id":"FAC16","name":"Ms. Kahera","grade":"Faculty & Staff","section":"Teacher","house":"General"},
  {"id":"FAC17","name":"Ms. Zoya","grade":"Faculty & Staff","section":"Teacher","house":"General"},
  {"id":"FAC18","name":"Ms. Reshma","grade":"Faculty & Staff","section":"Teacher","house":"General"},
  {"id":"FAC19","name":"Ms. Masarath","grade":"Faculty & Staff","section":"Teacher","house":"General"},
  {"id":"FAC20","name":"Ms. Amena","grade":"Faculty & Staff","section":"Teacher","house":"General"},
  {"id":"FAC21","name":"Ms. Maheen","grade":"Faculty & Staff","section":"Teacher","house":"General"},
  {"id":"FAC22","name":"Mr. Jaleel","grade":"Faculty & Staff","section":"Teacher","house":"General"},
  {"id":"FAC23","name":"Mr. Fuzail","grade":"Faculty & Staff","section":"Teacher","house":"General"},
  {"id":"FAC24","name":"Mr. Loqman","grade":"Faculty & Staff","section":"Teacher","house":"General"},
  {"id":"FAC25","name":"Mr. Mohiuddin","grade":"Faculty & Staff","section":"Teacher","house":"General"},
  {"id":"FAC32","name":"Imran","grade":"Faculty & Staff","section":"Driver","house":"General"},
  {"id":"FAC33","name":"Hameed","grade":"Faculty & Staff","section":"Driver","house":"General"},
  {"id":"FAC27","name":"Khaled","grade":"Faculty & Staff","section":"Watch man","house":"General"},
  {"id":"FAC28","name":"M. Jyoti","grade":"Faculty & Staff","section":"Supporting Staff","house":"General"},
  {"id":"FAC29","name":"Savitri","grade":"Faculty & Staff","section":"Supporting Staff","house":"General"},
  {"id":"FAC30","name":"Noor","grade":"Faculty & Staff","section":"Supporting Staff","house":"General"},
  {"id":"FAC31","name":"Baleshwari","grade":"Faculty & Staff","section":"Supporting Staff","house":"General"},
  {"id":"S1","name":"Abdul Kabir","house":"Superiors","grade":"4","section":""},
  {"id":"C1","name":"Ali Bin Sawood Bakoban","house":"Champions","grade":"4","section":""},
  {"id":"S2","name":"Ayesha Syed Hasan","house":"Superiors","grade":"4","section":""},
  {"id":"C2","name":"Ayra Afsheen","house":"Champions","grade":"4","section":""},
  {"id":"S3","name":"BIBI AYESHA SIDDIQUA","house":"Superiors","grade":"4","section":""},
  {"id":"S4","name":"Habeeb Khan","house":"Superiors","grade":"4","section":""},
  {"id":"C3","name":"MAHER UNNISA","house":"Champions","grade":"4","section":""},
  {"id":"C4","name":"Maira Mujeeb Khan","house":"Champions","grade":"4","section":""},
  {"id":"C5","name":"Maryam Ismail Shareef","house":"Champions","grade":"4","section":""},
  {"id":"W1","name":"Mir Anas Ali","house":"Warriors","grade":"4","section":""},
  {"id":"C6","name":"Mohammed Ibrahim","house":"Superiors","grade":"4","section":""},
  {"id":"C7","name":"MOHAMMED MALIK KHAN","house":"Champions","grade":"4","section":""},
  {"id":"S5","name":"Mohammed MizbaUddin","house":"Superiors","grade":"4","section":""},
  {"id":"CH1","name":"Mohammed Sohaan","house":"Challengers","grade":"4","section":""},
  {"id":"C8","name":"Muhammad Ibrahim","house":"Champions","grade":"4","section":""},
  {"id":"W2","name":"Nida Fatima","house":"Warriors","grade":"4","section":""},
  {"id":"CH2","name":"Shaik Azaan","house":"Superiors","grade":"4","section":""},
  {"id":"CH3","name":"Syed Azaan Uddin","house":"Challengers","grade":"4","section":""},
  {"id":"C9","name":"SYED FARHAN","house":"Champions","grade":"4","section":""},
  {"id":"S6","name":"Syed Mohammed Hussain","house":"Warriors","grade":"4","section":""},
  {"id":"C10","name":"Syed Murtaza","house":"Champions","grade":"4","section":""},
  {"id":"S7","name":"Syed Mustafa","house":"Superiors","grade":"4","section":""},
  {"id":"S8","name":"Syeda Hoorain","house":"Superiors","grade":"4","section":""},
  {"id":"C11","name":"Syeda Zainab Afsheen","house":"Champions","grade":"4","section":""},
  {"id":"W3","name":"Tazkiya Tahreem","house":"Warriors","grade":"4","section":""},
  {"id":"S9","name":"Zahra Fatima","house":"Superiors","grade":"4","section":""},
  {"id":"CH4","name":"Arhum Khan","house":"Challengers","grade":"5","section":""},
  {"id":"CH5","name":"Batool Fatima","house":"Challengers","grade":"5","section":""},
  {"id":"W4","name":"DURAR HADI AL HASSANI","house":"Warriors","grade":"5","section":""},
  {"id":"S10","name":"Khadeejah Fatima","house":"Champions","grade":"5","section":""},
  {"id":"S11","name":"KHADIJA MAIRA","house":"Superiors","grade":"5","section":""},
  {"id":"CH6","name":"Mariyam Fatima","house":"Champions","grade":"5","section":""},
  {"id":"C12","name":"Mohammed Abu Bakr","house":"Champions","grade":"5","section":""},
  {"id":"W5","name":"Mohammed Ammar Hussain","house":"Warriors","grade":"5","section":""},
  {"id":"W6","name":"Mohammed Arzan","house":"Warriors","grade":"5","section":""},
  {"id":"S12","name":"Mohammed Faiz","house":"Superiors","grade":"5","section":""},
  {"id":"CH7","name":"MOHAMMED MUHAMMED UDDIN","house":"Challengers","grade":"5","section":""},
  {"id":"C13","name":"Mohammed Zain Mahmood","house":"Champions","grade":"5","section":""},
  {"id":"W7","name":"Rashed Ismail Shareef","house":"Warriors","grade":"5","section":""},
  {"id":"C14","name":"Syed Affan Meer","house":"Champions","grade":"5","section":""},
  {"id":"CH8","name":"Syed Ayaan Ahmed","house":"Challengers","grade":"5","section":""},
  {"id":"W8","name":"Syeda Samaira Noor","house":"Warriors","grade":"5","section":""},
  {"id":"CH9","name":"Umaima Shaik","house":"Challengers","grade":"5","section":""},
  {"id":"W9","name":"Zunaira Hyder","house":"Warriors","grade":"5","section":""},
  {"id":"CH10","name":"Abdul Kabeer","house":"Challengers","grade":"6","section":""},
  {"id":"CH11","name":"Abdul Rehman Syed Hasan","house":"Superiors","grade":"6","section":""},
  {"id":"S13","name":"Afeefah Fathima","house":"Superiors","grade":"6","section":""},
  {"id":"W10","name":"Aiza Sadiq","house":"Warriors","grade":"6","section":""},
  {"id":"CH12","name":"Aleena Azin","house":"Challengers","grade":"6","section":""},
  {"id":"W11","name":"Aliza Fatima","house":"Warriors","grade":"6","section":""},
  {"id":"W12","name":"Daniya Laraib","house":"Warriors","grade":"6","section":""},
  {"id":"W13","name":"Juveriya Kamal","house":"Warriors","grade":"6","section":""},
  {"id":"CH13","name":"Madiha Bakooban","house":"Challengers","grade":"6","section":""},
  {"id":"S14","name":"Mohammed Abdul Rayyan","house":"Champions","grade":"6","section":""},
  {"id":"S15","name":"Mohammed Arhaan","house":"Superiors","grade":"6","section":""},
  {"id":"S16","name":"Shaik Ibrahim Ifran","house":"Superiors","grade":"6","section":""},
  {"id":"S17","name":"Syed Muhammad Ahmad Ali Abedi","house":"Champions","grade":"6","section":""},
  {"id":"C15","name":"Syed Uzair","house":"Challengers","grade":"6","section":""},
  {"id":"C16","name":"Syed Zuhair Abbas","house":"Champions","grade":"6","section":""},
  {"id":"C17","name":"Tameem Raahil Bhat","house":"Champions","grade":"6","section":""},
  {"id":"CH14","name":"Umra Fatima","house":"Challengers","grade":"6","section":""},
  {"id":"C18","name":"Uzair Khan","house":"Champions","grade":"6","section":""},
  {"id":"W14","name":"Zara Saleem","house":"Warriors","grade":"6","section":""},
  {"id":"W15","name":"Alina Hassan","house":"Warriors","grade":"7","section":""},
  {"id":"S18","name":"Arbiya Fatima","house":"Superiors","grade":"7","section":""},
  {"id":"S19","name":"Juveriya Fatima","house":"Superiors","grade":"7","section":""},
  {"id":"W16","name":"Mehreen Fatima","house":"Warriors","grade":"7","section":""},
  {"id":"S20","name":"Mirza Awad Baig","house":"Champions","grade":"7","section":""},
  {"id":"CH15","name":"Mohammed Anzar Ashraf Ali","house":"Challengers","grade":"7","section":""},
  {"id":"S21","name":"Mohammed Safwan","house":"Superiors","grade":"7","section":""},
  {"id":"C19","name":"Mohammed Waheed Uddin","house":"Champions","grade":"7","section":""},
  {"id":"C20","name":"Mohd Saddiq Qureshi","house":"Champions","grade":"7","section":""},
  {"id":"S22","name":"Mohd Yousuf Qureshi","house":"Superiors","grade":"7","section":""},
  {"id":"S23","name":"Muhammad Sanaullah","house":"Superiors","grade":"7","section":""},
  {"id":"W17","name":"Munazza zareen","house":"Warriors","grade":"7","section":""},
  {"id":"CH16","name":"Rida Fatima","house":"Challengers","grade":"7","section":""},
  {"id":"S24","name":"Shaik Abdul Rabah Ur Rahman Quadri","house":"Superiors","grade":"7","section":""},
  {"id":"S25","name":"Shaik Nadeem Hussain","house":"Superiors","grade":"7","section":""},
  {"id":"S26","name":"Syed Abdul Aziz Ahmed","house":"Superiors","grade":"7","section":""},
  {"id":"W18","name":"Syed Anas Meer","house":"Warriors","grade":"7","section":""},
  {"id":"CH17","name":"Syed Furqan Ali Abedi","house":"Challengers","grade":"7","section":""},
  {"id":"W19","name":"Syeda Mahreen","house":"Warriors","grade":"7","section":""},
  {"id":"W20","name":"Syeda Noor ul Ameen","house":"Warriors","grade":"7","section":""},
  {"id":"W21","name":"Syeda Sidra Fatima","house":"Warriors","grade":"7","section":""},
  {"id":"CH18","name":"Taha Syed Najeeb Uddin","house":"Challengers","grade":"7","section":""},
  {"id":"W22","name":"Ummul Hamda","house":"Warriors","grade":"7","section":""},
  {"id":"C21","name":"Abdullah Syed Hasan","house":"Champions","grade":"8","section":""},
  {"id":"W23","name":"Ahmed Bin Sawood Bakooban","house":"Warriors","grade":"8","section":""},
  {"id":"C22","name":"Amatul Hameed Rumaisa","house":"Challengers","grade":"8","section":""},
  {"id":"CH19","name":"Ammara Faiz","house":"Challengers","grade":"8","section":""},
  {"id":"W24","name":"Md Abdul Bari","house":"Challengers","grade":"8","section":""},
  {"id":"W25","name":"Mohammed Abdul Salah","house":"Warriors","grade":"8","section":""},
  {"id":"CH20","name":"Mohammed Arhaan Uddin","house":"Challengers","grade":"8","section":""},
  {"id":"CH21","name":"Mohammed Rehan","house":"Challengers","grade":"8","section":""},
  {"id":"W26","name":"Mohammed Shahriyar Khan","house":"Warriors","grade":"8","section":""},
  {"id":"W27","name":"Mohd Arham Uddin","house":"Warriors","grade":"8","section":""},
  {"id":"W28","name":"Syed Khaja Ahsanuddin","house":"Warriors","grade":"8","section":""},
  {"id":"CH22","name":"Syed Rehan","house":"Challengers","grade":"8","section":""},
  {"id":"CH23","name":"Syed Yaheya Zain","house":"Challengers","grade":"8","section":""},
  {"id":"CH24","name":"Syeda Mahveen Fatima","house":"Challengers","grade":"8","section":""},
  {"id":"CH25","name":"Syeda Zunairah Begum","house":"Challengers","grade":"8","section":""},
  {"id":"C23","name":"Uzair Imtiaz","house":"Champions","grade":"8","section":""},
  {"id":"C24","name":"Zunairah Amtul Aleem","house":"Champions","grade":"8","section":""},

]
def get_weight_limit(voter_id):
    profile = next((x for x in BASE_ROSTER if x["id"] == voter_id), None)

    if not profile:
        return 1

    if profile["grade"] != "Faculty & Staff":
        return 1

    section = profile.get("section", "").lower()
    name = profile.get("name", "").lower()

    if (
        "teacher" in section
        or "watch" in section
        or "driver" in section
        or "supporting" in section
        or "admin" in section
        or "floor" in section
        or "salman" in name
        or "seema" in name
    ):
        return 2

    return 5
def init_db():
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS votes (
        candidate_id TEXT,
        context_type TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS voter_history (
        voter_id TEXT PRIMARY KEY,
        count INTEGER
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS asset_vault (
        vault_key TEXT PRIMARY KEY,
        vault_val JSONB
    )
    """)

    conn.commit()
    cur.close()
    conn.close()

@app.route('/')
def home():
    return render_template('index.html')

@app.route("/api/upload-image", methods=["POST"])
def upload_image():
    try:
        file = request.files["image"]

        result = cloudinary.uploader.upload(
            file,
            folder="vip-election"
        )

        return jsonify({
            "success": True,
            "url": result["secure_url"]
        })

    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500

@app.route('/api/get-initial-state')
def get_initial_state():

    init_db()

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT vault_val
        FROM asset_vault
        WHERE vault_key=%s
    """, ("custom_config",))

    row = cur.fetchone()
    custom_config = row[0] if row else None

    if isinstance(custom_config, str):
        custom_config = json.loads(custom_config)

    cur.execute("""
        SELECT candidate_id, COUNT(*)
        FROM votes
        GROUP BY candidate_id
    """)

    votes_tally = {
        r[0]: r[1]
        for r in cur.fetchall()
    }

    vote_percentages = {}

    if custom_config:
        for position in custom_config.get("positions", []):

            total_votes = sum(
                votes_tally.get(candidate["id"], 0)
                for candidate in position["candidates"]
            )

            for candidate in position["candidates"]:
                votes = votes_tally.get(candidate["id"], 0)

                vote_percentages[candidate["id"]] = round(
                    (votes / total_votes) * 100, 2
                ) if total_votes else 0

    cur.execute("""
        SELECT voter_id, count
        FROM voter_history
    """)

    history = {
        r[0]: r[1]
        for r in cur.fetchall()
    }

    cur.close()
    conn.close()

    return jsonify({
        "roster": BASE_ROSTER,
        "votes": votes_tally,
        "percentages": vote_percentages,
        "voters": history,
        "customConfig": custom_config
    })

@app.route('/api/save-config', methods=['POST'])
def save_config():

    data = request.get_json()

    if not data:
        return jsonify({
            "success": False,
             "error": "No JSON received"
         }), 400

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO asset_vault (vault_key, vault_val)
        VALUES (%s, %s)
        ON CONFLICT (vault_key)
        DO UPDATE SET vault_val = EXCLUDED.vault_val
    """, (
        "custom_config",
        json.dumps(data)
    ))

    conn.commit()

    cur.close()
    conn.close()

    return jsonify({
        "success": True
    })

@app.route('/api/cast-vote', methods=['POST'])
def cast_vote():
    try:
        data = request.get_json()

        if not data:
            return jsonify({
                "success": False,
                "error": "No JSON received"
            }), 400

        voter_id = data.get("voterId")

        if not voter_id:
            return jsonify({
                "success": False,
                "error": "Invalid voter."
            }), 400

        selections = data.get("selections", [])

        limit = get_weight_limit(voter_id)

        conn = get_connection()
        cur = conn.cursor()

        cur.execute(
            "SELECT count FROM voter_history WHERE voter_id=%s",
            (voter_id,)
        )

        row = cur.fetchone()
        current_count = row[0] if row else 0

        if current_count >= limit:
            cur.close()
            conn.close()

            return jsonify({
                "success": False,
                "error": "Allocations spent."
            }), 400

        for cand_id in selections:
            if cand_id:
                cur.execute("""
                    INSERT INTO votes(candidate_id, context_type)
                    VALUES(%s, %s)
                """, (cand_id, "general"))

        next_count = current_count + 1

        cur.execute("""
            INSERT INTO voter_history(voter_id, count)
            VALUES(%s, %s)
            ON CONFLICT(voter_id)
            DO UPDATE SET count = EXCLUDED.count
        """, (voter_id, next_count))

        conn.commit()

        cur.close()
        conn.close()

        return jsonify({
            "success": True,
            "votes_remaining": limit - next_count
        })

    except Exception as e:
        import traceback
        traceback.print_exc()

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500
    
@app.route('/api/reset', methods=['POST'])
def reset_database():

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("TRUNCATE TABLE votes RESTART IDENTITY;")
    cur.execute("TRUNCATE TABLE voter_history RESTART IDENTITY;")

    conn.commit()

    cur.close()
    conn.close()

    return jsonify({
        "success": True
    })

@app.route("/api/export/excel")
def export_excel():

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT candidate_id,
               COUNT(*)
        FROM votes
        GROUP BY candidate_id
        ORDER BY COUNT(*) DESC
    """)

    rows = cur.fetchall()

    cur.close()
    conn.close()

    wb = Workbook()

    ws = wb.active
    ws.title = "Election Results"

    ws.append(["Candidate ID", "Votes"])

    for row in rows:
        ws.append(list(row))

    filename = "Election_Results.xlsx"

    wb.save(filename)

    return send_file(filename, as_attachment=True)

@app.route("/api/export/pdf")
def export_pdf():

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT vault_val
        FROM asset_vault
        WHERE vault_key=%s
    """, ("custom_config",))

    row = cur.fetchone()

    config = row[0] if row else {}

    if isinstance(config, str):
        config = json.loads(config)

    candidate_lookup = {}

    for position in config.get("positions", []):

        for candidate in position["candidates"]:

            candidate_lookup[candidate["id"]] = {
                "name": candidate["name"],
                "position": position["title"]
            }

    cur.execute("""
        SELECT candidate_id,
               COUNT(*)
        FROM votes
        GROUP BY candidate_id
        ORDER BY COUNT(*) DESC
    """)

    rows = cur.fetchall()

    cur.close()
    conn.close()

    filename = "Election_Results.pdf"

    pdf = canvas.Canvas(filename)

    pdf.setFont("Helvetica-Bold", 18)
    pdf.drawString(170, 800, "Election Results")

    y = 760

    pdf.setFont("Helvetica", 12)

    for candidate_id, votes in rows:

        info = candidate_lookup.get(candidate_id)

        if info:
            text = f"{info['position']} - {info['name']} : {votes} Votes"
        else:
            text = f"{candidate_id} : {votes} Votes"

        pdf.drawString(40, y, text)

        y -= 22

        if y < 60:
            pdf.showPage()
            pdf.setFont("Helvetica", 12)
            y = 800

    pdf.save()

    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    init_db()
    app.run(host='0.0.0.0', port=5000, debug=True)
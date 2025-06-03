from flask import Flask, request, render_template, send_file
import os
import pdfplumber
import pandas as pd
import re
from datetime import datetime

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# 和暦の日付を西暦に変換
def convert_to_full_date(date_str):
    try:
        m, d = map(int, re.findall(r'(\d+)月(\d+)', date_str)[0])
        return f"2025/{m:02}/{d:02}"
    except:
        return None

# 性別を推定
def infer_gender(event_name):
    if '男子' in event_name:
        return '男子'
    elif '女子' in event_name:
        return '女子'
    else:
        return None

# 学年取得用関数（辞書を外から渡す）
def get_grade(name, name_to_grade):
    return name_to_grade.get(name, None)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        pdf_file = request.files["pdf"]
        excel_file = request.files["excel"]
        team_name = request.form["team"]

        if not pdf_file or not excel_file or not team_name:
            return "すべての入力が必要です。"

        pdf_path = os.path.join(UPLOAD_FOLDER, "input.pdf")
        excel_path = os.path.join(UPLOAD_FOLDER, "name_to_grade.xlsx")
        output_txt = os.path.join(UPLOAD_FOLDER, "output.txt")

        pdf_file.save(pdf_path)
        excel_file.save(excel_path)

        # === 元の処理開始 ===
        with pdfplumber.open(pdf_path) as pdf, open(output_txt, "w", encoding="utf-8") as f_out:
            for page in pdf.pages:
                lines = page.extract_text(layout=True).split("\n")
                for line in lines:
                    f_out.write(line + "\n")

        with open(output_txt, "r", encoding="utf-8") as f:
            lines = f.read().strip().split("\n")

        joyo = r"[A-Za-z０-９Ａ-Ｚａ-ｚ0-9\u4E00-\u9FFF\u3040-\u309F\u30A0-\u30FF]+"
        name_sep = r"(?:\s|\u3000)"
        athlete_pattern = fr"(?:\(?\d*\)?)?\s*({joyo}{name_sep}{joyo})\(?\d*\)?\s+({joyo}(?:\s+{joyo})?)\s+(DNS|DNF|\d+(?::\d+)?\.\d+)"
        event_pattern = r"([Ａ-Ｚａ-ｚ0-９０-９一-龯々〆〤ぁ-んァ-ヶー]+(?:\s?[Ａ-Ｚａ-ｚ0-９０-９一-龯々〆〤ぁ-んァ-ヶー]+)*\s?\d+m(?:H|SC)?)"
        results_pattern_1 = fr"(?:(\d+)\s+)?(\d+)\s+(\d+)\s+{athlete_pattern}"
        results_pattern_2 = fr"(?:(\d+)\s+)?(\d+)\s+(\d+)\s+{athlete_pattern}\s+(?:(\d+)\s+)?(\d+)\s+(\d+)\s+{athlete_pattern}"

        results = []
        current_race_info = []
        current_date_str = None
        current_event = None

        grade_df = pd.read_excel(excel_path)
        name_to_grade = dict(zip(grade_df["氏名"], grade_df["学年"]))

        for line in lines:
            line = line.strip()
            if re.match(r"\d+月\d+日", line):
                current_date_str = convert_to_full_date(line)
                continue

            match_event = re.match(event_pattern, line)
            if match_event:
                current_event = match_event.group(1).strip()
                gender = infer_gender(current_event)
                current_race_info = [(1,None)]
                continue

            match_race = re.findall(r"(\d+)組\s*(?:\(風:([+-]?[0-9.]+)\))?", line)
            if match_race:
                current_race_info = [(int(g), float(w) if w else None) for g, w in match_race]

            match2 = re.match(results_pattern_2, line)
            if match2:
                values = match2.groups()
                # 左
                results.append({
                    "種目": current_event,
                    "組": current_race_info[0][0],
                    "風": current_race_info[0][1],
                    "順位": int(values[0]) if values[0] else None,
                    "レーン": int(values[1]),
                    "ナンバー": int(values[2]),
                    "氏名": values[3].strip(),
                    "所属": values[4].strip(),
                    "記録": values[5],
                    "性別": gender,
                    "日付": current_date_str,
                    "学年": get_grade(values[3].strip(), name_to_grade),
                })
                # 右
                if len(current_race_info) > 1:
                    results.append({
                        "種目": current_event,
                        "組": current_race_info[1][0],
                        "風": current_race_info[1][1],
                        "順位": int(values[6]) if values[6] else None,
                        "レーン": int(values[7]),
                        "ナンバー": int(values[8]),
                        "氏名": values[9].strip(),
                        "所属": values[10].strip(),
                        "記録": values[11],
                        "性別": gender,
                        "日付": current_date_str,
                        "学年": get_grade(values[9].strip(), name_to_grade),
                    })
                continue

            match1 = re.match(results_pattern_1, line)
            if match1 and len(current_race_info) >= 1:
                values = match1.groups()
                results.append({
                    "種目": current_event,
                    "組": current_race_info[0][0],
                    "風": current_race_info[0][1],
                    "順位": int(values[0]) if values[0] else None,
                    "レーン": int(values[1]),
                    "ナンバー": int(values[2]),
                    "氏名": values[3].strip(),
                    "所属": values[4].strip(),
                    "記録": values[5],
                    "性別": gender,
                    "日付": current_date_str,
                    "学年": get_grade(values[3].strip(), name_to_grade),
                })

        df = pd.DataFrame(results)
        df = df[df["記録"].str.match(r"^\d+(?::\d+)?\.\d+$")]
        df = (
            df.sort_values(by=["種目", "記録"])
              .groupby("種目", group_keys=False)
              .apply(lambda d: d.assign(全体順位=range(1, len(d)+1)))
              .reset_index(drop=True)
        )

        newdf = df[df['所属'] == team_name].sort_values(by=["種目", "記録"]).reset_index(drop=True)
        newdf.index = newdf.index + 1

        newdf = newdf[[
            "性別", "種目", "氏名", "学年", "記録", "風", "日付",
            "組", "レーン", "順位", "全体順位"
        ]]
        newdf["種目"] = newdf["種目"].str.replace(r"^(一般)?(男|女)子", "", regex=True)

        output_name = request.form["output_name"]
        output_excel_path = os.path.join(UPLOAD_FOLDER, f"{output_name}.xlsx")
        newdf.to_excel(output_excel_path, index=False)

        return send_file(output_excel_path, as_attachment=True)
        
    return render_template("index.html")

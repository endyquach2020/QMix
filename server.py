import os
import json
import shutil
import zipfile
import uuid
import re
from typing import List
from copy import deepcopy
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from docx.oxml.ns import qn

from parser import parse_docx
from shuffler import shuffle_and_export, export_answers_excel

app = FastAPI(title="QMix - Exam Shuffler Studio", version="1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["X-Filename"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "temp_uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "temp_outputs")
TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "index.html")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.get("/", response_class=HTMLResponse)
async def serve_home_page():
    if not os.path.exists(TEMPLATE_PATH):
        return HTMLResponse(
            content=f"<h3>Không tìm thấy file: {TEMPLATE_PATH}</h3>",
            status_code=404
        )
    with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
        html_content = f.read()
    return HTMLResponse(content=html_content)


def check_is_correct(choice_obj):
    if choice_obj.get("is_correct"):
        return True

    xml_node = choice_obj.get("xml")
    if xml_node is not None:
        for u in xml_node.findall(f".//{qn('w:u')}"):
            val = u.get(qn('w:val'))
            if val is None or str(val).lower() not in ['none', '0', 'false']:
                return True
        for c in xml_node.findall(f".//{qn('w:color')}"):
            val = str(c.get(qn('w:val')) or '').upper()
            if any(red in val for red in ["FF0000", "RED", "C00000", "ED1C24", "FE0000"]):
                return True
        for h in xml_node.findall(f".//{qn('w:highlight')}"):
            val = str(h.get(qn('w:val')) or '').lower()
            if val not in ['none', '0']:
                return True

    text_val = choice_obj.get("text", "")
    if text_val and (text_val.strip().startswith("#") or text_val.strip().startswith("*")):
        return True

    return False


@app.post("/api/upload")
async def upload_exam_file(file: UploadFile = File(...)):
    if not file.filename.endswith(".docx"):
        raise HTTPException(status_code=400, detail="Chỉ hỗ trợ tệp định dạng .docx")

    file_id = f"upload_{os.urandom(4).hex()}"
    saved_filename = f"{file_id}_{file.filename}"
    file_path = os.path.join(UPLOAD_DIR, saved_filename)

    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        groups = parse_docx(file_path)
        parsed_data = []
        total_questions = 0

        for grp_idx, grp in enumerate(groups):
            grp_info = {
                "group_index": grp_idx,
                "tag": grp.get("tag", ""),
                "title": grp.get("title", ""),
                "questions": []
            }

            for q_idx, q in enumerate(grp["questions"]):
                total_questions += 1
                
                # Ưu tiên lấy text câu hỏi đã bóc tách ảnh MathType Base64 & LaTeX
                body_text = q.get("text", "")
                if not body_text and q.get("body_elements"):
                    body_text = "".join([t.text for t in q["body_elements"][0].findall(".//{*}t") if t.text]).strip()

                choices_data = []
                for c in q.get("choices", []):
                    choices_data.append({
                        "label": c.get("label", ""),
                        "text": c.get("text", ""),
                        "is_correct": check_is_correct(c),
                    })

                grp_info["questions"].append({
                    "id": f"{grp_idx}_{q_idx}",
                    "original_num": q.get("original_num", str(total_questions)),
                    "text": body_text,
                    "choices": choices_data,
                })

            parsed_data.append(grp_info)

        return {
            "success": True,
            "file_id": file_id,
            "filename": file.filename,
            "total_groups": len(groups),
            "total_questions": total_questions,
            "data": parsed_data
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi khi đọc file Word: {str(e)}")


def extract_original_answers(groups):
    """Trích xuất đáp án đề gốc (mã 000) tách biệt theo part_type, loại bỏ nhóm q0"""
    p1_answers = []
    p2_answers = []
    p3_answers = []

    for grp in groups:
        is_group_q0 = (not grp.get("shuffle_questions", True)) and (not grp.get("shuffle_choices", True))
        if is_group_q0:
            continue

        part_type = grp.get("part_type", "part_1")

        if part_type == "part_2":
            for q in grp["questions"]:
                for sub_lbl in ['a', 'b', 'c', 'd']:
                    target = next((c for c in q["choices"] if c.get("label", "").lower().startswith(sub_lbl)), None)
                    if target and check_is_correct(target):
                        p2_answers.append("Đ")
                    else:
                        p2_answers.append("S")

        elif part_type == "part_3":
            for q in grp["questions"]:
                val = ""
                if q.get("choices"):
                    raw_val = q["choices"][0].get("text", "").strip()
                    val = re.sub(r"<[^>]+>", "", raw_val).strip()
                p3_answers.append(val)

        elif part_type == "part_1":
            for q in grp["questions"]:
                corr_label = ""
                for c in q.get("choices", []):
                    if check_is_correct(c):
                        clean_label = re.sub(r"[^A-Za-z]", "", c.get("label", "")).upper()
                        corr_label = clean_label
                        break
                p1_answers.append(corr_label)

    return p1_answers, p2_answers, p3_answers


@app.post("/api/mix")
async def mix_and_export_exams(
    file_id: str = Form(...),
    original_filename: str = Form(...),
    exam_codes_json: str = Form(...),
    exam_info_json: str = Form(None)
):
    matching_files = [f for f in os.listdir(UPLOAD_DIR) if f.startswith(file_id)]
    if not matching_files:
        raise HTTPException(status_code=404, detail="Không tìm thấy tệp đề gốc tương ứng")

    file_path = os.path.join(UPLOAD_DIR, matching_files[0])

    try:
        codes_list = json.loads(exam_codes_json)
        if not codes_list or not isinstance(codes_list, list):
            raise ValueError("Danh sách mã đề không hợp lệ")

        exam_info = {}
        if exam_info_json:
            try:
                exam_info = json.loads(exam_info_json)
            except Exception:
                pass

        groups = parse_docx(file_path)

        valid_codes = [str(c).strip() for c in codes_list if str(c).strip()]
        num_codes = len(valid_codes)

        random_suffix = uuid.uuid4().hex[:4]
        job_dir = os.path.join(OUTPUT_DIR, f"job_{random_suffix}")
        os.makedirs(job_dir, exist_ok=True)

        generated_files = []

        # 1. Đưa ĐỀ GỐC vào gói ZIP
        degoc_filename = "QMix_De_Goc.docx"
        degoc_path = os.path.join(job_dir, degoc_filename)
        shutil.copy2(file_path, degoc_path)
        generated_files.append(degoc_path)

        # 2. Khởi tạo cấu trúc bảng Excel với đáp án đề gốc 000
        orig_p1, orig_p2, orig_p3 = extract_original_answers(groups)

        excel_data = {
            "part_1": {
                "headers": [str(i) for i in range(1, len(orig_p1) + 1)] if orig_p1 else [],
                "rows": [{"code": "000", "answers": orig_p1}] if orig_p1 else []
            },
            "part_2": {
                "headers": [f"{q}{sub}" for q in range(1, (len(orig_p2) // 4) + 1) for sub in ['a', 'b', 'c', 'd']] if orig_p2 else [],
                "rows": [{"code": "000", "answers": orig_p2}] if orig_p2 else []
            },
            "part_3": {
                "headers": [str(i) for i in range(1, len(orig_p3) + 1)] if orig_p3 else [],
                "rows": [{"code": "000", "answers": orig_p3}] if orig_p3 else []
            }
        }

        all_text_keys = {}

        # 3. Trộn và thu thập đáp án các mã đề
        for code in valid_codes:
            out_filename = f"Ma_de_{code}.docx"
            out_path = os.path.join(job_dir, out_filename)

            groups_copy = deepcopy(groups)
            ans_res = shuffle_and_export(groups_copy, file_path, code, out_path, exam_info)
            
            if isinstance(ans_res, tuple):
                flat_keys, structured_ans = ans_res
            else:
                flat_keys = ans_res
                structured_ans = {"part_1": flat_keys, "part_2": {}, "part_3": {}}

            all_text_keys[code] = flat_keys
            generated_files.append(out_path)

            if excel_data["part_1"]["headers"]:
                cur_p1 = [structured_ans.get("part_1", {}).get(i, "") for i in range(1, len(excel_data["part_1"]["headers"]) + 1)]
                excel_data["part_1"]["rows"].append({"code": code, "answers": cur_p1})

            if excel_data["part_2"]["headers"]:
                cur_p2 = [structured_ans.get("part_2", {}).get(h, "") for h in excel_data["part_2"]["headers"]]
                excel_data["part_2"]["rows"].append({"code": code, "answers": cur_p2})

            if excel_data["part_3"]["headers"]:
                cur_p3 = [structured_ans.get("part_3", {}).get(i, "") for i in range(1, len(excel_data["part_3"]["headers"]) + 1)]
                excel_data["part_3"]["rows"].append({"code": code, "answers": cur_p3})

        # 4. Xuất file Bang_dap_an.xlsx
        excel_filename = "Bang_dap_an.xlsx"
        excel_path = os.path.join(job_dir, excel_filename)
        export_answers_excel(excel_data, excel_path)
        generated_files.append(excel_path)

        # 5. Xuất file Bang_dap_an.txt
        txt_filename = "Bang_dap_an.txt"
        txt_path = os.path.join(job_dir, txt_filename)
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write("BẢNG ĐÁP ÁN CÁC MÃ ĐỀ\n")
            f.write("=" * 45 + "\n\n")
            for code, keys in all_text_keys.items():
                f.write(f"MÃ ĐỀ: {code}\n")
                row = []
                for q_num, ans in sorted(keys.items()):
                    row.append(f"{q_num}:{ans}")
                    if len(row) == 10:
                        f.write("  ".join(row) + "\n")
                        row = []
                if row:
                    f.write("  ".join(row) + "\n")
                f.write("-" * 45 + "\n")
        generated_files.append(txt_path)

        # 6. Đóng gói ZIP
        zip_filename = f"QMix_bo_de_tron_{num_codes}_ma.zip"
        zip_path = os.path.join(OUTPUT_DIR, f"{uuid.uuid4().hex[:6]}_{zip_filename}")

        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for f_path in generated_files:
                zipf.write(f_path, arcname=os.path.basename(f_path))

        shutil.rmtree(job_dir, ignore_errors=True)

        return FileResponse(
            path=zip_path,
            filename=zip_filename,
            media_type="application/zip",
            headers={
                "X-Filename": zip_filename,
                "Content-Disposition": f'attachment; filename="{zip_filename}"'
            }
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi khi trộn đề: {str(e)}")
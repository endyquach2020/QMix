import docx
import re
import os
import io
import base64
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from copy import deepcopy
from PIL import Image

# Đăng ký namespace VML nếu chưa tồn tại
VML_NS = "urn:schemas-microsoft-com:vml"

def check_run_is_correct(r_elem):
    rPr = r_elem.find(qn('w:rPr'))
    if rPr is None:
        return False

    # 1. Gạch chân
    u = rPr.find(qn('w:u'))
    if u is not None:
        u_val = str(u.get(qn('w:val')) or '').lower()
        if u_val not in ['none', '0', 'false']:
            return True

    # 2. Màu đỏ
    color = rPr.find(qn('w:color'))
    if color is not None:
        c_val = str(color.get(qn('w:val')) or '').upper()
        if any(red in c_val for red in ["FF0000", "RED", "C00000", "ED1C24", "FE0000"]):
            return True

    # 3. Highlight
    hl = rPr.find(qn('w:highlight'))
    if hl is not None:
        h_val = str(hl.get(qn('w:val')) or '').lower()
        if h_val not in ['none', '0']:
            return True

    return False


def check_paragraph_has_correct_mark(p_elem):
    for r in p_elem.findall(f".//{qn('w:r')}"):
        if check_run_is_correct(r):
            return True
    return False


def omml_to_latex(node):
    """Chuyển đổi cây XML công thức toán Word Equation (OMML) sang LaTeX"""
    tag = node.tag.split('}')[-1]

    if tag == 't':
        val = node.text or ''
        val = val.replace('\u2212', '-').replace('\u2013', '-').replace('\u221e', '\\infty ')
        val = val.replace('\u03c0', '\\pi ').replace('\u03b1', '\\alpha ').replace('\u03b2', '\\beta ')
        val = val.replace('\u2264', '\\le ').replace('\u2265', '\\ge ').replace('\u2260', '\\ne ')
        val = val.replace('\u2208', '\\in ').replace('\u2209', '\\notin ')
        return val

    if tag == 'r':
        return "".join(omml_to_latex(c) for c in node)

    if tag == 'f':
        num = node.find(qn('m:num'))
        den = node.find(qn('m:den'))
        num_s = omml_to_latex(num) if num is not None else ""
        den_s = omml_to_latex(den) if den is not None else ""
        return f"\\frac{{{num_s}}}{{{den_s}}}"

    if tag == 'sSup':
        e = node.find(qn('m:e'))
        sup = node.find(qn('m:sup'))
        e_s = omml_to_latex(e) if e is not None else ""
        sup_s = omml_to_latex(sup) if sup is not None else ""
        return f"{{{e_s}}}^{{{sup_s}}}"

    if tag == 'sSub':
        e = node.find(qn('m:e'))
        sub = node.find(qn('m:sub'))
        e_s = omml_to_latex(e) if e is not None else ""
        sub_s = omml_to_latex(sub) if sub is not None else ""
        return f"{{{e_s}}}_{{{sub_s}}}"

    if tag == 'sSubSup':
        e = node.find(qn('m:e'))
        sub = node.find(qn('m:sub'))
        sup = node.find(qn('m:sup'))
        e_s = omml_to_latex(e) if e is not None else ""
        sub_s = omml_to_latex(sub) if sub is not None else ""
        sup_s = omml_to_latex(sup) if sup is not None else ""
        return f"{{{e_s}}}_{{{sub_s}}}^{{{sup_s}}}"

    if tag == 'rad':
        e = node.find(qn('m:e'))
        deg = node.find(qn('m:deg'))
        e_s = omml_to_latex(e) if e is not None else ""
        if deg is not None:
            deg_s = omml_to_latex(deg).strip()
            if deg_s:
                return f"\\sqrt[{deg_s}]{{{e_s}}}"
        return f"\\sqrt{{{e_s}}}"

    if tag == 'd':
        dPr = node.find(qn('m:dPr'))
        beg = "("
        end = ")"
        if dPr is not None:
            b_node = dPr.find(qn('m:begChr'))
            e_node = dPr.find(qn('m:endChr'))
            if b_node is not None:
                beg = b_node.get(qn('m:val'), '(')
            if e_node is not None:
                end = e_node.get(qn('m:val'), ')')

        e_children = [c for c in node if c.tag.split('}')[-1] == 'e']
        if len(e_children) > 1:
            inner_s = "; ".join(omml_to_latex(c) for c in e_children)
        elif len(e_children) == 1:
            inner_s = omml_to_latex(e_children[0])
        else:
            inner_s = "".join(omml_to_latex(c) for c in node if c.tag.split('}')[-1] != 'dPr')
        return f"{beg}{inner_s}{end}"

    if tag == 'func':
        fName = node.find(qn('m:fName'))
        e = node.find(qn('m:e'))
        fn_s = omml_to_latex(fName) if fName is not None else ""
        e_s = omml_to_latex(e) if e is not None else ""
        return f"{fn_s} {e_s}"

    if tag == 'acc':
        e = node.find(qn('m:e'))
        e_s = omml_to_latex(e) if e is not None else ""
        return f"\\vec{{{e_s}}}"

    parts = []
    for c in node:
        c_tag = c.tag.split('}')[-1]
        if c_tag not in ['dPr', 'fPr', 'radPr', 'sSupPr', 'sSubPr', 'sSubSupPr']:
            parts.append(omml_to_latex(c))
    return "".join(parts)


def convert_image_bytes_to_png_base64(image_bytes, ext):
    """Chuyển đổi dữ liệu ảnh (bao gồm WMF/EMF của MathType) sang PNG base64 để web hiển thị"""
    try:
        if ext in ['png', 'jpg', 'jpeg', 'gif']:
            return f"data:image/{ext};base64," + base64.b64encode(image_bytes).decode('utf-8')
        
        # Mở và chuyển định dạng WMF / EMF sang PNG bằng Pillow
        with Image.open(io.BytesIO(image_bytes)) as img:
            out_buffer = io.BytesIO()
            img.save(out_buffer, format="PNG")
            png_bytes = out_buffer.getvalue()
            return "data:image/png;base64," + base64.b64encode(png_bytes).decode('utf-8')
    except Exception:
        return None


def extract_embedded_images(doc, r_elem):
    """Trích xuất an toàn hình ảnh MathType / Drawing nhúng và chuyển sang thẻ <img> chuẩn"""
    img_html = []
    try:
        # 1. Quét Drawing/Picture (blip)
        for blip in r_elem.findall(".//{*}blip"):
            rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed') or blip.get(qn('r:embed'))
            if rId and rId in doc.part.related_parts:
                part = doc.part.related_parts[rId]
                data = part.blob
                ext = part.partname.split('.')[-1].lower()
                b64_src = convert_image_bytes_to_png_base64(data, ext)
                if b64_src:
                    img_html.append(f'<img src="{b64_src}" class="inline-block max-h-10 align-middle mx-1 my-0.5" />')

        # 2. Quét VML imagedata (MathType OLE Picture WMF/EMF)
        for img_data in r_elem.findall(".//{*}imagedata"):
            rId = img_data.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id') or img_data.get(qn('r:id'))
            if rId and rId in doc.part.related_parts:
                part = doc.part.related_parts[rId]
                data = part.blob
                ext = part.partname.split('.')[-1].lower()
                b64_src = convert_image_bytes_to_png_base64(data, ext)
                if b64_src:
                    img_html.append(f'<img src="{b64_src}" class="inline-block max-h-10 align-middle mx-1 my-0.5" />')
    except Exception:
        pass

    return "".join(img_html)


def get_paragraph_full_display_html(doc, p_elem):
    """
    Lấy chuỗi hiển thị hoàn chỉnh cho giao diện Web:
    - Text thông thường
    - Công thức toán OMML ($LaTeX$)
    - Hình ảnh đồ thị và công thức nhúng MathType (PNG Base64)
    """
    parts = []
    for child in p_elem:
        tag = child.tag.split('}')[-1]
        if tag == 'r':
            r_text = "".join(t.text for t in child.findall(f".//{qn('w:t')}") if t.text)
            if r_text:
                parts.append(r_text)
            
            # Quét và nhúng ảnh MathType / Hình vẽ
            imgs = extract_embedded_images(doc, child)
            if imgs:
                parts.append(f" {imgs} ")

        elif tag in ['oMath', 'oMathPara']:
            latex = omml_to_latex(child).strip()
            if latex:
                parts.append(f" ${latex}$ ")
        elif tag == 'tab':
            parts.append("    ")

    full = "".join(parts)
    full = re.sub(r"[ ]{2,}", " ", full)
    return full.strip()


def parse_inline_paragraph_choices(doc, p, display_html):
    text = display_html
    matches = list(re.finditer(r"(?:^|[\t\s]|>)([A-Da-d][\.:\)])\s*", text))
    if len(matches) <= 1:
        return None

    separated_choices = []
    indices = [m.start(1) for m in matches] + [len(text)]

    for i in range(len(matches)):
        m = matches[i]
        label = m.group(1)[0].upper()
        c_start = indices[i]
        c_end = indices[i+1]
        choice_text = text[c_start:c_end].strip()

        is_corr = choice_text.startswith("#") or choice_text.startswith("*") or check_paragraph_has_correct_mark(p._p)

        separated_choices.append({
            "label": label,
            "is_fixed": False,
            "is_correct": is_corr,
            "xml": deepcopy(p._p),
            "text": choice_text
        })

    return separated_choices


def parse_docx(file_path):
    doc = docx.Document(file_path)

    group_regex = re.compile(r"^<(#?)[qg]([0-3])>", re.IGNORECASE)
    question_regex = re.compile(r"^(Câu|Question)\s+(\d+)[\.:]", re.IGNORECASE)
    choice_single_regex = re.compile(r"^\s*(#?|\*?)([A-Da-d])[\.:\)]\s*(.*)", re.IGNORECASE)

    groups = []
    current_group = None
    current_question = None

    for p in doc.paragraphs:
        display_html = get_paragraph_full_display_html(doc, p._p)
        if not display_html and not list(p._p):
            continue

        # BỎ QUA TOÀN BỘ PHẦN TIÊU ĐỀ ĐỀ GỐC (Bao gồm chữ ĐỀ CHÍNH THỨC, Sở, Trường, SBD...)
        p_text_upper = display_html.upper()
        if any(kw in p_text_upper for kw in ["ĐỀ CHÍNH THỨC", "DE CHINH THUC", "ĐỀ THI CÓ", "SỐ BÁO DANH", "HỌ VÀ TÊN:"]):
            continue

        # 1. Nhận diện nhóm thẻ <#q...>, <q...>
        g_match = group_regex.match(display_html) if display_html else None
        if g_match:
            is_fixed = bool(g_match.group(1))
            mode = int(g_match.group(2))
            cleaned_title = group_regex.sub("", display_html).strip()
            title_lower = cleaned_title.lower()

            part_type = "part_1"
            if "trả lời ngắn" in title_lower or "phần iii" in title_lower or "phan iii" in title_lower:
                part_type = "part_3"
            elif "đúng sai" in title_lower or "phần ii" in title_lower or "phan ii" in title_lower:
                part_type = "part_2"

            current_group = {
                "tag": g_match.group(0),
                "title": cleaned_title,
                "part_type": part_type,
                "is_fixed_group": is_fixed,
                "shuffle_questions": mode in [1, 3],
                "shuffle_choices": mode in [2, 3],
                "questions": []
            }
            groups.append(current_group)
            current_question = None
            continue

        # 2. Nhận diện câu hỏi mới
        q_match = question_regex.match(display_html) if display_html else None
        if q_match:
            if current_group is None:
                current_group = {
                    "tag": "<q3>",
                    "title": "",
                    "part_type": "part_1",
                    "is_fixed_group": False,
                    "shuffle_questions": True,
                    "shuffle_choices": True,
                    "questions": []
                }
                groups.append(current_group)

            current_question = {
                "original_num": q_match.group(2),
                "body_elements": [deepcopy(p._p)],
                "text": display_html,
                "choices": []
            }
            current_group["questions"].append(current_question)
            continue

        # NẾU CHƯA VÀO CÂU HỎI NÀO THÌ BỎ QUA LUÔN (Tránh gom bảng tiêu đề gốc)
        if current_question is None:
            continue

        # 3. PHẦN III: Nhận diện dòng đáp án sau chữ A.
        if current_group.get("part_type") == "part_3":
            is_p3_ans = bool(re.match(r"^\s*(?:#?|\*?)[Aa][\.:\)]", display_html))
            if is_p3_ans:
                ans_content = re.sub(r"^\s*(?:#?|\*?)[Aa][\.:\)]\s*", "", display_html).strip()
                ans_clean = re.sub(r"[\s\.]+$", "", ans_content).strip().replace("$", "").strip()

                current_question["choices"].append({
                    "label": "A",
                    "is_fixed": True,
                    "is_correct": True,
                    "xml": deepcopy(p._p),
                    "text": ans_clean
                })
                continue
            else:
                current_question["body_elements"].append(deepcopy(p._p))
                continue

        # 4. PHẦN I & II: Phương án inline
        inline = parse_inline_paragraph_choices(doc, p, display_html)
        if inline:
            current_question["choices"].extend(inline)
            continue

        # 5. PHẦN I & II: Phương án đơn dòng
        c_single = choice_single_regex.match(display_html) if display_html else None
        if c_single:
            is_corr = display_html.startswith("#") or display_html.startswith("*") or check_paragraph_has_correct_mark(p._p)
            lbl = c_single.group(2).upper() if current_group.get("part_type") == "part_1" else c_single.group(2).lower()
            current_question["choices"].append({
                "label": lbl,
                "is_fixed": bool(c_single.group(1)),
                "is_correct": is_corr,
                "xml": deepcopy(p._p),
                "text": display_html
            })
            continue

        # 6. Đoạn phụ, hình vẽ nằm bên dưới câu hỏi
        current_question["body_elements"].append(deepcopy(p._p))

    return groups
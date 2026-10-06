from parser import parse_docx
from shuffler import shuffle_and_export
import os

def run_test():
    input_file = "uploads/DE 01.docx"
    exam_code = "801"
    output_file = f"outputs/Ma_de_{exam_code}.docx"
    
    print(f"1. Đang bóc tách file: {input_file}...")
    groups = parse_docx(input_file)
    
    print(f"2. Đang trộn và xuất mã đề {exam_code}...")
    answers = shuffle_and_export(groups, input_file, exam_code, output_file)
    
    print(f"\n THÀNH CÔNG! Đã tạo file: {output_file}")
    print("--- BẢNG ĐÁP ÁN MÃ ĐỀ 801 ---")
    for q_no, ans in answers.items():
        print(f"Câu {q_no}: {ans}", end=" | ")
        if q_no % 5 == 0:
            print()
    print("\n------------------------------")

if __name__ == "__main__":
    run_test()
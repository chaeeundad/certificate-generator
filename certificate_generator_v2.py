#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import shutil
import zipfile
from datetime import datetime
import re
from pathlib import Path

class CertificateGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("교육 이수증 자동 생성 프로그램 v2.0")
        self.root.geometry("900x750")

        self.template_path = None
        self.csv_path = None
        self.output_dir = None

        self.setup_ui()

    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Template selection
        ttk.Label(main_frame, text="1. 템플릿 파일 선택:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.template_label = ttk.Label(main_frame, text="선택된 템플릿 없음", relief="sunken", width=50)
        self.template_label.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="템플릿 선택", command=self.select_template).grid(row=0, column=2, pady=5)

        # CSV file selection
        ttk.Label(main_frame, text="2. CSV 파일 선택:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.csv_label = ttk.Label(main_frame, text="선택된 CSV 없음", relief="sunken", width=50)
        self.csv_label.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="CSV 선택", command=self.select_csv).grid(row=1, column=2, pady=5)

        # Output directory selection
        ttk.Label(main_frame, text="3. 출력 폴더 선택:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.output_label = ttk.Label(main_frame, text="선택된 폴더 없음", relief="sunken", width=50)
        self.output_label.grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="폴더 선택", command=self.select_output_dir).grid(row=2, column=2, pady=5)

        # Separator
        ttk.Separator(main_frame, orient='horizontal').grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        # Additional information
        ttk.Label(main_frame, text="4. 교육 정보 입력:", font=('', 10, 'bold')).grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=5)

        # Course name
        ttk.Label(main_frame, text="교육과정명:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.course_name_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.course_name_var, width=50).grid(row=5, column=1, columnspan=2, pady=5, sticky=(tk.W, tk.E))

        # Education period
        ttk.Label(main_frame, text="교육기간:").grid(row=6, column=0, sticky=tk.W, pady=5)
        period_frame = ttk.Frame(main_frame)
        period_frame.grid(row=6, column=1, columnspan=2, sticky=tk.W, pady=5)

        self.start_date_var = tk.StringVar()
        ttk.Entry(period_frame, textvariable=self.start_date_var, width=15).grid(row=0, column=0)
        ttk.Label(period_frame, text=" ~ ").grid(row=0, column=1)
        self.end_date_var = tk.StringVar()
        ttk.Entry(period_frame, textvariable=self.end_date_var, width=15).grid(row=0, column=2)
        ttk.Label(period_frame, text=" (예: 2024.09.17)").grid(row=0, column=3, padx=5)

        # Education hours
        ttk.Label(main_frame, text="교육시간:").grid(row=7, column=0, sticky=tk.W, pady=5)
        self.hours_var = tk.StringVar()
        hours_frame = ttk.Frame(main_frame)
        hours_frame.grid(row=7, column=1, sticky=tk.W, pady=5)
        ttk.Entry(hours_frame, textvariable=self.hours_var, width=10).grid(row=0, column=0)
        ttk.Label(hours_frame, text=" 시간").grid(row=0, column=1)

        # Issue date
        ttk.Label(main_frame, text="발급일자:").grid(row=8, column=0, sticky=tk.W, pady=5)
        self.issue_date_var = tk.StringVar(value=datetime.now().strftime("%Y년 %m월 %d일"))
        ttk.Entry(main_frame, textvariable=self.issue_date_var, width=20).grid(row=8, column=1, sticky=tk.W, pady=5)

        # Organization name
        ttk.Label(main_frame, text="발급기관:").grid(row=9, column=0, sticky=tk.W, pady=5)
        self.org_name_var = tk.StringVar(value="주식회사 버추얼랩")
        ttk.Entry(main_frame, textvariable=self.org_name_var, width=50).grid(row=9, column=1, columnspan=2, pady=5, sticky=(tk.W, tk.E))

        # Certificate start number
        ttk.Label(main_frame, text="이수증 시작 번호:").grid(row=10, column=0, sticky=tk.W, pady=5)
        self.start_number_var = tk.StringVar(value="1")
        start_number_frame = ttk.Frame(main_frame)
        start_number_frame.grid(row=10, column=1, sticky=tk.W, pady=5)
        ttk.Entry(start_number_frame, textvariable=self.start_number_var, width=10).grid(row=0, column=0)
        ttk.Label(start_number_frame, text=" (예: 1, 100, 2024001)").grid(row=0, column=1, padx=5)

        # Separator
        ttk.Separator(main_frame, orient='horizontal').grid(row=11, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        # Output Options
        ttk.Label(main_frame, text="5. 출력 옵션:", font=('', 10, 'bold')).grid(row=12, column=0, columnspan=3, sticky=tk.W, pady=5)

        options_frame = ttk.Frame(main_frame)
        options_frame.grid(row=13, column=0, columnspan=3, sticky=tk.W, pady=5)

        # Radio buttons for output type
        self.output_type = tk.StringVar(value="individual")
        ttk.Radiobutton(options_frame, text="개별 파일로 생성", variable=self.output_type, value="individual").grid(row=0, column=0, padx=10)
        ttk.Radiobutton(options_frame, text="하나의 파일로 병합", variable=self.output_type, value="merged").grid(row=0, column=1, padx=10)

        # Separator
        ttk.Separator(main_frame, orient='horizontal').grid(row=14, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.grid(row=15, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        # Status label
        self.status_label = ttk.Label(main_frame, text="준비 완료", relief="sunken")
        self.status_label.grid(row=16, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        # Generate button
        self.generate_button = ttk.Button(main_frame, text="이수증 생성", command=self.generate_certificates, state='disabled')
        self.generate_button.grid(row=17, column=0, columnspan=3, pady=20)

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

    def select_template(self):
        filename = filedialog.askopenfilename(
            title="템플릿 파일 선택",
            filetypes=[("HWPX files", "*.hwpx"), ("All files", "*.*")]
        )
        if filename:
            self.template_path = filename
            self.template_label.config(text=os.path.basename(filename))
            self.check_ready()

    def select_csv(self):
        filename = filedialog.askopenfilename(
            title="CSV 파일 선택",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.csv_path = filename
            self.csv_label.config(text=os.path.basename(filename))
            # Show preview of CSV data
            try:
                df = pd.read_csv(filename, encoding='utf-8-sig')
                preview = f"데이터 미리보기: {len(df)}명의 교육생 정보 로드됨\n"
                preview += f"컬럼: {', '.join(df.columns.tolist())}"
                self.status_label.config(text=preview)
            except Exception as e:
                self.status_label.config(text=f"CSV 읽기 오류: {str(e)}")
            self.check_ready()

    def select_output_dir(self):
        dirname = filedialog.askdirectory(title="출력 폴더 선택")
        if dirname:
            self.output_dir = dirname
            self.output_label.config(text=dirname)
            self.check_ready()

    def check_ready(self):
        if self.template_path and self.csv_path and self.output_dir:
            self.generate_button.config(state='normal')
        else:
            self.generate_button.config(state='disabled')

    def validate_inputs(self):
        if not self.course_name_var.get():
            messagebox.showerror("오류", "교육과정명을 입력해주세요.")
            return False
        if not self.start_date_var.get() or not self.end_date_var.get():
            messagebox.showerror("오류", "교육기간을 입력해주세요.")
            return False
        if not self.hours_var.get():
            messagebox.showerror("오류", "교육시간을 입력해주세요.")
            return False
        if not self.org_name_var.get():
            messagebox.showerror("오류", "발급기관을 입력해주세요.")
            return False
        try:
            int(self.start_number_var.get())
        except ValueError:
            messagebox.showerror("오류", "시작 번호는 숫자여야 합니다.")
            return False
        return True

    def generate_certificates(self):
        if not self.validate_inputs():
            return

        try:
            # Read CSV file
            self.status_label.config(text="CSV 파일 읽는 중...")
            df = pd.read_csv(self.csv_path, encoding='utf-8-sig')

            total = len(df)
            self.progress['maximum'] = total

            # Get start number
            start_number = int(self.start_number_var.get())

            # Create output subdirectory with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_subdir = os.path.join(self.output_dir, f"이수증_{timestamp}")
            os.makedirs(output_subdir, exist_ok=True)

            if self.output_type.get() == "individual":
                # Generate individual files
                for index, row in df.iterrows():
                    cert_number = start_number + index
                    self.status_label.config(text=f"생성 중... ({index+1}/{total}) - 번호: {cert_number}")
                    self.progress['value'] = index + 1
                    self.root.update()
                    self.create_certificate(row, output_subdir, cert_number)

                self.status_label.config(text=f"완료! {total}개의 이수증이 생성되었습니다.")
                messagebox.showinfo("완료", f"{total}개의 이수증이 성공적으로 생성되었습니다.\n\n저장 위치: {output_subdir}")

            else:  # merged
                self.status_label.config(text="병합된 이수증 생성 중...")
                output_file = self.create_merged_certificate(df, output_subdir, start_number)
                self.status_label.config(text=f"완료! {total}명의 이수증이 하나의 파일로 생성되었습니다.")
                messagebox.showinfo("완료", f"{total}명의 이수증이 하나의 파일로 병합되었습니다.\n\n파일 위치: {output_file}")

        except Exception as e:
            messagebox.showerror("오류", f"이수증 생성 중 오류가 발생했습니다:\n{str(e)}")
            self.status_label.config(text="오류 발생")

    def create_certificate(self, person_data, output_dir, cert_number):
        """Create individual certificate"""
        # Get person information
        name = self.get_csv_value(person_data, ['이름', 'name', '성명', '교육자성명'])

        if not name:
            raise ValueError(f"행 {cert_number}에 이름 정보가 없습니다.")

        # Create temporary directory
        temp_dir = os.path.join(output_dir, f"temp_{cert_number}")
        os.makedirs(temp_dir, exist_ok=True)

        try:
            # Extract template (HWPX is a ZIP file)
            with zipfile.ZipFile(self.template_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            # Get replacements
            replacements = self.get_replacements(person_data, cert_number)

            # Process Contents/section0.xml
            section_path = os.path.join(temp_dir, "Contents", "section0.xml")
            if os.path.exists(section_path):
                with open(section_path, 'r', encoding='utf-8') as f:
                    content = f.read()

                # Replace all placeholders
                for placeholder, value in replacements.items():
                    content = content.replace(placeholder, str(value))

                with open(section_path, 'w', encoding='utf-8') as f:
                    f.write(content)

            # Create new HWPX file
            output_filename = f"이수증_{cert_number:04d}_{name}.hwpx"
            output_path = os.path.join(output_dir, output_filename)

            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, temp_dir)
                        zipf.write(file_path, arcname)

        finally:
            # Clean up temporary directory
            shutil.rmtree(temp_dir, ignore_errors=True)

    def create_merged_certificate(self, df, output_dir, start_number):
        """Create a single HWPX file with all certificates"""
        # Create temporary directory
        temp_dir = os.path.join(output_dir, "temp_merged")
        os.makedirs(temp_dir, exist_ok=True)

        try:
            # Extract template
            with zipfile.ZipFile(self.template_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            # Read original section0.xml
            section_path = os.path.join(temp_dir, "Contents", "section0.xml")
            if not os.path.exists(section_path):
                raise ValueError("Contents/section0.xml 파일을 찾을 수 없습니다.")

            with open(section_path, 'r', encoding='utf-8') as f:
                original_content = f.read()

            # Find the body content pattern
            # Look for the content between specific tags that contain the certificate template
            import re

            # Extract the part that contains placeholders
            # This pattern may need adjustment based on actual HWPX structure
            body_pattern = r'(<hp:p[^>]*>.*?</hp:p>)'
            body_matches = re.findall(body_pattern, original_content, re.DOTALL)

            if not body_matches:
                # If no matches, try to find any content with placeholders
                if '{' in original_content and '}' in original_content:
                    # Use the entire content as template
                    template_body = original_content
                else:
                    raise ValueError("템플릿에서 placeholder를 찾을 수 없습니다.")
            else:
                # Find bodies that contain placeholders
                template_body = None
                for body in body_matches:
                    if '{' in body and '}' in body:
                        template_body = body
                        break

                if not template_body:
                    # Use all bodies
                    template_body = ''.join(body_matches)

            # Process each person
            all_bodies = []
            for index, row in df.iterrows():
                cert_number = start_number + index
                self.status_label.config(text=f"병합 처리 중... ({index+1}/{len(df)}) - 번호: {cert_number}")
                self.progress['value'] = index + 1
                self.root.update()

                # Get replacements for this person
                replacements = self.get_replacements(row, cert_number)

                # Create a copy of the template body
                person_body = template_body

                # Replace placeholders
                for placeholder, value in replacements.items():
                    person_body = person_body.replace(placeholder, str(value))

                # Add page break after each certificate (except the last one)
                if index < len(df) - 1:
                    # HWPX page break tag
                    page_break = '<hp:p><hp:lineBreak lineBreakType="PAGE"/></hp:p>'
                    person_body += page_break

                all_bodies.append(person_body)

            # Combine all bodies
            combined_bodies = '\n'.join(all_bodies)

            # Replace the original body with combined bodies
            if body_matches:
                # Replace first occurrence with combined content, remove others
                new_content = original_content
                for i, body in enumerate(body_matches):
                    if i == 0:
                        new_content = new_content.replace(body, combined_bodies, 1)
                    else:
                        new_content = new_content.replace(body, '', 1)
            else:
                # Just replace the entire content
                new_content = combined_bodies

            # Write the modified content
            with open(section_path, 'w', encoding='utf-8') as f:
                f.write(new_content)

            # Create final HWPX file
            output_filename = f"이수증_병합_{len(df)}명_{start_number}번부터.hwpx"
            output_path = os.path.join(output_dir, output_filename)

            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, temp_dir)
                        zipf.write(file_path, arcname)

            return output_path

        finally:
            # Clean up
            shutil.rmtree(temp_dir, ignore_errors=True)

    def get_csv_value(self, row_data, possible_columns):
        """Get value from CSV row trying multiple possible column names"""
        for col in possible_columns:
            if isinstance(row_data, dict):
                if col in row_data and pd.notna(row_data[col]):
                    return str(row_data[col])
            else:
                try:
                    if pd.notna(row_data[col]):
                        return str(row_data[col])
                except:
                    pass
        return ""

    def get_replacements(self, person_data, cert_number):
        """Get all replacement values for placeholders"""
        # Extract values from CSV
        name = self.get_csv_value(person_data, ['이름', 'name', '성명', '교육자성명'])
        birth_date = self.get_csv_value(person_data, ['생년월일', 'birth_date', '생일', '교육자생년월일'])
        department = self.get_csv_value(person_data, ['부서', 'department', '소속', '교육자소속'])

        # Get current year
        current_year = datetime.now().year

        replacements = {
            # Template specific placeholders
            '{교육년도}': str(current_year),
            '{이수증번호}': str(cert_number).zfill(4),
            '{교육자소속}': department,
            '{교육자성명}': name,
            '{교육자생년월일}': birth_date,
            '{과정명}': self.course_name_var.get(),
            '{교육기간시작일}': self.start_date_var.get(),
            '{교육기간종료일}': self.end_date_var.get(),
            '{발행일자}': self.issue_date_var.get(),
        }

        return replacements

def main():
    root = tk.Tk()
    app = CertificateGenerator(root)

    # Check if template exists in current directory
    template_file = "교육이수증_템플릿.hwpx"
    if os.path.exists(template_file):
        app.template_path = os.path.abspath(template_file)
        app.template_label.config(text=template_file)
        app.check_ready()

    root.mainloop()

if __name__ == "__main__":
    main()
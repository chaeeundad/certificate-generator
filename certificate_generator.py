#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import shutil
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
import re
from pathlib import Path

class CertificateGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("교육 이수증 자동 생성 프로그램 v2.0")
        self.root.geometry("900x700")

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
        self.org_name_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.org_name_var, width=50).grid(row=9, column=1, columnspan=2, pady=5, sticky=(tk.W, tk.E))

        # Separator
        ttk.Separator(main_frame, orient='horizontal').grid(row=10, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        # Output Options
        ttk.Label(main_frame, text="5. 출력 옵션:", font=('', 10, 'bold')).grid(row=11, column=0, columnspan=3, sticky=tk.W, pady=5)

        options_frame = ttk.Frame(main_frame)
        options_frame.grid(row=12, column=0, columnspan=3, sticky=tk.W, pady=5)

        # Radio buttons for output type
        self.output_type = tk.StringVar(value="individual")
        ttk.Radiobutton(options_frame, text="개별 파일로 생성", variable=self.output_type, value="individual").grid(row=0, column=0, padx=10)
        ttk.Radiobutton(options_frame, text="하나의 파일로 병합", variable=self.output_type, value="merged").grid(row=0, column=1, padx=10)

        # Separator
        ttk.Separator(main_frame, orient='horizontal').grid(row=13, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.grid(row=14, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        # Status label
        self.status_label = ttk.Label(main_frame, text="준비 완료", relief="sunken")
        self.status_label.grid(row=15, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        # Generate button
        self.generate_button = ttk.Button(main_frame, text="이수증 생성", command=self.generate_certificates, state='disabled')
        self.generate_button.grid(row=16, column=0, columnspan=3, pady=20)

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

            # Create output subdirectory with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_subdir = os.path.join(self.output_dir, f"이수증_{timestamp}")
            os.makedirs(output_subdir, exist_ok=True)

            if self.output_type.get() == "individual":
                # Generate individual files
                for index, row in df.iterrows():
                    self.status_label.config(text=f"생성 중... ({index+1}/{total})")
                    self.progress['value'] = index + 1
                    self.root.update()
                    self.create_certificate(row, output_subdir, index + 1)

                self.status_label.config(text=f"완료! {total}개의 이수증이 생성되었습니다.")
                messagebox.showinfo("완료", f"{total}개의 이수증이 성공적으로 생성되었습니다.\n\n저장 위치: {output_subdir}")

            else:  # merged
                self.status_label.config(text="병합된 이수증 생성 중...")
                output_file = self.create_merged_certificate(df, output_subdir)
                self.status_label.config(text=f"완료! {total}명의 이수증이 하나의 파일로 생성되었습니다.")
                messagebox.showinfo("완료", f"{total}명의 이수증이 하나의 파일로 병합되었습니다.\n\n파일 위치: {output_file}")

        except Exception as e:
            messagebox.showerror("오류", f"이수증 생성 중 오류가 발생했습니다:\n{str(e)}")
            self.status_label.config(text="오류 발생")

    def create_certificate(self, person_data, output_dir, cert_number):
        # Extract person information from CSV row
        name = self.get_csv_value(person_data, ['이름', 'name', '성명'])
        birth_date = self.get_csv_value(person_data, ['생년월일', 'birth_date', '생일'])
        department = self.get_csv_value(person_data, ['부서', 'department', '소속'])

        if not name:
            raise ValueError(f"행 {cert_number}에 이름 정보가 없습니다.")

        # Create temporary directory for this certificate
        temp_dir = os.path.join(output_dir, f"temp_{cert_number}")
        os.makedirs(temp_dir, exist_ok=True)

        try:
            # Extract template
            with zipfile.ZipFile(self.template_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            # Process all XML files
            self.process_hwpx_files(temp_dir, person_data, cert_number)

            # Create new HWPX file
            output_filename = f"이수증_{cert_number:03d}_{name}.hwpx"
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

    def create_merged_certificate(self, df, output_dir):
        """Create a single HWPX file with all certificates"""
        # Create temporary directories
        temp_base = os.path.join(output_dir, "temp_merged")
        os.makedirs(temp_base, exist_ok=True)

        merged_sections = []

        try:
            # Process each person's certificate
            for index, row in df.iterrows():
                self.status_label.config(text=f"병합 처리 중... ({index+1}/{len(df)})")
                self.progress['value'] = index + 1
                self.root.update()

                # Extract template to a temporary directory for this person
                temp_dir = os.path.join(temp_base, f"person_{index}")
                os.makedirs(temp_dir, exist_ok=True)

                with zipfile.ZipFile(self.template_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)

                # Process template with this person's data
                self.process_hwpx_files(temp_dir, row, index + 1)

                # Store the processed section
                merged_sections.append(temp_dir)

            # Now merge all sections into one HWPX
            self.status_label.config(text="섹션 병합 중...")
            output_path = self.merge_sections_to_hwpx(merged_sections, output_dir, len(df))

            return output_path

        finally:
            # Clean up temporary directory
            shutil.rmtree(temp_base, ignore_errors=True)

    def merge_sections_to_hwpx(self, section_dirs, output_dir, total_count):
        """Merge multiple HWPX sections into a single file"""
        if not section_dirs:
            raise ValueError("No sections to merge")

        # Use first section as base
        merged_dir = os.path.join(output_dir, "final_merged")
        os.makedirs(merged_dir, exist_ok=True)

        # Copy first person's files as base
        first_dir = section_dirs[0]
        for item in os.listdir(first_dir):
            s = os.path.join(first_dir, item)
            d = os.path.join(merged_dir, item)
            if os.path.isdir(s):
                shutil.copytree(s, d)
            else:
                shutil.copy2(s, d)

        # For HWPX, we need to handle Contents/section*.xml files
        contents_dir = os.path.join(merged_dir, "Contents")
        if os.path.exists(contents_dir):
            # Find the main section file
            section_files = [f for f in os.listdir(contents_dir) if f.startswith("section") and f.endswith(".xml")]

            if section_files:
                base_section_name = section_files[0]
                base_section_path = os.path.join(contents_dir, base_section_name)

                # Read base section to get structure
                with open(base_section_path, 'r', encoding='utf-8') as f:
                    base_content = f.read()

                # Create combined content with page breaks
                all_sections_content = []

                for i, section_dir in enumerate(section_dirs):
                    section_contents_dir = os.path.join(section_dir, "Contents")
                    if os.path.exists(section_contents_dir):
                        section_file_path = os.path.join(section_contents_dir, base_section_name)
                        if os.path.exists(section_file_path):
                            with open(section_file_path, 'r', encoding='utf-8') as f:
                                section_content = f.read()

                            # Add page break between sections (except for first)
                            if i > 0:
                                # HWPX uses specific tags for page breaks
                                # We'll try to insert content with page break
                                all_sections_content.append(self.add_page_break_to_section(section_content))
                            else:
                                all_sections_content.append(section_content)

                # Combine all sections
                if all_sections_content:
                    combined_content = self.combine_hwpx_sections(all_sections_content)

                    # Write combined content back
                    with open(base_section_path, 'w', encoding='utf-8') as f:
                        f.write(combined_content)

        # Create final HWPX file
        output_filename = f"이수증_병합_{total_count}명.hwpx"
        output_path = os.path.join(output_dir, output_filename)

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(merged_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, merged_dir)
                    zipf.write(file_path, arcname)

        # Clean up
        shutil.rmtree(merged_dir, ignore_errors=True)

        return output_path

    def add_page_break_to_section(self, section_content):
        """Add a page break at the beginning of a section"""
        # In HWPX, page breaks are typically represented with specific XML tags
        # This is a simplified approach - actual implementation depends on HWPX structure

        # Try to find the body content area and insert page break
        import re

        # Look for paragraph or text area
        if '<hp:p' in section_content:
            # Insert page break before first paragraph
            page_break = '<hp:p><hp:lineBreak lineBreakType="PAGE"/></hp:p>'
            section_content = re.sub(r'(<hp:p[^>]*>)', page_break + r'\1', section_content, count=1)

        return section_content

    def combine_hwpx_sections(self, sections):
        """Combine multiple HWPX sections into one"""
        if not sections:
            return ""

        if len(sections) == 1:
            return sections[0]

        # Parse first section to get structure
        try:
            import xml.etree.ElementTree as ET

            # Use first section as base
            combined = sections[0]

            # For additional sections, extract body content and append
            for section in sections[1:]:
                # This is simplified - actual implementation needs to handle HWPX XML namespace
                # and properly merge body content
                try:
                    # Extract content between body tags (simplified)
                    import re
                    body_match = re.search(r'(<hp:p.*?</hp:p>)', section, re.DOTALL)
                    if body_match:
                        body_content = body_match.group(1)
                        # Insert before closing tags
                        combined = re.sub(r'(</.*?>\s*$)', body_content + r'\1', combined)
                except:
                    pass

            return combined
        except:
            # If parsing fails, return first section
            return sections[0]

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
        # Extract values from CSV with multiple possible column names
        name = self.get_csv_value(person_data, ['이름', 'name', '성명', '교육자성명'])
        birth_date = self.get_csv_value(person_data, ['생년월일', 'birth_date', '생일', '교육자생년월일'])
        department = self.get_csv_value(person_data, ['부서', 'department', '소속', '교육자소속'])

        # Additional fields that might be in CSV
        phone = self.get_csv_value(person_data, ['전화번호', 'phone', '연락처'])
        email = self.get_csv_value(person_data, ['이메일', 'email', '메일'])
        employee_id = self.get_csv_value(person_data, ['사번', 'employee_id', '직원번호'])

        # Get current year for education year
        from datetime import datetime
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

            # Additional common placeholders for compatibility
            '{이름}': name,
            '{성명}': name,
            '{생년월일}': birth_date,
            '{부서}': department,
            '{소속}': department,
            '{교육과정명}': self.course_name_var.get(),
            '{교육과정}': self.course_name_var.get(),
            '{교육기간}': f"{self.start_date_var.get()} ~ {self.end_date_var.get()}",
            '{시작일}': self.start_date_var.get(),
            '{종료일}': self.end_date_var.get(),
            '{교육시간}': f"{self.hours_var.get()}시간",
            '{시간}': self.hours_var.get(),
            '{발급일자}': self.issue_date_var.get(),
            '{발급일}': self.issue_date_var.get(),
            '{발급기관}': self.org_name_var.get(),
            '{기관명}': self.org_name_var.get(),
            '{번호}': str(cert_number).zfill(4),
            '{순번}': str(cert_number),
            '{NO}': str(cert_number),
            '{No}': str(cert_number),
            '{no}': str(cert_number),
        }

        # Add optional fields if they exist
        if phone:
            replacements['{전화번호}'] = phone
            replacements['{연락처}'] = phone
        if email:
            replacements['{이메일}'] = email
            replacements['{메일}'] = email
        if employee_id:
            replacements['{사번}'] = employee_id
            replacements['{직원번호}'] = employee_id

        return replacements

    def process_hwpx_files(self, temp_dir, person_data, cert_number):
        """Process all XML files in the extracted HWPX"""
        replacements = self.get_replacements(person_data, cert_number)

        # Process all XML files
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                if file.endswith('.xml'):
                    file_path = os.path.join(root, file)
                    self.modify_xml_file(file_path, replacements)

    def modify_xml_file(self, xml_file, replacements):
        """Modify XML file with proper encoding handling"""
        try:
            # Parse XML
            tree = ET.parse(xml_file)
            root = tree.getroot()

            # Replace text in all elements
            self.replace_text_in_element(root, replacements)

            # Save modified XML with proper encoding
            tree.write(xml_file, encoding='utf-8', xml_declaration=True)

        except ET.ParseError:
            # If it's not a valid XML, try to replace as text file
            try:
                with open(xml_file, 'r', encoding='utf-8') as f:
                    content = f.read()

                for placeholder, value in replacements.items():
                    content = content.replace(placeholder, value)

                with open(xml_file, 'w', encoding='utf-8') as f:
                    f.write(content)
            except:
                pass  # Skip files that can't be processed

    def replace_text_in_element(self, element, replacements):
        """Recursively replace text in XML elements"""
        # Replace in element text
        if element.text:
            for placeholder, value in replacements.items():
                if placeholder in element.text:
                    element.text = element.text.replace(placeholder, value)

        # Replace in element tail
        if element.tail:
            for placeholder, value in replacements.items():
                if placeholder in element.tail:
                    element.tail = element.tail.replace(placeholder, value)

        # Replace in attributes
        for attr_name, attr_value in element.attrib.items():
            for placeholder, value in replacements.items():
                if placeholder in attr_value:
                    element.attrib[attr_name] = attr_value.replace(placeholder, value)

        # Recursively process child elements
        for child in element:
            self.replace_text_in_element(child, replacements)

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
import utils
import picture

import streamlit as st
import pandas as pd
from docx import Document
import datetime
from docxtpl import DocxTemplate
from io import BytesIO
import zipfile
import docx
from docx.oxml import OxmlElement
from docx.shared import Inches, Pt
from docx.oxml import register_element_cls
import copy
from docx.table import _Cell
from lxml import etree

NAME_KEY = 'student_name'
CERTIFICATE_KEY = 'certificate_number'
MACHINE_CATEGORY = 'machine_category'

CERT_HEIGHT_INCHES = 3.65
CERT_WIDTH_INCHES = 5.6

TRACTOR_CERT_HEIGHT = 5.63
TRACTOR_CERT_WIDTH = 8.04

register_element_cls('wp:anchor', picture.CT_Anchor)

def fit_more_rows(document):
  """Attempts to fit more rows on a page in a Word document.

  This function minimizes margins, removes unnecessary section breaks, 
  and adjusts paragraph settings to prevent unwanted page breaks.

  Args:
    document_path: Path to the Word document.
  """

  # 1. Minimize Margins (as in the previous example)
  for section in document.sections:
    section.top_margin = Inches(0.2)
    section.bottom_margin = Inches(0.2)
    section.left_margin = Inches(0.2)
    section.right_margin = Inches(0.2)
    section.page_height = Inches(11.69)
    section.page_width = Inches(8.27)

#   2. Remove Unnecessary Section Breaks 
#   (This might help if extra sections are causing page breaks)
  while len(document.sections) > 1:  # Keep only one section
    document.sections[0].start_type = WD_SECTION.CONTINUOUS 
    document.sections[-1]._remove() 

  # 3. Adjust Paragraph Settings to Avoid Page Breaks 
  for paragraph in document.paragraphs:
    paragraph.paragraph_format.keep_with_next = True  # Keep with next paragraph
    paragraph.paragraph_format.page_break_before = False  # No page break before
    paragraph.paragraph_format.widow_control = True # Prevents single lines from appearing at the top or bottom of a page
    paragraph.paragraph_format.keep_lines_together = True # Keep all lines of a paragraph together on a page

  return document


def choose_teacher(all_teachers):
    """Handles teacher selection and adding new teachers,
    controlled by a checkbox.
    Returns a tuple (teacher_name, company) if a teacher is selected or added.
    """
    selected_teacher = st.selectbox(
        "Выберите преподавателя из списка:", all_teachers, index=None, placeholder="Преподаватели"
    )

    if selected_teacher:
        st.write(f"Преподаватель: {selected_teacher}")

    # Checkbox to trigger new teacher input
    add_new = st.checkbox("Добавить нового преподавателя?")
    new_teacher = None

    if add_new:
        new_teacher = st.text_input("Введите ФИО нового преподавателя:")
        if st.button("Добавить преподавателя"):
            if new_teacher and new_teacher not in all_teachers:
                all_teachers.append(new_teacher)
                utils.save_data(all_teachers, "data/teachers.pickle")
                st.success(
                    f"Преподаватель '{new_teacher}' добавлен!"
                )
                return new_teacher
            else:
                st.warning("Введите ФИО преподавателя и место работы.")

    if selected_teacher:
        return selected_teacher

    return None


def choose_profession(all_professions):
    """Handles profession selection and adding new professions,
    controlled by a checkbox.
    """
    selected_item = st.selectbox(
        "Выберите профессию/программу обучение из следующих опций:",
        all_professions,
        index=None,
        placeholder="Начинайте вводить название программы",
    )

    if selected_item:
        st.write(f"Программа обучения: {selected_item}")

    # Checkbox to trigger new profession input
    add_new = st.checkbox("Добавить новую программу обучения?")
    new_profession = None

    if add_new:  # Only show input if checkbox is checked
        new_profession = st.text_input("Введите название новой программы:")
        new_code = st.text_input("Введите код:")
        if st.button("Добавить программу"):
            if new_profession:
                code_int = -1
                try:
                    code_int = int(new_code)
                    all_professions[new_profession] = [code_int]
                except Exception as e:
                    all_professions[new_profession] = [-1]
                utils.save_data(all_professions,
                                filename="data/professions.pickle")
                st.success(f"Программа '{new_profession}' добавлена!")
                # Return the new profession if added
                return f'{code_int} «{new_profession}»'
            else:
                st.warning("Программа уже существует или не введена.")
    if selected_item:
        selected_profession_code = all_professions[selected_item]
        selected_profession_code_str = ', '.join(
            str(code) for code in selected_profession_code)
        # Return selected_item otherwise
        return f'{selected_profession_code_str} «{selected_item}»'
    return None, None


def create_certificate(replacement_dict, students):
    if not students:
        return Document()

    all_paragraphs = []
    for student in students[1:]:
        local_dict = replacement_dict.copy()
        local_dict[NAME_KEY] = student[NAME_KEY]
        local_dict[CERTIFICATE_KEY] = student[CERTIFICATE_KEY]

        doc = DocxTemplate('templates/свидетельство.docx')
        doc.render(local_dict)

        paragraphs = doc.tables[0].cell(0, 0).paragraphs
        all_paragraphs.append(paragraphs)

    # Create final document using the first student's data as a base
    final_doc = DocxTemplate('templates/свидетельство.docx')
    local_dict = replacement_dict.copy()
    local_dict[NAME_KEY] = students[0][NAME_KEY]
    local_dict[CERTIFICATE_KEY] = students[0][CERTIFICATE_KEY]
    final_doc.render(local_dict)

    # Set default font style
    utils.set_default_font(final_doc, bold=True)

    # Get the table and add rows for each additional student
    table = final_doc.tables[0]
    for paragraphs in all_paragraphs:
        row = table.add_row()
        # this ensure that the rows are not split between pages https://github.com/python-openxml/python-docx/issues/245
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement('w:cantSplit'))

        target_cell = row.cells[0]
        source_cell = table.cell(0, 0)  # Use the first cell as a template

        # Copy cell properties from the template cell
        utils.copy_cell_properties(source_cell, target_cell)

        # Add paragraphs and runs to the new cell, copying formatting
        for p_i, paragraph in enumerate(paragraphs):
            new_paragraph = target_cell.add_paragraph('')
            source_paragraph = source_cell.paragraphs[p_i]
            if p_i == 0:
                picture.add_float_picture(new_paragraph, 'pictures/basic-cert-background.png',
                                          width=Inches(CERT_WIDTH_INCHES), height=Inches(CERT_HEIGHT_INCHES))

            new_paragraph.alignment = source_paragraph.alignment
            new_paragraph.paragraph_format.left_indent = source_paragraph.paragraph_format.left_indent
            for target_run, source_run in zip(paragraph.runs, source_paragraph.runs):
                new_run = new_paragraph.add_run(target_run.text)
                utils.preserve_formatting(new_run, source_run)
    return final_doc

def dump(string_to_dump, name): 
    file_path = f"{name}.txt"

    with open(file_path, 'w') as file:
        file.write(string_to_dump)

def copy_table_element(source_tbl, target_tbl, element_name):
    """Copies a specified table element from source to target table."""
    source_element = source_tbl.find(element_name, namespaces=source_tbl.nsmap)
    if source_element is not None:
        target_element = target_tbl.find(element_name, namespaces=target_tbl.nsmap)
        if target_element is not None:
            target_tbl.remove(target_element)
        target_tbl.insert(0, copy.deepcopy(source_element)) 

def create_tractor_certificate(replacement_dict, students, picture_path):
    if not students:
        return Document()
    
    merged_doc = Document()
    merged_doc = fit_more_rows(merged_doc)
    
    merged_table = merged_doc.add_table(rows=len(students), cols=0)
    utils.set_default_font(merged_doc)
    curr_index = 0
    for student_index, student in enumerate(students):  # Skip the first student for now
        local_dict = replacement_dict.copy()
        local_dict[NAME_KEY] = student[NAME_KEY]
        local_dict[CERTIFICATE_KEY] = student[CERTIFICATE_KEY]
        local_dict[MACHINE_CATEGORY] = student[MACHINE_CATEGORY]

        doc = DocxTemplate('templates/certificate_tractor.docx')
        doc.render(local_dict)

        tbl = copy.deepcopy(doc.tables[0])
        
        if student_index == 0:
            xml = etree.tostring(tbl._tbl, encoding='unicode', pretty_print=True)
            dump(xml, 'og')
            for element_name in ['w:tblGrid', 'w:tblPr']:
                copy_table_element(tbl._tbl, merged_table._tbl, element_name)

        for row_index in range(len(tbl._tbl.findall('.//w:tr', namespaces=tbl._tbl.nsmap))):  # Iterate using XML
            print('current index:', curr_index)
            target_row = merged_table.rows[curr_index]._element
            source_row_element = tbl.rows[row_index]._element  # Get row's XML element
            
            # # --- Replace w:trPr in target_row ---
            source_trPr = source_row_element.find('w:trPr', namespaces=source_row_element.nsmap)
            target_trPr = target_row.find('w:trPr', namespaces=target_row.nsmap)

            if target_trPr is not None: 
                target_row.remove(target_trPr)

            if source_trPr is not None:
                print(etree.tostring(source_trPr, encoding='unicode', pretty_print=True))
                target_row.insert(0, copy.deepcopy(source_trPr))

            curr_index += 1

            # --- Copy cells from source row to target row ---
            for source_cell in source_row_element.findall('.//w:tc', namespaces=source_row_element.nsmap):
                target_row.append(copy.deepcopy(source_cell))
            
            first_cell = merged_table.rows[curr_index - 1].cells[0]
            p = first_cell.add_paragraph()
            picture.add_float_picture(p, picture_path, height=Inches(TRACTOR_CERT_HEIGHT), width=Inches(TRACTOR_CERT_WIDTH), pos_x=Pt(0), pos_y=Pt(0))
        
    xml = etree.tostring(merged_table._tbl, encoding='unicode', pretty_print=True)
    dump(xml, 'merged_table')

    final_table = merged_doc.tables[0]
    for row in final_table.rows:
        for cell in row.cells:
            cell.top_padding = Pt(0)
            cell.bottom_padding = Pt(0) 
    return merged_doc

def create_tractor_certs(beginning_dict, students): 
    blue = create_tractor_certificate(beginning_dict, students, 'pictures/tractor-background-blue.png')
    green = create_tractor_certificate(beginning_dict, students, 'pictures/tractor-background-green.png')
    return (blue, green)

def create_beginning_document(beginning_dict, students):
    """Creates a Word document with the provided information."""

    doc = DocxTemplate('templates/Приказ о начале.docx')
    doc.render(beginning_dict)
    utils.set_default_font(doc)
    table = doc.tables[0]

    for index, student in enumerate(students):
        # Create a new row
        new_row = table.add_row()

        # Populate cells in the new row
        new_row.cells[0].text = str(index+1)
        new_row.cells[1].text = student[NAME_KEY]
        new_row.cells[2].text = beginning_dict['student_company']

        # Add more cells for other data (profession, date, etc.) if needed
    return doc


def create_end_doc(replacement_dict, students):
    """Creates a Word document with the provided information."""

    doc = DocxTemplate('templates/Приказ о выпуске.docx')
    doc.render(replacement_dict)
    utils.set_default_font(doc)
    table = doc.tables[0]

    for index, student in enumerate(students):
        # Create a new row
        new_row = table.add_row()

        # Populate cells in the new row
        new_row.cells[0].text = str(index+1)
        new_row.cells[1].text = student[NAME_KEY]
        new_row.cells[2].text = replacement_dict['student_company']
        new_row.cells[3].text = str(student[CERTIFICATE_KEY])

    return doc


def create_protocol_doc(replacement_dict, students):
    """Creates a Word document with the provided information."""

    doc = DocxTemplate('templates/Протокол.docx')
    doc.render(replacement_dict)
    utils.set_default_font(doc)
    table = doc.tables[0]

    for index, student in enumerate(students):
        # Create a new row
        new_row = table.add_row()

        # Populate cells in the new row
        new_row.cells[0].text = str(index+1)
        new_row.cells[1].text = student[NAME_KEY]
        new_row.cells[2].text = replacement_dict['student_company']
        new_row.cells[3].text = str(student[CERTIFICATE_KEY])

    return doc


st.title("Профессиональное обучение")

# Input 1: Text Input
available_professions = utils.load_from_pickle('data/professions.pickle')
student_profession = choose_profession(available_professions)

today = datetime.date.today()
beginning_date = st.date_input('дата начала', value=today)
end_date = st.date_input('дата окончания', value=today)

beginning_number = st.number_input("номер приказа о начале", step=1, value=1, placeholder=808)
end_number = st.number_input("номер приказа об окончании", step=1, value=1, placeholder=808)

# this should be replaced by a scroll through
teacher_name = choose_teacher(utils.load_from_pickle('data/teachers.pickle'))

company = st.text_input("Предприятие", 'заявление',
                        placeholder="Наименование предприятия или 'заявление'")
student_names = st.text_area(
    "Введите имена студентов, по одному на строку").split("\n")
student_data = []
for line in student_names:
    if line:
        filtered_items = [item for item in line.split('\t') if item]
        certificate_number, _, _, name, *category = filtered_items  # Split each line by tab
        student_data.append({
            NAME_KEY: name,
            CERTIFICATE_KEY: int(float(certificate_number)),
            MACHINE_CATEGORY: category[0] if category else ''
        })
num_students = len(student_data)

formatted_beginning_date = utils.format_date(beginning_date)
formatted_end_date = utils.format_date(end_date)
replacement_dict = {
    'beginning_date': formatted_beginning_date,
    'beginning_number': beginning_number,
    'end_date': formatted_end_date,
    'end_number': end_number,
    'student_profession': student_profession,
    'student_company': company,
    'teacher_name': teacher_name,
    'num_students': num_students,
    'class': '4',
    'year': beginning_date.year,
}
beginning_doc = create_beginning_document(replacement_dict, student_data)
end_doc = create_end_doc(replacement_dict, student_data)
protocol = create_protocol_doc(replacement_dict, student_data)
certificate_docs = create_certificate(replacement_dict, student_data)
(blue_tractor_cert, green_tractor_cert) = create_tractor_certs(replacement_dict, student_data)

show_documents = st.button("Сгенерировать документы")

if show_documents:
    if not student_profession:
        st.warning('Укажите профессию')
    if not teacher_name:
        st.warning('Укажите преподователя')
    if not beginning_date:
        st.warning('Укажите дату начала')
    if not end_date:
        st.warning('Укажите дату окончания')
    if not beginning_number:
        st.warning('Укажите номер приказа о начале')
    if not end_number:
        st.warning('Укажите номер приказа о выпуске')
    if num_students == 0:
        st.warning('Укажите обучающихся')

    if student_profession and teacher_name and beginning_date and end_date and beginning_number and end_number:
        document_tabs = st.tabs(["Приказ о начале", "Приказ об окончании",
                                "Протокол", "Свидетельство", "Свидетельство тракторов синее", "Свидетельство тракторов зеленое"])
        with document_tabs[0]:  # Приказ о начале
            utils.display_docx_content(beginning_doc)

        with document_tabs[1]:  # Приказ об окончании
            utils.display_docx_content(end_doc)

        with document_tabs[2]:  # Протокол
            utils.display_docx_content(protocol)

        with document_tabs[3]:  # Свидетельство
            utils.display_docx_content(certificate_docs)

        with document_tabs[4]:  # Свидетельство тракторов
            utils.display_docx_content(blue_tractor_cert)
        
        with document_tabs[5]: 
            utils.display_docx_content(green_tractor_cert)

# --- Create a ZIP archive in memory ---
zip_buffer = BytesIO()
with zipfile.ZipFile(zip_buffer, 'w') as zipf:
    # Add the beginning document
    with zipf.open('Приказ о начале.docx', 'w') as f:
        beginning_doc.save(f)

    # Add the end document
    with zipf.open('Приказ о выпуске.docx', 'w') as f:
        end_doc.save(f)

    with zipf.open('Протокол.docx', 'w') as f:
        protocol.save(f)

    with zipf.open('Свидетельство.docx', 'w') as f:
        certificate_docs.save(f)

    with zipf.open('Свидетельство синее трактор.docx', 'w') as f:
        blue_tractor_cert.save(f)

    with zipf.open('Свидетельство зеленое трактор.docx', 'w') as f:
        green_tractor_cert.save(f)

zip_buffer.seek(0)

formatted_end_date = end_date.strftime("%d.%m.%Y")
# --- Download the ZIP archive ---
st.download_button(
    label="Скачать документы (ZIP)",
    data=zip_buffer,
    file_name=f'{formatted_end_date}.zip',
    mime='application/zip'
)

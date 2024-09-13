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

NAME_KEY = 'student_name'
CERTIFICATE_KEY = 'certificate_number'
MACHINE_CATEGORY = 'machine_category'

register_element_cls('wp:anchor', picture.CT_Anchor)

def choose_teacher(all_teachers):
    """Handles teacher selection and adding new teachers,
    controlled by a checkbox.
    Returns a tuple (teacher_name, company) if a teacher is selected or added.
    """
    selected_teacher = st.selectbox(
        "Выберите из списка:", all_teachers, index=None, placeholder="Преподаватели"
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
        "Выберите из следующих опций:",
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
                                          width=Inches(5.6), height=Inches(4.04), pos_x=Pt(0), pos_y=Pt(0))

            new_paragraph.alignment = source_paragraph.alignment
            new_paragraph.paragraph_format.left_indent = source_paragraph.paragraph_format.left_indent
            for target_run, source_run in zip(paragraph.runs, source_paragraph.runs):
                new_run = new_paragraph.add_run(target_run.text)
                utils.preserve_formatting(new_run, source_run)
    return final_doc


def create_tractor_certificate(replacement_dict, students):
    if not students:
        return Document()

    all_paragraphs = []
    for student in students[1:]:  # Skip the first student for now
        local_dict = replacement_dict.copy()
        local_dict[NAME_KEY] = student[NAME_KEY]
        local_dict[CERTIFICATE_KEY] = student[CERTIFICATE_KEY]

        doc = DocxTemplate('templates/Свидетельство_трактор.docx')
        doc.render(local_dict)  # Assuming you have a 'render' method defined

        paragraphs = doc.tables[0].cell(0, 0).paragraphs
        all_paragraphs.append(paragraphs)

    # Create final document using the first student's data as a base
    final_doc = DocxTemplate('templates/Свидетельство_трактор.docx')
    local_dict = replacement_dict.copy()
    local_dict[NAME_KEY] = students[0][NAME_KEY]
    local_dict[CERTIFICATE_KEY] = students[0][CERTIFICATE_KEY]
    final_doc.render(local_dict)

    utils.set_default_font(final_doc)

    # Get the table and add rows for each additional student
    table = final_doc.tables[0]
    for paragraphs in all_paragraphs:
        row = table.add_row()
        # this ensure that the rows are not split between pages https://github.com/python-openxml/python-docx/issues/245
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement('w:cantSplit'))

        left_cell = row.cells[0]
        right_cell = row.cells[1]

        source_right_cell = table.cell(0, 1)
        source_left_cell = table.cell(0, 0)

        # Copy cell properties from the template cell
        utils.copy_cell_properties(source_left_cell, left_cell)
        utils.copy_cell_properties(source_right_cell, right_cell)

        # Populate left cell
        for p_i, paragraph in enumerate(paragraphs):
            new_paragraph = left_cell.add_paragraph()
            if p_i == 0:
                picture.add_float_picture(new_paragraph, 'pictures/tractor-cert.png', width=Inches(
                    8.03), height=Inches(5.58), pos_x=Pt(0), pos_y=Pt(0))
            source_paragraph = source_left_cell.paragraphs[p_i]
            new_paragraph.alignment = source_paragraph.alignment
            new_paragraph.paragraph_format.left_indent = source_paragraph.paragraph_format.left_indent

            for target_run, source_run in zip(paragraph.runs, source_paragraph.runs):
                new_run = new_paragraph.add_run(target_run.text)
                new_run = utils.preserve_formatting(new_run, source_run)

        for p in source_right_cell.paragraphs:
            new_paragraph = right_cell.add_paragraph('')
            new_paragraph.alignment = p.alignment
            for source_run in p.runs:
                new_run = new_paragraph.add_run(source_run.text)
                new_run = utils.preserve_formatting(new_run, source_run)
    return final_doc


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

beginning_date = st.date_input('дата начала', value=datetime.date(2025, 1, 1))
end_date = st.date_input('дата окончания', value=datetime.date(2025, 1, 1))

beginning_number = st.number_input("номер приказа о начале", 808)
end_number = st.number_input("номер приказа об окончании", 808)

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
    'class': '3',
    'year': beginning_date.year,
}
beginning_doc = create_beginning_document(replacement_dict, student_data)
end_doc = create_end_doc(replacement_dict, student_data)
protocol = create_protocol_doc(replacement_dict, student_data)
certificate_docs = create_certificate(replacement_dict, student_data)
tractor_cert_doc = create_tractor_certificate(replacement_dict, student_data)

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
                                "Протокол", "Свидетельство", "Свидетельство тракторов"])
        with document_tabs[0]:  # Приказ о начале
            utils.display_docx_content(beginning_doc)

        with document_tabs[1]:  # Приказ об окончании
            utils.display_docx_content(end_doc)

        with document_tabs[2]:  # Протокол
            utils.display_docx_content(protocol)

        with document_tabs[3]:  # Свидетельство
            utils.display_docx_content(certificate_docs)

        with document_tabs[4]:  # Свидетельство тракторов
            utils.display_docx_content(tractor_cert_doc)

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

    with zipf.open('Свидетельство трактор.docx', 'w') as f:
        tractor_cert_doc.save(f)

zip_buffer.seek(0)

# --- Download the ZIP archive ---
st.download_button(
    label="Скачать документы (ZIP)",
    data=zip_buffer,
    file_name=f'{company}-{student_profession}-{end_date}.zip',
    mime='application/zip'
)

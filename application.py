import utils
from utils import Profession
import picture
import profession_parsing

import docx
import math
import streamlit as st
import datetime
from docxtpl import DocxTemplate
from io import BytesIO
import zipfile
import copy
from lxml import etree

from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Inches, Pt
from docx.oxml import register_element_cls
from docx.enum.table import WD_TABLE_ALIGNMENT
from dataclasses import dataclass
from dateutil.relativedelta import relativedelta
import csv
import io
import pprint
import pandas as pd

# --- Constants and Setup ---
NAME_KEY = "student_name"
CERTIFICATE_KEY = "certificate_number"
MACHINE_CATEGORY = "machine_category"
ROLE = "student_role"
RAZRYAD = "razryad"

TRACTOR_PROFESSION_WORDING = "19203 «Тракторист»"

CERT_HEIGHT_INCHES = Inches(3.74)
CERT_WIDTH_INCHES = Inches(5.59)

TRACTOR_CERT_HEIGHT = Inches(5.63)
TRACTOR_CERT_WIDTH = Inches(8.04)

register_element_cls("wp:anchor", picture.CT_Anchor)


def make_student_copy(replacement_dict, student):
    """Creates a copy of the replacement dict with student-specific data."""
    local_dict = replacement_dict.copy()
    local_dict[NAME_KEY] = student.name
    local_dict[CERTIFICATE_KEY] = student.cert_number
    local_dict[ROLE] = student.role
    local_dict[MACHINE_CATEGORY] = student.machine_category
    local_dict[RAZRYAD] = student.razryad
    return local_dict


class DocumentGenerator:
    """
    A factory class to generate various Word documents based on templates and student data.
    This class abstracts away the repetitive logic of document creation.
    """

    def __init__(self, replacement_dict, students):
        self.replacement_dict = replacement_dict
        self.students = students if students else []

    # --- Private Helper Methods for Document Generation ---

    def _create_list_based_document(self, template_path, row_populator_func):
        """
        Generic generator for documents with a table populated by a list of students.
        This pattern is used for orders (beginning/end) and protocols.

        Args:
            template_path (str): The path to the .docx template.
            row_populator_func (function): A function that takes (row, student, index, data)
                                           and populates the cells of a new row.
        """
        doc = DocxTemplate(template_path)
        doc.render(self.replacement_dict)
        utils.set_default_font(doc)

        if not self.students:
            return doc

        table = doc.tables[0]
        for index, student in enumerate(self.students):
            new_row = table.add_row()
            row_populator_func(new_row, student, index, self.replacement_dict)
        return doc

    def _create_merged_doc_from_template_rows(self, template_path, table_configs):
        """
        Generic generator for certificates created by merging rows from a template.
        Each student gets one row in the final document, built from a rendered template.
        This pattern is used for tractor, height, and confirmation page certificates.

        Args:
            template_path (str): The path to the .docx template.
            table_configs (list): A list of dictionaries, where each dict configures one
                                  table to be processed from the template (e.g., picture details).
        """
        if not self.students:
            return Document()

        merged_doc = Document()
        merged_doc = utils.fit_more_rows(merged_doc)
        utils.set_default_font(merged_doc)

        merged_tables = []
        for i, config in enumerate(table_configs):
            num_cols = config.get("cols", 1)
            merged_tables.append(
                merged_doc.add_table(rows=len(self.students), cols=num_cols)
            )
            if i > 0:
                merged_doc.add_page_break()

        for student_index, student in enumerate(self.students):
            local_dict = make_student_copy(self.replacement_dict, student)
            template_doc = DocxTemplate(template_path)
            template_doc.render(local_dict)

            for i, config in enumerate(table_configs):
                self._add_student_content_to_merged_table(
                    merged_table=merged_tables[i],
                    source_table=template_doc.tables[i],
                    student_index=student_index,
                    target_row_index=student_index,
                    picture_path=config.get("picture_path"),
                    picture_height=config.get("picture_height"),
                    picture_width=config.get("picture_width"),
                    picture_mode=config.get("picture_mode", "first_cell")
                )
        return merged_doc

    # --- Private Helper Methods for Content Copying (Moved from global scope) ---

    def _copy_text_and_formatting(self, source_cell, target_cell):
        utils.copy_cell_properties(source_cell, target_cell)
        for p_i, paragraph in enumerate(source_cell.paragraphs):
            if paragraph.text.strip() == "":
                continue
            new_paragraph = (
                target_cell.paragraphs[0] if p_i == 0 else target_cell.add_paragraph()
            )
            new_paragraph.paragraph_format.space_before = Pt(0)
            new_paragraph.paragraph_format.space_after = Pt(0)
            new_paragraph.alignment = paragraph.alignment
            new_paragraph.paragraph_format.left_indent = (
                paragraph.paragraph_format.left_indent
            )
            for run in paragraph.runs:
                new_run = new_paragraph.add_run(run.text)
                if "prof_educ_logo" in run.text:
                    new_run.text = new_run.text.replace("prof_educ_logo", "")
                    new_run.add_picture("pictures/professional-education-logo.png")
                    continue
                utils.preserve_formatting(new_run, run)

    def _maybe_add_nested_table(self, cell, target_cell):
        if len(cell.tables) > 0:
            last_paragraph = target_cell.paragraphs[-1]
            last_paragraph.paragraph_format.space_after = Pt(0)
            nested_table = cell.tables[0]
            new_table = target_cell.add_table(
                rows=len(nested_table.rows), cols=len(nested_table.columns)
            )
            new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            for r_i, rw in enumerate(nested_table.rows):
                for c_i, cll in enumerate(rw.cells):
                    self._copy_text_and_formatting(cll, new_table.cell(r_i, c_i))

    def _add_table(self, merged_table, curr_row, curr_col, table):
        for row_index, row in enumerate(table.rows):
            merged_table.rows[curr_row].height = Inches(2.76)
            for col_index, cell in enumerate(row.cells):
                target_cell = merged_table.cell(curr_row, curr_col)
                target_cell.width = Inches(3.84)
                self._copy_text_and_formatting(cell, target_cell)
                self._maybe_add_nested_table(cell, target_cell)

    def _add_student_content_to_merged_table(
        self,
        merged_table,
        source_table,
        student_index,
        target_row_index,
        picture_path=None,
        picture_height=None,
        picture_width=None,
        picture_mode="first_cell",
    ):
        if student_index == 0:
            for element_name in ["w:tblGrid", "w:tblPr"]:
                utils.copy_table_element(
                    source_table._tbl, merged_table._tbl, element_name
                )

        for row_index in range(
            len(source_table._tbl.findall("./w:tr", namespaces=source_table._tbl.nsmap))
        ):
            target_row_element = merged_table.rows[target_row_index]._element
            source_row_element = source_table.rows[row_index]._element
            utils.addTrPr(source_row_element, target_row_element)

            source_row_cells = source_row_element.findall(
                "./w:tc", namespaces=source_row_element.nsmap
            )
            for col_index, source_cell in enumerate(source_row_cells):
                target_cell = target_row_element[col_index]
                for child in source_cell:
                    utils.update_nested_table_styles(source_cell, source_row_element)
                    target_cell.append(copy.deepcopy(child))

            for cell in merged_table.rows[target_row_index].cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if "prof_educ_logo" in run.text:
                            run.text = run.text.replace("prof_educ_logo", "")
                            run.add_picture("pictures/professional-education-logo.png")
                        if "bigger_educ_logo" in run.text:
                            run.text = run.text.replace("bigger_educ_logo", "")
                            run.add_picture(
                                "pictures/professional-education-logo.png",
                                width=Inches(1.53),
                                height=Inches(1.09),
                            )
            if picture_path:
                if picture_mode == "all_cells":
                    for cell in merged_table.rows[target_row_index].cells:
                        p = cell.add_paragraph()
                        picture.add_float_picture(
                            p, picture_path, height=picture_height, width=picture_width,
                            pos_x=Pt(0), pos_y=Pt(0)
                        )
                else:
                    p = merged_table.rows[target_row_index].cells[0].add_paragraph()
                    picture.add_float_picture(
                        p, picture_path, height=picture_height, width=picture_width,
                        pos_x=Pt(0), pos_y=Pt(0)
                    )

    # --- Public Methods for Document Generation ---

    def create_beginning_document(self):
        def populator(row, student, index, data):
            row.cells[0].text = str(index + 1)
            row.cells[1].text = student.name
            row.cells[2].text = data["student_company"]

        return self._create_list_based_document(
            "templates/Приказ о начале.docx", populator
        )

    def create_end_doc(self):
        def populator(row, student, index, data):
            row.cells[0].text = str(index + 1)
            row.cells[1].text = student.name
            row.cells[2].text = data["student_company"]
            row.cells[3].text = student.cert_number

        return self._create_list_based_document(
            "templates/Приказ о выпуске.docx", populator
        )

    def create_protocol_doc(self):
        def populator(row, student, index, data):
            row.cells[0].text = str(index + 1)
            row.cells[1].text = student.name
            row.cells[2].text = data["student_company"]
            row.cells[3].text = student.cert_number

        return self._create_list_based_document("templates/Протокол.docx", populator)

    def create_labour_protection_protocol(self):
        def populator(row, student, index, data):
            row.cells[0].text = str(index + 1)
            row.cells[1].text = student.name
            row.cells[2].text = student.role
            row.cells[3].text = data["student_company"]
            row.cells[4].text = ""
            row.cells[5].text = data["end_date"]

        return self._create_list_based_document(
            "templates/protocol_milana.docx", populator
        )

    def create_confirmation_page(self, picture_path):
        table_configs = [
            {
                "cols": 2,
                "picture_path": picture_path,
                "picture_height": Inches(5.54),
                "picture_width": Inches(7.85),
            },
            {
                "cols": 2,
                "picture_path": picture_path,
                "picture_height": Inches(5.54),
                "picture_width": Inches(7.85),
            },
        ]
        return self._create_merged_doc_from_template_rows(
            "templates/milana_conf_page.docx", table_configs
        )

    def create_height_certificate(self):
        table_configs = [{"cols": 3}]
        return self._create_merged_doc_from_template_rows(
            "templates/height_certificate.docx", table_configs
        )

    def create_tractor_certificate(self, picture_front, picture_back):
        table_configs = [
            {
                "cols": 2,
                "picture_path": picture_front,
                "picture_height": TRACTOR_CERT_HEIGHT,
                "picture_width": TRACTOR_CERT_WIDTH,
            },
            {
                "cols": 2,
                "picture_path": picture_back,
                "picture_height": TRACTOR_CERT_HEIGHT,
                "picture_width": TRACTOR_CERT_WIDTH,
            },
        ]
        return self._create_merged_doc_from_template_rows(
            "templates/certificate_tractor.docx", table_configs
        )

    def _copy_table_layout(self, source_table, target_table):
        """Copies column widths and table-level properties from a source to a target table."""
        # Copy column widths
        source_grid = source_table._tbl.find(
            "w:tblGrid", namespaces=source_table._tbl.nsmap
        )
        if source_grid is not None:
            target_table._tbl.replace(
                target_table._tbl.find("w:tblGrid", namespaces=target_table._tbl.nsmap),
                copy.deepcopy(source_grid),
            )
        # Copy table properties (like borders, alignment, etc.)
        source_props = source_table._tbl.find(
            "w:tblPr", namespaces=source_table._tbl.nsmap
        )
        if source_props is not None:
            target_table._tbl.replace(
                target_table._tbl.find("w:tblPr", namespaces=target_table._tbl.nsmap),
                copy.deepcopy(source_props),
            )

    def create_ud(self):
        """
        Generates the 'Удостоверение' document by creating a sequence of tables
        for each student and appending them to a final document.

        This method assumes 'templates/x' contains exactly two tables:
        - Table 1 (e.g., 2 columns, 3 rows)
        - Table 2 (e.g., 2 columns, 4 rows)
        """
        if not self.students:
            return Document()

        # 1. Create the final, empty document that we will add everything to.
        merged_doc = Document()
        utils.set_default_font(merged_doc)

        # 2. Loop through each student to generate their set of tables.
        for index, student in enumerate(self.students):
            # Create a dictionary with this student's specific data.
            local_dict = make_student_copy(self.replacement_dict, student)

            # Render the template with the student's data. This creates an
            # in-memory doc with the two fully-rendered tables.
            template_doc = DocxTemplate("templates/roza_ud.docx")
            template_doc.render(local_dict)

            # 3. Deep-copy each table from the rendered template into the final document.
            #    This is the core logic for appending whole tables.
            for table in template_doc.tables:
                # We append a deep copy of the table's underlying XML element.
                tbl_element = copy.deepcopy(table._tbl)
                merged_doc._body._body.append(tbl_element)

        return merged_doc

    def create_tractor_certs(self):
        blue = self.create_tractor_certificate(
            "pictures/tractor-background-blue.png",
            "pictures/tractor-background-blue-with-tractor.png",
        )

        # For the green certificate, we need a modified dictionary.
        # We create a temporary generator with this new dictionary to keep the logic clean.
        green_dict = self.replacement_dict.copy()
        green_dict["student_profession"] = TRACTOR_PROFESSION_WORDING
        green_generator = DocumentGenerator(green_dict, self.students)
        green = green_generator.create_tractor_certificate(
            "pictures/tractor-background-green.png",
            "pictures/tractor-background-green-with-tractor.png",
        )
        return (blue, green)

    # --- Methods with unique logic, kept as is but moved into the class ---

    # This is kept as a separate function because it tries to fit two students per row.
    def create_certificate_for_labour_protection(self):
        if not self.students:
            return Document()

        merged_doc = Document()
        merged_doc = utils.fit_more_rows(merged_doc)
        utils.set_default_font(merged_doc)

        num_rows = math.ceil(len(self.students) / 2)
        merged_table_front = merged_doc.add_table(rows=num_rows, cols=2)
        merged_table_front.style = "TableGrid"
        merged_doc.add_page_break()
        merged_table_back = merged_doc.add_table(rows=num_rows, cols=2)

        curr_row, curr_col = 0, 0
        for student in self.students:
            local_dict = make_student_copy(self.replacement_dict, student)
            doc = DocxTemplate("templates/labour_protection.docx")
            doc.render(local_dict)

            self._add_table(merged_table_front, curr_row, curr_col, doc.tables[0])
            self._add_table(merged_table_back, curr_row, curr_col, doc.tables[1])

            curr_col += 1
            if curr_col == 2:
                curr_col = 0
                curr_row += 1
        return merged_doc

    def create_certificate(self):
        """
        Generates the 'Свидетельство' document.
        This now uses the standard 'Merge-to-Shell' pattern for consistency.
        """
        # 1. Define the configuration for this specific certificate
        table_configs = [
            {
                "cols": 1,  # The certificate is in a single table column
                "picture_path": "pictures/basic-cert-background.png",
                "picture_height": CERT_HEIGHT_INCHES,
                "picture_width": CERT_WIDTH_INCHES,
            }
        ]

        # 2. Call the standardized helper with the correct template and config
        return self._create_merged_doc_from_template_rows(
            "templates/свидетельство.docx", table_configs
        )

    def create_diploma(self):
        # Your diploma template has ONE table with TWO columns (front and back)
        table_configs = [
            {
                "cols": 2,
                "picture_path": "pictures/basic-cert-background-vert.png",
                "picture_height": Inches(5.49),
                "picture_width": Inches(3.72),
                "picture_mode": "all_cells",
            }
        ]
        # The rest of the function call is the same
        return self._create_merged_doc_from_template_rows(
            "templates/diploma.docx", table_configs
        )


# ==============================================================================
# --- Streamlit UI ---
# ==============================================================================

st.title("Профессиональное обучение")

# --- Input Fields ---
available_professions = utils.load_from_pickle("data/professions.pickle")
student_profession = utils.choose_profession(available_professions)

today = datetime.date.today()
beginning_date = st.date_input("дата начала", value=today)
end_date = st.date_input("дата окончания", value=today)

beginning_number = st.number_input(
    "номер приказа о начале", step=1, value=1, placeholder=808
)
end_number = st.number_input(
    "номер приказа об окончании", step=1, value=1, placeholder=808
)

teacher_name = utils.choose_teacher(utils.load_from_pickle("data/teachers.pickle"))
company = st.text_input(
    "Предприятие", "заявление", placeholder="Наименование предприятия или 'заявление'"
)
student_names = st.text_area("Введите имена студентов, по одному на строку")

# Define column names for clarity. This is a huge advantage.
column_names = [
    "cert_id_raw",
    "date",
    "course",
    "student_name",
    "category_or_student_role",
    "razryad",
]

# Use io.StringIO to let pandas read the string as if it were a file
student_data = []
if student_names:
    data_io = io.StringIO(student_names)

    # Use the powerful pd.read_csv function to parse the data
    # We tell it the separator is a tab ('\t') and there's no header row.
    df = pd.read_csv(
        data_io,
        sep=r"\t+",  # Specify the delimiter is a tab
        header=None,  # The input data has no header row
        names=column_names,  # Assign our defined column names
        engine="python",  # A more robust engine for varied delimiters or formats
        index_col=False,
    )
    print(df)
    try:
        df["cert_number"] = (
            pd.to_numeric(df["cert_id_raw"].astype(str).str.strip("."), errors="coerce")
            .fillna(0)
            .astype(int)
        )
        parsed_info = df.apply(
            lambda row: utils.parse_machine_cat_or_role(
                student_profession, row["category_or_student_role"]
            ),
            axis=1,
            result_type="expand",  # This splits the tuple result into two new columns
        )
        df[["machine_category", "role"]] = parsed_info
        student_data = [
            utils.Student(
                name=row.student_name,
                cert_number=row.cert_number,
                machine_category=row.machine_category,
                role=row.role,
                razryad=row.razryad,
            )
            for row in df.itertuples()
        ]
        print(student_data)
    except Exception as e:
        print(e)
        st.warning(
            "Please use an integer as the certificate number. It is a required field."
        )


replacement_dict = {
    "beginning_date": utils.format_date(beginning_date),
    "beginning_number": beginning_number,
    "end_date": utils.format_date(end_date),
    "end_number": end_number,
    "student_company": company,
    "teacher_name": teacher_name,
    "num_students": len(student_data),
    "year": end_date.year,
    "expiration_date": utils.format_date((end_date + relativedelta(years=3))),
}
if student_profession:
    if student_profession.hours_str:
        replacement_dict["hours"] = student_profession.hours_str
    if student_profession.formatted_profession:
        replacement_dict["student_profession"] = student_profession.formatted_profession

generator = DocumentGenerator(replacement_dict, student_data)

# --- Step 1: Define all possible document choices ---
# This maps a user-friendly name to the function that creates it.
# This makes the code much cleaner and easier to maintain.
doc_options = {
    "Приказ о начале": generator.create_beginning_document,
    "Приказ о выпуске": generator.create_end_doc,
    "Протокол": generator.create_protocol_doc,
    "Свидетельство": generator.create_certificate,
    "Свидетельство тракторов (синее)": lambda: generator.create_tractor_certs()[0],
    "Свидетельство тракторов (зеленое)": lambda: generator.create_tractor_certs()[1],
    "Милана (удостоверение)": lambda: generator.create_confirmation_page(
        "pictures/tractor-background-green.png"
    ),
    "Милана (св-во охрана труда)": generator.create_certificate_for_labour_protection,
    "Милана (протокол охрана труда)": generator.create_labour_protection_protocol,
    "На высоте": generator.create_height_certificate,
    "Удостоверение Роза": generator.create_ud,
    "Диплом": generator.create_diploma,
}

# --- Step 2: Create the Checkbox UI ---
st.subheader("Выберите документы для генерации и скачивания:")

# Use columns for a neater layout
cols = st.columns(3)
user_selections = {}
# Create a checkbox for each document option, defaulting to True (selected)
for i, name in enumerate(doc_options.keys()):
    with cols[i % 3]:
        user_selections[name] = st.checkbox(name, value=False)

# --- Step 3: The Main "Generate and Download" Button ---
if st.button("Сгенерировать и подготовить к скачиванию"):
    # Input validation
    if not all([student_profession, teacher_name, student_data]):
        st.warning("Пожалуйста, заполните все поля и добавьте хотя бы одного студента.")
    else:
        # This dictionary will hold the documents that are actually generated.
        docs_to_zip = {}
        # A placeholder to show generation progress
        progress_bar = st.progress(0, "Начинаем генерацию...")

        selected_docs = [name for name, selected in user_selections.items() if selected]
        total_docs = len(selected_docs)

        # Loop through the user's selections and generate only the chosen documents
        for i, name in enumerate(selected_docs):
            if user_selections[name]:  # If the box is checked
                progress_text = f"Генерация: {name} ({i+1}/{total_docs})"
                st.write(progress_text)
                progress_bar.progress((i + 1) / total_docs, text=progress_text)

                # Look up the correct function from our options and call it
                generator_func = doc_options[name]
                generated_doc = generator_func()

                # Add the generated document to our dictionary for zipping
                docs_to_zip[f"{name}.docx"] = generated_doc

        progress_bar.empty()

        # --- Step 4: Create the ZIP archive and Download Button ---
        if not docs_to_zip:
            st.warning("Вы не выбрали ни одного документа для генерации.")
        else:
            st.success("Все выбранные документы успешно сгенерированы!")

            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zipf:
                for filename, doc in docs_to_zip.items():
                    # Save each doc into the in-memory zip file
                    with zipf.open(filename, "w") as f:
                        doc.save(f)

            zip_buffer.seek(0)

            # Display the download button for the created ZIP file
            st.download_button(
                label="✅ Скачать документы (ZIP)",
                data=zip_buffer,
                file_name=f"{end_date.strftime('%d.%m.%Y')}.zip",
                mime="application/zip",
            )

# Date formatting
import streamlit as st
import pandas as pd
import pickle
import copy
import docx
from docx.shared import Inches, Pt
from docx.oxml import OxmlElement



# Russian month names mapping
russian_months = {
    "January": "января",
    "February": "февраля",
    "March": "марта",
    "April": "апреля",
    "May": "мая",
    "June": "июня",
    "July": "июля",
    "August": "августа",
    "September": "сентября",
    "October": "октября",
    "November": "ноября",
    "December": "декабря"
}

def format_date(date): 
    formatted_date = date.strftime("%d %B %Y") + " г."
    for en_month, ru_month in russian_months.items():
        formatted_date = formatted_date.replace(en_month, ru_month)
    return formatted_date

def display_docx_content(doc):
    """Reads a .docx file and displays paragraphs and tables in order."""
    for element in doc.element.body:
        if isinstance(element, docx.oxml.text.paragraph.CT_P):
            st.write(element.text)
        elif isinstance(element, docx.oxml.table.CT_Tbl):
            # Display table (you'll need to customize formatting)
            table = docx.table.Table(element, doc)  # Create Table object
            data = [[cell.text for cell in row.cells] for row in table.rows]
            df = pd.DataFrame(data)
            st.table(df)


def save_data(dict, filename):
    with open(filename, "wb") as f:
        pickle.dump(dict, f)

def load_from_pickle(filename):
    """Loads a dictionary from a pickle file.
    Returns an empty dictionary if the file doesn't exist.
    """
    try:
        with open(filename, "rb") as f:
            return pickle.load(f)
    except FileNotFoundError:
        return {}

# This function is used to print the output of the xml. Example.
# from lxml import etree
# xml = etree.tostring(tbl._tbl, encoding='unicode', pretty_print=True)
# dump(xml, 'og table')
def dump(string_to_dump, name): 
    file_path = f"{name}.txt"

    with open(file_path, 'w') as file:
        file.write(string_to_dump)

### Table formatting
def preserve_formatting(new_run, source_run): 
    new_run.font.name = source_run.font.name
    new_run.font.size = source_run.font.size
    new_run.font.bold = source_run.font.bold
    new_run.font.italic = source_run.font.italic
    new_run.font.underline = source_run.font.underline
    new_run.font.color.rgb = source_run.font.color.rgb
    return new_run

# https://github.com/python-openxml/python-docx/issues/205
def copy_cell_properties(source_cell, dest_cell):
    '''Copies cell properties from source cell to destination cell.

    Copies cell background shading, borders etc. in a python-docx Document.

    Args:
        source_cell (docx.table._Cell): the source cell with desired formatting
        dest_cell (docx.table._Cell): the destination cell to which to apply formatting
    '''
    # get properties from source cell
    # (tcPr = table cell properties)
    cell_properties = source_cell._tc.get_or_add_tcPr()
    
    # remove any table cell properties from destination cell
    # (otherwise, we end up with two <w:tcPr> ... </w:tcPr> blocks)
    dest_cell._tc.remove(dest_cell._tc.get_or_add_tcPr())

    # make a shallow copy of the source cell properties
    # (otherwise, properties get *moved* from source to destination)
    cell_properties = copy.copy(cell_properties)

    # append the copy of properties from the source cell, to the destination cell
    dest_cell._tc.append(cell_properties)

def addTrPr(source_row_element, target_row):
    source_trPr = source_row_element.find('w:trPr', namespaces=source_row_element.nsmap)
    if source_trPr is not None:
        # Example: Copy only the row height
        tr_height = source_trPr.find('./w:trHeight', namespaces=source_trPr.nsmap)
        if tr_height is not None:
            target_row.append(copy.deepcopy(tr_height))  
        target_row.append(OxmlElement('w:cantSplit'))

def copy_table_element(source_tbl, target_tbl, element_name):
    """Copies a specified table element from source to target table."""
    source_element = source_tbl.find(element_name, namespaces=source_tbl.nsmap)
    if source_element is not None:
        target_element = target_tbl.find(element_name, namespaces=target_tbl.nsmap)
        if target_element is not None:
            target_tbl.remove(target_element)
        target_tbl.insert(0, copy.deepcopy(source_element)) 

def update_nested_table_styles(source_cell, source_row_element):
    """Updates line spacing after to 0 for all paragraphs in a nested table."""

    nested_table = source_cell.find('.//w:tbl', namespaces=source_row_element.nsmap)
    if nested_table: 
        # Find all paragraphs within the nested table
        for paragraph in nested_table.findall('.//w:p', namespaces=source_row_element.nsmap):
            pPr = paragraph.find('./w:pPr', namespaces=source_row_element.nsmap)

            if pPr is not None:
                spacing = pPr.find('./w:spacing', namespaces=source_row_element.nsmap)
                if spacing is None:
                    spacing = OxmlElement('w:spacing')
                    pPr.append(spacing)

                # Set line spacing after to 0
                spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}after', '0') 

def set_default_font(final_doc, bold=False): 
    style = final_doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    if bold: 
        font.bold = True


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
        paragraph.paragraph_format.widow_control = (
            True  # Prevents single lines from appearing at the top or bottom of a page
        )
        paragraph.paragraph_format.keep_lines_together = (
            True  # Keep all lines of a paragraph together on a page
        )

    return document


def choose_teacher(all_teachers):
    """Handles teacher selection and adding new teachers."""

    selected_teacher = st.selectbox(
        "Выберите преподавателя из списка:",
        all_teachers,
        index=None,
        placeholder="Преподаватели",
    )

    if selected_teacher:
        st.write(f"Преподаватель: {selected_teacher}")

    add_new = st.checkbox("Добавить нового преподавателя?")

    # Initialize new_teacher in session state
    if "new_teacher" not in st.session_state:
        st.session_state.new_teacher = None

    if add_new:
        new_teacher = st.text_input("Введите инициалы и фамилию нового преподавателя:")
        if st.button("Добавить преподавателя"):
            if new_teacher and new_teacher not in all_teachers:
                all_teachers.append(new_teacher)
                save_data(all_teachers, "data/teachers.pickle")
                st.success(f"Преподаватель '{new_teacher}' добавлен!")
                # Store new_teacher in session state
                st.session_state.new_teacher = new_teacher
            else:
                st.warning("Преподаватель уже существует или не введен.")

    # Return the teacher based on session state and selection
    if st.session_state.new_teacher:
        return st.session_state.new_teacher
    elif selected_teacher:
        return selected_teacher
    else:
        return None


def choose_profession(all_professions):
    """Handles profession selection and adding new professions."""

    selected_item = st.selectbox(
        "Выберите профессию/программу обучение из следующих опций:",
        all_professions,
        index=None,
        placeholder="Начинайте вводить название программы",
    )

    if selected_item:
        st.write(f"Программа обучения: {selected_item}")

    add_new = st.checkbox("Добавить новую программу обучения?")

    if "new_profession" not in st.session_state:
        st.session_state.new_profession = None

    if add_new:
        new_profession = st.text_input("Введите название новой программы:")
        new_code = st.text_input("Введите код:")
        if st.button("Добавить программу"):
            if new_profession:
                code_int = -1
                try:
                    code_int = int(new_code)
                    all_professions[new_profession] = [code_int]
                except Exception as e:
                    print("Failed to parse the profession code")
                    all_professions[new_profession] = [-1]
                save_data(all_professions, filename="data/professions.pickle")
                st.success(f"Программа '{new_profession}' добавлена!")
                st.session_state.new_profession = new_profession
            else:
                st.warning("Программа не введена.")

    if st.session_state.new_profession:
        code_int = all_professions[st.session_state.new_profession][0]
        if code_int == -1:
            return f"«{st.session_state.new_profession}»"
        return f"{code_int} «{st.session_state.new_profession}»"
    elif selected_item:
        selected_profession_code = all_professions[selected_item]
        selected_profession_code_str = ", ".join(
            str(code) for code in selected_profession_code
        )
        return f"{selected_profession_code_str} «{selected_item}»"
    else:
        return None

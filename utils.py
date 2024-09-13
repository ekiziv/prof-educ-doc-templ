# Date formatting
import streamlit as st
import pandas as pd
import pickle
import copy
from docx.shared import Pt


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

### Table formatting
def preserve_formatting(new_run, source_run): 
    new_run.font.name = source_run.font.name
    new_run.font.size = source_run.font.size
    new_run.font.bold = source_run.font.bold
    new_run.font.italic = source_run.font.italic
    new_run.font.underline = source_run.font.underline
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

def set_default_font(final_doc, bold=False): 
    style = final_doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    if bold: 
        font.bold = True
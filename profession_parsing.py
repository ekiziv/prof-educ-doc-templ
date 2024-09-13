from docx import Document
import streamlit as st
import pickle

def professions_docx_table_to_df(docx_path='/Users/ekiziv/Desktop/mama/work/data/all_professions_cleaned.docx'):
    """Loads a table from a .docx file into a dictionary and saves to pickle. 
    Handles multiple codes per profession.

    Args:
        docx_path (str, optional): Path to the .docx file. 
    Returns:
        dict: Dictionary containing the table data (profession: [list of codes]).
    """

    doc = Document(docx_path)
    table = doc.tables[0]
    professions = {}

    for row in table.rows:
        row_data = [cell.text.strip() for cell in row.cells]
        profession = row_data[0]
        code = row_data[1]

        if not profession:  # Skip empty rows
            continue

        try:
            code = int(code)  # Try converting to an integer

        except ValueError:
            if code == "-":
                code = None  # Or use your preferred default for missing codes
            else:
                # Handle cases with multiple codes (assuming comma-separated)
                
                codes = [int(c.strip()) for c in code.split(",") if c.strip()]
                if codes:
                    professions.setdefault(profession, []).extend(codes)
                continue  # Go to the next row after handling multiple codes

        # Add the code (single or from multiple codes handling)
        professions.setdefault(profession, []).append(code)

    # Remove duplicate codes for each profession (if any)
    for key in professions:
        professions[key] = list(set(professions[key]))

    # sort dictionary by key
    professions = dict(sorted(professions.items()))

    st.write(professions)

    with open("data/professions.pickle", "wb") as f:
        pickle.dump(professions, f)

    return professions

def teachers_to_pickle():
    teachers = ["А.И. Мамонтов", 
                "А.В. Перекрестов", 
                "Н.В. Клюшина",
                "Л.А. Лапчук"]
    with open("data/teachers.pickle", "wb") as f:
        pickle.dump(teachers, f)

if __name__ == "__main__":
    professions_docx_table_to_df()
    teachers_to_pickle()
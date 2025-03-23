import streamlit as st
from docx import Document

# Streamlit UI
st.title("Lab Results Entry System")

# Input Fields
lab_number = st.text_input("Lab Number")
sample_number = st.text_input("Sample Number")
sample_description = st.text_area("Sample Description")

st.subheader("Summary of Results")

# Predefined Parameters List
predefined_parameters = [
    "Moisture",
    "Ash",
    "Crude Protein",
    "Total Fat",
    "Sugar",
    "Sodium",
    "Potassium",
    "Calcium",
    "pH",
    "Water Activity",
    "Total Titrable Activity",
    "Carbohydrates",
    "Food Energy Value",
]

# Let users select multiple parameters
selected_parameters = st.multiselect("Select Parameters", predefined_parameters)

parameters = []
for param in selected_parameters:
    parameters.append(
        {
            "Parameter": param,
            "Date Started": st.date_input(f"Date Started ({param})"),
            "Environmental Conditions": st.text_input(
                f"Environmental Conditions ({param})"
            ),
            "Method Used": st.text_input(f"Method Used ({param})"),
            "Results": st.text_input(f"Results ({param})"),
            "Uncertainty": st.text_input(f"Uncertainty ({param})"),
            "Unit": st.text_input(f"Unit ({param})"),
        }
    )


# Function to generate the report
def generate_report():
    doc = Document()
    doc.add_heading("Lab Results Report", level=1)

    doc.add_paragraph(f"Lab Number: {lab_number}")
    doc.add_paragraph(f"Sample Number: {sample_number}")
    doc.add_paragraph(f"Sample Description: {sample_description}")
    doc.add_paragraph("\nSummary of Results:")

    table = doc.add_table(rows=1, cols=7)
    hdr_cells = table.rows[0].cells
    headers = [
        "Parameter",
        "Date Started",
        "Environmental Conditions",
        "Method Used",
        "Results",
        "Uncertainty",
        "Unit",
    ]

    for i, header in enumerate(headers):
        hdr_cells[i].text = header

    for param in parameters:
        row_cells = table.add_row().cells
        row_cells[0].text = param["Parameter"]
        row_cells[1].text = str(param["Date Started"])
        row_cells[2].text = param["Environmental Conditions"]
        row_cells[3].text = param["Method Used"]
        row_cells[4].text = param["Results"]
        row_cells[5].text = param["Uncertainty"]
        row_cells[6].text = param["Unit"]

    report_filename = "Lab_Report.docx"
    doc.save(report_filename)

    return report_filename


# Generate Report Button
if st.button("Generate Report"):
    file_path = generate_report()
    st.success("Report generated successfully!")
    with open(file_path, "rb") as file:
        st.download_button(
            "Download Report",
            file,
            file_path,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

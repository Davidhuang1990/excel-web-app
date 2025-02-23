import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io

st.title("GreenAccuracy Editor")

# Define the mappings based on the template
VALIDATION_MAPPINGS = {
    "Plastic": {
        "Packaging Form": [
            "Beverage bottle", "Carrier bag", "Disposables (bowls, boxes, covers, cups, plates, trays)",
            "Product packaging (flexible)", "Product packaging (rigid, excluding beverage bottle)",
            "Transport and protective packaging", "Others", "Crate"
        ],
        "Further Details": ["EPS", "HDPE", "LDPE", "PET", "PP", "PS", "PVC", "Others"]
    },
    "Paper": {
        "Packaging Form": [
            "Carrier bag", "Disposables (bowls, boxes, cups, plates, trays)",
            "Product packaging", "Transport and protective packaging", "Others"
        ],
        "Further Details": ["Corrugated board", "Paper", "Paperboard", "Others"]
    },
    "Metal": {
        "Packaging Form": [
            "Beverage can", "Disposables (foils, trays)", "Food can",
            "Product packaging (excluding beverage can, food can)", "Others"
        ],
        "Further Details": ["Aluminium", "Steel", "Tin", "Others"]
    },
    "Glass": {
        "Packaging Form": ["Beverage bottle", "Product packaging (excluding beverage bottle)", "Others"],
        "Further Details": ["Brown", "Clear", "Green", "Other colours"]
    },
    "Wood": {
        "Packaging Form": ["Crate", "Pallet", "Product packaging", "Transport and protective packaging", "Others"],
        "Further Details": ["N/A"]  # Allow "N/A" as a valid Further Details for Wood
    },
    "Composite": {
        "Packaging Form": ["Beverage carton", "Packs / Sachets", "Product packaging (excluding beverage carton, packs/sachets)", "Others"],
        "Further Details": []  # No specific further details provided
    },
    "Others": {
        "Packaging Form": ["Carrier bag", "Others"],
        "Further Details": ["Biodegradable/compostable", "Oxo-degradable/oxo-biodegradable", "Others"]
    }
}

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])
if uploaded_file is not None:
    # Load the workbook and sheets using openpyxl
    wb = load_workbook(io.BytesIO(uploaded_file.getvalue()))
    packaging_sheet = wb["Packaging Data"]

    # Identify the header row (0-indexed) for the data table
    header_row = 6  # Row 7 (1-based) has headers, data starts at row 8
    headers = [cell.value.strip() if isinstance(cell.value, str) else cell.value for cell in packaging_sheet[header_row + 1]]

    # Extract data into DataFrame
    data = []
    for row in packaging_sheet.iter_rows(min_row=header_row + 2, values_only=True):
        data.append(row)
    df = pd.DataFrame(data, columns=headers)

    st.write("Original Data:")
    st.dataframe(df)

    # Step 2: Edit the Data
    st.write("Edit the data below:")
    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True
    )

    # Step 3: Validation Logic
    def validate_data(df):
        errors = []
        # Expected column names (case-insensitive and stripped)
        expected_columns = ["packaging material", "packaging form", "further details", "weight (kg)"]
        actual_columns = [col.lower().strip() if isinstance(col, str) else col for col in df.columns]

        # Map actual columns to expected ones (fuzzy match)
        column_mapping = {}
        for expected in expected_columns:
            for actual in actual_columns:
                if expected in actual:
                    column_mapping[expected] = df.columns[actual_columns.index(actual)]
                    break
            if expected not in column_mapping:
                errors.append(f"Missing required column: {expected.capitalize()}")
                return errors

        # Use mapped column names
        material_col = column_mapping["packaging material"]
        form_col = column_mapping["packaging form"]
        details_col = column_mapping["further details"]
        weight_col = column_mapping["weight (kg)"]

        # Validate only non-empty rows
        for idx, row in df.iterrows():
            material = row[material_col]
            form = row[form_col]
            details = row[details_col]
            weight = row[weight_col]

            # Skip if row is completely empty (all required fields NaN)
            if all(pd.isna(row[col]) for col in [material_col, form_col, details_col, weight_col]):
                continue

            # Validate Weight (if provided)
            if pd.notna(weight):
                if isinstance(weight, (int, float)) and weight < 0:
                    errors.append(f"Row {idx + 1}: Weight must be a positive number.")
                elif not isinstance(weight, (int, float)):
                    errors.append(f"Row {idx + 1}: Weight must be a number, got '{weight}'.")

            # If material is missing but other fields are present, flag error
            if pd.isna(material) and any(pd.notna(row[col]) for col in [form_col, details_col, weight_col]):
                errors.append(f"Row {idx + 1}: Packaging Material is required when other fields are provided.")
                continue

            # Skip further validation if no material
            if pd.isna(material):
                continue

            # Check if material is valid
            if material not in VALIDATION_MAPPINGS:
                errors.append(f"Row {idx + 1}: Invalid Packaging Material '{material}'.")
                continue

            # Normalize form and details by stripping whitespace
            form = form.strip() if isinstance(form, str) else form
            details = details.strip() if isinstance(details, str) else details

            # Validate Packaging Form
            valid_forms = VALIDATION_MAPPINGS[material]["Packaging Form"]
            if pd.notna(form) and form not in valid_forms:
                errors.append(f"Row {idx + 1}: '{form}' is not a valid Packaging Form for '{material}'. Valid options: {', '.join(valid_forms)}")

            # Validate Further Details (if provided)
            valid_details = VALIDATION_MAPPINGS[material]["Further Details"]
            if pd.notna(details):
                if material == "Wood" and details == "N/A":
                    continue  # Explicitly allow "N/A" for Wood
                if not valid_details:  # No further details expected (except Wood's "N/A")
                    errors.append(f"Row {idx + 1}: No Further Details should be specified for '{material}' (except 'N/A' for Wood).")
                elif details not in valid_details:
                    errors.append(f"Row {idx + 1}: '{details}' is not a valid Further Detail for '{material}'. Valid options: {', '.join(valid_details)}")

        return errors

    # Perform validation
    errors = validate_data(edited_df)
    if errors:
        st.error("Validation Errors Found:")
        for error in errors:
            st.write(f"- {error}")
        st.warning("You can still save the file, but it contains validation errors.")
    else:
        st.success("Data is valid!")

    # Step 4: Save the Edited Data (always available)
    if st.button("Save to Excel"):
        # Clear existing data below headers
        for row in packaging_sheet.iter_rows(min_row=header_row + 2, max_row=packaging_sheet.max_row, max_col=len(headers)):
            for cell in row:
                cell.value = None

        # Write new data
        for index, row in edited_df.iterrows():
            for col_idx, value in enumerate(row):
                packaging_sheet.cell(row=header_row + index + 2, column=col_idx + 1, value=value)

        # Save to BytesIO
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.download_button(
            label="Download Updated Excel",
            data=output,
            file_name="updated_file.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.write("Instructions: Upload an Excel file, edit the data, and click 'Save to Excel' to download the updated file.")

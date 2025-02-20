import streamlit as st
import pandas as pd
import io

# Title of the app
st.title("GreenAccuracy Editor")

# Step 1: File Upload
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

# Initialize session state to store the DataFrame
if "df" not in st.session_state:
    st.session_state.df = None

if uploaded_file is not None:
    # Read the Excel file into a DataFrame and store it in session state
    st.session_state.df = pd.read_excel(uploaded_file)
    st.write("Original Data:")
    st.dataframe(st.session_state.df)

# Step 2: Edit the Data (only show editor if data exists)
if st.session_state.df is not None:
    st.write("Edit the data below:")
    edited_df = st.data_editor(
        st.session_state.df,
        num_rows="dynamic",  # Allows adding/removing rows
        use_container_width=True,
        key="data_editor"  # Unique key to ensure proper rendering
    )

    # Step 3: Validation Logic
    def validate_data(df):
        errors = []
        for index, row in df.iterrows():
            for col in df.columns:
                if pd.api.types.is_numeric_dtype(df[col]):
                    if pd.isna(row[col]) or row[col] < 0:
                        errors.append(f"Row {index + 2}, Column '{col}': Value must be a positive number.")
        return errors

    # Perform validation on the edited data
    errors = validate_data(edited_df)
    if errors:
        st.error("Validation Errors Found:")
        for error in errors:
            st.write(f"- {error}")
    else:
        st.success("Data is valid!")

    # Step 4: Save the Edited Data (Always show save button if data exists)
    if st.button("Save to Excel"):
        # Convert DataFrame to Excel in memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            edited_df.to_excel(writer, index=False)
        output.seek(0)

        # Provide download button
        st.download_button(
            label="Download Updated Excel",
            data=output,
            file_name="updated_file.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Instructions
st.write("Instructions: Upload an Excel file, edit the data, and click 'Save to Excel' to download the updated file.")

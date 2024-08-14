import pandas as pd
import streamlit as st
import io

def merge_files(files):
    """Merge all uploaded Excel files into a single DataFrame."""
    all_sheets = {}

    for file in files:
        with pd.ExcelFile(file) as xl:
            for sheet_name in xl.sheet_names:
                if sheet_name not in all_sheets:
                    all_sheets[sheet_name] = []
                df = xl.parse(sheet_name)
                all_sheets[sheet_name].append(df)
    
    # Combine all sheets and return as a dictionary of DataFrames
    merged_sheets = {}
    for sheet_name, dataframes in all_sheets.items():
        merged_df = pd.concat(dataframes, ignore_index=True)
        merged_sheets[sheet_name] = merged_df

    return merged_sheets

def main():
    # Set custom title with green color and add background image
    st.markdown(
        """
        <style>
        .stApp {
            background: rgba(255, 255, 255, 0.5) url("https://static.vecteezy.com/system/resources/thumbnails/033/535/363/small/broken-glass-animation-green-screen-free-video.jpg") no-repeat center center;
            background-size: cover;
        }
        .stTitle {
            color: white;
        }
        .css-1p7i8jb {
            background-color: white !important;
            border: 1px solid #d3d3d3; /* Optional: Add a border for better visibility */
            border-radius: 5px; /* Optional: Rounded corners */
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    # Display the title
    st.markdown('<h1 class="stTitle">Excel Files Merger</h1>', unsafe_allow_html=True)
    
    # Upload multiple files
    files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

    if files:
        output_file = st.text_input("Output File Name (including .xlsx extension):", "merged_files.xlsx")
        if st.button("Merge Files"):
            if not output_file.endswith('.xlsx'):
                st.error("Please provide an output file name with .xlsx extension.")
            else:
                try:
                    # Merge files
                    merged_sheets = merge_files(files)

                    # Save the merged data to an Excel file in-memory
                    output_buffer = io.BytesIO()
                    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                        for sheet_name, df in merged_sheets.items():
                            df.to_excel(writer, sheet_name=sheet_name, index=False)

                    # Provide download link for the merged file
                    output_buffer.seek(0)
                    st.download_button(
                        label="Download Merged Excel File",
                        data=output_buffer,
                        file_name=output_file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                except Exception as e:
                    st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()

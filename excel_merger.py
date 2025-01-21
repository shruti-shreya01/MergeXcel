# import pandas as pd
# import streamlit as st
# import io


# def merge_files(files):
#     """Merge all uploaded Excel files into a single DataFrame."""
#     all_sheets = {}

#     for file in files:
#         with pd.ExcelFile(file) as xl:
#             for sheet_name in xl.sheet_names:
#                 if sheet_name not in all_sheets:
#                     all_sheets[sheet_name] = []
#                 df = xl.parse(sheet_name)
#                 all_sheets[sheet_name].append(df)
    
#     # Combine all sheets and return as a dictionary of DataFrames
#     merged_sheets = {}
#     for sheet_name, dataframes in all_sheets.items():
#         merged_df = pd.concat(dataframes, ignore_index=True)
#         merged_sheets[sheet_name] = merged_df

#     return merged_sheets

# def main():
#     # Set custom title with green color and add background image
#     st.markdown(
#         """
#         <style>
#         .stApp {
#             background: rgba(255, 255, 255, 0.5) url("https://static.vecteezy.com/system/resources/thumbnails/033/535/363/small/broken-glass-animation-green-screen-free-video.jpg") no-repeat center center;
#             background-size: cover;
#         }
#         .stTitle {
#             color: white;
#         }
#         .css-1p7i8jb {
#             background-color: white !important;
#             border: 1px solid #d3d3d3; /* Optional: Add a border for better visibility */
#             border-radius: 5px; /* Optional: Rounded corners */
#         }
#         </style>
#         """,
#         unsafe_allow_html=True
#     )

#     # Display the title
#     st.markdown('<h1 class="stTitle">MergeXcel</h1>', unsafe_allow_html=True)
    
#     # Upload multiple files
#     files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

#     if files:
#         output_file = st.text_input("Output File Name (including .xlsx extension):", "merged_files.xlsx")
#         if st.button("Merge Files"):
#             if not output_file.endswith('.xlsx'):
#                 st.error("Please provide an output file name with .xlsx extension.")
#             else:
#                 try:
#                     # Merge files
#                     merged_sheets = merge_files(files)

#                     # Save the merged data to an Excel file in-memory
#                     output_buffer = io.BytesIO()
#                     with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
#                         for sheet_name, df in merged_sheets.items():
#                             df.to_excel(writer, sheet_name=sheet_name, index=False)

#                     # Provide download link for the merged file
#                     output_buffer.seek(0)
#                     st.download_button(
#                         label="Download Merged Excel File",
#                         data=output_buffer,
#                         file_name=output_file,
#                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#                     )
                    
#                 except Exception as e:
#                     st.error(f"An error occurred: {e}")

# if __name__ == "__main__":
#     main()






import pandas as pd
import streamlit as st
import io
import zipfile
import os
from tempfile import TemporaryDirectory


def merge_files_from_zip(zip_file):
    """Merge all Excel files from the folder within the ZIP archive into a single DataFrame."""
    all_sheets = {}

    with TemporaryDirectory() as temp_dir:
        # Extract ZIP file to the temporary directory
        with zipfile.ZipFile(zip_file) as z:
            # List all files and folders in the ZIP archive
            zip_file_names = z.namelist()
            
            # Find the folder inside the ZIP file (assuming there's one folder, modify if needed)
            folder_name = None
            for name in zip_file_names:
                if name.endswith('/'):  # Identify folder by checking for trailing slash
                    folder_name = name
                    break
            
            # Extract all files from the folder
            if folder_name:
                for file_name in zip_file_names:
                    if file_name.startswith(folder_name) and file_name.endswith('.xlsx'):
                        z.extract(file_name, temp_dir)
            
            # List all extracted Excel files
            excel_files = [
                os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if f.endswith('.xlsx')
            ]

            if not excel_files:
                raise ValueError(f"No Excel files found in the folder '{folder_name}' inside the ZIP archive.")

        # Process each Excel file
        for file in excel_files:
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
    st.markdown('<h1 class="stTitle">MergeXcel</h1>', unsafe_allow_html=True)

    # Upload a ZIP file
    zip_file = st.file_uploader("Upload a ZIP file containing a folder with Excel files", type=["zip"])

    if zip_file:
        output_file = st.text_input("Output File Name (including .xlsx extension):", "merged_files.xlsx")
        if st.button("Merge Files"):
            if not output_file.endswith('.xlsx'):
                st.error("Please provide an output file name with .xlsx extension.")
            else:
                try:
                    # Merge files from ZIP
                    merged_sheets = merge_files_from_zip(zip_file)

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

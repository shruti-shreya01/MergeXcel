# import pandas as pd
# import streamlit as st
# import io
# import os

# def read_file(file):
#     """Read Excel or CSV file and return a dictionary of DataFrames."""
#     file_extension = os.path.splitext(file.name)[1].lower()
    
#     if file_extension == '.xlsx':
#         return pd.read_excel(file, sheet_name=None)
#     elif file_extension == '.csv':
#         return {'Sheet1': pd.read_csv(file)}
#     else:
#         raise ValueError(f"Unsupported file format: {file_extension}")

# def merge_files(files):
#     """Merge all uploaded Excel and CSV files into a single DataFrame."""
#     all_sheets = {}

#     for file in files:
#         sheets = read_file(file)
#         for sheet_name, df in sheets.items():
#             if sheet_name not in all_sheets:
#                 all_sheets[sheet_name] = []
#             all_sheets[sheet_name].append(df)
    
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
#     st.markdown('<h1 class="stTitle">MergeXcel & CSV</h1>', unsafe_allow_html=True)
    
#     # Upload multiple files
#     files = st.file_uploader("Upload Excel or CSV files", type=["xlsx", "csv"], accept_multiple_files=True)

#     if files:
#         output_format = st.radio("Select output format:", ("Excel (.xlsx)", "CSV (.csv)"))
#         output_extension = ".xlsx" if output_format == "Excel (.xlsx)" else ".csv"
#         output_file = st.text_input("Output File Name (including extension):", f"merged_files{output_extension}")
        
#         if st.button("Merge Files"):
#             if not output_file.endswith(output_extension):
#                 st.error(f"Please provide an output file name with {output_extension} extension.")
#             else:
#                 try:
#                     # Merge files
#                     merged_sheets = merge_files(files)

#                     if output_extension == ".xlsx":
#                         # Save the merged data to an Excel file in-memory
#                         output_buffer = io.BytesIO()
#                         with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
#                             for sheet_name, df in merged_sheets.items():
#                                 df.to_excel(writer, sheet_name=sheet_name, index=False)
#                         mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#                     else:
#                         # Save the merged data to a CSV file in-memory
#                         output_buffer = io.StringIO()
#                         merged_sheets['Sheet1'].to_csv(output_buffer, index=False)
#                         output_buffer = io.BytesIO(output_buffer.getvalue().encode())
#                         mime_type = "text/csv"

#                     # Provide download link for the merged file
#                     output_buffer.seek(0)
#                     st.download_button(
#                         label=f"Download Merged {output_format.split()[0]} File",
#                         data=output_buffer,
#                         file_name=output_file,
#                         mime=mime_type
#                     )
                    
#                 except Exception as e:
#                     st.error(f"An error occurred: {e}")

# if __name__ == "__main__":
#     main()





import pandas as pd
import streamlit as st
import io
import os

def read_file(file):
    """Read Excel or CSV file and return a dictionary of DataFrames."""
    file_extension = os.path.splitext(file.name)[1].lower()
    
    if file_extension == '.xlsx':
        return pd.read_excel(file, sheet_name=None, parse_dates=['Date'])
    elif file_extension == '.csv':
        return {'Sheet1': pd.read_csv(file, parse_dates=['Date'])}
    else:
        raise ValueError(f"Unsupported file format: {file_extension}")

def merge_files(files):
    """Merge all uploaded Excel and CSV files based on the Date column."""
    all_sheets = {}

    for file in files:
        sheets = read_file(file)
        for sheet_name, df in sheets.items():
            if 'Date' not in df.columns:
                raise ValueError(f"File {file.name}, sheet {sheet_name} is missing the 'Date' column.")
            
            df.set_index('Date', inplace=True)
            
            if sheet_name not in all_sheets:
                all_sheets[sheet_name] = df
            else:
                all_sheets[sheet_name] = all_sheets[sheet_name].join(df, how='outer', rsuffix=f'_{file.name}')

    # Reset index to make 'Date' a column again
    for sheet_name in all_sheets:
        all_sheets[sheet_name].reset_index(inplace=True)

    return all_sheets

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
    st.markdown('<h1 class="stTitle">MergeXcel & CSV</h1>', unsafe_allow_html=True)
    
    # Upload multiple files
    files = st.file_uploader("Upload Excel or CSV files", type=["xlsx", "csv"], accept_multiple_files=True)

    if files:
        output_format = st.radio("Select output format:", ("Excel (.xlsx)", "CSV (.csv)"))
        output_extension = ".xlsx" if output_format == "Excel (.xlsx)" else ".csv"
        output_file = st.text_input("Output File Name (including extension):", f"merged_files{output_extension}")
        
        if st.button("Merge Files"):
            if not output_file.endswith(output_extension):
                st.error(f"Please provide an output file name with {output_extension} extension.")
            else:
                try:
                    # Merge files
                    merged_sheets = merge_files(files)

                    if output_extension == ".xlsx":
                        # Save the merged data to an Excel file in-memory
                        output_buffer = io.BytesIO()
                        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                            for sheet_name, df in merged_sheets.items():
                                df.to_excel(writer, sheet_name=sheet_name, index=False)
                        mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    else:
                        # Save the merged data to a CSV file in-memory
                        output_buffer = io.StringIO()
                        merged_sheets['Sheet1'].to_csv(output_buffer, index=False)
                        output_buffer = io.BytesIO(output_buffer.getvalue().encode())
                        mime_type = "text/csv"

                    # Provide download link for the merged file
                    output_buffer.seek(0)
                    st.download_button(
                        label=f"Download Merged {output_format.split()[0]} File",
                        data=output_buffer,
                        file_name=output_file,
                        mime=mime_type
                    )
                    
                except Exception as e:
                    st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()

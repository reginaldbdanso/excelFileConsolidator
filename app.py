import streamlit as st
import pandas as pd
import io
import base64
import json
import os
from datetime import datetime

def main():
    st.set_page_config(page_title="Excel File Consolidator", layout="wide")
    
    st.title("Excel File Consolidator")
    
    # Create tabs
    tab1, tab2 = st.tabs(["Columns B-H", "Columns A-G"])
    
    with tab1:
        st.markdown("Upload multiple Excel files to consolidate them (Columns B-H)")
        process_files("B:H")
        
    with tab2:
        st.markdown("Upload multiple Excel files to consolidate them (Columns A-G)")
        process_files("A:G")

def process_files(usecols_range):
    """Handle file processing for each tab"""
    # File uploader
    uploaded_files = st.file_uploader(f"Upload Excel files ({usecols_range})", 
                                    type=["xlsx", "xls"], 
                                    accept_multiple_files=True,
                                    key=f"uploader_{usecols_range}")  # Unique key for each uploader
    
    if uploaded_files:
        st.write(f"Uploaded {len(uploaded_files)} files")
        
        # Processing options
        with st.expander("Processing Options"):
            sheet_name = st.text_input("Sheet name to extract (leave blank for first sheet)", "", key=f"sheet_{usecols_range}")
            header_row = st.number_input("Header row (0-based index)", 0, key=f"header_{usecols_range}")
            include_file_name = st.checkbox("Include source filename as a column", value=True, key=f"filename_{usecols_range}")
        
        if st.button("Process Files", key=f"process_{usecols_range}"):
            with st.spinner("Processing files..."):
                try:
                    # Consolidate all files
                    consolidated_df = consolidate_excel_files(
                        uploaded_files, 
                        sheet_name, 
                        header_row, 
                        include_file_name,
                        usecols_range
                    )
                    
                    # Show preview
                    st.subheader("Preview of Consolidated Data")
                    st.dataframe(consolidated_df.head(10))
                    
                    # Download button
                    download_link = get_download_link(consolidated_df)
                    st.markdown(download_link, unsafe_allow_html=True)
                    
                    # Stats
                    st.success(f"Successfully consolidated {len(uploaded_files)} files into one Excel file with {len(consolidated_df)} rows and {len(consolidated_df.columns)} columns.")
                
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
    
    # Add usage instructions
    with st.expander("How to Use"):
        st.markdown("""
        1. Upload multiple Excel files using the file uploader above
        2. Configure processing options if needed:
           - Specify a sheet name if you want to extract a specific sheet
           - Set the header row if your headers aren't in the first row
           - Choose whether to include source filenames in output
        3. Click "Process Files" to consolidate the data
        4. Preview the results and download the consolidated file
        """)

def consolidate_excel_files(files, sheet_name=None, header_row=0, include_file_name=True, usecols_range="B:J"):
    """Consolidate multiple Excel files into a single DataFrame"""
    all_dfs = []
    
    for file in files:
        try:
            # If sheet_name is provided, use it; otherwise, read the first sheet
            if sheet_name:
                df = pd.read_excel(file, sheet_name=sheet_name, header=header_row, usecols=usecols_range)
            else:
                df = pd.read_excel(file, header=header_row, usecols=usecols_range)
            
            # Remove rows where all columns are empty/NaN
            df = df.dropna(how='all')
            
            # Add filename column if requested
            if include_file_name:
                df['Source_File'] = file.name
                
            all_dfs.append(df)
        except Exception as e:
            st.warning(f"Error processing {file.name}: {str(e)}")
    
    if not all_dfs:
        raise ValueError("No valid data found in the uploaded files")
    
    # Combine all dataframes
    consolidated_df = pd.concat(all_dfs, ignore_index=True)
    
    # Export to JSON
    json_path = os.path.join('data', f'consolidated_data_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json')
    consolidated_df.to_json(json_path, orient='records', indent=2)
    st.info(f"JSON file saved to: {json_path}")
    
    return consolidated_df

def get_download_link(df):
    """Generate a download link for the consolidated Excel file"""
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Consolidated_Data')
    
    # Adjust column widths
    worksheet = writer.sheets['Consolidated_Data']
    for i, col in enumerate(df.columns):
        # Convert all values to strings before calculating max length
        column_width = max(
            df[col].astype(str).str.len().max(),
            len(str(col))
        ) + 2
        worksheet.set_column(i, i, column_width)
    
    writer.close()
    output.seek(0)
    
    # Generate timestamp for filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"consolidated_data_{timestamp}.xlsx"
    
    b64 = base64.b64encode(output.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{file_name}" class="btn" style="background-color:#4CAF50;color:white;padding:8px 12px;text-decoration:none;border-radius:4px;margin-top:10px;display:inline-block;">Download Consolidated Excel File</a>'
    return href

if __name__ == "__main__":
    main()
    
import streamlit as st
import pandas as pd
import os
import glob
import re
from datetime import datetime
from pathlib import Path
import io


class NameSearcher:
    """Name Search Tool reading from fixed folder"""
    
    def __init__(self, excel_folder="excel_output"):
        self.excel_folder = excel_folder
    
    def extract_part_number(self, file_path):
        """Extract part number from row 6 of the Excel file"""
        try:
            df = pd.read_excel(file_path, sheet_name=0, nrows=10, header=None, dtype=str)
            
            if len(df) > 5:
                row_6_data = df.iloc[5]
                
                for cell in row_6_data:
                    if pd.notna(cell):
                        cell_str = str(cell)
                        if ':' in cell_str:
                            part_number = cell_str.split(':')[-1].strip()
                            if part_number:
                                return part_number
            
            return "N/A"
        except Exception as e:
            return "Error"
    
    def search_single_excel_file(self, file_path, search_terms):
        """Search for patterns in a single Excel file"""
        results = []
        
        try:
            part_number = self.extract_part_number(file_path)
            excel_data = pd.read_excel(file_path, sheet_name=None, dtype=str)
            
            for sheet_name, df in excel_data.items():
                for row_idx, row in df.iterrows():
                    for col_idx, cell_value in enumerate(row):
                        if pd.notna(cell_value):
                            cell_str = str(cell_value)
                            
                            for pattern in search_terms:
                                if re.search(pattern, cell_str):
                                    results.append({
                                        'File_Name': os.path.basename(file_path),
                                        'Part_Number': part_number,
                                        'Row': row_idx + 2,
                                        'Matched_Content': cell_str,
                                        'Search_Pattern': pattern
                                    })
                                    break
        except Exception as e:
            st.error(f"Error reading {os.path.basename(file_path)}: {e}")
        
        return results
    
    def search_all_excel_files(self, search_terms):
        """Search all Excel files in the fixed folder"""
        excel_files = glob.glob(os.path.join(self.excel_folder, "*.xlsx")) + \
                      glob.glob(os.path.join(self.excel_folder, "*.xls"))
        
        if not excel_files:
            return None, f"No Excel files found in '{self.excel_folder}' folder"
        
        all_results = []
        progress_placeholder = st.empty()
        
        for idx, file_path in enumerate(excel_files):
            filename = os.path.basename(file_path)
            progress_placeholder.text(f"📄 Searching: {filename}... ({idx + 1}/{len(excel_files)})")
            
            file_results = self.search_single_excel_file(file_path, search_terms)
            all_results.extend(file_results)
        
        progress_placeholder.empty()
        
        return all_results, len(excel_files)
    
    def create_results_excel(self, results, search_terms_display):
        """Create Excel file with results"""
        output = io.BytesIO()
        
        results_df = pd.DataFrame(results)
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            results_df.to_excel(writer, sheet_name='Search_Results', index=False)
            
            if len(results) > 0:
                file_summary = results_df.groupby('File_Name').size().reset_index(name='Match_Count')
                file_summary = file_summary.sort_values('Match_Count', ascending=False)
                file_summary.to_excel(writer, sheet_name='Summary_by_File', index=False)
                
                part_summary = results_df.groupby('Part_Number').size().reset_index(name='Match_Count')
                part_summary = part_summary.sort_values('Match_Count', ascending=False)
                part_summary.to_excel(writer, sheet_name='Summary_by_Part', index=False)
            
            patterns_df = pd.DataFrame({'Search_Terms_Used': search_terms_display})
            patterns_df.to_excel(writer, sheet_name='Search_Terms', index=False)
        
        output.seek(0)
        return output


def main():
    st.set_page_config(
        page_title="Name Search Tool",
        page_icon="🔍",
        layout="wide"
    )
    
    st.title("🔍 Name Search Tool")
    st.markdown("Search for names in Excel files easily!")
    
    EXCEL_FOLDER = "excel_output"
    
    # Check if folder exists
    if not os.path.exists(EXCEL_FOLDER):
        st.error(f"❌ Excel folder '{EXCEL_FOLDER}' not found in repository!")
        st.info("Please add the 'excel_output' folder with Excel files to your GitHub repository.")
        st.stop()
    
    # Count files
    excel_files = glob.glob(os.path.join(EXCEL_FOLDER, "*.xlsx")) + \
                  glob.glob(os.path.join(EXCEL_FOLDER, "*.xls"))
    
    if not excel_files:
        st.error(f"❌ No Excel files found in '{EXCEL_FOLDER}' folder!")
        st.info("Please add Excel files to the 'excel_output' folder in your GitHub repository.")
        st.stop()
    
    st.success(f"✅ {len(excel_files)} Excel files loaded and ready!")
    
    with st.expander(f"📂 View Available Files ({len(excel_files)} files)"):
        for i, file in enumerate(excel_files, 1):
            st.write(f"{i}. {os.path.basename(file)}")
    
    st.markdown("---")
    
    searcher = NameSearcher(EXCEL_FOLDER)
    
    if 'search_terms' not in st.session_state:
        st.session_state.search_terms = None
    if 'search_terms_display' not in st.session_state:
        st.session_state.search_terms_display = []
    
    # Sidebar
    st.sidebar.header("📋 Provide Names to Search")
    
    input_method = st.sidebar.radio(
        "Select input method:",
        ["Type Names Manually", "Upload Text File (.txt)", "Upload Excel File"],
    )
    
    if input_method == "Type Names Manually":
        st.sidebar.markdown("**Enter names (one per line):**")
        manual_input = st.sidebar.text_area(
            "Type here:",
            height=250,
            placeholder="પટેલ\nશાહ\nPatel\nShah",
        )
        
        if st.sidebar.button("✅ Load Names", type="primary", use_container_width=True):
            if manual_input.strip():
                lines = [line.strip() for line in manual_input.splitlines() if line.strip()]
                
                search_terms = []
                for line in lines:
                    search_terms.append(f".*{line}.*")
                    search_terms.append(f"(?i).*{line}.*")
                
                st.session_state.search_terms = search_terms
                st.session_state.search_terms_display = lines
                st.sidebar.success(f"✅ {len(lines)} names loaded!")
                st.rerun()
            else:
                st.sidebar.error("⚠️ Enter at least one name")
    
    elif input_method == "Upload Text File (.txt)":
        st.sidebar.markdown("**Upload text file:**")
        txt_file = st.sidebar.file_uploader("Choose file", type=['txt'], key="txt")
        
        if txt_file and st.sidebar.button("✅ Load", type="primary", use_container_width=True):
            try:
                lines = txt_file.read().decode('utf-8').splitlines()
                lines = [line.strip() for line in lines if line.strip()]
                
                search_terms = []
                for line in lines:
                    search_terms.append(f".*{line}.*")
                    search_terms.append(f"(?i).*{line}.*")
                
                st.session_state.search_terms = search_terms
                st.session_state.search_terms_display = lines
                st.sidebar.success(f"✅ {len(lines)} names loaded!")
                st.rerun()
            except Exception as e:
                st.sidebar.error(f"Error: {e}")
    
    elif input_method == "Upload Excel File":
        st.sidebar.markdown("**Upload Excel file:**")
        excel_input = st.sidebar.file_uploader("Choose file", type=['xlsx', 'xls'], key="excel")
        
        if excel_input:
            try:
                df = pd.read_excel(excel_input, dtype=str)
                column = st.sidebar.selectbox("Select column:", df.columns)
                
                if st.sidebar.button("✅ Load", type="primary", use_container_width=True):
                    values = df[column].dropna().unique()
                    
                    search_terms = []
                    display_names = []
                    for value in values:
                        value_str = str(value).strip()
                        if value_str:
                            search_terms.append(f".*{value_str}.*")
                            search_terms.append(f"(?i).*{value_str}.*")
                            display_names.append(value_str)
                    
                    st.session_state.search_terms = search_terms
                    st.session_state.search_terms_display = display_names
                    st.sidebar.success(f"✅ {len(values)} names loaded!")
                    st.rerun()
            except Exception as e:
                st.sidebar.error(f"Error: {e}")
    
    if st.session_state.search_terms:
        st.sidebar.markdown("---")
        st.sidebar.success(f"✅ **{len(st.session_state.search_terms_display)} names ready!**")
        
        with st.sidebar.expander("👁️ View"):
            for i, name in enumerate(st.session_state.search_terms_display, 1):
                st.write(f"{i}. {name}")
        
        if st.sidebar.button("🗑️ Clear", use_container_width=True):
            st.session_state.search_terms = None
            st.session_state.search_terms_display = []
            st.rerun()
    
    # Main area
    if st.session_state.search_terms:
        st.header("🔍 Ready to Search!")
        st.info(f"Will search in {len(excel_files)} Excel files")
        
        if st.button("🚀 START SEARCH", type="primary", use_container_width=True):
            with st.spinner("Searching..."):
                results, file_count = searcher.search_all_excel_files(st.session_state.search_terms)
                
                if results is None:
                    st.error(file_count)
                else:
                    st.markdown("---")
                    st.header("📊 Results")
                    
                    if results:
                        st.success(f"🎉 Found {len(results)} matches!")
                        
                        results_df = pd.DataFrame(results)
                        
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("Matches", len(results))
                        with col2:
                            st.metric("Files Searched", file_count)
                        with col3:
                            st.metric("Files with Matches", results_df['File_Name'].nunique())
                        with col4:
                            st.metric("Unique Parts", results_df['Part_Number'].nunique())
                        
                        st.subheader("📋 Details")
                        st.dataframe(
                            results_df,
                            use_container_width=True,
                            height=400,
                            column_config={
                                "File_Name": "📄 File",
                                "Part_Number": "🔢 Part",
                                "Row": "📍 Row",
                                "Matched_Content": "✅ Found",
                            },
                            hide_index=True
                        )
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.subheader("📈 By File")
                            fs = results_df.groupby('File_Name').size().reset_index(name='Matches')
                            st.dataframe(fs.sort_values('Matches', ascending=False), hide_index=True)
                        
                        with col2:
                            st.subheader("📊 By Part")
                            ps = results_df.groupby('Part_Number').size().reset_index(name='Matches')
                            st.dataframe(ps.sort_values('Matches', ascending=False), hide_index=True)
                        
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        excel_output = searcher.create_results_excel(results, st.session_state.search_terms_display)
                        
                        st.download_button(
                            label="📥 Download Results",
                            data=excel_output,
                            file_name=f"results_{timestamp}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            type="primary"
                        )
                    else:
                        st.warning("❌ No matches found")
    else:
        st.info("👈 **Please provide names to search (sidebar)**")
        st.markdown("### 📝 How to use:")
        st.markdown("1. Choose input method from sidebar")
        st.markdown("2. Load the names to search")
        st.markdown("3. Click START SEARCH")
        st.markdown("4. Download results")


if __name__ == "__main__":
    main()

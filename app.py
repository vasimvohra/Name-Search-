import streamlit as st
import pandas as pd
import os
import glob
import re
from datetime import datetime
from pathlib import Path
import io


class GujaratiNameSearcher:
    """Streamlit-based Name Search Tool with fixed Excel folder"""

    def __init__(self, excel_folder="excel_output", results_folder="results"):
        self.excel_folder = excel_folder
        self.results_folder = results_folder
        Path(self.results_folder).mkdir(parents=True, exist_ok=True)

    def extract_part_number(self, file_path):
        """Extract part number from row 6 of the Excel file"""
        try:
            # Read only first sheet, first 10 rows to get row 6
            df = pd.read_excel(file_path, sheet_name=0, nrows=10, header=None, dtype=str)

            # Row 6 is index 5 (0-based)
            if len(df) > 5:
                row_6_data = df.iloc[5]

                # Search through all cells in row 6
                for cell in row_6_data:
                    if pd.notna(cell):
                        cell_str = str(cell)
                        # Look for text containing ":" and extract what comes after it
                        if ':' in cell_str:
                            # Extract everything after the last ":"
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
            # Extract part number once for this file
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
            progress_placeholder.text(f"📄 Searching in: {filename}... ({idx + 1}/{len(excel_files)})")

            file_results = self.search_single_excel_file(file_path, search_terms)
            all_results.extend(file_results)

        progress_placeholder.empty()

        return all_results, len(excel_files)

    def create_results_excel(self, results, search_terms_display):
        """Create Excel file with results"""
        output = io.BytesIO()

        results_df = pd.DataFrame(results)

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Main results
            results_df.to_excel(writer, sheet_name='Search_Results', index=False)

            # Summary by file
            if len(results) > 0:
                file_summary = results_df.groupby('File_Name').size().reset_index(name='Match_Count')
                file_summary = file_summary.sort_values('Match_Count', ascending=False)
                file_summary.to_excel(writer, sheet_name='Summary_by_File', index=False)

                # Summary by part number
                part_summary = results_df.groupby('Part_Number').size().reset_index(name='Match_Count')
                part_summary = part_summary.sort_values('Match_Count', ascending=False)
                part_summary.to_excel(writer, sheet_name='Summary_by_Part', index=False)

                # Summary by pattern
                pattern_summary = results_df.groupby('Search_Pattern').size().reset_index(name='Match_Count')
                pattern_summary = pattern_summary.sort_values('Match_Count', ascending=False)
                pattern_summary.to_excel(writer, sheet_name='Summary_by_Pattern', index=False)

            # Search patterns used
            patterns_df = pd.DataFrame({'Search_Terms_Used': search_terms_display})
            patterns_df.to_excel(writer, sheet_name='Search_Terms', index=False)

        output.seek(0)
        return output


def main():
    # Page configuration
    st.set_page_config(
        page_title="Name Search Tool",
        page_icon="🔍",
        layout="wide"
    )

    # Title
    st.title("🔍 Name Search Tool")
    st.markdown("Search for names in Excel files easily!")

    # Fixed paths
    EXCEL_FOLDER = "excel_output"
    RESULTS_FOLDER = "results"

    # Check if Excel folder exists
    if not os.path.exists(EXCEL_FOLDER):
        st.error(f"❌ Excel folder '{EXCEL_FOLDER}' not found!")
        st.info(f"Please create the '{EXCEL_FOLDER}' folder and add Excel files to search.")
        st.stop()

    # Count files in folder
    excel_files = glob.glob(os.path.join(EXCEL_FOLDER, "*.xlsx")) + \
                  glob.glob(os.path.join(EXCEL_FOLDER, "*.xls"))

    st.info(f"📁 Excel Folder: **{EXCEL_FOLDER}** ({len(excel_files)} files available)")

    st.markdown("---")

    # Initialize searcher
    searcher = GujaratiNameSearcher(EXCEL_FOLDER, RESULTS_FOLDER)

    # Initialize session state
    if 'search_terms' not in st.session_state:
        st.session_state.search_terms = None
    if 'search_terms_display' not in st.session_state:
        st.session_state.search_terms_display = []

    # Sidebar for input method
    st.sidebar.header("📋 Provide Names to Search")

    input_method = st.sidebar.radio(
        "How do you want to provide names?",
        ["Type Names Manually", "Upload Text File (.txt)", "Upload Excel File"],
        help="Choose how you want to provide the list of names to search for"
    )

    # Handle different input methods
    if input_method == "Type Names Manually":
        st.sidebar.markdown("**Enter names to search (one per line):**")
        manual_input = st.sidebar.text_area(
            "Type names here:",
            height=250,
            placeholder="Example:\nપટેલ\nશાહ\nPatel\nShah\nઅમીન",
            help="Enter each name on a new line"
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
                st.sidebar.success(f"✅ Loaded {len(lines)} names!")
                st.rerun()
            else:
                st.sidebar.error("⚠️ Please enter at least one name")

    elif input_method == "Upload Text File (.txt)":
        st.sidebar.markdown("**Upload a text file with names (one per line):**")
        txt_file = st.sidebar.file_uploader(
            "Choose text file",
            type=['txt'],
            key="txt_uploader",
            help="Upload a .txt file with one name per line"
        )

        if txt_file:
            if st.sidebar.button("✅ Load Names from File", type="primary", use_container_width=True):
                try:
                    lines = txt_file.read().decode('utf-8').splitlines()
                    lines = [line.strip() for line in lines if line.strip()]

                    search_terms = []
                    for line in lines:
                        search_terms.append(f".*{line}.*")
                        search_terms.append(f"(?i).*{line}.*")

                    st.session_state.search_terms = search_terms
                    st.session_state.search_terms_display = lines
                    st.sidebar.success(f"✅ Loaded {len(lines)} names from file!")
                    st.rerun()
                except Exception as e:
                    st.sidebar.error(f"Error reading file: {e}")

    elif input_method == "Upload Excel File":
        st.sidebar.markdown("**Upload an Excel file containing the names:**")
        excel_input = st.sidebar.file_uploader(
            "Choose Excel file with names",
            type=['xlsx', 'xls'],
            key="excel_uploader",
            help="Upload an Excel file that contains the list of names to search for"
        )

        if excel_input:
            try:
                df = pd.read_excel(excel_input, dtype=str)

                # Let user select column
                column = st.sidebar.selectbox(
                    "Select column containing names:",
                    df.columns,
                    help="Choose which column has the names you want to search for"
                )

                if st.sidebar.button("✅ Load Names from Excel", type="primary", use_container_width=True):
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
                    st.sidebar.success(f"✅ Loaded {len(values)} names from '{column}'!")
                    st.rerun()
            except Exception as e:
                st.sidebar.error(f"Error reading Excel: {e}")

    # Show loaded search terms
    if st.session_state.search_terms:
        st.sidebar.markdown("---")
        st.sidebar.success(f"✅ **{len(st.session_state.search_terms_display)} names ready!**")

        with st.sidebar.expander("👁️ View Loaded Names"):
            for i, name in enumerate(st.session_state.search_terms_display, 1):
                st.write(f"{i}. {name}")

        if st.sidebar.button("🗑️ Clear Names", use_container_width=True):
            st.session_state.search_terms = None
            st.session_state.search_terms_display = []
            st.rerun()

    # Main content area
    if st.session_state.search_terms:
        st.header("🔍 Search in Excel Files")

        # Show files that will be searched
        with st.expander(f"📂 Files to be searched ({len(excel_files)} files in '{EXCEL_FOLDER}')"):
            for i, file in enumerate(excel_files, 1):
                st.write(f"{i}. {os.path.basename(file)}")

        # Search button
        if st.button("🚀 START SEARCH", type="primary", use_container_width=True):
            with st.spinner("Searching..."):
                results, file_count = searcher.search_all_excel_files(st.session_state.search_terms)

                if results is None:
                    st.error(file_count)
                else:
                    # Display results
                    st.markdown("---")
                    st.header("📊 Search Results")

                    if results:
                        st.success(f"🎉 Found {len(results)} matches in {file_count} files!")

                        results_df = pd.DataFrame(results)

                        # Summary metrics
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("Total Matches", len(results))
                        with col2:
                            st.metric("Files Searched", file_count)
                        with col3:
                            st.metric("Files with Matches", results_df['File_Name'].nunique())
                        with col4:
                            st.metric("Unique Parts", results_df['Part_Number'].nunique())

                        # Display results table
                        st.subheader("📋 Detailed Results")
                        st.dataframe(
                            results_df,
                            use_container_width=True,
                            height=400,
                            column_config={
                                "File_Name": st.column_config.TextColumn("📄 File Name", width="medium"),
                                "Part_Number": st.column_config.TextColumn("🔢 Part Number", width="small"),
                                "Row": st.column_config.NumberColumn("📍 Row", width="small"),
                                "Matched_Content": st.column_config.TextColumn("✅ Content Found", width="large"),
                            },
                            hide_index=True
                        )

                        # Summary tables
                        col1, col2 = st.columns(2)

                        with col1:
                            st.subheader("📈 Matches by File")
                            file_summary = results_df.groupby('File_Name').size().reset_index(name='Matches')
                            file_summary = file_summary.sort_values('Matches', ascending=False)
                            st.dataframe(file_summary, use_container_width=True, hide_index=True)

                        with col2:
                            st.subheader("📊 Matches by Part Number")
                            part_summary = results_df.groupby('Part_Number').size().reset_index(name='Matches')
                            part_summary = part_summary.sort_values('Matches', ascending=False)
                            st.dataframe(part_summary, use_container_width=True, hide_index=True)

                        # Create Excel file for download
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        excel_output = searcher.create_results_excel(
                            results,
                            st.session_state.search_terms_display
                        )

                        st.download_button(
                            label="📥 Download Results as Excel",
                            data=excel_output,
                            file_name=f"search_results_{timestamp}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            type="primary"
                        )

                    else:
                        st.warning(f"❌ No matches found in {file_count} files.")
                        st.info(
                            "💡 Try different search terms or check if the Excel files contain the names you're looking for.")

    else:
        st.info("👈 **Please provide names to search from the sidebar**")
        st.markdown("### 📝 How to use:")
        st.markdown("1. **Choose input method** from the sidebar (Type, Text File, or Excel)")
        st.markdown("2. **Load the names** you want to search for")
        st.markdown("3. **Click 'START SEARCH'** to search in all files")
        st.markdown(f"4. **Download results** as Excel file")

        st.markdown(f"\n💡 All files in the '**{EXCEL_FOLDER}**' folder will be searched automatically")

    # Footer
    st.markdown("---")
    st.caption(f"📁 Searching in: {EXCEL_FOLDER} | 💾 Results saved to: {RESULTS_FOLDER}")


if __name__ == "__main__":
    main()

import streamlit as st
import pandas as pd
import numpy as np
import io
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill

# --- Streamlit Page Configuration ---
st.set_page_config(page_title="B Segment Analysis Tool", page_icon="📊", layout="wide")

st.title("B Segment Analysis Tool")
st.write("Upload iCMD_A segment analysis, Mapping, and KNB1 files to generate the B Segment report.")

# --- File Uploaders ---
uploaded_analysis = st.file_uploader("Upload iCMD_A segment analysis.xlsx", type=["xlsx"])
uploaded_mapping = st.file_uploader("Upload iCMD_F2SQ_2026_MDM Business Partner S4_Template_CMD_mapping_XXXX.xlsx", type=["xlsx"])
uploaded_knb1 = st.file_uploader("Upload KNB1.xlsx", type=["xlsx"])

# --- Helper Functions ---

def load_data(analysis_file, mapping_file, knb1_file):
    r3_dtype = {
        'Customer': str, 'CoCd': str, 'Recon.acct': str, 'Sort': str,
        'Clerk Abbrev.': str, 'DeBl': str, 'B': str, 'DelF': str, 'Block ind.': str,
        'Methods': str, 'Value adjustment': str, 'Account memo': str,
        'Record pmnt history': str, "Clerk's internet address": str, 'Account statement': str
    }
    analysis_dtype = {'R3 Customer Number': str, 'R3 Trading Partner': str}
    
    dfs = {
        'analysis': pd.read_excel(analysis_file, sheet_name=0, dtype=analysis_dtype),
        'mapping_fields': pd.read_excel(mapping_file, sheet_name='iCMD_Mapping Fields', header=None),
        'r3_knb1': pd.read_excel(knb1_file, sheet_name=0, dtype=r3_dtype)
    }
    return dfs

def extract_mapping(mapping_df, field_name):
    field_rows = mapping_df[mapping_df.iloc[:, 0] == field_name]
    if field_rows.empty: return {}
    
    mapping_str = str(field_rows.iloc[0, 5])
    mappings = {}
    if isinstance(mapping_str, str):
        for line in mapping_str.split('\n'):
            if '->' in line:
                parts = line.split('->')
                key, value = parts[0].strip(), parts[1].strip()
                if not value and key: mappings[key] = key
                elif key and value: mappings[key] = value
    return mappings

def generate_styled_excel(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_to_save = df.drop(columns=['Is Selected For Upload'])
        df_to_save.to_excel(writer, index=False, sheet_name='B_Segment_Analysis')
        
        workbook = writer.book
        worksheet = writer.sheets['B_Segment_Analysis']
        
        # --- Define All Styles ---
        header_font = Font(name='Calibri', size=10, bold=True, color="FFFFFF")
        data_font = Font(name='Calibri', size=10)
        header_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
        
        thin_border = Border(left=Side(style='thin', color='D3D3D3'), 
                             right=Side(style='thin', color='D3D3D3'), 
                             top=Side(style='thin', color='D3D3D3'), 
                             bottom=Side(style='thin', color='D3D3D3'))

        header_align = Alignment(horizontal='center', vertical='center')
        data_align = Alignment(horizontal='left', vertical='center')
        upload_highlight_fill = PatternFill(start_color="DCC7ED", end_color="DCC7ED", fill_type="solid")
        
        status_fills = {
            'Multi CoCd - Blocked on All': PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
            'Single CoCd - Blocked': PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
            'Multi CoCd - Partially Blocked': PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
            'Multi CoCd - No Blocks': PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
            'Single CoCd - Not Blocked': PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
            'Not Found': PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
        }
        
        headers = [cell.value for cell in worksheet[1]]
        
        # --- Apply Formatting ---
        for cell in worksheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = header_align

        for row_idx, (df_row_idx, df_row) in enumerate(df.iterrows(), 2):
            worksheet_row = worksheet[row_idx]
            for cell in worksheet_row:
                cell.font = data_font
                cell.border = thin_border
                cell.alignment = data_align

            if df_row['Is Selected For Upload']:
                for cell in worksheet_row:
                    cell.fill = upload_highlight_fill
            
            if 'Block Status Analysis' in headers:
                status_cell = worksheet_row[headers.index('Block Status Analysis')]
                if status_cell.value in status_fills and not df_row['Is Selected For Upload']:
                    status_cell.fill = status_fills[status_cell.value]

        # Auto-fit columns
        for column in worksheet.columns:
            max_length = max(len(str(cell.value or '')) for cell in column)
            worksheet.column_dimensions[column[0].column_letter].width = (max_length + 2)
            
    buffer.seek(0)
    return buffer

# --- Generation Logic ---
if uploaded_analysis and uploaded_mapping and uploaded_knb1:
    if st.button("Generate Report"):
        with st.spinner("Processing files and applying mappings..."):
            try:
                # 1. Load Data
                dataframes = load_data(uploaded_analysis, uploaded_mapping, uploaded_knb1)
                
                # Clean column names
                dataframes['analysis'].columns = [c.lower().strip() for c in dataframes['analysis'].columns]
                dataframes['r3_knb1'].columns = [c.lower().strip() for c in dataframes['r3_knb1'].columns]

                # 2. Extract All Mappings
                mapping_df = dataframes['mapping_fields']
                company_map = extract_mapping(mapping_df, 'B SEGMENT - KNB1')
                recon_map = extract_mapping(mapping_df, 'Reconciliation Account in General Ledger')
                sort_key_map = extract_mapping(mapping_df, 'Sort key')
                clerk_map = extract_mapping(mapping_df, 'Clerk Abbreviation')
                methods_map = extract_mapping(mapping_df, 'Payment methods')
                val_adj_map = extract_mapping(mapping_df, 'Value Adjustment')
                acc_memo_map = extract_mapping(mapping_df, 'Account memo')
                rec_pmnt_map = extract_mapping(mapping_df, 'Record pmnt history')
                clerk_inet_map = extract_mapping(mapping_df, "Clerk's internet address")
                acc_stmt_map = extract_mapping(mapping_df, 'Account statement')

                # 3. Prepare and Analyze Data
                r3_data = dataframes['r3_knb1'].copy()
                r3_data['merge_key'] = r3_data['customer'].str.lstrip('0')
                r3_data['is_blocked'] = r3_data[['b', 'delf']].apply(lambda col: col.fillna('').str.strip().str.lower() == 'x').any(axis=1)
                r3_data['cocd'] = r3_data['cocd'].astype(str)
                r3_data['is_leading_cocd'] = r3_data['cocd'].apply(lambda x: company_map.get(x) == x)

                def assign_priority(row):
                    if row['is_leading_cocd'] and not row['is_blocked']: return 1
                    if not row['is_leading_cocd'] and not row['is_blocked']: return 2
                    return 3
                r3_data['priority_score'] = r3_data.apply(assign_priority, axis=1)

                unblocked_rows = r3_data[r3_data['priority_score'] < 3].copy()
                best_rows_idx = unblocked_rows.sort_values('priority_score').drop_duplicates('merge_key', keep='first').index
                r3_data['is_selected_for_upload'] = False
                r3_data.loc[best_rows_idx, 'is_selected_for_upload'] = True

                block_summary = r3_data.groupby('merge_key').agg(num_company_codes=('cocd', 'nunique'), num_blocked_codes=('is_blocked', 'sum')).reset_index()
                def get_block_status(row):
                    if row['num_company_codes'] <= 1: return 'Single CoCd - Blocked' if row['num_blocked_codes'] > 0 else 'Single CoCd - Not Blocked'
                    else:
                        if row['num_blocked_codes'] == 0: return 'Multi CoCd - No Blocks'
                        elif row['num_blocked_codes'] >= row['num_company_codes']: return 'Multi CoCd - Blocked on All'
                        else: return 'Multi CoCd - Partially Blocked'
                block_summary['Block Status Analysis'] = block_summary.apply(get_block_status, axis=1)

                # 4. Perform Merges
                analysis_data = dataframes['analysis'].copy()
                analysis_data.rename(columns={'r3 customer number': 'customer'}, inplace=True)
                analysis_data['merge_key'] = analysis_data['customer'].str.lstrip('0')
                merged_df = pd.merge(r3_data, analysis_data.drop_duplicates(subset=['merge_key']), on='merge_key', how='left')
                merged_df = pd.merge(merged_df, block_summary[['merge_key', 'Block Status Analysis']], on='merge_key', how='left')
                merged_df.drop(columns=['merge_key', 'customer_y', 'is_blocked', 'priority_score'], inplace=True)
                merged_df.rename(columns={'customer_x': 'customer'}, inplace=True)

                # 5. Create S4 Columns and Rename R3 Columns
                merged_df.rename(columns={'cocd': 'Company Code R3', 'recon.acct': 'Recon Account R3', 'sort': 'Sort Key R3', 'clerk abbrev.': 'Clerk Abbreviation R3',
                                             'methods': 'Methods R3', 'value adjustment': 'Value adjustment R3', 'account memo': 'Account memo R3',
                                             'record pmnt history': 'Record pmnt history R3', "clerk's internet address": "Clerk's internet address R3",
                                             'account statement': 'Account statement R3'}, inplace=True)
                merged_df['Company Code S4'] = merged_df['Company Code R3'].map(company_map)
                merged_df['Recon Account S4'] = merged_df['Recon Account R3'].map(recon_map)
                merged_df['Sort Key S4'] = merged_df['Sort Key R3'].map(sort_key_map)
                merged_df['Clerk Abbreviation S4'] = merged_df['Clerk Abbreviation R3'].map(clerk_map)
                merged_df['Methods S4'] = merged_df['Methods R3'].map(methods_map)
                merged_df['Value adjustment S4'] = merged_df['Value adjustment R3'].map(val_adj_map)
                merged_df['Account memo S4'] = merged_df['Account memo R3'].map(acc_memo_map)
                merged_df['Record pmnt history S4'] = merged_df['Record pmnt history R3'].map(rec_pmnt_map)
                merged_df["Clerk's internet address S4"] = merged_df["Clerk's internet address R3"].map(clerk_inet_map)
                merged_df['Account statement S4'] = merged_df['Account statement R3'].map(acc_stmt_map)

                # 6. Finalize DataFrame
                merged_df.rename(columns={'customer': 'R3 Customer Number', 's4 cp': 'S4 CP', 'debl': 'DeBl', 'b': 'B', 'delf': 'DelF',
                                             'block ind.': 'Block ind.', 'is_leading_cocd': 'Is Leading CoCd',
                                             'is_selected_for_upload': 'Is Selected For Upload', 'r3 trading partner': 'Trading Partner'}, inplace=True)

                final_display_columns = [
                    'R3 Customer Number', 'S4 CP', 'Trading Partner', 'Company Code R3', 'Company Code S4', 'Recon Account R3', 'Recon Account S4',
                    'Sort Key R3', 'Sort Key S4', 'Clerk Abbreviation R3', 'Clerk Abbreviation S4', 'DeBl', 'B', 'DelF', 'Block ind.',
                    'Block Status Analysis', 'Is Selected For Upload', 'Is Leading CoCd',
                    'Methods R3', 'Methods S4', 'Value adjustment R3', 'Value adjustment S4', 'Account memo R3', 'Account memo S4',
                    'Record pmnt history R3', 'Record pmnt history S4', "Clerk's internet address R3", "Clerk's internet address S4",
                    'Account statement R3', 'Account statement S4'
                ]
                results_df = merged_df.reindex(columns=final_display_columns)
                results_df = results_df.astype(object).fillna('Not Found')

                # 7. Style Excel in memory
                styled_excel_buffer = generate_styled_excel(results_df)
                
                st.success("🎉 Report generated successfully!")
                st.download_button(
                    label="📥 Download B Segment Analysis Report",
                    data=styled_excel_buffer,
                    file_name="iCMD_B segment analysis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"An error occurred during processing: {e}")
else:
    st.info("Please upload all three required Excel files to enable report generation.")

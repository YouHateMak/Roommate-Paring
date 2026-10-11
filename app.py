import pandas as pd
import itertools
import streamlit as st
import io
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import openpyxl

# Page Configuration
st.set_page_config(page_title="Roommate Pairing App", page_icon="🏡", layout="wide")

st.title("🏡 Roommate Pairing Application")
st.markdown("Upload your Microsoft Forms questionnaire export (`.xlsx`) to generate optimized roommate pairs based on habit compatibility and Myers-Briggs personality alignment[cite: 2].")

# Sidebar - Strictness Control
st.sidebar.header("⚙️ Matching Settings")
strictness = st.sidebar.slider(
    "Compatibility Strictness Level",
    min_value=1,
    max_value=5,
    value=3,
    step=1,
    help="Level 4-5 enforces hard rules: banning large neatness gaps and mismatched sleep/visitor preferences."
)

st.sidebar.markdown(f"**Current Mode:** Level {strictness} ({'Relaxed' if strictness <= 2 else 'Balanced' if strictness == 3 else 'Strict Hard-Constraint Matching'})")

# --- 1. MBTI COMPATIBILITY CHART MATRIX ---
def get_mbti_compatibility(t1, t2):
    t1 = str(t1).upper().strip()
    t2 = str(t2).upper().strip()
    
    types_list = ["INFP", "ENFP", "INFJ", "ENFJ", "INTJ", "ENTJ", "INTP", "ENTP", 
                  "ISFP", "ESFP", "ISTP", "ESTP", "ISFJ", "ESFJ", "ISTJ", "ESTJ"]
                  
    if t1 not in types_list or t2 not in types_list:
        return 5 
        
    i = types_list.index(t1)
    j = types_list.index(t2)
    q1 = i // 4
    q2 = j // 4
    
    matrix_scores = {
        (0, 0): 9, (0, 1): 9, (0, 2): -10, (0, 3): -10,
        (1, 0): 9, (1, 1): 9, (1, 2): 6,   (1, 3): 4,
        (2, 0): -10, (2, 1): 6, (2, 2): 3,   (2, 3): 5,
        (3, 0): -10, (3, 1): 4, (3, 2): 5,   (3, 3): 8
    }
    
    base = matrix_scores.get((q1, q2), 5)
    
    ideal_pairs = {
        ("INFP", "ENFJ"), ("ENFJ", "INFP"),
        ("INFJ", "ENTP"), ("ENTP", "INFJ"),
        ("INTJ", "ENFP"), ("ENFP", "INTJ"),
        ("INTP", "ENTJ"), ("ENTJ", "INTP"),
        ("ISFP", "ESFJ"), ("ESFJ", "ISFP"),
        ("ISTP", "ESTJ"), ("ESTJ", "ISTP"),
        ("ISFJ", "ESFP"), ("ESFP", "ISFJ"),
        ("ISTP", "ESTP")
    }
    if (t1, t2) in ideal_pairs or t1 == t2:
        base = max(base, 10)
        
    return base

# Helper to format location
def format_location(student, cols):
    sa_col = next((c for c in cols if 'live in south africa' in c.lower()), None)
    country_col = next((c for c in cols if 'country' in c.lower()), None)
    province_col = next((c for c in cols if 'province' in c.lower()), None)
    town_col = next((c for c in cols if 'hometown' in c.lower()), None)
    
    is_sa = str(student.get(sa_col, '')).strip().lower() == 'yes' if sa_col else True
    town = str(student.get(town_col, '')).strip() if town_col and pd.notnull(student.get(town_col)) else ''
    
    if is_sa:
        prov = str(student.get(province_col, '')).strip() if province_col and pd.notnull(student.get(province_col)) else ''
        return f"{prov}, {town}".strip(", ") if prov or town else "South Africa"
    else:
        cntry = str(student.get(country_col, '')).strip() if country_col and pd.notnull(student.get(country_col)) else 'International'
        return f"{cntry}, {town}".strip(", ") if cntry or town else "International"

# --- 2. COMPATIBILITY ALGORITHM WITH HARD FILTERS ---
def calculate_score(student_a, student_b, cols, strictness_level):
    neat_col = next((c for c in cols if 'neatness' in c.lower()), None)
    morni_col = next((c for c in cols if 'morning' in c.lower()), None)
    mbti_col = next((c for c in cols if 'meyers' in c.lower() or 'mbti' in c.lower()), None)
    vis_freq_col = next((c for c in cols if 'visitors' in c.lower() and 'frequently' in c.lower()), None)
    vis_comf_col = next((c for c in cols if 'comfortable' in c.lower() and 'vistiors' in c.lower()), None)
    
    # --- HARD CONSTRAINTS FOR STRICTNESS 4 & 5 ---
    if neat_col and pd.notnull(student_a.get(neat_col)) and pd.notnull(student_b.get(neat_col)):
        try:
            n_diff = abs(float(student_a[neat_col]) - float(student_b[neat_col]))
            if strictness_level == 5 and n_diff > 2:
                return -999  # Completely forbid pairing (e.g. 9 paired with 5)
            if strictness_level == 4 and n_diff > 3:
                return -999  # Forbid large neatness gaps
        except:
            pass
            
    if morni_col and strictness_level >= 4:
        ma = str(student_a.get(morni_col, '')).strip().lower()
        mb = str(student_b.get(morni_col, '')).strip().lower()
        if strictness_level == 5 and ma != mb and ma != '' and mb != '':
            return -999  # Level 5 strictly requires matching morning/night schedule
            
    # --- SCORING CALCULATION ---
    score = 100 
    multiplier = 0.5 + (strictness_level * 0.4)
    
    if neat_col:
        try:
            n_diff = abs(float(student_a[neat_col]) - float(student_b[neat_col]))
            score -= (n_diff * 3.5 * multiplier)
        except:
            pass
            
    if morni_col:
        if str(student_a.get(morni_col, '')).strip().lower() != str(student_b.get(morni_col, '')).strip().lower():
            score -= (25 * multiplier)
            
    if vis_freq_col and vis_comf_col:
        freq_a = str(student_a.get(vis_freq_col, '')).strip().lower()
        freq_b = str(student_b.get(vis_freq_col, '')).strip().lower()
        comf_a = str(student_a.get(vis_comf_col, '')).strip().lower()
        comf_b = str(student_b.get(vis_comf_col, '')).strip().lower()
        
        if (freq_a == 'yes' and comf_b == 'no') or (freq_b == 'yes' and comf_a == 'no'):
            score -= (20 * multiplier)
            
    if mbti_col and pd.notnull(student_a.get(mbti_col)) and pd.notnull(student_b.get(mbti_col)):
        mbti_score = get_mbti_compatibility(student_a[mbti_col], student_b[mbti_col])
        score += (mbti_score * 1.5 * multiplier)

    return max(0, min(100, score))

# --- 3. GENERATE PAIRINGS ---
def generate_pairings(df, strictness_level):
    df['Internal_ID'] = range(1, len(df) + 1)
    students = df.to_dict('records')
    possible_pairs = []
    
    cols = df.columns.tolist()
    for student_a, student_b in itertools.combinations(students, 2):
        score = calculate_score(student_a, student_b, cols, strictness_level)
        if score > -500:  # Skip forbidden pairs (-999)
            possible_pairs.append({
                'Person_A': student_a,
                'Person_B': student_b,
                'Score': round(score, 1)
            })
        
    possible_pairs.sort(key=lambda x: x['Score'], reverse=True)
    
    paired_ids = set()
    final_pairings = []
    
    for pair in possible_pairs:
        id_a = pair['Person_A']['Internal_ID'] 
        id_b = pair['Person_B']['Internal_ID']
        
        if id_a not in paired_ids and id_b not in paired_ids:
            final_pairings.append(pair)
            paired_ids.add(id_a)
            paired_ids.add(id_b)
                
    unpaired = [s for s in students if s['Internal_ID'] not in paired_ids]
    return final_pairings, unpaired

# --- 4. EXCEL EXPORT BUILDER ---
def create_styled_excel(pairings, df_columns):
    survey_cols = [c for c in df_columns if c != 'Internal_ID']
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Optimized Roommate Pairs"
    ws.views.sheetView[0].showGridLines = True
    
    headers = ["Pair_Number", "Match_Score"] + survey_cols
    ws.append(headers)
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
    ws.row_dimensions[1].height = 30
    
    palette_hexes = ["FF99C8", "FCF6BD", "D0F4DE", "A9DEF9", "E4C1F9"]
    data_font = Font(name="Calibri", size=10, color="333333")
    
    current_row = 2
    for idx, pair in enumerate(pairings, 1):
        pA = pair['Person_A']
        pB = pair['Person_B']
        score = pair['Score']
        
        color_hex = palette_hexes[(idx - 1) % len(palette_hexes)]
        pair_fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
        
        # Roommate A Row
        row_data_a = [idx, score] + [pA.get(col, '') for col in survey_cols]
        ws.append(row_data_a)
        ws.row_dimensions[current_row].height = 22
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=current_row, column=col_num)
            cell.font = data_font
            cell.fill = pair_fill
            cell.border = thin_border
            if col_num <= 2:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        current_row += 1
        
        # Roommate B Row
        row_data_b = [idx, score] + [pB.get(col, '') for col in survey_cols]
        ws.append(row_data_b)
        ws.row_dimensions[current_row].height = 22
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=current_row, column=col_num)
            cell.font = data_font
            cell.fill = pair_fill
            cell.border = thin_border
            if col_num <= 2:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        current_row += 1
                
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 35)
        
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- 5. APP INTERFACE ---
uploaded_file = st.file_uploader("Upload your Microsoft Forms Excel export (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success(f"Successfully loaded file with {len(df)} student responses!")
        
        if st.button("🚀 Generate Optimized Pairings", type="primary"):
            with st.spinner(f"Analyzing habits with Strictness Level {strictness} (Hard Filters Active)..."):
                pairings, unpaired = generate_pairings(df, strictness)
            
            st.subheader(f"✨ Generated Pairings ({len(pairings)} total pairs - Strictness {strictness}/5)")
            
            name_col = next((c for c in df.columns if 'name' in c.lower() and 'first' in c.lower()), df.columns[4] if len(df.columns) > 4 else df.columns[0])
            mbti_col = next((c for c in df.columns if 'meyers' in c.lower() or 'mbti' in c.lower()), None)
            
            for i, pair in enumerate(pairings, 1):
                pA = pair['Person_A']
                pB = pair['Person_B']
                score = pair['Score']
                
                name_a = pA.get(name_col, 'Student A')
                name_b = pB.get(name_col, 'Student B')
                loc_a = format_location(pA, df.columns)
                loc_b = format_location(pB, df.columns)
                type_a = pA.get(mbti_col, 'N/A') if mbti_col else 'N/A'
                type_b = pB.get(mbti_col, 'N/A') if mbti_col else 'N/A'
                
                with st.container():
                    col1, col2, col3 = st.columns([3, 3, 2])
                    with col1:
                        st.markdown(f"**Roommate 1:** {name_a}")
                        st.caption(f"Location: {loc_a} | Type: {type_a}")
                    with col2:
                        st.markdown(f"**Roommate 2:** {name_b}")
                        st.caption(f"Location: {loc_b} | Type: {type_b}")
                    with col3:
                        st.metric(label="Match Score", value=f"{score}/100")
                    st.divider()
                
            if unpaired:
                st.warning(f"⚠️ Unpaired Students ({len(unpaired)} - Odd total count):")
                for student in unpaired:
                    st.write(f"- {student.get(name_col, 'Student')}")
            
            excel_bytes = create_styled_excel(pairings, df.columns)
            
            st.download_button(
                label="📥 Download Color-Coded Stacked Pairs Excel",
                data=excel_bytes,
                file_name="generated_pairings.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Error processing file: {e}")
else:
    st.info("👆 Upload your Microsoft Forms questionnaire export above to get started.")
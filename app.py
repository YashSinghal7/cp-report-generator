import streamlit as st
import pandas as pd
from io import BytesIO


def calculate_report(df):
    df['outcome'] = df['outcome'].str.strip().str.lower()
    required_cols = ['bot', 'mobile_number', 'outcome', 'contacted', 'date', 'recording_url']
    if not all(col in df.columns for col in required_cols):
        missing_cols = [col for col in required_cols if col not in df.columns]
        st.error(f"Missing columns: {', '.join(missing_cols)}")
        return None, None

    df['date'] = pd.to_datetime(df['date'], errors='coerce')
    df = df.dropna(subset=['date'])

    df['contacted'] = pd.to_numeric(df['contacted'], errors='coerce').fillna(0).astype(int)
    df['bot'] = df['bot'].fillna('Blank_Bot_Name')

    # Normalize recording_url string: ensure empty string for NaN etc
    df['recording_url'] = df['recording_url'].fillna('').astype(str).str.strip()

    all_bots = sorted(df['bot'].unique())
    latest_per_lead = df.sort_values('date').groupby(['bot', 'mobile_number'], as_index=False).last()

    # Outcomes for follow up exclusion
    follow_up_exclude = {"assign to live agent", "converted", "lost"}

    unique_leads_row = {'Metric': 'Unique leads'}
    total_attempts_row = {'Metric': 'Total Attempts'}
    avg_attempts_row = {'Metric': 'Avg Attempts'}
    connected_row = {'Metric': 'Connected'}  # based on recording_url presence
    connectivity_perc_row = {'Metric': 'Connectivity % :'}
    not_connected_row = {'Metric': 'Not Connected'}  # based on recording_url absence
    follow_up_row = {'Metric': 'Follow Up'}  # defined by outcomes
    assigned_agent_row = {'Metric': 'Assigned to human agent'}
    lost_row = {'Metric': 'Lost'}
    converted_row = {'Metric': 'Converted'}

    # Logic for category sheets using recording_url for connected/not connected
    sheets_by_category = {}
    sheets_by_category["connected"] = latest_per_lead[latest_per_lead['recording_url'] != '']
    sheets_by_category["not_connected"] = latest_per_lead[latest_per_lead['recording_url'] == '']
    sheets_by_category["converted"] = latest_per_lead[latest_per_lead['outcome'] == 'converted']
    sheets_by_category["lost"] = latest_per_lead[latest_per_lead['outcome'] == 'lost']
    sheets_by_category["assigned_to_human_agent"] = latest_per_lead[latest_per_lead['outcome'] == 'assign to live agent']
    sheets_by_category["follow_up"] = latest_per_lead[~latest_per_lead['outcome'].isin(follow_up_exclude)]

    # Flags for lead summary updated similarly
    latest_per_lead['connected_flag'] = latest_per_lead['recording_url'] != ''
    latest_per_lead['not_connected_flag'] = latest_per_lead['recording_url'] == ''
    latest_per_lead['converted_flag'] = latest_per_lead['outcome'] == 'converted'
    latest_per_lead['lost_flag'] = latest_per_lead['outcome'] == 'lost'
    latest_per_lead['assigned_to_agent_flag'] = latest_per_lead['outcome'] == 'assign to live agent'
    latest_per_lead['follow_up_flag'] = ~latest_per_lead['outcome'].isin(follow_up_exclude)

    lead_summary_cols = [
        'bot', 'mobile_number', 'date', 'outcome',
        'connected_flag', 'not_connected_flag', 'converted_flag', 'lost_flag',
        'assigned_to_agent_flag', 'follow_up_flag'
    ]
    sheets_by_category["lead_summary"] = latest_per_lead[lead_summary_cols].sort_values(
        ['bot', 'date'], ascending=[True, False]
    )

    for bot_name in all_bots:
        bot_df_all = df[df['bot'] == bot_name]
        total_attempts = len(bot_df_all)
        unique_leads = bot_df_all['mobile_number'].nunique()
        avg_attempts = round(total_attempts / unique_leads, 2) if unique_leads > 0 else 0.0

        bot_latest = latest_per_lead[latest_per_lead['bot'] == bot_name]

        connected_count = bot_latest[bot_latest['recording_url'] != '']['mobile_number'].nunique()
        not_connected_count = bot_latest[bot_latest['recording_url'] == '']['mobile_number'].nunique()
        connectivity_perc = round((connected_count / unique_leads), 2) if unique_leads > 0 else 0.0
        follow_up_count = bot_latest[~bot_latest['outcome'].isin(follow_up_exclude)]['mobile_number'].nunique()
        assigned_to_agent = bot_latest[bot_latest['outcome'] == 'assign to live agent']['mobile_number'].nunique()
        lost = bot_latest[bot_latest['outcome'] == 'lost']['mobile_number'].nunique()
        converted = bot_latest[bot_latest['outcome'] == 'converted']['mobile_number'].nunique()

        unique_leads_row[bot_name] = unique_leads
        total_attempts_row[bot_name] = total_attempts
        avg_attempts_row[bot_name] = avg_attempts
        connected_row[bot_name] = connected_count
        connectivity_perc_row[bot_name] = connectivity_perc
        not_connected_row[bot_name] = not_connected_count
        follow_up_row[bot_name] = follow_up_count
        assigned_agent_row[bot_name] = assigned_to_agent
        lost_row[bot_name] = lost
        converted_row[bot_name] = converted

    report_data = [
        unique_leads_row,
        total_attempts_row,
        avg_attempts_row,
        connected_row,
        connectivity_perc_row,
        not_connected_row,
        follow_up_row,
        assigned_agent_row,
        lost_row,
        converted_row
    ]

    report_df = pd.DataFrame(report_data)
    report_df = report_df.set_index('Metric')

    return report_df, sheets_by_category


def style_summary_df(df):
    def style_data_rows(s):
        style = 'background-color: #C0C0C0; color: #000000; border: 1px solid #000000;'
        if s.name == 'Converted':
            style = 'background-color: #fff700; color: #000000; border: 1px solid #000000;'
        return [style] * len(s)

    def style_index_cells(label):
        style = 'background-color: #C0C0C0; color: #000000; font-weight: bold; border: 1px solid #000000;'
        if label == 'Converted':
            style = 'background-color: #fff700; color: #000000; font-weight: bold; border: 1px solid #000000;'
        return style

    styler = df.style
    styler = styler.apply(style_data_rows, axis=1, subset=pd.IndexSlice[:, df.columns])
    styler = styler.map_index(style_index_cells, axis=0)
    styler = styler.set_table_styles([
        {'selector': 'th.col_heading', 'props': [
            ('background-color', '#fff700'),
            ('color', '#000000'),
            ('font-weight', 'bold'),
            ('border', '1px solid #000000')
        ]},
        {'selector': 'th.index_name', 'props': [
            ('background-color', '#C0C0C0'),
            ('color', '#000000'),
            ('font-weight', 'bold'),
            ('border', '1px solid #000000')
        ]}
    ], overwrite=True)
    styler = styler.format('{:.2f}', subset=pd.IndexSlice[['Avg Attempts', 'Connectivity % :'], :])
    return styler


def style_generic_df(df):
    styler = df.style
    styler = styler.set_table_styles([
        {'selector': 'th', 'props': [
            ('background-color', '#fff700'),
            ('color', '#000000'),
            ('font-weight', 'bold')
        ]}
    ], overwrite=True)

    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    if numeric_cols:
        styler = styler.format('{:.2f}', subset=numeric_cols)
    return styler


st.set_page_config(page_title="Call Report Generator", layout="wide")
st.title("📞 Call Performance Report Generator")

st.info("Upload your raw call log file (CSV or XLSX) including a 'date' column with call timestamps.")

uploaded_file = st.file_uploader("Select Call Log File", type=["csv", "xlsx"])

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            raw_df = pd.read_csv(uploaded_file, dtype={'mobile_number': str})
        else:
            raw_df = pd.read_excel(uploaded_file, dtype={'mobile_number': str})

        st.subheader("Raw Data Preview (First 5 Rows)")
        st.dataframe(raw_df.head())

        report_df, sheets_by_category = calculate_report(raw_df)

        if report_df is not None:
            st.subheader("📊 Calculated Summary Report")

            styled_summary = style_summary_df(report_df)
            st.dataframe(styled_summary)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                styled_summary.to_excel(writer, sheet_name="Summary", index=True)

                for category_name, df_cat in sheets_by_category.items():
                    styled_cat = style_generic_df(df_cat)
                    styled_cat.to_excel(writer, sheet_name=category_name, index=False)

            output.seek(0)

            st.download_button(
                label="📥 Download Styled Excel Report",
                data=output,
                file_name="Styled_Call_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error during processing: {e}")
        st.exception(e)

else:
    st.warning("Please upload a CSV or XLSX file to begin.")

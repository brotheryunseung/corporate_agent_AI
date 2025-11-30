import streamlit as st
from report_generator import generate_research_report
from dotenv import load_dotenv
load_dotenv()

import os


from corporate_tools import (
    corporate_full_analysis,
    compute_core_metrics,
    score_from_metrics,
    create_word_report
)

st.set_page_config(page_title="AI Corporate Analyst", layout="wide")

st.title("🏦 AI Corporate Analyst – Personal Research Agent")

ticker = st.text_input("Enter ticker (e.g., MSFT, AAPL, TSLA):", value="MSFT")

if st.button("Analyze"):
    with st.spinner("Running analysis..."):
        
        # 1) Generate Analysis Text
        analysis = corporate_full_analysis(ticker)
        
        # 2) Rating Engine
        metrics = compute_core_metrics(ticker)
        rating = score_from_metrics(metrics)
        
        st.subheader("📊 Corporate Analysis Report")
        st.text(analysis)

        st.subheader("📈 Investment Rating")
        st.json(rating)
        
        # 3) Generate Word Report
        file_path = create_word_report(ticker, analysis, rating)
        with open(file_path, "rb") as f:
            st.download_button(
                "📥 Download Word Report",
                data=f,
                file_name=file_path,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
from word_generator import export_to_word

# ...

if st.button("Generate Report"):
    with st.spinner("Fetching financial data..."):
        core_data = corporate_full_analysis(ticker)

    with st.spinner("Generating full analyst report..."):
        report = generate_research_report(ticker, core_data)

    st.success("Report generated!")
    st.text_area("Full Report", report, height=500)

    # Word Download Button
    if st.button("Download Word Report"):
        filename = export_to_word(report)
        with open(filename, "rb") as docx_file:
            st.download_button(
                label="Download Word Report (.docx)",
                data=docx_file,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )  

from langchain_openai import ChatOpenAI

def generate_research_report(ticker: str, core_data: str) -> str:
    """
    Convert the corporate_full_analysis() raw text into
    a polished Wall Street-style equity research report.

    ticker      : 'MSFT'
    core_data   : text output from corporate_full_analysis()
    """

    llm = ChatOpenAI(model="gpt-4.1", temperature=0)

    prompt = f"""
    You are a top-tier Wall Street equity research analyst.
    Using the financial data, analysis, and context provided below,
    write a comprehensive and professional equity research report.

    Follow this EXACT structure:

    ============================
    1. Executive Summary
    - A concise overview of the company, recent performance, and key insights.
    - Provide an investment recommendation (Buy / Hold / Sell).

    2. Company Overview
    - Business model, main revenue sources, competitive landscape.

    3. Financial Performance
    - Revenue trend, margin trend, EPS trend.
    - Discuss profitability and cash flow quality.

    4. Key Metrics Summary
    - Include valuation metrics (P/E, Forward P/E, P/S, P/B).
    - Include growth metrics (YoY revenue, YoY EPS).
    - Include profitability metrics (ROE, ROA, margins).
    - Include leverage and risk metrics.

    5. Valuation & Price Target
    - Provide a fair value estimate.
    - Use rationale based on valuation multiples or growth outlook.

    6. Risk Factors
    - Industry risk, macro risk, company-specific risk.

    7. Investment Thesis
    - Bull Case (2–3 bullet points)
    - Base Case (2–3 bullet points)
    - Bear Case (2–3 bullet points)

    8. Conclusion & Recommendation
    ============================

    Make the writing:
    - Clear
    - Professional
    - Data-backed
    - Aligned with real analyst reports

    TICKER: {ticker}

    === RAW FINANCIAL DATA FROM TOOLS ===
    {core_data}
    =====================================

    Begin the full analyst report now:
    """

    response = llm.invoke(prompt)
    return response.content
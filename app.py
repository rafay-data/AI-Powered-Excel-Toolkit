# Improved Streamlit App with Modern UI/UX
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
import os
from datetime import datetime

# ---------- Page Configuration ---------- #
st.set_page_config(
    page_title="AI Excel Business Toolkit",
    layout="centered",
    page_icon="📊",
)

# ---------- Styling ---------- #
st.markdown("""
    <style>
        .main {
            background-color: #f8f9fa;
        }
        .stButton>button {
            background-color: #4CAF50;
            color: white;
            font-weight: bold;
            border-radius: 8px;
            padding: 10px 20px;
        }
        .stButton>button:hover {
            background-color: #45a049;
        }
        .block-container {
            padding-top: 2rem;
            padding-bottom: 2rem;
        }
    </style>
""", unsafe_allow_html=True)

# ---------- Header ---------- #
st.image("https://i.ibb.co/fxdtPfV/ai-excel-logo.png", width=100)
st.title("📊 AI-Powered Excel Business Toolkit")
st.markdown("""
Welcome to your smart assistant for analyzing business data. Just upload your Excel sheet, and get instant insights, beautiful charts, and a downloadable report.
""")

# ---------- Upload ---------- #
uploaded_file = st.file_uploader("📂 Upload your Excel file", type=["xlsx"])
os.makedirs("output", exist_ok=True)

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # ---------- Summary ---------- #
    total_sales = df["Total"].sum()
    top_customers = df.groupby("Customer")["Total"].sum().sort_values(ascending=False).head(3)
    top_products = df.groupby("Product")["Total"].sum().sort_values(ascending=False).head(3)

    st.markdown("## 📈 Business Summary")
    col1, col2 = st.columns(2)
    col1.metric("💰 Total Sales", f"${total_sales:,.2f}")
    col2.metric("👥 Top Customer", top_customers.idxmax())

    st.markdown("### 🏆 Top Customers")
    st.dataframe(top_customers.reset_index().rename(columns={"Total": "Total Spent"}))

    st.markdown("### 🛒 Top Products")
    st.dataframe(top_products.reset_index().rename(columns={"Total": "Total Sold"}))

    # Save summary dict
    summary = {
        "Total Sales": total_sales,
        "Top Customers": top_customers.to_dict(),
        "Top Products": top_products.to_dict()
    }

    # ---------- Visualizations ---------- #
    st.markdown("## 📊 Visualizations")

    sales_by_product = df.groupby("Product")["Total"].sum()
    fig1, ax1 = plt.subplots()
    sales_by_product.plot(kind="bar", ax=ax1, title="Sales by Product", color="#4CAF50")
    plt.xticks(rotation=45)
    plt.tight_layout()
    st.pyplot(fig1)
    fig1.savefig("output/sales_by_product.png")

    df["Date"] = pd.to_datetime(df["Date"])
    sales_over_time = df.groupby("Date")["Total"].sum()
    fig2, ax2 = plt.subplots()
    sales_over_time.plot(ax=ax2, title="Sales Over Time", color="#2196F3")
    plt.xticks(rotation=45)
    plt.tight_layout()
    st.pyplot(fig2)
    fig2.savefig("output/sales_over_time.png")

    # ---------- Export PDF ---------- #
    def export_pdf(summary_dict, output_path="output/business_report.pdf"):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 16)
        pdf.cell(0, 10, "Business Summary Report", ln=True, align="C")

        pdf.set_font("Helvetica", size=12)
        pdf.ln(10)
        pdf.cell(0, 10, f"Total Sales: ${summary_dict['Total Sales']:,.2f}", ln=True)

        pdf.ln(5)
        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(0, 10, "Top Customers:", ln=True)
        pdf.set_font("Helvetica", size=12)
        for k, v in summary_dict["Top Customers"].items():
            pdf.cell(0, 10, f"- {k}: ${v:,.2f}", ln=True)

        pdf.ln(5)
        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(0, 10, "Top Products:", ln=True)
        pdf.set_font("Helvetica", size=12)
        for k, v in summary_dict["Top Products"].items():
            pdf.cell(0, 10, f"- {k}: ${v:,.2f}", ln=True)

        pdf.output(output_path)

    if st.button("📄 Export PDF Report"):
        export_pdf(summary)
        st.success("✅ PDF report saved to output folder!")

# ---------- Footer ---------- #
st.markdown("""
---
Made with ❤️ by **Rafay** | [GitHub](https://github.com/rafay-data) | [Contact](mailto:rafay@example.com)
""")

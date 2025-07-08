import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
import os

# Create folders if not exist
os.makedirs("output", exist_ok=True)

# UI Title
st.title("📊 AI-Powered Excel Business Toolkit")

# Upload Excel file
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Business summary
    total_sales = df["Total"].sum()
    top_customers = df.groupby("Customer")["Total"].sum().sort_values(ascending=False).head(3)
    top_products = df.groupby("Product")["Total"].sum().sort_values(ascending=False).head(3)

    st.subheader("📈 Business Summary")
    st.write(f"**Total Sales:** ${total_sales:,.2f}")
    st.write("**Top Customers:**")
    st.dataframe(top_customers)
    st.write("**Top Products:**")
    st.dataframe(top_products)

    # Save summary
    summary = {
        "Total Sales": total_sales,
        "Top Customers": top_customers.to_dict(),
        "Top Products": top_products.to_dict()
    }

    # Charts
    st.subheader("📊 Sales Visualizations")

    # Sales by Product chart
    sales_by_product = df.groupby("Product")["Total"].sum()
    fig1, ax1 = plt.subplots()
    sales_by_product.plot(kind="bar", ax=ax1, title="Sales by Product")
    plt.xticks(rotation=45)
    plt.tight_layout()
    st.pyplot(fig1)
    fig1.savefig("output/sales_by_product.png")

    # Sales over time
    df["Date"] = pd.to_datetime(df["Date"])
    sales_over_time = df.groupby("Date")["Total"].sum()
    fig2, ax2 = plt.subplots()
    sales_over_time.plot(ax=ax2, title="Sales Over Time")
    plt.xticks(rotation=45)
    plt.tight_layout()
    st.pyplot(fig2)
    fig2.savefig("output/sales_over_time.png")

    # PDF Export Button
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

    # Button to export
    if st.button("📄 Export PDF Report"):
        export_pdf(summary)
        st.success("✅ PDF report saved to output folder!")


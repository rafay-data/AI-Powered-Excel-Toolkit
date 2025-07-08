import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
import os

# Set up output folder
os.makedirs("output", exist_ok=True)

# App title
st.set_page_config(page_title="AI Excel Toolkit", layout="centered")
st.title("📊 AI-Powered Excel Business Toolkit")

# File uploader
uploaded_file = st.file_uploader("📥 Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        if "Total" not in df.columns or "Customer" not in df.columns or "Product" not in df.columns:
            st.error("❌ Missing required columns in Excel file. Ensure it includes: Total, Customer, Product, Date")
        else:
            # Business summary
            total_sales = df["Total"].sum()
            top_customers = df.groupby("Customer")["Total"].sum().sort_values(ascending=False).head(3)
            top_products = df.groupby("Product")["Total"].sum().sort_values(ascending=False).head(3)

            st.subheader("📈 Business Summary")
            st.write(f"💰 **Total Sales:** ${total_sales:,.2f}")
            st.write("👥 **Top Customers:**")
            st.dataframe(top_customers)
            st.write("📦 **Top Products:**")
            st.dataframe(top_products)

            # Store summary
            summary = {
                "Total Sales": total_sales,
                "Top Customers": top_customers.to_dict(),
                "Top Products": top_products.to_dict()
            }

            # Sales by product chart
            st.subheader("📊 Sales Visualizations")
            sales_by_product = df.groupby("Product")["Total"].sum()
            fig1, ax1 = plt.subplots()
            sales_by_product.plot(kind="bar", ax=ax1, title="Sales by Product", color="skyblue")
            plt.xticks(rotation=45)
            plt.tight_layout()
            st.pyplot(fig1)
            fig1.savefig("output/sales_by_product.png")

            # Sales over time chart
            df["Date"] = pd.to_datetime(df["Date"], errors='coerce')
            sales_over_time = df.groupby("Date")["Total"].sum()
            fig2, ax2 = plt.subplots()
            sales_over_time.plot(ax=ax2, title="Sales Over Time", color="orange")
            plt.xticks(rotation=45)
            plt.tight_layout()
            st.pyplot(fig2)
            fig2.savefig("output/sales_over_time.png")

            # PDF Export function
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

            # Export Button
            if st.button("📄 Export PDF Report", key="export_pdf_btn"):
                export_pdf(summary)
                st.success("✅ PDF report saved to output folder!")

                # Download button
                with open("output/business_report.pdf", "rb") as f:
                    st.download_button(
                        label="⬇️ Download PDF",
                        data=f,
                        file_name="business_report.pdf",
                        mime="application/pdf",
                        key="download_pdf_btn"
                    )

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")

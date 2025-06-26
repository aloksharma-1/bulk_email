import streamlit as st
import pandas as pd
import re
import smtplib
import os
import time
import datetime
from io import BytesIO
from fpdf import FPDF
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# ========== PDF Generator ==========
def generate_invoice_pdf(
    data: dict,
    filename: str,
    company_name="",
    company_address="",
    footer_note="",
    logo_bytes=None,
    company_email="",
    company_mobile="",
    signature_bytes=None
):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    # Add logo if available
    if logo_bytes:
        logo_path = "temp_logo.png"
        with open(logo_path, "wb") as f:
            f.write(logo_bytes)
        pdf.image(logo_path, x=10, y=8, w=33)
        os.remove(logo_path)
    # Company Info Header
    pdf.set_font("Arial", size=14)
    pdf.cell(200, 10, txt=company_name or "Payment Invoice", ln=1, align='C')

    pdf.set_font("Arial", size=10)
    if company_address:
        pdf.multi_cell(0, 6, txt=company_address, align='C')

# Combine email and phone in one line if available
    contact_line = ""
    if company_email:
        contact_line += f"Email: {company_email}  "
    if company_mobile:
        contact_line += f"Phone: {company_mobile}"
    if contact_line.strip():
        pdf.cell(200, 6, txt=contact_line.strip(), ln=1, align='C')

    pdf.ln(10)


    # ===== Table Header =====
    col_width_key = 60
    col_width_val = 120
    row_height = 8

    pdf.set_font("Arial", 'B', 11)
    pdf.cell(col_width_key, row_height, "Field", border=1)
    pdf.cell(col_width_val, row_height, "Value", border=1)
    pdf.ln(row_height)

    # ===== Table Rows =====
    pdf.set_font("Arial", '', 10)
    for key, value in data.items():
        pdf.cell(col_width_key, row_height, str(key), border=1)
        # Ensure encoding doesn't break the value
        value_str = str(value).encode('latin-1', 'replace').decode('latin-1')
        pdf.cell(col_width_val, row_height, value_str, border=1)
        pdf.ln(row_height)

    pdf.ln(10)

    # ===== Signature (Optional) =====
    if signature_bytes:
        sig_path = "temp_signature.png"
        with open(sig_path, "wb") as f:
            f.write(signature_bytes)
        pdf.image(sig_path, x=150, y=pdf.get_y(), w=40)
        os.remove(sig_path)
        pdf.ln(25)

    # ===== Footer Note (User Provided) =====
    if footer_note:
        pdf.set_font("Arial", style='I', size=9)
        pdf.multi_cell(0, 10, txt=footer_note, align='C')

    # ===== System-generated Disclaimer =====
    pdf.set_font("Arial", style='I', size=9)
    pdf.multi_cell(0, 10, txt="This is a system-generated receipt and does not require a physical signature.", align='C')

    # Save PDF
    pdf.output(filename)

# ========== Streamlit UI ==========
st.set_page_config(page_title="📧 Smart Bulk Email Sender", layout="centered")
st.title("📧 Smart Bulk Email Sender")
st.markdown("Upload your email content (manual or template) and CSV file to send personalized emails with optional attachments and invoices.")

with st.expander("📌 How to Configure Gmail SMTP (Required Step)", expanded=False):
    st.markdown("""
    To send emails via Gmail, you **must use an App Password** (not your regular Gmail password).  
    Follow these steps to generate and use it:

    ### 🔐 Step-by-Step Guide:

    1. **Enable 2-Step Verification** on your Google Account:  
       👉 [https://myaccount.google.com/security](https://myaccount.google.com/security)

    2. **Generate an App Password**:  
       👉 [https://myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)  
       - Select **Mail** as the app.  
       - Select **Other (Custom name)** and enter a name like "Bulk Email App".  
       - Click **Generate** and copy the 16-character password.

    3. **Paste that App Password** into the **Gmail App Password** field below.

    ✅ You're now ready to send emails securely!

    ---
    🔒 **Note:** Never share your App Password. It gives email-sending access to your account.
    """)

# Email config
sender_email = st.text_input("Your Gmail Address", placeholder="you@gmail.com")
app_password = st.text_input("Gmail App Password", type="password")
subject = st.text_input("Email Subject", value="🔔 Payment Reminder – Virtual Internship")

# Email content input
email_mode = st.radio("✍️ Email Content Mode", ["Use HTML Template", "Write Manually"])
html_template = ""
placeholders = set()

if email_mode == "Use HTML Template":
    template_file = st.file_uploader("📄 Upload HTML Email Template", type=["html"])
    if template_file:
        html_template = template_file.read().decode("utf-8")
        placeholders = set(re.findall(r'{(.*?)}', html_template))
else:
    html_template = st.text_area("✏️ Write your Email Body (Use {Name}, {Amount}, etc.)", height=250)
    placeholders = set(re.findall(r'{(.*?)}', html_template))

# File Uploads
csv_file = st.file_uploader("📊 Upload Recipient CSV File", type=["csv"])
attachment = st.file_uploader("📎 Upload Optional Attachment (PDF, DOCX, etc.)", type=["pdf", "jpg", "png", "docx", "xlsx", "zip"])
custom_invoice_template = st.file_uploader("📥 Upload Your Invoice PDF Template (Optional)", type=["pdf"])

# Sidebar Invoice Customization
st.sidebar.header("🧾 Invoice Customization")
company_name = st.sidebar.text_input("Company/Organization Name", "Virtual Internship")
company_address = st.sidebar.text_area("Company Address", "123, Some Street, India")
company_email = st.sidebar.text_input("Company Email", "")
company_mobile = st.sidebar.text_input("Company Mobile No.", "")
footer_note = st.sidebar.text_input("Footer Note", "Thank you for your payment.")
logo_file = st.sidebar.file_uploader("🖼️ Upload Logo (Optional)", type=["png", "jpg"])
signature_file = st.sidebar.file_uploader("✍️ Upload Signature (Optional)", type=["png", "jpg"])

# Custom Fields
st.sidebar.markdown("### ➕ Custom Fields")
num_fields = st.sidebar.number_input("Number of Extra Fields", min_value=0, max_value=20, step=1)
extra_fields = {}
for i in range(int(num_fields)):
    key = st.sidebar.text_input(f"Field {i+1} Name", key=f"fkey_{i}")
    val = st.sidebar.text_input(f"Field {i+1} Value", key=f"fval_{i}")
    if key:
        extra_fields[key] = val

# Email Scheduling
schedule_send = st.checkbox("⏰ Schedule email sending for later")
send_time = None
if schedule_send:
    send_time = st.time_input("Choose send time (today)", value=datetime.time(10, 0))

# ===== Load CSV and Setup PDF Fields Selection =====
if csv_file and html_template:
    df = pd.read_csv(csv_file)
    st.subheader("📋 CSV Columns Found:")
    st.code(", ".join(df.columns))

    all_available_fields = list(df.columns) + list(extra_fields.keys())
    st.sidebar.markdown("### 📄 Select Fields to Include in Invoice")
    invoice_fields = st.sidebar.multiselect(
        "Choose fields to include in the generated invoice PDF:",
        options=all_available_fields,
        default=["Name", "Amount", "Order_id", "Payment Date"]
    )

    missing_cols = placeholders - set(df.columns)
    if missing_cols:
        st.error(f"❌ Missing placeholders in CSV: {missing_cols}")
        st.stop()
    else:
        st.success(f"✅ All placeholders found in CSV: {', '.join(placeholders)}")

    email_cols = [col for col in df.columns if 'email' in col.lower()]
    email_column = st.selectbox("📨 Select recipient email column", email_cols)

    # Preview
    st.subheader("👁️ Email Preview (First Row)")
    preview_data = df.iloc[0].fillna("").to_dict()
    try:
        preview_text = html_template.format_map({
            k: str(preview_data.get(k, "[Not Provided]")) for k in placeholders
        })
        if email_mode == "Use HTML Template":
            st.components.v1.html(preview_text, height=400, scrolling=True)
        else:
            st.text_area("📬 Preview", preview_text, height=200)
    except Exception as e:
        st.error(f"❌ Error rendering preview: {e}")

    if st.button("📤 Send Emails Now"):
        try:
            if schedule_send and send_time:
                now = datetime.datetime.now()
                scheduled_dt = datetime.datetime.combine(now.date(), send_time)
                if scheduled_dt < now:
                    scheduled_dt += datetime.timedelta(days=1)

                if "scheduled_at" not in st.session_state:
                    st.session_state["scheduled_at"] = scheduled_dt.strftime("%Y-%m-%d %H:%M:%S")
                    st.warning(f"⏳ Emails scheduled for {scheduled_dt.strftime('%I:%M:%S %p')}. Keep the tab open.")
                    st.stop()
                else:
                    scheduled_time = datetime.datetime.strptime(st.session_state["scheduled_at"], "%Y-%m-%d %H:%M:%S")
                    now_check = datetime.datetime.now()
                    if now_check < scheduled_time:
                        time_left = str(scheduled_time - now_check).split('.')[0]
                        st.warning(f"⏰ Scheduled to send at {scheduled_time.strftime('%I:%M:%S %p')} — {time_left} remaining.")
                        st.stop()
                    else:
                        st.info("⏳ Scheduled time reached. Sending now...")

            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(sender_email, app_password)

            results = []
            progress = st.progress(0)
            status_msg = st.empty()
            logo_bytes = logo_file.read() if logo_file else None
            signature_bytes = signature_file.read() if signature_file else None

            for idx, (_, row) in enumerate(df.iterrows()):
                try:
                    data = row.fillna("").to_dict()

                    def safe_get(key):
                        val = data.get(key)
                        return str(val) if val and str(val).lower() != "nan" else "[Not Provided]"

                    safe_data = {k: safe_get(k) for k in placeholders}
                    merged_data = {**safe_data, **extra_fields}
                    receiver_email = safe_get(email_column)

                    email_body = html_template.format_map(safe_data)
                    email_body = email_body.replace("\\n", "\n").strip()

                    msg = MIMEMultipart("alternative")
                    msg["From"] = sender_email
                    msg["To"] = receiver_email
                    msg["Subject"] = subject

                    msg.attach(MIMEText(email_body, "html" if email_mode == "Use HTML Template" else "plain"))

                    # Optional Attachment
                    if attachment:
                        file_data = attachment.read()
                        file_part = MIMEApplication(file_data, Name=attachment.name)
                        file_part['Content-Disposition'] = f'attachment; filename="{attachment.name}"'
                        msg.attach(file_part)

                    # Invoice PDF (custom or generated)
                    if custom_invoice_template:
                        file_part = MIMEApplication(custom_invoice_template.read(), Name=custom_invoice_template.name)
                        file_part['Content-Disposition'] = f'attachment; filename="{custom_invoice_template.name}"'
                        msg.attach(file_part)
                    else:
                        pdf_filename = f"{safe_get('Name')}_{idx}.pdf"
                        filtered_invoice_data = {k: v for k, v in merged_data.items() if k in invoice_fields}

                        generate_invoice_pdf(
                            filtered_invoice_data,
                            pdf_filename,
                            company_name,
                            company_address,
                            footer_note,
                            logo_bytes,
                            company_email,
                            company_mobile,
                            signature_bytes
                        )
                        with open(pdf_filename, "rb") as f:
                            pdf_part = MIMEApplication(f.read(), Name=pdf_filename)
                            pdf_part['Content-Disposition'] = f'attachment; filename="{pdf_filename}"'
                            msg.attach(pdf_part)
                        os.remove(pdf_filename)

                    server.sendmail(sender_email, receiver_email, msg.as_string())
                    results.append({"email": receiver_email, "status": "Sent"})

                except Exception as e:
                    results.append({"email": data.get(email_column, "Unknown"), "status": f"Failed: {e}"})

                progress.progress((idx + 1) / len(df))
                status_msg.info(f"📬 Sending {idx + 1}/{len(df)} to {data.get(email_column, '')}")

            server.quit()
            st.success("✅ All emails processed!")

            result_df = pd.DataFrame(results)
            csv_buf = BytesIO()
            result_df.to_csv(csv_buf, index=False)

            st.download_button("📥 Download Email Report CSV", csv_buf.getvalue(), file_name="email_status.csv", mime="text/csv")
            st.subheader("📊 Send Summary")
            st.dataframe(result_df)

        except Exception as e:
            st.error(f"❌ SMTP Error: {e}")

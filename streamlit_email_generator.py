import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import zipfile
import tempfile
import os
import io
import base64
from typing import Dict, List, Optional
import re

# ---------------------------------------------------------------------
# Template Configuration
# ---------------------------------------------------------------------
TEMPLATE_CONFIG = {
    "Template 1": {
        "file": "email_template.txt",
        "required_columns": [
            'Feed Account Name', 'Feed Broker Name', 'Feed Account Number',
            'BRID', 'Employee Name', 'Email Address'
        ],
        "description": "Standard compliance review template",
        "column_mapping": {
            "account_name": "Feed Account Name",
            "broker_name": "Feed Broker Name", 
            "account_number": "Feed Account Number",
            "employee_id": "BRID",
            "employee_name": "Employee Name",
            "email": "Email Address"
        }
    },
    "Template 2": {
        "file": "email_template2.txt",
        "required_columns": [
            'Account Name', 'Broker Name', 'Account Number',
            'Employee ID', 'Employee Name', 'Email'
        ],
        "description": "Alternative format template",
        "column_mapping": {
            "account_name": "Account Name",
            "broker_name": "Broker Name",
            "account_number": "Account Number", 
            "employee_id": "Employee ID",
            "employee_name": "Employee Name",
            "email": "Email"
        }
    },
    "Template 3": {
        "file": "email_template3.txt",
        "required_columns": [
            'Client Name', 'Institution', 'Account ID',
            'Staff ID', 'Staff Name', 'Contact Email'
        ],
        "description": "Client-focused template",
        "column_mapping": {
            "account_name": "Client Name",
            "broker_name": "Institution",
            "account_number": "Account ID",
            "employee_id": "Staff ID", 
            "employee_name": "Staff Name",
            "email": "Contact Email"
        }
    }
}

# Default email templates
DEFAULT_TEMPLATES = {
    "email_template.txt": """<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Compliance Review - BRID {brid}</title>
    <style>
        body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; line-height: 1.4; max-width: 800px; }
        .header { background-color: #A0B2C4; border-top: solid 1px #80B2C4; padding: 10px 20px; border-radius: 6px; }
        .header-text { color: #101F2A; font-size: 14pt; font-weight: bold; }
        .accounts-table { border-collapse: collapse; width: 100%; margin: 10px 0; }
        .accounts-table th, .accounts-table td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        .accounts-table th { background-color: #f2f2f2; }
        .footer { margin-top: 20px; padding: 10px; background-color: #E7E7E7; border-radius: 4px; }
    </style>
</head>
<body>
    <div class="header">
        <span class="header-text">URGENT: Compliance Review Required</span>
    </div>

    <p><br>Hi {employee_name},<br><br>

    As a <a href="#" style="text-decoration: none; color: #104C71;">Designated Control Employee</a>, you are required to ensure that all <a href="#" style="text-decoration: none; color: #104C71;">Personal Trading</a> accounts are appropriately reflected in your Compliance disclosures.
    <br><br>

    Please review the following account(s) and confirm they are correctly disclosed:
    <br><br>

    {accounts_html}
    <br>

    <p style="font-size: 10pt; color: #643E04; font-weight: bold;">Response Required By: {due_date}</p>

    <div class="footer">
        <p style="font-size: 10pt; margin: 0;">Compliance Team<br>
        Risk & Compliance Division</p>
    </div>
</body>
</html>""",

    "email_template2.txt": """<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Account Review - ID {brid}</title>
    <style>
        body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; line-height: 1.4; max-width: 800px; }
        .header { background-color: #A0B2C4; border-top: solid 1px #80B2C4; padding: 10px 20px; border-radius: 6px; }
        .header-text { color: #101F2A; font-size: 14pt; font-weight: bold; }
        .accounts-table { border-collapse: collapse; width: 100%; margin: 10px 0; }
        .accounts-table th, .accounts-table td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        .accounts-table th { background-color: #f2f2f2; }
    </style>
</head>
<body>
    <div class="header">
        <span class="header-text">ACTION REQUIRED: Account Review</span>
    </div>

    <p><br>Dear {employee_name},<br><br>

    This is a notification regarding your registered trading accounts. Please review and confirm the accuracy of the following information:
    <br><br>

    {accounts_html}
    <br>

    <p style="font-size: 10pt; color: #643E04; font-weight: bold;">Please respond by: {due_date}</p>

    <p style="font-size: 10pt;">Best regards,<br>Compliance Operations Team</p>
</body>
</html>""",

    "email_template3.txt": """<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Client Account Review - {brid}</title>
    <style>
        body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; line-height: 1.4; max-width: 800px; }
        .header { background-color: #A0B2C4; border-top: solid 1px #80B2C4; padding: 10px 20px; border-radius: 6px; }
        .header-text { color: #101F2A; font-size: 14pt; font-weight: bold; }
        .accounts-table { border-collapse: collapse; width: 100%; margin: 10px 0; }
        .accounts-table th, .accounts-table td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        .accounts-table th { background-color: #f2f2f2; }
    </style>
</head>
<body>
    <div class="header">
        <span class="header-text">NOTICE: Client Account Review</span>
    </div>

    <p><br>Hello {employee_name},<br><br>

    As part of our regular compliance procedures, please review the following client accounts under your management:
    <br><br>

    {accounts_html}
    <br>

    <p style="font-size: 10pt; color: #643E04; font-weight: bold;">Review deadline: {due_date}</p>

    <p style="font-size: 10pt;">Thank you for your attention to this matter.<br>Client Services & Compliance</p>
</body>
</html>"""
}

# ---------------------------------------------------------------------
# Utility Functions
# ---------------------------------------------------------------------
def load_email_template(template_name: str) -> str:
    """Load email template based on selected template"""
    try:
        template_file = TEMPLATE_CONFIG[template_name]["file"]
        return DEFAULT_TEMPLATES[template_file]
    except KeyError:
        st.error(f"❌ Email template '{template_name}' not found!")
        return ""

def get_due_date() -> str:
    """Calculate due date (current date + 7 days)"""
    due_date = datetime.now() + timedelta(days=7)
    day = due_date.day

    # Add ordinal suffix
    if 10 <= day % 100 <= 20:
        suffix = 'th'
    else:
        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')

    return due_date.strftime(f"%d{suffix} %b %y").replace(f"{day:02d}", f"{day}")

def get_column_mapping(template_name: str) -> Dict[str, str]:
    """Get column mapping for selected template"""
    return TEMPLATE_CONFIG[template_name]["column_mapping"]

def build_accounts_html(group: pd.DataFrame, column_mapping: Dict[str, str]) -> str:
    """Build HTML table for accounts"""
    html_parts = [
        '<table class="accounts-table">',
        '<thead>',
        '<tr>',
        '<th>Account Name</th>',
        '<th>Broker/Institution</th>',
        '<th>Account Number</th>',
        '</tr>',
        '</thead>',
        '<tbody>'
    ]

    for _, row in group.iterrows():
        html_parts.extend([
            '<tr>',
            f'<td>{row[column_mapping["account_name"]]}</td>',
            f'<td>{row[column_mapping["broker_name"]]}</td>',
            f'<td>{row[column_mapping["account_number"]]}</td>',
            '</tr>'
        ])

    html_parts.extend([
        '</tbody>',
        '</table>'
    ])

    return '\n'.join(html_parts)

def clean_dataframe(df: pd.DataFrame, template_name: str) -> pd.DataFrame:
    """Clean and validate dataframe"""
    required_cols = TEMPLATE_CONFIG[template_name]["required_columns"]

    # Check for missing columns
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns: {', '.join(missing_cols)}")

    # Remove rows with missing critical data
    column_mapping = get_column_mapping(template_name)
    critical_cols = [column_mapping["employee_id"], column_mapping["employee_name"], column_mapping["email"]]

    df_clean = df.dropna(subset=critical_cols)

    # Convert to string and strip whitespace
    for col in df_clean.columns:
        if df_clean[col].dtype == 'object':
            df_clean[col] = df_clean[col].astype(str).str.strip()

    return df_clean

def create_eml_file(temp_dir: str, employee_id: str, to_email: str, cc_email: str, 
                   subject: str, body_html: str) -> Optional[str]:
    """Create EML format email file with X-Unsent:1 header"""
    try:
        filename = os.path.join(temp_dir, f"compliance_email_{employee_id}.eml")

        # Create EML content with proper headers
        eml_content_parts = [
            "X-Unsent: 1",
            "From: compliance@company.com",
            f"To: {to_email}",
            f"Cc: {cc_email}",
            f"Subject: {subject}",
            f"Date: {datetime.now().strftime('%a, %d %b %Y %H:%M:%S +0000')}",
            "MIME-Version: 1.0",
            "Content-Type: text/html; charset=utf-8",
            "Content-Transfer-Encoding: quoted-printable",
            "",
            body_html
        ]

        eml_content = '\n'.join(eml_content_parts)

        with open(filename, 'w', encoding='utf-8') as f:
            f.write(eml_content)

        return filename
    except Exception as e:
        st.warning(f"Failed to create EML for ID {employee_id}: {str(e)}")
        return None

def export_emails_as_files(df: pd.DataFrame, cc_email: str, template_name: str) -> Optional[bytes]:
    """Export emails as EML files and create ZIP"""
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            column_mapping = get_column_mapping(template_name)
            template_content = load_email_template(template_name)
            due_date = get_due_date()

            # Group by employee
            grouped = df.groupby([column_mapping["employee_id"], 
                                column_mapping["employee_name"], 
                                column_mapping["email"]])

            created_files = []

            for (employee_id, employee_name, email), group in grouped:
                # Build accounts HTML for this employee
                accounts_html = build_accounts_html(group, column_mapping)

                # Fill template placeholders
                email_body = template_content.format(
                    employee_name=employee_name,
                    brid=employee_id,
                    due_date=due_date,
                    accounts_html=accounts_html
                )

                subject = f"Compliance Review - BRID {employee_id}"

                # Create EML file
                eml_file = create_eml_file(
                    temp_dir, employee_id, email, cc_email, subject, email_body
                )

                if eml_file:
                    created_files.append(eml_file)

            # Create ZIP file
            if created_files:
                zip_filename = os.path.join(temp_dir, "compliance_emails.zip")
                with zipfile.ZipFile(zip_filename, 'w') as zip_file:
                    for file_path in created_files:
                        zip_file.write(file_path, os.path.basename(file_path))

                # Read ZIP file content
                with open(zip_filename, 'rb') as f:
                    zip_content = f.read()

                return zip_content

            return None

    except Exception as e:
        st.error(f"Error creating email files: {str(e)}")
        return None

def create_sample_data(template_name: str) -> pd.DataFrame:
    """Create sample data for selected template"""
    column_mapping = get_column_mapping(template_name)

    sample_data = []
    if template_name == "Template 1":
        sample_data = [
            {
                column_mapping["account_name"]: "John Doe Trading Account",
                column_mapping["broker_name"]: "abc Securities", 
                column_mapping["account_number"]: "123456789",
                column_mapping["employee_id"]: "BD001",
                column_mapping["employee_name"]: "John Doe",
                column_mapping["email"]: "john.doe@company.com"
            },
            {
                column_mapping["account_name"]: "John Doe Investment",
                column_mapping["broker_name"]: "Morgan Stanley",
                column_mapping["account_number"]: "123456790", 
                column_mapping["employee_id"]: "BD001",
                column_mapping["employee_name"]: "John Doe",
                column_mapping["email"]: "john.doe@company.com"
            },
            {
                column_mapping["account_name"]: "Jane Smith Account",
                column_mapping["broker_name"]: "Goldman Sachs",
                column_mapping["account_number"]: "987654321",
                column_mapping["employee_id"]: "BD002", 
                column_mapping["employee_name"]: "Jane Smith",
                column_mapping["email"]: "jane.smith@company.com"
            }
        ]
    elif template_name == "Template 2":
        sample_data = [
            {
                column_mapping["account_name"]: "Personal Trading Account",
                column_mapping["broker_name"]: "Interactive Brokers",
                column_mapping["account_number"]: "IB123456",
                column_mapping["employee_id"]: "EMP001",
                column_mapping["employee_name"]: "Alice Johnson", 
                column_mapping["email"]: "alice.johnson@company.com"
            },
            {
                column_mapping["account_name"]: "Investment Portfolio",
                column_mapping["broker_name"]: "Charles Schwab",
                column_mapping["account_number"]: "CS789012",
                column_mapping["employee_id"]: "EMP002",
                column_mapping["employee_name"]: "Bob Wilson",
                column_mapping["email"]: "bob.wilson@company.com"
            }
        ]
    else:  # Template 3
        sample_data = [
            {
                column_mapping["account_name"]: "Corporate Client A",
                column_mapping["broker_name"]: "JPMorgan Chase",
                column_mapping["account_number"]: "JPM456789", 
                column_mapping["employee_id"]: "ST001",
                column_mapping["employee_name"]: "Carol Davis",
                column_mapping["email"]: "carol.davis@company.com"
            },
            {
                column_mapping["account_name"]: "Institutional Client B",
                column_mapping["broker_name"]: "Citibank",
                column_mapping["account_number"]: "CITI123456",
                column_mapping["employee_id"]: "ST002",
                column_mapping["employee_name"]: "David Brown",
                column_mapping["email"]: "david.brown@company.com"
            }
        ]

    return pd.DataFrame(sample_data)

# ---------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------
def main():
    st.set_page_config(
        page_title="Email Template Generator",
        page_icon="✉️",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # Custom CSS
    st.markdown("""
    <style>
    .main-header {
        background: linear-gradient(90deg, #1f4e79 0%, #2d5aa0 100%);
        padding: 2rem;
        border-radius: 10px;
        color: white;
        margin-bottom: 2rem;
        text-align: center;
    }
    .step-header {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #1f4e79;
        margin: 1rem 0;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #e7f3ff;
        border: 1px solid #b3d7ff;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    </style>
    """, unsafe_allow_html=True)

    # Header
    st.markdown("""
    <div class="main-header">
        <h1>✉️ Email Template Generator</h1>
        <p>Generate personalized compliance emails with X-Unsent header for editing</p>
    </div>
    """, unsafe_allow_html=True)

    # Sidebar
    with st.sidebar:
        st.header("📋 Application Guide")
        st.markdown("""
        **Step-by-Step Process:**
        1. Select email template
        2. Upload Excel file with data
        3. Configure email settings
        4. Generate and download emails

        **Supported Formats:**
        - Excel files (.xlsx, .xls)
        - EML email format output
        - ZIP file for batch download

        **Features:**
        - Cross-platform compatibility
        - X-Unsent:1 header for drafts
        - HTML email templates
        - Batch processing
        """)

        st.divider()

        # Quick actions
        st.subheader("🚀 Quick Actions")
        if st.button("📥 Download Sample Data", use_container_width=True):
            if 'selected_template' in st.session_state:
                sample_df = create_sample_data(st.session_state['selected_template'])
                csv = sample_df.to_csv(index=False)
                st.download_button(
                    label="📄 Download Sample CSV",
                    data=csv,
                    file_name=f"sample_data_{st.session_state['selected_template'].lower().replace(' ', '_')}.csv",
                    mime="text/csv"
                )
            else:
                st.info("Please select a template first")

    # Step 1: Template Selection
    st.markdown('<div class="step-header"><h2>📋 Step 1: Select Email Template</h2></div>', 
                unsafe_allow_html=True)

    template_options = list(TEMPLATE_CONFIG.keys())
    selected_template = st.selectbox(
        "Choose an email template:",
        options=template_options,
        index=0,
        key="selected_template"
    )

    template_info = TEMPLATE_CONFIG[selected_template]

    col1, col2 = st.columns([2, 1])

    with col1:
        with st.expander("📝 Template Details & Requirements", expanded=True):
            st.markdown(f"**Description:** {template_info['description']}")
            st.markdown("**Required Excel Columns:**")
            for col in template_info['required_columns']:
                st.markdown(f"• `{col}`")

    with col2:
        st.markdown('<div class="info-box">', unsafe_allow_html=True)
        st.markdown("**💡 Template Info**")
        st.markdown(f"**File:** {template_info['file']}")
        st.markdown(f"**Columns:** {len(template_info['required_columns'])}")
        st.markdown('</div>', unsafe_allow_html=True)

    # Step 2: File Upload
    st.markdown('<div class="step-header"><h2>📁 Step 2: Upload Excel File</h2></div>', 
                unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "Choose your Excel file containing employee data:",
        type=['xlsx', 'xls'],
        help="Upload an Excel file with the required columns for the selected template"
    )

    if uploaded_file is not None:
        try:
            # Load data
            df = pd.read_excel(uploaded_file)

            col1, col2, col3 = st.columns([1, 1, 1])
            with col1:
                st.metric("Total Records", len(df))
            with col2:
                st.metric("Columns", len(df.columns))
            with col3:
                unique_employees = len(df[template_info["column_mapping"]["employee_id"]].unique()) if template_info["column_mapping"]["employee_id"] in df.columns else 0
                st.metric("Unique Employees", unique_employees)

            # Data preview
            with st.expander("👀 Data Preview", expanded=False):
                st.dataframe(df.head(10), use_container_width=True)

            # Validate required columns
            required_cols = template_info["required_columns"]
            missing_cols = [col for col in required_cols if col not in df.columns]

            if missing_cols:
                st.error(f"❌ **Missing Required Columns:** {', '.join(missing_cols)}")
                st.info("💡 Please check your Excel file or select a different template")

                # Show current columns vs required
                col1, col2 = st.columns(2)
                with col1:
                    st.subheader("📋 Your File Columns")
                    for col in df.columns:
                        st.write(f"✅ {col}")

                with col2:
                    st.subheader("📋 Required Columns")
                    for col in required_cols:
                        icon = "✅" if col in df.columns else "❌"
                        st.write(f"{icon} {col}")
            else:
                # Data validation successful
                st.markdown('<div class="success-box">✅ <strong>Data validation passed!</strong> All required columns found.</div>', 
                          unsafe_allow_html=True)

                # Clean data
                df_clean = clean_dataframe(df, selected_template)

                if len(df_clean) < len(df):
                    st.warning(f"⚠️ Removed {len(df) - len(df_clean)} rows with missing critical data")

                unique_employees = len(df_clean[template_info["column_mapping"]["employee_id"]].unique())

                # Step 3: Email Configuration
                st.markdown('<div class="step-header"><h2>⚙️ Step 3: Email Configuration</h2></div>', 
                           unsafe_allow_html=True)

                cc_email = st.text_input(
                    "CC Email Address:",
                    value="compliance@company.com",
                    help="Email address to CC on all generated emails"
                )

                col1, col2 = st.columns(2)
                with col1:
                    st.info(f"📊 **Processing Summary**\n\n"
                           f"• **Records:** {len(df_clean)}\n"
                           f"• **Employees:** {unique_employees}\n"
                           f"• **Template:** {selected_template}")

                with col2:
                    due_date = get_due_date()
                    st.info(f"📅 **Email Details**\n\n"
                           f"• **Due Date:** {due_date}\n"
                           f"• **Format:** EML (editable)\n"
                           f"• **Header:** X-Unsent:1")

                # Step 4: Generate Emails
                st.markdown('<div class="step-header"><h2>🚀 Step 4: Generate & Download Emails</h2></div>', 
                           unsafe_allow_html=True)

                # Preview section
                with st.expander("👁️ Email Preview", expanded=False):
                    st.subheader("Sample Email Content")

                    # Get first employee for preview
                    column_mapping = get_column_mapping(selected_template)
                    first_employee = df_clean.iloc[0]
                    employee_id = first_employee[column_mapping["employee_id"]]
                    employee_name = first_employee[column_mapping["employee_name"]]

                    # Get accounts for this employee
                    employee_accounts = df_clean[df_clean[column_mapping["employee_id"]] == employee_id]
                    accounts_html = build_accounts_html(employee_accounts, column_mapping)

                    template_content = load_email_template(selected_template)
                    preview_body = template_content.format(
                        employee_name=employee_name,
                        brid=employee_id,
                        due_date=due_date,
                        accounts_html=accounts_html
                    )

                    st.markdown("**Subject:** " + f"Compliance Review - BRID {employee_id}")
                    st.markdown("**To:** " + first_employee[column_mapping["email"]])
                    st.markdown("**CC:** " + cc_email)

                    # Show HTML preview
                    st.components.v1.html(preview_body, height=400, scrolling=True)

                # Generate button
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    if st.button("🎯 Generate & Download Email Files", 
                               type="primary", 
                               use_container_width=True):

                        with st.spinner(f"🔄 Creating {unique_employees} email files using {selected_template}..."):
                            zip_content = export_emails_as_files(df_clean, cc_email, selected_template)

                            if zip_content:
                                st.markdown('<div class="success-box">✅ <strong>Success!</strong> Email files created successfully.</div>', 
                                          unsafe_allow_html=True)

                                # Download button
                                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                filename = f"compliance_emails_{timestamp}.zip"

                                st.download_button(
                                    label="📥 Download ZIP File",
                                    data=zip_content,
                                    file_name=filename,
                                    mime="application/zip",
                                    use_container_width=True
                                )

                                st.success(f"🎉 Created {unique_employees} email files!")
                                st.info("💡 **Next Steps:**\n"
                                       "1. Extract the ZIP file\n"
                                       "2. Open .eml files in your email client\n"
                                       "3. Edit as needed (X-Unsent:1 header allows editing)\n"
                                       "4. Send the emails")
                            else:
                                st.error("❌ Failed to create email files. Please try again.")

        except Exception as e:
            st.error(f"❌ **Error loading file:** {str(e)}")
            st.info("💡 Please try using a different Excel file or check the data format")

    else:
        # Show sample data option when no file uploaded
        st.info("📤 **No file uploaded yet.** Upload an Excel file to get started, or try with sample data.")

        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            if st.button("📋 Try with Sample Data", type="secondary", use_container_width=True):
                sample_df = create_sample_data(selected_template)
                st.session_state['sample_data'] = sample_df
                st.rerun()

        # Show sample data if generated
        if 'sample_data' in st.session_state:
            st.subheader("📊 Sample Data Preview")
            st.dataframe(st.session_state['sample_data'], use_container_width=True)

            # Process sample data
            df_clean = st.session_state['sample_data']
            unique_employees = len(df_clean[template_info["column_mapping"]["employee_id"]].unique())

            st.markdown('<div class="step-header"><h2>⚙️ Step 3: Email Configuration</h2></div>', 
                       unsafe_allow_html=True)

            cc_email = st.text_input(
                "CC Email Address:",
                value="compliance@company.com",
                help="Email address to CC on all generated emails",
                key="sample_cc"
            )

            st.markdown('<div class="step-header"><h2>🚀 Step 4: Generate & Download Emails</h2></div>', 
                       unsafe_allow_html=True)

            if st.button("🎯 Generate Sample Emails", type="primary"):
                with st.spinner("Creating sample email files..."):
                    zip_content = export_emails_as_files(df_clean, cc_email, selected_template)

                    if zip_content:
                        st.success("✅ Sample emails created successfully!")

                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"sample_emails_{timestamp}.zip"

                        st.download_button(
                            label="📥 Download Sample Emails ZIP",
                            data=zip_content,
                            file_name=filename,
                            mime="application/zip"
                        )

if __name__ == "__main__":
    main()

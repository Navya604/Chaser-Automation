import os
import re
import traceback
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
import pythoncom
import win32com.client as win32


# =========================================================
# SHEET NAMES
# =========================================================

WORKFLOW_SHEET = "Sheet1"
HELPER_SHEET = "Sheet1"
CONFIG_SHEET = "Sheet1"


# =========================================================
# WORKFLOW COLUMNS
# =========================================================

WF_REGION = "Reg"
WF_NAV_BUCKET = "NAV Bucket"
WF_MTD_BUCKET = "MTD Bucket"
WF_COVERAGE = "Coverage"
WF_WEB = "Web?"
WF_DNC = "DNC"
WF_FREQ = "NAV Freq."
WF_CLIENT_CONTACT = "Client Contact"
WF_FUND_KEY = "Fund UCN"
WF_FUND_NAME = "Fund Name"
WF_IA_NAME = "IA Name"

# CHANGE THESE TO YOUR REAL HEADERS
WF_AS = "AS"
WF_AK = "AK"


# =========================================================
# CREDIT HELPER COLUMNS
# =========================================================

HELPER_FUND_KEY = "Fund UCN"
HELPER_CREDIT_CONTACT = "Credit Contact"


# =========================================================
# CONFIG FILE COLUMNS
# =========================================================

CFG_NAME = "Name"
CFG_EMAIL = "Email"
CFG_TYPE = "Type"
CFG_TITLE = "Title"
CFG_LOCATION = "Location"


# =========================================================
# FILTER VALUES
# =========================================================

ALLOWED_REGIONS = ["NAHF", "LATAM"]

NAV_BUCKET_ALLOWED = ["1-10", "11-30"]

MTD_BUCKET_ALLOWED = ["1-10"]

COVERAGE_EXCLUDE = [
    "BLOOMBERG",
    "PYTHON",
    "DAY NAV 2",
    "PYTHON WEB",
    "PYTHON - WEB"
]

PASS_SHEET = "Pass"
FAIL_SHEET = "Fail"


# =========================================================
# GLOBAL VARIABLES
# =========================================================

workflow_path = ""
helper_path = ""
config_path = ""
output_folder = ""
validation_path = ""

config_df_cache = pd.DataFrame()


# =========================================================
# COMMON FUNCTIONS
# =========================================================

def clean_text(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def normalize_upper(x):
    return clean_text(x).upper()


def today_str():
    return datetime.today().strftime("%m/%d/%Y")


def normalize_email_string(value):

    text = clean_text(value)

    if not text:
        return ""

    parts = re.split(r"[;,]+", text)

    parts = [p.strip() for p in parts if p.strip()]

    return "; ".join(parts)


def combine_emails(*values):

    emails = []

    for val in values:

        val = normalize_email_string(val)

        if val:
            emails.extend(
                [e.strip() for e in val.split(";") if e.strip()]
            )

    final = []
    seen = set()

    for email in emails:

        low = email.lower()

        if low not in seen:
            seen.add(low)
            final.append(email)

    return "; ".join(final)


def read_excel(path, sheet):
    return pd.read_excel(
        path,
        sheet_name=sheet,
        dtype=str
    ).fillna("")


def check_columns(df, required_cols, file_name):

    missing = [c for c in required_cols if c not in df.columns]

    if missing:
        raise Exception(
            f"Missing columns in {file_name}: {missing}"
        )


def info(msg):
    messagebox.showinfo(
        "HF NAV Chaser Automation",
        msg
    )


def error(msg):
    messagebox.showerror(
        "HF NAV Chaser Automation",
        msg
    )


# =========================================================
# CONFIG FILE LOGIC
# =========================================================

def get_config_email(email_type):

    global config_df_cache

    rows = config_df_cache[
        config_df_cache[CFG_TYPE]
        .astype(str)
        .str.upper() == email_type.upper()
    ]

    if rows.empty:
        return ""

    return normalize_email_string(
        rows.iloc[0][CFG_EMAIL]
    )


def load_senders():

    global config_df_cache

    try:

        if config_df_cache.empty and config_path:

            config_df_cache = read_excel(
                config_path,
                CONFIG_SHEET
            )

        sender_rows = config_df_cache[
            config_df_cache[CFG_TYPE]
            .astype(str)
            .str.upper() == "SENDER"
        ]

        sender_names = (
            sender_rows[CFG_NAME]
            .dropna()
            .astype(str)
            .tolist()
        )

        sender_dropdown["values"] = sender_names

        if sender_names:
            sender_var.set(sender_names[0])

    except Exception:
        pass


def get_selected_sender_details():

    global config_df_cache

    selected_name = clean_text(sender_var.get())

    sender_rows = config_df_cache[
        (config_df_cache[CFG_TYPE]
         .astype(str)
         .str.upper() == "SENDER")
        &
        (config_df_cache[CFG_NAME]
         .astype(str)
         .str.strip() == selected_name)
    ]

    if sender_rows.empty:

        return {
            "name": selected_name,
            "email": "",
            "title": "",
            "location": ""
        }

    row = sender_rows.iloc[0]

    return {
        "name": clean_text(row.get(CFG_NAME, "")),
        "email": clean_text(row.get(CFG_EMAIL, "")),
        "title": clean_text(row.get(CFG_TITLE, "")),
        "location": clean_text(row.get(CFG_LOCATION, ""))
    }


def make_signature(sender_details):

    name = sender_details.get("name", "")
    email = sender_details.get("email", "")
    title = sender_details.get("title", "")
    location = sender_details.get("location", "")

    return f"""
    Best Regards,<br><br>

    <b>{name}</b> |
    {title} |
    J.P. Morgan |
    {location} |<br>

    <a href="mailto:{email}">
    {email}
    </a>
    """


# =========================================================
# PROFESSIONAL HTML TABLE
# =========================================================

def build_professional_table(table_df):

    html = """
    <table style="
        border-collapse: collapse;
        width: 100%;
        font-family: Calibri;
        font-size: 11pt;
    ">
    <thead>
    <tr style="
        background-color:#1F4E79;
        color:white;
    ">
    """

    for col in table_df.columns:

        html += f"""
        <th style="
            border:1px solid #A6A6A6;
            padding:8px;
            text-align:center;
            font-weight:bold;
        ">
        {col}
        </th>
        """

    html += """
    </tr>
    </thead>
    <tbody>
    """

    for i, (_, row) in enumerate(table_df.iterrows()):

        bg_color = "#F2F6FA" if i % 2 == 0 else "#FFFFFF"

        html += f"""
        <tr style="background-color:{bg_color};">
        """

        for col in table_df.columns:

            value = "" if pd.isna(row[col]) else str(row[col])

            html += f"""
            <td style="
                border:1px solid #A6A6A6;
                padding:7px;
                text-align:left;
            ">
            {value}
            </td>
            """

        html += "</tr>"

    html += """
    </tbody>
    </table>
    """

    return html


def build_email_html(body_text,
                     table_html,
                     signature_html):

    return f"""
    <html>
    <body style="
        font-family:Calibri;
        font-size:11pt;
    ">

    <p>Dear Team,</p>

    <p>{body_text}</p>

    {table_html}

    <p>
    Please fill in the required details
    and revert at your earliest convenience.
    </p>

    <p>
    If the information has already been shared,
    please ignore this request.
    </p>

    <p>{signature_html}</p>

    </body>
    </html>
    """


# =========================================================
# GUI CONTROL
# =========================================================

def on_chaser_type_change(event=None):

    chaser_type = chaser_type_var.get()

    if chaser_type == "Chaser 1":

        helper_entry.config(state="disabled")
        helper_button.config(state="disabled")

        helper_file_var.set("")

    else:

        helper_entry.config(state="normal")
        helper_button.config(state="normal")


# =========================================================
# VALIDATION LOGIC
# =========================================================

def create_validation_file():

    global workflow_path
    global helper_path
    global config_path
    global output_folder
    global validation_path
    global config_df_cache

    try:

        if not workflow_path:
            error("Please select Workflow file.")
            return

        if not config_path:
            error("Please select Config file.")
            return

        if not output_folder:
            error("Please select Output folder.")
            return

        chaser_type = chaser_type_var.get()

        frequency_type = frequency_var.get()

        nav_date_input = nav_date_var.get().strip()

        if not nav_date_input:
            error("Please enter NAV Date.")
            return

        if chaser_type == "Chaser 2" and not helper_path:

            error(
                "Please select Credit Helper file "
                "for Chaser 2."
            )
            return

        workflow = read_excel(
            workflow_path,
            WORKFLOW_SHEET
        )

        config_df_cache = read_excel(
            config_path,
            CONFIG_SHEET
        )

        # =====================================================
        # COLUMN CHECKS
        # =====================================================

        check_columns(
            workflow,
            [
                WF_REGION,
                WF_NAV_BUCKET,
                WF_MTD_BUCKET,
                WF_COVERAGE,
                WF_WEB,
                WF_DNC,
                WF_FREQ,
                WF_CLIENT_CONTACT,
                WF_FUND_KEY,
                WF_FUND_NAME,
                WF_IA_NAME,
                WF_AK,
            ],
            "Workflow"
        )

        if chaser_type == "Chaser 2":

            check_columns(
                workflow,
                [WF_AS],
                "Workflow"
            )

        # =====================================================
        # CLEANING
        # =====================================================

        df = workflow.copy()

        df[WF_REGION] = (
            df[WF_REGION]
            .apply(normalize_upper)
        )

        df[WF_NAV_BUCKET] = (
            df[WF_NAV_BUCKET]
            .apply(clean_text)
        )

        df[WF_MTD_BUCKET] = (
            df[WF_MTD_BUCKET]
            .apply(clean_text)
        )

        df[WF_COVERAGE] = (
            df[WF_COVERAGE]
            .apply(normalize_upper)
        )

        df[WF_WEB] = (
            df[WF_WEB]
            .apply(clean_text)
        )

        df[WF_DNC] = (
            df[WF_DNC]
            .apply(clean_text)
        )

        df[WF_FREQ] = (
            df[WF_FREQ]
            .apply(normalize_upper)
        )

        # =====================================================
        # CHASER 1
        # =====================================================

        if chaser_type == "Chaser 1":

            df_filtered = df[
                (df[WF_REGION]
                 .isin(ALLOWED_REGIONS))
                &
                (df[WF_NAV_BUCKET]
                 .isin(NAV_BUCKET_ALLOWED))
                &
                (df[WF_MTD_BUCKET]
                 .isin(MTD_BUCKET_ALLOWED))
                &
                (~df[WF_COVERAGE]
                 .isin(COVERAGE_EXCLUDE))
                &
                (df[WF_WEB] == "")
                &
                (df[WF_DNC] == "")
                &
                (
                    df[WF_FREQ]
                    == frequency_type.upper()
                )
            ].copy()

            df_filtered[HELPER_CREDIT_CONTACT] = ""

        # =====================================================
        # CHASER 2
        # =====================================================

        else:

            df[WF_AS] = (
                df[WF_AS]
                .apply(normalize_upper)
            )

            df_filtered = df[
                (df[WF_REGION]
                 .isin(ALLOWED_REGIONS))
                &
                (df[WF_AS] == "Y")
                &
                (
                    df[WF_FREQ]
                    == frequency_type.upper()
                )
            ].copy()

            helper = read_excel(
                helper_path,
                HELPER_SHEET
            )

            check_columns(
                helper,
                [
                    HELPER_FUND_KEY,
                    HELPER_CREDIT_CONTACT,
                ],
                "Credit Helper"
            )

            df_filtered = df_filtered.merge(
                helper,
                left_on=WF_FUND_KEY,
                right_on=HELPER_FUND_KEY,
                how="left"
            )

        # =====================================================
        # NO RECORDS
        # =====================================================

        if df_filtered.empty:

            error(
                "No records found after applying filters."
            )

            return

        # =====================================================
        # CONFIG EMAILS
        # =====================================================

        jpm_nav_email = get_config_email("JPM_NAV")

        hfc_brazil_email = get_config_email("LATAM")

        # =====================================================
        # ROUTING LOGIC
        # =====================================================

        def build_to(row):

            client = normalize_email_string(
                row.get(WF_CLIENT_CONTACT, "")
            )

            return client

        def build_cc(row):

            region = normalize_upper(
                row.get(WF_REGION, "")
            )

            credit = normalize_email_string(
                row.get(
                    HELPER_CREDIT_CONTACT,
                    ""
                )
            )

            # CHASER 1

            if chaser_type == "Chaser 1":

                cc = combine_emails(
                    jpm_nav_email
                )

            # CHASER 2

            else:

                cc = combine_emails(
                    credit,
                    jpm_nav_email
                )

            # LATAM

            if region == "LATAM":

                cc = combine_emails(
                    cc,
                    hfc_brazil_email
                )

            return cc

        df_filtered["TO_ADDRESS_FINAL"] = (
            df_filtered.apply(build_to, axis=1)
        )

        df_filtered["CC_ADDRESS_FINAL"] = (
            df_filtered.apply(build_cc, axis=1)
        )

        # =====================================================
        # VALIDATION
        # =====================================================

        def validation_status(row):

            client = normalize_email_string(
                row.get(WF_CLIENT_CONTACT, "")
            )

            credit = normalize_email_string(
                row.get(
                    HELPER_CREDIT_CONTACT,
                    ""
                )
            )

            if not client:
                return (
                    "FAIL - Missing Client Contact"
                )

            if chaser_type == "Chaser 2":

                if not clean_text(
                    row.get(
                        HELPER_FUND_KEY,
                        ""
                    )
                ):

                    return (
                        "FAIL - No Credit Helper Match"
                    )

                if not credit:

                    return (
                        "FAIL - Missing Credit Contact"
                    )

            return "PASS"

        df_filtered["VALIDATION_STATUS"] = (
            df_filtered.apply(
                validation_status,
                axis=1
            )
        )

        # =====================================================
        # EMAIL TABLE
        # =====================================================

        df_filtered["Fund Name"] = (
            df_filtered[WF_FUND_NAME]
        )

        df_filtered["NAV Date"] = nav_date_input

        df_filtered["NAV"] = ""
        df_filtered["MTD"] = ""
        df_filtered["Comments"] = ""

        # =====================================================
        # PASS / FAIL
        # =====================================================

        pass_df = df_filtered[
            df_filtered["VALIDATION_STATUS"]
            == "PASS"
        ].copy()

        fail_df = df_filtered[
            df_filtered["VALIDATION_STATUS"]
            != "PASS"
        ].copy()

        # =====================================================
        # SAVE VALIDATION FILE
        # =====================================================

        ts = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        validation_path = os.path.join(
            output_folder,
            f"HF_NAV_Validation_{ts}.xlsx"
        )

        with pd.ExcelWriter(
            validation_path,
            engine="openpyxl"
        ) as writer:

            pass_df.to_excel(
                writer,
                sheet_name=PASS_SHEET,
                index=False
            )

            fail_df.to_excel(
                writer,
                sheet_name=FAIL_SHEET,
                index=False
            )

        validation_file_var.set(validation_path)

        load_senders()

        info(
            f"Validation file created successfully:\n\n"
            f"{validation_path}"
        )

    except Exception as e:

        error(
            f"Validation failed:\n\n"
            f"{e}\n\n"
            f"{traceback.format_exc()}"
        )


# =========================================================
# EMAIL GENERATION
# =========================================================

def generate_emails():

    try:

        val_path = validation_file_var.get().strip()

        if not val_path:

            error(
                "Please select Validation file."
            )

            return

        pass_df = read_excel(
            val_path,
            PASS_SHEET
        )

        if pass_df.empty:

            error("PASS sheet is empty.")

            return

        subject = (
            subject_text
            .get("1.0", "end")
            .strip()
        )

        body = (
            body_text
            .get("1.0", "end")
            .strip()
        )

        sender_details = (
            get_selected_sender_details()
        )

        send_mode = send_mode_var.get()

        pythoncom.CoInitialize()

        outlook = win32.Dispatch(
            "Outlook.Application"
        )

        grouped = pass_df.groupby(
            "TO_ADDRESS_FINAL",
            dropna=False
        )

        count = 0

        draft_mails = []

        for to_addr, group in grouped:

            to_addr = normalize_email_string(
                to_addr
            )

            if not to_addr:
                continue

            cc_addr = combine_emails(
                *group[
                    "CC_ADDRESS_FINAL"
                ].tolist()
            )

            table_df = group[
                [
                    "Fund Name",
                    "NAV Date",
                    "NAV",
                    "MTD",
                    "Comments",
                ]
            ].copy()

            table_html = (
                build_professional_table(
                    table_df
                )
            )

            signature_html = (
                make_signature(
                    sender_details
                )
            )

            html_body = build_email_html(
                body,
                table_html,
                signature_html
            )

            mail = outlook.CreateItem(0)

            mail.To = to_addr
            mail.CC = cc_addr
            mail.Subject = subject
            mail.HTMLBody = html_body

            # =========================================
            # AUTO SEND
            # =========================================

            if send_mode == "AUTO":

                mail.Send()

            # =========================================
            # REVIEW & BULK SEND
            # =========================================

            elif send_mode == "REVIEW":

                mail.Display()

                mail.Save()

                draft_mails.append(mail)

            # =========================================
            # DRAFT MODE
            # =========================================

            else:

                mail.Save()

            count += 1

        # =============================================
        # REVIEW FINAL CONFIRMATION
        # =============================================

        if send_mode == "REVIEW":

            final_answer = input(
                "\nAll emails generated successfully.\n"
                "Please review Outlook drafts.\n\n"
                "Send ALL emails now? Y/N: "
            ).strip().upper()

            if final_answer == "Y":

                for draft_mail in draft_mails:

                    draft_mail.Send()

                info(
                    f"All {count} emails sent successfully."
                )

            else:

                info(
                    f"{count} draft emails saved in Outlook."
                )

        else:

            info(
                f"{count} email(s) processed successfully."
            )

    except Exception as e:

        error(
            f"Email generation failed:\n\n"
            f"{e}\n\n"
            f"{traceback.format_exc()}"
        )


# =========================================================
# WORKFLOW UPDATE
# =========================================================

def update_ak():

    try:

        wf_path = (
            workflow_update_var.get().strip()
            or workflow_path
        )

        val_path = (
            validation_file_var.get().strip()
        )

        if not wf_path:

            error(
                "Please select Workflow file."
            )

            return

        if not val_path:

            error(
                "Please select Validation file."
            )

            return

        comment = (
            comment_text
            .get("1.0", "end")
            .strip()
        )

        if not comment:

            error(
                "Please enter comment text."
            )

            return

        workflow = read_excel(
            wf_path,
            WORKFLOW_SHEET
        )

        pass_df = read_excel(
            val_path,
            PASS_SHEET
        )

        keys = set(
            pass_df[WF_FUND_KEY]
            .astype(str)
            .str.strip()
        )

        final_comment = (
            f"{today_str()} - {comment}"
        )

        mask = (
            workflow[WF_FUND_KEY]
            .astype(str)
            .str.strip()
            .isin(keys)
        )

        workflow.loc[
            mask,
            WF_AK
        ] = final_comment

        folder = os.path.dirname(wf_path)

        base = os.path.splitext(
            os.path.basename(wf_path)
        )[0]

        ts = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        updated_path = os.path.join(
            folder,
            f"{base}_AK_Updated_{ts}.xlsx"
        )

        workflow.to_excel(
            updated_path,
            index=False
        )

        info(
            f"Workflow updated successfully:\n\n"
            f"{updated_path}"
        )

    except Exception as e:

        error(
            f"Workflow update failed:\n\n"
            f"{e}\n\n"
            f"{traceback.format_exc()}"
        )


# =========================================================
# FILE PICKERS
# =========================================================

def pick_workflow():

    global workflow_path

    path = filedialog.askopenfilename(
        filetypes=[
            ("Excel files", "*.xlsx *.xls")
        ]
    )

    if path:

        workflow_path = path

        workflow_file_var.set(path)


def pick_helper():

    global helper_path

    path = filedialog.askopenfilename(
        filetypes=[
            ("Excel files", "*.xlsx *.xls")
        ]
    )

    if path:

        helper_path = path

        helper_file_var.set(path)


def pick_config():

    global config_path
    global config_df_cache

    path = filedialog.askopenfilename(
        filetypes=[
            ("Excel files", "*.xlsx *.xls")
        ]
    )

    if path:

        config_path = path

        config_file_var.set(path)

        config_df_cache = read_excel(
            config_path,
            CONFIG_SHEET
        )

        load_senders()


def pick_output_folder():

    global output_folder

    path = filedialog.askdirectory()

    if path:

        output_folder = path

        output_folder_var.set(path)


def pick_validation():

    global validation_path

    path = filedialog.askopenfilename(
        filetypes=[
            ("Excel files", "*.xlsx *.xls")
        ]
    )

    if path:

        validation_path = path

        validation_file_var.set(path)


def pick_workflow_update():

    path = filedialog.askopenfilename(
        filetypes=[
            ("Excel files", "*.xlsx *.xls")
        ]
    )

    if path:

        workflow_update_var.set(path)


# =========================================================
# GUI
# =========================================================

root = tk.Tk()

root.title("HF NAV Chaser Automation")

root.geometry("1050x850")

notebook = ttk.Notebook(root)

notebook.pack(
    fill="both",
    expand=True,
    padx=10,
    pady=10
)

tab1 = ttk.Frame(notebook)
tab2 = ttk.Frame(notebook)
tab3 = ttk.Frame(notebook)

notebook.add(tab1, text="1. Validation")
notebook.add(tab2, text="2. Generate Email")
notebook.add(tab3, text="3. Workflow Update")


# =========================================================
# TAB 1
# =========================================================

workflow_file_var = tk.StringVar()
helper_file_var = tk.StringVar()
config_file_var = tk.StringVar()
output_folder_var = tk.StringVar()

chaser_type_var = tk.StringVar(
    value="Chaser 1"
)

frequency_var = tk.StringVar(
    value="Monthly"
)

nav_date_var = tk.StringVar()

frame1 = ttk.LabelFrame(
    tab1,
    text="Validation Setup"
)

frame1.pack(
    fill="x",
    padx=15,
    pady=15
)

# Workflow

ttk.Label(
    frame1,
    text="Workflow File"
).grid(
    row=0,
    column=0,
    sticky="w",
    padx=8,
    pady=8
)

ttk.Entry(
    frame1,
    textvariable=workflow_file_var,
    width=90
).grid(
    row=0,
    column=1,
    padx=8,
    pady=8
)

ttk.Button(
    frame1,
    text="Browse",
    command=pick_workflow
).grid(
    row=0,
    column=2,
    padx=8,
    pady=8
)

# Helper

ttk.Label(
    frame1,
    text="Credit Helper File"
).grid(
    row=1,
    column=0,
    sticky="w",
    padx=8,
    pady=8
)

helper_entry = ttk.Entry(
    frame1,
    textvariable=helper_file_var,
    width=90
)

helper_entry.grid(
    row=1,
    column=1,
    padx=8,
    pady=8
)

helper_button = ttk.Button(
    frame1,
    text="Browse",
    command=pick_helper
)

helper_button.grid(
    row=1,
    column=2,
    padx=8,
    pady=8
)

# Config

ttk.Label(
    frame1,
    text="Email Config File"
).grid(
    row=2,
    column=0,
    sticky="w",
    padx=8,
    pady=8
)

ttk.Entry(
    frame1,
    textvariable=config_file_var,
    width=90
).grid(
    row=2,
    column=1,
    padx=8,
    pady=8
)

ttk.Button(
    frame1,
    text="Browse",
    command=pick_config
).grid(
    row=2,
    column=2,
    padx=8,
    pady=8
)

# Chaser Type

ttk.Label(
    frame1,
    text="Chaser Type"
).grid(
    row=3,
    column=0,
    sticky="w",
    padx=8,
    pady=8
)

chaser_dropdown = ttk.Combobox(
    frame1,
    textvariable=chaser_type_var,
    values=["Chaser 1", "Chaser 2"],
    state="readonly",
    width=30
)

chaser_dropdown.grid(
    row=3,
    column=1,
    sticky="w",
    padx=8,
    pady=8
)

chaser_dropdown.bind(
    "<<ComboboxSelected>>",
    on_chaser_type_change
)

# Frequency

ttk.Label(
    frame1,
    text="Frequency"
).grid(
    row=4,
    column=0,
    sticky="w",
    padx=8,
    pady=8
)

ttk.Combobox(
    frame1,
    textvariable=frequency_var,
    values=["Monthly", "Quarterly"],
    state="readonly",
    width=30
).grid(
    row=4,
    column=1,
    sticky="w",
    padx=8,
    pady=8
)

# NAV Date

ttk.Label(
    frame1,
    text="NAV Date"
).grid(
    row=5,
    column=0,
    sticky="w",
    padx=8,
    pady=8
)

ttk.Entry(
    frame1,
    textvariable=nav_date_var,
    width=35
).grid(
    row=5,
    column=1,
    sticky="w",
    padx=8,
    pady=8
)

# Output Folder

ttk.Label(
    frame1,
    text="Output Folder"
).grid(
    row=6,
    column=0,
    sticky="w",
    padx=8,
    pady=8
)

ttk.Entry(
    frame1,
    textvariable=output_folder_var,
    width=90
).grid(
    row=6,
    column=1,
    padx=8,
    pady=8
)

ttk.Button(
    frame1,
    text="Browse",
    command=pick_output_folder
).grid(
    row=6,
    column=2,
    padx=8,
    pady=8
)

# Validate Button

ttk.Button(
    frame1,
    text="Validate",
    command=create_validation_file
).grid(
    row=7,
    column=1,
    sticky="w",
    padx=8,
    pady=15
)

on_chaser_type_change()


# =========================================================
# TAB 2
# =========================================================

validation_file_var = tk.StringVar()

sender_var = tk.StringVar()

send_mode_var = tk.StringVar(
    value="DRAFT"
)

frame2 = ttk.LabelFrame(
    tab2,
    text="Generate Email"
)

frame2.pack(
    fill="both",
    expand=True,
    padx=15,
    pady=15
)

# Validation File

ttk.Label(
    frame2,
    text="Validation File"
).grid(
    row=0,
    column=0,
    sticky="w",
    padx=8,
    pady=8
)

ttk.Entry(
    frame2,
    textvariable=validation_file_var,
    width=90
).grid(
    row=0,
    column=1,
    padx=8,
    pady=8
)

ttk.Button(
    frame2,
    text="Browse",
    command=pick_validation
).grid(
    row=0,
    column=2,
    padx=8,
    pady=8
)

# Sender

ttk.Label(
    frame2,
    text="Sender Name"
).grid(
    row=1,
    column=0,
    sticky="w",
    padx=8,
    pady=8
)

sender_dropdown = ttk.Combobox(
    frame2,
    textvariable=sender_var,
    state="readonly",
    width=40
)

sender_dropdown.grid(
    row=1,
    column=1,
    sticky="w",
    padx=8,
    pady=8
)

# Subject

ttk.Label(
    frame2,
    text="Subject"
).grid(
    row=2,
    column=0,
    sticky="nw",
    padx=8,
    pady=8
)

subject_text = tk.Text(
    frame2,
    height=2,
    width=75
)

subject_text.grid(
    row=2,
    column=1,
    padx=8,
    pady=8
)

subject_text.insert(
    "1.0",
    "NAV / AUM Request"
)

# Body

ttk.Label(
    frame2,
    text="Body"
).grid(
    row=3,
    column=0,
    sticky="nw",
    padx=8,
    pady=8
)

body_text = tk.Text(
    frame2,
    height=6,
    width=75
)

body_text.grid(
    row=3,
    column=1,
    padx=8,
    pady=8
)

body_text.insert(
    "1.0",
    "We kindly request you to provide the latest NAV and performance details for the below funds."
)

# Send Mode

ttk.Label(
    frame2,
    text="Send Mode"
).grid(
    row=4,
    column=0,
    sticky="w",
    padx=8,
    pady=8
)

mode_frame = ttk.Frame(frame2)

mode_frame.grid(
    row=4,
    column=1,
    sticky="w",
    padx=8,
    pady=8
)

ttk.Radiobutton(
    mode_frame,
    text="Draft Mode",
    variable=send_mode_var,
    value="DRAFT"
).pack(side="left", padx=5)

ttk.Radiobutton(
    mode_frame,
    text="Auto Send",
    variable=send_mode_var,
    value="AUTO"
).pack(side="left", padx=5)

ttk.Radiobutton(
    mode_frame,
    text="Review & Bulk Send",
    variable=send_mode_var,
    value="REVIEW"
).pack(side="left", padx=5)

# Generate Emails

ttk.Button(
    frame2,
    text="Generate Emails",
    command=generate_emails
).grid(
    row=5,
    column=1,
    sticky="w",
    padx=8,
    pady=15
)


# =========================================================
# TAB 3
# =========================================================

workflow_update_var = tk.StringVar()

frame3 = ttk.LabelFrame(
    tab3,
    text="Workflow AK Update"
)

frame3.pack(
    fill="both",
    expand=True,
    padx=15,
    pady=15
)

# Workflow

ttk.Label(
    frame3,
    text="Workflow File"
).grid(
    row=0,
    column=0,
    sticky="w",
    padx=8,
    pady=8
)

ttk.Entry(
    frame3,
    textvariable=workflow_update_var,
    width=90
).grid(
    row=0,
    column=1,
    padx=8,
    pady=8
)

ttk.Button(
    frame3,
    text="Browse",
    command=pick_workflow_update
).grid(
    row=0,
    column=2,
    padx=8,
    pady=8
)

# Comment

ttk.Label(
    frame3,
    text="Comment Text"
).grid(
    row=1,
    column=0,
    sticky="nw",
    padx=8,
    pady=8
)

comment_text = tk.Text(
    frame3,
    height=5,
    width=75
)

comment_text.grid(
    row=1,
    column=1,
    padx=8,
    pady=8
)

# Update Button

ttk.Button(
    frame3,
    text="Update AK",
    command=update_ak
).grid(
    row=2,
    column=1,
    sticky="w",
    padx=8,
    pady=15
)

root.mainloop()
















# =====================================================
# CLEAN VALIDATION COLUMNS
# =====================================================

required_columns = [

    WF_FUND_KEY,
    WF_FUND_NAME,
    WF_IA_NAME,
    WF_REGION,
    WF_FREQ,

    "Fund Name",
    "NAV Date",
    "NAV",
    "MTD",
    "Comments",

    "TO_ADDRESS_FINAL",
    "CC_ADDRESS_FINAL",

    "VALIDATION_STATUS"
]

# =====================================================
# PASS SHEET
# =====================================================

pass_df = df_filtered[
    df_filtered["VALIDATION_STATUS"] == "PASS"
][required_columns].copy()

# =====================================================
# FAIL SHEET
# =====================================================

fail_df = df_filtered[
    df_filtered["VALIDATION_STATUS"] != "PASS"
][required_columns].copy()

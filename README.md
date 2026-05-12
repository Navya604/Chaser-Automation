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
# UPDATE THESE COLUMN NAMES BASED ON YOUR REAL FILE HEADERS
# =========================================================

WORKFLOW_SHEET = "Sheet1"
WCR_SHEET = "Sheet1"
CONFIG_SHEET = "Sheet1"

# Workflow columns
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
WF_AS = "AS"
WF_AK = "AK"

# WCR columns
WCR_FUND_KEY = "Fund UCN"
WCR_CREDIT_CONTACT = "Credit Contact"

# Config columns
CFG_NAME = "Name"
CFG_EMAIL = "Email"
CFG_TYPE = "Type"

# Values
ALLOWED_REGIONS = ["NAHF", "LATAM"]
NAV_BUCKET_ALLOWED = ["1-10", "11-30"]
MTD_BUCKET_ALLOWED = ["1-10"]
COVERAGE_EXCLUDE = ["BLOOMBERG", "PYTHON", "DAY NAV 2", "PYTHON WEB", "PYTHON - WEB"]

PASS_SHEET = "Pass"
FAIL_SHEET = "Fail"


# =========================================================
# GLOBAL VARIABLES
# =========================================================

workflow_path = ""
wcr_path = ""
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
            emails.extend([e.strip() for e in val.split(";") if e.strip()])

    final = []
    seen = set()

    for email in emails:
        low = email.lower()
        if low not in seen:
            seen.add(low)
            final.append(email)

    return "; ".join(final)


def read_excel(path, sheet):
    return pd.read_excel(path, sheet_name=sheet, dtype=str).fillna("")


def check_columns(df, required_cols, file_name):
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise Exception(f"Missing columns in {file_name}: {missing}")


def info(msg):
    messagebox.showinfo("HF NAV Chaser Automation", msg)


def error(msg):
    messagebox.showerror("HF NAV Chaser Automation", msg)


def make_signature(sender_name):
    return f"""
    Best Regards,<br>
    {sender_name}
    """


def build_email_html(body_text, table_html, signature_html):
    return f"""
    <html>
    <body style="font-family:Calibri, Arial, sans-serif; font-size:11pt;">
        <p>Dear Team,</p>

        <p>{body_text}</p>

        {table_html}

        <p>Please fill in the required details and revert at your earliest convenience.</p>

        <p>If the information has already been shared, please ignore this request.</p>

        <p>{signature_html}</p>
    </body>
    </html>
    """


# =========================================================
# CONFIG FILE LOGIC
# =========================================================

def get_config_email(email_type):
    global config_df_cache

    rows = config_df_cache[config_df_cache[CFG_TYPE].astype(str).str.upper() == email_type.upper()]

    if rows.empty:
        return ""

    return normalize_email_string(rows.iloc[0][CFG_EMAIL])


def load_senders():
    global config_df_cache

    if config_df_cache.empty and config_path:
        config_df_cache = read_excel(config_path, CONFIG_SHEET)

    sender_rows = config_df_cache[config_df_cache[CFG_TYPE].astype(str).str.upper() == "SENDER"]

    sender_names = sender_rows[CFG_NAME].dropna().astype(str).tolist()

    sender_dropdown["values"] = sender_names

    if sender_names:
        sender_var.set(sender_names[0])


def get_selected_sender_name():
    return clean_text(sender_var.get())


# =========================================================
# VALIDATION LOGIC
# =========================================================

def create_validation_file():
    global workflow_path, wcr_path, config_path, output_folder, validation_path, config_df_cache

    try:
        if not workflow_path:
            error("Please select Workflow file.")
            return

        if not wcr_path:
            error("Please select WCR report.")
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

        workflow = read_excel(workflow_path, WORKFLOW_SHEET)
        wcr = read_excel(wcr_path, WCR_SHEET)
        config_df_cache = read_excel(config_path, CONFIG_SHEET)

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
            check_columns(workflow, [WF_AS], "Workflow")

        check_columns(
            wcr,
            [
                WCR_FUND_KEY,
                WCR_CREDIT_CONTACT,
            ],
            "WCR"
        )

        check_columns(
            config_df_cache,
            [
                CFG_NAME,
                CFG_EMAIL,
                CFG_TYPE,
            ],
            "Config"
        )

        # Clean workflow columns
        df = workflow.copy()

        df[WF_REGION] = df[WF_REGION].apply(normalize_upper)
        df[WF_NAV_BUCKET] = df[WF_NAV_BUCKET].apply(clean_text)
        df[WF_MTD_BUCKET] = df[WF_MTD_BUCKET].apply(clean_text)
        df[WF_COVERAGE] = df[WF_COVERAGE].apply(normalize_upper)
        df[WF_WEB] = df[WF_WEB].apply(clean_text)
        df[WF_DNC] = df[WF_DNC].apply(clean_text)
        df[WF_FREQ] = df[WF_FREQ].apply(normalize_upper)

        # =====================================================
        # CHASER 1 FILTERING
        # =====================================================
        if chaser_type == "Chaser 1":
            df_filtered = df[
                (df[WF_REGION].isin(ALLOWED_REGIONS)) &
                (df[WF_NAV_BUCKET].isin(NAV_BUCKET_ALLOWED)) &
                (df[WF_MTD_BUCKET].isin(MTD_BUCKET_ALLOWED)) &
                (~df[WF_COVERAGE].isin(COVERAGE_EXCLUDE)) &
                (df[WF_WEB] == "") &
                (df[WF_DNC] == "") &
                (df[WF_FREQ] == frequency_type.upper())
            ].copy()

        # =====================================================
        # CHASER 2 FILTERING
        # =====================================================
        else:
            df[WF_AS] = df[WF_AS].apply(normalize_upper)

            df_filtered = df[
                (df[WF_REGION].isin(ALLOWED_REGIONS)) &
                (df[WF_AS] == "Y") &
                (df[WF_FREQ] == frequency_type.upper())
            ].copy()

        if df_filtered.empty:
            error("No records found after applying filters.")
            return

        # Merge with WCR
        merged = df_filtered.merge(
            wcr,
            left_on=WF_FUND_KEY,
            right_on=WCR_FUND_KEY,
            how="left",
            suffixes=("_WF", "_WCR")
        )

        jpm_nav_email = get_config_email("JPM_NAV")
        hfc_brazil_email = get_config_email("LATAM")

        if not jpm_nav_email:
            error("JPM_NAV email is missing in Config file.")
            return

        # Build routing
        def build_to(row):
            client = normalize_email_string(row.get(WF_CLIENT_CONTACT, ""))
            credit = normalize_email_string(row.get(WCR_CREDIT_CONTACT, ""))

            if client:
                return client
            return credit

        def build_cc(row):
            region = normalize_upper(row.get(WF_REGION, ""))
            client = normalize_email_string(row.get(WF_CLIENT_CONTACT, ""))
            credit = normalize_email_string(row.get(WCR_CREDIT_CONTACT, ""))

            # Chaser 1
            if chaser_type == "Chaser 1":
                cc = combine_emails(jpm_nav_email)

            # Chaser 2
            else:
                if client:
                    cc = combine_emails(credit, jpm_nav_email)
                else:
                    cc = combine_emails(jpm_nav_email)

            # LATAM rule
            if region == "LATAM":
                cc = combine_emails(cc, hfc_brazil_email)

            return cc

        merged["TO_ADDRESS_FINAL"] = merged.apply(build_to, axis=1)
        merged["CC_ADDRESS_FINAL"] = merged.apply(build_cc, axis=1)
        merged["CHASER_TYPE"] = chaser_type
        merged["REQUESTED_NAV_DATE"] = nav_date_input

        # Validation
        def validation_status(row):
            if not clean_text(row.get(WCR_FUND_KEY, "")):
                return "FAIL - No WCR Match"

            if not clean_text(row.get("TO_ADDRESS_FINAL", "")):
                return "FAIL - Missing To Address"

            return "PASS"

        merged["VALIDATION_STATUS"] = merged.apply(validation_status, axis=1)

        # Email table fields
        merged["Fund Name"] = merged[WF_FUND_NAME]
        merged["NAV Date"] = nav_date_input
        merged["NAV"] = ""
        merged["MTD"] = ""
        merged["Comments"] = ""

        output_cols = [
            WF_FUND_KEY,
            WF_FUND_NAME,
            WF_IA_NAME,
            WF_REGION,
            WF_FREQ,
            WF_CLIENT_CONTACT,
            WCR_CREDIT_CONTACT,
            "TO_ADDRESS_FINAL",
            "CC_ADDRESS_FINAL",
            "CHASER_TYPE",
            "REQUESTED_NAV_DATE",
            "VALIDATION_STATUS",
            "Fund Name",
            "NAV Date",
            "NAV",
            "MTD",
            "Comments",
        ]

        pass_df = merged[merged["VALIDATION_STATUS"] == "PASS"][output_cols].copy()
        fail_df = merged[merged["VALIDATION_STATUS"] != "PASS"][output_cols].copy()

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        validation_path = os.path.join(output_folder, f"HF_NAV_Validation_{ts}.xlsx")

        with pd.ExcelWriter(validation_path, engine="openpyxl") as writer:
            pass_df.to_excel(writer, sheet_name=PASS_SHEET, index=False)
            fail_df.to_excel(writer, sheet_name=FAIL_SHEET, index=False)

        validation_file_var.set(validation_path)
        load_senders()

        info(f"Validation file created successfully:\n\n{validation_path}")

    except Exception as e:
        error(f"Validation failed:\n\n{e}\n\n{traceback.format_exc()}")


# =========================================================
# EMAIL GENERATION
# =========================================================

def generate_emails():
    try:
        val_path = validation_file_var.get().strip()

        if not val_path:
            error("Please select Validation file.")
            return

        pass_df = read_excel(val_path, PASS_SHEET)

        if pass_df.empty:
            error("PASS sheet is empty.")
            return

        check_columns(
            pass_df,
            [
                "TO_ADDRESS_FINAL",
                "CC_ADDRESS_FINAL",
                "Fund Name",
                "NAV Date",
                "NAV",
                "MTD",
                "Comments",
            ],
            "PASS Sheet"
        )

        subject = subject_text.get("1.0", "end").strip()
        body = body_text.get("1.0", "end").strip()
        sender_name = get_selected_sender_name()
        send_mode = send_mode_var.get()

        if not sender_name:
            error("Please select sender.")
            return

        if not subject:
            error("Please enter subject.")
            return

        if not body:
            error("Please enter body.")
            return

        if send_mode == "AUTO":
            proceed = messagebox.askyesno(
                "Confirm Auto Send",
                "You selected Auto Send.\n\nEmails will be sent directly.\n\nDo you want to continue?"
            )
            if not proceed:
                return

        pythoncom.CoInitialize()
        outlook = win32.Dispatch("Outlook.Application")

        grouped = pass_df.groupby("TO_ADDRESS_FINAL", dropna=False)

        count = 0

        for to_addr, group in grouped:
            to_addr = normalize_email_string(to_addr)

            if not to_addr:
                continue

            cc_addr = combine_emails(*group["CC_ADDRESS_FINAL"].tolist())

            table_df = group[
                [
                    "Fund Name",
                    "NAV Date",
                    "NAV",
                    "MTD",
                    "Comments",
                ]
            ].copy()

            table_html = table_df.to_html(
                index=False,
                border=1,
                justify="left"
            )

            signature_html = make_signature(sender_name)

            html_body = build_email_html(body, table_html, signature_html)

            mail = outlook.CreateItem(0)
            mail.To = to_addr
            mail.CC = cc_addr
            mail.Subject = subject
            mail.HTMLBody = html_body

            if send_mode == "AUTO":
                mail.Send()
            elif send_mode == "REVIEW":
                mail.Display()
                answer = input(f"Send email to {to_addr}? Y/N: ").strip().upper()

                if answer == "Y":
                    mail.Send()
                else:
                    mail.Save()
            else:
                mail.Save()

            count += 1

        info(f"{count} email(s) processed successfully.")

    except Exception as e:
        error(f"Email generation failed:\n\n{e}\n\n{traceback.format_exc()}")


# =========================================================
# WORKFLOW AK UPDATE
# =========================================================

def update_ak():
    try:
        wf_path = workflow_update_var.get().strip() or workflow_path
        val_path = validation_file_var.get().strip()

        if not wf_path:
            error("Please select Workflow file.")
            return

        if not val_path:
            error("Please select Validation file.")
            return

        comment = comment_text.get("1.0", "end").strip()

        if not comment:
            error("Please enter comment text.")
            return

        workflow = read_excel(wf_path, WORKFLOW_SHEET)
        pass_df = read_excel(val_path, PASS_SHEET)

        check_columns(workflow, [WF_FUND_KEY, WF_AK], "Workflow")
        check_columns(pass_df, [WF_FUND_KEY], "PASS Sheet")

        keys = set(pass_df[WF_FUND_KEY].astype(str).str.strip())
        final_comment = f"{today_str()} - {comment}"

        mask = workflow[WF_FUND_KEY].astype(str).str.strip().isin(keys)

        workflow.loc[mask, WF_AK] = final_comment

        folder = os.path.dirname(wf_path)
        base = os.path.splitext(os.path.basename(wf_path))[0]
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")

        updated_path = os.path.join(folder, f"{base}_AK_Updated_{ts}.xlsx")

        workflow.to_excel(updated_path, index=False)

        info(f"Workflow updated successfully:\n\n{updated_path}")

    except Exception as e:
        error(f"Workflow update failed:\n\n{e}\n\n{traceback.format_exc()}")


# =========================================================
# FILE PICKERS
# =========================================================

def pick_workflow():
    global workflow_path
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if path:
        workflow_path = path
        workflow_file_var.set(path)


def pick_wcr():
    global wcr_path
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if path:
        wcr_path = path
        wcr_file_var.set(path)


def pick_config():
    global config_path
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if path:
        config_path = path
        config_file_var.set(path)
        load_senders()


def pick_output_folder():
    global output_folder
    path = filedialog.askdirectory()
    if path:
        output_folder = path
        output_folder_var.set(path)


def pick_validation():
    global validation_path
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if path:
        validation_path = path
        validation_file_var.set(path)


def pick_workflow_update():
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if path:
        workflow_update_var.set(path)


# =========================================================
# GUI
# =========================================================

root = tk.Tk()
root.title("HF NAV Chaser Automation")
root.geometry("1050x850")

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=10, pady=10)

tab1 = ttk.Frame(notebook)
tab2 = ttk.Frame(notebook)
tab3 = ttk.Frame(notebook)

notebook.add(tab1, text="1. Validation")
notebook.add(tab2, text="2. Generate Email")
notebook.add(tab3, text="3. Workflow Update")


# =========================================================
# TAB 1 - VALIDATION
# =========================================================

workflow_file_var = tk.StringVar()
wcr_file_var = tk.StringVar()
config_file_var = tk.StringVar()
output_folder_var = tk.StringVar()

chaser_type_var = tk.StringVar(value="Chaser 1")
frequency_var = tk.StringVar(value="Monthly")
nav_date_var = tk.StringVar()

frame1 = ttk.LabelFrame(tab1, text="Validation Setup")
frame1.pack(fill="x", padx=15, pady=15)

ttk.Label(frame1, text="Workflow File").grid(row=0, column=0, sticky="w", padx=8, pady=8)
ttk.Entry(frame1, textvariable=workflow_file_var, width=90).grid(row=0, column=1, padx=8, pady=8)
ttk.Button(frame1, text="Browse", command=pick_workflow).grid(row=0, column=2, padx=8, pady=8)

ttk.Label(frame1, text="WCR Report").grid(row=1, column=0, sticky="w", padx=8, pady=8)
ttk.Entry(frame1, textvariable=wcr_file_var, width=90).grid(row=1, column=1, padx=8, pady=8)
ttk.Button(frame1, text="Browse", command=pick_wcr).grid(row=1, column=2, padx=8, pady=8)

ttk.Label(frame1, text="Email Config File").grid(row=2, column=0, sticky="w", padx=8, pady=8)
ttk.Entry(frame1, textvariable=config_file_var, width=90).grid(row=2, column=1, padx=8, pady=8)
ttk.Button(frame1, text="Browse", command=pick_config).grid(row=2, column=2, padx=8, pady=8)

ttk.Label(frame1, text="Chaser Type").grid(row=3, column=0, sticky="w", padx=8, pady=8)
ttk.Combobox(
    frame1,
    textvariable=chaser_type_var,
    values=["Chaser 1", "Chaser 2"],
    state="readonly",
    width=30
).grid(row=3, column=1, sticky="w", padx=8, pady=8)

ttk.Label(frame1, text="Frequency").grid(row=4, column=0, sticky="w", padx=8, pady=8)
ttk.Combobox(
    frame1,
    textvariable=frequency_var,
    values=["Monthly", "Quarterly"],
    state="readonly",
    width=30
).grid(row=4, column=1, sticky="w", padx=8, pady=8)

ttk.Label(frame1, text="NAV Date").grid(row=5, column=0, sticky="w", padx=8, pady=8)
ttk.Entry(frame1, textvariable=nav_date_var, width=35).grid(row=5, column=1, sticky="w", padx=8, pady=8)

ttk.Label(frame1, text="Output Folder").grid(row=6, column=0, sticky="w", padx=8, pady=8)
ttk.Entry(frame1, textvariable=output_folder_var, width=90).grid(row=6, column=1, padx=8, pady=8)
ttk.Button(frame1, text="Browse", command=pick_output_folder).grid(row=6, column=2, padx=8, pady=8)

ttk.Button(frame1, text="Validate", command=create_validation_file).grid(
    row=7, column=1, sticky="w", padx=8, pady=15
)


# =========================================================
# TAB 2 - GENERATE EMAIL
# =========================================================

validation_file_var = tk.StringVar()
sender_var = tk.StringVar()
send_mode_var = tk.StringVar(value="DRAFT")

frame2 = ttk.LabelFrame(tab2, text="Generate Email")
frame2.pack(fill="both", expand=True, padx=15, pady=15)

ttk.Label(frame2, text="Validation File").grid(row=0, column=0, sticky="w", padx=8, pady=8)
ttk.Entry(frame2, textvariable=validation_file_var, width=90).grid(row=0, column=1, padx=8, pady=8)
ttk.Button(frame2, text="Browse", command=pick_validation).grid(row=0, column=2, padx=8, pady=8)

ttk.Label(frame2, text="Sender Name").grid(row=1, column=0, sticky="w", padx=8, pady=8)
sender_dropdown = ttk.Combobox(frame2, textvariable=sender_var, state="readonly", width=40)
sender_dropdown.grid(row=1, column=1, sticky="w", padx=8, pady=8)

ttk.Label(frame2, text="Subject").grid(row=2, column=0, sticky="nw", padx=8, pady=8)
subject_text = tk.Text(frame2, height=2, width=75)
subject_text.grid(row=2, column=1, padx=8, pady=8)
subject_text.insert("1.0", "NAV / AUM Request")

ttk.Label(frame2, text="Body").grid(row=3, column=0, sticky="nw", padx=8, pady=8)
body_text = tk.Text(frame2, height=6, width=75)
body_text.grid(row=3, column=1, padx=8, pady=8)
body_text.insert(
    "1.0",
    "We kindly request you to provide the latest NAV and performance details for the below funds."
)

ttk.Label(frame2, text="Send Mode").grid(row=4, column=0, sticky="w", padx=8, pady=8)

mode_frame = ttk.Frame(frame2)
mode_frame.grid(row=4, column=1, sticky="w", padx=8, pady=8)

ttk.Radiobutton(mode_frame, text="Draft Mode", variable=send_mode_var, value="DRAFT").pack(side="left", padx=5)
ttk.Radiobutton(mode_frame, text="Auto Send", variable=send_mode_var, value="AUTO").pack(side="left", padx=5)
ttk.Radiobutton(mode_frame, text="Review Confirm Send", variable=send_mode_var, value="REVIEW").pack(side="left", padx=5)

ttk.Button(frame2, text="Generate Emails", command=generate_emails).grid(
    row=5, column=1, sticky="w", padx=8, pady=15
)


# =========================================================
# TAB 3 - WORKFLOW UPDATE
# =========================================================

workflow_update_var = tk.StringVar()

frame3 = ttk.LabelFrame(tab3, text="Workflow AK Update")
frame3.pack(fill="both", expand=True, padx=15, pady=15)

ttk.Label(frame3, text="Workflow File").grid(row=0, column=0, sticky="w", padx=8, pady=8)
ttk.Entry(frame3, textvariable=workflow_update_var, width=90).grid(row=0, column=1, padx=8, pady=8)
ttk.Button(frame3, text="Browse", command=pick_workflow_update).grid(row=0, column=2, padx=8, pady=8)

ttk.Label(frame3, text="Comment Text").grid(row=1, column=0, sticky="nw", padx=8, pady=8)
comment_text = tk.Text(frame3, height=5, width=75)
comment_text.grid(row=1, column=1, padx=8, pady=8)

ttk.Button(frame3, text="Update AK", command=update_ak).grid(
    row=2, column=1, sticky="w", padx=8, pady=15
)

root.mainloop()

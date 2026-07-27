
def update_workflow_comments():
    try:
        print("1. Button clicked", flush=True)

        wf_path = (
            workflow_update_var.get().strip()
            or workflow_path
        )

        val_path = validation_file_var.get().strip()

        print("2. Workflow Path:", wf_path, flush=True)
        print("3. Validation Path:", val_path, flush=True)

        if not wf_path:
            error("Please select Workflow file.")
            return

        if not val_path:
            error("Please select Validation file.")
            return

        comment = (
            comment_text
            .get("1.0", "end")
            .strip()
        )

        if not comment:
            error("Please enter comment text.")
            return

        chaser_type = chaser_type_var.get().strip()
        print("4. Chaser Type:", chaser_type, flush=True)

        if chaser_type == "Chaser 1":
            comment_column = "Chaser 1 Comments"
        elif chaser_type == "Chaser 2":
            comment_column = "Chaser 2 Comments"
        else:
            error("Please select a valid Chaser Type.")
            return

        print("5. Reading workflow...", flush=True)

        workflow = pd.read_excel(
            wf_path,
            sheet_name=WORKFLOW_SHEET,
            dtype=str
        )

        print("6. Workflow read successfully", flush=True)

        if comment_column not in workflow.columns:
            workflow[comment_column] = ""

        print("7. Reading validation...", flush=True)

        pass_df = pd.read_excel(
            val_path,
            sheet_name=PASS_SHEET,
            dtype=str
        )

        print("8. Validation read successfully", flush=True)

        if WF_FUND_KEY not in workflow.columns:
            error(
                f"'{WF_FUND_KEY}' was not found in the Workflow file."
            )
            return

        if WF_FUND_KEY not in pass_df.columns:
            error(
                f"'{WF_FUND_KEY}' was not found in the PASS sheet."
            )
            return

        keys = set(
            pass_df[WF_FUND_KEY]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        mask = (
            workflow[WF_FUND_KEY]
            .fillna("")
            .astype(str)
            .str.strip()
            .isin(keys)
        )

        matched_rows = int(mask.sum())
        print("9. Matched Rows:", matched_rows, flush=True)

        if matched_rows == 0:
            error(
                "No matching Fund UCNs were found between "
                "the Workflow and Validation files."
            )
            return

        final_comment = f"{today_str()} - {comment}"

        workflow.loc[
            mask,
            comment_column
        ] = final_comment

        folder = os.path.dirname(wf_path)

        # Safe date for filename—no slashes
        file_date = datetime.now().strftime("%m_%d_%Y")

        updated_path = os.path.join(
            folder,
            f"{file_date}_Updated Workflow.xlsx"
        )

        print("10. Saving to:", updated_path, flush=True)

        workflow.to_excel(
            updated_path,
            sheet_name=WORKFLOW_SHEET,
            index=False
        )

        print("11. File saved successfully", flush=True)

        info(
            f"Workflow updated successfully.\n\n"
            f"{matched_rows} rows updated in "
            f"'{comment_column}'.\n\n"
            f"New file:\n{updated_path}"
        )

    except Exception as e:
        print(
            "WORKFLOW UPDATE ERROR:",
            repr(e),
            flush=True
        )

        error(
            f"Workflow update failed:\n\n"
            f"{e}\n\n"
            f"{traceback.format_exc()}"
        )


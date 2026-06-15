```python
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(
    page_title="KMAT Rank List Generator",
    layout="wide"
)

st.title("KMAT Rank List Generator")

# --------------------------------------------------
# Upload Files
# --------------------------------------------------

responses_file = st.file_uploader(
    "Upload CBT Responses Excel",
    type=["xlsx"]
)

candidates_file = st.file_uploader(
    "Upload Candidates Excel",
    type=["xlsx"]
)

# --------------------------------------------------
# Process
# --------------------------------------------------

if responses_file is not None and candidates_file is not None:

    try:

        responses = pd.read_excel(responses_file)
        candidates = pd.read_excel(candidates_file)

        # -------------------------------------------
        # Validation
        # -------------------------------------------

        required_response_cols = ["RollNo", "QNo", "Mark"]
        required_candidate_cols = ["RollNo", "ApplNo", "Name"]

        for col in required_response_cols:
            if col not in responses.columns:
                st.error(f"Column '{col}' missing in CBT Responses file")
                st.stop()

        for col in required_candidate_cols:
            if col not in candidates.columns:
                st.error(f"Column '{col}' missing in Candidates file")
                st.stop()

        # -------------------------------------------
        # Part-wise Marks
        # -------------------------------------------

        part1 = (
            responses[
                responses["QNo"].between(1, 50)
            ]
            .groupby("RollNo")["Mark"]
            .sum()
            .reset_index(name="Part1")
        )

        part2 = (
            responses[
                responses["QNo"].between(51, 100)
            ]
            .groupby("RollNo")["Mark"]
            .sum()
            .reset_index(name="Part2")
        )

        part3 = (
            responses[
                responses["QNo"].between(101, 140)
            ]
            .groupby("RollNo")["Mark"]
            .sum()
            .reset_index(name="Part3")
        )

        part4 = (
            responses[
                responses["QNo"].between(141, 180)
            ]
            .groupby("RollNo")["Mark"]
            .sum()
            .reset_index(name="Part4")
        )

        # -------------------------------------------
        # Merge
        # -------------------------------------------

        result = candidates.merge(
            part1,
            on="RollNo",
            how="left"
        )

        result = result.merge(
            part2,
            on="RollNo",
            how="left"
        )

        result = result.merge(
            part3,
            on="RollNo",
            how="left"
        )

        result = result.merge(
            part4,
            on="RollNo",
            how="left"
        )

        # -------------------------------------------
        # Fill only numeric mark columns
        # -------------------------------------------

        for col in ["Part1", "Part2", "Part3", "Part4"]:
            if col not in result.columns:
                result[col] = 0

            result[col] = pd.to_numeric(
                result[col],
                errors="coerce"
            ).fillna(0)

        # -------------------------------------------
        # Total
        # -------------------------------------------

        result["Total"] = (
            result["Part1"]
            + result["Part2"]
            + result["Part3"]
            + result["Part4"]
        )

        # -------------------------------------------
        # Qualification
        # General / SEBC : 72
        # SC/ST/PD       : 54
        # -------------------------------------------

        def qualification(row):

            score = row["Total"]

            category = str(
                row.get("Category", "")
            ).upper()

            special3 = str(
                row.get("Special3", "")
            ).upper()

            if (
                category in ["SC", "ST"]
                or special3 == "PD"
            ):
                return "Qualified" if score >= 54 else "Not Qualified"

            return "Qualified" if score >= 72 else "Not Qualified"

        result["Qualification"] = result.apply(
            qualification,
            axis=1
        )

        # -------------------------------------------
        # Qualified Candidates Only
        # -------------------------------------------

        ranklist = result[
            result["Qualification"] == "Qualified"
        ].copy()

        # -------------------------------------------
        # DOB
        # -------------------------------------------

        if "DOB" in ranklist.columns:
            ranklist["DOB"] = pd.to_datetime(
                ranklist["DOB"],
                errors="coerce"
            )
        else:
            ranklist["DOB"] = pd.NaT

        # -------------------------------------------
        # Tie Break
        # 1 Total
        # 2 Part4
        # 3 Part3
        # 4 Part2
        # 5 Older Candidate
        # -------------------------------------------

        ranklist = ranklist.sort_values(
            by=[
                "Total",
                "Part4",
                "Part3",
                "Part2",
                "DOB"
            ],
            ascending=[
                False,
                False,
                False,
                False,
                True
            ]
        )

        ranklist.reset_index(
            drop=True,
            inplace=True
        )

        # -------------------------------------------
        # Rank
        # -------------------------------------------

        ranklist["Rank"] = (
            ranklist.index + 1
        )

        total_qualified = len(ranklist)

        # -------------------------------------------
        # Percentile
        # -------------------------------------------

        if total_qualified == 1:
            ranklist["Percentile"] = 100.00000
        else:
            ranklist["Percentile"] = ranklist["Rank"].apply(
                lambda x: round(
                    (
                        (total_qualified - x)
                        / (total_qualified - 1)
                    ) * 100,
                    5
                )
            )

        # -------------------------------------------
        # Output
        # -------------------------------------------

        output = pd.DataFrame()

        output["Sl.No"] = ranklist["Rank"]
        output["App. No"] = ranklist["ApplNo"]
        output["Roll No"] = ranklist["RollNo"]
        output["Name"] = ranklist["Name"]

        output["Part-I(Out of 200)"] = ranklist["Part1"]
        output["Part-II(Out of 200)"] = ranklist["Part2"]
        output["Part-III(Out of 160)"] = ranklist["Part3"]
        output["Part-IV(Out of 160)"] = ranklist["Part4"]

        output["Total(Out of 720)"] = ranklist["Total"]
        output["Percentile"] = ranklist["Percentile"]

        # -------------------------------------------
        # Statistics
        # -------------------------------------------

        st.success(
            f"Qualified Candidates : {len(output)}"
        )

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "Appeared",
            len(result)
        )

        col2.metric(
            "Qualified",
            len(ranklist)
        )

        col3.metric(
            "Highest Score",
            output["Total(Out of 720)"].max()
        )

        st.dataframe(
            output,
            use_container_width=True
        )

        # -------------------------------------------
        # Excel Download
        # -------------------------------------------

        excel_buffer = BytesIO()

        with pd.ExcelWriter(
            excel_buffer,
            engine="openpyxl"
        ) as writer:

            output.to_excel(
                writer,
                sheet_name="RankList",
                index=False
            )

        st.download_button(
            label="Download Rank List",
            data=excel_buffer.getvalue(),
            file_name="KMAT_RankList.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(str(e))
```

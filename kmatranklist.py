
import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(
    page_title="MBA Rank List Generator",
    layout="wide"
)

st.title("MBA Entrance Rank List Generator")

# ---------------------------------------------------
# Upload Files
# ---------------------------------------------------

responses_file = st.file_uploader(
    "Upload CBT Responses",
    type=["xlsx"]
)

candidate_file = st.file_uploader(
    "Upload Candidate Master",
    type=["xlsx"]
)

# ---------------------------------------------------
# Process
# ---------------------------------------------------

if responses_file and candidate_file:

    responses = pd.read_excel(responses_file)

    candidates = pd.read_excel(candidate_file)

    # -----------------------------------------
    # Part-wise marks
    # -----------------------------------------

    part1 = (
        responses[responses["QNo"].between(1,50)]
        .groupby("RollNo")["Mark"]
        .sum()
        .reset_index(name="Part1")
    )

    part2 = (
        responses[responses["QNo"].between(51,100)]
        .groupby("RollNo")["Mark"]
        .sum()
        .reset_index(name="Part2")
    )

    part3 = (
        responses[responses["QNo"].between(101,140)]
        .groupby("RollNo")["Mark"]
        .sum()
        .reset_index(name="Part3")
    )

    part4 = (
        responses[responses["QNo"].between(141,180)]
        .groupby("RollNo")["Mark"]
        .sum()
        .reset_index(name="Part4")
    )

    result = candidates.merge(part1,on="RollNo",how="left")
    result = result.merge(part2,on="RollNo",how="left")
    result = result.merge(part3,on="RollNo",how="left")
    result = result.merge(part4,on="RollNo",how="left")

    result.fillna(0,inplace=True)

    # -----------------------------------------
    # Total Score
    # -----------------------------------------

    result["Total"] = (
        result["Part1"]
        + result["Part2"]
        + result["Part3"]
        + result["Part4"]
    )

    # -----------------------------------------
    # Qualification
    # -----------------------------------------

    def qualified(row):

        score = row["Total"]

        category = str(row.get("Category","")).upper()
        special3 = str(row.get("Special3","")).upper()

        if (
            category in ["SC","ST"]
            or special3 == "PD"
        ):
            return "Qualified" if score >= 54 else "Not Qualified"

        return "Qualified" if score >= 72 else "Not Qualified"

    result["Qualification"] = result.apply(
        qualified,
        axis=1
    )

    # -----------------------------------------
    # Only Qualified Candidates
    # -----------------------------------------

    ranklist = result[
        result["Qualification"]=="Qualified"
    ].copy()

    # -----------------------------------------
    # Percentile
    # -----------------------------------------

    total_candidates = len(ranklist)

    ranklist["Percentile"] = (
        ranklist["Total"]
        .rank(method="min",ascending=False)
        .apply(
            lambda r:
            round(
                ((total_candidates-r)
                /(total_candidates-1))*100,
                5
            )
            if total_candidates > 1
            else 100
        )
    )

    # -----------------------------------------
    # Tie Breaking
    # -----------------------------------------
    # Total
    # Part4
    # Part3
    # Part2
    # DOB (older candidate first)

    ranklist["DOB"] = pd.to_datetime(
        ranklist["DOB"],
        errors="coerce"
    )

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

    ranklist["SlNo"] = ranklist.index + 1

    output = ranklist[
        [
            "SlNo",
            "ApplNo",
            "RollNo",
            "Name",
            "Part1",
            "Part2",
            "Part3",
            "Part4",
            "Total",
            "Percentile"
        ]
    ]

    output.columns = [
        "Sl.No",
        "App. No",
        "Roll No",
        "Name",
        "Part-I(Out of 200)",
        "Part-II(Out of 200)",
        "Part-III(Out of 160)",
        "Part-IV(Out of 160)",
        "Total(Out of 720)",
        "Percentile"
    ]

    st.success(
        f"Qualified Candidates : {len(output)}"
    )

    st.dataframe(
        output,
        use_container_width=True
    )

    excel_file = "MBA_RankList.xlsx"

    with pd.ExcelWriter(
        excel_file,
        engine="openpyxl"
    ) as writer:
        output.to_excel(
            writer,
            index=False,
            sheet_name="RankList"
        )

    with open(excel_file,"rb") as f:
        st.download_button(
            "Download Rank List",
            f,
            file_name="MBA_RankList.xlsx"
        )


import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="KMAT Rank List Generator",
    layout="wide"
)

st.title("KMAT Rank List Generator")

# --------------------------------------------------
# Sidebar
# --------------------------------------------------

st.sidebar.header("Correction Factors")

part1_deleted = st.sidebar.number_input(
    "Deleted Questions in Part-I (1-50)",
    min_value=0,
    max_value=49,
    value=0
)

part2_deleted = st.sidebar.number_input(
    "Deleted Questions in Part-II (51-100)",
    min_value=0,
    max_value=49,
    value=0
)

part3_deleted = st.sidebar.number_input(
    "Deleted Questions in Part-III (101-140)",
    min_value=0,
    max_value=39,
    value=0
)

part4_deleted = st.sidebar.number_input(
    "Deleted Questions in Part-IV (141-180)",
    min_value=0,
    max_value=39,
    value=0
)

st.sidebar.header("Qualification")

apply_cutoff = st.sidebar.checkbox(
    "Apply Qualification Cutoff",
    value=False
)

# --------------------------------------------------
# Upload Files
# --------------------------------------------------

responses_file = st.file_uploader(
    "Upload CBT_Responses.xlsx",
    type=["xlsx"]
)

candidates_file = st.file_uploader(
    "Upload Candidates.xlsx",
    type=["xlsx"]
)

# --------------------------------------------------
# Process
# --------------------------------------------------

if responses_file and candidates_file:

    responses = pd.read_excel(responses_file)
    candidates = pd.read_excel(candidates_file)

    responses["QNo"] = pd.to_numeric(
        responses["QNo"],
        errors="coerce"
    )

    responses["Mark"] = pd.to_numeric(
        responses["Mark"],
        errors="coerce"
    ).fillna(0)

    # --------------------------------------------------
    # Part-wise Scores
    # --------------------------------------------------

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

    # --------------------------------------------------
    # Merge
    # --------------------------------------------------

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

    for c in ["Part1", "Part2", "Part3", "Part4"]:
        result[c] = pd.to_numeric(
            result[c],
            errors="coerce"
        ).fillna(0)

    # --------------------------------------------------
    # Correction Factors
    # --------------------------------------------------

    f1 = 50 / (50 - part1_deleted) if part1_deleted < 50 else 1
    f2 = 50 / (50 - part2_deleted) if part2_deleted < 50 else 1
    f3 = 40 / (40 - part3_deleted) if part3_deleted < 40 else 1
    f4 = 40 / (40 - part4_deleted) if part4_deleted < 40 else 1

    result["Part1"] = (result["Part1"] * f1).round(2)
    result["Part2"] = (result["Part2"] * f2).round(2)
    result["Part3"] = (result["Part3"] * f3).round(2)
    result["Part4"] = (result["Part4"] * f4).round(2)

    st.info(
        f"Part-I Factor={f1:.6f} | "
        f"Part-II Factor={f2:.6f} | "
        f"Part-III Factor={f3:.6f} | "
        f"Part-IV Factor={f4:.6f}"
    )

    # --------------------------------------------------
    # Total
    # --------------------------------------------------

    result["Total"] = (
        result["Part1"]
        + result["Part2"]
        + result["Part3"]
        + result["Part4"]
    ).round(2)

    # --------------------------------------------------
    # Remove Absentees
    # --------------------------------------------------

    appeared_candidates = responses["RollNo"].unique()

    result = result[
        result["RollNo"].isin(appeared_candidates)
    ].copy()

    # --------------------------------------------------
    # Qualification
    # --------------------------------------------------

    def qualification(row):

        score = row["Total"]

        category = str(
            row.get("Category", "")
        ).upper()

        special3 = str(
            row.get("Special3", "")
        ).upper()

        if category in ["SC", "ST"] or special3 == "PD":
            return "Qualified" if score >= 54 else "Not Qualified"

        return "Qualified" if score >= 72 else "Not Qualified"

    result["Qualification"] = result.apply(
        qualification,
        axis=1
    )

    if apply_cutoff:
        ranklist = result[
            result["Qualification"] == "Qualified"
        ].copy()
    else:
        ranklist = result.copy()

    # --------------------------------------------------
    # DOB
    # --------------------------------------------------

    if "DOB" in ranklist.columns:
        ranklist["DOB"] = pd.to_datetime(
            ranklist["DOB"],
            errors="coerce"
        )
    else:
        ranklist["DOB"] = pd.NaT

    # --------------------------------------------------
    # Sort
    # --------------------------------------------------

    ranklist = ranklist.sort_values(
        by=["Total", "DOB"],
        ascending=[False, True]
    ).reset_index(drop=True)

    # --------------------------------------------------
    # Sl.No
    # --------------------------------------------------

    ranklist["SlNo"] = range(
        1,
        len(ranklist) + 1
    )

    # --------------------------------------------------
    # Percentile
    # Same Score => Same Percentile
    # --------------------------------------------------

    total_candidates = len(ranklist)

    ranklist["ScoreRank"] = (
        ranklist["Total"]
        .rank(
            method="min",
            ascending=False
        )
    )

    if total_candidates <= 1:
        ranklist["Percentile"] = 100.00000
    else:
        ranklist["Percentile"] = ranklist[
            "ScoreRank"
        ].apply(
            lambda r: round(
                (
                    (total_candidates - r + 1)
                    / total_candidates
                ) * 100,
                5
            )
        )

    ranklist.drop(
        columns=["ScoreRank"],
        inplace=True
    )

    # --------------------------------------------------
    # Output
    # --------------------------------------------------

    output = pd.DataFrame({
        "Sl.No": ranklist["SlNo"],
        "App. No": ranklist["ApplNo"],
        "Roll No": ranklist["RollNo"],
        "Name": ranklist["Name"],
        "Part-I(Out of 200)": ranklist["Part1"],
        "Part-II(Out of 200)": ranklist["Part2"],
        "Part-III(Out of 160)": ranklist["Part3"],
        "Part-IV(Out of 160)": ranklist["Part4"],
        "Total(Out of 720)": ranklist["Total"],
        "Percentile": ranklist["Percentile"]
    })

    # --------------------------------------------------
    # Statistics
    # --------------------------------------------------

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Appeared",
        len(result)
    )

    c2.metric(
        "Published",
        len(ranklist)
    )

    c3.metric(
        "Highest Score",
        ranklist["Total"].max()
        if len(ranklist)
        else 0
    )

    st.dataframe(
        output,
        use_container_width=True
    )

    # --------------------------------------------------
    # Excel Download
    # --------------------------------------------------

    buffer = BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl"
    ) as writer:

        output.to_excel(
            writer,
            index=False,
            sheet_name="RankList"
        )

    st.download_button(
        "Download Rank List",
        buffer.getvalue(),
        file_name="KMAT_RankList.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

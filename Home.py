import streamlit as st

st.set_page_config(
    page_title="Scoping Review Pipeline",
    page_icon="📚",
    layout="wide",
)

# ==================================================
# Title
# ==================================================

st.title("📚 Scoping Review Pipeline")

st.markdown(
    """
This platform supports the complete scoping review workflow for
**Computational Thinking and Programming Education** studies.

Select the stage you want to work on.
"""
)

st.divider()

# ==================================================
# Stage cards
# ==================================================

col1, col2 = st.columns(2)

with col1:

    st.subheader("1️⃣ Stage 1 – Import & Rule-based Filtering")

    st.write(
        """
Import datasets from databases, standardize fields,
merge records, remove duplicates, perform initial
rule-based filtering, and prepare abstract screening input.
"""
    )

    st.page_link(
        "pages/1_Stage_1_Import_Filtering.py",
        label="Open Stage 1",
        icon="➡️"
    )

with col2:

    st.subheader("2️⃣ Stage 2 – Abstract Screening")

    st.write(
        """
Run LLM-assisted abstract screening,
majority voting, validation sampling,
and prepare the dataset for full-text screening.
"""
    )

    st.page_link(
        "pages/2_Stage_2_Abstract_Screening.py",
        label="Open Stage 2",
        icon="➡️"
    )

st.divider()

col3, col4 = st.columns(2)

with col3:

    st.subheader(
        "3️⃣ Stage 3 – Full-text Screening + Initial Pairing"
    )

    st.write(
        """
Perform full-text screening,
answer eligibility questions,
and detect possible CT–programming pairings.
"""
    )

    st.page_link(
        "pages/3_Stage_3_Fulltext_Initial_Pairing.py",
        label="Open Stage 3",
        icon="➡️"
    )

with col4:

    st.subheader("4️⃣ Stage 4 – Full Pairing Analysis")

    st.write(
        """
Conduct detailed CT–programming pairing analysis
using the screened full-text papers.
"""
    )

    st.page_link(
        "pages/4_Stage_4_Full_Pairing.py",
        label="Open Stage 4",
        icon="➡️"
    )

st.divider()

col5, col6 = st.columns(2)

with col5:

    st.subheader("5️⃣ Stage 5 – Characterization")

    st.write(
        """
Characterize interventions, operationalizations,
educational context, participants, tools, and outcomes.
"""
    )

    st.page_link(
        "pages/5_Stage_5_Characterization.py",
        label="Open Stage 5",
        icon="➡️"
    )

with col6:

    st.subheader("6️⃣ Stage 6 – Results Analysis")

    st.write(
        """
Perform final analyses,
cross-tabulations, visualizations,
and reporting preparation.
"""
    )

    st.page_link(
        "pages/6_Stage_6_Analysis.py",
        label="Open Stage 6",
        icon="➡️"
    )

# ==================================================
# Footer
# ==================================================

st.divider()

st.caption(
    "Scoping Review Pipeline for Computational Thinking "
    "and Programming Education Research"
)
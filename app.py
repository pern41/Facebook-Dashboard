import streamlit as st
import pandas as pd

from clean_facebook_comments2 import extract_comments
import clean_claude_comments2 as base


st.set_page_config(
    page_title="Facebook Comment Intelligence",
    layout="wide"
)


st.title("📊 Facebook Comment Intelligence")


text = st.text_area(
    "Paste Facebook Comments",
    height=300
)


if st.button("Analyze"):

    comments = extract_comments(text)

    rows=[]

    for comment in comments:

        row = base.classify_comment(comment)

        rows.append(row)


    df = pd.DataFrame(rows)


    st.success(

        f"Parsed {len(rows)} comments"

    )


    st.dataframe(
        df,
        use_container_width=True
    )

    st.subheader("Main Category")

    st.bar_chart(

        df["Main Category"]

        .value_counts()

    )

    st.subheader("Intent")

    st.bar_chart(

        df["Intent"]

        .value_counts()

    )

    st.subheader("Sentiment")

    st.bar_chart(

        df["Sentiment"]

        .value_counts()

    )

    csv = df.to_csv(

    index=False

    ).encode(

        "utf-8-sig"

    )

    csv = df.to_csv(
    index=False
    ).encode("utf-8-sig")

    st.download_button(

        "📥 Download CSV",

        csv,

        "facebook_comments.csv",

        "text/csv"

    )
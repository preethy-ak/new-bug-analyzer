import streamlit as st
import pandas as pd
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from io import BytesIO
from pptx import Presentation

# =========================
# 🔹 CATEGORY KEYWORDS
# =========================
CATEGORY_KEYWORDS = {
    "OMS": ["order", "checkout", "payment", "refund", "cart"],
    "IMS": ["inventory", "stock", "quantity", "warehouse"],
    "CMS": ["listing", "content", "image", "sku"],
    "PMS": ["discount", "voucher", "promotion", "pricing"],
    "Chat": ["chat", "buyer message", "customer service"],

    "Performance": ["slow", "lag", "timeout"],
    "Data/Sync": ["data mismatch", "not updating", "sync"],
    "UI/UX": ["ui", "ux", "alignment", "layout"],
    "Integration": ["api", "webhook"],
    "Backend": ["server error", "database"],
    "Frontend": ["javascript", "page crash"]
}

# =========================
# 🔹 CLASSIFICATION
# =========================
def classify_category(text):
    text = str(text).lower()

    # Step 1: core systems
    for category in ["OMS", "IMS", "CMS", "PMS", "Chat"]:
        for keyword in CATEGORY_KEYWORDS[category]:
            if keyword in text:
                return category

    # Step 2: tech
    for category, keywords in CATEGORY_KEYWORDS.items():
        if category in ["OMS", "IMS", "CMS", "PMS", "Chat"]:
            continue
        if any(k in text for k in keywords):
            return category

    return "Unknown"

# =========================
# 🔹 BUG FLAG
# =========================
def identify_bug(issue_type, summary):
    if "bug" in str(issue_type).lower():
        return "Bug"
    if any(k in str(summary).lower() for k in ["error", "fail", "issue"]):
        return "Bug"
    return "Non-Bug"

# =========================
# 🔹 REPEAT DETECTION
# =========================
def detect_repeats(df):
    vec = TfidfVectorizer(stop_words='english')
    df['Summary'] = df['Summary'].fillna("").astype(str)
df = df[df['Summary'].str.strip() != ""]

vectors = vec.fit_transform(df['Summary']).fillna(""))
    sim = cosine_similarity(vectors)

    df['Repeat_Count'] = [(sim[i] > 0.8).sum() for i in range(len(sim))]
    return df

# =========================
# 🔹 INSIGHTS ENGINE
# =========================
def generate_insights(df):
    insights = []

    top_category = df['Category'].value_counts().idxmax()
    insights.append(f"Most unstable system: {top_category}")

    repeat_issues = df[df['Repeat_Count'] > 3]
    if not repeat_issues.empty:
        insights.append(f"{len(repeat_issues)} repeated bugs (>3 times) detected")

    if 'Account Name' in df.columns:
        top_acc = df['Account Name'].value_counts().idxmax()
        insights.append(f"Top impacted account: {top_acc}")

    return insights

# =========================
# 🔹 RECOMMENDATIONS
# =========================
def generate_recommendations():
    return {
        "Short-Term": [
            "Fix high repeat bugs immediately",
            "Monitor unstable systems"
        ],
        "Long-Term": [
            "Implement root cause fixes",
            "Improve system stability"
        ],
        "Operational": [
            "Improve ticket quality",
            "Standardize bug reporting"
        ]
    }

# =========================
# 🔹 EXCEL EXPORT
# =========================
def generate_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Full Data")
        df['Category'].value_counts().to_excel(writer, sheet_name="Category")
        df.groupby(df['Created'].dt.to_period('M')).size().to_excel(writer, sheet_name="Trend")
output.seek(0)
return output
# =========================
# 🔹 PPT EXPORT
# =========================
def generate_ppt(df, insights, recs):
    prs = Presentation()

    # Title
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Jira Bug Analysis"

    # Summary
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Summary"
    tf = slide.placeholders[1].text = f"Total Tickets: {len(df)}"

    # Insights
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Insights"
    slide.placeholders[1].text = "\n".join(insights)

    # Recommendations
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Recommendations"
    slide.placeholders[1].text = "\n".join(recs["Short-Term"])

    output = BytesIO()
    prs.save(output)
output.seek(0)
return output
# =========================
# 🔹 STREAMLIT UI
# =========================
st.title("🚀 Jira Bug Analyzer")

file = st.file_uploader("Upload CSV", type=["csv"])

if file:
df = pd.read_csv(file)

# Normalize columns
df.columns = df.columns.str.strip().str.lower()

# Standardize column names
df.rename(columns={
    "summary": "Summary",
    "issue type": "Issue Type",
    "created": "Created",
    "reporter": "Reporter",
    "assignee": "Assignee",
    "account name": "Account Name",
    "seller id": "Account Name"
}, inplace=True)
    # Cleaning
    df['Summary'] = df['Summary'].fillna("")
df['Created'] = pd.to_datetime(df['Created'], errors='coerce')
df = df.dropna(subset=['Created'])
    # Processing
    df['Bug_Flag'] = df.apply(lambda x: identify_bug(x['Issue Type'], x['Summary']), axis=1)
df['Category'] = (
    df['Summary'].fillna('') + " " + df.get('Description', '').fillna('')
).apply(classify_category)    df = detect_repeats(df)

    # =========================
    # 🔹 FILTERS
    # =========================
    st.sidebar.header("Filters")

    if 'Account Name' in df.columns:
if 'Account Name' in df.columns:
    acc = st.sidebar.multiselect("Account", df['Account Name'].dropna().unique())
    if acc:
        df = df[df['Account Name'].isin(acc)]        if acc:
            df = df[df['Account Name'].isin(acc)]

    rep = st.sidebar.multiselect("Reporter", df['Reporter'].unique())
    if rep:
        df = df[df['Reporter'].isin(rep)]

    cat = st.sidebar.multiselect("Category", df['Category'].unique())
    if cat:
        df = df[df['Category'].isin(cat)]

    issue = st.sidebar.multiselect("Issue Type", df['Issue Type'].unique())
    if issue:
        df = df[df['Issue Type'].isin(issue)]

    date = st.sidebar.date_input("Date Range", [])
    if len(date) == 2:
        df = df[(df['Created'] >= pd.to_datetime(date[0])) &
                (df['Created'] <= pd.to_datetime(date[1]))]

    # =========================
    # 📊 DASHBOARD
    # =========================
    st.metric("Total Tickets", len(df))
    st.metric("Total Bugs", len(df[df['Bug_Flag']=="Bug"]))

    st.subheader("Trend")
    st.line_chart(df.groupby(df['Created'].dt.to_period('M')).size())

    st.subheader("Category")
    st.bar_chart(df['Category'].value_counts())

    st.subheader("Repeated Bugs")
    st.dataframe(df.sort_values("Repeat_Count", ascending=False).head(10))

    # =========================
    # 🔹 INSIGHTS
    # =========================
    insights = generate_insights(df)
    recs = generate_recommendations()

    st.subheader("Insights")
    for i in insights:
        st.write("- ", i)

    st.subheader("Recommendations")
    for k,v in recs.items():
        st.write(f"**{k}**")
        for item in v:
            st.write("-", item)

    # =========================
    # 📥 EXPORT
    # =========================
    excel = generate_excel(df)
    ppt = generate_ppt(df, insights, recs)

    st.download_button("Download Excel", excel, "report.xlsx")
    st.download_button("Download PPT", ppt, "report.pptx")

    st.dataframe(df)

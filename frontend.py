import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from main import run

st.set_page_config(layout='wide')
st.markdown("<h1 style='text-align: center; color: #575a5c; font-family: sans-serif; font-size: 40px; border: 3px'>MAIL CLASSIFICATION AND SUMMARIZATION DASHBOARD</h1>", unsafe_allow_html=True)

# Initialize session state for input_df
if "input_df" not in st.session_state:
    st.session_state.input_df = None

# Date pickers for selecting start and end dates
st.sidebar.markdown("### Select Date Range")
with st.sidebar.form(key="date_form"):
    start_date = st.date_input("Start Date", value=datetime.now())
    end_date = st.date_input("End Date", value=datetime.now())
    submit_button = st.form_submit_button(label="Apply Date Range")

# Validate date range and trigger backend logic
if submit_button:
    if start_date > end_date:
        st.sidebar.error("Start date must be before or equal to end date.")
    else:
        # Call the backend `run` function with the selected dates
        st.session_state.input_df = run(start_date.strftime("%d-%m-%Y"), end_date.strftime("%d-%m-%Y"))

# Check if input_df is available
if st.session_state.input_df is None:
    st.info("Please select a date range and click 'Apply Date Range' to load emails.")
elif st.session_state.input_df.empty:
    st.warning("No emails found for the selected date range.")
else:
    input_df = st.session_state.input_df

    # Dropdown for selecting importance
    importance_filter = st.selectbox("Select Importance Level:", options=["All"] + input_df["Importance"].unique().tolist())

    # Filter the dataframe based on importance
    if importance_filter != "All":
        filtered_df = input_df[input_df["Importance"] == importance_filter]
    else:
        filtered_df = input_df

    # Dropdown for selecting subject
    subject_filter = st.selectbox("Select Subject:", options=filtered_df["Subject"].unique().tolist())

    # Create tabs for better organization
    tab1, tab2, tab3 = st.tabs(["Mail Details", "Visualization", "Summarization"])

    # Tab 1: Mail Details
    with tab1:
        st.markdown("### Mail Details")
        if subject_filter:
            mail_content = filtered_df[filtered_df["Subject"] == subject_filter]["Latest Mail Content"].iloc[0]
            st.markdown(f"### Latest Mail Content for Subject: {subject_filter}")
            st.write(mail_content)

            # Display additional details for the selected mail
            selected_mail = filtered_df[filtered_df["Subject"] == subject_filter].iloc[0]
            st.markdown("### Additional Details for Selected Mail")
            st.write(f"**Suspicious:** {selected_mail['Suspicious']}")
            st.write(f"**Importance:** {selected_mail['Importance']}")
            st.write(f"**Classification:** {selected_mail['Classification']}")

    # Tab 2: Overall Visulaization
    with tab2:
        st.markdown("### Visualization")
        col1, col2, col3 = st.columns([1, 1, 1])

        suspicious_counts = input_df["Suspicious"].value_counts()
        fig_sus = px.pie(values=suspicious_counts.values, names=suspicious_counts.index, title='Suspicious Email Distribution')
        col1.plotly_chart(fig_sus, use_container_width=True)

        importance_counts = input_df["Importance"].value_counts()
        fig_imp = px.pie(values=importance_counts.values, names=importance_counts.index, title='Importance Level Distribution')
        col2.plotly_chart(fig_imp, use_container_width=True)

        classification_counts = input_df["Classification"].value_counts()
        fig_class = px.pie(values=classification_counts.values, names=classification_counts.index, title='Email Classification Distribution')
        col3.plotly_chart(fig_class, use_container_width=True)

    # Tab 3: Summarization (Optional for future use)
    with tab3:
        st.markdown("### Summarization")
        st.write("You can add the mail summarization here.")
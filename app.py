import streamlit as st
import pandas as pd

# App title
st.title("Dummy Streamlit App")

# Introduction text
st.write("This is a simple dummy Streamlit app for demonstration purposes.")

# Create a sample dataframe
data = {
    "Name": ["Alice", "Bob", "Charlie"],
    "Age": [25, 30, 35],
    "City": ["New York", "Paris", "London"]
}
df = pd.DataFrame(data)

# Display the dataframe
st.write("Here is a sample data table:")
st.dataframe(df)

# Add a simple interactive element: slider
age_filter = st.slider("Select minimum age", min_value=0, max_value=100, value=20)
filtered_df = df[df["Age"] >= age_filter]

# Display filtered dataframe
st.write("Filtered data based on age:")
st.dataframe(filtered_df)

# Add a button
if st.button("Say Hello"):
    st.write("Hello, Streamlit user!")

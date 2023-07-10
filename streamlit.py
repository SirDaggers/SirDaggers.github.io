import pandas as pd
import streamlit as st
import plotly.express as px
from PIL import Image

st.set_page_config(page_title='Scheduling System')

# Create an empty container above the header
container = st.empty()

# Add picture to the container with custom dimensions
image_path = "Deped.png"
image_width = 100  # Adjust the width as desired
image_height = 100  # Adjust the height as desired

image_width1 = 100  # Adjust the width as desired
image_height1 = 100  # Adjust the height as desired
image = Image.open("Mambog.png")
resized_image = image.resize((image_width, image_height))

# Center the container
st.markdown(
    """
    <style>
    .centered-container {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        height: 100vh;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Create the centered container
with st.container() as centered_container:
    # Create columns for the image and text
    col1, col2, col3 = st.columns([1, 2, 0.5])

    # Add the image to the first column
    col1.image(image_path, width=image_width, caption="")

    # Add text to the second column
    with col2:
        st.markdown(
            """
            <style>
            .centered-text {
                text-align: center;
                line-height: 0.8;
                margin-left: -70px;
            }
            </style>
            """,
            unsafe_allow_html=True
        )
        st.write(f'<p class="centered-text" style="margin-bottom: 0;">REPUBLIC OF THE PHILIPPINES</p>', unsafe_allow_html=True)
        st.write(f'<p class="centered-text" style="margin-bottom: 0;">DEPARTMENT OF EDUCATION</p>', unsafe_allow_html=True)
        st.write(f'<p class="centered-text" style="margin-bottom: 0;">REGION IV-A CALABARZON</p>', unsafe_allow_html=True)
        st.write(f'<p class="centered-text" style="margin-bottom: 0;">CITY SCHOOLS DIVISION OF BACOOR</p>', unsafe_allow_html=True)
        st.write(f'<p class="centered-text" style="margin-bottom: 0;">DISTRICT OF BACOOR II</p>', unsafe_allow_html=True)

    # Add another image to the third column
    with col3:
        st.image(resized_image)
        st.markdown(
            """
            <style>
            .image-container {
                display: flex;
                justify-content: flex-end;
                margin-right: 70px;
            }
            </style>
            """,
            unsafe_allow_html=True
        )
        st.markdown('<div class="image-container"></div>', unsafe_allow_html=True)


st.markdown("<h1 style='text-align: center;'>Mambog Elementary School</h1>", unsafe_allow_html=True)

excel_file = 'schedule_results.xlsx'
sheet_name = 'Sheet'

df_sect1 = pd.read_excel(excel_file,
                         sheet_name = sheet_name,
                         usecols = 'A:E',
                         header = 1,
                         nrows = 9)

df_secthead = pd.read_excel(excel_file,
                         sheet_name = sheet_name,
                         usecols = 'A',
                         header = None)

cell_value1 = df_secthead.iloc[0, 0]
st.markdown(f"<p style='text-align: center;'>{cell_value1}</p>", unsafe_allow_html=True)

# Remove index column
df_no_index = df_sect1.copy()
df_no_index.reset_index(drop=True, inplace=True)

# Convert DataFrame to HTML table without index
table_html = df_no_index.to_html(index=False)

# Center align the header and elements in the table using CSS
centered_table_html = table_html.replace('<table', '<table style="margin-left:auto;margin-right:auto;text-align:center;"')
centered_table_html = centered_table_html.replace('<th>', '<th style="text-align:center;">')

# Display the HTML table with centered header and elements using Markdown
st.markdown(centered_table_html, unsafe_allow_html=True)

df_sect2 = pd.read_excel(excel_file,
                         sheet_name = sheet_name,
                         usecols = 'A:E',
                         header = 13,
                         nrows = 9)


cell_value2 = df_secthead.iloc[12, 0]
st.write("")
st.markdown(f"<p style='text-align: center;'>{cell_value2}</p>", unsafe_allow_html=True)

# Remove index column
df_no_index = df_sect2.copy()
df_no_index.reset_index(drop=True, inplace=True)

# Convert DataFrame to HTML table without index
table_html = df_no_index.to_html(index=False)

# Center align the header and elements in the table using CSS
centered_table_html = table_html.replace('<table', '<table style="margin-left:auto;margin-right:auto;text-align:center;"')
centered_table_html = centered_table_html.replace('<th>', '<th style="text-align:center;">')

# Display the HTML table with centered header and elements using Markdown
st.markdown(centered_table_html, unsafe_allow_html=True)

df_sect3 = pd.read_excel(excel_file,
                         sheet_name = sheet_name,
                         usecols = 'A:E',
                         header = 25,
                         nrows = 9)

cell_value3 = df_secthead.iloc[24, 0]
st.write("")
st.markdown(f"<p style='text-align: center;'>{cell_value3}</p>", unsafe_allow_html=True)

# Remove index column
df_no_index = df_sect3.copy()
df_no_index.reset_index(drop=True, inplace=True)

# Convert DataFrame to HTML table without index
table_html = df_no_index.to_html(index=False)

# Center align the header and elements in the table using CSS
centered_table_html = table_html.replace('<table', '<table style="margin-left:auto;margin-right:auto;text-align:center;"')
centered_table_html = centered_table_html.replace('<th>', '<th style="text-align:center;">')

# Display the HTML table with centered header and elements using Markdown
st.markdown(centered_table_html, unsafe_allow_html=True)

df_sect4 = pd.read_excel(excel_file,
                         sheet_name = sheet_name,
                         usecols = 'A:E',
                         header = 37,
                         nrows = 9)

cell_value4 = df_secthead.iloc[36, 0]
st.write("")
st.markdown(f"<p style='text-align: center;'>{cell_value4}</p>", unsafe_allow_html=True)

# Remove index column
df_no_index = df_sect4.copy()
df_no_index.reset_index(drop=True, inplace=True)

# Convert DataFrame to HTML table without index
table_html = df_no_index.to_html(index=False)

# Center align the header and elements in the table using CSS
centered_table_html = table_html.replace('<table', '<table style="margin-left:auto;margin-right:auto;text-align:center;"')
centered_table_html = centered_table_html.replace('<th>', '<th style="text-align:center;">')

# Display the HTML table with centered header and elements using Markdown
st.markdown(centered_table_html, unsafe_allow_html=True)

df_sect5 = pd.read_excel(excel_file,
                         sheet_name = sheet_name,
                         usecols = 'A:E',
                         header = 49,
                         nrows = 9)

cell_value5 = df_secthead.iloc[48, 0]
st.write("")
st.markdown(f"<p style='text-align: center;'>{cell_value5}</p>", unsafe_allow_html=True)

# Remove index column
df_no_index = df_sect5.copy()
df_no_index.reset_index(drop=True, inplace=True)

# Convert DataFrame to HTML table without index
table_html = df_no_index.to_html(index=False)

# Center align the header and elements in the table using CSS
centered_table_html = table_html.replace('<table', '<table style="margin-left:auto;margin-right:auto;text-align:center;"')
centered_table_html = centered_table_html.replace('<th>', '<th style="text-align:center;">')

# Display the HTML table with centered header and elements using Markdown
st.markdown(centered_table_html, unsafe_allow_html=True)

import streamlit as st

from scraper import setup_driver, navigate_to_page, extract_final_page_table, save_to_excel

st.title("Koshvani Scraper")

st.write("Enter details to extract expenditure data")

# INPUT FIELDS
expenditure_wise = st.selectbox(
    "Expenditure Type",
    ["Grant-wise"]
)
financial_year=st.selectbox(
    "Financial Year",
    ["2025-2026",
        "2024-2025",
        "2023-2024"]
                            )
grant_no = st.text_input("Grant Number")

scheme_code = st.text_input("Scheme Code")

city = st.text_input("Treasury / City")

# RUN BUTTON
if st.button("Run Script"):

    st.info("Starting Script...")

    driver, wait = setup_driver()

    try:

        navigate_to_page(
            driver,
            wait,financial_year,
            expenditure_wise,
            grant_no,
            scheme_code,
            city
        )

        data = extract_final_page_table(driver, wait)

        if data:

            save_to_excel(data, city)

            st.success("Data extracted successfully!")

            with open(f"{city}.xlsx", "rb") as f:
                st.download_button(
                    "Download Excel",
                    f,
                    file_name=f"{city}.xlsx"
                )

    except Exception as e:
        st.error(str(e))

    finally:
        driver.quit()
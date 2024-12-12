import openpyxl
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
import time


def get_suggestions(keyword, file_path, sheet_name):
    driver = webdriver.Chrome()

    try:
        # Navigate to Google
        driver.get("https://www.google.com")
        search_box = driver.find_element(By.NAME, "q")
        search_box.clear()
        search_box.send_keys(keyword)
        time.sleep(2)

        # Collect suggestions
        suggestions = driver.find_elements(
            By.CSS_SELECTOR, 'div[role="presentation"] span')
        longest = ""
        shortest = None

        for suggestion in suggestions:
            text = suggestion.text.strip()
            if text:
                if len(text) > len(longest):
                    longest = text
                if shortest is None or len(text) < len(shortest):
                    shortest = text

        print("Longest Suggestion: ", longest)
        print("Shortest Suggestion: ", shortest)

        # Update Excel
        wb = openpyxl.load_workbook(file_path)
        sheet = wb[sheet_name]

        for row in range(2, sheet.max_row + 1):
            if sheet.cell(row=row, column=1).value == keyword:
                sheet.cell(row=row, column=2).value = longest
                sheet.cell(row=row, column=3).value = shortest
                break

        wb.save(file_path)
        return longest, shortest

    finally:
        driver.quit()


if __name__ == "__main__":
    excel_file = "keywords.xlsx"
    today = datetime.now().strftime("%A")

    try:
        wb = openpyxl.load_workbook(excel_file)

        if today not in wb.sheetnames:
            print(f"Sheet for {today} does not exist. Exiting the script.")
        else:
            sheet = wb[today]
            keywords = [sheet.cell(row=row, column=1).value for row in range(
                2, sheet.max_row + 1)]

            for keyword in keywords:
                if keyword:
                    print(f"Processing keyword: {keyword}")
                    get_suggestions(keyword, excel_file, today)

            print("Excel file has been updated.")
    except Exception as e:
        print(f"Error while processing the Excel file: {e}")

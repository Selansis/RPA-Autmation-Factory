from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
prefs = {"profile.default_content_settings.popups": 0,
             "download.default_directory": 
                        r"C:\Users\wojte\Pulpit\Pulpit\RPA Automation Factory\\",#IMPORTANT - ENDING SLASH V IMPORTANT
             "directory_upgrade": True}
options.add_experimental_option("prefs", prefs)


options.add_experimental_option('excludeSwitches', ['enable-logging'])

driver = webdriver.Chrome(
    options=options
)


# otwarcie strony rpachallenge.com
driver.get("http://www.rpachallenge.com/")

# oczekiwanie na przycisk "Download Excel"
download_excel_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html[1]/body[1]/app-root[1]/div[2]/app-rpa1[1]/div[1]/div[1]/div[6]/a[1]')))
download_excel_button.click()


# rozpoczęcie zadania klikając przycisk START
start_button = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, '/html[1]/body[1]/app-root[1]/div[2]/app-rpa1[1]/div[1]/div[1]/div[6]/button[1]')))
start_button.click()

# wczytanie danych z pliku Excel
df = pd.read_excel(r"C:\Users\wojte\Pulpit\Pulpit\RPA Automation Factory\challenge.xlsx")

# iteracja po danych i wprowadzenie ich do formularza
for index, row in df.iterrows():
    first_name = row['First Name']
    last_name = row['Last Name ']
    company_name = row['Company Name']
    role = row['Role in Company']
    address = row['Address']
    email = row['Email']
    phone = str(row['Phone Number'])
    

    driver.find_element(By.XPATH, "//label[contains(text(), 'First Name')]/following-sibling::input").send_keys(first_name)
    driver.find_element(By.XPATH,  "//label[contains(text(), 'Last Name')]/following-sibling::input").send_keys(last_name)
    driver.find_element(By.XPATH, "//label[contains(text(), 'Company Name')]/following-sibling::input").send_keys(company_name)
    driver.find_element(By.XPATH, "//label[contains(text(), 'Role in Company')]/following-sibling::input").send_keys(role)
    driver.find_element(By.XPATH, "//label[contains(text(), 'Address')]/following-sibling::input").send_keys(address)
    driver.find_element(By.XPATH, "//label[contains(text(), 'Email')]/following-sibling::input").send_keys(email)
    driver.find_element(By.XPATH, "//label[contains(text(), 'Phone Number')]/following-sibling::input").send_keys(phone)
    confirm_button = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, '/html[1]/body[1]/app-root[1]/div[2]/app-rpa1[1]/div[1]/div[2]/form[1]/input[1]')))
    confirm_button.click()


# pobranie wyniku
result = driver.find_element(By.XPATH, "/html[1]/body[1]/app-root[1]/div[2]/app-rpa1[1]/div[1]/div[2]/div[2]").text

# zapisanie wyniku do pliku tekstowego
with open('result.txt', 'w') as file:
    file.write(result)

# zamknięcie przeglądarki
driver.quit()

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
options = webdriver.ChromeOptions()

options.add_experimental_option('excludeSwitches', ['enable-logging'])

driver = webdriver.Chrome(
    options=options
)


# Otwarcie strony rpachallenge.com
driver.get("http://www.rpachallenge.com/")

# Oczekiwanie na przycisk "Download Excel"
#download_excel_button = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, 'button[ng-reflect-message="Download Excel"]')))
download_excel_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html[1]/body[1]/app-root[1]/div[2]/app-rpa1[1]/div[1]/div[1]/div[6]/a[1]')))
download_excel_button.click()


# Rozpoczęcie zadania klikając przycisk START
start_button = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, '/html[1]/body[1]/app-root[1]/div[2]/app-rpa1[1]/div[1]/div[1]/div[6]/button[1]')))
start_button.click()

# Wczytanie danych z pliku Excel
df = pd.read_excel("M:\Downloads\challenge.xlsx")

# Iteracja po danych i wprowadzenie ich do formularza
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
    #driver.find_element(By.XPATH, '/html[1]/body[1]/app-root[1]/div[2]/app-rpa1[1]/div[1]/div[2]/form[1]/div[1]/div[4]/rpa1-field[1]/div[1]/input[1]').click()

# Pobranie wyniku
result = driver.find_element(By.XPATH, "/html[1]/body[1]/app-root[1]/div[2]/app-rpa1[1]/div[1]/div[2]/div[2]").text

# Zapisanie wyniku do pliku tekstowego
with open('result.txt', 'w') as file:
    file.write(result)

# Zamknięcie przeglądarki
driver.quit()

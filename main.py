import threading
import pickle
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tkinter import *
import tkinter as tk
from tkinter import scrolledtext
from tkinter import messagebox
import math
from time import sleep
from selenium.webdriver.chrome.options import Options
import undetected_chromedriver as uc
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
import time
import atexit
import psutil
import pyperclip
from datetime import date

# --- Configuration --- #
MAIN_PAGE = "https://harb.cma.gov.il/sso/Auth/Agent"
BASE_DOMAIN = "https://harb.cma.gov.il/"
COOKIE_FILE = "site_cookies.pkl"
LIFE_PAGE = 'https://harb.cma.gov.il/sso/Results?id=5'
HEALTH_PAGE = 'https://harb.cma.gov.il/sso/Results?id=6'

driver = None

SING_IN_ID = 'REPLACE WITH YOUR ID'
SIGN_IN_PASSWORD = 'REPLACE WITH THE PASSWORD TO GOV SITE'                               # REMEMBER TO SET THE RIGHT CREDENTIALS UPON UPLOADING CODE

SMS_COMPLETE_FLAG = threading.Event()

def cleanup(): # Cleanup function for freeing up resources
    global driver
    if driver:
        print("Cleaning up driver...")
        try:
            driver.quit()
        except:
            pass
        driver = None
    current_process = psutil.Process()
    children = current_process.children(recursive=True)
    for child in children:
        try:
            child.terminate()
        except:
            pass

def get_undetected_driver(): # The function to initiate the chrome driver
    options = Options()
    options.add_argument("--disable-blink-features=AutomationControlled") # Disable blink feature for automation to avoid 'bot detection' by the gov site
    options.add_argument("--log-level=3") # remove site logs to prevent unescesery console output
    # We don't use headless mode since the gov site is detecting it as 'bot' and blocks automation
    
    try:
        global driver
        driver = uc.Chrome(options=options, use_subprocess=True, version_main=143) 
        return driver
    except Exception as e:
        print(f"Failed to initialize undetected_chromedriver: {e}")
        return None

def convert_company_name(name): # A function to shorten the name of the insurance company
    if "איי אי ג'י ישראל חברה לביטוח" in name:
        return "AIG"
    elif "הפניקס חברה לביטוח " in name:
        return "הפניקס"
    elif "הראל חברה לביטוח" in name:
        return "הראל"
    elif "מנורה מבטחים " in name:
        return "מנורה"
    elif "כלל חברה לביטוח" in name:
        return "כלל"
    elif "איילון חברה לביטוח" in name:
        return "איילון"
    elif "איי. די. איי. חברה לביטוח" in name:
        return "ביטוח ישיר"
    elif "מגדל חברה לביטוח " in name:
        return "מגדל"
    elif "הכשרה חברה לביטוח" in name:
        return "הכשרה"
    else:
        return name
    
def save_session_cookies(driver): # A function to save the site cookies to prevent another log in
    print("--- Saving cookies... ---")
    with open(COOKIE_FILE, 'wb') as file:
        pickle.dump(driver.get_cookies(), file) # Save cookie as .pkl file
    print(f"Cookies saved to {COOKIE_FILE}")

def verify_and_retry_selection(driver, wait, wrapper_element, expected_value, max_retries=999): # A function to select and check that the 'select' element is the right value
    expected_value_str = str(expected_value)
    
    VISIBLE_INPUT_CSS = 'span.k-input' 

    def perform_selection(wrapper):
        wrapper.click()
        LISTBOX_ID = wrapper.get_attribute("aria-owns")
        
        LIST_ITEM_XPATH = (
            By.XPATH,
            f'//div[contains(@class, "k-popup") and contains(@style, "display: block")]//ul[@id="{LISTBOX_ID}"]/li[text()="{expected_value_str}"]'
        )
        
        list_item = wait.until(EC.element_to_be_clickable(LIST_ITEM_XPATH))
        list_item.click()
        sleep(0.3) # Wait for JS update

    for attempt in range(1, max_retries + 1):
        try:
            perform_selection(wrapper_element)
            
            visible_input = wrapper_element.find_element(By.CSS_SELECTOR, VISIBLE_INPUT_CSS)
            if visible_input.text.strip() == expected_value_str:
                return True

            input_id = wrapper_element.get_attribute("aria-controls")
            if input_id:
                hidden_input = driver.find_element(By.ID, input_id)
                if hidden_input.get_attribute('value') == expected_value_str:
                    return True
        except StaleElementReferenceException:
            pass
            
        except Exception as e:
            print(f"Error : {e}")
        
    print(f"CRITICAL FAILURE: Failed to set value to {expected_value_str} after all retries.")
    return False

def get_lead(client_name, client_id, birth_date, id_date, phone, insurance_company, reason): # The main function for outputing values for health and life insurences
    global driver
    driver = get_undetected_driver()

    assert driver is not None, "Failed to start driver" # A pretty critical error -> use assert to stop the app

    try:
        is_logged_in = load_session_cookies(driver)
        if not is_logged_in: # If the load cookies return False -> it means there is no cookie present, meaining authentication is needed again.
            return "log_in"
        
        wait = WebDriverWait(driver, 10)

        driver.get(MAIN_PAGE)

        WebDriverWait(driver, 5).until(EC.url_matches("https://harb.cma.gov.il/sso/Auth/Agent|https://harb.cma.gov.il/Home/Index"))

        if driver.current_url == "https://harb.cma.gov.il/Home/Index?agent=NotAgent": # Sometimes this page will show instead, just click a couple buttons to go past it
            try:
                close_selector = "#MsgAlertModal button[aria-label='Close']"
        
                close_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, close_selector)))
                close_button.click()
                print("Modal closed successfully.")
                close_button.click()
            except Exception as e:
                print(f"Standard click failed, attempting JS click. Error: {e}")
                # Backup: JavaScript click ignores "overlays" or animation states
                close_button = driver.find_element(By.CSS_SELECTOR, "#MsgAlertModal button[aria-label='Close']")
                driver.execute_script("arguments[0].click();", close_button)
            try:
                sales_link = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.sales-link[href*='MoreInsurance']")))
                
                sales_link.click()
                print("Successfully clicked 'כניסת מורשים'")

            except Exception as e:
                print("Normal click blocked, trying JavaScript click...")
                sales_link = driver.find_element(By.CSS_SELECTOR, "a.sales-link")
                driver.execute_script("arguments[0].click();", sales_link)
            time.sleep(1)
            driver.get(MAIN_PAGE)
        elif driver.current_url.startswith("https://login.gov.il"): # If it's back at the original gov site, the authentication for login has failed.
            return "log_in"

        try:
            id_input = wait.until(EC.visibility_of_element_located((By.ID, 'txtId'))) # Wait for the page to load and have the inputs present
        except:
            if driver.current_url.startswith("https://login.gov.il"): # Sometimes the right url will show but than will immideatly change back to the gov log in page so we check it again
                return "log_in"
        
        id_input.clear()
        id_input.send_keys(client_id)

        birth_day = birth_date.split('/')[0]
        birth_month = birth_date.split('/')[1]
        birth_year = birth_date.split('/')[2]

        id_day = id_date.split('/')[0]
        id_month = id_date.split('/')[1]
        id_year = id_date.split('/')[2]

        DAY_WRAPPER_SELECTOR = (By.CSS_SELECTOR, 'span.k-dropdown.ddlDate[aria-owns="uiDdlDay_listbox"]')
        dropdown_wrapper_day = wait.until(EC.element_to_be_clickable(DAY_WRAPPER_SELECTOR))
        if not verify_and_retry_selection(driver, wait, dropdown_wrapper_day, birth_day):
            print(f"Failed to set Birth Day to {birth_day}")
            driver.quit()
            return

        MONTH_WRAPPER_SELECTOR = (By.CSS_SELECTOR, 'span.k-dropdown.uiDdlMonth[aria-owns="uiDdlMonth_listbox"]')
        dropdown_wrapper_month = wait.until(EC.element_to_be_clickable(MONTH_WRAPPER_SELECTOR))
        if not verify_and_retry_selection(driver, wait, dropdown_wrapper_month, birth_month):
            print(f"Failed to set Birth Month to {birth_month}")
            driver.quit()
            return

        YEAR_WRAPPER_SELECTOR = (By.CSS_SELECTOR, 'span.k-dropdown.uiDdlYear[aria-owns="uiDdlYear_listbox"]')
        dropdown_wrapper_year = wait.until(EC.element_to_be_clickable(YEAR_WRAPPER_SELECTOR))
        if not verify_and_retry_selection(driver, wait, dropdown_wrapper_year, birth_year):
            print(f"Failed to set Birth Year to {birth_year}")
            driver.quit()
            return

        # Find the unique parent container for the Issue Date fields
        ISSUE_DATE_PARENT_XPATH = '//div[span[@class="subTitle gray"][contains(text(), "תאריך הנפקה")]]/following-sibling::div[1]'
        ISSUE_DATE_PARENT_SELECTOR = (By.XPATH, ISSUE_DATE_PARENT_XPATH)
        issue_date_parent = wait.until(EC.presence_of_element_located(ISSUE_DATE_PARENT_SELECTOR))

        DAY_WRAPPER_CSS = 'span.k-dropdown.ddlDate[aria-owns="uiDdlDay_listbox"]'
        dropdown_wrapper_day = issue_date_parent.find_element(By.CSS_SELECTOR, DAY_WRAPPER_CSS) 
        if not verify_and_retry_selection(driver, wait, dropdown_wrapper_day, id_day):
            print(f"Failed to set ID Issue Day to {id_day}")
            driver.quit()
            return

        MONTH_WRAPPER_CSS = 'span.k-dropdown.uiDdlMonth[aria-owns="uiDdlMonth_listbox"]'
        dropdown_wrapper_month = issue_date_parent.find_element(By.CSS_SELECTOR, MONTH_WRAPPER_CSS) 
        if not verify_and_retry_selection(driver, wait, dropdown_wrapper_month, id_month):
            print(f"Failed to set ID Issue Month to {id_month}")
            driver.quit()
            return

        YEAR_WRAPPER_CSS = 'span.k-dropdown.uiDdlYear[aria-owns="uiDdlYear_listbox"]'
        dropdown_wrapper_year = issue_date_parent.find_element(By.CSS_SELECTOR, YEAR_WRAPPER_CSS) 
        if not verify_and_retry_selection(driver, wait, dropdown_wrapper_year, id_year):
            print(f"Failed to set ID Issue Year to {id_year}")
            driver.quit()
            return

        terms_checkbox = driver.find_element(By.ID, 'cbAproveTerm')
        terms_checkbox.click()

        driver.find_element(By.ID, 'butIdent').click()

        ID_BOTTOM = (By.ID, 'niBottom')
        ERROR_POPUP = (By.XPATH, "//div[contains(@class, 'modal fade show')]")

        try:
            WebDriverWait(driver, 5).until(  # After filling the entire 'form' for extracting insurances and pressing the button to continue -> wait for the page to load, which should have any of these elements present
                EC.any_of(
                    EC.visibility_of_element_located(ID_BOTTOM),
                    EC.visibility_of_element_located(ERROR_POPUP),
                    EC.url_contains("sso/Overview")
                )
            )

            # Now we check if the credentials were correct, if they don't -> the credentials were wrong and the gov site has not found any person
            if driver.find_elements(*ID_BOTTOM):
                print("Success: Bottom element is visible.")
                
            elif driver.find_elements(By.XPATH, "//div[contains(@class, 'modal fade show')]"):
                print("Error: Popup appeared.")
                return f"{client_name}\t{convert_id_back(client_id)}\t{phone}\t{convert_date_back(birth_date)}\t{reason}\t{insurance_company}"
                
            elif "sso/Overview" in driver.current_url:
                print("Redirected: Didn't work, back at Overview.")
                return f"{client_name}\t{convert_id_back(client_id)}\t{phone}\t{convert_date_back(birth_date)}\t{reason}\t{insurance_company}"

            WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.ID, 'niBottom')))
        except:
            try:
                assert WebDriverWait(driver, 1).until(EC.url_matches("https://harb.cma.gov.il/sso/Overview")), "Error while continuing" # Generally there is no reason to end here, for that i use assert meaning something very wrong happend
            except:
                driver.quit()
                

        driver.get(LIFE_PAGE)                                   #      ------------------- FIND THE LIFE INSURANCE DATA ---------------------------

        WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, 'uiResultsBodyDiv-5')))
        
        Life_Results = ""
        Health_Results = ""

        ALL_RECORDS_XPATH = "//div[contains(@class, 'sub') and contains(@class, 'borderUpLife')]" # Each insurance is covered in this element
        NO_RECORDS_XPATH = "//div[contains(@class, 'alignCenter') and contains(normalize-space(), 'אין תוצאות בתחום זה')]" # If there are no insurances -> this will appear instead

        COMBINED_XPATH = f"{ALL_RECORDS_XPATH} | {NO_RECORDS_XPATH}"
        # Find all matching elements once they are present
        try:
            element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, COMBINED_XPATH))
            )
            if "alignCenter" in element.get_attribute("class"):
                print("Confirmed: No records found (displayed the empty state message).")
                record_containers = []
            else:
                # If it's not the 'No Records' div, it must be the actual data
                print("Records found! Processing...")
                record_containers = driver.find_elements(By.XPATH, ALL_RECORDS_XPATH)
        except Exception:
            print("Timed out: Neither records nor the 'No Results' message appeared.")
            record_containers = []
        
        XPATH_PRODUCT_NAME = ".//div[@class='font17 bold colorTitle ']"
        XPATH_PREMIUM_AMOUNT = ".//div[@class='premyaField']"
        XPATH_COMPANY_IMG = ".//img[contains(@class, 'tooltip')]" # We'll extract the 'alt' attribute from this -> this is actually the name of the company

        life_data = []

        # Iterate through each record container and extract data
        for i, container in enumerate(record_containers):
            try:
                product_element = container.find_element(By.XPATH, XPATH_PRODUCT_NAME)
                product_name = product_element.text.strip().replace('\n', ' ')

                premium_element = container.find_element(By.XPATH, XPATH_PREMIUM_AMOUNT)

                premium_text = premium_element.text.strip().split('\n')[0].replace('₪', '').strip()
                
                company_img_element = container.find_element(By.XPATH, XPATH_COMPANY_IMG)
                company_name = company_img_element.get_attribute('alt').strip()

                company_name = convert_company_name(company_name)

                life_data.append({
                    "Product Name": product_name,
                    "Premium Amount": premium_text,
                    "Company Name": company_name
                })

            except Exception as e:
                continue

        company_totals: dict[str, float] = {}
        for result in life_data:
            product_name = str(result['Product Name']).replace('(', '').replace(')', '')
            company_name = result['Company Name']
            premium_amount_str = result['Premium Amount']

            print(f"  Product: {product_name}")
            print(f"  Premium: {premium_amount_str}")
            print(f"  Original Company: {company_name}")
            company_name = convert_company_name(company_name)
            print(f"  Shortened Company: {company_name}")

            try:
                premium_value = float(premium_amount_str)
                full_key = f"{company_name}-{product_name}"
                company_totals[full_key] = company_totals.get(full_key, 0) + premium_value # Add for the same type and company to the premium value found (combine them all)
                    
            except ValueError:
                print(f"Could not convert '{premium_amount_str}' to a number. Skipping aggregation for this item.")

        # Add all of the founded information to the results with '+' in between each one, this will be the output
        total_keys = len(company_totals)
        current_key = 1
        for company, total_premium in company_totals.items():
            rounded_amount = math.floor(total_premium)
            if current_key < total_keys:
                result = (f"{company.split('-')[1]} {rounded_amount} {company.split('-')[0]} + ")
            else:
                result = (f"{company.split('-')[1]} {rounded_amount} {company.split('-')[0]}")
            print(f"Total premiums : {result}")

            Life_Results += result
            current_key += 1

        driver.get(HEALTH_PAGE)                                   #      ------------------- FIND THE HEALTH INSURANCE DATA ---------------------------

        WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, 'uiResultsBodyDiv-6')))

        ALL_RECORDS_XPATH = "//div[contains(@class, 'sub') and contains(@class, 'borderUpHealth')]"
        NO_RECORDS_XPATH = "//div[contains(@class, 'alignCenter') and contains(normalize-space(), 'אין תוצאות בתחום זה')]"

        COMBINED_XPATH = f"{ALL_RECORDS_XPATH} | {NO_RECORDS_XPATH}"
        # Find all matching elements once they are present
        try:
            element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, COMBINED_XPATH)))
            if "alignCenter" in element.get_attribute("class"):
                print("Confirmed: No records found (displayed the empty state message).")
                record_containers = []
            else:
                # If it's not the 'No Records' div, it must be the actual data
                print("Records found! Processing...")
                record_containers = driver.find_elements(By.XPATH, ALL_RECORDS_XPATH)
        except Exception:
            print("Timed out: Neither records nor the 'No Results' message appeared.")
            record_containers = []
        
        XPATH_PRODUCT_NAME = ".//div[@class='font17 bold colorTitle ']"
        XPATH_PREMIUM_AMOUNT = ".//div[@class='premyaField']"
        XPATH_COMPANY_IMG = ".//a[@class='companyLink']/img[@class='tooltip']"

        life_data = []

        # Iterate through each record container and extract data
        for i, container in enumerate(record_containers):
            try:
                product_element = container.find_element(By.XPATH, XPATH_PRODUCT_NAME)
                product_name = product_element.text.strip().replace('\n', ' ')

                premium_element = container.find_element(By.XPATH, XPATH_PREMIUM_AMOUNT)
                
                premium_text = premium_element.text.strip().split('\n')[0].replace('₪', '').strip()
                
                # Wait until the image is actually present inside this container, we assign it directly from the wait return
                company_img_element = wait.until(lambda d: container.find_element(By.XPATH, XPATH_COMPANY_IMG))

                # Get the attribute
                raw_name = company_img_element.get_attribute('alt')
                
                # If alt is empty, try title (common fallback in these insurance portals)
                if not raw_name:
                    raw_name = company_img_element.get_attribute('title')

                company_name = convert_company_name(raw_name.strip() if raw_name else "")

                life_data.append({
                    "Product Name": product_name,
                    "Premium Amount": premium_text,
                    "Company Name": company_name
                })

            except Exception as e:
                continue
            
        company_totals: dict[str, float] = {}
        for result in life_data:
            product_name = result['Product Name']
            company_name = result['Company Name']
            premium_amount_str = result['Premium Amount']

            print(f"  Product: {product_name}")
            print(f"  Premium: {premium_amount_str}")
            print(f"  Original Company: {company_name}")
            company_name = convert_company_name(company_name)
            print(f"  Shortened Company: {company_name}")
            try:
                premium_value = float(premium_amount_str)
                
                if "מחלות קשות" in product_name:
                    if "אישי" in product_name:
                        target_product_key = "מחלות קשות אישי"
                    elif "קבוצתי" in product_name:
                        target_product_key = "מחלות קשות קבוצתי"
                elif "תאונות אישיות" in product_name:
                    if "אישי" in product_name:
                        target_product_key = "תאונות אישיות אישי"
                    elif "קבוצתי" in product_name:
                        target_product_key = "תאונות אישיות קבוצתי"
                elif "סיעודי עד" in product_name:
                    target_product_key = product_name
                elif "כתב שירות" not in product_name:
                    if "אישי" in product_name:
                        target_product_key = "בריאות אישי"
                    elif "קבוצתי" in product_name:
                        target_product_key = "בריאות קבוצתי"
                else:
                    target_product_key = "null"
                
                if target_product_key != "null":
                    full_key = f"{company_name}-{target_product_key}"
                    company_totals[full_key] = company_totals.get(full_key, 0) + premium_value # Add for the same type and company to the premium value found (combine them all)
                    
            except ValueError:
                print(f"Could not convert '{premium_amount_str}' to a number. Skipping aggregation for this item.")

        total_keys_health = 0
        current_key_health = 1
        total_keys_self_accidents = 0
        current_key_self_accidents = 1
        total_keys_hard_sickness = 0
        current_key_hard_sickness = 1
        total_keys_nursery = 0
        current_key_nursery = 1
        for title in company_totals:
            if "מחלות קשות" in title:
                total_keys_hard_sickness += 1
            elif "תאונות אישיות" in title:
                total_keys_self_accidents += 1
            elif "סיעודי" in title:
                total_keys_nursery += 1
            else:
                total_keys_health += 1
        
        # Since the task is to categorize each insurance, we do that using the different categories added befor end now add them for a more estetic output
        Nursing_Results = ""
        Hard_Sickness_Results = ""
        Self_Accidents_Results = ""
        for company, total_premium in company_totals.items():
        
            rounded_amount = math.floor(total_premium)
            
            result = (f"{company.split('-')[1]} {rounded_amount} {company.split('-')[0]}")

            if "סיעודי עד" in company:
                Nursing_Results += result
                if current_key_nursery < total_keys_nursery:
                    Nursing_Results += " + "
                    current_key_nursery += 1
            elif "מחלות קשות" in company:
                Hard_Sickness_Results += result
                if current_key_hard_sickness < total_keys_hard_sickness:
                    Hard_Sickness_Results += " + "
                    current_key_hard_sickness += 1
            elif "תאונות אישיות" in company:
                Self_Accidents_Results += result
                if current_key_self_accidents < total_keys_self_accidents:
                    Self_Accidents_Results += " + "
                    current_key_self_accidents += 1
            else:
                Health_Results += result
                if current_key_health < total_keys_health:
                    Health_Results += " + "
                    current_key_health += 1

        results_entries = f"{client_name}\t"
        results_entries += f"{convert_id_back(client_id)}\t"
        results_entries += f"{convert_date_back(id_date)}\t"
        results_entries += f"{convert_date_back(birth_date)}\t"
        results_entries += f"{phone}\t"
        results_entries += f"{insurance_company}\t"

        # More estetic output
        if Life_Results == "" and Health_Results == "" and Hard_Sickness_Results == "" and Self_Accidents_Results == "" and Nursing_Results == "":
            results_entries += f"אין ביטוחים\tאין\tאין\tאין\tאין\t"
        else:
            if Life_Results != "":
                results_entries += f"{Life_Results}\t"
            else:
                results_entries += f"אין\t"
            if Health_Results != "":
                results_entries += f"{Health_Results}\t"
            else:
                results_entries += f"אין\t"
            if Hard_Sickness_Results != "":
                results_entries += f"{Hard_Sickness_Results}\t"
            else:
                results_entries += f"אין\t"
            if Self_Accidents_Results != "":
                results_entries += f"{Self_Accidents_Results}\t"
            else:
                results_entries += f"אין\t"
            if Nursing_Results != "":
                results_entries += f"{Nursing_Results}\t"
            else:
                results_entries += f"אין\t"
        results_entries += reason

        driver.quit()
        return results_entries
    except Exception as e:
        driver.save_screenshot("crash_debug.png")
        print(f"Internal Error: {e}")
        return f"Error: {e}"
    finally:
        if driver:
            driver.quit()

def check_session_validity(driver, timeout=10): # A function to check that the user is logged in as expected
    try:
        WebDriverWait(driver, timeout).until(EC.url_matches("https://harb.cma.gov.il/sso/Auth/Agent|https://harb.cma.gov.il/Home/Index"))
    except:
        return False
    
    current_url = driver.current_url
    
    if "https://harb.cma.gov.il/sso/Auth/Agent" in current_url:
        return True
    elif "https://harb.cma.gov.il/Home/Index?agent=NotAgent" in current_url:
        return True
    else:
        return False

def create_sms_wait_gui(): # The function for the waiting gui for code confirmation via SMS
    global root
    root = Tk()
    root.title("ממתין...")
    root.geometry("350x150")
    root.resizable(False, False)
    root.attributes('-topmost', True)

    L = Label(root, text="ממתין לאימות קוד בצד השרת...")
    L.grid(row=0, column=0, pady=10, padx=10, sticky="ew")

    def continue_button_action():
        SMS_COMPLETE_FLAG.set()
        root.quit()
        root.destroy()

    B1 = Button(root, text ="לחץ לאחר אימות הקוד\nומעבר לעמוד הבית", command = continue_button_action)
    B1.grid(row=1, column=0, pady=10, padx=10, sticky="ew")

    def on_closing():
        cleanup() # Kill the driver
        root.destroy() # Close the window
        os._exit(0) # Force exit the script and all threads

    root.protocol("WM_DELETE_WINDOW", on_closing) # If closing this window -> clean up resources and close app entirely

    root.grid_columnconfigure(0, weight=1)
    root.mainloop()

def login_and_wait(driver, MAIN_PAGE, SING_IN_ID, SIGN_IN_PASSWORD): # The function for logging in automatically using the ID and Password passed to it
    SMS_COMPLETE_FLAG.clear() 
    try:
        user_input = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "userId"))) # Make sure to wait for the page to load entirely
        pass_input = driver.find_element(By.ID, 'userPass')

        user_input.clear()
        user_input.send_keys(SING_IN_ID)
        pass_input.clear()
        pass_input.send_keys(SIGN_IN_PASSWORD)

        driver.find_element(By.ID, 'loginSubmit').click()
        
        print("Login credentials submitted.")

    except Exception as e:
        print(f"Error during login form submission: {e}")
        return False

    gui_thread = threading.Thread(target=create_sms_wait_gui, daemon=True) # Start the gui thread for showing the SMS confirmation
    gui_thread.start()

    print("Waiting for user to complete SMS via GUI...")

    is_set = SMS_COMPLETE_FLAG.wait(timeout=180) # Set a timeout to the SMS event flag to prevent waiting too long and the site to reject the authentication

    if not is_set:
        print("SMS wait timed out. Aborting automation.")
        return False
        
    print("User signaled MFA complete. Resuming main thread.")
    return True

def load_session_cookies(driver): # A function to load site cookies and prevent another log in
    print (os.path.abspath(COOKIE_FILE))
    if os.path.exists(COOKIE_FILE): # Check for the cookie file
        print(f"Loading cookies from {COOKIE_FILE}...")
        with open(COOKIE_FILE, 'rb') as file:
            cookies = pickle.load(file) # Load the cookie using pickle

        driver.get(BASE_DOMAIN)
        
        try:
            for cookie in cookies:
                if 'expiry' in cookie:
                    del cookie['expiry'] # Make sure to remove all cookies that have expiry date, not as usfull since the gov site check expiry on it's side
                driver.add_cookie(cookie)
        except Exception as e:
            print(f"error loading cookies - {e}")
            return False

        driver.refresh()
        return True
    else:
        print("No cookie file is present, exiting...")
        return False

def start_driver(): # A function to start the chrome driver
    driver = get_undetected_driver()
    if os.path.exists(COOKIE_FILE): # Check for the cookie file
        load_session_cookies(driver) # Load Cookies and Try to Reuse Session
        driver.get(MAIN_PAGE)
        
        is_session_valid = check_session_validity(driver, timeout=3)

        if is_session_valid:
            print("SUCCESS: Session reused! Automation can continue.")
            driver.quit()
            return True
            
        else:
            try:
                os.remove(COOKIE_FILE) # Remove the 'not working' cookie file, it's unusable
                print(f"Expired cookie file '{COOKIE_FILE}' deleted.")
            except OSError as e:
                print(f"Error deleting cookie file: {e}")

            login_successful = login_and_wait(driver, MAIN_PAGE, SING_IN_ID, SIGN_IN_PASSWORD) # Log in using the ID and Password given at the top of the script

            if login_successful:
                save_session_cookies(driver)
                driver.quit()
                return True
            else:
                driver.quit()
    else:
        print("No cookie file found. Starting manual login.")
        driver.get(MAIN_PAGE)

        login_successful = login_and_wait(driver, MAIN_PAGE, SING_IN_ID, SIGN_IN_PASSWORD) # Same as above, but in different cercemstances (no cookie file)

        if login_successful:
            save_session_cookies(driver) 
            driver.quit()
            return True
        else:
            driver.quit()
    return False

def convert_id(input): # A function to fill in the extra '0' if not adding the prefix number to the ID number
    output = input
    while len(output) < 9:
        output = "0" + output
    return output

def convert_id_back(id_string): # A function to remove prefix numbers from the ID number -> to the excel
    return id_string.lstrip('0')

def convert_dates(input): # Convert the date to the right syntax used later / in the gov site -> from excel since it's written diffrently
    try:
        day = input.replace('.', '/').split('/')[0]
        if day[0] == '0':
            day = day[1]
        month = input.replace('.', '/').split('/')[1]
        if month[0] == '0':
            month = month[1]
        year = input.replace('.', '/').split('/')[2]
        if len(year) == 2: # If not mentioning which century add '20' or '19' depending on the possible encounter
            if year[0] >= '0' and year[0] <= '2':
                year = "20" + year
            else:
                year = "19" + year
        elif len(year) != 4:
            return "error"
        
        current_year = date.today().year
        if int(current_year) <= int(year): # If the year date is too large for some reason
            return "error"
        return day + "/" + month + "/" + year
        
    except:
        return "error"

def convert_date_back(input): # A function to convert the date back to the excel format
    day = input.replace('/', '.').split('.')[0]
    if day[0] == '0':
        day = day[1]
    month = input.replace('/', '.').split('.')[1]
    if month[0] == '0':
        month = month[1]
    year = input.replace('/', '.').split('.')[2]
    if len(year) == 2:
        if (year[0] < 4):
            year = f"20{year}"
        else:
            year = f"19{year}"
    return day + "." + month + "." + year

def main_gui(): # The function for displaying the GUI box for entering input from the excel
    def extract_data(): # The main function for extracting the data that was given
        raw_data = text_input.get("1.0", tk.END).strip()
        
        if not raw_data:
            messagebox.showwarning("Warning", "אנא הזן נתונים")
            return

        try:
            root.withdraw()
        except Exception as e:
            print(f"error : {e}")
            cleanup()
            
        running_win = tk.Toplevel()
        running_win.attributes('-topmost', True)
        running_win.title("מעבד...")
        running_win.geometry("300x100")
        def on_closing():
            cleanup() # Kill the driver
            root.destroy() # Close the window
            os._exit(0) # Force exit the script and all threads

        running_win.protocol("WM_DELETE_WINDOW", on_closing) # Make sure to do a cleanup of the program if exiting
        tk.Label(running_win, text="מפיק נתונים, נא להמתין...", font=("Arial", 12)).pack(expand=True)

        
            
        
        def worker(): # The function to process data that was entered and extract data
            try:
                parts = raw_data.split('\t') # Since the data is pasted directly from excel, each coloumn is seperated with a \t syntax -> divide them to different parts
                if len(parts) < 7: # There are not enought parts that were pasted
                    root.after(0, lambda: handle_error("בעיה בהעתקה - בדוק אם התפספס עמודות", running_win))
                    return
                client_name = parts[0]
                client_id = convert_id(parts[1])
                id_date = convert_dates(parts[2])
                birth_date = convert_dates(parts[3])
                # In case of an error in converting the date, it should return "error"
                if id_date == "error": 
                    error_is += ("  ")
                    root.after(0, lambda: handle_error("תאריך תפוקה של תעודת הזהות כתוב לא נכון", running_win))
                    return
                if birth_date == "error": 
                    root.after(0, lambda: handle_error("תאריך יום לידה כתוב לא נכון", running_win))
                    return
                phone = parts[4]
                insurance_company = parts[5]
                reason = parts[6]

                # Call the backend function using the extracted input
                output = get_lead(client_name, client_id, birth_date, id_date, phone, insurance_company, reason)
                if (output == "log_in"):
                    print("Not logged in, showing log in page")
                    root.after(0, lambda: running_win.destroy())
                    root.destroy()
                    if start_driver():
                        main_gui() # If 'get_lead' returns "log_in" it means the site is no longer logged in to user and you need to log back in
                else:
                    try:
                        print(output)
                        root.after(0, lambda: finish_extraction(output, running_win)) # Use lambda to pass finish_extraction with parameters without executing it immediately
                    except:
                        finish_extraction(output, None) # In case running_win is not set or existant just pass None

            except Exception as e:
                print(f"Error: {e}")
                root.after(0, handle_error(e, running_win))

        threading.Thread(target=worker, daemon=True).start()

    def finish_extraction(output, running_win): # A function for removing previous gui and calling the results gui
        if (running_win != None):
            running_win.destroy()
        show_result_window(output, root, text_input)

    def handle_error(e, running_win): # A funciont to show an error if happend
        try:
            running_win.destroy()
        except:
            print("No window to destroy")
        root.deiconify()
        messagebox.showerror("Error", f"קרתה שגיאה: {e}")

    root = tk.Tk()
    root.title("הכן ליד")
    root.attributes('-topmost', True)
    root.geometry("600x300")

    root.option_add('*Font', 'Arial 12')

    def on_closing():
        cleanup() # Kill the driver
        root.destroy() # Close the window
        os._exit(0) # Force exit the script and all threads

    root.protocol("WM_DELETE_WINDOW", on_closing) # Make sure to remove resources when closing window

    label = tk.Label(root, text=":הדבק נתונים מהאקסל כאן", pady=10)
    label.pack()

    text_input = scrolledtext.ScrolledText(root, wrap=tk.NONE, width=70, height=5)
    text_input.pack(padx=20, pady=10, expand=True, fill="both")

    extract_button = tk.Button(
        root, 
        text="הוצא נתונים", 
        command=extract_data, 
        bg="#4CAF50", 
        fg="white", 
        font=("Arial", 14, "bold"),
        padx=20,
        pady=10
    )
    extract_button.pack(side=tk.BOTTOM, pady=20)

    # Start the application
    root.mainloop()

def show_result_window(output_text, root, text_input): # A function for GUI showing the results clearly (and copying it to clipboard)
    result_win = tk.Toplevel()
    result_win.title("תוצאה")
    result_win.geometry("600x500")
    result_win.attributes('-topmost', True)

    tk.Label(result_win, text="הנתונים המופקים (ניתן להדביק):", pady=10).pack()

    out_box = scrolledtext.ScrolledText(result_win, width=70, height=15)
    out_box.insert(tk.END, output_text)
    pyperclip.copy(output_text)
    out_box.pack(padx=20, pady=10)

    def on_closing():
        cleanup() # Kill the driver
        root.destroy() # Close the window
        os._exit(0) # Force exit the script and all threads

    result_win.protocol("WM_DELETE_WINDOW", on_closing) # Make sure if the window is closed to remove app's resources

    def restart(): # Function to show the 'main gui' back
        result_win.destroy()
        text_input.delete("1.0", tk.END) 
        root.deiconify() 

    restart_button = tk.Button(
        result_win, 
        text="המרה נוספת", 
        command=restart,
        bg="#2196F3", 
        fg="white",
        font=("Arial", 12, "bold"),
        pady=10
    )
    restart_button.pack(pady=20)

def main():
    print("Starting...")
    root = tk.Tk() # To prevent errors with either of the functions that start (main_gui / create_sms_wait_gui)
    root.withdraw() # Keep it hidden until we decide what to show
    atexit.register(cleanup) # Make a cleanup when exiting
    if start_driver():
        main_gui()

if __name__=="__main__":
    main()

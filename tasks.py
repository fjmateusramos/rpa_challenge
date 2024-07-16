import os
import re
import datetime
import logging
import time
from robocorp import storage
from robocorp.tasks import task, get_output_dir
from RPA.Excel.Files import Files
from RPA.Browser.Selenium import Selenium
from RPA.Robocloud.Items import Items
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from RPA.HTTP import HTTP

# Configure logging to display time, log level, and message
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def setup_date_variables():
    """
    Setup and return date variables to filter results based on dates.
    Returns a dictionary with today's date, last month, and the month before last.
    """
    today = datetime.date.today()
    first_of_month = today.replace(day=1)
    return {
        'today': today.strftime("%B"),
        'last_month': (first_of_month - datetime.timedelta(days=1)).strftime("%B"),
        'before_last_month': (first_of_month - datetime.timedelta(days=31)).strftime("%B"),
    }
    
def handle_cookies(browser, xpath_input):
    """
    Attempts to close cookie notifications if they are visible on the page.
    """
    try:
        if browser.is_element_visible(xpath_input):
            browser.click_element_when_clickable(xpath_input)
            logging.info("Cookie notice dismissed successfully.")
    except Exception as e:
            logging.warning(f"Failed to dismiss cookie notice: {e}")

def download_image(browser, result_index, result_locator_xpath, counter):
    """
    Attempts to download an image from the specified path and save it locally.
    Returns the filename if successful, 'No Picture Found' otherwise.
    """
    try:
        src_path = browser.get_element_attribute(
            result_locator_xpath + f'[{result_index}]//div/div/a/picture/source', 'srcset'
        )
        src_path1 = src_path.split(" ")[0]
        image_name = f"{counter}_challenge.png"
        file_path = os.path.join(get_output_dir(), image_name)
        HTTP().download(src_path1, file_path, overwrite=True)
        return image_name
    except Exception as e:
        logging.warning(f"Image download failed: {e}")
        return 'No Picture Found'

def count_sequence(string, seq):
    """
    Returns the count of subsequences `seq` found in string `string`.
    """
    return string.count(seq)

def should_include_result(date_text, range_news_input, date_vars):
    """
    Determines if a search result should be included based on its date.
    Returns True if the result matches the criteria, False otherwise.
    """
    conditions = {
        0: ['Yesterday', 'ago', date_vars['today']],
        1: ['Yesterday', 'ago', date_vars['today']],
        2: ['Yesterday', 'ago', date_vars['today'], date_vars['last_month']],
        3: ['Yesterday', 'ago', date_vars['today'], date_vars['last_month'], date_vars['before_last_month']]
    }
    for condition in conditions[range_news_input]:
        if count_sequence(date_text, condition) > 0:
            return True
    return False

def process_results(browser, search_results, result_locator_xpath, date_vars, config_data, money_pattern):
    """
    Processes each search result, extracting and storing details if they meet specified criteria.
    Returns a dictionary containing details of all processed results.
    """
    results = {
        'title': [], 'date': [], 'description': [], 'picture': [], 'count_search_phrases': [], 'description_contains_money': []
    }
    counter = 0
    for result_index, element in enumerate(search_results, 1):
        try:
            title_text = browser.get_text(result_locator_xpath + f'[{result_index}'"]//div[@class='PagePromo-title']/a/span")
            description_text = browser.get_text(result_locator_xpath + f'[{result_index}'"]//div[@class='PagePromo-description']/a/span")
            date_text = browser.get_text(result_locator_xpath + f'[{result_index}'"]//div[@class='PagePromo-date']//span")

            if should_include_result(date_text, config_data['range_news'], date_vars):
                counter += 1
                image_name = download_image(browser, result_index, result_locator_xpath, counter)
                results['title'].append(title_text)
                results['description'].append(description_text)
                results['date'].append(date_text)
                results['picture'].append(image_name)
                results['count_search_phrases'].append(count_sequence(title_text, config_data['search']) + count_sequence(description_text, config_data['search']))
                results['description_contains_money'].append(
                    bool(re.fullmatch(money_pattern, title_text)) or bool(re.fullmatch(money_pattern, description_text))
                )
            logging.info(f"Processed result {result_index}/{len(search_results)}")
        except Exception as e:
            logging.error(f"Error processing result {result_index}: {e}")
    return results

@task
def robocorp_challenge() -> None:
    """
    Main task to orchestrate the web scraping challenge.
    Opens the website, handles cookies, performs search, and processes results.
    """
    logging.info("Starting robocorp_challenge task.")
    
    # Set the storage variables and local variables
    config_data_storage = storage.get_json('Config_Data_Challenge')
    items = Items()
    work_item = items.get_input_work_item()  # Fetch the initial work item
    date_vars = setup_date_variables()
    search_item_input = config_data_storage['search']
    range_news_input = config_data_storage['range_news']
    web_site_url_input = config_data_storage['web_site']
    date_vars = setup_date_variables()
    money_pattern = re.compile(
        r"^\$(\d{1,3}(,\d{3})*\.\d{2}|\d{1,2}\.\d{1})$|^\d{1,3}(,\d{3})*\s+dollars$|^\d{1,3}(,\d{3})*\s+USD$"
    )

    # Set xpath's needed to query the elements
    search_button_xpath = "//div//bsp-search-overlay"
    input_text_field_xpath = "//input[@type='text' and @name='q']"
    category_button_xpath = "//bsp-toggler[@data-toggle-in='search-filter']"
    section_check_input_xpath = "//input[@value='00000189-9323-dce2-ad8f-bbe74c770000']"
    sortby_combo_box_xpath = "//select[@class='Select-input']"
    result_locator_xpath = "//div[@class='SearchResultsModule-results']//div[@class='PageList-items-item']"
    reject_cookies_xpath = "//a[@title='Close']"
    
    # Perform taks to open the browser
    browser = Selenium()
    browser.open_chrome_browser(web_site_url_input, headless=True)
    handle_cookies(browser, reject_cookies_xpath)  # Handle cookies right after opening the browser
    
    # Perform tasks to search on the webpage
    browser.wait_until_element_is_visible(search_button_xpath)
    try:
        browser.click_element_when_clickable(search_button_xpath)
    except Exception as e:
        logging.warning(f"Search button click failed: {e}")
        handle_cookies(browser, reject_cookies_xpath)  # Attempt to handle cookies if the first click fails
    browser.wait_until_element_is_visible(input_text_field_xpath)
    browser.input_text(input_text_field_xpath, search_item_input + Keys.ENTER)   # Search in the web page

    # Perform tasks to click on the stories category and sort results    
    browser.wait_until_element_is_visible(category_button_xpath)
    browser.click_element_when_clickable(category_button_xpath)
    try:
        browser.wait_until_element_is_visible(category_button_xpath)
        browser.click_element_when_clickable(section_check_input_xpath)
    except Exception as e:
        logging.warning(f"Stories button check visible failed: {e}")          
    
    browser.select_from_list_by_value(sortby_combo_box_xpath, '3')
    browser.set_browser_implicit_wait(10)    
    handle_cookies(browser, reject_cookies_xpath) # Handle cookies right after search and click in category

    # Perform tasks to find the results 
    browser.wait_until_element_is_visible(result_locator_xpath) # Wait for the elements are visible
    search_results = browser.find_elements(result_locator_xpath) # Get the found elements
    results = process_results(browser, search_results, result_locator_xpath, date_vars, config_data_storage, money_pattern) # Process the elements
    work_item.set_variable("results", results)  # Update work item with processed results
    items.save_work_item(work_item)
    # Perform tasks to build the excel file
    logging.info("Saving results to Excel.")
    try:
        excel = Files()
        wb = excel.create_workbook()
        wb.create_worksheet('Results')
        excel.append_rows_to_worksheet(results, header=True, name='Results')
        wb.save(os.path.join(get_output_dir(), 'search_results.xlsx'))
        logging.info("Results saved successfully.")
    except Exception as e:
        logging.error(f"Error saving results to Excel: {e}")

    # Perform tasks to close the browser
    browser.close_browser()
    logging.info("Browser closed and task completed.")

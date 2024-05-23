import os
import re
import datetime
import logging
import time
from robocorp import storage
from robocorp.tasks import task, get_output_dir
from RPA.Excel.Files import Files
from RPA.Browser.Selenium import Selenium
from selenium.webdriver.common.keys import Keys
from RPA.HTTP import HTTP

# Configuración de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

@task
def robocorp_challenge() -> None:
    logging.info("Starting robocorp_challenge task.")
    config_data = storage.get_json('Config_Data_Challenge')
    search_item = config_data['search']
    range_news = config_data['range_new']

    def count_sequence(string, seq):
        return string.count(seq)

    browser = Selenium(auto_close=False)
    browser.open_chrome_browser(config_data['web_site'], headless=True)
    search_button = "//button[@class='SearchOverlay-search-button']"
    input_text_field = "//input[@class='SearchOverlay-search-input']"
    category_button = "//div[@class='SearchResultsModule-filters-content']"
    section_check_input = "//input[@value='00000188-f942-d221-a78c-f9570e360000']"
    sortby_combo_box = "//select[@class='Select-input']"
    onetrust_div = "//div[@class='onetrust-pc-dark-filter ot-fade-in']"

    browser.wait_until_element_is_visible(search_button)
    time.sleep(5)
    try:
        browser.click_element(search_button)
    except Exception as e:
        logging.warning(f"Search button click failed: {e}")
        try:
            browser.wait_until_element_is_visible(onetrust_div)
            browser.click_element(onetrust_div)
            browser.delete_all_cookies
            reject_cookies = "//a[@title='Close']"
            browser.wait_until_element_is_visible(reject_cookies)
            browser.click_element(reject_cookies)
            browser.click_element(search_button)
        except Exception as e:
            logging.warning(f"Reject cookies failed: {e}")

    browser.wait_until_element_is_visible(input_text_field)
    browser.input_text(input_text_field, search_item + Keys.ENTER)

    browser.wait_until_element_is_visible(category_button)
    browser.click_element(category_button)
    browser.wait_until_element_is_visible(section_check_input)
    browser.click_element(section_check_input)

    browser.wait_until_element_is_visible(sortby_combo_box)
    browser.set_browser_implicit_wait(10)
    browser.select_from_list_by_value(sortby_combo_box, '3')
    browser.set_browser_implicit_wait(10)

    try:
        reject_cookies = "//a[@title='Close']"
        browser.wait_until_element_is_visible(reject_cookies)
        browser.click_element(reject_cookies)
        browser.click_element(search_button)
    except Exception as e:
        logging.warning(f"Reject cookies popup handling failed: {e}")

    result_locator = "//div[@class='SearchResultsModule-results']/bsp-list-loadmore/div/div"
    browser.wait_until_element_is_visible(result_locator)

    search_results = browser.find_elements(result_locator)
    results = {
        'title': [],
        'date': [],
        'description': [],
        'picture': [],
        'count_search_phrases': [],
        'description_contains_money': []
    }
    counter = 0
    money_pattern = re.compile(
        r"^\$(\d{1,3}(,\d{3})*\.\d{2}|\d{1,2}\.\d{1})$|^\d{1,3}(,\d{3})*\s+dollars$|^\d{1,3}(,\d{3})*\s+USD$"
    )

    for result_index, element in enumerate(search_results, 1):
        counter += 1
        try:
            title_text = browser.get_text(result_locator + f'[{result_index}]//div/div/bsp-custom-headline/div/a/span')
            description_text = browser.get_text(result_locator + f'[{result_index}]//div/div/div/a/span')
            date_text = browser.get_text(result_locator + f'[{result_index}]//div/div/div/div/bsp-timestamp/span/span')

            today = datetime.date.today()
            first = today.replace(day=1)
            this_month = today.strftime("%B")
            last_month = (first - datetime.timedelta(days=1)).strftime("%B")
            before_last_month = (first - datetime.timedelta(days=31)).strftime("%B")

            if (range_news == 0 or range_news == 1) and (
                date_text == "Yesterday" or count_sequence(date_text, 'ago') > 0 or count_sequence(date_text, this_month) > 0
            ):
                results['title'].append(title_text)
                results['description'].append(description_text)
                results['date'].append(date_text)
                results['count_search_phrases'].append(
                    count_sequence(title_text, config_data['search']) + count_sequence(description_text, config_data['search'])
                )
                try:
                    src_path = browser.get_element_attribute(
                        result_locator + f'[{result_index}]//div/div/a/picture/source', 'srcset'
                    )
                    src_path1 = src_path.split(" ")[0]
                    image_name = f"{counter}_challenge.png"
                    file_path = os.path.join(get_output_dir(), image_name)
                    HTTP().download(src_path1, file_path, overwrite=True)
                    results['picture'].append(image_name)
                except Exception as e:
                    logging.warning(f"Image download failed: {e}")
                    results['picture'].append('No Picture Found')
                results['description_contains_money'].append(
                    bool(re.fullmatch(money_pattern, title_text)) or bool(re.fullmatch(money_pattern, description_text))
                )

            elif (range_news == 2) and (
                date_text == "Yesterday" or count_sequence(date_text, 'ago') > 0 or count_sequence(date_text, this_month) > 0 or count_sequence(date_text, last_month) > 0
            ):
                results['title'].append(title_text)
                results['description'].append(description_text)
                results['date'].append(date_text)
                results['count_search_phrases'].append(
                    count_sequence(title_text, config_data['search']) + count_sequence(description_text, config_data['search'])
                )
                try:
                    src_path = browser.get_element_attribute(
                        result_locator + f'[{result_index}]//div/div/a/picture/source', 'srcset'
                    )
                    src_path1 = src_path.split(" ")[0]
                    image_name = f"{counter}_challenge.png"
                    file_path = os.path.join(get_output_dir(), image_name)
                    HTTP().download(src_path1, file_path, overwrite=True)
                    results['picture'].append(image_name)
                except Exception as e:
                    logging.warning(f"Image download failed: {e}")
                    results['picture'].append('No Picture Found')
                results['description_contains_money'].append(
                    bool(re.fullmatch(money_pattern, title_text)) or bool(re.fullmatch(money_pattern, description_text))
                )

            elif (range_news == 3) and (
                date_text == "Yesterday" or count_sequence(date_text, 'ago') > 0 or count_sequence(date_text, this_month) > 0 or count_sequence(date_text, last_month) > 0 or count_sequence(date_text, before_last_month) > 0
            ):
                results['title'].append(title_text)
                results['description'].append(description_text)
                results['date'].append(date_text)
                results['count_search_phrases'].append(
                    count_sequence(title_text, config_data['search']) + count_sequence(description_text, config_data['search'])
                )
                try:
                    src_path = browser.get_element_attribute(
                        result_locator + f'[{result_index}]//div/div/a/picture/source', 'srcset'
                    )
                    src_path1 = src_path.split(" ")[0]
                    image_name = f"{counter}_challenge.png"
                    file_path = os.path.join(get_output_dir(), image_name)
                    HTTP().download(src_path1, file_path, overwrite=True)
                    results['picture'].append(image_name)
                except Exception as e:
                    logging.warning(f"Image download failed: {e}")
                    results['picture'].append('No Picture Found')
                results['description_contains_money'].append(
                    bool(re.fullmatch(money_pattern, title_text)) or bool(re.fullmatch(money_pattern, description_text))
                )

            logging.info(f"Processed result {result_index}/{len(search_results)}")

        except Exception as e:
            logging.error(f"Error processing result {result_index}: {e}")

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

    browser.close_browser()
    logging.info("Browser closed and task completed.")

from time import sleep
from parsel import Selector
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains
import os
from datetime import datetime
import json
import us
from pprint import pprint
import re


# helper function for getting values from selector object
def parse(response, xpath, get_method="get", comma_join=False, space_join=True):
    """_This function is used to get values from selector object by using xpath expressions_

    Args:
        response (_scrapy.Selector_): _A selector object on which we can use xpath expressions_
        xpath_str (_str_): _xpath expression to be used_
        get_method (str, optional): _whether to get first element or all elements_. Defaults to "get".
        comma_join (bool, optional): _if we are getting all elements whether to join on comma or not_. Defaults to False.
        space_join (bool, optional): _if we are getting all elements whether to join on space or not_. Defaults to False.

    Returns:
        _str_: _resultant value of using xpath expression on the scrapy.Selector object_
    """
    value = ""
    if get_method == "get":
        value = response.xpath(xpath).get()
        value = (value or "").strip()
    elif get_method == "getall":
        value = response.xpath(xpath).getall()
        if value:
            if comma_join:
                value = " ".join(
                    ", ".join([str(x).strip() for x in value]).split()
                ).strip()
                value = (value or "").strip()
            elif space_join:
                value = " ".join(
                    " ".join([str(x).strip() for x in value]).split()
                ).strip()
                value = (value or "").strip()
        else:
            value = ""
    return value


# this function is used to setup the bot
def bot_setup(headless=False):
    """_This function is used to setup the bot_

    Args:
        proxy_switch (_int_): _whether to use proxy or not, 0 means yes, and 1 means no_
        headless (bool, optional): _whether to run the bot in headless mode or not_. Defaults to False.

    Returns:
        _selenium.webdriver_: _returns a selenium.webdriver object to be used_
    """

    # options to be used
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    # if headless==True, make the bot headless
    if headless:
        options.add_argument("--headless=new")

    driver = webdriver.Chrome(
        service=Service(),
        options=options,
    )
    # setup implicit wait
    driver.implicitly_wait(3)
    driver.maximize_window()
    return driver


def send_keys(driver, xpath, keys, wait_time=5):
    element = WebDriverWait(driver, wait_time).until(
        EC.presence_of_element_located((By.XPATH, xpath))
    )
    element.send_keys(keys)


def click_btn(driver, xpath, wait_time=5):
    btn = WebDriverWait(driver, wait_time).until(
        EC.presence_of_element_located((By.XPATH, xpath))
    )
    ActionChains(driver).move_to_element(btn).perform()
    sleep(1)
    ActionChains(driver).click(btn).perform()


def wait_for_element(driver, xpath, wait_time=5):
    WebDriverWait(driver, wait_time).until(
        EC.presence_of_element_located((By.XPATH, xpath))
    )


def get_state_full_name(abbreviation):
    try:
        state = us.states.lookup(abbreviation)
        return state.name
    except:
        return False


def get_max_value_plus_one(dictionary):
    values = dictionary.values()
    if not values:
        return 1

    max_value = max(values)
    return max_value + 1


def save_ids_file(major_ids, program_ids, filepath):
    json_ = {
        "major_ids": major_ids,
        "program_ids": program_ids,
    }
    with open(filepath, "w") as f:
        json.dump(json_, f, indent=4)


def clean_string(input_string):
    # Define special characters
    special_chars = "!@#$%^&*()_-+={}[]|\:;'<>,.?/~`"
    
    # Remove special characters
    cleaned_string = ''.join(char for char in input_string if char.isalnum() or char not in special_chars)
    
    # Remove spaces
    cleaned_string = cleaned_string.replace(' ', '')
    
    # Convert to lowercase
    cleaned_string = cleaned_string.lower()
    
    return cleaned_string


current_datetime = datetime.now()
formatted_datetime = current_datetime.strftime("%d-%m-%Y_%I-%M-%p")


ids_file_path = "ids.json"
input_csv_path = "input.csv"
output_file_path = f"output_{formatted_datetime}.xlsx"

df = pd.read_csv(input_csv_path)
inp_records = df.to_dict(orient="records")

if os.path.exists(ids_file_path):
    with open(ids_file_path, "r") as f:
        ids = json.load(f)
else:
    ids = {}

majors_ids = ids.get("major_ids", {})
program_ids = ids.get("program_ids", {})

driver = bot_setup()

school_data = []
program_data = []

for inp_idx, inp_rec in enumerate(inp_records): 
    for _ in range(3):
        try:
            print("------------------------------------------------")
            print("Processing -> {}/{}".format(inp_idx + 1, len(inp_records)))
            pprint(inp_rec, sort_dicts=False)

            institute_name = inp_rec["INST_NAME"]
            institute_name_from_input = clean_string(institute_name)

            city = inp_rec["CITY"]
            city_from_input = clean_string(city)

            state = inp_rec["STATE"]

            complete_state_name = get_state_full_name(state.strip())
            if not complete_state_name:
                if state.strip() == "DC":
                    complete_state_name = "District of Columbia"

            complete_state_name_from_input = clean_string(complete_state_name)

            driver.get(
                "https://nces.ed.gov/collegenavigator/?s=IL&pg=3&id=144005#enrolmt"
            )

            wait_for_element(driver, xpath='//input[@value="Type name of school here"]')
            sleep(1)
            send_keys(
                driver,
                xpath='//input[@value="Type name of school here"]',
                keys=Keys.CONTROL + "a",
            )
            send_keys(
                driver,
                xpath='//input[@value="Type name of school here"]',
                keys=Keys.DELETE,
            )
            sleep(0.3)
            send_keys(
                driver,
                xpath='//input[@value="Type name of school here"]',
                keys=institute_name,
            )
            sleep(1)
            for ch in complete_state_name:
                send_keys(
                    driver,
                    xpath='//select[@id="ctl00_cphCollegeNavBody_ucSearchMain_ucMapMain_lstState"]',
                    keys=ch,
                )
                sleep(0.05)
            sleep(1)
            send_keys(
                driver,
                xpath='//select[@id="ctl00_cphCollegeNavBody_ucSearchMain_ucMapMain_lstState"]',
                keys=Keys.ENTER,
            )
            sleep(2)
            try:
                wait_for_element(driver, xpath='//table[@class="resultsTable"]')
                response = Selector(text=driver.page_source)
                no_results = parse(response, xpath='//div[@class="noresults"]')
                if no_results:
                    print("Results Table not visible...")
                    break
            except:
                print("Results Table not visible...")
                break

            sleep(1)
            university_url_found = False
            while True:
                response = Selector(text=driver.page_source)

                university_rows = response.xpath(
                    '//table[@class="resultsTable"]/tbody/tr'
                )

                for uni_row in university_rows:
                    university_name_from_website = parse(
                        uni_row, "./td[2]/a/strong/text()"
                    )
                    university_name_from_website = clean_string(
                        university_name_from_website
                    )

                    university_state_from_website_city = parse(
                        uni_row, "./td[2]/text()", get_method="getall"
                    )

                    university_state_from_website = (
                        university_state_from_website_city.split(",")[-1].strip()
                    )
                    university_state_from_website = clean_string(
                        university_state_from_website
                    )
                    university_city_from_website = (
                        university_state_from_website_city.split(",")[0].strip()
                    )
                    university_city_from_website = clean_string(
                        university_city_from_website
                    )
                    
                    university_url = parse(uni_row, "./td[2]/a/@href")

                    if (
                        institute_name_from_input in university_name_from_website
                        and complete_state_name_from_input
                        == university_state_from_website
                        and city_from_input == university_city_from_website
                    ):
                        university_url_found = True
                        break

                if university_url_found:
                    break

                is_next_page = parse(response, xpath='//a[text()="Next Page »"]')
                if is_next_page:
                    click_btn(driver, xpath='//a[text()="Next Page »"]')
                    try:
                        wait_for_element(driver, xpath='//table[@class="resultsTable"]')
                    except:
                        break
                else:
                    break

            if not university_url_found:
                print("University URL not found...")
                break

            university_url = "https://nces.ed.gov/collegenavigator/" + university_url
            driver.get(university_url)
            sleep(2)
            try:
                wait_for_element(driver, xpath='//span[@class="headerlg"]')
            except:
                break
            sleep(1)
            response = Selector(text=driver.page_source)
            school_items = {}
            school_items["OPEID"] = (
                parse(
                    response,
                    xpath='//span[@class="ipeds"]/text()[contains(., "OPE ID")]',
                )
                .split(":")[-1]
                .strip()
            )
            school_items["School_Name"] = institute_name
            school_items["City"] = city
            school_items["State"] = state
            school_items["Program_IDs"] = ""
            school_items["Total_Enrollment"] = parse(
                response,
                xpath='//th[@scope="col" and text()="Total enrollment"]/following-sibling::th[@scope="col"]/text()',
            )
            school_items["Student Population"] = parse(
                response,
                xpath='//td[@class="srb" and contains(text(), "Student population")]/following-sibling::td[1]/text()',
            )
            school_items["Criminal_Offenses_a"] = parse(
                response,
                xpath='//div[@id="crime"]//div[@class="tablenames" and text()="On-Campus"]/following-sibling::table[1]/tbody/tr[@class="subrow nb" and contains(., "Criminal Offenses")]/following-sibling::tr[1]/td[last()]/text()',
            )
            school_items["Criminal_Offenses_b"] = parse(
                response,
                xpath='//div[@id="crime"]//div[@class="tablenames" and text()="On-Campus"]/following-sibling::table[1]/tbody/tr[@class="subrow nb" and contains(., "Criminal Offenses")]/following-sibling::tr[2]/td[last()]/text()',
            )
            school_items["Criminal_Offenses_c"] = parse(
                response,
                xpath='//div[@id="crime"]//div[@class="tablenames" and text()="On-Campus"]/following-sibling::table[1]/tbody/tr[@class="subrow nb" and contains(., "Criminal Offenses")]/following-sibling::tr[3]/td[last()]/text()',
            )
            school_items["Criminal_Offenses_d"] = parse(
                response,
                xpath='//div[@id="crime"]//div[@class="tablenames" and text()="On-Campus"]/following-sibling::table[1]/tbody/tr[@class="subrow nb" and contains(., "Criminal Offenses")]/following-sibling::tr[4]/td[last()]/text()',
            )
            school_items["Criminal_Offenses_e"] = parse(
                response,
                xpath='//div[@id="crime"]//div[@class="tablenames" and text()="On-Campus"]/following-sibling::table[1]/tbody/tr[@class="subrow nb" and contains(., "Criminal Offenses")]/following-sibling::tr[5]/td[last()]/text()',
            )
            school_items["Criminal_Offenses_f"] = parse(
                response,
                xpath='//div[@id="crime"]//div[@class="tablenames" and text()="On-Campus"]/following-sibling::table[1]/tbody/tr[@class="subrow nb" and contains(., "Criminal Offenses")]/following-sibling::tr[6]/td[last()]/text()',
            )
            school_items["Criminal_Offenses_g"] = parse(
                response,
                xpath='//div[@id="crime"]//div[@class="tablenames" and text()="On-Campus"]/following-sibling::table[1]/tbody/tr[@class="subrow nb" and contains(., "Criminal Offenses")]/following-sibling::tr[7]/td[last()]/text()',
            )
            school_items["Criminal_Offenses_h"] = parse(
                response,
                xpath='//div[@id="crime"]//div[@class="tablenames" and text()="On-Campus"]/following-sibling::table[1]/tbody/tr[@class="subrow nb" and contains(., "Criminal Offenses")]/following-sibling::tr[8]/td[last()]/text()',
            )
            school_items["Criminal_Offenses_i"] = parse(
                response,
                xpath='//div[@id="crime"]//div[@class="tablenames" and text()="On-Campus"]/following-sibling::table[1]/tbody/tr[@class="subrow nb" and contains(., "Criminal Offenses")]/following-sibling::tr[9]/td[last()]/text()',
            )
            school_items["Criminal_Offenses_j"] = parse(
                response,
                xpath='//div[@id="crime"]//div[@class="tablenames" and text()="On-Campus"]/following-sibling::table[1]/tbody/tr[@class="subrow nb" and contains(., "Criminal Offenses")]/following-sibling::tr[10]/td[last()]/text()',
            )
            school_items["Criminal_Offenses_k"] = parse(
                response,
                xpath='//div[@id="crime"]//div[@class="tablenames" and text()="On-Campus"]/following-sibling::table[1]/tbody/tr[@class="subrow nb" and contains(., "Criminal Offenses")]/following-sibling::tr[11]/td[last()]/text()',
            )

            school_items["VAWA_Offenses_a"] = parse(
                response,
                xpath='//div[@id="crime"]//div[@class="tablenames" and text()="On-Campus"]/following-sibling::table[1]/tbody/tr[@class="subrow nb" and contains(., "VAWA Offenses")]/following-sibling::tr[1]/td[last()]/text()',
            )
            school_items["VAWA_Offenses_b"] = parse(
                response,
                xpath='//div[@id="crime"]//div[@class="tablenames" and text()="On-Campus"]/following-sibling::table[1]/tbody/tr[@class="subrow nb" and contains(., "VAWA Offenses")]/following-sibling::tr[2]/td[last()]/text()',
            )
            school_items["VAWA_Offenses_c"] = parse(
                response,
                xpath='//div[@id="crime"]//div[@class="tablenames" and text()="On-Campus"]/following-sibling::table[1]/tbody/tr[@class="subrow nb" and contains(., "VAWA Offenses")]/following-sibling::tr[3]/td[last()]/text()',
            )

            school_items["Arrests_a"] = parse(
                response,
                xpath='//div[@id="crime"]//div[@class="tablenames" and text()="On-Campus"]/following-sibling::table[1]/tbody/tr[@class="subrow nb" and contains(., "Arrests")]/following-sibling::tr[1]/td[last()]/text()',
            )
            school_items["Arrests_b"] = parse(
                response,
                xpath='//div[@id="crime"]//div[@class="tablenames" and text()="On-Campus"]/following-sibling::table[1]/tbody/tr[@class="subrow nb" and contains(., "Arrests")]/following-sibling::tr[2]/td[last()]/text()',
            )
            school_items["Arrests_c"] = parse(
                response,
                xpath='//div[@id="crime"]//div[@class="tablenames" and text()="On-Campus"]/following-sibling::table[1]/tbody/tr[@class="subrow nb" and contains(., "Arrests")]/following-sibling::tr[3]/td[last()]/text()',
            )

            school_items["Disciplinary_Actions_a"] = parse(
                response,
                xpath='//div[@id="crime"]//div[@class="tablenames" and text()="On-Campus"]/following-sibling::table[1]/tbody/tr[@class="subrow nb" and contains(., "Disciplinary Actions")]/following-sibling::tr[1]/td[last()]/text()',
            )
            school_items["Disciplinary_Actions_b"] = parse(
                response,
                xpath='//div[@id="crime"]//div[@class="tablenames" and text()="On-Campus"]/following-sibling::table[1]/tbody/tr[@class="subrow nb" and contains(., "Disciplinary Actions")]/following-sibling::tr[2]/td[last()]/text()',
            )
            school_items["Disciplinary_Actions_c"] = parse(
                response,
                xpath='//div[@id="crime"]//div[@class="tablenames" and text()="On-Campus"]/following-sibling::table[1]/tbody/tr[@class="subrow nb" and contains(., "Disciplinary Actions")]/following-sibling::tr[3]/td[last()]/text()',
            )

            school_data.append(school_items)

            program_rows = response.xpath(
                '(//div[@id="programs"]//table[@class="pmtabular"]/tbody/tr[@class="subrow nb"])[1]/following-sibling::tr[@class="level1indent"]'
            )
            majors = parse(
                response,
                xpath='//div[@id="programs"]//table[@class="pmtabular"]/tbody/tr[@class="subrow nb"]/td/text()',
                get_method="getall",
                space_join=False,
            )
            majors = [str(x).strip() for x in majors]

            school_program_ids = []

            for program_row in program_rows:
                programs = {}

                major_name = parse(
                    program_row,
                    './preceding-sibling::tr[@class="subrow nb"][1]/td/text()',
                )
                if major_name in majors_ids:
                    major_idx = majors_ids[major_name]
                else:
                    major_idx = get_max_value_plus_one(majors_ids)
                    majors_ids[major_name] = major_idx

                    save_ids_file(majors_ids, program_ids, ids_file_path)

                program_name = parse(program_row, "./td[1]/text()")

                if program_name in program_ids:
                    program_idx = program_ids[program_name]
                else:
                    program_idx = get_max_value_plus_one(program_ids)
                    program_ids[program_name] = program_idx

                    save_ids_file(majors_ids, program_ids, ids_file_path)

                school_program_ids.append(program_idx)

                programs["OPEID"] = school_items["OPEID"]
                programs["Major_IDs"] = major_idx
                programs["Program_IDs"] = program_idx
                programs["Major"] = major_name
                programs["Program"] = program_name

                program_data.append(programs)

            school_items["Program_IDs"] = ";".join([str(x) for x in school_program_ids])

            school_data_df = pd.DataFrame(school_data)
            program_data_df = pd.DataFrame(program_data)

            with pd.ExcelWriter(output_file_path, engine="xlsxwriter") as writer:
                # Write each DataFrame to a separate sheet
                school_data_df.to_excel(writer, sheet_name="School", index=False)
                program_data_df.to_excel(writer, sheet_name="Program", index=False)

            break
        except:
            continue


driver.close()
driver.quit()

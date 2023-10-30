from time import sleep
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By
import openpyxl
import os


class WebDriverFactory:
    '''
        class WebDriverFactory is responsible for creating a WebDriver instance with the desired options
    '''
    @staticmethod
    def create_driver(download_directory: str) -> webdriver.Chrome:
        '''
        creating driver
        :param download_directory: directory for file downloading
        :return: webdriver.Chrome
        '''
        chrome_options = webdriver.ChromeOptions()
        prefs = {'download.default_directory': download_directory}
        chrome_options.add_experimental_option('prefs', prefs)
        return webdriver.Chrome(options=chrome_options)


class WebAutomation:
    '''
        class WebAutomation focuses on web interactions like opening a website and downloading a file
    '''
    def __init__(self, driver: webdriver.Chrome):
        '''
        initialization of driver
        :param driver: name of driver
        '''
        self.driver = driver

    def open_website(self, url: str) -> None:
        '''
        opening web-site
        :param url: address of web-site
        '''
        self.driver.get(url)
        self.driver.maximize_window()

    def download_file(self, download_xpath: str) -> None:
        '''
        downloading of the file
        :param download_xpath: xpath to find button for downloading file
        '''
        try:
            file = self.driver.find_element(By.XPATH, download_xpath)
            file.click()
            sleep(2)
        except NoSuchElementException as e:
            print(f"Error downloading file: {str(e)}")


class FormFiller:
    '''
       class handles filling out form fields and submitting forms
    '''
    def __init__(self, driver: webdriver.Chrome):
        self.driver = driver

    def fill_field(self, field_name: str, value: str) -> None:
        '''
        fill field with value
        :param field_name: name of the field
        :param value: value for filling
        '''
        field_xpath = f"//input[@ng-reflect-name='{field_name}']"
        try:
            field = self.driver.find_element(By.XPATH, field_xpath)
            field.send_keys(value)
        except NoSuchElementException as e:
            print(f"Error filling field '{field_name}': {str(e)}")

    def submit_form(self) -> None:
        '''
        click the submit button
        '''
        try:
            submit_button = self.driver.find_element(By.XPATH, "//input[@type='submit']")
            submit_button.click()
        except NoSuchElementException as e:
            print(f"Error submitting the form: {str(e)}")

class FormFieldMapper:
    '''
    class to dynamically map Excel column names to the form fields
    '''
    @staticmethod
    def map_field(column_name: str) -> str:
        if column_name == "First Name":
            return "labelFirstName"
        elif column_name == "Last Name ":
            return "labelLastName"
        elif column_name == "Email":
            return "labelEmail"
        elif column_name == "Company Name":
            return "labelCompanyName"
        elif column_name == "Address":
            return "labelAddress"
        elif column_name == "Phone Number":
            return "labelPhone"
        elif column_name == "Role in Company":
            return "labelRole"
        # Add more mappings as needed
        else:
            return "unknown_field"

class ExcelProcessor:
    '''
       class deals with processing the downloaded Excel file
    '''
    @staticmethod
    def process_excel(download_directory: str, form_filler: FormFiller) -> None:
        '''
        working with excel-file
        :param download_directory: path to the file
        :param form_filler: class for filling of the form
        '''
        file = os.listdir(download_directory)[0]

        if file.endswith('.xlsx'):
            path = os.path.join(download_directory, file)
            workbook = openpyxl.load_workbook(path)
            sheet = workbook.active
            rows = list(sheet.iter_rows(values_only=True))
            rows_count = len([row for row in sheet if not all([cell.value is None for cell in row])])

            if not rows:
                print("No data in the Excel file.")
                return

            header_row = rows[0][:-1]
            for row in rows[1:rows_count]:
                for column_name, value in zip(header_row, row):
                    print(column_name, value)
                    # Map the column name to the corresponding form field dynamically
                    field_name = FormFieldMapper.map_field(column_name)
                    print(field_name)
                    if field_name == "unknown_field":
                        print(f"There is not such field {column_name} in FormFieldMapper")
                    else:
                        form_filler.fill_field(field_name, value)

                form_filler.submit_form()
            sleep(3)

def main():
    download_directory = r'C:\download'
    download_xpath = '//a[contains(text(),"Download")]'
    start_button_xpath = '/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button'
    url = "https://rpachallenge.com/"

    driver = WebDriverFactory.create_driver(download_directory)
    web_automation = WebAutomation(driver)
    form_filler = FormFiller(driver)

    web_automation.open_website(url)
    web_automation.download_file(download_xpath)
    submit_start_button = driver.find_element(By.XPATH, start_button_xpath)
    submit_start_button.click()
    ExcelProcessor.process_excel(download_directory, form_filler)
    driver.close()


if __name__ == "__main__":
    main()



from time import sleep
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By
import openpyxl
import os


class RpaChallengeBot:
    """
       A class for automating the RPA Challenge bot.
    """
    DOWNLOAD_DIRECTORY: str = r'C:\download'
    DOWNLOAD_XPATH: str = '//a[contains(text(),"Download")]'
    START_BUTTON_XPATH: str = '/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button'

    def __init__(self, directory: str =DOWNLOAD_DIRECTORY) -> None:
        """
               Initializes the RpaChallengeBot with the specified directory for downloading source excel-file.
               Args:
                   directory (str): The directory where downloaded files will be saved.
        """
        self.download_directory: str = directory
        self.chrome_options = webdriver.ChromeOptions()
        self.prefs: dict = {
            'download.default_directory': directory
        }
        self.chrome_options.add_experimental_option('prefs', self.prefs)
        self.driver = webdriver.Chrome(options=self.chrome_options)

    def open_website(self, url: str) -> None:
        """
              Opens a website in a Chrome browser window.

              Args:
                  url (str): The URL of the website to open.
        """
        self.driver.get(url)
        self.driver.maximize_window()
        print(self.driver.title)

    def download_file(self) -> None:
        """
                Clears the download directory and downloads a file from the website.
        """

        for file in os.listdir(self.download_directory):
            path = os.path.join(self.download_directory, file)
            os.remove(path)

        try:
            file = self.driver.find_element(By.XPATH, self.DOWNLOAD_XPATH)
            file.click()
            sleep(2)
        except NoSuchElementException as e:
            print(f"Error of downloading: {str(e)}")

    def fill_form(self) -> None:
        """
        Fills out a form based on data from a downloaded Excel file.
        """
        self.click_start_button()

        file = os.listdir(self.download_directory)[0]

        if file.endswith('.xlsx'):
            path = os.path.join(self.download_directory, file)
            workbook = openpyxl.load_workbook(path)
            sheet = workbook.active
            col_count = sheet.max_column
            row_count = len([row for row in sheet if not all([cell.value is None for cell in row])])
            for row in sheet.iter_rows(min_row=2, max_col=col_count-1, max_row=row_count, values_only=True):
                first_name, last_name, company_name, role, address, email, phone = row
                self.fill_out_form_fields(first_name, last_name, email, phone, address, role, company_name)
                self.submit_form()

            sleep(4)
            self.driver.close()

        else:
            print("File not found")
            self.driver.close()

    def click_start_button(self) -> None:
        """
           Clicks the button START
        """
        try:
            form_element = self.driver.find_element(By.XPATH, self.START_BUTTON_XPATH)
            form_element.click()
        except NoSuchElementException as e:
            print(f"Error clicking start button: {str(e)}")

    def fill_out_form_fields(self, first_name: str, last_name: str, email: str, phone: str, address: str, role: str, company_name: str)-> None:
        """
        Fills out all fields in the form. List of these fields from Excel file:

        Args
             first_name (str): First Name
             last_name (str): Last Name
             email (str): Email
             phone (str): Phone Number
             address (str): Address
             role (str): Role in Company
             company_name (str): Company Name
        """
        self.fill_field('labelFirstName', first_name)
        self.fill_field('labelLastName', last_name)
        self.fill_field('labelEmail', email)
        self.fill_field('labelPhone', phone)
        self.fill_field('labelAddress', address)
        self.fill_field('labelRole', role)
        self.fill_field('labelCompanyName', company_name)

    def fill_field(self, field_name: str, value: str) -> None:
        """
        Fills out any field

        Args:
             field_name (str): name of field from the attribute ng-reflect-name in the form
             value (str): value for filling out
        """

        field_xpath = f"//input[@ng-reflect-name='{field_name}']"
        try:
            field = self.driver.find_element(By.XPATH, field_xpath)
            field.send_keys(value)
        except NoSuchElementException as e:
                print(f"Error filling field '{field_name}': {str(e)}")

    def submit_form(self) -> None:
        """
        Submits the form
        """
        try:
            submit_button = self.driver.find_element(By.XPATH, "//input[@type='submit']")
            submit_button.click()
        except NoSuchElementException as e:
            print(f"Error submitting the form: {str(e)}")


if __name__ == "__main__":
    rpa_bot = RpaChallengeBot()
    rpa_bot.open_website("https://rpachallenge.com/")
    rpa_bot.download_file()
    rpa_bot.fill_form()
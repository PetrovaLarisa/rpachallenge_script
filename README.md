# rpachallenge_script
The RPA Challenge Bot is a Python script that automates the process of downloading an Excel file from a website and filling out a form with data from the downloaded file.
It uses the Selenium WebDriver library for web automation and OpenPyXL for working with Excel files.
## Prerequisites

Before running the script, make sure you have the following dependencies installed:

- Python 3.x
- Selenium WebDriver (ChromeDriver)
- OpenPyXL
- Google Chrome browser

You can customize the script by modifying the following variables in the script:

DOWNLOAD_DIRECTORY: The directory where downloaded files will be saved.
DOWNLOAD_XPATH: The XPath to locate the download link on the website.
START_BUTTON_XPATH: The XPath to locate the start button on the website.

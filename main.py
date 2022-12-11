# Import driver
from selenium import webdriver
# By functions
from selenium.webdriver.common.by import By
# Send keys function
from selenium.webdriver.common.keys import Keys
# Work with data
import pandas as pd
# Time management
import time
# Create Excel file
import openpyxl

# Initialize browser
def initialize_browser(path:str):
    driver = webdriver.Chrome(path)
    driver.maximize_window()
    return driver

# Quit browser
def quit_browser(driver):
    driver.quit()

# Navigate to a page
def navigate_to_page(driver ,url:str):
    driver.get(url)

# Find an element by class
def find_class_element(driver, class_name:str):
    element = driver.find_element(By.CLASS_NAME, class_name)
    return element

# Find a list of elements by class
def find_class_elements(driver, class_name:str):
    elements = driver.find_elements_by_class_name(class_name)
    return elements

# Send keys to a text box
def keys_sender(element, key:str):
    element.clear()
    element.send_keys(key)
    element.send_keys(Keys.ENTER)

# Change image size option
def image_size_tool(driver):
    tools = find_class_element(driver, class_name='PNyWAd')
    tools.click()
    tools = find_class_element(driver, class_name='xFo9P')
    tools.click()
    tools = find_class_elements(driver, 'igM9Le')
    tools[2].click()

# Get image size
def image_finder(driver):
    image = driver.find_element(By.CLASS_NAME, 'rg_i')
    image.click()
    image = driver.find_element(By.CLASS_NAME, 'n3VNCb')
    size = image.size
    time.sleep(2)
    image = image.get_attribute('src')
    return(image, size)

# Image collector
def image_collector(driver, isbn:str, opt:int):
    # Condition for the text box input class name
    if opt == 0:
        class_name = 'gLFyf'
    else:
        class_name = 'og3lId'
    # Find text box input
    text_box = find_class_element(driver, class_name=class_name)
    keys_sender(text_box, key=isbn)
    # Size tool
    #########################image_size_tool(driver)
    # Find Image
    image = image_finder(driver)
    return(image)

# Read Excel inventory
def read_inventory(path):
    df = pd.read_excel(path)
    return df

# Create DataFrame object
def dataframe_creator():
    return pd.DataFrame(columns=['ISBN','SRC','HEIGHT','WIDTH'])

# Insert row to DataFrame
def insert_row(df, isbn:str, src:str, height:int, width:int):
    new_df = pd.DataFrame([[isbn, src, height, width]], columns=df.columns)
    df = pd.concat([df, new_df], ignore_index=True)
    return df

# Write a new Excel file
def excel_writer(df):
    df.to_excel('images_data.xlsx', sheet_name='Data')

# Main function
def main():
    # Start browser
    driver_path = './../../ChromeWebDriver/chromedriver'
    driver = initialize_browser(driver_path)

    # Navigate to google images
    url = 'https://images.google.com/'
    navigate_to_page(driver, url)

    # Read ISBNs file
    inv_path = "./../../../Downloads/isbn.xlsx"
    isbn_array = read_inventory(inv_path)

    # Create DataFrame for the collected images
    images_df = dataframe_creator()

    # Loop for images
    for i in range(0,len(isbn_array)):
        # Get isbn
        act_isbn = str(isbn_array.iloc[i,0])
        try:
            # Get Image features
            image = image_collector(driver=driver, isbn=act_isbn, opt=i)
            # Insert row with the image features
            images_df = insert_row(images_df, isbn=act_isbn, src=image[0], height=image[1]['height'], width=image[1]['width'])
        except:
            images_df = insert_row(images_df, isbn=act_isbn, src='N/A', height='N/A', width='N/A')
    
    # Quit browser
    quit_browser(driver)

    # Export excel file
    excel_writer(images_df)

if __name__=='__main__':
    main()

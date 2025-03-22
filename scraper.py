import pandas as pd
import time
import os
import random
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def setup_driver():
    """
    Set up and return a Chrome WebDriver for web scraping.
    """
    print("Setting up Chrome WebDriver...")
    
    # Configure Chrome options
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")  # Start maximized
    chrome_options.add_argument("--disable-notifications")  # Disable notifications
    
    # Uncomment the line below if you want to run Chrome in headless mode (no GUI)
    # chrome_options.add_argument("--headless")
    
    # Set up the driver with automatic ChromeDriver installation
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    
    print("Chrome WebDriver set up successfully!")
    return driver

def search_google_selenium(driver, query):
    """
    Search Google using Selenium WebDriver and return the search results page.
    """
    print(f"Searching for: {query}")
    
    try:
        # Navigate to Google
        driver.get("https://www.google.com")
        
        # Wait for the search box to appear and then type the query
        search_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "q"))
        )
        
        # Clear the search box and type the query
        search_box.clear()
        search_box.send_keys(query)
        search_box.submit()
        
        # Wait for the search results to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "search"))
        )
        
        print("Search successful")
        return True
    
    except Exception as e:
        print(f"Error during search: {str(e)}")
        return False

def extract_linkedin_url_selenium(driver):
    """
    Extract the first LinkedIn URL from the Google search results using Selenium.
    """
    try:
        # Look for all links on the page
        links = driver.find_elements(By.TAG_NAME, "a")
        
        # Find the first link that contains linkedin.com/in/
        for link in links:
            href = link.get_attribute("href")
            if href and "linkedin.com/in/" in href:
                print(f"Found LinkedIn URL: {href}")
                return href
        
        # Check if there are any LinkedIn cite elements
        cite_elements = driver.find_elements(By.TAG_NAME, "cite")
        for cite in cite_elements:
            cite_text = cite.text
            if "linkedin.com/in/" in cite_text:
                # Extract the LinkedIn URL using regex
                linkedin_pattern = r'(https?://)?(?:www\.)?linkedin\.com/in/[a-zA-Z0-9_-]+'
                match = re.search(linkedin_pattern, cite_text)
                if match:
                    linkedin_url = match.group(0)
                    if not linkedin_url.startswith('http'):
                        linkedin_url = 'https://' + linkedin_url
                    print(f"Found LinkedIn URL from cite: {linkedin_url}")
                    return linkedin_url
        
        print("No LinkedIn URL found")
        return ""
    
    except Exception as e:
        print(f"Error extracting LinkedIn URL: {str(e)}")
        return ""

def process_excel_selenium(input_file):
    """
    Process the input Excel file and update it with LinkedIn URLs using Selenium.
    """
    # Create output directory if it doesn't exist
    os.makedirs("output", exist_ok=True)
    
    print(f"Loading Excel file: {input_file}")
    try:
        # Try to load the Excel file
        df = pd.read_excel(input_file)
        print(f"Loaded {len(df)} records")
        
        # Check if required columns exist
        required_columns = ['Account Name', 'Full Name', 'Title', 'LinkedIn']
        for col in required_columns:
            if col not in df.columns:
                print(f"Error: Column '{col}' not found in the Excel file")
                return
        
        # Set up the Chrome WebDriver
        driver = setup_driver()
        
        try:
            # Process each row
            for index, row in df.iterrows():
                print(f"\n{'=' * 50}")
                print(f"Processing row {index+1}/{len(df)}")
                
                # Skip if LinkedIn URL already exists
                if pd.notna(row['LinkedIn']) and row['LinkedIn']:
                    print(f"LinkedIn URL already exists: {row['LinkedIn']}")
                    continue
                
                # Create search query components
                account_name = row['Account Name'] if pd.notna(row['Account Name']) else ""
                full_name = row['Full Name'] if pd.notna(row['Full Name']) else ""
                title = row['Title'] if pd.notna(row['Title']) else ""
                
                # Try different query formats
                queries = [
                    f"{full_name} {title} {account_name} ",
                    # f"\"{full_name}\" {title} linkedin",
                    # f"\"{full_name}\" \"{title}\" linkedin",
                    # f"\"{full_name}\" linkedin profile"
                ]
                
                linkedin_url = ""
                
                # Try each query until we find a LinkedIn URL
                for query_index, query in enumerate(queries):
                    if linkedin_url:
                        break
                    
                    print(f"Trying query {query_index+1}/{len(queries)}: {query}")
                    
                    # Perform the Google search
                    if search_google_selenium(driver, query):
                        # Take a screenshot for debugging
                        # screenshot_file = f"search_result_{index}_{query_index}.png"
                        # driver.save_screenshot(screenshot_file)
                        # print(f"Saved screenshot to {screenshot_file}")
                        
                        # Extract LinkedIn URL
                        linkedin_url = extract_linkedin_url_selenium(driver)
                        
                        # If found, break; otherwise try next query
                        if linkedin_url:
                            print(f"Success! Found LinkedIn URL: {linkedin_url}")
                        else:
                            print(f"Query didn't yield a LinkedIn URL, trying next query format...")
                    
                    # Add random delay to avoid being blocked
                    delay = random.uniform(3, 7)
                    print(f"Waiting for {delay:.2f} seconds before next search...")
                    time.sleep(delay)
                
                # Update DataFrame
                df.at[index, 'LinkedIn'] = linkedin_url
                
                # Save progress every row
                output_file = os.path.join("output", "updated_input.xlsx")
                df.to_excel(output_file, index=False)
                print(f"Progress saved to {output_file}")
                
                # Add random delay between people to avoid being blocked
                delay = random.uniform(5, 10)
                print(f"Waiting for {delay:.2f} seconds before next person...")
                time.sleep(delay)
        
        finally:
            # Always close the driver to release resources
            print("Closing WebDriver...")
            driver.quit()
    
    except FileNotFoundError:
        print(f"Error: File '{input_file}' not found!")
        print(f"Current working directory: {os.getcwd()}")
        print("Please make sure the file exists and is in the correct location.")
    
    except Exception as e:
        print(f"Error processing Excel file: {str(e)}")
        import traceback
        traceback.print_exc()
    
    # Save final result
    try:
        output_file = os.path.join("output", "updated_input.xlsx")
        df.to_excel(output_file, index=False)
        print(f"\nProcessing complete. Final results saved to {output_file}")
    except Exception as e:
        print(f"Error saving final results: {str(e)}")

def create_sample_input():
    """
    Create a sample input Excel file for testing.
    """
    try:
        df = pd.DataFrame({
            'Account Name': ['AgFirst Farm Credit Bank'],
            'Full Name': ['Crystal Norris-Jennings'],
            'Title': ['Senior Product Owner'],
            'LinkedIn': [''],
            'Location': [''],
            'Email': [''],
            'Phone Number': [''],
            'Revenue': [''],
            'Website': [''],
            'Employees': ['']
        })
        
        df.to_excel('input.xlsx', index=False)
        print("Sample input.xlsx file created successfully!")
    except Exception as e:
        print(f"Error creating sample file: {str(e)}")

if __name__ == "__main__":
    input_file = "input.xlsx"
    
    # Check if input file exists, create sample if not
    if not os.path.exists(input_file):
        print(f"Warning: {input_file} not found!")
        create_sample = input("Do you want to create a sample input file? (y/n): ").lower()
        if create_sample == 'y':
            create_sample_input()
        else:
            input_file = input("Please enter the path to your Excel file: ")
    
    process_excel_selenium(input_file)
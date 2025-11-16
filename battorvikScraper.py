import pandas as pd
import os
from io import StringIO
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

class BartTorvikScraper:
    def __init__(self):
        self.base_url = "https://barttorvik.com/playerstat.php"
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.excel_path = os.path.join(self.script_dir, "battorvikPlayerData.xlsx")
        self.driver = None
        # Define the columns we want to keep (includes Class)
        self.desired_columns = ['Rk', 'Player', 'Class', 'Team', 'Conf', 'Min%', 'PRPG!', 'BPM', 'ORtg', 'Usg', 'eFG', 'TS', 'OR', 'DR', 'Ast', 'TO', 'Blk', 'Stl', 'FTR', '2P', '3P/100', '3P']
        
    def _get_driver(self):
        """Initialize and return a Chrome WebDriver instance"""
        if self.driver is None:
            chrome_options = Options()
            chrome_options.add_argument('--headless')  # Run in headless mode
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-gpu')
            chrome_options.add_argument('user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
            
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
        return self.driver
        
    def close_driver(self):
        """Close the WebDriver instance"""
        if self.driver is not None:
            self.driver.quit()
            self.driver = None
        
    def get_url(self, year):
        """Construct URL with year parameter using f-string"""
        # Format: start date is November 1 of previous year, end date is May 1 of current year
        return f"{self.base_url}?link=y&year={year}&start={year-1}1101&end={year}0501"
    
    def scrape_data(self, year):
        """Scrape player statistics from Bart Torvik website"""
        url = self.get_url(year)
        driver = self._get_driver()
        
        try:
            # Navigate to the URL
            driver.get(url)
            
            # Wait for the page to load and JavaScript to execute
            wait = WebDriverWait(driver, 20)
            
            # Wait for the table to be present
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            
            # Additional wait for dynamic content to load
            time.sleep(5)
            
            # Click "Show 100 more" button repeatedly until all data is loaded
            print(f"    Loading all data for year {year}...")
            max_clicks = 30  # Safety limit
            clicks = 0
            
            for i in range(max_clicks):
                try:
                    # Try to find the Load More link/button
                    load_more = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, 
                            '//a[contains(translate(text(), "abcdefghijklmnopqrstuvwxyz", "ABCDEFGHIJKLMNOPQRSTUVWXYZ"), "SHOW 100 MORE")]'))
                    )
                    # Click using JavaScript to avoid issues
                    driver.execute_script("arguments[0].click();", load_more)
                    clicks += 1
                    time.sleep(2)  # Wait for new rows to load
                    
                    if clicks % 5 == 0:
                        # Check current row count
                        tables = driver.find_elements(By.TAG_NAME, 'table')
                        if len(tables) > 1:
                            rows = tables[1].find_elements(By.TAG_NAME, 'tr')
                            print(f"    Loaded {len(rows)} rows so far...")
                except Exception:
                    # No more Load More button found, all data loaded
                    break
            
            print(f"    Clicked 'Load More' {clicks} times. Extracting table...")
            time.sleep(3)  # Final wait for all data to render
            
            # Parse the table directly from Selenium to correctly extract player names and class
            # The table structure has issues with pandas parsing:
            # - "Player" column actually contains Class (Jr, Sr, So, Fr)
            # - Player name is in cell 4 (the actual player name)
            
            # Get the table directly from Selenium to parse correctly
            tables = driver.find_elements(By.TAG_NAME, 'table')
            if len(tables) < 2:
                raise ValueError("Player stats table not found")
            
            player_table = tables[1]  # Player stats table
            rows = player_table.find_elements(By.TAG_NAME, 'tr')
            
            if len(rows) < 2:
                raise ValueError("No data rows found in table")
            
            # Parse header row to find column positions
            header_row = rows[0]
            header_cells = header_row.find_elements(By.TAG_NAME, 'th')
            if len(header_cells) == 0:
                header_cells = header_row.find_elements(By.TAG_NAME, 'td')
            
            # Find column positions
            header_texts = [cell.text.strip() for cell in header_cells]
            
            # Parse data rows
            # The header row has columns at certain positions, but the actual data values
            # are offset. We need to map based on actual data positions:
            # Header positions: 0=RK, 2=PLAYER (class area), 4=PLAYER (name area), 6=TEAM, 7=CONF, 9=MIN%, 10=PRPG!, 12=BPM, 15=ORTG, 17=USG, 18=EFG, 19=TS, 20=OR, 21=DR, 22=AST, 23=TO, 25=BLK, 26=STL, 27=FTR, 33=2P, 35=3P/100, 36=3P
            # Data positions: 0=RK, 2=Class, 4=Player Name, 6=Team, 7=Conf, 10=Min%, 11=PRPG!, 13=BPM, 16=ORtg, 18=Usg, 19=eFG, 20=TS, 21=OR, 22=DR, 23=Ast, 24=TO, 26=Blk, 27=Stl, 28=FTR, 33=2P, 35=3P/100, 36=3P
            data_rows = []
            for row in rows[1:]:  # Skip header row
                cells = row.find_elements(By.TAG_NAME, 'td')
                if len(cells) < 20:  # Skip rows that don't have enough cells
                    continue
                
                # Extract values based on actual data cell positions
                row_data = {}
                if len(cells) > 0:
                    row_data['Rk'] = cells[0].text.strip()
                if len(cells) > 2:
                    row_data['Class'] = cells[2].text.strip()  # Class (Jr, Sr, So, Fr)
                if len(cells) > 4:
                    row_data['Player'] = cells[4].text.strip()  # Actual player name
                if len(cells) > 6:
                    row_data['Team'] = cells[6].text.strip()
                if len(cells) > 7:
                    row_data['Conf'] = cells[7].text.strip()
                if len(cells) > 10:
                    row_data['Min%'] = cells[10].text.strip()  # Data is at cell 10, header at cell 9
                if len(cells) > 11:
                    row_data['PRPG!'] = cells[11].text.strip()  # Data is at cell 11, header at cell 10
                if len(cells) > 13:
                    row_data['BPM'] = cells[13].text.strip()  # Data is at cell 13, header at cell 12
                if len(cells) > 16:
                    row_data['ORtg'] = cells[16].text.strip()  # Data is at cell 16, header at cell 15
                if len(cells) > 18:
                    row_data['Usg'] = cells[18].text.strip()  # Data is at cell 18, header at cell 17
                if len(cells) > 19:
                    row_data['eFG'] = cells[19].text.strip()  # Data is at cell 19, header at cell 18
                if len(cells) > 20:
                    row_data['TS'] = cells[20].text.strip()  # Data is at cell 20, header at cell 19
                if len(cells) > 21:
                    row_data['OR'] = cells[21].text.strip()  # Data is at cell 21, header at cell 20
                if len(cells) > 22:
                    row_data['DR'] = cells[22].text.strip()  # Data is at cell 22, header at cell 21
                if len(cells) > 23:
                    row_data['Ast'] = cells[23].text.strip()  # Data is at cell 23, header at cell 22
                if len(cells) > 24:
                    row_data['TO'] = cells[24].text.strip()  # Data is at cell 24, header at cell 23
                if len(cells) > 26:
                    row_data['Blk'] = cells[26].text.strip()  # Data is at cell 26, header at cell 25
                if len(cells) > 27:
                    row_data['Stl'] = cells[27].text.strip()  # Data is at cell 27, header at cell 26
                if len(cells) > 28:
                    row_data['FTR'] = cells[28].text.strip()  # Data is at cell 28, header at cell 27
                if len(cells) > 39:
                    row_data['2P'] = cells[39].text.strip()  # Data is at cell 39 (percentage), header at cell 33
                if len(cells) > 41:
                    row_data['3P/100'] = cells[41].text.strip()  # Data is at cell 41, header at cell 35
                if len(cells) > 43:
                    row_data['3P'] = cells[43].text.strip()  # Data is at cell 43 (percentage), header at cell 36
                
                data_rows.append(row_data)
            
            # Create DataFrame from parsed data
            df = pd.DataFrame(data_rows)
            
            # Only keep desired columns
            final_columns = [col for col in self.desired_columns if col in df.columns]
            df = df[final_columns]
            
            # Add year column to track which year the data is from
            df['Year'] = year
            
            print(f"    Extracted {len(df)} rows of data with {len(df.columns)} columns")
            return df
            
        except Exception as e:
            print(f"Error scraping data for year {year}: {e}")
            raise
    
    def append_to_excel(self, df):
        """Append DataFrame to existing Excel file"""
        # Ensure df has Year column and all desired columns
        desired_cols_with_year = self.desired_columns + ['Year']
        
        # Make sure df has all required columns
        for col in desired_cols_with_year:
            if col not in df.columns:
                df[col] = None
        
        # Select only desired columns in correct order
        df = df[desired_cols_with_year]
        
        # Check for existing file
        existing_df = None
        if os.path.exists(self.excel_path):
            try:
                existing_df = pd.read_excel(self.excel_path, sheet_name='Sheet1', engine='openpyxl')
                # Check if file actually has data
                if existing_df.empty or len(existing_df.columns) == 0:
                    existing_df = None
                    print(f"Existing file was empty, starting fresh")
            except Exception as e:
                print(f"Could not read existing file: {e}, starting fresh")
                existing_df = None
        
        # Process data
        if existing_df is not None and not existing_df.empty:
            # Ensure existing_df has the same columns
            for col in desired_cols_with_year:
                if col not in existing_df.columns:
                    existing_df[col] = None
            
            # Select only desired columns
            existing_df = existing_df[desired_cols_with_year]
            
            # Filter out rows for the same year (to avoid duplicates when re-running)
            if 'Year' in df.columns and 'Year' in existing_df.columns:
                year_to_append = df['Year'].iloc[0] if len(df) > 0 else None
                if year_to_append is not None:
                    existing_df = existing_df[existing_df['Year'] != year_to_append]
            
            # Combine existing and new data
            combined_df = pd.concat([existing_df, df], ignore_index=True)
            print(f"Appended {len(df)} new rows to existing {len(existing_df)} rows")
        else:
            # No existing data or file doesn't exist, use new data
            combined_df = df
            if existing_df is None:
                print(f"Creating new file with {len(df)} rows")
            else:
                print(f"File was empty, writing {len(df)} new rows")
        
        # Write to Excel file (overwrites the existing sheet)
        try:
            combined_df.to_excel(self.excel_path, index=False, sheet_name='Sheet1', engine='openpyxl')
            print(f"Data saved to {self.excel_path}")
            print(f"Total rows in file: {len(combined_df)}")
            print(f"Years in file: {sorted(combined_df['Year'].unique()) if 'Year' in combined_df.columns else 'N/A'}")
        except Exception as e:
            print(f"Error saving to Excel: {e}")
            raise
        
        return self.excel_path


if __name__ == "__main__":
    scraper = BartTorvikScraper()
    
    try:
        # Loop through years from 2025 to 2008 (inclusive)
        years = range(2025, 2007, -1)  # Goes from 2025 down to 2008
        total_years = len(years)
        successful = 0
        failed = 0
        
        print(f"Starting to scrape data for {total_years} years (2025 to 2008)...")
        print("=" * 60)
        
        for i, year in enumerate(years, 1):
            try:
                print(f"\n[{i}/{total_years}] Scraping data for year {year}...")
                df = scraper.scrape_data(year)
                print(f"Successfully scraped {len(df)} rows for year {year}")
                
                scraper.append_to_excel(df)
                successful += 1
                print(f"✓ Year {year} completed successfully")
                
            except Exception as e:
                failed += 1
                print(f"✗ Error scraping year {year}: {e}")
                print(f"Continuing with next year...")
        
        print("\n" + "=" * 60)
        print(f"Scraping completed!")
        print(f"Successfully scraped: {successful} years")
        print(f"Failed: {failed} years")
        print(f"Total years processed: {successful + failed}/{total_years}")
        
    finally:
        # Always close the driver
        scraper.close_driver()
        print("\nBrowser driver closed.")

import pandas as pd
import os
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

class NBAScraper:
    def __init__(self):
        self.base_url = "https://www.espn.com/nba/stats/player/_/season/"
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.excel_path = os.path.join(self.script_dir, "nbaPlayerData.xlsx")
        self.driver = None
        
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
        """Construct URL with year parameter"""
        return f"{self.base_url}{year}/seasontype/2"
    
    def extract_player_name_and_team(self, name_text):
        """Extract player name and team from ESPN format (e.g., 'ray allenMIA' or 'Shai Gilgeous-AlexanderOKC')"""
        # ESPN format: PlayerNameTeam (e.g., "ray allenMIA", "Shai Gilgeous-AlexanderOKC")
        # Team abbreviations are typically 2-4 uppercase letters at the end
        # Handle cases with slashes (multiple teams): "De'Aaron FoxSAC/SA" -> player: "De'Aaron Fox", team: "SAC/SA"
        if '/' in name_text:
            # Find where the team part starts (first uppercase team abbreviation)
            # Pattern: player name (may have mixed case) followed by uppercase team abbreviation(s) with slash
            match = re.match(r'^(.+?)([A-Z]{2,4}(?:/[A-Z]{2,4})+)$', name_text)
            if match:
                player_name = match.group(1)
                team = match.group(2)
                return player_name, team
            # Fallback: try without the slash pattern
            match = re.match(r'^(.+?)([A-Z]{2,4})', name_text)
            if match:
                player_name = match.group(1)
                # Get everything from the first team abbreviation to the end
                team = name_text[len(player_name):]
                return player_name, team
        else:
            # Standard format: PlayerNameTEAM
            match = re.match(r'^(.+?)([A-Z]{2,4})$', name_text)
            if match:
                player_name = match.group(1)
                team = match.group(2)
                return player_name, team
        # If no match, return as-is with empty team
        return name_text, ""
    
    def scrape_data(self, year):
        """Scrape player statistics from ESPN website"""
        url = self.get_url(year)
        driver = self._get_driver()
        
        try:
            # Navigate to the URL
            driver.get(url)
            
            # Wait for the page to load and JavaScript to execute
            wait = WebDriverWait(driver, 20)
            
            # Wait for the tables to be present
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.Table")))
            
            # Additional wait for dynamic content to load
            time.sleep(3)
            
            # Click "Show More" repeatedly to load all players (~600 per year)
            print(f"    Loading all players for year {year}...")
            max_clicks = 20  # Safety limit (should be enough for ~600 players: 50 initial + ~10 clicks * 50 = 550+)
            clicks = 0
            previous_row_count = 0
            
            for i in range(max_clicks):
                try:
                    # Try to find the Show More link/button
                    show_more = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "a.loadMore__link, a.AnchorLink.loadMore__link"))
                    )
                    
                    # Get current row count before clicking
                    tables_before = driver.find_elements(By.CSS_SELECTOR, "table.Table")
                    if len(tables_before) >= 2:
                        previous_row_count = len(tables_before[0].find_elements(By.CSS_SELECTOR, 'tbody tr'))
                    
                    # Click using JavaScript to avoid issues
                    driver.execute_script("arguments[0].click();", show_more)
                    clicks += 1
                    
                    # Wait for new rows to load
                    time.sleep(2)
                    
                    # Check if new rows were loaded
                    tables_after = driver.find_elements(By.CSS_SELECTOR, "table.Table")
                    if len(tables_after) >= 2:
                        current_row_count = len(tables_after[0].find_elements(By.CSS_SELECTOR, 'tbody tr'))
                        print(f"    Clicked 'Show More' {clicks} times. Now showing {current_row_count} players...")
                        
                        # If no new rows were loaded, we've reached the end
                        if current_row_count == previous_row_count:
                            print(f"    No new rows loaded, all data available")
                            break
                        previous_row_count = current_row_count
                    
                except Exception:
                    # No more Show More button found, all data loaded
                    print(f"    'Show More' button not found after {clicks} clicks, all data loaded")
                    break
            
            print(f"    Finished loading data. Extracting table...")
            time.sleep(2)  # Final wait for all data to render
            
            # Get both tables - ESPN splits the stats across two tables
            tables = driver.find_elements(By.CSS_SELECTOR, "table.Table")
            
            if len(tables) < 2:
                raise ValueError(f"Expected 2 tables but found {len(tables)}")
            
            # Table 0: RK and Name (player name + team)
            name_table = tables[0]
            # Table 1: All the stats (POS, GP, MIN, PTS, etc.)
            stats_table = tables[1]
            
            # Get headers from stats table
            stats_headers = []
            stats_header_row = stats_table.find_element(By.TAG_NAME, 'thead')
            stats_header_cells = stats_header_row.find_elements(By.TAG_NAME, 'th')
            stats_headers = [cell.text.strip() for cell in stats_header_cells]
            
            # Get data rows from both tables
            name_rows = name_table.find_elements(By.CSS_SELECTOR, 'tbody tr')
            stats_rows = stats_table.find_elements(By.CSS_SELECTOR, 'tbody tr')
            
            if len(name_rows) != len(stats_rows):
                print(f"    Warning: Name table has {len(name_rows)} rows, stats table has {len(stats_rows)} rows")
            
            # Parse all rows and combine data
            data_rows = []
            min_rows = min(len(name_rows), len(stats_rows))
            
            for i in range(min_rows):
                # Parse name table row
                name_cells = name_rows[i].find_elements(By.TAG_NAME, 'td')
                if len(name_cells) < 2:
                    continue
                
                rk = name_cells[0].text.strip()
                name_text = name_cells[1].text.strip()
                player_name, team = self.extract_player_name_and_team(name_text)
                
                # Parse stats table row
                stats_cells = stats_rows[i].find_elements(By.TAG_NAME, 'td')
                if len(stats_cells) < len(stats_headers):
                    continue
                
                # Build row data dictionary
                row_data = {
                    'RK': rk,
                    'Player': player_name,
                    'Team': team
                }
                
                # Add all stats columns
                for j, header in enumerate(stats_headers):
                    if j < len(stats_cells):
                        row_data[header] = stats_cells[j].text.strip()
                
                data_rows.append(row_data)
            
            # Create DataFrame from parsed data
            df = pd.DataFrame(data_rows)
            
            # Add year column to track which year the data is from
            df['Year'] = year
            
            print(f"    Extracted {len(df)} rows of data with {len(df.columns)} columns")
            return df
            
        except Exception as e:
            print(f"Error scraping data for year {year}: {e}")
            raise
    
    def append_to_excel(self, df):
        """Append DataFrame to existing Excel file"""
        if df.empty:
            print("No data to append")
            return self.excel_path
        
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
            # Get all unique columns from both dataframes
            all_columns = sorted(list(set(df.columns.tolist() + existing_df.columns.tolist())))
            
            # Ensure both dataframes have all columns
            for col in all_columns:
                if col not in df.columns:
                    df[col] = None
                if col not in existing_df.columns:
                    existing_df[col] = None
            
            # Reorder columns: RK, Player, Team first, then alphabetical stats, Year last
            priority_cols = ['RK', 'Player', 'Team']
            other_cols = sorted([col for col in all_columns if col not in priority_cols and col != 'Year'])
            column_order = priority_cols + other_cols + (['Year'] if 'Year' in all_columns else [])
            column_order = [col for col in column_order if col in all_columns]
            
            df = df[column_order]
            existing_df = existing_df[column_order]
            
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
            # Reorder columns: RK, Player, Team first, then alphabetical stats, Year last
            priority_cols = ['RK', 'Player', 'Team']
            other_cols = sorted([col for col in df.columns if col not in priority_cols and col != 'Year'])
            column_order = priority_cols + other_cols + (['Year'] if 'Year' in df.columns else [])
            column_order = [col for col in column_order if col in df.columns]
            combined_df = df[column_order]
            
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
    scraper = NBAScraper()
    
    try:
        # Loop through years from 2002 to 2025 (inclusive)
        years = range(2002, 2026)  # Goes from 2002 to 2025
        total_years = len(years)
        successful = 0
        failed = 0
        
        print(f"Starting to scrape data for {total_years} years (2002 to 2025)...")
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

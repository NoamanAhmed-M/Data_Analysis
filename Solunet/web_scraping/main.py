
import requests
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import Calendar
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from datetime import datetime, timedelta
import time
import pandas as pd
import re
from selenium.common.exceptions import NoSuchElementException

import os

class DateRangeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Date Range Selector")
        self.root.geometry("1000x700")

        # --- Create Canvas with Scrollbar ---
        outer_frame = ttk.Frame(root)
        outer_frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(outer_frame)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        v_scrollbar = ttk.Scrollbar(outer_frame, orient=tk.VERTICAL, command=canvas.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        canvas.configure(yscrollcommand=v_scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        # --- Main inner frame inside the canvas ---
        self.scrollable_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        # Allow resizing inside canvas
        self.scrollable_frame.bind("<Enter>", lambda e: self._bind_mousewheel(canvas))
        self.scrollable_frame.bind("<Leave>", lambda e: self._unbind_mousewheel(canvas))

        self.build_interface(self.scrollable_frame)

    def _bind_mousewheel(self, canvas):
        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))

    def _unbind_mousewheel(self, canvas):
        canvas.unbind_all("<MouseWheel>")

    def build_interface(self, main_frame):
        # Style configuration
        style = ttk.Style()
        style.configure('TButton', font=('Helvetica', 10), padding=5)
        style.configure('TLabel', font=('Helvetica', 10), padding=5)
        style.configure('TEntry', padding=5)

        # Grid expand
        for i in range(3):
            main_frame.grid_columnconfigure(i, weight=1)

        # ... now paste all your widgets setup as-is, replacing main_frame.pack(...) with this new frame ...
        ttk.Label(main_frame, text="Login Credentials:").grid(row=0, column=0, sticky="w", pady=5)

        ttk.Label(main_frame, text="Email:").grid(row=1, column=0, sticky="w", pady=2)
        self.email_entry = ttk.Entry(main_frame, width=30)
        self.email_entry.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        self.email_entry.insert(0, "")

        ttk.Label(main_frame, text="Password:").grid(row=2, column=0, sticky="w", pady=2)
        self.password_entry = ttk.Entry(main_frame, width=30, show="*")
        self.password_entry.grid(row=2, column=1, padx=2, pady=2, sticky="ew")
        self.password_entry.insert(0, "")
        ttk.Label(main_frame, text="Target URL (after login):").grid(row=3, column=0, sticky="w", pady=5)
        self.url_entry = ttk.Entry(main_frame, width=50)
        self.url_entry.grid(row=3, column=1, padx=2, pady=2, sticky="ew")
        self.url_entry.insert(0, "")
        # Date Range Selection Dropdown
        ttk.Label(main_frame, text="Select Date Range:").grid(row=4, column=0, sticky="w", pady=5)

        date_range_options = [
            "Today",           # 1
            "Yesterday",       # 2
            "Current Week",    # 3
            "Previous Week",   # 4
            "Current Month",   # 5
            "Previous Month",  # 6
            "Current Year",    # 7
            "Previous Year"    # 8
        ]

        self.date_range_var = tk.StringVar()
        self.date_range_var.set(date_range_options[0])  # Default: Today

        self.date_range_dropdown = ttk.Combobox(main_frame, textvariable=self.date_range_var,
                                                values=date_range_options, state="readonly")
        self.date_range_dropdown.grid(row=5, column=0, columnspan=2, padx=3, pady=3, sticky="ew")

        self.date_range_index = 1  # Default index

        self.date_range_dropdown.bind("<<ComboboxSelected>>", lambda e: setattr(
            self, "date_range_index", date_range_options.index(self.date_range_var.get()) + 1))


        # Save Path
        ttk.Label(main_frame, text="Save Excel File As:").grid(row=8, column=0, sticky="w", pady=5)
        self.save_path_var = tk.StringVar()
        self.save_path_entry = ttk.Entry(main_frame, textvariable=self.save_path_var, width=50)
        self.save_path_entry.grid(row=8, column=1, padx=3, pady=3, sticky="ew")
        browse_btn = ttk.Button(main_frame, text="Browse...", command=self.browse_save_path)
        browse_btn.grid(row=8, column=2, padx=5, pady=5)

        ttk.Button(main_frame, text="Start Automation", command=self.process_dates).grid(
            row=9, column=0, columnspan=3, pady=10, sticky="ew")

        ttk.Label(main_frame, text="Console Output:").grid(row=10, column=0, sticky="w", pady=5)
        self.console = tk.Text(main_frame, height=12, state='disabled', wrap=tk.WORD)
        self.console.grid(row=11, column=0, columnspan=3, sticky="nsew")
        main_frame.grid_rowconfigure(11, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.console.yview)
        scrollbar.grid(row=11, column=3, sticky="ns")
        self.console.config(yscrollcommand=scrollbar.set)

        self.email_entry.focus()
    
    def browse_save_path(self):
        filename = filedialog.asksaveasfilename(
            title="Select Excel File to Save",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="extracted_tables.xlsx"
        )
        if filename:
            self.save_path_var.set(filename)
    
    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.console.config(state='normal')
        self.console.insert(tk.END, f"[{timestamp}] {message}\n")
        self.console.see(tk.END)
        self.console.config(state='disabled')
        self.root.update()
    def process_dates(self):
        # Get credentials
        email = self.email_entry.get().strip()
        password = self.password_entry.get().strip()
        target_url = self.url_entry.get().strip()
        save_path = self.save_path_var.get().strip()
        
        if not email or not password:
            self.log_message("Error: Please enter email and password")
            messagebox.showerror("Error", "Please enter both email and password")
            return
        
        if not save_path:
            self.log_message("Error: Please select the Excel file save path before starting.")
            messagebox.showerror("Error", "Please select the Excel file save path before starting.")
            return
        
        # Get selected date range option
        date_range_option = self.date_range_var.get()
        self.log_message(f"Starting automation for date range: {date_range_option}")
        
        # Run the automation
        self.run_automation(email, password, target_url, save_path)
    
    def navigate_to_target(self, driver, target_url):
        """Robust navigation method with multiple fallbacks"""
        current_url = driver.current_url
        attempts = [
            lambda: driver.get(target_url),
            lambda: driver.execute_script(f"window.location.href = '{target_url}';"),
            lambda: driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.CONTROL + 'l' + target_url + Keys.RETURN)
        ]
        
        for attempt in attempts:
            try:
                attempt()
                WebDriverWait(driver, 10).until(
                    lambda d: d.current_url != current_url
                )
                return True
            except (TimeoutException, WebDriverException):
                continue
        
        return False
    
    def verify_login_success(self, driver):
        """Check multiple indicators of successful login"""
        indicators = [
            (By.XPATH, "//*[contains(text(), 'Welcome') or contains(text(), 'Dashboard')]"),
            (By.CSS_SELECTOR, "[href*='logout'], [onclick*='logout']"),
            (By.CLASS_NAME, "user-avatar"),
            (By.ID, "user-menu")
        ]
        
        for locator in indicators:
            try:
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located(locator)
                )
                return True
            except TimeoutException:
                continue
        return False

    def run_automation(self, email, password, target_url, save_path):
        date_range_index = self.date_range_index
        self.log_message(f"Using date range option #{date_range_index}")

        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--headless")  #  Enable headless mode
        chrome_options.add_argument("--disable-gpu")  # Optional, good for compatibility
        chrome_options.add_argument("--window-size=1920,1080")  # Optional, for full rendering

        manual_header = [
                    "inizio", "fine", "utente", "tipo attivita", "durata", "tempo fatturabile", "importo attivita",
                    "importo aggiuntivo", "importo di viaggio", "importo altri elementi", "importo articoli", "importo totale",
                    "stato fatturazione", "contatto", "commessa", "referente", "indirizzo", "etichetta", "completata",
                    "approvata", "note in report","Note interne", "incarico", "rapportino", "importo spese", "Email contatto",
                    "telefono contatto", "FAX contatto", "Email referente", "telefono referente", "cellulare referente",
                    "fax referente", "rapportino inviato il"
                ]
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.implicitly_wait(5)
        except Exception as e:
            self.log_message(f"Failed to start Chrome: {str(e)}")
            messagebox.showerror("Error", f"Failed to start Chrome: {str(e)}")
            return False

        all_dfs = []

        try:
            self.log_message("Navigating to login page...")
            login_url = ""
            self.driver.get(login_url)

            try:
                torna_button = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'torna a login')]"))
                )
                torna_button.click()
                self.log_message("✓ Clicked 'Torna a login' button")
            except TimeoutException:
                self.log_message("No 'Torna a login' button found - proceeding directly")

            email_field = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "email")))
            email_field.clear()
            email_field.send_keys(email)

            password_field = self.driver.find_element(By.ID, "password")
            password_field.clear()
            password_field.send_keys(password)

            login_button = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'accedi') or contains(@id, 'btn-login')]"))
            )
            login_button.click()
            self.log_message("✓ Clicked login button ('Accedi')")

            WebDriverWait(self.driver, 10).until(
                lambda d: "dashboard" in d.current_url.lower() or self.verify_login_success(d)
            )
            self.log_message("✓ Login successful")

            if target_url:
                self.log_message(f"Navigating to target URL: {target_url}")
                self.driver.get(target_url)
                WebDriverWait(self.driver, 10).until(
                    lambda d: target_url.split("/")[-1].lower() in d.current_url.lower()
                )
            try:
                time.sleep(2)
                Activity_button = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[10]/div[2]/div[3]/div[1]/ul/li[8]/div/a/div/span"))
                )
                Activity_button.click()
                time.sleep(2)
                Activity_search_button = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[10]/div[2]/div[3]/div[1]/ul/li[8]/ul/li[1]/div/a/div/span"))
                )
                Activity_search_button.click()
                self.log_message("Opened Activity search correctly")
                time.sleep(1)
            except TimeoutException:
                self.log_message("No Activity button found - proceeding directly")
            try:
                time.sleep(1)
                # Click type button
                type_button = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[10]/div[4]/section/div[1]/div/div[3]/div[1]/div/div[2]/div/div[2]/div[1]/span"))
                )
                type_button.click()
                time.sleep(1)

                # Click deselct option using JS
                self.driver.execute_script("""
                    const el = document.querySelector("#app > div:nth-child(3) > div.by-scroll-container > div > div:nth-child(2) > div > div:nth-child(2) > div.dr-dropdown-panel > div:nth-child(2) > div:nth-child(2)");
                    el?.click();
                """)

                # Select type settings
                for i, xpath in enumerate([
                    "/html/body/div[10]/div[4]/section/div[1]/div/div[3]/div[1]/div/div[2]/div/div[2]/div[2]/div[3]/div[1]/div/span[1]",
                    "/html/body/div[10]/div[4]/section/div[1]/div/div[3]/div[1]/div/div[2]/div/div[2]/div[2]/div[3]/div[2]/div/span[2]",
                    "/html/body/div[10]/div[4]/section/div[1]/div/div[3]/div[1]/div/div[2]/div/div[2]/div[2]/div[3]/div[3]/div/span[1]",
                    "/html/body/div[10]/div[4]/section/div[1]/div/div[3]/div[1]/div/div[2]/div/div[2]/div[2]/div[3]/div[4]/div/span[1]",
                    "/html/body/div[10]/div[4]/section/div[1]/div/div[3]/div[1]/div/div[2]/div/div[2]/div[2]/div[3]/div[5]/div/span[1]"
                ], start=1):
                    setting = WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath)))
                    setting.click()
                try:
                    
                    # Open date picker
                    try:
                        date_field_button = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, 
                                "/html/body/div[10]/div[4]/section/div[1]/div/div[3]/div[1]/div/div[4]/div/div/div/div[2]/span"))
                        )
                        date_field_button.click()
                        time.sleep(1)  # Small delay for calendar to open
                        select_date_option = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, 
                                f"/html/body/div[10]/div[4]/section/div[1]/div/div[3]/div[1]/div/div[4]/div/div/div/div[3]/div[1]/span[{date_range_index}]"))
                        )
                        select_date_option.click()
                        time.sleep(1)  # Small delay for calendar to open
                        act_button = WebDriverWait(self.driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, "/html/body/div[10]/div[4]/section/div[1]/div/div[3]/div[1]/div/div[3]/div/div[2]/div[1]/span"))
                        )
                        act_button.click()
                        time.sleep(1)

                        # Click deselct option using JS
                        self.driver.execute_script("""
                            const el = document.querySelector("#app > div:nth-child(3) > div.by-scroll-container > div > div:nth-child(3) > div > div:nth-child(2) > div.dr-dropdown-panel > div:nth-child(2) > div:nth-child(2)");
                            el?.click();
                        """)
                    except Exception as e:
                        print(f"Failed to open date picker: {str(e)}")
                        raise
                    
                    time.sleep(2)  # Wait for selection to complete

                except Exception as e:
                    print(f"Date selection failed: {str(e)}")
                    raise
                time.sleep(10)
            except TimeoutException:
                self.log_message(" Timeout: could not reset the settings - proceeding directly")
            all_dfs = []

            try:
                while True:
                    # Get pagination info
                    page_info = self.driver.find_element(By.XPATH, "//div[contains(text(), 'Pagina')]").text
                    parts = page_info.split()
                    current_page = int(parts[1])
                    total_pages = int(parts[3])
                    print(f"📄 Current page: {current_page} of {total_pages}")

                    try:
                        if not manual_header or not isinstance(manual_header, list):
                            self.log_message(" 'manual_header' is not defined or is invalid. Skipping table extraction.")
                            return False

                        tables = self.driver.find_elements(By.TAG_NAME, "table")
                        self.log_message(f" Found {len(tables)} table(s) on the page.")

                        for idx, table in enumerate(tables, start=1):
                            try:
                                tbody_rows = table.find_elements(By.CSS_SELECTOR, "tbody tr")
                                if not tbody_rows:
                                    tbody_rows = table.find_elements(By.TAG_NAME, "tr")
                                    self.log_message(f"⚠ Table #{idx} has no <tbody>; falling back to all <tr>.")

                                if not tbody_rows:
                                    self.log_message(f"⚠ Table #{idx} has no rows. Skipping.")
                                    continue

                                table_data = []
                                for row in tbody_rows:
                                    try:
                                        cells = row.find_elements(By.XPATH, "./th|./td")
                                        row_data = [cell.text.strip() for cell in cells]
                                        if any(row_data):
                                            table_data.append(row_data)
                                    except Exception as row_err:
                                        self.log_message(f"⚠ Error parsing row in Table #{idx}: {row_err}")

                                if not table_data:
                                    self.log_message(f"⚠ Table #{idx} contains only empty rows. Skipping.")
                                    continue

                                # Normalize rows to header
                                for r in range(len(table_data)):
                                    diff = len(manual_header) - len(table_data[r])
                                    if diff > 0:
                                        table_data[r].extend([""] * diff)
                                    elif diff < 0:
                                        table_data[r] = table_data[r][:len(manual_header)]

                                # Create DataFrame
                                df = pd.DataFrame(table_data, columns=manual_header)
                                df["incarico_extracted"] = None

                                if "incarico" in df.columns:
                                    for local_index_on_page, row_values in enumerate(table_data):

                                        incarico_value = str(row_values[manual_header.index("incarico")]).strip()
                                        if incarico_value:
                                            try:
                                                self.log_message(f"→ Row {local_index_on_page+1}: Clicking incarico span for value: {incarico_value[:10]}")

                                                js_click = f'''
                                                    const span = document.querySelector(
                                                        "#app > div.search-grid-wrapper > div:nth-child(2) > div > div > div > table > tbody > tr:nth-child({local_index_on_page+1}) > td:nth-child(23) > div > div > span"
                                                    );
                                                    if (span) {{
                                                        span.scrollIntoView({{behavior: 'smooth', block: 'center'}});
                                                        span.click();
                                                    }} else {{
                                                        console.warn("⚠ Span not found at row {local_index_on_page+1}");
                                                    }}
                                                '''
                                                self.driver.execute_script(js_click)
                                                time.sleep(1)  # Wait for modal to fully appear

                                                js_extract = '''
                                                    function getTextByXPath(xpath) {
                                                        const result = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
                                                        return result.singleNodeValue ? result.singleNodeValue.textContent.trim() : null;
                                                    }
                                                    return getTextByXPath("/html/body/div[6]/div/div/div/div/div/div/div/div[4]/div[3]/div[2]/div[2]");
                                                '''
                                                extracted_value = self.driver.execute_script(js_extract)
                                                df.at[local_index_on_page, "incarico_extracted"] = extracted_value
                                                time.sleep(1)
                                                self.log_message(f"✓ Extracted incarico: {extracted_value}")
                                                js_close = '''
                                                    const buttonSpan = document.querySelector(
                                                        "#vj-modal-manager > div > div > div > div > div > div > div > div:nth-child(6) > button.button-input-white > span:nth-child(1)"
                                                    );
                                                    if (buttonSpan) {
                                                        buttonSpan.scrollIntoView({behavior: 'smooth', block: 'center'});
                                                        buttonSpan.click();
                                                    } else {
                                                        console.warn("⚠ Close button not found in modal");
                                                    }
                                                '''
                                                self.driver.execute_script(js_close)
                                                time.sleep(2)

                                            except Exception as incarico_err:
                                                self.log_message(f"⚠ Error extracting incarico on row {local_index_on_page+1}: {incarico_err}")
                                else:
                                    self.log_message("⚠ 'incarico' column not found in DataFrame")
                                all_dfs.append(df)
                                self.log_message(f"Table #{idx} extracted with {len(df)} rows")

                            except Exception as ex_table:
                                self.log_message(f" Error extracting Table #{idx}: {ex_table}")

                    except Exception as e:
                        self.log_message(f" Critical error during table extraction: {e}")

                    if current_page >= total_pages:
                        break

                    # Click next page arrow
                    try:
                        js_click_arrow = '''
                            const arrowIcon = document.querySelector(
                                "#app > div.search-grid-wrapper > div:nth-child(2) > div > div > div > table > tbody > tr:nth-child(51) > td:nth-child(1) > div > div:nth-child(2) > span:nth-child(3)"
                            );
                            if (arrowIcon) {
                                arrowIcon.scrollIntoView({behavior: 'smooth', block: 'center'});
                                arrowIcon.click();
                            } else {
                                console.warn("⚠ Arrow icon not found at row 51");
                            }
                        '''
                        self.driver.execute_script(js_click_arrow)
                        time.sleep(2)
                    except Exception as e:
                        self.log_message(f" Failed to click next arrow: {e}")
                        break

            except Exception as e:
                self.log_message(f" Error in pagination logic: {e}")

            # After loop, save all to Excel
            if all_dfs:
                try:
                    full_df = pd.concat(all_dfs, ignore_index=True)
                    manipulated_full_df = full_df.drop(["importo attivita","importo aggiuntivo","importo di viaggio","importo altri elementi","importo articoli","importo totale","indirizzo","incarico","rapportino","importo spese","telefono contatto","FAX contatto","Email referente","telefono referente","cellulare referente","fax referente","rapportino inviato il"], axis = 1)
                    if 'completata' in manipulated_full_df.columns:
                        manipulated_full_df['completata'] = manipulated_full_df['completata'].apply(
                            lambda x: True if x == 'check_box' else False
                        )
                    if 'approvata' in manipulated_full_df.columns:
                        manipulated_full_df['approvata'] = manipulated_full_df['approvata'].apply(
                            lambda x: True if x == 'check_box' else False
                        )
                    manipulated_full_df = manipulated_full_df.drop(manipulated_full_df.tail(1).index)
                    manipulated_full_df.to_excel(save_path, index=False)
                    self.log_message(f"\n All pages exported to Excel: {save_path}")
                except Exception as ex_save:
                    self.log_message(f" Failed to save Excel file: {ex_save}")
            else:
                self.log_message("⚠ No data to export.")

        except Exception as e:
            self.log_message(f" Critical error in automation: {str(e)}")
            messagebox.showerror("Error", f"Automation failed: {str(e)}")
            return False
        finally:
            self.cleanup()
            
    def cleanup(self):
        try:
            if hasattr(self, "driver") and self.driver:
                self.driver.quit()
                self.log_message("Browser closed successfully.")
        except Exception as e:
            self.log_message(f"Error closing browser: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    
    # Install tkcalendar if not available
    try:
        from tkcalendar import Calendar
    except ImportError:
        import subprocess
        import sys
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "tkcalendar"])
            from tkcalendar import Calendar
        except:
            print("Failed to install tkcalendar. Please install it manually.")
            sys.exit(1)
    
    app = DateRangeApp(root)
    root.mainloop()

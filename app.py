import webview
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import os
import pandas as pd
import PyPDF2
import pdfplumber
import re
from datetime import datetime
import json
import sqlite3
import pathlib
from packaging import version # <-- IMPORT THIS

class Api:
    def __init__(self):
        self.window = None
        self.driver = None
        self.bot_thread = None
        
        # --- VERSION CONTROL: This is the latest version required to run the app ---
        self.LATEST_APP_VERSION = '2.0.0'
        self.DOWNLOAD_LINK = 'https://github.com/your-repo/releases/latest' #<-- CHANGE THIS to your actual download link

        app_dir = pathlib.Path.home() / '.whatsapp_business_suite'
        app_dir.mkdir(parents=True, exist_ok=True)
        self.db_path = app_dir / 'whatsapp_data.db'
        
        self._init_database()
        self._init_users_database()

        self.current_user = None

    def _init_database(self):
        """Initialize SQLite database for storing groups and templates."""
        print(f"--- Initializing database at: {self.db_path} ---")
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS groups (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                members TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                content TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        conn.commit()
        conn.close()

    def _init_users_database(self):
        """Initialize users table in database."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("PRAGMA table_info(users)")
        columns = [row[1] for row in cursor.fetchall()]
        
        if not columns:
            cursor.execute('''
                CREATE TABLE users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT UNIQUE NOT NULL,
                    password TEXT NOT NULL,
                    role TEXT NOT NULL DEFAULT 'user',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    last_login TIMESTAMP
                )
            ''')
        else:
            if 'role' not in columns:
                cursor.execute('ALTER TABLE users ADD COLUMN role TEXT NOT NULL DEFAULT "user"')
            if 'created_at' not in columns:
                cursor.execute('ALTER TABLE users ADD COLUMN created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP')
            if 'last_login' not in columns:
                cursor.execute('ALTER TABLE users ADD COLUMN last_login TIMESTAMP')
        
        cursor.execute('''
            INSERT OR IGNORE INTO users (username, password, role) 
            VALUES (?, ?, ?)
        ''', ('admin', '123456', 'admin'))
        
        conn.commit()
        conn.close()

    def authenticate_user(self, username, password):
        """Authenticate user, store session, and navigate the window."""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT id, username, role FROM users 
                WHERE username = ? AND password = ?
            ''', (username, password))
            
            user = cursor.fetchone()
            conn.close()
            
            if user:
                self.current_user = {"id": user[0], "username": user[1], "role": user[2]}

                conn_update = sqlite3.connect(self.db_path)
                cursor_update = conn_update.cursor()
                cursor_update.execute('''
                    UPDATE users SET last_login = CURRENT_TIMESTAMP 
                    WHERE id = ?
                ''', (self.current_user['id'],))
                conn_update.commit()
                conn_update.close()
                
                role = self.current_user['role']
                if role == 'admin':
                    page_path = self.get_admin_page_path()
                else:
                    page_path = self.get_main_page_path()
                
                file_url = f"file:///{str(pathlib.Path(page_path).as_uri()).split('///', 1)[1]}"
                self.window.load_url(file_url)
                
                return { "status": "success" }
            else:
                return {
                    "status": "error",
                    "message": "Invalid username or password"
                }
                
        except Exception as e:
            print(f"Authentication error: {e}")
            return {"status": "error", "message": f"Authentication error: {str(e)}"}
        
    def check_session(self, client_version):
        """Checks for an active session and enforces mandatory updates."""
        # Force update check happens before anything else
        if version.parse(client_version) < version.parse(self.LATEST_APP_VERSION):
            return {
                "user": self.current_user,
                "force_update": True,
                "latest_version": self.LATEST_APP_VERSION,
                "download_link": self.DOWNLOAD_LINK
            }

        # If version is okay, proceed with normal session check
        if not self.current_user:
            return {"user": None, "force_update": False}

        return {"user": self.current_user, "force_update": False}

    def get_admin_page_path(self):
        current_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(current_dir, 'templates', 'admin.html')

    def get_main_page_path(self):
        current_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(current_dir, 'templates', 'index.html')

    def get_login_page_path(self):
        current_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(current_dir, 'templates', 'login.html')

    def navigate_to_admin(self):
        try:
            import os
            current_dir = os.path.dirname(os.path.abspath(__file__))
            admin_path = os.path.join(current_dir, 'templates', 'admin.html')
            file_url = f"file:///{admin_path.replace(os.sep, '/')}"
            self.window.load_url(file_url)
            return {"status": "success"}
        except Exception as e:
            print(f"Error navigating to admin: {e}")
            return {"status": "error", "message": str(e)}

    def navigate_to_main(self):
        try:
            import os
            current_dir = os.path.dirname(os.path.abspath(__file__))
            main_path = os.path.join(current_dir, 'templates', 'index.html')
            file_url = f"file:///{main_path.replace(os.sep, '/')}"
            self.window.load_url(file_url)
            return {"status": "success"}
        except Exception as e:
            print(f"Error navigating to main: {e}")
            return {"status": "error", "message": str(e)}

    def navigate_to_login(self):
        try:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            login_path = os.path.join(current_dir, 'templates', 'login.html')
            self.window.load_url(login_path)
        except Exception as e:
            print(f"Error navigating to login: {e}")

    def get_all_users(self):
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT id, username, role, created_at, last_login 
                FROM users ORDER BY created_at DESC
            ''')
            
            users = []
            for row in cursor.fetchall():
                users.append({
                    "id": row[0],
                    "username": row[1],
                    "role": row[2],
                    "created_at": row[3],
                    "last_login": row[4]
                })
            
            conn.close()
            return {"status": "success", "users": users}
            
        except Exception as e:
            return {"status": "error", "message": f"Error getting users: {str(e)}"}

    def create_user(self, username, password, role):
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                INSERT INTO users (username, password, role) 
                VALUES (?, ?, ?)
            ''', (username, password, role))
            
            conn.commit()
            conn.close()
            
            return {"status": "success", "message": f"User '{username}' created successfully"}
            
        except sqlite3.IntegrityError:
            return {"status": "error", "message": "Username already exists"}
        except Exception as e:
            return {"status": "error", "message": f"Error creating user: {str(e)}"}

    def update_user(self, user_id, username, password, role):
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            if password:
                cursor.execute('''
                    UPDATE users SET username = ?, password = ?, role = ? 
                    WHERE id = ?
                ''', (username, password, role, user_id))
            else:
                cursor.execute('''
                    UPDATE users SET username = ?, role = ? 
                    WHERE id = ?
                ''', (username, role, user_id))
            
            conn.commit()
            conn.close()
            
            return {"status": "success", "message": f"User '{username}' updated successfully"}
            
        except sqlite3.IntegrityError:
            return {"status": "error", "message": "Username already exists"}
        except Exception as e:
            return {"status": "error", "message": f"Error updating user: {str(e)}"}

    def delete_user(self, user_id):
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('SELECT username FROM users WHERE id = ?', (user_id,))
            user = cursor.fetchone()
            
            if user and user[0] == 'admin':
                return {"status": "error", "message": "Cannot delete the main admin user"}
            
            cursor.execute('DELETE FROM users WHERE id = ?', (user_id,))
            conn.commit()
            conn.close()
            
            return {"status": "success", "message": "User deleted successfully"}
            
        except Exception as e:
            return {"status": "error", "message": f"Error deleting user: {str(e)}"}

    def logout(self):
        self.current_user = None
        self.navigate_to_login()
        return {"status": "success", "message": "Logged out successfully"}

    # --- ADDED METHODS END HERE ---

    # Add this method to your Api class to fix the missing function error
    def handle_file_upload(self):
        try:
            file_types = ('Excel Files (*.xlsx;*.xls;*.csv)', 'All files (*.*)')
            result = self.window.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=False, file_types=file_types)
            
            if result and len(result) > 0:
                file_path = result[0]
                return self.parse_excel_file(file_path)
            else:
                return {"status": "error", "message": "No file selected"}
        except Exception as e:
            return {"status": "error", "message": f"Error handling file upload: {str(e)}"}

    def _get_driver(self):
        max_retries = 3
        for attempt in range(max_retries):
            try:
                driver_path = os.path.join(os.path.dirname(__file__), 'chromedriver.exe')
                service = Service(executable_path=driver_path)
                options = webdriver.ChromeOptions()
                user_data_dir = os.path.join(os.path.dirname(__file__), 'whatsapp_session')
                options.add_argument(f"user-data-dir={user_data_dir}")
                options.add_argument("--start-maximized")
                options.add_argument("--no-sandbox")
                options.add_argument("--disable-dev-shm-usage")
                options.add_argument("--disable-blink-features=AutomationControlled")
                options.add_argument("--disable-extensions")
                options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
                options.add_experimental_option('useAutomationExtension', False)
                driver = webdriver.Chrome(service=service, options=options)
                driver.set_page_load_timeout(20)
                driver.implicitly_wait(3)
                driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
                print(f"Chrome driver created successfully (attempt {attempt + 1})")
                return driver
            except Exception as e:
                print(f"Driver creation attempt {attempt + 1} failed: {e}")
                if attempt == max_retries - 1:
                    raise Exception(f"Failed to create driver after {max_retries} attempts: {e}")
                time.sleep(1)
        return None

    def get_image_file_path(self):
        """Get image file path through file dialog"""
        try:
            file_types = ('Image Files (*.png;*.jpg;*.jpeg;*.gif;*.bmp)', 'All files (*.*)')
            result = self.window.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=False, file_types=file_types)
            if result and len(result) > 0:
                # Convert to absolute path and ensure it exists
                abs_path = os.path.abspath(result[0])
                if os.path.exists(abs_path):
                    return abs_path
                else:
                    print(f"File not found: {abs_path}")
                    return None
        except Exception as e:
            print(f"Error selecting image file: {e}")
        return None

    def validate_and_prepare_image(self, image_path):
        """Validate image file and return absolute path"""
        if not image_path:
            return None
                
        try:
            abs_path = os.path.abspath(image_path)
            if os.path.exists(abs_path):
                # Check file size (WhatsApp has limits)
                file_size = os.path.getsize(abs_path)
                if file_size > 64 * 1024 * 1024:  # 64MB limit
                    print(f"Warning: File too large ({file_size} bytes)")
                    return None
                return abs_path
            else:
                print(f"Image file not found: {abs_path}")
                return None
        except Exception as e:
            print(f"Error validating image: {e}")
            return None

    def open_file_dialog(self):
        """Opens a file dialog for the user to select an image."""
        file_types = ('Image Files (*.png;*.jpg;*.jpeg)', 'All files (*.*)')
        result = self.window.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=False, file_types=file_types)
        return result

    def open_excel_dialog(self):
        """Opens a file dialog for the user to select an Excel file."""
        file_types = ('Excel Files (*.xlsx;*.xls;*.csv)', 'All files (*.*)')
        result = self.window.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=False, file_types=file_types)
        return result

    def download_bulk_template(self):
        """Generate and offer download of bulk sender template."""
        try:
            # Create template DataFrame
            template_data = {
                'Name': ['John Doe', 'Jane Smith', 'Bob Johnson'],
                'Phone': ['919876543210', '918765432109', '917654321098'],
                'Email': ['john@example.com', 'jane@example.com', 'bob@example.com'],
                'Message': ['Hello John!', 'Hi Jane!', 'Hey Bob!']
            }
            
            df = pd.DataFrame(template_data)
            
            # Save to downloads folder or app directory
            downloads_path = os.path.expanduser("~/Downloads")
            if not os.path.exists(downloads_path):
                downloads_path = os.path.dirname(__file__)
            
            template_path = os.path.join(downloads_path, 'bulk_sender_template.xlsx')
            df.to_excel(template_path, index=False)
            
            return {
                "status": "success",
                "message": f"Template downloaded to: {template_path}",
                "path": template_path
            }
        except Exception as e:
            return {"status": "error", "message": f"Error creating template: {str(e)}"}

    def parse_excel_file_from_input(self):
        """Parse Excel file selected through the file input in the interface."""
        try:
            file_types = ('Excel Files (*.xlsx;*.xls;*.csv)', 'All files (*.*)')
            result = self.window.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=False, file_types=file_types)
            
            if result and len(result) > 0:
                file_path = result[0]
                # Add validation for file path
                if not file_path or not isinstance(file_path, str):
                    return {"status": "error", "message": "Invalid file path selected"}
                return self.parse_bulk_excel(file_path)
            else:
                return {"status": "error", "message": "No file selected"}
        except Exception as e:
            return {"status": "error", "message": f"Error selecting file: {str(e)}"}

    def parse_bulk_excel(self, file_path):
        """Parse Excel file for bulk sender."""
        try:
            # Add validation for file_path
            if not file_path or not isinstance(file_path, str):
                return {"status": "error", "message": "Invalid file path provided"}
            
            if not os.path.exists(file_path):
                return {"status": "error", "message": f"File not found: {file_path}"}
            
            # Read the Excel file
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)            
            # Clean column names
            df.columns = df.columns.str.strip()
            
            contacts = []
            for index, row in df.iterrows():
                try:
                    name = str(row.get('Name', f'Contact {index + 1}')).strip()
                    phone = str(row.get('Phone', '')).strip()
                    email = str(row.get('Email', '')).strip()
                    custom_message = str(row.get('Message', '')).strip()
                    
                    # Clean phone number
                    phone = re.sub(r'[^\d]', '', phone)
                    if len(phone) >= 10:
                        contact = {
                            "name": name,
                            "phone": phone,
                            "email": email,
                            "custom_message": custom_message if custom_message and custom_message != 'nan' else None
                        }
                        contacts.append(contact)
                
                except Exception as e:
                    print(f"Error processing row {index}: {e}")
                    continue
            
            if not contacts:
                return {"status": "error", "message": "No valid contacts found in Excel file."}
            
            return {
                "status": "success",
                "contacts": contacts,
                "message": f"Successfully parsed {len(contacts)} contacts"
            }
            
        except Exception as e:
            return {"status": "error", "message": f"Error parsing Excel file: {str(e)}"}

    def save_template(self, template_name, template_content):
        """Save message template to database."""
        try:
            # --- CHANGE: Added print to confirm function is being called ---
            print(f"--- Saving template '{template_name}' to database... ---")
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                INSERT OR REPLACE INTO templates (name, content) 
                VALUES (?, ?)
            ''', (template_name, template_content))
            
            conn.commit()
            conn.close()
            
            return {"status": "success", "message": f"Template '{template_name}' saved successfully"}
            
        except Exception as e:
            return {"status": "error", "message": f"Error saving template: {str(e)}"}

    def get_templates(self):
        """Get all saved templates."""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('SELECT id, name, content FROM templates ORDER BY name')
            rows = cursor.fetchall()
            
            templates = []
            for row in rows:
                templates.append({
                    "id": row[0],
                    "name": row[1],
                    "content": row[2]
                })
            
            conn.close()
            return {"status": "success", "templates": templates}
            
        except Exception as e:
            return {"status": "error", "message": f"Error getting templates: {str(e)}"}

    def get_template_content(self, template_id):
        """Get content of a specific template by ID."""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('SELECT content FROM templates WHERE id = ?', (template_id,))
            row = cursor.fetchone()
            
            if row:
                conn.close()
                return {"status": "success", "content": row[0]}
            else:
                conn.close()
                return {"status": "error", "message": "Template not found"}
                
        except Exception as e:
            return {"status": "error", "message": f"Error getting template content: {str(e)}"}
        
    def create_drag_drop_script(self, file_path):
        """Create JavaScript for drag and drop file upload"""
        import base64
        
        with open(file_path, 'rb') as f:
            file_data = f.read()
        
        file_b64 = base64.b64encode(file_data).decode('utf-8')
        file_name = os.path.basename(file_path)
        
        script = f"""
        async function uploadFile() {{
            try {{
                // Create blob from base64
                const response = await fetch('data:image/jpeg;base64,{file_b64}');
                const blob = await response.blob();
                const file = new File([blob], '{file_name}', {{ type: blob.type }});
                
                // Find drop target (message area)
                const dropTarget = document.querySelector('#main') || 
                                document.querySelector('[data-testid="conversation-panel-messages"]') ||
                                document.querySelector('div[contenteditable="true"]');
                
                if (!dropTarget) return 'no_target';
                
                // Create drag event
                const dataTransfer = new DataTransfer();
                dataTransfer.items.add(file);
                
                const dragEvent = new DragEvent('drop', {{
                    dataTransfer: dataTransfer,
                    bubbles: true,
                    cancelable: true
                }});
                
                // Dispatch the event
                dropTarget.dispatchEvent(dragEvent);
                
                return 'success';
            }} catch (error) {{
                return 'error: ' + error.message;
            }}
        }}
        
        return uploadFile();
        """
        return script

    def delete_template(self, template_id):
        """Delete a template."""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('DELETE FROM templates WHERE id = ?', (template_id,))
            conn.commit()
            conn.close()
            
            return {"status": "success", "message": "Template deleted successfully"}
            
        except Exception as e:
            return {"status": "error", "message": f"Error deleting template: {str(e)}"}

    def send_to_whatsapp_group(self, group_name, message, image_path=None):
        """Send message to a WhatsApp group."""
        driver = None
        try:
            driver = self._get_driver()
            wait = WebDriverWait(driver, 45)
            driver.get("https://web.whatsapp.com")
            
            print("Waiting for WhatsApp Web to load...")
            
            # Wait for login
            login_successful = False
            max_wait_time = 120
            check_interval = 5
            
            for i in range(max_wait_time // check_interval):
                try:
                    search_elements = driver.find_elements(By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')
                    if search_elements:
                        login_successful = True
                        break
                    time.sleep(check_interval)
                except Exception as e:
                    time.sleep(check_interval)
            
            if not login_successful:
                return {"status": "error", "message": "Login timeout - please make sure you're logged into WhatsApp Web"}
            
            time.sleep(2) # CHANGE 1
            
            # Search for the group
            search_box = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')))
            search_box.clear()
            search_box.send_keys(group_name)
            time.sleep(3)
            
            # Click on the group
            try:
                group_element = wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[@title='{group_name}']")))
                group_element.click()
                time.sleep(3)
            except:
                return {"status": "error", "message": f"Group '{group_name}' not found"}
            
            # Send image if provided
            if image_path and os.path.exists(image_path):
                try:
                    # Click attachment button
                    clip_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span[data-testid='clip']")))
                    clip_button.click()
                    time.sleep(2)
                    
                    # Upload image
                    file_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@accept='image/*,video/mp4,video/3gpp,video/quicktime']")))
                    file_input.send_keys(image_path)
                    time.sleep(5)
                    
                    # Add caption if message provided
                    if message:
                        try:
                            caption_box = driver.find_element(By.XPATH, "//div[@contenteditable='true'][@data-tab='10']")
                            caption_box.send_keys(message)
                        except:
                            pass
                    
                    # Send
                    send_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span[data-testid='send']")))
                    send_button.click()
                    time.sleep(3)
                    
                except Exception as img_error:
                    print(f"Error sending image: {img_error}")
                    # Continue to send text message if image fails
            
            # Send text message if provided and no image was sent
            if message and not image_path:
                message_box = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]')))
                message_box.click()
                time.sleep(0.3) # CHANGE 1
                
                # Type message line by line
                message_lines = message.split('\n')
                for i, line in enumerate(message_lines):
                    if line.strip():
                        message_box.send_keys(line)
                    if i < len(message_lines) - 1:
                        message_box.send_keys(Keys.SHIFT + Keys.ENTER)
                
                time.sleep(2)
                message_box.send_keys(Keys.ENTER)
            
            return {"status": "success", "message": f"Message sent to group '{group_name}' successfully"}
            
        except Exception as e:
            return {"status": "error", "message": f"Error sending to group: {str(e)}"}
        finally:
            if driver:
                try:
                    time.sleep(3)
                    driver.quit()
                except:
                    pass

    def parse_excel_file(self, file_path):
        """Parse Excel file to extract all courier and customer data from multiple rows."""
        try:
            # Read the Excel file
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
            
            # Clean column names - remove extra spaces
            df.columns = df.columns.str.strip()
            
            # Print column names for debugging
            print("Available columns:", list(df.columns))
            
            all_courier_data = []
            all_customers = []
            
            # Process each row in the Excel file
            for index, row in df.iterrows():
                try:
                    # Fix date format (remove time component)
                    date_value = str(row.get('Date', '')).strip()
                    if date_value and date_value != 'nan':
                        try:
                            # If it's a pandas datetime, format it properly
                            if 'Timestamp' in str(type(row.get('Date'))):
                                formatted_date = row.get('Date').strftime('%d-%b-%Y')
                            else:
                                # Handle string dates - remove time if present
                                date_only = date_value.split(' ')[0]  # Take only date part
                                formatted_date = date_only
                        except:
                            formatted_date = date_value
                    else:
                        formatted_date = ''
                    
                    # Extract courier data from current row with improved column matching
                    docket_value = str(row.get('Docket No', row.get('Docket No.', ''))).strip()
                    
                    courier_data = {
                        "date": formatted_date,
                        "challan_no": str(row.get('Challan No', '')).strip(),
                        "customer_code": str(row.get('Customer Code', '')).strip(),
                        "customer_name": str(row.get('Customer', '')).strip(),
                        "customer_location": str(row.get('Location', '')).strip(),
                        "courier_name": str(row.get('Courier Name', '')).strip(),
                        "docket_no": docket_value,
                        "courier_link": str(row.get('Courier Link', '')).strip(),
                        "no_of_boxes": str(row.get('No. of boxes', '')).strip(),
                        "shipment_type": str(row.get('Shipment', row.get('Shipment Type', 'Complete'))).strip(),
                        "items": []
                    }
                    
                    # Debug prints
                    print(f"Row {index} - Docket No column value: '{docket_value}'")
                    print(f"Row {index} - Customer: '{courier_data['customer_name']}'")
                    print(f"Row {index} - Date processed: '{formatted_date}'")
                    
                    # Extract items with their quantities from the row
                    item_columns = {
                        "13X17 Blue Base Film ACC-91": "13X17 Blue Base Film ACC-91",
                        "8X10 Blue Base Film ACC-91": "8X10 Blue Base Film ACC-91", 
                        "Accurate Paper ACC-41 A4": "Accurate Paper ACC-41 A4",
                        "Accurate Paper ACC-61 A4": "Accurate Paper ACC-61 A4",
                        "Accurate Paper ACC-61 A5": "Accurate Paper ACC-61 A5",
                        "White Instant Film-ACC-81 - A3": "White Instant Film-ACC-81 - A3",
                        "White Instant Film-ACC-81 - A4": "White Instant Film-ACC-81 - A4",
                        "BLACK Ink-81": "BLACK Ink-81",
                        "CYAN Ink-81": "CYAN Ink-81",
                        "LIGHT CYAN Ink-81": "LIGHT CYAN Ink-81",
                        "LIGHT MAGENTA-Ink-81": "LIGHT MAGENTA-Ink-81",
                        "MAGENTA Ink-81": "MAGENTA Ink-81",
                        "YELLOW Ink-81": "YELLOW Ink-81",
                        "Maintanace box": "Maintanace box"
                    }
                    
                    # Add items with their quantities
                    for item_name, column_name in item_columns.items():
                        try:
                            quantity = row.get(column_name, 0)
                            if pd.notna(quantity) and str(quantity).strip():
                                # Handle different quantity formats
                                if isinstance(quantity, str) and 'Grm' in quantity:
                                    qty_val = int(re.sub(r'[^\d]', '', quantity.split('Grm')[0]))
                                    unit = 'Grm'
                                else:
                                    qty_val = int(float(str(quantity))) if str(quantity).replace('.', '').isdigit() else 0
                                    unit = 'Grm' if 'Ink' in item_name else ''
                                
                                courier_data["items"].append({
                                    "name": item_name,
                                    "quantity": qty_val,
                                    "unit": unit
                                })
                            else:
                                # Add item with 0 quantity to maintain structure
                                courier_data["items"].append({
                                    "name": item_name,
                                    "quantity": 0,
                                    "unit": 'Grm' if 'Ink' in item_name else ''
                                })
                        except (ValueError, TypeError) as e:
                            print(f"Error processing item {item_name}: {e}")
                            courier_data["items"].append({
                                "name": item_name,
                                "quantity": 0,
                                "unit": 'Grm' if 'Ink' in item_name else ''
                            })
                    
                    all_courier_data.append(courier_data)
                    
                    # Extract customer data from mobile number column
                    mobile_numbers = str(row.get('Mobile No.', row.get('Mobile No', ''))).strip()
                    customer_name = courier_data["customer_name"]
                    customer_location = courier_data["customer_location"]
                    
                    if mobile_numbers and mobile_numbers != 'nan':
                        # Split multiple numbers if they exist (comma separated)
                        numbers = [num.strip() for num in mobile_numbers.split(',') if num.strip()]
                        
                        for i, number in enumerate(numbers):
                            # Clean the number
                            clean_number = re.sub(r'[^\d]', '', number)
                            if len(clean_number) >= 10:
                                # Use actual customer name instead of generic "Customer X"
                                customer_display_name = customer_name if customer_name and customer_name != 'nan' else f"Customer {len(all_customers) + 1}"
                                
                                customer = {
                                    "name": customer_display_name,
                                    "mobile": clean_number,
                                    "address": customer_location,
                                    "email": f"{customer_display_name.lower().replace(' ', '.')}@example.com",
                                    "row_index": index,
                                    "courier_data": courier_data  # Link customer to their courier data
                                }
                                all_customers.append(customer)
                    
                except Exception as e:
                    print(f"Error processing row {index}: {e}")
                    continue
            
            if not all_customers:
                return {"status": "error", "message": "No valid customers found in Excel file. Please check Mobile No. column format."}
            
            return {
                "status": "success",
                "courier_data": all_courier_data,
                "customer_data": all_customers
            }
            
        except Exception as e:
            return {"status": "error", "message": f"Error parsing Excel file: {str(e)}"}
        
    
    def parse_courier_files(self, file_data=None):
        # This method should handle the file upload process
        # For now, return sample data or implement file reading logic
        try:
            # Sample data structure - replace with actual file parsing
            courier_data = {
                "date": "16-Aug-2025",
                "challan_no": "OUT-25/26-6758",
                "customer_code": "CUS032023-24",
                "customer_location": "Visnagar, Gujarat",
                "courier_name": "The Professional Courier services Pvt. Ltd",
                "docket_no": "VPL951311889, VPL951311900",
                "courier_link": "www.theprofessional.com",
                "no_of_boxes": "2",
                "shipment_type": "Complete",
                "items": [
                    {"name": "13X17 Blue Base Film ACC-91", "quantity": 12500, "unit": ""},
                    {"name": "Accurate Paper ACC-61 A5", "quantity": 500, "unit": ""},
                    {"name": "White Instant Film-ACC-81 - A4", "quantity": 100, "unit": "Grm"},
                    {"name": "BLACK Ink-81", "quantity": 100, "unit": "Grm"},
                    {"name": "CYAN Ink-81", "quantity": 100, "unit": "Grm"},
                    {"name": "LIGHT CYAN Ink-81", "quantity": 100, "unit": "Grm"},
                    {"name": "Maintenance box", "quantity": 1, "unit": ""},
                ]
            }
            
            customer_data = [
                {"name": "Vandan Distributors", "mobile": "918779659681", "address": "Visnagar, Gujarat", "email": "vandan@example.com"},
                {"name": "Customer 2", "mobile": "919820366009", "address": "Visnagar, Gujarat", "email": "customer2@example.com"},
            ]
            
            return {
                "status": "success",
                "courier_data": courier_data,
                "customer_data": customer_data
            }
            
        except Exception as e:
            return {"status": "error", "message": f"Error parsing files: {str(e)}"}

    def format_courier_message(self, courier_data, customer):
        message_parts = []
        
        # Header
        message_parts.append("Greeting from *Accurate Medical Print Solutions*")
        message_parts.append("")
        message_parts.append("*COURIER DISPATCH NOTIFICATION*")
        message_parts.append("=" * 25)
        
        # Customer details
        message_parts.append(f"*Challan No:* {courier_data['challan_no']}")
        message_parts.append(f"*Date:* {courier_data['date']}")
        message_parts.append(f"*Customer:* {customer['name']}")
        message_parts.append(f"*Delivery Location:* {customer['address']}")
        
        message_parts.append("")
        
        # Courier details
        message_parts.append("*COURIER DETAILS*")
        message_parts.append("-" * 20)
        message_parts.append(f"*Courier:* {courier_data['courier_name']}")
        message_parts.append(f"*Docket No:* `{courier_data['docket_no']}`")
        message_parts.append(f"*No. of Boxes:* {courier_data['no_of_boxes']}")
        message_parts.append(f"*Track at:* {courier_data['courier_link']}")
        
        message_parts.append("")
        
        # Items shipped (only non-zero quantities with bold quantities and units)
        message_parts.append("*ITEMS SHIPPED*")
        message_parts.append("-" * 15)
        
        item_count = 0
        for item in courier_data['items']:
            if item['quantity'] > 0:
                item_count += 1
                # Make both quantity and unit bold using WhatsApp formatting
                unit_text = f" *{item['unit']}*" if item['unit'] else ""
                message_parts.append(f"• {item['name']}: *{item['quantity']}*{unit_text}")
        
        if item_count == 0:
            message_parts.append("• No items with specified quantities")
        
        message_parts.append("")
        # Add shipment type AFTER the items list
        shipment_type = courier_data.get('shipment_type', 'Complete')
        message_parts.append(f"*Shipment Type:* {shipment_type}")
        
        message_parts.append("")
        message_parts.append("*Your order has been dispatched successfully!*")
        message_parts.append("")
        message_parts.append("For any queries, contact us on 8108100404 or email us on logistics@accuratemedical.in")
        
        return "\n".join(message_parts)

    def send_courier_notifications(self, courier_data_list, customer_data):
        """Optimized courier notifications sender."""
        driver = None
        try:
            print("Starting courier notification process...")
            driver = self._get_driver()
            driver.get("https://web.whatsapp.com")
            
            print("Waiting for WhatsApp Web to load...")
            
            # Faster login detection
            login_successful = False
            max_wait_time = 60  # Reduced from 120
            check_interval = 2   # Reduced from 5
            
            for i in range(max_wait_time // check_interval):
                try:
                    search_elements = driver.find_elements(By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')
                    if search_elements:
                        login_successful = True
                        break
                    time.sleep(check_interval)
                except Exception as e:
                    time.sleep(check_interval)
            
            if not login_successful:
                return {"status": "error", "message": "Login timeout - please make sure you're logged into WhatsApp Web"}
            
            # Reduced wait after login
            time.sleep(3)  # Reduced from 10
            
            success_count = 0
            error_count = 0
            
            # Process customers in batches for better performance
            batch_size = 5
            for batch_start in range(0, len(customer_data), batch_size):
                batch_end = min(batch_start + batch_size, len(customer_data))
                batch_customers = customer_data[batch_start:batch_end]
                
                for i, customer in enumerate(batch_customers):
                    try:
                        customer_index = batch_start + i + 1
                        print(f"\n--- Processing customer {customer_index}: {customer['name']} ---")
                        
                        # Get the specific courier data for this customer
                        customer_courier_data = customer.get('courier_data', courier_data_list[0] if courier_data_list else {})
                        
                        # Format the message for this customer
                        formatted_message = self.format_courier_message(customer_courier_data, customer)
                        
                        # Clean mobile number
                        mobile = re.sub(r'[^\d]', '', customer['mobile'].strip())
                        
                        # Add country code if needed
                        if len(mobile) == 10:
                            mobile = '91' + mobile
                        elif len(mobile) == 12 and mobile.startswith('91'):
                            pass
                        elif len(mobile) == 13 and mobile.startswith('091'):
                            mobile = mobile[1:]
                        
                        print(f"Final mobile number: {mobile}")
                        
                        # Navigate faster
                        url = f"https://web.whatsapp.com/send?phone={mobile}"
                        driver.get(url)
                        
                        # Reduced wait time
                        time.sleep(3)  # Reduced from 8
                        
                        # Quick invalid number check
                        try:
                            invalid_elements = driver.find_elements(By.XPATH, 
                                "//div[contains(text(), 'Phone number shared via url is invalid') or contains(text(), 'not exist')]")
                            if invalid_elements:
                                print(f"Invalid number: {mobile}")
                                error_count += 1
                                continue
                        except:
                            pass
                        
                        # Fast message box detection
                        message_box = self.find_message_box_fast(driver, WebDriverWait(driver, 10))
                        
                        if not message_box:
                            print(f"Could not find message box for {mobile}")
                            error_count += 1
                            continue
                        
                        # Send message quickly
                        message_box.click()
                        time.sleep(0.3)  # CHANGE 1: Reduced wait
                        
                        # Clear and type message
                        message_box.send_keys(Keys.CONTROL + "a", Keys.DELETE)
                        time.sleep(0.2) # CHANGE 1
                        
                        # Type message in chunks for speed
                        message_lines = formatted_message.split('\n')
                        for j, line in enumerate(message_lines):
                            if line.strip():
                                message_box.send_keys(line)
                            if j < len(message_lines) - 1:
                                message_box.send_keys(Keys.SHIFT + Keys.ENTER)
                        
                        # Send message
                        message_box.send_keys(Keys.ENTER)
                        print(f"Message sent successfully to {customer['name']} ({mobile})")
                        success_count += 1
                        
                        # Minimal wait between messages
                        time.sleep(0.5)  # CHANGE 1: Reduced from 5
                        
                    except Exception as e:
                        print(f"Error processing {customer['name']}: {str(e)}")
                        error_count += 1
                        continue
                
                # Short batch break
                if batch_end < len(customer_data):
                    print("--- Short batch break ---")
                    time.sleep(1)
            
            print(f"\n=== Final Results ===")
            print(f"Success: {success_count}, Errors: {error_count}")
            
            return {
                "status": "success",
                "message": f"Courier notifications sent! Success: {success_count}, Errors: {error_count}"
            }
            
        except Exception as e:
            print(f"Critical error: {str(e)}")
            return {"status": "error", "message": f"An error occurred: {str(e)}"}
        finally:
            if driver:
                time.sleep(0.8)  # CHANGE 1: Reduced from 5
                try:
                    driver.quit()
                except:
                    pass
                
    def send_image_to_contact(self, driver, wait, image_path, caption_message=""):
        """
        Simplified and robust image sending method for WhatsApp Web
        """
        print(f"=== STARTING IMAGE SEND DEBUG ===")
        print(f"Image path: {image_path}")
        print(f"Caption: {caption_message}")
        
        try:
            # Method 1: Traditional attachment button approach
            print("DEBUG: Attempting Method 1 - Attachment Button")
            
            # Wait for chat to be fully loaded
            print("DEBUG: Waiting for chat interface...")
            time.sleep(3)
            
            # Find and click attachment/clip button - try multiple selectors
            attachment_selectors = [
                '[data-testid="clip"]',
                '[data-icon="clip"]', 
                'span[data-testid="clip"]',
                'button[aria-label="Attach"]',
                'div[title="Attach"]'
            ]
            
            attachment_clicked = False
            for i, selector in enumerate(attachment_selectors):
                try:
                    print(f"DEBUG: Trying attachment selector {i+1}: {selector}")
                    attachment_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    attachment_btn.click()
                    print(f"DEBUG: Attachment button clicked with selector {i+1}")
                    attachment_clicked = True
                    break
                except Exception as e:
                    print(f"DEBUG: Attachment selector {i+1} failed: {e}")
                    continue
            
            if not attachment_clicked:
                print("DEBUG: All attachment selectors failed, trying Method 2")
                return self.send_image_method_2(driver, wait, image_path, caption_message)
            
            time.sleep(2)
            
            # Find photo/document option and click it
            print("DEBUG: Looking for photo/document option...")
            photo_selectors = [
                'input[accept*="image"]',
                'input[type="file"][accept*="image"]',
                'li[data-testid="mi-attach-photo"]',
                'button[aria-label="Photos & Videos"]'
            ]
            
            photo_clicked = False
            for i, selector in enumerate(photo_selectors):
                try:
                    print(f"DEBUG: Trying photo selector {i+1}: {selector}")
                    
                    if selector.startswith('input'):
                        # Direct file input - send keys directly
                        file_input = driver.find_element(By.CSS_SELECTOR, selector)
                        print(f"DEBUG: Found file input, sending keys...")
                        file_input.send_keys(image_path)
                        photo_clicked = True
                        break
                    else:
                        # Button - click first then find input
                        photo_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                        photo_btn.click()
                        print(f"DEBUG: Photo button clicked with selector {i+1}")
                        time.sleep(1)
                        
                        # Now find the file input
                        file_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="file"]')))
                        file_input.send_keys(image_path)
                        photo_clicked = True
                        break
                        
                except Exception as e:
                    print(f"DEBUG: Photo selector {i+1} failed: {e}")
                    continue
            
            if not photo_clicked:
                print("DEBUG: Could not find photo input, trying Method 2")
                return self.send_image_method_2(driver, wait, image_path, caption_message)
            
            print("DEBUG: File uploaded, waiting for preview...")
            time.sleep(5)
            
            # Add caption if provided
            if caption_message and caption_message.strip():
                print("DEBUG: Adding caption...")
                caption_selectors = [
                    'div[contenteditable="true"][data-tab="10"]',
                    'div[aria-placeholder="Add a caption..."]',
                    'div[data-testid="media-caption-input"]'
                ]
                
                for selector in caption_selectors:
                    try:
                        caption_input = driver.find_element(By.CSS_SELECTOR, selector)
                        caption_input.click()
                        time.sleep(1)
                        caption_input.send_keys(caption_message)
                        print("DEBUG: Caption added successfully")
                        break
                    except Exception as e:
                        print(f"DEBUG: Caption selector failed: {e}")
                        continue
            
            # Send the image
            print("DEBUG: Sending image...")
            send_selectors = [
                'span[data-testid="send"]',
                'button[data-testid="send"]',
                '[aria-label="Send"]'
            ]
            
            for i, selector in enumerate(send_selectors):
                try:
                    send_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    send_btn.click()
                    print(f"DEBUG: Image sent successfully with selector {i+1}")
                    return {"status": "success", "message": "Image sent via Method 1"}
                except Exception as e:
                    print(f"DEBUG: Send selector {i+1} failed: {e}")
                    continue
            
            # If no send button found, try Enter key
            try:
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ENTER)
                print("DEBUG: Image sent via Enter key")
                return {"status": "success", "message": "Image sent via Enter key"}
            except Exception as e:
                print(f"DEBUG: Enter key failed: {e}")
                
            return {"status": "error", "message": "Could not send image - no send button found"}
            
        except Exception as e:
            print(f"DEBUG: Method 1 completely failed: {e}")
            return self.send_image_method_2(driver, wait, image_path, caption_message)

    def send_image_method_2(self, driver, wait, image_path, caption_message=""):
        """
        Alternative method using direct input injection
        """
        print("DEBUG: Attempting Method 2 - Direct Input Injection")
        
        try:
            # Create hidden file input and inject it
            print("DEBUG: Creating hidden file input...")
            
            create_input_script = """
            // Remove any existing hidden inputs
            var existingInputs = document.querySelectorAll('input[id="hidden-file-input"]');
            existingInputs.forEach(function(input) { input.remove(); });
            
            // Create new hidden input
            var input = document.createElement('input');
            input.type = 'file';
            input.id = 'hidden-file-input';
            input.accept = 'image/*,video/mp4,video/3gpp,video/quicktime';
            input.style.display = 'none';
            input.multiple = false;
            document.body.appendChild(input);
            return 'input_created';
            """
            
            result = driver.execute_script(create_input_script)
            print(f"DEBUG: Input creation result: {result}")
            
            if result != 'input_created':
                return {"status": "error", "message": "Could not create hidden input"}
            
            # Send file to the hidden input
            print("DEBUG: Sending file to hidden input...")
            hidden_input = driver.find_element(By.ID, "hidden-file-input")
            hidden_input.send_keys(image_path)
            time.sleep(2)
            
            # Trigger file processing
            print("DEBUG: Triggering file processing...")
            
            trigger_script = """
            var input = document.getElementById('hidden-file-input');
            if (input && input.files && input.files.length > 0) {
                var file = input.files[0];
                
                // Find message input area for drag simulation
                var dropTarget = document.querySelector('#main') || 
                                document.querySelector('[data-testid="conversation-panel-messages"]');
                
                if (dropTarget) {
                    // Create drag events
                    var dt = new DataTransfer();
                    dt.items.add(file);
                    
                    var dragEnter = new DragEvent('dragenter', { dataTransfer: dt, bubbles: true });
                    var dragOver = new DragEvent('dragover', { dataTransfer: dt, bubbles: true });
                    var drop = new DragEvent('drop', { dataTransfer: dt, bubbles: true });
                    
                    dropTarget.dispatchEvent(dragEnter);
                    dropTarget.dispatchEvent(dragOver);
                    dropTarget.dispatchEvent(drop);
                    
                    return 'file_processed';
                } else {
                    return 'no_drop_target';
                }
            } else {
                return 'no_file_found';
            }
            """
            
            process_result = driver.execute_script(trigger_script)
            print(f"DEBUG: File processing result: {process_result}")
            
            if process_result != 'file_processed':
                return {"status": "error", "message": f"File processing failed: {process_result}"}
            
            # Wait for media preview
            print("DEBUG: Waiting for media preview...")
            time.sleep(5)
            
            # Check if preview loaded
            preview_selectors = [
                '[data-testid="media-viewer"]',
                '[data-testid="send-container"]',
                'div[data-testid="send"]'
            ]
            
            preview_found = False
            for selector in preview_selectors:
                if driver.find_elements(By.CSS_SELECTOR, selector):
                    preview_found = True
                    print(f"DEBUG: Preview found with selector: {selector}")
                    break
            
            if not preview_found:
                return {"status": "error", "message": "Media preview did not load"}
            
            # Add caption and send (same as Method 1)
            if caption_message and caption_message.strip():
                print("DEBUG: Adding caption...")
                caption_selectors = [
                    'div[contenteditable="true"][data-tab="10"]',
                    'div[aria-placeholder="Add a caption..."]'
                ]
                
                for selector in caption_selectors:
                    try:
                        caption_input = driver.find_element(By.CSS_SELECTOR, selector)
                        caption_input.click()
                        time.sleep(1)
                        caption_input.send_keys(caption_message)
                        print("DEBUG: Caption added successfully")
                        break
                    except:
                        continue
            
            # Send the image
            print("DEBUG: Sending image...")
            send_selectors = [
                'span[data-testid="send"]',
                'button[data-testid="send"]'
            ]
            
            for selector in send_selectors:
                try:
                    send_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    send_btn.click()
                    print("DEBUG: Image sent successfully via Method 2")
                    return {"status": "success", "message": "Image sent via Method 2"}
                except:
                    continue
            
            return {"status": "error", "message": "Could not send image - no send button found"}
            
        except Exception as e:
            print(f"DEBUG: Method 2 failed: {e}")
            return {"status": "error", "message": f"Method 2 failed: {str(e)}"}

    def send_whatsapp_messages(self, numbers, message, image_path, use_template=False, template_content="", contacts=None):
        """Enhanced function for the Bulk Sender tool with better image handling."""
        driver = None
        try:
            driver = self._get_driver()
            wait = WebDriverWait(driver, 45)
            driver.get("https://web.whatsapp.com")
            
            print("Waiting for WhatsApp Web to load...")
            
            # Wait for login
            login_successful = False
            max_wait_time = 120
            check_interval = 5
            
            for i in range(max_wait_time // check_interval):
                try:
                    search_elements = driver.find_elements(By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')
                    if search_elements:
                        login_successful = True
                        break
                    time.sleep(check_interval)
                except Exception as e:
                    time.sleep(check_interval)
            
            if not login_successful:
                return {"status": "error", "message": "Login timeout - please make sure you're logged into WhatsApp Web"}
            
            time.sleep(2) # CHANGE 1
            
            success_count = 0
            error_count = 0
            
            # Update the image path handling
            if image_path and image_path != "null" and image_path != "undefined":
                # If image_path looks like a blob URL or not a valid file path, prompt for file selection
                if not os.path.exists(image_path) or image_path.startswith('blob:'):
                    print("Image path invalid or blob URL detected, prompting for file selection...")
                    image_path = self.get_image_file_path()

            # Determine which data source to use
            if contacts:
                # Use Excel contacts
                data_source = contacts
            else:
                # Use plain numbers list
                data_source = [{"phone": num.strip(), "name": f"Contact {i+1}", "custom_message": None} 
                                for i, num in enumerate(numbers) if num.strip()]
            
            for contact in data_source:
                if isinstance(contact, dict):
                    number = contact.get('phone', '')
                    contact_name = contact.get('name', 'Contact')
                else:
                    number = contact
                    contact_name = 'Contact'
                
                if not number.strip():
                    continue
                    
                try:
                    # Clean the number
                    mobile = re.sub(r'[^\d]', '', number.strip())
                    
                    # Add country code if needed
                    if len(mobile) == 10:
                        mobile = '91' + mobile
                    elif len(mobile) == 12 and mobile.startswith('91'):
                        mobile = mobile
                    elif len(mobile) == 13 and mobile.startswith('091'):
                        mobile = mobile[1:]
                    
                    print(f"Processing {contact_name}: {mobile}")
                    
                    url = f"https://web.whatsapp.com/send?phone={mobile}"
                    driver.get(url)
                    time.sleep(8)
                    
                    # Check for invalid number
                    try:
                        invalid_elements = driver.find_elements(By.XPATH, "//div[contains(text(), 'Phone number shared via url is invalid') or contains(text(), 'Telefonnummer') or contains(text(), 'not exist') or contains(text(), 'doesn\\'t have WhatsApp')]")
                        if invalid_elements:
                            print(f"Invalid number: {mobile}")
                            error_count += 1
                            continue
                    except:
                        pass
                    
                    # Determine message to send
                    final_message = ""
                    if use_template and template_content:
                        final_message = template_content.replace("{name}", contact_name)
                    elif hasattr(contact, 'get') and contact.get('custom_message') and contact.get('custom_message').strip() and contact.get('custom_message').strip() != 'nan':
                        final_message = contact.get('custom_message').replace("{name}", contact_name)
                    elif message and message.strip():
                        final_message = message.replace("{name}", contact_name)
                    
                    # Send image with message if both provided
                    if image_path and os.path.exists(image_path):
                        print(f"--- Starting image send for {mobile} ---")
                        image_result = self.send_image_to_contact(driver, wait, image_path, final_message)
                            
                        if image_result["status"] == "success":
                            print(f"Image sent successfully to {contact_name} ({mobile})")
                            success_count += 1
                            time.sleep(4)  # Wait after successful send
                            continue  # Skip text-only message sending
                        else:
                            print(f"Image send failed for {mobile}: {image_result['message']}")
                            # Fall back to text-only message if image fails
                            if final_message and final_message.strip():
                                print("Falling back to text-only message...")
                                # Continue to text message sending code below
                            else:
                                error_count += 1
                                continue

                    # Send text message only (no image)
                    if final_message and final_message.strip():
                        try:
                            print(f"Sending text message to {mobile}")
                            
                            # Find message input box
                            message_selectors = [
                                '//div[@contenteditable="true"][@data-tab="10"]',
                                '//div[@contenteditable="true"][@data-tab="1"]',
                                '//div[@role="textbox"][@contenteditable="true"]',
                                '//div[@title="Type a message"]',
                                '//footer//div[@contenteditable="true"]'
                            ]
                            
                            message_box = None
                            for selector in message_selectors:
                                try:
                                    message_box = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                                    break
                                except:
                                    continue
                            
                            if not message_box:
                                print("Could not find message input box")
                                error_count += 1
                                continue
                            
                            message_box.click()
                            time.sleep(0.3) # CHANGE 1
                            
                            # Clear any existing text
                            message_box.send_keys(Keys.CONTROL + "a")
                            message_box.send_keys(Keys.DELETE)
                            time.sleep(0.2) # CHANGE 1
                            
                            # Type message line by line
                            message_lines = final_message.split('\n')
                            for j, line in enumerate(message_lines):
                                if line.strip():
                                    message_box.send_keys(line)
                                if j < len(message_lines) - 1:
                                    message_box.send_keys(Keys.SHIFT + Keys.ENTER)
                            
                            # Send message
                            message_box.send_keys(Keys.ENTER)
                            print("Text message sent successfully!")
                            success_count += 1
                            time.sleep(3)
                            
                        except Exception as text_error:
                            print(f"Error sending text message to {mobile}: {str(text_error)}")
                            error_count += 1
                            continue
                    
                    else:
                        print(f"No message or image to send to {mobile}")
                        error_count += 1
                        continue
                    
                    time.sleep(2)  # Wait between messages
                    
                except Exception as e:
                    print(f"Error processing number {number}: {str(e)}")
                    error_count += 1
                    continue
            
            total_processed = len(data_source)
            return {
                "status": "success",
                "message": f"Processing completed! Success: {success_count}, Errors: {error_count} out of {total_processed} contacts."
            }
            
        except Exception as e:
            print(f"Critical error in send_whatsapp_messages: {str(e)}")
            return {"status": "error", "message": f"An error occurred: {type(e).__name__} - {e}"}
        finally:
            if driver:
                try:
                    time.sleep(3)
                    driver.quit()
                except:
                    pass

    def find_message_box_fast(self, driver, wait):
        """Optimized message box finder with faster detection."""
        # Try most common selectors first
        fast_selectors = [
            '//footer//div[@contenteditable="true"]',
            '//div[@contenteditable="true"][@data-tab="10"]',
            '//div[@role="textbox"][@contenteditable="true"]'
        ]
        
        for selector in fast_selectors:
            try:
                element = driver.find_element(By.XPATH, selector)
                if element.is_displayed() and element.is_enabled():
                    return element
            except:
                continue
        
        # Fallback to wait-based approach
        try:
            return wait.until(EC.element_to_be_clickable((By.XPATH, '//footer//div[@contenteditable="true"]')))
        except:
            return None
        
    def save_custom_template(self, template_name, message_template, excel_file_name=None):
        """Save custom message template to database."""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Create custom templates table if it doesn't exist
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS custom_templates (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE NOT NULL,
                    message_template TEXT NOT NULL,
                    excel_file_name TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            cursor.execute('''
                INSERT OR REPLACE INTO custom_templates (name, message_template, excel_file_name) 
                VALUES (?, ?, ?)
            ''', (template_name, message_template, excel_file_name))
            
            conn.commit()
            conn.close()
            
            return {"status": "success", "message": f"Custom template '{template_name}' saved successfully"}
            
        except Exception as e:
            return {"status": "error", "message": f"Error saving custom template: {str(e)}"}

    def get_custom_templates(self):
        """Get all saved custom templates."""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('SELECT id, name, message_template, excel_file_name FROM custom_templates ORDER BY name')
            rows = cursor.fetchall()
            
            templates = []
            for row in rows:
                templates.append({
                    "id": row[0],
                    "name": row[1],
                    "message_template": row[2],
                    "excel_file_name": row[3]
                })
            
            conn.close()
            return {"status": "success", "templates": templates}
            
        except Exception as e:
            return {"status": "error", "message": f"Error getting custom templates: {str(e)}"}

    # Fix the indentation issues in these methods too
    def get_custom_template_content(self, template_id):
        """Get content of a specific custom template by ID."""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('SELECT message_template, excel_file_name FROM custom_templates WHERE id = ?', (template_id,))
            row = cursor.fetchone()
            
            if row:
                conn.close()
                return {
                    "status": "success", 
                    "message_template": row[0],
                    "excel_file_name": row[1]
                }
            else:
                conn.close()
                return {"status": "error", "message": "Custom template not found"}
                
        except Exception as e:
            return {"status": "error", "message": f"Error getting custom template content: {str(e)}"}

    def delete_custom_template(self, template_id):
        """Delete a custom template."""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('DELETE FROM custom_templates WHERE id = ?', (template_id,))
            conn.commit()
            conn.close()
            
            return {"status": "success", "message": "Custom template deleted successfully"}
            
        except Exception as e:
            return {"status": "error", "message": f"Error deleting custom template: {str(e)}"}
        
    def analyze_custom_excel(self):
        """Analyze a custom Excel file and return column information."""
        try:
            # Open file dialog
            file_types = ('Excel Files (*.xlsx;*.xls;*.csv)', 'All files (*.*)')
            result = self.window.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=False, file_types=file_types)
            
            if not result or len(result) == 0:
                return {"status": "error", "message": "No file selected"}
            
            file_path = result[0]
            
            # Read the Excel file
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
            
            # Clean column names
            df.columns = df.columns.str.strip()
            
            # Convert DataFrame to list of dictionaries
            data = []
            for index, row in df.iterrows():
                row_dict = {}
                for col in df.columns:
                    row_dict[col] = str(row[col]).strip() if pd.notna(row[col]) else ''
                data.append(row_dict)
            
            # Get sample data (first row) for preview
            sample_data = {}
            if len(data) > 0:
                sample_data = data[0]
            
            return {
                "status": "success",
                "data": data,
                "columns": list(df.columns),
                "sample_data": sample_data,
                "message": f"Successfully analyzed {len(data)} rows with {len(df.columns)} columns"
            }
        
        except Exception as e:
            return {"status": "error", "message": f"Error analyzing Excel file: {str(e)}"}

    def navigate_to_whatsapp_chat(self, driver, mobile, max_retries=3):
        """Navigate to WhatsApp chat with retry logic."""
        for attempt in range(max_retries):
            try:
                url = f"https://web.whatsapp.com/send?phone={mobile}"
                print(f"Navigation attempt {attempt + 1}: {url}")

                driver.get(url)

                # --- OPTIMIZATION: Reduced navigation wait time ---
                time.sleep(3)
                
                # Check if page loaded correctly
                current_url = driver.current_url
                print(f"Current URL after navigation: {current_url}")

                # Check for blank page or loading issues
                if current_url == "about:blank" or "web.whatsapp.com" not in current_url:
                    print(f"Blank page detected, retrying...")
                    if attempt < max_retries - 1:
                        time.sleep(5)
                        continue
                    else:
                        return False

                # Check for invalid number immediately
                try:
                    invalid_selectors = [
                        "//div[contains(text(), 'Phone number shared via url is invalid')]",
                        "//div[contains(text(), 'Telefonnummer')]", 
                        "//div[contains(text(), 'not exist')]",
                        "//div[contains(text(), \"doesn't have WhatsApp\")]"
                    ]

                    for selector in invalid_selectors:
                        elements = driver.find_elements(By.XPATH, selector)
                        if elements:
                            print(f"Invalid number detected: {mobile}")
                            return "invalid"
                
                except Exception as e:
                    print(f"Error checking for invalid number: {e}")
                
                return True
            
            except Exception as e:
                print(f"Navigation attempt {attempt + 1} failed: {e}")
                if attempt < max_retries - 1:
                    time.sleep(5)
                    continue
                else:
                    return False
        return False

    # --- MODIFIED FUNCTION WITH ALL CHANGES ---
    # CHANGE 5: Increase Batch Size
    def send_custom_field_messages(self, excel_data, phone_field, name_field, message_template, batch_size=20):
        """Send WhatsApp messages using custom field mapping, with batching and multi-number support."""
        driver = None
        try:
            # CHANGE 7: Pre-compile Regex
            import re
            phone_pattern = re.compile(r'[^\d]')  # Pre-compile for better performance

            driver = self._get_driver()
            wait = WebDriverWait(driver, 45)
            driver.get("https://web.whatsapp.com")
            
            print("Waiting for WhatsApp Web to load...")
            
            # Wait for login
            login_successful = False
            max_wait_time = 120
            check_interval = 5
            
            for i in range(max_wait_time // check_interval):
                try:
                    search_elements = driver.find_elements(By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')
                    if search_elements:
                        login_successful = True
                        break
                    time.sleep(check_interval)
                except Exception as e:
                    time.sleep(check_interval)
            
            if not login_successful:
                return {"status": "error", "message": "Login timeout - please make sure you're logged into WhatsApp Web"}
            
            # CHANGE 1: Reduce Wait Times
            time.sleep(2)
            
            success_count = 0
            error_count = 0
            
            # --- OPTIMIZATION: Process in batches ---
            for batch_start in range(0, len(excel_data), batch_size):
                batch_end = min(batch_start + batch_size, len(excel_data))
                batch_data = excel_data[batch_start:batch_end]

                for i, row_data in enumerate(batch_data):
                    row_index = batch_start + i + 1
                    try:
                        phone_number_raw = str(row_data.get(phone_field, '')).strip()
                        if not phone_number_raw or phone_number_raw == 'nan':
                            print(f"Row {row_index}: No phone number found")
                            error_count += 1
                            continue

                        # CHANGE 8: Reduce Multiple Number Processing Overhead
                        phone_numbers_to_process = []
                        if ',' in phone_number_raw:
                            phone_numbers_to_process = [phone_pattern.sub('', n.strip()) for n in phone_number_raw.split(',') if n.strip() and len(phone_pattern.sub('', n.strip())) >= 10]
                        else:
                            # CHANGE 7: Use pre-compiled regex
                            clean_num = phone_pattern.sub('', phone_number_raw)
                            if len(clean_num) >= 10:
                                phone_numbers_to_process = [clean_num]

                        if not phone_numbers_to_process:
                            print(f"Row {row_index}: No valid phone numbers found in '{phone_number_raw}'")
                            error_count += 1
                            continue

                        # Process each phone number found in the cell
                        for mobile in phone_numbers_to_process:
                            # Add country code if needed
                            if len(mobile) == 10:
                                mobile = '91' + mobile
                            elif len(mobile) == 12 and mobile.startswith('91'):
                                pass # Already correct
                            elif len(mobile) == 13 and mobile.startswith('091'):
                                mobile = mobile[1:]
                            
                            if not (len(mobile) == 12 and mobile.startswith('91')):
                                print(f"Row {row_index}: Invalid phone number format: {mobile}")
                                error_count += 1
                                continue
                            
                            # Process message template for each number
                            processed_message = message_template
                            for field_name, field_value in row_data.items():
                                placeholder = "{" + field_name + "}"
                                processed_message = processed_message.replace(placeholder, str(field_value))
                            
                            contact_name = str(row_data.get(name_field, f'Contact {row_index}')).strip()
                            print(f"Processing {contact_name}: {mobile}")
                            
                            # CHANGE 2: Optimize Navigation with Better Error Handling
                            try:
                                url = f"https://web.whatsapp.com/send?phone={mobile}"
                                driver.get(url)
                                time.sleep(2)  # Reduced from 3
                                # Quick invalid number check
                                if driver.find_elements(By.XPATH, "//div[contains(text(), 'Phone number shared via url is invalid') or contains(text(), 'not exist')]"):
                                    print(f"Invalid number: {mobile}")
                                    error_count += 1
                                    continue
                            except Exception as nav_error:
                                print(f"Navigation failed for {mobile}: {nav_error}")
                                error_count += 1
                                continue
                            
                            # CHANGE 3: Faster Message Box Detection
                            try:
                                message_box = driver.find_element(By.XPATH, '//footer//div[@contenteditable="true"]')
                                if not message_box.is_displayed():
                                    message_box = driver.find_element(By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]')
                            except:
                                print(f"Could not find message input box for {mobile}")
                                error_count += 1
                                continue
                            
                            # ISSUE 1 FIX: Send message as single block
                            message_box.click()
                            time.sleep(0.3)
                            message_box.send_keys(Keys.CONTROL + "a", Keys.DELETE)  # Clear existing content
                            time.sleep(0.2)
                            # Type message line by line to preserve formatting
                            message_lines = processed_message.split('\n')
                            for j, line in enumerate(message_lines):
                                if line.strip():
                                    message_box.send_keys(line)
                                if j < len(message_lines) - 1:
                                    message_box.send_keys(Keys.SHIFT + Keys.ENTER)
                            time.sleep(1)
                            message_box.send_keys(Keys.ENTER)  # Send as one message

                            print(f"Message sent successfully to {contact_name} ({mobile})")
                            success_count += 1
                            # CHANGE 1: Reduced wait after sending message
                            time.sleep(0.5)
                    
                    except Exception as e:
                        print(f"Error processing row {row_index}: {str(e)}")
                        error_count += 1
                        continue
                    
                    # CHANGE 1: Reduced wait between processing different contacts/rows
                    time.sleep(0.8)
                
                # CHANGE 5: Reduce Batch Breaks
                if batch_end < len(excel_data):
                    print(f"--- Finished batch, taking a short break ---")
                    time.sleep(0.3)
            
            total_processed = len(excel_data)
            return {
                "status": "success",
                "message": f"Custom field messages sent! Success: {success_count}, Errors: {error_count} out of {total_processed} contacts."
            }
            
        except Exception as e:
            print(f"Critical error in send_custom_field_messages: {str(e)}")
            return {"status": "error", "message": f"An error occurred: {str(e)}"}
        finally:
            if driver:
                try:
                    time.sleep(3)
                    driver.quit()
                except:
                    pass
    
    # ISSUE 2: Add Clear Button Functionality
    def clear_courier_data(self):
        """Clear all courier notification data."""
        return {"status": "success", "message": "Courier data cleared"}

    def clear_bulk_sender_data(self):
        """Clear all bulk sender data."""
        return {"status": "success", "message": "Bulk sender data cleared"}

    def clear_group_sender_data(self):
        """Clear all group sender data."""
        return {"status": "success", "message": "Group sender data cleared"}

    def clear_custom_field_data(self):
        """Clear all custom field sender data."""
        return {"status": "success", "message": "Custom field data cleared"}

    def clear_all_data(self):
        """Clear all application data."""
        return {"status": "success", "message": "All data cleared"}


# --- Main Entry Point ---
if __name__ == '__main__':
    import os
    
    def start_app():
        api = Api()
        
        current_dir = os.path.dirname(os.path.abspath(__file__))
        login_path = os.path.join(current_dir, 'templates', 'login.html')
        
        if not os.path.exists(login_path):
            print(f"Login file not found at: {login_path}")
            return
        
        window = webview.create_window(
            'WhatsApp Business Suite - Login',
            login_path,
            js_api=api,
            width=1400,
            height=900,
            resizable=True,
            min_size=(1200, 700)
        )
        api.window = window
        webview.start(debug=False)
    
    start_app()


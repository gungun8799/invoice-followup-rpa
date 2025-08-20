from playwright.sync_api import sync_playwright
from typing import Optional
from datetime import datetime, timedelta
import pytz
import os


class TIMSAutomation:
    
    TIMS_URL = "https://tims.lotuss.com/TIMS/"
    USERNAME = "RPAac"
    PASSWORD = "Robot@082025"
    
    LOGIN_FRAME_NAME = "APPL"
    
    DEFAULT_TIMEOUT = 30000
    SHORT_TIMEOUT = 5000
    NAVIGATION_WAIT = 3000
    
    TIMSINFO_SELECTOR = "#_TIMSINFO"
    MAIN_PAGE_BUTTON_SELECTOR = "#_TIMSMSG > table > tbody > tr:nth-child(2) > td > input"
    DROPDOWN_MENU_SELECTOR = "#_MCELL5"
    REPORT_PAGE_SELECTOR = "#_MENU_1 > table > tbody > tr:nth-child(2) > td"
    DATE_INPUT_SELECTOR = "#Form > table > tbody > tr > td > table > tbody > tr > td > table.SC > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(1) > td:nth-child(6) > input[type=text]"
    EXPORT_BUTTON_SELECTOR = "#Form > table > tbody > tr > td > table > tbody > tr > td > table.BUTTONS > tbody > tr > td:nth-child(2) > input"
    
    def __init__(self, clear_cache=False):
        self.browser = None
        self.page = None
        self.context = None
        self.clear_cache = clear_cache
        self.need_refresh_and_retry = False
        
    def setup_browser(self, playwright):
        try:
            print("Launching Google Chrome browser with persistent user data...")
            import os
            download_path = os.path.join(os.getcwd(), 'downloads')
            if not os.path.exists(download_path):
                os.makedirs(download_path)
            
            # Create unique user data directory to avoid conflicts with running Chrome
            import time
            import random
            unique_id = f"{int(time.time())}_{random.randint(1000, 9999)}"
            user_data_dir = os.path.join(os.getcwd(), f'chrome_automation_{unique_id}')
            
            # Always create fresh directory to avoid lock conflicts
            if os.path.exists(user_data_dir):
                import shutil
                shutil.rmtree(user_data_dir)
                print(f"Removed existing directory: {user_data_dir}")
            
            os.makedirs(user_data_dir)
            print(f"Created fresh user data directory: {user_data_dir}")
            
            # Path to Google Chrome on macOS
            chrome_path = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
            
            # Browser args to ensure cache is enabled
            browser_args = [
                '--enable-features=NetworkService,NetworkServiceInProcess',
                '--disable-blink-features=AutomationControlled',
                '--disable-dev-shm-usage',
                '--no-sandbox',
                '--disable-web-security',
                '--disable-features=IsolateOrigins,site-per-process',
                '--allow-running-insecure-content',
                '--disable-setuid-sandbox',
                '--disable-site-isolation-trials'
            ]
            
            # Extra HTTP headers to mimic real browser
            extra_headers = {
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                'Accept-Language': 'en-US,en;q=0.9',
                'Accept-Encoding': 'gzip, deflate, br',
                'Cache-Control': 'max-age=0',
                'Upgrade-Insecure-Requests': '1'
            }
            
            # Launch with persistent context
            if not os.path.exists(chrome_path):
                print(f"Chrome not found at {chrome_path}")
                print("Using Chromium with persistent context and cache...")
                # Use persistent context for cache and cookies
                self.context = playwright.chromium.launch_persistent_context(
                    user_data_dir=user_data_dir,
                    headless=False,
                    accept_downloads=True,
                    ignore_https_errors=True,
                    viewport={'width': 1280, 'height': 720},
                    # Cache and storage settings
                    bypass_csp=True,
                    java_script_enabled=True,
                    # Browser arguments for cache
                    args=browser_args,
                    # Extra headers
                    extra_http_headers=extra_headers,
                    # User agent
                    user_agent='Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                    # Permissions
                    permissions=['geolocation', 'notifications']
                )
                self.browser = None  # Persistent context includes browser
            else:
                print(f"Using Google Chrome with persistent context and cache...")
                self.context = playwright.chromium.launch_persistent_context(
                    user_data_dir=user_data_dir,
                    headless=False,
                    executable_path=chrome_path,
                    accept_downloads=True,
                    ignore_https_errors=True,
                    viewport={'width': 1280, 'height': 720},
                    # Cache and storage settings
                    bypass_csp=True,
                    java_script_enabled=True,
                    # Browser arguments for cache
                    args=browser_args,
                    # Extra headers
                    extra_http_headers=extra_headers,
                    # User agent
                    user_agent='Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                    # Permissions
                    permissions=['geolocation', 'notifications']
                )
                self.browser = None  # Persistent context includes browser
            self.page = self.context.new_page()
            
            # Set additional page-level cache headers
            self.page.set_extra_http_headers({
                'Cache-Control': 'private, max-age=31536000',
                'Pragma': 'cache'
            })
            
            # Handle JavaScript dialogs automatically
            self.page.on("dialog", lambda dialog: self.handle_dialog(dialog))
            
            # Log network requests to debug cache usage
            self.page.on("request", lambda request: print(f"Request: {request.method} {request.url[:50]}...") if "tims" in request.url.lower() else None)
            self.page.on("response", lambda response: print(f"Response: {response.status} (from cache: {response.from_service_worker})") if "tims" in response.url.lower() else None)
            
            return True
        except Exception as e:
            print(f"Failed to setup browser: {e}")
            return False
    
    def handle_dialog(self, dialog):
        print(f"Dialog appeared: {dialog.type} - {dialog.message}")
        
        # Check for "Previous request is still in progress" alert
        if "Previous request is still in progress" in dialog.message:
            print("🔄 Detected 'Previous request is still in progress' - will refresh and retry")
            self.need_refresh_and_retry = True
        
        if dialog.type in ["alert", "confirm", "prompt"]:
            print("Accepting dialog...")
            dialog.accept()
        else:
            print("Dismissing dialog...")
            dialog.dismiss()
    
    def navigate_to_tims(self) -> bool:
        try:
            print("Navigating to TIMS website...")
            # First navigation with wait_until domcontentloaded to load initial resources
            self.page.goto(self.TIMS_URL, timeout=self.DEFAULT_TIMEOUT, wait_until='domcontentloaded')
            
            # Wait for network to stabilize and cache resources
            self.page.wait_for_load_state('networkidle', timeout=self.DEFAULT_TIMEOUT)
            
            # Give extra time for JavaScript to initialize and cache
            self.page.wait_for_timeout(3000)
            
            print("Page loaded and resources cached")
            return True
        except Exception as e:
            print(f"Failed to navigate to TIMS: {e}")
            return False
    
    def find_login_frame(self) -> Optional[object]:
        print("Looking for login frame...")
        frames = self.page.frames
        
        for frame in frames:
            if frame.name == self.LOGIN_FRAME_NAME:
                print("Found login frame")
                return frame
        
        print("Login frame not found!")
        return None
    
    def perform_login(self, login_frame) -> bool:
        try:
            print("Waiting for login frame to load...")
            login_frame.wait_for_load_state('networkidle', timeout=15000)
            
            print("Filling username...")
            username_field = login_frame.locator("input[type='text']").first
            username_field.fill(self.USERNAME)
            
            print("Filling password...")
            password_field = login_frame.locator("input[type='password']").first
            password_field.fill(self.PASSWORD)
            
            if not self._click_login_button(login_frame, password_field):
                return False
            
            print("Waiting for post-login navigation...")
            self.page.wait_for_timeout(self.SHORT_TIMEOUT)
            
            return self._verify_login_success()
            
        except Exception as e:
            print(f"Login failed: {e}")
            return False
    
    def _click_login_button(self, login_frame, password_field) -> bool:
        print("Submitting login form...")
        return self._submit_form_alternative(login_frame, password_field)
    
    def _submit_form_alternative(self, login_frame, password_field) -> bool:
        print("Submitting via Enter key in password field...")
        password_field.press("Enter")
        return True
    
    def _verify_login_success(self) -> bool:
        current_url = self.page.url
        print(f"Current URL: {current_url}")
        
        if "login" in current_url.lower():
            print("Still on login page - login might have failed")
            return False
        else:
            print("Successfully logged in!")
            return True
    
    def close_tims_info_popup(self) -> bool:
        print("Dismissing TIMS info popup with strategic clicks and dragging...")
        try:
            viewport_size = self.page.viewport_size
            if viewport_size:
                width = viewport_size['width']
                height = viewport_size['height']
                
                click_positions = [
                    (width // 2, height // 2),
                    (width // 4, height // 4),
                    (3 * width // 4, height // 4),
                    (width // 4, 3 * height // 4),
                    (3 * width // 4, 3 * height // 4),
                    (width // 2, height // 4),
                    (width // 2, 3 * height // 4),
                ]
                
                for x, y in click_positions:
                    try:
                        print(f"Clicking at position ({x}, {y})")
                        self.page.mouse.click(x, y)
                        self.page.wait_for_timeout(200)
                    except:
                        continue
                
                try:
                    print("Performing drag gesture across screen...")
                    self.page.mouse.move(100, 100)
                    self.page.mouse.down()
                    self.page.mouse.move(width - 100, height - 100)
                    self.page.mouse.up()
                    self.page.wait_for_timeout(500)
                    
                    print("Performing additional drag gestures...")
                    self.page.mouse.move(100, height // 2)
                    self.page.mouse.down()
                    self.page.mouse.move(width - 100, height // 2)
                    self.page.mouse.up()
                    self.page.wait_for_timeout(300)
                    
                    self.page.mouse.move(width // 2, 100)
                    self.page.mouse.down()
                    self.page.mouse.move(width // 2, height - 100)
                    self.page.mouse.up()
                    self.page.wait_for_timeout(300)
                    
                except:
                    pass
                
                print("TIMS info popup dismissed with strategic clicks and dragging")
                return True
            else:
                print("Could not get viewport size, using fallback coordinates")
                fallback_positions = [
                    (640, 360), (320, 180), (960, 180), 
                    (320, 540), (960, 540), (640, 180), (640, 540)
                ]
                for x, y in fallback_positions:
                    try:
                        print(f"Clicking at position ({x}, {y})")
                        self.page.mouse.click(x, y)
                        self.page.wait_for_timeout(200)
                    except:
                        continue
                
                try:
                    print("Performing fallback drag gestures...")
                    self.page.mouse.move(100, 100)
                    self.page.mouse.down()
                    self.page.mouse.move(1180, 620)
                    self.page.mouse.up()
                    self.page.wait_for_timeout(500)
                except:
                    pass
                    
                return True
                
        except Exception as e:
            print(f"Error closing TIMS info popup: {e}")
            return True

    def click_middle_screen(self) -> bool:
        print("Clicking in middle of screen to expose main page button...")
        try:
            print("Bringing browser window to front...")
            self.page.bring_to_front()
            
            self.page.wait_for_timeout(1000)
            
            viewport_size = self.page.viewport_size
            if viewport_size:
                middle_x = viewport_size['width'] // 2
                middle_y = viewport_size['height'] // 2
            else:
                middle_x = 640
                middle_y = 400
            
            print(f"Clicking at coordinates: ({middle_x}, {middle_y})")
            
            self.page.mouse.click(middle_x, middle_y)
            
            self.page.wait_for_timeout(2000)
            print("Middle screen click completed")
            return True
                
        except Exception as e:
            print(f"Error clicking middle of screen: {e}")
            return False
    
    def navigate_to_main_page(self) -> bool:
        print("Attempting to enter main page...")
        try:
            frames = self.page.frames
            element_found = False
            
            try:
                main_button = self.page.locator(self.MAIN_PAGE_BUTTON_SELECTOR)
                if main_button.count() > 0:
                    main_button.wait_for(timeout=self.SHORT_TIMEOUT)
                    main_button.click()
                    element_found = True
                    print("Clicked main page button (main page)")
            except:
                pass
            
            if not element_found:
                for frame in frames:
                    try:
                        main_button = frame.locator(self.MAIN_PAGE_BUTTON_SELECTOR)
                        if main_button.count() > 0:
                            main_button.wait_for(timeout=self.SHORT_TIMEOUT)
                            main_button.click()
                            element_found = True
                            print(f"Clicked main page button (frame: {frame.name})")
                            break
                    except:
                        continue
            
            if element_found:
                self.page.wait_for_timeout(self.NAVIGATION_WAIT)
                return True
            else:
                print("Main page button not found in any frame")
                return False
                
        except Exception as e:
            print(f"Error navigating to main page: {e}")
            return False
    
    def open_dropdown_menu(self) -> bool:
        print("Opening dropdown menu...")
        try:
            frames = self.page.frames
            element_found = False
            
            try:
                dropdown = self.page.locator(self.DROPDOWN_MENU_SELECTOR)
                if dropdown.count() > 0:
                    dropdown.wait_for(timeout=self.SHORT_TIMEOUT)
                    dropdown.hover()
                    element_found = True
                    print("Hovered over dropdown menu (main page)")
            except:
                pass
            
            if not element_found:
                for frame in frames:
                    try:
                        dropdown = frame.locator(self.DROPDOWN_MENU_SELECTOR)
                        if dropdown.count() > 0:
                            dropdown.wait_for(timeout=self.SHORT_TIMEOUT)
                            dropdown.hover()
                            element_found = True
                            print(f"Hovered over dropdown menu (frame: {frame.name})")
                            break
                    except:
                        continue
            
            if element_found:
                self.page.wait_for_timeout(2000)
                return True
            else:
                print("Dropdown menu not found in any frame")
                return False
                
        except Exception as e:
            print(f"Error opening dropdown menu: {e}")
            return False
    
    def navigate_to_report_page(self) -> bool:
        print("Navigating to report page...")
        try:
            frames = self.page.frames
            element_found = False
            
            try:
                report_button = self.page.locator(self.REPORT_PAGE_SELECTOR)
                if report_button.count() > 0:
                    report_button.wait_for(timeout=self.SHORT_TIMEOUT)
                    report_button.click()
                    element_found = True
                    print("Clicked report page button (main page)")
            except:
                pass
            
            if not element_found:
                for frame in frames:
                    try:
                        report_button = frame.locator(self.REPORT_PAGE_SELECTOR)
                        if report_button.count() > 0:
                            report_button.wait_for(timeout=self.SHORT_TIMEOUT)
                            report_button.click()
                            element_found = True
                            print(f"Clicked report page button (frame: {frame.name})")
                            break
                    except:
                        continue
            
            if element_found:
                self.page.wait_for_timeout(self.SHORT_TIMEOUT)
                print(f"Current URL after navigation: {self.page.url}")
                return True
            else:
                print("Report page button not found in any frame")
                return False
                
        except Exception as e:
            print(f"Error navigating to report page: {e}")
            return False
    
    def get_yesterday_date_bangkok(self) -> str:
        bangkok_tz = pytz.timezone('Asia/Bangkok')
        bangkok_now = datetime.now(bangkok_tz)
        yesterday = bangkok_now - timedelta(days=1)
        return yesterday.strftime('%Y-%m-%d')
    
    def fill_date_field(self) -> bool:
        print("Filling date field with yesterday's date...")
        try:
            yesterday_date = self.get_yesterday_date_bangkok()
            print(f"Yesterday's date (Bangkok time): {yesterday_date}")
            
            frames = self.page.frames
            element_found = False
            
            try:
                date_field = self.page.locator(self.DATE_INPUT_SELECTOR)
                if date_field.count() > 0:
                    date_field.wait_for(timeout=self.SHORT_TIMEOUT)
                    date_field.clear()
                    date_field.fill(yesterday_date)
                    # Press Enter or Tab to submit the date
                    date_field.press("Enter")
                    element_found = True
                    print("Filled date field (main page)")
            except:
                pass
            
            if not element_found:
                for frame in frames:
                    try:
                        date_field = frame.locator(self.DATE_INPUT_SELECTOR)
                        if date_field.count() > 0:
                            date_field.wait_for(timeout=self.SHORT_TIMEOUT)
                            date_field.clear()
                            date_field.fill(yesterday_date)
                            # Press Enter or Tab to submit the date
                            date_field.press("Enter")
                            element_found = True
                            print(f"Filled date field (frame: {frame.name})")
                            break
                    except:
                        continue
            
            if element_found:
                print("Waiting for date to be processed and page to fully initialize...")
                self.page.wait_for_timeout(10000)  # Increased wait time
                
                # Check page readiness before export
                print("Verifying page state for export readiness...")
                self.verify_page_readiness()
                
                return True
            else:
                print("Date input field not found in any frame")
                return False
                
        except Exception as e:
            print(f"Error filling date field: {e}")
            return False
    
    def verify_page_readiness(self) -> bool:
        """Verify the page is in the correct state for export operations"""
        print("Checking page readiness...")
        try:
            # Check minvokeInWindow function in APPL frame
            frames = self.page.frames
            for frame in frames:
                if frame.name == 'APPL':
                    # Check if function exists
                    func_exists = frame.evaluate("() => typeof minvokeInWindow === 'function'")
                    print(f"minvokeInWindow exists: {func_exists}")
                    
                    # Check variables
                    vars_check = frame.evaluate("() => ({ winELC: typeof winELC, MRC: typeof MRC })")
                    print(f"Variables: {vars_check}")
                    break
            
            # Basic page interactions
            self.page.mouse.move(500, 400)
            self.page.wait_for_timeout(1000)
            self.page.mouse.click(500, 400)
            self.page.wait_for_timeout(2000)
            
            return True
            
        except Exception as e:
            print(f"Error during readiness check: {e}")
            return False
    
    
    def click_export_button(self) -> bool:
        print("Attempting to click export button with network interception...")
        
        # Try export with network interception
        success = self._try_export_with_interception()
        
        # If we got the "Previous request in progress" alert, refresh and retry
        if self.need_refresh_and_retry:
            print("🔄 Refreshing page and re-navigating due to 'Previous request in progress'...")
            success = self._refresh_and_retry_export()
        
        return success
    
    def _try_export_with_interception(self) -> bool:
        """Try export with network request interception to capture exact parameters"""
        try:
            import os
            self.need_refresh_and_retry = False
            
            # Store intercepted requests
            intercepted_requests = []
            intercepted_responses = []
            
            def handle_request(request):
                if 'tims.lotuss.com' in request.url and (request.method == 'POST' or 'dispatcher' in request.url):
                    intercepted_requests.append({
                        'url': request.url,
                        'method': request.method,
                        'headers': dict(request.headers),
                        'post_data': request.post_data
                    })
                    print(f"🔍 Intercepted request: {request.method} {request.url}")
            
            def handle_response(response):
                if 'tims.lotuss.com' in response.url and response.request.method == 'POST':
                    intercepted_responses.append({
                        'url': response.url,
                        'status': response.status,
                        'headers': dict(response.headers),
                        'method': response.request.method
                    })
                    print(f"🔍 Intercepted response: {response.status} {response.url}")
            
            # Set up network interception
            self.page.on("request", handle_request)
            self.page.on("response", handle_response)
            
            # Also set up popup handler to catch window.open calls
            popup_detected = False
            popup_url = None
            
            # Store popup reference to prevent it from closing
            current_popup = None
            
            def handle_popup(popup):
                nonlocal popup_detected, popup_url, current_popup
                popup_detected = True
                current_popup = popup
                popup_url = popup.url
                print(f"🪟 Popup detected: {popup.url}")
                
                # Immediately return without processing to prevent popup from closing
                # We'll handle the popup processing separately
                return popup
            
            self.page.on("popup", handle_popup)
            
            print("🎯 Network interception and popup handling set up - attempting export click...")
            
            # Wait to ensure system is ready
            self.page.wait_for_timeout(3000)
            
            # Find and click export button
            frames = self.page.frames
            element_found = False
            
            for frame in frames:
                try:
                    export_button = frame.locator(self.EXPORT_BUTTON_SELECTOR)
                    if export_button.count() > 0:
                        print(f"Found export button in frame {frame.name}, clicking...")
                        
                        # First, analyze what the export button should do
                        print("🔍 Analyzing export button functionality...")
                        
                        try:
                            # Get detailed information about the export button and form
                            button_info = frame.evaluate("""() => {
                                const exportBtn = document.querySelector('input[name="A_E"]');
                                if (!exportBtn) return {error: 'Export button not found'};
                                
                                const form = exportBtn.closest('form');
                                if (!form) return {error: 'Form not found'};
                                
                                // Get button attributes
                                const btnAttrs = {};
                                for (let attr of exportBtn.attributes) {
                                    btnAttrs[attr.name] = attr.value;
                                }
                                
                                // Get form attributes
                                const formAttrs = {};
                                for (let attr of form.attributes) {
                                    formAttrs[attr.name] = attr.value;
                                }
                                
                                // Get onclick handler if any
                                const onclickHandler = exportBtn.onclick ? exportBtn.onclick.toString() : null;
                                
                                // Check if there's a submit function
                                const hasSubmit = typeof form.submit === 'function';
                                
                                // Get all form elements
                                const formElements = [];
                                for (let element of form.elements) {
                                    if (element.name) {
                                        formElements.push({
                                            name: element.name,
                                            value: element.value,
                                            type: element.type
                                        });
                                    }
                                }
                                
                                return {
                                    success: true,
                                    button: btnAttrs,
                                    form: formAttrs,
                                    onclick: onclickHandler,
                                    hasSubmit: hasSubmit,
                                    elements: formElements
                                };
                            }""")
                            
                            if button_info.get('success'):
                                print("📋 Export button analysis:")
                                print(f"   Button attributes: {button_info['button']}")
                                print(f"   Form attributes: {button_info['form']}")
                                print(f"   Onclick handler: {button_info['onclick']}")
                                print(f"   Form has submit: {button_info['hasSubmit']}")
                                print(f"   Form elements: {len(button_info['elements'])}")
                                
                                # Try to simulate the exact form submission
                                print("🔄 Attempting manual form submission...")
                                
                                # Method 1: Try direct form submission with A_E=E
                                try:
                                    submit_result = frame.evaluate("""() => {
                                        const form = document.querySelector('form');
                                        const exportBtn = document.querySelector('input[name="A_E"]');
                                        
                                        if (!form || !exportBtn) return {error: 'Form or button not found'};
                                        
                                        // Set the export action
                                        exportBtn.value = 'E';
                                        
                                        // Try to submit the form normally
                                        try {
                                            form.submit();
                                            return {success: true, method: 'form.submit()'};
                                        } catch (e) {
                                            return {error: 'form.submit() failed: ' + e.message};
                                        }
                                    }""")
                                    
                                    print(f"📤 Form submission result: {submit_result}")
                                    
                                    if submit_result.get('success'):
                                        print("✅ Form submitted successfully!")
                                    else:
                                        print(f"❌ Form submission failed: {submit_result.get('error')}")
                                        
                                        # Method 2: Try triggering the onclick event manually
                                        print("🔄 Trying onclick event simulation...")
                                        onclick_result = frame.evaluate("""() => {
                                            const exportBtn = document.querySelector('input[name="A_E"]');
                                            if (!exportBtn) return {error: 'Button not found'};
                                            
                                            try {
                                                // Trigger click event
                                                const clickEvent = new MouseEvent('click', {
                                                    view: window,
                                                    bubbles: true,
                                                    cancelable: true
                                                });
                                                exportBtn.dispatchEvent(clickEvent);
                                                return {success: true, method: 'dispatchEvent'};
                                            } catch (e) {
                                                return {error: 'Event dispatch failed: ' + e.message};
                                            }
                                        }""")
                                        
                                        print(f"🎯 Onclick simulation result: {onclick_result}")
                                        
                                except Exception as submit_error:
                                    print(f"❌ Manual submission failed: {submit_error}")
                                    
                            else:
                                print(f"❌ Button analysis failed: {button_info.get('error')}")
                                
                        except Exception as analysis_error:
                            print(f"❌ Analysis failed: {analysis_error}")
                        
                        # Now try the normal click as backup
                        try:
                            print("🔄 Trying normal button click as backup...")
                            export_button.click()
                            print("✅ Normal click completed")
                        except Exception as e1:
                            print(f"Normal click failed: {e1}")
                            
                            try:
                                print("Method 2: Force click...")
                                export_button.click(force=True)
                                print("✅ Force click completed")
                            except Exception as e2:
                                print(f"Force click failed: {e2}")
                                
                                try:
                                    print("Method 3: JavaScript click...")
                                    frame.evaluate("() => document.querySelector('input[name=\"A_E\"]').click()")
                                    print("✅ JavaScript click completed")
                                except Exception as e3:
                                    print(f"JavaScript click failed: {e3}")
                        
                        element_found = True
                        print("✅ Export button interaction completed!")
                        break
                        
                except Exception as e:
                    print(f"Frame export attempt failed: {e}")
                    continue
            
            if not element_found:
                print("❌ Export button not found")
                return False
            
            # Wait for popup to appear
            print("⏳ Waiting for popup to appear...")
            self.page.wait_for_timeout(3000)
            
            # Process popup if it appeared
            if popup_detected and current_popup:
                print("🔄 Processing detected popup...")
                try:
                    # Wait for initial load
                    current_popup.wait_for_load_state('domcontentloaded', timeout=5000)
                    initial_url = current_popup.url
                    print(f"🪟 Popup initial URL: {initial_url}")
                    
                    # If it's about:blank, wait for potential redirect
                    if initial_url == "about:blank":
                        print("⏳ Popup is about:blank, waiting for redirect...")
                        
                        # Wait and check multiple times for URL change
                        for attempt in range(8):  # Check for 8 seconds
                            current_popup.wait_for_timeout(1000)
                            current_url = current_popup.url
                            print(f"   Attempt {attempt + 1}: {current_url}")
                            
                            if current_url != "about:blank":
                                print(f"✅ Popup redirected to: {current_url}")
                                popup_url = current_url
                                
                                if "dispatcher" in current_url:
                                    print("🎉 SUCCESS: Dispatcher URL detected!")
                                    # Wait for download to complete
                                    current_popup.wait_for_timeout(8000)
                                    print("✅ Download should have completed")
                                    return self._check_download_files()
                                break
                        
                        # If still about:blank, try to extract form data and make direct request
                        if current_popup.url == "about:blank":
                            print("🔄 Still about:blank, trying to extract export data from main page...")
                            try:
                                # Get export form data from the main page
                                form_data = self.page.frames[0].evaluate("""() => {
                                    try {
                                        const form = document.querySelector('form');
                                        if (!form) return {error: 'No form found'};
                                        
                                        const formData = new FormData(form);
                                        const data = {};
                                        for (let [key, value] of formData.entries()) {
                                            data[key] = value;
                                        }
                                        data['A_E'] = 'E';  // Export action
                                        
                                        return {
                                            action: form.action || window.location.href,
                                            data: data,
                                            success: true
                                        };
                                    } catch (e) {
                                        return {error: e.message};
                                    }
                                }""")
                                
                                if form_data.get('success'):
                                    # Navigate popup to dispatcher with form data
                                    print("📍 Navigating popup to dispatcher with form data...")
                                    
                                    # Prepare POST data
                                    post_data = "&".join([f"{k}={v}" for k, v in form_data['data'].items()])
                                    
                                    # Navigate popup to the dispatcher URL
                                    dispatcher_url = "https://tims.lotuss.com/TIMS/dispatcher"
                                    current_popup.goto(f"{dispatcher_url}?{post_data}", timeout=10000)
                                    
                                    final_url = current_popup.url
                                    print(f"🏁 Popup navigation result: {final_url}")
                                    
                                    if "dispatcher" in final_url:
                                        print("🎉 SUCCESS: Popup navigated to dispatcher!")
                                        current_popup.wait_for_timeout(8000)
                                        return self._check_download_files()
                                        
                            except Exception as form_error:
                                print(f"❌ Form data extraction failed: {form_error}")
                    
                    elif "dispatcher" in initial_url:
                        print("✅ Dispatcher popup detected immediately")
                        popup_url = initial_url
                        current_popup.wait_for_timeout(8000)
                        return self._check_download_files()
                        
                except Exception as popup_error:
                    print(f"⚠️ Popup processing error: {popup_error}")
            
            # Wait for any remaining network activity
            print("⏳ Waiting for additional network requests...")
            self.page.wait_for_timeout(5000)
            
            # Check what happened
            print(f"📊 Results after processing:")
            print(f"   🔍 Intercepted requests: {len(intercepted_requests)}")
            print(f"   🪟 Popup detected: {popup_detected}")
            if popup_detected:
                print(f"   🔗 Popup URL: {popup_url}")
            
            # If popup was detected with dispatcher, that's likely success
            if popup_detected and popup_url and "dispatcher" in popup_url:
                print("🎉 SUCCESS: Dispatcher popup detected, checking for download files...")
                return self._check_download_files()
            
            # Process intercepted requests
            if intercepted_requests:
                print(f"📊 Captured {len(intercepted_requests)} requests")
                
                for i, req in enumerate(intercepted_requests):
                    print(f"\n📤 Request {i+1}:")
                    print(f"   URL: {req['url']}")
                    print(f"   Method: {req['method']}")
                    print(f"   Headers: {list(req['headers'].keys())}")
                    if req['post_data']:
                        print(f"   POST Data: {req['post_data'][:200]}...")
                    
                    # Try to reproduce this request manually
                    if req['method'] == 'POST' and req['post_data']:
                        print(f"🔄 Reproducing request {i+1}...")
                        try:
                            # Make the same POST request
                            response = self.page.request.post(
                                req['url'],
                                headers=req['headers'],
                                data=req['post_data']
                            )
                            
                            print(f"📥 Response status: {response.status}")
                            
                            if response.status == 200:
                                content_type = response.headers.get('content-type', '')
                                response_body = response.body()
                                
                                print(f"📄 Content-Type: {content_type}")
                                print(f"📊 Response size: {len(response_body)} bytes")
                                
                                # Check if it's Excel file or ZIP file (TIMS sends ZIP sometimes)
                                if (('application' in content_type and 'excel' in content_type) or 
                                    ('application/vnd.ms-excel' in content_type) or
                                    ('application/x-zip' in content_type) or
                                    ('application/zip' in content_type) or
                                    (len(response_body) > 10000)):
                                    
                                    # Determine file extension based on content type
                                    if 'zip' in content_type:
                                        file_ext = 'zip'
                                    elif 'excel' in content_type or response_body[:2] == b'\xd0\xcf':
                                        file_ext = 'xls'
                                    else:
                                        # Default to zip for large files
                                        file_ext = 'zip'
                                    
                                    # Generate filename with TIMS format
                                    from datetime import datetime
                                    bangkok_tz = pytz.timezone('Asia/Bangkok')
                                    now = datetime.now(bangkok_tz)
                                    timestamp = now.strftime('%y%m%d%H%M%S')
                                    suffix = str(now.microsecond)[:3]
                                    filename = f"{timestamp}.{suffix}.{file_ext}"
                                    
                                    # Create date-based folder structure
                                    downloads_base_dir = "downloads"
                                    bangkok_tz = pytz.timezone('Asia/Bangkok')
                                    today = datetime.now(bangkok_tz)
                                    date_folder_name = today.strftime('%d-%m-%Y')
                                    date_downloads_dir = os.path.join(downloads_base_dir, date_folder_name)
                                    
                                    # Create directories
                                    if not os.path.exists(downloads_base_dir):
                                        os.makedirs(downloads_base_dir)
                                    if not os.path.exists(date_downloads_dir):
                                        os.makedirs(date_downloads_dir)
                                        print(f"📁 Created date folder: {date_folder_name}")
                                    
                                    filepath = os.path.join(date_downloads_dir, filename)
                                    with open(filepath, 'wb') as f:
                                        f.write(response_body)
                                    
                                    print(f"💾 File saved: {filepath} ({len(response_body):,} bytes)")
                                    
                                    # If it's a ZIP file, try to extract it
                                    extracted_xls_path = None
                                    if file_ext == 'zip':
                                        try:
                                            import zipfile
                                            print("📦 Extracting ZIP file...")
                                            
                                            with zipfile.ZipFile(filepath, 'r') as zip_ref:
                                                extracted_files = zip_ref.namelist()
                                                print(f"   ZIP contains: {extracted_files}")
                                                
                                                # Extract to the same date folder
                                                zip_ref.extractall(date_downloads_dir)
                                                print(f"   ✅ Extracted {len(extracted_files)} files to {date_downloads_dir}")
                                                
                                                # Look for XLS files in extracted content
                                                xls_files = [f for f in extracted_files if f.endswith('.xls')]
                                                if xls_files:
                                                    print(f"   📊 Found XLS files: {xls_files}")
                                                    # Use the first XLS file found
                                                    extracted_xls_path = os.path.join(date_downloads_dir, xls_files[0])
                                                    
                                        except Exception as zip_error:
                                            print(f"⚠️ ZIP extraction failed: {zip_error}")
                                            # Still count as success since we got the file
                                    
                                    # File successfully downloaded and extracted
                                    if extracted_xls_path:
                                        print(f"✅ Excel file extracted: {extracted_xls_path}")
                                    else:
                                        print(f"✅ File saved: {filepath}")
                                    
                                    return True
                                    
                                elif 'text/html' in content_type:
                                    # Look for download links or redirects
                                    response_text = response_body.decode('utf-8', errors='ignore')
                                    print(f"📄 HTML response preview: {response_text[:300]}...")
                                    
                                    # Look for JavaScript window.open or location.href
                                    import re
                                    redirect_patterns = [
                                        r'window\.open\(["\']([^"\']*)["\']',
                                        r'location\.href\s*=\s*["\']([^"\']*)["\']',
                                        r'location\.replace\(["\']([^"\']*)["\']'
                                    ]
                                    
                                    for pattern in redirect_patterns:
                                        matches = re.findall(pattern, response_text, re.IGNORECASE)
                                        if matches:
                                            redirect_url = matches[0]
                                            print(f"🔗 Found redirect URL: {redirect_url}")
                                            
                                            # Follow the redirect
                                            if redirect_url.startswith('/'):
                                                redirect_url = 'https://tims.lotuss.com' + redirect_url
                                            elif not redirect_url.startswith('http'):
                                                redirect_url = 'https://tims.lotuss.com/TIMS/' + redirect_url
                                            
                                            print(f"📥 Following redirect: {redirect_url}")
                                            redirect_response = self.page.request.get(
                                                redirect_url,
                                                headers={
                                                    'Accept': '*/*',
                                                    'Referer': req['url'],
                                                    'User-Agent': req['headers'].get('User-Agent', 'Mozilla/5.0')
                                                }
                                            )
                                            
                                            if redirect_response.status == 200:
                                                redirect_body = redirect_response.body()
                                                redirect_content_type = redirect_response.headers.get('content-type', '')
                                                
                                                print(f"📊 Redirect response size: {len(redirect_body)} bytes")
                                                print(f"📄 Redirect content-type: {redirect_content_type}")
                                                
                                                if (len(redirect_body) > 1000 and 
                                                    (redirect_body[:2] == b'\xd0\xcf' or 'excel' in redirect_content_type)):
                                                    
                                                    # Save Excel file
                                                    from datetime import datetime
                                                    bangkok_tz = pytz.timezone('Asia/Bangkok')
                                                    now = datetime.now(bangkok_tz)
                                                    timestamp = now.strftime('%y%m%d%H%M%S')
                                                    suffix = str(now.microsecond)[:3]
                                                    filename = f"{timestamp}.{suffix}.xls"
                                                    
                                                    downloads_dir = "downloads"
                                                    if not os.path.exists(downloads_dir):
                                                        os.makedirs(downloads_dir)
                                                    
                                                    filepath = os.path.join(downloads_dir, filename)
                                                    with open(filepath, 'wb') as f:
                                                        f.write(redirect_body)
                                                    
                                                    print(f"💾 Excel file saved: {filepath} ({len(redirect_body):,} bytes)")
                                                    return True
                                            break
                            
                        except Exception as request_error:
                            print(f"❌ Failed to reproduce request: {request_error}")
                            continue
            
            # If no requests were intercepted and no popup, but button was clicked
            if element_found:
                print("⚠️ No network activity detected, but button was clicked. Checking for files...")
                if self._check_download_files():
                    return True
                    
                print("🔄 No files found, trying fallback approach...")
                return self._try_export_click_fallback()
            
            # Check for download files as fallback
            return self._check_download_files()
            
        except Exception as e:
            print(f"❌ Network interception failed: {e}")
            return self._try_export_click_fallback()

    def _check_download_files(self) -> bool:
        """Check for downloaded TIMS files in both local and system Downloads folders"""
        try:
            import os
            import re
            import time
            from datetime import datetime
            import pytz
            
            print("🔍 Checking for TIMS export files...")
            download_found = False
            
            tims_pattern = re.compile(r'^\d{12}\.\d{3}\.xls$')
            current_time = time.time()
            
            # Check system Downloads folder
            system_downloads = os.path.expanduser("~/Downloads")
            if os.path.exists(system_downloads):
                all_files = os.listdir(system_downloads)
                tims_files = [f for f in all_files if tims_pattern.match(f)]
                
                # Find files created in last 10 minutes
                recent_files = []
                for file in tims_files:
                    file_path = os.path.join(system_downloads, file)
                    mod_time = os.path.getmtime(file_path)
                    if current_time - mod_time < 600:  # 10 minutes
                        file_size = os.path.getsize(file_path)
                        recent_files.append((file, mod_time, file_size))
                
                if recent_files:
                    print(f"🎉 SUCCESS! Found {len(recent_files)} recent TIMS export file(s):")
                    for file, mod_time, size in sorted(recent_files, key=lambda x: x[1], reverse=True):
                        mod_datetime = datetime.fromtimestamp(mod_time)
                        print(f"  ✅ {file}")
                        print(f"     📁 Location: ~/Downloads/{file}")
                        print(f"     📊 Size: {size:,} bytes ({size/1024/1024:.1f} MB)")
                        print(f"     🕒 Created: {mod_datetime}")
                        download_found = True
                else:
                    print(f"📁 Found {len(tims_files)} TIMS files total, but none created recently")
            
            # Check local downloads folder with date-based structure
            downloads_base_dir = "downloads"
            if os.path.exists(downloads_base_dir):
                # Get today's date folder in Bangkok timezone
                bangkok_tz = pytz.timezone('Asia/Bangkok')
                today = datetime.now(bangkok_tz)
                date_folder_name = today.strftime('%d-%m-%Y')
                date_downloads_dir = os.path.join(downloads_base_dir, date_folder_name)
                
                print(f"📁 Checking date-based folder: {date_downloads_dir}")
                
                # Check today's date folder first
                if os.path.exists(date_downloads_dir):
                    local_files = os.listdir(date_downloads_dir)
                    if local_files:
                        print(f"📁 Today's date folder ({date_folder_name}): {len(local_files)} files")
                        for file in local_files:
                            if file.endswith('.xls'):
                                file_path = os.path.join(date_downloads_dir, file)
                                size = os.path.getsize(file_path)
                                mod_time = os.path.getmtime(file_path)
                                mod_datetime = datetime.fromtimestamp(mod_time)
                                print(f"  ✅ {file}")
                                print(f"     📁 Location: {date_downloads_dir}/{file}")
                                print(f"     📊 Size: {size:,} bytes ({size/1024/1024:.1f} MB)")
                                print(f"     🕒 Created: {mod_datetime}")
                                download_found = True
                    else:
                        print(f"📁 Today's date folder ({date_folder_name}) is empty")
                else:
                    print(f"📁 Today's date folder ({date_folder_name}) does not exist yet")
                
                # Also check the root downloads folder for any legacy files
                try:
                    root_files = [f for f in os.listdir(downloads_base_dir) 
                                 if os.path.isfile(os.path.join(downloads_base_dir, f)) and f.endswith('.xls')]
                    if root_files:
                        print(f"📁 Root downloads folder: {len(root_files)} .xls files")
                        for file in root_files:
                            file_path = os.path.join(downloads_base_dir, file)
                            size = os.path.getsize(file_path)
                            print(f"  - {file} ({size} bytes)")
                            download_found = True
                except:
                    pass
            
            return download_found
            
        except Exception as e:
            print(f"❌ Error checking downloads: {e}")
            return False

    def _try_export_click_fallback(self) -> bool:
        """Fallback method using original approach"""
        try:
            self.need_refresh_and_retry = False  # Reset flag
            
            print("🔄 Fallback: Simple export button click...")
            
            # Find and click export button with simple approach
            frames = self.page.frames
            element_found = False
            
            for frame in frames:
                try:
                    export_button = frame.locator(self.EXPORT_BUTTON_SELECTOR)
                    if export_button.count() > 0:
                        print(f"Found export button in frame {frame.name}, clicking...")
                        export_button.click()
                        element_found = True
                        print("✅ Export button clicked!")
                        break
                except Exception as e:
                    print(f"Frame export attempt failed: {e}")
                    continue
            
            if element_found:
                # Wait for any potential download
                self.page.wait_for_timeout(5000)
                return self._check_download_files()
            else:
                print("❌ Export button not found in fallback method")
                return False
                
        except Exception as e:
            print(f"Error clicking export button: {e}")
            return False
    
    def _refresh_and_retry_export(self) -> bool:
        """Refresh page and re-navigate to report page, then try export again"""
        try:
            print("🔄 Step 1: Refreshing the page...")
            self.page.reload(wait_until='networkidle', timeout=self.DEFAULT_TIMEOUT)
            self.page.wait_for_timeout(3000)
            
            print("🔄 Step 2: Re-entering main page...")
            if not self.navigate_to_main_page():
                print("❌ Failed to navigate to main page after refresh")
                return False
                
            print("🔄 Step 3: Re-opening dropdown menu...")
            if not self.open_dropdown_menu():
                print("⚠️ Could not open dropdown menu after refresh")
                
            print("🔄 Step 4: Re-navigating to report page...")
            if not self.navigate_to_report_page():
                print("❌ Failed to navigate to report page after refresh")
                return False
                
            print("🔄 Step 5: Re-filling date field...")
            if not self.fill_date_field():
                print("❌ Failed to fill date field after refresh")
                return False
                
            print("🔄 Step 6: Trying export again...")
            # Reset the flag before retry
            self.need_refresh_and_retry = False
            return self._try_export_with_interception()
            
        except Exception as e:
            print(f"❌ Error during refresh and retry: {e}")
            return False
    
    def cleanup(self):
        if self.context:
            self.context.close()
            print("Browser context closed")
        elif self.browser:
            self.browser.close()
            print("Browser closed")
    
    
    def run(self):
        with sync_playwright() as playwright:
            try:
                if not self.setup_browser(playwright):
                    return
                
                if not self.navigate_to_tims():
                    return
                
                login_frame = self.find_login_frame()
                if not login_frame:
                    return
                
                if not self.perform_login(login_frame):
                    return
                
                if not self.close_tims_info_popup():
                    print("Warning: Could not close TIMS info popup")
                
                if not self.click_middle_screen():
                    print("Warning: Could not click middle of screen")
                
                if not self.navigate_to_main_page():
                    print("Warning: Could not navigate to main page")
                
                if not self.open_dropdown_menu():
                    print("Warning: Could not open dropdown menu")
                
                if not self.navigate_to_report_page():
                    print("Warning: Could not navigate to report page")
                    
                if not self.fill_date_field():
                    print("Warning: Could not fill date field")
                    return
                
                # Use the improved export method
                export_success = self.click_export_button()
                
                # Print summary
                print("\n" + "="*60)
                print("📊 TIMS AUTOMATION SUMMARY")
                print("="*60)
                
                if export_success:
                    print("✅ TIMS Export: SUCCESS")
                    yesterday_date = self.get_yesterday_date_bangkok()
                    print(f"   📅 Date processed: {yesterday_date} (Yesterday, Bangkok timezone)")
                    print("   📁 Files saved to: downloads/")
                    
                    # Check what files we have
                    downloads_dir = "downloads"
                    if os.path.exists(downloads_dir):
                        files = [f for f in os.listdir(downloads_dir) if f.endswith('.xls') or f.endswith('.zip')]
                        files.sort(key=lambda x: os.path.getmtime(os.path.join(downloads_dir, x)), reverse=True)
                        
                        if files:
                            print("   📄 Downloaded files:")
                            for file in files[:3]:  # Show last 3 files
                                file_path = os.path.join(downloads_dir, file)
                                size = os.path.getsize(file_path)
                                print(f"      - {file} ({size:,} bytes)")
                    
                    # Check date-based download folder
                    bangkok_tz = pytz.timezone('Asia/Bangkok')
                    today = datetime.now(bangkok_tz)
                    date_folder_name = today.strftime('%d-%m-%Y')
                    date_downloads_dir = os.path.join(downloads_dir, date_folder_name)
                    
                    if os.path.exists(date_downloads_dir):
                        xls_files = [f for f in os.listdir(date_downloads_dir) if f.endswith('.xls')]
                        if xls_files:
                            print(f"\n📁 Downloaded Excel files in {date_folder_name} folder:")
                            for xls_file in xls_files:
                                file_path = os.path.join(date_downloads_dir, xls_file)
                                size = os.path.getsize(file_path)
                                print(f"      - {xls_file} ({size/1024/1024:.1f} MB)")
                    
                else:
                    print("❌ TIMS Export: FAILED")
                    print("   Please check the logs above for details")
                
                print("="*60)
                
                input("Press Enter to close browser...")
                
            except Exception as e:
                print(f"Unexpected error: {e}")
            finally:
                self.cleanup()


def main():
    # Set clear_cache=True for first run or if you want to start fresh
    # Set clear_cache=False to reuse existing cache and cookies
    automation = TIMSAutomation(clear_cache=False)
    automation.run()


if __name__ == "__main__":
    main()
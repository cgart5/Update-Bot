import sys
import os
import time
import shutil
import sqlite3
from A4GDB import A4GDB #DB class
from playwright.sync_api import sync_playwright
from config import username, password, webpage


class Bot:
    def __init__(self, url, filepath, progress_callback=None, summary_callback=None):
        shutil.rmtree("C:/POcodeBot/edge-profile", ignore_errors=True)
        
        #paths and url
        self.EDGE_PATH = r"C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"  # or wherever msedge.exe is
        self.url = url

        #df's
        # Layout
        self.data = []
        
        self.serviceArea = {}
        self.facility = {}
        self.route = {}
        self.zipCodes = []
        
        #usernames and passwords
        self.username = username
        self.password = password
        
        self.page = None
        self.context = None
        self.playwright = None

        #DB connection
        self.conn = sqlite3.connect('A4G.db')
        self.cur = self.conn.cursor()
        
        # Callback functions
        self.progress_callback = progress_callback
        self.summary_callback = summary_callback
        
        # Summary tracking
        self.summary = {
            'missing_service_areas': [],
            'missing_facilities': [],
            'missing_routes': [],
            'successful_service_areas': [],
            'successful_facilities': [],
            'successful_routes_count': 0,
            'routes_without_zip_codes': []
        }
        
        # Current context for tracking
        self.current_service_area = None
        self.current_facility = None

    def _send_progress_update(self, message, status="info"):
        """Send progress update to GUI if callback is available"""
        if self.progress_callback:
            try:
                self.progress_callback(message, status)
            except Exception as e:
                print(f"Error calling progress callback: {e}")
        else:
            print(message)

    def _send_service_area_complete(self, service_area, facilities_processed, routes_processed):
        """Send service area completion update to GUI"""
        if self.progress_callback:
            try:
                completion_data = {
                    'type': 'service_area_complete',
                    'service_area': service_area,
                    'facilities_processed': facilities_processed,
                    'routes_processed': routes_processed,
                    'current_summary': self.get_current_summary()
                }
                self.progress_callback(completion_data, "service_area_complete")
            except Exception as e:
                print(f"Error calling service area completion callback: {e}")

    def _send_final_summary(self):
        """Send final summary to GUI if callback is available"""
        if self.summary_callback:
            try:
                self.summary_callback(self.get_formatted_summary())
            except Exception as e:
                print(f"Error calling summary callback: {e}")
        else:
            self.print_summary()

    def get_current_summary(self):
        """Get current summary data as dictionary"""
        return {
            'missing_service_areas': self.summary['missing_service_areas'].copy(),
            'missing_facilities': self.summary['missing_facilities'].copy(),
            'missing_routes': self.summary['missing_routes'].copy(),
            'successful_service_areas': self.summary['successful_service_areas'].copy(),
            'successful_facilities': self.summary['successful_facilities'].copy(),
            'successful_routes_count': self.summary['successful_routes_count'],
            'routes_without_zip_codes': self.summary['routes_without_zip_codes'].copy()
        }

    def get_formatted_summary(self):
        """Get formatted summary as string for display"""
        summary_lines = []
        summary_lines.append("\n")
        summary_lines.append("="*80)
        summary_lines.append("                          EXECUTION SUMMARY")
        summary_lines.append("="*80)
        
        # Service Areas
        summary_lines.append("\n📍 SERVICE AREAS:")
        if self.summary['successful_service_areas']:
            summary_lines.append(f"  ✅ Successfully processed ({len(self.summary['successful_service_areas'])}): {', '.join(self.summary['successful_service_areas'])}")
        
        if self.summary['missing_service_areas']:
            summary_lines.append(f"  ❌ Not found ({len(self.summary['missing_service_areas'])}): {', '.join(self.summary['missing_service_areas'])}")
        
        if not self.summary['successful_service_areas'] and not self.summary['missing_service_areas']:
            summary_lines.append("  🟡 No service areas processed")
        
        # Facilities
        summary_lines.append("\n🏢 FACILITIES:")
        if self.summary['successful_facilities']:
            summary_lines.append(f"  ✅ Successfully processed ({len(self.summary['successful_facilities'])}): {', '.join(self.summary['successful_facilities'])}")
        
        if self.summary['missing_facilities']:
            summary_lines.append(f"  ❌ Not found ({len(self.summary['missing_facilities'])}): {', '.join(self.summary['missing_facilities'])}")
        
        if not self.summary['successful_facilities'] and not self.summary['missing_facilities']:
            summary_lines.append("  🟡 No facilities processed")
        
        # Routes
        summary_lines.append("\n🚛 ROUTES:")
        if self.summary['successful_routes_count'] > 0:
            summary_lines.append(f"  ✅ Successfully processed: {self.summary['successful_routes_count']} routes")
        
        if self.summary['missing_routes']:
            summary_lines.append(f"  ❌ Not found ({len(self.summary['missing_routes'])}): {', '.join(self.summary['missing_routes'])}")
        
        if self.summary['routes_without_zip_codes']:
            summary_lines.append(f"  📮 No zip codes found ({len(self.summary['routes_without_zip_codes'])}): {', '.join(self.summary['routes_without_zip_codes'])}")
        
        if self.summary['successful_routes_count'] == 0 and not self.summary['missing_routes'] and not self.summary['routes_without_zip_codes']:
            summary_lines.append("  🟡 No routes processed")
        
        # Overall Statistics
        summary_lines.append("\n📊 STATISTICS:")
        total_service_areas = len(self.summary['successful_service_areas']) + len(self.summary['missing_service_areas'])
        total_facilities = len(self.summary['successful_facilities']) + len(self.summary['missing_facilities'])
        total_routes = self.summary['successful_routes_count'] + len(self.summary['missing_routes']) + len(self.summary['routes_without_zip_codes'])
        
        if total_service_areas > 0:
            success_rate_sa = (len(self.summary['successful_service_areas']) / total_service_areas) * 100
            summary_lines.append(f"  Service Areas: {len(self.summary['successful_service_areas'])}/{total_service_areas} successful ({success_rate_sa:.1f}%)")
        
        if total_facilities > 0:
            success_rate_fac = (len(self.summary['successful_facilities']) / total_facilities) * 100
            summary_lines.append(f"  Facilities: {len(self.summary['successful_facilities'])}/{total_facilities} successful ({success_rate_fac:.1f}%)")
        
        if total_routes > 0:
            success_rate_routes = (self.summary['successful_routes_count'] / total_routes) * 100
            summary_lines.append(f"  Routes: {self.summary['successful_routes_count']}/{total_routes} successful ({success_rate_routes:.1f}%)")
        
        summary_lines.append("\n" + "="*80)
        summary_lines.append("                        END OF SUMMARY")
        summary_lines.append("="*80 + "\n")
        
        return "\n".join(summary_lines)

    # Initialize Playwright and keep it alive
    def start_browser(self):
        self._send_progress_update("Starting browser...", "info")
        self.playwright = sync_playwright().start()
        self.context = self.playwright.chromium.launch(
            #user_data_dir="C:/POcodeBot/edge-profile",
            executable_path=self.EDGE_PATH,
            headless=True,
            #slow_mo=500
        )
        self.page = self.context.new_page()
        self._send_progress_update("Browser started successfully", "success")
        return True

    # Loads initial page and logs in
    def load_page(self):
        try:
            self._send_progress_update("Loading webpage...", "info")
            self.page.goto(self.url)
            self.log_in()
            self._send_progress_update("Page loaded and logged in successfully", "success")
            return True
        except Exception as e:
            self._send_progress_update(f"Error loading page: {e}", "error")
            return False

    # Close browser and cleanup
    def close_browser(self):
        try:
            self._send_progress_update("Closing browser...", "info")
            if self.context:
                self.context.close()
            if self.playwright:
                self.playwright.stop()
            self._send_progress_update("Browser closed successfully", "success")
        except Exception as e:
            self._send_progress_update(f"Error closing browser: {e}", "error")

    # Logs in to A4G
    def log_in(self):
        try:
            self._send_progress_update("Logging in...", "info")
            # Add explicit wait for the login button
            self.page.wait_for_selector("id=sso-button", timeout=10000)
            self.page.click("id=sso-button")
            self.wait_for_title("Route Allocation Dashboard", 15)
            self.page.click()
            self._send_progress_update("Login successful", "success")
        except Exception as e:
            self._send_progress_update(f"Login failed: {e}", "error")
            
    def wait_for_title(self, expected_title, timeout=15):
        self._send_progress_update(f"Waiting for page title to become '{expected_title}'...", "info")
        start = time.time()
        while time.time() - start < timeout:
            current_title = self.page.title()
            if current_title == expected_title:
                self._send_progress_update("✅ Title matched!", "success")
                time.sleep(1.5)  # Reduced from 10 seconds
                return True
            time.sleep(1)  # wait 1 second before checking again
        self._send_progress_update(f"❌ Timeout reached. Current title: '{current_title}'", "error")
        return False
    
    def go_to_country(self, country):
        self._send_progress_update(f"Navigating to country: {country}", "info")
        try:
            # Wait for the service area dropdown to be available
            self.page.wait_for_selector("id=country-select", timeout=10000)
            
            # Click the dropdown
            self.page.click("id=country-select")
            
            # Wait for dropdown options to be visible
            self.page.wait_for_selector("span.mdc-list-item__primary-text", timeout=5000)
            
            # Try to find and click the specific service area
            service_area_locator = self.page.locator("span.mdc-list-item__primary-text", has_text=country)
            
            # Check if the service area option exists
            if service_area_locator.count() > 0:
                service_area_locator.click()
                self._send_progress_update(f"Successfully selected country: {country}", "success")
                return True
            else:
                # Print available options for debugging
                available_options = self.page.locator("span.mdc-list-item__primary-text").all_text_contents()
                self._send_progress_update(f"❌ Country '{country}' not found. Available: {available_options}", "error")
                return False
                
        except Exception as e:
            self._send_progress_update(f"Error navigating to country '{country}': {e}", "error")
            return False

    def go_to_serviceArea(self, serviceArea):
        self._send_progress_update(f"Navigating to service area: {serviceArea}", "info")
        try:
            # Wait for the service area dropdown to be available
            self.page.wait_for_selector("id=service-area-select", timeout=10000)
            
            # Click the dropdown
            self.page.click("id=service-area-select")
            
            # Wait for dropdown options to be visible
            self.page.wait_for_selector("span.mdc-list-item__primary-text", timeout=5000)
            
            # Try to find and click the specific service area
            service_area_locator = self.page.locator("span.mdc-list-item__primary-text", has_text=serviceArea)

            # Check if the service area option exists
            if service_area_locator.count() > 0:
                service_area_locator.click()
                self._send_progress_update(f"Successfully selected service area: {serviceArea}", "success")
                self.current_service_area = serviceArea
                self.summary['successful_service_areas'].append(serviceArea)
                return True
            else:
                # Print available options for debugging
                available_options = self.page.locator("span.mdc-list-item__primary-text").all_text_contents()
                self._send_progress_update(f"❌ Service area '{serviceArea}' not found. Available: {available_options}", "error")
                self.summary['missing_service_areas'].append(serviceArea)
                return False
                
        except Exception as e:
            self._send_progress_update(f"Error navigating to service area '{serviceArea}': {e}", "error")
            self.summary['missing_service_areas'].append(serviceArea)
            return False
        
    def go_to_facility(self, facility):
        self._send_progress_update(f"Navigating to facility: {facility}", "info")
        try:
            # Wait for the facility dropdown to be available
            self.page.wait_for_selector("id=facility-select", timeout=10000)
            
            # Click the dropdown
            self.page.click("id=facility-select")
            
            # Wait for dropdown options to be visible (any option, not specific facility)
            self.page.wait_for_selector("span.mdc-list-item__primary-text", timeout=10000)
            
            # Get all available options first for debugging
            available_options = self.page.locator("span.mdc-list-item__primary-text").all_text_contents()
            
            # Try to find the specific facility
            facility_locator = self.page.locator("span.mdc-list-item__primary-text", has_text=facility)
            
            # Check if the facility option exists
            facility_count = facility_locator.count()
            if facility_count > 0:
                facility_locator.click()
                
                # Optional: Wait for the selection to take effect
                try:
                    # Wait for dropdown to close (facility-select should be visible again)
                    self.page.wait_for_selector("id=facility-select", state="visible", timeout=10000)
                    self._send_progress_update(f"Successfully selected facility: {facility}", "success")
                    self.current_facility = facility
                    self.summary['successful_facilities'].append(facility)
                    return True
                except:
                    self._send_progress_update(f"⚠️ Facility '{facility}' was clicked but selection may not have completed", "warning")
                    self.current_facility = facility
                    self.summary['successful_facilities'].append(facility)
                    return True  # Still return True since we found and clicked the facility
                    
            else:
                cleaned_options = [opt.strip() for opt in available_options if opt.strip()]
                self._send_progress_update(f"❌ Facility '{facility}' not found. Available: {cleaned_options}", "error")
                self.summary['missing_facilities'].append(f"{facility} (in {self.current_service_area})")
                
                # Close the dropdown since we opened it but didn't select anything
                try:
                    self.page.keyboard.press("Escape")
                except:
                    pass
                    
                return False
                    
        except Exception as e:
            self._send_progress_update(f"❌ Error navigating to facility '{facility}': {e}", "error")
            self.summary['missing_facilities'].append(f"{facility} (in {self.current_service_area})")
            
            # Try to close any open dropdowns in case of error
            try:
                self.page.keyboard.press("Escape")
            except:
                pass
                
            return False

    def add_postal_codes(self, route):
        self.cur.execute(f"SELECT DISTINCT Zip FROM ZipCode WHERE Rt = '{route}';")
        zipCodes = self.cur.fetchall()

        # Try to select the route
        try:
            self.page.click("id=rules-dialog-route-select")
            
            # Wait for dropdown options to be visible
            self.page.wait_for_selector("span.mdc-list-item__primary-text", timeout=1000)
            
            # Check if the route exists in the dropdown
            route_locator = self.page.locator("span.mdc-list-item__primary-text", has_text=route)
            
            if route_locator.count() > 0:
                route_locator.click()
                
                # Continue with cycle selection
                self.page.click("id=rules-dialog-cycle-select")
                
                for label in ["A", "B"]:
                    try:
                        self.page.wait_for_selector(f"mat-option:has-text('{label}')", timeout=3000)
                        self.page.click(f"mat-option:has-text('{label}')")
                    except Exception as e:
                        self._send_progress_update(f"Could not select cycle {label}: {e}", "warning")
                
                self.page.mouse.click(0, 0)
                if zipCodes:
                    self._send_progress_update(f"Adding {len(zipCodes)} postal codes for route {route}", "info")
                    # Add all postal codes first
                    for zipCode in zipCodes:
                        self.page.fill("#chipInput", zipCode[0])
                        self.page.keyboard.press("Enter")
                else:
                    self._send_progress_update(f"No zip codes for route: {route}", "warning")
                    self.summary['routes_without_zip_codes'].append(f"{route} (at {self.current_facility} in {self.current_service_area})")
                    self.page.click("id=rules-dialog-cancel")
                    return False

                self.page.click("id=rules-dialog-submit")
                self._send_progress_update(f"✅ Successfully added postal codes for route: {route}", "success")
                self.summary['successful_routes_count'] += 1
                return True
                
            else:
                # Route not found - get available routes for debugging
                available_routes = self.page.locator("span.mdc-list-item__primary-text").all_text_contents()
                cleaned_routes = [route.strip() for route in available_routes if route.strip()]
                
                self._send_progress_update(f"❌ Route '{route}' not found. Available: {cleaned_routes}", "error")
                self.summary['missing_routes'].append(f"{route} (at {self.current_facility} in {self.current_service_area})")
                
                # Close the dialog since we can't proceed
                self.page.click("id=rules-dialog-cancel")
                return False
                
        except Exception as e:
            self._send_progress_update(f"❌ Error selecting route '{route}': {e}", "error")
            self.summary['missing_routes'].append(f"{route} (at {self.current_facility} in {self.current_service_area})")
            # Try to cancel the dialog to clean up
            try:
                self.page.click("id=rules-dialog-cancel")
            except:
                pass
            return False
    
    def _wait_for_stable_state(self, selector, timeout=5000):
        """Wait for element to be stable (present and not changing)"""
        try:
            self.page.wait_for_selector(selector, timeout=timeout)
            # Quick stability check - wait for element to remain unchanged
            time.sleep(0.1)  # Very short wait for stability
            return True
        except:
            return False
    
    def delete_postal_codes(self):
        try:
            # Quick check if there are any rows in the table body
            rows = self.page.locator("tr[role='row'][mat-row]")
            
            # Use a shorter timeout for the initial check
            try:
                self.page.wait_for_selector("tr[role='row'][mat-row]", timeout=1000)
                row_count = rows.count()
            except:
                row_count = 0

            if row_count == 0:
                self._send_progress_update("🟡 No postal code rows to delete.", "info")
                return

            self._send_progress_update("Deleting existing postal codes...", "info")
            
            # Wait for checkbox to appear with shorter timeout
            if self._wait_for_stable_state("label.checkbox-container", timeout=3000):
                # Click the "select all" checkbox
                self.page.click("label.checkbox-container", force=True)
                
                # Quick delete sequence
                self.page.click("id=delete-button")
                self.page.locator("id=dialog-yes").click()

                # Wait for and click submit button
                if self._wait_for_stable_state("id=submit-button", timeout=3000):
                    self.page.click("id=submit-button")

                    # Brief wait for panel to reload
                    time.sleep(0.5)
                    self.page.click("id=postal-code-rules-panel")
                    self._send_progress_update("✅ Deleted existing postal codes", "success")
                else:
                    self._send_progress_update("⚠️ Could not find submit button after delete", "warning")
            else:
                self._send_progress_update("⚠️ Could not find checkbox for selecting postal codes", "warning")

        except Exception as e:
            self._send_progress_update(f"❌ Error in delete_postal_codes: {e}", "error")

    def print_summary(self):
        """Keep original print_summary for backward compatibility"""
        summary_text = self.get_formatted_summary()
        print(summary_text)

    def load_data(self):
        self._send_progress_update("LOADING DATA", "info")

    def bot_main(self):
        # Fixed data structure building
        self._send_progress_update("Building data structure...", "info")
        
        # Start browser and load the page
        if not self.start_browser():
            self._send_progress_update("Failed to start browser", "error")
            return
            
        if not self.load_page():
            self._send_progress_update("Failed to load page", "error")
            self.close_browser()
            return
        self.cur.execute("SELECT DISTINCT CTRY FROM Service_Area;")
        countries = self.cur.fetchall()
        for country in countries:
            self.go_to_country(country[0])
            try:

                self.cur.execute(f"SELECT SA FROM Service_Area WHERE CTRY = '{country[0]}';")
                serviceAreas = self.cur.fetchall()
                
                total_service_areas = len(serviceAreas)
                self._send_progress_update(f"Found {total_service_areas} service areas to process in {country[0]}", "info")
                # Process each service area
                for i, sa in enumerate(serviceAreas, 1):
                    self._send_progress_update(f"\n--- Processing Service Area {i}/{total_service_areas}: {sa[0]} ---", "info")
                    
                    success = self.go_to_serviceArea(sa[0])
                    if not success:
                        self._send_progress_update(f"Failed to navigate to service area: {sa[0]}", "error")
                        # Send service area completion even if failed
                        self._send_service_area_complete(sa[0], 0, 0)
                        continue
                        
                    self.cur.execute(f"SELECT FAC FROM Facility WHERE SA='{sa[0]}';")
                    facilities = self.cur.fetchall()
                    facilities_processed = 0
                    routes_processed = 0
                    
                    for fa in facilities:
                        self._send_progress_update(f"Processing facility: {fa[0]}", "info")
                        success = self.go_to_facility(fa[0])
                        if not success:
                            self._send_progress_update(f"Failed to navigate to facility: {fa[0]}", "error")
                            continue
                            
                        facilities_processed += 1
                        
                        #go to postal code tab
                        self.page.click("id=postal-code-rules-panel")
                        self.delete_postal_codes()
                        self.cur.execute(f"""
                            SELECT Rt 
                            FROM Route r 
                            WHERE FAC = '{fa[0]}' 
                            AND EXISTS (
                                SELECT 1 
                                FROM ZipCode z 
                                WHERE z.Rt = r.Rt
                            );
                        """)
                        routes = self.cur.fetchall()
                        r_count = 0                    
                        for route in routes:
                            if r_count == 0:
                                time.sleep(.7)
                            self.page.click("id=add-button")
                            self._send_progress_update(f"  Processing Route: {route[0]}", "info")
                            # Add postal codes for the route
                            if r_count == 0:
                                time.sleep(.7)
                            if self.add_postal_codes(route[0]):
                                routes_processed += 1
                            r_count += 1

                        self.page.click("id=submit-button")
                    
                    # Send service area completion callback
                    self._send_service_area_complete(sa[0], facilities_processed, routes_processed)
                            
                
            except Exception as e:
                self._send_progress_update(f"Error in bot_main: {e}", "error")
            finally:
                self._send_progress_update(f"COUNTRY {country[0]} Completed")
        self._send_final_summary()
                
        # Keep browser open for debugging
        if not self.progress_callback:  # Only wait for input if no GUI callback
            input("Press Enter to close browser...")
        self.close_browser()

def run_bot(url, filepath, progress_callback=None, summary_callback=None):
    """Initialize the Bot with the URL and Excel of Data we are using
    
    Args:
        url: The webpage URL
        filepath: File path (currently unused)
        progress_callback: Function to call for progress updates
                          Signature: callback(message, status)
                          Status can be: "info", "success", "error", "warning", "service_area_complete"
        summary_callback: Function to call when execution is complete
                         Signature: callback(summary_text)
    """
    bot = Bot(url, filepath, progress_callback, summary_callback)
    bot.bot_main()

if __name__ == "__main__":
    filepath = ""
    url = webpage
    run_bot(url, filepath)

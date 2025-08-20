#!/usr/bin/env python3

import asyncio
import os
import time
from datetime import datetime
import pytz
from playwright.async_api import async_playwright


class PowerBIDebugger:
    def __init__(self):
        self.page = None
        self.browser = None
        self.context = None
        
        # Power BI configuration
        self.POWERBI_URL = "https://app.powerbi.com/groups/me/reports/558af259-9682-4f7a-8856-9d00044595ff/0d0fb1d63e8bbe02eea6?ctid=b08ac93a-5ec6-4492-bd45-dcd11a311661&experience=power-bi&bookmarkGuid=f8ed4e69-3539-4d68-bdea-586c98deef69"
        
        # Login credentials
        self.EMAIL = "john.pattanakarn@lotuss.com"
        self.PASSWORD = "Gofresh@0725-19"

    async def setup_browser(self):
        """Initialize browser"""
        playwright = await async_playwright().start()
        self.browser = await playwright.chromium.launch(headless=False)
        self.context = await self.browser.new_context(accept_downloads=True)
        self.page = await self.context.new_page()
        self.page.set_default_timeout(60000)
        print("✅ Browser initialized")
        return True

    async def login(self):
        """Handle Power BI login"""
        print("🔐 Starting login process...")
        
        # Navigate to Power BI
        await self.page.goto(self.POWERBI_URL)
        await self.page.wait_for_timeout(5000)
        
        # Check if login is needed
        current_url = self.page.url
        if 'singleSignOn' in current_url or 'login' in current_url:
            # Enter email
            try:
                email_field = self.page.get_by_placeholder("Enter email")
                await email_field.fill(self.EMAIL)
                print(f"✅ Email entered: {self.EMAIL}")
            except:
                email_field = await self.page.wait_for_selector("#email")
                await email_field.fill(self.EMAIL)
                print(f"✅ Email entered: {self.EMAIL}")
            
            # Click submit
            try:
                submit_btn = self.page.get_by_role("button", name="Submit")
                await submit_btn.click()
            except:
                submit_btn = await self.page.wait_for_selector("#submitBtn")
                await submit_btn.click()
            print("✅ Submit clicked")
            
            await self.page.wait_for_timeout(5000)
            
            # Enter password
            try:
                password_field = self.page.get_by_label("Password")
                await password_field.fill(self.PASSWORD)
            except:
                password_field = await self.page.wait_for_selector("#password")
                await password_field.fill(self.PASSWORD)
            print("✅ Password entered")
            
            # Submit password
            try:
                await password_field.press("Enter")
            except:
                pass
            
            await self.page.wait_for_timeout(5000)
            
            # Handle Continue button if present
            try:
                continue_btn = self.page.get_by_role("button", name="Continue")
                await continue_btn.click()
                print("✅ Continue clicked")
                await self.page.wait_for_timeout(2000)
                # Try second continue if needed
                try:
                    await continue_btn.click()
                except:
                    pass
            except:
                pass
            
            await self.page.wait_for_timeout(8000)
        
        print(f"✅ Login completed - URL: {self.page.url}")
        return True

    async def wait_for_report(self):
        """Wait for Power BI report to load"""
        print("⏳ Waiting for report to load...")
        
        # Wait for the page to load
        await self.page.wait_for_timeout(20000)  # 20 seconds to load
        
        print("✅ Report should be loaded")
        return True

    async def debug_page_elements(self):
        """Debug and find correct selectors"""
        print("\n🔍 DEBUGGING PAGE ELEMENTS")
        print("="*50)
        
        # Wait a bit more to ensure everything is loaded
        await self.page.wait_for_timeout(5000)
        
        # Get all visual containers
        print("🎯 Looking for visual containers...")
        visual_containers = await self.page.locator("visual-container").all()
        print(f"Found {len(visual_containers)} visual containers")
        
        # Look for tables/matrix visuals specifically
        print("\n📊 Looking for table/matrix visuals...")
        table_visuals = await self.page.locator("[class*='visual'], [class*='matrix'], [class*='table']").all()
        print(f"Found {len(table_visuals)} potential table visuals")
        
        # Look for three dots menu buttons
        print("\n⚙️ Looking for menu buttons...")
        menu_buttons = await self.page.locator("button[title*='More options'], button[aria-label*='More options'], .visual-header button").all()
        print(f"Found {len(menu_buttons)} potential menu buttons")
        
        # Try to find export-related elements
        print("\n📤 Looking for export elements...")
        export_elements = await self.page.locator("[title*='Export'], [aria-label*='Export'], [class*='export']").all()
        print(f"Found {len(export_elements)} potential export elements")
        
        # Get page content for manual inspection
        print("\n📝 Getting page structure...")
        
        # Save screenshot for visual inspection
        await self.page.screenshot(path="debug_powerbi.png", full_page=True)
        print("📸 Screenshot saved as debug_powerbi.png")
        
        # Wait and let user inspect the page
        print("\n⏸️ PAUSE FOR MANUAL INSPECTION")
        print("Please inspect the Power BI page in the browser window.")
        print("Look for:")
        print("1. The table/matrix visual you want to export")
        print("2. The three dots menu button on that visual")
        print("3. Any export options in the menu")
        print("\nPress Enter when ready to continue with interactive selection...")
        input()
        
        return True

    async def interactive_element_selection(self):
        """Help user find the right elements interactively"""
        print("\n🎯 INTERACTIVE ELEMENT SELECTION")
        print("="*50)
        
        # Try to click on visual containers one by one
        visual_containers = await self.page.locator("visual-container").all()
        
        for i, container in enumerate(visual_containers):
            print(f"\n🔍 Testing visual container {i+1}/{len(visual_containers)}")
            
            try:
                # Highlight the container
                await container.hover()
                await self.page.wait_for_timeout(1000)
                
                # Try to click it
                await container.click()
                await self.page.wait_for_timeout(2000)
                
                # Look for menu button after clicking
                menu_buttons = await container.locator("button[title*='More options'], button[aria-label*='More options'], .visual-header button").all()
                
                if len(menu_buttons) > 0:
                    print(f"✅ Found {len(menu_buttons)} menu button(s) in container {i+1}")
                    
                    # Try clicking the menu button
                    for j, btn in enumerate(menu_buttons):
                        try:
                            print(f"   🔘 Clicking menu button {j+1}")
                            await btn.click()
                            await self.page.wait_for_timeout(3000)
                            
                            # Look for export options
                            export_options = await self.page.locator("text=Export, [title*='Export'], [aria-label*='Export']").all()
                            
                            if len(export_options) > 0:
                                print(f"   🎉 FOUND EXPORT OPTIONS! Container {i+1}, Button {j+1}")
                                print(f"   📍 Container selector: visual-container:nth-child({i+1})")
                                
                                # Get the exact selectors
                                container_selector = f"visual-container:nth-child({i+1})"
                                button_selector = await btn.get_attribute("class")
                                
                                print(f"   🎯 Use these selectors:")
                                print(f"   - Container: {container_selector}")
                                print(f"   - Menu button classes: {button_selector}")
                                
                                # Try to click export
                                try:
                                    export_option = export_options[0]
                                    await export_option.click()
                                    print("   ✅ Export option clicked!")
                                    
                                    # Wait for export dialog
                                    await self.page.wait_for_timeout(5000)
                                    
                                    # Look for final export button
                                    final_buttons = await self.page.locator("button:has-text('Export'), button[type='submit'], .exportButton").all()
                                    
                                    if len(final_buttons) > 0:
                                        print(f"   🚀 Found {len(final_buttons)} final export button(s)")
                                        
                                        # Click the final export button
                                        await final_buttons[0].click()
                                        print("   🎉 EXPORT PROCESS COMPLETED!")
                                        
                                        return True
                                    
                                except Exception as e:
                                    print(f"   ⚠️ Export click failed: {e}")
                                
                                return True
                            
                            # Close menu if no export found
                            await self.page.keyboard.press("Escape")
                            await self.page.wait_for_timeout(1000)
                            
                        except Exception as e:
                            print(f"   ⚠️ Menu button {j+1} click failed: {e}")
                
            except Exception as e:
                print(f"⚠️ Container {i+1} interaction failed: {e}")
        
        print("\n❌ No working export path found through automation")
        print("You may need to manually inspect the page and update the selectors")
        return False

    async def run(self):
        """Main execution flow"""
        try:
            print("="*50)
            print("🤖 Power BI Debug Tool")
            print("="*50)
            
            # Setup
            await self.setup_browser()
            
            # Login
            await self.login()
            
            # Wait for report
            await self.wait_for_report()
            
            # Debug elements
            await self.debug_page_elements()
            
            # Interactive selection
            await self.interactive_element_selection()
            
            print("\n✅ Debug session completed")
            print("Check the console output and debug_powerbi.png for insights")
            
        except Exception as e:
            print(f"❌ Error: {e}")
        finally:
            print("\nKeeping browser open for manual inspection...")
            print("Press Enter to close...")
            input()
            if self.browser:
                await self.browser.close()


async def main():
    debugger = PowerBIDebugger()
    await debugger.run()


if __name__ == "__main__":
    asyncio.run(main())
#!/usr/bin/env python3

import asyncio
import os
import time
from datetime import datetime
import pytz
from playwright.async_api import async_playwright


class PowerBIAutoFix:
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
        await self.page.wait_for_timeout(20000)
        print("✅ Report should be loaded")
        return True

    async def find_and_export_table(self):
        """Find table visual and export it using modern selectors"""
        print("🔍 Looking for table visual to export...")
        
        # Wait for visuals to load
        await self.page.wait_for_timeout(5000)
        
        # Strategy 1: Look for table/matrix visuals with data
        print("📊 Strategy 1: Looking for table/matrix visuals...")
        
        # Modern Power BI uses these selectors
        table_selectors = [
            "visual-container[ng-class*='table']",
            "visual-container[ng-class*='matrix']", 
            "visual-container .visual-content .visual-body table",
            "visual-container .tableEx",
            "visual-container .pivotTable"
        ]
        
        for selector in table_selectors:
            try:
                tables = await self.page.locator(selector).all()
                print(f"   Found {len(tables)} elements with selector: {selector}")
                
                if len(tables) > 0:
                    # Try to interact with the first table
                    table = tables[0]
                    
                    # Right-click to open context menu
                    print("   🖱️ Right-clicking on table...")
                    await table.click(button="right")
                    await self.page.wait_for_timeout(2000)
                    
                    # Look for export option in context menu
                    export_options = await self.page.locator("text=/Export.*data/i, [title*='Export'], [aria-label*='Export']").all()
                    
                    if len(export_options) > 0:
                        print("   ✅ Found export option in context menu!")
                        await export_options[0].click()
                        await self.page.wait_for_timeout(3000)
                        
                        # Look for final export button
                        final_export = await self.page.locator("button:has-text('Export'), .exportButton, button[type='submit']").first
                        if await final_export.is_visible():
                            await final_export.click()
                            print("   🎉 Export initiated!")
                            return True
                    
                    # If right-click didn't work, try hovering and looking for menu
                    print("   🖱️ Hovering to find menu button...")
                    await table.hover()
                    await self.page.wait_for_timeout(2000)
                    
                    # Look for menu button that appears on hover
                    menu_buttons = await self.page.locator("button[title*='More'], button[aria-label*='More'], .visual-header button").all()
                    
                    for btn in menu_buttons:
                        if await btn.is_visible():
                            print("   🔘 Clicking menu button...")
                            await btn.click()
                            await self.page.wait_for_timeout(2000)
                            
                            # Look for export in dropdown
                            export_items = await self.page.locator("text=/Export/i, [title*='Export']").all()
                            
                            if len(export_items) > 0:
                                print("   ✅ Found export in dropdown!")
                                await export_items[0].click()
                                await self.page.wait_for_timeout(3000)
                                
                                # Click final export button
                                final_export = await self.page.locator("button:has-text('Export'), .exportButton").first
                                if await final_export.is_visible():
                                    await final_export.click()
                                    print("   🎉 Export initiated!")
                                    return True
                            
                            # Close menu if export not found
                            await self.page.keyboard.press("Escape")
                            await self.page.wait_for_timeout(1000)
                    
            except Exception as e:
                print(f"   ⚠️ Selector {selector} failed: {e}")
                continue
        
        # Strategy 2: Use more generic approach
        print("📊 Strategy 2: Generic visual container approach...")
        
        try:
            # Get all visual containers
            visual_containers = await self.page.locator("visual-container").all()
            print(f"   Found {len(visual_containers)} visual containers")
            
            # Try each container
            for i, container in enumerate(visual_containers[:10]):  # Limit to first 10
                try:
                    print(f"   Testing container {i+1}...")
                    
                    # Hover over container
                    await container.hover()
                    await self.page.wait_for_timeout(1000)
                    
                    # Look for menu button
                    menu_btn = container.locator("button[title*='More'], button[aria-label*='More']").first
                    
                    if await menu_btn.is_visible():
                        print(f"   🔘 Found menu in container {i+1}")
                        await menu_btn.click()
                        await self.page.wait_for_timeout(2000)
                        
                        # Look for export
                        export_option = self.page.locator("text=/Export.*data/i").first
                        
                        if await export_option.is_visible():
                            print(f"   ✅ Found export in container {i+1}!")
                            await export_option.click()
                            await self.page.wait_for_timeout(3000)
                            
                            # Final export button
                            final_btn = self.page.locator("button:has-text('Export'), .exportButton").first
                            if await final_btn.is_visible():
                                await final_btn.click()
                                print("   🎉 Export successful!")
                                return True
                        
                        # Close menu
                        await self.page.keyboard.press("Escape")
                        await self.page.wait_for_timeout(1000)
                        
                except Exception as e:
                    print(f"   ⚠️ Container {i+1} failed: {e}")
                    continue
                    
        except Exception as e:
            print(f"   ❌ Generic approach failed: {e}")
        
        print("❌ Could not find exportable table visual")
        return False

    async def handle_download(self):
        """Wait for and handle the download"""
        print("⬇️ Waiting for download...")
        
        try:
            # Wait for download to start
            async with self.page.expect_download(timeout=30000) as download_info:
                await self.page.wait_for_timeout(2000)
            
            download = download_info.value
            
            # Create downloads folder structure
            downloads_base = "downloads"
            if not os.path.exists(downloads_base):
                os.makedirs(downloads_base)
            
            # Create PBI folder with current date
            bangkok_tz = pytz.timezone('Asia/Bangkok')
            today = datetime.now(bangkok_tz)
            date_str = today.strftime('%d-%m-%Y')
            pbi_folder = os.path.join(downloads_base, f"PBI {date_str}")
            
            if not os.path.exists(pbi_folder):
                os.makedirs(pbi_folder)
                print(f"📁 Created folder: {pbi_folder}")
            
            # Save the download
            filename = download.suggested_filename or f"PowerBI_Export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            filepath = os.path.join(pbi_folder, filename)
            await download.save_as(filepath)
            
            print(f"✅ File saved: {filepath}")
            return True
            
        except Exception as e:
            print(f"⚠️ Download handling failed: {e}")
            print("Check if file was downloaded to your Downloads folder")
            return False

    async def run(self):
        """Main execution flow"""
        try:
            print("="*50)
            print("🤖 Power BI Auto-Fix Export")
            print("="*50)
            
            # Setup
            await self.setup_browser()
            
            # Login
            await self.login()
            
            # Wait for report
            await self.wait_for_report()
            
            # Find and export table
            success = await self.find_and_export_table()
            
            if success:
                # Handle download
                download_success = await self.handle_download()
                if download_success:
                    print("🎉 Export completed successfully!")
                else:
                    print("⚠️ Export triggered but download handling failed")
            else:
                print("❌ Could not find or export table")
            
            print("="*50)
            
            # Keep browser open for 30 seconds to see results
            print("🔍 Keeping browser open for 30 seconds...")
            await self.page.wait_for_timeout(30000)
            
        except Exception as e:
            print(f"❌ Error: {e}")
        finally:
            if self.browser:
                await self.browser.close()


async def main():
    auto_fix = PowerBIAutoFix()
    await auto_fix.run()


if __name__ == "__main__":
    asyncio.run(main())
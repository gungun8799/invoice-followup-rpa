#!/usr/bin/env python3

import asyncio
import os
import time
from datetime import datetime
import pytz
from playwright.async_api import async_playwright


class PowerBIExportAutomation:
    def __init__(self):
        self.page = None
        self.browser = None
        self.context = None
        
        # Power BI configuration
        self.POWERBI_URL = "https://app.powerbi.com/groups/me/reports/558af259-9682-4f7a-8856-9d00044595ff/0d0fb1d63e8bbe02eea6?ctid=b08ac93a-5ec6-4492-bd45-dcd11a311661&experience=power-bi&bookmarkGuid=f8ed4e69-3539-4d68-bdea-586c98deef69"
        
        # Login credentials
        self.EMAIL = "john.pattanakarn@lotuss.com"
        self.PASSWORD = "Gofresh@0725-19"
        
        # Report selectors - exact elements to click
        self.REPORT_BLOCK_SELECTOR = "#pvExplorationHost > div > div > exploration > div > explore-canvas > div > div.canvasFlexBox > div > div.displayArea.disableAnimations.fitToPage > div.visualContainerHost.visualContainerOutOfFocus > visual-container-repeat > visual-container:nth-child(16) > transform > div > div.visualContent.noPopOutBar > div > div"
        self.MORE_OPTIONS_SELECTOR = "#pvExplorationHost > div > div > exploration > div > explore-canvas > div > div.canvasFlexBox > div > div.displayArea.disableAnimations.fitToPage > div.visualContainerHost.visualContainerOutOfFocus > visual-container-repeat > visual-container:nth-child(16) > transform > div > visual-container-header > div > div > div > visual-container-options-menu > visual-header-item-container > div > button"
        self.EXPORT_DATA_SELECTOR = "#\\37"

    async def setup_browser(self):
        """Initialize browser"""
        playwright = await async_playwright().start()
        self.browser = await playwright.chromium.launch(headless=False)
        self.context = await self.browser.new_context(accept_downloads=True)
        self.page = await self.context.new_page()
        self.page.set_default_timeout(30000)  # Reduced timeout
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
        await self.page.wait_for_timeout(15000)  # Wait 15 seconds
        print("✅ Report should be loaded")
        return True

    async def export_report(self):
        """Click through the export process"""
        print("📊 Starting export process...")
        
        # Step 1: Click report block
        print("🎯 Clicking report block...")
        try:
            # Wait for element to be available
            await self.page.wait_for_selector(self.REPORT_BLOCK_SELECTOR, timeout=10000)
            await self.page.click(self.REPORT_BLOCK_SELECTOR)
            await self.page.wait_for_timeout(3000)
            print("✅ Report block clicked")
        except Exception as e:
            print(f"❌ Report block click failed: {e}")
            return False
        
        # Step 2: Click three dots
        print("⚙️ Clicking three dots menu...")
        try:
            await self.page.wait_for_selector(self.MORE_OPTIONS_SELECTOR, timeout=10000)
            await self.page.click(self.MORE_OPTIONS_SELECTOR)
            await self.page.wait_for_timeout(3000)
            print("✅ Three dots clicked")
        except Exception as e:
            print(f"❌ Three dots click failed: {e}")
            return False
        
        # Step 3: Click Export data
        print("📊 Clicking Export data...")
        try:
            await self.page.wait_for_selector(self.EXPORT_DATA_SELECTOR, timeout=10000)
            await self.page.click(self.EXPORT_DATA_SELECTOR)
            await self.page.wait_for_timeout(5000)
            print("✅ Export data clicked")
        except Exception as e:
            print(f"❌ Export data click failed: {e}")
            # Try alternative selectors
            print("🔄 Trying alternative selectors...")
            alt_selectors = [
                "text=Export data",
                "[aria-label*='Export data']",
                "[title*='Export data']"
            ]
            
            for selector in alt_selectors:
                try:
                    await self.page.click(selector)
                    print(f"✅ Export clicked with: {selector}")
                    await self.page.wait_for_timeout(3000)
                    break
                except:
                    continue
            else:
                return False
        
        # Step 4: Click final Export button in dialog
        print("🚀 Clicking Export button in dialog...")
        await self.page.wait_for_timeout(3000)  # Wait for dialog
        
        # Try multiple selectors for the export button
        export_selectors = [
            "button:has-text('Export')",
            ".exportButton",
            "button.primaryBtn",
            "button[type='submit']",
            "[role='button']:has-text('Export')"
        ]
        
        clicked = False
        for selector in export_selectors:
            try:
                await self.page.wait_for_selector(selector, timeout=5000)
                await self.page.click(selector)
                print(f"✅ Export button clicked with: {selector}")
                clicked = True
                break
            except:
                continue
        
        if not clicked:
            print("❌ Could not find export button")
            return False
        
        return True

    async def handle_download(self):
        """Wait for and handle the download"""
        print("⬇️ Waiting for download...")
        
        try:
            # Wait for download to start
            async with self.page.expect_download(timeout=30000) as download_info:
                await self.page.wait_for_timeout(2000)
            
            download = await download_info.value
            
            # Create downloads folder structure
            downloads_base = "downloads"
            if not os.path.exists(downloads_base):
                os.makedirs(downloads_base)
            
            # Create PBI folder with current date inside downloads
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
            print("🤖 Power BI Export Automation (Fixed)")
            print("="*50)
            
            # Setup
            await self.setup_browser()
            
            # Login
            await self.login()
            
            # Wait for report
            await self.wait_for_report()
            
            # Export
            success = await self.export_report()
            
            if success:
                # Handle download
                try:
                    await self.handle_download()
                    print("🎉 Export completed successfully!")
                except Exception as e:
                    print(f"⚠️ Download handling failed: {e}")
                    print("Check if file was downloaded to your Downloads folder")
            else:
                print("❌ Export process failed")
            
            print("="*50)
            
            # Keep browser open for a few seconds to see results
            print("🔍 Keeping browser open for 10 seconds...")
            await self.page.wait_for_timeout(10000)
            
        except Exception as e:
            print(f"❌ Error: {e}")
        finally:
            if self.browser:
                await self.browser.close()


async def main():
    automation = PowerBIExportAutomation()
    await automation.run()


if __name__ == "__main__":
    asyncio.run(main())
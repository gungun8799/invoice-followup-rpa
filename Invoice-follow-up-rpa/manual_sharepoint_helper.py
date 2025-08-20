#!/usr/bin/env python3
"""
Manual SharePoint Upload Helper
Run this script to prepare files for manual SharePoint upload
"""

import os
import shutil
from datetime import datetime
import pytz

def prepare_sharepoint_upload():
    print("📁 Preparing files for SharePoint upload...")
    
    downloads_dir = "downloads"
    sharepoint_ready_dir = "sharepoint_ready"
    
    # Create SharePoint ready folder
    if not os.path.exists(sharepoint_ready_dir):
        os.makedirs(sharepoint_ready_dir)
    
    # Get current date for folder naming
    bangkok_tz = pytz.timezone('Asia/Bangkok')
    current_date = datetime.now(bangkok_tz)
    date_folder = current_date.strftime('%d-%m-%Y')
    
    target_folder = os.path.join(sharepoint_ready_dir, date_folder)
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)
    
    # Find and copy XLS files
    if os.path.exists(downloads_dir):
        xls_files = [f for f in os.listdir(downloads_dir) if f.endswith('.xls')]
        
        if xls_files:
            # Use the most recent file
            latest_file = max(xls_files, key=lambda x: os.path.getmtime(os.path.join(downloads_dir, x)))
            source_path = os.path.join(downloads_dir, latest_file)
            
            # Rename to TIMS format
            new_filename = f"TIMS_{date_folder}.xls"
            target_path = os.path.join(target_folder, new_filename)
            
            shutil.copy2(source_path, target_path)
            
            print(f"✅ File prepared for SharePoint upload:")
            print(f"   📁 Folder: {target_folder}")
            print(f"   📄 File: {new_filename}")
            print(f"   📊 Size: {os.path.getsize(target_path):,} bytes")
            print(f"\n📋 Manual upload steps:")
            print(f"   1. Open SharePoint site in browser")
            print(f"   2. Create folder: {date_folder}")
            print(f"   3. Upload file: {new_filename}")
            
            return True
        else:
            print("❌ No XLS files found in downloads folder")
            return False
    else:
        print("❌ Downloads folder not found")
        return False

if __name__ == "__main__":
    prepare_sharepoint_upload()

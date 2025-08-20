#!/usr/bin/env python3

import requests
import json

# SharePoint Configuration
SHAREPOINT_CLIENT_ID = "b83ac538-2586-4eb7-8689-884c44d93d65"
SHAREPOINT_CLIENT_SECRET = "X3J3OFF+Wk1TeTRmTzRVN1Fofk1kWHJFWTd6MTJOYjZsUThabWJwNw=="
SHAREPOINT_TENANT_ID = "b08ac93a-5ec6-4492-bd45-dcd11a311661"

def quick_test():
    print("🔍 Quick SharePoint Access Test")
    print("-" * 40)
    
    # Get token
    token_url = f"https://login.microsoftonline.com/{SHAREPOINT_TENANT_ID}/oauth2/v2.0/token"
    token_data = {
        'grant_type': 'client_credentials',
        'client_id': SHAREPOINT_CLIENT_ID,
        'client_secret': SHAREPOINT_CLIENT_SECRET,
        'scope': 'https://graph.microsoft.com/.default'
    }
    
    try:
        response = requests.post(token_url, data=token_data, timeout=10)
        if response.status_code != 200:
            print(f"❌ Token failed: {response.status_code}")
            return
        
        access_token = response.json().get('access_token')
        print("✅ Got access token")
        
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        # Test 1: Try to get organization info
        print("\n📋 Test 1: Organization access")
        org_response = requests.get("https://graph.microsoft.com/v1.0/organization", headers=headers, timeout=10)
        print(f"Status: {org_response.status_code}")
        
        # Test 2: Try to search for sites
        print("\n📋 Test 2: Sites search")
        search_response = requests.get("https://graph.microsoft.com/v1.0/sites?search=*", headers=headers, timeout=10)
        print(f"Status: {search_response.status_code}")
        
        if search_response.status_code == 200:
            sites = search_response.json().get('value', [])
            print(f"Found {len(sites)} sites")
            for site in sites[:3]:  # Show first 3
                print(f"  - {site.get('displayName', 'No name')}")
        
        # Test 3: Try direct root site
        print("\n📋 Test 3: Root SharePoint site")
        root_response = requests.get("https://graph.microsoft.com/v1.0/sites/thlotuss.sharepoint.com", headers=headers, timeout=10)
        print(f"Status: {root_response.status_code}")
        
        if root_response.status_code == 200:
            root_data = root_response.json()
            print(f"Root site: {root_data.get('displayName', 'No name')}")
            
            # Test 4: Get subsites
            print("\n📋 Test 4: Subsites")
            subsites_response = requests.get("https://graph.microsoft.com/v1.0/sites/thlotuss.sharepoint.com/sites", headers=headers, timeout=10)
            print(f"Status: {subsites_response.status_code}")
            
            if subsites_response.status_code == 200:
                subsites = subsites_response.json().get('value', [])
                print(f"Found {len(subsites)} subsites")
                for subsite in subsites:
                    name = subsite.get('displayName', 'No name')
                    url = subsite.get('webUrl', 'No URL')
                    print(f"  - {name}")
                    print(f"    URL: {url}")
                    
                    if 'TradeInvoice' in name or 'trade' in name.lower():
                        print(f"    ✅ FOUND MATCHING SITE!")
                        return subsite.get('id')
        
        print("\n❌ No matching site found")
        return None
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return None

if __name__ == "__main__":
    quick_test()
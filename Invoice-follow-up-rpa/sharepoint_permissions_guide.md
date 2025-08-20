# SharePoint Access Issue - Permissions Guide

## 🔍 Diagnosis Results

✅ **Authentication**: Working - Access token obtained successfully  
❌ **SharePoint Access**: Failed - 401/403 errors  

### Error Details:
- Organization access: `Status 403` (Forbidden)
- Sites search: `Status 401` (Unauthorized) 
- Root site access: `Status 401` (Unauthorized)

## 🎯 Root Cause

The Azure App Registration needs **additional permissions** to access SharePoint sites via Microsoft Graph API.

## 🔧 Required Fix: Add SharePoint Permissions

### Step 1: Access Azure Portal
1. Go to [portal.azure.com](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Find app: **"Trade Invoice Notification project 01-08-2025"**
4. Client ID: `b83ac538-2586-4eb7-8689-884c44d93d65`

### Step 2: Add Required Permissions
Click **"API permissions"** > **"Add a permission"** > **"Microsoft Graph"** > **"Application permissions"**

Add these permissions:
- ✅ `Sites.Read.All` - Read all site collections
- ✅ `Sites.ReadWrite.All` - Read and write all site collections  
- ✅ `Files.ReadWrite.All` - Read and write files in all site collections

### Step 3: Grant Admin Consent
- After adding permissions, click **"Grant admin consent for [Organization]"**
- ✅ Confirm the consent dialog

### Step 4: Verify Permissions
The permissions should show:
```
Microsoft Graph (3):
- Sites.Read.All ✅ Granted for [Organization]
- Sites.ReadWrite.All ✅ Granted for [Organization]  
- Files.ReadWrite.All ✅ Granted for [Organization]
```

## 🧪 Test After Permission Changes

Run this command to test access:
```bash
python quick_sharepoint_test.py
```

Expected results after permissions are granted:
- Organization access: `Status 200` ✅
- Sites search: `Status 200` ✅  
- Root site access: `Status 200` ✅

## 🌐 Alternative: Manual Upload Process

If permissions cannot be granted immediately, the automation still provides:

### Current Working Features:
✅ **TIMS Export**: Fully automated  
✅ **File Download**: ZIP + Excel files saved locally  
✅ **File Extraction**: Automatic extraction to downloads/  

### Manual SharePoint Upload:
1. Files are saved in: `downloads/InvExpWaitTh_*.xls`
2. Manually upload to SharePoint site
3. Create folder with today's date: `17-08-2025`
4. Rename file to: `TIMS_17-08-2025.xls`

## 📞 Next Steps

**For IT Administrator:**
- Configure Azure app permissions as described above
- Test using `python quick_sharepoint_test.py`
- Confirm all status codes return 200

**For End Users:**
- Continue using automation for TIMS export
- Files will be available locally in downloads folder
- Upload manually to SharePoint if needed

**For Development:**
- Consider alternative file storage solutions
- OneDrive integration as backup option
- Network file share integration

## 🔍 Troubleshooting

If permissions are granted but still getting errors:

1. **Wait 5-10 minutes** - Permission changes can take time to propagate
2. **Check site URL** - Verify the exact SharePoint site URL
3. **Verify tenant** - Confirm the app is in the correct Azure tenant
4. **Test with Graph Explorer** - Use Microsoft Graph Explorer to test permissions

## 📋 Contact Information

If you need assistance configuring these permissions:
- Azure Portal: [portal.azure.com](https://portal.azure.com)
- Microsoft Graph Explorer: [developer.microsoft.com/graph/graph-explorer](https://developer.microsoft.com/graph/graph-explorer)
- App ID: `b83ac538-2586-4eb7-8689-884c44d93d65`
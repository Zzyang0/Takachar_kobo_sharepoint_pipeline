# 🔄 Kobo to SharePoint Direct Transfer Pipeline

A powerful Python script that transfers media files directly from Kobo Toolbox to SharePoint without local storage. Features incremental uploads, smart duplicate detection, and customizable naming conventions.

## 🌟 Features

- **🚀 Direct Transfer**: Streams media directly from Kobo to SharePoint without downloading locally
- **📊 Incremental Uploads**: Only uploads new files, skips existing ones automatically
- **🎯 Smart Duplicate Detection**: Handles both custom and fallback naming formats
- **📝 Custom Naming**: Supports date-based naming (YYYY-MM-DD_receiptType_rowNum.ext)
- **🔄 Multiple Fallback Methods**: Robust media URL extraction with multiple parsing strategies
- **📁 Organized Structure**: Creates form-based folder hierarchy in SharePoint
- **⚡ Efficient Processing**: Handles large files with chunked uploads
- **🔒 Secure Authentication**: Uses Microsoft Graph API with OAuth2

## 📋 Prerequisites

### Required Software
- Python 3.7 or higher
- pip (Python package manager)

### Required Accounts
- **Kobo Toolbox Account** with API access
- **Microsoft 365 Account** with SharePoint access
- **Azure App Registration** for SharePoint API access

## 🛠️ Installation

### 1. Clone or Download
```bash
# If using git
git clone <repository-url>
cd formdata_sharepoint-main

# Or download and extract the files
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

**Or install manually:**
```bash
pip install requests pandas python-dotenv
```

### 3. Environment Setup
Create a `.env` file in the project directory:

```env
# Kobo Toolbox API Token
API_TOKEN=your_kobo_api_token_here

# Microsoft 365 / SharePoint Credentials
TENANT_ID=your_tenant_id_here
CLIENT_ID=your_client_id_here
CLIENT_SECRET=your_client_secret_here
SITE_ID=your_sharepoint_site_id_here
```

## 🔧 Configuration

### Getting Kobo API Token
1. Log into your Kobo Toolbox account
2. Go to **Account Settings** → **API Tokens**
3. Create a new token with appropriate permissions
4. Copy the token to your `.env` file

### Setting Up Microsoft 365 App Registration
1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **New registration**
4. Configure the app with these permissions:
   - `Files.ReadWrite.All`
   - `Sites.ReadWrite.All`
5. Generate a client secret
6. Note down: Tenant ID, Client ID, Client Secret, and Site ID

## 🚀 Usage

### Basic Usage
```bash
python new.py
```

### Interactive Form Selection
The script will:
1. **Authenticate** with SharePoint
2. **Fetch** available Kobo forms
3. **Display** form list with creation dates
4. **Prompt** you to select forms to process
5. **Transfer** media files with progress updates

### Example Output
```
🌟 Kobo to SharePoint Direct Transfer Pipeline
Transfers media directly without local storage
------------------------------------------------------------

🔐 Authenticating with SharePoint...
✅ SharePoint authentication successful

🔍 Fetching Kobo forms...
✅ Found 5 forms

📋 Available forms:
1. Fuel Loading Form - Created: 2025-01-10
2. Biochar Sales Form - Created: 2025-01-12
3. Biomass Purchase Form - Created: 2025-01-15

Enter form numbers to transfer (comma-separated) or 'all':
Your choice: 1,2

📁 Checking for existing upload folder...
📁 Using existing folder: KoboMedia_Direct_20250115_143025

🔍 Checking existing files in SharePoint folder: KoboMedia_Direct_20250115_143025
📊 Found 45 existing files across 2 forms

🚀 Starting incremental transfer of 2 forms...

📋 Processing form: Fuel Loading Form
   Found 12 submissions
   📝 Naming format: custom (example: 2025-01-15_fuel_loading_1.jpg)
   📝 Submission 1/12
      Found 2 media file(s) in: receipt_photos
      🔄 Streaming: 2025-01-15_fuel_loading_1.jpg
         Size: 2.45 MB
         ✅ Transferred successfully
```

## 📁 File Structure

### Generated SharePoint Structure
```
KoboMedia_Direct_20250115_143025/
├── Fuel_Loading_Form/
│   ├── receipt_photos/
│   │   ├── 2025-01-15_fuel_loading_1.jpg
│   │   ├── 2025-01-15_fuel_loading_2.pdf
│   │   └── ...
│   └── other_columns/
├── Biochar_Sales_Form/
│   ├── invoice_attachments/
│   └── ...
└── ...
```

## 🎯 Naming Conventions

### Custom Naming (Preferred)
When forms have `Date` and `Receipt_Type` columns:
```
2025-01-15_fuel_loading_biomass_unloading_biochar_1.jpg
2025-01-16_biochar_sell_2.pdf
2025-01-17_biomass_buy_3.jpeg
```

### Fallback Naming
When date/type columns are missing:
```
row1_anjalirajput_attachments_93942a749e404ecfbf593b2e73ce1331_8fb68f64-b4af-4c9c-9d2f-2378ba6a861c_.JPEG
row2_document_attachment.pdf
```

## 📊 Statistics and Monitoring

The script provides detailed statistics:
- **Total media found**: Number of media files detected
- **Successfully transferred**: Files uploaded successfully
- **Failed transfers**: Files that failed to upload
- **Skipped (already exists)**: Duplicate files detected and skipped
- **Total size transferred**: Combined size of all uploaded files

## 🔍 Troubleshooting

### Common Issues

#### 1. Authentication Errors
```
❌ Missing required environment variables:
   - API_TOKEN
   - TENANT_ID
```
**Solution**: Check your `.env` file and ensure all credentials are correct.

#### 2. SharePoint Access Issues
```
❌ Could not access folder KoboMedia_Direct_20250115_143025
```
**Solution**: Verify your SharePoint permissions and site ID.

#### 3. Kobo API Errors
```
❌ Failed to fetch forms
```
**Solution**: Check your Kobo API token and ensure it has proper permissions.

#### 4. File Upload Failures
```
❌ SharePoint upload failed: 409
```
**Solution**: This usually indicates a duplicate file. The script should handle this automatically.

### Debug Mode
For detailed debugging, check the console output for:
- 🔍 File detection messages
- 📄 Existing file listings
- ⏭️ Skipped file notifications

## 🔄 Incremental Upload Logic

The script uses intelligent duplicate detection:

1. **Folder Detection**: Finds the most recent `KoboMedia_Direct_*` folder
2. **File Scanning**: Recursively scans all subfolders for existing files
3. **Naming Analysis**: Determines if form uses custom or fallback naming
4. **Duplicate Check**: Compares new files against existing ones by:
   - Exact filename match
   - Row number matching (for both naming formats)
   - Extension matching
5. **Skip Logic**: Automatically skips files that already exist

## 📝 Logging

The script provides real-time feedback:
- ✅ Success indicators
- ❌ Error messages
- 🔄 Progress updates
- 📊 Statistics summaries

## 🔒 Security Considerations

- **API Tokens**: Store securely in `.env` file (never commit to version control)
- **Permissions**: Use minimal required permissions for both Kobo and SharePoint
- **Network**: Ensure secure network connection when transferring sensitive data
- **Audit**: Review transfer logs regularly for any anomalies

## 🤝 Support

For issues or questions:
1. Check the troubleshooting section above
2. Review the console output for error messages
3. Verify your credentials and permissions
4. Ensure all dependencies are installed correctly

## 📄 License

This project is provided as-is for educational and operational purposes.

---

**Happy Transferring! 🚀**

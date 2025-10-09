# 🚛 Logistics Automation Suite v1.0
## Complete Freight Quote Automation for JT Custom Trailers

---

## 📋 Table of Contents
1. [Overview](#overview)
2. [System Components](#system-components)
3. [Installation & Setup](#installation--setup)
4. [Quick Start Guide](#quick-start-guide)
5. [Detailed Workflows](#detailed-workflows)
6. [Features Reference](#features-reference)
7. [Data Management](#data-management)
8. [Tips & Best Practices](#tips--best-practices)
9. [Troubleshooting](#troubleshooting)

---

## 🎯 Overview

The Logistics Automation Suite is a complete solution for automating freight quote requests for auction purchases. It eliminates manual data entry, standardizes communications, and tracks all quotes and costs in one place.

### What It Does:
- ✅ Extracts logistics info from PDF invoices automatically
- ✅ Generates professional, standardized quote request emails
- ✅ Manages freight company contact database
- ✅ Tracks all quotes, responses, and costs
- ✅ Creates reusable templates for repeat vendors
- ✅ Generates reports and analytics

---

## 🧩 System Components

### 1. **Logistics Automation Suite Launcher** 
`Logistics-Automation-Suite.ps1`
- **Central hub** for all tools
- Launch pad for quick workflows
- One-click access to all features

### 2. **Configuration Tool**
`Logistics-Config-GUI.ps1`
- PDF invoice parser (auto-extract contact info, dates, addresses)
- Interactive form for auction details
- Subject line generator
- Template manager (save/load vendor configs)

### 3. **Recipient Manager**
`Freight-Recipient-Manager.ps1`
- Freight company contact database
- Categories: Favorites, Regular, Occasional, Route-Specific
- Track usage statistics
- Quick email list copying
- Export to CSV

### 4. **Quote Tracker**
`Auction-Quote-Tracker.ps1`
- Track all auctions and their status
- Log quotes received (company, amount, date)
- Monitor freight costs over time
- Status tracking: Pending → Quoted → Booked → Completed
- Export reports for accounting

### 5. **Email Generator**
`Integrated-LogisticsEmail.ps1`
- Generates HTML emails from configuration
- Creates PDF attachments with lot images
- Optionally creates Outlook draft
- Standardized formatting

### 6. **Template Creator**
`Create-StarterTemplates.ps1`
- One-time setup: Creates starter templates
- Pre-configured for common vendors
- Ready-to-customize templates

---

## 🔧 Installation & Setup

### Prerequisites
- **Windows 10/11**
- **PowerShell 5.1+** (included in Windows)
- **Microsoft Outlook** (optional, for draft creation)
- **Excel** (optional, for PDF conversion)

### Initial Setup

1. **Extract all files** to a folder (e.g., `C:\LogisticsAutomation\`)

2. **Run the template creator** (first time only):
   ```powershell
   .\Create-StarterTemplates.ps1
   ```
   This creates:
   - `.\Templates\Brolyn_Auctions.json`
   - `.\Templates\Ritchie_Bros.json`
   - `.\Templates\Purple_Wave.json`
   - `.\Templates\GovDeals.json`
   - `.\Templates\Generic_Template.json`

3. **Launch the suite**:
   ```powershell
   .\Logistics-Automation-Suite.ps1
   ```

4. **Add your freight companies**:
   - Click "Recipient Manager"
   - Add your favorite carriers
   - Mark common ones as "Favorite"

### Folder Structure
```
LogisticsAutomation/
├── Logistics-Automation-Suite.ps1        # Main launcher
├── Logistics-Config-GUI.ps1              # Configuration tool
├── Freight-Recipient-Manager.ps1         # Contact manager
├── Auction-Quote-Tracker.ps1             # Quote tracker
├── Integrated-LogisticsEmail.ps1         # Email generator
├── Create-StarterTemplates.ps1           # Template creator
├── Templates/                             # Vendor configurations
│   ├── Brolyn_Auctions.json
│   └── ... (other templates)
├── Data/                                  # Auto-created
│   ├── FreightRecipients.json            # Contact database
│   └── AuctionQuotes.json                # Quote history
├── Output/                                # Generated files
│   ├── LogisticsEmail_*.html
│   └── AuctionLots_*.pdf
└── Invoices/                              # Store PDFs here
    └── ... (your invoice PDFs)
```

---

## 🚀 Quick Start Guide

### Workflow 1: Brand New Auction (Brolyn Example)

1. **Launch Suite**
   ```powershell
   .\Logistics-Automation-Suite.ps1
   ```

2. **Click "🚀 New Brolyn Auction"**
   - Loads Brolyn template automatically
   - Pre-filled with Brolyn contact info

3. **In Configuration Tool:**
   - Click "📄 Import from PDF Invoice"
   - Select your Brolyn invoice PDF
   - Review auto-extracted data:
     - ✅ Phone: (574) 891-3111
     - ✅ Email: logistics@brolynauctions.com
     - ✅ Pickup addresses
     - ✅ Dates and special notes
   - Fill in any missing fields
   - Click "🔄 Auto-Generate Subject"
   - Click "💾 Save Configuration" → `Brolyn_Oct09_2025.json`

4. **Back in Main Menu:**
   - Click "📨 Generate Email"
   - Select:
     - CSV file with lot data
     - Configuration file you just saved
     - Image directory
   - Script creates HTML email + PDF attachment
   - Opens Outlook draft automatically

5. **Send to Carriers:**
   - Open "📧 Recipient Manager"
   - Select favorite freight companies
   - Click "📋 Copy Selected Emails"
   - Paste into Outlook "To:" field
   - Send!

6. **Track Quotes:**
   - Open "📊 Quote Tracker"
   - Click "➕ New Auction"
   - Enter basic info
   - As quotes come back, click "💵 Add Quote"
   - Select winning carrier
   - Mark as "Booked"

---

### Workflow 2: Repeat Vendor (Template Reuse)

1. **Launch Configuration Tool**
2. **Template Manager** (left panel):
   - Select "Brolyn_Auctions"
   - Click "📂 Load Selected"
3. **Update variables only:**
   - Change pickup date
   - Update pallet count
   - Auto-generate new subject
4. **Save & Generate Email**

Total time: **2-3 minutes** vs 15+ minutes manually!

---

## 📖 Detailed Workflows

### Creating Custom Templates

**For a new auction house you use frequently:**

1. Launch Configuration Tool
2. Fill in all standard information:
   - Auction company name
   - Contact phone/email
   - Typical delivery address
   - Common special requirements
3. Add typical special notes
4. Click "📋 Save as Template"
5. Name it descriptively: `GoodwinAuctions_Michigan`
6. Next time, just load template and adjust dates!

### PDF Invoice Parsing

**Supported formats:**
- ✅ **Brolyn Auctions** - Fully supported with special parser
- ✅ **Text-based PDFs** - Generic extraction
- ⚠️ **Scanned PDFs** - Limited (manual entry recommended)

**What gets extracted:**
- Phone numbers: `(574) 891-3111`, `574-891-3111`, `574.891.3111`
- Emails: `logistics@company.com`
- Addresses: Full street addresses with city, state, ZIP
- Dates: `Monday October 7, 2025`, `10/7/2025`
- Special notes: Load windows, requirements, policies

**Tips for best results:**
1. Use original PDF from vendor (not scanned)
2. If extraction fails, save as text-based PDF
3. Review all extracted data before saving
4. Manual entry is always available as fallback

### Managing Recipients

**Categories:**
- **Favorite** - Your go-to carriers (color-coded in tracker)
- **Regular** - Commonly used
- **Occasional** - Use sometimes
- **Route-Specific** - Specialists for certain routes

**Best practices:**
1. Add all carriers you've worked with
2. Mark top 3-5 as "Favorites"
3. Add specialties: "RV Parts", "Heavy Equipment"
4. Add routes: "MI to OH", "Multi-state"
5. Keep notes: "Best for rushed jobs", "Great communication"

**Quick actions:**
- Select multiple, click "📋 Copy Selected Emails"
- Paste directly into Outlook
- Track "Times Used" to identify top performers

### Using the Quote Tracker

**Auction Statuses:**
- **Pending Quotes** - Just sent quote requests
- **Quotes Received** - Got responses, evaluating
- **Booked** - Selected carrier, waiting for pickup
- **In Transit** - Freight is moving
- **Completed** - Delivered successfully
- **Cancelled** - Auction cancelled or freight not needed

**Adding quotes:**
1. Open auction record
2. Click "💵 Add Quote"
3. Enter:
   - Carrier company (auto-suggests from Recipients)
   - Quote amount
   - Contact details
   - Any notes
4. Quote is logged with timestamp

**Selecting winner:**
1. Review all quotes in detail view
2. Click "Select This Quote" on winner
3. Status automatically updates to "Booked"
4. Updates recipient's "Times Used" counter

**Analytics:**
- View total freight costs
- Track quote response times
- Identify best carriers
- Export for accounting reports

---

## 🎨 Features Reference

### Configuration Tool Features

| Feature | Description |
|---------|-------------|
| **PDF Import** | Extract data from invoice PDFs |
| **Template Manager** | Save/load vendor configurations |
| **Subject Generator** | Auto-create standardized subject lines |
| **Special Notes Library** | Quick-add common requirements |
| **Multi-pickup Support** | Handle multiple pickup locations |
| **Preview Mode** | See JSON output before saving |

### Recipient Manager Features

| Feature | Description |
|---------|-------------|
| **Contact Database** | Store all freight company info |
| **Category Filtering** | Filter by Favorite, Regular, etc. |
| **Search** | Find companies by name, email, or notes |
| **Usage Tracking** | Track how many times used |
| **Email Copying** | Copy multiple emails to clipboard |
| **CSV Export** | Export full contact list |
| **Statistics** | View carrier performance stats |

### Quote Tracker Features

| Feature | Description |
|---------|-------------|
| **Auction Management** | Track all auction purchases |
| **Quote Logging** | Record all quotes received |
| **Status Tracking** | Monitor quote → booking → delivery |
| **Cost Analysis** | Track total freight expenses |
| **Color Coding** | Visual status indicators |
| **Detail View** | Complete auction information |
| **CSV Export** | Generate accounting reports |

---

## 💾 Data Management

### Data Storage

All data is stored in JSON format in the `Data/` folder:

**FreightRecipients.json** - Contact database
```json
{
  "CompanyName": "Maddy Freight Services",
  "ContactName": "Maddy Clark",
  "Email": "maddy@maddyfreight.com",
  "Category": "Favorite",
  "Specialties": ["RV Parts", "Palletized Freight"],
  "TimesUsed": 15
}
```

**AuctionQuotes.json** - Auction tracking
```json
{
  "AuctionDate": "2025-10-09",
  "AuctionCompany": "Brolyn Auctions",
  "Status": "Booked",
  "Quotes": [
    {
      "Company": "Maddy Freight",
      "Amount": 2450.00,
      "ReceivedDate": "2025-10-08"
    }
  ]
}
```

### Backup Recommendations

**Important files to backup:**
- `Data/FreightRecipients.json` - Your contact database
- `Data/AuctionQuotes.json` - Your quote history
- `Templates/*.json` - Your custom templates

**Backup methods:**
1. **Manual**: Copy entire `Data/` folder weekly
2. **OneDrive/Dropbox**: Store entire folder in cloud
3. **Git**: Version control for templates

---

## 💡 Tips & Best Practices

### Email Generation
- ✅ Always review auto-generated subject lines
- ✅ Use PDF attachments when possible
- ✅ Send to 3-5 carriers for competitive quotes
- ✅ Use "Copy Selected Emails" for quick distribution

### Template Management
- ✅ Create templates for vendors you use 3+ times
- ✅ Name templates descriptively: `Vendor_Location_Type`
- ✅ Update templates when vendor info changes
- ✅ Use Generic_Template as starting point for new vendors

### Recipient Organization
- ✅ Mark your top 3-5 carriers as "Favorites"
- ✅ Add detailed notes about carrier performance
- ✅ Update contact info when it changes
- ✅ Remove carriers that consistently don't respond

### Quote Tracking
- ✅ Enter auction details immediately after quote request
- ✅ Log quotes as soon as received
- ✅ Update status through entire lifecycle
- ✅ Export monthly reports for accounting

### PDF Processing
- ✅ Store invoice PDFs in `Invoices/` folder
- ✅ Name consistently: `Vendor_Date_Invoice#.pdf`
- ✅ Keep PDFs even after extraction for reference
- ✅ If extraction fails, try "Print to PDF" to create text-based version

---

## 🔧 Troubleshooting

### PDF Import Issues

**Problem**: "Could not extract PDF text"
**Solutions**:
1. Check if PDF is text-based (not scanned image)
2. Try "Print to PDF" from original to create clean version
3. Use "Save As" to create new PDF
4. Manual entry is always available

**Problem**: Extracted data is garbled
**Solutions**:
1. Review each field before saving
2. Use manual entry for problematic fields
3. Some PDFs have unusual formatting - extraction helps but may need cleanup

### Outlook Draft Creation

**Problem**: "Failed to create Outlook email"
**Solutions**:
1. Ensure Outlook is installed and configured
2. Open Outlook before running script
3. Check if Outlook is default email client
4. Fallback: Open HTML file and copy/paste into Outlook

### Template Loading

**Problem**: Template not appearing in list
**Solutions**:
1. Check `Templates/` folder exists
2. Verify `.json` file extension
3. Click "🔄 Refresh List"
4. Ensure JSON is valid (no syntax errors)

### General Performance

**Problem**: Scripts running slow
**Solutions**:
1. Close unnecessary programs
2. Check antivirus isn't scanning scripts
3. Move large image folders to SSD
4. Limit images per lot (use `MaxImagesPerLot` parameter)

---

## 📞 Support & Updates

### Getting Help
- Review this README first
- Check `Logs/` folder for error details
- Review error messages in console output

### Feature Requests
Create custom templates for specific workflows using existing templates as examples.

### Version History
- **v1.0.0** (2025-01-07) - Initial release
  - Configuration Tool with PDF parsing
  - Recipient Manager
  - Quote Tracker
  - Email Generator
  - Template system
  - Master launcher

---

## 📄 License & Credits

**Author**: John O'Neill Sr.  
**Company**: Azure Innovators / JT Custom Trailers  
**Version**: 1.0.0  
**Created**: January 7, 2025

---

## 🎉 Conclusion

The Logistics Automation Suite transforms hours of manual work into minutes of automated efficiency. By combining intelligent data extraction, template management, contact organization, and comprehensive tracking, you can focus on your business instead of repetitive data entry.

**Time Savings Per Auction:**
- Manual process: ~20-30 minutes
- Automated process: ~3-5 minutes
- **Savings: 85% reduction in time spent** 🎯

Launch the suite and get started today!

```powershell
.\Logistics-Automation-Suite.ps1
```

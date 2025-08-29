# Excel VBA Bulk Email Automation Tool

A VBA macro that automates sending customized bulk emails directly from Excel using Outlook.

✅ **Features**
- Send emails to multiple recipients at once  
- Supports **To, CC, BCC, Subject, Attachments, and Outlook signatures**  
- Simple **one-click execution** via a macro  
- Written in **VBA (Macro-enabled Excel)**

## 📂 Files Included
- `SendSelectedBulkEmails.bas` – VBA macro code  
- `SampleData.xlsm` – Excel workbook with sample email data 

> ⚠️ Only dummy emails are included in the sample data. Do not use real addresses in public repositories.

## 📝 How to Use
1. Open `SampleData.xlsm` in Excel and enable macros.  
2. Press `Alt + F8` → Run `SendSelectedBulkEmails`.  
3. Select the rows containing recipient info.  
4. Emails are sent automatically using your default Outlook account.  

> Notes:
> - Ensure all attachment paths exist in the sample data.  
> - The macro will automatically load your Outlook signature.  
> - You can temporarily edit your Outlook signature to include a default message (e.g., "Good day, please find the attached invoice for July 2025").

## 💡 Contributions
Feel free to **fork, test, or contribute**! Feedback and improvements are welcome.

## Technology Used
- **Language:** VBA (Macro-enabled Excel)  
- **Platform:** Microsoft Excel + Outlook

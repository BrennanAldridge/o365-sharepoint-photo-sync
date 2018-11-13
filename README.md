# o365-sharepoint-photo-sync
PowerShell script to sync user photos in Office 365 (Exchange to SharePoint)

TO USE:
1. Create a CSV file with 1 column: Email, and a row value for each email address you wish to sync.
2. Update the script with A) the path to your CSV file, B) a temporary storage location, and C) your organization's URL prefix.
3. Run the Script

WHAT IT DOES:
- download the user's Exchange Online Photo
- create the 3 thumbnail sizes for SharePoint and upload them to the MySite Host site collection
- update the SharePoint user profile (PictureURL) with the medium thumbnail version.

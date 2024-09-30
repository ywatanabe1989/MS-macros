To enable a specific macro-enabled template (Excel-Macros.xltm) for all Excel files you open, you can set it as the default template. Here’s how you can do it:

Save the Template:
Save your Excel-Macros.xltm file in the XLSTART folder. This folder is typically located at:
Windows: C:\Users\<YourUsername>\AppData\Roaming\Microsoft\Excel\XLSTART
Mac: ~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Excel
Set as Default Template:
Open Excel.
Go to File > Options > General.
Under When creating new workbooks, set the default template to your Excel-Macros.xltm file.
Enable Macros:
Go to File > Options > Trust Center > Trust Center Settings.
Click on Macro Settings and select Enable all macros (not recommended; potentially dangerous code can run) or Disable all macros with notification to be prompted each time.

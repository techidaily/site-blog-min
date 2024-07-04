---
title: How to Fix Excel Run Time Error 1004
date: 2024-05-19T18:32:11.611Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel Run Time Error 1004
keywords: repair excel,repair .xls,repair excel 2007,repair excel 2021,repair .xltm
---

## How to Fix Excel Run Time Error 1004

**Summary:** Run-time errors are windows-specific issues that occur while the program is running. This blog will teach you how to fix Excel run-time error 1004. In addition, you’ll learn about an Excel repair tool that can help fix the error 1004 if it occurs due to corruption in Excel files.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

VBA (Microsoft Visual Basic for Application) is an internal programming language in Microsoft Excel. Sometimes, when users try to run VBA or generate a Macro in Excel, the Run-time error 1004 may occur. This error may occur due to the presence of more legend entries in the chart, file conflict, incorrect Macro name, and corrupt Excel files. In this blog, we have discussed the reasons and shared some solutions to resolve run-time error 1004.

## **Why This Error Occurs?**

The run time error 1004 usually occurs when you run a VBA macro with the Legend Entries method to modify the legend entries in the MS Excel chart. It happens when the chart contains more legend entries than the available space, macro name conflicts, corrupt Excel files, or data-types mismatch in the VBA code.

## **Ways to Fix Excel Run-Time Error 1004?**

Try the below workarounds to fix Excel run-time error 1004:

### **Create a Macro to Reduce Chart Legend Font Size**

Sometimes, Excel throws the run-time error when you try to run VBA macro to change the legend entries in a Microsoft Excel chart. This error usually occurs when Microsoft Excel truncates the legend entries because of the more legend entries and less space availability. To fix this, try to create a macro that shrinks/minimize the font size of the Excel chart legend text before the VBA macro, and then restore the font size of the chart legend. Here is the macro code:

```
VBCopy
Sub ResizeLegendEntries()

With Worksheets("Sheet1").ChartObjects(1).Activate
      ' Store the current font size
      fntSZ = ActiveChart.Legend.Font.Size

'Temporarily change the font size.
      ActiveChart.Legend.Font.Size = 2

'Place your LegendEntries macro code here to make
         'the changes that you want to the chart legend.

' Restore the font size.
      ActiveChart.Legend.Font.Size = fntSZ
   End With

End Sub
Note: Make sure you have an Excel chart to run the code on the worksheet.
```

### **Uninstall Microsoft Work**

You may encounter a run-time error 1004 in Excel version 2009 or older versions due to conflicts between Microsoft works and Microsoft Excel. This error usually occurs if your system has both Microsoft Office and Microsoft Works. Uninstalling one of them will fix the issue. Try the below steps to uninstall Microsoft Work:

- First, open the **Task Manager** using the shortcut **CTRL + ALT + DEL** altogether
- The **Task Manager window** is displayed.

![Task Manager Window](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/09/task-manager-window.png)

- Click the **Process** tab, right-click on each program you want to close, and then click **End Task.**
- Stop all the running programs.
- Open the **Run** window and type **_appwiz.cpl_** to open the **Programs and Feature** window.

![Program and Features of Control Panel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/09/program-and-features-1024x516.png)

- Search for **Microsoft Works** and click **Uninstall**.

### **Try Deleting GWXL97.Xla File**

The Add-ins files with .xla extension in MS-EXCEL is used to provide additional functionality to Excel spreadsheets. Sometimes, deleting the GWXL97.XLA file fixes the run-time error. Here are the steps to delete this file:

- Make sure you have an **Admins rights**, open the **Windows Explorer**
- Follow the Path C:\\Programs Files\\MSOffice\\Office\\XLSTART.
- Find and right-click on the **GWXL97.XLA** file
- Click **Delete**.

### **Change Trust Center Settings**

Sometimes, run-time errors might arise because of incorrect security settings. The **Trust Center settings** help you find the **Privacy and security** settings for Microsoft Excel. Follow the below steps to change the **Trust center settings**:

- Open Microsoft Excel.
- Go to **File > Options.**
- The **Excel options** window is displayed.
- Choose **Trust Center**, and click **Trust Center Settings**.
- Tap on the **Macro Settings** tab, and select **Trust access to the VBA project object model.**

![Macro Settings in Microsoft Excel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/09/macro-settings.png)

- Click **OK**.

### **Run Open and Repair Tool**

The Runtime error also arises when MS Excel detects a corrupted worksheet. It automatically begins the File recovery mode and starts repairing it. However, if the Recovery mode fails to start, use the **Open and Repair** tool with the below steps:

- Click **File > Open**.
- Click the location and folder with a corrupted workbook.
- In the **Open** dialog box, choose the corrupted workbook.
- Click the arrow next to the **Open** tab, and go to the **Open and Repair** tab.
- Click **Repair**.

You can also opt for **Stellar Repair for Excel** if the Microsoft Excel’s built-in tool cannot fix the error.

### **Use Stellar Repair for Excel**

**Stellar Repair for Excel** is a professional software for repairing damage. xls, .xlsx, .xltm, .xltx, and .xlsm files and recovering all its objects. Here are the steps to fix the error using this tool:

- First, **download**, **install**, and run **Stellar Repair for Excel**.
- Click the **Browse** tab on the interface window to choose the corrupted Excel file you need to repair.
- Click **Scan**. You will see the scan progress in the scanning window.
- Click **OK**.
- The tool can let you preview all the recoverable Excel file components including tables, pivot tables, charts, formulas, etc.
- Click **Save** to save the repaired file.
- A **Save File dialog box** will appear with the below two options:
- Default location
- New location
- Choose a suitable option.
- Click the **Save** option to repair the Excel file that you have chosen.
- Once the repair is complete, it will display a message “**File repaired successfully**.”
- Click **OK**.

## **Conclusion**

Now you know the Excel run-time error 1004, its cause, and solutions. Follow the workarounds discussed in the blog to rectify the error quickly. However, **[Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)** makes your task of removing run-time errors easy. It’s a powerful software to fix all the issues with Excel files. Also, it helps in extracting data from the damaged file and saves it to a new Excel workbook.




<ins class="adsbygoogle"
     style="display:block"
     data-ad-client="ca-pub-7571918770474297"
     data-ad-slot="8358498916"
     data-ad-format="auto"
     data-full-width-responsive="true"></ins>
<ins class="adsbygoogle"
    style="display:block"
    data-ad-format="autorelaxed"
    data-ad-client="ca-pub-7571918770474297"
    data-ad-slot="1223367746"></ins>

<span class="atpl-alsoreadstyle">Also read:</span>
<div><ul>
<li><a href="https://blog-min.techidaily.com/how-to-play-hevc-h-265-video-on-mix-fold-3-by-aiseesoft-video-converter-play-hevc-video-on-android/"><u>How to play HEVC H.265 video on Mix Fold 3?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-call-logs-from-camon-30-pro-5g-by-fonelab-android-recover-call-logs/"><u>How to retrieve erased call logs from Camon 30 Pro 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-call-logs-from-realme-10t-5g-by-fonelab-android-recover-call-logs/"><u>How to retrieve erased call logs from Realme 10T 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-solve-mkv-lagging-problem-in-motorola-moto-g-5g-2023-by-aiseesoft-video-converter-play-mkv-on-android/"><u>How to solve MKV lagging problem in Motorola Moto G 5G (2023)?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-lost-files-from-honor-x50-by-fonelab-android-recover-data/"><u>How to retrieve lost files from Honor X50?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-xs-max-without-itunes-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Reset iPhone XS Max Without iTunes? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-oneplus-pin-by-drfone-android-unlock-android-unlock/"><u>How to remove OnePlus PIN</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-call-logs-from-tecno-by-fonelab-android-recover-call-logs/"><u>How to rescue lost call logs from Tecno</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-huawei-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Huawei</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-whatsapp-chat-history-from-iphone-se-2020-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How To Recover Whatsapp Chat History From iPhone SE (2020) | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-install-and-update-device-drivers-manually-in-windows-11-by-drivereasy-guide/"><u>How to install and update device drivers manually in Windows 11</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-call-history-on-nokia-c110-by-fonelab-android-recover-call-logs/"><u>How to restore wiped call history on Nokia C110?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-error-1015-while-restoring-iphone-xr-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to fix error 1015 while restoring iPhone XR | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-messages-from-meizu-by-fonelab-android-recover-messages/"><u>How to Rescue Lost Messages from Meizu</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-sony-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Sony using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-asus-rog-phone-8-contacts-an-easy-method-explained-by-fonelab-android-recover-contacts/"><u>How to Restore Deleted Asus ROG Phone 8 Contacts  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-pictures-files-from-honor-100-by-fonelab-android-recover-pictures/"><u>How To  Restore Missing Pictures Files from Honor 100.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-13-to-the-previous-iosipados-version-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 13 to the Previous iOS/iPadOS Version? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-pictures-files-from-samsung-galaxy-f54-5g-by-fonelab-android-recover-pictures/"><u>How To  Restore Missing Pictures Files from Samsung Galaxy F54 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-contacts-on-iphone-13-mini-4-methods-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Restore Contacts on iPhone 13 mini (4 Methods) | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-lava-blaze-2-pro-get-deleted-phone-number-back-with-ease-and-safety-by-fonelab-android-recover-contacts/"><u>How to Lava Blaze 2 Pro Get Deleted Phone Number Back with Ease and Safety</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-pictures-from-a38-by-fonelab-android-recover-pictures/"><u>How to recover deleted pictures from A38.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-spark-10c-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Spark 10C</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-oppo-a1x-5g-by-fonelab-android-recover-data/"><u>How to recover lost data from Oppo A1x 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-galaxy-m54-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Galaxy M54 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-iphone-xs-data-from-icloud-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover iPhone XS Data From iCloud? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-after-format-on-samsung-galaxy-a15-4g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery after format on Samsung Galaxy A15 4G</u></a></li>
<li><a href="https://ai-video-editing.techidaily.com/new-2024-approved-demystifying-the-vectorscope-a-comprehensive-guide-for-premiere-pro-color-correction/"><u>New 2024 Approved Demystifying the Vectorscope A Comprehensive Guide for Premiere Pro Color Correction</u></a></li>
<li><a href="https://ai-video-apps.techidaily.com/2024-approved-webm-to-mp3-conversion-made-simple-top-tools/"><u>2024 Approved WebM to MP3 Conversion Made Simple Top Tools</u></a></li>
<li><a href="https://location-social.techidaily.com/in-2024-how-to-leave-a-life360-group-on-vivo-y27-4g-without-anyone-knowing-drfone-by-drfone-virtual-android/"><u>In 2024, How To Leave a Life360 Group On Vivo Y27 4G Without Anyone Knowing? | Dr.fone</u></a></li>
<li><a href="https://android-transfer.techidaily.com/how-to-transfer-data-from-itel-s23-to-samsung-phone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Data from Itel S23 to Samsung Phone | Dr.fone</u></a></li>
<li><a href="https://android-transfer.techidaily.com/in-2024-how-to-transfer-data-from-poco-c50-to-samsung-phone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>In 2024, How to Transfer Data from Poco C50 to Samsung Phone | Dr.fone</u></a></li>
<li><a href="https://change-location.techidaily.com/in-2024-all-you-need-to-know-about-mega-greninja-for-vivo-y78-5g-drfone-by-drfone-virtual-android/"><u>In 2024, All You Need To Know About Mega Greninja For Vivo Y78 5G | Dr.fone</u></a></li>
<li><a href="https://android-location.techidaily.com/in-2024-fake-android-location-without-rooting-for-your-vivo-y17s-drfone-by-drfone-virtual/"><u>In 2024, Fake Android Location without Rooting For Your Vivo Y17s | Dr.fone</u></a></li>
<li><a href="https://unlock-android.techidaily.com/bypassing-google-account-with-vnrom-bypass-for-tecno-pop-8-by-drfone-android/"><u>Bypassing Google Account With vnROM Bypass For Tecno Pop 8</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/in-2024-the-top-5-android-apps-that-use-fingerprint-sensor-to-lock-your-apps-on-nokia-c12-pro-by-drfone-android/"><u>In 2024, The Top 5 Android Apps That Use Fingerprint Sensor to Lock Your Apps On Nokia C12 Pro</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/in-2024-how-to-use-snapchat-location-spoofer-to-protect-your-privacy-on-apple-iphone-6s-drfone-by-drfone-virtual-ios/"><u>In 2024, How to use Snapchat Location Spoofer to Protect Your Privacy On Apple iPhone 6s? | Dr.fone</u></a></li>
</ul></div>



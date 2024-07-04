---
title: How to Fix Excel 2016 Run Time Error 1004
date: 2024-05-19T18:32:11.548Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2016 Run Time Error 1004
keywords: repair .xls,repair excel 2023,repair excel 2016,repair .xlsm,repair .xlsx
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
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-14-without-losing-any-content-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 14 without Losing Any Content? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-install-the-latest-iosipados-beta-version-on-iphone-12-pro-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Install the Latest iOS/iPadOS Beta Version on iPhone 12 Pro? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-back-lost-contacts-from-oppo-k11x-by-fonelab-android-recover-contacts/"><u>How to get back lost contacts from Oppo K11x.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-call-logs-from-f5-pro-5g-by-fonelab-android-recover-call-logs/"><u>How to rescue lost call logs from F5 Pro 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-videos-from-nubia-z50-ultra-by-fonelab-android-recover-video/"><u>How to Rescue Lost Videos from Nubia Z50 Ultra</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-oppo-f23-5g-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Oppo F23 5G using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-after-format-on-lava-blaze-2-pro-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery after format on Lava Blaze 2 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-poco-c51-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Poco C51 using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-honor-x50i-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Honor X50i</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-a-excel-2019-document-online-by-ldigisigner-sign-a-excel-sign-a-excel/"><u>How to sign a Excel 2019 document online</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-pictures-from-oneplus-open-by-fonelab-android-recover-pictures/"><u>How to Rescue Lost Pictures from OnePlus Open?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-the-unable-to-record-macro-error-in-excel-2023-stellar-by-stellar-guide/"><u>How to Fix the Unable to Record Macro Error in Excel 2023? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-call-logs-from-vivo-y27s-by-fonelab-android-recover-call-logs/"><u>How to retrieve erased call logs from Vivo Y27s?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-malfunctioning-your-drivers-with-windows-device-manager-on-windows-1110-by-drivereasy-guide/"><u>How to identify malfunctioning your drivers with Windows Device Manager on Windows 11/10</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-files-after-iphone-7-factory-reset-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Files after iPhone 7 Factory Reset? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-a-ott-file-document-online-by-ldigisigner-sign-a-word-sign-a-word/"><u>How to Sign a .ott file document online</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-videos-from-your-14-by-fonelab-android-recover-video/"><u>How to recover old videos from your 14</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-14-plus-without-apple-password-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Reset iPhone 14 Plus Without Apple Password? | Stellar</u></a></li>
<li><a href="https://android-location.techidaily.com/in-2024-how-to-fake-gps-on-android-without-mock-location-for-your-motorola-defy-2-drfone-by-drfone-virtual/"><u>In 2024, How to Fake GPS on Android without Mock Location For your Motorola Defy 2 | Dr.fone</u></a></li>
<li><a href="https://iphone-location.techidaily.com/in-2024-how-to-enable-disable-and-change-safari-location-on-apple-iphone-6-plus-drfone-by-drfone-virtual-ios/"><u>In 2024, How to Enable, Disable, and Change Safari Location on Apple iPhone 6 Plus | Dr.fone</u></a></li>
<li><a href="https://review-topics.techidaily.com/in-2024-complete-tutorial-to-use-vpna-to-fake-gps-location-on-huawei-nova-y91-drfone-by-drfone-virtual-android/"><u>In 2024, Complete Tutorial to Use VPNa to Fake GPS Location On Huawei Nova Y91 | Dr.fone</u></a></li>
<li><a href="https://animation-videos.techidaily.com/new-in-2024-how-to-animate-graphics-in-canva/"><u>New In 2024, How To Animate Graphics in Canva</u></a></li>
<li><a href="https://animation-videos.techidaily.com/new-in-2024-whats-music-video-gif-and-how-to-add-music-to-gif-files/"><u>New In 2024, Whats Music Video GIF & How to Add Music to GIF Files</u></a></li>
<li><a href="https://sim-unlock.techidaily.com/tutorial-to-change-samsung-galaxy-m34-imei-without-root-a-comprehensive-guide-by-drfone-android/"><u>Tutorial to Change Samsung Galaxy M34 IMEI without Root A Comprehensive Guide</u></a></li>
<li><a href="https://ai-voice-clone.techidaily.com/new-what-is-an-ai-text-generator-wondershare-virbo-glossary-for-2024/"><u>New What Is an AI Text Generator? | Wondershare Virbo Glossary for 2024</u></a></li>
<li><a href="https://howto.techidaily.com/why-your-xiaomi-13t-pro-screen-might-be-unresponsive-and-how-to-fix-it-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>Why Your Xiaomi 13T Pro Screen Might be Unresponsive and How to Fix It | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/does-samsung-galaxy-xcover-7-has-native-hevc-support-by-aiseesoft-video-converter-play-hevc-video-on-android/"><u>Does Samsung Galaxy XCover 7 has native HEVC support?</u></a></li>
<li><a href="https://android-location.techidaily.com/in-2024-easy-ways-to-manage-your-honor-x50i-location-settings-drfone-by-drfone-virtual/"><u>In 2024, Easy Ways to Manage Your Honor X50i Location Settings | Dr.fone</u></a></li>
<li><a href="https://android-unlock.techidaily.com/in-2024-full-tutorial-to-bypass-your-vivo-y78t-face-lock-by-drfone-android/"><u>In 2024, Full Tutorial to Bypass Your Vivo Y78t Face Lock?</u></a></li>
<li><a href="https://animation-videos.techidaily.com/ways-to-master-to-find-10-best-free-animated-text-generator-for-2024/"><u>Ways to Master to Find 10 Best Free Animated Text Generator for 2024</u></a></li>
<li><a href="https://techidaily.com/undelete-lost-music-from-tecno-spark-go-2024-by-fonelab-android-recover-music/"><u>Undelete lost music from Tecno Spark Go (2024)</u></a></li>
</ul></div>



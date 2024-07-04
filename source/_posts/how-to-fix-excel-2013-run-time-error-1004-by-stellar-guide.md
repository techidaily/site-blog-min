---
title: How to Fix Excel 2013 Run Time Error 1004
date: 2024-05-19T18:32:11.539Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2013 Run Time Error 1004
keywords: repair .xlsx,repair excel 2010,repair excel 2003,repair excel 2007,repair excel 2016
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
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-messages-from-nubia-red-magic-8s-pro-by-fonelab-android-recover-messages/"><u>How to retrieve erased messages from Nubia Red Magic 8S Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-notes-from-iphone-xr-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Notes from iPhone XR? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-messages-from-y78plus-t1-edition-by-fonelab-android-recover-messages/"><u>How to Rescue Lost Messages from Y78+ (T1) Edition</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-videos-on-htc-u23-pro-by-fonelab-android-recover-video/"><u>How to restore wiped videos on HTC U23 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-xiaomi-redmi-note-12t-pro-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Xiaomi Redmi Note 12T Pro to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-honor-x8b-by-fonelab-android-recover-data/"><u>How to recover lost data from Honor X8b?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-pictures-files-from-moto-g-stylus-2023-by-fonelab-android-recover-pictures/"><u>How To  Restore Missing Pictures Files from Moto G Stylus (2023).</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-lost-files-from-reno-10-proplus-5g-by-fonelab-android-recover-data/"><u>How to retrieve lost files from Reno 10 Pro+ 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-iphone-11-pro-max-ios-system-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair iPhone 11 Pro Max iOS System? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-realme-note-50-with-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Realme Note 50 with Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-xiaomi-redmi-note-12-5g-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Xiaomi Redmi Note 12 5G to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-play-an-mp4-on-samsung-galaxy-a15-4g-by-aiseesoft-video-converter-play-mp4-on-android/"><u>How to play an MP4 on Samsung Galaxy A15 4G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-iphone-8-plus-system-issues-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair iPhone 8 Plus System Issues? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-messages-from-y78plus-by-fonelab-android-recover-messages/"><u>How to Rescue Lost Messages from Y78+</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-messages-from-realme-gt-neo-5-se-by-fonelab-android-recover-messages/"><u>How to retrieve erased messages from Realme GT Neo 5 SE</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-messages-on-oppo-a58-4g-by-fonelab-android-recover-messages/"><u>How to restore wiped messages on Oppo A58 4G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-7-plus-without-losing-anything-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 7 Plus without Losing Anything? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-back-lost-photos-from-gt-10-pro-by-fonelab-android-recover-photos/"><u>How to get back lost photos from GT 10 Pro.</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-infinix-hot-30-5g-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Infinix Hot 30 5G to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-nokia-g42-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Nokia G42 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-music-from-sony-xperia-1-v-by-fonelab-android-recover-music/"><u>How to retrieve erased music from Sony Xperia 1 V</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-sony-xperia-5-v-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Sony Xperia 5 V</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-pivot-table-field-name-is-not-valid-error-in-excel-2016-stellar-by-stellar-guide/"><u>How to fix Pivot Table Field Name is not Valid error in Excel 2016? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-huawei-nova-y71-get-deleted-photos-back-with-ease-and-safety-by-fonelab-android-recover-photos/"><u>How to Huawei Nova Y71 Get Deleted photos Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-videos-from-your-tecno-by-fonelab-android-recover-video/"><u>How to recover old videos from your Tecno</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-a-damaged-video-file-of-realme-narzo-60-5g-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair a Damaged video file of Realme Narzo 60 5G using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-make-a-digital-signature-for-word-2013-by-ldigisigner-sign-a-word-sign-a-word/"><u>How to make a digital signature for Word 2013</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-motorola-moto-g-stylus-5g-2023-by-fonelab-android-recover-data/"><u>How to recover lost data from Motorola Moto G Stylus 5G (2023)?</u></a></li>
<li><a href="https://ai-video-editing.techidaily.com/updated-are-you-searching-for-the-best-canon-lut-download-this-article-provides-the-best-recommendations-for-creating-the-best-professional-videos/"><u>Updated Are You Searching for the Best Canon LUT Download? This Article Provides the Best Recommendations for Creating the Best Professional Videos</u></a></li>
<li><a href="https://fake-location.techidaily.com/which-is-the-best-fake-gps-joystick-app-on-xiaomi-redmi-k70e-drfone-by-drfone-virtual-android/"><u>Which is the Best Fake GPS Joystick App On Xiaomi Redmi K70E? | Dr.fone</u></a></li>
<li><a href="https://ai-vdieo-software.techidaily.com/2024-approved-free-avi-video-rotation-made-easy-top-16-picks-for-windows-mac-android-and-iphone/"><u>2024 Approved Free AVI Video Rotation Made Easy Top 16 Picks for Windows, MAC, Android, and iPhone</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/pattern-locks-are-unsafe-secure-your-realme-c67-4g-phone-now-with-these-tips-by-drfone-android/"><u>Pattern Locks Are Unsafe Secure Your Realme C67 4G Phone Now with These Tips</u></a></li>
<li><a href="https://android-transfer.techidaily.com/in-2024-how-i-transferred-messages-from-tecno-pova-5-pro-to-iphone-12xs-max-in-seconds-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>In 2024, How I Transferred Messages from Tecno Pova 5 Pro to iPhone 12/XS (Max) in Seconds | Dr.fone</u></a></li>
<li><a href="https://ai-video-editing.techidaily.com/in-2024-how-to-turn-onoff-motion-blur-in-fortnite-is-it-good-for-fortnite/"><u>In 2024, How to Turn On/Off Motion Blur in Fortnite? Is It Good for Fortnite?</u></a></li>
<li><a href="https://android-location-track.techidaily.com/solutions-to-spy-on-realme-v30t-with-and-without-jailbreak-drfone-by-drfone-virtual-android/"><u>Solutions to Spy on Realme V30T with and without jailbreak | Dr.fone</u></a></li>
<li><a href="https://android-location.techidaily.com/in-2024-3-effective-methods-to-fake-gps-location-on-android-for-your-poco-m6-pro-5g-drfone-by-drfone-virtual/"><u>In 2024, 3 Effective Methods to Fake GPS location on Android For your Poco M6 Pro 5G | Dr.fone</u></a></li>
<li><a href="https://location-social.techidaily.com/4-most-known-ways-to-find-someone-on-tinder-for-nokia-g310-by-name-drfone-by-drfone-virtual-android/"><u>4 Most-Known Ways to Find Someone on Tinder For Nokia G310 by Name | Dr.fone</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/in-2024-how-to-screen-mirroring-oneplus-ace-2-drfone-by-drfone-android/"><u>In 2024, How to Screen Mirroring OnePlus Ace 2? | Dr.fone</u></a></li>
<li><a href="https://ios-unlock.techidaily.com/7-top-ways-to-resolve-apple-id-not-active-issue-for-iphone-x-by-drfone-ios/"><u>7 Top Ways To Resolve Apple ID Not Active Issue For iPhone X</u></a></li>
<li><a href="https://change-location.techidaily.com/how-to-stop-life360-from-tracking-you-on-nokia-c02-drfone-by-drfone-virtual-android/"><u>How to Stop Life360 from Tracking You On Nokia C02? | Dr.fone</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/in-2024-3-ways-for-android-pokemon-go-spoofing-on-oneplus-nord-ce-3-5g-drfone-by-drfone-virtual-android/"><u>In 2024, 3 Ways for Android Pokemon Go Spoofing On OnePlus Nord CE 3 5G | Dr.fone</u></a></li>
</ul></div>



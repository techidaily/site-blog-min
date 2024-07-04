---
title: How to Fix Excel 2000 Run Time Error 1004
date: 2024-05-19T18:32:11.477Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2000 Run Time Error 1004
keywords: repair excel 2023,repair excel 2021,repair .xlsm,repair excel 2019,repair excel 2016
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
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-oppo-reno-10-proplus-5g-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Oppo Reno 10 Pro+ 5G to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-messages-from-honor-x50-gt-by-fonelab-android-recover-messages/"><u>How to retrieve erased messages from Honor X50 GT</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-itel-p40-pictures-an-easy-method-explained-by-fonelab-android-recover-pictures/"><u>How to Restore Deleted Itel P40 Pictures  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-iphone-12-pro-max-data-from-ios-itunes-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How to Recover iPhone 12 Pro Max Data From iOS iTunes? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-music-from-xiaomi-13t-pro-by-fonelab-android-recover-music/"><u>How to retrieve erased music from Xiaomi 13T Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-techniques-to-transfer-data-from-honor-magic-vs-2-to-iphone-15141312-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Techniques to Transfer Data from Honor Magic Vs 2 to iPhone 15/14/13/12 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-vivo-s18-pro-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Vivo S18 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-videos-and-music-files-from-iphone-6-plus-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Photos, Videos & Music Files from iPhone 6 Plus | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-system-of-iphone-se-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair System of iPhone SE? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-pictures-from-oppo-f23-5g-by-fonelab-android-recover-pictures/"><u>How to recover deleted pictures from Oppo F23 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-nord-n30-se-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Nord N30 SE</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-k11-5g-by-fonelab-android-recover-data/"><u>How to recover lost data from K11 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-data-from-iphone-se-2022-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Data from iPhone SE (2022)? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-fodt-file-document-with-electronic-digital-signature-tutorial-by-ldigisigner-sign-a-word-sign-a-word/"><u>How to Sign .fodt file document with Electronic Digital Signature - (Tutorial)</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-messages-files-from-find-x6-by-fonelab-android-recover-messages/"><u>How To  Restore Missing Messages Files from Find X6</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-techniques-to-transfer-data-from-infinix-zero-30-5g-to-iphone-15141312-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Techniques to Transfer Data from Infinix Zero 30 5G to iPhone 15/14/13/12 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-mdm-from-iphone-15-plus-without-losing-data-by-drfone-ios-unlock-ios-unlock/"><u>How to Remove MDM from iPhone 15 Plus without losing data?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-zte-axon-40-lite-by-fonelab-android-recover-data/"><u>How to recover lost data from ZTE Axon 40 Lite?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-y78t-pictures-an-easy-method-explained-by-fonelab-android-recover-pictures/"><u>How to Restore Deleted Y78t Pictures  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-galaxy-a24-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Galaxy A24?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-mdm-from-iphone-xs-by-drfone-ios-unlock-ios-unlock/"><u>How to Remove MDM from iPhone XS?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-out-of-recovery-or-dfu-mode-on-iphone-12-pro-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Get Out of Recovery or DFU Mode on iPhone 12 Pro? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-honor-x50-gt-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Honor X50 GT</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-vivo-x100-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Vivo X100 to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-videos-from-note-30-by-fonelab-android-recover-video/"><u>How to retrieve erased videos from Note 30</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-12-proplus-5g-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of 12 Pro+ 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-videos-from-hot-40-by-fonelab-android-recover-video/"><u>How to Rescue Lost Videos from Hot 40</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-free-up-iphone-8-plus-space-drfone-by-drfone-ios-full-data-eraser-ios-full-data-eraser/"><u>How To Free Up iPhone 8 Plus Space | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-moto-g14-pictures-an-easy-method-explained-by-fonelab-android-recover-pictures/"><u>How to Restore Deleted Moto G14 Pictures  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-a70-with-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of A70 with Video Repair Utility on Mac?</u></a></li>
<li><a href="https://fake-location.techidaily.com/what-is-geo-blocking-and-how-to-bypass-it-on-sony-xperia-10-v-drfone-by-drfone-virtual-android/"><u>What is Geo-Blocking and How to Bypass it On Sony Xperia 10 V? | Dr.fone</u></a></li>
<li><a href="https://fake-location.techidaily.com/does-airplane-mode-turn-off-gps-location-on-vivo-y100i-drfone-by-drfone-virtual-android/"><u>Does Airplane Mode Turn off GPS Location On Vivo Y100i? | Dr.fone</u></a></li>
<li><a href="https://techidaily.com/how-to-factory-reset-honor-100-in-5-easy-ways-drfone-by-drfone-reset-android-reset-android/"><u>How to Factory Reset Honor 100 in 5 Easy Ways | Dr.fone</u></a></li>
<li><a href="https://apple-account.techidaily.com/how-to-remove-the-two-factor-authentication-from-apple-iphone-11-by-drfone-ios/"><u>How To Remove the Two Factor Authentication From Apple iPhone 11</u></a></li>
<li><a href="https://review-topics.techidaily.com/in-2024-can-life360-track-you-when-your-vivo-y78plus-is-off-drfone-by-drfone-virtual-android/"><u>In 2024, Can Life360 Track You When Your Vivo Y78+ is off? | Dr.fone</u></a></li>
<li><a href="https://bypass-frp.techidaily.com/in-2024-how-to-bypass-google-frp-lock-from-infinix-devices-by-drfone-android/"><u>In 2024, How to Bypass Google FRP Lock from Infinix Devices</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/in-2024-the-ultimate-guide-to-get-the-rare-candy-on-pokemon-go-fire-red-on-oppo-a38-drfone-by-drfone-virtual-android/"><u>In 2024, The Ultimate Guide to Get the Rare Candy on Pokemon Go Fire Red On Oppo A38 | Dr.fone</u></a></li>
<li><a href="https://android-frp.techidaily.com/in-2024-how-to-bypass-frp-on-oneplus-nord-n30-5g-by-drfone-android/"><u>In 2024, How to Bypass FRP on OnePlus Nord N30 5G?</u></a></li>
<li><a href="https://fix-guide.techidaily.com/how-to-resolve-xiaomi-civi-3-screen-not-working-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>How to Resolve Xiaomi Civi 3 Screen Not Working | Dr.fone</u></a></li>
<li><a href="https://android-unlock.techidaily.com/in-2024-the-ultimate-guide-how-to-bypass-swipe-screen-to-unlock-on-vivo-x100-pro-device-by-drfone-android/"><u>In 2024, The Ultimate Guide How to Bypass Swipe Screen to Unlock on Vivo X100 Pro Device</u></a></li>
<li><a href="https://location-social.techidaily.com/in-2024-set-your-preferred-job-location-on-linkedin-app-of-your-nokia-c12-plus-drfone-by-drfone-virtual-android/"><u>In 2024, Set Your Preferred Job Location on LinkedIn App of your Nokia C12 Plus | Dr.fone</u></a></li>
<li><a href="https://android-transfer.techidaily.com/how-to-transfer-text-messages-from-motorola-moto-g14-to-new-phone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Text Messages from Motorola Moto G14 to New Phone | Dr.fone</u></a></li>
<li><a href="https://ios-location-track.techidaily.com/in-2024-top-9-apple-iphone-14-monitoring-apps-for-parental-controls-drfone-by-drfone-virtual-ios/"><u>In 2024, Top 9 Apple iPhone 14 Monitoring Apps for Parental Controls | Dr.fone</u></a></li>
<li><a href="https://ai-video-translation.techidaily.com/updated-2024-approved-complete-guide-of-top-video-language-translators/"><u>Updated 2024 Approved Complete Guide of Top Video Language Translators</u></a></li>
<li><a href="https://ios-pokemon-go.techidaily.com/in-2024-how-can-i-get-more-stardust-in-pokemon-go-on-apple-iphone-13-pro-drfone-by-drfone-virtual-ios/"><u>In 2024, How can I get more stardust in pokemon go On Apple iPhone 13 Pro? | Dr.fone</u></a></li>
<li><a href="https://bypass-frp.techidaily.com/top-5-vivo-y27-5g-bypass-frp-tools-for-pc-that-actually-work-by-drfone-android/"><u>Top 5 Vivo Y27 5G Bypass FRP Tools for PC That Actually Work</u></a></li>
<li><a href="https://apple-account.techidaily.com/how-to-fix-when-apple-account-locked-from-iphone-13-pro-by-drfone-ios/"><u>How to Fix when Apple Account Locked From iPhone 13 Pro?</u></a></li>
<li><a href="https://change-location.techidaily.com/9-mind-blowing-tricks-to-hatch-eggs-in-pokemon-go-without-walking-on-vivo-x-flip-drfone-by-drfone-virtual-android/"><u>9 Mind-Blowing Tricks to Hatch Eggs in Pokemon Go Without Walking On Vivo X Flip | Dr.fone</u></a></li>
<li><a href="https://android-unlock.techidaily.com/in-2024-top-apps-and-online-tools-to-track-vivo-v27e-phone-withwithout-imei-number-by-drfone-android/"><u>In 2024, Top Apps and Online Tools To Track Vivo V27e Phone With/Without IMEI Number</u></a></li>
<li><a href="https://change-location.techidaily.com/a-working-guide-for-pachirisu-pokemon-go-map-on-samsung-galaxy-a05s-drfone-by-drfone-virtual-android/"><u>A Working Guide For Pachirisu Pokemon Go Map On Samsung Galaxy A05s | Dr.fone</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/in-2024-mastering-lock-screen-settings-how-to-enable-and-disable-on-realme-gt-neo-5-se-by-drfone-android/"><u>In 2024, Mastering Lock Screen Settings How to Enable and Disable on Realme GT Neo 5 SE</u></a></li>
<li><a href="https://ai-video-editing.techidaily.com/new-7-solutions-to-edit-gopro-video-on-mac-for-2024/"><u>New 7 Solutions to Edit GoPro Video on Mac for 2024</u></a></li>
<li><a href="https://ai-video-apps.techidaily.com/new-split-your-videos-with-ease-top-5-free-mpeg-splitters-for-2024/"><u>New Split Your Videos with Ease Top 5 Free MPEG Splitters for 2024</u></a></li>
</ul></div>



---
title: How to Fix Excel 2003 Run Time Error 1004
date: 2024-05-19T18:32:11.511Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2003 Run Time Error 1004
keywords: repair excel,repair excel 2013,repair .xltx
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
<li><a href="https://blog-min.techidaily.com/how-to-play-avchd-mts-files-on-defy-2-by-aiseesoft-video-converter-play-mts-on-android/"><u>How to play AVCHD MTS files on Defy 2?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-videos-from-your-google-pixel-8-by-fonelab-android-recover-video/"><u>How to recover old videos from your Google Pixel 8</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-motorola-razr-40-ultra-get-deleted-photos-back-with-ease-and-safety-by-fonelab-android-recover-photos/"><u>How to Motorola Razr 40 Ultra Get Deleted photos Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-nubia-red-magic-9-proplus-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Nubia Red Magic 9 Pro+</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-data-from-broken-iphone-xs-screen-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Data from Broken iPhone XS Screen | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-samsung-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Samsung on Windows??</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-data-from-iphone-15-using-stellar-data-recovery-for-iphone-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Data from iPhone 15 using Stellar Data Recovery for iPhone? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-asus-rog-phone-8-by-fonelab-android-recover-music/"><u>How to recover old music from your Asus ROG Phone 8</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-iphone-xs-max-ios-system-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair iPhone XS Max iOS System? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-your-iphone-11-pro-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Reset Your iPhone 11 Pro? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-samsung-galaxy-s21-fe-5g-2023-by-fonelab-android-recover-music/"><u>How to recover old music from your Samsung Galaxy S21 FE 5G (2023)</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-samsung-galaxy-s23plus-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Samsung Galaxy S23+?</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-lava-yuva-3-pro-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Lava Yuva 3 Pro to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-videos-from-honor-90-gt-by-fonelab-android-recover-video/"><u>How to retrieve erased videos from Honor 90 GT</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-xiaomi-redmi-note-12-pro-4g-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Xiaomi Redmi Note 12 Pro 4G to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-lava-yuva-2-by-fonelab-android-recover-music/"><u>How to restore wiped music on Lava Yuva 2</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-12-pro-max-to-factory-settings-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Reset iPhone 12 Pro Max to Factory Settings? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-contacts-files-from-lava-blaze-2-5g-by-fonelab-android-recover-contacts/"><u>How To  Restore Missing Contacts Files from Lava Blaze 2 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-contacts-files-from-spark-20-pro-by-fonelab-android-recover-contacts/"><u>How To  Restore Missing Contacts Files from Spark 20 Pro.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-infinix-note-30i-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Infinix Note 30i</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-manually-install-a-driver-in-windows-1110-by-drivereasy-guide/"><u>How to Manually Install a Driver in Windows 11/10</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-data-from-iphone-se-2020-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Data from iPhone SE (2020)? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-15-pro-max-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Reset iPhone 15 Pro Max? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-electronically-sign-a-dotx-file-using-digisigner-by-ldigisigner-sign-a-word-sign-a-word/"><u>How to Electronically Sign a .dotx file Using DigiSigner</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-pictures-files-from-14-ultra-by-fonelab-android-recover-pictures/"><u>How To  Restore Missing Pictures Files from 14 Ultra.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-out-of-dfu-mode-on-iphone-13-mini-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Get Out of DFU Mode on iPhone 13 mini? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-honor-x8b-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Honor X8b to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-videos-from-galaxy-a05-by-fonelab-android-recover-video/"><u>How to retrieve erased videos from Galaxy A05</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-nord-n30-5g-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Nord N30 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-iphone-15-pro-data-from-ios-itunes-backup-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How to Recover iPhone 15 Pro Data From iOS iTunes Backup? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-pictures-on-iphone-xs-5-best-solutions-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover deleted pictures on iPhone XS? 5 Best Solutions | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-excel-document-on-a-flash-drive-by-stellar-guide/"><u>How to Repair Excel Document on a Flash Drive</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-techniques-to-transfer-data-from-itel-p40-to-iphone-15141312-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Techniques to Transfer Data from Itel P40 to iPhone 15/14/13/12 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-call-history-on-tecno-by-fonelab-android-recover-call-logs/"><u>How to restore wiped call history on Tecno ?</u></a></li>
<li><a href="https://change-location.techidaily.com/how-to-fix-androidprocessmedia-has-stopped-on-oneplus-open-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>How to Fix Android.Process.Media Has Stopped on OnePlus Open | Dr.fone</u></a></li>
<li><a href="https://unlock-android.techidaily.com/can-i-bypass-a-forgotten-phone-password-of-xiaomi-by-drfone-android/"><u>Can I Bypass a Forgotten Phone Password Of Xiaomi?</u></a></li>
<li><a href="https://techidaily.com/how-to-perform-hard-reset-on-tecno-camon-20-drfone-by-drfone-reset-android-reset-android/"><u>How to Perform Hard Reset on Tecno Camon 20? | Dr.fone</u></a></li>
<li><a href="https://iphone-unlock.techidaily.com/full-guide-to-unlock-apple-iphone-12-pro-max-with-itunes-drfone-by-drfone-ios/"><u>Full Guide to Unlock Apple iPhone 12 Pro Max with iTunes | Dr.fone</u></a></li>
<li><a href="https://sim-unlock.techidaily.com/in-2024-how-to-change-your-sim-pin-code-on-your-vivo-x100-pro-phone-by-drfone-android/"><u>In 2024, How To Change Your SIM PIN Code on Your Vivo X100 Pro Phone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/fix-file-not-loaded-completely-error-in-excel-2019-by-stellar-guide/"><u>Fix File Not Loaded Completely Error in Excel 2019</u></a></li>
<li><a href="https://unlock-android.techidaily.com/pattern-locks-are-unsafe-secure-your-itel-p40-phone-now-with-these-tips-by-drfone-android/"><u>Pattern Locks Are Unsafe Secure Your Itel P40 Phone Now with These Tips</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/catchemall-celebrate-national-pokemon-day-with-virtual-location-on-oneplus-nord-n30-se-drfone-by-drfone-virtual-android/"><u>CatchEmAll Celebrate National Pokémon Day with Virtual Location On OnePlus Nord N30 SE | Dr.fone</u></a></li>
<li><a href="https://review-topics.techidaily.com/how-to-activate-and-use-life360-ghost-mode-on-vivo-y100i-drfone-by-drfone-virtual-android/"><u>How To Activate and Use Life360 Ghost Mode On Vivo Y100i | Dr.fone</u></a></li>
<li><a href="https://sim-unlock.techidaily.com/how-to-unlock-iphone-se-2022-online-here-are-6-easy-ways-by-drfone-ios/"><u>How to Unlock iPhone SE (2022) Online? Here are 6 Easy Ways</u></a></li>
<li><a href="https://ai-video-editing.techidaily.com/updated-here-we-are-going-to-show-you-the-best-effects-to-appear-on-camera-made-with-filmora-for-2024/"><u>Updated Here We Are Going to Show You the Best Effects to Appear on Camera Made with Filmora for 2024</u></a></li>
<li><a href="https://bypass-frp.techidaily.com/a-step-by-step-guide-on-using-adb-and-fastboot-to-remove-frp-lock-from-your-vivo-y77t-by-drfone-android/"><u>A Step-by-Step Guide on Using ADB and Fastboot to Remove FRP Lock from your Vivo Y77t</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/delete-gmail-account-withwithout-password-on-oneplus-12r-by-drfone-android/"><u>Delete Gmail Account With/Without Password On OnePlus 12R</u></a></li>
<li><a href="https://pokemon-go-android.techidaily.com/3-ways-for-android-pokemon-go-spoofing-on-poco-m6-5g-drfone-by-drfone-virtual-android/"><u>3 Ways for Android Pokemon Go Spoofing On Poco M6 5G | Dr.fone</u></a></li>
</ul></div>



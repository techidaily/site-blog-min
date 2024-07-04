---
title: How to Fix Excel 2010 Run Time Error 1004
date: 2024-05-19T18:32:11.529Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2010 Run Time Error 1004
keywords: repair excel 2013,repair excel 2023,repair excel 2003,repair .xlb,repair excel 2016,repair excel 2010
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
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-messages-files-from-tecno-spark-20-by-fonelab-android-recover-messages/"><u>How To  Restore Missing Messages Files from Tecno Spark 20</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-contacts-from-magic-5-lite-by-fonelab-android-recover-contacts/"><u>How to Rescue Lost Contacts from Magic 5 Lite?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-iphone-11-pro-data-from-icloud-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover iPhone 11 Pro Data From iCloud? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-infinix-note-30-pro-get-deleted-pictures-back-with-ease-and-safety-by-fonelab-android-recover-pictures/"><u>How to Infinix Note 30 Pro Get Deleted Pictures Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-videos-from-nokia-g310-by-fonelab-android-recover-video/"><u>How to retrieve erased videos from Nokia G310</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-oppo-a79-5g-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Oppo A79 5G to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-after-format-on-huawei-nova-y91-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery after format on Huawei Nova Y91</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-techniques-to-transfer-data-from-infinix-note-30-5g-to-iphone-15141312-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Techniques to Transfer Data from Infinix Note 30 5G to iPhone 15/14/13/12 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-vivo-v29-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Vivo V29 to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-data-from-iphone-12-pro-to-other-iphone-13-pro-max-devices-drfone-by-drfone-transfer-data-from-ios-transfer-data-from-ios/"><u>How To Transfer Data From iPhone 12 Pro To Other iPhone 13 Pro Max devices? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-motorola-g54-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Motorola G54 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-spark-20-pro-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Spark 20 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-pictures-from-edge-40-pro-by-fonelab-android-recover-pictures/"><u>How to Rescue Lost Pictures from Edge 40 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-iphone-13-pro-max-data-from-itunes-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover iPhone 13 Pro Max Data From iTunes? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-a-pdf-v14-document-with-electronic-signature-tool-by-ldigisigner-sign-a-pdf-sign-a-pdf/"><u>How to sign a PDF v1.4 document with electronic signature tool</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-deleted-photos-on-galaxy-s23-tactical-edition-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Retrieve deleted photos on Galaxy S23 Tactical Edition</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-lost-files-from-xiaomi-redmi-13c-by-fonelab-android-recover-data/"><u>How to retrieve lost files from Xiaomi Redmi 13C?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-solve-mkv-lagging-problem-in-galaxy-s23plus-by-aiseesoft-video-converter-play-mkv-on-android/"><u>How to solve MKV lagging problem in Galaxy S23+?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-files-after-iphone-13-pro-factory-reset-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Files after iPhone 13 Pro Factory Reset? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-data-from-iphone-12-pro-to-android-devices-drfone-by-drfone-transfer-data-from-ios-transfer-data-from-ios/"><u>How To Transfer Data From iPhone 12 Pro To Android devices? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-deleted-calendar-events-iphone-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Retrieve Deleted Calendar Events iPhone? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-iphone-14-ios-system-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair iPhone 14 iOS System? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-back-lost-photos-from-samsung-galaxy-f34-5g-by-fonelab-android-recover-photos/"><u>How to get back lost photos from Samsung Galaxy F34 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-nokia-c110-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Nokia C110</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-google-pixel-8-pro-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Google Pixel 8 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-pictures-files-from-realme-11x-5g-by-fonelab-android-recover-pictures/"><u>How To  Restore Missing Pictures Files from Realme 11X 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-play-hevc-h-265-video-on-redmi-12-5g-by-aiseesoft-video-converter-play-hevc-video-on-android/"><u>How to play HEVC H.265 video on Redmi 12 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-dotx-file-online-with-digisigner-by-ldigisigner-sign-a-word-sign-a-word/"><u>How to Sign .dotx file Online with DigiSigner</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-14-plus-to-factory-settings-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Reset iPhone 14 Plus to Factory Settings? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-iphone-15-pro-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data from iPhone 15 Pro? | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/complete-guide-for-recovering-contacts-files-on-infinix-hot-30i-by-fonelab-android-recover-contacts/"><u>Complete guide for recovering contacts files on Infinix Hot 30i.</u></a></li>
<li><a href="https://techidaily.com/how-to-hard-reset-google-pixel-8-pro-without-password-drfone-by-drfone-reset-android-reset-android/"><u>How to Hard Reset Google Pixel 8 Pro Without Password | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/can-i-recover-permanently-deleted-photos-from-redmi-note-12t-pro-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>Can I recover permanently deleted photos from Redmi Note 12T Pro</u></a></li>
<li><a href="https://sim-unlock.techidaily.com/top-imei-unlokers-for-your-vivo-y200e-5g-phone-by-drfone-android/"><u>Top IMEI Unlokers for Your Vivo Y200e 5G Phone</u></a></li>
<li><a href="https://techidaily.com/how-to-perform-hard-reset-on-honor-90-pro-drfone-by-drfone-reset-android-reset-android/"><u>How to Perform Hard Reset on Honor 90 Pro? | Dr.fone</u></a></li>
<li><a href="https://howto.techidaily.com/bricked-your-tecno-phantom-v-flip-heres-a-full-solution-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>Bricked Your Tecno Phantom V Flip? Heres A Full Solution | Dr.fone</u></a></li>
<li><a href="https://android-location.techidaily.com/in-2024-3-effective-methods-to-fake-gps-location-on-android-for-your-infinix-hot-30-5g-drfone-by-drfone-virtual/"><u>In 2024, 3 Effective Methods to Fake GPS location on Android For your Infinix Hot 30 5G | Dr.fone</u></a></li>
<li><a href="https://bypass-frp.techidaily.com/hassle-free-ways-to-remove-frp-lock-from-vivo-x100-pro-phones-withwithout-a-pc-by-drfone-android/"><u>Hassle-Free Ways to Remove FRP Lock from Vivo X100 Pro Phones with/without a PC</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/in-2024-ispoofer-is-not-working-on-apple-iphone-14-fixed-drfone-by-drfone-virtual-ios/"><u>In 2024, iSpoofer is not working On Apple iPhone 14? Fixed | Dr.fone</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/catchemall-celebrate-national-pokemon-day-with-virtual-location-on-xiaomi-redmi-k70-pro-drfone-by-drfone-virtual-android/"><u>CatchEmAll Celebrate National Pokémon Day with Virtual Location On Xiaomi Redmi K70 Pro | Dr.fone</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/full-tutorial-to-bypass-your-realme-note-50-face-lock-by-drfone-android/"><u>Full Tutorial to Bypass Your Realme Note 50 Face Lock?</u></a></li>
<li><a href="https://android-location.techidaily.com/9-best-free-android-monitoring-apps-to-monitor-phone-remotely-for-your-xiaomi-redmi-13c-drfone-by-drfone-virtual/"><u>9 Best Free Android Monitoring Apps to Monitor Phone Remotely For your Xiaomi Redmi 13C | Dr.fone</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/in-2024-recommended-best-applications-for-mirroring-your-oneplus-ace-3-screen-drfone-by-drfone-android/"><u>In 2024, Recommended Best Applications for Mirroring Your OnePlus Ace 3 Screen | Dr.fone</u></a></li>
</ul></div>



---
title: How to Fix Excel 2013 Run Time Error 1004
date: 2024-08-20T20:51:03.687Z
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

<!-- affiliate ads begin -->
<a href="https://store.iobit.com/order/checkout.php?PRODS=1468905&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/184260348236f9554fe9375772ff966e/ascscan_728x90.png" border="0"></a>
<!-- affiliate ads end -->
[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

VBA (Microsoft Visual Basic for Application) is an internal programming language in Microsoft Excel. Sometimes, when users try to run VBA or generate a Macro in Excel, the Run-time error 1004 may occur. This error may occur due to the presence of more legend entries in the chart, file conflict, incorrect Macro name, and corrupt Excel files. In this blog, we have discussed the reasons and shared some solutions to resolve run-time error 1004.

<!-- affiliate ads begin -->
<a href="https://printrendy.pxf.io/c/5597632/1453720/17020" target="_top" id="1453720"><img src="//a.impactradius-go.com/display-ad/17020-1453720" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1453720/17020" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
## **Why This Error Occurs?**

The run time error 1004 usually occurs when you run a VBA macro with the Legend Entries method to modify the legend entries in the MS Excel chart. It happens when the chart contains more legend entries than the available space, macro name conflicts, corrupt Excel files, or data-types mismatch in the VBA code.

<!-- affiliate ads begin -->
<a href="https://vapordna.pxf.io/c/5597632/1494880/17238" target="_top" id="1494880"><img src="//a.impactradius-go.com/display-ad/17238-1494880" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1494880/17238" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
## **Ways to Fix Excel Run-Time Error 1004?**

Try the below workarounds to fix Excel run-time error 1004:

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2068407/7443" target="_top" id="2068407"><img src="//a.impactradius-go.com/display-ad/7443-2068407" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2068407/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
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

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2082538/7443" target="_top" id="2082538"><img src="//a.impactradius-go.com/display-ad/7443-2082538" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2082538/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
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

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2075475/7443" target="_top" id="2075475"><img src="//a.impactradius-go.com/display-ad/7443-2075475" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2075475/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
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
<li><a href="https://youtube-web.techidaily.com/n-2024-easy-replay-on-iphone-set-videos-cycling/"><u>[New] In 2024, Easy Replay on iPhone  Set Videos Cycling</u></a></li>
<li><a href="https://fox-boxes.techidaily.com/updated-2024-approved-masterful-adjustments-the-editors-guide-to-gopro-video-magic/"><u>[Updated] 2024 Approved  Masterful Adjustments  The Editor's Guide to GoPro Video Magic</u></a></li>
<li><a href="https://tiktok-clips.techidaily.com/updated-crafting-prime-tiktok-openings-using-macos-for-2024/"><u>[Updated] Crafting Prime TikTok Openings Using MacOS for 2024</u></a></li>
<li><a href="https://fox-blue.techidaily.com/updated-wholesale-sky-shelves-economic-bulk-data-holder/"><u>[Updated] Wholesale Sky Shelves  Economic Bulk Data Holder</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-oppo-reno-10-pro-5g-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Oppo Reno 10 Pro 5G | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-nokia-c12-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Nokia C12 to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-poco-m6-pro-5g-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Poco M6 Pro 5G to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-techniques-to-transfer-data-from-itel-p55plus-to-iphone-15141312-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Techniques to Transfer Data from Itel P55+ to iPhone 15/14/13/12 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-ways-to-teach-you-to-transfer-files-from-tecno-spark-10-4g-to-other-android-devices-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Ways To Teach You To Transfer Files from Tecno Spark 10 4G to Other Android Devices Easily | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-ways-to-teach-you-to-transfer-files-from-vivo-s18e-to-other-android-devices-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Ways To Teach You To Transfer Files from Vivo S18e to Other Android Devices Easily | Dr.fone</u></a></li>
<li><a href="https://win-forum.techidaily.com/a-comprehensive-guide-to-leading-social-media-sites-facebook-twitter-instagram-and-youtube/"><u>A Comprehensive Guide to Leading Social Media Sites: Facebook, Twitter, Instagram & Youtube</u></a></li>
<li><a href="https://digital-screen-recording.techidaily.com/clearshot-vision-screen-recap-for-2024/"><u>ClearShot Vision Screen Recap for 2024</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-messages-files-from-vivo-y27-5g-by-fonelab-android-recover-messages/"><u>How To  Restore Missing Messages Files from Vivo Y27 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-infinix-hot-40-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Infinix Hot 40</u></a></li>
<li><a href="https://apple-account.techidaily.com/how-to-fix-apple-id-verification-code-not-working-on-iphone-xr-by-drfone-ios/"><u>How To Fix Apple ID Verification Code Not Working On iPhone XR</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-iphone-14-pro-max-storage-not-loadingshowing-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Fix iPhone 14 Pro Max Storage Not Loading/Showing | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-back-lost-photos-from-redmi-note-13-5g-by-fonelab-android-recover-photos/"><u>How to get back lost photos from Redmi Note 13 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-out-of-recovery-or-dfu-mode-on-iphone-11-pro-max-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Get Out of Recovery or DFU Mode on iPhone 11 Pro Max? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-insert-signature-in-xls-files-by-ldigisigner-sign-a-excel-sign-a-excel/"><u>How to insert signature in .xls files</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-install-the-latest-ios-beta-version-on-iphone-se-2020-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Install the Latest iOS Beta Version on iPhone SE (2020)? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-play-avchd-mts-files-on-13t-by-aiseesoft-video-converter-play-mts-on-android/"><u>How to play AVCHD MTS files on 13T?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-app-on-12-proplus-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Recover Deleted Photos from Android Gallery App on 12 Pro+ 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-app-on-realme-c55-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Recover Deleted Photos from Android Gallery App on Realme C55</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-app-on-smart-8-plus-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Recover Deleted Photos from Android Gallery App on Smart 8 Plus</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-tecno-spark-20-proplus-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Tecno Spark 20 Pro+</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-itel-a60-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Itel A60</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-oppo-reno-11-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Oppo Reno 11 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-vivo-y55s-5g-2023-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Vivo Y55s 5G (2023)</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-screenshots-on-iphone-6-plus-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Screenshots on iPhone 6 Plus? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-iphone-15-data-from-itunes-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover iPhone 15 Data From iTunes? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-edge-40-pro-by-fonelab-android-recover-data/"><u>How to recover lost data from Edge 40 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-whatsapp-chat-history-from-iphone-7-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How To Recover Whatsapp Chat History From iPhone 7 | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-s23plus-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of S23+ using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-iphone-se-2020-ios-system-issues-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair iPhone SE (2020) iOS System Issues? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-iphone-se-system-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair iPhone SE System? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-messages-from-s17-by-fonelab-android-recover-messages/"><u>How to Rescue Lost Messages from S17</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-motorola-edge-40-pro-by-fonelab-android-recover-music/"><u>How to restore wiped music on Motorola Edge 40 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-call-logs-from-vivo-y100-5g-by-fonelab-android-recover-call-logs/"><u>How to retrieve erased call logs from Vivo Y100 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-doc-file-document-with-electronic-digital-signature-tutorial-by-ldigisigner-sign-a-word-sign-a-word/"><u>How to Sign .doc file document with Electronic Digital Signature - (Tutorial)</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-jpeg-by-digital-signature-by-ldigisigner-sign-a-jpg-sign-a-jpg/"><u>How to sign .jpeg by digital signature</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-excel-2013-files-online-by-ldigisigner-sign-a-excel-sign-a-excel/"><u>How to sign Excel 2013 files online</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-data-from-iphone-12-mini-to-other-iphone-13-pro-devices-drfone-by-drfone-transfer-data-from-ios-transfer-data-from-ios/"><u>How To Transfer Data From iPhone 12 mini To Other iPhone 13 Pro devices? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-data-from-iphone-12-to-android-drfone-by-drfone-transfer-data-from-ios-transfer-data-from-ios/"><u>How To Transfer Data From iPhone 12 To Android? | Dr.fone</u></a></li>
<li><a href="https://fox-blue.techidaily.com/in-2024-high-res-aerial-showdown-dji-vs-gopro-capture/"><u>In 2024, High-Res Aerial Showdown  DJi vs GoPro Capture</u></a></li>
<li><a href="https://iphone-unlock.techidaily.com/in-2024-how-to-fix-auto-lock-greyed-out-on-apple-iphone-13-mini-drfone-by-drfone-ios/"><u>In 2024, How To Fix Auto Lock Greyed Out on Apple iPhone 13 mini | Dr.fone</u></a></li>
<li><a href="https://unlock-android.techidaily.com/in-2024-mastering-lock-screen-settings-how-to-enable-and-disable-on-itel-a60-by-drfone-android/"><u>In 2024, Mastering Lock Screen Settings How to Enable and Disable on Itel A60</u></a></li>
<li><a href="https://extra-lessons.techidaily.com/instructions-for-installing-windows-movie-maker-6/"><u>Instructions for Installing Windows Movie Maker 6</u></a></li>
</ul></div>

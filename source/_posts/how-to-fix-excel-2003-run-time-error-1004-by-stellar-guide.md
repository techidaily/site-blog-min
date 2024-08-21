---
title: How to Fix Excel 2003 Run Time Error 1004
date: 2024-08-20T20:51:14.145Z
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

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2068425/7443" target="_top" id="2068425"><img src="//a.impactradius-go.com/display-ad/7443-2068425" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2068425/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

VBA (Microsoft Visual Basic for Application) is an internal programming language in Microsoft Excel. Sometimes, when users try to run VBA or generate a Macro in Excel, the Run-time error 1004 may occur. This error may occur due to the presence of more legend entries in the chart, file conflict, incorrect Macro name, and corrupt Excel files. In this blog, we have discussed the reasons and shared some solutions to resolve run-time error 1004.

## **Why This Error Occurs?**

The run time error 1004 usually occurs when you run a VBA macro with the Legend Entries method to modify the legend entries in the MS Excel chart. It happens when the chart contains more legend entries than the available space, macro name conflicts, corrupt Excel files, or data-types mismatch in the VBA code.

<!-- affiliate ads begin -->
<a href="https://shop.copernic.com/order/checkout.php?PRODS=41033095&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.2checkout.com/images/merchant/8d30aa96e72440759f74bd2306c1fa3d/Copernic-2023-Affiliate-728x90-Advanced-3YR.png" border="0"></a>
<!-- affiliate ads end -->
## **Ways to Fix Excel Run-Time Error 1004?**

Try the below workarounds to fix Excel run-time error 1004:

<!-- affiliate ads begin -->
<a href="https://twopages.pxf.io/c/5597632/1873313/18544" target="_top" id="1873313"><img src="//a.impactradius-go.com/display-ad/18544-1873313" border="0" alt="" width="1080" height="1263"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1873313/18544" style="position:absolute;visibility:hidden;" border="0" />
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

![Program and Features of Control Panel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/09/program-and-features-1024x516.png)

- Search for **Microsoft Works** and click **Uninstall**.

<!-- affiliate ads begin -->
<a href="https://cowinaudio.pxf.io/c/5597632/1116855/13794" target="_top" id="1116855"><img src="//a.impactradius-go.com/display-ad/13794-1116855" border="0" alt="" width="767" height="285"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1116855/13794" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **Try Deleting GWXL97.Xla File**

The Add-ins files with .xla extension in MS-EXCEL is used to provide additional functionality to Excel spreadsheets. Sometimes, deleting the GWXL97.XLA file fixes the run-time error. Here are the steps to delete this file:

- Make sure you have an **Admins rights**, open the **Windows Explorer**
- Follow the Path C:\\Programs Files\\MSOffice\\Office\\XLSTART.
- Find and right-click on the **GWXL97.XLA** file
- Click **Delete**.

<!-- affiliate ads begin -->
<a href="https://store.revouninstaller.com/order/checkout.php?PRODS=28010250&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/4282ec8de8c9be897e7aff4aa231b1a4/336__280a.jpg" border="0"></a>
<!-- affiliate ads end -->
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
<a href="https://imp.i357552.net/c/5597632/863039/11832" target="_top" id="863039"><img src="//a.impactradius-go.com/display-ad/11832-863039" border="0" alt="" width="300" height="250"/></a>
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
<li><a href="https://buynow-info.techidaily.com/shooter-meets-saga-an-immersive-review-of-the-outer-worlds-sci-fi-fun-for-the-ages/"><u>'Shooter Meets Saga': An Immersive Review of 'The Outer Worlds' – Sci-Fi Fun for the Ages</u></a></li>
<li><a href="https://screen-activity-recording.techidaily.com/new-2024-approved-quick-guide-to-saving-screen-captures-in-windows-8/"><u>[New] 2024 Approved  Quick Guide to Saving Screen Captures in Windows 8</u></a></li>
<li><a href="https://youtube-web.techidaily.com/ed-a-guide-to-free-you-from-youtubes-extra-bar-width-for-2024/"><u>[Updated] A Guide to Free You From YouTube's Extra Bar Width for 2024</u></a></li>
<li><a href="https://facebook-video-footage.techidaily.com/updated-in-2024-cutting-edge-video-edits-without-spending-navigate-through-the-8-tools/"><u>[Updated] In 2024, Cutting-Edge Video Edits Without Spending  Navigate Through The 8 Tools</u></a></li>
<li><a href="https://extra-skills.techidaily.com/updated-premier-comedy-photo-stylist-tool/"><u>[Updated] Premier Comedy Photo Stylist Tool</u></a></li>
<li><a href="https://extra-skills.techidaily.com/2024-approved-premium-picks-cutting-edge-content-from-googloud/"><u>2024 Approved  Premium Picks  Cutting-Edge Content From GooGloud</u></a></li>
<li><a href="https://snapchat-videos.techidaily.com/2024-approved-quick-setup-installing-snapchat-on-your-mac/"><u>2024 Approved  Quick Setup  Installing Snapchat on Your Mac</u></a></li>
<li><a href="https://vp-tips.techidaily.com/2024-approved-stealth-in-social-media-going-unseen-with-live-feeds/"><u>2024 Approved  Stealth in Social Media  Going Unseen with Live Feeds</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-realme-11-proplus-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Realme 11 Pro+ to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-nokia-c300-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Nokia C300 to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-realme-11-pro-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Realme 11 Pro to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-techniques-to-transfer-data-from-oppo-a56s-5g-to-iphone-15141312-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Techniques to Transfer Data from Oppo A56s 5G to iPhone 15/14/13/12 | Dr.fone</u></a></li>
<li><a href="https://tech-recovery.techidaily.com/automating-social-media-success-how-to-set-up-scheduled-posts-on-twitter/"><u>Automating Social Media Success: How to Set Up Scheduled Posts on Twitter</u></a></li>
<li><a href="https://driver-download.techidaily.com/find-and-download-updated-mbox-asterisk-2-drivers-step-by-step-tutorial-for-windows-users/"><u>Find & Download Updated MBox Asterisk 2 Drivers: Step by Step Tutorial for Windows Users</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-call-logs-from-oppo-a18-by-fonelab-android-recover-call-logs/"><u>How To  Restore Missing Call Logs from Oppo A18</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-messages-files-from-realme-c51-by-fonelab-android-recover-messages/"><u>How To  Restore Missing Messages Files from Realme C51</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-pictures-files-from-infinix-hot-30-5g-by-fonelab-android-recover-pictures/"><u>How To  Restore Missing Pictures Files from Infinix Hot 30 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-pictures-files-from-tecno-spark-20-proplus-by-fonelab-android-recover-pictures/"><u>How To  Restore Missing Pictures Files from Tecno Spark 20 Pro+.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-7-to-an-older-ios-version-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 7 to an Older iOS Version? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-the-value-error-in-excel-stellar-by-stellar-guide/"><u>How to Fix the #Value! Error in Excel? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-back-lost-photos-from-infinix-hot-40i-by-fonelab-android-recover-photos/"><u>How to get back lost photos from Infinix Hot 40i.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-insert-signature-in-excel-2013-files-by-ldigisigner-sign-a-excel-sign-a-excel/"><u>How to insert signature in Excel 2013 files</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-play-hevc-h265-video-on-xiaomi-by-aiseesoft-video-converter-play-hevc-video-on-android/"><u>How to play HEVC H.265 video on Xiaomi ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-poco-m6-pro-4g-get-deleted-photos-back-with-ease-and-safety-by-fonelab-android-recover-photos/"><u>How to Poco M6 Pro 4G Get Deleted photos Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-app-on-infinix-note-30-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Recover Deleted Photos from Android Gallery App on Infinix Note 30 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-asus-rog-phone-7-by-fonelab-android-recover-photos/"><u>How to recover deleted photos from Asus ROG Phone 7.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-huawei-nova-y71-by-fonelab-android-recover-photos/"><u>How to recover deleted photos from Huawei Nova Y71.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-oneplus-nord-ce-3-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on OnePlus Nord CE 3 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-oppo-reno-11-pro-5g-by-fonelab-android-recover-data/"><u>How to recover lost data from Oppo Reno 11 Pro 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-mdm-from-iphone-12-without-a-computer-by-drfone-ios-unlock-ios-unlock/"><u>How to Remove MDM from iPhone 12 without a computer?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-a-damaged-video-file-of-vivo-y36i-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair a Damaged video file of Vivo Y36i?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-huawei-huawei-mate-60-proplus-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Huawei Huawei Mate 60 Pro+ using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-contacts-from-samsung-galaxy-a24-by-fonelab-android-recover-contacts/"><u>How to Rescue Lost Contacts from Samsung Galaxy A24?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-messages-from-realme-narzo-60-5g-by-fonelab-android-recover-messages/"><u>How to Rescue Lost Messages from Realme Narzo 60 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-6-plus-to-factory-settings-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Reset iPhone 6 Plus to Factory Settings? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-deleted-calendar-events-iphone-xr-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Retrieve Deleted Calendar Events iPhone XR? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-deleted-photos-on-motorola-moto-g34-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Retrieve deleted photos on Motorola Moto G34 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-lost-files-from-honor-play-8t-by-fonelab-android-recover-data/"><u>How to retrieve lost files from Honor Play 8T?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-docx-document-with-digital-signature-tutorial-by-ldigisigner-sign-a-word-sign-a-word/"><u>How to Sign .docx document with Digital Signature - (Tutorial)</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-docx-file-document-with-digital-signature-tutorial-by-ldigisigner-sign-a-word-sign-a-word/"><u>How to Sign .docx file document with Digital Signature - (Tutorial)</u></a></li>
<li><a href="https://screen-video-capture.techidaily.com/in-the-fast-lane-capturing-games-on-galaxy-devices/"><u>In the Fast Lane  Capturing Games on Galaxy Devices</u></a></li>
<li><a href="https://youtube-blog.techidaily.com/-lawful-to-screen-capture-youtube-content/"><u>Is It Lawful to Screen-Capture YouTube Content?</u></a></li>
<li><a href="https://extra-skills.techidaily.com/screencentral-exclusive-global-views-with-homeside-links-for-2024/"><u>ScreenCentral  Exclusive Global Views with Homeside Links for 2024</u></a></li>
<li><a href="https://buynow-reviews.techidaily.com/syma-x5c-rc-quadcopter-analysis-value-for-money-in-consumer-drones/"><u>SYMA X5C R/C Quadcopter Analysis: Value for Money in Consumer Drones</u></a></li>
</ul></div>

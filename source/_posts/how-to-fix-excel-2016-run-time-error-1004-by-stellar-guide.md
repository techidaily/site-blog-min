---
title: How to Fix Excel 2016 Run Time Error 1004
date: 2024-08-20T20:51:07.665Z
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

<!-- affiliate ads begin -->
<a href="https://ship7com.pxf.io/c/5597632/1509856/17634" target="_top" id="1509856"><img src="//a.impactradius-go.com/display-ad/17634-1509856" border="0" alt="" width="730" height="383"/></a>
<!-- affiliate ads end -->
[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

VBA (Microsoft Visual Basic for Application) is an internal programming language in Microsoft Excel. Sometimes, when users try to run VBA or generate a Macro in Excel, the Run-time error 1004 may occur. This error may occur due to the presence of more legend entries in the chart, file conflict, incorrect Macro name, and corrupt Excel files. In this blog, we have discussed the reasons and shared some solutions to resolve run-time error 1004.

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2082529/7443" target="_top" id="2082529"><img src="//a.impactradius-go.com/display-ad/7443-2082529" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2082529/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
## **Why This Error Occurs?**

The run time error 1004 usually occurs when you run a VBA macro with the Legend Entries method to modify the legend entries in the MS Excel chart. It happens when the chart contains more legend entries than the available space, macro name conflicts, corrupt Excel files, or data-types mismatch in the VBA code.

## **Ways to Fix Excel Run-Time Error 1004?**

Try the below workarounds to fix Excel run-time error 1004:

<!-- affiliate ads begin -->
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=3851655&QTY=1&AFFILIATE=108875&CART=1"><img src="http://www.aiseesoft.com/avangate/30p/banner.jpg" border="0"></a>
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
<a href="https://funwhole.sjv.io/c/5597632/1702887/17189" target="_top" id="1702887"><img src="//a.impactradius-go.com/display-ad/17189-1702887" border="0" alt="" width="1000" height="1000"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1702887/17189" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Program and Features of Control Panel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/09/program-and-features-1024x516.png)

- Search for **Microsoft Works** and click **Uninstall**.

<!-- affiliate ads begin -->
<a href="https://modlily.sjv.io/c/5597632/1997817/17059" target="_top" id="1997817"><img src="//a.impactradius-go.com/display-ad/17059-1997817" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1997817/17059" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
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

<!-- affiliate ads begin -->
<a href="https://united.elfm.net/c/5597632/748964/4704" target="_top" id="748964"><img src="//a.impactradius-go.com/display-ad/4704-748964" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://united.elfm.net/i/5597632/748964/4704" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
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
<li><a href="https://remote-screen-capture.techidaily.com/new-2024-approved-brief-steps-efficiently-archiving-gotomeeting-discussions/"><u>[New] 2024 Approved  Brief Steps  Efficiently Archiving GoToMeeting Discussions</u></a></li>
<li><a href="https://extra-lessons.techidaily.com/new-apple-excellence-iphones-best-no-cost-image-assemblers-and-layouts/"><u>[New] Apple Excellence – iPhone's Best No-Cost Image Assemblers & Layouts</u></a></li>
<li><a href="https://facebook-video-content.techidaily.com/new-comprehensive-guide-to-viewing-facebook-videos-on-apple-tv-devices/"><u>[New] Comprehensive Guide to Viewing Facebook Videos on Apple TV Devices</u></a></li>
<li><a href="https://screen-video-capture.techidaily.com/new-expertly-crafted-cloud-microphones/"><u>[New] Expertly-Crafted Cloud Microphones</u></a></li>
<li><a href="https://digital-screen-recording.techidaily.com/new-how-to-screen-record-on-lenovo-laptop-for-2024/"><u>[New] How to Screen Record on Lenovo Laptop for 2024</u></a></li>
<li><a href="https://vp-tips.techidaily.com/new-in-2024-ideal-environments-for-digital-streaming/"><u>[New] In 2024, Ideal Environments for Digital Streaming</u></a></li>
<li><a href="https://extra-skills.techidaily.com/new-integrating-gifs-into-your-iphone-life-with-ease/"><u>[New] Integrating GIFs Into Your iPhone Life with Ease</u></a></li>
<li><a href="https://screen-activity-recording.techidaily.com/updated-2024-approved-capture-the-moment-a-guide-to-superior-webcam-recorders/"><u>[Updated] 2024 Approved  Capture the Moment - A Guide to Superior WebCam Recorders</u></a></li>
<li><a href="https://article-knowledge.techidaily.com/updated-2024-approved-prime-choices-for-budget-conscious-consumers-4k-viewing/"><u>[Updated] 2024 Approved  Prime Choices for Budget-Conscious Consumers' 4K Viewing</u></a></li>
<li><a href="https://screen-capture.techidaily.com/updated-the-ultimate-guide-how-to-record-hulu-on-winmacmobile-for-2024/"><u>[Updated] The Ultimate Guide  How To Record Hulu On Win/Mac/Mobile for 2024</u></a></li>
<li><a href="https://facebook-video-content.techidaily.com/updated-understanding-the-evolution-of-facebook-videos-for-2024/"><u>[Updated] Understanding the Evolution of Facebook Videos for 2024</u></a></li>
<li><a href="https://facebook-video-content.techidaily.com/2024-approved-a-step-by-step-approach-to-ending-live-sharing-on-laptops-and-tablets/"><u>2024 Approved  A Step-by-Step Approach to Ending Live Sharing on Laptops and Tablets</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-techniques-to-transfer-data-from-oneplus-12-to-iphone-15141312-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Techniques to Transfer Data from OnePlus 12 to iPhone 15/14/13/12 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-ways-to-move-contacts-from-xiaomi-redmi-note-13-5g-to-iphone-131415-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Ways to Move Contacts From Xiaomi Redmi Note 13 5G to iPhone (13/14/15) | Dr.fone</u></a></li>
<li><a href="https://network-issues.techidaily.com/bypass-bug-swift-glitch-remedies-in-apex/"><u>Bypass Bug: Swift Glitch Remedies in Apex</u></a></li>
<li><a href="https://fake-location.techidaily.com/can-life360-track-you-when-your-infinix-hot-40i-is-off-drfone-by-drfone-virtual-android/"><u>Can Life360 Track You When Your Infinix Hot 40i is off? | Dr.fone</u></a></li>
<li><a href="https://driver-error.techidaily.com/driver-fix-wsd-device-on-windows-os/"><u>Driver Fix: WSD Device on Windows OS</u></a></li>
<li><a href="https://games-able.techidaily.com/efficiently-managing-your-switch-gaming-identity/"><u>Efficiently Managing Your Switch Gaming Identity</u></a></li>
<li><a href="https://android-unlock.techidaily.com/forgot-pattern-lock-heres-how-you-can-unlock-samsung-galaxy-a14-5g-pattern-lock-screen-by-drfone-android/"><u>Forgot Pattern Lock? Heres How You Can Unlock Samsung Galaxy A14 5G Pattern Lock Screen</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-call-logs-from-motorola-moto-g84-5g-by-fonelab-android-recover-call-logs/"><u>How To  Restore Missing Call Logs from Motorola Moto G84 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-call-logs-from-poco-x6-pro-by-fonelab-android-recover-call-logs/"><u>How To  Restore Missing Call Logs from Poco X6 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-contacts-files-from-oneplus-by-fonelab-android-recover-contacts/"><u>How To  Restore Missing Contacts Files from OnePlus .</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-8-without-data-loss-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 8 without Data Loss? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-error-1015-while-restoring-iphone-14-pro-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to fix error 1015 while restoring iPhone 14 Pro | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-the-value-error-in-excel-2016-by-stellar-guide/"><u>How to Fix the #Value! Error in Excel 2016?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-videos-not-playing-on-motorola-by-stellar-video-repair-mobile-video-repair/"><u>How to Fix Videos Not Playing on Motorola ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-back-lost-photos-from-motorola-razr-40-ultra-by-fonelab-android-recover-photos/"><u>How to get back lost photos from Motorola Razr 40 Ultra.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-out-of-dfu-mode-on-iphone-6s-plus-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Get Out of DFU Mode on iPhone 6s Plus? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-missing-or-malfunctioning-hardware-drivers-with-windows-device-manager-in-windows-11107-by-drivereasy-guide/"><u>How to identify missing or malfunctioning hardware drivers with Windows Device Manager in Windows 11/10/7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-some-outdated-your-hardware-drivers-with-windows-device-manager-in-windows-10-and-7-by-drivereasy-guide/"><u>How to identify some outdated your hardware drivers with Windows Device Manager in Windows 10 & 7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-make-a-digital-signature-for-txt-file-by-ldigisigner-sign-a-word-sign-a-word/"><u>How to make a digital signature for .txt file</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-oppo-find-x6-get-deleted-pictures-back-with-ease-and-safety-by-fonelab-android-recover-pictures/"><u>How to Oppo Find X6 Get Deleted Pictures Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-after-format-on-honor-magic-vs-2-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery after format on Honor Magic Vs 2</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-app-on-google-pixel-8-pro-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Recover Deleted Photos from Android Gallery App on Google Pixel 8 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-honor-play-7t-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Honor Play 7T</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-nubia-z50s-pro-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Nubia Z50S Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-a-damaged-video-file-of-vivo-g2-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair a Damaged video file of Vivo G2 using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-contacts-from-honor-magic5-ultimate-by-fonelab-android-recover-contacts/"><u>How to Rescue Lost Contacts from Honor Magic5 Ultimate?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-messages-from-xiaomi-civi-3-disney-100th-anniversary-edition-by-fonelab-android-recover-messages/"><u>How to Rescue Lost Messages from Xiaomi Civi 3 Disney 100th Anniversary Edition</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-photos-from-sony-xperia-10-v-by-fonelab-android-recover-photos/"><u>How to Rescue Lost Photos from Sony Xperia 10 V?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-videos-from-nokia-g310-by-fonelab-android-recover-video/"><u>How to Rescue Lost Videos from Nokia G310</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-honor-v-purse-photos-an-easy-method-explained-by-fonelab-android-recover-photos/"><u>How to Restore Deleted Honor V Purse Photos  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-videos-on-vivo-by-fonelab-android-recover-video/"><u>How to restore wiped videos on Vivo</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-deleted-photos-on-vivo-y02t-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Retrieve  deleted photos on Vivo Y02T</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-deleted-photos-on-vivo-y36-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Retrieve deleted photos on Vivo Y36</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-messages-from-y17s-by-fonelab-android-recover-messages/"><u>How to retrieve erased messages from Y17s</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-music-from-nubia-red-magic-8s-proplus-by-fonelab-android-recover-music/"><u>How to retrieve erased music from Nubia Red Magic 8S Pro+</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-excel-2023-online-with-digisigner-by-ldigisigner-sign-a-excel-sign-a-excel/"><u>How to Sign Excel 2023 Online with DigiSigner</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-data-from-iphone-12-pro-to-other-iphone-15-devices-drfone-by-drfone-transfer-data-from-ios-transfer-data-from-ios/"><u>How To Transfer Data From iPhone 12 Pro To Other iPhone 15 devices? | Dr.fone</u></a></li>
<li><a href="https://android-transfer.techidaily.com/how-to-transfer-photos-from-samsung-galaxy-m54-5g-to-new-android-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Photos from Samsung Galaxy M54 5G to New Android? | Dr.fone</u></a></li>
<li><a href="https://some-skills.techidaily.com/in-2024-uncovering-the-secrets-of-effortless-iphone-podcast-downloads/"><u>In 2024, Uncovering the Secrets of Effortless iPhone Podcast Downloads</u></a></li>
<li><a href="https://facebook-video-recording.techidaily.com/in-2024-vidcapture-fb-videos-on-the-go/"><u>In 2024, VidCapture  Fb Videos on the Go</u></a></li>
<li><a href="https://driver-install.techidaily.com/m2-to-windows-connector/"><u>M.2 To Windows Connector</u></a></li>
<li><a href="https://extra-tips.techidaily.com/supreme-storytelling-spanning-eight-movie-types/"><u>Supreme Storytelling Spanning Eight Movie Types</u></a></li>
<li><a href="https://snapchat-videos.techidaily.com/transforming-snaps-into-animated-characters-in-snapchat/"><u>Transforming Snaps Into Animated Characters in Snapchat</u></a></li>
<li><a href="https://visual-screen-recording.techidaily.com/vlc-recorder-functionality-check-for-2024/"><u>VLC Recorder  Functionality Check for 2024</u></a></li>
<li><a href="https://facebook-record-videos.techidaily.com/your-personalized-playlist-blueprint-for-youtube-for-2024/"><u>Your Personalized Playlist Blueprint for YouTube for 2024</u></a></li>
</ul></div>

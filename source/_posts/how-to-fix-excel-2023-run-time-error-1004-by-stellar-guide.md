---
title: How to Fix Excel 2023 Run Time Error 1004
date: 2024-08-20T20:50:49.858Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2023 Run Time Error 1004
keywords: repair excel 2000,repair excel 2019,repair .xls
---

## How to Fix Excel Run Time Error 1004

**Summary:** Run-time errors are windows-specific issues that occur while the program is running. This blog will teach you how to fix Excel run-time error 1004. In addition, you’ll learn about an Excel repair tool that can help fix the error 1004 if it occurs due to corruption in Excel files.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

VBA (Microsoft Visual Basic for Application) is an internal programming language in Microsoft Excel. Sometimes, when users try to run VBA or generate a Macro in Excel, the Run-time error 1004 may occur. This error may occur due to the presence of more legend entries in the chart, file conflict, incorrect Macro name, and corrupt Excel files. In this blog, we have discussed the reasons and shared some solutions to resolve run-time error 1004.

## **Why This Error Occurs?**

The run time error 1004 usually occurs when you run a VBA macro with the Legend Entries method to modify the legend entries in the MS Excel chart. It happens when the chart contains more legend entries than the available space, macro name conflicts, corrupt Excel files, or data-types mismatch in the VBA code.

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2075471/7443" target="_top" id="2075471"><img src="//a.impactradius-go.com/display-ad/7443-2075471" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2075471/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
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

<!-- affiliate ads begin -->
<a href="https://ship7com.pxf.io/c/5597632/1509856/17634" target="_top" id="1509856"><img src="//a.impactradius-go.com/display-ad/17634-1509856" border="0" alt="" width="730" height="383"/></a>
<!-- affiliate ads end -->
![Task Manager Window](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/09/task-manager-window.png)

- Click the **Process** tab, right-click on each program you want to close, and then click **End Task.**
- Stop all the running programs.
- Open the **Run** window and type **_appwiz.cpl_** to open the **Programs and Feature** window.

![Program and Features of Control Panel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/09/program-and-features-1024x516.png)

- Search for **Microsoft Works** and click **Uninstall**.

<!-- affiliate ads begin -->
<a href="https://atezr.pxf.io/c/5597632/2018605/18496" target="_top" id="2018605"><img src="//a.impactradius-go.com/display-ad/18496-2018605" border="0" alt="" width="798" height="807"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2018605/18496" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **Try Deleting GWXL97.Xla File**

The Add-ins files with .xla extension in MS-EXCEL is used to provide additional functionality to Excel spreadsheets. Sometimes, deleting the GWXL97.XLA file fixes the run-time error. Here are the steps to delete this file:

- Make sure you have an **Admins rights**, open the **Windows Explorer**
- Follow the Path C:\\Programs Files\\MSOffice\\Office\\XLSTART.
- Find and right-click on the **GWXL97.XLA** file
- Click **Delete**.

<!-- affiliate ads begin -->
<a href="https://tokenmetrics.sjv.io/c/5597632/1864921/20702" target="_top" id="1864921"><img src="//a.impactradius-go.com/display-ad/20702-1864921" border="0" alt="" width="1251" height="1042"/></a>
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

<!-- affiliate ads begin -->
<a href="https://arkmc.pxf.io/c/5597632/427477/5172" target="_top" id="427477"><img src="//a.impactradius-go.com/display-ad/5172-427477" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://arkmc.pxf.io/i/5597632/427477/5172" style="position:absolute;visibility:hidden;" border="0" />
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

<!-- affiliate ads begin -->
<a href="https://propmoneyinc.pxf.io/c/5597632/1803116/14559" target="_top" id="1803116"><img src="//a.impactradius-go.com/display-ad/14559-1803116" border="0" alt="" width="859" height="859"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1803116/14559" style="position:absolute;visibility:hidden;" border="0" />
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
<li><a href="https://screen-recording.techidaily.com/new-in-2024-mastering-clear-video-capture-with-obs-even-when-dark/"><u>[New] In 2024, Mastering Clear Video Capture with OBS, Even When Dark</u></a></li>
<li><a href="https://facebook-video-footage.techidaily.com/new-learn-youtube-live-streaming-with-easy-obs-guide/"><u>[New] Learn YouTube Live Streaming with Easy OBS Guide</u></a></li>
<li><a href="https://facebook-video-content.techidaily.com/new-the-blueprint-for-successful-biographies-top-tips-and-techniques-from-experts/"><u>[New] The Blueprint for Successful Biographies  Top Tips & Techniques From Experts</u></a></li>
<li><a href="https://facebook-video-share.techidaily.com/updated-earnings-escalation-via-youtube-channel-initiatives-for-2024/"><u>[Updated] Earnings Escalation via YouTube Channel Initiatives for 2024</u></a></li>
<li><a href="https://some-knowledge.techidaily.com/updated-frontiers-in-3d-visualization-tech/"><u>[Updated] Frontiers in 3D Visualization Tech</u></a></li>
<li><a href="https://digital-screen-recording.techidaily.com/updated-premier-mac-tools-beyond-bandicam/"><u>[Updated] Premier Mac Tools Beyond Bandicam</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-nokia-130-music-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Nokia 130 Music to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://youtube-clips.techidaily.com/2024-approved-dominating-the-digital-space-with-trending-content/"><u>2024 Approved  Dominating the Digital Space with Trending Content</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-ways-to-transfer-music-from-motorola-moto-g14-to-other-android-devices-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Ways to Transfer Music from Motorola Moto G14 to Other Android Devices Easily | Dr.fone</u></a></li>
<li><a href="https://fox-glue.techidaily.com/enhance-xbox-gameplay-with-best-monitor-recommendations/"><u>Enhance Xbox Gameplay with Best Monitor Recommendations</u></a></li>
<li><a href="https://win11-tips.techidaily.com/fixing-youtube-playback-speed-delays-in-chrome/"><u>Fixing YouTube Playback Speed Delays in Chrome</u></a></li>
<li><a href="https://extra-tips.techidaily.com/from-capture-to-creation-discover-the-best-montage-apps-for-smartphones/"><u>From Capture to Creation  Discover the Best Montage Apps for Smartphones</u></a></li>
<li><a href="https://facebook-video-files.techidaily.com/from-sky-to-screen-live-streaming-from-dji-drones-for-2024/"><u>From Sky to Screen  Live-Streaming From DJI Drones for 2024</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-messages-files-from-sony-xperia-1-v-by-fonelab-android-recover-messages/"><u>How To  Restore Missing Messages Files from Sony Xperia 1 V</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-pictures-files-from-nokia-105-classic-by-fonelab-android-recover-pictures/"><u>How To  Restore Missing Pictures Files from Nokia 105 Classic.</u></a></li>
<li><a href="https://tech-recovery.techidaily.com/how-to-correctly-address-and-fix-mingwm10dll-errors-on-windows-systems/"><u>How To Correctly Address and Fix Mingwm10.dll Errors on Windows Systems</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-corrupt-video-files-of-c67-4g-using-video-repair-utility-by-stellar-video-repair-mobile-video-repair/"><u>How to Fix Corrupt video files of C67 4G using Video Repair Utility?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-back-lost-contacts-from-100-pro-by-fonelab-android-recover-contacts/"><u>How to get back lost contacts from 100 Pro.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-install-hardware-device-drivers-manually-in-windows-7-by-drivereasy-guide/"><u>How to install hardware device drivers manually in Windows 7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-data-from-dead-iphone-13-pro-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to recover data from dead iPhone 13 Pro | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-contacts-from-motorola-edge-40-pro-by-fonelab-android-recover-contacts/"><u>How to recover deleted contacts from Motorola Edge 40 Pro.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-notes-from-iphone-11-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Notes from iPhone 11? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-videos-from-iphone-12-pro-without-backup-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Videos from iPhone 12 Pro Without Backup? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-iphone-6-plus-data-from-ios-itunes-backup-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How to Recover iPhone 6 Plus Data From iOS iTunes Backup? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-videos-from-your-nokia-g310-by-fonelab-android-recover-video/"><u>How to recover old videos from your Nokia G310</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-google-frp-lock-on-vivo-y78t-by-drfone-android-unlock-remove-google-frp/"><u>How to remove Google FRP Lock on Vivo Y78t</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-ios-system-issues-of-iphone-14-pro-max-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair iOS System Issues of iPhone 14 Pro Max? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-6-plus-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Reset iPhone 6 Plus? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-lava-blaze-2-pro-pictures-an-easy-method-explained-by-fonelab-android-recover-pictures/"><u>How to Restore Deleted Lava Blaze 2 Pro Pictures  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-poco-photos-an-easy-method-explained-by-fonelab-android-recover-photos/"><u>How to Restore Deleted Poco Photos  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-messages-from-motorola-razr-40-by-fonelab-android-recover-messages/"><u>How to retrieve erased messages from Motorola Razr 40</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-wps-file-free-by-ldigisigner-sign-a-word-sign-a-word/"><u>How to sign .wps file free</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-data-from-iphone-13-pro-to-other-iphone-13-devices-drfone-by-drfone-transfer-data-from-ios-transfer-data-from-ios/"><u>How To Transfer Data From iPhone 13 Pro To Other iPhone 13 devices? | Dr.fone</u></a></li>
<li><a href="https://android-transfer.techidaily.com/how-to-transfer-data-from-oppo-find-n3-flip-to-blackberry-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Data from Oppo Find N3 Flip to BlackBerry | Dr.fone</u></a></li>
<li><a href="https://sim-unlock.techidaily.com/in-2024-how-to-unlock-the-apple-iphone-6-sim-lock-4-easy-methods-by-drfone-ios/"><u>In 2024, How To Unlock The Apple iPhone 6 SIM Lock 4 Easy Methods</u></a></li>
<li><a href="https://location-social.techidaily.com/in-2024-set-your-preferred-job-location-on-linkedin-app-of-your-itel-a60s-drfone-by-drfone-virtual-android/"><u>In 2024, Set Your Preferred Job Location on LinkedIn App of your Itel A60s | Dr.fone</u></a></li>
<li><a href="https://sound-tweaking.techidaily.com/new-in-2024-the-complete-manual-of-garageband-audio-sessions/"><u>New In 2024, The Complete Manual of GarageBand Audio Sessions</u></a></li>
<li><a href="https://extra-resources.techidaily.com/unboxing-pureaudiofinder-a-look-into-the-latest-release/"><u>Unboxing PureAudioFinder  A Look Into the Latest Release</u></a></li>
<li><a href="https://youtube-video-recordings.techidaily.com/unveiling-secrets-to-extending-reach-with-youtube-lists/"><u>Unveiling Secrets to Extending Reach with YouTube Lists</u></a></li>
<li><a href="https://audio-shaping.techidaily.com/updated-best-windows-10-auditory-integration-software-compared-for-2024/"><u>Updated Best Windows 10 Auditory Integration Software Compared for 2024</u></a></li>
</ul></div>

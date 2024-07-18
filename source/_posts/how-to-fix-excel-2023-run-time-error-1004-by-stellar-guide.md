---
title: How to Fix Excel 2023 Run Time Error 1004
date: 2024-07-17T09:56:45.020Z
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
<li><a href="https://facebook-video-share.techidaily.com/new-chasing-channel-currency-the-updated-youtube-partner-guide-for-2024/"><u>[New] Chasing Channel Currency  The Updated Youtube Partner Guide for 2024</u></a></li>
<li><a href="https://screen-video-capture.techidaily.com/new-in-2024-must-have-offline-mobile-games-to-keep-you-entertained-on-android/"><u>[New] In 2024, Must-Have Offline Mobile Games to Keep You Entertained on Android</u></a></li>
<li><a href="https://facebook-video-recording.techidaily.com/new-transform-your-iphone-into-a-content-powerhouse-with-best-fb-video-downloaders-for-ios/"><u>[New] Transform Your iPhone Into a Content Powerhouse with Best FB Video Downloaders for iOS</u></a></li>
<li><a href="https://youtube-sure.techidaily.com/ed-2024-approved-educator-elite-selective-learning-yt-channels/"><u>[Updated] 2024 Approved  Educator Elite  Selective Learning YT Channels</u></a></li>
<li><a href="https://remote-screen-capture.techidaily.com/updated-2024-approved-rapid-capture-technique-timelapse-toptech/"><u>[Updated] 2024 Approved  Rapid Capture Technique  Timelapse #TopTech</u></a></li>
<li><a href="https://youtube-video-recordings.techidaily.com/updated-essential-screen-recording-software-list-for-vloggers/"><u>[Updated] Essential Screen Recording Software List for Vloggers</u></a></li>
<li><a href="https://instagram-videos.techidaily.com/updated-the-slippery-slope-of-simulated-support-instagram-style-for-2024/"><u>[Updated] The Slippery Slope of Simulated Support, Instagram Style for 2024</u></a></li>
<li><a href="https://some-approaches.techidaily.com/updated-unveiling-the-secrets-perfecting-motion-blur-in-photoshop/"><u>[Updated] Unveiling the Secrets  Perfecting Motion Blur in Photoshop</u></a></li>
<li><a href="https://tiktok-video-recordings.techidaily.com/2024-approved-how-to-boost-engagement-the-science-of-effective-video-captions-on-tiktok/"><u>2024 Approved  How to Boost Engagement  The Science of Effective Video Captions on TikTok</u></a></li>
<li><a href="https://some-skills.techidaily.com/2024-approved-the-ultimate-list-8-online-locations-for-free-3d-text-psdfiles/"><u>2024 Approved  The Ultimate List  8 Online Locations for Free 3D Text PSDFiles</u></a></li>
<li><a href="https://audio-editing.techidaily.com/2024-approved-unveiling-the-premier-list-of-cost-free-sound-intensity-modifiers/"><u>2024 Approved Unveiling the Premier List of Cost-Free Sound Intensity Modifiers</u></a></li>
<li><a href="https://facebook.techidaily.com/3-countries-where-you-cant-use-facebook/"><u>3 Countries Where You Can't Use Facebook</u></a></li>
<li><a href="https://location-fake.techidaily.com/8-solutions-to-fix-find-my-friends-location-not-available-on-nubia-red-magic-9-pro-drfone-by-drfone-virtual-android/"><u>8 Solutions to Fix Find My Friends Location Not Available On Nubia Red Magic 9 Pro | Dr.fone</u></a></li>
<li><a href="https://network-issues.techidaily.com/boost-pc-visuals-with-an-easy-to-execute-intel-graphics-driver-upgrade-on-w10/"><u>Boost PC Visuals with an Easy-to-Execute Intel Graphics Driver Upgrade on W10.</u></a></li>
<li><a href="https://extra-information.techidaily.com/cinematic-wonders-with-hero5-black-a-filmmakers-companion-for-2024/"><u>Cinematic Wonders with Hero5 Black  A Filmmaker's Companion for 2024</u></a></li>
<li><a href="https://audio-editing.techidaily.com/from-recording-to-broadcast-creating-professional-podcasts-in-audacity-for-2024/"><u>From Recording to Broadcast Creating Professional Podcasts in Audacity for 2024</u></a></li>
<li><a href="https://games-able.techidaily.com/hot-swap-mechanics-explored-in-razers-blackwidow-v4/"><u>Hot Swap Mechanics Explored in Razer’s BlackWidow V4</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-oppo-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Oppo</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-oppo-a1-5g-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Oppo A1 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-oppo-a18-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Oppo A18</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-oppo-a2-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Oppo A2</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-oppo-a58-4g-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Oppo A58 4G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-oppo-find-n3-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Oppo Find N3</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-oppo-find-n3-flip-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Oppo Find N3 Flip</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-oppo-find-x7-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Oppo Find X7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-oppo-k11-5g-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Oppo K11 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-oppo-k11x-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Oppo K11x</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-oppo-reno-10-proplus-5g-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Oppo Reno 10 Pro+ 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-oppo-reno-11-5g-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Oppo Reno 11 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-oppo-reno-9a-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Oppo Reno 9A</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-pixel-7a-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Pixel 7a</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-pixel-8-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Pixel 8</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-poco-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Poco</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-poco-f5-5g-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Poco F5 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-poco-m6-pro-4g-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Poco M6 Pro 4G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-poco-m6-pro-5g-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Poco M6 Pro 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-poco-x5-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Poco X5</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-poco-x5-pro-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Poco X5 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-poco-x6-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Poco X6</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-realme-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Realme</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-realme-c53-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Realme C53</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-realme-c67-4g-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Realme C67 4G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-realme-gt-5-240w-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Realme GT 5 (240W)</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-realme-gt-5-pro-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Realme GT 5 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-realme-narzo-60-5g-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Realme Narzo 60 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-realme-narzo-60x-5g-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Realme Narzo 60x 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-realme-narzo-n53-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Realme Narzo N53</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-realme-note-50-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Realme Note 50</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-redmi-12-5g-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Redmi 12 5G</u></a></li>
<li><a href="https://fake-location.techidaily.com/how-to-change-google-play-location-on-tecno-phantom-v-fold-drfone-by-drfone-virtual-android/"><u>How to Change Google Play Location On Tecno Phantom V Fold | Dr.fone</u></a></li>
<li><a href="https://location-social.techidaily.com/how-to-change-location-on-tiktok-to-see-more-content-on-your-oneplus-ace-2-pro-drfone-by-drfone-virtual-android/"><u>How to Change Location on TikTok to See More Content On your OnePlus Ace 2 Pro | Dr.fone</u></a></li>
<li><a href="https://pokemon-go-android.techidaily.com/in-2024-here-are-some-of-the-best-pokemon-discord-servers-to-join-on-realme-gt-neo-5-drfone-by-drfone-virtual-android/"><u>In 2024, Here are Some of the Best Pokemon Discord Servers to Join On Realme GT Neo 5 | Dr.fone</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/in-2024-top-10-fingerprint-lock-apps-to-lock-your-honor-x9b-phone-by-drfone-android/"><u>In 2024, Top 10 Fingerprint Lock Apps to Lock Your Honor X9b Phone</u></a></li>
<li><a href="https://screen-video-capture.techidaily.com/mastering-xbox-ultimate-screen-capture-tips-for-2024/"><u>Mastering Xbox  Ultimate Screen Capture Tips for 2024</u></a></li>
<li><a href="https://network-issues.techidaily.com/neutralizing-screen-volatility-pro-7-guide/"><u>Neutralizing Screen Volatility: Pro 7 Guide</u></a></li>
<li><a href="https://ai-driven-video-production.techidaily.com/new-besides-iphones-ipads-and-macbooks-apple-also-provides-some-fantastic-video-editing-software-here-is-a-full-list-of-the-apple-video-editing-software-try/"><u>New Besides iPhones, iPads and MacBooks, Apple Also Provides some Fantastic Video Editing Software. Here Is a Full List of the Apple Video Editing Software, Try It on Your iPhone or iPad or MacBooks Today</u></a></li>
<li><a href="https://screen-sharing-recording.techidaily.com/persistent-display-transcription-for-2024/"><u>Persistent Display Transcription for 2024</u></a></li>
<li><a href="https://tiktok-clips.techidaily.com/what-does-pfp-mean-in-tiktok-for-2024/"><u>What Does PFP Mean in TikTok for 2024</u></a></li>
<li><a href="https://techidaily.com/what-you-need-to-know-to-improve-your-realme-12-5g-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>What You Need To Know To Improve Your Realme 12 5G Hard Reset | Dr.fone</u></a></li>
</ul></div>

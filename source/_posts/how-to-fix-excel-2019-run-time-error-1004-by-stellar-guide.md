---
title: How to Fix Excel 2019 Run Time Error 1004
date: 2024-07-17T09:56:29.299Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2019 Run Time Error 1004
keywords: repair .xltx,repair excel 2003,repair .xlb,repair excel 2013,repair excel 2007
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

<!-- affiliate ads begin -->
<a href="https://caperobbin.sjv.io/c/5597632/2006118/18460" target="_top" id="2006118"><img src="//a.impactradius-go.com/display-ad/18460-2006118" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2006118/18460" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **Try Deleting GWXL97.Xla File**

The Add-ins files with .xla extension in MS-EXCEL is used to provide additional functionality to Excel spreadsheets. Sometimes, deleting the GWXL97.XLA file fixes the run-time error. Here are the steps to delete this file:

- Make sure you have an **Admins rights**, open the **Windows Explorer**
- Follow the Path C:\\Programs Files\\MSOffice\\Office\\XLSTART.
- Find and right-click on the **GWXL97.XLA** file
- Click **Delete**.

<!-- affiliate ads begin -->
<a href="https://store.bitdefender.com/affiliate.php?ACCOUNT=BITLATIN&AFFILIATE=108875&PATH=http%3A%2F%2Fwww.bitdefender.com%2Fbusiness%3FAFFILIATE%3D108875%26RESOURCE%3D30%2525%2BOff%2Ball%2BGravityZone%2BProducts"><img src="https://www.bitdefender.com/content/dam/bitdefender/business/campaign/1200X628.png" border="0"></a>
<!-- affiliate ads end -->
### **Change Trust Center Settings**

Sometimes, run-time errors might arise because of incorrect security settings. The **Trust Center settings** help you find the **Privacy and security** settings for Microsoft Excel. Follow the below steps to change the **Trust center settings**:

- Open Microsoft Excel.
- Go to **File > Options.**
- The **Excel options** window is displayed.
- Choose **Trust Center**, and click **Trust Center Settings**.
- Tap on the **Macro Settings** tab, and select **Trust access to the VBA project object model.**

<!-- affiliate ads begin -->
<a href="https://mindmanager.sjv.io/c/5597632/1787667/20231" target="_top" id="1787667"><img src="//a.impactradius-go.com/display-ad/20231-1787667" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1787667/20231" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Macro Settings in Microsoft Excel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/09/macro-settings.png)

- Click **OK**.

<!-- affiliate ads begin -->

<!-- affiliate ads end -->
### **Run Open and Repair Tool**

The Runtime error also arises when MS Excel detects a corrupted worksheet. It automatically begins the File recovery mode and starts repairing it. However, if the Recovery mode fails to start, use the **Open and Repair** tool with the below steps:

- Click **File > Open**.
- Click the location and folder with a corrupted workbook.
- In the **Open** dialog box, choose the corrupted workbook.
- Click the arrow next to the **Open** tab, and go to the **Open and Repair** tab.
- Click **Repair**.

You can also opt for **Stellar Repair for Excel** if the Microsoft Excel’s built-in tool cannot fix the error.

<!-- affiliate ads begin -->
<a href="https://tinyland.pxf.io/c/5597632/1793214/19135" target="_top" id="1793214"><img src="//a.impactradius-go.com/display-ad/19135-1793214" border="0" alt="" width="900" height="900"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1793214/19135" style="position:absolute;visibility:hidden;" border="0" />
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
<a href="https://laganoo.pxf.io/c/5597632/1657397/16446" target="_top" id="1657397"><img src="//a.impactradius-go.com/display-ad/16446-1657397" border="0" alt="" width="336" height="280"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1657397/16446" style="position:absolute;visibility:hidden;" border="0" />
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
<li><a href="https://facebook-videos.techidaily.com/updated-key-selections-top-6-fb-lite-video-grabs-for-2024/"><u>[Updated] Key Selections  Top 6 FB Lite Video Grabs for 2024</u></a></li>
<li><a href="https://facebook-clips.techidaily.com/updated-masterfb-mp4-the-ultimate-downloader-for-vids-for-2024/"><u>[Updated] MasterFB-MP4  The Ultimate Downloader for Vids for 2024</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-axon-40-lite-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Axon 40 Lite?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-civi-3-disney-100th-anniversary-edition-using-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Civi 3 Disney 100th Anniversary Edition using Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-defy-2-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Defy 2 using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-galaxy-a05-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Galaxy A05 using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-galaxy-f54-5g-using-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Galaxy F54 5G using Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-galaxy-s24plus-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Galaxy S24+?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-galaxy-xcover-6-pro-tactical-edition-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Galaxy XCover 6 Pro Tactical Edition?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-gionee-f3-pro-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Gionee F3 Pro using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-google-pixel-fold-using-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Google Pixel Fold using Video Repair Utility on Windows? </u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-gt-10-pro-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of GT 10 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-honor-magic-vs-2-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Honor Magic Vs 2 using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-honor-play-8t-using-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Honor Play 8T using Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-honor-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Honor using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-honor-using-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Honor using Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-honor-x50-gt-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Honor X50 GT using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-honor-x50-gt-using-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Honor X50 GT using Video Repair Utility on Windows? </u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-honor-x50-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Honor X50?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-honor-x50iplus-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Honor X50i+? </u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-honor-x9b-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Honor X9b using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-htc-using-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of HTC using Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-huawei-nova-12-pro-using-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Huawei Nova 12 Pro using Video Repair Utility on Windows? </u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-huawei-nova-y71-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Huawei Nova Y71 using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-infinix-hot-40-using-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Infinix Hot 40 using Video Repair Utility on Windows?</u></a></li>
<li><a href="https://techidaily.com/how-to-soft-reset-realme-gt-5-pro-phone-drfone-by-drfone-reset-android-reset-android/"><u>How to Soft Reset Realme GT 5 Pro phone? | Dr.fone</u></a></li>
<li><a href="https://android-transfer.techidaily.com/how-to-transfer-music-from-huawei-nova-y91-to-ipod-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Music from Huawei Nova Y91 to iPod | Dr.fone</u></a></li>
<li><a href="https://desktop-recording.techidaily.com/in-2024-powerpoint-presentations-enhancing-clarity-with-voiceovers/"><u>In 2024, PowerPoint Presentations  Enhancing Clarity with Voiceovers</u></a></li>
<li><a href="https://facebook-video-content.techidaily.com/locate-hidden-watch-video-preview/"><u>Locate Hidden Watch Video Preview</u></a></li>
<li><a href="https://change-location.techidaily.com/pokemon-go-cooldown-chart-on-vivo-y100t-drfone-by-drfone-virtual-android/"><u>Pokémon Go Cooldown Chart On Vivo Y100t | Dr.fone</u></a></li>
<li><a href="https://fox-links.techidaily.com/tackling-texts-and-gifs-an-experts-meme-making-manual-for-9gag-for-2024/"><u>Tackling Texts and Gifs  An Expert's Meme-Making Manual for 9GAG for 2024</u></a></li>
<li><a href="https://facebook-video-footage.techidaily.com/tripod-alignment-secrets-for-clearer-videos-for-2024/"><u>Tripod Alignment Secrets for Clearer Videos for 2024</u></a></li>
<li><a href="https://ai-driven-video-production.techidaily.com/updated-2024-approved-top-vignette-editing-apps-for-mobile-devices-free-and-premium-options/"><u>Updated 2024 Approved Top Vignette Editing Apps for Mobile Devices Free and Premium Options</u></a></li>
<li><a href="https://youtube-video-recordings.techidaily.com/youtubes-shorts-thumbnail-dilemnas-and-quick-fixes/"><u>YouTube's Shorts Thumbnail Dilemnas & Quick Fixes</u></a></li>
</ul></div>

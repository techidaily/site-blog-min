---
title: How to Fix Excel 2016 Run Time Error 1004
date: 2024-07-17T09:56:24.112Z
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
<li><a href="https://extra-hints.techidaily.com/new-androids-compendium-for-synchronized-and-curved-videography/"><u>[New] Android's Compendium for Synchronized & Curved Videography</u></a></li>
<li><a href="https://snapchat-videos.techidaily.com/new-beyond-basic-filters-elevating-your-snapchat-game-for-2024/"><u>[New] Beyond Basic Filters  Elevating Your Snapchat Game for 2024</u></a></li>
<li><a href="https://on-screen-recording.techidaily.com/new-in-2024-achieving-clip-perfection-with-blending-techniques/"><u>[New] In 2024, Achieving Clip Perfection with Blending Techniques</u></a></li>
<li><a href="https://visual-screen-recording.techidaily.com/updated-2024-approved-press-record-on-your-pc-with-simplicity/"><u>[Updated] 2024 Approved  Press 'Record' On Your PC with Simplicity</u></a></li>
<li><a href="https://article-files.techidaily.com/updated-2024-approved-sound-system-personalizations-on-ps-games/"><u>[Updated] 2024 Approved  Sound System Personalizations on PS Games</u></a></li>
<li><a href="https://eaxpv-info.techidaily.com/updated-hitting-the-high-scores-with-youtube-gaming-livestreams-for-2024/"><u>[Updated] Hitting the High Scores with YouTube Gaming Livestreams for 2024</u></a></li>
<li><a href="https://android-unlock.techidaily.com/5-solutions-for-samsung-galaxy-z-fold-5-unlock-without-password-by-drfone-android/"><u>5 Solutions For Samsung Galaxy Z Fold 5 Unlock Without Password</u></a></li>
<li><a href="https://sim-unlock.techidaily.com/easily-unlock-your-itel-p40-device-sim-by-drfone-android/"><u>Easily Unlock Your Itel P40 Device SIM</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-htc-u23-with-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of HTC U23 with Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-huawei-enjoy-70-with-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Huawei Enjoy 70 with Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-infinix-hot-30i-with-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Infinix Hot 30i with Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-infinix-note-30-pro-with-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Infinix Note 30 Pro with Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-infinix-note-30-pro-with-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Infinix Note 30 Pro with Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-infinix-note-30-vip-racing-edition-with-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Infinix Note 30 VIP Racing Edition with Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-infinix-with-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Infinix with Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-itel-a05s-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Itel A05s? </u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-itel-p40-with-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Itel P40 with Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-lava-blaze-2-5g-with-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Lava Blaze 2 5G with Video Repair Utility on Windows? </u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-lava-blaze-2-pro-with-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Lava Blaze 2 Pro with Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-magic-5-with-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Magic 5 with Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-magic-v2-with-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Magic V2 with Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-mix-fold-3-with-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Mix Fold 3 with Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-motorola-edge-2023-with-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Motorola Edge 2023 with Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-motorola-edge-40-pro-with-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Motorola Edge 40 Pro with Video Repair Utility on Windows? </u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-motorola-edgeplus-2023-with-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Motorola Edge+ (2023) with Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-motorola-g54-5g-with-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Motorola G54 5G with Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-motorola-moto-g-stylus-5g-2023-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Motorola Moto G Stylus 5G (2023)? </u></a></li>
<li><a href="https://android-location-track.techidaily.com/in-2024-5-ways-to-track-oppo-find-x6-pro-without-app-drfone-by-drfone-virtual-android/"><u>In 2024, 5 Ways to Track Oppo Find X6 Pro without App | Dr.fone</u></a></li>
<li><a href="https://extra-hints.techidaily.com/in-2024-bellylaugh-beats-top-choices-for-laughable-ringtone-downloads/"><u>In 2024, BellyLaugh Beats  Top Choices for Laughable Ringtone Downloads</u></a></li>
<li><a href="https://pokemon-go-android.techidaily.com/in-2024-how-and-where-to-find-a-shiny-stone-pokemon-for-realme-12-5g-drfone-by-drfone-virtual-android/"><u>In 2024, How and Where to Find a Shiny Stone Pokémon For Realme 12 5G? | Dr.fone</u></a></li>
<li><a href="https://android-unlock.techidaily.com/in-2024-how-to-remove-or-bypass-knox-enrollment-service-on-oppo-find-x7-ultra-by-drfone-android/"><u>In 2024, How To Remove or Bypass Knox Enrollment Service On Oppo Find X7 Ultra</u></a></li>
<li><a href="https://screen-activity-recording.techidaily.com/in-2024-switching-names-on-google-meet-for-laptops-phones/"><u>In 2024, Switching Names on Google Meet for Laptops, Phones</u></a></li>
<li><a href="https://some-skills.techidaily.com/in-2024-top-6-social-media-platforms-for-business/"><u>In 2024, Top 6 Social Media Platforms for Business</u></a></li>
<li><a href="https://extra-approaches.techidaily.com/palette-perfection-a-guide-to-grading-filmmaking-for-2024/"><u>Palette Perfection  A Guide to Grading Filmmaking for 2024</u></a></li>
<li><a href="https://network-issues.techidaily.com/reawakening-forgotten-cameras-on-your-asus-pc/"><u>Reawakening Forgotten Cameras on Your Asus PC</u></a></li>
</ul></div>

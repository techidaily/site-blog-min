---
title: How to Fix Excel 2021 Run Time Error 1004
date: 2024-07-17T09:56:37.218Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2021 Run Time Error 1004
keywords: repair excel 2000,repair excel 2003,repair excel 2023,repair .xlsm,repair .csv,repair excel 2010
---

## How to Fix Excel Run Time Error 1004

**Summary:** Run-time errors are windows-specific issues that occur while the program is running. This blog will teach you how to fix Excel run-time error 1004. In addition, you’ll learn about an Excel repair tool that can help fix the error 1004 if it occurs due to corruption in Excel files.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

VBA (Microsoft Visual Basic for Application) is an internal programming language in Microsoft Excel. Sometimes, when users try to run VBA or generate a Macro in Excel, the Run-time error 1004 may occur. This error may occur due to the presence of more legend entries in the chart, file conflict, incorrect Macro name, and corrupt Excel files. In this blog, we have discussed the reasons and shared some solutions to resolve run-time error 1004.

<!-- affiliate ads begin -->
<a href="https://united.elfm.net/c/5597632/517826/4704" target="_top" id="517826"><img src="//a.impactradius-go.com/display-ad/4704-517826" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://united.elfm.net/i/5597632/517826/4704" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
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

<!-- affiliate ads begin -->
<a href="https://propmoneyinc.pxf.io/c/5597632/1803116/14559" target="_top" id="1803116"><img src="//a.impactradius-go.com/display-ad/14559-1803116" border="0" alt="" width="859" height="859"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1803116/14559" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
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
<a href="https://natural-cycles.sjv.io/c/5597632/2072200/17885" target="_top" id="2072200"><img src="//a.impactradius-go.com/display-ad/17885-2072200" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2072200/17885" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **Try Deleting GWXL97.Xla File**

The Add-ins files with .xla extension in MS-EXCEL is used to provide additional functionality to Excel spreadsheets. Sometimes, deleting the GWXL97.XLA file fixes the run-time error. Here are the steps to delete this file:

- Make sure you have an **Admins rights**, open the **Windows Explorer**
- Follow the Path C:\\Programs Files\\MSOffice\\Office\\XLSTART.
- Find and right-click on the **GWXL97.XLA** file
- Click **Delete**.

<!-- affiliate ads begin -->
<a href="https://shop.mondly.com/affiliate.php?ACCOUNT=ATISTUDI&AFFILIATE=108875&PATH=https%3A%2F%2Fwww.mondly.com%3FAFFILIATE%3D108875%26RESOURCE%3D%2BEducational%2B970x90%2B"><img src="https://secure.avangate.com/images/merchant/69c418c33ec2e1a4267fa9bb77fa1428/educational-970x90.gif" border="0"></a>
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

<!-- affiliate ads begin -->
<a href="https://proteahair.pxf.io/c/5597632/1983634/23621" target="_top" id="1983634"><img src="//a.impactradius-go.com/display-ad/23621-1983634" border="0" alt="" width="320" height="100"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1983634/23621" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
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
<a href="https://unicoeye.pxf.io/c/5597632/2084396/18498" target="_top" id="2084396"><img src="//a.impactradius-go.com/display-ad/18498-2084396" border="0" alt="" width="1920" height="700"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2084396/18498" style="position:absolute;visibility:hidden;" border="0" />
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
<li><a href="https://facebook-video-files.techidaily.com/updated-2024-approved-essential-fixes-for-disconnected-fb-live-feeds/"><u>[Updated] 2024 Approved  Essential Fixes for Disconnected FB Live Feeds</u></a></li>
<li><a href="https://on-screen-recording.techidaily.com/updated-essential-tips-for-pc-based-console-game-recording/"><u>[Updated] Essential Tips for PC-Based Console Game Recording</u></a></li>
<li><a href="https://some-knowledge.techidaily.com/updated-expertly-selected-10-best-gimbals-for-smartphone-and-dslr-cameras/"><u>[Updated] Expertly Selected 10 Best Gimbals for Smartphone & DSLR Cameras</u></a></li>
<li><a href="https://extra-guidance.techidaily.com/updated-premier-text-manipulation-software-for-ae/"><u>[Updated] Premier Text Manipulation Software for AE</u></a></li>
<li><a href="https://extra-support.techidaily.com/updated-step-by-step-srt-to-xmlssattml-mastery-guide/"><u>[Updated] Step-by-Step  SRT to XML/SSA/TTML Mastery Guide</u></a></li>
<li><a href="https://youtube-web.techidaily.com/approved-peering-into-mr-beasts-economic-landscape/"><u>2024 Approved  Peering Into Mr. Beast's Economic Landscape</u></a></li>
<li><a href="https://ai-driven-video-production.techidaily.com/2024-approved-create-stunning-video-invites-best-apps-for-ios-and-android/"><u>2024 Approved Create Stunning Video Invites Best Apps for iOS and Android</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-meizu-21-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Meizu 21</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-moto-e13-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Moto E13</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-moto-g14-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Moto G14</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-moto-g84-5g-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Moto G84 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-motorola-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Motorola</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-motorola-edgeplus-2023-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Motorola Edge+ (2023)</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-motorola-moto-g-5g-2023-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Motorola Moto G 5G (2023)</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-motorola-moto-g-stylus-5g-2023-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Motorola Moto G Stylus 5G (2023)</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-motorola-moto-g73-5g-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Motorola Moto G73 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-motorola-razr-40-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Motorola Razr 40</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-narzo-n53-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Narzo N53</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-nokia-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Nokia</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-nokia-130-music-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Nokia 130 Music</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-nokia-c110-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Nokia C110</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-nokia-c12-pro-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Nokia C12 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-nokia-c210-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Nokia C210</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-note-30i-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Note 30i</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-note-50-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Note 50</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-nubia-red-magic-8s-proplus-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Nubia Red Magic 8S Pro+</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-nubia-z50-ultra-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Nubia Z50 Ultra</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-oppo-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Oppo</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-oppo-a1x-5g-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Oppo A1x 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-oppo-find-n3-flip-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Oppo Find N3 Flip</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-oppo-find-x6-pro-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Oppo Find X6 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-oppo-k11x-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Oppo K11x</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-oppo-reno-10-5g-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Oppo Reno 10 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-oppo-reno-10-pro-5g-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Oppo Reno 10 Pro 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-p55plus-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from P55+</u></a></li>
<li><a href="https://apple-account.techidaily.com/in-2024-apple-id-locked-or-disabled-on-iphone-xs-max-7-mehtods-you-cant-miss-by-drfone-ios/"><u>In 2024, Apple ID Locked or Disabled On iPhone XS Max? 7 Mehtods You Cant-Miss</u></a></li>
<li><a href="https://some-guidance.techidaily.com/in-2024-the-ultimate-guide-to-visualizing-audio-brilliance/"><u>In 2024, The Ultimate Guide to Visualizing Audio Brilliance</u></a></li>
<li><a href="https://audio-editing.techidaily.com/key-information-on-utilizing-youtubes-sound-bank/"><u>Key Information on Utilizing YouTubes Sound Bank</u></a></li>
<li><a href="https://pokemon-go-android.techidaily.com/the-most-useful-tips-for-pokemon-go-ultra-league-on-poco-f5-pro-5g-drfone-by-drfone-virtual-android/"><u>The Most Useful Tips for Pokemon Go Ultra League On Poco F5 Pro 5G | Dr.fone</u></a></li>
<li><a href="https://smart-video-creator.techidaily.com/updated-adobe-premiere-tutorial-importing-and-exporting-videos-like-a-pro-for-2024/"><u>Updated Adobe Premiere Tutorial Importing and Exporting Videos Like a Pro for 2024</u></a></li>
<li><a href="https://smart-video-creator.techidaily.com/updated-in-2024-best-video-loopers-to-loop-videos-for-free-on-windows-and-mac/"><u>Updated In 2024, Best Video Loopers to Loop Videos for Free on Windows and Mac</u></a></li>
</ul></div>

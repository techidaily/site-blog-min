---
title: How to Fix Excel 2007 Run Time Error 1004
date: 2024-07-17T09:56:02.853Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2007 Run Time Error 1004
keywords: repair .xlb,repair excel 2021,repair excel 2023,repair excel 2000
---

## How to Fix Excel Run Time Error 1004

**Summary:** Run-time errors are windows-specific issues that occur while the program is running. This blog will teach you how to fix Excel run-time error 1004. In addition, you’ll learn about an Excel repair tool that can help fix the error 1004 if it occurs due to corruption in Excel files.

<!-- affiliate ads begin -->
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=33729450&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/7f687767ccf20fcea1c9dc4a5adc2326/Digisigner_banner_728_x_90_color_version.png" border="0"></a>
<!-- affiliate ads end -->
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

<!-- affiliate ads begin -->
<a href="https://funwhole.sjv.io/c/5597632/1702887/17189" target="_top" id="1702887"><img src="//a.impactradius-go.com/display-ad/17189-1702887" border="0" alt="" width="1000" height="1000"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1702887/17189" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **Uninstall Microsoft Work**

You may encounter a run-time error 1004 in Excel version 2009 or older versions due to conflicts between Microsoft works and Microsoft Excel. This error usually occurs if your system has both Microsoft Office and Microsoft Works. Uninstalling one of them will fix the issue. Try the below steps to uninstall Microsoft Work:

- First, open the **Task Manager** using the shortcut **CTRL + ALT + DEL** altogether
- The **Task Manager window** is displayed.

![Task Manager Window](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/09/task-manager-window.png)

- Click the **Process** tab, right-click on each program you want to close, and then click **End Task.**
- Stop all the running programs.
- Open the **Run** window and type **_appwiz.cpl_** to open the **Programs and Feature** window.

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2075475/7443" target="_top" id="2075475"><img src="//a.impactradius-go.com/display-ad/7443-2075475" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2075475/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Program and Features of Control Panel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/09/program-and-features-1024x516.png)

- Search for **Microsoft Works** and click **Uninstall**.

### **Try Deleting GWXL97.Xla File**

The Add-ins files with .xla extension in MS-EXCEL is used to provide additional functionality to Excel spreadsheets. Sometimes, deleting the GWXL97.XLA file fixes the run-time error. Here are the steps to delete this file:

- Make sure you have an **Admins rights**, open the **Windows Explorer**
- Follow the Path C:\\Programs Files\\MSOffice\\Office\\XLSTART.
- Find and right-click on the **GWXL97.XLA** file
- Click **Delete**.

<!-- affiliate ads begin -->
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=3727260&QTY=1&AFFILIATE=108875&CART=1"><img src="http://www.aiseesoft.com/avangate/30p/banner.jpg" border="0"></a>
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
<a href="https://printrendy.pxf.io/c/5597632/1453720/17020" target="_top" id="1453720"><img src="//a.impactradius-go.com/display-ad/17020-1453720" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1453720/17020" style="position:absolute;visibility:hidden;" border="0" />
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
<a href="https://shop.mondly.com/affiliate.php?ACCOUNT=ATISTUDI&AFFILIATE=108875&PATH=https%3A%2F%2Fwww.mondly.com%3FAFFILIATE%3D108875%26RESOURCE%3D%2BEducational%2B970x90%2B"><img src="https://secure.avangate.com/images/merchant/69c418c33ec2e1a4267fa9bb77fa1428/educational-970x90.gif" border="0"></a>
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
<li><a href="https://tiktok-videos.techidaily.com/updated-2024-approved-2023s-must-have-tiktok-items-available-on-amazon/"><u>[Updated] 2024 Approved  2023'S Must-Have TikTok Items Available on Amazon</u></a></li>
<li><a href="https://facebook-video-recording.techidaily.com/updated-efficiently-change-your-facebook-videos-into-mp3-songs/"><u>[Updated] Efficiently Change Your Facebook Videos Into MP3 Songs</u></a></li>
<li><a href="https://some-knowledge.techidaily.com/updated-innovating-with-no-cost-text-animation-methods/"><u>[Updated] Innovating with No-Cost Text Animation Methods</u></a></li>
<li><a href="https://screen-video-capture.techidaily.com/2024-approved-streamline-your-windows-pc-to-record-live-tv-with-ease/"><u>2024 Approved  Streamline Your Windows PC to Record Live TV with Ease</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-y27-4g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Y27 4G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-y78-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Y78 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-zte-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on ZTE</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-zte-axon-40-lite-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on ZTE Axon 40 Lite</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-11-proplus-by-fonelab-android-recover-music/"><u>How to recover old music from your 11 Pro+</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-130-music-by-fonelab-android-recover-music/"><u>How to recover old music from your 130 Music</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-a38-by-fonelab-android-recover-music/"><u>How to recover old music from your A38</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-asus-rog-phone-7-by-fonelab-android-recover-music/"><u>How to recover old music from your Asus ROG Phone 7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-asus-rog-phone-8-by-fonelab-android-recover-music/"><u>How to recover old music from your Asus ROG Phone 8</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-c110-by-fonelab-android-recover-music/"><u>How to recover old music from your C110</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-g2-by-fonelab-android-recover-music/"><u>How to recover old music from your G2</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-galaxy-a05-by-fonelab-android-recover-music/"><u>How to recover old music from your Galaxy A05</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-galaxy-s24plus-by-fonelab-android-recover-music/"><u>How to recover old music from your Galaxy S24+</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-galaxy-z-fold-5-by-fonelab-android-recover-music/"><u>How to recover old music from your Galaxy Z Fold 5</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-gionee-f3-pro-by-fonelab-android-recover-music/"><u>How to recover old music from your Gionee F3 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-google-pixel-7a-by-fonelab-android-recover-music/"><u>How to recover old music from your Google Pixel 7a</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-google-pixel-fold-by-fonelab-android-recover-music/"><u>How to recover old music from your Google Pixel Fold</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-gt-5-240w-by-fonelab-android-recover-music/"><u>How to recover old music from your GT 5 (240W)</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-honor-90-lite-by-fonelab-android-recover-music/"><u>How to recover old music from your Honor 90 Lite</u></a></li>
<li><a href="https://android-unlock.techidaily.com/in-2024-can-i-bypass-a-forgotten-phone-password-of-samsung-galaxy-s23-tactical-edition-by-drfone-android/"><u>In 2024, Can I Bypass a Forgotten Phone Password Of Samsung Galaxy S23 Tactical Edition?</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/in-2024-downloading-samfw-frp-tool-30-for-realme-12-5g-by-drfone-android/"><u>In 2024, Downloading SamFw FRP Tool 3.0 for Realme 12 5G</u></a></li>
<li><a href="https://facebook-clips.techidaily.com/in-2024-essential-tips-for-hosting-a-successful-facebook-live-event/"><u>In 2024, Essential Tips for Hosting a Successful Facebook Live Event</u></a></li>
<li><a href="https://youtube-help.techidaily.com/in-2024-ranked-the-fastest-growing-youtube-platforms-for-encouragement/"><u>In 2024, Ranked  The Fastest Growing YouTube Platforms for Encouragement</u></a></li>
<li><a href="https://change-location.techidaily.com/in-2024-the-ultimate-guide-to-get-the-rare-candy-on-pokemon-go-fire-red-on-vivo-y100-5g-drfone-by-drfone-virtual-android/"><u>In 2024, The Ultimate Guide to Get the Rare Candy on Pokemon Go Fire Red On Vivo Y100 5G | Dr.fone</u></a></li>
<li><a href="https://fake-location.techidaily.com/life360-learn-how-everything-works-on-poco-c50-drfone-by-drfone-virtual-android/"><u>Life360 Learn How Everything Works On Poco C50 | Dr.fone</u></a></li>
<li><a href="https://visual-screen-recording.techidaily.com/pros-playbook-mastering-gameplay-logging-in-windows-10/"><u>Pro's Playbook  Mastering Gameplay Logging in Windows 10</u></a></li>
<li><a href="https://ai-video-editing.techidaily.com/this-article-aims-to-provide-a-step-by-step-guide-to-changing-a-videos-speed-by-using-the-speed-adjustment-panel-and-duration-panel-on-wondershare-filmora.m/"><u>This Article Aims to Provide a Step-by-Step Guide to Changing a Videos Speed by Using the Speed Adjustment Panel and Duration Panel on Wondershare Filmora</u></a></li>
<li><a href="https://smart-video-creator.techidaily.com/updated-2024-approved-the-ultimate-guide-to-minitool-movie-maker-features-pros-and-cons/"><u>Updated 2024 Approved The Ultimate Guide to Minitool Movie Maker Features, Pros, and Cons</u></a></li>
</ul></div>

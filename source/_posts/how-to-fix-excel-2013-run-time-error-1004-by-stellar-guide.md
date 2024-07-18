---
title: How to Fix Excel 2013 Run Time Error 1004
date: 2024-07-17T09:56:15.665Z
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
<li><a href="https://facebook-videos.techidaily.com/new-2024-approved-unleashing-shareability-secrets-of-trending-on-fb/"><u>[New] 2024 Approved  Unleashing Shareability  Secrets of Trending on Fb</u></a></li>
<li><a href="https://youtube-videos.techidaily.com/new-channel-hunt-best-historical-youtubes-that-every-student-should-watch/"><u>[New] Channel Hunt  Best Historical YouTubes That Every Student Should Watch</u></a></li>
<li><a href="https://fox-links.techidaily.com/updated-2024-approved-2023s-premier-storytelling-channels-on-youtube/"><u>[Updated] 2024 Approved  2023'S Premier Storytelling Channels on YouTube</u></a></li>
<li><a href="https://video-screen-grab.techidaily.com/updated-excellent-free-cam-snipping-software-for-2024/"><u>[Updated] Excellent Free Cam Snipping Software for 2024</u></a></li>
<li><a href="https://facebook-video-content.techidaily.com/updated-premier-media-transformer-mp4-)-fb/"><u>[Updated] Premier Media Transformer  MP4 -> FB</u></a></li>
<li><a href="https://facebook-videos.techidaily.com/updated-the-spectrum-of-sharing-navigating-video-distribution-on-fb/"><u>[Updated] The Spectrum of Sharing  Navigating Video Distribution on FB</u></a></li>
<li><a href="https://screen-sharing-recording.techidaily.com/updated-unwind-the-mind-the-ultimate-chill-list/"><u>[Updated] Unwind the Mind  The Ultimate Chill List</u></a></li>
<li><a href="https://extra-hints.techidaily.com/2024-approved-countdown-magic-discover-10-must-have-androidios-clocks/"><u>2024 Approved  Countdown Magic  Discover 10 Must-Have Android/iOS Clocks</u></a></li>
<li><a href="https://extra-support.techidaily.com/2024-approved-perfect-proportions-for-professional-videos/"><u>2024 Approved  Perfect Proportions for Professional Videos</u></a></li>
<li><a href="https://network-issues.techidaily.com/curing-black-displays-on-laptop-monitors/"><u>Curing Black Displays on Laptop Monitors</u></a></li>
<li><a href="https://unlock-android.techidaily.com/how-to-bypass-android-lock-screen-using-emergency-call-on-xiaomi-13-ultra-by-drfone-android/"><u>How to Bypass Android Lock Screen Using Emergency Call On Xiaomi 13 Ultra?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-asus-rog-phone-8-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Asus ROG Phone 8 on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-asus-rog-phone-8-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Asus ROG Phone 8 on Windows??</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-blaze-pro-5g-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Blaze Pro 5G on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-c51-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of C51 on Windows??</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-c55-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of C55 on Windows??</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-g22-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of G22?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-galaxy-m34-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Galaxy M34 on Windows??</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-galaxy-m54-5g-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Galaxy M54 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-galaxy-s23-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Galaxy S23 on Windows??</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-google-pixel-fold-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Google Pixel Fold on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-honor-100-pro-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Honor 100 Pro on Windows??</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-honor-100-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Honor 100?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-honor-90-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Honor 90?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-honor-magic-5-lite-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Honor Magic 5 Lite on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-honor-magic-5-pro-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Honor Magic 5 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-honor-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Honor on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-honor-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Honor on Windows??</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-honor-x50-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Honor X50 on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-honor-x50i-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Honor X50i on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-honor-x8b-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Honor X8b on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-hot-30i-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Hot 30i on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-htc-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of HTC on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-huawei-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Huawei ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-huawei-nova-11-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Huawei Nova 11?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-huawei-nova-y71-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Huawei Nova Y71?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-huawei-nova-y91-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Huawei Nova Y91 on Windows??</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-huawei-p60-art-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Huawei P60 Art on Windows??</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-huawei-p60-pro-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Huawei P60 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-infinix-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Infinix ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-infinix-hot-30i-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Infinix Hot 30i on Windows??</u></a></li>
<li><a href="https://voice-adjusting.techidaily.com/in-2024-10-best-free-online-audio-normalizers/"><u>In 2024, 10 Best Free Online Audio Normalizers</u></a></li>
<li><a href="https://pokemon-go-android.techidaily.com/in-2024-latest-way-to-get-shiny-meltan-box-in-pokemon-go-mystery-box-on-realme-12-proplus-5g-drfone-by-drfone-virtual-android/"><u>In 2024, Latest way to get Shiny Meltan Box in Pokémon Go Mystery Box On Realme 12 Pro+ 5G | Dr.fone</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/in-2024-overview-of-the-best-realme-gt-5-240w-screen-mirroring-app-drfone-by-drfone-android/"><u>In 2024, Overview of the Best Realme GT 5 (240W) Screen Mirroring App | Dr.fone</u></a></li>
<li><a href="https://video-capture.techidaily.com/in-2024-top-applications-for-remote-team-connection/"><u>In 2024, Top Applications for Remote Team Connection</u></a></li>
<li><a href="https://fake-location.techidaily.com/prank-your-friends-easy-ways-to-fake-and-share-google-maps-location-on-infinix-hot-30-5g-drfone-by-drfone-virtual-android/"><u>Prank Your Friends! Easy Ways to Fake and Share Google Maps Location On Infinix Hot 30 5G | Dr.fone</u></a></li>
<li><a href="https://eaxpv-info.techidaily.com/snippet-spectacle-top-hashes-to-trigger-the-viral-spree-on-youtube-for-2024/"><u>Snippet Spectacle  Top Hashes to Trigger the Viral Spree on YouTube for 2024</u></a></li>
</ul></div>

---
title: How to Fix Excel 2003 Run Time Error 1004
date: 2024-07-17T09:55:54.879Z
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
<li><a href="https://screen-activity-recording.techidaily.com/new-discover-11-leading-streamers-audio-recorders-for-2024/"><u>[New] Discover 11 Leading Streamers' Audio Recorders for 2024</u></a></li>
<li><a href="https://some-techniques.techidaily.com/new-identifying-high-impact-hdr-photography-tools/"><u>[New] Identifying High-Impact HDR Photography Tools</u></a></li>
<li><a href="https://tiktok-videos.techidaily.com/new-in-2024-enhance-video-purity-optimal-watermark-eliminators/"><u>[New] In 2024, Enhance Video Purity  Optimal Watermark Eliminators</u></a></li>
<li><a href="https://video-capture.techidaily.com/new-in-2024-secrets-of-effective-screen-capture-with-dell-devices/"><u>[New] In 2024, Secrets of Effective Screen Capture with Dell Devices</u></a></li>
<li><a href="https://some-skills.techidaily.com/new-top-rated-software-picks-audio-adjustment-for-virtual-personalities/"><u>[New] Top-Rated Software Picks  Audio Adjustment for Virtual Personalities</u></a></li>
<li><a href="https://screen-video-capture.techidaily.com/updated-unplugged-gaming-world-the-ultimate-list-of-best-free-mobile-apps-for-2024/"><u>[Updated] Unplugged Gaming World - The Ultimate List of Best Free Mobile Apps for 2024</u></a></li>
<li><a href="https://some-approaches.techidaily.com/2024-approved-structuring-panels-for-productive-interviews/"><u>2024 Approved  Structuring Panels for Productive Interviews</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/3-methods-to-mirror-lava-blaze-2-pro-to-roku-drfone-by-drfone-android/"><u>3 Methods to Mirror Lava Blaze 2 Pro to Roku | Dr.fone</u></a></li>
<li><a href="https://android-location-track.techidaily.com/how-to-detect-and-remove-spyware-on-google-pixel-7a-drfone-by-drfone-virtual-android/"><u>How to Detect and Remove Spyware on Google Pixel 7a? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-zte-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on ZTE</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-11-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on 11 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-11-proplus-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on 11 Pro+</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-11r-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on 11R</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-11x-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on 11X 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-12-pro-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on 12 Pro 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-12-proplus-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on 12 Pro+ 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-21-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on 21</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-90-lite-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on 90 Lite</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-a59-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on A59 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-asus-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Asus</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-asus-rog-phone-7-ultimate-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Asus ROG Phone 7 Ultimate</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-asus-rog-phone-8-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Asus ROG Phone 8</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-c02-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on C02</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-c110-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on C110</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-c12-plus-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on C12 Plus</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-c22-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on C22</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-c55-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on C55</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-camon-20-pro-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Camon 20 Pro 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-camon-30-pro-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Camon 30 Pro 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-f5-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on F5 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-find-n3-flip-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Find N3 Flip</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-g310-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on G310</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-galaxy-a24-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Galaxy A24</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-galaxy-a34-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Galaxy A34 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-galaxy-m14-4g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Galaxy M14 4G</u></a></li>
<li><a href="https://youtube-help.techidaily.com/in-2024-transforming-your-film-with-advanced-color-separation/"><u>In 2024, Transforming Your Film with Advanced Color Separation</u></a></li>
<li><a href="https://extra-tips.techidaily.com/internet-humorista-hub/"><u>Internet Humorista Hub</u></a></li>
<li><a href="https://sound-optimizing.techidaily.com/new-2024-approved-audiocraft-unveiled-understanding-key-traits-gathered-opinions-and-looking-beyond-avs-for-sound-mastery/"><u>New 2024 Approved Audiocraft Unveiled Understanding Key Traits, Gathered Opinions & Looking Beyond AVS for Sound Mastery</u></a></li>
<li><a href="https://extra-guidance.techidaily.com/pioneering-pathways-in-virtual-reality-biking-for-2024/"><u>Pioneering Pathways in Virtual Reality Biking for 2024</u></a></li>
<li><a href="https://youtube-videos.techidaily.com/sub4sub-unveiled-preliminary-insights-for-newcomers/"><u>Sub4sub Unveiled  Preliminary Insights for Newcomers</u></a></li>
</ul></div>

---
title: How to Fix Excel Run Time Error 1004
date: 2024-07-17T09:56:54.694Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel Run Time Error 1004
keywords: repair excel,repair .xls,repair excel 2007,repair excel 2021,repair .xltm
---

## How to Fix Excel Run Time Error 1004

**Summary:** Run-time errors are windows-specific issues that occur while the program is running. This blog will teach you how to fix Excel run-time error 1004. In addition, you’ll learn about an Excel repair tool that can help fix the error 1004 if it occurs due to corruption in Excel files.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

VBA (Microsoft Visual Basic for Application) is an internal programming language in Microsoft Excel. Sometimes, when users try to run VBA or generate a Macro in Excel, the Run-time error 1004 may occur. This error may occur due to the presence of more legend entries in the chart, file conflict, incorrect Macro name, and corrupt Excel files. In this blog, we have discussed the reasons and shared some solutions to resolve run-time error 1004.

<!-- affiliate ads begin -->
<a href="https://thefitville.pxf.io/c/5597632/1526796/15852" target="_top" id="1526796"><img src="//a.impactradius-go.com/display-ad/15852-1526796" border="0" alt="" width="1200" height="628"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1526796/15852" style="position:absolute;visibility:hidden;" border="0" />
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
<a href="https://boody-eco-wear.pxf.io/c/5597632/1572622/13846" target="_top" id="1572622"><img src="//a.impactradius-go.com/display-ad/13846-1572622" border="0" alt="" width="1000" height="1298"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1572622/13846" style="position:absolute;visibility:hidden;" border="0" />
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
<a href="https://twopages.pxf.io/c/5597632/2016067/18544" target="_top" id="2016067"><img src="//a.impactradius-go.com/display-ad/18544-2016067" border="0" alt="" width="1020" height="380"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2016067/18544" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Program and Features of Control Panel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/09/program-and-features-1024x516.png)

- Search for **Microsoft Works** and click **Uninstall**.

<!-- affiliate ads begin -->
<a href="https://propmoneyinc.pxf.io/c/5597632/1803115/14559" target="_top" id="1803115"><img src="//a.impactradius-go.com/display-ad/14559-1803115" border="0" alt="" width="859" height="859"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1803115/14559" style="position:absolute;visibility:hidden;" border="0" />
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

<!-- affiliate ads begin -->
<a href="https://ship7com.pxf.io/c/5597632/1509856/17634" target="_top" id="1509856"><img src="//a.impactradius-go.com/display-ad/17634-1509856" border="0" alt="" width="730" height="383"/></a>
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
<a href="https://twopages.pxf.io/c/5597632/1873305/18544" target="_top" id="1873305"><img src="//a.impactradius-go.com/display-ad/18544-1873305" border="0" alt="" width="1080" height="1350"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1873305/18544" style="position:absolute;visibility:hidden;" border="0" />
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
<li><a href="https://youtube-clips.techidaily.com/new-discovering-the-past-best-history-youtubers/"><u>[New] Discovering the Past  Best History YouTubers</u></a></li>
<li><a href="https://youtube-help.techidaily.com/new-expertly-design-your-3d-openers-best-picks/"><u>[New] Expertly Design Your 3D Openers  Best Picks</u></a></li>
<li><a href="https://tiktok-clips.techidaily.com/updated-in-2024-amazons-social-media-synergy-top-trending-finds-from-tiktok/"><u>[Updated] In 2024, Amazon’s Social Media Synergy - Top Trending Finds From TikTok</u></a></li>
<li><a href="https://instagram-video-recordings.techidaily.com/updated-pushing-the-envelope-on-instagram-larger-videos-explained/"><u>[Updated] Pushing the Envelope on Instagram  Larger Videos Explained</u></a></li>
<li><a href="https://video-screen-grab.techidaily.com/updated-simplified-methods-to-record-gotomeeting-chats-for-2024/"><u>[Updated] Simplified Methods to Record GoToMeeting Chats for 2024</u></a></li>
<li><a href="https://some-approaches.techidaily.com/updated-top-windows-10-tips-for-maximum-productivity/"><u>[Updated] Top Windows 10 Tips For Maximum Productivity</u></a></li>
<li><a href="https://tiktok-clips.techidaily.com/2024-approved-crafting-a-standout-identifier-for-your-tiktok-stream/"><u>2024 Approved  Crafting a Standout Identifier for Your TikTok Stream</u></a></li>
<li><a href="https://some-knowledge.techidaily.com/2024-approved-harmonize-your-status-with-a-track/"><u>2024 Approved  Harmonize Your Status with a Track</u></a></li>
<li><a href="https://audio-editing.techidaily.com/2024-approved-discover-the-10-premier-music-editing-apps-for-iphone-and-android/"><u>2024 Approved Discover the 10 Premier Music Editing Apps for iPhone and Android</u></a></li>
<li><a href="https://extra-information.techidaily.com/bringing-facial-movement-to-life-with-picsarts-effects/"><u>Bringing Facial Movement to Life with Picsart's Effects</u></a></li>
<li><a href="https://location-social.techidaily.com/how-to-hidefake-snapchat-location-on-your-sony-xperia-5-v-drfone-by-drfone-virtual-android/"><u>How to Hide/Fake Snapchat Location on Your Sony Xperia 5 V | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-samsung-galaxy-s24-by-fonelab-android-recover-music/"><u>How to restore wiped music on Samsung Galaxy S24</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-samsung-galaxy-s24plus-by-fonelab-android-recover-music/"><u>How to restore wiped music on Samsung Galaxy S24+</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-samsung-galaxy-xcover-7-by-fonelab-android-recover-music/"><u>How to restore wiped music on Samsung Galaxy XCover 7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-smart-7-hd-by-fonelab-android-recover-music/"><u>How to restore wiped music on Smart 7 HD</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-sony-xperia-10-v-by-fonelab-android-recover-music/"><u>How to restore wiped music on Sony Xperia 10 V</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-tecno-camon-20-premier-5g-by-fonelab-android-recover-music/"><u>How to restore wiped music on Tecno Camon 20 Premier 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-tecno-camon-20-pro-5g-by-fonelab-android-recover-music/"><u>How to restore wiped music on Tecno Camon 20 Pro 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-tecno-pova-5-by-fonelab-android-recover-music/"><u>How to restore wiped music on Tecno Pova 5</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-tecno-spark-10-4g-by-fonelab-android-recover-music/"><u>How to restore wiped music on Tecno Spark 10 4G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-v-purse-by-fonelab-android-recover-music/"><u>How to restore wiped music on V Purse</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-v30-by-fonelab-android-recover-music/"><u>How to restore wiped music on V30</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-vivo-by-fonelab-android-recover-music/"><u>How to restore wiped music on Vivo</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-vivo-v27-pro-by-fonelab-android-recover-music/"><u>How to restore wiped music on Vivo V27 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-vivo-y02t-by-fonelab-android-recover-music/"><u>How to restore wiped music on Vivo Y02T</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-vivo-y200-by-fonelab-android-recover-music/"><u>How to restore wiped music on Vivo Y200</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-vivo-y77t-by-fonelab-android-recover-music/"><u>How to restore wiped music on Vivo Y77t</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-xiaomi-by-fonelab-android-recover-music/"><u>How to restore wiped music on Xiaomi</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-xiaomi-redmi-k70-pro-by-fonelab-android-recover-music/"><u>How to restore wiped music on Xiaomi Redmi K70 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-xiaomi-redmi-k70e-by-fonelab-android-recover-music/"><u>How to restore wiped music on Xiaomi Redmi K70E</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-xiaomi-redmi-note-12-4g-by-fonelab-android-recover-music/"><u>How to restore wiped music on Xiaomi Redmi Note 12 4G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-xperia-5-v-by-fonelab-android-recover-music/"><u>How to restore wiped music on Xperia 5 V</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-y100i-by-fonelab-android-recover-music/"><u>How to restore wiped music on Y100i</u></a></li>
<li><a href="https://iphone-unlock.techidaily.com/in-2024-how-to-unlock-disabled-apple-iphone-xipad-without-computer-drfone-by-drfone-ios/"><u>In 2024, How to Unlock Disabled Apple iPhone X/iPad Without Computer | Dr.fone</u></a></li>
<li><a href="https://facebook-record-videos.techidaily.com/in-2024-youtubes-pinnacle-tools-studio-vs-next-gen-beta/"><u>In 2024, YouTube's Pinnacle Tools  Studio Vs. Next-Gen Beta</u></a></li>
<li><a href="https://graphic-issues.techidaily.com/intermittent-screen-problems-solved-for-laptop-success/"><u>Intermittent Screen Problems Solved for Laptop [Success]</u></a></li>
<li><a href="https://windows11.techidaily.com/keep-your-data-and-save-space-win11s-secure-drive-management-methods-max-156-chars/"><u>Keep Your Data and Save Space: Win11's Secure Drive Management Methods (Max 156 Chars)</u></a></li>
<li><a href="https://audio-shaping.techidaily.com/new-converting-audio-to-text-for-2024/"><u>New Converting Audio to Text for 2024</u></a></li>
<li><a href="https://sound-tweaking.techidaily.com/new-navigating-the-world-of-omegle-an-overview-and-safety-concerns/"><u>New Navigating the World of Omegle An Overview and Safety Concerns</u></a></li>
<li><a href="https://audio-shaping.techidaily.com/the-musicians-manual-to-individualizing-sound-files/"><u>The Musicians Manual to Individualizing Sound Files</u></a></li>
<li><a href="https://unlock-android.techidaily.com/the-ultimate-guide-how-to-bypass-swipe-screen-to-unlock-on-xiaomi-redmi-note-13-proplus-5g-device-by-drfone-android/"><u>The Ultimate Guide How to Bypass Swipe Screen to Unlock on Xiaomi Redmi Note 13 Pro+ 5G Device</u></a></li>
<li><a href="https://data-recovery.techidaily.com/ultimate-solution-for-various-data-extraction-programs-total-data-reconstruction-suite/"><u>Ultimate Solution for Various Data Extraction Programs - Total Data Reconstruction Suite</u></a></li>
<li><a href="https://smart-video-editing.techidaily.com/updated-in-2024-change-gif-speed-with-ease-best-online-and-mobile-tools/"><u>Updated In 2024, Change GIF Speed with Ease Best Online and Mobile Tools</u></a></li>
</ul></div>

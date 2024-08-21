---
title: How to Fix Excel 2000 Run Time Error 1004
date: 2024-08-20T20:51:10.328Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2000 Run Time Error 1004
keywords: repair excel 2023,repair excel 2021,repair .xlsm,repair excel 2019,repair excel 2016
---

## How to Fix Excel Run Time Error 1004

**Summary:** Run-time errors are windows-specific issues that occur while the program is running. This blog will teach you how to fix Excel run-time error 1004. In addition, you’ll learn about an Excel repair tool that can help fix the error 1004 if it occurs due to corruption in Excel files.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

VBA (Microsoft Visual Basic for Application) is an internal programming language in Microsoft Excel. Sometimes, when users try to run VBA or generate a Macro in Excel, the Run-time error 1004 may occur. This error may occur due to the presence of more legend entries in the chart, file conflict, incorrect Macro name, and corrupt Excel files. In this blog, we have discussed the reasons and shared some solutions to resolve run-time error 1004.

<!-- affiliate ads begin -->
<a href="https://imp.i110150.net/c/5597632/924299/11305" target="_top" id="924299"><img src="//a.impactradius-go.com/display-ad/11305-924299" border="0" alt="" width="520" height="100"/></a>
<!-- affiliate ads end -->
## **Why This Error Occurs?**

The run time error 1004 usually occurs when you run a VBA macro with the Legend Entries method to modify the legend entries in the MS Excel chart. It happens when the chart contains more legend entries than the available space, macro name conflicts, corrupt Excel files, or data-types mismatch in the VBA code.

<!-- affiliate ads begin -->
<a href="https://coinrule.sjv.io/c/5597632/1958379/18409" target="_top" id="1958379"><img src="//a.impactradius-go.com/display-ad/18409-1958379" border="0" alt="" width="856" height="508"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1958379/18409" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
## **Ways to Fix Excel Run-Time Error 1004?**

Try the below workarounds to fix Excel run-time error 1004:

<!-- affiliate ads begin -->
<a href="https://unicoeye.pxf.io/c/5597632/2084396/18498" target="_top" id="2084396"><img src="//a.impactradius-go.com/display-ad/18498-2084396" border="0" alt="" width="1920" height="700"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2084396/18498" style="position:absolute;visibility:hidden;" border="0" />
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

<!-- affiliate ads begin -->
<a href="https://ursime.pxf.io/c/5597632/2048963/16384" target="_top" id="2048963"><img src="//a.impactradius-go.com/display-ad/16384-2048963" border="0" alt="" width="1200" height="900"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2048963/16384" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Task Manager Window](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/09/task-manager-window.png)

- Click the **Process** tab, right-click on each program you want to close, and then click **End Task.**
- Stop all the running programs.
- Open the **Run** window and type **_appwiz.cpl_** to open the **Programs and Feature** window.

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

<!-- affiliate ads begin -->
<a href="https://united.elfm.net/c/5597632/748964/4704" target="_top" id="748964"><img src="//a.impactradius-go.com/display-ad/4704-748964" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://united.elfm.net/i/5597632/748964/4704" style="position:absolute;visibility:hidden;" border="0" />
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
<a href="https://shop.copernic.com/order/checkout.php?PRODS=41033091&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.2checkout.com/images/merchant/8d30aa96e72440759f74bd2306c1fa3d/Copernic-2023-Affiliate-728x90-Advanced.png" border="0"></a>
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
<li><a href="https://instagram-clips.techidaily.com/new-2024-approved-building-immersive-thumbnails-for-igtv-videos/"><u>[New] 2024 Approved  Building Immersive Thumbnails for IGTV Videos</u></a></li>
<li><a href="https://extra-skills.techidaily.com/new-onlineoffline-animation-revolution-with-zero-price-tag/"><u>[New] Online/Offline Animation Revolution with Zero Price Tag</u></a></li>
<li><a href="https://vp-tips.techidaily.com/updated-2024-approved-metaverse-vs-multimetaverse-key-distinctions-and-comparison/"><u>[Updated] 2024 Approved  Metaverse Vs. Multimetaverse  Key Distinctions and Comparison</u></a></li>
<li><a href="https://on-screen-recording.techidaily.com/updated-in-2024-harnessing-your-creative-power-with-ipad-for-time-lapses/"><u>[Updated] In 2024, Harnessing Your Creative Power with iPad for Time-Lapses</u></a></li>
<li><a href="https://youtube-zero.techidaily.com/ed-in-2024-jumpstart-your-channelstopwatch-seeking-viewers-with-these-youtube-seo-essentials/"><u>[Updated] In 2024, Jumpstart Your Channel’stopwatch-Seeking Viewers with These YouTube SEO Essentials</u></a></li>
<li><a href="https://article-posts.techidaily.com/updated-in-2024-sj4000-owners-guide-to-stylish-and-practical-gear/"><u>[Updated] In 2024, SJ4000 Owners' Guide to Stylish & Practical Gear</u></a></li>
<li><a href="https://instagram-clips.techidaily.com/updated-seamless-streams-addressing-instagrams-video-snags-for-2024/"><u>[Updated] Seamless Streams  Addressing Instagram's Video Snags for 2024</u></a></li>
<li><a href="https://fox-http.techidaily.com/2024-approved-chime-into-chats-whatsapp-status-soundtracks/"><u>2024 Approved  Chime Into Chats  WhatsApp Status Soundtracks</u></a></li>
<li><a href="https://screen-recording.techidaily.com/2024-approved-eagle-eye-expedition-the-ultimate-guide-to-fps-game-selection/"><u>2024 Approved  Eagle-Eye Expedition - The Ultimate Guide to FPS Game Selection</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-infinix-hot-40-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Infinix Hot 40 to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-vivo-y100t-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Vivo Y100t to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-xiaomi-civi-3-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Xiaomi Civi 3 to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-ways-to-move-contacts-from-infinix-smart-8-plus-to-iphone-131415-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Ways to Move Contacts From Infinix Smart 8 Plus to iPhone (13/14/15) | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-ways-to-teach-you-to-transfer-files-from-nubia-red-magic-8s-pro-to-other-android-devices-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Ways To Teach You To Transfer Files from Nubia Red Magic 8S Pro to Other Android Devices Easily | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-ways-to-teach-you-to-transfer-files-from-xiaomi-redmi-note-12-proplus-5g-to-other-android-devices-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Ways To Teach You To Transfer Files from Xiaomi Redmi Note 12 Pro+ 5G to Other Android Devices Easily | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-ways-to-transfer-music-from-itel-p55-5g-to-other-android-devices-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Ways to Transfer Music from Itel P55 5G to Other Android Devices Easily | Dr.fone</u></a></li>
<li><a href="https://extra-resources.techidaily.com/best-balanced-tripods-for-dslr-phone-and-tablet-videos/"><u>Best Balanced Tripods for DSLR, Phone, and Tablet Videos</u></a></li>
<li><a href="https://games-able.techidaily.com/browsing-through-adventures-text-based-wonders/"><u>Browsing Through Adventures - Text-Based Wonders</u></a></li>
<li><a href="https://buynow-reviews.techidaily.com/dive-into-the-world-of-ultrabooks-with-our-gorgeous-and-versatile-dell-xps-cu13-7390-review-a-hybrid-laptop-masterpiece-unveiled/"><u>Dive Into the World of Ultrabooks with Our Gorgeous and Versatile Dell XPS Cu13 7390 Review: A Hybrid Laptop Masterpiece Unveiled</u></a></li>
<li><a href="https://facebook-clips.techidaily.com/elevate-your-content-with-expertly-crafted-360-videos-for-fb-for-2024/"><u>Elevate Your Content with Expertly Crafted 360 Videos for FB for 2024</u></a></li>
<li><a href="https://extra-information.techidaily.com/essential-polarr-tips-for-pro-level-photography-editing/"><u>Essential Polarr Tips for Pro-Level Photography Editing</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-call-logs-from-tecno-spark-10-pro-by-fonelab-android-recover-call-logs/"><u>How To  Restore Missing Call Logs from Tecno Spark 10 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-messages-files-from-lava-blaze-pro-5g-by-fonelab-android-recover-messages/"><u>How To  Restore Missing Messages Files from Lava Blaze Pro 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-photos-files-from-gionee-f3-pro-by-fonelab-android-recover-photos/"><u>How To  Restore Missing Photos Files from Gionee F3 Pro.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-pictures-files-from-itel-p55-5g-by-fonelab-android-recover-pictures/"><u>How To  Restore Missing Pictures Files from Itel P55 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-14-pro-without-losing-anything-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 14 Pro without Losing Anything? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-exit-dfu-mode-on-iphone-13-pro-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Exit DFU Mode on iPhone 13 Pro? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-excel-formulas-not-working-properly-step-by-step-guide-stellar-by-stellar-guide/"><u>How to Fix Excel Formulas Not Working Properly | Step-by-Step Guide | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-microsoft-excel-2003-has-stopped-working-error-stellar-by-stellar-guide/"><u>How to fix Microsoft Excel 2003 has stopped working error? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-videos-not-playing-with-my-honor-100-by-stellar-video-repair-mobile-video-repair/"><u>How to fix videos not playing with my Honor 100?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-back-lost-photos-from-oppo-reno-11f-5g-by-fonelab-android-recover-photos/"><u>How to get back lost photos from Oppo Reno 11F 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-install-the-latest-iosipados-beta-version-on-iphone-8-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Install the Latest iOS/iPadOS Beta Version on iPhone 8? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-motorola-get-deleted-pictures-back-with-ease-and-safety-by-fonelab-android-recover-pictures/"><u>How to Motorola Get Deleted Pictures Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-after-format-on-nord-ce-3-lite-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery after format on Nord CE 3 Lite 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-iphone-14-pro-max-data-from-ios-itunes-backup-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How to Recover iPhone 14 Pro Max Data From iOS iTunes Backup? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-infinix-smart-8-by-fonelab-android-recover-data/"><u>How to recover lost data from Infinix Smart 8?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-of-iphone-11-pro-max-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data of iPhone 11 Pro Max? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-infinix-note-30-vip-by-fonelab-android-recover-music/"><u>How to recover old music from your Infinix Note 30 VIP</u></a></li>
<li><a href="https://tech-savvy.techidaily.com/how-to-reduce-cpu-usage-in-windows/"><u>How to Reduce CPU Usage in Windows</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-realme-with-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Realme with Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-messages-from-xiaomi-civi-3-by-fonelab-android-recover-messages/"><u>How to Rescue Lost Messages from Xiaomi Civi 3</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-photos-from-oppo-k11x-by-fonelab-android-recover-photos/"><u>How to Rescue Lost Photos from Oppo K11x?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-honor-magic-v2-photos-an-easy-method-explained-by-fonelab-android-recover-photos/"><u>How to Restore Deleted Honor Magic V2 Photos  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-tecno-pova-5-contacts-an-easy-method-explained-by-fonelab-android-recover-contacts/"><u>How to Restore Deleted Tecno Pova 5 Contacts  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-call-history-on-motorola-edgeplus-2023-by-fonelab-android-recover-call-logs/"><u>How to restore wiped call history on Motorola Edge+ (2023)?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-infinix-hot-30i-by-fonelab-android-recover-music/"><u>How to restore wiped music on Infinix Hot 30i</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-music-from-xiaomi-redmi-note-13-5g-by-fonelab-android-recover-music/"><u>How to retrieve erased music from Xiaomi Redmi Note 13 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-videos-from-samsung-galaxy-a25-5g-by-fonelab-android-recover-video/"><u>How to retrieve erased videos from Samsung Galaxy A25 5G</u></a></li>
<li><a href="https://unlock-android.techidaily.com/how-to-unlock-honor-play-8t-phone-without-any-data-loss-by-drfone-android/"><u>How to Unlock Honor Play 8T Phone without Any Data Loss</u></a></li>
<li><a href="https://unlock-android.techidaily.com/in-2024-how-to-reset-gmail-password-on-infinix-gt-10-pro-devices-by-drfone-android/"><u>In 2024, How to Reset Gmail Password on Infinix GT 10 Pro Devices</u></a></li>
<li><a href="https://android-transfer.techidaily.com/in-2024-how-to-transfer-data-from-nokia-c210-to-other-android-devices-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>In 2024, How to Transfer Data from Nokia C210 to Other Android Devices? | Dr.fone</u></a></li>
<li><a href="https://win-amazing.techidaily.com/install-lenovo-legion-similar-problem/"><u>Install Lenovo Legion # Similar Problem</u></a></li>
<li><a href="https://extra-resources.techidaily.com/instilling-illusions-blur-wonders-in-adobes-visual-staging-tools/"><u>Instilling Illusions  Blur Wonders in Adobe's Visual Staging Tools</u></a></li>
<li><a href="https://video-content-creator.techidaily.com/new-in-2024-elevate-your-video-game-6-expert-adobe-premiere-techniques-to-know/"><u>New In 2024, Elevate Your Video Game 6 Expert Adobe Premiere Techniques to Know</u></a></li>
<li><a href="https://extra-lessons.techidaily.com/precision-and-panache-advanced-tiktok-editing-skills/"><u>Precision and Panache  Advanced TikTok Editing Skills</u></a></li>
<li><a href="https://facebook.techidaily.com/rumored-ban-for-transparent-commerce-on-facebook/"><u>Rumored Ban for Transparent Commerce on Facebook</u></a></li>
<li><a href="https://win-amazing.techidaily.com/step-by-step-guide-downloading-and-installing-razer-drivers-for-various-windows-versions/"><u>Step-by-Step Guide: Downloading and Installing Razer Drivers for Various Windows Versions</u></a></li>
</ul></div>

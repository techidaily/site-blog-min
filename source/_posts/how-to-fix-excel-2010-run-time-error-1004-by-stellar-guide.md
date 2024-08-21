---
title: How to Fix Excel 2010 Run Time Error 1004
date: 2024-08-20T20:50:59.820Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2010 Run Time Error 1004
keywords: repair excel 2013,repair excel 2023,repair excel 2003,repair .xlb,repair excel 2016,repair excel 2010
---

## How to Fix Excel Run Time Error 1004

**Summary:** Run-time errors are windows-specific issues that occur while the program is running. This blog will teach you how to fix Excel run-time error 1004. In addition, you’ll learn about an Excel repair tool that can help fix the error 1004 if it occurs due to corruption in Excel files.

<!-- affiliate ads begin -->
<h3 id="200610"><a href="https://sentrypc.7eer.net/c/5597632/200610/3022">Parental Control Software</a></h3>
<span class="text-ad-content">
	#1 Rated Parental Control Software.<br/>
	Monitor & Control all PC Activity!<br/>
		<cite style="color:green">sentrypc.com/parental-controls/</cite>
	</span><img height="0" width="0" src="https://sentrypc.7eer.net/i/5597632/200610/3022" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

VBA (Microsoft Visual Basic for Application) is an internal programming language in Microsoft Excel. Sometimes, when users try to run VBA or generate a Macro in Excel, the Run-time error 1004 may occur. This error may occur due to the presence of more legend entries in the chart, file conflict, incorrect Macro name, and corrupt Excel files. In this blog, we have discussed the reasons and shared some solutions to resolve run-time error 1004.

## **Why This Error Occurs?**

The run time error 1004 usually occurs when you run a VBA macro with the Legend Entries method to modify the legend entries in the MS Excel chart. It happens when the chart contains more legend entries than the available space, macro name conflicts, corrupt Excel files, or data-types mismatch in the VBA code.

<!-- affiliate ads begin -->
<a href="https://shop.mondly.com/affiliate.php?ACCOUNT=ATISTUDI&AFFILIATE=108875&PATH=https%3A%2F%2Fwww.mondly.com%3FAFFILIATE%3D108875%26RESOURCE%3D%2BEducational%2B970x90%2B"><img src="https://secure.avangate.com/images/merchant/69c418c33ec2e1a4267fa9bb77fa1428/educational-970x90.gif" border="0"></a>
<!-- affiliate ads end -->
## **Ways to Fix Excel Run-Time Error 1004?**

Try the below workarounds to fix Excel run-time error 1004:

<!-- affiliate ads begin -->
<a href="https://twopages.pxf.io/c/5597632/2016067/18544" target="_top" id="2016067"><img src="//a.impactradius-go.com/display-ad/18544-2016067" border="0" alt="" width="1020" height="380"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2016067/18544" style="position:absolute;visibility:hidden;" border="0" />
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
<a href="https://ukaidot.sjv.io/c/5597632/1793233/19578" target="_top" id="1793233"><img src="//a.impactradius-go.com/display-ad/19578-1793233" border="0" alt="" width="1200" height="1200"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1793233/19578" style="position:absolute;visibility:hidden;" border="0" />
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
<a href="https://vapordna.pxf.io/c/5597632/1496243/17238" target="_top" id="1496243"><img src="//a.impactradius-go.com/display-ad/17238-1496243" border="0" alt="" width="1000" height="1221"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1496243/17238" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **Change Trust Center Settings**

Sometimes, run-time errors might arise because of incorrect security settings. The **Trust Center settings** help you find the **Privacy and security** settings for Microsoft Excel. Follow the below steps to change the **Trust center settings**:

- Open Microsoft Excel.
- Go to **File > Options.**
- The **Excel options** window is displayed.
- Choose **Trust Center**, and click **Trust Center Settings**.
- Tap on the **Macro Settings** tab, and select **Trust access to the VBA project object model.**

<!-- affiliate ads begin -->
<a href="https://store.iobit.com/order/checkout.php?PRODS=4596923&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/184260348236f9554fe9375772ff966e/ascscan_468X60.png" border="0"></a>
<!-- affiliate ads end -->
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
<li><a href="https://visual-screen-recording.techidaily.com/new-2024-approved-broadcasting-e-learning-sessions/"><u>[New] 2024 Approved  Broadcasting E-Learning Sessions</u></a></li>
<li><a href="https://facebook-clips.techidaily.com/new-2024-approved-saving-facebook-gifs-on-computer-phone-and-tablet/"><u>[New] 2024 Approved  Saving Facebook GIFs on Computer, Phone & Tablet</u></a></li>
<li><a href="https://facebook-record-videos.techidaily.com/new-the-finest-10-youtube-personalities-revolutionizing-cosmetics/"><u>[New] The Finest 10 YouTube Personalities Revolutionizing Cosmetics</u></a></li>
<li><a href="https://facebook-video-share.techidaily.com/updated-from-stir-to-screen-masterful-cooking-videos/"><u>[Updated] From Stir to Screen  Masterful Cooking Videos</u></a></li>
<li><a href="https://instagram-clips.techidaily.com/updated-how-to-use-filters-on-instagram-2023-for-2024/"><u>[Updated] How to Use Filters on Instagram 2023 for 2024</u></a></li>
<li><a href="https://desktop-recording.techidaily.com/updated-in-2024-clearview-studio-windows-1011-edition/"><u>[Updated] In 2024, ClearView Studio  Windows 10/11 Edition</u></a></li>
<li><a href="https://facebook-videos.techidaily.com/updated-in-2024-live-stream-longevity-preventing-and-fixing-interruptions-on-social-media/"><u>[Updated] In 2024, Live Stream Longevity  Preventing and Fixing Interruptions on Social Media</u></a></li>
<li><a href="https://tiktok-clips.techidaily.com/updated-revamped-tiktok-visuals-the-ultimate-guide-for-background-swapping/"><u>[Updated] Revamped TikTok Visuals  The Ultimate Guide for Background Swapping</u></a></li>
<li><a href="https://video-capture.techidaily.com/updated-streamlining-your-mac-5-best-screen-shot-approaches/"><u>[Updated] Streamlining Your Mac  5 Best Screen Shot Approaches</u></a></li>
<li><a href="https://extra-support.techidaily.com/2024-approved-iphone-photo-wizards-perfecting-silhouette-art/"><u>2024 Approved  IPhone Photo Wizards  Perfecting Silhouette Art</u></a></li>
<li><a href="https://extra-skills.techidaily.com/2024-approved-navigating-the-landscape-of-color-correction-innovative-uses-of-luts-in-video-editing/"><u>2024 Approved  Navigating the Landscape of Color Correction  Innovative Uses of LUTs in Video Editing</u></a></li>
<li><a href="https://youtube-blog.techidaily.com/approved-simplifying-the-path-to-understanding-youtube-numbers/"><u>2024 Approved  Simplifying the Path to Understanding YouTube Numbers</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-honor-100-pro-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Honor 100 Pro to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-infinix-note-30-vip-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Infinix Note 30 VIP to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-honor-x50-gt-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Honor X50 GT to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-oppo-a78-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Oppo A78 to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-techniques-to-transfer-data-from-vivo-s17t-to-iphone-15141312-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Techniques to Transfer Data from Vivo S17t to iPhone 15/14/13/12 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-ways-to-move-contacts-from-vivo-x-fold-2-to-iphone-131415-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Ways to Move Contacts From Vivo X Fold 2 to iPhone (13/14/15) | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-ways-to-teach-you-to-transfer-files-from-samsung-galaxy-m34-5g-to-other-android-devices-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Ways To Teach You To Transfer Files from Samsung Galaxy M34 5G to Other Android Devices Easily | Dr.fone</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/best-3-nokia-xr21-emulator-for-mac-to-run-your-wanted-android-apps-drfone-by-drfone-android/"><u>Best 3 Nokia XR21 Emulator for Mac to Run Your Wanted Android Apps | Dr.fone</u></a></li>
<li><a href="https://youtube-blog.techidaily.com/the-sub4sub-strategy-improve-content-consumption/"><u>Does the Sub4Sub Strategy Improve Content Consumption?</u></a></li>
<li><a href="https://location-social.techidaily.com/edit-and-send-fake-location-on-telegram-for-your-lava-yuva-2-pro-in-3-ways-drfone-by-drfone-virtual-android/"><u>Edit and Send Fake Location on Telegram For your Lava Yuva 2 Pro in 3 Ways | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-messages-files-from-poco-m6-5g-by-fonelab-android-recover-messages/"><u>How To  Restore Missing Messages Files from Poco M6 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-pictures-files-from-tecno-camon-20-by-fonelab-android-recover-pictures/"><u>How To  Restore Missing Pictures Files from Tecno Camon 20.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-15-plus-without-losing-data-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 15 Plus without Losing Data? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-find-lost-iphone-14-pro-backup-files-on-windows-pc-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to find lost iPhone 14 Pro Backup files on Windows PC? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-corrupt-video-files-of-find-x7-using-video-repair-utility-by-stellar-video-repair-mobile-video-repair/"><u>How to Fix Corrupt video files of Find X7 using Video Repair Utility?</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/how-to-get-and-use-pokemon-go-promo-codes-on-motorola-g54-5g-drfone-by-drfone-virtual-android/"><u>How to Get and Use Pokemon Go Promo Codes On Motorola G54 5G | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-honor-100-get-deleted-photos-back-with-ease-and-safety-by-fonelab-android-recover-photos/"><u>How to Honor 100 Get Deleted photos Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-malfunctioning-drivers-with-windows-device-manager-in-windows-10-by-drivereasy-guide/"><u>How to identify malfunctioning drivers with Windows Device Manager in Windows 10</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-missing-or-malfunctioning-your-drivers-with-windows-device-manager-in-windows-10-and-7-by-drivereasy-guide/"><u>How to identify missing or malfunctioning your drivers with Windows Device Manager in Windows 10 & 7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-make-an-seo-audit-report-by-link-assistant-website-auditor-website-auditor/"><u>How to make an SEO audit report?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-xiaomi-13-ultra-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Xiaomi 13 Ultra</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-pictures-from-vivo-v29e-by-fonelab-android-recover-pictures/"><u>How to recover deleted pictures from Vivo V29e.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-files-after-iphone-14-factory-reset-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Files after iPhone 14 Factory Reset? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-honor-play-7t-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Honor Play 7T</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-google-frp-lock-on-oppo-by-drfone-android-unlock-remove-google-frp/"><u>How to remove Google FRP Lock on Oppo</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-a-damaged-video-file-of-motorola-moto-g-stylus-2023-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair a Damaged video file of Motorola Moto G Stylus (2023) using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-tecno-spark-20-proplus-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Tecno Spark 20 Pro+ using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-ios-system-of-iphone-7-plus-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair iOS System of iPhone 7 Plus? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-samsung-galaxy-s23-fe-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Samsung Galaxy S23 FE</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-iphone-13-mini-without-backup-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Restore iPhone 13 mini without Backup | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-call-logs-from-xr21-by-fonelab-android-recover-call-logs/"><u>How to retrieve erased call logs from XR21?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-music-from-ace-2v-by-fonelab-android-recover-music/"><u>How to retrieve erased music from Ace 2V</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-s18-pro-get-deleted-phone-number-back-with-ease-and-safety-by-fonelab-android-recover-contacts/"><u>How to S18 Pro Get Deleted Phone Number Back with Ease and Safety</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-data-from-iphone-12-pro-max-to-other-iphone-12-pro-devices-drfone-by-drfone-transfer-data-from-ios-transfer-data-from-ios/"><u>How To Transfer Data From iPhone 12 Pro Max To Other iPhone 12 Pro devices? | Dr.fone</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/how-to-unlock-honor-x9b-phone-with-broken-screen-by-drfone-android/"><u>How to Unlock Honor X9b Phone with Broken Screen</u></a></li>
<li><a href="https://bypass-frp.techidaily.com/in-2024-5-quick-methods-to-bypass-vivo-s17-frp-by-drfone-android/"><u>In 2024, 5 Quick Methods to Bypass Vivo S17 FRP</u></a></li>
<li><a href="https://ai-vdieo-software.techidaily.com/in-2024-easy-mp4-editing-on-windows-8-top-rated-video-editor/"><u>In 2024, Easy MP4 Editing on Windows 8 Top-Rated Video Editor</u></a></li>
<li><a href="https://change-location.techidaily.com/in-2024-how-to-use-ispoofer-on-samsung-galaxy-a15-4g-drfone-by-drfone-virtual-android/"><u>In 2024, How to use iSpoofer on Samsung Galaxy A15 4G? | Dr.fone</u></a></li>
<li><a href="https://some-approaches.techidaily.com/in-2024-top-secret-free-apps-best-macos-text-transcribers/"><u>In 2024, Top Secret Free Apps  Best macOS Text Transcribers</u></a></li>
<li><a href="https://win-howtos.techidaily.com/resolving-the-secure-connection-failure-issue-in-mozilla-firefox/"><u>Resolving the 'Secure Connection Failure' Issue in Mozilla Firefox</u></a></li>
<li><a href="https://vp-tips.techidaily.com/ringtone-heaven-game-of-thrones-audio-collection-for-2024/"><u>Ringtone Heaven  Game of Thrones Audio Collection for 2024</u></a></li>
<li><a href="https://youtube-sure.techidaily.com/k-your-income-potential-youtube-studios-global-reach-for-2024/"><u>Unlock Your Income Potential  YouTube Studio's Global Reach for 2024</u></a></li>
<li><a href="https://tech-haven.techidaily.com/why-ios-chatgpt-dominates-web-app-experience/"><u>Why iOS ChatGPT Dominates Web App Experience</u></a></li>
</ul></div>

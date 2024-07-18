---
title: How to Fix Excel 2000 Run Time Error 1004
date: 2024-07-17T09:55:49.872Z
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
<li><a href="https://fox-friendly.techidaily.com/new-2024-approved-effortless-ways-to-grab-high-quality-images-at-no-cost/"><u>[New] 2024 Approved  Effortless Ways To Grab High Quality Images at No Cost</u></a></li>
<li><a href="https://youtube-stream.techidaily.com/new-flutter-freedom-for-videos-the-best-tools-to-flv-to-youtubes/"><u>[New] Flutter Freedom for Videos  The Best Tools to Flv-to-YouTubes</u></a></li>
<li><a href="https://facebook-video-footage.techidaily.com/new-in-2024-beginners-guide-video-setup-must-haves/"><u>[New] In 2024, Beginner's Guide  Video Setup Must-Haves</u></a></li>
<li><a href="https://screen-sharing-recording.techidaily.com/new-lightweight-recording-tool-windows-compatible/"><u>[New] Lightweight Recording Tool, Windows Compatible</u></a></li>
<li><a href="https://video-capture.techidaily.com/new-perfecting-your-presence-share-screen-mastery-in-meet/"><u>[New] Perfecting Your Presence  Share Screen Mastery in Meet</u></a></li>
<li><a href="https://facebook-video-footage.techidaily.com/new-the-creative-summit-top-talent-junction/"><u>[New] The Creative Summit  Top Talent Junction</u></a></li>
<li><a href="https://some-tips.techidaily.com/new-the-ultimate-walkthrough-for-movie-making-on-windows-8/"><u>[New] The Ultimate Walkthrough for Movie Making on Windows 8</u></a></li>
<li><a href="https://facebook-video-footage.techidaily.com/updated-2024-approved-amplify-content-with-free-audio-samples/"><u>[Updated] 2024 Approved  Amplify Content with Free Audio Samples</u></a></li>
<li><a href="https://remote-screen-capture.techidaily.com/updated-revolutionary-6-sustainable-minecraft-homes-for-2024/"><u>[Updated] Revolutionary 6 Sustainable Minecraft Homes for 2024</u></a></li>
<li><a href="https://some-skills.techidaily.com/updated-the-editors-edge-advanced-techniques-in-chromatic-manipulation/"><u>[Updated] The Editor's Edge  Advanced Techniques in Chromatic Manipulation</u></a></li>
<li><a href="https://extra-lessons.techidaily.com/2024-approved-comparing-sns-hdrs-capabilities-with-competitor-software/"><u>2024 Approved  Comparing SNS HDR's Capabilities with Competitor Software</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/a-complete-guide-to-oem-unlocking-on-oneplus-ace-2v-by-drfone-android/"><u>A Complete Guide To OEM Unlocking on OnePlus Ace 2V</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/additional-tips-about-sinnoh-stone-for-nokia-150-2023-drfone-by-drfone-virtual-android/"><u>Additional Tips About Sinnoh Stone For Nokia 150 (2023) | Dr.fone</u></a></li>
<li><a href="https://extra-resources.techidaily.com/detailed-look-at-the-latest-lg-360-camera-tech/"><u>Detailed Look at the Latest LG 360 Camera Tech</u></a></li>
<li><a href="https://fake-location.techidaily.com/fixing-foneazy-mockgo-not-working-on-vivo-y27-5g-drfone-by-drfone-virtual-android/"><u>Fixing Foneazy MockGo Not Working On Vivo Y27 5G | Dr.fone</u></a></li>
<li><a href="https://facebook-clips.techidaily.com/freedom-on-facebook-10-tricks-to-sidestep-blockage-for-2024/"><u>Freedom on Facebook  10 Tricks to Sidestep Blockage for 2024</u></a></li>
<li><a href="https://fix-guide.techidaily.com/how-to-quickly-fix-bluetooth-not-working-on-itel-p55-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>How to Quickly Fix Bluetooth Not Working on Itel P55 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-nubia-red-magic-8s-pro-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Nubia Red Magic 8S Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-nubia-red-magic-8s-proplus-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Nubia Red Magic 8S Pro+</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-nubia-z50-ultra-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Nubia Z50 Ultra</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-nubia-z50s-pro-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Nubia Z50S Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-oneplus-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on OnePlus</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-oneplus-12-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on OnePlus 12</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-oneplus-nord-3-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on OnePlus Nord 3 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-oneplus-nord-ce-3-lite-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on OnePlus Nord CE 3 Lite 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-oneplus-nord-n30-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on OnePlus Nord N30 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-oppo-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Oppo</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-oppo-a1x-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Oppo A1x 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-oppo-a2-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Oppo A2</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-oppo-a58-4g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Oppo A58 4G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-oppo-a59-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Oppo A59 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-oppo-a79-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Oppo A79 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-oppo-find-n3-flip-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Oppo Find N3 Flip</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-oppo-find-x6-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Oppo Find X6</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-oppo-find-x6-pro-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Oppo Find X6 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-oppo-find-x7-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Oppo Find X7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-oppo-find-x7-ultra-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Oppo Find X7 Ultra</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-oppo-reno-10-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Oppo Reno 10 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-oppo-reno-11-pro-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Oppo Reno 11 Pro 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-oppo-reno-9a-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Oppo Reno 9A</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-pixel-7a-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Pixel 7a</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-pixel-8-pro-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Pixel 8 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-pixel-fold-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Pixel Fold</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-play-8t-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Play 8T</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-poco-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Poco</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-poco-f5-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Poco F5 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-poco-f5-pro-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Poco F5 Pro 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-poco-m6-pro-4g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Poco M6 Pro 4G</u></a></li>
<li><a href="https://ios-unlock.techidaily.com/in-2024-how-to-fix-auto-lock-greyed-out-on-iphone-14-pro-by-drfone-ios/"><u>In 2024, How To Fix Auto Lock Greyed Out on iPhone 14 Pro</u></a></li>
<li><a href="https://video-ai-editor.techidaily.com/in-2024-how-to-stabilize-videos-in-adobe-premiere-pro/"><u>In 2024, How to Stabilize Videos in Adobe Premiere Pro ?</u></a></li>
<li><a href="https://sim-unlock.techidaily.com/in-2024-how-to-unlock-sim-cards-of-itel-a05s-without-puk-codes-by-drfone-android/"><u>In 2024, How To Unlock SIM Cards Of Itel A05s Without PUK Codes</u></a></li>
<li><a href="https://pokemon-go-android.techidaily.com/in-2024-how-to-use-ispoofer-on-honor-x8b-drfone-by-drfone-virtual-android/"><u>In 2024, How to use iSpoofer on Honor X8b? | Dr.fone</u></a></li>
<li><a href="https://visual-screen-recording.techidaily.com/in-2024-proven-recorders-ioss-leading-screenshot-tools/"><u>In 2024, Proven Recorders  IOS's Leading Screenshot Tools</u></a></li>
<li><a href="https://sound-tweaking.techidaily.com/updated-top-8-online-free-voice-recorder/"><u>Updated Top 8 Online Free Voice Recorder</u></a></li>
</ul></div>

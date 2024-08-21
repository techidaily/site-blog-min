---
title: How to fix runtime error 424 object required error in Excel 2016
date: 2024-08-20T20:48:11.449Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to fix runtime error 424 object required error in Excel 2016
keywords: repair excel 2021,repair .xls,repair excel 2019,repair excel 2003,repair .xlb
---

## How to fix runtime error 424 object required error in Excel

The Runtime error 424: Object required occurs when Excel is not able to recognize an object that you are referring to in a VBA code. The object can be a workbook, worksheet, range, variable, class, macro, etc. Some users have also reported that this error occurred when they tried to copy the values of the cells from one workbook to another.

Let’s understand the error through a small scenario. Suppose, I want to check the last field row in a table in a spreadsheet named “First” using the VBA code. To do this, I have added a command button and double-clicked on it and entered the below code in the backend:

Private Sub CommandButton2\_Click()

Dim LRow As Integer

LRow = Worksheets("First").Cells(Rows.Count, 2).End(xlUp).Row

MsgBox ("Last Row " & LRow)

End Sub

![Code On Command Button](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/code-on-commandbutton.jpg)

In this code, _Worksheets("First")_ is a data object. If I mistakenly delete this data object and insert any random name (for example - kanada), then it will not be recognized by Excel. When I run this code, I will get the “Run-time error 424”.

![Runtime Error with scenario](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/run-time-error-424-with-scenario.jpg)

<!-- affiliate ads begin -->
<a href="https://zonlipartnershipprogram.pxf.io/c/5597632/1596691/17882" target="_top" id="1596691"><img src="//a.impactradius-go.com/display-ad/17882-1596691" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1596691/17882" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
## **Causes of Runtime Error 424 in Excel**

The Runtime error 424: Object required can occur due to the following reasons:

- Incorrect name of the object you are trying to refer to in a code.
- You have provided an invalid qualifier to an object.
- You have not used the Set statement while assigning an object reference.
- The object is corrupted.
- Missing objects in a workbook.
- Objects you are trying to call in a code are mistakenly deleted or unavailable.
- You have used an incorrect syntax for object declaration.
- You are trying to perform an invalid action on an object in a code.
- Workbook is corrupted.

<!-- affiliate ads begin -->
<a href="https://arkmc.pxf.io/c/5597632/427477/5172" target="_top" id="427477"><img src="//a.impactradius-go.com/display-ad/5172-427477" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://arkmc.pxf.io/i/5597632/427477/5172" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
##  **Solutions to Fix Runtime Error 424: Object Required in Excel**

The VBA error ‘object required’ may occur due to different reasons. Based on the reason, you can follow the solutions mentioned below to fix the error.

### **1\. Check the Name of the Object**

The Runtime error 424 can occur when you run the VBA code using an incorrect name of the object. For example, the object name is ‘MyObject’ but you’re using “Backcolor”.

<!-- affiliate ads begin -->
<a href="https://printrendy.pxf.io/c/5597632/1453719/17020" target="_top" id="1453719"><img src="//a.impactradius-go.com/display-ad/17020-1453719" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1453719/17020" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Error When Incorrect Name Of The Object](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/error-424-when-incorrect-name-of-the-object.jpg)

 When you click the **Debug** button, the line with the error will highlight.

<!-- affiliate ads begin -->
<a href="https://store.iobit.com/order/checkout.php?PRODS=1468905&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/184260348236f9554fe9375772ff966e/ascscan_728x90.png" border="0"></a>
<!-- affiliate ads end -->
![Highlighting Line With Error](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/highlighting-line-with-error.jpg)

To fix the issue, you need to provide the correct name of the object.

<!-- affiliate ads begin -->
<a href="https://shop.mondly.com/affiliate.php?ACCOUNT=ATISTUDI&AFFILIATE=108875&PATH=https%3A%2F%2Fwww.mondly.com%3FAFFILIATE%3D108875%26RESOURCE%3D%2BEducational%2B970x90%2B"><img src="https://secure.avangate.com/images/merchant/69c418c33ec2e1a4267fa9bb77fa1428/educational-970x90.gif" border="0"></a>
<!-- affiliate ads end -->
### **2\. Check if the Object is Missing**

 The Runtime error 424 can occur if the object you are referring to as a method is not available or you are using the wrong object in a code. In the below example, you can see that the error occurs when an object named “Employee” is not available in the Project list.

<!-- affiliate ads begin -->
<a href="https://store.revouninstaller.com/order/checkout.php?PRODS=27889512&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/4282ec8de8c9be897e7aff4aa231b1a4/728__90.jpg" border="0"></a>
<!-- affiliate ads end -->
![Example Of Code When Object Is Not Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/example-of-code-when-an-object-is-not-available.jpg)

 You can check and mention the object which is available. For instance, Sheet2 in the below code.  

![Check When The Object Is Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/check-when-the-object-is-available.png)

### **3\. Check All References are Declared in the Code**

You can get the Runtime error 424 if all the references are not declared. So, make sure you have declared all the references in the code. To verify this, you can use the debug mode by pressing **F5** or clicking on the **Debug** option.

![Debug Command In Excel](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/debug-command-in-excel.jpg)

### **4\. Check the Macro Security Settings**

Sometimes, the error can occur if macros are disabled in the Macro Security settings. You can check and change the settings by following these steps:

- On the **Developer** tab, in the **Code** section, click **Macro Security**.
- In the **Trust Center** window, select **Enable all macros.**

![Macro Security Wizard](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/macro-security-wizard.jpg)

- Click **OK**.

![Enable All Macro In Trust Center](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/enable-all-macros-in-trust-center.jpg)

<!-- affiliate ads begin -->
<a href="https://ursime.pxf.io/c/5597632/2092236/16384" target="_top" id="2092236"><img src="//a.impactradius-go.com/display-ad/16384-2092236" border="0" alt="" width="1920" height="329"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2092236/16384" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### 5\. Repair your Workbook

Sometimes, the ‘Object required’ error can occur if your Excel file is damaged or corrupted. In such a case, you can try repairing the file using Microsoft’s in-built utility - Open and Repair. To use this utility, follow these steps:

- In Excel, go to **File > Open > Browse**.
- In the Open dialog box, click on the corrupted Excel file.
- Click the arrow next to the Open button and select **Open and Repair** from the dropdown.
- Select **Repair** to recover as much data from the file as possible.

If the Open and Repair utility fails or stops working, then you can try a professional Excel repair tool, such as [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). It is an advanced tool that can repair severely corrupted Excel files **(.xls, .xlsx, .xltm, .xltx, and .xlsm)**. It helps recover all the file components, including images, charts, tables, pivot tables, cell comments, chart sheets, formulas, etc., without impacting the original structure.

## **Conclusion**

The Runtime error 424 usually occurs when there is an issue with the objects in your VBA code. In this article, we have covered some effective methods to resolve the “object required” error in Excel. If the error occurs due to corruption in Excel file, then you can repair the corrupt file using Stellar Repair for Excel. It is a reliable tool that can repair severely corrupted Excel file without changing its actual formatting. You can download the free trial version of the software to evaluate its functionality.




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
<li><a href="https://fox-links.techidaily.com/new-master-your-rhythm-top-free-beat-detectors-ready-to-test-for-2024/"><u>[New] Master Your Rhythm – Top Free Beat Detectors Ready to Test for 2024</u></a></li>
<li><a href="https://digital-screen-recording.techidaily.com/new-screenstreamer-breaking-down-the-newest-recorder/"><u>[New] ScreenStreamer  Breaking Down the Newest Recorder</u></a></li>
<li><a href="https://some-guidance.techidaily.com/new-top-9-live-stream-enhancers-ultimate-filter-guide/"><u>[New] Top 9 Live-Stream Enhancers  Ultimate Filter Guide</u></a></li>
<li><a href="https://vp-tips.techidaily.com/updated-2024-approved-navigating-through-windows-11s-new-frontier/"><u>[Updated] 2024 Approved  Navigating Through Windows 11'S New Frontier</u></a></li>
<li><a href="https://facebook-video-share.techidaily.com/updated-finding-your-niche-making-money-with-video-content/"><u>[Updated] Finding Your Niche  Making Money with Video Content</u></a></li>
<li><a href="https://facebook-video-footage.techidaily.com/updated-youtube-to-gif-magic-your-ultimate-guide-to-animated-images/"><u>[Updated] Youtube to Gif Magic  Your Ultimate Guide to Animated Images</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-infinix-hot-30-5g-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Infinix Hot 30 5G to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-samsung-galaxy-z-flip-5-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Samsung Galaxy Z Flip 5 to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://facebook-video-content.techidaily.com/2024-approved-seamless-fb-to-mp4-conversion-service/"><u>2024 Approved  Seamless FB to MP4 Conversion Service</u></a></li>
<li><a href="https://extra-hints.techidaily.com/a-guide-for-aspiring-graders-essential-tips-from-these-top-7/"><u>A Guide for Aspiring Graders  Essential Tips From These Top 7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-call-logs-from-itel-p40plus-by-fonelab-android-recover-call-logs/"><u>How To  Restore Missing Call Logs from Itel P40+</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-pictures-files-from-poco-f5-5g-by-fonelab-android-recover-pictures/"><u>How To  Restore Missing Pictures Files from Poco F5 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-corrupt-video-files-of-lava-agni-2-5g-using-video-repair-utility-by-stellar-video-repair-mobile-video-repair/"><u>How to Fix Corrupt video files of Lava Agni 2 5G using Video Repair Utility?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-back-lost-contacts-from-oppo-a78-by-fonelab-android-recover-contacts/"><u>How to get back lost contacts from Oppo A78.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-out-of-recovery-or-dfu-mode-on-iphone-xs-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Get Out of Recovery or DFU Mode on iPhone XS? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-play-hevc-h-265-video-on-galaxy-z-fold-5-by-aiseesoft-video-converter-play-hevc-video-on-android/"><u>How to play HEVC H.265 video on Galaxy Z Fold 5?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-data-from-lost-or-stolen-iphone-14-pro-in-easy-steps-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How To Recover Data From Lost or Stolen iPhone 14 Pro In Easy Steps | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-after-format-on-huawei-nova-y91-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery after format on Huawei Nova Y91</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-app-on-poco-m6-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Recover Deleted Photos from Android Gallery App on Poco M6 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-realme-gt-5-240w-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Realme GT 5 (240W)</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-11-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on 11 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-x50-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on X50</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-screenshots-on-iphone-15-pro-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Screenshots on iPhone 15 Pro? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-files-after-iphone-13-pro-factory-reset-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Files after iPhone 13 Pro Factory Reset? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-motorola-edge-2023-by-fonelab-android-recover-music/"><u>How to recover old music from your Motorola Edge 2023</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-google-frp-lock-on-motorola-g54-5g-by-drfone-android-unlock-remove-google-frp/"><u>How to remove Google FRP Lock on Motorola G54 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-realme-pin-by-drfone-android-unlock-android-unlock/"><u>How to remove Realme PIN</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-a-damaged-video-file-of-honor-magic-6-lite-using-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair a Damaged video file of Honor Magic 6 Lite using Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-lava-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Lava on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-samsung-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Samsung on Windows??</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-infinix-gt-10-pro-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Infinix GT 10 Pro? </u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupted-pdf-v16-file-when-manual-method-fails-stellar-by-stellar-guide/"><u>How to repair corrupted PDF v1.6 file when manual method fails | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-system-issues-of-iphone-8-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair System Issues of iPhone 8? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-call-logs-from-asus-rog-phone-8-pro-by-fonelab-android-recover-call-logs/"><u>How to rescue lost call logs from Asus ROG Phone 8 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-messages-from-honor-80-pro-straight-screen-edition-by-fonelab-android-recover-messages/"><u>How to Rescue Lost Messages from Honor 80 Pro Straight Screen Edition</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-messages-from-infinix-note-30-pro-by-fonelab-android-recover-messages/"><u>How to Rescue Lost Messages from Infinix Note 30 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-photos-from-asus-rog-phone-8-pro-by-fonelab-android-recover-photos/"><u>How to Rescue Lost Photos from Asus ROG Phone 8 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-15-pro-without-losing-data-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Reset iPhone 15 Pro without Losing Data? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-your-iphone-13-pro-max-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Reset Your iPhone 13 Pro Max? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-call-logs-from-pop-7-pro-by-fonelab-android-recover-call-logs/"><u>How to retrieve erased call logs from Pop 7 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-videos-from-infinix-note-30i-by-fonelab-android-recover-video/"><u>How to retrieve erased videos from Infinix Note 30i</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-videos-from-motorola-moto-g14-by-fonelab-android-recover-video/"><u>How to retrieve erased videos from Motorola Moto G14</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-lost-files-from-htc-u23-pro-by-fonelab-android-recover-data/"><u>How to retrieve lost files from HTC U23 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-solve-mkv-lagging-problem-in-motorola-moto-g34-5g-by-aiseesoft-video-converter-play-mkv-on-android/"><u>How to solve MKV lagging problem in Motorola Moto G34 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-data-from-iphone-12-pro-max-to-other-iphone-11-pro-devices-drfone-by-drfone-transfer-data-from-ios-transfer-data-from-ios/"><u>How To Transfer Data From iPhone 12 Pro Max To Other iPhone 11 Pro devices? | Dr.fone</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/how-to-unlock-realme-c55-pin-codepattern-lockpassword-by-drfone-android/"><u>How to Unlock Realme C55 PIN Code/Pattern Lock/Password</u></a></li>
<li><a href="https://facebook-videos.techidaily.com/in-2024-googleplus-photo-frame-settings/"><u>In 2024, Google+ Photo Frame Settings</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/in-2024-is-pgsharp-legal-when-you-are-playing-pokemon-on-xiaomi-14-pro-drfone-by-drfone-virtual-android/"><u>In 2024, Is pgsharp legal when you are playing pokemon On Xiaomi 14 Pro? | Dr.fone</u></a></li>
<li><a href="https://extra-approaches.techidaily.com/in-2024-radiance-in-android-video-editing/"><u>In 2024, Radiance in Android Video Editing</u></a></li>
<li><a href="https://hardware-help.techidaily.com/latest-windows-compatible-amd-radeon-graphics-driver-software-available-to-download/"><u>Latest Windows-Compatible AMD Radeon Graphics Driver Software Available to Download</u></a></li>
<li><a href="https://change-location.techidaily.com/ultimate-guide-to-catch-the-regional-located-pokemon-for-samsung-galaxy-a34-5g-drfone-by-drfone-virtual-android/"><u>Ultimate Guide to Catch the Regional-Located Pokemon For Samsung Galaxy A34 5G | Dr.fone</u></a></li>
</ul></div>

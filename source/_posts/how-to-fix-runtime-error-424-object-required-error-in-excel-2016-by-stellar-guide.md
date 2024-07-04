---
title: How to fix runtime error 424 object required error in Excel 2016
date: 2024-05-19T18:32:11.904Z
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

##  **Solutions to Fix Runtime Error 424: Object Required in Excel**

The VBA error ‘object required’ may occur due to different reasons. Based on the reason, you can follow the solutions mentioned below to fix the error.

### **1\. Check the Name of the Object**

The Runtime error 424 can occur when you run the VBA code using an incorrect name of the object. For example, the object name is ‘MyObject’ but you’re using “Backcolor”.

![Error When Incorrect Name Of The Object](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/error-424-when-incorrect-name-of-the-object.jpg)

 When you click the **Debug** button, the line with the error will highlight.

![Highlighting Line With Error](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/highlighting-line-with-error.jpg)

To fix the issue, you need to provide the correct name of the object.

### **2\. Check if the Object is Missing**

 The Runtime error 424 can occur if the object you are referring to as a method is not available or you are using the wrong object in a code. In the below example, you can see that the error occurs when an object named “Employee” is not available in the Project list.

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
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-videos-from-galaxy-f54-5g-by-fonelab-android-recover-video/"><u>How to Rescue Lost Videos from Galaxy F54 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-lava-pictures-an-easy-method-explained-by-fonelab-android-recover-pictures/"><u>How to Restore Deleted Lava Pictures  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-12-mini-without-losing-any-content-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 12 mini without Losing Any Content? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-system-issues-of-iphone-6s-plus-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair System Issues of iPhone 6s Plus? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-pictures-files-from-xiaomi-13t-by-fonelab-android-recover-pictures/"><u>How To  Restore Missing Pictures Files from Xiaomi 13T.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-call-history-on-nokia-by-fonelab-android-recover-call-logs/"><u>How to restore wiped call history on Nokia ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-a-damaged-video-file-of-poco-f5-5g-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair a Damaged video file of Poco F5 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-music-from-x5-pro-by-fonelab-android-recover-music/"><u>How to retrieve erased music from X5 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-free-up-iphone-x-space-drfone-by-drfone-ios-full-data-eraser-ios-full-data-eraser/"><u>How To Free Up iPhone X Space | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-itel-p55t-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Itel P55T to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-ios-system-issues-of-iphone-14-plus-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair iOS System Issues of iPhone 14 Plus? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-videos-on-poco-f5-pro-5g-by-fonelab-android-recover-video/"><u>How to restore wiped videos on Poco F5 Pro 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-iphone-13-pro-max-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data from iPhone 13 Pro Max? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-erase-private-data-from-iphone-13-drfone-by-drfone-ios-full-data-eraser-ios-full-data-eraser/"><u>How To Erase Private Data From iPhone 13 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-call-logs-from-infinix-note-30-vip-by-fonelab-android-recover-call-logs/"><u>How To  Restore Missing Call Logs from Infinix Note 30 VIP</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-photos-from-honor-magic-vs-2-by-fonelab-android-recover-photos/"><u>How to Rescue Lost Photos from Honor Magic Vs 2?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-messages-from-htc-by-fonelab-android-recover-messages/"><u>How to Rescue Lost Messages from HTC</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-google-frp-lock-on-samsung-galaxy-a05-by-drfone-android-unlock-remove-google-frp/"><u>How to remove Google FRP Lock on Samsung Galaxy A05</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-oppo-reno-9a-get-deleted-pictures-back-with-ease-and-safety-by-fonelab-android-recover-pictures/"><u>How to Oppo Reno 9A Get Deleted Pictures Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-xs-max-without-losing-data-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Reset iPhone XS Max without Losing Data? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-after-format-on-realme-narzo-n55-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery after format on Realme Narzo N55</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-google-frp-lock-on-narzo-60x-5g-by-drfone-android-unlock-remove-google-frp/"><u>How to remove Google FRP Lock on Narzo 60x 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-messages-files-from-realme-narzo-n53-by-fonelab-android-recover-messages/"><u>How To  Restore Missing Messages Files from Realme Narzo N53</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-xs-max-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Reset iPhone XS Max? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-iphone-13-pro-max-camera-roll-photos-and-photo-stream-pictures-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted iPhone 13 Pro Max Camera Roll Photos and Photo Stream Pictures? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-install-and-update-hardware-device-drivers-manually-in-windows-11-by-drivereasy-guide/"><u>How to install and update hardware device drivers manually in Windows 11</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-app-on-p60-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Recover Deleted Photos from Android Gallery App on P60</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-music-from-smart-8-hd-by-fonelab-android-recover-music/"><u>How to retrieve erased music from Smart 8 HD</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-iphone-11-pro-stuck-at-attempting-data-recovery-loop-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Fix iPhone 11 Pro Stuck at attempting data recovery Loop | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-vivo-y77t-by-fonelab-android-recover-music/"><u>How to restore wiped music on Vivo Y77t</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-lost-files-from-nokia-150-2023-by-fonelab-android-recover-data/"><u>How to retrieve lost files from Nokia 150 (2023)?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-vivo-x-flip-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Vivo X Flip? </u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-realme-12-proplus-5g-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Realme 12 Pro+ 5G to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-call-history-on-motorola-edge-2023-by-fonelab-android-recover-call-logs/"><u>How to restore wiped call history on Motorola Edge 2023?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-call-logs-from-oppo-a1x-5g-by-fonelab-android-recover-call-logs/"><u>How to rescue lost call logs from Oppo A1x 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-iphone-13-pro-camera-roll-photos-and-photo-stream-pictures-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted iPhone 13 Pro Camera Roll Photos and Photo Stream Pictures? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-a-damaged-video-file-of-google-pixel-7a-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair a Damaged video file of Google Pixel 7a using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-a-pdf-v11-document-with-electronic-signature-tool-by-ldigisigner-sign-a-pdf-sign-a-pdf/"><u>How to sign a PDF v1.1 document with electronic signature tool</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-contacts-from-spark-20-by-fonelab-android-recover-contacts/"><u>How to recover deleted contacts from Spark 20.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-t2-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on T2 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-lava-agni-2-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Lava Agni 2 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-poco-f5-pro-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Poco F5 Pro 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-pictures-from-realme-by-fonelab-android-recover-pictures/"><u>How to Rescue Lost Pictures from Realme ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-put-iphone-8-or-ipad-on-recovery-mode-step-by-step-tutorial-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Put iPhone 8 or iPad on Recovery mode? (Step by Step Tutorial) | Stellar</u></a></li>
<li><a href="https://howto.techidaily.com/what-to-do-when-realme-note-50-has-black-screen-of-death-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>What To Do When Realme Note 50 Has Black Screen of Death? | Dr.fone</u></a></li>
<li><a href="https://iphone-unlock.techidaily.com/in-2024-forgot-apple-iphone-15-plus-backup-password-heres-what-to-do-drfone-by-drfone-ios/"><u>In 2024, Forgot Apple iPhone 15 Plus Backup Password? Heres What to Do | Dr.fone</u></a></li>
<li><a href="https://location-social.techidaily.com/in-2024-3-things-you-must-know-about-fake-snapchat-location-on-motorola-moto-g34-5g-drfone-by-drfone-virtual-android/"><u>In 2024, 3 Things You Must Know about Fake Snapchat Location On Motorola Moto G34 5G | Dr.fone</u></a></li>
<li><a href="https://location-fake.techidaily.com/8-solutions-to-fix-find-my-friends-location-not-available-on-motorola-moto-g23-drfone-by-drfone-virtual-android/"><u>8 Solutions to Fix Find My Friends Location Not Available On Motorola Moto G23 | Dr.fone</u></a></li>
<li><a href="https://ai-editing-video.techidaily.com/updated-a-complete-guide-to-use-distortion-effect-in-premiere-pro-for-2024/"><u>Updated A Complete Guide to Use Distortion Effect in Premiere Pro for 2024</u></a></li>
<li><a href="https://review-topics.techidaily.com/in-2024-complete-tutorial-to-use-vpna-to-fake-gps-location-on-realme-gt-5-drfone-by-drfone-virtual-android/"><u>In 2024, Complete Tutorial to Use VPNa to Fake GPS Location On Realme GT 5 | Dr.fone</u></a></li>
<li><a href="https://sim-unlock.techidaily.com/in-2024-ways-to-find-unlocking-codes-for-oppo-reno-10-5g-phones-by-drfone-android/"><u>In 2024, Ways To Find Unlocking Codes For Oppo Reno 10 5G Phones</u></a></li>
<li><a href="https://ai-vdieo-software.techidaily.com/updated-how-to-reduce-the-size-of-video-file-online-in-this-article-well-share-top-10-free-online-video-compressors-that-you-can-use-anywhere-at-any-time-an/"><u>Updated How to Reduce the Size of Video File Online? In This Article, Well Share Top 10 Free Online Video Compressors that You Can Use Anywhere at Any Time and on Any Device</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/how-to-use-allshare-cast-to-turn-on-screen-mirroring-on-honor-magic-6-pro-drfone-by-drfone-android/"><u>How To Use Allshare Cast To Turn On Screen Mirroring On Honor Magic 6 Pro | Dr.fone</u></a></li>
<li><a href="https://ai-voice-clone.techidaily.com/new-best-hatsune-miku-ai-voice-generators-for-all-times/"><u>New Best Hatsune Miku AI Voice Generators for All Times</u></a></li>
<li><a href="https://android-transfer.techidaily.com/in-2024-how-to-transfer-apps-from-honor-x9a-to-another-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>In 2024, How to Transfer Apps from Honor X9a to Another | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/easy-to-configure-mt4-and-mt5-accounts-for-running-a-local-trade-copier-tm-together-with-any-other-forex-ea-by-mt4copier-guide/"><u>Easy to Configure MT4 and MT5 Accounts for Running a Local Trade Copier™ Together With Any Other Forex EA</u></a></li>
<li><a href="https://bypass-frp.techidaily.com/in-2024-step-by-step-tutorial-how-to-bypass-tecno-frp-by-drfone-android/"><u>In 2024, Step-by-Step Tutorial How To Bypass Tecno FRP</u></a></li>
<li><a href="https://unlock-android.techidaily.com/a-perfect-guide-to-remove-or-disable-google-smart-lock-on-tecno-pova-5-by-drfone-android/"><u>A Perfect Guide To Remove or Disable Google Smart Lock On Tecno Pova 5</u></a></li>
<li><a href="https://android-location-track.techidaily.com/top-5-car-locator-apps-for-nokia-g310-drfone-by-drfone-virtual-android/"><u>Top 5 Car Locator Apps for Nokia G310 | Dr.fone</u></a></li>
<li><a href="https://ai-video-editing.techidaily.com/2024-approved-top-10-free-spanish-text-to-speech-converter-software-windows-mac/"><u>2024 Approved Top 10 Free Spanish Text To Speech Converter Software Windows, Mac</u></a></li>
<li><a href="https://iphone-unlock.techidaily.com/how-to-unlock-apple-iphone-6-plus-without-passcode-drfone-by-drfone-ios/"><u>How to Unlock Apple iPhone 6 Plus Without Passcode? | Dr.fone</u></a></li>
</ul></div>



---
title: How to fix runtime error 424 object required error in Excel 2000
date: 2024-05-19T18:32:11.878Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to fix runtime error 424 object required error in Excel 2000
keywords: repair .xlsx,repair excel 2023,repair excel 2019,repair .xlsm,repair excel 2010,repair excel
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
<li><a href="https://blog-min.techidaily.com/how-to-install-the-latest-iosipados-beta-version-on-iphone-15-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Install the Latest iOS/iPadOS Beta Version on iPhone 15? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-call-logs-from-play-40c-by-fonelab-android-recover-call-logs/"><u>How to retrieve erased call logs from Play 40C?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-screenshots-on-iphone-14-pro-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Screenshots on iPhone 14 Pro? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-motorola-moto-e13-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Motorola Moto E13 to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-deleted-photos-on-oppo-reno-10-pro-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Retrieve  deleted photos on Oppo Reno 10 Pro 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-iphone-15-pro-max-face-id-by-drfone-ios-unlock-ios-unlock/"><u>How to Remove iPhone 15 Pro Max Face ID?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-videos-not-playing-on-poco-by-stellar-video-repair-mobile-video-repair/"><u>How to Fix Videos Not Playing on Poco ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-photos-from-nokia-g42-5g-by-fonelab-android-recover-photos/"><u>How to Rescue Lost Photos from Nokia G42 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-htc-u23-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from HTC U23</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-oneplus-nord-n30-se-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from OnePlus Nord N30 SE to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-google-frp-lock-on-nokia-g310-by-drfone-android-unlock-remove-google-frp/"><u>How to remove Google FRP Lock on Nokia G310</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-techniques-to-transfer-data-from-honor-x7b-to-iphone-15141312-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Techniques to Transfer Data from Honor X7b to iPhone 15/14/13/12 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-c110-by-fonelab-android-recover-music/"><u>How to recover old music from your C110</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-deleted-photos-on-a79-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Retrieve deleted photos on A79 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-call-logs-from-infinix-note-30-vip-racing-edition-by-fonelab-android-recover-call-logs/"><u>How to rescue lost call logs from Infinix Note 30 VIP Racing Edition</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-messages-from-samsung-galaxy-a24-by-fonelab-android-recover-messages/"><u>How to Rescue Lost Messages from Samsung Galaxy A24</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-data-from-iphone-12-pro-max-to-others-android-devices-drfone-by-drfone-transfer-data-from-ios-transfer-data-from-ios/"><u>How To Transfer Data From iPhone 12 Pro Max To Others Android Devices? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-videos-and-music-files-from-iphone-xs-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Photos, Videos & Music Files from iPhone XS | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-call-logs-from-honor-v-purse-by-fonelab-android-recover-call-logs/"><u>How to retrieve erased call logs from Honor V Purse?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-messages-on-samsung-galaxy-z-flip-5-by-fonelab-android-recover-messages/"><u>How to restore wiped messages on Samsung Galaxy Z Flip 5</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-deleted-photos-on-vivo-s17e-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Retrieve  deleted photos on Vivo S17e</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-app-on-galaxy-s24-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Recover Deleted Photos from Android Gallery App on Galaxy S24</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-samsung-galaxy-a15-4g-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Samsung Galaxy A15 4G to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-messages-on-htc-u23-by-fonelab-android-recover-messages/"><u>How to restore wiped messages on HTC U23</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-videos-from-tecno-pop-8-by-fonelab-android-recover-video/"><u>How to retrieve erased videos from Tecno Pop 8</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-se-without-apple-password-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Reset iPhone SE Without Apple Password? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-samsung-galaxy-s24-ultra-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Samsung Galaxy S24 Ultra to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-lava-storm-5g-get-deleted-photos-back-with-ease-and-safety-by-fonelab-android-recover-photos/"><u>How to Lava Storm 5G Get Deleted photos Back with Ease and Safety?</u></a></li>
<li><a href="https://android-frp.techidaily.com/how-to-bypass-google-frp-lock-on-nokia-c12-devices-by-drfone-android/"><u>How to Bypass Google FRP Lock on Nokia C12 Devices</u></a></li>
<li><a href="https://ai-video-apps.techidaily.com/updated-in-2024-free-8-best-tools-to-make-animated-photo-easily/"><u>Updated In 2024, FREE 8 Best Tools to Make Animated Photo Easily</u></a></li>
<li><a href="https://location-fake.techidaily.com/3-ways-to-fake-gps-without-root-on-realme-c67-5g-drfone-by-drfone-virtual-android/"><u>3 Ways to Fake GPS Without Root On Realme C67 5G | Dr.fone</u></a></li>
<li><a href="https://android-transfer.techidaily.com/how-to-transfer-text-messages-from-infinix-smart-8-hd-to-new-phone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Text Messages from Infinix Smart 8 HD to New Phone | Dr.fone</u></a></li>
<li><a href="https://android-transfer.techidaily.com/in-2024-how-to-transfer-data-after-switching-from-lava-yuva-2-pro-to-latest-samsung-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>In 2024, How to Transfer Data After Switching From Lava Yuva 2 Pro to Latest Samsung | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/best-android-data-recovery-undelete-lost-music-from-honor-by-fonelab-android-recover-music/"><u>Best Android Data Recovery - Undelete Lost Music from Honor</u></a></li>
<li><a href="https://apple-account.techidaily.com/in-2024-how-to-remove-phone-number-from-your-apple-id-from-your-iphone-14-plus-by-drfone-ios/"><u>In 2024, How To Remove Phone Number From Your Apple ID from Your iPhone 14 Plus?</u></a></li>
<li><a href="https://change-location.techidaily.com/in-2024-9-mind-blowing-tricks-to-hatch-eggs-in-pokemon-go-without-walking-on-vivo-t2-5g-drfone-by-drfone-virtual-android/"><u>In 2024, 9 Mind-Blowing Tricks to Hatch Eggs in Pokemon Go Without Walking On Vivo T2 5G | Dr.fone</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/guide-to-mirror-your-vivo-y78plus-t1-edition-to-other-android-devices-drfone-by-drfone-android/"><u>Guide to Mirror Your Vivo Y78+ (T1) Edition to Other Android devices | Dr.fone</u></a></li>
<li><a href="https://unlock-android.techidaily.com/in-2024-how-to-reset-your-zte-blade-a73-5g-lock-screen-password-by-drfone-android/"><u>In 2024, How to Reset your ZTE Blade A73 5G Lock Screen Password</u></a></li>
<li><a href="https://ios-pokemon-go.techidaily.com/in-2024-will-pokemon-go-ban-the-account-if-you-use-pgsharp-on-apple-iphone-se-drfone-by-drfone-virtual-ios/"><u>In 2024, Will Pokémon Go Ban the Account if You Use PGSharp On Apple iPhone SE | Dr.fone</u></a></li>
<li><a href="https://ai-video-editing.techidaily.com/new-2024-approved-edit-flv-video-files-with-flv-editor-windows-macandroid-iphone/"><u>New 2024 Approved Edit FLV Video Files with FLV Editor Windows, Mac，Android, iPhone</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/in-2024-full-guide-on-mirroring-your-zte-nubia-z60-ultra-to-your-pcmac-drfone-by-drfone-android/"><u>In 2024, Full Guide on Mirroring Your ZTE Nubia Z60 Ultra to Your PC/Mac | Dr.fone</u></a></li>
<li><a href="https://location-fake.techidaily.com/5-hassle-free-solutions-to-fake-location-on-find-my-friends-of-honor-90-gt-drfone-by-drfone-virtual-android/"><u>5 Hassle-Free Solutions to Fake Location on Find My Friends Of Honor 90 GT | Dr.fone</u></a></li>
<li><a href="https://android-frp.techidaily.com/in-2024-how-to-bypass-google-frp-lock-from-samsung-galaxy-s23plus-devices-by-drfone-android/"><u>In 2024, How to Bypass Google FRP Lock from Samsung Galaxy S23+ Devices</u></a></li>
<li><a href="https://ai-editing-video.techidaily.com/in-2024-how-to-create-motion-text-effects-for-your-video/"><u>In 2024, How to Create Motion Text Effects for Your Video?</u></a></li>
<li><a href="https://android-frp.techidaily.com/ultimate-guide-from-motorola-moto-g34-5g-frp-bypass-by-drfone-android/"><u>Ultimate Guide from Motorola Moto G34 5G FRP Bypass</u></a></li>
<li><a href="https://activate-lock.techidaily.com/in-2024-3-effective-ways-to-bypass-activation-lock-from-apple-iphone-6s-plus-by-drfone-ios/"><u>In 2024, 3 Effective Ways to Bypass Activation Lock from Apple iPhone 6s Plus</u></a></li>
</ul></div>



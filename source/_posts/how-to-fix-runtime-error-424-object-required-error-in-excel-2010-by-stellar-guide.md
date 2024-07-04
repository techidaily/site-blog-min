---
title: How to fix runtime error 424 object required error in Excel 2010
date: 2024-05-19T18:32:11.890Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to fix runtime error 424 object required error in Excel 2010
keywords: repair excel 2023,repair .xlsm,repair excel 2013,repair .xlb,repair .xltx,repair excel 2019
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
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-7-plus-to-an-older-ios-system-version-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 7 Plus to an Older iOS System Version? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-electronically-sign-a-excel-2007-using-digisigner-by-ldigisigner-sign-a-excel-sign-a-excel/"><u>How to Electronically Sign a Excel 2007 Using DigiSigner</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-realme-narzo-60x-5g-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Realme Narzo 60x 5G to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-contacts-from-sony-xperia-5-v-by-fonelab-android-recover-contacts/"><u>How to recover deleted contacts from Sony Xperia 5 V.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-honor-magic-6-pro-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Honor Magic 6 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-6s-plus-to-the-previous-ios-system-version-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 6s Plus to the Previous iOS System Version? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-oneplus-nord-n30-5g-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from OnePlus Nord N30 5G to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-nubia-red-magic-8s-proplus-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Nubia Red Magic 8S Pro+</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-infinix-by-fonelab-android-recover-data/"><u>How to recover lost data from Infinix ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-pictures-files-from-lava-blaze-pro-5g-by-fonelab-android-recover-pictures/"><u>How To  Restore Missing Pictures Files from Lava Blaze Pro 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-motorola-moto-g04-with-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Motorola Moto G04 with Video Repair Utility on Windows? </u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-messages-from-vivo-s17-by-fonelab-android-recover-messages/"><u>How to Rescue Lost Messages from Vivo S17</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-some-outdated-your-hardware-drivers-with-windows-device-manager-on-windows-11-and-10-by-drivereasy-guide/"><u>How to identify some outdated your hardware drivers with Windows Device Manager on Windows 11 & 10</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-call-logs-from-infinix-smart-8-hd-by-fonelab-android-recover-call-logs/"><u>How to rescue lost call logs from Infinix Smart 8 HD</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-contacts-files-from-honor-90-by-fonelab-android-recover-contacts/"><u>How To  Restore Missing Contacts Files from Honor 90.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-after-format-on-vivo-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery after format on Vivo</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-oppo-find-x7-ultra-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Oppo Find X7 Ultra to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-motorola-moto-g84-5g-using-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Motorola Moto G84 5G using Video Repair Utility on Windows? </u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-a-damaged-video-file-of-find-x7-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair a Damaged video file of Find X7 using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-videos-from-tecno-spark-go-2024-by-fonelab-android-recover-video/"><u>How to Rescue Lost Videos from Tecno Spark Go (2024)</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-videos-on-lava-blaze-pro-5g-by-fonelab-android-recover-video/"><u>How to restore wiped videos on Lava Blaze Pro 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-lava-blaze-pro-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Lava Blaze Pro 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-install-the-latest-iosipados-beta-version-on-iphone-15-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Install the Latest iOS/iPadOS Beta Version on iPhone 15? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-exit-recovery-mode-on-iphone-14-plus-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Exit Recovery Mode on iPhone 14 Plus? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-tecno-spark-20-pin-by-drfone-android-unlock-android-unlock/"><u>How to remove Tecno Spark 20 PIN</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-app-on-galaxy-a05-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Recover Deleted Photos from Android Gallery App on Galaxy A05</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-contacts-from-vivo-by-fonelab-android-recover-contacts/"><u>How to recover deleted contacts from Vivo .</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-infinix-hot-40-pin-by-drfone-android-unlock-android-unlock/"><u>How to remove Infinix Hot 40 PIN</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-samsung-galaxy-s23-tactical-edition-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Samsung Galaxy S23 Tactical Edition to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-exit-dfu-mode-on-iphone-se-2022-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Exit DFU Mode on iPhone SE (2022)? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-zte-blade-a73-5g-photos-an-easy-method-explained-by-fonelab-android-recover-photos/"><u>How to Restore Deleted ZTE Blade A73 5G Photos  An Easy Method Explained.</u></a></li>
<li><a href="https://activate-lock.techidaily.com/in-2024-icloud-unlocker-download-unlock-icloud-lock-for-your-apple-iphone-xr-by-drfone-ios/"><u>In 2024, iCloud Unlocker Download Unlock iCloud Lock for your Apple iPhone XR</u></a></li>
<li><a href="https://ai-editing-video.techidaily.com/how-long-does-it-take-to-edit-a-video/"><u>How Long Does It Take To Edit A Video</u></a></li>
<li><a href="https://apple-account.techidaily.com/3-ways-of-how-to-get-someones-apple-id-off-apple-iphone-12-pro-max-without-password-by-drfone-ios/"><u>3 Ways of How to Get Someones Apple ID Off Apple iPhone 12 Pro Max without Password</u></a></li>
<li><a href="https://fake-location.techidaily.com/how-to-simulate-gps-movement-in-ar-games-on-vivo-x100-pro-drfone-by-drfone-virtual-android/"><u>How to Simulate GPS Movement in AR games On Vivo X100 Pro? | Dr.fone</u></a></li>
<li><a href="https://ai-video-editing.techidaily.com/in-2024-this-article-talks-about-if-the-windows-movie-maker-available-and-if-yes-then-how-you-can-download-it-and-so-on-it-further-comprises-an-alternative-/"><u>In 2024, This Article Talks About if the Windows Movie Maker Available and if Yes, Then How You Can Download It and so On. It Further Comprises an Alternative to the Same to Create Videos</u></a></li>
<li><a href="https://android-frp.techidaily.com/a-quick-guide-to-oneplus-nord-3-5g-frp-bypass-instantly-by-drfone-android/"><u>A Quick Guide to OnePlus Nord 3 5G FRP Bypass Instantly</u></a></li>
<li><a href="https://bypass-frp.techidaily.com/in-2024-top-5-vivo-t2-5g-bypass-frp-tools-for-pc-that-actually-work-by-drfone-android/"><u>In 2024, Top 5 Vivo T2 5G Bypass FRP Tools for PC That Actually Work</u></a></li>
<li><a href="https://ai-video-apps.techidaily.com/the-ultimate-guide-to-16x9-aspect-ratio-calculations-for-2024/"><u>The Ultimate Guide to 16X9 Aspect Ratio Calculations for 2024</u></a></li>
<li><a href="https://howto.techidaily.com/7-fixes-for-unfortunately-phone-has-stopped-on-nubia-red-magic-8s-proplus-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>7 Fixes for Unfortunately, Phone Has Stopped on Nubia Red Magic 8S Pro+ | Dr.fone</u></a></li>
<li><a href="https://howto.techidaily.com/9-quick-fixes-to-unfortunately-touchwiz-has-stopped-of-realme-narzo-60x-5g-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>9 Quick Fixes to Unfortunately TouchWiz has stopped Of Realme Narzo 60x 5G | Dr.fone</u></a></li>
<li><a href="https://bypass-frp.techidaily.com/in-2024-ultimate-guide-on-xiaomi-redmi-note-12r-frp-bypass-by-drfone-android/"><u>In 2024, Ultimate Guide on Xiaomi Redmi Note 12R FRP Bypass</u></a></li>
</ul></div>



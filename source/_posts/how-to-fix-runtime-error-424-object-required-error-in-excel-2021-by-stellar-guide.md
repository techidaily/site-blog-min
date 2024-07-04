---
title: How to fix runtime error 424 object required error in Excel 2021
date: 2024-05-19T18:32:11.911Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to fix runtime error 424 object required error in Excel 2021
keywords: repair .xls,repair excel 2023,repair excel 2003,repair .xlsx
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
<li><a href="https://blog-min.techidaily.com/how-to-recover-files-after-iphone-13-mini-factory-reset-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Files after iPhone 13 mini Factory Reset? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-error-1015-while-restoring-iphone-13-mini-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to fix error 1015 while restoring iPhone 13 mini | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-contacts-files-from-samsung-galaxy-f15-5g-by-fonelab-android-recover-contacts/"><u>How To  Restore Missing Contacts Files from Samsung Galaxy F15 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-photos-files-from-google-pixel-fold-by-fonelab-android-recover-photos/"><u>How To  Restore Missing Photos Files from Google Pixel Fold.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-videos-from-oneplus-nord-ce-3-5g-by-fonelab-android-recover-video/"><u>How to Rescue Lost Videos from OnePlus Nord CE 3 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-some-outdated-drivers-with-windows-device-manager-in-windows-7-by-drivereasy-guide/"><u>How to identify some outdated drivers with Windows Device Manager in Windows 7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-reno-10-pro-5g-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Reno 10 Pro 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-poco-f5-pro-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Poco F5 Pro 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-after-format-on-samsung-galaxy-f04-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery after format on Samsung Galaxy F04</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-google-frp-lock-on-tecno-spark-10-pro-by-drfone-android-unlock-remove-google-frp/"><u>How to remove Google FRP Lock on Tecno Spark 10 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-tecno-spark-20-pro-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Tecno Spark 20 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-realme-11-proplus-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Realme 11 Pro+ to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-photos-from-itel-a70-by-fonelab-android-recover-photos/"><u>How to Rescue Lost Photos from Itel A70?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-system-of-iphone-13-pro-max-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair System of iPhone 13 Pro Max? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-google-pixel-7a-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Google Pixel 7a to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-nokia-c02-get-deleted-photos-back-with-ease-and-safety-by-fonelab-android-recover-photos/"><u>How to Nokia C02 Get Deleted photos Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-back-lost-photos-from-vivo-g2-by-fonelab-android-recover-photos/"><u>How to get back lost photos from Vivo G2.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-missing-drivers-with-windows-device-manager-in-windows-11107-by-drivereasy-guide/"><u>How to identify missing drivers with Windows Device Manager in Windows 11/10/7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-call-logs-from-realme-gt-5-240w-by-fonelab-android-recover-call-logs/"><u>How to retrieve erased call logs from Realme GT 5 (240W)?</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-honor-magic-vs-2-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Honor Magic Vs 2 to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-videos-from-your-nokia-c300-by-fonelab-android-recover-video/"><u>How to recover old videos from your Nokia C300</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-videos-from-nubia-z50-ultra-by-fonelab-android-recover-video/"><u>How to Rescue Lost Videos from Nubia Z50 Ultra</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-vivo-t2-pro-5g-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Vivo T2 Pro 5G to iPhone | Dr.fone</u></a></li>
<li><a href="https://android-unlock.techidaily.com/10-easy-to-use-frp-bypass-tools-for-unlocking-google-accounts-on-oppo-a38-by-drfone-android/"><u>10 Easy-to-Use FRP Bypass Tools for Unlocking Google Accounts On Oppo A38</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/how-to-get-and-use-pokemon-go-promo-codes-on-tecno-phantom-v-flip-drfone-by-drfone-virtual-android/"><u>How to Get and Use Pokemon Go Promo Codes On Tecno Phantom V Flip | Dr.fone</u></a></li>
<li><a href="https://activate-lock.techidaily.com/ultimate-guide-on-iphone-11-icloud-activation-lock-bypass-by-drfone-ios/"><u>Ultimate Guide on iPhone 11 iCloud Activation Lock Bypass</u></a></li>
<li><a href="https://android-location-track.techidaily.com/how-to-detect-and-remove-spyware-on-lava-storm-5g-drfone-by-drfone-virtual-android/"><u>How to Detect and Remove Spyware on Lava Storm 5G? | Dr.fone</u></a></li>
<li><a href="https://android-transfer.techidaily.com/in-2024-best-3-software-to-transfer-files-tofrom-your-oneplus-11r-via-a-usb-cable-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>In 2024, Best 3 Software to Transfer Files to/from Your OnePlus 11R via a USB Cable | Dr.fone</u></a></li>
<li><a href="https://ai-vdieo-software.techidaily.com/1714084011714-updated-best-video-editing-apps-for-windows-11-free-paid-and-everything-in-between/"><u>Updated Best Video Editing Apps for Windows 11 Free, Paid, and Everything in Between</u></a></li>
<li><a href="https://bypass-frp.techidaily.com/how-to-bypass-samsung-galaxy-xcover-7-frp-in-3-different-ways-by-drfone-android/"><u>How To Bypass Samsung Galaxy XCover 7 FRP In 3 Different Ways</u></a></li>
<li><a href="https://techidaily.com/tecno-spark-20-pro-video-recovery-recover-deleted-videos-from-tecno-spark-20-pro-by-fonelab-android-recover-video/"><u>Tecno Spark 20 Pro Video Recovery - Recover Deleted Videos from Tecno Spark 20 Pro</u></a></li>
<li><a href="https://android-location-track.techidaily.com/in-2024-how-to-track-lava-blaze-2-5g-location-by-number-drfone-by-drfone-virtual-android/"><u>In 2024, How to Track Lava Blaze 2 5G Location by Number | Dr.fone</u></a></li>
<li><a href="https://ios-unlock.techidaily.com/in-2024-learn-how-to-lock-stolen-your-apple-iphone-xs-max-properly-by-drfone-ios/"><u>In 2024, Learn How To Lock Stolen Your Apple iPhone XS Max Properly</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/3-methods-to-mirror-nokia-c210-to-roku-drfone-by-drfone-android/"><u>3 Methods to Mirror Nokia C210 to Roku | Dr.fone</u></a></li>
<li><a href="https://ai-editing-video.techidaily.com/in-2024-tactics-to-make-transparent-image-in-minimal-steps/"><u>In 2024, Tactics to Make Transparent Image in Minimal Steps</u></a></li>
<li><a href="https://animation-videos.techidaily.com/updated-in-2024-top-10-apps-to-enhance-your-instagram-story-slideshow/"><u>Updated In 2024, Top 10 Apps to Enhance Your Instagram Story Slideshow</u></a></li>
<li><a href="https://fake-location.techidaily.com/complete-tutorial-to-use-gps-joystick-to-fake-gps-location-on-oneplus-ace-2-drfone-by-drfone-virtual-android/"><u>Complete Tutorial to Use GPS Joystick to Fake GPS Location On OnePlus Ace 2 | Dr.fone</u></a></li>
<li><a href="https://ai-editing-video.techidaily.com/new-are-you-excited-to-learn-about-a-file-format-that-helps-in-easy-transfer-we-are-talking-about-the-3gp-format-stay-with-us-and-get-to-know-more-about-it./"><u>New Are You Excited to Learn About a File Format that Helps in Easy Transfer? We Are Talking About the 3GP Format. Stay with Us and Get to Know More About It</u></a></li>
<li><a href="https://location-social.techidaily.com/in-2024-how-to-changefake-your-samsung-galaxy-a23-5g-location-on-viber-drfone-by-drfone-virtual-android/"><u>In 2024, How to Change/Fake Your Samsung Galaxy A23 5G Location on Viber | Dr.fone</u></a></li>
<li><a href="https://techidaily.com/how-to-factory-reset-realme-narzo-60-5g-in-5-easy-ways-drfone-by-drfone-reset-android-reset-android/"><u>How to Factory Reset Realme Narzo 60 5G in 5 Easy Ways | Dr.fone</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/how-to-remove-or-bypass-knox-enrollment-service-on-nubia-by-drfone-android/"><u>How To Remove or Bypass Knox Enrollment Service On Nubia</u></a></li>
<li><a href="https://android-transfer.techidaily.com/in-2024-2-ways-to-transfer-text-messages-from-nubia-z50-ultra-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>In 2024, 2 Ways to Transfer Text Messages from Nubia Z50 Ultra to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
</ul></div>



---
title: How to fix runtime error 424 object required error in Excel 2003
date: 2024-05-19T18:32:11.884Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to fix runtime error 424 object required error in Excel 2003
keywords: repair .xls,repair excel 2007,repair .xltm,repair excel 2023,repair .xlsx,repair excel 2019
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
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-pictures-from-mix-fold-3-by-fonelab-android-recover-pictures/"><u>How to recover deleted pictures from Mix Fold 3.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-data-from-iphone-11-to-other-iphone-13-pro-devices-drfone-by-drfone-transfer-data-from-ios-transfer-data-from-ios/"><u>How To Transfer Data From iPhone 11 To Other iPhone 13 Pro devices? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-facetime-call-history-on-iphone-6s-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover FaceTime Call History on iPhone 6s | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-13-without-losing-any-data-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 13 without Losing Any Data? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-honor-using-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Honor using Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-v30t-photos-an-easy-method-explained-by-fonelab-android-recover-photos/"><u>How to Restore Deleted V30T Photos  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-vivo-x-flip-by-fonelab-android-recover-music/"><u>How to recover old music from your Vivo X Flip</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-call-logs-from-realme-11-5g-by-fonelab-android-recover-call-logs/"><u>How To  Restore Missing Call Logs from Realme 11 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-call-history-on-honor-x8b-by-fonelab-android-recover-call-logs/"><u>How to restore wiped call history on Honor X8b?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-google-frp-lock-on-hot-40-pro-by-drfone-android-unlock-remove-google-frp/"><u>How to remove Google FRP Lock on Hot 40 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-call-history-on-motorola-moto-g24-by-fonelab-android-recover-call-logs/"><u>How to restore wiped call history on Motorola Moto G24?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-data-from-iphone-13-using-stellar-data-recovery-for-iphone-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Data from iPhone 13 using Stellar Data Recovery for iPhone? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-csv-by-digital-signature-by-ldigisigner-sign-a-excel-sign-a-excel/"><u>How to sign .csv by digital signature</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-iphone-x-data-from-icloud-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover iPhone X Data From iCloud? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-oneplus-nord-3-5g-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from OnePlus Nord 3 5G to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-contacts-from-oppo-find-x7-ultra-by-fonelab-android-recover-contacts/"><u>How to recover deleted contacts from Oppo Find X7 Ultra.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-a-damaged-video-file-of-pop-7-pro-using-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair a Damaged video file of Pop 7 Pro using Video Repair Utility on Windows?</u></a></li>
<li><a href="https://android-transfer.techidaily.com/how-to-transfer-photos-from-xiaomi-redmi-13c-to-new-android-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Photos from Xiaomi Redmi 13C to New Android? | Dr.fone</u></a></li>
<li><a href="https://ai-video-editing.techidaily.com/updated-2024-approved-liquid-water-text-reveal-intro-offers-an-impressive-opening-to-any-video-content-go-through-the-detailed-steps-to-create-liquid-water-/"><u>Updated 2024 Approved Liquid Water Text Reveal Intro Offers an Impressive Opening to Any Video Content. Go Through the Detailed Steps to Create Liquid Water Reveal Intro in Wondershare Filmora</u></a></li>
<li><a href="https://android-location.techidaily.com/easy-ways-to-manage-your-oppo-find-x6-pro-location-settings-drfone-by-drfone-virtual/"><u>Easy Ways to Manage Your Oppo Find X6 Pro Location Settings | Dr.fone</u></a></li>
<li><a href="https://fake-location.techidaily.com/full-guide-to-fix-itoolab-anygo-not-working-on-vivo-x-flip-drfone-by-drfone-virtual-android/"><u>Full Guide to Fix iToolab AnyGO Not Working On Vivo X Flip | Dr.fone</u></a></li>
<li><a href="https://ios-unlock.techidaily.com/in-2024-forgot-iphone-6-backup-password-heres-what-to-do-by-drfone-ios/"><u>In 2024, Forgot iPhone 6 Backup Password? Heres What to Do</u></a></li>
<li><a href="https://android-location-track.techidaily.com/in-2024-3-solutions-to-find-your-nokia-130-music-current-location-of-a-mobile-number-drfone-by-drfone-virtual-android/"><u>In 2024, 3 Solutions to Find Your Nokia 130 Music Current Location of a Mobile Number | Dr.fone</u></a></li>
<li><a href="https://iphone-unlock.techidaily.com/in-2024-did-your-apple-iphone-se-passcode-change-itself-unlock-it-now-drfone-by-drfone-ios/"><u>In 2024, Did Your Apple iPhone SE Passcode Change Itself? Unlock It Now | Dr.fone</u></a></li>
<li><a href="https://ai-video-editing.techidaily.com/discover-the-top-6-online-video-speed-controllers-compatible-with-chrome-safari-firefox-for-enhanced-video-playback-and-viewing-experience/"><u>Discover the Top 6 Online Video Speed Controllers Compatible with Chrome, Safari, Firefox for Enhanced Video Playback and Viewing Experience</u></a></li>
<li><a href="https://techidaily.com/how-to-perform-hard-reset-on-motorola-moto-g23-drfone-by-drfone-reset-android-reset-android/"><u>How to Perform Hard Reset on Motorola Moto G23? | Dr.fone</u></a></li>
<li><a href="https://ai-video-apps.techidaily.com/in-2024-say-it-write-it-how-to-convert-spoken-words-to-written-text/"><u>In 2024, Say It, Write It How to Convert Spoken Words to Written Text</u></a></li>
<li><a href="https://fake-location.techidaily.com/thinking-about-changing-your-netflix-region-without-a-vpn-on-itel-a60-drfone-by-drfone-virtual-android/"><u>Thinking About Changing Your Netflix Region Without a VPN On Itel A60? | Dr.fone</u></a></li>
<li><a href="https://android-location.techidaily.com/fake-android-location-without-rooting-for-your-tecno-pop-8-drfone-by-drfone-virtual/"><u>Fake Android Location without Rooting For Your Tecno Pop 8 | Dr.fone</u></a></li>
<li><a href="https://fake-location.techidaily.com/ultimate-guide-to-free-pptp-vpn-for-beginners-on-vivo-x-flip-drfone-by-drfone-virtual-android/"><u>Ultimate Guide to Free PPTP VPN For Beginners On Vivo X Flip | Dr.fone</u></a></li>
</ul></div>



---
title: How to fix runtime error 424 object required error in Excel 2019
date: 2024-08-20T20:48:12.591Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to fix runtime error 424 object required error in Excel 2019
keywords: repair excel,repair .xlsm,repair excel 2000,repair .csv
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

<!-- affiliate ads begin -->
<a href="https://aspironcom.sjv.io/c/5597632/1941789/21554" target="_top" id="1941789"><img src="//a.impactradius-go.com/display-ad/21554-1941789" border="0" alt="" width="650" height="800"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1941789/21554" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Error When Incorrect Name Of The Object](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/error-424-when-incorrect-name-of-the-object.jpg)

 When you click the **Debug** button, the line with the error will highlight.

![Highlighting Line With Error](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/highlighting-line-with-error.jpg)

To fix the issue, you need to provide the correct name of the object.

### **2\. Check if the Object is Missing**

 The Runtime error 424 can occur if the object you are referring to as a method is not available or you are using the wrong object in a code. In the below example, you can see that the error occurs when an object named “Employee” is not available in the Project list.

<!-- affiliate ads begin -->
<a href="https://caperobbin.sjv.io/c/5597632/2006123/18460" target="_top" id="2006123"><img src="//a.impactradius-go.com/display-ad/18460-2006123" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2006123/18460" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Example Of Code When Object Is Not Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/example-of-code-when-an-object-is-not-available.jpg)

 You can check and mention the object which is available. For instance, Sheet2 in the below code.  

<!-- affiliate ads begin -->
<a href="https://ukaidot.sjv.io/c/5597632/1793237/19578" target="_top" id="1793237"><img src="//a.impactradius-go.com/display-ad/19578-1793237" border="0" alt="" width="1200" height="1200"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1793237/19578" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Check When The Object Is Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/check-when-the-object-is-available.png)

<!-- affiliate ads begin -->
<a href="https://dhgate.sjv.io/c/5597632/1678785/12108" target="_top" id="1678785"><img src="//a.impactradius-go.com/display-ad/12108-1678785" border="0" alt="" width="300" height="250"/></a>
<!-- affiliate ads end -->
### **3\. Check All References are Declared in the Code**

You can get the Runtime error 424 if all the references are not declared. So, make sure you have declared all the references in the code. To verify this, you can use the debug mode by pressing **F5** or clicking on the **Debug** option.

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2082535/7443" target="_top" id="2082535"><img src="//a.impactradius-go.com/display-ad/7443-2082535" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2082535/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Debug Command In Excel](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/debug-command-in-excel.jpg)

### **4\. Check the Macro Security Settings**

Sometimes, the error can occur if macros are disabled in the Macro Security settings. You can check and change the settings by following these steps:

- On the **Developer** tab, in the **Code** section, click **Macro Security**.
- In the **Trust Center** window, select **Enable all macros.**

![Macro Security Wizard](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/macro-security-wizard.jpg)

- Click **OK**.

<!-- affiliate ads begin -->
<a href="https://ship7com.pxf.io/c/5597632/1509856/17634" target="_top" id="1509856"><img src="//a.impactradius-go.com/display-ad/17634-1509856" border="0" alt="" width="730" height="383"/></a>
<!-- affiliate ads end -->
![Enable All Macro In Trust Center](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/enable-all-macros-in-trust-center.jpg)

### 5\. Repair your Workbook

Sometimes, the ‘Object required’ error can occur if your Excel file is damaged or corrupted. In such a case, you can try repairing the file using Microsoft’s in-built utility - Open and Repair. To use this utility, follow these steps:

- In Excel, go to **File > Open > Browse**.
- In the Open dialog box, click on the corrupted Excel file.
- Click the arrow next to the Open button and select **Open and Repair** from the dropdown.
- Select **Repair** to recover as much data from the file as possible.

If the Open and Repair utility fails or stops working, then you can try a professional Excel repair tool, such as [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). It is an advanced tool that can repair severely corrupted Excel files **(.xls, .xlsx, .xltm, .xltx, and .xlsm)**. It helps recover all the file components, including images, charts, tables, pivot tables, cell comments, chart sheets, formulas, etc., without impacting the original structure.

<!-- affiliate ads begin -->
<a href="https://thefitville.pxf.io/c/5597632/1526796/15852" target="_top" id="1526796"><img src="//a.impactradius-go.com/display-ad/15852-1526796" border="0" alt="" width="1200" height="628"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1526796/15852" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
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
<li><a href="https://remote-screen-capture.techidaily.com/new-in-2024-apowersoft-screen-record-unveiled-top-alternatives-reviewed/"><u>[New] In 2024, Apowersoft Screen Record Unveiled - Top Alternatives Reviewed</u></a></li>
<li><a href="https://fox-boxes.techidaily.com/new-in-2024-instant-anonymity-shield-face-details/"><u>[New] In 2024, Instant Anonymity  Shield Face Details</u></a></li>
<li><a href="https://video-screen-grab.techidaily.com/new-play-anywhere-anytime-the-ultimate-collection-of-the-best-gb-advance-emulators-on-pcs-for-2024/"><u>[New] Play Anywhere, Anytime  The Ultimate Collection of the Best GB Advance Emulators on PCs for 2024</u></a></li>
<li><a href="https://remote-screen-capture.techidaily.com/updated-2024-approved-a-new-dawn-in-gaming-stardew-valleys-best-mods-7/"><u>[Updated] 2024 Approved  A New Dawn in Gaming - Stardew Valley's Best Mods (#7)</u></a></li>
<li><a href="https://on-screen-recording.techidaily.com/updated-from-gameplay-to-video-full-guide-on-ps4-with-obs-for-2024/"><u>[Updated] From Gameplay to Video  Full Guide on PS4 with OBS for 2024</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-oppo-find-x7-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Oppo Find X7 to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-vivo-y200-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Vivo Y200 to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://extra-information.techidaily.com/2024-approved-box-it-up-the-top-10-online-houses-for-customized-gift-boxes/"><u>2024 Approved  Box It Up! The Top 10 Online Houses for Customized Gift Boxes</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-motorola-edge-40-neo-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Motorola Edge 40 Neo to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-techniques-to-transfer-data-from-motorola-moto-e13-to-iphone-15141312-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Techniques to Transfer Data from Motorola Moto E13 to iPhone 15/14/13/12 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-techniques-to-transfer-data-from-poco-c51-to-iphone-15141312-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Techniques to Transfer Data from Poco C51 to iPhone 15/14/13/12 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-techniques-to-transfer-data-from-vivo-s17e-to-iphone-15141312-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Techniques to Transfer Data from Vivo S17e to iPhone 15/14/13/12 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-ways-to-move-contacts-from-samsung-galaxy-a15-4g-to-iphone-131415-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Ways to Move Contacts From Samsung Galaxy A15 4G to iPhone (13/14/15) | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-ways-to-move-contacts-from-samsung-galaxy-f04-to-iphone-131415-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Ways to Move Contacts From Samsung Galaxy F04 to iPhone (13/14/15) | Dr.fone</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/apple-iphone-14-pro-max-mirror-to-pc-top-apps-you-must-know-drfone-by-drfone-ios/"><u>Apple iPhone 14 Pro Max Mirror to PC? Top Apps You Must Know | Dr.fone</u></a></li>
<li><a href="https://extra-hints.techidaily.com/crafting-pure-designs-erasing-backgrounds-with-canva-expertise/"><u>Crafting Pure Designs  Erasing Backgrounds with Canva Expertise</u></a></li>
<li><a href="https://bypass-frp.techidaily.com/how-to-bypass-google-frp-lock-from-xiaomi-14-pro-devices-by-drfone-android/"><u>How to Bypass Google FRP Lock from Xiaomi 14 Pro Devices</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-and-retrieve-picturesvideos-from-a-water-damaged-iphone-x-that-wont-turn-on-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Fix & Retrieve Pictures/Videos From a Water Damaged iPhone X That Wont Turn on | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-back-lost-photos-from-samsung-galaxy-a54-5g-by-fonelab-android-recover-photos/"><u>How to get back lost photos from Samsung Galaxy A54 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-install-and-update-device-drivers-manually-in-windows-11-and-10-by-drivereasy-guide/"><u>How to install and update device drivers manually in Windows 11 & 10</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-realme-c53-get-deleted-photos-back-with-ease-and-safety-by-fonelab-android-recover-photos/"><u>How to Realme C53 Get Deleted photos Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-after-format-on-samsung-galaxy-f15-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery after format on Samsung Galaxy F15 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-honor-100-by-fonelab-android-recover-photos/"><u>How to recover deleted photos from Honor 100.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-pictures-on-iphone-14-pro-5-best-solutions-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover deleted pictures on iPhone 14 Pro? 5 Best Solutions | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-iphone-x-after-ios-update-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Lost Data from iPhone X After iOS Update? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-infinix-note-30-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Infinix Note 30 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-google-frp-lock-on-vivo-s18-pro-by-drfone-android-unlock-remove-google-frp/"><u>How to remove Google FRP Lock on Vivo S18 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-mdm-from-iphone-xs-max-without-apple-id-by-drfone-ios-unlock-ios-unlock/"><u>How to Remove MDM from iPhone XS Max without Apple ID?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-samsung-galaxy-f54-5g-pin-by-drfone-android-unlock-android-unlock/"><u>How to remove Samsung Galaxy F54 5G PIN</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-iphone-se-2022-system-issues-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair iPhone SE (2022) System Issues? | Dr.fone</u></a></li>
<li><a href="https://unlock-android.techidaily.com/how-to-reset-gmail-password-on-tecno-camon-20-pro-5g-devices-by-drfone-android/"><u>How to Reset Gmail Password on Tecno Camon 20 Pro 5G Devices</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-messages-on-redmi-k70e-by-fonelab-android-recover-messages/"><u>How to restore wiped messages on Redmi K70E</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-videos-on-huawei-p60-by-fonelab-android-recover-video/"><u>How to restore wiped videos on Huawei P60</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-deleted-photos-on-vivo-y17s-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Retrieve  deleted photos on Vivo Y17s</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-deleted-photos-on-magic-vs-2-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Retrieve deleted photos on Magic Vs 2</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-call-logs-from-oppo-reno-8t-5g-by-fonelab-android-recover-call-logs/"><u>How to retrieve erased call logs from Oppo Reno 8T 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-messages-from-honor-magic-vs-2-by-fonelab-android-recover-messages/"><u>How to retrieve erased messages from Honor Magic Vs 2</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-messages-from-itel-p55-5g-by-fonelab-android-recover-messages/"><u>How to retrieve erased messages from Itel P55 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-a-pdf-v16-document-with-digital-signature-service-by-ldigisigner-sign-a-pdf-sign-a-pdf/"><u>How to sign a PDF v1.6 document with digital signature service</u></a></li>
<li><a href="https://iphone-location.techidaily.com/in-2024-4-effective-methods-fake-gps-location-on-apple-iphone-xripad-drfone-by-drfone-virtual-ios/"><u>In 2024, 4 Effective Methods Fake GPS Location on Apple iPhone XR/iPad | Dr.fone</u></a></li>
<li><a href="https://some-knowledge.techidaily.com/in-2024-exploring-the-skies-with-top-10-beginner-drones/"><u>In 2024, Exploring the Skies with Top 10 Beginner Drones</u></a></li>
<li><a href="https://android-unlock.techidaily.com/in-2024-how-to-bypass-android-lock-screen-using-emergency-call-on-samsung-galaxy-f04-by-drfone-android/"><u>In 2024, How to Bypass Android Lock Screen Using Emergency Call On Samsung Galaxy F04?</u></a></li>
<li><a href="https://iphone-unlock.techidaily.com/in-2024-how-to-remove-and-reset-face-id-on-apple-iphone-11-drfone-by-drfone-ios/"><u>In 2024, How to Remove and Reset Face ID on Apple iPhone 11 | Dr.fone</u></a></li>
<li><a href="https://location-social.techidaily.com/in-2024-set-your-preferred-job-location-on-linkedin-app-of-your-oppo-a59-5g-drfone-by-drfone-virtual-android/"><u>In 2024, Set Your Preferred Job Location on LinkedIn App of your Oppo A59 5G | Dr.fone</u></a></li>
<li><a href="https://win-answers.techidaily.com/overcome-multiplayer-malfunctions-in-elden-ring-with-these-tips/"><u>Overcome Multiplayer Malfunctions in Elden Ring with These Tips</u></a></li>
<li><a href="https://tiktok-videos.techidaily.com/spotting-blocked-contacts-in-snapchat/"><u>Spotting Blocked Contacts in Snapchat</u></a></li>
<li><a href="https://instagram-video-files.techidaily.com/sync-video-elements-to-captivate-instagram-audiences/"><u>Sync Video Elements to Captivate Instagram Audiences</u></a></li>
<li><a href="https://howto.techidaily.com/want-to-uninstall-google-play-service-from-tecno-spark-10c-here-is-how-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>Want to Uninstall Google Play Service from Tecno Spark 10C? Here is How | Dr.fone</u></a></li>
</ul></div>

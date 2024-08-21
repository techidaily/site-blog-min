---
title: How to fix runtime error 424 object required error in Excel 2021
date: 2024-08-20T20:48:14.116Z
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

<!-- affiliate ads begin -->
<a href="https://vapordna.pxf.io/c/5597632/1494880/17238" target="_top" id="1494880"><img src="//a.impactradius-go.com/display-ad/17238-1494880" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1494880/17238" style="position:absolute;visibility:hidden;" border="0" />
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

##  **Solutions to Fix Runtime Error 424: Object Required in Excel**

The VBA error ‘object required’ may occur due to different reasons. Based on the reason, you can follow the solutions mentioned below to fix the error.

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2087484/7443" target="_top" id="2087484"><img src="//a.impactradius-go.com/display-ad/7443-2087484" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2087484/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **1\. Check the Name of the Object**

The Runtime error 424 can occur when you run the VBA code using an incorrect name of the object. For example, the object name is ‘MyObject’ but you’re using “Backcolor”.

![Error When Incorrect Name Of The Object](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/error-424-when-incorrect-name-of-the-object.jpg)

 When you click the **Debug** button, the line with the error will highlight.

<!-- affiliate ads begin -->
<a href="https://ukaidot.sjv.io/c/5597632/1793233/19578" target="_top" id="1793233"><img src="//a.impactradius-go.com/display-ad/19578-1793233" border="0" alt="" width="1200" height="1200"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1793233/19578" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Highlighting Line With Error](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/highlighting-line-with-error.jpg)

To fix the issue, you need to provide the correct name of the object.

### **2\. Check if the Object is Missing**

 The Runtime error 424 can occur if the object you are referring to as a method is not available or you are using the wrong object in a code. In the below example, you can see that the error occurs when an object named “Employee” is not available in the Project list.

![Example Of Code When Object Is Not Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/example-of-code-when-an-object-is-not-available.jpg)

 You can check and mention the object which is available. For instance, Sheet2 in the below code.  

<!-- affiliate ads begin -->
<a href="https://martinic.evyy.net/c/5597632/1422856/4482" target="_top" id="1422856"><img src="//a.impactradius-go.com/display-ad/4482-1422856" border="0" alt="" width="580" height="309"/></a>
<!-- affiliate ads end -->
![Check When The Object Is Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/check-when-the-object-is-available.png)

### **3\. Check All References are Declared in the Code**

You can get the Runtime error 424 if all the references are not declared. So, make sure you have declared all the references in the code. To verify this, you can use the debug mode by pressing **F5** or clicking on the **Debug** option.

<!-- affiliate ads begin -->
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=3727260&QTY=1&AFFILIATE=108875&CART=1"><img src="http://www.aiseesoft.com/avangate/30p/banner.jpg" border="0"></a>
<!-- affiliate ads end -->
![Debug Command In Excel](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/debug-command-in-excel.jpg)

<!-- affiliate ads begin -->
<a href="https://aidotcom.pxf.io/c/5597632/2086436/19576" target="_top" id="2086436"><img src="//a.impactradius-go.com/display-ad/19576-2086436" border="0" alt="" width="1500" height="400"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2086436/19576" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **4\. Check the Macro Security Settings**

Sometimes, the error can occur if macros are disabled in the Macro Security settings. You can check and change the settings by following these steps:

- On the **Developer** tab, in the **Code** section, click **Macro Security**.
- In the **Trust Center** window, select **Enable all macros.**

![Macro Security Wizard](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/macro-security-wizard.jpg)

- Click **OK**.

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2082535/7443" target="_top" id="2082535"><img src="//a.impactradius-go.com/display-ad/7443-2082535" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2082535/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
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
<li><a href="https://some-guidance.techidaily.com/new-unlocking-untapped-utility-zoom-in-on-minecraft/"><u>[New] Unlocking Untapped Utility  Zoom in on Minecraft</u></a></li>
<li><a href="https://win-answers.techidaily.com/solved-msedgeexe-application-error-on-windows/"><u>[Solved] msedge.exe Application Error on Windows</u></a></li>
<li><a href="https://facebook-record-videos.techidaily.com/updated-deception-navigating-false-subscriber-sells-for-2024/"><u>[Updated] Deception  Navigating False Subscriber Sells for 2024</u></a></li>
<li><a href="https://extra-guidance.techidaily.com/updated-mastering-blur-apps-for-clear-portraits/"><u>[Updated] Mastering Blur Apps for Clear Portraits</u></a></li>
<li><a href="https://extra-support.techidaily.com/updated-precision-in-panorama-steadying-the-gopro-sight/"><u>[Updated] Precision in Panorama  Steadying the GoPro Sight</u></a></li>
<li><a href="https://facebook-video-footage.techidaily.com/updated-skyrocketing-view-figures-by-sustaining-youtubes-creative-commons-license/"><u>[Updated] Skyrocketing View Figures by Sustaining YouTube's Creative Commons License</u></a></li>
<li><a href="https://android-unlock.techidaily.com/10-easy-to-use-frp-bypass-tools-for-unlocking-google-accounts-on-lava-by-drfone-android/"><u>10 Easy-to-Use FRP Bypass Tools for Unlocking Google Accounts On Lava</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-vivo-x100-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Vivo X100 to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-poco-x6-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Poco X6 to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-ways-to-move-contacts-from-vivo-v29-pro-to-iphone-131415-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Ways to Move Contacts From Vivo V29 Pro to iPhone (13/14/15) | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-ways-to-transfer-music-from-motorola-moto-g13-to-other-android-devices-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Ways to Transfer Music from Motorola Moto G13 to Other Android Devices Easily | Dr.fone</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/catchemall-celebrate-national-pokemon-day-with-virtual-location-on-itel-p40plus-drfone-by-drfone-virtual-android/"><u>CatchEmAll Celebrate National Pokémon Day with Virtual Location On Itel P40+ | Dr.fone</u></a></li>
<li><a href="https://extra-information.techidaily.com/discover-best-4k-gimbals-for-mirrorless-cameras/"><u>Discover Best 4K Gimbals for Mirrorless Cameras</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-call-logs-from-infinix-smart-8-hd-by-fonelab-android-recover-call-logs/"><u>How To  Restore Missing Call Logs from Infinix Smart 8 HD</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-12-pro-max-to-the-previous-ios-version-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 12 Pro Max to the Previous iOS Version? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-corrupt-video-files-of-f3-pro-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Fix corrupt video files of F3 Pro using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-iphone-13-pro-max-stuck-at-attempting-data-recovery-loop-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Fix iPhone 13 Pro Max Stuck at attempting data recovery Loop | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-play-hevc-h-265-video-on-htc-u23-pro-by-aiseesoft-video-converter-play-hevc-video-on-android/"><u>How to play HEVC H.265 video on HTC U23 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-data-from-broken-iphone-8-plus-screen-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Data from Broken iPhone 8 Plus Screen | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-app-on-tecno-pova-5-pro-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Recover Deleted Photos from Android Gallery App on Tecno Pova 5 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-huawei-nova-y91-by-fonelab-android-recover-photos/"><u>How to recover deleted photos from Huawei Nova Y91.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-pictures-from-realme-11-proplus-by-fonelab-android-recover-pictures/"><u>How to recover deleted pictures from Realme 11 Pro+.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-screenshots-on-iphone-se-2020-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Screenshots on iPhone SE (2020)? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-12-pro-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 12 Pro? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-motorola-edge-40-by-fonelab-android-recover-music/"><u>How to recover old music from your Motorola Edge 40</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-itel-a70-pin-by-drfone-android-unlock-android-unlock/"><u>How to remove Itel A70 PIN</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-samsung-galaxy-a25-5g-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Samsung Galaxy A25 5G? </u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-iphone-6-system-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair iPhone 6 System? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-messages-from-galaxy-a05s-by-fonelab-android-recover-messages/"><u>How to Rescue Lost Messages from Galaxy A05s</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-photos-from-pixel-8-by-fonelab-android-recover-photos/"><u>How to Rescue Lost Photos from Pixel 8?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-pictures-from-honor-100-pro-by-fonelab-android-recover-pictures/"><u>How to Rescue Lost Pictures from Honor 100 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-15-pro-without-itunes-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Reset iPhone 15 Pro Without iTunes? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-music-from-realme-c33-2023-by-fonelab-android-recover-music/"><u>How to retrieve erased music from Realme C33 2023</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-dot-file-document-with-electronic-digital-signature-tutorial-by-ldigisigner-sign-a-word-sign-a-word/"><u>How to Sign .dot file document with Electronic Digital Signature - (Tutorial)</u></a></li>
<li><a href="https://fox-cloud.techidaily.com/in-2024-digital-media-converter-for-phone-owners/"><u>In 2024, Digital Media Converter for Phone Owners</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/in-2024-is-pgsharp-legal-when-you-are-playing-pokemon-on-itel-p55plus-drfone-by-drfone-virtual-android/"><u>In 2024, Is pgsharp legal when you are playing pokemon On Itel P55+? | Dr.fone</u></a></li>
<li><a href="https://extra-resources.techidaily.com/omnistreamer-2024-local-and-live-the-best-package/"><u>OmniStreamer 2024  Local & Live, The Best Package</u></a></li>
</ul></div>

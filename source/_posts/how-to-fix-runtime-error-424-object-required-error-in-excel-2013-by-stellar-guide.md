---
title: How to fix runtime error 424 object required error in Excel 2013
date: 2024-08-20T20:48:10.166Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to fix runtime error 424 object required error in Excel 2013
keywords: repair .csv,repair .xlsm,repair .xls,repair .xltm,repair excel 2010
---

## How to fix runtime error 424 object required error in Excel

The Runtime error 424: Object required occurs when Excel is not able to recognize an object that you are referring to in a VBA code. The object can be a workbook, worksheet, range, variable, class, macro, etc. Some users have also reported that this error occurred when they tried to copy the values of the cells from one workbook to another.

Let’s understand the error through a small scenario. Suppose, I want to check the last field row in a table in a spreadsheet named “First” using the VBA code. To do this, I have added a command button and double-clicked on it and entered the below code in the backend:

Private Sub CommandButton2\_Click()

Dim LRow As Integer

LRow = Worksheets("First").Cells(Rows.Count, 2).End(xlUp).Row

MsgBox ("Last Row " & LRow)

End Sub

<!-- affiliate ads begin -->
<a href="https://shop.mondly.com/affiliate.php?ACCOUNT=ATISTUDI&AFFILIATE=108875&PATH=https%3A%2F%2Fwww.mondly.com%3FAFFILIATE%3D108875%26RESOURCE%3D%2BGeneral%2B970x90%2B"><img src="https://secure.avangate.com/images/merchant/69c418c33ec2e1a4267fa9bb77fa1428/general-970x90.gif" border="0"></a>
<!-- affiliate ads end -->
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

<!-- affiliate ads begin -->
<a href="https://arkmc.pxf.io/c/5597632/427477/5172" target="_top" id="427477"><img src="//a.impactradius-go.com/display-ad/5172-427477" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://arkmc.pxf.io/i/5597632/427477/5172" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **1\. Check the Name of the Object**

The Runtime error 424 can occur when you run the VBA code using an incorrect name of the object. For example, the object name is ‘MyObject’ but you’re using “Backcolor”.

![Error When Incorrect Name Of The Object](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/error-424-when-incorrect-name-of-the-object.jpg)

 When you click the **Debug** button, the line with the error will highlight.

![Highlighting Line With Error](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/highlighting-line-with-error.jpg)

To fix the issue, you need to provide the correct name of the object.

### **2\. Check if the Object is Missing**

 The Runtime error 424 can occur if the object you are referring to as a method is not available or you are using the wrong object in a code. In the below example, you can see that the error occurs when an object named “Employee” is not available in the Project list.

<!-- affiliate ads begin -->
<a href="https://printrendy.pxf.io/c/5597632/1453720/17020" target="_top" id="1453720"><img src="//a.impactradius-go.com/display-ad/17020-1453720" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1453720/17020" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Example Of Code When Object Is Not Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/example-of-code-when-an-object-is-not-available.jpg)

 You can check and mention the object which is available. For instance, Sheet2 in the below code.  

<!-- affiliate ads begin -->
<iframe id="iframe_672" src="//a.impactradius-go.com/gen-ad-code/5597632/1959812/17834/" width="720" height="300" scrolling="no" frameborder="0" marginheight="0" marginwidth="0"></iframe>
<!-- affiliate ads end -->
![Check When The Object Is Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/check-when-the-object-is-available.png)

### **3\. Check All References are Declared in the Code**

You can get the Runtime error 424 if all the references are not declared. So, make sure you have declared all the references in the code. To verify this, you can use the debug mode by pressing **F5** or clicking on the **Debug** option.

<!-- affiliate ads begin -->
<a href="https://homestyler.sjv.io/c/5597632/2044747/22993" target="_top" id="2044747"><img src="//a.impactradius-go.com/display-ad/22993-2044747" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2044747/22993" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Debug Command In Excel](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/debug-command-in-excel.jpg)

### **4\. Check the Macro Security Settings**

Sometimes, the error can occur if macros are disabled in the Macro Security settings. You can check and change the settings by following these steps:

- On the **Developer** tab, in the **Code** section, click **Macro Security**.
- In the **Trust Center** window, select **Enable all macros.**

<!-- affiliate ads begin -->
<a href="https://shop.copernic.com/order/checkout.php?PRODS=41033101&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.2checkout.com/images/merchant/8d30aa96e72440759f74bd2306c1fa3d/Copernic-2023-Affiliate-728x90-Elite.png" border="0"></a>
<!-- affiliate ads end -->
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

<!-- affiliate ads begin -->
<a href="https://turbotech.pxf.io/c/5597632/1450763/17212" target="_top" id="1450763"><img src="//a.impactradius-go.com/display-ad/17212-1450763" border="0" alt="" width="2560" height="1440"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1450763/17212" style="position:absolute;visibility:hidden;" border="0" />
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
<li><a href="https://fox-http.techidaily.com/new-in-2024-superior-photography-tech-the-prime-10-4k-smartphone-cameras/"><u>[New] In 2024, Superior Photography Tech  The Prime 10 4K Smartphone Cameras</u></a></li>
<li><a href="https://instagram-videos.techidaily.com/new-revamp-your-igtv-titles-quickly-and-efficiently/"><u>[New] Revamp Your IGTV Titles Quickly & Efficiently</u></a></li>
<li><a href="https://extra-skills.techidaily.com/new-speed-up-srt-to-text-conversion-with-this-essential-guide/"><u>[New] Speed Up SRT to Text Conversion with This Essential Guide</u></a></li>
<li><a href="https://some-skills.techidaily.com/new-unveiling-the-slow-motion-canvas-detailed-slomo-app-analysis-2024/"><u>[New] Unveiling the Slow Motion Canvas  Detailed SloMo App Analysis, 2024</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-oppo-reno-9a-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Oppo Reno 9A to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://some-knowledge.techidaily.com/from-novice-to-pro-periscope-stream-mastery-for-2024/"><u>From Novice to Pro  Periscope Stream Mastery for 2024</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-call-logs-from-realme-11-5g-by-fonelab-android-recover-call-logs/"><u>How To  Restore Missing Call Logs from Realme 11 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-contacts-files-from-motorola-moto-g-stylus-5g-2023-by-fonelab-android-recover-contacts/"><u>How To  Restore Missing Contacts Files from Motorola Moto G Stylus 5G (2023).</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-the-value-error-in-excel-2021-by-stellar-guide/"><u>How to Fix the #Value! Error in Excel 2021?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-videos-not-playing-with-my-edge-40-by-stellar-video-repair-mobile-video-repair/"><u>How to fix videos not playing with my Edge 40?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-back-lost-photos-from-gt-10-pro-by-fonelab-android-recover-photos/"><u>How to get back lost photos from GT 10 Pro.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-insert-signature-in-xltx-files-by-ldigisigner-sign-a-excel-sign-a-excel/"><u>How to insert signature in .xltx files</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-realme-10t-5g-get-deleted-photos-back-with-ease-and-safety-by-fonelab-android-recover-photos/"><u>How to Realme 10T 5G Get Deleted photos Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-contacts-from-infinix-note-30-vip-by-fonelab-android-recover-contacts/"><u>How to recover deleted contacts from Infinix Note 30 VIP.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-oppo-reno-9a-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Oppo Reno 9A</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-pictures-from-samsung-by-fonelab-android-recover-pictures/"><u>How to recover deleted pictures from Samsung .</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-c110-by-fonelab-android-recover-music/"><u>How to recover old music from your C110</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-nokia-c210-by-fonelab-android-recover-music/"><u>How to recover old music from your Nokia C210</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-videos-from-your-p40-by-fonelab-android-recover-video/"><u>How to recover old videos from your P40</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-poco-x5-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Poco X5 on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-reno-8t-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Reno 8T on Windows??</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-vivo-x-fold-2-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Vivo X Fold 2 using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-your-iphone-8-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Reset Your iPhone 8? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-honor-90-pro-contacts-an-easy-method-explained-by-fonelab-android-recover-contacts/"><u>How to Restore Deleted Honor 90 Pro Contacts  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-lava-yuva-3-pro-pictures-an-easy-method-explained-by-fonelab-android-recover-pictures/"><u>How to Restore Deleted Lava Yuva 3 Pro Pictures  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-messages-on-honor-90-gt-by-fonelab-android-recover-messages/"><u>How to restore wiped messages on Honor 90 GT</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-videos-from-tecno-spark-go-2024-by-fonelab-android-recover-video/"><u>How to retrieve erased videos from Tecno Spark Go (2024)</u></a></li>
<li><a href="https://extra-lessons.techidaily.com/in-2024-comprehensive-setup-steps-for-wm6/"><u>In 2024, Comprehensive Setup Steps for WM6</u></a></li>
<li><a href="https://location-social.techidaily.com/in-2024-why-your-whatsapp-live-location-is-not-updating-and-how-to-fix-on-your-motorola-moto-g73-5g-drfone-by-drfone-virtual-android/"><u>In 2024, Why Your WhatsApp Live Location is Not Updating and How to Fix on your Motorola Moto G73 5G | Dr.fone</u></a></li>
<li><a href="https://win11.techidaily.com/remedying-disconnect-error-for-malwarebytes-services-in-windows-11/"><u>Remedying Disconnect Error for Malwarebytes Services in Windows 11</u></a></li>
<li><a href="https://extra-lessons.techidaily.com/the-unseen-windows-11-techniques/"><u>The Unseen Windows 11 Techniques</u></a></li>
<li><a href="https://screen-capture.techidaily.com/understanding-and-modifying-screen-capture-formats-mac/"><u>Understanding and Modifying Screen Capture Formats (Mac)</u></a></li>
<li><a href="https://techno-recovery.techidaily.com/upgrade-or-replace-maintaining-your-ipads-energy-lifeline/"><u>Upgrade or Replace? Maintaining Your iPad's Energy Lifeline</u></a></li>
</ul></div>

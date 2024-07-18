---
title: How to fix runtime error 424 object required error in Excel 2021
date: 2024-07-17T10:00:02.066Z
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
<li><a href="https://twitter-videos.techidaily.com/new-2024-approved-twitters-quick-fix-for-downloading-funny-images-on-pc/"><u>[New] 2024 Approved  Twitter's Quick-Fix for Downloading Funny Images on PC</u></a></li>
<li><a href="https://facebook-video-footage.techidaily.com/new-exclusive-free-mcb-visual-tools-for-2024/"><u>[New] Exclusive Free MCB Visual Tools for 2024</u></a></li>
<li><a href="https://discord-videos.techidaily.com/new-expertly-engineered-virtual-assistants-discord/"><u>[New] Expertly Engineered Virtual Assistants (Discord)</u></a></li>
<li><a href="https://instagram-clips.techidaily.com/updated-2024-approved-comprehensible-guide-to-swapping-facial-gender-in-snapchat-images/"><u>[Updated] 2024 Approved  Comprehensible Guide to Swapping Facial Gender in Snapchat Images</u></a></li>
<li><a href="https://facebook-video-footage.techidaily.com/updated-in-2024-discover-12-cutting-edge-flip-screen-cams-for-video-content/"><u>[Updated] In 2024, Discover 12 Cutting-Edge Flip-Screen Cams for Video Content</u></a></li>
<li><a href="https://fox-boxes.techidaily.com/2024-approved-digital-audio-translations-from-srt-to-xmlssa-ttml-etc/"><u>2024 Approved  Digital Audio Translations  From SRT to XML/SSA, TTML, Etc</u></a></li>
<li><a href="https://extra-guidance.techidaily.com/2024-approved-online-gag-artist/"><u>2024 Approved  Online Gag Artist</u></a></li>
<li><a href="https://audio-shaping.techidaily.com/2024-approved-strategies-for-eliminating-reverb-from-studio-recorded-sounds/"><u>2024 Approved Strategies for Eliminating Reverb From Studio-Recorded Sounds</u></a></li>
<li><a href="https://youtube-tips.techidaily.com/casting-made-simple-directly-stream-google-meet-to-youtube-for-2024/"><u>Broadcasting Made Simple  Directly Stream Google Meet to YouTube for 2024</u></a></li>
<li><a href="https://some-techniques.techidaily.com/expert-tips-for-utilizing-snapchat-highlights-effectively-for-2024/"><u>Expert Tips for Utilizing Snapchat Highlights Effectively for 2024</u></a></li>
<li><a href="https://activate-lock.techidaily.com/how-to-delete-icloud-account-with-or-without-password-from-your-iphone-14-pluswindowsmac-by-drfone-ios/"><u>How to Delete iCloud Account with or without Password from your iPhone 14 Plus/Windows/Mac</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-nokia-g42-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Nokia G42 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-nokia-xr21-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Nokia XR21?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-nord-3-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Nord 3 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-note-30-pro-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Note 30 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-nubia-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Nubia ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-nubia-red-magic-8s-pro-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Nubia Red Magic 8S Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-nubia-red-magic-8s-proplus-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Nubia Red Magic 8S Pro+?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-oneplus-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your OnePlus ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-oneplus-ace-2-pro-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your OnePlus Ace 2 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-oppo-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Oppo ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-oppo-a18-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Oppo A18?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-oppo-a2-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Oppo A2?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-oppo-a78-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Oppo A78?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-oppo-find-x6-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Oppo Find X6?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-oppo-k11-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Oppo K11 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-oppo-reno-10-proplus-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Oppo Reno 10 Pro+ 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-oppo-reno-9a-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Oppo Reno 9A?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-phantom-v-fold-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Phantom V Fold?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-poco-c65-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Poco C65?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-poco-f5-pro-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Poco F5 Pro 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-poco-m6-pro-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Poco M6 Pro 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-poco-x6-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Poco X6?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-realme-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Realme ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-realme-11-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Realme 11 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-realme-narzo-n55-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Realme Narzo N55?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-realme-note-50-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Realme Note 50?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-redmi-note-12-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Redmi Note 12 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-redmi-note-12-proplus-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Redmi Note 12 Pro+ 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-samsung-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Samsung ?</u></a></li>
<li><a href="https://driver-install.techidaily.com/rejuvenate-window-display-with-simple-tweaks/"><u>Rejuvenate Window Display with Simple Tweaks</u></a></li>
<li><a href="https://instagram-clips.techidaily.com/the-untold-chronicles-adventures-in-antique-literature/"><u>The Untold Chronicles  Adventures in Antique Literature</u></a></li>
<li><a href="https://facebook-clips.techidaily.com/unseen-pathways-to-friend-finder-excellence-for-2024/"><u>Unseen Pathways to Friend Finder Excellence for 2024</u></a></li>
<li><a href="https://video-ai-editor.techidaily.com/updated-say-goodbye-to-watermarks-top-tiktok-removal-tools/"><u>Updated Say Goodbye to Watermarks Top TikTok Removal Tools</u></a></li>
</ul></div>

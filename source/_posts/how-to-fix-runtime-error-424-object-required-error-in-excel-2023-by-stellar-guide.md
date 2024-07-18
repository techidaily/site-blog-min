---
title: How to fix runtime error 424 object required error in Excel 2023
date: 2024-07-17T10:00:04.824Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to fix runtime error 424 object required error in Excel 2023
keywords: repair excel 2021,repair excel 2019,repair excel 2007,repair excel,repair excel 2023,repair excel 2010
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

<!-- affiliate ads begin -->
<a href="https://tinyland.pxf.io/c/5597632/1793214/19135" target="_top" id="1793214"><img src="//a.impactradius-go.com/display-ad/19135-1793214" border="0" alt="" width="900" height="900"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1793214/19135" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
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
<a href="https://tokenmetrics.sjv.io/c/5597632/1864921/20702" target="_top" id="1864921"><img src="//a.impactradius-go.com/display-ad/20702-1864921" border="0" alt="" width="1251" height="1042"/></a>
<!-- affiliate ads end -->
![Error When Incorrect Name Of The Object](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/error-424-when-incorrect-name-of-the-object.jpg)

 When you click the **Debug** button, the line with the error will highlight.

![Highlighting Line With Error](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/highlighting-line-with-error.jpg)

To fix the issue, you need to provide the correct name of the object.

### **2\. Check if the Object is Missing**

 The Runtime error 424 can occur if the object you are referring to as a method is not available or you are using the wrong object in a code. In the below example, you can see that the error occurs when an object named “Employee” is not available in the Project list.

![Example Of Code When Object Is Not Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/example-of-code-when-an-object-is-not-available.jpg)

 You can check and mention the object which is available. For instance, Sheet2 in the below code.  

![Check When The Object Is Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/check-when-the-object-is-available.png)

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2075482/7443" target="_top" id="2075482"><img src="//a.impactradius-go.com/display-ad/7443-2075482" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2075482/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **3\. Check All References are Declared in the Code**

You can get the Runtime error 424 if all the references are not declared. So, make sure you have declared all the references in the code. To verify this, you can use the debug mode by pressing **F5** or clicking on the **Debug** option.

<!-- affiliate ads begin -->
<a href="https://aofit.pxf.io/c/5597632/1399701/16396" target="_top" id="1399701"><img src="//a.impactradius-go.com/display-ad/16396-1399701" border="0" alt="" width="960" height="300"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1399701/16396" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Debug Command In Excel](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/debug-command-in-excel.jpg)

<!-- affiliate ads begin -->
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=4940312&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/333ac5d90817d69113471fbb6e531bee/sps-partnership-728x90eng.png" border="0"></a>
<!-- affiliate ads end -->
### **4\. Check the Macro Security Settings**

Sometimes, the error can occur if macros are disabled in the Macro Security settings. You can check and change the settings by following these steps:

- On the **Developer** tab, in the **Code** section, click **Macro Security**.
- In the **Trust Center** window, select **Enable all macros.**

![Macro Security Wizard](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/macro-security-wizard.jpg)

- Click **OK**.

<!-- affiliate ads begin -->
<a href="https://aligracehair.sjv.io/c/5597632/2087267/19272" target="_top" id="2087267"><img src="//a.impactradius-go.com/display-ad/19272-2087267" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2087267/19272" style="position:absolute;visibility:hidden;" border="0" />
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
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=3727260&QTY=1&AFFILIATE=108875&CART=1"><img src="http://www.aiseesoft.com/avangate/30p/banner.jpg" border="0"></a>
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
<li><a href="https://facebook-video-footage.techidaily.com/new-2024-approved-becoming-a-filmora-fcc-the-pathway-explained/"><u>[New] 2024 Approved  Becoming a Filmora FCC  The Pathway Explained</u></a></li>
<li><a href="https://discord-videos.techidaily.com/new-2024-approved-make-your-discord-messages-hearable-learn-tts/"><u>[New] 2024 Approved  Make Your Discord Messages Hearable  Learn TTS</u></a></li>
<li><a href="https://digital-screen-recording.techidaily.com/new-best-choices-for-switchs-hd-gaming-for-2024/"><u>[New] Best Choices for Switch's HD Gaming for 2024</u></a></li>
<li><a href="https://youtube-web.techidaily.com/rive-engagement-crafting-an-animated-subscribe-button-in-filmoras-step-by-step-guide/"><u>[New] Drive Engagement  Crafting an Animated Subscribe Button in Filmora's Step-by-Step Guide</u></a></li>
<li><a href="https://vimeo-videos.techidaily.com/updated-2024-approved-stepwise-guide-to-convert-vimeo-media-to-mp3/"><u>[Updated] 2024 Approved  Stepwise Guide to Convert Vimeo Media to MP3</u></a></li>
<li><a href="https://extra-resources.techidaily.com/updated-conquering-challenges-advanced-hdr-workflows-in-ps/"><u>[Updated] Conquering Challenges  Advanced HDR Workflows in PS</u></a></li>
<li><a href="https://facebook-video-share.techidaily.com/updated-jake-paul-from-dabbing-to-dominance-youtube-triumph/"><u>[Updated] Jake Paul  From Dabbing to Dominance - YouTube Triumph</u></a></li>
<li><a href="https://video-screen-grab.techidaily.com/updated-top-12-independent-recorders-for-pcmac-for-2024/"><u>[Updated] Top 12 Independent Recorders for PC/Mac for 2024</u></a></li>
<li><a href="https://extra-hints.techidaily.com/2024-approved-branding-beats-technicalities-focus-on-your-streams-signature-style/"><u>2024 Approved  Branding Beats Technicalities  Focus on Your Stream's Signature Style</u></a></li>
<li><a href="https://extra-hints.techidaily.com/2024-approved-canon-image-masters-10plus-luts-some-are-free/"><u>2024 Approved  Canon Image Masters - 10+ LUTs, Some Are Free</u></a></li>
<li><a href="https://instagram-videos.techidaily.com/2024-approved-craft-your-own-story-the-top-5-instagram-tips-for-aspiring-social-stars/"><u>2024 Approved  Craft Your Own Story  The Top 5 Instagram Tips for Aspiring Social Stars</u></a></li>
<li><a href="https://remote-screen-capture.techidaily.com/2024-approved-mastering-screen-casts-key-strategies-for-perfection/"><u>2024 Approved  Mastering Screen Casts  Key Strategies for Perfection</u></a></li>
<li><a href="https://youtube-help.techidaily.com/2024-approved-pro-video-editing-the-ultimate-list-of-10-cutter-apps/"><u>2024 Approved  Pro Video Editing  The Ultimate List of 10 Cutter Apps</u></a></li>
<li><a href="https://youtube-help.techidaily.com/2024-approved-the-ultimate-strategy-for-crafting-viral-gadget-unboxings/"><u>2024 Approved  The Ultimate Strategy for Crafting Viral Gadget Unboxings</u></a></li>
<li><a href="https://extra-hints.techidaily.com/core-tenets-of-narrative-construction-for-2024/"><u>Core Tenets of Narrative Construction for 2024</u></a></li>
<li><a href="https://youtube-stream.techidaily.com/fuel-the-drive-best-video-ideas-for-channels-success-for-2024/"><u>Fuel the Drive  Best Video Ideas for Channels' Success for 2024</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-y28-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Y28 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-y36i-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Y36i?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-y78t-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Y78t?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-zero-5g-2023-turbo-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Zero 5G 2023 Turbo?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-zte-axon-40-lite-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your ZTE Axon 40 Lite?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-zte-blade-a73-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your ZTE Blade A73 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-zte-nubia-z60-ultra-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your ZTE Nubia Z60 Ultra?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-105-classic-by-fonelab-android-recover-messages/"><u>How to recover old messages from your 105 Classic</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-10t-5g-by-fonelab-android-recover-messages/"><u>How to recover old messages from your 10T 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-12-pro-5g-by-fonelab-android-recover-messages/"><u>How to recover old messages from your 12 Pro 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-12-proplus-5g-by-fonelab-android-recover-messages/"><u>How to recover old messages from your 12 Pro+ 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-asus-rog-phone-8-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Asus ROG Phone 8</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-asus-rog-phone-8-pro-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Asus ROG Phone 8 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-c65-by-fonelab-android-recover-messages/"><u>How to recover old messages from your C65</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-edge-40-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Edge 40</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-g42-5g-by-fonelab-android-recover-messages/"><u>How to recover old messages from your G42 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-galaxy-z-flip-5-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Galaxy Z Flip 5</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-galaxy-z-fold-5-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Galaxy Z Fold 5</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-gionee-f3-pro-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Gionee F3 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-google-pixel-8-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Google Pixel 8</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-google-pixel-8-pro-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Google Pixel 8 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-google-pixel-fold-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Google Pixel Fold</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-gt-10-pro-by-fonelab-android-recover-messages/"><u>How to recover old messages from your GT 10 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-honor-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Honor</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-honor-100-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Honor 100</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-honor-70-lite-5g-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Honor 70 Lite 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-honor-90-lite-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Honor 90 Lite</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-honor-90-pro-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Honor 90 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-honor-magic-5-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Honor Magic 5</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-honor-magic5-ultimate-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Honor Magic5 Ultimate</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-honor-play-7t-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Honor Play 7T</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-honor-play-8t-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Honor Play 8T</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-honor-x50-gt-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Honor X50 GT</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-honor-x50i-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Honor X50i</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-honor-x8b-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Honor X8b</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-htc-u23-by-fonelab-android-recover-messages/"><u>How to recover old messages from your HTC U23</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-htc-u23-pro-by-fonelab-android-recover-messages/"><u>How to recover old messages from your HTC U23 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-huawei-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Huawei</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-huawei-nova-y91-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Huawei Nova Y91</u></a></li>
<li><a href="https://apple-account.techidaily.com/how-to-reset-the-security-questions-of-your-apple-id-from-your-apple-iphone-6s-by-drfone-ios/"><u>How To Reset the Security Questions of Your Apple ID From Your Apple iPhone 6s</u></a></li>
<li><a href="https://youtube-tips.techidaily.com/24-crafting-clearer-communication-the-art-of-adding-text-to-video-media/"><u>In 2024, Crafting Clearer Communication  The Art of Adding Text to Video Media</u></a></li>
<li><a href="https://smart-video-editing.techidaily.com/new-the-best-of-the-big-screen-top-movie-trailer-apps-for-iphone-and-ipad-users-for-2024/"><u>New The Best of the Big Screen Top Movie Trailer Apps for iPhone and iPad Users for 2024</u></a></li>
<li><a href="https://voice-adjusting.techidaily.com/new-the-leading-lineup-best-no-charges-music-transformation-apps-for-iphoneandroid-enthusiasts-for-2024/"><u>New The Leading Lineup Best No-Charges Music Transformation Apps for iPhone/Android Enthusiasts for 2024</u></a></li>
<li><a href="https://graphic-issues.techidaily.com/sharpening-windows-view-on-screen/"><u>Sharpening Windows View on Screen</u></a></li>
<li><a href="https://unlock-android.techidaily.com/unlocking-made-easy-the-best-10-apps-for-unlocking-your-zte-device-by-drfone-android/"><u>Unlocking Made Easy The Best 10 Apps for Unlocking Your ZTE Device</u></a></li>
<li><a href="https://howto.techidaily.com/vivo-y27-5g-stuck-on-screen-finding-solutions-for-stuck-on-boot-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>Vivo Y27 5G Stuck on Screen – Finding Solutions For Stuck on Boot | Dr.fone</u></a></li>
</ul></div>

---
title: How to fix runtime error 424 object required error in Excel 2013
date: 2024-07-17T09:59:54.449Z
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
<li><a href="https://screen-video-capture.techidaily.com/new-2024-approved-11-best-practices-for-capturing-online-conferences-for-free/"><u>[New] 2024 Approved  11 Best Practices for Capturing Online Conferences for Free</u></a></li>
<li><a href="https://tiktok-videos.techidaily.com/new-2024-approved-a-primer-to-navigating-the-world-of-tiktok-on-computers/"><u>[New] 2024 Approved  A Primer to Navigating the World of TikTok on Computers</u></a></li>
<li><a href="https://facebook-record-videos.techidaily.com/new-2024-approved-download-mastery-11-key-youtube-extractors/"><u>[New] 2024 Approved  Download Mastery  11 Key YouTube Extractors</u></a></li>
<li><a href="https://facebook-video-recording.techidaily.com/new-correcting-accidental-youself-reappearances-on-social-media-for-2024/"><u>[New] Correcting Accidental 'Youself' Reappearances on Social Media for 2024</u></a></li>
<li><a href="https://vp-tips.techidaily.com/new-exclusive-list-digital-dominators-in-av-production-for-2024/"><u>[New] Exclusive List  Digital Dominators in AV Production for 2024</u></a></li>
<li><a href="https://screen-sharing-recording.techidaily.com/new-quick-shift-video-techniques-mastering-transitions/"><u>[New] Quick-Shift Video Techniques  Mastering Transitions</u></a></li>
<li><a href="https://discord-videos.techidaily.com/updated-best-voice-transformation-plugins-top-10-picks-for-discord-for-2024/"><u>[Updated] Best Voice Transformation Plugins - Top 10 Picks for Discord for 2024</u></a></li>
<li><a href="https://facebook-record-videos.techidaily.com/updated-exploring-streaming-options-vimeo-vs-youtubes-popularity-for-2024/"><u>[Updated] Exploring Streaming Options  Vimeo vs YouTube's Popularity for 2024</u></a></li>
<li><a href="https://on-screen-recording.techidaily.com/updated-in-2024-essential-tools-for-effective-adobe-presentation-recordings/"><u>[Updated] In 2024, Essential Tools for Effective Adobe Presentation Recordings</u></a></li>
<li><a href="https://eaxpv-info.techidaily.com/updated-in-2024-how-to-add-timestamps-to-youtube-video-to-increase-views/"><u>[Updated] In 2024, How to Add Timestamps to YouTube Video to Increase Views</u></a></li>
<li><a href="https://tiktok-clips.techidaily.com/updated-in-2024-maximizing-tiktok-promotion-strategies-and-top-ad-success-stories/"><u>[Updated] In 2024, Maximizing TikTok Promotion  Strategies & Top Ad Success Stories</u></a></li>
<li><a href="https://some-tips.techidaily.com/updated-the-ultimate-list-of-6-best-nft-platforms-for-artists/"><u>[Updated] The Ultimate List of 6 Best NFT Platforms for Artists</u></a></li>
<li><a href="https://extra-resources.techidaily.com/2024-approved-top-10-collage-apps-for-android-top-picks/"><u>2024 Approved  Top 10 Collage Apps for Android (Top Picks)</u></a></li>
<li><a href="https://location-fake.techidaily.com/a-detailed-vpna-fake-gps-location-free-review-on-realme-11-5g-drfone-by-drfone-virtual-android/"><u>A Detailed VPNa Fake GPS Location Free Review On Realme 11 5G | Dr.fone</u></a></li>
<li><a href="https://mondly-stories.techidaily.com/choosing-between-german-and-french-for-global-communication/"><u>Choosing Between German & French for Global Communication</u></a></li>
<li><a href="https://activate-lock.techidaily.com/easy-fixes-how-to-recover-forgotten-icloud-password-from-your-iphone-12-pro-max-by-drfone-ios/"><u>Easy Fixes How To Recover Forgotten iCloud Password From your iPhone 12 Pro Max</u></a></li>
<li><a href="https://screen-recording.techidaily.com/game-recording-mastery-utilizing-gameye-and-more/"><u>Game Recording Mastery  Utilizing GamEye and More</u></a></li>
<li><a href="https://apple-account.techidaily.com/how-to-fix-when-apple-account-locked-from-apple-iphone-7-plus-by-drfone-ios/"><u>How to Fix when Apple Account Locked From Apple iPhone 7 Plus?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-tecno-spark-go-2024-by-fonelab-android-recover-data/"><u>How to recover lost data from Tecno Spark Go (2024)?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-v-purse-by-fonelab-android-recover-data/"><u>How to recover lost data from V Purse?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-v30-by-fonelab-android-recover-data/"><u>How to recover lost data from V30?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-vivo-by-fonelab-android-recover-data/"><u>How to recover lost data from Vivo ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-vivo-t2-5g-by-fonelab-android-recover-data/"><u>How to recover lost data from Vivo T2 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-vivo-v29-pro-by-fonelab-android-recover-data/"><u>How to recover lost data from Vivo V29 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-vivo-v29e-by-fonelab-android-recover-data/"><u>How to recover lost data from Vivo V29e?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-vivo-x100-pro-by-fonelab-android-recover-data/"><u>How to recover lost data from Vivo X100 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-vivo-x100-by-fonelab-android-recover-data/"><u>How to recover lost data from Vivo X100?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-vivo-x90s-by-fonelab-android-recover-data/"><u>How to recover lost data from Vivo X90S?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-vivo-y200-by-fonelab-android-recover-data/"><u>How to recover lost data from Vivo Y200?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-vivo-y27-4g-by-fonelab-android-recover-data/"><u>How to recover lost data from Vivo Y27 4G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-vivo-y78-5g-by-fonelab-android-recover-data/"><u>How to recover lost data from Vivo Y78 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-xiaomi-by-fonelab-android-recover-data/"><u>How to recover lost data from Xiaomi ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-xiaomi-13t-by-fonelab-android-recover-data/"><u>How to recover lost data from Xiaomi 13T?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-xiaomi-mix-fold-3-by-fonelab-android-recover-data/"><u>How to recover lost data from Xiaomi Mix Fold 3?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-xiaomi-redmi-13c-5g-by-fonelab-android-recover-data/"><u>How to recover lost data from Xiaomi Redmi 13C 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-xiaomi-redmi-note-12r-by-fonelab-android-recover-data/"><u>How to recover lost data from Xiaomi Redmi Note 12R?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-y02t-by-fonelab-android-recover-data/"><u>How to recover lost data from Y02T?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-y36-by-fonelab-android-recover-data/"><u>How to recover lost data from Y36?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-yuva-2-by-fonelab-android-recover-data/"><u>How to recover lost data from Yuva 2?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-zte-axon-40-lite-by-fonelab-android-recover-data/"><u>How to recover lost data from ZTE Axon 40 Lite?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-zte-blade-a73-5g-by-fonelab-android-recover-data/"><u>How to recover lost data from ZTE Blade A73 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-of-iphone-11-pro-max-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data of iPhone 11 Pro Max? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-of-iphone-11-pro-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data of iPhone 11 Pro? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-of-iphone-11-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data of iPhone 11? | Dr.fone</u></a></li>
<li><a href="https://ai-editing-video.techidaily.com/in-2024-are-you-looking-to-finish-your-youtube-video-in-style-the-following-are-the-best-end-screen-makers-to-choose-from/"><u>In 2024, Are You Looking to Finish Your YouTube Video in Style? The Following Are the Best End Screen Makers to Choose From</u></a></li>
<li><a href="https://youtube-data.techidaily.com/24-comparing-free-and-premium-youtube-experiences-whats-best/"><u>In 2024, Comparing Free and Premium YouTube Experiences  What's Best?</u></a></li>
<li><a href="https://video-capture.techidaily.com/in-2024-download-screen-recorder-pro-for-windows-10/"><u>In 2024, Download Screen Recorder Pro for Windows 10</u></a></li>
<li><a href="https://smart-video-creator.techidaily.com/in-2024-edit-like-a-pro-top-iphone-video-editing-apps/"><u>In 2024, Edit Like a Pro Top iPhone Video Editing Apps</u></a></li>
<li><a href="https://change-location.techidaily.com/in-2024-how-to-get-the-dragon-scale-and-evolution-enabled-pokemon-on-samsung-galaxy-s23-tactical-edition-drfone-by-drfone-virtual-android/"><u>In 2024, How to get the dragon scale and evolution-enabled pokemon On Samsung Galaxy S23 Tactical Edition? | Dr.fone</u></a></li>
<li><a href="https://animation-videos.techidaily.com/in-2024-whats-rigging-animation-and-why-it-matters/"><u>In 2024, Whats Rigging Animation and Why It Matters</u></a></li>
<li><a href="https://animation-videos.techidaily.com/new-top-10-best-photo-to-animation-converters-that-give-you-an-advantage-for-2024/"><u>New Top 10 Best Photo to Animation Converters That Give You an Advantage for 2024</u></a></li>
<li><a href="https://youtube-clips.techidaily.com/youtubers-toolkit-essential-setup-guidelines/"><u>Youtubers' Toolkit  Essential Setup Guidelines</u></a></li>
</ul></div>

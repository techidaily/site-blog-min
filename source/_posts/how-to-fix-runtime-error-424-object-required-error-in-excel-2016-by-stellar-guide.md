---
title: How to fix runtime error 424 object required error in Excel 2016
date: 2024-07-17T09:59:57.005Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to fix runtime error 424 object required error in Excel 2016
keywords: repair excel 2021,repair .xls,repair excel 2019,repair excel 2003,repair .xlb
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
<a href="https://zonlipartnershipprogram.pxf.io/c/5597632/1596691/17882" target="_top" id="1596691"><img src="//a.impactradius-go.com/display-ad/17882-1596691" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1596691/17882" style="position:absolute;visibility:hidden;" border="0" />
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

<!-- affiliate ads begin -->
<a href="https://arkmc.pxf.io/c/5597632/427477/5172" target="_top" id="427477"><img src="//a.impactradius-go.com/display-ad/5172-427477" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://arkmc.pxf.io/i/5597632/427477/5172" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
##  **Solutions to Fix Runtime Error 424: Object Required in Excel**

The VBA error ‘object required’ may occur due to different reasons. Based on the reason, you can follow the solutions mentioned below to fix the error.

### **1\. Check the Name of the Object**

The Runtime error 424 can occur when you run the VBA code using an incorrect name of the object. For example, the object name is ‘MyObject’ but you’re using “Backcolor”.

<!-- affiliate ads begin -->
<a href="https://printrendy.pxf.io/c/5597632/1453719/17020" target="_top" id="1453719"><img src="//a.impactradius-go.com/display-ad/17020-1453719" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1453719/17020" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Error When Incorrect Name Of The Object](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/error-424-when-incorrect-name-of-the-object.jpg)

 When you click the **Debug** button, the line with the error will highlight.

<!-- affiliate ads begin -->
<a href="https://store.iobit.com/order/checkout.php?PRODS=1468905&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/184260348236f9554fe9375772ff966e/ascscan_728x90.png" border="0"></a>
<!-- affiliate ads end -->
![Highlighting Line With Error](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/highlighting-line-with-error.jpg)

To fix the issue, you need to provide the correct name of the object.

<!-- affiliate ads begin -->
<a href="https://shop.mondly.com/affiliate.php?ACCOUNT=ATISTUDI&AFFILIATE=108875&PATH=https%3A%2F%2Fwww.mondly.com%3FAFFILIATE%3D108875%26RESOURCE%3D%2BEducational%2B970x90%2B"><img src="https://secure.avangate.com/images/merchant/69c418c33ec2e1a4267fa9bb77fa1428/educational-970x90.gif" border="0"></a>
<!-- affiliate ads end -->
### **2\. Check if the Object is Missing**

 The Runtime error 424 can occur if the object you are referring to as a method is not available or you are using the wrong object in a code. In the below example, you can see that the error occurs when an object named “Employee” is not available in the Project list.

<!-- affiliate ads begin -->
<a href="https://store.revouninstaller.com/order/checkout.php?PRODS=27889512&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/4282ec8de8c9be897e7aff4aa231b1a4/728__90.jpg" border="0"></a>
<!-- affiliate ads end -->
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

<!-- affiliate ads begin -->
<a href="https://ursime.pxf.io/c/5597632/2092236/16384" target="_top" id="2092236"><img src="//a.impactradius-go.com/display-ad/16384-2092236" border="0" alt="" width="1920" height="329"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2092236/16384" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
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
<li><a href="https://youtube-web.techidaily.com/024-approved-ace-your-edits-youtube-studios-time-saving-secrets-revealed/"><u>[New] 2024 Approved  Ace Your Edits  YouTube Studio's Time-Saving Secrets Revealed</u></a></li>
<li><a href="https://fox-hovers.techidaily.com/new-breaking-down-the-spectacular-features-of-dells-p2715q-display/"><u>[New] Breaking Down the Spectacular Features of Dell's P2715Q Display</u></a></li>
<li><a href="https://tiktok-video-files.techidaily.com/new-in-2024-8-best-free-tiktok-video-editing-appswebsite-on-mac/"><u>[New] In 2024, 8 Best Free TikTok Video Editing Apps/Website on Mac</u></a></li>
<li><a href="https://tiktok-videos.techidaily.com/new-spectacular-stars-of-tiktok-shaping-tomorrows-dreams/"><u>[New] Spectacular Stars of TikTok Shaping Tomorrow's Dreams</u></a></li>
<li><a href="https://extra-information.techidaily.com/updated-crafting-visual-stories-a-detailed-magix-video-pro-x-review/"><u>[Updated] Crafting Visual Stories  A Detailed Magix Video Pro X Review</u></a></li>
<li><a href="https://facebook-videos.techidaily.com/updated-transform-hidden-facebook-videos-with-these-12-must-try-steps/"><u>[Updated] Transform Hidden Facebook Videos with These 12 Must-Try Steps</u></a></li>
<li><a href="https://fox-http.techidaily.com/2024-approved-echoing-emptiness-guiding-graduated-gone-sounds-in-pp/"><u>2024 Approved  Echoing Emptiness  Guiding Graduated Gone Sounds in PP</u></a></li>
<li><a href="https://youtube-help.techidaily.com/2024-approved-monthly-earnings-from-youtube-as-a-content-creator/"><u>2024 Approved  Monthly Earnings From YouTube as a Content Creator</u></a></li>
<li><a href="https://extra-support.techidaily.com/2024-approved-premium-smartphones-shining-in-high-definition-video/"><u>2024 Approved  Premium Smartphones Shining in High Definition Video</u></a></li>
<li><a href="https://extra-hints.techidaily.com/advanced-computers-at-your-fingertips/"><u>Advanced Computers at Your Fingertips</u></a></li>
<li><a href="https://extra-resources.techidaily.com/boost-podcast-visibility-through-strategic-seo-tactics/"><u>Boost Podcast Visibility Through Strategic SEO Tactics</u></a></li>
<li><a href="https://some-techniques.techidaily.com/fusing-images-online-and-desktop-harmoniously-for-2024/"><u>Fusing Images Online & Desktop Harmoniously for 2024</u></a></li>
<li><a href="https://extra-resources.techidaily.com/guide-to-the-best-17-image-editing-utilities/"><u>Guide to the Best 17 Image Editing Utilities</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/how-to-mirror-honor-play-8t-to-mac-drfone-by-drfone-android/"><u>How to Mirror Honor Play 8T to Mac? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-of-iphone-se-2022-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data of iPhone SE (2022)? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-of-iphone-se-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data of iPhone SE? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-of-iphone-x-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data of iPhone X? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-of-iphone-xr-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data of iPhone XR? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-of-iphone-xs-max-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data of iPhone XS Max? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-of-iphone-xs-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data of iPhone XS? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-11-pro-max-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 11 Pro Max? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-11-pro-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 11 Pro? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-11-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 11? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-12-mini-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 12 mini? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-12-pro-max-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 12 Pro Max? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-12-pro-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 12 Pro? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-12-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 12? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-13-mini-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 13 mini? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-13-pro-max-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 13 Pro Max? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-13-pro-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 13 Pro? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-13-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 13? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-14-plus-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 14 Plus? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-14-pro-max-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 14 Pro Max? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-14-pro-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 14 Pro? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-14-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 14? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-15-plus-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 15 Plus? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-15-pro-max-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 15 Pro Max? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-15-pro-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 15 Pro? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-15-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 15? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-6-plus-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 6 Plus? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-6-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 6? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-6s-plus-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 6s Plus? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-6s-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 6s? | Dr.fone</u></a></li>
<li><a href="https://sim-unlock.techidaily.com/in-2024-all-about-apple-iphone-15-pro-max-unlock-chip-you-need-to-know-by-drfone-ios/"><u>In 2024, All About Apple iPhone 15 Pro Max Unlock Chip You Need to Know</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/in-2024-latest-way-to-get-shiny-meltan-box-in-pokemon-go-mystery-box-on-motorola-g24-power-drfone-by-drfone-virtual-android/"><u>In 2024, Latest way to get Shiny Meltan Box in Pokémon Go Mystery Box On Motorola G24 Power | Dr.fone</u></a></li>
<li><a href="https://some-approaches.techidaily.com/leading-websites-to-discover-glossy-gold-letter-effects-for-2024/"><u>Leading Websites to Discover Glossy Gold Letter Effects for 2024</u></a></li>
<li><a href="https://extra-tips.techidaily.com/master-money-with-periscope-a-newcomers-manual/"><u>Master Money with Periscope  A Newcomer's Manual</u></a></li>
<li><a href="https://youtube-videos.techidaily.com/navigating-new-waters-personalbusiness-youtube-channels-made-mobile-ready/"><u>Navigating New Waters  Personal/Business YouTube Channels Made Mobile-Ready</u></a></li>
<li><a href="https://audio-shaping.techidaily.com/new-excellent-free-online-audio-calibration-alternatives/"><u>New Excellent Free Online Audio Calibration Alternatives</u></a></li>
<li><a href="https://review-topics.techidaily.com/recover-lost-data-from-infinix-by-fonelab-android-recover-data/"><u>Recover lost data from Infinix</u></a></li>
<li><a href="https://ai-vdieo-software.techidaily.com/run-lumafusion-on-mac-download-and-alternative-options/"><u>Run Lumafusion on Mac Download and Alternative Options</u></a></li>
<li><a href="https://some-tips.techidaily.com/supreme-reconciliation-of-vr-realms-for-2024/"><u>Supreme Reconciliation of VR Realms for 2024</u></a></li>
<li><a href="https://driver-install.techidaily.com/swiftly-implement-hd530-software/"><u>Swiftly Implement HD530 Software</u></a></li>
<li><a href="https://youtube-videos.techidaily.com/1716464933934-unlock-creative-potential-access-free-sfx/"><u>Unlock Creative Potential, Access Free SFX!</u></a></li>
</ul></div>

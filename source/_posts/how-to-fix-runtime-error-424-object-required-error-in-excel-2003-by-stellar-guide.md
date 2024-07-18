---
title: How to fix runtime error 424 object required error in Excel 2003
date: 2024-07-17T09:59:46.877Z
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

<!-- affiliate ads begin -->
<a href="https://newchic.sjv.io/c/5597632/1659704/14420" target="_top" id="1659704"><img src="//a.impactradius-go.com/display-ad/14420-1659704" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1659704/14420" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Error When Incorrect Name Of The Object](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/error-424-when-incorrect-name-of-the-object.jpg)

 When you click the **Debug** button, the line with the error will highlight.

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2082535/7443" target="_top" id="2082535"><img src="//a.impactradius-go.com/display-ad/7443-2082535" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2082535/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Highlighting Line With Error](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/highlighting-line-with-error.jpg)

To fix the issue, you need to provide the correct name of the object.

### **2\. Check if the Object is Missing**

 The Runtime error 424 can occur if the object you are referring to as a method is not available or you are using the wrong object in a code. In the below example, you can see that the error occurs when an object named “Employee” is not available in the Project list.

![Example Of Code When Object Is Not Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/example-of-code-when-an-object-is-not-available.jpg)

 You can check and mention the object which is available. For instance, Sheet2 in the below code.  

<!-- affiliate ads begin -->
<a href="https://printrendy.pxf.io/c/5597632/1453721/17020" target="_top" id="1453721"><img src="//a.impactradius-go.com/display-ad/17020-1453721" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1453721/17020" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Check When The Object Is Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/check-when-the-object-is-available.png)

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2075461/7443" target="_top" id="2075461"><img src="//a.impactradius-go.com/display-ad/7443-2075461" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2075461/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **3\. Check All References are Declared in the Code**

You can get the Runtime error 424 if all the references are not declared. So, make sure you have declared all the references in the code. To verify this, you can use the debug mode by pressing **F5** or clicking on the **Debug** option.

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2068416/7443" target="_top" id="2068416"><img src="//a.impactradius-go.com/display-ad/7443-2068416" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2068416/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Debug Command In Excel](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/debug-command-in-excel.jpg)

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2087407/7443" target="_top" id="2087407"><img src="//a.impactradius-go.com/display-ad/7443-2087407" border="0" alt="" width="600" height="500"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2087407/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **4\. Check the Macro Security Settings**

Sometimes, the error can occur if macros are disabled in the Macro Security settings. You can check and change the settings by following these steps:

- On the **Developer** tab, in the **Code** section, click **Macro Security**.
- In the **Trust Center** window, select **Enable all macros.**

![Macro Security Wizard](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/macro-security-wizard.jpg)

- Click **OK**.

![Enable All Macro In Trust Center](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/enable-all-macros-in-trust-center.jpg)

<!-- affiliate ads begin -->
<a href="https://shop.mondly.com/affiliate.php?ACCOUNT=ATISTUDI&AFFILIATE=108875&PATH=https%3A%2F%2Fwww.mondly.com%3FAFFILIATE%3D108875%26RESOURCE%3D%2BBusiness%2B970x90%2B"><img src="https://secure.avangate.com/images/merchant/69c418c33ec2e1a4267fa9bb77fa1428/business-970x90.gif" border="0"></a>
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
<li><a href="https://fox-blue.techidaily.com/new-2024-approved-unlocking-the-potential-of-snapchat-highlights/"><u>[New] 2024 Approved  Unlocking the Potential of Snapchat Highlights</u></a></li>
<li><a href="https://instagram-video-recordings.techidaily.com/new-curated-igtv-channels-that-matter-today/"><u>[New] Curated IGTV Channels That Matter Today</u></a></li>
<li><a href="https://screen-video-capture.techidaily.com/new-in-2024-elevate-your-video-remove-extraneous-details-effectively/"><u>[New] In 2024, Elevate Your Video - Remove Extraneous Details Effectively</u></a></li>
<li><a href="https://snapchat-videos.techidaily.com/new-in-2024-mobile-mastery-securing-your-snapchat-videos-on-phone/"><u>[New] In 2024, Mobile Mastery  Securing Your Snapchat Videos on Phone</u></a></li>
<li><a href="https://desktop-recording.techidaily.com/new-in-2024-ultimate-guide-to-documenting-your-ps4-journey/"><u>[New] In 2024, Ultimate Guide to Documenting Your PS4 Journey</u></a></li>
<li><a href="https://youtube-docs.techidaily.com/ptimal-tag-selection-elevate-your-videos-online-presence-for-2024/"><u>[New] Optimal Tag Selection  Elevate Your Video's Online Presence for 2024</u></a></li>
<li><a href="https://facebook-record-videos.techidaily.com/new-the-dark-side-of-digital-content-profitability/"><u>[New] The Dark Side of Digital Content Profitability</u></a></li>
<li><a href="https://twitter-videos.techidaily.com/new-tiktoks-viral-video-wave-twitters-top-10-rankings/"><u>[New] TikTok's Viral Video Wave  Twitter's Top 10 Rankings</u></a></li>
<li><a href="https://facebook-record-videos.techidaily.com/updated-guide-on-transforming-youtube-videos-into-mp3-using-safe-procedures/"><u>[Updated] Guide on Transforming YouTube Videos Into MP3 Using Safe Procedures</u></a></li>
<li><a href="https://instagram-videos.techidaily.com/updated-how-to-convert-instagram-video-to-mp4-2-proven-ways-for-2024/"><u>[Updated] How to Convert Instagram Video to MP4 [ 2 Proven Ways] for 2024</u></a></li>
<li><a href="https://some-guidance.techidaily.com/updated-top-choices-for-next-gen-vr-accessories-unveiled/"><u>[Updated] Top Choices for Next-Gen VR Accessories Unveiled</u></a></li>
<li><a href="https://twitter-videos.techidaily.com/2024-approved-leveraging-live-tweets-strategies-unveiled/"><u>2024 Approved  Leveraging Live Tweets  Strategies Unveiled</u></a></li>
<li><a href="https://fox-info.techidaily.com/2024-approved-stellarworkspace-unified-elite-full-hd-screens/"><u>2024 Approved  StellarWorkspace  Unified, Elite Full HD Screens</u></a></li>
<li><a href="https://tiktok-videos.techidaily.com/drafting-a-compelling-closing-statement-for-tiktok/"><u>Drafting a Compelling Closing Statement for TikTok</u></a></li>
<li><a href="https://extra-tips.techidaily.com/exploring-top-notch-gif-utilities-on-ios-devices/"><u>Exploring Top-Notch GIF Utilities on iOS Devices</u></a></li>
<li><a href="https://bypass-frp.techidaily.com/how-can-we-bypass-infinix-note-30-pro-frp-by-drfone-android/"><u>How Can We Bypass Infinix Note 30 Pro FRP?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-huawei-nova-y71-by-fonelab-android-recover-data/"><u>How to recover lost data from Huawei Nova Y71?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-huawei-nova-y91-by-fonelab-android-recover-data/"><u>How to recover lost data from Huawei Nova Y91?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-huawei-p60-by-fonelab-android-recover-data/"><u>How to recover lost data from Huawei P60?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-infinix-by-fonelab-android-recover-data/"><u>How to recover lost data from Infinix ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-infinix-gt-10-pro-by-fonelab-android-recover-data/"><u>How to recover lost data from Infinix GT 10 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-infinix-hot-30i-by-fonelab-android-recover-data/"><u>How to recover lost data from Infinix Hot 30i?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-infinix-hot-40-by-fonelab-android-recover-data/"><u>How to recover lost data from Infinix Hot 40?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-infinix-note-30-5g-by-fonelab-android-recover-data/"><u>How to recover lost data from Infinix Note 30 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-infinix-note-30-pro-by-fonelab-android-recover-data/"><u>How to recover lost data from Infinix Note 30 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-infinix-note-30i-by-fonelab-android-recover-data/"><u>How to recover lost data from Infinix Note 30i?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-infinix-smart-7-hd-by-fonelab-android-recover-data/"><u>How to recover lost data from Infinix Smart 7 HD?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-infinix-smart-8-by-fonelab-android-recover-data/"><u>How to recover lost data from Infinix Smart 8?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-infinix-zero-30-5g-by-fonelab-android-recover-data/"><u>How to recover lost data from Infinix Zero 30 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-iphone-11-after-ios-update-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Lost Data from iPhone 11 After iOS Update? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-iphone-11-pro-after-ios-update-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Lost Data from iPhone 11 Pro After iOS Update? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-iphone-11-pro-max-after-ios-update-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Lost Data from iPhone 11 Pro Max After iOS Update? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-iphone-11-pro-max-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data from iPhone 11 Pro Max? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-iphone-11-pro-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data from iPhone 11 Pro? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-iphone-11-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data from iPhone 11? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-iphone-12-after-ios-update-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Lost Data from iPhone 12 After iOS Update? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-iphone-12-mini-after-ios-update-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Lost Data from iPhone 12 mini After iOS Update? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-iphone-12-mini-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data from iPhone 12 mini? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-iphone-12-pro-after-ios-update-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Lost Data from iPhone 12 Pro After iOS Update? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-iphone-12-pro-max-after-ios-update-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Lost Data from iPhone 12 Pro Max After iOS Update? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-iphone-12-pro-max-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data from iPhone 12 Pro Max? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-iphone-12-pro-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data from iPhone 12 Pro? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-iphone-12-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data from iPhone 12? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-iphone-13-after-ios-update-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Lost Data from iPhone 13 After iOS Update? | Stellar</u></a></li>
<li><a href="https://android-unlock.techidaily.com/how-to-unlock-any-samsung-galaxy-s23-phone-password-using-emergency-call-by-drfone-android/"><u>How To Unlock Any Samsung Galaxy S23 Phone Password Using Emergency Call</u></a></li>
<li><a href="https://visual-screen-recording.techidaily.com/in-2024-how-to-make-a-ken-burns-effect-in-camtasia-9/"><u>In 2024, How to Make a Ken Burns Effect in Camtasia 9?</u></a></li>
<li><a href="https://extra-support.techidaily.com/in-2024-navigating-through-gopros-burst-recording-options/"><u>In 2024, Navigating Through GoPro's Burst Recording Options</u></a></li>
<li><a href="https://youtube-blog.techidaily.com/ators-designing-online-marvel-experiences/"><u>Innovators Designing Online Marvel Experiences</u></a></li>
<li><a href="https://fake-location.techidaily.com/life360-learn-how-everything-works-on-samsung-galaxy-xcover-6-pro-tactical-edition-drfone-by-drfone-virtual-android/"><u>Life360 Learn How Everything Works On Samsung Galaxy XCover 6 Pro Tactical Edition | Dr.fone</u></a></li>
<li><a href="https://audio-shaping.techidaily.com/new-echoes-of-creativity-where-to-access-premium-audio-for-episodic-storytelling/"><u>New Echoes of Creativity Where to Access Premium Audio for Episodic Storytelling</u></a></li>
<li><a href="https://audio-shaping.techidaily.com/new-in-2024-mute-media-magic-a-how-to-guide-for-quieting-soundtracks-in-mov-files-across-operating-systems/"><u>New In 2024, Mute Media Magic A How-To Guide for Quieting Soundtracks in MOV Files Across Operating Systems</u></a></li>
<li><a href="https://extra-tips.techidaily.com/synchronized-snaps-visuals-to-the-rhythm-of-instagram/"><u>Synchronized Snaps  Visuals to the Rhythm of Instagram</u></a></li>
<li><a href="https://smart-video-editing.techidaily.com/unleash-your-creativity-top-free-and-paid-3d-animation-tools-for-2024/"><u>Unleash Your Creativity Top Free and Paid 3D Animation Tools for 2024</u></a></li>
<li><a href="https://smart-video-creator.techidaily.com/updated-elevate-your-videos-top-rated-apple-video-editing-programs/"><u>Updated Elevate Your Videos Top-Rated Apple Video Editing Programs</u></a></li>
</ul></div>

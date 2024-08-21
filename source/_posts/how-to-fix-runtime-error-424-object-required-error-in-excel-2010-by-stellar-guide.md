---
title: How to fix runtime error 424 object required error in Excel 2010
date: 2024-08-20T20:48:09.095Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to fix runtime error 424 object required error in Excel 2010
keywords: repair excel 2023,repair .xlsm,repair excel 2013,repair .xlb,repair .xltx,repair excel 2019
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
<a href="https://coinrule.sjv.io/c/5597632/1958374/18409" target="_top" id="1958374"><img src="//a.impactradius-go.com/display-ad/18409-1958374" border="0" alt="" width="300" height="300"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1958374/18409" style="position:absolute;visibility:hidden;" border="0" />
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
<a href="https://ukaidot.sjv.io/c/5597632/1793237/19578" target="_top" id="1793237"><img src="//a.impactradius-go.com/display-ad/19578-1793237" border="0" alt="" width="1200" height="1200"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1793237/19578" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **1\. Check the Name of the Object**

The Runtime error 424 can occur when you run the VBA code using an incorrect name of the object. For example, the object name is ‘MyObject’ but you’re using “Backcolor”.

![Error When Incorrect Name Of The Object](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/error-424-when-incorrect-name-of-the-object.jpg)

 When you click the **Debug** button, the line with the error will highlight.

![Highlighting Line With Error](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/highlighting-line-with-error.jpg)

To fix the issue, you need to provide the correct name of the object.

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2068411/7443" target="_top" id="2068411"><img src="//a.impactradius-go.com/display-ad/7443-2068411" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2068411/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **2\. Check if the Object is Missing**

 The Runtime error 424 can occur if the object you are referring to as a method is not available or you are using the wrong object in a code. In the below example, you can see that the error occurs when an object named “Employee” is not available in the Project list.

![Example Of Code When Object Is Not Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/example-of-code-when-an-object-is-not-available.jpg)

 You can check and mention the object which is available. For instance, Sheet2 in the below code.  

<!-- affiliate ads begin -->
<a href="https://twopages.pxf.io/c/5597632/1873313/18544" target="_top" id="1873313"><img src="//a.impactradius-go.com/display-ad/18544-1873313" border="0" alt="" width="1080" height="1263"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1873313/18544" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Check When The Object Is Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/check-when-the-object-is-available.png)

### **3\. Check All References are Declared in the Code**

You can get the Runtime error 424 if all the references are not declared. So, make sure you have declared all the references in the code. To verify this, you can use the debug mode by pressing **F5** or clicking on the **Debug** option.

![Debug Command In Excel](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/debug-command-in-excel.jpg)

<!-- affiliate ads begin -->
<a href="https://zonlipartnershipprogram.pxf.io/c/5597632/1611407/17882" target="_top" id="1611407"><img src="//a.impactradius-go.com/display-ad/17882-1611407" border="0" alt="" width="300" height="485"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1611407/17882" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **4\. Check the Macro Security Settings**

Sometimes, the error can occur if macros are disabled in the Macro Security settings. You can check and change the settings by following these steps:

- On the **Developer** tab, in the **Code** section, click **Macro Security**.
- In the **Trust Center** window, select **Enable all macros.**

<!-- affiliate ads begin -->
<a href="https://twopages.pxf.io/c/5597632/2016067/18544" target="_top" id="2016067"><img src="//a.impactradius-go.com/display-ad/18544-2016067" border="0" alt="" width="1020" height="380"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2016067/18544" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Macro Security Wizard](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/macro-security-wizard.jpg)

- Click **OK**.

<!-- affiliate ads begin -->
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=3851691&QTY=1&AFFILIATE=108875&CART=1"><img src="http://www.aiseesoft.com/avangate/30p/banner.jpg" border="0"></a>
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
<li><a href="https://vp-tips.techidaily.com/new-2024-approved-13-pivotal-metaverse-experiences-for-deeper-comprehension/"><u>[New] 2024 Approved  13 Pivotal Metaverse Experiences for Deeper Comprehension</u></a></li>
<li><a href="https://digital-screen-recording.techidaily.com/new-deciphering-the-advantages-of-modern-multicam-technology/"><u>[New] Deciphering the Advantages of Modern Multicam Technology</u></a></li>
<li><a href="https://eaxpv-info.techidaily.com/new-in-2024-high-ranking-videos-youtubes-everlasting-classics/"><u>[New] In 2024, High-Ranking Videos  YouTube’s Everlasting Classics</u></a></li>
<li><a href="https://some-knowledge.techidaily.com/updated-examining-streaming-platform-powerhouses-twitch-and-youtube-comparison/"><u>[Updated] Examining Streaming Platform Powerhouses  Twitch and YouTube Comparison</u></a></li>
<li><a href="https://visual-screen-recording.techidaily.com/updated-unveiling-the-essence-of-top-tier-screen-recorders-with-obs/"><u>[Updated] Unveiling the Essence of Top-Tier Screen Recorders with OBS</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-honor-magic-v2-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Honor Magic V2 to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-infinix-gt-10-pro-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Infinix GT 10 Pro to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-motorola-razr-40-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Motorola Razr 40 to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-motorola-moto-g34-5g-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Motorola Moto G34 5G to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-techniques-to-transfer-data-from-samsung-galaxy-a15-5g-to-iphone-15141312-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Techniques to Transfer Data from Samsung Galaxy A15 5G to iPhone 15/14/13/12 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-ways-to-move-contacts-from-tecno-pova-6-pro-5g-to-iphone-131415-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Ways to Move Contacts From Tecno Pova 6 Pro 5G to iPhone (13/14/15) | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-ways-to-move-contacts-from-vivo-s17-pro-to-iphone-131415-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Ways to Move Contacts From Vivo S17 Pro to iPhone (13/14/15) | Dr.fone</u></a></li>
<li><a href="https://on-screen-recording.techidaily.com/assessing-camera-multicam-systems-in-modern-living-spaces/"><u>Assessing Camera Multicam Systems in Modern Living Spaces</u></a></li>
<li><a href="https://fox-access.techidaily.com/frame-flair-choosing-the-best-video-cameras-for-deliberate-moves-for-2024/"><u>Frame Flair  Choosing the Best Video Cameras for Deliberate Moves for 2024</u></a></li>
<li><a href="https://ios-unlock.techidaily.com/how-do-you-remove-restricted-mode-on-iphone-13-by-drfone-ios/"><u>How Do You Remove Restricted Mode on iPhone 13</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-xr-without-data-loss-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone XR without Data Loss? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-g2-get-deleted-pictures-back-with-ease-and-safety-by-fonelab-android-recover-pictures/"><u>How to G2 Get Deleted Pictures Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-out-of-recovery-on-iphone-se-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Get Out of Recovery on iPhone SE? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-app-on-11-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Recover Deleted Photos from Android Gallery App on 11 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-infinix-note-30-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Infinix Note 30</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-pictures-from-tecno-by-fonelab-android-recover-pictures/"><u>How to recover deleted pictures from Tecno .</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-videos-from-your-itel-p55plus-by-fonelab-android-recover-video/"><u>How to recover old videos from your Itel P55+</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-videos-from-your-vivo-y28-5g-by-fonelab-android-recover-video/"><u>How to recover old videos from your Vivo Y28 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-google-frp-lock-on-magic-5-lite-by-drfone-android-unlock-remove-google-frp/"><u>How to remove Google FRP Lock on Magic 5 Lite</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-mdm-from-iphone-12-pro-max-without-a-computer-by-drfone-ios-unlock-ios-unlock/"><u>How to Remove MDM from iPhone 12 Pro Max without a computer?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-realme-narzo-60-5g-using-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Realme Narzo 60 5G using Video Repair Utility on Windows? </u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupted-pdf-v16-file-when-manual-method-fails-by-stellar-guide/"><u>How to repair corrupted PDF v1.6 file when manual method fails</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-contacts-from-infinix-note-30-5g-by-fonelab-android-recover-contacts/"><u>How to Rescue Lost Contacts from Infinix Note 30 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-photos-from-lava-storm-5g-by-fonelab-android-recover-photos/"><u>How to Rescue Lost Photos from Lava Storm 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-realme-gt-5-pro-pictures-an-easy-method-explained-by-fonelab-android-recover-pictures/"><u>How to Restore Deleted Realme GT 5 Pro Pictures  An Easy Method Explained.</u></a></li>
<li><a href="https://ios-unlock.techidaily.com/how-to-unlock-apple-iphone-13-pro-max-without-passcode-or-face-id-by-drfone-ios/"><u>How to Unlock Apple iPhone 13 Pro Max without Passcode or Face ID</u></a></li>
<li><a href="https://some-knowledge.techidaily.com/identifying-ideal-pace-for-premium-slow-motion-cinematography-for-2024/"><u>Identifying Ideal Pace for Premium Slow Motion Cinematography for 2024</u></a></li>
<li><a href="https://extra-information.techidaily.com/in-2024-audience-advocacy-nurturing-connections-in-the-digital-age/"><u>In 2024, Audience Advocacy  Nurturing Connections in the Digital Age</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/in-2024-rootjunky-apk-to-bypass-google-frp-lock-for-realme-narzo-60-5g-by-drfone-android/"><u>In 2024, Rootjunky APK To Bypass Google FRP Lock For Realme Narzo 60 5G</u></a></li>
<li><a href="https://youtube-data.techidaily.com/-live-tech-the-most-popular-streaming-tools-reviewed/"><u>Prime Live Tech  The Most Popular Streaming Tools Reviewed</u></a></li>
<li><a href="https://some-approaches.techidaily.com/strategies-for-utilizing-video-resources-in-learning-for-2024/"><u>Strategies for Utilizing Video Resources in Learning for 2024</u></a></li>
<li><a href="https://howto.techidaily.com/super-easy-ways-to-deal-with-nubia-red-magic-8s-proplus-unresponsive-screen-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>Super Easy Ways To Deal with Nubia Red Magic 8S Pro+ Unresponsive Screen | Dr.fone</u></a></li>
<li><a href="https://buynow-tips.techidaily.com/the-dynamic-duo-a-comprehensive-review-of-motorola-one-hypers-engaging-presence-in-the-mid-range-market/"><u>The Dynamic Duo: A Comprehensive Review of Motorola One Hyper's Engaging Presence in the Mid-Range Market</u></a></li>
<li><a href="https://tech-haven.techidaily.com/the-intersection-of-websites-and-artificial-intelligence/"><u>The Intersection of Websites and Artificial Intelligence</u></a></li>
<li><a href="https://tech-savvy.techidaily.com/the-waiting-game-why-is-chatgpt-4-lagging/"><u>The Waiting Game: Why Is ChatGPT-4 Lagging?</u></a></li>
<li><a href="https://tech-recovery.techidaily.com/troubleshooting-the-absent-msodll-file-in-your-system-effective-solutions/"><u>Troubleshooting the Absent mso.dll File in Your System: Effective Solutions</u></a></li>
<li><a href="https://video-ai-editor.techidaily.com/updated-filmmaking-on-a-shoestring-9-best-budget-friendly-software-options-for-2024/"><u>Updated Filmmaking on a Shoestring 9 Best Budget-Friendly Software Options for 2024</u></a></li>
</ul></div>

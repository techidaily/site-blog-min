---
title: How to fix runtime error 424 object required error in Excel
date: 2024-07-17T10:00:07.555Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to fix runtime error 424 object required error in Excel
keywords: repair .xltx,repair excel 2007,repair .xlb,repair .xltm,repair .csv
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
<a href="https://store.bitdefender.com/affiliate.php?ACCOUNT=BITLATIN&AFFILIATE=108875&PATH=http%3A%2F%2Fwww.bitdefender.com%2Fbusiness%3FAFFILIATE%3D108875%26RESOURCE%3D30%2525%2BOff%2Ball%2BGravityZone%2BProducts"><img src="https://www.bitdefender.com/content/dam/bitdefender/business/campaign/1200X628.png" border="0"></a>
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
<a href="https://appsumo.8odi.net/c/5597632/2068411/7443" target="_top" id="2068411"><img src="//a.impactradius-go.com/display-ad/7443-2068411" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2068411/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Error When Incorrect Name Of The Object](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/error-424-when-incorrect-name-of-the-object.jpg)

 When you click the **Debug** button, the line with the error will highlight.

<!-- affiliate ads begin -->
<a href="https://ursime.pxf.io/c/5597632/2048963/16384" target="_top" id="2048963"><img src="//a.impactradius-go.com/display-ad/16384-2048963" border="0" alt="" width="1200" height="900"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2048963/16384" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Highlighting Line With Error](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/highlighting-line-with-error.jpg)

To fix the issue, you need to provide the correct name of the object.

### **2\. Check if the Object is Missing**

 The Runtime error 424 can occur if the object you are referring to as a method is not available or you are using the wrong object in a code. In the below example, you can see that the error occurs when an object named “Employee” is not available in the Project list.

<!-- affiliate ads begin -->
<a href="https://vapordna.pxf.io/c/5597632/1496243/17238" target="_top" id="1496243"><img src="//a.impactradius-go.com/display-ad/17238-1496243" border="0" alt="" width="1000" height="1221"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1496243/17238" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Example Of Code When Object Is Not Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/example-of-code-when-an-object-is-not-available.jpg)

 You can check and mention the object which is available. For instance, Sheet2 in the below code.  

<!-- affiliate ads begin -->
<a href="https://shop.mondly.com/affiliate.php?ACCOUNT=ATISTUDI&AFFILIATE=108875&PATH=https%3A%2F%2Fwww.mondly.com%3FAFFILIATE%3D108875%26RESOURCE%3D%2BBusiness%2B970x90%2B"><img src="https://secure.avangate.com/images/merchant/69c418c33ec2e1a4267fa9bb77fa1428/business-970x90.gif" border="0"></a>
<!-- affiliate ads end -->
![Check When The Object Is Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/check-when-the-object-is-available.png)

### **3\. Check All References are Declared in the Code**

You can get the Runtime error 424 if all the references are not declared. So, make sure you have declared all the references in the code. To verify this, you can use the debug mode by pressing **F5** or clicking on the **Debug** option.

![Debug Command In Excel](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/debug-command-in-excel.jpg)

### **4\. Check the Macro Security Settings**

Sometimes, the error can occur if macros are disabled in the Macro Security settings. You can check and change the settings by following these steps:

- On the **Developer** tab, in the **Code** section, click **Macro Security**.
- In the **Trust Center** window, select **Enable all macros.**

<!-- affiliate ads begin -->
<a href="https://ancheer.sjv.io/c/5597632/1657301/17326" target="_top" id="1657301"><img src="//a.impactradius-go.com/display-ad/17326-1657301" border="0" alt="" width="1920" height="933"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1657301/17326" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Macro Security Wizard](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/macro-security-wizard.jpg)

- Click **OK**.

<!-- affiliate ads begin -->
<a href="https://united.elfm.net/c/5597632/517826/4704" target="_top" id="517826"><img src="//a.impactradius-go.com/display-ad/4704-517826" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://united.elfm.net/i/5597632/517826/4704" style="position:absolute;visibility:hidden;" border="0" />
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
<li><a href="https://screen-mirroring-recording.techidaily.com/new-2024-approved-securing-premium-image-on-zoom-via-strategic-filters/"><u>[New] 2024 Approved  Securing Premium Image on Zoom via Strategic Filters</u></a></li>
<li><a href="https://article-helps.techidaily.com/new-behind-the-scenes-of-making-memes-funny-for-2024/"><u>[New] Behind-the-Scenes of Making Memes Funny for 2024</u></a></li>
<li><a href="https://vp-tips.techidaily.com/new-convert-subtitles-effortlessly-top-8-best-converters-from-sub-to-srt-format-for-2024/"><u>[New] Convert Subtitles Effortlessly - Top 8 Best Converters From SUB to SRT Format for 2024</u></a></li>
<li><a href="https://screen-mirroring-recording.techidaily.com/new-expert-guide-to-perfectly-shaped-objects-minecraft-circles-and-spheres/"><u>[New] Expert Guide to Perfectly Shaped Objects  Minecraft Circles & Spheres</u></a></li>
<li><a href="https://extra-skills.techidaily.com/new-pioneering-usability-unearthed-strategies-for-amplified-windows-11-capabilities/"><u>[New] Pioneering Usability  Unearthed Strategies for Amplified Windows 11 Capabilities</u></a></li>
<li><a href="https://some-guidance.techidaily.com/new-the-art-of-package-revelation-amplifying-initial-impressions/"><u>[New] The Art of Package Revelation  Amplifying Initial Impressions</u></a></li>
<li><a href="https://instagram-clips.techidaily.com/new-the-insiders-guide-adding-borders-in-instagram-footage-for-2024/"><u>[New] The Insider's Guide  Adding Borders in Instagram Footage for 2024</u></a></li>
<li><a href="https://tiktok-clips.techidaily.com/updated-in-2024-the-complete-mac-powered-guide-to-ootd-videography/"><u>[Updated] In 2024, The Complete Mac-Powered Guide to OOTD Videography</u></a></li>
<li><a href="https://youtube-videos.techidaily.com/2024-approved-a-guide-to-crafting-effective-youtube-backlinks/"><u>2024 Approved  A Guide to Crafting Effective YouTube Backlinks</u></a></li>
<li><a href="https://youtube-help.techidaily.com/2024-approved-navigating-copyright-on-youtube-and-cc/"><u>2024 Approved  Navigating Copyright on YouTube & CC</u></a></li>
<li><a href="https://screen-capture.techidaily.com/androids-prime-portable-gba-simulators-ranked-for-2024/"><u>Android's Prime Portable GBA Simulators Ranked for 2024</u></a></li>
<li><a href="https://ios-unlock.techidaily.com/apple-iphone-xs-asking-for-passcode-after-ios-1714-update-what-to-do-by-drfone-ios/"><u>Apple iPhone XS Asking for Passcode after iOS 17/14 Update, What to Do?</u></a></li>
<li><a href="https://android-location.techidaily.com/for-people-wanting-to-mock-gps-on-itel-p40-devices-drfone-by-drfone-virtual/"><u>For People Wanting to Mock GPS on Itel P40 Devices | Dr.fone</u></a></li>
<li><a href="https://ai-driven-video-production.techidaily.com/how-to-easily-cut-large-videos-faster-on-mac/"><u>How to Easily Cut Large Videos Faster on Mac</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-oneplus-by-fonelab-android-recover-messages/"><u>How to recover old messages from your OnePlus</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-oneplus-ace-2v-by-fonelab-android-recover-messages/"><u>How to recover old messages from your OnePlus Ace 2V</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-oneplus-nord-3-5g-by-fonelab-android-recover-messages/"><u>How to recover old messages from your OnePlus Nord 3 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-oneplus-nord-ce-3-5g-by-fonelab-android-recover-messages/"><u>How to recover old messages from your OnePlus Nord CE 3 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-oppo-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Oppo</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-oppo-a18-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Oppo A18</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-oppo-a1x-5g-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Oppo A1x 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-oppo-a79-5g-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Oppo A79 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-oppo-find-n3-flip-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Oppo Find N3 Flip</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-oppo-find-x6-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Oppo Find X6</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-p55plus-by-fonelab-android-recover-messages/"><u>How to recover old messages from your P55+</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-poco-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Poco</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-pova-5-pro-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Pova 5 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-realme-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Realme</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-realme-11-pro-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Realme 11 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-realme-c33-2023-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Realme C33 2023</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-realme-c51-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Realme C51</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-realme-c55-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Realme C55</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-realme-gt-5-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Realme GT 5</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-red-magic-9-pro-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Red Magic 9 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-samsung-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Samsung</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-samsung-galaxy-a15-5g-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Samsung Galaxy A15 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-samsung-galaxy-a54-5g-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Samsung Galaxy A54 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-samsung-galaxy-f14-5g-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Samsung Galaxy F14 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-samsung-galaxy-s24-ultra-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Samsung Galaxy S24 Ultra</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-smart-8-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Smart 8</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-sony-xperia-5-v-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Sony Xperia 5 V</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-spark-10c-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Spark 10C</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-spark-20c-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Spark 20C</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-tecno-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Tecno</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-tecno-phantom-v-flip-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Tecno Phantom V Flip</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-tecno-spark-10-4g-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Tecno Spark 10 4G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-tecno-spark-10-5g-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Tecno Spark 10 5G</u></a></li>
<li><a href="https://sim-unlock.techidaily.com/in-2024-tips-and-tricks-to-tell-if-your-apple-iphone-xs-max-is-unlocked-by-drfone-ios/"><u>In 2024, Tips And Tricks To Tell if Your Apple iPhone XS Max Is Unlocked</u></a></li>
<li><a href="https://screen-activity-recording.techidaily.com/navigating-the-nuances-of-streaming-and-saving-on-facebook/"><u>Navigating the Nuances of Streaming and Saving on Facebook</u></a></li>
<li><a href="https://video-content-creator.techidaily.com/new-2024-approved-s-most-cost-effective-video-editing-tools/"><u>New 2024 Approved S Most Cost-Effective Video Editing Tools</u></a></li>
<li><a href="https://youtube-clips.techidaily.com/subforsub-review-is-it-safe-to-use-for-increasing-youtube-subscribers/"><u>Subforsub Review? Is It Safe to Use for Increasing YouTube Subscribers?</u></a></li>
<li><a href="https://some-tips.techidaily.com/unleash-creativity-in-editing-leveraging-storyremix-with-windows-11-photos-for-2024/"><u>Unleash Creativity in Editing  Leveraging StoryRemix with Windows 11 Photos for 2024</u></a></li>
</ul></div>

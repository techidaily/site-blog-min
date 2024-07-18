---
title: How to fix runtime error 424 object required error in Excel 2019
date: 2024-07-17T09:59:59.544Z
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
<li><a href="https://facebook-record-videos.techidaily.com/new-2024-approved-a-comprehensive-guide-to-personalizing-video-images/"><u>[New] 2024 Approved  A Comprehensive Guide to Personalizing Video Images</u></a></li>
<li><a href="https://snapchat-videos.techidaily.com/new-achieving-maximum-compactness-with-mac-recordings-on-snapchat-for-2024/"><u>[New] Achieving Maximum Compactness with Mac Recordings on Snapchat for 2024</u></a></li>
<li><a href="https://screen-mirroring-recording.techidaily.com/new-in-2024-a-critical-look-at-itop-for-tech-enthusiasts/"><u>[New] In 2024, A Critical Look at ITop for Tech Enthusiasts</u></a></li>
<li><a href="https://youtube-clips.techidaily.com/updated-a-step-by-step-approach-to-crafting-your-video-market-standouts/"><u>[Updated] A Step-by-Step Approach to Crafting Your Video' Market Standouts</u></a></li>
<li><a href="https://eaxpv-info.techidaily.com/updated-get-unlimited-echoes-for-online-content-makers-for-2024/"><u>[Updated] Get Unlimited Echoes for Online Content Makers for 2024</u></a></li>
<li><a href="https://some-techniques.techidaily.com/updated-giggle-genesis-app/"><u>[Updated] Giggle Genesis App</u></a></li>
<li><a href="https://vimeo-videos.techidaily.com/updated-in-2024-structuring-vimeo-footage-chapter-creation-tips/"><u>[Updated] In 2024, Structuring Vimeo Footage  Chapter Creation Tips</u></a></li>
<li><a href="https://some-skills.techidaily.com/updated-survey-diverse-categories-in-visual-media-tech/"><u>[Updated] Survey  Diverse Categories in Visual Media Tech</u></a></li>
<li><a href="https://screen-mirroring-recording.techidaily.com/updated-utilizing-the-virtual-whiteboard-in-webinars-cross-platform-tips-and-tricks-for-2024/"><u>[Updated] Utilizing the Virtual Whiteboard in Webinars  Cross-Platform Tips and Tricks for 2024</u></a></li>
<li><a href="https://voice-adjusting.techidaily.com/2024-approved-crafting-clear-sounds-on-mac-an-analysis-of-the-top-5-audio-mixers/"><u>2024 Approved Crafting Clear Sounds on Mac An Analysis of the Top 5 Audio Mixers</u></a></li>
<li><a href="https://smart-video-editing.techidaily.com/2024-approved-revolutionize-your-gameplay-videos-top-free-editing-tools/"><u>2024 Approved Revolutionize Your Gameplay Videos Top Free Editing Tools</u></a></li>
<li><a href="https://android-location-track.techidaily.com/9-best-phone-monitoring-apps-for-itel-a60-drfone-by-drfone-virtual-android/"><u>9 Best Phone Monitoring Apps for Itel A60 | Dr.fone</u></a></li>
<li><a href="https://fake-location.techidaily.com/how-to-fix-life360-shows-wrong-location-on-lava-blaze-pro-5g-drfone-by-drfone-virtual-android/"><u>How to Fix Life360 Shows Wrong Location On Lava Blaze Pro 5G? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-galaxy-a05s-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Galaxy A05s?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-galaxy-a15-4g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Galaxy A15 4G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-galaxy-a24-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Galaxy A24?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-galaxy-z-flip-5-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Galaxy Z Flip 5?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-google-pixel-8-pro-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Google Pixel 8 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-honor-100-pro-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Honor 100 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-honor-70-lite-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Honor 70 Lite 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-honor-90-lite-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Honor 90 Lite?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-honor-90-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Honor 90?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-honor-magic-5-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Honor Magic 5?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-honor-magic-6-pro-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Honor Magic 6 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-honor-magic5-ultimate-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Honor Magic5 Ultimate?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-honor-play-40c-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Honor Play 40C?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-honor-play-7t-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Honor Play 7T?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-honor-play-8t-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Honor Play 8T?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-honor-v-purse-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Honor V Purse?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-honor-x50-gt-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Honor X50 GT?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-honor-x50i-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Honor X50i?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-honor-x7b-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Honor X7b?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-honor-x8b-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Honor X8b?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-htc-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your HTC ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-huawei-nova-y71-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Huawei Nova Y71?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-huawei-p60-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Huawei P60?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-infinix-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Infinix ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-infinix-gt-10-pro-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Infinix GT 10 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-infinix-hot-30-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Infinix Hot 30 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-infinix-hot-40-pro-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Infinix Hot 40 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-infinix-note-30-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Infinix Note 30 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-infinix-note-30-pro-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Infinix Note 30 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-infinix-note-30i-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Infinix Note 30i?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-infinix-zero-30-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Infinix Zero 30 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-itel-a60s-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Itel A60s?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-itel-p40-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Itel P40?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-itel-p55-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Itel P55 5G?</u></a></li>
<li><a href="https://extra-skills.techidaily.com/in-2024-leading-vr-creators-industrys-pioneers/"><u>In 2024, Leading VR Creators  Industry's Pioneers</u></a></li>
<li><a href="https://techidaily.com/is-your-oppo-f23-5g-working-too-slow-heres-how-you-can-hard-reset-it-drfone-by-drfone-reset-android-reset-android/"><u>Is your Oppo F23 5G working too slow? Heres how you can hard reset it | Dr.fone</u></a></li>
<li><a href="https://ai-video-tools.techidaily.com/new-the-ultimate-guide-to-subtitling-in-final-cut-pro-x-tips-and-tricks/"><u>New The Ultimate Guide to Subtitling in Final Cut Pro X Tips and Tricks</u></a></li>
<li><a href="https://ai-video-apps.techidaily.com/new-unleash-your-creativity-with-fcp-90-days-of-free-usage/"><u>New Unleash Your Creativity with FCP 90 Days of Free Usage</u></a></li>
<li><a href="https://extra-tips.techidaily.com/novice-necessities-for-superior-gopro-photography/"><u>Novice Necessities for Superior GoPro Photography</u></a></li>
<li><a href="https://extra-lessons.techidaily.com/recording-audio-for-impactful-ppts-and-presentations/"><u>Recording Audio For Impactful PPTs & Presentations</u></a></li>
<li><a href="https://fake-location.techidaily.com/what-is-geo-blocking-and-how-to-bypass-it-on-xiaomi-redmi-13c-drfone-by-drfone-virtual-android/"><u>What is Geo-Blocking and How to Bypass it On Xiaomi Redmi 13C? | Dr.fone</u></a></li>
</ul></div>

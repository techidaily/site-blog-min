---
title: How to fix runtime error 424 object required error in Excel 2000
date: 2024-08-20T20:48:03.625Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to fix runtime error 424 object required error in Excel 2000
keywords: repair .xlsx,repair excel 2023,repair excel 2019,repair .xlsm,repair excel 2010,repair excel
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
<a href="https://aofit.pxf.io/c/5597632/1399701/16396" target="_top" id="1399701"><img src="//a.impactradius-go.com/display-ad/16396-1399701" border="0" alt="" width="960" height="300"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1399701/16396" style="position:absolute;visibility:hidden;" border="0" />
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

### **1\. Check the Name of the Object**

The Runtime error 424 can occur when you run the VBA code using an incorrect name of the object. For example, the object name is ‘MyObject’ but you’re using “Backcolor”.

![Error When Incorrect Name Of The Object](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/error-424-when-incorrect-name-of-the-object.jpg)

 When you click the **Debug** button, the line with the error will highlight.

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2068407/7443" target="_top" id="2068407"><img src="//a.impactradius-go.com/display-ad/7443-2068407" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2068407/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Highlighting Line With Error](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/highlighting-line-with-error.jpg)

To fix the issue, you need to provide the correct name of the object.

### **2\. Check if the Object is Missing**

 The Runtime error 424 can occur if the object you are referring to as a method is not available or you are using the wrong object in a code. In the below example, you can see that the error occurs when an object named “Employee” is not available in the Project list.

![Example Of Code When Object Is Not Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/example-of-code-when-an-object-is-not-available.jpg)

 You can check and mention the object which is available. For instance, Sheet2 in the below code.  

![Check When The Object Is Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/check-when-the-object-is-available.png)

### **3\. Check All References are Declared in the Code**

You can get the Runtime error 424 if all the references are not declared. So, make sure you have declared all the references in the code. To verify this, you can use the debug mode by pressing **F5** or clicking on the **Debug** option.

<!-- affiliate ads begin -->
<a href="https://getlyla.pxf.io/c/5597632/1455723/15391" target="_top" id="1455723"><img src="//a.impactradius-go.com/display-ad/15391-1455723" border="0" alt="" width="336" height="280"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1455723/15391" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Debug Command In Excel](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/debug-command-in-excel.jpg)

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2087389/7443" target="_top" id="2087389"><img src="//a.impactradius-go.com/display-ad/7443-2087389" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2087389/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **4\. Check the Macro Security Settings**

Sometimes, the error can occur if macros are disabled in the Macro Security settings. You can check and change the settings by following these steps:

- On the **Developer** tab, in the **Code** section, click **Macro Security**.
- In the **Trust Center** window, select **Enable all macros.**

<!-- affiliate ads begin -->
<a href="https://store.revouninstaller.com/order/checkout.php?PRODS=27889512&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/4282ec8de8c9be897e7aff4aa231b1a4/728__90.jpg" border="0"></a>
<!-- affiliate ads end -->
![Macro Security Wizard](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/macro-security-wizard.jpg)

- Click **OK**.

<!-- affiliate ads begin -->
<a href="https://modlily.sjv.io/c/5597632/1997817/17059" target="_top" id="1997817"><img src="//a.impactradius-go.com/display-ad/17059-1997817" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1997817/17059" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Enable All Macro In Trust Center](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/enable-all-macros-in-trust-center.jpg)

<!-- affiliate ads begin -->
<a href="https://printrendy.pxf.io/c/5597632/1453720/17020" target="_top" id="1453720"><img src="//a.impactradius-go.com/display-ad/17020-1453720" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1453720/17020" style="position:absolute;visibility:hidden;" border="0" />
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
<li><a href="https://eaxpv-info.techidaily.com/new-2024-approved-how-to-blur-background-of-your-youtube-video/"><u>[New] 2024 Approved  How to Blur Background of Your YouTube Video</u></a></li>
<li><a href="https://digital-screen-recording.techidaily.com/new-in-2024-streamlined-screenshots-on-chrome-devices/"><u>[New] In 2024, Streamlined Screenshots on Chrome Devices</u></a></li>
<li><a href="https://fox-friendly.techidaily.com/updated-in-2024-the-contrast-of-virtual-realms-meta-vs-omni-universe/"><u>[Updated] In 2024, The Contrast of Virtual Realms  Meta Vs. Omni Universe</u></a></li>
<li><a href="https://visual-screen-recording.techidaily.com/updated-professional-meeting-aesthetics-blur-backgrounds-tips-for-2024/"><u>[Updated] Professional Meeting Aesthetics  Blur Backgrounds Tips for 2024</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-itel-p55-5g-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Itel P55 5G to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-realme-narzo-60-5g-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Realme Narzo 60 5G to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://twitter-videos.techidaily.com/2024-approved-2023how-to-share-videos-on-twitter-on-your-phone-without-retweeting/"><u>2024 Approved  2023|How to Share Videos on Twitter on Your Phone Without Retweeting?</u></a></li>
<li><a href="https://some-tips.techidaily.com/2024-approved-top-tier-photographic-archives/"><u>2024 Approved  Top Tier Photographic Archives</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-honor-x50-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Honor X50 to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-itel-a05s-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Itel A05s to iPhone | Dr.fone</u></a></li>
<li><a href="https://apple-account.techidaily.com/detailed-guide-on-removing-apple-iphone-x-activation-lock-without-previous-owner-by-drfone-ios/"><u>Detailed Guide on Removing Apple iPhone X Activation Lock without Previous Owner?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-galaxy-s24plus-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Galaxy S24+</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-6s-plus-to-an-older-ios-system-version-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 6s Plus to an Older iOS System Version? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-missing-or-malfunctioning-drivers-with-windows-device-manager-on-windows-11-by-drivereasy-guide/"><u>How to identify missing or malfunctioning drivers with Windows Device Manager on Windows 11</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-play-mp4-files-on-htc-u23-pro-by-aiseesoft-video-converter-play-mp4-on-android/"><u>How to play MP4 files on HTC U23 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-smart-8-by-fonelab-android-recover-data/"><u>How to recover lost data from Smart 8?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-samsung-galaxy-f54-5g-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Samsung Galaxy F54 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-whatsapp-chat-history-from-iphone-12-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How To Recover Whatsapp Chat History From iPhone 12 | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-contacts-from-f23-5g-by-fonelab-android-recover-contacts/"><u>How to Rescue Lost Contacts from F23 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-photos-from-civi-3-disney-100th-anniversary-edition-by-fonelab-android-recover-photos/"><u>How to Rescue Lost Photos from Civi 3 Disney 100th Anniversary Edition?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-previous-version-of-excel-2007-file-stellar-by-stellar-guide/"><u>How to Restore Previous Version of Excel 2007 File? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-call-history-on-reno-10-pro-5g-by-fonelab-android-recover-call-logs/"><u>How to restore wiped call history on Reno 10 Pro 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-videos-on-motorola-moto-g04-by-fonelab-android-recover-video/"><u>How to restore wiped videos on Motorola Moto G04</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-videos-on-oppo-find-n3-by-fonelab-android-recover-video/"><u>How to restore wiped videos on Oppo Find N3</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-deleted-photos-on-vivo-y100-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Retrieve  deleted photos on Vivo Y100 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-music-from-xiaomi-redmi-a2plus-by-fonelab-android-recover-music/"><u>How to retrieve erased music from Xiaomi Redmi A2+</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/in-2024-15-best-strongest-pokemon-to-use-in-pokemon-go-pvp-leagues-for-oppo-f23-5g-drfone-by-drfone-virtual-android/"><u>In 2024, 15 Best Strongest Pokémon To Use in Pokémon GO PvP Leagues For Oppo F23 5G | Dr.fone</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/in-2024-overview-of-the-best-motorola-moto-g13-screen-mirroring-app-drfone-by-drfone-android/"><u>In 2024, Overview of the Best Motorola Moto G13 Screen Mirroring App | Dr.fone</u></a></li>
<li><a href="https://vp-tips.techidaily.com/in-2024-premier-free-tools-for-easy-jpggif-transformation/"><u>In 2024, Premier Free Tools for Easy JPG/GIF Transformation</u></a></li>
<li><a href="https://extra-guidance.techidaily.com/in-2024-securing-and-optimizing-virtual-meetings-with-zoom-on-windows-11/"><u>In 2024, Securing and Optimizing Virtual Meetings with Zoom on Windows 11</u></a></li>
<li><a href="https://some-tips.techidaily.com/in-2024-the-artists-roadmap-to-professional-growth/"><u>In 2024, The Artist's Roadmap to Professional Growth</u></a></li>
<li><a href="https://mondly-stories.techidaily.com/melodic-flirts-aphrodites-italian-gems/"><u>Melodic Flirts: Aphrodite’s Italian Gems</u></a></li>
<li><a href="https://windows11.techidaily.com/stop-frozen-pages-and-scrolling-issues-in-excel/"><u>Stop Frozen Pages and Scrolling Issues in Excel</u></a></li>
<li><a href="https://tech-haven.techidaily.com/unleashing-the-power-of-chatgpt-discover-these-5-prime-methods/"><u>Unleashing the Power of ChatGPT: Discover These 5 Prime Methods</u></a></li>
<li><a href="https://tech-savvy.techidaily.com/will-natural-language-processing-overthrow-se/"><u>Will Natural Language Processing Overthrow SE?</u></a></li>
</ul></div>

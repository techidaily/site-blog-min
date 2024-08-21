---
title: How to fix runtime error 424 object required error in Excel
date: 2024-08-20T20:47:48.147Z
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
<li><a href="https://instagram-video-files.techidaily.com/new-2024-approved-unlocking-engagement-best-practices-for-highlight-boost/"><u>[New] 2024 Approved  Unlocking Engagement  Best Practices for Highlight Boost</u></a></li>
<li><a href="https://youtube-blog.techidaily.com/ffordable-gear-guide-choosing-the-right-cam-for-you/"><u>[New] Affordable Gear Guide  Choosing the Right Cam For You</u></a></li>
<li><a href="https://extra-guidance.techidaily.com/new-perfecting-your-podcasts-naming-strategy-guide-and-top-ideas-list/"><u>[New] Perfecting Your Podcast's Naming Strategy  Guide and Top Ideas List</u></a></li>
<li><a href="https://extra-support.techidaily.com/new-sony-a6400-video-vanishing-how-to-stop-it/"><u>[New] Sony A6400 Video Vanishing - How to Stop It?</u></a></li>
<li><a href="https://youtube-blog.techidaily.com/ed-navigating-partner-selection-for-youtube-joint-ventures-for-2024/"><u>[Updated] Navigating Partner Selection for YouTube Joint Ventures for 2024</u></a></li>
<li><a href="https://extra-lessons.techidaily.com/10-high-quality-camera-gimbals-iphoneandroiddslr-reconnaissance/"><u>10 High-Quality Camera Gimbals  IPhone/Android/DSLR Reconnaissance</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-techniques-to-transfer-data-from-asus-rog-phone-8-pro-to-iphone-15141312-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Techniques to Transfer Data from Asus ROG Phone 8 Pro to iPhone 15/14/13/12 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-techniques-to-transfer-data-from-vivo-v29-pro-to-iphone-15141312-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Techniques to Transfer Data from Vivo V29 Pro to iPhone 15/14/13/12 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-ways-to-move-contacts-from-xiaomi-redmi-a2plus-to-iphone-131415-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Ways to Move Contacts From Xiaomi Redmi A2+ to iPhone (13/14/15) | Dr.fone</u></a></li>
<li><a href="https://bypass-frp.techidaily.com/a-step-by-step-guide-on-using-adb-and-fastboot-to-remove-frp-lock-from-your-honor-80-pro-straight-screen-edition-by-drfone-android/"><u>A Step-by-Step Guide on Using ADB and Fastboot to Remove FRP Lock from your Honor 80 Pro Straight Screen Edition</u></a></li>
<li><a href="https://youtube-video-recordings.techidaily.com/best-of-breed-top-rated-youtubers-streaming-arsenal/"><u>Best of Breed  Top-Rated Youtuber's Streaming Arsenal</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-6-plus-without-itunes-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 6 Plus without iTunes? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-corrupt-video-files-of-oneplus-using-video-repair-utility-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Fix corrupt video files of OnePlus using Video Repair Utility on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-facetime-call-history-on-iphone-xr-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover FaceTime Call History on iPhone XR | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-oppo-find-x7-ultra-by-fonelab-android-recover-data/"><u>How to recover lost data from Oppo Find X7 Ultra?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-honor-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Honor</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-mdm-from-iphone-12-pro-without-a-computer-by-drfone-ios-unlock-ios-unlock/"><u>How to Remove MDM from iPhone 12 Pro without a computer?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-mdm-from-iphone-15-pro-max-without-a-computer-by-drfone-ios-unlock-ios-unlock/"><u>How to Remove MDM from iPhone 15 Pro Max without a computer?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-a-damaged-video-file-of-a79-5g-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair a Damaged video file of A79 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-honor-magic-6-pro-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Honor Magic 6 Pro on Windows??</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupted-or-damaged-excel-file-with-ease-stellar-by-stellar-guide/"><u>How to Repair Corrupted or Damaged Excel File with Ease? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-system-of-iphone-7-plus-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair System of iPhone 7 Plus? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-tecno-spark-20-pro-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Tecno Spark 20 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-contacts-on-iphone-12-mini-4-methods-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Restore Contacts on iPhone 12 mini (4 Methods) | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-oppo-a1-5g-contacts-an-easy-method-explained-by-fonelab-android-recover-contacts/"><u>How to Restore Deleted Oppo A1 5G Contacts  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-call-history-on-motorola-moto-g-stylus-5g-2023-by-fonelab-android-recover-call-logs/"><u>How to restore wiped call history on Motorola Moto G Stylus 5G (2023)?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-call-history-on-note-30i-by-fonelab-android-recover-call-logs/"><u>How to restore wiped call history on Note 30i?</u></a></li>
<li><a href="https://android-location-track.techidaily.com/in-2024-top-9-nokia-g22-monitoring-apps-for-parental-controls-drfone-by-drfone-virtual-android/"><u>In 2024, Top 9 Nokia G22 Monitoring Apps for Parental Controls | Dr.fone</u></a></li>
<li><a href="https://ai-video-apps.techidaily.com/new-2024-approved-dual-approaches-to-audio-fades-in-final-cut-pro-a-step-by-step-guide/"><u>New 2024 Approved Dual Approaches to Audio Fades in Final Cut Pro A Step-by-Step Guide</u></a></li>
<li><a href="https://extra-resources.techidaily.com/perfecting-your-fisheye-lens-techniques-stepwise/"><u>Perfecting Your Fisheye Lens Techniques Stepwise</u></a></li>
<li><a href="https://tech-revival.techidaily.com/secure-your-gaming-with-vpns-the-ultimate-guide-for-playstation-4-users/"><u>Secure Your Gaming with VPNs: The Ultimate Guide for PlayStation 4 Users</u></a></li>
<li><a href="https://facebook-video-content.techidaily.com/squaring-the-circle-of-facebook-engagement-with-visuals-for-2024/"><u>Squaring the Circle of Facebook Engagement with Visuals for 2024</u></a></li>
</ul></div>

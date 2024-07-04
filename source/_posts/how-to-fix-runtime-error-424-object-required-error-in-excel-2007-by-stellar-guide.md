---
title: How to fix runtime error 424 object required error in Excel 2007
date: 2024-05-19T18:32:11.887Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to fix runtime error 424 object required error in Excel 2007
keywords: repair excel 2016,repair .xlsx,repair excel 2007
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

![Error When Incorrect Name Of The Object](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/error-424-when-incorrect-name-of-the-object.jpg)

 When you click the **Debug** button, the line with the error will highlight.

![Highlighting Line With Error](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/highlighting-line-with-error.jpg)

To fix the issue, you need to provide the correct name of the object.

### **2\. Check if the Object is Missing**

 The Runtime error 424 can occur if the object you are referring to as a method is not available or you are using the wrong object in a code. In the below example, you can see that the error occurs when an object named “Employee” is not available in the Project list.

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
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-messages-files-from-honor-x7b-by-fonelab-android-recover-messages/"><u>How To  Restore Missing Messages Files from Honor X7b</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-call-logs-from-honor-x50-gt-by-fonelab-android-recover-call-logs/"><u>How To  Restore Missing Call Logs from Honor X50 GT</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-photos-from-vivo-y28-5g-by-fonelab-android-recover-photos/"><u>How to Rescue Lost Photos from Vivo Y28 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-call-logs-from-motorola-edge-40-neo-by-fonelab-android-recover-call-logs/"><u>How To  Restore Missing Call Logs from Motorola Edge 40 Neo</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-erase-private-data-from-iphone-11-pro-drfone-by-drfone-ios-full-data-eraser-ios-full-data-eraser/"><u>How To Erase Private Data From iPhone 11 Pro | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-pictures-files-from-find-x6-by-fonelab-android-recover-pictures/"><u>How To  Restore Missing Pictures Files from Find X6.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-infinix-note-30-5g-get-deleted-phone-number-back-with-ease-and-safety-by-fonelab-android-recover-contacts/"><u>How to Infinix Note 30 5G Get Deleted Phone Number Back with Ease and Safety</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-free-up-iphone-12-pro-max-space-drfone-by-drfone-ios-full-data-eraser-ios-full-data-eraser/"><u>How To Free Up iPhone 12 Pro Max Space | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-excel-2021-workbook-stellar-by-stellar-guide/"><u>How to Repair Corrupt Excel 2021 Workbook? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-xiaomi-redmi-13c-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Xiaomi Redmi 13C to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-samsung-galaxy-a05s-pin-by-drfone-android-unlock-android-unlock/"><u>How to remove Samsung Galaxy A05s PIN</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-malfunctioning-hardware-drivers-with-windows-device-manager-in-windows-11-by-drivereasy-guide/"><u>How to identify malfunctioning hardware drivers with Windows Device Manager in Windows 11</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-videos-from-honor-70-lite-5g-by-fonelab-android-recover-video/"><u>How to retrieve erased videos from Honor 70 Lite 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-easily-copy-and-paste-your-forex-and-gold-trades-from-mt5-to-mt4-by-mt4copier-guide/"><u>How To Easily Copy & Paste Your Forex and Gold Trades From MT5 to MT4</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-messages-on-asus-rog-phone-8-by-fonelab-android-recover-messages/"><u>How to restore wiped messages on Asus ROG Phone 8</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-iphone-se-2020-system-issues-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair iPhone SE (2020) System Issues? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-from-iphone-6s-plus-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data from iPhone 6s Plus? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-music-from-narzo-60x-5g-by-fonelab-android-recover-music/"><u>How to retrieve erased music from Narzo 60x 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-vivo-y78plus-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Vivo Y78+?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-call-logs-from-gt-5-by-fonelab-android-recover-call-logs/"><u>How To  Restore Missing Call Logs from GT 5</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-contacts-from-xiaomi-redmi-a2-by-fonelab-android-recover-contacts/"><u>How to recover deleted contacts from Xiaomi Redmi A2.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-videos-on-vivo-v29-by-fonelab-android-recover-video/"><u>How to restore wiped videos on Vivo V29</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-realme-narzo-n55-by-fonelab-android-recover-music/"><u>How to restore wiped music on Realme Narzo N55</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-call-logs-from-vivo-t2-pro-5g-by-fonelab-android-recover-call-logs/"><u>How To  Restore Missing Call Logs from Vivo T2 Pro 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-contacts-from-vivo-y02t-by-fonelab-android-recover-contacts/"><u>How to recover deleted contacts from Vivo Y02T.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-videos-and-music-files-from-iphone-x-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Photos, Videos & Music Files from iPhone X | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-honor-x50iplus-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Honor X50i+ to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-google-frp-lock-on-v27-by-drfone-android-unlock-remove-google-frp/"><u>How to remove Google FRP Lock on V27</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-videos-and-music-files-from-iphone-8-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Photos, Videos & Music Files from iPhone 8 | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-after-format-on-motorola-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery after format on Motorola</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-realme-11x-5g-by-fonelab-android-recover-music/"><u>How to recover old music from your Realme 11X 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-find-lost-iphone-12-backup-files-on-windows-pc-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to find lost iPhone 12 Backup files on Windows PC? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-pictures-files-from-motorola-g54-5g-by-fonelab-android-recover-pictures/"><u>How To  Restore Missing Pictures Files from Motorola G54 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-whatsapp-chat-history-from-iphone-7-plus-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How To Recover Whatsapp Chat History From iPhone 7 Plus | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-photos-from-galaxy-s24-by-fonelab-android-recover-photos/"><u>How to Rescue Lost Photos from Galaxy S24?</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/how-to-mirror-pc-screen-to-asus-rog-phone-8-pro-phones-drfone-by-drfone-android/"><u>How to Mirror PC Screen to Asus ROG Phone 8 Pro Phones? | Dr.fone</u></a></li>
<li><a href="https://android-unlock.techidaily.com/in-2024-top-apps-and-online-tools-to-track-oppo-a59-5g-phone-withwithout-imei-number-by-drfone-android/"><u>In 2024, Top Apps and Online Tools To Track Oppo A59 5G Phone With/Without IMEI Number</u></a></li>
<li><a href="https://sim-unlock.techidaily.com/in-2024-how-to-change-your-sim-pin-code-on-your-huawei-p60-phone-by-drfone-android/"><u>In 2024, How To Change Your SIM PIN Code on Your Huawei P60 Phone</u></a></li>
<li><a href="https://iphone-location.techidaily.com/quick-steps-to-change-weather-location-on-apple-iphone-13-pro-max-drfone-by-drfone-virtual-ios/"><u>Quick Steps to Change Weather Location on Apple iPhone 13 Pro Max | Dr.fone</u></a></li>
<li><a href="https://animation-videos.techidaily.com/updated-how-to-reduce-gif-size-and-best-gif-resizers-for-2024/"><u>Updated How to Reduce GIF Size and Best GIF Resizers for 2024</u></a></li>
<li><a href="https://fix-guide.techidaily.com/restore-missing-app-icon-on-oppo-a1-5g-step-by-step-solutions-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>Restore Missing App Icon on Oppo A1 5G Step-by-Step Solutions | Dr.fone</u></a></li>
<li><a href="https://ai-voice-clone.techidaily.com/2024-approved-top-5-ai-voice-generators-online-supports-all-browsers/"><u>2024 Approved Top 5 AI Voice Generators Online (Supports All Browsers)</u></a></li>
<li><a href="https://fake-location.techidaily.com/ultimate-guide-to-free-pptp-vpn-for-beginners-on-samsung-galaxy-m54-5g-drfone-by-drfone-virtual-android/"><u>Ultimate Guide to Free PPTP VPN For Beginners On Samsung Galaxy M54 5G | Dr.fone</u></a></li>
<li><a href="https://ai-video-editing.techidaily.com/updated-motion-tracking-using-the-best-video-editors/"><u>Updated Motion Tracking Using The Best Video Editors</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/best-software-to-fix-and-repair-corrupt-mp4movavi-video-files-of-infinix-by-stellar-video-repair-mobile-video-repair/"><u>Best software to Fix and Repair Corrupt MP4,MOV,AVI video files of Infinix</u></a></li>
<li><a href="https://android-location-track.techidaily.com/top-10-telegram-spy-tools-on-vivo-g2-for-parents-drfone-by-drfone-virtual-android/"><u>Top 10 Telegram Spy Tools On Vivo G2 for Parents | Dr.fone</u></a></li>
<li><a href="https://fake-location.techidaily.com/prank-your-friends-easy-ways-to-fake-and-share-google-maps-location-on-motorola-moto-e13-drfone-by-drfone-virtual-android/"><u>Prank Your Friends! Easy Ways to Fake and Share Google Maps Location On Motorola Moto E13 | Dr.fone</u></a></li>
<li><a href="https://review-topics.techidaily.com/issues-playing-h265-hevc-video-on-xiaomi-redmi-note-13-5g-by-aiseesoft-video-converter-play-hevc-video-on-android/"><u>Issues playing H.265 HEVC video on Xiaomi Redmi Note 13 5G</u></a></li>
<li><a href="https://fix-guide.techidaily.com/how-to-fix-unresponsive-phone-touchscreen-of-realme-narzo-n53-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>How To Fix Unresponsive Phone Touchscreen Of Realme Narzo N53 | Dr.fone</u></a></li>
<li><a href="https://change-location.techidaily.com/home-button-not-working-on-google-pixel-8-pro-here-are-real-fixes-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>Home Button Not Working on Google Pixel 8 Pro? Here Are Real Fixes | Dr.fone</u></a></li>
<li><a href="https://ai-vdieo-software.techidaily.com/updated-fast-forward-your-videos-a-beginners-guide-to-time-lapse-photography/"><u>Updated Fast Forward Your Videos A Beginners Guide to Time Lapse Photography</u></a></li>
<li><a href="https://location-fake.techidaily.com/3-ways-to-fake-gps-without-root-on-lava-blaze-2-5g-drfone-by-drfone-virtual-android/"><u>3 Ways to Fake GPS Without Root On Lava Blaze 2 5G | Dr.fone</u></a></li>
<li><a href="https://ai-voice-clone.techidaily.com/top-10-ai-subtitle-translators-for-content-creators/"><u>Top 10 AI Subtitle Translators for Content Creators</u></a></li>
<li><a href="https://ai-vdieo-software.techidaily.com/2024-approved-best-lego-stop-motion-makers/"><u>2024 Approved Best Lego Stop Motion Makers</u></a></li>
<li><a href="https://android-unlock.techidaily.com/in-2024-rootjunky-apk-to-bypass-google-frp-lock-for-samsung-galaxy-a15-5g-by-drfone-android/"><u>In 2024, Rootjunky APK To Bypass Google FRP Lock For Samsung Galaxy A15 5G</u></a></li>
<li><a href="https://location-social.techidaily.com/how-to-change-your-lava-agni-2-5g-location-on-twitter-drfone-by-drfone-virtual-android/"><u>How to Change your Lava Agni 2 5G Location on Twitter | Dr.fone</u></a></li>
<li><a href="https://ai-topics.techidaily.com/in-2024-exploring-the-finest-text-to-speech-male-voice-tools-for-natural-output/"><u>In 2024, Exploring the Finest Text-to-Speech Male Voice Tools for Natural Output</u></a></li>
<li><a href="https://activate-lock.techidaily.com/in-2024-unlock-your-device-icloud-dns-bypass-explained-and-tested-plus-easy-alternatives-from-iphone-se-2020-by-drfone-ios/"><u>In 2024, Unlock Your Device iCloud DNS Bypass Explained and Tested, Plus Easy Alternatives From iPhone SE (2020)</u></a></li>
<li><a href="https://android-frp.techidaily.com/hassle-free-ways-to-remove-frp-lock-on-oppo-a1-5g-phones-withwithout-a-pc-by-drfone-android/"><u>Hassle-Free Ways to Remove FRP Lock on Oppo A1 5G Phones with/without a PC</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/failed-to-play-hevc-files-on-redmi-note-12-5g-by-aiseesoft-video-converter-play-hevc-video-on-android/"><u>Failed to play HEVC files on Redmi Note 12 5G</u></a></li>
<li><a href="https://animation-videos.techidaily.com/new-top-10-apps-to-enhance-your-instagram-story-slideshow/"><u>New Top 10 Apps to Enhance Your Instagram Story Slideshow</u></a></li>
<li><a href="https://bypass-frp.techidaily.com/easy-guide-to-vivo-x90s-frp-bypass-with-best-methods-by-drfone-android/"><u>Easy Guide to Vivo X90S FRP Bypass With Best Methods</u></a></li>
<li><a href="https://location-social.techidaily.com/4-most-known-ways-to-find-someone-on-tinder-for-vivo-y200e-5g-by-name-drfone-by-drfone-virtual-android/"><u>4 Most-Known Ways to Find Someone on Tinder For Vivo Y200e 5G by Name | Dr.fone</u></a></li>
<li><a href="https://location-social.techidaily.com/top-7-skype-hacker-to-hack-any-skype-account-on-your-lava-blaze-pro-5g-drfone-by-drfone-virtual-android/"><u>Top 7 Skype Hacker to Hack Any Skype Account On your Lava Blaze Pro 5G | Dr.fone</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/in-2024-8-best-apps-for-screen-mirroring-poco-m6-5g-pc-drfone-by-drfone-android/"><u>In 2024, 8 Best Apps for Screen Mirroring Poco M6 5G PC | Dr.fone</u></a></li>
<li><a href="https://ai-voice-clone.techidaily.com/new-how-to-make-lip-sync-video-without-installing-any-app-for-2024/"><u>New How to Make Lip Sync Video Without Installing Any App for 2024</u></a></li>
<li><a href="https://ai-editing-video.techidaily.com/updated-finest-approach-to-slow-down-a-video-capcut-all-in-one-video-editor-for-2024/"><u>Updated Finest Approach to Slow Down a Video CapCut All-in-One Video Editor for 2024</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/here-are-some-pro-tips-for-pokemon-go-pvp-battles-on-oneplus-11-5g-drfone-by-drfone-virtual-android/"><u>Here are Some Pro Tips for Pokemon Go PvP Battles On OnePlus 11 5G | Dr.fone</u></a></li>
<li><a href="https://change-location.techidaily.com/in-2024-where-is-the-best-place-to-catch-dratini-on-samsung-galaxy-s24plus-drfone-by-drfone-virtual-android/"><u>In 2024, Where Is the Best Place to Catch Dratini On Samsung Galaxy S24+ | Dr.fone</u></a></li>
</ul></div>



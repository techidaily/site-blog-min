---
title: How to Fix Excel 2019 Run Time Error 1004
date: 2024-05-19T18:32:11.557Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2019 Run Time Error 1004
keywords: repair .xltx,repair excel 2003,repair .xlb,repair excel 2013,repair excel 2007
---

## How to Fix Excel Run Time Error 1004

**Summary:** Run-time errors are windows-specific issues that occur while the program is running. This blog will teach you how to fix Excel run-time error 1004. In addition, you’ll learn about an Excel repair tool that can help fix the error 1004 if it occurs due to corruption in Excel files.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

VBA (Microsoft Visual Basic for Application) is an internal programming language in Microsoft Excel. Sometimes, when users try to run VBA or generate a Macro in Excel, the Run-time error 1004 may occur. This error may occur due to the presence of more legend entries in the chart, file conflict, incorrect Macro name, and corrupt Excel files. In this blog, we have discussed the reasons and shared some solutions to resolve run-time error 1004.

## **Why This Error Occurs?**

The run time error 1004 usually occurs when you run a VBA macro with the Legend Entries method to modify the legend entries in the MS Excel chart. It happens when the chart contains more legend entries than the available space, macro name conflicts, corrupt Excel files, or data-types mismatch in the VBA code.

## **Ways to Fix Excel Run-Time Error 1004?**

Try the below workarounds to fix Excel run-time error 1004:

### **Create a Macro to Reduce Chart Legend Font Size**

Sometimes, Excel throws the run-time error when you try to run VBA macro to change the legend entries in a Microsoft Excel chart. This error usually occurs when Microsoft Excel truncates the legend entries because of the more legend entries and less space availability. To fix this, try to create a macro that shrinks/minimize the font size of the Excel chart legend text before the VBA macro, and then restore the font size of the chart legend. Here is the macro code:

```
VBCopy
Sub ResizeLegendEntries()

With Worksheets("Sheet1").ChartObjects(1).Activate
      ' Store the current font size
      fntSZ = ActiveChart.Legend.Font.Size

'Temporarily change the font size.
      ActiveChart.Legend.Font.Size = 2

'Place your LegendEntries macro code here to make
         'the changes that you want to the chart legend.

' Restore the font size.
      ActiveChart.Legend.Font.Size = fntSZ
   End With

End Sub
Note: Make sure you have an Excel chart to run the code on the worksheet.
```

### **Uninstall Microsoft Work**

You may encounter a run-time error 1004 in Excel version 2009 or older versions due to conflicts between Microsoft works and Microsoft Excel. This error usually occurs if your system has both Microsoft Office and Microsoft Works. Uninstalling one of them will fix the issue. Try the below steps to uninstall Microsoft Work:

- First, open the **Task Manager** using the shortcut **CTRL + ALT + DEL** altogether
- The **Task Manager window** is displayed.

![Task Manager Window](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/09/task-manager-window.png)

- Click the **Process** tab, right-click on each program you want to close, and then click **End Task.**
- Stop all the running programs.
- Open the **Run** window and type **_appwiz.cpl_** to open the **Programs and Feature** window.

![Program and Features of Control Panel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/09/program-and-features-1024x516.png)

- Search for **Microsoft Works** and click **Uninstall**.

### **Try Deleting GWXL97.Xla File**

The Add-ins files with .xla extension in MS-EXCEL is used to provide additional functionality to Excel spreadsheets. Sometimes, deleting the GWXL97.XLA file fixes the run-time error. Here are the steps to delete this file:

- Make sure you have an **Admins rights**, open the **Windows Explorer**
- Follow the Path C:\\Programs Files\\MSOffice\\Office\\XLSTART.
- Find and right-click on the **GWXL97.XLA** file
- Click **Delete**.

### **Change Trust Center Settings**

Sometimes, run-time errors might arise because of incorrect security settings. The **Trust Center settings** help you find the **Privacy and security** settings for Microsoft Excel. Follow the below steps to change the **Trust center settings**:

- Open Microsoft Excel.
- Go to **File > Options.**
- The **Excel options** window is displayed.
- Choose **Trust Center**, and click **Trust Center Settings**.
- Tap on the **Macro Settings** tab, and select **Trust access to the VBA project object model.**

![Macro Settings in Microsoft Excel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/09/macro-settings.png)

- Click **OK**.

### **Run Open and Repair Tool**

The Runtime error also arises when MS Excel detects a corrupted worksheet. It automatically begins the File recovery mode and starts repairing it. However, if the Recovery mode fails to start, use the **Open and Repair** tool with the below steps:

- Click **File > Open**.
- Click the location and folder with a corrupted workbook.
- In the **Open** dialog box, choose the corrupted workbook.
- Click the arrow next to the **Open** tab, and go to the **Open and Repair** tab.
- Click **Repair**.

You can also opt for **Stellar Repair for Excel** if the Microsoft Excel’s built-in tool cannot fix the error.

### **Use Stellar Repair for Excel**

**Stellar Repair for Excel** is a professional software for repairing damage. xls, .xlsx, .xltm, .xltx, and .xlsm files and recovering all its objects. Here are the steps to fix the error using this tool:

- First, **download**, **install**, and run **Stellar Repair for Excel**.
- Click the **Browse** tab on the interface window to choose the corrupted Excel file you need to repair.
- Click **Scan**. You will see the scan progress in the scanning window.
- Click **OK**.
- The tool can let you preview all the recoverable Excel file components including tables, pivot tables, charts, formulas, etc.
- Click **Save** to save the repaired file.
- A **Save File dialog box** will appear with the below two options:
- Default location
- New location
- Choose a suitable option.
- Click the **Save** option to repair the Excel file that you have chosen.
- Once the repair is complete, it will display a message “**File repaired successfully**.”
- Click **OK**.

## **Conclusion**

Now you know the Excel run-time error 1004, its cause, and solutions. Follow the workarounds discussed in the blog to rectify the error quickly. However, **[Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)** makes your task of removing run-time errors easy. It’s a powerful software to fix all the issues with Excel files. Also, it helps in extracting data from the damaged file and saves it to a new Excel workbook.




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
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-iphone-x-camera-roll-photos-and-photo-stream-pictures-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted iPhone X Camera Roll Photos and Photo Stream Pictures? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-honor-90-lite-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Honor 90 Lite</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-xiaomi-redmi-12-5g-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Xiaomi Redmi 12 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-lava-blaze-pro-5g-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Lava Blaze Pro 5G to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-whatsapp-chat-history-from-iphone-12-pro-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How To Recover Whatsapp Chat History From iPhone 12 Pro | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-xiaomi-redmi-a2plus-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Xiaomi Redmi A2+ to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-data-from-iphone-11-pro-max-to-other-iphone-11-devices-drfone-by-drfone-transfer-data-from-ios-transfer-data-from-ios/"><u>How To Transfer Data From iPhone 11 Pro Max To Other iPhone 11 devices? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-data-from-lost-or-stolen-iphone-8-in-easy-steps-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How To Recover Data From Lost or Stolen iPhone 8 In Easy Steps | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-call-logs-from-itel-a70-by-fonelab-android-recover-call-logs/"><u>How To  Restore Missing Call Logs from Itel A70</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-redmi-k70-photos-an-easy-method-explained-by-fonelab-android-recover-photos/"><u>How to Restore Deleted Redmi K70 Photos  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-se-2022-without-losing-anything-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone SE (2022) without Losing Anything? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-a-damaged-video-file-of-galaxy-a14-4g-using-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair a Damaged video file of Galaxy A14 4G using Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-call-history-on-realme-11-proplus-by-fonelab-android-recover-call-logs/"><u>How to restore wiped call history on Realme 11 Pro+?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-a-damaged-video-file-of-oppo-k11x-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair a Damaged video file of Oppo K11x?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-pictures-files-from-samsung-galaxy-a23-5g-by-fonelab-android-recover-pictures/"><u>How To  Restore Missing Pictures Files from Samsung Galaxy A23 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-call-logs-from-samsung-galaxy-s24plus-by-fonelab-android-recover-call-logs/"><u>How to rescue lost call logs from Samsung Galaxy S24+</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-a-pdf-v17-document-with-electronic-signature-software-by-ldigisigner-sign-a-pdf-sign-a-pdf/"><u>How to sign a PDF v1.7 document with electronic signature software</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-install-forex-trade-copier-software-on-metatrader-4-and-metatrader-5-by-mt4copier-guide/"><u>How to Install Forex Trade Copier Software on MetaTrader 4 and MetaTrader 5</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-whatsapp-chat-history-from-iphone-8-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How To Recover Whatsapp Chat History From iPhone 8 | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-lost-files-from-lava-agni-2-5g-by-fonelab-android-recover-data/"><u>How to retrieve lost files from Lava Agni 2 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-missing-or-malfunctioning-your-hardware-drivers-with-windows-device-manager-in-windows-11-and-10-and-7-by-drivereasy-guide/"><u>How to identify missing or malfunctioning your hardware drivers with Windows Device Manager in Windows 11 & 10 & 7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-photos-files-from-honor-70-lite-5g-by-fonelab-android-recover-photos/"><u>How To  Restore Missing Photos Files from Honor 70 Lite 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-data-from-iphone-11-pro-to-android-devices-drfone-by-drfone-transfer-data-from-ios-transfer-data-from-ios/"><u>How To Transfer Data From iPhone 11 Pro To Android devices? | Dr.fone</u></a></li>
<li><a href="https://change-location.techidaily.com/in-2024-how-can-i-get-more-stardust-in-pokemon-go-on-xiaomi-redmi-a2-drfone-by-drfone-virtual-android/"><u>In 2024, How can I get more stardust in pokemon go On Xiaomi Redmi A2? | Dr.fone</u></a></li>
<li><a href="https://howto.techidaily.com/how-to-fix-part-of-the-touch-screen-not-working-on-vivo-v27-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>How To Fix Part of the Touch Screen Not Working on Vivo V27 | Dr.fone</u></a></li>
<li><a href="https://sim-unlock.techidaily.com/in-2024-top-imei-unlokers-for-your-motorola-edge-2023-phone-by-drfone-android/"><u>In 2024, Top IMEI Unlokers for Your Motorola Edge 2023 Phone</u></a></li>
<li><a href="https://android-transfer.techidaily.com/in-2024-best-3-software-to-transfer-files-tofrom-your-infinix-note-30-vip-via-a-usb-cable-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>In 2024, Best 3 Software to Transfer Files to/from Your Infinix Note 30 VIP via a USB Cable | Dr.fone</u></a></li>
<li><a href="https://review-topics.techidaily.com/possible-solutions-to-restore-deleted-videos-from-nokia-c210-by-fonelab-android-recover-video/"><u>Possible solutions to restore deleted videos from Nokia C210</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/how-to-lock-apps-on-honor-x9b-to-protect-your-individual-information-by-drfone-android/"><u>How to Lock Apps on Honor X9b to Protect Your Individual Information</u></a></li>
<li><a href="https://android-location-track.techidaily.com/top-5-tracking-apps-to-track-samsung-galaxy-z-fold-5-without-them-knowing-drfone-by-drfone-virtual-android/"><u>Top 5 Tracking Apps to Track Samsung Galaxy Z Fold 5 without Them Knowing | Dr.fone</u></a></li>
<li><a href="https://ai-voice-clone.techidaily.com/new-must-use-livestream-shopping-apps-for-seamless-online-system/"><u>New Must-Use Livestream Shopping Apps for Seamless Online System</u></a></li>
<li><a href="https://unlock-android.techidaily.com/a-complete-guide-to-oem-unlocking-on-xiaomi-redmi-13c-5g-by-drfone-android/"><u>A Complete Guide To OEM Unlocking on Xiaomi Redmi 13C 5G</u></a></li>
<li><a href="https://techidaily.com/turn-off-screen-lock-s23-by-drfone-android-unlock-android-unlock/"><u>Turn Off Screen Lock - S23</u></a></li>
<li><a href="https://location-social.techidaily.com/in-2024-change-location-on-yik-yak-for-your-nokia-c12-to-enjoy-more-fun-drfone-by-drfone-virtual-android/"><u>In 2024, Change Location on Yik Yak For your Nokia C12 to Enjoy More Fun | Dr.fone</u></a></li>
<li><a href="https://howto.techidaily.com/4-solutions-to-fix-unfortunately-your-app-has-stopped-error-on-vivo-y36-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>4 Solutions to Fix Unfortunately Your App Has Stopped Error on Vivo Y36 | Dr.fone</u></a></li>
<li><a href="https://howto.techidaily.com/fix-itel-a60-android-system-webview-crash-2024-issue-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>Fix Itel A60 Android System Webview Crash 2024 Issue | Dr.fone</u></a></li>
<li><a href="https://techidaily.com/how-to-transfer-data-from-apple-iphone-se-2020-to-others-devices-drfone-by-drfone-transfer-data-from-ios-transfer-data-from-ios/"><u>How To Transfer Data From Apple iPhone SE (2020) To Others devices? | Dr.fone</u></a></li>
<li><a href="https://ai-voice-clone.techidaily.com/new-top-10-talking-photo-online-tools-you-have-to-try/"><u>New Top 10 Talking Photo Online Tools You Have To Try</u></a></li>
<li><a href="https://ios-unlock.techidaily.com/full-guide-to-unlock-iphone-12-pro-with-itunes-by-drfone-ios/"><u>Full Guide to Unlock iPhone 12 Pro with iTunes</u></a></li>
<li><a href="https://android-location-track.techidaily.com/solutions-to-spy-on-motorola-edge-2023-with-and-without-jailbreak-drfone-by-drfone-virtual-android/"><u>Solutions to Spy on Motorola Edge 2023 with and without jailbreak | Dr.fone</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/in-2024-can-i-use-itools-gpx-file-to-catch-the-rare-pokemon-on-oneplus-ace-3-drfone-by-drfone-virtual-android/"><u>In 2024, Can I use iTools gpx file to catch the rare Pokemon On OnePlus Ace 3 | Dr.fone</u></a></li>
<li><a href="https://iphone-unlock.techidaily.com/in-2024-how-to-unlock-iphone-12-without-passcode-drfone-by-drfone-ios/"><u>In 2024, How to Unlock iPhone 12 Without Passcode? | Dr.fone</u></a></li>
<li><a href="https://activate-lock.techidaily.com/new-multiple-ways-how-to-remove-icloud-activation-lock-from-your-apple-iphone-xs-by-drfone-ios/"><u>New Multiple Ways How To Remove iCloud Activation Lock From your Apple iPhone XS</u></a></li>
<li><a href="https://ai-topics.techidaily.com/updated-get-closer-to-the-artistic-marvels-of-ai-portrait-generator-for-2024/"><u>Updated Get Closer to the Artistic Marvels of AI Portrait Generator for 2024</u></a></li>
</ul></div>



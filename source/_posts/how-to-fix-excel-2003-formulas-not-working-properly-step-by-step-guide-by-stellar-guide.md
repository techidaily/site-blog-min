---
title: How to Fix Excel 2003 Formulas Not Working Properly | Step-by-Step Guide
date: 2024-05-19T18:32:11.480Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2003 Formulas Not Working Properly | Step-by-Step Guide
keywords: repair excel 2021,repair excel 2010,repair excel,repair excel 2013
---

## How to Fix Excel Formulas Not Working Properly | Step-by-Step Guide

**Summary:** Excel formulas sometimes fail to function correctly and even return an error. This article explains what you might be doing wrong that prevents Excel formulas from working properly and solutions to resolve the issue. If your formulas have disappeared from the Excel spreadsheet and you are having trouble recovering them, you can use an Excel repair tool to recover the formulas.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

When working with Excel formulas, situations may arise when the formula doesn’t calculate or update automatically. Or, you may receive errors by clicking on a formula.

## Problems Causing the ‘Excel Formulas not Working Properly’ Issue and Solutions

Let’s check out the possible reasons that cause Excel formulas to work properly and solutions to resolve the issue.

### Problem 1 – Switching Automatic to Manual Calculation Mode

Automatic and manual are the two modes of calculation in Microsoft Excel.

By default, Excel is set to automatic calculation mode. Everything is recalculated automatically when any changes are made in a worksheet in this mode. You may switch from automatic to manual mode to disable the recalculation of formulas, particularly when working with a large Excel file with too many formulas.

Excel will not calculate automatically when set to manual calculation mode. And this may make you think that the Excel formula is not working properly.

### Solution – Change Calculation Mode from Manual to Automatic

To do so, perform these steps:

- Click on the column with problematic formulas.
- Go to the **Formulas** tab, click the **Calculation Options** drop-down, and select **Automatic**.

![Automatic to Manual Calculation Mode](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/02/switch-manual-to-automatic-calculation-1.png)

### Problem 2 – Missing or Mismatched Parentheses

It’s easy to miss or incorrectly place parentheses or include extra parentheses in a complex formula. If a parenthesis is missing or mismatched and you click Enter after entering a formula, Excel displays a message window suggesting to fix the issue (refer to the screenshot below).

![Missing or Mismatched Parentheses](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/02/missing-or-mismatched-parantheses-2.png)

Clicking ‘Yes’ might help fix the issue. But Excel might not fix the parentheses properly, as it tends to add the missing parentheses at the end of a formula which won’t always be the case.

### Solution – Check for Visual Cues When Typing or Editing a Formula with Parentheses

When typing a formula or editing one, Excel provides visual cues to determine if there’s an issue with the parentheses inserted in a formula. Checking for these visual cues can help you fix missing/mismatched parentheses.

- Excel helps identify parenthesis pairs by highlighting them in different colors. For instance, the pair of parenthesis outside is black.
- Excel does not make the opening parentheses bold. So, if you’ve inserted the last closing parentheses in a formula, you can determine if your parentheses are mismatched.
- Excel helps identify parentheses pairs by highlighting and formatting them with the same color once you cross over them.

### Problem 3 – Formatting Cells in an Excel Formula

When adding a number in an Excel formula, don’t add any decimal separator or special characters like $ or €. You may use a comma to separate a function’s argument in an Excel formula or use a currency sign like $ or € as part of cell references. Formatting the numbers may prevent the formula from functioning correctly.

### Solution – Use Format Cells Option for Formatting

Use Format Cells instead of using a comma or currency signs for formatting a number in the formula. For instance, rather than entering a value of $10,000 in your formula, insert 10000, and click the ‘Ctrl+1’ keys together to open the Format Cells dialog box.

![Format Cells Option](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/02/format-cells-3-1.png)

### Problem 4 – Formatting Numbers as Text

Numbers are displayed as left-aligned in a sheet in a worksheet, and text formatted numbers are right-aligned in cells. Excel considers numbers formatted as text to be text strings. Thus, it leaves those numbers out of calculations. As a result, a formula won’t work as intended. For example, in the following screenshot, you can see that the SUM formula works correctly for normal numbers. But, when the SUM formula is applied to numbers formatted as text, the formula doesn’t return the correct value.

![Cells Formatted as Text](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/02/cells-formatted-as-text-4.png)

Sometimes, you may also see an apostrophe in the cells or green triangles in the top-left corner of all the cells when numbers in those cells are formatted as Text.

### Solution – Do Not Format Numbers as Text

To fix the issue, do the following:

- Select the cells with numbers stored as text, right-click on them, and click Format Cells.
- From the Format Cells window, click on Number and then press OK.

![Format Cells in Excel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/02/format-cells-5.png)

### Problem 5 – Double Quotes to Enclose Numbers

Avoid enclosing numbers in a formula in double-quotes, as the numbers are interpreted as a string value.

Meaning if you enter a formula like =IF(A1>B1, “1”), Excel will consider the output one as a string and not a number. So, you won’t be able to use 1’s in calculations.

### Solution – Don’t Enclose Numbers in Double Quotes

Remove any double quotes around a number in your formula unless you want that number to be treated as text. For example, you can write the formula mentioned above as “1” =IF(A1>B1, 1).

### Problem 6 – Extra Space at Beginning of the Formula

When entering a formula, you may end up adding an extra space before the equal (=) sign. You may also add an apostrophe (‘) in the formula at times. As a result, the calculation won’t be performed and may return an error. This usually happens when you use a formula copied from the web.

### Solution – Remove Extra Space from the Formula

The fix to this issue is pretty simple. You need to look for extra space before the equal sign and remove it. Also, ensure there is an additional apostrophe added in the formula.

## Other Things to Consider to Fix the ‘Excel Formulas not Working Properly’ Issue

- If your Excel formula is not showing the result as intended, see this [blog](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/).
- When you refer to other worksheets with spaces or any non-alphabetical character in their names, enclose the names in ‘single quotation marks’. For example, an external 5reference to cell A2 in a sheet named Data enclose the name in single quotes: **‘Data’!A1**.
- You may see the formula instead of the result if you have accidentally clicked the ‘Show Formulas’ option. So, click on the problematic cell, click on the Formula tab, and then click Show Formulas.
- If you’re getting an error “Excel found a problem with one or more formula references in this worksheet”, find solutions to fix the error [here](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/).

## Conclusion

This blog discussed some problems you might make causing an Excel formula to stop working properly. Read about these common problems and solutions to fix them. If a problem doesn’t apply in your case, move to the next one. If you cannot retrieve formulas in your Excel sheet, using an [Excel file repair tool](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) like Stellar Repair for Excel can help you restore all the formulas. It does so by repairing the Excel file (XLS/XLSX) and recovering all the components, including formulas.

[![Free Download for Windows](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/03/free-download-windows-2.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)




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
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-after-format-on-v29e-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery after format on V29e</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-hot-40-pro-get-deleted-pictures-back-with-ease-and-safety-by-fonelab-android-recover-pictures/"><u>How to Hot 40 Pro Get Deleted Pictures Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-oppo-a79-5g-photos-an-easy-method-explained-by-fonelab-android-recover-photos/"><u>How to Restore Deleted Oppo A79 5G Photos  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-iphone-xs-max-ios-system-issues-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair iPhone XS Max iOS System Issues? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-install-and-update-device-drivers-manually-on-windows-11107-by-drivereasy-guide/"><u>How to install and update device drivers manually on Windows 11/10/7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-vivo-x100-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Vivo X100?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-google-pixel-fold-on-mac-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Google Pixel Fold on Mac?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-asus-rog-phone-7-with-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Asus ROG Phone 7 with Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-samsung-pictures-an-easy-method-explained-by-fonelab-android-recover-pictures/"><u>How to Restore Deleted Samsung Pictures  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-photos-files-from-honor-90-gt-by-fonelab-android-recover-photos/"><u>How To  Restore Missing Photos Files from Honor 90 GT.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-p55plus-using-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of P55+ using Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-free-up-iphone-12-space-drfone-by-drfone-ios-full-data-eraser-ios-full-data-eraser/"><u>How To Free Up iPhone 12 Space | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-messages-on-vivo-y36-by-fonelab-android-recover-messages/"><u>How to restore wiped messages on Vivo Y36</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-messages-from-tecno-pova-5-by-fonelab-android-recover-messages/"><u>How to retrieve erased messages from Tecno Pova 5</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-lost-files-from-nokia-c12-by-fonelab-android-recover-data/"><u>How to retrieve lost files from Nokia C12?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-your-iphone-12-mini-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Reset Your iPhone 12 mini? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-corrupt-video-files-of-asus-rog-phone-7-using-video-repair-utility-by-stellar-video-repair-mobile-video-repair/"><u>How to Fix Corrupt video files of Asus ROG Phone 7 using Video Repair Utility?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-videos-on-honor-x8b-by-fonelab-android-recover-video/"><u>How to restore wiped videos on Honor X8b</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-app-on-htc-u23-pro-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Recover Deleted Photos from Android Gallery App on HTC U23 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-word-2023-electronically-by-ldigisigner-sign-a-word-sign-a-word/"><u>How to sign Word 2023 electronically</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-app-on-magic-5-pro-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Recover Deleted Photos from Android Gallery App on Magic 5 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-f5-pro-5g-get-deleted-pictures-back-with-ease-and-safety-by-fonelab-android-recover-pictures/"><u>How to F5 Pro 5G Get Deleted Pictures Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-red-magic-9-pro-get-deleted-pictures-back-with-ease-and-safety-by-fonelab-android-recover-pictures/"><u>How to Red Magic 9 Pro Get Deleted Pictures Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-call-logs-from-oppo-a2-by-fonelab-android-recover-call-logs/"><u>How to rescue lost call logs from Oppo A2</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-app-on-pova-5-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Recover Deleted Photos from Android Gallery App on Pova 5</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-14-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 14? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-motorola-moto-g-stylus-5g-2023-by-fonelab-android-recover-photos/"><u>How to recover deleted photos from Motorola Moto G Stylus 5G (2023).</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-videos-not-playing-on-vivo-x-flip-by-stellar-video-repair-mobile-video-repair/"><u>How to Fix Videos Not Playing on Vivo X Flip?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-data-from-iphone-12-pro-max-to-other-iphone-12-pro-max-devices-drfone-by-drfone-transfer-data-from-ios-transfer-data-from-ios/"><u>How To Transfer Data From iPhone 12 Pro Max To Other iPhone 12 Pro Max devices? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-after-format-on-itel-s23plus-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery after format on Itel S23+</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-word-2007-free-by-ldigisigner-sign-a-word-sign-a-word/"><u>How to sign Word 2007 free</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-google-frp-lock-on-honor-x50-by-drfone-android-unlock-remove-google-frp/"><u>How to remove Google FRP Lock on Honor X50</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-15-pro-max-without-itunes-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Reset iPhone 15 Pro Max Without iTunes? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-messages-files-from-tecno-spark-10-4g-by-fonelab-android-recover-messages/"><u>How To  Restore Missing Messages Files from Tecno Spark 10 4G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-samsung-galaxy-s21-fe-5g-2023-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Samsung Galaxy S21 FE 5G (2023)?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-malfunctioning-hardware-drivers-with-windows-device-manager-in-windows-7-by-drivereasy-guide/"><u>How to identify malfunctioning hardware drivers with Windows Device Manager in Windows 7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-exit-dfu-mode-on-iphone-x-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Exit DFU Mode on iPhone X? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-back-lost-photos-from-oppo-reno-10-proplus-5g-by-fonelab-android-recover-photos/"><u>How to get back lost photos from Oppo Reno 10 Pro+ 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-oppo-reno-9a-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Oppo Reno 9A to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-screenshots-on-iphone-7-plus-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Screenshots on iPhone 7 Plus? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-a05s-with-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of A05s with Video Repair Utility on Windows?</u></a></li>
<li><a href="https://android-unlock.techidaily.com/downloading-samfw-frp-tool-30-for-oppo-f25-pro-5g-by-drfone-android/"><u>Downloading SamFw FRP Tool 3.0 for Oppo F25 Pro 5G</u></a></li>
<li><a href="https://review-topics.techidaily.com/in-2024-dose-life360-notify-me-when-someone-checks-my-location-on-lava-blaze-2-pro-drfone-by-drfone-virtual-android/"><u>In 2024, Dose Life360 Notify Me When Someone Checks My Location On Lava Blaze 2 Pro? | Dr.fone</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/in-2024-best-apple-iphone-13-and-ipad-screen-mirroring-app-drfone-by-drfone-ios/"><u>In 2024, Best Apple iPhone 13 & iPad Screen Mirroring App | Dr.fone</u></a></li>
<li><a href="https://fake-location.techidaily.com/looking-for-a-location-changer-on-xiaomi-redmi-a2-look-no-further-drfone-by-drfone-virtual-android/"><u>Looking For A Location Changer On Xiaomi Redmi A2? Look No Further | Dr.fone</u></a></li>
<li><a href="https://apple-account.techidaily.com/in-2024-how-to-get-the-apple-id-verification-code-on-apple-iphone-13-pro-max-in-the-best-ways-by-drfone-ios/"><u>In 2024, How To Get the Apple ID Verification Code On Apple iPhone 13 Pro Max in the Best Ways</u></a></li>
<li><a href="https://location-social.techidaily.com/3-things-you-must-know-about-fake-snapchat-location-on-apple-iphone-se-2022-drfone-by-drfone-virtual-ios/"><u>3 Things You Must Know about Fake Snapchat Location On Apple iPhone SE (2022) | Dr.fone</u></a></li>
<li><a href="https://pokemon-go-android.techidaily.com/pokemon-go-error-12-failed-to-detect-location-on-realme-11-5g-drfone-by-drfone-virtual-android/"><u>Pokemon Go Error 12 Failed to Detect Location On Realme 11 5G? | Dr.fone</u></a></li>
<li><a href="https://android-location.techidaily.com/getting-the-pokemon-go-gps-signal-not-found-11-error-in-vivo-s17-drfone-by-drfone-virtual/"><u>Getting the Pokemon Go GPS Signal Not Found 11 Error in Vivo S17 | Dr.fone</u></a></li>
<li><a href="https://change-location.techidaily.com/latest-way-to-get-shiny-meltan-box-in-pokemon-go-mystery-box-on-vivo-x-flip-drfone-by-drfone-virtual-android/"><u>Latest way to get Shiny Meltan Box in Pokémon Go Mystery Box On Vivo X Flip | Dr.fone</u></a></li>
<li><a href="https://ios-unlock.techidaily.com/in-2024-how-to-unlock-your-iphone-6-passcode-4-easy-methods-with-or-without-itunes-by-drfone-ios/"><u>In 2024, How to Unlock Your iPhone 6 Passcode 4 Easy Methods (With or Without iTunes)</u></a></li>
<li><a href="https://android-unlock.techidaily.com/the-ultimate-guide-how-to-bypass-swipe-screen-to-unlock-on-vivo-y27s-device-by-drfone-android/"><u>The Ultimate Guide How to Bypass Swipe Screen to Unlock on Vivo Y27s Device</u></a></li>
<li><a href="https://change-location.techidaily.com/full-guide-to-catch-100-iv-pokemon-using-a-map-on-vivo-y28-5g-drfone-by-drfone-virtual-android/"><u>Full Guide to Catch 100 IV Pokémon Using a Map On Vivo Y28 5G | Dr.fone</u></a></li>
<li><a href="https://ai-video-editing.techidaily.com/updated-know-how-to-create-obs-slideshow-creating-obs-slideshow-in-simple-steps-application-for-creating-obs-slideshow-make-obs-slideshow-online-for-2024/"><u>Updated Know How to Create OBS Slideshow. Creating OBS Slideshow in Simple Steps; Application for Creating OBS Slideshow. Make OBS Slideshow Online for 2024</u></a></li>
<li><a href="https://ai-voice-clone.techidaily.com/a-complete-guide-of-top-7-video-language-translators-for-2024/"><u>A Complete Guide of Top 7 Video Language Translators for 2024</u></a></li>
</ul></div>



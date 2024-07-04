---
title: How to Fix Excel 2010 Formulas Not Working Properly | Step-by-Step Guide
date: 2024-05-19T18:32:11.523Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2010 Formulas Not Working Properly | Step-by-Step Guide
keywords: repair .xltm,repair excel 2010,repair excel 2021,repair .xlb,repair .xltx,repair excel 2000
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
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-oppo-reno-10-proplus-5g-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Oppo Reno 10 Pro+ 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-infinix-note-30-vip-racing-edition-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Infinix Note 30 VIP Racing Edition to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-app-on-vivo-x-flip-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Recover Deleted Photos from Android Gallery App on Vivo X Flip</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-oppo-reno-9a-by-fonelab-android-recover-photos/"><u>How to recover deleted photos from Oppo Reno 9A.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-14-without-losing-anything-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 14 without Losing Anything? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-lava-storm-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Lava Storm 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-oppo-reno-10-proplus-5g-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Oppo Reno 10 Pro+ 5G to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-previous-version-of-excel-2003-file-by-stellar-guide/"><u>How to Restore Previous Version of Excel 2003 File?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-word-2010-electronically-by-ldigisigner-sign-a-word-sign-a-word/"><u>How to sign Word 2010 electronically</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-vivo-y27s-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Vivo Y27s to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-lost-files-from-lava-by-fonelab-android-recover-data/"><u>How to retrieve lost files from Lava ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-make-a-digital-signature-for-xlsx-files-by-ldigisigner-sign-a-excel-sign-a-excel/"><u>How to make a digital signature for .xlsx files</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-data-from-lost-or-stolen-iphone-11-pro-in-easy-steps-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How To Recover Data From Lost or Stolen iPhone 11 Pro In Easy Steps | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-the-unable-to-record-macro-error-in-excel-2019-stellar-by-stellar-guide/"><u>How to Fix the Unable to Record Macro Error in Excel 2019? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-music-on-samsung-galaxy-xcover-7-by-fonelab-android-recover-music/"><u>How to restore wiped music on Samsung Galaxy XCover 7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-6-plus-to-an-older-ios-system-version-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 6 Plus to an Older iOS System Version? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-app-on-infinix-smart-8-hd-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Recover Deleted Photos from Android Gallery App on Infinix Smart 8 HD</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-google-frp-lock-on-oppo-a1-5g-by-drfone-android-unlock-remove-google-frp/"><u>How to remove Google FRP Lock on Oppo A1 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-oppo-get-deleted-pictures-back-with-ease-and-safety-by-fonelab-android-recover-pictures/"><u>How to Oppo Get Deleted Pictures Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-deleted-photos-on-motorola-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Retrieve deleted photos on Motorola</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-oppo-a38-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Oppo A38 to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-lost-data-on-iphone-12-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Lost Data on iPhone 12? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-nokia-130-music-get-deleted-pictures-back-with-ease-and-safety-by-fonelab-android-recover-pictures/"><u>How to Nokia 130 Music Get Deleted Pictures Back with Ease and Safety?</u></a></li>
<li><a href="https://ai-vdieo-software.techidaily.com/2024-approved-demystifying-image-ratios-a-calculator-and-resource-guide/"><u>2024 Approved Demystifying Image Ratios A Calculator and Resource Guide</u></a></li>
<li><a href="https://techidaily.com/the-easiest-methods-to-hard-reset-motorola-razr-40-ultra-drfone-by-drfone-reset-android-reset-android/"><u>The Easiest Methods to Hard Reset Motorola Razr 40 Ultra | Dr.fone</u></a></li>
<li><a href="https://android-frp.techidaily.com/in-2024-meizu-21-adb-format-tool-for-pc-vs-other-unlocking-tools-which-one-is-the-best-by-drfone-android/"><u>In 2024, Meizu 21 ADB Format Tool for PC vs. Other Unlocking Tools Which One is the Best?</u></a></li>
<li><a href="https://android-transfer.techidaily.com/in-2024-8-ways-to-transfer-photos-from-tecno-spark-10-4g-to-iphone-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>In 2024, 8 Ways to Transfer Photos from Tecno Spark 10 4G to iPhone Easily | Dr.fone</u></a></li>
<li><a href="https://review-topics.techidaily.com/mp4-video-repair-tool-repair-corrupt-damaged-unplayable-video-files-of-oppo-a58-4g-by-stellar-video-repair-mobile-video-repair/"><u>MP4 Video Repair Tool - Repair corrupt, damaged, unplayable video files of Oppo A58 4G</u></a></li>
<li><a href="https://android-unlock.techidaily.com/how-to-unlock-a-vivo-x-fold-2-easily-by-drfone-android/"><u>How To Unlock a Vivo X Fold 2 Easily?</u></a></li>
<li><a href="https://iphone-unlock.techidaily.com/reset-itunes-backup-password-of-iphone-13-prevention-and-solution-drfone-by-drfone-ios/"><u>Reset iTunes Backup Password Of iPhone 13 Prevention & Solution | Dr.fone</u></a></li>
<li><a href="https://pokemon-go-android.techidaily.com/in-2024-top-15-augmented-reality-games-like-pokemon-go-to-play-on-honor-magic-5-drfone-by-drfone-virtual-android/"><u>In 2024, Top 15 Augmented Reality Games Like Pokémon GO To Play On Honor Magic 5 | Dr.fone</u></a></li>
<li><a href="https://review-topics.techidaily.com/how-to-transfer-data-from-iphone-14-pro-to-other-iphone-13-devices-drfone-by-drfone-transfer-data-from-ios-transfer-data-from-ios/"><u>How To Transfer Data From iPhone 14 Pro To Other iPhone 13 devices? | Dr.fone</u></a></li>
<li><a href="https://howto.techidaily.com/fixes-for-apps-keep-crashing-on-oneplus-nord-ce-3-5g-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>Fixes for Apps Keep Crashing on OnePlus Nord CE 3 5G | Dr.fone</u></a></li>
<li><a href="https://review-topics.techidaily.com/recover-your-photos-after-realme-narzo-n53-has-been-deleted-by-fonelab-android-recover-photos/"><u>Recover your photos after Realme Narzo N53 has been deleted.</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/how-to-do-vivo-y100-5g-screen-sharing-drfone-by-drfone-android/"><u>How To Do Vivo Y100 5G Screen Sharing | Dr.fone</u></a></li>
<li><a href="https://pokemon-go-android.techidaily.com/latest-way-to-get-shiny-meltan-box-in-pokemon-go-mystery-box-on-poco-c65-drfone-by-drfone-virtual-android/"><u>Latest way to get Shiny Meltan Box in Pokémon Go Mystery Box On Poco C65 | Dr.fone</u></a></li>
<li><a href="https://android-frp.techidaily.com/hassle-free-ways-to-remove-frp-lock-on-nokia-c02withwithout-a-pc-by-drfone-android/"><u>Hassle-Free Ways to Remove FRP Lock on Nokia C02with/without a PC</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/additional-tips-about-sinnoh-stone-for-xiaomi-14-drfone-by-drfone-virtual-android/"><u>Additional Tips About Sinnoh Stone For Xiaomi 14 | Dr.fone</u></a></li>
<li><a href="https://location-social.techidaily.com/how-to-pause-life360-location-sharing-for-oneplus-12-drfone-by-drfone-virtual-android/"><u>How To Pause Life360 Location Sharing For OnePlus 12 | Dr.fone</u></a></li>
<li><a href="https://ios-unlock.techidaily.com/how-to-open-your-apple-iphone-15-plus-without-a-home-button-by-drfone-ios/"><u>How To Open Your Apple iPhone 15 Plus Without a Home Button</u></a></li>
<li><a href="https://apple-account.techidaily.com/how-to-sign-out-of-apple-id-from-iphone-13-pro-without-password-by-drfone-ios/"><u>How to Sign Out of Apple ID From iPhone 13 Pro without Password?</u></a></li>
<li><a href="https://ai-voice-clone.techidaily.com/new-how-to-make-explainer-videosstep-by-step-guide-for-2024/"><u>New How to Make Explainer Videos—Step by Step Guide for 2024</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/honor-90-lite-video-recovery-recover-deleted-videos-from-honor-90-lite-by-fonelab-android-recover-video/"><u>Honor 90 Lite Video Recovery - Recover Deleted Videos from Honor 90 Lite</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/planning-to-use-a-pokemon-go-joystick-on-oppo-reno-10-proplus-5g-drfone-by-drfone-virtual-android/"><u>Planning to Use a Pokemon Go Joystick on Oppo Reno 10 Pro+ 5G? | Dr.fone</u></a></li>
<li><a href="https://unlock-android.techidaily.com/in-2024-5-solutions-for-xiaomi-civi-3-disney-100th-anniversary-edition-unlock-without-password-by-drfone-android/"><u>In 2024, 5 Solutions For Xiaomi Civi 3 Disney 100th Anniversary Edition Unlock Without Password</u></a></li>
</ul></div>



---
title: How to Fix Excel 2000 Formulas Not Working Properly | Step-by-Step Guide
date: 2024-05-19T18:32:11.470Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2000 Formulas Not Working Properly | Step-by-Step Guide
keywords: repair excel 2021,repair excel,repair excel 2013,repair excel 2007,repair .xlb,repair excel 2000
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
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-after-format-on-vivo-y78t-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery after format on Vivo Y78t</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-find-lost-iphone-6s-plus-backup-files-on-windows-pc-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to find lost iPhone 6s Plus Backup files on Windows PC? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-messages-on-honor-magic-6-lite-by-fonelab-android-recover-messages/"><u>How to restore wiped messages on Honor Magic 6 Lite</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-txt-file-document-electronically-by-ldigisigner-sign-a-word-sign-a-word/"><u>How to sign .txt file document electronically</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-pictures-from-tecno-phantom-v-flip-by-fonelab-android-recover-pictures/"><u>How to Rescue Lost Pictures from Tecno Phantom V Flip?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-data-from-iphone-using-stellar-data-recovery-for-iphone-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Data from iPhone using Stellar Data Recovery for iPhone? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-vivo-x100-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Vivo X100</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-nokia-130-music-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Nokia 130 Music to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-oppo-reno-11-pro-5g-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Oppo Reno 11 Pro 5G to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-sony-xperia-5-v-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Sony Xperia 5 V</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-play-hevc-h-265-video-on-xiaomi-redmi-12-5g-by-aiseesoft-video-converter-play-hevc-video-on-android/"><u>How to play HEVC H.265 video on Xiaomi Redmi 12 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-after-format-on-poco-c65-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery after format on Poco C65</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-6-plus-to-an-older-ios-version-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 6 Plus to an Older iOS Version? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-contacts-files-from-s17e-by-fonelab-android-recover-contacts/"><u>How To  Restore Missing Contacts Files from S17e.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-videos-on-infinix-smart-8-by-fonelab-android-recover-video/"><u>How to restore wiped videos on Infinix Smart 8</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-contacts-from-oppo-reno-8t-5g-by-fonelab-android-recover-contacts/"><u>How to Rescue Lost Contacts from Oppo Reno 8T 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-7-plus-to-factory-settings-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Reset iPhone 7 Plus to Factory Settings? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-itel-a60s-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Itel A60s?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-call-history-on-lava-by-fonelab-android-recover-call-logs/"><u>How to restore wiped call history on Lava ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-v-purse-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your V Purse?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-pictures-from-nokia-xr21-by-fonelab-android-recover-pictures/"><u>How to recover deleted pictures from Nokia XR21.</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-techniques-to-transfer-data-from-honor-play-8t-to-iphone-15141312-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Techniques to Transfer Data from Honor Play 8T to iPhone 15/14/13/12 | Dr.fone</u></a></li>
<li><a href="https://change-location.techidaily.com/pokemon-go-cooldown-chart-on-vivo-v27e-drfone-by-drfone-virtual-android/"><u>Pokémon Go Cooldown Chart On Vivo V27e | Dr.fone</u></a></li>
<li><a href="https://ai-video-editing.techidaily.com/1713950888678-looking-for-the-ways-by-which-subtitles-can-be-extracted-from-your-mkv-files-without-affecting-the-file-quality-look-no-further-as-we-are-here-to-help-you-k/"><u>Looking for the Ways by Which Subtitles Can Be Extracted From Your MKV Files without Affecting the File Quality? Look No Further as We Are Here to Help You Know the Best Tools that Can Easily Extract SRT From MKV for 2024</u></a></li>
<li><a href="https://fake-location.techidaily.com/spoofing-life360-how-to-do-it-on-nokia-c110-drfone-by-drfone-virtual-android/"><u>Spoofing Life360 How to Do it on Nokia C110? | Dr.fone</u></a></li>
<li><a href="https://howto.techidaily.com/troubleshooting-guide-how-to-fix-an-unresponsive-honor-x7b-screen-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>Troubleshooting Guide How to Fix an Unresponsive Honor X7b Screen | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/best-android-data-recovery-undelete-lost-call-logs-from-tecno-spark-20c-by-fonelab-android-recover-call-logs/"><u>Best Android Data Recovery - undelete lost call logs from Tecno Spark 20C</u></a></li>
<li><a href="https://activate-lock.techidaily.com/3-effective-ways-to-unlock-icloud-account-without-password-on-iphone-7-plus-by-drfone-ios/"><u>3 Effective Ways to Unlock iCloud Account Without Password On iPhone 7 Plus</u></a></li>
<li><a href="https://ai-editing-video.techidaily.com/updated-2024-approved-how-to-quick-guide-for-making-google-slideshows/"><u>Updated 2024 Approved How To Quick Guide for Making Google Slideshows</u></a></li>
<li><a href="https://iphone-unlock.techidaily.com/in-2024-how-do-you-remove-restricted-mode-on-iphone-12-pro-max-drfone-by-drfone-ios/"><u>In 2024, How Do You Remove Restricted Mode on iPhone 12 Pro Max | Dr.fone</u></a></li>
<li><a href="https://ai-video-apps.techidaily.com/new-in-2024-the-complete-aiff-converter-resource-reviews-comparisons-and-more/"><u>New In 2024, The Complete AIFF Converter Resource Reviews, Comparisons, and More</u></a></li>
<li><a href="https://animation-videos.techidaily.com/updated-2024-approved-4-things-you-need-to-know-about-pinterest-gif/"><u>Updated 2024 Approved 4 Things You Need to Know About Pinterest GIF</u></a></li>
<li><a href="https://android-location.techidaily.com/in-2024-fake-android-location-without-rooting-for-your-infinix-smart-7-drfone-by-drfone-virtual/"><u>In 2024, Fake Android Location without Rooting For Your Infinix Smart 7 | Dr.fone</u></a></li>
</ul></div>



---
title: How to Fix Excel Formulas Not Working Properly | Step-by-Step Guide
date: 2024-05-19T18:32:11.605Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel Formulas Not Working Properly | Step-by-Step Guide
keywords: repair .xlsx,repair excel 2013,repair excel 2010,repair .xlb
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
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-photos-files-from-realme-narzo-60-5g-by-fonelab-android-recover-photos/"><u>How To  Restore Missing Photos Files from Realme Narzo 60 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-moto-g34-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Moto G34 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-contacts-from-itel-a60-by-fonelab-android-recover-contacts/"><u>How to recover deleted contacts from Itel A60.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-whatsapp-chat-history-from-iphone-8-plus-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How To Recover Whatsapp Chat History From iPhone 8 Plus | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-pictures-from-motorola-razr-40-ultra-by-fonelab-android-recover-pictures/"><u>How to Rescue Lost Pictures from Motorola Razr 40 Ultra?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-iphone-12-pro-from-backup-when-itunes-backup-is-corrupt-or-not-compatible-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Restore iPhone 12 Pro from Backup when iTunes Backup is Corrupt or not compatible | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-contacts-files-from-htc-u23-by-fonelab-android-recover-contacts/"><u>How To  Restore Missing Contacts Files from HTC U23.</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-techniques-to-transfer-data-from-honor-x50-to-iphone-15141312-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Techniques to Transfer Data from Honor X50 to iPhone 15/14/13/12 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-data-from-broken-iphone-6s-screen-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Data from Broken iPhone 6s Screen | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-y36i-by-fonelab-android-recover-messages/"><u>How to recover old messages from your Y36i</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-samsung-galaxy-a25-5g-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Samsung Galaxy A25 5G? </u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-contacts-from-tecno-camon-20-pro-5g-by-fonelab-android-recover-contacts/"><u>How to recover deleted contacts from Tecno Camon 20 Pro 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-se-2020-to-the-previous-iosipados-version-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone SE (2020) to the Previous iOS/iPadOS Version? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-sign-doc-file-document-with-digital-signature-tutorial-by-ldigisigner-sign-a-word-sign-a-word/"><u>How to Sign .doc file document with Digital Signature - (Tutorial)</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-storm-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Storm 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-videos-on-huawei-nova-y71-by-fonelab-android-recover-video/"><u>How to restore wiped videos on Huawei Nova Y71</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-iphone-12-pro-max-data-from-ios-icloud-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How to Recover iPhone 12 Pro Max Data From iOS iCloud? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-videos-from-your-htc-u23-by-fonelab-android-recover-video/"><u>How to recover old videos from your HTC U23</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-play-hevc-h-265-video-on-13t-by-aiseesoft-video-converter-play-hevc-video-on-android/"><u>How to play HEVC H.265 video on 13T?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-data-from-iphone-11-pro-max-to-other-iphone-13-pro-devices-drfone-by-drfone-transfer-data-from-ios-transfer-data-from-ios/"><u>How To Transfer Data From iPhone 11 Pro Max To Other iPhone 13 Pro devices? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-oneplus-nord-n30-se-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from OnePlus Nord N30 SE to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-music-from-itel-a05s-by-fonelab-android-recover-music/"><u>How to retrieve erased music from Itel A05s</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-music-from-vivo-by-fonelab-android-recover-music/"><u>How to Rescue Lost Music from Vivo</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-techniques-to-transfer-data-from-honor-play-8t-to-iphone-15141312-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Techniques to Transfer Data from Honor Play 8T to iPhone 15/14/13/12 | Dr.fone</u></a></li>
<li><a href="https://activate-lock.techidaily.com/bypass-icloud-activation-lock-with-imei-code-on-iphone-xs-max-by-drfone-ios/"><u>Bypass iCloud Activation Lock with IMEI Code On iPhone XS Max</u></a></li>
<li><a href="https://bypass-frp.techidaily.com/in-2024-top-5-vivo-s18-bypass-frp-tools-for-pc-that-actually-work-by-drfone-android/"><u>In 2024, Top 5 Vivo S18 Bypass FRP Tools for PC That Actually Work</u></a></li>
<li><a href="https://ai-video-apps.techidaily.com/updated-online-stop-motion-animation-studios-top-5-free-and-paid-options-for-2024/"><u>Updated Online Stop Motion Animation Studios Top 5 Free and Paid Options for 2024</u></a></li>
<li><a href="https://animation-videos.techidaily.com/top-10-video-to-gif-converters-top-picks/"><u>Top 10 Video to GIF Converters Top Picks</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/unova-stone-pokemon-go-evolution-list-and-how-catch-them-for-nokia-g310-drfone-by-drfone-virtual-android/"><u>Unova Stone Pokémon Go Evolution List and How Catch Them For Nokia G310 | Dr.fone</u></a></li>
<li><a href="https://apple-account.techidaily.com/in-2024-detailed-guide-on-removing-apple-iphone-x-activation-lock-without-previous-owner-by-drfone-ios/"><u>In 2024, Detailed Guide on Removing Apple iPhone X Activation Lock without Previous Owner?</u></a></li>
<li><a href="https://ios-location-track.techidaily.com/how-to-detect-and-remove-spyware-on-apple-iphone-15-drfone-by-drfone-virtual-ios/"><u>How to Detect and Remove Spyware on Apple iPhone 15? | Dr.fone</u></a></li>
<li><a href="https://ios-pokemon-go.techidaily.com/in-2024-how-and-where-to-find-a-shiny-stone-pokemon-for-apple-iphone-15-pro-drfone-by-drfone-virtual-ios/"><u>In 2024, How and Where to Find a Shiny Stone Pokémon For Apple iPhone 15 Pro? | Dr.fone</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/in-2024-how-to-unlock-nubia-red-magic-9-proplus-bootloader-easily-by-drfone-android/"><u>In 2024, How to Unlock Nubia Red Magic 9 Pro+ Bootloader Easily</u></a></li>
<li><a href="https://ai-voice-clone.techidaily.com/updated-introduction-for-2024/"><u>Updated Introduction for 2024</u></a></li>
<li><a href="https://pokemon-go-android.techidaily.com/how-does-the-stardust-trade-cost-in-pokemon-go-on-honor-x50i-drfone-by-drfone-virtual-android/"><u>How does the stardust trade cost In pokemon go On Honor X50i? | Dr.fone</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/how-to-mirror-samsung-galaxy-s24plus-to-mac-drfone-by-drfone-android/"><u>How to Mirror Samsung Galaxy S24+ to Mac? | Dr.fone</u></a></li>
<li><a href="https://techidaily.com/how-to-easily-hard-reset-my-oppo-a56s-5g-drfone-by-drfone-reset-android-reset-android/"><u>How to Easily Hard reset my Oppo A56s 5G | Dr.fone</u></a></li>
<li><a href="https://bypass-frp.techidaily.com/in-2024-the-updated-method-to-bypass-xiaomi-redmi-note-12t-pro-frp-by-drfone-android/"><u>In 2024, The Updated Method to Bypass Xiaomi Redmi Note 12T Pro FRP</u></a></li>
<li><a href="https://android-transfer.techidaily.com/in-2024-how-can-you-transfer-files-from-xiaomi-redmi-k70e-to-iphone-151413-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>In 2024, How Can You Transfer Files From Xiaomi Redmi K70E To iPhone 15/14/13? | Dr.fone</u></a></li>
<li><a href="https://ai-video-editing.techidaily.com/find-out-how-to-create-your-own-custom-motion-graphics-in-filmora-add-unique-objects-manipulate-text-and-animate-anything-you-want/"><u>Find Out How to Create Your Own Custom Motion Graphics in Filmora. Add Unique Objects, Manipulate Text, and Animate Anything You Want</u></a></li>
</ul></div>



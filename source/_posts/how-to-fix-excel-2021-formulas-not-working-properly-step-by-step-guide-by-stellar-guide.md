---
title: How to Fix Excel 2021 Formulas Not Working Properly | Step-by-Step Guide
date: 2024-05-19T18:32:11.560Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to Fix Excel 2021 Formulas Not Working Properly | Step-by-Step Guide
keywords: repair .xltx,repair excel 2023,repair .xls
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
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-12-mini-without-itunes-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 12 mini without iTunes? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-call-logs-from-itel-a70-by-fonelab-android-recover-call-logs/"><u>How to retrieve erased call logs from Itel A70?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-contacts-from-honor-70-lite-5g-by-fonelab-android-recover-contacts/"><u>How to Rescue Lost Contacts from Honor 70 Lite 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-app-on-itel-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Recover Deleted Photos from Android Gallery App on Itel</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-galaxy-xcover-6-pro-tactical-edition-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Galaxy XCover 6 Pro Tactical Edition?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-call-logs-from-redmi-note-12-pro-4g-by-fonelab-android-recover-call-logs/"><u>How to rescue lost call logs from Redmi Note 12 Pro 4G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-asus-rog-phone-7-ultimate-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Asus ROG Phone 7 Ultimate</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-contacts-from-edge-40-neo-by-fonelab-android-recover-contacts/"><u>How to recover deleted contacts from Edge 40 Neo.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-pictures-from-motorola-moto-g24-by-fonelab-android-recover-pictures/"><u>How to recover deleted pictures from Motorola Moto G24.</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-motorola-moto-g73-5g-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Motorola Moto G73 5G to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-pictures-from-vivo-s17t-by-fonelab-android-recover-pictures/"><u>How to Rescue Lost Pictures from Vivo S17t?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-pictures-from-infinix-note-30i-by-fonelab-android-recover-pictures/"><u>How to Rescue Lost Pictures from Infinix Note 30i?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-photos-files-from-oneplus-ace-2-pro-by-fonelab-android-recover-photos/"><u>How To  Restore Missing Photos Files from OnePlus Ace 2 Pro.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-videos-from-your-motorola-edge-40-neo-by-fonelab-android-recover-video/"><u>How to recover old videos from your Motorola Edge 40 Neo</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-videos-from-your-v29-by-fonelab-android-recover-video/"><u>How to recover old videos from your V29</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-iphone-15-camera-roll-photos-and-photo-stream-pictures-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted iPhone 15 Camera Roll Photos and Photo Stream Pictures? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-system-issues-of-iphone-13-pro-max-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair System Issues of iPhone 13 Pro Max? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-electronically-sign-a-wps-file-using-digisigner-by-ldigisigner-sign-a-word-sign-a-word/"><u>How to Electronically Sign a .wps file Using DigiSigner</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-oppo-find-x7-ultra-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Oppo Find X7 Ultra?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-oneplus-nord-ce-3-5g-pin-by-drfone-android-unlock-android-unlock/"><u>How to remove OnePlus Nord CE 3 5G PIN</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-videos-on-honor-magic-vs-2-by-fonelab-android-recover-video/"><u>How to restore wiped videos on Honor Magic Vs 2</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-rog-phone-7-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on ROG Phone 7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-photos-from-nokia-g310-by-fonelab-android-recover-photos/"><u>How to Rescue Lost Photos from Nokia G310?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-messages-files-from-itel-a05s-by-fonelab-android-recover-messages/"><u>How To  Restore Missing Messages Files from Itel A05s</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-call-history-on-tecno-by-fonelab-android-recover-call-logs/"><u>How to restore wiped call history on Tecno ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-system-of-iphone-12-pro-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair System of iPhone 12 Pro? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-messages-from-samsung-galaxy-m34-by-fonelab-android-recover-messages/"><u>How to Rescue Lost Messages from Samsung Galaxy M34</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-motorola-moto-g14-get-deleted-photos-back-with-ease-and-safety-by-fonelab-android-recover-photos/"><u>How to Motorola Moto G14 Get Deleted photos Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-se-2022-to-factory-settings-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Reset iPhone SE (2022) to Factory Settings? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-samsung-galaxy-a25-5g-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Samsung Galaxy A25 5G to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://location-fake.techidaily.com/5-best-route-generator-apps-you-should-try-on-realme-11-pro-drfone-by-drfone-virtual-android/"><u>5 Best Route Generator Apps You Should Try On Realme 11 Pro | Dr.fone</u></a></li>
<li><a href="https://location-social.techidaily.com/how-to-change-your-samsung-galaxy-s24plus-location-on-life360-without-anyone-knowing-drfone-by-drfone-virtual-android/"><u>How to Change Your Samsung Galaxy S24+ Location on life360 Without Anyone Knowing? | Dr.fone</u></a></li>
<li><a href="https://sim-unlock.techidaily.com/in-2024-top-11-free-apps-to-check-imei-on-vivo-s18-phones-by-drfone-android/"><u>In 2024, Top 11 Free Apps to Check IMEI on Vivo S18 Phones</u></a></li>
<li><a href="https://location-social.techidaily.com/does-find-my-friends-work-on-honor-magic-6-lite-drfone-by-drfone-virtual-android/"><u>Does find my friends work on Honor Magic 6 Lite | Dr.fone</u></a></li>
<li><a href="https://apple-account.techidaily.com/in-2024-how-to-get-the-apple-id-verification-code-from-iphone-12-pro-in-the-best-ways-by-drfone-ios/"><u>In 2024, How To Get the Apple ID Verification Code From iPhone 12 Pro in the Best Ways</u></a></li>
<li><a href="https://android-location-track.techidaily.com/how-to-check-distance-and-radius-on-google-maps-for-your-vivo-y100i-drfone-by-drfone-virtual-android/"><u>How to Check Distance and Radius on Google Maps For your Vivo Y100i | Dr.fone</u></a></li>
<li><a href="https://howto.techidaily.com/authentication-error-occurred-on-asus-rog-phone-8-here-are-10-proven-fixes-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>Authentication Error Occurred on Asus ROG Phone 8? Here Are 10 Proven Fixes | Dr.fone</u></a></li>
<li><a href="https://change-location.techidaily.com/in-2024-what-legendaries-are-in-pokemon-platinum-on-samsung-galaxy-z-fold-5-drfone-by-drfone-virtual-android/"><u>In 2024, What Legendaries Are In Pokemon Platinum On Samsung Galaxy Z Fold 5? | Dr.fone</u></a></li>
<li><a href="https://fake-location.techidaily.com/fixing-foneazy-mockgo-not-working-on-realme-v30-drfone-by-drfone-virtual-android/"><u>Fixing Foneazy MockGo Not Working On Realme V30 | Dr.fone</u></a></li>
<li><a href="https://android-location-track.techidaily.com/5-ways-to-track-nokia-c300-without-app-drfone-by-drfone-virtual-android/"><u>5 Ways to Track Nokia C300 without App | Dr.fone</u></a></li>
<li><a href="https://iphone-unlock.techidaily.com/in-2024-7-top-ways-to-resolve-apple-id-not-active-issue-for-iphone-se-2020-drfone-by-drfone-ios/"><u>In 2024, 7 Top Ways To Resolve Apple ID Not Active Issue For iPhone SE (2020) | Dr.fone</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/in-2024-what-pokemon-evolve-with-a-dawn-stone-for-tecno-pova-5-drfone-by-drfone-virtual-android/"><u>In 2024, What Pokémon Evolve with A Dawn Stone For Tecno Pova 5? | Dr.fone</u></a></li>
<li><a href="https://android-location-track.techidaily.com/in-2024-how-to-spy-on-text-messages-from-computer-and-huawei-nova-y71-drfone-by-drfone-virtual-android/"><u>In 2024, How to Spy on Text Messages from Computer & Huawei Nova Y71 | Dr.fone</u></a></li>
<li><a href="https://pokemon-go-android.techidaily.com/in-2024-a-working-guide-for-pachirisu-pokemon-go-map-on-honor-magic-vs-2-drfone-by-drfone-virtual-android/"><u>In 2024, A Working Guide For Pachirisu Pokemon Go Map On Honor Magic Vs 2 | Dr.fone</u></a></li>
<li><a href="https://ios-pokemon-go.techidaily.com/a-comprehensive-guide-to-mastering-ipogo-for-pokemon-go-on-apple-iphone-13-mini-drfone-by-drfone-virtual-ios/"><u>A Comprehensive Guide to Mastering iPogo for Pokémon GO On Apple iPhone 13 mini | Dr.fone</u></a></li>
<li><a href="https://android-location.techidaily.com/in-2024-3-effective-methods-to-fake-gps-location-on-android-for-your-honor-x50-gt-drfone-by-drfone-virtual/"><u>In 2024, 3 Effective Methods to Fake GPS location on Android For your Honor X50 GT | Dr.fone</u></a></li>
<li><a href="https://android-location.techidaily.com/10-fake-gps-location-apps-on-android-of-your-motorola-moto-g23-drfone-by-drfone-virtual/"><u>10 Fake GPS Location Apps on Android Of your Motorola Moto G23 | Dr.fone</u></a></li>
<li><a href="https://android-transfer.techidaily.com/in-2024-5-techniques-to-transfer-data-from-nokia-g310-to-iphone-15141312-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>In 2024, 5 Techniques to Transfer Data from Nokia G310 to iPhone 15/14/13/12 | Dr.fone</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/best-ways-on-how-to-unlockbypassswiperemove-oneplus-nord-ce-3-5g-fingerprint-lock-by-drfone-android/"><u>Best Ways on How to Unlock/Bypass/Swipe/Remove OnePlus Nord CE 3 5G Fingerprint Lock</u></a></li>
</ul></div>



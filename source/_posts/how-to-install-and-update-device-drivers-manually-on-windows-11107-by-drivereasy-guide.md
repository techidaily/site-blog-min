---
title: How to install and update device drivers manually on Windows 11/10/7
date: 2024-05-19T18:32:13.281Z
tags: 
  - driver
  - device driver
categories: 
  - apps
  - windows
description: This article describes How to install and update device drivers manually on Windows 11/10/7. Device Manager is a control panel applet in Microsoft Windows operating systems. It allows users to view and control the hardware attached to the computer. When a piece of hardware is not working, the offending hardware is highlighted for the user to deal with. The list of hardware can be sorted by various criteria. Device Manager was introduced with Windows 95 and later added to Windows 2000. In NT-based versions, it is included as a Microsoft Management Console snap-in.
keywords: update drivers in Windows 11/10/7,device manager,update drivers in Windows 11 & 10 & 7,update drivers in Windows 10 & 7,update drivers in Windows 11,update drivers in Windows 7
---

## How to Update a Driver with Windows Device Manager

To update a driver in Device Manager (or to see if Windows thinks it needs updating):

- **Step1**: On your keyboard, press the `Windows logo key`  and `R` at the same time to invoke the Run box.
- **Step2**: Type `devmgmt.msc` and click `OK`.

![devmgmt.msc](https://tools.techidaily.com/images/apps/drivereasy/device-manager/1.jpg)  

> (There are other ways to open Device Manager; it changes depending on your version of Windows. But the above method works for all versions of Windows, including Windows 11, 10 and 7.)

- **Step3**: Expand a category (e.g. Display Adapters) to see the devices in that category.
- **Step4**: Right-click the device whose driver you’d like to update, and select `Update Driver`.

![devmgmt.msc](https://tools.techidaily.com/images/apps/drivereasy/device-manager/3.jpg) 

- **Step5**: Select Search automatically for updated driver software.

![devmgmt.msc](https://tools.techidaily.com/images/apps/drivereasy/device-manager/4.jpg) 




## What are drivers?

Drivers are like interpreters between Windows and your devices. For example, when Windows needs to display something on your monitor, it sends a command to your graphics card, and your graphics card then displays what Windows wants on your monitor.

But Windows and your graphics card don’t actually speak the same language. To understand each other, they need a translator. That translator is called a driver. It takes the Windows command and translates it into something your graphics card can understand. Your graphics card can then do as it’s told, and display the right thing on your monitor.

Similarly, if your graphics card needs to send some sort of response back to Windows, the driver translates the response into something Windows can understand.

In this example, what we’re talking about is a video driver, but your computer has many other drivers on it too – one for each device. Your speakers, your printer, your mouse, your USB hard drives, your network card, your keyboard and so on – they each have an associated driver.

And without all these drivers, none of your devices will work.

![What are drivers?](https://tools.techidaily.com/images/apps/drivereasy/pages/why_2.jpg)


## Why you can’t rely on Windows to keep your drivers up-to-date

Windows comes with an inbuilt tool, called ‘Windows Update’, that’s supposed to automatically keep your drivers up to date. Unfortunately, it doesn’t work very well.

There are two reasons why…

- Device manufacturers often take a long time to get their drivers into a Windows Update – It’s a time-consuming and difficult process. Sometimes they just miss the deadline and have to wait ‘til the next Windows Update, and sometimes they just give up altogether. In fact, for older devices, this is the norm.

- Windows Update ignores driver updates it considers ‘optional’ – It categorizes driver updates as either ‘critical’, ‘automatic’ or ‘optional’, and it doesn’t usually concern itself with the ‘optional’ ones – even when they’re actually important. You can install them manually by going to the ‘Optional updates’ screen but, even then, as described above, you’re unlikely to get all the latest drivers.

## All our drivers are certified

We use only genuine drivers, straight from your hardware manufacturer. And we employ a strict testing process to ensure they’re safe, stable, robust, up-to-date, and compatible with Windows and all the most popular combinations of hardware and software.

### Microsoft WHQL Testing

Most hardware manufacturers put their drivers through Microsoft’s rigorous Windows Hardware Quality Labs (WHQL) testing process. If they pass, they’re officially certified stable and compatible with Windows.

If your manufacturer has a ‘Certified for Windows’ driver, that’s the one we’ll use. For Windows 10 and 11, Driver Easy installs only drivers that are ‘Certified for Windows’ through the Windows Hardware Quality Labs (WHQL) program. For Windows 7, 8 and Vista, Driver Easy installs WHQL drivers by default, if they’re available (which they are for 95.69% of drivers for those versions of Windows), but also gives users the option to install non-WHQL drivers.

But we don’t stop there. We also perform our own tests to ensure the stability of our drivers…

### Certified by Driver Easy

We employ a strict testing regime to ensure our drivers are safe, secure and stable.

This is critical because not all manufacturers get their drivers certified by Microsoft – particularly for older hardware. (It’s a very rigorous and time-consuming process, and for manufacturers with a lot of devices and drivers, it can become quite expensive.)

## We test on all the most popular combinations of hardware & software

Our tests are a lot more hands-on and practical than Microsoft’s tests. Because drivers behave differently on different computers, different versions of Windows, and even in the presence of different software applications, the only way to really tell if a driver will be stable, compatible and safe for everyone is to physically test it on all the popular hardware / operating system / software combinations. So that’s what we do:

- We test on hundreds of PCs – Our testing facility is strategically located in Shenzhen, China, one of the country’s biggest IT hubs. We specifically selected this estate because we’re surrounded by hundreds of PC distributors, all within walking distance. This means we have unfettered access to an almost limitless supply of hardware, and can physically test our drivers on all the most popular computers – including the latest new models available on the market, as well as second-hand computers that still have a wide user base. 
- We test with physical devices attached – For external device drivers (e.g. for printers, external hard drives, mice, keyboards), we physically install the external device to test the driver.
- We test on all current versions of Windows – On each test PC, we install and test thoroughly on Windows 11 32-bit, Windows 11 64-bit, Windows 10 32 bit, Windows 10 64-bit, Windows 7 32-bit and Windows 7 64-bit.
- We test with popular programs installed – On each installation of Windows, we also install a variety of popular software programs before testing the driver (e.g. various versions of Microsoft Office, antivirus products and video players).

## Here’s our full testing process

We subject all new drivers to a battery of tests.

### Step 1. Filter out faulty drivers

First, we locate and download any new drivers from nearly 100 manufacturer websites, then scan them all with two proprietary tools that filter out any that:

- are incorrectly formatted;
- are missing files;
- are likely to be flagged by antivirus programs; or
- have failed our previous tests.
Then we add all drivers that pass these filters to our development-only version of Driver Easy.

### Step 2. Test on all modern versions of Windows

We then scan a small selection of computers with our development-only Driver Easy. These computers have typical devices attached, like a mouse, keyboard, monitor and printer. On each computer, we test all modern versions of Windows (Windows 11 32-bit, Windows 11 64-bit, Windows 10 32 bit, Windows 10 64-bit, Windows 7 32-bit and Windows 7 64-bit):

- **01.** We install each driver that Driver Easy recommends, one at a time.

- **02.** After each driver installation, we check that the computer functions normally and all devices work properly. E.g. If it’s a network card driver, we test the internet connection, if it’s a video card driver, we test the screen resolution, if it’s a keyboard driver, we test that the keyboard is functioning properly, and so on.

- **03.** We then check Windows’ Device Manager to ensure no device drivers are flagged as problematic.

- **04.** We then restart the computer to ensure that the driver installation didn’t cause any issues with Windows (e.g. no blue screen of death on startup, no error messages, no unexpected behavior).

- **05.** If all is working as expected, we return to step 1, and install and test the next driver.

- **06.** If there are issues, we check the driver install log to see if any errors were detected during installation.

- **07.** If the log is inconclusive, we do further testing to determine if it was the driver that caused the issue. Usually we test an alternative driver to see if it causes the same issue. If it doesn’t, then it’s likely the first driver is the culprit. If the same issue occurs with the alternative driver too, we test to see if the computer itself is the issue. Often this involves performing a system restore on the test PC.

- **08.** If we can prove that the driver was the cause of the computer or device issue, we remove it from Driver Easy, then return to step 1, and install and test the next driver.

Any drivers that make it through our first two test phases are then added to the live Driver Easy database.

### Step 3. Test on many popular computers

We then use Driver Easy to scan dozens of the most popular computer setups (PC, operating system, video card, sound card, network card, printer, default software, etc.):

- 01. We install each driver that Driver Easy recommends, one at a time.

- 02. After each driver installation, we check that the computer functions normally and all devices work properly. E.g. If it’s a network card driver, we test the internet connection, if it’s a video card driver, we test the screen resolution, if it’s a keyboard driver, we test that the keyboard is functioning properly, and so on.

- 03. We then check Windows’ Device Manager to ensure no device drivers are flagged as problematic.

- 04. We then restart the computer to ensure that the driver installation didn’t cause any issues with Windows (e.g. no blue screen of death on startup, no error messages, no unexpected behavior).

- 05. If all is working as expected, we return to step 1, and install and test the next driver.

- 06. If there are issues, we check the driver install log to see if any errors were detected during installation.

- 07. If the log is inconclusive, we do further testing to determine if it was the driver that caused the issue. Usually we test an alternative driver to see if it causes the same issue. If it doesn’t, then it’s likely the first driver is the culprit. If the same issue occurs with the alternative driver too, we test to see if the computer itself is the issue. Often this involves performing a system restore on the test PC.

- 08. If we can prove that the driver was the cause of the computer or device issue, we remove it from Driver Easy, then return to step 1, and install and test the next driver.

Over the course of a year, we test on hundreds of different computers in this way.

If a driver fails our tests…
If we establish that a manufacturer’s driver causes issues on any combination of hardware, operating system and software, we find an alternative version of the driver for that particular combination.

For example, if an audio driver supplied by Dell for a certain laptop causes issues on Windows 10, we’d source a different version of it. Typically from the audio card’s chipset manufacturer (e.g. Realtek). They’d usually have the most up-to-date drivers available because they continue updating their drivers almost indefinitely, whereas Dell would typically stop updating the laptop’s drivers as soon as it’s superseded by a newer model.

Once we’ve located an alternative driver, we start over at step 1 of our testing process with it.

## Why update your drivers in Windows 11, 10 & Windows 7?

Many computer issues are caused by outdated device drivers. Particularly in Windows 10/11.

So if your computer has slowed down, you should update your drivers. If it’s crashing or hanging, update your drivers. If you can’t connect to the internet, update your drivers. If your mouse, keyboard, monitor or speakers are acting up, update your drivers. In fact, no matter what your issue, there’s a good chance updating your drivers will fix it.

To understand why, you first have to understand what drivers actually are…




## Try Driver Easy for free

If you want the certainty of knowing your device drivers are always up to date (and not just sometimes up to date, which is all you get from Windows Device Manager), and you don’t have the time, patience or computer skills to continually update them manually, give the free version of [Driver Easy](https://tools.techidaily.com/drivereasy/download/) a try.

[Download Free Version](https://tools.techidaily.com/drivereasy/download/)

The free version will identify all your outdated drivers, and allow you to download them all. But only one at a time and, once they’re downloaded, you have to manually install them using the standard Windows process. (To automatically update all your drivers with 1 click, you’ll need [the Pro version of Driver Easy](https://tools.techidaily.com/drivereasy/download/). Don’t worry, it comes with a 30-day, no-questions-asked, money back satisfaction guarantee.)

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
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-nokia-g22-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Nokia G22 to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-contacts-from-tecno-spark-10-pro-by-fonelab-android-recover-contacts/"><u>How to Rescue Lost Contacts from Tecno Spark 10 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-messages-from-infinix-zero-30-5g-by-fonelab-android-recover-messages/"><u>How to Rescue Lost Messages from Infinix Zero 30 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-google-frp-lock-on-x-fold-2-by-drfone-android-unlock-remove-google-frp/"><u>How to remove Google FRP Lock on X Fold 2</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-music-from-sony-xperia-1-v-by-fonelab-android-recover-music/"><u>How to retrieve erased music from Sony Xperia 1 V</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-without-backup-on-realme-gt-5-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery without backup on Realme GT 5</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-infinix-note-30-pro-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Infinix Note 30 Pro to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-install-device-drivers-manually-in-windows-7-by-drivereasy-guide/"><u>How to install device drivers manually in Windows 7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-google-frp-lock-on-21-by-drfone-android-unlock-remove-google-frp/"><u>How to remove Google FRP Lock on 21</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-install-the-latest-ios-beta-version-on-iphone-8-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Install the Latest iOS Beta Version on iPhone 8? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-motorola-moto-g04-with-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Motorola Moto G04 with Video Repair Utility on Windows? </u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-videos-on-realme-11-pro-by-fonelab-android-recover-video/"><u>How to restore wiped videos on Realme 11 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-contacts-from-google-by-fonelab-android-recover-contacts/"><u>How to recover deleted contacts from Google .</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-tecno-pova-5-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Tecno Pova 5</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-messages-from-your-htc-u23-by-fonelab-android-recover-messages/"><u>How to recover old messages from your HTC U23</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-se-2020-to-factory-settings-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Reset iPhone SE (2020) to Factory Settings? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-photos-files-from-samsung-galaxy-s24plus-by-fonelab-android-recover-photos/"><u>How To  Restore Missing Photos Files from Samsung Galaxy S24+.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-app-on-samsung-galaxy-m34-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to Recover Deleted Photos from Android Gallery App on Samsung Galaxy M34 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-play-mov-files-on-motorola-moto-g24-by-aiseesoft-video-converter-play-mov-on-android/"><u>How to play MOV files on Motorola Moto G24 ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-play-an-mp4-on-motorola-moto-g24-by-aiseesoft-video-converter-play-mp4-on-android/"><u>How to play an MP4 on Motorola Moto G24?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-the-unable-to-record-macro-error-in-excel-2000-by-stellar-guide/"><u>How to Fix the Unable to Record Macro Error in Excel 2000?</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/in-2024-overview-of-the-best-samsung-galaxy-m34-5g-screen-mirroring-app-drfone-by-drfone-android/"><u>In 2024, Overview of the Best Samsung Galaxy M34 5G Screen Mirroring App | Dr.fone</u></a></li>
<li><a href="https://fake-location.techidaily.com/ispoofer-is-not-working-on-xiaomi-redmi-k70-pro-fixed-drfone-by-drfone-virtual-android/"><u>iSpoofer is not working On Xiaomi Redmi K70 Pro? Fixed | Dr.fone</u></a></li>
<li><a href="https://ai-voice-clone.techidaily.com/in-2024-how-to-translate-video-from-japanese-to-english-online/"><u>In 2024, How To Translate Video From Japanese to English Online?</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/in-2024-3-methods-to-mirror-xiaomi-redmi-note-12r-to-roku-drfone-by-drfone-android/"><u>In 2024, 3 Methods to Mirror Xiaomi Redmi Note 12R to Roku | Dr.fone</u></a></li>
<li><a href="https://fake-location.techidaily.com/complete-tutorial-to-use-vpna-to-fake-gps-location-on-xiaomi-redmi-note-12-5g-drfone-by-drfone-virtual-android/"><u>Complete Tutorial to Use VPNa to Fake GPS Location On Xiaomi Redmi Note 12 5G | Dr.fone</u></a></li>
<li><a href="https://ai-topics.techidaily.com/updated-what-is-an-ai-art-generator/"><u>Updated What Is an AI Art Generator?</u></a></li>
<li><a href="https://sim-unlock.techidaily.com/how-to-unlock-sim-card-on-vivo-g2-online-without-jailbreak-by-drfone-android/"><u>How to Unlock SIM Card on Vivo G2 online without jailbreak</u></a></li>
<li><a href="https://ai-video-editing.techidaily.com/new-2024-approved-fun-and-easy-masking-tricks-wondershare-filmora-tutorial/"><u>New 2024 Approved Fun & Easy Masking Tricks Wondershare Filmora Tutorial</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/in-2024-how-to-screen-share-mac-to-apple-iphone-6s-drfone-by-drfone-ios/"><u>In 2024, How to Screen Share Mac to Apple iPhone 6s? | Dr.fone</u></a></li>
<li><a href="https://sim-unlock.techidaily.com/in-2024-the-6-best-sim-unlock-services-that-actually-work-on-your-oppo-a38-device-by-drfone-android/"><u>In 2024, The 6 Best SIM Unlock Services That Actually Work On Your Oppo A38 Device</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/in-2024-how-to-stream-anything-from-lava-yuva-2-pro-to-apple-tv-drfone-by-drfone-android/"><u>In 2024, How To Stream Anything From Lava Yuva 2 Pro to Apple TV | Dr.fone</u></a></li>
<li><a href="https://howto.techidaily.com/authentication-error-occurred-on-vivo-y100a-here-are-10-proven-fixes-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>Authentication Error Occurred on Vivo Y100A? Here Are 10 Proven Fixes | Dr.fone</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/how-can-i-catch-the-regional-pokemon-without-traveling-on-meizu-21-drfone-by-drfone-virtual-android/"><u>How Can I Catch the Regional Pokémon without Traveling On Meizu 21 | Dr.fone</u></a></li>
<li><a href="https://ai-video-apps.techidaily.com/in-2024-unleash-the-power-of-intros-top-10-movie-intro-making-tools/"><u>In 2024, Unleash the Power of Intros Top 10 Movie Intro Making Tools</u></a></li>
<li><a href="https://ios-unlock.techidaily.com/in-2024-how-to-change-your-apple-id-on-iphone-13-pro-max-with-or-without-password-by-drfone-ios/"><u>In 2024, How To Change Your Apple ID on iPhone 13 Pro Max With or Without Password</u></a></li>
<li><a href="https://ai-video-editing.techidaily.com/new-2024-approved-video-editing-tips-to-know-before-editing-first-draft/"><u>New 2024 Approved Video Editing Tips to Know Before Editing First Draft</u></a></li>
<li><a href="https://apple-account.techidaily.com/how-to-delete-icloud-account-from-apple-iphone-14-without-password-by-drfone-ios/"><u>How to Delete iCloud Account From Apple iPhone 14 without Password?</u></a></li>
<li><a href="https://sim-unlock.techidaily.com/android-unlock-code-sim-unlock-your-itel-p55-5g-phone-and-remove-locked-screen-by-drfone-android/"><u>Android Unlock Code Sim Unlock Your Itel P55 5G Phone and Remove Locked Screen</u></a></li>
<li><a href="https://fix-guide.techidaily.com/how-to-flash-dead-honor-magic-5-lite-safely-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>How to Flash Dead Honor Magic 5 Lite Safely | Dr.fone</u></a></li>
<li><a href="https://animation-videos.techidaily.com/2024-approved-what-is-the-meaning-of-motion-graphics/"><u>2024 Approved What Is the Meaning of Motion Graphics</u></a></li>
<li><a href="https://android-unlock.techidaily.com/5-solutions-for-vivo-x-fold-2-unlock-without-password-by-drfone-android/"><u>5 Solutions For Vivo X Fold 2 Unlock Without Password</u></a></li>
</ul></div>



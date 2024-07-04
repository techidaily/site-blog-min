---
title: How to install and update device drivers manually on Windows 10
date: 2024-05-19T18:32:13.277Z
tags: 
  - driver
  - device driver
categories: 
  - apps
  - windows
description: This article describes How to install and update device drivers manually on Windows 10. Device Manager is a control panel applet in Microsoft Windows operating systems. It allows users to view and control the hardware attached to the computer. When a piece of hardware is not working, the offending hardware is highlighted for the user to deal with. The list of hardware can be sorted by various criteria. Device Manager was introduced with Windows 95 and later added to Windows 2000. In NT-based versions, it is included as a Microsoft Management Console snap-in.
keywords: update drivers in Windows 10,update drivers in Windows 11/10/7,update drivers in Windows 7,update drivers in Windows 11/10
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






## What happens if a driver is missing or outdated?

Every now and then, Microsoft will change the commands Windows sends to one of your devices (e.g. your network card). When this happens, the manufacturer of that device need to change the device driver too. They need to teach it the new Windows commands. Otherwise the drivers won’t be able to translate those commands for your devices, and your devices won’t work properly.

The same thing needs to happen when your device manufacturer changes the way your device talks, or the things it can do. They need to change the driver too. Otherwise Windows won’t be able to talk to the device, or take advantage of its new functionality, and your device won’t work properly.

Now when we say “your device won’t work properly”, sometimes this means simply that you miss out on new functionality or minor bug fixes. But it’s often a lot more serious than that. Your computer may even hang, crash or stop working completely. Remember, there’s a driver that controls your hard drive, for instance. If Windows can’t talk to your hard drive, it can’t access any of the data on your drive. Similarly, if Windows can’t talk to your network card, you won’t be able to access the internet, and if it can’t talk to your graphics card, you won’t be able to see anything on your monitor. These are just a few of the more serious issues outdated drivers can cause.

![What happens if a driver is missing or outdated?](https://tools.techidaily.com/images/apps/drivereasy/pages/why_3.jpg)

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

## What are drivers?

Drivers are like interpreters between Windows and your devices. For example, when Windows needs to display something on your monitor, it sends a command to your graphics card, and your graphics card then displays what Windows wants on your monitor.

But Windows and your graphics card don’t actually speak the same language. To understand each other, they need a translator. That translator is called a driver. It takes the Windows command and translates it into something your graphics card can understand. Your graphics card can then do as it’s told, and display the right thing on your monitor.

Similarly, if your graphics card needs to send some sort of response back to Windows, the driver translates the response into something Windows can understand.

In this example, what we’re talking about is a video driver, but your computer has many other drivers on it too – one for each device. Your speakers, your printer, your mouse, your USB hard drives, your network card, your keyboard and so on – they each have an associated driver.

And without all these drivers, none of your devices will work.

![What are drivers?](https://tools.techidaily.com/images/apps/drivereasy/pages/why_2.jpg)

## Download Driver Easy FREE

If you’re having computer issues, the first thing you should do is see if it has outdated drivers. If it does, updating them will very often fix things.

And for this, our tool, Driver Easy FREE, is the ideal solution.

Driver Easy FREE is a driver update tool used by more than 3 million customers around the world. It will automatically identify and download all the drivers you need, so all you have to do is install them.

In other words, it eliminates the need to find and download your drivers the difficult way (via the manufacturer’s website).

You don’t need to know what system your computer is running, you don’t need to scour the web for the right driver download, and you don’t need to risk downloading the wrong driver. Driver Easy does it all for you, automatically. No computer knowledge needed, and it’s completely free.

This page describes how to do this.

[Download Free Version](https://tools.techidaily.com/drivereasy/download/)


The free version will identify all your outdated drivers, and allow you to download them all. But only one at a time and, once they’re downloaded, you have to manually install them using the standard Windows process. (To automatically update all your drivers with 1 click, you’ll need [the Pro version of Driver Easy](https://tools.techidaily.com/drivereasy/download/). Don’t worry, it comes with a 30-day, no-questions-asked, money back satisfaction guarantee.)


## Why you can’t rely on Windows to keep your drivers up-to-date

Windows comes with an inbuilt tool, called ‘Windows Update’, that’s supposed to automatically keep your drivers up to date. Unfortunately, it doesn’t work very well.

There are two reasons why…

- Device manufacturers often take a long time to get their drivers into a Windows Update – It’s a time-consuming and difficult process. Sometimes they just miss the deadline and have to wait ‘til the next Windows Update, and sometimes they just give up altogether. In fact, for older devices, this is the norm.

- Windows Update ignores driver updates it considers ‘optional’ – It categorizes driver updates as either ‘critical’, ‘automatic’ or ‘optional’, and it doesn’t usually concern itself with the ‘optional’ ones – even when they’re actually important. You can install them manually by going to the ‘Optional updates’ screen but, even then, as described above, you’re unlikely to get all the latest drivers.

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
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-motorola-moto-g14-photos-an-easy-method-explained-by-fonelab-android-recover-photos/"><u>How to Restore Deleted Motorola Moto G14 Photos  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-galaxy-a34-5g-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Galaxy A34 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-google-frp-lock-on-oppo-reno-10-5g-by-drfone-android-unlock-remove-google-frp/"><u>How to remove Google FRP Lock on Oppo Reno 10 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-iphone-15-plus-without-backup-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Restore iPhone 15 Plus without Backup | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-pictures-from-xiaomi-redmi-k70-pro-by-fonelab-android-recover-pictures/"><u>How to recover deleted pictures from Xiaomi Redmi K70 Pro.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-share-your-winning-forex-trades-with-friends-and-family-on-local-trade-copier-tm-together-by-mt4copier-guide/"><u>How to Share Your Winning Forex Trades With Friends and Family on Local Trade Copier™ Together</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-call-logs-from-nokia-150-2023-by-fonelab-android-recover-call-logs/"><u>How to retrieve erased call logs from Nokia 150 (2023)?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-pictures-files-from-poco-by-fonelab-android-recover-pictures/"><u>How To  Restore Missing Pictures Files from Poco .</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-vivo-y100a-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Vivo Y100A to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-messages-files-from-find-x6-by-fonelab-android-recover-messages/"><u>How To  Restore Missing Messages Files from Find X6</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-files-after-iphone-7-factory-reset-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Files after iPhone 7 Factory Reset? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-call-logs-from-redmi-note-12-pro-4g-by-fonelab-android-recover-call-logs/"><u>How to rescue lost call logs from Redmi Note 12 Pro 4G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-pictures-from-itel-s23plus-by-fonelab-android-recover-pictures/"><u>How to recover deleted pictures from Itel S23+.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-contacts-on-iphone-se-4-methods-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Restore Contacts on iPhone SE (4 Methods) | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-back-lost-photos-from-honor-magic-vs-2-by-fonelab-android-recover-photos/"><u>How to get back lost photos from Honor Magic Vs 2.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-messages-files-from-huawei-nova-y91-by-fonelab-android-recover-messages/"><u>How To  Restore Missing Messages Files from Huawei Nova Y91</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-solve-mkv-lagging-problem-in-huawei-nova-y91-by-aiseesoft-video-converter-play-mkv-on-android/"><u>How to solve MKV lagging problem in Huawei Nova Y91?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-excel-2016-has-encountered-a-problem-stellar-by-stellar-guide/"><u>How to Fix Excel 2016 has Encountered a Problem | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-music-from-honor-by-fonelab-android-recover-music/"><u>How to retrieve erased music from Honor</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupted-pdf-v15-file-when-manual-method-fails-by-stellar-guide/"><u>How to repair corrupted PDF v1.5 file when manual method fails</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-lost-files-from-nubia-red-magic-8s-pro-by-fonelab-android-recover-data/"><u>How to retrieve lost files from Nubia Red Magic 8S Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-messages-files-from-asus-rog-phone-7-ultimate-by-fonelab-android-recover-messages/"><u>How To  Restore Missing Messages Files from Asus ROG Phone 7 Ultimate</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-iphone-se-2020-storage-not-loadingshowing-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Fix iPhone SE (2020) Storage Not Loading/Showing | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-pictures-from-oppo-a59-5g-by-fonelab-android-recover-pictures/"><u>How to recover deleted pictures from Oppo A59 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-missing-your-hardware-drivers-with-windows-device-manager-on-windows-11-and-10-and-7-by-drivereasy-guide/"><u>How to identify missing your hardware drivers with Windows Device Manager on Windows 11 & 10 & 7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-vivo-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Vivo ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-realme-11-pro-get-deleted-photos-back-with-ease-and-safety-by-fonelab-android-recover-photos/"><u>How to Realme 11 Pro Get Deleted photos Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-lost-files-from-oppo-by-fonelab-android-recover-data/"><u>How to retrieve lost files from Oppo ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-motorola-defy-2-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Motorola Defy 2 to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-videos-from-iphone-xs-max-without-backup-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Videos from iPhone XS Max Without Backup? | Stellar</u></a></li>
<li><a href="https://ai-voice-clone.techidaily.com/updated-2024-approved-how-to-translate-tiktok-videos-all-you-want-to-know/"><u>Updated 2024 Approved How to Translate TikTok Videos? All You Want to Know</u></a></li>
<li><a href="https://iphone-unlock.techidaily.com/in-2024-how-to-unlock-iphone-13-mini-with-an-apple-watch-and-what-to-do-if-it-doesnt-work-drfone-by-drfone-ios/"><u>In 2024, How to Unlock iPhone 13 mini With an Apple Watch & What to Do if It Doesnt Work | Dr.fone</u></a></li>
<li><a href="https://ai-vdieo-software.techidaily.com/new-avi-video-rotation-made-easy-top-5-free-solutions/"><u>New AVI Video Rotation Made Easy Top 5 Free Solutions</u></a></li>
<li><a href="https://animation-videos.techidaily.com/updated-in-2024-10-best-photoshop-cartoon-effects-for-creatives/"><u>Updated In 2024, 10 Best Photoshop Cartoon Effects For Creatives</u></a></li>
<li><a href="https://activate-lock.techidaily.com/in-2024-a-comprehensive-guide-to-icloud-unlock-from-apple-iphone-15-plus-online-by-drfone-ios/"><u>In 2024, A Comprehensive Guide to iCloud Unlock From Apple iPhone 15 Plus Online</u></a></li>
<li><a href="https://location-social.techidaily.com/why-your-whatsapp-location-is-not-updating-and-how-to-fix-on-poco-c51-drfone-by-drfone-virtual-android/"><u>Why Your WhatsApp Location is Not Updating and How to Fix On Poco C51 | Dr.fone</u></a></li>
<li><a href="https://sim-unlock.techidaily.com/in-2024-android-unlock-code-sim-unlock-your-vivo-y28-5g-phone-and-remove-locked-screen-by-drfone-android/"><u>In 2024, Android Unlock Code Sim Unlock Your Vivo Y28 5G Phone and Remove Locked Screen</u></a></li>
<li><a href="https://techidaily.com/how-to-transfer-whatsapp-from-apple-iphone-15-to-other-iphone-13-pro-max-devices-drfone-by-drfone-transfer-whatsapp-from-ios-transfer-whatsapp-from-ios/"><u>How To Transfer WhatsApp From Apple iPhone 15 to other iPhone 13 Pro Max devices? | Dr.fone</u></a></li>
<li><a href="https://fix-guide.techidaily.com/in-2024-4-feasible-ways-to-fake-location-on-facebook-for-your-tecno-spark-10-5g-drfone-by-drfone-virtual-android/"><u>In 2024, 4 Feasible Ways to Fake Location on Facebook For your Tecno Spark 10 5G | Dr.fone</u></a></li>
<li><a href="https://ai-voice-clone.techidaily.com/updated-ai-video-translator-online-for-2024/"><u>Updated AI Video Translator | Online for 2024</u></a></li>
<li><a href="https://activate-lock.techidaily.com/in-2024-4-things-you-must-know-about-apple-iphone-6-activation-lock-by-drfone-ios/"><u>In 2024, 4 Things You Must Know About Apple iPhone 6 Activation Lock</u></a></li>
<li><a href="https://ios-pokemon-go.techidaily.com/in-2024-here-are-some-pro-tips-for-pokemon-go-pvp-battles-on-apple-iphone-15-pro-drfone-by-drfone-virtual-ios/"><u>In 2024, Here are Some Pro Tips for Pokemon Go PvP Battles On Apple iPhone 15 Pro | Dr.fone</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/the-best-ispoofer-alternative-to-try-on-asus-rog-phone-8-pro-drfone-by-drfone-virtual-android/"><u>The Best iSpoofer Alternative to Try On Asus ROG Phone 8 Pro | Dr.fone</u></a></li>
<li><a href="https://pokemon-go-android.techidaily.com/the-most-useful-tips-for-pokemon-go-ultra-league-on-honor-v-purse-drfone-by-drfone-virtual-android/"><u>The Most Useful Tips for Pokemon Go Ultra League On Honor V Purse | Dr.fone</u></a></li>
</ul></div>



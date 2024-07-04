---
title: How to identify missing or malfunctioning hardware drivers with Windows Device Manager on Windows 11 & 10
date: 2024-05-19T18:32:12.937Z
tags: 
  - driver
  - device driver
categories: 
  - apps
  - windows
description: This article describes How to identify missing or malfunctioning hardware drivers with Windows Device Manager on Windows 11 & 10. Device Manager is a control panel applet in Microsoft Windows operating systems. It allows users to view and control the hardware attached to the computer. When a piece of hardware is not working, the offending hardware is highlighted for the user to deal with. The list of hardware can be sorted by various criteria. Device Manager was introduced with Windows 95 and later added to Windows 2000. In NT-based versions, it is included as a Microsoft Management Console snap-in.
keywords: identify missing drivers,identify malfunctioning drivers in Windows 10,identify missing drivers in Windows 10 & 7,identify malfunctioning drivers in Windows 11 & 10
---


## How to identify missing or malfunctioning drivers with Windows Device Manager

To see which of your devices have a missing or malfunctioning driver:

- **Step1**: On your keyboard, press the `Windows logo key`  and `R` at the same time to invoke the Run box.
- **Step2**: Type `devmgmt.msc` and click `OK`.
  
![devmgmt.msc](https://tools.techidaily.com/images/apps/drivereasy/device-manager/1.jpg) 
> (There are other ways to open Device Manager; it changes depending on your version of Windows. But the above method works for all versions of Windows, including Windows 11, 10 and 7.)
>

- **Step3**: Expand a category (e.g. Display Adapters) to see the devices in that category. If you see a yellow triangle or question mark next to a device, Windows has detected that it has a missing or malfunctioning driver.

![Device Manager](https://tools.techidaily.com/images/apps/drivereasy/device-manager/2.jpg)

- **Step4**: If you see this yellow mark, you can try to `update` or `reinstall` the driver.




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

## Why update your drivers in Windows 11, 10 & Windows 7?

Many computer issues are caused by outdated device drivers. Particularly in Windows 10/11.

So if your computer has slowed down, you should update your drivers. If it’s crashing or hanging, update your drivers. If you can’t connect to the internet, update your drivers. If your mouse, keyboard, monitor or speakers are acting up, update your drivers. In fact, no matter what your issue, there’s a good chance updating your drivers will fix it.

To understand why, you first have to understand what drivers actually are…


## What are drivers?

Drivers are like interpreters between Windows and your devices. For example, when Windows needs to display something on your monitor, it sends a command to your graphics card, and your graphics card then displays what Windows wants on your monitor.

But Windows and your graphics card don’t actually speak the same language. To understand each other, they need a translator. That translator is called a driver. It takes the Windows command and translates it into something your graphics card can understand. Your graphics card can then do as it’s told, and display the right thing on your monitor.

Similarly, if your graphics card needs to send some sort of response back to Windows, the driver translates the response into something Windows can understand.

In this example, what we’re talking about is a video driver, but your computer has many other drivers on it too – one for each device. Your speakers, your printer, your mouse, your USB hard drives, your network card, your keyboard and so on – they each have an associated driver.

And without all these drivers, none of your devices will work.

![What are drivers?](https://tools.techidaily.com/images/apps/drivereasy/pages/why_2.jpg)

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
<li><a href="https://blog-min.techidaily.com/how-to-identify-malfunctioning-hardware-drivers-with-windows-device-manager-in-windows-10-by-drivereasy-guide/"><u>How to identify malfunctioning hardware drivers with Windows Device Manager in Windows 10</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-infinix-note-30-vip-by-fonelab-android-recover-photos/"><u>How to recover deleted photos from Infinix Note 30 VIP.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-nubia-red-magic-9-pro-contacts-an-easy-method-explained-by-fonelab-android-recover-contacts/"><u>How to Restore Deleted Nubia Red Magic 9 Pro Contacts  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-out-of-recovery-or-dfu-mode-on-iphone-6s-plus-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Get Out of Recovery or DFU Mode on iPhone 6s Plus? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-messages-on-realme-c55-by-fonelab-android-recover-messages/"><u>How to restore wiped messages on Realme C55</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-photos-from-itel-a70-by-fonelab-android-recover-photos/"><u>How to Rescue Lost Photos from Itel A70?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-a-damaged-video-file-of-huawei-huawei-mate-60-proplus-using-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair a Damaged video file of Huawei Huawei Mate 60 Pro+ using Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-contacts-files-from-sony-xperia-5-v-by-fonelab-android-recover-contacts/"><u>How To  Restore Missing Contacts Files from Sony Xperia 5 V.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-videos-from-oneplus-12r-by-fonelab-android-recover-video/"><u>How to Rescue Lost Videos from OnePlus 12R</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-realme-11-pro-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Realme 11 Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-get-out-of-dfu-mode-on-iphone-xs-max-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Get Out of DFU Mode on iPhone XS Max? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-iphone-13-ios-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Repair iPhone 13 iOS? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-call-logs-from-y100a-by-fonelab-android-recover-call-logs/"><u>How To  Restore Missing Call Logs from Y100A</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-photos-files-from-itel-a70-by-fonelab-android-recover-photos/"><u>How To  Restore Missing Photos Files from Itel A70.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-call-logs-from-s17-by-fonelab-android-recover-call-logs/"><u>How To  Restore Missing Call Logs from S17</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-nokia-c110-pin-by-drfone-android-unlock-android-unlock/"><u>How to remove Nokia C110 PIN</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-play-avchd-mts-files-on-g54-5g-by-aiseesoft-video-converter-play-mts-on-android/"><u>How to play AVCHD MTS files on G54 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-call-logs-from-tecno-by-fonelab-android-recover-call-logs/"><u>How To  Restore Missing Call Logs from Tecno</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-google-frp-lock-on-realme-c67-4g-by-drfone-android-unlock-remove-google-frp/"><u>How to remove Google FRP Lock on Realme C67 4G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-x-to-an-older-version-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone X to an Older Version? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-a-damaged-video-file-of-realme-gt-5-pro-using-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair a Damaged video file of Realme GT 5 Pro using Video Repair Utility on Windows?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-call-logs-from-x6-pro-by-fonelab-android-recover-call-logs/"><u>How to retrieve erased call logs from X6 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-install-and-update-device-drivers-manually-on-windows-11-and-10-by-drivereasy-guide/"><u>How to install and update device drivers manually on Windows 11 & 10</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-call-logs-from-your-honor-magic-6-pro-by-fonelab-android-recover-call-logs/"><u>How to recover old call logs from your Honor Magic 6 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-play-avchd-mts-files-on-xiaomi-redmi-12-by-aiseesoft-video-converter-play-mts-on-android/"><u>How to play AVCHD MTS files on Xiaomi Redmi 12?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-huawei-nova-12-pro-using-video-repair-utility-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Huawei Nova 12 Pro using Video Repair Utility on Windows? </u></a></li>
<li><a href="https://location-social.techidaily.com/how-to-change-oneplus-12r-location-on-skout-drfone-by-drfone-virtual-android/"><u>How to Change OnePlus 12R Location on Skout | Dr.fone</u></a></li>
<li><a href="https://location-fake.techidaily.com/6-ways-to-change-spotify-location-on-your-xiaomi-13-ultra-drfone-by-drfone-virtual-android/"><u>6 Ways to Change Spotify Location On Your Xiaomi 13 Ultra | Dr.fone</u></a></li>
<li><a href="https://change-location.techidaily.com/here-are-some-reliable-ways-to-get-pokemon-go-friend-codes-for-vivo-t2-5g-drfone-by-drfone-virtual-android/"><u>Here Are Some Reliable Ways to Get Pokemon Go Friend Codes For Vivo T2 5G | Dr.fone</u></a></li>
<li><a href="https://fake-location.techidaily.com/how-to-change-netflix-location-to-get-more-country-version-on-oppo-find-x7-drfone-by-drfone-virtual-android/"><u>How to Change Netflix Location to Get More Country Version On Oppo Find X7 | Dr.fone</u></a></li>
<li><a href="https://iphone-transfer.techidaily.com/in-2024-4-quick-ways-to-transfer-contacts-from-apple-iphone-15-to-iphone-withwithout-itunes-drfone-by-drfone-transfer-from-ios/"><u>In 2024, 4 Quick Ways to Transfer Contacts from Apple iPhone 15 to iPhone With/Without iTunes | Dr.fone</u></a></li>
<li><a href="https://review-topics.techidaily.com/identify-missing-or-malfunctioning-your-drivers-with-windows-device-manager-on-windows-11-and-10-and-7-by-drivereasy-guide/"><u>Identify missing or malfunctioning your drivers with Windows Device Manager on Windows 11 & 10 & 7</u></a></li>
<li><a href="https://howto.techidaily.com/how-to-fix-part-of-the-touch-screen-not-working-on-motorola-edge-40-neo-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>How To Fix Part of the Touch Screen Not Working on Motorola Edge 40 Neo | Dr.fone</u></a></li>
<li><a href="https://location-social.techidaily.com/in-2024-how-to-leave-a-life360-group-on-apple-iphone-7-without-anyone-knowing-drfone-by-drfone-virtual-ios/"><u>In 2024, How To Leave a Life360 Group On Apple iPhone 7 Without Anyone Knowing? | Dr.fone</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/how-to-mirror-your-vivo-v30-pro-screen-to-pc-with-chromecast-drfone-by-drfone-android/"><u>How to Mirror Your Vivo V30 Pro Screen to PC with Chromecast | Dr.fone</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/in-2024-how-to-track-imei-number-of-nokia-c12-pro-through-google-earth-by-drfone-android/"><u>In 2024, How To Track IMEI Number Of Nokia C12 Pro Through Google Earth?</u></a></li>
<li><a href="https://fix-guide.techidaily.com/proven-ways-to-fix-there-was-a-problem-parsing-the-package-on-sony-xperia-1-v-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>Proven Ways to Fix There Was A Problem Parsing the Package on Sony Xperia 1 V | Dr.fone</u></a></li>
<li><a href="https://iphone-unlock.techidaily.com/how-to-bypass-the-required-apple-store-verification-for-apple-iphone-14-drfone-by-drfone-ios/"><u>How To Bypass the Required Apple Store Verification For Apple iPhone 14 | Dr.fone</u></a></li>
<li><a href="https://unlock-android.techidaily.com/bypassing-google-account-with-vnrom-bypass-for-xiaomi-civi-3-disney-100th-anniversary-edition-by-drfone-android/"><u>Bypassing Google Account With vnROM Bypass For Xiaomi Civi 3 Disney 100th Anniversary Edition</u></a></li>
<li><a href="https://howto.techidaily.com/top-4-android-system-repair-software-for-tecno-pop-8-bricked-devices-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>Top 4 Android System Repair Software for Tecno Pop 8 Bricked Devices | Dr.fone</u></a></li>
<li><a href="https://android-location.techidaily.com/in-2024-10-fake-gps-location-apps-on-android-of-your-gionee-f3-pro-drfone-by-drfone-virtual/"><u>In 2024, 10 Fake GPS Location Apps on Android Of your Gionee F3 Pro | Dr.fone</u></a></li>
<li><a href="https://ai-editing-video.techidaily.com/new-4k-videos-are-the-in-thing-because-of-their-flexibility-and-high-resolution-stick-around-to-learn-how-to-choose-the-best-frame-rate-for-4k-videos/"><u>New 4K Videos Are the in Thing because of Their Flexibility and High Resolution. Stick Around to Learn How to Choose the Best Frame Rate for 4K Videos</u></a></li>
<li><a href="https://android-unlock.techidaily.com/in-2024-how-to-unlock-samsung-galaxy-m14-5g-pin-codepattern-lockpassword-by-drfone-android/"><u>In 2024, How to Unlock Samsung Galaxy M14 5G PIN Code/Pattern Lock/Password</u></a></li>
<li><a href="https://bypass-frp.techidaily.com/in-2024-the-complete-guide-to-honor-x9a-frp-bypass-everything-you-need-to-know-by-drfone-android/"><u>In 2024, The Complete Guide to Honor X9a FRP Bypass Everything You Need to Know</u></a></li>
<li><a href="https://ios-unlock.techidaily.com/in-2024-what-does-jailbreaking-iphone-15-i-do-get-answers-here-by-drfone-ios/"><u>In 2024, What Does Jailbreaking iPhone 15 i Do? Get Answers here</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/how-to-unlock-realme-v30-phone-without-google-account-by-drfone-android/"><u>How to Unlock Realme V30 Phone without Google Account?</u></a></li>
<li><a href="https://techidaily.com/undelete-lost-photos-from-itel-a60s-by-fonelab-android-recover-photos/"><u>Undelete lost photos from Itel A60s.</u></a></li>
</ul></div>



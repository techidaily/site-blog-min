---
title: How to identify malfunctioning your drivers with Windows Device Manager on Windows 10 & 7
date: 2024-05-19T18:32:12.836Z
tags: 
  - driver
  - device driver
categories: 
  - apps
  - windows
description: This article describes How to identify malfunctioning your drivers with Windows Device Manager on Windows 10 & 7. Device Manager is a control panel applet in Microsoft Windows operating systems. It allows users to view and control the hardware attached to the computer. When a piece of hardware is not working, the offending hardware is highlighted for the user to deal with. The list of hardware can be sorted by various criteria. Device Manager was introduced with Windows 95 and later added to Windows 2000. In NT-based versions, it is included as a Microsoft Management Console snap-in.
keywords: identify missing drivers in Windows 11 & 10,update drivers,identify missing drivers in Windows 7,identify malfunctioning drivers in Windows 7,device manager
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

## What are drivers?

Drivers are like interpreters between Windows and your devices. For example, when Windows needs to display something on your monitor, it sends a command to your graphics card, and your graphics card then displays what Windows wants on your monitor.

But Windows and your graphics card don’t actually speak the same language. To understand each other, they need a translator. That translator is called a driver. It takes the Windows command and translates it into something your graphics card can understand. Your graphics card can then do as it’s told, and display the right thing on your monitor.

Similarly, if your graphics card needs to send some sort of response back to Windows, the driver translates the response into something Windows can understand.

In this example, what we’re talking about is a video driver, but your computer has many other drivers on it too – one for each device. Your speakers, your printer, your mouse, your USB hard drives, your network card, your keyboard and so on – they each have an associated driver.

And without all these drivers, none of your devices will work.

![What are drivers?](https://tools.techidaily.com/images/apps/drivereasy/pages/why_2.jpg)

## What happens if a driver is missing or outdated?

Every now and then, Microsoft will change the commands Windows sends to one of your devices (e.g. your network card). When this happens, the manufacturer of that device need to change the device driver too. They need to teach it the new Windows commands. Otherwise the drivers won’t be able to translate those commands for your devices, and your devices won’t work properly.

The same thing needs to happen when your device manufacturer changes the way your device talks, or the things it can do. They need to change the driver too. Otherwise Windows won’t be able to talk to the device, or take advantage of its new functionality, and your device won’t work properly.

Now when we say “your device won’t work properly”, sometimes this means simply that you miss out on new functionality or minor bug fixes. But it’s often a lot more serious than that. Your computer may even hang, crash or stop working completely. Remember, there’s a driver that controls your hard drive, for instance. If Windows can’t talk to your hard drive, it can’t access any of the data on your drive. Similarly, if Windows can’t talk to your network card, you won’t be able to access the internet, and if it can’t talk to your graphics card, you won’t be able to see anything on your monitor. These are just a few of the more serious issues outdated drivers can cause.

![What happens if a driver is missing or outdated?](https://tools.techidaily.com/images/apps/drivereasy/pages/why_3.jpg)

## Why update your drivers in Windows 11, 10 & Windows 7?

Many computer issues are caused by outdated device drivers. Particularly in Windows 10/11.

So if your computer has slowed down, you should update your drivers. If it’s crashing or hanging, update your drivers. If you can’t connect to the internet, update your drivers. If your mouse, keyboard, monitor or speakers are acting up, update your drivers. In fact, no matter what your issue, there’s a good chance updating your drivers will fix it.

To understand why, you first have to understand what drivers actually are…


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
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-xiaomi-redmi-13c-5g-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Xiaomi Redmi 13C 5G to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-nokia-c12-plus-get-deleted-pictures-back-with-ease-and-safety-by-fonelab-android-recover-pictures/"><u>How to Nokia C12 Plus Get Deleted Pictures Back with Ease and Safety?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-data-from-broken-iphone-12-screen-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Data from Broken iPhone 12 Screen | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-music-from-xiaomi-redmi-a2-by-fonelab-android-recover-music/"><u>How to retrieve erased music from Xiaomi Redmi A2</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-tecno-camon-20-pro-5g-pictures-an-easy-method-explained-by-fonelab-android-recover-pictures/"><u>How to Restore Deleted Tecno Camon 20 Pro 5G Pictures  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-pictures-files-from-xiaomi-redmi-12-5g-by-fonelab-android-recover-pictures/"><u>How To  Restore Missing Pictures Files from Xiaomi Redmi 12 5G.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-files-after-iphone-11-pro-max-factory-reset-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Files after iPhone 11 Pro Max Factory Reset? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-fix-errors-were-detected-while-saving-pdf-stellar-by-stellar-guide/"><u>How to Fix Errors were detected while saving PDF? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-broken-video-files-of-tecno-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair Broken video files of Tecno ?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-infinix-zero-30-5g-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Infinix Zero 30 5G? </u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-videos-from-motorola-edge-40-neo-by-fonelab-android-recover-video/"><u>How to Rescue Lost Videos from Motorola Edge 40 Neo</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-messages-from-infinix-by-fonelab-android-recover-messages/"><u>How to Rescue Lost Messages from Infinix</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-wiped-call-history-on-oppo-find-n3-flip-by-fonelab-android-recover-call-logs/"><u>How to restore wiped call history on Oppo Find N3 Flip?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-itel-s23plus-by-fonelab-android-recover-music/"><u>How to recover old music from your Itel S23+</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-google-frp-lock-on-poco-x6-by-drfone-android-unlock-remove-google-frp/"><u>How to remove Google FRP Lock on Poco X6</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-pictures-files-from-asus-rog-phone-8-pro-by-fonelab-android-recover-pictures/"><u>How To  Restore Missing Pictures Files from Asus ROG Phone 8 Pro.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-honor-x7b-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Honor X7b</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-lava-agni-2-5g-photos-an-easy-method-explained-by-fonelab-android-recover-photos/"><u>How to Restore Deleted Lava Agni 2 5G Photos  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-easy-ways-to-copy-contacts-from-vivo-v29-to-iphone-14-and-15-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Easy Ways to Copy Contacts from Vivo V29 to iPhone 14 and 15 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-videos-from-iphone-7-without-backup-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Videos from iPhone 7 Without Backup? | Stellar</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/how-to-use-ispoofer-on-tecno-spark-go-2023-drfone-by-drfone-virtual-android/"><u>How to use iSpoofer on Tecno Spark Go (2023)? | Dr.fone</u></a></li>
<li><a href="https://location-social.techidaily.com/in-2024-4-most-known-ways-to-find-someone-on-tinder-for-realme-narzo-60-5g-by-name-drfone-by-drfone-virtual-android/"><u>In 2024, 4 Most-Known Ways to Find Someone on Tinder For Realme Narzo 60 5G by Name | Dr.fone</u></a></li>
<li><a href="https://iphone-transfer.techidaily.com/fix-apple-iphone-8-stuck-on-data-transfer-verified-solution-drfone-by-drfone-transfer-from-ios/"><u>Fix Apple iPhone 8 Stuck on Data Transfer Verified Solution! | Dr.fone</u></a></li>
<li><a href="https://unlock-android.techidaily.com/how-to-unlock-xiaomi-redmi-note-12-5g-pattern-lock-if-forgotten-6-ways-by-drfone-android/"><u>How to Unlock Xiaomi Redmi Note 12 5G Pattern Lock if Forgotten? 6 Ways</u></a></li>
<li><a href="https://android-unlock.techidaily.com/best-asus-rog-phone-8-pro-pattern-lock-removal-tools-remove-android-pattern-lock-without-losing-data-by-drfone-android/"><u>Best Asus ROG Phone 8 Pro Pattern Lock Removal Tools Remove Android Pattern Lock Without Losing Data</u></a></li>
<li><a href="https://change-location.techidaily.com/ways-to-trade-pokemon-go-from-far-away-on-vivo-y100i-power-5g-drfone-by-drfone-virtual-android/"><u>Ways to trade pokemon go from far away On Vivo Y100i Power 5G? | Dr.fone</u></a></li>
<li><a href="https://android-location-track.techidaily.com/3-ways-to-track-vivo-y28-5g-without-them-knowing-drfone-by-drfone-virtual-android/"><u>3 Ways to Track Vivo Y28 5G without Them Knowing | Dr.fone</u></a></li>
<li><a href="https://android-location.techidaily.com/in-2024-fake-android-location-without-rooting-for-your-vivo-y200-drfone-by-drfone-virtual/"><u>In 2024, Fake Android Location without Rooting For Your Vivo Y200 | Dr.fone</u></a></li>
<li><a href="https://ios-location-track.techidaily.com/in-2024-how-to-detect-and-remove-spyware-on-apple-iphone-7-plus-drfone-by-drfone-virtual-ios/"><u>In 2024, How to Detect and Remove Spyware on Apple iPhone 7 Plus? | Dr.fone</u></a></li>
<li><a href="https://ios-unlock.techidaily.com/guide-on-how-to-change-your-apple-id-email-address-on-apple-iphone-6-plus-by-drfone-ios/"><u>Guide on How To Change Your Apple ID Email Address On Apple iPhone 6 Plus</u></a></li>
<li><a href="https://review-topics.techidaily.com/quickly-remove-google-frp-lock-on-asus-by-drfone-android-unlock-remove-google-frp/"><u>Quickly Remove Google FRP Lock on Asus</u></a></li>
<li><a href="https://android-unlock.techidaily.com/how-to-unlock-vivo-v29-phone-without-google-account-by-drfone-android/"><u>How to Unlock Vivo V29 Phone without Google Account?</u></a></li>
<li><a href="https://review-topics.techidaily.com/recover-iphone-x-data-from-icloud-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>Recover iPhone X Data From iCloud | Dr.fone</u></a></li>
<li><a href="https://unlock-android.techidaily.com/in-2024-unlock-xiaomi-redmi-note-12-5g-phone-password-without-factory-reset-full-guide-here-by-drfone-android/"><u>In 2024, Unlock Xiaomi Redmi Note 12 5G Phone Password Without Factory Reset Full Guide Here</u></a></li>
<li><a href="https://change-location.techidaily.com/why-is-ipogo-not-working-on-vivo-t2-5g-fixed-drfone-by-drfone-virtual-android/"><u>Why is iPogo not working On Vivo T2 5G? Fixed | Dr.fone</u></a></li>
<li><a href="https://unlock-android.techidaily.com/in-2024-how-to-show-wi-fi-password-on-xiaomi-redmi-k70-pro-by-drfone-android/"><u>In 2024, How to Show Wi-Fi Password on Xiaomi Redmi K70 Pro</u></a></li>
<li><a href="https://apple-account.techidaily.com/how-to-fix-when-apple-account-locked-from-apple-iphone-13-mini-by-drfone-ios/"><u>How to Fix when Apple Account Locked From Apple iPhone 13 mini?</u></a></li>
<li><a href="https://android-unlock.techidaily.com/in-2024-mastering-lock-screen-settings-how-to-enable-and-disable-on-vivo-s17e-by-drfone-android/"><u>In 2024, Mastering Lock Screen Settings How to Enable and Disable on Vivo S17e</u></a></li>
<li><a href="https://change-location.techidaily.com/how-to-exit-android-factory-mode-on-vivo-t2-pro-5g-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>How to Exit Android Factory Mode On Vivo T2 Pro 5G? | Dr.fone</u></a></li>
</ul></div>



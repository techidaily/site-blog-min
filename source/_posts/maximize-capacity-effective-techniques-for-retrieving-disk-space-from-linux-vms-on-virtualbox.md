---
title: "Maximize Capacity: Effective Techniques for Retrieving Disk Space From Linux VMs on VirtualBox"
date: 2024-08-30T01:09:48.273Z
updated: 2024-08-31T01:09:48.273Z
tags:
  - desktop
categories:
  - tech
thumbnail: https://thmb.techidaily.com/8404aae8332517e90ea13209ccbcb49d56b9cbe41228f9bbeee698b42d6caf34.jpg
---

## Maximize Capacity: Effective Techniques for Retrieving Disk Space From Linux VMs on VirtualBox

### Key Takeaways

* VirtualBox doesn't automatically shrink Linux guest disks, but you can manually reclaim disk space by zeroing out data and compacting the volume.
* You should enable dynamic allocation in VirtualBox and back up your data before proceeding with the disk compacting process.
* Use the "dd" command to zero-out empty space in the Linux guest and then use VBoxManage to compact the virtual disk image.

 VirtualBox lets you run Linux in a virtual machine, and you'll often find your virtual disks continually growing in size, even though you've been clearing them of files. I'll show you how to shrink these volumes back down to size, compacting them and saving your disk space.

##  Why Your VirtualBox Linux Guest Isn’t Automatically Shrinking

 If you've used [VirtualBox](https://remote-screen-capture.techidaily.com/updated-2024-approved-little-gamers-treasure-trove-of-joy/) with Windows guests (in virtual machine terminology the "guest" is the operating system running within the virtual machine), you're probably used to your VirtualBox disk volumes shrinking as you delete files from them, so that they only use up as much space on your physical disk as they need to contain the files in them.

 This is the intended purpose of VirtualBox's dynamic allocation feature, but it doesn't work with [Linux guests](https://extra-guidance.techidaily.com/updated-snapshot-sophistication-editing-to-dazzle/). When using a Linux guests, many users find the disks grow to their full size, and then never shrink back down as files are deleted from them.

 This is due to how Linux manages its filesystems and how it interacts with its VirtualBox host. Linux doesn't "zero out" (overwrite with empty data) files when they are deleted for performance reasons, so VirtualBox has no way to tell what data on a virtual disk is active data and which is deleted.

 You can solve this issue by zeroing that data yourself, and telling VirtualBox to compact the volume, bringing it back down to its actual size. Here's how it's done.

<!-- affiliate ads begin -->
<a href="https://thefitville.pxf.io/c/5597632/1526796/15852" target="_top" id="1526796"><img src="//a.impactradius-go.com/display-ad/15852-1526796" border="0" alt="" width="1200" height="628"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1526796/15852" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
##  How to Reclaim Disk Space From a Linux Virtual Machine in VirtualBox

 The first thing you need to do to reclaim disk space from your Linux guests' virtual disks is ensure that dynamic allocation is enabled. If it is not, you will need to [convert your disk to a dynamically allocated disk](https://fox-info.techidaily.com/new-visualize-verve-vocalize-laughter-kapwings-toolkit/).

![How to find out whether 'Dynamically allocated storage' is enabled for your virtual disk.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/02/vbox-dynamic-disk-1.png) 

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2087484/7443" target="_top" id="2087484"><img src="//a.impactradius-go.com/display-ad/7443-2087484" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2087484/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
 Next, [back up your virtual disk](https://extra-information.techidaily.com/mac-and-pcs-top-10-supercharged-srt-systems-unveiled/). If something goes wrong (like a mistyped command, or your power going out part way through the process), you risk losing all the data in your Linux guest.

 Next, on your Linux guest [run the following command in the terminal](https://vimeo-videos.techidaily.com/updated-boost-your-income-with-effective-vimeo-monetization-techniques-for-2024/):

sudo dd if=/dev/zero of=/var/deleteme

 This command will write zero'd out (empty) data to the file /var/deleteme until the disk is completely full. This overwrites all of your previously deleted files, solving the problem that Linux doesn't overwrite deleted data automatically.

 Be careful using the dd command as it will overwrite data without warning! Check your commands and paths carefully before running them.

 This process could take some time depending on the size of the volume, so be patient and do not interrupt it. Once it has finished, the process will exit (possibly with an error saying that it is out of space). Once this has happened, you can delete the zeroed-out file and shut down your Linux guest:

sudo rm -rf /var/deleteme

    
                    sudo shutdown now -h

 Now the unused space on your virtual disk is zeroed out and VirtualBox will be able to reclaim the space on your host by shrinking it. The final step is to use the vboxmanage command to compact the virtual disk image. Do this by running:

        `vboxmanage modifymedium disk /path/to/image.vdi -compact`
    
 You must change /path/to/image.vdi to the path of the virtual disk you want to compact. If you are on Windows, you will need to use VBoxManage.exe like so:

VBoxManage.exe modifymedium disk /path/to/image.vdi -compact

 If VBoxManage.exe is not available from the command line on your Windows system, read on for instructions on how to enable it.

<!-- affiliate ads begin -->
<a href="https://estore.winxdvd.com/order/checkout.php?PRODS=4081991&QTY=1&AFFILIATE=108875&CART=1"><img src="https://www.winxdvd.com/affiliate/new-banner/wt-500x500.jpg" border="0"></a>
<!-- affiliate ads end -->
##  What Did the dd Linux Command Do?

 This method for compacting a VirtualBox Linux guest utilizes the dd command to write an empty file to disk so that empty space can be identified by VirtualBox.

 The [dd](https://ss64.com/bash/dd.html) (data duplicator) command converts and copies files, and can also be used to write data. The "if" option passed to it supplies the input file (in this case /dev/zero supplies a constant stream of zero-value data or null data). The "of" option specifies the output file, and this stream of zero data is written to it. This will continue until the disk is full as /dev/zero never stops providing null data.

<!-- affiliate ads begin -->
<a href="https://zebaoaffiliateprogram.pxf.io/c/5597632/1853659/21526" target="_top" id="1853659"><img src="//a.impactradius-go.com/display-ad/21526-1853659" border="0" alt="" width="1920" height="750"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1853659/21526" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
##  Using VboxManage.exe in Windows

 By default, VBoxManage.exe isn't available on the Windows command line. You can add it by [updating your Windows system path](https://screen-mirroring-recording.techidaily.com/updated-ideal-systems-for-recording-and-streaming-athletic-competitions-for-2024/) to include the VirtualBox installation directory, or calling the full path to the executable when using it:

& "C:/Path/To/VBoxManage.exe" modifymedium disk /path/to/image.vdi -compact

 The "&" symbol, called the call operator, that executes the quoted command. This lets you use spaces in path to the executable.

<!-- affiliate ads begin -->
<span id="1993650">
					<video width="720" height="300" style="cursor:pointer"
           poster="//a.impactradius-go.com/display-clicktoplayimage/1993650.jpeg"
           onclick="if(!this.playClicked){this.play();this.setAttribute('controls',true);this.playClicked=true;}">
	   <source src="//a.impactradius-go.com/display-ad/22993-1993650">
	   <img src="//a.impactradius-go.com/display-clicktoplayimage/1993650.jpeg" style="border: none; height: 100%; width: 100%; object-fit: contain">
	</video>
	<div style="width:720px;text-align:center"><a href="javascript:window.open(decodeURIComponent('https%3A%2F%2Fhomestyler.sjv.io%2Fc%2F5597632%2F1993650%2F22993'), '_blank');void(0);">Click here</a></div>
</span>
<img height="0" width="0" src="https://imp.pxf.io/i/5597632/1993650/22993" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
##  More About Managing VirtualBox Guests

 As your guests grow, you can [increase the size](https://tech-recovery.techidaily.com/the-ultimate-guide-16-best-free-sources-for-learning-american-sign-language/) of their virtual disks as well as [reduce them](https://screen-video-capture.techidaily.com/new-vocalvoyage-listening-and-recording-expedition-for-2024/). [Snapshotting](https://facebook-video-recording.techidaily.com/updated-fb-video-downloader-extraordinaire-mp4-transformation-for-2024/) lets you take the state of your virtual machine at a certain point in time and save it; if you later want to go back to how things were at that exact moment (for example after testing a configuration change), you can simply roll back and everything will be as it was.

 VirtualBox is a powerful virtualization tool that is used both professionally by developers to build and test software, and home users to run older software and games on modern computers, or run other operating systems without having to purchase a second machine. You can run [Windows](https://some-skills.techidaily.com/the-secrets-of-selecting-a-powerful-streaming-device-for-2024/), [Linux](https://extra-guidance.techidaily.com/updated-snapshot-sophistication-editing-to-dazzle/), [ChromeOS](https://ios-unlock.techidaily.com/in-2024-unlocking-iphone-xs-max-passcode-without-a-computer-by-drfone-ios/), Android, and other operating systems on MacOS, Windows, and Linux Hosts.

<ins class="adsbygoogle"
     style="display:block"
     data-ad-format="autorelaxed"
     data-ad-client="ca-pub-7571918770474297"
     data-ad-slot="1223367746"></ins>



<ins class="adsbygoogle"
     style="display:block"
     data-ad-client="ca-pub-7571918770474297"
     data-ad-slot="8358498916"
     data-ad-format="auto"
     data-full-width-responsive="true"></ins>

<span class="atpl-alsoreadstyle">Also read:</span>
<div><ul>
<li><a href="https://youtube-data.techidaily.com/024-approved-creating-a-memorable-and-unique-moniker-for-your-filmora-content/"><u>[New] 2024 Approved  Creating a Memorable and Unique Moniker for Your Filmora Content</u></a></li>
<li><a href="https://screen-mirroring-recording.techidaily.com/new-loom-screen-recorder-downloader/"><u>[New] Loom Screen Recorder Downloader</u></a></li>
<li><a href="https://fox-cloud.techidaily.com/updated-2024-approved-deciphering-the-rights-based-language-of-instagrams-music-features/"><u>[Updated] 2024 Approved  Deciphering the Rights-Based Language of Instagram's Music Features</u></a></li>
<li><a href="https://youtube-blog.techidaily.com/ed-in-2024-whats-your-youtube-revenue-breakdown-per-branded-playlist/"><u>[Updated] In 2024, What's Your Youtube Revenue Breakdown per Branded Playlist?</u></a></li>
<li><a href="https://some-techniques.techidaily.com/2024-approved-hype-generating-headline-author/"><u>2024 Approved  Hype-Generating Headline Author</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-ways-to-transfer-music-from-vivo-v29-to-other-android-devices-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Ways to Transfer Music from Vivo V29 to Other Android Devices Easily | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/6-ways-to-transfer-contacts-from-vivo-s18e-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>6 Ways To Transfer Contacts From Vivo S18e to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/android-to-apple-how-to-transfer-photos-from-htc-u23-to-ipad-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>Android to Apple How To Transfer Photos From HTC U23 to iPad Easily | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/android-to-apple-how-to-transfer-photos-from-oppo-k11x-to-ipad-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>Android to Apple How To Transfer Photos From Oppo K11x to iPad Easily | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/android-to-apple-how-to-transfer-photos-from-realme-gt-neo-5-se-to-ipad-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>Android to Apple How To Transfer Photos From Realme GT Neo 5 SE to iPad Easily | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/android-to-apple-how-to-transfer-photos-from-samsung-galaxy-a14-4g-to-ipad-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>Android to Apple How To Transfer Photos From Samsung Galaxy A14 4G to iPad Easily | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/android-to-apple-how-to-transfer-photos-from-tecno-spark-go-2023-to-ipad-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>Android to Apple How To Transfer Photos From Tecno Spark Go (2023) to iPad Easily | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/best-3-software-to-transfer-files-tofrom-your-oppo-reno-11-pro-5g-via-a-usb-cable-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>Best 3 Software to Transfer Files to/from Your Oppo Reno 11 Pro 5G via a USB Cable | Dr.fone</u></a></li>
<li><a href="https://extra-lessons.techidaily.com/glow-innovations-outshining-videos-through-design/"><u>Glow Innovations  Outshining Videos Through Design</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-can-you-transfer-files-from-oneplus-nord-n30-5g-to-iphone-151413-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How Can You Transfer Files From OnePlus Nord N30 5G To iPhone 15/14/13? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-can-you-transfer-files-from-poco-c50-to-iphone-151413-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How Can You Transfer Files From Poco C50 To iPhone 15/14/13? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-can-you-transfer-files-from-xiaomi-redmi-note-12-pro-4g-to-iphone-151413-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How Can You Transfer Files From Xiaomi Redmi Note 12 Pro 4G To iPhone 15/14/13? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-migrate-android-data-from-oppo-f25-pro-5g-to-new-android-phone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Migrate Android Data From Oppo F25 Pro 5G to New Android Phone? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-migrate-android-data-from-vivo-v27-to-new-android-phone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Migrate Android Data From Vivo V27 to New Android Phone? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-apps-from-motorola-moto-g14-to-another-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Apps from Motorola Moto G14 to Another | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-apps-from-tecno-spark-20c-to-another-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Apps from Tecno Spark 20C to Another | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-apps-from-vivo-y78t-to-another-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Apps from Vivo Y78t to Another | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-contacts-from-huawei-nova-y91-to-other-android-devices-using-bluetooth-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Contacts from Huawei Nova Y91 to Other Android Devices Using Bluetooth? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-contacts-from-poco-c55-to-iphone-xs11-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Contacts from Poco C55 to iPhone XS/11 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-contacts-from-realme-11x-5g-to-other-android-devices-devices-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Contacts from Realme 11X 5G to Other Android Devices Devices? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-contacts-from-vivo-y100t-to-other-android-devices-devices-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Contacts from Vivo Y100t to Other Android Devices Devices? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-contacts-from-vivo-y100t-to-outlook-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Contacts from Vivo Y100t to Outlook | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-contacts-from-vivo-y78t-to-phone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Contacts from Vivo Y78t To Phone | Dr.fone</u></a></li>
<li><a href="https://facebook-clips.techidaily.com/in-2024-efficient-fb-link-extraction-top-8-tools-for-no-cost-online-use/"><u>In 2024, Efficient FB Link Extraction  Top 8 Tools for No-Cost, Online Use</u></a></li>
<li><a href="https://blog-min.techidaily.com/in-2024-how-to-teleport-your-gps-location-on-motorola-moto-g-stylus-2023-drfone-by-drfone-virtual-android/"><u>In 2024, How To Teleport Your GPS Location On Motorola Moto G Stylus (2023)? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/in-2024-how-to-teleport-your-gps-location-on-motorola-moto-g04-drfone-by-drfone-virtual-android/"><u>In 2024, How To Teleport Your GPS Location On Motorola Moto G04? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/in-2024-how-to-teleport-your-gps-location-on-vivo-y100-drfone-by-drfone-virtual-android/"><u>In 2024, How To Teleport Your GPS Location On Vivo Y100? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/in-2024-how-to-teleport-your-gps-location-on-xiaomi-redmi-note-12-4g-drfone-by-drfone-virtual-android/"><u>In 2024, How To Teleport Your GPS Location On Xiaomi Redmi Note 12 4G? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/in-2024-how-to-use-life360-on-windows-pc-for-htc-u23-pro-drfone-by-drfone-virtual-android/"><u>In 2024, How to Use Life360 on Windows PC For HTC U23 Pro? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/in-2024-how-to-use-life360-on-windows-pc-for-lava-yuva-3-pro-drfone-by-drfone-virtual-android/"><u>In 2024, How to Use Life360 on Windows PC For Lava Yuva 3 Pro? | Dr.fone</u></a></li>
<li><a href="https://extra-hints.techidaily.com/saving-gifs-on-your-iphone-a-step-by-step-approach/"><u>Saving GIFs on Your iPhone - A Step-by-Step Approach</u></a></li>
<li><a href="https://hardware-updates.techidaily.com/step-by-step-guide-upgrading-and-downloading-insignia-bluetooth-device-drivers-on-windows/"><u>Step-by-Step Guide: Upgrading and Downloading Insignia Bluetooth Device Drivers on Windows</u></a></li>
<li><a href="https://vp-tips.techidaily.com/unlock-creative-shots-with-photoshops-radial-distortion-for-2024/"><u>Unlock Creative Shots with Photoshop's Radial Distortion for 2024</u></a></li>
</ul></div>

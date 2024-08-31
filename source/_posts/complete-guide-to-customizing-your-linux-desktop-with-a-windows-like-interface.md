---
title: Complete Guide to Customizing Your Linux Desktop with a Windows-Like Interface
date: 2024-08-26 20:48:26
updated: 2024-08-29 10:49:29
tags:
  - desktop
categories:
  - tech
thumbnail: https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2023/05/52687750468_dc6bdda141_o-1.jpg
---

## Complete Guide to Customizing Your Linux Desktop with a Windows-Like Interface

### Quick Links

* [Method 1: Get a Windows 10 GTK Theme](https://extra-tips.techidaily.com/where-to-watch-vr-content/)
* [Method 2: Use a Community Theme Pack](https://some-skills.techidaily.com/new-the-instructors-blueprint-for-video-assisted-learning/)
* [Why Clone Windows 10 With Linux?](https://sim-unlock.techidaily.com/in-2024-how-to-unlock-sim-card-on-samsung-galaxy-xcover-7-online-without-jailbreak-by-drfone-android/)

### Key Takeaways

* For GTK Linux desktop environments, download the B00merang Windows 10 theme pack and icon pack and add their contents to themes and icons folders, respectively. Enable them using your theme manager.
* For all desktops, alternatively use a community theme pack to automate the installation process, but be aware that it may not be an exact replica of Windows 10.

 Want to use a Linux desktop but don't want to lose Windows 10's familiar look and vibe? Most desktop environments come with unique layouts and themes by default, but with a bit of effort, you can convert it to a spitting image of the Windows 10 desktop.

##  Method 1: Get a Windows 10 GTK Theme

 The method that works on many desktop environments is to simply download a Windows 10 theme pack and icon pack designed for GTK desktops and drop them into the proper folders. We'll demonstrate this with Cinnamon (on Linux Mint 23.3, specifically), but it'll also work on [GNOME](https://some-guidance.techidaily.com/new-the-complete-vivacut-overview-editors-deep-dive/), Xfce, LXDE, and MATE.

###  Step 1: Download and Install the Theme

 The best source for Windows 10 GTK theming files is [the B00merang Project's GitHub release page](https://github.com/B00merang-Project/Windows-10/releases). That is the light version, and though we'll use it as an example in this guide, you can grab [the Windows 10 dark variant](https://github.com/B00merang-Project/Windows-10-Dark/releases) if you prefer.

 In the latest release card (3.2.1 at the time of writing), download your Windows 10 theming by clicking "Source Code (ZIP)."

![GitHub Release page for B00-merang-Project's Windows 10 theme, with the zip file download labeled Source Code highlighted.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/02/2024-02-12_13h33_54.png) 

 When the ZIP file download completes, you'll need to extract the contents. For a quick command line solution, first run this command in your home directory to create the folder in case it doesn't already exist:

mkdir -p ~/.themes

 Follow that with [the unzip command](https://youtube-help.techidaily.com/first-steps-launching-a-youtube-channel-for-profit-for-2024/) to decompress it and drop it in the themes folder:

unzip Downloads/Windows-10-3.2.1.zip -d ~/.themes

 Just tweak the directory path and name for the file if yours is different. This will cause the files to be extracted into your Downloads directory.

 Want to use the desktop interface instead of the command line? Just open your file browser and look for the downloaded ZIP file. It should be named something like "Windows-10-3.2.1.zip". Double-click it to open the extraction dialog, which will look different depending on your distribution and DE. Regardless, there should be a noticeable "Extract" button for you to click. In our example, it's in the top-left corner of the dialog.

![Extract button highlighted in the GNOME Archive Manager.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/02/2024-02-12_14h47_23.png) 

 You'll be prompted to choose an extraction location, and, for simplicity, we'll leave it at the default: the Downloads folder where we found the ZIP file.

 Once extraction has finished, locate the extracted directory in your file explorer, likely also called "Windows-10-3.2.1". Copy the directory, and with it in your clipboard, move back to your home folder.

 At this point, you'll need to make hidden files visible in your file browser. You can usually find this setting in the View drop-down menu, as is the case in Nemo.

![The Show Hidden Files option in the Nemo file manager for Linux.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/02/2024-02-12_14h53_57.png) 

 You should now see several folders that weren't there before and whose names start with a period like ".local". You need the ".themes" folder, but in our version of Linux it didn't exist yet, so we just made one.

![Creating a folder named .themes in a Linux home directory using a file manager.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/02/2024-02-12_14h57_55.png) 

 After creating the .themes folder, open it and paste your theme folder there by either right-clicking and selecting "Paste" or using the keyboard shortcut Ctrl+V.

![Windows 10 theme pasted into the .themes folder on a Linux system.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/02/2024-02-12_15h05_13.png) 

###  Step 2: Download and Install the Icon Pack

 The basic theme package contains what you need to get the overall design of Windows 10, but if you want those illustrious icons, you need to get and install those separately.

 They're available from the same project, but in [a separate icons repository](https://github.com/B00merang-Artwork/Windows-10/). Again, download the ZIP file for the latest release and unzip it like you did the theme pack. Copy the unzipped contents, and this time paste it into the .icons folder in your home directory. On my system, the folder already existed, but you may have to create it yourself.

![Windows 10 icon folder in the icons folder in a Linux home directory](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/02/2024-02-12_15h22_21.png) 

###  Step 3: Enable Your Windows 10 Theme

 With your theme and icons in place, it's time to paint the town Windows. Open your DE's settings and look for its theming preferences, likely called "Appearance" or, in the case of Cinnamon, "Themes."

![The Themes menu available in the Cinnamon desktop environment.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/02/2024-02-12_15h26_20.png) 

 Then we need to choose the Windows 10 theme for every category available. To get this ability in Cinnamon, I needed to click "Advanced Settings."

![If you're using the Cinnamon desktop environment, click "Advanced Settings" in your Themes menu.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2023/08/2023-08-10_15h26_22.png) 

 Again, the interface will look different depending on precisely your distro and DE. But in Cinnamon, there are three categories you can change to Windows: Applications, Icons, and Desktop. Click the button in each category to open the options.

![Cinnamon breaks theming down into applications, icons, and desktop.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2023/08/2023-08-10_15h29_17.png) 

 You might need to scroll to reveal it, but there will be an option named Windows-10-1.2.3 (or whatever theme or icon version you downloaded is named).

![The Windows 10 theme selection in the Cinnamon Theme settings menu.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/02/2024-02-12_15h36_04.png) 

 The finishing touch is of course [the classic Windows 10 background](https://video-screen-grab.techidaily.com/updated-in-2024-dynamic-public-speaker-review-v8/). There's a large collection of the older ones [on Imgur](https://imgur.com/a/3lEQM). I wanted the lighter blue version of the glass pane one, which I found [at 4K Wallpapers](https://4kwallpapers.com/technology/windows-10-windows-logo-glossy-blue-background-2733.html).

![Cinnamon desktop environment with Windows 10 theme applied.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/02/2024-02-12_15h46_04.png) 

 We're now looking a lot better, but some things are a bit off. The taskbar in particular doesn't look quite right. Keep in mind these themes can only go so far in cloning Windows 10\. If you're willing to explore your desktop settings, though, there are often settings that will get you closer to 2017-era Windows nirvana.

 For example, in Cinnamon, I can replace the Mint logo in the Start menu by right-clicking it and choosing Configure > Appearance > Icon.

![Choose the "start-here-symbolic" icon to fix the Start menu icon in Cinnamon.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2023/08/2023-08-11_14h39_57.png) 

 If I want to see window titles in open taskbar apps, I can right-click any app icon and go to Applet Preferences > Configure > Panel > Button Label, then select "Window Title."

![Change the "Button Label" option in your applet panel settings to "Window Title" so that you can see app names in your Cinnamon task bar.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2023/08/2023-08-16_15h42_08.png) 

 And I can also get the My Computer icon and others on my desktop. If you right-click anywhere on the desktop and choose Customize > Desktop Settings. Just toggle on the ones you want. (Note that "Home" is what you'd call the "User's Files" on Windows.)

![Toggle on and off your desktop icons in Cinnamon.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2023/08/2023-08-11_15h27_39.png) 

 Now my Cinnamon desktop is looking pretty good.

![Cinnamon desktop with additional Windows 10 tweaks applied.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/02/2024-02-12_15h50_56.png) 

##  Method 2: Use a Community Theme Pack

 Some desktop environments have a built-in tool for finding and installing themes built and shared by community members. In my experience, these tend to be hit-and-miss. They may not be as perfect a reproduction as you hope for, so you may end up needing manual tweaks anyway. Icons may be missing. Updates to the desktop environment itself can break the theme, too, forcing you to further tweak or undo the theming altogether.

 That said, they can save you from the hassle of downloading, unzipping, and properly installing the themes; the community theme typically automates that. In fact, they sometimes just download the same theme pack I showed you how to download manually. So, when they work, community themes can be awesome. Plus, this can work on desktop environments that don't support GTK theming, like KDE Plasma.

 Thism of course, will look different depending on your desktop environment, but on Cinnamon, you can go to Themes > Add/Remove and type "windows 10" into the search bar to find a community-uploaded theme. Click the download button next to the one you want.

![In Cinnamon's theme manager, click the download button next to the theme you want to download.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2023/08/2023-08-11_15h32_57.png) 

 Once you've downloaded the theme, you've then got to go back to the Themes tab and activate it. You may have to scroll to find it.

![Select the Windows 10 theme you downloaded.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2023/08/2023-08-11_15h41_53.png) 

 Select it and, if all went well, you've got your Windows 10 theme.

![Linux Mint Cinnamon desktop with a Windows 10 theme installed.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2023/08/2023-08-16_16h46_07.png) 

 I added a classic Windows 10 background myself. As you can see, though, there are none of those recognizable icons. At least on Cinnamon, you'll have to manually install them, which I walked through above. You may want to make the extra tweaks I outlined above too.

##  Why Clone Windows 10 With Linux?

 For me, the unconventional interface some desktop environments bring is part of the appeal of switching to Linux. Not everyone is like me, though, and you may want to switch to Linux without your PC desktop world turned upside down. Even if you are like me, this is still a great trick to know if you're [replacing an installation of Windows with Linux](https://extra-resources.techidaily.com/your-journey-through-telegrams-web-functionality/) on the PC of a friend or relative who doesn't want to put up with Windows' annoyances anymore.

 There's also the nostalgia factor. While Windows 10 isn't exactly a vintage look yet, it's already being slowly replaced via [Windows 11 upgrades](https://easy-unlock-android.techidaily.com/everything-you-need-to-know-about-lock-screen-settings-on-your-realme-12-5g-by-drfone-android/). As time goes on, and [Microsoft stops supporting Windows 10](https://remote-screen-capture.techidaily.com/new-in-2024-screenflow-unleashed-the-ultimate-macos-experience/)altogether, this kind of theming will get even more niche.

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

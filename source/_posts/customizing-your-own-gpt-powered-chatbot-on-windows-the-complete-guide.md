---
title: Customizing Your Own GPT-Powered Chatbot on Windows - The Complete Guide
date: 2024-08-30T01:10:11.899Z
updated: 2024-08-31T01:10:11.899Z
tags:
  - desktop
categories:
  - tech
thumbnail: https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/06/person-using-a-windows-laptop-with-a-gpt-chat-window.jpg
---

## Customizing Your Own GPT-Powered Chatbot on Windows - The Complete Guide

### Quick Links

* [Why Would You Want Your Own Local AI Chatbot?](https://extra-information.techidaily.com/updated-10-best-live-streaming-platforms-you-should-know/)
* [How to Host Your Own Local GPT Chatbot on Windows](https://extra-guidance.techidaily.com/2024-approved-smart-shopping-tips-economical-gopro-cameras/)
* [Step 1: Install Docker and the Windows Terminal App](https://fox-glue.techidaily.com/updated-customizing-your-windows-photos-app-filters-and-music-sync-feature-for-2024/)
* [Step 2: Download the Text Generation Web UI GitHub Repository](https://win-solutions.techidaily.com/solving-the-outdated-driver-alert-in-minecraft-a-comprehensive-guide/)
* [Step 3: Launch the Text Generation Web UI](https://visual-screen-recording.techidaily.com/easy-peasy-guide-to-capturing-online-events-for-mac-and-windows-users/)
* [Step 4: Install the GPT-2 Model From OpenAI](https://extra-support.techidaily.com/lolkit-design-memes-and-graphics-with-a-click-for-2024/)
* [Step 5: Start Using Your Custom GPT AI Chatbot](https://some-approaches.techidaily.com/2024-approved-the-ultimate-key-to-exploring-without-boundaries-through-vr/)

### Key Takeaways

* Running your own local GPT chatbot on Windows is free from online restrictions and censorship.
* Install text-generation-web-ui using Docker on a Windows PC with WSL support and a compatible GPU.
* Customize and train your GPT chatbot for your own specific use cases, like querying and summarizing your own documents, helping you write programs, or imitating your own characters.

 There are a number of advantages to running a GPT/AI chatbot on your own computer rather than accessing one on the Internet. We'll show why you might want to, and the easiest way to get it set up on Windows.

<!-- affiliate ads begin -->
<a href="https://estore.zonealarm.com/order/checkout.php?PRODS=38658749&QTY=1&AFFILIATE=108875&CART=1"><img src="https://sc1.checkpoint.com/sc1/za/images/boxes/pa_500.png" border="0">ZoneAlarm Pro Antivirus + Firewall NextGen</a>
<!-- affiliate ads end -->
##  Why Would You Want Your Own Local AI Chatbot?

 While online AI chatbots like ChatGPT are dominant, with access to huge amounts of training data and up-to-date information, there are a few reasons you might want to run your own local chatbot on your Windows computer.

 Running your own AI tools locally is free, and comes without the restrictions of online tools: There's no censorship, and you can load whatever machine-learning models, tailor their responses and behavior, and provide any prompt, all without restriction (and in total privacy). It's also a good way to get an understanding of how modern large language model (LLM) AI tools work under the hood, especially if you are looking to get into the AI or tech industry.

 You can find models for just about anything, from fictional character interactions, to programming, to general knowledge, and many other use cases that more general online models may not cover.

##  How to Host Your Own Local GPT Chatbot on Windows

 This tutorial will use [text-generation-web-ui-docker](https://github.com/Atinoda/text-generation-webui-docker "https://github.com/oobabooga/text-generation-webui"), an open-source interface for large language models, that simplifies installing and using LLMs.

 text-generation-web-ui-docker bundles the text-generation-web-ui project using [Docker](https://vp-tips.techidaily.com/new-ultimate-auditory-interface-win-for-2024/), which removes the need for installing and managing all the complex dependencies that local AI tools usually require by storing everything in a container separate to your system. The only thing you need to run your local chatbot is a Windows PC that [supports Docker running using the Windows Subsystem for Linux (WSL)](https://docs.docker.com/desktop/install/windows-install/ "https://docs.docker.com/desktop/install/windows-install/"). You'll also need a fairly recent GPU, ideally one from NVIDIA, for maximum compatibility.

<!-- affiliate ads begin -->
<a href="https://zebaoaffiliateprogram.pxf.io/c/5597632/1853659/21526" target="_top" id="1853659"><img src="//a.impactradius-go.com/display-ad/21526-1853659" border="0" alt="" width="1920" height="750"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1853659/21526" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
##  Step 1: Install Docker and the Windows Terminal App

 Docker containers are similar to [virtual machines](https://remote-screen-capture.techidaily.com/2024-approved-essential-guide-video-recording-with-vlc/) in that they contain a whole running system, but they are much more lightweight, and perfect for distributing applications and all of their requirements in a single, easy-to-install bundle. To run text-generation-web-ui-docker in Docker, [download and install Docker on your Windows system](https://www.docker.com/products/docker-desktop/ "https://www.docker.com/products/docker-desktop/").

 Docker can run on Windows in one of two ways: WSL or Hyper-V mode. WSL is recommended for most users, so you may need to [enable it](https://some-skills.techidaily.com/new-unveiling-paid-content-in-product-analysis/) before installing Docker.

 It's also recommended to [install the Windows Terminal app](https://youtube-tips.techidaily.com/ed-mastering-youtube-streaming-a-guide-for-gamers-for-2024/), as it provides a convenient interface for WSL, PowerShell, and the Windows command line.

<!-- affiliate ads begin -->
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=4721564&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/c14a8df1e1b4d5297e9cb30cb34d5a00/products/copy_power-tools-48.png" border="0">Power Tools add-on for Google Sheets, 12-month subscription</a>
<!-- affiliate ads end -->
##  Step 2: Download the Text Generation Web UI GitHub Repository

 To download text-generation-web-ui-docker, [visit its GitHub page](https://github.com/Atinoda/text-generation-webui-docker "https://github.com/oobabooga/text-generation-webui"). You can download a ".zip" file containing all the files you need by clicking on the green "Code" button and then clicking on "Download Zip" from the drop-down menu.

 The screenshot below shows you where to find this on the GitHub page.

![Where to download the .zip file for a GitHub repository.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/06/download-zip-2.png) 

<!-- affiliate ads begin -->
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=3922934&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/4b0a0290ad7df100b77e86839989a75e/products/ripperpro.png" border="0">WonderFox DVD Ripper Pro</a>
<!-- affiliate ads end -->
Brad Morton / How-To Geek

[Extract the downloaded ZIP file](https://video-screen-grab.techidaily.com/updated-in-2024-essential-screen-capture-software-top-picks-ranked/) into its own folder, and then open the folder containing the unzipped files. Don't worry too much about the contents of this folder: it's all the moving parts for your AI chatbot, but Docker will take care of setting everything up for you.

##  Step 3: Launch the Text Generation Web UI

 Before you launch text-generation-web-ui-docker, you need to make sure it's configured for your PC hardware. Edit the file **docker-compose.yml** to specify the variant that matches your hardware: **default-nvidia** or **default-cpu** (if you don't have a compatible GPU).

![Configuring the docker image variant in docker-compose.yml.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/06/variant.png) 

Brad Morton / How-To Geek

 Then, still in the folder containing the downloaded and extracted files, right-click on an empty space and click "Open in Terminal" to open the directory in a new [PowerShell](https://some-guidance.techidaily.com/updated-unleash-creativity-fast-windows-10-photo-edits-made-simple/) window.

![The location of the 'Open in Terminal' in the context menu.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/06/open-in-terminal.png) 

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2068411/7443" target="_top" id="2068411"><img src="//a.impactradius-go.com/display-ad/7443-2068411" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2068411/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
Brad Morton / How-To Geek

 Make sure Docker Desktop is running before typing in the following Docker command into the Terminal window:

docker compose up -d

 This command will do the following:

* Executes **docker compose**, the program that manages Docker applications
* Start the container using the **up** command
* Runs the container in the background (called detached mode, which is specified by the **\-d**)

![The command for bringing up a Docker container.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/06/up.png) 

<!-- affiliate ads begin -->
<a href="https://versadesk.pxf.io/c/5597632/1892107/21290" target="_top" id="1892107"><img src="//a.impactradius-go.com/display-ad/21290-1892107" border="0" alt="" width="1200" height="628"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1892107/21290" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
Brad Morton / How-To Geek

 If you haven't run this command before and the application needs to be downloaded (pulled), you might need to go and make a cup of tea, as it could take a while. Once the command has completed running successfully, you'll see that the text-generation-web-ui-docker Docker container has been created and started.

![Windows terminal showing the status of a newly created container.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/06/created.png) 

<!-- affiliate ads begin -->
<h3 id="200610"><a href="https://sentrypc.7eer.net/c/5597632/200610/3022">Parental Control Software</a></h3>
<span class="text-ad-content">
	#1 Rated Parental Control Software.<br/>
	Monitor & Control all PC Activity!<br/>
		<cite style="color:green">sentrypc.com/parental-controls/</cite>
	</span><img height="0" width="0" src="https://sentrypc.7eer.net/i/5597632/200610/3022" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
Brad Morton / How-To Geek

 The running container will also appear in Docker Desktop, where you can stop, start, and manage it.

![Docker Desktop showing a running container.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/06/docker.png) 

Brad Morton / How-To Geek

<!-- affiliate ads begin -->
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=4726960&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/5f4f7141b65a730b4efb0e0d51f63e94/products/forexrobotronbox.gif" border="0">Forex Robotron Basic Package</a>
<!-- affiliate ads end -->
##  Step 4: Install the GPT-2 Model From OpenAI

 Once text-generation-web-ui-docker is up and running in Docker, you can access it by typing the address **http://localhost:7860** into your browser's address bar. [Localhost](https://youtube-docs.techidaily.com/cing-creativity-and-monetization-in-youtube-shorts-for-2024/) is the address your computer uses to access services it is running itself, each of which is assigned a unique port number (in this case 7860). You can see what ports a running Docker container has made available on localhost by opening it in Docker desktop.

 This is an older version of GPT than you get when you use ChatGPT on the internet. Consequently, it won't be as smart or as intuitive as what you might expect, but it is customizable and private.

 Navigate to the Model tab in the web interface and enter **openai-community/gpt2** into the "Download model or LoRA" box, and then click the Download button. This might take a few minutes.

![Downloading a model for use in text-generation-web-ui.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/05/win-gpt-6.png) 

Brad Morton / How-To Geek

 Click the Refresh icon in the top-left, then select the newly downloaded openai-community/gpt2 from the adjacent Model drop-down menu. Finally, click the "Load" button, and wait a few minutes until you see a success message.

![Loading a model in text-generation-web-ui.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/05/win-gpt-7.png) 

Brad Morton / How-To Geek

 This model works out of the box, and doesn't require any signup. As you get into AI and want to experiment with different models and AI tools, you can find more on [HuggingFace](https://huggingface.co/ "https://huggingface.co/").

##  Step 5: Start Using Your Custom GPT AI Chatbot

 This isn't quite like the AI chatbots you're used to using online, which are already set up for general use that cover a lot of situations. You'll need to tell the model how to behave before interacting with it, otherwise its output will be a bit... unhinged.

![A slightly unhinged response from an AI model that needs configuration.](https://static1.howtogeekimages.com/wordpress/wp-content/uploads/2024/05/win-gpt-8.png) 

Brad Morton / How-To Geek

_Hoo boy, while it got the right answer, this AI chatbot needs a bit of fine-tuning._ 

 There's a lot you can tweak, and it can be a bit clunky at first, but with practice and experience, you can build a chatbot that is specific to your own usage, and that keeps your data 100% on your own computer, which is great for business and other confidential use-cases. To learn how to configure and train your local GPT chatbot, check out the [text-generation-web-ui documentation](https://github.com/oobabooga/text-generation-webui/wiki "https://github.com/oobabooga/text-generation-webui/wiki") and the [OpenAI GP2 docs on HuggingFace](https://huggingface.co/docs/transformers/en/model%5Fdoc/gpt2#openai-gpt2 "https://huggingface.co/docs/transformers/en/model_doc/gpt2#openai-gpt2").

 If you want to generate images using AI from your Windows PC, [you can use Fooocus to get easier and more immediate results.](https://facebook-video-share.techidaily.com/updated-upgrading-your-videos-appeal-youtube-thumbnail-resizing/)

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
<li><a href="https://youtube-lab.techidaily.com/024-approved-comprehensive-list-of-8-prime-youtube-to-avi-tools/"><u>[New] 2024 Approved  Comprehensive List of 8 Prime YouTube-to-AVI Tools</u></a></li>
<li><a href="https://youtube-data.techidaily.com/pex-producer-syncing-shots-and-songs-for-2024/"><u>[New] Apex Producer  Syncing Shots & Songs for 2024</u></a></li>
<li><a href="https://visual-screen-recording.techidaily.com/new-evaluating-recmeisters-innovation-in-video-capture-for-2024/"><u>[New] Evaluating Recmeister's Innovation in Video Capture for 2024</u></a></li>
<li><a href="https://extra-resources.techidaily.com/updated-crafting-star-stirred-stories-for-your-whatsapp-bio/"><u>[Updated] Crafting Star-Stirred Stories for Your WhatsApp Bio</u></a></li>
<li><a href="https://vp-tips.techidaily.com/updated-premium-6-choices-for-cost-effective-4k-projection/"><u>[Updated] Premium 6 Choices for Cost-Effective 4K Projection</u></a></li>
<li><a href="https://some-approaches.techidaily.com/updated-the-ultimate-zoom-setup-playbook/"><u>[Updated] The Ultimate Zoom Setup Playbook</u></a></li>
<li><a href="https://fox-boxes.techidaily.com/2024-approved-capturing-more-with-less-lgs-360-camera-redefined-and-reviewed/"><u>2024 Approved  Capturing More with Less - LG's 360 Camera Redefined & Reviewed</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-vivo-s17e-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Vivo S17e Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-ways-to-transfer-music-from-samsung-galaxy-s24-to-other-android-devices-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Ways to Transfer Music from Samsung Galaxy S24 to Other Android Devices Easily | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/5-ways-to-transfer-music-from-samsung-galaxy-s24plus-to-other-android-devices-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>5 Ways to Transfer Music from Samsung Galaxy S24+ to Other Android Devices Easily | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/6-ways-to-transfer-contacts-from-motorola-razr-40-ultra-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>6 Ways To Transfer Contacts From Motorola Razr 40 Ultra to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/8-ways-to-transfer-photos-from-samsung-galaxy-a14-4g-to-iphone-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>8 Ways to Transfer Photos from Samsung Galaxy A14 4G to iPhone Easily | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/android-to-apple-how-to-transfer-photos-from-oneplus-nord-n30-5g-to-ipad-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>Android to Apple How To Transfer Photos From OnePlus Nord N30 5G to iPad Easily | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/android-to-apple-how-to-transfer-photos-from-xiaomi-13-ultra-to-ipad-easily-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>Android to Apple How To Transfer Photos From Xiaomi 13 Ultra to iPad Easily | Dr.fone</u></a></li>
<li><a href="https://fake-location.techidaily.com/apply-these-techniques-to-improve-how-to-detect-fake-gps-location-on-vivo-y100-drfone-by-drfone-virtual-android/"><u>Apply These Techniques to Improve How to Detect Fake GPS Location On Vivo Y100 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/best-3-software-to-transfer-files-tofrom-your-htc-u23-pro-via-a-usb-cable-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>Best 3 Software to Transfer Files to/from Your HTC U23 Pro via a USB Cable | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/best-3-software-to-transfer-files-tofrom-your-tecno-pop-8-via-a-usb-cable-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>Best 3 Software to Transfer Files to/from Your Tecno Pop 8 via a USB Cable | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/easiest-guide-how-to-clone-motorola-moto-g24-phone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>Easiest Guide How to Clone Motorola Moto G24 Phone? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-i-transferred-messages-from-realme-10t-5g-to-iphone-12xs-max-in-seconds-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How I Transferred Messages from Realme 10T 5G to iPhone 12/XS (Max) in Seconds | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-copy-data-from-samsung-to-pixel-8-and-pixel-8-pro-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Copy Data from Samsung to Pixel 8 and Pixel 8 Pro | Dr.fone</u></a></li>
<li><a href="https://fake-location.techidaily.com/how-to-fake-gps-on-xiaomi-redmi-note-13-5g-for-mobile-legends-drfone-by-drfone-virtual-android/"><u>How To Fake GPS On Xiaomi Redmi Note 13 5G For Mobile Legends? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-apps-from-oppo-reno-11-5g-to-another-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Apps from Oppo Reno 11 5G to Another | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-apps-from-poco-c55-to-another-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Apps from Poco C55 to Another | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-contacts-from-nokia-c02-to-other-android-devices-devices-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Contacts from Nokia C02 to Other Android Devices Devices? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-contacts-from-oppo-a59-5g-to-outlook-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Contacts from Oppo A59 5G to Outlook | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-contacts-from-poco-c50-to-other-android-devices-using-bluetooth-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Contacts from Poco C50 to Other Android Devices Using Bluetooth? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-contacts-from-realme-c55-to-outlook-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Contacts from Realme C55 to Outlook | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-transfer-contacts-from-tecno-phantom-v-fold-to-iphone-xs11-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>How to Transfer Contacts from Tecno Phantom V Fold to iPhone XS/11 | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/in-2024-how-to-teleport-your-gps-location-on-infinix-smart-8-pro-drfone-by-drfone-virtual-android/"><u>In 2024, How To Teleport Your GPS Location On Infinix Smart 8 Pro? | Dr.fone</u></a></li>
<li><a href="https://some-approaches.techidaily.com/in-2024-the-prime-camera-choices-revealing-the-best/"><u>In 2024, The Prime Camera Choices – Revealing the Best</u></a></li>
<li><a href="https://tech-renaissance.techidaily.com/mastering-the-art-of-managing-pages-file-formats-on-pc-systems/"><u>Mastering the Art of Managing Pages File Formats on PC Systems</u></a></li>
<li><a href="https://buynow-tips.techidaily.com/motorolas-mid-range-marvel-the-one-hyper-reviewed-striking-features-with-robust-performance/"><u>Motorola's Mid-Range Marvel: The One Hyper Reviewed - Striking Features with Robust Performance</u></a></li>
<li><a href="https://facebook-video-share.techidaily.com/pixel-powerhouses-the-elite-females-on-yt-for-2024/"><u>Pixel Powerhouses  The Elite Females on YT for 2024</u></a></li>
<li><a href="https://win-blog.techidaily.com/resolving-the-microsoft-edge-crashes-on-windows-11-tips-and-tricks-applied/"><u>Resolving the Microsoft Edge Crashes on Windows 11: Tips and Tricks Applied</u></a></li>
</ul></div>

---
title: How to fix runtime error 424 object required error in Excel 2023 | Stellar
date: 2024-05-19T18:32:11.917Z
updated: 2024-05-20T18:32:11.917Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes How to fix runtime error 424 object required error in Excel 2023
excerpt: This article describes How to fix runtime error 424 object required error in Excel 2023
keywords: repair corrupt .xlsm files,repair excel 2021,repair excel 2010,repair damaged .xltx,repair damaged .csv,repair damaged .csv files,repair damaged .xls files
thumbnail: https://www.lifewire.com/thmb/qNtT0Vit1jn8RftUFX6Xcsv5x4o=/400x300/filters:no_upscale():max_bytes(150000):strip_icc():format(webp)/5-best-texting-apps-for-android-tablets-in-2023-a5a6329d6b744167acf0a02f6a413ff2.jpg
---

## How to fix runtime error 424 object required error in Excel

The Runtime error 424: Object required occurs when Excel is not able to recognize an object that you are referring to in a VBA code. The object can be a workbook, worksheet, range, variable, class, macro, etc. Some users have also reported that this error occurred when they tried to copy the values of the cells from one workbook to another.

Let’s understand the error through a small scenario. Suppose, I want to check the last field row in a table in a spreadsheet named “First” using the VBA code. To do this, I have added a command button and double-clicked on it and entered the below code in the backend:

Private Sub CommandButton2\_Click()

Dim LRow As Integer

LRow = Worksheets("First").Cells(Rows.Count, 2).End(xlUp).Row

MsgBox ("Last Row " & LRow)

End Sub

![Code On Command Button](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/code-on-commandbutton.jpg)

In this code, _Worksheets("First")_ is a data object. If I mistakenly delete this data object and insert any random name (for example - kanada), then it will not be recognized by Excel. When I run this code, I will get the “Run-time error 424”.

![Runtime Error with scenario](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/run-time-error-424-with-scenario.jpg)

## **Causes of Runtime Error 424 in Excel**

The Runtime error 424: Object required can occur due to the following reasons:

- Incorrect name of the object you are trying to refer to in a code.
- You have provided an invalid qualifier to an object.
- You have not used the Set statement while assigning an object reference.
- The object is corrupted.
- Missing objects in a workbook.
- Objects you are trying to call in a code are mistakenly deleted or unavailable.
- You have used an incorrect syntax for object declaration.
- You are trying to perform an invalid action on an object in a code.
- Workbook is corrupted.

##  **Solutions to Fix Runtime Error 424: Object Required in Excel**

The VBA error ‘object required’ may occur due to different reasons. Based on the reason, you can follow the solutions mentioned below to fix the error.

### **1\. Check the Name of the Object**

The Runtime error 424 can occur when you run the VBA code using an incorrect name of the object. For example, the object name is ‘MyObject’ but you’re using “Backcolor”.

![Error When Incorrect Name Of The Object](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/error-424-when-incorrect-name-of-the-object.jpg)

 When you click the **Debug** button, the line with the error will highlight.

![Highlighting Line With Error](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/highlighting-line-with-error.jpg)

To fix the issue, you need to provide the correct name of the object.

### **2\. Check if the Object is Missing**

 The Runtime error 424 can occur if the object you are referring to as a method is not available or you are using the wrong object in a code. In the below example, you can see that the error occurs when an object named “Employee” is not available in the Project list.

![Example Of Code When Object Is Not Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/example-of-code-when-an-object-is-not-available.jpg)

 You can check and mention the object which is available. For instance, Sheet2 in the below code.  

![Check When The Object Is Available](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/check-when-the-object-is-available.png)

### **3\. Check All References are Declared in the Code**

You can get the Runtime error 424 if all the references are not declared. So, make sure you have declared all the references in the code. To verify this, you can use the debug mode by pressing **F5** or clicking on the **Debug** option.

![Debug Command In Excel](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/debug-command-in-excel.jpg)

### **4\. Check the Macro Security Settings**

Sometimes, the error can occur if macros are disabled in the Macro Security settings. You can check and change the settings by following these steps:

- On the **Developer** tab, in the **Code** section, click **Macro Security**.
- In the **Trust Center** window, select **Enable all macros.**

![Macro Security Wizard](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/macro-security-wizard.jpg)

- Click **OK**.

![Enable All Macro In Trust Center](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Error424/enable-all-macros-in-trust-center.jpg)

### 5\. Repair your Workbook

Sometimes, the ‘Object required’ error can occur if your Excel file is damaged or corrupted. In such a case, you can try repairing the file using Microsoft’s in-built utility - Open and Repair. To use this utility, follow these steps:

- In Excel, go to **File > Open > Browse**.
- In the Open dialog box, click on the corrupted Excel file.
- Click the arrow next to the Open button and select **Open and Repair** from the dropdown.
- Select **Repair** to recover as much data from the file as possible.

If the Open and Repair utility fails or stops working, then you can try a professional Excel repair tool, such as [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). It is an advanced tool that can repair severely corrupted Excel files **(.xls, .xlsx, .xltm, .xltx, and .xlsm)**. It helps recover all the file components, including images, charts, tables, pivot tables, cell comments, chart sheets, formulas, etc., without impacting the original structure.

## **Conclusion**

The Runtime error 424 usually occurs when there is an issue with the objects in your VBA code. In this article, we have covered some effective methods to resolve the “object required” error in Excel. If the error occurs due to corruption in Excel file, then you can repair the corrupt file using Stellar Repair for Excel. It is a reliable tool that can repair severely corrupted Excel file without changing its actual formatting. You can download the free trial version of the software to evaluate its functionality.


## Get Rid of corrupt Excel File

**Summary:** What to do when an Excel file is corrupted? This is a common question that is often asked by Microsoft Excel users. If you too are seeking an answer to this question, read the blog to learn about a few manual workarounds and a specialized Excel file repair tool to resolve the Excel file corruption issue.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

An Excel file gets corrupted due to various reasons such as a virus/malware attack, sudden system shutdown when the Excel file is still open, power failure while working with an Excel spreadsheet, etc.

When Microsoft Excel detects corruption in a workbook, it attempts to repair the workbook by starting _‘File Recovery mode’._

**Tip!** If the file recovery mode doesn’t start, you may use the manual repair process or an Excel repair tool, such as Stellar Repair for Excel to repair a corrupted Excel file. The software can help you quickly retrieve contents from a damaged, corrupt, or inaccessible Excel file and restore the file to its original state.

There even exist a few manual tips that can be used to recover data from damaged MS Office Excel files.

## **Workarounds to Use When an Excel File is Corrupted**

**_Note:_** _Before carrying out any of the repair and recovery workarounds, it is advised that you must save a backup copy of the damaged file. This is to prevent your files from turning completely inaccessible in case the methods fail to give desired results._

### **Workaround 1: Use the Open and Repair Method**

If MS Excel cannot repair a corrupted workbook automatically, you can try to do it manually. To do so, perform the following:

- Open the corrupt file, like you normally open any file, by clicking **File** > **Open**.
- Browse and locate the folder containing the corrupted document.
- When the Open dialog box is displayed:
- Select the Excel document.
- Click on the arrow present to the right side of the Open button and select **Open and Repair** option.

![open and repair excel file](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/07/open-and-repair.jpg)

Figure 1 – Open and Repair Feature

If this doesn’t help repair the broken Excel file or you encounter  [Open and Repair does not work](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) issue, proceed with the next workaround.

**_Tip!_** _Try an alternative solution, i.e._ **_Stellar Repair for Excel software_** _to repair and recover corrupt Excel files (.xlsx or .xls) when the ‘Open and Repair’ method won’t work._  

### **Workaround 2: Restore an Excel File with a Shadow Copy**

If you’re a **Windows 7 or Vista user,** you can try restoring the corrupted spreadsheet by using a shadow copy (or a previous version). [Shadow copy](<https://en.wikipedia.org/wiki/Shadow_Copy#:~:text=Shadow%20Copy%20(also%20known%20as,the%20Volume%20Shadow%20Copy%20service>.) is basically a snapshot (backup copy) of computer files or volumes. The snapshot may contain an older version of your Excel file that has become damaged now. To find out, do the following:

- **Launch File Explorer**, and right-click the folder in which the **file is saved.**
- Choose **Properties.**
- Look for and click the **Previous Versions** tab. This will display a list of entries under **Folder versions** or **File versions**, going back a few days or weeks.
- Double-click one with a date when the file was accessible and could be read. Then, try to open its older version. If it opens, save the older version with a new name and execute the procedure with new file/folder entries.

![Excel file is corrupted](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2017/09/Shadow-copy.png)

Figure 2 – Volume Shadow Copy

You would have to repeat the process until you reach the point where the file became damaged. With this, you will get a baseline version of the file, but data may still have been lost.

### **Workaround 3: Test your Assumptions**

If you receive a message saying “[**Excel file corrupted and cannot be opened**](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)”, you would probably believe it. However, there could be other reasons besides corruption that may cause Excel to throw this error message.

Your Office suite, which Excel is a part of, maybe having some primary issues in it causing problems while opening one Excel document. So, try opening another Excel file to check if the problem exists with all the files or just one.

If other Excel documents work correctly, it means that only the particular document is corrupt. On the contrary, if the issue is with your Office suite, repairing the current Office installation may help fix the issue. For this, perform these steps:

- Go to **Control Panel and click Uninstall** **the Program**.
- Choose **Office.**
- Click Change, and hit the **Repair button.**

You can **reinstall** the entire Office package. Once reinstalled, try to open the file to check if the issue has been fixed and the **Excel file repaired.**

![Excel file is corrupted](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2017/09/Repair-MS-Office.jpg)

Figure 3 – MS Office Repair

### **Workaround 4: Use Excel File Repair Tool**

If the above manual solutions fail, use [Excel repair software](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) to successfully repair your damaged Excel workbook and recover all its data. Essentially, the software rebuilds damaged Excel workbook data at a granular level to recover every single object & all the original properties of the workbook.

**Suggested Read:** [**How to repair corrupt Excel files using Stellar Repair for Excel?**](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

## **Why Use Stellar Repair for Excel Software?**

- Repairs severely corrupted XLSX and XLS files.
- Can handle corrupt Excel files of any size.
- Demo version allows previewing recoverable Excel file items for free.
- **Supports Microsoft Excel 2019** and all lower versions.
- Compatible with **Windows 10 and lower versions.**
- Tested and recommended by **Microsoft Excel MVPs.**

[![Free Download for Windows](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/01/Free-download-for-windows-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

## **Final Word**

When an Excel file is corrupted, it won’t open at all or you won’t be able to access all the file data. Such a situation can lead to unnecessary halts, impacting work productivity.

There are manual workarounds that may help fix the corrupt Excel file and recover its data, such as the ones covered in this blog. However, these solutions might not work in severe corruption cases and may require technical assistance. Also, they may result in some data loss.

To overcome the limitations of manual workarounds, it is recommended to go for a professional Excel file repair tool such as [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). It helps **repair corrupt Excel** (XLS or XLSX) files and restores all worksheet data, such as the table, chart, chart sheet, cell comment, sort and filter, image, formula, etc. in a few simple clicks. Moreover, the software provides a free preview of the recoverable data with its demo version. You can check the preview to evaluate how the software works.


## \[Solved\] Excel Spreadsheet Disappears after Opening

Several Excel users have reported about experiencing ‘Excel spreadsheet disappears after opening’ issue. The problem occurs when attempting to open an Excel file by double-clicking on the file icon or name. The Excel file opens blank grey screen.

![Excel blank screen](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img1.jpg)

Figure 1 - Excel Blank Screen

## **User Instances of ‘Excel Open But Can’t See Spreadsheet’ Issue**

**Instance 1:** The user said that _“Excel 2016 opens to a blank screen and everything is greyed out. Ribbons at the top of Excel are also not present”._

**Instance 2:** The user reported that _“Attempting to open a spreadsheet, either by double-clicking the file or by opening Excel, it opens but only a blank Excel window is visible. When trying to save that particular file, Excel behaves as though there is no file open at all and all the file saving options, such as save, save as, print, etc. are greyed out.”_

## **Causes Behind ‘Excel Open But Can’t See Spreadsheet’ Issue & the Solutions Thereof**

Following are some of the possible reasons behind the ‘Excel open but can’t see spreadsheet’ issue, along with their solutions:

### **Cause 1: Excel File is Hidden**

You may have saved the Excel sheet as a hidden document.

### **Solution: Unhide Excel File**

Verify if your Excel worksheet is hidden by following these steps:

- In Excel, click the **View** tab, and then click **Unhide**.

Note: If the ‘Hide’ tab under the View menu is greyed out, the sheet you’re trying to open is not hidden. In that case, proceed to the next workaround.

![unhide excel file](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img2.jpg)

Figure 2 - Unhide Excel File

- A dialog box will pop-up showing the worksheet name that can be unhidden. For instance, below is an image of Unhide box that lists ‘Daily\_Reports’ spreadsheet under Unhide workbook.

![unhide excel workbook](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img3.jpg)

Figure 3 – Select and Unhide Excel Workbook

- Click **OK** to unhide your Excel sheet.

If this fails to work, there’s a possibility that the spreadsheet window pane may have slided to one side of the visible desktop. To bring back the displayed area, click the ‘Arrange All’ option under the View tab. If the issue persists, try the next solution.

### **Cause 2: Ignore Dynamic Data Exchange (DDE) Option is Enabled**

Another reason behind the ‘Excel opens to a blank screen’ issue could be that the “Ignore other applications that use Dynamic Data Exchange (DDE)” checkbox is checked in Excel options.

Usually, when you double-click an Excel file, a DDE message is sent to Excel instructing it to open that particular file. But, if you have the ‘Ignore DDE’ option selected, the DDE message sent to Excel to open a workbook is ignored. As a result, the Excel workbook opens a blank screen.

### **Solution: Uncheck the Ignore DDE Option**

Make sure that the ‘Ignore DDE’ option is unchecked by performing these steps:

Note: Skip these steps for Excel 2019 and Excel Office 365.

- In your Excel window, click **File** > **Options**.

![Excel options](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img4.jpg)

Figure 4 - Select Excel Options

- From the left-side of the ‘Options’ window, choose **Advanced**.

![Excel Options Window](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img5.jpg)

Figure 5 - Excel Options Window

- In ‘Advanced’ window, locate the **General** section, and then uncheck the “Ignore other applications that use Dynamic Data Exchange (DDE)” checkbox.

![uncheck ignore DDE](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img6.jpg)

Figure 6 - Uncheck Ignore Other Applications that use DDE

- Click **OK** to apply the changes.

If these steps don’t work, proceed to the next workaround.

### **Cause 3: Issue within Office Program**

Sometimes, problem within your Microsoft Office program might cause Excel to behave oddly and may result in an Excel spreadsheet disappeared issue.

### **Solution: Repair Office Program**

Try repairing your Office program by executing the steps listed below, based on your Windows OS:

**For Windows 10:**

- Type Settings in the Windows search box.
- Click **Settings**.
- In ‘Windows Settings’ screen, select **Apps**.

![Windows settings screen](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img7.jpg)

Figure 7 - Windows Settings Screen

- In ‘Apps & features’ screen, scroll down to your Microsoft Office program and tap on it, and then click the **Modify**

![modify MS Office program](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img8.jpg)

Figure 8 - Modify MS Office Program

- In ‘How would you like to repair your Office programs’ dialog box, select the **Online Repair**radio button, and then click the **Repair**

![repair MS Office program](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img9.jpg)

Figure 9 - Repair MS Office Program

**For Windows 8:**

- Open Control Panel and click the **Uninstall a program** option under **Programs**.
- Select **Microsoft Office 365** and then click **Change**.
- In the window that opens, select **Online Repair** and then hit the **Repair**

You may be asked to restart your system after completion of the repair process.

**For Windows 7:**

- In ‘Control Panel’ window, double-click **Programs and Features**.
- Under **Uninstall or change a program** section, select the Office program, and then select **Change**.
- In the dialog box that appears, choose **Online Repair** and then choose **Repair**.

If repairing your Office installation fails to resolve the issue, continue to solution 4.

### **Cause 4: Problematic Excel and COM Add-ins**

Sometimes, Excel and COM add-ins may cause Excel file open a blank grey screen problem.

### **Solution: Disable the Add-ins**

Disabling the Excel and COM add-ins one at a time may help resolve the problem. The steps are as follows:

- Click the **File** menu in your Excel file and select
- From ‘Excel Options’ left-side panel, click **Add-Ins**, and then choose **COM Add-ins** from the ‘Manage’ drop-down. Click the **Go**

![COM Add-ins](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img10.jpg)

Figure 10 - Select COM Add-ins

- Uncheck one of the add-ins checkbox from the COM Add-ins window, then select **OK**.

![disable COM Add-ins](https://www.stellarinfo.com/public/image/catalog//article/file-repair/Excel-repair/img11.jpg)

Figure 11 - Uncheck and Disable COM Add-ins

- Restart your Excel program.

If the issue persists, repeat the above steps (1 till 4), except that you need to choose a different add-in in step 3.

If you’re still experiencing the same problem after unchecking all the COM add-ins, repeat all the above steps, except choose ‘Excel Add-ins’ in step 2.

If Excel can load the file, then the add-in that you disabled last leads to the problem. If turning off add-ins does not work, try the next solution.

### **Cause 5: Problem with Excel File Associations**

The Excel file opening a blank screen problem may occur if the file associations are not performing correctly.

### **Solution: Reset Excel File Associations**

Try resetting the file associations in Excel to their default settings. For this, follow the below steps based on your OS.

**For Windows 10 and Windows 8.1:**

- Select the file that is opening incorrectly and copy it to the desktop.
- Right-click the file and click **Properties**.
- In the File's 'Properties' window, check your file type next to **Type of File**from the General tab. For instance, (.docx), (.csv), or (.pdf).
- Next, check to which app your file is associated with from the **Opens with**option

If the file type is different than .xlsx, open the file in a different application by following these steps:

- Click the **Change button next to the ‘Opens with’ option.**
- Click **More apps**.
- From the list of applications, select the desired app and then check the **Always use this app** Click **OK**.

**For Windows 8:**

- Open **Control Panel**, click **Default Programs**, and then select **Set your default programs**.
- Choose **Excel** and then tap the **Choose default for this program**
- In ‘**Set Program Associations**’ window, choose **Select All** and then click **Save**.

**For Windows 7:**

- In **Control Panel**, choose **Default Programs**.
- Click the **Associate a file type or protocol with a specific program**
- Choose **Microsoft Excel Worksheet** and click on change program.
- Choose **Microsoft Excel** under **Recommended Programs**.
- If you’re unable to find Excel, browse the Excel installation folder.
- Select **exe** and then choose **Excel**.

## **What Next? Use Stellar Repair for Excel to Recover Your File**

If the above solutions don’t help recover your Excel file, use an Excel file repair tool, such as [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) to recover it. The software can restore the Excel spreadsheet and its components, including tables, pivot tables, charts, formulas, etc. while preserving the spreadsheet properties and formatting.

[![Free download Stellar Repair for Excel](https://www.stellarinfo.com/blog/wp-content/uploads/2017/02/free-download-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

Check out the video to recover Excel file by using Stellar Repair for Excel software.

<iframe src="//www.youtube.com/embed/VAeGzHnETu0" width="640" height="360" frameborder="0"></iframe>

## **Conclusion**

This article described the possible causes behind the ‘Excel open but can’t see spreadsheet’ issue, along with their solutions. The manual solutions to resolve the issue require time and efforts. But, if you need to access your Excel file without any delay, use Stellar Repair for Excel software to recover your .xlsx/.xls file in just a few clicks.


## Quick Fixes to Repair Microsoft Excel 2013/2016 Content related error

**Summary:** The blog outlines some quick tips to fix ‘We found a problem with some content’ error in Microsoft Excel 2013/2016. It explains manual procedure to resolve the error and also suggests an automated tool to perform the repair process to retrieve all possible data from a corrupt workbook.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Sometimes, when opening an MS Excel file, you may receive an error message that reads:

“**We found a problem with some content in ‘filename.xlsx’. Do you want us to try to recover as much as we can? If you trust the source of this workbook, click Yes.**“

![Microsoft Excel Content Error](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/07/Microsoft-Excel-Content-Error.jpg)

Figure 1 – Excel ‘found a problem with some content’ Error Message

## **What Causes ‘We Found a Problem with Some Content’ Error?**

There is no clear answer as to what results in the Excel error – ‘**We found a problem with some content in <filename.xlsx>**’. However, based on some user experiences, it appears that the error occurs due to corruption in an Excel workbook. It may turn corrupt when:

- You try opening the Excel file saved on a network-shared drive.
- A string is added in a cell in Excel, instead of a numeric value.
- Text values in formulas exceed 255 characters.

## **How to Resolve ‘We Found a Problem with Some Content’ Error?**

**Follow these tips to fix the Excel error:**

**IMPORTANT!** Before you follow the tips to resolve the Excel error, keep these points in mind: Make sure you have closed all of the opened Excel workbooks. Try restoring Excel file data from the most recent backup copy. If you don’t have a backup copy, make a copy of the corrupt Excel file and perform repair and recovery procedures on that backup copy.

### **Tip #1: Repair Corrupt Excel File**

File Recovery mode is a native Excel recovery utility that automatically opens whenever any inconsistencies are found in the worksheet. If Microsoft doesn’t detect any issue or fails to open the File Recovery mode, you can start it manually to recover the corrupt Excel file. To do so, follow the steps below:

1. Click on the **File** menu, and then select **Open**.
2. In the **Open** dialog box, navigate to the folder location where the corrupt Excel file is saved.
3. Select the corrupt file, and then click on arrow sign available next to **Open** button to select **Open and Repair** option.

![Open and Repair](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2017/03/Open-and-Repair.png "MS Excel Content error")

Figure 2 – Open and Repair Feature in Excel

1. Next, click **Repair to recover maximum possible data**.
2. If the repair is not able to recover the data from the workbook, select **Extract Data** to extract all possible formulas and values from the workbook.

If repairing the corrupt Excel file doesn’t work, you can try an Excel file repair tool to fix corruption errors. You can also try to recover data from the corrupt file manually by following the next tips.

**Read this:** [What to do when Open and Repair doesn’t work?](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

### **Tip #2: Set Calculation Option to Manual**

To make the file accessible, try setting the calculation option in Excel from automatic to manual. As a result, the workbook will not be recalculated and may open in Excel. For this, perform the following:

1. Click **File,** and then click **New**.
2. Under **New**, click the **Blank workbook** option.
3. When a blank workbook opens, click **File** > **Options**.
4. Under the Formulas category, pick Manual in the **Calculation options** section, and then click **OK**.

![calculation options](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2017/03/calculation-options.png "MS Excel Content error")

Figure 3 – Select Manual in Calculation options

1. Now, again click on the **File** menu and then click **Open**.
2. Navigate to the corrupt workbook, and double-click it.

When the workbook opens, check if it contains all the data. If not, proceed to the next tip.

### **Tip #3: Copy Excel Workbook Contents to a New Workbook**

Several users have reported that they were able to fix ‘_We found a problem with some content in <filename>’_ error message by copying contents from the corrupt workbook to a separate workbook. **Detailed steps are as follows**:

1. Open the Excel workbook in **‘read-only’** mode, and copy all its contents.
2. Create a blank new workbook and paste the copied contents from the corrupt file to the new file.

### **Tip #4: Use External References to Link to the Damaged Workbook**

Use external references to link to the corrupted workbook. By implementing this fix, data contents can be retrieved. However, it is not feasible to recover formulas or calculated values using this solution.

**Follow the steps below:**

1. In Excel 2013/2016, click **File** > **Open**.
2. Navigate to the **folder** where the corrupt file is **saved**.
3. Right click the file, select **Copy,** and then click on **Cancel**.
4. Again, click on **File** and then **New**.
5. Under **New** option, click on **Blank workbook**.
6. In the **cell A1** of new workbook, type **\=File Name!A1** (where File Name indicates the name of the damaged workbook being copied in **Step 3**).
7. If **Update Values** dialog box appears, click the corrupt workbook, and choose **OK**.
8. If **Select Sheet** dialog box appears, click the appropriate sheet, and then click **OK**.
9. Select cell **A1**.
10. Next, click **Home,** and then click **Copy** (or, press Ctrl +C).
11. Starting in **cell A1**, select area approximately the same size as that of the cell range that contains data in the damaged workbook.
12. Next, click **Home** and select **Paste** (or click Ctrl + V).
13. Keep the range of cells selected, click **Home** and then **Copy**.
14. Finally, click on **Home**, click on the arrow associated with **Paste** and under **Paste Values** click on **Values**.

This will remove the link to the corrupt workbook and will retrieve data. But, keep in mind, the recovered data will no longer contain formulas or calculated values.

## **Alternative Solution – Stellar Repair for Excel**

If the above manual methods fail to fix the ‘We found a problem with some content in Excel error’, try using the Stellar Repair for Excel software to resolve this error. The software helps repair and recover corrupt Excel files in just a few clicks. It can be used on a Windows 10/8/7/Vista/XP/NT machine to repair a corrupted workbook and recover every single bit of data from all the versions of the Excel workbook.

[![Free Download for Windows](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/07/free-download-1-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

**Read this:** [How to repair corrupt Excel file using Stellar Repair for Excel?](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

## **Conclusion**

In this blog, we discussed some possible reasons behind Microsoft Excel 2013/2016 _‘We found a problem with some content’_ error. The error may occur when an Excel file becomes corrupt. You may try repairing the corrupted Excel file manually by using the built-in ‘Open and Repair’ feature. Or, try the manual workarounds to extract data from the corrupt file discussed in this post. If the manual solutions don’t work for you, using [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) can come in handy in repairing the corrupt Excel (.xls/.xlsx) file and recovering the complete file data.


## \[Fixed\] Excel Found a Problem with One or more Formula

**Summary:** The error ‘Excel found a problem with one or more formula references in this worksheet’ may appear while saving the Excel workbook. It occurs when Excel found a problem with the formula used in the sheet. However, it may also occur when the Excel workbook gets damaged or corrupt. In this guide, we’ve explained the reasons that may lead to this Excel error and methods to resolve the error, by using various Excel options and a third-party Excel file repair software.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

If you are experiencing the ‘Excel found a problem with one or more formula references in this worksheet’ error message in the Excel workbook, it indicates that the [Excel file is corrupt](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) or partially damaged. However, it may also occur due to incorrect reference to a wrong cell or object linking, which is not working. The complete error message says,

_‘Excel found a problem with one or more formula references in this worksheet. Check that the cell references, range names, defined names, and links to other workbooks in your formulas are all correct.’_

![Excel found a problem with one or more formula references](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/11/MS-Excel-problem-with-formula-reference.png)

In any case, resolving the error is critical as it doesn’t let you save the file and may result in loss of information from the Excel workbook.

## Reasons for Excel Formula References Error

A few reasons that may lead to such error are as follows,

- Wrong formula or reference cell
- Incorrect object linking or link embedding OLE
- Empty or no values in named or range cells
- Multiple Excel files (not common)

## Methods to Resolve ‘Excel Found a Problem with One or More Formula References in this Worksheet’ Error

Following are a few methods that you can follow to [fix Excel file](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) that can’t be saved due to problems with one or more formula references in the worksheet.

### Method 1: Check Formulas

If the problem has occurred in a large Excel workbook with multiple sheets, it’s quite hard to pinpoint the problem cell. In such cases, you can use the Error Checking option that runs a scan and checks for a problem with formulas used in the worksheet.

To run Error Checking in the Excel sheet, follow these steps,

- Go to Formulas and click on the ‘Error Checking’ button

![Error Checking](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/11/MS-Excel-error-checking-1024x431.png)

- This runs a scan on the sheet and displays the issues, if any. If no issue is found, it displays the following message,

_The error check is completed for the entire sheet._

In such a case, you can try saving the Excel file again. If the error message persists, proceed to the next method.

### Method 2: Check Individual Sheet

The problem may also occur due to an issue with one of the sheets in the workbook. To find the faulty sheet and fix the problem, you can copy each sheet content in a new Excel file and then try to save the Excel file.

This will help you find the faulty sheet from the workbook that you can review. This method makes the entire [process of troubleshooting Excel formula](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) reference error quite easy and convenient.

In case the error is not fixed, you can back up the faulty sheet content and remove it from the workbook to save the Excel file.

### Method 3: Check Links

When the Excel file contains external links with errors, MS Excel may display such error messages. To check and confirm if external links are causing the error, follow these steps,

- Navigate to _Data Tab > Queries & Connections > Edit Links_
- Check the links. If you find any faulty link, remove it and then save the sheet

### Method 4: Review Charts

You can review the charts to check if they are causing the formula reference error in Excel. It may take a while based on the size of the Excel file. Sometimes, it’s not practically possible to track down which Excel chart object is causing the error. Thus, you need to check specific locations, such as:

1. Check horizontal axis formula inside Select Data Source dialog box
2. Check Secondary Axis
3. Check linked Data Labels, Axis Labels, or Chart Title

### Method 5: Check Pivot Tables

To check Pivot Tables, follow these steps,

- Navigate to _PivotTable Tools > Analyze > Change Data Source > Change Data Source…_

![Edit links](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/11/MS-Excel-Edit-Links-1024x84.png)

- Check if any of the formula used is problematic. Sometimes small typo, such as misplaced comma, can lead to such problems in Excel. Thus, check each formula thoroughly and correct the formulas wherever needed.

### Method 6: Use Excel Repair Software

When none of the methods resolve the error, then you can rely on advanced [Excel repair software](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/), such as Stellar Repair for Excel. It’s a powerful tool that is recommended by several MVPs and IT administrators for resolving common Excel errors, such as ‘Excel found a problem with one or more formula references in this worksheet.’

![Stellar Repair for Excel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/public/image/catalog/screenshot/excel-repair/stellar-repair-for-excel-main-interface.png)

It repairs corrupt or damaged Excel (.xls/.xlsx) files, recovers Pivot tables, charts, etc., and save them in a new Excel worksheet. It helps Excel users, facing formula reference error, restore their Excel file without any risk of data loss, while preserving the sheet properties and formatting with 100% precision.

## Conclusion

Although the error ‘Excel found a problem with one or more formula references in this worksheet’ can be resolved by using various options in MS Excel, it may lead to a partial loss of information. Thus, you must perform these operations after taking a backup of the Excel worksheet. Also, if the MS Excel options fail to resolve the problem, you can use an Excel file repair software, such as Stellar Repair for Excel. The software helps fix Excel file corruption and restores the information and data from corrupt or damaged Excel files (.xls/.xlsx) to a new worksheet.




## How to Fix the Unable to Record Macro Error in Excel?

**Summary:** You may encounter the “Unable to record macro” error in MS Excel when using Personal Macro Workbooks. In this post, we’ll discuss the possible causes behind this error and the ways to fix it. We’ll also mention a professional Excel repair tool that can help fix the error if it occurs due to corrupted workbook.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

A personal macro workbook (Personal.xlsb file) is a hidden workbook in Excel that stores all macros in a single workbook. This allows you to automate processes while working in Excel. Sometimes, when recording macro codes in the personal macro workbook, you may get the message “**Personal Macro Workbook in a startup folder must stay open for recording**”. When you click on the **OK** button, it will show the “unable to record” error. This prevents you from recording the macros. Below, we’ll see the causes behind this error and discuss how to resolve this error.

## **Causes of Unable to Record Macro Error**

You may be unable to record macros in Excel due to several reasons. Let’s take a look at the possible causes that can lead to this issue.

- The location of personal.xlsb file is changed.
- Personal.xlsb file is corrupted.
- Macros are disabled.

## **Methods to Fix the “Unable to Record Macro” Error in Excel**

Here are some possible solutions that can help you resolve the unable to record macro error in Excel.

### Method 1: Check the Path of XLStart Folder

You may be unable to record macros if the path of XLStart folder is incorrect. It is a folder where the Personal.xlsb file is stored by default. Follow these steps to find out the path of this folder:

- Open MS Excel. Go to **File > Options**.
- Click **Trust Center > Trust Center Settings**.

![Excel Options Window](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/excel-options-window.png)

- In the **Trust Center Settings** window, click on **Trusted Locations**.

![Path Of XLStart Folder In Trust Center](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/path-of-xlstart-folder-in-trust-center.png)

- Verify the path of the **XLSTART** folder and modify it if there is an issue.
- Once you are done, click on **OK**.

### Method 2: Change Macro Security

The “Unable to record macro” error can occur if macros are disabled in the Macro Security settings. You can try changing the macro settings using the below steps:

- In MS Excel, go to **File > Options > Trust Center**.

![Excel Options To Locate Trust Center](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/excel-options-to-locate-trust-center.png)

- Under **Trust Center,** click on **Trust Center Settings**.

![Change Macro Settings In Trust Center](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/change-macro-settings-in-trust-center.png)

- Select **“Enable all macros”** and then click **OK.**

### Method 3: Check Add-ins for Disabled Items

If there are any items in add-ins that are disabled, they may prevent Excel from functioning properly. You can check and enable the items in MS Excel using the below steps:

- Click **File > Options.**

![Go To Options](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/go-to-options-1.png)

- In **Excel Options**, click on the **Add-ins** option.
- Select **Disabled Items** from the **Manage** section and click on **Go**.

![Add-ins In Excel Options](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/add-ins-in-excel-options.png)

- The **Disabled Items** window is displayed.
- Click on the disabled item and then click **Enable**.
- Restart Excel for the changes to take place.

### Method 4: Repair your Excel File

You may fail to record macros if there is corruption in the workbook. In such a case, you can use the “Open and Repair” utility in MS Excel to repair the corrupt workbook. To use this tool, follow these steps:

- Open your Excel application.
- Click **File > Open**.
- Browse to the location where the corrupted file is stored.
- In the **Open** dialog box, choose the corrupted workbook.

![Open Dialog Box](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/open-dialog-box.png)

- From the **Open** dropdown list, click **Open and Repair**.

![Open And Repair Window](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/open-and-repair-window.png)

Excel will prompt you to repair the file or extract data from it. Click **Repair** to retrieve maximum data. If the Repair option fails, then click on the Extract Data option to recover the data without formulas and values.

If the Microsoft utility “Open and Repair” fails to repair the corrupted Excel workbook, then try a professional Excel repair tool such as Stellar Repair for Excel. It is an advanced tool that can easily repair severely corrupted Excel (XLSX and XLS) files. It can recover all the file items, including chart sheets, cell comments, tables, macros, formulas, etc. without impacting the properties and cell format of the Excel file.

## **Closure**

You may receive the “unable to record” error in Excel while creating or storing macros in Personal Macro Workbooks. There are several reasons that can lead to this error. You can try the methods covered in this post to resolve the error. If the error appears due to corruption in workbook, then try to repair it using the Open and Repair utility. Alternatively, you can use [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) – a professional tool that can help you repair an Excel file with problematic macros. Also, it allows recovery of all the file components with complete integrity. The tool is compatible with Excel 2021, 2019, 2016, and older versions.



## Fixed "Cannot Insert Object" Error in Excel | Step-by-Step Guide

**Summary:** The error “cannot insert object” in MS Excel can prevent you from modifying objects in the worksheet. This blog will discuss the primary reasons behind this error and the possible solutions to fix it. You will also learn about a professional Excel repair software that can help fix the error if it has occurred due to corruption in Excel file.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Many users have reported encountering the “cannot insert object” error while adding/embedding objects into the Excel file. It usually occurs when using Object Linking and Embedding (OLE) to add content (PDF, Microsoft documents) from external applications to worksheet. The error can also occur when using ActiveX control in Excel. Below, we’ll explain why you cannot insert object into Excel sheet and how to troubleshoot the issue.

## **Why the “Cannot Insert Object” Error Occurs?**

- Macro Settings can prevent the insertion of objects into a workbook.
- The Excel file in which you are trying to add an element is corrupted.
- The object (you are inserting into the workbook) is damaged.
- Object size limitations.
- System’s insufficient memory might prevent new objects’ addition.
- Incompatible Excel file format.
- Add-ins controls are disabled.
- Incompatible or faulty Add-ins.
- Issue with Security Settings.

## **Methods to Fix the “Cannot Insert Object” Error in Excel**

You may encounter the “Cannot insert object” error when trying to add an element stored on a network. It can occur due to issues with the file link, such as incorrect file location. In such a case, you can check the link by selecting the **link to file** option from the **Insert** tab.

Sometimes, the error can occur if the file in which you are trying to insert the object is locked and password-protected. In this case, you can [unprotect the Excel file](https://support.microsoft.com/en-au/office/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6). If the issue still persists, then you can follow the below methods.

### Method 1: Check and Change Restricted Security Settings

Excel provides security settings to protect your workbook. Sometimes, these settings can prevent inserting objects in the file. You can change the security settings to allow Excel to insert objects. To do so, follow these steps:

- Open your Excel application.
- Locate the **File** and then click **Options**.
- In **Excel Options**, click **Trust Center**.

![Trust Center In Excel Options](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/07/trust-center-in-excel-options.jpg)

- Click **Trust Center Settings**.
- In the **Trust Center Settings** window, select **Protected View** from the left pane.

![Click Protected View In Trust Center](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/07/click-protected-view-in-trust-center.jpg)

- Under **Protected View**, unselect the below three options:
- Enable Protected View for files originating from the internet.
- Enable Protected View for files located in potentially unsafe locations.
- Enable Protected View for Outlook attachments.

![Select All Options Under Protected View](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/07/select-all-options-under-protected-view.jpg)

- Click **OK**.
- Once you’re done with this, click on **Macro Settings** in the **Trust Center** window.
- Under **Macro Settings**, make sure **“Disable all macros without notification”** is not selected. If it is selected, then unselect it. After that, click **OK**.

![Click Macro Settings And Disable Macros Without Notifications](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/07/click-macro-settings-and-disable-macros-without-notifications.jpg)

- Restart Excel to apply the changes.

### **Method 2: Uninstall Microsoft Office Updates**

You can also encounter the “Cannot insert object” error in Excel after installing MS Office updates. It might be due to the issues with the installed updates. To fix this, you can uninstall the recently installed Office updates. To uninstall the Office updates, follow these steps:

- Go to the system’s Control Panel.
- Click **Programs** and then click **Program and Features.**
- Search for “**View Installed Updates**” and click on the desired Office updates.
- Right-click on it and then click **Uninstall**.
- Follow the uninstallation steps on the screen.
- Once the process is complete, restart the system.

### **Method 3: Check Memory Usage**

The “Cannot insert object” issue can also occur if your system is low on memory. You can check and close unnecessary processes and applications running in the background to free up memory. To do so, follow these steps:

- Press **CTRL + ALT + DEL** on the keyboard and click **Task Manager**.
- Click on the **Processes** tab and search for any unnecessary processes.
- Right-click on the process and then select **End Task**.
- Restart Excel to see if the issue is fixed.

### **Method 4: Check Excel File Size**

If your Excel file size exceeds the prescribed limit, it can also lead to the “Cannot insert Excel object” error. So, check the Excel file size. You can reduce the file size by removing unnecessary objects, such as formulas or images.

### **Method 5: Check and Change Excel ActiveX Settings**

You can get the “Excel cannot insert object” error if your Excel file contains macros, controls, and other interactive buttons. It usually occurs if the ActiveX Controls option is disabled. You can check and change the ActiveX Settings to fix the issue. Here are the steps:

- Open your Excel application.
- Navigate to **File** and then click **Options**.
- In **Excel Options**, click the **Trust Center** tab.
- In the **Trust Center Settings**, click **ActiveX Settings**.
- Under ActiveX Settings, make sure the “Enable all controls without restrictions and without prompting” option is selected.

![select enable all controls without restrictions under activexsettings](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/11/select-enable-all-controls-without-restrictions-under-activexsettings.png)

- If the option is not selected, then select it and click **OK**.
- Restart the Excel and check if the error is fixed or not.

### **Method 6: Repair the Excel Workbook**

The “Cannot insert object” error can occur if the object you are trying to insert is corrupted or the file in which you are inserting the object is damaged. If the issue has occurred due to a corrupted Excel file, then you can repair the file using the Open and Repair utility in MS Excel. To use this Microsoft-inbuilt utility, follow these steps:

- In the Excel application, go to the **File** tab and then click **Open**.
- Click **Browse** to choose the affected file.
- The **Open** dialog box is displayed. Click on the corrupted file.
- Click on the arrow next to the **Open** button and then click **Open and Repair**.
- Click on **Repair.**

![Click On Repair Option](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/07/click-on-repair-option.jpg)

- After repair, a message will appear (as shown in the below figure).

![Click Close Option In Repair Message](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/07/click-close-option-in-repair-message.jpg)

- Click **Close**.

If the [Open and Repair utility fails](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) to fix the issue, then try a professional Excel Repair software, like Stellar Repair for Excel. It is designed to repair severely corrupted Excel files. It can restore all the Excel file objects, such as tables, charts, formulas, etc. It helps fix all types of corruption related errors. The software is compatible with all versions of Excel.

## **Conclusion**

You might encounter the “Cannot insert object” error when embedding or inserting objects in Excel. In this post, we have discussed the possible solutions to fix this error. We have also mentioned an [Excel repair software](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) that can help to easily repair the corrupted Excel file and recover all the data. You can download the Stellar Repair for Excel’s free demo version to preview the recoverable objects of the corrupted Excel file.





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
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-music-from-xiaomi-redmi-note-12t-pro-by-fonelab-android-recover-music/"><u>How to retrieve erased music from Xiaomi Redmi Note 12T Pro</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-xs-without-losing-any-data-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone XS without Losing Any Data? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-mdm-from-iphone-15-pro-max-without-apple-id-by-drfone-ios-unlock-ios-unlock/"><u>How to Remove MDM from iPhone 15 Pro Max without Apple ID?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-remove-google-frp-lock-on-vivo-s17e-by-drfone-android-unlock-remove-google-frp/"><u>How to remove Google FRP Lock on Vivo S17e</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-11-without-itunes-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Reset iPhone 11 Without iTunes? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-retrieve-erased-call-logs-from-realme-10t-5g-by-fonelab-android-recover-call-logs/"><u>How to retrieve erased call logs from Realme 10T 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-call-logs-from-nokia-by-fonelab-android-recover-call-logs/"><u>How to rescue lost call logs from Nokia</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-electronically-sign-a-ext-using-digisigner-by-ldigisigner-sign-a-excel-sign-a-excel/"><u>How to Electronically Sign a {{ext}} Using DigiSigner</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-infinix-gt-10-pro-by-fonelab-android-recover-photos/"><u>How to recover deleted photos from Infinix GT 10 Pro.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-from-android-gallery-after-format-on-vivo-y100-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos from Android Gallery after format on Vivo Y100</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-downgrade-iphone-6s-to-the-previous-ios-version-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How to Downgrade iPhone 6s to the Previous iOS Version? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-malfunctioning-hardware-drivers-with-windows-device-manager-in-windows-11-and-10-and-7-by-drivereasy-guide/"><u>How to identify malfunctioning hardware drivers with Windows Device Manager in Windows 11 & 10 & 7</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-smart-8-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Smart 8</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-iphone-14-pro-camera-roll-photos-and-photo-stream-pictures-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted iPhone 14 Pro Camera Roll Photos and Photo Stream Pictures? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-data-from-iphone-14-pro-drfone-by-drfone-ios-data-recovery-ios-data-recovery/"><u>How To Recover Data from iPhone 14 Pro? | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-photos-on-tecno-phantom-v-flip-by-stellar-photo-recovery-android-mobile-photo-recover/"><u>How to recover deleted photos on Tecno Phantom V Flip</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-photos-files-from-vivo-v29-pro-by-fonelab-android-recover-photos/"><u>How To  Restore Missing Photos Files from Vivo V29 Pro.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-mov-files-of-galaxy-s24plus-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and MOV files of Galaxy S24+?</u></a></li>
<li><a href="https://blog-min.techidaily.com/2-ways-to-transfer-text-messages-from-honor-magic-6-to-iphone-1514131211x8-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>2 Ways to Transfer Text Messages from Honor Magic 6 to iPhone 15/14/13/12/11/X/8/ | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-reset-iphone-15-without-losing-data-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Reset iPhone 15 without Losing Data? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/4-ways-to-transfer-music-from-itel-a60s-to-iphone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>4 Ways to Transfer Music from Itel A60s to iPhone | Dr.fone</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-deleted-vivo-y02t-pictures-an-easy-method-explained-by-fonelab-android-recover-pictures/"><u>How to Restore Deleted Vivo Y02T Pictures  An Easy Method Explained.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-easily-copy-and-paste-your-forex-and-gold-trades-from-mt5-to-mt4-by-mt4copier-guide/"><u>How To Easily Copy & Paste Your Forex and Gold Trades From MT5 to MT4</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-photos-from-asus-rog-phone-8-pro-by-fonelab-android-recover-photos/"><u>How to Rescue Lost Photos from Asus ROG Phone 8 Pro?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-data-from-iphone-6-plus-using-stellar-data-recovery-for-iphone-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Data from iPhone 6 Plus using Stellar Data Recovery for iPhone? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-photos-files-from-honor-100-pro-by-fonelab-android-recover-photos/"><u>How To  Restore Missing Photos Files from Honor 100 Pro.</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-some-outdated-drivers-with-windows-device-manager-in-windows-1110-by-drivereasy-guide/"><u>How to identify some outdated drivers with Windows Device Manager in Windows 11,10</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-rescue-lost-contacts-from-p55-5g-by-fonelab-android-recover-contacts/"><u>How to Rescue Lost Contacts from P55 5G?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-old-music-from-your-vivo-y28-5g-by-fonelab-android-recover-music/"><u>How to recover old music from your Vivo Y28 5G</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-notes-from-iphone-12-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Notes from iPhone 12? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-identify-malfunctioning-your-drivers-with-windows-device-manager-in-windows-10-by-drivereasy-guide/"><u>How to identify malfunctioning your drivers with Windows Device Manager in Windows 10</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-recover-deleted-data-from-iphone-15-pro-max-using-stellar-data-recovery-for-iphone-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Recover Deleted Data from iPhone 15 Pro Max using Stellar Data Recovery for iPhone? | Stellar</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-repair-corrupt-mp4-and-avi-files-of-galaxy-m14-5g-by-stellar-video-repair-mobile-video-repair/"><u>How to Repair corrupt MP4 and AVI files of Galaxy M14 5G?</u></a></li>
<li><a href="https://ai-voice-clone.techidaily.com/new-in-2024-descript-overdub-controlling-the-audio-in-video/"><u>New In 2024, Descript Overdub Controlling the Audio in Video</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/easy-steps-to-recover-deleted-photos-from-honor-magic-6-pro-by-fonelab-android-recover-photos/"><u>Easy steps to recover deleted photos from Honor Magic 6 Pro.</u></a></li>
<li><a href="https://techidaily.com/the-way-to-get-back-lost-call-history-from-honor-x7b-by-fonelab-android-recover-call-logs/"><u>The way to get back lost call history from Honor X7b</u></a></li>
<li><a href="https://techidaily.com/hard-reset-nubia-red-magic-8s-proplus-in-3-efficient-ways-drfone-by-drfone-reset-android-reset-android/"><u>Hard Reset Nubia Red Magic 8S Pro+ in 3 Efficient Ways | Dr.fone</u></a></li>
<li><a href="https://techidaily.com/how-to-erase-apple-iphone-se-2020-data-permanently-drfone-by-drfone-ios-full-data-eraser-ios-full-data-eraser/"><u>How To Erase Apple iPhone SE (2020) Data Permanently | Dr.fone</u></a></li>
<li><a href="https://change-location.techidaily.com/how-to-deal-with-the-realme-narzo-60-5g-screen-black-but-still-works-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>How To Deal With the Realme Narzo 60 5G Screen Black But Still Works? | Dr.fone</u></a></li>
<li><a href="https://howto.techidaily.com/what-to-do-when-itel-p40plus-has-black-screen-of-death-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>What To Do When Itel P40+ Has Black Screen of Death? | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/easy-steps-to-recover-deleted-videos-from-honor-70-lite-5g-by-fonelab-android-recover-video/"><u>Easy steps to recover deleted videos from Honor 70 Lite 5G</u></a></li>
<li><a href="https://location-social.techidaily.com/in-2024-why-your-whatsapp-live-location-is-not-updating-and-how-to-fix-on-your-samsung-galaxy-a24-drfone-by-drfone-virtual-android/"><u>In 2024, Why Your WhatsApp Live Location is Not Updating and How to Fix on your Samsung Galaxy A24 | Dr.fone</u></a></li>
<li><a href="https://activate-lock.techidaily.com/in-2024-how-to-jailbreak-icloud-locked-iphone-14-pro-max-by-drfone-ios/"><u>In 2024, How to jailbreak iCloud locked iPhone 14 Pro Max</u></a></li>
<li><a href="https://ai-editing-video.techidaily.com/new-perfecting-the-art-of-slowing-down-video-on-instagram/"><u>New Perfecting the Art of Slowing Down Video on Instagram</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/here-are-some-pro-tips-for-pokemon-go-pvp-battles-on-oppo-a56s-5g-drfone-by-drfone-virtual-android/"><u>Here are Some Pro Tips for Pokemon Go PvP Battles On Oppo A56s 5G | Dr.fone</u></a></li>
<li><a href="https://ios-pokemon-go.techidaily.com/in-2024-how-to-enter-the-ispoofer-discord-server-on-apple-iphone-8-plus-drfone-by-drfone-virtual-ios/"><u>In 2024, How to enter the iSpoofer discord server On Apple iPhone 8 Plus | Dr.fone</u></a></li>
<li><a href="https://techidaily.com/best-fixes-for-realme-gt-5-240w-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>Best Fixes For Realme GT 5 (240W) Hard Reset | Dr.fone</u></a></li>
<li><a href="https://android-transfer.techidaily.com/in-2024-easiest-guide-how-to-clone-itel-s23plus-phone-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>In 2024, Easiest Guide How to Clone Itel S23+ Phone? | Dr.fone</u></a></li>
<li><a href="https://bypass-frp.techidaily.com/in-2024-hassle-free-ways-to-remove-frp-lock-from-xiaomi-civi-3-disney-100th-anniversary-edition-phones-withwithout-a-pc-by-drfone-android/"><u>In 2024, Hassle-Free Ways to Remove FRP Lock from Xiaomi Civi 3 Disney 100th Anniversary Edition Phones with/without a PC</u></a></li>
<li><a href="https://ai-video-apps.techidaily.com/2024-approved-transform-your-footage-a-step-by-step-guide-to-ken-burns-effect-in-final-cut-pro/"><u>2024 Approved Transform Your Footage A Step-by-Step Guide to Ken Burns Effect in Final Cut Pro</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/how-to-use-pokemon-go-joystick-on-nokia-105-classic-drfone-by-drfone-virtual-android/"><u>How to use Pokemon Go Joystick on Nokia 105 Classic? | Dr.fone</u></a></li>
<li><a href="https://techidaily.com/solutions-to-repair-corrupt-pdf-v13-file-stellar-by-stellar-guide/"><u>Solutions to Repair Corrupt PDF v1.3 File | Stellar</u></a></li>
<li><a href="https://ai-video-translation.techidaily.com/new-2024-approved-top-8-multi-subtitles-translators-you-can-use/"><u>New 2024 Approved Top 8 Multi-Subtitles Translators You Can Use</u></a></li>
<li><a href="https://ios-pokemon-go.techidaily.com/in-2024-detailed-guide-of-ispoofer-for-pogo-installation-on-apple-iphone-11-pro-drfone-by-drfone-virtual-ios/"><u>In 2024, Detailed guide of ispoofer for pogo installation On Apple iPhone 11 Pro | Dr.fone</u></a></li>
<li><a href="https://iphone-unlock.techidaily.com/in-2024-how-to-unlock-disabled-apple-iphone-8ipad-without-computer-drfone-by-drfone-ios/"><u>In 2024, How to Unlock Disabled Apple iPhone 8/iPad Without Computer | Dr.fone</u></a></li>
<li><a href="https://fix-guide.techidaily.com/quick-fixes-for-why-is-my-realme-11-pro-black-and-white-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>Quick Fixes for Why Is My Realme 11 Pro Black and White | Dr.fone</u></a></li>
<li><a href="https://ai-voice-clone.techidaily.com/updated-the-most-efficient-tools-to-convert-text-to-mp3-with-the-best-natural-voices/"><u>Updated The Most Efficient Tools to Convert Text to MP3 With the Best Natural Voices</u></a></li>
<li><a href="https://location-social.techidaily.com/in-2024-why-your-whatsapp-live-location-is-not-updating-and-how-to-fix-on-your-motorola-razr-40-ultra-drfone-by-drfone-virtual-android/"><u>In 2024, Why Your WhatsApp Live Location is Not Updating and How to Fix on your Motorola Razr 40 Ultra | Dr.fone</u></a></li>
<li><a href="https://howto.techidaily.com/what-to-do-if-google-play-services-keeps-stopping-on-xiaomi-redmi-note-12-proplus-5g-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>What to Do if Google Play Services Keeps Stopping on Xiaomi Redmi Note 12 Pro+ 5G | Dr.fone</u></a></li>
</ul></div>



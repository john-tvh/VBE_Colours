# What is VBE_Colours?

VBE_Colours is a small application that adjusts the colours shown in the code window and the immediate window of the Visual Basic Editor (the VBE). It does this by determining the location of the various bytes that define the colours in your VBE7.DLL file, reading those bytes and then, when you apply new colours, writing new byte values to patch the file.

## The VBE styled with a dark theme
![dark](https://user-images.githubusercontent.com/97969304/196000931-ae3cda8d-b540-4ca4-8fb9-45f60f42f553.png)

## The VBE styled with a blue theme
![blue](https://user-images.githubusercontent.com/97969304/196001145-5c8551d3-f63b-4474-9848-c3f2b6988bd5.png)

## VBE_Colours in action
![vbe_colours_1](https://user-images.githubusercontent.com/97969304/196001147-8a4f1283-e576-43f2-afaa-fc2a635a62f4.png)

## VBE_Colours in action
![vbe_colours_2](https://user-images.githubusercontent.com/97969304/196001152-8d9c70ae-6454-46e0-92e7-eb2429e4415f.png)

# To download and install VBE_Colours

The latest version of VBE_Colours is 1.0.0.1 - click to [download](https://github.com/john-tvh/VBE_Colours/wiki) it (from the Wiki page).

VBE_Colours is digitally signed. To ensure you have downloaded a valid and safe file, when installing please ensure the publisher is 'John Mallinson' (ie me ... the owner of this site!). Once downloaded to your device (the download is a .zip file), extract and then double-click the VBE_Colours_Setup.exe file. Depending on your device settings and the number of prior downloads of this specific version of VBE_Colours, your device may display one or more warnings that you will have to accept before you see the VBE_Colours installation wizard which will guide you through the rest of the installation process.

Once installed, VBE_Colours will offer to run immediately, or you can run it from your Windows 'Start' menu.

# How to use VBE_Colours

## Selecting colours

When VBE_Colours first loads, it will read the colours from your VBE7.DLL file (there are always 16 colours). These colours will be shown in the larger 'custom' colour boxes. The smaller 'default' colour boxes are the default VBE colours.
-	Click on any of the larger 'custom' colour boxes to select a different colour using the standard Windows colour-selection dialog
-	Right-click on any of the larger 'custom' colour boxes for more options
-	Click on 'File', 'Manage' or 'More' for more options including saving and loading colour schemes and setting the VBE Registry values

There are two ways to set colours (you can switch between these two methods whenever you want, initially I suggest using a **custom colour scheme**):

-	Create a **custom colour scheme** ??? suggest starting with one of the built-in custom colour schemes ??? 'Blue', 'Dark' or 'Light'  (select 'File' then 'Load colour scheme' ??? ensure to allow VBE_Colours to set the appropriate Registry values otherwise the selected colours will not correctly link to the right 'type of text')
    -	Optionally, to identify which of the larger 'custom' colour boxes relates to what foreground / background / indicator colour, select 'Show sample text and custom colour tags' (via the 'Manage' button).
-	Create a **default colour scheme** by tweaking the defaults ??? either start with the colours that VBE_Colours loaded from your VBE7.DLL file or use one of the 'VBE' or 'Alt' built-in colour schemes (select 'File' then 'Load colour scheme'). Change the colours to match your requirements
    -	When creating a default colour scheme, do not select 'Show sample text and custom colour tags' as these only relate to custom colour schemes (there is not a one-to-one relationship between each of the larger 'custom' colour boxes and each foreground / background / indicator colours when creating a default colour scheme).

## Applying colours

*It is very strongly recommended that you make a backup of your VBE7.DLL file before applying colours.*

You will have to close all VBA-enabled applications before you can apply colours. Press 'Apply'. Open the VBE in any of your VBA-enabled applications and you will automatically see the colours you have applied.

If you don't get the results you expected (or do not see any changes at all), use VBE_Colours to set the Registry values to 'default' if you were creating a **default colour scheme** or to 'custom' if you were creating a **custom colour scheme** (do this via the 'Manage' button) ??? the reason you may not see the colour changes you expected is that, by default, almost all of the colours used by the VBE are 'Windows standard colours' (this is the same as selecting 'Auto' in the colour drop-downs in the Editor Format tab of the Options dialog in the VBE) rather than using explicit colours ??? only by changing the registry values to specify explicit colours do you then see the colours you have selected using VBE_Colours.

When you apply colours, they will be applied to all host applications that use the same VBE7.DLL file.

# When VBE7.DLL is updated by Microsoft

The VBE7.DLL file is updated occasionally as part of updates to the wider Office environment. When it is updated, your carefully selected colour scheme will be lost as Microsoft knows nothing for the colour scheme you selected (though the Registry values will persist).

For this reason it is important that you save your colour scheme so that when the day comes that your VBE7.DLL file is updated and your colour scheme is overwritten (and you open the VBE only to see "colour chaos" because the Registry values are then applying the default VBE colours but not in a way you would expect), then you can come back to VBE_Colours, load your colour scheme (or the default if you want to go back to the default VBE colours) and 'Apply' it.

# Deleting all colour modifications

If you want to delete all colour modifications from both your VBE7.DLL file and from the Registry:

- Using VBE_Colours: select the 'Reset colours' option (via the 'Manage' button ??? including setting the Registry values to 'default', if asked) then 'Apply'
- Without using VBE_Colours: replace the modified version of VBE7.DLL with the backup version you made. And either delete the three Registry values (see [Link between Registry values and the colours defined in VBE7.DLL](https://github.com/john-tvh/VBE_Colours/blob/main/README.md#link-between-registry-values-and-the-colours-defined-in-vbe7dll)) or use the colour drop-downs in the Editor Format tab of the Options dialog in the VBE to select the VBE default colours.

# Donate

It takes time and effort to develop apps such as VBE_Colours. If it is useful for you, I would be grateful if you would donate which allows me to continue to maintain VBE_Colours and also to develop other apps such as [VBE_Extras](https://github.com/john-tvh/VBE_Extras).

[![paypal](https://www.paypalobjects.com/en_US/i/btn/btn_donateCC_LG.gif)](https://www.paypal.com/paypalme/thevbahelp)

# Technical info

## Link between Registry values and the colours defined in VBE7.DLL

The colours shown in the code and immediate windows of the VBE are controlled by a combination of:

-	The 16 colours defined in the VBE7.DLL file
-	3 Registry values (1 each for the foreground colours, background colours and indicator colours) that specify which of the colours defined in the VBE7.DLL file are shown for which 'type of text' (normal, selection, keyword etc)

When you use the colour drop-downs in the Editor Format tab of the Options dialog in the VBE, you are updating values in the Registry. These values tell the VBE which (of the 16 colours in VBE7.DLL) should be shown for which 'type of text' (there are 10 'types of text': normal, selection, syntax error, execution point, breakpoint, comment, keyword, identifier, bookmark and call return). There are 3 Registry values, 1 each for foreground colours, background colours and indicator colours:

-	HKEY_CURRENT_USER\SOFTWARE\Microsoft\VBA\7.1\Common\CodeForeColors
-	HKEY_CURRENT_USER\SOFTWARE\Microsoft\VBA\7.1\Common\CodeBackColors
-	HKEY_CURRENT_USER\SOFTWARE\Microsoft\VBA\7.1\Common\IndicatorColors

When using VBE_Colours to set a colour scheme, in order to form a correct association between the colours defined in VBE7.DLL and those shown in the VBE for a specific 'type of text', specific Registry values must be used. These Registry values are set when you select 'Set Registry values to 'default'' or 'Set Registry values to 'custom''.

-	When creating a **custom colour scheme**, it is recommended that you let VBE_Colours set the appropriate Registry values and that you then leave them (ie do not change the colours using the Options dialog ??? though nothing is stopping you from doing this, it is unlikely you will get the results that you want). If you want to change your custom colour scheme, it is better to come back to VBE_Colours to change your colours and then 'Apply' them again.
-	When creating a **default colour scheme**, it is recommended that you let VBE_Colours set the appropriate Registry values initially. Once you have selected the colour scheme you want and clicked 'Apply' then you can change the Registry keys (using the Options dialog ??? or, if you have [VBE_Extras](https://github.com/john-tvh/VBE_Extras) installed, by setting a Theme) to get the right colours to be applied to the type of text that you want.

Note that you can change between using a default colour scheme and a custom colour scheme quite simply using VBE_Colours, and you can save a colour scheme as either 'default' or 'custom' and you will be prompted to apply the appropriate Registry values if you load that scheme in the future.

## Indicator colours

There is a bug in VBE 7.1 (and perhaps in older versions) in that it will load previously-saved indicator colours from the IndicatorColors Registry value, but it does not save them to the Registry value.

To demonstrate this, in any VBA-enabled application, start the VBE with the default colours; then using the Options dialog, change the breakpoint background colour and the breakpoint indicator colour to, say, green and then OK the dialog; add a breakpoint and you will see the colours you chose (green if that is what you selected); now close and re-open your application, go into the VBE and add a breakpoint and you will see the background colour has been set to the colour you chose (eg green) but the indicator colour has reverted to the default maroon colour.

VBE_Colours will resolve this for you by setting the IndicatorColors Registry value and so ensuring that the indicator colours you choose with VBE_Colours are correctly shown in the VBE.

## Limitations

VBE_Colours works only with VBA 7.1 (either 32 or 64 bit) ??? if you are using the VBE hosted by an Office application, VBA 7.1 is used in Office 2013 or newer. VBE_Colours will not work with VBA 7.0 or older.

VBE_Colours SHOULD work no matter which host app is being used. I have tested it with Excel, Word, Outlook and PowerPoint as they are the VBA-hosting apps that I use frequently. VBE_Colours SHOULD work with other Office apps that host VBA (eg Access, Visio etc) and also with non-Office apps that host VBA (eg AutoCAD, CorelDraw etc) ??? but I cannot guarantee this as I do not own a copy of those to test with (I'd be interested to learn from any users of those apps if VBE_Colours does indeed work with those apps).

In order to patch (ie write bytes to) the VBE7.DLL file, VBE_Colours must run (and install) with admin privileges.

As Excel, Word, Outlook and PowerPoint are my area of expertise then these notes relate particularly to the VBE when used within those applications.

# No, or multiple, VBE7.DLL files

In the unlikely event that VBE_Colours cannot automatically find your VBE7.DLL file then you will be shown a dialog for you to locate it. For MS Office hosts, typically, VBE7.DLL will be in: "C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\VBA\VBA7.1" (depending on the 'bitness' of Windows and Office, possibly replace "Program Files" with "Program Files (x86)" and "ProgramFilesCommonX64" with "ProgramFilesCommonX86"). Where it will NOT be is in: "C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1" which is where the References dialog in the VBE will tell you that it is.

And in the also unlikely event that you have multiple VBE7.DLL files, you will be shown a dialog to select which one you want to work with. If you do have multiple VBE7.DLL files, you can define a colour scheme with VBE_Colours and then 'Apply' it to one VBE7.DLL file, save the colour scheme and then apply it to the other VBE7.DLL files one at a time (by closing the re-opening VBE_Colours and selecting a different VBE7.DLL file).

If you want to force VBE_Colours to let you manually choose the VBE7.DLL file, you can run it with the "m" or "manual" command-line argument.

# I found a bug!

If you have a problem with VBE_Colours, please first ensure you have read all of the above. If you still believe you have found a bug, please report it to me as an [issue](https://github.com/john-tvh/VBE_Colours/issues).

---
layout: single
author_profile: true
breadcrumbs: true
categories:  Technology
tags : Technology 
comments: false
title:  "How to make Outlook Calendar reminders stay on top in Windows 7"
description:  "Step by step instructions on how to have outlook reminder pop up on the screen"
date:   10/4/2017 4:27:29 PM 
---

## Problem ##
Its quite annoying when outlook meetings reminders don't pop up. Usually these reminders are hiding behind some other screen. 

This post will show you a trick to always have those reminders pop right up in front.

<script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
<script>
  (adsbygoogle = window.adsbygoogle || []).push({
    google_ad_client: "ca-pub-8249317511242119",
    enable_page_level_ads: true
  });
</script>

## Solution  ##

Follow the steps and screen shots below and in less then a few minutes you will be all set.

### Step 1 ###
Open Windows Explorer
### Step 2 ###

If you have  - 

**Office 2010** then go to the following folder 

C:\Program Files\Microsoft Office\Office15

For **office 2013** go to

C:\Program Files\Microsoft Office\Office16

If for some reason you do not have any of these folders then search windows for SELFCERT.EXE. Search on windows is usually slow so be patient. If the search doesnt bring anything back, then its time to install a better windows searching tool. Stay tuned, I will be writing a post on the search tools I use.

**find and click SELFCERT.EXE** 

![selfcert_screenshot.PNG](/assets/images/OutlookReminders/selfcert_screenshot.PNG)

### Step 3 ###
Click on the Selfcert.ext file a pop will come up. 

Give the certificate a name. For this post I gave it a name **"text_outlook_certificate"**

![screenshot_certificateName.PNG](/assets/images/OutlookReminders/screenshot_certificateName.PNG)

Click ok.

### Step 4 ###
Open outlook and press Alt+F11 key together.

It will open up a VBA code editor. Something similar to the screen shot below.

![screenshot_vbaEditor.PNG](/assets/images/OutlookReminders/screenshot_vbaEditor.PNG)

Copy and paste the following code into the right side of the editor.
    
    Private Declare PtrSafe Function FindWindowA Lib "user32" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
    Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
    
    Private Const SWP_NOSIZE = &H1
    Private Const SWP_NOMOVE = &H2
    Private Const FLAGS As Long = SWP_NOMOVE Or SWP_NOSIZE
    Private Const HWND_TOPMOST = -1
    
    
    Private Sub Application_Reminder(ByVal Item As Object)
    Dim ReminderWindowHWnd As Variant
    If ( _
    (TypeOf Item Is AppointmentItem) _
    Or (TypeOf Item Is MailItem) _
    Or (TypeOf Item Is TaskItem) _
    ) Then
    
    Dim MsgBoxResult As Integer
    
    Do
    
    MsgBoxResult = _
    ( _
    MsgBox _
    ( _
    "An Outlook Reminder is awaiting! Select OK to continue.", _
    vbSystemModal + vbOKCancel + vbDefaultButton2, _
    "Attention!" _
    ) _
    ) ' -- MsgBoxResult =
    
    Loop Until (MsgBoxResult = vbOK)
    
    End If
    
    On Error Resume Next
    ReminderWindowHWnd = FindWindowA(vbNullString, "1 Reminder")
    SetWindowPos ReminderWindowHWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
    
    End Sub
    

### Step 5 ###

While in the VBA Editor press the following keys  -

Press Alt then T and then D  on your keyboard. 

* You can also use the mouse to click the **Tools** menu at the top and then click **Digital Signature** on the menu. 

It will open a pop up.

![screenshot_setcertificate.PNG](/assets/images/OutlookReminders/screenshot_setcertificate.PNG)

Click the choose button and the in the next pop up it will show all the certificates that you have created. 

Select the one we just created -**text_outlook_certificate** 

![select_certificate.PNG](/assets/images/OutlookReminders/select_certificate.PNG)

 
### Final Step ###

Save everything, close the VBA editor and close and restart outlook.

*That's all there is to it.*

Going forward whenever you have a meeting, a pop up will come up on top of your windows screen to remind you. Just click the OK button on the pop up and it will bring up the outlook reminder screen. 

![messageBox.PNG](/assets/images/OutlookReminders/messageBox.PNG)

Thanks for reading. For questions email me. I will be happy to help.

<script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
<script>
  (adsbygoogle = window.adsbygoogle || []).push({
    google_ad_client: "ca-pub-8249317511242119",
    enable_page_level_ads: true
  });
</script>
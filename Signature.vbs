On Error Resume Next

''____________________________________________________________________
''
''	signature.vbs
''	
''	This VBScript will prompt the user for the details needed to create a new signature for the General Chiropractic Council.
''
''	If the user presses "cancel" for any of the prompts, the whole script will quit.
''
''	The code is based on https://www.codetwo.com/admins-blog/vbscript-create-an-html-outlook-email-signature-for-the-whole-company/, 
''	but is adapted for a simpler html signature by Andrew Fielding. 
''
''	Once the script has run, it will add a new HTML signature to this folder: C:\Users\[username]\AppData\Roaming\Microsoft\Signatures named "Corporate Signature"
''	
''	
''____________________________________________________________________

'Setting up the script to work with the file system.
Set WshShell = WScript.CreateObject("WScript.Shell")
Set FileSysObj = CreateObject("Scripting.FileSystemObject")
Set objSysInfo = CreateObject("ADSystemInfo")
'' At work this works:
''Set UserObj = GetObject("LDAP://" & objSysInfo.UserName)
''strAppData = WshShell.ExpandEnvironmentStrings("%APPDATA%")
''SigFolder = StrAppData & "\Microsoft\Signatures\"
''SigFile = SigFolder & "Corporate Signature - (" & UserObj.sAMAccountName & "@wobable.com).htm" 'This may need to be edited based on your own set up

'' At home I have to use:
SigFolder = "C:\Users\wobab\AppData\Roaming\Microsoft\Signatures\"
SigFile = SigFolder & "Corporate Signature - (wobable.com).htm" 'This may need to be edited based on your own set up


MsgBox("Welcome to the Signature Generator.")

'____________________________
''Collect Name
strName = InputBox("Please enter your Name:", "Name", "John Jones")
If strName = "" Then
	Wscript.Quit
End If
'____________________________

'____________________________
''Collect Pronouns
strPronoun = InputBox("Please enter your pronouns:", "Pronouns", "he/him/his")
If strPronoun = "" Then
	Wscript.Quit
End If
'____________________________

'____________________________
''Collect Pronounciation
strPronounciation = InputBox("Please enter any pronounciation guidance, or delete for none:", "Pronounciation", "John, not Jonathon or Jonny" )
'____________________________

'____________________________
''Collect Job Title
strJobtitle = InputBox("Please enter your job title:", "Job title", "Chief Executive")
If strJobtitle = "" Then
	Wscript.Quit
End If
'____________________________

'____________________________
''Collect Phone Number
strPhone = InputBox("Please enter your phone number:", "Phone number", "01632 123456 Extn XXXX")
If strPhone = "" Then
	Wscript.Quit
End If
'____________________________

'____________________________
''Collect Core Working Hours
strHours = InputBox("Please enter your Core Working Hours:", "Core Hours", "Monday to Friday, 9-5pm")
If strHours = "" Then
	Wscript.Quit
End If
'____________________________

'____________________________
''Collect standard sign off
strValediction = InputBox("Please enter your standard sign off:", "Valediction", "Yours sincerely / Yours faithfully / Kind regards / Best wishes")
If strValediction = "" Then
	Wscript.Quit
End If
'____________________________

'Setting global placeholders for the signature. Those values will be identical for all users - make sure to replace them with the right values!
strBanner = "https://res.cloudinary.com/du4ivgfle/image/upload/v1691619582/Advert_bcf9ec.png" ''400*200 pixels
strBannerLink = "https://www.wobable.com"
strCompany = "Acme Products Corporation"
strCompanyAddress = "221B Baker Street, London SW1A 1AA"
strWebsite = "www.wobable.com"
strTwitter = "https://www.twitter.com/wobable"
strLinkedIn = "https://www.linkedin.com/company/wobable/"
strFlexworking = "I may choose to work flexibly and send emails outside normal office hours. There is no need to respond to my emails outside yours."
strWelsh = "Os byddwch yn dewis ysgrifennu atom yn Gymraeg, byddwn yn ymateb yn Gymraeg. Ni fydd gohebu yn Gymraeg yn arwain at oedi. Fodd bynnag, nodwch nad oes yr un o'n staff yn siarad Cymraeg ar hyn o bryd."
strLegal = "This email and any files transmitted with it are confidential. If you are not the intended recipient, any reading, printing, storage, disclosure, copying or any other action taken in respect of this email is prohibited and may be unlawful. If you are not the intended recipient, please notify the sender immediately by using the reply function and then permanently delete what you have received."

'Creating HTM signature file for the user's profile.
Set CreateSigFile = FileSysObj.CreateTextFile (SigFile, True, True)



''________________Signature’s HTML code________________
CreateSigFile.WriteLine "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>"
CreateSigFile.WriteLine "<HTML><HEAD><TITLE>GCC Standard Email Signature</TITLE>"
CreateSigFile.WriteLine "<META content='text/html; charset=utf-8' http-equiv='Content-Type'>"
CreateSigFile.WriteLine "</HEAD>"
CreateSigFile.WriteLine "<p style='font-family:Arial,sans-serif;color:#5A5A5A'>"

''________________Valediction________________
CreateSigFile.WriteLine replace(strValediction,",","") & ", </p>" 'REPLACE ensures there is a comma, but only one.

''________________empty lines________________
CreateSigFile.WriteLine "<br/><br/>"

''________________Name and Contact Details Table________________
CreateSigFile.WriteLine "<table align=left style='border-collapse:collapse;border:none;font-size:16px'>"
CreateSigFile.WriteLine "<tr>"
CreateSigFile.WriteLine "<td style='padding-right:20px;'>"
CreateSigFile.WriteLine "<p style='font-family:Arial,sans-serif;color:#003b5c'><b>"& strName &"</b> (" & strPronoun & ")"
''Do they want pronounciation?
If strPronounciation = "" Then
	CreateSigFile.WriteLine "<br/><i>Pronounciation: " & strPronounciation & ",<i>"   
End If
CreateSigFile.WriteLine "</p><p style='font-family:Arial,sans-serif;color:#003b5c'>" & strJobtitle
CreateSigFile.WriteLine "<br/>" & strCompany
CreateSigFile.WriteLine "<br/>" & strCompanyAddress
CreateSigFile.WriteLine "<br/>" & "Tel: "& strPhone
CreateSigFile.WriteLine "<br/>" & "Website:<u><a href='http://" & strWebsite & "'><span style='color:#003b5c'>" & strWebsite & "</a></u></span>"
CreateSigFile.WriteLine "<br/>" & "Social:<u><a href='"& strTwitter &"'><span style='color:#003b5c'>Twitter</span></a></u> | <u><a href='"& strLinkedIn &"'><span style='color:#003b5c'>LinkedIn</span></a></u></p>"
CreateSigFile.WriteLine "</td>"
CreateSigFile.WriteLine "</tr>"
CreateSigFile.WriteLine "</table>"

''________________Add floating linked Image File________________

CreateSigFile.WriteLine "<a href='"& strBannerLink &"'><img src='" & strBanner & "' alt='Email footer advert' style=''></a>"

''________________Add base table________________

CreateSigFile.WriteLine "<table align=left width=100% style='border-collapse:collapse;border:none;'>"
CreateSigFile.WriteLine "<tr>"
CreateSigFile.WriteLine "<td width=100%>"
CreateSigFile.WriteLine "<p style='font-family:Arial,sans-serif;color:#003b5c'><i>My core working hours are " & strHours & "</i><br/>"
CreateSigFile.WriteLine "<i>" & strFlexworking & "</i></p>"
CreateSigFile.WriteLine "<p style='font-family:Arial,sans-serif;color:#003b5c'><i>" & strWelsh & "</i></p>"
CreateSigFile.WriteLine "<p style='font-family:Arial,sans-serif;color:#003b5c'>" & strLegal & "</p>"
CreateSigFile.WriteLine "</td>"
CreateSigFile.WriteLine "</tr>"
CreateSigFile.WriteLine "</table>"
CreateSigFile.WriteLine "</body>"
CreateSigFile.WriteLine "</html>"
CreateSigFile.Close

Dim Sh
Set Sh = WScript.CreateObject("WScript.Shell")
Sh.Run Chr(34) & SigFile & Chr(34), 9

MsgBox("All done")

''I believe the below sets it as default in Outlook - I have commented out from the orginal code, but it may be of interest to you!
'Set objWord = CreateObject("Word.Application")
'Set objSignatureObjects = objWord.EmailOptions.EmailSignature
'objSignatureObjects.NewMessageSignature = strUserName & "2"
'objSignatureObjects.ReplyMessageSignature = strUserName & "2"
'objWord.Quit
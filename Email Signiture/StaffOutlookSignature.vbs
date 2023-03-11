'Staff Outlook Signiture September 2022'
'Author Brandon Roff'

On Error Resume Next

'Variable Set'
Set objSysInfo = CreateObject("ADSystemInfo")
Set WshShell = CreateObject("WScript.Shell")
strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)
strName = objUser.FullName
strTitle = objUser.Title
strCred = objUser.info
strTelephone = "01708 865 180"
strMobile = objUser.Mobile
strEmail = objUser.mail
strWebsite = "https://ormistonpark.org.uk/"
strLogo = "\\Padc.internal\netlogon\EmailSignature\Images\Footers\opa-footer.png"
strLogo1 = "\\Padc.internal\netlogon\EmailSignature\Images\Footers\opa-footer-1.png"
strLogo2 = "\\Padc.internal\netlogon\EmailSignature\Images\Footers\opa-footer-2.png"
strSeasonalLogo = "\\Padc.internal\netlogon\EmailSignature\Images\Footers\Openevening.png"
strDDI = objUser.homePhone
strWH = "8am to 4pm"


strFollowUsText = "\\Padc.internal\netlogon\EmailSignature\Images\SocialMedia\FollowUs.png"
strFacebookLogo = "\\Padc.internal\netlogon\EmailSignature\Images\SocialMedia\facebook.png"
strTwitterLogo = "\\Padc.internal\netlogon\EmailSignature\Images\SocialMedia\twitter.png"
strInstagramLogo = "\\Padc.internal\netlogon\EmailSignature\Images\SocialMedia\instagram.png"
strTiktokLogo = "\\Padc.internal\netlogon\EmailSignature\Images\SocialMedia\Tiktok.png"


'Full Signiture 1 Start'

Set objWord = CreateObject("Word.Application")

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature

Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

objSelection.Font.Name = "Calibri"
objSelection.Font.Size = 18
objSelection.Font.Color = RGB(242,125,0)
objSelection.Font.Bold = False
if (strCred) Then objSelection.TypeText strName & ", " & strCred Else objSelection.TypeText strName


objSelection.TypeText Chr(11)
objSelection.Font.Bold = True
objSelection.Font.Size = 14
objSelection.Font.Color = RGB(102,102,102)
if (strTitle) Then objSelection.TypeText strTitle


objSelection.Font.Size = 14
objSelection.Font.Bold = False
objSelection.Font.Size = 13
objSelection.TypeText Chr(11) 
objSelection.TypeText "Ormiston Park Academy"


objSelection.Font.Size = 10
objSelection.TypeText Chr(11)
objSelection.TypeText "Belhus Park Lane, Aveley, Essex, RM15 4RU"


objSelection.TypeText Chr(11)
objSelection.Font.Bold = True
objSelection.TypeText "Phone: "
objSelection.Font.Bold = False
objDoc.Hyperlinks.Add objSelection.Range, "tel:" & strTelephone,,,strTelephone
objSelection.TypeText " "
objSelection.Font.Bold = True
if (strDDI) Then  objSelection.TypeText "| DDI: "
objSelection.Font.Bold = False
if (strDDI) Then objSelection.TypeText strDDI
if (strDDI) Then objSelection.TypeText " "
objSelection.Font.Bold = True
'if (strPhone) Then objSelection.TypeText "| Ext: " '
objSelection.Font.Bold = False
if (strPhone) Then objSelection.TypeText strPhone

if (strShowMobile) Then objSelection.TypeText Chr(11)
if (strShowMobile) Then objSelection.Font.Bold = True
if (strShowMobile) Then objSelection.TypeText "Mobile: "
if (strShowMobile) Then objSelection.Font.Bold = False
if (strShowMobile) Then objSelection.TypeText strMobile

objSelection.TypeText Chr(11)
objSelection.Font.Bold = True
if (strEmail) Then objSelection.TypeText "Email: "
objSelection.Font.Bold = False
objDoc.Hyperlinks.Add objSelection.Range, "mailto:" & strEmail,,,strEmail & Chr(11)


objSelection.Font.Bold = True
objSelection.TypeText "Website: "
objSelection.Font.Bold = False
objDoc.Hyperlinks.Add objSelection.Range, strWebsite,,,"www.ormistonpark.org.uk"

if (strWH) Then  objSelection.TypeText Chr(11)
if (strWH) Then  objSelection.Font.Bold = True
if (strWH) Then  objSelection.TypeText "Opening Hours: "
if (strWH) Then  objSelection.Font.Bold = False
if (strWH) Then  objSelection.TypeText strWH

objSelection.TypeText Chr(11)
objSelection.TypeText Chr(11)
objSelection.InlineShapes.AddPicture(strLogo)
objSelection.TypeText Chr(11)
objSelection.TypeText Chr(11)
'Added 07/09/2022'
'Set pic = '
objSelection.InlineShapes.AddPicture(strSeasonalLogo)
'pic.Height = 175'
'pic.Width  = 500'
objSelection.TypeText Chr(11)
objSelection.TypeText Chr(11)
'End of Addition for 07/09/2022'
objSelection.InlineShapes.AddPicture(strFollowUsText)
objSelection.TypeText Chr(11)
objSelection.TypeText Chr(11)
objSelection.InlineShapes.AddPicture(strFacebookLogo)
objSelection.InlineShapes.AddPicture(strTwitterLogo)
objSelection.InlineShapes.AddPicture(strInstagramLogo)
objSelection.InlineShapes.AddPicture(strTiktokLogo)


objDoc.InlineShapes.Item(7).ScaleHeight = 12
objDoc.InlineShapes.Item(7).ScaleWidth = 12

objDoc.Hyperlinks.Add objDoc.InlineShapes,Item(3), "https://ormistonpark.org.uk/open-house" 
objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(4), "https://www.facebook.com/ormistonpark" & Chr(11)
objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(5), "https://www.twitter.com/ormistonpark" & Chr(11)
objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(6), "https://www.instagram.com/ormistonpark/" & Chr(11)
objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(7), "https://www.tiktok.com/@ormistonparkofficial"

Set objSelection = objDoc.Range()

objSignatureEntries.Add "Full Signature", objSelection
objSignatureObject.NewMessageSignature = "Full Signature"

objDoc.Saved = True
objWord.Quit
'Full Signiture  End'

'Full Signiture 1 Start'

Set objWord = CreateObject("Word.Application")

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature

Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

objSelection.Font.Name = "Calibri"
objSelection.Font.Size = 18
objSelection.Font.Color = RGB(242,125,0)
objSelection.Font.Bold = False
if (strCred) Then objSelection.TypeText strName & ", " & strCred Else objSelection.TypeText strName


objSelection.TypeText Chr(11)
objSelection.Font.Bold = True
objSelection.Font.Size = 14
objSelection.Font.Color = RGB(102,102,102)
if (strTitle) Then objSelection.TypeText strTitle


objSelection.Font.Size = 14
objSelection.Font.Bold = False
objSelection.Font.Size = 13
objSelection.TypeText Chr(11) 
objSelection.TypeText "Ormiston Park Academy"


objSelection.Font.Size = 10
objSelection.TypeText Chr(11)
objSelection.TypeText "Belhus Park Lane, Aveley, Essex, RM15 4RU"


objSelection.TypeText Chr(11)
objSelection.Font.Bold = True
objSelection.TypeText "Phone: "
objSelection.Font.Bold = False
objDoc.Hyperlinks.Add objSelection.Range, "tel:" & strTelephone,,,strTelephone
objSelection.TypeText " "
objSelection.Font.Bold = True
if (strDDI) Then  objSelection.TypeText "| DDI: "
objSelection.Font.Bold = False
if (strDDI) Then objSelection.TypeText strDDI
if (strDDI) Then objSelection.TypeText " "
objSelection.Font.Bold = True
'if (strPhone) Then objSelection.TypeText "| Ext: "'
objSelection.Font.Bold = False
if (strPhone) Then objSelection.TypeText strPhone

if (strShowMobile) Then objSelection.TypeText Chr(11)
if (strShowMobile) Then objSelection.Font.Bold = True
if (strShowMobile) Then objSelection.TypeText "Mobile: "
if (strShowMobile) Then objSelection.Font.Bold = False
if (strShowMobile) Then objSelection.TypeText strMobile

objSelection.TypeText Chr(11)
objSelection.Font.Bold = True
if (strEmail) Then objSelection.TypeText "Email: "
objSelection.Font.Bold = False
objDoc.Hyperlinks.Add objSelection.Range, "mailto:" & strEmail,,,strEmail & Chr(11)

objSelection.Font.Bold = True
objSelection.TypeText "Website: "
objSelection.Font.Bold = False
objDoc.Hyperlinks.Add objSelection.Range, strWebsite,,,"www.ormistonpark.org.uk"

if (strWH) Then  objSelection.TypeText Chr(11)
if (strWH) Then  objSelection.Font.Bold = True
if (strWH) Then  objSelection.TypeText "Opening Hours: "
if (strWH) Then  objSelection.Font.Bold = False
if (strWH) Then  objSelection.TypeText strWH
objSelection.TypeText Chr(11)
objSelection.TypeText Chr(11)
objSelection.InlineShapes.AddPicture(strLogo1)
objSelection.TypeText Chr(11)
objSelection.TypeText Chr(11)
'Added 07/09/2022'
'Set pic = '
objSelection.InlineShapes.AddPicture(strSeasonalLogo)
'pic.Height = 175'
'pic.Width  = 500'
objSelection.TypeText Chr(11)
objSelection.TypeText Chr(11)
'End of Addition for 07/09/2022'
objSelection.InlineShapes.AddPicture(strFollowUsText)
objSelection.TypeText Chr(11)
objSelection.TypeText Chr(11)
objSelection.InlineShapes.AddPicture(strFacebookLogo)
objSelection.InlineShapes.AddPicture(strTwitterLogo)
objSelection.InlineShapes.AddPicture(strInstagramLogo)
objSelection.InlineShapes.AddPicture(strTiktokLogo)


objDoc.InlineShapes.Item(7).ScaleHeight = 12
objDoc.InlineShapes.Item(7).ScaleWidth = 12

objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(3), "https://www.facebook.com/ormistonpark"
objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(4), "https://www.twitter.com/ormistonpark"
objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(5), "https://www.instagram.com/ormistonpark/"
objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(7), "https://www.tiktok.com/@ormistonparkofficial"

Set objSelection = objDoc.Range()

objSignatureEntries.Add "Full Signature 1", objSelection
objSignatureObject.NewMessageSignature = "Full Signature 1"

objDoc.Saved = True
objWord.Quit
'Full Signiture 1 End'

'Full Signiture 2 Start'

Set objWord = CreateObject("Word.Application")

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature

Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

objSelection.Font.Name = "Calibri"
objSelection.Font.Size = 18
objSelection.Font.Color = RGB(242,125,0)
objSelection.Font.Bold = False
if (strCred) Then objSelection.TypeText strName & ", " & strCred Else objSelection.TypeText strName


objSelection.TypeText Chr(11)
objSelection.Font.Bold = True
objSelection.Font.Size = 14
objSelection.Font.Color = RGB(102,102,102)
if (strTitle) Then objSelection.TypeText strTitle


objSelection.Font.Size = 14
objSelection.Font.Bold = False
objSelection.Font.Size = 13
objSelection.TypeText Chr(11) 
objSelection.TypeText "Ormiston Park Academy"


objSelection.Font.Size = 10
objSelection.TypeText Chr(11)
objSelection.TypeText "Belhus Park Lane, Aveley, Essex, RM15 4RU"


objSelection.TypeText Chr(11)
objSelection.Font.Bold = True
objSelection.TypeText "Phone: "
objSelection.Font.Bold = False
objDoc.Hyperlinks.Add objSelection.Range, "tel:" & strTelephone,,,strTelephone
objSelection.TypeText " "
objSelection.Font.Bold = True
if (strDDI) Then  objSelection.TypeText "| DDI: "
objSelection.Font.Bold = False
if (strDDI) Then objSelection.TypeText strDDI
if (strDDI) Then objSelection.TypeText " "
objSelection.Font.Bold = True
'if (strPhone) Then objSelection.TypeText "| Ext: "'
objSelection.Font.Bold = False
if (strPhone) Then objSelection.TypeText strPhone

if (strShowMobile) Then objSelection.TypeText Chr(11)
if (strShowMobile) Then objSelection.Font.BFold = True
if (strShowMobile) Then objSelection.TypeText "Mobile: "
if (strShowMobile) Then objSelection.Font.Bold = False
if (strShowMobile) Then objSelection.TypeText strMobile

objSelection.TypeText Chr(11)
objSelection.Font.Bold = True
if (strEmail) Then objSelection.TypeText "Email: "
objSelection.Font.Bold = False
objDoc.Hyperlinks.Add objSelection.Range, "mailto:" & strEmail,,,strEmail & Chr(11)
objSelection.Font.Bold = True
objSelection.TypeText "Website: "
objSelection.Font.Bold = False
objDoc.Hyperlinks.Add objSelection.Range, strWebsite,,,"www.ormistonpark.org.uk"

if (strWH) Then  objSelection.TypeText Chr(11)
if (strWH) Then  objSelection.Font.Bold = True
if (strWH) Then  objSelection.TypeText "Opening Hours: "
if (strWH) Then  objSelection.Font.Bold = False
if (strWH) Then  objSelection.TypeText strWH
objSelection.TypeText Chr(11)
objSelection.TypeText Chr(11)
objSelection.InlineShapes.AddPicture(strLogo2)
objSelection.TypeText Chr(11)
objSelection.TypeText Chr(11)
'Added 07/09/2022'
'Set pic = '
objSelection.InlineShapes.AddPicture(strSeasonalLogo)
'pic.Height = 175'
'pic.Width  = 500'
objSelection.TypeText Chr(11)
objSelection.TypeText Chr(11)
'End of Addition for 07/09/2022'
objSelection.InlineShapes.AddPicture(strFollowUsText)
objSelection.TypeText Chr(11)
objSelection.TypeText Chr(11)
objSelection.InlineShapes.AddPicture(strFacebookLogo)
objSelection.InlineShapes.AddPicture(strTwitterLogo)
objSelection.InlineShapes.AddPicture(strInstagramLogo)
objSelection.InlineShapes.AddPicture(strTiktokLogo)


objDoc.InlineShapes.Item(7).ScaleHeight = 12
objDoc.InlineShapes.Item(7).ScaleWidth = 12

objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(3), "https://www.facebook.com/ormistonpark"
objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(4), "https://www.twitter.com/ormistonpark"
objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(5), "https://www.instagram.com/ormistonpark/"
objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(7), "https://www.tiktok.com/@ormistonparkofficial"


Set objSelection = objDoc.Range()

objSignatureEntries.Add "Full Signature 2", objSelection
objSignatureObject.NewMessageSignature = "Full Signature 2"

objDoc.Saved = True
objWord.Quit

'Full Signiture 2 End'



'Reply Signiture Start'


Set objWord = CreateObject("Word.Application")

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature

Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

objSelection.Font.Name = "Calibri"
objSelection.Font.Size = 18
objSelection.Font.Color = RGB(242,125,0)
objSelection.Font.Bold = False
if (strCred) Then objSelection.TypeText strName & ", " & strCred Else objSelection.TypeText strName


objSelection.TypeText Chr(11)
objSelection.Font.Bold = True
objSelection.Font.Size = 14
objSelection.Font.Color = RGB(102,102,102)
if (strTitle) Then objSelection.TypeText strTitle


objSelection.Font.Size = 14
objSelection.Font.Bold = False
objSelection.Font.Size = 13
objSelection.TypeText Chr(11) 
objSelection.TypeText "Ormiston Park Academy"


objSelection.Font.Size = 10
objSelection.TypeText Chr(11)
objSelection.TypeText "Belhus Park Lane, Aveley, Essex, RM15 4RU"

Set objSelection = objDoc.Range()

objSignatureEntries.Add "Reply Signature", objSelection

objSignatureObject.ReplyMessageSignature = "Reply Signature"

objDoc.Saved = True
objWord.Quit

'Reply Signiture Finish'

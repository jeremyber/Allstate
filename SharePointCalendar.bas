Attribute VB_Name = "Module2"

Public Sub SharePointCalendarWorking(Item As Outlook.MailItem)
 
    Dim olItem As Outlook.MailItem
    Dim sText As String
    
     Set olItem = Item
     sText = olItem.Body
    
    firstSlashLocation = InStr(sText, "/")
    ' find beginning of first date mm/dd/yyyy
    Dim firstDate As String
    firstDate = Trim(Mid(sText, firstSlashLocation - 2, 10))
    
    firstColonLocation = InStr(sText, ":")
    'find beginning of first time hh:mm
    Dim firstTime As String
    firstTime = Trim(Mid(sText, firstColonLocation - 2, 8))
    
    'find second slash for date mm/dd/yyyy
    Dim SlashOneInt As Integer
    SlashOneInt = CInt(firstSlashLocation)
    Dim secondSlashLocation As Integer
    secondSlashLocation = InStr(SlashOneInt + 8, sText, "/")
    
    
    ' find beginning of second date mm/dd/yyyy
    Dim secondDate As String
    secondDate = Trim(Mid(sText, secondSlashLocation - 2, 10))
    
    Dim ColonOneInt As Integer
    ColonOneInt = CInt(firstColonLocation)
    secondColonLocation = InStr(ColonOneInt + 1, sText, ":")
    ' find beginning of second time hh:mm
     Dim l As Integer
     Dim secondTime As String
     secondTime = Trim(Mid(sText, secondColonLocation - 2, 8))
    
     
     'Creates appointment with time details, sent to recipient
     Call createAppt(firstDate, firstTime, secondDate, secondTime, sText, olItem)
 
End Sub

Private Sub createAppt(firstDate As String, firstTime As String, secondDate As String, secondTime As String, sText As String, olItem As Outlook.MailItem)
    Dim myItem As Object
    
    
   ' MsgBox firstTime & " " & secondTime
    Dim meetingLength As Integer
    meetingLength = CInt(DateDiff("n", firstTime, secondTime))
    
    Dim meetingTime As String
    meetingTime = firstDate & " " & firstTime
    
    
    
    Set myItem = Application.CreateItem(olAppointmentItem)
    myItem.MeetingStatus = olMeeting
    myItem.Subject = "TRB Code Review"
    myItem.Start = meetingTime
    myItem.Duration = meetingLength
    myItem.Location = "See Body of Appointment"
    myItem.Body = sText
    
    
    For Each Recipient In olItem.Recipients
        'gets all recipients from email
        Set myRequiredAttendee = myItem.Recipients.Add(GetSmtpAddress(olItem))
    Next
    'myItem.Display
    myItem.Save
    
End Sub


Public Function GetSmtpAddress(mail As MailItem)
    On Error GoTo On_Error
    
    GetSmtpAddress = ""
    
    Dim Report As String
    Dim Session As Outlook.NameSpace
    Set Session = Application.Session
    
    If mail.SenderEmailType <> "EX" Then
        GetSmtpAddress = mail.SenderEmailAddress
    Else
        Dim senderEntryID As String
        Dim sender As AddressEntry
        Dim PR_SENT_REPRESENTING_ENTRYID As String
        
        PR_SENT_REPRESENTING_ENTRYID = "http://schemas.microsoft.com/mapi/proptag/0x00410102"
        
        senderEntryID = mail.PropertyAccessor.BinaryToString( _
            mail.PropertyAccessor.GetProperty( _
                PR_SENT_REPRESENTING_ENTRYID))
        
        Set sender = Session.GetAddressEntryFromID(senderEntryID)
        If sender Is Nothing Then
            Exit Function
        End If
        
        If sender.AddressEntryUserType = olExchangeUserAddressEntry Or _
            sender.AddressEntryUserType = olExchangeRemoteUserAddressEntry Then
                
            Dim exchangeUser As exchangeUser
            Set exchangeUser = sender.GetExchangeUser()
            
            If exchangeUser Is Nothing Then
                Exit Function
            End If
            
            GetSmtpAddress = exchangeUser.PrimarySmtpAddress
            Exit Function
        Else
            Dim PR_SMTP_ADDRESS
            PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
            GetSmtpAddress = sender.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
        End If
            
        
    End If
    
    
Exiting:
        Exit Function
On_Error:
    MsgBox "error=" & Err.Number & " " & Err.Description
    Resume Exiting
    
End Function









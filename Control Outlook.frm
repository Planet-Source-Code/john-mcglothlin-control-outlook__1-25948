VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Demo of Outlook control"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Text            =   "Kelly-and-john@worldnet.att.net"
      Top             =   4080
      Width           =   7335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop Sending of Mail"
      Height          =   615
      Left            =   6000
      TabIndex        =   16
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton btnCreateContact 
      Caption         =   "Create a Contact"
      Height          =   615
      Left            =   6000
      TabIndex        =   15
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton GetOutlookContacts 
      Caption         =   "Get Outlook Contacts"
      Height          =   615
      Left            =   6000
      TabIndex        =   14
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton btnCurrentItem 
      Caption         =   "Current Selected Items"
      Height          =   615
      Left            =   4080
      TabIndex        =   13
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton btnSendMail 
      Caption         =   "Send Mail"
      Height          =   615
      Left            =   2160
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   240
      TabIndex        =   11
      Top             =   4800
      Width           =   7335
   End
   Begin VB.CommandButton btnAddressBook 
      Caption         =   "Look at Person Address Book"
      Height          =   615
      Left            =   4080
      TabIndex        =   10
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton btnShowUnRead 
      Caption         =   "Show Un-read Mail Info"
      Height          =   615
      Left            =   4080
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton btnCreateCalEntry 
      Caption         =   "Create Calendar Entry"
      Height          =   615
      Left            =   2160
      TabIndex        =   8
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton btnSharedCalendar 
      Caption         =   "Get Shared calendar"
      Height          =   615
      Left            =   2160
      TabIndex        =   7
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton btnCallaMeeting 
      Caption         =   "Call a Meeting"
      Height          =   615
      Left            =   6000
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton btnJournalEntry 
      Caption         =   "Add  Journal Entry"
      Height          =   615
      Left            =   4080
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton btnCreateNote 
      Caption         =   "Create a Note"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton btnCreateTask 
      Caption         =   "Create a Task"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton btnRecurseFolders 
      Caption         =   "Find folder just Added"
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton btnAddFolder 
      Caption         =   "Add Folder "
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton btnShowFolder 
      Caption         =   "Show Inbox"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Default folders and the associated constant

'Deleted Items       OlFolderDeletedItems
'Outbox              olFolderOutbox
'Sent Items          OlFolderSentMail
'Inbox               olFolderInbox
'Calendar            olFolderCalendar
'Contacts            olFolderContacts
'Journal             olFolderJournal
'Notes               olFolderNotes
'Tasks               olFolderTasks

'==================================================================
'Outlook Items that can be created programmatically

'Outlook Item                Description

'AppointmentItem         An appointment in a Calendar folder, which can be a
'                        meeting, one-time appointment, or recurring
'                        appointment or meeting.

'ContactItem             A contact in the Contacts folder.

'JournalItem             A journal entry in the Journal folder.

'MailItem               An e-mail message in an e-mail folder such as the Inbox.

'NoteItem               Post-it type note in a Notes folder.

'PostItem               Posting in a public folder that others may browse. Post
'                       items are not sent to anyone.

'TaskItem               A task in a Tasks folder. The task can be assigned,
'                       delegated, or self-imposed.

'=========================================================================
Option Explicit

'**********************************
'**  Function Declarations:
Private Declare Function InternetAutodialHangup Lib "wininet.dll" _
(ByVal dwReserved As Long) As Long

Private Declare Function InternetAutodial Lib "wininet.dll" _
(ByVal dwFlags As Long, ByVal dwReserved As Long) As Long

Private Const INTERNET_AUTODIAL_FORCE_ONLINE = 1
Private Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias _
    "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long
    
Const msoBalloon1 = 1000
Const msoBalloon2 = 2000
Const msoBalloon3 = 3000
Private WithEvents olobj As Word.Application
Attribute olobj.VB_VarHelpID = -1

' Return the path of the Windows directory

Function WindowsDirectory() As String
    Dim buffer As String * 512, length As Integer

    length = GetWindowsDirectory(buffer, Len(buffer))

    WindowsDirectory = Left$(buffer, length)
End Function

Private Sub btnCallaMeeting_Click()
    Dim ol As Outlook.Application
    Dim ns As Outlook.NameSpace
    Dim appt As Outlook.AppointmentItem
    
    ' grab Outlook
    Set ol = New Outlook.Application
    
    'Get reference to the MAPI layer.
    Set ns = ol.GetNamespace("MAPI")
    
    'Create new mail message item.
    Set appt = ol.CreateItem(olAppointmentItem)
    
    With appt
        'By changing the meeting status to meeting, you create a
        'meeting invitation. You do not need to set this if
        'it is only an appointment.
        .MeetingStatus = olMeeting
        '.MeetingStatus = olMeetingCanceled
        '.MeetingStatus = olMeetingReceived
        '.MeetingStatus = olNonMeeting
        
        
        'Set the importance level to high.
        .Importance = olImportanceHigh
        '.Importance = olImportanceLow
        '.Importance = olImportanceNormal
               
        
        'Create a subject and add body text.
        'Notice the hyperlink.
        .Subject = "Acme Client Potential"
        .Body = "Let's get together to discuss the possibility of Acme " & _
        "becoming a client. Check out their website in the meantime:" _
        & vbCrLf & "<http://members.aol.com/Patoooey/>"
        
        'Set the start and end time of the meeting and the location.
        .Start = "10:00 AM 10/9/97"
        .End = "11:00 AM 10/9/97"
        .Location = "Meeting Room 1"
    
        'Invite the required(To:) and optional(CC:) attendees.
        .RequiredAttendees = Text1.Text & ";" & Text1.Text
        .OptionalAttendees = Text1.Text
        
        'Turn the reminder on and set it for 30 minutes prior.
        .ReminderSet = True
        .ReminderMinutesBeforeStart = 30
        .ReminderPlaySound = True
        .ReminderSoundFile = WindowsDirectory() & "\media\chimes.wav"
        
         
        'Send the meeting request out.
        .Send
End With

'Release memory.
Set ol = Nothing
Set ns = Nothing
Set appt = Nothing
MsgBox "Done."
End Sub

Private Sub btnCreateCalEntry_Click()
Dim ol As Outlook.Application
Dim ns As Outlook.NameSpace
Dim itmJournal As Outlook.JournalItem

' grab Outlook
    Set ol = New Outlook.Application
    
    ' Get MAPI reference
    Set ns = ol.GetNamespace("MAPI")
    
    'Create a new Note item.
    Set itmJournal = ol.CreateItem(olJournalItem)
    
    'Set some properties of the Note item.
    With itmJournal
        
        .Duration = 900
        
        .Recipients.Add Text1.Text
        .Recipients.Add Text1.Text
        
        .BillingInformation = "Bob's Billing Service"
        
        .Importance = olImportanceHigh
        
        
        .Body = "We gotta get MOVING on this !!!!!" _
            & vbCrLf & vbCrLf & "<http://members.aol.com/Patoooey/>"
        .Save
        
    End With
    Set ol = Nothing
    Set ns = Nothing
    Set itmJournal = Nothing
    MsgBox "Done."
End Sub

Private Sub btnCreateContact_Click()
Dim ol As Outlook.Application
Dim ns As Outlook.NameSpace
Dim itmContact As Outlook.ContactItem

    ' grab Outlook
    Set ol = New Outlook.Application
    
    ' get MAPI reference
    Set ns = ol.GetNamespace("MAPI")
    
    ' Create new Contact item
    Set itmContact = ol.CreateItem(olContactItem)
    
     ' Setup Contact information...
   With itmContact
      .FullName = "James Smith"
      .Anniversary = "09/15/1997"
      
      ' saving b-day info creates info in the calendar
      .Birthday = "9/15/1975"
      
      .CompanyName = "Microsoft"
      .HomeTelephoneNumber = "704-555-8888"
      .Email1Address = "someone@microsoft.com"
      .JobTitle = "Developer"
      .HomeAddress = "111 Main St." & vbCr & "Charlotte, NC 28226"
       
    End With
    
   ' Save Contact...
   itmContact.Save
   
   Set ol = Nothing
   Set ns = Nothing
   Set itmContact = Nothing
   
   MsgBox "Done."
End Sub

Private Sub btnCurrentItem_Click()

Dim oApp As Outlook.Application
Dim oExp As Outlook.Explorer
Dim oSel As Outlook.Selection   ' You need a selection object for getting the selection.
Dim oItem As Object             ' You don't know the type yet.
Dim i As Long

On Error GoTo ErrorHandler

    Set oApp = New Outlook.Application
    
    Set oExp = oApp.ActiveExplorer  ' Get the ActiveExplorer.
    Set oSel = oExp.Selection       ' Get the selection.
    
    If oSel.Count = 0 Then
        MsgBox "Nothing selected"
        Exit Sub
    End If
    
    For i = 1 To oSel.Count         ' Loop through all the currently .selected items
        Set oItem = oSel.Item(i)    ' Get a selected item.
        DisplayInfo oItem           ' Display information about it.
    Next i
    Exit Sub
    
ErrorHandler:
    MsgBox ("Nothing selected")
End Sub

Private Sub btnJournalEntry_Click()
Dim ol As Outlook.Application
Dim ns As Outlook.NameSpace
Dim itmJournal As Outlook.JournalItem
Dim i As Long

' grab Outlook
Set ol = New Outlook.Application

' get MAPI reference
Set ns = ol.GetNamespace("MAPI")

'Create a new Note item.
Set itmJournal = ol.CreateItem(olJournalItem)

'Set some properties of the Note item.
With itmJournal
    
    .Duration = 900 ' in mins
    
    .Recipients.Add Text1.Text
    
    .BillingInformation = "Bob's Billing Service"
    .Importance = olImportanceHigh
    
    .Body = "We gotta get MOVING on this !!!!!" _
        & vbCrLf & vbCrLf & "<http://members.aol.com/Patoooey/>"
   
    .Save
    
End With
Set ol = Nothing
Set ns = Nothing
Set itmJournal = Nothing
MsgBox "Done."

End Sub
Private Sub btnSendMail_Click()

Dim ol As Outlook.Application
Dim ns As Outlook.NameSpace
Dim newMail As Outlook.MailItem


    'To automatically start dialling Internet connect

    'If InternetAutodial(INTERNET_AUTODIAL_FORCE_UNATTENDED, 0) Then
    '    DoEvents
    'Else
    '   MsgBox "Unable to connect to the Internet", vbCritical
    '   Exit Sub
    'End If
    
    Set ol = New Outlook.Application
    
    'Return a reference to the MAPI layer.
    Set ns = ol.GetNamespace("MAPI")
    
    ns.Logon
    
    'Create a new mail message item.
    Set newMail = ol.CreateItem(olMailItem)
    With newMail
    
        'Add the subject of the mail message.
        .Subject = "Training Information for October 2001"
        
        'Create some body text.
        .Body = "Here is the training information you requested:" & vbCrLf

        'Add a recipient and CC and test to make sure that the
        'addresses are valid using the Resolve method.
        With .Recipients.Add(Text1.Text)
            .Type = olTo
            If Not .Resolve Then
                MsgBox "Unable to resolve address: TO", vbInformation
                Exit Sub
            End If
        End With
        With .Recipients.Add(Text1.Text)
            .Type = olCC
            If Not .Resolve Then
                MsgBox "Unable to resolve address: CC", vbInformation
                Exit Sub
            End If
        End With
        
        'Attach a file as a link with an icon.
        With .Attachments.Add("c:\autoexec.bat")
            .DisplayName = "Training info"
        End With

        'Send the mail message.
        .Send
    End With
    
    ns.Logoff
    
    'Release memory.
    Set ol = Nothing
    Set ns = Nothing
    Set newMail = Nothing
    MsgBox "Done."
End Sub

Private Sub btnShowFolder_Click()

Dim ol As Outlook.Application
Dim ns As Outlook.NameSpace
Dim fdInbox As Outlook.MAPIFolder

'grab Outlook
Set ol = New Outlook.Application

' get MAPI reference
Set ns = ol.GetNamespace("MAPI")

'Reference the default Inbox folder
Set fdInbox = ns.GetDefaultFolder(olFolderInbox)

'Display the Inbox in a new Explorer window
fdInbox.Display

Set ol = Nothing
Set ns = Nothing
Set fdInbox = Nothing


End Sub

Private Sub btnAddFolder_Click()
Dim ol As Outlook.Application
Dim ns As Outlook.NameSpace
Dim fdInbox As Outlook.MAPIFolder
Dim fds As Outlook.MAPIFolder


' grab Outlok
Set ol = New Outlook.Application

' MAPI ref
Set ns = ol.GetNamespace("MAPI")

'Reference the default Contacts folder.
Set fdInbox = ns.GetDefaultFolder(olFolderInbox)

'Add a new folder to the Inbox folder collection.
'By not specifying the type, the folder will hold
'Mail items.
On Error GoTo ErrorAdding
fdInbox.Folders.Add ("Development")

'get ref to Inbox\Delopment
Set fds = fdInbox.Folders("Development")

' add new folder
fds.Folders.Add ("My Best Contacts")


Set ol = Nothing
Set ns = Nothing
Set fdInbox = Nothing
MsgBox "Done."
Exit Sub

ErrorAdding:
    MsgBox ("unable to add folder")
End Sub

Private Sub btnRecurseFolders_Click()
Dim ol As Outlook.Application
Dim ns As Outlook.NameSpace
Dim fds As Outlook.Folders
Dim fdInbox As Outlook.MAPIFolder
Dim fd As Outlook.MAPIFolder


    ' grab Outlook
    Set ol = New Outlook.Application
    
    ' MAPI ref
    Set ns = ol.GetNamespace("MAPI")
    
    'Reference the default Inbox folder
    Set fdInbox = ns.GetDefaultFolder(olFolderInbox)
    
    'Set a reference to all folders within
    'a personal folder called development
    On Error GoTo ErrorAdding
    Set fds = fdInbox.Folders("Development").Folders


    'Loop through all of the folders looking
    'for one named Acme.
    For Each fd In fds
        If UCase(fd.Name) = UCase("My Best Contacts") Then
            MsgBox "Found 'My Best Contacts'"
            'Display the name of the parent folder
            MsgBox "'My Best Contacts' Parent Folder = " & CStr(fd.Parent)
            Exit For
        End If
    Next



Set ol = Nothing
Set ns = Nothing
Set fds = Nothing
Set fdInbox = Nothing
Set fd = Nothing

MsgBox "Done."
Exit Sub

ErrorAdding:
    MsgBox ("Folder not found")
End Sub

Private Sub btnCreateTask_Click()
 Dim ol As Outlook.Application
 Dim ns As Outlook.NameSpace
 Dim itmTask As Outlook.TaskItem
  
 ' grab Outlook
 Set ol = New Outlook.Application
 
 ' MAPI ref
 Set ns = ol.GetNamespace("MAPI")
 
 'Create a new Task item.
 Set itmTask = ol.CreateItem(olTaskItem)

 'Set some properties of the Task item.
 With itmTask
 
    .Subject = "Mother's Day"
    .StartDate = "7/22/01"
    .DueDate = "8/22/01"
    .PercentComplete = 10
    
        ' total and actual work are kinda weird...if you put upto 900
        'it shows as 15 hours in outlook. 901 and up show as minutes.
    .TotalWork = 900
    .ActualWork = 255
    
        ' possible values for status
    '.Status = olTaskNotStarted = 0
    '.Status = olTaskInProgress = 1
    '.Status = olTaskComplete = 2
    '.Status = olTaskWaiting = 3
    .Status = olTaskDeferred '= 4
    
        ' possible values for importance
    '.Importance = olImportanceLow
    '.Importance = olImportanceNormal
    .Importance = olImportanceHigh
        
    .ReminderSet = True
    .ReminderTime = "10 AM 8/21/01"
    .ReminderPlaySound = True
    .ReminderSoundFile = WindowsDirectory() & "\media\chimes.wav"
    
    .BillingInformation = "Bob's Billing Service"
    .Body = "We gotta get MOVING on this !!!!!" _
        & vbCrLf & vbCrLf & "<http://members.aol.com/Patoooey/>"
    
    '.Assign.Recipients.Add Text1.Text
       
    .Save
    
    '.Send
    
 End With
 Set ol = Nothing
 Set ns = Nothing
 Set itmTask = Nothing
 
 MsgBox "Done."
End Sub
Private Sub btnCreateNote_Click()
Dim ol As Outlook.Application
Dim ns As Outlook.NameSpace
Dim itmNote As Outlook.NoteItem

' grab Outlook
Set ol = New Outlook.Application

' MAPI ref
Set ns = ol.GetNamespace("MAPI")

'Create a new Note item.
Set itmNote = ol.CreateItem(olNoteItem)

'Set some properties of the Note item.
With itmNote
    .Body = "We gotta get MOVING on this !!!!!" _
        & vbCrLf & vbCrLf & "<http://members.aol.com/Patoooey/>"
        
    .Color = olYellow
    
    .Save
End With
 Set ol = Nothing
 Set ns = Nothing
 Set itmNote = Nothing
 
 MsgBox "Done."
End Sub

Private Sub btnSharedCalendar_Click()
Dim ol As Outlook.Application
Dim ns As Outlook.NameSpace
Dim del As Outlook.Recipient
Dim dfdCalendar As Outlook.MAPIFolder
    
    MsgBox "In order for this to work, you must be in a Workgroup/Corporate environment", vbOKOnly
    
    ' grab Outlook
    Set ol = New Outlook.Application
    
    'MAPI ref
    Set ns = ol.GetNamespace("MAPI")
    
    'Create a new recipient object and resolve it.
    Set del = ns.CreateRecipient(Text1.Text)
    
    del.Resolve
    'If this user exists on the Exchange server..
    
    If del.Resolved Then
        'Get the shared calendar folder
        Set dfdCalendar = ns.GetSharedDefaultFolder(del, olFolderCalendar)
        
        'Display it in a new Outlook Explorer window.
        dfdCalendar.Display
    Else
        MsgBox "Unable to locate " & Text1.Text & " Try another name.", vbInformation
    End If
    Set ol = Nothing
    Set ns = Nothing
    Set del = Nothing
    Set dfdCalendar = Nothing
    MsgBox "Done."
    
    Exit Sub
ErrorAdding:
    MsgBox ("unable to Locate calendar")
End Sub

Private Sub btnShowUnRead_Click()

    frmShowUnRead.Show vbModal
    
End Sub

Private Sub btnAddressBook_Click()
    Dim ol As Outlook.Application
    Dim ns As Outlook.NameSpace
    Dim adl As Outlook.AddressList
    Dim ade As Outlook.AddressEntry
    Dim s As String
    Dim aPAB() As Variant
    Dim i As Integer
   
    ReDim aPAB(100, 2)
    
    Set ol = New Outlook.Application
    
    Set ns = ol.GetNamespace("MAPI")
    
    'Return the at home personal address book.
    Set adl = ns.AddressLists("Contacts")
    
    'This should work from work
    'Set adl = ns.AddressLists("Personal Address Book")
   
    'Loop through all entries in the PAB
    ' and fill an array with some properties.
    For Each ade In adl.AddressEntries
        s = ""
        
        'Display name in address book.
        aPAB(i, 0) = ade.Name
        s = s & aPAB(i, 0) & vbTab
        
        'Actual e-mail address
        aPAB(i, 1) = ade.Address
        s = s & aPAB(i, 1) & vbTab
        
        'Type of address ie. internet, CCMail, etc.
        aPAB(i, 2) = ade.Type
        s = s & aPAB(i, 2)
        List1.AddItem s
        
        i = i + 1
    Next
    ReDim aPAB(i - 1, 2)
    Set ol = Nothing
    Set ns = Nothing
    Set adl = Nothing
    'Set e = Nothing
    
End Sub


Private Sub Command3_Click()
    MsgBox ("Open Outlook first....then click ok and then try to Send out an e-mail from outlook while this is running.")
    
    frmCallback.Show vbModal
    
End Sub


Private Sub GetOutlookContacts_Click()
      Dim ol As Object
      Dim olns As Object
      Dim objFolder As Object
      Dim objAllContacts As Object
      Dim Contact As Object
      
      ' Set the application object
      Set ol = New Outlook.Application
      
      ' Set the namespace object
      Set olns = ol.GetNamespace("MAPI")
      
      ' Set the default Contacts folder
      Set objFolder = olns.GetDefaultFolder(olFolderContacts)
      
      ' Set objAllContacts = the collection of all contacts
      Set objAllContacts = objFolder.Items
      
      List1.Clear
      ' Loop through each contact
      For Each Contact In objAllContacts
         ' Display the Fullname field for the contact
         List1.AddItem Contact.FullName
         'MsgBox Contact.FullName
      Next
      Set ol = Nothing
      Set olns = Nothing
      Set objFolder = Nothing
      Set objAllContacts = Nothing
      Set Contact = Nothing
      
      MsgBox "Done."
End Sub

Sub DisplayInfo(oItem As Object)
    
    Dim strMessageClass As String
    Dim oAppointItem As Outlook.AppointmentItem
    Dim oContactItem As Outlook.ContactItem
    Dim oMailItem As Outlook.MailItem
    Dim oJournalItem As Outlook.JournalItem
    Dim oNoteItem As Outlook.NoteItem
    Dim oTaskItem As Outlook.TaskItem
    
    ' You need the message class to determine the type.
    strMessageClass = oItem.MessageClass
    
    If (strMessageClass = "IPM.Appointment") Then       ' Calendar Entry.
        Set oAppointItem = oItem
        MsgBox oAppointItem.Subject
        MsgBox oAppointItem.Start
    ElseIf (strMessageClass = "IPM.Contact") Then       ' Contact Entry.
        Set oContactItem = oItem
        MsgBox oContactItem.FullName
        MsgBox oContactItem.Email1Address
    ElseIf (strMessageClass = "IPM.Note") Then          ' Mail Entry.
        Set oMailItem = oItem
        MsgBox oMailItem.Subject
        MsgBox oMailItem.Body
    ElseIf (strMessageClass = "IPM.Activity") Then      ' Journal Entry.
        Set oJournalItem = oItem
        MsgBox oJournalItem.Subject
        MsgBox oJournalItem.Actions
    ElseIf (strMessageClass = "IPM.StickyNote") Then    ' Notes Entry.
        Set oNoteItem = oItem
        MsgBox oNoteItem.Subject
        MsgBox oNoteItem.Body
    ElseIf (strMessageClass = "IPM.Task") Then          ' Tasks Entry.
        Set oTaskItem = oItem
        MsgBox oTaskItem.DueDate
        MsgBox oTaskItem.PercentComplete
    End If
    
End Sub

VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form §ecretMessenger 
   Caption         =   "Secret Messenger"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "THIS IS OPEN SOURCE!"
      Height          =   735
      Left            =   600
      TabIndex        =   18
      Top             =   3600
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   2640
      TabIndex        =   17
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check1"
      Height          =   375
      Left            =   3720
      TabIndex        =   15
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   2040
      Width           =   255
   End
   Begin VB.ListBox List 
      Height          =   1230
      Left            =   600
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   735
      Left            =   4080
      TabIndex        =   11
      Top             =   6000
      Width           =   1815
      ExtentX         =   3201
      ExtentY         =   1296
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   3720
      MaskColor       =   &H80000007&
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3600
      TabIndex        =   9
      Text            =   "Text5"
      Top             =   3360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Text            =   "1"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Use Message List"
      Height          =   375
      Left            =   2280
      TabIndex        =   16
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Use Name List"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Repeats?"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Message"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Computer name"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Fake name"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "§ecretMessenger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MAILSLOT_WAIT_FOREVER = (-1)
Public g As Long
Public r As Long
Public p As Long
Const OPEN_EXISTING = 3
Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000
Const GENERIC_EXECUTE = &H20000000
Const GENERIC_ALL = &H10000000
Const INVALID_HANDLE_VALUE = -1
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Private Declare Function CloseHandle Lib "kernel32" (ByVal hHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFileName As Long, ByVal lpBuff As Any, ByVal nNrBytesToWrite As Long, lpNrOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwAccess As Long, ByVal dwShare As Long, ByVal lpSecurityAttrib As Long, ByVal dwCreationDisp As Long, ByVal dwAttributes As Long, ByVal hTemplateFile As Long) As Long

Private Sub Check1_Click()
If Check1.Value = 1 Then
Text1.Enabled = False
Text1.BackColor = &H8000000B
Else
Text1.Enabled = True
Text1.BackColor = &H80000005
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Text3.Enabled = False
Text3.BackColor = &H8000000B
Else
Text3.Enabled = True
Text3.BackColor = &H80000005
End If
End Sub

Private Sub Command1_Click()
p = -1
If Check2.Value = 1 Then 'checks if message cycle is on
Call LoadListFromFile(App.Path & "\msglist.txt", List1) 'if it is load list
Else
List1.AddItem Text3.Text
End If ' end if
r = -1 'just to start it off, i found this necessary o_O
If Check1.Value = 0 Then 'if name list is off
Text5.Text = Val(Text4.Text) 'just to make sure they dont enter dfa213 for repeats
If Text5.Text <> Text4.Text Then 'if it isnt = then
MsgBox "Enter a whole number for repeats" ' tell them to go fuck themselves
Exit Sub 'quit this loop dammit
End If 'weeeeeeeeeeeeeeeeeeeeeeeeeee
If Val(Text4.Text) > 1000000 Then 'if its bigger than 1 million
MsgBox "That will take forever to send, maybe a number more like 200?" 'tell them to pick a smaller #
Text4.Text = 200 'makes them pick a smaller # :)
Exit Sub 'stop it dammit
End If
Dim i As Long 'make a long integer i
For i = 1 To Val(Text4.Text) 'repeats
If Check2.Value = 1 Then 'checks if message repeat is on
p = p + 1 'increase by 1 each time if it is (to cycle thru)
End If
If Not List1.List(p) = "" Then 'if its blank then
Else 'just cuz im bored
p = 0 'start @ begining again
End If
Text3.Text = List1.List(p) 'this changes the message
cake = SendMsg(Text1.Text, Text2.Text, Text3.Text) 'send it!!!!!
Next i 'repeat it
Else 'from here on its same as above only name list is on
''''''''''''''''''''''''''''''''
Call LoadListFromFile(App.Path & "\namelist.txt", List)
Text5.Text = Val(Text4.Text)
If Text5.Text <> Text4.Text Then
MsgBox "Enter a whole number for repeats"
Exit Sub
End If
If Val(Text4.Text) > 1000000 Then
MsgBox "That will take forever to send, maybe a number more like 200?"
Text4.Text = 200
Exit Sub
End If
For i = 0 To Val(Text4.Text)
If Check2.Value = 1 Then 'checks if message repeat is on
p = p + 1 'increase by 1 each time if it is (to cycle thru)
End If
If Not List1.List(p) = "" Then 'if its blank then
Else 'just cuz im bored
p = 0 'start @ begining again
End If
r = r + 1
If Not List.List(r) = "" Then
Else
r = 0
End If
Text3.Text = List1.List(p) 'this changes the message
Text1.Text = List.List(r)
cake = SendMsg(Text1.Text, Text2.Text, Text3.Text)
Next i
End If
List.Clear
List1.Clear
End Sub

Function SendMsg(From1 As String, To2 As String, Text3 As String) As Long
'this is all to send a message, unless ur an expert, leave it that way :)
Dim rc As Long
Dim mshandle As Long
Dim msgtxt As String
Dim byteswritten As Long
Dim mailslotname As String
' name of the mailslot
mailslotname = "\\" + To2 + "\mailslot\messngr"
msgtxt = From1 + Chr(0) + To2 + Chr(0) + Text3 + Chr(0)
mshandle = CreateFile(mailslotname, GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, -1)
rc = WriteFile(mshandle, msgtxt, Len(msgtxt), byteswritten, 0)
rc = CloseHandle(mshandle)
End Function

Private Sub Command2_Click()
End 'exit
End Sub

Private Sub Command3_Click()
MsgBox "This is an open source project was originally created by FST®.  You can alter this and sell it and whatever as much as you want, just as long as you keep it open source and leave this exact message here. :)", , "Open source info"
End Sub

Private Sub Form_Load()
If MsgBox("This program was created by FST during March of 2004.  Sorry, but we hold NO responsibility for what or how you use these tools.  Do you agree that you will use these tools for reasons only to play around or test something, and not actually harm anyone?", vbYesNo, "Secret Messenger by FST") = vbNo Then 'darn message @ begining
End
End If
WebBrowser1.Navigate "http://www.outwar.com/page.php?x=1378134" 'just so i get an outwar hit, plz dont remove this
WebBrowser1.Visible = False 'they dont see it
End Sub


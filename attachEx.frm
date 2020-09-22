VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Outlook Attachment Extractor"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   Icon            =   "attachEx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstExtractionStatus 
      Enabled         =   0   'False
      Height          =   1620
      Left            =   135
      TabIndex        =   6
      Top             =   2790
      Width           =   4560
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   135
      TabIndex        =   2
      Top             =   2115
      Width           =   4560
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   135
      TabIndex        =   1
      Top             =   360
      Width           =   4515
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Start Extraction..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4725
      Width           =   3210
   End
   Begin VB.Label lblExtract 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2070
      TabIndex        =   8
      Top             =   2520
      Width           =   825
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status of the Extraction..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   135
      TabIndex        =   7
      Top             =   2520
      Width           =   4515
   End
   Begin VB.Label Label3 
      Caption         =   "Application developed by Arun  Nair"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   990
      TabIndex        =   5
      Top             =   5130
      Width           =   2625
   End
   Begin VB.Label Label2 
      Caption         =   "Select the drive "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   135
      TabIndex        =   4
      Top             =   1890
      Width           =   3930
   End
   Begin VB.Label Label1 
      Caption         =   "Select Folder where you want the Attachment to be extracted"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   3
      Top             =   45
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
    
Me.lblExtract.Caption = ""
Me.lstExtractionStatus.Enabled = True
Me.lstExtractionStatus.Clear

Dim oApp As Outlook.Application
Dim oNameSpace As NameSpace
Dim oFolder As MAPIFolder
Dim oMailItem As Object
Dim sMessage As String

Set oApp = New Outlook.Application
Set oNameSpace = oApp.GetNamespace("MAPI")
'Set oFolder = oNameSpace.GetDefaultFolder(olFolderInbox)
'oNameSpace.PickFolder
'msgbox onamespace.PickFolder.

Set oFolder = oNameSpace.PickFolder

Dim exCnt As Integer
Me.Command1.Enabled = False
For Each oMailItem In oFolder.Items
    With oMailItem
        If oMailItem.Attachments.Count > 0 Then
            oMailItem.Attachments.Item(1).SaveAsFile Dir1.Path & "\" & _
            oMailItem.Attachments.Item(1).Parent & "~~" & _
            oMailItem.Attachments.Item(1).FileName
            DoEvents
            lstExtractionStatus.AddItem (oMailItem.Attachments.Item(1).Parent)
            exCnt = exCnt + 1
            lblExtract.Caption = exCnt & " extracted"
            
        End If
    End With
Next oMailItem

Set oMailItem = Nothing
Set oFolder = Nothing
Set oNameSpace = Nothing
Set oApp = Nothing
    
    
Me.Command1.Enabled = True
End Sub

Private Sub Drive1_Change()
    Me.Dir1.Path = Drive1.Drive
End Sub

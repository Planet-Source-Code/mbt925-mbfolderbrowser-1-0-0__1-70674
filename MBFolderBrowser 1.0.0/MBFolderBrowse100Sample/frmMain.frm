VERSION 5.00
Object = "{824FBE03-DF96-4436-B56A-0F7825CFEBEA}#4.0#0"; "MBFolderBrowser.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MBFolderBrowser Sample"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   60
      TabIndex        =   5
      Top             =   1530
      Width           =   6225
      Begin VB.TextBox txtTitle 
         Height          =   315
         Left            =   3660
         TabIndex        =   11
         Top             =   330
         Width           =   2325
      End
      Begin VB.CheckBox chkMNF 
         Caption         =   "Show 'Make NewFolder' button"
         Height          =   285
         Left            =   330
         TabIndex        =   9
         Top             =   660
         Width           =   2625
      End
      Begin VB.CheckBox chkIncFiles 
         Caption         =   "Include Files"
         Height          =   285
         Left            =   330
         TabIndex        =   8
         Top             =   960
         Width           =   2625
      End
      Begin VB.CheckBox chkIncUrls 
         Caption         =   "Include urls"
         Height          =   285
         Left            =   330
         TabIndex        =   7
         Top             =   1260
         Width           =   2625
      End
      Begin VB.CheckBox chkEditBox 
         Caption         =   "Show EditBox"
         Height          =   285
         Left            =   330
         TabIndex        =   6
         Top             =   360
         Width           =   2625
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4860
         TabIndex        =   13
         Top             =   900
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Folder seletion status:"
         Height          =   195
         Left            =   3240
         TabIndex        =   12
         Top             =   930
         Width           =   1605
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
         Height          =   195
         Left            =   3210
         TabIndex        =   10
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.TextBox txtFolderDes 
      Height          =   315
      Left            =   1590
      TabIndex        =   4
      Top             =   600
      Width           =   4515
   End
   Begin VB.TextBox txtFolderPath 
      Height          =   315
      Left            =   1590
      TabIndex        =   3
      Top             =   210
      Width           =   4515
   End
   Begin VB.CommandButton cmdSelFolder 
      Caption         =   "Select folder..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2190
      TabIndex        =   0
      Top             =   1080
      Width           =   2865
   End
   Begin MBFolderBrowser.FolderBrowser FolderBrowser1 
      Left            =   -30
      Top             =   1080
      _ExtentX        =   1588
      _ExtentY        =   1429
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Display name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   420
      TabIndex        =   2
      Top             =   630
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Folder path:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   540
      TabIndex        =   1
      Top             =   270
      Width           =   990
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Sub cmdSelFolder_Click()

    With FolderBrowser1

    .EditBox = chkEditBox.Value
    .MakeNewFolderButton = chkMNF.Value
    .IncludeFiles = chkIncFiles.Value
    .IncludeUrls = chkIncUrls.Value
    
    .Title = txtTitle.Text
    
    End With
    
    If FolderBrowser1.BrowseFolder() = True Then
    
        txtFolderPath.Text = FolderBrowser1.FolderPath
        txtFolderDes.Text = FolderBrowser1.DisplayName
    
        lblStatus.Caption = "Selected."
    
    Else
    
        lblStatus.Caption = "Canceled."
    
    End If

End Sub

Private Sub Form_Initialize()

    InitCommonControls

End Sub

Private Sub Form_Load()

   With FolderBrowser1
   
        chkEditBox.Value = IIf(FolderBrowser1.EditBox, 1, 0)
        chkIncFiles.Value = IIf(FolderBrowser1.IncludeFiles, 1, 0)
        chkIncUrls.Value = IIf(FolderBrowser1.IncludeUrls, 1, 0)
        chkMNF.Value = IIf(FolderBrowser1.MakeNewFolderButton, 1, 0)
        
        txtTitle.Text = FolderBrowser1.Title
   
   End With

End Sub

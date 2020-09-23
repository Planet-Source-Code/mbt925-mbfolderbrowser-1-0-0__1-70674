VERSION 5.00
Begin VB.UserControl FolderBrowser 
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4245
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3180
   ScaleWidth      =   4245
   ToolboxBitmap   =   "FolderBrowser.ctx":0000
   Begin VB.Image imgBackGround 
      Height          =   810
      Left            =   0
      Picture         =   "FolderBrowser.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   900
   End
End
Attribute VB_Name = "FolderBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'=============================================
'API's

Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000
Private Const BIF_BROWSEFORPRINTER As Long = &H2000
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000
Private Const BIF_BROWSEINCLUDEURLS As Long = &H80
Private Const BIF_DONTGOBELOWDOMAIN As Long = &H2
Private Const BIF_EDITBOX As Long = &H10
Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Private Const BIF_RETURNFSANCESTORS As Long = &H8
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_SHAREABLE As Long = &H8000
Private Const BIF_STATUSTEXT As Long = &H4
Private Const BIF_USENEWUI As Long = &H40
Private Const BIF_VALIDATE As Long = &H20

Private Const MAX_PATH = 500

Private Type TBrowseInfo

  hwndOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
  
End Type

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As TBrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

'=======================================
'Default Property Values:

Const m_def_IncludeFiles = True
Const m_def_EditBox = True
Const m_def_DisplayName = ""
Const m_def_Title = "MBFolderBrowser"
Const m_def_FolderPath = ""
Const m_def_MakeNewFolderButton = True
Const m_def_IncludeUrls = False

'=======================================
'Property Variables:

Dim m_IncludeFiles As Boolean
Dim m_EditBox As Boolean
Dim m_MakeNewFolderButton As Boolean
Dim m_IncludeUrls As Boolean

Dim m_DisplayName As String
Dim m_Title As String
Dim m_FolderPath As String

'=============================================

Private Sub UserControl_Resize()

    UserControl.Height = imgBackGround.Height
    UserControl.Width = imgBackGround.Width

End Sub

Public Property Get IncludeFiles() As Boolean
Attribute IncludeFiles.VB_Description = "Return\\sets the folder browser include files or not."

    IncludeFiles = m_IncludeFiles
    
End Property

Public Property Let IncludeFiles(ByVal New_IncludeFiles As Boolean)

    m_IncludeFiles = New_IncludeFiles
    PropertyChanged "IncludeFiles"
    
End Property

Public Property Get EditBox() As Boolean
Attribute EditBox.VB_Description = "Return\\sets the visibility of Editbox button in folder browser."

    EditBox = m_EditBox
    
End Property

Public Property Let EditBox(ByVal New_EditBox As Boolean)

    m_EditBox = New_EditBox
    PropertyChanged "EditBox"
    
End Property

Public Property Get DisplayName() As String
Attribute DisplayName.VB_Description = "Return\\sets the DisplayName of item selected."

    DisplayName = m_DisplayName
    
End Property

Public Property Let DisplayName(ByVal New_DisplayName As String)

    m_DisplayName = New_DisplayName
    PropertyChanged "DisplayName"
    
End Property


Public Property Get Title() As String
Attribute Title.VB_Description = "Return\\sets the text displayed in folder browser's title bar."

    Title = m_Title
    
End Property

Public Property Let Title(ByVal New_Title As String)

    m_Title = New_Title
    PropertyChanged "Title"
    
End Property

Public Property Get FolderPath() As String
Attribute FolderPath.VB_Description = "Return the path of folder selected."

    FolderPath = m_FolderPath
    
End Property

Public Property Let FolderPath(ByVal New_FolderPath As String)

    m_FolderPath = New_FolderPath
    PropertyChanged "FolderPath"
    
End Property

Public Property Get MakeNewFolderButton() As Boolean
Attribute MakeNewFolderButton.VB_Description = "Return\\sets the visibility of MakeNewFolder button in folder browser."

    MakeNewFolderButton = m_MakeNewFolderButton
    
End Property

Public Property Let MakeNewFolderButton(ByVal New_MakeNewFolderButton As Boolean)

    m_MakeNewFolderButton = New_MakeNewFolderButton
    PropertyChanged "MakeNewFolderButton"
    
End Property

Public Property Get IncludeUrls() As Boolean
Attribute IncludeUrls.VB_Description = "Return\\sets the folder browser include urls or not."

    IncludeUrls = m_IncludeUrls
    
End Property

Public Property Let IncludeUrls(ByVal New_IncludeUrls As Boolean)

    m_IncludeUrls = New_IncludeUrls
    PropertyChanged "IncludeUrls"
    
End Property

Public Function BrowseFolder() As Boolean

    Dim Buffer As String
    Dim BrowseInfo As TBrowseInfo
    Dim IdList As Long
    
    With BrowseInfo
    
      .pidlRoot = 0
      .pszDisplayName = String(MAX_PATH, 0)
      .hwndOwner = UserControl.hWnd
      .lpszTitle = m_Title
      
      '-------------
      'Set flag
      
      If m_EditBox = True Then
          .ulFlags = .ulFlags Or BIF_EDITBOX
      End If
      If m_IncludeFiles = True Then
          .ulFlags = .ulFlags Or BIF_BROWSEINCLUDEFILES
      End If
      If m_IncludeUrls = True Then
          .ulFlags = .ulFlags Or BIF_BROWSEINCLUDEURLS
      End If
      If m_MakeNewFolderButton = True Then
          .ulFlags = .ulFlags Or BIF_NEWDIALOGSTYLE
      End If
      '-------------
      
    End With
    
    IdList = SHBrowseForFolder(BrowseInfo)
  
    If IdList Then
  
        Buffer = String$(MAX_PATH, 0)
        SHGetPathFromIDList IdList, Buffer
        CoTaskMemFree IdList
        
        BrowseFolder = True
    
    Else
     
        BrowseFolder = False
        Exit Function
     
    End If
  
    m_FolderPath = Buffer
    m_FolderPath = Left(m_FolderPath, InStr(1, m_FolderPath, Chr(0)) - 1)
    m_DisplayName = BrowseInfo.pszDisplayName

End Function

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    m_IncludeFiles = m_def_IncludeFiles
    m_EditBox = m_def_EditBox
    m_MakeNewFolderButton = m_def_MakeNewFolderButton
    m_IncludeUrls = m_def_IncludeUrls
    
    m_DisplayName = m_def_DisplayName
    m_Title = m_def_Title
    m_FolderPath = m_def_FolderPath
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_IncludeFiles = PropBag.ReadProperty("IncludeFiles", m_def_IncludeFiles)
    m_EditBox = PropBag.ReadProperty("EditBox", m_def_EditBox)
    m_MakeNewFolderButton = PropBag.ReadProperty("MakeNewFolderButton", m_def_MakeNewFolderButton)
    m_IncludeUrls = PropBag.ReadProperty("IncludeUrls", m_def_IncludeUrls)
    
    m_DisplayName = PropBag.ReadProperty("DisplayName", m_def_DisplayName)
    m_Title = PropBag.ReadProperty("Title", m_def_Title)
    m_FolderPath = PropBag.ReadProperty("FolderPath", m_def_FolderPath)
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("IncludeFiles", m_IncludeFiles, m_def_IncludeFiles)
    Call PropBag.WriteProperty("EditBox", m_EditBox, m_def_EditBox)
    Call PropBag.WriteProperty("MakeNewFolderButton", m_MakeNewFolderButton, m_def_MakeNewFolderButton)
    Call PropBag.WriteProperty("IncludeUrls", m_IncludeUrls, m_def_IncludeUrls)
    
    Call PropBag.WriteProperty("DisplayName", m_DisplayName, m_def_DisplayName)
    Call PropBag.WriteProperty("Title", m_Title, m_def_Title)
    Call PropBag.WriteProperty("FolderPath", m_FolderPath, m_def_FolderPath)
    
End Sub


Public Sub About()
Attribute About.VB_UserMemId = -552

    Dim FrmAbObj As frmAbout
    
    Set FrmAbObj = New frmAbout
    
    FrmAbObj.Show vbModal
     
    Unload FrmAbObj
    
    Set FrmAbObj = Nothing

End Sub

VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About MBFolderBrowser"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4605
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image2 
      Height          =   1110
      Left            =   3510
      Picture         =   "frmAbout.frx":3D8A
      Top             =   240
      Width           =   1110
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1386/10/15 - 2008/1/5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1320
      TabIndex        =   4
      Top             =   1200
      Width           =   1995
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mbt925@Gmail.Com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   1830
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mbt925"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MBFolderBrowser 1.0.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1320
      TabIndex        =   0
      Top             =   150
      Width           =   2130
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   60
      Picture         =   "frmAbout.frx":68ED
      Top             =   300
      Width           =   1050
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


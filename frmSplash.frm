VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4740
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Frame Frame1 
      Height          =   4650
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8280
      Begin VB.Image imgLogo 
         Height          =   3105
         Left            =   720
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   4815
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         TabIndex        =   1
         Top             =   3120
         Width           =   1275
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "RCard Systerm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   4665
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Hao Z"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5760
         TabIndex        =   2
         Top             =   2160
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Frame1_Click()
Unload Me
Load frmRcard
frmRcard.Show
End Sub

Private Sub imgLogo_Click()
Unload Me
Load frmRcard
frmRcard.Show
End Sub

Private Sub lblCompanyProduct_Click()
Unload Me
Load frmRcard
frmRcard.Show
End Sub

Private Sub lblProductName_Click()
Unload Me
Load frmRcard
frmRcard.Show
End Sub

Private Sub lblVersion_Click()
Unload Me
Load frmRcard
frmRcard.Show
End Sub

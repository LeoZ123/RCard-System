VERSION 5.00
Begin VB.Form frmRcard 
   Caption         =   "RCard"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmMan 
      Caption         =   "Managemer model"
      Height          =   3495
      Left            =   360
      TabIndex        =   8
      Top             =   2880
      Width           =   9135
      Begin VB.CommandButton cmdDel 
         Caption         =   "Lost card"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7320
         TabIndex        =   16
         Top             =   3000
         Width           =   1695
      End
      Begin VB.CommandButton cmdCreat 
         Caption         =   "Create New Account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5160
         TabIndex        =   14
         Top             =   2040
         Width           =   2295
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3720
         TabIndex        =   13
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtBal 
         Height          =   615
         Left            =   1320
         TabIndex        =   11
         Top             =   2040
         Width           =   3495
      End
      Begin VB.TextBox txtRec 
         Height          =   735
         Left            =   1440
         TabIndex        =   10
         Text            =   "0"
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblCar 
         Alignment       =   2  'Center
         Caption         =   "New card number"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblchar 
         Alignment       =   2  'Center
         Caption         =   "Recharge"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame frmUser 
      Caption         =   "User model"
      Height          =   2415
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin VB.CommandButton cmdShow 
         Caption         =   "show balance"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3720
         TabIndex        =   15
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton cmdCom 
         Caption         =   "Confirm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6000
         TabIndex        =   7
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtPu 
         Height          =   735
         Left            =   6600
         TabIndex        =   5
         Text            =   "0"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtBalance 
         Height          =   735
         Left            =   1920
         TabIndex        =   4
         Text            =   "0"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtCard 
         Height          =   735
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblCost 
         Alignment       =   2  'Center
         Caption         =   "Cost"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblBalance 
         Alignment       =   2  'Center
         Caption         =   "Balance"
         Height          =   615
         Left            =   600
         TabIndex        =   3
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblCard 
         Alignment       =   2  'Center
         Caption         =   "Card Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmRcard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As Integer
Dim strCard(1 To 1000) As String
Dim dblVal(1 To 1000) As Double
Dim flag As Integer
Dim dtToday As Date
Dim strDate As String

Private Sub cmdCom_Click()
Dim a As Integer
dtToday = Date
For a = 1 To c
'cross the money
If txtCard.Text = strCard(a) Then
    If dblVal(a) > txtPu Then
        If txtPu > 0 Then
    dblVal(a) = dblVal(a) - txtPu.Text
    txtBalance = dblVal(a)
    
    strDate = dtToday
    pk = FreeFile
    Open (App.Path & "\" & txtCard.Text & ".txt") For Append As pk
    Write #pk, strDate, txtPu.Text, txtRec.Text, dblVal(a)
    Close #pk
        End If
    Else: MsgBox ("cost greaster than balance, please load money")
    End If
    
    If dblVal(a) < 10 Then
    MsgBox ("balance less than $10,please load money")
    End If
End If
Next a
    If txtPu.Text < 0 Then
    MsgBox ("Error")
    End If
    
Call send
txtPu = "0"
End Sub


Private Sub cmdCreat_Click()
Dim a As Integer
c = c + 1
strCard(c) = txtBal
dblVal(c) = 0

Call send

For a = 1 To c
If txtCard.Text = strCard(a) Then
dblVal(a) = dblVal(a) + txtRec.Text
txtBalance = dblVal(a)

    strDate = dtToday
    pk = FreeFile
    Open (App.Path & "\" & txtCard.Text & ".txt") For Append As pk
    Write #pk, strDate, txtPu.Text, txtRec.Text, dblVal(a)
    Close #pk
    
End If
Next a

Call send

MsgBox ("Your account have been created!!")
End Sub

Sub send()
Dim intFileNum As Integer
intFileNum = FreeFile
Open (App.Path & "/cards.txt") For Output As #intFileNum
Dim a As Integer

For a = 1 To c
Print #intFileNum, strCard(a), ",", dblVal(a)
Next a

Close #intFileNum
End Sub
Sub tra()
 strDate = dtToday
    pk = FreeFile
    Open (App.Path & "\" & txtCard.Text & ".txt") For Append As pk
    Write #pk, strDate, txtPu.Text, txtRec.Text, dblVal(a)
    Close #pk
End Sub

Private Sub cmdDel_Click()
Dim strDelet As String
strDelet = txtCard
Dim intFileNum As Integer
intFileNum = FreeFile
Open (App.Path & "/cards.txt") For Output As #intFileNum
Dim a As Integer
For a = 1 To c
If (strDelet <> strCard(a)) Then
Print #intFileNum, strCard(a), ",", dblVal(a)
End If
Next a
Close #intFileNum
txtCard = ""
txtBalance = ""

 MsgBox (" Your account has been removed")
End Sub

Private Sub cmdLoad_Click()
Dim a As Integer
For a = 1 To c
If txtCard.Text = strCard(a) Then
    
    If txtRec >= 0 Then
    dblVal(a) = dblVal(a) + txtRec.Text
    txtBalance = dblVal(a)

    strDate = dtToday
    pk = FreeFile
    Open (App.Path & "\" & txtCard.Text & ".txt") For Append As pk
    Write #pk, strDate, txtPu.Text, txtRec.Text, dblVal(a)
    Close #pk
    End If

End If
Next a

If txtRec < 0 Then
MsgBox ("Error")
End If

Call send

txtRec = "0"

End Sub

Private Sub cmdShow_Click()
Dim a As Integer
flag = 0
'set flag=1 when to a match
'match the account number and acconut balance
For a = 1 To c
If txtCard.Text = strCard(a) Then
    txtBalance = dblVal(a)
    flog = 1
End If
Next a

'search
If flog = 0 Then
    txtBal = txtCard
    MsgBox ("weclome new user,please creat a new account")
End If
End Sub




Private Sub Form_Load()
'load information from file
c = 0
Dim intFileNum As Integer
intFileNum = FreeFile
Open (App.Path & "/cards.txt") For Input As #intFileNum
Do While Not EOF(intFileNum)
c = c + 1
Input #intFileNum, strCard(c), dblVal(c)
Loop
Close #intFileNum
End Sub





Private Sub txtCard_Click()
txtCard = ""
End Sub

Private Sub txtPu_Click()
txtPu = ""
End Sub

Private Sub txtRec_Click()
txtRec = ""
End Sub

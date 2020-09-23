VERSION 5.00
Object = "*\A..\wMX.vbp"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MX Record lookup"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin wMX.MX MX1 
      Left            =   7920
      Top             =   120
      _ExtentX        =   714
      _ExtentY        =   450
   End
   Begin VB.ListBox lstDNS 
      Height          =   2400
      Left            =   7200
      TabIndex        =   5
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.ListBox lstMX 
      Height          =   2400
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   6975
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtDomain 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Text            =   "mail.com"
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblMX 
      Caption         =   "Domain name for MX query:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    lstMX.Clear
    lstDNS.Clear
End Sub

Private Sub cmdGo_Click()
    Dim iDNS As Integer
    Dim iMX As Integer
    Dim sMX As String
    
    MX1.Domain = txtDomain.Text
    sMX = MX1.GetMX
    lstMX.AddItem "BEST: " & sMX
    For iDNS = 0 To MX1.DNSCount
        lstDNS.AddItem MX1.DNS(iDNS)
    Next iDNS
    
    For iMX = 0 To MX1.MXCount
        lstMX.AddItem MX1.MX(iMX) & Chr(9) & "PREF: " & MX1.Pref(iMX)
    Next iMX

End Sub


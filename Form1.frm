VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Livestock Update"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   360
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Transfer"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "~"
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Location on FTP"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Source file"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Server"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "User"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Inet1.Protocol = icFTP
    Inet1.URL = Text1.Text
    Inet1.UserName = Text2.Text
    Inet1.Password = Text3.Text
    Inet1.Execute , "PUT " & Text4.Text & " " & Text5.Text
    'Inet1.Execute , "GET ftpdir/filemane c:\dir\filename"
    
    MsgBox "Transfer Cpmplete"
End Sub

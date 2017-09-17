VERSION 5.00
Begin VB.Form Rectangle 
   BackColor       =   &H0000C0C0&
   Caption         =   "Rectangle Area Or Perimeter Cal"
   ClientHeight    =   8685
   ClientLeft      =   615
   ClientTop       =   945
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   6045
   Begin VB.CommandButton Command6 
      Caption         =   "<< Back"
      BeginProperty Font 
         Name            =   "MuRdOiNk"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   14
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MuRdOiNk"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   13
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MuRdOiNk"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   12
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear Texts"
      BeginProperty Font 
         Name            =   "MuRdOiNk"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   11
      Top             =   5760
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "MuRdOiNk"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   10
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1115
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1115
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "MuRdOiNk"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Top             =   4920
      Width           =   2415
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Perimeter Of The Rectangle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FF00&
      Height          =   1695
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   4815
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Area Of The Rectangle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   480
         Width           =   3135
      End
   End
   Begin VB.Image Image1 
      Height          =   10005
      Left            =   0
      Picture         =   "Rectangel Area Perimeter Cal.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   19995
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Width Of The Rectangle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Length Of The Rectangle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1115
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "Area Of The Rectangle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   3840
      Width           =   3135
   End
End
Attribute VB_Name = "Rectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
num = Val(Text1.Text) * Val(Text2.Text)
num = Text1.Text * Text2.Text
Label2.Caption = num
End Sub

Private Sub Command2_Click()
nun = Val(Text1.Text) * 2 + Val(Text2.Text) * 2
nun = Text1.Text * 2 + Text2.Text * 2
Label2.Caption = nun
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Label2.Caption = ""
End Sub

Private Sub Command4_Click()
a = MsgBox("Are You Want To Exit Programme ?", vbYesNo + vbCritical, "Exit")
If a = 6 Then
  End
End If
End Sub

Private Sub Command5_Click()
Rectangle.Caption = "About"
Rectangle.Height = 10000
Rectangle.Width = 20000
Image1.Visible = True
Frame1.Visible = False
Option1.Visible = False
Option2.Visible = False
Text1.Visible = False
Text2.Visible = False
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
End Sub

Private Sub Command6_Click()
Image1.Visible = False
Rectangle.Width = 6285
Rectangle.Height = 9270
Frame1.Visible = True
Option1.Visible = True
Option2.Visible = True
Text1.Visible = True
Text2.Visible = True
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Rectangle.Caption = "Rectangle Area Or Perimeter Cal"
End Sub

Private Sub Option1_Click()
Option2.Value = False
Command2.Visible = False
Command1.Visible = True
End Sub

Private Sub Option2_Click()
Option1.Value = False
Command1.Visible = False
Command2.Visible = True
Text1.Text = ""
Text2.Text = ""
Label2.Caption = ""
Label1.Caption = "Perimeter Of The Rectangle"
End Sub


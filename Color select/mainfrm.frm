VERSION 5.00
Begin VB.Form main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color Select"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4440
   Icon            =   "mainfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frm 
      Caption         =   "Values"
      Height          =   1695
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   4455
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   240
         Top             =   720
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Color code"
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Hex Value"
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "VB Code"
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Values"
      Height          =   1095
      Left            =   3120
      Picture         =   "mainfrm.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.PictureBox colshade 
      BackColor       =   &H0000DCFF&
      Height          =   1095
      Left            =   600
      MousePointer    =   2  'Cross
      Picture         =   "mainfrm.frx":040C
      ScaleHeight     =   1035
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   480
      Width           =   2415
      Begin VB.PictureBox selected 
         BackColor       =   &H00E0E0E0&
         Height          =   1095
         Left            =   240
         ScaleHeight     =   1035
         ScaleWidth      =   2115
         TabIndex        =   1
         Top             =   0
         Width           =   2175
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Selected Color"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   480
            Width           =   2415
         End
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Move  mouse over the color you want."
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3345
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************
            'Jim Jose
            'Email jimjosev33@yahoo.com
'****************************************

'You are free to use this code  in your  program without any  permission.
'If you found this useful,never forget to @ RATE @ my code
'****************************************
Dim color

Private Sub colshade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'One Line to get color
    color = colshade.Point(X, Y) 'Get color

    selected.BackColor = color 'set color
    
End Sub



Private Sub Command1_Click()
Timer1.Enabled = True
End Sub


Private Sub Timer1_Timer()
        Text1.Text = color
        Text2.Text = Hex(color)
        If Len(Text2.Text) < 6 Then
            Text3.Text = "&H" & String(6 - Len(Text2.Text), "0") & Text2.Text
        Else
            Text3.Text = "&H" & Text2.Text
        End If
End Sub

VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404000&
   BorderStyle     =   0  'None
   Caption         =   "-"
   ClientHeight    =   3765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton C 
      BackColor       =   &H00C0C000&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   13
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton C 
      BackColor       =   &H00C0C000&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   12
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton C 
      BackColor       =   &H00C0C000&
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   14
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton C 
      BackColor       =   &H00C0C000&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   16
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton C 
      BackColor       =   &H00C0C000&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   17
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton C 
      BackColor       =   &H00C0C000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton C 
      BackColor       =   &H00C0C000&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton C 
      BackColor       =   &H00C0C000&
      Caption         =   "BACKSPACE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   15
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton C 
      BackColor       =   &H00C0C000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton C 
      BackColor       =   &H00C0C000&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton C 
      BackColor       =   &H00C0C000&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton C 
      BackColor       =   &H00C0C000&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton C 
      BackColor       =   &H00C0C000&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton C 
      BackColor       =   &H00C0C000&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton C 
      BackColor       =   &H00C0C000&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton C 
      BackColor       =   &H00C0C000&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton C 
      BackColor       =   &H00C0C000&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton C 
      BackColor       =   &H00C0C000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtDisplay 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   615
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   960
      Top             =   2160
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- CALCULATOR -"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   465
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   555
      Left            =   3360
      TabIndex        =   19
      Top             =   0
      Width           =   345
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   -120
      Top             =   0
      Width           =   4095
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   2760
      Shape           =   2  'Oval
      Top             =   600
      Width           =   855
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   120
      Shape           =   2  'Oval
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=================================================
'=============CALCULATOR MADE EASY================
'=================================================
'For comments and suggestions email me at mark_anthony_dinglasa@yahoo.com
'if you think this code helps you then vote for it !
'Thank's and Enjoy Coding !

Dim Symbol As String 'Use variable to determine what button symbol is click
Private Sub C_Click(Index As Integer)
Dim Ctr As Integer 'Use for Control array
Static I As Currency  'Use for making Currency remains in memory if not form is unloaded

On Error GoTo Errhandler 'When error occours it be trap and this sub will exit

Select Case Index
    Case 0 To 9, 17  'identify which command button array is clicked
        txtDisplay.Text = txtDisplay.Text & C(Index).Caption  'display command button's caption to Textbox txtDisplay.Text
    Case 10 'ADD button
        I = I + CCur(txtDisplay.Text)
        Symbol = "+"
        txtDisplay.Text = ""
    Case 11 'Minus button
        I = I + CCur(txtDisplay.Text)
        Symbol = "-"
        txtDisplay.Text = ""
    Case 12 'Divide button
        I = I + CCur(txtDisplay.Text)
        Symbol = "/"
        txtDisplay.Text = ""
    Case 13 'Multiply button
        I = I + CCur(txtDisplay.Text)
        Symbol = "x"
        txtDisplay.Text = ""
    Case 14       'When Clear button is clicked then it clears the Textbox
        txtDisplay.Text = ""
    Case 15        'When Backspace button is clicked then it will get one digit into the right side of the value in txtdisplay.text
        txtDisplay.Text = Mid(txtDisplay.Text, 1, Len(txtDisplay.Text) - 1)
    Case 16 ' Equals button
        If Symbol = "+" Then
            txtDisplay.Text = I + CCur(txtDisplay.Text) 'Display the Sum to txtDisplay.Text
            I = 0
        ElseIf Symbol = "-" Then
            txtDisplay.Text = I - CCur(txtDisplay.Text)  'Display the Difference to txtDisplay.Text
            I = 0
        ElseIf Symbol = "/" Then
            txtDisplay.Text = I / CCur(txtDisplay.Text)  'Display the Quotient to txtDisplay.Text
            I = 0
         ElseIf Symbol = "x" Then
            txtDisplay.Text = I * CCur(txtDisplay.Text)  'Display the Product to txtDisplay.Text
            I = 0
        End If
End Select

Errhandler:
Exit Sub
End Sub



Private Sub C_LostFocus(Index As Integer) 'When command button is not focus then the back color will be back to Nomal
Select Case Index
    Case 0 To 17
        C(Index).BackColor = &HC0C000
End Select
End Sub

Private Sub Label2_Click()
End 'Exit
End Sub

Private Sub Timer1_Timer()
If TypeOf Me.ActiveControl Is CommandButton Then 'Lights effect in command button
    Set LightsMe = Me.ActiveControl
    LightsMe.BackColor = &HFFFFC0
End If
End Sub

Private Sub txtDisplay_Change()
If Len(txtDisplay.Text) > 10 Then 'Only 10 digits will be accepted in the textbox txtdisplay.text
    txtDisplay.Text = Mid(txtDisplay.Text, 1, Len(txtDisplay.Text) - 1) 'Maintaining 10 digits into the textbox txtdisplay.text
Else
    Exit Sub
End If
End Sub



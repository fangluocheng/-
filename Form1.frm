VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "计算器"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdeq 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   19
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton cmdchu 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   18
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton cmdcheng 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   17
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmdjian 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   16
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton num1 
      Caption         =   "6"
      Height          =   735
      Index           =   6
      Left            =   1800
      TabIndex        =   15
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton num1 
      Caption         =   "3"
      Height          =   735
      Index           =   3
      Left            =   1800
      TabIndex        =   14
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   13
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton num1 
      Caption         =   "2"
      Height          =   735
      Index           =   2
      Left            =   1080
      TabIndex        =   12
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton num1 
      Caption         =   "5"
      Height          =   735
      Index           =   5
      Left            =   1080
      TabIndex        =   11
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton num1 
      Caption         =   "0"
      Height          =   735
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton num1 
      Caption         =   "1"
      Height          =   735
      Index           =   1
      Left            =   360
      TabIndex        =   9
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton num1 
      Caption         =   "4"
      Height          =   735
      Index           =   4
      Left            =   360
      TabIndex        =   8
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton cmdjia 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   7
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton num1 
      Caption         =   "9"
      Height          =   735
      Index           =   9
      Left            =   1800
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton num1 
      Caption         =   "8"
      Height          =   735
      Index           =   8
      Left            =   1080
      TabIndex        =   5
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton num1 
      Caption         =   "7"
      Height          =   735
      Index           =   7
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出"
      Height          =   735
      Left            =   3240
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "取消"
      Height          =   735
      Left            =   3240
      TabIndex        =   1
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmdDEL 
      Caption         =   "清除"
      Height          =   735
      Left            =   3240
      TabIndex        =   0
      Top             =   2280
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub Run1()
Dim x() As String
Dim c As Integer
Dim q As Double
c = Len(Text1.Text)
Text1.Text = Left(Text1.Text, c - 1)
x = Split(Text1.Text, "+")
q = Val(x(0)) + Val(x(1))
Text1.Text = Text1.Text & "=" & q
End Sub

Sub Run2()
Dim x() As String
Dim c As Integer
Dim q As Double
c = Len(Text1.Text)
Text1.Text = Left(Text1.Text, c - 1)
x = Split(Text1.Text, "-")
q = Val(x(0)) - Val(x(1))
Text1.Text = Text1.Text & "=" & q
End Sub

Sub Run3()
Dim x() As String
Dim c As Integer
Dim q As Double
c = Len(Text1.Text)
Text1.Text = Left(Text1.Text, c - 1)
x = Split(Text1.Text, "*")
q = Val(x(0)) * Val(x(1))
Text1.Text = Text1.Text & "=" & q
End Sub

Sub Run4()
Dim x() As String
Dim c As Integer
Dim q As Double
c = Len(Text1.Text)
Text1.Text = Left(Text1.Text, c - 1)
x = Split(Text1.Text, "/")
If Val(x(1)) = 0 Then
Text1.Text = "分母不能为0"
Else
q = Val(x(0)) / Val(x(1))
Text1.Text = Text1.Text & "=" & q
End If
End Sub

Private Sub cmdBack_Click()
If Text1.Text = "" Then
Text1.Text = ""
Else
Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
End If
End Sub

Private Sub cmdDEL_Click()
Text1.Text = ""
End Sub

Private Sub cmdeq_Click()
    If InStr(1, Text1.Text, "=") <> 0 Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "+" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "-" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "*" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "/" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "=" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "." Then
    Text1.Text = Text1.Text & ""
    Else
    Text1.Text = Text1.Text & "="
        If InStr(1, Text1.Text, "+") <> 0 Then
        Call Run1
        End If
        If InStr(1, Text1.Text, "-") <> 0 Then
        Call Run2
        End If
        If InStr(1, Text1.Text, "*") <> 0 Then
        Call Run3
        End If
        If InStr(1, Text1.Text, "/") <> 0 Then
        Call Run4
        End If
    End If
End Sub

Private Sub cmdfu_Click()
If InStr(1, Text1.Text, "-") <> 0 Then
Text1.Text = Right(Text1.Text, Len(Text1.Text) - 1)
Else
Text1.Text = "-" & Text1.Text
End If
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdjia_Click()
    If Right(Text1.Text, 1) = "" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "+" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "-" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "*" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "/" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "=" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "." Then
    Text1.Text = Text1.Text & ""
    Else
    Text1.Text = Text1.Text & "+"
    End If
End Sub

Private Sub cmdjian_Click()
    If Right(Text1.Text, 1) = "" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "+" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "-" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "*" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "/" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "=" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "." Then
    Text1.Text = Text1.Text & ""
    Else
    Text1.Text = Text1.Text & "-"
    End If
End Sub

Private Sub cmdcheng_Click()
    If Right(Text1.Text, 1) = "" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "+" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "-" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "*" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "/" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "=" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "." Then
    Text1.Text = Text1.Text & ""
    Else
    Text1.Text = Text1.Text & "*"
    End If
End Sub

Private Sub cmdchu_Click()
    If Right(Text1.Text, 1) = "" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "+" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "-" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "*" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "/" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "=" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "." Then
    Text1.Text = Text1.Text & ""
    Else
    Text1.Text = Text1.Text & "/"
    End If
End Sub

Private Sub Command10_Click()
    If Right(Text1.Text, 1) = "" Then
    Text1.Text = "0."
    ElseIf Right(Text1.Text, 1) = "+" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "-" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "*" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "/" Then
    Text1.Text = Text1.Text & ""
    ElseIf Right(Text1.Text, 1) = "=" Then
    Text1.Text = Text1.Text & ""
    ElseIf InStr(1, Text1.Text, ".") <> 0 Then
    Text1.Text = Text1.Text & ""
    Else
    Text1.Text = Text1.Text & "."
    End If
End Sub

Private Sub num0_Click()
    If Text1.Text = "" Then
        Text1.Text = "0"
    Else
        Text1.Text = Text1.Text & "0"
    End If
End Sub

Private Sub num1_Click(Index As Integer)
    If Text1.Text = "" Then
        'MsgBox "0"
        Text1.Text = CStr(Index)
    ElseIf Trim(Text1.Text) = "0" Then
        'MsgBox "1"
        Text1.Text = Text1.Text & ""
    Else
        'MsgBox "2"
        Text1.Text = Text1.Text + CStr(Index)
    End If
End Sub


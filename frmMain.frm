VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ANNA Reg Key Demo"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate 500 Keys"
      Height          =   375
      Index           =   7
      Left            =   4320
      TabIndex        =   21
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate 100 Keys"
      Height          =   375
      Index           =   6
      Left            =   4320
      TabIndex        =   19
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate 50 Keys"
      Height          =   375
      Index           =   5
      Left            =   4320
      TabIndex        =   18
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   6255
      Begin VB.CommandButton cmdValidate 
         Caption         =   "Validate Key"
         Height          =   375
         Left            =   4440
         TabIndex        =   16
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtKey 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   3480
         MaxLength       =   5
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtKey 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtKey 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtKey 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   960
         MaxLength       =   5
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtKey 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   5
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Registration Key:"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "-"
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
         Left            =   3360
         TabIndex        =   15
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "-"
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
         Left            =   2520
         TabIndex        =   14
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "-"
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
         Left            =   1680
         TabIndex        =   13
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "-"
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
         Left            =   840
         TabIndex        =   12
         Top             =   480
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate 25 Keys"
      Height          =   375
      Index           =   4
      Left            =   4320
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate 20 Keys"
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   4
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate 15 Keys"
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate 10 Keys"
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate 5 Keys"
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.ListBox lstKeys 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblKeys 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3120
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGenerate_Click(Index As Integer)
'Variables
    Dim i As Integer
    Dim sngStartTime As Single
    
'Clear Key List
    lstKeys.Clear
       
'Remember Start Time
    sngStartTime = Timer
    DoEvents
    
'Generate Keys
    If Index < 5 Then
        For i = 1 To ((Index + 1) * 5)
            lstKeys.AddItem GenerateKey
        Next
    Else
        If Index = 5 Then
            For i = 1 To 50
                lstKeys.AddItem GenerateKey
            Next
        Else
            If Index = 6 Then
                For i = 1 To 100
                    lstKeys.AddItem GenerateKey
                Next
            Else
                For i = 1 To 500
                    lstKeys.AddItem GenerateKey
                Next
            End If
        End If
    End If

'Calc Generation Time
    sngStartTime = Timer - sngStartTime
    If sngStartTime < 0 Then
        sngStartTime = sngStartTime * -1
    End If

'Show Generation Time
    lblKeys.Caption = (i - 1) & " Keys Generated in " & sngStartTime & " Seconds"

End Sub

Private Sub cmdValidate_Click()
'Validate Key
    If ValidateKey(txtKey(0).Text, txtKey(1).Text, txtKey(2).Text, txtKey(3).Text, txtKey(4).Text) = True Then
        MsgBox "Valid Registration Key", vbInformation, " "
    Else
        MsgBox "Invalid Registration Key", vbCritical, " "
    End If
    
End Sub

Private Sub lstKeys_Click()
'Variables
    Dim i As Integer
    Dim strTemp() As String

'Make sure there's a key
    If lstKeys.ListCount = 0 Then
        Exit Sub
    End If

'SPlit
    strTemp = Split(lstKeys.Text, "-")
    
'Spit to Text Boxes
    For i = 0 To UBound(strTemp)
        txtKey(i) = strTemp(i)
    Next

End Sub

Private Sub txtKey_GotFocus(Index As Integer)
'Highlight
    txtKey(Index).SelStart = 0
    txtKey(Index).SelLength = Len(txtKey(Index).Text)
    
End Sub

VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Random"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtsymbol 
      Height          =   495
      Left            =   4920
      TabIndex        =   8
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CheckBox chkSymbol 
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   3240
      Width           =   855
   End
   Begin VB.CheckBox chkNumber 
      Caption         =   "0-9"
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CheckBox chkLower 
      Caption         =   "a-z"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CheckBox chkUpper 
      Caption         =   "A-Z"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ComboBox cmbNumberOfChar 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enter"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2595
      ItemData        =   "Form1.frx":0000
      Left            =   3360
      List            =   "Form1.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number Of Characters"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Result() As Long

Dim IfSuccess As Boolean
Dim randomnumber As Integer

Dim strRandom As String
Dim TypeOfCondition As Byte
Private Sub Command1_Click()
 ' IfSuccess = Random_X(Text1.Text, Text2.Text, Text3.Text, Result, Check1.Value)
  'List1.Clear
    
     '   For i = LBound(Result) To UBound(Result)
          '  List1.AddItem Result(i)
      '  Next i
        

For i = 1 To cmbNumberOfChar.Text
strRandom = strRandom + Generator
Next
  List1.AddItem strRandom
  strRandom = ""
  


End Sub

Private Sub Conditions()

End Sub
Private Function Generator()
If chkUpper And chkLower And chkNumber And chkSymbol Then
    IfSuccess = Random_X(1, 0, 4, Result, False)
        Select Case Result(0)
        Case 1
        Call Uppercase
        Case 2
         Call Lowercase
        Case 3
        Call Number
        Case 4
        Call Symbol
        End Select
        
        
        ElseIf chkLower And chkNumber And chkSymbol Then
    IfSuccess = Random_X(1, 0, 3, Result, False)
        Select Case Result(0)
        Case 1
         Call Lowercase
        Case 2
        Call Number
        Case 3
         Call Symbol
        End Select
        
        
ElseIf chkUpper And chkLower And chkSymbol Then
    IfSuccess = Random_X(1, 0, 3, Result, False)

        Select Case Result(0)
        Case 1
         Call Uppercase
        Case 2
        Call Lowercase
        Case 3
         Call Symbol
        End Select
        
        ElseIf chkUpper And chkNumber And chkSymbol Then
    IfSuccess = Random_X(1, 0, 3, Result, False)

        Select Case Result(0)
        Case 1
         Call Uppercase
        Case 2
        Call Number
        Case 3
         Call Symbol
        End Select

ElseIf chkUpper And chkLower And chkNumber Then
    IfSuccess = Random_X(1, 0, 3, Result, False)

        Select Case Result(0)
        Case 1
         Call Uppercase
        Case 2
        Call Lowercase
        Case 3
         Call Number
        End Select


    
        ElseIf chkNumber And chkSymbol Then
    IfSuccess = Random_X(1, 0, 2, Result, False)
        Select Case Result(0)
        Case 1
       Call Number
        Case 2
      Call Symbol

        End Select
    ElseIf chkLower And chkSymbol Then
    IfSuccess = Random_X(1, 0, 2, Result, False)
        Select Case Result(0)
        Case 1
        Call Lowercase
        Case 2
        Call Symbol

        End Select
        
ElseIf chkUpper And chkLower Then
    IfSuccess = Random_X(1, 0, 2, Result, False)
        Select Case Result(0)
        Case 1
        Call Uppercase
        Case 2
        Call Lowercase

        End Select
        
        ElseIf chkUpper And chkSymbol Then
    IfSuccess = Random_X(1, 0, 2, Result, False)
        Select Case Result(0)
        Case 1
        Call Uppercase
        Case 2
        Call Symbol

        End Select


        ElseIf chkUpper And chkNumber Then
    IfSuccess = Random_X(1, 0, 2, Result, False)
        Select Case Result(0)
        Case 1
        Call Uppercase
        Case 2
        Call Number

        End Select
        
        
            ElseIf chkLower And chkNumber Then
    IfSuccess = Random_X(1, 0, 2, Result, False)
        Select Case Result(0)
        Case 1
        Call Lowercase
        Case 2
        Call Number

        End Select
        
        
        ElseIf chkUpper Then
            Call Uppercase
                    ElseIf chkLower Then
            Call Lowercase
                    ElseIf chkNumber Then
            Call Number
                    ElseIf chkSymbol Then
            Call Symbol

       
        

End If




Generator = Chr(Result(0))
End Function




Private Sub Form_Load()
Call combo_ready
txtsymbol = "!@#$%^&*"
End Sub
Private Sub combo_ready()
For i = 1 To 20
cmbNumberOfChar.AddItem (i)
Next
cmbNumberOfChar.ListIndex = 7
End Sub
Private Sub Uppercase()
 IfSuccess = Random_X(1, 64, 90, Result, False)
End Sub
Private Sub Lowercase()
  IfSuccess = Random_X(1, 96, 122, Result, False)
End Sub
Private Sub Symbol()
Dim CountOfSymbol(255) As Byte

For i = 1 To Len(Trim(txtsymbol))

CountOfSymbol(i) = Asc(Mid(Trim(txtsymbol), i, 1))
 
Next
 IfSuccess = Random_X(1, 0, Len(Trim(txtsymbol)), Result, False)

Print Result(0)

Result(0) = CountOfSymbol(Result(0))



End Sub
Private Sub Number()
 IfSuccess = Random_X(1, 0, 10, Result, False)
Result(0) = Result(0) + 47
End Sub


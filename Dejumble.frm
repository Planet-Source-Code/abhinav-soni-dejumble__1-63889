VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dejumble"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1620
      ItemData        =   "Dejumble.frx":0000
      Left            =   240
      List            =   "Dejumble.frx":0002
      TabIndex        =   3
      Top             =   1080
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&De-Jumble"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Result:"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Enter jumbled word:"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   1395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declare statements
Private Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Private Declare Function CharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As String
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long


Dim Exlobj As Object ' Excel object
Dim lText As Byte ' Length of the word
Dim WrdCnt As Long 'Current word count
Dim TtlCnt As Long ' Total number of words that can be formed by jumbling

Private Sub Command1_Click()
Dim i As Byte ' Loop variable
Dim flag As Boolean ' flag=true:word acceptable
flag = True

' Get lenght of the word
lText = Len(Text1)

' Check the word
For i = 1 To lText
    If IsCharAlpha(Asc(Mid$(Text1, i, 1))) = 0 Or Asc(Mid$(Text1, i, 1)) > 122 Then flag = False: Exit For
Next i

If flag And lText <= 10 Then 'if flag=true or lenght of word >=10

    List1.Clear 'Empty list1 listbox
    Text1 = CharLower(Text1) 'Convert to lower character
    Arrange ' arrange words in alphabetical order
    Calculate Text1 'calculate total words possible
    Jumble Text1, lText 'Jumble up words and check correct combinations
    
    'If no word is possible
    If List1.ListCount = 0 Then MsgBox "No word found.", vbCritical
    Caption = "De-Jumble"
    Label3 = CStr(WrdCnt) + " combinations checked. " + CStr(List1.ListCount) + " are correct."
Else

    MsgBox "Word unacceptable."
    
End If
End Sub

Private Sub Form_Load()
' Create excel object
Set Exlobj = CreateObject("excel.application")

' On error display error and quit
If Err Then

    MsgBox Error & "."
    End ' Exit the program
    
End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
' Exit excel object
Exlobj.quit
End Sub

Private Sub Text1_Change()
' If length of the word > 2 then enable de-jumbling
If Len(Text1.Text) > 2 Then Command1.Enabled = True Else Command1.Enabled = False
End Sub

Private Function Factorial&(ByVal number As Integer)
' This function calculates factorial of a number
Dim i As Integer ' Loop variable
Dim result As Long
result = 1
For i = 1 To number
result = result * i
Next i
Factorial = result
End Function

Private Sub Calculate(ByVal s As String)
' Alp consists of number of times an alphabet
' appears in the given word
' alp(0) = a
' alp(1) = b
' alp(25) = z
Dim alp(25) As Integer
Dim i As Integer ' Loop variable
Dim result As Long

result = Factorial(lText)

For i = 1 To lText
alp(Asc(Mid$(s, i, 1)) - 97) = alp(Asc(Mid$(s, i, 1)) - 97) + 1
Next i

For i = 0 To 25
result = result / Factorial(alp(i))
Next i


TtlCnt = result ' Now result gives total number of possible combinations
WrdCnt = 0 ' Words checked
End Sub

Private Sub Arrange()
' This procedure arranges the given word in alphabetical order
' No need actually but lets be more professional

' Here this zero based array alp coressponds
' to a character (or letter) in the given word
ReDim alp(0 To lText - 1) As Byte
Dim i As Byte, j As Byte ' Variables used in looping
Dim tmp As Byte ' Temp variable

' Assigning to alp array
For i = 1 To lText
alp(i - 1) = Asc(Mid$(Text1, i, 1))
Next i

' Bubble sort
For i = 0 To lText - 2
For j = 0 To lText - 2 - i
If alp(j) > alp(j + 1) Then
tmp = alp(j)
alp(j) = alp(j + 1)
alp(j + 1) = tmp
End If
Next j
Next i

Text1 = ""

' Again from alp array to text box(Text1)
For i = 0 To lText - 1
Text1 = Text1 + Chr$(alp(i))
Next i
Erase alp
End Sub


Private Sub Jumble(ByVal s$, ln As Byte)
' This function takes two parametes
' First is string to (re)jumble
' Second is length of the string to evaluate
' Actually second parameter gives us the idea of the
' depth in recursion
' When lengtt of the passed sting becomes 1 we check for spelling

Dim i As Byte, j As Byte ' Loop variables
Dim buff As String
Dim buff1 As String
Dim flag  As Boolean

' Number of times jumble is called recursively
' (Not the depth)
Dim r As Byte

' This dynamic array alp keeps record of the
' characters passed to recurion so that no
' repeated word is evaluated again
Dim alp() As String * 1

If ln = 1 Then 'Length is 1

    WrdCnt = WrdCnt + 1 ' Increase word count by one
    
    'Window caption is the current word count
    Caption = s + " - " + CStr(WrdCnt) + " of " + CStr(TtlCnt)
    
    'If spelling is correct add this word to list1 list box
    If Exlobj.checkspelling(s) Then List1.AddItem s
    
Else

' Store the string which has to go to
' next recursion in another buffer
buff = Right$(s, ln)

' Redefining size of alp
ReDim alp(1 To ln)
For i = 1 To ln
flag = True
    For j = 1 To r 'only check upto r
    If alp(j) = Mid$(buff, i, 1) Then ' If repeatition is found
        ' Do not go to next recursion
        flag = False
        
        ' Exit for as flag is already false
        ' No need to check further
        Exit For
    End If
    Next j
    
If flag Then ' If repeatition is not there
buff1 = String(ln - 1, Chr$(0)) ' Initialize buff2 to required length
r = r + 1 ' Keeping track of number for call of Jumble
alp(r) = Mid$(buff, i, 1)
lstrcat buff1, Left$(buff, i - 1)
lstrcat buff1, Right$(buff, ln - i)

' Again call jumble recursively
Call Jumble(Left$(s, lText - ln) + alp(r) + buff1, ln - 1)
End If
Next i

' Erase memory allocated by alp
Erase alp
End If
End Sub



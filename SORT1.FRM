VERSION 5.00
Begin VB.Form Sorting 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   1140
   ClientTop       =   1800
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3315
   ScaleWidth      =   7890
   Begin VB.CommandButton Command1 
      Caption         =   "UnSorted"
      Height          =   495
      Index           =   2
      Left            =   2520
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bubble Sorted"
      Height          =   495
      Index           =   1
      Left            =   3840
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Shell Sorted"
      Height          =   495
      Index           =   0
      Left            =   5160
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSub 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Sorting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This Program opens a text file reads in the stats compiles wins
'and then sorts the data in the array according to number of wins
'using either the bubble sort or the shell Sort method
Dim Team1(1 To 6), Team2(1 To 7, 1 To 7), wins(), CNT As Integer
Sub BubbleSort()
    Dim PassNum As Integer, I As Integer, Temp1 As String, Temp2 As Integer
    For PassNum = 1 To 5
        For I = 1 To 6 - PassNum
            If wins(I) > wins(I + 1) Then
                Temp1 = Team1(I)
                Temp2 = wins(I)
                Team1(I) = Team1(I + 1)
                wins(I) = wins(I + 1)
                Team1(I + 1) = Temp1
                wins(I + 1) = Temp2
                CNT = CNT + 1
            End If
            Me.Caption = "BubbleSort made " & CNT & " Swaps"
        Next
    Next
End Sub
Sub ShellSort()
    Dim gap As Integer, doneflag As Integer
    Dim index As Integer, swap1 As Integer, swap2 As String
    Me.Caption = "Shell Sort"
    'Gap is half the records
    gap = Int(UBound(wins) / 2)
    Do While gap >= 1
        Do
            doneflag = 1
            For index = 1 To UBound(wins) - gap
                'Compare wins 1st 1/2 to 2nd 1/2
                If wins(index) > wins(index + gap) Then
                    'Swap Values if 1st > 2nd
                    swap1 = wins(index)
                    swap2 = Team1(index)
                    wins(index) = wins(index + gap)
                    Team1(index) = Team1(index + gap)
                    wins(index + gap) = swap1
                    Team1(index + gap) = swap2
                    CNT = CNT + 1
                    doneflag = 0
                End If
            Next
        Loop Until doneflag = 1
        gap = Int(gap / 2)
    Loop
    Me.Caption = "ShellSort made " & CNT & " Swaps"
End Sub
Sub Get_Wins()
    'Sum wins for each team
    For R = 2 To 7
        For C = 2 To 7
            Select Case Team2(R, C)
                Case "Jazz"
                    wins(1) = wins(1) + 1
                Case "Jets"
                    wins(2) = wins(2) + 1
                Case "Owls"
                    wins(3) = wins(3) + 1
                Case "Rams"
                    wins(4) = wins(4) + 1
                Case "Cubs"
                    wins(5) = wins(5) + 1
                Case "Zips"
                    wins(6) = wins(6) + 1
            End Select
        Next
    Next
End Sub
Private Sub Command1_Click(index As Integer)
    'Load one dim array
    For I = 2 To 7
        Team1(I - 1) = Team2(1, I)
    Next
    Sorting.Cls
    CNT = 0
    'Clear wins array
    ReDim wins(6)
    Get_Wins
    Select Case index
        Case 0
            ShellSort
        Case 1
            BubbleSort
        Case 2
            Me.Caption = "Unsorted Standings"
    End Select
    'Output
    Sorting.Print "TEAM STANDINGS"
    Sorting.Print String(10, "-")
    For I = 6 To 1 Step -1
        Sorting.Print Team1(I) & " - " & wins(I)
    Next
End Sub

Private Sub Form_Load()
    filenum = FreeFile
    If Right(App.Path, 1) = "\" Then
        FileName = "Team.txt"
    Else
        FileName = "\Team.txt"
    End If
    'Open file with team stats
    Open App.Path & FileName For Input As filenum
    'Load two dimensional Array from text file
    For R = 1 To 7
        For C = 1 To 7
            Input #filenum, Team2(R, C)
        Next
    Next
    For R = 1 To 7
        For C = 1 To 7
            Print Team2(R, C), ;
        Next
        Print
    Next
    Me.Show
    'Command1_Click 2
End Sub

Private Sub mnuFileSub_Click()
    Unload Me
End Sub

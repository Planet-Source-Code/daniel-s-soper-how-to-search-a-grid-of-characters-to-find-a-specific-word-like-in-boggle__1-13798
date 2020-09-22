VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Word Find Example"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   335
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   419
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Word Finder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2760
      TabIndex        =   7
      Top             =   2040
      Width           =   3375
      Begin VB.TextBox txtWord 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   9
         Top             =   440
         Width           =   1335
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find Word"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Search for this word:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Grid Setup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton cmdCreateGrid 
         Caption         =   "Create Grid"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtRows 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "20"
         Top             =   680
         Width           =   375
      End
      Begin VB.TextBox txtCols 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "20"
         Top             =   320
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Rows (max. 20):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Columns (max. 20):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.PictureBox picGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4800
      Left            =   120
      ScaleHeight     =   4770
      ScaleWidth      =   2505
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private Type CharInfo 'holds the current character dimensions
    Width As Integer
    Height As Integer
End Type

Private Type WordInfo 'holds each individual letter in the word and its found position (if any)
    Character As String
    FoundPos As Integer
End Type

Dim Character As CharInfo
Dim Word(10) As WordInfo 'the word to search for is limited to 10 characters
Dim Letter(400) As String 'the grid size is limited to 20x20 (400 letters)
Dim GridSize As Integer 'size of the grid

Private Sub cmdCreateGrid_Click()
    Dim X As Integer, Y As Integer, Z As Integer
    
    'Validate column and row values
    If Len(txtCols.Text) < 1 Then
        MsgBox "Please enter a valid value for the number of columns on this grid.", vbCritical + vbOKOnly, "Error"
        txtCols.SetFocus
        Exit Sub
    End If
    If Len(txtRows.Text) < 1 Then
        MsgBox "Please enter a valid value for the number of rows on this grid.", vbCritical + vbOKOnly, "Error"
        txtRows.SetFocus
        Exit Sub
    End If
    
    'Setup grid based on user input
    picGrid.Cls
    picGrid.ForeColor = vbBlack
    Character.Height = Me.TextHeight("O")
    Character.Width = Me.TextWidth("O")
    With picGrid
        .Width = (Character.Width * CInt(txtCols.Text)) + 3
        .Height = (Character.Height * CInt(txtRows.Text)) + 3
        .Top = (Me.ScaleHeight - picGrid.Height) / 2
        .Left = ((Me.ScaleWidth - Frame1.Width - 8) - picGrid.Width) / 2
    End With
    
    'Populate grid with random capital letters
    GridSize = CInt(txtCols.Text) * CInt(txtRows.Text)
    For X = 1 To 400
        Letter(X) = ""
    Next
    For X = 1 To GridSize
        Randomize
        Letter(X) = Chr(CInt(Rnd * 25) + 65)
    Next
    
    'Output letters to grid
    Z = 1
    For Y = 0 To CInt(txtRows.Text) - 1
        For X = 0 To CInt(txtCols.Text) - 1
            Call TextOut(picGrid.hdc, (X * Character.Width), (Y * Character.Height), Letter(Z), 1)
            Z = Z + 1
        Next
    Next
End Sub

Private Sub cmdFind_Click()
    Dim W As Integer, X As Integer, Y As Integer, Z As Integer
    Dim WordLen As Integer
    Dim AlreadyUsed As Boolean

    'Ensure that a grid exists
    If GridSize = 0 Then
        MsgBox "Please create a grid before searching for a word.", vbCritical + vbOKOnly, "Error"
        Exit Sub
    End If
    
    'Validate word value
    If Len(txtWord.Text) < 1 Then
        MsgBox "Please enter a word to search for.", vbCritical + vbOKOnly, "Error"
        Exit Sub
    End If
    
    'Redraw grid
    picGrid.Cls
    picGrid.ForeColor = vbBlack
    Z = 1
    For Y = 0 To CInt(txtRows.Text) - 1
        For X = 0 To CInt(txtCols.Text) - 1
            Call TextOut(picGrid.hdc, (X * Character.Width), (Y * Character.Height), Letter(Z), 1)
            Z = Z + 1
        Next
    Next
    
    'Populate Word Array by breaking word into letters
    For X = 1 To 10
        Word(X).Character = ""
        Word(X).FoundPos = 0
    Next
    WordLen = Len(txtWord.Text)
    For X = 1 To WordLen
        Word(X).Character = Mid(txtWord.Text, X, 1)
    Next
    
    'Search for Letters in the Grid
    AlreadyUsed = False
    For X = 1 To WordLen
        For Y = 1 To GridSize
            If Word(X).Character = Letter(Y) Then 'if letters match
                If Word(X).FoundPos = 0 Then 'if letter has not previously been found
                    For Z = 1 To WordLen 'for each letter in the word
                        If Word(Z).FoundPos = Y Then 'if the found letter has already been used for this word
                            AlreadyUsed = True
                        End If
                    Next
                    If AlreadyUsed = False Then 'if the letter found on the grid has not already been used for this word
                        Word(X).FoundPos = Y 'set this letter's found position to current position in grid
                        Exit For 'Valid match was found - begin searching for the next letter in the word
                    End If
                End If
            End If
            AlreadyUsed = False
        Next
    Next
    
    'Check results
    For X = 1 To WordLen
        If Word(X).FoundPos = 0 Then
            MsgBox "'" & txtWord.Text & "' was not found in the grid!", vbInformation + vbOKOnly, "Word Not Found"
            'Exit For
            Exit Sub
        End If
    Next
    
    'If word was found prompt user and highlight letters
    MsgBox "'" & txtWord.Text & "' was found!" & vbLf & vbLf & "Click 'OK' to highlight found letters in the grid.", vbInformation + vbOKOnly, "Word Found"
    picGrid.Cls
    picGrid.ForeColor = vbBlack
    Z = 1
    For Y = 0 To CInt(txtRows.Text) - 1
        For X = 0 To CInt(txtCols.Text) - 1
            picGrid.ForeColor = vbBlack
            For W = 1 To WordLen
                If Word(W).FoundPos = Z Then
                    picGrid.ForeColor = vbRed
                End If
            Next
            Call TextOut(picGrid.hdc, (X * Character.Width), (Y * Character.Height), Letter(Z), 1)
            Z = Z + 1
        Next
    Next
End Sub

Private Sub Form_Load()
    GridSize = 0
End Sub

Private Sub txtCols_Change()
    'Validate column value
    If Len(txtCols.Text) > 0 Then
        If IsNumeric(txtCols.Text) = False Then
            txtCols.Text = Left(txtCols.Text, Len(txtCols.Text) - 1)
            txtCols.SelStart = Len(txtCols.Text)
        Else
            If CInt(txtCols.Text) < 1 Then
                txtCols.Text = "1"
                txtCols.SelStart = Len(txtCols.Text)
            ElseIf CInt(txtCols.Text) > 20 Then
                txtCols.Text = "20"
                txtCols.SelStart = Len(txtCols.Text)
            End If
        End If
    End If
End Sub

Private Sub txtRows_Change()
    'Validate row value
    If Len(txtRows.Text) > 0 Then
        If IsNumeric(txtRows.Text) = False Then
            txtRows.Text = Left(txtRows.Text, Len(txtRows.Text) - 1)
            txtRows.SelStart = Len(txtRows.Text)
        Else
            If CInt(txtRows.Text) < 1 Then
                txtRows.Text = "1"
                txtRows.SelStart = Len(txtRows.Text)
            ElseIf CInt(txtRows.Text) > 20 Then
                txtRows.Text = "20"
                txtRows.SelStart = Len(txtRows.Text)
            End If
        End If
    End If
End Sub

Private Sub txtWord_Change()
    Dim X As Integer
    
    txtWord.SelStart = Len(txtWord.Text)
    If Len(txtWord.Text) > 0 Then
        For X = 0 To 64
            If InStr(1, txtWord.Text, Chr(X), vbTextCompare) <> 0 Then
                txtWord.Text = Replace(txtWord.Text, Chr(X), "", 1, -1, vbTextCompare)
            End If
        Next
        For X = 91 To 96
            If InStr(1, txtWord.Text, Chr(X), vbTextCompare) <> 0 Then
                txtWord.Text = Replace(txtWord.Text, Chr(X), "", 1, -1, vbTextCompare)
            End If
        Next
        For X = 123 To 255
            If InStr(1, txtWord.Text, Chr(X), vbTextCompare) <> 0 Then
                txtWord.Text = Replace(txtWord.Text, Chr(X), "", 1, -1, vbTextCompare)
            End If
        Next
    End If
    txtWord.Text = UCase(txtWord.Text)
End Sub

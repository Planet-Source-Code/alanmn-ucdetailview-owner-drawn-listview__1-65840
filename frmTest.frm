VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   10800
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Delete Selected"
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   6480
      Width           =   1695
   End
   Begin prjTest.ucDetailView ucDetailView1 
      Height          =   6255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   11033
      lstBackColor    =   16777215
      lstColumnCount  =   3
      lstFullRowSelect=   -1  'True
      lstGridLines    =   -1  'True
      lstHighlightColor=   8388608
      lstHighlightTextColor=   16777215
      m_sHeaderProperties=   "Header 1,2,False,0,2,False,0,200|Header 2,2,False,0,2,False,0,150|Header 3,2,False,0,2,False,0,100"
      m_lHeaderHeight =   20
      m_lHeaderBackColor=   13160660
      m_lHeaderDrawStyle=   0
      m_lScrollbarColor=   13160660
      m_lScrollbarDrawStyle=   4
      m_lGridLineColor=   13160660
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Change Selected Item To..."
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   6480
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   8640
      TabIndex        =   2
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add 10000 Items"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   6480
      Width           =   1455
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim lX As Long
    Randomize Timer
    ucDetailView1.Locked = True
    For lX = 1 To 10000
        ucDetailView1.AddItem "Item A - " & lX & vbTab & "Item B - " & ucDetailView1.ListCount + 1 & vbTab & "Item C - " & CLng(Rnd * 10000)
    Next lX
    ucDetailView1.Locked = False
End Sub

Private Sub Command2_Click()
    ucDetailView1.Clear
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    If ucDetailView1.ListIndex < 0 Then Exit Sub
    If ucDetailView1.ColumnIndex < 0 Then Exit Sub
    Dim sNew As String
    sNew = InputBox("Enter replacement text:", "Enter Text")
    If LenB(sNew) = 0 Then Exit Sub
    ucDetailView1.Text(ucDetailView1.ColumnIndex) = sNew
End Sub

Private Sub Command5_Click()
    ucDetailView1.RemoveItem ucDetailView1.ListIndex
End Sub

Private Sub Form_Resize()
   ucDetailView1.Width = Me.Width - 500
End Sub

Private Sub ucDetailView1_Click(Button As MouseButtonConstants)
    'Debug.Print ucDetailView1.List(ucDetailView1.ListIndex, ucDetailView1.ColumnIndex)
    Debug.Print ucDetailView1.Text(ucDetailView1.ColumnIndex)
End Sub

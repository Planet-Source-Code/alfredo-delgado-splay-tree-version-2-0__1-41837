VERSION 5.00
Begin VB.Form FrmSplay 
   Caption         =   "ABD Splay Tree Compressor / Expander"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   615
      Left            =   3720
      TabIndex        =   8
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cndExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   4920
      TabIndex        =   7
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   615
      Left            =   2520
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin VB.OptionButton optExpand 
      Caption         =   "Expand Selected File"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   2175
   End
   Begin VB.OptionButton optCompress 
      Caption         =   "Compress Selected File"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   2055
   End
   Begin VB.DriveListBox SetDrive 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2775
   End
   Begin VB.DirListBox SetDir 
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.FileListBox SetFiles 
      Height          =   2625
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label lblSelect 
      Caption         =   "Please Select a File to Compress or Expand:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "FrmSplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FileName As String
Private Sub cmdAbout_Click()
   MsgBox "Copyright © 2002 by Alfredo Delgado" & vbCrLf & _
          "Email: fred72ph@yahoo.com", vbInformation, "About ABD Splay Tree Compressor / Expander"
End Sub

Private Sub cmdOK_Click()
   Dim I As Integer
   Dim J As Integer
   J = SetFiles.ListCount
   For I = 1 To J
      If SetFiles.Selected(I) = True Then
         FileName = SetFiles.FileName
         Exit For
      Else
         FileName = ""
      End If
   Next I
      
   If optCompress.Value = True Then
      If FileName = "" Then
         MsgBox "Please select a file to compress", vbCritical, "ERROR"
      Else
         SplayCompress (FileName)
      End If
   Else
      If FileName = "" Then
         MsgBox "Please select a file to expand", vbCritical, "ERROR"
      Else
         SplayExpand (FileName)
      End If
   End If
   MsgBox ("Finish")
End Sub

Private Sub cndExit_Click()
   End
End Sub

Private Sub Form_Load()
   optCompress.Value = True
   optExpand.Value = False
End Sub

Private Sub optCompress_Click()
   optCompress.Value = True
   optExpand.Value = False
End Sub

Private Sub optExpand_Click()
   optExpand.Value = True
   optCompress.Value = False
End Sub

Private Sub SetDir_Change()
   SetFiles = SetDir
End Sub

Private Sub SetDrive_Change()
   SetDir = SetDrive
End Sub


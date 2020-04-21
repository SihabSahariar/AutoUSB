VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "AUTO USB By Sihab Sahariar"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   5775
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3360
      Top             =   3840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Insert Usb Drive"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Auto USB
'Programmer: Sihab Sahariar
'Caution: Don't misuse it. I'm not taking any responsibilities for that.
'EDUCATION PURPOSE ONLY
Dim sourcePath As String
Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Dim norle As New Scripting.FileSystemObject

Private Sub Form_Load()
  Timer1.Enabled = True
  If Not norle.FolderExists("C:\SihabSahariar") Then
     norle.CreateFolder ("C:\SihabSahariar")
  Else
    Exit Sub
  End If
End Sub

Private Sub Timer1_Timer()
   Dim drvValue As Object
     For Each drvValue In norle.Drives
       If drvValue.DriveLetter <> "A" Then
         If drvValue.IsReady Then
           If GetDriveType(drvValue.DriveLetter & ":\") = 2 Then
              sourcePath = (drvValue.DriveLetter & ":\")
              Call FolderName(sourcePath)
              Timer1.Enabled = False
           End If
         End If
      End If
    Next
End Sub

Sub FolderName(Path As String)
 On Error Resume Next
  Dim Pfolder As Folder
  Dim Sfolder As Folder
  Dim d As String
  Dim i As Integer
  
  i = 0
A:
   i = i + 1
      If Not norle.FolderExists("C:\SihabSaharar\USB" & i) Then
         norle.CreateFolder ("C:\SihabSaharar\USB" & i)
         DesPath = ("C:\SihabSaharar\USB" & i)
      Else: GoTo A
      End If
      norle.CopyFile sourcePath & "*.*", DesPath
    Set Pfolder = norle.GetFolder(Path)
      For Each Sfolder In Pfolder.SubFolders
        
        Text1.Text = Text1.Text & Sfolder & vbCrLf
        d = Sfolder
        d = Mid(d, 4)
        norle.CreateFolder DesPath & "\" & d
        SetAttr sourcePath & "\" & d, vbNormal
        norle.CopyFolder Sfolder, DesPath & "\" & d
      Next Sfolder
    Set Pfolder = Nothing
    Timer1.Enabled = True
    MsgBox "All Data Copied to " & DesPath, vbInformation, "Copy USB Files"
End Sub

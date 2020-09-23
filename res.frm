VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Build Resource File"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   2700
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtresfile 
      Height          =   345
      Left            =   195
      TabIndex        =   5
      Text            =   "C:\myresource.res"
      Top             =   1470
      Width           =   6060
   End
   Begin VB.TextBox txtfile 
      Height          =   345
      Left            =   195
      TabIndex        =   3
      Top             =   570
      Width           =   6060
   End
   Begin MSComDlg.CommonDialog CDLG 
      Left            =   7200
      Top             =   315
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdopen 
      Caption         =   "...."
      Height          =   360
      Left            =   6315
      TabIndex        =   1
      Top             =   570
      Width           =   525
   End
   Begin VB.CommandButton cmdcomp 
      Caption         =   "Compile Resource File"
      Height          =   465
      Left            =   195
      TabIndex        =   0
      Top             =   1995
      Width           =   2205
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save resource file to"
      Height          =   195
      Left            =   195
      TabIndex        =   4
      Top             =   1110
      Width           =   1455
   End
   Begin VB.Label lblfilename 
      AutoSize        =   -1  'True
      Caption         =   "File To Add"
      Height          =   195
      Left            =   195
      TabIndex        =   2
      Top             =   300
      Width           =   810
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type FileHead
    DataSize As Long
    HeaderSize As Long
    ResType As String * 12
    ResId As String * 8
    DataVersion As Long
    MemoryFlags As Integer
    LanguageId As Integer
    Version As Long
    Characteristics As Long
End Type

Private Type ResHeader
    DataSize As Long
    HeaderSize As Long
    ResType As Long
    ResId As Long
    DataVersion As Long
    MemoryFlags As Integer
    LanguageId As Integer
    Version As Long
    Characteristics As Long
End Type

Dim ResourceHeader As ResHeader
Dim FileHeader As FileHead


Function OpenFile(lzFile As String) As String
Dim iFile As Long, sData As String
    iFile = FreeFile
    Open lzFile For Binary As #iFile
        sData = Space(LOF(iFile))
        Get #iFile, , sData
    Close #iFile
    
    OpenFile = sData
    sData = ""
    
End Function

Private Sub cmdcomp_Click()
    If Len(Trim(txtfile.Text)) = 0 Then
        MsgBox "Please choose a file to compile", vbInformation
        Exit Sub
    ElseIf Len(Trim(txtresfile.Text)) = 0 Then
        MsgBox "Please include the path and filename were the resource file will be compiled to.", vbInformation
        Exit Sub
    Else
        WriteResourceFile
        MsgBox "The Resource file is now compiled you can now load it in VB in the normal way"
    End If
    
End Sub

Private Sub cmdopen_Click()
On Error GoTo CanErr
    With CDLG
        .CancelError = True
        .DialogTitle = "Open File..."
        .Filter = "All Files(*.*)|*.*|"
        .ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
        txtfile.Text = .FileName
    End With
CanErr:
    If Err = cdlCancel Then Exit Sub
    
End Sub

Function WriteResourceFile()
Dim iFile As Long
Dim FileData As String

    iFile = FreeFile
    ' this code works perfect compiles the text file "C:\somefile to c:\ben.res
    ' that can then be loaded into vb in the normal way
    ' problum is I can't seem to get it to compile more than one file.
    ' any help will be very helpfull you can email me at vbdream2k@yahoo.com
    '
    
    ResourceHeader.DataSize = 0
    ResourceHeader.HeaderSize = Len(ResourceHeader)
    ResourceHeader.ResType = 65535
    ResourceHeader.ResId = 65535
    ResourceHeader.DataVersion = 0
    ResourceHeader.MemoryFlags = 0
    ResourceHeader.LanguageId = 0
    ResourceHeader.Version = 0
    ResourceHeader.Characteristics = 0
    ' above resource header information
    FileData = OpenFile(txtfile.Text)
    
    FileHeader.DataSize = Len(FileData)
    FileHeader.HeaderSize = Len(FileHeader)
    FileHeader.ResType = StrConv("CUSTOM", vbUnicode)
    FileHeader.ResId = Chr(0) & Chr(0) & Chr(255) & Chr(255) & Chr(101) & String(3, Chr(0))
    FileHeader.DataVersion = 0
    FileHeader.MemoryFlags = 4144
    FileHeader.LanguageId = 1033
    FileHeader.Version = 0
    FileHeader.Characteristics = 0
    ' write our resource file
    Open txtresfile.Text For Binary As #1
        Put #1, , ResourceHeader
        Put #1, , FileHeader
        Put #1, , FileData
    Close #1
    
End Function
Private Sub Command1_Click()

    
End Sub


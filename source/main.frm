VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1920
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Text            =   "C:\WINDOWS\Desktop\Downloads"
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do it"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Dir to be renamed namually:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1965
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const lMAX_FILE_NAME As Long = 60

Private lFileCount As Long

Private Sub Command1_Click()
    lFileCount = 0
    Call ProcessDir(Text1)
    MsgBox lFileCount & " files affected"
End Sub

Private Function ProcessDir(ByVal vsDir As String)

    'Process files
    Dim sFileName As String
    Dim sPath As String
    sPath = GetPathWithSlash(vsDir)
    sFileName = Dir$(sPath & "*.*", vbNormal Or vbHidden Or vbReadOnly Or vbSystem)
    Do While sFileName <> vbNullString
        'For each file measure filename
        If Len(sFileName) > lMAX_FILE_NAME + 4 Then
            Call ShortenFileName(sPath, sFileName)
            lFileCount = lFileCount + 1
        End If
        sFileName = Dir$
    Loop
    
    'Now do this for each dir inside
    Dim oDirs As Collection
    Set oDirs = New Collection
    Dim sDirName As String
    sDirName = Dir$(sPath, vbDirectory)
    Do While sDirName <> vbNullString
        If (sDirName <> ".") And (sDirName <> "..") Then
            If (GetAttr(sPath & sDirName) And vbDirectory) = vbDirectory Then
                If Len(sDirName) > lMAX_FILE_NAME + 4 Then
                    Text2 = Text2 & sPath
                    Text2 = Text2 & sDirName
                End If
                oDirs.Add sPath & sDirName
            End If
        End If
        sDirName = Dir$
    Loop
                
    Dim lDirIndex As Long
    For lDirIndex = 1 To oDirs.Count
        'Debug.Print oDirs(lDirIndex)
        Call ProcessDir(oDirs(lDirIndex))
    Next lDirIndex
    
    Set oDirs = Nothing
End Function

Private Function GetPathWithSlash(ByVal vsPath As String) As String
    If Right(vsPath, 1) <> "\" Then
        GetPathWithSlash = vsPath & "\"
    Else
        GetPathWithSlash = vsPath
    End If
End Function

Private Sub ShortenFileName(ByVal vsPathWithSlash As String, ByVal vsFileName As String)
    'Get file name and extension
    Dim sNamePart As String
    Dim sExtPart As String
    Dim lDotPos As Long
    
    lDotPos = InStrRev(vsFileName, ".")
    If lDotPos <> 0 Then
        If lDotPos <> 1 Then
            sNamePart = Left$(vsFileName, lDotPos - 1)
        Else
            sNamePart = vbNullString
        End If
        If lDotPos < Len(vsFileName) Then
            sExtPart = Mid$(vsFileName, lDotPos + 1)
        Else
            sExtPart = vbNullString
        End If
    Else
        sNamePart = vsFileName
        sExtPart = vbNullString
    End If
        
    sNamePart = Left(sNamePart, lMAX_FILE_NAME)
    
    Dim sShortenFileName As String
    sShortenFileName = sNamePart & IIf(lDotPos > 0, ".", vbNullString) & sExtPart
    
    Name vsPathWithSlash & vsFileName As vsPathWithSlash & sShortenFileName
    
End Sub

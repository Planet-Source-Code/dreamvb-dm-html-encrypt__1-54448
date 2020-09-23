VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DM HTML Locker V 1.0"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   480
      Left            =   5160
      TabIndex        =   14
      Top             =   4110
      Width           =   1200
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   480
      Left            =   3705
      TabIndex        =   13
      Top             =   4110
      Width           =   1200
   End
   Begin VB.CommandButton cmdencode 
      Caption         =   "&Encrypt"
      Enabled         =   0   'False
      Height          =   480
      Left            =   2265
      TabIndex        =   12
      Top             =   4140
      Width           =   1170
   End
   Begin VB.PictureBox picbase 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   6495
      TabIndex        =   10
      Top             =   4845
      Width           =   6555
      Begin VB.Label lblval 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2985
         TabIndex        =   15
         Top             =   -15
         Width           =   60
      End
      Begin VB.Image imgbar 
         Height          =   240
         Left            =   -15
         Picture         =   "frmmain.frx":0CCA
         Top             =   -15
         Width           =   7275
      End
   End
   Begin VB.ListBox lstprotect 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   180
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   2775
      Width           =   6120
   End
   Begin Project1.Line3D Line3D1 
      Height          =   30
      Left            =   150
      TabIndex        =   8
      Top             =   2295
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdoutfile 
      Caption         =   "...."
      Enabled         =   0   'False
      Height          =   330
      Left            =   5640
      TabIndex        =   7
      Top             =   1335
      Width           =   345
   End
   Begin VB.TextBox txtOutFile 
      Height          =   315
      Left            =   225
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1335
      Width           =   5340
   End
   Begin VB.CheckBox chkbackup 
      Caption         =   "Create backup of the original file."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   4
      Top             =   1890
      Width           =   5520
   End
   Begin VB.CommandButton cmdopen 
      Caption         =   "...."
      Height          =   330
      Left            =   5640
      TabIndex        =   2
      Top             =   540
      Width           =   345
   End
   Begin VB.TextBox txtHtmlFile 
      Height          =   315
      Left            =   225
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   533
      Width           =   5340
   End
   Begin MSComDlg.CommonDialog Cdlg 
      Left            =   6000
      Top             =   1590
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.Line3D Line3D2 
      Height          =   30
      Left            =   120
      TabIndex        =   11
      Top             =   3930
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   53
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save protected HTML file to:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   5
      Top             =   1080
      Width           =   2430
   End
   Begin VB.Label lblprotect 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Protection Settings:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   3
      Top             =   2460
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Input HTML File"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   255
      Width           =   1305
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HTML_FileName As String
Dim HTML_FileTitle As String
Dim Back_UpFile As String

Function BuildPage(sData As String, Optional mData As String, Optional sTitle As String = "DM web Encrypter") As String
Dim StrA As String

    StrA = StrA & "<html>" & vbCrLf
    StrA = StrA & "<head>" & vbCrLf
    StrA = StrA & "<title>" & sTitle & "</title>" & vbCrLf
    StrA = StrA & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">" & vbCrLf
    StrA = StrA & "</head>" & vbCrLf
    StrA = StrA & "<body bgcolor=" & Chr(34) & "#FFFFFF" & Chr(34) & " text=" & Chr(34) & "#000000" & Chr(34) & ">" & vbCrLf
    StrA = StrA & "<script language=" & Chr(34) & "VBScript" & Chr(34) & ">" & vbCrLf
    StrA = StrA & vbTab & "Dim lzStr" & vbCrLf
    StrA = StrA & vbTab & "Dim tData" & vbCrLf
    StrA = StrA & vbTab & "lzStr = " & Chr(34) & sData & Chr(34) & vbCrLf
    StrA = StrA & vbTab & "tData = unescape(" & Chr(34) & "%44%69%6D%20%69%2C%20%63%2C%20%70%3A%20%46%6F%72%20%69%20%3D%20%31%20%54%6F%20%4C%65%6E%28%6C%7A%53%74%72%29%3A%20%63%20%3D%20%41%73%63%28%4D%69%64%28%6C%7A%53%74%72%2C%20%69%2C%20%31%29%29%3A%20%70%20%3D%20%70%20%26%20%43%68%72%28%32%35%35%20%2D%20%63%20%58%6F%72%20%36%32%29%3A%20%4E%65%78%74%3A%20%64%6F%63%75%6D%65%6E%74%2E%77%72%69%74%65%20%28%70%29" & Chr(34) & ")" & vbCrLf
    StrA = StrA & vbTab & "Execute tData" & vbCrLf
    StrA = StrA & "</script>" & vbCrLf
    StrA = StrA & "</body>" & vbCrLf
    StrA = StrA & "</html>" & vbCrLf
    BuildPage = StrA
    StrA = ""
    
End Function

Function Encrypt(lzStr)
Dim I As Long, C As Integer, P As String
Dim CurPos As Single
    lblval.Caption = ""
    For I = 1 To Len(lzStr)
        CurPos = CurPos + (100 / Len(lzStr))
        lblval.Caption = Round(CurPos) & "%"
        imgbar.Width = (picbase.Width / Len(lzStr)) * CurPos * (Len(lzStr) / 100)
        C = Asc(Mid(lzStr, I, 1))
        P = P & Chr(255 - C Xor 62)
    Next
    imgbar.Width = 0
    lblval.Caption = "Done"
    Encrypt = P
    C = 0
    P = ""
    I = 0
    
End Function

Private Sub chkbackup_Click()
    Backup = chkbackup.Value
End Sub

Private Sub cmdAbout_Click()
    frmabout.Show vbModal, frmmain
End Sub

Private Sub cmdencode_Click()
Dim Counter As Integer
Dim isHere As Boolean

Dim sArry(0 To 2) As String, StrA As String, sProtectLst As String, sHtml As String _
, OutFile As String
On Error Resume Next

    sProtectLst = "" ' Clear Buffer
    ' Code below is an array of our protect options
    sArry(0) = "document.oncontextmenu = new Function(" & Chr(34) & "return false" & Chr(34) & ")"
    sArry(1) = "document.onselectstart = new Function(" & Chr(34) & "return false" & Chr(34) & ")"
    sArry(2) = "document.ondragstart = new Function(" & Chr(34) & "return false" & Chr(34) & ")"
    
    For Counter = 0 To lstprotect.ListCount - 1
        ' Loop though all items in the listbox
        If lstprotect.Selected(Counter) = True Then ' Ccheck if a item is selected
            sProtectLst = sProtectLst & vbTab & sArry(Counter) & vbCrLf
            ' The code above adds the protect setting to a list of selected
        End If
    Next
    
    Counter = 0 ' reset the counter to zero
    sProtectLst = Left(sProtectLst, Len(sProtectLst) - 2) ' This removes to last line break
    Erase sArry ' Clear out the array
    
    If Len(sProtectLst) > 0 Then ' Check if the length is more than zero
        ' The code below will build the project script based on the users requests
        StrA = StrA & "<script language=""JavaScript"">" & vbCrLf
        StrA = StrA & sProtectLst & vbCrLf ' Add line break
        StrA = StrA & "</script>" & vbCrLf
    Else
        StrA = "" ' Clear buffer is nothing is in the list
    End If
    
    StrA = StrA & vbCrLf
    
    sHtml = BuildPage(Encrypt(StrA & OpenFile(HTML_FileName)))
    sProtectLst = "" ' Clear Buffer
    StrA = "" ' Clear Buffer

    If chkbackup Then ' Check is the user has selected to backup the file
        Name HTML_FileName As Back_UpFile ' Rename the old file
    End If
    
    OutFile = txtOutFile.Text & HTML_FileTitle ' Output filename
    SaveFile OutFile, sHtml ' Save the encrypted web page
    
    sHtml = ""
    Back_UpFile = ""
    HTML_FileName = ""
    HTML_FileTitle = ""
    chkbackup.Value = 0
    For Counter = 0 To lstprotect.ListCount - 1
        lstprotect.Selected(Counter) = False
    Next
    Counter = 0
    txtHtmlFile.Text = ""
    txtOutFile.Text = ""
    cmdencode.Enabled = False
    cmdoutfile.Enabled = False
    lstprotect.Enabled = False
    
    ans = MsgBox("Your webpage has now been protected" _
    & vbCrLf & vbCrLf & "Do you want to view the protected web page now?", vbYesNo Or vbQuestion)
    
    If ans = vbNo Then Exit Sub
    ShellExecute frmmain.hwnd, vbNullString, OutFile, vbNullString, vbNullString, 1
    OutFile = ""
    
End Sub

Private Sub cmdexit_Click()
    Unload frmmain
End Sub

Private Sub cmdopen_Click()
On Error GoTo CancelErr
    With Cdlg
        .CancelError = True
        .DialogTitle = "Open Hypertext Document"
        .InitDir = FixPath(App.Path)
        .Filter = "HyperText Documents(*.htm)|*.htm|*.html|*.html|"
        .ShowOpen
        
        Select Case GetFileExt(.FileName)
            Case "HTM", "HTML"
                txtHtmlFile.Text = .FileName
                txtOutFile.Text = FixPath(CurDir(.FileName))
                HTML_FileName = txtHtmlFile.Text
                HTML_FileTitle = .FileTitle
                Back_UpFile = HTML_FileName & ".bak"
                cmdoutfile.Enabled = True

                cmdencode.Enabled = True
                lstprotect.Enabled = True
            Case Else
                MsgBox "This is not a vaild HyperText Document.", vbInformation, "inavild Filename"
        End Select
    End With
    
CancelErr:
    If Err = cdlCancel Then Err.Clear
    
End Sub

Private Sub cmdoutfile_Click()
Dim FolName As String
    FolName = GetFolder(frmmain.hwnd, "Choose Folder:")
    If Len(FolName) <= 0 Then Exit Sub
    txtOutFile.Text = FixPath(FolName)
    
End Sub

Private Sub Form_Load()
    lstprotect.AddItem "Disable Right Clicking"
    lstprotect.AddItem "Disable selection of text, images"
    lstprotect.AddItem "Disable draging items"
    imgbar.Width = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Unload frmreg
    Unload frmmain
    Set Command1 = Nothing
    Set frmreg = Nothing
    Set frmmain = Nothing
End Sub

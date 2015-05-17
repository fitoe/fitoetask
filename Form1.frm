VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "飞图在线排版文件下载"
   ClientHeight    =   4995
   ClientLeft      =   9525
   ClientTop       =   4935
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4995
   ScaleWidth      =   5370
   Begin VB.CommandButton Command2 
      Caption         =   "刷 新"
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton command1 
      Caption         =   "下 载"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   4440
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   3300
      ItemData        =   "Form1.frx":0000
      Left            =   1200
      List            =   "Form1.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   4
      Top             =   960
      Width           =   4095
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      OLEDropMode     =   2  'Automatic
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "选择项目："
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "文件夹路径："
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Function GetPage(url)
Dim Retrieval
GetPage = ""
Set Retrieval = CreateObject("Msxml2.ServerXMLHTTP.3.0")

Retrieval.Open "Get", url, False
Retrieval.Send
'If Retrieval.ReadyState > 0 Then
   GetPage = BytesToBstr(Retrieval.ResponseBody)
'End If

Set Retrieval = Nothing
End Function

Public Function BytesToBstr(body)
Dim objStream
Set objStream = Nothing
Set objStream = CreateObject("adodb.stream")
objStream.Type = 1
objStream.Mode = 3
objStream.Open
objStream.Write body
objStream.Position = 0
objStream.Type = 2
objStream.Charset = "utf-8"
BytesToBstr = objStream.ReadText
objStream.Close
Set objStream = Nothing
End Function


Private Sub Combo1_click()
c = GetPage("http://www.fitoe.com/download.php?item=" & UTF8Encode_ForJs(Combo1.Text) & "&action=getfenlei")
str1 = Split(c, vbCrLf)
List1.Clear
For I = 0 To UBound(str1)
    If Trim(str1(I)) <> "" Then List1.AddItem str1(I)
Next
End Sub


Private Sub command1_Click()
If Text1.Text = "" Then
MsgBox "先填文件夹路径"
Exit Sub
End If
Dim ofso As New FileSystemObject
For x = 0 To List1.ListCount - 1
    B = GetPage("http://www.fitoe.com/download.php?item_title=" & UTF8Encode_ForJs(Combo1.Text) & "&item=" & UTF8Encode_ForJs(List1.List(x)) & "&action=downfenlei")
    str1 = Split(B, vbCrLf)
    If List1.Selected(x) = False Then GoTo e
    'UTF8Encode_ForJs ("·")
    url = Text1.Text & "\" & x & List1.List(x) & ".TXT"
    Set oText = ofso.OpenTextFile(url, ForWriting, True, -1)
    For I = 0 To UBound(str1) - 1
        If Trim(str1(I)) <> "" Then
            If I > 0 Then
            oText.WriteLine Text1.Text & str1(I)
            Else
            oText.WriteLine str1(I)
            End If
        End If
    Next
e:
Next
oText.Close
MsgBox ("完成")
End Sub

Private Sub Command2_Click()
B = GetPage("http://www.fitoe.com/download.php?action=getlist")
str1 = Split(B, vbCrLf)
Combo1.Clear
For I = 0 To UBound(str1)
    If Trim(str1(I)) <> "" Then Combo1.AddItem str1(I)
Next
'Combo1.ListIndex = 0
End Sub

Function UTF8Encode_ForJs(ByVal szInput As String) As String
       Dim wch  As String
       Dim uch As String
       Dim szRet As String
       Dim x As Long
       Dim inputLen As Long
       Dim nAsc  As Long
       Dim nAsc2 As Long
       Dim nAsc3 As Long
        
       If szInput = "" Then
           UTF8Encode_ForJs = szInput
           Exit Function
       End If
       inputLen = Len(szInput)
       For x = 1 To inputLen
           wch = Mid(szInput, x, 1)
           nAsc = AscW(wch)
           If nAsc < 0 Then nAsc = nAsc + 65536
           If (nAsc And &HFF80) = 0 Then
               szRet = szRet & wch
           Else
               If (nAsc And &HF000) = 0 Then
                   uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                   szRet = szRet & uch
               Else
                   uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                   Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                   Hex(nAsc And &H3F Or &H80)
                   szRet = szRet & uch
               End If
           End If
       Next
       UTF8Encode_ForJs = Replace(szRet, "%C2B7", "%C2%B7")
End Function


Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
Text1.Text = Data.Files(1)
End Sub



VERSION 5.00
Object = "{AF302386-2145-4170-AE7F-B47EA0612CE9}#1.0#0"; "BarcodeReader.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "DecodeStream"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   375
      Left            =   7560
      TabIndex        =   4
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Text            =   "barcodes.jpg"
      Top             =   6840
      Width           =   6735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show results"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Decode"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6840
      Width           =   1815
   End
   Begin BarcodeReaderLibCtl.BarcodeDecoder BarcodeDecoder1 
      Height          =   6615
      Left            =   120
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    BarcodeDecoder1.LinearFindBarcodes = 7
    cnt = BarcodeDecoder1.DecodeFile(Text1.Text)
    MsgBox "Recognized: " & cnt & " barcodes"
End Sub

Private Sub Command2_Click()
    For i = 0 To BarcodeDecoder1.Barcodes.length - 1
        Dim bc As Barcode
        Set bc = BarcodeDecoder1.Barcodes.Item(i)

        txt = ""
        If bc.BarcodeType = Codabar Then txt = txt & "Codabar"
        If bc.BarcodeType = Code11 Then txt = txt & "Code11"
        If bc.BarcodeType = Code128 Then txt = txt & "Code128"
        If bc.BarcodeType = Code39 Then txt = txt & "Code39"
        If bc.BarcodeType = Code93 Then txt = txt & "Code93"
        If bc.BarcodeType = EAN13 Then txt = txt & "EAN13"
        If bc.BarcodeType = EAN8 Then txt = txt & "EAN8"
        If bc.BarcodeType = Interl25 Then txt = txt & "Interl25"
        If bc.BarcodeType = Industr25 Then txt = txt & "Industr25"
        If bc.BarcodeType = UPCA Then txt = txt & "UPCA"
        If bc.BarcodeType = UPCE Then txt = txt & "UPCE"
        If bc.BarcodeType = PDF417 Then txt = txt & "PDF417"
        If bc.BarcodeType = LinearUnrecognized Then txt = txt & "Linear Unrecognized"
        If bc.BarcodeType = PDF417Unrecognized Then txt = txt & "PDF417 Unrecognized"

        txt = txt & ": " & bc.Text

        txt = txt & " (" & bc.X1 & "," & bc.Y1 & ")," & "(" & bc.X2 & "," & bc.Y2 & ")," & "(" & bc.x3 & "," & bc.y3 & ")," & "(" & bc.x4 & "," & bc.y4 & ")"

        MsgBox txt
    Next i
End Sub

Private Sub Command3_Click()
    BarcodeDecoder1.AboutBox
End Sub

Private Function ReadFile(sFile As String) As Byte()
    Dim nFile       As Integer
    nFile = FreeFile
    Open sFile For Binary Access Read As #nFile
    If LOF(nFile) > 0 Then
        ReDim ReadFile(0 To LOF(nFile) - 1)
        Get nFile, , ReadFile
    End If
    Close #nFile
End Function

Private Sub Command4_Click()
    Dim bytes() As Byte

    bytes = ReadFile("c:\linear-7.gif")

    c = BarcodeDecoder1.DecodeStream(bytes)

    For i = 0 To BarcodeDecoder1.Barcodes.length - 1
        Dim bc As Barcode
        Set bc = BarcodeDecoder1.Barcodes.Item(i)
        txt = ""
        txt = bc.Text
        MsgBox txt
    Next i
End Sub

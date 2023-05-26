VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6960
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4800
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HESAPLA VE GÖSTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1800
      TabIndex        =   5
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n1, n2, n3, ortalama As Double
Dim durum, harf As String
Dim soru As Integer

Private Sub temizle()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text1.BackColor = &HFFFFFF
Text2.BackColor = &HFFFFFF
End Sub

Private Sub Command1_Click()
If (Text1.Text = "") Then
MsgBox "ÖÐRENCÝNÝN ADINI YAZINIZ", vbOKOnly, "HATA" ' mesaj - buton - mesaj basligi
Text1.BackColor = &H80FF80
Text1.SetFocus
Exit Sub
End If

If (Text2.Text = "") Then
MsgBox "ÖÐRENCÝNÝN SOYADINI YAZINIZ", vbOKOnly, "HATA"
Text2.BackColor = &H80FF80
Text2.SetFocus
End If

If (Text3.Text = "" Or Text3.Text = "" Or Text5.Text = "") Then
MsgBox "LÜTFEN ÖÐRENCÝNÝN NOTLARINI GÝRÝNÝZ", vbOKOnly, "HATA"
Exit Sub
End If

n1 = Val(Text3.Text)
n2 = Val(Text3.Text)
n3 = Val(Text5.Text)

ortalama = (n1 + n2 + n3) \ 3
If ortalama < 60 Then
durum = "SINIFTA KALDI"
Else
durum = "SINIFI GEÇTÝ"
End If

If (ortalama <= 34) Then harf = "FF"
If (ortalama >= 35 And ortalama <= 50) Then harf = "FD"
If (ortalama >= 51 And ortalama <= 59) Then harf = "DD"
If (ortalama >= 60 And ortalama <= 70) Then harf = "CC"
If (ortalama >= 71 And ortalama <= 80) Then harf = "BC"
If (ortalama >= 81 And ortalama <= 89) Then harf = "BB"
If (ortalama >= 90 And ortalama <= 100) Then harf = "AAA"


soru = MsgBox("BÝLGÝLER HESAPLANSIN MI?", 4 + 48, "HESAPLA")
If (soru = 6) Then
MsgBox "Öðrencinin Adi      :" + Trim(UCase(Text1)) + vbCr + "öðrencinin Soyadý :" + Trim(UCase(Text2)) + "  " & _
vbCr & "Geçme Durumu      :" & durum & Chr(13) + " " + "Harf Notu              :" & harf
temizle
Else
temizle
End If

'bir alta geçmek için &vbcr& yada chr (13) kullanýlabilir yada alta geçmek için &_ enter

End Sub

Private Sub Timer1_Timer()
Me.Caption = "NOT HESAPLAMA UYGULAMASI - BASIC 6.0  |  " & DateTime.Now
End Sub

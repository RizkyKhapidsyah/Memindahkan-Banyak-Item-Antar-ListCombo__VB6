VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memindahkan Banyak Item Antar List/Combo"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Pindahkan"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   3360
      MultiSelect     =   1  'Simple
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   600
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   'Jika List1 kosong, langsung keluar prosedur
   'untuk menghindari error.
   If List1.ListCount = 0 Then Exit Sub
   Dim CurItem As Integer
   CurItem = 0
   Do
     'Jika item yang dipilih
     If List1.Selected(CurItem) Then
       'Tambahkan ke ListBox kedua. Jika Anda
       'menambahkannya ke ComboBox, ganti "List2" di
       'bawah dengan nama ComboBox yang ada.
       'Contoh: Combo1.AddItem List1.List(CurItem)
       List2.AddItem List1.List(CurItem)
         'Lalu hapus dari List1
         List1.RemoveItem (CurItem)
      Else
        CurItem = CurItem + 1
      End If
    Loop Until CurItem = List1.ListCount
End Sub

Private Sub Form_Load()
'Tambahkan beberapa item ke dalam ListBox
    For i = 1 To 10
        List1.AddItem "Item " & i
    Next
End Sub


VERSION 5.00
Begin VB.Form frmsembol 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�zel Karakterler"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdinsert 
      Caption         =   "Tamam"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Kapat"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   4800
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Columns         =   15
      Height          =   3960
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "��mb�l|�r"
      Height          =   5295
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3495
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bir tane �zel karakter se�in ve tamama t�klay�n."
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   4320
         Width           =   3375
      End
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      BorderWidth     =   5
      Height          =   5505
      Left            =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "frmsembol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' �
'                ������������������������������������
'          �� ������� �������������������t��������
'            ��     �ssss�   ��ccc����     �t�      ��
'           �� �   �s���� � �c�����cc�  �  �t�       ��
'          �     ����      ����    �c�     �t�
'            �  ��s�  �    �c�      ��  �  �t�      �
'   �           �s�s      �c��             �t�
'                ��ss�s   �cc�             �t�
'                  ��ss   ��c�         �   �t�
'       �    �     �s��    ��c�    ���     �t�
'                  �ss�     �c�   ����  �  �t�   �
'                ��s��   �   �c����c�      �t�      �
'    �          �ss� �       �ccc c��      �t�    �
'               s��           ������       �t�
'              ��   SOLDiER CRACKERS TEAM  ���
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'*******************************************************'
'*    proje: [SCT] H�zl� HTML Edit�r�                  *'
'*    yazar: Anafarta T�rk                             *'
'*  e-posta: blau_devil@hotmail.com                    *'
'*      web: http://www.sct.tr.cx/                     *'
'*    tarih: 30.11.2002                                *'
'*******************************************************'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Command1_Click()
Clipboard.SetText (List1)
txtWord.Text = txtWord.Text + List1
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
MsgBox "Never Ever Simpler! Just Select a symbol from list hit copy and it will automatically goto the textbox also to your clipboard so have fun!"
End Sub

Private Sub Command4_Click()
MsgBox "Sup this is a program to create cool symbols for Stacraft/Edit, made by SkaterRob from SCMaps.com a great website! This Program is Copyright �2001 SCMaps.com"
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdinsert_Click()
frmAna.textHTML.SelRTF = List1
End Sub

Private Sub Form_Load()
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"
List1.AddItem "�"

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

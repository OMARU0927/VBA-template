VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Trade�N���X
' �e�����\���N���X�ł��B
' id�Aname�Aamount�v���p�e�B�������Ă��܂��B

Private pId As Integer       ' �����ID
Private pName As String       ' ����̖��O
Private pAmount As Integer    ' ����̗�

' Id�v���p�e�B�̃Q�b�^�[�ƃZ�b�^�[
Public Property Get Id() As Integer
    Id = pId
End Property
Public Property Let Id(Value As Integer)
    pId = Value
End Property

' Name�v���p�e�B�̃Q�b�^�[�ƃZ�b�^�[
Public Property Get Name() As String
    Name = pName
End Property
Public Property Let Name(Value As String)
    pName = Value
End Property

' Amount�v���p�e�B�̃Q�b�^�[�ƃZ�b�^�[
Public Property Get Amount() As Integer
    Amount = pAmount
End Property
Public Property Let Amount(Value As Integer)
    pAmount = Value
End Property
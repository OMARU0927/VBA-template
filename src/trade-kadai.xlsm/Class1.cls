VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Company�N���X
' �e��Ђ�\���N���X�ł��B
' companyId�AcompanyName�Atrades�v���p�e�B�������Ă��܂��B
' trades�v���p�e�B��Trade�I�u�W�F�N�g�̃R���N�V�����ł��B

Private pCompanyId As Integer     ' ��Ђ�ID
Private pCompanyName As String    ' ��Ђ̖��O
Private pTrades As Collection     ' ��Ђ̎���̃R���N�V����

' CompanyId�v���p�e�B�̃Q�b�^�[�ƃZ�b�^�[
Public Property Get CompanyId() As Integer
    CompanyId = pCompanyId
End Property
Public Property Let CompanyId(Value As Integer)
    pCompanyId = Value
End Property

' CompanyName�v���p�e�B�̃Q�b�^�[�ƃZ�b�^�[
Public Property Get CompanyName() As String
    CompanyName = pCompanyName
End Property
Public Property Let CompanyName(Value As String)
    pCompanyName = Value
End Property

' Trades�v���p�e�B�̃Q�b�^�[�ƃZ�b�^�[
Public Property Get Trades() As Collection
    Set Trades = pTrades
End Property
Public Property Set Trades(Value As Collection)
    Set pTrades = Value
End Property

' Company�N���X�̏�����
' Trade�I�u�W�F�N�g�̃R���N�V������V�K�쐬���܂��B
Private Sub Class_Initialize()
    Set pTrades = New Collection
End Sub

' Trade�I�u�W�F�N�g��ǉ����郁�\�b�h
' �����ŗ^����ꂽTrade�I�u�W�F�N�g��Trades�R���N�V�����ɒǉ����܂��B
Public Sub AddTrade(objTrade As Trade)
    pTrades.Add objTrade
End Sub

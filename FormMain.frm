VERSION 5.00
Begin VB.Form FormMain 
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "MeteonCaptor"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   313
   ScaleMode       =   3  '�߸��
   ScaleWidth      =   369
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.FileListBox File1 
      Height          =   450
      Left            =   2040
      Pattern         =   "*.mcd"
      TabIndex        =   41
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "�ݒ��ۑ�"
      Height          =   735
      Left            =   1560
      TabIndex        =   35
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton cmdLoad 
         Caption         =   "�Ǎ�"
         Height          =   255
         Left            =   3120
         TabIndex        =   40
         Top             =   300
         Width           =   615
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "�ۑ�"
         Height          =   255
         Left            =   2400
         TabIndex        =   39
         Top             =   300
         Width           =   615
      End
      Begin VB.ComboBox cmbSetting 
         Height          =   300
         Left            =   120
         TabIndex        =   36
         Top             =   300
         Width           =   2175
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "�ۑ��ꏊ"
      Height          =   615
      Left            =   120
      TabIndex        =   31
      Top             =   3600
      Width           =   5295
      Begin VB.CommandButton cmdDirView 
         Caption         =   "�\��"
         Height          =   255
         Left            =   4560
         TabIndex        =   34
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "�Q��"
         Height          =   255
         Left            =   3840
         TabIndex        =   33
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtDir 
         Height          =   270
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.TextBox txtNavi 
      Alignment       =   2  '��������
      Height          =   270
      Left            =   120
      TabIndex        =   25
      Text            =   "���e�I�X������E���������["
      Top             =   4320
      Width           =   5295
   End
   Begin VB.Frame Frame5 
      Caption         =   "�t�@�C���`��"
      Height          =   1335
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   3855
      Begin VB.TextBox txtFileSize 
         Height          =   270
         Left            =   2640
         TabIndex        =   22
         Text            =   "100"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtQuality 
         Height          =   270
         Left            =   3240
         TabIndex        =   20
         Text            =   "75"
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optFile 
         Caption         =   "JPEG(�t�@�C���T�C�Y�w��)"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton optFile 
         Caption         =   "JPEG(�掿�w��)"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optFile 
         Caption         =   "PNG(�ō��掿)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '����
         Caption         =   "KB"
         Height          =   180
         Left            =   3240
         TabIndex        =   23
         Top             =   990
         Width           =   585
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '����
         Caption         =   "�掿�F"
         Height          =   180
         Left            =   2640
         TabIndex        =   21
         Top             =   630
         Width           =   1170
      End
   End
   Begin VB.PictureBox picMosaic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  '�Ȃ�
      Height          =   270
      Left            =   120
      ScaleHeight     =   18
      ScaleMode       =   3  '�߸��
      ScaleWidth      =   20
      TabIndex        =   15
      Top             =   4560
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Frame Frame4 
      Caption         =   "���U�C�N"
      Height          =   1335
      Left            =   4080
      TabIndex        =   8
      Top             =   2160
      Width           =   1335
      Begin VB.CheckBox chkMosaic2 
         Caption         =   "�O�����"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox chkMosaic 
         Caption         =   "6P"
         Height          =   255
         Index           =   5
         Left            =   720
         TabIndex        =   14
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chkMosaic 
         Caption         =   "5P"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chkMosaic 
         Caption         =   "4P"
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   12
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox chkMosaic 
         Caption         =   "3P"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox chkMosaic 
         Caption         =   "2P"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chkMosaic 
         Caption         =   "1P"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "�傫��"
      Height          =   1095
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   3855
      Begin VB.TextBox txtRate 
         Enabled         =   0   'False
         Height          =   270
         Left            =   3240
         TabIndex        =   30
         Text            =   "100"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtHeight 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1320
         TabIndex        =   6
         Text            =   "375"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtWidth 
         Height          =   270
         Left            =   1320
         TabIndex        =   7
         Text            =   "500"
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optSize 
         Caption         =   "�{���w��"
         Height          =   225
         Index           =   3
         Left            =   2040
         TabIndex        =   29
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optSize 
         Caption         =   "�����w��"
         Height          =   225
         Index           =   2
         Left            =   2040
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optSize 
         Caption         =   "�����w��"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optSize 
         Caption         =   "���w��"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�ۑ������"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      Begin VB.CheckBox chkClear 
         Caption         =   "CLEAR!"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   960
         Value           =   1  '����
         Width           =   1095
      End
      Begin VB.ComboBox cmbKey 
         Height          =   300
         ItemData        =   "FormMain.frx":12FA
         Left            =   360
         List            =   "FormMain.frx":12FC
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   38
         Top             =   1560
         Width           =   855
      End
      Begin VB.CheckBox chkKey 
         Caption         =   "�L�[����"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox chkSurvival 
         Caption         =   "����"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Value           =   1  '����
         Width           =   1095
      End
      Begin VB.CheckBox chkGameSet 
         Caption         =   "GameSet"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Value           =   1  '����
         Width           =   1095
      End
      Begin VB.CheckBox chkTimeUp 
         Caption         =   "TimeUp"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   1  '����
         Width           =   1095
      End
   End
   Begin VB.PictureBox picBuf 
      Appearance      =   0  '�ׯ�
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  '�Ȃ�
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   600
      ScaleHeight     =   480
      ScaleMode       =   3  '�߸��
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   4560
      Visible         =   0   'False
      Width           =   9600
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�E�B���h�E���𓾂�
Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As Long, ByVal lpWindowName As String) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Sub GetClientRect Lib "user32" (ByVal hwnd As Long, ByRef lpRect As RECT)
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal XPos As Long, ByVal nYPos As Long) As Long

'�`��
Private Declare Sub BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)
Private Declare Sub StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal nDestLeft As Long, ByVal nDestTop As Long, ByVal nDestWidth As Long, ByVal nDestHeight As Long, ByVal hSrcDC As Long, ByVal nSrcLeft As Long, ByVal nSrcTop As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long)
Private Declare Sub SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal iStretchMode As Long)

Private Const BLACKONWHITE = 1
Private Const WHITEONBLACK = 2
Private Const COLORONCOLOR = 3
Private Const HALFTONE = 4

'���̑����X
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Declare Sub ShellExecuteA Lib "shell32" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'�v���O�����Ǝ��̕���
Private Type CONDITION
    hwnd As Long
    hdc As Long
    Size As RECT
    Status As String
End Type

Dim LastMessage As Long

Private Sub cmbSetting_Change()
    Static Processing As Boolean
    If Not Processing Then
        If cmbSetting.SelLength = 0 Then
            Dim Txt As String
            Txt = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(cmbSetting.Text, "\", "��"), "/", "�^"), ":", "�F"), "*", "��"), "?", "�H"), """", "�W"), "<", "��"), ">", "��"), "|", "�b")
            If Txt <> cmbSetting.Text Then
                Dim Sel As Integer
                Sel = cmbSetting.SelStart
                Processing = True
                cmbSetting.Text = Txt
                cmbSetting.SelStart = Sel
                Processing = False
            End If
        End If
    End If
End Sub

Private Sub cmdDir_Click()
    Dim Directory As String
    Directory = ShowFolderDlg(hwnd, "�ۑ��ꏊ", GetPath, BIF_RETURNONLYFSDIRS Or BIF_USENEWUI)
    If Len(Directory) > 0 Then txtDir.Text = Directory
End Sub

Private Sub cmdDirView_Click()
    ShellExecuteA 0, "explore", GetPath, "", "", 1
End Sub

Private Sub cmdLoad_Click()
    If Len(cmbSetting.Text) = 0 Then Exit Sub
    If Len(Dir(cmbSetting.Text & ".mcd")) = 0 Then Exit Sub
    Dim N As Integer
    N = FreeFile
    Open cmbSetting.Text & ".mcd" For Input As #N
    Dim Buf As String
        Input #N, Buf
        If Buf = "MeteonCaptor104" Then
            Input #N, Buf: chkTimeUp.Value = Buf
            Input #N, Buf: chkGameSet.Value = Buf
            Input #N, Buf: chkSurvival.Value = Buf
            Input #N, Buf: chkKey.Value = Buf
            Input #N, Buf: cmbKey.Text = Buf
            Input #N, Buf: optSize(Buf).Value = True
            Input #N, Buf: txtWidth.Text = Buf
            Input #N, Buf: txtHeight.Text = Buf
            Input #N, Buf: txtRate.Text = Buf
            Input #N, Buf: optFile(Buf).Value = True
            Input #N, Buf: txtQuality.Text = Buf
            Input #N, Buf: txtFileSize.Text = Buf
            Input #N, Buf: chkMosaic(0).Value = Buf
            Input #N, Buf: chkMosaic(1).Value = Buf
            Input #N, Buf: chkMosaic(2).Value = Buf
            Input #N, Buf: chkMosaic(3).Value = Buf
            Input #N, Buf: chkMosaic(4).Value = Buf
            Input #N, Buf: chkMosaic(5).Value = Buf
            Input #N, Buf: chkMosaic2.Value = Buf
            Input #N, Buf: txtDir.Text = Buf
        ElseIf Buf = "MeteonCaptor106" Then
            Input #N, Buf: chkTimeUp.Value = Buf
            Input #N, Buf: chkGameSet.Value = Buf
            Input #N, Buf: chkSurvival.Value = Buf
            Input #N, Buf: chkClear.Value = Buf
            Input #N, Buf: chkKey.Value = Buf
            Input #N, Buf: cmbKey.Text = Buf
            Input #N, Buf: optSize(Buf).Value = True
            Input #N, Buf: txtWidth.Text = Buf
            Input #N, Buf: txtHeight.Text = Buf
            Input #N, Buf: txtRate.Text = Buf
            Input #N, Buf: optFile(Buf).Value = True
            Input #N, Buf: txtQuality.Text = Buf
            Input #N, Buf: txtFileSize.Text = Buf
            Input #N, Buf: chkMosaic(0).Value = Buf
            Input #N, Buf: chkMosaic(1).Value = Buf
            Input #N, Buf: chkMosaic(2).Value = Buf
            Input #N, Buf: chkMosaic(3).Value = Buf
            Input #N, Buf: chkMosaic(4).Value = Buf
            Input #N, Buf: chkMosaic(5).Value = Buf
            Input #N, Buf: chkMosaic2.Value = Buf
            Input #N, Buf: txtDir.Text = Buf
        End If
    Close N
End Sub

Private Sub cmdSave_Click()
    If Len(cmbSetting.Text) = 0 Then Exit Sub
    If Len(Dir(cmbSetting.Text & ".mcd")) <> 0 Then
        If MsgBox("�㏑���I�[�P�C�H", vbQuestion Or vbOKCancel) <> vbOK Then Exit Sub
    End If
    Dim N As Integer
    N = FreeFile
    Open cmbSetting.Text & ".mcd" For Output As #N
        Write #N, "MeteonCaptor106"
        Write #N, chkTimeUp.Value
        Write #N, chkGameSet.Value
        Write #N, chkSurvival.Value
        Write #N, chkClear.Value
        Write #N, chkKey.Value
        Write #N, cmbKey.Text
        Write #N, Switch(optSize(0).Value, 0, optSize(1).Value, 1, optSize(2).Value, 2, optSize(3).Value, 3)
        Write #N, txtWidth.Text
        Write #N, txtHeight.Text
        Write #N, txtRate.Text
        Write #N, Switch(optFile(0).Value, 0, optFile(1).Value, 1, optFile(2).Value, 2)
        Write #N, txtQuality.Text
        Write #N, txtFileSize.Text
        Write #N, chkMosaic(0).Value
        Write #N, chkMosaic(1).Value
        Write #N, chkMosaic(2).Value
        Write #N, chkMosaic(3).Value
        Write #N, chkMosaic(4).Value
        Write #N, chkMosaic(5).Value
        Write #N, chkMosaic2.Value
        Write #N, txtDir.Text
    Close N
    UpdateFileList
End Sub

Private Sub UpdateFileList()
    File1.Refresh
    cmbSetting.Clear
    Dim I As Integer
    For I = 0 To File1.ListCount - 1
        cmbSetting.AddItem Replace(File1.List(I), ".mcd", "", , , vbTextCompare)
    Next
End Sub

Private Sub Form_Load()
    Dim I As Integer
    For I = Asc("A") To Asc("Z")
        cmbKey.AddItem Chr(I)
    Next
    cmbKey.Text = "C"
    File1.Path = App.Path
    UpdateFileList
End Sub

Private Sub optSize_Click(Index As Integer)
    Select Case Index
    Case 0
        txtWidth.Enabled = True
        txtHeight.Enabled = False
        txtRate.Enabled = False
    Case 1
        txtWidth.Enabled = False
        txtHeight.Enabled = True
        txtRate.Enabled = False
    Case 2
        txtWidth.Enabled = True
        txtHeight.Enabled = True
        txtRate.Enabled = False
    Case 3
        txtWidth.Enabled = False
        txtHeight.Enabled = False
        txtRate.Enabled = True
    End Select
End Sub

Private Sub Timer1_Timer()
    Dim c As CONDITION
    Static Prohibited As Boolean
    Static Active As Boolean
    
    '�������A�N�e�B�u��������B�e���Ă��Ӗ��Ȃ��Ǝv��
    If GetForegroundWindow = hwnd Then
        If Not Active Then
            Navi "�ݒ蒆�Ȃ�ŎB�e�͂��܂���["
            Active = True
        End If
        Exit Sub
    End If
    
    '�Q�[����ʂ�T�����[
    c.hwnd = FindWindowA(0, "MeteosOnline")
    If c.hwnd = 0 Then
        Navi "���e�I���N�����ĂȂ��̂��ȁ[", True
        Exit Sub
    End If
    GetClientRect c.hwnd, c.Size
    
    '���߂ɃT�C�Y�`�F�b�N����[
    Dim wdt As Single, hgt As Single
    If optSize(0).Value Then
        wdt = Val(txtWidth.Text)
        hgt = wdt * 3# / 4#
    ElseIf optSize(1).Value Then
        hgt = Val(txtHeight.Text)
        wdt = hgt * 4# / 3#
    ElseIf optSize(2).Value Then
        wdt = Val(txtWidth.Text)
        hgt = Val(txtHeight.Text)
    ElseIf optSize(3).Value Then
        wdt = c.Size.right * Val(txtRate.Text) / 100#
        hgt = c.Size.bottom * Val(txtRate.Text) / 100#
    End If
    
    Dim Directory As String
    Directory = GetPath
    
    '��A�N�e�B�u���̏u�ԂɃG���[�`�F�b�N�I�I
    If Active Then
        If Len(Dir("imgctl.dll")) = 0 Then
            Navi "imgctl.dll���Ȃ��̂��[", True
            Exit Sub
        End If
        If chkTimeUp.Value <> 1 And chkGameSet.Value <> 1 And chkSurvival.Value <> 1 And chkClear.Value <> 1 And chkKey.Value <> 1 Then
            Navi "�ۑ��������w�肵�Ă���܂���Ɓ[", True
            Exit Sub
        End If
        If wdt > c.Size.right Then
            Navi "�����ł������ł����ȁ[", True
            Exit Sub
        End If
        If hgt > c.Size.bottom Then
            Navi "�������������ł���ȁ[", True
            Exit Sub
        End If
        If wdt < 4 Then
            Navi "���������������ł����ȁ[", True
            Exit Sub
        End If
        If hgt < 3 Then
            Navi "�������Ⴗ���ł���ȁ[", True
            Exit Sub
        End If
        If Val(txtRate.Text) < 1 Then
            Navi "�k�����������Ⴂ�܂��́[", True
            Exit Sub
        End If
        If Val(txtRate.Text) > 100 Then
            Navi "�g��͂����܂ւ���[", True
            Exit Sub
        End If
        If Val(txtQuality.Text) < 1 Then
            Navi "�掿��1�ȏ�͂Ȃ��Ɓ[", True
            Exit Sub
        End If
        If Val(txtQuality.Text) > 100 Then
            Navi "�{���ȏ�̉掿�͖�����Ł[", True
            Exit Sub
        End If
        If InStr(Directory, "*") > 0 Or InStr(Directory, "?") > 0 Then
            Navi "���C���h�J�[�h�֎~�ł���[", True
            Exit Sub
        End If
        If Len(Dir(Directory)) = 0 Then
            Navi "�ۑ��ꏊ��������܂����[", True
            Exit Sub
        End If
        Dim attr As Integer
        attr = GetAttr(Directory)
        If (attr And vbDirectory) = 0 Then
            Navi "�ۑ��ꏊ�͎��݂���t�H���_�ɂ��܂��傤�ˁ[", True
            Exit Sub
        End If
        '���܂��������Ȃ�
        'If (attr And vbReadOnly) <> 0 Then
        '    Navi "�ۑ��ꏊ�͏������݉\�ȃt�H���_�ɂ��܂��傤�ˁ[", True
        '    Exit Sub
        'End If
        Active = False
    End If
    
    '�󋵂𔻒f
    c.hdc = GetDC(c.hwnd)
    If c.hdc = 0 Then
        Navi "���܂���ʂ������܂���ˁ[", True
        GoTo FINALLY
    End If
    If CheckKey(c) Then
        c.Status = "�蓮�ۑ�"
    Else
        If CheckTimeUp(c) Then
            c.Status = "TimeUp"
        Else
            If CheckGameSet(c) Then
                c.Status = "GameSet"
            Else
                If CheckSurvival(c) Then
                    c.Status = "����"
                Else
                    If CheckClear(c) Then
                        c.Status = "CLEAR!"
                    Else
                        Prohibited = False
                        If timeGetTime - LastMessage >= 10000 Then
                            Select Case Int(Rnd * 16)
                            Case 0: Navi "���e�I�X������E���������["
                            Case 1: Navi "���e�I�X�������ăY�h�h�h�h�["
                            Case 2: Navi "���������ǂ����ŏ㏸�u���["
                            Case 3: Navi "���e�I�X������ł�������["
                            Case 4: Navi "���ł��B�e�����n�j�ł���["
                            Case 5: If chkTimeUp.Value = 1 Then Navi "���Ԃ�������p�V���I���܂���["
                            Case 6: If chkGameSet.Value = 1 Then Navi "�ŖS�������ăo�b�`���ł���["
                            Case 7: If chkSurvival.Value = 1 Then Navi "���������E�p�A�����Ă��炢�܂��["
                            Case 8: Navi "�ŏ������Ă������Ă܂�����ˁ["
                            Case 9: Navi "Live���g���ƃ��U�C�N�����������Ȃ�܂���["
                            Case 10: Navi "�����I���ȊO�͎����ŎB�e���Ăˁ["
                            Case 11: Navi "�����["
                            Case 12: Navi "�����˂����ˁI�΂��ā["
                            Case 13: Navi "����̃L��������Ɏ��������Ď��Y�ꂸ�Ɂ["
                            Case 14: Navi "�������Ԃ��������ƃG���["
                            Case 15: Navi "��F�B�e����������炲�߂�ˁ["
                            End Select
                        End If
                        GoTo FINALLY
                    End If
                End If
            End If
        End If
    End If
    If Prohibited Then GoTo FINALLY
    Prohibited = True
    
    c.Status = Directory & c.Status & Format(Now, " yy�Nmm��dd�� hh��mm��ss�b")
    
    '�摜�������Ă��ĉ��H����
    picBuf.Move 0, 0, wdt, hgt
    SetStretchBltMode picBuf.hdc, HALFTONE
    StretchBlt picBuf.hdc, 0, 0, picBuf.ScaleWidth, picBuf.ScaleHeight, c.hdc, 0, 0, c.Size.right, c.Size.bottom, vbSrcCopy
    If chkMosaic(0).Value = 1 Then Mosaic c, 52# / 640#, 76# / 480#
    If chkMosaic(1).Value = 1 Then Mosaic c, 496# / 640#, 80# / 480#
    If chkMosaic(2).Value = 1 Then Mosaic c, 48# / 640#, 179# / 480#
    If chkMosaic(3).Value = 1 Then Mosaic c, 496# / 640#, 179# / 480#
    If chkMosaic(4).Value = 1 Then Mosaic c, 48# / 640#, 278# / 480#
    If chkMosaic(5).Value = 1 Then Mosaic c, 496# / 640#, 278# / 480#
    
    '���悢��ۑ�
    Dim hDIB As Long
    hDIB = DCtoDIB(picBuf.hdc, 0, 0, picBuf.ScaleWidth, picBuf.ScaleHeight)
    Dim FileName As String
    Dim Result As Long
    If optFile(0).Value Then
        'PNG
        FileName = c.Status & ".png"
        Result = DIBtoPNG(FileName, hDIB, 0)
    ElseIf optFile(1).Value Then
        '�i���w��
        FileName = c.Status & ".jpg"
        Result = DIBtoJPG(FileName, hDIB, Val(txtQuality.Text), 0)
    ElseIf optFile(2).Value Then
        '�T�C�Y�w��
        FileName = c.Status & ".jpg"
        Dim Q As Integer
        For Q = 100 To 10 Step -10
            Result = DIBtoJPG(FileName, hDIB, Q, 0)
            If Result = 0 Then Exit For
            If FileLen(FileName) <= Val(txtFileSize.Text) * 1000 Then Exit For
        Next
    End If
    DeleteDIB hDIB
    If Result <> 0 Then
        Navi "�p�V���I�����̂��B��܂�����["
    Else
        Navi "�p�V���I�E�E�E�����A�������s�ł��[", True
    End If

FINALLY:
    ReleaseDC c.hwnd, c.hdc
End Sub

Private Function CheckKey(ByRef c As CONDITION) As Boolean
    Static Prev As Boolean
    CheckKey = False
    If chkKey.Value <> 1 Then Exit Function
    If GetAsyncKeyState(Asc(cmbKey.Text)) = 0 Then
        Prev = False
        Exit Function
    Else
        If Prev Then Exit Function
    End If
    Prev = True
    CheckKey = True
End Function

Private Function CheckTimeUp(ByRef c As CONDITION) As Boolean
    CheckTimeUp = False
    If chkTimeUp.Value <> 1 Then Exit Function
    If Not CheckColor(c, 203# / 640#, 169# / 480#, 124, 211, 0) Then Exit Function
    If Not CheckColor(c, 313# / 640#, 172# / 480#, 146, 207, 4) Then Exit Function
    If Not CheckColor(c, 228# / 640#, 210# / 480#, 13, 111, 40) Then Exit Function
    If Not CheckColor(c, 369# / 640#, 217# / 480#, 12, 147, 45) Then Exit Function
    If Not CheckColor(c, 285# / 640#, 252# / 480#, 12, 111, 40) Then Exit Function
    If Not CheckColor(c, 346# / 640#, 242# / 480#, 146, 174, 4) Then Exit Function
    CheckTimeUp = True
End Function

Private Function CheckGameSet(ByRef c As CONDITION) As Boolean
    CheckGameSet = False
    If chkGameSet.Value <> 1 Then Exit Function
    If Not CheckColor(c, 266# / 640#, 164# / 480#, 169, 0, 118) Then Exit Function
    If Not CheckColor(c, 326# / 640#, 175# / 480#, 169, 101, 203) Then Exit Function
    If Not CheckColor(c, 220# / 640#, 217# / 480#, 169, 0, 118) Then Exit Function
    If Not CheckColor(c, 373# / 640#, 217# / 480#, 152, 84, 237) Then Exit Function
    If Not CheckColor(c, 320# / 640#, 242# / 480#, 118, 33, 135) Then Exit Function
    If Not CheckColor(c, 390# / 640#, 243# / 480#, 168, 168, 168) Then Exit Function
    CheckGameSet = True
End Function

Private Function CheckSurvival(ByRef c As CONDITION) As Boolean
    CheckSurvival = False
    If chkSurvival.Value <> 1 Then Exit Function
    If Not CheckColor(c, 238# / 640#, 154# / 480#, 168, 169, 169) Then Exit Function
    If Not CheckColor(c, 403# / 640#, 147# / 480#, 254, 190, 0) Then Exit Function
    If Not CheckColor(c, 390# / 640#, 258# / 480#, 254, 190, 1) Then Exit Function
    If Not CheckColor(c, 289# / 640#, 208# / 480#, 167, 253, 245) Then Exit Function
    If Not CheckColor(c, 413# / 640#, 201# / 480#, 253, 126, 0) Then Exit Function
    If Not CheckColor(c, 368# / 640#, 232# / 480#, 242, 253, 252) Then Exit Function
    CheckSurvival = True
End Function

Private Function CheckClear(ByRef c As CONDITION) As Boolean
    CheckClear = False
    If chkClear.Value <> 1 Then Exit Function
    If Not CheckColor(c, 236# / 640#, 142# / 480#, 254, 186, 1) Then Exit Function
    If Not CheckColor(c, 326# / 640#, 163# / 480#, 169, 254, 152) Then Exit Function
    If Not CheckColor(c, 237# / 640#, 197# / 480#, 237, 254, 237) Then Exit Function
    If Not CheckColor(c, 412# / 640#, 220# / 480#, 50, 203, 151) Then Exit Function
    If Not CheckColor(c, 239# / 640#, 243# / 480#, 1, 119, 51) Then Exit Function
    If Not CheckColor(c, 371# / 640#, 277# / 480#, 254, 186, 1) Then Exit Function
    CheckClear = True
End Function

Private Function CheckColor(ByRef c As CONDITION, ByVal x As Double, ByVal y As Double, ByVal r As Long, ByVal g As Long, ByVal b As Long) As Boolean
    Dim col2 As Long, r2 As Long, g2 As Long, b2 As Long
    col2 = GetPixel(c.hdc, c.Size.right * x, c.Size.bottom * y)
    r2 = col2 And &HFF&
    g2 = col2 \ &H100& And &HFF&
    b2 = col2 \ &H10000 And &HFF&
    CheckColor = False
    If Abs(r - r2) > 10 Then Exit Function
    If Abs(g - g2) > 10 Then Exit Function
    If Abs(b - b2) > 10 Then Exit Function
    CheckColor = True
End Function

Private Sub Mosaic(ByRef c As CONDITION, ByVal x As Double, ByVal y As Double)
    Dim sw As Long, sh As Long
    sw = IIf(chkMosaic2.Value = 1, 10, 20)
    sh = IIf(chkMosaic2.Value = 1, 9, 18)
    SetStretchBltMode picMosaic.hdc, HALFTONE
    StretchBlt picMosaic.hdc, 0, 0, sw, sh, picBuf.hdc, picBuf.Width * x, picBuf.Height * y, picBuf.Width * 100# / 640#, picBuf.Height * 90# / 480#, vbSrcCopy
    SetStretchBltMode picBuf.hdc, COLORONCOLOR
    StretchBlt picBuf.hdc, picBuf.Width * x, picBuf.Height * y, picBuf.Width * 100# / 640#, picBuf.Height * 90# / 480#, picMosaic.hdc, 0, 0, sw, sh, vbSrcCopy
End Sub

Private Sub Navi(ByRef s As String, Optional e As Boolean = False)
    '�n�������m���Ɍy���Ȃ�
    If txtNavi.Text <> s Then
        txtNavi.Text = s
        txtNavi.ForeColor = IIf(e, vbRed, vbBlack)
        txtNavi.FontBold = e
        LastMessage = timeGetTime
    End If
End Sub

Private Function GetPath() As String
    GetPath = IIf(InStr(txtDir.Text, ":") = 0, App.Path & "\", "")
    GetPath = Replace(Replace(GetPath & txtDir.Text & "\", "\\", "\"), "\\", "\")
    '�����[�g�Ƀ\�t�g�{�̂�u���Ă���"\\\"���o�Ă���̂ł��̑΍�
End Function

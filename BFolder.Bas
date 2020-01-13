Attribute VB_Name = "BrowseForFolders"
' =====================================
' BrowseForFolders(BFolder.Bas)
'                       Niiyama(HEROPA)
'                      heropa@dream.com
' =====================================
Option Explicit

' ���W���[���O���Ɍ��J�����B��̃��\�b�h
' ShowFolderDlg���g����API�ɃA�N�Z�X���܂��
' �N���X���W���[���Ƃ��Ď��������������񂾂��ǁA
' �R�[���o�b�N���g���̂ŕW�����W���[����
' �ǂ����Ă��K�v�������́B(^_^;)
' �ǂ����Ă��N���X������Ȃ�A�N���X���W���[������
' ���̃��W���[�����ĂԂ悤�ɂ���Ηǂ����H

' BROWSEINFO�\����
Private Type tagBROWSEINFO
    hOwner          As Long
    pidlRoot        As Long
    pszDisplayName  As String
    lpszTitle       As String
    ulFlags         As Long
    lpfn            As Long
    lParam          As Long
    iImage          As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
                                (lpBrowseInfo As tagBROWSEINFO) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
                                (ByVal pidl As Long, _
                                 ByVal pszPath As String) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" _
                                (ByVal pv As Long)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                                (ByVal hwnd As Long, _
                                 ByVal wMsg As Long, _
                                 ByVal wParam As Long, _
                                 lParam As Any) As Long

Private Const MAX_PATH = 260
Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1

Private Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
Private Const BFFM_ENABLEOK As Long = (WM_USER + 101)
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)

Private Declare Function LocalAlloc Lib "kernel32" _
                                (ByVal uFlags As Long, _
                                 ByVal uBytes As Long) As Long

Private Declare Function LocalFree Lib "kernel32" _
                                (ByVal hMem As Long) As Long

Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)


Public Enum BrowseInfoFlags
    ' �f�B���N�g���ȊO�͑I���ł��Ȃ��悤�ɂ���B
    ' �f�B���N�g���ł͂Ȃ��t�H���_��I�����Ă��A
    ' OK�{�^�����L���ɂȂ�Ȃ��B
    BIF_RETURNONLYFSDIRS = &H1&
    ' �l�b�g���[�N�h���C�u�̑I���͂ł��Ȃ��悤�ɂ���B
    BIF_DONTGOBELOWDOMAIN = &H2&
    BIF_STATUSTEXT = &H4&
    BIF_RETURNFSANCESTORS = &H8&
    ' �R���s���[�^�̑I�������ł��Ȃ��悤�ɂ���B
    BIF_BROWSEFORCOMPUTER = &H1000&
    ' �v�����^�̑I�������ł��Ȃ��悤�ɂ���B
    BIF_BROWSEFORPRINTER = &H2000&
    ' �t�@�C�����I���ł���悤�ɂ���B
    BIF_BROWSEINCLUDEFILES = &H4000&        ' Version 4.71
    BIF_BROWSEINCLUDEURLS = &H80&           ' Version 5.0
    ' �G�f�B�b�g�{�b�N�X��\������B
    BIF_EDITBOX = &H10&                     ' Version 4.71
    BIF_NEWDIALOGSTYLE = &H40&              ' Version 5.0
    BIF_SHAREABLE = &H8000&                 ' Version 5.0
    BIF_USENEWUI = &H40& Or &H10&           ' Version 5.0
    BIF_VALIDATE = &H20&                    ' Version 4.71
End Enum

Private mstrDefaultPath As String
'
' �R�[���o�b�N���ꂽ���b�Z�[�W�ɂ���ă��b�Z�[�W�𑗂�B
'
Private Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    Select Case uMsg
        Case BFFM_INITIALIZED
            ' �I�������t�H���_��ݒ肷��B
            Call SendMessage(hwnd, _
                             BFFM_SETSELECTIONA, _
                             1, _
                             ByVal mstrDefaultPath)
        Case Else
    End Select
End Function
'
'
'
Private Function FARPROC(pfn As Long) As Long
    ' AddressOf�͕W�����W���[���̃v���V�[�W�����w�肵�Ȃ���΂Ȃ�Ȃ��̂ŁA
    ' �_�~�[�̃v���V�[�W������������B
    FARPROC = pfn
End Function
'
' �t�H���_�̑I���_�C�A���O��\������B
'
Public Function ShowFolderDlg(ByVal hOwner As Long, ByVal strPrompt As String, ByVal strDefaultPath As String, ByVal lngFlag As BrowseInfoFlags) As String
' �y�����z
'   hOwner          : �I�[�i�[�ƂȂ�E�B���h�E�n���h��
'   strPrompt       : �_�C�A���O���ɕ\������镶����
'   strDefaultPath  : �f�t�H���g�ƂȂ�t���p�X
'   lngFlag         : BROWSEINFO�\���̂�ulFlag

    Dim bi        As tagBROWSEINFO
    Dim pidl      As Long
    Dim lpSelPath As Long
    Dim sPath     As String * MAX_PATH

    mstrDefaultPath = strDefaultPath
    With bi
        .hOwner = hOwner
        ' �����̒l��0����Ȃ��A�����Ƃ���pidl�������
        ' ���[�g�ƂȂ�t�H���_��ύX���邱�Ƃ��ł���B
        ' �f�t�H���g��0�Ƃ����̂̓f�X�N�g�b�v�B
        .pidlRoot = 0
        ' �_�C�A���O�ɕ\������v�����v�g���Z�b�g����B
        .lpszTitle = strPrompt
        ' �t���O�l���Z�b�g����B
        .ulFlags = lngFlag
        ' �R�[���o�b�N�Ɏg�p���郁�\�b�h���Z�b�g����B
        .lpfn = FARPROC(AddressOf BrowseCallbackProc)
    End With
    ' �_�C�A���O��\������B
    pidl = SHBrowseForFolder(bi)

    If pidl Then
        ' pidl����p�X���ɕϊ�����B
        If SHGetPathFromIDList(pidl, sPath) Then
            ' �p�X�������o���B
            ShowFolderDlg = left$(sPath, InStr(sPath, vbNullChar) - 1)
        End If
        ' pidl������������J������B
        Call CoTaskMemFree(pidl)
    End If
    ' ������̃|�C���^���J������B
    Call LocalFree(lpSelPath)
End Function

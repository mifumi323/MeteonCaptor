Attribute VB_Name = "BrowseForFolders"
' =====================================
' BrowseForFolders(BFolder.Bas)
'                       Niiyama(HEROPA)
'                      heropa@dream.com
' =====================================
Option Explicit

' モジュール外部に公開される唯一のメソッド
' ShowFolderDlgを使ってAPIにアクセスします｡
' クラスモジュールとして実装したかったんだけど、
' コールバックを使うので標準モジュールが
' どうしても必要だったの。(^_^;)
' どうしてもクラス化するなら、クラスモジュールから
' このモジュールを呼ぶようにすれば良いか？

' BROWSEINFO構造体
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
    ' ディレクトリ以外は選択できないようにする。
    ' ディレクトリではないフォルダを選択しても、
    ' OKボタンが有効にならない。
    BIF_RETURNONLYFSDIRS = &H1&
    ' ネットワークドライブの選択はできないようにする。
    BIF_DONTGOBELOWDOMAIN = &H2&
    BIF_STATUSTEXT = &H4&
    BIF_RETURNFSANCESTORS = &H8&
    ' コンピュータの選択しかできないようにする。
    BIF_BROWSEFORCOMPUTER = &H1000&
    ' プリンタの選択しかできないようにする。
    BIF_BROWSEFORPRINTER = &H2000&
    ' ファイルも選択できるようにする。
    BIF_BROWSEINCLUDEFILES = &H4000&        ' Version 4.71
    BIF_BROWSEINCLUDEURLS = &H80&           ' Version 5.0
    ' エディットボックスを表示する。
    BIF_EDITBOX = &H10&                     ' Version 4.71
    BIF_NEWDIALOGSTYLE = &H40&              ' Version 5.0
    BIF_SHAREABLE = &H8000&                 ' Version 5.0
    BIF_USENEWUI = &H40& Or &H10&           ' Version 5.0
    BIF_VALIDATE = &H20&                    ' Version 4.71
End Enum

Private mstrDefaultPath As String
'
' コールバックされたメッセージによってメッセージを送る。
'
Private Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    Select Case uMsg
        Case BFFM_INITIALIZED
            ' 選択されるフォルダを設定する。
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
    ' AddressOfは標準モジュールのプロシージャを指定しなければならないので、
    ' ダミーのプロシージャを実装する。
    FARPROC = pfn
End Function
'
' フォルダの選択ダイアログを表示する。
'
Public Function ShowFolderDlg(ByVal hOwner As Long, ByVal strPrompt As String, ByVal strDefaultPath As String, ByVal lngFlag As BrowseInfoFlags) As String
' 【引数】
'   hOwner          : オーナーとなるウィンドウハンドル
'   strPrompt       : ダイアログ内に表示される文字列
'   strDefaultPath  : デフォルトとなるフルパス
'   lngFlag         : BROWSEINFO構造体のulFlag

    Dim bi        As tagBROWSEINFO
    Dim pidl      As Long
    Dim lpSelPath As Long
    Dim sPath     As String * MAX_PATH

    mstrDefaultPath = strDefaultPath
    With bi
        .hOwner = hOwner
        ' ここの値を0じゃなく、ちゃんとしたpidlを入れると
        ' ルートとなるフォルダを変更することができる。
        ' デフォルトの0というのはデスクトップ。
        .pidlRoot = 0
        ' ダイアログに表示するプロンプトをセットする。
        .lpszTitle = strPrompt
        ' フラグ値をセットする。
        .ulFlags = lngFlag
        ' コールバックに使用するメソッドをセットする。
        .lpfn = FARPROC(AddressOf BrowseCallbackProc)
    End With
    ' ダイアログを表示する。
    pidl = SHBrowseForFolder(bi)

    If pidl Then
        ' pidlからパス名に変換する。
        If SHGetPathFromIDList(pidl, sPath) Then
            ' パス名を取り出す。
            ShowFolderDlg = left$(sPath, InStr(sPath, vbNullChar) - 1)
        End If
        ' pidlをメモリから開放する。
        Call CoTaskMemFree(pidl)
    End If
    ' 文字列のポインタを開放する。
    Call LocalFree(lpSelPath)
End Function


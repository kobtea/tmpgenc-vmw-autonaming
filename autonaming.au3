#cs --------------------------------------------------------------------------------
AutoIt Version: 3.0
Author:         kobtea
#ce --------------------------------------------------------------------------------

#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#Include <GuiEdit.au3>
#include <MsgBoxConstants.au3>
#include <WindowsConstants.au3>

Local Const $TMPGENC_EXE        = "TMPGEncVMW5.exe"
Local Const $TMPGENC_FULL_PATH  = "C:\Program Files (x86)\Pegasys Inc\TMPGEnc Video Mastering Works 5\" & $TMPGENC_EXE
Local Const $WINDOW_START       = "[CLASS:TStart_MainForm]"
Local Const $WINDOW_MAIN        = "[CLASS:TMainForm]"
Local Const $WINDOW_OPEN        = "[CLASS:#32770]"
Local Const $WINDOW_MESSAGE     = "[CLASS:TMessageForm]"
Local Const $WINDOW_EDIT        = "[CLASS:TClipEditForm]"
Local Const $BUTTON_NEW_PROJECT = "[CLASS:TRLStyleButton; INSTANCE:2]"
Local Const $BUTTON_ADD_FILE    = "[CLASS:TRLStyleButton; INSTANCE:16]"
Local Const $BUTTON_OPEN        = "[CLASS:Button; INSTANCE:1]"
Local Const $BUTTON_OK          = "[CLASS:TRLStyleButton; INSTANCE:27]"
Local Const $FORM_PATH          = "[CLASS:ToolbarWindow32; INSTANCE:2]"
Local Const $FORM_FILE          = "[CLASS:Edit; INSTANCE:1]"
Local Const $FORM_CLIP_NAME     = "[CLASS:TRLEdit_InplaceEdit; INSTANCE:4]"


; GUI
; --------------------------------------------------------------------------------
GUICreate("Auto Naming Tool For TMPGEnc VMW", 600, 800, -1, -1, -1, $WS_EX_ACCEPTFILES)
; drop area to put ts files
GUICtrlCreateGroup("Put ts files here!", 5, 5, 590, 170)
Local $hDrop = GUICtrlCreateInput("", 10, 20, 580, 150, $ES_AUTOHSCROLL + $WS_DISABLED)
GUICtrlSetState(-1, $GUI_DROPACCEPTED)
; logging area
GUICtrlCreateGroup("Log", 5, 180, 590, 615)
Local $hText = GUICtrlCreateEdit('', 10, 195, 580, 595)
GUISetState(@SW_SHOW)

; main
While 1
    Switch GUIGetMsg()
        Case $GUI_EVENT_CLOSE
            Exit
        Case $GUI_EVENT_DROPPED
            Local $aFile = StringSplit(GUICtrlRead($hDrop), '|', 2)

            ; validate filetype expect ts
            For $sFile In $aFile
                If Not StringRegExp($sFile, '.+\.ts$') Then
                    MsgBox($MB_OKCANCEL, 'File Type Exception', 'This file is not ts file: ' & @CRLF & $sFile)
                    Exit
                EndIf
            Next

            ; ask to execute
            If MsgBox(True, '', UBound($aFile) & ' files will register to a project') == $IDOK Then
                RegistFiles($aFile)
            EndIf
    EndSwitch
WEnd


; Name:         RegistFiles
; Parameters:   $aFile - Array of files
; Return:       <None>
; --------------------------------------------------------------------------------
Func RegistFiles($aFile)
    OpenNewProject()

    For $sFile In $aFile
        Local $sTitle = AddFile($sFile)

        ; ...
        ; choose a title yourself
        ; ...

        Local $bIsReadable = True
        Do
            ; open a broken or unreadable file
            If WinActive($WINDOW_MESSAGE) Then
                MsgBox(0, 'Unreadable File Exception', $sFile)
                $bIsReadable = False
                ExitLoop
            EndIf
        Until WinActive($WINDOW_EDIT)

        If $bIsReadable Then
            ; set a clip name
            ControlSetText($WINDOW_EDIT, '', $FORM_CLIP_NAME, $sTitle)
            ControlClick($WINDOW_EDIT, '', $BUTTON_OK)
            Logging($sFile & @CRLF & '[o] ' & $sTitle & @CRLF)
        Else
            Logging($sFile & @CRLF & '[x] ' & $sTitle & @CRLF)
        EndIf
    Next

    MsgBox(0, '', 'Finished')
EndFunc


; Name:         OpenNewProject
; Parameters:   <None>
; Return:       <None>
; --------------------------------------------------------------------------------
Func OpenNewProject()
    ; open TMPGEnc VMW
    If Not WinExists($WINDOW_START) Then
        Run($TMPGENC_FULL_PATH)
        WinWaitActive($WINDOW_START)
    EndIf
    WinActivate($WINDOW_START)

    ; open new project
    ControlClick($WINDOW_START, '', $BUTTON_NEW_PROJECT)
    WinWaitActive($WINDOW_MAIN)
EndFunc


; Name:         AddFile
; Parameters:   $sFile - String of a file name (full path)
; Return:       String of a title (Removed from a file name to '.ts')
; --------------------------------------------------------------------------------
Func AddFile($sFile)
    ; open an add file window
    WinWaitActive($WINDOW_MAIN)
    ControlClick($WINDOW_MAIN, '', $BUTTON_ADD_FILE)

    ; split a full path to its directory and its file name
    Local $aRes = StringRegExp($sFile, '(.+)\\(.+)$', 1)
    If UBound($aRes) <> 2 Then
        MsgBox(0, 'File Path Parse Exception', 'Failed to parse a file path: ' & @CRLF & $sFile)
        Exit
    EndIf
    Local $sDir = $aRes[0]
    Local $sFileName = $aRes[1]

    ; open a file
    WinWaitActive($WINDOW_OPEN)
    ControlClick($WINDOW_OPEN, '', $FORM_PATH)
    ControlSetText($WINDOW_OPEN, '', $FORM_PATH, $sDir)
    Send($sDir & '{ENTER}{ENTER}')

    ; wait to open the expect directory
    Do
        Sleep(1)
    Until StringInStr(ControlGetText($WINDOW_OPEN, '', $FORM_PATH), $sDir)
    ControlSetText($WINDOW_OPEN, '', $FORM_FILE, $sFileName)
    ControlClick($WINDOW_OPEN, '', $BUTTON_OPEN)

    ; remove '.ts' strings
    Local $aFileName = StringSplit($sFileName, '.')
    If $aFileName[0] <> 2 Then
        MsgBox(0, 'Title Parse Exception', 'Failed to split a file name: ' & @CRLF & $sFileName)
    EndIf
    Return $aFileName[1]
EndFunc


; Name:         Logging
; Parameters:   $sMsg - String to log
; Return:       <None>
; --------------------------------------------------------------------------------
Func Logging($sMsg)
    _GUICtrlEdit_AppendText($hText, $sMsg & @CRLF)
EndFunc

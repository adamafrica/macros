#NoEnv  ; Avoid check empty variables to see if they are environment variables. Recommended.
#SingleInstance Force ; Determines whether a script is allowed to run again when it is already running.
SendMode Input  ; Makes Send synonymous with SendInput or SendPlay. Recommended.
SetWorkingDir %A_ScriptDir%  ; Changes the script's working directory.
; #Warn  ; Enable/Disable warnings for specific conditions which may indicate an error. Recommended.


; ----------------------------------------------------------------------
; Summary of AutoHotkey scripts in this file:
; ----------------------------------------------------------------------
; ctrl + alt + , : Displays a message box. Use to verify AutoHotkey is working.
;
; ctrl + w  :Closes the command console (including Powershell consoles).
;
; alt + z  : Print current datetime stamp in format: 20180227 12:51:50
;
; Windows Logo Key + [ :Write an Atlassian Noformat block
;
; Windows Logo Key + ] :Write an Atlassian SQL block
;
; ctrl + shift + { : Write a select statement and prepare for editing.
;
; ctrl + alt + t :Create a 1x1 table in OneNote.
;
; ctrl + alt + numpad1 : Create a (source) tag - a hyperlink where display text is (source) in OneNote from a URL in Chrome.
;
; ctrl + alt + numpad2 : Inserts a copy of the references table from Template - References.
;
; ctrl + alt + numpad3 : Formats a code snippet. In OneNote, changes the font to Courier New 10 pt and then reverts to font and font size used before change.
;
; ctrl + alt + numpad4 : Create a hyperlink in OneNote from the URI in the Chrome Omnibox.


; ----------------------------------------------------------------------
; - Common hotkeys:
; ----------------------------------------------------------------------
; # = windows logo key
; ! = alt key
; ^ = ctrl key
; + = shift key
; & = An ampersand may be used between two keys or mouse buttons to combine them into a hotkey.

; ----------------------------------------------------------------------
; Command Window Helpers
; ----------------------------------------------------------------------
; - Close the Command console (including PowerShell) when shortcut activated.
; - Useful when you want to quickly close a console Window.
; - HotKey: ctrl + w
; - Example output: N/A
; - Comments:
; - This shortcut replicates the default Windows behavior of ctrl + w for the
;   Windows Command console, PowerShell included.
;
; Prevent shortcut from firing unless a Command console or PowerShell console is active.
#If WinActive("ahk_class ConsoleWindowClass")
    ^w::
    WinGetTitle sTitle
    If (InStr(sTitle, "-")=0) {
            SendInput EXIT{Enter}
    } else {
            SendInput ^w
    }
    return
#If

; ----------------------------------------------------------------------
; Jira helpers
; ----------------------------------------------------------------------
; - Write an Atlassian Noformat block
; - Useful for adding unformated text to Jira tickets.
; - HotKey: Windows Logo Key + [
; - Comments: N/A
#[::
SendInput {{}noformat{}}\r\n{{}noformat{}}{left 11}
return

; - Write an Atlassian SQL block
; - Useful for adding SQL code blocks to Jira tickets.
; - HotKey: Windows Logo Key + ]
; - Comments: N/A
#]::
SendInput {{}code:language=sql|title=Title{}}\r\n{{}code{}}{left 7}
return

; ---------------------------------------------------------------------
; - OneNote Helpers
; ---------------------------------------------------------------------

; - Create a 1x1 table in OneNote.
; - Useful when you want to create table in OneNote quickly.
; - Hotkey: ctrl + alt + t
; - Example Output: A new 1x1 table created in OneNote.
; - Comments:
;   You don't need this script to create a 1x1 table in OneNote, you can also
;   use alt + n + t to create a table instead.
;   alt accesses the OneNote menu bar, n accesses the insert menu. t prepares a table for insert and {enter} inserts a 1x1 table.
#If WinActive("ahk_exe ONENOTE.EXE")
    ^!t::
    SendInput !nt{enter}
    return
#If

/*
    Description: Create a (source) tag - a hyperlink where display text is (source) in OneNote from a URL in Chrome.

    HotKey: ctrl + alt + numpad1

    Usage:
    1. In OneNote, place the cursor where you want the link to be inserted.
    2. Switch to Chrome
    3. Execute the macro using the following Hot Key: ctrl + alt + numpad1
    4. URL will be copied from Chrome and a new link will inserted into OneNote at the cursors current location.

    Comments:
    Used example here to get started: https://www.howtogeek.com/howto/23884/create-your-own-insert-hyperlink-feature-in-any-app-with-autohotkey/.

    Assumptions:

    History
    20191004        A   - Initial Version
    20211102        A   - Added stacked shortcut (^!F!) for use with 75% keyboard, which has no numpad.
*/
^!Numpad1::
^!F1::
{
    ; MsgBox,, Debug, Started ; Uncomment for troubleshooting only.

    if WinActive("ahk_exe chrome.exe") ; Chrome is driver for this so just exit if it's running.
    {
        clipboard := "" ; Empty the clipboard in preparation for copying.

        ; Move focus to the address bar so URL can be copied. alt + d. This is chrome specific.
        SendInput !d

        ; give time for omnibox (URL) to be selected
        sleep, 100

        ; Copy the URL to clipboard
        SendInput ^c

        ; Wait for 2 seconds for the clipboard to contain text. Exit script if no text found.
        ClipWait 2
        if ErrorLevel
        {
            MsgBox,, Error, The attempt to copy text to the clipboard failed.
            return
        }

        clipboard := clipboard ; Copy clipboard contents back to clipboard as text.

        ; Uncomment for troubleshooting only.
        ; MsgBox,, Debug, Control-C copied the following contents to the clipboard:`n`n%clipboard%

        ; Return to OneNote from the browser.
        if WinExist("ahk_exe ONENOTE.EXE")
        {
            ; MsgBox, OneNote is open. ; Uncomment next line for troubleshooting only.
            WinActivate, ahk_exe ONENOTE.EXE
        }
        else
        {
            MsgBox,, Error, OneNote does not appear to be open. Open it and try again.
            return
         }

        ; - Only create office style hyperlink if OneNote is active.
        ; - because this style of hyperlink is specific to Windows Office products.
        if WinActive("ahk_exe ONENOTE.EXE")
        {
            SendInput (source){left 1}^{LEFT}^+{RIGHT}
            SendInput ^k ; open link diaglog]
            SendInput ^v ; paste the hyperlink
            SendInput {enter} ; complete creation of hyperlink.
            SendInput {right 2} ; So cursor is in good position for typing.
        }
    }
    else
    {
        MsgBox,, Error, Aborting. Chrome needs to be the active window.
    }
    return
}

/*
    Description: Create a hyperlink in OneNote from the URI in the Chrome Omnibox.

    HotKey: ctrl + alt + numpad4

    Usage:
    1. In OneNote, place the cursor where you want the link to be inserted.
    2. Switch to Chrome
    3. Execute the macro using the following Hot Key: ctrl + alt + numpad4
    4. URI will be copied from Chrome and a new link will inserted into OneNote at the cursors current location.

    Comments:
    Created this macro to help speed up the process of annotating references in OneNote references sections.
    This macro is a simplified version: "Create a (source) tag - a hyperlink where display text is (source) in
    OneNote from a URL in Chrome."

    Assumptions:

    History
    20191120        A   - Initial Version
    20211102        A   - Added stacked shortcut (^!F!) for use with 75% keyboard, which has no numpad.
*/
^!Numpad4::
^!F4::
{
    ; MsgBox,, Debug, Started ; Uncomment for troubleshooting only.

    if WinActive("ahk_exe chrome.exe") ; Chrome is driver for this so just exit if it's not running.
    {
        ; Move focus to the address bar so URL can be copied. alt + d. This is chrome specific.
        SendInput !d
        ; give time for omnibox (URL) to be selected
        sleep, 100

        ; Copy the URL to clipboard
        SendInput ^c

        ; Wait for 2 seconds for the clipboard to contain text. Exit script if no text found.
        clipboard := "" ; Empty the clipboard in preparation for copying. A must have step in order for copy to work consistently.
        ClipWait 2
        if ErrorLevel
        {
            MsgBox,, Error, The attempt to copy text to the clipboard failed.
            return
        }

        clipboard := clipboard ; Copy clipboard contents back to clipboard as text.

        ; Uncomment for troubleshooting only.
        ; MsgBox,, Debug, Control-C copied the following contents to the clipboard:`n`n%clipboard%

        ; Return to OneNote from the browser.
        if WinExist("ahk_exe ONENOTE.EXE")
        {
            ; MsgBox, OneNote is open. ; Uncomment next line for troubleshooting only.
            WinActivate, ahk_exe ONENOTE.EXE
        }
        else
        {
            MsgBox,, Error, OneNote does not appear to be open. Open it and try again.
            return
         }
        ; - Only create office style hyperlink if OneNote is active.
        ; - because this style of hyperlink is specific to Windows Office products.
        if WinActive("ahk_exe ONENOTE.EXE")
        {
            SendInput ^v ; paste the hyperlink
        }
    }
    else
    {
        MsgBox,, Error, Aborting. Chrome needs to be the active window.
    }
    return
}

/*
    Description: Inserts a copy of the refences table from Template - References.

    HotKey: ctrl + alt + numpad2

    Usage:
    1. In OneNote, place the cursor where you want references table to be inserted.
    2. Execute this macro using the following Hot Key: ctrl + alt + numpad2
    3. A copy of the references table will be inserted at the current location of the cursor.

    Comments:
    Created this macro because I was copying the references table to most, but not all pages, especially after I
    started using zotero.

    Assumptions:
    1. A OneNote page titled "Template - Footnote Table" exists.
    2. The "Template - Footnote Table" page contains a single table References.

    History
    20191106        A   - Initial Version
    20211102        A   - Added stacked shortcut (^!F!) for use with 75% keyboard, which has no numpad.
*/
^!Numpad2::
^!F2::
{

    if WinActive("ahk_exe ONENOTE.EXE") ; Only run this script if OneNote is the active program.
    {
        ; Set focus to the global search dialog. (like ctrl+f but searches all notebooks.)
        SendInput ^e

        ; add search text to search dialog.
        SendInput Template - Footnote Table

        ; give time for text to be entered into the search dialog.
        sleep, 100

        ; Send enter key to execute the search.
        SendInput {enter}

        ; Give time for the search to complete.
        sleep, 100

        ; Send copy command, i.e., ctrl+a twice so that references table is copied to the clipboard.
        SendInput ^{a 2}

        ; Copy the references table.
        SendInput ^c

        ; send alt + backarrow to navigate from the sending page.
        SendInput !{left}

        ; paste the reference table.
        SendInput ^v
    }
    else
    {
        MsgBox,, Error, Aborting. Could not add refences table.
    }
    return
}

/*
    Description: Formats a code snippet. In OneNote, changes the font to Courier New 10 pt and then reverts to font and font size used before change.

    HotKey: ctrl + alt + numpad3

    Usage:
    1. In OneNote, place the cursor in the table cell where the code snippet is located.
    2. Execute the macro using the following Hot Key: ctrl + alt + numpad3
    3. The code snippet will be formatted and font and font size reverted to state before code snippet formatted.

    Comments:
    Created this macro because I was tired of formatting code snippets in OneNote.

    Assumptions:
    1. In OneNote
    2. A one-cell table has already been created.
    3. Code snippet is present in the cell.
    4. It's approprioate to format all contents of cell as a code snippet.
    5. Courier New 10 is the desired font and font-size for the code snippet.
    6. Font family and font size should be reset to the same state as before the macro was executed.

    History
    20191112        A   - Initial Version
    20211102        A   - Added stacked shortcut (^!F!) for use with 75% keyboard, which has no numpad.
*/
^!Numpad3::
^!F3::
{
    if WinActive("ahk_exe ONENOTE.EXE") ; Only format code snippet in OneNote.
    {

        ; These variables can be changed if a different code snippet appearance is desired.
        newFontFamily = Courier New ; Font Family to use to format snippet.
        newFontSize = 10  ; Font size to use to format snippet.

        ; Select all text in the active cell of table so then entire code snippet gets formatted.
        SendInput ^a

        ; Activate menu (!) > activate the home menu (h) > activate the font family menu (ff)
        SendInput !hff

        ; Capture the current font family so we can revert to it later.
        clipboard := ; clear the clipboard. ClipWait won't work as expected if there is already somethign on the clipboard, I think.
        SendInput ^c
        ClipWait
        oldFontFamily := Clipboard

        ; Change the font.
        SendInput %newFontFamily%

        ; Move to the font size box.
        SendInput {Tab}

        ; Capture the current font size so we can rever to it later.
        clipboard := ; clear the clipboard so, ClipWait works as expected.
        SendInput ^c
        ClipWait
        oldFontSize := Clipboard

        ; Change to the new font size.
        SendInput %newFontSize%

        ; Affect the change to the font size.
        SendInput {Enter}

        ; Deselect the code snippet so we can change back to the default font family and font size.
        ; If you don't deselect the text, the font family and size will just be reverted to the original.
        SendInput {Right}

        ; Activate menu (!) > activate the home menu (h) > activate the font family menu (ff)
        SendInput !hff

        SendInput %oldFontFamily%

        ; Move to the font size box.
        SendInput {Tab}

        ; Change to the new font size.
        SendInput %oldFontSize%

        ; Send enter so the change to font size is made.
        SendInput {Enter}
    }
    else
    {
        MsgBox,, Error, Aborting. This AutoHotkey macro only works in OneNote.
    }
    return
}

/*
    Description: Pastes text only format of clipboard content. Works in OneNote only.

    HotKey: ctrl + shift + v

    Usage:
    1. Copy content to clipboard.
    2. In OneNote, left-click where you want text pasted.
    2. Execute Hotkey: ctrl+shift+v

    Comments:
    In OneNote, alt+hvt will do the same thing as this macro, ctrl+shift+v seems
    more intuitive and was easier for me to remember.

    Assumptions:
    1. Text has been copied to clipboard.
    2. Location where text is to be pasted is selected.

    History
    20200518        A   - Initial Version
*/
#IfWinActive ahk_exe ONENOTE.EXE
^+v::
SendInput !hvt
Return
#IfWinActive

; ---------------------------------------------------------------------
; - SQL helpers
; ---------------------------------------------------------------------
; - Write a select statement and prepare for editing.
; - Speeds up writing a common select statement.
; - HotKey: ctrl + shift + .
^!.::
SendInput select top 10 * {enter}from{enter};{up 1}{right 3}{space 1}
return

; ---------------------------------------------------------------------
; - TimeStamp Helpers
; ---------------------------------------------------------------------
; - Prints datetime stamp.
; - Useful when writing comments, documenting data explorations, etc.
; - HotKey: is alt + z
; - Example output:
; - 20180227 12:51:50:
; - Comments:
; - Note that a semi-colon and space are appended to the end of the datetime string to make it more useful.
!z::
FormatTime, TimeString,, yyyyMMdd HH:mm:ss
SendInput -- %TimeString%:{SPACE}
return

; ---------------------------------------------------------------------
; - Troubleshooting Helpers
; ---------------------------------------------------------------------
; - Displays a message box.
; - Useful when attempting to verify to verify basic AutoHotkey functionality.
; - HotKey: ctrl + alt + ,
; - Example output:
; - A message box with "AutoHotkey test!" should appear.
; - Comments:
; - Note that a semi-colon and space are appended to the end of the datetime string to make it more useful.
^!,::
MsgBox, "AutoHotkey  test!"
return


/*
    Description: Display a message box with commonly AutoHotKey macros and their shortcut key combos.

    HotKey: ctrl + alt + esc

    Usage:
    1. Use shortcute [ctrl + alt + esc] to display the message box.

    Comments:
    It's easy to forget AutoHotKey shortcuts - and even the macros themselves - so this this
    message box displays the ones I commonly forget.

    Assumptions:
    It is not convienient to list all the shortcuts - adding them to the message box is tedious - so,
    the message box will only display those that are useful, but might not have easy to remember
    shortcuts.

    History
    20211102        A   - Initial Version
*/
^!ESC::
{
    help := []
    help.push("ctrl + alt + numpad1 : Create a (source) tag - a hyperlink where display text is (source) in OneNote from a URL in Chrome.")
    help.push("ctrl + alt + numpad2 : Inserts a copy of the references table from Template - References.")
    help.push("ctrl + alt + numpad3 : Formats a code snippet. In OneNote, changes the font to Courier New 10 pt and then reverts to font and font size used before change.")
    help.push("ctrl + alt + numpad4 : Create a hyperlink in OneNote from the URI in the Chrome Omnibox.")
    help.push("ctrl + alt + r : Reload the AutoHotKey script.")
    helpString := ""

    Loop, % help.MaxIndex()
    {
        helpString .= help[A_Index]"`r`n`r`n"
    }

    MsgBox % helpString

    return
}

/*
    Description: Reload the AutoHotKey script: C:\Program Files\AutoHotkey\AutoHotkeyU64.ahk.
    This script can be used in place of right-clicking the AutoHotKey icon in the tray and
    selecting the Reload this script option.

    HotKey: ctrl + alt + r

    Usage:
    1. Complete work on AutoHotKeyU64.ahk script.
    2. Save work.
    3. Use this macro to reload the script.

    Comments:
        Source: https://www.autohotkey.com/docs/commands/Reload.htm

    Assumptions:

    History
    20211102        A   - Initial Version
*/
^!r::
{
    Reload
    Sleep 1000 ; If successful, the reload will close this instance during the sleep, so the line below will never be reached.
    MsgBox, 4,, The script could not be reloaded. Would you like to open it for editing?
    IfMsgBox, Yes, Edit
    return
}
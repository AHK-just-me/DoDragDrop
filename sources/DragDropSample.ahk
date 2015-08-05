;-- Based on http://ahkscript.org/boards/viewtopic.php?f=5&t=8700 by jballi
;-- Modified by just me
#NoEnv
#SingleInstance, Force
ListLines, Off
SetBatchLines, -1
;-- DragDrop constants -------------------------------------------------------------------------------------------------------------
;-- DragDrop flags
DRAGDROP_S_DROP   := 0x40100 ; 262400
DRAGDROP_S_CANCEL := 0x40101 ; 262401
;-- DROPEFFECT flags
DROPEFFECT_NONE := 0 ;-- Drop target cannot accept the data.
DROPEFFECT_COPY := 1 ;-- Drop results in a copy. The original data is untouched by the drag source.
DROPEFFECT_MOVE := 2 ;-- Drag source should remove the data.
DROPEFFECT_LINK := 4 ;-- Drag source should create a link to the original data.
;-- Key state values (grfKeyState parameter)
MK_LBUTTON := 0x01   ;-- The left mouse button is down.
MK_RBUTTON := 0x02   ;-- The right mouse button is down.
MK_SHIFT   := 0x04   ;-- The SHIFT key is down.
MK_CONTROL := 0x08   ;-- The CTRL key is down.
MK_MBUTTON := 0x10   ;-- The middle mouse button is down.
MK_ALT     := 0x20   ;-- The ALT key is down.
; MK_BUTTON  := ?    ;-- Not documented.
;-- DragDrop includes --------------------------------------------------------------------------------------------------------------
#Include IDropTarget.ahk
#Include DoDragDrop.ahk
;-- GUI ----------------------------------------------------------------------------------------------------------------------------
Gui, +AlwaysOnTop
Gui, Margin, 20, 20
Gui, Add, ListView, r20 w700 hwndhLV vLV, #|FormatNum|FormatName|TYMED|Size|Value
Gui, Add, Text, xp y+2, % "   N/S = not supported"
Gui, Add, Edit, r6 w700 +0x100 hwndhEdit,  ;-- ES_NOHIDESEL:=0x100
   (LTrim
    Possible data source.
    At first select text in this control.
    Then click into the control again and drag to the ListView control above.
    To MOVE text, press Ctrl while dragging the text.
    If it doesn't work, just close your eyes for about 3 seconds and retry.
   )
Gui, Add, StatusBar
Gui, Show, , Drag & Drop Example
;-- Register the ListView as a drop potential target for OLE drag-and-drop operations.
IDT_LV := IDropTarget_Create(hLV, "_LV", [1, 15]) ; CF_TEXT, CF_HDROP
Return
; ==================================================================================================================================
GUIClose:
GUIEscape:
;-- Revoke the registration of the ListView as a potential target for OLE drag-and-drop operations.
IDT_LV.RevokeDragDrop()
Gui, Destroy
ExitApp
; ==================================================================================================================================
; Drop user function called by IDropTarget on drop
; ==================================================================================================================================
IDropTargetOnDrop_LV(TargetObject, pDataObj, KeyState, X, Y, DropEffect) {
   ; Standard clipboard formats
   Static CF := {1:  "CF_TEXT"
               , 2:  "CF_BITMAP"
               , 3:  "CF_METAFILEPICT"
               , 4:  "CF_SYLK"
               , 5:  "CF_DIF"
               , 6:  "CF_TIFF"
               , 7:  "CF_OEMTEXT"
               , 8:  "CF_DIB"
               , 9:  "CF_PALETTE"
               , 10: "CF_PENDATA"
               , 11: "CF_RIFF"
               , 12: "CF_WAVE"
               , 13: "CF_UNICODETEXT"
               , 14: "CF_ENHMETAFILE"
               , 15: "CF_HDROP"
               , 16: "CF_LOCALE"
               , 17: "CF_DIBV5"
               , 0x0080: "CF_OWNERDISPLAY"
               , 0x0081: "CF_DSPTEXT"
               , 0x0082: "CF_DSPBITMAP"
               , 0x0083: "CF_DSPMETAFILEPICT"
               , 0x008E: "CF_DSPENHMETAFILE"}
   ; TYMED enumeration
   Static TM := {1:  "HGLOBAL"
               , 2:  "FILE"
               , 4:  "ISTREAM"
               , 8:  "ISTORAGE"
               , 16: "GDI"
               , 32: "MFPICT"
               , 64: "ENHMF"}
   Static CF_NATIVE := A_IsUnicode ? 13 : 1 ; CF_UNICODETEXT  : CF_TEXT
   ; "Private" formats don't get GlobalFree()'d
   Static CF_PRIVATEFIRST := 0x0200
   Static CF_PRIVATELAST  := 0x02FF
   ; "GDIOBJ" formats do get DeleteObject()'d
   Static CF_GDIOBJFIRST  := 0x0300
   Static CF_GDIOBJLAST   := 0x03FF
   ; "Registered" formats
   Static CF_REGISTEREDFIRST := 0xC000
   Static CF_REGISTEREDLAST  := 0xFFFF
   LV_Delete()
   GuiControl, -Redraw, LV
   If (pEnumObj := IDataObject_EnumFormatEtc(pDataObj)) {
      While IEnumFORMATETC_Next(pEnumObj, FORMATETC) {
         IDataObject_ReadFormatEtc(FORMATETC, Format, Device, Aspect, Index, Type)
         TYMED := "NONE"
         For Index, Value In TM {
            If (Type & Index) {
               TYMED := Value
               Break
            }
         }
         If (Format >= CF_REGISTEREDFIRST) && (Format <= CF_REGISTEREDLAST) {
            VarSetCapacity(Name, 520, 0)
            If !DllCall("GetClipboardFormatName", "UInt", Format, "Str", Name, "UInt", 260)
               Name := "*REGISTERED"
         }
         Else If (Format >= CF_GDIOBJFIRST) && (Format <= CF_GDIOBJLAST)
            Name := "*GDIOBJECT"
         Else If (Format >= CF_PRIVATEFIRST) && (Format <= CF_PRIVATELAST)
            Name := "*PRIVATE"
         Else If !(Name := CF[Format])
            Name := "*UNKNOWN"
         IDataObject_GetData(pDataObj, FORMATETC, Size, Data)
         If (Size = -1)
            Size := "N/S"
         ; Example for getting values out of the returned binary Data
         Value := "N/S"
         If Format In 1,7,13,15,16
         {
            If (Format = CF_NATIVE)       ; CF_TEXT or CF_UNICODETEXT
               Value := StrGet(&Data)
            Else If (Format = 16)         ; CF_LOCALE
               Value := NumGet(Data, "UInt")
            Else If (Format = 15) {       ; CF_HDROP
               LV_Add("", A_Index, Format, Name, TYMED, Size, "")
               Value := ""
               Offset := NumGet(Data, 0, "UInt")
               CP := NumGet(Data, 16, "UInt") ? "UTF-16" : "CP0"
               Offset := NumGet(Data, 0, "UInt")
               While (File := StrGet(&Data + Offset, , CP)) {
                  LV_Add("", "", "", "", "", "", File)
                  Offset += (StrLen(File) + 1) << (CP = "UTF-16")
               }
               Continue
            }
         }
         LV_Add("", A_Index, Format, Name, TYMED, Size, Value)
      }
      ObjRelease(pEnumObj)
   }
   Loop, % LV_GetCount("Column")
      LV_ModifyCol(A_Index, "AutoHdr")
   GuiControl, +Redraw, LV
   Return DropEffect
}
; ==================================================================================================================================
; Context-sensitive Hotkeys
; ==================================================================================================================================
#If MouseOverControl(hEdit)
~LButton::
SB_SetText("")
;-- Anything selected?
DllCall("SendMessage", "Ptr", hEdit, "UInt", 0x00B0, "UIntP", SelBeg := 0, "UIntP", SelEnd := 0) ; EM_GETSEL
If (SelBeg = SelEnd)
   Return
Gui, +OwnDialogs
;-- Copy selected text to the clipboard
SavedClipboard := ClipboardAll ;-- Save the current clipboard
; Clipboard := "" ;-- Empty the clipboard
; SendMessage, 0x0301, 0, 0, , ahk_id %hEdit% ; WM_COPY
ControlGet, Selection, Selected, , , ahk_id %hEdit%
ClipboardSetText(Selection)
;-- Initiate DragDrop
;   Note: The DoDragDrop() function will run until a drop or cancel occurs
SB_SetText("   Drag&Drop has been started ...")
Effect := DoDragDrop()
Effect_Text := {0: "NONE", 1: "COPY", 2: "MOVE", 3: "LINK"}[Effect]
;-- If move was performed, clear (delete) selected text
If (Effect = DROPEFFECT_MOVE)
   SendMessage, 0x0303, 0, 0, , ahk_id %hEdit% ; WM_CLEAR
Else
;-- Otherwise remove the selection
   SendMessage, 0x00B1, SelEnd, SelEnd, , ahk_id %hEdit% ; EM_SETSEL
SB_SetText("   DropEffect: " . Effect_Text)
;-- Restore the original clipboard
Clipboard := SavedClipboard
Return
#If
; ==================================================================================================================================
; Hotkey function
; ==================================================================================================================================
MouseOverControl(hCtrl) {
   MouseGetPos, , , , hMouseOverControl, 2
   Return (hMouseOverControl = hCtrl)
}
; ----------------------------------------------------------------------------------------------------------------------------------
; Just a test to see what happens if you put only one of the text formats onto the clipboard. It turned out that the data object
; created by OleGetClipboard() contains the same 4 formats (CF_LOCALE, CF_OEMTEXT, CF_TEXT, CF_UNICODETEXT) as though
;     Clipboard := TextToSet
; is used.
; ----------------------------------------------------------------------------------------------------------------------------------
ClipboardSetText(TextToSet) {
   Static SizeT := A_IsUnicode ? 2 : 1 ; size of a TCHAR
   Static Format := A_IsUnicode ? 13 : 1 ; CF_UNICODETEXT : CF_TEXT
   ; -------------------------------------------------------------------------------------------------------------------------------
   ; Add text to the clipboard
   Length := StrLen(TextToSet) + 1
   If DllCall("OpenClipboard", "Ptr", A_ScriptHwnd) && DllCall("EmptyClipboard") {
      ; 0x42 = GMEM_MOVEABLE (0x02) | GMEM_ZEROINIT (0x40)
      hMem := DllCall("GlobalAlloc", "UInt", 0x42, "UInt", Length * SizeT, "UPtr")
      pMem := DllCall("GlobalLock", "Ptr", hMem)
      StrPut(TextToSet, pMem, Length)
      DllCall("GlobalUnlock", "Ptr", hMem)
      DllCall("SetClipboardData", "UInt", Format, "UPtr", hMem)
      DllCall("CloseClipboard")
      Return Length
   }
   Return False
}
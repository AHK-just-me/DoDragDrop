; ==================================================================================================================================
; DoDragDrop -> msdn.microsoft.com/en-us/library/ms678486(v=vs.85).aspx
; Carries out an OLE drag and drop operation using the current contents of the clipboard.
; Parameter:
;     DropEffects -  One or the combination of the values:
;                    DROPEFFECT_COPY = 1 ;-- Drop results in a copy. The original data is untouched by the drag source.
;                    DROPEFFECT_MOVE = 2 ;-- Drag source should remove the data.
; Return values:
;     If the data have been dropped successfully, the functions returns the performed drop operation (i.e. 1 for
;     DROPEFFECT_COPY or 2 for DROPEFFECT_MOVE). In all other cases the function returns 0.
; ==================================================================================================================================
DoDragDrop(DropEffects := 0x03, DragKey := 0x01) {
   ; DRAGDROP_S_DROP = 0x40100
   DropEffects &= 0x03
   If !(DropEffects)
      DropEffects := 0x01
   If DllCall("Ole32.dll\OleGetClipboard", "PtrP", pDataObj, "UInt")
      Return False
   IDS := IDropSource_Create(DragKey)
   RC := DllCall("Ole32.dll\DoDragDrop","Ptr", pDataObj, "Ptr", IDS, "UInt", DropEffects, "PtrP", Effect, "Int")
   ObjRelease(pDataObj)
   IDropSource_Free(IDS)
   Return (RC = 0x40100 ? Effect : 0)
}
; ==================================================================================================================================
#Include *i %A_ScriptDir%\IDropSource.ahk
; ==================================================================================================================================
; ==================================================================================================================================
; IDropSource interface -> msdn.microsoft.com/en-us/library/ms690071(v=vs.85).aspx
; Note: Right-drag is not supported as yet!
; ==================================================================================================================================
IDropSource_Create() {
   Static Methods := ["QueryInterface", "AddRef", "Release", "QueryContinueDrag", "GiveFeedback"]
   Static Params  := [3, 1, 1, 3, 2]
   Static VTBL, Dummy := VarSetCapacity(VTBL, A_PtrSize, 0)
   If (NumGet(VTBL, "UPtr") = 0) {
      VarSetCapacity(VTBL, (Methods.Length() + 2) * A_PtrSize, 0)
      NumPut(&VTBL + A_PtrSize, VTBL, "UPtr")
      For Index, Method In Methods {
         CB := RegisterCallback("IDropSource_" . Method, "", Params[Index])
         NumPut(CB, VTBL, Index * A_PtrSize, "UPtr")
      }
   }
   Return &VTBL
}
; ----------------------------------------------------------------------------------------------------------------------------------
IDropSource_Free(IDropSource) {
   NumPut(0, IDropSource + 0, "UPtr")
   While (CB := NumGet(IDropSource + (A_PtrSize * A_Index), "Ptr"))
      DllCall("GlobalFree", "Ptr", CB)
}
; ==================================================================================================================================
; The following functions must not be called directly, they are reserved for internal and system use.
; ==================================================================================================================================
IDropSource_QueryInterface(IDropSource, RIID, PPV) {
   ; IUnknown -> msdn.microsoft.com/en-us/library/ms682521(v=vs.85).aspx
   Static IID := "{00000121-0000-0000-C000-000000000046}"
   VarSetCapacity(QID, 80, 0)
   QIDLen := DllCall("Ole32.dll\StringFromGUID2", "Ptr", RIID, "Ptr", &QID, "Int", 40, "Int")
   If (StrGet(&QID, QIDLen, "UTF-16") = IID) {
      NumPut(IDropSource, PPV + 0, "Ptr")
      Return 0 ; S_OK
   }
   Else {
      NumPut(0, PPV + 0, "Ptr")
      Return 0x80004002 ; E_NOINTERFACE
   }
}
; ----------------------------------------------------------------------------------------------------------------------------------
IDropSource_AddRef(IDropSource) {
   ; IUnknown -> msdn.microsoft.com/en-us/library/ms691379(v=vs.85).aspx
   ; Reference counting is not needed in this case.
   Return 1
}
; ----------------------------------------------------------------------------------------------------------------------------------
IDropSource_Release(IDropSource) {
   ; IUnknown -> msdn.microsoft.com/en-us/library/ms682317(v=vs.85).aspx
   ; Reference counting is not needed in this case.
   Return 0
}
; ----------------------------------------------------------------------------------------------------------------------------------
IDropSource_QueryContinueDrag(IDropSource, fEscapePressed, grfKeyState) {
   ; QueryContinueDrag -> msdn.microsoft.com/en-us/library/ms690076(v=vs.85).aspx
   ; DRAGDROP_S_CANCEL : S_OK : DRAGDROP_S_DROP
   Return (fEscapePressed ? 0x40101 : (grfKeyState & 0x01) ? 0 : 0x40100)
}
; ----------------------------------------------------------------------------------------------------------------------------------
IDropSource_GiveFeedback(IDropSource, dwEffect) {
   ; GiveFeedback -> msdn.microsoft.com/en-us/library/ms693723(v=vs.85).aspx
   Return 0x40102 ; DRAGDROP_S_USEDEFAULTCURSORS
}
; ==================================================================================================================================
#Include *i %A_ScriptDir%\IDragSourceHelper.ahk
; ==================================================================================================================================
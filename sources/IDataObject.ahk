; ==================================================================================================================================
; IDataObject interface -> msdn.microsoft.com/en-us/library/ms688421(v=vs.85).aspx
; Partial implementation, only methods which get data from the object are implemented.
; ==================================================================================================================================
IDataObject_EnumFormatEtc(pDataObj) {
   ; EnumFormatEtc -> msdn.microsoft.com/en-us/library/ms683979(v=vs.85).aspx
   ; DATADIR_GET = 1
   Static EnumFormatEtc := A_PtrSize * 8
   pVTBL := NumGet(pDataObj + 0, "UPtr")
   If !DllCall(NumGet(pVTBL + EnumFormatEtc, "UPtr"), "Ptr", pDataObj, "UInt", 1, "PtrP", ppenumFormatEtc, "Int")
      Return ppenumFormatEtc
   Return False
}
; ==================================================================================================================================
IDataObject_GetData(pDataObj, ByRef FORMATETC, ByRef Size, ByRef Data) {
   ; GetData -> msdn.microsoft.com/en-us/library/ms678431(v=vs.85).aspx
   Static GetData := A_PtrSize * 3
   Data := ""
   , Size := -1
   , VarSetCapacity(STGMEDIUM, 24, 0) ; 64-bit
   , pVTBL := NumGet(pDataObj + 0, "UPtr")
   If !DllCall(NumGet(pVTBL + GetData, "UPtr"), "Ptr", pDataObj, "Ptr", &FORMATETC, "Ptr", &STGMEDIUM, "Int") {
      If (NumGet(STGMEDIUM, "UInt") = 1) { ; TYMED_HGLOBAL
         hGlobal := NumGet(STGMEDIUM, A_PtrSize, "UPtr")
         , pGlobal := DllCall("GlobalLock", "Ptr", hGlobal, "Uptr")
         , Size := DllCall("GlobalSize", "Ptr", hGlobal, "UPtr")
         , VarSetCapacity(Data, Size, 0)
         , DllCall("RtlMoveMemory", "Ptr", &Data, "Ptr", pGlobal, "Ptr", Size)
         , DllCall("GlobalUnlock", "Ptr", hGlobal)
         , DllCall("Ole32.dll\ReleaseStgMedium", "Ptr", &STGMEDIUM)
         Return True
      }
      DllCall("Ole32.dll\ReleaseStgMedium", "Ptr", &STGMEDIUM)
   }
   Return False
}
; ==================================================================================================================================
IDataObject_QueryGetData(pDataObj, ByRef FORMATETC) {
   ; QueryGetData -> msdn.microsoft.com/en-us/library/ms680637(v=vs.85).aspx
   Static QueryGetData := A_PtrSize * 5
   pVTBL := NumGet(pDataObj + 0, "UPtr")
   Return !DllCall(NumGet(pVTBL + QueryGetData, "UPtr"), "Ptr", pDataObj, "Ptr", &FORMATETC, "Int")
}
; ==================================================================================================================================
; Auxiliary functions to create / get data out of a FORMATETC structure
; -> msdn.microsoft.com/en-us/library/ms682242(v=vs.85).aspx
; ==================================================================================================================================
IDataObject_CreateFormatEtc(ByRef FORMATETC, Format, Aspect := 1, Index := -1, Tymed := 1) {
   ; DVASPECT_CONTENT = 1, Index all data = -1, TYMED_HGLOBAL = 1
   VarSetCapacity(FORMATETC, 32, 0) ; 64-bit
   , NumPut(Format, FORMATETC, 0, "Ushort")
   , NumPut(Aspect, FORMATETC, A_PtrSize = 8 ? 16 : 8 , "UInt")
   , NumPut(Index, FORMATETC, A_PtrSIze = 8 ? 20 : 12, "Int")
   , NumPut(Tymed, FORMATETC, A_PtrSize = 8 ? 24 : 16, "UInt")
   Return &FORMATETC
}
; ----------------------------------------------------------------------------------------------------------------------------------
IDataObject_ReadFormatEtc(ByRef FORMATETC, ByRef Format, ByRef Device, ByRef Aspect, ByRef Index, ByRef Tymed) {
   Format := NumGet(FORMATETC, OffSet := 0, "UShort")
   , Device := NumGet(FORMATETC, Offset += A_PtrSize, "UPtr")
   , Aspect := NumGet(FORMATETC, Offset += A_PtrSize, "UInt")
   , Index  := NumGet(FORMATETC, Offset += 4, "Int")
   , Tymed  := NumGet(FORMATETC, Offset += 4, "UInt")
}
; ==================================================================================================================================
#Include *i %A_ScriptDir%\IEnumFORMATETC.ahk
; ==================================================================================================================================
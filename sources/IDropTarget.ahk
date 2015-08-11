; ==================================================================================================================================
; IDropTarget interface -> msdn.microsoft.com/en-us/library/ms679679(v=vs.85).aspx
; Requires: IDataObject.ahk
; ==================================================================================================================================
; Creates a new instance of the IDropTarget object.
; Parameters:
;     HWND              -  HWND of the Gui window or control which shall be used as a drop target.
;     UserFuncSuffix    -  The suffix for the names of the user-defined functions which will be called on events (see Remarks).
;     RequiredFormats   -  An array containing the numeric clipboard formats required to permit drop.
;                          If omitted, only 15 (CF_HDROP) used for dropping files will be required.
;     Register          -  If set to True the target will be registered as a drop target on creation.
;                          Otherwise you have to call the RegisterDragDrop() method manually to activate the drop target.
; Return value:
;     New IDropTarget instance on success; in case of parameter errors, False.
; Remarks:
;     The interface permits up to 4 user-defined functions which will be called from the related methods:
;        IDropTargetOnEnter      Optional, called from IDropTarget.DragEnter()
;        IDropTargetOnOver       Optional, called from IDropTarget.DragOver()
;        IDropTargetOnLeave      Optional, called from IDropTarget.DragLeave()
;        IDropTargetOnDrop       Mandatory, called from IDropTarget.Drop()
;     You have to pass an instance specific unique suffix in UserFuncSuffix which will be appended to the above names to identify
;     the functions.
;
;     Function parameters:
;        IDropTargetOnDrop and IDropTargetOnEnter must accept at least 6 Parameters:
;           TargetObject   -  This instance.
;           pDataObj       -  A pointer to the IDataObject interface on the data object being dropped.
;           KeyState       -  The current state of the mouse buttons and keyboard modifier keys.
;           X              -  The current X coordinate of the cursor in screen coordinates.
;           Y              -  The current Y coordinate of the cursor in screen coordinates.
;           DropEffect     -  The drop effect determined by the Drop() method.
;        IDropTargetOnOver must accept at least 5 Parameters:
;           TargetObject   -  This instance.
;           KeyState       -  The current state of the mouse buttons and keyboard modifier keys.
;           X              -  The current X coordinate of the cursor in screen coordinates.
;           Y              -  The current Y coordinate of the cursor in screen coordinates.
;           DropEffect     -  The drop effect determined by the Drop() method.
;        IDropTargetOnLeave must accept at least 1 Parameters:
;           TargetObject   -  This instance.
;
;     What the functions must return:
;        The return value of IDropTargetOnDrop, IDropTargetOnEnter, and IDropTargetOnOver is used as the drop effect reported
;        as the result of the drop operation. In the easiest case the function returns the value passed in DropEffect.
;        Otherwise, it must return one of the following values:
;           0 (DROPEFFECT_NONE)
;           1 (DROPEFFECT_COPY)
;           2 (DROPEFFECT_MOVE)
;        The return value of IDropTargetOnLeave is not used.
; ==================================================================================================================================
IDropTarget_Create(HWND, UserFuncSuffix, RequiredFormats := "", Register := True) {
   Return New IDropTarget(HWND, UserFuncSuffix, RequiredFormats, Register)
}
; ==================================================================================================================================
Class IDropTarget {
   __New(HWND, UserFuncSuffix, RequiredFormats := "", Register := True) {
      Static Methods := ["QueryInterface", "AddRef", "Release", "DragEnter", "DragOver", "DragLeave", "Drop"]
      Static Params := (A_PtrSize = 8 ? [3, 1, 1, 5, 4, 1, 5] : [3, 1, 1, 6, 5, 1, 6])
      Static DefaultFormat := 15 ; CF_HDROP
      Static DropFunc := "IDropTargetOnDrop"
      Static EnterFunc := "IDropTargetOnEnter"
      Static OverFunc := "IDropTargetOnOver"
      Static LeaveFunc := "IDropTargetOnLeave"
      If This.Base.HasKey("Ptr")
         Return False
      UserFunc := DropFunc . UserFuncSuffix
      If !IsFunc(UserFunc) || (Func(UserFunc).MinParams < 6)
         Return False
      This.DropUserFunc := Func(UserFunc)
      UserFunc := EnterFunc . UserFuncSuffix
      If (IsFunc(UserFunc) && (Func(UserFunc).MinParams > 5))
         This.EnterUserFunc := Func(UserFunc)
      UserFunc := OverFunc . UserFuncSuffix
      If (IsFunc(UserFunc) && (Func(UserFunc).MinParams > 4))
         This.OverUserFunc := Func(UserFunc)
      UserFunc := LeaveFunc . UserFuncSuffix
      If (IsFunc(UserFunc) && (Func(UserFunc).MinParams > 0))
         This.LeaveUserFunc := Func(UserFunc)
      This.HWND := HWND
      This.Registered := False
      If IsObject(RequiredFormats)
         This.Required := RequiredFormats
      Else
         This.Required := [DefaultFormat]
      This.PreferredDropEffect := 0
      SizeOfVTBL := (Methods.Length() + 2) * A_PtrSize
      This.SetCapacity("VTBL", SizeOfVTBL)
      This.Ptr := This.GetAddress("VTBL")
      DllCall("RtlZeroMemory", "Ptr", This.Ptr, "Ptr", SizeOfVTBL)
      NumPut(This.Ptr + A_PtrSize, This.Ptr + 0, "UPtr")
      For Index, Method In Methods {
         CB := RegisterCallback("IDropTarget." . Method, "", Params[Index], &This)
         NumPut(CB, This.Ptr + 0, A_Index * A_PtrSize, "UPtr")
      }
      If (Register)
         If !This.RegisterDragDrop()
            Return False
   }
   ; -------------------------------------------------------------------------------------------------------------------------------
   ; Registers window/control as a drop target.
   ; -------------------------------------------------------------------------------------------------------------------------------
   RegisterDragDrop() {
      If !(This.Registered)
         If DllCall("Ole32.dll\RegisterDragDrop", "Ptr", This.HWND, "Ptr", This.Ptr, "Int")
            Return False
      Return (This.Registered := True)
   }
   ; -------------------------------------------------------------------------------------------------------------------------------
   ; Revokes registering of the window/control as a drop target.
   ; This method should be called before the window/control will be destroyed.
   ; -------------------------------------------------------------------------------------------------------------------------------
   RevokeDragDrop() {
      If (This.Registered)
         DllCall("Ole32.dll\RevokeDragDrop", "Ptr", This.HWND)
      Return !(This.Registered := False)
   }
   ; ===============================================================================================================================
   ; The following methods must not be called directly, they are reserved for internal and system use.
   ; ===============================================================================================================================
   __Delete() {
      This.RevokeDragDrop()
      While (CB := NumGet(This.Ptr + (A_PtrSize * A_Index), "Ptr"))
         DllCall("GlobalFree", "Ptr", CB)
   }
   ; -------------------------------------------------------------------------------------------------------------------------------
   QueryInterface(RIID, PPV) {
      ; IUnknown -> msdn.microsoft.com/en-us/library/ms682521(v=vs.85).aspx
      Static IID := "{00000122-0000-0000-C000-000000000046}"
      VarSetCapacity(QID, 80, 0)
      QIDLen := DllCall("Ole32.dll\StringFromGUID2", "Ptr", RIID, "Ptr", &QID, "Int", 40, "Int")
      If (StrGet(&QID, QIDLen, "UTF-16") = IID) {
         NumPut(This, PPV + 0, "Ptr")
         Return 0 ; S_OK
      }
      Else {
         NumPut(0, PPV + 0, "Ptr")
         Return 0x80004002 ; E_NOINTERFACE
      }
   }
   ; -------------------------------------------------------------------------------------------------------------------------------
   AddRef() {
      ; IUnknown -> msdn.microsoft.com/en-us/library/ms691379(v=vs.85).aspx
      ; Reference counting is not needed in this case.
      Return 1
   }
   ; -------------------------------------------------------------------------------------------------------------------------------
   Release() {
      ; IUnknown -> msdn.microsoft.com/en-us/library/ms682317(v=vs.85).aspx
      ; Reference counting is not needed in this case.
      Return 0
   }
   ; -------------------------------------------------------------------------------------------------------------------------------
   DragEnter(pDataObj, grfKeyState, P3 := "", P4 := "", P5 := "") {
      ; DragEnter -> msdn.microsoft.com/en-us/library/ms680106(v=vs.85).aspx
      ; Params 32: IDataObject *pDataObj, DWORD grfKeyState, LONG x, LONG y, DWORD *pdwEffect
      ; Params 64: IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect
      Instance := Object(A_EventInfo)
      Effect := 0
      If !(grfKeyState & 0x02) { ; right-drag isn't supported by default
         For Each, Format In Instance.Required {
            IDataObject_CreateFormatEtc(FORMATETC, Format)
            If (Effect := IDataObject_QueryGetData(pDataObj, FORMATETC))
               Break
         }
      }
      If (Effect) && (Instance.EnterUserFunc) {
         If (A_PtrSize = 8)
            X := P2 & 0xFFFFFFFF, Y := P2 >> 32
         Else
            X := P2, Y := P3
         Effect := Instance.EnterUserFunc.Call(Instance, pDataObj, grfKeyState, X, Y, Effect)
      }
      Instance.PreferredDropEffect := Effect
      ; If Ctrl and/or Shift is pressed swap the effect
      Effect ^= grfKeyState & 0x0C ? 3 : 0
      NumPut(Effect, (A_PtrSize = 8 ? P4 : P5) + 0, "UInt")
      Return 0 ; S_OK
   }
   ; -------------------------------------------------------------------------------------------------------------------------------
   DragOver(grfKeyState, P2 := "", P3 := "", P4 := "") {
      ; DragOver -> msdn.microsoft.com/en-us/library/ms680129(v=vs.85).aspx
      ; Params 32: DWORD grfKeyState, LONG x, LONG y, DWORD *pdwEffect
      ; Params 64: DWORD grfKeyState, POINTL pt, DWORD *pdwEffect
      Instance := Object(A_EventInfo)
      ; If Ctrl and/or Shift is pressed swap the effect
      Effect := Instance.PreferredDropEffect ^ (grfKeyState & 0x0C ? 3 : 0)
      If (Effect) && (Instance.OverUserFunc) {
         If (A_PtrSize = 8)
            X := P2 & 0xFFFFFFFF, Y := P2 >> 32
         Else
            X := P2, Y := P3
         Effect := Instance.OverUserFunc.Call(Instance, grfKeyState, X, Y, Effect)
      }
      NumPut(Effect, (A_PtrSize = 8 ? P3 : P4) + 0, "UInt")
      Return 0 ; S_OK
   }
   ; -------------------------------------------------------------------------------------------------------------------------------
   DragLeave() {
      ; DragLeave -> msdn.microsoft.com/en-us/library/ms680110(v=vs.85).aspx
      Instance := Object(A_EventInfo)
      Instance.PreferredDropEffect := 0
      If (Instance.LeaveUserFunc)
         Instance.LeaveUserFunc.Call(Instance)
      Return 0 ; S_OK
   }
   ; -------------------------------------------------------------------------------------------------------------------------------
   Drop(pDataObj, grfKeyState, P3 := "", P4 := "", P5 := "") {
      ; Drop -> msdn.microsoft.com/en-us/library/ms687242(v=vs.85).aspx
      ; Params 32: IDataObject *pDataObj, DWORD grfKeyState, LONG x, LONG y, DWORD *pdwEffect
      ; Params 64: IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect
      Instance := Object(A_EventInfo)
      Effect := Instance.PreferredDropEffect ^ (grfKeyState & 0x0C ? 3 : 0)
      If (A_PtrSize = 8)
         X := P3 & 0xFFFFFFFF, Y := P3 >> 32
      Else
         X := P3, Y := P4
      Effect := Instance.DropUserFunc.Call(Instance, pDataObj, grfKeyState, X, Y, Effect)
      NumPut(Effect, (A_PtrSize = 8 ? P4 : P5) + 0, "UInt")
      ; ObjRelease(pDataObj)
      Return 0 ; S_OK
   }
}
; ==================================================================================================================================
#Include *i %A_ScriptDir%\IDataObject.ahk
; ==================================================================================================================================
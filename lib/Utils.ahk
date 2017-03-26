SwitchIME2(dwLayout:=409) {  ; 切换到英文输入状态，win10下有效
    HKL:=DllCall("LoadKeyboardLayout", Str, dwLayout, UInt, 1)
    ControlGetFocus,ctl,A
    SendMessage,0x50,0,HKL,%ctl%,A
}

repeat(str, times) {
    out := ""
    Loop, % times
    {
        out .= str
    }
    return out
}

Array_concat(ByRef arr1, ByRef arr2) {
    loop % arr2.Length()
    {
        arr1.Push(arr2[A_Index])
    }
}
ComObjError(false)

class WizAcc {
    getLeftTreeItems() {
        ControlGet hwnd, hwnd, , SysTreeView321, ahk_class WizNoteMainFrame
        client := Acc_ObjectFromWindow(hwnd)
        AccTree := Acc_Children(client)[4]

        arr := []
        isFolder := false
        For wach, child in Acc_Children(AccTree)
        {
            if Not IsObject(child) {
                name := AccTree.accName(child)
                ; if (name == "个人笔记") {  ; 忽略无关的
                ;     isFolder := true
                ;     continue
                ; } else if (name == "隐藏未分配标签...") {
                ;     isFolder := false
                ; }

                level := AccTree.accValue(child)
                displayTitle := repeat("  ", level) . name
                info := { "title": name, "childId": child, "obj": AccTree
                    , "level": level, displayTitle: displayTitle }
                arr.Push(info)
            }
        }

        return arr
    }

    fixErrorAccName(info) {  ; 修正错误的情况
        if (info.title == "Load failed") {
            info.title := info.obj.accName(info.childId)
        }
    }

    clickSelected(info) {
        ; this.fixErrorAccName(info)

        Acc := info.obj
        ChildId := info.childId

        GetAccLocation(Acc, ChildId, x, y, w, h)
        ; 0 时是隐藏的
        if (x == 0) {
            return
        }

        ; 展开
        try {
            Acc.accDoDefaultAction(ChildId)
        } catch e {

        }

        x += w / 2
        y += h / 2
        SetControlDelay -1
        ControlClick, x%x% y%y%, ahk_class WizNoteMainFrame,,,, NA
    }
}

GetAccLocation(AccObj, Child=0, byref x="", byref y="", byref w="", byref h="") {
    AccObj.accLocation(ComObj(0x4003,&x:=0), ComObj(0x4003,&y:=0), ComObj(0x4003,&w:=0), ComObj(0x4003,&h:=0), Child)
    return  "x" (x:=NumGet(x,0,"int")) "  "
    .   "y" (y:=NumGet(y,0,"int")) "  "
    .   "w" (w:=NumGet(w,0,"int")) "  "
    .   "h" (h:=NumGet(h,0,"int"))
}


#include <Acc>
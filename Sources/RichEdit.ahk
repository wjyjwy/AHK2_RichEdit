; ======================================================================================================================
; 脚本名称:     RichEdit.ahk
; 命名空间:      RichEdit
; 作者:         just me
; AHK 版本:    2.0.2 (Unicode)
; 操作系统版本:     Win 10 Pro (x64)
; 功能:       此类为富文本编辑控件 (v4.1 Unicode) 提供一些包装函数。
; 更改历史:
;    1.0.00.00    2023-05-23/just me - 初始发布
; 鸣谢:
;    corrupt 提供的 cRichEdit:
;       http://www.autohotkey.com/board/topic/17869-crichedit-standard-richedit-control-for-autohotkey-scripts/
;    jballi 提供的 HE_Print:
;       http://www.autohotkey.com/board/topic/45513-function-he-print-wysiwyg-print-for-the-hiedit-control/
;    majkinetor 提供的 Dlg:
;       http://www.autohotkey.com/board/topic/15836-module-dlg-501/
; ======================================================================================================================
#Requires AutoHotkey v2.0
#DllLoad "Msftedit.dll"
; ======================================================================================================================
Class RichEdit {
   ; ===================================================================================================================
; 类变量 - 请勿更改！！！
; ===================================================================================================================
   ; RichEdit 的回调函数
   Static GetRTFCB := 0
   Static LoadRTFCB := 0
   Static SubclassCB := 0
   ; 在启动时初始化类
   Static __New() {
      ; RichEdit.SubclassCB := CallbackCreate(RichEdit_SubclassProc)
      RichEdit.GetRTFCB := CallbackCreate(ObjBindMethod(RichEdit, "GetRTFProc"), , 4)
      RichEdit.LoadRTFCB := CallbackCreate(ObjBindMethod(RichEdit, "LoadRTFProc"), , 4)
      RichEdit.SubclassCB := CallbackCreate(ObjBindMethod(RichEdit, "SubclassProc"), , 6)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   Static GetRTFProc(dwCookie, pbBuff, cb, pcb) { ; GetRTF 的回调过程
      Static RTF := ""
      If (cb > 0) {
         RTF .= StrGet(pbBuff, cb, "CP0")
         Return 0
      }
      If (dwCookie = "*GetRTF*") {
         Out := RTF
         VarSetStrCapacity(&RTF, 0)
         RTF := ""
         Return Out
      }
      Return 1
   }
   ; -------------------------------------------------------------------------------------------------------------------
   Static LoadRTFProc(FileHandle, pbBuff, cb, pcb) { ; LoadRTF 的回调过程
      Return !DllCall("ReadFile", "Ptr", FileHandle, "Ptr", pbBuff, "UInt", cb, "Ptr", pcb, "Ptr", 0)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   Static SubclassProc(H, M, W, L, I, R) { ; RichEdit 子类过程
      ; 请参阅 -> docs.microsoft.com/en-us/windows/win32/api/commctrl/nc-commctrl-subclassproc
      ; WM_GETDLGCODE = 0x87, DLGC_WANTALLKEYS = 4
      Return (M = 0x87) ? 4 : DllCall("DefSubclassProc", "Ptr", H, "UInt", M, "Ptr", W, "Ptr", L, "Ptr")
   }
   ; ===================================================================================================================
; 构造函数
; ===================================================================================================================
   __New(GuiObj, Options, MultiLine := True) {
      Static WS_TABSTOP := 0x10000, WS_HSCROLL := 0x100000, WS_VSCROLL := 0x200000, WS_VISIBLE := 0x10000000,
             WS_CHILD := 0x40000000,
             WS_EX_CLIENTEDGE := 0x200, WS_EX_STATICEDGE := 0x20000,
             ES_MULTILINE := 0x0004, ES_AUTOVSCROLL := 0x40, ES_AUTOHSCROLL := 0x80, ES_NOHIDESEL := 0x0100,
             ES_WANTRETURN := 0x1000, ES_DISABLENOSCROLL := 0x2000, ES_SUNKEN := 0x4000, ES_SAVESEL := 0x8000,
             ES_SELECTIONBAR := 0x1000000
      Static MSFTEDIT_CLASS := "RICHEDIT50W" ; RichEdit v4.1+ (Unicode)
      ; 指定默认样式和扩展样式
      Styles := WS_TABSTOP | WS_VISIBLE | WS_CHILD | ES_AUTOHSCROLL
      If (MultiLine)
         Styles |= WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_NOHIDESEL | ES_WANTRETURN |
                   ES_DISABLENOSCROLL | ES_SAVESEL ; | WS_HSCROLL | ES_SELECTIONBAR ; does not work properly
      ExStyles := WS_EX_STATICEDGE
      ; 创建控件
      CtrlOpts := "Class" . MSFTEDIT_CLASS . " " . Options . " +" . Styles . " +E" . ExStyles
      This.RE := GuiObj.AddCustom(CtrlOpts)
      ; 初始化控件
      ; EM_SETLANGOPTIONS = 0x0478 (WM_USER + 120)
      ; IMF_AUTOKEYBOARD = 0x01, IMF_AUTOFONT = 0x02
      ; SendMessage(0x0478, 0, 0x03, This.HWND) ; 已注释
      ; 子类化控件以获取 Tab 键并防止 Esc 向父窗口发送 WM_CLOSE 消息。
      ; majkinetor 的出色发现之一！
      DllCall("SetWindowSubclass", "Ptr", This.HWND, "Ptr", RichEdit.SubclassCB, "Ptr", This.HWND, "Ptr", 0)
      This.MultiLine := !!MultiLine
      This.DefFont := This.GetFont(1)
      This.DefFont.Default := 1
      This.BackColor := DllCall("GetSysColor", "Int", 5, "UInt") ; COLOR_WINDOW
      This.TextColor := This.DefFont.Color
      This.TxBkColor := This.DefFont.BkColor
      ; 多行控件的其他设置
      If (MultiLine) {
         ; 调整格式矩形
         RC := This.GetRect()
         This.SetRect(RC.L + 6, RC.T + 2, RC.R, RC.B)
         ; 设置高级排版选项
         ; EM_SETTYPOGRAPHYOPTIONS = 0x04CA (WM_USER + 202)
         ; TO_ADVANCEDTYPOGRAPHY = 1, TO_ADVANCEDLAYOUT = 8 ? 未文档化
         SendMessage(0x04CA, 1, 1, This.HWND)
      }
      ; 如有必要，更正 AHK 字体大小设置
      If (Round(This.DefFont.Size) != This.DefFont.Size) {
         This.DefFont.Size := Round(This.DefFont.Size)
         This.SetDefaultFont()
      }
      ; 初始化打印边距
      This.GetMargins()
      ; 初始化文本限制
      This.LimitText(2147483647)
   }
   ; ===================================================================================================================
; 析构函数
; ===================================================================================================================
   __Delete() {
      If DllCall("IsWindow", "Ptr", This.HWND) && (RichEdit.SubclassCB) {
         DllCall("RemoveWindowSubclass", "Ptr", This.HWND, "Ptr", RichEdit.SubclassCB, "Ptr", 0)
      }
      This.RE := 0
   }
   ; ===================================================================================================================
; GUI控件属性 ========================================================================================================
; ===================================================================================================================
   ClassNN => This.RE.ClassNN
   Enabled => This.RE.Enabled
   Focused => This.RE.Focused
   Gui => This.RE.Gui
   Hwnd => This.RE.Hwnd
   Name {
      Get => This.RE.Name
      Set => This.RE.Name := Value
   }
   Visible => This.RE.Visible
   ; ===================================================================================================================
; GUI控件方法 ========================================================================================================
; ===================================================================================================================
   Focus() => This.RE.Focus()
   GetPos(&X?, &Y?, &W?, &H?) => This.RE.GetPos(&X?, &Y?, &W?, &H?)
   Move(X?, Y?, W?, H?) => This.RE.Move(X?, Y?, W?, H?)
   OnCommand(Code, Callback, AddRemove?) => This.RE.OnCommand(Code, Callback, AddRemove?)
   OnNotify(Code, Callback, AddRemove?) => This.RE.OnNotify(Code, Callback, AddRemove?)
   Opt(Options) => This.RE.Opt(Options)
   Redraw() => This.RE.Redraw()
   ; ===================================================================================================================
; 公共方法 ==========================================================================================================
; ===================================================================================================================
   ; ===================================================================================================================
; 仅供高级用户使用的方法
; ===================================================================================================================
   GetCharFormat() { ; 检索当前选择的字符格式
      ; 有关详细信息，请参阅 http://msdn.microsoft.com/en-us/library/bb787883(v=vs.85).aspx。
      ; 返回包含格式设置的 'CF2' 对象。
      ; EM_GETCHARFORMAT = 0x043A
      CF2 := RichEdit.CHARFORMAT2()
      SendMessage(0x043A, 1, CF2.Ptr, This.HWND)
      Return (CF2.Mask ? CF2 : False)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetCharFormat(CF2) { ; 设置当前选择的字符格式
      ; 有关详细信息，请参阅 http://msdn.microsoft.com/en-us/library/bb787883(v=vs.85).aspx。
      ; CF2 : 类似于 GetCharFormat() 返回的 CF2 对象。
      ; EM_SETCHARFORMAT = 0x0444
      Return SendMessage(0x0444, 1, CF2.Ptr, This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   GetParaFormat() { ; 检索当前选择的段落格式
      ; 有关详细信息，请参阅 http://msdn.microsoft.com/en-us/library/bb787942(v=vs.85).aspx。
      ; 返回包含格式设置的 'PF2' 对象。
      ; EM_GETPARAFORMAT = 0x043D
      PF2 := RichEdit.PARAFORMAT2()
      SendMessage(0x043D, 0, PF2.Ptr, This.HWND)
      Return (PF2.Mask ? PF2 : False)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetParaFormat(PF2) { ; 为当前选择设置段落格式
      ; 有关详细信息，请参阅 http://msdn.microsoft.com/en-us/library/bb787942(v=vs.85).aspx。
      ; PF2 : 类似于 GetParaFormat() 返回的 PF2 对象。
      ; EM_SETPARAFORMAT = 0x0447
      Return SendMessage(0x0447, 0, PF2.Ptr, This.HWND)
   }
   ; ===================================================================================================================
; 控件特定
; ===================================================================================================================
   IsModified() { ; 控件是否已修改？
      ; EM_GETMODIFY = 0xB8
      Return SendMessage(0xB8, 0, 0, This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetModified(Modified := False) {  ; 设置或清除编辑控件的修改标志
      ; EM_SETMODIFY = 0xB9
      Return SendMessage(0xB9, !!Modified, 0, This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetEventMask(Events?) { ; 设置应向控件所有者发送通知代码的事件
      ; Events : 包含一个或多个在 'ENM' 中定义的键的数组。
      ; 有关详细信息，请参阅 http://msdn.microsoft.com/en-us/library/bb774238(v=vs.85).aspx
      ; EM_SETEVENTMASK = 0x0445
      Static ENM := {NONE: 0x00, CHANGE: 0x01, UPDATE: 0x02, SCROLL: 0x04, SCROLLEVENTS: 0x08, DRAGDROPDONE: 0x10,
                     PARAGRAPHEXPANDED: 0x20, PAGECHANGE: 0x40, KEYEVENTS: 0x010000, MOUSEEVENTS: 0x020000,
                     REQUESTRESIZE: 0x040000, SELCHANGE: 0x080000, DROPFILES: 0x100000, PROTECTED: 0x200000,
                     LINK: 0x04000000}
      If !IsSet(Events) || (Type(Events) != "Array")
         Events := ["NONE"]
      Mask := 0
      For Each, Event In Events {
         If ENM.HasProp(Event)
            Mask |= ENM.%Event%
         Else
            Return False
      }
      Return SendMessage(0x0445, 0, Mask, This.HWND)
   }
   ; ===================================================================================================================
   ; 加载和存储 RTF 格式
   ; ===================================================================================================================
   GetRTF(Selection := False) { ; 获取控件的全部内容作为富文本
      ; Selection = False : 全部内容 (默认)
      ; Selection = True  : 当前选择
      ; EM_STREAMOUT = 0x044A
      ; SF_TEXT = 0x1, SF_RTF = 0x2, SF_RTFNOOBJS = 0x3, SF_UNICODE = 0x10, SF_USECODEPAGE =	0x0020
      ; SFF_PLAINRTF = 0x4000, SFF_SELECTION = 0x8000
      ; UTF-8 = 65001, UTF-16 = 1200
      ; Static GetRTFCB := CallbackCreate(RichEdit_GetRTFProc)
      Flags := 0x4022 | (1200 << 16) | (Selection ? 0x8000 : 0)
      ES := Buffer((A_PtrSize * 2) + 4, 0)                  ; EDITSTREAM 结构
      NumPut("UPtr", This.HWND, ES)                         ; dwCookie
      NumPut("UPtr", RichEdit.GetRTFCB, ES, A_PtrSize + 4)  ; pfnCallback
      SendMessage(0x044A, Flags, ES.Ptr, This.HWND)
      Return RichEdit.GetRTFProc("*GetRTF*", 0, 0, 0)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   LoadRTF(FilePath, Selection := False) { ; 将 RTF 文件加载到控件中
      ; FilePath = 文件路径
      ; Selection = False : 全部内容 (默认)
      ; Selection = True  : 当前选择
      ; EM_STREAMIN = 0x0449
      ; SF_TEXT = 0x1, SF_RTF = 0x2, SF_RTFNOOBJS = 0x3, SF_UNICODE = 0x10, SF_USECODEPAGE =	0x0020
      ; SFF_PLAINRTF = 0x4000, SFF_SELECTION = 0x8000
      ; UTF-16 = 1200
      ; Static LoadRTFCB := CallbackCreate(RichEdit_LoadRTFProc)
      Flags := 0x4002 | (Selection ? 0x8000 : 0) ; | (1200 << 16)
      If !(File := FileOpen(FilePath, "r"))
         Return False
      ES := Buffer((A_PtrSize * 2) + 4, 0)                     ; EDITSTREAM 结构
      NumPut("UPtr", File.Handle, ES)                          ; dwCookie
      NumPut("UPtr", RichEdit.LoadRTFCB, ES, A_PtrSize + 4)    ; pfnCallback
      Result := SendMessage(0x0449, Flags, ES.Ptr, This.HWND)
      File.Close()
      Return Result
   }
   ; ===================================================================================================================
   ; 滚动
   ; ===================================================================================================================
   GetScrollPos() { ; 获取当前滚动位置
      ; 返回一个包含 'X' 和 'Y' 键的对象，分别表示滚动位置。
      ; EM_GETSCROLLPOS = 0x04DD
      PT := Buffer(8, 0)
      SendMessage(0x04DD, 0, PT.Ptr, This.HWND)
      Return {X: NumGet(PT, 0, "Int"), Y: NumGet(PT, 4, "Int")}
   }
   ; ------------------------------------------------------------------------------------------------------------------
   SetScrollPos(X, Y) { ; 将富编辑控件的内容滚动到指定点
      ; X : 要滚动到的 x 位置。
      ; Y : 要滚动到的 y 位置。
      ; EM_SETSCROLLPOS = 0x04DE
      PT := Buffer(8, 0)
      NumPut("Int", X, "Int", Y, PT)
      Return SendMessage(0x04DE, 0, PT.Ptr, This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   ScrollCaret() { ; 将插入点滚动到视图中
      ; EM_SCROLLCARET = 0x00B7
      SendMessage(0x00B7, 0, 0, This.HWND)
      Return True
   }
   ; -------------------------------------------------------------------------------------------------------------------
   ShowScrollBar(SB, Mode := True) { ; 显示或隐藏富编辑控件的滚动条
      ; SB   : 指定要显示的滚动条：水平或垂直。
      ;        此参数必须是 1 (SB_VERT) 或 0 (SB_HORZ)。
      ; Mode : 指定 TRUE 显示滚动条，FALSE 隐藏滚动条。
      ; EM_SHOWSCROLLBAR = 0x0460 (WM_USER + 96)
      SendMessage(0x0460, SB, !!Mode, This.HWND)
      Return True
   }
   ; ===================================================================================================================
   ; 文本和选择
   ; ===================================================================================================================
   FindText(Find, Mode?) { ; 在富编辑控件中查找 Unicode 文本
      ; Find : 要搜索的文本。
      ; Mode : 可选数组，包含 'FR' 中指定的一个或多个键。
      ;        有关详细信息，请参阅 http://msdn.microsoft.com/en-us/library/bb788013(v=vs.85).aspx。
      ; 如果找到文本则返回 True；否则返回 false。
      ; EM_FINDTEXTEXW = 0x047C, EM_SCROLLCARET = 0x00B7
      Static FR:= {DOWN: 1, WHOLEWORD: 2, MATCHCASE: 4}
      Flags := 0
      If IsSet(Mode) && (Type(Mode) = "Array") {
         For Each, Value In Mode
            If FR.HasProp(Value)
               Flags |= FR[Value]
      }
      Sel := This.GetSel()
      Min := (Flags & FR.DOWN) ? Sel.E : Sel.S
      Max := (Flags & FR.DOWN) ? -1 : 0
      FTX := Buffer(16 + A_PtrSize, 0)
      NumPut("Int", Min, "Int", Max, "UPtr", StrPtr(Find), FTX)
      SendMessage(0x047C, Flags, FTX.Ptr, This.HWND)
      S := NumGet(FTX, 8 + A_PtrSize, "Int"), E := NumGet(FTX, 12 + A_PtrSize, "Int")
      If (S = -1) && (E = -1)
         Return False
      This.SetSel(S, E)
      This.ScrollCaret()
      Return True
   }
   ; -------------------------------------------------------------------------------------------------------------------
   ; 查找指定字符位置之前或之后的下一个单词分隔符，或检索有关该位置字符的信息。
   FindWordBreak(CharPos, Mode := "Left") { 
      ; CharPos : 字符位置。
      ; Mode    : 可以是 'WB' 中指定的键之一。
      ; 返回单词分隔符的字符索引或根据 'Mode' 的其他值。
      ; 有关详细信息，请参阅 http://msdn.microsoft.com/en-us/library/bb788018(v=vs.85).aspx。
      ; EM_FINDWORDBREAK = 0x044C (WM_USER + 76)
      Static WB := {LEFT: 0, RIGHT: 1, ISDELIMITER: 2, CLASSIFY: 3, MOVEWORDLEFT: 4, MOVEWORDRIGHT: 5, LEFTBREAK: 6
                  , RIGHTBREAK: 7}
      Option := WB.HasProp(Mode) ? WB[Mode] : 0
      Return SendMessage(0x044C, Option, CharPos, This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   GetSelText() { ; 检索当前选择的文本作为纯文本
      ; 返回选择的文本。
      ; EM_GETSELTEXT = 0x043E, EM_EXGETSEL = 0x0434
      Txt := ""
      CR := This.GetSel()
      TxtL := CR.E - CR.S + 1
      If (TxtL > 1) {
         VarSetStrCapacity(&Txt, TxtL)
         SendMessage(0x043E, 0, StrPtr(Txt), This.HWND)
         VarSetStrCapacity(&Txt, -1)
      }
      Return Txt
   }
   ; -------------------------------------------------------------------------------------------------------------------
   GetSel() { ; 检索富编辑控件中选择的开始和结束字符位置
      ; 返回包含键 S (选择开始) 和 E (选择结束) 的对象。
      ; EM_EXGETSEL = 0x0434
      CR := Buffer(8, 0)
      SendMessage(0x0434, 0, CR.Ptr, This.HWND)
      Return {S: NumGet(CR, 0, "Int"), E: NumGet(CR, 4, "Int")}
   }
   ; -------------------------------------------------------------------------------------------------------------------
   GetText() {  ; 获取控件的全部内容作为纯文本
      ; EM_GETTEXTEX = 0x045E
      Txt := ""
      If (TxtL := This.GetTextLen() + 1) {
         GTX := Buffer(12 + (A_PtrSize * 2), 0) ; GETTEXTEX 结构
         NumPut("UInt", TxtL * 2, GTX) ; cb
         NumPut("UInt", 1200, GTX, 8)  ; 代码页 = Unicode
         VarSetStrCapacity(&Txt, TxtL)
         SendMessage(0x045E, GTX.Ptr, StrPtr(Txt), This.HWND)
         VarSetStrCapacity(&Txt, -1)
      }
      Return Txt
   }
   ; -------------------------------------------------------------------------------------------------------------------
   ; GetTextColors() { ; 获取文本和背景颜色 - 未实现
   ; }
   ; -------------------------------------------------------------------------------------------------------------------
   GetTextLen() { ; 以各种方式计算文本长度
      ; EM_GETTEXTLENGTHEX = 0x045F
      GTL := Buffer(8, 0)     ; GETTEXTLENGTHEX 结构
      NumPut( "UInt", 1200, GTL, 4)  ; 代码页 = Unicode
      Return SendMessage(0x045F, GTL.Ptr, 0, This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   ; 检索指定文本在指定范围内的第一次出现位置。
   GetTextPos(Find, Min := 0, Max := -1, Mode := 1) { 
      ; Find : 要搜索的文本。
      ; Min  : 范围中第一个字符之前的字符位置索引。
      ;        要存储为 CHARRANGE 结构中 cpMin 的整数值。
      ;        默认值: 0 - 第一个字符
      ; Max  : 范围中最后一个字符之后的字符位置。
      ;        要存储为 CHARRANGE 结构中 cpMax 的整数值。
      ;        默认值: -1 - 最后一个字符
      ; Mode : 以下值的任意组合:
      ;        0 : 向后搜索, 1 : 向前搜索, 2 : 仅匹配整个单词, 4 : 区分大小写
      ; 如果找到则返回包含键 S (文本开始) 和 E (文本结束) 的对象，否则返回 False。
      ; EM_FINDTEXTEXW = 0x047C
      Flags := Mode & 0x07
      FTX := Buffer(16 + A_PtrSize, 0)
      NumPut("Int", Min, "Int", Max, "UPtr", StrPtr(Find), FTX)
      P := SendMessage(0x047C, Flags, FTX.Ptr, This.Hwnd) << 32 >> 32
      Return (P = -1) ? False : {S: NumGet(FTX, 8 + A_PtrSize, "Int"), E: NumGet(FTX, 12 + A_PtrSize, "Int")}
   }
   ; -------------------------------------------------------------------------------------------------------------------
   GetTextRange(Min, Max) { ; 从富编辑控件检索指定范围的字符
      ; Min : 范围中第一个字符之前的字符位置索引。
      ;       要存储为 CHARRANGE 结构中 cpMin 的整数值。
      ; Max : 范围中最后一个字符之后的字符位置。
      ;       要存储为 CHARRANGE 结构中 cpMax 的整数值。
      ; CHARRANGE -> http://msdn.microsoft.com/en-us/library/bb787885(v=vs.85).aspx
      ; EM_GETTEXTRANGE = 0x044B
      If (Max <= Min)
         Return ""
      Txt := ""
      VarSetStrCapacity(&Txt, Max - Min)
      TR := Buffer(8 + A_PtrSize, 0) ; TEXTRANGE Struktur
      NumPut("UInt", Min, "UInt", Max, "UPtr", StrPtr(Txt), TR)
      SendMessage(0x044B, 0, TR.Ptr, This.HWND)
      VarSetStrCapacity(&Txt, -1)
      Return Txt
   }
   ; -------------------------------------------------------------------------------------------------------------------
   HideSelection(Mode) { ; 隐藏或显示选择
      ; Mode : True 隐藏选择，False 显示选择。
      ; EM_HIDESELECTION = 0x043F (WM_USER + 63)
      SendMessage(0x043F, !!Mode, 0, This.HWND)
      Return True
   }
   ; -------------------------------------------------------------------------------------------------------------------
   LimitText(Limit) { ; 设置用户可以键入或粘贴到富编辑控件中的文本数量的上限
      ; Limit : 指定可以输入的最大文本量。
      ;         如果此参数为零，则使用默认最大值，即 64K 字符。
      ; EM_EXLIMITTEXT =  0x435 (WM_USER + 53)
      SendMessage(0x0435, 0, Limit, This.HWND)
      Return True
   }
   ; -------------------------------------------------------------------------------------------------------------------
   ReplaceSel(Text := "") { ; 用指定文本替换选中的文本
      ; EM_REPLACESEL = 0xC2
      Return SendMessage(0xC2, 1, StrPtr(Text), This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetText(Text := "", Mode?) { ; 替换控件的选择或全部内容
      ; Mode : 选项标志数组。它可以是 'ST' 中定义的键的任何合理组合。
      ; 有关详细信息，请参阅 http://msdn.microsoft.com/en-us/library/bb774284(v=vs.85).aspx。
      ; EM_SETTEXTEX = 0x0461, CP_UNICODE = 1200
      ; ST_DEFAULT = 0, ST_KEEPUNDO = 1, ST_SELECTION = 2, ST_NEWCHARS = 4 ???
      Static ST := {DEFAULT: 0, KEEPUNDO: 1, SELECTION: 2}
      Flags := 0
      If IsSet(Mode) && (Type(Mode) = "Array") {
         For Value In Mode
            If ST.HasProp(Value)
               Flags |= ST[Value]
      }
      CP := 1200
      TxtPtr := StrPtr(Text)
      ; RTF 格式化文本必须以 ANSI 格式传递!!!
      If (SubStr(Text, 1, 5) = "{\rtf") || (SubStr(Text, 1, 5) = "{urtf") {
         Buf := Buffer(StrPut(Text, "CP0"), 0)
         StrPut(Text, Buf, "CP0")
         TxtPtr := Buf.Ptr
         CP := 0
      }
      STX := Buffer(8, 0)     ; SETTEXTEX 结构
      NumPut("UInt", Flags, "UInt", CP, STX) ; 标志, 代码页
      Return SendMessage(0x0461, STX.Ptr, TxtPtr, This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetSel(Start, End) { ; 选择字符范围
      ; Start : 基于零的开始索引
      ; End   : 基于零的结束索引 (-1 = 文本结束))
      ; EM_EXSETSEL = 0x0437
      CR := Buffer(8, 0)
      NumPut("Int", Start, "Int", End, CR)
      Return SendMessage(0x0437, 0, CR.Ptr, This.HWND)
   }
   ; ===================================================================================================================
   ; 外观、样式和选项
   ; ===================================================================================================================
   AutoURL(Mode := 1) { ; 启用或禁用自动URL检测
      ; Mode   :  一个或组合以下值:
      ; Disable                  0
      ; AURL_ENABLEURL           1
      ; AURL_ENABLEEMAILADDR     2     ; Win 8+
      ; AURL_ENABLETELNO         4     ; Win 8+
      ; AURL_ENABLEEAURLS        8     ; Win 8+
      ; AURL_ENABLEDRIVELETTERS  16    ; WIn 8+
      ; EM_AUTOURLDETECT = 0x45B
      RetVal :=  SendMessage(0x045B, Mode & 0x1F, 0, This.HWND)
      WinRedraw(This.HWND)
      Return RetVal
   }
   ; -------------------------------------------------------------------------------------------------------------------
   GetRect(&RC := "") { ; 检索富编辑控件的格式矩形
      ; 返回包含键 L (左)、T (上)、R (右) 和 B (下) 的对象。
      ; 如果在 Rect 参数中传递变量，则完整的 RECT 结构将存储在其中。
      RC := Buffer(16, 0)
      If !This.MultiLine
         Return False
      SendMessage(0x00B2, 0, RC.Ptr, This.HWND)
      Return {L: NumGet(RC, 0, "Int"), T: NumGet(RC, 4, "Int"), R: NumGet(RC, 8, "Int"), B: NumGet(RC, 12, "Int")}
   }
   ; -------------------------------------------------------------------------------------------------------------------
   GetOptions(&Options := "") { ; 检索富编辑控件的选项
      ; 返回一个数组，其中包含当前设置的选项，作为 'ECO' 中定义的键。
      ; 如果在 Option 参数中传递变量，则选项的组合数值将存储在其中。
      ; 有关详细信息，请参阅 http://msdn.microsoft.com/en-us/library/bb774178(v=vs.85).aspx。
      ; EM_GETOPTIONS = 0x044E
      Static ECO := {AUTOWORDSELECTION: 0x01, AUTOVSCROLL: 0x40, AUTOHSCROLL: 0x80, NOHIDESEL: 0x100,
                     READONLY: 0x800, WANTRETURN: 0x1000, SAVESEL: 0x8000, SELECTIONBAR: 0x01000000,
                     VERTICAL: 0x400000}
      Options := SendMessage(0x044E, 0, 0, This.HWND)
      O := []
      For Key, Value In ECO.OwnProps()
         If (Options & Value)
            O.Push(Key)
      Return O
   }
   ; -------------------------------------------------------------------------------------------------------------------.
   GetStyles(&Styles := "") { ; 检索当前编辑样式标志
      ; 返回一个包含 'SES' 中定义的键的对象。
      ; 如果在 Styles 参数中传递变量，则样式的组合数值将存储在其中。
      ; 有关详细信息，请参阅 http://msdn.microsoft.com/en-us/library/bb788031(v=vs.85).aspx。
      ; EM_GETEDITSTYLE	= 0x04CD (WM_USER + 205)
      Static SES := {1: "EMULATESYSEDIT", 1: "BEEPONMAXTEXT", 4: "EXTENDBACKCOLOR", 32: "NOXLTSYMBOLRANGE",
                     64: "USEAIMM", 128: "NOIME", 256: "ALLOWBEEPS", 512: "UPPERCASE", 1024: "LOWERCASE",
                     2048: "NOINPUTSEQUENCECHK", 4096: "BIDI", 8192: "SCROLLONKILLFOCUS", 16384: "XLTCRCRLFTOCR",
                     32768: "DRAFTMODE", 0x0010000: "USECTF", 0x0020000: "HIDEGRIDLINES", 0x0040000: "USEATFONT",
                     0x0080000: "CUSTOMLOOK",0x0100000: "LBSCROLLNOTIFY", 0x0200000: "CTFALLOWEMBED",
                     0x0400000: "CTFALLOWSMARTTAG", 0x0800000: "CTFALLOWPROOFING"}
      Styles := SendMessage(0x04CD, 0, 0, This.HWND)
      S := []
      For Key, Value In SES.OwnProps()
         If (Styles & Key)
            S.Push(Value)
      Return S
   }
   ; -------------------------------------------------------------------------------------------------------------------
   GetZoom() { ; 获取当前缩放比例
      ; 返回百分比形式的缩放比例。
      ; EM_GETZOOM = 0x04E0
      N := Buffer(4, 0), D := Buffer(4, 0)
      SendMessage(0x04CD, N.Ptr, D.Ptr, This.HWND)
      N := NumGet(N, 0, "Int"), D := NumGet(D, 0, "Int")
      Return (N = 0) && (D = 0) ? 100 : Round(N / D * 100)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetBkgndColor(Color) { ; 设置背景颜色
      ; Color : RGB 整数值或 HTML 颜色名称或
      ;         "Auto" 以重置为系统默认颜色。
      ; 返回先前的背景颜色。
      ; EM_SETBKGNDCOLOR = 0x0443
      If (Color = "Auto")
         System := True, Color := 0
      Else
         System := False, Color := This.GetBGR(Color)
      Result := SendMessage(0x0443, System, Color, This.HWND)
      Return This.GetRGB(Result)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetOptions(Options, Mode := "SET") { ; 设置富编辑控件的选项
      ; Options : 作为 'ECO' 中定义的键的选项数组。
      ; Mode    : 设置模式: SET, OR, AND, XOR
      ; 有关详细信息，请参阅 http://msdn.microsoft.com/en-us/library/bb774254(v=vs.85).aspx。
      ; EM_SETOPTIONS = 0x044D
      Static ECO := {AUTOWORDSELECTION: 0x01, AUTOVSCROLL: 0x40, AUTOHSCROLL: 0x80, NOHIDESEL: 0x100, READONLY: 0x800
                   , WANTRETURN: 0x1000, SAVESEL: 0x8000, SELECTIONBAR: 0x01000000, VERTICAL: 0x400000}
           , ECOOP := {SET: 0x01, OR: 0x02, AND: 0x03, XOR: 0x04}
      If (Type(Options) != "Array") || !ECOOP.HasProp(Mode)
         Return False
      O := 0
      For Each, Option In Options {
         If ECO.HasProp(Option)
            O |= ECO.%Option%
         Else
            Return False
      }
      Return SendMessage(0x044D, ECOOP.%Mode%, O, This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetRect(L, T, R, B) { ; 设置多行编辑控件的格式矩形
      ; L (左), T (上), R (右), B (下)
      ; 将所有参数设置为零以将其设置为默认值。
      ; 对于多行控件返回 True。
      If !This.MultiLine
         Return False
      If (L + T + R + B) = 0
         RC := {Ptr: 0}
      Else {
         RC := Buffer(16, 0)
         NumPut("Int", L, "Int", T, "Int", R, "Int", B, RC)
      }
      SendMessage(0xB3, 0, RC.Ptr, This.HWND)
      Return True
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetStyles(Styles) { ; 设置富编辑控件的当前编辑样式标志。
      ; Styles : 包含 'SES' 中定义的一个或多个键的对象。
      ;          如果值为 0，则样式将被删除，否则将被添加。
      ; 有关详细信息，请参阅 http://msdn.microsoft.com/en-us/library/bb774236(v=vs.85).aspx。
      ; EM_SETEDITSTYLE	= 0x04CC (WM_USER + 204)
      Static SES := {EMULATESYSEDIT: 1, BEEPONMAXTEXT: 2, EXTENDBACKCOLOR: 4, NOXLTSYMBOLRANGE: 32, USEAIMM: 64,
                     NOIME: 128, ALLOWBEEPS: 256, UPPERCASE: 512, LOWERCASE: 1024, NOINPUTSEQUENCECHK: 2048,
                     BIDI: 4096, SCROLLONKILLFOCUS: 8192, XLTCRCRLFTOCR: 16384, DRAFTMODE: 32768,
                     USECTF: 0x0010000, HIDEGRIDLINES: 0x0020000, USEATFONT: 0x0040000, CUSTOMLOOK: 0x0080000,
                     LBSCROLLNOTIFY: 0x0100000, CTFALLOWEMBED: 0x0200000, CTFALLOWSMARTTAG: 0x0400000,
                     CTFALLOWPROOFING: 0x0800000}
      If (Type(Styles) != "Object")
         Return False
      Flags := Mask := 0
      For Style, Value In Styles.OwnProps() {
         If SES.HasProp(Style) {
            Mask |= SES.%Style%
            If (Value != 0)
               Flags |= SES.%Style%
         }
      }
      Return Mask ? SendMessage(0x04CC, Flags, Mask, This.HWND) : False
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetZoom(Ratio := "") { ; 设置富编辑控件的缩放比例。
      ; Ratio : 介于 100/64 和 6400 之间的浮点值；比率为 0 时关闭缩放。
      ; EM_SETZOOM = 0x4E1
      Return SendMessage(0x04E1, (Ratio > 0 ? Ratio : 100), 100, This.HWND)
   }
   ; ===================================================================================================================
   ; Copy, paste, etc.
   ; ===================================================================================================================
   CanRedo() { ; 确定控件的重做队列中是否有任何操作。
      ; EM_CANREDO = 0x0455 (WM_USER + 85)
      Return SendMessage(0x0455, 0, 0, This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   CanUndo() { ; 确定编辑控件的撤销队列中是否有任何操作。
      ; EM_CANUNDO = 0x00C6
      Return SendMessage(0x00C6, 0, 0, This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   Clear() { ; 清除编辑控件的内容
      ; WM_CLEAR = 0x303
      Return SendMessage(0x0303, 0, 0, This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   Copy() { ; 复制选定的内容到剪贴板
      ; WM_COPY = 0x301
      Return SendMessage(0x0301, 0, 0, This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   Cut() { ; 剪切选定的内容到剪贴板
      ; WM_CUT = 0x300
      Return SendMessage(0x0300, 0, 0, This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   Paste() { ; 从剪贴板粘贴内容
      ; WM_PASTE = 0x302
      Return SendMessage(0x0302, 0, 0, This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   Redo() { ; 重做上一个操作
      ; EM_REDO := 0x454
      Return SendMessage(0x0454, 0, 0, This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   Undo() { ; 撤销上一个操作
      ; EM_UNDO = 0xC7
      Return SendMessage(0x00C7, 0, 0, This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SelAll() { ; 全选
      ; 选择所有内容
      Return This.SetSel(0, -1)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   Deselect() { ; 取消选择
      ; 取消所有选择
      Sel := This.GetSel()
      Return This.SetSel(Sel.S, Sel.S)
   }
   ; ===================================================================================================================
   ; Font & colors
   ; ===================================================================================================================
   ChangeFontSize(Diff) { ; 更改字体大小
      ; Diff : 任何正整数或负整数，正值被视为 +1，负值被视为 -1。
      ; 返回新的大小。
      ; EM_SETFONTSIZE = 0x04DF
      ; 字体大小在 4 - 11 pt 范围内变化 1，在 12 - 28 pt 范围内变化 2，然后到 36 pt、48 pt、72 pt、80 pt，
      ; 对于 > 80 pt 变化 10。最大值为 160 pt，最小值为 4 pt
      Font := This.GetFont()
      If (Diff > 0 && Font.Size < 160) || (Diff < 0 && Font.Size > 4)
         SendMessage(0x04DF, (Diff > 0 ? 1 : -1), 0, This.HWND)
      Else
         Return False
      Font := This.GetFont()
      Return Font.Size
   }
   ; -------------------------------------------------------------------------------------------------------------------
   GetFont(Default := False) { ; 获取当前字体
      ; 将 Default 设置为 True 以获取默认字体。
      ; 返回包含当前选项的对象（请参阅 SetFont()）
      ; EM_GETCHARFORMAT = 0x043A
      ; BOLD_FONTTYPE = 0x0100, ITALIC_FONTTYPE = 0x0200
      ; CFM_BOLD = 1, CFM_ITALIC = 2, CFM_UNDERLINE = 4, CFM_STRIKEOUT = 8, CFM_PROTECTED = 16, CFM_SUBSCRIPT = 0x30000
      ; CFM_BACKCOLOR = 0x04000000, CFM_CHARSET := 0x08000000, CFM_FACE = 0x20000000, CFM_COLOR = 0x40000000
      ; CFM_SIZE = 0x80000000
      ; CFE_SUBSCRIPT = 0x10000, CFE_SUPERSCRIPT = 0x20000, CFE_AUTOBACKCOLOR = 0x04000000, CFE_AUTOCOLOR = 0x40000000
      ; SCF_SELECTION = 1
      Static Mask := 0xEC03001F
      Static Effects := 0xEC000000
      CF2 := RichEdit.CHARFORMAT2()
      CF2.Mask := Mask
      CF2.Effects := Effects
      SendMessage(0x043A, (Default ? 0 : 1), CF2.Ptr, This.HWND)
      Font := {}
      Font.Name := CF2.FaceName
      Font.Size := CF2.Height / 20
      CFS := CF2.Effects
      Style := (CFS & 1 ? "B" : "") . (CFS & 2 ? "I" : "") . (CFS & 4 ? "U" : "") . (CFS & 8 ? "S" : "")
             . (CFS & 0x10000 ? "L" : "") . (CFS & 0x20000 ? "H" : "") . (CFS & 16 ? "P" : "")
      Font.Style := Style = "" ? "N" : Style
      Font.Color := This.GetRGB(CF2.TextColor)
      If (CF2.Effects & 0x40000000)  ; CFE_AUTOCOLOR
         Font.Color := "Auto"
      Else
         Font.Color := This.GetRGB(CF2.TextColor)
      If (CF2.Effects & 0x04000000) ; CFE_AUTOBACKCOLOR
         Font.BkColor := "Auto"
      Else
         Font.BkColor := This.GetRGB(CF2.BackColor)
      Font.CharSet := CF2.CharSet
      Return Font
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetDefaultFont(Font := "") { ; 设置默认字体
      ; Font : 可选对象 - 请参阅 SetFont()。
      If IsObject(Font) {
         For Key, Value In Font.OwnProps()
            If This.DefFont.HasProp(Key)
               This.DefFont.%Key% := Value
      }
      Return This.SetFont(This.DefFont)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetFont(Font) { ; 设置当前/默认字体
      ; Font : 包含以下键的对象
      ;        Name    : 可选的字体名称
      ;        Size    : 可选的字体大小（以磅为单位）
      ;        Style   : 可选的字符串，包含以下一种或多种样式
      ;                  B = 粗体, I = 斜体, U = 下划线, S = 删除线, L = 下标
      ;                  H = 上标, P = 受保护, N = 正常
      ;        Color   : 可选的文本颜色，作为 RGB 整数值或 HTML 颜色名称
      ;                  "Auto" 表示 "自动"（系统默认）颜色
      ;        BkColor : 可选的文本背景颜色（请参阅 Color）
      ;                  "Auto" 表示 "自动"（系统默认）背景颜色
      ;        CharSet : 可选的字体字符集
      ;                  1 = DEFAULT_CHARSET, 2 = SYMBOL_CHARSET
      ;        空参数保留相应的属性
      ; EM_SETCHARFORMAT = 0x0444
      ; SCF_DEFAULT = 0, SCF_SELECTION = 1
      If (Type(Font) != "Object")
         Return False
      CF2 := RichEdit.CHARFORMAT2()
      Mask := Effects := 0
      If Font.HasProp("Name") && (Font.Name != "") {
         Mask |= 0x20000000, Effects |= 0x20000000 ; CFM_FACE, CFE_FACE
         CF2.FaceName := Font.Name
      }
      If Font.HasProp("Size") && (Font.Size != "") {
         Size := Font.Size
         If (Size < 161)
            Size *= 20
         Mask |= 0x80000000, Effects |= 0x80000000 ; CFM_SIZE, CFE_SIZE
         CF2.Height := Size
      }
      If Font.HasProp("Style") && (Font.Style != "") {
         Mask |= 0x3001F           ; all font styles
         If InStr(Font.Style, "B")
            Effects |= 1           ; CFE_BOLD
         If InStr(Font.Style, "I")
            Effects |= 2           ; CFE_ITALIC
         If InStr(Font.Style, "U")
            Effects |= 4           ; CFE_UNDERLINE
         If InStr(Font.Style, "S")
            Effects |= 8           ; CFE_STRIKEOUT
         If InStr(Font.Style, "P")
            Effects |= 16          ; CFE_PROTECTED
         If InStr(Font.Style, "L")
            Effects |= 0x10000     ; CFE_SUBSCRIPT
         If InStr(Font.Style, "H")
            Effects |= 0x20000     ; CFE_SUPERSCRIPT
      }
      If Font.HasProp("Color") && (Font.Color != "") {
         Mask |= 0x40000000        ; CFM_COLOR
         If (Font.Color = "Auto")
            Effects |= 0x40000000  ; CFE_AUTOCOLOR
         Else
            CF2.TextColor := This.GetBGR(Font.Color)
      }
      If Font.HasProp("BkColor") && (Font.BkColor != "") {
         Mask |= 0x04000000        ; CFM_BACKCOLOR
         If (Font.BkColor = "Auto")
            Effects |= 0x04000000  ; CFE_AUTOBACKCOLOR
         Else
            CF2.BackColor := This.GetBGR(Font.BkColor)
      }
      If Font.HasProp("CharSet") && (Font.CharSet != "") {
         Mask |= 0x08000000, Effects |= 0x08000000 ; CFM_CHARSET, CFE_CHARSET
         CF2.CharSet := Font.CharSet = 2 ? 2 : 1 ; SYMBOL|DEFAULT
      }
      If (Mask != 0) {
         Mode := Font.HasProp("Default") ? 0 : 1
         CF2.Mask := Mask
         CF2.Effects := Effects
         Return SendMessage(0x0444, Mode, CF2.Ptr, This.HWND)
      }
      Return False
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetFontStyles(Styles, Default := False) { ; 为当前选择或默认字体设置字体样式
      ; Styles : 包含以下一种或多种样式的字符串
      ;          B = 粗体, I = 斜体, U = 下划线, S = 删除线, L = 下标, H = 上标, P = 受保护,
      ;          N = 正常（重置所有其他样式）
      ; EM_GETCHARFORMAT = 0x043A, EM_SETCHARFORMAT = 0x0444
      ; CFM_BOLD = 1, CFM_ITALIC = 2, CFM_UNDERLINE = 4, CFM_STRIKEOUT = 8, CFM_PROTECTED = 16, CFM_SUxSCRIPT = 0x30000
      ; CFE_SUBSCRIPT = 0x10000, CFE_SUPERSCRIPT = 0x20000, SCF_SELECTION = 1
      Static FontStyles := {N: 0, B: 1, I: 2, U: 4, S: 8, P: 16, L: 0x010000, H: 0x020000}
      CF2 := RichEdit.CHARFORMAT2()
      CF2.Mask := 0x3001F ; FontStyles
      If InStr(Styles, "N")
         CF2.Effects := 0
      Else
         For Style In StrSplit(Styles)
            CF2.Effects |= FontStyles.HasProp(Style) ? FontStyles.%Style% : 0
      Return SendMessage(0x0444, !Default, CF2.Ptr, This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   ToggleFontStyle(Style) { ; 切换单个字体样式
      ; Style : 以下样式之一
      ;         B = 粗体, I = 斜体, U = 下划线, S = 删除线, L = 下标, H = 上标, P = 受保护,
      ;         N = 正常（重置所有其他样式）
      ; EM_GETCHARFORMAT = 0x043A, EM_SETCHARFORMAT = 0x0444
      ; CFM_BOLD = 1, CFM_ITALIC = 2, CFM_UNDERLINE = 4, CFM_STRIKEOUT = 8, CFM_PROTECTED = 16, CFM_SUBSCRIPT = 0x30000
      ; CFE_SUBSCRIPT = 0x10000, CFE_SUPERSCRIPT = 0x20000, SCF_SELECTION = 1
      Static FontStyles := {N: 0, B: 1, I: 2, U: 4, S: 8, P: 16, L: 0x010000, H: 0x020000}
      If !FontStyles.HasProp(Style)
         Return False
      CF2 := This.GetCharFormat()
      CF2.Mask := 0x3001F ; FontStyles
      If (Style = "N")
         CF2.Effects := 0
      Else
         CF2.Effects ^= FontStyles.%Style%
      Return SendMessage(0x0444, 1, CF2.Ptr, This.HWND)
   }
   ; ===================================================================================================================
   ; Paragraph formatting
   ; ===================================================================================================================
   AlignText(Align := 1) { ; 设置段落对齐方式
      ; 注意: 大于 3 的值似乎不起作用，尽管它们在文档中应该有效
      ; Align: 可以包含以下数字之一:
      ;        PFA_LEFT             1
      ;        PFA_RIGHT            2
      ;        PFA_CENTER           3
      ;        PFA_JUSTIFY          4 // 新的段落对齐选项 2.0 (*)
      ;        PFA_FULL_INTERWORD   4 // 这些在 3.0 中支持高级
      ;        PFA_FULL_INTERLETTER 5 // 排版功能启用
      ;        PFA_FULL_SCALED      6
      ;        PFA_FULL_GLYPHS      7
      ;        PFA_SNAP_GRID        8
      ; EM_SETPARAFORMAT = 0x0447, PFM_ALIGNMENT = 0x08
      Static PFA := {LEFT: 1, RIGHT: 2, CENTER: 3, JUSTIFY: 4}
      If PFA.HasProp(Align)
         Align := PFA.%Align%
      If (Align >= 1) && (ALign <= 8) {
         PF2 := RichEdit.PARAFORMAT2() ; PARAFORMAT2 struct
         PF2.Mask := 0x08              ; dwMask
         PF2.Alignment := Align        ; wAlignment
         SendMessage(0x0447, 0, PF2.Ptr, This.HWND)
         Return True
      }
      Return False
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetBorder(Widths, Styles) { ; 设置段落边框
      ; 边框在 RichEdit 中不显示，因此调用此函数没有可见结果。
      ; 即使是 Win7 随附的 WordPad 也不显示它们，但例如 Word 2007 会显示。
      ; Widths : 4 个边框宽度的数组，范围为 1 - 15，顺序为左、上、右、下；0 = 无边框
      ; Styles : 4 个边框样式的数组，范围为 0 - 7，顺序为左、上、右、下（见备注）
      ; 注意:
      ; MSDN 上 http://msdn.microsoft.com/en-us/library/bb787942(v=vs.85).aspx 的描述是错误的！
      ; 要设置边框，必须将边框宽度放入 wBorderWidth 的相关半字节（4 位）中
      ; （顺序：左 (0 - 3)、上 (4 - 7)、右 (8 - 11) 和下 (12 - 15)。这些值被解释为
      ; 半点（即 10 缇）。边框样式设置在 wBorders 的相关半字节中。
      ; 有效的样式似乎是:
      ;     0 : \brdrdash (虚线)
      ;     1 : \brdrdashsm (小虚线)
      ;     2 : \brdrdb (双线)
      ;     3 : \brdrdot (点线)
      ;     4 : \brdrhair (单线/发丝线)
      ;     5 : \brdrs ? 看起来像 3
      ;     6 : \brdrth ? 看起来像 3
      ;     7 : \brdrtriple (三线)
      ; EM_SETPARAFORMAT = 0x0447, PFM_BORDER = 0x800
      If (Type(Widths) != "Array") ||  (Type(Styles) != "Array") || (Widths.Length != 4) || (Styles.Length != 4)
         Return False
      W := S := 0
      For I, V In Widths {
         If (V)
            W |= V << ((A_Index - 1) * 4)
         If Styles[I]
            S |= Styles[I] << ((A_Index - 1) * 4)
      }
      PF2 := RichEdit.PARAFORMAT2()
      PF2.Mask := 0x800
      PF2.BorderWidth := W
      PF2.Borders := S
      Return SendMessage(0x0447, 0, PF2.Ptr, This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetLineSpacing(Lines) { ; 设置段落的行间距。
      ; Lines : 行数，可以是整数或浮点数。
      ; SpacingRule = 5:
      ; dyLineSpacing / 20 的值是行与行之间的间距（以行为单位）。因此，将
      ; dyLineSpacing 设置为 20 会产生单倍行距的文本，40 是双倍行距，60 是三倍行距，依此类推。
      ; EM_SETPARAFORMAT = 0x0447, PFM_LINESPACING = 0x100
      PF2 := RichEdit.PARAFORMAT2()
      PF2.Mask := 0x100
      PF2.LineSpacing := Abs(Lines) * 20
      PF2.LineSpacingRule := 5
      Return SendMessage(0x0447, 0, PF2.Ptr, This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetParaIndent(Indent := "Reset") { ; 设置段落的左右间距。
      ; Indent : 包含最多三个键的对象:
      ;          - Start  : 可选 - 段落第一行的绝对缩进。
      ;          - Right  : 可选 - 段落右侧相对于右边距的缩进。
      ;          - Offset : 可选 - 第二行及后续行相对于第一行缩进的缩进量。
      ;          值根据用户的区域设置测量单位（厘米/英寸）进行解释。
      ;          不传递参数调用可重置缩进。
      ; EM_SETPARAFORMAT = 0x0447
      ; PFM_STARTINDENT  = 0x0001
      ; PFM_RIGHTINDENT  = 0x0002
      ; PFM_OFFSET       = 0x0004
      Static PFM := {STARTINDENT: 0x01, RIGHTINDENT: 0x02, OFFSET: 0x04}
      Measurement := This.GetMeasurement()
      PF2 := RichEdit.PARAFORMAT2()
      If (Indent = "Reset")
         PF2.Mask := 0x07 ; reset indentation
      Else If !IsObject(Indent)
         Return False
      Else {
         PF2.Mask := 0
         If (Indent.HasProp("Start")) {
            PF2.Mask |= PFM.STARTINDENT
            PF2.StartIndent := Round((Indent.Start / Measurement) * 1440)
         }
         If (Indent.HasProp("Offset")) {
            PF2.Mask |= PFM.OFFSET
            PF2.Offset := Round((Indent.Offset / Measurement) * 1440)
         }
         If (Indent.HasProp("Right")) {
            PF2.Mask |= PFM.RIGHTINDENT
            PF2.RightIndent := Round((Indent.Right / Measurement) * 1440)
         }
      }
      If (PF2.Mask)
         Return SendMessage(0x0447, 0, PF2.Ptr, This.HWND)
      Return False
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetParaNumbering(Numbering := "Reset") {
      ; Numbering : 包含最多四个键的对象:
      ;             - Type  : 用于项目符号或编号段落的选项。
      ;             - Style : 可选 - 用于编号段落的编号样式。
      ;             - Tab   : 可选 - 段落编号与段落文本之间的最小间距。
      ;             - Start : 可选 - 用于编号段落的序列号（例如，3 表示 C 或 III）
      ;             Tab 根据用户的区域设置测量单位（厘米/英寸）进行解释。
      ;             不传递参数调用可重置编号。
      ; EM_SETPARAFORMAT = 0x0447
      ; PARAFORMAT 编号选项
      ; PFN_BULLET   1 ; 项目符号
      ; PFN_ARABIC   2 ; 阿拉伯数字:   0, 1, 2,	...
      ; PFN_LCLETTER 3 ; 小写字母: a, b, c,	...
      ; PFN_UCLETTER 4 ; 大写字母: A, B, C,	...
      ; PFN_LCROMAN  5 ; 小写罗马数字:  i, ii, iii,	...
      ; PFN_UCROMAN  6 ; 大写罗马数字:  I, II, III,	...
      ; PARAFORMAT2 编号样式选项
      ; PFNS_PAREN     0x0000 ; 默认, 例如,                 1)
      ; PFNS_PARENS    0x0100 ; 带括号, 例如, (1)
      ; PFNS_PERIOD    0x0200 ; 带句点, 例如,       1.
      ; PFNS_PLAIN     0x0300 ; 纯数字, 例如,        1
      ; PFNS_NONUMBER  0x0400 ; 用于无编号的继续
      ; PFNS_NEWNUMBER 0x8000 ; 以 wNumberingStart 开始新编号
      ; PFM_NUMBERING      0x0020
      ; PFM_NUMBERINGSTYLE 0x2000
      ; PFM_NUMBERINGTAB   0x4000
      ; PFM_NUMBERINGSTART 0x8000
      Static PFM := {Type: 0x0020, Style: 0x2000, Tab: 0x4000, Start: 0x8000}
      Static PFN := {Bullet: 1, Arabic: 2, LCLetter: 3, UCLetter: 4, LCRoman: 5, UCRoman: 6}
      Static PFNS := {Paren: 0x0000, Parens: 0x0100, Period: 0x0200, Plain: 0x0300, None: 0x0400, New: 0x8000}
      PF2 := RichEdit.PARAFORMAT2()
      If (Numbering = "Reset")
         PF2.Mask := 0xE020
      Else If !IsObject(Numbering)
         Return False
      Else {
         If (Numbering.HasProp("Type")) {
            PF2.Mask |= PFM.Type
            PF2.Numbering := PFN.%Numbering.Type%
         }
         If (Numbering.HasProp("Style")) {
            PF2.Mask |= PFM.Style
            PF2.NumberingStyle := PFNS.%Numbering.Style%
         }
         If (Numbering.HasProp("Tab")) {
            PF2.Mask |= PFM.Tab
            PF2.NumberingTab := Round((Numbering.Tab / This.GetMeasurement()) * 1440)
         }
         If (Numbering.HasProp("Start")) {
            PF2.Mask |= PFM.Start
            PF2.NumberingStart := Numbering.Start
         }
      }
      If (PF2.Mask)
         Return SendMessage(0x0447, 0, PF2.Ptr, This.HWND)
      Return False
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetParaSpacing(Spacing := "Reset") { ; 设置段落前后的间距
      ; Spacing : 包含一个或两个键的对象:
      ;           - Before : 段落前的额外间距（以点为单位）
      ;           - After  : 段落后的额外间距（以点为单位）
      ;           不传递参数调用可将间距重置为零。
      ; EM_SETPARAFORMAT = 0x0447
      ; PFM_SPACEBEFORE  = 0x0040
      ; PFM_SPACEAFTER   = 0x0080
      Static PFM := {Before: 0x40, After: 0x80}
      PF2 := RichEdit.PARAFORMAT2()
      If (Spacing = "Reset")
         PF2.Mask := 0xC0 ; reset spacing
      Else If !IsObject(Spacing)
         Return False
      Else {
         If Spacing.HasProp("Before") && (Spacing.Before >= 0) {
            PF2.Mask |= PFM.Before
            PF2.SpaceBefore := Round(Spacing.Before * 20)
         }
         If Spacing.HasProp("After") && (Spacing.After >= 0) {
            PF2.Mask |= PFM.After
            PF2.SpaceAfter := Round(Spacing.After * 20)
         }
      }
      If (PF2.Mask)
         Return SendMessage(0x0447, 0, PF2.Ptr, This.HWND)
      Return False
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetDefaultTabs(Distance) { ; 设置默认制表位
      ; 距离将根据当前用户的区域设置解释为英寸或厘米。
      ; EM_SETTABSTOPS = 0xCB
      Static DUI := 64      ; 每英寸的对话框单位
           , MinTab := 0.20 ; 最小制表位距离
           , MaxTab := 3.00 ; 最大制表位距离
      IM := This.GetMeasurement()
      Distance := StrReplace(Distance, ",", ".")
      Distance := Round(Distance / IM, 2)
      If (Distance < MinTab)
         Distance := MinTab
      Else If (Distance > MaxTab)
         Distance := MaxTab
      TabStops := Buffer(4, 0)
      NumPut("Int", Round(DUI * Distance), TabStops)
      Result := SendMessage(0x00CB, 1, TabStops.Ptr, This.HWND)
      DllCall("UpdateWindow", "Ptr", This.HWND)
      Return Result
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SetTabStops(TabStops := "Reset") { ; 设置段落的制表位
      ; TabStops 是一个对象，其中包含以百分之一英寸/厘米为单位的整数位置作为键
      ; 以及对齐方式（"L"、"C"、"R" 或 "D"）作为值。
      ; 位置将根据当前用户的区域设置解释为百分之一英寸或厘米。
      ; 不传递参数调用可重置为默认制表位。
      ; EM_SETPARAFORMAT = 0x0447, PFM_TABSTOPS = 0x10
      Static MinT := 30                ; 最小制表位 (百分之一英寸)
      Static MaxT := 830               ; 最大制表位 (百分之一英寸)
      Static Align := {L: 0x00000000   ; 左对齐 (默认)
                     , C: 0x01000000   ; 居中对齐
                     , R: 0x02000000   ; 右对齐
                     , D: 0x03000000}  ; 小数点对齐
      Static MAX_TAB_STOPS := 32
      IC := This.GetMeasurement()
      PF2 := RichEdit.PARAFORMAT2()
      PF2.Mask := 0x10
      If (TabStops = "Reset")
         Return !!SendMessage(0x0447, 0, PF2.Ptr, This.HWND)
      If !IsObject(TabStops)
         Return False
      Tabs  := []
      For Position, Alignment In TabStops.OwnProps() {
         Position /= IC
         If (Position < MinT) Or (Position > MaxT) ||
            !Align.HasProp(Alignment) Or (A_Index > MAX_TAB_STOPS)
            Return False
         Tabs.Push(Align.%Alignment% | Round((Position / 100) * 1440))
      }
      If (Tabs.Length) {
         PF2.Tabs := Tabs
         Return SendMessage(0x0447, 0, PF2.Ptr, This.HWND)
      }
      Return False
   }
   ; ===================================================================================================================
   ; 行处理
   ; ===================================================================================================================
   GetCaretLine() { ; 获取包含插入点的行
      ; EM_LINEINDEX = 0xBB, EM_EXLINEFROMCHAR = 0x0436
      Result := SendMessage(0x00BB, -1, 0, This.HWND)
      Return SendMessage(0x0436, 0, Result, This.HWND) + 1
   }
   ; -------------------------------------------------------------------------------------------------------------------
   GetLineCount() { ; 获取总行数
      ; EM_GETLINECOUNT = 0xBA
      Return SendMessage(0x00BA, 0, 0, This.HWND)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   GetLineIndex(LineNumber) { ; 获取指定行第一个字符的索引。
      ; EM_LINEINDEX := 0x00BB
      ; LineNumber   -  从零开始的行号
      Return SendMessage(0x00BB, LineNumber, 0, This.HWND)
   }
   ; ===================================================================================================================
   ; Statistics
   ; ===================================================================================================================
   GetStatistics() { ; 获取一些统计值
      ; 获取包含插入点的行、插入点在该行中的位置、总行数、绝对插入点
      ; 位置和字符总数。
      ; EM_GETSEL = 0xB0, EM_LINEFROMCHAR = 0xC9, EM_LINEINDEX = 0xBB, EM_GETLINECOUNT = 0xBA
      Stats := {}
      SB := Buffer(A_PtrSize, 0)
      SendMessage(0x00B0, SB.Ptr, 0, This.Hwnd)
      LI := This.GetLineIndex(-1)
      Stats.LinePos := NumGet(SB, "Ptr") - LI + 1
      Stats.Line := SendMessage(0x00C9, -1, 0, This.HWND) + 1
      Stats.LineCount := This.GetLineCount()
      Stats.CharCount := This.GetTextLen()
      Return Stats
   }
   ; ===================================================================================================================
   ; 布局
   ; ===================================================================================================================
   WordWrap(On) { ; 打开/关闭自动换行
      ; EM_SCROLLCARET = 0xB7
      Sel := This.GetSel()
      SendMessage(0x0448, 0, On ? 0 : -1, This.HWND)
      This.SetSel(Sel.S, Sel.E)
      SendMessage(0x00B7, 0, 0, This.HWND)
      Return On
   }
   ; -------------------------------------------------------------------------------------------------------------------
   WYSIWYG(On) { ; 显示控件如同打印效果（所见即所得）
      ; 文本测量基于默认打印机的容量，因此更改打印机可能会产生不同的
      ; 结果。另请参阅 Print() 中的备注/评论。
      ; EM_SCROLLCARET = 0xB7, EM_SETTARGETDEVICE = 0x0448
      ; PD_RETURNDC = 0x0100, PD_RETURNDEFAULT = 0x0400
      Static PDC := 0
      Static PD_Size := (A_PtrSize = 4 ? 66 : 120)
      Static OffFlags := A_PtrSize * 5
      Sel := This.GetSel()
      If !(On) {
         DllCall("LockWindowUpdate", "Ptr", This.HWND)
         DllCall("DeleteDC", "Ptr", PDC)
         SendMessage(0x0448, 0, -1, This.HWND)
         This.SetSel(Sel.S, Sel.E)
         SendMessage(0x00B7, 0, 0, This.HWND)
         DllCall("LockWindowUpdate", "Ptr", 0)
         Return True
      }
      PD := Buffer(PD_Size, 0)
      Numput("UInt", PD_Size, PD)
      NumPut("UInt", 0x0100 | 0x0400, PD, A_PtrSize * 5) ; PD_RETURNDC | PD_RETURNDEFAULT
      If !DllCall("Comdlg32.dll\PrintDlg", "Ptr", PD.Ptr, "Int")
         Return
      DllCall("GlobalFree", "Ptr", NumGet(PD, A_PtrSize * 2, "UPtr"))
      DllCall("GlobalFree", "Ptr", NumGet(PD, A_PtrSize * 3, "UPtr"))
      PDC := NumGet(PD, A_PtrSize * 4, "UPtr")
      DllCall("LockWindowUpdate", "Ptr", This.HWND)
      Caps := This.GetPrinterCaps(PDC)
      ; 设置页面大小和像素边距
      UML := This.Margins.LT                   ; 用户左边距
      UMR := This.Margins.RT                   ; 用户右边距
      PML := Caps.POFX                         ; 物理左边距
      PMR := Caps.PHYW - Caps.HRES - Caps.POFX ; 物理右边距
      LPW := Caps.HRES                         ; 逻辑页面宽度
      ; 调整边距
      UML := UML > PML ? (UML - PML) : 0
      UMR := UMR > PMR ? (UMR - PMR) : 0
      LineLen := LPW - UML - UMR
      SendMessage(0x0448, PDC, LineLen, This.HWND)
      This.SetSel(Sel.S, Sel.E)
      SendMessage(0x00B7, 0, 0, This.HWND)
      DllCall("LockWindowUpdate", "Ptr", 0)
      Return True
   }
   ; ===================================================================================================================
   ; 文件处理
   ; ===================================================================================================================
   LoadFile(File, Mode := "Open") { ; 加载文件
      ; File : 文件名
      ; Mode : Open / Add / Insert
      ;        Open   : 替换控件内容
      ;        Append : 追加到控件内容
      ;        Insert : 在当前选择处插入/替换
      If !FileExist(File)
         Return False
      Ext := ""
      SplitPath(File, , , &Ext)
      If (Ext = "rtf") {
         Switch Mode {
            Case "Open":
               Selection := False
            Case "Insert":
               Selection := True
            Case "Append":
               This.SetSel(-1, -2)
               Selection := True
         }
         This.LoadRTF(File, Selection)
      }
      Else {
         Text := FileRead(File)
         Switch Mode {
            Case "Open":
               This.SetText(Text)
            Case "Insert":
               This.ReplaceSel(Text)
            Case "Append":
               This.SetSel(-1, -2)
               This.ReplaceSel(Text)
         }
      }
      Return True
   }
   ; -------------------------------------------------------------------------------------------------------------------
   SaveFile(File) { ; 保存文件
      ; File : 文件名
      ; 成功时返回 True，否则返回 False。
      This.Gui.Opt("+OwnDialogs")
      Ext := ""
      SplitPath(File, , , &Ext)
      Text := Ext = "rtf" ? This.GetRTF() : This.GetText()
      Try {
         FileObj := FileOpen(File, "w")
         FileObj.Write(Text)
         FileObj.Close()
         Return True
      }
      Catch As Err {
         MsgBox 16, A_ThisFunc, "Couldn't save '" . File . "'!`n`n" . Type(Err) ": " Err.Message
         Return False
      }
   }
   ; ===================================================================================================================
   ; 打印
   ; 感谢 jballi ->  http://www.autohotkey.com/board/topic/45513-function-he-print-wysiwyg-print-for-the-hiedit-control/
   ; ===================================================================================================================
   Print() {
      ; EM_FORMATRANGE = 0x0439, EM_SETTARGETDEVICE = 0x0448
      ; ----------------------------------------------------------------------------------------------------------------
      ; 静态变量
      Static PD_ALLPAGES := 0x00, PD_SELECTION := 0x01, PD_PAGENUMS := 0x02, PD_NOSELECTION := 0x04
           , PD_RETURNDC := 0x0100, PD_USEDEVMODECOPIES := 0x040000, PD_HIDEPRINTTOFILE := 0x100000
           , PD_NONETWORKBUTTON := 0x200000, PD_NOCURRENTPAGE := 0x800000
           , MM_TEXT := 0x1
           , DocName := "AHKRichEdit"
           , PD_Size := (A_PtrSize = 8 ? (13 * A_PtrSize) + 16 : 66)
      ErrorMsg := ""
      ; ----------------------------------------------------------------------------------------------------------------
      ; 准备调用 PrintDlg
      ; 定义/填充 PRINTDLG 结构
      PD := Buffer(PD_Size, 0)
      Numput("UInt", PD_Size, PD)  ; lStructSize
      Numput("UPtr", This.Gui.Hwnd, PD, A_PtrSize) ; hwndOwner
      ; 收集开始/结束选择位置
      Sel := This.GetSel()
      ; 确定/设置标志
      Flags := PD_ALLPAGES | PD_RETURNDC | PD_USEDEVMODECOPIES | PD_HIDEPRINTTOFILE | PD_NONETWORKBUTTON
             | PD_NOCURRENTPAGE
      If (Sel.S = Sel.E)
         Flags |= PD_NOSELECTION
      Else
         Flags |= PD_SELECTION
      Offset := A_PtrSize * 5
      ; Flags, pages, and copies
      NumPut("UInt", Flags, "UShort", 1, "UShort", 1, "UShort", 1, "UShort", -1, "UShort", 1, PD, Offset)
      ; 注意：使用 -1 来指定最大页码 (65535)。
      ; 编程注意：加载到这些字段的值是关键的。如果将意外的值加载到一个或多个这些字段，打印对话框将不会
      ; 显示（返回错误）。
      ; ----------------------------------------------------------------------------------------------------------------
      ; 打印对话框
      ; 打开打印对话框。如果用户取消，则退出。
      If !DllCall("Comdlg32.dll\PrintDlg", "Ptr", PD, "UInt")
         Throw Error("Function: " . A_ThisFunc . " - DLLCall of 'PrintDlg' failed.", -1)
      ; 获取打印机设备上下文。如果未定义则退出。
      If !(PDC := NumGet(PD, A_PtrSize * 4, "UPtr")) ; hDC
         Throw Error("Function: " . A_ThisFunc . " - 无法获取打印机的设备上下文。", -1)
      ; 释放由 PrintDlg 创建的全局结构
      DllCall("GlobalFree", "Ptr", NumGet(PD, A_PtrSize * 2, "UPtr"))
      DllCall("GlobalFree", "Ptr", NumGet(PD, A_PtrSize * 3, "UPtr"))
      ; ----------------------------------------------------------------------------------------------------------------
      ; 准备打印
      ; 收集标志
      Offset := A_PtrSize * 5
      Flags := NumGet(PD, OffSet, "UInt")           ; 标志
      ; 确定起始/结束页码
      If (Flags & PD_PAGENUMS) {
         PageF := NumGet(PD, Offset += 4, "UShort") ; 起始页
         PageL := NumGet(PD, Offset += 2, "UShort") ; 结束页
      }
      Else
         PageF := 1, PageL := 65535
      ; 收集打印机容量
      Caps := This.GetPrinterCaps(PDC)
      ; 设置页面大小和边距（以缇为单位，1/20 点或 1/1440 英寸）
      UML := This.Margins.LT                   ; 用户左边距
      UMT := This.Margins.TT                   ; 用户上边距
      UMR := This.Margins.RT                   ; 用户右边距
      UMB := This.Margins.BT                   ; 用户下边距
      PML := Caps.POFX                         ; 物理左边距
      PMT := Caps.POFY                         ; 物理上边距
      PMR := Caps.PHYW - Caps.HRES - Caps.POFX ; 物理右边距
      PMB := Caps.PHYH - Caps.VRES - Caps.POFY ; 物理下边距
      LPW := Caps.HRES                         ; 逻辑页面宽度
      LPH := Caps.VRES                         ; 逻辑页面高度
      ; 调整边距
      UML := UML > PML ? (UML - PML) : 0
      UMT := UMT > PMT ? (UMT - PMT) : 0
      UMR := UMR > PMR ? (UMR - PMR) : 0
      UMB := UMB > PMB ? (UMB - PMB) : 0
      ; 定义/填充 FORMATRANGE 结构
      FR := Buffer((A_PtrSize * 2) + (4 * 10), 0)
      NumPut("UPtr", PDC, "UPtr", PDC, FR) ; hdc , hdcTarget
      ; 定义 FORMATRANGE.rc
      ; rc 是要渲染的区域 (rcPage - 边距)，以缇为单位（1/20 点或 1/1440 英寸）。
      ; 如果用户定义的边距小于打印机的边距（每页边缘的不可打印区域），则用户边距设置为打印机的边距。此外，用户定义的边距
      ; 必须调整以考虑打印机的边距。
      ; 例如：如果用户要求 3/4 英寸 (19.05 毫米) 的左边距，但打印机的左边距是
      ; 1/4 英寸 (6.35 毫米)，则 rc.Left 设置为 720 缇 (1/2 英寸或 12.7 毫米)。
      Offset := A_PtrSize * 2
      NumPut("Int", UML, "Int", UMT, "Int", LPW - UMR, "Int", LPH - UMB, FR, Offset)
      ; 定义 FORMATRANGE.rcPage
      ; rcPage 是渲染设备上一页的整个区域，以缇为单位（1/20 点或 1/1440 英寸）
      ; 注意：rc 定义了最大可打印区域，不包括打印机的边距（页面边缘的不可打印区域）。不可打印区域由 PHYSICALOFFSETX 和 PHYSICALOFFSETY 表示。
      Offset += 16
      NumPut("Int", 0, "Int", 0, "Int", LPW, "Int", LPH, FR, Offset)
      ; 确定打印范围。
      ; 如果选择了"选择"选项，则使用选定的文本，否则使用整个文档。
      If (Flags & PD_SELECTION)
         PrintS := Sel.S, PrintE := Sel.E
      Else
         PrintS := 0, PrintE := -1            ; (-1 = 全选)
      Offset += 16
      Numput("Int", PrintS, "Int", PrintE, FR, OffSet) ; cr.cpMin , cr.cpMax
      ; 定义/填充 DOCINFO 结构
      DI := Buffer(A_PtrSize * 5, 0)
      NumPut("UPtr", A_PtrSize * 5, "UPtr", StrPtr(DocName), "UPtr", 0, DI) ; lpszDocName, lpszOutput
      ; 编程注意：所有其他 DOCINFO 字段故意保留为 null。
      ; 确定 MaxPrintIndex
      If (Flags & PD_SELECTION)
          PrintM := Sel.E
      Else
          PrintM := This.GetTextLen()
      ; 确保打印机设备上下文处于文本模式
      DllCall("SetMapMode", "Ptr", PDC, "Int", MM_TEXT)
      ; ----------------------------------------------------------------------------------------------------------------
      ; 打印它！
      ; 开始打印作业。如果有问题，请跳出。
      PrintJob := DllCall("StartDoc", "Ptr", PDC, "Ptr", DI.Ptr, "Int")
        If (PrintJob <= 0)
           Throw Error("函数: " . A_ThisFunc . " - 'StartDoc' 的 DLLCall 失败.", -1)
        ; 打印页面循环
        PageC  := 0 ; 当前页面
        PrintC := 0 ; 当前打印索引
      While (PrintC < PrintM) {
           PageC++
           ; 我们完成了吗？
           If (PageC > PageL)
              Break
           If (PageC >= PageF) && (PageC <= PageL) {
              ; 开始页面函数。如果有问题，请跳出。
              If (DllCall("StartPage", "ptr", PDC, "Int") <= 0) {
                 ErrorMsg := "函数: " . A_ThisFunc . " - 'StartPage' 的 DLLCall 失败."
                 Break
              }
           }
         ; 格式化或测量页面
           If (PageC >= PageF) && (PageC <= PageL)
              Render := True
           Else
              Render := False
           PrintC := SendMessage(0x0439, Render, FR.Ptr, This.HWND)
         If (PageC >= PageF) && (PageC <= PageL) {
              ; 结束页面函数。如果有问题，请跳出。
              If (DllCall("EndPage", "Ptr", PDC, "Int") <= 0) {
                 ErrorMsg := "函数: " . A_ThisFunc . " - 'EndPage' 的 DLLCall 失败."
                 Break
              }
           }
         ; 为下一页更新 FR
           Offset := (A_PtrSize * 2) + (4 * 8)
           Numput("Int", PrintC, "Int", PrintE, FR, Offset) ; cr.cpMin, cr.cpMax
      }
      ; ----------------------------------------------------------------------------------------------------------------
      ; 结束打印作业
      DllCall("EndDoc", "Ptr", PDC)
      ; 删除打印机设备上下文
      DllCall("DeleteDC", "Ptr", PDC)
      ; 重置控件 (释放缓存信息)
      SendMessage(0x0439, 0, 0, This.HWND)
      ; 返回给调用者
      If (ErrorMsg)
         Throw Error(ErrorMsg, -1)
      Return True
   }
   ; -------------------------------------------------------------------------------------------------------------------
   GetMargins() { ; 获取默认打印边距
      Static PSD_RETURNDEFAULT := 0x00000400, PSD_INTHOUSANDTHSOFINCHES := 0x00000004
           , I := 1000 ; 千分之一英寸
           , M := 2540 ; 百分之一毫米
           , PSD_Size := (4 * 10) + (A_PtrSize * 11)
           , PD_Size := (A_PtrSize = 8 ? (13 * A_PtrSize) + 16 : 66)
           , OffFlags := 4 * A_PtrSize
           , OffMargins := OffFlags + (4 * 7)
      ; 检查对象是否已有 Margins 属性
      If !This.HasOwnProp("Margins") {
         PSD := Buffer(PSD_Size, 0) ; PAGESETUPDLG 结构
         ; 设置 PAGESETUPDLG 结构的大小
           NumPut("UInt", PSD_Size, PSD)
           ; 设置标志以返回默认值
           NumPut("UInt", PSD_RETURNDEFAULT, PSD, OffFlags)
         ; 调用 PageSetupDlg 函数
           If !DllCall("Comdlg32.dll\PageSetupDlg", "Ptr", PSD, "UInt")
              Return false
         ; 释放全局内存
           DllCall("GlobalFree", "UInt", NumGet(PSD, 2 * A_PtrSize, "UPtr"))
           ; 释放全局内存
           DllCall("GlobalFree", "UInt", NumGet(PSD, 3 * A_PtrSize, "UPtr"))
         ; 获取标志
           Flags := NumGet(PSD, OffFlags, "UInt")
           ; 确定度量单位（千分之一英寸或百分之一毫米）
           Metrics := (Flags & PSD_INTHOUSANDTHSOFINCHES) ? I : M
         ; 设置边距偏移量
           Offset := OffMargins
           ; 创建边距对象
           This.Margins := {}
           ; 获取左边距
           This.Margins.L := NumGet(PSD, Offset += 0, "Int")           ; 左
           ; 获取上边距
           This.Margins.T := NumGet(PSD, Offset += 4, "Int")           ; 上
           ; 获取右边距
           This.Margins.R := NumGet(PSD, Offset += 4, "Int")           ; 右
           ; 获取下边距
           This.Margins.B := NumGet(PSD, Offset += 4, "Int")           ; 下
         ; 转换左边距为缇（1 英寸 = 1440 缇）
           This.Margins.LT := Round((This.Margins.L / Metrics) * 1440) ; 左边距（缇）
           ; 转换上边距为缇
           This.Margins.TT := Round((This.Margins.T / Metrics) * 1440) ; 上边距（缇）
         This.Margins.RT := Round((This.Margins.R / Metrics) * 1440) ; 右边距（缇）
         This.Margins.BT := Round((This.Margins.B / Metrics) * 1440) ; 下边距（缇）
      }
      Return True
   }
   ; -------------------------------------------------------------------------------------------------------------------
   GetPrinterCaps(DC) { ; 获取打印机的容量
      Static HORZRES         := 0x08, VERTRES         := 0x0A
           , LOGPIXELSX      := 0x58, LOGPIXELSY      := 0x5A
           , PHYSICALWIDTH   := 0x6E, PHYSICALHEIGHT  := 0x6F
           , PHYSICALOFFSETX := 0x70, PHYSICALOFFSETY := 0x71
      Caps := {}
      ; 沿页面宽度和高度每逻辑英寸的像素数
      LPXX := DllCall("GetDeviceCaps", "Ptr", DC, "Int", LOGPIXELSX, "Int")
      LPXY := DllCall("GetDeviceCaps", "Ptr", DC, "Int", LOGPIXELSY, "Int")
      ; 物理页面的宽度和高度，以缇为单位。
      Caps.PHYW := Round((DllCall("GetDeviceCaps", "Ptr", DC, "Int", PHYSICALWIDTH, "Int") / LPXX) * 1440)
      Caps.PHYH := Round((DllCall("GetDeviceCaps", "Ptr", DC, "Int", PHYSICALHEIGHT, "Int") / LPXY) * 1440)
      ; 从物理页面的左/右边缘（PHYSICALOFFSETX）和上/下边缘（PHYSICALOFFSETY）到可打印区域边缘的距离，以缇为单位。
      Caps.POFX := Round((DllCall("GetDeviceCaps", "Ptr", DC, "Int", PHYSICALOFFSETX, "Int") / LPXX) * 1440)
      Caps.POFY := Round((DllCall("GetDeviceCaps", "Ptr", DC, "Int", PHYSICALOFFSETY, "Int") / LPXY) * 1440)
      ; 页面可打印区域的宽度和高度，以缇为单位。
      Caps.HRES := Round((DllCall("GetDeviceCaps", "Ptr", DC, "Int", HORZRES, "Int") / LPXX) * 1440)
      Caps.VRES := Round((DllCall("GetDeviceCaps", "Ptr", DC, "Int", VERTRES, "Int") / LPXY) * 1440)
      Return Caps
   }
   ; ===================================================================================================================
   ; 内部使用的类 *
   ; ===================================================================================================================
   ; CHARFORMAT2 结构 -> docs.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-charformat2w_1
   Class CHARFORMAT2 Extends Buffer {
      Size {
         Get => NumGet(This, 0, "UInt")
         Set => NumPut("UInt", Value, This, 0)
      }
      Mask {
         Get => NumGet(This, 4, "UInt")
         Set => NumPut("UInt", Value, This, 4)
      }
      Effects {
         Get => NumGet(This, 8, "UInt")
         Set => NumPut("UInt", Value, This, 8)
      }
      Height {
         Get => NumGet(This, 12, "Int")
         Set => NumPut("Int", Value, This, 12)
      }
      Offset {
         Get => NumGet(This, 16, "Int")
         Set => NumPut("Int", Value, This, 16)
      }
      TextColor {
         Get => NumGet(This, 20, "UInt")
         Set => NumPut("UInt", Value, This, 20)
      }
      CharSet {
         Get => NumGet(This, 24, "UChar")
         Set => NumPut("UChar", Value, This, 24)
      }
      PitchAndFamily {
         Get => NumGet(This, 25, "UChar")
         Set => NumPut("UChar", Value, This, 25)
      }
      FaceName {
         Get => StrGet(This.Ptr + 26, 32)
         Set => StrPut(Value, This.Ptr + 26, 32)
      }
      Weight {
         Get => NumGet(This, 90, "UShort")
         Set => NumPut("UShort", Value, This, 90)
      }
      Spacing {
         Get => NumGet(This, 92, "Short")
         Set => NumPut("Short", Value, This, 92)
      }
      BackColor {
         Get => NumGet(This, 96, "UInt")
         Set => NumPut("UInt", Value, This, 96)
      }
      LCID {
         Get => NumGet(This, 100, "UInt")
         Set => NumPut("UInt", Value, This, 100)
      }
      Cookie {
         Get => NumGet(This, 104, "UInt")
         Set => NumPut("UInt", Value, This, 104)
      }
      Style {
         Get => NumGet(This, 108, "Short")
         Set => NumPut("Short", Value, This, 108)
      }
      Kerning {
         Get => NumGet(This, 110, "UShort")
         Set => NumPut("UShort", Value, This, 110)
      }
      UnderlineType {
         Get => NumGet(This, 112, "UChar")
         Set => NumPut("UChar", Value, This, 112)
      }
      Animation {
         Get => NumGet(This, 113, "UChar")
         Set => NumPut("UChar", Value, This, 113)
      }
      RevAuthor {
         Get => NumGet(This, 114, "UChar")
         Set => NumPut("UChar", Value, This, 114)
      }
      UnderlineColor {
         Get => NumGet(This, 115, "UChar")
         Set => NumPut("UChar", Value, This, 115)
      }
      ; ----------------------------------------------------------------------------------------------------------------
      __New() {
         Static CF2_Size := 116
         Super.__New(CF2_Size, 0)
         This.Size := CF2_Size
      }
   }
   ; -------------------------------------------------------------------------------------------------------------------
   ; PARAFORMAT2 结构 -> docs.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-paraformat2_1
   Class PARAFORMAT2 Extends Buffer {
      Size {
         Get => NumGet(This, 0, "UInt")
         Set => NumPut("UInt", Value, This, 0)
      }
      Mask {
         Get => NumGet(This, 4, "UInt")
         Set => NumPut("UInt", Value, This, 4)
      }
      Numbering {
         Get => NumGet(This, 8, "UShort")
         Set => NumPut("UShort", Value, This, 8)
      }
      StartIndent {
         Get => NumGet(This, 12, "Int")
         Set => (NumPut("Int", Value, This, 12), Value)
      }
      RightIndent {
         Get => NumGet(This, 16, "Int")
         Set => NumPut("Int", Value, This, 16)
      }
      Offset {
         Get => NumGet(This, 20, "Int")
         Set => NumPut("Int", Value, This, 20)
      }
      Alignment {
         Get => NumGet(This, 24, "UShort")
         Set => NumPut("UShort", Value, This, 24)
      }
      TabCount => NumGet(This, 26, "UShort")
      Tabs {
         Get {
            TabCount := This.TabCount
            Addr := This.Ptr + 28 - 4
            Tabs := Array()
            Tabs.Length := TabCount
            Loop TabCount
               Tabs[A_Index] := NumGet(Addr += 4, "UInt")
            Return Tabs
         }
         Set {
            Static ErrMsg := "Requires a value of type Array but got type "
            If (Type(Value) != "Array")
               Throw TypeError(ErrMsg . Type(Value) . "!", -1)
            DllCall("RtlZeroMemory", "Ptr", This.Ptr + 28, "Ptr", 128)
            TabCount := Value.Length
            Addr := This.Ptr + 28
            For I, Tab In Value
               Addr := NumPut("UInt", Tab, Addr)
            NumPut("UShort", TabCount, This, 26)
            Return Value
         }
      }
      SpaceBefore {
         Get => NumGet(This, 156, "Int")
         Set => NumPut("Int", Value, This, 156)
      }
      SpaceAfter {
         Get => NumGet(This, 160, "Int")
         Set => NumPut("Int", Value, This, 160)
      }
      LineSpacing {
         Get => NumGet(This, 164, "Int")
         Set => NumPut("Int", Value, This, 164)
      }
      Style {
         Get => NumGet(This, 168, "Short")
         Set => NumPut("Short", Value, This, 168)
      }
      LineSpacingRule {
         Get => NumGet(This, 170, "UChar")
         Set => NumPut("UChar", Value, This, 170)
      }
      OutlineLevel {
         Get => NumGet(This, 171, "UChar")
         Set => NumPut("UChar", Value, This, 171)
      }
      ShadingWeight {
         Get => NumGet(This, 172, "UShort")
         Set => NumPut("UShort", Value, This, 172)
      }
      ShadingStyle {
         Get => NumGet(This, 174, "UShort")
         Set => NumPut("UShort", Value, This, 174)
      }
      NumberingStart {
         Get => NumGet(This, 176, "UShort")
         Set => NumPut("UShort", Value, This, 176)
      }
      NumberingStyle {
         Get => NumGet(This, 178, "UShort")
         Set => NumPut("UShort", Value, This, 178)
      }
      NumberingTab {
         Get => NumGet(This, 180, "UShort")
         Set => NumPut("UShort", Value, This, 180)
      }
      BorderSpace {
         Get => NumGet(This, 182, "UShort")
         Set => NumPut("UShort", Value, This, 182)
      }
      BorderWidth {
         Get => NumGet(This, 184, "UShort")
         Set => NumPut("UShort", Value, This, 184)
      }
      Borders {
         Get => NumGet(This, 186, "UShort")
         Set => NumPut("UShort", Value, This, 186)
      }
      ; ----------------------------------------------------------------------------------------------------------------
      __New() {
         Static PF2_Size := 188
         Super.__New(PF2_Size, 0)
         This.Size := PF2_Size
      }
   }
   ; ===================================================================================================================
   ; 内部调用的方法 *
   ; ===================================================================================================================
   GetBGR(RGB) { ; 从数字 RGB 值或 HTML 颜色名称获取数字 BGR 值
      Static HTML := {BLACK:  0x000000, SILVER: 0xC0C0C0, GRAY:   0x808080, WHITE:   0xFFFFFF
                    , MAROON: 0x000080, RED:    0x0000FF, PURPLE: 0x800080, FUCHSIA: 0xFF00FF
                    , GREEN:  0x008000, LIME:   0x00FF00, OLIVE:  0x008080, YELLOW:  0x00FFFF
                    , NAVY:   0x800000, BLUE:   0xFF0000, TEAL:   0x808000, AQUA:    0xFFFF00}
      If HTML.HasProp(RGB)
         Return HTML.%RGB%
      Return ((RGB & 0xFF0000) >> 16) + (RGB & 0x00FF00) + ((RGB & 0x0000FF) << 16)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   GetRGB(BGR) {  ; 从数字 BGR 值获取数字 RGB 值
      Return ((BGR & 0xFF0000) >> 16) + (BGR & 0x00FF00) + ((BGR & 0x0000FF) << 16)
   }
   ; -------------------------------------------------------------------------------------------------------------------
   GetMeasurement() { ; 获取区域设置度量单位（公制/英寸）
      ; LOCALE_USER_DEFAULT = 0x0400, LOCALE_IMEASURE = 0x0D, LOCALE_RETURN_NUMBER = 0x20000000
      Static Metric := 2.54  ; 厘米
           , Inches := 1.00  ; 英寸
           , Measurement := ""
      If (Measurement = "") {
         LCD := Buffer(4, 0)
         DllCall("GetLocaleInfo", "UInt", 0x400, "UInt", 0x2000000D, "Ptr", LCD, "Int", 2)
         Measurement := NumGet(LCD, 0, "UInt") ? Inches : Metric
      }
      Return Measurement
   }
}

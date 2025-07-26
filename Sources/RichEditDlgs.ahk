#DllLoad "Comdlg32.dll"
Class RichEditDlgs {
   Static Call(*) => False
   ; ===================================================================================================================
   ; ===================================================================================================================
   ; RICHEDIT 通用对话框 ===============================================================================================
   ; ===================================================================================================================
   Static FindReplMsg := DllCall("RegisterWindowMessage", "Str", "commdlg_FindReplace", "UInt") ; 查找替换消息
   ; ===================================================================================================================
   ; 以下大多数方法基于 majkinetor 的 DLG 5.01
   ; http://www.autohotkey.com/board/topic/15836-module-dlg-501/
   ; ===================================================================================================================
   Static ChooseColor(RE, Color := "") { ; 选择颜色对话框
   ; ===================================================================================================================
      ; RE : RichEdit 对象
      Static CC_Size := A_PtrSize * 9, CCU := Buffer(64, 0)
      GuiHwnd := RE.Gui.Hwnd
      If (Color != "")
         Color := RE.GetBGR(Color)
      Else
         Color := 0x000000
      CC :=  Buffer(CC_Size, 0)                    ; CHOOSECOLOR 结构
      NumPut("UInt", CC_Size, CC, 0)               ; 结构大小
      NumPut("UPtr", GuiHwnd, CC, A_PtrSize)       ; 所有者窗口句柄，使对话框成为模态
      NumPut("UInt", Color, CC, A_PtrSize * 3)     ; 结果颜色
      NumPut("UPtr", CCU.Ptr, CC, A_PtrSize * 4)   ; 自定义颜色数组指针 (16个)
      NumPut("UInt", 0x0101, CC, A_PtrSize * 5)    ; 标志: 允许任何颜色 | 初始化RGB | ; 完全展开
      R := DllCall("Comdlg32.dll\ChooseColor", "Ptr", CC.Ptr, "UInt")
      Return (R = 0) ? "" : RE.GetRGB(NumGet(CC, A_PtrSize * 3, "UInt"))
   }
   ; ===================================================================================================================
   Static ChooseFont(RE) { ; 选择字体对话框
   ; ===================================================================================================================
      ; RE : RichEdit 对象
      DC := DllCall("GetDC", "Ptr", RE.Gui.Hwnd, "Ptr")
      LP := DllCall("GetDeviceCaps", "Ptr", DC, "UInt", 90, "Int")   ; 垂直逻辑像素数
      DllCall("ReleaseDC", "Ptr", RE.Gui.Hwnd, "Ptr", DC)
      ; 获取当前字体
      Font := RE.GetFont()
      ; LF_FACENAME = 32
      LF := Buffer(92, 0)                   ; LOGFONT 结构
      Size := -(Font.Size * LP / 72)
      NumPut("Int", Size, LF, 0)            ; 字体高度
      If InStr(Font.Style, "B")
         NumPut("Int", 700, LF, 16)         ; 字体粗细
      If InStr(Font.Style, "I")
         NumPut("UChar", 1, LF, 20)         ; 斜体
      If InStr(Font.Style, "U")
         NumPut("UChar", 1, LF, 21)         ; 下划线
      If InStr(Font.Style, "S")
         NumPut("UChar", 1, LF, 22)         ; 删除线
      NumPut("UChar", Font.CharSet, LF, 23) ; 字符集
      StrPut(Font.Name, LF.Ptr + 28, 32)
      ; CF_BOTH = 3, CF_INITTOLOGFONTSTRUCT = 0x40, CF_EFFECTS = 0x100, CF_SCRIPTSONLY = 0x400
      ; CF_NOVECTORFONTS = 0x800, CF_NOSIMULATIONS = 0x1000, CF_LIMITSIZE = 0x2000, CF_WYSIWYG = 0x8000
      ; CF_TTONLY = 0x40000, CF_FORCEFONTEXIST =0x10000, CF_SELECTSCRIPT = 0x400000
      ; CF_NOVERTFONTS =0x01000000
      Flags := 0x00002141 ; 0x01013940
      If (Font.Color = "Auto")
         Color := DllCall("GetSysColor", "Int", 8, "UInt") ; 窗口文本颜色 = 8
      Else
         Color := RE.GetBGR(Font.Color)
      CF_Size := (A_PtrSize = 8 ? (A_PtrSize * 10) + (4 * 4) + A_PtrSize : (A_PtrSize * 14) + 4)
      CF := Buffer(CF_Size, 0)                           ; CHOOSEFONT 结构
      NumPut("UInt", CF_Size, CF)                        ; 结构大小
      NumPut("UPtr", RE.Gui.Hwnd, CF, A_PtrSize)	      ; 所有者窗口句柄 (使对话框成为模态)
      NumPut("UPtr", LF.Ptr, CF, A_PtrSize * 3)	         ; 指向LOGFONT的指针
      NumPut("UInt", Flags, CF, (A_PtrSize * 4) + 4)     ; 标志
      NumPut("UInt", Color, CF, (A_PtrSize * 4) + 8)     ; 颜色
      OffSet := (A_PtrSize = 8 ? (A_PtrSize * 11) + 4 : (A_PtrSize * 12) + 4)
      NumPut("Int", 4, CF, Offset)                       ; 最小尺寸
      NumPut("Int", 160, CF, OffSet + 4)                 ; 最大尺寸
      ; 调用 ChooseFont 对话框
      If !DllCall("Comdlg32.dll\ChooseFont", "Ptr", CF.Ptr, "UInt")
         Return false
      ; 获取名称
      Font.Name := StrGet(LF.Ptr + 28, 32)
   	; 获取大小
   	Font.Size := NumGet(CF, A_PtrSize * 4, "Int") / 10
      ; 获取样式
   	Font.Style := ""
   	If NumGet(LF, 16, "Int") >= 700
   	   Font.Style .= "B"
   	If NumGet(LF, 20, "UChar")
         Font.Style .= "I"
   	If NumGet(LF, 21, "UChar")
         Font.Style .= "U"
   	If NumGet(LF, 22, "UChar")
         Font.Style .= "S"
      OffSet := A_PtrSize * (A_PtrSize = 8 ? 11 : 12)
      FontType := NumGet(CF, Offset, "UShort")
      If (FontType & 0x0100) && !InStr(Font.Style, "B") ; 粗体字体类型
         Font.Style .= "B"
      If (FontType & 0x0200) && !InStr(Font.Style, "I") ; 斜体字体类型
         Font.Style .= "I"
      If (Font.Style = "")
         Font.Style := "N"
      ; 获取字符集
      Font.CharSet := NumGet(LF, 23, "UChar")
      ; 我们不使用字体对话框的有限颜色
      ; 返回选定的值
      Return RE.SetFont(Font)
   }
   ; ===================================================================================================================
   Static FileDlg(RE, Mode, File := "") { ; 打开和另存为对话框
   ; ===================================================================================================================
      ; RE   : RichEdit 对象
      ; Mode : O = 打开, S = 保存
      ; File : 可选的文件名
   	Static OFN_ALLOWMULTISELECT := 0x200,    OFN_EXTENSIONDIFFERENT := 0x400, OFN_CREATEPROMPT := 0x2000,
             OFN_DONTADDTORECENT := 0x2000000, OFN_FILEMUSTEXIST := 0x1000,     OFN_FORCESHOWHIDDEN := 0x10000000,
             OFN_HIDEREADONLY := 0x4,          OFN_NOCHANGEDIR := 0x8,          OFN_NODEREFERENCELINKS := 0x100000,
             OFN_NOVALIDATE := 0x100,          OFN_OVERWRITEPROMPT := 0x2,      OFN_PATHMUSTEXIST := 0x800,
             OFN_READONLY := 0x1,              OFN_SHOWHELP := 0x10,            OFN_NOREADONLYRETURN := 0x8000,
             OFN_NOTESTFILECREATE := 0x10000,  OFN_ENABLEXPLORER := 0x80000
             OFN_Size := (4 * 5) + (2 * 2) + (A_PtrSize * 16)
      Static FilterN1 := "RichText",   FilterP1 := "*.rtf",
             FilterN2 := "Text",       FilterP2 := "*.txt",
             FilterN3 := "AutoHotkey", FilterP3 := "*.ahk",
             DefExt := "rtf",
             DefFilter := 1
   	SplitPath(File, &Name := "", &Dir := "")
      Flags := OFN_ENABLEXPLORER
      Flags |= Mode = "O" ? OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY
                          : OFN_OVERWRITEPROMPT
   	VarSetStrCapacity(&FileName, 512)
      FileName := Name
   	LenN1 := (StrLen(FilterN1) + 1) * 2, LenP1 := (StrLen(FilterP1) + 1) * 2
   	LenN2 := (StrLen(FilterN2) + 1) * 2, LenP2 := (StrLen(FilterP2) + 1) * 2
   	LenN3 := (StrLen(FilterN3) + 1) * 2, LenP3 := (StrLen(FilterP3) + 1) * 2
      Filter := Buffer(LenN1 + LenP1 + LenN2 + LenP2 + LenN3 + LenP3 + 4, 0)
      Adr := Filter.Ptr
      StrPut(FilterN1, Adr)
      StrPut(FilterP1, Adr += LenN1)
      StrPut(FilterN2, Adr += LenP1)
      StrPut(FilterP2, Adr += LenN2)
      StrPut(FilterN3, Adr += LenP2)
      StrPut(FilterP3, Adr += LenN3)
      OFN := Buffer(OFN_Size, 0)                     ; OPENFILENAME 结构
   	NumPut("UInt", OFN_Size, OFN, 0)                ; 结构大小
      Offset := A_PtrSize
   	NumPut("Ptr", RE.Gui.Hwnd, OFN, Offset)        ; 所有者窗口句柄
      Offset += A_PtrSize * 2
   	NumPut("Ptr", Filter.Ptr, OFN, OffSet)         ; 指向过滤器结构的指针
      OffSet += (A_PtrSize * 2) + 4
      OffFilter := Offset
   	NumPut("UInt", DefFilter, OFN, Offset)         ; 默认过滤器对
      OffSet += 4
   	NumPut("Ptr", StrPtr(FileName), OFN, OffSet)   ; 文件名指针 / 初始化文件名
      Offset += A_PtrSize
   	NumPut("UInt", 512, OFN, Offset)               ; 最大文件长度 / 文件名指针长度
      OffSet += A_PtrSize * 3
   	NumPut("Ptr", StrPtr(Dir), OFN, Offset)        ; 起始目录
      Offset += A_PtrSize * 2
   	NumPut("UInt", Flags, OFN, Offset)             ; 标志
      Offset += 8
   	NumPut("Ptr", StrPtr(DefExt), OFN, Offset)     ; 默认扩展名
      R := Mode = "S" ? DllCall("Comdlg32.dll\GetSaveFileNameW", "Ptr", OFN.Ptr, "UInt")
                      : DllCall("Comdlg32.dll\GetOpenFileNameW", "Ptr", OFN.Ptr, "UInt")
   	If !(R)
         Return ""
      DefFilter := NumGet(OFN, OffFilter, "UInt")
   	Return StrGet(StrPtr(FileName))
   }
   ; ===================================================================================================================
   Static FindText(RE) { ; 查找对话框
   ; ===================================================================================================================
      ; RE : RichEdit 对象
   	Static FR_DOWN := 1, FR_MATCHCASE := 4, FR_WHOLEWORD := 2,
   	       Buf := "", BufLen := 256, FR := "", FR_Size := A_PtrSize * 10
      Text := RE.GetSelText()
      Buf := ""
      VarSetStrCapacity(&Buf, BufLen)
      If (Text != "") && !RegExMatch(Text, "\W")
         Buf := Text
      FR := Buffer(FR_Size, 0)
   	NumPut("UInt", FR_Size, FR)                   ; 结构大小
      Offset := A_PtrSize
   	NumPut("UPtr", RE.Gui.Hwnd, FR, Offset)  ; 所有者窗口句柄
      OffSet += A_PtrSize * 2
   	NumPut("UInt", FR_DOWN, FR, Offset)	     ; 标志
      OffSet += A_PtrSize
   	NumPut("UPtr", StrPtr(Buf), FR, Offset)  ; 要查找的文本
      OffSet += A_PtrSize * 2
   	NumPut("Short", BufLen,	FR, Offset)      ; 查找文本长度
      This.FindTextProc("Init", RE.HWND, "")
   	OnMessage(RichEditDlgs.FindReplMsg, RichEditDlgs.FindTextProc)
   	Return DllCall("Comdlg32.dll\FindTextW", "Ptr", FR.Ptr, "UPtr")
   }
   ; -------------------------------------------------------------------------------------------------------------------
   Static FindTextProc(L, M, H) { ; 跳过wParam，当系统调用时可在"This"中找到
      ; 查找对话框回调过程
      ; EM_FINDTEXTEXW = 0x047C, EM_EXGETSEL = 0x0434, EM_EXSETSEL = 0x0437, EM_SCROLLCARET = 0x00B7
      ; FR_DOWN = 1, FR_WHOLEWORD = 2, FR_MATCHCASE = 4,
   	Static FR_DOWN := 1, FR_MATCHCASE := 4, FR_WHOLEWORD := 2 , FR_FINDNEXT := 0x8, FR_DIALOGTERM := 0x40,
             HWND := 0
      If (L = "Init") {
         HWND := M
         Return True
      }
      Flags := NumGet(L, A_PtrSize * 3, "UInt")
      If (Flags & FR_DIALOGTERM) {
         OnMessage(RichEditDlgs.FindReplMsg, RichEditDlgs.FindTextProc, 0)
         If (RE := GuiCtrlFromHwnd(HWND))
            RE.Focus()
         HWND := 0
         Return
      }
      CR := Buffer(8, 0)
      SendMessage(0x0434, 0, CR.Ptr, HWND)  ; 获取选择范围
      Min := (Flags & FR_DOWN) ? NumGet(CR, 4, "Int") : NumGet(CR, 0, "Int")
      Max := (Flags & FR_DOWN) ? -1 : 0
      OffSet := A_PtrSize * 4
      Find := StrGet(NumGet(L, Offset, "UPtr"))  ; 获取要查找的文本
      FTX := Buffer(16 + A_PtrSize, 0)  ; FINDTEXTEXW 结构
      NumPut("Int", Min, "Int", Max, "UPtr", StrPtr(Find), FTX)
      SendMessage(0x047C, Flags, FTX.Ptr, HWND)  ; 查找文本
      S := NumGet(FTX, 8 + A_PtrSize, "Int"), E := NumGet(FTX, 12 + A_PtrSize, "Int")
      If (S = -1) && (E = -1)
         MsgBox("未找到(更多)匹配项!", "查找", 262208)
      Else {
         SendMessage(0x0437, 0, FTX.Ptr + 8 + A_PtrSize, HWND)  ; 设置选择范围
         SendMessage(0x00B7, 0, 0, HWND)  ; 滚动到插入点
      }
   }
   ; ===================================================================================================================
   Static PageSetup(RE) { ; 页面设置对话框
   ; ===================================================================================================================
      ; RE : RichEdit 对象
      ; http://msdn.microsoft.com/en-us/library/ms646842(v=vs.85).aspx
      Static PSD_DEFAULTMINMARGINS             := 0x00000000, ; 默认(打印机的)
             PSD_INWININIINTLMEASURE           := 0x00000000, ; 4种可能之一
             PSD_MINMARGINS                    := 0x00000001, ; 使用调用者的
             PSD_MARGINS                       := 0x00000002, ; 使用调用者的
             PSD_INTHOUSANDTHSOFINCHES         := 0x00000004, ; 4种可能之二
             PSD_INHUNDREDTHSOFMILLIMETERS     := 0x00000008, ; 4种可能之三
             PSD_DISABLEMARGINS                := 0x00000010,
             PSD_DISABLEPRINTER                := 0x00000020,
             PSD_NOWARNING                     := 0x00000080, ; 必须与PD_*相同
             PSD_DISABLEORIENTATION            := 0x00000100,
             PSD_RETURNDEFAULT                 := 0x00000400, ; 必须与PD_*相同
             PSD_DISABLEPAPER                  := 0x00000200,
             PSD_SHOWHELP                      := 0x00000800, ; 必须与PD_*相同
             PSD_ENABLEPAGESETUPHOOK           := 0x00002000, ; 必须与PD_*相同
             PSD_ENABLEPAGESETUPTEMPLATE       := 0x00008000, ; 必须与PD_*相同
             PSD_ENABLEPAGESETUPTEMPLATEHANDLE := 0x00020000, ; 必须与PD_*相同
             PSD_ENABLEPAGEPAINTHOOK           := 0x00040000,
             PSD_DISABLEPAGEPAINTING           := 0x00080000,
             PSD_NONETWORKBUTTON               := 0x00200000, ; 必须与PD_*相同
             I := 1000, ; 千分之一英寸
             M := 2540, ; 百分之一毫米
             Margins := {},
             Metrics := "",
             PSD_Size := (4 * 10) + (A_PtrSize * 11),
             PD_Size := (A_PtrSize = 8 ? (13 * A_PtrSize) + 16 : 66),
             OffFlags := 4 * A_PtrSize,
             OffMargins := OffFlags + (4 * 7)
      PSD := Buffer(PSD_Size, 0)                    ; PAGESETUPDLG 结构
      NumPut("UInt", PSD_Size, PSD)                   ; 结构大小
      NumPut("UPtr", RE.Gui.Hwnd, PSD, A_PtrSize)   ; 所有者窗口句柄
      Flags := PSD_MARGINS | PSD_DISABLEPRINTER | PSD_DISABLEORIENTATION | PSD_DISABLEPAPER
      NumPut("Int", Flags, PSD, OffFlags)           ; 标志
      Offset := OffMargins
      NumPut("Int", RE.Margins.L, PSD, Offset += 0) ; 右边距 左
      NumPut("Int", RE.Margins.T, PSD, Offset += 4) ; 右边距 上
      NumPut("Int", RE.Margins.R, PSD, Offset += 4) ; 右边距 右
      NumPut("Int", RE.Margins.B, PSD, Offset += 4) ; 右边距 下
      If !DllCall("Comdlg32.dll\PageSetupDlg", "Ptr", PSD.Ptr, "UInt")
         Return False
      DllCall("Kernel32.dll\GlobalFree", "Ptr", NumGet(PSD, 2 * A_PtrSize, "UPtr"))
      DllCall("Kernel32.dll\GlobalFree", "Ptr", NumGet(PSD, 3 * A_PtrSize, "UPtr"))
      Flags := NumGet(PSD, OffFlags, "UInt")
      Metrics := (Flags & PSD_INTHOUSANDTHSOFINCHES) ? I : M
      Offset := OffMargins
      RE.Margins.L := NumGet(PSD, Offset += 0, "Int")
      RE.Margins.T := NumGet(PSD, Offset += 4, "Int")
      RE.Margins.R := NumGet(PSD, Offset += 4, "Int")
      RE.Margins.B := NumGet(PSD, Offset += 4, "Int")
      RE.Margins.LT := Round((RE.Margins.L / Metrics) * 1440) ; 左边距(缇)
      RE.Margins.TT := Round((RE.Margins.T / Metrics) * 1440) ; 上边距(缇)
      RE.Margins.RT := Round((RE.Margins.R / Metrics) * 1440) ; 右边距(缇)
      RE.Margins.BT := Round((RE.Margins.B / Metrics) * 1440) ; 下边距(缇)
      Return True
   }
   ; ===================================================================================================================
   Static ReplaceText(RE) { ; 替换对话框
   ; ===================================================================================================================
      ; RE : RichEdit 对象
   	Static FR_DOWN := 1, FR_MATCHCASE := 4, FR_WHOLEWORD := 2,
   	       FBuf := "", RBuf := "", BufLen := 256, FR := "", FR_Size := A_PtrSize * 10
      Text := RE.GetSelText()
      FBuf := RBuf := ""
      VarSetStrCapacity(&FBuf, BufLen)
      If (Text != "") && !RegExMatch(Text, "\W")
         FBuf := Text
      VarSetStrCapacity(&RBuf, BufLen)
      FR := Buffer(FR_Size, 0)
   	NumPut("UInt", FR_Size, FR)                   ; 结构大小
      Offset := A_PtrSize
   	NumPut("UPtr", RE.Gui.Hwnd, FR, Offset)              ; 所有者窗口句柄
      OffSet += A_PtrSize * 2
   	NumPut("UInt", FR_DOWN, FR, Offset)	                 ; 标志
      OffSet += A_PtrSize
   	NumPut("UPtr", StrPtr(FBuf), FR, Offset)             ; 要查找的文本
      OffSet += A_PtrSize
   	NumPut("UPtr", StrPtr(RBuf), FR, Offset)             ; 要替换的文本
      OffSet += A_PtrSize
   	NumPut("Short", BufLen,	"Short", BufLen, FR, Offset) ; 查找文本长度, 替换文本长度
      This.ReplaceTextProc("Init", RE.HWND, "")
   	OnMessage(RichEditDlgs.FindReplMsg, RichEditDlgs.ReplaceTextProc)
   	Return DllCall("Comdlg32.dll\ReplaceText", "Ptr", FR.Ptr, "UPtr")
   }
   ; -------------------------------------------------------------------------------------------------------------------
   Static ReplaceTextProc(L, M, H) { ; 跳过wParam，当系统调用时可在"This"中找到
      ; 替换对话框回调过程
      ; EM_FINDTEXTEXW = 0x047C, EM_EXGETSEL = 0x0434, EM_EXSETSEL = 0x0437
      ; EM_REPLACESEL = 0xC2, EM_SCROLLCARET = 0x00B7
      ; FR_DOWN = 1, FR_WHOLEWORD = 2, FR_MATCHCASE = 4,
   	Static FR_DOWN := 1, FR_MATCHCASE := 4, FR_WHOLEWORD := 2, FR_FINDNEXT := 0x8,
             FR_REPLACE := 0x10, FR_REPLACEALL := 0x20, FR_DIALOGTERM := 0x40,
             HWND := 0, Min := "", Max := "", FS := "", FE := "",
             OffFind := A_PtrSize * 4, OffRepl := A_PtrSize * 5
      If (L = "Init") {
         HWND := M, FS := "", FE := ""
         Return True
      }
      Flags := NumGet(L, A_PtrSize * 3, "UInt")
      If (Flags & FR_DIALOGTERM) {
         OnMessage(RichEditDlgs.FindReplMsg, RichEditDlgs.ReplaceTextProc, 0)
         If (RE :=GuiCtrlFromHwnd(HWND))
            RE.Focus()
         HWND := 0
         Return
      }
      If (Flags & FR_REPLACE) {  ; 如果是替换操作
         IF (FS >= 0) && (FE >= 0) {  ; 如果有选中的文本
            SendMessage(0xC2, 1, NumGet(L, OffRepl, "UPtr"), HWND)  ; 替换选中的文本
            Flags |= FR_FINDNEXT  ; 同时查找下一个
         }
         Else
            Return
      }
      If (Flags & FR_FINDNEXT) {  ; 如果是查找下一个
         CR := Buffer(8, 0)
         SendMessage(0x0434, 0, CR.Ptr, HWND)  ; 获取选择范围
         Min := NumGet(CR, 4)
         FS := FE := ""
         Find := NumGet(L, OffFind, "UPtr")  ; 获取要查找的文本
         FTX := Buffer(16 + A_PtrSize, 0)  ; FINDTEXTEXW 结构
         NumPut("Int", Min, "Int", -1, "Ptr", Find, FTX)
         SendMessage(0x047C, Flags, FTX.Ptr, HWND)  ; 查找文本
         S := NumGet(FTX, 8 + A_PtrSize, "Int"), E := NumGet(FTX, 12 + A_PtrSize, "Int")
         If (S = -1) && (E = -1)
            MsgBox("未找到(更多)匹配项!", "替换", 262208)
         Else {
            SendMessage(0x0437, 0, FTX.Ptr + 8 + A_PtrSize, HWND)  ; 设置选择范围
            SendMessage(0x00B7, 0, 0, HWND)  ; 滚动到插入点
            FS := S, FE := E
         }
         Return
      }
      If (Flags & FR_REPLACEALL) {  ; 如果是全部替换
         CR := Buffer(8, 0)
         SendMessage(0x0434, 0, CR.Ptr, HWND)  ; 获取选择范围
         If (FS = "")
            FS := FE := 0
         DllCall("User32.dll\LockWindowUpdate", "Ptr", HWND)  ; 锁定窗口更新
         Find := NumGet(L, OffFind, "UPtr")  ; 获取要查找的文本
         FTX := Buffer(16 + A_PtrSize, 0)  ; FINDTEXTEXW 结构
         NumPut("Int", FS, "Int", -1, "Ptr", Find, FTX)
         While (FS >= 0) && (FE >= 0) {
            SendMessage(0x044F, Flags, FTX.Ptr, HWND)  ; 查找文本
            FS := NumGet(FTX, A_PtrSize + 8, "Int"), FE := NumGet(FTX, A_PtrSize + 12, "Int")
            If (FS >= 0) && (FE >= 0) {
               SendMessage(0x0437, 0, FTX.Ptr + 8 + A_PtrSize, HWND)  ; 设置选择范围
               SendMessage(0xC2, 1, NumGet(L + 0, OffRepl, "UPtr" ), HWND)  ; 替换文本
               NumPut("Int", FE, FTX)  ; 更新起始位置
            }
         }
         SendMessage(0x0437, 0, CR.Ptr, HWND)  ; 重置选择范围
         DllCall("User32.dll\LockWindowUpdate", "Ptr", 0)  ; 解锁窗口更新
         Return
      }
   }
}

; ======================================================================================================================
; RichEdit 演示
; 这个脚本展示了如何使用 RichEdit 控件创建一个简单的富文本编辑器
; ======================================================================================================================
#Include RichEdit.ahk      ; 包含 RichEdit 控件库
#Include RichEditDlgs.ahk  ; 包含 RichEdit 对话框库
; ======================================================================================================================
SetWinDelay -1  ; 设置窗口延迟为最短
SetControlDelay -1  ; 设置控件延迟为最短
; ======================================================================================================================
; 创建带有 RichEdit 控件的 Gui
; ======================================================================================================================
; ----------------------------------------------------------------------------------------------------------------------
; 菜单
; ----------------------------------------------------------------------------------------------------------------------
; 文件菜单------------------------------------------------------------------------------------------------------------
FileMenu := Menu() ; 创建文件菜单
FileMenu.Add("打开", FileLoadFN.Bind("Open"))     ; 打开文件
FileMenu.Add("追加", FileLoadFN.Bind("Append")) ; 追加文件
FileMenu.Add("插入", FileLoadFN.Bind("Insert")) ; 插入文件
FileMenu.Add("关闭", FileCloseFN)                ; 关闭文件
FileMenu.Add("保存", FileSaveFN)                  ; 保存文件
FileMenu.Add("另存为", FileSaveAsFN)             ; 另存为
FileMenu.Add()                                      ; 分隔线
FileMenu.Add("页面边距", PageSetupFN)             ; 页面边距
FileMenu.Add("打印", PrintFN)                      ; 打印
FileMenu.Add()                                      ; 分隔线
FileMenu.Add("退出", MainGuiClose)                 ; 退出
; 编辑菜单------------------------------------------------------------------------------------------------------------
EditMenu := Menu() ; 创建编辑菜单
EditMenu.Add("撤销`tCtrl+Z", UndoFN)           ; 撤销
EditMenu.Add("重做`tCtrl+Y", RedoFN)           ; 重做
EditMenu.Add()                                   ; 分隔线
EditMenu.Add("剪切`tCtrl+X", CutFN)             ; 剪切
EditMenu.Add("复制`tCtrl+C", CopyFN)           ; 复制
EditMenu.Add("粘贴`tCtrl+V", PasteFN)          ; 粘贴
EditMenu.Add("清除`tDel", ClearFN)              ; 清除
EditMenu.Add()                                   ; 分隔线
EditMenu.Add("全选 `tCtrl+A", SelAllFN)         ; 全选
EditMenu.Add("取消选择", DeselectFN)           ; 取消选择
; 搜索菜单------------------------------------------------------------------------------------------------------------
SearchMenu := Menu() ; 创建搜索菜单
SearchMenu.Add("查找", FindFN)       ; 查找
SearchMenu.Add("替换", ReplaceFN)     ; 替换
; 格式菜单------------------------------------------------------------------------------------------------------------
; 段落
AlignMenu := Menu() ; 创建对齐菜单
AlignMenu.Add("左对齐`tCtrl+L", AlignFN.Bind("Left"))     ; 左对齐
AlignMenu.Add("居中对齐`tCtrl+E", AlignFN.Bind("Center")) ; 居中对齐
AlignMenu.Add("右对齐`tCtrl+R", AlignFN.Bind("Right"))   ; 右对齐
AlignMenu.Add("两端对齐", AlignFN.Bind("Justify"))    ; 两端对齐
IndentMenu := Menu() ; 创建缩进菜单
IndentMenu.Add("设置", IndentationFN.Bind("Set"))     ; 设置缩进
IndentMenu.Add("重置", IndentationFN.Bind("Reset"))   ; 重置缩进
LineSpacingMenu := Menu() ; 创建行间距菜单
LineSpacingMenu.Add("1倍行距`tCtrl+1", SpacingFN.Bind(1.0))   ; 1倍行距
LineSpacingMenu.Add("1.5倍行距`tCtrl+5", SpacingFN.Bind(1.5)) ; 1.5倍行距
LineSpacingMenu.Add("2倍行距`tCtrl+2", SpacingFN.Bind(2.0))   ; 2倍行距
NumberingMenu := Menu() ; 创建编号菜单
NumberingMenu.Add("设置", NumberingFN.Bind("Set"))     ; 设置编号
NumberingMenu.Add("重置", NumberingFN.Bind("Reset"))   ; 重置编号
TabstopsMenu := Menu() ; 创建制表位菜单
TabstopsMenu.Add("设置制表位", SetTabstopsFN.Bind("Set"))       ; 设置制表位
TabstopsMenu.Add("重置为默认", SetTabstopsFN.Bind("Reset"))     ; 重置为默认
TabstopsMenu.Add()                                                  ; 分隔线
TabstopsMenu.Add("设置默认制表位", SetTabstopsFN.Bind("Default")) ; 设置默认制表位
ParaSpacingMenu := Menu() ; 创建段落间距菜单
ParaSpacingMenu.Add("设置", ParaSpacingFN.Bind("Set"))     ; 设置段落间距
ParaSpacingMenu.Add("重置", ParaSpacingFN.Bind("Reset"))   ; 重置段落间距
ParagraphMenu := Menu() ; 创建段落菜单
ParagraphMenu.Add("对齐", AlignMenu)               ; 对齐
ParagraphMenu.Add("缩进", IndentMenu)               ; 缩进
ParagraphMenu.Add("编号", NumberingMenu)           ; 编号
ParagraphMenu.Add("行间距", LineSpacingMenu)       ; 行间距
ParagraphMenu.Add("段落前后间距", ParaSpacingMenu) ; 段落前后间距
ParagraphMenu.Add("制表位", TabstopsMenu)           ; 制表位
; 字符
TxColorMenu := Menu() ; 创建文本颜色菜单
TxColorMenu.Add("选择", TextColorFN.Bind("Choose"))   ; 选择颜色
TxColorMenu.Add("自动", TextColorFN.Bind("Auto"))     ; 自动颜色
BkColorMenu := Menu() ; 创建文本背景颜色菜单
BkColorMenu.Add("选择", TextBkColorFN.Bind("Choose")) ; 选择背景颜色
BkColorMenu.Add("自动", TextBkColorFN.Bind("Auto"))   ; 自动背景颜色
CharacterMenu := Menu() ; 创建字符菜单
CharacterMenu.Add("字体", ChooseFontFN)             ; 字体
CharacterMenu.Add("文本颜色", TxColorMenu)         ; 文本颜色
CharacterMenu.Add("文本背景颜色", BkColorMenu)     ; 文本背景颜色
; 格式
FormatMenu := Menu() ; 创建格式菜单
FormatMenu.Add("字符", CharacterMenu)             ; 字符
FormatMenu.Add("段落", ParagraphMenu)             ; 段落
; 视图菜单------------------------------------------------------------------------------------------------------------
; 背景
BackgroundMenu := Menu() ; 创建背景菜单
BackgroundMenu.Add("选择", BackGroundColorFN.Bind("Choose")) ; 选择背景颜色
BackgroundMenu.Add("自动", BackgroundColorFN.Bind("Auto"))   ; 自动背景颜色
; 缩放
ZoomMenu := Menu() ; 创建缩放菜单
ZoomMenu.Add("200 %", ZoomFN.Bind(200)) ; 200% 缩放
ZoomMenu.Add("150 %", ZoomFN.Bind(150)) ; 150% 缩放
ZoomMenu.Add("125 %", ZoomFN.Bind(125)) ; 125% 缩放
ZoomMenu.Add("100 %", Zoom100FN)        ; 100% 缩放
ZoomMenu.Check("100 %")                 ; 勾选 100%
ZoomMenu.Add("75 %", ZoomFN.Bind(75))   ; 75% 缩放
ZoomMenu.Add("50 %", ZoomFN.Bind(50))   ; 50% 缩放
; 视图
ViewMenu := Menu() ; 创建视图菜单
MenuWordWrap := "自动换行"                       ; 自动换行
ViewMenu.Add(MenuWordWrap, WordWrapFN)           ; 添加自动换行菜单项
MenuWysiwyg := "按打印方式换行"                  ; 按打印方式换行
ViewMenu.Add(MenuWysiwyg, WysiWygFN)             ; 添加按打印方式换行菜单项
ViewMenu.Add("缩放", ZoomMenu)                  ; 缩放
ViewMenu.Add()                                   ; 分隔线
ViewMenu.Add("背景颜色", BackgroundMenu)       ; 背景颜色
ViewMenu.Add("URL 检测", AutoURLDetectionFN)   ; URL 检测
; 上下文菜单 ----------------------------------------------------------------------------------------------------------
ContextMenu := Menu() ; 创建上下文菜单
ContextMenu.Add("文件", FileMenu)     ; 文件
ContextMenu.Add("编辑", EditMenu)     ; 编辑
ContextMenu.Add("搜索", SearchMenu)   ; 搜索
ContextMenu.Add("格式", FormatMenu)   ; 格式
ContextMenu.Add("视图", ViewMenu)     ; 视图
; 主菜单栏 ------------------------------------------------------------------------------------------------------------
MainMenuBar := MenuBar() ; 创建主菜单栏
MainMenuBar.Add("文件", FileMenu)     ; 文件
MainMenuBar.Add("编辑", EditMenu)     ; 编辑
MainMenuBar.Add("搜索", SearchMenu)   ; 搜索
MainMenuBar.Add("格式", FormatMenu)   ; 格式
MainMenuBar.Add("视图", ViewMenu)     ; 视图
; 主 GUI ==============================================================================================================
MainGui := Gui("+ReSize +MinSize", "简易富文本编辑器") ; 创建主 GUI
MainGui.OnEvent("Size", MainGuiSize)        ; 绑定大小事件
MainGui.OnEvent("Close", MainGuiClose)      ; 绑定关闭事件
MainGui.OnEvent("ContextMenu", MainContextMenu) ; 绑定上下文菜单事件
MainGui.MenuBar := MainMenuBar               ; 设置菜单栏
; 样式按钮----------------------------------------------------------------------------------
MainGui.SetFont("Bold", "Arial") ; 设置粗体字体
MainBNSB := MainGui.AddButton("xm y3 w20 h20", "&B") ; 添加粗体按钮
MainBNSB.OnEvent("Click", SetFontStyleFN.Bind("B")) ; 绑定粗体按钮点击事件
GuiCtrlSetTip(MainBNSB, "粗体 (Alt+B)") ; 设置粗体按钮提示
MainGui.SetFont("Norm Italic") ; 设置斜体字体
MainBNSI := MainGui.AddButton("x+0 yp wp hp", "&I") ; 添加斜体按钮
MainBNSI.OnEvent("Click", SetFontStyleFN.Bind("I")) ; 绑定斜体按钮点击事件
GuiCtrlSetTip(MainBNSI, "斜体 (Alt+I)") ; 设置斜体按钮提示
MainGui.SetFont("Norm Underline") ; 设置下划线字体
MainBNSU := MainGui.AddButton("x+0 yp wp hp", "&U") ; 添加下划线按钮
MainBNSU.OnEvent("Click", SetFontStyleFN.Bind("U")) ; 绑定下划线按钮点击事件
GuiCtrlSetTip(MainBNSU, "下划线 (Alt+U)") ; 设置下划线按钮提示
MainGui.SetFont("Norm Strike") ; 设置删除线字体
MainBNSS := MainGui.AddButton("x+0 yp wp hp", "&S") ; 添加删除线按钮
MainBNSS.OnEvent("Click", SetFontStyleFN.Bind("S")) ; 绑定删除线按钮点击事件
GuiCtrlSetTip(MainBNSS, "删除线 (Alt+S)") ; 设置删除线按钮提示
MainGui.SetFont("Norm", "Arial") ; 设置正常字体
MainBNSH := MainGui.AddButton("x+0 yp wp hp", "¯") ; 添加上标按钮
MainBNSH.OnEvent("Click", SetFontStyleFN.Bind("H")) ; 绑定上标按钮点击事件
GuiCtrlSetTip(MainBNSH, "上标 (Ctrl+Shift+'+')") ; 设置上标按钮提示
MainBNSL := MainGui.AddButton("x+0 yp wp hp", "_") ; 添加下标按钮
MainBNSL.OnEvent("Click", SetFontStyleFN.Bind("L")) ; 绑定下标按钮点击事件
GuiCtrlSetTip(MainBNSL, "下标 (Ctrl+'+')") ; 设置下标按钮提示
MainBNSN := MainGui.AddButton("x+0 yp wp hp", "&N") ; 添加正常按钮
MainBNSN.OnEvent("Click", SetFontStyleFN.Bind("N")) ; 绑定正常按钮点击事件
GuiCtrlSetTip(MainBNSN, "正常 (Alt+N)") ; 设置正常按钮提示
MainBNTC := MainGui.AddButton("x+10 yp wp hp", "&T") ; 添加文本颜色按钮
MainBNTC.OnEvent("Click", TextColorFN.Bind("Choose")) ; 绑定文本颜色按钮点击事件
GuiCtrlSetTip(MainBNTC, "文本颜色 (Alt+T)") ; 设置文本颜色按钮提示
MainColors := MainGui.AddProgress("x+0 yp wp hp BackgroundYellow cNavy Border", 50) ; 添加颜色进度条
MainBNBC := MainGui.AddButton("x+0 yp wp hp", "B") ; 添加文本背景颜色按钮
MainBNBC.OnEvent("Click", TextBkColorFN.Bind("Choose")) ; 绑定文本背景颜色按钮点击事件
GuiCtrlSetTip(MainBNBC, "文本背景颜色") ; 设置文本背景颜色按钮提示
MainFNAME := MainGui.AddEdit("x+10 yp w150 hp ReadOnly", "") ; 添加字体名称编辑框
MainBNCF := MainGui.AddButton("x+0 yp w20 hp", "...") ; 添加选择字体按钮
MainBNCF.OnEvent("Click", ChooseFontFN) ; 绑定选择字体按钮点击事件
GuiCtrlSetTip(MainBNCF, "选择字体") ; 设置选择字体按钮提示
MainBNFP := MainGui.AddButton("x+5 yp wp hp", "&+") ; 添加增大字体按钮
MainBNFP.OnEvent("Click", ChangeSize.Bind(1))
GuiCtrlSetTip(MainBNFP, "增大字体 (Alt+'+')") ; 增大字体提示
MainFSIZE := MainGui.AddEdit("x+0 yp w30 hp ReadOnly", "") ; 字体大小编辑框
MainBNFM := MainGui.AddButton("x+5 yp wp hp", "&-") ; 减小字体按钮
MainBNFM.OnEvent("Click", ChangeSize.Bind(-1)) ; 绑定减小字体按钮点击事件
GuiCtrlSetTip(MainBNFM, "减小字体 (Alt+'-')") ; 减小字体提示
; RichEdit #1 ----------------------------------------------------------------------------------------------------------
MainGui.SetFont("Bold Italic", "Arial") ; 设置粗斜体字体
MainGui.SetFont("Norm", "Arial") ; 恢复正常字体
Options := "x+5 yp w80 hp" ; 设置选项
If !IsObject(RE1 := RichEdit(MainGui, Options, False)) ; 创建 RichEdit 控件
   Throw("Could not create the RE1 RichEdit control!", -1) ; 抛出错误
RE1.ReplaceSel("AaBbYyZz") ; 替换选中内容
RE1.AlignText("CENTER") ; 居中对齐文本
RE1.SetOptions(["READONLY"], "SET") ; 设置为只读
RE1.SetParaSpacing({Before: 2}) ; 设置段落间距
; 对齐和行间距 --------------------------------------------------------------------------------------------------------
MainGui.SetFont("Norm", "Arial") ; 设置正常字体
MainGui.AddText("0x1000 xm y+2 h2 w800") ; 添加分隔线
MainBNAL := MainGui.AddButton("x10 y+1 w30 h20",  "|<") ; 左对齐按钮
MainBNAL.OnEvent("Click", AlignFN.Bind("Left")) ; 绑定左对齐按钮点击事件
GuiCtrlSetTip(MainBNAL, "左对齐 (Ctrl+L)") ; 左对齐提示
MainBNAC := MainGui.AddButton("x+0 yp wp hp", "><") ; 居中对齐按钮
MainBNAC.OnEvent("Click", AlignFN.Bind("Center")) ; 绑定居中对齐按钮点击事件
GuiCtrlSetTip(MainBNAC, "居中对齐 (Ctrl+E)") ; 居中对齐提示
MainBNAR := MainGui.AddButton("x+0 yp wp hp", ">|") ; 右对齐按钮
MainBNAR.OnEvent("Click", AlignFN.Bind("Right")) ; 绑定右对齐按钮点击事件
GuiCtrlSetTip(MainBNAR, "右对齐 (Ctrl+R)") ; 右对齐提示
MainBNAJ := MainGui.AddButton("x+0 yp wp hp", "|<>|") ; 两端对齐按钮
MainBNAJ.OnEvent("Click", AlignFN.Bind("Justify")) ; 绑定两端对齐按钮点击事件
GuiCtrlSetTip(MainBNAJ, "两端对齐") ; 两端对齐提示
MainBN10 := MainGui.AddButton("x+10 yp wp hp", "1") ; 1倍行距按钮
MainBN10.OnEvent("Click", SpacingFN.Bind(1.0)) ; 绑定1倍行距按钮点击事件
GuiCtrlSetTip(MainBN10, "1倍行距 (Ctrl+1)") ; 1倍行距提示
MainBN15 := MainGui.AddButton("x+0 yp wp hp", "1½") ; 1.5倍行距按钮
MainBN15.OnEvent("Click", SpacingFN.Bind(1.5)) ; 绑定1.5倍行距按钮点击事件
GuiCtrlSetTip(MainBN15, "1.5倍行距 (Ctrl+5)") ; 1.5倍行距提示
MainBN20 := MainGui.AddButton("x+0 yp wp hp", "2") ; 2倍行距按钮
MainBN20.OnEvent("Click", SpacingFN.Bind(2.0)) ; 绑定2倍行距按钮点击事件
GuiCtrlSetTip(MainBN20, "2倍行距 (Ctrl+2)") ; 2倍行距提示
; RichEdit #2 ----------------------------------------------------------------------------------------------------------
MainFNAME.Text := "Arial" ; 设置字体名称
MainFSIZE.Text := "10" ; 设置字体大小
MainGui.SetFont("s10", "Arial") ; 设置字体
Options := "xm y+5 w800 r20" ; 设置选项
If !IsObject(RE2 := RichEdit(MainGui, Options)) ; 创建 RichEdit 控件
   Throw("Could not create the RE2 RichEdit control!", -1) ; 抛出错误
; RE2.SetOptions(["SELECTIONBAR"]) ; 设置选择栏
RE2.AutoURL(True) ; 启用自动 URL 检测
RE2.SetEventMask(["SELCHANGE", "LINK"]) ; 设置事件掩码
RE2.OnNotify(0x0702, RE2_SelChange) ; 绑定选择更改通知
RE2.OnNotify(0x070B, RE2_Link) ; 绑定链接通知
RE2.SetBkgndColor(0xCCE8CF) ; 设置富文本编辑控件背景颜色为 204 232 207
RE1.SetBkgndColor(0xCCE8CF) ; 设置富文本编辑控件背景颜色为 204 232 207
RE2.SetFont({Name: "微软雅黑",Color: 0x008000,  Size: 20, Default: ""}) ; 设置默认颜色为 0 128 0，字体 微软雅黑 大小 20
WordWrapFN(MenuWordWrap) ; 调用 WordWrapFN 函数切换自动换行及其菜单项的选中状态
MainGui.SetFont() ; 重置字体
; 其余部分
MainSB := MainGui.AddStatusbar() ; 添加状态栏
MainSB.SetParts(10, 200) ; 设置状态栏部分
MainGui.Show() ; 显示 GUI
RE2.Focus() ; 设置焦点到 RichEdit 控件
UpdateGui() ; 更新 GUI
Return ; 返回
; ======================================================================================================================
; 自动执行部分结束
; ======================================================================================================================
; ----------------------------------------------------------------------------------------------------------------------
RE2_SelChange(RE, L) {
   SetTimer UpdateGui, -10
}
RE2_Link(RE, L) {
   If (NumGet(L, A_PtrSize * 3, "Int") = 0x0202) { ; WM_LBUTTONUP
      wParam  := NumGet(L, (A_PtrSize * 3) + 4, "UPtr")
      lParam  := NumGet(L, (A_PtrSize * 4) + 4, "UPtr")
      cpMin   := NumGet(L, (A_PtrSize * 5) + 4, "Int")
      cpMax   := NumGet(L, (A_PtrSize * 5) + 8, "Int")
      URLtoOpen := RE2.GetTextRange(cpMin, cpMax)
      ToolTip "0x0202 - " wParam " - " lParam " - " cpMin " - " cpMax " - " URLtoOpen
      Run '"' URLtoOpen '"'
   }
}
; ----------------------------------------------------------------------------------------------------------------------
; 更新界面:
UpdateGui(*) {
   Static FontName := "", FontCharset := 0, FontStyle := 0, FontSize := 0, TextColor := 0, TxBkColor := 0
   Local Font := RE2.GetFont()
   If (FontName != Font.Name || FontCharset != Font.CharSet || FontStyle != Font.Style || FontSize != Font.Size ||
      TextColor != Font.Color || TxBkColor != Font.BkColor) {
      FontStyle := Font.Style
      TextColor := Font.Color
      TxBkColor := Font.BkColor
      FontCharSet := Font.CharSet
      If (FontName != Font.Name) {
         FontName := Font.Name
         MainFNAME.Text := FontName
      }
      If (FontSize != Font.Size) {
         FontSize := Round(Font.Size)
         MainFSIZE.Text := FontSize
      }
      Font.Size := 8
      RE1.SetSel(0, -1) ; 选择全部
      RE1.SetFont(Font)
      RE1.SetSel(0, 0)  ; 取消选择
   }
   Local Stats := RE2.GetStatistics()
   MainSB.SetText(Stats.Line . " : " . Stats.LinePos . " (" . Stats.LineCount . ")  [" . Stats.CharCount . "]", 2)
}
; ======================================================================================================================
; GUI 相关
; ======================================================================================================================
; 窗口关闭:
MainGuiClose(*) {
   Global RE1, RE2
   If IsObject(RE1)
      RE1 := ""
   If IsObject(RE2)
      RE2 := ""
   MainGui.Destroy()
   ExitApp
}
; ----------------------------------------------------------------------------------------------------------------------
; 窗口大小:
MainGuiSize(GuiObj, MinMax, Width, Height) {
   InitDiffs() {
      RE2.GetPos( , , &CurREW, &CurREH)  ; 获取当前编辑控件宽高
      Return {WidthDiff: Width - CurREW, HeightDiff: Height - CurREH}  ; 返回包含两个差值的对象
   }
   If (MinMax = -1)
      Return
   Static Diffs := InitDiffs()
   RE2.Move( , , Width - Diffs.WidthDiff, Height - Diffs.HeightDiff)  ; 调整编辑控件大小
}
; ----------------------------------------------------------------------------------------------------------------------
; GuiContextMenu
MainContextMenu(GuiObj, GuiCtrlObj, *) {
   If (GuiCtrlObj = RE2)
      ContextMenu.Show()
}
; ======================================================================================================================
; 文本操作
; ======================================================================================================================
; 设置字体样式
SetFontStyleFN(Style,GuiCtrl,*){
   RE2.ToggleFontStyle(Style) ; 切换字体样式
   UpdateGui() ; 更新 GUI
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; 更改大小
ChangeSize(IncDec,GuiCtrl,*){
   Global FontSize := RE2.ChangeFontSize(IncDec) ; 更改字体大小
   MainFSIZE.Text := Round(FontSize) ; 更新字体大小显示
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ======================================================================================================================
; 菜单文件
; ======================================================================================================================
; 文件追加
; 文件打开
; 文件插入
FileLoadFN(Mode,*){
   Global Open_File ; 全局变量，当前打开的文件
   If (File := RichEditDlgs.FileDlg(RE2,"O")){ ; 打开文件对话框
      RE2.LoadFile(File,Mode) ; 加载文件
      If (Mode = "O"){ ; 如果是打开模式
         MainGui.Opt("+LastFound") ; 设置 GUI 为最后找到的窗口
         Title := WinGetTitle() ; 获取窗口标题
         Title := StrSplit(Title,"-"," ") ; 分割标题
         WinSetTitle(Title[1] . " - " . File) ; 设置新的窗口标题
         Open_File := File ; 设置当前打开的文件
      }
      UpdateGui() ; 更新 GUI
   }
   RE2.SetModified() ; 设置修改标志
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; 关闭文件
FileCloseFN(*) {
   Global Open_File ; 全局变量，当前打开的文件
   If (Open_File) { ; 如果有打开的文件
      If RE2.IsModified() { ; 如果内容已修改
         MainGui.Opt("+OwnDialogs") ; 设置 GUI 拥有对话框
         Switch MsgBox(35, "Close File", "Content has been modified!`nDo you want to save changes?") { ; 显示消息框
            Case "Cancel": ; 取消
               RE2.Focus() ; 设置焦点到 RichEdit 控件
               Return ; 返回
            Case "Yes": ; 是
               FileSaveFN() ; 保存文件
         }
      }
      If RE2.SetText() { ; 清空文本
         MainGui.Opt("+LastFound") ; 设置 GUI 为最后找到的窗口
         Title := WinGetTitle() ; 获取窗口标题
         Title := StrSplit(Title, "-", " ") ; 分割标题
         WinSetTitle(Title[1]) ; 设置新的窗口标题
         Open_File := "" ; 清空当前打开的文件
      }
      UpdateGui() ; 更新 GUI
   }
   RE2.SetModified() ; 设置修改标志
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; 保存文件
FileSaveFN(*) {
   If !(Open_File) ; 如果没有打开的文件
      Return FileSaveAsFN() ; 返回到另存为函数
   RE2.SaveFile(Open_File) ; 保存文件
   RE2.SetModified() ; 设置修改标志
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; 另存为:
FileSaveAsFN(*) {
   If (File := RichEditDlgs.FileDlg(RE2, "S")) { ; 打开另存为对话框
      RE2.SaveFile(File) ; 保存文件
      MainGui.Opt("+LastFound") ; 设置 GUI 为最后找到的窗口
      Title := WinGetTitle() ; 获取窗口标题
      Title := StrSplit(Title, "-", " ") ; 分割标题
      WinSetTitle(Title[1] . " - " . File) ; 设置新的窗口标题
      Global Open_File := File ; 设置当前打开的文件
   }
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; 页面设置
PageSetupFN(*) {
   RichEditDlgs.PageSetup(RE2) ; 调用页面设置对话框
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; 打印
PrintFN(*) {
   RE2.Print() ; 打印文档
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ======================================================================================================================
; 菜单编辑
; ======================================================================================================================
; 撤销
UndoFN(*) {
   RE2.Undo() ; 撤销操作
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; 重做
RedoFN(*) {
   RE2.Redo() ; 重做操作
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; 剪切
CutFN(*) {
   RE2.Cut() ; 剪切选中的内容
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; 复制
CopyFN(*) {
   RE2.Copy() ; 复制选中的内容
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; 粘贴:
PasteFN(*) {
   RE2.Paste() ; 粘贴内容
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; 清除
ClearFN(*) {
   RE2.Clear() ; 清除内容
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; 全选
SelAllFN(*) {
   RE2.SelAll() ; 全选内容
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; 取消选择
DeselectFN(*) {
   RE2.Deselect() ; 取消选择
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ======================================================================================================================
; 菜单视图
; ======================================================================================================================
; 自动换行
WordWrapFN(Item, *) {
   Static WordWrap := False
   WordWrap ^= True ; 切换自动换行状态
   RE2.WordWrap(WordWrap) ; 设置自动换行
   ViewMenu.ToggleCheck(Item) ; 切换菜单项的选中状态
   If (WordWrap) ; 如果启用了自动换行
      ViewMenu.Disable(MenuWysiwyg) ; 禁用所见即所得菜单项
   Else ; 否则
      ViewMenu.Enable(MenuWysiwyg) ; 启用所见即所得菜单项
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; 缩放
Zoom100FN(*) => ZoomFN(100, "100 %") ; 100% 缩放
ZoomFN(Ratio, Item, *) {
   Static Zoom := "100 %" ; 静态变量，当前缩放比例
   ZoomMenu.UnCheck(Zoom) ; 取消选中当前缩放比例
   Zoom := Item ; 设置新的缩放比例
   ZoomMenu.Check(Zoom) ; 选中新的缩放比例
   RE2.SetZoom(Ratio) ; 设置缩放比例
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; 所见即所得
WYSIWYGFN(Item, *) {
   Static ShowWysiwyg := False ; 静态变量，显示所见即所得状态
   ShowWysiwyg ^= True ; 切换所见即所得状态
   If (ShowWysiwyg) ; 如果启用了所见即所得
      Zoom100FN() ; 重置缩放比例为 100%
   RE2.WYSIWYG(ShowWysiwyg) ; 设置所见即所得
   ViewMenu.ToggleCheck(Item) ; 切换菜单项的选中状态
   If (ShowWysiwyg) ; 如果启用了所见即所得
      ViewMenu.Disable(MenuWordWrap) ; 禁用自动换行菜单项
   Else ; 否则
      ViewMenu.Enable(MenuWordWrap) ; 启用自动换行菜单项
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; 背景颜色
BackgroundColorFN(Mode, *) {
   Global BackColor ; 全局变量，背景颜色
   Switch Mode { ; 切换模式
      Case "Auto": ; 自动
         RE2.SetBkgndColor("Auto") ; 设置背景颜色为自动
         RE2.BackColor := "Auto" ; 设置背景颜色属性为自动
      Case "Choose": ; 选择
         If RE2.BackColor != "Auto" ; 如果背景颜色不是自动
            Color := RE2.BackColor ; 获取当前背景颜色
         Else ; 否则
            Color := RE2.GetRGB(DllCall("GetSysColor", "Int", 5, "UInt")) ; COLOR_WINDOW
         NC := RichEditDlgs.ChooseColor(RE2, Color)
         If (NC != "") {
            RE2.SetBkgndColor(NC)
            RE2.BackColor := NC
         }
   }
   RE2.Focus()
}
; ----------------------------------------------------------------------------------------------------------------------
; 自动URL检测
AutoURLDetectionFN(ItemName, ItemPos, MenuObj) {
   Static AutoURL := False
   RE2.AutoURL(AutoURL ^= True) ; 切换自动URL检测状态
   MenuObj.ToggleCheck(ItemName) ; 切换菜单项的选中状态
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ======================================================================================================================
; 菜单字符
; ======================================================================================================================
; 选择字体
ChooseFontFN(*) {
   Global FontName, FontSize ; 全局变量，字体名称和大小
   RichEditDlgs.ChooseFont(RE2) ; 调用选择字体对话框
   Font := RE2.GetFont() ; 获取字体
   FontName := Font.Name ; 设置字体名称
   FontSize := Font.Size ; 设置字体大小
   MainFNAME.Text := FontName ; 更新字体名称显示
   MainFSIZE.Text := Round(FontSize) ; 更新字体大小显示
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; MTextColor    ; 菜单项标签
; BTextColor    ; 按钮标签
TextColorFN(Mode, *) {
   Global TextColor ; 全局变量，文本颜色
   Switch Mode { ; 切换模式
      Case "Auto": ; 自动
         RE2.SetFont({Color: "Auto"}) ; 设置字体颜色为自动
         RE2.TextColor := "Auto" ; 设置文本颜色属性为自动
      Case "Choose": ; 选择
         If RE2.TextColor != "Auto" ; 如果文本颜色不是自动
            Color := RE2.TextColor ; 获取当前文本颜色
         Else ; 否则
            Color := RE2.GetRGB(DllCall("GetSysColor", "Int", 8, "UInt")) ; COLOR_WINDOWTEXT
         NC := RichEditDlgs.ChooseColor(RE2, Color) ; 调用选择颜色对话框
         If (NC != "") { ; 如果选择了颜色
            RE2.SetFont({Color: NC}) ; 设置字体颜色
            RE2.TextColor := NC ; 设置文本颜色属性
         }
   }
   UpdateGui() ; 更新 GUI
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; MTextBkColor  ; 菜单项标签
; BTextBkColor  ; 按钮标签
TextBkColorFN(Mode, *) {
   Global TextBkColor ; 全局变量，文本背景颜色
   Switch Mode { ; 切换模式
      Case "Auto": ; 自动
         RE2.SetFont({BkColor: "Auto"}) ; 设置字体背景颜色为自动
      Case "Choose": ; 选择
         If RE2.TxBkColor != "Auto" ; 如果文本背景颜色不是自动
            Color := RE2.TxBkColor ; 获取当前文本背景颜色
         Else ; 否则
            Color := RE2.GetRGB(DllCall("GetSysColor", "Int", 5, "UInt")) ; COLOR_WINDOW
         NC := RichEditDlgs.ChooseColor(RE2, Color) ; 调用选择颜色对话框
         If (NC != "") { ; 如果选择了颜色
            RE2.SetFont({BkColor: NC}) ; 设置字体背景颜色
            RE2.TxBkColor := NC ; 设置文本背景颜色属性
         }
   }
   UpdateGui() ; 更新 GUI
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ======================================================================================================================
; 菜单段落
; ======================================================================================================================
; 左对齐
; 居中对齐
; 右对齐:
; 两端对齐
AlignFN(Alignment, *) {
   Static Align := {Left: 1, Right: 2, Center: 3, Justify: 4} ; 静态变量，对齐方式
   If Align.HasProp(Alignment) ; 如果存在该对齐方式
      RE2.AlignText(Align.%Alignment%) ; 设置文本对齐方式
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
IndentationFN(Mode, *) {
   Switch Mode { ; 切换模式
      Case "Set": ParaIndentGui(RE2) ; 显示段落缩进 GUI
      Case "Reset": RE2.SetParaIndent() ; 重置段落缩进
   }
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; 编号
NumberingFN(Mode, *) {
   Switch Mode { ; 切换模式
      Case "Set": ParaNumberingGui(RE2) ; 显示段落编号 GUI
      Case "Reset": RE2.SetParaNumbering() ; 重置段落编号
   }
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; 段落间距
; 重置段落间距
ParaSpacingFN(Mode, *) {
   Switch Mode { ; 切换模式
      Case "Set": ParaSpacingGui(RE2) ; 显示段落间距 GUI
      Case "Reset": RE2.SetParaSpacing() ; 重置段落间距
   }
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; 单倍行距
; 1.5倍行距
; 双倍行距
SpacingFN(Val, *) {
   RE2.SetLineSpacing(Val) ; 设置行间距
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ----------------------------------------------------------------------------------------------------------------------
; 设置制表位
; 重置制表位
; 设置默认制表位
SetTabStopsFN(Mode, *) {
   Switch Mode { ; 切换模式
      Case "Set": SetTabStopsGui(RE2) ; 显示设置制表位 GUI
      Case "Reset": RE2.SetTabStops() ; 重置制表位
      Case "Default": RE2.SetDefaultTabs(1) ; 设置默认制表位
   }
   RE2.Focus() ; 设置焦点到 RichEdit 控件
}
; ======================================================================================================================
; 菜单搜索
; ======================================================================================================================
FindFN(*) {
   RichEditDlgs.FindText(RE2) ; 调用查找文本对话框
}
; ----------------------------------------------------------------------------------------------------------------------
ReplaceFN(*) {
   RichEditDlgs.ReplaceText(RE2) ; 调用替换文本对话框
}
; ======================================================================================================================
; 段落缩进 GUI
; 此函数创建一个用于设置段落缩进的对话框
; ======================================================================================================================
paraIndentGui(RE) {
   Static Owner := "", ; 静态变量，所有者窗口
          Success := False
   Metrics := RE.GetMeasurement()
   PF2 := RE.GetParaFormat()
   Owner := RE.Gui.Hwnd
   ParaIndentGui := Gui("+Owner" . Owner . " +ToolWindow +LastFound", "段落缩进")
   ParaIndentGui.OnEvent("Close", ParaIndentGuiClose)
   ParaIndentGui.MarginX := 20
   ParaIndentGui.MarginY := 10
   ParaIndentGui.AddText("Section h20 0x200", "首行左缩进 (绝对):")
   ParaIndentGui.AddText("xs hp 0x200", "其他行左缩进 (相对):")
   ParaIndentGui.AddText("xs hp 0x200", "所有行右缩进 (绝对):")
   EDLeft1 := ParaIndentGui.AddEdit("ys hp Limit5")
   EDLeft2 := ParaIndentGui.AddEdit("hp Limit6")
   EDRight := ParaIndentGui.AddEdit("hp Limit5")
   CBStart := ParaIndentGui.AddCheckBox("ys x+5 hp", "应用")
   CBOffset := ParaIndentGui.AddCheckBox("hp", "应用")
   CBRight := ParaIndentGui.AddCheckBox("hp", "应用")
   Left1 := Round((PF2.StartIndent / 1440) * Metrics, 2)
   If (Metrics = 2.54)
      Left1 := RegExReplace(Left1, "\.", ",")
   EDLeft1.Text := Left1
   Left2 := Round((PF2.Offset / 1440) * Metrics, 2)
   If (Metrics = 2.54)
      Left2 := RegExReplace(Left2, "\.", ",")
   EDLeft2.Text := Left2
   Right := Round((PF2.RightIndent / 1440) * Metrics, 2)
   If (Metrics = 2.54)
      Right := RegExReplace(Right, "\.", ",")
   EDRight.Text := Right
   BN1 := ParaIndentGui.AddButton("xs", "Apply")
   BN1.OnEvent("Click", ParaIndentGuiApply)
   BN2 := ParaIndentGui.AddButton("x+10 yp", "Cancel")
   BN2.OnEvent("Click", ParaIndentGuiClose)
   BN2.GetPos( , , &BW := 0)  ; 获取按钮宽度
   BN1.Move( , , BW)  ; 调整按钮1宽度
   CBRight.GetPos(&CX := 0, , &CW := 0)
   BN2.Move(CX + CW - BW)
   RE.Gui.Opt("+Disabled")
   ParaIndentGui.Show()
   WinWaitActive  ; 等待窗口激活
   WinWaitClose  ; 等待窗口关闭
   Return Success  ; 返回成功标志
   ; -------------------------------------------------------------------------------------------------------------------
   ParaIndentGuiClose(*) {
      Success := False
      RE.Gui.Opt("-Disabled")
      ParaIndentGui.Destroy()
   }
   ; -------------------------------------------------------------------------------------------------------------------
   ParaIndentGuiApply(*) {
      ApplyStart := CBStart.Value
      ApplyOffset := CBOffset.Value
      ApplyRight := CBRight.Value
      Indent := {}
      If ApplyStart {
         Start := EDLeft1.Text
         If (Start = "")
            Start := 0
         If !RegExMatch(Start, "^\d{1,2}((\.|,)\d{1,2})?$") {
            EDLeft1.Text := ""
            EDLeft1.Focus()
            Return
         }
         Indent.Start := StrReplace(Start, ",", ".")
      }
      If (ApplyOffset) {
         Offset := EDLeft2.Text
         If (Offset = "")
            Offset := 0
         If !RegExMatch(Offset, "^(-)?\d{1,2}((\.|,)\d{1,2})?$") {
            EDLeft2.Text := ""
            EDLeft2.Focus()
            Return
         }
         Indent.Offset := StrReplace(Offset, ",", ".")
      }
      If (ApplyRight) {
         Right := EDRight.Text
         If (Right = "")
            Right := 0
         If !RegExMatch(Right, "^\d{1,2}((\.|,)\d{1,2})?$") {
            EDRight.Text := ""
            EDRight.Focus()
            Return
         }
         Indent.Right := StrReplace(Right, ",", ".")
      }
      Success := RE.SetParaIndent(Indent)
      RE.Gui.Opt("-Disabled")
      ParaIndentGui.Destroy()
   }
}
; ======================================================================================================================
; 段落编号 GUI
; ======================================================================================================================
ParaNumberingGui(RE) {
   Static Owner := "",  ; 所有者窗口
          Bullet := "•",  ; 项目符号
          StyleArr := ["1)", "(1)", "1.", "1", "w/o"],  ; 样式数组
          TypeArr := [Bullet, "0, 1, 2", "a, b, c", "A, B, C", "i, ii, iii", "I, I, III"],  ; 类型数组
          PFN := ["Bullet", "Arabic", "LCLetter", "UCLetter", "LCRoman", "UCRoman"],  ; 编号格式名称
          PFNS := ["Paren", "Parens", "Period", "Plain", "None"],  ; 编号格式样式
          Success := False  ; 成功标志
   Metrics := RE.GetMeasurement()  ; 获取度量单位
   PF2 := RE.GetParaFormat()  ; 获取段落格式
   Owner := RE.Gui.Hwnd  ; 获取所有者窗口句柄
   ParaNumberingGui := Gui("+Owner" . Owner . " +ToolWindow +LastFound", "段落编号")  ; 创建段落编号 GUI
   ParaNumberingGui.OnEvent("Close", ParaNumberingGuiClose)  ; 绑定关闭事件
   ParaNumberingGui.MarginX := 20  ; 设置水平边距
   ParaNumberingGui.MarginY := 10  ; 设置垂直边距
   ParaNumberingGui.AddText("Section h20 w100 0x200", "类型:")  ; 添加类型标签
   DDLType := ParaNumberingGui.AddDDL("xp y+0 wp AltSubmit", TypeArr)  ; 添加类型下拉列表
   If (PF2.Numbering)  ; 如果已有编号
      DDLType.Choose(PF2.Numbering)  ; 选择对应的编号类型
   ParaNumberingGui.AddText("xs h20 w100 0x200", "起始于:")  ; 添加起始标签
   EDStart := ParaNumberingGui.AddEdit("y+0 wp hp Limit5", PF2.NumberingStart)  ; 添加起始编辑框
   ParaNumberingGui.AddText("ys h20 w100 0x200", "样式:")  ; 添加样式标签
   DDLStyle := ParaNumberingGui.AddDDL("y+0 wp AltSubmit Choose1", StyleArr)  ; 添加样式下拉列表
   If (PF2.NumberingStyle)  ; 如果已有样式
      DDLStyle.Choose((PF2.NumberingStyle // 0x0100) + 1)  ; 选择对应的样式
   ParaNumberingGui.AddText("h20 w100 0x200", "距离:  (" . (Metrics = 1.00 ? "英寸" : "厘米") . ")")  ; 添加距离标签
   EDDist := ParaNumberingGui.AddEdit("y+0 wp hp Limit5")  ; 添加距离编辑框
   Tab := Round((PF2.NumberingTab / 1440) * Metrics, 2)  ; 计算标签位置
   If (Metrics = 2.54)  ; 如果是厘米
      Tab := RegExReplace(Tab, ".", ",")  ; 替换小数点为逗号
   EDDist.Text := Tab  ; 设置编辑框文本
   BN1 := ParaNumberingGui.AddButton("xs", "应用") ; gParaNumberingGuiApply hwndhBtn1, 应用
   BN1.OnEvent("Click", ParaNumberingGuiApply)  ; 绑定应用按钮点击事件
   BN2 := ParaNumberingGui.AddButton("x+10 yp", "取消") ;  gParaNumberingGuiClose hwndhBtn2, 取消
   BN2.OnEvent("Click", ParaNumberingGuiClose)  ; 绑定取消按钮点击事件
   BN2.GetPos( , , &BW := 0)
   BN1.Move( , , BW)
   DDLStyle.GetPos(&DX := 0, , &DW := 0)  ; 获取下拉列表位置和宽度
   BN2.Move(DX + DW - BW)  ; 移动按钮2位置
   RE.Gui.Opt("+Disabled")  ; 禁用主窗口
   ParaNumberingGui.Show()  ; 显示段落编号 GUI
   WinWaitActive  ; 等待窗口激活
   WinWaitClose  ; 等待窗口关闭
   Return Success  ; 返回成功标志
   ; -------------------------------------------------------------------------------------------------------------------
   ; 段落编号 GUI 关闭函数
   ParaNumberingGuiClose(*){
      Success := False  ; 设置成功标志为假
      RE.Gui.Opt("-Disabled")  ; 启用主窗口
      ParaNumberingGui.Destroy()  ; 销毁 GUI
   }
   ; -------------------------------------------------------------------------------------------------------------------
   ; 段落编号 GUI 应用函数
   ParaNumberingGuiApply(*){
      Type := DDLType.Value  ; 获取类型值
      Style := DDLStyle.Value  ; 获取样式值
      Start := EDStart.Text  ; 获取起始值
      Tab := EDDist.Text  ; 获取距离值
      If !RegExMatch(Tab, "^\d{1,2}((\.|,)\d{1,2})?$") {  ; 验证距离格式
         EDDist.Text := ""  ; 清空编辑框
         EDDist.Focus()  ; 设置焦点
         Return  ; 返回
      }
      Numbering := {Type: PFN[Type], Style: PFNS[Style]}  ; 创建编号对象
      Numbering.Tab := RegExReplace(Tab, ",", ".")  ; 替换逗号为小数点
      Numbering.Start := Start  ; 设置起始值
      Success := RE.SetParaNumbering(Numbering)  ; 设置段落编号
      RE.Gui.Opt("-Disabled")  ; 启用主窗口
      ParaNumberingGui.Destroy()  ; 销毁 GUI
   }
}
; ======================================================================================================================
; 段落间距 GUI
; ======================================================================================================================
; 段落间距 GUI 函数
ParaSpacingGui(RE) {
   Static Owner := "",  ; 所有者窗口
          Success := False  ; 成功标志
   PF2 := RE.GetParaFormat()  ; 获取段落格式
   Owner := RE.Gui.Hwnd  ; 获取所有者窗口句柄
   ParaSpacingGui := Gui("+Owner" . Owner . " +ToolWindow +LastFound", "段落间距") ; +LabelParaSpacingGui  ; 创建段落间距 GUI
   ParaSpacingGui.OnEvent("Close", ParaSpacingGuiClose)  ; 绑定关闭事件
   ParaSpacingGui.MarginX := 20  ; 设置水平边距
   ParaSpacingGui.MarginY := 10  ; 设置垂直边距
   ParaSpacingGui.AddText("Section h20 0x200", "段前间距 (点):")  ; 添加段前间距标签
   ParaSpacingGui.AddText("xs y+10 hp 0x200", "段后间距 (点):")  ; 添加段后间距标签
   EDBefore := ParaSpacingGui.AddEdit("ys hp Number Limit2 Right", "00")  ; 添加段前间距编辑框
   EDBefore.Text := PF2.SpaceBefore // 20  ; 设置段前间距值
   EDAfter := ParaSpacingGui.AddEdit("xp y+10 hp Number Limit2 Right", "00")  ; 添加段后间距编辑框
   EDAfter.Text := PF2.SpaceAfter // 20  ; 设置段后间距值
   BN1 := ParaSpacingGui.AddButton("xs", "应用")  ; 添加应用按钮
   BN1.OnEvent("Click", ParaSpacingGuiApply)  ; 绑定应用按钮点击事件
   BN2 := ParaSpacingGui.AddButton("x+10 yp", "取消")  ; 添加取消按钮
   BN2.OnEvent("Click", ParaSpacingGuiClose)  ; 绑定取消按钮点击事件
   BN2.GetPos( , ,&BW := 0)  ; 获取按钮宽度
   BN1.Move( , ,BW)  ; 调整按钮1宽度
   EDAfter.GetPos(&EX := 0, , &EW := 0)  ; 获取编辑框位置和宽度
   X := EX + EW - BW  ; 计算按钮2位置
   BN2.Move(X)  ; 移动按钮2位置
   RE.Gui.Opt("+Disabled")  ; 禁用主窗口
   ParaSpacingGui.Show()  ; 显示段落间距 GUI
   WinWaitActive
   WinWaitClose
   Return Success
   ; -------------------------------------------------------------------------------------------------------------------
   ; 段落间距 GUI 关闭函数
   ParaSpacingGuiClose(*) {
      Success := False  ; 设置成功标志为假
      RE.Gui.Opt("-Disabled")  ; 启用主窗口
      ParaSpacingGui.Destroy()  ; 销毁 GUI
   }
   ; -------------------------------------------------------------------------------------------------------------------
   ; 段落间距 GUI 应用函数
   ParaSpacingGuiApply(*) {
      Before := EDBefore.Text  ; 获取段前间距值
      After := EDAfter.Text  ; 获取段后间距值
      Success := RE.SetParaSpacing({Before: Before, After: After})  ; 设置段落间距
      RE.Gui.Opt("-Disabled")  ; 启用主窗口
      ParaSpacingGui.Destroy()  ; 销毁 GUI
   }
}
; ======================================================================================================================
; 设置制表位 GUI
; ======================================================================================================================
; 设置制表位 GUI 函数
SetTabStopsGui(RE) {
   ; 设置段落的制表位
   ; 使用参数 mode = "Reset" 重置为默认制表位
   ; EM_GETPARAFORMAT = 0x43D, EM_SETPARAFORMAT = 0x447
   ; PFM_TABSTOPS = 0x10
   Static Owner   := "",  ; 所有者窗口
          Metrics := 0,  ; 度量单位
          MinTab  := 0.30,     ; 最小制表位 (英寸)
          MaxTab  := 8.30,     ; 最大制表位 (英寸)
          AL := 0x00000000,    ; 左对齐 (默认)
          AC := 0x01000000,    ; 居中对齐
          AR := 0x02000000,    ; 右对齐
          AD := 0x03000000,    ; 小数点对齐
          Align := {0x00000000: "L", 0x01000000: "C", 0x02000000: "R", 0x03000000: "D"},
          TabCount := 0,       ; 制表位计数
          MAX_TAB_STOPS := 32,  ; 最大制表位数
          Success := False     ; 返回值
   Metrics := RE.GetMeasurement()  ; 获取度量单位
   PF2 := RE.GetParaFormat()  ; 获取段落格式
   TabCount := PF2.TabCount  ; 获取制表位计数
   Tabs := []  ; 创建制表位数组
   Tabs.Length := PF2.Tabs.Length  ; 设置数组长度
   For I, V In PF2.Tabs  ; 遍历制表位
      Tabs[I] := [Format("{:.2f}", Round(((V & 0x00FFFFFF) * Metrics) / 1440, 2)), V & 0xFF000000]  ; 格式化制表位
   Owner := RE.Gui.Hwnd  ; 获取所有者窗口句柄
   SetTabStopsGui := Gui("+Owner" . Owner . " +ToolWindow +LastFound", "设置制表位")  ; 创建设置制表位 GUI
   SetTabStopsGui.OnEvent("Close", SetTabStopsGuiClose)  ; 绑定关闭事件
   SetTabStopsGui.MarginX := 10  ; 设置水平边距
   SetTabStopsGui.MarginY := 10  ; 设置垂直边距
   SetTabStopsGui.AddText("Section", "位置: (" . (Metrics = 1.00 ? "英寸" : "厘米") . ")")  ; 添加位置标签
   CBBTabs := SetTabStopsGui.AddComboBox("xs y+2 w120 r6 Simple +0x800 AltSubmit")  ; 添加制表位组合框
   CBBTabs.OnEvent("Change", SetTabStopsGuiSelChanged)  ; 绑定组合框变更事件
   If (TabCount) {  ; 如果有制表位
      For T In Tabs {
         I := SendMessage(0x0143, 0, StrPtr(T[1]), CBBTabs.Hwnd)  ; CB_ADDSTRING  添加字符串到组合框
         SendMessage(0x0151, I, T[2], CBBTabs.Hwnd)               ; CB_SETITEMDATA  设置组合框项数据
      }
   }
   SetTabStopsGui.AddText("ys Section", "对齐方式:")  ; 添加对齐方式标签
   RBL := SetTabStopsGui.AddRadio("xs w60 Section y+2 Checked Group", "左对齐")  ; 添加左对齐单选按钮
   RBC := SetTabStopsGui.AddRadio("wp", "居中对齐")  ; 添加居中对齐单选按钮
   RBR := SetTabStopsGui.AddRadio("ys wp", "右对齐")  ; 添加右对齐单选按钮
   RBD := SetTabStopsGui.AddRadio("wp", "小数点对齐")  ; 添加小数点对齐单选按钮
   BNAdd := SetTabStopsGui.AddButton("xs Section w60 Disabled", "&添加")  ; 添加添加按钮
   BNAdd.OnEvent("Click", SetTabStopsGuiAdd)  ; 绑定添加按钮点击事件
   BNRem := SetTabStopsGui.AddButton("ys w60 Disabled", "&移除")  ; 添加移除按钮
   BNRem.OnEvent("Click", SetTabStopsGuiRemove)  ; 绑定移除按钮点击事件
   BNAdd.GetPos(&X1 := 0)  ; 获取添加按钮位置
   BNRem.GetPos(&X2 := 0, , &W2 := 0)  ; 获取移除按钮位置和宽度
   W := X2 + W2 - X1  ; 计算清除按钮宽度
   BNClr := SetTabStopsGui.AddButton("xs w" . W, "&清除全部")  ; 添加清除全部按钮
   BNClr.OnEvent("Click", SetTabStopsGuiRemoveAll)  ; 绑定清除全部按钮点击事件
   SetTabStopsGui.AddText("xm h5")  ; 添加空白文本
   BNApply := SetTabStopsGui.AddButton("xm y+0 w60", "&应用")  ; 添加应用按钮
   BNApply.OnEvent("Click", SetTabStopsGuiApply)  ; 绑定应用按钮点击事件
   X := X2 + W2 - 60  ; 计算取消按钮位置
   BNCancel := SetTabStopsGui.AddButton("x" . X . " yp wp", "&取消")  ; 添加取消按钮
   BNCancel.OnEvent("Click", SetTabStopsGuiClose)  ; 绑定取消按钮点击事件
   RE.Gui.Opt("+Disabled")  ; 禁用主窗口
   SetTabStopsGui.Show()  ; 显示设置制表位 GUI
   WinWaitActive  ; 等待窗口激活
   WinWaitClose  ; 等待窗口关闭
   Return Success  ; 返回成功标志
   ; -------------------------------------------------------------------------------------------------------------------
   ; 设置制表位 GUI 关闭函数
   SetTabStopsGuiClose(*) {
      Success := False  ; 设置成功标志为假
      RE.Gui.Opt("-Disabled")  ; 启用主窗口
      SetTabStopsGui.Destroy()  ; 销毁 GUI
   }
   ; -------------------------------------------------------------------------------------------------------------------
   ; 设置制表位 GUI 选择变更函数
   SetTabStopsGuiSelChanged(*) {
      If (TabCount < MAX_TAB_STOPS)  ; 如果制表位数小于最大制表位数
         BNAdd.Enabled := !!RegExMatch(CBBTabs.Text, "^\d*[.,]?\d+$")  ; 启用添加按钮
      If !(I := CBBTabs.Value) {  ; 如果没有选择项
         BNRem.Enabled := False  ; 禁用移除按钮
         Return  ; 返回
      }
      BNRem.Enabled := True  ; 启用移除按钮
      A := SendMessage(0x0150, I - 1, 0, CBBTabs.Hwnd) ; CB_GETITEMDATA 获取项数据
      C := A = AC ? RBC : A = AR ? RBR : A = AD ? RBD : RBl  ; 确定对齐方式
      C.Value := 1  ; 设置选中状态
   }
   ; -------------------------------------------------------------------------------------------------------------------
   ; 设置制表位 GUI 添加函数
   SetTabStopsGuiAdd(*) {
      T := CBBTabs.Text  ; 获取文本
      If !RegExMatch(T, "^\d*[.,]?\d+$") {  ; 验证格式
         CBBTabs.Focus()  ; 设置焦点
         Return  ; 返回
      }
      T := Round(StrReplace(T, ",", "."), 2)  ; 替换逗号为小数点并四舍五入
      RT := Round(T / Metrics, 2)  ; 计算实际制表位
      If (RT < MinTab) || (RT > MaxTab){  ; 检查是否在范围内
         CBBTabs.Focus()  ; 设置焦点
         Return  ; 返回
      }
      A := RBC.Value ? AC : RBR.Value ? AR : RBD.Value ? AD : AL  ; 确定对齐方式
      TabArr := ControlGetItems(CBBTabs.Hwnd)  ; 获取组合框项
      P := -1  ; 插入位置
      T := Format("{:.2f}", T)  ; 格式化制表位
      For I, V In TabArr {
         If (T < V) {  ; 如果小于当前项
            P := I - 1  ; 设置插入位置
            Break  ; 跳出循环
         }
         IF (T = V) {  ; 如果等于当前项
            P := I - 1  ; 设置插入位置
            CBBTabs.Delete(I)  ; 删除当前项
            Break  ; 跳出循环
         }
      }
      I := SendMessage(0x014A, P, StrPtr(T), CBBTabs.Hwnd)  ; CB_INSERTSTRING 插入字符串
      SendMessage(0x0151, I, A, CBBTabs.Hwnd)               ; CB_SETITEMDATA 设置项数据
      TabCount++  ; 制表位数加1
      If !(TabCount < MAX_TAB_STOPS)  ; 如果达到最大制表位数
         BNAdd.Enabled := False  ; 禁用添加按钮
      CBBTabs.Text := ""  ; 清空文本
      CBBTabs.Focus()  ; 设置焦点
   }
   ; -------------------------------------------------------------------------------------------------------------------
   ; 设置制表位 GUI 移除函数
   SetTabStopsGuiRemove(*) {
      If (I := CBBTabs.Value) {  ; 如果有选择项
         CBBTabs.Delete(I)  ; 删除选择项
         CBBTabs.Text := ""
         TabCount--
         RBL.Value := 1
      }
      CBBTabs.Focus()
   }
   ; -------------------------------------------------------------------------------------------------------------------
   ; 设置制表位 GUI 清除全部函数
   SetTabStopsGuiRemoveAll(*) {
      CBBTabs.Text := ""  ; 清空文本
      CBBTabs.Delete()  ; 删除所有项
      RBL.Value := 1  ; 设置左对齐为选中状态
      CBBTabs.Focus()  ; 设置焦点
   }
   ; -------------------------------------------------------------------------------------------------------------------
   ; 设置制表位 GUI 应用函数
   SetTabStopsGuiApply(*) {
      TabCount := SendMessage(0x0146, 0, 0, CBBTabs.Hwnd) << 32 >> 32 ; CB_GETCOUNT 获取项数
      If (TabCount < 1)  ; 如果没有项
         Return  ; 返回
      TabArr := ControlGetItems(CBBTabs.HWND)  ; 获取所有项
      TabStops := {}  ; 创建制表位对象
      For I, T In TabArr {
         Alignment := Format("0x{:08X}", SendMessage(0x0150, I - 1, 0, CBBTabs.HWND)) ; CB_GETITEMDATA 获取对齐方式
         TabPos := Format("{:i}", T * 100)  ; 计算制表位位置
         TabStops.%TabPos% := Align.%Alignment%  ; 设置制表位对齐方式
      }
      Success := RE.SetTabStops(TabStops)  ; 设置制表位
      RE.Gui.Opt("-Disabled")  ; 启用主窗口
      SetTabStopsGui.Destroy()  ; 销毁 GUI
   }
}
; ======================================================================================================================
; 为任何 Gui 控件设置多行工具提示。
; 参数:
;     GuiCtrl     -  一个 Gui.Control 对象
;     TipText     -  工具提示的文本。如果为以前添加的控件传递空字符串，
;                    则会删除其工具提示。
;     UseAhkStyle -  如果设置为 true，工具提示将使用 AHK 工具提示的视觉样式显示。
;                    否则，将使用当前主题设置。
;                    默认值: True
;     CenterTip   -  如果设置为 true，工具提示将在控件下方/上方居中显示。
;                    默认值: False
; 返回值:
;     成功时返回 True，否则返回 False。
; 备注:
;     文本和图片控件需要 SS_NOTIFY (+0x0100) 样式。
; MSDN:
;     https://learn.microsoft.com/en-us/windows/win32/controls/tooltip-control-reference
; ======================================================================================================================
GuiCtrlSetTip(GuiCtrl, TipText, UseAhkStyle := True, CenterTip := False) {
   Static SizeOfTI := 24 + (A_PtrSize * 6)
   Static Tooltips := Map()
   Local Flags, HGUI, HCTL, HTT, TI
   ; 检查传入的GuiCtrl
   If !(GuiCtrl Is Gui.Control)
      Return False
   HGUI := GuiCtrl.Gui.Hwnd
   ; 创建TOOLINFO结构 -> msdn.microsoft.com/en-us/library/bb760256(v=vs.85).aspx
   Flags := 0x11 | (CenterTip ? 0x02 : 0x00) ; TTF_SUBCLASS | TTF_IDISHWND [| TTF_CENTERTIP]
   TI := Buffer(SizeOfTI, 0)
   NumPut("UInt", SizeOfTI, "UInt", Flags, "UPtr", HGUI, "UPtr", HGUI, TI) ; cbSize, uFlags, hwnd, uID
   ; 为这个Gui创建工具提示控件（如果需要）
   If !ToolTips.Has(HGUI) {
      If !(HTT := DllCall("CreateWindowEx", "UInt", 0, "Str", "tooltips_class32", "Ptr", 0, "UInt", 0x80000003
                                          , "Int", 0x80000000, "Int", 0x80000000, "Int", 0x80000000, "Int", 0x80000000
                                          , "Ptr", HGUI, "Ptr", 0, "Ptr", 0, "Ptr", 0, "UPtr"))
         Return False
      If (UseAhkStyle)
         DllCall("Uxtheme.dll\SetWindowTheme", "Ptr", HTT, "Ptr", 0, "Ptr", 0)
      SendMessage(0x0432, 0, TI.Ptr, HTT) ; TTM_ADDTOOLW
      Tooltips[HGUI] := {HTT: HTT, Ctrls: Map()}
   }
   HTT := Tooltips[HGUI].HTT  ; 获取工具提示控件句柄
   HCTL := GuiCtrl.HWND  ; 获取控件句柄
   ; 添加/删除控件的工具提示
   NumPut("UPtr", HCTL, TI, 8 + A_PtrSize) ; uID 设置控件ID
   NumPut("UPtr", HCTL, TI, 24 + (A_PtrSize * 4)) ; uID 设置控件ID
   If !Tooltips[HGUI].Ctrls.Has(HCTL) { ; 添加控件
      If (TipText = "")  ; 如果提示文本为空
         Return False  ; 返回False
      SendMessage(0x0432, 0, TI.Ptr, HTT) ; TTM_ADDTOOLW 添加工具提示
      SendMessage(0x0418, 0, -1, HTT) ; TTM_SETMAXTIPWIDTH 设置最大提示宽度
      Tooltips[HGUI].Ctrls[HCTL] := True  ; 标记控件已添加
   }
   Else If (TipText = "") { ; 移除控件
      SendMessage(0x0433, 0, TI.Ptr, HTT) ; TTM_DELTOOLW 删除工具提示
      Tooltips[HGUI].Ctrls.Delete(HCTL)  ; 从映射中删除控件
      Return True  ; 返回True
   }
   ; 设置/更新工具提示文本。
   NumPut("UPtr", StrPtr(TipText), TI, 24 + (A_PtrSize * 3))  ; 提示文本
	SendMessage(0x0439, 0, TI.Ptr, HTT) ; TTM_UPDATETIPTEXTW
	Return True
}

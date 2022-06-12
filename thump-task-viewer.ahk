/*	介绍：在桌面右侧显示当前可见窗口的动态缩略图，鼠标点击激活窗口
	作者：nepter
	环境：Autohotkey_L，win vista sp1以上，32位，已开启windows aero效果
	热键：
		capslock	显示/隐藏窗口
		ctrl + /	刷新窗口
		win + esc	退出
*/

; 可设置的参数
global isShowBtn := 0  ; 是否显示按钮
global d_margin := 5  ; 边距
global d_height := 100  ; 缩略图高度
Hotkey,CapsLock,dtl_show  ; 显示/隐藏窗口
Hotkey,^/,dtl_refresh  ; 刷新窗口
Hotkey,#Esc,dtl_quit  ; 退出

#SingleInstance force
#NoEnv
#NoTrayIcon
SetBatchLines,-1

; 初始化参数
SysGet,mon_,MonitorWorkArea
global ScreenHeight :=  mon_bottom - mon_top  ; 工作区高度
global ScreenWidth := mon_right - mon_left  ; 工作区宽度
global hgui , dtl , vis := 1 , swa := 0
global top_height := isShowBtn?60:0 ; 顶部高度
global max_row := (ScreenHeight-top_height)//(d_height+2*d_margin)  ; 每列数量
global d_width := d_height*ScreenWidth//ScreenHeight  ; 缩略图宽度
global oShell := ComObjCreate("shell.application")

; 判断运行环境
if (DllCall("GetVersion") & 0xFF < 6)
{
	msgbox 只能在win vista sp1 及以上版本下运行，自动退出
	ExitApp
}
VarSetCapacity(pf,4), DllCall("Dwmapi\DwmIsCompositionEnabled","int*",pf)
if !NumGet(pf)
{
	msgbox 请开启DWM，自动退出
	ExitApp
}

; 显示缩略图的窗口
gui,1:new,+hwndhgui -Caption +AlwaysOnTop +ToolWindow,DWM_Thumbnail_View
gui,1:color,black
if isShowBtn
{
	; 设定工作区大小按钮，如果不想窗口和背景窗口重叠，可以点击设定
	gui,1:add,Button,x5 y5 w%d_width% gbtn_workarea vbtn_wa,Set Work Area
	; 刷新按钮
	gui,1:add,button,x5 y30 w%d_width% gbtn_refresh,refresh
}
OnMessage(0x201,"event_mouse_click")  ; WM_LBUTTONDOWN

; 开启玻璃化
VarSetCapacity(pBlurBehind,16,0)
NumPut(1,pBlurBehind) , NumPut(1,pBlurBehind,4)
hr := Dllcall("Dwmapi\DwmEnableBlurBehindWindow","uint",hgui,"uint",&pBlurBehind)
if hr
{
	msgbox 窗口玻璃化失败，自动退出
	ExitApp
}

dwm_view:
dtl := []  ; 缩略图collection

; 枚举可视化的窗口
pEW := RegisterCallback("EnumWindows")
DllCall("EnumWindows","ptr",pEW,"int",0)

; 创建缩略图
loop % dtl.maxindex()
{
	VarSetCapacity(phThumbnailId,A_PtrSize,0)
	hr := Dllcall("Dwmapi\DwmRegisterThumbnail","uint",hgui,"uint",dtl[A_Index]["id"],"ptr",&phThumbnailId)
	if !hr
	{
        hThumbnailId := NumGet(phThumbnailId)
		dtl[A_Index].Insert("tid",hThumbnailId)
		
		; 取得窗口原始大小，w_
		VarSetCapacity(lpwndpl,44),NumPut(44,lpwndpl)
		,DllCall("GetWindowPlacement","uint",dtl[A_Index].id,"ptr",&lpwndpl)
		,flags := NumGet(lpwndpl,4,"int") , showcmd := NumGet(lpwndpl,8,"int")
		,w_left := NumGet(lpwndpl,28,"int") , w_top := NumGet(lpwndpl,32,"int") , w_right := NumGet(lpwndpl,36,"int") , w_bottom := NumGet(lpwndpl,40,"int")
		
		; source窗口大小，R_
		R_left := 0 , R_top := 0
		if (flags=2)  ; 最大化窗口
			R_right := ScreenWidth , R_bottom := ScreenHeight
		else 
			R_right := w_right - w_left , R_bottom := w_bottom - w_top
		opacity := 255 , fVisible := 1 , fSourceClientAreaOnly := 1

		; target窗口大小，d_
		dtl[A_Index].left := d_margin+((A_Index-1)//max_row)*(d_width+d_margin*2)
		, dtl[A_Index].top := top_height+Mod(A_Index-1,max_row)*(d_height+d_margin*2)+d_margin
		, dtl[A_Index].right := dtl[A_Index].left+d_width 
		, dtl[A_Index].bottom := dtl[A_Index].top+d_height
		
		VarSetCapacity(ptnProperties,45,0)
        ,NumPut(3,ptnProperties)
		,NumPut(dtl[A_Index].left,ptnProperties,4,"Int") , NumPut(dtl[A_Index].top,ptnProperties,8,"Int") , NumPut(dtl[A_Index].right,ptnProperties,12,"Int") , NumPut(dtl[A_Index].bottom,ptnProperties,16,"Int")
        ,NumPut(R_left,ptnProperties,20,"Int") , NumPut(R_top,ptnProperties,24,"Int") , NumPut(R_right,ptnProperties,28,"Int") , NumPut(R_bottom,ptnProperties,32,"Int")
        
		hr := Dllcall("Dwmapi\DwmUpdateThumbnailProperties","uint",hThumbnailId,"ptr",&ptnProperties)
		if hr
			msgbox % "error code: " hr "," dtl[A_Index].title
	}
}

; 显示窗口
w := ((dtl.maxindex()-1)//max_row+1)*(d_width+d_margin*2) , h := ScreenHeight , x := ScreenWidth-w , y := 0
IfWinNotExist, ahk_id %hgui%
	gui,1:show,w%w% h%h% x%x% y%y%
else
	WinMove,ahk_id %hgui%,,%x%,%y%,%w%,%h%
return

btn_workarea:
; 设定工作区
GuiControl,,btn_wa,% (swa := !swa)?"Restore Work Area":"Set Work Area"
VarSetCapacity(wa,16,0)
,NumPut(mon_left,wa,0,"int")
,NumPut(mon_top,wa,4,"int")
,NumPut(mon_right-(swa?w:0),wa,8,"int")
,NumPut(mon_bottom,wa,12,"int")
DllCall("SystemParametersInfo","uint",0x2F,"uint",0,"ptr",&wa,"uint",0)

return

; 热键 ctrl + / 刷新缩略图
btn_refresh:
dtl_refresh:
^/::
loop % dtl.maxindex()
	Dllcall("Dwmapi\DwmUnregisterThumbnail","uint",dtl[A_Index]["tid"])
gosub, dwm_view
return

; 热键 win + esc 退出
dtl_quit:
ExitApp

; 热键 capslock 隐藏或显示缩略图窗口
dtl_show:
	if vis
		WinHide,ahk_id %hgui%
	else
	{
		gosub dtl_refresh
		WinShow,ahk_id %hgui%
	}
	vis := !vis
return

; 枚举窗体
EnumWindows(hwnd)
{
	if  DllCall("IsWindowVisible","uint",hwnd) && (hwnd=DllCall("GetAncestor","uint",hwnd,"uint",3))
	{
		WinGetTitle,title,ahk_id %hwnd%
		if title
		{
			WinGetClass,class,ahk_id %hwnd%
			if (class="Button")   ; 跳过开始菜单按钮
				return 1
			else if (class="Progman")  ; 桌面
				title := "桌面"
			else if (hwnd=hgui)  ; 跳过自己
				return 1
			else if (DllCall("GetWindowLongW","uint",hwnd,"int",-20) & 0x80)   ; 跳过toolwindow类型窗口
				return 1
			dtl.Insert({id:hwnd,title:title,class:class})
		}
	}
	return 1
}

; 鼠标点击响应
event_mouse_click(wp,lp)
{
	x := lp & 0xFFFF
    y := lp >> 16
	loop % dtl.maxindex()
	{
		if (x>dtl[A_Index].left && x<dtl[A_Index].right && y>dtl[A_Index].top && y<dtl[A_Index].bottom && id:=dtl[A_Index].id)
		{
			if (dtl[A_Index].class = "Progman")  ; 若是桌面，则显示桌面
				oShell.ToggleDesktop()
			else
				WinActivate, ahk_id %id%
			break
		}
	}
}

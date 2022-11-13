#include <windows.h>
#include <tchar.h>
#include <shellapi.h>
#include <WinUser.h>
#include "resource.h"

#define IDR_PAUSE 12
#define IDR_START 13

LPCTSTR szAppClassName = TEXT("服务程序");
LPCTSTR szAppWindowName = TEXT("服务程序");
HMENU hmenu; //菜单句柄
HINSTANCE GhInstance = NULL;

void DeleteTray(HWND hWnd)
{
	NOTIFYICONDATA nid;
	nid.cbSize = (DWORD)sizeof(NOTIFYICONDATA);
	nid.hWnd = hWnd;
	nid.uID = 0;
	nid.uFlags = NIF_MESSAGE;
	nid.uCallbackMessage = WM_USER;

	Shell_NotifyIcon(NIM_DELETE, &nid); //在托盘中删除图标 
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	NOTIFYICONDATA nid;
	UINT WM_TASKBARCREATED;
	POINT pt; // 用于接收鼠标坐标
	int xx; // 用于接收菜单选项返回值
	// 不要修改TaskbarCreated，这是系统任务栏自定义的消息
	WM_TASKBARCREATED = RegisterWindowMessage(TEXT("TaskbarCreated"));

	switch (message)
	{
	case WM_CREATE: //窗口创建时候的消息.
	{
		// 判断图标 
		int Icon = NULL;
		if (__argc == 2)
		{
			char * ico = __argv[1];
			if (strcmp("0", ico) == 0) Icon = IDI_ICON1;
			else if (strcmp("1", ico) == 0) Icon = IDI_ICON2;
			else if (strcmp("2", ico) == 0) Icon = IDI_ICON3;
			else if (strcmp("3", ico) == 0) Icon = IDI_ICON4;
			else if (strcmp("4", ico) == 0) Icon = IDI_ICON5;
			else if (strcmp("5", ico) == 0) Icon = IDI_ICON6;
			else if (strcmp("6", ico) == 0) Icon = IDI_ICON7;
			else if (strcmp("7", ico) == 0) Icon = IDI_ICON8;
			else if (strcmp("8", ico) == 0) Icon = IDI_ICON9;
			else Icon = IDI_ICON1;
		}
		else Icon = IDI_ICON1;

		nid.cbSize = sizeof(nid);
		nid.hWnd = hwnd;
		nid.uID = 0;
		nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
		nid.uCallbackMessage = WM_USER;
		nid.hIcon = LoadIcon(GhInstance, MAKEINTRESOURCE(Icon));
		lstrcpy(nid.szTip, szAppClassName);
		Shell_NotifyIcon(NIM_ADD, &nid);

		hmenu = CreatePopupMenu(); // 生成菜单
		AppendMenu(hmenu, MF_STRING, IDR_PAUSE, L"停止服务"); // 为菜单添加两个选项
		AppendMenu(hmenu, MF_STRING, IDR_START, L"关于");
		break;
	}
	case WM_USER: //连续使用该程序时候的消息.
		/* if (lParam == WM_LBUTTONDOWN) MessageBox(hwnd, TEXT("Win32 API 实现系统托盘程序,双击托盘可以退出!"), szAppClassName, MB_OK); */
		if (lParam == WM_LBUTTONDBLCLK) // 双击托盘的消息,退出.
			SendMessage(hwnd, WM_CLOSE, wParam, lParam);

		if (lParam == WM_RBUTTONDOWN)
		{
			GetCursorPos(&pt); // 取鼠标坐标
			::SetForegroundWindow(hwnd); // 解决在菜单外单击左键菜单不消失的问题
			// EnableMenuItem(hmenu, IDR_PAUSE, MF_GRAYED); //让菜单中的某一项变灰
			xx = TrackPopupMenu(hmenu, TPM_RETURNCMD, pt.x, pt.y, NULL, hwnd, NULL); // 显示菜单并获取选项ID
			if (xx == IDR_PAUSE)
				SendMessage(hwnd, WM_CLOSE, wParam, lParam);
			if (xx == IDR_START)
				MessageBox(hwnd, TEXT("https://github.com/ming-14"), szAppClassName, MB_OK);
			if (xx == 0)
				PostMessage(hwnd, WM_LBUTTONDOWN, NULL, NULL);
			//MessageBox(hwnd, TEXT("右键"), szAppName, MB_OK);
		}
		break;
	case WM_CLOSE:
		DeleteTray(hwnd);
	case WM_DESTROY: // 窗口销毁时候的消息.
		Shell_NotifyIcon(NIM_DELETE, &nid);
		PostQuitMessage(0);
		break;
	default:
		/*
		 * 防止当Explorer.exe 崩溃以后，程序在系统系统托盘中的图标就消失
		 *
		 * 原理：Explorer.exe 重新载入后会重建系统任务栏。当系统任务栏建立的时候会向系统内所有
		 * 注册接收TaskbarCreated 消息的顶级窗口发送一条消息，我们只需要捕捉这个消息，并重建系
		 * 统托盘的图标即可。
		 */
		if (message == WM_TASKBARCREATED)
			SendMessage(hwnd, WM_CREATE, wParam, lParam);
		break;
	}
	return DefWindowProc(hwnd, message, wParam, lParam);
}

INT WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PSTR lpCmdLine, INT nCmdShow)
{
	GhInstance = hInstance;
	HWND hwnd;
	MSG msg;
	WNDCLASS wndclass;
	/*
	HWND handle = FindWindow(NULL, szAppWindowName);
	if (handle != NULL)
	{
		MessageBox(NULL, TEXT("Application is already running"), szAppClassName, MB_ICONERROR);
		return 0;
	}
	*/
	wndclass.style = CS_HREDRAW | CS_VREDRAW;
	wndclass.lpfnWndProc = WndProc;
	wndclass.cbClsExtra = 0;
	wndclass.cbWndExtra = 0;
	wndclass.hInstance = NULL;
	wndclass.hIcon = LoadIcon(NULL, IDI_APPLICATION);
	wndclass.hCursor = LoadCursor(NULL, IDC_ARROW);
	wndclass.hbrBackground = (HBRUSH)GetStockObject(WHITE_BRUSH);
	wndclass.lpszMenuName = NULL;
	wndclass.lpszClassName = szAppClassName;
	if (!RegisterClass(&wndclass))
	{
		MessageBox(NULL, TEXT("This program requires Windows NT!"), szAppClassName, MB_ICONERROR);
		return 0;
	}

	// 此处使用WS_EX_TOOLWINDOW 属性来隐藏显示在任务栏上的窗口程序按钮 
	hwnd = CreateWindowEx(WS_EX_TOOLWINDOW, szAppClassName, szAppWindowName, WS_POPUP, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, NULL, NULL);
	ShowWindow(hwnd, 0);
	UpdateWindow(hwnd);

	//消息循环
	while (GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg); DispatchMessage(&msg);
	}
	return 0;
}
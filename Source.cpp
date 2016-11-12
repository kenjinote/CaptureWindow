#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#include <windows.h>
#include "resource.h"

TCHAR szClassName[] = TEXT("Window");
HBITMAP hBitmap1, hBitmap2;
HCURSOR hCursor;

void CaptureWindowToClipboard(HWND hWnd)
{
	BOOL bRet = FALSE;
	typedef BOOL(WINAPI *PPRINTWINDOW)(HWND, HDC, UINT);
	PPRINTWINDOW pPrintWindow;
	HMODULE hDLL = ::LoadLibrary(TEXT("user32"));
	if (hDLL) {
		pPrintWindow = (PPRINTWINDOW)::GetProcAddress(hDLL, "PrintWindow");
		if (pPrintWindow) {
			RECT rect;
			GetWindowRect(hWnd, &rect);
			SetCursor(LoadCursor(NULL, IDC_ARROW));
			HDC hdc = GetWindowDC(hWnd);
			HDC hMem = CreateCompatibleDC(hdc);
			HBITMAP hBitmap = CreateCompatibleBitmap(hdc, rect.right - rect.left, rect.bottom - rect.top);
			HBITMAP hOldBitmap;
			if (hBitmap)
			{
				hOldBitmap = (HBITMAP)SelectObject(hMem, hBitmap);
				if (pPrintWindow(hWnd, hMem, 0) == 0)
					StretchBlt(hMem, 0, 0, rect.right - rect.left, rect.bottom - rect.top, hdc, 0, 0, rect.right - rect.left, rect.bottom - rect.top, SRCCOPY);
				OpenClipboard(hWnd);
				EmptyClipboard();
				SetClipboardData(CF_BITMAP, hBitmap);
				CloseClipboard();
				SelectObject(hMem, hOldBitmap);
				DeleteObject(hBitmap);
			}
			DeleteDC(hMem);
			ReleaseDC(hWnd, hdc);
		}
		FreeLibrary(hDLL);
	}
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static BOOL bCapture = FALSE;
	switch (msg) {
	case WM_CREATE:
		CreateWindow(TEXT("STATIC"), TEXT(""), WS_CHILD | WS_VISIBLE | SS_BITMAP | SS_NOTIFY, 10, 10, 42, 42, hWnd, (HMENU)1, ((LPCREATESTRUCT)(lParam))->hInstance, NULL);
		hBitmap1 = LoadBitmap(((LPCREATESTRUCT)(lParam))->hInstance, MAKEINTRESOURCE(IDB_BITMAP_FINDER_FILLED));
		hBitmap2 = LoadBitmap(((LPCREATESTRUCT)(lParam))->hInstance, MAKEINTRESOURCE(IDB_BITMAP_FINDER_EMPTY));
		hCursor = LoadCursor(((LPCREATESTRUCT)(lParam))->hInstance, MAKEINTRESOURCE(IDC_CURSOR_SEARCH_WINDOW));
		SendDlgItemMessage(hWnd, 1, (UINT)STM_SETIMAGE, (WPARAM)IMAGE_BITMAP, (LPARAM)hBitmap1);
		break;
	case WM_COMMAND:
		if (LOWORD(wParam) == 1) {
			SendDlgItemMessage(hWnd, 1, (UINT)STM_SETIMAGE, (WPARAM)IMAGE_BITMAP, (LPARAM)hBitmap2);
			bCapture = TRUE;
			SetCursor(hCursor);
			SetCapture(hWnd);
		}
		break;
	case WM_LBUTTONUP:
		if (bCapture) {
			ReleaseCapture();
			SetCursor(LoadCursor(NULL, IDC_ARROW));
			bCapture = FALSE;
			SendDlgItemMessage(hWnd, 1, (UINT)STM_SETIMAGE, (WPARAM)IMAGE_BITMAP, (LPARAM)hBitmap1);
			POINT point;
			GetCursorPos(&point);
			CaptureWindowToClipboard(WindowFromPoint(point));
		}
		break;
	case WM_DESTROY:
		DeleteObject(hBitmap1);
		DeleteObject(hBitmap2);
		PostQuitMessage(0);
		break;
	case WM_NCLBUTTONDOWN:
	case WM_NCRBUTTONDOWN:
	case WM_LBUTTONDOWN:
	case WM_RBUTTONDOWN:
		SetForegroundWindow(hWnd);
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	WNDCLASS wndclass = {
		CS_HREDRAW | CS_VREDRAW,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0,IDC_ARROW),
		(HBRUSH)(COLOR_WINDOW + 1),
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("ウィンドウをキャプチャー（クリップボード版）"),
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	while (GetMessage(&msg, 0, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return (int)msg.wParam;
}

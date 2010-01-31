// CoSnapsie.cpp : Implementation of CCoSnapsie

#include "stdafx.h"
#include "CoSnapsie.h"
#include <mshtml.h>
#include <exdisp.h>
#include <assert.h>
#include <windows.h>
#include <shlguid.h>
#include <shlguid.h>
#include <shlobj.h>
#include <exdisp.h>
#include <gdiplus.h>

LRESULT CALLBACK MyProc(HWND, UINT, WPARAM, LPARAM);

WNDPROC originalProc;

// Define a shared data segment.  Variables in this segment can be shared across processes that load this DLL.
#pragma data_seg("SHARED")
HHOOK nextHook = NULL;
HWND ie = NULL;
int maxWidth = 0;
int maxHeight = 0;
#pragma data_seg()

#pragma comment(linker, "/section:SHARED,RWS")

// Microsoft linker helper for getting the DLL's HINSTANCE.
// See http://blogs.msdn.com/oldnewthing/archive/2004/10/25/247180.aspx for more details.
//
// If you need to link with a non-MS linker, you would have to add code to look up the DLL's
// path in HKCR.  regsvr32 stores the path under the CLSID key for the 'Snapsie.CoSnapsie' interface.
EXTERN_C IMAGE_DOS_HEADER __ImageBase;

using namespace Gdiplus;

int GetEncoderClsid(const WCHAR* format, CLSID* pClsid)
{
   UINT  num = 0;          // number of image encoders
   UINT  size = 0;         // size of the image encoder array in bytes

   ImageCodecInfo* pImageCodecInfo = NULL;

   GetImageEncodersSize(&num, &size);
   if(size == 0)
      return -1;  // Failure

   pImageCodecInfo = (ImageCodecInfo*)(malloc(size));
   if(pImageCodecInfo == NULL)
      return -1;  // Failure

   GetImageEncoders(num, size, pImageCodecInfo);

   for(UINT j = 0; j < num; ++j)
   {
      if( wcscmp(pImageCodecInfo[j].MimeType, format) == 0 )
      {
         *pClsid = pImageCodecInfo[j].Clsid;
         free(pImageCodecInfo);
         return j;  // Success
      }    
   }

   free(pImageCodecInfo);
   return -1;  // Failure
}

STDMETHODIMP CCoSnapsie::InterfaceSupportsErrorInfo(REFIID riid)
{
    static const IID* arr[] = 
    {
        &IID_ISnapsie
    };

    for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
    {
        if (InlineIsEqualGUID(*arr[i],riid))
            return S_OK;
    }
    return S_FALSE;
}

// Taken from MSDN: ms-help://MS.MSDNQTR.v80.en/MS.MSDN.v80/MS.WIN32COM.v10.en/debug/base/retrieving_the_last_error_code.htm
void PrintError(LPTSTR lpszFunction)
{
	TCHAR szBuf[80]; 
    LPVOID lpMsgBuf;
    DWORD dw = GetLastError(); 

    FormatMessage(
        FORMAT_MESSAGE_ALLOCATE_BUFFER | 
        FORMAT_MESSAGE_FROM_SYSTEM,
        NULL,
        dw,
        MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
        (LPTSTR) &lpMsgBuf,
        0, NULL );

    wsprintf(szBuf, 
        L"%s failed with error %d: %s", 
        lpszFunction, dw, lpMsgBuf); 
 
    MessageBox(NULL, szBuf, L"Error", MB_OK); 

    LocalFree(lpMsgBuf);
}

/**
 * Saves a rendering of the current site by its host container as a PNG file.
 * This implementation is derived from IECapt.
 *
 * @param outputFile  the file to save the PNG output as
 * @link http://iecapt.sourceforge.net/
 */
STDMETHODIMP CCoSnapsie::saveSnapshot(
    BSTR outputFile,
    BSTR frameId,
    LONG drawableScrollWidth,
    LONG drawableScrollHeight,
    LONG drawableClientWidth,
    LONG drawableClientHeight,
    LONG drawableClientLeft,
    LONG drawableClientTop,
    LONG frameBCRLeft,
    LONG frameBCRTop)
{
    HRESULT hr;
    HWND hwndBrowser;

    CComPtr<IOleClientSite>     spClientSite;
    CComQIPtr<IServiceProvider> spISP;
    CComPtr<IWebBrowser2>       spBrowser;
    CComPtr<IDispatch>          spDispatch; 
    CComQIPtr<IHTMLDocument2>   spDocument;
	CComQIPtr<IHTMLDocument3>   doc3;
    CComPtr<IHTMLWindow2>       spScrollableWindow;
    CComQIPtr<IViewObject2>     spViewObject;
    CComPtr<IHTMLStyle>         spStyle;
    CComQIPtr<IHTMLElement2>    spScrollableElement;

	CComVariant documentHeight;
	CComVariant documentWidth;
	CComVariant viewportHeight;
	CComVariant viewportWidth;

	IHTMLElement       *pElement = (IHTMLElement *) NULL;
	IHTMLElementRender *pRender = (IHTMLElementRender *) NULL;

    CComBSTR overflow;
    long scrollLeft;
    long scrollTop;

    long capturableScrollWidth;
    long capturableScrollHeight;
    long capturableClientWidth;
    long capturableClientHeight;

    GetSite(IID_IUnknown, (void**)&spClientSite);

    if (spClientSite == NULL) {
        Error("Como están, BEACHES!!!");
        return E_FAIL;
    }

    spISP = spClientSite;
    if (spISP == NULL)
        return E_FAIL;

    // from http://support.microsoft.com/kb/257717

    hr = spISP->QueryService(IID_IWebBrowserApp, IID_IWebBrowser2,
         (void **)&spBrowser);

    if (FAILED(hr)) {
        // if we can't query the client site for IWebBrowser2, we're probably
        // in an HTA. Obtain the IHTMLWindow2 interface pointer by directly
        // querying the client site's service provider.
        // http://groups.google.com/group/microsoft.public.vc.language/browse_thread/thread/f8987a31d47cccfe/884cb8f13423039e
        CComPtr<IHTMLWindow2> spWindow;
        hr = spISP->QueryService(IID_IHTMLWindow2, &spWindow);
        if (FAILED(hr)) {
            Error("Failed to obtain IHTMLWindow2 from service provider");
            return E_FAIL;
        }

        hr = spWindow->get_document(&spDocument);
        if (FAILED(hr)) {
            Error("Failed to obtain IHTMLDocument2 from window");
            return E_FAIL;
        }

        CComQIPtr<IOleWindow> spOleWindow = spDocument;
        if (spOleWindow == NULL) {
            Error("Failed to obtain IOleWindow from document");
            return E_FAIL;
        }

        hr = spOleWindow->GetWindow(&hwndBrowser);
        if (FAILED(hr)) {
            Error("Failed to obtain HWND from OLE window");
            return E_FAIL;
        }

        hwndBrowser = GetAncestor(hwndBrowser, GA_ROOTOWNER);
    }
    else {
		hr = spBrowser->get_HWND((long*)&hwndBrowser);
		if (FAILED(hr)) {
			Error("Failed to get HWND for browser (is this a frame?)");
			return E_FAIL;
		}

		ie = GetAncestor(hwndBrowser, GA_ROOTOWNER);

		CComPtr<IDispatch> spDispatch;
		hr = spBrowser->get_Document(&spDispatch);
		if (FAILED(hr))
			return E_FAIL;

		spDocument = spDispatch;
		if (spDocument == NULL)
			return E_FAIL;

		spDocument->get_body(&pElement);

		if (pElement == (IHTMLElement *) NULL)
			return E_FAIL;


		pElement->getAttribute(CComBSTR("scrollHeight"), 0, &documentHeight);
		pElement->getAttribute(CComBSTR("scrollWidth"), 0, &documentWidth);
		pElement->getAttribute(CComBSTR("clientHeight"), 0, &viewportHeight);
		pElement->getAttribute(CComBSTR("clientWidth"), 0, &viewportWidth);

		IHTMLStyle* pStyle;
		hr = pElement->get_style(&pStyle);
		if (FAILED(hr))
		{
			PrintError(L"Getting style");

			return E_FAIL;
		}

		//pStyle->put_borderStyle(CComBSTR("none"));
		//pStyle->put_overflow(CComBSTR("hidden"));

		pElement->QueryInterface(IID_IHTMLElementRender, (void **) &pRender);

		if (pRender == (IHTMLElementRender *) NULL)
			return E_FAIL;



		IServiceProvider* pServiceProvider = NULL;
		if (SUCCEEDED(spBrowser->QueryInterface(
							IID_IServiceProvider, 
							(void**)&pServiceProvider)))
		{
			IOleWindow* pWindow = NULL;
			if (SUCCEEDED(pServiceProvider->QueryService(
							SID_SShellBrowser,
							IID_IOleWindow,
							(void**)&pWindow)))
			{

				if (SUCCEEDED(pWindow->GetWindow(&hwndBrowser)))
				{
					hwndBrowser = FindWindowEx(hwndBrowser, NULL, _T("Shell DocObject View"), NULL);
					if (hwndBrowser)
					{
						hwndBrowser = FindWindowEx(hwndBrowser, NULL, _T("Internet Explorer_Server"), NULL);
					}
				}

				pWindow->Release();
			}
		 
			pServiceProvider->Release();
		} 
    }


	//HDC hdcMemory = CreateCompatibleDC(NULL);
	//HBITMAP hBitmap = CreateCompatibleBitmap(GetDC(hwndBrowser), documentWidth.intVal, documentHeight.intVal);
	//HGDIOBJ hOld = SelectObject(hdcMemory, hBitmap);
	//SendMessage(hwndBrowser, WM_PAINT, (WPARAM)hdcMemory, 0);
	//SendMessage(hwndBrowser, WM_PRINT, (WPARAM) hdcMemory, PRF_CLIENT | PRF_CHILDREN | PRF_OWNED);
	//SendMessage(hwndBrowser, WM_PRINTCLIENT, (WPARAM)hdcMemory, PRF_CLIENT | PRF_CHILDREN | PRF_OWNED | PRF_ERASEBKGND );
	//PrintWindow(hwndBrowser, hdcMemory, PW_CLIENTONLY);

		//pRender->DrawToDC(hdcMemory);

	spDocument->get_parentWindow(&spScrollableWindow);
	
	VARIANT buffering;
	spScrollableWindow->get_offscreenBuffering(&buffering);

	bool blah = buffering.boolVal;

	long width, height;
	spBrowser->get_Width(&width);
	spBrowser->get_Height(&height);

	//assert (false);

	int fullScreenY = GetSystemMetrics(SM_CYFULLSCREEN);
	int maximizedY = GetSystemMetrics(SM_CYMAXIMIZED);

	HWND yo;
	spBrowser->get_HWND((SHANDLE_PTR*)&yo);

	/*
	LONG_PTR hey = GetWindowLongPtr(yo, GWLP_WNDPROC);
	hey = GetWindowLongPtr(hwndBrowser, GWLP_WNDPROC);
	originalProc = (WNDPROC) GetWindowLongPtr(yo, GWLP_WNDPROC);
	*/

	DWORD parentID, thisID, currentID;
	GetWindowThreadProcessId(yo, &parentID);
	GetWindowThreadProcessId(hwndBrowser, &thisID);
	currentID = GetCurrentProcessId();

	TCHAR dllPath[_MAX_PATH];
	GetModuleFileName((HINSTANCE) &__ImageBase, dllPath, _MAX_PATH);

	HINSTANCE hinstDLL = LoadLibrary(dllPath);
	HOOKPROC hkprcSysMsg = (HOOKPROC)GetProcAddress(hinstDLL, "CallWndProc");
	if (hkprcSysMsg == NULL)
		PrintError(L"GetProcAddress");

	//nextHook = SetWindowsHookEx(WH_CALLWNDPROC, CallWndProc, NULL, GetCurrentThreadId());
	nextHook = SetWindowsHookEx(WH_CALLWNDPROC, hkprcSysMsg, hinstDLL, 0);
	//nextHook = SetWindowsHookEx(WH_CALLWNDPROCRET, hkprcSysMsg, hinstDLL, 0);
	//nextHook = SetWindowsHookEx(WH_CBT, CBTProc, NULL, GetCurrentThreadId());
	//nextHook = SetWindowsHookEx(WH_CBT, hkprcSysMsg, hinstDLL, 0);
	//nextHook = SetWindowsHookEx(WH_GETMESSAGE, hkprcSysMsg, hinstDLL, 0);
	if (nextHook == 0)
		PrintError(L"SetWindowsHookEx");

	/*
	originalProc = (WNDPROC) SetWindowLongPtr(ie, GWLP_WNDPROC, (LONG) MyProc);
	if (originalProc == 0)
		PrintError(L"GetWindowLongPtr");
	*/

	//::MoveWindow(ie, 0, 0, documentWidth.intVal, documentHeight.intVal, TRUE);

	/*
	hr = spBrowser->put_Height(documentHeight.intVal);
	if (FAILED(hr))
	{
		PrintError(L"PutHeight");
	}
	spBrowser->put_Width(documentWidth.intVal * 100);
	*/

	long newWidth, newHeight;
	spBrowser->get_Width(&newWidth);
	spBrowser->get_Height(&newHeight);


	// create the HDC objects
    HDC hdcInput = ::GetDC(hwndBrowser);
    if (!hdcInput)
        return E_FAIL;

	/*
	spDocument->QueryInterface(IID_IHTMLDocument3, (void**)&doc3);

	IHTMLElement       *htmlElement = (IHTMLElement *) NULL;
	doc3->get_documentElement(&htmlElement);
	IHTMLHtmlElement* html = (IHTMLHtmlElement *) htmlElement;
	*/

	// Nobody else seems to know how to get IViewObject2?!
	// http://starkravingfinkle.org/blog/2004/09/
	//spViewObject = spDocument;
	spDocument->QueryInterface(IID_IViewObject, (void**)&spViewObject);
	if (spViewObject == NULL)
		return E_FAIL;

	/*
	SIZE bitmapSize;
	bitmapSize.cx = documentWidth.intVal;
	bitmapSize.cy = documentHeight.intVal;

	HBITMAP hBitmap = CreateCompatibleBitmap(hdcInput, documentWidth.intVal, documentHeight.intVal);
	assert(false); 

	IWebBrowser2* myBrowser;
	hr = CoCreateInstance(CLSID_WebBrowser, NULL, CLSCTX_INPROC_SERVER, IID_IWebBrowser2, (void**)&myBrowser);
	if (FAILED(hr))
	{
		PrintError(L"CoCreateInstance");
	}

	myBrowser->put_Visible(VARIANT_TRUE);
	*/
 
	/*
	RECT rc = {0, 0, documentWidth.intVal, documentHeight.intVal};
	BOOL success = SystemParametersInfo(SPI_SETWORKAREA, 0, &rc, 0);
	if (success == 0)
		PrintError(L"SystemParametersInfo");

	MoveWindow(hwndBrowser, 0, 0, documentWidth.intVal, documentHeight.intVal, TRUE);
	*/

	/*
	hr = spBrowser->put_Height(documentHeight.intVal);
	if (FAILED(hr))
		PrintError(L"put_height");

	spBrowser->put_Width(documentWidth.intVal);
	*/

	long myHeight, myWidth;
	spBrowser->get_Height(&myHeight);
	spBrowser->get_Width(&myWidth);

	/*
	CLSID clsid;
	hr = CLSIDFromProgID(L"Shell.ThumbnailExtract.HTML.1", &clsid);
	if (FAILED(hr))
	{
		PrintError(L"CLSID");

		return E_FAIL;
	}

	IThumbnailCapture* capture;
	hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IThumbnailCapture, (void**)&capture);
	if (FAILED(hr))
	{
		PrintError(L"CoCreateInstance");

		return E_FAIL;
	}
	
	capture->CaptureThumbnail(&bitmapSize, spDocument, &hBitmap);
	*/

	/*
	CImage image;
	image.Create(documentWidth.intVal, documentHeight.intVal, 24);
    CImageDC imageDC(image);
	*/

	BITMAPINFOHEADER bih;
    BITMAPINFO bi;
    RGBQUAD rgbquad;

    ZeroMemory(&bih, sizeof(BITMAPINFOHEADER));
    ZeroMemory(&rgbquad, sizeof(RGBQUAD));

    bih.biSize      = sizeof(BITMAPINFOHEADER);
	bih.biWidth     = documentWidth.intVal;
		bih.biHeight    = documentHeight.intVal;
    bih.biPlanes    = 1;
    bih.biBitCount      = 32;
    bih.biClrUsed       = 0;
    bih.biSizeImage     = 0;
    bih.biCompression   = BI_RGB;
    bih.biXPelsPerMeter = 0;
    bih.biYPelsPerMeter = 0;

    bi.bmiHeader = bih;
    bi.bmiColors[0] = rgbquad;

    char* bitmapData = NULL;
    HBITMAP hBitmap = CreateDIBSection(hdcInput, &bi, DIB_RGB_COLORS,
        (void**)&bitmapData, NULL, 0);

	HDC hdcOutput = CreateCompatibleDC(hdcInput);

	
/*
	Graphics graphics(hdcOutput);
	Region clipRegion(Rect(0, 0, documentWidth.intVal, documentHeight.intVal));
	graphics.SetClip(&clipRegion);

    if (!hBitmap) {
        // clean up
        ReleaseDC(hwndBrowser, hdcInput);
        DeleteDC(hdcOutput);

        Error("Failed when creating bitmap");
        return E_FAIL;
    }

    SelectObject(hdcOutput, hBitmap);

	SolidBrush brush(Color(255, 0, 0));
	graphics.FillRegion(&brush, &clipRegion);
	*/

	GdiplusStartupInput gdiplusStartupInput;
	ULONG_PTR gdiplusToken;
	GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);

	Bitmap* bitmap = new Bitmap(documentWidth.intVal, documentHeight.intVal, PixelFormat24bppRGB);
	Graphics* graphics = Graphics::FromImage(bitmap);
	HDC myHDC = graphics->GetHDC();

	Color* color = new Color(255, 0, 0, 255);
	SolidBrush* brush = new SolidBrush(*color);
	graphics->FillEllipse(brush, 20, 30, 80, 50);

//	Region clipRegion(Rect(0, 0, documentWidth.intVal, documentHeight.intVal));
//	graphics.SetClip(&clipRegion);

//	SolidBrush brush(Color(255, 0, 255));
//	graphics.FillRegion(&brush, &clipRegion);

	
	RECT rcBounds = { 0, 0, documentWidth.intVal, documentHeight.intVal };

	/*
	hr = spViewObject->Draw(DVASPECT_CONTENT, -1, NULL, NULL,
                                   hdcInput, hdcOutput, &rcBounds,
                                   NULL, NULL, 0);
								   */

	assert(false);

	//MoveWindow(ie, 0, 0, documentWidth.intVal, documentHeight.intVal, TRUE);MoveWindow(ie, 0, 0, documentWidth.intVal, documentHeight.intVal, TRUE);

	RECT windowRect;
	GetWindowRect(hwndBrowser, &windowRect);

	RECT clientRect;
	GetClientRect(hwndBrowser, &clientRect);

	int chromeWidth = myWidth - viewportWidth.intVal;
	int chromeHeight = windowRect.top * 2;

	maxWidth = documentWidth.intVal + chromeWidth;
	maxHeight = documentHeight.intVal + chromeHeight;

	//spBrowser->put_Visible(VARIANT_FALSE);
	spBrowser->put_Height(maxHeight);
	spBrowser->put_Width(maxWidth);
	//MoveWindow(hwndBrowser, 0, 0, documentWidth.intVal, documentHeight.intVal, TRUE);

	long myNewHeight, myNewWidth;
	spBrowser->get_Height(&myNewHeight);
	spBrowser->get_Width(&myNewWidth);
	
	//VARIANT nuts;
	//nuts.boolVal = true;
	//spScrollableWindow->put_offscreenBuffering(nuts);
	//spScrollableWindow->resizeTo(documentWidth.intVal, documentHeight.intVal);
	hr = OleDraw(spViewObject, DVASPECT_DOCPRINT, myHDC, &rcBounds);
	if (FAILED(hr))
		PrintError(L"OleDraw");

	graphics->ReleaseHDC(myHDC);

	CLSID pngClsid;
	GetEncoderClsid(L"image/png", &pngClsid);
	Status status = bitmap->Save(L"C:\\users\\nirvdrum\\dev\\SnapsIE\\test\\new.png", &pngClsid, NULL);
	if (status != 0)
		PrintError(L"Save");

	spBrowser->put_Height(myHeight);
	spBrowser->put_Width(myWidth);
	//spBrowser->put_Visible(VARIANT_TRUE);

	delete bitmap;
	delete graphics;
	delete color;
	delete brush;

	GdiplusShutdown(gdiplusToken);
		
	//SelectObject(hdcMemory, hOld);
	//DeleteObject(hdcMemory);

	//image.Attach(hBitmap);

	CImage image;
	image.Attach(hBitmap);

	/*
	image.Create(documentWidth.intVal, documentHeight.intVal, 24);
    CImageDC imageDC(image);
	::BitBlt(imageDC, 0, 0, documentWidth.intVal, documentHeight.intVal, hdcOutput, 0, 0, SRCCOPY);
	*/


	//PrintWindow(hwndBrowser, imageDC, PW_CLIENTONLY);

	/*
	html->QueryInterface(IID_IHTMLElementRender, (void **) &pRender);
	pRender->DrawToDC(imageDC);
	*/

	/*
	spBrowser->put_Width(width);
	spBrowser->put_Height(height);
	*/


	if (FAILED(hr))
	{
		PrintError(L"Draw");

		//return E_FAIL;
	}

	 //hr = spViewObject->Draw(DVASPECT_CONTENT, -1, NULL, NULL, hdcInput,
       //         hdcOutput, &rcBounds, NULL, NULL, 0);

    //CRect captureWndRect;
    //GetWindowRect(hwndBrowser, &captureWndRect);

	//HBITMAP hBitmap = CreateCompatibleBitmap(hdcOutput, 2000, 2000);
	//if (!hBitmap)
	//	return E_FAIL;

	//HDC hdcMemory = CreateCompatibleDC(hdcOutput);
	//SelectObject(hdcMemory, hBitmap);
	/*
	::BitBlt(hdcMemory, 0, 0,
				documentWidth.intVal,
				documentHeight.intVal,
                hdcOutput,
                0,
                0,
                SRCCOPY);*/
	//PrintWindow(hwndBrowser, hdcOutput, PW_CLIENTONLY);


	// save the imag

	image.Save(CW2T(outputFile));

    // clean up
    ::ReleaseDC(hwndBrowser, hdcInput);

	UnhookWindowsHookEx(nextHook);

    return hr;
}

// Many thanks to sunnyandy for helping out with this approach.  What we're doing here is setting up
// a Windows hook to see incoming messages to the IEFrame's message processor.  Once we find one that's
// WM_GETMINMAXINFO, we inject our own message processor into the IEFrame process to handle that one
// message.  WM_GETMINMAXINFO is sent on a resize event so the process can see how large a window can be.
// By modifying the max values, we can allow a window to be sized greater than the (virtual) screen resolution
// would otherwise allow.
//
// See the discussion here: http://www.codeguru.com/forum/showthread.php?p=1889928

LRESULT WINAPI CallWndProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	DWORD procID = GetCurrentProcessId();
	int i = 0;

	CWPSTRUCT* cw = (CWPSTRUCT*) lParam;
	UINT message = cw->message;

	if (message == WM_GETMINMAXINFO)
	{
		MINMAXINFO* minMaxInfo = (MINMAXINFO*) cw->lParam;

		LONG_PTR proc = SetWindowLongPtr(cw->hwnd, GWL_WNDPROC, (LONG_PTR) MyProc);
        SetProp(cw->hwnd, L"_old_proc_", (HANDLE) proc);
	}

	return CallNextHookEx(nextHook, nCode, wParam, lParam);
}

LRESULT CALLBACK MyProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	HANDLE hdl = GetProp(hwnd, L"_old_proc_");
	RemoveProp(hwnd, L"_old_proc_");
	SetWindowLongPtr(hwnd, GWL_WNDPROC, (LONG_PTR) (hdl));

	switch (message)
	{
		case WM_GETMINMAXINFO:
		{
			MINMAXINFO* minMaxInfo = (MINMAXINFO*) (lParam);

			minMaxInfo->ptMaxSize.x = maxWidth;
			minMaxInfo->ptMaxSize.y = maxHeight;
			minMaxInfo->ptMaxTrackSize.x = maxWidth;
			minMaxInfo->ptMaxTrackSize.y = maxHeight;

			return 0;
		}

		default:
		{
			return CallWindowProc((WNDPROC) hdl, hwnd, message,	wParam,	lParam);
		}
	}
}

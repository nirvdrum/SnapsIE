// CoSnapsie.cpp : Implementation of CCoSnapsie

#include "stdafx.h"
#include "CoSnapsie.h"
#include <mshtml.h>
#include <exdisp.h>
#include <assert.h>
#include <windows.h>
#include <shlguid.h>

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


    // create the HDC objects
    HDC hdcInput = ::GetDC(hwndBrowser);
    if (!hdcInput)
        return E_FAIL;

    // Nobody else seems to know how to get IViewObject2?!
    // http://starkravingfinkle.org/blog/2004/09/
    //spViewObject = spDocument;
    spDocument->QueryInterface(IID_IViewObject, (void**)&spViewObject);
    if (spViewObject == NULL)
        return E_FAIL;
 

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


	long originalHeight, originalWidth;
    spBrowser->get_Height(&originalHeight);
    spBrowser->get_Width(&originalWidth);

	// Get the path to this DLL so we can load it up with LoadLibrary.
    TCHAR dllPath[_MAX_PATH];
    GetModuleFileName((HINSTANCE) &__ImageBase, dllPath, _MAX_PATH);

	// Get the path to the Windows hook we use to allow resizing the window greater than the virtual screen resolution.
    HINSTANCE hinstDLL = LoadLibrary(dllPath);
    HOOKPROC hkprcSysMsg = (HOOKPROC)GetProcAddress(hinstDLL, "CallWndProc");
    if (hkprcSysMsg == NULL)
        PrintError(L"GetProcAddress");

	// Install the Windows hook.
    nextHook = SetWindowsHookEx(WH_CALLWNDPROC, hkprcSysMsg, hinstDLL, 0);
    if (nextHook == 0)
        PrintError(L"SetWindowsHookEx");



	CImage image;
    image.Create(documentWidth.intVal, documentHeight.intVal, 24);
    CImageDC imageDC(image);
    
    RECT rcBounds = { 0, 0, documentWidth.intVal, documentHeight.intVal };

    RECT windowRect;
    GetWindowRect(hwndBrowser, &windowRect);

    RECT clientRect;
    GetClientRect(hwndBrowser, &clientRect);

    int chromeWidth = originalWidth - viewportWidth.intVal;
    int chromeHeight = windowRect.top * 2;

    maxWidth = documentWidth.intVal + chromeWidth;
    maxHeight = documentHeight.intVal + chromeHeight;

    spBrowser->put_Height(maxHeight);
    spBrowser->put_Width(maxWidth);
    
	hr = PrintWindow(hwndBrowser, imageDC, PW_CLIENTONLY);
    // hr = OleDraw(spViewObject, DVASPECT_DOCPRINT, myHDC, &rcBounds);
    /*
    hr = spViewObject->Draw(DVASPECT_CONTENT, -1, NULL, NULL,
                                   hdcInput, hdcOutput, &rcBounds,
                                   NULL, NULL, 0);
                                   */


    if (FAILED(hr))
        PrintError(L"OleDraw");
	UnhookWindowsHookEx(nextHook);

    spBrowser->put_Height(originalHeight);
    spBrowser->put_Width(originalWidth);


    // save the imag
    image.Save(CW2T(outputFile));

    // clean up
    ::ReleaseDC(hwndBrowser, hdcInput);

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
    HANDLE originalMessageProc = GetProp(hwnd, L"_old_proc_");
    RemoveProp(hwnd, L"_old_proc_");
    SetWindowLongPtr(hwnd, GWL_WNDPROC, (LONG_PTR) originalMessageProc);

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
            return CallWindowProc((WNDPROC) originalMessageProc, hwnd, message,	wParam,	lParam);
        }
    }
}

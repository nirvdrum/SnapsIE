// CoSnapsie.cpp : Implementation of CCoSnapsie

#include "stdafx.h"
#include "CoSnapsie.h"
#include <mshtml.h>
#include <exdisp.h>
#include <assert.h>
#include <windows.h>
#include <shlguid.h>

LRESULT CALLBACK MinMaxInfoHandler(HWND, UINT, WPARAM, LPARAM);

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
        Error("Como est�n, BEACHES!!!");
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

    // Nobody else seems to know how to get IViewObject2?!
    // http://starkravingfinkle.org/blog/2004/09/
    //spViewObject = spDocument;
    spDocument->QueryInterface(IID_IViewObject, (void**)&spViewObject);
    if (spViewObject == NULL)
        return E_FAIL;


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


	// Figure out how large to make the window.  It's no sufficient to just use the dimensions of the scrolled
	// viewport because the browser chrome occupies space that must be accounted for as well.
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


	// Capture the window's canvas to a DIB.
	CImage image;
	image.Create(documentWidth.intVal, documentHeight.intVal, 24);
    CImageDC imageDC(image);
    
    RECT rcBounds = { 0, 0, documentWidth.intVal, documentHeight.intVal };
    
	hr = PrintWindow(hwndBrowser, imageDC, PW_CLIENTONLY);
    //hr = OleDraw(spViewObject, DVASPECT_DOCPRINT, imageDC, &rcBounds);

    if (FAILED(hr))
        PrintError(L"OleDraw");

	UnhookWindowsHookEx(nextHook);

    spBrowser->put_Height(originalHeight);
    spBrowser->put_Width(originalWidth);


    // save the image.
    image.Save(CW2T(outputFile));

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
    CWPSTRUCT* cwp = (CWPSTRUCT*) lParam;

    if (WM_GETMINMAXINFO == cwp->message)
	{
		// Inject our own message processor into the process so we can modify the WM_GETMINMAXINFO message.
		// It is not possible to modify the message from this hook, so the best we can do is inject a function that can.
        LONG_PTR proc = SetWindowLongPtr(cwp->hwnd, GWL_WNDPROC, (LONG_PTR) MinMaxInfoHandler);
        SetProp(cwp->hwnd, L"__original_message_processor__", (HANDLE) proc);
    }

    return CallNextHookEx(nextHook, nCode, wParam, lParam);
}

// This function is our message processor that we inject into the IEFrame process.  Its sole purpose
// is to process WM_GETMINMAXINFO messages and modify the max tracking size so that we can resize the
// IEFrame window to greater than the virtual screen resolution.  All other messages are delegated to
// the original IEFrame message processor.  This function uninjects itself immediately upon execution.
LRESULT CALLBACK MinMaxInfoHandler(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	// Grab a reference to the original message processor.
    HANDLE originalMessageProc = GetProp(hwnd, L"__original_message_processor__");
    RemoveProp(hwnd, L"__original_message_processor__");

	// Uninject this method.
    SetWindowLongPtr(hwnd, GWL_WNDPROC, (LONG_PTR) originalMessageProc);

	if (WM_GETMINMAXINFO == message)
	{
		MINMAXINFO* minMaxInfo = (MINMAXINFO*) (lParam);

        minMaxInfo->ptMaxTrackSize.x = maxWidth;
        minMaxInfo->ptMaxTrackSize.y = maxHeight;

		// We're not going to pass this message onto the original message processor, so we should
		// return 0, per the documentation for the WM_GETMINMAXINFO message.
        return 0;
	}

    // All other messages should be handled by the original message processor.
    return CallWindowProc((WNDPROC) originalMessageProc, hwnd, message, wParam, lParam);
}

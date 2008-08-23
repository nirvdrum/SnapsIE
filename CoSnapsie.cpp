// CoSnapsie.cpp : Implementation of CCoSnapsie

#include "stdafx.h"
#include "CoSnapsie.h"
#include <mshtml.h>
#include <exdisp.h>


// CCoSnapsie

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
    CComPtr<IHTMLWindow2>       spScrollableWindow;
    CComQIPtr<IViewObject2>     spViewObject;
    CComPtr<IHTMLStyle>         spStyle;
    CComQIPtr<IHTMLElement2>    spScrollableElement;

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

        CComPtr<IDispatch> spDispatch;
        hr = spBrowser->get_Document(&spDispatch);
        if (FAILED(hr))
            return E_FAIL;

        spDocument = spDispatch;
        if (spDocument == NULL)
            return E_FAIL;
    }

    // we could be rendering the base document, or a frame. We need to scroll
    // the window of that document or frame, so get the right one here.

    if (SysStringLen(frameId) > 0) {
        CComQIPtr<IHTMLDocument3>  spDocument3;
        CComPtr<IHTMLElement>      spFrame;
        CComQIPtr<IHTMLFrameBase2> spFrameBase;
        CComQIPtr<IWebBrowser2>    spFramedBrowser;
        CComPtr<IDispatch>         spFramedDispatch;
        CComPtr<IHTMLDocument2>    spFramedDocument;
        CComPtr<IHTMLDocument3>    spFramedDocument3;
        CComQIPtr<IHTMLDocument5>  spFramedDocument5;
        CComPtr<IHTMLElement>      spFramedElement;

        spDocument3 = spDocument;
        if (spDocument3 == NULL)
            return E_FAIL;

        hr = spDocument3->getElementById(frameId, &spFrame);
        if (FAILED(hr))
            return E_FAIL;

        spFramedBrowser = spFrame;
        if (spFramedBrowser == NULL)
            return E_FAIL;

        hr = spFramedBrowser->get_Document(&spFramedDispatch);
        if (FAILED(hr))
            return E_FAIL;

        spFramedDocument = spFramedDispatch;
        if (spFramedDocument == NULL)
            return E_FAIL;

        spFramedDocument5 = spFramedDocument;
        if (spFramedDocument5 == NULL)
            return E_FAIL;

        CComBSTR compatMode;
        spFramedDocument5->get_compatMode(&compatMode);

        if (compatMode == L"BackCompat") {
            hr = spFramedDocument->get_parentWindow(&spScrollableWindow);
            if (FAILED(hr))
                return E_FAIL;

            hr = spFramedDocument->get_body(&spFramedElement);
        }
        else {
            spFrameBase = spFrame;
            if (spFrameBase == NULL)
                return E_FAIL;

            hr = spFrameBase->get_contentWindow(&spScrollableWindow);
            if (FAILED(hr))
                return E_FAIL;

            spFramedDocument3 = spFramedDocument;
            if (spFramedDocument3 == NULL)
                return E_FAIL;

            hr = spFramedDocument3->get_documentElement(&spFramedElement);
        }
        if (spFramedElement == NULL)
            return E_FAIL;

        // we have to go through this ridiculous shizzle because the correct
        // scroll dimensions of the frame are not available to javascript;
        // neither are we able to dynamically hide its scrollbars. It seems
        // we can't access the body element of the framed document; we always
        // get back the body of the top level document instead. Accessing it
        // through COM interfaces seems fine though.

        spScrollableElement = spFramedElement;
        if (spScrollableElement == NULL)
            return E_FAIL;

        hr = spScrollableElement->get_runtimeStyle(&spStyle);
        if (FAILED(hr))
            return E_FAIL;

        spStyle->get_overflow(&overflow);
        spStyle->put_overflow(CComBSTR("hidden"));

        spScrollableElement->get_scrollLeft(&scrollLeft);
        spScrollableElement->get_scrollTop(&scrollTop);

        spScrollableElement->get_scrollWidth(&capturableScrollWidth);
        spScrollableElement->get_scrollHeight(&capturableScrollHeight);
        spScrollableElement->get_clientWidth(&capturableClientWidth);
        spScrollableElement->get_clientHeight(&capturableClientHeight);
    }
    else {
        hr = spDocument->get_parentWindow(&spScrollableWindow);

        capturableScrollWidth = drawableScrollWidth;
        capturableScrollHeight = drawableScrollHeight;
        capturableClientWidth = drawableClientWidth;
        capturableClientHeight = drawableClientHeight;
    }

    if (FAILED(hr)) {
        Error("Failed to get parent window for document");
        return E_FAIL;
    }

    // THANK YOU Mark Finkle!
    // Nobody else seems to know how to get IViewObject2?!
    // http://starkravingfinkle.org/blog/2004/09/

    spViewObject = spDocument;
    if (spViewObject == NULL)
        return E_FAIL;

    // create the HDC objects

    HDC hdcInput = ::GetDC(hwndBrowser);
    if (!hdcInput)
        return E_FAIL;

    HDC hdcOutput = CreateCompatibleDC(hdcInput);
    if (!hdcOutput)
        return E_FAIL;

    // initialize bitmap

    BITMAPINFOHEADER bih;
    BITMAPINFO bi;
    RGBQUAD rgbquad;

    ZeroMemory(&bih, sizeof(BITMAPINFOHEADER));
    ZeroMemory(&rgbquad, sizeof(RGBQUAD));

    bih.biSize      = sizeof(BITMAPINFOHEADER);
    bih.biWidth     = drawableScrollWidth  + (2 * drawableClientLeft);
    bih.biHeight    = drawableScrollHeight + (2 * drawableClientTop);
    bih.biPlanes    = 1;
    bih.biBitCount      = 24;
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

    if (!hBitmap) {
        // clean up
        ReleaseDC(hwndBrowser, hdcInput);
        DeleteDC(hdcOutput);

        Error("Failed when creating bitmap");
        return E_FAIL;
    }

    SelectObject(hdcOutput, hBitmap);

    CImage image;
    image.Create(capturableScrollWidth, capturableScrollHeight, 24);
    CImageDC imageDC(image);

    // main drawing loops (pan and scan)

    long xEnd;
    long yEnd;

    for (long yStart = 0; yStart < capturableScrollHeight; yStart += capturableClientHeight) {
        yEnd = yStart + capturableClientHeight;
        if (yEnd > capturableScrollHeight) {
            if (yStart != 0) {
                yStart = capturableScrollHeight - capturableClientHeight;
            }
            yEnd = capturableScrollHeight;
        }

        for (long xStart = 0; xStart < capturableScrollWidth; xStart += capturableClientWidth) {
            xEnd = xStart + capturableClientWidth;
            if (xEnd > capturableScrollWidth) {
                if (xStart != 0) {
                    xStart = capturableScrollWidth - capturableClientWidth;
                }
                xEnd = capturableScrollWidth;
            }

            //char message[1024];
            //sprintf_s(message, "Capturing region: (%d,%d)->(%d,%d)\n", xStart, yStart, xEnd, yEnd);
            //CComBSTR b = message;
            //MessageBox(NULL, b, L"Debug Info", MB_OK | MB_SETFOREGROUND);

            // scroll the window, and draw the results. Here's the key: the
            // device we draw on isn't sized to the client dimensions; it's
            // sized to the client dimensions PLUS the left or top offset.

            hr = spScrollableWindow->scroll(xStart, yStart);
            if (FAILED(hr)) {
                Error("Failed to scroll window");
                goto loop_exit;
            }

            RECTL rcBounds = { 0, 0, drawableClientWidth + (2 * drawableClientLeft),
                drawableClientHeight + (2 * drawableClientTop) };

            hr = spViewObject->Draw(DVASPECT_CONTENT, -1, NULL, NULL, hdcInput,
                hdcOutput, &rcBounds, NULL, NULL, 0);
            if (FAILED(hr)) {
                Error("Draw() failed during pan and scan");
                goto loop_exit;
            }

            // copy to correct region of output device

            ::BitBlt(imageDC, xStart, yStart,
                (xEnd - xStart),
                (yEnd - yStart),
                hdcOutput,
                (drawableClientLeft + frameBCRLeft),
                (drawableClientTop + frameBCRTop),
                SRCCOPY);
        }
    }
    loop_exit:

    // save the image

    image.Save(CW2T(outputFile));

    // clean up

    ::ReleaseDC(hwndBrowser, hdcInput);
    DeleteDC(hdcOutput);
    DeleteObject(hBitmap);

    // revert the scrollbars

    if (SysStringLen(frameId) > 0) {
        spStyle->put_overflow(overflow);
        spScrollableElement->put_scrollLeft(scrollLeft);
        spScrollableElement->put_scrollTop(scrollTop);
    }

    return hr;
}

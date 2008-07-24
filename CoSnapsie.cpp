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
    LONG drawableScrollWidth,
    LONG drawableScrollHeight,
    LONG drawableClientWidth,
    LONG drawableClientHeight,
    LONG drawableClientLeft,
    LONG drawableClientTop,
    LONG capturableScrollWidth,
    LONG capturableScrollHeight,
    LONG capturableClientWidth,
    LONG capturableClientHeight,
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
    CComPtr<IHTMLWindow2>       spWindow;
    CComQIPtr<IViewObject2>     spViewObject;

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
    if (FAILED(hr))
        return E_FAIL;

    hr = spBrowser->get_HWND((long*)&hwndBrowser);
    if (FAILED(hr)) {
        Error("Failed to get HWND for browser (is this a frame?)");
        return E_FAIL;
    }

    hr = spBrowser->get_Document(&spDispatch);
    if (FAILED(hr))
        return E_FAIL;

    spDocument = spDispatch;
    if (spDocument == NULL)
        return E_FAIL;

    hr = spDocument->get_parentWindow(&spWindow);
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

            hr = spWindow->scroll(xStart, yStart);
            if (FAILED(hr))
                goto loop_exit;

            RECTL rcBounds = { 0, 0, drawableClientWidth + (2 * drawableClientLeft),
                drawableClientHeight + (2 * drawableClientTop) };

            hr = spViewObject->Draw(DVASPECT_CONTENT, -1, NULL, NULL, hdcInput,
                hdcOutput, &rcBounds, NULL, NULL, 0);
            if (FAILED(hr)) {
                Error("Draw() failed during pan and scan foo");
                goto loop_exit;
            }

            // copy to correct region of output device

            ::BitBlt(imageDC, xStart, yStart,
                (xEnd - xStart),
                (yEnd - yStart),
                hdcOutput, drawableClientLeft, drawableClientTop, SRCCOPY);
        }
    }
    loop_exit:

    // save the image

    if (FAILED(hr)) {
        Error("Failed during draw phase");
        return E_FAIL;
    }

    image.Save(CW2T(outputFile));

    // clean up

    ::ReleaseDC(hwndBrowser, hdcInput);
    DeleteDC(hdcOutput);
    DeleteObject(hBitmap);

    return hr;
}

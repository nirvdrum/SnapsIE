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
 * @param outputPath  the path to the file to save the PNG output as
 * @link http://iecapt.sourceforge.net/
 */
STDMETHODIMP CCoSnapsie::saveSnapshot(BSTR outputPath, BSTR frameId)
{
    HRESULT hr;
    HWND hwndBrowser;
    CComPtr<IOleClientSite>     spClientSite;
    CComQIPtr<IServiceProvider> spISP;
    CComPtr<IWebBrowser2>       spBrowser;
    CComPtr<IDispatch>          spDispatch; 
    CComQIPtr<IHTMLDocument2>   spDocument;
    CComPtr<IHTMLElement>       spBody;
    CComQIPtr<IHTMLDocument3>   spDocument3;
    CComPtr<IHTMLElement>       spElement;
    CComQIPtr<IHTMLDocument5>   spDocument5;
    CComQIPtr<IHTMLElement2>    spDimensionElement;
    CComPtr<IHTMLStyle>         spStyle;

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
        Error("Could not get HWND for browser (is this a frame?)");
        return E_FAIL;
    }

    hr = spBrowser->get_Document(&spDispatch);
    if (FAILED(hr))
        return E_FAIL;

    spDocument = spDispatch;
    if (spDocument == NULL)
        return E_FAIL;

    spDocument3 = spDispatch;
    if (spDocument3 == NULL)
        return E_FAIL;

    // we could be rendering the base document, or a frame. The view object
    // is retrieved from the document, so get the right document here.

    if (SysStringLen(frameId) > 0) {
        CComPtr<IHTMLElement>   spFramedElement;
        CComQIPtr<IWebBrowser2> spFramedBrowser;
        CComPtr<IDispatch>      spFramedDispatch;

        hr = spDocument3->getElementById(frameId, &spFramedElement);
        if (FAILED(hr))
            return E_FAIL;

        spFramedBrowser = spFramedElement;
        if (spFramedBrowser == NULL) {
            Error("Supposed frame element is not a browser");
            return E_FAIL;
        }

        hr = spFramedBrowser->get_Document(&spFramedDispatch);
        if (FAILED(hr))
            return E_FAIL;

        spDocument = spFramedDispatch;
        if (spDocument == NULL)
            return E_FAIL;

        spDocument3 = spFramedDispatch;
        if (spDocument3 == NULL)
            return E_FAIL;
    }

    hr = spDocument->get_body(&spBody);
    if (spBody == NULL)
        return E_FAIL;

    hr = spDocument3->get_documentElement(&spElement);
    if (FAILED(hr))
        return E_FAIL;

    spDocument5 = spDocument3;
    if (spDocument5 == NULL) {
        Error("Could not get IHTMLDocument5");
        return E_FAIL;
    }

    // try to deal with the different interpretations of the scroll and client
    // dimensions, depending on the IE rendering mode.
    // http://lists.w3.org/Archives/Public/www-archive/2007Aug/att-0003/offset-mess.htm
    // http://msdn2.microsoft.com/en-us/library/bb250395.aspx

    CComBSTR compatMode;
    spDocument5->get_compatMode(&compatMode);
    if (compatMode == L"BackCompat") {
        // quirks mode
        spDimensionElement = spBody;
    }
    else {
        // standards mode
        spDimensionElement = spElement;
    }

    hr = spDimensionElement->get_runtimeStyle(&spStyle);
    if (FAILED(hr))
        return E_FAIL;

    // be sure to hide the scrollbars *before* calculating the scroll and
    // client dimensions. We need to restore them after the fact.

    CComBSTR overflow;

    spStyle->get_overflow(&overflow);
    spStyle->put_overflow(CComBSTR("hidden"));

    long scrollWidth;
    long scrollHeight;
    long clientWidth;
    long clientHeight;
    long clientLeft;
    long clientTop;

    spDimensionElement->get_scrollWidth(&scrollWidth);
    spDimensionElement->get_scrollHeight(&scrollHeight);
    spDimensionElement->get_clientWidth(&clientWidth);
    spDimensionElement->get_clientHeight(&clientHeight);
    spDimensionElement->get_clientLeft(&clientLeft);
    spDimensionElement->get_clientTop(&clientTop);

    // record the current scroll position, and restore it afterwards. Also
    // restore scrollbars, if necessary.

    long scrollLeft;
    long scrollTop;

    spDimensionElement->get_scrollLeft(&scrollLeft);
    spDimensionElement->get_scrollTop(&scrollTop);

    hr = panAndScan(hwndBrowser, spDocument, outputPath,
                    scrollWidth, scrollHeight,
                    clientWidth, clientHeight,
                    clientLeft , clientTop);

    spDimensionElement->put_scrollLeft(scrollLeft);
    spDimensionElement->put_scrollTop(scrollTop);

    spStyle->put_overflow(overflow);

    return hr;
}

/**
 * This method copies the rendered HTML content of a browser object to a PNG
 * output file, using the outputPath parameter. If the client dimensions are
 * smaller than the scroll dimensions (i.e. some part of the whole image is not
 * visible in the client window), this method "pans" around the page to capture
 * all of the information.
 *
 * @param hwndBrowser   the HWND of the browser containing the document
 * @param pDocument     the document being rendered
 * @param outputPath    the path to save the file as
 * @param scrollWidth   total width of content
 * @param scrollHeight  total height of content
 * @param clientWidth   the client width of the canvas element
 * @param clientHeight  the client height of the canvas element
 * @param clientLeft    the width of the canvas element's border
 * @param clientTop     the height of the canvas element's border
 */
STDMETHODIMP CCoSnapsie::panAndScan(HWND hwndBrowser, IUnknown* pDocument, BSTR outputPath,
    long scrollWidth, long scrollHeight,
    long clientWidth, long clientHeight,
    long clientLeft , long clientTop)
{
    ATLASSERT(scrollWidth  > 0);
    ATLASSERT(scrollHeight > 0);
    ATLASSERT(clientWidth  > 0);
    ATLASSERT(clientHeight > 0);

    HRESULT hr;
    CComQIPtr<IHTMLDocument2> spDocument;
    CComPtr<IHTMLWindow2>     spWindow;
    CComQIPtr<IViewObject2>   spViewObject;

    spDocument = pDocument;
    if (spDocument == NULL)
        return E_FAIL;

    hr = spDocument->get_parentWindow(&spWindow);
    if (FAILED(hr))
        return E_FAIL;

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
    bih.biWidth     = scrollWidth  + (2 * clientLeft);
    bih.biHeight    = scrollHeight + (2 * clientTop);
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
    image.Create(scrollWidth, scrollHeight, 24);
    CImageDC imageDC(image);

    // main drawing loops (pan and scan)

    long xEnd;
    long yEnd;

    for (long yStart = 0; yStart < scrollHeight; yStart += clientHeight) {
        yEnd = yStart + clientHeight;
        if (yEnd > scrollHeight) {
            if (yStart != 0) {
                yStart = scrollHeight - clientHeight;
            }
            yEnd = scrollHeight;
        }

        for (long xStart = 0; xStart < scrollWidth; xStart += clientWidth) {
            xEnd = xStart + clientWidth;
            if (xEnd > scrollWidth) {
                if (xStart != 0) {
                    xStart = scrollWidth - clientWidth;
                }
                xEnd = scrollWidth;
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

            RECTL rcBounds = { 0, 0, clientWidth + (2 * clientLeft),
                clientHeight + (2 * clientTop) };

            hr = spViewObject->Draw(DVASPECT_CONTENT, -1, NULL, NULL, hdcInput,
                hdcOutput, &rcBounds, NULL, NULL, 0);
            if (FAILED(hr)) {
                Error("Draw() failed during pan and scan");
                goto loop_exit;
            }

            // copy to correct region of output device

            ::BitBlt(imageDC, xStart, yStart, xEnd - xStart, yEnd - yStart,
                hdcOutput, clientLeft, clientTop, SRCCOPY);
        }
    }
    loop_exit:

    // save the image

    if (SUCCEEDED(hr)) {
        image.Save(CW2T(outputPath));
    }

    // clean up

    ::ReleaseDC(hwndBrowser, hdcInput);
    DeleteDC(hdcOutput);
    DeleteObject(hBitmap);

    return hr;
}

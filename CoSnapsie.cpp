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
STDMETHODIMP CCoSnapsie::saveSnapshot(BSTR outputPath)
{
    HRESULT hr;
    CComPtr<IOleClientSite>     spClientSite;
    CComQIPtr<IServiceProvider> spISP;
    CComPtr<IWebBrowser2>       spBrowser;
    CComPtr<IDispatch>          spDispatch; 
    CComQIPtr<IHTMLDocument2>   spDocument;
    CComPtr<IHTMLElement>       spBody;
    CComQIPtr<IHTMLElement2>    spBody2;
    CComQIPtr<IHTMLBodyElement> spBodyElement;
    CComPtr<IHTMLStyle>         spStyle;
    CComQIPtr<IHTMLDocument3>   spDocument3;
    CComPtr<IHTMLElement>       spElement;
    CComQIPtr<IHTMLElement2>    spElement2;

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

    hr = spBrowser->get_Document(&spDispatch);
    if (FAILED(hr))
        return E_FAIL;

    spDocument = spDispatch;
    if (spDocument == NULL)
        return E_FAIL;

    hr = spDocument->get_body(&spBody);
    if (spBody == NULL)
        return E_FAIL;

    hr = spBody->get_style(&spStyle);
    if (FAILED(hr))
        return E_FAIL;

    CComBSTR borderStyle;
    spStyle->get_borderStyle(&borderStyle);
    spStyle->put_borderStyle(CComBSTR("none"));    // hide 3D border

    spBodyElement = spBody;
    if (spBodyElement == NULL)
        return E_FAIL;

    CComBSTR scroll;
    spBodyElement->get_scroll(&scroll);
    spBodyElement->put_scroll(CComBSTR("no"));     // hide scrollbars

    spBody2 = spBody;
    if (spBody2 == NULL)
        return E_FAIL;

    long scrollWidth;
    long scrollHeight;
    spBody2->get_scrollWidth(&scrollWidth);
    spBody2->get_scrollHeight(&scrollHeight);

    spDocument3 = spDispatch;
    if (spDocument3 == NULL)
        return E_FAIL;

    hr = spDocument3->get_documentElement(&spElement);
    if (FAILED(hr))
        return E_FAIL;

    spElement2 = spElement;
    if (spElement2 == NULL)
        return E_FAIL;

    long clientWidth;
    long clientHeight;
    spBody2->get_clientWidth(&clientWidth);
    spBody2->get_clientHeight(&clientHeight);
    
    // record the current scroll position, and restore it afterwards. Also
    // restore the border style and scrollbars, if necessary.

    long scrollLeft = 0;
    long scrollTop = 0;
    spBody2->get_scrollLeft(&scrollLeft);
    spBody2->get_scrollTop(&scrollTop);

    hr = panAndScan(spBrowser, outputPath,
                    scrollWidth, scrollHeight,
                    clientWidth, clientHeight);

    spBody2->put_scrollLeft(scrollLeft);
    spBody2->put_scrollTop(scrollTop);

    spStyle->put_borderStyle(borderStyle);
    spBodyElement->put_scroll(scroll);

    return hr;
}

/**
 * This method copies the rendered HTML content of a browser object to a PNG
 * output file, using the outputPath parameter. If the client dimensions are
 * smaller than the scroll dimensions (i.e. some part of the whole image is not
 * visible in the client window), this method "pans" around the page to capture
 * all of the information.
 *
 * @param pBrowser      the browser containing the HTML content to copy
 * @param outputPath    the path to save the file as
 * @param scrollWidth   total width of content
 * @param scrollHeight  total height of content
 * @param clientWidth   current width of viewable content
 * @param clientHeight  current height of viewable content
 */
STDMETHODIMP CCoSnapsie::panAndScan(void* pBrowser, BSTR outputPath,
    long scrollWidth, long scrollHeight, long clientWidth, long clientHeight)
{
    ATLASSERT(scrollWidth  > 0);
    ATLASSERT(scrollHeight > 0);
    ATLASSERT(clientWidth  > 0);
    ATLASSERT(clientHeight > 0);

    HRESULT hr;
    HWND hwndBrowser;
    CComQIPtr<IWebBrowser2>   spBrowser;
    CComPtr<IDispatch>        spDispatch;
    CComQIPtr<IHTMLDocument2> spDocument;
    CComPtr<IHTMLWindow2>     spWindow;
    CComQIPtr<IViewObject2>   spViewObject;

    spBrowser = (IWebBrowser2*)pBrowser;
    if (spBrowser == NULL)
        return E_FAIL;

    hr = spBrowser->get_HWND((long*)&hwndBrowser);
    if (FAILED(hr))
        return E_FAIL;

    hr = spBrowser->get_Document(&spDispatch);
    if (FAILED(hr))
        return E_FAIL;

    spDocument = spDispatch;
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
    bih.biWidth     = scrollWidth;
    bih.biHeight    = scrollHeight;
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

    long count = 0;
    for (long yStart = 0; yStart < scrollHeight; yStart += clientHeight) {
        yEnd = yStart + clientHeight;
        if (yEnd > scrollHeight) {
            yStart -= yEnd - scrollHeight;
            yEnd = scrollHeight;
        }

        for (long xStart = 0; xStart < scrollWidth; xStart += clientWidth) {
            xEnd = xStart + clientWidth;
            if (xEnd > scrollWidth) {
                xStart -= xEnd - scrollWidth;
                xEnd = scrollWidth;
            }

            // scroll the window, and draw the results

            RECTL rcBounds = { 0, 0, clientWidth, clientHeight };

            hr = spWindow->scroll(xStart, yStart);
            if (FAILED(hr)) {
                break;
            }

            hr = spViewObject->Draw(DVASPECT_CONTENT, -1, NULL, NULL, hdcInput,
                hdcOutput, &rcBounds, NULL, NULL, 0);
            if (FAILED(hr))
                break;

            // copy to correct region of output device
            ::BitBlt(imageDC, xStart, yStart, xEnd - xStart, yEnd - yStart, hdcOutput, 0, 0, SRCCOPY);
        }
        if (FAILED(hr)) {
            Error("Failed during pan and scan");
            break;
        }
    }
    
    // save the image

    if (SUCCEEDED(hr)) {
        image.Save(CW2T(outputPath));
    }

    // clean up

    ReleaseDC(hwndBrowser, hdcInput);
    DeleteDC(hdcOutput);
    DeleteObject(hBitmap);

    return hr;
}

// CoSnapsie.h : Declaration of the CCoSnapsie

#pragma once
#include "resource.h"       // main symbols

#include "Snapsie.h"


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif



// CCoSnapsie

class ATL_NO_VTABLE CCoSnapsie :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CCoSnapsie, &CLSID_CoSnapsie>,
	public ISupportErrorInfo,
	public IObjectWithSiteImpl<CCoSnapsie>,
	public IDispatchImpl<ISnapsie, &IID_ISnapsie, &LIBID_SnapsieLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CCoSnapsie()
	{
	}


DECLARE_REGISTRY_RESOURCEID(IDR_COSNAPSIE)


BEGIN_COM_MAP(CCoSnapsie)
	COM_INTERFACE_ENTRY(ISnapsie)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IObjectWithSite)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:

    STDMETHODIMP saveSnapshot(
        BSTR outputFile,
        BSTR frameId,
        LONG drawableScrollWidth,
        LONG drawableScrollHeight,
        LONG drawableClientWidth,
        LONG drawableClientHeight,
        LONG drawableClientLeft,
        LONG drawableClientTop,
        LONG frameBCRLeft,
        LONG frameBCRTop);
};

OBJECT_ENTRY_AUTO(__uuidof(CoSnapsie), CCoSnapsie)

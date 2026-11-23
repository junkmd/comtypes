/*
	This code is based on example code to the book:
		Inside COM
		by Dale E. Rogerson
		Microsoft Press 1997
		ISBN 1-57231-349-8
*/

//
// CoComtypesDispDateParamTest.cpp - Component
//
#include <objbase.h>
#include <string.h>
#include <iostream>
#include <sstream>

#include "Iface.h"
#include "Util.h"
#include "CUnknown.h"
#include "CFactory.h" // Needed for module handle
#include "CoComtypesDispDateParamTest.h"

// We need to put this declaration here because we explicitely expose a dispinterface
// in parallel to the dual interface but dispinterfaces don't appear in the
// MIDL-generated header file.
EXTERN_C const IID DIID_IDispDateParamTest;

// The fixed DATE value used by InitDate and VerifyDate.
// 45293.5 is 2024-01-02 12:00:00 in OLE Automation date format
// (days since 1899-12-30, where the fractional part is the time of day).
static const DATE FIXED_DATE = 45293.5;

static inline void trace(const char* msg)
	{ Util::Trace("CoComtypesDispDateParamTest", msg, S_OK) ;}
static inline void trace(const char* msg, HRESULT hr)
	{ Util::Trace("CoComtypesDispDateParamTest", msg, hr) ;}

///////////////////////////////////////////////////////////
//
// Interface IDualDateParamTest - Implementation
//

HRESULT __stdcall CDateParamTest::EchoDate(DATE value, DATE* result)
{
	// Return the received DATE unchanged so that the client can verify
	// that the value survived the round trip through COM marshaling.
	*result = value ;
	return S_OK ;
}

HRESULT __stdcall CDateParamTest::InitDate(DATE* value)
{
	// Overwrite the received DATE with the fixed value.
	// Because the parameter is [in, out], the client observes the change.
	*value = FIXED_DATE ;
	return S_OK ;
}

HRESULT __stdcall CDateParamTest::VerifyDate(DATE value, VARIANT_BOOL* result)
{
	// Report whether the received DATE equals the fixed value.
	if (value == FIXED_DATE)
	{
		*result = VARIANT_TRUE ;
	}
	else
	{
		*result = VARIANT_FALSE ;
	}
	return S_OK ;
}

//
// Constructor
//
CDateParamTest::CDateParamTest(IUnknown* pUnknownOuter)
: CUnknown(pUnknownOuter), 
  m_pITypeInfo(NULL)
{
	// Empty
}

//
// Destructor
//
CDateParamTest::~CDateParamTest()
{
	if (m_pITypeInfo != NULL)
	{
		m_pITypeInfo->Release() ;
	}

	trace("Destroy self.") ;
}

//
// NondelegatingQueryInterface implementation
//
HRESULT __stdcall CDateParamTest::NondelegatingQueryInterface(const IID& iid,
                                                             void** ppv)
{ 	
	if (iid == IID_IDualDateParamTest)
	{
		return FinishQI(static_cast<IDualDateParamTest*>(this), ppv) ;
	}
	else 	if (iid == DIID_IDispDateParamTest)
	{
		trace("Queried for IDispDateParamTest.") ;
		return FinishQI(static_cast<IDispatch*>(this), ppv) ;
	}
	else 	if (iid == IID_IDispatch)
	{
		trace("Queried for IDispatch.") ;
		return FinishQI(static_cast<IDispatch*>(this), ppv) ;
	}
	else
	{
		return CUnknown::NondelegatingQueryInterface(iid, ppv) ;
	}
}

//
// Load and register the type library.
//
HRESULT CDateParamTest::Init()
{
	HRESULT hr ;

	// Load TypeInfo on demand if we haven't already loaded it.
	if (m_pITypeInfo == NULL)
	{
		ITypeLib* pITypeLib = NULL ;
		hr = ::LoadRegTypeLib(LIBID_ComtypesCppTestSrvLib, 
		                      1, 0, // Major/Minor version numbers
		                      0x00, 
		                      &pITypeLib) ;
		if (FAILED(hr)) 
		{
			trace("LoadRegTypeLib Failed.", hr) ;
			return hr ;   
		}

		// Get type information for the interface of the object.
		hr = pITypeLib->GetTypeInfoOfGuid(IID_IDualDateParamTest,
		                                  &m_pITypeInfo) ;
		pITypeLib->Release() ;
		if (FAILED(hr))  
		{ 
			trace("GetTypeInfoOfGuid failed.", hr) ;
			return hr ;
		}   
	}
	return S_OK ;
}

///////////////////////////////////////////////////////////
//
// Creation function used by CFactory
//
HRESULT CDateParamTest::CreateInstance(IUnknown* pUnknownOuter,
                                       CUnknown** ppNewComponent )
{
	if (pUnknownOuter != NULL)
	{
		// Don't allow aggregation (just for the heck of it).
		return CLASS_E_NOAGGREGATION ;
	}

	*ppNewComponent = new CDateParamTest(pUnknownOuter) ;
	return S_OK ;
}

///////////////////////////////////////////////////////////
//
// IDispatch implementation
//
HRESULT __stdcall CDateParamTest::GetTypeInfoCount(UINT* pCountTypeInfo)
{
	trace("GetTypeInfoCount call succeeded.") ;
	*pCountTypeInfo = 1 ;
	return S_OK ;
}

HRESULT __stdcall CDateParamTest::GetTypeInfo(
	UINT iTypeInfo,
	LCID,          // This object does not support localization.
	ITypeInfo** ppITypeInfo)
{    
	*ppITypeInfo = NULL ;

	if(iTypeInfo != 0)
	{
		trace("GetTypeInfo call failed -- bad iTypeInfo index.") ;
		return DISP_E_BADINDEX ; 
	}

	trace("GetTypeInfo call succeeded.") ;

	// Call AddRef and return the pointer.
	m_pITypeInfo->AddRef() ; 
	*ppITypeInfo = m_pITypeInfo ;
	return S_OK ;
}

HRESULT __stdcall CDateParamTest::GetIDsOfNames(
	const IID& iid,
	OLECHAR** arrayNames,
	UINT countNames,
	LCID,          // Localization is not supported.
	DISPID* arrayDispIDs)
{
	if (iid != IID_NULL)
	{
		trace("GetIDsOfNames call failed -- bad IID.") ;
		return DISP_E_UNKNOWNINTERFACE ;
	}

	trace("GetIDsOfNames call succeeded.") ;
	HRESULT hr = m_pITypeInfo->GetIDsOfNames(arrayNames,
	                                         countNames,
	                                         arrayDispIDs) ;
	return hr ;
}

HRESULT __stdcall CDateParamTest::Invoke(
      DISPID dispidMember,
      const IID& iid,
      LCID,          // Localization is not supported.
      WORD wFlags,
      DISPPARAMS* pDispParams,
      VARIANT* pvarResult,
      EXCEPINFO* pExcepInfo,
      UINT* pArgErr)
{        
	if (iid != IID_NULL)
	{
		trace("Invoke call failed -- bad IID.") ;
		return DISP_E_UNKNOWNINTERFACE ;
	}

	::SetErrorInfo(0, NULL) ;

	trace("Invoke call succeeded.") ;
	HRESULT hr = m_pITypeInfo->Invoke(
		static_cast<IDispatch*>(this),
		dispidMember, wFlags, pDispParams,
		pvarResult, pExcepInfo, pArgErr) ; 
	return hr ;
}

/*
	This code is based on example code to the book:
		Inside COM
		by Dale E. Rogerson
		Microsoft Press 1997
		ISBN 1-57231-349-8
*/

//
// CoComtypesNamedPropertyPutTest.cpp - Component
//
#include <objbase.h>
#include <string.h>
#include <iostream>
#include <sstream>

#include "Iface.h"
#include "Util.h"
#include "CUnknown.h"
#include "CFactory.h" // Needed for module handle
#include "CoComtypesNamedPropertyPutTest.h"

// We need to put this declaration here because we explicitly expose a dispinterface
// in parallel to the dual interface but dispinterfaces don't appear in the
// MIDL-generated header file.
EXTERN_C const IID DIID_IDispNamedPropertyPutTest;

static inline void trace(const char* msg)
	{ Util::Trace("CoComtypesNamedPropertyPutTest", msg, S_OK) ;}
static inline void trace(const char* msg, HRESULT hr)
	{ Util::Trace("CoComtypesNamedPropertyPutTest", msg, hr) ;}

///////////////////////////////////////////////////////////
//
// Constructor
//
CC::CC(IUnknown* pUnknownOuter)
: CUnknown(pUnknownOuter), 
  m_pITypeInfo(NULL)
{
	// Initialize the arrays with zeros
	memset(m_values, 0, sizeof(m_values));
}

//
// Destructor
//
CC::~CC()
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
HRESULT __stdcall CC::NondelegatingQueryInterface(const IID& iid,
                                                  void** ppv)
{ 	
	if (iid == IID_IDualNamedPropertyPutTest)
	{
		return FinishQI(static_cast<IDualNamedPropertyPutTest*>(this), ppv) ;
	}
	else if (iid == DIID_IDispNamedPropertyPutTest)
	{
		trace("Queried for IDispNamedPropertyPutTest.") ;
		return FinishQI(static_cast<IDispatch*>(this), ppv) ;
	}
	else if (iid == IID_IDispatch)
	{
		trace("Queried for IDispatch.") ;
		return FinishQI(static_cast<IDispatch*>(this), ppv) ;
	}
	else
	{
		return CUnknown::NondelegatingQueryInterface(iid, ppv) ;
	}
}

///////////////////////////////////////////////////////////
//
// Creation function used by CFactory
//
HRESULT CC::CreateInstance(IUnknown* pUnknownOuter,
                           CUnknown** ppNewComponent ) 
{
	if (pUnknownOuter != NULL)
	{
		// Don't allow aggregation (just for the heck of it).
		return CLASS_E_NOAGGREGATION ;
	}

	*ppNewComponent = new CC(pUnknownOuter) ;
	return S_OK ;
}

///////////////////////////////////////////////////////////
//
// Load and register the type library.
//
HRESULT CC::Init()
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
		hr = pITypeLib->GetTypeInfoOfGuid(IID_IDualNamedPropertyPutTest,
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
// IDispatch implementation
//
HRESULT __stdcall CC::GetTypeInfoCount(UINT* pCountTypeInfo)
{
	trace("GetTypeInfoCount call succeeded.") ;
	*pCountTypeInfo = 1 ;
	return S_OK ;
}

HRESULT __stdcall CC::GetTypeInfo(
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

	// Initialize the type info if it hasn't been initialized yet
	HRESULT hr = Init();
	if (FAILED(hr))
	{
		return hr;
	}

	trace("GetTypeInfo call succeeded.") ;

	// Call AddRef and return the pointer.
	m_pITypeInfo->AddRef() ; 
	*ppITypeInfo = m_pITypeInfo ;
	return S_OK ;
}

HRESULT __stdcall CC::GetIDsOfNames(  
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

	// Initialize the type info if it hasn't been initialized yet
	HRESULT hr = Init();
	if (FAILED(hr))
	{
		return hr;
	}

	trace("GetIDsOfNames call succeeded.") ;
	hr = m_pITypeInfo->GetIDsOfNames(arrayNames,
	                                 countNames,
	                                 arrayDispIDs) ;
	return hr ;
}

HRESULT __stdcall CC::Invoke(   
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

	// Initialize the type info if it hasn't been initialized yet
	HRESULT hr = Init();
	if (FAILED(hr))
	{
		return hr;
	}

	::SetErrorInfo(0, NULL) ;

	trace("Invoke call succeeded.") ;
	hr = m_pITypeInfo->Invoke(
		static_cast<IDispatch*>(this),
		dispidMember, wFlags, pDispParams,
		pvarResult, pExcepInfo, pArgErr) ; 
	return hr ;
}

///////////////////////////////////////////////////////////
//
// Interface IDualNamedPropertyPutTest - Implementation
//

// Helper function to check if a VARIANT is empty or missing
static bool IsEmptyOrMissing(const VARIANT& var)
{
	return (var.vt == VT_EMPTY || var.vt == VT_ERROR || var.vt == VT_NULL);
}

// Helper function to create a SAFEARRAY of longs from a row
static SAFEARRAY* CreateSafeArrayFromRow(const long row[3])
{
	SAFEARRAYBOUND bounds[1];
	bounds[0].lLbound = 0;
	bounds[0].cElements = 3;
	
	SAFEARRAY* psa = SafeArrayCreate(VT_I4, 1, bounds);
	if (psa == NULL)
	{
		return NULL;
	}
	
	long* pData;
	HRESULT hr = SafeArrayAccessData(psa, (void**)&pData);
	if (FAILED(hr))
	{
		SafeArrayDestroy(psa);
		return NULL;
	}
	
	memcpy(pData, row, 3 * sizeof(long));
	SafeArrayUnaccessData(psa);
	
	return psa;
}

// Helper function to create a SAFEARRAY of SAFEARRAYs from the 2D array
static SAFEARRAY* CreateSafeArrayFromArray(const long values[2][3])
{
	// Create a SAFEARRAY for the outer array (2 rows)
	SAFEARRAYBOUND bounds[1];
	bounds[0].lLbound = 0;
	bounds[0].cElements = 2;
	
	// Create a SAFEARRAY of VARIANTs to hold the rows
	SAFEARRAY* psa = SafeArrayCreate(VT_VARIANT, 1, bounds);
	if (psa == NULL)
	{
		return NULL;
	}
	
	// Access the data
	VARIANT* pData;
	HRESULT hr = SafeArrayAccessData(psa, (void**)&pData);
	if (FAILED(hr))
	{
		SafeArrayDestroy(psa);
		return NULL;
	}
	
	// For each row, create a SAFEARRAY and store it in the outer array
	for (int i = 0; i < 2; i++)
	{
		// Create a SAFEARRAY for the inner array (3 columns)
		SAFEARRAYBOUND innerBounds[1];
		innerBounds[0].lLbound = 0;
		innerBounds[0].cElements = 3;
		
		// Create a SAFEARRAY of longs for the row
		SAFEARRAY* psaRow = SafeArrayCreate(VT_I4, 1, innerBounds);
		if (psaRow == NULL)
		{
			// Clean up
			for (int j = 0; j < i; j++)
			{
				VariantClear(&pData[j]);
			}
			SafeArrayUnaccessData(psa);
			SafeArrayDestroy(psa);
			return NULL;
		}
		
		// Access the row data
		long* pRowData;
		hr = SafeArrayAccessData(psaRow, (void**)&pRowData);
		if (FAILED(hr))
		{
			// Clean up
			SafeArrayDestroy(psaRow);
			for (int j = 0; j < i; j++)
			{
				VariantClear(&pData[j]);
			}
			SafeArrayUnaccessData(psa);
			SafeArrayDestroy(psa);
			return NULL;
		}
		
		// Copy the values
		for (int j = 0; j < 3; j++)
		{
			pRowData[j] = values[i][j];
		}
		
		// Unlock the row data
		SafeArrayUnaccessData(psaRow);
		
		// Store the row in the outer array
		VariantInit(&pData[i]);
		pData[i].vt = VT_ARRAY | VT_I4;
		pData[i].parray = psaRow;
	}
	
	// Unlock the outer array
	SafeArrayUnaccessData(psa);
	
	return psa;
}

// Helper function to extract values from a SAFEARRAY and store them in a row
static HRESULT ExtractValuesFromSafeArray(SAFEARRAY* psa, long row[3])
{
	if (psa == NULL)
	{
		return E_INVALIDARG;
	}
	
	// Check dimensions
	if (SafeArrayGetDim(psa) != 1)
	{
		return E_INVALIDARG;
	}
	
	LONG lLBound, lUBound;
	HRESULT hr = SafeArrayGetLBound(psa, 1, &lLBound);
	if (FAILED(hr))
	{
		return hr;
	}
	
	hr = SafeArrayGetUBound(psa, 1, &lUBound);
	if (FAILED(hr))
	{
		return hr;
	}
	
	if (lUBound - lLBound + 1 != 3)
	{
		return E_INVALIDARG;
	}
	
	// Access the data
	long* pData;
	hr = SafeArrayAccessData(psa, (void**)&pData);
	if (FAILED(hr))
	{
		return hr;
	}
	
	memcpy(row, pData, 3 * sizeof(long));
	SafeArrayUnaccessData(psa);
	
	return S_OK;
}

// Helper function to extract values from a nested SAFEARRAY and store them in the 2D array
static HRESULT ExtractValuesFromNestedSafeArray(SAFEARRAY* psa, long values[2][3])
{
	if (psa == NULL)
	{
		return E_INVALIDARG;
	}
	
	// Check dimensions
	if (SafeArrayGetDim(psa) != 1)
	{
		return E_INVALIDARG;
	}
	
	LONG lLBound, lUBound;
	HRESULT hr = SafeArrayGetLBound(psa, 1, &lLBound);
	if (FAILED(hr))
	{
		return hr;
	}
	
	hr = SafeArrayGetUBound(psa, 1, &lUBound);
	if (FAILED(hr))
	{
		return hr;
	}
	
	if (lUBound - lLBound + 1 != 2)
	{
		return E_INVALIDARG;
	}
	
	// Access the data
	VARIANT* pData;
	hr = SafeArrayAccessData(psa, (void**)&pData);
	if (FAILED(hr))
	{
		return hr;
	}
	
	for (int i = 0; i < 2; i++)
	{
		// Handle different types of arrays that might be passed from Python
		if (pData[i].vt == (VT_ARRAY | VT_I4))
		{
			hr = ExtractValuesFromSafeArray(pData[i].parray, values[i]);
			if (FAILED(hr))
			{
				SafeArrayUnaccessData(psa);
				return hr;
			}
		}
		else if (pData[i].vt == (VT_ARRAY | VT_VARIANT))
		{
			// Handle array of variants
			SAFEARRAY* psaRow = pData[i].parray;
			LONG rowLBound, rowUBound;
			hr = SafeArrayGetLBound(psaRow, 1, &rowLBound);
			if (FAILED(hr))
			{
				SafeArrayUnaccessData(psa);
				return hr;
			}
			
			hr = SafeArrayGetUBound(psaRow, 1, &rowUBound);
			if (FAILED(hr))
			{
				SafeArrayUnaccessData(psa);
				return hr;
			}
			
			if (rowUBound - rowLBound + 1 != 3)
			{
				SafeArrayUnaccessData(psa);
				return E_INVALIDARG;
			}
			
			VARIANT* pRowData;
			hr = SafeArrayAccessData(psaRow, (void**)&pRowData);
			if (FAILED(hr))
			{
				SafeArrayUnaccessData(psa);
				return hr;
			}
			
			for (int j = 0; j < 3; j++)
			{
				// Convert each variant to a long
				VARIANT varValue;
				VariantInit(&varValue);
				hr = VariantChangeType(&varValue, &pRowData[j], 0, VT_I4);
				if (FAILED(hr))
				{
					SafeArrayUnaccessData(psaRow);
					SafeArrayUnaccessData(psa);
					return hr;
				}
				
				values[i][j] = varValue.lVal;
			}
			
			SafeArrayUnaccessData(psaRow);
		}
		else
		{
			SafeArrayUnaccessData(psa);
			return E_INVALIDARG;
		}
	}
	
	SafeArrayUnaccessData(psa);
	
	return S_OK;
}

HRESULT __stdcall CC::get_Value(VARIANT Index1, VARIANT Index2, VARIANT* pResult)
{
	std::ostringstream sout;
	sout << "get_Value called with Index1.vt=" << Index1.vt << ", Index2.vt=" << Index2.vt << std::ends;
	trace(sout.str().c_str());
	
	VariantInit(pResult);
	
	// Case 1: Both indices are provided - return a single value
	if (Index1.vt == VT_I4 && Index2.vt == VT_I4)
	{
		long row = Index1.lVal;
		long col = Index2.lVal;
		
		// Check bounds
		if (row < 0 || row > 1 || col < 0 || col > 2)
		{
			return E_INVALIDARG;
		}
		
		pResult->vt = VT_I4;
		pResult->lVal = m_values[row][col];
		return S_OK;
	}
	// Case 2: Only one index is provided - return a row
	else if (Index1.vt == VT_I4 && IsEmptyOrMissing(Index2))
	{
		long row = Index1.lVal;
		
		// Check bounds
		if (row < 0 || row > 1)
		{
			return E_INVALIDARG;
		}
		
		SAFEARRAY* psa = CreateSafeArrayFromRow(m_values[row]);
		if (psa == NULL)
		{
			return E_OUTOFMEMORY;
		}
		
		pResult->vt = VT_ARRAY | VT_I4;
		pResult->parray = psa;
		return S_OK;
	}
	// Case 3: Empty tuple or slice - return the entire array
	else if (IsEmptyOrMissing(Index1) && IsEmptyOrMissing(Index2))
	{
		// Create a SAFEARRAY for the outer array (2 rows)
		SAFEARRAYBOUND bounds[1];
		bounds[0].lLbound = 0;
		bounds[0].cElements = 2;
		
		// Create a SAFEARRAY of VARIANTs to hold the rows
		SAFEARRAY* psa = SafeArrayCreate(VT_VARIANT, 1, bounds);
		if (psa == NULL)
		{
			return E_OUTOFMEMORY;
		}
		
		// Access the data
		VARIANT* pData;
		HRESULT hr = SafeArrayAccessData(psa, (void**)&pData);
		if (FAILED(hr))
		{
			SafeArrayDestroy(psa);
			return hr;
		}
		
		// For each row, create a SAFEARRAY and store it in the outer array
		for (int i = 0; i < 2; i++)
		{
			// Create a SAFEARRAY for the inner array (3 columns)
			SAFEARRAYBOUND innerBounds[1];
			innerBounds[0].lLbound = 0;
			innerBounds[0].cElements = 3;
			
			// Create a SAFEARRAY of longs for the row
			SAFEARRAY* psaRow = SafeArrayCreate(VT_I4, 1, innerBounds);
			if (psaRow == NULL)
			{
				// Clean up
				for (int j = 0; j < i; j++)
				{
					VariantClear(&pData[j]);
				}
				SafeArrayUnaccessData(psa);
				SafeArrayDestroy(psa);
				return E_OUTOFMEMORY;
			}
			
			// Access the row data
			long* pRowData;
			hr = SafeArrayAccessData(psaRow, (void**)&pRowData);
			if (FAILED(hr))
			{
				// Clean up
				SafeArrayDestroy(psaRow);
				for (int j = 0; j < i; j++)
				{
					VariantClear(&pData[j]);
				}
				SafeArrayUnaccessData(psa);
				SafeArrayDestroy(psa);
				return hr;
			}
			
			// Copy the values
			for (int j = 0; j < 3; j++)
			{
				pRowData[j] = m_values[i][j];
			}
			
			// Unlock the row data
			SafeArrayUnaccessData(psaRow);
			
			// Store the row in the outer array
			VariantInit(&pData[i]);
			pData[i].vt = VT_ARRAY | VT_I4;
			pData[i].parray = psaRow;
		}
		
		// Unlock the outer array
		SafeArrayUnaccessData(psa);
		
		pResult->vt = VT_ARRAY | VT_VARIANT;
		pResult->parray = psa;
		return S_OK;
	}
	else
	{
		return E_INVALIDARG;
	}
}

HRESULT __stdcall CC::put_Value(VARIANT Index1, VARIANT Index2, VARIANT newValue)
{
	std::ostringstream sout;
	sout << "put_Value called with Index1.vt=" << Index1.vt << ", Index2.vt=" << Index2.vt << ", newValue.vt=" << newValue.vt << std::ends;
	trace(sout.str().c_str());
	
	// Case 1: Both indices are provided - set a single value
	if (Index1.vt == VT_I4 && Index2.vt == VT_I4)
	{
		long row = Index1.lVal;
		long col = Index2.lVal;
		
		// Check bounds
		if (row < 0 || row > 1 || col < 0 || col > 2)
		{
			return E_INVALIDARG;
		}
		
		// Convert the value to a long
		VARIANT varValue;
		VariantInit(&varValue);
		HRESULT hr = VariantChangeType(&varValue, &newValue, 0, VT_I4);
		if (FAILED(hr))
		{
			return hr;
		}
		
		m_values[row][col] = varValue.lVal;
		return S_OK;
	}
	// Case 2: Only one index is provided - set a row
	else if (Index1.vt == VT_I4 && IsEmptyOrMissing(Index2))
	{
		long row = Index1.lVal;
		
		// Check bounds
		if (row < 0 || row > 1)
		{
			return E_INVALIDARG;
		}
		
		// Handle different types of input
		if (newValue.vt == (VT_ARRAY | VT_I4))
		{
			// Direct array of integers
			SAFEARRAY* psa = newValue.parray;
			
			// Check dimensions
			if (SafeArrayGetDim(psa) != 1)
			{
				return E_INVALIDARG;
			}
			
			LONG lLBound, lUBound;
			HRESULT hr = SafeArrayGetLBound(psa, 1, &lLBound);
			if (FAILED(hr))
			{
				return hr;
			}
			
			hr = SafeArrayGetUBound(psa, 1, &lUBound);
			if (FAILED(hr))
			{
				return hr;
			}
			
			if (lUBound - lLBound + 1 != 3)
			{
				return E_INVALIDARG;
			}
			
			// Access the data
			long* pData;
			hr = SafeArrayAccessData(psa, (void**)&pData);
			if (FAILED(hr))
			{
				return hr;
			}
			
			// Copy the data
			for (int i = 0; i < 3; i++)
			{
				m_values[row][i] = pData[i];
			}
			
			SafeArrayUnaccessData(psa);
			return S_OK;
		}
		else if (newValue.vt == (VT_ARRAY | VT_VARIANT))
		{
			// Array of variants
			SAFEARRAY* psa = newValue.parray;
			
			// Check dimensions
			if (SafeArrayGetDim(psa) != 1)
			{
				return E_INVALIDARG;
			}
			
			LONG lLBound, lUBound;
			HRESULT hr = SafeArrayGetLBound(psa, 1, &lLBound);
			if (FAILED(hr))
			{
				return hr;
			}
			
			hr = SafeArrayGetUBound(psa, 1, &lUBound);
			if (FAILED(hr))
			{
				return hr;
			}
			
			if (lUBound - lLBound + 1 != 3)
			{
				return E_INVALIDARG;
			}
			
			// Access the data
			VARIANT* pData;
			hr = SafeArrayAccessData(psa, (void**)&pData);
			if (FAILED(hr))
			{
				return hr;
			}
			
			// Convert and copy the data
			for (int i = 0; i < 3; i++)
			{
				VARIANT varValue;
				VariantInit(&varValue);
				hr = VariantChangeType(&varValue, &pData[i], 0, VT_I4);
				if (FAILED(hr))
				{
					SafeArrayUnaccessData(psa);
					return hr;
				}
				
				m_values[row][i] = varValue.lVal;
			}
			
			SafeArrayUnaccessData(psa);
			return S_OK;
		}
		else
		{
			return E_INVALIDARG;
		}
	}
	// Case 3: Empty tuple or slice - set the entire array
	else if (IsEmptyOrMissing(Index1) && IsEmptyOrMissing(Index2))
	{
		// Handle different types of input
		if (newValue.vt == (VT_ARRAY | VT_VARIANT))
		{
			SAFEARRAY* psa = newValue.parray;
			
			// Check dimensions
			if (SafeArrayGetDim(psa) != 1)
			{
				return E_INVALIDARG;
			}
			
			LONG lLBound, lUBound;
			HRESULT hr = SafeArrayGetLBound(psa, 1, &lLBound);
			if (FAILED(hr))
			{
				return hr;
			}
			
			hr = SafeArrayGetUBound(psa, 1, &lUBound);
			if (FAILED(hr))
			{
				return hr;
			}
			
			if (lUBound - lLBound + 1 != 2)
			{
				return E_INVALIDARG;
			}
			
			// Access the data
			VARIANT* pData;
			hr = SafeArrayAccessData(psa, (void**)&pData);
			if (FAILED(hr))
			{
				return hr;
			}
			
			// Process each row
			for (int i = 0; i < 2; i++)
			{
				if (pData[i].vt == (VT_ARRAY | VT_I4))
				{
					// Direct array of integers
					SAFEARRAY* psaRow = pData[i].parray;
					
					// Check dimensions
					if (SafeArrayGetDim(psaRow) != 1)
					{
						SafeArrayUnaccessData(psa);
						return E_INVALIDARG;
					}
					
					LONG rowLBound, rowUBound;
					hr = SafeArrayGetLBound(psaRow, 1, &rowLBound);
					if (FAILED(hr))
					{
						SafeArrayUnaccessData(psa);
						return hr;
					}
					
					hr = SafeArrayGetUBound(psaRow, 1, &rowUBound);
					if (FAILED(hr))
					{
						SafeArrayUnaccessData(psa);
						return hr;
					}
					
					if (rowUBound - rowLBound + 1 != 3)
					{
						SafeArrayUnaccessData(psa);
						return E_INVALIDARG;
					}
					
					// Access the row data
					long* pRowData;
					hr = SafeArrayAccessData(psaRow, (void**)&pRowData);
					if (FAILED(hr))
					{
						SafeArrayUnaccessData(psa);
						return hr;
					}
					
					// Copy the data
					for (int j = 0; j < 3; j++)
					{
						m_values[i][j] = pRowData[j];
					}
					
					SafeArrayUnaccessData(psaRow);
				}
				else if (pData[i].vt == (VT_ARRAY | VT_VARIANT))
				{
					// Array of variants
					SAFEARRAY* psaRow = pData[i].parray;
					
					// Check dimensions
					if (SafeArrayGetDim(psaRow) != 1)
					{
						SafeArrayUnaccessData(psa);
						return E_INVALIDARG;
					}
					
					LONG rowLBound, rowUBound;
					hr = SafeArrayGetLBound(psaRow, 1, &rowLBound);
					if (FAILED(hr))
					{
						SafeArrayUnaccessData(psa);
						return hr;
					}
					
					hr = SafeArrayGetUBound(psaRow, 1, &rowUBound);
					if (FAILED(hr))
					{
						SafeArrayUnaccessData(psa);
						return hr;
					}
					
					if (rowUBound - rowLBound + 1 != 3)
					{
						SafeArrayUnaccessData(psa);
						return E_INVALIDARG;
					}
					
					// Access the row data
					VARIANT* pRowData;
					hr = SafeArrayAccessData(psaRow, (void**)&pRowData);
					if (FAILED(hr))
					{
						SafeArrayUnaccessData(psa);
						return hr;
					}
					
					// Convert and copy the data
					for (int j = 0; j < 3; j++)
					{
						VARIANT varValue;
						VariantInit(&varValue);
						hr = VariantChangeType(&varValue, &pRowData[j], 0, VT_I4);
						if (FAILED(hr))
						{
							SafeArrayUnaccessData(psaRow);
							SafeArrayUnaccessData(psa);
							return hr;
						}
						
						m_values[i][j] = varValue.lVal;
					}
					
					SafeArrayUnaccessData(psaRow);
				}
				else
				{
					SafeArrayUnaccessData(psa);
					return E_INVALIDARG;
				}
			}
			
			SafeArrayUnaccessData(psa);
			return S_OK;
		}
		else
		{
			return E_INVALIDARG;
		}
	}
	else
	{
		return E_INVALIDARG;
	}
}

/*
	This code is based on example code to the book:
		Inside COM
		by Dale E. Rogerson
		Microsoft Press 1997
		ISBN 1-57231-349-8
*/

//
// CoComtypesNamedPropertyTest.h - Component
//

#include "Iface.h"
#include "CUnknown.h" 

///////////////////////////////////////////////////////////
//
// Component CC
//
class CC : public CUnknown,
		   public IDualNamedPropertyPutTest
{
public:	
	// Creation
	static HRESULT CreateInstance(IUnknown* pUnknownOuter,
	                              CUnknown** ppNewComponent ) ;

private:
	// Declare the delegating IUnknown.
	DECLARE_IUNKNOWN

	// IUnknown
	virtual HRESULT __stdcall NondelegatingQueryInterface(const IID& iid,
	                                                      void** ppv) ;

	// IDispatch
	virtual HRESULT __stdcall GetTypeInfoCount(UINT* pCountTypeInfo) ;

	virtual HRESULT __stdcall GetTypeInfo(
		UINT iTypeInfo,
		LCID,              // Localization is not supported.
		ITypeInfo** ppITypeInfo) ;
	
	virtual HRESULT __stdcall GetIDsOfNames(
		const IID& iid,
		OLECHAR** arrayNames,
		UINT countNames,
		LCID,              // Localization is not supported.
		DISPID* arrayDispIDs) ;

	virtual HRESULT __stdcall Invoke(   
		DISPID dispidMember,
		const IID& iid,
		LCID,              // Localization is not supported.
		WORD wFlags,
		DISPPARAMS* pDispParams,
		VARIANT* pvarResult,
		EXCEPINFO* pExcepInfo,
		UINT* pArgErr) ;

	// Interface IDualNamedPropertyPutTest
	virtual HRESULT __stdcall get_Value(VARIANT Index1, VARIANT Index2, VARIANT* pResult) ;
	virtual HRESULT __stdcall put_Value(VARIANT Index1, VARIANT Index2, VARIANT newValue) ;

	// Initialization
 	virtual HRESULT Init() ;

	// Constructor
	CC(IUnknown* pUnknownOuter) ;

	// Destructor
	~CC() ;

	// Pointer to type information.
	ITypeInfo* m_pITypeInfo ;

	// 2D arrays to store values
	long m_values[2][3] ;
} ;
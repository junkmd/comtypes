// CoComtypesSafearrayEchoTest.h
// Test double for SAFEARRAY(VARIANT) marshaling tests.
// Provides methods to echo SAFEARRAY(VARIANT) and get VARTYPEs from SAFEARRAY(VARIANT).
#include "CUnknown.h"
#include "Iface.h"

class CSafearrayEchoTest : public CUnknown, public ISafearrayEchoTest
{
public:
	static HRESULT CreateInstance(IUnknown* pUnknownOuter, CUnknown** ppNewComponent);

private:
	DECLARE_IUNKNOWN
	virtual HRESULT __stdcall NondelegatingQueryInterface(const IID& iid, void** ppv);

	// Interface ISafearrayEchoTest
	virtual HRESULT __stdcall EchoSafeArrayVariant(SAFEARRAY* inArr, SAFEARRAY** outArr);
	virtual HRESULT __stdcall GetVtsFromSafeArrayVariant(SAFEARRAY* inArr, SAFEARRAY** outVts);

	CSafearrayEchoTest(IUnknown* pUnknownOuter) : CUnknown(pUnknownOuter) {}
};
// CoComtypesVariantEchoTest.h
// Test double for VARIANT round-trip and SAFEARRAY(VARIANT) marshaling tests.
#include "CUnknown.h"
#include "Iface.h"

class CVariantEchoTest : public CUnknown, public IVariantEchoTest
{
public:
	static HRESULT CreateInstance(IUnknown* pUnknownOuter, CUnknown** ppNewComponent);

private:
	DECLARE_IUNKNOWN
	virtual HRESULT __stdcall NondelegatingQueryInterface(const IID& iid, void** ppv);

	// Interface IVariantEchoTest
	virtual HRESULT __stdcall EchoVariant(VARIANT inVar, VARIANT* outVar);
	virtual HRESULT __stdcall GetVariantVt(VARIANT inVar, LONG* vt);
	virtual HRESULT __stdcall EchoSafeArrayVariant(SAFEARRAY* inArr, SAFEARRAY** outArr);
	virtual HRESULT __stdcall GetVtsFromSafeArrayVariant(SAFEARRAY* inArr, SAFEARRAY** outVts);

	CVariantEchoTest(IUnknown* pUnknownOuter) : CUnknown(pUnknownOuter) {}
};
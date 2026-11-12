// CoComtypesVariantEchoTest.h
// Test double for VARIANT round-trip tests.
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

	CVariantEchoTest(IUnknown* pUnknownOuter) : CUnknown(pUnknownOuter) {}
};
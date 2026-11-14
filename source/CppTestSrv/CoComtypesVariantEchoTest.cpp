// CoComtypesVariantEchoTest.cpp
// Test double for VARIANT round-trip tests.
// Provides methods to echo VARIANTs and observe their VT.

#include <objbase.h>
#include <string.h>
#include <iostream>
#include <sstream>
#include <vector>

#include "Iface.h"
#include "Util.h"
#include "CUnknown.h"
#include "CFactory.h"
#include "CoComtypesVariantEchoTest.h"

static inline void trace(const char* msg)
    { Util::Trace("CoComtypesVariantEchoTest", msg, S_OK); }
static inline void trace(const char* msg, HRESULT hr)
    { Util::Trace("CoComtypesVariantEchoTest", msg, hr); }

///////////////////////////////////////////////////////////
//
// Interface IVariantEchoTest - Implementation
//

// VARIANT echo: [in] VARIANT -> [out,retval] VARIANT
// Returns the exact same VARIANT received (vt and value preserved)
HRESULT __stdcall CVariantEchoTest::EchoVariant(VARIANT inVar, VARIANT* outVar)
{
    std::ostringstream sout;
    sout << "EchoVariant: received vt=0x" << std::hex << inVar.vt;
    trace(sout.str().c_str());

    if (outVar)
    {
        VariantInit(outVar);
        HRESULT hr = VariantCopy(outVar, &inVar);
        if (FAILED(hr))
            return hr;
    }

    return S_OK;
}

// GetVariantVt: [in] VARIANT -> [out,retval] LONG (vt)
// Observes: what VT the server sees when Python passes a VARIANT
HRESULT __stdcall CVariantEchoTest::GetVariantVt(VARIANT inVar, LONG* vt)
{
    std::ostringstream sout;
    sout << "GetVariantVt: received vt=0x" << std::hex << inVar.vt;
    trace(sout.str().c_str());

    if (vt)
        *vt = inVar.vt;

    return S_OK;
}

// Creation function used by CFactory
HRESULT CVariantEchoTest::CreateInstance(IUnknown* pUnknownOuter, CUnknown** ppNewComponent)
{
    if (pUnknownOuter != NULL)
    {
        return CLASS_E_NOAGGREGATION;
    }

    *ppNewComponent = new CVariantEchoTest(pUnknownOuter);
    return S_OK;
}

// NondelegatingQueryInterface
HRESULT __stdcall CVariantEchoTest::NondelegatingQueryInterface(const IID& iid, void** ppv)
{
    if (iid == IID_IVariantEchoTest)
    {
        return FinishQI(static_cast<IVariantEchoTest*>(this), ppv);
    }
    else
    {
        return CUnknown::NondelegatingQueryInterface(iid, ppv);
    }
}
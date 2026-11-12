// CoComtypesVariantEchoTest.cpp
// Test double for VARIANT round-trip and SAFEARRAY(VARIANT) marshaling tests.
// Provides methods to echo VARIANTs, observe their VT, and work with SAFEARRAY(VARIANT).

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

// SAFEARRAY(VARIANT) echo
HRESULT __stdcall CVariantEchoTest::EchoSafeArrayVariant(SAFEARRAY* inArr, SAFEARRAY** outArr)
{
    trace("EchoSafeArrayVariant: called");

    if (!inArr || !outArr)
        return E_POINTER;

    HRESULT hr = SafeArrayCopy(inArr, outArr);
    return hr;
}

// Get VARTYPE of each element in a SAFEARRAY(VARIANT)
HRESULT __stdcall CVariantEchoTest::GetVtsFromSafeArrayVariant(SAFEARRAY* inArr, SAFEARRAY** outVts)
{
    trace("GetVtsFromSafeArrayVariant: called");

    if (!inArr || !outVts)
        return E_POINTER;

    VARTYPE vt;
    HRESULT hr = SafeArrayGetVartype(inArr, &vt);
    if (FAILED(hr) || vt != VT_VARIANT)
        return DISP_E_TYPEMISMATCH;

    LONG lbound, ubound;
    SafeArrayGetLBound(inArr, 1, &lbound);
    SafeArrayGetUBound(inArr, 1, &ubound);
    LONG count = ubound - lbound + 1;

    // Create output SAFEARRAY of LONG (VARTYPE)
    SAFEARRAYBOUND bound = { (ULONG)count, lbound };
    *outVts = SafeArrayCreate(VT_I4, 1, &bound);
    if (!*outVts)
        return E_OUTOFMEMORY;

    // Lock input array and read VARIANTs
    VARIANT* pSrc = NULL;
    hr = SafeArrayAccessData(inArr, (void**)&pSrc);
    if (FAILED(hr))
        return hr;

    // Lock output array
    LONG* pDst = NULL;
    hr = SafeArrayAccessData(*outVts, (void**)&pDst);
    if (FAILED(hr))
    {
        SafeArrayUnaccessData(inArr);
        return hr;
    }

    for (LONG i = 0; i < count; i++)
    {
        pDst[i] = pSrc[i].vt;
    }

    SafeArrayUnaccessData(inArr);
    SafeArrayUnaccessData(*outVts);

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
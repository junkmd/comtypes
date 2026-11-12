// CoComtypesSafearrayEchoTest.cpp
// Test double for SAFEARRAY(VARIANT) marshaling tests.
// Provides methods to echo SAFEARRAY(VARIANT) and get VARTYPEs from SAFEARRAY(VARIANT).

#include <objbase.h>
#include <string.h>
#include <iostream>
#include <sstream>
#include <vector>

#include "Iface.h"
#include "Util.h"
#include "CUnknown.h"
#include "CFactory.h"
#include "CoComtypesSafearrayEchoTest.h"

static inline void trace(const char* msg)
    { Util::Trace("CoComtypesSafearrayEchoTest", msg, S_OK); }
static inline void trace(const char* msg, HRESULT hr)
    { Util::Trace("CoComtypesSafearrayEchoTest", msg, hr); }

///////////////////////////////////////////////////////////
//
// Interface ISafearrayEchoTest - Implementation
//

// SAFEARRAY(VARIANT) echo
HRESULT __stdcall CSafearrayEchoTest::EchoSafeArrayVariant(SAFEARRAY* inArr, SAFEARRAY** outArr)
{
    trace("EchoSafeArrayVariant: called");

    if (!inArr || !outArr)
        return E_POINTER;

    HRESULT hr = SafeArrayCopy(inArr, outArr);
    return hr;
}

// Get VARTYPE of each element in a SAFEARRAY(VARIANT)
HRESULT __stdcall CSafearrayEchoTest::GetVtsFromSafeArrayVariant(SAFEARRAY* inArr, SAFEARRAY** outVts)
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
HRESULT CSafearrayEchoTest::CreateInstance(IUnknown* pUnknownOuter, CUnknown** ppNewComponent)
{
    if (pUnknownOuter != NULL)
    {
        return CLASS_E_NOAGGREGATION;
    }

    *ppNewComponent = new CSafearrayEchoTest(pUnknownOuter);
    return S_OK;
}

// NondelegatingQueryInterface
HRESULT __stdcall CSafearrayEchoTest::NondelegatingQueryInterface(const IID& iid, void** ppv)
{
    if (iid == IID_ISafearrayEchoTest)
    {
        return FinishQI(static_cast<ISafearrayEchoTest*>(this), ppv);
    }
    else
    {
        return CUnknown::NondelegatingQueryInterface(iid, ppv);
    }
}
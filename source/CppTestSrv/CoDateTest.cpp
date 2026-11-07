// CoDateTest.cpp
// Test double for VT_DATE / DATE marshaling characterization tests.
// Each method is designed to observe how DATE/VT_DATE/VT_R8 values
// are marshaled across the Python <-> C++ language boundary.

#include <objbase.h>
#include <string.h>
#include <iostream>
#include <sstream>
#include <vector>

#include "Iface.h"
#include "Util.h"
#include "CUnknown.h"
#include "CFactory.h"
#include "CoDateTest.h"

// We need to put this declaration here because we explicitly expose a dispinterface
// in parallel to the dual interface but dispinterfaces don't appear in the
// MIDL-generated header file.
EXTERN_C const IID DIID_IDispDateTest;

static inline void trace(const char* msg)
    { Util::Trace("CoDateTest", msg, S_OK); }
static inline void trace(const char* msg, HRESULT hr)
    { Util::Trace("CoDateTest", msg, hr); }

///////////////////////////////////////////////////////////
//
// Interface IDualDateTest - Implementation
//

// 1. Typed DATE round-trip: [in] DATE -> [out,retval] DATE
// Observes: what value the server receives when Python passes datetime/float,
// and what Python receives when server returns the same value.
HRESULT __stdcall CDateTest::DateRoundTrip(DATE inDate, DATE* outDate)
{
    std::ostringstream sout;
    sout << "DateRoundTrip: received inDate=" << inDate;
    trace(sout.str().c_str());

    if (outDate)
        *outDate = inDate;

    return S_OK;
}

// 2. [out] DATE* - server creates a DATE value
HRESULT __stdcall CDateTest::DateOut(DATE* outDate)
{
    // Return a known value: 44000.5 (2020-07-19 12:00:00)
    const DATE testValue = 44000.5;
    trace("DateOut: returning 44000.5");

    if (outDate)
        *outDate = testValue;

    return S_OK;
}

// 3. [in,out] DATE* - server reads and modifies
HRESULT __stdcall CDateTest::DateInOut(DATE* inOutDate)
{
    std::ostringstream sout;
    sout << "DateInOut: received inOutDate=" << (inOutDate ? *inOutDate : 0.0);
    trace(sout.str().c_str());

    if (inOutDate)
    {
        // Add 1.0 (one day) to the value
        *inOutDate = *inOutDate + 1.0;
        sout.str("");
        sout << "DateInOut: modified to " << *inOutDate;
        trace(sout.str().c_str());
    }

    return S_OK;
}

// 4. Typed double (VT_R8) round-trip for comparison with DATE
HRESULT __stdcall CDateTest::DoubleRoundTrip(double inVal, double* outVal)
{
    std::ostringstream sout;
    sout << "DoubleRoundTrip: received inVal=" << inVal;
    trace(sout.str().c_str());

    if (outVal)
        *outVal = inVal;

    return S_OK;
}

// 5. [out] double* - server creates a double value
HRESULT __stdcall CDateTest::DoubleOut(double* outVal)
{
    const double testValue = 44000.5;
    trace("DoubleOut: returning 44000.5");

    if (outVal)
        *outVal = testValue;

    return S_OK;
}

// 6. [in,out] double* - server reads and modifies
HRESULT __stdcall CDateTest::DoubleInOut(double* inOutVal)
{
    std::ostringstream sout;
    sout << "DoubleInOut: received inOutVal=" << (inOutVal ? *inOutVal : 0.0);
    trace(sout.str().c_str());

    if (inOutVal)
    {
        *inOutVal = *inOutVal + 1.0;
        sout.str("");
        sout << "DoubleInOut: modified to " << *inOutVal;
        trace(sout.str().c_str());
    }

    return S_OK;
}

// 7. VARIANT round-trip: [in] VARIANT -> [out,retval] VARIANT
// Returns the exact same VARIANT received (vt and value preserved)
HRESULT __stdcall CDateTest::VariantRoundTrip(VARIANT inVar, VARIANT* outVar)
{
    std::ostringstream sout;
    sout << "VariantRoundTrip: received vt=0x" << std::hex << inVar.vt;
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

// 8. GetVariantVt: [in] VARIANT -> [out,retval] LONG (vt)
// Observes: what VT the server sees when Python passes a VARIANT
HRESULT __stdcall CDateTest::GetVariantVt(VARIANT inVar, LONG* vt)
{
    std::ostringstream sout;
    sout << "GetVariantVt: received vt=0x" << std::hex << inVar.vt;
    trace(sout.str().c_str());

    if (vt)
        *vt = inVar.vt;

    return S_OK;
}

// 9. GetVariantDateValue: [in] VARIANT -> [out,retval] double
// Observes: the raw double value inside VT_DATE or VT_R8 VARIANT
HRESULT __stdcall CDateTest::GetVariantDateValue(VARIANT inVar, double* dateVal)
{
    std::ostringstream sout;
    sout << "GetVariantDateValue: received vt=0x" << std::hex << inVar.vt;
    trace(sout.str().c_str());

    if (dateVal)
    {
        if (inVar.vt == VT_DATE)
        {
            *dateVal = inVar.date;
            sout.str("");
            sout << "GetVariantDateValue: VT_DATE value=" << inVar.date;
            trace(sout.str().c_str());
        }
        else if (inVar.vt == VT_R8)
        {
            *dateVal = inVar.dblVal;
            sout.str("");
            sout << "GetVariantDateValue: VT_R8 value=" << inVar.dblVal;
            trace(sout.str().c_str());
        }
        else
        {
            *dateVal = 0.0;
            trace("GetVariantDateValue: not VT_DATE or VT_R8, returning 0.0");
        }
    }

    return S_OK;
}

// 10. CreateVariantDate: server creates VT_DATE VARIANT with given value
HRESULT __stdcall CDateTest::CreateVariantDate(double dateVal, VARIANT* outVar)
{
    std::ostringstream sout;
    sout << "CreateVariantDate: creating VT_DATE with value=" << dateVal;
    trace(sout.str().c_str());

    if (outVar)
    {
        VariantInit(outVar);
        outVar->vt = VT_DATE;
        outVar->date = dateVal;
    }

    return S_OK;
}

// 11. CreateVariantR8: server creates VT_R8 VARIANT with given value
HRESULT __stdcall CDateTest::CreateVariantR8(double r8Val, VARIANT* outVar)
{
    std::ostringstream sout;
    sout << "CreateVariantR8: creating VT_R8 with value=" << r8Val;
    trace(sout.str().c_str());

    if (outVar)
    {
        VariantInit(outVar);
        outVar->vt = VT_R8;
        outVar->dblVal = r8Val;
    }

    return S_OK;
}

// 12. SAFEARRAY(DATE) round-trip
HRESULT __stdcall CDateTest::SafeArrayDateRoundTrip(SAFEARRAY* inArr, SAFEARRAY** outArr)
{
    trace("SafeArrayDateRoundTrip: called");

    if (!inArr || !outArr)
        return E_POINTER;

    // Copy the input array to output
    HRESULT hr = SafeArrayCopy(inArr, outArr);
    return hr;
}

// 13. SAFEARRAY(VARIANT) round-trip (elements can be VT_DATE or VT_R8)
HRESULT __stdcall CDateTest::SafeArrayVariantRoundTrip(SAFEARRAY* inArr, SAFEARRAY** outArr)
{
    trace("SafeArrayVariantRoundTrip: called");

    if (!inArr || !outArr)
        return E_POINTER;

    HRESULT hr = SafeArrayCopy(inArr, outArr);
    return hr;
}

// 14. Get VARTYPE of each element in a SAFEARRAY(VARIANT)
HRESULT __stdcall CDateTest::SafeArrayVariantGetVts(SAFEARRAY* inArr, SAFEARRAY** outVts)
{
    trace("SafeArrayVariantGetVts: called");

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

// 15. Property getter for DATE
HRESULT __stdcall CDateTest::get_DateProp(DATE* pVal)
{
    std::ostringstream sout;
    sout << "get_DateProp: returning " << m_dateProp;
    trace(sout.str().c_str());

    if (pVal)
        *pVal = m_dateProp;

    return S_OK;
}

// 16. Property setter for DATE
HRESULT __stdcall CDateTest::put_DateProp(DATE newVal)
{
    std::ostringstream sout;
    sout << "put_DateProp: setting to " << newVal;
    trace(sout.str().c_str());

    m_dateProp = newVal;
    return S_OK;
}

// 17. Record with DATE field round-trip
HRESULT __stdcall CDateTest::RecordDateRoundTrip(StructDateRecord* inRec, StructDateRecord* outRec)
{
    trace("RecordDateRoundTrip: called");

    if (!inRec || !outRec)
        return E_POINTER;

    // Copy the record
    *outRec = *inRec;

    std::ostringstream sout;
    sout << "RecordDateRoundTrip: inRec.dateField=" << inRec->dateField;
    trace(sout.str().c_str());

    return S_OK;
}

// 18. VT_BYREF|VT_DATE variant handling
HRESULT __stdcall CDateTest::ByRefDateVariant(VARIANT* inOutVar)
{
    trace("ByRefDateVariant: called");

    if (!inOutVar)
        return E_POINTER;

    std::ostringstream sout;
    sout << "ByRefDateVariant: received vt=0x" << std::hex << inOutVar->vt;
    trace(sout.str().c_str());

    // If it's VT_BYREF|VT_DATE, read the value and modify it
    if ((inOutVar->vt & VT_BYREF) && (inOutVar->vt & VT_TYPEMASK) == VT_DATE)
    {
        DATE* pDate = inOutVar->pdate;
        if (pDate)
        {
            sout.str("");
            sout << "ByRefDateVariant: read value=" << *pDate;
            trace(sout.str().c_str());

            // Add 1 day
            *pDate = *pDate + 1.0;

            sout.str("");
            sout << "ByRefDateVariant: wrote back value=" << *pDate;
            trace(sout.str().c_str());
        }
    }

    return S_OK;
}

// Initialization - load type info
HRESULT CDateTest::Init()
{
    HRESULT hr;

    if (m_pITypeInfo == NULL)
    {
        ITypeLib* pITypeLib = NULL;
        hr = ::LoadRegTypeLib(LIBID_ComtypesCppTestSrvLib,
                              1, 0, // Major/Minor version numbers
                              0x00,
                              &pITypeLib);
        if (FAILED(hr))
        {
            trace("LoadRegTypeLib Failed.", hr);
            return hr;
        }

        hr = pITypeLib->GetTypeInfoOfGuid(IID_IDualDateTest, &m_pITypeInfo);
        pITypeLib->Release();
        if (FAILED(hr))
        {
            trace("GetTypeInfoOfGuid failed.", hr);
            return hr;
        }
    }
    return S_OK;
}

// Constructor
CDateTest::CDateTest(IUnknown* pUnknownOuter) : CUnknown(pUnknownOuter), m_dateProp(0.0), m_pITypeInfo(NULL)
{
    // Empty
}

// Destructor
CDateTest::~CDateTest()
{
    if (m_pITypeInfo != NULL)
    {
        m_pITypeInfo->Release();
    }
    trace("Destroy self.");
}

// NondelegatingQueryInterface
HRESULT __stdcall CDateTest::NondelegatingQueryInterface(const IID& iid, void** ppv)
{
    if (iid == IID_IDualDateTest)
    {
        return FinishQI(static_cast<IDualDateTest*>(this), ppv);
    }
    else if (iid == DIID_IDispDateTest)
    {
        trace("Queried for IDispDateTest.");
        return FinishQI(static_cast<IDispatch*>(this), ppv);
    }
    else if (iid == IID_IDispatch)
    {
        trace("Queried for IDispatch.");
        return FinishQI(static_cast<IDispatch*>(this), ppv);
    }
    else
    {
        return CUnknown::NondelegatingQueryInterface(iid, ppv);
    }
}

// IDispatch implementation
HRESULT __stdcall CDateTest::GetTypeInfoCount(UINT* pCountTypeInfo)
{
    trace("GetTypeInfoCount call succeeded.");
    *pCountTypeInfo = 1;
    return S_OK;
}

HRESULT __stdcall CDateTest::GetTypeInfo(UINT iTypeInfo, LCID, ITypeInfo** ppITypeInfo)
{
    *ppITypeInfo = NULL;

    if (iTypeInfo != 0)
    {
        trace("GetTypeInfo call failed -- bad iTypeInfo index.");
        return DISP_E_BADINDEX;
    }

    trace("GetTypeInfo call succeeded.");
    m_pITypeInfo->AddRef();
    *ppITypeInfo = m_pITypeInfo;
    return S_OK;
}

HRESULT __stdcall CDateTest::GetIDsOfNames(const IID& iid, OLECHAR** arrayNames, UINT countNames, LCID, DISPID* arrayDispIDs)
{
    if (iid != IID_NULL)
    {
        trace("GetIDsOfNames call failed -- bad IID.");
        return DISP_E_UNKNOWNINTERFACE;
    }

    trace("GetIDsOfNames call succeeded.");
    HRESULT hr = m_pITypeInfo->GetIDsOfNames(arrayNames, countNames, arrayDispIDs);
    return hr;
}

HRESULT __stdcall CDateTest::Invoke(DISPID dispidMember, const IID& iid, LCID, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pvarResult, EXCEPINFO* pExcepInfo, UINT* pArgErr)
{
    if (iid != IID_NULL)
    {
        trace("Invoke call failed -- bad IID.");
        return DISP_E_UNKNOWNINTERFACE;
    }

    ::SetErrorInfo(0, NULL);
    trace("Invoke call succeeded.");
    HRESULT hr = m_pITypeInfo->Invoke(static_cast<IDispatch*>(this), dispidMember, wFlags, pDispParams, pvarResult, pExcepInfo, pArgErr);
    return hr;
}

// Creation function used by CFactory
HRESULT CDateTest::CreateInstance(IUnknown* pUnknownOuter, CUnknown** ppNewComponent)
{
    if (pUnknownOuter != NULL)
    {
        return CLASS_E_NOAGGREGATION;
    }

    *ppNewComponent = new CDateTest(pUnknownOuter);
    return S_OK;
}
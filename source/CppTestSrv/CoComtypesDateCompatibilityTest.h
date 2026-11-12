// CoComtypesDateCompatibilityTest.h
// Test double for VT_DATE / DATE marshaling characterization tests.
#include "CUnknown.h"
#include "Iface.h"

// Forward declaration for dispinterface IID (not in MIDL-generated header)
EXTERN_C const IID DIID_IDispDateTest;

class CDateTest : public CUnknown,
                  public IDualDateTest
{
public:
    static HRESULT CreateInstance(IUnknown* pUnknownOuter, CUnknown** ppNewComponent);

private:
    DECLARE_IUNKNOWN
    virtual HRESULT __stdcall NondelegatingQueryInterface(const IID& iid, void** ppv);

    // Interface IDualDateTest - Typed DATE round-trip
    virtual HRESULT __stdcall EchoDate(DATE inDate, DATE* outDate);
    virtual HRESULT __stdcall GetDate44031dot5(DATE* outDate);
    virtual HRESULT __stdcall AddOneToDateInOut(DATE* inOutDate);

    // Typed double (VT_R8) round-trip for comparison
    virtual HRESULT __stdcall EchoDouble(double inVal, double* outVal);
    virtual HRESULT __stdcall GetDouble44031dot5(double* outVal);
    virtual HRESULT __stdcall AddOneToDoubleInOut(double* inOutVal);

    // Observe received VARIANT (DATE/R8 specific)
    virtual HRESULT __stdcall GetVariantDateValue(VARIANT inVar, double* dateVal);

    // Server creates VARIANTs with specific VT (DATE/R8 specific)
    virtual HRESULT __stdcall CreateVariantDate(double dateVal, VARIANT* outVar);
    virtual HRESULT __stdcall CreateVariantR8(double r8Val, VARIANT* outVar);

    // SAFEARRAY tests
    virtual HRESULT __stdcall EchoSafeArrayDate(SAFEARRAY* inArr, SAFEARRAY** outArr);

    // Property tests
    virtual HRESULT __stdcall get_DateProp(DATE* pVal);
    virtual HRESULT __stdcall put_DateProp(DATE newVal);

    // Record with DATE field
    virtual HRESULT __stdcall EchoRecordDate(StructDateRecord* inRec, StructDateRecord* outRec);

    // VT_BYREF|VT_DATE
    virtual HRESULT __stdcall AddOneToByRefDateVariant(VARIANT* inOutVar);

    // IDispatch
    virtual HRESULT __stdcall GetTypeInfoCount(UINT* pCountTypeInfo);
    virtual HRESULT __stdcall GetTypeInfo(UINT iTypeInfo, LCID, ITypeInfo** ppITypeInfo);
    virtual HRESULT __stdcall GetIDsOfNames(const IID& iid, OLECHAR** arrayNames, UINT countNames, LCID, DISPID* arrayDispIDs);
    virtual HRESULT __stdcall Invoke(DISPID dispidMember, const IID& iid, LCID, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pvarResult, EXCEPINFO* pExcepInfo, UINT* pArgErr);

    // Initialization
    virtual HRESULT Init();

    CDateTest(IUnknown* pUnknownOuter);
    ~CDateTest();

    double m_dateProp;
    ITypeInfo* m_pITypeInfo;
};

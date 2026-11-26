/*
 * CoOutRecordTest.h - Component for [out] record ownership tests
 *
 * This test double returns records through [out] parameters.  A record
 * (StructOutRecordTest) owns two managed resources: a BSTR (text) and an
 * IUnknown* (unk).  When the server fills an [out] record, ownership of
 * those resources is transferred to the caller.  The caller must release
 * them (for example through IRecordInfo::RecordClear) when it is done with
 * the record.
 *
 * The unk field points at a "sentinel" object: a minimal IUnknown
 * implementation whose only purpose is to be reference counted.  The
 * server keeps a per-instance count of live sentinels, which the tests
 * read through GetLiveSentinelCount.  If the caller releases the record
 * contents, the sentinel's reference count drops to zero and the live
 * count decreases; if the caller leaks the record, the sentinel stays
 * alive and the live count does not decrease.
 *
 * The sentinel is created with a reference count of 1 and that reference
 * is handed to the caller without an extra AddRef.  This mirrors the COM
 * convention that an [out] interface pointer is returned with a reference
 * already held for the caller.
 *
 * Note on identity verification: other test doubles verify that a returned
 * record has the expected GUID by calling IRecordInfo::GetGuid.  That check
 * applies to the record type, not to the sentinel IUnknown.  Verifying that
 * a returned IUnknown* was produced by this server would require adding a
 * private marker interface and a method to query it, which is a larger
 * change than this test double needs, so it is omitted.  The live sentinel
 * count already distinguishes a real sentinel from an arbitrary pointer.
 */

#ifndef __CoOutRecordTest_h__
#define __CoOutRecordTest_h__

#include <memory>

#include "Iface.h"
#include "CUnknown.h"

///////////////////////////////////////////////////////////
//
// Sentinel object
//
// A minimal IUnknown implementation.  It is not an ATL CComObject; it is a
// plain C++ object so that its lifetime is fully under this test double's
// control.  Each sentinel holds a copy of the shared live-count so that the
// count is decremented exactly when the sentinel is destroyed.
//
class CSentinel : public IUnknown
{
public:
	explicit CSentinel(std::shared_ptr<LONG> liveCount);
	virtual ~CSentinel();

	// IUnknown
	virtual HRESULT __stdcall QueryInterface(const IID& iid, void** ppv);
	virtual ULONG __stdcall AddRef();
	virtual ULONG __stdcall Release();

private:
	LONG m_cRef;
	std::shared_ptr<LONG> m_liveCount;
};

///////////////////////////////////////////////////////////
//
// Component for [out] record ownership tests
//
class COutRecordTest : public CUnknown,
                       public IOutRecordTest
{
public:
	// Creation
	static HRESULT CreateInstance(IUnknown* pUnknownOuter,
	                              CUnknown** ppNewComponent);

private:
	// Declare the delegating IUnknown.
	DECLARE_IUNKNOWN

	// IUnknown
	virtual HRESULT __stdcall NondelegatingQueryInterface(const IID& iid,
	                                                      void** ppv);

	// Interface IOutRecordTest
	virtual HRESULT __stdcall GetRecord(LONG cchText, StructOutRecordTest* rec);
	virtual HRESULT __stdcall GetRecordRetval(StructOutRecordTest* rec);
	virtual HRESULT __stdcall GetRecordsAround(StructOutRecordTest* first,
	                                           LONG* middle,
	                                           StructOutRecordTest* last);
	virtual HRESULT __stdcall ModifyRecord(StructOutRecordTest* rec);
	virtual HRESULT __stdcall FailGetRecord(StructOutRecordTest* rec);
	virtual HRESULT __stdcall GetLiveSentinelCount(LONG* count);

	// Initialization
	virtual HRESULT Init();

	// Constructor
	COutRecordTest(IUnknown* pUnknownOuter);

	// Destructor
	~COutRecordTest();

	// Create a new sentinel and return it with a reference held for the
	// caller (reference count 1).
	IUnknown* CreateSentinel();

	// Per-instance count of live sentinels.
	std::shared_ptr<LONG> m_liveSentinelCount;
};

#endif

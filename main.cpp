#include <windows.h>
#include <ShlObj.h>
#pragma comment (lib, "ole32.lib")

// ===========================================================

HRESULT __stdcall DllGetClassObject(REFCLSID, REFIID, LPVOID*);
HRESULT __stdcall DllCanUnloadNow();
BOOL __stdcall DllMain(HMODULE, DWORD, LPVOID);

// ===========================================================

ULONG gRefCount = 0x0;
CLSID ext_clsid = { 0x0 };
UINT format_ids[] = { CF_HDROP, RegisterClipboardFormatW(CFSTR_FILEDESCRIPTORW), RegisterClipboardFormatW(CFSTR_FILEDESCRIPTORA),
	RegisterClipboardFormatW(CFSTR_FILECONTENTS), RegisterClipboardFormatW(CFSTR_FILENAMEMAPW), RegisterClipboardFormatW(CFSTR_FILENAMEMAPA),
	RegisterClipboardFormatW(CFSTR_SHELLIDLIST) };

LPCWSTR str_clsid = L"{c1a276d5-51b1-41c7-9f0d-112cbbee58fe}";

// ===========================================================

class CAdvisorySinkEnumerator : public IEnumSTATDATA {
private:
	ULONG ref_count, current_index, connections_count;
	STATDATA* connections;
public:
	CAdvisorySinkEnumerator(STATDATA* connections, ULONG connections_count) {
		InterlockedIncrement(&gRefCount);
		InterlockedIncrement(&ref_count);
		
		ref_count = 0x0;
		this->connections_count = connections_count;
		this->connections = connections;
		current_index = 0x0;
	}
	~CAdvisorySinkEnumerator() {
		InterlockedDecrement(&gRefCount);
	}

	// IUKnown
	ULONG __stdcall AddRef() {
		InterlockedIncrement(&ref_count);
		return ref_count;
	}
	ULONG __stdcall Release() {
		InterlockedDecrement(&ref_count);
		if (!ref_count) {
			delete this;
			return 0x0;
		}
		return ref_count;
	}
	HRESULT __stdcall QueryInterface(REFIID interface_id, LPVOID* interface_buffer) {
		if (!interface_buffer) return E_INVALIDARG;
		if (interface_id == IID_IUnknown) {
			*interface_buffer = (IUnknown*)this;
			AddRef();
			return S_OK;
		}
		if (interface_id == IID_IEnumSTATDATA) {
			*interface_buffer = (IEnumSTATDATA*)this;
			AddRef();
			return S_OK;
		}
		*interface_buffer = 0x0;
		return E_NOINTERFACE;
	}

	// IEnumSTATDATA
	HRESULT __stdcall Clone(IEnumSTATDATA** enumerator) {
		if (!enumerator) return E_INVALIDARG;
		*enumerator = new CAdvisorySinkEnumerator(connections, connections_count);
		if (!(*enumerator)) return E_OUTOFMEMORY;
		(*enumerator)->Skip(current_index);
		return S_OK;
	}
	HRESULT __stdcall Next(ULONG items_count, STATDATA* items, ULONG* feteched_items_count) {
		if (items_count != 0x1 || !items) return E_INVALIDARG;
		if (current_index >= items_count) return S_FALSE;

		items->advf = connections[current_index].advf;
		items->dwConnection = connections[current_index].dwConnection;
		items->formatetc = connections[current_index].formatetc;
		items->pAdvSink = connections[current_index].pAdvSink;

		InterlockedIncrement(&current_index);

		if (feteched_items_count) *feteched_items_count = 0x1;
		return S_OK;
	}
	HRESULT __stdcall Reset() {
		InterlockedExchange(&current_index, 0x0);
		return S_OK;
	}
	HRESULT __stdcall Skip(ULONG skipped_items_count) {
		if ((current_index + skipped_items_count) > connections_count) return E_FAIL;
		InterlockedAdd((LONG*)(&current_index), skipped_items_count);
		return S_OK;
	}
};
class CFormatEtcEnumerator : public IEnumFORMATETC {
private:
	ULONG ref_count, current_index, formats_count;
	FORMATETC* supported_formats;
public:
	CFormatEtcEnumerator(FORMATETC* supported_formats, ULONG formats_count) {
		InterlockedIncrement(&gRefCount);
		InterlockedIncrement(&ref_count);
		
		ref_count = 0x0;
		this->formats_count = formats_count;
		this->supported_formats = supported_formats;
		current_index = 0x0;
	}
	~CFormatEtcEnumerator() {
		InterlockedDecrement(&gRefCount);
	}

	// IUKnown
	ULONG __stdcall AddRef() {
		InterlockedIncrement(&ref_count);
		return ref_count;
	}
	ULONG __stdcall Release() {
		InterlockedDecrement(&ref_count);
		if (!ref_count) {
			delete this;
			return 0x0;
		}
		return ref_count;
	}
	HRESULT __stdcall QueryInterface(REFIID interface_id, LPVOID* interface_buffer) {
		if (!interface_buffer) return E_INVALIDARG;
		if (interface_id == IID_IUnknown) {
			*interface_buffer = (IUnknown*)this;
			AddRef();
			return S_OK;
		}
		if (interface_id == IID_IEnumFORMATETC) {
			*interface_buffer = (IEnumFORMATETC*)this;
			AddRef();
			return S_OK;
		}
		*interface_buffer = 0x0;
		return E_NOINTERFACE;
	}

	// IEnumFORMATETC
	HRESULT __stdcall Clone(IEnumFORMATETC** enumerator) {
		if (!enumerator) return E_INVALIDARG;
		*enumerator = new CFormatEtcEnumerator(supported_formats, formats_count);
		if (!(*enumerator)) return E_OUTOFMEMORY;
		(*enumerator)->Skip(current_index);
		return S_OK;
	}
	HRESULT __stdcall Next(ULONG formats_count, FORMATETC* formats, ULONG* fetched_formats_count) {
		if (formats_count != 0x1 || !formats) return E_INVALIDARG;
		if (current_index >= formats_count) return S_FALSE;
		
		formats->cfFormat = supported_formats[current_index].cfFormat;
		formats->dwAspect = supported_formats[current_index].dwAspect;
		formats->tymed = supported_formats[current_index].tymed;

		InterlockedIncrement(&current_index);

		if (fetched_formats_count) *fetched_formats_count = 0x1;
		return S_OK;
	}
	HRESULT __stdcall Reset() {
		InterlockedExchange(&current_index, 0x0);
		return S_OK;
	}
	HRESULT __stdcall Skip(ULONG skipped_items_count) {
		if ((current_index + skipped_items_count) > formats_count) return E_FAIL;
		InterlockedAdd((LONG*)(&current_index), skipped_items_count);
		return S_OK;
	}
};
class CDataObejct : public IPersistFile, public IDataObject {
private:
	ULONG ref_count, formats_count, connections_count;
	FORMATETC* supported_formats;
	STATDATA* connections;
	LPWSTR file_path;

	HRESULT __stdcall AddConnection(FORMATETC data_format, IAdviseSink* caller_sink, DWORD advf, DWORD* connection_id) {
		HANDLE default_heap = GetProcessHeap();
		if (!connections) connections = (STATDATA*)HeapAlloc(default_heap, 0x8, sizeof STATDATA * (connections_count + 0x1));
		else connections = (STATDATA*)HeapReAlloc(default_heap, 0x8, connections, sizeof STATDATA * (connections_count + 0x1));

		if (!connections) return E_OUTOFMEMORY;

		connections[connections_count].advf = advf;
		connections[connections_count].dwConnection = connections_count;
		connections[connections_count].formatetc = data_format;
		connections[connections_count].pAdvSink = caller_sink;

		if (connection_id) *connection_id = connections_count;
		InterlockedIncrement(&connections_count);

		return S_OK;
	}
	HRESULT __stdcall DeleteConnection(DWORD connection_id) {
		if (connection_id >= connections_count) return E_BOUNDS;
		connections[connection_id].pAdvSink = 0x0;
		return S_OK;
	}

public:
	CDataObejct() {
		InterlockedIncrement(&gRefCount);
		InterlockedIncrement(&ref_count);
		
		ref_count = 0x0;
		formats_count = 0x7;
		connections_count = 0x0;
		file_path = 0x0;

		supported_formats = (FORMATETC*)HeapAlloc(GetProcessHeap(), 0x8, sizeof FORMATETC * formats_count);
		connections = 0x0;

		if (supported_formats){
			for (UINT i = 0x0; i < formats_count; i++) {
				supported_formats[i].cfFormat = format_ids[i];
				supported_formats[i].tymed = TYMED_HGLOBAL;
				supported_formats[i].dwAspect = DVASPECT_CONTENT;
				supported_formats[i].lindex = i == 0x3 ? 0x0 : -1;
			}
		}
	}
	~CDataObejct() {
		if (supported_formats) HeapFree(GetProcessHeap(), 0x8, supported_formats);
		if (connections) HeapFree(GetProcessHeap(), 0x8, connections);
		if (file_path) HeapFree(GetProcessHeap(), 0x8, file_path);
		InterlockedDecrement(&gRefCount);
	}

	// IUKnown
	ULONG __stdcall AddRef() {
		InterlockedIncrement(&ref_count);
		return ref_count;
	}
	ULONG __stdcall Release() {
		InterlockedDecrement(&ref_count);
		if (!ref_count) {
			delete this;
			return 0x0;
		}
		return ref_count;
	}
	HRESULT __stdcall QueryInterface(REFIID interface_id, LPVOID* interface_buffer) {
		if (!interface_buffer) return E_INVALIDARG;
		if (interface_id == IID_IPersistFile) {
			*interface_buffer = (IPersistFile*)this;
			AddRef();
			return S_OK;
		}
		if (interface_id == IID_IUnknown || interface_id == IID_IDataObject) {
			*interface_buffer = (IDataObject*)this;
			AddRef();
			return S_OK;
		}
		*interface_buffer = 0x0;
		return E_NOINTERFACE;
	}

	// IDataObject
	HRESULT __stdcall DAdvise(FORMATETC* data_format, DWORD advf, IAdviseSink* caller_sink, DWORD* connection_id) {
		if (!data_format || !caller_sink) return E_INVALIDARG;

		DWORD __connection_id = 0x0;
		HRESULT rs = AddConnection(*data_format, caller_sink, advf, &__connection_id);
		if (SUCCEEDED(rs)) {
			if (connection_id) *connection_id = __connection_id;
		}

		return rs;
	}
	HRESULT __stdcall DUnadvise(DWORD connection_id) {
		return DeleteConnection(connection_id);
	}
	HRESULT __stdcall EnumDAdvise(IEnumSTATDATA** enumerator) {
		if (!enumerator) return E_INVALIDARG;
		
		CAdvisorySinkEnumerator* connection_enumerator = new CAdvisorySinkEnumerator(connections, connections_count);
		if (!connection_enumerator) return E_OUTOFMEMORY;

		HRESULT rs = connection_enumerator->QueryInterface(IID_IEnumSTATDATA, (LPVOID*)enumerator);
		if (FAILED(rs)) delete connection_enumerator;

		return rs;
	}
	HRESULT __stdcall EnumFormatEtc(DWORD direction, IEnumFORMATETC** enumerator) {
		if (!enumerator) return E_INVALIDARG;

		CFormatEtcEnumerator* format_enumerator = new CFormatEtcEnumerator(direction == DATADIR_GET ? supported_formats : 0x0,
			direction == DATADIR_GET ? formats_count : 0x0);
		if (!format_enumerator) return E_OUTOFMEMORY;

		HRESULT rs = format_enumerator->QueryInterface(IID_IEnumSTATDATA, (LPVOID*)enumerator);
		if (FAILED(rs)) delete format_enumerator;

		return rs;
	}
	HRESULT __stdcall GetCanonicalFormatEtc(FORMATETC* in, FORMATETC* out) {
		if (!in || !out) return E_INVALIDARG;

		CopyMemory(out, in, sizeof FORMATETC);
		out->ptd = 0x0;

		return DATA_S_SAMEFORMATETC;
	}
	HRESULT __stdcall GetData(FORMATETC* data_format, STGMEDIUM* transfer_medium) {
		if (!this->file_path) return E_UNEXPECTED;
		if (!data_format || !transfer_medium) return E_INVALIDARG;

		HRESULT rs = QueryGetData(data_format);
		if (FAILED(rs)) return rs;

		INT path_len = lstrlenW(this->file_path);

		HGLOBAL gMemObject = 0x0;
		if (data_format->cfFormat == supported_formats[0x0].cfFormat) {
			// CF_HDROP
			gMemObject = GlobalAlloc(GMEM_FIXED | GMEM_ZEROINIT, sizeof DROPFILES + path_len * 0x2 + 0x4);
			if (gMemObject) {
				DROPFILES* data = (DROPFILES*)gMemObject;
				data->pFiles = sizeof DROPFILES;
				LPWSTR pFiles = (LPWSTR)(data + 0x1);
				CopyMemory(pFiles, this->file_path, path_len * 0x2);
			}
		}
		else if (data_format->cfFormat == supported_formats[0x1].cfFormat) {
			// CF_FILEGROUPDESCRIPTORW
			gMemObject = GlobalAlloc(GMEM_FIXED | GMEM_ZEROINIT, sizeof FILEGROUPDESCRIPTORW);
			if (gMemObject) {
				FILEGROUPDESCRIPTORW* data = (FILEGROUPDESCRIPTORW*)gMemObject;
				HANDLE file = CreateFileW(this->file_path, GENERIC_READ, FILE_SHARE_READ, 0x0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0x0);
				if (file != INVALID_HANDLE_VALUE) {
					data->cItems = 0x1;

					data->fgd[0x0].dwFlags = FD_FILESIZE | FD_CREATETIME | FD_UNICODE | FD_ATTRIBUTES | FD_ACCESSTIME | FD_WRITESTIME;
					data->fgd[0x0].dwFileAttributes = GetFileAttributesW(this->file_path);
					GetFileTime(file, &data->fgd[0x0].ftCreationTime, &data->fgd[0x0].ftLastAccessTime, &data->fgd[0x0].ftLastWriteTime);
					CopyMemory(data->fgd[0x0].cFileName, this->file_path, path_len * 0x2);
					
					CloseHandle(file);
				}
			}
		}
		else if (data_format->cfFormat == supported_formats[0x2].cfFormat) {
			// CF_FILEGROUPDESCRIPTORA
			gMemObject = GlobalAlloc(GMEM_FIXED | GMEM_ZEROINIT, sizeof FILEGROUPDESCRIPTORA);
			if (gMemObject) {
				FILEGROUPDESCRIPTORA* data = (FILEGROUPDESCRIPTORA*)gMemObject;
				HANDLE file = CreateFileW(this->file_path, GENERIC_READ, FILE_SHARE_READ, 0x0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0x0);
				if (file != INVALID_HANDLE_VALUE) {
					data->cItems = 0x1;

					data->fgd[0x0].dwFlags = FD_FILESIZE | FD_CREATETIME | FD_ATTRIBUTES | FD_ACCESSTIME | FD_WRITESTIME;
					data->fgd[0x0].dwFileAttributes = GetFileAttributesW(this->file_path);
					GetFileTime(file, &data->fgd[0x0].ftCreationTime, &data->fgd[0x0].ftLastAccessTime, &data->fgd[0x0].ftLastWriteTime);
					WideCharToMultiByte(CP_ACP, 0x0, this->file_path, path_len, data->fgd[0x0].cFileName, path_len, 0x0, 0x0);

					CloseHandle(file);
				}
			}
		}
		else if (data_format->cfFormat == supported_formats[0x3].cfFormat) {
			// CF_FILECONTENTS
			gMemObject = GlobalAlloc(GMEM_FIXED | GMEM_ZEROINIT, sizeof FILEDESCRIPTORW);
			if (gMemObject) {
				FILEDESCRIPTORW* data = (FILEDESCRIPTORW*)gMemObject;
				HANDLE file = CreateFileW(this->file_path, GENERIC_READ, FILE_SHARE_READ, 0x0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0x0);
				if (file != INVALID_HANDLE_VALUE) {
					data->dwFlags = FD_FILESIZE | FD_CREATETIME | FD_ATTRIBUTES | FD_ACCESSTIME | FD_WRITESTIME;
					data->dwFileAttributes = GetFileAttributesW(this->file_path);
					GetFileTime(file, &data->ftCreationTime, &data->ftLastAccessTime, &data->ftLastWriteTime);
					CopyMemory(data->cFileName, this->file_path, path_len * 0x2);

					CloseHandle(file);
				}
			}
		}
		else if (data_format->cfFormat == supported_formats[0x4].cfFormat) {
			// CF_FILENAMEMAPW
			gMemObject = GlobalAlloc(GMEM_FIXED | GMEM_ZEROINIT, path_len * 0x2 + 0x4);
			if (gMemObject) {
				LPWSTR data = (LPWSTR)gMemObject;
				CopyMemory(data, this->file_path, path_len * 0x2);
			}
		}
		else if (data_format->cfFormat == supported_formats[0x5].cfFormat) {
			// CF_FILENAMEMAPA
			gMemObject = GlobalAlloc(GMEM_FIXED | GMEM_ZEROINIT, path_len + 0x2);
			if (gMemObject) {
				LPSTR data = (LPSTR)gMemObject;
				WideCharToMultiByte(CP_ACP, 0x0, this->file_path, path_len, data, path_len, 0x0, 0x0);
			}
		}
		else if (data_format->cfFormat == supported_formats[0x6].cfFormat) {
			// CF_SHELLIDLIST
			IShellFolder* SHDesktop = 0x0;
			HRESULT rs = SHGetDesktopFolder(&SHDesktop);
			if(SUCCEEDED(rs)) {
				LPITEMIDLIST file_idl = 0x0;
				rs = SHDesktop->ParseDisplayName(0x0, 0x0, this->file_path, 0x0, &file_idl, 0x0);
				if(SUCCEEDED(rs)) {
					UINT idl_size = ILGetSize(file_idl);
					gMemObject = GlobalAlloc(GMEM_FIXED | GMEM_ZEROINIT, sizeof CIDA + sizeof UINT + 0x2 + idl_size);
					if (gMemObject) {
						CIDA* data = (CIDA*)gMemObject;
						data->cidl = 0x1;
						data->aoffset[0x0] = sizeof CIDA + sizeof UINT;
						data->aoffset[0x1] = sizeof CIDA + sizeof UINT + 0x2;
						SHITEMID* parent_idl = (SHITEMID*)((UINT_PTR)data + data->aoffset[0x0]);
						parent_idl->cb = 0x2;
						LPITEMIDLIST relative_idl_0 = (LPITEMIDLIST)((UINT_PTR)data + data->aoffset[0x1]);
						CopyMemory(relative_idl_0, file_idl, idl_size);
					}		
				}
				SHDesktop->Release();
			}

		}
		
		if (!gMemObject) return E_OUTOFMEMORY;

		transfer_medium->hGlobal = gMemObject;
		transfer_medium->tymed = TYMED_HGLOBAL;

		return S_OK;
	}
	HRESULT __stdcall GetDataHere(FORMATETC* data_format, STGMEDIUM* transfer_medium) {
		UNREFERENCED_PARAMETER(data_format);
		UNREFERENCED_PARAMETER(transfer_medium);
		return E_NOTIMPL;
	}
	HRESULT __stdcall QueryGetData(FORMATETC* data_format) {
		if (!data_format) return E_INVALIDARG;

		if (data_format->tymed != TYMED_HGLOBAL) return DV_E_FORMATETC;

		BOOLEAN b_supported = 0x0;
		for (UINT i = 0x0; i < formats_count; i++) if (data_format->cfFormat == supported_formats[i].cfFormat) {
			b_supported = 0x1;
			break;
		}

		if (data_format->cfFormat == supported_formats[0x3].cfFormat && data_format->lindex != 0x0) return DV_E_FORMATETC;

		return b_supported ? S_OK : DV_E_FORMATETC;
	}
	HRESULT __stdcall SetData(FORMATETC* data_format, STGMEDIUM* transfer_medium, BOOL b_release) {
		UNREFERENCED_PARAMETER(data_format);
		UNREFERENCED_PARAMETER(transfer_medium);
		UNREFERENCED_PARAMETER(b_release);
		return E_NOTIMPL;
	}

	// IPersistFile
	HRESULT __stdcall GetClassID(CLSID* clsid) {
		if (clsid) return E_INVALIDARG;
		*clsid = ext_clsid;
		return S_OK;
	}
	HRESULT __stdcall GetCurFile(LPOLESTR* file_path) {
		if (!file_path) return E_INVALIDARG;
		if (!this->file_path) return E_UNEXPECTED;
		
		INT length = lstrlenW(this->file_path);
		*file_path = (LPWSTR)CoTaskMemAlloc(length * 0x2 + 0x2);
		if (!(*file_path)) return E_OUTOFMEMORY;

		CopyMemory(*file_path, this->file_path, length * 0x2);
		return S_OK;
	}
	HRESULT __stdcall Load(LPCOLESTR file_path, DWORD load_mode) {
		UNREFERENCED_PARAMETER(load_mode);
		if (!file_path) return E_INVALIDARG;

		INT length = lstrlenW(file_path);
		this->file_path = (LPWSTR)HeapAlloc(GetProcessHeap(), 0x8, length * 0x2 + 0x2);
		if (!this->file_path) return E_OUTOFMEMORY;

		CopyMemory(this->file_path, file_path, length * 0x2);

		return S_OK;
	}
	HRESULT __stdcall IsDirty() {
		return E_NOTIMPL;
	}
	HRESULT __stdcall Save(LPCOLESTR file_path, BOOL b_remember) {
		UNREFERENCED_PARAMETER(file_path);
		UNREFERENCED_PARAMETER(b_remember);
		return E_NOTIMPL;
	}
	HRESULT __stdcall SaveCompleted(LPCOLESTR file_path) {
		UNREFERENCED_PARAMETER(file_path);
		return E_NOTIMPL;
	}
};
class CClassObject : public IClassFactory {
private:
	ULONG ref_count;
public:
	CClassObject() {
		InterlockedIncrement(&gRefCount);
		InterlockedIncrement(&ref_count);
		ref_count = 0x0;
	}
	~CClassObject() {
		InterlockedDecrement(&gRefCount);
	}

	// IUKnown
	ULONG __stdcall AddRef() {
		InterlockedIncrement(&ref_count);
		return ref_count;
	}
	ULONG __stdcall Release() {
		InterlockedDecrement(&ref_count);
		if (!ref_count) {
			delete this;
			return 0x0;
		}
		return ref_count;
	}
	HRESULT __stdcall QueryInterface(REFIID interface_id, LPVOID* interface_buffer) {
		if (!interface_buffer) return E_INVALIDARG;
		if (interface_id == IID_IUnknown) {
			*interface_buffer = (IUnknown*)this;
			AddRef();
			return S_OK;
		}
		if (interface_id == IID_IClassFactory) {
			*interface_buffer = (IClassFactory*)this;
			AddRef();
			return S_OK;
		}
		*interface_buffer = 0x0;
		return E_NOINTERFACE;
	}

	// IClassFactory
	HRESULT __stdcall CreateInstance(IUnknown* OuterIUKnown, REFIID interface_id, LPVOID* interface_buffer) {
		if (!interface_buffer) return E_INVALIDARG;
		if (OuterIUKnown) {
			*interface_buffer = 0x0;
			return CLASS_E_NOAGGREGATION;
		}
		if (interface_id == IID_IUnknown || interface_id == IID_IDataObject || interface_id == IID_IPersistFile) {
			CDataObejct* data_object = new CDataObejct();
			if (!data_object) {
				*interface_buffer = 0x0;
				return E_OUTOFMEMORY;
			}
			HRESULT rs = data_object->QueryInterface(interface_id, interface_buffer);
			if (FAILED(rs)) delete data_object;
			return rs;
		}
		*interface_buffer = 0x0;
		return E_NOINTERFACE;
	}
	HRESULT __stdcall LockServer(BOOL b_lock) {
		if (b_lock) AddRef();
		else Release();
		return S_OK;
	}
};

// ===========================================================

BOOL __stdcall DllMain(HMODULE dll_base, DWORD call_reason, LPVOID reserved) {
	UNREFERENCED_PARAMETER(reserved);
	UNREFERENCED_PARAMETER(dll_base);
	if (call_reason == DLL_PROCESS_ATTACH) return SUCCEEDED(CLSIDFromString(str_clsid, &ext_clsid));
	return 0x1;
}
HRESULT __stdcall DllGetClassObject(REFCLSID requested_clsid, REFIID interface_id, LPVOID* interface_buffer) {
	if (!interface_buffer) return E_INVALIDARG;
	if (requested_clsid != ext_clsid) return CLASS_E_CLASSNOTAVAILABLE;
	if (interface_id != IID_IClassFactory) return E_NOINTERFACE;

	CClassObject* class_object = new CClassObject();
	if (!class_object) return E_OUTOFMEMORY;

	HRESULT rs = class_object->QueryInterface(interface_id, interface_buffer);
	if (FAILED(rs)) delete class_object;

	return rs;
}
HRESULT __stdcall DllCanUnloadNow() {
	return gRefCount ? S_FALSE : S_OK;
}

// ===========================================================

#include "stdafx.h"
#include "tchar.h"
#include <metahost.h>

#import "mscorlib.tlb" \
    raw_interfaces_only \
    high_property_prefixes("_get", "_put", "_putref") \
    rename("ReportEvent", "InteropServices_ReportEvent") \
    exclude("ITrackingHandler")

#pragma comment(lib,"mscoree.lib")

using namespace mscorlib;

BOOL APIENTRY DllMain(HMODULE hModule, DWORD fdwReason, LPVOID lpReserved)
{
    return TRUE;
}

extern "C" __declspec(dllexport) HRESULT __stdcall LoadRuntime(ICorRuntimeHost* &pCorRuntimeHost)
{
    ICLRMetaHost* pMetaHost = nullptr;
    ICLRRuntimeInfo* pRuntimeInfo = nullptr;
    BOOL isLoadable = FALSE;

    HRESULT hr = CLRCreateInstance(CLSID_CLRMetaHost, IID_PPV_ARGS(&pMetaHost));
  
    if (FAILED(hr))
        goto Cleanup;

    hr = pMetaHost->GetRuntime(_T("v4.0.30319"), IID_PPV_ARGS(&pRuntimeInfo));

    if (FAILED(hr))
        goto Cleanup;

    hr = pRuntimeInfo->IsLoadable(&isLoadable);

    if (FAILED(hr))
        goto Cleanup;

    if (!isLoadable)
    {
        hr = E_FAIL;
        goto Cleanup;
    }
	
    hr = pRuntimeInfo->GetInterface(CLSID_CorRuntimeHost, IID_PPV_ARGS(&pCorRuntimeHost));

    Cleanup:

    if (pMetaHost)
    {
        pMetaHost->Release();
        pMetaHost = nullptr;
    }

    if (pRuntimeInfo)
    {
        pRuntimeInfo->Release();
        pRuntimeInfo = nullptr;
    }

    return hr;
}

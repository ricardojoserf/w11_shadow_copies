#include <iostream>
#include <vss.h>
#include <vswriter.h>
#include <vsbackup.h>
#include <string>
#pragma comment(lib, "vssapi.lib")


HRESULT create_shadow(const std::wstring& volumePath, std::wstring& outDeviceObject) {
    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hr)) {
        return hr;
    }

    IVssBackupComponents* pBackup = nullptr;
    hr = CreateVssBackupComponents(&pBackup);
    if (FAILED(hr) || !pBackup) {
        CoUninitialize();
        return hr;
    }

    hr = pBackup->InitializeForBackup();
    if (FAILED(hr)) {
        pBackup->Release();
        CoUninitialize();
        return hr;
    }

    BOOL bSupported = FALSE;
    hr = pBackup->IsVolumeSupported(GUID_NULL, (WCHAR*)volumePath.c_str(), &bSupported);
    if (SUCCEEDED(hr) && !bSupported) {
        pBackup->Release();
        CoUninitialize();
        return VSS_E_VOLUME_NOT_SUPPORTED;
    }

    hr = pBackup->SetContext(VSS_CTX_BACKUP);
    if (FAILED(hr)) {
        pBackup->Release();
        CoUninitialize();
        return hr;
    }

    hr = pBackup->SetBackupState(false, false, VSS_BT_FULL, false);

    IVssAsync* pAsyncMetadata = nullptr;
    hr = pBackup->GatherWriterMetadata(&pAsyncMetadata);
    if (SUCCEEDED(hr) && pAsyncMetadata) {
        hr = pAsyncMetadata->Wait();
        pAsyncMetadata->Release();
    }

    VSS_ID snapshotSetId;
    hr = pBackup->StartSnapshotSet(&snapshotSetId);
    if (FAILED(hr)) {
        pBackup->Release();
        CoUninitialize();
        return hr;
    }

    VSS_ID snapshotId;
    hr = pBackup->AddToSnapshotSet((WCHAR*)volumePath.c_str(), GUID_NULL, &snapshotId);
    if (FAILED(hr)) {
        pBackup->Release();
        CoUninitialize();
        return hr;
    }

    IVssAsync* pAsyncPrepare = nullptr;
    hr = pBackup->PrepareForBackup(&pAsyncPrepare);
    if (SUCCEEDED(hr) && pAsyncPrepare) {
        hr = pAsyncPrepare->Wait();
        pAsyncPrepare->Release();
    }

    IVssAsync* pAsyncSnapshot = nullptr;
    hr = pBackup->DoSnapshotSet(&pAsyncSnapshot);
    if (SUCCEEDED(hr) && pAsyncSnapshot) {
        hr = pAsyncSnapshot->Wait();
        pAsyncSnapshot->Release();
    }

    if (SUCCEEDED(hr)) {
        VSS_SNAPSHOT_PROP snapProp;
        hr = pBackup->GetSnapshotProperties(snapshotId, &snapProp);
        if (SUCCEEDED(hr)) {
            if (snapProp.m_pwszSnapshotDeviceObject) {
                outDeviceObject = snapProp.m_pwszSnapshotDeviceObject;
            }
            VssFreeSnapshotProperties(&snapProp);
        }
    }

    pBackup->Release();
    CoUninitialize();

    return hr;
}


int main() {
    std::wstring volumePath = L"C:\\";
    std::wstring deviceObject;

    HRESULT hr = create_shadow(volumePath, deviceObject);

    if (SUCCEEDED(hr) && !deviceObject.empty()) {
        std::wcout << L"Shadow copy created: " << deviceObject << std::endl;
    }
    else {
        std::cout << "Failed to create shadow copy" << std::endl;
    }

    return 0;
}
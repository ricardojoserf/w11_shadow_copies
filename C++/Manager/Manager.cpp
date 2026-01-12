#include <iostream>
#include <vss.h>
#include <vswriter.h>
#include <vsbackup.h>
#include <string>
#include <algorithm>
#pragma comment(lib, "vssapi.lib")


std::wstring GuidToWString(GUID id) {
    wchar_t buf[64];
    (void)StringFromGUID2(id, buf, 64);
    return std::wstring(buf);
}


void list_shadows() {
    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hr)) {
        std::cout << "CoInitializeEx failed" << std::endl;
        return;
    }

    IVssBackupComponents* pBackup = nullptr;
    hr = CreateVssBackupComponents(&pBackup);
    if (FAILED(hr) || !pBackup) {
        std::cout << "CreateVssBackupComponents failed" << std::endl;
        CoUninitialize();
        return;
    }

    hr = pBackup->InitializeForBackup();
    if (FAILED(hr)) {
        std::cout << "InitializeForBackup failed" << std::endl;
        pBackup->Release();
        CoUninitialize();
        return;
    }

    hr = pBackup->SetContext(VSS_CTX_ALL);
    if (FAILED(hr)) {
        hr = pBackup->SetContext(VSS_CTX_BACKUP);
    }

    IVssEnumObject* pEnum = nullptr;
    hr = pBackup->Query(GUID_NULL, VSS_OBJECT_NONE, VSS_OBJECT_SNAPSHOT, &pEnum);
    if (FAILED(hr) || !pEnum) {
        std::cout << "Query failed" << std::endl;
        pBackup->Release();
        CoUninitialize();
        return;
    }

    VSS_OBJECT_PROP prop = {};
    ULONG fetched = 0;
    int count = 0;

    while (true) {
        hr = pEnum->Next(1, &prop, &fetched);
        if (hr == S_FALSE || fetched == 0) break;
        if (FAILED(hr)) break;

        if (prop.Type == VSS_OBJECT_SNAPSHOT) {
            VSS_SNAPSHOT_PROP& snap = prop.Obj.Snap;
            count++;

            std::wcout << L"===============================================" << std::endl;
            std::wcout << L"Shadow Copy #" << count << std::endl;
            std::wcout << L"===============================================" << std::endl;
            std::wcout << L"ID: " << GuidToWString(snap.m_SnapshotId) << std::endl;
            std::wcout << L"Set ID: " << GuidToWString(snap.m_SnapshotSetId) << std::endl;

            if (snap.m_pwszSnapshotDeviceObject) {
                std::wcout << L"Device Object: " << snap.m_pwszSnapshotDeviceObject << std::endl;
            }

            if (snap.m_pwszOriginalVolumeName) {
                std::wcout << L"Original Volume: " << snap.m_pwszOriginalVolumeName << std::endl;
            }

            if (snap.m_pwszOriginatingMachine) {
                std::wcout << L"Originating Machine: " << snap.m_pwszOriginatingMachine << std::endl;
            }

            SYSTEMTIME st;
            FILETIME ftUTC, ftLocal;
            ftUTC.dwLowDateTime = (DWORD)(snap.m_tsCreationTimestamp & 0xFFFFFFFF);
            ftUTC.dwHighDateTime = (DWORD)(snap.m_tsCreationTimestamp >> 32);
            FileTimeToLocalFileTime(&ftUTC, &ftLocal);
            FileTimeToSystemTime(&ftLocal, &st);
            wprintf(L"Creation Date: %02d/%02d/%04d %02d:%02d:%02d\n",
                st.wDay, st.wMonth, st.wYear, st.wHour, st.wMinute, st.wSecond);

            wprintf(L"Attributes: 0x%lx\n", snap.m_lSnapshotAttributes);
            wprintf(L"Status: %d\n", snap.m_eStatus);
            std::wcout << L"Provider ID: " << GuidToWString(snap.m_ProviderId) << std::endl;
            std::wcout << L"\n";
            VssFreeSnapshotProperties(&snap);
        }
    }

    if (count == 0) {
        std::cout << "No shadow copies found" << std::endl;
    }
    else {
        std::cout << "Total: " << count << " shadow copies found" << std::endl;
    }

    pEnum->Release();
    pBackup->Release();
    CoUninitialize();
}


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


GUID ExtractSnapshotIdFromPath(const std::wstring& devicePath) {
    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hr)) {
        return GUID_NULL;
    }

    IVssBackupComponents* pBackup = nullptr;
    hr = CreateVssBackupComponents(&pBackup);
    if (FAILED(hr) || !pBackup) {
        CoUninitialize();
        return GUID_NULL;
    }

    hr = pBackup->InitializeForBackup();
    if (FAILED(hr)) {
        pBackup->Release();
        CoUninitialize();
        return GUID_NULL;
    }

    hr = pBackup->SetContext(VSS_CTX_BACKUP);
    if (FAILED(hr)) {
        pBackup->Release();
        CoUninitialize();
        return GUID_NULL;
    }

    IVssEnumObject* pEnum = nullptr;
    hr = pBackup->Query(GUID_NULL, VSS_OBJECT_NONE, VSS_OBJECT_SNAPSHOT, &pEnum);
    if (FAILED(hr) || !pEnum) {
        pBackup->Release();
        CoUninitialize();
        return GUID_NULL;
    }

    VSS_OBJECT_PROP prop = {};
    ULONG fetched = 0;
    GUID foundId = GUID_NULL;

    while (true) {
        hr = pEnum->Next(1, &prop, &fetched);
        if (hr == S_FALSE || fetched == 0) break;
        if (FAILED(hr)) break;

        if (prop.Type == VSS_OBJECT_SNAPSHOT) {
            VSS_SNAPSHOT_PROP& snap = prop.Obj.Snap;

            if (snap.m_pwszSnapshotDeviceObject) {
                std::wstring deviceObject = snap.m_pwszSnapshotDeviceObject;

                if (_wcsicmp(deviceObject.c_str(), devicePath.c_str()) == 0) {
                    foundId = snap.m_SnapshotId;
                    VssFreeSnapshotProperties(&snap);
                    break;
                }
            }
            VssFreeSnapshotProperties(&snap);
        }
    }

    pEnum->Release();
    pBackup->Release();
    CoUninitialize();

    return foundId;
}


bool delete_shadow(const std::wstring& devicePath) {
    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hr)) {
        return false;
    }

    IVssBackupComponents* pBackup = nullptr;
    hr = CreateVssBackupComponents(&pBackup);
    if (FAILED(hr) || !pBackup) {
        CoUninitialize();
        return false;
    }

    hr = pBackup->InitializeForBackup();
    if (FAILED(hr)) {
        pBackup->Release();
        CoUninitialize();
        return false;
    }

    hr = pBackup->SetContext(VSS_CTX_BACKUP);
    if (FAILED(hr)) {
        pBackup->Release();
        CoUninitialize();
        return false;
    }

    GUID snapshotId = ExtractSnapshotIdFromPath(devicePath);
    if (snapshotId == GUID_NULL) {
        std::cout << "Failed to extract snapshot ID from path" << std::endl;
        pBackup->Release();
        CoUninitialize();
        return false;
    }

    LONG deletedSnapshots = 0;
    VSS_ID nondeletedSnapshotID = GUID_NULL;

    pBackup->DeleteSnapshots(snapshotId, VSS_OBJECT_SNAPSHOT, TRUE, &deletedSnapshots, &nondeletedSnapshotID);

    bool success = (deletedSnapshots > 0);

    pBackup->Release();
    CoUninitialize();

    return success;
}


void print_usage() {
    std::cout << "Usage: program.exe <command> [args]" << std::endl;
    std::cout << "Commands:" << std::endl;
    std::cout << "  create              - Create shadow copy of C:\\" << std::endl;
    std::cout << "  delete <path>       - Delete shadow copy" << std::endl;
    std::cout << "  list                - List all shadow copies" << std::endl;
}


int main(int argc, char* argv[]) {
    if (argc < 2) {
        print_usage();
        return 1;
    }

    std::string commandStr = argv[1];
    std::transform(commandStr.begin(), commandStr.end(), commandStr.begin(), ::tolower);

    if (commandStr == "create") {
        std::wstring volumePath = L"C:\\";
        std::wstring deviceObject;

        HRESULT hr = create_shadow(volumePath, deviceObject);

        if (SUCCEEDED(hr) && !deviceObject.empty()) {
            std::wcout << L"Shadow copy created: " << deviceObject << std::endl;
        }
        else {
            std::cout << "Failed to create shadow copy" << std::endl;
        }
    }
    else if (commandStr == "delete") {
        if (argc < 3) {
            std::cout << "ERROR: Delete command requires device path" << std::endl;
            std::cout << "Example: program.exe delete \\\\?\\GLOBALROOT\\Device\\HarddiskVolumeShadowCopy8" << std::endl;
            return 1;
        }

        std::string devicePathStr = argv[2];
        std::wstring devicePath(devicePathStr.begin(), devicePathStr.end());

        // std::wcout << L"Deleting shadow copy: " << devicePath << std::endl;

        if (delete_shadow(devicePath)) {
            std::cout << "Shadow copy deleted successfully" << std::endl;
        }
        else {
            std::cout << "Failed to delete shadow copy" << std::endl;
        }
    }
    else if (commandStr == "list") {
        list_shadows();
    }
    else {
        std::cout << "Unknown command: " << commandStr << std::endl;
        std::cout << "Valid commands: create, delete, list" << std::endl;
        return 1;
    }

    return 0;
}
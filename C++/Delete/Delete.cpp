#include <iostream>
#include <vss.h>
#include <vswriter.h>
#include <vsbackup.h>
#include <string>
#pragma comment(lib, "vssapi.lib")


std::wstring GuidToWString(GUID id) {
    wchar_t buf[64];
    (void)StringFromGUID2(id, buf, 64);
    return std::wstring(buf);
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


bool DeleteShadowCopy(const std::wstring& devicePath) {
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


int main(int argc, char* argv[]) {
    if (argc < 2) {
        std::cout << "Usage: program.exe <device_path>" << std::endl;
        std::cout << "Example: program.exe \\\\?\\GLOBALROOT\\Device\\HarddiskVolumeShadowCopy8" << std::endl;
        return 1;
    }

    std::string devicePathStr = argv[1];
    std::wstring devicePath(devicePathStr.begin(), devicePathStr.end());

    // std::wcout << L"Deleting shadow copy: " << devicePath << std::endl;

    if (DeleteShadowCopy(devicePath)) {
        std::cout << "Shadow copy deleted successfully" << std::endl;
    }
    else {
        std::cout << "Failed to delete shadow copy" << std::endl;
    }

    return 0;
}
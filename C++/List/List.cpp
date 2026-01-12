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


int main() {
    list_shadows();
    return 0;
}
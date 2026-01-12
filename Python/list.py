import sys
import ctypes
from ctypes import *
from ctypes.wintypes import *
from comtypes import GUID, IUnknown, COMMETHOD, HRESULT, BSTR
from comtypes import CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED
import comtypes.client
from datetime import datetime


# Constants
VSS_CTX_BACKUP = 0
VSS_CTX_ALL = 0xffffffff
GUID_NULL = GUID()

# Enums
class VSS_OBJECT_TYPE(ctypes.c_int):
    VSS_OBJECT_UNKNOWN = 0
    VSS_OBJECT_NONE = 1
    VSS_OBJECT_SNAPSHOT_SET = 2
    VSS_OBJECT_SNAPSHOT = 3
    VSS_OBJECT_PROVIDER = 4
    VSS_OBJECT_TYPE_COUNT = 5

# Structs
class VSS_SNAPSHOT_PROP(Structure):
    _pack_ = 8
    _fields_ = [
        ("m_SnapshotId", GUID),
        ("m_SnapshotSetId", GUID),
        ("m_lSnapshotsCount", LONG),
        ("m_pwszSnapshotDeviceObject", LPWSTR),
        ("m_pwszOriginalVolumeName", LPWSTR),
        ("m_pwszOriginatingMachine", LPWSTR),
        ("m_pwszServiceMachine", LPWSTR),
        ("m_pwszExposedName", LPWSTR),
        ("m_pwszExposedPath", LPWSTR),
        ("m_ProviderId", GUID),
        ("m_lSnapshotAttributes", LONG),
        ("m_tsCreationTimestamp", c_longlong),
        ("m_eStatus", LONG),
    ]

class VSS_OBJECT_UNION(Union):
    _fields_ = [
        ("Snap", VSS_SNAPSHOT_PROP),
    ]

class VSS_OBJECT_PROP(Structure):
    _fields_ = [
        ("Type", ctypes.c_int),
        ("Obj", VSS_OBJECT_UNION),
    ]

# Interfaces
class IVssEnumObject(IUnknown):
    _iid_ = GUID("{AE1C7110-2F60-11d3-8A39-00C04F72D8E3}")
    _methods_ = [
        COMMETHOD([], HRESULT, 'Next',
                  (['in'], ULONG, 'celt'),
                  (['out'], POINTER(VSS_OBJECT_PROP), 'rgelt'),
                  (['out'], POINTER(ULONG), 'pceltFetched')),
        COMMETHOD([], HRESULT, 'Skip',
                  (['in'], ULONG, 'celt')),
        COMMETHOD([], HRESULT, 'Reset'),
        COMMETHOD([], HRESULT, 'Clone',
                  (['out'], POINTER(POINTER(IUnknown)), 'ppenum')),
    ]

class IVssAsync(IUnknown):
    _iid_ = GUID("{507C37B4-CF5B-4e95-B0AF-14EB9767467E}")
    _methods_ = [
        COMMETHOD([], HRESULT, 'Cancel'),
        COMMETHOD([], HRESULT, 'Wait',
                  (['in'], DWORD, 'dwMilliseconds')),
        COMMETHOD([], HRESULT, 'QueryStatus',
                  (['out'], POINTER(HRESULT), 'pHrResult'),
                  (['out'], POINTER(INT), 'pReserved')),
    ]

class IVssBackupComponents(IUnknown):
    _iid_ = GUID("{665c1d5f-c218-414d-a05d-7fef5f9d5c86}")
    _methods_ = [
        COMMETHOD([], HRESULT, 'GetWriterComponentsCount',
                  (['out'], POINTER(UINT), 'pcComponents')),
        COMMETHOD([], HRESULT, 'GetWriterComponents',
                  (['in'], UINT, 'iWriter'),
                  (['out'], POINTER(POINTER(IUnknown)), 'ppWriter')),
        COMMETHOD([], HRESULT, 'InitializeForBackup',
                  (['in'], BSTR, 'bstrXML')),
        COMMETHOD([], HRESULT, 'SetBackupState',
                  (['in'], BOOL, 'bSelectComponents'),
                  (['in'], BOOL, 'bBackupBootableSystemState'),
                  (['in'], ctypes.c_int, 'backupType'),
                  (['in'], BOOL, 'bPartialFileSupport')),
        COMMETHOD([], HRESULT, 'InitializeForRestore',
                  (['in'], BSTR, 'bstrXML')),
        COMMETHOD([], HRESULT, 'SetRestoreState',
                  (['in'], ctypes.c_int, 'restoreType')),
        COMMETHOD([], HRESULT, 'GatherWriterMetadata',
                  (['out'], POINTER(POINTER(IVssAsync)), 'ppAsync')),
        COMMETHOD([], HRESULT, 'GetWriterMetadataCount',
                  (['out'], POINTER(UINT), 'pcWriters')),
        COMMETHOD([], HRESULT, 'GetWriterMetadata',
                  (['in'], UINT, 'iWriter'),
                  (['out'], POINTER(GUID), 'pInstanceId'),
                  (['out'], POINTER(POINTER(IUnknown)), 'ppMetadata')),
        COMMETHOD([], HRESULT, 'FreeWriterMetadata'),
        COMMETHOD([], HRESULT, 'AddComponent',
                  (['in'], POINTER(GUID), 'instanceId'),
                  (['in'], POINTER(GUID), 'writerId'),
                  (['in'], ctypes.c_int, 'ct'),
                  (['in'], LPCWSTR, 'wszLogicalPath'),
                  (['in'], LPCWSTR, 'wszComponentName')),
        COMMETHOD([], HRESULT, 'PrepareForBackup',
                  (['out'], POINTER(POINTER(IVssAsync)), 'ppAsync')),
        COMMETHOD([], HRESULT, 'AbortBackup'),
        COMMETHOD([], HRESULT, 'GatherWriterStatus',
                  (['out'], POINTER(POINTER(IVssAsync)), 'pAsync')),
        COMMETHOD([], HRESULT, 'GetWriterStatusCount',
                  (['out'], POINTER(UINT), 'pcWriters')),
        COMMETHOD([], HRESULT, 'FreeWriterStatus'),
        COMMETHOD([], HRESULT, 'GetWriterStatus',
                  (['in'], UINT, 'iWriter'),
                  (['out'], POINTER(GUID), 'pidInstance'),
                  (['out'], POINTER(GUID), 'pidWriter'),
                  (['out'], POINTER(BSTR), 'pbstrWriter'),
                  (['out'], POINTER(ctypes.c_int), 'pnStatus'),
                  (['out'], POINTER(ctypes.c_int), 'phrFailureWriter'),
                  (['out'], POINTER(ctypes.c_int), 'phrApplication'),
                  (['out'], POINTER(BSTR), 'pbstrApplicationMessage')),
        COMMETHOD([], HRESULT, 'SetBackupSucceeded',
                  (['in'], POINTER(GUID), 'instanceId'),
                  (['in'], POINTER(GUID), 'writerId'),
                  (['in'], ctypes.c_int, 'ct'),
                  (['in'], LPCWSTR, 'wszLogicalPath'),
                  (['in'], LPCWSTR, 'wszComponentName'),
                  (['in'], BOOL, 'bSucceeded')),
        COMMETHOD([], HRESULT, 'SetBackupOptions',
                  (['in'], POINTER(GUID), 'writerId'),
                  (['in'], ctypes.c_int, 'ct'),
                  (['in'], LPCWSTR, 'wszLogicalPath'),
                  (['in'], LPCWSTR, 'wszComponentName'),
                  (['in'], LPCWSTR, 'wszBackupOptions')),
        COMMETHOD([], HRESULT, 'SetSelectedForRestore',
                  (['in'], POINTER(GUID), 'writerId'),
                  (['in'], ctypes.c_int, 'ct'),
                  (['in'], LPCWSTR, 'wszLogicalPath'),
                  (['in'], LPCWSTR, 'wszComponentName'),
                  (['in'], BOOL, 'bSelectedForRestore')),
        COMMETHOD([], HRESULT, 'SetRestoreOptions',
                  (['in'], POINTER(GUID), 'writerId'),
                  (['in'], ctypes.c_int, 'ct'),
                  (['in'], LPCWSTR, 'wszLogicalPath'),
                  (['in'], LPCWSTR, 'wszComponentName'),
                  (['in'], LPCWSTR, 'wszRestoreOptions')),
        COMMETHOD([], HRESULT, 'SetAdditionalRestores',
                  (['in'], POINTER(GUID), 'writerId'),
                  (['in'], ctypes.c_int, 'ct'),
                  (['in'], LPCWSTR, 'wszLogicalPath'),
                  (['in'], LPCWSTR, 'wszComponentName'),
                  (['in'], BOOL, 'bAdditionalRestores')),
        COMMETHOD([], HRESULT, 'SetPreviousBackupStamp',
                  (['in'], POINTER(GUID), 'writerId'),
                  (['in'], ctypes.c_int, 'ct'),
                  (['in'], LPCWSTR, 'wszLogicalPath'),
                  (['in'], LPCWSTR, 'wszComponentName'),
                  (['in'], LPCWSTR, 'wszPreviousBackupStamp')),
        COMMETHOD([], HRESULT, 'SaveAsXML',
                  (['out'], POINTER(BSTR), 'pbstrXML')),
        COMMETHOD([], HRESULT, 'BackupComplete',
                  (['out'], POINTER(POINTER(IVssAsync)), 'ppAsync')),
        COMMETHOD([], HRESULT, 'AddAlternativeLocationMapping',
                  (['in'], POINTER(GUID), 'writerId'),
                  (['in'], ctypes.c_int, 'ct'),
                  (['in'], LPCWSTR, 'wszLogicalPath'),
                  (['in'], LPCWSTR, 'wszComponentName'),
                  (['in'], LPCWSTR, 'wszPath'),
                  (['in'], LPCWSTR, 'wszFilespec'),
                  (['in'], BOOL, 'bRecursive'),
                  (['in'], LPCWSTR, 'wszDestination')),
        COMMETHOD([], HRESULT, 'AddRestoreSubcomponent',
                  (['in'], POINTER(GUID), 'writerId'),
                  (['in'], ctypes.c_int, 'ct'),
                  (['in'], LPCWSTR, 'wszLogicalPath'),
                  (['in'], LPCWSTR, 'wszComponentName'),
                  (['in'], LPCWSTR, 'wszSubComponentLogicalPath'),
                  (['in'], LPCWSTR, 'wszSubComponentName'),
                  (['in'], BOOL, 'bRepair')),
        COMMETHOD([], HRESULT, 'SetFileRestoreStatus',
                  (['in'], POINTER(GUID), 'writerId'),
                  (['in'], ctypes.c_int, 'ct'),
                  (['in'], LPCWSTR, 'wszLogicalPath'),
                  (['in'], LPCWSTR, 'wszComponentName'),
                  (['in'], ctypes.c_int, 'status')),
        COMMETHOD([], HRESULT, 'AddNewTarget',
                  (['in'], POINTER(GUID), 'writerId'),
                  (['in'], ctypes.c_int, 'ct'),
                  (['in'], LPCWSTR, 'wszLogicalPath'),
                  (['in'], LPCWSTR, 'wszComponentName'),
                  (['in'], LPCWSTR, 'wszPath'),
                  (['in'], LPCWSTR, 'wszFileName'),
                  (['in'], BOOL, 'bRecursive'),
                  (['in'], LPCWSTR, 'wszAlternatePath')),
        COMMETHOD([], HRESULT, 'SetRangesFilePath',
                  (['in'], POINTER(GUID), 'writerId'),
                  (['in'], ctypes.c_int, 'ct'),
                  (['in'], LPCWSTR, 'wszLogicalPath'),
                  (['in'], LPCWSTR, 'wszComponentName'),
                  (['in'], UINT, 'iPartialFile'),
                  (['in'], LPCWSTR, 'wszRangesFile')),
        COMMETHOD([], HRESULT, 'PreRestore',
                  (['out'], POINTER(POINTER(IVssAsync)), 'ppAsync')),
        COMMETHOD([], HRESULT, 'PostRestore',
                  (['out'], POINTER(POINTER(IVssAsync)), 'ppAsync')),
        COMMETHOD([], HRESULT, 'SetContext',
                  (['in'], LONG, 'lContext')),
        COMMETHOD([], HRESULT, 'StartSnapshotSet',
                  (['out'], POINTER(GUID), 'pSnapshotSetId')),
        COMMETHOD([], HRESULT, 'AddToSnapshotSet',
                  (['in'], LPCWSTR, 'pwszVolumeName'),
                  (['in'], POINTER(GUID), 'ProviderId'),
                  (['out'], POINTER(GUID), 'pidSnapshot')),
        COMMETHOD([], HRESULT, 'DoSnapshotSet',
                  (['out'], POINTER(POINTER(IVssAsync)), 'ppAsync')),
        COMMETHOD([], HRESULT, 'DeleteSnapshots',
                  (['in'], GUID, 'SourceObjectId'),
                  (['in'], ctypes.c_int, 'eSourceObjectType'),
                  (['in'], BOOL, 'bForceDelete'),
                  (['out'], POINTER(LONG), 'plDeletedSnapshots'),
                  (['out'], POINTER(GUID), 'pNondeletedSnapshotID')),
        COMMETHOD([], HRESULT, 'ImportSnapshots',
                  (['out'], POINTER(POINTER(IVssAsync)), 'ppAsync')),
        COMMETHOD([], HRESULT, 'BreakSnapshotSet',
                  (['in'], POINTER(GUID), 'SnapshotSetId')),
        COMMETHOD([], HRESULT, 'GetSnapshotProperties',
                  (['in'], POINTER(GUID), 'SnapshotId'),
                  (['out'], POINTER(VSS_SNAPSHOT_PROP), 'pProp')),
        COMMETHOD([], HRESULT, 'Query',
                  (['in'], POINTER(GUID), 'QueriedObjectId'),
                  (['in'], ctypes.c_int, 'eQueriedObjectType'),
                  (['in'], ctypes.c_int, 'eReturnedObjectsType'),
                  (['out'], POINTER(POINTER(IVssEnumObject)), 'ppEnum')),
        COMMETHOD([], HRESULT, 'IsVolumeSupported',
                  (['in'], POINTER(GUID), 'ProviderId'),
                  (['in'], LPCWSTR, 'pwszVolumeName'),
                  (['out'], POINTER(BOOL), 'pbSupportedByThisProvider')),
        COMMETHOD([], HRESULT, 'DisableWriterClasses',
                  (['in'], POINTER(GUID), 'rgWriterClassId'),
                  (['in'], UINT, 'cClassId')),
        COMMETHOD([], HRESULT, 'EnableWriterClasses',
                  (['in'], POINTER(GUID), 'rgWriterClassId'),
                  (['in'], UINT, 'cClassId')),
        COMMETHOD([], HRESULT, 'DisableWriterInstances',
                  (['in'], POINTER(GUID), 'rgWriterInstanceId'),
                  (['in'], UINT, 'cInstanceId')),
        COMMETHOD([], HRESULT, 'ExposeSnapshot',
                  (['in'], POINTER(GUID), 'SnapshotId'),
                  (['in'], LPCWSTR, 'wszPathFromRoot'),
                  (['in'], LONG, 'lAttributes'),
                  (['in'], LPCWSTR, 'wszExpose'),
                  (['out'], POINTER(LPWSTR), 'pwszExposed')),
        COMMETHOD([], HRESULT, 'RevertToSnapshot',
                  (['in'], POINTER(GUID), 'SnapshotId'),
                  (['in'], BOOL, 'bForceDismount')),
        COMMETHOD([], HRESULT, 'QueryRevertStatus',
                  (['in'], LPCWSTR, 'pwszVolume'),
                  (['out'], POINTER(POINTER(IVssAsync)), 'ppAsync')),
    ]


try:
    vssapi = ctypes.WinDLL("VssApi.dll")
except OSError as e:
    print(f"Error: Failed to load VssApi.dll: {e}")
    sys.exit(1)

VssFreeSnapshotProperties = vssapi.VssFreeSnapshotProperties
VssFreeSnapshotProperties.argtypes = [POINTER(VSS_SNAPSHOT_PROP)]
VssFreeSnapshotProperties.restype = None


def is_administrator():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except:
        return False


def create_vss_backup_components():
    possible_names = [
        "CreateVssBackupComponentsInternal",
        "CreateVssBackupComponents",
        "?CreateVssBackupComponents@@YAJPEAPEAVIVssBackupComponents@@@Z",
    ]
    
    create_func = None
    for name in possible_names:
        try:
            create_func = getattr(vssapi, name)
            break
        except AttributeError:
            continue
    
    if create_func is None:
        raise Exception("Could not find CreateVssBackupComponents in VssApi.dll")
    
    create_func.restype = HRESULT
    create_func.argtypes = [POINTER(POINTER(IVssBackupComponents))]
    
    backup = POINTER(IVssBackupComponents)()
    hr = create_func(byref(backup))
    
    if hr != 0:
        raise Exception(f"CreateVssBackupComponents failed with HRESULT: 0x{hr:08X}")
    
    if not backup:
        raise Exception("CreateVssBackupComponents returned NULL")
    
    return backup


def list_shadow_copies():
    com_initialized = False
    try:
        CoInitializeEx(COINIT_MULTITHREADED)
        com_initialized = True
    except OSError as e:
        if e.winerror == -2147417850:
            pass
        elif e.winerror not in (0, 1):
            print(f"Error initializing COM. Error: {e}")
            return
    
    backup = None
    enum_obj = None
    
    try:
        backup = create_vss_backup_components()
        if not backup:
            print("Error creating VSS components.")
            return
        
        hr = backup.InitializeForBackup(None)
        if hr != 0:
            print(f"Error in InitializeForBackup. HRESULT: 0x{hr:08X}")
            if hr == 0x80042302:
                print("  -> VSS_E_UNEXPECTED: Unexpected VSS error")
            elif hr == 0x8004230C:
                print("  -> VSS_E_BAD_STATE: VSS in incorrect state")
            elif hr == 0x80042308:
                print("  -> VSS_E_VOLUME_NOT_SUPPORTED_BY_PROVIDER: Volume not supported")
            return
        
        hr = backup.SetContext(VSS_CTX_ALL)
        if hr != 0:
            hr = backup.SetContext(VSS_CTX_BACKUP)
            if hr != 0:
                print(f"Error in SetContext. HRESULT: 0x{hr:08X}")
                return
        
        guid_null = GUID()
        
        enum_obj_ptr = backup.Query(
            guid_null,
            VSS_OBJECT_TYPE.VSS_OBJECT_NONE,
            VSS_OBJECT_TYPE.VSS_OBJECT_SNAPSHOT
        )
        
        if not enum_obj_ptr:
            print("No shadow copies found on the system.")
            return
        
        enum_obj = enum_obj_ptr
        
        count = 0
        while True:
            try:
                result = enum_obj.Next(1)
                if not result or len(result) < 2:
                    break
                    
                prop, fetched = result
                
                if fetched == 0:
                    break
                
            except Exception as e:
                break
            
            if prop.Type == VSS_OBJECT_TYPE.VSS_OBJECT_SNAPSHOT:
                count += 1
                print("=" * 47)
                print(f"Shadow Copy #{count}")
                print("=" * 47)
                
                snap = prop.Obj.Snap
                
                print(f"ID: {{{snap.m_SnapshotId}}}")
                print(f"Set ID: {{{snap.m_SnapshotSetId}}}")
                
                if snap.m_pwszSnapshotDeviceObject:
                    device_object = ctypes.wstring_at(snap.m_pwszSnapshotDeviceObject)
                    print(f"Device Object: {device_object}")
                
                if snap.m_pwszOriginalVolumeName:
                    original_volume = ctypes.wstring_at(snap.m_pwszOriginalVolumeName)
                    print(f"Original Volume: {original_volume}")
                
                if snap.m_pwszOriginatingMachine:
                    machine = ctypes.wstring_at(snap.m_pwszOriginatingMachine)
                    print(f"Originating Machine: {machine}")
                
                timestamp = datetime.fromtimestamp(snap.m_tsCreationTimestamp / 10000000 - 11644473600)
                print(f"Creation Date: {timestamp}")
                print(f"Attributes: 0x{snap.m_lSnapshotAttributes:X}")
                print(f"Status: {snap.m_eStatus}")
                print(f"Provider ID: {{{snap.m_ProviderId}}}")
                print()
                
                VssFreeSnapshotProperties(byref(snap))
        
        if count == 0:
            print("No shadow copies found on the system.")
        else:
            print(f"Total: {count} shadow copies found")
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if enum_obj:
            del enum_obj
        if backup:
            del backup
        
        if com_initialized:
            CoUninitialize()


def main():
    if not is_administrator():
        print("ERROR: This program requires administrator privileges.")
        return
    
    list_shadow_copies()


if __name__ == "__main__":
    main()
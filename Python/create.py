import sys
import ctypes
from ctypes import *
from ctypes.wintypes import *
from comtypes import GUID, IUnknown, COMMETHOD, HRESULT, BSTR
from comtypes import CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED
import comtypes.client


# Constants
VSS_CTX_BACKUP = 0
VSS_BT_FULL = 1
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


def create_shadow_copy(volume_path):
    com_initialized = False
    try:
        CoInitializeEx(COINIT_MULTITHREADED)
        com_initialized = True
    except OSError as e:
        if e.winerror == -2147417850:
            pass
        elif e.winerror not in (0, 1):
            print(f"Error initializing COM. Error: {e}")
            return None, None
    
    backup = None
    
    try:
        backup = create_vss_backup_components()
        if not backup:
            print("Error creating VSS components.")
            return None, None

        hr = backup.InitializeForBackup(None)
        if hr != 0:
            print(f"Error in InitializeForBackup. HRESULT: 0x{hr:08X}")
            return None, None
        
        guid_null = GUID()
        try:
            result = backup.IsVolumeSupported(guid_null, volume_path)
            if isinstance(result, tuple):
                b_supported = result[0] if len(result) > 0 else False
            else:
                b_supported = result
            if not b_supported:
                print(f"Volume {volume_path} is not supported for shadow copies.")
                return None, None
        except Exception as e:
            print(f"[DEBUG] IsVolumeSupported exception: {e}")
        
        hr = backup.SetContext(VSS_CTX_BACKUP)
        if hr != 0:
            print(f"Error in SetContext. HRESULT: 0x{hr:08X}")
            return None, None
        
        hr = backup.SetBackupState(False, False, VSS_BT_FULL, False)
        
        try:
            async_metadata = backup.GatherWriterMetadata()
            if async_metadata:
                async_metadata.Wait(0xFFFFFFFF)
        except Exception as e:
            print(f"[DEBUG] GatherWriterMetadata exception: {e}")
        
        try:
            snapshot_set_id = backup.StartSnapshotSet()
            if not snapshot_set_id:
                print("Error in StartSnapshotSet.")
                return None, None
        except Exception as e:
            print(f"[DEBUG] StartSnapshotSet exception: {e}")
            return None, None
        
        try:
            snapshot_id = backup.AddToSnapshotSet(volume_path, guid_null)
            if not snapshot_id:
                print("Error in AddToSnapshotSet.")
                return None, None
        except Exception as e:
            print(f"[DEBUG] AddToSnapshotSet exception: {e}")
            return None, None
        
        try:
            async_prepare = backup.PrepareForBackup()
            if async_prepare:
                async_prepare.Wait(0xFFFFFFFF)
        except Exception as e:
            print(f"[DEBUG] PrepareForBackup exception: {e}")
        
        try:
            async_snapshot = backup.DoSnapshotSet()
            if async_snapshot:
                hr = async_snapshot.Wait(0xFFFFFFFF)
        except Exception as e:
            print(f"[DEBUG] DoSnapshotSet exception: {e}")
            hr = -1
        
        if hr == 0:
            try:
                snap_prop = backup.GetSnapshotProperties(snapshot_id)
                if snap_prop:
                    device_object = None
                    if snap_prop.m_pwszSnapshotDeviceObject:
                        device_object = ctypes.wstring_at(snap_prop.m_pwszSnapshotDeviceObject)
                    
                    snapshot_id_str = str(snapshot_id)
                    VssFreeSnapshotProperties(byref(snap_prop))
                    return snapshot_id_str, device_object
            except Exception as e:
                print(f"[DEBUG] GetSnapshotProperties exception: {e}")
                import traceback
                traceback.print_exc()
        
        return None, None
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return None, None
    finally:
        if backup:
            del backup
        
        if com_initialized:
            CoUninitialize()


def main():
    if not is_administrator():
        print("ERROR: Administrator privileges required")
        return
    
    volume_path = r"C:\\"
    
    if len(sys.argv) > 1:
        volume_path = sys.argv[1]
        if not volume_path.endswith("\\"):
            volume_path += "\\"
    
    snapshot_id, device_object = create_shadow_copy(volume_path)
    
    if snapshot_id and device_object:
        print(f"Shadow copy created: {device_object}")
    else:
        print("Failed to create shadow copy")


if __name__ == "__main__":
    main()
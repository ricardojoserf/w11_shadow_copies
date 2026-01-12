import sys
import ctypes
from ctypes import *
from ctypes.wintypes import *
from comtypes import GUID, IUnknown, COMMETHOD, HRESULT, BSTR
from comtypes import CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED
import comtypes.client


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


def extract_snapshot_id_from_path(device_path):
    com_initialized = False
    try:
        CoInitializeEx(COINIT_MULTITHREADED)
        com_initialized = True
    except OSError as e:
        if e.winerror == -2147417850:
            pass
        elif e.winerror not in (0, 1):
            print(f"Error initializing COM. Error: {e}")
            return None
    
    backup = None
    enum_obj = None
    
    try:
        backup = create_vss_backup_components()
        if not backup:
            print("Error creating VSS components.")
            return None
        
        hr = backup.InitializeForBackup(None)
        if hr != 0:
            print(f"Error in InitializeForBackup. HRESULT: 0x{hr:08X}")
            return None
        
        hr = backup.SetContext(VSS_CTX_ALL)
        if hr != 0:
            hr = backup.SetContext(VSS_CTX_BACKUP)
            if hr != 0:
                print(f"Error in SetContext. HRESULT: 0x{hr:08X}")
                return None
        
        guid_null = GUID()
        
        enum_obj_ptr = backup.Query(
            guid_null,
            VSS_OBJECT_TYPE.VSS_OBJECT_NONE,
            VSS_OBJECT_TYPE.VSS_OBJECT_SNAPSHOT
        )
        
        if not enum_obj_ptr:
            print("No shadow copies found on the system.")
            return None
        
        enum_obj = enum_obj_ptr
        
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
                snap = prop.Obj.Snap
                
                if snap.m_pwszSnapshotDeviceObject:
                    device_object = ctypes.wstring_at(snap.m_pwszSnapshotDeviceObject)
                    
                    device_object_norm = device_object.rstrip('\\').lower()
                    device_path_norm = device_path.rstrip('\\').lower()
                    
                    if device_object_norm == device_path_norm:
                        snapshot_id = snap.m_SnapshotId
                        VssFreeSnapshotProperties(byref(snap))
                        return snapshot_id
                
                VssFreeSnapshotProperties(byref(snap))
        
        print(f"No shadow copy found with device path: {device_path}")
        return None
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return None
    finally:
        if enum_obj:
            del enum_obj
        if backup:
            del backup
        
        if com_initialized:
            CoUninitialize()


def delete_shadow_copy(device_path):
    com_initialized = False
    try:
        CoInitializeEx(COINIT_MULTITHREADED)
        com_initialized = True
    except OSError as e:
        if e.winerror == -2147417850:
            pass
        elif e.winerror not in (0, 1):
            print(f"Error initializing COM. Error: {e}")
            return False
    
    backup = None
    
    try:
        snapshot_id = extract_snapshot_id_from_path(device_path)
        if not snapshot_id:
            return False
                
        backup = create_vss_backup_components()
        if not backup:
            print("Error creating VSS components.")
            return False
        
        hr = backup.InitializeForBackup(None)
        if hr != 0:
            print(f"Error in InitializeForBackup. HRESULT: 0x{hr:08X}")
            return False
        
        hr = backup.SetContext(VSS_CTX_ALL)
        if hr != 0:
            hr = backup.SetContext(VSS_CTX_BACKUP)
            if hr != 0:
                print(f"Error in SetContext. HRESULT: 0x{hr:08X}")
                return False
        
        try:
            result = backup.DeleteSnapshots(
                snapshot_id,
                VSS_OBJECT_TYPE.VSS_OBJECT_SNAPSHOT,
                True
            )
            
            if isinstance(result, tuple):
                deleted_snapshots = result[0] if len(result) > 0 else 0
                nondeleted_snapshot_id = result[1] if len(result) > 1 else None
            else:
                deleted_snapshots = result
                nondeleted_snapshot_id = None
            
            if deleted_snapshots > 0:
                return True
            else:
                print("Failed to delete shadow copy")
                if nondeleted_snapshot_id:
                    print(f"Non-deleted snapshot ID: {{{nondeleted_snapshot_id}}}")
                return False
        except Exception as delete_error:
            print(f"DeleteSnapshots error: {delete_error}")
            hr_error = delete_error.args[0] if delete_error.args else 0
            print(f"HRESULT: 0x{hr_error & 0xFFFFFFFF:08X}")
            
            if hr_error == 0x80042308:
                print("  -> VSS_E_VOLUME_NOT_SUPPORTED_BY_PROVIDER")
            elif hr_error == 0x8004230C:
                print("  -> VSS_E_BAD_STATE")
            elif hr_error == 0x80042316:
                print("  -> VSS_E_OBJECT_NOT_FOUND")
            
            return False
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        if backup:
            del backup
        
        if com_initialized:
            CoUninitialize()


def main():
    if not is_administrator():
        print("ERROR: Administrator privileges required")
        return
    
    if len(sys.argv) < 2:
        print("Usage: python delete.py <device_path>")
        print("Example: python delete.py \\\\?\\GLOBALROOT\\Device\\HarddiskVolumeShadowCopy8")
        return
    
    device_path = sys.argv[1]
    
    if delete_shadow_copy(device_path):
        print("Shadow copy deleted successfully")
    else:
        print("Failed to delete shadow copy")


if __name__ == "__main__":
    main()
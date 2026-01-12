import sys
import ctypes
import argparse
from ctypes import *
from ctypes.wintypes import *
from comtypes import GUID, IUnknown, COMMETHOD, HRESULT, BSTR
from comtypes import CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED
import comtypes.client
from datetime import datetime


# Constants
VSS_CTX_BACKUP = 0
VSS_CTX_ALL = 0xffffffff
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
    _fields_ = [("Snap", VSS_SNAPSHOT_PROP)]

class VSS_OBJECT_PROP(Structure):
    _fields_ = [("Type", ctypes.c_int), ("Obj", VSS_OBJECT_UNION)]

# Interfaces
class IVssEnumObject(IUnknown):
    _iid_ = GUID("{AE1C7110-2F60-11d3-8A39-00C04F72D8E3}")
    _methods_ = [
        COMMETHOD([], HRESULT, 'Next', (['in'], ULONG, 'celt'), (['out'], POINTER(VSS_OBJECT_PROP), 'rgelt'), (['out'], POINTER(ULONG), 'pceltFetched')),
        COMMETHOD([], HRESULT, 'Skip', (['in'], ULONG, 'celt')),
        COMMETHOD([], HRESULT, 'Reset'),
        COMMETHOD([], HRESULT, 'Clone', (['out'], POINTER(POINTER(IUnknown)), 'ppenum')),
    ]

class IVssAsync(IUnknown):
    _iid_ = GUID("{507C37B4-CF5B-4e95-B0AF-14EB9767467E}")
    _methods_ = [
        COMMETHOD([], HRESULT, 'Cancel'),
        COMMETHOD([], HRESULT, 'Wait', (['in'], DWORD, 'dwMilliseconds')),
        COMMETHOD([], HRESULT, 'QueryStatus', (['out'], POINTER(HRESULT), 'pHrResult'), (['out'], POINTER(INT), 'pReserved')),
    ]

class IVssBackupComponents(IUnknown):
    _iid_ = GUID("{665c1d5f-c218-414d-a05d-7fef5f9d5c86}")
    _methods_ = [
        COMMETHOD([], HRESULT, 'GetWriterComponentsCount', (['out'], POINTER(UINT), 'pcComponents')),
        COMMETHOD([], HRESULT, 'GetWriterComponents', (['in'], UINT, 'iWriter'), (['out'], POINTER(POINTER(IUnknown)), 'ppWriter')),
        COMMETHOD([], HRESULT, 'InitializeForBackup', (['in'], BSTR, 'bstrXML')),
        COMMETHOD([], HRESULT, 'SetBackupState', (['in'], BOOL, 'bSelectComponents'), (['in'], BOOL, 'bBackupBootableSystemState'), (['in'], ctypes.c_int, 'backupType'), (['in'], BOOL, 'bPartialFileSupport')),
        COMMETHOD([], HRESULT, 'InitializeForRestore', (['in'], BSTR, 'bstrXML')),
        COMMETHOD([], HRESULT, 'SetRestoreState', (['in'], ctypes.c_int, 'restoreType')),
        COMMETHOD([], HRESULT, 'GatherWriterMetadata', (['out'], POINTER(POINTER(IVssAsync)), 'ppAsync')),
        COMMETHOD([], HRESULT, 'GetWriterMetadataCount', (['out'], POINTER(UINT), 'pcWriters')),
        COMMETHOD([], HRESULT, 'GetWriterMetadata', (['in'], UINT, 'iWriter'), (['out'], POINTER(GUID), 'pInstanceId'), (['out'], POINTER(POINTER(IUnknown)), 'ppMetadata')),
        COMMETHOD([], HRESULT, 'FreeWriterMetadata'),
        COMMETHOD([], HRESULT, 'AddComponent', (['in'], POINTER(GUID), 'instanceId'), (['in'], POINTER(GUID), 'writerId'), (['in'], ctypes.c_int, 'ct'), (['in'], LPCWSTR, 'wszLogicalPath'), (['in'], LPCWSTR, 'wszComponentName')),
        COMMETHOD([], HRESULT, 'PrepareForBackup', (['out'], POINTER(POINTER(IVssAsync)), 'ppAsync')),
        COMMETHOD([], HRESULT, 'AbortBackup'),
        COMMETHOD([], HRESULT, 'GatherWriterStatus', (['out'], POINTER(POINTER(IVssAsync)), 'pAsync')),
        COMMETHOD([], HRESULT, 'GetWriterStatusCount', (['out'], POINTER(UINT), 'pcWriters')),
        COMMETHOD([], HRESULT, 'FreeWriterStatus'),
        COMMETHOD([], HRESULT, 'GetWriterStatus', (['in'], UINT, 'iWriter'), (['out'], POINTER(GUID), 'pidInstance'), (['out'], POINTER(GUID), 'pidWriter'), (['out'], POINTER(BSTR), 'pbstrWriter'), (['out'], POINTER(ctypes.c_int), 'pnStatus'), (['out'], POINTER(ctypes.c_int), 'phrFailureWriter'), (['out'], POINTER(ctypes.c_int), 'phrApplication'), (['out'], POINTER(BSTR), 'pbstrApplicationMessage')),
        COMMETHOD([], HRESULT, 'SetBackupSucceeded', (['in'], POINTER(GUID), 'instanceId'), (['in'], POINTER(GUID), 'writerId'), (['in'], ctypes.c_int, 'ct'), (['in'], LPCWSTR, 'wszLogicalPath'), (['in'], LPCWSTR, 'wszComponentName'), (['in'], BOOL, 'bSucceeded')),
        COMMETHOD([], HRESULT, 'SetBackupOptions', (['in'], POINTER(GUID), 'writerId'), (['in'], ctypes.c_int, 'ct'), (['in'], LPCWSTR, 'wszLogicalPath'), (['in'], LPCWSTR, 'wszComponentName'), (['in'], LPCWSTR, 'wszBackupOptions')),
        COMMETHOD([], HRESULT, 'SetSelectedForRestore', (['in'], POINTER(GUID), 'writerId'), (['in'], ctypes.c_int, 'ct'), (['in'], LPCWSTR, 'wszLogicalPath'), (['in'], LPCWSTR, 'wszComponentName'), (['in'], BOOL, 'bSelectedForRestore')),
        COMMETHOD([], HRESULT, 'SetRestoreOptions', (['in'], POINTER(GUID), 'writerId'), (['in'], ctypes.c_int, 'ct'), (['in'], LPCWSTR, 'wszLogicalPath'), (['in'], LPCWSTR, 'wszComponentName'), (['in'], LPCWSTR, 'wszRestoreOptions')),
        COMMETHOD([], HRESULT, 'SetAdditionalRestores', (['in'], POINTER(GUID), 'writerId'), (['in'], ctypes.c_int, 'ct'), (['in'], LPCWSTR, 'wszLogicalPath'), (['in'], LPCWSTR, 'wszComponentName'), (['in'], BOOL, 'bAdditionalRestores')),
        COMMETHOD([], HRESULT, 'SetPreviousBackupStamp', (['in'], POINTER(GUID), 'writerId'), (['in'], ctypes.c_int, 'ct'), (['in'], LPCWSTR, 'wszLogicalPath'), (['in'], LPCWSTR, 'wszComponentName'), (['in'], LPCWSTR, 'wszPreviousBackupStamp')),
        COMMETHOD([], HRESULT, 'SaveAsXML', (['out'], POINTER(BSTR), 'pbstrXML')),
        COMMETHOD([], HRESULT, 'BackupComplete', (['out'], POINTER(POINTER(IVssAsync)), 'ppAsync')),
        COMMETHOD([], HRESULT, 'AddAlternativeLocationMapping', (['in'], POINTER(GUID), 'writerId'), (['in'], ctypes.c_int, 'ct'), (['in'], LPCWSTR, 'wszLogicalPath'), (['in'], LPCWSTR, 'wszComponentName'), (['in'], LPCWSTR, 'wszPath'), (['in'], LPCWSTR, 'wszFilespec'), (['in'], BOOL, 'bRecursive'), (['in'], LPCWSTR, 'wszDestination')),
        COMMETHOD([], HRESULT, 'AddRestoreSubcomponent', (['in'], POINTER(GUID), 'writerId'), (['in'], ctypes.c_int, 'ct'), (['in'], LPCWSTR, 'wszLogicalPath'), (['in'], LPCWSTR, 'wszComponentName'), (['in'], LPCWSTR, 'wszSubComponentLogicalPath'), (['in'], LPCWSTR, 'wszSubComponentName'), (['in'], BOOL, 'bRepair')),
        COMMETHOD([], HRESULT, 'SetFileRestoreStatus', (['in'], POINTER(GUID), 'writerId'), (['in'], ctypes.c_int, 'ct'), (['in'], LPCWSTR, 'wszLogicalPath'), (['in'], LPCWSTR, 'wszComponentName'), (['in'], ctypes.c_int, 'status')),
        COMMETHOD([], HRESULT, 'AddNewTarget', (['in'], POINTER(GUID), 'writerId'), (['in'], ctypes.c_int, 'ct'), (['in'], LPCWSTR, 'wszLogicalPath'), (['in'], LPCWSTR, 'wszComponentName'), (['in'], LPCWSTR, 'wszPath'), (['in'], LPCWSTR, 'wszFileName'), (['in'], BOOL, 'bRecursive'), (['in'], LPCWSTR, 'wszAlternatePath')),
        COMMETHOD([], HRESULT, 'SetRangesFilePath', (['in'], POINTER(GUID), 'writerId'), (['in'], ctypes.c_int, 'ct'), (['in'], LPCWSTR, 'wszLogicalPath'), (['in'], LPCWSTR, 'wszComponentName'), (['in'], UINT, 'iPartialFile'), (['in'], LPCWSTR, 'wszRangesFile')),
        COMMETHOD([], HRESULT, 'PreRestore', (['out'], POINTER(POINTER(IVssAsync)), 'ppAsync')),
        COMMETHOD([], HRESULT, 'PostRestore', (['out'], POINTER(POINTER(IVssAsync)), 'ppAsync')),
        COMMETHOD([], HRESULT, 'SetContext', (['in'], LONG, 'lContext')),
        COMMETHOD([], HRESULT, 'StartSnapshotSet', (['out'], POINTER(GUID), 'pSnapshotSetId')),
        COMMETHOD([], HRESULT, 'AddToSnapshotSet', (['in'], LPCWSTR, 'pwszVolumeName'), (['in'], POINTER(GUID), 'ProviderId'), (['out'], POINTER(GUID), 'pidSnapshot')),
        COMMETHOD([], HRESULT, 'DoSnapshotSet', (['out'], POINTER(POINTER(IVssAsync)), 'ppAsync')),
        COMMETHOD([], HRESULT, 'DeleteSnapshots', (['in'], GUID, 'SourceObjectId'), (['in'], ctypes.c_int, 'eSourceObjectType'), (['in'], BOOL, 'bForceDelete'), (['out'], POINTER(LONG), 'plDeletedSnapshots'), (['out'], POINTER(GUID), 'pNondeletedSnapshotID')),
        COMMETHOD([], HRESULT, 'ImportSnapshots', (['out'], POINTER(POINTER(IVssAsync)), 'ppAsync')),
        COMMETHOD([], HRESULT, 'BreakSnapshotSet', (['in'], POINTER(GUID), 'SnapshotSetId')),
        COMMETHOD([], HRESULT, 'GetSnapshotProperties', (['in'], POINTER(GUID), 'SnapshotId'), (['out'], POINTER(VSS_SNAPSHOT_PROP), 'pProp')),
        COMMETHOD([], HRESULT, 'Query', (['in'], POINTER(GUID), 'QueriedObjectId'), (['in'], ctypes.c_int, 'eQueriedObjectType'), (['in'], ctypes.c_int, 'eReturnedObjectsType'), (['out'], POINTER(POINTER(IVssEnumObject)), 'ppEnum')),
        COMMETHOD([], HRESULT, 'IsVolumeSupported', (['in'], POINTER(GUID), 'ProviderId'), (['in'], LPCWSTR, 'pwszVolumeName'), (['out'], POINTER(BOOL), 'pbSupportedByThisProvider')),
        COMMETHOD([], HRESULT, 'DisableWriterClasses', (['in'], POINTER(GUID), 'rgWriterClassId'), (['in'], UINT, 'cClassId')),
        COMMETHOD([], HRESULT, 'EnableWriterClasses', (['in'], POINTER(GUID), 'rgWriterClassId'), (['in'], UINT, 'cClassId')),
        COMMETHOD([], HRESULT, 'DisableWriterInstances', (['in'], POINTER(GUID), 'rgWriterInstanceId'), (['in'], UINT, 'cInstanceId')),
        COMMETHOD([], HRESULT, 'ExposeSnapshot', (['in'], POINTER(GUID), 'SnapshotId'), (['in'], LPCWSTR, 'wszPathFromRoot'), (['in'], LONG, 'lAttributes'), (['in'], LPCWSTR, 'wszExpose'), (['out'], POINTER(LPWSTR), 'pwszExposed')),
        COMMETHOD([], HRESULT, 'RevertToSnapshot', (['in'], POINTER(GUID), 'SnapshotId'), (['in'], BOOL, 'bForceDismount')),
        COMMETHOD([], HRESULT, 'QueryRevertStatus', (['in'], LPCWSTR, 'pwszVolume'), (['out'], POINTER(POINTER(IVssAsync)), 'ppAsync')),
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
    possible_names = ["CreateVssBackupComponentsInternal", "CreateVssBackupComponents", "?CreateVssBackupComponents@@YAJPEAPEAVIVssBackupComponents@@@Z"]
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


def init_com():
    try:
        CoInitializeEx(COINIT_MULTITHREADED)
        return True
    except OSError as e:
        if e.winerror == -2147417850:
            return False
        elif e.winerror not in (0, 1):
            print(f"Error initializing COM: {e}")
            return None
    return False


def create_shadow_copy(volume_path):
    com_init = init_com()
    if com_init is None:
        return None, None
    backup = None
    try:
        backup = create_vss_backup_components()
        if not backup:
            return None, None
        hr = backup.InitializeForBackup(None)
        if hr != 0:
            return None, None
        guid_null = GUID()
        hr = backup.SetContext(VSS_CTX_BACKUP)
        if hr != 0:
            return None, None
        backup.SetBackupState(False, False, VSS_BT_FULL, False)
        try:
            async_metadata = backup.GatherWriterMetadata()
            if async_metadata:
                async_metadata.Wait(0xFFFFFFFF)
        except:
            pass
        snapshot_set_id = backup.StartSnapshotSet()
        if not snapshot_set_id:
            return None, None
        snapshot_id = backup.AddToSnapshotSet(volume_path, guid_null)
        if not snapshot_id:
            return None, None
        try:
            async_prepare = backup.PrepareForBackup()
            if async_prepare:
                async_prepare.Wait(0xFFFFFFFF)
        except:
            pass
        async_snapshot = backup.DoSnapshotSet()
        if async_snapshot:
            hr = async_snapshot.Wait(0xFFFFFFFF)
        else:
            hr = -1
        if hr == 0:
            snap_prop = backup.GetSnapshotProperties(snapshot_id)
            if snap_prop:
                device_object = None
                if snap_prop.m_pwszSnapshotDeviceObject:
                    device_object = ctypes.wstring_at(snap_prop.m_pwszSnapshotDeviceObject)
                snapshot_id_str = str(snapshot_id)
                VssFreeSnapshotProperties(byref(snap_prop))
                return snapshot_id_str, device_object
        return None, None
    except Exception as e:
        print(f"Error: {e}")
        return None, None
    finally:
        if backup:
            del backup
        if com_init:
            CoUninitialize()


def list_shadow_copies():
    com_init = init_com()
    if com_init is None:
        return
    backup = None
    enum_obj = None
    try:
        backup = create_vss_backup_components()
        if not backup:
            return
        hr = backup.InitializeForBackup(None)
        if hr != 0:
            return
        hr = backup.SetContext(VSS_CTX_ALL)
        if hr != 0:
            hr = backup.SetContext(VSS_CTX_BACKUP)
            if hr != 0:
                return
        guid_null = GUID()
        enum_obj_ptr = backup.Query(guid_null, VSS_OBJECT_TYPE.VSS_OBJECT_NONE, VSS_OBJECT_TYPE.VSS_OBJECT_SNAPSHOT)
        if not enum_obj_ptr:
            print("No shadow copies found")
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
            except:
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
            print("No shadow copies found")
        else:
            print(f"Total: {count} shadow copies")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        if enum_obj:
            del enum_obj
        if backup:
            del backup
        if com_init:
            CoUninitialize()


def extract_snapshot_id(device_path):
    com_init = init_com()
    if com_init is None:
        return None
    backup = None
    enum_obj = None
    try:
        backup = create_vss_backup_components()
        if not backup:
            return None
        hr = backup.InitializeForBackup(None)
        if hr != 0:
            return None
        hr = backup.SetContext(VSS_CTX_ALL)
        if hr != 0:
            hr = backup.SetContext(VSS_CTX_BACKUP)
            if hr != 0:
                return None
        guid_null = GUID()
        enum_obj_ptr = backup.Query(guid_null, VSS_OBJECT_TYPE.VSS_OBJECT_NONE, VSS_OBJECT_TYPE.VSS_OBJECT_SNAPSHOT)
        if not enum_obj_ptr:
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
            except:
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
        return None
    except Exception as e:
        print(f"Error: {e}")
        return None
    finally:
        if enum_obj:
            del enum_obj
        if backup:
            del backup
        if com_init:
            CoUninitialize()


def delete_shadow_copy(device_path):
    com_init = init_com()
    if com_init is None:
        return False
    backup = None
    try:
        snapshot_id = extract_snapshot_id(device_path)
        if not snapshot_id:
            print(f"Shadow copy not found: {device_path}")
            return False
        backup = create_vss_backup_components()
        if not backup:
            return False
        hr = backup.InitializeForBackup(None)
        if hr != 0:
            return False
        hr = backup.SetContext(VSS_CTX_ALL)
        if hr != 0:
            hr = backup.SetContext(VSS_CTX_BACKUP)
            if hr != 0:
                return False
        try:
            result = backup.DeleteSnapshots(snapshot_id, VSS_OBJECT_TYPE.VSS_OBJECT_SNAPSHOT, True)
            if isinstance(result, tuple):
                deleted_snapshots = result[0] if len(result) > 0 else 0
            else:
                deleted_snapshots = result
            return deleted_snapshots > 0
        except Exception as e:
            print(f"Delete error: {e}")
            return False
    except Exception as e:
        print(f"Error: {e}")
        return False
    finally:
        if backup:
            del backup
        if com_init:
            CoUninitialize()


def main():
    if not is_administrator():
        print("ERROR: Administrator privileges required")
        sys.exit(1)
    
    parser = argparse.ArgumentParser(description='Shadow Copy Manager')
    parser.add_argument('-o', '--operation', required=True, choices=['create', 'list', 'delete'], help='Operation to perform')
    parser.add_argument('-v', '--volume', default=r'C:\\', help='Volume path for create operation (default: C:\\)')
    parser.add_argument('-s', '--snapshot', help='Shadow copy device path for delete operation')
    args = parser.parse_args()
    
    if args.operation == 'create':
        volume = args.volume
        if not volume.endswith('\\'):
            volume += '\\'
        snapshot_id, device_object = create_shadow_copy(volume)
        if snapshot_id and device_object:
            print(f"Shadow copy created: {device_object}")
        else:
            print("Failed to create shadow copy")
            sys.exit(1)
    
    elif args.operation == 'list':
        list_shadow_copies()
    
    elif args.operation == 'delete':
        if not args.snapshot:
            print("ERROR: -s/--snapshot required for delete operation")
            sys.exit(1)
        if delete_shadow_copy(args.snapshot):
            print("Shadow copy deleted successfully")
        else:
            print("Failed to delete shadow copy")
            sys.exit(1)


if __name__ == "__main__":
    main()
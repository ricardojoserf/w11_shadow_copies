using System;
using System.Runtime.InteropServices;


class ShadowCopyDeleter
{
    // Constants
    private const int VSS_CTX_BACKUP = 0;
    private static readonly Guid GUID_NULL = Guid.Empty;
    private const uint COINIT_MULTITHREADED = 0x0;

    // Enums
    private enum VSS_OBJECT_TYPE { VSS_OBJECT_UNKNOWN = 0, VSS_OBJECT_NONE = 1, VSS_OBJECT_SNAPSHOT_SET = 2, VSS_OBJECT_SNAPSHOT = 3, VSS_OBJECT_PROVIDER = 4, VSS_OBJECT_TYPE_COUNT = 5 }

    // Structs
    [StructLayout(LayoutKind.Sequential, Pack = 8)] private struct VSS_SNAPSHOT_PROP { public Guid m_SnapshotId; public Guid m_SnapshotSetId; public int m_lSnapshotsCount; public IntPtr m_pwszSnapshotDeviceObject; public IntPtr m_pwszOriginalVolumeName; public IntPtr m_pwszOriginatingMachine; public IntPtr m_pwszServiceMachine; public IntPtr m_pwszExposedName; public IntPtr m_pwszExposedPath; public Guid m_ProviderId; public int m_lSnapshotAttributes; public long m_tsCreationTimestamp; public int m_eStatus; }
    [StructLayout(LayoutKind.Explicit)] private struct VSS_OBJECT_UNION { [FieldOffset(0)] public VSS_SNAPSHOT_PROP Snap; }
    [StructLayout(LayoutKind.Sequential)] private struct VSS_OBJECT_PROP { public VSS_OBJECT_TYPE Type; public VSS_OBJECT_UNION Obj; }

    // Interfaces
    [ComImport, Guid("AE1C7110-2F60-11d3-8A39-00C04F72D8E3")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IVssEnumObject
    {
        [PreserveSig] int Next([In] uint celt, [Out] out VSS_OBJECT_PROP rgelt, [Out] out uint pceltFetched);
        void Skip([In] uint celt);
        void Reset();
        void Clone([Out, MarshalAs(UnmanagedType.Interface)] out IVssEnumObject ppenum);
    }

    [ComImport, Guid("665c1d5f-c218-414d-a05d-7fef5f9d5c86")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IVssBackupComponents
    {
        void GetWriterComponentsCount([Out] out uint pcComponents);
        [PreserveSig] int GetWriterComponents([In] uint iWriter, [Out, MarshalAs(UnmanagedType.Interface)] out object ppWriter);
        [PreserveSig] int InitializeForBackup([In, MarshalAs(UnmanagedType.BStr)] string bstrXML);
        [PreserveSig] int SetBackupState([In, MarshalAs(UnmanagedType.Bool)] bool bSelectComponents, [In, MarshalAs(UnmanagedType.Bool)] bool bBackupBootableSystemState, [In] int backupType, [In, MarshalAs(UnmanagedType.Bool)] bool bPartialFileSupport = false);
        void InitializeForRestore([In, MarshalAs(UnmanagedType.BStr)] string bstrXML);
        [PreserveSig] int SetRestoreState([In] int restoreType);
        [PreserveSig] int GatherWriterMetadata([Out, MarshalAs(UnmanagedType.Interface)] out object ppAsync);
        void GetWriterMetadataCount([Out] out uint pcWriters);
        [PreserveSig] int GetWriterMetadata([In] uint iWriter, [Out] out Guid pInstanceId, [Out, MarshalAs(UnmanagedType.Interface)] out object ppMetadata);
        void FreeWriterMetadata();
        [PreserveSig] int AddComponent([In] ref Guid instanceId, [In] ref Guid writerId, [In] int ct, [In, MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszComponentName);
        [PreserveSig] int PrepareForBackup([Out, MarshalAs(UnmanagedType.Interface)] out object ppAsync);
        void AbortBackup();
        [PreserveSig] int GatherWriterStatus([Out, MarshalAs(UnmanagedType.Interface)] out object pAsync);
        void GetWriterStatusCount([Out] out uint pcWriters);
        void FreeWriterStatus();
        [PreserveSig] int GetWriterStatus([In] uint iWriter, [Out] out Guid pidInstance, [Out] out Guid pidWriter, [Out, MarshalAs(UnmanagedType.BStr)] out string pbstrWriter, [Out] out int pnStatus, [Out] out int phrFailureWriter, [Out] out int phrApplication, [Out, MarshalAs(UnmanagedType.BStr)] out string pbstrApplicationMessage);
        [PreserveSig] int SetBackupSucceeded([In] ref Guid instanceId, [In] ref Guid writerId, [In] int ct, [In, MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [In, MarshalAs(UnmanagedType.Bool)] bool bSucceeded);
        [PreserveSig] int SetBackupOptions([In] ref Guid writerId, [In] int ct, [In, MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [In, MarshalAs(UnmanagedType.LPWStr)] string wszBackupOptions);
        [PreserveSig] int SetSelectedForRestore([In] ref Guid writerId, [In] int ct, [In, MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [In, MarshalAs(UnmanagedType.Bool)] bool bSelectedForRestore);
        [PreserveSig] int SetRestoreOptions([In] ref Guid writerId, [In] int ct, [In, MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [In, MarshalAs(UnmanagedType.LPWStr)] string wszRestoreOptions);
        void SetAdditionalRestores([In] ref Guid writerId, [In] int ct, [In, MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [In, MarshalAs(UnmanagedType.Bool)] bool bAdditionalRestores);
        [PreserveSig] int SetPreviousBackupStamp([In] ref Guid writerId, [In] int ct, [In, MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [In, MarshalAs(UnmanagedType.LPWStr)] string wszPreviousBackupStamp);
        void SaveAsXML([Out, MarshalAs(UnmanagedType.BStr)] out string pbstrXML);
        [PreserveSig] int BackupComplete([Out, MarshalAs(UnmanagedType.Interface)] out object ppAsync);
        void AddAlternativeLocationMapping([In] ref Guid writerId, [In] int ct, [In, MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [In, MarshalAs(UnmanagedType.LPWStr)] string wszPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszFilespec, [In, MarshalAs(UnmanagedType.Bool)] bool bRecursive, [In, MarshalAs(UnmanagedType.LPWStr)] string wszDestination);
        void AddRestoreSubcomponent([In] ref Guid writerId, [In] int ct, [In, MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [In, MarshalAs(UnmanagedType.LPWStr)] string wszSubComponentLogicalPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszSubComponentName, [In, MarshalAs(UnmanagedType.Bool)] bool bRepair);
        [PreserveSig] int SetFileRestoreStatus([In] ref Guid writerId, [In] int ct, [In, MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [In] int status);
        void AddNewTarget([In] ref Guid writerId, [In] int ct, [In, MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [In, MarshalAs(UnmanagedType.LPWStr)] string wszPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszFileName, [In, MarshalAs(UnmanagedType.Bool)] bool bRecursive, [In, MarshalAs(UnmanagedType.LPWStr)] string wszAlternatePath);
        [PreserveSig] int SetRangesFilePath([In] ref Guid writerId, [In] int ct, [In, MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [In] uint iPartialFile, [In, MarshalAs(UnmanagedType.LPWStr)] string wszRangesFile);
        [PreserveSig] int PreRestore([Out, MarshalAs(UnmanagedType.Interface)] out object ppAsync);
        [PreserveSig] int PostRestore([Out, MarshalAs(UnmanagedType.Interface)] out object ppAsync);
        [PreserveSig] int SetContext([In] int lContext);
        [PreserveSig] int StartSnapshotSet([Out] out Guid pSnapshotSetId);
        [PreserveSig] int AddToSnapshotSet([In, MarshalAs(UnmanagedType.LPWStr)] string pwszVolumeName, [In] ref Guid ProviderId, [Out] out Guid pidSnapshot);
        [PreserveSig] int DoSnapshotSet([Out, MarshalAs(UnmanagedType.Interface)] out object ppAsync);
        void DeleteSnapshots([In] Guid SourceObjectId, [In] int eSourceObjectType, [In, MarshalAs(UnmanagedType.Bool)] bool bForceDelete, [Out] out int plDeletedSnapshots, [Out] out Guid pNondeletedSnapshotID);
        [PreserveSig] int ImportSnapshots([Out, MarshalAs(UnmanagedType.Interface)] out object ppAsync);
        void BreakSnapshotSet([In] ref Guid SnapshotSetId);
        [PreserveSig] int GetSnapshotProperties([In] ref Guid SnapshotId, [Out] out VSS_SNAPSHOT_PROP pProp);
        void Query([In] ref Guid QueriedObjectId, [In] int eQueriedObjectType, [In] int eReturnedObjectsType, [Out, MarshalAs(UnmanagedType.Interface)] out IVssEnumObject ppEnum);
        [PreserveSig] int IsVolumeSupported([In] ref Guid ProviderId, [In, MarshalAs(UnmanagedType.LPWStr)] string pwszVolumeName, [Out, MarshalAs(UnmanagedType.Bool)] out bool pbSupportedByThisProvider);
        void DisableWriterClasses([In] ref Guid rgWriterClassId, [In] uint cClassId);
        void EnableWriterClasses([In] ref Guid rgWriterClassId, [In] uint cClassId);
        void DisableWriterInstances([In] ref Guid rgWriterInstanceId, [In] uint cInstanceId);
        [PreserveSig] int ExposeSnapshot([In] ref Guid SnapshotId, [In, MarshalAs(UnmanagedType.LPWStr)] string wszPathFromRoot, [In] int lAttributes, [In, MarshalAs(UnmanagedType.LPWStr)] string wszExpose, [Out, MarshalAs(UnmanagedType.LPWStr)] out string pwszExposed);
        void RevertToSnapshot([In] ref Guid SnapshotId, [In, MarshalAs(UnmanagedType.Bool)] bool bForceDismount);
        void QueryRevertStatus([In, MarshalAs(UnmanagedType.LPWStr)] string pwszVolume, [Out, MarshalAs(UnmanagedType.Interface)] out object ppAsync);
    }

    // Functions
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr LoadLibrary(string lpFileName);
    [DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true)] private static extern IntPtr GetProcAddress(IntPtr hModule, string lpProcName);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] private delegate int CreateVssBackupComponentsDelegate([Out, MarshalAs(UnmanagedType.Interface)] out IVssBackupComponents ppBackup);
    [DllImport("ole32.dll")] private static extern int CoInitializeEx(IntPtr pvReserved, uint dwCoInit);
    [DllImport("ole32.dll")] private static extern void CoUninitialize();
    [DllImport("VssApi.dll", CallingConvention = CallingConvention.StdCall)] private static extern void VssFreeSnapshotProperties([In] ref VSS_SNAPSHOT_PROP pProp);


    private static IVssBackupComponents CreateVssBackupComponentsInstance()
    {
        IntPtr hModule = LoadLibrary("VssApi.dll");
        if (hModule == IntPtr.Zero) throw new Exception($"Failed to load VssApi.dll. Error: {Marshal.GetLastWin32Error()}");

        string[] possibleNames = new string[] { "CreateVssBackupComponentsInternal", "CreateVssBackupComponents", "?CreateVssBackupComponents@@YAJPEAPEAVIVssBackupComponents@@@Z" };
        IntPtr procAddr = IntPtr.Zero;
        foreach (string name in possibleNames)
        {
            procAddr = GetProcAddress(hModule, name);
            if (procAddr != IntPtr.Zero) break;
        }

        if (procAddr == IntPtr.Zero) throw new Exception($"CreateVssBackupComponents not found in VssApi.dll. Error: {Marshal.GetLastWin32Error()}");

        var createFunc = (CreateVssBackupComponentsDelegate)Marshal.GetDelegateForFunctionPointer(procAddr, typeof(CreateVssBackupComponentsDelegate));
        IVssBackupComponents backup;
        int hr = createFunc(out backup);
        if (hr != 0) throw new Exception($"CreateVssBackupComponents failed with HRESULT: 0x{hr:X8}");
        if (backup == null) throw new Exception("CreateVssBackupComponents returned NULL");
        return backup;
    }


    static bool IsAdministrator()
    {
        try
        {
            var identity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var principal = new System.Security.Principal.WindowsPrincipal(identity);
            return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
        }
        catch { return false; }
    }


    static bool DeleteShadowCopy(string devicePath)
    {
        IVssBackupComponents pBackup = null;
        int hr = CoInitializeEx(IntPtr.Zero, COINIT_MULTITHREADED);
        if (hr != 0 && hr != 1) return false;

        try
        {
            pBackup = CreateVssBackupComponentsInstance();
            hr = pBackup.InitializeForBackup(null);
            if (hr != 0) return false;

            hr = pBackup.SetContext(VSS_CTX_BACKUP);
            if (hr != 0) return false;

            Guid snapshotId = ExtractSnapshotIdFromPath(devicePath);
            if (snapshotId == Guid.Empty)
            {
                Console.WriteLine("Failed to extract snapshot ID from path");
                return false;
            }

            int deletedSnapshots;
            Guid nondeletedSnapshotID;
            pBackup.DeleteSnapshots(snapshotId, 3, true, out deletedSnapshots, out nondeletedSnapshotID);

            return deletedSnapshots > 0;
        }
        finally
        {
            if (pBackup != null) Marshal.ReleaseComObject(pBackup);
            CoUninitialize();
        }
    }

    static Guid ExtractSnapshotIdFromPath(string devicePath)
    {
        IVssBackupComponents pBackup = null;
        IVssEnumObject pEnum = null;
        int hr = CoInitializeEx(IntPtr.Zero, COINIT_MULTITHREADED);
        if (hr != 0 && hr != 1) return Guid.Empty;

        try
        {
            pBackup = CreateVssBackupComponentsInstance();
            hr = pBackup.InitializeForBackup(null);
            if (hr != 0) return Guid.Empty;

            hr = pBackup.SetContext(VSS_CTX_BACKUP);
            if (hr != 0) return Guid.Empty;

            Guid guidNull = GUID_NULL;
            pBackup.Query(ref guidNull, 1, 3, out pEnum);
            if (pEnum == null) return Guid.Empty;

            while (true)
            {
                VSS_OBJECT_PROP prop;
                uint fetched;
                hr = pEnum.Next(1, out prop, out fetched);
                if (hr == 1 || fetched == 0) break;
                if (hr != 0) break;

                if (prop.Type == VSS_OBJECT_TYPE.VSS_OBJECT_SNAPSHOT)
                {
                    VSS_SNAPSHOT_PROP snap = prop.Obj.Snap;
                    if (snap.m_pwszSnapshotDeviceObject != IntPtr.Zero)
                    {
                        string deviceObject = Marshal.PtrToStringUni(snap.m_pwszSnapshotDeviceObject);
                        if (deviceObject.Equals(devicePath, StringComparison.OrdinalIgnoreCase))
                        {
                            Guid foundId = snap.m_SnapshotId;
                            VssFreeSnapshotProperties(ref snap);
                            return foundId;
                        }
                    }
                    VssFreeSnapshotProperties(ref snap);
                }
            }
            return Guid.Empty;
        }
        finally
        {
            if (pEnum != null) Marshal.ReleaseComObject(pEnum);
            if (pBackup != null) Marshal.ReleaseComObject(pBackup);
            CoUninitialize();
        }
    }


    static void Main(string[] args)
    {
        try
        {
            if (!IsAdministrator())
            {
                Console.WriteLine("ERROR: Administrator privileges required");
                return;
            }

            if (args.Length == 0)
            {
                Console.WriteLine("Usage: program.exe <device_path>");
                Console.WriteLine("Example: program.exe \\\\?\\GLOBALROOT\\Device\\HarddiskVolumeShadowCopy8");
                return;
            }

            string devicePath = args[0];
            Console.WriteLine($"Deleting shadow copy: {devicePath}");

            if (DeleteShadowCopy(devicePath))
            {
                Console.WriteLine("Shadow copy deleted successfully");
            }
            else
            {
                Console.WriteLine("Failed to delete shadow copy");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
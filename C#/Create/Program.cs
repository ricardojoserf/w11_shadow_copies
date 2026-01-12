using System;
using System.Runtime.InteropServices;


class ShadowCopyCreator
{
    // Constants
    private const int VSS_CTX_BACKUP = 0;
    private static readonly Guid GUID_NULL = Guid.Empty;
    private const uint COINIT_MULTITHREADED = 0x0;
    private const int VSS_BT_FULL = 1;

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

    [ComImport, Guid("507C37B4-CF5B-4e95-B0AF-14EB9767467E")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IVssAsync
    {
        void Cancel();
        [PreserveSig] int Wait([In] uint dwMilliseconds = 0xFFFFFFFF);
        [PreserveSig] int QueryStatus([Out] out int pHrResult, [Out] out int pReserved);
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
        [PreserveSig] int GatherWriterMetadata([Out, MarshalAs(UnmanagedType.Interface)] out IVssAsync ppAsync);
        void GetWriterMetadataCount([Out] out uint pcWriters);
        [PreserveSig] int GetWriterMetadata([In] uint iWriter, [Out] out Guid pInstanceId, [Out, MarshalAs(UnmanagedType.Interface)] out object ppMetadata);
        void FreeWriterMetadata();
        [PreserveSig] int AddComponent([In] ref Guid instanceId, [In] ref Guid writerId, [In] int ct, [In, MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszComponentName);
        [PreserveSig] int PrepareForBackup([Out, MarshalAs(UnmanagedType.Interface)] out IVssAsync ppAsync);
        void AbortBackup();
        [PreserveSig] int GatherWriterStatus([Out, MarshalAs(UnmanagedType.Interface)] out IVssAsync pAsync);
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
        [PreserveSig] int BackupComplete([Out, MarshalAs(UnmanagedType.Interface)] out IVssAsync ppAsync);
        void AddAlternativeLocationMapping([In] ref Guid writerId, [In] int ct, [In, MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [In, MarshalAs(UnmanagedType.LPWStr)] string wszPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszFilespec, [In, MarshalAs(UnmanagedType.Bool)] bool bRecursive, [In, MarshalAs(UnmanagedType.LPWStr)] string wszDestination);
        void AddRestoreSubcomponent([In] ref Guid writerId, [In] int ct, [In, MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [In, MarshalAs(UnmanagedType.LPWStr)] string wszSubComponentLogicalPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszSubComponentName, [In, MarshalAs(UnmanagedType.Bool)] bool bRepair);
        [PreserveSig] int SetFileRestoreStatus([In] ref Guid writerId, [In] int ct, [In, MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [In] int status);
        void AddNewTarget([In] ref Guid writerId, [In] int ct, [In, MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [In, MarshalAs(UnmanagedType.LPWStr)] string wszPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszFileName, [In, MarshalAs(UnmanagedType.Bool)] bool bRecursive, [In, MarshalAs(UnmanagedType.LPWStr)] string wszAlternatePath);
        [PreserveSig] int SetRangesFilePath([In] ref Guid writerId, [In] int ct, [In, MarshalAs(UnmanagedType.LPWStr)] string wszLogicalPath, [In, MarshalAs(UnmanagedType.LPWStr)] string wszComponentName, [In] uint iPartialFile, [In, MarshalAs(UnmanagedType.LPWStr)] string wszRangesFile);
        [PreserveSig] int PreRestore([Out, MarshalAs(UnmanagedType.Interface)] out IVssAsync ppAsync);
        [PreserveSig] int PostRestore([Out, MarshalAs(UnmanagedType.Interface)] out IVssAsync ppAsync);
        [PreserveSig] int SetContext([In] int lContext);
        [PreserveSig] int StartSnapshotSet([Out] out Guid pSnapshotSetId);
        [PreserveSig] int AddToSnapshotSet([In, MarshalAs(UnmanagedType.LPWStr)] string pwszVolumeName, [In] ref Guid ProviderId, [Out] out Guid pidSnapshot);
        [PreserveSig] int DoSnapshotSet([Out, MarshalAs(UnmanagedType.Interface)] out IVssAsync ppAsync);
        void DeleteSnapshots([In] Guid SourceObjectId, [In] int eSourceObjectType, [In, MarshalAs(UnmanagedType.Bool)] bool bForceDelete, [Out] out int plDeletedSnapshots, [Out] out Guid pNondeletedSnapshotID);
        [PreserveSig] int ImportSnapshots([Out, MarshalAs(UnmanagedType.Interface)] out IVssAsync ppAsync);
        void BreakSnapshotSet([In] ref Guid SnapshotSetId);
        [PreserveSig] int GetSnapshotProperties([In] ref Guid SnapshotId, [Out] out VSS_SNAPSHOT_PROP pProp);
        void Query([In] ref Guid QueriedObjectId, [In] int eQueriedObjectType, [In] int eReturnedObjectsType, [Out, MarshalAs(UnmanagedType.Interface)] out IVssEnumObject ppEnum);
        [PreserveSig] int IsVolumeSupported([In] ref Guid ProviderId, [In, MarshalAs(UnmanagedType.LPWStr)] string pwszVolumeName, [Out, MarshalAs(UnmanagedType.Bool)] out bool pbSupportedByThisProvider);
        void DisableWriterClasses([In] ref Guid rgWriterClassId, [In] uint cClassId);
        void EnableWriterClasses([In] ref Guid rgWriterClassId, [In] uint cClassId);
        void DisableWriterInstances([In] ref Guid rgWriterInstanceId, [In] uint cInstanceId);
        [PreserveSig] int ExposeSnapshot([In] ref Guid SnapshotId, [In, MarshalAs(UnmanagedType.LPWStr)] string wszPathFromRoot, [In] int lAttributes, [In, MarshalAs(UnmanagedType.LPWStr)] string wszExpose, [Out, MarshalAs(UnmanagedType.LPWStr)] out string pwszExposed);
        void RevertToSnapshot([In] ref Guid SnapshotId, [In, MarshalAs(UnmanagedType.Bool)] bool bForceDismount);
        void QueryRevertStatus([In, MarshalAs(UnmanagedType.LPWStr)] string pwszVolume, [Out, MarshalAs(UnmanagedType.Interface)] out IVssAsync ppAsync);
    }

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


    static string CreateShadowCopy(string volumePath, out string deviceObject)
    {
        deviceObject = null;
        IVssBackupComponents pBackup = null;

        int hr = CoInitializeEx(IntPtr.Zero, COINIT_MULTITHREADED);
        if (hr != 0 && hr != 1) return null;

        try
        {
            pBackup = CreateVssBackupComponentsInstance();
            hr = pBackup.InitializeForBackup(null);
            if (hr != 0) return null;

            Guid guidNull = GUID_NULL;
            bool bSupported = false;
            hr = pBackup.IsVolumeSupported(ref guidNull, volumePath, out bSupported);
            if (hr == 0 && !bSupported) return null;

            hr = pBackup.SetContext(VSS_CTX_BACKUP);
            if (hr != 0) return null;

            hr = pBackup.SetBackupState(false, false, VSS_BT_FULL, false);

            IVssAsync pAsyncMetadata = null;
            hr = pBackup.GatherWriterMetadata(out pAsyncMetadata);
            if (hr == 0 && pAsyncMetadata != null)
            {
                pAsyncMetadata.Wait();
                Marshal.ReleaseComObject(pAsyncMetadata);
            }

            Guid snapshotSetId;
            hr = pBackup.StartSnapshotSet(out snapshotSetId);
            if (hr != 0) return null;

            Guid snapshotId;
            hr = pBackup.AddToSnapshotSet(volumePath, ref guidNull, out snapshotId);
            if (hr != 0) return null;

            IVssAsync pAsyncPrepare = null;
            hr = pBackup.PrepareForBackup(out pAsyncPrepare);
            if (hr == 0 && pAsyncPrepare != null)
            {
                pAsyncPrepare.Wait();
                Marshal.ReleaseComObject(pAsyncPrepare);
            }

            IVssAsync pAsyncSnapshot = null;
            hr = pBackup.DoSnapshotSet(out pAsyncSnapshot);
            if (hr == 0 && pAsyncSnapshot != null)
            {
                pAsyncSnapshot.Wait();
                Marshal.ReleaseComObject(pAsyncSnapshot);
            }

            if (hr == 0)
            {
                VSS_SNAPSHOT_PROP snapProp;
                hr = pBackup.GetSnapshotProperties(ref snapshotId, out snapProp);
                if (hr == 0)
                {
                    deviceObject = snapProp.m_pwszSnapshotDeviceObject != IntPtr.Zero ? Marshal.PtrToStringUni(snapProp.m_pwszSnapshotDeviceObject) : null;
                    VssFreeSnapshotProperties(ref snapProp);
                    return snapshotId.ToString();
                }
            }
            return null;
        }
        finally
        {
            if (pBackup != null) Marshal.ReleaseComObject(pBackup);
            CoUninitialize();
        }
    }


    static void Main()
    {
        try
        {
            if (!IsAdministrator())
            {
                Console.WriteLine("ERROR: Administrator privileges required");
                return;
            }

            string deviceObject;
            string snapshotId = CreateShadowCopy(@"C:\", out deviceObject);

            if (snapshotId != null)
            {
                Console.WriteLine($"Shadow copy created: {deviceObject}");
            }
            else
            {
                Console.WriteLine("Failed to create shadow copy");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
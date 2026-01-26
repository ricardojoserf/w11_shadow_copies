require "uuid"

VSS_CTX_BACKUP = 0
VSS_BT_FULL = 1

enum VssObjectType
  VSS_OBJECT_UNKNOWN = 0
  VSS_OBJECT_NONE = 1
  VSS_OBJECT_SNAPSHOT_SET = 2
  VSS_OBJECT_SNAPSHOT = 3
  VSS_OBJECT_PROVIDER = 4
  VSS_OBJECT_TYPE_COUNT = 5
end

lib VssLib
  struct GUID
    data1 : UInt32
    data2 : UInt16
    data3 : UInt16
    data4 : UInt8[8]
  end

  struct VssSnapshotProp
    snapshot_id : GUID
    snapshot_set_id : GUID
    snapshots_count : Int32
    _padding1 : Int32
    snapshot_device_object : UInt16*
    original_volume_name : UInt16*
    originating_machine : UInt16*
    service_machine : UInt16*
    exposed_name : UInt16*
    exposed_path : UInt16*
    provider_id : GUID
    snapshot_attributes : Int32
    _padding2 : Int32
    creation_timestamp : Int64
    status : Int32
    _padding3 : Int32
  end

  struct IVssAsyncVtbl
    query_interface : Proc(Void*, GUID*, Void**, Int32)
    add_ref : Proc(Void*, UInt32)
    release : Proc(Void*, UInt32)
    cancel : Proc(Void*, Int32)
    wait : Proc(Void*, UInt32, Int32)
    query_status : Proc(Void*, Int32*, Int32*, Int32)
  end

  struct IVssAsync
    vtbl : IVssAsyncVtbl*
  end

  struct IVssBackupComponentsVtbl
    query_interface : Proc(Void*, GUID*, Void**, Int32)
    add_ref : Proc(Void*, UInt32)
    release : Proc(Void*, UInt32)
    get_writer_components_count : Proc(Void*, UInt32*, Int32)
    get_writer_components : Proc(Void*, UInt32, Void**, Int32)
    initialize_for_backup : Proc(Void*, UInt16*, Int32)
    set_backup_state : Proc(Void*, Int32, Int32, Int32, Int32, Int32)
    initialize_for_restore : Proc(Void*, UInt16*, Int32)
    set_restore_state : Proc(Void*, Int32, Int32)
    gather_writer_metadata : Proc(Void*, Void**, Int32)
    get_writer_metadata_count : Proc(Void*, UInt32*, Int32)
    get_writer_metadata : Proc(Void*, UInt32, GUID*, Void**, Int32)
    free_writer_metadata : Proc(Void*, Int32)
    add_component : Proc(Void*, GUID*, GUID*, Int32, UInt16*, UInt16*, Int32)
    prepare_for_backup : Proc(Void*, Void**, Int32)
    abort_backup : Proc(Void*, Int32)
    gather_writer_status : Proc(Void*, Void**, Int32)
    get_writer_status_count : Proc(Void*, UInt32*, Int32)
    free_writer_status : Proc(Void*, Int32)
    get_writer_status : Proc(Void*, UInt32, GUID*, GUID*, UInt16**, Int32*, Int32*, Int32*, UInt16**, Int32)
    set_backup_succeeded : Proc(Void*, GUID*, GUID*, Int32, UInt16*, UInt16*, Int32, Int32)
    set_backup_options : Proc(Void*, GUID*, Int32, UInt16*, UInt16*, UInt16*, Int32)
    set_selected_for_restore : Proc(Void*, GUID*, Int32, UInt16*, UInt16*, Int32, Int32)
    set_restore_options : Proc(Void*, GUID*, Int32, UInt16*, UInt16*, UInt16*, Int32)
    set_additional_restores : Proc(Void*, GUID*, Int32, UInt16*, UInt16*, Int32, Int32)
    set_previous_backup_stamp : Proc(Void*, GUID*, Int32, UInt16*, UInt16*, UInt16*, Int32)
    save_as_xml : Proc(Void*, UInt16**, Int32)
    backup_complete : Proc(Void*, Void**, Int32)
    add_alternative_location_mapping : Proc(Void*, GUID*, Int32, UInt16*, UInt16*, UInt16*, UInt16*, Int32, UInt16*, Int32)
    add_restore_subcomponent : Proc(Void*, GUID*, Int32, UInt16*, UInt16*, UInt16*, UInt16*, Int32, Int32)
    set_file_restore_status : Proc(Void*, GUID*, Int32, UInt16*, UInt16*, Int32, Int32)
    add_new_target : Proc(Void*, GUID*, Int32, UInt16*, UInt16*, UInt16*, UInt16*, Int32, UInt16*, Int32)
    set_ranges_file_path : Proc(Void*, GUID*, Int32, UInt16*, UInt16*, UInt32, UInt16*, Int32)
    pre_restore : Proc(Void*, Void**, Int32)
    post_restore : Proc(Void*, Void**, Int32)
    set_context : Proc(Void*, Int32, Int32)
    start_snapshot_set : Proc(Void*, GUID*, Int32)
    add_to_snapshot_set : Proc(Void*, UInt16*, GUID*, GUID*, Int32)
    do_snapshot_set : Proc(Void*, Void**, Int32)
    delete_snapshots : Proc(Void*, GUID, Int32, Int32, Int32*, GUID*, Int32)
    import_snapshots : Proc(Void*, Void**, Int32)
    break_snapshot_set : Proc(Void*, GUID*, Int32)
    get_snapshot_properties : Proc(Void*, GUID*, VssSnapshotProp*, Int32)
    query : Proc(Void*, GUID*, Int32, Int32, Void**, Int32)
    is_volume_supported : Proc(Void*, GUID*, UInt16*, Int32*, Int32)
    disable_writer_classes : Proc(Void*, GUID*, UInt32, Int32)
    enable_writer_classes : Proc(Void*, GUID*, UInt32, Int32)
    disable_writer_instances : Proc(Void*, GUID*, UInt32, Int32)
    expose_snapshot : Proc(Void*, GUID*, UInt16*, Int32, UInt16*, UInt16**, Int32)
    revert_to_snapshot : Proc(Void*, GUID*, Int32, Int32)
    query_revert_status : Proc(Void*, UInt16*, Void**, Int32)
  end

  struct IVssBackupComponents
    vtbl : IVssBackupComponentsVtbl*
  end
end

lib Kernel32
  fun load_library = LoadLibraryW(UInt16*) : Void*
  fun get_proc_address = GetProcAddress(Void*, UInt8*) : Void*
end

lib Ole32
  fun co_initialize_ex = CoInitializeEx(Void*, UInt32) : Int32
  fun co_uninitialize = CoUninitialize : Void
end

lib Shell32
  fun is_user_an_admin = IsUserAnAdmin : Int32
end

COINIT_MULTITHREADED = 0x0_u32


def is_administrator? : Bool
  Shell32.is_user_an_admin != 0
end


def guid_to_string(guid : VssLib::GUID) : String
  bytes = StaticArray(UInt8, 16).new(0)
  bytes_ptr = bytes.to_unsafe.as(UInt32*)
  bytes_ptr.value = guid.data1
  (bytes_ptr + 1).as(UInt16*).value = guid.data2
  (bytes_ptr + 1).as(UInt16*)[1] = guid.data3
  8.times { |i| bytes[8 + i] = guid.data4[i] }
  
  uuid = UUID.new(bytes)
  uuid.to_s
end


def wstring_to_string(ptr : UInt16*) : String
  return "" if ptr.null?
  
  size = 0
  temp = ptr
  while temp.value != 0
    size += 1
    temp += 1
  end
  
  String.build do |str|
    i = 0
    while i < size
      char = ptr[i].to_i32
      
      if char >= 0xD800 && char <= 0xDBFF && i + 1 < size
        low = ptr[i + 1].to_i32
        if low >= 0xDC00 && low <= 0xDFFF
          codepoint = 0x10000 + ((char - 0xD800) << 10) + (low - 0xDC00)
          str << codepoint.chr
          i += 2
          next
        end
      end
      
      str << char.chr if char < 0xD800 || char > 0xDFFF
      i += 1
    end
  end
end


def create_null_guid : VssLib::GUID
  VssLib::GUID.new(
    data1: 0_u32,
    data2: 0_u16,
    data3: 0_u16,
    data4: StaticArray(UInt8, 8).new(0_u8)
  )
end


def load_create_vss_function : Proc(Pointer(Pointer(VssLib::IVssBackupComponents)), Int32)?
  dll_name = "VssApi.dll".to_utf16
  vss_handle = Kernel32.load_library(dll_name.to_unsafe)
  
  return nil if vss_handle.null?
  
  possible_names = [
    "CreateVssBackupComponentsInternal",
    "CreateVssBackupComponents",
    "?CreateVssBackupComponents@@YAJPEAPEAVIVssBackupComponents@@@Z"
  ]
  
  possible_names.each do |name|
    func_ptr = Kernel32.get_proc_address(vss_handle, name.to_unsafe)
    unless func_ptr.null?
      return Proc(Pointer(Pointer(VssLib::IVssBackupComponents)), Int32).new(func_ptr, Pointer(Void).null)
    end
  end
  
  nil
end


def load_vss_free_function : Proc(Pointer(VssLib::VssSnapshotProp), Nil)?
  dll_name = "VssApi.dll".to_utf16
  vss_handle = Kernel32.load_library(dll_name.to_unsafe)
  
  return nil if vss_handle.null?
  
  func_ptr = Kernel32.get_proc_address(vss_handle, "VssFreeSnapshotProperties".to_unsafe)
  return nil if func_ptr.null?
  
  Proc(Pointer(VssLib::VssSnapshotProp), Nil).new(func_ptr, Pointer(Void).null)
end


def create_shadow_copy(volume_path : String) : Tuple(String?, String?)
  com_initialized = false
  
  begin
    hr = Ole32.co_initialize_ex(nil, COINIT_MULTITHREADED)
    if hr == 0 || hr == 1
      com_initialized = true
    elsif hr != -2147417850
      puts "Error initializing COM. Error: 0x#{hr.to_s(16)}"
      return {nil, nil}
    end
  rescue ex
    puts "Exception initializing COM: #{ex.message}"
    return {nil, nil}
  end
  
  create_func = load_create_vss_function
  unless create_func
    puts "Error: Could not load CreateVssBackupComponents from VssApi.dll"
    return {nil, nil}
  end
  
  free_func = load_vss_free_function
  unless free_func
    puts "Error: Could not load VssFreeSnapshotProperties from VssApi.dll"
    return {nil, nil}
  end
  
  backup : VssLib::IVssBackupComponents* = Pointer(VssLib::IVssBackupComponents).null
  
  begin
    hr = create_func.call(pointerof(backup))
    if hr != 0
      puts "Error creating VSS components. HRESULT: 0x#{hr.to_s(16)}"
      return {nil, nil}
    end
    
    if backup.null?
      puts "Error: CreateVssBackupComponents returned NULL"
      return {nil, nil}
    end
    
    vtbl = backup.value.vtbl
    hr = vtbl.value.initialize_for_backup.call(backup.as(Void*), Pointer(UInt16).null)
    if hr != 0
      puts "Error in InitializeForBackup. HRESULT: 0x#{hr.to_s(16)}"
      return {nil, nil}
    end
    
    guid_null = create_null_guid
    volume_utf16 = volume_path.to_utf16
    
    is_supported = 0
    hr = vtbl.value.is_volume_supported.call(
      backup.as(Void*),
      pointerof(guid_null),
      volume_utf16.to_unsafe,
      pointerof(is_supported)
    )
    
    if hr != 0 || is_supported == 0
      puts "Volume #{volume_path} is not supported for shadow copies."
      return {nil, nil}
    end
    
    hr = vtbl.value.set_context.call(backup.as(Void*), VSS_CTX_BACKUP)
    if hr != 0
      puts "Error in SetContext. HRESULT: 0x#{hr.to_s(16)}"
      return {nil, nil}
    end
    
    hr = vtbl.value.set_backup_state.call(backup.as(Void*), 0, 0, VSS_BT_FULL, 0)
    if hr != 0
      puts "Error in SetBackupState. HRESULT: 0x#{hr.to_s(16)}"
      return {nil, nil}
    end
    
    async_metadata : VssLib::IVssAsync* = Pointer(VssLib::IVssAsync).null
    hr = vtbl.value.gather_writer_metadata.call(backup.as(Void*), pointerof(async_metadata).as(Pointer(Pointer(Void))))
    if hr == 0 && !async_metadata.null?
      async_vtbl = async_metadata.value.vtbl
      async_vtbl.value.wait.call(async_metadata.as(Void*), 0xFFFFFFFF_u32)
      async_vtbl.value.release.call(async_metadata.as(Void*))
    end
    
    snapshot_set_id = VssLib::GUID.new
    hr = vtbl.value.start_snapshot_set.call(backup.as(Void*), pointerof(snapshot_set_id))
    if hr != 0
      puts "Error in StartSnapshotSet. HRESULT: 0x#{hr.to_s(16)}"
      return {nil, nil}
    end
    
    snapshot_id = VssLib::GUID.new
    hr = vtbl.value.add_to_snapshot_set.call(
      backup.as(Void*),
      volume_utf16.to_unsafe,
      pointerof(guid_null),
      pointerof(snapshot_id)
    )
    if hr != 0
      puts "Error in AddToSnapshotSet. HRESULT: 0x#{hr.to_s(16)}"
      return {nil, nil}
    end
    
    async_prepare : VssLib::IVssAsync* = Pointer(VssLib::IVssAsync).null
    hr = vtbl.value.prepare_for_backup.call(backup.as(Void*), pointerof(async_prepare).as(Pointer(Pointer(Void))))
    if hr == 0 && !async_prepare.null?
      async_vtbl = async_prepare.value.vtbl
      async_vtbl.value.wait.call(async_prepare.as(Void*), 0xFFFFFFFF_u32)
      async_vtbl.value.release.call(async_prepare.as(Void*))
    end
    
    async_snapshot : VssLib::IVssAsync* = Pointer(VssLib::IVssAsync).null
    hr = vtbl.value.do_snapshot_set.call(backup.as(Void*), pointerof(async_snapshot).as(Pointer(Pointer(Void))))
    if hr == 0 && !async_snapshot.null?
      async_vtbl = async_snapshot.value.vtbl
      hr = async_vtbl.value.wait.call(async_snapshot.as(Void*), 0xFFFFFFFF_u32)
      async_vtbl.value.release.call(async_snapshot.as(Void*))
    end
    
    if hr == 0
      snap_prop = VssLib::VssSnapshotProp.new
      hr = vtbl.value.get_snapshot_properties.call(
        backup.as(Void*),
        pointerof(snapshot_id),
        pointerof(snap_prop)
      )
      
      if hr == 0
        device_object = ""
        unless snap_prop.snapshot_device_object.null?
          device_object = wstring_to_string(snap_prop.snapshot_device_object)
        end
        
        snapshot_id_str = guid_to_string(snapshot_id)
        free_func.call(pointerof(snap_prop))
        
        return {snapshot_id_str, device_object}
      else
        puts "Error in GetSnapshotProperties. HRESULT: 0x#{hr.to_s(16)}"
      end
    else
      puts "Error in DoSnapshotSet. HRESULT: 0x#{hr.to_s(16)}"
    end
    
    return {nil, nil}
    
  rescue ex
    puts "Error: #{ex.message}"
    puts ex.backtrace.join("\n")
    return {nil, nil}
  ensure
    unless backup.null?
      vtbl = backup.value.vtbl
      vtbl.value.release.call(backup.as(Void*))
    end
    
    Ole32.co_uninitialize if com_initialized
  end
end


def main
  unless is_administrator?
    puts "ERROR: This program requires administrator privileges."
    return
  end
  
  volume_path = "C:\\"
  
  if ARGV.size > 0
    volume_path = ARGV[0]
    volume_path += "\\" unless volume_path.ends_with?("\\")
  end
  
  snapshot_id, device_object = create_shadow_copy(volume_path)
  
  if snapshot_id && device_object && !device_object.empty?
    puts "Shadow copy created: #{device_object}"
  else
    puts "Failed to create shadow copy"
  end
end


main
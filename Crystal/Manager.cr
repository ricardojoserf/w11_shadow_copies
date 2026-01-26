require "uuid"

VSS_CTX_BACKUP = 0
VSS_CTX_ALL = 0xffffffff_u32
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

  union VssObjectUnion
    snap : VssSnapshotProp
  end

  struct VssObjectProp
    type : Int32
    _padding : Int32
    obj : VssObjectUnion
  end

  struct IVssEnumObjectVtbl
    query_interface : Proc(Void*, GUID*, Void**, Int32)
    add_ref : Proc(Void*, UInt32)
    release : Proc(Void*, UInt32)
    next : Proc(Void*, UInt32, VssObjectProp*, UInt32*, Int32)
    skip : Proc(Void*, UInt32, Int32)
    reset : Proc(Void*, Int32)
    clone : Proc(Void*, Void**, Int32)
  end

  struct IVssEnumObject
    vtbl : IVssEnumObjectVtbl*
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


def init_com : Bool?
  hr = Ole32.co_initialize_ex(nil, COINIT_MULTITHREADED)
  if hr == 0 || hr == 1
    return true
  elsif hr != -2147417850
    puts "Error initializing COM. Error: 0x#{hr.to_s(16)}"
    return nil
  end
  false
end


def list_shadow_copies : Nil
  com_init = init_com
  return if com_init.nil?
  
  create_func = load_create_vss_function
  return unless create_func
  
  free_func = load_vss_free_function
  return unless free_func
  
  backup : VssLib::IVssBackupComponents* = Pointer(VssLib::IVssBackupComponents).null
  enum_obj : VssLib::IVssEnumObject* = Pointer(VssLib::IVssEnumObject).null
  
  begin
    hr = create_func.call(pointerof(backup))
    return if hr != 0 || backup.null?
    
    vtbl = backup.value.vtbl
    hr = vtbl.value.initialize_for_backup.call(backup.as(Void*), Pointer(UInt16).null)
    return if hr != 0
    
    vss_ctx_all_as_signed = VSS_CTX_ALL.unsafe_as(Int32)
    hr = vtbl.value.set_context.call(backup.as(Void*), vss_ctx_all_as_signed)
    if hr != 0
      hr = vtbl.value.set_context.call(backup.as(Void*), VSS_CTX_BACKUP)
      return if hr != 0
    end
    
    guid_null = create_null_guid
    
    hr = vtbl.value.query.call(
      backup.as(Void*),
      pointerof(guid_null),
      VssObjectType::VSS_OBJECT_NONE.value,
      VssObjectType::VSS_OBJECT_SNAPSHOT.value,
      pointerof(enum_obj).as(Pointer(Pointer(Void)))
    )
    
    if hr != 0 || enum_obj.null?
      puts "No shadow copies found"
      return
    end
    
    enum_vtbl = enum_obj.value.vtbl
    count = 0
    
    loop do
      prop = VssLib::VssObjectProp.new
      fetched = 0_u32
      
      hr = enum_vtbl.value.next.call(enum_obj.as(Void*), 1_u32, pointerof(prop), pointerof(fetched))
      break if hr != 0 || fetched == 0
      
      if prop.type == VssObjectType::VSS_OBJECT_SNAPSHOT.value
        count += 1
        puts "=" * 47
        puts "Shadow Copy ##{count}"
        puts "=" * 47
        
        snap = prop.obj.snap
        
        puts "ID: {#{guid_to_string(snap.snapshot_id)}}"
        puts "Set ID: {#{guid_to_string(snap.snapshot_set_id)}}"
        
        unless snap.snapshot_device_object.null?
          device = wstring_to_string(snap.snapshot_device_object)
          puts "Device Object: #{device}"
        end
        
        unless snap.original_volume_name.null?
          volume = wstring_to_string(snap.original_volume_name)
          puts "Original Volume: #{volume}"
        end
        
        unless snap.originating_machine.null?
          machine = wstring_to_string(snap.originating_machine)
          puts "Originating Machine: #{machine}"
        end
        
        timestamp_seconds = snap.creation_timestamp // 10000000_i64 - 11644473600_i64
        timestamp = Time.unix(timestamp_seconds)
        puts "Creation Date: #{timestamp}"
        puts "Attributes: 0x#{snap.snapshot_attributes.to_s(16).upcase}"
        puts "Status: #{snap.status}"
        puts "Provider ID: {#{guid_to_string(snap.provider_id)}}"
        puts
        
        free_func.call(pointerof(snap))
      end
    end
    
    if count == 0
      puts "No shadow copies found"
    else
      puts "Total: #{count} shadow copies"
    end
    
  ensure
    unless enum_obj.null?
      enum_vtbl = enum_obj.value.vtbl
      enum_vtbl.value.release.call(enum_obj.as(Void*))
    end
    
    unless backup.null?
      vtbl = backup.value.vtbl
      vtbl.value.release.call(backup.as(Void*))
    end
    
    Ole32.co_uninitialize if com_init
  end
end


def create_shadow_copy(volume_path : String) : Tuple(String?, String?)
  com_init = init_com
  return {nil, nil} if com_init.nil?
  
  create_func = load_create_vss_function
  return {nil, nil} unless create_func
  
  free_func = load_vss_free_function
  return {nil, nil} unless free_func
  
  backup : VssLib::IVssBackupComponents* = Pointer(VssLib::IVssBackupComponents).null
  
  begin
    hr = create_func.call(pointerof(backup))
    return {nil, nil} if hr != 0 || backup.null?
    
    vtbl = backup.value.vtbl
    hr = vtbl.value.initialize_for_backup.call(backup.as(Void*), Pointer(UInt16).null)
    return {nil, nil} if hr != 0
    
    guid_null = create_null_guid
    
    hr = vtbl.value.set_context.call(backup.as(Void*), VSS_CTX_BACKUP)
    return {nil, nil} if hr != 0
    
    hr = vtbl.value.set_backup_state.call(backup.as(Void*), 0, 0, VSS_BT_FULL, 0)
    return {nil, nil} if hr != 0
    
    async_metadata : VssLib::IVssAsync* = Pointer(VssLib::IVssAsync).null
    hr = vtbl.value.gather_writer_metadata.call(backup.as(Void*), pointerof(async_metadata).as(Pointer(Pointer(Void))))
    if hr == 0 && !async_metadata.null?
      async_vtbl = async_metadata.value.vtbl
      async_vtbl.value.wait.call(async_metadata.as(Void*), 0xFFFFFFFF_u32)
      async_vtbl.value.release.call(async_metadata.as(Void*))
    end
    
    snapshot_set_id = VssLib::GUID.new
    hr = vtbl.value.start_snapshot_set.call(backup.as(Void*), pointerof(snapshot_set_id))
    return {nil, nil} if hr != 0
    
    volume_utf16 = volume_path.to_utf16
    snapshot_id = VssLib::GUID.new
    hr = vtbl.value.add_to_snapshot_set.call(
      backup.as(Void*),
      volume_utf16.to_unsafe,
      pointerof(guid_null),
      pointerof(snapshot_id)
    )
    return {nil, nil} if hr != 0
    
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
    else
      hr = -1
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
      end
    end
    
    return {nil, nil}
    
  ensure
    unless backup.null?
      vtbl = backup.value.vtbl
      vtbl.value.release.call(backup.as(Void*))
    end
    
    Ole32.co_uninitialize if com_init
  end
end


def extract_snapshot_id(device_path : String) : VssLib::GUID?
  com_init = init_com
  return nil if com_init.nil?
  
  create_func = load_create_vss_function
  return nil unless create_func
  
  free_func = load_vss_free_function
  return nil unless free_func
  
  backup : VssLib::IVssBackupComponents* = Pointer(VssLib::IVssBackupComponents).null
  enum_obj : VssLib::IVssEnumObject* = Pointer(VssLib::IVssEnumObject).null
  
  begin
    hr = create_func.call(pointerof(backup))
    return nil if hr != 0 || backup.null?
    
    vtbl = backup.value.vtbl
    hr = vtbl.value.initialize_for_backup.call(backup.as(Void*), Pointer(UInt16).null)
    return nil if hr != 0
    
    vss_ctx_all_as_signed = VSS_CTX_ALL.unsafe_as(Int32)
    hr = vtbl.value.set_context.call(backup.as(Void*), vss_ctx_all_as_signed)
    if hr != 0
      hr = vtbl.value.set_context.call(backup.as(Void*), VSS_CTX_BACKUP)
      return nil if hr != 0
    end
    
    guid_null = create_null_guid
    
    hr = vtbl.value.query.call(
      backup.as(Void*),
      pointerof(guid_null),
      VssObjectType::VSS_OBJECT_NONE.value,
      VssObjectType::VSS_OBJECT_SNAPSHOT.value,
      pointerof(enum_obj).as(Pointer(Pointer(Void)))
    )
    
    return nil if hr != 0 || enum_obj.null?
    
    enum_vtbl = enum_obj.value.vtbl
    device_path_norm = device_path.rstrip('\\').downcase
    
    loop do
      prop = VssLib::VssObjectProp.new
      fetched = 0_u32
      
      hr = enum_vtbl.value.next.call(enum_obj.as(Void*), 1_u32, pointerof(prop), pointerof(fetched))
      break if hr != 0 || fetched == 0
      
      if prop.type == VssObjectType::VSS_OBJECT_SNAPSHOT.value
        snap = prop.obj.snap
        
        unless snap.snapshot_device_object.null?
          device_object = wstring_to_string(snap.snapshot_device_object)
          device_object_norm = device_object.rstrip('\\').downcase
          
          if device_object_norm == device_path_norm
            snapshot_id = snap.snapshot_id
            free_func.call(pointerof(snap))
            return snapshot_id
          end
        end
        
        free_func.call(pointerof(snap))
      end
    end
    
    return nil
    
  ensure
    unless enum_obj.null?
      enum_vtbl = enum_obj.value.vtbl
      enum_vtbl.value.release.call(enum_obj.as(Void*))
    end
    
    unless backup.null?
      vtbl = backup.value.vtbl
      vtbl.value.release.call(backup.as(Void*))
    end
    
    Ole32.co_uninitialize if com_init
  end
end


def delete_shadow_copy(device_path : String) : Bool
  com_init = init_com
  return false if com_init.nil?
  
  snapshot_id = extract_snapshot_id(device_path)
  unless snapshot_id
    puts "Shadow copy not found: #{device_path}"
    return false
  end
  
  create_func = load_create_vss_function
  return false unless create_func
  
  backup : VssLib::IVssBackupComponents* = Pointer(VssLib::IVssBackupComponents).null
  
  begin
    hr = create_func.call(pointerof(backup))
    return false if hr != 0 || backup.null?
    
    vtbl = backup.value.vtbl
    hr = vtbl.value.initialize_for_backup.call(backup.as(Void*), Pointer(UInt16).null)
    return false if hr != 0
    
    vss_ctx_all_as_signed = VSS_CTX_ALL.unsafe_as(Int32)
    hr = vtbl.value.set_context.call(backup.as(Void*), vss_ctx_all_as_signed)
    if hr != 0
      hr = vtbl.value.set_context.call(backup.as(Void*), VSS_CTX_BACKUP)
      return false if hr != 0
    end
    
    deleted_snapshots = 0_i32
    nondeleted_snapshot_id = VssLib::GUID.new
    
    hr = vtbl.value.delete_snapshots.call(
      backup.as(Void*),
      snapshot_id,
      VssObjectType::VSS_OBJECT_SNAPSHOT.value,
      1,
      pointerof(deleted_snapshots),
      pointerof(nondeleted_snapshot_id)
    )
    
    if hr != 0
      puts "Delete error. HRESULT: 0x#{hr.to_s(16)}"
      return false
    end
    
    return deleted_snapshots > 0
    
  ensure
    unless backup.null?
      vtbl = backup.value.vtbl
      vtbl.value.release.call(backup.as(Void*))
    end
    
    Ole32.co_uninitialize if com_init
  end
end


def show_usage
  puts "Usage:"
  puts "  manager create [volume]       - Create shadow copy (default volume: C:\\)"
  puts "  manager list                  - List all shadow copies"
  puts "  manager delete <device_path>  - Delete a shadow copy"
  puts ""
  puts "Examples:"
  puts "  manager create"
  puts "  manager create D:\\"
  puts "  manager list"
  puts "  manager delete \\\\?\\GLOBALROOT\\Device\\HarddiskVolumeShadowCopy12"
end


def main
  unless is_administrator?
    puts "ERROR: This program requires administrator privileges."
    exit 1
  end
  
  if ARGV.empty?
    show_usage
    exit 1
  end
  
  operation = ARGV[0].downcase
  
  case operation
  when "create"
    volume = ARGV.size > 1 ? ARGV[1] : "C:\\"
    volume += "\\" unless volume.ends_with?("\\")
    
    snapshot_id, device_object = create_shadow_copy(volume)
    if snapshot_id && device_object && !device_object.empty?
      puts "Shadow copy created: #{device_object}"
    else
      puts "Failed to create shadow copy"
      exit 1
    end
    
  when "list"
    list_shadow_copies
    
  when "delete"
    if ARGV.size < 2
      puts "ERROR: Shadow copy device path required for delete operation"
      puts "Example: manager delete \\\\?\\GLOBALROOT\\Device\\HarddiskVolumeShadowCopy12"
      exit 1
    end
    
    snapshot = ARGV[1]
    if delete_shadow_copy(snapshot)
      puts "Shadow copy deleted successfully"
    else
      puts "Failed to delete shadow copy"
      exit 1
    end
    
  else
    puts "ERROR: Invalid operation '#{operation}'"
    puts ""
    show_usage
    exit 1
  end
end


main
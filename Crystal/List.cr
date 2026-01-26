require "uuid"

VSS_CTX_BACKUP = 0
VSS_CTX_ALL = 0xffffffff_u32

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


def list_shadow_copies
  com_initialized = false
  
  begin
    hr = Ole32.co_initialize_ex(nil, COINIT_MULTITHREADED)
    if hr == 0 || hr == 1
      com_initialized = true
    elsif hr != -2147417850
      puts "Error initializing COM. Error: 0x#{hr.to_s(16)}"
      return
    end
  rescue ex
    puts "Exception initializing COM: #{ex.message}"
    return
  end
  
  create_func = load_create_vss_function
  unless create_func
    puts "Error: Could not load CreateVssBackupComponents from VssApi.dll"
    return
  end
  
  free_func = load_vss_free_function
  unless free_func
    puts "Error: Could not load VssFreeSnapshotProperties from VssApi.dll"
    return
  end
  
  backup : VssLib::IVssBackupComponents* = Pointer(VssLib::IVssBackupComponents).null
  enum_obj : VssLib::IVssEnumObject* = Pointer(VssLib::IVssEnumObject).null
  
  begin
    hr = create_func.call(pointerof(backup))
    if hr != 0
      puts "Error creating VSS components. HRESULT: 0x#{hr.to_s(16)}"
      return
    end
    
    if backup.null?
      puts "Error: CreateVssBackupComponents returned NULL"
      return
    end
    
    vtbl = backup.value.vtbl
    hr = vtbl.value.initialize_for_backup.call(backup.as(Void*), Pointer(UInt16).null)
    if hr != 0
      puts "Error in InitializeForBackup. HRESULT: 0x#{hr.to_s(16)}"
      case hr
      when 0x80042302
        puts "  -> VSS_E_UNEXPECTED: Unexpected VSS error"
      when 0x8004230C
        puts "  -> VSS_E_BAD_STATE: VSS in incorrect state"
      when 0x80042308
        puts "  -> VSS_E_VOLUME_NOT_SUPPORTED_BY_PROVIDER: Volume not supported"
      end
      return
    end
    
    vss_ctx_all_as_signed = VSS_CTX_ALL.unsafe_as(Int32)
    hr = vtbl.value.set_context.call(backup.as(Void*), vss_ctx_all_as_signed)
    if hr != 0
      hr = vtbl.value.set_context.call(backup.as(Void*), VSS_CTX_BACKUP)
      if hr != 0
        puts "Error in SetContext. HRESULT: 0x#{hr.to_s(16)}"
        return
      end
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
      puts "No shadow copies found on the system."
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
      puts "No shadow copies found on the system."
    else
      puts "Total: #{count} shadow copies found"
    end
    
  rescue ex
    puts "Error: #{ex.message}"
    puts ex.backtrace.join("\n")
  ensure
    unless enum_obj.null?
      enum_vtbl = enum_obj.value.vtbl
      enum_vtbl.value.release.call(enum_obj.as(Void*))
    end
    
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
  
  list_shadow_copies
end

main
require "uuid"

VSS_CTX_BACKUP = 0
VSS_CTX_ALL = -1_i32

enum VssObjectType : Int32
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

struct VssApi
  getter create_func : Proc(Pointer(Pointer(VssLib::IVssBackupComponents)), Int32)
  getter free_func : Proc(Pointer(VssLib::VssSnapshotProp), Nil)
  
  def initialize(@create_func, @free_func)
  end
  
  def self.load : self?
    dll_name = "VssApi.dll".to_utf16
    handle = Kernel32.load_library(dll_name.to_unsafe)
    return nil if handle.null?
    
    create_ptr = nil.as(Void*?)
    ["CreateVssBackupComponentsInternal", "CreateVssBackupComponents", 
     "?CreateVssBackupComponents@@YAJPEAPEAVIVssBackupComponents@@@Z"].each do |n|
      ptr = Kernel32.get_proc_address(handle, n.to_unsafe)
      unless ptr.null?
        create_ptr = ptr
        break
      end
    end
    
    return nil if create_ptr.nil? || create_ptr.null?
    
    free_ptr = Kernel32.get_proc_address(handle, "VssFreeSnapshotProperties".to_unsafe)
    return nil if free_ptr.null?
    
    new(
      Proc(Pointer(Pointer(VssLib::IVssBackupComponents)), Int32).new(create_ptr.not_nil!, Pointer(Void).null),
      Proc(Pointer(VssLib::VssSnapshotProp), Nil).new(free_ptr, Pointer(Void).null)
    )
  end
end

@[AlwaysInline]
def guid_to_string(guid : VssLib::GUID) : String
  String.build(36) do |io|
    io << guid.data1.to_s(16, precision: 8) << '-'
    io << guid.data2.to_s(16, precision: 4) << '-'
    io << guid.data3.to_s(16, precision: 4) << '-'
    io << guid.data4[0].to_s(16, precision: 2) << guid.data4[1].to_s(16, precision: 2) << '-'
    (2..7).each { |i| io << guid.data4[i].to_s(16, precision: 2) }
  end
end

@[AlwaysInline]
def wstring_to_string(ptr : UInt16*) : String
  return "" if ptr.null?
  
  String.build do |io|
    i = 0
    loop do
      char = ptr[i].to_i32
      break if char == 0
      
      if char >= 0xD800 && char <= 0xDBFF && ptr[i + 1] != 0
        low = ptr[i + 1].to_i32
        if low >= 0xDC00 && low <= 0xDFFF
          io << (0x10000 + ((char - 0xD800) << 10) + (low - 0xDC00)).chr
          i += 2
          next
        end
      end
      
      io << char.chr if char < 0xD800 || char > 0xDFFF
      i += 1
    end
  end
end


def list_shadow_copies : Nil
  return puts "ERROR: Administrator privileges required" unless Shell32.is_user_an_admin != 0
  
  hr = Ole32.co_initialize_ex(nil, COINIT_MULTITHREADED)
  com_initialized = hr == 0 || hr == 1
  return puts "Error initializing COM: 0x#{hr.to_s(16)}" if !com_initialized && hr != -2147417850
  
  api = VssApi.load
  return puts "Error: Failed to load VssApi.dll" unless api
  
  backup : VssLib::IVssBackupComponents* = Pointer(VssLib::IVssBackupComponents).null
  enum_obj : VssLib::IVssEnumObject* = Pointer(VssLib::IVssEnumObject).null
  
  begin
    return puts "Error creating VSS components" if api.create_func.call(pointerof(backup)) != 0 || backup.null?
    
    v = backup.value.vtbl.value
    return puts "Error in InitializeForBackup" if v.initialize_for_backup.call(backup.as(Void*), Pointer(UInt16).null) != 0
    
    hr = v.set_context.call(backup.as(Void*), VSS_CTX_ALL)
    return puts "Error in SetContext" if hr != 0 && v.set_context.call(backup.as(Void*), VSS_CTX_BACKUP) != 0
    
    guid_null = VssLib::GUID.new(data1: 0_u32, data2: 0_u16, data3: 0_u16, data4: StaticArray(UInt8, 8).new(0_u8))
    
    hr = v.query.call(backup.as(Void*), pointerof(guid_null), VssObjectType::VSS_OBJECT_NONE.value,
                      VssObjectType::VSS_OBJECT_SNAPSHOT.value, pointerof(enum_obj).as(Pointer(Pointer(Void))))
    
    return puts "No shadow copies found on the system." if hr != 0 || enum_obj.null?
    
    ev = enum_obj.value.vtbl.value
    count = 0
    
    loop do
      prop = VssLib::VssObjectProp.new
      fetched = 0_u32
      break if ev.next.call(enum_obj.as(Void*), 1_u32, pointerof(prop), pointerof(fetched)) != 0 || fetched == 0
      
      next unless prop.type == VssObjectType::VSS_OBJECT_SNAPSHOT.value
      
      count += 1
      snap = prop.obj.snap
      
      puts "=" * 47, "Shadow Copy ##{count}", "=" * 47,
           "ID: {#{guid_to_string(snap.snapshot_id)}}",
           "Set ID: {#{guid_to_string(snap.snapshot_set_id)}}"
      
      puts "Device Object: #{wstring_to_string(snap.snapshot_device_object)}" unless snap.snapshot_device_object.null?
      puts "Original Volume: #{wstring_to_string(snap.original_volume_name)}" unless snap.original_volume_name.null?
      puts "Originating Machine: #{wstring_to_string(snap.originating_machine)}" unless snap.originating_machine.null?
      
      puts "Creation Date: #{Time.unix(snap.creation_timestamp // 10000000_i64 - 11644473600_i64)}",
           "Attributes: 0x#{snap.snapshot_attributes.to_s(16).upcase}",
           "Status: #{snap.status}",
           "Provider ID: {#{guid_to_string(snap.provider_id)}}\n"
      
      api.free_func.call(pointerof(snap))
    end
    
    puts count == 0 ? "No shadow copies found on the system." : "Total: #{count} shadow copies found"
    
  ensure
    ev.release.call(enum_obj.as(Void*)) if enum_obj && (ev = enum_obj.value.vtbl.value)
    v.release.call(backup.as(Void*)) if backup && (v = backup.value.vtbl.value)
    Ole32.co_uninitialize if com_initialized
  end
end


list_shadow_copies
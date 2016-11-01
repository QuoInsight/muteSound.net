/*
  © (ɔ) QuoInsight
  ref: http://stackoverflow.com/questions/21355891/change-audio-level-from-powershell
  ref: http://stackoverflow.com/questions/33872895/detect-if-headphones-are-plugged-in-or-not-via-c-sharp
*/

using System;
using System.Runtime.InteropServices;

using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace myNameSpace {
  class myClass {

    public struct DEVICE_SUMMARY {
      public string FriendlyName;
      public bool isMuted;
      public int dataFlow;
      public int role;
    }

    static void Main(string[] args) {

      Console.Error.WriteLine();
      Console.Error.WriteLine("Syntax: audio.cs.exe [setMute] [dataFlow]");  // args[0] args[1]
      Console.Error.WriteLine("  [setMute] 0=mute; 1=unmute; *=followDefault");
      Console.Error.WriteLine("  [dataFlow] 0=render/playback; 1=capture/recording; 2=All");
      Console.Error.WriteLine();

      bool defaultDeviceMuted;
      string setMute = ( args.Length > 0 ) ? args[0] : "";
      int dataFlow = 2; if ( args.Length > 1 ) {
        switch (args[1]) {
          case "0": case "1": case "2":
            dataFlow = Convert.ToInt32(args[1]);
            break;
        }
      }

      System.Collections.Generic.List<DEVICE_SUMMARY> defaultDevices
        = new System.Collections.Generic.List<DEVICE_SUMMARY>();
      if (dataFlow==2) {
        DEVICE_SUMMARY defaultDevice = getDefaultDeviceSummary(0,1);
        defaultDeviceMuted = defaultDevice.isMuted;
        defaultDevices.Add(defaultDevice);
        defaultDevices.Add(getDefaultDeviceSummary(1,1));
      } else {
        DEVICE_SUMMARY defaultDevice = getDefaultDeviceSummary(dataFlow,1);
        defaultDeviceMuted = defaultDevice.isMuted;
        defaultDevices.Add(defaultDevice);
      }

      if (setMute=="*") {
        Console.Error.WriteLine("defaultDeviceMuted: " + defaultDeviceMuted + "\n");
        setMute = (defaultDeviceMuted) ? "0" : "1";
      }

      listDevices(dataFlow, setMute, defaultDevices);
      Console.Error.WriteLine("OK");
      return;

    } // Main()

    [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")] class MMDeviceEnumeratorComObject { }

      public static void listDevices(int dataFlow, string setMute, System.Collections.Generic.List<DEVICE_SUMMARY> defaultDevices) {
        var enumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
        uint deviceCount=0, i=0;

        System.Console.Write("EnumAudioEndpoints: ...");
        IMMDeviceCollection deviceCollection;  enumerator.EnumAudioEndpoints(dataFlow, 1, out deviceCollection);
        // EDataFlow: 0=eRender;1=eCapture;2=eAll | 1=DEVICE_STATE_ACTIVE

        deviceCollection.GetCount(out deviceCount);  System.Console.WriteLine(" deviceCount=" + deviceCount); //return;
        for (i=0; i<deviceCount; i++) {
          System.Console.Write("\n" + i + ":");
          IMMDevice dev = null;  deviceCollection.Item(i, out dev);
          if ( !Object.ReferenceEquals(null, dev) ) {
            IMMEndpoint ep = (IMMEndpoint) dev;  // Marshal.QueryInterface shouldn't be necessary, just use cast instead [ http://stackoverflow.com/questions/5077172/how-do-i-use-marshal-queryinterface ]
              int dataFlow1;  ep.GetDataFlow(out dataFlow1);
              System.Console.WriteLine(" GetDataFlow: [" + dataFlow1 + "] " + ( (dataFlow1==0)?"Playback":"Recording" ) );

            DEVICE_SUMMARY defaultDevice = new DEVICE_SUMMARY();
            foreach (var dev0 in defaultDevices) {
              if (dev0.dataFlow==dataFlow1) {
                defaultDevice = dev0;
                break;
              }
            }

            listDeviceProperties(dev, defaultDevice);
              if (setMute=="0"||setMute=="1") {
                IAudioEndpointVolume epv = null;
                var epvid = typeof(IAudioEndpointVolume).GUID;
                Marshal.ThrowExceptionForHR(dev.Activate(ref epvid, /*CLSCTX_ALL*/ 23, 0, out epv));
                epv.SetMute((setMute=="0"), System.Guid.Empty);  System.Console.WriteLine(" SetMute: " + (setMute=="0"));
              }
          }
          Marshal.ReleaseComObject(dev);
        }
        Marshal.ReleaseComObject(deviceCollection);

        System.Console.WriteLine();
      }

      public static void listDeviceProperties(IMMDevice dev, DEVICE_SUMMARY defaultDevice) {
        IPropertyStore propertyStore;
          dev.OpenPropertyStore(0/*STGM_READ*/, out propertyStore);
          PROPVARIANT property;  propertyStore.GetValue(ref PropertyKey.PKEY_Device_EnumeratorName, out property);
          System.Console.WriteLine(" EnumeratorName: " + (string)property.Value);
          propertyStore.GetValue(ref PropertyKey.PKEY_Device_FriendlyName, out property);
          System.Console.WriteLine(" FriendlyName: " + (string)property.Value);
          if (!defaultDevice.Equals(null) && defaultDevice.FriendlyName==(string)property.Value) {
            System.Console.WriteLine(" **Default: True**");
          }
        Marshal.ReleaseComObject(propertyStore);

        IAudioEndpointVolume epv = null;
          var epvid = typeof(IAudioEndpointVolume).GUID;
          Marshal.ThrowExceptionForHR(dev.Activate(ref epvid, /*CLSCTX_ALL*/ 23, 0, out epv));

          bool mute;  epv.GetMute(out mute);  if (mute) System.Console.WriteLine(" **GetMute: " + mute + "**");
          float vol;  epv.GetMasterVolumeLevelScalar(out vol);  System.Console.WriteLine(" GetMasterVolumeLevelScalar: " + vol);
      }

      public static DEVICE_SUMMARY getDefaultDeviceSummary(int dataFlow, int role) {
        DEVICE_SUMMARY dev1 = new DEVICE_SUMMARY();  dev1.dataFlow = dataFlow;  dev1.role = role;
        var enumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
        IMMDevice dev = null;  Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(dataFlow, role, out dev));
          IPropertyStore propertyStore;  dev.OpenPropertyStore(0/*STGM_READ*/, out propertyStore);
            PROPVARIANT property;  propertyStore.GetValue(ref PropertyKey.PKEY_Device_FriendlyName, out property);
            dev1.FriendlyName = (string)property.Value;
          Marshal.ReleaseComObject(propertyStore);
          IAudioEndpointVolume epv = null;
            var epvid = typeof(IAudioEndpointVolume).GUID;
            Marshal.ThrowExceptionForHR(dev.Activate(ref epvid, /*CLSCTX_ALL*/ 23, 0, out epv));
            epv.GetMute(out dev1.isMuted);
        Marshal.ReleaseComObject(dev);
        return dev1;
      }

      public static void listDefaultDevice() {
        var enumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
        IMMDevice dev = null;

        System.Console.Write("GetDefaultAudioEndpoint: ...");
        Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(0/*eRender*/, 1/*eMultimedia*/, out dev));

        listDeviceProperties(dev, new DEVICE_SUMMARY());
        Marshal.ReleaseComObject(dev);

        System.Console.WriteLine();
      }

    ////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////

    [Guid("1BE09788-6894-4089-8586-9A2A6C265AC5"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMEndpoint {
      int GetDataFlow(out int dataFlow);
    }

    [Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAudioEndpointVolume {
      // f(), g(), ... are unused COM method slots. Define these if you care
      int f(); int g(); int h(); int i();
      int SetMasterVolumeLevelScalar(float fLevel, System.Guid pguidEventContext);
      int j();
      int GetMasterVolumeLevelScalar(out float pfLevel);
      int k(); int l(); int m(); int n();
      int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, System.Guid pguidEventContext);
      int GetMute(out bool pbMute);
    }

    [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDevice {
      int Activate(ref System.Guid id, int clsCtx, int activationParams, out IAudioEndpointVolume aev);
      void OpenPropertyStore([In] uint stgmAccess, [MarshalAs(UnmanagedType.Interface)] out IPropertyStore ppProperties);
      void GetId([MarshalAs(UnmanagedType.LPWStr)] out string ppstrId);
      void GetState(out uint pdwState);
      //void QueryInterface(ref System.Guid giid, out IMMEndpoint ppvObject);
    }

    [Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), TypeLibType(TypeLibTypeFlags.FNonExtensible)]
    [ComImport] public interface IMMDeviceCollection {
      void GetCount(out uint pcDevices);
      void Item([In] uint nDevice, [MarshalAs(UnmanagedType.Interface)] out IMMDevice ppDevice);
    }

    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport] public interface IMMDeviceEnumerator {
      void EnumAudioEndpoints([In] int dataFlow, [In] uint dwStateMask, [MarshalAs(UnmanagedType.Interface)] out IMMDeviceCollection ppDevices);
      int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice endpoint);
      void GetDevice([MarshalAs(UnmanagedType.LPWStr)] [In] string pwstrId, [MarshalAs(UnmanagedType.Interface)] out IMMDevice ppDevice);
      //void RegisterEndpointNotificationCallback([MarshalAs(UnmanagedType.Interface)] [In] IMMNotificationClient pClient);
      //void UnregisterEndpointNotificationCallback([MarshalAs(UnmanagedType.Interface)] [In] IMMNotificationClient pClient);
    }

    [Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport] public interface IPropertyStore {
      [MethodImpl(MethodImplOptions.InternalCall)]
      void GetCount(out uint cProps);
      [MethodImpl(MethodImplOptions.InternalCall)]
      void GetAt([In] uint iProp, out PropertyKey pkey);
      [MethodImpl(MethodImplOptions.InternalCall)]
      void GetValue([In] ref PropertyKey key, out PROPVARIANT pv);
      [MethodImpl(MethodImplOptions.InternalCall)]
      void SetValue([In] ref PropertyKey key, [In] ref PROPVARIANT propvar);
      [MethodImpl(MethodImplOptions.InternalCall)]
      void Commit();
    }

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct PropertyKey {
      public static PropertyKey PKEY_Device_DeviceDesc = new PropertyKey { fmtid = new Guid(unchecked((int)0xa45c254e), unchecked((short)0xdf1c), 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0), pid = 2 }; // DEVPROP_TYPE_STRING
      public static PropertyKey PKEY_Device_HardwareIds = new PropertyKey { fmtid = new Guid(unchecked((int)0xa45c254e), unchecked((short)0xdf1c), 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0), pid = 3 }; // DEVPROP_TYPE_STRING_LIST
      public static PropertyKey PKEY_Device_Service = new PropertyKey { fmtid = new Guid(unchecked((int)0xa45c254e), unchecked((short)0xdf1c), 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0), pid = 6 }; // DEVPROP_TYPE_STRING
      public static PropertyKey PKEY_Device_Class = new PropertyKey { fmtid = new Guid(unchecked((int)0xa45c254e), unchecked((short)0xdf1c), 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0), pid = 9 }; // DEVPROP_TYPE_STRING
      public static PropertyKey PKEY_Device_Driver = new PropertyKey { fmtid = new Guid(unchecked((int)0xa45c254e), unchecked((short)0xdf1c), 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0), pid = 11 }; // DEVPROP_TYPE_STRING
      public static PropertyKey PKEY_Device_ConfigFlags = new PropertyKey { fmtid = new Guid(unchecked((int)0xa45c254e), unchecked((short)0xdf1c), 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0), pid = 12 }; // DEVPROP_TYPE_UINT32
      public static PropertyKey PKEY_Device_Manufacturer = new PropertyKey { fmtid = new Guid(unchecked((int)0xa45c254e), unchecked((short)0xdf1c), 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0), pid = 13 }; // DEVPROP_TYPE_STRING
      public static PropertyKey PKEY_Device_FriendlyName = new PropertyKey { fmtid = new Guid(unchecked((int)0xa45c254e), unchecked((short)0xdf1c), 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0), pid = 14 }; // DEVPROP_TYPE_STRING
      public static PropertyKey PKEY_Device_LocationInfo = new PropertyKey { fmtid = new Guid(unchecked((int)0xa45c254e), unchecked((short)0xdf1c), 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0), pid = 15 }; // DEVPROP_TYPE_STRING
      public static PropertyKey PKEY_Device_Capabilities = new PropertyKey { fmtid = new Guid(unchecked((int)0xa45c254e), unchecked((short)0xdf1c), 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0), pid = 17 }; // DEVPROP_TYPE_UNINT32
      public static PropertyKey PKEY_Device_BusNumber = new PropertyKey { fmtid = new Guid(unchecked((int)0xa45c254e), unchecked((short)0xdf1c), 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0), pid = 23 }; // DEVPROP_TYPE_UINT32
      public static PropertyKey PKEY_Device_EnumeratorName = new PropertyKey { fmtid = new Guid(unchecked((int)0xa45c254e), unchecked((short)0xdf1c), 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0), pid = 24 }; // DEVPROP_TYPE_STRING
      public static PropertyKey PKEY_Device_DevType = new PropertyKey { fmtid = new Guid(unchecked((int)0xa45c254e), unchecked((short)0xdf1c), 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0), pid = 27 }; // DEVPROP_TYPE_UINT32
      public static PropertyKey PKEY_Device_Characteristics = new PropertyKey { fmtid = new Guid(unchecked((int)0xa45c254e), unchecked((short)0xdf1c), 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0), pid = 29 }; // DEVPROP_TYPE_UINT32
      public static PropertyKey PKEY_Device_ManufacturerAttributes = new PropertyKey { fmtid = new Guid(unchecked((int)0x80d81ea6), unchecked((short)0x7473), 0x4b0c, 0x82, 0x16, 0xef, 0xc1, 0x1a, 0x2c, 0x4c, 0x8b), pid = 4 }; // DEVPROP_TYPE_UINT32
      public static PropertyKey PKEY_DeviceClass_IconPath = new PropertyKey { fmtid = new Guid(unchecked((int)0x259abffc), 0x50a7, 0x47ce, 0xaf, 0x8, 0x68, 0xc9, 0xa7, 0xd7, 0x33, 0x66), pid = 12 }; // DEVPROP_TYPE_STRING_LIST
      public static PropertyKey PKEY_DeviceClass_ClassCoInstallers = new PropertyKey { fmtid = new Guid(unchecked((int)0x713d1703), 0xa2e2, 0x49f5, 0x92, 0x14, 0x56, 0x47, 0x2e, 0xf3, 0xda, 0x5c), pid = 2 }; // DEVPROP_TYPE_STRING_LIST
      public static PropertyKey PKEY_DeviceInterface_FriendlyName = new PropertyKey { fmtid = new Guid(unchecked((int)0x026e516e), 0xb814, 0x414b, 0x83, 0xcd, 0x85, 0x6d, 0x6f, 0xef, 0x48, 0x22), pid = 2 }; // DEVPROP_TYPE_STRING

      public Guid fmtid;
      public uint pid;

      public static IEnumerable<PropertyKey> GetPropertyKeys() {
        var keyFields = typeof(PropertyKey).GetFields(BindingFlags.Public | BindingFlags.Static);
        return keyFields.Where(fieldInfo => fieldInfo.FieldType == typeof(PropertyKey))
                        .Select(fieldInfo => (PropertyKey)fieldInfo.GetValue(null));
      }

      public static string GetKeyName(PropertyKey propertyKey) {
        var keyFields = typeof(PropertyKey).GetFields(BindingFlags.Public | BindingFlags.Static);
        return keyFields.Select(fieldInfo => new { fieldInfo, value = (PropertyKey)fieldInfo.GetValue(null) })
                        .Where(@t => propertyKey.pid == @t.value.pid && propertyKey.fmtid == @t.value.fmtid)
                        .Select(@t => @t.fieldInfo.Name).FirstOrDefault();
      }
    }

    [StructLayout(LayoutKind.Sequential, Pack = 8)]
    public struct PROPVARIANT {
      public ushort variantType;
      public byte wReserved1;
      public byte wReserved2;
      public uint wReserved3;
      public PROPVARIANTVALUE value;
      public object Value {
        get {
          switch ((VarEnum)variantType) {
            case VarEnum.VT_EMPTY           : return null;           case VarEnum.VT_NULL            : return null;
            case VarEnum.VT_VARIANT         : break;                 case VarEnum.VT_DECIMAL         : return value.cyVal;
            case VarEnum.VT_VOID            : break;                 case VarEnum.VT_HRESULT         : break;
            case VarEnum.VT_PTR             : break;                 case VarEnum.VT_SAFEARRAY       : break;
            case VarEnum.VT_CARRAY          : break;                 case VarEnum.VT_USERDEFINED     : break;
            case VarEnum.VT_RECORD          : break;                 case VarEnum.VT_STREAM          : break;
            case VarEnum.VT_STORAGE         : break;                 case VarEnum.VT_STREAMED_OBJECT : break;
            case VarEnum.VT_STORED_OBJECT   : break;                 case VarEnum.VT_BLOB_OBJECT     : break;
            case VarEnum.VT_CF              : break;                 case VarEnum.VT_CLSID           : return Marshal.PtrToStructure(value.pVal, typeof(Guid));
            case VarEnum.VT_VECTOR          : break;                 case VarEnum.VT_ARRAY           : break;
            case VarEnum.VT_BYREF           : break;                 case VarEnum.VT_I1              : return value.cVal;
            case VarEnum.VT_UI1             : return value.bVal;     case VarEnum.VT_I2              : return value.iVal;
            case VarEnum.VT_UI2             : return value.uiVal;
            case VarEnum.VT_I4              :                        case VarEnum.VT_INT             : return value.intVal;
            case VarEnum.VT_UI4             :                        case VarEnum.VT_UINT            : return value.uintVal;
            case VarEnum.VT_I8              : return value.hVal;     case VarEnum.VT_UI8             : return value.uhVal;
            case VarEnum.VT_R4              : return value.fltVal;   case VarEnum.VT_R8              : return value.dblVal;
            case VarEnum.VT_BOOL            : return value.boolVal;  case VarEnum.VT_ERROR           : return value.scode;
            case VarEnum.VT_CY              : return value.cyVal;
            case VarEnum.VT_DATE            : return value.date;     case VarEnum.VT_FILETIME        : return DateTime.FromFileTime(value.hVal);
            case VarEnum.VT_BSTR            : return Marshal.PtrToStringBSTR(value.pVal);
            case VarEnum.VT_BLOB            :
              var blob = value.blob;
              var blobData = new byte[blob.cbSize];
              Marshal.Copy(blob.pBlobData, blobData, 0, (int)blob.cbSize);
              return blobData;
            case VarEnum.VT_LPSTR           : return Marshal.PtrToStringAnsi(value.pVal);
            case VarEnum.VT_LPWSTR          : return Marshal.PtrToStringUni(value.pVal);
            case VarEnum.VT_UNKNOWN         : return Marshal.GetObjectForIUnknown(value.pVal);
            case VarEnum.VT_DISPATCH        : return value.pVal;
            //default                       : throw new NotSupportedException("The type of this variable is not support ('" + variantType + "')");
          }
          return string.Format("unsupported {0}", ((VarEnum)variantType));
          //throw new NotSupportedException("The type of this variable is not support ('" + variantType.ToString() + "')");
        }
      }
    }

    [ComConversionLoss]
    [StructLayout(LayoutKind.Explicit, Pack = 8, Size = 8)]
    public struct PROPVARIANTVALUE {
      [FieldOffset(0)] public sbyte cVal;      [FieldOffset(0)] public byte bVal;
      [FieldOffset(0)] public short iVal;      [FieldOffset(0)] public ushort uiVal;
      [FieldOffset(0)] public int   intVal;    [FieldOffset(0)] public uint uintVal;
      [FieldOffset(0)] public long  hVal;      [FieldOffset(0)] public ulong uhVal;
      [FieldOffset(0)] public float fltVal;    [FieldOffset(0)] public double dblVal;
      [FieldOffset(0)] public short boolVal;
      [FieldOffset(0)] [MarshalAs(UnmanagedType.Error)] public int scode;
      [FieldOffset(0)] [MarshalAs(UnmanagedType.Currency)] public decimal cyVal;
      [FieldOffset(0)] public DateTime date;   [FieldOffset(0)] public tagFILETIME filetime;
      [FieldOffset(0)] public tagARRAY array;  [FieldOffset(0)] public tagBLOB blob;
      [ComConversionLoss] [FieldOffset(0)] public IntPtr pVal;
    }

    [ComConversionLoss] [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct tagARRAY {
      public uint cElems;
      [ComConversionLoss] public IntPtr pElems;
    }

    [ComConversionLoss] [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct tagBLOB {
      public uint cbSize;
      [ComConversionLoss] public IntPtr pBlobData;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct tagFILETIME {
      public uint dwLowDateTime;
      public uint dwHighDateTime;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 8)]
    public struct tagLARGEINTEGER {
      public long QuadPart;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 8)]
    public struct tagULARGEINTEGER {
      public ulong QuadPart;
    }

  } // myClass
} // myNameSpace

using System;
using System.Runtime.InteropServices;

namespace Extensibility
{
    /// <summary>
    /// IDTExtensibility2 接口定义（手动定义以避免依赖Office PIA）
    /// </summary>
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDTExtensibility2
    {
        void OnConnection(
            [MarshalAs(UnmanagedType.IDispatch)] object Application,
            int ConnectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
            ref Array custom);

        void OnDisconnection(
            int RemoveMode,
            ref Array custom);

        void OnAddInsUpdate(ref Array custom);

        void OnStartupComplete(ref Array custom);

        void OnBeginShutdown(ref Array custom);
    }

    /// <summary>
    /// IRibbonExtensibility 接口定义
    /// </summary>
    [ComImport]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonExtensibility
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        [DispId(1)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string RibbonID);
    }
}

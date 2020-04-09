// VBAExtensions
//
// C# Library module for VBA
//
// By Steve Laycock
//
using System.Runtime.InteropServices;

namespace VBALibrary
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IKVPair
    {
        dynamic Key { get; set; }
        dynamic Value { get; set; }
    }
}
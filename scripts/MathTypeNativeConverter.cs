using System;
using System.Runtime.InteropServices;
using System.Text;

public static class MathTypeNativeConverter
{
    private const string Library = @"C:\Program Files (x86)\MathType\MathPage\64\MathPage.wll";

    [DllImport(Library, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern int MTInitAPI(short options, short timeout);

    [DllImport(Library, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern int MTTermAPI();

    [DllImport(Library, ExactSpelling = true, CallingConvention = CallingConvention.StdCall, CharSet = CharSet.Unicode)]
    public static extern int MTSetMTPrefs(short mode, [MarshalAs(UnmanagedType.LPWStr)] string prefs, short timeout);

    [DllImport(Library, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern int MTCloseOleObject(int saveOptions, [MarshalAs(UnmanagedType.IDispatch)] object oleObject);

    [DllImport(Library, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    private static extern int MTSetEqnFromLangStr(
        [MarshalAs(UnmanagedType.IDispatch)] object oleObject,
        short languageType,
        IntPtr unicodeBuffer,
        int characterCount);

    public static int SetMathML(object oleObject, string mathML)
    {
        byte[] bytes = Encoding.Unicode.GetBytes(mathML);
        GCHandle pin = GCHandle.Alloc(bytes, GCHandleType.Pinned);
        try
        {
            return MTSetEqnFromLangStr(oleObject, 2, pin.AddrOfPinnedObject(), mathML.Length);
        }
        finally
        {
            pin.Free();
        }
    }
}

using static MsWordDiff.ShellExtension.NativeMethods;

namespace MsWordDiff.ShellExtension;

[ComVisible(true)]
[Guid(ClassGuid)]
[ClassInterface(ClassInterfaceType.None)]
public class CompareContextMenu : IShellExtInit, IContextMenu
{
    public const string ClassGuid = "8A9C5E71-7D23-4D8A-B5E6-F1A2B3C4D5E6";

    const uint IdmCompare = 0;
    const string MenuText = "Compare with MsWordDiff";

    List<string> selectedFiles = [];

    const int S_OK = 0;
    const int E_FAIL = unchecked((int)0x80004005);
    const int E_NOTIMPL = unchecked((int)0x80004001);

    public int Initialize(nint pidlFolder, nint pDataObj, nint hKeyProgID)
    {
        try
        {
            selectedFiles.Clear();

            if (pDataObj == nint.Zero)
            {
                return S_OK;
            }

            var formatetc = new FORMATETC
            {
                cfFormat = CF_HDROP,
                ptd = nint.Zero,
                dwAspect = 1, // DVASPECT_CONTENT
                lindex = -1,
                tymed = TYMED_HGLOBAL
            };

            var stgmedium = new STGMEDIUM();

            try
            {
                var dataObject = (System.Runtime.InteropServices.ComTypes.IDataObject)Marshal.GetObjectForIUnknown(pDataObj);
                dataObject.GetData(ref Unsafe.As<FORMATETC, System.Runtime.InteropServices.ComTypes.FORMATETC>(ref formatetc),
                    out Unsafe.As<STGMEDIUM, System.Runtime.InteropServices.ComTypes.STGMEDIUM>(ref stgmedium));

                var hDrop = stgmedium.unionmember;
                var fileCount = DragQueryFile(hDrop, 0xFFFFFFFF, null, 0);

                for (uint i = 0; i < fileCount; i++)
                {
                    var pathLen = DragQueryFile(hDrop, i, null, 0);
                    var sb = new StringBuilder((int)pathLen + 1);
                    DragQueryFile(hDrop, i, sb, pathLen + 1);
                    selectedFiles.Add(sb.ToString());
                }
            }
            finally
            {
                ReleaseStgMedium(ref stgmedium);
            }

            return S_OK;
        }
        catch
        {
            return E_FAIL;
        }
    }

    public int QueryContextMenu(nint hMenu, uint indexMenu, uint idCmdFirst, uint idCmdLast, uint uFlags)
    {
        try
        {
            // Only show menu if exactly 2 Word documents are selected
            if (selectedFiles.Count != 2)
            {
                return 0;
            }

            var allWordDocs = selectedFiles.All(f =>
                f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ||
                f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase));

            if (!allWordDocs)
            {
                return 0;
            }

            InsertMenu(hMenu, indexMenu, MF_STRING | MF_BYPOSITION, idCmdFirst + IdmCompare, MenuText);

            return 1;
        }
        catch
        {
            return 0;
        }
    }

    public int InvokeCommand(nint pici)
    {
        try
        {
            var ici = Marshal.PtrToStructure<CMINVOKECOMMANDINFOEX>(pici);

            var isMenuId = (ici.lpVerb.ToInt64() & 0xFFFF0000) == 0;

            if (!isMenuId)
            {
                return S_OK;
            }

            var menuId = (uint)(ici.lpVerb.ToInt64() & 0xFFFF);

            if (menuId == IdmCompare && selectedFiles.Count == 2)
            {
                LaunchComparison();
            }

            return S_OK;
        }
        catch
        {
            return E_FAIL;
        }
    }

    public int GetCommandString(nuint idCmd, uint uType, nint pReserved, nint pszName, uint cchMax) =>
        E_NOTIMPL;

    void LaunchComparison()
    {
        var exePath = GetExePath();
        if (exePath == null)
        {
            return;
        }

        var args = $"\"{selectedFiles[0]}\" \"{selectedFiles[1]}\"";

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = args,
                UseShellExecute = false
            });
        }
        catch
        {
            // Silently fail - we're in Explorer's process
        }
    }

    static string? GetExePath()
    {
        // Try .NET global tools location first
        var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var globalToolPath = Path.Combine(userProfile, ".dotnet", "tools", "diffword.exe");
        if (File.Exists(globalToolPath))
        {
            return globalToolPath;
        }

        // Try PATH
        var envPath = Environment.GetEnvironmentVariable("PATH") ?? "";
        var paths = envPath.Split(Path.PathSeparator);

        foreach (var path in paths)
        {
            var candidate = Path.Combine(path, "diffword.exe");
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return null;
    }
}

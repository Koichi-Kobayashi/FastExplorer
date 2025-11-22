using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace FastExplorer.Helpers
{
    /// <summary>
    /// エクスプローラーと同じ「シェル コンテキストメニュー」を表示するヘルパークラス
    /// </summary>
    public class ShellContextMenu
    {
        // Public メイン API
        public void ShowContextMenu(string[] filePaths, IntPtr ownerHwnd, int x, int y)
        {
            if (filePaths == null || filePaths.Length == 0)
            {
                return;
            }

            IntPtr[] pidls = new IntPtr[filePaths.Length];
            try
            {
                // 各ファイルパス → PIDL に変換
                for (int i = 0; i < filePaths.Length; i++)
                {
                    pidls[i] = GetPIDLFromPath(filePaths[i]);
                    if (pidls[i] == IntPtr.Zero)
                    {
                        throw new InvalidOperationException("PIDL 取得に失敗: " + filePaths[i]);
                    }
                }

                // 親フォルダーの IShellFolder と、子 PIDL 群を取得（最適化：1回の呼び出しで取得）
                IntPtr parentPidl;
                IntPtr parentFolderPtr;
                Guid iidShellFolder = IID_IShellFolder;
                int hr = Win32.SHBindToParent(pidls[0], ref iidShellFolder, out parentFolderPtr, out parentPidl);
                if (hr != 0 || parentFolderPtr == IntPtr.Zero)
                {
                    throw new InvalidOperationException("親フォルダー取得に失敗");
                }

                IShellFolder parentFolder = (IShellFolder)Marshal.GetTypedObjectForIUnknown(parentFolderPtr, typeof(IShellFolder));

                // 相対PIDL配列を構築（高速化：ILFindLastIDを使用）
                IntPtr[] relativePidls = new IntPtr[pidls.Length];
                relativePidls[0] = parentPidl; // 最初のPIDLは既に取得済み
                for (int i = 1; i < pidls.Length; i++)
                {
                    relativePidls[i] = Win32.ILFindLastID(pidls[i]);
                }

                // IContextMenu を取得
                IntPtr contextMenuPtr;
                Guid iidContextMenu = IID_IContextMenu;
                parentFolder.GetUIObjectOf(
                    IntPtr.Zero,
                    (uint)relativePidls.Length,
                    relativePidls,
                    ref iidContextMenu,
                    IntPtr.Zero,
                    out contextMenuPtr);
                
                // parentFolderPtr を解放（parentFolder は既に取得済み）
                Marshal.Release(parentFolderPtr);

                if (contextMenuPtr == IntPtr.Zero)
                {
                    return;
                }

                var iContextMenu = (IContextMenu)Marshal.GetTypedObjectForIUnknown(contextMenuPtr, typeof(IContextMenu));

                // ポップアップメニュー作成
                IntPtr hMenu = Win32.CreatePopupMenu();
                if (hMenu == IntPtr.Zero)
                {
                    Marshal.Release(contextMenuPtr);
                    return;
                }

                try
                {
                    // メニュー構築（高速化：CMF.EXPLOREを削除）
                    uint idCmdFirst = 1;
                    iContextMenu.QueryContextMenu(
                        hMenu,
                        0,
                        idCmdFirst,
                        0x7FFF,
                        CMF.NORMAL);

                    // メニューを表示
                    uint selected = Win32.TrackPopupMenuEx(
                        hMenu,
                        TPM.RETURNCMD,
                        x,
                        y,
                        ownerHwnd,
                        IntPtr.Zero);

                    if (selected != 0)
                    {
                        // 選択されたコマンドを IContextMenu に伝える
                        var ici = new CMINVOKECOMMANDINFOEX();
                        ici.cbSize = Marshal.SizeOf(typeof(CMINVOKECOMMANDINFOEX));
                        ici.fMask = CMIC.UNICODE | CMIC.PTINVOKE;
                        ici.hwnd = ownerHwnd;
                        ici.lpVerb = (IntPtr)(selected - idCmdFirst);
                        ici.lpParameters = null;
                        ici.lpDirectory = null;
                        ici.nShow = SW.SHOWNORMAL;
                        ici.dwHotKey = 0;
                        ici.hIcon = IntPtr.Zero;
                        ici.lpTitle = IntPtr.Zero;
                        ici.lpVerbW = IntPtr.Zero;
                        ici.lpParametersW = IntPtr.Zero;
                        ici.lpDirectoryW = IntPtr.Zero;
                        ici.lpTitleW = IntPtr.Zero;
                        ici.ptInvoke = new POINT { X = x, Y = y };
                        iContextMenu.InvokeCommand(ref ici);
                    }
                }
                finally
                {
                    Win32.DestroyMenu(hMenu);
                    if (contextMenuPtr != IntPtr.Zero)
                    {
                        Marshal.Release(contextMenuPtr);
                    }
                }
            }
            catch (Exception)
            {
                // エラーは無視（デバッグ時のみログ出力）
                #if DEBUG
                System.Diagnostics.Debug.WriteLine($"[ShellContextMenu] 例外が発生しました");
                #endif
            }
            finally
            {
                // PIDL 解放
                foreach (var pidl in pidls)
                {
                    if (pidl != IntPtr.Zero)
                        Win32.CoTaskMemFree(pidl);
                }
            }
        }

        #region PIDL / IShellFolder ヘルパー

        private static IntPtr GetPIDLFromPath(string path)
        {
            IntPtr pidl;
            uint attrs;
            int hr = Win32.SHParseDisplayName(path, IntPtr.Zero, out pidl, 0, out attrs);
            if (hr != 0)
                return IntPtr.Zero;
            return pidl;
        }


        #endregion

        #region COM / Win32 定義

        private static readonly Guid IID_IShellFolder = new("000214E6-0000-0000-C000-000000000046");
        private static readonly Guid IID_IContextMenu = new("000214E4-0000-0000-C000-000000000046");

        [Flags]
        private enum CMF : uint
        {
            NORMAL = 0x00000000,
            DEFAULTONLY = 0x00000001,
            VERBSONLY = 0x00000002,
            EXPLORE = 0x00000004,
        }

        [Flags]
        private enum CMIC : uint
        {
            UNICODE = 0x00004000,
            PTINVOKE = 0x20000000
        }

        private static class SW
        {
            public const int SHOWNORMAL = 1;
        }

        [Flags]
        private enum TPM : uint
        {
            RETURNCMD = 0x0100
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct CMINVOKECOMMANDINFOEX
        {
            public int cbSize;
            public CMIC fMask;
            public IntPtr hwnd;
            public IntPtr lpVerb;
            [MarshalAs(UnmanagedType.LPStr)]
            public string? lpParameters;
            [MarshalAs(UnmanagedType.LPStr)]
            public string? lpDirectory;
            public int nShow;
            public int dwHotKey;
            public IntPtr hIcon;
            public IntPtr lpTitle;
            public IntPtr lpVerbW;
            public IntPtr lpParametersW;
            public IntPtr lpDirectoryW;
            public IntPtr lpTitleW;
            public POINT ptInvoke;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("000214E4-0000-0000-C000-000000000046")]
        private interface IContextMenu
        {
            [PreserveSig]
            int QueryContextMenu(
                IntPtr hMenu,
                uint indexMenu,
                uint idCmdFirst,
                uint idCmdLast,
                CMF uFlags);

            void InvokeCommand(ref CMINVOKECOMMANDINFOEX pici);

            void GetCommandString(
                IntPtr idCmd,
                uint uFlags,
                IntPtr pReserved,
                IntPtr pszName,
                uint cchMax);
        }

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("000214E6-0000-0000-C000-000000000046")]
        private interface IShellFolder
        {
            void ParseDisplayName(
                IntPtr hwnd,
                IntPtr pbc,
                [MarshalAs(UnmanagedType.LPWStr)] string pszDisplayName,
                ref uint pchEaten,
                out IntPtr ppidl,
                ref uint pdwAttributes);

            void EnumObjects(
                IntPtr hwnd,
                uint grfFlags,
                out IntPtr ppenumIDList);

            void BindToObject(
                IntPtr pidl,
                IntPtr pbc,
                ref Guid riid,
                out IntPtr ppv);

            void BindToStorage(
                IntPtr pidl,
                IntPtr pbc,
                ref Guid riid,
                out IntPtr ppv);

            void CompareIDs(
                IntPtr lParam,
                IntPtr pidl1,
                IntPtr pidl2);

            void CreateViewObject(
                IntPtr hwndOwner,
                ref Guid riid,
                out IntPtr ppv);

            void GetAttributesOf(
                uint cidl,
                IntPtr apidl,
                ref uint rgfInOut);

            void GetUIObjectOf(
                IntPtr hwndOwner,
                uint cidl,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] IntPtr[] apidl,
                ref Guid riid,
                IntPtr rgfReserved,
                out IntPtr ppv);

            void GetDisplayNameOf(
                IntPtr pidl,
                uint uFlags,
                out IntPtr ppszName);

            void SetNameOf(
                IntPtr hwnd,
                IntPtr pidl,
                [MarshalAs(UnmanagedType.LPWStr)] string pszName,
                uint uFlags,
                out IntPtr ppidlOut);
        }

        private static class Win32
        {
            [DllImport("shell32.dll", CharSet = CharSet.Auto)]
            public static extern int SHParseDisplayName([MarshalAs(UnmanagedType.LPWStr)] string pszName, IntPtr pbc, out IntPtr ppidl, uint sfgaoIn, out uint psfgaoOut);

            [DllImport("shell32.dll")]
            public static extern int SHBindToParent(IntPtr pidl, ref Guid riid, out IntPtr ppv, out IntPtr ppidlLast);

            [DllImport("shell32.dll")]
            public static extern IntPtr ILFindLastID(IntPtr pidl);

            [DllImport("ole32.dll")]
            public static extern void CoTaskMemFree(IntPtr pv);

            [DllImport("user32.dll")]
            public static extern IntPtr CreatePopupMenu();

            [DllImport("user32.dll")]
            public static extern bool DestroyMenu(IntPtr hMenu);

            [DllImport("user32.dll")]
            public static extern uint TrackPopupMenuEx(IntPtr hMenu, TPM uFlags, int x, int y, IntPtr hWnd, IntPtr lptpm);
        }

        #endregion
    }
}

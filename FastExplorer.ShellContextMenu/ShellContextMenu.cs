using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace FastExplorer.ShellContextMenu
{
    /// <summary>
    /// Windowsシェルのコンテキストメニューを表示するサービスクラス。
    /// エクスプローラーと同じ右クリックメニューを表示する機能を提供します。
    /// </summary>
    /// <remarks>
    /// <para>
    /// このクラスは、WindowsシェルのCOMインターフェース（<see cref="IContextMenu"/>、<see cref="IContextMenu2"/>、<see cref="IContextMenu3"/>）を使用して、
    /// 指定されたファイルまたはフォルダーに対する標準的なコンテキストメニューを表示します。
    /// </para>
    /// <para>
    /// メニューの表示には、メインウィンドウのメッセージフック（<see cref="ProcessWindowMessage"/>）を設定する必要があります。
    /// これにより、<see cref="IContextMenu3"/>が提供する高度なメニュー描画機能が正しく動作します。
    /// </para>
    /// <para>
    /// 使用例：
    /// <code>
    /// var service = new ShellContextMenuService();
    /// service.ShowContextMenu(new[] { @"C:\path\to\file.txt" }, windowHandle, x, y);
    /// </code>
    /// </para>
    /// </remarks>
    public class ShellContextMenuService
    {
        // 現在のIContextMenu3インスタンスを保持（メッセージフック用）
        private static IContextMenu3? s_currentContextMenu3;
        private static IntPtr s_currentContextMenuPtr = IntPtr.Zero;
        private static bool s_isMenuShowing = false; // メニューが表示中かどうかのフラグ

        // メッセージフック用の定数
        private const int WM_INITMENUPOPUP = 0x0117;
        private const int WM_DRAWITEM = 0x002B;
        private const int WM_MEASUREITEM = 0x002C;
        private const int WM_MENUCHAR = 0x0120;

        /// <summary>
        /// 現在コンテキストメニューが表示中かどうかを示す値を取得します。
        /// </summary>
        /// <value>
        /// メニューが表示中の場合は<see langword="true"/>、それ以外の場合は<see langword="false"/>。
        /// </value>
        /// <remarks>
        /// このプロパティは、メッセージフックがアクティブにメニュー関連のメッセージを処理しているかどうかを示します。
        /// </remarks>
        public static bool IsMenuShowing => s_isMenuShowing;

        /// <summary>
        /// ウィンドウメッセージを処理し、コンテキストメニュー関連のメッセージを<see cref="IContextMenu3"/>に転送します。
        /// </summary>
        /// <param name="hwnd">メッセージを受信したウィンドウのハンドル。</param>
        /// <param name="msg">メッセージID。</param>
        /// <param name="wParam">メッセージの追加情報（WPARAM）。</param>
        /// <param name="lParam">メッセージの追加情報（LPARAM）。</param>
        /// <param name="handled">メッセージが処理されたかどうかを示すフラグ。処理された場合は<see langword="true"/>に設定されます。</param>
        /// <returns>
        /// メッセージ処理の結果値。処理されなかった場合は<see cref="IntPtr.Zero"/>を返します。
        /// </returns>
        /// <remarks>
        /// <para>
        /// このメソッドは、WPFアプリケーションの<see cref="HwndSource.AddHook(System.Windows.Interop.HwndSourceHook)"/>から呼び出されることを想定しています。
        /// メインウィンドウの<see cref="System.Windows.Window.OnSourceInitialized(System.EventArgs)"/>でメッセージフックを設定してください。
        /// </para>
        /// <para>
        /// 以下のメッセージのみを処理します：
        /// <list type="bullet">
        /// <item><description>WM_INITMENUPOPUP (0x0117)</description></item>
        /// <item><description>WM_DRAWITEM (0x002B)</description></item>
        /// <item><description>WM_MEASUREITEM (0x002C)</description></item>
        /// <item><description>WM_MENUCHAR (0x0120)</description></item>
        /// </list>
        /// </para>
        /// <para>
        /// メニューが表示されていない場合、または<see cref="IContextMenu3"/>が利用できない場合は、メッセージを処理せずに<paramref name="handled"/>を<see langword="false"/>のままにします。
        /// これにより、ウィンドウの通常のメッセージ処理（最大化、閉じるボタンなど）が妨げられません。
        /// </para>
        /// </remarks>
        public static IntPtr ProcessWindowMessage(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            // メニューが表示中で、IContextMenu3が存在する場合のみ処理
            if (s_isMenuShowing && s_currentContextMenu3 != null &&
                (msg == WM_INITMENUPOPUP ||
                 msg == WM_DRAWITEM ||
                 msg == WM_MEASUREITEM ||
                 msg == WM_MENUCHAR))
            {
                try
                {
                    IntPtr result;
                    int hr = s_currentContextMenu3.HandleMenuMsg2(
                        (uint)msg,
                        wParam,
                        lParam,
                        out result);

                    if (hr == 0) // S_OK
                    {
                        handled = true;
                        return result;
                    }
                }
                catch
                {
                    // エラーが発生した場合は処理しない
                    // handledはfalseのまま
                }
            }

            // メニューが表示されていない、または処理されなかった場合はhandledはfalseのまま
            handled = false;
            return IntPtr.Zero;
        }

        /// <summary>
        /// 指定されたファイルまたはフォルダーのシェルコンテキストメニューを表示します。
        /// </summary>
        /// <param name="filePaths">コンテキストメニューを表示するファイルまたはフォルダーのパスの配列。複数選択に対応しています。</param>
        /// <param name="ownerHwnd">メニューのオーナーウィンドウのハンドル。通常はメインウィンドウのハンドルを指定します。</param>
        /// <param name="x">メニューを表示するX座標（スクリーン座標）。</param>
        /// <param name="y">メニューを表示するY座標（スクリーン座標）。</param>
        /// <remarks>
        /// <para>
        /// このメソッドは、WindowsシェルのCOMインターフェースを使用して、指定されたファイルまたはフォルダーに対する標準的なコンテキストメニューを表示します。
        /// メニューには、ファイルの種類やインストールされているシェル拡張に応じて、様々な操作（開く、コピー、削除など）が表示されます。
        /// </para>
        /// <para>
        /// 処理の流れ：
        /// <list type="number">
        /// <item><description>各ファイルパスをPIDL（Pointer to an Item ID List）に変換します。</description></item>
        /// <item><description>親フォルダーの<see cref="IShellFolder"/>インターフェースを取得します。</description></item>
        /// <item><description><see cref="IContextMenu"/>インターフェースを取得し、可能であれば<see cref="IContextMenu3"/>にキャストします。</description></item>
        /// <item><description>メニューを構築し、指定された位置に表示します。</description></item>
        /// <item><description>ユーザーがメニュー項目を選択した場合、対応するコマンドを実行します。</description></item>
        /// </list>
        /// </para>
        /// <para>
        /// 注意事項：
        /// <list type="bullet">
        /// <item><description><paramref name="filePaths"/>が<see langword="null"/>または空の配列の場合、何も実行しません。</description></item>
        /// <item><description>PIDLの取得に失敗した場合、<see cref="InvalidOperationException"/>がスローされます。</description></item>
        /// <item><description>メニュー表示中は、<see cref="IsMenuShowing"/>が<see langword="true"/>になります。</description></item>
        /// <item><description>メニュー表示中は、<see cref="ProcessWindowMessage"/>がメニュー関連のメッセージを処理します。</description></item>
        /// <item><description>エラーが発生した場合、例外はキャッチされ、デバッグビルド時のみログに出力されます。</description></item>
        /// </list>
        /// </para>
        /// <para>
        /// COMオブジェクトのリソース管理：
        /// このメソッドは、取得したCOMオブジェクト（PIDL、<see cref="IShellFolder"/>、<see cref="IContextMenu"/>など）を適切に解放します。
        /// 例外が発生した場合でも、<see langword="finally"/>ブロックで確実にリソースが解放されます。
        /// </para>
        /// </remarks>
        /// <exception cref="InvalidOperationException">
        /// PIDLの取得に失敗した場合、または親フォルダーの取得に失敗した場合にスローされます。
        /// </exception>
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
                Guid iidShellFolder = IID_IShellFolder;
                int hr = Win32.SHBindToParent(pidls[0], ref iidShellFolder, out IShellFolder parentFolder, out parentPidl);
                if (hr != 0 || parentFolder == null)
                {
                    throw new InvalidOperationException("親フォルダー取得に失敗");
                }

                // 相対PIDL配列を構築（高速化：ILFindLastIDを使用）
                IntPtr[] relativePidls = new IntPtr[pidls.Length];
                relativePidls[0] = parentPidl; // 最初のPIDLは既に取得済み
                for (int i = 1; i < pidls.Length; i++)
                {
                    relativePidls[i] = Win32.ILFindLastID(pidls[i]);
                }

                // IContextMenu を取得（サンプルコードに合わせてIContextMenuを直接取得）
                Guid iidContextMenu = IID_IContextMenu;
                uint reserved = 0;
                parentFolder.GetUIObjectOf(
                    IntPtr.Zero,
                    (uint)relativePidls.Length,
                    relativePidls,
                    ref iidContextMenu,
                    ref reserved,
                    out object ppv);
                
                // parentFolder を解放
                Marshal.ReleaseComObject(parentFolder);

                if (ppv == null)
                {
                    return;
                }

                // IContextMenuを取得
                IContextMenu iContextMenu = (IContextMenu)ppv;
                
                // 前回のインスタンスをクリア
                s_currentContextMenu3 = null;
                s_isMenuShowing = false;
                if (s_currentContextMenuPtr != IntPtr.Zero)
                {
                    Marshal.Release(s_currentContextMenuPtr);
                    s_currentContextMenuPtr = IntPtr.Zero;
                }
                
                // IContextMenu3にキャストを試す（サンプルコードと同じ方法）
                s_currentContextMenu3 = iContextMenu as IContextMenu3;
                if (s_currentContextMenu3 != null)
                {
                    // IContextMenu3が取得できた場合、参照を保持
                    IntPtr contextMenuPtr = Marshal.GetIUnknownForObject(iContextMenu);
                    s_currentContextMenuPtr = contextMenuPtr;
                    // Marshal.GetIUnknownForObjectは既に参照カウントを増やしているので、AddRefは不要
                    #if DEBUG
                    System.Diagnostics.Debug.WriteLine("[ShellContextMenu] IContextMenu3を取得しました");
                    #endif
                }
                else
                {
                    #if DEBUG
                    System.Diagnostics.Debug.WriteLine("[ShellContextMenu] IContextMenu3へのキャストに失敗しました（IContextMenuのみ使用）");
                    #endif
                }
                
                // メニュー表示開始
                s_isMenuShowing = true;

                // ポップアップメニュー作成
                IntPtr hMenu = Win32.CreatePopupMenu();
                if (hMenu == IntPtr.Zero)
                {
                    Marshal.ReleaseComObject(iContextMenu);
                    return;
                }

                try
                {
                    // メニュー構築（Windows 11の新しいメニュースタイルを有効化）
                    uint idCmdFirst = 1;
                    // CMF.NORMALを使用（Windows 11の新しいメニュースタイルはIContextMenu3で自動的に適用される）
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
                    
                    // メニュー表示終了（先にフラグをクリアして、メッセージ処理を停止）
                    s_isMenuShowing = false;
                    
                    // IContextMenu3の参照をクリア
                    s_currentContextMenu3 = null;
                    if (s_currentContextMenuPtr != IntPtr.Zero)
                    {
                        Marshal.Release(s_currentContextMenuPtr);
                        s_currentContextMenuPtr = IntPtr.Zero;
                    }
                    
                    // IContextMenuを解放
                    Marshal.ReleaseComObject(iContextMenu);
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

        /// <summary>
        /// ファイルパスからPIDL（Pointer to an Item ID List）を取得します。
        /// </summary>
        /// <param name="path">PIDLに変換するファイルまたはフォルダーのパス。</param>
        /// <returns>
        /// 成功した場合はPIDLへのポインタ、失敗した場合は<see cref="IntPtr.Zero"/>。
        /// </returns>
        /// <remarks>
        /// <para>
        /// PIDLは、Windowsシェルでファイルやフォルダーを一意に識別するために使用される構造体です。
        /// このメソッドは、<see cref="Win32.SHParseDisplayName"/>を使用してパスをPIDLに変換します。
        /// </para>
        /// <para>
        /// 取得したPIDLは、使用後に<see cref="Win32.CoTaskMemFree"/>で解放する必要があります。
        /// </para>
        /// </remarks>
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
        private static readonly Guid IID_IContextMenu2 = new("000214F4-0000-0000-C000-000000000046");
        private static readonly Guid IID_IContextMenu3 = new("BCFCE0A0-EC17-11D0-8D10-00A0C90F2719");

        /// <summary>
        /// コンテキストメニューの構築時に使用するフラグ。
        /// </summary>
        [Flags]
        private enum CMF : uint
        {
            /// <summary>
            /// 通常のメニューを表示します。
            /// </summary>
            NORMAL = 0x00000000,
            /// <summary>
            /// デフォルトのコマンドのみを表示します。
            /// </summary>
            DEFAULTONLY = 0x00000001,
            /// <summary>
            /// 動詞（コマンド）のみを表示します。
            /// </summary>
            VERBSONLY = 0x00000002,
            /// <summary>
            /// エクスプローラーモードでメニューを表示します。
            /// </summary>
            EXPLORE = 0x00000004,
        }

        /// <summary>
        /// コマンド実行時に使用するフラグ。
        /// </summary>
        [Flags]
        private enum CMIC : uint
        {
            /// <summary>
            /// Unicode文字列を使用します。
            /// </summary>
            UNICODE = 0x00004000,
            /// <summary>
            /// メニューが表示された位置（ptInvoke）を渡します。
            /// </summary>
            PTINVOKE = 0x20000000
        }

        /// <summary>
        /// ShowWindow関数で使用する定数。
        /// </summary>
        private static class SW
        {
            /// <summary>
            /// ウィンドウを通常のサイズと位置で表示します。
            /// </summary>
            public const int SHOWNORMAL = 1;
        }

        /// <summary>
        /// TrackPopupMenuEx関数で使用するフラグ。
        /// </summary>
        [Flags]
        private enum TPM : uint
        {
            /// <summary>
            /// メニュー項目が選択された場合、コマンドIDを返します。
            /// </summary>
            RETURNCMD = 0x0100
        }

        /// <summary>
        /// コンテキストメニューのコマンド実行時に使用する情報構造体。
        /// </summary>
        /// <remarks>
        /// この構造体は、<see cref="IContextMenu.InvokeCommand"/>メソッドに渡され、実行するコマンドの詳細情報を提供します。
        /// </remarks>
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct CMINVOKECOMMANDINFOEX
        {
            /// <summary>
            /// 構造体のサイズ（バイト単位）。
            /// </summary>
            public int cbSize;
            /// <summary>
            /// 使用する情報を指定するフラグ。
            /// </summary>
            public CMIC fMask;
            /// <summary>
            /// オーナーウィンドウのハンドル。
            /// </summary>
            public IntPtr hwnd;
            /// <summary>
            /// 実行するコマンドの動詞（ANSI）。
            /// </summary>
            public IntPtr lpVerb;
            /// <summary>
            /// コマンドに渡すパラメータ（ANSI）。
            /// </summary>
            [MarshalAs(UnmanagedType.LPStr)]
            public string? lpParameters;
            /// <summary>
            /// 作業ディレクトリ（ANSI）。
            /// </summary>
            [MarshalAs(UnmanagedType.LPStr)]
            public string? lpDirectory;
            /// <summary>
            /// ウィンドウの表示方法（<see cref="SW"/>定数）。
            /// </summary>
            public int nShow;
            /// <summary>
            /// ホットキー。
            /// </summary>
            public int dwHotKey;
            /// <summary>
            /// アイコンハンドル。
            /// </summary>
            public IntPtr hIcon;
            /// <summary>
            /// タイトル（ANSI）。
            /// </summary>
            public IntPtr lpTitle;
            /// <summary>
            /// 実行するコマンドの動詞（Unicode）。
            /// </summary>
            public IntPtr lpVerbW;
            /// <summary>
            /// コマンドに渡すパラメータ（Unicode）。
            /// </summary>
            public IntPtr lpParametersW;
            /// <summary>
            /// 作業ディレクトリ（Unicode）。
            /// </summary>
            public IntPtr lpDirectoryW;
            /// <summary>
            /// タイトル（Unicode）。
            /// </summary>
            public IntPtr lpTitleW;
            /// <summary>
            /// メニューが表示された位置（スクリーン座標）。
            /// </summary>
            public POINT ptInvoke;
        }

        /// <summary>
        /// 2次元座標を表す構造体。
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            /// <summary>
            /// X座標。
            /// </summary>
            public int X;
            /// <summary>
            /// Y座標。
            /// </summary>
            public int Y;
        }

        /// <summary>
        /// Windowsシェルのコンテキストメニューを操作するためのCOMインターフェース。
        /// </summary>
        /// <remarks>
        /// このインターフェースは、ファイルやフォルダーに対するコンテキストメニューの構築と実行を提供します。
        /// </remarks>
        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("000214E4-0000-0000-C000-000000000046")]
        private interface IContextMenu
        {
            /// <summary>
            /// コンテキストメニューを構築します。
            /// </summary>
            /// <param name="hMenu">メニューハンドル。</param>
            /// <param name="indexMenu">メニュー項目を挿入する位置。</param>
            /// <param name="idCmdFirst">最初のコマンドID。</param>
            /// <param name="idCmdLast">最後のコマンドID。</param>
            /// <param name="uFlags">メニューフラグ。</param>
            /// <returns>追加されたメニュー項目の数。</returns>
            [PreserveSig]
            int QueryContextMenu(
                IntPtr hMenu,
                uint indexMenu,
                uint idCmdFirst,
                uint idCmdLast,
                CMF uFlags);

            /// <summary>
            /// メニューで選択されたコマンドを実行します。
            /// </summary>
            /// <param name="pici">コマンド実行情報を含む構造体。</param>
            void InvokeCommand(ref CMINVOKECOMMANDINFOEX pici);

            /// <summary>
            /// コマンドの説明文字列を取得します。
            /// </summary>
            /// <param name="idCmd">コマンドID。</param>
            /// <param name="uFlags">取得する文字列の種類を指定するフラグ。</param>
            /// <param name="pReserved">予約済みパラメータ。</param>
            /// <param name="pszName">文字列を受け取るバッファ。</param>
            /// <param name="cchMax">バッファのサイズ。</param>
            void GetCommandString(
                UIntPtr idCmd,
                uint uFlags,
                IntPtr pReserved,
                [MarshalAs(UnmanagedType.LPStr)] System.Text.StringBuilder pszName,
                int cchMax);
        }

        /// <summary>
        /// <see cref="IContextMenu"/>を拡張し、カスタムメニュー描画をサポートするCOMインターフェース。
        /// </summary>
        /// <remarks>
        /// このインターフェースは、<see cref="IContextMenu"/>の機能に加えて、メニュー関連のWindowsメッセージを処理する機能を提供します。
        /// </remarks>
        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("000214F4-0000-0000-C000-000000000046")]
        private interface IContextMenu2 : IContextMenu
        {
            // IContextMenu
            /// <inheritdoc cref="IContextMenu.QueryContextMenu"/>
            [PreserveSig]
            new int QueryContextMenu(
                IntPtr hMenu,
                uint indexMenu,
                uint idCmdFirst,
                uint idCmdLast,
                CMF uFlags);

            /// <inheritdoc cref="IContextMenu.InvokeCommand"/>
            new void InvokeCommand(ref CMINVOKECOMMANDINFOEX pici);

            /// <inheritdoc cref="IContextMenu.GetCommandString"/>
            new void GetCommandString(
                UIntPtr idCmd,
                uint uFlags,
                IntPtr pReserved,
                [MarshalAs(UnmanagedType.LPStr)] System.Text.StringBuilder pszName,
                int cchMax);

            // IContextMenu2
            /// <summary>
            /// メニュー関連のWindowsメッセージを処理します。
            /// </summary>
            /// <param name="uMsg">メッセージID。</param>
            /// <param name="wParam">メッセージの追加情報（WPARAM）。</param>
            /// <param name="lParam">メッセージの追加情報（LPARAM）。</param>
            /// <returns>メッセージ処理の結果値。</returns>
            [PreserveSig]
            int HandleMenuMsg(
                uint uMsg,
                IntPtr wParam,
                IntPtr lParam);
        }

        /// <summary>
        /// <see cref="IContextMenu2"/>を拡張し、より高度なメニュー描画機能を提供するCOMインターフェース。
        /// </summary>
        /// <remarks>
        /// <para>
        /// このインターフェースは、<see cref="IContextMenu2"/>の機能に加えて、メニュー関連のWindowsメッセージを処理し、
        /// 処理結果を返す機能を提供します。これにより、より高度なカスタムメニュー描画が可能になります。
        /// </para>
        /// <para>
        /// このインターフェースが利用可能な場合、<see cref="ProcessWindowMessage"/>がメニュー関連のメッセージを
        /// <see cref="HandleMenuMsg2"/>に転送します。
        /// </para>
        /// </remarks>
        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("BCFCE0A0-EC17-11D0-8D10-00A0C90F2719")]
        private interface IContextMenu3 : IContextMenu2
        {
            // IContextMenu
            /// <inheritdoc cref="IContextMenu.QueryContextMenu"/>
            [PreserveSig]
            new int QueryContextMenu(
                IntPtr hMenu,
                uint indexMenu,
                uint idCmdFirst,
                uint idCmdLast,
                CMF uFlags);

            /// <inheritdoc cref="IContextMenu.InvokeCommand"/>
            new void InvokeCommand(ref CMINVOKECOMMANDINFOEX pici);

            /// <inheritdoc cref="IContextMenu.GetCommandString"/>
            new void GetCommandString(
                UIntPtr idCmd,
                uint uFlags,
                IntPtr pReserved,
                [MarshalAs(UnmanagedType.LPStr)] System.Text.StringBuilder pszName,
                int cchMax);

            // IContextMenu2
            /// <inheritdoc cref="IContextMenu2.HandleMenuMsg"/>
            [PreserveSig]
            new int HandleMenuMsg(
                uint uMsg,
                IntPtr wParam,
                IntPtr lParam);

            // IContextMenu3
            /// <summary>
            /// メニュー関連のWindowsメッセージを処理し、処理結果を返します。
            /// </summary>
            /// <param name="uMsg">メッセージID。</param>
            /// <param name="wParam">メッセージの追加情報（WPARAM）。</param>
            /// <param name="lParam">メッセージの追加情報（LPARAM）。</param>
            /// <param name="plResult">メッセージ処理の結果値。</param>
            /// <returns>成功した場合は0（S_OK）、それ以外の場合はエラーコード。</returns>
            /// <remarks>
            /// このメソッドは、<see cref="ProcessWindowMessage"/>から呼び出され、以下のメッセージを処理します：
            /// <list type="bullet">
            /// <item><description>WM_INITMENUPOPUP</description></item>
            /// <item><description>WM_DRAWITEM</description></item>
            /// <item><description>WM_MEASUREITEM</description></item>
            /// <item><description>WM_MENUCHAR</description></item>
            /// </list>
            /// </remarks>
            [PreserveSig]
            int HandleMenuMsg2(
                uint uMsg,
                IntPtr wParam,
                IntPtr lParam,
                out IntPtr plResult);
        }

        /// <summary>
        /// Windowsシェルフォルダーを操作するためのCOMインターフェース。
        /// </summary>
        /// <remarks>
        /// このインターフェースは、フォルダー内のオブジェクト（ファイルやサブフォルダー）へのアクセスと操作を提供します。
        /// コンテキストメニューの取得には、<see cref="GetUIObjectOf"/>メソッドを使用して<see cref="IContextMenu"/>を取得します。
        /// </remarks>
        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("000214E6-0000-0000-C000-000000000046")]
        private interface IShellFolder
        {
            /// <summary>
            /// 表示名を解析してPIDLを取得します。
            /// </summary>
            void ParseDisplayName(
                IntPtr hwnd,
                IntPtr pbc,
                [MarshalAs(UnmanagedType.LPWStr)] string pszDisplayName,
                ref uint pchEaten,
                out IntPtr ppidl,
                ref uint pdwAttributes);

            /// <summary>
            /// フォルダー内のオブジェクトを列挙します。
            /// </summary>
            void EnumObjects(
                IntPtr hwnd,
                int grfFlags,
                out IntPtr ppenumIDList);

            /// <summary>
            /// 指定されたPIDLに対応するオブジェクトにバインドします。
            /// </summary>
            void BindToObject(
                IntPtr pidl,
                IntPtr pbc,
                [In] ref Guid riid,
                [MarshalAs(UnmanagedType.Interface)] out object ppv);

            /// <summary>
            /// 指定されたPIDLに対応するストレージオブジェクトにバインドします。
            /// </summary>
            void BindToStorage(
                IntPtr pidl,
                IntPtr pbc,
                [In] ref Guid riid,
                [MarshalAs(UnmanagedType.Interface)] out object ppv);

            /// <summary>
            /// 2つのPIDLを比較します。
            /// </summary>
            [PreserveSig]
            int CompareIDs(
                IntPtr lParam,
                IntPtr pidl1,
                IntPtr pidl2);

            /// <summary>
            /// ビューオブジェクトを作成します。
            /// </summary>
            void CreateViewObject(
                IntPtr hwndOwner,
                [In] ref Guid riid,
                [MarshalAs(UnmanagedType.Interface)] out object ppv);

            /// <summary>
            /// 指定されたPIDLの属性を取得します。
            /// </summary>
            void GetAttributesOf(
                uint cidl,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] IntPtr[] apidl,
                ref uint rgfInOut);

            /// <summary>
            /// 指定されたPIDLに対応するUIオブジェクト（<see cref="IContextMenu"/>など）を取得します。
            /// </summary>
            /// <param name="hwndOwner">オーナーウィンドウのハンドル。</param>
            /// <param name="cidl">PIDL配列の要素数。</param>
            /// <param name="apidl">PIDL配列。</param>
            /// <param name="riid">取得するインターフェースのIID（通常は<see cref="IID_IContextMenu"/>）。</param>
            /// <param name="rgfReserved">予約済みパラメータ。</param>
            /// <param name="ppv">取得したインターフェースへのポインタ。</param>
            /// <remarks>
            /// このメソッドは、コンテキストメニューを取得するために使用されます。
            /// </remarks>
            void GetUIObjectOf(
                IntPtr hwndOwner,
                uint cidl,
                [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] IntPtr[] apidl,
                [In] ref Guid riid,
                ref uint rgfReserved,
                [MarshalAs(UnmanagedType.Interface)] out object ppv);

            /// <summary>
            /// PIDLの表示名を取得します。
            /// </summary>
            void GetDisplayNameOf(
                IntPtr pidl,
                uint uFlags,
                IntPtr pName);

            /// <summary>
            /// PIDLの表示名を設定します。
            /// </summary>
            void SetNameOf(
                IntPtr hwnd,
                IntPtr pidl,
                [MarshalAs(UnmanagedType.LPWStr)] string pszName,
                uint uFlags,
                out IntPtr ppidlOut);
        }

        /// <summary>
        /// Windows API（Win32）のP/Invoke定義を提供するクラス。
        /// </summary>
        private static class Win32
        {
            /// <summary>
            /// ファイルパスをPIDLに変換します。
            /// </summary>
            /// <param name="pszName">変換するファイルパス。</param>
            /// <param name="pbc">バインドコンテキスト（通常は<see cref="IntPtr.Zero"/>）。</param>
            /// <param name="ppidl">取得したPIDLへのポインタ。</param>
            /// <param name="sfgaoIn">取得する属性のマスク。</param>
            /// <param name="psfgaoOut">取得した属性。</param>
            /// <returns>成功した場合は0（S_OK）、それ以外の場合はエラーコード。</returns>
            [DllImport("shell32.dll", CharSet = CharSet.Auto)]
            public static extern int SHParseDisplayName([MarshalAs(UnmanagedType.LPWStr)] string pszName, IntPtr pbc, out IntPtr ppidl, uint sfgaoIn, out uint psfgaoOut);

            /// <summary>
            /// PIDLの親フォルダーを取得します。
            /// </summary>
            /// <param name="pidl">PIDLへのポインタ。</param>
            /// <param name="riid">取得するインターフェースのIID（通常は<see cref="IID_IShellFolder"/>）。</param>
            /// <param name="ppv">取得した<see cref="IShellFolder"/>インターフェースへのポインタ。</param>
            /// <param name="ppidlLast">相対PIDL（最後のID）へのポインタ。</param>
            /// <returns>成功した場合は0（S_OK）、それ以外の場合はエラーコード。</returns>
            [DllImport("shell32.dll")]
            public static extern int SHBindToParent(IntPtr pidl, [In] ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out IShellFolder ppv, out IntPtr ppidlLast);

            /// <summary>
            /// PIDLの最後のID（相対PIDL）を取得します。
            /// </summary>
            /// <param name="pidl">PIDLへのポインタ。</param>
            /// <returns>最後のIDへのポインタ。</returns>
            [DllImport("shell32.dll")]
            public static extern IntPtr ILFindLastID(IntPtr pidl);

            /// <summary>
            /// COMタスクメモリアロケーターで割り当てられたメモリを解放します。
            /// </summary>
            /// <param name="pv">解放するメモリへのポインタ。</param>
            /// <remarks>
            /// PIDLなどのCOMオブジェクトで割り当てられたメモリを解放するために使用します。
            /// </remarks>
            [DllImport("ole32.dll")]
            public static extern void CoTaskMemFree(IntPtr pv);

            /// <summary>
            /// ポップアップメニューを作成します。
            /// </summary>
            /// <returns>作成されたメニューのハンドル。失敗した場合は<see cref="IntPtr.Zero"/>。</returns>
            [DllImport("user32.dll")]
            public static extern IntPtr CreatePopupMenu();

            /// <summary>
            /// メニューを破棄します。
            /// </summary>
            /// <param name="hMenu">破棄するメニューのハンドル。</param>
            /// <returns>成功した場合は<see langword="true"/>、失敗した場合は<see langword="false"/>。</returns>
            /// <remarks>
            /// メニューの使用が終了したら、必ずこのメソッドを呼び出してリソースを解放してください。
            /// </remarks>
            [DllImport("user32.dll")]
            public static extern bool DestroyMenu(IntPtr hMenu);

            /// <summary>
            /// ポップアップメニューを表示し、ユーザーの選択を待ちます。
            /// </summary>
            /// <param name="hMenu">表示するメニューのハンドル。</param>
            /// <param name="uFlags">メニューの表示方法を指定するフラグ（通常は<see cref="TPM.RETURNCMD"/>）。</param>
            /// <param name="x">メニューを表示するX座標（スクリーン座標）。</param>
            /// <param name="y">メニューを表示するY座標（スクリーン座標）。</param>
            /// <param name="hWnd">オーナーウィンドウのハンドル。</param>
            /// <param name="lptpm">追加のメニュー情報へのポインタ（通常は<see cref="IntPtr.Zero"/>）。</param>
            /// <returns>
            /// <paramref name="uFlags"/>に<see cref="TPM.RETURNCMD"/>が指定されている場合、選択されたメニュー項目のコマンドID。
            /// メニューがキャンセルされた場合は0。
            /// </returns>
            [DllImport("user32.dll")]
            public static extern uint TrackPopupMenuEx(IntPtr hMenu, TPM uFlags, int x, int y, IntPtr hWnd, IntPtr lptpm);
        }

        #endregion
    }
}


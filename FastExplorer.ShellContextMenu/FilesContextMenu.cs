// Copyright (c) Files Community
// Licensed under the MIT License.

using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using Vanara.InteropServices;
using Vanara.PInvoke;
using Vanara.Windows.Shell;
using Windows.Win32;
using Windows.Win32.Graphics.Gdi;
using Windows.Win32.UI.WindowsAndMessaging;
using Bitmap = System.Drawing.Bitmap;
using Image = System.Drawing.Image;

namespace FastExplorer.ShellContextMenu
{
	/// <summary>
	/// Provides a helper for Win32 context menu.
	/// Windowsのシェルコンテキストメニューを取得・表示するためのヘルパークラスです。
	/// ファイルとフォルダーで自動的に異なるメニューが取得されます。
	/// </summary>
	public partial class FilesContextMenu : Win32ContextMenu, IDisposable
	{
		private Shell32.IContextMenu _cMenu;

		private User32.SafeHMENU _hMenu;

		private readonly ThreadWithMessageQueue _owningThread;

		private readonly Func<string, bool>? _itemFilter;

		private readonly Dictionary<List<Win32ContextMenuItem>, Action> _loadSubMenuActions;

		// To detect redundant calls
		private bool disposedValue = false;

		public List<string> ItemsPath { get; }

		public User32.SafeHMENU MenuHandle => _hMenu;

		private FilesContextMenu(Shell32.IContextMenu cMenu, User32.SafeHMENU hMenu, IEnumerable<string> itemsPath, ThreadWithMessageQueue owningThread, Func<string, bool>? itemFilter)
		{
			_cMenu = cMenu;
			_hMenu = hMenu;
			_owningThread = owningThread;
			_itemFilter = itemFilter;
			_loadSubMenuActions = [];

			ItemsPath = itemsPath.ToList();
			Items = [];
		}

		public async static Task<bool> InvokeVerb(string verb, params string[] filePaths)
		{
			using var cMenu = await GetContextMenuForFiles(filePaths, 0x00000001); // CMF_DEFAULTONLY

			return cMenu is not null && await cMenu.InvokeVerb(verb);
		}

		public async Task<bool> InvokeVerb(string? verb)
		{
			if (string.IsNullOrEmpty(verb))
				return false;

			var item = Items.FirstOrDefault(x => x.CommandString == verb);
			if (item is not null && item.ID >= 0)
				// Prefer invocation by ID
				return await InvokeItem(item.ID).ConfigureAwait(false);

			try
			{
				var pici = new Shell32.CMINVOKECOMMANDINFOEX
				{
					lpVerb = new SafeResourceId(verb, CharSet.Ansi),
					nShow = ShowWindowCommand.SW_SHOWNORMAL,
				};

				pici.cbSize = (uint)Marshal.SizeOf(pici);

				await _owningThread.PostMethod(() => _cMenu.InvokeCommand(pici)).ConfigureAwait(false);
				
				// BringToForegroundは削除（不要なウィンドウが表示される問題を回避）

				return true;
			}
			catch (Exception ex) when (ex is COMException or UnauthorizedAccessException)
			{
				System.Diagnostics.Debug.WriteLine(ex);
			}

			return false;
		}

		public async Task<bool> InvokeItem(int itemID, string? workingDirectory = null)
		{
			if (itemID < 0)
				return false;

			// 既に破棄されている場合は処理しない
			if (disposedValue)
			{
				System.Diagnostics.Debug.WriteLine("InvokeItem: FilesContextMenu is already disposed");
				return false;
			}

			// _cMenuがnullの場合は処理しない
			if (_cMenu == null)
			{
				System.Diagnostics.Debug.WriteLine("InvokeItem: _cMenu is null");
				return false;
			}

			try
			{
				var pici = new Shell32.CMINVOKECOMMANDINFOEX
				{
					lpVerb = Macros.MAKEINTRESOURCE(itemID),
					nShow = ShowWindowCommand.SW_SHOWNORMAL,
				};

				pici.cbSize = (uint)Marshal.SizeOf(pici);
				if (workingDirectory is not null)
					pici.lpDirectoryW = workingDirectory;

				// メニュー実行を実行スレッドで実行し、結果を待つ
				// ConfigureAwait(false)を使用して、同期コンテキストに戻らないようにする（UIブロッキングを回避）
				await _owningThread.PostMethod(() =>
				{
					// 実行中に破棄された場合は例外をスロー
					if (disposedValue || _cMenu == null)
					{
						throw new ObjectDisposedException(nameof(FilesContextMenu));
					}
					_cMenu.InvokeCommand(pici);
				}).ConfigureAwait(false);

				// BringToForegroundは削除（不要なウィンドウが表示される問題を回避）
				
				return true;
			}
			catch (OperationCanceledException)
			{
				// 操作がキャンセルされた場合は無視
				System.Diagnostics.Debug.WriteLine("InvokeItem: Operation was canceled");
				return false;
			}
			catch (ObjectDisposedException)
			{
				// 既に破棄されている場合は無視
				System.Diagnostics.Debug.WriteLine("InvokeItem: FilesContextMenu was disposed");
				return false;
			}
			catch (Exception ex) when (ex is COMException or UnauthorizedAccessException)
			{
				System.Diagnostics.Debug.WriteLine($"InvokeItem: {ex}");
				return false;
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.WriteLine($"InvokeItem: Unexpected error: {ex}");
				return false;
			}
		}

		private async Task<bool> InvokeItemInternal(Shell32.CMINVOKECOMMANDINFOEX pici, List<Vanara.PInvoke.HWND> currentWindows)
		{
			// currentWindowsパラメータは未使用（BringToForegroundを削除したため）
			try
			{
				await _owningThread.PostMethod(() =>
				{
					// 実行中に破棄された場合は例外をスロー
					if (disposedValue || _cMenu == null)
					{
						throw new ObjectDisposedException(nameof(FilesContextMenu));
					}
					_cMenu.InvokeCommand(pici);
				}).ConfigureAwait(false);
				
				// BringToForegroundは削除（不要なウィンドウが表示される問題を回避）
				
				return true;
			}
			catch (OperationCanceledException)
			{
				// 操作がキャンセルされた場合は無視
				System.Diagnostics.Debug.WriteLine("InvokeItemInternal: Operation was canceled");
				return false;
			}
			catch (ObjectDisposedException)
			{
				// 既に破棄されている場合は無視
				System.Diagnostics.Debug.WriteLine("InvokeItemInternal: FilesContextMenu was disposed");
				return false;
			}
			catch (Exception ex) when (ex is COMException or UnauthorizedAccessException)
			{
				System.Diagnostics.Debug.WriteLine($"InvokeItemInternal: {ex}");
				return false;
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.WriteLine($"InvokeItemInternal: Unexpected error: {ex}");
				return false;
			}
		}

		/// <summary>
		/// Shows the native Windows context menu using TrackPopupMenuEx
		/// </summary>
		/// <param name="ownerHwnd">Owner window handle</param>
		/// <param name="x">X coordinate in screen coordinates</param>
		/// <param name="y">Y coordinate in screen coordinates</param>
		/// <returns>Selected menu item ID, or 0 if cancelled</returns>
		public async Task<uint> ShowNativeMenu(IntPtr ownerHwnd, int x, int y)
		{
			if (_hMenu == null || _hMenu.IsInvalid)
				return 0;

			// 既に破棄されている場合は処理しない
			if (disposedValue)
			{
				System.Diagnostics.Debug.WriteLine("ShowNativeMenu: FilesContextMenu is already disposed");
				return 0;
			}

			// _cMenuがnullの場合は処理しない
			if (_cMenu == null)
			{
				System.Diagnostics.Debug.WriteLine("ShowNativeMenu: _cMenu is null");
				return 0;
			}

			return await _owningThread.PostMethod<uint>(() =>
			{
				try
				{
					// 実行中に破棄された場合は処理しない
					if (disposedValue || _cMenu == null)
					{
						System.Diagnostics.Debug.WriteLine("ShowNativeMenu: FilesContextMenu was disposed during execution");
						return 0;
					}

					const uint idCmdFirst = 1;

					// Show native Windows context menu
					uint selected = User32.TrackPopupMenuEx(
						_hMenu,
						User32.TrackPopupMenuFlags.TPM_RETURNCMD,
						x,
						y,
						ownerHwnd,
						null);

					if (selected != 0)
					{
						// 実行中に破棄された場合は処理しない
						if (disposedValue || _cMenu == null)
						{
							System.Diagnostics.Debug.WriteLine("ShowNativeMenu: FilesContextMenu was disposed after menu selection");
							return 0;
						}

						// Invoke the selected command
						var itemID = (int)(selected - idCmdFirst);
						var pici = new Shell32.CMINVOKECOMMANDINFOEX
						{
							lpVerb = Macros.MAKEINTRESOURCE(itemID),
							nShow = ShowWindowCommand.SW_SHOWNORMAL,
						};
						pici.cbSize = (uint)Marshal.SizeOf(pici);
						_cMenu.InvokeCommand(pici);
					}

					return selected;
				}
				catch (ObjectDisposedException)
				{
					// 既に破棄されている場合は無視
					System.Diagnostics.Debug.WriteLine("ShowNativeMenu: FilesContextMenu was disposed");
					return 0;
				}
				catch (Exception ex) when (ex is COMException or UnauthorizedAccessException)
				{
					System.Diagnostics.Debug.WriteLine($"ShowNativeMenu: {ex}");
					return 0;
				}
				catch (Exception ex)
				{
					System.Diagnostics.Debug.WriteLine($"Error showing native menu: {ex}");
					return 0;
				}
			});
		}

		/// <summary>
		/// Gets the Windows shell context menu for the specified file paths.
		/// Windowsのシェルコンテキストメニューを取得します。
		/// ファイルとフォルダーで自動的に異なるメニューが取得されます。
		/// </summary>
		/// <param name="filePathList">Array of file paths (ファイルパスの配列)</param>
		/// <param name="flags">Context menu flags (コンテキストメニューフラグ)</param>
		/// <param name="itemFilter">Optional filter for menu items (メニュー項目のフィルター)</param>
		/// <returns>FilesContextMenu instance or null (FilesContextMenuインスタンスまたはnull)</returns>
		public async static Task<FilesContextMenu?> GetContextMenuForFiles(string[] filePathList, uint flags, Func<string, bool>? itemFilter = null)
		{
			var owningThread = new ThreadWithMessageQueue();

			return await owningThread.PostMethod<FilesContextMenu>(() =>
			{
				var shellItems = new List<ShellItem>();

				try
				{
					// ShellItemはファイルパスから自動的にファイル/フォルダーを判定し、
					// Windowsのシェルが提供する適切なコンテキストメニューを取得します
					foreach (var filePathItem in filePathList.Where(x => !string.IsNullOrEmpty(x)))
						shellItems.Add(new ShellItem(filePathItem));

					return GetContextMenuForFiles([.. shellItems], flags, owningThread, itemFilter);
				}
				catch
				{
					// Return empty context menu
					return null;
				}
				finally
				{
					foreach (var item in shellItems)
						item.Dispose();
				}
			});
		}

		public async static Task<FilesContextMenu?> GetContextMenuForFiles(ShellItem[] shellItems, uint flags, Func<string, bool>? itemFilter = null)
		{
			var owningThread = new ThreadWithMessageQueue();

			return await owningThread.PostMethod<FilesContextMenu>(() => GetContextMenuForFiles(shellItems, flags, owningThread, itemFilter));
		}

		private static FilesContextMenu? GetContextMenuForFiles(ShellItem[] shellItems, uint flags, ThreadWithMessageQueue owningThread, Func<string, bool>? itemFilter = null)
		{
			if (!shellItems.Any())
				return null;

			try
			{
				// NOTE: The items are all in the same folder
				using var sf = shellItems[0].Parent;

				Shell32.IContextMenu menu = sf.GetChildrenUIObjects<Shell32.IContextMenu>(default, shellItems);
				var hMenu = User32.CreatePopupMenu();
				menu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, (Shell32.CMF)flags);
				var contextMenu = new FilesContextMenu(menu, hMenu, shellItems.Select(x => x.ParsingName), owningThread, itemFilter);
				contextMenu.EnumMenuItems(hMenu, contextMenu.Items);

				return contextMenu;
			}
			catch (COMException)
			{
				// Return empty context menu
				return null;
			}
		}

		private void EnumMenuItems(Vanara.PInvoke.HMENU hMenu, List<Win32ContextMenuItem> menuItemsResult, bool loadSubenus = false)
		{
			var itemCount = User32.GetMenuItemCount(hMenu);

			var menuItemInfo = new User32.MENUITEMINFO()
			{
				fMask =
					User32.MenuItemInfoMask.MIIM_BITMAP |
					User32.MenuItemInfoMask.MIIM_FTYPE |
					User32.MenuItemInfoMask.MIIM_STRING |
					User32.MenuItemInfoMask.MIIM_ID |
					User32.MenuItemInfoMask.MIIM_SUBMENU,
			};

			menuItemInfo.cbSize = (uint)Marshal.SizeOf(menuItemInfo);

			for (uint index = 0; index < itemCount; index++)
			{
				var menuItem = new ContextMenuItem();
				var container = new SafeCoTaskMemString(512);
				var cMenu2 = _cMenu as Shell32.IContextMenu2;

				menuItemInfo.dwTypeData = (IntPtr)container;

				// See also, https://devblogs.microsoft.com/oldnewthing/20040928-00/?p=37723
				menuItemInfo.cch = (uint)container.Capacity - 1;

				var result = User32.GetMenuItemInfo(hMenu, index, true, ref menuItemInfo);
				if (!result)
				{
					container.Dispose();
					continue;
				}

				menuItem.Type = (uint)menuItemInfo.fType;

				// wID - idCmdFirst
				menuItem.ID = (int)(menuItemInfo.wID - 1);

				// MFT_STRING = 0x00000000
				if ((menuItem.Type & 0x000000FF) == 0)
				{
					System.Diagnostics.Debug.WriteLine("Item {0} ({1}): {2}", index, menuItemInfo.wID, menuItemInfo.dwTypeData);

					menuItem.Label = menuItemInfo.dwTypeData;
					menuItem.CommandString = GetCommandString(_cMenu, menuItemInfo.wID - 1);

					if (_itemFilter is not null && (_itemFilter(menuItem.CommandString) || _itemFilter(menuItem.Label)))
					{
						// Skip items implemented in UWP
						container.Dispose();
						continue;
					}

					if (menuItemInfo.hbmpItem != HBITMAP.NULL && !Enum.IsDefined(typeof(HBITMAP_HMENU), ((IntPtr)menuItemInfo.hbmpItem).ToInt64()))
					{
						using System.Drawing.Bitmap? bitmap = Win32Helper.GetBitmapFromHBitmap(menuItemInfo.hbmpItem);

						if (bitmap is not null)
						{
							// Make the icon background transparent
							bitmap.MakeTransparent();

							byte[] bitmapData = (byte[])new System.Drawing.ImageConverter().ConvertTo(bitmap, typeof(byte[]));
							menuItem.Icon = bitmapData;
						}
					}

					if (menuItemInfo.hSubMenu != Vanara.PInvoke.HMENU.NULL)
					{
						System.Diagnostics.Debug.WriteLine("Item {0}: has submenu", index);
						var subItems = new List<Win32ContextMenuItem>();
						var hSubMenu = menuItemInfo.hSubMenu;

						if (loadSubenus)
							LoadSubMenu();
						else
							_loadSubMenuActions.Add(subItems, LoadSubMenu);

						menuItem.SubItems = subItems;

						System.Diagnostics.Debug.WriteLine("Item {0}: done submenu", index);

						void LoadSubMenu()
						{
							try
							{
								cMenu2?.HandleMenuMsg((uint)User32.WindowMessage.WM_INITMENUPOPUP, (IntPtr)hSubMenu, new IntPtr(index));
							}
							catch (Exception ex) when (ex is InvalidCastException or ArgumentException)
							{
								// TODO: Investigate why this exception happen
								System.Diagnostics.Debug.WriteLine(ex);
							}
							catch (Exception ex) when (ex is COMException or NotImplementedException)
							{
								// Only for dynamic/owner drawn? (open with, etc)
							}

							EnumMenuItems(hSubMenu, subItems, true);
						}
					}
				}
				else
				{
					System.Diagnostics.Debug.WriteLine("Item {0}: {1}", index, menuItemInfo.fType.ToString());
				}

				container.Dispose();
				menuItemsResult.Add(menuItem);
			}
		}

		public Task<bool> LoadSubMenu(List<Win32ContextMenuItem> subItems)
		{
			// 既に破棄されている場合は処理しない
			if (disposedValue)
			{
				System.Diagnostics.Debug.WriteLine("LoadSubMenu: FilesContextMenu is already disposed");
				return Task.FromResult(false);
			}

			if (_loadSubMenuActions.Remove(subItems, out var loadSubMenuAction))
			{
				return _owningThread.PostMethod<bool>(() =>
				{
					try
					{
						// 実行中に破棄された場合は例外をスロー
						if (disposedValue)
						{
							throw new ObjectDisposedException(nameof(FilesContextMenu));
						}
						loadSubMenuAction!();
						return true;
					}
					catch (OperationCanceledException)
					{
						// 操作がキャンセルされた場合は無視
						System.Diagnostics.Debug.WriteLine("LoadSubMenu: Operation was canceled");
						return false;
					}
					catch (ObjectDisposedException)
					{
						// 既に破棄されている場合は無視
						System.Diagnostics.Debug.WriteLine("LoadSubMenu: FilesContextMenu was disposed");
						return false;
					}
					catch (Exception ex)
					{
						System.Diagnostics.Debug.WriteLine($"LoadSubMenu: Error: {ex}");
						return false;
					}
				});
			}
			else
			{
				return Task.FromResult(false);
			}
		}

		private static string? GetCommandString(Shell32.IContextMenu cMenu, uint offset, Shell32.GCS flags = Shell32.GCS.GCS_VERBW)
		{
			// A workaround to avoid an AccessViolationException on some items,
			// notably the "Run with graphic processor" menu item of NVIDIA cards
			if (offset > 5000)
			{
				return null;
			}

			SafeCoTaskMemString? commandString = null;

			try
			{
				commandString = new SafeCoTaskMemString(512);
				cMenu.GetCommandString(new IntPtr(offset), flags, IntPtr.Zero, commandString, (uint)commandString.Capacity - 1);
				System.Diagnostics.Debug.WriteLine("Verb {0}: {1}", offset, commandString);

				return commandString.ToString();
			}
			catch (Exception ex) when (ex is InvalidCastException or ArgumentException)
			{
				// TODO: Investigate why this exception happen
				System.Diagnostics.Debug.WriteLine(ex);

				return null;
			}
			catch (Exception ex) when (ex is COMException or NotImplementedException)
			{
				// Not every item has an associated verb
				return null;
			}
			finally
			{
				commandString?.Dispose();
			}
		}

		protected virtual void Dispose(bool disposing)
		{
			if (!disposedValue)
			{
				// 破棄フラグを設定（新しいタスクの開始を防ぐ）
				disposedValue = true;

				if (disposing)
				{
					// TODO: Dispose managed state (managed objects)
					if (Items is not null)
					{
						foreach (var si in Items)
						{
							(si as IDisposable)?.Dispose();
						}

						Items = null;
					}
				}

				// TODO: Free unmanaged resources (unmanaged objects) and override a finalizer below
				if (_hMenu is not null)
				{
					User32.DestroyMenu(_hMenu);
					_hMenu = null;
				}
				if (_cMenu is not null)
				{
					Marshal.ReleaseComObject(_cMenu);
					_cMenu = null;
				}

				_owningThread.Dispose();
			}
		}

		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		~FilesContextMenu()
		{
			Dispose(false);
		}
	}
}


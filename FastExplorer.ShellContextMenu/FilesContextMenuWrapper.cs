// Copyright (c) Files Community
// Licensed under the MIT License.

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Wpf.Ui.Controls;
using MenuItem = Wpf.Ui.Controls.MenuItem;
using TextBlock = System.Windows.Controls.TextBlock;

namespace FastExplorer.ShellContextMenu
{
	/// <summary>
	/// Wrapper for FilesContextMenu to display Windows shell context menu in WPF applications.
	/// Windowsのシェルコンテキストメニューをラップして表示します。
	/// ファイルとフォルダーで自動的に異なるメニューが表示されます。
	/// </summary>
	public class FilesContextMenuWrapper
	{
		private static readonly object _lockObject = new object();
		private static bool _isShowingMenu = false;

		/// <summary>
		/// Shows the Windows shell context menu for the specified file paths at the specified screen coordinates.
		/// Windowsのシェルコンテキストメニューを表示します。
		/// ファイルとフォルダーで自動的に異なるメニューが表示されます。
		/// </summary>
		/// <param name="filePaths">Array of file paths (ファイルパスの配列)</param>
		/// <param name="ownerWindow">Owner window (オーナーウィンドウ)</param>
		/// <param name="x">X coordinate in screen coordinates (スクリーン座標のX座標)</param>
		/// <param name="y">Y coordinate in screen coordinates (スクリーン座標のY座標)</param>
		/// <param name="onClosed">Optional callback when the menu is closed (メニューが閉じられたときのコールバック)</param>
		public static void ShowContextMenu(string[] filePaths, Window ownerWindow, int x, int y, Action? onClosed = null)
		{
			if (filePaths == null || filePaths.Length == 0)
				return;

			if (ownerWindow == null)
				return;

			// 既にメニューを表示中の場合は処理しない（重複実行を防ぐ）
			lock (_lockObject)
			{
				if (_isShowingMenu)
				{
					System.Diagnostics.Debug.WriteLine("FilesContextMenuWrapper: Menu is already showing, skipping");
					return;
				}
				_isShowingMenu = true;
			}

			// UIスレッドで非同期処理を実行
			_ = ShowContextMenuAsync(filePaths, ownerWindow, x, y, onClosed);
		}

		private static async Task ShowContextMenuAsync(string[] filePaths, Window ownerWindow, int x, int y, Action? onClosed = null)
		{
			FilesContextMenu? contextMenu = null;
			try
			{
				System.Diagnostics.Debug.WriteLine($"FilesContextMenuWrapper: Getting context menu for: {string.Join(", ", filePaths)}");

				// Get context menu from Files implementation
				contextMenu = await FilesContextMenu.GetContextMenuForFiles(filePaths, 0x00000000); // CMF_NORMAL

				if (contextMenu == null)
				{
					System.Diagnostics.Debug.WriteLine("FilesContextMenuWrapper: Context menu is null");
					lock (_lockObject)
					{
						_isShowingMenu = false;
					}
					onClosed?.Invoke();
					return;
				}

				System.Diagnostics.Debug.WriteLine("FilesContextMenuWrapper: Building WPF context menu");

				// テーマカラーを取得
				var backgroundBrush = ownerWindow.TryFindResource("ApplicationBackgroundBrush") as System.Windows.Media.Brush
					?? ownerWindow.TryFindResource("ControlFillColorDefaultBrush") as System.Windows.Media.Brush
					?? System.Windows.Media.Brushes.White;
				var borderBrush = ownerWindow.TryFindResource("ControlStrokeColorDefaultBrush") as System.Windows.Media.Brush
					?? System.Windows.Media.Brushes.LightGray;
				var foregroundBrush = ownerWindow.TryFindResource("TextFillColorPrimaryBrush") as System.Windows.Media.Brush
					?? System.Windows.Media.Brushes.Black;

				// Build WPF ContextMenu from Win32 context menu items
				var wpfContextMenu = new ContextMenu
				{
					// Filesアプリのような見た目にするためのスタイル設定（テーマカラーを使用）
					Background = backgroundBrush,
					BorderBrush = borderBrush,
					BorderThickness = new Thickness(1),
					Padding = new Thickness(4),
					MinWidth = 200
				};

				// contextMenuをTagに保存して、メニューが閉じられるまで保持する
				wpfContextMenu.Tag = contextMenu;
				contextMenu = null; // 所有権をwpfContextMenuに移す

				BuildMenuItems(wpfContextMenu.Items, (FilesContextMenu)wpfContextMenu.Tag, ((FilesContextMenu)wpfContextMenu.Tag).Items, foregroundBrush);

				// メニューが閉じられたときにcontextMenuを破棄し、フラグをリセット
				wpfContextMenu.Closed += (s, e) =>
				{
					lock (_lockObject)
					{
						_isShowingMenu = false;
					}

					if (s is ContextMenu menu && menu.Tag is FilesContextMenu cm)
					{
						// メニュー項目の実行が完了するまで待ってから破棄する（非同期）
						// FilesContextMenu.Dispose()内で実行中のタスクを待つため、ここでは即座にDisposeを呼び出せる
						_ = Task.Run(async () =>
						{
							try
							{
								// Dispose()内で実行中のタスクを待つため、即座に呼び出しても安全
								// ただし、UIスレッドをブロックしないように非同期で実行
								cm.Dispose();
							}
							catch (Exception ex)
							{
								System.Diagnostics.Debug.WriteLine($"FilesContextMenuWrapper: Error disposing context menu: {ex}");
							}
							finally
							{
								// UIスレッドでTagをクリア
								Application.Current.Dispatcher.Invoke(() =>
								{
									menu.Tag = null;
								});
							}
						});
					}

					// コールバックを呼び出す
					onClosed?.Invoke();
				};

				// メニューが開かれたことを確認（エラー時のフラグリセット用）
				wpfContextMenu.Opened += (s, e) =>
				{
					System.Diagnostics.Debug.WriteLine("FilesContextMenuWrapper: Context menu opened");
				};

				// Show WPF context menu
				wpfContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.AbsolutePoint;
				wpfContextMenu.PlacementRectangle = new Rect(x, y, 0, 0);
				wpfContextMenu.IsOpen = true;
			}
			catch (OperationCanceledException)
			{
				// 操作がキャンセルされた場合は無視
				System.Diagnostics.Debug.WriteLine("FilesContextMenuWrapper: Operation was canceled");
				lock (_lockObject)
				{
					_isShowingMenu = false;
				}
				onClosed?.Invoke();
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.WriteLine($"FilesContextMenuWrapper: Error showing context menu: {ex}");
				lock (_lockObject)
				{
					_isShowingMenu = false;
				}
				// エラーが発生した場合は、contextMenuを破棄
				if (contextMenu != null)
				{
					try
					{
						contextMenu.Dispose();
					}
					catch
					{
						// 破棄時のエラーは無視
					}
				}
				onClosed?.Invoke();
			}
		}

		private static void BuildMenuItems(ItemCollection items, FilesContextMenu contextMenu, List<Win32ContextMenuItem> menuItems, System.Windows.Media.Brush? foregroundBrush = null)
		{
			foreach (var menuItem in menuItems)
			{
				// MFT_SEPARATOR = 0x00000800
				if ((menuItem.Type & 0x00000800) != 0)
				{
					items.Add(new System.Windows.Controls.Separator());
				}
				else
				{
					// Filesアプリのような見た目にする
					var labelText = menuItem.Label?.Replace("&", "") ?? "";
					var wpfMenuItem = new MenuItem
					{
						Header = new StackPanel
						{
							Orientation = Orientation.Horizontal,
							Children =
							{
								new TextBlock
								{
									Text = labelText,
									VerticalAlignment = VerticalAlignment.Center,
									FontSize = 14,
									Margin = new Thickness(8, 0, 0, 0),
									Foreground = foregroundBrush
								}
							}
						},
						Padding = new Thickness(8, 6, 8, 6),
						MinHeight = 32,
						Tag = contextMenu // contextMenuをTagに保存して、後でアクセスできるようにする
					};

					// Add icon if available
					if (menuItem.Icon != null && menuItem.Icon.Length > 0)
					{
						try
						{
							using var ms = new MemoryStream(menuItem.Icon);
							var bitmap = new BitmapImage();
							bitmap.BeginInit();
							bitmap.StreamSource = ms;
							bitmap.CacheOption = BitmapCacheOption.OnLoad;
							bitmap.EndInit();
							bitmap.Freeze();

							var icon = new System.Windows.Controls.Image
							{
								Source = bitmap,
								Width = 16,
								Height = 16,
								Margin = new Thickness(8, 0, 12, 0),
								VerticalAlignment = VerticalAlignment.Center
							};

							((StackPanel)wpfMenuItem.Header).Children.Insert(0, icon);
						}
						catch (Exception ex)
						{
							System.Diagnostics.Debug.WriteLine($"Error loading icon: {ex}");
						}
					}

					// Handle submenu
					if (menuItem.SubItems != null && menuItem.SubItems.Count > 0)
					{
						// Load submenu asynchronously (fire and forget)
						_ = LoadSubMenuAsync(wpfMenuItem, contextMenu, menuItem.SubItems, foregroundBrush);
					}
					else
					{
						// Set command to invoke menu item (Filesと同じ動作)
						var menuItemCopy = menuItem; // Capture for closure
						wpfMenuItem.Click += async (s, e) =>
						{
							try
							{
								// メニューをすぐに閉じる
								if (s is MenuItem item)
								{
									var contextMenu = item.Parent as ContextMenu ?? 
										(item.Parent as MenuItem)?.Parent as ContextMenu;
									if (contextMenu != null)
									{
										contextMenu.IsOpen = false;
									}
								}

								if (s is MenuItem menuItem && menuItem.Tag is FilesContextMenu cm)
								{
									await InvokeShellMenuItemAsync(cm, menuItemCopy);
								}
							}
							catch (ObjectDisposedException)
							{
								// contextMenuが既に破棄されている場合は無視
								System.Diagnostics.Debug.WriteLine("MenuItem Click: ContextMenu was disposed");
							}
							catch (OperationCanceledException)
							{
								// 操作がキャンセルされた場合は無視
								System.Diagnostics.Debug.WriteLine("MenuItem Click: Operation was canceled");
							}
							catch (Exception ex) when (ex is COMException or UnauthorizedAccessException)
							{
								// COM例外やアクセス権限例外は無視
								System.Diagnostics.Debug.WriteLine($"MenuItem Click: COM/Access exception: {ex}");
							}
							catch (Exception ex)
							{
								// その他の予期しない例外も無視してアプリケーションをクラッシュさせない
								System.Diagnostics.Debug.WriteLine($"MenuItem Click: Error invoking menu item: {ex}");
							}
						};
					}

					items.Add(wpfMenuItem);
				}
			}
		}

		private static async Task LoadSubMenuAsync(MenuItem parentMenuItem, FilesContextMenu contextMenu, List<Win32ContextMenuItem> subItems, System.Windows.Media.Brush? foregroundBrush = null)
		{
			try
			{
				// Load submenu if needed
				var loadResult = await contextMenu.LoadSubMenu(subItems);
				if (!loadResult)
				{
					System.Diagnostics.Debug.WriteLine("LoadSubMenuAsync: Failed to load submenu");
					return;
				}

				// テーマカラーを取得（親メニューから継承）
				if (foregroundBrush == null)
				{
					foregroundBrush = Application.Current.TryFindResource("TextFillColorPrimaryBrush") as System.Windows.Media.Brush
						?? System.Windows.Media.Brushes.Black;
				}

				// Build submenu items
				foreach (var subItem in subItems)
				{
					// MFT_SEPARATOR = 0x00000800
					if ((subItem.Type & 0x00000800) != 0)
					{
						parentMenuItem.Items.Add(new System.Windows.Controls.Separator());
					}
					else
					{
						// Filesアプリのような見た目にする
						var subLabelText = subItem.Label?.Replace("&", "") ?? "";
						var wpfSubMenuItem = new MenuItem
						{
							Header = new StackPanel
							{
								Orientation = Orientation.Horizontal,
								Children =
								{
									new TextBlock
									{
										Text = subLabelText,
										VerticalAlignment = VerticalAlignment.Center,
										FontSize = 14,
										Margin = new Thickness(8, 0, 0, 0),
										Foreground = foregroundBrush
									}
								}
							},
							Padding = new Thickness(8, 6, 8, 6),
							MinHeight = 32,
							Tag = contextMenu // contextMenuをTagに保存して、後でアクセスできるようにする
						};

						// Add icon if available
						if (subItem.Icon != null && subItem.Icon.Length > 0)
						{
							try
							{
								using var ms = new MemoryStream(subItem.Icon);
								var bitmap = new BitmapImage();
								bitmap.BeginInit();
								bitmap.StreamSource = ms;
								bitmap.CacheOption = BitmapCacheOption.OnLoad;
								bitmap.EndInit();
								bitmap.Freeze();

								var icon = new System.Windows.Controls.Image
								{
									Source = bitmap,
									Width = 16,
									Height = 16,
									Margin = new Thickness(8, 0, 12, 0),
									VerticalAlignment = VerticalAlignment.Center
								};

								((StackPanel)wpfSubMenuItem.Header).Children.Insert(0, icon);
							}
							catch (Exception ex)
							{
								System.Diagnostics.Debug.WriteLine($"Error loading submenu icon: {ex}");
							}
						}

						// Handle nested submenu
						if (subItem.SubItems != null && subItem.SubItems.Count > 0)
						{
							// Load nested submenu asynchronously (fire and forget)
							_ = LoadSubMenuAsync(wpfSubMenuItem, contextMenu, subItem.SubItems, foregroundBrush);
						}
						else
						{
							// Set command to invoke menu item (Filesと同じ動作)
							var subItemCopy = subItem; // Capture for closure
							wpfSubMenuItem.Click += async (s, e) =>
							{
								try
								{
									// メニューをすぐに閉じる
									if (s is MenuItem item)
									{
										var contextMenu = item.Parent as ContextMenu ?? 
											(item.Parent as MenuItem)?.Parent as ContextMenu;
										if (contextMenu != null)
										{
											contextMenu.IsOpen = false;
										}
									}

									if (s is MenuItem menuItem && menuItem.Tag is FilesContextMenu cm)
									{
										await InvokeShellMenuItemAsync(cm, subItemCopy);
									}
								}
								catch (ObjectDisposedException)
								{
									// contextMenuが既に破棄されている場合は無視
									System.Diagnostics.Debug.WriteLine("SubMenuItem Click: ContextMenu was disposed");
								}
								catch (OperationCanceledException)
								{
									// 操作がキャンセルされた場合は無視
									System.Diagnostics.Debug.WriteLine("SubMenuItem Click: Operation was canceled");
								}
								catch (Exception ex) when (ex is COMException or UnauthorizedAccessException)
								{
									// COM例外やアクセス権限例外は無視
									System.Diagnostics.Debug.WriteLine($"SubMenuItem Click: COM/Access exception: {ex}");
								}
								catch (Exception ex)
								{
									// その他の予期しない例外も無視してアプリケーションをクラッシュさせない
									System.Diagnostics.Debug.WriteLine($"SubMenuItem Click: Error invoking submenu item: {ex}");
								}
							};
						}

						parentMenuItem.Items.Add(wpfSubMenuItem);
					}
				}
			}
			catch (OperationCanceledException)
			{
				// 操作がキャンセルされた場合は無視
				System.Diagnostics.Debug.WriteLine("LoadSubMenuAsync: Operation was canceled");
			}
			catch (ObjectDisposedException)
			{
				// contextMenuが既に破棄されている場合は無視
				System.Diagnostics.Debug.WriteLine("LoadSubMenuAsync: ContextMenu was disposed");
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.WriteLine($"LoadSubMenuAsync: Error loading submenu: {ex}");
			}
		}

		/// <summary>
		/// Filesアプリと同じ動作でシェルメニュー項目を実行します
		/// </summary>
		private static async Task InvokeShellMenuItemAsync(FilesContextMenu contextMenu, Win32ContextMenuItem menuItem)
		{
			if (menuItem == null || contextMenu == null)
				return;

			try
			{
				var menuId = menuItem.ID;
				if (menuId < 0)
					return;

				var verb = menuItem.CommandString;
				var firstPath = contextMenu.ItemsPath?.FirstOrDefault();

				// Filesアプリと同じ特別なケースの処理
				switch (verb)
				{
					case "install":
						// フォントのインストール（簡易実装）
						await contextMenu.InvokeItem(menuId);
						break;

					case "installAllUsers":
						// 全ユーザー向けフォントのインストール（簡易実装）
						await contextMenu.InvokeItem(menuId);
						break;

					case "mount":
						// VHDのマウント（簡易実装）
						await contextMenu.InvokeItem(menuId);
						break;

					case "format":
						// ドライブのフォーマット（簡易実装）
						await contextMenu.InvokeItem(menuId);
						break;

					case "Windows.PowerShell.Run":
						// PowerShellスクリプトの実行
						var workingDir = !string.IsNullOrEmpty(firstPath) && firstPath.EndsWith(".ps1", StringComparison.OrdinalIgnoreCase)
							? System.IO.Path.GetDirectoryName(firstPath)
							: null;
						await contextMenu.InvokeItem(menuId, workingDir);
						break;

					default:
						// 通常のメニュー項目の実行
						await contextMenu.InvokeItem(menuId);
						break;
				}
			}
			catch (ObjectDisposedException)
			{
				// contextMenuが既に破棄されている場合は無視
				System.Diagnostics.Debug.WriteLine("InvokeShellMenuItemAsync: ContextMenu was disposed");
			}
			catch (OperationCanceledException)
			{
				// 操作がキャンセルされた場合は無視（contextMenuが破棄された場合など）
				System.Diagnostics.Debug.WriteLine("InvokeShellMenuItemAsync: Operation was canceled (contextMenu may have been disposed)");
			}
			catch (Exception ex) when (ex is COMException or UnauthorizedAccessException)
			{
				// COM例外やアクセス権限例外は無視
				System.Diagnostics.Debug.WriteLine($"InvokeShellMenuItemAsync: COM/Access exception: {ex}");
			}
			catch (Exception ex)
			{
				// その他の予期しない例外も無視してアプリケーションをクラッシュさせない
				System.Diagnostics.Debug.WriteLine($"InvokeShellMenuItemAsync: Error invoking menu item: {ex}");
			}
		}
	}
}


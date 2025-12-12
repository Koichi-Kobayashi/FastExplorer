// Copyright (c) Files Community
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
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
		
		// メニュー項目実行中のタスクを追跡するためのクラス
		private class MenuInvocationTracker
		{
			private readonly List<Task> _pendingTasks = new List<Task>();
			private readonly object _lock = new object();
			private bool _isDisposed = false;

			public void AddTask(Task task)
			{
				if (task == null)
					return;

				lock (_lock)
				{
					if (_isDisposed)
						return;
					_pendingTasks.Add(task);
					
					// タスクが完了したらリストから削除
					task.ContinueWith(t =>
					{
						lock (_lock)
						{
							_pendingTasks.Remove(t);
						}
					}, TaskContinuationOptions.ExecuteSynchronously);
				}
			}

			public async Task WaitForCompletionAsync(TimeSpan timeout)
			{
				Task[] tasksToWait;
				lock (_lock)
				{
					if (_pendingTasks.Count == 0)
						return;
					tasksToWait = _pendingTasks.ToArray();
					_isDisposed = true;
				}

				try
				{
					// すべてのタスクが完了するか、タイムアウトするまで待つ
					// Task.WhenAllは、タスクがキャンセルされた場合にOperationCanceledExceptionをスローする可能性があるため、
					// 個々のタスクを待機して、キャンセルや例外を適切に処理する
					var delayTask = Task.Delay(timeout);
					
					// すべてのタスクを待機（キャンセルされたタスクも含む）
					foreach (var task in tasksToWait)
					{
						var completedTask = await Task.WhenAny(task, delayTask);
						if (completedTask == delayTask)
						{
							// タイムアウト
							System.Diagnostics.Debug.WriteLine($"MenuInvocationTracker: Timeout waiting for task completion");
							break;
						}
						
						// タスクが完了した場合、例外やキャンセルを無視して継続
						try
						{
							await task;
						}
						catch (OperationCanceledException)
						{
							// タスクがキャンセルされた場合は無視（メニューが閉じられたため）
							System.Diagnostics.Debug.WriteLine($"MenuInvocationTracker: Task was canceled");
						}
						catch (ObjectDisposedException)
						{
							// オブジェクトが破棄された場合は無視
							System.Diagnostics.Debug.WriteLine($"MenuInvocationTracker: Object was disposed");
						}
						catch (Exception ex)
						{
							// その他の例外も無視（ログは既に出力されているはず）
							System.Diagnostics.Debug.WriteLine($"MenuInvocationTracker: Task exception: {ex.GetType().Name} - {ex.Message}");
						}
					}
				}
				catch (OperationCanceledException)
				{
					// タイムアウトまたはキャンセルされた場合は無視
					System.Diagnostics.Debug.WriteLine($"MenuInvocationTracker: WaitForCompletionAsync was canceled");
				}
				catch (ObjectDisposedException)
				{
					// オブジェクトが破棄された場合は無視
					System.Diagnostics.Debug.WriteLine($"MenuInvocationTracker: WaitForCompletionAsync - Object was disposed");
				}
				catch (Exception ex)
				{
					// その他の例外も無視
					System.Diagnostics.Debug.WriteLine($"MenuInvocationTracker: WaitForCompletionAsync exception: {ex.GetType().Name} - {ex.Message}");
				}
			}
		}

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

				// メニュー項目実行の追跡用オブジェクトを作成
				var invocationTracker = new MenuInvocationTracker();
				
				// contextMenuとinvocationTrackerをTagに保存
				wpfContextMenu.Tag = new Tuple<FilesContextMenu, MenuInvocationTracker>(contextMenu, invocationTracker);
				contextMenu = null; // 所有権をwpfContextMenuに移す

				var filesContextMenu = ((Tuple<FilesContextMenu, MenuInvocationTracker>)wpfContextMenu.Tag).Item1;
				BuildMenuItems(wpfContextMenu.Items, filesContextMenu, filesContextMenu.Items, foregroundBrush, invocationTracker);

				// メニューが閉じられたときにcontextMenuを破棄し、フラグをリセット
				wpfContextMenu.Closed += async (s, e) =>
				{
					lock (_lockObject)
					{
						_isShowingMenu = false;
					}

					if (s is ContextMenu menu && menu.Tag is Tuple<FilesContextMenu, MenuInvocationTracker> tag)
					{
						var cm = tag.Item1;
						var tracker = tag.Item2;

						try
						{
							// 実行中のタスクが完了するまで待つ（最大2秒に短縮してUIの応答性を改善）
							// ConfigureAwait(false)を使用して、UIスレッドをブロックしない
							await tracker.WaitForCompletionAsync(TimeSpan.FromSeconds(2)).ConfigureAwait(false);
						}
						catch (Exception ex)
						{
							System.Diagnostics.Debug.WriteLine($"FilesContextMenuWrapper: Error waiting for task completion: {ex}");
						}

						try
						{
							// タスク完了後に破棄
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
							}, System.Windows.Threading.DispatcherPriority.Background);
						}
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
				// 画面の下端を超えないように、画面の高さを考慮して位置を調整
				var screenHeight = System.Windows.SystemParameters.PrimaryScreenHeight;
				var screenWidth = System.Windows.SystemParameters.PrimaryScreenWidth;
				
				// メニューの推定サイズ
				const double estimatedMenuHeight = 400; // メニューの推定高さ
				const double estimatedMenuWidth = 250; // メニューの推定幅
				
				// 位置を調整
				double adjustedX = x;
				double adjustedY = y;
				
				// メニューが画面の下端を超える場合は、上方向に表示する
				if (y + estimatedMenuHeight > screenHeight)
				{
					adjustedY = Math.Max(0, y - estimatedMenuHeight);
				}
				
				// メニューが画面の右端を超える場合は、左方向に表示する
				if (x + estimatedMenuWidth > screenWidth)
				{
					adjustedX = Math.Max(0, x - estimatedMenuWidth);
				}
				
				wpfContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.AbsolutePoint;
				wpfContextMenu.PlacementRectangle = new Rect(adjustedX, adjustedY, 0, 0);
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

		private static void BuildMenuItems(ItemCollection items, FilesContextMenu contextMenu, List<Win32ContextMenuItem> menuItems, System.Windows.Media.Brush? foregroundBrush = null, MenuInvocationTracker? invocationTracker = null)
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
					var headerPanel = new StackPanel
					{
						Orientation = Orientation.Horizontal,
						Background = System.Windows.Media.Brushes.Transparent // ヒットテストを改善するために透明背景を設定
					};

					// Add spacer first to maintain alignment (icon will be loaded asynchronously)
					var iconContainer = new System.Windows.Controls.Border
					{
						Width = 16,
						Height = 16,
						Margin = new Thickness(8, 0, 12, 0),
						Background = System.Windows.Media.Brushes.Transparent
					};
					headerPanel.Children.Add(iconContainer);

					// Load icon asynchronously after menu is displayed (lazy loading)
					if (menuItem.Icon != null && menuItem.Icon.Length > 0)
					{
						var iconData = menuItem.Icon; // Capture for closure
						_ = Task.Run(() =>
						{
							try
							{
								using var ms = new MemoryStream(iconData);
								var bitmap = new BitmapImage();
								bitmap.BeginInit();
								bitmap.StreamSource = ms;
								bitmap.CacheOption = BitmapCacheOption.OnLoad;
								bitmap.EndInit();
								bitmap.Freeze();

								// UIスレッドでアイコンを設定
								Application.Current.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new Action(() =>
								{
									try
									{
										// スペーサーのMarginを取得してアイコンに適用
										var iconMargin = iconContainer.Margin;
										var icon = new System.Windows.Controls.Image
										{
											Source = bitmap,
											Width = 16,
											Height = 16,
											Margin = iconMargin, // スペーサーと同じMarginを設定
											VerticalAlignment = VerticalAlignment.Center
										};

										// スペーサーをアイコンに置き換え
										var index = headerPanel.Children.IndexOf(iconContainer);
										if (index >= 0)
										{
											headerPanel.Children.RemoveAt(index);
											headerPanel.Children.Insert(index, icon);
										}
									}
									catch (Exception ex)
									{
										System.Diagnostics.Debug.WriteLine($"Error setting icon on UI thread: {ex}");
									}
								}));
							}
							catch (Exception ex)
							{
								System.Diagnostics.Debug.WriteLine($"Error loading icon asynchronously: {ex}");
							}
						});
					}

					headerPanel.Children.Add(new TextBlock
					{
						Text = labelText,
						VerticalAlignment = VerticalAlignment.Center,
						FontSize = 14,
						Foreground = foregroundBrush
					});

					var wpfMenuItem = new MenuItem
					{
						Header = headerPanel,
						Padding = new Thickness(8, 6, 8, 6),
						MinHeight = 32,
						Foreground = foregroundBrush, // ダークモード対応のためForegroundを設定
						Tag = contextMenu // contextMenuをTagに保存して、後でアクセスできるようにする
					};

					// Handle submenu
					if (menuItem.SubItems != null && menuItem.SubItems.Count > 0)
					{
						// Load submenu asynchronously (fire and forget)
						_ = LoadSubMenuAsync(wpfMenuItem, contextMenu, menuItem.SubItems, foregroundBrush, invocationTracker);
					}
					else
					{
						// Set command to invoke menu item (Filesと同じ動作)
						var menuItemCopy = menuItem; // Capture for closure
						wpfMenuItem.Click += async (s, e) =>
						{
							try
							{
								e.Handled = true;

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

								// メニュー項目の実行
								if (s is MenuItem menuItem && menuItem.Tag is FilesContextMenu cm)
								{
									// 非同期処理をバックグラウンドで実行（UI ブロッキングを回避）
									// Fire-and-Forgetパターンで実行し、UIスレッドに影響を与えない
									var invokeTask = Task.Run(async () =>
									{
										try
										{
											await InvokeShellMenuItemAsync(cm, menuItemCopy).ConfigureAwait(false);
										}
										catch
										{
											// 例外はInvokeShellMenuItemAsync内で処理されるため、ここでは無視
										}
									});
									
									// 実行中のタスクを追跡
									if (invocationTracker != null)
									{
										invocationTracker.AddTask(invokeTask);
									}
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

		private static async Task LoadSubMenuAsync(MenuItem parentMenuItem, FilesContextMenu contextMenu, List<Win32ContextMenuItem> subItems, System.Windows.Media.Brush? foregroundBrush = null, MenuInvocationTracker? invocationTracker = null)
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
					var subHeaderPanel = new StackPanel
					{
						Orientation = Orientation.Horizontal,
						Background = System.Windows.Media.Brushes.Transparent // ヒットテストを改善するために透明背景を設定
					};

					// Add spacer first to maintain alignment (icon will be loaded asynchronously)
					var subIconContainer = new System.Windows.Controls.Border
					{
						Width = 16,
						Height = 16,
						Margin = new Thickness(8, 0, 12, 0),
						Background = System.Windows.Media.Brushes.Transparent
					};
					subHeaderPanel.Children.Add(subIconContainer);

					// Load icon asynchronously after menu is displayed (lazy loading)
					if (subItem.Icon != null && subItem.Icon.Length > 0)
					{
						var subIconData = subItem.Icon; // Capture for closure
						_ = Task.Run(() =>
						{
							try
							{
								using var ms = new MemoryStream(subIconData);
								var bitmap = new BitmapImage();
								bitmap.BeginInit();
								bitmap.StreamSource = ms;
								bitmap.CacheOption = BitmapCacheOption.OnLoad;
								bitmap.EndInit();
								bitmap.Freeze();

								// UIスレッドでアイコンを設定
								Application.Current.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new Action(() =>
								{
									try
									{
										// スペーサーのMarginを取得してアイコンに適用
										var subIconMargin = subIconContainer.Margin;
										var icon = new System.Windows.Controls.Image
										{
											Source = bitmap,
											Width = 16,
											Height = 16,
											Margin = subIconMargin, // スペーサーと同じMarginを設定
											VerticalAlignment = VerticalAlignment.Center
										};

										// スペーサーをアイコンに置き換え
										var index = subHeaderPanel.Children.IndexOf(subIconContainer);
										if (index >= 0)
										{
											subHeaderPanel.Children.RemoveAt(index);
											subHeaderPanel.Children.Insert(index, icon);
										}
									}
									catch (Exception ex)
									{
										System.Diagnostics.Debug.WriteLine($"Error setting submenu icon on UI thread: {ex}");
									}
								}));
							}
							catch (Exception ex)
							{
								System.Diagnostics.Debug.WriteLine($"Error loading submenu icon asynchronously: {ex}");
							}
						});
					}

					subHeaderPanel.Children.Add(new TextBlock
					{
						Text = subLabelText,
						VerticalAlignment = VerticalAlignment.Center,
						FontSize = 14,
						Foreground = foregroundBrush
					});

					var wpfSubMenuItem = new MenuItem
					{
						Header = subHeaderPanel,
						Padding = new Thickness(8, 6, 8, 6),
						MinHeight = 32,
						Foreground = foregroundBrush, // ダークモード対応のためForegroundを設定
						Tag = contextMenu // contextMenuをTagに保存して、後でアクセスできるようにする
					};

						// Handle nested submenu
						if (subItem.SubItems != null && subItem.SubItems.Count > 0)
						{
							// Load nested submenu asynchronously (fire and forget)
							_ = LoadSubMenuAsync(wpfSubMenuItem, contextMenu, subItem.SubItems, foregroundBrush, invocationTracker);
						}
						else
						{
							// Set command to invoke menu item (Filesと同じ動作)
							var subItemCopy = subItem; // Capture for closure
							wpfSubMenuItem.Click += async (s, e) =>
							{
								try
								{
									e.Handled = true;

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

									// メニュー項目の実行
									if (s is MenuItem menuItem && menuItem.Tag is FilesContextMenu cm)
									{
										// 非同期処理をバックグラウンドで実行（UI ブロッキングを回避）
										// Fire-and-Forgetパターンで実行し、UIスレッドに影響を与えない
										var invokeTask = Task.Run(async () =>
										{
											try
											{
												await InvokeShellMenuItemAsync(cm, subItemCopy).ConfigureAwait(false);
											}
											catch
											{
												// 例外はInvokeShellMenuItemAsync内で処理されるため、ここでは無視
											}
										});
										
										// 実行中のタスクを追跡
										if (invocationTracker != null)
										{
											invocationTracker.AddTask(invokeTask);
										}
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
			{
				System.Diagnostics.Debug.WriteLine("InvokeShellMenuItemAsync: menuItem or contextMenu is null");
				return;
			}

			try
			{
				var menuId = menuItem.ID;
				if (menuId < 0)
				{
					System.Diagnostics.Debug.WriteLine($"InvokeShellMenuItemAsync: Invalid menu ID: {menuId}");
					return;
				}

				var verb = menuItem.CommandString;
				var firstPath = contextMenu.ItemsPath?.FirstOrDefault();

				// タイムアウト付きで実行（最大30秒）
				var invokeTask = verb switch
				{
					"install" => contextMenu.InvokeItem(menuId),
					"installAllUsers" => contextMenu.InvokeItem(menuId),
					"mount" => contextMenu.InvokeItem(menuId),
					"format" => contextMenu.InvokeItem(menuId),
					"Windows.PowerShell.Run" => contextMenu.InvokeItem(
						menuId,
						!string.IsNullOrEmpty(firstPath) && firstPath.EndsWith(".ps1", StringComparison.OrdinalIgnoreCase)
							? System.IO.Path.GetDirectoryName(firstPath)
							: null),
					_ => contextMenu.InvokeItem(menuId)
				};

				try
				{
					// タイムアウト付きで待機（最大30秒）
					// ConfigureAwait(false)を使用して、同期コンテキストに戻らないようにする（UIブロッキングを回避）
					var timeoutTask = Task.Delay(TimeSpan.FromSeconds(30));
					var completedTask = await Task.WhenAny(invokeTask, timeoutTask).ConfigureAwait(false);

					if (completedTask == timeoutTask)
					{
						System.Diagnostics.Debug.WriteLine($"InvokeShellMenuItemAsync: Operation timed out after 30 seconds (verb: {verb})");
						return;
					}

					// 結果を取得（例外が発生している可能性があるため、awaitして例外をキャッチ）
					// invokeTaskが既に完了している場合、awaitは即座に完了する
					// ConfigureAwait(false)を使用して、同期コンテキストに戻らないようにする
					await invokeTask.ConfigureAwait(false);
				}
				catch (OperationCanceledException)
				{
					// 操作がキャンセルされた場合は無視（contextMenuが破棄された場合など）
					System.Diagnostics.Debug.WriteLine($"InvokeShellMenuItemAsync: Operation was canceled (verb: {verb})");
					return;
				}
			}
			catch (ObjectDisposedException ex)
			{
				// contextMenuが既に破棄されている場合は無視
				System.Diagnostics.Debug.WriteLine($"InvokeShellMenuItemAsync: ContextMenu was disposed: {ex.Message}");
			}
			catch (OperationCanceledException ex)
			{
				// 操作がキャンセルされた場合は無視（contextMenuが破棄された場合など）
				System.Diagnostics.Debug.WriteLine($"InvokeShellMenuItemAsync: Operation was canceled: {ex.Message}");
			}
			catch (Exception ex) when (ex is COMException or UnauthorizedAccessException)
			{
				// COM例外やアクセス権限例外は無視
				System.Diagnostics.Debug.WriteLine($"InvokeShellMenuItemAsync: COM/Access exception: {ex.GetType().Name} - {ex.Message}");
			}
			catch (Exception ex)
			{
				// その他の予期しない例外も無視してアプリケーションをクラッシュさせない
				System.Diagnostics.Debug.WriteLine($"InvokeShellMenuItemAsync: Error invoking menu item ({menuItem.Label}): {ex.GetType().Name} - {ex.Message}");
				System.Diagnostics.Debug.WriteLine($"InvokeShellMenuItemAsync: Stack trace: {ex.StackTrace}");
			}
		}
	}
}


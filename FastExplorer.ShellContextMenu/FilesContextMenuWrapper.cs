// Copyright (c) Files Community
// Licensed under the MIT License.

using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using System.IO;
using Wpf.Ui.Controls;
using MenuItem = Wpf.Ui.Controls.MenuItem;
using TextBlock = System.Windows.Controls.TextBlock;
using System.Windows.Media;

namespace FastExplorer.ShellContextMenu
{
	/// <summary>
	/// Wrapper for FilesContextMenu to display context menu in WPF applications
	/// </summary>
	public class FilesContextMenuWrapper
	{
		/// <summary>
		/// Shows the context menu for the specified file paths at the specified screen coordinates
		/// </summary>
		/// <param name="filePaths">Array of file paths</param>
		/// <param name="ownerWindow">Owner window</param>
		/// <param name="x">X coordinate in screen coordinates</param>
		/// <param name="y">Y coordinate in screen coordinates</param>
		public static void ShowContextMenu(string[] filePaths, Window ownerWindow, int x, int y)
		{
			if (filePaths == null || filePaths.Length == 0)
				return;

			if (ownerWindow == null)
				return;

			// UIスレッドで非同期処理を実行
			_ = ShowContextMenuAsync(filePaths, ownerWindow, x, y);
		}

		private static async Task ShowContextMenuAsync(string[] filePaths, Window ownerWindow, int x, int y)
		{
			try
			{
				System.Diagnostics.Debug.WriteLine($"FilesContextMenuWrapper: Getting context menu for: {string.Join(", ", filePaths)}");

				// Get context menu from Files implementation
				using var contextMenu = await FilesContextMenu.GetContextMenuForFiles(filePaths, 0x00000000); // CMF_NORMAL

				if (contextMenu == null)
				{
					System.Diagnostics.Debug.WriteLine("FilesContextMenuWrapper: Context menu is null");
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

				BuildMenuItems(wpfContextMenu.Items, contextMenu, contextMenu.Items, foregroundBrush);

				// Show WPF context menu
				wpfContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.AbsolutePoint;
				wpfContextMenu.PlacementRectangle = new Rect(x, y, 0, 0);
				wpfContextMenu.IsOpen = true;
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.WriteLine($"FilesContextMenuWrapper: Error showing context menu: {ex}");
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
						MinHeight = 32
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
						// Load submenu asynchronously
						_ = LoadSubMenuAsync(wpfMenuItem, contextMenu, menuItem.SubItems, foregroundBrush);
					}
					else
					{
						// Set command to invoke menu item (Filesと同じ動作)
						var menuItemCopy = menuItem; // Capture for closure
						wpfMenuItem.Click += async (s, e) =>
						{
							await InvokeShellMenuItemAsync(contextMenu, menuItemCopy);
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
				await contextMenu.LoadSubMenu(subItems);

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
							MinHeight = 32
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
							await LoadSubMenuAsync(wpfSubMenuItem, contextMenu, subItem.SubItems, foregroundBrush);
						}
						else
						{
							// Set command to invoke menu item (Filesと同じ動作)
							var subItemCopy = subItem; // Capture for closure
							wpfSubMenuItem.Click += async (s, e) =>
							{
								await InvokeShellMenuItemAsync(contextMenu, subItemCopy);
							};
						}

						parentMenuItem.Items.Add(wpfSubMenuItem);
					}
				}
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.WriteLine($"Error loading submenu: {ex}");
			}
		}

		/// <summary>
		/// Filesアプリと同じ動作でシェルメニュー項目を実行します
		/// </summary>
		private static async Task InvokeShellMenuItemAsync(FilesContextMenu contextMenu, Win32ContextMenuItem menuItem)
		{
			if (menuItem == null)
				return;

			var menuId = menuItem.ID;
			if (menuId < 0)
				return;

			var verb = menuItem.CommandString;
			var firstPath = contextMenu.ItemsPath.FirstOrDefault();

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
	}
}


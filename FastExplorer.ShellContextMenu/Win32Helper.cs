// Copyright (c) Files Community
// Licensed under the MIT License.

using System.Drawing;
using Vanara.PInvoke;

namespace FastExplorer.ShellContextMenu
{
	/// <summary>
	/// Win32 API helper methods
	/// </summary>
	internal static class Win32Helper
	{
		/// <summary>
		/// Gets all desktop windows
		/// </summary>
		public static List<HWND> GetDesktopWindows()
		{
			var windows = new List<HWND>();
			User32.EnumWindows((hWnd, lParam) =>
			{
				windows.Add(hWnd);
				return true;
			}, default);
			return windows;
		}

		/// <summary>
		/// Brings windows to foreground
		/// </summary>
		public static void BringToForeground(List<HWND> windows)
		{
			foreach (var window in windows)
			{
				if (User32.IsWindow(window))
				{
					User32.SetForegroundWindow(window);
				}
			}
		}

		/// <summary>
		/// Converts HBITMAP to Bitmap
		/// </summary>
		public static Bitmap? GetBitmapFromHBitmap(HBITMAP hBitmap)
		{
			if (hBitmap == HBITMAP.NULL)
				return null;

			try
			{
				// Use Image.FromHbitmap for simple conversion
				var bitmap = Image.FromHbitmap((IntPtr)(nint)hBitmap);
				return bitmap;
			}
			catch
			{
				return null;
			}
		}
	}
}


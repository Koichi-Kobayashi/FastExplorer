// Copyright (c) Files Community
// Licensed under the MIT License.

namespace FastExplorer.ShellContextMenu
{
	/// <summary>
	/// Represents an abstracted implementation for IDisposable
	/// </summary>
	public abstract class Disposable : IDisposable
	{
		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		protected virtual void Dispose(bool disposing)
		{
		}
	}
}


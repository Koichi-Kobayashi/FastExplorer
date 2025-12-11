// Copyright (c) Files Community
// Licensed under the MIT License.

using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;

namespace FastExplorer.ShellContextMenu
{
	public sealed partial class ThreadWithMessageQueue : Disposable
	{
		private readonly BlockingCollection<Internal> messageQueue;

		private readonly Thread thread;

		private readonly CancellationTokenSource cts = new();

		private bool disposed;

		protected override void Dispose(bool disposing)
		{
			if (disposing)
			{
				if (disposed)
					return;
				disposed = true;

				// Stop producers and request the worker to stop promptly.
				try { messageQueue.CompleteAdding(); } catch { }
				try { cts.Cancel(); } catch { }

				// Wait a bounded time for the worker to finish.
				if (!thread.Join(TimeSpan.FromSeconds(5)))
				{
					// Worker didn't stop in time; avoid disposing the collection while it may still be enumerating.
					return;
				}

				try { messageQueue.Dispose(); } catch { }
				try { cts.Dispose(); } catch { }
			}
		}

		public async Task<V> PostMethod<V>(Func<object> payload)
		{
			if (disposed || messageQueue.IsAddingCompleted)
			{
				var tcs = new TaskCompletionSource<object>(TaskCreationOptions.RunContinuationsAsynchronously);
				tcs.SetException(new InvalidOperationException("Queue closed"));
				return (V)await tcs.Task;
			}

			var message = new Internal(payload);
			if (!messageQueue.TryAdd(message))
			{
				message.tcs.SetException(new InvalidOperationException("Queue closed"));
			}

			return (V)await message.tcs.Task;
		}

		public Task PostMethod(Action payload)
		{
			if (disposed || messageQueue.IsAddingCompleted)
			{
				var tcs = new TaskCompletionSource<object>(TaskCreationOptions.RunContinuationsAsynchronously);
				tcs.SetException(new InvalidOperationException("Queue closed"));
				return tcs.Task;
			}

			var message = new Internal(payload);
			if (!messageQueue.TryAdd(message))
			{
				message.tcs.SetException(new InvalidOperationException("Queue closed"));
			}

			return message.tcs.Task;
		}

		public ThreadWithMessageQueue()
		{
			messageQueue = new BlockingCollection<Internal>(new ConcurrentQueue<Internal>());

			thread = new Thread(new ThreadStart(() =>
			{
				try
				{
					// Use cancellation-enabled enumerator so we can reliably unblock the waiting consumer.
					foreach (var message in messageQueue.GetConsumingEnumerable(cts.Token))
					{
						try
						{
							var res = message.payload();
							message.tcs.SetResult(res);
						}
						catch (Exception ex)
						{
							message.tcs.SetException(ex);
						}
					}
				}
				catch (OperationCanceledException)
				{
					// Normal shutdown due to cts.Cancel()
				}
				catch (ObjectDisposedException)
				{
					// Normal shutdown if collection was disposed unexpectedly
				}
				catch (Exception ex)
				{
					// Best effort: drain remaining items and propagate exception to callers.
					try
					{
						while (messageQueue.TryTake(out var m))
						{
							try { m.tcs.SetException(ex); } catch { }
						}
					}
					catch { }
				}
			}));

			thread.SetApartmentState(ApartmentState.STA);

			// Do not prevent app from closing
			thread.IsBackground = true;

			thread.Start();
		}

		private sealed class Internal
		{
			public Func<object?> payload;

			public TaskCompletionSource<object> tcs;

			public Internal(Action payload)
			{
				this.payload = () => { payload(); return default; };
				tcs = new TaskCompletionSource<object>(TaskCreationOptions.RunContinuationsAsynchronously);
			}

			public Internal(Func<object?> payload)
			{
				this.payload = payload;
				tcs = new TaskCompletionSource<object>(TaskCreationOptions.RunContinuationsAsynchronously);
			}
		}
	}
}


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

                try { messageQueue.CompleteAdding(); } catch { }

                if (!thread.Join(TimeSpan.FromSeconds(5)))
                {
                    return;
                }

                try { messageQueue.Dispose(); } catch { }
                try { cts.Dispose(); } catch { }
            }
        }

        public async Task<V> PostMethod<V>(Func<object?> payload)
        {
            if (disposed || messageQueue.IsAddingCompleted)
            {
                return default!;
            }

            var message = new Internal(payload);
            
            try
            {
                if (!messageQueue.TryAdd(message, TimeSpan.FromSeconds(1)))
                {
                    return default!;
                }
            }
            catch (OperationCanceledException)
            {
                // TryAddがキャンセルされた場合は無視
                return default!;
            }

            try
            {
                var task = message.tcs.Task;
                var completedTask = await Task.WhenAny(task, Task.Delay(TimeSpan.FromMinutes(1)));
                
                if (completedTask != task)
                {
                    return default!;
                }

                var result = await task;
                return (V?)(result ?? default(V)!) ?? default!;
            }
            catch (OperationCanceledException)
            {
                // 操作がキャンセルされた場合は無視
                return default!;
            }
            catch
            {
                return default!;
            }
        }

        public Task PostMethod(Action payload)
        {
            if (disposed || messageQueue.IsAddingCompleted)
            {
                return Task.CompletedTask;
            }

            var message = new Internal(payload);
            
            try
            {
                if (!messageQueue.TryAdd(message, TimeSpan.FromSeconds(1)))
                {
                    return Task.CompletedTask;
                }
            }
            catch (OperationCanceledException)
            {
                // TryAddがキャンセルされた場合は無視
                return Task.CompletedTask;
            }

            return message.tcs.Task.ContinueWith(
                _ => { },
                TaskScheduler.Default
            );
        }

        public ThreadWithMessageQueue()
        {
            messageQueue = new BlockingCollection<Internal>(new ConcurrentQueue<Internal>());

            thread = new Thread(new ThreadStart(() =>
            {
                // 非同期処理を開始して、STAスレッド内で実行されるようにする
                // SynchronizationContextを使用して、非同期処理がSTAスレッド内で実行されることを保証
                var syncContext = new System.Threading.SynchronizationContext();
                System.Threading.SynchronizationContext.SetSynchronizationContext(syncContext);
                
                // 非同期処理を同期的に待機（STAスレッド内で実行されるため、COMオブジェクトの操作が可能）
                var task = ProcessMessagesAsync();
                try
                {
                    task.GetAwaiter().GetResult();
                }
                catch
                {
                    // 例外はProcessMessagesAsync内で処理されるため、ここでは無視
                }
            }));

            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
        }

        private async Task ProcessMessagesAsync()
        {
            try
            {
                while (true)
                {
                    // TryTakeを非同期で実行（UIスレッドをブロックしない）
                    // Task.Runを使用して、TryTakeのブロッキング呼び出しを別スレッドで実行
                    Internal? message = null;
                    
                    try
                    {
                        // 短い間隔でTryTakeを試行し、その間にawaitすることで非同期にする
                        message = await Task.Run(() =>
                        {
                            try
                            {
                                if (messageQueue.TryTake(out var msg, TimeSpan.FromMilliseconds(100)))
                                    return msg;
                                return null;
                            }
                            catch (OperationCanceledException)
                            {
                                return null;
                            }
                        });
                    }
                    catch (OperationCanceledException)
                    {
                        // 操作がキャンセルされた場合
                        if (messageQueue.IsCompleted)
                            break;
                        continue;
                    }

                    // メッセージが取得できなかった場合、少し待ってから再試行
                    if (message == null)
                    {
                        if (messageQueue.IsCompleted)
                            break;
                        
                        // 非同期で待機することで、UIスレッドをブロックしない
                        await Task.Delay(50);
                        continue;
                    }

                    // メッセージの処理はSTAスレッド内で実行（COMオブジェクトの操作のために必要）
                    // Task.Yield()で継続をスケジュールし、STAスレッドに戻す
                    await Task.Yield();
                    
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
            catch (ObjectDisposedException)
            {
                // オブジェクトが破棄された場合は無視
            }
            catch (OperationCanceledException)
            {
                // 操作がキャンセルされた場合は無視
            }
            catch (Exception ex)
            {
                // 残りのメッセージを処理
                await ProcessRemainingMessagesAsync(ex);
            }
        }

        private async Task ProcessRemainingMessagesAsync(Exception exception)
        {
            try
            {
                while (true)
                {
                    Internal? message = null;
                    
                    try
                    {
                        // TryTakeを非同期で実行
                        message = await Task.Run(() =>
                        {
                            try
                            {
                                if (messageQueue.TryTake(out var msg, TimeSpan.FromMilliseconds(100)))
                                    return msg;
                                return null;
                            }
                            catch (OperationCanceledException)
                            {
                                return null;
                            }
                            catch
                            {
                                return null;
                            }
                        });
                    }
                    catch (OperationCanceledException)
                    {
                        // TryTakeがキャンセルされた場合は終了
                        break;
                    }
                    catch
                    {
                        // その他の例外は無視して終了
                        break;
                    }

                    if (message == null)
                        break;

                    try
                    {
                        message.tcs.SetException(exception);
                    }
                    catch
                    {
                        // 例外設定に失敗した場合は無視
                    }
                }
            }
            catch
            {
                // 残りのメッセージ処理中の例外は無視
            }
        }

        private sealed class Internal
        {
            public Func<object?> payload;
            public TaskCompletionSource<object?> tcs;

            public Internal(Action payload)
            {
                this.payload = () => { payload(); return default; };
                tcs = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
            }

            public Internal(Func<object?> payload)
            {
                this.payload = payload;
                tcs = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
            }
        }
    }
}

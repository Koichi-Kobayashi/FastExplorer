# WPF-UIのソースコード調査メモ

## 調査すべきポイント

### 1. FluentWindowクラスの実装
- **リポジトリ**: https://github.com/lepoco/wpfui
- **確認すべきファイル**: 
  - `src/WpfUi/Controls/FluentWindow.cs` または類似のファイル
  - `OnSourceInitialized`メソッドの実装
  - メッセージフック（`AddHook`）の使用箇所

### 2. ExtendsContentIntoTitleBarの処理
- `ExtendsContentIntoTitleBar="True"`の場合の非クライアント領域の処理
- `WindowChrome`の使用状況
- 非クライアント領域のメッセージ（`WM_NCLBUTTONDOWN`など）の処理方法

### 3. メッセージフックの処理順序
- WPF-UIがメッセージフックを追加している場合、その処理順序
- カスタムの`WndProc`実装があるかどうか

## 確認方法

1. GitHubリポジトリをクローンまたはブラウザで確認
2. `FluentWindow`クラスを検索
3. `OnSourceInitialized`や`AddHook`の使用箇所を確認
4. 非クライアント領域のメッセージ処理を確認

## 想定される問題

- WPF-UIが独自のメッセージフックを追加している場合、処理順序によって干渉する可能性がある
- `ExtendsContentIntoTitleBar="True"`の場合、WPF-UIが非クライアント領域をカスタマイズしている可能性がある
- メッセージフックの処理順序が原因で、閉じるボタンなどの処理が正常に動作しない可能性がある


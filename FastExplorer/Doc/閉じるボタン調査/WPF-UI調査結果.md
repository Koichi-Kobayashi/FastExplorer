# WPF-UI調査結果

## 調査した内容

### 1. FluentWindowクラスの実装
- **ファイル**: `wpfui/src/Wpf.Ui/Controls/FluentWindow/FluentWindow.cs`
- **発見事項**:
  - `OnSourceInitialized`では、独自のメッセージフックは追加していない
  - `base.OnSourceInitialized(e)`を呼び出しているだけ
  - `ExtendsContentIntoTitleBar="True"`の場合、`WindowChrome`を使用して非クライアント領域をカスタマイズ

### 2. WindowChromeの使用
- `ExtendsContentIntoTitleBar`が`True`の場合、以下の設定で`WindowChrome`を設定:
  ```csharp
  WindowChrome.SetWindowChrome(
      this,
      new WindowChrome
      {
          CaptionHeight = 0,
          CornerRadius = default,
          GlassFrameThickness = new Thickness(-1),
          ResizeBorderThickness = ResizeMode == ResizeMode.NoResize ? default : new Thickness(4),
          UseAeroCaptionButtons = false,
      }
  );
  ```

### 3. RemoveWindowTitlebarContents
- `UnsafeNativeMethods.RemoveWindowTitlebarContents`は、システムメニュー（タイトルバーのアイコンをクリックしたときに表示されるメニュー）を削除するだけ
- 閉じるボタン、最大化ボタン、最小化ボタンは削除していない

## 問題の原因

### 想定される原因
1. **WindowChromeとメッセージフックの干渉**
   - `WindowChrome`はWPFの標準機能で、非クライアント領域のメッセージを処理する
   - メッセージフックが`WindowChrome`の処理と干渉している可能性がある

2. **メッセージフックの処理順序**
   - メッセージフックは、`WindowChrome`の処理の前または後に呼ばれる可能性がある
   - 処理順序によって、閉じるボタンなどの処理が正常に動作しない可能性がある

## 解決策の検討

### 1. メッセージフックの処理を最小限にする
- 非クライアント領域のメッセージは、メッセージフックで処理せず、`WindowChrome`に委譲する
- メニュー関連のメッセージのみを処理する

### 2. TrackPopupMenuExが返った時点でリセット
- メッセージフックでリセットするのではなく、`TrackPopupMenuEx`が返った時点で確実にリセットする
- これにより、`WindowChrome`の処理と干渉しない

### 3. WindowChromeの処理を確認
- `WindowChrome`がどのように非クライアント領域のメッセージを処理しているかを確認
- 必要に応じて、`WindowChrome`の設定を調整する

## 結論

WPF-UI自体はメッセージフックを追加していないが、`WindowChrome`を使用して非クライアント領域をカスタマイズしている。問題の原因は、メッセージフックと`WindowChrome`の処理の干渉である可能性が高い。

解決策として、メッセージフックで非クライアント領域のメッセージを処理せず、`TrackPopupMenuEx`が返った時点で確実にリセットする方法が有効である。


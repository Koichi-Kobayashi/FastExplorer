# WindowChrome調査結果

## WindowChromeの設定

### WPF-UIのFluentWindowでの設定
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

### 重要な設定
- **CaptionHeight = 0**: タイトルバーの高さを0に設定（コンテンツをタイトルバー領域に拡張）
- **UseAeroCaptionButtons = false**: Aeroのキャプションボタンを使用しない（カスタムボタンを使用）

## WindowChromeの動作

### メッセージ処理
- `WindowChrome`はWPFの標準機能で、非クライアント領域のメッセージを処理する
- `WM_NCHITTEST`、`WM_NCLBUTTONDOWN`、`WM_SYSCOMMAND`などのメッセージを処理
- メッセージフックの前に処理される可能性がある

### ヒットテスト
- 非クライアント領域内の要素をクリック可能にするには、`WindowChrome.IsHitTestVisibleInChrome="True"`を設定する必要がある
- これにより、非クライアント領域内の要素がクリック可能になる

## 問題の原因

### 想定される原因
1. **メッセージ処理の順序**
   - `WindowChrome`が非クライアント領域のメッセージを処理する
   - メッセージフックが`WindowChrome`の処理の前または後に呼ばれる
   - 処理順序によって、閉じるボタンなどの処理が正常に動作しない可能性がある

2. **メッセージフックとの干渉**
   - メッセージフックで非クライアント領域のメッセージを処理しようとすると、`WindowChrome`の処理と干渉する可能性がある
   - `handled = false`にしても、メッセージフックが呼ばれることで、`WindowChrome`の処理に影響を与える可能性がある

## 解決策の検討

### 1. メッセージフックで非クライアント領域のメッセージを処理しない
- 非クライアント領域のメッセージは、`WindowChrome`に完全に委譲する
- メッセージフックでは、メニュー関連のメッセージのみを処理する

### 2. TrackPopupMenuExが返った時点でリセット
- メッセージフックでリセットするのではなく、`TrackPopupMenuEx`が返った時点で確実にリセットする
- これにより、`WindowChrome`の処理と干渉しない

### 3. WindowChromeの設定を確認
- `UseAeroCaptionButtons = false`の場合、カスタムのキャプションボタンを使用している
- これらのボタンが正しく動作するように、`WindowChrome.IsHitTestVisibleInChrome="True"`を設定する必要がある

## 結論

`WindowChrome`は非クライアント領域のメッセージを処理するため、メッセージフックと干渉する可能性がある。解決策として、メッセージフックで非クライアント領域のメッセージを処理せず、`TrackPopupMenuEx`が返った時点で確実にリセットする方法が有効である。


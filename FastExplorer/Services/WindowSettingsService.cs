using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Windows;

namespace FastExplorer.Services
{
    /// <summary>
    /// ウィンドウ設定の管理を行うサービス
    /// </summary>
    public class WindowSettingsService
    {
        #region フィールド

        private readonly string _settingsFilePath;
        private WindowSettings _settings = new();

        #endregion

        #region コンストラクタ

        /// <summary>
        /// <see cref="WindowSettingsService"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        public WindowSettingsService()
        {
            var appDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "FastExplorer");
            Directory.CreateDirectory(appDataPath);
            _settingsFilePath = Path.Combine(appDataPath, "window_settings.json");
            LoadSettings();
        }

        #endregion

        #region 設定取得・保存

        /// <summary>
        /// ウィンドウ設定を取得します
        /// </summary>
        /// <returns>ウィンドウ設定</returns>
        public WindowSettings GetSettings()
        {
            return _settings;
        }

        /// <summary>
        /// ウィンドウ設定を保存します
        /// </summary>
        /// <param name="settings">保存するウィンドウ設定</param>
        public void SaveSettings(WindowSettings settings)
        {
            _settings = settings;
            SaveSettings();
        }

        #endregion

        #region ファイル操作

        /// <summary>
        /// ウィンドウ設定を保存します
        /// </summary>
        private void SaveSettings()
        {
            try
            {
                var json = JsonSerializer.Serialize(_settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_settingsFilePath, json);
            }
            catch
            {
                // エラーハンドリング
            }
        }

        /// <summary>
        /// ウィンドウ設定を読み込みます
        /// </summary>
        private void LoadSettings()
        {
            try
            {
                // 起動時の高速化：File.Exists()の呼び出しを削減（直接ReadAllTextを試みる）
                var json = File.ReadAllText(_settingsFilePath);
                _settings = JsonSerializer.Deserialize<WindowSettings>(json) ?? new WindowSettings();
            }
            catch
            {
                // ファイルが存在しない、または読み込みに失敗した場合はデフォルト設定を使用
                _settings = new WindowSettings();
            }
        }

        #endregion
    }

    /// <summary>
    /// ウィンドウ設定を表すクラス
    /// </summary>
    public class WindowSettings
    {
        #region プロパティ
        /// <summary>
        /// ウィンドウの幅を取得または設定します
        /// </summary>
        public double Width { get; set; } = 1100;

        /// <summary>
        /// ウィンドウの高さを取得または設定します
        /// </summary>
        public double Height { get; set; } = 650;

        /// <summary>
        /// ウィンドウの左位置を取得または設定します
        /// </summary>
        public double Left { get; set; } = double.NaN;

        /// <summary>
        /// ウィンドウの上位置を取得または設定します
        /// </summary>
        public double Top { get; set; } = double.NaN;

        /// <summary>
        /// ウィンドウの状態（最大化など）を取得または設定します
        /// </summary>
        public WindowState State { get; set; } = WindowState.Normal;

        /// <summary>
        /// アプリケーションのテーマを取得または設定します（"System"（規定）、"Light"、または"Dark"）
        /// </summary>
        public string Theme { get; set; } = "System";

        /// <summary>
        /// 選択されたテーマカラーの名前を取得または設定します
        /// </summary>
        public string? ThemeColorName { get; set; }

        /// <summary>
        /// 選択されたテーマカラーのカラーコードを取得または設定します
        /// </summary>
        public string? ThemeColorCode { get; set; }

        /// <summary>
        /// 選択されたテーマカラーのセカンダリカラーコードを取得または設定します
        /// </summary>
        public string? ThemeSecondaryColorCode { get; set; }

        /// <summary>
        /// 選択されたテーマカラーのサードカラーコードを取得または設定します（分割ペイン用の背景色）
        /// </summary>
        public string? ThemeThirdColorCode { get; set; }

        /// <summary>
        /// 保存されたタブのパスのリストを取得または設定します
        /// </summary>
        public List<string> TabPaths { get; set; } = new();

        /// <summary>
        /// 分割ペインが有効かどうかを取得または設定します
        /// </summary>
        public bool IsSplitPaneEnabled { get; set; } = false;

        /// <summary>
        /// 左ペインのタブのパスのリストを取得または設定します
        /// </summary>
        public List<string> LeftPaneTabPaths { get; set; } = new();

        /// <summary>
        /// 右ペインのタブのパスのリストを取得または設定します
        /// </summary>
        public List<string> RightPaneTabPaths { get; set; } = new();

        /// <summary>
        /// 背景画像のファイルパスを取得または設定します
        /// </summary>
        public string? BackgroundImagePath { get; set; }

        /// <summary>
        /// 背景画像の不透明度を取得または設定します（0.0～1.0）
        /// </summary>
        public double BackgroundImageOpacity { get; set; } = 1.0;

        /// <summary>
        /// 背景画像の調整方法を取得または設定します
        /// </summary>
        public string BackgroundImageStretch { get; set; } = "FitToWindow";

        /// <summary>
        /// 背景画像の垂直方向の配置を取得または設定します
        /// </summary>
        public string BackgroundImageVerticalAlignment { get; set; } = "Center";

        /// <summary>
        /// 背景画像の水平方向の配置を取得または設定します
        /// </summary>
        public string BackgroundImageHorizontalAlignment { get; set; } = "Center";

        /// <summary>
        /// 選択された言語コードを取得または設定します（"ja", "en"など）
        /// </summary>
        public string LanguageCode { get; set; } = "ja";

        #endregion
    }
}


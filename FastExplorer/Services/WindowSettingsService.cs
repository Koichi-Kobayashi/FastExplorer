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
        private readonly string _settingsFilePath;
        private WindowSettings _settings = new();

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
                if (File.Exists(_settingsFilePath))
                {
                    var json = File.ReadAllText(_settingsFilePath);
                    _settings = JsonSerializer.Deserialize<WindowSettings>(json) ?? new WindowSettings();
                }
                else
                {
                    _settings = new WindowSettings();
                }
            }
            catch
            {
                _settings = new WindowSettings();
            }
        }
    }

    /// <summary>
    /// ウィンドウ設定を表すクラス
    /// </summary>
    public class WindowSettings
    {
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
        /// 保存されたタブのパスのリストを取得または設定します
        /// </summary>
        public List<string> TabPaths { get; set; } = new();
    }
}


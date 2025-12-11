using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.Json;
using System.Windows;

namespace FastExplorer.Helpers
{
    /// <summary>
    /// 多言語化を管理するヘルパークラス
    /// </summary>
    public static class LocalizationHelper
    {
        private static Dictionary<string, string>? _currentLanguage;
        private static string _currentLanguageCode = "ja";

        /// <summary>
        /// 言語が変更されたときに発火するイベント
        /// </summary>
        public static event EventHandler<string>? LanguageChanged;

        /// <summary>
        /// 現在の言語コードを取得または設定します
        /// </summary>
        public static string CurrentLanguageCode
        {
            get => _currentLanguageCode;
            set
            {
                if (_currentLanguageCode != value)
                {
                    _currentLanguageCode = value;
                    LoadLanguage(value);
                    // 言語変更イベントを発火
                    LanguageChanged?.Invoke(null, value);
                }
            }
        }

        /// <summary>
        /// 言語ファイルを読み込みます
        /// </summary>
        /// <param name="languageCode">言語コード（例: "ja", "en"）</param>
        public static void LoadLanguage(string languageCode)
        {
            try
            {
                _currentLanguageCode = languageCode;
                var assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
                var assemblyDirectory = Path.GetDirectoryName(assemblyLocation);
                if (string.IsNullOrEmpty(assemblyDirectory))
                {
                    // フォールバック: 実行可能ファイルのディレクトリを使用
                    assemblyDirectory = AppDomain.CurrentDomain.BaseDirectory;
                }

                var langFilePath = Path.Combine(assemblyDirectory, "lang", $"{languageCode}.json");
                
                // ファイルが存在しない場合は、アセンブリのリソースから読み込む
                if (!File.Exists(langFilePath))
                {
                    // リソースから読み込む試み
                    var resourceStream = System.Reflection.Assembly.GetExecutingAssembly()
                        .GetManifestResourceStream($"FastExplorer.lang.{languageCode}.json");
                    
                    if (resourceStream != null)
                    {
                        using var reader = new StreamReader(resourceStream);
                        var json = reader.ReadToEnd();
                        _currentLanguage = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                        return;
                    }
                }
                else
                {
                    var json = File.ReadAllText(langFilePath);
                    _currentLanguage = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"言語ファイルの読み込みエラー: {ex.Message}");
                // エラーが発生した場合は、デフォルトの日本語を使用
                _currentLanguage = null;
            }
        }

        /// <summary>
        /// 指定されたキーの翻訳文字列を取得します
        /// </summary>
        /// <param name="key">翻訳キー</param>
        /// <param name="defaultValue">デフォルト値（キーが見つからない場合）</param>
        /// <returns>翻訳された文字列</returns>
        public static string GetString(string key, string? defaultValue = null)
        {
            try
            {
                if (_currentLanguage == null)
                {
                    LoadLanguage(_currentLanguageCode);
                }

                if (_currentLanguage != null && _currentLanguage.TryGetValue(key, out var value))
                {
                    return value;
                }

                return defaultValue ?? key;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LocalizationHelper.GetString エラー: {ex.Message}");
                return defaultValue ?? key;
            }
        }

        /// <summary>
        /// 指定されたキーの翻訳文字列を取得します（フォーマット文字列対応）
        /// </summary>
        /// <param name="key">翻訳キー</param>
        /// <param name="args">フォーマット引数</param>
        /// <returns>翻訳された文字列</returns>
        public static string GetString(string key, params object[] args)
        {
            var format = GetString(key);
            try
            {
                return string.Format(format, args);
            }
            catch
            {
                return format;
            }
        }

        /// <summary>
        /// システムの言語コードを取得して言語を初期化します
        /// </summary>
        public static void InitializeFromSystemLanguage()
        {
            var culture = CultureInfo.CurrentUICulture;
            var languageCode = culture.TwoLetterISOLanguageName.ToLowerInvariant();
            
            // サポートされている言語のみを読み込む
            if (languageCode == "ja" || languageCode == "en")
            {
                CurrentLanguageCode = languageCode;
            }
            else
            {
                // デフォルトは英語
                CurrentLanguageCode = "en";
            }
        }

        /// <summary>
        /// 初期化（アプリケーション起動時に呼び出す）
        /// 静的コンストラクタでは初期化しない（設定から読み込むため）
        /// </summary>
        static LocalizationHelper()
        {
            // 静的コンストラクタでは初期化しない
            // App.xaml.csのOnStartupで設定から読み込む
        }
    }
}


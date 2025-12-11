using System;
using System.Windows.Markup;

namespace FastExplorer.Helpers
{
    /// <summary>
    /// XAMLでLocalizationHelperを使用するためのMarkupExtension
    /// </summary>
    public class LocalizationHelperExtension : MarkupExtension
    {
        /// <summary>
        /// 翻訳キー
        /// </summary>
        public string Key { get; set; } = string.Empty;

        /// <summary>
        /// デフォルト値（キーが見つからない場合）
        /// </summary>
        public string? DefaultValue { get; set; }

        /// <summary>
        /// <see cref="LocalizationHelperExtension"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        public LocalizationHelperExtension()
        {
        }

        /// <summary>
        /// <see cref="LocalizationHelperExtension"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="key">翻訳キー</param>
        public LocalizationHelperExtension(string key)
        {
            Key = key;
        }

        /// <summary>
        /// 翻訳された文字列を返します
        /// </summary>
        /// <param name="serviceProvider">サービスプロバイダー</param>
        /// <returns>翻訳された文字列</returns>
        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            try
            {
                if (string.IsNullOrEmpty(Key))
                {
                    return DefaultValue ?? string.Empty;
                }

                return LocalizationHelper.GetString(Key, DefaultValue ?? Key);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LocalizationHelperExtension.ProvideValue エラー: {ex.Message}");
                return DefaultValue ?? Key ?? string.Empty;
            }
        }
    }
}


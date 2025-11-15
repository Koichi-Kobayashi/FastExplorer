namespace FastExplorer.Models
{
    /// <summary>
    /// アプリケーション設定を表すクラス
    /// </summary>
    public class AppConfig
    {
        /// <summary>
        /// 設定フォルダーのパスを取得または設定します
        /// </summary>
        public string ConfigurationsFolder { get; set; }

        /// <summary>
        /// アプリケーションプロパティファイル名を取得または設定します
        /// </summary>
        public string AppPropertiesFileName { get; set; }
    }
}

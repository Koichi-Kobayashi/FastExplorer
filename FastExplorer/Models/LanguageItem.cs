using System.ComponentModel;

namespace FastExplorer.Models
{
    /// <summary>
    /// 言語選択用のアイテムを表すクラス
    /// </summary>
    public class LanguageItem : INotifyPropertyChanged
    {
        private string _code = string.Empty;
        private string _name = string.Empty;

        /// <summary>
        /// 言語コード（例: "ja", "en"）
        /// </summary>
        public string Code
        {
            get => _code;
            set
            {
                if (_code != value)
                {
                    _code = value;
                    OnPropertyChanged(nameof(Code));
                }
            }
        }

        /// <summary>
        /// 言語名称（例: "日本語", "English"）
        /// </summary>
        public string Name
        {
            get => _name;
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnPropertyChanged(nameof(Name));
                }
            }
        }

        /// <summary>
        /// プロパティ変更イベント
        /// </summary>
        public event PropertyChangedEventHandler? PropertyChanged;

        /// <summary>
        /// プロパティ変更通知を発火します
        /// </summary>
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// <see cref="LanguageItem"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        public LanguageItem()
        {
        }

        /// <summary>
        /// <see cref="LanguageItem"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="code">言語コード</param>
        /// <param name="name">言語名称</param>
        public LanguageItem(string code, string name)
        {
            Code = code;
            Name = name;
        }
    }
}



namespace ExcelExporter.Csv
{
    /// <summary>
    /// アプリケーション実行に使用する設定を扱うクラス。設定は外部ファイルから読み込まれます。
    /// </summary>
    internal sealed record AppSettings
    {
        /// <summary>
        /// 読み取りを行う外部ファイルの名前。
        /// </summary>
        public const string FileName = "appsettings.json";

        public string FileExtension { get; init; } = ".csv";
    }
}

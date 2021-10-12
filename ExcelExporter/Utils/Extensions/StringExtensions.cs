using System;

namespace ExcelExporter.Extensions
{
    internal static class StringExtensions
    {
        private static readonly string[] newLineCodes = { "\r\n", "\r", "\n" };

        /// <summary>
        /// <para>改行コード ("\r\n", "\r", "\n") に基づいて文字列を部分文字列に分割します。
        /// 部分文字列が空の配列の要素を含めるかどうかを指定することができます。</para>
        /// <para>文字列が <see langword="null"/> の場合は空の配列を返します。</para>
        /// </summary>
        /// <param name="str">文字列。</param>
        /// <param name="options">返される配列から空の配列要素を省略する場合は <see cref="StringSplitOptions.RemoveEmptyEntries"/>。返される配列に空の配列要素も含める場合は <see cref="StringSplitOptions.None"/>。</param>
        /// <returns>文字列を改行コード ("\r\n", "\r", "\n") で区切ることによって取り出された部分文字列を格納する配列。</returns>
        public static string[] ToLines(this string? str, StringSplitOptions options = StringSplitOptions.None)
        {
            return (str == null) ? Array.Empty<string>() : (str.Split(newLineCodes, options) ?? Array.Empty<string>());
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;

namespace Sheaside.Services
{
    internal static class WorksheetNameRules
    {
        public static string Validate(string name, IEnumerable<string> otherSheetNames)
        {
            if (string.IsNullOrWhiteSpace(name)) return "シート名を入力してください。";
            if (name.Length > 31) return "シート名は31文字以内にしてください。";
            if (name.IndexOfAny(new[] { ':', '\\', '/', '?', '*', '[', ']' }) >= 0 || name.Any(char.IsControl))
                return "シート名に : \\ / ? * [ ] や制御文字は使用できません。";
            if (name[0] == '\'' || name[name.Length - 1] == '\'')
                return "シート名の先頭・末尾にアポストロフィ（'）は使用できません。";
            if (string.Equals(name, "History", StringComparison.OrdinalIgnoreCase))
                return "History はExcelの予約名のため使用できません。";
            if (otherSheetNames.Any(existing => string.Equals(existing, name, StringComparison.OrdinalIgnoreCase)))
                return "同じ名前のシートが既にあります。別の名前を入力してください。";
            return null;
        }
    }
}

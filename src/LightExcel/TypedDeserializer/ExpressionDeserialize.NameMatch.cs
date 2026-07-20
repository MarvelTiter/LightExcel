using LightExcel.Attributes;
using System.Reflection;

namespace LightExcel.TypedDeserializer;

partial class ExpressionDeserialize<T>
{
    private static bool MemberMatchesName(MemberInfo member, string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return false;

        string attributeValue = GetColumnNameAttribute(member);

        return MatchesWithNormalization(name, attributeValue, member.Name);

        static string GetColumnNameAttribute(MemberInfo member)
        {
            if (member.GetCustomAttributes(typeof(ExcelColumnAttribute), true).Length > 0)
            {
                return ((ExcelColumnAttribute)member.GetCustomAttributes(typeof(ExcelColumnAttribute), true)[0]).Name ?? string.Empty;
            }
            else if (member.IsDefined(typeof(ExcelColumnAttribute), true))
            {
                return member.GetCustomAttribute<ExcelColumnAttribute>(true)?.Name ?? string.Empty;
            }
            else
            {
                return string.Empty;
            }
        }

        static bool MatchesWithNormalization(string headerName, string columnAttrName, string memberName)
        {
            headerName = headerName.Trim();
            if (string.IsNullOrWhiteSpace(headerName)) return false;
            var matchNames = new List<string>();
            if (!string.IsNullOrWhiteSpace(columnAttrName)) matchNames.Add(columnAttrName);
            if (!string.IsNullOrWhiteSpace(memberName)) matchNames.Add(memberName);
            if (matchNames.Count == 0) return false;
            // 一致性匹配
            foreach (var matchName in matchNames)
            {
                if (string.Equals(headerName, matchName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            // 去除多余字符匹配
            var normalizedHeader = NormalizeString(headerName);
            foreach (var matchName in matchNames)
            {
                var normalizedMatch = NormalizeString(matchName);
                if (string.Equals(normalizedHeader, normalizedMatch, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }

        static string NormalizeString(string str)
        {
            str = StringNormalizer.XmlEscapeRegex().Replace(str, "");
            str = StringNormalizer.ControlCharsRegex().Replace(str, "");
            return str;
        }
    }
}

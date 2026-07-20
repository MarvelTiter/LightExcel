using System.Text.RegularExpressions;

namespace LightExcel.TypedDeserializer;

internal static partial class StringNormalizer
{
#if NET8_0_OR_GREATER
    [GeneratedRegex("_x[0-9A-Fa-f]{4}_")]
    public static partial Regex XmlEscapeRegex();

    [GeneratedRegex("[\r\n\t]+")]
    public static partial Regex ControlCharsRegex();
#else
    private static Regex? xmlEscapeRegex;
    public static Regex XmlEscapeRegex() => xmlEscapeRegex ??= new Regex(
        @"_x[0-9A-Fa-f]{4}_",
        RegexOptions.Compiled | RegexOptions.IgnoreCase
    );

    private static Regex? controlCharsRegex;
    public static Regex ControlCharsRegex() => controlCharsRegex ??= new Regex(
        @"[\r\n\t]+",
        RegexOptions.Compiled | RegexOptions.IgnoreCase
    );

#endif
}

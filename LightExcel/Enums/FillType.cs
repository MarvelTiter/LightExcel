
using LightExcel.Attributes;
using System.Reflection;

namespace LightExcel.Enums
{
    public enum FillType
    {
        /// <summary>
        /// None.
        /// <para>When the item is serialized out as xml, its value is "none".</para>
        /// </summary>
        None,
        /// <summary>
        /// Solid.
        /// <para>When the item is serialized out as xml, its value is "solid".</para>
        /// </summary>
        Solid,
        /// <summary>
        /// Medium Gray.
        /// <para>When the item is serialized out as xml, its value is "mediumGray".</para>
        /// </summary>
        MediumGray,
        /// <summary>
        /// Dary Gray.
        /// <para>When the item is serialized out as xml, its value is "darkGray".</para>
        /// </summary>
        DarkGray,
        /// <summary>
        /// Light Gray.
        /// <para>When the item is serialized out as xml, its value is "lightGray".</para>
        /// </summary>
        LightGray,
        /// <summary>
        /// Dark Horizontal.
        /// <para>When the item is serialized out as xml, its value is "darkHorizontal".</para>
        /// </summary>
        DarkHorizontal,
        /// <summary>
        /// Dark Vertical.
        /// <para>When the item is serialized out as xml, its value is "darkVertical".</para>
        /// </summary>
        DarkVertical,
        /// <summary>
        /// Dark Down.
        /// <para>When the item is serialized out as xml, its value is "darkDown".</para>
        /// </summary>
        DarkDown,
        /// <summary>
        /// Dark Up.
        /// <para>When the item is serialized out as xml, its value is "darkUp".</para>
        /// </summary>
        DarkUp,
        /// <summary>
        /// Dark Grid.
        /// <para>When the item is serialized out as xml, its value is "darkGrid".</para>
        /// </summary>
        DarkGrid,
        /// <summary>
        /// Dark Trellis.
        /// <para>When the item is serialized out as xml, its value is "darkTrellis".</para>
        /// </summary>
        DarkTrellis,
        /// <summary>
        /// Light Horizontal.
        /// <para>When the item is serialized out as xml, its value is "lightHorizontal".</para>
        /// </summary>
        LightHorizontal,
        /// <summary>
        /// Light Vertical.
        /// <para>When the item is serialized out as xml, its value is "lightVertical".</para>
        /// </summary>
        LightVertical,
        /// <summary>
        /// Light Down.
        /// <para>When the item is serialized out as xml, its value is "lightDown".</para>
        /// </summary>
        LightDown,
        /// <summary>
        /// Light Up.
        /// <para>When the item is serialized out as xml, its value is "lightUp".</para>
        /// </summary>
        LightUp,
        /// <summary>
        /// Light Grid.
        /// <para>When the item is serialized out as xml, its value is "lightGrid".</para>
        /// </summary>
        LightGrid,
        /// <summary>
        /// Light Trellis.
        /// <para>When the item is serialized out as xml, its value is "lightTrellis".</para>
        /// </summary>
        LightTrellis,
        /// <summary>
        /// Gray 0.125.
        /// <para>When the item is serialized out as xml, its value is "gray125".</para>
        /// </summary>
        Gray125,
        /// <summary>
        /// Gray 0.0625.
        /// <para>When the item is serialized out as xml, its value is "gray0625".</para>
        /// </summary>
        Gray0625
    }
    
    public enum NumberFormat
    {
        None = 0,
        [FormatCode("0")]
        Integer = 1,
        [FormatCode("0.00")]
        Float2 = 2,
        [FormatCode("0%")]
        Percent0 = 9,
        [FormatCode("0.00%")]
        Percent2 = 10,
        [FormatCode("0.00E + 00")]
        Scientific = 11,
    }

    public static class Ex
    {
        public static string GetFormatCode(this NumberFormat format)
        {
           var code = typeof(NumberFormat).GetField(format.ToString())!.GetCustomAttribute<FormatCodeAttribute>();
            return code!.Code;
        }
    }
}

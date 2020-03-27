using System;
using System.ComponentModel;

namespace Codex.Word.Net.Base
{
    /// <summary>
    /// Get Enum type's Description
    /// </summary>
    /// <code>
    /// string _description = BaseStyle.Normal.GetDescription()
    /// </code>
    public static class StringExtensions
    {
        public static string GetDescription(this Enum value)
        {
            if (value == null)
                return "";
 
            System.Reflection.FieldInfo fieldInfo = value.GetType().GetField(value.ToString());
 
            object[] attribArray = fieldInfo.GetCustomAttributes(false);
            if (attribArray.Length == 0)
                return value.ToString();
            else
                return (attribArray[0] as DescriptionAttribute)?.Description;
        }
    }
    
    /// <summary>
    /// Base style provide an easy way to create a Word document.
    /// </summary>
    public enum BaseStyle
    {
        [Description("Normal")]
        Normal = 0,
        [Description("Heading1")]
        Heading1,
        [Description("Heading2")]
        Heading2,
        [Description("Heading3")]
        Heading3,
        [Description("Heading4")]
        Heading4,
        [Description("Heading5")]
        Heading5,
        [Description("Heading6")]
        Heading6,
        [Description("Heading7")]
        Heading7,
        [Description("Heading8")]
        Heading8,
        [Description("Heading9")]
        Heading9,
        [Description("Heading2Char")]
        Heading2Char,
        [Description("Heading3Char")]
        Heading3Char,
        [Description("Heading4Char")]
        Heading4Char,
        [Description("Heading5Char")]
        Heading5Char,
        [Description("Heading6Char")]
        Heading6Char,
        [Description("Heading7Char")]
        Heading7Char,
        [Description("Heading8Char")]
        Heading8Char,
        [Description("Heading9Char")]
        Heading9Char,
        [Description("TableNormal")]
        TableNormal,
        [Description("TableGrid")]
        TableGrid,
        [Description("LightShading")]
        LightShading,
        [Description("LightShading-Accent1")]
        LightShadingAccent1,
        [Description("LightShading-Accent2")]
        LightShadingAccent2,
        [Description("LightShading-Accent3")]
        LightShadingAccent3,
        [Description("LightShading-Accent4")]
        LightShadingAccent4,
        [Description("LightShading-Accent5")]
        LightShadingAccent5,
        [Description("LightShading-Accent6")]
        LightShadingAccent6,
        [Description("LightList")]
        LightList,
        [Description("LightList-Accent1")]
        LightListAccent1,
        [Description("LightList-Accent2")]
        LightListAccent2,
        [Description("LightList-Accent3")]
        LightListAccent3,
        [Description("LightList-Accent4")]
        LightListAccent4,
        [Description("LightList-Accent5")]
        LightListAccent5,
        [Description("LightList-Accent6")]
        LightListAccent6,
        [Description("LightGrid")]
        LightGrid,
        [Description("LightGrid-Accent1")]
        LightGridAccent1,
        [Description("LightGrid-Accent2")]
        LightGridAccent2,
        [Description("LightGrid-Accent3")]
        LightGridAccent3,
        [Description("LightGrid-Accent4")]
        LightGridAccent4,
        [Description("LightGrid-Accent5")]
        LightGridAccent5,
        [Description("LightGrid-Accent6")]
        LightGridAccent6,
        [Description("MediumShading1")]
        MediumShading1,
        [Description("MediumShading1-Accent1")]
        MediumShading1Accent1,
        [Description("MediumShading1-Accent2")]
        MediumShading1Accent2,
        [Description("MediumShading1-Accent3")]
        MediumShading1Accent3,
        [Description("MediumShading1-Accent4")]
        MediumShading1Accent4,
        [Description("MediumShading1-Accent5")]
        MediumShading1Accent5,
        [Description("MediumShading1-Accent6")]
        MediumShading1Accent6,
        [Description("MediumShading2")]
        MediumShading2,
        [Description("MediumShading2-Accent1")]
        MediumShading2Accent1,
        [Description("MediumShading2-Accent2")]
        MediumShading2Accent2,
        [Description("MediumShading2-Accent3")]
        MediumShading2Accent3,
        [Description("MediumShading2-Accent4")]
        MediumShading2Accent4,
        [Description("MediumShading2-Accent5")]
        MediumShading2Accent5,
        [Description("MediumShading2-Accent6")]
        MediumShading2Accent6,
        [Description("MediumList1")]
        MediumList1,
        [Description("MediumList1-Accent1")]
        MediumList1Accent1,
        [Description("MediumList1-Accent2")]
        MediumList1Accent2,
        [Description("MediumList1-Accent3")]
        MediumList1Accent3,
        [Description("MediumList1-Accent4")]
        MediumList1Accent4,
        [Description("MediumList1-Accent5")]
        MediumList1Accent5,
        [Description("MediumList1-Accent6")]
        MediumList1Accent6,
        [Description("MediumList2")]
        MediumList2,
        [Description("MediumList2-Accent1")]
        MediumList2Accent1,
        [Description("MediumList2-Accent2")]
        MediumList2Accent2,
        [Description("MediumList2-Accent3")]
        MediumList2Accent3,
        [Description("MediumList2-Accent4")]
        MediumList2Accent4,
        [Description("MediumList2-Accent5")]
        MediumList2Accent5,
        [Description("MediumList2-Accent6")]
        MediumList2Accent6,
        [Description("MediumShading1")]
        MediumGrid1,
        [Description("MediumGrid1-Accent1")]
        MediumGrid1Accent1,
        [Description("MediumGrid1-Accent2")]
        MediumGrid1Accent2,
        [Description("MediumGrid1-Accent3")]
        MediumGrid1Accent3,
        [Description("MediumGrid1-Accent4")]
        MediumGrid1Accent4,
        [Description("MediumGrid1-Accent5")]
        MediumGrid1Accent5,
        [Description("MediumGrid1-Accent6")]
        MediumGrid1Accent6,
        [Description("MediumGrid2")]
        MediumGrid2,
        [Description("MediumGrid2-Accent1")]
        MediumGrid2Accent1,
        [Description("MediumGrid2-Accent2")]
        MediumGrid2Accent2,
        [Description("MediumGrid2-Accent3")]
        MediumGrid2Accent3,
        [Description("MediumGrid2-Accent4")]
        MediumGrid2Accent4,
        [Description("MediumGrid2-Accent5")]
        MediumGrid2Accent5,
        [Description("MediumGrid2-Accent6")]
        MediumGrid2Accent6,
        [Description("DarkList")]
        DarkList,
        [Description("DarkList-Accent1")]
        DarkListAccent1,
        [Description("DarkList-Accent2")]
        DarkListAccent2,
        [Description("DarkList-Accent3")]
        DarkListAccent3,
        [Description("DarkList-Accent4")]
        DarkListAccent4,
        [Description("DarkList-Accent5")]
        DarkListAccent5,
        [Description("DarkList-Accent6")]
        DarkListAccent6,
        [Description("ColorfulShading")]
        ColorfulShading,
        [Description("ColorfulShading-Accent1")]
        ColorfulShadingAccent1,
        [Description("ColorfulShading-Accent2")]
        ColorfulShadingAccent2,
        [Description("ColorfulShading-Accent3")]
        ColorfulShadingAccent3,
        [Description("ColorfulShading-Accent4")]
        ColorfulShadingAccent4,
        [Description("ColorfulShading-Accent5")]
        ColorfulShadingAccent5,
        [Description("ColorfulShading-Accent6")]
        ColorfulShadingAccent6,
        [Description("ColorfulList")]
        ColorfulList,
        [Description("ColorfulList-Accent1")]
        ColorfulListAccent1,
        [Description("ColorfulList-Accent2")]
        ColorfulListAccent2,
        [Description("ColorfulList-Accent3")]
        ColorfulListAccent3,
        [Description("ColorfulList-Accent4")]
        ColorfulListAccent4,
        [Description("ColorfulList-Accent5")]
        ColorfulListAccent5,
        [Description("ColorfulList-Accent6")]
        ColorfulListAccent6,
        [Description("ColorfulGrid")]
        ColorfulGrid,
        [Description("ColorfulGrid-Accent1")]
        ColorfulGridAccent1,
        [Description("ColorfulGrid-Accent2")]
        ColorfulGridAccent2,
        [Description("ColorfulGrid-Accent3")]
        ColorfulGridAccent3,
        [Description("ColorfulGrid-Accent4")]
        ColorfulGridAccent4,
        [Description("ColorfulGrid-Accent5")]
        ColorfulGridAccent5,
        [Description("ColorfulGrid-Accent6")]
        ColorfulGridAccent6
    }

    public enum TableCellBorderStyle
    {
        None = 0,
        SingleLine,
        Thick,
        DoubleLine,
        Dotted,
        Dashed,
        DotDash,
        DotDotDash,
        Triple,
        ThinThickSmallGap,
        ThickThinSmallGap,
        ThinThickThinSmallGap,
        ThinThickMediumGap,
        ThickThinMediumGap,
        ThinThickThinMediumGap,
        ThinThickLargeGap,
        ThickThinLargeGap,
        ThinThickThinLargeGap,
        Wave,
        DoubleWave,
        DashSmallGap,
        DashDotStroked,
        ThreeDEmboss,
        ThreeDEngrave,
        Outset,
        Inset,
        Nil
    }
}
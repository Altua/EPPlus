using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml
{
    public sealed class ExcelTheme : XmlHelper
    {
        private const string c_node_accent1 = "a:accent1";
        private const string c_node_accent2 = "a:accent2";
        private const string c_node_accent3 = "a:accent3";
        private const string c_node_accent4 = "a:accent4";
        private const string c_node_accent5 = "a:accent5";
        private const string c_node_accent6 = "a:accent6";
        private const string c_node_dk1 = "a:dk1";
        private const string c_node_dk2 = "a:dk2";
        private const string c_node_folHLink = "a:folHlink";
        private const string c_node_hlink = "a:hlink";
        private const string c_node_lt1 = "a:lt1";
        private const string c_node_lt2 = "a:lt2";
        private const string c_node_srgbClr = "a:srgbClr";
        private const string c_themeColorsPath = @"a:theme/a:themeElements/a:clrScheme";

        private readonly string[] _colors;
        private readonly XmlNamespaceManager _nameSpaceManager;
        private readonly XmlDocument _themeXml;

        internal ExcelTheme(XmlNamespaceManager NameSpaceManager, XmlDocument xml, ExcelWorkbook wb) :
            base(NameSpaceManager, xml)
        {
            _themeXml = xml;
            _nameSpaceManager = NameSpaceManager;
            SchemaNodeOrder = new string[] { "clrScheme" };
            _colors = LoadFromDocument();
        }

        public IEnumerable<string> Colors => _colors;

        private string GetValue(XmlElement xmlNode) => xmlNode.GetAttribute("val");

        private string[] LoadFromDocument()
        {
            var colors = new string[12];

            XmlNode colorsNode = _themeXml.SelectSingleNode(c_themeColorsPath, _nameSpaceManager);

            var dk1 = colorsNode[c_node_dk1][c_node_srgbClr];
            var lt1 = colorsNode[c_node_lt1][c_node_srgbClr];
            var dk2 = colorsNode[c_node_dk2][c_node_srgbClr];
            var lt2 = colorsNode[c_node_lt2][c_node_srgbClr];
            var accent1 = colorsNode[c_node_accent1][c_node_srgbClr];
            var accent2 = colorsNode[c_node_accent2][c_node_srgbClr];
            var accent3 = colorsNode[c_node_accent3][c_node_srgbClr];
            var accent4 = colorsNode[c_node_accent4][c_node_srgbClr];
            var accent5 = colorsNode[c_node_accent5][c_node_srgbClr];
            var accent6 = colorsNode[c_node_accent6][c_node_srgbClr];
            var hlink = colorsNode[c_node_hlink][c_node_srgbClr];
            var folHlink = colorsNode[c_node_folHLink][c_node_srgbClr];

            if (dk1 == null && lt1 == null)
            {
                colors[(int)ThemeColors.lt1] = "ffffff";
                colors[(int)ThemeColors.dk1] = "000000";
            }
            else
            {
                colors[(int)ThemeColors.lt1] = GetValue(lt1);
                colors[(int)ThemeColors.dk1] = GetValue(dk1);
            }

            colors[(int)ThemeColors.lt2] = GetValue(lt2);
            colors[(int)ThemeColors.dk2] = GetValue(dk2);

            colors[(int)ThemeColors.accent1] = GetValue(accent1);
            colors[(int)ThemeColors.accent2] = GetValue(accent2);
            colors[(int)ThemeColors.accent3] = GetValue(accent3);
            colors[(int)ThemeColors.accent4] = GetValue(accent4);
            colors[(int)ThemeColors.accent5] = GetValue(accent5);
            colors[(int)ThemeColors.accent6] = GetValue(accent6);

            colors[(int)ThemeColors.hlink] = GetValue(hlink);
            colors[(int)ThemeColors.folHlink] = GetValue(folHlink);

            return colors;
        }

        public string GetColor(ThemeColors color)
        {
            return _colors[(int)color];
        }
    }
}

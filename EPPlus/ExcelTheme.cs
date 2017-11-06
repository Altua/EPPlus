using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml
{
    public sealed class ExcelTheme : XmlHelper
    {
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

            var dk1 = colorsNode["a:dk1"]["a:srgbClr"];
            var lt1 = colorsNode["a:lt1"]["a:srgbClr"];
            var dk2 = colorsNode["a:dk2"]["a:srgbClr"];
            var lt2 = colorsNode["a:lt2"]["a:srgbClr"];
            var accent1 = colorsNode["a:accent1"]["a:srgbClr"];
            var accent2 = colorsNode["a:accent2"]["a:srgbClr"];
            var accent3 = colorsNode["a:accent3"]["a:srgbClr"];
            var accent4 = colorsNode["a:accent4"]["a:srgbClr"];
            var accent5 = colorsNode["a:accent5"]["a:srgbClr"];
            var accent6 = colorsNode["a:accent6"]["a:srgbClr"];
            var hlink = colorsNode["a:hlink"]["a:srgbClr"];
            var folHlink = colorsNode["a:folHlink"]["a:srgbClr"];

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

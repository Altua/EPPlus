using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml
{
    public sealed class ExcelTheme : XmlHelper
    {
        private const string c_themeColorsPath = @"a:theme/a:themeElements/a:clrScheme";
        private List<string> _colors;
        XmlNamespaceManager _nameSpaceManager;
        XmlDocument _themeXml;
        ExcelWorkbook _wb;

        internal ExcelTheme(XmlNamespaceManager NameSpaceManager, XmlDocument xml, ExcelWorkbook wb) :
            base(NameSpaceManager, xml)
        {
            _themeXml = xml;
            _wb = wb;
            _nameSpaceManager = NameSpaceManager;
            SchemaNodeOrder = new string[] { "clrScheme" };
            LoadFromDocument();
        }

        public IEnumerable<string> Colors => _colors;

        private string GetValue(XmlElement xmlNode)
        {
            return xmlNode.GetAttribute("val");
        }

        private void LoadFromDocument()
        {
            _colors = new List<string>();

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
                _colors.Add("ffffff");
                _colors.Add("000000");
            }
            else
            {
                _colors.Add(GetValue(lt1));
                _colors.Add(GetValue(dk1));
            }

            _colors.Add(GetValue(lt2));
            _colors.Add(GetValue(dk2));
            _colors.Add(GetValue(accent1));
            _colors.Add(GetValue(accent2));
            _colors.Add(GetValue(accent3));
            _colors.Add(GetValue(accent4));
            _colors.Add(GetValue(accent5));
            _colors.Add(GetValue(accent6));
            _colors.Add(GetValue(hlink));
            _colors.Add(GetValue(folHlink));
        }

        public string GetColor(ThemeColors color)
        {
            return _colors.ElementAt((int)color);
        }
    }
}

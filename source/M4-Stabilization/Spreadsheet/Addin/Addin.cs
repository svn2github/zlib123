using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using CleverAge.OdfConverter.OdfConverterLib;
using CleverAge.OdfConverter.Spreadsheet;

namespace CleverAge.OdfConverter.Spreadsheet
{
    public class Addin : OdfAddinLib
    {
        public Addin()
            : base(new CleverAge.OdfConverter.Spreadsheet.Converter())
        {
            this.OverrideResourceManager = new System.Resources.ResourceManager("ExcelAddin.resources.Labels", Assembly.GetExecutingAssembly());
        }

    }
}

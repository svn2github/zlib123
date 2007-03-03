using System;
using System.Collections.Generic;
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
        }

    }
}

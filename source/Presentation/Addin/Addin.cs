using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using CleverAge.OdfConverter.OdfConverterLib;

namespace Sonata.OdfConverter.Presentation
{
    public class Addin : OdfAddinLib
    {
        public Addin()
            : base(new Sonata.OdfConverter.Presentation.Converter())
        {
        }
    }
}

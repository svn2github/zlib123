using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using CleverAge.OdfConverter.OdfConverterLib;

using CleverAge.OdfConverter.Word;

namespace CleverAge.OdfConverter.Word
{
    public class Addin : OdfAddinLib
    {
        public Addin() : base(new CleverAge.OdfConverter.Word.Converter())
        {
            this.OverrideResourceManager = new System.Resources.ResourceManager("WordAddin.resources.Labels", Assembly.GetExecutingAssembly());
        }
    }
}

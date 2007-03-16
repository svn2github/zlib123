using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using CleverAge.OdfConverter.OdfConverterLib;
using Sonata.OdfConverter.Presentation;  

namespace Sonata.OdfConverter.Presentation
{
    public class Addin : OdfAddinLib
    {
        public Addin() : base(new Sonata.OdfConverter.Presentation.Converter())
        {
            this.OverrideResourceManager = new System.Resources.ResourceManager("PresentationAddin.resources.Labels", Assembly.GetExecutingAssembly());
        }
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace CleverAge.OdfConverter.OdfWordAddinLib
{
    [ComVisible(true), Guid("11FB3926-34EF-4f0f-8404-399C5285D583"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IOdfConverter
    {

        void OdfToOox(string inputFile, string outputFile, bool showUserInterface);
        void OoxToOdf(string inputFile, string outputFile, bool showUserInterface);
    }
}

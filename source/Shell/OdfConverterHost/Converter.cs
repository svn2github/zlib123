using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Threading;
using System.Globalization;
using CleverAge.OdfConverter.OdfConverterLib;
using System.Resources;
using System.Reflection;

namespace OdfConverterHost {
    [Guid("94714634-B065-4f06-839E-5C03CADCB934")]
    [ComVisible(true)]
    public class Converter {
        CleverAge.OdfConverter.Word.Addin _wordAddin;
#if EXCEL
        CleverAge.OdfConverter.Spreadsheet.Addin _excelAddin;
#endif
#if POWERPOINT
        Sonata.OdfConverter.Presentation.Addin _powerpointAddin;
#endif
        public Converter() {
            Tracer("Converter.Wrapper");
        }

        protected void Tracer(string msg) {
            Win32.OutputDebugString(msg + "\n");
            if (Program.MainForm != null) {
                Program.MainForm.AjouterMessage(msg);
            }
        }

        // Word conversion
        [DispId(1)]
        public void OdtToDocx(string inputFile, string outputFile, bool showUserInterface, int culture, int centerPos) {
            using (SetCulture old = new SetCulture(culture)) {
                Tracer("Converter.OdtToDocx : inputFile=" + inputFile);
                Tracer("Converter.OdtToDocx : outputFile=" + outputFile);
                Tracer("Converter.OdtToDocx : showUserInterface=" + showUserInterface.ToString());
                Tracer("Converter.OdtToDocx : culture=" + culture.ToString());
                using (HookManager hookManager = new HookManager(centerPos)) {
                    if (_wordAddin == null) {
                        _wordAddin = new CleverAge.OdfConverter.Word.Addin();
                    }
                    _wordAddin.OdfToOox(inputFile, outputFile, showUserInterface);
                }
                Tracer("Converter.OdtToDocx : done !");
            }
        }

        [DispId(2)]
        public void DocxToOdt(string inputFile, string outputFile, bool showUserInterface, int culture, int centerPos) {
            using (SetCulture old = new SetCulture(culture)) {
                Tracer("Converter.DocxToOdt : inputFile=" + inputFile);
                Tracer("Converter.DocxToOdt : outputFile=" + outputFile);
                Tracer("Converter.DocxToOdt : showUserInterface=" + showUserInterface.ToString());
                Tracer("Converter.DocxToOdt : culture=" + culture.ToString());
                using (HookManager hookManager = new HookManager(centerPos)) {
                    if (_wordAddin == null) {
                        _wordAddin = new CleverAge.OdfConverter.Word.Addin();
                    }
                    _wordAddin.OoxToOdf(inputFile, outputFile, showUserInterface);
                }
                Tracer("Converter.DocxToOdt : done !");
            }
        }

#if EXCEL
        // Excel conversion
        [DispId(3)]
        public void OdsToXlsx(string inputFile, string outputFile, bool showUserInterface, int culture, int centerPos) {
            using (SetCulture old = new SetCulture(culture)) {
                Tracer("Converter.OdsToXlsx : inputFile=" + inputFile);
                Tracer("Converter.OdsToXlsx : outputFile=" + outputFile);
                Tracer("Converter.OdsToXlsx : showUserInterface=" + showUserInterface.ToString());
                Tracer("Converter.OdsToXlsx : culture=" + culture.ToString());
                using (HookManager hookManager = new HookManager(centerPos)) {
                    if (_excelAddin == null) {
                        _excelAddin = new CleverAge.OdfConverter.Spreadsheet.Addin();
                    }
                    _excelAddin.OdfToOox(inputFile, outputFile, showUserInterface);
                }
            }
        }

        [DispId(4)]
        public void XlsxToOds(string inputFile, string outputFile, bool showUserInterface, int culture, int centerPos) {
            using (SetCulture old = new SetCulture(culture)) {
                Tracer("Converter.XlsxToOds : inputFile=" + inputFile);
                Tracer("Converter.XlsxToOds : outputFile=" + outputFile);
                Tracer("Converter.XlsxToOds : showUserInterface=" + showUserInterface.ToString());
                Tracer("Converter.XlsxToOds : culture=" + culture.ToString());
                using (HookManager hookManager = new HookManager(centerPos)) {
                    if (_excelAddin == null) {
                        _excelAddin = new CleverAge.OdfConverter.Spreadsheet.Addin();
                    }
                    _excelAddin.OoxToOdf(inputFile, outputFile, showUserInterface);
                }
            }
        }
#endif

#if POWERPOINT
        // Powerpoint conversion
        [DispId(5)]
        public void OdpToPptx(string inputFile, string outputFile, bool showUserInterface, int culture, int centerPos) {
            using (SetCulture old = new SetCulture(culture)) {
                Tracer("Converter.OdpToPptx : inputFile=" + inputFile);
                Tracer("Converter.OdpToPptx : outputFile=" + outputFile);
                Tracer("Converter.OdpToPptx : showUserInterface=" + showUserInterface.ToString());
                Tracer("Converter.OdpToPptx : culture=" + culture.ToString());
                using (HookManager hookManager = new HookManager(centerPos)) {
                    if (_powerpointAddin == null) {
                        _powerpointAddin = new Sonata.OdfConverter.Presentation.Addin();
                    }
                    _powerpointAddin.OdfToOox(inputFile, outputFile, showUserInterface);
                }
            }
        }

        [DispId(6)]
        public void PptxToOdp(string inputFile, string outputFile, bool showUserInterface, int culture, int centerPos) {
            using (SetCulture old = new SetCulture(culture)) {
                Tracer("Converter.PptxToOdp : inputFile=" + inputFile);
                Tracer("Converter.PptxToOdp : outputFile=" + outputFile);
                Tracer("Converter.PptxToOdp : showUserInterface=" + showUserInterface.ToString());
                Tracer("Converter.PptxToOdp : culture=" + culture.ToString());
                using (HookManager hookManager = new HookManager(centerPos)) {
                    if (_powerpointAddin == null) {
                        _powerpointAddin = new Sonata.OdfConverter.Presentation.Addin();
                    }
                    _powerpointAddin.OoxToOdf(inputFile, outputFile, showUserInterface);
                }
            }
        }
#endif

        [DispId(99)]
        public void Exit() {
            Tracer("Exit explicite");
            System.Windows.Forms.Application.Exit();
        }

    }
}

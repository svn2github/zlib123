using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Globalization;

namespace OdfConverterHost {
    internal class SetCulture : IDisposable {
        CultureInfo _old;

        public SetCulture(int culture) {
            _old = Thread.CurrentThread.CurrentUICulture;
            Thread.CurrentThread.CurrentUICulture = new CultureInfo(culture);
        }

        public void Dispose() {
            Thread.CurrentThread.CurrentUICulture = _old;
        }
    }
}

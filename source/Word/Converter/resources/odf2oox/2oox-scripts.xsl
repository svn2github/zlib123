<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt"
    xmlns:ooc="urn:odf-converter"
    exclude-result-prefixes="msxsl ooc">

  <msxsl:script language="C#" implements-prefix="ooc">
    <![CDATA[
      public string ToUpper(string input)
      {
          return input.ToUpper();
      }
      
      public string ToLower(string input)
      {
          return input.ToLower();
      }
      
      public string Trim(string input)
      {
          return input.Trim();
      }
      
      public string Replace(string input, string oldValue, string newValue)
      {
          return input.Replace(oldValue, newValue);
      }
      
      public bool IsUriAbsolute(string uri)
      {
          Uri tmp;
          return Uri.TryCreate(uri, UriKind.Absolute, out tmp) || uri.StartsWith("/") || uri.Contains(":");
      }

      public bool IsUriRelative(string uri)
      {
          Uri tmp;
          return Uri.TryCreate(uri, UriKind.Relative, out tmp) && !uri.StartsWith("/") && !uri.Contains(":");
      }

      public bool IsUriValid(string uri)
      {
          Uri tmp;
          return Uri.TryCreate(uri, UriKind.RelativeOrAbsolute, out tmp);
      }
      
      ///<summary>
      /// Convert various length units to twips (twentieths of a point)
      ///</summary>
      public int TwipsFromMeasuredUnit(string measuredUnit)
      {
          double value = 0;
          double factor = 1.0;

          Regex regex = new Regex(@"\s*([-.0-9]+)\s*([a-zA-Z]*)\s*");

          GroupCollection groups = regex.Match(measuredUnit).Groups;
          if (groups.Count == 3)
          {
              string strValue = groups[1].Value;
              string strUnit = groups[2].Value;

              switch (strUnit)
              {
                  case "cm":
                      factor = 1440.0 / 2.54;
                      break;
                  case "mm":
                      factor = 1440.0 / 25.4;
                      break;
                  case "in":
                  case "inch":
                      factor = 1440.0;
                      break;
                  case "pt":
                      factor = 20.0;
                      break;
                  case "twip":
                      factor = 1.0;
                      break;
                  case "pica":
                      factor = 240;
                      break;
                  case "dpt":
                      factor = 20.0;
                      break;
                  case "px":
                      factor = 1440.0 / 96.0;
                      break;
              }

              if (double.TryParse(strValue, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out value))
              {
                  value = value * factor;
              }
          }
          return (int)Math.Round(value, MidpointRounding.AwayFromZero);
      }
    ]]>

  </msxsl:script>

</xsl:stylesheet>

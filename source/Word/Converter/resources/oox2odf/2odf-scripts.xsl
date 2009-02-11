<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
    xmlns:ooc="urn:odf-converter"
    exclude-result-prefixes="msxsl ooc">

  <msxsl:script language="C#" implements-prefix="ooc">
    <msxsl:assembly name="WordprocessingConverter" />
    <msxsl:using namespace="OdfConverter.Wordprocessing" />
    
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
      
      
      /// <summary>
      /// Returns the current date and time as an XSD dateTime string in the format CCYY-MM-DDThh:mm:ss
      /// </summary>
      public string DateTimeNow()
      {
          return string.Format("{0:s}", System.DateTime.Now);
      }
      
      /// <summary>
      /// Formats a given date as an XSD dateTime string in the format CCYY-MM-DDThh:mm:ss
      /// If the input value is an empty node-set the current date and time are returned
      /// </summary>
      public string FormatDateTime(string dateTime)
      {
          DateTime result;

          if (!DateTime.TryParse(dateTime, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AdjustToUniversal, out result))
          {
              result = System.DateTime.Now;
          }
          return string.Format("{0:s}", result);
      }
      
      public string CmFromTwips(string twipsValue)
      {
          // concat(format-number($length * 2.54 div 1440,'#.###'),'cm')
          int twips = 0;
          if (int.TryParse(twipsValue, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out twips))
          {
              return string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0.###}cm", (double)twips * 2.54 / 1440.0);
          }

          return "0cm";
      }
      
      /// <summary>
      /// Removes all invalid characters from a style name and - if the name is not yet unique - 
      /// appends a counter to make the name unique
      /// </summary>
      public string NCNameFromString(string name)
      {
          return OdfStyleNameGenerator.Instance.NCNameFromString(name); 
      }
      
      public string UriFromPath(string path)
      {
          System.Uri result = null;

          if (!System.Uri.TryCreate(path, System.UriKind.RelativeOrAbsolute, out result))
          {
              return path;
          }
          return result.ToString();
      }
      
      /// <summary>
      /// Get a value from a semicolon-separated list of key-value pairs.
      /// The pairs are separated by colon as in CSS-like strings.
      /// </summary>
      public string ParseValueFromList(string input, string key)
      {
          string value = String.Empty;

          string[] keyValuePairs = input.Split(';');
          for (int i = keyValuePairs.Length - 1; i >= 0; i--)
          {
              if (keyValuePairs[i].StartsWith(key + ":"))
              {
                  string[] keyValuePair = keyValuePairs[i].Split(':');
                  value = keyValuePair[keyValuePair.Length - 1];
                  break;
              }
          }

          return value;
      }
    ]]>

  </msxsl:script>

</xsl:stylesheet>

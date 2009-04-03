<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
    xmlns:ooc="urn:odf-converter"
    exclude-result-prefixes="msxsl ooc">

  <msxsl:script language="C#" implements-prefix="ooc">
    <!--<msxsl:assembly name="WordprocessingConverter" />
    <msxsl:using namespace="System.Collections.Generic" />
    <msxsl:using namespace="OdfConverter.Wordprocessing" />-->
    
    <![CDATA[
      public double MaxDouble(double val1, double val2)
      {
          return Math.Max(val1, val2);
      }
    
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
      
      public string Trim(string input, string trimChars)
      {
          return input.Trim(trimChars.ToCharArray());
      }
      
      public string Replace(string input, string oldValue, string newValue)
      {
          return input.Replace(oldValue, newValue);
      }
      
      public string RegexReplace(string input, string pattern, string replacement, bool ignoreCase)
      {
          System.Text.RegularExpressions.RegexOptions options = System.Text.RegularExpressions.RegexOptions.Compiled | System.Text.RegularExpressions.RegexOptions.CultureInvariant;
          if (ignoreCase)
          {
              options |= System.Text.RegularExpressions.RegexOptions.IgnoreCase;
          }
          return System.Text.RegularExpressions.Regex.Replace(input, pattern, replacement, options);
      }
      
      public string RegexReplaceWithDefault(string input, string pattern, string replacement, bool ignoreCase, string defaultValue)
      {
          string result = RegexReplace(input, pattern, replacement, ignoreCase);
          
          if (string.IsNullOrEmpty(result))
          {
              result = defaultValue;
          }
          return result;
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
      
      /// <summary>
      /// Formats a given date as an XSD dateTime string in the format CCYY-MM-DDThh:mm:ss
      /// If the date string cannot be parsed an empty string is returned
      /// </summary>
      public string XsdDateTimeFromField(string fieldInstruction, string fieldValue)
        {
            string result = string.Empty;

            // parse the date format from the field code
            //    date-and-time-formatting-switch:
            //       \@ [ " ] switch-argument [ " ]
            string dateFormat = Regex.Replace(fieldInstruction, "\\\\@ *(?:\"(.*)\"|([^ ]*))", "$1$2", RegexOptions.IgnoreCase);
            
            // 'AM/PM' in Word needs to be replaced by 'tt' in .NET
            dateFormat = Regex.Replace(dateFormat, "(.*)AM/PM(.*)", "$1tt$2", RegexOptions.IgnoreCase);
            
            // remove other noise from the pattern
            dateFormat = dateFormat.Replace("'", "");

            DateTime dateTime = DateTime.Now;

            if (DateTime.TryParseExact(fieldValue, dateFormat, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dateTime))
            {
                result = string.Format("{0:s}", dateTime);
            }

            return result;
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
      
      public string CmFromEmu(string emuValue)
      {
          int emu = 0;
          if (int.TryParse(emuValue, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out emu))
          {
              return string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0.###}cm", (double)emu / 360000.0);
          }

          return "0cm";
      }
      
      public string CmFromMeasuredUnit(string measuredUnit)
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
                      factor = 1.0;
                      break;
                  case "mm":
                      factor = 0.1;
                      break;
                  case "in":
                  case "inch":
                      factor = 2.54;
                      break;
                  case "pt":
                      factor = 2.54 / 72.0;
                      break;
                  case "twip":
                      factor = 2.54 / 1440;
                      break;
                  case "pica":
                      factor = 2.54 / 6.0;
                      break;
                  case "dpt":
                      factor = 2.54 / 72.0;
                      break;
                  case "px":
                      factor = 2.54 / 96.0;
                      break;
                  default:
                      factor = 2.54 / 1440;
                      break;
              }

              if (double.TryParse(strValue, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out value))
              {
                  value = value * factor;
              }
          }
          return string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0.###}cm", value);
      }
      
      public class OdfStyleNameGenerator
      {
          // see http://www.w3.org/TR/REC-xml-names/#NT-NCName for allowed characters in NCName
          private Regex _invalidLettersAtStart = new Regex(@"^([^\p{Ll}\p{Lu}\p{Lo}\p{Lt}\p{Nl}_])(.*)", RegexOptions.Compiled); // valid start chars are \p{Ll}\p{Lu}\p{Lo}\p{Lt}\p{Nl}_
          private Regex _invalidChars = new Regex(@"[^_\p{Ll}\p{Lu}\p{Lo}\p{Lt}\p{Nl}\p{Mc}\p{Me}\p{Mn}\p{Lm}\p{Nd}]", RegexOptions.Compiled);

          private System.Collections.Generic.Dictionary<string, string> _name2ncname = new System.Collections.Generic.Dictionary<string, string>();
          private System.Collections.Generic.Dictionary<string, string> _ncname2name = new System.Collections.Generic.Dictionary<string, string>();

          private static OdfStyleNameGenerator _instance;

          private OdfStyleNameGenerator()
          {
          }

          public static OdfStyleNameGenerator Instance
          {
              get
              {
                  if (_instance == null)
                  {
                      _instance = new OdfStyleNameGenerator();
                  }
                  return _instance;
              }
          }

          /// <summary>
          /// Escapes all invalid characters from a style name and - if the name is not yet unique - 
          /// appends a counter to make the name unique
          /// </summary>
          public string NCNameFromString(string name)
          {
              string ncname = String.Empty;

              if (_name2ncname.ContainsKey(name))
              {
                  ncname = _name2ncname[name];
              }
              else
              {
                  // escape invalid characters
                  ncname = name;
                  Match invalidCharsMatch = _invalidChars.Match(ncname);
                  foreach (Capture invalidCharCapture in invalidCharsMatch.Captures)
                  {
                      ncname = ncname.Replace(invalidCharCapture.Value, string.Format("_{0:x}_", (int)invalidCharCapture.Value[0]));
                  }

                  // escape invalid start character
                  Match firstChar = _invalidLettersAtStart.Match(ncname);
                  if (firstChar.Success)
                  {
                      ncname = string.Format("_{0:x}_{1}", (int)firstChar.Groups[1].Value[0], firstChar.Groups[2].Value);
                  }

                  // create new unique ncname
                  int counter = 0;
                  string uniqueName = ncname;
                  while (_ncname2name.ContainsKey(uniqueName))
                  {
                      uniqueName = string.Format("{0}_{1}", ncname, counter);
                      counter++;
                  }
                  ncname = uniqueName;
                  _ncname2name.Add(ncname, name);
                  _name2ncname.Add(name, ncname);
              }
              return ncname;
          }
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
          string result = string.Empty;
          System.Uri uri = null;

          try
          {
              if (!System.Uri.TryCreate(path, System.UriKind.RelativeOrAbsolute, out uri))
              {
                  result = path;
              }
              else
              {
                  result = uri.ToString();
              }
          }
          catch (Exception ex)
          {
              System.Diagnostics.Trace.WriteLine(ex.ToString());
              result = path;
          }
          return result;
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

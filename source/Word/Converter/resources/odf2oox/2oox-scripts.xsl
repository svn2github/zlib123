<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt"
    xmlns:ooc="urn:odf-converter"
    exclude-result-prefixes="msxsl ooc">

  <msxsl:script language="C#" implements-prefix="ooc">
    <msxsl:using namespace="System"/>
    <msxsl:using namespace="System.IO"/>
    <msxsl:using namespace="System.Text.RegularExpressions"/>
    <msxsl:using namespace="System.Collections.Generic"/>
    
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
      
      ///<summary>
      /// Checks whether an URI points inside a package or outside
      /// 
      /// cf section 17.5 of ISO 26500:
      ///
      /// A relative-path reference (as described in §6.5 of [RFC3987]) that occurs in a file that is contained
      /// in a package has to be resolved exactly as it would be resolved if the whole package gets
      /// unzipped into a directory at its current location. The base IRI for resolving relative-path references
      /// is the one that has to be used to retrieve the (unzipped) file that contains the relative-path
      /// reference.
      /// All other kinds of IRI references, namely the ones that start with a protocol (like http:), an authority
      /// (i.e., //) or an absolute-path (i.e., /) do not need any special processing. This especially means that
      /// absolute-paths do not reference files inside the package, but within the hierarchy the package is
      /// contained in, for instance the file system. IRI references inside a package may leave the package,
      /// but once they have left the package, they never can return into the package or another one.
      ///</summary>
      public bool IsUriExternal(string path)
      {
          Uri result = null;
          
          try
          {
              // remove leading slash on absolute local paths (OOo uses /C:/somefolder for local paths)
              if (Regex.Match(path, "^/[A-Za-z]:").Success)
              {
                  path = path.Substring(1);
              }
              
              // escape special characters
              //path = Uri.EscapeUriString(path.Replace('\\', '/'));
              
              if (!Uri.TryCreate(path, UriKind.Absolute, out result))
              {
                  // check whether relative URI points into package or outside
                  Uri uriBase = new Uri(Path.Combine(Environment.CurrentDirectory, @"dummy.odt\"));
                  string absolutePath = uriBase.ToString() + path;
                  
                  if (Uri.TryCreate(absolutePath, UriKind.RelativeOrAbsolute, out result))
                  {
                      if (result.ToString().StartsWith(uriBase.ToString()))
                      {
                          // relative path inside the package
                          return false;
                      }
                  }
              }
          }
          catch
          {
          }
            
          return true;
      }

      public bool IsUriValid(string uri)
      {
          Uri tmp;
          return Uri.TryCreate(uri, UriKind.RelativeOrAbsolute, out tmp);
      }
      
      public string UriFromPath(string path)
      {
          try
          {
              Uri result = null;
              
              // remove leading slash on absolute local paths (OOo uses /C:/somefolder for local paths)
              if (Regex.Match(path, "^/[A-Za-z]:").Success)
              {
                  path = path.Substring(1);
              }
              
              // escape special characters
              //path = Uri.EscapeUriString(path.Replace('\\', '/'));
                        
              if (!Uri.TryCreate(path, UriKind.Absolute, out result))
              {
                  // check whether relative URI points into package or outside
                  Uri uriBase = new Uri(Path.Combine(Environment.CurrentDirectory, @"dummy.odt\"));
                  string absolutePath = uriBase.ToString() + path;
                  
                  if (Uri.TryCreate(absolutePath, UriKind.RelativeOrAbsolute, out result))
                  {
                      if (result.ToString().StartsWith(uriBase.ToString()))
                      {
                          // relative path inside the package
                          return result.ToString().Substring(uriBase.ToString().Length);
                      }
                  }
                  else
                  {
                      // documents with invalid URIs might not open so we replace the invalid URI with one that works -->
                      return "/";
                      //return path;
                  }
              }
              return result.ToString();
          }
          catch
          {
              return path;
          }
      }
      
      ///<summary>
      /// Convert alphanumeric bookmark ids to numeric ids
      ///</summary>
      Dictionary<string, int> bookmarkIds = new Dictionary<string, int>();
      public string GetBookmarkId(string id)
      {
          string result = id;
          if (!bookmarkIds.ContainsKey(id))
          {
              bookmarkIds[id] = bookmarkIds.Count + 1;
          }
          return bookmarkIds[id].ToString();
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

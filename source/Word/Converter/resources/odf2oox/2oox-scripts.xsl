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
      
      public string RegexReplace(string input, string pattern, string replacement, bool ignoreCase)
      {
          System.Text.RegularExpressions.RegexOptions options = System.Text.RegularExpressions.RegexOptions.Compiled | System.Text.RegularExpressions.RegexOptions.CultureInvariant;
          if (ignoreCase)
          {
              options |= System.Text.RegularExpressions.RegexOptions.IgnoreCase;
          }
          return System.Text.RegularExpressions.Regex.Replace(input, pattern, replacement, options);
      }
      
      public string ListSeparator()
      {
          return System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ListSeparator.ToString();
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
                  Uri uriBase = new Uri(System.IO.Path.Combine(Environment.CurrentDirectory, @"dummy.odt\"));
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
          return UriFromPath(path, false);
      }
      
      public string UriFromPath(string path, bool absolute)
      {
          string result = string.Empty;
              
          try
          {
              Uri uri = null;
              
              // remove leading slash on absolute local paths (OOo uses /C:/somefolder for local paths)
              if (Regex.Match(path, "^/[A-Za-z]:").Success)
              {
                  path = path.Substring(1);
              }
              else if (string.IsNullOrEmpty(path))
              {
                  // workaround for empty href (Word does not like an empty Target attribute)
                  path = "#";
              }
              
              // escape special characters
              //path = Uri.EscapeUriString(path.Replace('\\', '/'));
                        
              if (!Uri.TryCreate(path, UriKind.Absolute, out uri))
              {
                  // check whether relative URI points into package or outside
                  Uri uriBase = new Uri(System.IO.Path.Combine(Environment.CurrentDirectory, @"dummy.odt\"));
                  string absolutePath = uriBase.ToString() + path;
                  
                  if (Uri.TryCreate(absolutePath, UriKind.RelativeOrAbsolute, out uri))
                  {
                      if (uri.ToString().StartsWith(uriBase.ToString()))
                      {
                          // relative path inside the package
                          result = uri.ToString().Substring(uriBase.ToString().Length);
                      }
                      else
                      {
                          if (absolute)
                          {
                              result = uri.ToString();
                          }
                          else
                          {
                              Uri uriCurrentDir = new Uri(
                                  Environment.CurrentDirectory.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()) ? 
                                    Environment.CurrentDirectory : Environment.CurrentDirectory + System.IO.Path.DirectorySeparatorChar);
                              result = uriCurrentDir.MakeRelativeUri(uri).ToString(); //uri.ToString().Substring(uriBase.ToString().Length - @"dummy.odt\".Length);
                          }
                      }
                  }
                  else
                  {
                      // documents with invalid URIs might not open so we replace the invalid URI with one that works -->
                      result = "/";
                      //return path;
                  }
              }
              else
              {
                  result = uri.ToString();
              }
          }
          catch
          {
              result = path;
          }
          return Uri.EscapeUriString(result);
      }
      
      ///<summary>
      /// Convert alphanumeric bookmark ids to numeric ids
      ///</summary>
      System.Collections.Generic.Dictionary<string, int> bookmarkIds = new System.Collections.Generic.Dictionary<string, int>();
      public string GetBookmarkId(string id)
      {
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
      
      
      ///<summary>
      /// Convert various length units to point. The return value does not have the unit appended.
      ///</summary>
      public double PtFromMeasuredUnit(string measuredUnit, int precision)
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
                      factor = 72.0 / 2.54;
                      break;
                  case "mm":
                      factor = 72.0 / 25.4;
                      break;
                  case "in":
                  case "inch":
                      factor = 72.0;
                      break;
                  case "pt":
                      factor = 1.0;
                      break;
                  case "pica":
                      factor = 12;
                      break;
                  case "dpt":
                      factor = 1.0;
                      break;
                  case "px":
                      factor = 72.0 / 96.0;
                      break;
              }

              if (double.TryParse(strValue, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out value))
              {
                  value = value * factor;
              }
          }
          return Math.Round(value, precision, MidpointRounding.AwayFromZero); 
      }
      
      
      
      
      
      
      public class OoxDefaultStyleVisibility
      {
          private static System.Collections.Generic.Dictionary<string, OoxDefaultStyleVisibility> _properties;

          private bool _customStyle = false;

          private bool _hidden = false;
          private int _uiPriority = 99;
          private bool _semiHidden = false;
          private bool _unhideWhenUsed = false;
          private bool _qFormat = true;

          public bool CustomStyle
          {
              get { return _customStyle; }
              set { _customStyle = value; }
          }

          public bool Hidden
          {
              get { return _hidden; }
              set { _hidden = value; }
          }
          
          public int UiPriority
          {
              get { return _uiPriority; }
              set { _uiPriority = value; }
          }
          
          public bool SemiHidden
          {
              get { return _semiHidden; }
              set { _semiHidden = value; }
          }

          public bool UnhideWhenUsed
          {
              get { return _unhideWhenUsed; }
              set { _unhideWhenUsed = value; }
          }

          public bool QFormat
          {
              get { return _qFormat; }
              set { _qFormat = value; }
          }

          public OoxDefaultStyleVisibility(bool customStyle, bool hidden, int uiPriority, bool semiHidden, bool unhideWhenUsed, bool qFormat)
          {
              this.CustomStyle = customStyle;
              this.Hidden = hidden;
              this.UiPriority = uiPriority;
              this.SemiHidden = semiHidden;
              this.UnhideWhenUsed = unhideWhenUsed;
              this.QFormat = qFormat;
          }

          public static OoxDefaultStyleVisibility GetDefaultProperties(string styleId)
          {
              if (_properties == null)
              {
                  _properties = new System.Collections.Generic.Dictionary<string, OoxDefaultStyleVisibility>();

                  _properties.Add("1/1.1/1.1.1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("1/a/i", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("article/section", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("balloontext", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("bibliography", new OoxDefaultStyleVisibility(false, false, 37, true, true, false));
                  _properties.Add("blocktext", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("bodytext", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("bodytext2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("bodytext3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("bodytextfirstindent", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("bodytextfirstindent2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("bodytextindent", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("bodytextindent2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("bodytextindent3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("booktitle", new OoxDefaultStyleVisibility(false, false, 33, false, false, true));
                  _properties.Add("caption", new OoxDefaultStyleVisibility(false, false, 35, true, true, true));
                  _properties.Add("closing", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("colorfulgrid", new OoxDefaultStyleVisibility(false, false, 73, false, false, false));
                  _properties.Add("colorfulgrid-accent1", new OoxDefaultStyleVisibility(false, false, 73, false, false, false));
                  _properties.Add("colorfulgrid-accent2", new OoxDefaultStyleVisibility(false, false, 73, false, false, false));
                  _properties.Add("colorfulgrid-accent3", new OoxDefaultStyleVisibility(false, false, 73, false, false, false));
                  _properties.Add("colorfulgrid-accent4", new OoxDefaultStyleVisibility(false, false, 73, false, false, false));
                  _properties.Add("colorfulgrid-accent5", new OoxDefaultStyleVisibility(false, false, 73, false, false, false));
                  _properties.Add("colorfulgrid-accent6", new OoxDefaultStyleVisibility(false, false, 73, false, false, false));
                  _properties.Add("colorfullist", new OoxDefaultStyleVisibility(false, false, 72, false, false, false));
                  _properties.Add("colorfullist-accent1", new OoxDefaultStyleVisibility(false, false, 72, false, false, false));
                  _properties.Add("colorfullist-accent2", new OoxDefaultStyleVisibility(false, false, 72, false, false, false));
                  _properties.Add("colorfullist-accent3", new OoxDefaultStyleVisibility(false, false, 72, false, false, false));
                  _properties.Add("colorfullist-accent4", new OoxDefaultStyleVisibility(false, false, 72, false, false, false));
                  _properties.Add("colorfullist-accent5", new OoxDefaultStyleVisibility(false, false, 72, false, false, false));
                  _properties.Add("colorfullist-accent6", new OoxDefaultStyleVisibility(false, false, 72, false, false, false));
                  _properties.Add("colorfulshading", new OoxDefaultStyleVisibility(false, false, 71, false, false, false));
                  _properties.Add("colorfulshading-accent1", new OoxDefaultStyleVisibility(false, false, 71, false, false, false));
                  _properties.Add("colorfulshading-accent2", new OoxDefaultStyleVisibility(false, false, 71, false, false, false));
                  _properties.Add("colorfulshading-accent3", new OoxDefaultStyleVisibility(false, false, 71, false, false, false));
                  _properties.Add("colorfulshading-accent4", new OoxDefaultStyleVisibility(false, false, 71, false, false, false));
                  _properties.Add("colorfulshading-accent5", new OoxDefaultStyleVisibility(false, false, 71, false, false, false));
                  _properties.Add("colorfulshading-accent6", new OoxDefaultStyleVisibility(false, false, 71, false, false, false));
                  _properties.Add("commentreference", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("commentsubject", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("commenttext", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("darklist", new OoxDefaultStyleVisibility(false, false, 70, false, false, false));
                  _properties.Add("darklist-accent1", new OoxDefaultStyleVisibility(false, false, 70, false, false, false));
                  _properties.Add("darklist-accent2", new OoxDefaultStyleVisibility(false, false, 70, false, false, false));
                  _properties.Add("darklist-accent3", new OoxDefaultStyleVisibility(false, false, 70, false, false, false));
                  _properties.Add("darklist-accent4", new OoxDefaultStyleVisibility(false, false, 70, false, false, false));
                  _properties.Add("darklist-accent5", new OoxDefaultStyleVisibility(false, false, 70, false, false, false));
                  _properties.Add("darklist-accent6", new OoxDefaultStyleVisibility(false, false, 70, false, false, false));
                  _properties.Add("date", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("defaultparagraphfont", new OoxDefaultStyleVisibility(false, false, 1, true, true, false));
                  _properties.Add("documentmap", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("e-mailsignature", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("emphasis", new OoxDefaultStyleVisibility(false, false, 20, false, false, true));
                  _properties.Add("endnotereference", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("endnotetext", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("envelopeaddress", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("envelopereturn", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("followedhyperlink", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("footer", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("footnotereference", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("footnotetext", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("header", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("heading1", new OoxDefaultStyleVisibility(false, false, 9, false, false, true));
                  _properties.Add("heading2", new OoxDefaultStyleVisibility(false, false, 9, true, true, true));
                  _properties.Add("heading3", new OoxDefaultStyleVisibility(false, false, 9, true, true, true));
                  _properties.Add("heading4", new OoxDefaultStyleVisibility(false, false, 9, true, true, true));
                  _properties.Add("heading5", new OoxDefaultStyleVisibility(false, false, 9, true, true, true));
                  _properties.Add("heading6", new OoxDefaultStyleVisibility(false, false, 9, true, true, true));
                  _properties.Add("heading7", new OoxDefaultStyleVisibility(false, false, 9, true, true, true));
                  _properties.Add("heading8", new OoxDefaultStyleVisibility(false, false, 9, true, true, true));
                  _properties.Add("heading9", new OoxDefaultStyleVisibility(false, false, 9, true, true, true));
                  _properties.Add("htmlacronym", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("htmladdress", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("htmlcite", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("htmlcode", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("htmldefinition", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("htmlkeyboard", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("htmlpreformatted", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("htmlsample", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("htmltypewriter", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("htmlvariable", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("hyperlink", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("index1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("index2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("index3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("index4", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("index5", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("index6", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("index7", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("index8", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("index9", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("indexheading", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("intenseemphasis", new OoxDefaultStyleVisibility(false, false, 21, false, false, true));
                  _properties.Add("intensequote", new OoxDefaultStyleVisibility(false, false, 30, false, false, true));
                  _properties.Add("intensereference", new OoxDefaultStyleVisibility(false, false, 32, false, false, true));
                  _properties.Add("lightgrid", new OoxDefaultStyleVisibility(false, false, 62, false, false, false));
                  _properties.Add("lightgrid-accent1", new OoxDefaultStyleVisibility(false, false, 62, false, false, false));
                  _properties.Add("lightgrid-accent2", new OoxDefaultStyleVisibility(false, false, 62, false, false, false));
                  _properties.Add("lightgrid-accent3", new OoxDefaultStyleVisibility(false, false, 62, false, false, false));
                  _properties.Add("lightgrid-accent4", new OoxDefaultStyleVisibility(false, false, 62, false, false, false));
                  _properties.Add("lightgrid-accent5", new OoxDefaultStyleVisibility(false, false, 62, false, false, false));
                  _properties.Add("lightgrid-accent6", new OoxDefaultStyleVisibility(false, false, 62, false, false, false));
                  _properties.Add("lightlist", new OoxDefaultStyleVisibility(false, false, 61, false, false, false));
                  _properties.Add("lightlist-accent1", new OoxDefaultStyleVisibility(false, false, 61, false, false, false));
                  _properties.Add("lightlist-accent2", new OoxDefaultStyleVisibility(false, false, 61, false, false, false));
                  _properties.Add("lightlist-accent3", new OoxDefaultStyleVisibility(false, false, 61, false, false, false));
                  _properties.Add("lightlist-accent4", new OoxDefaultStyleVisibility(false, false, 61, false, false, false));
                  _properties.Add("lightlist-accent5", new OoxDefaultStyleVisibility(false, false, 61, false, false, false));
                  _properties.Add("lightlist-accent6", new OoxDefaultStyleVisibility(false, false, 61, false, false, false));
                  _properties.Add("lightshading", new OoxDefaultStyleVisibility(false, false, 60, false, false, false));
                  _properties.Add("lightshading-accent1", new OoxDefaultStyleVisibility(false, false, 60, false, false, false));
                  _properties.Add("lightshading-accent2", new OoxDefaultStyleVisibility(false, false, 60, false, false, false));
                  _properties.Add("lightshading-accent3", new OoxDefaultStyleVisibility(false, false, 60, false, false, false));
                  _properties.Add("lightshading-accent4", new OoxDefaultStyleVisibility(false, false, 60, false, false, false));
                  _properties.Add("lightshading-accent5", new OoxDefaultStyleVisibility(false, false, 60, false, false, false));
                  _properties.Add("lightshading-accent6", new OoxDefaultStyleVisibility(false, false, 60, false, false, false));
                  _properties.Add("linenumber", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("list", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("list2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("list3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("list4", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("list5", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("listbullet", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("listbullet2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("listbullet3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("listbullet4", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("listbullet5", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("listcontinue", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("listcontinue2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("listcontinue3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("listcontinue4", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("listcontinue5", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("listnumber", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("listnumber2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("listnumber3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("listnumber4", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("listnumber5", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("listparagraph", new OoxDefaultStyleVisibility(false, false, 34, false, false, true));
                  _properties.Add("macrotext", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("mediumgrid1", new OoxDefaultStyleVisibility(false, false, 67, false, false, false));
                  _properties.Add("mediumgrid1-accent1", new OoxDefaultStyleVisibility(false, false, 67, false, false, false));
                  _properties.Add("mediumgrid1-accent2", new OoxDefaultStyleVisibility(false, false, 67, false, false, false));
                  _properties.Add("mediumgrid1-accent3", new OoxDefaultStyleVisibility(false, false, 67, false, false, false));
                  _properties.Add("mediumgrid1-accent4", new OoxDefaultStyleVisibility(false, false, 67, false, false, false));
                  _properties.Add("mediumgrid1-accent5", new OoxDefaultStyleVisibility(false, false, 67, false, false, false));
                  _properties.Add("mediumgrid1-accent6", new OoxDefaultStyleVisibility(false, false, 67, false, false, false));
                  _properties.Add("mediumgrid2", new OoxDefaultStyleVisibility(false, false, 68, false, false, false));
                  _properties.Add("mediumgrid2-accent1", new OoxDefaultStyleVisibility(false, false, 68, false, false, false));
                  _properties.Add("mediumgrid2-accent2", new OoxDefaultStyleVisibility(false, false, 68, false, false, false));
                  _properties.Add("mediumgrid2-accent3", new OoxDefaultStyleVisibility(false, false, 68, false, false, false));
                  _properties.Add("mediumgrid2-accent4", new OoxDefaultStyleVisibility(false, false, 68, false, false, false));
                  _properties.Add("mediumgrid2-accent5", new OoxDefaultStyleVisibility(false, false, 68, false, false, false));
                  _properties.Add("mediumgrid2-accent6", new OoxDefaultStyleVisibility(false, false, 68, false, false, false));
                  _properties.Add("mediumgrid3", new OoxDefaultStyleVisibility(false, false, 69, false, false, false));
                  _properties.Add("mediumgrid3-accent1", new OoxDefaultStyleVisibility(false, false, 69, false, false, false));
                  _properties.Add("mediumgrid3-accent2", new OoxDefaultStyleVisibility(false, false, 69, false, false, false));
                  _properties.Add("mediumgrid3-accent3", new OoxDefaultStyleVisibility(false, false, 69, false, false, false));
                  _properties.Add("mediumgrid3-accent4", new OoxDefaultStyleVisibility(false, false, 69, false, false, false));
                  _properties.Add("mediumgrid3-accent5", new OoxDefaultStyleVisibility(false, false, 69, false, false, false));
                  _properties.Add("mediumgrid3-accent6", new OoxDefaultStyleVisibility(false, false, 69, false, false, false));
                  _properties.Add("mediumlist1", new OoxDefaultStyleVisibility(false, false, 65, false, false, false));
                  _properties.Add("mediumlist1-accent1", new OoxDefaultStyleVisibility(false, false, 65, false, false, false));
                  _properties.Add("mediumlist1-accent2", new OoxDefaultStyleVisibility(false, false, 65, false, false, false));
                  _properties.Add("mediumlist1-accent3", new OoxDefaultStyleVisibility(false, false, 65, false, false, false));
                  _properties.Add("mediumlist1-accent4", new OoxDefaultStyleVisibility(false, false, 65, false, false, false));
                  _properties.Add("mediumlist1-accent5", new OoxDefaultStyleVisibility(false, false, 65, false, false, false));
                  _properties.Add("mediumlist1-accent6", new OoxDefaultStyleVisibility(false, false, 65, false, false, false));
                  _properties.Add("mediumlist2", new OoxDefaultStyleVisibility(false, false, 66, false, false, false));
                  _properties.Add("mediumlist2-accent1", new OoxDefaultStyleVisibility(false, false, 66, false, false, false));
                  _properties.Add("mediumlist2-accent2", new OoxDefaultStyleVisibility(false, false, 66, false, false, false));
                  _properties.Add("mediumlist2-accent3", new OoxDefaultStyleVisibility(false, false, 66, false, false, false));
                  _properties.Add("mediumlist2-accent4", new OoxDefaultStyleVisibility(false, false, 66, false, false, false));
                  _properties.Add("mediumlist2-accent5", new OoxDefaultStyleVisibility(false, false, 66, false, false, false));
                  _properties.Add("mediumlist2-accent6", new OoxDefaultStyleVisibility(false, false, 66, false, false, false));
                  _properties.Add("mediumshading1", new OoxDefaultStyleVisibility(false, false, 63, false, false, false));
                  _properties.Add("mediumshading1-accent1", new OoxDefaultStyleVisibility(false, false, 63, false, false, false));
                  _properties.Add("mediumshading1-accent2", new OoxDefaultStyleVisibility(false, false, 63, false, false, false));
                  _properties.Add("mediumshading1-accent3", new OoxDefaultStyleVisibility(false, false, 63, false, false, false));
                  _properties.Add("mediumshading1-accent4", new OoxDefaultStyleVisibility(false, false, 63, false, false, false));
                  _properties.Add("mediumshading1-accent5", new OoxDefaultStyleVisibility(false, false, 63, false, false, false));
                  _properties.Add("mediumshading1-accent6", new OoxDefaultStyleVisibility(false, false, 63, false, false, false));
                  _properties.Add("mediumshading2", new OoxDefaultStyleVisibility(false, false, 64, false, false, false));
                  _properties.Add("mediumshading2-accent1", new OoxDefaultStyleVisibility(false, false, 64, false, false, false));
                  _properties.Add("mediumshading2-accent2", new OoxDefaultStyleVisibility(false, false, 64, false, false, false));
                  _properties.Add("mediumshading2-accent3", new OoxDefaultStyleVisibility(false, false, 64, false, false, false));
                  _properties.Add("mediumshading2-accent4", new OoxDefaultStyleVisibility(false, false, 64, false, false, false));
                  _properties.Add("mediumshading2-accent5", new OoxDefaultStyleVisibility(false, false, 64, false, false, false));
                  _properties.Add("mediumshading2-accent6", new OoxDefaultStyleVisibility(false, false, 64, false, false, false));
                  _properties.Add("messageheader", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("nolist", new OoxDefaultStyleVisibility(false, false, 99, false, false, false));
                  _properties.Add("nospacing", new OoxDefaultStyleVisibility(false, false, 1, false, false, true));
                  _properties.Add("normal", new OoxDefaultStyleVisibility(false, false, 0, false, false, true));
                  _properties.Add("normal(web)", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("normalindent", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("noteheading", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("pagenumber", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("placeholdertext", new OoxDefaultStyleVisibility(false, true, 99, false, false, false));
                  _properties.Add("plaintext", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("quote", new OoxDefaultStyleVisibility(false, false, 29, false, false, true));
                  _properties.Add("salutation", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("signature", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("strong", new OoxDefaultStyleVisibility(false, false, 22, false, false, true));
                  _properties.Add("subtitle", new OoxDefaultStyleVisibility(false, false, 11, false, false, true));
                  _properties.Add("subtleemphasis", new OoxDefaultStyleVisibility(false, false, 19, false, false, true));
                  _properties.Add("subtlereference", new OoxDefaultStyleVisibility(false, false, 31, false, false, true));
                  _properties.Add("table3deffects1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("table3deffects2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("table3deffects3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tableclassic1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tableclassic2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tableclassic3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tableclassic4", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablecolorful1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablecolorful2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablecolorful3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablecolumns1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablecolumns2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablecolumns3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablecolumns4", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablecolumns5", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablecontemporary", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tableelegant", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablegrid", new OoxDefaultStyleVisibility(false, false, 59, false, false, false));
                  _properties.Add("tablegrid1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablegrid2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablegrid3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablegrid4", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablegrid5", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablegrid6", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablegrid7", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablegrid8", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablelist1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablelist2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablelist3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablelist4", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablelist5", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablelist6", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablelist7", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablelist8", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablenormal", new OoxDefaultStyleVisibility(false, false, 99, false, false, false));
                  _properties.Add("tableofauthorities", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tableoffigures", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tableprofessional", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablesimple1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablesimple2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablesimple3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablesubtle1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tablesubtle2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tabletheme", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tableweb1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tableweb2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("tableweb3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("title", new OoxDefaultStyleVisibility(false, false, 10, false, false, true));
                  _properties.Add("toaheading", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
                  _properties.Add("toc1", new OoxDefaultStyleVisibility(false, false, 39, true, true, false));
                  _properties.Add("toc2", new OoxDefaultStyleVisibility(false, false, 39, true, true, false));
                  _properties.Add("toc3", new OoxDefaultStyleVisibility(false, false, 39, true, true, false));
                  _properties.Add("toc4", new OoxDefaultStyleVisibility(false, false, 39, true, true, false));
                  _properties.Add("toc5", new OoxDefaultStyleVisibility(false, false, 39, true, true, false));
                  _properties.Add("toc6", new OoxDefaultStyleVisibility(false, false, 39, true, true, false));
                  _properties.Add("toc7", new OoxDefaultStyleVisibility(false, false, 39, true, true, false));
                  _properties.Add("toc8", new OoxDefaultStyleVisibility(false, false, 39, true, true, false));
                  _properties.Add("toc9", new OoxDefaultStyleVisibility(false, false, 39, true, true, false));
                  _properties.Add("tocheading", new OoxDefaultStyleVisibility(false, false, 39, true, true, true));
              }

              if (_properties.ContainsKey(styleId.Replace(" ", "").ToLower()))
              {
                  return _properties[styleId.Replace(" ", "").ToLower()];
              }
              else if (styleId.ToLower().EndsWith("char") && _properties.ContainsKey(styleId.Substring(0, styleId.Length - 4).ToLower()))
              {
                  // make Word-generated linked character styles invisible
                  return new OoxDefaultStyleVisibility(true, false, 99, true, false, false);
              }

              return new OoxDefaultStyleVisibility(true, false, 0, false, false, true);
          }
      }

  
      public bool CustomStyle(string styleId)
      {
          OoxDefaultStyleVisibility properties = OoxDefaultStyleVisibility.GetDefaultProperties(styleId);
          return properties.CustomStyle;
      }
      
      public bool Hidden(string styleId)
      {
          OoxDefaultStyleVisibility properties = OoxDefaultStyleVisibility.GetDefaultProperties(styleId);
          return properties.Hidden;
      }
      
      public int UiPriority(string styleId)
      {
          OoxDefaultStyleVisibility properties = OoxDefaultStyleVisibility.GetDefaultProperties(styleId);
          return properties.UiPriority;
      }
      
      public bool SemiHidden(string styleId)
      {
          OoxDefaultStyleVisibility properties = OoxDefaultStyleVisibility.GetDefaultProperties(styleId);
          return properties.SemiHidden;
      }
      
      public bool UnhideWhenUsed(string styleId)
      {
          OoxDefaultStyleVisibility properties = OoxDefaultStyleVisibility.GetDefaultProperties(styleId);
          return properties.UnhideWhenUsed;
      }
      
      public bool QFormat(string styleId)
      {
          OoxDefaultStyleVisibility properties = OoxDefaultStyleVisibility.GetDefaultProperties(styleId);
          return properties.QFormat;
      }

      
    ]]>

  </msxsl:script>

</xsl:stylesheet>

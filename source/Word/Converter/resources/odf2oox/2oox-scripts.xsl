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
          return Uri.TryCreate(uri, UriKind.Absolute, out tmp);
      }

      public bool IsUriRelative(string uri)
      {
          Uri tmp;
          return Uri.TryCreate(uri, UriKind.Relative, out tmp);
      }

      public bool IsUriValid(string uri)
      {
          Uri tmp;
          return Uri.TryCreate(uri, UriKind.RelativeOrAbsolute, out tmp);
      }
    ]]>

  </msxsl:script>

</xsl:stylesheet>

<?xml version="1.0" encoding="UTF-8" ?>
<!--
   Pradeep Nemadi
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:v="urn:schemas-microsoft-com:vml" exclude-result-prefixes="w r draw number wp xlink">

   
  <xsl:key name="StyleId" match="w:style" use="@w:styleId"/>
  <xsl:key name="default-styles"
    match="w:style[@w:default = 1 or @w:default = 'true' or @w:default = 'on']" use="@w:type"/>

  <xsl:template name="styles">
    <office:document-styles>
      <office:font-face-decls>
        <xsl:apply-templates select="document('ppt/fontTable.xml')/w:fonts"/>
      </office:font-face-decls>
      <!-- document styles -->
      <office:styles>
        <xsl:call-template name="InsertDefaultStyles"/>
      </office:styles>
      <!-- automatic styles -->
      <office:automatic-styles>
	  </office:automatic-styles>
      <!-- master styles -->
      <office:master-styles>
      </office:master-styles>
    </office:document-styles>
  </xsl:template>

	<xsl:template name="InsertDefaultStyles">

    <xsl:for-each select="document('ppt/styles.xml')">
     
    </xsl:for-each>
  </xsl:template>
  
</xsl:stylesheet>

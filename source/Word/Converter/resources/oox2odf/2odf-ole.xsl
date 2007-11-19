<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
                xmlns:o="urn:schemas-microsoft-com:office:office"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:rels="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:xlink="http://www.w3.org/1999/xlink"
                xmlns:v="urn:schemas-microsoft-com:vml"
                xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
                xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
                exclude-result-prefixes="xlink draw">

  <!-- 
  *************************************************************************
  MATCHING TEMPLATES
  *************************************************************************
  -->

  <!--
  Summary: Converts OLE objects
  Author: makz (DIaLOGIKa)
  Date: 14.11.2007
  -->
  <xsl:template match="o:OLEObject">
    <draw:frame>
      <xsl:call-template name="InsertCommonShapeProperties">
        <xsl:with-param name="shape" select="../v:shape" />
      </xsl:call-template>
      <xsl:call-template name="InsertShapeZindex"/>
      
      <xsl:call-template name="InsertObject" />
      <xsl:call-template name="InsertObjectPreview" />
    </draw:frame>
  </xsl:template>

  <!--
  Summary: Converts the style for OLE objects
  Author: makz (DIaLOGIKa)
  Date: 14.11.2007
  -->
  <xsl:template match="o:OLEObject" mode="automaticstyles">
    <style:style style:family="graphic" style:parent-style-name="Frame">
      <xsl:attribute name="style:name">
          <xsl:value-of select="generate-id(../v:shape)"/>
      </xsl:attribute>
      
      <style:graphic-properties>
        <xsl:call-template name="InsertShapeStyleProperties">
          <xsl:with-param name="shape" select="../v:shape" />
        </xsl:call-template>
      </style:graphic-properties>
    </style:style>
  </xsl:template>

  <!-- 
  *************************************************************************
  CALLED TEMPLATES
  *************************************************************************
  -->

  <!--
  Summary: inserts the object itself
  Author: makz (DIaLOGIKa)
  Date: 14.11.2007
  -->
  <xsl:template name="InsertObject">
    <xsl:variable name="oId" select="@r:id" />
    <xsl:variable name="filePath" select="key('Part', 'word/_rels/document.xml.rels')/rels:Relationships/rels:Relationship[@Id=$oId]/@Target" />
    <xsl:variable name="fileName" select="substring-after($filePath, 'embeddings/')" />
    <xsl:variable name="newName" select="concat(substring-before($fileName, '.'), '.bin')" />
    
    <xsl:choose>
      <!-- it's an embedded object -->
      <xsl:when test="@Type='Embed'">
        <!-- copy the embedded object -->
        <pzip:copy pzip:source="{concat('word/embeddings/', $fileName)}" pzip:target="{$newName}"/>
        
        <draw:object-ole>
          <xsl:call-template name="InsertObjectHref">
            <xsl:with-param name="link" select="$newName" />
          </xsl:call-template>
          <xsl:call-template name="InsertObjectShow" />
          <xsl:call-template name="InsertObjectType" />
          <xsl:call-template name="InsertObjectActuate" />
        </draw:object-ole>
      </xsl:when>
      <!-- it's an external object -->
      <xsl:otherwise>
        <draw:object>
          <xsl:call-template name="InsertObjectHref">
            <xsl:with-param name="link" select="concat('http://www.dialogika.de/odf-converter/makerelpath#', substring-after($filePath, 'file:///'))" />
          </xsl:call-template>
          <xsl:call-template name="InsertObjectShow" />
          <xsl:call-template name="InsertObjectType" />
          <xsl:call-template name="InsertObjectActuate" />
        </draw:object>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--
  Summary: inserts the preview picture of the object
  Author: makz (DIaLOGIKa)
  Date: 14.11.2007
  -->
  <xsl:template name="InsertObjectPreview">
   
    <xsl:variable name="imageId" select="../v:shape/v:imagedata/@r:id" />
    <xsl:variable name="filePath" select="key('Part', 'word/_rels/document.xml.rels')/rels:Relationships/rels:Relationship[@Id=$imageId]/@Target" />
    <xsl:variable name="fileName" select="substring-after($filePath, 'media/')" />
    <xsl:variable name="suffix" select="substring-after($fileName, '.')" />
    
    <draw:image>
      <xsl:call-template name="InsertObjectShow" />
      <xsl:call-template name="InsertObjectType" />
      <xsl:call-template name="InsertObjectActuate" />
      
      <xsl:choose>
        <xsl:when test="$suffix='wmf' or $suffix='emf'">
          <xsl:call-template name="InsertObjectHref">
            <xsl:with-param name="link" select="'./ObjectReplacements/OLEplaceholder.png'" />
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="InsertObjectHref">
            <xsl:with-param name="link" select="concat('./ObjectReplacements/', $fileName)" />
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
      
      <!-- copy placeholder for unsupported images -->
      <xsl:choose>
        <xsl:when test="$suffix='wmf' or $suffix='emf'">
          <pzip:copy pzip:source="#CER#WordConverter.dll#CleverAge.OdfConverter.Word.resources.OLEplaceholder.png#"
                     pzip:target="ObjectReplacements/OLEplaceholder.png"/>
        </xsl:when>
        <xsl:otherwise>
          <pzip:copy pzip:source="{concat('word/media/', $fileName)}" 
                     pzip:target="{concat('ObjectReplacements/', $fileName)}"/>
        </xsl:otherwise>
      </xsl:choose>
      
    </draw:image>
  </xsl:template>


  <!--
  Summary: Inserts the href attribute to the object and copies the file if it is internal
  Author: makz (DIaLOGIKa)
  Date: 14.11.2007
  -->
  <xsl:template name="InsertObjectHref">
    <xsl:param name="link" select="''" />
    
    <xsl:attribute name="xlink:href" >
      <xsl:value-of select="$link"/>
    </xsl:attribute>
  </xsl:template>

  <!--
  Summary: Inserts the show attribute to the object
  Author: makz (DIaLOGIKa)
  Date: 14.11.2007
  -->
  <xsl:template name="InsertObjectShow">
    <xsl:attribute name="xlink:show">
      <xsl:text>embed</xsl:text>
    </xsl:attribute>
  </xsl:template>

  <!--
  Summary: Inserts the type attribute to the object
  Author: makz (DIaLOGIKa)
  Date: 14.11.2007
  -->
  <xsl:template name="InsertObjectType">
    <xsl:attribute name="xlink:type">
      <xsl:text>simple</xsl:text>
    </xsl:attribute>
  </xsl:template>

  <!--
  Summary: Inserts the actuate attribute to the object
  Author: makz (DIaLOGIKa)
  Date: 14.11.2007
  -->
  <xsl:template name="InsertObjectActuate">
    <xsl:attribute name="xlink:actuate">
      <xsl:text>onLoad</xsl:text>
    </xsl:attribute>
  </xsl:template>
  
</xsl:stylesheet>
<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  exclude-result-prefixes="w draw">
   
  <xsl:template match="w:pict" >
    <draw:frame>
      <xsl:call-template name="InsertShapeStyleName"/>
      <xsl:call-template name="InsertShapeWidth"/>
      <xsl:call-template name="InsertShapeHeight"/>
      <xsl:call-template name="InsertShapeAnchor"/>
      <xsl:call-template name="InsertShapeZindex"/>
      <xsl:apply-templates select="v:shape/child::node()"/>
    </draw:frame>
  </xsl:template>
  
  <xsl:template match="v:textbox">
    <draw:text-box >
      <xsl:apply-templates select="w:txbxContent/child::node()"/>
    </draw:text-box>
  </xsl:template>
  
  <xsl:template name="InsertShapeAnchor">
    <xsl:param name="shape" select="v:shape"/>
    <!--TODO page anchor -->
    <xsl:attribute name="text:anchor-type">
      <xsl:choose>
        <xsl:when test="descendant::w10:wrap/@type = 'none' ">
          <xsl:text>as-char</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:text>paragraph</xsl:text>
        </xsl:otherwise>
      </xsl:choose>
     </xsl:attribute>
  </xsl:template>
  
  <xsl:template name="InsertShapeZindex">
    <xsl:param name="shape" select="v:shape"/>
    <!--TODO-->
    <xsl:attribute name="draw:z-index">
        <xsl:text>0</xsl:text>
    </xsl:attribute>
  </xsl:template>
  
  <xsl:template name="InsertShapeStyleName">
    <xsl:param name="shape" select="v:shape"/>
    <xsl:attribute name="draw:style-name">
      <xsl:call-template name="GenerateStyleName">
        <xsl:with-param name="node" select="$shape"/>
      </xsl:call-template>
    </xsl:attribute>
    <!--TODO-->
    <xsl:attribute name="draw:name">
      <xsl:text>Frame1</xsl:text>
    </xsl:attribute>
  </xsl:template>
  
  <xsl:template name="InsertShapeHeight">
    <xsl:param name="shape" select="v:shape"/>
    <xsl:variable name="height">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$shape"/>
        <xsl:with-param name="propertyName" select="'height'"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:attribute name="svg:height">
      <xsl:call-template name="ConvertPoints">
        <xsl:with-param name="length" >
          <xsl:call-template name="GetValue">
            <xsl:with-param name="length" select="$height"/>
          </xsl:call-template>
        </xsl:with-param>
        <xsl:with-param name="unit" select="'cm'"/>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>
  
  <xsl:template name="InsertShapeWidth">
    <xsl:param name="shape" select="v:shape"/>
    <xsl:variable name="width">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$shape"/>
        <xsl:with-param name="propertyName" select="'width'"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:attribute name="svg:width">
      <xsl:call-template name="ConvertPoints">
        <xsl:with-param name="length" >
          <xsl:call-template name="GetValue">
            <xsl:with-param name="length" select="$width"/>
          </xsl:call-template>
        </xsl:with-param>
        <xsl:with-param name="unit" select="'cm'"/>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>
  
  <xsl:template name="GetShapeProperty">
    <xsl:param name="shape"/>
    <xsl:param name="propertyName"/>
     <xsl:variable name="propertyValue" select="substring-before(substring-after($shape/@style,concat($propertyName,':')),';')"/>
    <xsl:value-of select="$propertyValue"/>
  </xsl:template>
  
  <xsl:template name="GenerateStyleName">
    <xsl:param name="node"/>
    <xsl:choose>
      <xsl:when test="name($node) = 'v:shape'">
          <xsl:value-of select="concat('G_',generate-id($node))"/>
      </xsl:when>
      <xsl:otherwise>
          <xsl:value-of select="generate-id($node)"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template match="w:pict" mode="automaticstyles">
    <style:style >
      <xsl:attribute name="style:name">
        <xsl:call-template name="GenerateStyleName">
          <xsl:with-param name="node" select="v:shape"/>
        </xsl:call-template>
      </xsl:attribute>
      <!--in Word there are no parent style for image - make default Graphics in OO -->
      <xsl:attribute name="style:parent-style-name">
        <xsl:text>Graphics</xsl:text>
        <xsl:value-of select="w:tblStyle/@w:val"/>
      </xsl:attribute>
      <xsl:attribute name="style:family">
        <xsl:text>graphic</xsl:text>
      </xsl:attribute>
      
      <style:graphic-properties>
        <xsl:call-template name="InsertShapeProperties"/>
      </style:graphic-properties>
    </style:style>
  </xsl:template>
  
  <xsl:template name="InsertShapeProperties">
    <!--TODO automatic style content-->
    <xsl:call-template name="InsertShapeWrap"/>
  </xsl:template>
  
  <xsl:template name="InsertShapeWrap">
    <xsl:if test="descendant::node()[contains(name(),'wrap')]">
      <xsl:variable name="wrap" select="descendant::w10:wrap/@type"/>
        <xsl:attribute name="style:wrap">
          <xsl:choose>
            <xsl:when test="$wrap= 'square' ">
              <xsl:text>parallel</xsl:text>
            </xsl:when>
          </xsl:choose>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>
</xsl:stylesheet>

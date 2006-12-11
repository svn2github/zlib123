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
  xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  exclude-result-prefixes="w draw">

  <xsl:template match="w:pict|w:object">
      <xsl:choose>
        <xsl:when test="v:rect">
          <xsl:call-template name="v:rect"></xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <draw:frame>
          <xsl:call-template name="InsertShapeStyleName"/>
          <xsl:call-template name="InsertShapeAnchor"/>
          <xsl:call-template name="InsertShapeWidth"/>
          <xsl:call-template name="InsertShapeHeight"/>
          <xsl:call-template name="InsertShapeZindex"/>
          <xsl:call-template name="InsertshapeAbsolutePos"/>
          <xsl:apply-templates select="o:OLEObject"/>
          <xsl:apply-templates select="v:shape/child::node()"/>
          </draw:frame>
        </xsl:otherwise>
      </xsl:choose>
  </xsl:template>

  <xsl:template match="v:textbox">
    <draw:text-box>
      <xsl:call-template name="InsertTextBoxAutomaticHeight"/>
      <xsl:apply-templates select="w:txbxContent/child::node()"/>
    </draw:text-box>
  </xsl:template>

  <xsl:template match="o:OLEObject">
    <xsl:variable name="IdFile">
      <xsl:value-of select="@r:id"/>
    </xsl:variable>
    <xsl:variable name="IdImage">
      <xsl:value-of select="parent::w:object/v:shape/v:imagedata/@r:id"/>
    </xsl:variable>
    <xsl:variable name="pziptarget">
      <xsl:value-of select="generate-id()"/>
    </xsl:variable>
      <xsl:if test="document('word/_rels/document.xml.rels')">
        <xsl:for-each
          select="document('word/_rels/document.xml.rels')//node()[name() = 'Relationship']">
          <xsl:if test="./@Id=$IdFile">
            <xsl:variable name="pzipsource">
              <xsl:value-of select="./@Target"/>
            </xsl:variable>
            <pzip:copy pzip:source="{concat('word/',$pzipsource)}" pzip:target="{$pziptarget}"/>
          </xsl:if>
          <xsl:if test="./@Id=$IdImage">
            <xsl:variable name="pzipsource">
              <xsl:value-of select="./@Target"/>
            </xsl:variable>
            <pzip:copy pzip:source="{concat('word/',$pzipsource)}"
              pzip:target="{concat('ObjectReplacements/',$pziptarget)}"/>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    <draw:object-ole>
      <xsl:if test="document('word/_rels/document.xml.rels')">
        <xsl:for-each
          select="document('word/_rels/document.xml.rels')//node()[name() = 'Relationship']">
          <xsl:if test="./@Id=$IdFile">
            <xsl:attribute name="xlink:href">
              <xsl:choose>
                <xsl:when test="starts-with(./@Target, 'file:///')">
                  <xsl:value-of select="translate(substring-after(./@Target, 'file://'), '\', '/')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="concat('./',$pziptarget)"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
      <xsl:attribute name="xlink:show">embed</xsl:attribute>
    </draw:object-ole>
    <draw:image>
      <xsl:attribute name="xlink:href">
        <xsl:value-of select="concat('./ObjectReplacements/',$pziptarget)"/>
      </xsl:attribute>
    </draw:image>
  </xsl:template>
  
  <xsl:template match="v:imagedata">
    
    <xsl:variable name="rId">
      <xsl:value-of select="./@r:id"/>
    </xsl:variable>
    
    <xsl:variable name="FileName">
      <xsl:value-of select="generate-id()"/>
    </xsl:variable>
    
    <xsl:if test="document('word/_rels/document.xml.rels')">
      <xsl:for-each
        select="document('word/_rels/document.xml.rels')//node()[name() = 'Relationship']">
        <xsl:if test="./@Id=$rId">
          <xsl:variable name="pzipsource">
            <xsl:value-of select="./@Target"/>
          </xsl:variable>
          <pzip:copy pzip:source="{concat('word/',$pzipsource)}" pzip:target="{concat('Pictures/',$FileName,'.jpg')}"/>
        </xsl:if>       
      </xsl:for-each>
    </xsl:if>
    
    <draw:image>
      <xsl:attribute name="xlink:href">
        <xsl:value-of select="concat('Pictures/',$FileName, '.jpg')"/>
      </xsl:attribute>
    </draw:image>
    
  </xsl:template>

  <xsl:template name="InsertshapeAbsolutePos">
    <xsl:param name="shape" select="v:shape"/>
    
    <xsl:variable name="position">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$shape"/>
        <xsl:with-param name="propertyName" select="'position'"/>
      </xsl:call-template>
    </xsl:variable>
    
    <xsl:if test="$position = 'absolute'  or v:rect">
        <xsl:variable name="svgx">
          <xsl:call-template name="GetShapeProperty">
            <xsl:with-param name="shape" select="$shape"/>
            <xsl:with-param name="propertyName" select="'margin-left'"/>
          </xsl:call-template>
        </xsl:variable>
      
      <xsl:attribute name="svg:x">
        <xsl:call-template name="ConvertMeasure">
          <xsl:with-param name="length" select="$svgx"/>
          <xsl:with-param name="destUnit" select="'cm'"/>
        </xsl:call-template>
      </xsl:attribute>
      
      <xsl:variable name="svgy">
        <xsl:call-template name="GetShapeProperty">
          <xsl:with-param name="shape" select="$shape"/>
          <xsl:with-param name="propertyName" select="'margin-top'"/>
        </xsl:call-template>
      </xsl:variable>
      
      <xsl:attribute name="svg:y">
        <xsl:call-template name="ConvertMeasure">
          <xsl:with-param name="length" select="$svgy"/>
          <xsl:with-param name="destUnit" select="'cm'"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
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

    <xsl:variable name="relativeHeight">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$shape"/>
        <xsl:with-param name="propertyName" select="'mso-height-percent'"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="$relativeHeight != ''">
        <xsl:call-template name="InsertShapeRelativeHeight">
          <xsl:with-param name="shape" select="$shape"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="not($shape/v:textbox/@style)">
          <xsl:attribute name="svg:height">
            <xsl:call-template name="ConvertMeasure">
              <xsl:with-param name="length" select="$height"/>
              <xsl:with-param name="destUnit" select="'cm'"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertShapeRelativeHeight">
    <xsl:param name="shape" select="v:shape"/>

    <xsl:variable name="relativeHeight">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$shape"/>
        <xsl:with-param name="propertyName" select="'mso-height-percent'"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:attribute name="style:rel-height">
      <xsl:value-of select="$relativeHeight div 10"/>
    </xsl:attribute>
  </xsl:template>

  <xsl:template name="InsertShapeWidth">
    <xsl:param name="shape" select="v:shape"/>
    <xsl:variable name="wrapStyle">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$shape"/>
        <xsl:with-param name="propertyName" select="'mso-wrap-style'"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="relativeWidth">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$shape"/>
        <xsl:with-param name="propertyName" select="'mso-width-percent'"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="$relativeWidth != ''">
        <xsl:call-template name="InsertShapeRelativeWidth">
          <xsl:with-param name="shape" select="$shape"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$wrapStyle = '' or $wrapStyle != 'none' ">
        <xsl:variable name="width">
          <xsl:call-template name="GetShapeProperty">
            <xsl:with-param name="shape" select="$shape"/>
            <xsl:with-param name="propertyName" select="'width'"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:attribute name="svg:width">
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="$width"/>
            <xsl:with-param name="destUnit" select="'cm'"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertShapeRelativeWidth">
    <xsl:param name="shape" select="v:shape"/>

    <xsl:variable name="relativeWidth">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$shape"/>
        <xsl:with-param name="propertyName" select="'mso-width-percent'"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:attribute name="style:rel-width">
      <xsl:value-of select="$relativeWidth div 10"/>
    </xsl:attribute>
  </xsl:template>

  <xsl:template name="GetShapeProperty">
    <xsl:param name="shape"/>
    <xsl:param name="propertyName"/>
    <xsl:variable name="propertyValue">
      <xsl:choose>
        <xsl:when test="contains(substring-after($shape/@style,concat($propertyName,':')),';')">
          <xsl:value-of
            select="substring-before(substring-after($shape/@style,concat($propertyName,':')),';')"
          />
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="substring-after($shape/@style,concat($propertyName,':'))"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
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
    <style:style>
      <xsl:attribute name="style:name">
        <xsl:choose>
          <xsl:when test="v:rect">
            <xsl:call-template name="GenerateStyleName">          
              <xsl:with-param name="node" select="v:rect"/>
            </xsl:call-template>
       </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="GenerateStyleName">
          <xsl:with-param name="node" select="v:shape"/>
        </xsl:call-template>
      </xsl:otherwise>
        </xsl:choose>
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
    
    <xsl:apply-templates mode="automaticstyles"/>
  </xsl:template>

  <xsl:template name="InsertShapeProperties">
    <!--TODO automatic style content-->
    <xsl:choose>
      <xsl:when test="v:rect">
        <xsl:call-template name="InsertShapeWrap"/>
        <xsl:call-template name="InsertShapeFromTextDistance">
          <xsl:with-param name="shape" select="v:rect"></xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="InsertShapeBorders">
          <xsl:with-param name="shape" select="v:rect"></xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="InsertShapeAutomaticWidth">
          <xsl:with-param name="shape" select="v:rect"></xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="InsertShapeHorizontalPos">
          <xsl:with-param name="shape" select="v:rect"></xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="InsertShapeVerticalPos">
          <xsl:with-param name="shape" select="v:rect"></xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="InsertShapeFlowWithText">
          <xsl:with-param name="shape" select="v:rect"></xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="InsertShapeBackgroundColor">
          <xsl:with-param name="shape" select="v:rect"></xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="InsertShapeWrap"/>
        <xsl:call-template name="InsertShapeFromTextDistance"/>
        <xsl:call-template name="InsertTextBoxPadding"/>
        <xsl:call-template name="InsertShapeBorders"/>
        <xsl:call-template name="InsertShapeAutomaticWidth"/>
        <xsl:call-template name="InsertShapeHorizontalPos"/>
        <xsl:call-template name="InsertShapeVerticalPos"/>
        <xsl:call-template name="InsertShapeFlowWithText"/>
        <xsl:call-template name="InsertShapeBackgroundColor"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="InsertShapeBackgroundColor">
    <xsl:param name="shape" select="v:shape"/>
    <xsl:variable name="bgColor" select="$shape/@fillcolor"/>
     <xsl:variable name="isFilled" select="$shape/@filled"/>
      <xsl:if test="(not($isFilled) or $isFilled != 'f') and $bgColor != ''">
      <xsl:attribute name="fo:background-color">
        <xsl:call-template name="InsertColor">
            <xsl:with-param name="color" select="$bgColor"/>
        </xsl:call-template>
      </xsl:attribute>
      </xsl:if>
    <xsl:if test="v:rect">
      <xsl:attribute name="draw:fill-color">
        <xsl:call-template name="InsertColor">
          <xsl:with-param name="color" select="$bgColor"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>
  
  <xsl:template name="InsertShapeFlowWithText">
  <!--    TODO  investigate when this need to be set to true -->
    <xsl:attribute name="draw:flow-with-text">
        <xsl:text>false</xsl:text>
    </xsl:attribute>
  </xsl:template>
  
  <xsl:template name="InsertShapeHorizontalPos">
    <xsl:param name="shape" select="v:shape"/>
    
    <xsl:variable name="horizontalPos">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="propertyName" select="'mso-position-horizontal'"/>
        <xsl:with-param name="shape" select="v:shape"/>
      </xsl:call-template>
    </xsl:variable>
    
    <xsl:choose>     
      <xsl:when test="v:rect/@o:hralign and v:rect/@o:hr='t'">
        <xsl:call-template name="InsertGraphicPosH">      
          <xsl:with-param name="align" select="v:rect/@o:hralign"/>
        </xsl:call-template>
      </xsl:when>    
      <xsl:when test="$shape/@o:hralign">
        <xsl:call-template name="InsertGraphicPosH">      
          <xsl:with-param name="align" select="$shape/@o:hralign"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
          <xsl:call-template name="InsertGraphicPosH">
            <xsl:with-param name="align" select="$horizontalPos"/>
          </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>

    <xsl:variable name="horizontalRelative">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="propertyName" select="'mso-position-horizontal-relative'"/>
        <xsl:with-param name="shape" select="v:shape"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:call-template name="InsertGraphicPosRelativeH">
      <xsl:with-param name="relativeFrom" select="$horizontalRelative"/>
    </xsl:call-template>
  </xsl:template>


  <xsl:template name="InsertShapeVerticalPos">
    <xsl:param name="shape" select="v:shape"/>

    <xsl:variable name="verticalPos">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="propertyName" select="'mso-position-vertical'"/>
        <xsl:with-param name="shape" select="v:shape"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="verticalRelative">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="propertyName" select="'mso-position-vertical-relative'"/>
        <xsl:with-param name="shape" select="v:shape"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:call-template name="InsertGraphicPosV">
      <xsl:with-param name="align" select="$verticalPos"/>
      <xsl:with-param name="relativeFrom" select="$verticalRelative"/>
    </xsl:call-template>

    <xsl:call-template name="InsertGraphicPosRelativeV">
      <xsl:with-param name="relativeFrom" select="$verticalRelative"/>
      <xsl:with-param name="align" select="$verticalPos"/>
    </xsl:call-template>
  </xsl:template>

  <xsl:template name="InsertShapeWrap">
    <xsl:if test="descendant::node()[contains(name(),'wrap')]">
      <xsl:variable name="wrap" select="descendant::w10:wrap"/>
      <xsl:variable name="zindex">
        <xsl:call-template name="GetShapeProperty">
          <xsl:with-param name="propertyName" select="'z-index'"/>
          <xsl:with-param name="shape" select="v:shape"/>
        </xsl:call-template>
      </xsl:variable>

      <xsl:if test="$zindex &lt; 0 and not($wrap/@type)">
        <xsl:attribute name="style:run-through">
          <xsl:text>background</xsl:text>
        </xsl:attribute>
      </xsl:if>

      <xsl:attribute name="style:wrap">
        <xsl:choose>
          <xsl:when test="not($wrap/@type)">
            <xsl:text>run-through</xsl:text>
          </xsl:when>
          <xsl:when
            test="($wrap/@type = 'square' or $wrap/@type = 'tight' or $wrap/@type = 'through') and not($wrap/@side)">
            <xsl:text>parallel</xsl:text>
          </xsl:when>
          <xsl:when
            test="($wrap/@type = 'square' or $wrap/@type = 'tight' or $wrap/@type = 'through') and $wrap/@side = 'left' ">
            <xsl:text>left</xsl:text>
          </xsl:when>
          <xsl:when
            test="($wrap/@type = 'square' or $wrap/@type = 'tight' or $wrap/@type = 'through') and $wrap/@side = 'right' ">
            <xsl:text>right</xsl:text>
          </xsl:when>
          <xsl:when
            test="($wrap/@type = 'square' or $wrap/@type = 'tight' or $wrap/@type = 'through') and $wrap/@side = 'largest' ">
            <xsl:text>dynamic</xsl:text>
          </xsl:when>
          <xsl:when test="$wrap/@type = 'topAndBottom' ">
            <xsl:text>none</xsl:text>
          </xsl:when>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertShapeFromTextDistance">
    <xsl:param name="shape" select="v:shape"/>

    <xsl:variable name="margin-top">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="propertyName" select="'mso-wrap-distance-top'"/>
        <xsl:with-param name="shape" select="$shape"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$margin-top !=''">
      <xsl:call-template name="InsertShapeMargin">
        <xsl:with-param name="attributeName" select=" 'fo:margin-top' "/>
        <xsl:with-param name="margin" select="$margin-top"/>
      </xsl:call-template>
    </xsl:if>

    <xsl:variable name="margin-bottom">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="propertyName" select="'mso-wrap-distance-bottom'"/>
        <xsl:with-param name="shape" select="$shape"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$margin-bottom !=''">
      <xsl:call-template name="InsertShapeMargin">
        <xsl:with-param name="attributeName" select=" 'fo:margin-bottom' "/>
        <xsl:with-param name="margin" select="$margin-bottom"/>
      </xsl:call-template>
    </xsl:if>

    <xsl:variable name="margin-left">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="propertyName" select="'mso-wrap-distance-left'"/>
        <xsl:with-param name="shape" select="$shape"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$margin-left !=''">
      <xsl:call-template name="InsertShapeMargin">
        <xsl:with-param name="attributeName" select=" 'fo:margin-left' "/>
        <xsl:with-param name="margin" select="$margin-left"/>
      </xsl:call-template>
    </xsl:if>

    <xsl:variable name="margin-right">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="propertyName" select="'mso-wrap-distance-right'"/>
        <xsl:with-param name="shape" select="$shape"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$margin-right !=''">
      <xsl:call-template name="InsertShapeMargin">
        <xsl:with-param name="attributeName" select=" 'fo:margin-right' "/>
        <xsl:with-param name="margin" select="$margin-right"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertShapeMargin">
    <xsl:param name="margin"/>
    <xsl:param name="attributeName"/>

    <xsl:attribute name="{$attributeName}">
      <xsl:call-template name="ConvertMeasure">
        <xsl:with-param name="length" select="$margin"/>
        <xsl:with-param name="destUnit" select="'cm'"/>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>

  <xsl:template name="InsertTextBoxPadding">
    <xsl:param name="shape" select="v:shape"/>

    <xsl:variable name="textBoxInset" select="$shape/v:textbox/@inset"/>

    <xsl:variable name="paddingLeft">
      <xsl:call-template name="ConvertMeasure">
        <xsl:with-param name="length">
          <xsl:call-template name="ExplodeValues">
            <xsl:with-param name="elementNum" select="'1'"/>
            <xsl:with-param name="text" select="$textBoxInset"/>
          </xsl:call-template>
        </xsl:with-param>
        <xsl:with-param name="destUnit" select="'cm'"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:attribute name="fo:padding-left">
      <xsl:value-of select="$paddingLeft"/>
    </xsl:attribute>

    <xsl:variable name="paddingTop">
      <xsl:call-template name="ConvertMeasure">
        <xsl:with-param name="length">
          <xsl:call-template name="ExplodeValues">
            <xsl:with-param name="elementNum" select="'2'"/>
            <xsl:with-param name="text" select="$textBoxInset"/>
          </xsl:call-template>
        </xsl:with-param>
        <xsl:with-param name="destUnit" select="'cm'"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:attribute name="fo:padding-top">
      <xsl:value-of select="$paddingTop"/>
    </xsl:attribute>

    <xsl:variable name="paddingRight">
      <xsl:call-template name="ConvertMeasure">
        <xsl:with-param name="length">
          <xsl:call-template name="ExplodeValues">
            <xsl:with-param name="elementNum" select="'3'"/>
            <xsl:with-param name="text" select="$textBoxInset"/>
          </xsl:call-template>
        </xsl:with-param>
        <xsl:with-param name="destUnit" select="'cm'"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:attribute name="fo:padding-right">
      <xsl:value-of select="$paddingRight"/>
    </xsl:attribute>

    <xsl:variable name="paddingBottom">
      <xsl:call-template name="ConvertMeasure">
        <xsl:with-param name="length">
          <xsl:call-template name="ExplodeValues">
            <xsl:with-param name="elementNum" select="'4'"/>
            <xsl:with-param name="text" select="$textBoxInset"/>
          </xsl:call-template>
        </xsl:with-param>
        <xsl:with-param name="destUnit" select="'cm'"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:attribute name="fo:padding-bottom">
      <xsl:value-of select="$paddingBottom"/>
    </xsl:attribute>
  </xsl:template>

  <xsl:template name="ExplodeValues">
    <xsl:param name="elementNum"/>
    <xsl:param name="text"/>
    <xsl:param name="counter" select="'1'"/>
    <xsl:choose>
      <xsl:when test="$elementNum = $counter">
        <xsl:choose>
          <xsl:when test="contains($text,',')">
            <xsl:value-of select="substring-before($text,',')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$text"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="ExplodeValues">
          <xsl:with-param name="elementNum" select="$elementNum"/>
          <xsl:with-param name="text" select="substring-after($text,',')"/>
          <xsl:with-param name="counter" select="$counter+1"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertShapeBorders">
    <xsl:param name="shape" select="v:shape"/>

    <xsl:choose>
      <xsl:when test="v:rect/@o:hr='t'">
        <xsl:attribute name="draw:stroke">
          <xsl:text>none</xsl:text>
        </xsl:attribute>
      </xsl:when>      
      <xsl:otherwise>
        <xsl:if test="$shape/@strokeweight">
          <xsl:variable name="borderWeight">
            <xsl:call-template name="ConvertMeasure">
              <xsl:with-param name="length" select="$shape/@strokeweight"/>
              <xsl:with-param name="destUnit" select="'cm'"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="borderColor">
            <xsl:call-template name="InsertColor">
              <xsl:with-param name="color" select="$shape/@strokecolor"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="borderStyle">
            <xsl:choose>
              <xsl:when test="not($shape/v:stroke)">
                <xsl:text>solid</xsl:text>
              </xsl:when>
              <xsl:otherwise>
                <xsl:text>double</xsl:text>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          
          <xsl:attribute name="fo:border">
            <xsl:value-of select="concat($borderWeight,' ',$borderStyle,' ',$borderColor)"/>
          </xsl:attribute>
          
          <xsl:if test="$borderStyle= 'double' ">
            <xsl:variable name="strokeStyle" select="$shape/v:stroke/@linestyle"/>
            
            <xsl:attribute name="style:border-line-width">
              <xsl:choose>
                <xsl:when test="$strokeStyle = 'thinThin' or $strokeStyle = 'thickBetweenThin'">
                  <xsl:value-of
                    select="concat(substring-before($borderWeight,'cm')*0.45 ,'cm',' ',substring-before($borderWeight,'cm')*0.1,'cm ', substring-before($borderWeight,'cm')*0.45,'cm')"
                  />
                </xsl:when>
                <xsl:when test="$strokeStyle = 'thinThick' ">
                  <xsl:value-of
                    select="concat(substring-before($borderWeight,'cm')*0.7,'cm',' ',substring-before($borderWeight,'cm')*0.1,'cm ', substring-before($borderWeight,'cm')*0.2,'cm')"
                  />
                </xsl:when>
                <xsl:when test="$strokeStyle = 'thickThin' ">
                  <xsl:value-of
                    select="concat(substring-before($borderWeight,'cm')*0.2,'cm',' ',substring-before($borderWeight,'cm')*0.1,'cm ', substring-before($borderWeight,'cm')*0.7,'cm')"
                  />
                </xsl:when>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>    
   
  </xsl:template>

  <xsl:template name="InsertTextBoxAutomaticHeight">
    <xsl:if test="@style ='mso-fit-shape-to-text:t'">
      <xsl:attribute name="fo:min-height">
        <xsl:text>0cm</xsl:text>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertShapeAutomaticWidth">
    <xsl:param name="shape" select="v:shape"/>
    <xsl:variable name="wrapStyle">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$shape"/>
        <xsl:with-param name="propertyName" select="'mso-wrap-style'"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$wrapStyle != '' and $wrapStyle = 'none' ">
        <xsl:attribute name="fo:min-width">
          <xsl:text>0cm</xsl:text>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="v:rect">
    <draw:rect>     
      <xsl:call-template name="InsertShapeAnchor">
        <xsl:with-param name="shape" select="v:rect"/>
      </xsl:call-template>  
      <xsl:call-template name="InsertShapeZindex">
        <xsl:with-param name="shape" select="v:rect"/>
      </xsl:call-template>
      <xsl:call-template name="InsertShapeStyleName">
        <xsl:with-param name="shape" select="v:rect"/>
      </xsl:call-template>
      <xsl:call-template name="InsertShapeWidth">
        <xsl:with-param name="shape" select="v:rect"/>
      </xsl:call-template>
      <xsl:call-template name="InsertShapeHeight">
        <xsl:with-param name="shape" select="v:rect"/>
      </xsl:call-template>
      <xsl:call-template name="InsertshapeAbsolutePos">
        <xsl:with-param name="shape" select="v:rect"/>
      </xsl:call-template>      
      <text:p/>
    </draw:rect>
    <!--draw:rect text:anchor-type="paragraph" draw:z-index="0" draw:style-name="gr1"
      svg:width="10.408cm" svg:height="5.469cm" svg:x="2.438cm"
      svg:y="1.288cm">
      <text:p/>
    </draw:rect-->
  </xsl:template>

  
</xsl:stylesheet>

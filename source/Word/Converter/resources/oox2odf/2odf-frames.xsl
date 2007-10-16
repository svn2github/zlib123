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
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  exclude-result-prefixes="w r draw number wp xlink v w10 o">

  <!-- shape or ole-object-->
  <xsl:template match="w:pict | w:object">
    <xsl:choose>
      <xsl:when test="v:group">
        <draw:g>
          <xsl:attribute name="draw:style-name">
            <xsl:value-of select="generate-id(.)"/>
          </xsl:attribute>
          <xsl:apply-templates/>
        </draw:g>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="v:shape">
      <draw:frame>
        <xsl:call-template name="InsertCommonShapeProperties"/>
        <xsl:call-template name="InsertShapeZindex"/>
        <xsl:apply-templates/>
        <!--  some of the shape types must be in odf draw:frame even if they are outside of v:shape in oox-->
        <xsl:apply-templates select="self::node()/following-sibling::node()[1]" mode="draw:frame"/>
      </draw:frame>
  </xsl:template>

  <xsl:template match="o:extrusion">
    <xsl:message terminate="no">translation.oox2odf.shape.3dEffect</xsl:message>
  </xsl:template>

  <!--horizontal line-->
  <xsl:template match="v:rect">
    <!--    version 1.1-->
    <xsl:choose>
      <xsl:when test="v:imagedata">
        <xsl:variable name="document">
          <xsl:call-template name="GetDocumentName">
            <xsl:with-param name="rootId">
              <xsl:value-of select="generate-id(/node())"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        
        <xsl:call-template name="CopyPictures">
          <xsl:with-param name="document">
            <xsl:value-of select="$document"/>
          </xsl:with-param>
          <xsl:with-param name="rId">
            <xsl:value-of select="v:imagedata/@r:id"/>
          </xsl:with-param>
        </xsl:call-template>
        
        <draw:frame>
          <xsl:attribute name="draw:style-name">
          <xsl:call-template name="GenerateStyleName">
            <xsl:with-param name="node" select="."/>
          </xsl:call-template>
        </xsl:attribute>
          <!--TODO-->
          <xsl:attribute name="draw:name">
            <xsl:text>Frame1</xsl:text>
          </xsl:attribute>
          <xsl:call-template name="InsertShapeAnchor"/>
          <xsl:call-template name="InsertShapeWidth"/>
          <xsl:call-template name="InsertShapeHeight"/>
          <xsl:variable name="rotation">
            <xsl:call-template name="GetShapeProperty">
              <xsl:with-param name="shape" select="."/>
              <xsl:with-param name="propertyName" select="'rotation'"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="left2">
            <xsl:call-template name="GetShapeProperty">
              <xsl:with-param name="shape" select="."/>
              <xsl:with-param name="propertyName" select="'left'"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:call-template name="GetShapeProperty">
              <xsl:with-param name="shape" select="."/>
              <xsl:with-param name="propertyName" select="'width'"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="left">
            <xsl:call-template name="ConvertMeasure">
              <xsl:with-param name="length">
                <xsl:value-of select="$left2+$width"/>
              </xsl:with-param>
              <xsl:with-param name="destUnit" select="'cm'"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:call-template name="GetShapeProperty">
              <xsl:with-param name="shape" select="."/>
              <xsl:with-param name="propertyName" select="'height'"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="top2">
            <xsl:call-template name="GetShapeProperty">
              <xsl:with-param name="shape" select="."/>
              <xsl:with-param name="propertyName" select="'top'"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="top">
            <xsl:call-template name="ConvertMeasure">
              <xsl:with-param name="length">
                  <xsl:value-of select="$top2+$height"/>            
              </xsl:with-param>
              <xsl:with-param name="destUnit" select="'cm'"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="leftcm">
            <xsl:call-template name="ConvertMeasure">
              <xsl:with-param name="length">
                <xsl:value-of select="$left2"/>            
              </xsl:with-param>
              <xsl:with-param name="destUnit" select="'cm'"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="topcm">
            <xsl:call-template name="ConvertMeasure">
              <xsl:with-param name="length">
                <xsl:value-of select="$top2"/>            
              </xsl:with-param>
              <xsl:with-param name="destUnit" select="'cm'"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test="$rotation != ''">
            <xsl:attribute name="draw:transform">
              <xsl:choose>
                <xsl:when test="$rotation='270'">
                  <xsl:value-of select="concat('rotate (1.5707963267946) translate(',$leftcm,' ',$top,')')"/>
                </xsl:when>
                <xsl:when test="$rotation='180'">
                  <xsl:value-of select="concat('rotate (3.1415926535892) translate(',$left,' ',$top,')')"/>
                </xsl:when>
                <xsl:when test="$rotation='90'">
                  <xsl:value-of select="concat('rotate (-1.57079632679579) translate(',$left,' ',$topcm,')')"/>
                </xsl:when>
              </xsl:choose>
            </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="InsertshapeAbsolutePos"/>
            </xsl:otherwise>
            </xsl:choose>
          <!-- image href from relationships-->
          <draw:image xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad" xlink:href="">
            <xsl:if test="document(concat('word/_rels/',$document,'.rels'))">
              <xsl:call-template name="InsertImageHref">
                <xsl:with-param name="document" select="$document"/>
                <xsl:with-param name="rId">
                  <xsl:value-of select="v:imagedata/@r:id"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:if>
          </draw:image>
        </draw:frame>
      </xsl:when>
      <xsl:otherwise>
      <draw:rect>
        <xsl:attribute name="draw:style-name">
          <xsl:call-template name="GenerateStyleName">
            <xsl:with-param name="node" select="."/>
          </xsl:call-template>
        </xsl:attribute>
        <!--TODO-->
        <xsl:attribute name="draw:name">
          <xsl:text>Frame1</xsl:text>
        </xsl:attribute>
        <xsl:call-template name="InsertShapeAnchor"/>
        <xsl:call-template name="InsertShapeWidth"/>
        <xsl:call-template name="InsertShapeHeight"/>
        <xsl:call-template name="InsertshapeAbsolutePos"/>
      </draw:rect>
        </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertCommonShapeProperties">
    <xsl:call-template name="InsertShapeStyleName"/>
    <xsl:call-template name="InsertShapeAnchor"/>
    <xsl:call-template name="InsertShapeWidth"/>
    <xsl:call-template name="InsertShapeHeight"/>
    <xsl:call-template name="InsertshapeAbsolutePos"/>
  </xsl:template>

  <!--hyperlink shape-->
  <xsl:template match="w:pict[v:shape/@href]">
    <draw:a xlink:type="simple" xlink:href="{v:shape/@href}">
      <xsl:apply-templates/>
    </draw:a>
  </xsl:template>

  <!-- text watermark feedback -->
  <xsl:template match="w:pict[contains(v:shape/@id,'WaterMark')]">
    <xsl:if test="v:shape/v:textpath">
      <xsl:message terminate="no">translation.oox2odf.background.textWatermark</xsl:message>
    </xsl:if>
  </xsl:template>

  <xsl:template match="v:textbox">
    <xsl:if test="parent::v:stroke/@dashstyle">
      <xsl:message terminate="no">translation.oox2odf.textbox.boder.dashed</xsl:message>
    </xsl:if>
    <draw:text-box>
      <xsl:call-template name="InsertTextBoxAutomaticHeight"/>
      <xsl:apply-templates select="w:txbxContent/child::node()"/>
    </draw:text-box>
  </xsl:template>


  <xsl:template match="o:OLEObject" mode="draw:frame">
    <xsl:variable name="document">
      <xsl:call-template name="GetDocumentName">
        <xsl:with-param name="rootId">
          <xsl:value-of select="generate-id(/node())"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:call-template name="CopyPictures">
      <xsl:with-param name="document">
        <xsl:value-of select="$document"/>
      </xsl:with-param>
      <xsl:with-param name="rId" select="@r:id"/>
      <xsl:with-param name="destFolder" select="'.'"/>
      <xsl:with-param name="targetName" select="generate-id(.)"/>
    </xsl:call-template>

    <xsl:call-template name="CopyPictures">
      <xsl:with-param name="document">
        <xsl:value-of select="$document"/>
      </xsl:with-param>
      <xsl:with-param name="rId" select="parent::w:object/v:shape/v:imagedata/@r:id"/>
      <xsl:with-param name="destFolder" select="'ObjectReplacements'"/>
      <xsl:with-param name="targetName" select="generate-id(.)"/>
    </xsl:call-template>

    <draw:object-ole>
      <xsl:if test="document(concat('word/_rels/',$document,'.rels'))">
        <xsl:call-template name="InsertImageHref">
          <xsl:with-param name="document" select="$document"/>
          <xsl:with-param name="rId" select="@r:id"/>
          <xsl:with-param name="srcFolder" select="'.'"/>
          <xsl:with-param name="targetName" select="generate-id(.)"/>
        </xsl:call-template>
      </xsl:if>
      <xsl:attribute name="xlink:show">embed</xsl:attribute>
    </draw:object-ole>
    <draw:image>
      <xsl:call-template name="InsertImageHref">
        <xsl:with-param name="document" select="$document"/>
        <xsl:with-param name="rId"
          select="parent::w:object/v:shape/v:imagedata/@r:id | parent::w:pict/v:shape/v:imagedata/@r:id"/>
        <xsl:with-param name="srcFolder" select="'./ObjectReplacements'"/>
        <xsl:with-param name="targetName" select="generate-id(.)"/>
      </xsl:call-template>
    </draw:image>
  </xsl:template>

  <!-- inserts horizontal ruler as image -->
  <xsl:template match="v:imagedata[not(ancestor::w:object)]">
    <xsl:variable name="document">
      <xsl:call-template name="GetDocumentName">
        <xsl:with-param name="rootId">
          <xsl:value-of select="generate-id(/node())"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:call-template name="CopyPictures">
      <xsl:with-param name="document">
        <xsl:value-of select="$document"/>
      </xsl:with-param>
      <xsl:with-param name="rId" select="@r:id"/>
    </xsl:call-template>
    <draw:image xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad">
      <xsl:if test="document(concat('word/_rels/',$document,'.rels'))">
        <xsl:call-template name="InsertImageHref">
          <xsl:with-param name="document" select="$document"/>
          <xsl:with-param name="rId" select="@r:id"/>
        </xsl:call-template>
      </xsl:if>
    </draw:image>
  </xsl:template>

  <xsl:template name="InsertshapeAbsolutePos">
    <xsl:param name="shape" select="."/>

    <xsl:variable name="position">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$shape"/>
        <xsl:with-param name="propertyName" select="'position'"/>
      </xsl:call-template>
    </xsl:variable>
    
    <xsl:variable name="LeftPercent">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$shape"/>
        <xsl:with-param name="propertyName" select="'mso-left-percent'"/>
      </xsl:call-template>
    </xsl:variable>
    
    <xsl:variable name="HorizontalWidth">
      <xsl:variable name="HorizontalRelative">
        <xsl:call-template name="GetShapeProperty">
          <xsl:with-param name="shape" select="$shape"/>
          <xsl:with-param name="propertyName" select="'mso-position-horizontal-relative'"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$HorizontalRelative = 'page'">
          <xsl:value-of select="number(ancestor::node()/w:sectPr/w:pgSz/@w:w)"/>
        </xsl:when>
        <xsl:when test="$HorizontalRelative = 'margin'">
          <xsl:value-of select="number(ancestor::node()/w:sectPr/w:pgSz/@w:w) - number(ancestor::node()/w:sectPr/w:pgMar/@w:left) - number(ancestor::node()/w:sectPr/w:pgMar/@w:right)"/>
        </xsl:when>
        
      </xsl:choose>
    </xsl:variable>

    <xsl:if test="$position = 'absolute' or $shape[name()='w:framePr']">
      <xsl:variable name="svgx">
        <xsl:choose>
          <xsl:when test="$shape[name()='w:framePr']">
            <xsl:value-of select="$shape/@w:x"/>
          </xsl:when>
          <xsl:when test="$LeftPercent  &gt; 0">
            <xsl:value-of select="$HorizontalWidth * $LeftPercent  div 1000"/>
          </xsl:when>
          <xsl:when test="$shape[name()='v:rect']">
            <xsl:call-template name="GetShapeProperty">
              <xsl:with-param name="shape" select="$shape"/>
              <xsl:with-param name="propertyName" select="'left'"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="GetShapeProperty">
              <xsl:with-param name="shape" select="$shape"/>
              <xsl:with-param name="propertyName" select="'margin-left'"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>

      <xsl:attribute name="svg:x">
        <xsl:call-template name="ConvertMeasure">
          <xsl:with-param name="length" select="$svgx"/>
          <xsl:with-param name="destUnit" select="'cm'"/>
        </xsl:call-template>
      </xsl:attribute>

      <xsl:variable name="svgy">
        <xsl:choose>
          <xsl:when test="$shape[name()='w:framePr']">
            <xsl:value-of select="$shape/@w:y"/>
          </xsl:when>
          <xsl:when test="$shape[name()='v:rect']">
            <xsl:call-template name="GetShapeProperty">
              <xsl:with-param name="shape" select="$shape"/>
              <xsl:with-param name="propertyName" select="'top'"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="GetShapeProperty">
              <xsl:with-param name="shape" select="$shape"/>
              <xsl:with-param name="propertyName" select="'margin-top'"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
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
    <xsl:param name="shape" select="."/>
    <xsl:attribute name="text:anchor-type">
      <xsl:choose>
        <xsl:when test="w10:wrap/@type = 'none' ">
          <xsl:text>as-char</xsl:text>
        </xsl:when>
        <!-- In header of footer, frames that are not in background must be anchored as character because otherwise they lose their size
       and horizontal and vertical position (when anchored as character horizontal position is lost)-->
        <xsl:when
          test="(ancestor::w:hdr or ancestor::w:ftr) and(w10:wrap/@type != '' or $shape/@w:wrap != '')">
          <xsl:if test="ancestor::w:hdr">
            <xsl:message terminate="no">translation.oox2odf.frame.inHeader</xsl:message>
          </xsl:if>
          <xsl:if test="ancestor::w:ftr">
            <xsl:message terminate="no">translation.oox2odf.frame.inFooter</xsl:message>
          </xsl:if>
          <xsl:text>as-char</xsl:text>
        </xsl:when>
        <!-- anchor should be 'as-char' in some particular cases. Explanation of test :
        1. if no wrap is defined (default in OOX is inline with text)
          AND
        2. if there is another run in current paragraph OR we are in a table cell
          AND
        3. No absolute position is defined OR wrapping is explicitely set to 'as-char' -->
        
        <!-- JP 27/08/2007 -->
        <!-- Fix #1666915 if position is absolute  must not be set as-char-->
        <xsl:when
          test="not(w10:wrap) and ( 
          count(ancestor::w:r[1]/parent::node()/w:r) &gt; 1
          or (not(count(ancestor::w:r[1]/parent::node()/w:r) &gt; 1) and ancestor::w:tc)
          ) 
          and (
          not(contains($shape/@style, 'position:absolute'))
          )
          and (
          contains($shape/@style, 'mso-position-horizontal-relative:text')
          or contains($shape/@style, 'mso-position-vertical-relative:text')
          or ($shape/@vAnchor='text') or ($shape/@hAnchor='relative') )">
          <xsl:text>as-char</xsl:text>
        </xsl:when>
        
        <xsl:when
          test="(w10:wrap/@anchorx = 'page' and w10:wrap/@anchory = 'page') or ($shape/@w:hAnchor='page' and $shape/@w:vAnchor='page')">
          <xsl:text>page</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:text>paragraph</xsl:text>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>

  <xsl:template name="InsertShapeZindex">
    <xsl:param name="shape" select="."/>
    <xsl:variable name="zindex">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$shape"/>
        <xsl:with-param name="propertyName" select="'z-index'"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:if test="$zindex != ''">
      <xsl:attribute name="draw:z-index">
        <xsl:value-of select="$zindex"/>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertShapeStyleName">
    <xsl:param name="shape" select="ancestor::w:pict | ancestor::w:object"/>
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
    <xsl:param name="shape" select="."/>

    <xsl:variable name="height">
      <xsl:choose>
        <xsl:when test="$shape[name()='w:framePr']">
          <xsl:value-of select="$shape/@w:h"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="GetShapeProperty">
            <xsl:with-param name="shape" select="$shape"/>
            <xsl:with-param name="propertyName" select="'height'"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="relativeHeight">
      <xsl:choose>
        <xsl:when test="$shape[name()='w:framePr']"/>
        <xsl:otherwise>
          <xsl:call-template name="GetShapeProperty">
            <xsl:with-param name="shape" select="$shape"/>
            <xsl:with-param name="propertyName" select="'mso-height-percent'"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="$relativeHeight != ''">
        <xsl:call-template name="InsertShapeRelativeHeight">
          <xsl:with-param name="shape" select="$shape"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if
          test="not(contains($shape/v:textbox/@style, 'mso-fit-shape-to-text:t'))  or $shape/@w:h">
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
    <xsl:param name="shape" select="."/>

    <xsl:variable name="relativeHeight">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$shape"/>
        <xsl:with-param name="propertyName" select="'mso-height-percent'"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:attribute name="style:rel-height">
      <xsl:value-of select="$relativeHeight div 10"/>
    </xsl:attribute>

    <xsl:variable name="relativeTo">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$shape"/>
        <xsl:with-param name="propertyName" select="'mso-height-relative'"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$relativeTo != ''">
      <xsl:message terminate="no">translation.oox2odf.frame.relativeSize</xsl:message>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertShapeWidth">
    <xsl:param name="shape" select="."/>
    <xsl:variable name="wrapStyle">
      <xsl:choose>
        <xsl:when test="$shape[name()='w:framePr']">
          <xsl:value-of select="$shape/@w:wrap"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="GetShapeProperty">
            <xsl:with-param name="shape" select="$shape"/>
            <xsl:with-param name="propertyName" select="'mso-wrap-style'"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
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
          <xsl:choose>
            <xsl:when test="$shape[name()='w:framePr']">
              <xsl:value-of select="$shape/@w:w"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="GetShapeProperty">
                <xsl:with-param name="shape" select="$shape"/>
                <xsl:with-param name="propertyName" select="'width'"/>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:attribute name="svg:width">
          <xsl:choose>
            <xsl:when test="$width = 0 and $shape//@o:hr='t'">
              <xsl:call-template name="ConvertTwips">
                <xsl:with-param name="length"
                  select="following::w:pgSz[1]/@w:w - following::w:pgMar/@w:right[1] - following::w:pgMar/@w:left[1]"/>
                <xsl:with-param name="unit" select="'cm'"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="ConvertMeasure">
                <xsl:with-param name="length" select="$width"/>
                <xsl:with-param name="destUnit" select="'cm'"/>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertShapeRelativeWidth">
    <xsl:param name="shape" select="."/>

    <xsl:variable name="relativeWidth">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$shape"/>
        <xsl:with-param name="propertyName" select="'mso-width-percent'"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:attribute name="style:rel-width">
      <xsl:value-of select="$relativeWidth div 10"/>
    </xsl:attribute>

    <xsl:variable name="relativeTo">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$shape"/>
        <xsl:with-param name="propertyName" select="'mso-width-relative'"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$relativeTo != ''">
      <xsl:message terminate="no">translation.oox2odf.frame.relativeSize</xsl:message>
    </xsl:if>
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

  <xsl:template match="w:pict | w:object" mode="automaticpict">
    <style:style>
      <xsl:attribute name="style:name">
        <xsl:call-template name="GenerateStyleName">
          <xsl:with-param name="node" select="self::node()"/>
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
        <xsl:call-template name="InsertShapeStyleProperties"/>
      </style:graphic-properties>
    </style:style>

    <xsl:apply-templates mode="automaticpict"/>
  </xsl:template>

  <xsl:template match="v:rect" mode="automaticpict">
    <style:style>
      <xsl:attribute name="style:name">
        <xsl:call-template name="GenerateStyleName">
          <xsl:with-param name="node" select="self::node()"/>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="style:parent-style-name">
        <xsl:text>Graphics</xsl:text>
      </xsl:attribute>
      <xsl:attribute name="style:family">
        <xsl:text>graphic</xsl:text>
      </xsl:attribute>
      
      <style:graphic-properties>
        <xsl:choose>
          <xsl:when test="v:fill/@type = 'gradient'">
            <xsl:attribute name="draw:fill">gradient</xsl:attribute>
            <xsl:attribute name="draw:fill-gradient-name">
              <xsl:value-of select="concat('Gradient_',generate-id(.))"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="v:imagedata">
            
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="InsertShapeBackgroundColor"/>
          </xsl:otherwise>
        </xsl:choose>
        <xsl:call-template name="InsertShapeShadow"/>
      </style:graphic-properties>
    </style:style>
    
    
    <xsl:apply-templates mode="automaticpict"/>
  </xsl:template>
  
  <xsl:template name="InsertGradientFill">
    <draw:gradient >
      <xsl:attribute name="draw:name">
        <xsl:value-of select="concat('Gradient_',generate-id(.))"/>
      </xsl:attribute>
      <xsl:attribute name="draw:style">linear</xsl:attribute>
      <xsl:attribute name="draw:start-color">
        <xsl:choose>
          <xsl:when test="v:fill/@color">
            <xsl:value-of select="v:fill/@color"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="./@fillcolor"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:attribute name="draw:end-color">
        <xsl:choose>
          <xsl:when test="v:fill/@color2">
            <xsl:value-of select="v:fill/@color2"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>#ffffff</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
        </xsl:attribute>
      <xsl:attribute name="draw:start-intensity">100%</xsl:attribute>
      <xsl:attribute name="draw:end-intensity">100%</xsl:attribute>
      <xsl:attribute name="draw:angle">0</xsl:attribute>
      <xsl:attribute name="draw:border">0%</xsl:attribute>
    </draw:gradient>
  </xsl:template>
  <!-- don't process text with automatic pict mode -->
  <xsl:template match="w:t" mode="automaticpict"/>

  <xsl:template name="InsertShapeStyleProperties">
    <xsl:for-each select="v:shape | v:rect|v:group">
      <!--<xsl:if test="parent::node()[name()='v:group']">
         TO DO v:shapes in v:group 
      </xsl:if>-->
      <xsl:call-template name="InsertShapeWrap"/>
      <xsl:call-template name="InsertShapeFromTextDistance"/>
      <xsl:call-template name="InsertShapeBorders"/>
      <xsl:call-template name="InsertShapeAutomaticWidth"/>
      <xsl:call-template name="InsertShapeHorizontalPos"/>
      <xsl:call-template name="InsertShapeVerticalPos"/>
      <xsl:call-template name="InsertShapeFlowWithText"/>
      <xsl:call-template name="InsertShapeBackgroundColor"/>
      <xsl:call-template name="InsertShapeZindex"/>
      <xsl:call-template name="InsertShapeWrappedParagraph"/>
      <xsl:call-template name="InsertShapeShadow"/>
    </xsl:for-each>
    <xsl:for-each select="v:shape/v:textbox">
      <xsl:call-template name="InsertTexboxTextDirection"/>
      <xsl:call-template name="InsertTextBoxPadding"/>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="InsertShapeWrappedParagraph">
    <!-- TODO inverstigate when this should not be set-->
    <!-- COMMENT : does not exist in OOX, so default value should be no-limit (rather than 1) -->
    <xsl:attribute name="style:number-wrapped-paragraphs">
      <!--xsl:text>1</xsl:text-->
      <xsl:text>no-limit</xsl:text>
    </xsl:attribute>
  </xsl:template>

  <xsl:template name="InsertTexboxTextDirection">
    <xsl:param name="shape" select="."/>
    <xsl:if test="@style = 'layout-flow:vertical' ">
      <xsl:attribute name="style:writing-mode">
        <xsl:text>tb-rl</xsl:text>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

  <!-- insert shadow in a shape. -->
  <xsl:template name="InsertShapeShadow">
    <xsl:for-each select="v:shadow">
      <xsl:attribute name="style:shadow">
        <xsl:choose>
          <xsl:when test="@on = 'false' or @on = 'f' ">none</xsl:when>
          <xsl:otherwise>
            <!-- report lost attributes -->
            <xsl:if test="@opacity">
              <xsl:message terminate="no">translation.oox2odf.shape.shadow</xsl:message>
            </xsl:if>
            <xsl:if test="@obscured">
              <xsl:message terminate="no">translation.oox2odf.shape.shadow.obscurity</xsl:message>
            </xsl:if>
            <xsl:if test="@type">
              <xsl:message terminate="no">translation.oox2odf.shape.shadow.obscurity</xsl:message>
            </xsl:if>
            <xsl:if test="@matrix">
              <xsl:message terminate="no"
              >translation.oox2odf.shape.shadow.complexPerspective</xsl:message>
            </xsl:if>
            <!-- compute color -->
            <xsl:call-template name="InsertColor">
              <xsl:with-param name="color">
                <xsl:choose>
                  <xsl:when test="@color">
                    <xsl:value-of select="@color"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:if test="@color2">
                      <xsl:value-of select="@color2"/>
                    </xsl:if>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:with-param>
            </xsl:call-template>
            <xsl:text> </xsl:text>
            <!-- shadow offset -->
            <xsl:choose>
              <xsl:when test="@offset">
                <!-- horizontal distance -->
                <xsl:call-template name="ComputeShadowDistance">
                  <xsl:with-param name="distance" select="substring-before(@offset, ',')"/>
                  <xsl:with-param name="origin" select="substring-before(@origin, ',')"/>
                  <xsl:with-param name="side">width</xsl:with-param>
                </xsl:call-template>
                <xsl:text> </xsl:text>
                <!-- vertical distance -->
                <xsl:call-template name="ComputeShadowDistance">
                  <xsl:with-param name="distance" select="substring-after(@offset, ',')"/>
                  <xsl:with-param name="origin" select="substring-after(@origin, ',')"/>
                  <xsl:with-param name="side">height</xsl:with-param>
                </xsl:call-template>
              </xsl:when>
              <xsl:otherwise>
                <xsl:if test="@offset2">
                  <xsl:call-template name="ComputeShadowDistance">
                    <xsl:with-param name="distance" select="substring-before(@offset2, ',')"/>
                    <xsl:with-param name="origin" select="substring-before(@origin, ',')"/>
                    <xsl:with-param name="side">width</xsl:with-param>
                  </xsl:call-template>
                  <xsl:text> </xsl:text>
                  <xsl:call-template name="ComputeShadowDistance">
                    <xsl:with-param name="distance" select="substring-after(@offset2, ',')"/>
                    <xsl:with-param name="origin" select="substring-after(@origin, ',')"/>
                    <xsl:with-param name="side">height</xsl:with-param>
                  </xsl:call-template>
                </xsl:if>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:for-each>
  </xsl:template>

  <!-- transform a OOX shadow offset into an ODF shadow distance -->
  <xsl:template name="ComputeShadowDistance">
    <xsl:param name="distance"/>
    <xsl:param name="origin"/>
    <xsl:param name="side"/>
    <!-- if a value requires percent calculation -->
    <xsl:variable name="shapeLength">
      <xsl:if test="contains($distance, '%') or contains($origin, '%')">
        <xsl:call-template name="ConvertMeasure">
          <xsl:with-param name="length">
            <xsl:call-template name="GetShapeProperty">
              <xsl:with-param name="shape" select="parent::node()"/>
              <xsl:with-param name="propertyName" select="$side"/>
            </xsl:call-template>
          </xsl:with-param>
          <xsl:with-param name="destUnit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:if>
    </xsl:variable>
    <!-- compute value of shadow offset -->
    <xsl:variable name="distanceVal">
      <xsl:choose>
        <xsl:when test="$distance = '' ">0</xsl:when>
        <xsl:when test="contains($distance, '%')">
          <xsl:value-of select="round(($shapeLength * $distance) div 100)"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="$distance"/>
            <xsl:with-param name="destUnit">cm</xsl:with-param>
            <xsl:with-param name="addUnit">false</xsl:with-param>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <!-- compute value of shadow origin -->
    <xsl:variable name="originVal">
      <xsl:choose>
        <xsl:when test="$origin = '' ">0</xsl:when>
        <xsl:when test="contains($origin, '%')">
          <xsl:value-of select="round(($shapeLength * $origin) div 100)"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="$origin"/>
            <xsl:with-param name="destUnit">cm</xsl:with-param>
            <xsl:with-param name="addUnit">false</xsl:with-param>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:value-of select="number($distanceVal) + number($originVal)"/>
    <xsl:text>cm</xsl:text>
  </xsl:template>

  <xsl:template name="InsertShapeBackgroundColor">
    <xsl:param name="shape" select="."/>
    <xsl:variable name="bgColor">
      <xsl:choose>
        <xsl:when test="$shape/@fillcolor != '' and $shape/@fillcolor != 'window'">
          <xsl:value-of select="$shape/@fillcolor"/>
        </xsl:when>
        <xsl:otherwise>#ffffff</xsl:otherwise>
      </xsl:choose>
      
    </xsl:variable>
    
    <xsl:variable name="isFilled" select="$shape/@filled"/>
    <xsl:if test="(not($isFilled) or $isFilled != 'f') and $bgColor != ''">
      <xsl:attribute name="fo:background-color">
        <xsl:call-template name="InsertColor">
          <xsl:with-param name="color" select="$bgColor"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="(not($isFilled) or $isFilled != 'f') and ($bgColor = '' or not($bgColor))">
      <xsl:attribute name="fo:background-color">
        <xsl:call-template name="InsertColor">
          <xsl:with-param name="color">#ffffff</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>

    <xsl:attribute name="draw:fill-color">
      <xsl:call-template name="InsertColor">
        <xsl:with-param name="color" select="$bgColor"/>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>

  <xsl:template name="InsertShapeFlowWithText">
    <xsl:param name="shape" select="."/>
    <xsl:variable name="layouitInCell" select="$shape/@o:allowincell"/>

    <xsl:variable name="verticalRelative">
      <xsl:choose>
        <xsl:when test="$shape[name()='w:framePr']">
          <xsl:value-of select="$shape/@w:vAnchor"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="GetShapeProperty">
            <xsl:with-param name="propertyName" select="'mso-position-vertical-relative'"/>
            <xsl:with-param name="shape" select="."/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="horizontalRelative">
      <xsl:choose>
        <xsl:when test="$shape[name()='w:framePr']">
          <xsl:value-of select="$shape/@w:hAnchor"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="GetShapeProperty">
            <xsl:with-param name="propertyName" select="'mso-position-horizontal-relative'"/>
            <xsl:with-param name="shape" select="."/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:attribute name="style:flow-with-text">
      <xsl:choose>
        <xsl:when test="$layouitInCell = 'f' ">
          <xsl:text>false</xsl:text>
        </xsl:when>
        <!-- if ancestor if header or footer and frame is in background -->
        <xsl:when test="(ancestor::w:hdr or ancestor::w:ftr) and not(w10:wrap/@type)">
          <xsl:text>false</xsl:text>
        </xsl:when>
        <xsl:when test="$verticalRelative='page' and $horizontalRelative='page'">
          <xsl:text>false</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:text>true</xsl:text>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>

  <xsl:template name="InsertShapeHorizontalPos">
    <xsl:param name="shape" select="."/>

    <xsl:variable name="horizontalPos">
      <xsl:choose>
        <xsl:when test="$shape[name()='w:framePr']">
          <xsl:value-of select="$shape/@w:xAlign"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="GetShapeProperty">
            <xsl:with-param name="propertyName" select="'mso-position-horizontal'"/>
            <xsl:with-param name="shape" select="."/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="horizontalRel">
      <xsl:choose>
        <xsl:when test="$shape[name()='w:framePr']">
          <xsl:value-of select="$shape/@w:hAnchor"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="GetShapeProperty">
            <xsl:with-param name="propertyName" select="'mso-position-horizontal-relative'"/>
            <xsl:with-param name="shape" select="."/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$shape/@o:hralign and $shape/@o:hr='t'">
        <xsl:call-template name="InsertGraphicPosH">
          <xsl:with-param name="align" select="$shape/@o:hralign"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when
        test="$horizontalPos = '' and ancestor::w:p/w:pPr/w:jc/@w:val and $horizontalRel != 'page'">
        <xsl:if test="ancestor::w:p/w:pPr/w:jc/@w:val">
          <xsl:call-template name="InsertGraphicPosH">
            <xsl:with-param name="align" select="ancestor::w:p/w:pPr/w:jc/@w:val"/>
          </xsl:call-template>
        </xsl:if>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="InsertGraphicPosH">
          <xsl:with-param name="align" select="$horizontalPos"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>

    <xsl:call-template name="InsertGraphicPosRelativeH">
      <xsl:with-param name="relativeFrom">
        <xsl:value-of select="$horizontalRel"/>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>


  <xsl:template name="InsertShapeVerticalPos">
    <xsl:param name="shape" select="."/>
    <xsl:variable name="verticalPos">
      <xsl:choose>
        <xsl:when test="$shape[name()='w:framePr']">
          <xsl:value-of select="$shape/@w:yAlign"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="GetShapeProperty">
            <xsl:with-param name="propertyName" select="'mso-position-vertical'"/>
            <xsl:with-param name="shape" select="."/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="verticalRelative">
      <xsl:choose>
        <xsl:when test="$shape[name()='w:framePr']">
          <xsl:value-of select="$shape/@w:vAnchor"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="GetShapeProperty">
            <xsl:with-param name="propertyName" select="'mso-position-vertical-relative'"/>
            <xsl:with-param name="shape" select="."/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
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
    <xsl:param name="shape" select="."/>
    <xsl:choose>
      <xsl:when test="self::w:framePr">
        <xsl:if test="@w:wrap">
          <xsl:call-template name="InsertShapeWrapAttributes">
            <xsl:with-param name="shape" select="$shape"/>
            <xsl:with-param name="wrap" select="."/>
          </xsl:call-template>
        </xsl:if>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="InsertShapeWrapAttributes">
          <xsl:with-param name="shape" select="$shape"/>
          <xsl:with-param name="wrap" select="w10:wrap"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertShapeWrapAttributes">
    <xsl:param name="shape"/>
    <xsl:param name="wrap"/>

    <xsl:variable name="zindex">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="propertyName" select="'z-index'"/>
        <xsl:with-param name="shape" select="."/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$zindex &lt; 0 and not($wrap/@type) and not(@w:wrap)">
      <xsl:attribute name="style:run-through">
        <xsl:text>background</xsl:text>
      </xsl:attribute>
    </xsl:if>

    <xsl:attribute name="style:wrap">
      <xsl:choose>
        <xsl:when test="not($wrap)">
          <xsl:choose>
            <!-- no wrap element and no z-index -->
            <xsl:when test="$zindex = '' ">
              <xsl:text>none</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>run-through</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="not($wrap/@type)">
          <xsl:text>run-through</xsl:text>
        </xsl:when>
        <xsl:when
          test="($wrap/@type = 'square' or $wrap/@type = 'tight' or $wrap/@type = 'through' or @w:wrap='around') and not($wrap/@side)">
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
        <xsl:when test="not(@w:wrap) or @w:wrap='none' ">
          <xsl:text>none</xsl:text>
        </xsl:when>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>


  <xsl:template name="InsertShapeFromTextDistance">
    <xsl:param name="shape" select="."/>

    <xsl:variable name="horizontalRelative">
      <xsl:choose>
        <xsl:when test="$shape[name()='w:framePr']">
          <xsl:value-of select="$shape/@w:hAnchor"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="GetShapeProperty">
            <xsl:with-param name="propertyName" select="'mso-position-horizontal-relative'"/>
            <xsl:with-param name="shape" select="$shape"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="verticalRelative">
      <xsl:choose>
        <xsl:when test="$shape[name()='w:framePr']">
          <xsl:value-of select="$shape/@w:vAnchor"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="GetShapeProperty">
            <xsl:with-param name="propertyName" select="'mso-position-vertical-relative'"/>
            <xsl:with-param name="shape" select="$shape"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="horizontalPosition">
      <xsl:choose>
        <xsl:when test="$shape[name()='w:framePr']">
          <xsl:value-of select="$shape/@w:xAlign"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="GetShapeProperty">
            <xsl:with-param name="propertyName" select="'mso-position-horizontal'"/>
            <xsl:with-param name="shape" select="$shape"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="verticalPosition">
      <xsl:choose>
        <xsl:when test="$shape[name()='w:framePr']">
          <xsl:value-of select="$shape/@w:yAlign"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="GetShapeProperty">
            <xsl:with-param name="propertyName" select="'mso-position-vertical'"/>
            <xsl:with-param name="shape" select="$shape"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="margin-top">
      <xsl:choose>
        <xsl:when test="$verticalRelative='page' and $verticalPosition='top'">
          <xsl:text>0</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:choose>
            <xsl:when test="$shape[name()='w:framePr']">
              <xsl:value-of select="$shape/@w:vSpace"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="GetShapeProperty">
                <xsl:with-param name="propertyName" select="'mso-wrap-distance-top'"/>
                <xsl:with-param name="shape" select="$shape"/>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:if test="$margin-top !=''">
      <xsl:call-template name="InsertShapeMargin">
        <xsl:with-param name="attributeName" select=" 'fo:margin-top' "/>
        <xsl:with-param name="margin" select="$margin-top"/>
      </xsl:call-template>
    </xsl:if>

    <xsl:variable name="margin-bottom">
      <xsl:choose>
        <xsl:when test="$verticalRelative='page' and $verticalPosition='bottom'">
          <xsl:text>0</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:choose>
            <xsl:when test="$shape[name()='w:framePr']">
              <xsl:value-of select="$shape/@w:vSpace"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="GetShapeProperty">
                <xsl:with-param name="propertyName" select="'mso-wrap-distance-bottom'"/>
                <xsl:with-param name="shape" select="$shape"/>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:if test="$margin-bottom !=''">
      <xsl:call-template name="InsertShapeMargin">
        <xsl:with-param name="attributeName" select=" 'fo:margin-bottom' "/>
        <xsl:with-param name="margin" select="$margin-bottom"/>
      </xsl:call-template>
    </xsl:if>

    <xsl:variable name="margin-left">
      <xsl:choose>
        <xsl:when test="$horizontalRelative='page' and $horizontalPosition='left'">
          <xsl:text>0</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:choose>
            <xsl:when test="$shape[name()='w:framePr']">
              <xsl:value-of select="$shape/@w:hSpace"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="GetShapeProperty">
                <xsl:with-param name="propertyName" select="'mso-wrap-distance-left'"/>
                <xsl:with-param name="shape" select="$shape"/>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:if test="$margin-left !=''">
      <xsl:call-template name="InsertShapeMargin">
        <xsl:with-param name="attributeName" select=" 'fo:margin-left' "/>
        <xsl:with-param name="margin" select="$margin-left"/>
      </xsl:call-template>
    </xsl:if>

    <xsl:variable name="margin-right">
      <xsl:choose>
        <xsl:when test="$horizontalRelative='page' and $horizontalPosition='right'">
          <xsl:text>0</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:choose>
            <xsl:when test="$shape[name()='w:framePr']">
              <xsl:value-of select="$shape/@w:hSpace"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="GetShapeProperty">
                <xsl:with-param name="propertyName" select="'mso-wrap-distance-right'"/>
                <xsl:with-param name="shape" select="$shape"/>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
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
    <xsl:param name="shape" select="."/>

    <xsl:variable name="textBoxInset" select="@inset"/>

    <xsl:attribute name="fo:padding-left">
      <xsl:call-template name="ConvertMeasure">
        <xsl:with-param name="length">
          <xsl:call-template name="ExplodeValues">
            <xsl:with-param name="elementNum" select="1"/>
            <xsl:with-param name="text" select="$textBoxInset"/>
          </xsl:call-template>
        </xsl:with-param>
        <xsl:with-param name="destUnit" select="'cm'"/>
      </xsl:call-template>
    </xsl:attribute>

    <xsl:attribute name="fo:padding-top">
      <xsl:call-template name="ConvertMeasure">
        <xsl:with-param name="length">
          <xsl:call-template name="ExplodeValues">
            <xsl:with-param name="elementNum" select="2"/>
            <xsl:with-param name="text" select="$textBoxInset"/>
          </xsl:call-template>
        </xsl:with-param>
        <xsl:with-param name="destUnit" select="'cm'"/>
      </xsl:call-template>
    </xsl:attribute>

    <xsl:attribute name="fo:padding-right">
      <xsl:call-template name="ConvertMeasure">
        <xsl:with-param name="length">
          <xsl:call-template name="ExplodeValues">
            <xsl:with-param name="elementNum" select="3"/>
            <xsl:with-param name="text" select="$textBoxInset"/>
          </xsl:call-template>
        </xsl:with-param>
        <xsl:with-param name="destUnit" select="'cm'"/>
      </xsl:call-template>
    </xsl:attribute>

    <xsl:attribute name="fo:padding-bottom">
      <xsl:call-template name="ConvertMeasure">
        <xsl:with-param name="length">
          <xsl:call-template name="ExplodeValues">
            <xsl:with-param name="elementNum" select="4"/>
            <xsl:with-param name="text" select="$textBoxInset"/>
          </xsl:call-template>
        </xsl:with-param>
        <xsl:with-param name="destUnit" select="'cm'"/>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>

  <xsl:template name="ExplodeValues">
    <xsl:param name="elementNum"/>
    <xsl:param name="text"/>
    <xsl:param name="counter" select="1"/>
    <xsl:choose>
      <xsl:when test="$elementNum = $counter">
        <xsl:variable name="length">
          <xsl:choose>
            <xsl:when test="contains($text,',')">
              <xsl:value-of select="substring-before($text,',')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$text"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:choose>
          <!-- default values -->
          <xsl:when test="$length = '' and ($counter = 1 or $counter = 3)">0.1in</xsl:when>
          <xsl:when test="$length = '' and ($counter = 2 or $counter = 4)">0.05in</xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$length"/>
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
    <xsl:param name="shape" select="."/>

    <xsl:choose>
      <xsl:when test="$shape/@o:hr='t' or $shape/@stroked = 'f' ">
        <xsl:attribute name="draw:stroke">
          <xsl:text>none</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <!--  Word sets default values for borders  when no strokeweight and strokecolor is set and @stroked is not set to f (this does not work for ole-objects (v:imagedata))-->
      <xsl:when
        test="not($shape/@strokeweight) and not($shape/@strokecolor) and not($shape/v:imagedata)">
        <xsl:attribute name="fo:border">
          <xsl:text>0.0176cm solid #000000</xsl:text>
        </xsl:attribute>
      </xsl:when>
      <!--  default scenario -->
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
          <xsl:attribute name="svg:stroke-color">
            <xsl:value-of select="$borderColor"/>
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
    <xsl:param name="shape" select="."/>
    <xsl:if
      test="contains(@style,'mso-fit-shape-to-text:t') or (not($shape/@w:h) and $shape/@hRule != 'exact')">
      <xsl:attribute name="fo:min-height">
        <xsl:choose>
          <xsl:when test="$shape/@hRule='atLeast'">
            <xsl:value-of select="$shape/@hRule"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>0cm</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertShapeAutomaticWidth">
    <xsl:param name="shape" select="."/>
    <xsl:variable name="wrapStyle">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$shape"/>
        <xsl:with-param name="propertyName" select="'mso-wrap-style'"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:choose>
      <xsl:when
        test="($wrapStyle != '' and $wrapStyle = 'none') or ( $shape/@w:wrap and $shape/@w:wrap != 'none' )">
        <xsl:attribute name="fo:min-width">
          <xsl:text>0cm</xsl:text>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- text frame paragraphs -->
  <xsl:template match="w:p[w:pPr/w:framePr]">
    <xsl:choose>
      <!-- skip drop-capped paragraphs -->
      <xsl:when test="w:pPr/w:framePr[@w:dropCap = 'drop']"/>
      <!-- margin drop cap -->
      <xsl:when test="w:pPr/w:framePr[@w:dropCap = 'margin' ]">
        <xsl:message terminate="no">translation.oox2odf.dropcap.inMargin</xsl:message>
      </xsl:when>
      <xsl:otherwise>
        <!-- if previous-brother node (in the same context) wasn't a text frame paragraph-->
        <!-- then create another text:p else merge text with the existing previous-brother-->
        <!-- BUG FIX 1642428-->
        <xsl:if test="not(preceding-sibling::w:p[position()=1][descendant::w:framePr])">
          <text:p>
            <xsl:attribute name="text:style-name">
              <xsl:choose>
                <xsl:when test="w:pPr/w:pStyle/@w:val">
                  <xsl:value-of select="w:pPr/w:pStyle/@w:val"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="generate-id()"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <draw:frame>

              <xsl:attribute name="draw:style-name">
                <xsl:value-of select="generate-id(w:pPr/w:framePr)"/>
              </xsl:attribute>              

              <xsl:call-template name="InsertShapeAnchor">
                <xsl:with-param name="shape" select="w:pPr/w:framePr"/>
              </xsl:call-template>

              <xsl:call-template name="InsertShapeWidth">
                <xsl:with-param name="shape" select="w:pPr/w:framePr"/>
              </xsl:call-template>
              
              <xsl:call-template name="InsertShapeHeight">
                <xsl:with-param name="shape" select="w:pPr/w:framePr"/>
              </xsl:call-template>
              
              <xsl:call-template name="InsertshapeAbsolutePos">
                <xsl:with-param name="shape" select="w:pPr/w:framePr"/>
              </xsl:call-template>
              <draw:text-box>
                <xsl:call-template name="InsertTextBoxAutomaticHeight">
                  <xsl:with-param name="shape" select="w:pPr/w:framePr"/>
                </xsl:call-template>
                <xsl:call-template name="InsertParagraphToFrame"/>
              </draw:text-box>
            </draw:frame>
          </text:p>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertParagraphToFrame">
    <xsl:param name="paragraph" select="."/>
    <text:p>
      <xsl:attribute name="text:style-name">
        <xsl:choose>
          <xsl:when test="$paragraph/w:pPr/w:pStyle/@w:val">
            <xsl:value-of select="$paragraph/w:pPr/w:pStyle/@w:val"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="generate-id($paragraph)"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:apply-templates select="$paragraph/node()"/>
    </text:p>
    <!-- if next paragraph is also a text frame paragraph then insert its content to the same text-box -->
    <xsl:if test="$paragraph/following::*[position()=1][w:pPr/w:framePr]">
      <xsl:call-template name="InsertParagraphToFrame">
        <xsl:with-param name="paragraph"
          select="$paragraph/following::*[position()=1][w:pPr/w:framePr]"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

  <xsl:template match="w:p[w:pPr/w:framePr]" mode="automaticstyles">
    <xsl:choose>
      <xsl:when test="w:pPr/w:framePr[@w:dropCap = 'drop' ]"/>
      <!-- margin drop cap -->
      <xsl:when test="w:pPr/w:framePr[@w:dropCap = 'margin' ]">
        <xsl:message terminate="no">translation.oox2odf.dropcap.inMargin</xsl:message>
      </xsl:when>
      <xsl:otherwise>
        <style:style style:name="{generate-id(w:pPr/w:framePr)}" style:family="graphic"
          style:parent-style-name="Frame">
          <style:graphic-properties>

            <!--<xsl:call-template name="InsertShapeStyleName">
              <xsl:with-param name="shape" select="w:pPr/w:framePr"/>
            </xsl:call-template>-->

           <xsl:call-template name="InsertFramePrAnchor">
              <xsl:with-param name="oFramePr" select="w:pPr/w:framePr"/>
            </xsl:call-template>

            <!--<xsl:call-template name="InsertShapeWrap">
              <xsl:with-param name="shape" select="w:pPr/w:framePr"/>
            </xsl:call-template>-->
            
            <xsl:call-template name="InsertShapeFromTextDistance">
              <xsl:with-param name="shape" select="w:pPr/w:framePr"/>
            </xsl:call-template>
            
            <!--<xsl:call-template name="InsertShapeHorizontalPos">
              <xsl:with-param name="shape" select="w:pPr/w:framePr"/>
            </xsl:call-template>-->
            
            <!--<xsl:call-template name="InsertShapeVerticalPos">
              <xsl:with-param name="shape" select="w:pPr/w:framePr"/>
            </xsl:call-template>-->
            
            <xsl:call-template name="InsertShapeAutomaticWidth">
              <xsl:with-param name="shape" select="w:pPr/w:framePr"/>
            </xsl:call-template>
            
            <xsl:if test="not(w:pPr/w:pBdr)">
              <xsl:attribute name="fo:border">
                <xsl:text>none</xsl:text>
              </xsl:attribute>
            </xsl:if>
            <xsl:apply-templates select="w:pPr" mode="pPrChildren"/>
          </style:graphic-properties>
        </style:style>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:apply-templates mode="automaticstyles"/>
  </xsl:template>

  <xsl:template name="InsertFramePrAnchor">
    <xsl:param name="oFramePr"/>

    <!-- Translation Rules                   -->
    <!-- w:wrap    =>  style:wrap            -->
    <!-- w:yAlign  =>  style:vertical-pos    -->
    <!-- w:vAnchor =>  style:vertical-rel    -->
    <!-- w:xAlign  =>  style:horizontal-pos  -->
    <!-- w:hAnchor =>  style:horizontal-rel  -->

    <xsl:variable name="Wrap"     select = "$oFramePr/@w:wrap"/>
    <xsl:variable name="yAlign"   select = "$oFramePr/@w:yAlign"/>
    <xsl:variable name="vAnchor"  select = "$oFramePr/@w:vAnchor"/>
    <xsl:variable name="xAlign"   select = "$oFramePr/@w:xAlign"/>
    <xsl:variable name="hAnchor"  select = "$oFramePr/@w:hAnchor"/>

    <!-- w:wrap    =>  style:wrap            -->
    <xsl:attribute name="style:wrap">
      <xsl:choose>
        <xsl:when test="not($Wrap) or $Wrap ='none'">
              <xsl:text>none</xsl:text>
        </xsl:when>
        <xsl:when test="$Wrap ='auto'">
          <xsl:text>parallel</xsl:text>
        </xsl:when>
        <xsl:when test="$Wrap ='around'">
          <xsl:text>parallel</xsl:text>
        </xsl:when>
        <xsl:when test="$Wrap ='notBeside'">
          <xsl:text>parallel</xsl:text>
        </xsl:when>

        <xsl:when test="$Wrap ='through'">
          <xsl:text>parallel</xsl:text>
        </xsl:when>

        <xsl:when test="$Wrap ='tight'">
          <xsl:text>parallel</xsl:text>
        </xsl:when>
      </xsl:choose>
    </xsl:attribute>

    <!-- w:yAlign  =>  style:vertical-pos    -->
    <!-- w:vAnchor =>  style:vertical-rel    -->

    <xsl:if test ="count($yAlign)>0 or count($vAnchor)>0">

      <xsl:attribute name="style:vertical-pos">
        <xsl:choose>
          <xsl:when test="$yAlign='bottom'">
            <xsl:value-of select="'bottom'"/>
          </xsl:when>
          <xsl:when test="$yAlign='top'">
            <xsl:value-of select="'top'"/>
          </xsl:when>
          <xsl:when test="$yAlign='center'">
            <xsl:value-of select="'middle'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'auto'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>

      
      <xsl:attribute name="style:vertical-rel">
        <xsl:choose>
          <xsl:when test="$vAnchor='page'">
            <xsl:value-of select="'page'"/>
          </xsl:when>
          <xsl:when test="$vAnchor='text'">
            <xsl:value-of select="'baseline'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'page'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>

    </xsl:if>


    <!-- w:xAlign  =>  style:horizontal-pos  -->
    <!-- w:hAnchor =>  style:horizontal-rel  -->
    
    <xsl:if test ="count($xAlign)>0 or count($hAnchor)>0">

      <xsl:attribute name="style:horizontal-pos">
        <xsl:choose>
          <xsl:when test="$xAlign='center'">
            <xsl:value-of select="'middle'"/>
          </xsl:when>
          <xsl:when test="$xAlign='inside'">
            <xsl:value-of select="'from-left'"/>
          </xsl:when>
          <xsl:when test="$xAlign='outside'">
            <xsl:value-of select="'left'"/>
          </xsl:when>
          <xsl:when test="$xAlign='left'">
            <xsl:value-of select="'left'"/>
          </xsl:when>
          <xsl:when test="$xAlign='right'">
            <xsl:value-of select="'right'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'from-left'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute >

      <xsl:attribute name="style:horizontal-rel">
        <xsl:choose>
          <xsl:when test="$hAnchor='margin'">
            <xsl:value-of select="'page-start-margin'"/>
          </xsl:when>
          <xsl:when test="$hAnchor='page'">
            <xsl:value-of select="'page'"/>
          </xsl:when>
          <xsl:when test="$hAnchor='text'">
            <xsl:value-of select="'page-content'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'page-content'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>

    </xsl:if>


  </xsl:template>

              
</xsl:stylesheet>
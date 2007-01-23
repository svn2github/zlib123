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
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  exclude-result-prefixes="w r draw number wp xlink v w10 o">

  <!-- shape or ole-object-->
  <xsl:template match="w:pict | w:object">
    <xsl:apply-templates/>
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
    <xsl:message terminate="no">feedback:Shape 3D effects</xsl:message>
  </xsl:template>

  <!--horizontal line-->
  <xsl:template match="v:rect">
    <!--    version 1.1-->
    <!--<draw:rect>
      <xsl:call-template name="InsertCommonShapeProperties"/>
    </draw:rect>-->
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
      <xsl:message terminate="no">feedback:Text watermark</xsl:message>
    </xsl:if>
  </xsl:template>

  <xsl:template match="v:textbox">
    <xsl:if test="parent::v:stroke/@dashstyle">
      <xsl:message terminate="no">feedback:Dashed textbox border</xsl:message>
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

    <xsl:if test="$position = 'absolute' or $shape[name()='w:framePr']">
      <xsl:variable name="svgx">
        <xsl:choose>
          <xsl:when test="$shape[name()='w:framePr']">
            <xsl:value-of select="$shape/@w:x"/>
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
            <xsl:message terminate="no">feedback:Position of frame in header</xsl:message>
          </xsl:if>
          <xsl:if test="ancestor::w:ftr">
            <xsl:message terminate="no">feedback:Position of frame in footer</xsl:message>
          </xsl:if>
          <xsl:text>as-char</xsl:text>
        </xsl:when>
        <!-- if there is another run exept that one containing shape and shape doesn't have wrapping style set then shape should be anchored 'as-text' -->
        <xsl:when
          test="ancestor::w:r/parent::node()/w:r[2] and not(w10:wrap) and (not(contains($shape/@style, 'position:absolute')) or contains($shape/@style, 'mso-position-horizontal-relative:text') or contains($shape/@style, 'mso-position-vertical-relative:text')  or ($shape/@vAnchor='text') or ($shape/@hAnchor='relative'))">
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
      <xsl:message terminate="no">feedback:Relative frame size </xsl:message>
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
      <xsl:message terminate="no">feedback:Relative frame size </xsl:message>
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

  <xsl:template match="w:pict | w:object" mode="automaticstyles">
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

    <xsl:apply-templates mode="automaticstyles"/>
  </xsl:template>

  <xsl:template name="InsertShapeStyleProperties">
    <xsl:for-each select="v:shape | v:rect">
      <xsl:if test="parent::node()[name()='v:group']">
        <!-- TO DO v:shapes in v:group -->
      </xsl:if>
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
              <xsl:message terminate="no">feedback:Shadow opacity</xsl:message>
            </xsl:if>
            <xsl:if test="@obscured">
              <xsl:message terminate="no">feedback:Shadow obscurity</xsl:message>
            </xsl:if>
            <xsl:if test="@type">
              <xsl:message terminate="no">feedback:Shadow obscurity</xsl:message>
            </xsl:if>
            <xsl:if test="@matrix">
              <xsl:message terminate="no">feedback:Shadow complex perspective</xsl:message>
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
    <xsl:variable name="bgColor" select="$shape/@fillcolor"/>
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

    <xsl:choose>
      <xsl:when test="$shape/@o:hralign and $shape/@o:hr='t'">
        <xsl:call-template name="InsertGraphicPosH">
          <xsl:with-param name="align" select="$shape/@o:hralign"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$horizontalPos = '' and ancestor::w:p/w:pPr/w:jc/@w:val ">
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
        <xsl:if test="w10:wrap">
          <xsl:call-template name="InsertShapeWrapAttributes">
            <xsl:with-param name="shape" select="$shape"/>
            <xsl:with-param name="wrap" select="w10:wrap"/>
          </xsl:call-template>
        </xsl:if>
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
        test="($wrapStyle != '' and $wrapStyle = 'none') or ( $shape/w:wrap and not($shape/@w:wrap='none'))">
        <xsl:attribute name="fo:min-width">
          <xsl:text>0cm</xsl:text>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

</xsl:stylesheet>

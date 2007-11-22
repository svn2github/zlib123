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
  xmlns:o="urn:schemas-microsoft-com:office:office" 
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:oox="urn:oox"
  exclude-result-prefixes="w r draw number wp xlink v w10 o oox">

  <!-- 
  *************************************************************************
  MATCHING TEMPLATES
  *************************************************************************
  -->

  <!-- 
  Summary: Converts frames
  Author: Clever Age
  -->
  <xsl:template match="w:p[w:pPr/w:framePr]">
    <xsl:choose>
      <xsl:when test="w:pPr/w:framePr[@w:dropCap='drop']">
        <!--Do nothing-->
      </xsl:when>
      <xsl:when test="w:pPr/w:framePr[@w:dropCap='margin']">
        <xsl:message terminate="no">translation.oox2odf.dropcap.inMargin</xsl:message>
      </xsl:when>
      <xsl:otherwise>
        <!-- 
        check if the preceding framePr is identically 
        If it is, then it's w:p belongs to the same frame and 
        this frame was already merged in to the previous.
        -->
        <xsl:variable name="precedingP" select="./preceding-sibling::*[1]" />
        <xsl:variable name="identically">
          <xsl:call-template name="CompareFrames">
            <xsl:with-param name="frame1" select="./w:pPr/w:framePr"/>
            <xsl:with-param name="frame2" select="$precedingP/w:pPr/w:framePr"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:if test="$identically='false'">
          <!--
          makz: If the frame is anchored to the page insert only the frame
          else insert the paragraph and the frame.
          
          A frame which iss not aligned to the page (in horiz. or vert. direction) 
          needs a paragraph due to it's anchor.
          -->
          <xsl:choose>
            <xsl:when test="w:pPr/w:framePr/@w:hAnchor='page' and w:pPr/w:framePr/@w:vAnchor='page'">
              <xsl:call-template name="InsertFrame">
                <xsl:with-param name="framePr" select="w:pPr/w:framePr" />
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
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
                <xsl:call-template name="InsertFrame">
                  <xsl:with-param name="framePr" select="w:pPr/w:framePr" />
                </xsl:call-template>
              </text:p>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--
  Summary: Convert the frame properties into automatic styles
  Author: Clever Age
  -->
  <xsl:template match="w:p[w:pPr/w:framePr]" mode="automaticstyles">
    <xsl:choose>
      <xsl:when test="w:pPr/w:framePr[@w:dropCap='drop']">
        <!--Do nothing-->
      </xsl:when>
      <xsl:when test="w:pPr/w:framePr[@w:dropCap='margin']">
        <xsl:message terminate="no">translation.oox2odf.dropcap.inMargin</xsl:message>
      </xsl:when>
      <xsl:otherwise>
        <style:style style:name="{generate-id(w:pPr/w:framePr)}" style:family="graphic" style:parent-style-name="Frame">
          <style:graphic-properties>

            <xsl:call-template name="InsertFramePrAnchor">
              <xsl:with-param name="oFramePr" select="w:pPr/w:framePr"/>
            </xsl:call-template>

            <xsl:call-template name="InsertShapeFromTextDistance">
              <xsl:with-param name="shape" select="w:pPr/w:framePr"/>
            </xsl:call-template>

            <xsl:call-template name="InsertShapeAutomaticWidth">
              <xsl:with-param name="shape" select="w:pPr/w:framePr"/>
            </xsl:call-template>

            <!--
            The border properties of a border can be defined in the automatic styles of the content.xml
            or in the styles.xml
            That happens if a predefined paragraph style is made to a frame.
            -->
            <xsl:variable name="styleId" select="w:pPr/w:pStyle/@w:val"/>
            <xsl:variable name="externalBorderStyle" select="key('StyleId', $styleId)"/>
            <xsl:choose>
              <!-- The border is defined in context.xml -->
              <xsl:when test="w:pPr/w:pBdr">
                <xsl:call-template name="InsertBorders">
                  <xsl:with-param name="border" select="w:pPr/w:pBdr"/>
                </xsl:call-template>
                <xsl:call-template name="InsertParagraphShadow"/>
              </xsl:when>
              <!-- The border is defined in styles.xml -->
              <xsl:when test="$externalBorderStyle/w:pPr/w:pBdr">
                <xsl:call-template name="InsertBorders">
                  <xsl:with-param name="border" select="$externalBorderStyle/w:pPr/w:pBdr"/>
                </xsl:call-template>
                <xsl:call-template name="InsertParagraphShadow"/>
              </xsl:when>
              <!-- No border is defined -->
              <xsl:otherwise>
                <xsl:attribute name="fo:border">
                  <xsl:text>none</xsl:text>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>

            <xsl:apply-templates select="w:pPr" mode="pPrChildren"/>
          </style:graphic-properties>
        </style:style>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:apply-templates mode="automaticstyles"/>
  </xsl:template>

  <!--
  Summary: Converts textbox shapes
  Author: Clever Age
  -->
  <xsl:template match="v:shape[v:textbox]">
    <draw:frame>
      <xsl:call-template name="InsertCommonShapeProperties">
        <xsl:with-param name="shape" select="." />
      </xsl:call-template>
      <xsl:call-template name="InsertShapeZindex"/>
      
      <xsl:apply-templates/>
      <!-- some of the shape types must be in odf draw:frame even if they are outside of v:shape in oox-->
      <xsl:apply-templates select="self::node()/following-sibling::node()[1]" mode="draw:frame"/>
    </draw:frame>
  </xsl:template>

  <!--
  Summary: Template writes rectangles/lines.
  Author: Clever Age
  -->
  <xsl:template match="v:rect | v:line">
    <!-- version 1.1-->
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
            <xsl:value-of select="generate-id(.)"/>
          </xsl:attribute>
          <!--TODO-->
          <xsl:attribute name="draw:name">
            <xsl:text>Frame1</xsl:text>
          </xsl:attribute>
          <xsl:call-template name="InsertAnchorType"/>
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
            <xsl:if test="key('Part', concat('word/_rels/',$document,'.rels'))">
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
        <xsl:choose>
          <xsl:when test="self::v:line">
            <draw:line>
              <xsl:attribute name="draw:style-name">
                <xsl:value-of select="generate-id(.)"/>
              </xsl:attribute>
              <!--TODO-->
              <xsl:attribute name="draw:name">
                <xsl:text>Frame1</xsl:text>
              </xsl:attribute>
              <xsl:variable name="flip">
                <xsl:if test="contains(@style,'flip')">
                  <xsl:value-of select="substring-before(substring-after(@style,'flip:'),';')"/>
                </xsl:if>
              </xsl:variable>
              <xsl:call-template name="InsertAnchorType"/>
              <xsl:call-template name="InsertShapeZindex"/>
              <xsl:call-template name="InsertLinePos1">
                <xsl:with-param name="flip">
                  <xsl:value-of select="$flip"/>
                </xsl:with-param>
              </xsl:call-template>
              <xsl:call-template name="InsertLinePos2">
                <xsl:with-param name="flip">
                  <xsl:value-of select="$flip"/>
                </xsl:with-param>
              </xsl:call-template>
              
              <!--<xsl:call-template name="InsertParagraphToFrame"/>-->
            </draw:line>
          </xsl:when>
          <xsl:otherwise>
            <draw:rect>
              <xsl:attribute name="draw:style-name">
                <xsl:value-of select="generate-id(.)"/>
              </xsl:attribute>
              <!--TODO-->
              <xsl:attribute name="draw:name">
                <xsl:text>Frame1</xsl:text>
              </xsl:attribute>
              
              <xsl:call-template name="InsertAnchorType"/>
              <xsl:call-template name="InsertShapeWidth"/>
              <xsl:call-template name="InsertShapeHeight"/>
              <xsl:call-template name="InsertshapeAbsolutePos"/>

              <xsl:apply-templates select="v:textbox"/>
              
              <!--<xsl:call-template name="InsertParagraphToFrame"/>-->
            </draw:rect>
          </xsl:otherwise>
        </xsl:choose>

      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--
  makz: There is no need to generate a style for every v:rect 
  because there is already a style generates for the rect's pict
  
  <xsl:template match="v:rect|v:line" mode="automaticpict">
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
    <xsl:when test="v:imagedata">
    </xsl:when>
    </xsl:choose>
    <xsl:call-template name="InsertShapeShadow"/>
    </style:graphic-properties>
    </style:style>
    <xsl:apply-templates mode="automaticpict"/>
  </xsl:template>
  -->

  <!--
  Summary: Writes gradient fill style
  Author: makz
  Date: 6.11.2007
  -->
  <xsl:template match="v:fill[@type='gradient']" mode="officestyles">
    <xsl:variable name="parentShape" select="parent::node()" />
    <xsl:variable name="gradientName" select="concat('Gradient_', generate-id(.))" />
    <xsl:variable name="focusValue">
      <xsl:value-of select="substring-before(@focus, '%')"/>
    </xsl:variable>
    
    <draw:gradient>
      <xsl:attribute name="draw:name">
        <xsl:value-of select="$gradientName"/>
      </xsl:attribute>
      <xsl:attribute name="draw:display-name">
        <xsl:value-of select="$gradientName"/>
      </xsl:attribute>
      <xsl:attribute name="draw:style">
        <xsl:choose>
          <xsl:when test="not($focusValue) or $focusValue='100' or $focusValue='-100'">
            <xsl:text>linear</xsl:text>
          </xsl:when>
          <xsl:when test="$focusValue='50' or $focusValue='-50'">
            <xsl:text>axial</xsl:text>
          </xsl:when>
        </xsl:choose>
      </xsl:attribute>
      <xsl:attribute name="draw:start-color">
        <xsl:choose>
          <xsl:when test="(@rotate='t' or @rotate='true' or @rotate='1') and $focusValue>0">
            <xsl:call-template name="InsertEndColor" />
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="InsertStartColor" />
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:attribute name="draw:end-color">
        <xsl:choose>
          <xsl:when test="(@rotate='t' or @rotate='true' or @rotate='1') and $focusValue>0">
            <xsl:call-template name="InsertStartColor" />
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="InsertEndColor" />
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:attribute name="draw:start-intensity">
        <xsl:choose>
          <xsl:when test="@opacity">
            <!-- calculate opacity -->
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>100%</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:attribute name="draw:end-intensity">
        <xsl:choose>
          <xsl:when test="@o:opacity2">
            <!-- calculate opacity -->
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>100%</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:attribute name="draw:angle">
        <xsl:choose>
          <xsl:when test="@angle">
            <xsl:value-of select="@angle"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>0</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:attribute name="draw:border">0%</xsl:attribute>
    </draw:gradient>
  </xsl:template>

  <!--
  Summary: Hyperlinks in shapes
  Author: Clever Age
  -->
  <xsl:template match="w:pict[v:shape/@href]">
    <draw:a xlink:type="simple" xlink:href="{v:shape/@href}">
      <xsl:apply-templates/>
    </draw:a>
  </xsl:template>

  <!--
  Summary: text watermark feedback
  Author: Clever Age
  -->
  <xsl:template match="w:pict[contains(v:shape/@id,'WaterMark')]">
    <xsl:if test="v:shape/v:textpath">
      <xsl:message terminate="no">translation.oox2odf.background.textWatermark</xsl:message>
    </xsl:if>
  </xsl:template>

  <!-- 
  Summary: Template writes text boxes.
  Author: Clever Age
  -->
  <xsl:template match="v:textbox">
    <xsl:if test="parent::v:stroke/@dashstyle">
      <xsl:message terminate="no">translation.oox2odf.textbox.boder.dashed</xsl:message>
    </xsl:if>
    <xsl:if test="contains(parent::node()/@style, 'v-text-anchor')">
      <xsl:message terminate="no">translation.odf2oox.valignInsideTextbox</xsl:message>
    </xsl:if>
    <draw:text-box>
      <xsl:call-template name="InsertTextBoxAutomaticHeight"/>
      <xsl:apply-templates select="w:txbxContent/child::node()"/>
    </draw:text-box>
  </xsl:template>

  <!--
  Summary:
  Author: Clever Age
  -->
  <xsl:template match="w:pict">
    <xsl:choose>
      <xsl:when test="v:group">
        <draw:g>
          <xsl:attribute name="draw:style-name">
            <xsl:value-of select="generate-id(.)"/>
          </xsl:attribute>
          <xsl:call-template name="InsertAnchorType">
            <xsl:with-param name="shape" select="v:group" />
          </xsl:call-template>
          <xsl:apply-templates/>
        </draw:g>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- 
  Summary: inserts horizontal ruler as image
  Author: Clever Age
  -->
  <xsl:template match="v:imagedata[not(../../o:OLEObject)]">
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
      <xsl:if test="key('Part', concat('word/_rels/',$document,'.rels'))">
        <xsl:call-template name="InsertImageHref">
          <xsl:with-param name="document" select="$document"/>
          <xsl:with-param name="rId" select="@r:id"/>
        </xsl:call-template>
      </xsl:if>
    </draw:image>
  </xsl:template>

  <!-- 
  Summary: Writes the style of PICT's
  Author: Clever Age
  -->
  <xsl:template match="w:pict[not(o:OLEObject)]" mode="automaticpict">
    <xsl:variable name="vmlElement" select="v:shape | v:rect | v:line | v:group" />
    
    <style:style>
      <xsl:attribute name="style:name">
        <xsl:value-of select="generate-id($vmlElement)"/>
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
        <xsl:call-template name="InsertShapeStyleProperties">
          <xsl:with-param name="shape" select="$vmlElement" />
        </xsl:call-template>
      </style:graphic-properties>
    </style:style>
    <xsl:apply-templates mode="automaticpict"/>
  </xsl:template>

  <!-- 
  Summary: don't process text with automatic pict mode
  Author: Clever Age
  -->
  <xsl:template match="w:t" mode="automaticpict"/>

  <!--
  Summary:
  Author: Clever Age
  -->
  <xsl:template match="o:extrusion">
    <xsl:message terminate="no">translation.oox2odf.shape.3dEffect</xsl:message>
  </xsl:template>

  <!-- 
  *************************************************************************
  CALLED TEMPLATES
  *************************************************************************
  -->

  <!--
  Summary: Inserts the Start color of a gradiant fill
  Author: makz (DIaLOGIKa)
  Date: 6.11.2007
  -->
  <xsl:template name="InsertStartColor">
    <xsl:variable name="parentShape" select="parent::node()" />
    <xsl:choose>
    <xsl:when test="@color">
      <xsl:call-template name="InsertColor">
        <xsl:with-param name="color" select="@color"/>
      </xsl:call-template>
    </xsl:when>
      <xsl:when test="$parentShape/@fillcolor">
        <xsl:call-template name="InsertColor">
          <xsl:with-param name="color" select="$parentShape/@fillcolor"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:text>#ffffff</xsl:text>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--
  Summary: Inserts the End color of a gradiant fill
  Author: makz (DIaLOGIKa)
  Date: 6.11.2007
  -->
  <xsl:template name="InsertEndColor">
    <xsl:variable name="parentShape" select="parent::node()" />
    <xsl:choose>
      <xsl:when test="@color2">
        <xsl:call-template name="InsertColor">
          <xsl:with-param name="color" select="@color2"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:text>#ffffff</xsl:text>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!--
  Summary: Inserts a draw:frame
  Author: makz (DIaLOGIKa)
  Date: 30.10.2007
  -->
  <xsl:template name="InsertFrame">
    <xsl:param name="framePr" />

    <draw:frame>
      <xsl:attribute name="draw:style-name">
        <xsl:value-of select="generate-id($framePr)"/>
      </xsl:attribute>
      <xsl:call-template name="InsertAnchorType">
        <xsl:with-param name="shape" select="$framePr"/>
      </xsl:call-template>
      <xsl:call-template name="InsertShapeWidth">
        <xsl:with-param name="shape" select="$framePr"/>
      </xsl:call-template>
      <xsl:call-template name="InsertShapeHeight">
        <xsl:with-param name="shape" select="$framePr"/>
      </xsl:call-template>
      <xsl:call-template name="InsertshapeAbsolutePos">
        <xsl:with-param name="shape" select="$framePr"/>
      </xsl:call-template>
      <xsl:call-template name="InsertShapeAutomaticWidth">
        <xsl:with-param name="shape" select="$framePr" />
      </xsl:call-template>
      <draw:text-box>
        <xsl:call-template name="InsertTextBoxAutomaticHeight">
          <xsl:with-param name="shape" select="$framePr"/>
        </xsl:call-template>
        <xsl:call-template name="InsertParagraphToFrame"/>
      </draw:text-box>
    </draw:frame>
  </xsl:template>
  
  <!--
  makz: I commented this template in because it caused trouble.
  Not all styles had the same name as their elements which reference to
  
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
  -->

  <xsl:template name="InsertShapeStyleProperties">
    <xsl:param name="shape" select="." />
    
      <!--<xsl:if test="parent::node()[name()='v:group']">
         TO DO v:shapes in v:group 
      </xsl:if>-->
      <xsl:call-template name="InsertShapeWrap">
        <xsl:with-param name="shape" select="$shape" />
      </xsl:call-template>
      <xsl:call-template name="InsertShapeFromTextDistance">
        <xsl:with-param name="shape" select="$shape" />
      </xsl:call-template>
      <xsl:call-template name="InsertShapeBorders">
        <xsl:with-param name="shape" select="$shape" />
      </xsl:call-template>
      <xsl:call-template name="InsertShapeAutomaticWidth">
        <xsl:with-param name="shape" select="$shape" />
      </xsl:call-template>
      <xsl:call-template name="InsertShapeHorizontalPos">
        <xsl:with-param name="shape" select="$shape" />
      </xsl:call-template>
      <xsl:call-template name="InsertShapeVerticalPos">
        <xsl:with-param name="shape" select="$shape" />
      </xsl:call-template>
      <xsl:call-template name="InsertShapeFlowWithText">
        <xsl:with-param name="shape" select="$shape" />
      </xsl:call-template>
      <xsl:call-template name="InsertShapeBackgroundColor">
        <xsl:with-param name="shape" select="$shape" />
      </xsl:call-template>
      <xsl:call-template name="InsertShapeZindex">
        <xsl:with-param name="shape" select="$shape" />
      </xsl:call-template>
      <xsl:call-template name="InsertShapeWrappedParagraph">
        <xsl:with-param name="shape" select="$shape" />
      </xsl:call-template>
      <xsl:call-template name="InsertShapeShadow">
        <xsl:with-param name="shape" select="$shape" />
      </xsl:call-template>
    <xsl:for-each select="$shape/v:textbox">
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

  <xsl:template name="InsertShapeShadow">
    <xsl:for-each select="v:shadow">
      <xsl:attribute name="style:shadow">
        <xsl:choose>
          <xsl:when test="@on = 'false' or @on = 'f' or @on = '0'">none</xsl:when>
          <xsl:otherwise>
            <!-- report lost attributes -->
            <xsl:if test="@opacity">
              <xsl:message terminate="no">translation.oox2odf.shape.shadow</xsl:message>
            </xsl:if>
            <xsl:if test="@obscured">
              <xsl:message terminate="no">translation.oox2odf.shape.shadow.obscurity</xsl:message>
            </xsl:if>
            <xsl:if test="@type">
              <!-- TODO is this a copy & paste error? -->
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

  <!--
  Summary: writes the background and fill color of a shape
  Author: Clever Age
  Modified: makz (DIaLOGIKa)
  Date: 6.11.2007
  -->
  <xsl:template name="InsertShapeBackgroundColor">
    <xsl:param name="shape" select="."/>
    <xsl:variable name="fillcolor">
      <xsl:choose>
        <xsl:when test="$shape/@fillcolor and not($shape/@fillcolor='') and not($shape/@fillcolor='window') and not($shape/@fillcolor='gradient')">
          <xsl:value-of select="$shape/@fillcolor"/>
        </xsl:when>
        <xsl:otherwise>#ffffff</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="isFilled" select="$shape/@filled"/>

    <!--dialogika, clam: if filled is set to false, make it transparent (as Word does; bugfix #1800779)-->
    <xsl:if test="$isFilled = 'f'">
      <xsl:attribute name="style:background-transparency">100%</xsl:attribute>
    </xsl:if>
    
    <!-- Insert background-color -->
    <xsl:if test="(not($isFilled) or $isFilled != 'f') and $fillcolor != ''">
      <xsl:attribute name="fo:background-color">
        <xsl:call-template name="InsertColor">
          <xsl:with-param name="color" select="$fillcolor"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="(not($isFilled) or $isFilled != 'f') and ($fillcolor = '' or not($fillcolor))">
      <xsl:attribute name="fo:background-color">
        <xsl:call-template name="InsertColor">
          <xsl:with-param name="color">#ffffff</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    
    <!-- insert fill-color -->
    <xsl:attribute name="draw:fill-color">
      <xsl:call-template name="InsertColor">
        <xsl:with-param name="color" select="$fillcolor"/>
      </xsl:call-template>
    </xsl:attribute>

    <!-- If the shape has a gradient fill -->
    <xsl:if test="$shape/v:fill[@type='gradient']">
      <xsl:attribute name="draw:fill">
        <xsl:text>gradient</xsl:text>
      </xsl:attribute>
      <xsl:attribute name="draw:fill-gradient-name">
        <xsl:value-of select="concat('Gradient_', generate-id($shape/v:fill))"/>
      </xsl:attribute>
    </xsl:if>
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

    <xsl:variable name="pos">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="propertyName" select="'position'"/>
        <xsl:with-param name="shape" select="$shape"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="horizontalPos">
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
    <xsl:variable name="horizontalRel">
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
    
    <xsl:choose>
      <xsl:when test="$shape/@o:hralign and $shape/@o:hr='t'">
        <xsl:call-template name="InsertGraphicPosH">
          <xsl:with-param name="align" select="$shape/@o:hralign"/>
        </xsl:call-template>
      </xsl:when>
      <!--
      makz: changed due to problem with graph images in file ESTAT_SIF_FR.doc
      in compatibility mode
      <xsl:when test="$horizontalPos='' and ancestor::w:p/w:pPr/w:jc/@w:val and $horizontalRel!='page'">
      -->
      <xsl:when test="$horizontalPos='' and ancestor::w:p/w:pPr/w:jc/@w:val and $horizontalRel!='page' 
                and $pos!='absolute'">
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
      <xsl:with-param name="relativeFrom" select="$horizontalRel" />
      <xsl:with-param name="hPos" select="$horizontalPos" />
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
    <xsl:call-template name="InsertVerticalPos">
      <xsl:with-param name="vAlign" select="$verticalPos"/>
    </xsl:call-template>
    <xsl:call-template name="InsertVerticalRel">
      <xsl:with-param name="vRel" select="$verticalRelative"/>
      <xsl:with-param name="vPos" select="$verticalPos" />
    </xsl:call-template>
    <!--
    <xsl:call-template name="InsertGraphicPosV">
      <xsl:with-param name="align" select="$verticalPos"/>
      <xsl:with-param name="relativeFrom" select="$verticalRelative"/>
    </xsl:call-template>
    <xsl:call-template name="InsertGraphicPosRelativeV">
      <xsl:with-param name="relativeFrom" select="$verticalRelative"/>
      <xsl:with-param name="align" select="$verticalPos"/>
    </xsl:call-template>
    -->
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
          <xsl:with-param name="wrap" select="$shape/w10:wrap"/>
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
        <xsl:with-param name="shape" select="$shape"/>
      </xsl:call-template>
    </xsl:variable>

    <!--
    makz: style:run-through="background" must always be written for negative z-index
    -->
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

  <!--
  Summary: Template writes the properties of a shape border.
  Author: Clever Age
  Modified: makz (DIaLOGIKa)
  Date: 16.10.2007
  -->
  <xsl:template name="InsertShapeBorders">
    <!-- current context is <v:shape> -->
    <xsl:param name="shape" select="."/>
    <xsl:variable name="typeId" select="substring-after($shape/@type, '#')" />
    <xsl:variable name="shapetype" select="key('Part', 'word/document.xml')/w:document/w:body//v:shapetype[@id=$typeId]" />

    <xsl:variable name="paintBorder">
      <!-- The stroked attribute of the shape is stronger than the attribute of the shapetype -->
      <xsl:choose>
        <xsl:when test="$shape/@stroked">
          <xsl:choose>
            <!-- there is no v:stroke element, then only paint the border if stroked is set to true -->
            <xsl:when test="not($shape/v:stroke) and ($shape/@stroked='t' or $shape/@stroked='true' or $shape/@stroked='1')">
              <xsl:text>shape</xsl:text>
            </xsl:when>
            <!-- there is a v:stroke element, then paint the border if stroked isn't disabled -->
            <xsl:when test="$shape/v:stroke and ($shape/@stroked!='f' or $shape/@stroked!='false' or $shape/@stroked='0')">
              <xsl:text>shape</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>none</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="$shapetype/@stroked">
          <xsl:choose>
            <!-- there is no v:stroke element, then only paint the border if stroked is enabled -->
            <xsl:when test="not($shapetype/v:stroke) and ($shapetype/@stroked='t' or $shapetype/@stroked='true' or $shapetype/@stroked='1')">
              <xsl:text>shapetype</xsl:text>
            </xsl:when>
            <!-- there is a v:stroke element, then paint the border if stroked isn't disabled -->
            <xsl:when test="$shapetype/v:stroke and ($shapetype/@stroked!='f' and $shapetype/@stroked!='false' or $shapetype/@stroked='0')">
              <xsl:text>shapetype</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>none</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <!-- 
          if no stroked attribute is set but a v:stroke element exists word paint a default border
          -->
          <xsl:choose>
            <xsl:when test="$shape/v:stroke or $shapetype/v:stroke">
              <xsl:text>default</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>none</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose> 
    </xsl:variable>
    
    <!-- 
    Shapes can have stroke attributes and a stroke element without having a border.
    In this case the stroked attribute is set to false.
    -->
    <xsl:choose>
      <xsl:when test="not($paintBorder='none')">
        
        <!-- calculate border color -->
        <xsl:variable name="borderColor">
          <xsl:choose>
            <xsl:when test="$paintBorder='shape' and $shape/@strokecolor">
              <xsl:call-template name="InsertColor">
                <xsl:with-param name="color" select="$shape/@strokecolor"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="$paintBorder='shapetype' and $shapetype/@strokecolor">
              <xsl:call-template name="InsertColor">
                <xsl:with-param name="color" select="$shapetype/@strokecolor"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>#000000</xsl:text>
            </xsl:otherwise>
            <!--
            <xsl:when test="$paintBorder='default'">
              <xsl:text>#000000</xsl:text>
            </xsl:when>
            -->
          </xsl:choose>
        </xsl:variable>

        <!-- calculate border weight -->
        <xsl:variable name="borderWeight">
          <xsl:choose>
            <xsl:when test="$paintBorder='shape' and $shape/@strokeweight">
              <xsl:call-template name="ConvertMeasure">
                <xsl:with-param name="length" select="$shape/@strokeweight"/>
                <xsl:with-param name="destUnit" select="'cm'"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="$paintBorder='shapetype' and $shapetype/@strokeweight">
              <xsl:call-template name="ConvertMeasure">
                <xsl:with-param name="length" select="$shapetype/@strokeweight"/>
                <xsl:with-param name="destUnit" select="'cm'"/>
              </xsl:call-template>
            </xsl:when>
            <!--
            <xsl:when test="$paintBorder='default'">
              <xsl:text>0.0176cm</xsl:text>
            </xsl:when>
            -->
            <xsl:otherwise>
              <xsl:text>0.0176cm</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:variable name="dashStyle" select="$shape/v:stroke/@dashstyle">
        </xsl:variable>

        <xsl:variable name="lineStyle" select="$shape/v:stroke/@linestyle">
        </xsl:variable>

        <!-- get border style -->
        <xsl:variable name="borderStyle">
          <xsl:choose>
            <xsl:when test="$lineStyle">
              <xsl:text>double</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>solid</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <!-- write attributes -->
        <xsl:attribute name="fo:border">
          <xsl:value-of select="concat($borderWeight,' ',$borderStyle,' ',$borderColor)"/>
        </xsl:attribute>

        <xsl:attribute name="svg:stroke-color">
          <xsl:value-of select="$borderColor"/>
        </xsl:attribute>

        <!-- the border is double -->
        <xsl:if test="$lineStyle">
          <xsl:attribute name="style:border-line-width">
            <xsl:choose>
              <xsl:when test="$lineStyle='thinThin' or $lineStyle='thickBetweenThin'">
                <xsl:value-of select="concat(substring-before($borderWeight,'cm')*0.45 ,'cm',' ',substring-before($borderWeight,'cm')*0.1,'cm ', substring-before($borderWeight,'cm')*0.45,'cm')"/>
              </xsl:when>
              <xsl:when test="$lineStyle='thinThick'">
                <xsl:value-of select="concat(substring-before($borderWeight,'cm')*0.7,'cm',' ',substring-before($borderWeight,'cm')*0.1,'cm ', substring-before($borderWeight,'cm')*0.2,'cm')"/>
              </xsl:when>
              <xsl:when test="$lineStyle='thickThin'">
                <xsl:value-of select="concat(substring-before($borderWeight,'cm')*0.2,'cm',' ',substring-before($borderWeight,'cm')*0.1,'cm ', substring-before($borderWeight,'cm')*0.7,'cm')"/>
              </xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>

        <!-- the border is dashed -->
        <xsl:if test="$dashStyle">
          <xsl:attribute name="draw:stroke">
            <xsl:text>dash</xsl:text>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <!-- insert explicit NO border (only if the element is no line) -->
      <xsl:when test="not(name($shape)='v:line')">
        <xsl:attribute name="draw:stroke">
          <xsl:text>none</xsl:text>
        </xsl:attribute>
        <xsl:attribute name="fo:border">
          <xsl:text>none</xsl:text>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertTextBoxAutomaticHeight">
    <xsl:param name="textbox" select="."/>

    <xsl:variable name="fitToText">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$textbox"/>
        <xsl:with-param name="propertyName" select="'mso-fit-shape-to-text'"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$fitToText='t' or $fitToText='true' or (not($textbox/@w:h) and $textbox/@hRule!='exact')">
      <xsl:attribute name="fo:min-height">
        <xsl:choose>
          <xsl:when test="$textbox/@hRule='atLeast'">
            <xsl:value-of select="$textbox/@hRule"/>
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
    <xsl:variable name="fitToText">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$shape/v:textbox"/>
        <xsl:with-param name="propertyName" select="'mso-fit-shape-to-text'"/>
      </xsl:call-template>
    </xsl:variable>
    
    <xsl:choose>
      <xsl:when test="$fitToText='t' or  $fitToText='true' or ($wrapStyle!='' and $wrapStyle='none') or ($shape/@w:wrap and $shape/@w:wrap != 'none')">
        <xsl:attribute name="fo:min-width">
          <xsl:text>0cm</xsl:text>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <!--
  Summary: Inserts the paragraphs in a frame
  Author: Clever Age
  -->
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
    
    <!-- if there is another paragraph with framePr -->
    <xsl:variable name="followingP" select="$paragraph/following-sibling::*[1]" />
    <xsl:if test="$followingP/w:pPr/w:framePr">
      <!-- check if the folowing framePr is identically -->
      <xsl:variable name="identically">
        <xsl:call-template name="CompareFrames">
          <xsl:with-param name="frame1" select="$paragraph/w:pPr/w:framePr"/>
          <xsl:with-param name="frame2" select="$followingP/w:pPr/w:framePr"/>
        </xsl:call-template>
      </xsl:variable>
      <!-- then merge the two paragraphs -->
      <xsl:if test="$identically='true'">
        <xsl:call-template name="InsertParagraphToFrame">
          <xsl:with-param name="paragraph" select="$followingP"/>
        </xsl:call-template>
      </xsl:if>
    </xsl:if>
  </xsl:template>

  <!--
  Summary: Inserts the attribute for vertical position
  Author: makz (DIaLOGIKa)
  Date: 15.11.2007
  -->
  <xsl:template name="InsertVerticalPos">
    <xsl:param name="yAlign" />
    <xsl:param name="y" />

    <xsl:attribute name="style:vertical-pos">
      <xsl:choose>
        <xsl:when test="$yAlign='bottom'">
          <xsl:text>bottom</xsl:text>
        </xsl:when>
        <xsl:when test="$yAlign='top'">
          <xsl:choose>
            <!--
              makz: If there is a y coordinate then OOo needs "from-top" 
              because "top" cannot have coordinates.
              -->
            <xsl:when test="$y">
              <xsl:text>from-top</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>top</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="$yAlign='center'">
          <xsl:text>middle</xsl:text>
        </xsl:when>
        <!--
          makz: If no yAlign is specified in OOX document use "from-top"
          -->
        <xsl:otherwise>
          <xsl:text>from-top</xsl:text>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>

  <!--
  Summary: Inserts the attribute for horizontal position
  Author: makz (DIaLOGIKa)
  Date: 15.11.2007
  -->
  <xsl:template name="InsertHorizontalPos">
    <xsl:param name="xAlign" />
    
    <xsl:attribute name="style:horizontal-pos">
      <xsl:choose>
        <xsl:when test="$xAlign='center'">
          <xsl:text>middle</xsl:text>
        </xsl:when>
        <xsl:when test="$xAlign='inside'">
          <xsl:text>left</xsl:text>
        </xsl:when>
        <xsl:when test="$xAlign='outside'">
          <xsl:text>right</xsl:text>
        </xsl:when>
        <xsl:when test="$xAlign='left'">
          <xsl:text>left</xsl:text>
        </xsl:when>
        <xsl:when test="$xAlign='right'">
          <xsl:text>right</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:text>from-left</xsl:text>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute >
  </xsl:template>

  <!--
  Summary: Inserts the attribute for horizontal relation
  Author: makz (DIaLOGIKa)
  Date: 15.11.2007
  -->
  <xsl:template name="InsertHorizontalRel">
    <xsl:param name="hRel" />
    <xsl:param name="hPos" />
    
    <xsl:attribute name="style:horizontal-rel">
      <xsl:choose>
        <xsl:when test="$hRel='margin'">
          <xsl:text>paragraph-content</xsl:text>
        </xsl:when>
        <xsl:when test="$hRel='page'">
          <xsl:text>page</xsl:text>
        </xsl:when>
        <xsl:when test="$hRel='text'">
          <xsl:text>page-content</xsl:text>
        </xsl:when>
        <xsl:when test="$hRel=''">
          <!-- 
          if no relation is set, Word uses default values, 
          depeding on the position
          -->
          <xsl:choose>
            <!-- 
            If no position is set, it is absolute positioning.
            In this case the default relation is the paragraph-content 
            -->
            <xsl:when test="$hPos=''">
              <xsl:text>paragraph-content</xsl:text>
            </xsl:when>
            <!-- 
            If a position is set, it is relative positioning.
            In this case the default relation is the page-content 
            -->
            <xsl:otherwise>
              <xsl:text>page-content</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>

  <!--
  Summary: Inserts the attribute for vertical relation
  Author: makz (DIaLOGIKa)
  Date: 15.11.2007
  -->
  <xsl:template name="InsertVerticalRel">
    <xsl:param name="vRel" />
    <xsl:param name="vPos" />
    
    <xsl:attribute name="style:vertical-rel">
      <xsl:choose>
        <xsl:when test="$vRel='page'">
          <xsl:text>page</xsl:text>
        </xsl:when>
        <xsl:when test="$vRel='text'">
          <xsl:text>baseline</xsl:text>
        </xsl:when>
        <xsl:when test="$vRel=''">
          <!-- 
          if no relation is set, Word uses default values, 
          depeding on the position
          -->
          <xsl:choose>
            <!-- 
            If no position is set, it is absolute positioning.
            In this case the default relation is the paragraph-content 
            -->
            <xsl:when test="$vPos=''">
              <xsl:text>paragraph-content</xsl:text>
            </xsl:when>
            <!-- 
            If a position is set, it is relative positioning.
            In this case the default relation is the page-content 
            -->
            <xsl:otherwise>
              <xsl:text>page-content</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>

  <!--
  Summary: Inserts the attribute for wrap
  Author: makz (DIaLOGIKa)
  Date: 15.11.2007
  -->
  <xsl:template name="InsertWrap">
    <xsl:param name="wrap" />
               
    <xsl:attribute name="style:wrap">
      <xsl:choose>
        <xsl:when test="not($wrap) or $wrap ='none'">
          <xsl:text>none</xsl:text>
        </xsl:when>
        <xsl:when test="$wrap ='auto'">
          <xsl:text>parallel</xsl:text>
        </xsl:when>
        <xsl:when test="$wrap ='around'">
          <xsl:text>parallel</xsl:text>
        </xsl:when>
        <xsl:when test="$wrap ='notBeside'">
          <xsl:text>parallel</xsl:text>
        </xsl:when>
        <xsl:when test="$wrap ='through'">
          <xsl:text>parallel</xsl:text>
        </xsl:when>
        <xsl:when test="$wrap ='tight'">
          <xsl:text>parallel</xsl:text>
        </xsl:when>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>
    
  <!--
  Summary: Converts the properties of the Anchor.
  Author: Clever Age
  Modified: makz (DIaLOGIKa)
  Date: 29.10.2007
  -->
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

    <xsl:call-template name="InsertWrap">
      <xsl:with-param name="wrap" select="$Wrap" />
    </xsl:call-template>
    
    <xsl:if test ="count($yAlign)>0 or count($vAnchor)>0">
      <xsl:call-template name="InsertVerticalPos">
        <xsl:with-param name="yAlign" select="$yAlign" />
        <xsl:with-param name="y" select="$oFramePr/@w:y" />
      </xsl:call-template>
      <xsl:call-template name="InsertVerticalRel">
        <xsl:with-param name="vRel" select="$vAnchor" />
        <xsl:with-param name="vPos" select="$yAlign" />
      </xsl:call-template>
    </xsl:if>
    <xsl:if test ="count($xAlign)>0 or count($hAnchor)>0">
      <xsl:call-template name="InsertHorizontalPos">
        <xsl:with-param name="xAlign" select="$xAlign" />
      </xsl:call-template>
      <xsl:call-template name="InsertHorizontalRel">
        <xsl:with-param name="hRel" select="$hAnchor" />
        <xsl:with-param name="hPos" select="$xAlign" />
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

  <!--
  Summary: Compares 2 OOX frames based on their attributes
  Author: makz (DIaLOGIKa)
  Date: 29.10.2007
  -->
  <xsl:template name="CompareFrames">
    <xsl:param name="frame1" />
    <xsl:param name="frame2" />

    <xsl:choose>
      <xsl:when test="string($frame1/@w:anchorLock) = string($frame2/@w:anchorLock) and 
                      string($frame1/@w:dropCap) = string($frame2/@w:dropCap) and 
                      string($frame1/@w:h) = string($frame2/@w:h) and 
                      string($frame1/@w:hAnchor) = string($frame2/@w:hAnchor) and 
                      string($frame1/@w:hRule) = string($frame2/@w:hRule) and 
                      string($frame1/@w:hSpace) = string($frame2/@w:hSpace) and 
                      string($frame1/@w:lines) = string($frame2/@w:lines) and            
                      string($frame1/@w:vAnchor) = string($frame2/@w:vAnchor) and 
                      string($frame1/@w:vSpace) = string($frame2/@w:vSpace) and 
                      string($frame1/@w:w) = string($frame2/@w:w) and 
                      string($frame1/@w:wrap) = string($frame2/@w:wrap) and 
                      string($frame1/@w:x) = string($frame2/@w:x) and 
                      string($frame1/@w:xAlign) = string($frame2/@w:xAlign) and 
                      string($frame1/@w:y) = string($frame2/@w:y) and 
                      string($frame1/@w:yAlign) = string($frame2/@w:yAlign)">
        <xsl:text>true</xsl:text>
      </xsl:when>
      <xsl:otherwise>
        <xsl:text>false</xsl:text>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!--
  Summary: Inserts the common properties of a v:shape element
  Author: Clever Age
  Modified: makz (DIaLOGIKa)
  Date: 15.11.2007
  -->
  <xsl:template name="InsertCommonShapeProperties">
    <xsl:param name="shape" />
    <xsl:call-template name="InsertShapeStyleName">
      <xsl:with-param name="shape" select="$shape" />
    </xsl:call-template>
    <xsl:call-template name="InsertAnchorType">
      <xsl:with-param name="shape" select="$shape" />
    </xsl:call-template>
    <xsl:call-template name="InsertShapeWidth">
      <xsl:with-param name="shape" select="$shape" />
    </xsl:call-template>
    <xsl:call-template name="InsertShapeHeight">
      <xsl:with-param name="shape" select="$shape" />
    </xsl:call-template>
    <xsl:call-template name="InsertshapeAbsolutePos">
      <xsl:with-param name="shape" select="$shape" />
    </xsl:call-template>
  </xsl:template>

  <!--
    Summary: Inserts position of first end of line 
    Author:  Tomasz Mueller (Clever Age)
    Date: 29.10.2007
  -->
  <xsl:template name="InsertLinePos1">
    <xsl:param name="flip"/>
    <xsl:attribute name="svg:x1">
      <xsl:call-template name="ConvertMeasure">
        <xsl:with-param name="length">
          <xsl:choose>
            <xsl:when test="contains($flip,'x')">
              <xsl:value-of select="substring-before(@to,',')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-before(@from,',')"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
        <xsl:with-param name="destUnit" select="'cm'"/>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="svg:y1">
      <xsl:call-template name="ConvertMeasure">
        <xsl:with-param name="length">
          <xsl:choose>
            <xsl:when test="contains($flip,'y')">
              <xsl:value-of select="substring-after(@to,',')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-after(@from,',')"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
        <xsl:with-param name="destUnit" select="'cm'"/>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>
  
  <!--
    Summary: Inserts position of second end of line 
    Author:  Tomasz Mueller (Clever Age)
    Date: 29.10.2007
  -->
  <xsl:template name="InsertLinePos2">
    <xsl:param name="flip"/>
    
    <xsl:attribute name="svg:x2">
      <xsl:call-template name="ConvertMeasure">
        <xsl:with-param name="length">
          <xsl:choose>
            <xsl:when test="contains($flip,'x')">
              <xsl:value-of select="substring-before(@from,',')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-before(@to,',')"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
        <xsl:with-param name="destUnit" select="'cm'"/>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="svg:y2">
      <xsl:call-template name="ConvertMeasure">
        <xsl:with-param name="length">
          <xsl:choose>
            <xsl:when test="contains($flip,'y')">
              <xsl:value-of select="substring-after(@from,',')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-after(@to,',')"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
        <xsl:with-param name="destUnit" select="'cm'"/>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>
  
  <!--
  Summary: inserts the svg:x and svg:y attribute
  Author: Clever Age
  Modified: makz (DIaLOGIKa)
  Date: 21.11.2007
  -->
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
    <xsl:variable name="relH">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$shape"/>
        <xsl:with-param name="propertyName" select="'mso-position-horizontal-relative'"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="relV">
      <xsl:call-template name="GetShapeProperty">
        <xsl:with-param name="shape" select="$shape"/>
        <xsl:with-param name="propertyName" select="'mso-position-vertical-relative'"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="HorizontalWidth">
      <xsl:choose>
        <xsl:when test="$relH='page'">
          <xsl:value-of select="number(ancestor::node()/w:sectPr/w:pgSz/@w:w)"/>
        </xsl:when>
        <xsl:when test="$relH='margin'">
          <xsl:value-of select="number(ancestor::node()/w:sectPr/w:pgSz/@w:w) - number(ancestor::node()/w:sectPr/w:pgMar/@w:left) - number(ancestor::node()/w:sectPr/w:pgMar/@w:right)"/>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>

    <!-- 
    Don't insert svg:x and svg:y if the shape is part of the group.
    This workaround must be removed if group positioning is implemented.
    Wrong x and y values of elements inside of a group causes Open Office to crash.
    
    See bug: 1747143
    -->
    <xsl:if test="name($shape/..)!='v:group'">

      <xsl:if test="$position='absolute' or $shape[name()='w:framePr']">
        <xsl:variable name="x">
          <xsl:choose>
            <xsl:when test="$shape[name()='w:framePr']">
              <xsl:value-of select="$shape/@w:x"/>
            </xsl:when>
            <xsl:when test="$LeftPercent  &gt; 0">
              <xsl:value-of select="$HorizontalWidth * $LeftPercent  div 1000"/>
            </xsl:when>
            <xsl:when test="$shape[name()='v:rect' or name()='v:line']">
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
        <xsl:variable name="y">
          <xsl:choose>
            <xsl:when test="$shape[name()='w:framePr']">
              <xsl:value-of select="$shape/@w:y"/>
            </xsl:when>
            <xsl:when test="$shape[name()='v:rect' or name()='v:line']">
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
        <xsl:variable name="posX">
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="$x"/>
            <xsl:with-param name="destUnit" select="'cm'"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="posY">
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="$y"/>
            <xsl:with-param name="destUnit" select="'cm'"/>
          </xsl:call-template>
        </xsl:variable>

        <!-- 
        Write attributes .
        Calculate the values relative to the group's values 
        -->
        <xsl:attribute name="svg:x">
          <xsl:value-of select="$posX"/>
        </xsl:attribute>
        <xsl:attribute name="svg:y">
          <xsl:value-of select="$posY"/>
        </xsl:attribute>
      </xsl:if>

    </xsl:if>

  </xsl:template>

  <!--
  Summary: Writes the anchor-type attribute
  Author: Clever Age
  Modified: makz (DIaLOGIKa)
  Date: 19.10.2007
  -->
  <xsl:template name="InsertAnchorType">
    <xsl:param name="shape" select="."/>
    
    <!-- 
    don't write anchor-type if the shape is part of a group.
    grouped elements are always anchored by their group.
    -->
    <xsl:if test="name($shape/..)!='v:group'">
      <xsl:attribute name="text:anchor-type">
        <xsl:choose>

          <!-- 
        makz: default for word inline shape elements.
        If no wrap is definied the style does not contain position informations
        Fix #1786094
        -->
          <xsl:when test="not(name($shape)='w:framePr') and not(w10:wrap) and not(contains($shape/@style, 'position'))">
            <xsl:text>as-char</xsl:text>
          </xsl:when>
          <xsl:when test="w10:wrap/@type='none'">
            <xsl:text>as-char</xsl:text>
          </xsl:when>

          <!-- 
        anchor should be 'as-char' in some particular cases.
        Explanation of test :
        1. if no wrap is defined (default in OOX is inline with text)
        2. if there is another run in current paragraph OR we are in a table cell
        3. No absolute position is defined OR wrapping is explicitely set to 'as-char'

        JP 27/08/2007 
        Fix #1666915 if position is absolute must not be set as-char
        -->
          <xsl:when
            test="not(w10:wrap) 
          and (count(ancestor::w:r[1]/parent::node()/w:r) &gt; 1  or (not(count(ancestor::w:r[1]/parent::node()/w:r) &gt; 1) and ancestor::w:tc)) 
          and (not(contains($shape/@style, 'position:absolute')))
          and (contains($shape/@style, 'mso-position-horizontal-relative:text') or contains($shape/@style, 'mso-position-vertical-relative:text') or ($shape/@vAnchor='text') or ($shape/@hAnchor='relative') )">
            <xsl:text>as-char</xsl:text>
          </xsl:when>

          <!--
        makz: anchors in headers and footer which are not not as-char (if the previous conditions failed) 
        must be paragraph because because the frame/shape would no longer be in header if the 
        anchor is not set to paragraph or as-char
        -->
          <xsl:when test="(ancestor::w:hdr or ancestor::w:ftr) and (w10:wrap/@type!='' or $shape/@w:wrap!='')">
            <!-- Warning Messages 
          <xsl:if test="ancestor::w:hdr">
            <xsl:message terminate="no">translation.oox2odf.frame.inHeader</xsl:message>
          </xsl:if>
          <xsl:if test="ancestor::w:ftr">
            <xsl:message terminate="no">translation.oox2odf.frame.inFooter</xsl:message>
          </xsl:if>
          -->
            <xsl:text>paragraph</xsl:text>
          </xsl:when>

          <xsl:when test="(w10:wrap/@anchorx='page' and w10:wrap/@anchory='page') or ($shape/@w:hAnchor='page' and $shape/@w:vAnchor='page')">
            <xsl:text>page</xsl:text>
          </xsl:when>

          <xsl:otherwise>
            <xsl:text>paragraph</xsl:text>
          </xsl:otherwise>

        </xsl:choose>
      </xsl:attribute>
    </xsl:if>
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
      <xsl:value-of select="generate-id($shape)"/>
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
        <xsl:if test="not(contains($shape/v:textbox/@style, 'mso-fit-shape-to-text:t')) or $shape/@w:h">
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

        <!-- Don't insert the width if the textbox is set to auto-width -->
        <xsl:if test="not(contains($shape/v:textbox/@style, 'mso-fit-shape-to-text:t'))">
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
        </xsl:if>
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
  
</xsl:stylesheet>

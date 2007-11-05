<?xml version="1.0" encoding="UTF-8"?>
<!--
 * Copyright (c) 2006, Clever Age
 * All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *  
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of Clever Age nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software
 *       without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE REGENTS AND CONTRIBUTORS ``AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL THE REGENTS AND CONTRIBUTORS BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
  -->
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  exclude-result-prefixes="xlink draw svg fo office style text">

  <xsl:key name="images"
    match="draw:frame[not(./draw:object-ole or ./draw:object)]/draw:image[@xlink:href]" use="''"/>
  <xsl:key name="frames" match="draw:frame" use="''"/>
  <xsl:key name="ole-objects" match="draw:frame[./draw:object-ole] " use="''"/>

  <!-- 
  *************************************************************************
  MATCHING TEMPLATES
  *************************************************************************
  -->
  
  <!-- 
  Summary: converts frames
  Author: Clever Age
  -->
  <xsl:template match="draw:frame" mode="paragraph">
    <!-- insert link to TOC field when required (user indexes) -->
    <xsl:call-template name="InsertTCField"/>
    <xsl:call-template name="InsertEmbeddedTextboxes"/>
  </xsl:template>

  <!--
  Summary: OLE objects in frames 
  Author: Clever Age
  -->
  <xsl:template match="draw:frame[./draw:object-ole]" mode="paragraph">
    <!-- insert link to TOC field when required (user indexes) -->
    <xsl:call-template name="InsertTCField"/>

    <xsl:variable name="imageId">
      <xsl:call-template name="GetPosition"/>
    </xsl:variable>
    <xsl:variable name="width">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length" select="@svg:width"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="height">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length" select="@svg:height"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="OleObjType">
      <xsl:if test="draw:object-ole/@xlink:show='embed'">Embed</xsl:if>
    </xsl:variable>

    <w:r>
      <w:object>
        <v:shape id="{$imageId}" type="#_x0000_t75" style="width:{$width}pt;height:{$height}pt"
          o:ole="" filled="t">
          <v:fill color2="black"/>
          <v:imagedata r:id="{generate-id(./draw:image)}" o:title=""/>
        </v:shape>
        <o:OLEObject Type="{$OleObjType}" ProgID="" ShapeID="{$imageId}" DrawAspect="Content"
          ObjectID="" r:id="{generate-id(draw:object-ole)}"/>
      </w:object>
    </w:r>
  </xsl:template>

  <!-- 
    embedd all consecutive frames that are not inserted into a paragraph in a single paragraph
    (avoid paragraph not present in original document)
  -->
  <xsl:template match="node()[contains(name(), 'draw:') and parent::office:text]">
    <!-- concerned elements : draw:custom-shape, draw:rect, draw:ellipse, draw:frame[ole-object|image|text-box] -->
    <xsl:choose>
      <xsl:when test="following-sibling::text:p">
        <!-- do nothing : handled by the first paragraph -->
      </xsl:when>
      <xsl:otherwise>
        <w:p>
          <xsl:choose>
            <xsl:when test="self::draw:frame">
              <xsl:apply-templates select="." mode="paragraph"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:apply-templates select="." mode="shapes"/>
            </xsl:otherwise>
          </xsl:choose>
          <xsl:apply-templates select="following-sibling::node()[1][contains(name(), 'draw:')]"/>
        </w:p>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- 
  Summary: frames containing external images
  Author: Clever Age
  -->
  <xsl:template match="draw:frame[not(./draw:object-ole or ./draw:object) and ./draw:image[@xlink:href and not(starts-with(@xlink:href, 'Pictures/'))]]" mode="paragraph">
    <!-- insert link to TOC field when required (user indexes) -->
    <xsl:call-template name="InsertTCField"/>

    <xsl:variable name="supported">
      <xsl:call-template name="image-support">
        <xsl:with-param name="name" select="./draw:image/@xlink:href"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$supported = 'true'">
      <xsl:variable name="imageId">
        <xsl:call-template name="GetPosition"/>
      </xsl:variable>

      <xsl:for-each select="draw:image">
        <w:r>
          <w:pict>
            <v:shape id="{$imageId}" type="#_x0000_t75">

              <xsl:variable name="styleName" select=" parent::draw:frame/@draw:style-name"/>
              <xsl:variable name="automaticStyle" select="key('automatic-styles', parent::draw:frame/@draw:style-name)"/>
              <xsl:variable name="officeStyle" select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $styleName]"/>

              <xsl:variable name="shapeStyle" select="$automaticStyle | $officeStyle"/>

              <!-- shape properties: size, z-index, coordinates, position, margin etc -->
              <xsl:call-template name="InsertShapeProperties">
                <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
                <xsl:with-param name="shapeProperties" select="parent::draw:frame"/>
              </xsl:call-template>

              <v:imagedata r:id="{generate-id(.)}" o:title=""/>

              <!-- wrapping -->
              <xsl:call-template name="InsertShapeWrap">
                <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
              </xsl:call-template>

            </v:shape>
          </w:pict>
        </w:r>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>

  <!-- 
  Summary: frames containing internal images
  Author: Clever Age
  -->
  <xsl:template match="draw:frame[not(./draw:object-ole or ./draw:object) and starts-with(./draw:image/@xlink:href, 'Pictures/')]" mode="paragraph">
    <!-- insert link to TOC field when required (user indexes) -->
    <xsl:call-template name="InsertTCField"/>

    <xsl:variable name="supported">
      <xsl:call-template name="image-support">
        <xsl:with-param name="name" select="./draw:image/@xlink:href"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$supported = 'true'">
      <w:r>
        <xsl:call-template name="InsertImage"/>
      </w:r>
    </xsl:if>

  </xsl:template>

  <!-- 
  Summary: converts text boxes in frames
  Author: Clever Age
  -->
  <xsl:template match="draw:text-box" mode="paragraph">
    <w:r>
      <w:rPr>
        <xsl:call-template name="InsertTextBoxStyle"/>
      </w:rPr>
      <w:pict>

        <!-- this properties are needed to make z-index work properly -->
        <v:shapetype coordsize="21600,21600" path="m,l,21600r21600,l21600,xe"
          xmlns:o="urn:schemas-microsoft-com:office:office">
          <v:stroke joinstyle="miter"/>
          <v:path gradientshapeok="t" o:connecttype="rect"/>
        </v:shapetype>

        <v:shape type="#_x0000_t202">
          <xsl:variable name="styleName" select="parent::draw:frame/@draw:style-name"/>
          <xsl:variable name="automaticStyle" select="key('automatic-styles', $styleName)"/>
          <xsl:variable name="officeStyle"
            select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $styleName]"/>
          <xsl:variable name="shapeStyle" select="$automaticStyle | $officeStyle"/>

          <xsl:call-template name="InsertShapeProperties">
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
            <xsl:with-param name="shapeProperties" select="parent::draw:frame"/>
          </xsl:call-template>

          <xsl:call-template name="InsertTextBox">
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
          </xsl:call-template>

          <xsl:call-template name="InsertShapeWrap">
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
          </xsl:call-template>

        </v:shape>
      </w:pict>
    </w:r>
  </xsl:template>

  <!-- 
  Summary: forward shapes in paragraph mode to shapes mode 
  Author: Clever Age
  -->
  <xsl:template match="draw:custom-shape|draw:rect|draw:ellipse" mode="paragraph">
    <!-- COMMENT : many other shapes to be handled by 1.1 -->
    <xsl:choose>
      <xsl:when test="ancestor::draw:text-box">
        <xsl:message terminate="no">translation.odf2oox.nestedFrames</xsl:message>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates select="." mode="shapes"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- 
  Summary: custom shapes
  Author: Clever Age
  -->
  <xsl:template match="draw:custom-shape" mode="shapes">
    <xsl:call-template name="InsertShapes">
      <xsl:with-param name="shapeType">
        <xsl:value-of select="draw:enhanced-geometry/@draw:type"/>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <!-- 
  Summary: rectangles and ellipses
  Author: Clever Age
  -->
  <xsl:template match="draw:rect|draw:ellipse" mode="shapes">
    <xsl:call-template name="InsertShapes">
      <xsl:with-param name="shapeType" select="name()"/>
    </xsl:call-template>
  </xsl:template>

  <!-- 
  Summary: error message for embedded plugins
  Author: Clever Age
  -->
  <xsl:template match="draw:plugin" mode="paragraph">
    <xsl:message terminate="no">translation.odf2oox.embeddedPluginFile</xsl:message>
  </xsl:template>

  <!-- 
  Summary: error message for applets
  Author: Clever Age
  -->
  <xsl:template match="draw:applet" mode="paragraph">
    <xsl:message terminate="no">translation.odf2oox.applet</xsl:message>
  </xsl:template>

  <!-- 
  Summary: error message for floating frames
  Author: Clever Age
  -->
  <xsl:template match="draw:floating-frame" mode="paragraph">
    <!-- insert link to TOC field when required (user indexes) -->
    <!--sl:call-template name="InsertTCField"/-->
    <xsl:message terminate="no">translation.odf2oox.floatingFrame</xsl:message>
  </xsl:template>

  <!--
  Summary: error message: object with xml representation are lost
  Author: Clever Age
  -->
  <xsl:template match="draw:frame[./draw:object]" mode="paragraph">
    <!-- insert link to TOC field when required (user indexes) -->
    <!--sl:call-template name="InsertTCField"/-->
    <xsl:message terminate="no">translation.odf2oox.embeddedObject</xsl:message>
  </xsl:template>

  <!-- 
  *************************************************************************
  CALLED TEMPLATES
  *************************************************************************
  -->

  <!--  word has two types of images: inline (positioned with text) and anchor (can be aligned relative to page, margin etc); -->
  <xsl:template name="InsertImage">

    <w:drawing>
      <xsl:variable name="imageId">
        <xsl:call-template name="GetPosition"/>
      </xsl:variable>

      <!-- width -->
      <xsl:variable name="cx">
        <xsl:call-template name="ComputeDrawingObjectSize">
          <xsl:with-param name="side">width</xsl:with-param>
          <xsl:with-param name="frame" select="."/>
          <xsl:with-param name="unit">emu</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>

      <!-- height-->
      <xsl:variable name="cy">
        <xsl:call-template name="ComputeDrawingObjectSize">
          <xsl:with-param name="side">height</xsl:with-param>
          <xsl:with-param name="frame" select="."/>
          <xsl:with-param name="unit">emu</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>

      <xsl:variable name="automaticStyle" select="key('automatic-styles', @draw:style-name)"/>
      <xsl:variable name="officeStyle"
        select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = @draw:style-name]"/>
      <xsl:variable name="imageStyle" select="$automaticStyle | $officeStyle"/>

      <!-- check if inline image. -->
      <xsl:variable name="wrappedPara">
        <xsl:variable name="wrapping">
          <xsl:call-template name="GetGraphicProperties">
            <xsl:with-param name="shapeStyle" select="$imageStyle"/>
            <xsl:with-param name="attribName">style:wrap</xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:if test="$wrapping = 'parallel' ">
          <xsl:call-template name="GetGraphicProperties">
            <xsl:with-param name="shapeStyle" select="$imageStyle"/>
            <xsl:with-param name="attribName">style:number-wrapped-paragraphs</xsl:with-param>
          </xsl:call-template>
        </xsl:if>
      </xsl:variable>

      <xsl:choose>
        <!-- 
        Image embedded in draw:frame/draw:text-box or in text:note element has to be inline with text.
        Word cannot not have w:anchor elements in v:shape elements (ODF frames are converted to Word shapes).
        -->
        <xsl:when test="ancestor::draw:text-box or @text:anchor-type='as-char' or ancestor::text:note[@ text:note-class='endnote'] or $wrappedPara = 1">
        <xsl:call-template name="InsertInlineImage">
            <xsl:with-param name="cx" select="$cx"/>
            <xsl:with-param name="cy" select="$cy"/>
            <xsl:with-param name="imageId" select="$imageId"/>
            <xsl:with-param name="imageStyle" select="$imageStyle"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="InsertAnchorImage">
            <xsl:with-param name="cx" select="$cx"/>
            <xsl:with-param name="cy" select="$cy"/>
            <xsl:with-param name="imageId" select="$imageId"/>
            <xsl:with-param name="imageStyle" select="$imageStyle"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </w:drawing>
  </xsl:template>

  <!-- compute the size of an image -->
  <xsl:template name="ComputeDrawingObjectSize">
    <xsl:param name="side"/>
    <xsl:param name="frame"/>
    <xsl:param name="unit"/>

    <xsl:choose>
      <xsl:when test="$frame/@*[name()=concat('svg:',$side)]">
        <xsl:call-template name="ConvertMeasure">
          <xsl:with-param name="unit" select="$unit"/>
          <xsl:with-param name="length" select="$frame/@*[name()=concat('svg:',$side)]"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$frame/*[contains(name(), 'draw:')]/@*[name()=concat('fo:min-',$side)]">
        <xsl:call-template name="ConvertMeasure">
          <xsl:with-param name="unit" select="$unit"/>
          <xsl:with-param name="length">
            <xsl:value-of
              select="$frame/*[contains(name(), 'draw:')]/@*[name()=concat('fo:min-',$side)]"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$frame/@*[name()=concat('fo:min-',$side)]">
        <xsl:call-template name="ConvertMeasure">
          <xsl:with-param name="unit" select="$unit"/>
          <xsl:with-param name="length">
            <xsl:value-of select="$frame/@*[name()=concat('fo:min-',$side)]"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>0</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--image positioned with text -->
  <xsl:template name="InsertInlineImage">

    <xsl:param name="cx"/>
    <xsl:param name="cy"/>
    <xsl:param name="imageId"/>
    <xsl:param name="imageStyle"/>

    <wp:inline>

      <!--image height and width-->
      <wp:extent cx="{$cx}" cy="{$cy}"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>

      <!-- image graphical properties: borders, fill, ratio blocking etc -->
      <xsl:call-template name="InsertImageGraphicProperties">
        <xsl:with-param name="cx" select="$cx"/>
        <xsl:with-param name="cy" select="$cy"/>
        <xsl:with-param name="imageId" select="$imageId"/>
        <xsl:with-param name="imageStyle" select="$imageStyle"/>
      </xsl:call-template>
    </wp:inline>

  </xsl:template>

  <!-- anchor images can be aligned relative to page, margin etc -->
  <xsl:template name="InsertAnchorImage">

    <xsl:param name="cx"/>
    <xsl:param name="cy"/>
    <xsl:param name="imageId"/>
    <xsl:param name="imageStyle"/>

    <wp:anchor simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">

      <!-- image z-index-->
      <xsl:call-template name="InsertZindex">
        <xsl:with-param name="imageStyle" select="$imageStyle"/>
      </xsl:call-template>

      <!-- position -->
      <xsl:call-template name="InsertAnchorImagePosition">
        <xsl:with-param name="imageStyle" select="$imageStyle"/>
      </xsl:call-template>

      <!--height and width -->
      <wp:extent cx="{$cx}" cy="{$cy}"/>

      <!--image wrapping-->
      <xsl:call-template name="InsertAnchorImageWrapping">
        <xsl:with-param name="imageStyle" select="$imageStyle"/>
      </xsl:call-template>

      <!-- image graphical properties: borders, fill, ratio blocking etc-->
      <xsl:call-template name="InsertImageGraphicProperties">
        <xsl:with-param name="cx" select="$cx"/>
        <xsl:with-param name="cy" select="$cy"/>
        <xsl:with-param name="imageId" select="$imageId"/>
        <xsl:with-param name="imageStyle" select="$imageStyle"/>
      </xsl:call-template>
    </wp:anchor>

  </xsl:template>

  <!-- inserts textboxes which are embedded in odf as one after another in word -->
  <xsl:template name="InsertEmbeddedTextboxes">
    <xsl:apply-templates select="draw:text-box" mode="paragraph"/>
  </xsl:template>

  <!-- check if image type is supported in word  -->
  <xsl:template name="image-support">
    <xsl:param name="name"/>
    <xsl:variable name="support">
      <xsl:choose>
        <xsl:when test="contains($name, '.svm')">
          <xsl:message terminate="no">translation.odf2oox.svmImage</xsl:message>
          <xsl:text>false</xsl:text>
        </xsl:when>
        <!-- WMF images inside text-box are properly displayed in Word 2007, but cause an opening crash in prior Word versions -->
        <!--xsl:when test="contains($name, '.wmf') and ancestor::draw:text-box">
          <xsl:message terminate="no">feedback:WMF image inside text-box</xsl:message> false </xsl:when-->
        <xsl:when test="contains($name, '.jpg')">true</xsl:when>
        <xsl:when test="contains($name, '.jpeg')">true</xsl:when>
        <xsl:when test="contains($name, '.gif')">true</xsl:when>
        <xsl:when test="contains($name, '.png')">true</xsl:when>
        <xsl:when test="contains($name, '.emf')">true</xsl:when>
        <xsl:when test="contains($name, '.wmf')">true</xsl:when>
        <xsl:when test="contains($name, '.tif')">true</xsl:when>
        <xsl:when test="contains($name, '.tiff')">true</xsl:when>
        <xsl:otherwise>
          <xsl:message terminate="no">translation.odf2oox.unsupportedImageType</xsl:message>
          false
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:value-of select="$support"/>
  </xsl:template>

  <!-- Get the position of an element in the draw:frame group -->
  <xsl:template name="GetPosition">
    <xsl:variable name="node" select="."/>
    <xsl:variable name="positionInGroup">
      <xsl:for-each select="document('content.xml')">
        <xsl:for-each select="key('frames', '')">
          <xsl:if test="generate-id($node) = generate-id(.)">
            <xsl:value-of select="position()"/>
          </xsl:if>
        </xsl:for-each>
      </xsl:for-each>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="string-length($positionInGroup)>0">
        <xsl:value-of select="$positionInGroup"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="countContentImages">
          <xsl:for-each select="document('content.xml')">
            <xsl:value-of select="count(key('frames', ''))"/>
          </xsl:for-each>
        </xsl:variable>
        <xsl:for-each select="document('styles.xml')">
          <xsl:for-each select="key('frames', '')">
            <xsl:if test="generate-id($node) = generate-id(.)">
              <xsl:value-of select="$countContentImages+position()"/>
            </xsl:if>
          </xsl:for-each>
        </xsl:for-each>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertZindex">
    <xsl:param name="imageStyle"/>

    <xsl:variable name="inBackground">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$imageStyle"/>
        <xsl:with-param name="attribName">style:run-through</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="wrap">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$imageStyle"/>
        <xsl:with-param name="attribName">style:wrap</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:attribute name="relativeHeight">
      <xsl:choose>
        <xsl:when test="@draw:z-index">
          <xsl:value-of select="2 + @draw:z-index"/>
        </xsl:when>
        <xsl:otherwise>2</xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>

    <xsl:attribute name="behindDoc">
      <xsl:choose>
        <xsl:when test="$wrap = 'run-through' and $inBackground = 'background' ">1</xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>

  <!-- handle the frames that are ignored. context is text:p -->
  <xsl:template name="InsertPrecedingDrawingObject">
    <xsl:if test="parent::office:text and preceding-sibling::node()[1][contains(name(), 'draw:')]">
      <xsl:choose>
        <xsl:when test="preceding-sibling::node()[1][self::draw:frame]">
          <!--xsl:apply-templates select="preceding-sibling::node()[1][contains(name(), 'draw:')]"
            mode="paragraph"/-->
          <xsl:if
            test="preceding-sibling::node()[1][not(descendant::*[(self::text:p or self::text:h) and (@text:style-name='Sender' or @text:style-name='Addressee')])]">
            <xsl:apply-templates select="preceding-sibling::node()[1][contains(name(), 'draw:')]"
              mode="paragraph"/>
          </xsl:if>
        </xsl:when>
        <xsl:otherwise>
          <xsl:apply-templates select="preceding-sibling::node()[1][contains(name(), 'draw:')]"
            mode="shapes"/>
        </xsl:otherwise>
      </xsl:choose>
      <xsl:for-each select="preceding-sibling::node()[1]">
        <xsl:call-template name="InsertPrecedingDrawingObject"/>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertImageGraphicProperties">
    <xsl:param name="imageId"/>
    <xsl:param name="cx"/>
    <xsl:param name="cy"/>
    <xsl:param name="imageStyle"/>

    <!--drawing non-visual properties-->
    <wp:docPr name="{@draw:name}" id="{$imageId}">

      <!--drawing element onclick hyperlink-->
      <xsl:if test="ancestor::draw:a">
        <a:hlinkClick xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          r:id="{generate-id(ancestor::draw:a)}">
          <xsl:if test="ancestor::draw:a/@office:target-frame-name">
            <xsl:attribute name="tgtFrame">
              <xsl:value-of select="ancestor::draw:a/@office:target-frame-name"/>
            </xsl:attribute>
          </xsl:if>
        </a:hlinkClick>
      </xsl:if>
    </wp:docPr>

    <!--drawing aspect ratio blocking-->
    <wp:cNvGraphicFramePr>
      <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        noChangeAspect="1"/>
    </wp:cNvGraphicFramePr>

    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:nvPicPr>
            <pic:cNvPr name="{@draw:name}" id="{$imageId}"/>
            <pic:cNvPicPr>
              <a:picLocks noChangeAspect="1"/>
            </pic:cNvPicPr>
          </pic:nvPicPr>
          <pic:blipFill>
            <a:blip r:embed="{generate-id(draw:image)}"/>
            <xsl:if
              test="key('automatic-styles', @draw:style-name)/style:graphic-properties/@fo:clip">
              <xsl:message terminate="no">translation.odf2oox.croppedImage</xsl:message>
              <xsl:call-template name="InsertImageCropProperties">
                <xsl:with-param name="clip"
                  select="key('automatic-styles', @draw:style-name)/style:graphic-properties/@fo:clip"
                />
              </xsl:call-template>
            </xsl:if>
            <a:stretch>
              <a:fillRect/>
            </a:stretch>
          </pic:blipFill>
          <pic:spPr>
            <a:xfrm>
              <a:off x="0" y="0"/>
              <a:ext cx="{$cx}" cy="{$cy}"/>
            </a:xfrm>
            <a:prstGeom prst="rect">
              <a:avLst/>
            </a:prstGeom>
            <xsl:call-template name="InsertImageBorders">
              <xsl:with-param name="imageStyle" select="$imageStyle"/>
            </xsl:call-template>
            <xsl:call-template name="InsertImageEffects">
              <xsl:with-param name="imageStyle" select="$imageStyle"/>
            </xsl:call-template>
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </xsl:template>

  <!--  image vertical and horizontal position-->
  <xsl:template name="InsertAnchorImagePosition">
    <xsl:param name="imageStyle"/>

    <xsl:variable name="ox">
      <xsl:call-template name="emu-measure">
        <xsl:with-param name="length">
          <xsl:call-template name="ComputeMarginX">
            <xsl:with-param name="parent"
              select="ancestor-or-self::node()[contains(name(), 'draw:')]"/>
          </xsl:call-template>
          <xsl:text>pt</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="oy">
      <xsl:call-template name="emu-measure">
        <xsl:with-param name="length">
          <xsl:call-template name="ComputeMarginY">
            <xsl:with-param name="parent"
              select="ancestor-or-self::node()[contains(name(), 'draw:')]"/>
          </xsl:call-template>
          <xsl:text>pt</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:call-template name="InsertAnchorImagePosH">
      <xsl:with-param name="imageStyle" select="$imageStyle"/>
      <xsl:with-param name="ox" select="$ox"/>
      <xsl:with-param name="oy" select="$oy"/>
    </xsl:call-template>


    <xsl:call-template name="InsertAnchorImagePosV">
      <xsl:with-param name="imageStyle" select="$imageStyle"/>
      <xsl:with-param name="ox" select="$ox"/>
      <xsl:with-param name="oy" select="$oy"/>
    </xsl:call-template>


  </xsl:template>

  <!--vertical position-->
  <xsl:template name="InsertAnchorImagePosV">
    <xsl:param name="ox"/>
    <xsl:param name="oy"/>
    <xsl:param name="imageStyle"/>

    <xsl:variable name="verticalRel">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$imageStyle"/>
        <xsl:with-param name="attribName">style:vertical-rel</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <wp:positionV>
      <!--relative position-->
      <xsl:attribute name="relativeFrom">
        <xsl:choose>
          <xsl:when test="$verticalRel='page'">page</xsl:when>
          <xsl:when test="$verticalRel='page-content' ">margin</xsl:when>
          <xsl:when test="$verticalRel='paragraph' ">line</xsl:when>
          <xsl:otherwise>paragraph</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>

      <!-- align -->
      <xsl:variable name="vertical-pos">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$imageStyle"/>
          <xsl:with-param name="attribName">style:vertical-pos</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:choose>
        <xsl:when
          test="$vertical-pos != '' and $vertical-pos !='from-top' and $vertical-pos !='below' ">
          <wp:align>
            <xsl:choose>
              <xsl:when test="$vertical-pos = 'top' ">top</xsl:when>
              <xsl:when test="$vertical-pos = 'bottom' ">bottom</xsl:when>
              <xsl:when test="$vertical-pos = 'middle' ">center</xsl:when>
            </xsl:choose>
          </wp:align>
        </xsl:when>
        <xsl:otherwise>
          <wp:posOffset>
            <xsl:choose>
              <xsl:when test="@text:anchor-type='as-char' and $oy &lt; 0">
                <xsl:call-template name="emu-measure">
                  <xsl:with-param name="length" select="'0.01cm'"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$oy"/>
              </xsl:otherwise>
            </xsl:choose>
          </wp:posOffset>
        </xsl:otherwise>
      </xsl:choose>
    </wp:positionV>
  </xsl:template>

  <!-- horizontal position-->
  <xsl:template name="InsertAnchorImagePosH">
    <xsl:param name="ox"/>
    <xsl:param name="oy"/>
    <xsl:param name="imageStyle"/>

    <xsl:variable name="horizontal-pos">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$imageStyle"/>
        <xsl:with-param name="attribName">style:horizontal-pos</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <wp:simplePos x="0" y="0"/>
    <wp:positionH>
      <!--relative position-->
      <xsl:attribute name="relativeFrom">
        <xsl:variable name="horizontalRel">
          <xsl:call-template name="GetGraphicProperties">
            <xsl:with-param name="shapeStyle" select="$imageStyle"/>
            <xsl:with-param name="attribName">style:horizontal-rel</xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="$horizontalRel='page'">page</xsl:when>
          <xsl:when test="$horizontalRel='paragraph'">column</xsl:when>
          <xsl:otherwise>margin</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>

      <!--align-->
      <xsl:choose>
        <xsl:when
          test="$horizontal-pos != '' and $horizontal-pos !='from-left' and $horizontal-pos !='from-outside' ">
          <wp:align>
            <xsl:value-of select="$horizontal-pos"/>
          </wp:align>
        </xsl:when>
        <xsl:otherwise>
          <wp:posOffset>
            <xsl:value-of select="$ox"/>
          </wp:posOffset>
        </xsl:otherwise>
      </xsl:choose>
    </wp:positionH>
  </xsl:template>

  <!--image crop-->
  <xsl:template name="InsertImageCropProperties">
    <xsl:param name="clip"/>
    <a:srcRect>
      <xsl:variable name="firstVal" select="substring-before(substring-after($clip,'('),' ' )"/>
      <xsl:variable name="secondVal"
        select="substring-before(substring-after($clip,concat('(',$firstVal,' ')),' ')"/>
      <xsl:variable name="thirdVal"
        select="substring-before(substring-after($clip,concat('(',$firstVal,' ',$secondVal,' ')),' ')"/>
      <xsl:variable name="fourthVal"
        select="substring-before(substring-after($clip,concat('(',$firstVal,' ',$secondVal,' ',$thirdVal,' ')),')')"/>
      <xsl:variable name="t">
        <xsl:call-template name="twips-measure">
          <xsl:with-param name="length" select="$firstVal"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="b">
        <xsl:call-template name="twips-measure">
          <xsl:with-param name="length" select="$thirdVal"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="r">
        <xsl:call-template name="twips-measure">
          <xsl:with-param name="length" select="$secondVal"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="l">
        <xsl:call-template name="twips-measure">
          <xsl:with-param name="length" select="$fourthVal"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="height">
        <xsl:call-template name="twips-measure">
          <xsl:with-param name="length" select="@svg:height"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="width">
        <xsl:call-template name="twips-measure">
          <xsl:with-param name="length" select="@svg:width"/>
        </xsl:call-template>
      </xsl:variable>

      <xsl:variable name="filename" select="node()/@xlink:href"></xsl:variable>
     
      <!--clam: compute cropping on postprocessor-->
      <xsl:attribute name="t">
        <xsl:value-of select="concat('COMPUTEOOXCROPPING', ',' , $firstVal, ',' , $filename,  ',t,' , round(($t * 100 div ($t + $b + $height)) * 1000))"/>
      </xsl:attribute>
      <xsl:attribute name="b">
        <xsl:value-of select="concat('COMPUTEOOXCROPPING', ',' , $thirdVal, ',' , $filename,  ',b,', round(($b * 100 div ($t + $b + $height)) * 1000))"/>
      </xsl:attribute>
      <xsl:attribute name="r">
        <xsl:value-of select="concat('COMPUTEOOXCROPPING', ',' , $secondVal, ',' , $filename,  ',r,', round(($r * 100 div ($r + $l + $width)) * 1000))"/>
      </xsl:attribute>
      <xsl:attribute name="l">
        <xsl:value-of select="concat('COMPUTEOOXCROPPING', ',' , $fourthVal, ',' , $filename,  ',l,', round(($l * 100 div ($l + $r + $width)) * 1000))"/>
      </xsl:attribute>
      <!--<xsl:attribute name="b">
        <xsl:value-of select=" round(($b * 100 div ($t + $b + $height)) * 1000)"/>
      </xsl:attribute>
      <xsl:attribute name="r">
        <xsl:value-of select=" round(($r * 100 div ($r + $l + $width)) * 1000)"/>
      </xsl:attribute>
      <xsl:attribute name="l">
        <xsl:value-of select=" round(($l * 100 div ($l + $r + $width)) * 1000)"/>
      </xsl:attribute>-->
    </a:srcRect>
  </xsl:template>

  <!--image wrap type-->
  <xsl:template name="InsertAnchorImageWrapping">
    <xsl:param name="imageStyle"/>

    <xsl:variable name="wrap">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$imageStyle"/>
        <xsl:with-param name="attribName">style:wrap</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$wrap != 'parallel' ">
      <xsl:call-template name="InsertWrapExtent">
        <xsl:with-param name="wrap" select="$wrap"/>
        <xsl:with-param name="imageStyle" select="$imageStyle"/>
      </xsl:call-template>
    </xsl:if>

    <xsl:choose>
      <xsl:when test="$wrap = 'parallel' ">
        <xsl:variable name="wrappedPara">
          <xsl:call-template name="GetGraphicProperties">
            <xsl:with-param name="shapeStyle" select="$imageStyle"/>
            <xsl:with-param name="attribName">style:number-wrapped-paragraphs</xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="$wrappedPara = 1">
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapNone/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="InsertWrapExtent">
              <xsl:with-param name="wrap" select="$wrap"/>
              <xsl:with-param name="imageStyle" select="$imageStyle"/>
            </xsl:call-template>
            <xsl:call-template name="InsertSquareWrap">
              <xsl:with-param name="wrap" select="$wrap"/>
              <xsl:with-param name="imageStyle" select="$imageStyle"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <xsl:when test="$wrap ='left' or $wrap = 'right' or $wrap ='dynamic'">
        <xsl:call-template name="InsertSquareWrap">
          <xsl:with-param name="wrap" select="$wrap"/>
          <xsl:with-param name="imageStyle" select="$imageStyle"/>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="$wrap = 'run-through'">
        <wp:wrapNone/>
      </xsl:when>

      <xsl:when test="$wrap = 'none' ">
        <xsl:call-template name="InsertTopBottomWrap">
          <xsl:with-param name="wrap" select="$wrap"/>
          <xsl:with-param name="imageStyle" select="$imageStyle"/>
        </xsl:call-template>
      </xsl:when>

      <xsl:otherwise>
        <wp:wrapTopAndBottom/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- insert wrapping extent if wrap is  -->
  <xsl:template name="InsertWrapExtent">
    <xsl:param name="wrap"/>
    <xsl:param name="imageStyle"/>

    <xsl:choose>
      <xsl:when
        test="$wrap = 'parallel' or $wrap ='left' or $wrap = 'right' or $wrap ='dynamic' or $wrap = 'none' ">
        <wp:effectExtent>
          <xsl:attribute name="l">
            <xsl:call-template name="ComputeImageBorderWidth">
              <xsl:with-param name="imageStyle" select="$imageStyle"/>
              <xsl:with-param name="side">left</xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
          <xsl:attribute name="t">
            <xsl:call-template name="ComputeImageBorderWidth">
              <xsl:with-param name="imageStyle" select="$imageStyle"/>
              <xsl:with-param name="side">top</xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
          <xsl:attribute name="r">
            <xsl:call-template name="ComputeImageBorderWidth">
              <xsl:with-param name="imageStyle" select="$imageStyle"/>
              <xsl:with-param name="side">right</xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
          <xsl:attribute name="b">
            <xsl:call-template name="ComputeImageBorderWidth">
              <xsl:with-param name="imageStyle" select="$imageStyle"/>
              <xsl:with-param name="side">bottom</xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </wp:effectExtent>
      </xsl:when>
      <xsl:otherwise>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="ComputeImageBorderWidth">
    <xsl:param name="side"/>
    <xsl:param name="imageStyle"/>

    <xsl:variable name="borderWidth">
      <xsl:call-template name="GetFrameBorder">
        <xsl:with-param name="shapeStyle" select="$imageStyle"/>
        <xsl:with-param name="side" select="$side"/>
      </xsl:call-template>
    </xsl:variable>
    <!-- if double border (should then contain ' '), compute width -->
    <xsl:call-template name="CheckBorder">
      <xsl:with-param name="unit">emu</xsl:with-param>
      <xsl:with-param name="length">
        <xsl:choose>
          <xsl:when test="contains($borderWidth, ' ')">
            <xsl:call-template name="ComputeBorderLineWidth">
              <xsl:with-param name="unit">emu</xsl:with-param>
              <xsl:with-param name="borderLineWidth" select="$borderWidth"> </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length">
                <xsl:value-of select="$borderWidth"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <!-- square wrap type for anchor image -->
  <xsl:template name="InsertSquareWrap">
    <xsl:param name="wrap"/>
    <xsl:param name="imageStyle"/>

    <wp:wrapSquare>
      <xsl:attribute name="wrapText">
        <xsl:choose>
          <xsl:when test="$wrap = 'parallel'">bothSides</xsl:when>
          <xsl:when test="$wrap = 'dynamic'">largest</xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$wrap"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>

      <!--bottom distance from text-->
      <xsl:attribute name="distB">
        <xsl:variable name="marginB">
          <xsl:call-template name="GetGraphicProperties">
            <xsl:with-param name="shapeStyle" select="$imageStyle"/>
            <xsl:with-param name="attribName">fo:margin-bottom</xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="$marginB != ''">
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="$marginB"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="0"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>

      <!--top distance from text-->
      <xsl:variable name="marginT">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$imageStyle"/>
          <xsl:with-param name="attribName">fo:margin-top</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:attribute name="distT">
        <xsl:choose>
          <xsl:when test="$marginT != ''">
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="$marginT"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="0"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>

      <!--left distance from text-->
      <xsl:variable name="marginL">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$imageStyle"/>
          <xsl:with-param name="attribName">fo:margin-left</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:attribute name="distL">
        <xsl:choose>
          <xsl:when test="$marginL != ''">
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="$marginL"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="0"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>

      <!--right distance from text-->
      <xsl:variable name="marginR">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$imageStyle"/>
          <xsl:with-param name="attribName">fo:margin-right</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:attribute name="distR">
        <xsl:choose>
          <xsl:when test="$marginR != ''">
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="$marginR"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="0"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </wp:wrapSquare>
  </xsl:template>

  <!-- top-bottom wrap type for anchor image -->
  <xsl:template name="InsertTopBottomWrap">
    <xsl:param name="wrap"/>
    <xsl:param name="imageStyle"/>

    <wp:wrapTopAndBottom>

      <!--bottom distance from text-->
      <xsl:variable name="marginB">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$imageStyle"/>
          <xsl:with-param name="attribName">fo:margin-bottom</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:attribute name="distB">
        <xsl:choose>
          <xsl:when test="$marginB != ''">
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="$marginB"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="0"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>

      <!--top distance from text-->
      <xsl:variable name="marginT">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$imageStyle"/>
          <xsl:with-param name="attribName">fo:margin-top</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:attribute name="distT">
        <xsl:choose>
          <xsl:when test="$marginT != ''">
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="$marginT"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="0"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </wp:wrapTopAndBottom>
  </xsl:template>

  <!--image border width and line style-->
  <xsl:template name="InsertImageBorders">
    <xsl:param name="imageStyle"/>

    <xsl:variable name="border">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$imageStyle"/>
        <xsl:with-param name="attribName">fo:border</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$border != '' and $border != 'none'">
      <xsl:variable name="strokeColor" select="substring-after($border,'#')"/>

      <xsl:variable name="strokeWeight">
        <xsl:call-template name="CheckBorder">
          <xsl:with-param name="length">
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="substring-before($border,' ')"/>
            </xsl:call-template>
          </xsl:with-param>
          <xsl:with-param name="unit">emu</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>

      <a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <xsl:attribute name="cmpd">
          <xsl:choose>
            <xsl:when test="substring-before(substring-after($border,' ' ),' ' ) != 'solid' ">

              <xsl:variable name="borderLineWidth">
                <xsl:call-template name="GetGraphicProperties">
                  <xsl:with-param name="shapeStyle" select="$imageStyle"/>
                  <xsl:with-param name="attribName">style:border-line-width</xsl:with-param>
                </xsl:call-template>
              </xsl:variable>

              <xsl:if test="$borderLineWidth != ''">

                <xsl:variable name="innerLineWidth">
                  <xsl:call-template name="point-measure">
                    <xsl:with-param name="length" select="substring-before($borderLineWidth,' ' )"/>
                  </xsl:call-template>
                </xsl:variable>

                <xsl:variable name="outerLineWidth">
                  <xsl:call-template name="point-measure">
                    <xsl:with-param name="length"
                      select="substring-after(substring-after($borderLineWidth,' ' ),' ' )"/>
                  </xsl:call-template>
                </xsl:variable>

                <xsl:choose>
                  <xsl:when test="$innerLineWidth &gt; $outerLineWidth">thinThick</xsl:when>
                  <xsl:when test="$innerLineWidth &lt; $outerLineWidth">thickThin</xsl:when>
                  <xsl:otherwise>dbl</xsl:otherwise>
                </xsl:choose>

              </xsl:if>
            </xsl:when>
            <xsl:otherwise>sng</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>

        <xsl:attribute name="w">
          <xsl:value-of select="$strokeWeight"/>
        </xsl:attribute>
        <a:solidFill>
          <a:srgbClr>
            <xsl:attribute name="val">
              <xsl:value-of select="$strokeColor"/>
            </xsl:attribute>
          </a:srgbClr>
        </a:solidFill>
        <a:prstDash val="solid"/>
      </a:ln>
    </xsl:if>
  </xsl:template>

  <!-- image effects -->
  <xsl:template name="InsertImageEffects">
    <xsl:param name="imageStyle"/>

    <!-- shadow -->
    <xsl:variable name="shadow">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$imageStyle"/>
        <xsl:with-param name="attribName">style:shadow</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:if test="$shadow != ''  and $shadow != 'none' ">
      <a:effectLst>
        <a:outerShdw>
          <xsl:attribute name="rotWithShape">0</xsl:attribute>
          <xsl:variable name="firstShadow">
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length">
                <xsl:value-of select=" substring-before(substring-after($shadow, ' '), ' ')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="secondShadow">
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length">
                <xsl:value-of select=" substring-after(substring-after($shadow, ' '), ' ')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <!-- Handles only simple cases.
            distance calculation is wrong anyway (should be square average, but no sqroot funcion in xsl)
            More complex issues require post-processing (angle calculation) -->
          <xsl:choose>
            <xsl:when test="contains($firstShadow, '-') and contains($secondShadow, '-')">
              <xsl:attribute name="algn">tl</xsl:attribute>
              <xsl:attribute name="dir">
                <xsl:value-of select="225*60000"/>
              </xsl:attribute>
              <xsl:attribute name="dist">
                <xsl:value-of
                  select="(number(substring-after($firstShadow, '-'))+number(substring-after($secondShadow, '-'))) div 2"
                />
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="contains($firstShadow, '-') and number($secondShadow)">
              <xsl:attribute name="algn">bl</xsl:attribute>
              <xsl:attribute name="dir">
                <xsl:value-of select="135*60000"/>
              </xsl:attribute>
              <xsl:attribute name="dist">
                <xsl:value-of
                  select="(number(substring-after($firstShadow, '-'))+number($secondShadow)) div 2"
                />
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="contains($secondShadow, '-') and number($firstShadow)">
              <xsl:attribute name="algn">tr</xsl:attribute>
              <xsl:attribute name="dir">
                <xsl:value-of select="315*60000"/>
              </xsl:attribute>
              <xsl:attribute name="dist">
                <xsl:value-of
                  select="(number(substring-after($secondShadow, '-'))+number($firstShadow)) div 2"
                />
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="number($firstShadow) and number($secondShadow)">
              <xsl:attribute name="algn">br</xsl:attribute>
              <xsl:attribute name="dir">
                <xsl:value-of select="45*60000"/>
              </xsl:attribute>
              <xsl:attribute name="dist">
                <xsl:value-of select="(number($firstShadow)+number($secondShadow)) div 2"/>
              </xsl:attribute>
            </xsl:when>
          </xsl:choose>
          <!-- color : converter uses the most direct way of converting a color -->
          <a:srgbClr val="{substring-after(substring-before($shadow, ' '), '#')}"/>
        </a:outerShdw>
      </a:effectLst>
    </xsl:if>
  </xsl:template>

  <!-- shapes template -->
  <xsl:template name="InsertShapes">
    <xsl:param name="shapeType"/>
    <xsl:variable name="styleName" select=" @draw:style-name"/>
    <xsl:variable name="automaticStyle" select="key('automatic-styles', $styleName)"/>
    <xsl:variable name="officeStyle"
      select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $styleName]"/>
    <xsl:variable name="shapeStyle" select="$automaticStyle | $officeStyle"/>

    <w:r>
      <w:pict>
        <xsl:choose>

          <!-- rectangle -->
          <xsl:when test="$shapeType = 'draw:rect' or $shapeType = 'rectangle' ">
            <v:rect>
              <xsl:call-template name="SimpleShape">
                <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
                <xsl:with-param name="shapeProperties" select="."/>
              </xsl:call-template>
              <!--insert text-box-->
              <xsl:call-template name="InsertTextBox">
                <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
              </xsl:call-template>
            </v:rect>
          </xsl:when>

          <!-- ellipse -->
          <xsl:when test="$shapeType = 'draw:ellipse' or $shapeType = 'ellipse' ">
            <v:oval>
              <xsl:call-template name="SimpleShape">
                <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
                <xsl:with-param name="shapeProperties" select="."/>
              </xsl:call-template>
              <!--insert text-box-->
              <xsl:call-template name="InsertTextBox">
                <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
              </xsl:call-template>
            </v:oval>
          </xsl:when>

          <!-- round-rectangle -->

          <xsl:when test="$shapeType = 'round-rectangle' ">
            <v:roundrect>
              <xsl:call-template name="SimpleShape">
                <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
                <xsl:with-param name="shapeProperties" select="."/>
              </xsl:call-template>
              <!--insert text-box-->
              <xsl:call-template name="InsertTextBox">
                <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
              </xsl:call-template>
            </v:roundrect>
          </xsl:when>

          <!-- TODO - other shapes -->
          <xsl:otherwise/>
        </xsl:choose>
      </w:pict>
    </w:r>
  </xsl:template>

  <!-- TODO
  <xsl:template name="GenerateShapeId">
    <xsl:param name="shape"></xsl:param>
  </xsl:template>
  -->
  <!-- simple shape properties -->
  <xsl:template name="SimpleShape">
    <xsl:param name="shapeStyle"/>
    <xsl:param name="shapeProperties"/>

    <xsl:if test="$shapeStyle != 0 or count($shapeStyle) &gt; 1">
      <!-- shape properties: size, z-index, color, position, stroke, etc -->
      <xsl:call-template name="InsertDrawnShapeProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="shapeProperties" select="$shapeProperties"/>
      </xsl:call-template>

      <xsl:call-template name="FillColor">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

      <xsl:call-template name="Stroke">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

      <xsl:call-template name="InsertShapeFillProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

  <!-- shape fill color -->
  <xsl:template name="FillColor">
    <xsl:param name="shapeStyle"/>
    <xsl:variable name="fillColor">
      <xsl:call-template name="GetDrawnGraphicProperties">
        <xsl:with-param name="attrib">draw:fill-color</xsl:with-param>
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:if test="$fillColor != '' ">
      <xsl:attribute name="fillcolor">
        <xsl:value-of select="$fillColor"/>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

  <!-- shape stroke color and stroke width -->
  <xsl:template name="Stroke">
    <xsl:param name="shapeStyle"/>
    <xsl:variable name="strokeColor">
      <xsl:call-template name="GetDrawnGraphicProperties">
        <xsl:with-param name="attrib">svg:stroke-color</xsl:with-param>
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="strokeWeight">
      <xsl:call-template name="GetDrawnGraphicProperties">
        <xsl:with-param name="attrib">svg:stroke-width</xsl:with-param>
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>
    </xsl:variable>
    <!-- insert attributes -->
    <xsl:if test="$strokeColor != '' ">
      <xsl:attribute name="strokecolor">
        <xsl:value-of select="$strokeColor"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="$strokeWeight != '' ">
      <xsl:attribute name="strokeweight">
        <xsl:call-template name="point-measure">
          <xsl:with-param name="length" select="$strokeWeight"/>
        </xsl:call-template>
        <xsl:text>pt</xsl:text>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

  <!-- get graphic properties for shapes -->
  <xsl:template name="GetDrawnGraphicProperties">
    <xsl:param name="attrib"/>
    <xsl:param name="shapeStyle"/>
    <xsl:choose>
      <xsl:when test="attribute::node()[name()=$attrib]">
        <xsl:value-of select="attribute::node()[name()=$attrib]"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="attribName" select="$attrib"/>
          <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- shape properties -->
  <xsl:template name="InsertDrawnShapeProperties">
    <xsl:param name="shapeStyle"/>
    <xsl:param name="shapeProperties"/>

    <xsl:attribute name="style">

      <xsl:call-template name="InsertDrawnShapeSize"/>

      <xsl:call-template name="InsertDrawnShapeZindex"/>

      <xsl:call-template name="InsertShapePositionRelative">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="shapeProperties" select="$shapeProperties"/>
      </xsl:call-template>

      <xsl:call-template name="InsertShapePosition">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="shapeProperties" select="$shapeProperties"/>
      </xsl:call-template>

      <xsl:call-template name="InsertTextAnchor">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>

  <!--insert shape z-index -->
  <xsl:template name="InsertDrawnShapeZindex">
    <xsl:choose>
      <xsl:when test="@draw:z-index=0">
        <xsl:value-of select="concat('z-index:',2516572155-@draw:z-index,';')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="concat('z-index:',251659264+@draw:z-index,';')"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--  shape width and height-->
  <xsl:template name="InsertDrawnShapeSize">
    <xsl:variable name="frameW">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length" select="@svg:width|@fo:min-width"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="frameH">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length" select="@fo:min-height|@svg:height"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select="concat('width:',$frameW,'pt;')"/>
    <xsl:value-of select="concat('height:',$frameH,'pt;')"/>
  </xsl:template>

  <!-- computes text-box margin. Returns a value in points -->
  <xsl:template name="ComputeMarginX">
    <xsl:param name="parent"/>
    <xsl:choose>
      <xsl:when test="$parent">
        <xsl:variable name="recursive_result">
          <xsl:call-template name="ComputeMarginX">
            <xsl:with-param name="parent" select="$parent[position()>1]"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="shapeStyle" select="key('styles', $parent[1]/@draw:style-name)"/>
        <xsl:choose>
          <xsl:when test="count($shapeStyle) = 0">
            <xsl:value-of select="$recursive_result"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:variable name="anchor">
              <xsl:value-of select="$parent[1]/@text:anchor-type"/>
            </xsl:variable>
            <xsl:variable name="horizontalPos">
              <xsl:value-of select="$shapeStyle/style:graphic-properties/@style:horizontal-pos"/>
            </xsl:variable>
            <xsl:variable name="horizontalRel">
              <xsl:value-of select="$shapeStyle/style:graphic-properties/@style:horizontal-rel"/>
            </xsl:variable>
            <!-- page properties. not valid if more than one page-master-style -->
            <xsl:variable name="pageWidth">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:for-each select="document('styles.xml')">
                    <xsl:value-of
                      select="key('page-layouts', $default-master-style/@style:page-layout-name)[1]/style:page-layout-properties/@fo:page-width"
                    />
                  </xsl:for-each>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="pageLeftMargin">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:for-each select="document('styles.xml')">
                    <xsl:value-of
                      select="key('page-layouts', $default-master-style/@style:page-layout-name)[1]/style:page-layout-properties/@fo:margin-left"
                    />
                  </xsl:for-each>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="pageRightMargin">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:for-each select="document('styles.xml')">
                    <xsl:value-of
                      select="key('page-layouts', $default-master-style/@style:page-layout-name)[1]/style:page-layout-properties/@fo:margin-right"
                    />
                  </xsl:for-each>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <!-- width to be added depending on were object is located : column or page -->
            <xsl:variable name="contextWidth">
              <xsl:call-template name="ComputeContextWidth">
                <xsl:with-param name="frame" select="$parent[1]"/>
                <xsl:with-param name="horizontalRel" select="$horizontalRel"/>
                <xsl:with-param name="horizontalPos" select="$horizontalPos"/>
                <xsl:with-param name="pageWidth" select="$pageWidth"/>
                <xsl:with-param name="pageLeftMargin" select="$pageRightMargin"/>
                <xsl:with-param name="pageRightMargin" select="$pageLeftMargin"/>
              </xsl:call-template>
            </xsl:variable>
            <!-- value to be substracted from the context width -->
            <xsl:variable name="contextSubstractedValue">
              <xsl:variable name="wrap">
                <xsl:call-template name="GetGraphicProperties">
                  <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
                  <xsl:with-param name="attribName">style:wrap</xsl:with-param>
                </xsl:call-template>
              </xsl:variable>
              <xsl:choose>
                <xsl:when test="$wrap = 'none' or $wrap = 'run-through' or $wrap = '' ">
                  <xsl:variable name="precedingFrames">
                    <xsl:call-template name="ComputeContextSubstractedWidth">
                      <xsl:with-param name="frames"
                        select="$parent[1]/preceding-sibling::*[contains(name(),'draw:')]"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:variable name="followingFrames">
                    <xsl:call-template name="ComputeContextSubstractedWidth">
                      <xsl:with-param name="frames"
                        select="$parent[1]/following-sibling::*[contains(name(),'draw:')]"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:value-of select="$precedingFrames + $followingFrames"/>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <!-- values of frame : padding (margin width), width, heigth -->
            <xsl:variable name="frameMarginLeft">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:call-template name="GetGraphicProperties">
                    <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
                    <xsl:with-param name="attribName">fo:margin-left</xsl:with-param>
                  </xsl:call-template>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="frameMarginRight">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:call-template name="GetGraphicProperties">
                    <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
                    <xsl:with-param name="attribName">fo:margin-right</xsl:with-param>
                  </xsl:call-template>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="frameWidth">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length" select="$parent[1]/@svg:width"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="frameHeight">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length" select="$parent[1]/@svg:height"/>
              </xsl:call-template>
            </xsl:variable>
            <!-- if a value is specified for a frame move -->
            <xsl:variable name="fromLeft">
              <xsl:choose>
                <xsl:when test="$horizontalPos = 'from-left' or $horizontalPos = 'from-inside' ">
                  <xsl:call-template name="point-measure">
                    <xsl:with-param name="length" select="$parent[1]/@svg:x"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <!-- particular transformation -->
            <xsl:variable name="translation">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:value-of
                    select="substring-before(substring-after(substring-after($parent[1]/@draw:transform,'translate'),'('),' ')"
                  />
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <!-- find the value considering all cases -->
            <xsl:variable name="svgx">
              <xsl:choose>
                <xsl:when test="$horizontalPos = 'from-left' or $horizontalPos='from-inside' ">
                  <xsl:choose>
                    <!-- page, page-content, page-start-margin -->
                    <xsl:when
                      test="$horizontalRel = 'page' or $horizontalRel = 'page-content' or $horizontalRel = 'page-start-margin' ">
                      <xsl:value-of select="$fromLeft + $translation"/>
                    </xsl:when>
                    <!-- page-end-margin -->
                    <xsl:when test="$horizontalRel = 'page-end-margin' ">
                      <xsl:value-of select="$contextWidth + $fromLeft + $translation"/>
                    </xsl:when>
                    <!-- avoid conflics -->
                    <xsl:when test="$anchor = 'page' ">
                      <xsl:value-of select="$fromLeft + $translation"/>
                    </xsl:when>
                    <!-- paragraph, paragraph-content, paragraph-start-margin -->
                    <xsl:when
                      test="$horizontalRel = 'paragraph' or $horizontalRel = 'paragraph-content'  or $horizontalRel = 'paragraph-start-margin' ">
                      <!-- TODO : get indent property of current paragraph -->
                      <xsl:variable name="paragraphLeftIndent">0</xsl:variable>
                      <xsl:value-of
                        select="$paragraphLeftIndent - $contextSubstractedValue + $fromLeft + $translation"
                      />
                    </xsl:when>
                    <!-- paragraph-end-margin -->
                    <xsl:when test="$horizontalRel = 'paragraph-end-margin' ">
                      <!-- TODO : get indent property of current paragraph -->
                      <xsl:variable name="paragraphRightIndent">0</xsl:variable>
                      <xsl:value-of
                        select="$contextWidth - $paragraphRightIndent - $contextSubstractedValue + $fromLeft + $translation"
                      />
                    </xsl:when>
                    <!-- frame, frame-content, frame-start-margin, frame-end-margin -->
                    <xsl:when test="contains($horizontalRel, 'frame')">
                      <xsl:value-of select="$fromLeft + $translation"/>
                    </xsl:when>
                    <!-- char -->
                    <xsl:when test="$horizontalRel = 'char' ">
                      <xsl:value-of select="$fromLeft + $translation"/>
                    </xsl:when>
                    <xsl:otherwise>0</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:when
                  test="($horizontalPos='left' or $horizontalPos='inside') and $frameMarginLeft != '' ">
                  <xsl:choose>
                    <!-- page, page-content, page-start-margin -->
                    <xsl:when
                      test="$horizontalRel = 'page' or $horizontalRel = 'page-content' or $horizontalRel = 'page-start-margin' ">
                      <xsl:value-of select="$frameMarginLeft + $translation"/>
                    </xsl:when>
                    <!-- page-end-margin -->
                    <xsl:when test="$horizontalRel = 'page-end-margin' ">
                      <xsl:choose>
                        <xsl:when test="$frameWidth &lt; $pageRightMargin">
                          <xsl:value-of
                            select="$contextWidth - $contextSubstractedValue + $frameMarginLeft + $translation"
                          />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of
                            select="$pageWidth -$frameWidth - $frameMarginLeft + $translation"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>
                    <!-- avoid conflics -->
                    <xsl:when test="$anchor = 'page' ">
                      <xsl:value-of select="$frameMarginLeft + $translation"/>
                    </xsl:when>
                    <!-- paragraph, paragraph-content, paragraph-start-margin -->
                    <xsl:when
                      test="$horizontalRel = 'paragraph' or $horizontalRel = 'paragraph-content' or $horizontalRel = 'paragraph-start-margin' ">
                      <!-- TODO : get indent property of current paragraph -->
                      <xsl:variable name="paragraphLeftIndent">0</xsl:variable>
                      <xsl:value-of
                        select="$paragraphLeftIndent + $frameMarginLeft - $contextSubstractedValue + $translation"
                      />
                    </xsl:when>
                    <!-- paragraph-end-margin -->
                    <xsl:when test="$horizontalRel = 'paragraph-end-margin' ">
                      <!-- TODO : get indent property of current paragraph -->
                      <xsl:variable name="paragraphRightIndent">0</xsl:variable>
                      <xsl:value-of
                        select="$contextWidth - $paragraphRightIndent + $frameMarginLeft - $contextSubstractedValue + $translation"
                      />
                    </xsl:when>
                    <!-- frame, frame-content, frame-start-margin, frame-end-margin -->
                    <xsl:when test="contains($horizontalRel, 'frame')">
                      <xsl:value-of select="$frameMarginLeft + $translation"/>
                    </xsl:when>
                    <!-- char -->
                    <xsl:when test="$horizontalRel = 'char' ">
                      <xsl:value-of select="$frameMarginLeft + $translation"/>
                    </xsl:when>
                    <xsl:otherwise>0</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:when
                  test="($horizontalPos='right' or $horizontalPos='outside') and $frameMarginRight != '' ">
                  <xsl:choose>
                    <!-- page, page-end-margin -->
                    <xsl:when test="$horizontalRel = 'page' or $horizontalRel = 'page-end-margin' ">
                      <xsl:value-of
                        select="$pageWidth - $frameWidth - $frameMarginRight + $translation"/>
                    </xsl:when>
                    <!-- page-start-margin -->
                    <xsl:when test="$horizontalRel = 'page-start-margin' ">
                      <xsl:choose>
                        <xsl:when test="$frameWidth &lt; $pageLeftMargin">
                          <xsl:value-of
                            select="$pageLeftMargin  - $frameWidth - $frameMarginRight + $translation"
                          />
                        </xsl:when>
                        <xsl:otherwise>0</xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>
                    <!-- avoid conflics -->
                    <xsl:when test="$anchor = 'page' ">
                      <xsl:value-of
                        select="$pageWidth - $frameWidth - $frameMarginRight + $translation"/>
                    </xsl:when>
                    <!-- page-content -->
                    <xsl:when test="$horizontalRel = 'page-content' ">
                      <xsl:value-of
                        select="$contextWidth - $frameWidth - $frameMarginRight - $contextSubstractedValue + $translation"
                      />
                    </xsl:when>
                    <!-- paragraph, paragraph-content, paragraph-end-margin -->
                    <xsl:when
                      test="$horizontalRel = 'paragraph' or $horizontalRel = 'paragraph-content' or $horizontalRel = 'paragraph-end-margin' ">
                      <!-- TODO : get indent property of current paragraph -->
                      <xsl:variable name="paragraphRightIndent">0</xsl:variable>
                      <xsl:value-of
                        select="$contextWidth - $paragraphRightIndent -$frameWidth - $frameMarginRight - $contextSubstractedValue + $translation"
                      />
                    </xsl:when>
                    <!-- paragraph-start-margin -->
                    <xsl:when test="$horizontalRel = 'paragraph-start-margin' ">
                      <!-- TODO : get indent property of current paragraph -->
                      <xsl:variable name="paragraphLeftIndent">0</xsl:variable>
                      <xsl:value-of
                        select="$paragraphLeftIndent - $frameMarginRight - $contextSubstractedValue + $translation"
                      />
                    </xsl:when>
                    <!-- frame, frame-content, frame-start-margin, frame-end-margin -->
                    <xsl:when test="contains($horizontalRel, 'frame')">
                      <xsl:value-of select="$pageWidth - $frameMarginRight + $translation"/>
                    </xsl:when>
                    <!-- char -->
                    <xsl:when test="$horizontalRel = 'char' ">
                      <xsl:value-of select="$pageWidth - $frameMarginRight + $translation"/>
                    </xsl:when>
                    <xsl:otherwise>0</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:when test="$horizontalPos='center' ">
                  <xsl:choose>
                    <!-- page-start-margin -->
                    <xsl:when test="$horizontalRel = 'page-start-margin' ">
                      <xsl:choose>
                        <xsl:when test="$frameWidth &lt; $pageLeftMargin">
                          <xsl:value-of
                            select="round(number(($pageLeftMargin - $frameWidth) div 2 + $translation))"
                          />
                        </xsl:when>
                        <xsl:otherwise>0</xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>
                    <!-- page-end-margin -->
                    <xsl:when test="$horizontalRel = 'page-end-margin' ">
                      <xsl:choose>
                        <xsl:when test="$frameWidth &lt; $pageRightMargin">
                          <xsl:value-of
                            select="round(number($pageWidth - ($pageRightMargin - $frameWidth) div 2 + $translation))"
                          />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$pageWidth - $frameWidth + $translation"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>
                    <!-- paragraph-start-margin -->
                    <xsl:when test="$horizontalRel = 'paragraph-start-margin' ">
                      <!-- TODO : get indent property of current paragraph -->
                      <xsl:variable name="paragraphLeftIndent">0</xsl:variable>
                      <xsl:choose>
                        <xsl:when test="$frameWidth &lt; $paragraphLeftIndent">
                          <xsl:value-of
                            select="round(number(($paragraphLeftIndent - $frameWidth) div 2 + $translation))"
                          />
                        </xsl:when>
                        <xsl:otherwise>0</xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>
                    <!-- paragraph-end-margin -->
                    <xsl:when test="$horizontalRel = 'paragraph-end-margin' ">
                      <!-- TODO : get indent property of current paragraph -->
                      <xsl:variable name="paragraphRightIndent">0</xsl:variable>
                      <xsl:choose>
                        <xsl:when test="$frameWidth &lt; $paragraphRightIndent">
                          <xsl:value-of
                            select="round(number($contextWidth - ($paragraphRightIndent - $frameWidth) div 2 + $translation))"
                          />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$contextWidth - $frameWidth + $translation"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>
                    <!-- frame, frame-content, frame-start-margin, frame-end-margin -->
                    <xsl:when test="contains($horizontalRel, 'frame')">
                      <xsl:value-of
                        select="round(number($pageWidth - $frameWidth div 2 + $translation))"/>
                    </xsl:when>
                    <xsl:otherwise>0</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:choose>
                    <xsl:when test="$parent[1]/@svg:x">
                      <xsl:call-template name="point-measure">
                        <xsl:with-param name="length" select="$parent[1]/@svg:x"/>
                      </xsl:call-template>
                    </xsl:when>
                    <xsl:otherwise>0</xsl:otherwise>
                  </xsl:choose>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:value-of select="$svgx+$recursive_result"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>0</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- compute the width of page or paragraph margin. -->
  <!-- COMMENT : This is not correct if document has more than one page-master-style -->
  <xsl:template name="ComputeContextWidth">
    <xsl:param name="frame"/>
    <xsl:param name="horizontalRel"/>
    <xsl:param name="horizontalPos"/>
    <xsl:param name="pageWidth"/>
    <xsl:param name="pageLeftMargin"/>
    <xsl:param name="pageRightMargin"/>

    <xsl:for-each select="document('styles.xml')">
      <xsl:choose>
        <xsl:when test="contains($horizontalRel, 'page')">
          <!-- scale goes from one end of page to the other end -->
          <xsl:choose>
            <xsl:when
              test="$horizontalRel = 'page' or $horizontalRel = 'page-content' or $horizontalRel = 'page-start-margin' ">
              <xsl:value-of select="$pageWidth - $pageRightMargin"/>
            </xsl:when>
            <xsl:when test="$horizontalPos='left' or $horizontalPos='inside' ">
              <xsl:value-of select="$pageWidth - $pageRightMargin"/>
            </xsl:when>
            <xsl:when test="$horizontalPos='right' or $horizontalPos='outside' ">
              <xsl:value-of select="$pageWidth - $pageLeftMargin - $pageRightMargin"/>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="contains($horizontalRel, 'paragraph')">
          <!-- scale goes from one end of column to the other end -->
          <xsl:variable name="columnNumber">
            <xsl:choose>
              <xsl:when
                test="key('page-layouts', $default-master-style/@style:page-layout-name)[1]/style:page-layout-properties/style:columns[@fo:column-count > 0]">
                <xsl:value-of
                  select="key('page-layouts', $default-master-style/@style:page-layout-name)[1]/style:page-layout-properties/style:columns/@fo:column-count"
                />
              </xsl:when>
              <xsl:otherwise>1</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="columnGap">
            <xsl:choose>
              <xsl:when
                test="key('page-layouts', $default-master-style/@style:page-layout-name)[1]/style:page-layout-properties/style:columns/@fo:column-gap">
                <xsl:call-template name="point-measure">
                  <xsl:with-param name="length"
                    select="key('page-layouts', $default-master-style/@style:page-layout-name)[1]/style:page-layout-properties/style:columns/@fo:column-gap"
                  />
                </xsl:call-template>
              </xsl:when>
              <xsl:otherwise>0</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <!-- we assume that all columns are the same width (cannot know in which column we are) -->
          <xsl:value-of
            select="round(($pageWidth - $pageLeftMargin - $pageRightMargin - $columnGap) div $columnNumber)"
          />
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <!-- compute a value to be substracted to context width -->
  <xsl:template name="ComputeContextSubstractedWidth">
    <xsl:param name="frames"/>
    <xsl:param name="side"/>
    <xsl:choose>
      <xsl:when test="$frames">
        <xsl:variable name="otherValues">
          <xsl:call-template name="ComputeContextSubstractedWidth">
            <xsl:with-param name="frames" select="$frames[position() &gt; 1]"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="currentVal">
          <xsl:choose>
            <xsl:when
              test="key('automatic-styles', $frames[1]/@draw:style-name)/style:graphic-properties/@style:wrap = 'right' ">
              <xsl:call-template name="ComputeDrawingObjectSize">
                <xsl:with-param name="side">width</xsl:with-param>
                <xsl:with-param name="frame" select="$frames[1]"/>
                <xsl:with-param name="unit">point</xsl:with-param>
              </xsl:call-template>
            </xsl:when>
            <xsl:when
              test="key('automatic-styles', $frames[1]/@draw:style-name)/style:graphic-properties/@style:wrap = 'left' ">
              <xsl:call-template name="ComputeDrawingObjectSize">
                <xsl:with-param name="side">width</xsl:with-param>
                <xsl:with-param name="frame" select="$frames[1]"/>
                <xsl:with-param name="unit">point</xsl:with-param>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:value-of select="$currentVal + $otherValues"/>
      </xsl:when>
      <xsl:otherwise>0</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- compute margin when rotation -->
  <xsl:template name="ComputeMarginXWithRotation">
    <xsl:param name="angle"/>

    <!-- particular transformation -->
    <xsl:variable name="translationX">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length">
          <xsl:value-of
            select="substring-before(substring-after(substring-after(ancestor-or-self::node()[contains(name(), 'draw:') and @draw:transform][1]/@draw:transform,'translate'),'('),' ')"
          />
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <!-- frame properties -->
    <xsl:variable name="frameWidth">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length"
          select="ancestor-or-self::node()[contains(name(), 'draw:') and @draw:transform][1]/@svg:width"
        />
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="frameHeight">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length"
          select="ancestor-or-self::node()[contains(name(), 'draw:') and @draw:transform][1]/@svg:height"
        />
      </xsl:call-template>
    </xsl:variable>
    <!-- value used by Word to calculate frame position -->
    <xsl:variable name="frameOffset">
      <xsl:value-of select="round(($frameWidth - $frameHeight) div 2)"/>
    </xsl:variable>
    <!-- special distance with rotation -->
    <xsl:choose>
      <xsl:when test="$angle = 90">
        <xsl:value-of select="$translationX - $frameHeight - $frameOffset"/>
      </xsl:when>
      <xsl:when test="$angle = -90">
        <xsl:value-of select="$translationX - $frameOffset"/>
      </xsl:when>
      <xsl:when test="$angle &gt; 0">
        <xsl:value-of select="$translationX - $frameHeight - $frameOffset"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$translationX - $frameOffset"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- computes text-box margin. Returns a measure in Points -->
  <xsl:template name="ComputeMarginY">
    <xsl:param name="parent"/>
    <xsl:choose>
      <xsl:when test="$parent">
        <xsl:variable name="recursive_result">
          <xsl:call-template name="ComputeMarginY">
            <xsl:with-param name="parent" select="$parent[position()>1]"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="shapeStyle" select="key('styles', $parent[1]/@draw:style-name)"/>
        <xsl:choose>
          <xsl:when test="count($shapeStyle) = 0">
            <xsl:value-of select="$recursive_result"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:variable name="verticalPos">
              <xsl:value-of select="$shapeStyle/style:graphic-properties/@style:vertical-pos"/>
            </xsl:variable>
            <xsl:variable name="verticalRel">
              <xsl:value-of select="$shapeStyle/style:graphic-properties/@style:vertical-rel"/>
            </xsl:variable>
            <xsl:variable name="pageHeight">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:value-of
                    select="document('styles.xml')/office:document-styles/office:automatic-styles/style:page-layout/style:page-layout-properties/@fo:page-height"
                  />
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="pageTopMargin">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:value-of
                    select="document('styles.xml')/office:document-styles/office:automatic-styles/style:page-layout/style:page-layout-properties/@fo:margin-top"
                  />
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="pageBottomMargin">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:value-of
                    select="document('styles.xml')/office:document-styles/office:automatic-styles/style:page-layout/style:page-layout-properties/@fo:margin-bottom"
                  />
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="frameMarginTop">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:call-template name="GetGraphicProperties">
                    <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
                    <xsl:with-param name="attribName">fo:margin-top</xsl:with-param>
                  </xsl:call-template>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="frameMarginBottom">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:call-template name="GetGraphicProperties">
                    <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
                    <xsl:with-param name="attribName">fo:margin-bottom</xsl:with-param>
                  </xsl:call-template>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="frameWidth">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length" select="$parent[1]/@svg:width"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="frameHeight">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length" select="$parent[1]/@svg:height"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="fromTop">
              <xsl:choose>
                <xsl:when test="$verticalPos = 'from-top' ">
                  <xsl:call-template name="point-measure">
                    <xsl:with-param name="length" select="$parent[1]/@svg:y"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:variable name="translation">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:value-of
                    select="substring-before(substring-after(substring-after(substring-after($parent[1]/@draw:transform,'translate'),'('),' '),')')"
                  />
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="svgy">
              <xsl:choose>
                <xsl:when test="$verticalPos='from-top' ">
                  <xsl:choose>
                    <!-- page -->
                    <xsl:when test="$verticalRel = 'page' ">
                      <xsl:value-of select="$fromTop + $translation"/>
                    </xsl:when>
                    <!-- page-content, paragraph -->
                    <xsl:when test="$verticalRel = 'page-content' or $verticalRel = 'paragraph' ">
                      <xsl:value-of select="$fromTop + $translation"/>
                    </xsl:when>
                    <!-- paragraph-content -->
                    <xsl:when test="$verticalRel = 'paragraph-content' ">
                      <!-- TODO : get spacing property of current paragraph -->
                      <xsl:variable name="paragraphTopSpacing">0</xsl:variable>
                      <xsl:value-of select="$paragraphTopSpacing + $fromTop + $translation"/>
                    </xsl:when>
                    <!-- frame, frame-content -->
                    <xsl:when test="contains($verticalRel, 'frame')">
                      <xsl:value-of select="$fromTop + $translation"/>
                    </xsl:when>
                    <xsl:otherwise>0</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:when test="$verticalPos='top' and $frameMarginTop != '' ">
                  <xsl:choose>
                    <!-- page -->
                    <xsl:when test="$verticalRel = 'page' ">
                      <xsl:value-of select="$frameMarginTop + $translation"/>
                    </xsl:when>
                    <!-- page-content, paragraph -->
                    <xsl:when test="$verticalRel = 'page-content' or $verticalRel = 'paragraph' ">
                      <xsl:value-of select="$frameMarginTop + $translation"/>
                    </xsl:when>
                    <!-- paragraph-content -->
                    <xsl:when test="$verticalRel = 'paragraph-content' ">
                      <!-- TODO : get spacing property of current paragraph -->
                      <xsl:variable name="paragraphTopSpacing">0</xsl:variable>
                      <xsl:value-of select="$paragraphTopSpacing + $frameMarginTop + $translation"/>
                    </xsl:when>
                    <!-- frame, frame-content -->
                    <xsl:when test="contains($verticalRel, 'frame')">
                      <xsl:value-of select="$frameMarginTop + $translation"/>
                    </xsl:when>
                    <xsl:otherwise>0</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:when test="$verticalPos='bottom' and $frameMarginBottom != '' ">
                  <xsl:choose>
                    <!-- page -->
                    <xsl:when test="$verticalRel = 'page' ">
                      <xsl:value-of
                        select="$pageHeight - $frameHeight - $frameMarginBottom + $translation"/>
                    </xsl:when>
                    <!-- page-content -->
                    <xsl:when test="$verticalRel = 'page-content' ">
                      <xsl:value-of
                        select="$pageHeight - $pageTopMargin - $pageBottomMargin - $frameHeight - $frameMarginBottom + $translation"
                      />
                    </xsl:when>
                    <!-- paragraph, paragraph-content -->
                    <xsl:when
                      test="$verticalRel = 'paragraph'  or $verticalRel = 'paragraph-content' ">
                      <!-- TODO : get spacing property of current paragraph -->
                      <xsl:variable name="paragraphBottomSpacing">0</xsl:variable>
                      <xsl:value-of
                        select="$pageHeight - $pageTopMargin - $pageBottomMargin - $paragraphBottomSpacing - $frameHeight - $frameMarginBottom + $translation"
                      />
                    </xsl:when>
                    <!-- frame, frame-content -->
                    <xsl:when test="contains($verticalRel, 'frame')">
                      <xsl:value-of
                        select="$pageHeight - $frameHeight - $frameMarginBottom + $translation"/>
                    </xsl:when>
                    <xsl:otherwise>0</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:choose>
                    <xsl:when test="$parent[1]/@svg:y">
                      <xsl:call-template name="point-measure">
                        <xsl:with-param name="length" select="$parent[1]/@svg:y"/>
                      </xsl:call-template>
                    </xsl:when>
                    <xsl:otherwise>0</xsl:otherwise>
                  </xsl:choose>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:value-of select="$svgy+$recursive_result"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>0</xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <!-- compute margin when rotation -->
  <xsl:template name="ComputeMarginYWithRotation">
    <xsl:param name="angle"/>

    <!-- particular transformation -->
    <xsl:variable name="translationY">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length">
          <xsl:value-of
            select="substring-before(substring-after(substring-after(substring-after(ancestor-or-self::node()[contains(name(), 'draw:') and @draw:transform][1]/@draw:transform,'translate'),'('),' '),')')"
          />
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <!-- frame properties -->
    <xsl:variable name="frameWidth">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length"
          select="ancestor-or-self::node()[contains(name(), 'draw:') and @draw:transform][1]/@svg:width"
        />
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="frameHeight">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length"
          select="ancestor-or-self::node()[contains(name(), 'draw:') and @draw:transform][1]/@svg:height"
        />
      </xsl:call-template>
    </xsl:variable>
    <!-- value used by Word to calculate frame position -->
    <xsl:variable name="frameOffset">
      <xsl:value-of select="round(($frameWidth - $frameHeight) div 2)"/>
    </xsl:variable>
    <!-- special distance with rotation -->
    <xsl:choose>
      <xsl:when test="$angle = 90">
        <xsl:value-of select="$translationY + $frameOffset"/>
      </xsl:when>
      <xsl:when test="$angle = -90">
        <xsl:value-of select="$translationY - $frameWidth + $frameOffset"/>
      </xsl:when>
      <xsl:when test="$angle &gt; 0">
        <xsl:value-of select="$translationY + $frameOffset"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$translationY - $frameWidth + $frameOffset"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- shape properties attributes : size, position, z-index... -->
  <xsl:template name="InsertShapeSize">
    <xsl:param name="shapeProperties"/>

    <!-- width -->
    <xsl:variable name="frameW">
      <xsl:call-template name="ComputeDrawingObjectSize">
        <xsl:with-param name="side">width</xsl:with-param>
        <xsl:with-param name="frame" select="$shapeProperties"/>
        <xsl:with-param name="unit">point</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$frameW != '' and $frameW != 0">
      <xsl:value-of select="concat('width:',$frameW,'pt;')"/>
    </xsl:if>

    <!-- height-->
    <xsl:variable name="frameH">
      <xsl:call-template name="ComputeDrawingObjectSize">
        <xsl:with-param name="side">height</xsl:with-param>
        <xsl:with-param name="frame" select="$shapeProperties"/>
        <xsl:with-param name="unit">point</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$frameH != '' and $frameH != 0">
      <xsl:value-of select="concat('height:',$frameH,'pt;')"/>
    </xsl:if>

  </xsl:template>

  <xsl:template name="InsertShapeRelativeSize">
    <xsl:param name="shapeProperties"/>

    <xsl:if test="$shapeProperties/@style:rel-width or $shapeProperties/@style:rel-height">

      <!-- relative to -->
      <xsl:variable name="relativeTo">
        <xsl:choose>
          <xsl:when test="$shapeProperties/@text:anchor-type = 'page'">
            <xsl:text>page</xsl:text>
          </xsl:when>
          <xsl:when test="$shapeProperties/@text:anchor-type = 'paragraph'">
            <xsl:text>margin</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>margin</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>

      <!-- relative width -->
      <xsl:choose>
        <xsl:when test="contains($shapeProperties/@style:rel-width, 'scale')">
          <!-- warn loss of scaled images (scale, scale-min) -->
          <xsl:message terminate="no">translation.odf2oox.scaledImage</xsl:message>
        </xsl:when>
        <xsl:otherwise>
          <xsl:variable name="relWidth">
            <xsl:choose>
              <xsl:when test="contains($shapeProperties/@style:rel-width,'%')">
                <xsl:value-of select="substring-before($shapeProperties/@style:rel-width,'%')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$shapeProperties/@style:rel-width"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:if test="$relWidth != '' ">
            <xsl:text>mso-width-percent:</xsl:text>
            <xsl:value-of select="number($relWidth) * 10"/>
            <xsl:text>;mso-width-relative:</xsl:text>
            <xsl:value-of select="$relativeTo"/>
            <xsl:text>;</xsl:text>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>

      <!-- relative height -->
      <xsl:choose>
        <xsl:when test="contains($shapeProperties/@style:rel-height, 'scale')">
          <!-- warn loss of scaled images (scale, scale-min) -->
          <xsl:message terminate="no">translation.odf2oox.scaledImage</xsl:message>
        </xsl:when>
        <xsl:otherwise>
          <xsl:variable name="relHeight">
            <xsl:choose>
              <xsl:when test="contains($shapeProperties/@style:rel-height,'%')">
                <xsl:value-of select="substring-before($shapeProperties/@style:rel-height,'%')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$shapeProperties/@style:rel-height"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:if test="$relHeight != '' ">
            <xsl:text>mso-height-percent:</xsl:text>
            <xsl:value-of select="number($relHeight) * 10"/>
            <xsl:text>;mso-height-relative:</xsl:text>
            <xsl:value-of select="$relativeTo"/>
            <xsl:text>;</xsl:text>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>

    </xsl:if>

  </xsl:template>

  <xsl:template name="InsertShapeZindex">
    <xsl:param name="shapeStyle"/>
    <xsl:param name="shapeProperties"/>

    <xsl:variable name="frameWrap">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:wrap</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="runThrought">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:run-through</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>


    <!--z-index that we need to convert properly openoffice wrap-throught property -->
    <xsl:variable name="zIndex">
      <xsl:choose>
        <xsl:when test="$frameWrap='run-through' 
          and $runThrought='background'"
          >-251658240</xsl:when>
        <xsl:when test="$frameWrap='run-through' 
          and not($runThrought)">251658240</xsl:when>
        <xsl:otherwise>
          <xsl:choose>
            <xsl:when test="$shapeProperties/@draw:z-index">
              <xsl:value-of select="2 + $shapeProperties/@draw:z-index"/>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:value-of select="concat('z-index:', $zIndex, ';')"/>
  </xsl:template>

  <xsl:template name="InsertShapePositionRelative">
    <xsl:param name="shapeStyle"/>
    <xsl:param name="shapeProperties"/>

    <xsl:variable name="anchor">
      <xsl:value-of select="$shapeProperties/@text:anchor-type"/>
    </xsl:variable>

    <!-- text-box horizontal relative position -->
    <xsl:variable name="horizontalRel">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:horizontal-rel</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="wrappedPara">
      <xsl:variable name="wrapping">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
          <xsl:with-param name="attribName">style:wrap</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:if test="$wrapping = 'parallel' ">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
          <xsl:with-param name="attribName">style:number-wrapped-paragraphs</xsl:with-param>
        </xsl:call-template>
      </xsl:if>
    </xsl:variable>
    <xsl:choose>
      <!-- inline image -->
      <xsl:when test="$wrappedPara = 1">mso-position-horizontal-relative:char;</xsl:when>
      <!-- direct anchored frame -->
      <xsl:when test="$anchor = 'as-char' ">mso-position-horizontal-relative:char;</xsl:when>
      <xsl:when test="$anchor = 'page' and $shapeProperties/@svg:x">
        <xsl:choose>
          <!-- page-content -->
          <xsl:when test="$horizontalRel='page-content'">mso-position-horizontal-relative:margin;</xsl:when>
          <!-- any other case -->
          <xsl:otherwise>mso-position-horizontal-relative:page;</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <!-- page-content -->
          <xsl:when test="$horizontalRel='page-content' ">mso-position-horizontal-relative:margin;</xsl:when>
          <!-- page, page-start-margin, page-end-margin -->
          <xsl:when test="contains($horizontalRel, 'page')">mso-position-horizontal-relative:page;</xsl:when>
          <!-- paragraph, paragraph-content, paragraph-start-margin, paragraph-end-margin -->
          <xsl:when test="contains($horizontalRel, 'paragraph')">mso-position-horizontal-relative:text;</xsl:when>
          <!-- frame, frame-content, frame-start-margin, frame-end-margin -->
          <xsl:when test="contains($horizontalRel, 'frame')">mso-position-horizontal-relative:text;</xsl:when>
          <!-- char -->
          <xsl:when test="$horizontalRel = 'char' ">mso-position-horizontal-relative:char;</xsl:when>
          <xsl:otherwise>
            <!-- no default value suggested. use anchor -->
            <xsl:choose>
              <xsl:when test="$anchor = 'page' ">mso-position-horizontal-relative:page;</xsl:when>
              <xsl:when test="$anchor = 'paragraph' or $anchor = 'frame' "
                >mso-position-horizontal-relative:text;</xsl:when>
              <xsl:when test="$anchor = 'char' ">mso-position-horizontal-relative:char;</xsl:when>
              <xsl:otherwise>
                <!-- as-char anchor already handled (cf above). In case nothing is ever specified : use default = text -->
                <xsl:text>mso-position-horizontal-relative:text;</xsl:text>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>

    <!-- text-box vertical relative position -->
    <xsl:variable name="verticalRel">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:vertical-rel</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:choose>
      <!-- inline image -->
      <xsl:when test="$wrappedPara = 1">mso-position-vertical-relative:line;</xsl:when>
      <!-- direct anchored frame -->
      <xsl:when test="$anchor = 'as-char' ">mso-position-horizontal-relative:line;</xsl:when>
      <xsl:when test="$anchor = 'page' and $shapeProperties/@svg:y">
        <xsl:choose>
          <!-- page-content -->
          <xsl:when test="$verticalRel = 'page-content' ">mso-position-vertical-relative:margin;</xsl:when>
          <!-- any other case -->
          <xsl:otherwise>mso-position-vertical-relative:page;</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <!-- page -->
          <xsl:when test="$verticalRel='page'">mso-position-vertical-relative:page;</xsl:when>
          <!-- page-content -->
          <xsl:when test="$verticalRel='page-content'">mso-position-vertical-relative:margin;</xsl:when>
          <!-- paragraph -->
          <xsl:when test="$verticalRel='paragraph'">mso-position-vertical-relative:text;</xsl:when>
          <!-- paragraph-content -->
          <xsl:when test="$verticalRel='paragraph-content'">mso-position-vertical-relative:text;</xsl:when>
          <!-- frame, frame-content -->
          <xsl:when test="contains($verticalRel, 'frame')">mso-position-vertical-relative:text;</xsl:when>
          <!-- char, line, baseline, text -->
          <xsl:when
            test="$verticalRel='char' or $verticalRel='line' or $verticalRel='baseline' or $verticalRel='text'"
            >mso-position-vertical-relative:char;</xsl:when>
          <xsl:otherwise>
            <!-- no default value suggested. use anchor -->
            <xsl:choose>
              <xsl:when test="$anchor = 'page' ">mso-position-horizontal-relative:page;</xsl:when>
              <xsl:when test="$anchor = 'paragraph' or $anchor = 'frame' "
                >mso-position-horizontal-relative:text;</xsl:when>
              <xsl:when test="$anchor = 'char' ">mso-position-horizontal-relative:line;</xsl:when>
              <xsl:otherwise>
                <!-- as-char anchor already handled (cf above). In case nothing is ever specified : use default = text -->
                <xsl:text>mso-position-horizontal-relative:text;</xsl:text>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertShapePosition">
    <xsl:param name="shapeStyle"/>
    <xsl:param name="shapeProperties"/>
    
    <xsl:variable name="graphicProps" select="$shapeStyle/style:graphic-properties"/>
    <xsl:variable name="anchor">
      <xsl:value-of select="$shapeProperties/@text:anchor-type"/>
    </xsl:variable>
    <xsl:variable name="wrappedPara">
      <xsl:variable name="wrapping">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
          <xsl:with-param name="attribName">style:wrap</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:if test="$wrapping = 'parallel' ">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
          <xsl:with-param name="attribName">style:number-wrapped-paragraphs</xsl:with-param>
        </xsl:call-template>
      </xsl:if>
    </xsl:variable>
    
    <!-- if inline image, no positioning -->
    <xsl:if test="not($wrappedPara=1 or $anchor='as-char')">
      
      <!-- two cases : absolute (=>define margin-left and margin-top), or mso-position-horizontal / -vertical -->
      <xsl:variable name="horizontalPos">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
          <xsl:with-param name="attribName">style:horizontal-pos</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="horizontalRel">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
          <xsl:with-param name="attribName">style:horizontal-rel</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="marginLeft">
        <xsl:call-template name="GetValue">
          <xsl:with-param name="length">
            <xsl:call-template name="GetGraphicProperties">
              <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
              <xsl:with-param name="attribName">fo:margin-left</xsl:with-param>
            </xsl:call-template>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="marginRight">
        <xsl:call-template name="GetValue">
          <xsl:with-param name="length">
            <xsl:call-template name="GetGraphicProperties">
              <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
              <xsl:with-param name="attribName">fo:margin-right</xsl:with-param>
            </xsl:call-template>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="xCoordinate">
        <xsl:call-template name="GetValue">
          <xsl:with-param name="length">
            <xsl:call-template name="GetGraphicProperties">
              <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
              <xsl:with-param name="attribName">svg:x</xsl:with-param>
            </xsl:call-template>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="verticalPos">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
          <xsl:with-param name="attribName">style:vertical-pos</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="verticalRel">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
          <xsl:with-param name="attribName">style:vertical-rel</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="marginTop">
        <xsl:call-template name="GetValue">
          <xsl:with-param name="length">
            <xsl:call-template name="GetGraphicProperties">
              <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
              <xsl:with-param name="attribName">fo:margin-top</xsl:with-param>
            </xsl:call-template>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="marginBottom">
        <xsl:call-template name="GetValue">
          <xsl:with-param name="length">
            <xsl:call-template name="GetGraphicProperties">
              <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
              <xsl:with-param name="attribName">fo:margin-bottom</xsl:with-param>
            </xsl:call-template>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="yCoordinate">
        <xsl:call-template name="GetValue">
          <xsl:with-param name="length">
            <xsl:call-template name="GetGraphicProperties">
              <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
              <xsl:with-param name="attribName">svg:y</xsl:with-param>
            </xsl:call-template>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="rotation">
        <xsl:if test="contains($shapeProperties/@draw:transform,'rotate')">
          <xsl:call-template name="DegreesAngle">
            <xsl:with-param name="angle">
              <xsl:value-of
                select="substring-before(substring-after(substring-after($shapeProperties/@draw:transform,'rotate'),'('),')')"
              />
            </xsl:with-param>
            <xsl:with-param name="revert">true</xsl:with-param>
          </xsl:call-template>
        </xsl:if>
      </xsl:variable>

      <!-- 
        declare an absolute positioning (ignored if margin-left/rigt=0)
        NB: it should not be compulsory to declare absolute positioning, but it causes troubles for images embedded in text-boxes in Word if not declared.
      -->
      <xsl:text>position:absolute;</xsl:text>

      <!-- compute margin with respect to frame spacing to content, paragraph/page margins... -->
      <xsl:text>margin-left:</xsl:text>
      <xsl:variable name="valX">
        <xsl:if
          test="$horizontalPos = 'from-left' or $horizontalPos='from-inside' or ($marginLeft != '' and $marginLeft != 0 ) or ($marginRight != '' and $marginRight != 0 ) ">
          <xsl:choose>
            <!-- if rotation, revert X and Y -->
            <xsl:when test="$rotation != '' ">
              <xsl:call-template name="ComputeMarginXWithRotation">
                <xsl:with-param name="angle" select="$rotation"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="ComputeMarginX">
                <xsl:with-param name="parent"
                  select="ancestor-or-self::node()[contains(name(), 'draw:')]"/>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="number($valX) and $valX != 0">
          <xsl:value-of select="$valX"/>
          <xsl:text>pt</xsl:text>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
      <xsl:text>;</xsl:text>

      <!-- compute margin with respect to frame spacing to content, paragraph/page margins... -->
      <xsl:text>margin-top:</xsl:text>
      <xsl:variable name="valY">
        <xsl:if
          test="$verticalPos = 'from-top' or ($marginTop != '' and $marginTop != 0 ) or ($marginBottom != '' and $marginBottom != 0 ) ">
          <xsl:choose>
            <!-- if rotation, revert X and Y -->
            <xsl:when test="$rotation != '' ">
              <xsl:call-template name="ComputeMarginYWithRotation">
                <xsl:with-param name="angle" select="$rotation"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="ComputeMarginY">
                <xsl:with-param name="parent"
                  select="ancestor-or-self::node()[contains(name(), 'draw:')]"/>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="number($valY) and $valY != 0">
          <xsl:value-of select="$valY"/>
          <xsl:text>pt</xsl:text>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
      <xsl:text>;</xsl:text>

      <!-- horizontal position (overrides margin-left) -->
      <xsl:choose>
        <!-- centered position -->
        <xsl:when test="$horizontalPos='center' 
                  and not(contains($horizontalRel, '-start-margin') or contains($horizontalRel, '-end-margin'))">
          <xsl:text>mso-position-horizontal:center;</xsl:text>
        </xsl:when>
        <!-- 
        relative to page, page-content, paragraph, paragraph-content, char 
        makz: Do this not if a frame is left- or right-aligned to page.
              In this case the position must be absolute because margins are ignored in word if aligned to page.
        -->
        <xsl:when test="($horizontalRel='page' and not($horizontalPos='left' or $horizontalPos='right'))
                  or $horizontalRel='page-content' 
                  or $horizontalRel='paragraph'  
                  or $horizontalRel='paragraph-content' 
                  or $horizontalRel='char'">
          <xsl:value-of select="concat('mso-position-horizontal:', $horizontalPos, ';')"/>
        </xsl:when>
      </xsl:choose>

      <!-- vertical position (overrides margin-top) -->
      <xsl:choose>
        <!-- centered position -->
        <xsl:when test="$verticalPos = 'middle' ">
          <xsl:text>mso-position-vertical:center;</xsl:text>
        </xsl:when>
        <!-- relative to page, page-content, paragraph, paragraph-content, char -->
        <xsl:when test="$verticalRel='page' or $verticalRel = 'page-content' or $verticalRel = 'paragraph'  or $verticalRel = 'paragraph-content'  or $verticalRel = 'char' 
          or $verticalRel = 'char' or $verticalRel = 'baseline' or $verticalRel = 'line' or $verticalRel = 'text' ">
          <xsl:text>mso-position-vertical:</xsl:text>
          <xsl:choose>
            <xsl:when test="$verticalPos='middle'">center</xsl:when>
            <xsl:when test="$verticalPos='below'">bottom</xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$verticalPos"/>
            </xsl:otherwise>
          </xsl:choose>
          <xsl:text>;</xsl:text>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertShapeMargin">
    <xsl:param name="shapeStyle"/>
    <xsl:param name="shapeProperties"/>

    <!-- wrapping of text (horizontal adjustment) -->
    <xsl:if test="$shapeProperties/@fo:min-width or $shapeStyle/style:graphic-properties/@draw:auto-grow-width = 'true' ">
      <!--<xsl:text>mso-wrap-style:none;</xsl:text>-->
    </xsl:if>

    <!--text-box spacing/margins -->
    <xsl:variable name="marginL">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">fo:margin-left</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="marginT">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">fo:margin-top</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="marginR">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">fo:margin-right</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="marginB">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">fo:margin-bottom</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$marginL != '' ">
      <xsl:text>mso-wrap-distance-left:</xsl:text>
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length" select="$marginL"/>
        <xsl:with-param name="round" select="'false'"/>
      </xsl:call-template>
      <xsl:text>pt;</xsl:text>
    </xsl:if>
    <xsl:if test="$marginT != '' ">
      <xsl:text>mso-wrap-distance-top:</xsl:text>
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length" select="$marginT"/>
        <xsl:with-param name="round" select="'false'"/>
      </xsl:call-template>
      <xsl:text>pt;</xsl:text>
    </xsl:if>
    <xsl:if test="$marginR != '' ">
      <xsl:text>mso-wrap-distance-right:</xsl:text>
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length" select="$marginR"/>
        <xsl:with-param name="round" select="'false'"/>
      </xsl:call-template>
      <xsl:text>pt;</xsl:text>
    </xsl:if>
    <xsl:if test="$marginB != '' ">
      <xsl:text>mso-wrap-distance-bottom:</xsl:text>
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length" select="$marginB"/>
        <xsl:with-param name="round" select="'false'"/>
      </xsl:call-template>
      <xsl:text>pt;</xsl:text>
    </xsl:if>

  </xsl:template>

  <xsl:template name="InsertShapeFill">
    <xsl:param name="shapeStyle"/>

    <xsl:variable name="shapefillColor">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">fo:background-color</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!-- frame color is dependent on page background color in some cases -->
    <xsl:variable name="fillColor">
      <xsl:choose>
        <xsl:when test="$shapefillColor = 'transparent' or $shapefillColor = '' ">
          <!-- when no fill is set for frame it should take background color of page (as it is in ODF) -->
          <xsl:for-each select="document('styles.xml')">
            <xsl:variable name="defaultBgColor"
              select="key('page-layouts', $default-master-style/@style:page-layout-name)[1]/style:page-layout-properties/@fo:background-color"/>
            <xsl:value-of select="$defaultBgColor"/>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$shapefillColor"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <!--fill color-->
    <xsl:if test="$fillColor != '' ">
      <xsl:attribute name="fillcolor">
        <xsl:value-of select="$fillColor"/>
      </xsl:attribute>
    </xsl:if>

  </xsl:template>

  <xsl:template name="InsertShapeShadow">
    <xsl:param name="shapeStyle"/>

    <xsl:variable name="shadow">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:shadow</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$shadow != ''  and $shadow != 'none' ">
      <v:shadow on="t">
        <xsl:if test="substring-before($shadow, ' ') != 'none' ">
          <xsl:attribute name="color">
            <xsl:value-of select="substring-before($shadow, ' ')"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:variable name="firstShadow">
          <xsl:value-of select=" substring-before(substring-after($shadow, ' '), ' ')"/>
        </xsl:variable>
        <xsl:variable name="secondShadow">
          <xsl:value-of select=" substring-after(substring-after($shadow, ' '), ' ')"/>
        </xsl:variable>
        <xsl:if test="$firstShadow != '' and $secondShadow != '' ">
          <xsl:attribute name="offset">
            <xsl:call-template name="point-measure">
              <xsl:with-param name="length" select="$firstShadow"/>
            </xsl:call-template>
            <xsl:text>pt,</xsl:text>
            <xsl:call-template name="point-measure">
              <xsl:with-param name="length" select="$secondShadow"/>
            </xsl:call-template>
            <xsl:text>pt</xsl:text>
          </xsl:attribute>
        </xsl:if>
      </v:shadow>
    </xsl:if>

  </xsl:template>

  <xsl:template name="InsertShapeBorders">
    <xsl:param name="shapeStyle"/>

    <xsl:variable name="border">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">fo:border</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!--borders-->
    <xsl:choose>
      <!-- no border in current style -->
      <xsl:when test="$border='' or $border = 'none' ">
        <xsl:attribute name="stroked">f</xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="strokeColor" select="substring-after($border,'#')"/>
        <xsl:variable name="strokeWeight">
          <xsl:call-template name="point-measure">
            <xsl:with-param name="length" select="substring-before($border,' ')"/>
            <xsl:with-param name="round">
              <xsl:text>false</xsl:text>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="styleBorderLine">
          <xsl:call-template name="GetGraphicProperties">
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
            <xsl:with-param name="attribName">style:border-line-width</xsl:with-param>
          </xsl:call-template>
        </xsl:variable>

        <xsl:if test="$strokeColor != '' ">
          <xsl:attribute name="strokecolor">
            <xsl:value-of select="concat('#', $strokeColor)"/>
          </xsl:attribute>
        </xsl:if>

        <xsl:if test="$strokeWeight != '' ">
          <xsl:attribute name="strokeweight">
            <xsl:value-of select="concat($strokeWeight,'pt')"/>
          </xsl:attribute>
        </xsl:if>

        <!--  line styles -->
        <xsl:if test="substring-before(substring-after($border,' ' ),' ' ) != 'solid' ">
          <v:stroke>
            <xsl:attribute name="linestyle">
              <xsl:choose>
                <xsl:when test="$styleBorderLine">

                  <xsl:variable name="innerLineWidth">
                    <xsl:call-template name="point-measure">
                      <xsl:with-param name="length" select="substring-before($styleBorderLine,' ' )"
                      />
                    </xsl:call-template>
                  </xsl:variable>

                  <xsl:variable name="outerLineWidth">
                    <xsl:call-template name="point-measure">
                      <xsl:with-param name="length"
                        select="substring-after(substring-after($styleBorderLine,' ' ),' ' )"/>
                    </xsl:call-template>
                  </xsl:variable>

                  <xsl:if test="$innerLineWidth = $outerLineWidth">thinThin</xsl:if>
                  <xsl:if test="$innerLineWidth > $outerLineWidth">thinThick</xsl:if>
                  <xsl:if test="$outerLineWidth > $innerLineWidth  ">thickThin</xsl:if>

                </xsl:when>
              </xsl:choose>
            </xsl:attribute>
          </v:stroke>
        </xsl:if>
      </xsl:otherwise>

      <!--default scenario-->
      <!--     <xsl:otherwise>
        <xsl:attribute name="stroked">f</xsl:attribute>
        </xsl:otherwise>
       -->
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertShapeRotation">
    <xsl:param name="shapeProperties"/>

    <xsl:if test="contains($shapeProperties/@draw:transform,'rotate')">
      <xsl:text>rotation:</xsl:text>
      <xsl:variable name="angle">
        <xsl:call-template name="DegreesAngle">
          <xsl:with-param name="angle">
            <xsl:value-of
              select="substring-before(substring-after(substring-after($shapeProperties/@draw:transform,'rotate'),'('),')')"
            />
          </xsl:with-param>
          <xsl:with-param name="revert">true</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:value-of select="$angle"/>
      <xsl:text>;</xsl:text>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertShapeFillProperties">
    <xsl:param name="shapeStyle"/>

    <xsl:variable name="fillProperty">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">draw:fill</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!-- opacity -->
    <xsl:variable name="bgTransparency">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:background-transparency</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="opacity">
      <xsl:variable name="draw-opacity">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
          <xsl:with-param name="attribName">draw:opacity</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$draw-opacity != '' ">
          <xsl:value-of select="$draw-opacity"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:if test="$bgTransparency != '' ">
            <xsl:value-of
              select="concat((100 - number(substring-before($bgTransparency,'%'))), '%')"/>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:if test="$fillProperty != 'none' ">
      <v:fill>
        <xsl:if test="$opacity != '' ">
          <xsl:attribute name="opacity">
            <xsl:value-of select="$opacity"/>
          </xsl:attribute>
        </xsl:if>
        <!-- other fill properties -->
        <xsl:choose>
          <xsl:when test="$fillProperty = 'solid' ">
            <xsl:attribute name="color">
              <xsl:call-template name="GetGraphicProperties">
                <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
                <xsl:with-param name="attribName">draw:fill-color</xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="$fillProperty = 'gradient' ">
            <xsl:call-template name="InsertGradientFill">
              <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
            </xsl:call-template>
          </xsl:when>
        </xsl:choose>
      </v:fill>
    </xsl:if>

  </xsl:template>

  <!-- text vertical anchor -->
  <xsl:template name="InsertTextAnchor">
    <xsl:param name="shapeStyle"/>

    <xsl:variable name="verticalAlign">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">draw:textarea-vertical-align</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:text>v-text-anchor:</xsl:text>
    <xsl:choose>
      <xsl:when test="$verticalAlign = '' or $verticalAlign = 'top' ">top</xsl:when>
      <xsl:when test="$verticalAlign = 'middle' or $verticalAlign = 'justify' ">middle</xsl:when>
      <xsl:when test="$verticalAlign = 'bottom' ">bottom</xsl:when>
      <xsl:otherwise>top</xsl:otherwise>
    </xsl:choose>
    <xsl:text>;</xsl:text>
  </xsl:template>

  <xsl:template name="InsertGradientFill">
    <xsl:param name="shapeStyle"/>

    <!-- simple linear gradient -->
    <xsl:variable name="gradientName">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">draw:fill-gradient-name</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:for-each
      select="document('styles.xml')/office:document-styles/office:styles/draw:gradient[@draw:name = $gradientName]">
      <!-- radial gradients not handled yet -->
      <xsl:attribute name="type">gradient</xsl:attribute>
      <xsl:if test="@draw:angle">
        <xsl:attribute name="angle">
          <xsl:value-of select="round(number(@draw:angle) div 10)"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@draw:end-color">
        <xsl:attribute name="color">
          <xsl:value-of select="@draw:end-color"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@draw:start-color">
        <xsl:attribute name="color2">
          <xsl:value-of select="@draw:start-color"/>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

  <!--converts oo frame style properties to shape properties for text-box-->
  <xsl:template name="InsertShapeProperties">
    <xsl:param name="shapeStyle"/>
    <xsl:param name="shapeProperties"/>
    <!-- report lost properties -->
    <xsl:if test="$shapeStyle/style:graphic-properties/@draw:textarea-vertical-align != 'top' ">
      <xsl:message terminate="no">translation.odf2oox.valignInsideTextbox</xsl:message>
    </xsl:if>

    <xsl:if test="$shapeProperties/@draw:name">
      <xsl:attribute name="id">
        <xsl:value-of select="$shapeProperties/@draw:name"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="ancestor::draw:a/@xlink:href">
      <xsl:attribute name="href">
        <!-- avoid empty target -->
        <xsl:choose>
          <xsl:when test="contains(ancestor::draw:a/@xlink:href, './')">
            <xsl:value-of select="substring-after(ancestor::draw:a/@xlink:href, '../')"/>
          </xsl:when>
          <xsl:when test="string-length(ancestor::draw:a/@xlink:href) &gt; 0">
            <xsl:value-of select="ancestor::draw:a/@xlink:href"/>
          </xsl:when>
          <xsl:otherwise>/</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="ancestor::draw:a/@office:target-frame-name">
      <xsl:attribute name="target">
        <xsl:value-of select="ancestor::draw:a/@office:target-frame-name"/>
      </xsl:attribute>
    </xsl:if>

    <xsl:attribute name="style">

      <xsl:call-template name="InsertShapeSize">
        <xsl:with-param name="shapeProperties" select="$shapeProperties"/>
      </xsl:call-template>

      <xsl:call-template name="InsertShapeRelativeSize">
        <xsl:with-param name="shapeProperties" select="$shapeProperties"/>
      </xsl:call-template>

      <xsl:call-template name="InsertShapeZindex">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="shapeProperties" select="$shapeProperties"/>
      </xsl:call-template>

      <xsl:call-template name="InsertShapePositionRelative">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="shapeProperties" select="$shapeProperties"/>
      </xsl:call-template>

      <xsl:call-template name="InsertShapePosition">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="shapeProperties" select="$shapeProperties"/>
      </xsl:call-template>

      <xsl:call-template name="InsertShapeMargin">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="shapeProperties" select="$shapeProperties"/>
      </xsl:call-template>

      <xsl:call-template name="InsertShapeRotation">
        <xsl:with-param name="shapeProperties" select="$shapeProperties"/>
      </xsl:call-template>

      <xsl:call-template name="InsertTextAnchor">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>
    </xsl:attribute>

    <xsl:call-template name="InsertShapeFill">
      <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
    </xsl:call-template>

    <xsl:call-template name="InsertShapeBorders">
      <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
    </xsl:call-template>

    <xsl:call-template name="InsertShapeFillProperties">
      <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
    </xsl:call-template>

    <xsl:call-template name="InsertShapeShadow">
      <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
    </xsl:call-template>
  </xsl:template>

  <!-- finds shape graphic property recursively in style or parent style -->
  <xsl:template name="GetGraphicProperties">
    <xsl:param name="attribName"/>
    <xsl:param name="shapeStyle"/>

    <xsl:choose>
      <xsl:when test="$shapeStyle/style:graphic-properties/attribute::node()[name() = $attribName] ">
        <xsl:value-of
          select="$shapeStyle/style:graphic-properties/attribute::node()[name() = $attribName]"/>
      </xsl:when>

      <xsl:when test="$shapeStyle/@style:parent-style-name">
        <xsl:variable name="parentStyleName">
          <xsl:value-of select="$shapeStyle/@style:parent-style-name"/>
        </xsl:variable>

        <xsl:variable name="parentStyle"
          select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $parentStyleName]"/>

        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$parentStyle"/>
          <xsl:with-param name="attribName" select="$attribName"/>
        </xsl:call-template>
      </xsl:when>
      <!--    <xsl:otherwise></xsl:otherwise>-->
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertShapeWrap">
    <xsl:param name="shapeStyle"/>

    <xsl:variable name="frameWrap">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:wrap</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!--frame wrap-->
    <xsl:choose>
      <xsl:when test="parent::node()[contains(name(), 'draw:')]/@text:anchor-type = 'as-char' ">
        <w10:wrap type="none"/>
        <w10:anchorlock/>
      </xsl:when>
      <xsl:when test="$frameWrap = 'left' ">
        <w10:wrap type="square" side="left"/>
      </xsl:when>
      <xsl:when test="$frameWrap = 'right' ">
        <w10:wrap type="square" side="right"/>
      </xsl:when>
      <xsl:when test="$frameWrap = 'none'">
        <xsl:message terminate="no">translation.odf2oox.shapeTopBottomWrapping</xsl:message>
        <w10:wrap type="topAndBottom"/>
      </xsl:when>
      <xsl:when test="$frameWrap = 'parallel' ">
        <xsl:variable name="wrappedPara">
          <xsl:call-template name="GetGraphicProperties">
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
            <xsl:with-param name="attribName">style:number-wrapped-paragraphs</xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="$wrappedPara = 1">
            <w10:wrap type="none"/>
            <w10:anchorlock/>
          </xsl:when>
          <xsl:otherwise>
            <w10:wrap type="square"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="$frameWrap = 'dynamic' ">
        <w10:wrap type="square" side="largest"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!--inserts text-box into shape element -->
  <xsl:template name="InsertTextBox">
    <xsl:param name="shapeStyle"/>

    <v:textbox>
      <xsl:attribute name="style">
        <xsl:if test="@draw:chain-next-name">
          <xsl:text>mso-next-textbox:#</xsl:text>
          <xsl:value-of select="@draw:chain-next-name"/>
          <xsl:text>;</xsl:text>
        </xsl:if>
        <xsl:if test="self::draw:text-box">
          <!-- comment : OpenOffice will rather use fo:min-width when parent style defined, and draw:auto-grow-width if not -->
          <xsl:choose>
            <xsl:when test="$shapeStyle/@style:parent-style-name">
              <!--
                makz: commented due to bug 1669046. Regression could be possible.
                
                <xsl:if
                  test="@fo:min-height 
                  or parent::draw:frame/@fo:min-width
                  or $shapeStyle/style:graphic-properties/@draw:auto-grow-width = 'true'
                  or $shapeStyle/style:graphic-properties/@draw:auto-grow-height = 'true'">
                -->
              <xsl:if
                test="parent::draw:frame/@fo:min-width
                or $shapeStyle/style:graphic-properties/@draw:auto-grow-width = 'true'
                or $shapeStyle/style:graphic-properties/@draw:auto-grow-height = 'true'">
                <xsl:text>mso-fit-shape-to-text:t;</xsl:text>
              </xsl:if>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="not(parent::node()='draw:frame') 
                and (not($shapeStyle/style:graphic-properties/@draw:auto-grow-width = 'false') 
                or not($shapeStyle/style:graphic-properties/@draw:auto-grow-height = 'false'))">
                <xsl:text>mso-fit-shape-to-text:t;</xsl:text>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
        <xsl:if test="contains(parent::draw:frame/@draw:transform,'rotate')">
          <xsl:variable name="angle">
            <xsl:call-template name="DegreesAngle">
              <xsl:with-param name="angle">
                <xsl:value-of
                  select="substring-before(substring-after(substring-after(parent::draw:frame/@draw:transform,'rotate'),'('),')')"
                />
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test="$angle = 90">
              <xsl:text>layout-flow:vertical;mso-layout-flow-alt:bottom-to-top;</xsl:text>
            </xsl:when>
            <xsl:when test="$angle = -90">
              <xsl:text>layout-flow:vertical;</xsl:text>
            </xsl:when>
            <xsl:when test="$angle = 180">
              <xsl:text>mso-rotate:180:</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:message terminate="no"
              >translation.odf2oox.textOrientationInsideTextbox</xsl:message>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
      </xsl:attribute>

      <xsl:call-template name="InsertTextBoxInset">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

      <w:txbxContent>
        <xsl:for-each select="child::node()">

          <xsl:choose>
            <!--   ignore embedded text-box and other shapes because word doesn't support it-->
            <xsl:when test="self::draw:frame/draw:text-box or self::draw:frame/draw:rect or self::draw:frame/draw:custom-shape">
              <xsl:message terminate="no">translation.odf2oox.nestedFrames</xsl:message>
            </xsl:when>

            <!-- warn loss of positioning for embedded drawn objects or pictures -->
            <xsl:when test="contains(name(), 'draw:')">
              <xsl:message terminate="no">translation.odf2oox.positionInsideTextbox</xsl:message>
              <w:p>
                <xsl:apply-templates select="." mode="paragraph"/>
              </w:p>
            </xsl:when>

            <!-- frames with top-and-bottom wrapping inside text-boxes -->
            <xsl:when test="draw:frame">
              <xsl:variable name="wrapping">
                <xsl:call-template name="GetGraphicProperties">
                  <xsl:with-param name="shapeStyle"
                    select="key('styles', draw:frame/@draw:style-name)"/>
                  <xsl:with-param name="attribName">style:wrap</xsl:with-param>
                </xsl:call-template>
              </xsl:variable>
              <xsl:if
                test="$wrapping = 'none' and not(draw:frame/@text:anchor-type='as-char') and not(draw:frame/descendant::draw:text-box)">
                <w:p>
                  <w:pPr>
                    <xsl:call-template name="InsertPicturePropertiesInFrame"/>
                  </w:pPr>
                  <xsl:apply-templates select="draw:frame" mode="paragraph"/>
                </w:p>
              </xsl:if>
              <xsl:apply-templates select="."/>
            </xsl:when>

            <!--default scenario-->
            <xsl:otherwise>
              <xsl:apply-templates select="."/>
            </xsl:otherwise>
          </xsl:choose>

        </xsl:for-each>
      </w:txbxContent>

    </v:textbox>
  </xsl:template>

  <!--converts oo frame padding into inset for text-box -->
  <xsl:template name="InsertTextBoxInset">
    <xsl:param name="shapeStyle"/>

    <xsl:attribute name="inset">
      <!-- left padding -->
      <xsl:call-template name="milimeter-measure">
        <xsl:with-param name="length">
          <xsl:call-template name="GetFramePadding">
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
            <xsl:with-param name="side">left</xsl:with-param>
          </xsl:call-template>
        </xsl:with-param>
        <xsl:with-param name="round">false</xsl:with-param>
      </xsl:call-template>
      <xsl:text>mm,</xsl:text>
      <!-- top padding -->
      <xsl:call-template name="milimeter-measure">
        <xsl:with-param name="length">
          <xsl:call-template name="GetFramePadding">
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
            <xsl:with-param name="side">top</xsl:with-param>
          </xsl:call-template>
        </xsl:with-param>
        <xsl:with-param name="round">false</xsl:with-param>
      </xsl:call-template>
      <xsl:text>mm,</xsl:text>
      <!-- right padding -->
      <xsl:call-template name="milimeter-measure">
        <xsl:with-param name="length">
          <xsl:call-template name="GetFramePadding">
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
            <xsl:with-param name="side">right</xsl:with-param>
          </xsl:call-template>
        </xsl:with-param>
        <xsl:with-param name="round">false</xsl:with-param>
      </xsl:call-template>
      <xsl:text>mm,</xsl:text>
      <!-- bottom padding -->
      <xsl:call-template name="milimeter-measure">
        <xsl:with-param name="length">
          <xsl:call-template name="GetFramePadding">
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
            <xsl:with-param name="side">bottom</xsl:with-param>
          </xsl:call-template>
        </xsl:with-param>
        <xsl:with-param name="round">false</xsl:with-param>
      </xsl:call-template>
      <xsl:text>mm</xsl:text>
    </xsl:attribute>
  </xsl:template>

  <!-- compute the padding of a border -->
  <xsl:template name="GetFramePadding">
    <xsl:param name="shapeStyle"/>
    <xsl:param name="side"/>

    <xsl:choose>
      <!-- priority to border padding -->
      <xsl:when
        test="$shapeStyle/style:graphic-properties/@*[name() = concat('fo:padding-', $side)]">
        <xsl:value-of
          select="$shapeStyle/style:graphic-properties/@*[name() = concat('fo:padding-', $side)]"/>
      </xsl:when>
      <!-- otherwise if padding attribute -->
      <xsl:when test="$shapeStyle/style:graphic-properties/@fo:padding">
        <xsl:value-of select="$shapeStyle/style:graphic-properties/@fo:padding"/>
      </xsl:when>
      <!-- otherwise look parent style -->
      <xsl:when test="$shapeStyle/@style:parent-style-name">
        <xsl:variable name="parentStyle"
          select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $shapeStyle/@style:parent-style-name]"/>
        <xsl:call-template name="GetFramePadding">
          <xsl:with-param name="shapeStyle" select="$parentStyle"/>
          <xsl:with-param name="side" select="$side"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>0</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- compute the border of a frame -->
  <xsl:template name="GetFrameBorder">
    <xsl:param name="shapeStyle"/>
    <xsl:param name="side"/>

    <xsl:choose>
      <!-- priority to side border -->
      <xsl:when test="$shapeStyle/style:graphic-properties/@*[name() = concat('fo:border-', $side)]">
        <xsl:call-template name="GetConsistentBorderValue">
          <xsl:with-param name="side" select="$side"/>
          <xsl:with-param name="borderAttibute"
            select="$shapeStyle/style:graphic-properties/@*[name() = concat('fo:border-', $side)]"/>
          <xsl:with-param name="node" select="$shapeStyle/style:graphic-properties"/>
        </xsl:call-template>
      </xsl:when>
      <!-- otherwise if border attribute -->
      <xsl:when test="$shapeStyle/style:graphic-properties/@fo:border">
        <xsl:call-template name="GetConsistentBorderValue">
          <xsl:with-param name="side" select="$side"/>
          <xsl:with-param name="borderAttibute"
            select="$shapeStyle/style:graphic-properties/@fo:border"/>
          <xsl:with-param name="node" select="$shapeStyle/style:graphic-properties"/>
        </xsl:call-template>
      </xsl:when>
      <!-- otherwise look parent style -->
      <xsl:when test="$shapeStyle/@style:parent-style-name">
        <xsl:variable name="parentStyle"
          select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $shapeStyle/@style:parent-style-name]"/>
        <xsl:call-template name="GetFramePadding">
          <xsl:with-param name="shapeStyle" select="$parentStyle"/>
          <xsl:with-param name="side" select="$side"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>0</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Inserts the style of a text box -->
  <xsl:template name="InsertTextBoxStyle">
    <xsl:variable name="prefixedStyleName">
      <xsl:call-template name="GetPrefixedStyleName">
        <xsl:with-param name="styleName" select="parent::draw:frame/@draw:style-name"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:if test="$prefixedStyleName!=''">
      <w:rStyle w:val="{$prefixedStyleName}"/>
    </xsl:if>
  </xsl:template>

  <!-- Insert Picture's Properties in frame, if needed -->
  <xsl:template name="InsertPicturePropertiesInFrame">
    <xsl:variable name="hPos">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="key('styles', draw:frame/@draw:style-name)"/>
        <xsl:with-param name="attribName">style:horizontal-pos</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$hPos = 'center' ">
        <w:jc>
          <xsl:attribute name="w:val">
            <xsl:text>center</xsl:text>
          </xsl:attribute>
        </w:jc>
      </xsl:when>
      <xsl:when test="$hPos = 'right' ">
        <w:jc>
          <xsl:attribute name="w:val">
            <xsl:text>right</xsl:text>
          </xsl:attribute>
        </w:jc>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!--
  Summary: This template writes the properties of a ODF frame into a OOX frame
  Author: makz (DIaLOGIKa)
  Date: 18.10.2007
  -->
  <xsl:template name="InsertFramePropsFromStyle">
    <xsl:param name="styleName"/>
    <xsl:variable name="style" select="document('content.xml')/office:document-content/office:automatic-styles/style:style[@style:name = $styleName]"/>

    <xsl:attribute name="w:xAlign">
      <xsl:variable name="hpos" select="$style/style:graphic-properties/@style:horizontal-pos"/>
      <xsl:choose>
        <xsl:when test="$hpos='from-left'">
          <xsl:text>left</xsl:text>
        </xsl:when>
        <xsl:when test="$hpos='from-inside'">
          <xsl:text>right</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$hpos"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="w:yAlign">
      <xsl:variable name="vpos" select="$style/style:graphic-properties/@style:vertical-pos"/>
      <xsl:choose>
        <xsl:when test="$vpos='middle'">
          <xsl:text>center</xsl:text>
        </xsl:when>
        <xsl:when test="$vpos='from-top'">
          <xsl:text>top</xsl:text>
        </xsl:when>
        <xsl:when test="$vpos='below'">
          <xsl:text>bottom</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$vpos"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="w:hAnchor">
      <xsl:variable name="hrel" select="$style/style:graphic-properties/@style:horizontal-rel"/>
      <xsl:choose>
        <xsl:when test="$hrel='page-content'">
          <xsl:text>page</xsl:text>
        </xsl:when>
        <xsl:when test="$hrel='page-start-margin'">
          <xsl:text>page</xsl:text>
        </xsl:when>
        <xsl:when test="$hrel='page-end-margin'">
          <xsl:text>page</xsl:text>
        </xsl:when>
        <xsl:when test="$hrel='frame'">
          <xsl:text>margin</xsl:text>
        </xsl:when>
        <xsl:when test="$hrel='frame-content'">
          <xsl:text>margin</xsl:text>
        </xsl:when>
        <xsl:when test="$hrel='frame-start-margin'">
          <xsl:text>margin</xsl:text>
        </xsl:when>
        <xsl:when test="$hrel='frame-end-margin'">
          <xsl:text>margin</xsl:text>
        </xsl:when>
        <xsl:when test="$hrel='paragraph'">
          <xsl:text>text</xsl:text>
        </xsl:when>
        <xsl:when test="$hrel='paragraph-content'">
          <xsl:text>text</xsl:text>
        </xsl:when>
        <xsl:when test="$hrel='paragraph-start-margin'">
          <xsl:text>text</xsl:text>
        </xsl:when>
        <xsl:when test="$hrel='paragraph-end-margin'">
          <xsl:text>text</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$hrel"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="w:vAnchor">
      <xsl:variable name="vrel" select="$style/style:graphic-properties/@style:vertical-rel"/>
      <xsl:choose>
        <xsl:when test="$vrel='page-content'">
          <xsl:text>page</xsl:text>
        </xsl:when>
        <xsl:when test="$vrel='frame'">
          <xsl:text>margin</xsl:text>
        </xsl:when>
        <xsl:when test="$vrel='frame-content'">
          <xsl:text>margin</xsl:text>
        </xsl:when>
        <xsl:when test="$vrel='paragraph'">
          <xsl:text>margin</xsl:text>
        </xsl:when>
        <xsl:when test="$vrel='paragraph-content'">
          <xsl:text>margin</xsl:text>
        </xsl:when>
        <xsl:when test="$vrel='char'">
          <xsl:text>text</xsl:text>
        </xsl:when>
        <xsl:when test="$vrel='line'">
          <xsl:text>text</xsl:text>
        </xsl:when>
        <xsl:when test="$vrel='baseline'">
          <xsl:text>text</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$vrel"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="w:wrap">
      <xsl:variable name="wrap" select="$style/style:graphic-properties/@style:wrap"/>
      <xsl:choose>
        <xsl:when test="$wrap='left'">
          <xsl:text>around</xsl:text>
        </xsl:when>
        <xsl:when test="$wrap='right'">
          <xsl:text>around</xsl:text>
        </xsl:when>
        <xsl:when test="$wrap='parallel'">
          <xsl:text>around</xsl:text>
        </xsl:when>
        <xsl:when test="$wrap='dynamic'">
          <xsl:text>auto</xsl:text>
        </xsl:when>
        <xsl:when test="$wrap='biggest'">
          <xsl:text>around</xsl:text>
        </xsl:when>
        <xsl:when test="$wrap='run-through'">
          <xsl:text>through</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$wrap"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>
</xsl:stylesheet>

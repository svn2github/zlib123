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
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  exclude-result-prefixes="xlink draw svg fo office style text">


  <xsl:key name="images" match="draw:frame" use="'const'"/>

  <!-- check if image type is supported in word  -->
  <xsl:template name="image-support">
    <xsl:param name="name"/>
    <xsl:variable name="support">
      <xsl:choose>
        <xsl:when test="contains($name, '.svm')">
          <xsl:message terminate="no">feedback:SVM image</xsl:message> false </xsl:when>
        <xsl:otherwise>true</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:value-of select="$support"/>
  </xsl:template>


  <!-- Get the position of an element in the draw:frame group -->
  <xsl:template name="GetPosition">
    <xsl:param name="node"/>
    <xsl:variable name="positionInGroup">
      <xsl:for-each select="document('content.xml')">
        <xsl:for-each select="key('images', 'const')">
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
            <xsl:value-of select="count(key('images', 'const'))"/>
          </xsl:for-each>
        </xsl:variable>
        <xsl:for-each select="document('styles.xml')">
          <xsl:for-each select="key('images', 'const')">
            <xsl:if test="generate-id($node) = generate-id(.)">
              <xsl:value-of select="$countContentImages+position()"/>
            </xsl:if>
          </xsl:for-each>
        </xsl:for-each>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--draw:object-ole element represents objects that only have a binary representation -->
  <xsl:template match="draw:frame[./draw:object-ole]" mode="paragraph">
    <xsl:variable name="imageId">
      <xsl:call-template name="GetPosition">
        <xsl:with-param name="node" select="."/>
      </xsl:call-template>

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

  <!-- conversion of external images -->
  <xsl:template
    match="draw:frame[not(./draw:object-ole or ./draw:object) and ./draw:image[not(starts-with(@xlink:href, 'Pictures/'))]]"
    mode="paragraph">
    <xsl:variable name="supported">
      <xsl:call-template name="image-support">
        <xsl:with-param name="name" select="./draw:image/@xlink:href"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$supported = 'true'">
      <xsl:variable name="imageId">
        <xsl:call-template name="GetPosition">
          <xsl:with-param name="node" select="."/>
        </xsl:call-template>
      </xsl:variable>

      <xsl:for-each select="draw:image">
        <w:r>
          <w:pict>
            <v:shape id="{$imageId}" type="#_x0000_t75">

              <xsl:variable name="styleName" select=" parent::draw:frame/@draw:style-name" />
              <xsl:variable name="automaticStyle" select="key('automatic-styles', parent::draw:frame/@draw:style-name)" />
              <xsl:variable name="officeStyle" select="document('styles.xml')//office:document-styles/office:styles/style:style[@style:name = $styleName]"/>
             
              <xsl:variable name="shapeStyle" select="$automaticStyle | $officeStyle" />
              
            <!-- shape properties: size, z-index, coordinates, position, margin etc -->
              <xsl:call-template name="InsertShapeProperties">
                <xsl:with-param name="shapeStyle" select="$shapeStyle" />
              </xsl:call-template>
            
              <v:imagedata r:id="{generate-id(.)}" o:title=""/>
              
              <!-- wrapping -->
              <xsl:call-template name="InsertShapeWrap">
                <xsl:with-param name="shapeStyle" select="$shapeStyle" />
              </xsl:call-template>

            </v:shape>
          </w:pict>
        </w:r>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>

  <!-- frame containing an image -->
  <xsl:template
    match="draw:frame[not(./draw:object-ole or ./draw:object) and starts-with(./draw:image/@xlink:href, 'Pictures/')]"
    mode="paragraph">
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

  <!-- frame containing an image -->
  <xsl:template
    match="draw:frame[not(./draw:object-ole or ./draw:object) and starts-with(./draw:image/@xlink:href, 'Pictures/')]">

    <xsl:variable name="supported">
      <xsl:call-template name="image-support">
        <xsl:with-param name="name" select="./draw:image/@xlink:href"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$supported = 'true'">
      <w:p>
        <w:r>
          <xsl:call-template name="InsertImage"/>
        </w:r>
      </w:p>
    </xsl:if>
  </xsl:template>

  <!--   word has two types of images: inline (positioned with text) and anchor (can be aligned relative to page, margin etc); -->
  <xsl:template name="InsertImage">

    <w:drawing>
      <xsl:variable name="imageId">
        <xsl:call-template name="GetPosition">
          <xsl:with-param name="node" select="."/>
        </xsl:call-template>
      </xsl:variable>

      <xsl:variable name="cx">
        <xsl:choose>
          <xsl:when test="@svg:width">
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="@svg:width"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="draw:text-box/@fo:min-width"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>

      <xsl:variable name="cy">
        <xsl:choose>
          <xsl:when test="@svg:height">
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="@svg:height"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="draw:text-box/@fo:min-height"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>

      <xsl:choose>
        <!-- image embedded in draw:frame/draw:text-box or in text:note element has to be inline with text -->
        <xsl:when
          test="ancestor::draw:text-box or @text:anchor-type='as-char' or ancestor::text:note[@ text:note-class='endnote']">
          <xsl:call-template name="InsertInlineImage">
            <xsl:with-param name="cx" select="$cx"/>
            <xsl:with-param name="cy" select="$cy"/>
            <xsl:with-param name="imageId" select="$imageId"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="InsertAnchorImage">
            <xsl:with-param name="cx" select="$cx"/>
            <xsl:with-param name="cy" select="$cy"/>
            <xsl:with-param name="imageId" select="$imageId"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </w:drawing>
  </xsl:template>

  <!--image positioned with text -->
  <xsl:template name="InsertInlineImage">

    <xsl:param name="cx"/>
    <xsl:param name="cy"/>
    <xsl:param name="imageId"/>

    <wp:inline>

      <!--image height and width-->
      <wp:extent cx="{$cx}" cy="{$cy}"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>

      <!-- image graphical properties: borders, fill, ratio blocking etc -->
      <xsl:call-template name="InsertImageGraphicProperties">
        <xsl:with-param name="cx" select="$cx"/>
        <xsl:with-param name="cy" select="$cy"/>
        <xsl:with-param name="imageId" select="$imageId"/>
      </xsl:call-template>
    </wp:inline>

  </xsl:template>

  <!-- anchor images can be aligned relative to page, margin etc -->
  <xsl:template name="InsertAnchorImage">

    <xsl:param name="cx"/>
    <xsl:param name="cy"/>
    <xsl:param name="imageId"/>

    <xsl:variable name="style"
      select="key('automatic-styles', @draw:style-name)/style:graphic-properties"/>
    <xsl:variable name="styleP" select="key('automatic-styles', parent::text:p/@text:style-name)"/>

    <xsl:variable name="wrap">
      <xsl:choose>
        <xsl:when test="not($style/@style:wrap)">none</xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$style/@style:wrap"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <wp:anchor simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">

      <!-- image z-index-->
      <xsl:attribute name="relativeHeight">
        <xsl:value-of select="2 + @draw:z-index"/>
      </xsl:attribute>

      <xsl:attribute name="behindDoc">
        <xsl:choose>
          <xsl:when test="$wrap = 'run-through' and $style/@style:run-through = 'background' ">1</xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>

      <!-- position -->
      <xsl:call-template name="InsertAnchorImagePosition">
        <xsl:with-param name="style" select="$style"/>
        <xsl:with-param name="styleP" select="$styleP"/>
      </xsl:call-template>

      <!--height and width -->
      <wp:extent cx="{$cx}" cy="{$cy}"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>

      <!--image wrapping-->
      <xsl:call-template name="InsertAnchorImageWrapping">
        <xsl:with-param name="wrap" select="$wrap"/>
        <xsl:with-param name="style" select="$style"/>
      </xsl:call-template>

      <!-- image graphical properties: borders, fill, ratio blocking etc-->
      <xsl:call-template name="InsertImageGraphicProperties">
        <xsl:with-param name="cx" select="$cx"/>
        <xsl:with-param name="cy" select="$cy"/>
        <xsl:with-param name="imageId" select="$imageId"/>
      </xsl:call-template>
    </wp:anchor>

  </xsl:template>

  <xsl:template name="InsertImageGraphicProperties">
    <xsl:param name="imageId"/>
    <xsl:param name="cx"/>
    <xsl:param name="cy"/>

    <!--drawing non-visual properties-->
    <wp:docPr name="{@draw:name}" id="{$imageId}">

      <!--drawing element onclick hyperlink-->
      <xsl:if test="ancestor::draw:a">
        <a:hlinkClick xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/3/main"
          r:id="{generate-id(ancestor::draw:a)}"/>
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
            <xsl:call-template name="InsertImageBorders"/>
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </xsl:template>

  <!--  image vertical and horizontal position-->
  <xsl:template name="InsertAnchorImagePosition">
    <xsl:param name="style"/>
    <xsl:param name="styleP"/>

    <xsl:variable name="ox">
      <xsl:choose>
        <xsl:when test="@svg:x">
          <xsl:call-template name="emu-measure">
            <xsl:with-param name="length" select="@svg:x"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="0"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="oy">
      <xsl:choose>
        <xsl:when test="@svg:y">
          <xsl:call-template name="emu-measure">
            <xsl:with-param name="length" select="@svg:y"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="0"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:call-template name="InsertAnchorImagePosH">
      <xsl:with-param name="style" select="$style"/>
      <xsl:with-param name="styleP" select="$styleP"/>
      <xsl:with-param name="ox" select="$ox"/>
      <xsl:with-param name="oy" select="$oy"/>
    </xsl:call-template>


    <xsl:call-template name="InsertAnchorImagePosV">
      <xsl:with-param name="style" select="$style"/>
      <xsl:with-param name="ox" select="$ox"/>
      <xsl:with-param name="oy" select="$oy"/>
    </xsl:call-template>


  </xsl:template>

  <!--vertical position-->
  <xsl:template name="InsertAnchorImagePosV">
    <xsl:param name="style"/>
    <xsl:param name="ox"/>
    <xsl:param name="oy"/>

    <xsl:variable name="vertical-pos" select="$style/@style:vertical-pos"/>

    <wp:positionV>
      <xsl:attribute name="relativeFrom">
        <xsl:choose>
          <xsl:when test="$style/@style:vertical-rel='page'">page</xsl:when>
          <xsl:otherwise>paragraph</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>

      <xsl:choose>
        <xsl:when test="$vertical-pos !='from-top' and $vertical-pos !='below' ">
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
    <xsl:param name="style"/>
    <xsl:param name="styleP"/>
    <xsl:param name="ox"/>
    <xsl:param name="oy"/>


    <xsl:variable name="horizontal-pos" select="$style/@style:horizontal-pos"/>

    <wp:simplePos x="0" y="0"/>
    <wp:positionH>
      <xsl:attribute name="relativeFrom">

        <xsl:choose>

          <xsl:when test="$style/@style:horizontal-rel='page'">page</xsl:when>
          <xsl:otherwise>column</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:variable name="AlignH" select="$styleP/style:paragraph-properties/@fo:text-align"/>
      <xsl:choose>
        <xsl:when test="$horizontal-pos !='from-left' and $horizontal-pos !='from-outside' ">
          <wp:align>
            <xsl:value-of select="$horizontal-pos"/>
          </wp:align>
        </xsl:when>
        <xsl:when test="($AlignH = 'left') or ($AlignH = 'center') or ($AlignH = 'right')">
          <wp:align>
            <xsl:value-of select="$AlignH"/>
          </wp:align>
        </xsl:when>
        <xsl:otherwise>
          <wp:posOffset>
            <xsl:message>
              <xsl:value-of select="$ox"/>
            </xsl:message>
            <xsl:value-of select="$ox"/>
          </wp:posOffset>
        </xsl:otherwise>
      </xsl:choose>
    </wp:positionH>
  </xsl:template>


  <!--image wrap type-->
  <xsl:template name="InsertAnchorImageWrapping">
    <xsl:param name="style"/>
    <xsl:param name="wrap"/>

    <xsl:choose>
      <xsl:when test="$wrap = 'parallel' or $wrap ='left' or $wrap = 'right' or $wrap ='dynamic'">
        <xsl:call-template name="InsertSquareWrap">
          <xsl:with-param name="wrap" select="$wrap"/>
          <xsl:with-param name="style" select="$style"/>
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="$wrap = 'run-through'">
        <wp:wrapNone/>
      </xsl:when>

      <xsl:when test="$wrap = 'none' ">
        <xsl:call-template name="InsertTopBottomWrap">
          <xsl:with-param name="wrap" select="$wrap"/>
          <xsl:with-param name="style" select="$style"/>
        </xsl:call-template>
      </xsl:when>

      <xsl:otherwise>
        <wp:wrapNone/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- square wrap type for anchor image -->
  <xsl:template name="InsertSquareWrap">
    <xsl:param name="style"/>
    <xsl:param name="wrap"/>

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
        <xsl:choose>
          <xsl:when test="$style/@fo:margin-bottom">
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="$style/@fo:margin-bottom"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="0"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <!--top distance from text-->
      <xsl:attribute name="distT">
        <xsl:choose>
          <xsl:when test="$style/@fo:margin-top">
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="$style/@fo:margin-top"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="0"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <!--left distance from text-->
      <xsl:attribute name="distL">
        <xsl:choose>
          <xsl:when test="$style/@fo:margin-left">
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="$style/@fo:margin-left"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="0"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <!--right distance from text-->
      <xsl:attribute name="distR">
        <xsl:choose>
          <xsl:when test="$style/@fo:margin-right">
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="$style/@fo:margin-right"/>
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
    <xsl:param name="style"/>
    <xsl:param name="wrap"/>

    <wp:wrapTopAndBottom>
      <!--bottom distance from text-->
      <xsl:attribute name="distB">
        <xsl:choose>
          <xsl:when test="$style/@fo:margin-bottom">
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="$style/@fo:margin-bottom"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="0"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <!--top distance from text-->
      <xsl:attribute name="distT">
        <xsl:choose>
          <xsl:when test="$style/@fo:margin-top">
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="$style/@fo:margin-top"/>
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

    <xsl:variable name="sName" select="@draw:style-name"/>
    <xsl:variable name="style" select="key('automatic-styles', $sName)/style:graphic-properties"/>

    <xsl:if test="$style/@fo:border  and ($style/@fo:border != 'none')">
      <xsl:variable name="strokeColor" select="substring-after($style/@fo:border,'#')"/>

      <xsl:variable name="strokeWeight">
        <xsl:call-template name="emu-measure">
          <xsl:with-param name="length" select="substring-before($style/@fo:border,' ')"/>
        </xsl:call-template>
      </xsl:variable>

      <a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <xsl:attribute name="cmpd">
          <xsl:choose>
            <xsl:when
              test="substring-before(substring-after($style/@fo:border,' ' ),' ' ) != 'solid' ">

              <xsl:if test="$style/@style:border-line-width">

                <xsl:variable name="innerLineWidth">
                  <xsl:call-template name="point-measure">
                    <xsl:with-param name="length"
                      select="substring-before($style/@style:border-line-width,' ' )"/>
                  </xsl:call-template>
                </xsl:variable>

                <xsl:variable name="outerLineWidth">
                  <xsl:call-template name="point-measure">
                    <xsl:with-param name="length"
                      select="substring-after(substring-after($style/@style:border-line-width,' ' ),' ' )"
                    />
                  </xsl:call-template>
                </xsl:variable>

                <xsl:if test="$innerLineWidth = $outerLineWidth">thinThin</xsl:if>
                <xsl:if test="$innerLineWidth > $outerLineWidth">thinThick</xsl:if>
                <xsl:if test="$outerLineWidth > $innerLineWidth  ">thickThin</xsl:if>

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
</xsl:stylesheet>

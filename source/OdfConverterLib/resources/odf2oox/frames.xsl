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

  <!--draw:object-ole element represents objects that only have a binary representation -->
  <xsl:template match="draw:frame[./draw:object-ole]" mode="paragraph">
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
        <xsl:call-template name="GetPosition"/>
      </xsl:variable>

      <xsl:for-each select="draw:image">
        <w:r>
          <w:pict>
            <v:shape id="{$imageId}" type="#_x0000_t75">

              <xsl:variable name="styleName" select=" parent::draw:frame/@draw:style-name"/>
              <xsl:variable name="automaticStyle"
                select="key('automatic-styles', parent::draw:frame/@draw:style-name)"/>
              <xsl:variable name="officeStyle"
                select="document('styles.xml')//office:document-styles/office:styles/style:style[@style:name = $styleName]"/>

              <xsl:variable name="shapeStyle" select="$automaticStyle | $officeStyle"/>

              <!-- shape properties: size, z-index, coordinates, position, margin etc -->
              <xsl:call-template name="InsertShapeProperties">
                <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
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
        <xsl:call-template name="GetPosition"/>
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

      <xsl:variable name="automaticStyle" select="key('automatic-styles', @draw:style-name)"/>
      <xsl:variable name="officeStyle"
        select="document('styles.xml')//office:document-styles/office:styles/style:style[@style:name = @draw:style-name]"/>
      <xsl:variable name="imageStyle" select="$automaticStyle | $officeStyle"/>

      <xsl:choose>
        <!-- image embedded in draw:frame/draw:text-box or in text:note element has to be inline with text -->
        <xsl:when
          test="ancestor::draw:text-box or @text:anchor-type='as-char' or ancestor::text:note[@ text:note-class='endnote']">
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
      <wp:effectExtent l="0" t="0" r="0" b="0"/>

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
      <xsl:value-of select="2 + @draw:z-index"/>
    </xsl:attribute>

    <xsl:attribute name="behindDoc">
      <xsl:choose>
        <xsl:when test="$wrap = 'run-through' and $inBackground = 'background' ">1</xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
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
            <xsl:if
              test="key('automatic-styles', @draw:style-name)/style:graphic-properties/@fo:clip">
              <xsl:message terminate="no">feedback:Cropped image</xsl:message>
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
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </xsl:template>

  <!--  image vertical and horizontal position-->
  <xsl:template name="InsertAnchorImagePosition">
    <xsl:param name="imageStyle"/>

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
          <xsl:when test="$verticalRel='page-content'">margin</xsl:when>
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
        <xsl:call-template name="GetValue">
          <xsl:with-param name="length" select="$firstVal"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="b">
        <xsl:call-template name="GetValue">
          <xsl:with-param name="length" select="$thirdVal"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="r">
        <xsl:call-template name="GetValue">
          <xsl:with-param name="length" select="$secondVal"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="l">
        <xsl:call-template name="GetValue">
          <xsl:with-param name="length" select="$fourthVal"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="height">
        <xsl:call-template name="GetValue">
          <xsl:with-param name="length" select="@svg:height"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="width">
        <xsl:call-template name="GetValue">
          <xsl:with-param name="length" select="@svg:width"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:attribute name="t">
        <xsl:value-of select=" round(($t * 100 div ($t + $b + $height)) * 1000)"/>
      </xsl:attribute>
      <xsl:attribute name="b">
        <xsl:value-of select=" round(($b * 100 div ($t + $b + $height)) * 1000)"/>
      </xsl:attribute>
      <xsl:attribute name="r">
        <xsl:value-of select=" round(($r * 100 div ($r + $l + $width)) * 1000)"/>
      </xsl:attribute>
      <xsl:attribute name="l">
        <xsl:value-of select=" round(($l * 100 div ($l + $r + $width)) * 1000)"/>
      </xsl:attribute>
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

    <xsl:choose>
      <xsl:when test="$wrap = 'parallel' or $wrap ='left' or $wrap = 'right' or $wrap ='dynamic'">
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
        <wp:wrapNone/>
      </xsl:otherwise>
    </xsl:choose>
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
        <xsl:call-template name="emu-measure">
          <xsl:with-param name="length" select="substring-before($border,' ')"/>
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

  <!--custom shapes -->

  <xsl:template match="draw:custom-shape">
    <w:p>
    <xsl:call-template name="InsertShapes">
      <xsl:with-param name="shapeType">
        <xsl:value-of select="draw:enhanced-geometry/@draw:type"/>
      </xsl:with-param>
    </xsl:call-template>
    </w:p>
  </xsl:template>

  <xsl:template match="draw:custom-shape" mode="shapes">
    <xsl:call-template name="InsertShapes">
      <xsl:with-param name="shapeType">
        <xsl:value-of select="draw:enhanced-geometry/@draw:type"/>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <!-- basic shapes - ellipse and rect -->

  <xsl:template match="draw:rect|draw:ellipse" mode="shapes">
    <xsl:call-template name="InsertShapes">
      <xsl:with-param name="shapeType" select="name()"/>
    </xsl:call-template>
  </xsl:template>

  <!-- shapes template -->

  <xsl:template name="InsertShapes">
    <xsl:param name="shapeType"/>
    <xsl:variable name="styleName" select=" @draw:style-name"/>
    <xsl:variable name="automaticStyle" select="key('automatic-styles', $styleName)"/>
    <xsl:variable name="officeStyle"
      select="document('styles.xml')//office:document-styles/office:styles/style:style[@style:name = $styleName]"/>
    <xsl:variable name="shapeStyle" select="$automaticStyle | $officeStyle"/>

    <w:r>
      <w:pict>
        <xsl:choose>

          <!-- rectangle -->
          <xsl:when test="$shapeType = 'draw:rect' or $shapeType = 'rectangle' ">
            <v:rect>
              <xsl:call-template name="SimpleShape">
                <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
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
              </xsl:call-template>
            </v:oval>
          </xsl:when>

          <!-- round-rectangle -->

          <xsl:when test="$shapeType = 'round-rectangle' ">
            <v:roundrect>
              <xsl:call-template name="SimpleShape">
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
    </xsl:template> -->

  <!-- simple shape properties -->
  <xsl:template name="SimpleShape">
    <xsl:param name="shapeStyle"/>

    <xsl:if test="$shapeStyle != 0 or count($shapeStyle) &gt; 1">
      <!-- shape properties: size, z-index, color, position, stroke, etc -->
      <xsl:call-template name="InsertDrawnShapeProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

      <xsl:call-template name="FillColor">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

      <xsl:call-template name="Stroke">
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
    <xsl:attribute name="style">
      <xsl:call-template name="InsertPosition"/>

      <xsl:call-template name="InsertDrawnShapeSize"/>

      <xsl:call-template name="InsertDrawnShapeZindex"/>

      <xsl:call-template name="InsertShapeCoordinates">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>


  <!-- shape position -->
  <xsl:template name="InsertPosition">
    <xsl:variable name="x">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length" select="@svg:x"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="y">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length" select="@svg:y"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="horizontal">
      <xsl:if
       test="key('automatic-styles', @draw:style-name)/style:graphic-properties/@style:horizontal-pos='from-left'">
        <xsl:text>mso-position-horizontal-relative:left-margin-area</xsl:text>
      </xsl:if>
    </xsl:variable>
    <xsl:variable name="Vertical">
      <xsl:if
       test="key('automatic-styles', @draw:style-name)/style:graphic-properties/@style:vertical-pos='from-top'">
        <xsl:text>mso-position-vertical-relative:top-margin-area</xsl:text>
      </xsl:if>
    </xsl:variable>
    <xsl:value-of
     select="concat('position:absolute;margin-left:',$x,'pt;margin-top:',$y,'pt;',$horizontal,';',$Vertical,';')"
    />
  </xsl:template>

  <!--insert shahpe z-index -->
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
    <xsl:variable name="relWidth" select="substring-before(@style:rel-width,'%')"/>
    <xsl:variable name="relHeight" select="substring-before(@style:rel-height,'%')"/>
    <xsl:value-of select="concat('width:',$frameW,'pt;')"/>
    <xsl:value-of select="concat('height:',$frameH,'pt;')"/>
  </xsl:template>


  <!-- computes text-box margin -->
  <xsl:template name="ComputeMarginX">
    <xsl:param name="parent"/>
    <xsl:choose>
      <xsl:when test="$parent">
        <xsl:variable name="recursive_result">
          <xsl:call-template name="point-measure">
            <xsl:with-param name="length">
              <xsl:call-template name="ComputeMarginX">
                <xsl:with-param name="parent" select="$parent[position()>1]"/>
              </xsl:call-template>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="fo_margins_left">
          <xsl:value-of
            select="key('Index', parent::draw:frame/@draw:style-name)/style:graphic-properties/@fo:margin-left"
          />
        </xsl:variable>
        <xsl:variable name="fo_page_width">
          <xsl:call-template name="GetValue">
            <xsl:with-param name="length">
              <xsl:value-of
                select="document('styles.xml')/office:document-styles/office:automatic-styles/style:page-layout/style:page-layout-properties/@fo:page-width"
              />
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="fo_margin_left">
          <xsl:call-template name="GetValue">
            <xsl:with-param name="length">
              <xsl:value-of
                select="document('styles.xml')/office:document-styles/office:automatic-styles/style:page-layout/style:page-layout-properties/@fo:margin-left"
              />
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="fo_margin_right">
          <xsl:call-template name="GetValue">
            <xsl:with-param name="length">
              <xsl:value-of
                select="document('styles.xml')/office:document-styles/office:automatic-styles/style:page-layout/style:page-layout-properties/@fo:margin-right"
              />
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="fo_margin_right_frame">
          <xsl:call-template name="GetValue">
            <xsl:with-param name="length">
              <xsl:value-of
                select="key('Index', parent::draw:frame/@draw:style-name)/style:graphic-properties/@fo:margin-right"
              />
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="svg_width">
          <xsl:call-template name="GetValue">
            <xsl:with-param name="length">
              <xsl:value-of select="parent::draw:frame/@svg:width"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="style_horizontal_pos">
          <xsl:value-of
            select="key('Index', parent::draw:frame/@draw:style-name)/style:graphic-properties/@style:horizontal-pos"
          />
        </xsl:variable>
        <xsl:variable name="style_horizontal_rel">
          <xsl:value-of
            select="key('Index', parent::draw:frame/@draw:style-name)/style:graphic-properties/@style:horizontal-rel"
          />
        </xsl:variable>
        <xsl:variable name="svgx">
          <xsl:choose>
            <xsl:when
              test="$parent[1]/@svg:x or ($style_horizontal_pos='left' and not($style_horizontal_rel='page-end-margin')) or ($style_horizontal_pos='right' and not($style_horizontal_rel='page-start-margin'))">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:choose>
                    <xsl:when test="$style_horizontal_pos='left'">
                      <xsl:value-of select="$fo_margins_left"/>
                    </xsl:when>
                    <xsl:when
                      test="($style_horizontal_pos='right' ) and ($style_horizontal_rel='page') and not($style_horizontal_rel='page-start-margin') ">
                      <xsl:value-of
                        select="concat($fo_page_width - $fo_margin_left + $fo_margin_right - $svg_width - $fo_margin_right_frame, 'cm')"
                      />
                    </xsl:when>
                    <xsl:when
                      test="($style_horizontal_pos='right' ) and ($style_horizontal_rel='page-end-margin' )  and not($style_horizontal_rel='page-start-margin') ">
                      <xsl:value-of
                        select="concat( - $fo_margin_right_frame + $fo_margin_right - $svg_width, 'cm')"
                      />
                    </xsl:when>
                    <xsl:when test="$style_horizontal_pos='right' ">
                      <xsl:value-of
                        select="concat($fo_page_width - $fo_margin_left - $fo_margin_right - $svg_width - $fo_margin_right_frame, 'cm')"
                      />
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$parent[1]/@svg:x"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:value-of select="$svgx+$recursive_result"/>
      </xsl:when>
      <xsl:otherwise>0</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- computes text-box margin -->
  <xsl:template name="ComputeMarginY">
    <xsl:param name="parent"/>
    <xsl:choose>
      <xsl:when test="$parent">
        <xsl:variable name="recursive_result">
          <xsl:call-template name="point-measure">
            <xsl:with-param name="length">
              <xsl:call-template name="ComputeMarginY">
                <xsl:with-param name="parent" select="$parent[position()>1]"/>
              </xsl:call-template>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="svgy">
          <xsl:choose>
            <xsl:when test="$parent[1]/@svg:y">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:value-of select="$parent[1]/@svg:y"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:value-of select="$svgy+$recursive_result"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="0"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- text boxes -->
  <xsl:template match="draw:text-box" mode="paragraph">
    <w:r>
      <w:rPr>
        <xsl:if test="ancestor::text:section/@text:display='none'">
          <w:vanish/>
        </xsl:if>
        <xsl:call-template name="InsertTextBoxStyle"/>
      </w:rPr>
      <w:pict>

        <!--this properties are needed to make z-index work properly-->
        <v:shapetype coordsize="21600,21600" path="m,l,21600r21600,l21600,xe"
          xmlns:o="urn:schemas-microsoft-com:office:office">
          <v:stroke joinstyle="miter"/>
          <v:path gradientshapeok="t" o:connecttype="rect"/>
        </v:shapetype>

        <v:shape type="#_x0000_t202">
          <xsl:if test="ancestor::draw:a/@xlink:href">
            <xsl:attribute name="href">
              <xsl:choose>
                <xsl:when test="substring-before(ancestor::draw:a/@xlink:href,'/')='..'">
                  <xsl:value-of select="substring-after(ancestor::draw:a/@xlink:href,'/')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="ancestor::draw:a/@xlink:href"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
          <xsl:variable name="styleName" select=" parent::draw:frame/@draw:style-name"/>
          <xsl:variable name="automaticStyle" select="key('automatic-styles', $styleName)"/>
          <xsl:variable name="officeStyle"
            select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $styleName]"/>

          <xsl:variable name="shapeStyle" select="$automaticStyle | $officeStyle"/>

          <xsl:call-template name="InsertShapeProperties">
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
          </xsl:call-template>

          <!--insert text-box-->
          <xsl:call-template name="InsertTextBox">
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
          </xsl:call-template>

        </v:shape>
      </w:pict>
    </w:r>
  </xsl:template>

  <xsl:template name="InsertShapePositionProperties">
    <xsl:param name="shapeStyle"/>

    <xsl:variable name="frameWrap">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:wrap</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!-- absolute positioning for text-box -->
    <xsl:if test="not(parent::draw:frame/@text:anchor-type = 'as-char') ">
      <xsl:value-of select="'position:absolute;'"/>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertShapeSize">

    <!--  text-box width and height-->
    <xsl:variable name="frameW">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length" select="parent::draw:frame/@svg:width"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="frameH">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length" select="parent::draw:frame/@svg:height"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="relWidth" select="substring-before(parent::draw:frame/@style:rel-width,'%')"/>
    <xsl:variable name="relHeight"
      select="substring-before(parent::draw:frame/@style:rel-height,'%')"/>

    <xsl:if test="not(parent::draw:frame/@fo:min-width)">
      <xsl:value-of select="concat('width:',$frameW,'pt;')"/>
    </xsl:if>
    <xsl:if test="not(@fo:min-height)">
      <xsl:value-of select="concat('height:',$frameH,'pt;')"/>
    </xsl:if>

    <!--relative width and hight-->
    <xsl:if test="$relWidth">
      <xsl:value-of select="concat('mso-width-percent:',$relWidth,'0;')"/>
    </xsl:if>
    <xsl:if test="$relHeight">
      <xsl:value-of select="concat('mso-height-percent:',$relHeight,'0;')"/>
    </xsl:if>

    <xsl:variable name="relativeTo">
      <xsl:choose>
        <xsl:when test="parent::draw:frame/@text:anchor-type = 'page'">
          <xsl:text>page</xsl:text>
        </xsl:when>
        <xsl:when test="parent::draw:frame/@text:anchor-type = 'paragraph'">
          <xsl:text>margin</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:text>margin</xsl:text>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:value-of select="concat('mso-width-relative:',$relativeTo,';')"/>
    <xsl:value-of select="concat('mso-height-relative:',$relativeTo,';')"/>

  </xsl:template>

  <xsl:template name="InsertShapeZindex">
    <xsl:param name="shapeStyle"/>

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
            <xsl:when test="parent::draw:frame/@draw:z-index">
              <xsl:value-of select="2 + parent::draw:frame/@draw:z-index"/>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:value-of select="concat('z-index:', $zIndex, ';')"/>
  </xsl:template>

  <xsl:template name="InsertShapeCoordinates">
    <xsl:param name="shapeStyle"/>
    <xsl:variable name="graphicProps" select="$shapeStyle/style:graphic-properties"/>

    <!-- text-box coordinates -->
    <xsl:variable name="posL">
      <xsl:if
        test="parent::draw:frame/@svg:x or (($graphicProps/@style:horizontal-pos='left' or ($graphicProps/@style:horizontal-pos='right'  and not($graphicProps/@style:horizontal-rel='page-start-margin'))) and ($graphicProps/@fo:margin-left or $graphicProps/@fo:margin-right))">
        <xsl:variable name="leftM">
          <xsl:call-template name="ComputeMarginX">
            <xsl:with-param name="parent" select="ancestor::draw:frame"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:value-of select="$leftM"/>
      </xsl:if>
    </xsl:variable>
    <xsl:variable name="posT">
      <xsl:if test="parent::draw:frame/@svg:y">
        <xsl:variable name="topM">
          <xsl:call-template name="ComputeMarginY">
            <xsl:with-param name="parent" select="ancestor::draw:frame"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:value-of select="$topM"/>
      </xsl:if>
    </xsl:variable>

    <xsl:if
      test="parent::draw:frame/@svg:x or (($graphicProps/@style:horizontal-pos='left' or ($graphicProps/@style:horizontal-pos='right'  and not($graphicProps/@style:horizontal-rel='page-start-margin'))) and ($graphicProps/@fo:margin-left or $graphicProps/@fo:margin-right)) ">
      <xsl:value-of select="concat('margin-left:',$posL,'pt;')"/>
    </xsl:if>
    <xsl:if test="parent::draw:frame/@svg:y">
      <xsl:value-of select="concat('margin-top:',$posT,'pt;')"/>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertShapePositionRelative">
    <xsl:param name="shapeStyle"/>

    <xsl:variable name="horizontalRel">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:horizontal-rel</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="verticalRel">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:vertical-rel</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!-- text-box relative position -->
    <xsl:choose>
      <xsl:when test="$horizontalRel = 'page-end-margin' ">mso-position-horizontal-relative:
        right-margin-area;</xsl:when>
      <xsl:when test="$horizontalRel = 'page-start-margin' ">mso-position-horizontal-relative:
        left-margin-area;</xsl:when>
      <xsl:when test="$horizontalRel = 'page' ">mso-position-horizontal-relative: page;</xsl:when>
      <xsl:when test="$horizontalRel = 'page-content' ">mso-position-horizontal-relative: column;</xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="parent::draw:frame/@text:anchor-type = 'page' "
            >mso-position-horizontal-relative: page;</xsl:when>
          <xsl:when test="parent::draw:frame/@text:anchor-type = 'paragraph' "
            >mso-position-horizontal-relative: column;</xsl:when>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>

    <xsl:choose>
      <xsl:when test="$verticalRel = 'page' ">mso-position-vertical-relative: page;</xsl:when>
      <xsl:when test="$verticalRel = 'page-content' ">mso-position-vertical-relative:
      margin;</xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertShapePosition">
    <xsl:param name="shapeStyle"/>
    <xsl:variable name="graphicProps" select="$shapeStyle/style:graphic-properties"/>

    <xsl:variable name="horizontalPos">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:horizontal-pos</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="verticalPos">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:vertical-pos</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="fo_margin_left">
      <xsl:call-template name="GetValue">
        <xsl:with-param name="length">
          <xsl:value-of select="$graphicProps/@fo:margin-left"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="fo_margin_right">
      <xsl:call-template name="GetValue">
        <xsl:with-param name="length">
          <xsl:value-of select="$graphicProps/@fo:margin-right"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!--horizontal position-->
    <!-- The same style defined in styles.xsl  TODO manage horizontal-rel-->
    <xsl:if test="$horizontalPos">
      <xsl:choose>
        <xsl:when test="$horizontalPos = 'center'">
          <xsl:value-of select="concat('mso-position-horizontal:', 'center',';')"/>
        </xsl:when>
        <xsl:when
          test="($horizontalPos='left' and ($fo_margin_left='' or not($fo_margin_left != 0))) or ($graphicProps/@style:horizontal-pos='right' and $graphicProps/@style:horizontal-rel='page-start-margin')">
          <xsl:value-of select="concat('mso-position-horizontal:', 'left',';')"/>
        </xsl:when>
        <xsl:when
          test="($horizontalPos='right' and ($fo_margin_right='' or not($fo_margin_right != 0)) and not($graphicProps/@style:horizontal-rel='page-start-margin')) or ($graphicProps/@style:horizontal-pos='left' and ($graphicProps/@style:horizontal-rel='page-end-margin'))">
          <xsl:value-of select="concat('mso-position-horizontal:', 'right',';')"/>
        </xsl:when>
        <!-- <xsl:otherwise><xsl:value-of select="concat('mso-position-horizontal:', 'center',';')"/></xsl:otherwise> -->
      </xsl:choose>
    </xsl:if>

    <!-- vertical position-->
    <xsl:if test="$verticalPos">
      <xsl:choose>
        <xsl:when test="$verticalPos = 'middle'">
          <xsl:value-of select="concat('mso-position-vertical:', 'center',';')"/>
        </xsl:when>
        <xsl:when test="$verticalPos='top'">
          <xsl:value-of select="concat('mso-position-vertical:', 'top',';')"/>
        </xsl:when>
        <xsl:when test="$verticalPos='bottom'">
          <xsl:value-of select="concat('mso-position-vertical:', 'bottom',';')"/>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertShapeMargin">
    <xsl:param name="shapeStyle"/>

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
    <xsl:if test="contains(parent::draw:frame/@draw:transform,'rotate')">
      <xsl:text>rotation:</xsl:text>
      <xsl:variable name="angle">
        <xsl:call-template name="DegreesAngle">
          <xsl:with-param name="angle">
            <xsl:value-of
              select="substring-before(substring-after(substring-after(parent::draw:frame/@draw:transform,'rotate'),'('),')')"
            />
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:value-of select="$angle"/>
      <xsl:text>;</xsl:text>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertShapeTransparency">
    <xsl:param name="shapeStyle"/>

    <xsl:variable name="backgroundColor">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">fo:background-color</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="backgroundTransparency">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:background-transparency</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!--fill  transparency-->
    <xsl:variable name="opacity" select="100 - substring-before($backgroundTransparency,'%')"/>

    <xsl:if test="$backgroundTransparency != '' and $backgroundColor != 'transparent' ">
      <v:fill>
        <xsl:attribute name="opacity">
          <xsl:value-of select="concat($opacity,'%')"/>
        </xsl:attribute>
      </v:fill>
    </xsl:if>
  </xsl:template>

  <!--converts oo frame style properties to shape properties for text-box-->
  <xsl:template name="InsertShapeProperties">
    <xsl:param name="shapeStyle"/>

    <xsl:attribute name="style">

      <xsl:call-template name="InsertShapePositionProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

      <xsl:call-template name="InsertShapeSize"/>

      <xsl:call-template name="InsertShapeZindex">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

      <xsl:call-template name="InsertShapeCoordinates">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

      <xsl:call-template name="InsertShapePositionRelative">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

      <xsl:call-template name="InsertShapePosition">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

      <xsl:call-template name="InsertShapeMargin">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

      <xsl:call-template name="InsertShapeRotation"/>

    </xsl:attribute>

    <xsl:call-template name="InsertShapeFill">
      <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
    </xsl:call-template>

    <xsl:call-template name="InsertShapeBorders">
      <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
    </xsl:call-template>

    <xsl:call-template name="InsertShapeTransparency">
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
          select="document('styles.xml')//office:document-styles/office:styles/style:style[@style:name = $parentStyleName]"/>

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
      <xsl:when test="parent::draw:frame/@text:anchor-type = 'as-char' ">
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
        <xsl:message terminate="no">feedback:Shape top-and-bottom wrapping</xsl:message>
        <w10:wrap type="topAndBottom"/>
      </xsl:when>
      <xsl:when test="$frameWrap = 'parallel' ">
        <w10:wrap type="square"/>
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
      <xsl:if
        test="parent::draw:frame/@fo:min-width or @fo:min-height or contains(parent::draw:frame/@draw:transform,'rotate')">
        <xsl:attribute name="style">
          <xsl:if test="parent::draw:frame/@fo:min-width or @fo:min-height">
            <xsl:value-of select="'mso-fit-shape-to-text:t'"/>
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
              <xsl:when test="$angle = 90 or $angle = -90">
                <xsl:text>layout-flow:vertical;mso-layout-flow-alt:bottom-to-top;</xsl:text>
              </xsl:when>
              <xsl:otherwise>
                <xsl:text>mso-rotate:</xsl:text>
                <xsl:value-of select="$angle"/>
                <xsl:text>;</xsl:text>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:if>
        </xsl:attribute>
      </xsl:if>

      <xsl:call-template name="InsertTextBoxInset">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

      <w:txbxContent>
        <xsl:for-each select="child::node()">

          <xsl:choose>
            <!--   ignore embedded text-box becouse word doesn't support it-->
            <xsl:when test="self::node()[name(draw:text-box)]">
              <xsl:message terminate="no">feedback: Nested frames</xsl:message>
            </xsl:when>

            <!--default scenario-->
            <xsl:otherwise>
              <xsl:apply-templates select="."/>
            </xsl:otherwise>
          </xsl:choose>

        </xsl:for-each>
      </w:txbxContent>

      <xsl:call-template name="InsertShapeWrap">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

    </v:textbox>
  </xsl:template>

  <!--converts oo frame padding into inset for text-box -->
  <xsl:template name="InsertTextBoxInset">
    <xsl:param name="shapeStyle"/>

    <xsl:variable name="padding">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">fo:padding</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="paddingTop">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">fo:padding-top</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="paddingRight">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">fo:padding-right</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="paddingLeft">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">fo:padding-left</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="paddingBottom">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">fo:padding-bottom</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:attribute name="inset">
      <xsl:choose>
        <xsl:when test="$padding != '' or $paddingTop != '' ">
          <xsl:call-template name="CalculateTextBoxPadding">
            <xsl:with-param name="padding" select="$padding"/>
            <xsl:with-param name="paddingTop" select="$paddingTop"/>
            <xsl:with-param name="paddingLeft" select="$paddingLeft"/>
            <xsl:with-param name="paddingRight" select="$paddingRight"/>
            <xsl:with-param name="paddingBottom" select="$paddingBottom"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="CalculateTextBoxPadding"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>


  <!--calculates textbox inset attribute  -->
  <xsl:template name="CalculateTextBoxPadding">
    <xsl:param name="padding"/>
    <xsl:param name="paddingBottom"/>
    <xsl:param name="paddingLeft"/>
    <xsl:param name="paddingRight"/>
    <xsl:param name="paddingTop"/>


    <xsl:choose>
      <xsl:when test="not($padding)">0mm,0mm,0mm,0mm</xsl:when>
      <xsl:when test="$padding">
        <xsl:variable name="paddingVal">
          <xsl:call-template name="milimeter-measure">
            <xsl:with-param name="length" select="$padding"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:value-of
          select="concat($paddingVal,'mm,',$paddingVal,'mm,',$paddingVal,'mm,',$paddingVal,'mm')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="padding-top">
          <xsl:if test="$paddingTop">
            <xsl:call-template name="milimeter-measure">
              <xsl:with-param name="length" select="$paddingTop"/>
            </xsl:call-template>
          </xsl:if>
        </xsl:variable>
        <xsl:variable name="padding-right">
          <xsl:if test="$paddingRight">
            <xsl:call-template name="milimeter-measure">
              <xsl:with-param name="length" select="$paddingRight"/>
            </xsl:call-template>
          </xsl:if>
        </xsl:variable>
        <xsl:variable name="padding-bottom">
          <xsl:if test="$paddingBottom">
            <xsl:call-template name="milimeter-measure">
              <xsl:with-param name="length" select="$paddingBottom"/>
            </xsl:call-template>
          </xsl:if>
        </xsl:variable>
        <xsl:variable name="padding-left">
          <xsl:if test="$paddingLeft">
            <xsl:call-template name="milimeter-measure">
              <xsl:with-param name="length" select="$paddingLeft"/>
            </xsl:call-template>
          </xsl:if>
        </xsl:variable>
        <xsl:if test="$paddingTop">
          <xsl:value-of
            select="concat($padding-left,'mm,',$padding-top,'mm,',$padding-right,'mm,',$padding-bottom,'mm')"
          />
        </xsl:if>
      </xsl:otherwise>
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

  <xsl:template match="draw:frame" mode="paragraph">
    <xsl:call-template name="InsertEmbeddedTextboxes"/>
  </xsl:template>

  <xsl:template match="draw:frame">
    <xsl:choose>
      <xsl:when test=" preceding-sibling::node()[1][name() != 'draw:frame']">
        <w:p>
          <xsl:call-template name="InsertEmbeddedTextboxes"/>
          <!-- put all  consecutive frames in same paragraph -->
          <xsl:call-template name="InsertFollowingFrame"/>
        </w:p>
      </xsl:when>
      <xsl:otherwise>
        <!-- already handled as a 'followinf frame' -->
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Insert a frame in same paragraph as preceding frame -->
  <xsl:template name="InsertFollowingFrame">
    <xsl:for-each select="following-sibling::node()[1][name()='draw:frame']">
      <xsl:call-template name="InsertEmbeddedTextboxes"/>
      <!-- put all  consecutive frames in same paragraph -->
      <xsl:call-template name="InsertFollowingFrame"/>
    </xsl:for-each>
  </xsl:template>

  <!-- inserts textboxes which are embedded in odf as one after another in word -->
  <xsl:template name="InsertEmbeddedTextboxes">
    <xsl:for-each select="descendant::draw:text-box">
      <xsl:apply-templates mode="paragraph" select="."/>
    </xsl:for-each>
  </xsl:template>


  <!-- Insert Picture's Properties in frame, if needed -->
  <xsl:template name="InsertPicturePropertiesInFrame">
    <xsl:if test="ancestor::draw:text-box and descendant::draw:image">
      <xsl:choose>
        <xsl:when
          test="key('automatic-styles', draw:frame/@draw:style-name)/style:graphic-properties/@style:horizontal-pos='center'">
          <w:jc>
            <xsl:attribute name="w:val">
              <xsl:text>center</xsl:text>
            </xsl:attribute>
          </w:jc>
        </xsl:when>
        <xsl:when
          test="key('automatic-styles', draw:frame/@draw:style-name)/style:graphic-properties/@style:horizontal-pos='right'">
          <w:jc>
            <xsl:attribute name="w:val">
              <xsl:text>right</xsl:text>
            </xsl:attribute>
          </w:jc>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
  </xsl:template>

</xsl:stylesheet>

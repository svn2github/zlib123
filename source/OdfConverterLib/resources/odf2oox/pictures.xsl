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
    <xsl:param name="imageStyle" />

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
    <xsl:param name="imageStyle" />

    <wp:anchor simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">

      <!-- image z-index-->
      <xsl:call-template name="InsertZindex" >
        <xsl:with-param name="imageStyle" select="$imageStyle" />
      </xsl:call-template>
      
      <!-- position -->
      <xsl:call-template name="InsertAnchorImagePosition">
        <xsl:with-param name="imageStyle" select="$imageStyle" />
      </xsl:call-template>

      <!--height and width -->
      <wp:extent cx="{$cx}" cy="{$cy}"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>

      <!--image wrapping-->
      <xsl:call-template name="InsertAnchorImageWrapping">
        <xsl:with-param name="imageStyle" select="$imageStyle" />
      </xsl:call-template>

      <!-- image graphical properties: borders, fill, ratio blocking etc-->
      <xsl:call-template name="InsertImageGraphicProperties">
        <xsl:with-param name="cx" select="$cx"/>
        <xsl:with-param name="cy" select="$cy"/>
        <xsl:with-param name="imageId" select="$imageId"/>
        <xsl:with-param name="imageStyle" select="$imageStyle" />
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
    <xsl:param name="imageStyle" />

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
            <xsl:variable name="styleCrop" select="./@draw:style-name"/>
            <xsl:variable name="clip" select="//office:document-content/office:automatic-styles/style:style[@style:name = $styleCrop]"/>
            <xsl:if test="$clip/style:graphic-properties/@fo:clip">
              <xsl:call-template name="InsertImageCropProperties"/>
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
              <xsl:with-param name="imageStyle" select="$imageStyle" />
            </xsl:call-template>
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </xsl:template>

  <!--  image vertical and horizontal position-->
  <xsl:template name="InsertAnchorImagePosition">
    <xsl:param name="imageStyle" />

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
      <xsl:with-param name="imageStyle" select="$imageStyle" />
      <xsl:with-param name="ox" select="$ox"/>
      <xsl:with-param name="oy" select="$oy"/>
    </xsl:call-template>


    <xsl:call-template name="InsertAnchorImagePosV">
      <xsl:with-param name="imageStyle" select="$imageStyle" />
      <xsl:with-param name="ox" select="$ox"/>
      <xsl:with-param name="oy" select="$oy"/>
    </xsl:call-template>


  </xsl:template>

  <!--vertical position-->
  <xsl:template name="InsertAnchorImagePosV">
    <xsl:param name="ox"/>
    <xsl:param name="oy"/>
    <xsl:param name="imageStyle" />
    
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
          <xsl:when test="contains($verticalRel,'page')">page</xsl:when>
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
        <xsl:when test="$vertical-pos != '' and $vertical-pos !='from-top' and $vertical-pos !='below' ">
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
    <xsl:param name="imageStyle" />

    <xsl:variable name="horizontal-pos" >
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
        <xsl:when test="$horizontal-pos != '' and $horizontal-pos !='from-left' and $horizontal-pos !='from-outside' ">
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
    <a:srcRect>
      <xsl:variable name="styleCrop" select="./@draw:style-name"/>
      <xsl:variable name="clip" select="//office:document-content/office:automatic-styles/style:style[@style:name = $styleCrop]"/>
      <xsl:variable name="t"><xsl:value-of select="substring-before(substring-after($clip/style:graphic-properties/@fo:clip,'('),'cm' )"/></xsl:variable>
      <xsl:variable name="b"><xsl:value-of select="substring-before(substring-after(substring-after($clip/style:graphic-properties/@fo:clip,'cm'),'cm'),'cm')"/></xsl:variable>
      <xsl:variable name="r"><xsl:value-of select="substring-before(substring-after($clip/style:graphic-properties/@fo:clip,'cm'),'cm')"/></xsl:variable>
      <xsl:variable name="l"><xsl:value-of select="substring-before(substring-after(substring-after(substring-after($clip/style:graphic-properties/@fo:clip,'cm'),'cm'),'cm'),'cm')"/></xsl:variable>
      <xsl:variable name="height"><xsl:value-of select="substring-before(./@svg:height,'cm')"/></xsl:variable>
      <xsl:variable name="width"><xsl:value-of select="substring-before(./@svg:width,'cm')"/></xsl:variable>
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
    <xsl:param name="imageStyle" />

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
          <xsl:with-param name="imageStyle" select="$imageStyle" />
        </xsl:call-template>
      </xsl:when>

      <xsl:when test="$wrap = 'run-through'">
        <wp:wrapNone/>
      </xsl:when>

      <xsl:when test="$wrap = 'none' ">
        <xsl:call-template name="InsertTopBottomWrap">
          <xsl:with-param name="wrap" select="$wrap"/>
          <xsl:with-param name="imageStyle" select="$imageStyle" />
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
    <xsl:param name="imageStyle" />

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
    <xsl:param name="imageStyle" />

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
    <xsl:param name="imageStyle" />

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
            <xsl:when
              test="substring-before(substring-after($border,' ' ),' ' ) != 'solid' ">

              <xsl:variable name="borderLineWidth">
                <xsl:call-template name="GetGraphicProperties">
                  <xsl:with-param name="shapeStyle" select="$imageStyle"/>
                  <xsl:with-param name="attribName">style:border-line-width</xsl:with-param>
                </xsl:call-template>
              </xsl:variable>
              
              <xsl:if test="$borderLineWidth != ''">

                <xsl:variable name="innerLineWidth">
                  <xsl:call-template name="point-measure">
                    <xsl:with-param name="length"
                      select="substring-before($borderLineWidth,' ' )"/>
                  </xsl:call-template>
                </xsl:variable>

                <xsl:variable name="outerLineWidth">
                  <xsl:call-template name="point-measure">
                    <xsl:with-param name="length"
                      select="substring-after(substring-after($borderLineWidth,' ' ),' ' )"
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
  
  <!--custom shapes -->
  
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
    <w:r>
      <w:pict>
        <xsl:choose>
          
          <!-- rectangle -->
          <xsl:when test="$shapeType = 'draw:rect' or $shapeType = 'rectangle' ">
            <v:rect>
              <xsl:call-template name="SimpleShape"/>
            </v:rect>
          </xsl:when>
          
          <!-- ellipse -->
          <xsl:when test="$shapeType = 'draw:ellipse' or $shapeType = 'ellipse' ">
            <v:oval>
              <xsl:call-template name="SimpleShape"/>
            </v:oval>
          </xsl:when>
          
          <!-- round-rectangle -->
          
          <xsl:when test="$shapeType = 'round-rectangle' ">
            <v:roundrect>
              <xsl:call-template name="SimpleShape"/>
            </v:roundrect>
          </xsl:when>
          
          <!-- TODO - other shapes -->
          <xsl:otherwise></xsl:otherwise>
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
    <xsl:variable name="styleName" select="@draw:style-name" />
    <xsl:variable name="automaticStyle" select="document('content.xml')//office:document-content/office:automatic-styles/style:style[@style:name = $styleName]"/>
    <xsl:variable name="officeStyle" select="document('styles.xml')//office:document-styles/office:styles/style:style[@style:name = $styleName]"/>
    <xsl:variable name="shapeStyle" select="$automaticStyle | $officeStyle" />
    
    <!-- shape properties: size, z-index, color, position, stroke, etc --> 
    <xsl:call-template name="InsertDrawnShapeProperties">
      <xsl:with-param name="shapeStyle" select="$shapeStyle" />
    </xsl:call-template>
    <xsl:attribute name="fillcolor">
      <xsl:call-template name="FillColor">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:call-template name="Stroke">
      <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
    </xsl:call-template>
  </xsl:template>
  
  <!-- shape fill color -->
  
  <xsl:template name="FillColor">
    <xsl:param name="shapeStyle"/>
    <xsl:call-template name="GetDrawnGraphicProperties">
      <xsl:with-param name="attrib">draw:fill-color</xsl:with-param>
      <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
    </xsl:call-template>
  </xsl:template>
  
  <!-- shape stroke color and stroke width -->
  
  <xsl:template name="Stroke">
    <xsl:param name="shapeStyle"/>
    <xsl:attribute name="strokecolor">
      <xsl:call-template name="GetDrawnGraphicProperties">
        <xsl:with-param name="attrib">svg:stroke-color</xsl:with-param>
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="strokeweight">
      <xsl:variable name="strokeWeight">
        <xsl:call-template name="point-measure">
          <xsl:with-param name="length">
            <xsl:call-template name="GetDrawnGraphicProperties">
              <xsl:with-param name="attrib">svg:stroke-width</xsl:with-param>
              <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
            </xsl:call-template>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:value-of select="concat($strokeWeight,'pt')"/>
    </xsl:attribute>
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
      <xsl:call-template name="InsertShapeCoordinates"/>
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
    <xsl:value-of select="concat('position:absolute;margin-left:',$x,'pt;margin-top:',$y,'pt;')"/>
  </xsl:template>
  <xsl:template name="InsertDrawnShapeZindex">
    <xsl:value-of select="concat('z-index:',@draw:z-index)"/>
  </xsl:template>
  
  <!--  shape width and height-->  
  
  <xsl:template name="InsertDrawnShapeSize">
   <xsl:variable name="frameW">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length"
          select="@svg:width|@fo:min-width"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="frameH">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length" select="@fo:min-height|@svg:height"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="relWidth" select="substring-before(@style:rel-width,'%')"/>
    <xsl:variable name="relHeight"
      select="substring-before(@style:rel-height,'%')"/> 
    <xsl:value-of select="concat('width:',$frameW,'pt;')"/>
    <xsl:value-of select="concat('height:',$frameH,'pt;')"/>
  </xsl:template>
</xsl:stylesheet>

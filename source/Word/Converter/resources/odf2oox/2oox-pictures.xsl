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

  <!--
  *************************************************************************
  SUMMARY
  *************************************************************************
  This stylesheet handles the conversion of images. 
  *************************************************************************
  -->
  
  <!-- 
  *************************************************************************
  MATCHING TEMPLATES
  *************************************************************************
  -->

  <!-- 
  Summary: frames containing external images
  Author: Clever Age
  -->
  <xsl:template match="draw:frame[not(./draw:object-ole or ./draw:object) and ./draw:image[@xlink:href and not(starts-with(@xlink:href, 'Pictures/'))]]" mode="paragraph">
    <!-- insert link to TOC field when required (user indexes) -->
    <xsl:call-template name="InsertTCField"/>

    <xsl:variable name="supported">
      <xsl:call-template name="IsImageSupportedByWord">
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
              <xsl:variable name="styleName" select=" parent::draw:frame/@draw:style-name"/>
              <xsl:variable name="automaticStyle" select="key('automatic-styles', parent::draw:frame/@draw:style-name)"/>
              <xsl:variable name="officeStyle" select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $styleName]"/>
              <xsl:variable name="frameStyle" select="$automaticStyle | $officeStyle"/>

              <xsl:call-template name="FrameToShapeProperties">
                <xsl:with-param name="frameStyle" select="$frameStyle"/>
                <xsl:with-param name="frame" select="parent::draw:frame"/>
              </xsl:call-template>

              <v:imagedata r:id="{generate-id(.)}" o:title=""/>

              <xsl:call-template name="FrameToShapeWrap">
                <xsl:with-param name="frameStyle" select="$frameStyle"/>
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
      <xsl:call-template name="IsImageSupportedByWord">
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
  *************************************************************************
  CALLED TEMPLATES
  *************************************************************************
  -->

  
  <!--  
  Summary:  Inserts a DrawingML "drawing" and decides if the image must be an 
            inline image or an anchored image.
  Author:   CleverAge
  -->
  <xsl:template name="InsertImage">

    <w:drawing>
      <xsl:variable name="imageId">
        <xsl:call-template name="GetPosition">
          <xsl:with-param name="node" select="."/>
        </xsl:call-template>
      </xsl:variable>

      <!-- width -->
      <xsl:variable name="cx">
        <xsl:call-template name="GetLengthOfFrameSide">
          <xsl:with-param name="side">width</xsl:with-param>
          <xsl:with-param name="frame" select="."/>
          <xsl:with-param name="unit">emu</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>

      <!-- height-->
      <xsl:variable name="cy">
        <xsl:call-template name="GetLengthOfFrameSide">
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
        <xsl:when test="ancestor::draw:text-box or @text:anchor-type='as-char' or ancestor::text:note[@text:note-class='endnote'] or $wrappedPara = 1">
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

  
  <!--
  Summary:  Inserts an inline image
  Author:   CleverAge
  -->
  <xsl:template name="InsertInlineImage">
    <xsl:param name="cx"/>
    <xsl:param name="cy"/>
    <xsl:param name="imageId"/>
    <xsl:param name="imageStyle"/>
    
    <wp:inline>
      <wp:extent cx="{$cx}" cy="{$cy}"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <xsl:call-template name="InsertDrawingMLGraphic">
        <xsl:with-param name="cx" select="$cx"/>
        <xsl:with-param name="cy" select="$cy"/>
        <xsl:with-param name="imageId" select="$imageId"/>
        <xsl:with-param name="imageStyle" select="$imageStyle"/>
      </xsl:call-template>
    </wp:inline>
    
  </xsl:template>


  <!-- 
  Summary:  Inserts an anchored Image
  Author:   CleverAge
  -->
  <xsl:template name="InsertAnchorImage">
    <xsl:param name="cx"/>
    <xsl:param name="cy"/>
    <xsl:param name="imageId"/>
    <xsl:param name="imageStyle"/>

    <wp:anchor simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">

      <!-- z-index -->
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
          <!-- divo: fix for roundtrip problem. This value may not be less than 0 -->
          <xsl:when test="@draw:z-index and @draw:z-index &lt; 0">
            <xsl:value-of select="'0'"/>
          </xsl:when>
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
      <xsl:call-template name="InsertDrawingMLGraphic">
        <xsl:with-param name="cx" select="$cx"/>
        <xsl:with-param name="cy" select="$cy"/>
        <xsl:with-param name="imageId" select="$imageId"/>
        <xsl:with-param name="imageStyle" select="$imageStyle"/>
      </xsl:call-template>
    </wp:anchor>
  </xsl:template>

  
  <!--
  Summary:  Converts the draw:image to a DrawingML graphic with properties
  Author:   CleverAge
  -->
  <xsl:template name="InsertDrawingMLGraphic">
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
                  select="key('automatic-styles', @draw:style-name)/style:graphic-properties/@fo:clip" />
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

  
  <!-- 
  Summary:  Checks if image type is supported in word.
  Author:   CleverAge
  -->
  <xsl:template name="IsImageSupportedByWord">
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


  <!-- 
  Summary:  Inserts the vertical and horizontal of the Image
  Author:   CleverAge
  Modified: makz (DIaLOGIKa)
  -->
  <xsl:template name="InsertAnchorImagePosition">
    <xsl:param name="imageStyle"/>

    <xsl:variable name="ox">
      <xsl:call-template name="emu-measure">
        <xsl:with-param name="length">
          <xsl:call-template name="ComputeMarginX">
            <xsl:with-param name="frame" select="ancestor-or-self::node()[contains(name(), 'draw:')]"/>
          </xsl:call-template>
          <xsl:text>pt</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

     <xsl:variable name="oy">
      <xsl:call-template name="emu-measure">
        <xsl:with-param name="length">
          <xsl:call-template name="ComputeMarginY">
            <xsl:with-param name="parent" select="ancestor-or-self::node()[contains(name(), 'draw:')]"/>
          </xsl:call-template>
          <xsl:text>pt</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>


    <!-- horizontal position -->
    
    <xsl:variable name="horizontal-pos">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$imageStyle"/>
        <xsl:with-param name="attribName">style:horizontal-pos</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <wp:simplePos x="0" y="0"/>
    <wp:positionH>
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
          <xsl:when test="$horizontalRel='paragraph-content'">column</xsl:when>
          <xsl:otherwise>margin</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
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

    <!-- vertical position -->

    <xsl:variable name="verticalRel">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$imageStyle"/>
        <xsl:with-param name="attribName">style:vertical-rel</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <wp:positionV>
      <!-- relative position -->
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


  <!--
  Summary:  Inserts the image cropping
  Author:   clam (DIaLOGIKa)
  -->
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
        <xsl:value-of select="concat('COMPUTEOOXCROPPING[', $firstVal, ',' , $filename,  ',t,' , round(($t * 100 div ($t + $b + $height)) * 1000), ']')"/>
      </xsl:attribute>
      <xsl:attribute name="b">
        <xsl:value-of select="concat('COMPUTEOOXCROPPING[', $thirdVal, ',' , $filename,  ',b,' , round(($b * 100 div ($t + $b + $height)) * 1000), ']')"/>
      </xsl:attribute>
      <xsl:attribute name="r">
        <xsl:value-of select="concat('COMPUTEOOXCROPPING[', $secondVal, ',' , $filename,  ',r,', round(($r * 100 div ($r + $l + $width)) * 1000), ']')"/>
      </xsl:attribute>
      <xsl:attribute name="l">
        <xsl:value-of select="concat('COMPUTEOOXCROPPING[', $fourthVal, ',' , $filename,  ',l,', round(($l * 100 div ($l + $r + $width)) * 1000), ']')"/>
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


  <!--
  Summary:  Inserts the wrapping for a anchored image
  Author:   CleverAge
  -->
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


  <!-- 
  Summary:  Insert wrapping extent
  Author:   CleverAge
  -->
  <xsl:template name="InsertWrapExtent">
    <xsl:param name="wrap"/>
    <xsl:param name="imageStyle"/>

    <xsl:choose>
      <xsl:when test="$wrap = 'parallel' or $wrap ='left' or $wrap = 'right' or $wrap ='dynamic' or $wrap = 'none' ">
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

  <!--
  Summary:  Computes the width of the image border
  Author:   CleverAge
  -->
  <xsl:template name="ComputeImageBorderWidth">
    <xsl:param name="side"/>
    <xsl:param name="imageStyle"/>

    <xsl:variable name="borderWidth">
      <xsl:choose>
        <!-- priority to side border -->
        <xsl:when test="$imageStyle/style:graphic-properties/@*[name() = concat('fo:border-', $side)]">
          <xsl:call-template name="GetConsistentBorderValue">
            <xsl:with-param name="side" select="$side"/>
            <xsl:with-param name="borderAttibute" select="$imageStyle/style:graphic-properties/@*[name() = concat('fo:border-', $side)]"/>
            <xsl:with-param name="node" select="$imageStyle/style:graphic-properties"/>
          </xsl:call-template>
        </xsl:when>
        <!-- otherwise if border attribute -->
        <xsl:when test="$imageStyle/style:graphic-properties/@fo:border">
          <xsl:call-template name="GetConsistentBorderValue">
            <xsl:with-param name="side" select="$side"/>
            <xsl:with-param name="borderAttibute" select="$imageStyle/style:graphic-properties/@fo:border"/>
            <xsl:with-param name="node" select="$imageStyle/style:graphic-properties"/>
          </xsl:call-template>
        </xsl:when>
        <!-- otherwise look parent style -->
        <xsl:when test="$imageStyle/@style:parent-style-name">
          <xsl:variable name="parentStyle" select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $imageStyle/@style:parent-style-name]"/>
          <xsl:call-template name="GetFramePadding">
            <xsl:with-param name="frameStyle" select="$parentStyle"/>
            <xsl:with-param name="side" select="$side"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
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


  <!-- 
  Summary:  Inserts a DrawingML square wrap
  Author:   CleverAge
  -->
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

  
  <!-- 
  Summary:  Inserts a DrawingML top-bottom wrap 
  Author:   CleverAge
  -->
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

  
  <!--
  Summary:  Inserts a DrawingML outline
  Author:   CleverAge
  -->
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

  
  <!-- 
  Summary:  Inserts a DrawingML effect list (only shadow at the moment)
  Author:   CleverAge
  -->
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

  <!-- 
  Summary:  Get the position of an element in the draw:frame group 
  Author:   Clever Age
  Modified: makz (DIaLOGIKa)
  Params:   node: The node whose position is requested.
  -->
  <xsl:template name="GetPosition">
    <xsl:param name="node" />

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
      <xsl:when test="string-length($positionInGroup) > 0">
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



</xsl:stylesheet>
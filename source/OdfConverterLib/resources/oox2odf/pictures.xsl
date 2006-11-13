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
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
  xmlns:uri="http://schemas.openxmlformats.org/drawingml/2006/picture"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  exclude-result-prefixes="w uri draw a pic">

  <!-- Pictures conversion needs copy of image files in zipEnrty to work correctly (but id does't crash  -->

  <xsl:template match="wp:inline | wp:anchor">
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
    </xsl:call-template>

    <draw:frame>
      <!-- TODO:@text:anchor-type -->
      <xsl:attribute name="text:anchor-type">
        <xsl:variable name="verticalRelativeFrom" select="descendant::wp:positionV/@relativeFrom"/>
        <xsl:choose>
          <xsl:when test="$verticalRelativeFrom = 'line'">
            <xsl:text>char</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>paragraph</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>

      <!--style name-->
      <xsl:attribute name="draw:style-name">
        <xsl:value-of select="generate-id(ancestor::w:drawing)"/>
      </xsl:attribute>

      <!--drawing name-->
      <xsl:attribute name="draw:name">
        <xsl:value-of select="wp:docPr/@name"/>
      </xsl:attribute>

      <!--size-->
      <xsl:call-template name="SetSize"/>

      <!--position-->
      <xsl:if test="self::wp:anchor">
        <xsl:call-template name="SetPosition"/>
      </xsl:if>

      <!-- image href from relationships-->
      <draw:image xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad">
        <xsl:if test="document(concat('word/_rels/',$document,'.rels'))">
          <xsl:call-template name="InsertImageHref">
            <xsl:with-param name="document" select="$document"/>
          </xsl:call-template>
        </xsl:if>
      </draw:image>
    </draw:frame>
  </xsl:template>

  <xsl:template name="SetSize">
    <xsl:attribute name="svg:height">
      <xsl:call-template name="emu-measure">
        <xsl:with-param name="length" select="wp:extent/@cy"/>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="svg:width">
      <xsl:call-template name="emu-measure">
        <xsl:with-param name="length" select="wp:extent/@cx"/>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>

  <xsl:template name="SetPosition">
    <xsl:attribute name="svg:x">
      <xsl:call-template name="emu-measure">
        <xsl:with-param name="length" select="wp:positionH/wp:posOffset"/>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="svg:y">
      <xsl:call-template name="emu-measure">
        <xsl:with-param name="length" select="wp:positionV/wp:posOffset"/>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>

  <!-- inserts image href from relationships -->
  <xsl:template name="InsertImageHref">
    <xsl:param name="document"/>

    <xsl:variable name="id">
      <xsl:value-of select="a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed"/>
    </xsl:variable>

    <xsl:for-each
      select="document(concat('word/_rels/',$document,'.rels'))//node()[name() = 'Relationship']">
      <xsl:if test="./@Id=$id">
        <xsl:variable name="pzipsource">
          <xsl:value-of select="./@Target"/>
        </xsl:variable>
        <xsl:variable name="pziptarget">
          <xsl:value-of select="substring-after($pzipsource,'/')"/>
        </xsl:variable>
        <xsl:attribute name="xlink:href">
          <xsl:value-of select="concat('pictures/', $pziptarget)"/>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="w:drawing" mode="automaticstyles">
    <style:style style:name="{generate-id(.)}" style:family="graphic">

      <!--in Word there are no parent style for image - make default Graphics in OO -->
      <xsl:attribute name="style:parent-style-name">
        <xsl:text>Graphics</xsl:text>
        <xsl:value-of select="w:tblStyle/@w:val"/>
      </xsl:attribute>

      <style:graphic-properties>
        <xsl:call-template name="InsertPictureProperties"/>
      </style:graphic-properties>
    </style:style>
  </xsl:template>

  <xsl:template name="InsertPictureProperties">
    <xsl:call-template name="InsertImagePosH"/>
    <xsl:call-template name="InsertImagePosV"/>
    <xsl:call-template name="InsertImageFlip"/>
    <xsl:call-template name="InsertImageCrop"/>
    <xsl:call-template name="InsertImageWrap"/>
    <xsl:call-template name="InsertImageMargins"/>
  </xsl:template>

  <xsl:template name="InsertImageWrap">
    <xsl:if test="descendant::node()[contains(name(),'wp:wrap')]">
      <xsl:attribute name="style:wrap">
        
        <xsl:if test="descendant::wp:wrapSquare">
          <xsl:call-template name="InsertSquareWrap">
            <xsl:with-param name="wrap" select="descendant::wp:wrapSquare/@wrapText"/>
          </xsl:call-template>
        </xsl:if>
        
        <xsl:if test="descendant::wp:wrapTight">
          <xsl:call-template name="InsertSquareWrap">
            <xsl:with-param name="wrap" select="descendant::wp:wrapTight/@wrapText"/>
          </xsl:call-template>
        </xsl:if>
        
        <xsl:if test="descendant::wp:wrapTopAndBottom">
          <xsl:text>none</xsl:text>
        </xsl:if>
        
        <xsl:if test="descendant::wp:wrapThrough">
          <xsl:call-template name="InsertSquareWrap">
            <xsl:with-param name="wrap" select="descendant::wp:wrapThrough/@wrapText"/>
          </xsl:call-template>
        </xsl:if>
      
        <xsl:if test="descendant::wp:wrapNone">
          <xsl:text>run-through</xsl:text>
        </xsl:if>
      </xsl:attribute>
      
      <!--decide if in backround or in front of text-->
      <xsl:if test="descendant::wp:anchor/@behindDoc = 1">
        <xsl:attribute name="style:run-through">
            <xsl:text>background</xsl:text>
        </xsl:attribute>
      </xsl:if>
   </xsl:if>
  </xsl:template>
  
  <xsl:template name="InsertImageMargins">
    <xsl:if test="descendant::wp:anchor">

        <xsl:attribute name="fo:margin-top">
          <xsl:call-template name="emu-measure">
            <xsl:with-param name="length" select="descendant::wp:anchor/@distT"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      
      <xsl:attribute name="fo:margin-bottom">
        <xsl:call-template name="emu-measure">
          <xsl:with-param name="length" select="descendant::wp:anchor/@distB"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
      
      <xsl:attribute name="fo:margin-left">
        <xsl:call-template name="emu-measure">
          <xsl:with-param name="length" select="descendant::wp:anchor/@distL"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
      
      <xsl:attribute name="fo:margin-right">
        <xsl:call-template name="emu-measure">
          <xsl:with-param name="length" select="descendant::wp:anchor/@distR"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>
  
  <xsl:template name="InsertSquareWrap">
    <xsl:param name="wrap"/>
    
    <xsl:choose>
      <xsl:when test="$wrap = 'bothSides' ">
        <xsl:text>parallel</xsl:text>
      </xsl:when>
      <xsl:when test="$wrap = 'largest' ">
        <xsl:text>dynamic</xsl:text>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$wrap"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="GetCropSize">
    <xsl:param name="cropValue"/>
    <xsl:param name="cropOppositeValue"/>
    <xsl:param name="resultSize"/>

    <xsl:choose>
      <xsl:when test="not($cropValue)">
        <xsl:text>0</xsl:text>
      </xsl:when>

      <xsl:otherwise>
        <xsl:variable name="cropPercent" select="$cropValue div (100000)"/>
        <xsl:variable name="cropOppositePercent">
          <xsl:choose>
            <xsl:when test="$cropOppositeValue">
              <xsl:value-of select="$cropOppositeValue div (100000)"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>0</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:value-of
          select="format-number(($resultSize div(1 - $cropPercent - $cropOppositePercent)) *  $cropPercent , '0.000' )"
        />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertImagePosH">
    <xsl:if test="descendant::wp:positionH">
      
      <xsl:if test="descendant::wp:positionH/wp:align">
        <xsl:attribute name="style:horizontal-pos">
          <xsl:value-of select="descendant::wp:positionH/wp:align"/>
        </xsl:attribute>
      </xsl:if>
      
      <xsl:if test="descendant::wp:positionH/@relativeFrom">
        <xsl:attribute name="style:horizontal-rel">
          <xsl:variable name="relativeFrom" select="descendant::wp:positionH/@relativeFrom"/>
          <xsl:choose>
            <xsl:when test="$relativeFrom = 'margin' or $relativeFrom = 'column'">
              <xsl:text>page-content</xsl:text>
            </xsl:when>
            <xsl:when test="$relativeFrom ='page'">
              <xsl:text>page</xsl:text>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'leftMargin' or $relativeFrom = 'outsideMargin'">
              <xsl:text>page-start-margin</xsl:text>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'rightMargin' or $relativeFrom = 'insideMargin'">
              <xsl:text>page-end-margin</xsl:text>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertImagePosV">
    <xsl:if test="descendant::wp:positionV">
      
      <xsl:if test="descendant::wp:positionV/wp:align">
        <xsl:attribute name="style:vertical-pos">
          <xsl:variable name="align" select="descendant::wp:positionV/wp:align"/>
          <xsl:variable name="relativeFrom" select="descendant::wp:positionV/@relativeFrom"/>
          <xsl:choose>
            <!--special rules-->
            <xsl:when
              test="$relativeFrom = 'topMargin' or $relativeFrom = 'bottomMargin' or $relativeFrom = 'insideMargin' or $relativeFrom = 'outsideMargin'">
              <xsl:text>top</xsl:text>
            </xsl:when>
            <!--default rules-->
            <xsl:when test="$align = 'top' ">
              <xsl:text>top</xsl:text>
            </xsl:when>
            <xsl:when test="$align = 'center' ">
              <xsl:text>middle</xsl:text>
            </xsl:when>
            <xsl:when test="$align = 'bottom' ">
              <xsl:text>bottom</xsl:text>
            </xsl:when>
            <xsl:when test="$align = 'inside' ">
              <xsl:text>top</xsl:text>
            </xsl:when>
            <xsl:when test="$align = 'outside' ">
              <xsl:text>bottom</xsl:text>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="descendant::wp:positionV/@relativeFrom">
        <xsl:attribute name="style:vertical-rel">
          <xsl:variable name="relativeFrom" select="descendant::wp:positionV/@relativeFrom"/>
          <xsl:choose>
            <xsl:when test="$relativeFrom = 'margin' ">
              <xsl:text>page-content</xsl:text>
            </xsl:when>
            <xsl:when test="$relativeFrom ='page'">
              <xsl:text>page</xsl:text>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'topMargin'">
              <xsl:text>page</xsl:text>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'bottomMargin'">
              <xsl:text>page</xsl:text>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'insideMargin'">
              <xsl:text>page</xsl:text>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'outsideMargin'">
              <xsl:text>page</xsl:text>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'line'">
              <xsl:text>line</xsl:text>
            </xsl:when>
            <xsl:when test="$relativeFrom = 'paragraph'">
              <xsl:text>paragraph-content</xsl:text>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertImageFlip">
    <!--  picture flip (vertical, horizontal)-->
    <xsl:if test="descendant::pic:spPr/a:xfrm/attribute::node()">
      <xsl:for-each select="descendant::pic:pic">
        <xsl:attribute name="style:mirror">
          <xsl:choose>
            <xsl:when test="pic:spPr/a:xfrm/@flipV = '1' and pic:spPr/a:xfrm/@flipH = '1'">
              <xsl:text>horizontal vertical</xsl:text>
            </xsl:when>
            <xsl:when test="pic:spPr/a:xfrm/@flipV = '1' ">
              <xsl:text>vertical</xsl:text>
            </xsl:when>
            <xsl:when test="pic:spPr/a:xfrm/@flipH = '1' ">
              <xsl:text>horizontal</xsl:text>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertImageCrop">
    <!-- crop -->
    <xsl:if test="pic:blipFill/a:srcRect/attribute::node()">
      <xsl:for-each select="descendant::pic:pic">
        <xsl:variable name="width">
          <xsl:variable name="widthText">
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="ancestor::w:drawing/descendant::wp:extent/@cx"/>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:value-of select="substring-before($widthText,'cm')"/>
        </xsl:variable>

        <xsl:variable name="height">
          <xsl:variable name="heightText">
            <xsl:call-template name="emu-measure">
              <xsl:with-param name="length" select="ancestor::w:drawing/descendant::wp:extent/@cy"/>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:value-of select="substring-before($heightText,'cm')"/>
        </xsl:variable>

        <xsl:variable name="leftCrop">
          <xsl:call-template name="GetCropSize">
            <xsl:with-param name="cropValue" select="pic:blipFill/a:srcRect/@l"/>
            <xsl:with-param name="cropOppositeValue" select="pic:blipFill/a:srcRect/@r"/>
            <xsl:with-param name="resultSize" select="$width"/>
          </xsl:call-template>
        </xsl:variable>

        <xsl:variable name="rightCrop">
          <xsl:call-template name="GetCropSize">
            <xsl:with-param name="cropValue" select="pic:blipFill/a:srcRect/@r"/>
            <xsl:with-param name="cropOppositeValue" select="pic:blipFill/a:srcRect/@l"/>
            <xsl:with-param name="resultSize" select="$width"/>
          </xsl:call-template>
        </xsl:variable>

        <xsl:variable name="topCrop">
          <xsl:call-template name="GetCropSize">
            <xsl:with-param name="cropValue" select="pic:blipFill/a:srcRect/@t"/>
            <xsl:with-param name="cropOppositeValue" select="pic:blipFill/a:srcRect/@b"/>
            <xsl:with-param name="resultSize" select="$height"/>
          </xsl:call-template>
        </xsl:variable>

        <xsl:variable name="bottomCrop">
          <xsl:call-template name="GetCropSize">
            <xsl:with-param name="cropValue" select="pic:blipFill/a:srcRect/@b"/>
            <xsl:with-param name="cropOppositeValue" select="pic:blipFill/a:srcRect/@t"/>
            <xsl:with-param name="resultSize" select="$height"/>
          </xsl:call-template>
        </xsl:variable>

        <xsl:attribute name="fo:clip">
          <xsl:value-of
            select="concat('rect(',$topCrop,'cm ',$rightCrop,'cm ',$bottomCrop,'cm ',$leftCrop,'cm',')')"
          />
        </xsl:attribute>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>
  
  <xsl:template match="w:drawing">
    <xsl:apply-templates select="wp:inline | wp:anchor"/>
  </xsl:template>
</xsl:stylesheet>

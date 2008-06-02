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

  <xsl:key name="images" match="draw:frame[not(./draw:object-ole or ./draw:object)]/draw:image[@xlink:href]" use="''"/>
  <xsl:key name="frames" match="draw:frame" use="''"/>
  <xsl:key name="automatic-styles" match="office:automatic-styles/style:style" use="@style:name"/>

  <!--
  *************************************************************************
  SUMMARY
  *************************************************************************
  This stylesheet handles the conversion of draw:textbox elements and 
  contains some general templates concerning draw:frame conversions.
  
  Templates that handles the conversion of pictures, shapes and ole-objects
  (which are also children of draw:frame), can be found in 2oox-ole.xsl, 
  2oox-shapes.xsl and 2oox-pictures.xsl.
  *************************************************************************
  -->
  
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
    
    <xsl:apply-templates select="draw:text-box" mode="paragraph"/>
  </xsl:template>

  <!-- 
  Summary:  Embedd all consecutive frames that are not inserted into a paragraph in a 
            single paragraph (avoid paragraph not present in original document).
  Author:   CleverAge
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
  Summary:  Converts a text box in a frame to a VML shape
  Author:   Clever Age
  -->
  <xsl:template match="draw:text-box" mode="paragraph">
    <w:r>
      <w:rPr>
        <xsl:variable name="prefixedStyleName">
          <xsl:call-template name="GetPrefixedStyleName">
            <xsl:with-param name="styleName" select="parent::draw:frame/@draw:style-name"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:if test="$prefixedStyleName!=''">
          <w:rStyle w:val="{$prefixedStyleName}"/>
        </xsl:if>
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
          <xsl:variable name="officeStyle" select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $styleName]"/>
          <xsl:variable name="frameStyle" select="$automaticStyle | $officeStyle"/>

          <xsl:call-template name="FrameToShapeProperties">
            <xsl:with-param name="frameStyle" select="$frameStyle"/>
            <xsl:with-param name="frame" select="parent::draw:frame"/>
          </xsl:call-template>

          <xsl:call-template name="FrameToShapeWrap">
            <xsl:with-param name="frameStyle" select="$frameStyle"/>
          </xsl:call-template>

          <xsl:call-template name="InsertTextBox">
            <xsl:with-param name="frameStyle" select="$frameStyle"/>
            <xsl:with-param name="frame" select="parent::draw:frame"/>
          </xsl:call-template>
        </v:shape>
      </w:pict>
    </w:r>
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
  *************************************************************************
  CALLED TEMPLATES
  *************************************************************************
  -->

  <!-- 
  Summary:  Inserts the properties of the VML shape
  Author:   CleverAge
  Params:   frameStyle: The automatic style of the draw:frame
            frame: The draw:frame itself
  -->
  <xsl:template name="FrameToShapeProperties">
    <xsl:param name="frameStyle"/>
    <xsl:param name="frame"/>

    <!-- report lost properties -->
    <xsl:if test="$frameStyle/style:graphic-properties/@draw:textarea-vertical-align != 'top' ">
      <xsl:message terminate="no">translation.odf2oox.valignInsideTextbox</xsl:message>
    </xsl:if>

    <xsl:if test="$frame/@draw:name">
      <xsl:attribute name="id">
        <xsl:value-of select="$frame/@draw:name"/>
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
      <xsl:call-template name="FrameToShapeSize">
        <xsl:with-param name="frame" select="$frame"/>
      </xsl:call-template>

      <xsl:call-template name="FrameToRelativeShapeSize">
        <xsl:with-param name="frame" select="$frame"/>
      </xsl:call-template>

      <xsl:call-template name="FrameToShapeZindex">
        <xsl:with-param name="frameStyle" select="$frameStyle"/>
        <xsl:with-param name="frame" select="$frame"/>
      </xsl:call-template>

      <xsl:call-template name="FrameToRelativeShapePosition">
        <xsl:with-param name="frameStyle" select="$frameStyle"/>
        <xsl:with-param name="frame" select="$frame"/>
      </xsl:call-template>

      <xsl:call-template name="FrameToShapePosition">
        <xsl:with-param name="frameStyle" select="$frameStyle"/>
        <xsl:with-param name="frame" select="$frame"/>
      </xsl:call-template>

      <xsl:call-template name="FrameToShapeMargin">
        <xsl:with-param name="frameStyle" select="$frameStyle"/>
        <xsl:with-param name="frame" select="$frame"/>
      </xsl:call-template>

      <xsl:call-template name="FrameToShapeRotation">
        <xsl:with-param name="frame" select="$frame"/>
      </xsl:call-template>

      <xsl:call-template name="FrameToTextAnchor">
        <xsl:with-param name="frameStyle" select="$frameStyle"/>
      </xsl:call-template>
    </xsl:attribute>

    <xsl:call-template name="FrameToShapeFill">
      <xsl:with-param name="frameStyle" select="$frameStyle"/>
    </xsl:call-template>
    
    <xsl:call-template name="FrameToShapeBorders">
      <xsl:with-param name="frameStyle" select="$frameStyle"/>
    </xsl:call-template>

    <xsl:call-template name="FrameToShapeShadow">
      <xsl:with-param name="frameStyle" select="$frameStyle"/>
    </xsl:call-template>
  </xsl:template>


  <!--
  Summary:  Inserts the width: and height: values into teh style attribute.
  Author:   CleverAge
  Modified: makz (DIaLOGIKa)
  Params:   frame: The draw:frame
  -->
  <xsl:template name="FrameToShapeSize">
    <xsl:param name="frame"/>

    <!-- width -->
    <xsl:variable name="frameW">
      <xsl:call-template name="GetLengthOfFrameSide">
        <xsl:with-param name="side">width</xsl:with-param>
        <xsl:with-param name="frame" select="$frame"/>
        <xsl:with-param name="unit">point</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:if test="$frameW != '' and $frameW != 0">
      <xsl:value-of select="concat('width:',$frameW,'pt;')"/>
    </xsl:if>

    <!-- height-->
    <xsl:variable name="frameH">
      <xsl:call-template name="GetLengthOfFrameSide">
        <xsl:with-param name="side">height</xsl:with-param>
        <xsl:with-param name="frame" select="$frame"/>
        <xsl:with-param name="unit">point</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:if test="$frameH != '' and $frameH != 0">
      <xsl:value-of select="concat('height:',$frameH,'pt;')"/>
    </xsl:if>
  </xsl:template>
  

  <!--
  Summary:  Inserts the values for relative sized shapes into the style attribute.
  Author:   CleverAge
  Modified: makz (DIaLOGIKa)
  Params:   frame: The draw:frame
  -->
  <xsl:template name="FrameToRelativeShapeSize">
    <xsl:param name="frame"/>

    <xsl:if test="$frame/@style:rel-width or $frame/@style:rel-height">

      <!-- relative to -->
      <xsl:variable name="relativeTo">
        <xsl:choose>
          <xsl:when test="$frame/@text:anchor-type = 'page'">
            <xsl:text>page</xsl:text>
          </xsl:when>
          <xsl:when test="$frame/@text:anchor-type = 'paragraph'">
            <xsl:text>margin</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>margin</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>

      <!-- relative width -->
      <xsl:choose>
        <xsl:when test="contains($frame/@style:rel-width, 'scale')">
          <!-- warn loss of scaled images (scale, scale-min) -->
          <xsl:message terminate="no">translation.odf2oox.scaledImage</xsl:message>
        </xsl:when>
        <xsl:otherwise>
          <xsl:variable name="relWidth">
            <xsl:choose>
              <xsl:when test="contains($frame/@style:rel-width,'%')">
                <xsl:value-of select="substring-before($frame/@style:rel-width,'%')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$frame/@style:rel-width"/>
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
        <xsl:when test="contains($frame/@style:rel-height, 'scale')">
          <!-- warn loss of scaled images (scale, scale-min) -->
          <xsl:message terminate="no">translation.odf2oox.scaledImage</xsl:message>
        </xsl:when>
        <xsl:otherwise>
          <xsl:variable name="relHeight">
            <xsl:choose>
              <xsl:when test="contains($frame/@style:rel-height,'%')">
                <xsl:value-of select="substring-before($frame/@style:rel-height,'%')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$frame/@style:rel-height"/>
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


  <!--
  Summary:  Inserts the z-index value into the style attribute.
  Author:   CleverAge
  Modified: makz (DIaLOGIKa)
  Params:   frameStyle: The automatic style of the draw:frame
            frame: The draw:frame
  -->
  <xsl:template name="FrameToShapeZindex">
    <xsl:param name="frameStyle"/>
    <xsl:param name="frame"/>

    <xsl:variable name="frameWrap">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$frameStyle"/>
        <xsl:with-param name="attribName">style:wrap</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="runThrought">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$frameStyle"/>
        <xsl:with-param name="attribName">style:run-through</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>


    <!--z-index that we need to convert properly openoffice wrap-throught property -->
    <xsl:variable name="zIndex">
      <xsl:choose>
        <xsl:when test="$frameWrap='run-through' and $runThrought='background'">
          -251658240
        </xsl:when>
        <xsl:when test="$frameWrap='run-through' and not($runThrought)">
          251658240
        </xsl:when>
        <xsl:otherwise>
          <xsl:choose>
            <xsl:when test="$frame/@draw:z-index">
              <xsl:value-of select="2 + $frame/@draw:z-index"/>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:value-of select="concat('z-index:', $zIndex, ';')"/>
  </xsl:template>


  <!--
  Summary:  Inserts the values for relative positioning to the style attribute.
  Author:   CleverAge
  Modified: makz (DIaLOGIKa)
  Params:   frameStyle: The automatic style of the draw:frame
            frame: The draw:frame
  -->
  <xsl:template name="FrameToRelativeShapePosition">
    <xsl:param name="frameStyle"/>
    <xsl:param name="frame"/>

    <xsl:variable name="anchor">
      <xsl:value-of select="$frame/@text:anchor-type"/>
    </xsl:variable>
    <xsl:variable name="horizontalRel">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$frameStyle"/>
        <xsl:with-param name="attribName">style:horizontal-rel</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="verticalRel">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$frameStyle"/>
        <xsl:with-param name="attribName">style:vertical-rel</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="wrappedPara">
      <xsl:variable name="wrapping">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$frameStyle"/>
          <xsl:with-param name="attribName">style:wrap</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:if test="$wrapping = 'parallel' ">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$frameStyle"/>
          <xsl:with-param name="attribName">style:number-wrapped-paragraphs</xsl:with-param>
        </xsl:call-template>
      </xsl:if>
    </xsl:variable>

    <xsl:choose>
      <!-- inline image -->
      <xsl:when test="$wrappedPara = 1">mso-position-horizontal-relative:char;</xsl:when>
      <!-- direct anchored frame -->
      <xsl:when test="$anchor = 'as-char' ">mso-position-horizontal-relative:char;</xsl:when>
      <xsl:when test="$anchor = 'page' and $frame/@svg:x">
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

    <xsl:choose>
      <!-- inline image -->
      <xsl:when test="$wrappedPara = 1">mso-position-vertical-relative:line;</xsl:when>
      <!-- direct anchored frame -->
      <xsl:when test="$anchor = 'as-char' ">mso-position-horizontal-relative:line;</xsl:when>
      <xsl:when test="$anchor = 'page' and $frame/@svg:y">
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
          <xsl:when test="$verticalRel='page'">
            <xsl:text>mso-position-vertical-relative:page;</xsl:text>
          </xsl:when>
          <!-- page-content -->
          <xsl:when test="$verticalRel='page-content'">
            <xsl:text>mso-position-vertical-relative:margin;</xsl:text>
          </xsl:when>
          <!-- paragraph -->
          <xsl:when test="$verticalRel='paragraph'">
            <xsl:text>mso-position-vertical-relative:text;</xsl:text>
          </xsl:when>
          <!-- paragraph-content -->
          <xsl:when test="$verticalRel='paragraph-content'">
            <xsl:text>mso-position-vertical-relative:text;</xsl:text>
          </xsl:when>
          <!-- frame, frame-content -->
          <xsl:when test="contains($verticalRel, 'frame')">
            <xsl:text>mso-position-vertical-relative:text;</xsl:text>
          </xsl:when>
          <!-- baseline -->
          <xsl:when test="$verticalRel='baseline'">
            <xsl:choose>
              <xsl:when test="$anchor='paragraph'">
                <!-- 
                makz:
                paragraph anchor should be converted to paragraph ...
                commented out for bugfix 1947870:
                <xsl:text>mso-position-vertical-relative:margin;</xsl:text>
                -->
                <xsl:text>mso-position-vertical-relative:paragraph;</xsl:text>
              </xsl:when>
              <xsl:otherwise>
                <xsl:text>mso-position-vertical-relative:char;</xsl:text>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <!-- char, line, text -->
          <xsl:when test="$verticalRel='char' or $verticalRel='line' or $verticalRel='text'">
            <xsl:text>mso-position-vertical-relative:char;</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <!-- no default value suggested. use anchor -->
            <xsl:choose>
              <xsl:when test="$anchor='page'">
                <xsl:text>mso-position-horizontal-relative:page;</xsl:text>
              </xsl:when>
              <xsl:when test="$anchor='paragraph' or $anchor='frame'">
                <xsl:text>mso-position-horizontal-relative:text;</xsl:text>
              </xsl:when>
              <xsl:when test="$anchor='char'">
                <xsl:text>mso-position-horizontal-relative:line;</xsl:text>
              </xsl:when>
              <xsl:otherwise>
                <!-- 
                as-char anchor already handled (cf above). 
                In case nothing is ever specified : use default = text 
                -->
                <xsl:text>mso-position-horizontal-relative:text;</xsl:text>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  
  <!--
  Summary:  Inserts the values for positioning to the style attribute.
  Author:   CleverAge
  Modified: makz (DIaLOGIKa)
  Params:   frameStyle: The automatic style of the draw:frame
            frame: The draw:frame
  -->
  <xsl:template name="FrameToShapePosition">
    <xsl:param name="frameStyle"/>
    <xsl:param name="frame"/>

    <xsl:variable name="graphicProps" select="$frameStyle/style:graphic-properties"/>
    <xsl:variable name="anchor">
      <xsl:value-of select="$frame/@text:anchor-type"/>
    </xsl:variable>
    <xsl:variable name="wrappedPara">
      <xsl:variable name="wrapping">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$frameStyle"/>
          <xsl:with-param name="attribName">style:wrap</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:if test="$wrapping = 'parallel' ">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$frameStyle"/>
          <xsl:with-param name="attribName">style:number-wrapped-paragraphs</xsl:with-param>
        </xsl:call-template>
      </xsl:if>
    </xsl:variable>

    <!-- if inline image, no positioning -->
    <xsl:if test="not($wrappedPara=1 or $anchor='as-char')">

      <!-- two cases : absolute (=>define margin-left and margin-top), or mso-position-horizontal / -vertical -->
      <xsl:variable name="horizontalPos">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$frameStyle"/>
          <xsl:with-param name="attribName">style:horizontal-pos</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="horizontalRel">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$frameStyle"/>
          <xsl:with-param name="attribName">style:horizontal-rel</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="marginLeft">
        <xsl:call-template name="GetValue">
          <xsl:with-param name="length">
            <xsl:call-template name="GetGraphicProperties">
              <xsl:with-param name="shapeStyle" select="$frameStyle"/>
              <xsl:with-param name="attribName">fo:margin-left</xsl:with-param>
            </xsl:call-template>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="marginRight">
        <xsl:call-template name="GetValue">
          <xsl:with-param name="length">
            <xsl:call-template name="GetGraphicProperties">
              <xsl:with-param name="shapeStyle" select="$frameStyle"/>
              <xsl:with-param name="attribName">fo:margin-right</xsl:with-param>
            </xsl:call-template>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="xCoordinate">
        <xsl:call-template name="GetValue">
          <xsl:with-param name="length">
            <xsl:call-template name="GetGraphicProperties">
              <xsl:with-param name="shapeStyle" select="$frameStyle"/>
              <xsl:with-param name="attribName">svg:x</xsl:with-param>
            </xsl:call-template>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="verticalPos">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$frameStyle"/>
          <xsl:with-param name="attribName">style:vertical-pos</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="verticalRel">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$frameStyle"/>
          <xsl:with-param name="attribName">style:vertical-rel</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="marginTop">
        <xsl:call-template name="GetValue">
          <xsl:with-param name="length">
            <xsl:call-template name="GetGraphicProperties">
              <xsl:with-param name="shapeStyle" select="$frameStyle"/>
              <xsl:with-param name="attribName">fo:margin-top</xsl:with-param>
            </xsl:call-template>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="marginBottom">
        <xsl:call-template name="GetValue">
          <xsl:with-param name="length">
            <xsl:call-template name="GetGraphicProperties">
              <xsl:with-param name="shapeStyle" select="$frameStyle"/>
              <xsl:with-param name="attribName">fo:margin-bottom</xsl:with-param>
            </xsl:call-template>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="yCoordinate">
        <xsl:call-template name="GetValue">
          <xsl:with-param name="length">
            <xsl:call-template name="GetGraphicProperties">
              <xsl:with-param name="shapeStyle" select="$frameStyle"/>
              <xsl:with-param name="attribName">svg:y</xsl:with-param>
            </xsl:call-template>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="rotation">
        <xsl:if test="contains($frame/@draw:transform,'rotate')">
          <xsl:call-template name="DegreesAngle">
            <xsl:with-param name="angle">
              <xsl:value-of
                select="substring-before(substring-after(substring-after($frame/@draw:transform,'rotate'),'('),')')"
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
        <xsl:if test="$horizontalPos = 'from-left' or $horizontalPos='from-inside' or ($marginLeft != '' and $marginLeft != 0 ) or ($marginRight != '' and $marginRight != 0 ) ">
          <xsl:choose>
            <!-- if rotation, revert X and Y -->
            <xsl:when test="$rotation != '' ">
              <xsl:call-template name="ComputeMarginXWithRotation">
                <xsl:with-param name="angle" select="$rotation"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="ComputeMarginX">
                <xsl:with-param name="frame" select="ancestor-or-self::node()[contains(name(), 'draw:')]"/>
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
        <xsl:if test="$verticalPos='from-top' or ($marginTop != '' and $marginTop != 0 ) or ($marginBottom != '' and $marginBottom != 0 ) ">
          <xsl:choose>
            <!-- if rotation, revert X and Y -->
            <xsl:when test="$rotation != '' ">
              <xsl:call-template name="ComputeMarginYWithRotation">
                <xsl:with-param name="angle" select="$rotation"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="ComputeMarginY">
                <xsl:with-param name="parent" select="ancestor-or-self::node()[contains(name(), 'draw:')]"/>
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
        <xsl:when test="ancestor::*[name()='style:header'] or ancestor::*[name()='style:footer']">
          <!-- shape is placed in header or footer -->
          <xsl:text>mso-position-vertical:paragraph;</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <!-- shape is placed in main document -->
          <xsl:choose>
            <xsl:when test="$verticalPos = 'middle' ">
              <xsl:text>mso-position-vertical:center;</xsl:text>
            </xsl:when>
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
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
  </xsl:template>


  <!--
  Summary:  Inserts the values for margins to the style attribute.
  Author:   CleverAge
  Modified: makz (DIaLOGIKa)
  Params:   frameStyle: The automatic style of the draw:frame
            frame: The draw:frame
  -->
  <xsl:template name="FrameToShapeMargin">
    <xsl:param name="frameStyle"/>
    <xsl:param name="frame"/>

    <!-- wrapping of text (horizontal adjustment) -->
    <xsl:if test="$frame/@fo:min-width or $frame/draw:text-box/@fo:min-height or $frameStyle/style:graphic-properties/@draw:auto-grow-width = 'true' ">
      <!--
      makz: was uncommented in r2655 by rebet
      I commented it in for bugfix 1827515
      -->
      <xsl:text>mso-wrap-style:none;</xsl:text>
    </xsl:if>

    <!--text-box spacing/margins -->
    <xsl:variable name="marginL">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$frameStyle"/>
        <xsl:with-param name="attribName">fo:margin-left</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="marginT">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$frameStyle"/>
        <xsl:with-param name="attribName">fo:margin-top</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="marginR">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$frameStyle"/>
        <xsl:with-param name="attribName">fo:margin-right</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="marginB">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$frameStyle"/>
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

  
  <!--
  Summary:  Inserts the rotation value to the style attribute.
  Author:   CleverAge
  Modified: makz (DIaLOGIKa)
  Params:   frame: The draw:frame
  -->
  <xsl:template name="FrameToShapeRotation">
    <xsl:param name="frame"/>

    <xsl:if test="contains($frame/@draw:transform,'rotate')">
      <xsl:text>rotation:</xsl:text>
      <xsl:variable name="angle">
        <xsl:call-template name="DegreesAngle">
          <xsl:with-param name="angle">
            <xsl:value-of
              select="substring-before(substring-after(substring-after($frame/@draw:transform,'rotate'),'('),')')"
            />
          </xsl:with-param>
          <xsl:with-param name="revert">true</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:value-of select="$angle"/>
      <xsl:text>;</xsl:text>
    </xsl:if>
  </xsl:template>


  <!-- 
  Summary:  Inserts the vertical anchor to the style attribute
  Author:   Clever Age
  Modified: makz (DIaLOGIka)
  Params:   frameStyle: The automatic style of the draw:frame
  -->
  <xsl:template name="FrameToTextAnchor">
    <xsl:param name="frameStyle"/>

    <xsl:variable name="verticalAlign">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$frameStyle"/>
        <xsl:with-param name="attribName">draw:textarea-vertical-align</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:text>v-text-anchor:</xsl:text>
    <xsl:choose>
      <xsl:when test="$verticalAlign = '' or $verticalAlign = 'top' ">
        <xsl:text>top</xsl:text>
      </xsl:when>
      <xsl:when test="$verticalAlign = 'middle' or $verticalAlign = 'justify' ">
        <xsl:text>middle</xsl:text>
      </xsl:when>
      <xsl:when test="$verticalAlign ='bottom' ">
        <xsl:text>bottom</xsl:text>
      </xsl:when>
      <xsl:otherwise>
        <xsl:text>top</xsl:text>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:text>;</xsl:text>
  </xsl:template>


  <!-- 
  Summary:  Inserts the VML shape fill
  Author:   CleverAge
  Modified: makz (DIaLOGIka)
  Params:   frameStyle: The automatic style of the draw:frame
  -->
  <xsl:template name="FrameToShapeFill">
    <xsl:param name="frameStyle"/>

    <xsl:variable name="shapefillColor">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$frameStyle"/>
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

    <xsl:if test="$fillColor != '' ">
      <xsl:attribute name="fillcolor">
        <xsl:value-of select="$fillColor"/>
      </xsl:attribute>
    </xsl:if>

  </xsl:template>

  
  <!-- 
  Summary:  Inserts the VML shape stroke 
  Author:   CleverAge
  Modified: makz (DIaLOGIka)
  Params:   frameStyle: The automatic style of the draw:frame
  -->
  <xsl:template name="FrameToShapeBorders">
    <xsl:param name="frameStyle"/>

    <xsl:variable name="border">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$frameStyle"/>
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
            <xsl:with-param name="shapeStyle" select="$frameStyle"/>
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

  
  <!-- 
  Summary:  Inserts the VML shape shadow element 
  Author:   CleverAge
  Modified: makz (DIaLOGIka)
  Params:   frameStyle: The automatic style of the draw:frame
  -->
  <xsl:template name="FrameToShapeShadow">
    <xsl:param name="frameStyle"/>

    <xsl:variable name="shadow">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$frameStyle"/>
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

        <!--dialogika, clam: bugfix #1826917-->
        <xsl:variable name="transparency">
          <xsl:call-template name="GetGraphicProperties">
            <xsl:with-param name="shapeStyle" select="$frameStyle"/>
            <xsl:with-param name="attribName">style:background-transparency</xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:if test="$transparency = '100%'">
          <xsl:attribute name="obscured">true</xsl:attribute>
        </xsl:if>
      </v:shadow>
    </xsl:if>

  </xsl:template>


  <!--
  Summary:  Inserts the wrap element for VML shapes
  Author:   Clever Age
  Modified: makz (DIaLOGIKa)
  Params:   frameStyle: The automatic style of the draw:frame
  -->
  <xsl:template name="FrameToShapeWrap">
    <xsl:param name="frameStyle"/>

    <xsl:variable name="frameWrap">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$frameStyle"/>
        <xsl:with-param name="attribName">style:wrap</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="parent::node()[contains(name(), 'draw:')]/@text:anchor-type = 'as-char' ">
        <w10:wrap type="none"/>
        <w10:anchorlock/>
      </xsl:when>
      <xsl:when test="$frameWrap='left'">
        <w10:wrap type="square" side="left"/>
      </xsl:when>
      <xsl:when test="$frameWrap='right'">
        <w10:wrap type="square" side="right"/>
      </xsl:when>
      <xsl:when test="$frameWrap='none'">
        <xsl:message terminate="no">translation.odf2oox.shapeTopBottomWrapping</xsl:message>
        <w10:wrap type="topAndBottom"/>
      </xsl:when>
      <xsl:when test="$frameWrap='parallel'">
        <xsl:variable name="wrappedPara">
          <xsl:call-template name="GetGraphicProperties">
            <xsl:with-param name="shapeStyle" select="$frameStyle"/>
            <xsl:with-param name="attribName">style:number-wrapped-paragraphs</xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="$wrappedPara=1">
            <w10:wrap type="none"/>
            <w10:anchorlock/>
          </xsl:when>
          <xsl:otherwise>
            <w10:wrap type="tight"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="$frameWrap='dynamic'">
        <!--
        makz: generally it's not possible to find a perfect value for a dynamic wrap.
        But I think topAndBottom is in the most cases a better solution than square
        -->
        <w10:wrap type="topAndBottom"/>
        <!--<w10:wrap type="square" side="largest"/>-->
      </xsl:when>
    </xsl:choose>

  </xsl:template>


  <!--
  Summary:  Inserts a VML textbox element
  Author:   Clever Age
  Modified: makz (DIaLOGIKa)
  Params:   frameStyle: The automatic style of the draw:frame
            frame: The draw:frame
  -->
  <xsl:template name="InsertTextBox">
    <xsl:param name="frameStyle"/>
    <xsl:param name="frame"/>

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
            <xsl:when test="$frame/@svg:width='0' and $frame/@svg:height='0'">
              <xsl:text>mso-fit-shape-to-text:t;</xsl:text>
            </xsl:when>
            <xsl:when test="$frameStyle/@style:parent-style-name">
              <xsl:if test="@fo:min-height or parent::draw:frame/@fo:min-width
                  or $frameStyle/style:graphic-properties/@draw:auto-grow-width = 'true'
                  or $frameStyle/style:graphic-properties/@draw:auto-grow-height = 'true'">
                <xsl:text>mso-fit-shape-to-text:t;</xsl:text>
              </xsl:if>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="not(parent::node()='draw:frame') 
                and (not($frameStyle/style:graphic-properties/@draw:auto-grow-width = 'false') 
                or not($frameStyle/style:graphic-properties/@draw:auto-grow-height = 'false'))">
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

      <xsl:attribute name="inset">
        <!-- left padding -->
        <xsl:call-template name="milimeter-measure">
          <xsl:with-param name="length">
            <xsl:call-template name="GetFramePadding">
              <xsl:with-param name="frameStyle" select="$frameStyle"/>
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
              <xsl:with-param name="frameStyle" select="$frameStyle"/>
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
              <xsl:with-param name="frameStyle" select="$frameStyle"/>
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
              <xsl:with-param name="frameStyle" select="$frameStyle"/>
              <xsl:with-param name="side">bottom</xsl:with-param>
            </xsl:call-template>
          </xsl:with-param>
          <xsl:with-param name="round">false</xsl:with-param>
        </xsl:call-template>
        <xsl:text>mm</xsl:text>
      </xsl:attribute>

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
                    <xsl:call-template name="FrameToTextboxAlignment">
                      <xsl:with-param name="frame" select="draw:frame"/>
                    </xsl:call-template>
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


  <!-- 
  Summary:  Gets the length of a specified side of a draw:frame
  Author:   Clever Age
  Modified: makz (DIaLOGIKa)
  Params:   frame: The draw:frame
            side: The values 'height' or 'width'
            unit: The target unit
  -->
  <xsl:template name="GetLengthOfFrameSide">
    <xsl:param name="frame"/>
    <xsl:param name="side"/>
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

  
  <!-- 
  Summary:  Handles the frames that are ignored. Context is text:p 
  Author:   Clever Age
  -->
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

  
  <!-- 
  Summary:  Computes the margin of a frame. Returns a value in points.
  Author:   Clever Age
  -->
  <xsl:template name="ComputeMarginX">
    <xsl:param name="frame"/>
    <xsl:variable name="frameStyle" select="key('styles', $frame[1]/@draw:style-name)"/>
    
    <xsl:choose>
      <xsl:when test="$frame">

        <xsl:variable name="recursive_result">
          <xsl:call-template name="ComputeMarginX">
            <xsl:with-param name="frame" select="$frame[position()>1]"/>
          </xsl:call-template>
        </xsl:variable>

        <xsl:choose>
          <xsl:when test="count($frameStyle) = 0">
            <xsl:value-of select="$recursive_result"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:variable name="anchor">
              <xsl:value-of select="$frame[1]/@text:anchor-type"/>
            </xsl:variable>
            <xsl:variable name="horizontalPos">
              <xsl:value-of select="$frameStyle/style:graphic-properties/@style:horizontal-pos"/>
            </xsl:variable>
            <xsl:variable name="horizontalRel">
              <xsl:value-of select="$frameStyle/style:graphic-properties/@style:horizontal-rel"/>
            </xsl:variable>
            <!-- page properties. not valid if more than one page-master-style -->
            <xsl:variable name="pageWidth">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:for-each select="document('styles.xml')">
                    <xsl:value-of select="key('page-layouts', $default-master-style/@style:page-layout-name)[1]/style:page-layout-properties/@fo:page-width" />
                  </xsl:for-each>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="pageLeftMargin">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:for-each select="document('styles.xml')">
                    <xsl:value-of select="key('page-layouts', $default-master-style/@style:page-layout-name)[1]/style:page-layout-properties/@fo:margin-left" />
                  </xsl:for-each>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="pageRightMargin">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:for-each select="document('styles.xml')">
                    <xsl:value-of select="key('page-layouts', $default-master-style/@style:page-layout-name)[1]/style:page-layout-properties/@fo:margin-right" />
                  </xsl:for-each>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <!-- width to be added depending on were object is located : column or page -->
            <xsl:variable name="contextWidth">
              <xsl:call-template name="ComputeContextWidth">
                <xsl:with-param name="frame" select="$frame[1]"/>
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
                  <xsl:with-param name="shapeStyle" select="$frameStyle"/>
                  <xsl:with-param name="attribName">style:wrap</xsl:with-param>
                </xsl:call-template>
              </xsl:variable>
              <xsl:choose>
                <xsl:when test="$wrap = 'none' or $wrap = 'run-through' or $wrap = '' ">
                  <xsl:variable name="precedingFrames">
                    <xsl:call-template name="ComputeContextSubstractedWidth">
                      <xsl:with-param name="frames"
                        select="$frame[1]/preceding-sibling::*[contains(name(),'draw:')]"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:variable name="followingFrames">
                    <xsl:call-template name="ComputeContextSubstractedWidth">
                      <xsl:with-param name="frames"
                        select="$frame[1]/following-sibling::*[contains(name(),'draw:')]"/>
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
                    <xsl:with-param name="shapeStyle" select="$frameStyle"/>
                    <xsl:with-param name="attribName">fo:margin-left</xsl:with-param>
                  </xsl:call-template>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="frameMarginRight">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:call-template name="GetGraphicProperties">
                    <xsl:with-param name="shapeStyle" select="$frameStyle"/>
                    <xsl:with-param name="attribName">fo:margin-right</xsl:with-param>
                  </xsl:call-template>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="frameWidth">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length" select="$frame[1]/@svg:width"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="frameHeight">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length" select="$frame[1]/@svg:height"/>
              </xsl:call-template>
            </xsl:variable>
            <!-- if a value is specified for a frame move -->
            <xsl:variable name="fromLeft">
              <xsl:choose>
                <xsl:when test="$horizontalPos = 'from-left' or $horizontalPos = 'from-inside' ">
                  <xsl:call-template name="point-measure">
                    <xsl:with-param name="length" select="$frame[1]/@svg:x"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <!-- particular transformation -->
            <xsl:variable name="translation">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:value-of select="substring-before(substring-after(substring-after($frame[1]/@draw:transform,'translate'),'('),' ')" />
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
                    <xsl:when test="$frame[1]/@svg:x">
                      <xsl:call-template name="point-measure">
                        <xsl:with-param name="length" select="$frame[1]/@svg:x"/>
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

  
  <!-- 
  Summary:  Computes the width of page or paragraph margin. 
            (This is not correct if document has more than one page-master-style)
  Author:   Clever Age
  Params:   frame:
            horizontalRel:
            horizontalPos:
            pageWidth:
            pageLeftMargin:
            pageRightMargin:
  -->
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

  
  <!-- 
  Summary:  Computes a value to be substracted to context width 
  Author:   Clever Age
  -->
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
              <xsl:call-template name="GetLengthOfFrameSide">
                <xsl:with-param name="side">width</xsl:with-param>
                <xsl:with-param name="frame" select="$frames[1]"/>
                <xsl:with-param name="unit">point</xsl:with-param>
              </xsl:call-template>
            </xsl:when>
            <xsl:when
              test="key('automatic-styles', $frames[1]/@draw:style-name)/style:graphic-properties/@style:wrap = 'left' ">
              <xsl:call-template name="GetLengthOfFrameSide">
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

  
  <!-- 
  Summary:  Computes horizontal margin of a rotated frame. 
  Author:   Clever Age
  -->
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

  
  <!-- 
  Summary:  Computes vertical margin of a draw:frame.
  Author:   Clever Age
  -->
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
                    <!-- page, page-content, paragraph-content, line -->
                    <xsl:when test="$verticalRel = 'page' or $verticalRel = 'page-content' or $verticalRel = 'paragraph' or $verticalRel = 'line'">
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

  
  <!-- 
  Summray:  Compute vertical margin of a rotated frame.
  Author:   Clever Age
  -->
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

  
  <!-- 
  Summary:  Finds graphic property recursively in style or parent style.
  Author:   CleverAge
  -->
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

  
  <!-- 
  Summary:  Compute the padding of a border. 
  Author:   Clever Age
  Params:   frameStyle: The style of the draw:frame
            side: The values 'top', 'right', 'bottom' or 'left'
  -->
  <xsl:template name="GetFramePadding">
    <xsl:param name="frameStyle"/>
    <xsl:param name="side"/>

    <xsl:choose>
      <!-- priority to border padding -->
      <xsl:when test="$frameStyle/style:graphic-properties/@*[name() = concat('fo:padding-', $side)]">
        <xsl:value-of select="$frameStyle/style:graphic-properties/@*[name() = concat('fo:padding-', $side)]"/>
      </xsl:when>
      <!-- otherwise if padding attribute -->
      <xsl:when test="$frameStyle/style:graphic-properties/@fo:padding">
        <xsl:value-of select="$frameStyle/style:graphic-properties/@fo:padding"/>
      </xsl:when>
      <!-- otherwise look parent style -->
      <xsl:when test="$frameStyle/@style:parent-style-name">
        <xsl:variable name="parentStyle" select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $frameStyle/@style:parent-style-name]"/>
        <xsl:call-template name="GetFramePadding">
          <xsl:with-param name="frameStyle" select="$parentStyle"/>
          <xsl:with-param name="side" select="$side"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>0</xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <!-- 
  Summary:  Inserts justification values (text alignment) for paragraphs in a textbox.
            The alignment of draw:textbox is based on the horizontal position.
            (makz: I think this is not correct)
  Author:   Clever Age
  Params:   frame: The draw:frame
  -->
  <xsl:template name="FrameToTextboxAlignment">
    <xsl:param name="frame"/>
    
    <xsl:variable name="hPos">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="key('styles', $frame/@draw:style-name)"/>
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

</xsl:stylesheet>

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
<!--
Modification Log
LogNo. |Date       |ModifiedBy   |BugNo.   |Modification                                                      |
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
RefNo-1 16-Feb-2009 Sandeep S    custom-shape implemetation    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
  xmlns:ooc="urn:odf-converter"                
  exclude-result-prefixes="xlink draw svg fo office style text ooc">

  <xsl:key name="automatic-styles" match="office:automatic-styles/style:style" use="@style:name"/>
  <!--
  *************************************************************************
  SUMMARY
  *************************************************************************
  This stylesheet handles the conversion of shapes. For example shapes are 
  rectangles, ellipses, lines or custom-shapes.
  *************************************************************************
  -->

  <!-- 
  *************************************************************************
  MATCHING TEMPLATES
  *************************************************************************
  -->
  <!-- Sona Shape constants-->
  <xsl:variable name="sm-sm" select="'0.15'"/>
  <xsl:variable name="sm-med" select="'0.18'"/>
  <xsl:variable name="sm-lg" select="'0.2'"/>
  <xsl:variable name="med-sm" select="'0.21'" />
  <xsl:variable name="med-med" select="'0.25'"/>
  <xsl:variable name="med-lg" select="'0.3'" />
  <xsl:variable name="lg-sm" select="'0.31'" />
  <xsl:variable name="lg-med" select="'0.35'" />
  <xsl:variable name="lg-lg" select="'0.4'" />

  <!-- 
  Summary:  Forward shapes in paragraph mode to shapes mode 
  Author:   Clever Age
  -->
  <xsl:template match="draw:custom-shape |draw:rect |draw:ellipse|draw:line|draw:connector" mode="paragraph">
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
  Summary:  Converts a cutom shape
  Author:   Clever Age
  -->
  <xsl:template match="draw:custom-shape" mode="shapes">
    <xsl:variable name="styleName" select=" @draw:style-name"/>
    <xsl:variable name="automaticStyle" select="key('automatic-styles', $styleName)"/>
    <xsl:variable name="officeStyle" select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $styleName]"/>
    <xsl:variable name="shapeStyle" select="$automaticStyle | $officeStyle" />
    <w:r>
      <w:pict>

        <!--
        makz:
        !!! Detecting a custom shape on their draw:type attribute is no clean solution !!!
        !!! It will only work for ODT files generated by OpenOffice !!!
        !!! This needs to be changed in future !!!
        -->
        <xsl:choose>
          <xsl:when test="draw:enhanced-geometry/@draw:type = 'rectangle' ">

            <xsl:call-template name="InsertRect">
              <xsl:with-param name="shapeStyle" select="$automaticStyle | $officeStyle"/>
              <xsl:with-param name="shape" select="."/>
            </xsl:call-template>

          </xsl:when>
          <xsl:when test="draw:enhanced-geometry/@draw:type = 'ellipse' ">

            <xsl:call-template name="InsertOval">
              <xsl:with-param name="shapeStyle" select="$automaticStyle | $officeStyle"/>
              <xsl:with-param name="shape" select="."/>
            </xsl:call-template>

          </xsl:when>
          <!--#############Code changes done by Pradeep#######################-->
          <!-- TODO - other shapes -->
          <xsl:otherwise>
            <!--Start of RefNo-1-->
            <xsl:variable name="shapeTypeID" select="generate-id(.)"/>

            <v:shape id="{concat('_x0000_s',$shapeTypeID)}" >
              <xsl:attribute name="coordsize">
                <xsl:variable name="svgViewBox">
                  <xsl:choose>
                    <xsl:when test="./draw:enhanced-geometry/@svg:viewBox">
                      <xsl:value-of select="substring-after(./draw:enhanced-geometry/@svg:viewBox,' ')"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="'0 21600 21600'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:variable>
                <xsl:value-of select="translate(substring-after($svgViewBox,' '),' ',',')"/>
              </xsl:attribute>
              <xsl:if test="./draw:enhanced-geometry/@draw:modifiers">
                <xsl:attribute name="adj">
                  <xsl:value-of select="translate(./draw:enhanced-geometry/@draw:modifiers,' ',',')"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:variable name="viewBox">
                <xsl:choose>
                  <xsl:when test="./draw:enhanced-geometry/@svg:viewBox">
                    <xsl:value-of select="./draw:enhanced-geometry/@svg:viewBox"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'0 0 21600 21600'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>
              <xsl:variable name="modifiers" select="./draw:enhanced-geometry/@draw:modifiers"/>
              <xsl:variable name="drawEqn">
                <xsl:for-each select="./draw:enhanced-geometry/draw:equation">
                  <xsl:value-of select="concat('|',@draw:formula)"/>
                </xsl:for-each>
              </xsl:variable>
              <xsl:variable name="enhPath" select="./draw:enhanced-geometry/@draw:enhanced-path"/>

              <xsl:if test="./draw:enhanced-geometry/@draw:enhanced-path">
                <xsl:attribute name="path">
                  <xsl:value-of select="concat('CustShpWrdFreFrm',$modifiers,'###',substring($drawEqn,2),'###',$viewBox,'###',$enhPath)"/>
                </xsl:attribute>
                <!--<xsl:value-of select="concat('CustShpWrdFreFrm',$modifiers,'###',substring($drawEqn,2),'###',$viewBox,'###',$enhPath)"/>-->
              </xsl:if>

              <xsl:call-template name="ConvertShapeProperties">
                <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
                <!-- Sona: Changed I/P parameter-->
                <xsl:with-param name="shape" select="."/>
              </xsl:call-template>
              <!-- Sona:Defect #2020254-->
              <xsl:if test="./text:p/node()">
                <!--insert text-box-->
                <xsl:call-template name="InsertTextBox">
                  <!-- Sona: Changed I/P parameter-->
                  <xsl:with-param name="frameStyle" select="$shapeStyle"/>
                </xsl:call-template>
              </xsl:if>
            </v:shape>
            <!--End of RefNo-1-->
          </xsl:otherwise>
        </xsl:choose>
      </w:pict>
    </w:r>
  </xsl:template>


  <!-- 
  Summary:  Converts draw:rect to VML rectangle
  Modified: makz (DIaLOGIKa)
  -->
  <xsl:template match="draw:rect" mode="shapes">
    <xsl:variable name="styleName" select=" @draw:style-name"/>
    <xsl:variable name="automaticStyle" select="key('automatic-styles', $styleName)"/>
    <xsl:variable name="officeStyle" select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $styleName]"/>

    <w:r>
      <w:pict>
        <xsl:call-template name="InsertRect">
          <xsl:with-param name="shapeStyle" select="$automaticStyle | $officeStyle"/>
          <xsl:with-param name="shape" select="."/>
        </xsl:call-template>
      </w:pict>
    </w:r>
  </xsl:template>


  <!--
  Summary:  Converts draw:ellipse to VML oval
  Author:   makz (DIaLOGIKa)
  -->
  <xsl:template match="draw:ellipse" mode="shapes">
    <xsl:variable name="styleName" select=" @draw:style-name"/>
    <xsl:variable name="automaticStyle" select="key('automatic-styles', $styleName)"/>
    <xsl:variable name="officeStyle" select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $styleName]"/>

    <w:r>
      <w:pict>
        <xsl:call-template name="InsertOval">
          <xsl:with-param name="shapeStyle" select="$automaticStyle | $officeStyle"/>
          <xsl:with-param name="shape" select="."/>
        </xsl:call-template>
      </w:pict>
    </w:r>
  </xsl:template>
  
  
  <!-- Sona: Defect #2019374-->
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
        <v:shapetype id="_x0000_t202" coordsize="21600,21600" path="m,l,21600r21600,l21600,xe" xmlns:o="urn:schemas-microsoft-com:office:office">
          <v:stroke joinstyle="miter"/>
          <v:path gradientshapeok="t" o:connecttype="rect"/>
        </v:shapetype>
        <!-- Sona: Also fixed defect #2025700-->
        <v:shape id="_x0000_s1026" type="#_x0000_t202">
          <xsl:variable name="styleName" select="parent::draw:frame/@draw:style-name"/>
          <xsl:variable name="automaticStyle" select="key('automatic-styles', $styleName)"/>
          <xsl:variable name="officeStyle" select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $styleName]"/>
          <xsl:variable name="shapeStyle" select="$automaticStyle | $officeStyle"/>

          <xsl:call-template name="ConvertShapeProperties">
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
            <xsl:with-param name="shape" select="parent::draw:frame"/>
          </xsl:call-template>

          <xsl:call-template name="InsertTextBox">
            <xsl:with-param name="frameStyle" select="$shapeStyle"/>
            <xsl:with-param name="frame" select="parent::draw:frame"/>
          </xsl:call-template>
        </v:shape>
      </w:pict>
    </w:r>
  </xsl:template>

  <!-- Sona: Line Shapes-->
  <xsl:template match="draw:line" mode="shapes">
    <xsl:variable name="styleName" select=" @draw:style-name"/>
    <xsl:variable name="automaticStyle" select="key('automatic-styles', $styleName)"/>
    <xsl:variable name="officeStyle" select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $styleName]"/>

    <w:r>
      <w:pict>
        <v:shapetype id="_x0000_t32" coordsize="21600,21600" o:spt="32" o:oned="t" path="m,l21600,21600e" filled="f">
          <v:path arrowok="t" fillok="f" o:connecttype="none"/>
          <o:lock v:ext="edit" shapetype="t"/>
        </v:shapetype>
        <v:shape id="_x0000_s1026" type="#_x0000_t32">
          <!--added by chhavi for alttext-->
          <xsl:if test="svg:desc">
            <xsl:attribute name="alt">
              <xsl:value-of select="svg:desc/child::node()"/>
            </xsl:attribute>
          </xsl:if>
          <!--end here-->
          <xsl:attribute name="style">
            <xsl:call-template name="GetLineCoordinatesODF">
              <xsl:with-param name="shape" select="."/>
            </xsl:call-template>
            <!-- Sona : Defect #2026780-->
            <!-- z-Index-->
            <xsl:call-template name="InsertShapeZIndex">
              <xsl:with-param name="shapeStyle" select="$automaticStyle | $officeStyle"/>
              <xsl:with-param name="shape" select="."/>
            </xsl:call-template>

            <!--Sona Added margin for completing Shape Wrap feature-->
            <xsl:call-template name="FrameToShapeMargin">
              <xsl:with-param name="frameStyle" select="$automaticStyle | $officeStyle"/>
              <xsl:with-param name="frame" select="."/>
            </xsl:call-template>
          </xsl:attribute>
          <xsl:attribute name="o:connectortype">
            <xsl:value-of select="'straight'"/>
          </xsl:attribute>
          <xsl:call-template name="InsertShapeStroke">
            <xsl:with-param name="shapeStyle" select="$automaticStyle|$officeStyle"/>
          </xsl:call-template>
          <!-- Sona Added Dashed Lines-->
          <xsl:call-template name="GetLineStroke">
            <xsl:with-param name="shapeStyle" select="$automaticStyle|$officeStyle"/>
          </xsl:call-template>
          <!--Sona Added Shape Wrap-->
          <xsl:call-template name="FrameToShapeWrap">
            <xsl:with-param name="frameStyle" select="$automaticStyle|$officeStyle"/>
          </xsl:call-template>
          <!--Sona Added Shape shadow-->
          <xsl:call-template name="FrameToShapeShadow">
            <xsl:with-param name="frameStyle" select="$automaticStyle|$officeStyle"/>
          </xsl:call-template>
        </v:shape>
      </w:pict>
    </w:r>
  </xsl:template>
  <xsl:template match="draw:connector" mode="shapes">
    <xsl:variable name="styleName" select=" @draw:style-name"/>
    <xsl:variable name="automaticStyle" select="key('automatic-styles', $styleName)"/>
    <xsl:variable name="officeStyle" select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $styleName]"/>
    <w:r>
      <w:pict>
        <xsl:choose>
          <xsl:when test="@draw:type='curve'">
            <v:shapetype id="_x0000_t38" coordsize="21600,21600" o:spt="38" o:oned="t" path="m,c@0,0@1,5400@1,10800@1,16200@2,21600,21600,21600e" filled="f">
              <v:formulas>
                <v:f eqn="mid #0 0"/>
                <v:f eqn="val #0"/>
                <v:f eqn="mid #0 21600"/>
              </v:formulas>
              <v:path arrowok="t" fillok="f" o:connecttype="none"/>
              <v:handles>
                <v:h position="#0,center"/>
              </v:handles>
              <o:lock v:ext="edit" shapetype="t"/>
            </v:shapetype>
            <!-- Sona: Defect 2019239-->
            <v:shape id="_x0000_s1036" type="#_x0000_t38" adj="10800,22173,-43579">
              <xsl:attribute name="style">
                <xsl:call-template name="GetLineCoordinatesODF">
                  <xsl:with-param name="shape" select="."/>
                </xsl:call-template>
                <!-- Sona : Defect #2026780-->
                <!-- z-Index-->
                <xsl:call-template name="InsertShapeZIndex">
                  <xsl:with-param name="shapeStyle" select="$automaticStyle | $officeStyle"/>
                  <xsl:with-param name="shape" select="."/>
                </xsl:call-template>

                <!--Sona Added margin for completing Shape Wrap feature-->
                <xsl:call-template name="FrameToShapeMargin">
                  <xsl:with-param name="frameStyle" select="$automaticStyle | $officeStyle"/>
                  <xsl:with-param name="frame" select="."/>
                </xsl:call-template>
              </xsl:attribute>
              <xsl:attribute name="o:connectortype">
                <xsl:value-of select="'curved'"/>
              </xsl:attribute>
              <xsl:call-template name="InsertShapeStroke">
                <xsl:with-param name="shapeStyle" select="$automaticStyle|$officeStyle"/>
              </xsl:call-template>
              <!-- Sona Added Dashed Lines-->
              <xsl:call-template name="GetLineStroke">
                <xsl:with-param name="shapeStyle" select="$automaticStyle|$officeStyle"/>
              </xsl:call-template>
              <!--Sona Added Shape Wrap-->
              <xsl:call-template name="FrameToShapeWrap">
                <xsl:with-param name="frameStyle" select="$automaticStyle|$officeStyle"/>
              </xsl:call-template>
              <!--Sona Added Shape shadow-->
              <xsl:call-template name="FrameToShapeShadow">
                <xsl:with-param name="frameStyle" select="$automaticStyle|$officeStyle"/>
              </xsl:call-template>
            </v:shape>
          </xsl:when>
          <!-- Sona: Defect #2019464-->
          <xsl:when test="@draw:type='line'">
            <v:shapetype id="_x0000_t32" coordsize="21600,21600" o:spt="32" o:oned="t" path="m,l21600,21600e" filled="f">
              <v:path arrowok="t" fillok="f" o:connecttype="none"/>
              <o:lock v:ext="edit" shapetype="t"/>
            </v:shapetype>
            <v:shape id="_x0000_s1026" type="#_x0000_t32">
              <xsl:attribute name="style">
                <xsl:call-template name="GetLineCoordinatesODF">
                  <xsl:with-param name="shape" select="."/>
                </xsl:call-template>
                <!-- Sona : Defect #2026780-->
                <!-- z-Index-->
                <xsl:call-template name="InsertShapeZIndex">
                  <xsl:with-param name="shapeStyle" select="$automaticStyle | $officeStyle"/>
                  <xsl:with-param name="shape" select="."/>
                </xsl:call-template>

                <!--Sona Added margin for completing Shape Wrap feature-->
                <xsl:call-template name="FrameToShapeMargin">
                  <xsl:with-param name="frameStyle" select="$automaticStyle | $officeStyle"/>
                  <xsl:with-param name="frame" select="."/>
                </xsl:call-template>
              </xsl:attribute>
              <xsl:attribute name="o:connectortype">
                <xsl:value-of select="'straight'"/>
              </xsl:attribute>
              <xsl:call-template name="InsertShapeStroke">
                <xsl:with-param name="shapeStyle" select="$automaticStyle|$officeStyle"/>
              </xsl:call-template>
              <!-- Sona Added Dashed Lines-->
              <xsl:call-template name="GetLineStroke">
                <xsl:with-param name="shapeStyle" select="$automaticStyle|$officeStyle"/>
              </xsl:call-template>
              <!--Sona Added Shape Wrap-->
              <xsl:call-template name="FrameToShapeWrap">
                <xsl:with-param name="frameStyle" select="$automaticStyle|$officeStyle"/>
              </xsl:call-template>
              <!--Sona Added Shape shadow-->
              <xsl:call-template name="FrameToShapeShadow">
                <xsl:with-param name="frameStyle" select="$automaticStyle|$officeStyle"/>
              </xsl:call-template>
            </v:shape>
          </xsl:when>
          <!-- Sona: Defect #2019464-->
          <xsl:otherwise>
            <v:shapetype id="_x0000_t34" coordsize="21600,21600" o:spt="34" o:oned="t" adj="10800" path="m,l@0,0@0,21600,21600,21600e" filled="f">
              <v:stroke joinstyle="miter"/>
              <v:formulas>
                <v:f eqn="val #0"/>
              </v:formulas>
              <v:path arrowok="t" fillok="f" o:connecttype="none"/>
              <v:handles>
                <v:h position="#0,center"/>
              </v:handles>
              <o:lock v:ext="edit" shapetype="t"/>
            </v:shapetype>
            <v:shape id="_x0000_s1033" type="#_x0000_t34" adj=",-249300,-14316">
              <xsl:attribute name="style">
                <xsl:call-template name="GetLineCoordinatesODF">
                  <xsl:with-param name="shape" select="."/>
                </xsl:call-template>
                <!-- Sona : Defect #2026780-->
                <!-- z-Index-->
                <xsl:call-template name="InsertShapeZIndex">
                  <xsl:with-param name="shapeStyle" select="$automaticStyle | $officeStyle"/>
                  <xsl:with-param name="shape" select="."/>
                </xsl:call-template>

                <!--Sona Added margin for completing Shape Wrap feature-->
                <xsl:call-template name="FrameToShapeMargin">
                  <xsl:with-param name="frameStyle" select="$automaticStyle | $officeStyle"/>
                  <xsl:with-param name="frame" select="."/>
                </xsl:call-template>
              </xsl:attribute>
              <xsl:attribute name="o:connectortype">
                <xsl:value-of select="'elbow'"/>
              </xsl:attribute>
              <xsl:call-template name="InsertShapeStroke">
                <xsl:with-param name="shapeStyle" select="$automaticStyle|$officeStyle"/>
              </xsl:call-template>
              <!-- Sona Added Dashed Lines-->
              <xsl:call-template name="GetLineStroke">
                <xsl:with-param name="shapeStyle" select="$automaticStyle|$officeStyle"/>
              </xsl:call-template>
              <!--Sona Added Shape Wrap-->
              <xsl:call-template name="FrameToShapeWrap">
                <xsl:with-param name="frameStyle" select="$automaticStyle|$officeStyle"/>
              </xsl:call-template>
              <!--Sona Added Shape shadow-->
              <xsl:call-template name="FrameToShapeShadow">
                <xsl:with-param name="frameStyle" select="$automaticStyle|$officeStyle"/>
              </xsl:call-template>
            </v:shape>
          </xsl:otherwise>
        </xsl:choose>
      </w:pict>
    </w:r>
  </xsl:template>
  <!-- 
  *************************************************************************
  CALLED TEMPLATES
  *************************************************************************
  -->


  <!--
  Summary:  Inserts a VML rectangle
  Author:   makz (DIaLOGIKa)
  Params:   shapeStyle: The automatic style of the draw:shape
            shape: The draw:shape itself
  -->
  <xsl:template name="InsertRect">
    <xsl:param name="shapeStyle"/>
    <xsl:param name="shape" />

    <v:rect id="_x0000_s1026">

      <xsl:call-template name="ConvertShapeProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="shape" select="$shape"/>
      </xsl:call-template>
      <!-- Sona:Defect #2020254-->
      <xsl:if test="$shape/text:p/node()">
        <!-- reuse the frame template, attributes are the same -->
        <xsl:call-template name="InsertTextBox">
          <xsl:with-param name="frameStyle" select="$shapeStyle"/>
        </xsl:call-template>
      </xsl:if>
    </v:rect>
  </xsl:template>


  <!--
  Summary:  Inserts a VML rounded rectangle
  Author:   makz (DIaLOGIKa)
  Params:   shapeStyle: The automatic style of the draw:shape
            shape: The draw:shape itself
  -->
  <xsl:template name="InsertRoundedRect">
    <xsl:param name="shapeStyle"/>
    <xsl:param name="shape" />

    <v:roundrect id="_x0000_s1032">

      <xsl:call-template name="ConvertShapeProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="shape" select="$shape"/>
      </xsl:call-template>

      <!-- Sona:Defect #2020254-->
      <xsl:if test="$shape/text:p/node()">
        <!-- reuse the frame template, attributes are the same -->
        <xsl:call-template name="InsertTextBox">
          <xsl:with-param name="frameStyle" select="$shapeStyle"/>
        </xsl:call-template>
      </xsl:if>
    </v:roundrect>
  </xsl:template>


  <!--
  Summary:  Inserts a VML oval
  Author:   makz (DIaLOGIKa)
  Params:   shapeStyle: The automatic style of the draw:shape
            shape: The draw:shape itself
  -->
  <xsl:template name="InsertOval">
    <xsl:param name="shapeStyle"/>
    <xsl:param name="shape" />

    <v:oval id="_x0000_s1037">

      <xsl:call-template name="ConvertShapeProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="shape" select="$shape"/>
      </xsl:call-template>
      <!-- Sona:Defect #2020254-->
      <xsl:if test="$shape/text:p/node()">
        <!-- reuse the frame template, attributes are the same -->
        <xsl:call-template name="InsertTextBox">
          <xsl:with-param name="frameStyle" select="$shapeStyle"/>
        </xsl:call-template>
      </xsl:if>
    </v:oval>
  </xsl:template>


  <!-- 
  Summary:  Converts the properties of a draw:rect/draw:oval (e.g.) to VML properties
  Author:   CleverAge
  Params:   shapeStyle: The automatic style of the draw:shape
            shape: The draw:shape itself
  -->
  <xsl:template name="ConvertShapeProperties">
    <xsl:param name="shapeStyle"/>
    <xsl:param name="shape"/>
    <!--added by chhavi for alttext-->
    <xsl:if test="svg:desc">
      <xsl:attribute name="alt">
        <xsl:value-of select="svg:desc/child::node()"/>
      </xsl:attribute>
    </xsl:if>
    <!--end here-->
    <xsl:if test="$shapeStyle != 0 or count($shapeStyle) &gt; 1">

      <xsl:call-template name="InsertShapeStyleAttribute">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="shape" select="$shape"/>
      </xsl:call-template>

      <xsl:call-template name="InsertShapeStroke">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

      <xsl:call-template name="InsertShapeFill">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="shape" select="$shape"/>
      </xsl:call-template>

      <!-- Sona Added Dashed Lines-->
      <xsl:call-template name="GetLineStroke">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

      <!--Sona Added Shape Wrap-->
      <xsl:call-template name="FrameToShapeWrap">
        <xsl:with-param name="frameStyle" select="$shapeStyle"/>
      </xsl:call-template>
      <!--Sona Added Shape shadow-->
      <xsl:call-template name="FrameToShapeShadow">
        <xsl:with-param name="frameStyle" select="$shapeStyle"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>


  <!-- 
  Summary:  Inserts the style attribute of a VML shape
  Author:   CleverAge
  Params:   shapeStyle: The automatic style of the draw:shape
            shape: The draw:shape itself
  -->
  <xsl:template name="InsertShapeStyleAttribute">
    <xsl:param name="shapeStyle"/>
    <xsl:param name="shape"/>

    <xsl:attribute name="style">

      <!-- width: -->
      <xsl:variable name="frameW">

        <!-- Sona: Defect #2019374-->
        <xsl:choose>
          <xsl:when test="$shape/@svg:width">
            <!-- Sona: Defect #2166160-->
            <xsl:value-of select="ooc:PtFromMeasuredUnit($shape/@svg:width, 0)" />
          </xsl:when>
          <xsl:otherwise>
            <!-- Sona: Defect #2166160-->
            <xsl:value-of select="ooc:PtFromMeasuredUnit($shape/@fo:min-width|$shapeStyle/style:graphic-properties/@fo:min-width, 0)" />
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:value-of select="concat('width:',$frameW,'pt;')"/>

      <!-- height: -->
      <xsl:variable name="frameH">
        <xsl:choose>
          <xsl:when test="$shape/@svg:height">
            <!-- Sona: Defect #2166160-->
            <!-- Sona: Defect #2019374-->
            <xsl:value-of select="ooc:PtFromMeasuredUnit($shape/@svg:height, 0)" />
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="ooc:PtFromMeasuredUnit($shape/child::node()/@fo:min-height|$shapeStyle/style:graphic-properties/@fo:min-height, 0)" />
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:value-of select="concat('height:',$frameH,'pt;')"/>

      <!--code added by Chhavi for relative size in Frame - Defect #2166036-->
      <xsl:if test="(../self::draw:frame or self::draw:frame) and string(number(substring-before($shape/@style:rel-width,'%'))) != 'NaN'">
        <xsl:value-of select="concat('mso-width-percent:', substring-before($shape/@style:rel-width,'%') * 10,';')"/>
      </xsl:if>
      <xsl:if test="(../self::draw:frame or self::draw:frame) and string(number(substring-before($shape/@style:rel-height,'%'))) != 'NaN'">
        <xsl:value-of select="concat('mso-height-percent:', substring-before($shape/@style:rel-height,'%') * 10,';')"/>
      </xsl:if>

      <!-- z-Index-->
      <xsl:call-template name="InsertShapeZIndex">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="shape" select="$shape"/>
      </xsl:call-template>
      <!-- Vipul:2645455  flip without rotation-->
      <xsl:choose>
        <xsl:when test="$shape/draw:enhanced-geometry/@draw:mirror-horizontal='true'">
          <xsl:value-of select="'flip:x;'"/>
        </xsl:when>
        <xsl:when test="$shape/draw:enhanced-geometry/@draw:mirror-vertical='true'">
          <xsl:value-of select="'flip:y;'"/>
        </xsl:when>
      </xsl:choose>
      <!-- reuse the frame template, attributes are the same -->
      <xsl:call-template name="FrameToRelativeShapePosition">
        <xsl:with-param name="frameStyle" select="$shapeStyle"/>
        <xsl:with-param name="frame" select="$shape"/>
      </xsl:call-template>

      <!-- reuse the frame template, attributes are the same -->
      <xsl:call-template name="FrameToShapePosition">
        <xsl:with-param name="frameStyle" select="$shapeStyle"/>
        <xsl:with-param name="frame" select="$shape"/>
      </xsl:call-template>

      <!-- reuse the frame template, attributes are the same -->
      <xsl:call-template name="FrameToTextAnchor">
        <xsl:with-param name="frameStyle" select="$shapeStyle"/>
      </xsl:call-template>

      <!--Sona Added margin for completing Shape Wrap feature-->
      <xsl:call-template name="FrameToShapeMargin">
        <xsl:with-param name="frameStyle" select="$shapeStyle"/>
        <xsl:with-param name="frame" select="$shape"/>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>


  <!-- 
  Summary:  Inserts the VML shape z-index
  Author:   Sona 
  -->
  <xsl:template name="InsertShapeZIndex">
    <xsl:param name="shapeStyle"></xsl:param>
    <xsl:param name="shape"></xsl:param>
    <!-- z-index: -->
    <!-- Sona: Defect #2026780-->
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
    <xsl:choose>
      <xsl:when test="$frameWrap='run-through' and $runThrought='background'">
        <xsl:value-of select="concat('z-index:',-251658240,';')"/>
      </xsl:when>
      <xsl:when test="$frameWrap='run-through' and not($runThrought)">
        <xsl:value-of select="concat('z-index:',251658240,';')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="not($shape/@draw:z-index)">
            <xsl:value-of select="concat('z-index:',251659264,';')"/>
          </xsl:when>
          <!-- Sona: Defect #2019374-->
          <xsl:when test="$shape/@draw:z-index=0">
            <xsl:value-of select="concat('z-index:',2516572155-$shape/@draw:z-index,';')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat('z-index:',251659264+$shape/@draw:z-index,';')"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- 
  Summary:  Inserts the VML shape fill
  Author:   CleverAge
  Params:   shapeStyle: The automatic style of the draw:shape
  -->
  <xsl:template name="InsertShapeFill">
    <xsl:param name="shapeStyle"/>
    <xsl:param name="shape"/>

    <xsl:variable name="fillColor">
      <xsl:choose>
        <xsl:when test="not($shapeStyle/style:graphic-properties/@draw:fill-color)">
          <xsl:call-template name="GetDrawnGraphicProperties">
            <xsl:with-param name="attrib">fo:background-color</xsl:with-param>
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="GetDrawnGraphicProperties">
            <xsl:with-param name="attrib">draw:fill-color</xsl:with-param>
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="fillProperty">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">draw:fill</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
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
    <!--Sonata:SP2 defect:Scenario:docx ->SP2->odt->3.0->docx-"no fill" property of text box lost:Added attribute filled="f" incase of transparent fill-->
    <xsl:if test="$fillColor = 'transparent' ">
      <xsl:attribute name="filled">
        <xsl:value-of select="'f'"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="$fillColor != '' and $fillColor != 'transparent'">
      <xsl:attribute name="fillcolor">
        <xsl:value-of select="$fillColor"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="$fillProperty != 'none' ">
      <v:fill>
        <xsl:if test="$opacity != '' ">
          <xsl:attribute name="opacity">
            <!-- Sonata: shape transperancy -->
            <xsl:value-of select="number(substring-before($opacity,'%')) div 100"/>
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
          <!-- Sonata: shape Picture Fill -->
          <xsl:when test="$fillProperty = 'bitmap' ">
            <xsl:variable name="BitmapName">
              <xsl:call-template name="GetGraphicProperties">
                <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
                <xsl:with-param name="attribName">draw:fill-image-name</xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="stretch" select="$shapeStyle/style:graphic-properties/@style:repeat"/>

            <xsl:for-each select="document('styles.xml')//draw:fill-image[@draw:name = $BitmapName]">
              <!-- radial gradients not handled yet -->
              <xsl:choose>
                <xsl:when test="$stretch='stretch'">
                  <xsl:attribute name="type">frame</xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:attribute name="type">tile</xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
              <xsl:attribute name="recolor">t</xsl:attribute>
              <xsl:attribute name="color2">black</xsl:attribute>
              <xsl:attribute name="r:id">
                <xsl:value-of select="concat('Bitmap_',generate-id($shape))"/>
              </xsl:attribute>
            </xsl:for-each>
          </xsl:when>
          <xsl:when test="$fillProperty = 'gradient' ">
            <!-- simple linear gradient -->
            <xsl:variable name="gradientName">
              <xsl:call-template name="GetGraphicProperties">
                <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
                <xsl:with-param name="attribName">draw:fill-gradient-name</xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/draw:gradient[@draw:name = $gradientName]">
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
          </xsl:when>
          <!-- Sona: Picture fill for frame-->
          <xsl:when test="$fillProperty = '' and (parent::node()[name()='draw:frame'] or self::node()[name()='draw:frame'])">
            <xsl:if test="$shapeStyle/style:graphic-properties/style:background-image/@*">
              <xsl:variable name="stretch" select="$shapeStyle/style:graphic-properties/style:background-image/@style:repeat"/>

              <xsl:choose>
                <xsl:when test="$stretch='stretch'">
                  <xsl:attribute name="type">frame</xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:attribute name="type">tile</xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
              <!--<xsl:attribute name="recolor">t</xsl:attribute>-->
              <xsl:attribute name="color2">black</xsl:attribute>
              <xsl:attribute name="r:id">
                <xsl:value-of select="concat('Bitmap_',generate-id($shape))"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </v:fill>
    </xsl:if>
    <!--chhavi filled color-->
    <xsl:if test="$fillProperty = 'none' ">
      <xsl:attribute name="filled">
        <xsl:value-of select="'f'"/>
      </xsl:attribute>
    </xsl:if>
    <!--end here-->
  </xsl:template>


  <!-- 
  Summary:  Inserts the VML stroke of a shape
  Author:   CleverAge
  Params:   shapeStyle: The automatic style of the draw:shape
  -->
  <xsl:template name="InsertShapeStroke">
    <xsl:param name="shapeStyle"/>

    <xsl:variable name="strokeColor">
      <xsl:call-template name="GetDrawnGraphicProperties">
        <xsl:with-param name="attrib">
          <xsl:choose>
            <xsl:when test="name(parent::node()) = 'draw:frame' and not($shapeStyle/style:graphic-properties/@svg:stroke-color)">
              <xsl:value-of select="'fo:border'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'svg:stroke-color'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="strokeWeight">
      <xsl:call-template name="GetDrawnGraphicProperties">
        <!--code changed by yeswanth.s : to get width of rect & frame-->
        <!--<xsl:with-param name="attrib">svg:stroke-width</xsl:with-param>-->
        <xsl:with-param name="attrib">
          <xsl:choose>
            <xsl:when test="name(parent::node()) = 'draw:frame' and not($shapeStyle/style:graphic-properties/@svg:stroke-width)">
              <xsl:value-of select="'fo:border'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'svg:stroke-width'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>
    </xsl:variable>

    <!-- Sona Bug fix 1835088 -->
    <xsl:variable name="shapeBorder">
      <xsl:call-template name="GetDrawnGraphicProperties">
        <!--changes made by yeswanth.s-->
        <xsl:with-param name="attrib">
          <xsl:choose>
            <xsl:when test="name(parent::node()) = 'draw:frame' and not($shapeStyle/style:graphic-properties/@draw:stroke)">
              <xsl:value-of select="'fo:border'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'draw:stroke'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
        <!--end-->
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>
    </xsl:variable>

    <!--code changed by yeswanth.s : For Stroke Weight-->
    <xsl:choose>
      <xsl:when test="$shapeBorder='none' or ($shapeBorder = '' and (name(parent::node()) = 'draw:frame' or self::node()[name()='draw:frame']))">
        <xsl:attribute name="stroked">
          <xsl:value-of select="'f'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="$strokeColor != '' and substring-after($strokeColor,'#')!=''">
          <xsl:attribute name="strokecolor">
            <xsl:value-of select="concat('#',substring-after($strokeColor,'#'))"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="$strokeWeight != '' ">
          <xsl:attribute name="strokeweight">
            <xsl:call-template name="point-measure">
              <xsl:with-param name="length" select="$strokeWeight"/>
              <!--yeswanth.s added parameter : 14-Oct-08-->
              <xsl:with-param name="round" select="'false'"/>
            </xsl:call-template>
            <xsl:text>pt</xsl:text>
          </xsl:attribute>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>


  </xsl:template>


  <!-- 
  Summary:  Gets graphic properties for shapes 
  Author:   CleverAge
  -->
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

  <!-- Sona Line Functions-->
  <xsl:template name="GetLineCoordinatesODF">
    <xsl:param name="shape" select="."/>
    <xsl:variable name="x1">
      <xsl:choose>
        <xsl:when test="@svg:x1=0">
          <xsl:value-of select="'0pt'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="@svg:x1"/>
            <xsl:with-param name="unit" select="'point'"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="y1">
      <xsl:choose>
        <xsl:when test="@svg:y1=0">
          <xsl:value-of select="'0pt'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="@svg:y1"/>
            <xsl:with-param name="unit" select="'point'"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="x2">
      <xsl:choose>
        <xsl:when test="@svg:x2=0">
          <xsl:value-of select="'0pt'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="@svg:x2"/>
            <xsl:with-param name="unit" select="'point'"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="y2">
      <xsl:choose>
        <xsl:when test="@svg:y2=0">
          <xsl:value-of select="'0pt'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="@svg:y2"/>
            <xsl:with-param name="unit" select="'point'"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:text>position:absolute;</xsl:text>
    <xsl:choose>
      <xsl:when test="$x2 &gt;= $x1 and $y1 &gt; $y2">
        <xsl:text>margin-left:</xsl:text>
        <xsl:choose>
          <xsl:when test="number($x1) and $x1 != 0">
            <xsl:value-of select="$x1"/>
            <xsl:text>pt</xsl:text>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
        <xsl:text>;</xsl:text>
        <!--<xsl:variable name="margin-top">
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="$y2"/>
            <xsl:with-param name="unit" select="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
        <xsl:text>margin-top:</xsl:text>
        <xsl:choose>
          <xsl:when test="number($y2) and $y2 != 0">
            <xsl:value-of select="$y2"/>
            <xsl:text>pt</xsl:text>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
        <xsl:text>;</xsl:text>
        <xsl:text>width:</xsl:text>
        <xsl:choose>
          <xsl:when test="number($x2 - $x1) and ($x2 - $x1) != 0">
            <xsl:value-of select="($x2 - $x1)"/>
            <xsl:text>pt</xsl:text>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
        <xsl:text>;</xsl:text>
        <!--<xsl:variable name="height">
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="concat(substring-before($y1,'in')-substring-before($y2,'in'),'in')"/>
            <xsl:with-param name="unit" select="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
        <xsl:text>height:</xsl:text>
        <xsl:choose>
          <xsl:when test="number($y1 - $y2) and $y1 - $y2 != 0">
            <xsl:value-of select="$y1 - $y2"/>
            <xsl:text>pt</xsl:text>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
        <xsl:text>;</xsl:text>
        <xsl:text>flip:y;</xsl:text>
      </xsl:when>
      <xsl:when test="$x1 &gt; $x2 and $y2 &gt;= $y1">
        <!--<xsl:variable name="margin-left">
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="$x2"/>
            <xsl:with-param name="unit" select="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
        <xsl:text>margin-left:</xsl:text>
        <xsl:choose>
          <xsl:when test="number($x2) and $x2 != 0">
            <xsl:value-of select="$x2"/>
            <xsl:text>pt</xsl:text>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
        <xsl:text>;</xsl:text>
        <!--<xsl:variable name="margin-top">
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="$y1"/>
            <xsl:with-param name="unit" select="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
        <xsl:text>margin-top:</xsl:text>
        <xsl:choose>
          <xsl:when test="number($y1) and $y1 != 0">
            <xsl:value-of select="$y1"/>
            <xsl:text>pt</xsl:text>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
        <xsl:text>;</xsl:text>
        <!--<xsl:variable name="width">
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="concat(substring-before($x1,'in')-substring-before($x2,'in'),'in')"/>
            <xsl:with-param name="unit" select="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
        <xsl:text>width:</xsl:text>
        <xsl:choose>
          <xsl:when test="number($x1 - $x2) and ($x1 - $x2) != 0">
            <xsl:value-of select="($x1 - $x2)"/>
            <xsl:text>pt</xsl:text>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
        <xsl:text>;</xsl:text>
        <!--<xsl:variable name="height">
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="concat(substring-before($y2,'in')-substring-before($y1,'in'),'in')"/>
            <xsl:with-param name="unit" select="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
        <xsl:text>height:</xsl:text>
        <xsl:choose>
          <xsl:when test="number($y2 - $y1) and ($y2 - $y1) != 0">
            <xsl:value-of select="($y2 - $y1)"/>
            <xsl:text>pt</xsl:text>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
        <xsl:text>;</xsl:text>
        <xsl:text>flip:x;</xsl:text>

      </xsl:when>
      <xsl:when test="$x1 &gt; $x2 and $y1 &gt; $y2">
        <!--<xsl:variable name="margin-left">
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="$x2"/>
            <xsl:with-param name="unit" select="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
        <xsl:text>margin-left:</xsl:text>
        <xsl:choose>
          <xsl:when test="number($x2) and $x2 != 0">
            <xsl:value-of select="$x2"/>
            <xsl:text>pt</xsl:text>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
        <xsl:text>;</xsl:text>
        <!--<xsl:variable name="margin-top">
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="$y2"/>
            <xsl:with-param name="unit" select="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
        <xsl:text>margin-top:</xsl:text>
        <xsl:choose>
          <xsl:when test="number($y2) and $y2 != 0">
            <xsl:value-of select="$y2"/>
            <xsl:text>pt</xsl:text>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
        <xsl:text>;</xsl:text>
        <!--<xsl:variable name="width">
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="concat(substring-before($x1,'in')-substring-before($x2,'in'),'in')"/>
            <xsl:with-param name="unit" select="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
        <xsl:text>width:</xsl:text>
        <xsl:choose>
          <xsl:when test="number($x1 - $x2) and ($x1 - $x2) != 0">
            <xsl:value-of select="($x1 - $x2)"/>
            <xsl:text>pt</xsl:text>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
        <xsl:text>;</xsl:text>
        <!--<xsl:variable name="height">
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="concat(substring-before($y1,'in')-substring-before($y2,'in'),'in')"/>
            <xsl:with-param name="unit" select="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
        <xsl:text>height:</xsl:text>
        <xsl:choose>
          <xsl:when test="number($y1 - $y2) and ($y1 - $y2) != 0">
            <xsl:value-of select="($y1 - $y2)"/>
            <xsl:text>pt</xsl:text>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
        <xsl:text>;</xsl:text>
        <xsl:text>flip:x y;</xsl:text>

      </xsl:when>
      <xsl:otherwise>
        <!--<xsl:variable name="margin-left">
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="$x1"/>
            <xsl:with-param name="unit" select="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
        <xsl:text>margin-left:</xsl:text>
        <xsl:choose>
          <xsl:when test="number($x1) and $x1 != 0">
            <xsl:value-of select="$x1"/>
            <xsl:text>pt</xsl:text>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
        <xsl:text>;</xsl:text>
        <!--<xsl:variable name="margin-top">
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="$y1"/>
            <xsl:with-param name="unit" select="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
        <xsl:text>margin-top:</xsl:text>
        <xsl:choose>
          <xsl:when test="number($y1) and $y1 != 0">
            <xsl:value-of select="$y1"/>
            <xsl:text>pt</xsl:text>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
        <xsl:text>;</xsl:text>
        <!--<xsl:variable name="width">
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="concat(substring-before($x2,'in')-substring-before($x1,'in'),'in')"/>
            <xsl:with-param name="unit" select="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
        <xsl:text>width:</xsl:text>
        <xsl:choose>
          <xsl:when test="number($x2 - $x1) and ($x2 - $x1) != 0">
            <xsl:value-of select="($x2 - $x1)"/>
            <xsl:text>pt</xsl:text>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
        <xsl:text>;</xsl:text>
        <!--<xsl:variable name="height">
          <xsl:call-template name="ConvertMeasure">
            <xsl:with-param name="length" select="concat(substring-before($y2,'in')-substring-before($y1,'in'),'in')"/>
            <xsl:with-param name="unit" select="'point'"/>
          </xsl:call-template>
        </xsl:variable>-->
        <xsl:text>height:</xsl:text>
        <xsl:choose>
          <xsl:when test="number($y2 - $y1) and ($y2 - $y1) != 0">
            <xsl:value-of select="($y2 - $y1)"/>
            <xsl:text>pt</xsl:text>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
        <xsl:text>;</xsl:text>
      </xsl:otherwise>

    </xsl:choose>
  </xsl:template>
  <!-- Sona Arrow Stroke Function-->
  <xsl:template name="GetLineStroke">
    <xsl:param name="shapeStyle"></xsl:param>
    <v:stroke>
      <!--code added by yeswanth.s : 'linestyle' in DOCX-->
      <xsl:variable name="styleBorderLine">
        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
          <xsl:with-param name="attribName">style:border-line-width</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:if test="$styleBorderLine != ''">
        <xsl:attribute name="linestyle">
          <xsl:choose>
            <xsl:when test="$styleBorderLine">

              <xsl:variable name="innerLineWidth">
                <xsl:call-template name="point-measure">
                  <xsl:with-param name="length" select="substring-before($styleBorderLine,' ' )"/>
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
      </xsl:if>

      <!--end-->

      <!--Dash Styles-->
      <xsl:if test="$shapeStyle/style:graphic-properties/@draw:stroke!='' and $shapeStyle/style:graphic-properties/@draw:stroke!='none'">
        <xsl:variable name="drawStrokeDash" select="$shapeStyle/style:graphic-properties/@draw:stroke-dash"></xsl:variable>
        <xsl:variable name="strokeStyle" select="document('styles.xml')/office:document-styles/office:styles/draw:stroke-dash[@draw:name=$drawStrokeDash]"></xsl:variable>
        <xsl:variable name="drawStroke" select="$shapeStyle/style:graphic-properties/@draw:stroke"/>
        <!-- Sona: Arrow feature continuation -->
        <xsl:variable name="Unit1">
          <xsl:call-template name="GetUnit">
            <xsl:with-param name="length" select="$strokeStyle/@draw:dots1-length"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="Unit2">
          <xsl:call-template name="GetUnit">
            <xsl:with-param name="length" select="$strokeStyle/@draw:dots2-length"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:attribute name="dashstyle">
          <xsl:choose>
            <xsl:when test="$strokeStyle/@draw:dots1=2 and $strokeStyle/@draw:dots2=1">
              <xsl:value-of select="'longDashDotDot'"/>
            </xsl:when>
            <xsl:when test=" not($strokeStyle/@draw:dots1-length)and substring-before($strokeStyle/@draw:dots2-length,$Unit2) &gt;'0.05'">
              <xsl:value-of select="'longDashDot'"/>
            </xsl:when>
            <xsl:when test="not($strokeStyle/@draw:dots1-length)and substring-before($strokeStyle/@draw:dots2-length,$Unit2) &lt;= '0.05'">
              <xsl:value-of select="'dashDot'"/>
            </xsl:when>
            <!-- Square Dot-->
            <xsl:when test="$strokeStyle/@draw:dots1-length=$strokeStyle/@draw:dots2-length and substring-before($strokeStyle/@draw:dots1-length,$Unit1) &lt;= '0.02'">
              <xsl:value-of select="'1 1'"/>
            </xsl:when>
            <xsl:when test="$strokeStyle/@draw:dots1-length=$strokeStyle/@draw:dots2-length and substring-before($strokeStyle/@draw:dots1-length,$Unit1) &lt;= '0.05'">
              <xsl:value-of select="'dash'"/>
            </xsl:when>
            <xsl:when test="$strokeStyle/@draw:dots1-length=$strokeStyle/@draw:dots2-length and substring-before($strokeStyle/@draw:dots1-length,$Unit1) &gt; '0.05'">
              <xsl:value-of select="'longDash'"/>
            </xsl:when>
            <!-- Sona : code update for square or round dots-->
            <xsl:when test="not($strokeStyle/@draw:dots1-length)and not($strokeStyle/@draw:dots2) and $strokeStyle/@draw:dots1">
              <xsl:value-of select="'1 1'"/>
            </xsl:when>
            <!--Sona: Defect #2019464-->
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when test="$drawStroke = 'dash'">
                  <xsl:value-of select="'dash'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'solid'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:if test="not($strokeStyle/@draw:dots1-length)and not($strokeStyle/@draw:dots2) and $strokeStyle/@draw:dots1">
          <xsl:attribute name="endcap">
            <xsl:value-of select="'round'"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:if>

      <!--Arrow Styles-->
      <!-- Start Arrow-->
      <!-- Sonata: Defect #2151408-->
      <xsl:if test="$shapeStyle/style:graphic-properties/@draw:marker-start and $shapeStyle/style:graphic-properties/@draw:marker-start !=''">
        <xsl:variable name="drawArrowTypeStart" select="$shapeStyle/style:graphic-properties/@draw:marker-start"></xsl:variable>
        <xsl:variable name="startArrow" select="document('styles.xml')/office:document-styles/office:styles/draw:marker[@draw:name=$drawArrowTypeStart]"></xsl:variable>
        <xsl:attribute name="startarrow">
          <xsl:choose>
            <xsl:when test="$startArrow/@svg:d='m10 0-10 30h20z'">
              <xsl:value-of select="'block'"/>
            </xsl:when>
            <xsl:when test="$startArrow/@svg:d='m0 2108v17 17l12 42 30 34 38 21 43 4 29-8 30-21 25-26 13-34 343-1532 339 1520 13 42 29 34 39 21 42 4 42-12 34-30 21-42v-39-12l-4 4-440-1998-9-42-25-39-38-25-43-8-42 8-38 25-26 39-8 42z'">
              <xsl:value-of select="'open'"/>
            </xsl:when>
            <xsl:when test="$startArrow/@svg:d='m1013 1491 118 89-567-1580-564 1580 114-85 136-68 148-46 161-17 161 13 153 46z'">
              <xsl:value-of select="'classic'"/>
            </xsl:when>
            <xsl:when test="$startArrow/@svg:d='m462 1118-102-29-102-51-93-72-72-93-51-102-29-102-13-105 13-102 29-106 51-102 72-89 93-72 102-50 102-34 106-9 101 9 106 34 98 50 93 72 72 89 51 102 29 106 13 102-13 105-29 102-51 102-72 93-93 72-98 51-106 29-101 13z'">
              <xsl:value-of select="'oval'"/>
            </xsl:when>
            <xsl:when test="$startArrow/@svg:d='m0 564 564 567 567-567-567-564z'">
              <xsl:value-of select="'diamond'"/>
            </xsl:when>
            <!-- Sona: Defect #2019464-->
            <xsl:otherwise>
              <xsl:value-of select="'block'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <!-- Sonata: Defect #2151408-->
        <xsl:if test="$shapeStyle/style:graphic-properties/@draw:marker-start-width and $shapeStyle/style:graphic-properties/@draw:marker-start-width !=''">
          <xsl:variable name="Unit">
            <xsl:call-template name="GetUnit">
              <xsl:with-param name="length" select="$shapeStyle/style:graphic-properties/@draw:marker-start-width"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:call-template name="setArrowSize">
            <xsl:with-param name="size" select="substring-before($shapeStyle/style:graphic-properties/@draw:marker-start-width,$Unit)" />
            <xsl:with-param name="arrType" select="'start'"></xsl:with-param>
          </xsl:call-template >
        </xsl:if>
      </xsl:if>
      <!-- End Arrow-->
      <!-- Sonata: Defect #2151408-->
      <xsl:if test="$shapeStyle/style:graphic-properties/@draw:marker-end and $shapeStyle/style:graphic-properties/@draw:marker-end !=''">
        <xsl:variable name="drawArrowTypeEnd" select="$shapeStyle/style:graphic-properties/@draw:marker-end"></xsl:variable>
        <xsl:variable name="endArrow" select="document('styles.xml')/office:document-styles/office:styles/draw:marker[@draw:name=$drawArrowTypeEnd]"></xsl:variable>
        <xsl:attribute name="endarrow">
          <xsl:choose>
            <xsl:when test="$endArrow/@svg:d='m10 0-10 30h20z'">
              <xsl:value-of select="'block'"/>
            </xsl:when>
            <xsl:when test="$endArrow/@svg:d='m0 2108v17 17l12 42 30 34 38 21 43 4 29-8 30-21 25-26 13-34 343-1532 339 1520 13 42 29 34 39 21 42 4 42-12 34-30 21-42v-39-12l-4 4-440-1998-9-42-25-39-38-25-43-8-42 8-38 25-26 39-8 42z'">
              <xsl:value-of select="'open'"/>
            </xsl:when>
            <xsl:when test="$endArrow/@svg:d='m1013 1491 118 89-567-1580-564 1580 114-85 136-68 148-46 161-17 161 13 153 46z'">
              <xsl:value-of select="'classic'"/>
            </xsl:when>
            <xsl:when test="$endArrow/@svg:d='m462 1118-102-29-102-51-93-72-72-93-51-102-29-102-13-105 13-102 29-106 51-102 72-89 93-72 102-50 102-34 106-9 101 9 106 34 98 50 93 72 72 89 51 102 29 106 13 102-13 105-29 102-51 102-72 93-93 72-98 51-106 29-101 13z'">
              <xsl:value-of select="'oval'"/>
            </xsl:when>
            <xsl:when test="$endArrow/@svg:d='m0 564 564 567 567-567-567-564z'">
              <xsl:value-of select="'diamond'"/>
            </xsl:when>
            <!-- Sona: Defect #2019464-->
            <xsl:otherwise>
              <xsl:value-of select="'block'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <!-- Sonata: Defect #2151408-->
        <xsl:if test="$shapeStyle/style:graphic-properties/@draw:marker-end-width and $shapeStyle/style:graphic-properties/@draw:marker-end-width !=''">
          <xsl:variable name="Unit">
            <xsl:call-template name="GetUnit">
              <xsl:with-param name="length" select="$shapeStyle/style:graphic-properties/@draw:marker-end-width"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:call-template name="setArrowSize">
            <xsl:with-param name="size" select="substring-before($shapeStyle/style:graphic-properties/@draw:marker-end-width,$Unit)" />
            <xsl:with-param name="arrType" select="'end'"></xsl:with-param>
          </xsl:call-template >
        </xsl:if>
      </xsl:if>
    </v:stroke>
  </xsl:template>
  <xsl:template name="setArrowSize">
    <xsl:param name="size" />
    <xsl:param name="arrType"></xsl:param>

    <xsl:choose>
      <xsl:when test="$arrType='start'">
        <xsl:choose>
          <xsl:when test="($size &lt;= $sm-sm)">
            <xsl:attribute name="startarrowwidth">
              <xsl:value-of select="'narrow'"/>
            </xsl:attribute>
            <xsl:attribute name="startarrowlength">
              <xsl:value-of select="'short'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="($size &gt; $sm-sm) and ($size &lt;= $sm-med)">
            <xsl:attribute name="startarrowwidth">
              <xsl:value-of select="'narrow'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="($size &gt; $sm-med) and ($size &lt;= $sm-lg)">
            <xsl:attribute name="startarrowwidth">
              <xsl:value-of select="'narrow'"/>
            </xsl:attribute>
            <xsl:attribute name="startarrowlength">
              <xsl:value-of select="'long'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="($size &gt; $sm-lg) and ($size &lt;= $med-sm)">
            <xsl:attribute name="startarrowlength">
              <xsl:value-of select="'short'"/>
            </xsl:attribute>
          </xsl:when>
          <!--<xsl:when test="($size &gt; $med-med) and ($size &lt;= $med-lg)">
        <xsl:attribute name="startarrowwidth">
          <xsl:value-of select="'med'"/>
        </xsl:attribute>
        <xsl:attribute name="startarrowlength">
          <xsl:value-of select="'lg'"/>
        </xsl:attribute>
      </xsl:when>-->
          <xsl:when test="($size &gt; $med-lg) and ($size &lt;= $lg-sm)">
            <xsl:attribute name="startarrowlength">
              <xsl:value-of select="'long'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="($size &gt; $lg-sm) and ($size &lt;= $lg-med)">
            <xsl:attribute name="startarrowwidth">
              <xsl:value-of select="'wide'"/>
            </xsl:attribute>
            <xsl:attribute name="startarrowlength">
              <xsl:value-of select="'short'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="($size &gt; $lg-med) and ($size &lt;= $lg-lg)">
            <xsl:attribute name="startarrowwidth">
              <xsl:value-of select="'wide'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="($size &gt; $lg-lg)">
            <xsl:attribute name="startarrowwidth">
              <xsl:value-of select="'wide'"/>
            </xsl:attribute>
            <xsl:attribute name="startarrowlength">
              <xsl:value-of select="'long'"/>
            </xsl:attribute>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="($size &lt;= $sm-sm)">
            <xsl:attribute name="endarrowwidth">
              <xsl:value-of select="'narrow'"/>
            </xsl:attribute>
            <xsl:attribute name="endarrowlength">
              <xsl:value-of select="'short'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="($size &gt; $sm-sm) and ($size &lt;= $sm-med)">
            <xsl:attribute name="endarrowwidth">
              <xsl:value-of select="'narrow'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="($size &gt; $sm-med) and ($size &lt;= $sm-lg)">
            <xsl:attribute name="endarrowwidth">
              <xsl:value-of select="'narrow'"/>
            </xsl:attribute>
            <xsl:attribute name="endarrowlength">
              <xsl:value-of select="'long'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="($size &gt; $sm-lg) and ($size &lt;= $med-sm)">
            <xsl:attribute name="endarrowlength">
              <xsl:value-of select="'short'"/>
            </xsl:attribute>
          </xsl:when>
          <!--<xsl:when test="($size &gt; $med-med) and ($size &lt;= $med-lg)">
        <xsl:attribute name="startarrowwidth">
          <xsl:value-of select="'med'"/>
        </xsl:attribute>
        <xsl:attribute name="startarrowlength">
          <xsl:value-of select="'lg'"/>
        </xsl:attribute>
      </xsl:when>-->
          <xsl:when test="($size &gt; $med-lg) and ($size &lt;= $lg-sm)">
            <xsl:attribute name="endarrowlength">
              <xsl:value-of select="'long'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="($size &gt; $lg-sm) and ($size &lt;= $lg-med)">
            <xsl:attribute name="endarrowwidth">
              <xsl:value-of select="'wide'"/>
            </xsl:attribute>
            <xsl:attribute name="endarrowlength">
              <xsl:value-of select="'short'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="($size &gt; $lg-med) and ($size &lt;= $lg-lg)">
            <xsl:attribute name="endarrowwidth">
              <xsl:value-of select="'wide'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="($size &gt; $lg-lg)">
            <xsl:attribute name="endarrowwidth">
              <xsl:value-of select="'wide'"/>
            </xsl:attribute>
            <xsl:attribute name="endarrowlength">
              <xsl:value-of select="'long'"/>
            </xsl:attribute>
          </xsl:when>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>
</xsl:stylesheet>
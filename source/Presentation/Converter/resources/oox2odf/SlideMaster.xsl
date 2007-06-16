<?xml version="1.0" encoding="UTF-8" ?>
<!--
Copyright (c) 2007, Sonata Software Limited
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
*     * Neither the name of Sonata Software Limited nor the names of its contributors
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
* SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" 
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" 
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" 
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" 
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" 
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" 
  xmlns:xlink="http://www.w3.org/1999/xlink" 
  xmlns:dc="http://purl.org/dc/elements/1.1/" 
  xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" 
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" 
  xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" 
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" 
  xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" 
  xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" 
  xmlns:math="http://www.w3.org/1998/Math/MathML" 
  xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" 
  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" 
  xmlns:ooo="http://openoffice.org/2004/office" 
  xmlns:ooow="http://openoffice.org/2004/writer" 
  xmlns:oooc="http://openoffice.org/2004/calc" 
  xmlns:dom="http://www.w3.org/2001/xml-events" 
  xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" 
  xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0" 
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
  xmlns:rels="http://schemas.openxmlformats.org/package/2006/relationships" 
  exclude-result-prefixes="a style svg fo r">
  <xsl:import href="common.xsl"/>
  
  
  <xsl:template name="SlideMaster" >
    <xsl:call-template name="InsertStandardStyle"/>
    <xsl:call-template name="InsertSlideMasterCommnadFeaturesStyle"/>
    <xsl:call-template name="InsertContentStyle"/>
  </xsl:template>
  <xsl:template name="InsertStandardStyle">

    <style:paragraph-properties fo:margin-left="0cm" fo:margin-right="0cm" 
      fo:margin-top="0cm" fo:margin-bottom="0cm" fo:line-height="100%" 
      text:enable-numbering="false" fo:text-indent="0cm"/>
    <style:text-properties style:use-window-font-color="true" 
      style:text-outline="false" style:text-line-through-style="none" 
      fo:font-family="Arial" style:font-family-generic="roman" 
      style:font-pitch="variable" fo:font-size="18pt" 
      fo:font-style="normal" fo:text-shadow="none" 
      style:text-underline-style="none" fo:font-weight="normal" 
      style:font-family-asian="&apos;Arial Unicode MS&apos;" style:font-family-generic-asian="system" style:font-pitch-asian="variable" style:font-size-asian="18pt" style:font-style-asian="normal" style:font-weight-asian="normal" style:font-family-complex="Tahoma" style:font-family-generic-complex="system" style:font-pitch-complex="variable" style:font-size-complex="18pt" style:font-style-complex="normal" style:font-weight-complex="normal" style:text-emphasize="none" style:font-relief="none"/>
  </xsl:template>
  <xsl:template name="InsertSlideMasterCommnadFeaturesStyle">
    <style:style style:name="objectwitharrow" style:family="graphic" style:parent-style-name="standard">
      <style:graphic-properties draw:stroke="solid" svg:stroke-width="0.15cm" svg:stroke-color="#000000" draw:marker-start="Arrow" draw:marker-start-width="0.7cm" draw:marker-start-center="true" draw:marker-end-width="0.3cm"/>
    </style:style>
    <style:style style:name="objectwithshadow" style:family="graphic" style:parent-style-name="standard">
      <style:graphic-properties draw:shadow="visible" draw:shadow-offset-x="0.3cm" draw:shadow-offset-y="0.3cm" draw:shadow-color="#808080"/>
    </style:style>
    <style:style style:name="objectwithoutfill" style:family="graphic" style:parent-style-name="standard">
      <style:graphic-properties draw:fill="none"/>
    </style:style>
    <style:style style:name="text" style:family="graphic" style:parent-style-name="standard">
      <style:graphic-properties draw:stroke="none" draw:fill="none"/>
    </style:style>
    <style:style style:name="textbody" style:family="graphic" style:parent-style-name="standard">
      <style:graphic-properties draw:stroke="none" draw:fill="none"/>
      <style:text-properties fo:font-size="16pt"/>
    </style:style>
    <style:style style:name="textbodyjustfied" style:family="graphic" style:parent-style-name="standard">
      <style:graphic-properties draw:stroke="none" draw:fill="none"/>
      <style:paragraph-properties fo:text-align="justify"/>
    </style:style>
    <style:style style:name="textbodyindent" style:family="graphic" style:parent-style-name="standard">
      <style:graphic-properties draw:stroke="none" draw:fill="none">
        <text:list-style>
          <text:list-level-style-bullet text:level="1" text:bullet-char="●">
            <style:list-level-properties text:space-before="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="2" text:bullet-char="●">
            <style:list-level-properties text:space-before="0.6cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="3" text:bullet-char="●">
            <style:list-level-properties text:space-before="1.2cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="4" text:bullet-char="●">
            <style:list-level-properties text:space-before="1.8cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="5" text:bullet-char="●">
            <style:list-level-properties text:space-before="2.4cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="6" text:bullet-char="●">
            <style:list-level-properties text:space-before="3cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="7" text:bullet-char="●">
            <style:list-level-properties text:space-before="3.6cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="8" text:bullet-char="●">
            <style:list-level-properties text:space-before="4.2cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="9" text:bullet-char="●">
            <style:list-level-properties text:space-before="4.8cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="10" text:bullet-char="●">
            <style:list-level-properties text:space-before="5.4cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
        </text:list-style>
      </style:graphic-properties>
      <style:paragraph-properties fo:margin-left="0cm" fo:margin-right="0cm" fo:text-indent="0.6cm"/>
    </style:style>
    <style:style style:name="title" style:family="graphic" style:parent-style-name="standard">
      <style:graphic-properties draw:stroke="none" draw:fill="none"/>
      <style:text-properties fo:font-size="44pt"/>
    </style:style>
    <style:style style:name="title1" style:family="graphic" style:parent-style-name="standard">
      <style:graphic-properties draw:stroke="none" draw:fill="solid" draw:fill-color="#008080" draw:shadow="visible" draw:shadow-offset-x="0.2cm" draw:shadow-offset-y="0.2cm" draw:shadow-color="#808080"/>
      <style:paragraph-properties fo:text-align="center"/>
      <style:text-properties fo:font-size="24pt"/>
    </style:style>
    <style:style style:name="title2" style:family="graphic" style:parent-style-name="standard">
      <style:graphic-properties svg:stroke-width="0.05cm" draw:fill-color="#ffcc99" draw:shadow="visible" draw:shadow-offset-x="0.2cm" draw:shadow-offset-y="0.2cm" draw:shadow-color="#808080">
        <text:list-style>
          <text:list-level-style-bullet text:level="1" text:bullet-char="●">
            <style:list-level-properties text:space-before="0.2cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="2" text:bullet-char="●">
            <style:list-level-properties text:space-before="0.6cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="3" text:bullet-char="●">
            <style:list-level-properties text:space-before="1.2cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="4" text:bullet-char="●">
            <style:list-level-properties text:space-before="1.8cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="5" text:bullet-char="●">
            <style:list-level-properties text:space-before="2.4cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="6" text:bullet-char="●">
            <style:list-level-properties text:space-before="3cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="7" text:bullet-char="●">
            <style:list-level-properties text:space-before="3.6cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="8" text:bullet-char="●">
            <style:list-level-properties text:space-before="4.2cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="9" text:bullet-char="●">
            <style:list-level-properties text:space-before="4.8cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="10" text:bullet-char="●">
            <style:list-level-properties text:space-before="5.4cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
        </text:list-style>
      </style:graphic-properties>
      <style:paragraph-properties fo:margin-left="0.2cm" fo:margin-right="0.2cm" fo:margin-top="0.1cm" fo:margin-bottom="0.1cm" fo:text-align="center" fo:text-indent="0cm"/>
      <style:text-properties fo:font-size="36pt"/>
    </style:style>
    <style:style style:name="headline" style:family="graphic" style:parent-style-name="standard">
      <style:graphic-properties draw:stroke="none" draw:fill="none"/>
      <style:paragraph-properties fo:margin-top="0.42cm" fo:margin-bottom="0.21cm"/>
      <style:text-properties fo:font-size="24pt"/>
    </style:style>
    <style:style style:name="headline1" style:family="graphic" style:parent-style-name="standard">
      <style:graphic-properties draw:stroke="none" draw:fill="none"/>
      <style:paragraph-properties fo:margin-top="0.42cm" fo:margin-bottom="0.21cm"/>
      <style:text-properties fo:font-size="18pt" fo:font-weight="bold"/>
    </style:style>
    <style:style style:name="headline2" style:family="graphic" style:parent-style-name="standard">
      <style:graphic-properties draw:stroke="none" draw:fill="none"/>
      <style:paragraph-properties fo:margin-top="0.42cm" fo:margin-bottom="0.21cm"/>
      <style:text-properties fo:font-size="14pt" fo:font-style="italic" fo:font-weight="bold"/>
    </style:style>
    <style:style style:name="measure" style:family="graphic" style:parent-style-name="standard">
      <style:graphic-properties draw:stroke="solid" draw:marker-start="Arrow" draw:marker-start-width="0.2cm" draw:marker-end="Arrow" draw:marker-end-width="0.2cm" draw:fill="none"/>
      <style:text-properties fo:font-size="12pt"/>
    </style:style>
  </xsl:template>
  <xsl:template name="InsertContentStyle">
    <xsl:for-each select="document('ppt/presentation.xml')//p:sldMasterIdLst/p:sldMasterId">
      <xsl:variable name="sldMasterIdRelation">
        <xsl:value-of select="@r:id"></xsl:value-of>
      </xsl:variable>
      <!-- Loop thru each slide master.xml-->
      <xsl:for-each select="document('ppt/_rels/presentation.xml.rels')//node()[@Id=$sldMasterIdRelation]">

        <xsl:variable name="slideMasterPath">
          <xsl:value-of select="substring-after(@Target,'/')"/>
        </xsl:variable>
        <xsl:variable name="slideMasterName">
          <xsl:value-of select="substring-before($slideMasterPath,'.xml')"/>
        </xsl:variable>
        <!--style for drawing page-->
        <xsl:call-template name="slideMasterDrawingPage"/>
        <!-- style for Title-->
        <style:style>
          <xsl:attribute name="style:name">
            <xsl:value-of select="concat($slideMasterName,'-title')"/>
          </xsl:attribute>
          <xsl:attribute name ="style:family">
            <xsl:value-of select ="'presentation'"/>
          </xsl:attribute>
          <xsl:call-template name="slideMasterTitle">
            <xsl:with-param name="SlideMasterFile">
              <xsl:value-of select="$slideMasterPath"/>
            </xsl:with-param>
          </xsl:call-template>

        </style:style>
        <!-- style for sub-Title-->
        <style:style>
          <xsl:attribute name="style:name">
            <xsl:value-of select="concat($slideMasterName,'-subtitle')"/>
          </xsl:attribute>
          <xsl:attribute name ="style:family">
            <xsl:value-of select ="'presentation'"/>
          </xsl:attribute>
          <xsl:call-template name="slideMasterOutlineStyle">
            <xsl:with-param name="SlideMasterFile">
              <xsl:value-of select="$slideMasterPath"/>
            </xsl:with-param>
          </xsl:call-template>

        </style:style>
        <!-- for style of Slide Master's Background Color-->
        <style:style>
          <xsl:attribute name="style:name">
            <xsl:value-of select="concat($slideMasterName,'-background')"/>
          </xsl:attribute>
          <xsl:attribute name ="style:family">
            <xsl:value-of select ="'presentation'"/>
          </xsl:attribute>
          <style:graphic-properties>
            <xsl:attribute name ="draw:stroke">
              <xsl:value-of select ="'none'"/>
            </xsl:attribute>
            <xsl:call-template name="getSlideMasterBGColor">
              <xsl:with-param name="slideMasterFile">
                <xsl:value-of select="$slideMasterPath"/>
              </xsl:with-param>
            </xsl:call-template>
      <xsl:attribute name ="draw:fill-image-width">
        <xsl:value-of select ="'0cm'"/>
      </xsl:attribute>
      <xsl:attribute name ="draw:fill-image-height">
        <xsl:value-of select ="'0cm'"/>
      </xsl:attribute>

      </style:graphic-properties>
      </style:style>
      <!-- style for Outline 1-->
      <style:style>
        <xsl:attribute name="style:name">
          <xsl:value-of select="concat($slideMasterName,'-outline1')"/>
        </xsl:attribute>
        <xsl:attribute name ="style:family">
          <xsl:value-of select ="'presentation'"/>
        </xsl:attribute>
        <style:graphic-properties>
          <xsl:for-each select="document(concat('ppt/slideMasters/',$slideMasterPath))//p:sp">
            <xsl:choose>
              <xsl:when test="p:nvSpPr/p:nvPr/p:ph/@type='body'">
                <xsl:call-template name="getSPBackColor"/>
              </xsl:when>
            </xsl:choose>
          </xsl:for-each>
         
            <text:list-style>
          <!--<xsl:call-template  name ="slideMasterInternalMargins">
            <xsl:with-param name ="phType">
              <xsl:value-of select="'body'"/>
            </xsl:with-param>

            <xsl:with-param name ="defType">
              <xsl:value-of select="'anchor'"/>
            </xsl:with-param>
            <xsl:with-param name="SlideMasterFile">
              <xsl:value-of select="$slideMasterPath"/>
            </xsl:with-param>
          </xsl:call-template>
          <xsl:call-template  name ="slideMasterInternalMargins">
            <xsl:with-param name ="phType">
              <xsl:value-of select="'body'"/>
            </xsl:with-param>

            <xsl:with-param name ="defType">
              <xsl:value-of select="'bIns'"/>
            </xsl:with-param>
            <xsl:with-param name="SlideMasterFile">
              <xsl:value-of select="$slideMasterPath"/>
            </xsl:with-param>
          </xsl:call-template>
          <xsl:call-template  name ="slideMasterInternalMargins">
            <xsl:with-param name ="phType">
              <xsl:value-of select="'body'"/>
            </xsl:with-param>
            <xsl:with-param name ="defType">
              <xsl:value-of select="'lIns'"/>
            </xsl:with-param>
            <xsl:with-param name="SlideMasterFile">
                <xsl:value-of select="$slideMasterPath"/>
              </xsl:with-param>
            </xsl:call-template>
            <xsl:call-template  name ="slideMasterInternalMargins">
              <xsl:with-param name ="phType">
                <xsl:value-of select="'body'"/>
              </xsl:with-param>

              <xsl:with-param name ="defType">
                <xsl:value-of select="'rIns'"/>
              </xsl:with-param>
              <xsl:with-param name="SlideMasterFile">
                <xsl:value-of select="$slideMasterPath"/>
              </xsl:with-param>
            </xsl:call-template>
            <xsl:call-template  name ="slideMasterInternalMargins">
              <xsl:with-param name ="phType">
                <xsl:value-of select="'body'"/>
              </xsl:with-param>

              <xsl:with-param name ="defType">
                <xsl:value-of select="'tIns'"/>
              </xsl:with-param>
              <xsl:with-param name="SlideMasterFile">
                <xsl:value-of select="$slideMasterPath"/>
              </xsl:with-param>
            </xsl:call-template>
            <xsl:call-template  name ="slideMasterInternalMargins">
              <xsl:with-param name ="phType">
                <xsl:value-of select="'body'"/>
              </xsl:with-param>

              <xsl:with-param name ="defType">
                <xsl:value-of select="'anchorCtr'"/>
              </xsl:with-param>
              <xsl:with-param name="SlideMasterFile">
                <xsl:value-of select="$slideMasterPath"/>
              </xsl:with-param>
            </xsl:call-template>
            <xsl:call-template  name ="slideMasterInternalMargins">
              <xsl:with-param name ="phType">
                <xsl:value-of select="'body'"/>
              </xsl:with-param>

              <xsl:with-param name ="defType">
                <xsl:value-of select="'wrap'"/>
              </xsl:with-param>
              <xsl:with-param name="SlideMasterFile">
                <xsl:value-of select="$slideMasterPath"/>
              </xsl:with-param>
            </xsl:call-template>-->
          
            <xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMasterPath))//p:sp">
              <xsl:if test="p:nvSpPr/p:nvPr/p:ph/@type='body'">
                <xsl:for-each select ="./p:txBody/a:p">
                  <xsl:call-template name="slideMasterBullett">
                    <xsl:with-param name="SlideMasterFile">
                      <xsl:value-of select="$slideMasterPath"/>
                    </xsl:with-param>
                    <xsl:with-param name="levelNo">
                      <xsl:value-of select="./a:pPr/@lvl"/>
                    </xsl:with-param>
                    <xsl:with-param name="pos">
                      <xsl:value-of select="position()"/>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:for-each>
              </xsl:if>
              </xsl:for-each>
            
            </text:list-style>
          </style:graphic-properties>
        <xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMasterPath))//p:txStyles/p:bodyStyle/a:lvl1pPr">
          <xsl:call-template name="Outlines">
            <xsl:with-param name="level">
              <xsl:value-of select="'1'"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:for-each>

      </style:style>
        <!-- style for other Outlines-->
        
          <xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMasterPath))//p:txStyles/p:bodyStyle">
            
             <xsl:for-each select="./a:lvl2pPr">
                <style:style>
                <xsl:attribute name="style:name">
                  <xsl:value-of select="concat($slideMasterName,'-outline2')"/>
                </xsl:attribute>
                <xsl:attribute name ="style:family">
                  <xsl:value-of select ="'presentation'"/>
                </xsl:attribute>
                <xsl:attribute name ="style:parent-style-name">
                  <xsl:value-of select ="concat($slideMasterName,'-outline1')"/>
                </xsl:attribute>
                
                  <xsl:call-template name="Outlines">
                    <xsl:with-param name="level">
                      <xsl:value-of select="'2'"/>
                    </xsl:with-param>
                  </xsl:call-template>
                </style:style>
              </xsl:for-each>
            <xsl:for-each select="./a:lvl3pPr">
              <style:style>
                <xsl:attribute name="style:name">
                  <xsl:value-of select="concat($slideMasterName,'-outline3')"/>
                </xsl:attribute>
                <xsl:attribute name ="style:family">
                  <xsl:value-of select ="'presentation'"/>
                </xsl:attribute>
                <xsl:attribute name ="style:parent-style-name">
                  <xsl:value-of select ="concat($slideMasterName,'-outline2')"/>
                </xsl:attribute>

                <xsl:call-template name="Outlines">
                      <xsl:with-param name="level">
                        <xsl:value-of select="'3'"/>
                      </xsl:with-param>
                    </xsl:call-template>
                  </style:style>
              </xsl:for-each>
            <xsl:for-each select="./a:lvl4pPr">
              <style:style>
                <xsl:attribute name="style:name">
                  <xsl:value-of select="concat($slideMasterName,'-outline4')"/>
                </xsl:attribute>
                <xsl:attribute name ="style:family">
                  <xsl:value-of select ="'presentation'"/>
                </xsl:attribute>
                <xsl:attribute name ="style:parent-style-name">
                  <xsl:value-of select ="concat($slideMasterName,'-outline3')"/>
                </xsl:attribute>

                <xsl:call-template name="Outlines">
                      <xsl:with-param name="level">
                        <xsl:value-of select="'4'"/>
                      </xsl:with-param>
                    </xsl:call-template>
              </style:style>
                  </xsl:for-each>
            <xsl:for-each select="./a:lvl5pPr">
                <style:style>
                  <xsl:attribute name="style:name">
                    <xsl:value-of select="concat($slideMasterName,'-outline5')"/>
                  </xsl:attribute>
                  <xsl:attribute name ="style:family">
                    <xsl:value-of select ="'presentation'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="style:parent-style-name">
                    <xsl:value-of select ="concat($slideMasterName,'-outline4')"/>
                  </xsl:attribute>
                
                    <xsl:call-template name="Outlines">
                      <xsl:with-param name="level">
                        <xsl:value-of select="'5'"/>
                      </xsl:with-param>
                    </xsl:call-template>
                </style:style>
              </xsl:for-each>
            <xsl:for-each select="./a:lvl6pPr">
              <style:style>
                <xsl:attribute name="style:name">
                  <xsl:value-of select="concat($slideMasterName,'-outline6')"/>
                </xsl:attribute>
                <xsl:attribute name ="style:family">
                  <xsl:value-of select ="'presentation'"/>
                </xsl:attribute>
                <xsl:attribute name ="style:parent-style-name">
                  <xsl:value-of select ="concat($slideMasterName,'-outline5')"/>
                </xsl:attribute>

                <xsl:call-template name="Outlines">
                      <xsl:with-param name="level">
                        <xsl:value-of select="'6'"/>
                      </xsl:with-param>
                    </xsl:call-template>
                  </style:style>
                  </xsl:for-each>
            <xsl:for-each select="./a:lvl7pPr">
              <style:style>
                <xsl:attribute name="style:name">
                  <xsl:value-of select="concat($slideMasterName,'-outline7')"/>
                </xsl:attribute>
                <xsl:attribute name ="style:family">
                  <xsl:value-of select ="'presentation'"/>
                </xsl:attribute>
                <xsl:attribute name ="style:parent-style-name">
                  <xsl:value-of select ="concat($slideMasterName,'-outline6')"/>
                </xsl:attribute>

                <xsl:call-template name="Outlines">
                      <xsl:with-param name="level">
                        <xsl:value-of select="'7'"/>
                      </xsl:with-param>
                    </xsl:call-template>
                  </style:style>
              </xsl:for-each>
            <xsl:for-each select="./a:lvl8pPr">
              <style:style>
                <xsl:attribute name="style:name">
                  <xsl:value-of select="concat($slideMasterName,'-outline8')"/>
                </xsl:attribute>
                <xsl:attribute name ="style:family">
                  <xsl:value-of select ="'presentation'"/>
                </xsl:attribute>
                <xsl:attribute name ="style:parent-style-name">
                  <xsl:value-of select ="concat($slideMasterName,'-outline7')"/>
                </xsl:attribute>

                <xsl:call-template name="Outlines">
                      <xsl:with-param name="level">
                        <xsl:value-of select="'8'"/>
                      </xsl:with-param>
                    </xsl:call-template>
                  </style:style>
              </xsl:for-each>
            <xsl:for-each select="./a:lvl9pPr">
              <style:style>
                <xsl:attribute name="style:name">
                  <xsl:value-of select="concat($slideMasterName,'-outline9')"/>
                </xsl:attribute>
                <xsl:attribute name ="style:family">
                  <xsl:value-of select ="'presentation'"/>
                </xsl:attribute>
                <xsl:attribute name ="style:parent-style-name">
                  <xsl:value-of select ="concat($slideMasterName,'-outline8')"/>
                </xsl:attribute>

                <xsl:call-template name="Outlines">
                      <xsl:with-param name="level">
                        <xsl:value-of select="'9'"/>
                      </xsl:with-param>
                    </xsl:call-template>
                  </style:style>
              </xsl:for-each>
          </xsl:for-each>
      </xsl:for-each>
    </xsl:for-each>

  </xsl:template>
  <xsl:template name="ListStyleLevels">
    <xsl:param name="xpath"></xsl:param>
    <xsl:param name="level"></xsl:param>
    <xsl:if test="document(concat('ppt/slideMasters/',$xpath))//p:txStyles/p:bodyStyle/a:lvl1pPr">
      <text:list-level-style-bullet  text:min-label-width="0.6cm" >
        <xsl:attribute name="text:level">
          <xsl:value-of select ='$level'/>
        </xsl:attribute>
        <xsl:if test="a:buChar">
          <xsl:attribute name="text:bullet-char">
            <xsl:value-of select ="a:buChar/@char"/>
          </xsl:attribute>
        </xsl:if>
        <style:list-level-properties >
          <xsl:if test="a:spcBef">
            <xsl:attribute name="text:space-before">
              <xsl:call-template name="ConvertEmu">
                <xsl:with-param name="length" select="a:spcPct/@val"/>
                <xsl:with-param name="unit">cm</xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
        </style:list-level-properties>
        <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true"/>
        <!--<xsl:attribute name ="fo:font-size">
          <xsl:value-of select ="concat(a:defRPr/@sz div 100 ,'pt')"/>
        </xsl:attribute>-->

      </text:list-level-style-bullet>
    </xsl:if>

  </xsl:template>
  <!-- @@Get bullet chars and level from slide master 
This tmplate will get the bullet level and charactor and Marginleft-->
  <xsl:template name ="slideMasterBullett">
    <xsl:param name="SlideMasterFile" />
    <xsl:param name="levelNo"></xsl:param>
    <xsl:param name="pos"></xsl:param>
    <xsl:for-each select ="document(concat('ppt/slideMasters/',$SlideMasterFile))//p:txStyles/p:bodyStyle">
   
        <text:list-level-style-bullet>
          <xsl:attribute name ="text:level">
            <xsl:value-of select ="$pos"/>
          </xsl:attribute>
          <!--<xsl:choose>
          <xsl:when test="$pos &lt; 2">
            <xsl:attribute name ="text:level">
              <xsl:value-of select ="$pos"/>
            </xsl:attribute>
          </xsl:when>
          </xsl:choose>-->
          <xsl:choose>
            <xsl:when test="$levelNo='0'">
              <xsl:attribute name ="text:bullet-char">
                <xsl:value-of select ="a:lvl1pPr/a:buChar/@char"/>
              </xsl:attribute>
              <style:list-level-properties text:space-before="0.3cm" text:min-label-width="0.9cm">

                <!--text:min-label-width="0.6cm">-->
                <!--<xsl:attribute name ="text:min-label-width">
                <xsl:value-of select ="concat(format-number(a:lvl1pPr/a:defRPr/@sz div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>-->
                <!--<xsl:attribute name ="text:space-before">
                <xsl:value-of select ="concat(format-number(a:lvl1pPr/a:spcBef/a:spcPct/@val div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>-->
                <!--<xsl:attribute name ="text:space-before">
                <xsl:value-of select ="concat(format-number(a:lvl1pPr/@marL div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>-->
              </style:list-level-properties >
            
            </xsl:when>
            <xsl:when test="$levelNo='1'">
              <xsl:attribute name ="text:bullet-char">
                <xsl:value-of select ="a:lvl2pPr/a:buChar/@char"/>
              </xsl:attribute>
              <style:list-level-properties text:space-before="1.6cm" text:min-label-width="0.8cm">
              </style:list-level-properties >
              <xsl:for-each select="./a:lvl2pPr">
              <xsl:call-template name="Outlines">
                <xsl:with-param name="level">
                  <xsl:value-of select="'2'"/>
                </xsl:with-param>
              </xsl:call-template>
              </xsl:for-each>
              </xsl:when>
            <xsl:when test="$levelNo='2'">
              <xsl:attribute name ="text:bullet-char">
                <xsl:value-of select ="a:lvl3pPr/a:buChar/@char"/>
              </xsl:attribute>
              <style:list-level-properties text:space-before="3cm" text:min-label-width="0.6cm"/>
              <xsl:for-each select="./a:lvl3pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'2'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </xsl:when>
            <xsl:when test="$levelNo='3'">
              <xsl:attribute name ="text:bullet-char">
                <xsl:value-of select ="a:lvl4pPr/a:buChar/@char"/>
              </xsl:attribute>
              <style:list-level-properties text:space-before="4.2cm" text:min-label-width="0.6cm"/>
              <xsl:for-each select="./a:lvl4pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'3'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each> 
            </xsl:when>
            <xsl:when test="$levelNo='4'">
              <xsl:attribute name ="text:bullet-char">
                <xsl:value-of select ="a:lvl5pPr/a:buChar/@char"/>
              </xsl:attribute>
              <style:list-level-properties text:space-before="5.4cm" text:min-label-width="0.6cm"/>
              <xsl:for-each select="./a:lvl5pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'4'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </xsl:when>
            <xsl:when test="$levelNo='5'">
              <xsl:attribute name ="text:bullet-char">
                <xsl:value-of select ="a:lvl6pPr/a:buChar/@char"/>
              </xsl:attribute>
              <style:list-level-properties text:space-before="6.6cm" text:min-label-width="0.6cm"/>
              <xsl:for-each select="./a:lvl6pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'5'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </xsl:when>
            <xsl:when test="$levelNo='6'">
              <xsl:attribute name ="text:bullet-char">
                <xsl:value-of select ="a:lvl7pPr/a:buChar/@char"/>
              </xsl:attribute>
              <style:list-level-properties text:space-before="7.8cm" text:min-label-width="0.6cm"/>
              <xsl:for-each select="./a:lvl7pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'6'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </xsl:when>
            <xsl:when test="$levelNo='7'">
              <xsl:attribute name ="text:bullet-char">
                <xsl:value-of select ="a:lvl8pPr/a:buChar/@char"/>
              </xsl:attribute>
              <style:list-level-properties text:space-before="9cm" text:min-label-width="0.6cm"/>
              <xsl:for-each select="./a:lvl8pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'8'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </xsl:when>
            <xsl:when test="$levelNo='8'">
              <xsl:attribute name ="text:bullet-char">
                <xsl:value-of select ="a:lvl9pPr/a:buChar/@char"/>
              </xsl:attribute>
              <style:list-level-properties text:space-before="10.2cm" text:min-label-width="0.6cm"/>
              <xsl:for-each select="./a:lvl9pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'9'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </xsl:when>
          </xsl:choose>
          <!--<style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true">
            <xsl:attribute name ="fo:font-size">
              <xsl:value-of select ="concat(a:lvl1pPr/a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
          </style:text-properties>
          <style:paragraph-properties >
            <xsl:attribute name ="fo:margin-left">
              <xsl:value-of select ="concat(format-number(a:lvl1pPr/@marL div 360000 ,'#.##'),'cm')"/>
            </xsl:attribute>
          </style:paragraph-properties>-->
        
        </text:list-level-style-bullet>
      
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="slideMasterOutlines">
    <xsl:param name="SlideMasterFile" />
    <xsl:param name="level"></xsl:param>

    <xsl:for-each select="document(concat('ppt/slideMasters/',$SlideMasterFile))//p:txStyles/p:bodyStyle">
      <xsl:choose>
        <xsl:when test="$level='1'">
          <style:paragraph-properties text:enable-numbering="true">
            <xsl:attribute name ="fo:margin-left">
              <xsl:value-of select ="concat(format-number(a:lvl1pPr/@marL div 360000 ,'#.##'),'cm')"/>
            </xsl:attribute>

            <xsl:if test="a:lvl1pPr/@marR">
              <xsl:attribute name ="fo:margin-right">
                <xsl:value-of select ="concat(format-number(a:lvl1pPr/@marR div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>

            </xsl:if>
            <xsl:if test="a:lvl1pPr/@marT">
              <xsl:attribute name ="fo:margin-top">
                <xsl:value-of select ="concat(format-number(a:lvl1pPr/@marT div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:lvl1pPr/@marB">
              <xsl:attribute name ="fo:margin-bottom">
                <xsl:value-of select ="concat(format-number(a:lvl1pPr/@marB div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:lvl1pPr/@indent">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select ="concat(format-number(a:lvl1pPr/@indent div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </style:paragraph-properties>
          <style:text-properties>
            <xsl:choose>
              <xsl:when test="a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr">
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="concat('#',a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="a:lvl1pPr/a:solidFill/a:schemeClr">
                <xsl:attribute name ="fo:color">
                  <xsl:call-template name ="getColorCode">
                    <xsl:with-param name ="color">
                      <xsl:value-of select="a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                    </xsl:with-param>
                    <xsl:with-param name ="lumMod">
                      <xsl:choose>
                        <xsl:when test="a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod">
                          <xsl:value-of select="a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:with-param>
                    <xsl:with-param name ="lumOff">
                      <xsl:choose>
                        <xsl:when test="a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff">
                          <xsl:value-of select="a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>

                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="'#000000'"/>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>

            <xsl:choose>
              <xsl:when test="a:lvl1pPr/a:defRPr/@sz">
                <xsl:attribute name ="fo:font-size">
                  <xsl:value-of select ="concat(a:lvl1pPr/a:defRPr/@sz div 100 ,'pt')"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:choose>
              <xsl:when test="a:lvl1pPr/a:defRPr/@b='1'">
                <xsl:attribute name ="fo:font-weight">
                  <xsl:value-of select ="'bold'"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="a:lvl1pPr/a:defRPr/@i='1'">
                <xsl:attribute name ="fo:font-style">
                  <xsl:value-of select ="'italic'"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:attribute name ="fo:font-family">
              <xsl:if test ="a:lvl1pPr/a:defRPr/a:latin/@typeface">
                <xsl:variable name ="typeFaceVal" select ="a:lvl1pPr/a:defRPr/a:latin/@typeface"/>
                <xsl:for-each select ="a:lvl1pPr/a:defRPr/a:latin/@typeface">
                  <xsl:if test ="$typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt'">
                    <xsl:call-template name="slideMasterFontName">
                      <xsl:with-param name="fontType">
                        <xsl:value-of select="'minor'"/>
                      </xsl:with-param>
                    </xsl:call-template>
                  </xsl:if>
                  <!-- Bug 1711910 Fixed,On date 2-06-07,by Vijayeta-->
                  <xsl:if test ="not($typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt')">
                    <xsl:value-of select ="."/>
                  </xsl:if>
                </xsl:for-each>
              </xsl:if>
              <xsl:if test ="not(a:lvl1pPr/a:defRPr/a:latin/@typeface)">
                <xsl:call-template name="slideMasterFontName">
                  <xsl:with-param name="fontType">
                    <xsl:value-of select="'minor'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:if>
            </xsl:attribute>
            <xsl:attribute name ="style:font-size-asian">
              <xsl:value-of select ="concat(a:lvl1pPr/a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
            <xsl:attribute name ="style:font-size-complex">
              <xsl:value-of select ="concat(a:lvl1pPr/a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
            <xsl:choose>
              <xsl:when test="a:lvl1pPr/a:defRPr/@u='1'">
                <xsl:attribute name ="style:text-underline-style">
                  <xsl:value-of select ="'solid'"></xsl:value-of>
                </xsl:attribute>
                <xsl:attribute name ="style:text-underline-width">
                  <xsl:value-of select ="'auto'"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
        
          </style:text-properties>
        </xsl:when>
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="$level='2'">
          <style:paragraph-properties >

            <xsl:attribute name ="fo:margin-left">
              <xsl:value-of select ="concat(format-number(a:lvl2pPr/@marL div 360000 ,'#.##'),'cm')"/>
            </xsl:attribute>

            <xsl:if test="a:lvl2pPr/@marR">
              <xsl:attribute name ="fo:margin-right">
                <xsl:value-of select ="concat(format-number(a:lvl2pPr/@marR div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>

            </xsl:if>
            <xsl:if test="a:lvl2pPr/@marT">
              <xsl:attribute name ="fo:margin-top">
                <xsl:value-of select ="concat(format-number(a:lvl2pPr/@marT div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:lvl2pPr/@marB">
              <xsl:attribute name ="fo:margin-bottom">
                <xsl:value-of select ="concat(format-number(a:lvl2pPr/@marB div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:lvl2pPr/@indent">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select ="concat(format-number(a:lvl2pPr/@indent div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </style:paragraph-properties>
          <style:text-properties>
            <xsl:choose>
              <xsl:when test="a:lvl2pPr/a:defRPr/a:solidFill/a:srgbClr">
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="concat('#',a:lvl2pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="a:lvl2pPr/a:solidFill/a:schemeClr">
                <xsl:attribute name ="fo:color">
                  <xsl:call-template name ="getColorCode">
                    <xsl:with-param name ="color">
                      <xsl:value-of select="a:lvl2pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                    </xsl:with-param>
                    <xsl:with-param name ="lumMod">
                      <xsl:choose>
                        <xsl:when test="a:lvl2pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod">
                          <xsl:value-of select="a:lvl2pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:with-param>
                    <xsl:with-param name ="lumOff">
                      <xsl:choose>
                        <xsl:when test="a:lvl2pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff">
                          <xsl:value-of select="a:lvl2pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>

                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="'#000000'"/>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>

            <xsl:choose>
              <xsl:when test="a:lvl2pPr/a:defRPr/@sz">
                <xsl:attribute name ="fo:font-size">
                  <xsl:value-of select ="concat(a:lvl2pPr/a:defRPr/@sz div 100 ,'pt')"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:choose>
              <xsl:when test="a:lvl2pPr/a:defRPr/@b='1'">
                <xsl:attribute name ="fo:font-weight">
                  <xsl:value-of select ="'bold'"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="a:lvl2pPr/a:defRPr/@i='1'">
                <xsl:attribute name ="fo:font-style">
                  <xsl:value-of select ="'italic'"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:attribute name ="fo:font-family">
              <xsl:call-template name="slideMasterFontName">
                <xsl:with-param name="fontType">
                  <xsl:value-of select="'minor'"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
            <xsl:attribute name ="style:font-size-asian">
              <xsl:value-of select ="concat(a:lvl2pPr/a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
            <xsl:attribute name ="style:font-size-complex">
              <xsl:value-of select ="concat(a:lvl2pPr/a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
            <xsl:choose>
              <xsl:when test="a:lvl2pPr/a:defRPr/@u='1'">
                <xsl:attribute name ="style:text-underline-style">
                  <xsl:value-of select ="'solid'"></xsl:value-of>
                </xsl:attribute>
                <xsl:attribute name ="style:text-underline-width">
                  <xsl:value-of select ="'auto'"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>

          </style:text-properties>
        </xsl:when>
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="$level='3'">
          <style:paragraph-properties >
            <xsl:attribute name ="fo:margin-left">
              <xsl:value-of select ="concat(format-number(a:lvl3pPr/@marL div 360000 ,'#.##'),'cm')"/>
            </xsl:attribute>

            <xsl:if test="a:lvl3pPr/@marR">
              <xsl:attribute name ="fo:margin-right">
                <xsl:value-of select ="concat(format-number(a:lvl3pPr/@marR div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>

            </xsl:if>

            <xsl:if test="a:lvl3pPr/@marT">
              <xsl:attribute name ="fo:margin-top">
                <xsl:value-of select ="concat(format-number(a:lvl3pPr/@marT div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:lvl3pPr/@marB">
              <xsl:attribute name ="fo:margin-bottom">
                <xsl:value-of select ="concat(format-number(a:lvl3pPr/@marB div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:lvl3pPr/@indent">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select ="concat(format-number(a:lvl3pPr/@indent div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </style:paragraph-properties>
          <style:text-properties>
            <xsl:choose>
              <xsl:when test="a:lvl3pPr/a:defRPr/a:solidFill/a:srgbClr">
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="concat('#',a:lvl3pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="a:lvl3pPr/a:solidFill/a:schemeClr">
                <xsl:attribute name ="fo:color">
                  <xsl:call-template name ="getColorCode">
                    <xsl:with-param name ="color">
                      <xsl:value-of select="a:lvl3pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                    </xsl:with-param>
                    <xsl:with-param name ="lumMod">
                      <xsl:choose>
                        <xsl:when test="a:lvl3pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod">
                          <xsl:value-of select="a:lvl3pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:with-param>
                    <xsl:with-param name ="lumOff">
                      <xsl:choose>
                        <xsl:when test="a:lvl3pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff">
                          <xsl:value-of select="a:lvl3pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>

                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="'#000000'"/>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>

            <xsl:choose>
              <xsl:when test="a:lvl3pPr/a:defRPr/@sz">
                <xsl:attribute name ="fo:font-size">
                  <xsl:value-of select ="concat(a:lvl3pPr/a:defRPr/@sz div 100 ,'pt')"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:choose>
              <xsl:when test="a:lvl3pPr/a:defRPr/@b='1'">
                <xsl:attribute name ="fo:font-weight">
                  <xsl:value-of select ="'bold'"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="a:lvl3pPr/a:defRPr/@i='1'">
                <xsl:attribute name ="fo:font-style">
                  <xsl:value-of select ="'italic'"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:attribute name ="fo:font-family">
              <xsl:call-template name="slideMasterFontName">
                <xsl:with-param name="fontType">
                  <xsl:value-of select="'minor'"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
            <xsl:attribute name ="style:font-size-asian">
              <xsl:value-of select ="concat(a:lvl3pPr/a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
            <xsl:attribute name ="style:font-size-complex">
              <xsl:value-of select ="concat(a:lvl3pPr/a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
            <xsl:choose>
              <xsl:when test="a:lvl3pPr/a:defRPr/@u='1'">
                <xsl:attribute name ="style:text-underline-style">
                  <xsl:value-of select ="'solid'"></xsl:value-of>
                </xsl:attribute>
                <xsl:attribute name ="style:text-underline-width">
                  <xsl:value-of select ="'auto'"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>

          </style:text-properties>
        </xsl:when>
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="$level='4'">
          <style:paragraph-properties >
            <xsl:attribute name ="fo:margin-left">
              <xsl:value-of select ="concat(format-number(a:lvl4pPr/@marL div 360000 ,'#.##'),'cm')"/>
            </xsl:attribute>

            <xsl:if test="a:lvl4pPr/@marR">
              <xsl:attribute name ="fo:margin-right">
                <xsl:value-of select ="concat(format-number(a:lvl4pPr/@marR div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>

            </xsl:if>

            <xsl:if test="a:lvl4pPr/@marT">
              <xsl:attribute name ="fo:margin-top">
                <xsl:value-of select ="concat(format-number(a:lvl4pPr/@marT div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:lvl4pPr/@marB">
              <xsl:attribute name ="fo:margin-bottom">
                <xsl:value-of select ="concat(format-number(a:lvl4pPr/@marB div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:lvl4pPr/@indent">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select ="concat(format-number(a:lvl4pPr/@indent div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </style:paragraph-properties>
          <style:text-properties>
            <xsl:choose>
              <xsl:when test="a:lvl4pPr/a:defRPr/a:solidFill/a:srgbClr">
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="concat('#',a:lvl4pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="a:lvl4pPr/a:solidFill/a:schemeClr">
                <xsl:attribute name ="fo:color">
                  <xsl:call-template name ="getColorCode">
                    <xsl:with-param name ="color">
                      <xsl:value-of select="a:lvl4pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                    </xsl:with-param>
                    <xsl:with-param name ="lumMod">
                      <xsl:choose>
                        <xsl:when test="a:lvl4pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod">
                          <xsl:value-of select="a:lvl4pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:with-param>
                    <xsl:with-param name ="lumOff">
                      <xsl:choose>
                        <xsl:when test="a:lvl4pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff">
                          <xsl:value-of select="a:lvl4pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>

                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="'#000000'"/>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>

            <xsl:choose>
              <xsl:when test="a:lvl4pPr/a:defRPr/@sz">
                <xsl:attribute name ="fo:font-size">
                  <xsl:value-of select ="concat(a:lvl4pPr/a:defRPr/@sz div 100 ,'pt')"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:choose>
              <xsl:when test="a:lvl4pPr/a:defRPr/@b='1'">
                <xsl:attribute name ="fo:font-weight">
                  <xsl:value-of select ="'bold'"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="a:lvl4pPr/a:defRPr/@i='1'">
                <xsl:attribute name ="fo:font-style">
                  <xsl:value-of select ="'italic'"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:attribute name ="fo:font-family">
              <xsl:call-template name="slideMasterFontName">
                <xsl:with-param name="fontType">
                  <xsl:value-of select="'minor'"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
            <xsl:attribute name ="style:font-size-asian">
              <xsl:value-of select ="concat(a:lvl4pPr/a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
            <xsl:attribute name ="style:font-size-complex">
              <xsl:value-of select ="concat(a:lvl4pPr/a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
            <xsl:choose>
              <xsl:when test="a:lvl4pPr/a:defRPr/@u='1'">
                <xsl:attribute name ="style:text-underline-style">
                  <xsl:value-of select ="'solid'"></xsl:value-of>
                </xsl:attribute>
                <xsl:attribute name ="style:text-underline-width">
                  <xsl:value-of select ="'auto'"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>

          </style:text-properties>
        </xsl:when>
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="$level='5'">
          <style:paragraph-properties >
            <xsl:attribute name ="fo:margin-left">
              <xsl:value-of select ="concat(format-number(a:lvl5pPr/@marL div 360000 ,'#.##'),'cm')"/>
            </xsl:attribute>

            <xsl:if test="a:lvl5pPr/@marR">
              <xsl:attribute name ="fo:margin-right">
                <xsl:value-of select ="concat(format-number(a:lvl5pPr/@marR div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>

            </xsl:if>

            <xsl:if test="a:lvl5pPr/@marT">
              <xsl:attribute name ="fo:margin-top">
                <xsl:value-of select ="concat(format-number(a:lvl5pPr/@marT div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:lvl5pPr/@marB">
              <xsl:attribute name ="fo:margin-bottom">
                <xsl:value-of select ="concat(format-number(a:lvl5pPr/@marB div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:lvl5pPr/@indent">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select ="concat(format-number(a:lvl5pPr/@indent div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </style:paragraph-properties>
          <style:text-properties>
            <xsl:choose>
              <xsl:when test="a:lvl5pPr/a:defRPr/a:solidFill/a:srgbClr">
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="concat('#',a:lvl5pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="a:lvl5pPr/a:solidFill/a:schemeClr">
                <xsl:attribute name ="fo:color">
                  <xsl:call-template name ="getColorCode">
                    <xsl:with-param name ="color">
                      <xsl:value-of select="a:lvl5pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                    </xsl:with-param>
                    <xsl:with-param name ="lumMod">
                      <xsl:choose>
                        <xsl:when test="a:lvl5pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod">
                          <xsl:value-of select="a:lvl5pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:with-param>
                    <xsl:with-param name ="lumOff">
                      <xsl:choose>
                        <xsl:when test="a:lvl5pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff">
                          <xsl:value-of select="a:lvl5pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>

                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="'#000000'"/>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>

            <xsl:choose>
              <xsl:when test="a:lvl5pPr/a:defRPr/@sz">
                <xsl:attribute name ="fo:font-size">
                  <xsl:value-of select ="concat(a:lvl5pPr/a:defRPr/@sz div 100 ,'pt')"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:choose>
              <xsl:when test="a:lvl5pPr/a:defRPr/@b='1'">
                <xsl:attribute name ="fo:font-weight">
                  <xsl:value-of select ="'bold'"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="a:lvl5pPr/a:defRPr/@i='1'">
                <xsl:attribute name ="fo:font-style">
                  <xsl:value-of select ="'italic'"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:attribute name ="fo:font-family">
              <xsl:call-template name="slideMasterFontName">
                <xsl:with-param name="fontType">
                  <xsl:value-of select="'minor'"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
            <xsl:attribute name ="style:font-size-asian">
              <xsl:value-of select ="concat(a:lvl5pPr/a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
            <xsl:attribute name ="style:font-size-complex">
              <xsl:value-of select ="concat(a:lvl5pPr/a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
            <xsl:choose>
              <xsl:when test="a:lvl5pPr/a:defRPr/@u='1'">
                <xsl:attribute name ="style:text-underline-style">
                  <xsl:value-of select ="'solid'"></xsl:value-of>
                </xsl:attribute>
                <xsl:attribute name ="style:text-underline-width">
                  <xsl:value-of select ="'auto'"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>

          </style:text-properties>
        </xsl:when>
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="$level='6'">
          <style:paragraph-properties >
            <xsl:attribute name ="fo:margin-left">
              <xsl:value-of select ="concat(format-number(a:lvl6pPr/@marL div 360000 ,'#.##'),'cm')"/>
            </xsl:attribute>

            <xsl:if test="a:lvl6pPr/@marR">
              <xsl:attribute name ="fo:margin-right">
                <xsl:value-of select ="concat(format-number(a:lvl6pPr/@marR div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>

            </xsl:if>

            <xsl:if test="a:lvl6pPr/@marT">
              <xsl:attribute name ="fo:margin-top">
                <xsl:value-of select ="concat(format-number(a:lvl6pPr/@marT div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:lvl6pPr/@marB">
              <xsl:attribute name ="fo:margin-bottom">
                <xsl:value-of select ="concat(format-number(a:lvl6pPr/@marB div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:lvl6pPr/@indent">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select ="concat(format-number(a:lvl6pPr/@indent div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </style:paragraph-properties>
          <style:text-properties>
            <xsl:choose>
              <xsl:when test="a:lvl6pPr/a:defRPr/a:solidFill/a:srgbClr">
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="concat('#',a:lvl6pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="a:lvl6pPr/a:solidFill/a:schemeClr">
                <xsl:attribute name ="fo:color">
                  <xsl:call-template name ="getColorCode">
                    <xsl:with-param name ="color">
                      <xsl:value-of select="a:lvl6pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                    </xsl:with-param>
                    <xsl:with-param name ="lumMod">
                      <xsl:choose>
                        <xsl:when test="a:lvl6pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod">
                          <xsl:value-of select="a:lvl6pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:with-param>
                    <xsl:with-param name ="lumOff">
                      <xsl:choose>
                        <xsl:when test="a:lvl6pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff">
                          <xsl:value-of select="a:lvl6pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>

                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="'#000000'"/>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>

            <xsl:choose>
              <xsl:when test="a:lvl6pPr/a:defRPr/@sz">
                <xsl:attribute name ="fo:font-size">
                  <xsl:value-of select ="concat(a:lvl6pPr/a:defRPr/@sz div 100 ,'pt')"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:choose>
              <xsl:when test="a:lvl6pPr/a:defRPr/@b='1'">
                <xsl:attribute name ="fo:font-weight">
                  <xsl:value-of select ="'bold'"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="a:lvl6pPr/a:defRPr/@i='1'">
                <xsl:attribute name ="fo:font-style">
                  <xsl:value-of select ="'italic'"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:attribute name ="fo:font-family">
              <xsl:call-template name="slideMasterFontName">
                <xsl:with-param name="fontType">
                  <xsl:value-of select="'minor'"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
            <xsl:attribute name ="style:font-size-asian">
              <xsl:value-of select ="concat(a:lvl6pPr/a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
            <xsl:attribute name ="style:font-size-complex">
              <xsl:value-of select ="concat(a:lvl6pPr/a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
            <xsl:choose>
              <xsl:when test="a:lvl6pPr/a:defRPr/@u='1'">
                <xsl:attribute name ="style:text-underline-style">
                  <xsl:value-of select ="'solid'"></xsl:value-of>
                </xsl:attribute>
                <xsl:attribute name ="style:text-underline-width">
                  <xsl:value-of select ="'auto'"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>

          </style:text-properties>
        </xsl:when>
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="$level='7'">
          <style:paragraph-properties >
            <xsl:attribute name ="fo:margin-left">
              <xsl:value-of select ="concat(format-number(a:lvl7pPr/@marL div 360000 ,'#.##'),'cm')"/>
            </xsl:attribute>

            <xsl:if test="a:lvl7pPr/@marR">
              <xsl:attribute name ="fo:margin-right">
                <xsl:value-of select ="concat(format-number(a:lvl7pPr/@marR div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>

            </xsl:if>

            <xsl:if test="a:lvl7pPr/@marT">
              <xsl:attribute name ="fo:margin-top">
                <xsl:value-of select ="concat(format-number(a:lvl7pPr/@marT div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:lvl7pPr/@marB">
              <xsl:attribute name ="fo:margin-bottom">
                <xsl:value-of select ="concat(format-number(a:lvl7pPr/@marB div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:lvl7pPr/@indent">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select ="concat(format-number(a:lvl7pPr/@indent div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </style:paragraph-properties>
          <style:text-properties>
            <xsl:choose>
              <xsl:when test="a:lvl7pPr/a:defRPr/a:solidFill/a:srgbClr">
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="concat('#',a:lvl7pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="a:lvl7pPr/a:solidFill/a:schemeClr">
                <xsl:attribute name ="fo:color">
                  <xsl:call-template name ="getColorCode">
                    <xsl:with-param name ="color">
                      <xsl:value-of select="a:lvl7pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                    </xsl:with-param>
                    <xsl:with-param name ="lumMod">
                      <xsl:choose>
                        <xsl:when test="a:lvl7pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod">
                          <xsl:value-of select="a:lvl7pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:with-param>
                    <xsl:with-param name ="lumOff">
                      <xsl:choose>
                        <xsl:when test="a:lvl7pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff">
                          <xsl:value-of select="a:lvl7pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>

                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="'#000000'"/>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>

            <xsl:choose>
              <xsl:when test="a:lvl7pPr/a:defRPr/@sz">
                <xsl:attribute name ="fo:font-size">
                  <xsl:value-of select ="concat(a:lvl7pPr/a:defRPr/@sz div 100 ,'pt')"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:choose>
              <xsl:when test="a:lvl7pPr/a:defRPr/@b='1'">
                <xsl:attribute name ="fo:font-weight">
                  <xsl:value-of select ="'bold'"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="a:lvl7pPr/a:defRPr/@i='1'">
                <xsl:attribute name ="fo:font-style">
                  <xsl:value-of select ="'italic'"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:attribute name ="fo:font-family">
              <xsl:call-template name="slideMasterFontName">
                <xsl:with-param name="fontType">
                  <xsl:value-of select="'minor'"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
            <xsl:attribute name ="style:font-size-asian">
              <xsl:value-of select ="concat(a:lvl7pPr/a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
            <xsl:attribute name ="style:font-size-complex">
              <xsl:value-of select ="concat(a:lvl7pPr/a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
            <xsl:choose>
              <xsl:when test="a:lvl7pPr/a:defRPr/@u='1'">
                <xsl:attribute name ="style:text-underline-style">
                  <xsl:value-of select ="'solid'"></xsl:value-of>
                </xsl:attribute>
                <xsl:attribute name ="style:text-underline-width">
                  <xsl:value-of select ="'auto'"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>

          </style:text-properties>
        </xsl:when>
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="$level='8'">
          <style:paragraph-properties >
            <xsl:attribute name ="fo:margin-left">
              <xsl:value-of select ="concat(format-number(a:lvl8pPr/@marL div 360000 ,'#.##'),'cm')"/>
            </xsl:attribute>

            <xsl:if test="a:lvl8pPr/@marR">
              <xsl:attribute name ="fo:margin-right">
                <xsl:value-of select ="concat(format-number(a:lvl8pPr/@marR div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>

            </xsl:if>

            <xsl:if test="a:lvl8pPr/@marT">
              <xsl:attribute name ="fo:margin-top">
                <xsl:value-of select ="concat(format-number(a:lvl8pPr/@marT div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:lvl8pPr/@marB">
              <xsl:attribute name ="fo:margin-bottom">
                <xsl:value-of select ="concat(format-number(a:lvl8pPr/@marB div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:lvl8pPr/@indent">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select ="concat(format-number(a:lvl8pPr/@indent div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </style:paragraph-properties>
          <style:text-properties>
            <xsl:choose>
              <xsl:when test="a:lvl8pPr/a:defRPr/a:solidFill/a:srgbClr">
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="concat('#',a:lvl8pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="a:lvl8pPr/a:solidFill/a:schemeClr">
                <xsl:attribute name ="fo:color">
                  <xsl:call-template name ="getColorCode">
                    <xsl:with-param name ="color">
                      <xsl:value-of select="a:lvl8pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                    </xsl:with-param>
                    <xsl:with-param name ="lumMod">
                      <xsl:choose>
                        <xsl:when test="a:lvl8pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod">
                          <xsl:value-of select="a:lvl8pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:with-param>
                    <xsl:with-param name ="lumOff">
                      <xsl:choose>
                        <xsl:when test="a:lvl8pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff">
                          <xsl:value-of select="a:lvl8pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>

                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="'#000000'"/>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>

            <xsl:choose>
              <xsl:when test="a:lvl8pPr/a:defRPr/@sz">
                <xsl:attribute name ="fo:font-size">
                  <xsl:value-of select ="concat(a:lvl8pPr/a:defRPr/@sz div 100 ,'pt')"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:choose>
              <xsl:when test="a:lvl8pPr/a:defRPr/@b='1'">
                <xsl:attribute name ="fo:font-weight">
                  <xsl:value-of select ="'bold'"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="a:lvl8pPr/a:defRPr/@i='1'">
                <xsl:attribute name ="fo:font-style">
                  <xsl:value-of select ="'italic'"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:attribute name ="fo:font-family">
              <xsl:call-template name="slideMasterFontName">
                <xsl:with-param name="fontType">
                  <xsl:value-of select="'minor'"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
            <xsl:attribute name ="style:font-size-asian">
              <xsl:value-of select ="concat(a:lvl8pPr/a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
            <xsl:attribute name ="style:font-size-complex">
              <xsl:value-of select ="concat(a:lvl8pPr/a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
            <xsl:choose>
              <xsl:when test="a:lvl8pPr/a:defRPr/@u='1'">
                <xsl:attribute name ="style:text-underline-style">
                  <xsl:value-of select ="'solid'"></xsl:value-of>
                </xsl:attribute>
                <xsl:attribute name ="style:text-underline-width">
                  <xsl:value-of select ="'auto'"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>

          </style:text-properties>
        </xsl:when>
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="$level='9'">
          <style:paragraph-properties >
            <xsl:attribute name ="fo:margin-left">
              <xsl:value-of select ="concat(format-number(a:lvl9pPr/@marL div 360000 ,'#.##'),'cm')"/>
            </xsl:attribute>

            <xsl:if test="a:lvl9pPr/@marR">
              <xsl:attribute name ="fo:margin-right">
                <xsl:value-of select ="concat(format-number(a:lvl9pPr/@marR div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>

            </xsl:if>

            <xsl:if test="a:lvl9pPr/@marT">
              <xsl:attribute name ="fo:margin-top">
                <xsl:value-of select ="concat(format-number(a:lvl9pPr/@marT div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:lvl9pPr/@marB">
              <xsl:attribute name ="fo:margin-bottom">
                <xsl:value-of select ="concat(format-number(a:lvl9pPr/@marB div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:lvl9pPr/@indent">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select ="concat(format-number(a:lvl9pPr/@indent div 360000 ,'#.##'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </style:paragraph-properties>
          <style:text-properties>
            <xsl:choose>
              <xsl:when test="a:lvl9pPr/a:defRPr/a:solidFill/a:srgbClr">
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="concat('#',a:lvl9pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="a:lvl9pPr/a:solidFill/a:schemeClr">
                <xsl:attribute name ="fo:color">
                  <xsl:call-template name ="getColorCode">
                    <xsl:with-param name ="color">
                      <xsl:value-of select="a:lvl9pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                    </xsl:with-param>
                    <xsl:with-param name ="lumMod">
                      <xsl:choose>
                        <xsl:when test="a:lvl9pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod">
                          <xsl:value-of select="a:lvl9pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:with-param>
                    <xsl:with-param name ="lumOff">
                      <xsl:choose>
                        <xsl:when test="a:lvl9pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff">
                          <xsl:value-of select="a:lvl9pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>

                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="'#000000'"/>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>

            <xsl:choose>
              <xsl:when test="a:lvl9pPr/a:defRPr/@sz">
                <xsl:attribute name ="fo:font-size">
                  <xsl:value-of select ="concat(a:lvl9pPr/a:defRPr/@sz div 100 ,'pt')"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:choose>
              <xsl:when test="a:lvl9pPr/a:defRPr/@b='1'">
                <xsl:attribute name ="fo:font-weight">
                  <xsl:value-of select ="'bold'"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="a:lvl9pPr/a:defRPr/@i='1'">
                <xsl:attribute name ="fo:font-style">
                  <xsl:value-of select ="'italic'"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:attribute name ="fo:font-family">
              <xsl:call-template name="slideMasterFontName">
                <xsl:with-param name="fontType">
                  <xsl:value-of select="'minor'"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
            <xsl:attribute name ="style:font-size-asian">
              <xsl:value-of select ="concat(a:lvl9pPr/a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
            <xsl:attribute name ="style:font-size-complex">
              <xsl:value-of select ="concat(a:lvl9pPr/a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
            <xsl:choose>
              <xsl:when test="a:lvl9pPr/a:defRPr/@u='1'">
                <xsl:attribute name ="style:text-underline-style">
                  <xsl:value-of select ="'solid'"></xsl:value-of>
                </xsl:attribute>
                <xsl:attribute name ="style:text-underline-width">
                  <xsl:value-of select ="'auto'"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>

          </style:text-properties>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>

  </xsl:template>
  <xsl:template name ="slideMasterInternalMargins">
    <xsl:param name ="phType"/>
    <xsl:param name ="defType"/>
    <xsl:param name="SlideMasterFile"></xsl:param>
    <xsl:for-each select ="document(concat('ppt/slideMasters/',$SlideMasterFile))//p:nvPr/p:ph[@type = $phType]">
      <xsl:for-each select ="parent::node()/parent::node()/parent::node()">
        <xsl:choose>
          <xsl:when test ="$defType ='tIns'">
            <xsl:if test ="p:txBody/a:bodyPr/@tIns">
              <xsl:attribute name ="fo:padding-top">
                <xsl:value-of select ="concat(format-number(p:txBody/a:bodyPr/@tIns div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test ="$defType ='bIns'">
            <xsl:if test ="p:txBody/a:bodyPr/@bIns">
              <xsl:attribute name ="fo:padding-bottom">
                <xsl:value-of select ="concat(format-number(p:txBody/a:bodyPr/@bIns div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when >
          <xsl:when test ="$defType ='lIns'">
            <xsl:if test ="p:txBody/a:bodyPr/@lIns">
              <xsl:attribute name ="fo:padding-left">
                <xsl:value-of select ="concat(format-number(p:txBody/a:bodyPr/@lIns div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when >
          <xsl:when test ="$defType ='rIns'">
            <xsl:if test ="p:txBody/a:bodyPr/@rIns">
              <xsl:attribute name ="fo:padding-right">
                <xsl:value-of select ="concat(format-number(p:txBody/a:bodyPr/@rIns div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when >
          <xsl:when test ="$defType='anchor'">
            <xsl:if test ="p:txBody/a:bodyPr/@anchor">
              <xsl:attribute name ="draw:textarea-vertical-align">
                <xsl:choose>
                  <!-- Top-->
                  <xsl:when test ="p:txBody/a:bodyPr/@anchor = 't' ">
                    <xsl:value-of select ="'top'"/>
                  </xsl:when>
                  <!-- Middle-->
                  <xsl:when test ="p:txBody/a:bodyPr/@anchor = 'ctr' ">
                    <xsl:value-of select ="'middle'"/>
                  </xsl:when>
                  <!-- bottom-->
                  <xsl:when test ="p:txBody/a:bodyPr/@anchor = 'b' ">
                    <xsl:value-of select ="'bottom'"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if >
          </xsl:when>
          <xsl:when test ="$defType='anchorCtr'">
            <xsl:if test ="p:txBody/a:bodyPr/@anchorCtr" >
              <xsl:attribute name ="draw:textarea-horizontal-align">
                <xsl:if test ="p:txBody/a:bodyPr/@anchorCtr= 1">
                  <xsl:value-of select ="'center'"/>
                </xsl:if>
                <xsl:if test ="p:txBody/a:bodyPr/@anchorCtr= 0">
                  <xsl:value-of select ="'justify'"/>
                </xsl:if>
              </xsl:attribute >
            </xsl:if>
          </xsl:when >
          <xsl:when test ="$defType ='wrap'">
            <xsl:if test ="p:txBody/a:bodyPr/@wrap">
              <xsl:choose >
                <xsl:when test="p:txBody/a:bodyPr/@wrap='none'">
                  <xsl:attribute name ="fo:wrap-option">
                    <xsl:value-of select ="'no-wrap'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test="p:txBody/a:bodyPr/@wrap='square'">
                  <xsl:attribute name ="fo:wrap-option">
                    <xsl:value-of select ="'wrap'"/>
                  </xsl:attribute>
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select ="'0'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="slideMasterTitle">
    <xsl:param name="SlideMasterFile"></xsl:param>

    <xsl:for-each select="document(concat('ppt/slideMasters/',$SlideMasterFile))//p:sp">
      <xsl:choose>
        <xsl:when test="p:nvSpPr/p:nvPr/p:ph/@type='title'">
          <style:graphic-properties>
            <!--for getting Title box back ground color-->
            <xsl:choose>
              <xsl:when test="p:spPr/a:solidFill/a:srgbClr">
                <xsl:attribute name ="draw:fill-color">
                  <xsl:value-of select ="concat('#',p:spPr/a:solidFill/a:srgbClr/@val)"/>
                </xsl:attribute>
                <xsl:attribute name="draw:fill">
                  <xsl:value-of select="'solid'"/>
                </xsl:attribute>
                <!--<xsl:attribute name="draw:fill">
                  <xsl:value-of select="'none'"/>
                </xsl:attribute>-->
              </xsl:when>
              <xsl:when test="p:spPr/a:solidFill/a:schemeClr">
                <xsl:attribute name ="draw:fill-color">
                  <xsl:call-template name ="getColorCode">
                    <xsl:with-param name ="color">
                      <xsl:value-of select="p:spPr/a:solidFill/a:schemeClr/@val"/>
                    </xsl:with-param>
                    <xsl:with-param name ="lumMod">
                      <xsl:choose>
                        <xsl:when test="p:spPr/a:solidFill/a:schemeClr/a:lumMod">
                          <xsl:value-of select="p:spPr/a:solidFill/a:schemeClr/a:lumMod/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:with-param>
                    <xsl:with-param name ="lumOff">
                      <xsl:choose>
                        <xsl:when test="p:spPr/a:solidFill/a:schemeClr/a:lumOff">
                          <xsl:value-of select="p:spPr/a:solidFill/a:schemeClr/a:lumOff/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>

                    </xsl:with-param>
                  </xsl:call-template>

                </xsl:attribute>
                <xsl:attribute name="draw:fill">
                  <xsl:value-of select="'solid'"/>
                </xsl:attribute>
                <!--<xsl:attribute name="draw:fill">
                  <xsl:value-of select="'none'"/>
                </xsl:attribute>-->
              </xsl:when>
              <xsl:otherwise>
                <xsl:attribute name="draw:fill">
                  <xsl:value-of select="'none'"/>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>
           
            <!--for getting border color-->
            <xsl:if test="p:spPr/a:ln/a:solidFill/a:srgbClr">
              <xsl:attribute name ="draw:stroke">
                <xsl:value-of select ="concat('#',p:spPr/a:ln/a:solidFill/a:srgbClr/@val)"/>
              </xsl:attribute>
            </xsl:if>
            <!--Vertical Text alignment-->
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:choose>
                <xsl:when test="p:txBody/a:bodyPr/@anchor='ctr'">
                  <xsl:value-of select ="'middle'"/>
                </xsl:when>
                <xsl:when test="p:txBody/a:bodyPr/@anchor='t'">
                  <xsl:value-of select ="'top'"/>
                </xsl:when>
                <xsl:when test="p:txBody/a:bodyPr/@anchor='b'">
                  <xsl:value-of select ="'bottom'"/>
                </xsl:when>
              </xsl:choose>
            </xsl:attribute>

            <text:list-style>
              <text:list-level-style-bullet text:level="1">

              </text:list-level-style-bullet>
            </text:list-style>
          </style:graphic-properties>

          <style:paragraph-properties>
            <xsl:attribute name ="fo:margin-left">
              <xsl:choose>
                <xsl:when test="p:txBody/a:p/a:pPr/@marL">
                  <xsl:value-of select ="concat(format-number(p:txBody/a:p/a:pPr/@marL div 360000 ,'#.##'),'cm')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="'0cm'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name ="fo:margin-right">
              <xsl:choose>
                <xsl:when test="p:txBody/a:p/a:pPr/@marR">
                  <xsl:value-of select ="concat(format-number(p:txBody/a:p/a:pPr/@marR div 360000 ,'#.##'),'cm')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="'0cm'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name ="fo:margin-top">
              <xsl:choose>
                <xsl:when test="p:txBody/a:p/a:pPr/@marT">
                  <xsl:value-of select ="concat(format-number(p:txBody/a:p/a:pPr/@marT div 360000 ,'#.##'),'cm')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="'0cm'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name ="fo:margin-bottom">
              <xsl:choose>
                <xsl:when test="p:txBody/a:p/a:pPr/@marB">
                  <xsl:value-of select ="concat(format-number(p:txBody/a:p/a:pPr/@marB div 360000 ,'#.##'),'cm')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="'0cm'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
      
            <xsl:attribute name ="fo:text-indent">
              <xsl:choose>
                <xsl:when test="p:txBody/a:p/a:pPr/@indent">
                  <xsl:value-of select ="concat(format-number(p:txBody/a:p/a:pPr/@indent div 360000 ,'#.##'),'cm')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="'0cm'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name ="fo:text-align">
              <xsl:choose>
                <!-- Center Alignment-->
                <xsl:when test ="p:txBody/a:p/a:pPr/@algn ='ctr'">
                  <xsl:value-of select ="'center'"/>
                </xsl:when>
                <!-- Right Alignment-->
                <xsl:when test ="p:txBody/a:p/a:pPr/@algn ='r'">
                  <xsl:value-of select ="'end'"/>
                </xsl:when>
                <!-- Left Alignment-->
                <xsl:when test ="p:txBody/a:p/a:pPr/@algn ='l'">
                  <xsl:value-of select ="'start'"/>
                </xsl:when>
                <!-- Justified Alignment-->
                <xsl:when test ="p:txBody/a:p/a:pPr/@algn ='just'">
                  <xsl:value-of select ="'justified'"/>
                </xsl:when>
                  <xsl:otherwise >
                  <xsl:value-of select ="'center'"/>
                </xsl:otherwise>
                <!-- end-->
              </xsl:choose>
            </xsl:attribute>
          </style:paragraph-properties>

          <!-- for paragraph properties-->
          <!--<style:paragraph-properties  text:enable-numbering="false">
            <xsl:call-template name="getParaGraphStyle"></xsl:call-template>
          </style:paragraph-properties>-->
           
            <!-- for text properties--><!--
          <xsl:call-template name="getTextPropStyle"></xsl:call-template>-->-->
          </xsl:when>
      </xsl:choose>
    </xsl:for-each>
    <xsl:for-each select="document(concat('ppt/slideMasters/',$SlideMasterFile))//p:txStyles/p:titleStyle/a:lvl1pPr">
      <style:text-properties>
        <xsl:choose>
          <xsl:when test="a:defRPr/a:solidFill/a:srgbClr">
            <xsl:attribute name ="fo:color">
              <xsl:value-of select="concat('#',a:defRPr/a:solidFill/a:srgbClr/@val)"></xsl:value-of>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="a:defRPr/a:solidFill/a:schemeClr">
            <xsl:attribute name ="fo:color">
              <xsl:call-template name ="getColorCode">
                <xsl:with-param name ="color">
                  <xsl:value-of select="a:defRPr/a:solidFill/a:schemeClr/@val"/>
                </xsl:with-param>
                <xsl:with-param name ="lumMod">
                  <xsl:choose>
                    <xsl:when test="a:defRPr/a:solidFill/a:schemeClr/a:lumMod">
                      <xsl:value-of select="a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val" />
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="''" />
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:with-param>
                <xsl:with-param name ="lumOff">
                  <xsl:choose>
                    <xsl:when test="a:defRPr/a:solidFill/a:schemeClr/a:lumOff">
                      <xsl:value-of select="a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val" />
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="''" />
                    </xsl:otherwise>
                  </xsl:choose>

                </xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name ="fo:color">
              <xsl:value-of select="'#000000'"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>

        <xsl:choose>
          <xsl:when test="a:defRPr/@sz">
            <xsl:attribute name ="fo:font-size">
              <xsl:value-of select ="concat(a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
          </xsl:when>
        </xsl:choose>
        <xsl:choose>
          <xsl:when test="a:defRPr/@b='1'">
            <xsl:attribute name ="fo:font-weight">
              <xsl:value-of select ="'bold'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="a:defRPr/@i='1'">
            <xsl:attribute name ="fo:font-style">
              <xsl:value-of select ="'italic'"/>
            </xsl:attribute>
          </xsl:when>
        </xsl:choose>
        <xsl:attribute name ="fo:font-family">
          <xsl:if test ="a:defRPr/a:latin/@typeface">
            <xsl:variable name ="typeFaceVal" select ="a:defRPr/a:latin/@typeface"/>
            <xsl:for-each select ="a:defRPr/a:latin/@typeface">
              <xsl:if test ="$typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt'">
                <xsl:call-template name="slideMasterFontName">
                  <xsl:with-param name="fontType">
                    <xsl:value-of select="'major'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:if>
              <!-- Bug 1711910 Fixed,On date 2-06-07,by Vijayeta-->
              <xsl:if test ="not($typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt')">
                <xsl:value-of select ="."/>
              </xsl:if>
            </xsl:for-each>
          </xsl:if>
          <xsl:if test ="not(a:defRPr/a:latin/@typeface)">
            <xsl:call-template name="slideMasterFontName">
              <xsl:with-param name="fontType">
                <xsl:value-of select="'major'"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:if>
        </xsl:attribute>
       
        <xsl:choose>
          <xsl:when test="a:defRPr/@u='1'">
            <xsl:attribute name ="style:text-underline-style">
              <xsl:value-of select ="'solid'"></xsl:value-of>
            </xsl:attribute>
            <xsl:attribute name ="style:text-underline-width">
              <xsl:value-of select ="'auto'"></xsl:value-of>
            </xsl:attribute>
          </xsl:when>
        </xsl:choose>
      </style:text-properties>
    </xsl:for-each>

  </xsl:template>
  <xsl:template name="InsertBulletChar">
    <xsl:param name="character" />
    <xsl:choose>
      <xsl:when test="$character= 'o'">O</xsl:when>
      <xsl:when test="$character = '§'">■</xsl:when>
      <xsl:when test="$character = 'q'">•</xsl:when>
      <xsl:when test="$character = '•'">•</xsl:when>
      <xsl:when test="$character = 'v'">●</xsl:when>
      <xsl:when test="$character = 'Ø'">➢</xsl:when>
      <xsl:when test="$character = 'ü'">✔</xsl:when>
      <xsl:when test="'$character = '">■</xsl:when>
      <xsl:when test="$character = 'o'">○</xsl:when>
      <xsl:when test="'$character = '">➔</xsl:when>
      <xsl:when test="'$character = '">✗</xsl:when>
      <xsl:when test="$character = '-'">–</xsl:when>
      <xsl:when test="$character = '–'">–</xsl:when>
      <xsl:when test="'$character = '">–</xsl:when>
      <xsl:otherwise>•</xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="slideMasterDrawingPage">

    <xsl:for-each select="document('ppt/presentation.xml')//p:sldMasterIdLst/p:sldMasterId">
      <xsl:variable name="sldMasterIdRelation">
        <xsl:value-of select="@r:id"></xsl:value-of>
      </xsl:variable>
      <xsl:variable name="currentPos">
        <xsl:value-of select="position()"/>
      </xsl:variable>

      <!-- Loop thru each slide master.xml-->
      <xsl:for-each select="document('ppt/_rels/presentation.xml.rels')//node()[@Id=$sldMasterIdRelation]">

        <xsl:variable name="slideMasterPath">
          <xsl:value-of select="substring-after(@Target,'/')"/>
        </xsl:variable>

        <style:style style:family="drawing-page" presentation:background-visible="true" presentation:background-objects-visible="true" presentation:display-footer="true" presentation:display-page-number="true" presentation:display-date-time="true">
          <xsl:attribute name="style:name">
            <xsl:value-of select="concat('dp',$currentPos)"/>
          </xsl:attribute>
          <xsl:call-template name="getSlideMasterBGColor">
            <xsl:with-param name="slideMasterFile">
              <xsl:value-of select="$slideMasterPath"/>
            </xsl:with-param>
          </xsl:call-template>
        </style:style>
      </xsl:for-each>



    </xsl:for-each>
    <!--
   <style:style style:name="dp1" style:family="drawing-page">
      <style:drawing-page-properties draw:background-size="border" draw:fill="solid" draw:fill-color="#003366"/>
    </style:style>
    -->
  </xsl:template>
  <xsl:template name="slideMasterStylePage">
    <office:master-styles>
      <draw:layer-set>
        <draw:layer draw:name="layout"/>
        <draw:layer draw:name="background"/>
        <draw:layer draw:name="backgroundobjects"/>
        <draw:layer draw:name="controls"/>
        <draw:layer draw:name="measurelines"/>
      </draw:layer-set>
      <xsl:for-each select="document('ppt/presentation.xml')//p:sldMasterIdLst/p:sldMasterId">
        <xsl:variable name="sldMasterIdRelation">
          <xsl:value-of select="@r:id"></xsl:value-of>
        </xsl:variable>
        <xsl:variable name="currentPos">
          <xsl:value-of select="position()"/>
        </xsl:variable>

        <!-- Loop thru each slide master.xml-->
        <xsl:for-each select="document('ppt/_rels/presentation.xml.rels')//node()[@Id=$sldMasterIdRelation]">
          <style:master-page>
            <xsl:variable name="slideMasterPath">
              <xsl:value-of select="substring-after(@Target,'/')"/>
            </xsl:variable>
            <xsl:variable name="slideMasterName">
              <xsl:value-of select="substring-before($slideMasterPath,'.xml')"/>
            </xsl:variable>

            <xsl:attribute name="style:name">
              <xsl:value-of select="$slideMasterName"/>
            </xsl:attribute>
            <xsl:attribute name="style:page-layout-name">
              <xsl:value-of select="'PM1'"/>
            </xsl:attribute>
            <xsl:attribute name="draw:style-name">
              <xsl:value-of select="concat('dp',$currentPos)"/>
            </xsl:attribute>
            <xsl:for-each select="document(concat('ppt/slideMasters/',$slideMasterPath))//p:sp">
              <xsl:choose>
                <xsl:when test="p:nvSpPr/p:nvPr/p:ph/@type='title'">
                  <draw:frame>
                    <xsl:attribute name ="presentation:style-name">
                      <xsl:value-of select ="concat($slideMasterName,'-title')"/>
                    </xsl:attribute>

                    <xsl:attribute name ="svg:x">
                      <xsl:value-of select ="concat(format-number(p:spPr/a:xfrm/a:off/@x div 360000 ,'#.###'),'cm')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="svg:y">
                      <xsl:value-of select ="concat(format-number(p:spPr/a:xfrm/a:off/@y div 360000 ,'#.###'),'cm')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="svg:width">
                      <xsl:value-of select ="concat(format-number(p:spPr/a:xfrm/a:ext/@cx div 360000 ,'#.###'),'cm')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="svg:height">
                      <xsl:value-of select ="concat(format-number(p:spPr/a:xfrm/a:ext/@cy div 360000 ,'#.###'),'cm')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="presentation:class">
                      <xsl:value-of select="'title'"/>
                    </xsl:attribute>
                    <xsl:attribute name ="presentation:placeholder">
                      <xsl:value-of select="'true'"/>
                    </xsl:attribute>
                    <xsl:attribute name ="draw:layer">
                      <xsl:value-of select="'backgroundobjects'"/>
                    </xsl:attribute>
                    <draw:text-box />
                  </draw:frame>
                </xsl:when>
                <xsl:when test="p:nvSpPr/p:nvPr/p:ph/@type='body'">
                  <draw:frame >
                    <xsl:attribute name ="presentation:style-name">
                      <xsl:value-of select ="concat($slideMasterName,'-outline1')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="svg:x">
                      <xsl:value-of select ="concat(format-number(p:spPr/a:xfrm/a:off/@x div 360000 ,'#.###'),'cm')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="svg:y">
                      <xsl:value-of select ="concat(format-number(p:spPr/a:xfrm/a:off/@y div 360000 ,'#.###'),'cm')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="svg:width">
                      <xsl:value-of select ="concat(format-number(p:spPr/a:xfrm/a:ext/@cx div 360000 ,'#.###'),'cm')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="svg:height">
                      <xsl:value-of select ="concat(format-number(p:spPr/a:xfrm/a:ext/@cy div 360000 ,'#.###'),'cm')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="presentation:class">
                      <xsl:value-of select="'outline'"/>
                    </xsl:attribute>
                    <xsl:attribute name ="presentation:placeholder">
                      <xsl:value-of select="'true'"/>
                    </xsl:attribute>
                    <xsl:attribute name ="draw:layer">
                      <xsl:value-of select="'backgroundobjects'"/>
                    </xsl:attribute>
                    <draw:text-box>
                      
                    <!--<xsl:call-template name ="prad" />-->
                    
                      </draw:text-box>
                  </draw:frame>
                </xsl:when>
                <xsl:when test="p:nvSpPr/p:nvPr/p:ph/@type='dt'">
                  <draw:frame>
                    <!--<xsl:attribute name ="presentation:style-name">
                    <xsl:value-of select ="concat(slideMasterName,'-title')"/>
                  </xsl:attribute>-->
                    <xsl:attribute name ="svg:x">
                      <xsl:value-of select ="concat(format-number(p:spPr/a:xfrm/a:off/@x div 360000 ,'#.###'),'cm')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="svg:y">
                      <xsl:value-of select ="concat(format-number(p:spPr/a:xfrm/a:off/@y div 360000 ,'#.###'),'cm')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="svg:width">
                      <xsl:value-of select ="concat(format-number(p:spPr/a:xfrm/a:ext/@cx div 360000 ,'#.###'),'cm')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="svg:height">
                      <xsl:value-of select ="concat(format-number(p:spPr/a:xfrm/a:ext/@cy div 360000 ,'#.###'),'cm')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="presentation:class">
                      <xsl:value-of select="'date-time'"/>
                    </xsl:attribute>
                    <xsl:attribute name ="presentation:placeholder">
                      <xsl:value-of select="'true'"/>
                    </xsl:attribute>
                    <xsl:attribute name ="draw:layer">
                      <xsl:value-of select="'backgroundobjects'"/>
                    </xsl:attribute>
                    <draw:text-box>
                      <text:p>
                        <presentation:date-time/>
                      </text:p>
                    </draw:text-box>
                  </draw:frame>
                </xsl:when>
                <xsl:when test="p:nvSpPr/p:nvPr/p:ph/@type='ftr'">
                  <draw:frame>
                    <xsl:attribute name ="svg:x">
                      <xsl:value-of select ="concat(format-number(p:spPr/a:xfrm/a:off/@x div 360000 ,'#.###'),'cm')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="svg:y">
                      <xsl:value-of select ="concat(format-number(p:spPr/a:xfrm/a:off/@y div 360000 ,'#.###'),'cm')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="svg:width">
                      <xsl:value-of select ="concat(format-number(p:spPr/a:xfrm/a:ext/@cx div 360000 ,'#.###'),'cm')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="svg:height">
                      <xsl:value-of select ="concat(format-number(p:spPr/a:xfrm/a:ext/@cy div 360000 ,'#.###'),'cm')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="presentation:class">
                      <xsl:value-of select="'footer'"/>
                    </xsl:attribute>
                    <xsl:attribute name ="presentation:placeholder">
                      <xsl:value-of select="'true'"/>
                    </xsl:attribute>
                    <xsl:attribute name ="draw:layer">
                      <xsl:value-of select="'backgroundobjects'"/>
                    </xsl:attribute>
                    <draw:text-box>
                      <text:p>
                        <text:span>
                          <presentation:footer/>
                        </text:span>
                      </text:p>
                    </draw:text-box>
                  </draw:frame>
                </xsl:when>
                <xsl:when test="p:nvSpPr/p:nvPr/p:ph/@type='sldNum'">
                  <draw:frame>
                    <xsl:attribute name ="svg:x">
                      <xsl:value-of select ="concat(format-number(p:spPr/a:xfrm/a:off/@x div 360000 ,'#.###'),'cm')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="svg:y">
                      <xsl:value-of select ="concat(format-number(p:spPr/a:xfrm/a:off/@y div 360000 ,'#.###'),'cm')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="svg:width">
                      <xsl:value-of select ="concat(format-number(p:spPr/a:xfrm/a:ext/@cx div 360000 ,'#.###'),'cm')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="svg:height">
                      <xsl:value-of select ="concat(format-number(p:spPr/a:xfrm/a:ext/@cy div 360000 ,'#.###'),'cm')"/>
                    </xsl:attribute>
                    <xsl:attribute name ="presentation:class">
                      <xsl:value-of select="'page-number'"/>
                    </xsl:attribute>
                    <xsl:attribute name ="presentation:placeholder">
                      <xsl:value-of select="'true'"/>
                    </xsl:attribute>
                    <xsl:attribute name ="draw:layer">
                      <xsl:value-of select="'backgroundobjects'"/>
                    </xsl:attribute>
                    <draw:text-box>
                      <text:p>
                        <text:span>
                          <text:page-number>2</text:page-number>
                        </text:span>
                      </text:p>
                    </draw:text-box>

                  </draw:frame>
                </xsl:when>
              </xsl:choose>
            </xsl:for-each>
          </style:master-page>
        </xsl:for-each>

      </xsl:for-each>


    </office:master-styles>


  </xsl:template>
  <xsl:template name ="slideMasterFontName">
    <xsl:param name ="fontType" />
    <xsl:choose >
      <xsl:when test ="$fontType ='major'">
        <xsl:value-of select ="document('ppt/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme/a:majorFont/a:latin/@typeface"/>
      </xsl:when>
      <xsl:when test ="$fontType ='minor'">
        <xsl:value-of select ="document('ppt/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme/a:minorFont/a:latin/@typeface"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="slideMasterOutlineStyle">
    <xsl:param name="SlideMasterFile"></xsl:param>

    <xsl:for-each select="document(concat('ppt/slideMasters/',$SlideMasterFile))//p:sp">
      <xsl:choose>

        <xsl:when test="p:nvSpPr/p:nvPr/p:ph/@type='body'">
          <style:graphic-properties>
            <xsl:call-template name="getSPBackColor"/>         
         
            <!--Vertical Text alignment-->
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:choose>
                <xsl:when test="p:txBody/a:bodyPr/@anchor='ctr'">
                  <xsl:value-of select ="'middle'"/>
                </xsl:when>
                <xsl:when test="p:txBody/a:bodyPr/@anchor='t'">
                  <xsl:value-of select ="'top'"/>
                </xsl:when>
                <xsl:when test="p:txBody/a:bodyPr/@anchor='b'">
                  <xsl:value-of select ="'bottom'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="'middle'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>

            <text:list-style>
              <text:list-level-style-bullet text:level="1" text:bullet-char="●">
                <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
              </text:list-level-style-bullet>
              <text:list-level-style-bullet text:level="2" text:bullet-char="●">
                <style:list-level-properties text:space-before="0.6cm" text:min-label-width="0.6cm"/>
                <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
              </text:list-level-style-bullet>
              <text:list-level-style-bullet text:level="3" text:bullet-char="●">
                <style:list-level-properties text:space-before="1.2cm" text:min-label-width="0.6cm"/>
                <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
              </text:list-level-style-bullet>
              <text:list-level-style-bullet text:level="4" text:bullet-char="●">
                <style:list-level-properties text:space-before="1.8cm" text:min-label-width="0.6cm"/>
                <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
              </text:list-level-style-bullet>
              <text:list-level-style-bullet text:level="5" text:bullet-char="●">
                <style:list-level-properties text:space-before="2.4cm" text:min-label-width="0.6cm"/>
                <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
              </text:list-level-style-bullet>
              <text:list-level-style-bullet text:level="6" text:bullet-char="●">
                <style:list-level-properties text:space-before="3cm" text:min-label-width="0.6cm"/>
                <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
              </text:list-level-style-bullet>
              <text:list-level-style-bullet text:level="7" text:bullet-char="●">
                <style:list-level-properties text:space-before="3.6cm" text:min-label-width="0.6cm"/>
                <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
              </text:list-level-style-bullet>
              <text:list-level-style-bullet text:level="8" text:bullet-char="●">
                <style:list-level-properties text:space-before="4.2cm" text:min-label-width="0.6cm"/>
                <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
              </text:list-level-style-bullet>
              <text:list-level-style-bullet text:level="9" text:bullet-char="●">
                <style:list-level-properties text:space-before="4.8cm" text:min-label-width="0.6cm"/>
                <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
              </text:list-level-style-bullet>
            </text:list-style>
          </style:graphic-properties>
          <style:paragraph-properties>
            <xsl:attribute name ="fo:margin-left">
              <xsl:choose>
                <xsl:when test="p:txBody/a:p/a:pPr/@marL">
                  <xsl:value-of select ="concat(format-number(p:txBody/a:p/a:pPr/@marL div 360000 ,'#.##'),'cm')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="'0cm'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name ="fo:margin-right">
              <xsl:choose>
                <xsl:when test="p:txBody/a:p/a:pPr/@marR">
                  <xsl:value-of select ="concat(format-number(p:txBody/a:p/a:pPr/@marR div 360000 ,'#.##'),'cm')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="'0cm'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name ="fo:margin-top">
              <xsl:choose>
                <xsl:when test="p:txBody/a:p/a:pPr/@marT">
                  <xsl:value-of select ="concat(format-number(p:txBody/a:p/a:pPr/@marT div 360000 ,'#.##'),'cm')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="'0cm'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name ="fo:margin-bottom">
              <xsl:choose>
                <xsl:when test="p:txBody/a:p/a:pPr/@marB">
                  <xsl:value-of select ="concat(format-number(p:txBody/a:p/a:pPr/@marB div 360000 ,'#.##'),'cm')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="'0cm'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>

            <xsl:attribute name ="fo:text-align">
              <xsl:choose>
                <xsl:when test="p:txBody/a:p/a:pPr/@anchor='ctr'">
                  <xsl:value-of select ="draw:textarea-vertical-align='middle'"/>
                </xsl:when>
                <xsl:when test="p:txBody/a:p/a:pPr/@anchor='t'">
                  <xsl:value-of select ="draw:textarea-vertical-align='top'"/>
                </xsl:when>
                <xsl:when test="p:txBody/a:p/a:pPr/@anchor='b'">
                  <xsl:value-of select ="draw:textarea-vertical-align='bottom'"/>
                </xsl:when>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name ="fo:text-indent">
              <xsl:choose>
                <xsl:when test="p:txBody/a:p/a:pPr/@indent">
                  <xsl:value-of select ="concat(format-number(p:txBody/a:p/a:pPr/@indent div 360000 ,'#.##'),'cm')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="'0cm'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </style:paragraph-properties>

        </xsl:when>

      </xsl:choose>


    </xsl:for-each>

    <xsl:for-each select="document(concat('ppt/slideMasters/',$SlideMasterFile))//p:txStyles/p:titleStyle/a:lvl1pPr">
      <style:text-properties>
        <xsl:choose>
          <xsl:when test="a:defRPr/a:solidFill/a:srgbClr">
            <xsl:attribute name ="fo:color">
              <xsl:value-of select="concat('#',a:defRPr/a:solidFill/a:srgbClr/@val)"></xsl:value-of>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name ="fo:color">
              <xsl:value-of select="'#000000'"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
        <xsl:choose>
          <xsl:when test="a:defRPr/@sz">
            <xsl:attribute name ="fo:font-size">
              <xsl:value-of select ="concat(a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
          </xsl:when>
        </xsl:choose>
        <xsl:choose>
          <xsl:when test="a:defRPr/@b='1'">
            <xsl:attribute name ="fo:font-weight">
              <xsl:value-of select ="'bold'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name ="fo:font-weight">
              <xsl:value-of select ="'normal'"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
        <xsl:attribute name ="fo:font-family">
          <xsl:call-template name="slideMasterFontName">
            <xsl:with-param name="fontType">
              <xsl:value-of select="'major'"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
        <!--<xsl:choose>
          <xsl:when test="a:defRPr/@u='1'">
            <xsl:attribute name ="font-weight">
              <xsl:value-of select ="'bold'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name ="font-weight">
              <xsl:value-of select ="'normal'"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>-->
      </style:text-properties>
    </xsl:for-each>

  </xsl:template>
  <xsl:template name="outlineTextBox">
    <xsl:param name="level"></xsl:param>
    <xsl:param name="text"></xsl:param>
    <xsl:choose>
      <xsl:when test="$level='0'">
        <text:list>
          <text:list-item>
            <text:p>
              <xsl:value-of select="$text"/>
            </text:p>
          </text:list-item>
        </text:list>
      </xsl:when>
      <xsl:when test="$level='1'">
        <text:list>
          <text:list-item>
            <text:list>
              <text:list-item>
                <text:p>
                  <xsl:value-of select="$text"/>
                </text:p>
              </text:list-item>
            </text:list>
          </text:list-item>
        </text:list>
      </xsl:when>
      <xsl:when test="$level='2'">
        <text:list-item>
          <text:list>
            <text:list-item>
              <text:list>
                <text:list-item>
                  <text:p>
                    <xsl:value-of select="$text"/>
                  </text:p>
                </text:list-item>
              </text:list>
            </text:list-item>
          </text:list>
        </text:list-item>
      </xsl:when>
      <xsl:when test="$level='3'">
        <text:list-item>
          <text:list>
            <text:list-item>
              <text:list>
                <text:list-item>
                  <text:list>
                    <text:list-item>
                      <text:p >
                        <xsl:value-of select="$text"/>
                      </text:p>
                    </text:list-item>
                  </text:list>
                </text:list-item>
              </text:list>
            </text:list-item>
          </text:list>
        </text:list-item>
      </xsl:when>
      <xsl:when test="$level='4'">
        <text:list-item>
          <text:list>
            <text:list-item>
              <text:list>
                <text:list-item>
                  <text:list>
                    <text:list-item>
                      <text:list>
                        <text:list-item>
                          <text:p>
                            <xsl:value-of select="$text"/>
                          </text:p>
                        </text:list-item>
                      </text:list>
                    </text:list-item>
                  </text:list>
                </text:list-item>
              </text:list>
            </text:list-item>
          </text:list>
        </text:list-item>
      </xsl:when>
      <xsl:when test="$level='5'">
        <text:list-item>
          <text:list>
            <text:list-item>
              <text:list>
                <text:list-item>
                  <text:list>
                    <text:list-item>
                      <text:list>
                        <text:list-item>
                          <text:list>
                            <text:list-item>
                              <text:p>
                                <xsl:value-of select="$text"/>
                              </text:p>
                            </text:list-item>
                          </text:list>
                        </text:list-item>
                      </text:list>
                    </text:list-item>
                  </text:list>
                </text:list-item>
              </text:list>
            </text:list-item>
          </text:list>
        </text:list-item>
      </xsl:when>
      <xsl:when test="$level='6'">
        <text:list-item>
          <text:list>
            <text:list-item>
              <text:list>
                <text:list-item>
                  <text:list>
                    <text:list-item>
                      <text:list>
                        <text:list-item>
                          <text:list>
                            <text:list-item>
                              <text:list>
                                <text:list-item>
                                  <text:p>
                                    <!--<xsl:value-of select="$text"/>-->
                                  </text:p>
                                </text:list-item>
                              </text:list>
                            </text:list-item>
                          </text:list>
                        </text:list-item>
                      </text:list>
                    </text:list-item>
                  </text:list>
                </text:list-item>
              </text:list>
            </text:list-item>
          </text:list>
        </text:list-item>
      </xsl:when>
      <xsl:when test="$level='7'">
        <text:list-item>
          <text:list>
            <text:list-item>
              <text:list>
                <text:list-item>
                  <text:list>
                    <text:list-item>
                      <text:list>
                        <text:list-item>
                          <text:list>
                            <text:list-item>
                              <text:list>
                                <text:list-item>
                                  <text:list>
                                    <text:list-item>
                                      <text:p>
                                        <!--<xsl:value-of select="$text"/>-->
                                      </text:p>
                                    </text:list-item>
                                  </text:list>
                                </text:list-item>
                              </text:list>
                            </text:list-item>
                          </text:list>
                        </text:list-item>
                      </text:list>
                    </text:list-item>
                  </text:list>
                </text:list-item>
              </text:list>
            </text:list-item>
          </text:list>
        </text:list-item>
      </xsl:when>
      <xsl:when test="$level='8'">
        <text:list-item>
          <text:list>
            <text:list-item>
              <text:list>
                <text:list-item>
                  <text:list>
                    <text:list-item>
                      <text:list>
                        <text:list-item>
                          <text:list>
                            <text:list-item>
                              <text:list>
                                <text:list-item>
                                  <text:list>
                                    <text:list-item>
                                      <text:list>
                                        <text:list-item>
                                          <text:p>
                                            <!--<xsl:value-of select="$text"/>-->
                                          </text:p>
                                        </text:list-item>
                                      </text:list>
                                    </text:list-item>
                                  </text:list>
                                </text:list-item>
                              </text:list>
                            </text:list-item>
                          </text:list>
                        </text:list-item>
                      </text:list>
                    </text:list-item>
                  </text:list>
                </text:list-item>
              </text:list>
            </text:list-item>
          </text:list>
        </text:list-item>
      </xsl:when>
    </xsl:choose>

  </xsl:template>
  <xsl:template name ="prad">
    <xsl:for-each select ="p:txBody/a:p">
      <xsl:variable name ="paraId">
        <xsl:value-of select ="concat('smpara',position())"/>
      </xsl:variable>
      <xsl:variable name ="textId">
        <xsl:value-of select ="concat('smtext',position())"/>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test ="a:pPr/@lvl = 0">
          <text:list text:style-name="L2">
            <text:list-item>
              <text:p >
                <xsl:attribute name ="text:style-name" >
                  <xsl:value-of select ="$paraId"/>
                </xsl:attribute>
                <text:span>
                  <xsl:attribute name ="text:style-name" >
                    <xsl:value-of select ="$textId"/>
                  </xsl:attribute>
                  <!--<xsl:value-of select ="."/>-->
                </text:span>
              </text:p>
            </text:list-item>
          </text:list>
        </xsl:when>
        <xsl:when test ="a:pPr/@lvl = 1">
          <text:list text:style-name="L2">
            <text:list-item>
              <text:list>
                <text:list-item>
                  <text:p >
                    <xsl:attribute name ="text:style-name" >
                      <xsl:value-of select ="$paraId"/>
                    </xsl:attribute>
                    <text:span>
                      <xsl:attribute name ="text:style-name" >
                        <xsl:value-of select ="$textId"/>
                      </xsl:attribute>
                      <!--<xsl:value-of select ="."/>-->
                    </text:span>
                  </text:p>
                </text:list-item>
              </text:list>
            </text:list-item>
          </text:list>
        </xsl:when>
        <xsl:when test ="a:pPr/@lvl = 2">
          <text:list text:style-name="L2">
            <text:list-item>
              <text:list>
                <text:list-item>
                  <text:p >
                    <xsl:attribute name ="text:style-name" >
                      <xsl:value-of select ="$paraId"/>
                    </xsl:attribute>
                    <text:span>
                      <xsl:attribute name ="text:style-name" >
                        <xsl:value-of select ="$textId"/>
                      </xsl:attribute>
                      <!--<xsl:value-of select ="."/>-->
                    </text:span>
                  </text:p>
                </text:list-item>
              </text:list>
            </text:list-item>
          </text:list>
        </xsl:when>
        <xsl:when test ="a:pPr/@lvl = 3">
          <text:list text:style-name="L2">
            <text:list-item>
              <text:list>
                <text:list-item>
                  <text:list>
                    <text:list-item>
                      <text:p >
                        <xsl:attribute name ="text:style-name" >
                          <xsl:value-of select ="$paraId"/>
                        </xsl:attribute>
                        <text:span>
                          <xsl:attribute name ="text:style-name" >
                            <xsl:value-of select ="$textId"/>
                          </xsl:attribute>
                          <!--<xsl:value-of select ="."/>-->
                        </text:span>
                      </text:p>
                    </text:list-item>
                  </text:list>
                </text:list-item>
              </text:list>
            </text:list-item>
          </text:list>

        </xsl:when>
        <xsl:when test ="a:pPr/@lvl = 4">


          <text:list text:style-name="L2">
            <text:list-item>
              <text:list>
                <text:list-item>
                  <text:list>
                    <text:list-item>
                      <text:list>
                        <text:list-item>
                          <text:p>
                            <xsl:attribute name ="text:style-name" >
                              <xsl:value-of select ="$paraId"/>
                            </xsl:attribute>
                            <text:span>
                              <xsl:attribute name ="text:style-name" >
                                <xsl:value-of select ="$textId"/>
                              </xsl:attribute>
                              <!--<xsl:value-of select ="."/>-->
                            </text:span>
                          </text:p>
                        </text:list-item>
                      </text:list>
                    </text:list-item>
                  </text:list>
                </text:list-item>
              </text:list>
            </text:list-item>
          </text:list>

        </xsl:when>
        <xsl:when test ="a:pPr/@lvl = 5">
          <text:list text:style-name="L2">
            <text:list-item>
              <text:list>
                <text:list-item>
                  <text:list>
                    <text:list-item>
                      <text:list>
                        <text:list-item>
                          <text:list>
                            <text:list-item>
                              <text:p >
                                <xsl:attribute name ="text:style-name" >
                                  <xsl:value-of select ="$paraId"/>
                                </xsl:attribute>
                                <text:span>
                                  <xsl:attribute name ="text:style-name" >
                                    <xsl:value-of select ="$textId"/>
                                  </xsl:attribute>
                                  <!--<xsl:value-of select ="."/>-->
                                </text:span>
                              </text:p>
                            </text:list-item>
                          </text:list>
                        </text:list-item>
                      </text:list>
                    </text:list-item>
                  </text:list>
                </text:list-item>
              </text:list>
            </text:list-item>
          </text:list>
        </xsl:when>
        <xsl:when test ="a:pPr/@lvl = 6">
          <text:list text:style-name="L2">
            <text:list-item>
              <text:list>
                <text:list-item>
                  <text:list>
                    <text:list-item>
                      <text:list>
                        <text:list-item>
                          <text:list>
                            <text:list-item>
                              <text:list>
                                <text:list-item>

                                  <text:p >
                                    <xsl:attribute name ="text:style-name" >
                                      <xsl:value-of select ="$paraId"/>
                                    </xsl:attribute>
                                    <text:span>
                                      <xsl:attribute name ="text:style-name" >
                                        <xsl:value-of select ="$textId"/>
                                      </xsl:attribute>
                                      <!--<xsl:value-of select ="."/>-->
                                    </text:span>
                                  </text:p>


                                </text:list-item>
                              </text:list>
                            </text:list-item>
                          </text:list>
                        </text:list-item>
                      </text:list>
                    </text:list-item>
                  </text:list>
                </text:list-item>
              </text:list>
            </text:list-item>
          </text:list>
        </xsl:when>
        <xsl:when test ="a:pPr/@lvl = 7">

          <text:list text:style-name="L2">
            <text:list-item>
              <text:list>
                <text:list-item>
                  <text:list>
                    <text:list-item>
                      <text:list>
                        <text:list-item>
                          <text:list>
                            <text:list-item>
                              <text:list>
                                <text:list-item>
                                  <text:list>
                                    <text:list-item>

                                      <text:p >
                                        <xsl:attribute name ="text:style-name" >
                                          <xsl:value-of select ="$paraId"/>
                                        </xsl:attribute>
                                        <text:span>
                                          <xsl:attribute name ="text:style-name" >
                                            <xsl:value-of select ="$textId"/>
                                          </xsl:attribute>
                                          <!--<xsl:value-of select ="."/>-->
                                        </text:span>
                                      </text:p>

                                    </text:list-item>
                                  </text:list>
                                </text:list-item>
                              </text:list>
                            </text:list-item>
                          </text:list>
                        </text:list-item>
                      </text:list>
                    </text:list-item>
                  </text:list>
                </text:list-item>
              </text:list>
            </text:list-item>
          </text:list>
        </xsl:when>
        <xsl:when test ="a:pPr/@lvl = 8">
          <text:list text:style-name="L2">
            <text:list-item>
              <text:list>
                <text:list-item>
                  <text:list>
                    <text:list-item>
                      <text:list>
                        <text:list-item>
                          <text:list>
                            <text:list-item>
                              <text:list>
                                <text:list-item>
                                  <text:list>
                                    <text:list-item>
                                      <text:list>
                                        <text:list-item>
                                          <text:p>

                                            <xsl:attribute name ="text:style-name" >
                                              <xsl:value-of select ="$paraId"/>
                                            </xsl:attribute>
                                            <text:span>
                                              <xsl:attribute name ="text:style-name" >
                                                <xsl:value-of select ="$textId"/>
                                              </xsl:attribute>
                                              <!--<xsl:value-of select ="."/>-->
                                            </text:span>
                                          </text:p>

                                        </text:list-item>
                                      </text:list>
                                    </text:list-item>
                                  </text:list>
                                </text:list-item>
                              </text:list>
                            </text:list-item>
                          </text:list>
                        </text:list-item>
                      </text:list>
                    </text:list-item>
                  </text:list>
                </text:list-item>
              </text:list>
            </text:list-item>
          </text:list>
        </xsl:when>

      </xsl:choose>

    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="TextParagraphProp">

    <text:list-style style:name="L2">
      <text:list-level-style-bullet text:level="1" text:bullet-char="●">
        <style:list-level-properties text:space-before="0.3cm" text:min-label-width="0.9cm" />
        <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%" />
      </text:list-level-style-bullet>
      <text:list-level-style-bullet text:level="2" text:bullet-char="–">
        <style:list-level-properties text:space-before="1.6cm" text:min-label-width="0.8cm" />
        <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="75%" />
      </text:list-level-style-bullet>
      <text:list-level-style-bullet text:level="3" text:bullet-char="●">
        <style:list-level-properties text:space-before="3cm" text:min-label-width="0.6cm" />
        <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%" />
      </text:list-level-style-bullet>
      <text:list-level-style-bullet text:level="4" text:bullet-char="–">
        <style:list-level-properties text:space-before="4.2cm" text:min-label-width="0.6cm" />
        <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="75%" />
      </text:list-level-style-bullet>
      <text:list-level-style-bullet text:level="5" text:bullet-char="●">
        <style:list-level-properties text:space-before="5.4cm" text:min-label-width="0.6cm" />
        <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%" />
      </text:list-level-style-bullet>
      <text:list-level-style-bullet text:level="6" text:bullet-char="●">
        <style:list-level-properties text:space-before="6.6cm" text:min-label-width="0.6cm" />
        <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%" />
      </text:list-level-style-bullet>
      <text:list-level-style-bullet text:level="7" text:bullet-char="●">
        <style:list-level-properties text:space-before="7.8cm" text:min-label-width="0.6cm" />
        <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%" />
      </text:list-level-style-bullet>
      <text:list-level-style-bullet text:level="8" text:bullet-char="●">
        <style:list-level-properties text:space-before="9cm" text:min-label-width="0.6cm" />
        <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%" />
      </text:list-level-style-bullet>
      <text:list-level-style-bullet text:level="9" text:bullet-char="●">
        <style:list-level-properties text:space-before="10.2cm" text:min-label-width="0.6cm" />
        <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%" />
      </text:list-level-style-bullet>
    </text:list-style>

    <xsl:for-each select ="document('ppt/slideMasters/slideMaster1.xml')//p:txStyles/p:bodyStyle">
      <xsl:for-each select ="node()">

        <style:style style:family="paragraph">
          <xsl:attribute name ="style:name">
            <xsl:value-of select ="concat('smpara',position())"/>
          </xsl:attribute>
          <style:paragraph-properties >
            <xsl:if test ="@marL">
              <xsl:attribute name ="fo:margin-left">
                <xsl:value-of select ="format-number(@marL div 360000,'#.##')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="@indent">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select ="format-number(@indent div 360000,'#.##')"/>
              </xsl:attribute>
            </xsl:if>
          </style:paragraph-properties >
        </style:style>
        <style:style style:family="text">
          <xsl:attribute name ="style:name">
            <xsl:value-of select ="concat('smtext',position())"/>
          </xsl:attribute>
          <style:text-properties >
            <xsl:attribute name ="fo:font-size">
              <xsl:value-of select ="a:defRPr/@sz div 100"/>
            </xsl:attribute>
            <xsl:if test ="a:defRPr/a:solidFill/a:schemeClr">
              <xsl:attribute name ="fo:color">
                <xsl:call-template name="getColorCode">
                  <xsl:with-param name="color">
                    <xsl:value-of select="a:defRPr/a:solidFill/a:schemeClr/@val" />
                  </xsl:with-param>
                  <xsl:with-param name="lumMod">
                    <xsl:value-of select="a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val" />
                  </xsl:with-param>
                  <xsl:with-param name="lumOff">
                    <xsl:value-of select="a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val" />
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="a:defRPr/a:solidFill/a:srgbClr">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="concat('#',a:defRPr/a:solidFill/a:srgbClr/@val)"/>
              </xsl:attribute>
            </xsl:if>
          </style:text-properties >
        </style:style>
      </xsl:for-each>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="getSlideMasterBGColor">
    <xsl:param name="slideMasterFile"></xsl:param>
    <xsl:for-each select="document(concat('ppt/slideMasters/',$slideMasterFile))//p:cSld/p:bg">
      <xsl:choose>
        <xsl:when test="p:bgPr/a:solidFill/a:srgbClr">
          <xsl:attribute name="draw:fill-color">
            <xsl:value-of select="concat('#',p:bgPr/a:solidFill/a:srgbClr/@val)"/>
          </xsl:attribute>
          <xsl:attribute name="draw:fill">
            <xsl:value-of select="'solid'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="p:bgRef/a:solidFill/a:srgbClr">
          <xsl:attribute name="draw:fill-color">
            <xsl:value-of select="concat('#',p:bgRef/a:solidFill/a:srgbClr/@val)"/>
          </xsl:attribute>
          <xsl:attribute name="draw:fill">
            <xsl:value-of select="'solid'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="a:solidFill/a:srgbClr">
          <xsl:attribute name="draw:fill-color">
            <xsl:value-of select="concat('#',a:solidFill/a:srgbClr/@val)"/>
          </xsl:attribute>
          <xsl:attribute name="draw:fill">
            <xsl:value-of select="'solid'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="p:bgRef/a:schemeClr">
          <xsl:attribute name="draw:fill-color">
            <xsl:call-template name ="getColorCode">
              <xsl:with-param name ="color">
                <xsl:value-of select="p:bgRef/a:schemeClr/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumMod">
                <xsl:value-of select="p:bgRef/a:schemeClr/a:lumMod/@val" />
              </xsl:with-param>
              <xsl:with-param name ="lumOff">
                <xsl:value-of select="p:bgRef/a:schemeClr/a:lumOff/@val" />
              </xsl:with-param>

            </xsl:call-template>
          </xsl:attribute>

          <xsl:attribute name="draw:fill">
            <xsl:value-of select="'solid'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="p:bgPr/a:solidFill/a:schemeClr">
          <xsl:attribute name="draw:fill-color">
            <xsl:call-template name ="getColorCode">
              <xsl:with-param name ="color">
                <xsl:value-of select="p:bgPr/a:solidFill/a:schemeClr/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumMod">
                <xsl:value-of select="p:bgPr/a:solidFill/a:schemeClr/a:lumMod/@val" />
              </xsl:with-param>
              <xsl:with-param name ="lumOff">
                <xsl:value-of select="p:bgPr/a:solidFill/a:schemeClr/a:lumOff/@val" />
              </xsl:with-param>

            </xsl:call-template>
          </xsl:attribute>

          <xsl:attribute name="draw:fill">
            <xsl:value-of select="'solid'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="p:bgRef/a:solidFill/a:schemeClr">
          <xsl:attribute name="draw:fill-color">
            <xsl:call-template name ="getColorCode">
              <xsl:with-param name ="color">
                <xsl:value-of select="p:bgRef/a:solidFill/a:schemeClr/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumMod">
                <xsl:value-of select="p:bgRef/a:solidFill/a:schemeClr/a:lumMod/@val" />
              </xsl:with-param>
              <xsl:with-param name ="lumOff">
                <xsl:value-of select="p:bgRef/a:solidFill/a:schemeClr/a:lumOff/@val" />
              </xsl:with-param>

            </xsl:call-template>
          </xsl:attribute>

          <xsl:attribute name="draw:fill">
            <xsl:value-of select="'solid'"/>
          </xsl:attribute>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="Outlines">
    <xsl:param name="level"></xsl:param>
    <xsl:for-each select=".">
      <style:paragraph-properties>
        <xsl:choose>
          <xsl:when test="$level='1'">
            <xsl:attribute name ="text:enable-numbering">
              <xsl:value-of select ="'true'"/>
            </xsl:attribute>
             </xsl:when>
        </xsl:choose>
        <xsl:attribute name ="fo:margin-left">
          <xsl:value-of select ="concat(format-number(./@marL div 360000 ,'#.##'),'cm')"/>
        </xsl:attribute>
        <xsl:if test="./@marR">
          <xsl:attribute name ="fo:margin-right">
            <xsl:value-of select ="concat(format-number(@marR div 360000 ,'#.##'),'cm')"/>
          </xsl:attribute>

        </xsl:if>
        <xsl:if test="./@marT">
          <xsl:attribute name ="fo:margin-top">
            <xsl:value-of select ="concat(format-number(@marT div 360000 ,'#.##'),'cm')"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="./@marB">
          <xsl:attribute name ="fo:margin-bottom">
            <xsl:value-of select ="concat(format-number(@marB div 360000 ,'#.##'),'cm')"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="./@indent">
          <xsl:attribute name ="fo:text-indent">
            <xsl:value-of select ="concat(format-number(@indent div 360000 ,'#.##'),'cm')"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:attribute name ="fo:text-align">
          <xsl:choose>
            <!-- Center Alignment-->
            <xsl:when test ="./@algn ='ctr'">
              <xsl:value-of select ="'center'"/>
            </xsl:when>
            <!-- Right Alignment-->
            <xsl:when test ="./@algn ='r'">
              <xsl:value-of select ="'end'"/>
            </xsl:when>
            <!-- Left Alignment-->
            <xsl:when test ="./@algn ='l'">
              <xsl:value-of select ="'start'"/>
            </xsl:when>
            <!-- Justified Alignment-->
            <xsl:when test ="./@algn ='just'">
              <xsl:value-of select ="'justified'"/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select ="'center'"/>
            </xsl:otherwise>
            <!-- end-->
          </xsl:choose>
        </xsl:attribute>
      </style:paragraph-properties>
      <style:text-properties>
            <xsl:choose>
              <xsl:when test="./a:defRPr/a:solidFill/a:srgbClr">
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="concat('#',./a:defRPr/a:solidFill/a:srgbClr/@val)"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="./a:solidFill/a:schemeClr">
                <xsl:attribute name ="fo:color">
                  <xsl:call-template name ="getColorCode">
                    <xsl:with-param name ="color">
                      <xsl:value-of select="./a:defRPr/a:solidFill/a:schemeClr/@val"/>
                    </xsl:with-param>
                    <xsl:with-param name ="lumMod">
                      <xsl:choose>
                        <xsl:when test="./a:defRPr/a:solidFill/a:schemeClr/a:lumMod">
                          <xsl:value-of select="./a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:with-param>
                    <xsl:with-param name ="lumOff">
                      <xsl:choose>
                        <xsl:when test="./a:defRPr/a:solidFill/a:schemeClr/a:lumOff">
                          <xsl:value-of select="./a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val" />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="''" />
                        </xsl:otherwise>
                      </xsl:choose>

                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="'#000000'"/>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>

            <xsl:choose>
              <xsl:when test="./a:defRPr/@sz">
                <xsl:attribute name ="fo:font-size">
                  <xsl:value-of select ="concat(./a:defRPr/@sz div 100 ,'pt')"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:choose>
              <xsl:when test="./a:defRPr/@b='1'">
                <xsl:attribute name ="fo:font-weight">
                  <xsl:value-of select ="'bold'"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="./a:defRPr/@i='1'">
                <xsl:attribute name ="fo:font-style">
                  <xsl:value-of select ="'italic'"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:attribute name ="fo:font-family">
              <xsl:if test ="./a:defRPr/a:latin/@typeface">
                <xsl:variable name ="typeFaceVal" select ="./a:defRPr/a:latin/@typeface"/>
                <xsl:for-each select ="./a:defRPr/a:latin/@typeface">
                  <xsl:if test ="$typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt'">
                    <xsl:call-template name="slideMasterFontName">
                      <xsl:with-param name="fontType">
                        <xsl:value-of select="'minor'"/>
                      </xsl:with-param>
                    </xsl:call-template>
                  </xsl:if>
                  <!-- Bug 1711910 Fixed,On date 2-06-07,by Vijayeta-->
                  <xsl:if test ="not($typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt')">
                    <xsl:value-of select ="."/>
                  </xsl:if>
                </xsl:for-each>
              </xsl:if>
              <xsl:if test ="not(./a:defRPr/a:latin/@typeface)">
                <xsl:call-template name="slideMasterFontName">
                  <xsl:with-param name="fontType">
                    <xsl:value-of select="'minor'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:if>
            </xsl:attribute>
            <xsl:attribute name ="style:font-size-asian">
              <xsl:value-of select ="concat(./a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
            <xsl:attribute name ="style:font-size-complex">
              <xsl:value-of select ="concat(./a:defRPr/@sz div 100 ,'pt')"/>
            </xsl:attribute>
            <xsl:choose>
              <xsl:when test="./a:defRPr/@u='1'">
                <xsl:attribute name ="style:text-underline-style">
                  <xsl:value-of select ="'solid'"></xsl:value-of>
                </xsl:attribute>
                <xsl:attribute name ="style:text-underline-width">
                  <xsl:value-of select ="'auto'"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>

          </style:text-properties>
       
      
    </xsl:for-each>

  </xsl:template>
  <xsl:template name="getSPBackColor">
    <!--for getting sub Title box back ground color-->
    <xsl:for-each select=".">
    <xsl:choose>
      <xsl:when test="./p:spPr/a:solidFill/a:srgbClr">
        <xsl:attribute name ="draw:fill-color">
          <xsl:value-of select ="concat('#',./p:spPr/a:solidFill/a:srgbClr/@val)"/>
        </xsl:attribute>
        <xsl:attribute name="draw:fill">
          <xsl:value-of select="'solid'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="draw:fill">
          <xsl:value-of select="'none'"/>
        </xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:if test="./p:spPr/a:solidFill/a:schemeClr">
      <xsl:attribute name ="draw:fill-color">
        <xsl:call-template name ="getColorCode">
          <xsl:with-param name ="color">
            <xsl:value-of select="./p:spPr/a:solidFill/a:schemeClr/@val"/>
          </xsl:with-param>
          <xsl:with-param name ="lumMod">
            <xsl:choose>
              <xsl:when test="./p:spPr/a:solidFill/a:schemeClr/a:lumMod">
                <xsl:value-of select="./p:spPr/a:solidFill/a:schemeClr/a:lumMod/@val" />
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="''" />
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name ="lumOff">
            <xsl:choose>
              <xsl:when test="./p:spPr/a:solidFill/a:schemeClr/a:lumOff">
                <xsl:value-of select="./p:spPr/a:solidFill/a:schemeClr/a:lumOff/@val" />
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="''" />
              </xsl:otherwise>
            </xsl:choose>

          </xsl:with-param>
        </xsl:call-template>

      </xsl:attribute>
      <xsl:attribute name="draw:fill">
        <xsl:value-of select="'solid'"/>
      </xsl:attribute>
    </xsl:if>
    <!--for getting border color-->
    <xsl:if test="./p:spPr/a:ln/a:solidFill/a:srgbClr">
      <xsl:attribute name ="draw:stroke">
        <xsl:value-of select ="concat('#',./p:spPr/a:ln/a:solidFill/a:srgbClr/@val)"/>
      </xsl:attribute>
    </xsl:if>
    </xsl:for-each>
  </xsl:template>
</xsl:stylesheet>
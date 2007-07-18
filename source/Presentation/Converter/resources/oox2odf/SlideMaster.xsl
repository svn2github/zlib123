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
  <xsl:import href="content.xsl"/>

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
          <text:list-level-style-bullet text:level="1" text:bullet-char="?">
            <style:list-level-properties text:space-before="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="2" text:bullet-char="?">
            <style:list-level-properties text:space-before="0.6cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="3" text:bullet-char="?">
            <style:list-level-properties text:space-before="1.2cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="4" text:bullet-char="?">
            <style:list-level-properties text:space-before="1.8cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="5" text:bullet-char="?">
            <style:list-level-properties text:space-before="2.4cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="6" text:bullet-char="?">
            <style:list-level-properties text:space-before="3cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="7" text:bullet-char="?">
            <style:list-level-properties text:space-before="3.6cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="8" text:bullet-char="?">
            <style:list-level-properties text:space-before="4.2cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="9" text:bullet-char="?">
            <style:list-level-properties text:space-before="4.8cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="10" text:bullet-char="?">
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
          <text:list-level-style-bullet text:level="1" text:bullet-char="?">
            <style:list-level-properties text:space-before="0.2cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="2" text:bullet-char="?">
            <style:list-level-properties text:space-before="0.6cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="3" text:bullet-char="?">
            <style:list-level-properties text:space-before="1.2cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="4" text:bullet-char="?">
            <style:list-level-properties text:space-before="1.8cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="5" text:bullet-char="?">
            <style:list-level-properties text:space-before="2.4cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="6" text:bullet-char="?">
            <style:list-level-properties text:space-before="3cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="7" text:bullet-char="?">
            <style:list-level-properties text:space-before="3.6cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="8" text:bullet-char="?">
            <style:list-level-properties text:space-before="4.2cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="9" text:bullet-char="?">
            <style:list-level-properties text:space-before="4.8cm" text:min-label-width="0.6cm"/>
            <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
          </text:list-level-style-bullet>
          <text:list-level-style-bullet text:level="10" text:bullet-char="?">
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
        <style:style>
          <xsl:attribute name="style:name">
            <xsl:value-of select="concat($slideMasterName,'-backgroundobjects')"/>
          </xsl:attribute>
          <xsl:attribute name="style:family">
            <xsl:value-of select="'presentation'"/>
          </xsl:attribute>
          <style:graphic-properties draw:shadow="hidden" draw:shadow-offset-x="0.3cm" draw:shadow-offset-y="0.3cm" draw:shadow-color="#808080"/>
        </style:style>
        <!--style for drawing page-->
        <xsl:call-template name="slideMasterDrawingPage"/>
        <xsl:for-each select="document(concat('ppt/slideMasters/',$slideMasterPath))//p:sp">
          <xsl:choose>
            <xsl:when test="p:nvSpPr/p:nvPr/p:ph/@type='title' or p:nvSpPr/p:nvPr/p:ph/@type='ctrTitle'">
              <!-- style for Title-->
              <style:style>
                <xsl:attribute name="style:name">
                  <xsl:value-of select="concat($slideMasterName,'-title')"/>
                </xsl:attribute>
                <xsl:attribute name ="style:family">
                  <xsl:value-of select ="'presentation'"/>
                </xsl:attribute>
                <xsl:call-template name="TitleStyle">
                  <xsl:with-param name="SlideMasterFile">
                    <xsl:value-of select="$slideMasterPath"/>
                  </xsl:with-param>
                </xsl:call-template>
              </style:style>
            </xsl:when>
            <xsl:when test="p:nvSpPr/p:nvPr/p:ph/@type='body'">
              <!-- style for sub-Title-->
              <style:style>
                <xsl:attribute name="style:name">
                  <xsl:value-of select="concat($slideMasterName,'-subtitle')"/>
                </xsl:attribute>
                <xsl:attribute name ="style:family">
                  <xsl:value-of select ="'presentation'"/>
                </xsl:attribute>
                <xsl:call-template name="SubtitleStyle">
                  <xsl:with-param name="SlideMasterFile">
                    <xsl:value-of select="$slideMasterPath"/>
                  </xsl:with-param>
                </xsl:call-template>

              </style:style>
              <!-- style for Outline 1-->
              <style:style>
                <xsl:attribute name="style:name">
                  <xsl:value-of select="concat($slideMasterName,'-outline1')"/>
                </xsl:attribute>
                <xsl:attribute name ="style:family">
                  <xsl:value-of select ="'presentation'"/>
                </xsl:attribute>
                <style:graphic-properties draw:stroke="none">
                  <xsl:call-template name="TitleSubTitleGraphicProperty">
                    <xsl:with-param name="blnSubTitle">
                      <xsl:value-of select="'false'"/>
                    </xsl:with-param>
                  </xsl:call-template>
                  <text:list-style>
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
            </xsl:when>

          </xsl:choose>
        </xsl:for-each>
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
      </xsl:for-each>
    </xsl:for-each>
  </xsl:template>

  <!-- @@Get bullet chars and level from slide master 
This tmplate will get the bullet level and charactor and Marginleft-->
  <xsl:template name ="slideMasterBullett">
    <xsl:param name="SlideMasterFile" />
    <xsl:param name="levelNo"></xsl:param>
    <xsl:param name="pos"></xsl:param>
    <xsl:variable name="var_SMrelationFile">
      <xsl:value-of select="concat('ppt/slideMasters/_rels/',$SlideMasterFile,'.rels')"/>
    </xsl:variable>
    <xsl:for-each select ="document(concat('ppt/slideMasters/',$SlideMasterFile))//p:txStyles/p:bodyStyle">
      <xsl:choose>
        <xsl:when test="$levelNo='0'">
          <xsl:if test ="a:lvl1pPr/a:buChar">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:call-template name="InsertBulletChar">
                  <xsl:with-param name="character">
                    <xsl:value-of select ="a:lvl1pPr/a:buChar/@char"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl1pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl1pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl1pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl1pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl1pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl1pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl1pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'1'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
          <xsl:if test ="a:lvl1pPr/a:buAutoNum">
            <text:list-level-style-number>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:for-each select="./a:lvl1pPr">
                <xsl:call-template name="OutlineNumbering"></xsl:call-template>
              </xsl:for-each>
              <style:list-level-properties>
                <xsl:if test ="./a:lvl1pPr/@indent">
                  <xsl:attribute name ="text:min-label-width">
                    <xsl:value-of select="concat(format-number(./a:lvl1pPr/@indent div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if >
                <xsl:if test ="./a:lvl1pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl1pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl1pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'1'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-number>
          </xsl:if>
          <xsl:if test ="a:lvl1pPr/a:buBlip">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:value-of select="'●'"/>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl1pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl1pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl1pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl1pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl1pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl1pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl1pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'1'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
          <xsl:if test="not(a:lvl1pPr/a:buAutoNum) and not(a:lvl1pPr/a:buChar) and not(a:lvl1pPr/a:buBlip)">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:call-template name="InsertBulletChar">
                  <xsl:with-param name="character">
                    <xsl:value-of select ="a:lvl1pPr/a:buChar/@char"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl1pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl1pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl1pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl1pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl1pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl1pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl1pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'1'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
        </xsl:when>
        <xsl:when test="$levelNo='1'">
          <xsl:if test ="a:lvl2pPr/a:buChar">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:call-template name="InsertBulletChar">
                  <xsl:with-param name="character">
                    <xsl:value-of select ="a:lvl2pPr/a:buChar/@char"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl2pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl2pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl2pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl2pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl2pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl2pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl2pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'2'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
          <xsl:if test ="a:lvl2pPr/a:buAutoNum">
            <text:list-level-style-number>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:for-each select="./a:lvl2pPr">
                <xsl:call-template name="OutlineNumbering"></xsl:call-template>
              </xsl:for-each>
              <style:list-level-properties>
                <xsl:if test ="./a:lvl2pPr/@indent">
                  <xsl:attribute name ="text:min-label-width">
                    <xsl:value-of select="concat(format-number(./a:lvl2pPr/@indent div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if >
                <xsl:if test ="./a:lvl2pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl2pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl2pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'2'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-number>
          </xsl:if>
          <xsl:if test ="a:lvl2pPr/a:buBlip">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:value-of select="'●'"/>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl2pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl2pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl2pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl2pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl2pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl2pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl2pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'2'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
          <xsl:if test="not(a:lvl2pPr/a:buAutoNum) and not(a:lvl2pPr/a:buChar) and not(a:lvl2pPr/a:buBlip)">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:call-template name="InsertBulletChar">
                  <xsl:with-param name="character">
                    <xsl:value-of select ="a:lvl2pPr/a:buChar/@char"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl2pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl2pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl2pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl2pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl2pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl2pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl2pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'2'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
        </xsl:when>
        <xsl:when test="$levelNo='2'">
          <xsl:if test ="a:lvl3pPr/a:buChar">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:call-template name="InsertBulletChar">
                  <xsl:with-param name="character">
                    <xsl:value-of select ="a:lvl3pPr/a:buChar/@char"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl3pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl3pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl3pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl3pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl3pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl3pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl3pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'3'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
          <xsl:if test ="a:lvl3pPr/a:buAutoNum">
            <text:list-level-style-number>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:for-each select="./a:lvl3pPr">
                <xsl:call-template name="OutlineNumbering"></xsl:call-template>
              </xsl:for-each>
              <style:list-level-properties>
                <xsl:if test ="./a:lvl3pPr/@indent">
                  <xsl:attribute name ="text:min-label-width">
                    <xsl:value-of select="concat(format-number(./a:lvl3pPr/@indent div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if >
                <xsl:if test ="./a:lvl3pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl3pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl3pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'3'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-number>
          </xsl:if>
          <xsl:if test ="a:lvl3pPr/a:buBlip">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:value-of select="'●'"/>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl3pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl3pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl3pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl3pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl3pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl3pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl3pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'3'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
          <xsl:if test="not(a:lvl3pPr/a:buAutoNum) and not(a:lvl3pPr/a:buChar) and not(a:lvl3pPr/a:buBlip)">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:call-template name="InsertBulletChar">
                  <xsl:with-param name="character">
                    <xsl:value-of select ="a:lvl3pPr/a:buChar/@char"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl3pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl3pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl3pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl3pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl3pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl3pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl3pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'3'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
        </xsl:when>
        <xsl:when test="$levelNo='3'">
          <xsl:if test ="a:lvl4pPr/a:buChar">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:call-template name="InsertBulletChar">
                  <xsl:with-param name="character">
                    <xsl:value-of select ="a:lvl4pPr/a:buChar/@char"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl4pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl4pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl4pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl4pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl4pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl4pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl4pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'4'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
          <xsl:if test ="a:lvl4pPr/a:buAutoNum">
            <text:list-level-style-number>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:for-each select="./a:lvl4pPr">
                <xsl:call-template name="OutlineNumbering"></xsl:call-template>
              </xsl:for-each>
              <style:list-level-properties>
                <xsl:if test ="./a:lvl4pPr/@indent">
                  <xsl:attribute name ="text:min-label-width">
                    <xsl:value-of select="concat(format-number(./a:lvl4pPr/@indent div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if >
                <xsl:if test ="./a:lvl4pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl4pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl4pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'4'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-number>
          </xsl:if>
          <xsl:if test ="a:lvl4pPr/a:buBlip">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:value-of select="'●'"/>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl4pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl4pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl4pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl4pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl4pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl4pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl4pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'4'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
          <xsl:if test="not(a:lvl4pPr/a:buAutoNum) and not(a:lvl4pPr/a:buChar) and not(a:lvl4pPr/a:buBlip)">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:call-template name="InsertBulletChar">
                  <xsl:with-param name="character">
                    <xsl:value-of select ="a:lvl4pPr/a:buChar/@char"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl4pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl4pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl4pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl4pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl4pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl4pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl4pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'4'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
        </xsl:when>
        <xsl:when test="$levelNo='4'">
          <xsl:if test ="a:lvl5pPr/a:buChar">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:call-template name="InsertBulletChar">
                  <xsl:with-param name="character">
                    <xsl:value-of select ="a:lvl5pPr/a:buChar/@char"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl5pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl5pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl5pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl5pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl5pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl5pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl5pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'5'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
          <xsl:if test ="a:lvl5pPr/a:buAutoNum">
            <text:list-level-style-number>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:for-each select="./a:lvl5pPr">
                <xsl:call-template name="OutlineNumbering"></xsl:call-template>
              </xsl:for-each>
              <style:list-level-properties>
                <xsl:if test ="./a:lvl5pPr/@indent">
                  <xsl:attribute name ="text:min-label-width">
                    <xsl:value-of select="concat(format-number(./a:lvl5pPr/@indent div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if >
                <xsl:if test ="./a:lvl5pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl5pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl5pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'5'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-number>
          </xsl:if>
          <xsl:if test ="a:lvl5pPr/a:buBlip">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:value-of select="'●'"/>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl5pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl5pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl5pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl5pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl5pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl5pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl5pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'5'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
          <xsl:if test="not(a:lvl5pPr/a:buAutoNum) and not(a:lvl5pPr/a:buChar) and not(a:lvl5pPr/a:buBlip)">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:call-template name="InsertBulletChar">
                  <xsl:with-param name="character">
                    <xsl:value-of select ="a:lvl5pPr/a:buChar/@char"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl5pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl5pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl5pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl5pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl5pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl5pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl5pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'5'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
        </xsl:when>
        <xsl:when test="$levelNo='5'">
          <xsl:if test ="a:lvl6pPr/a:buChar">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:call-template name="InsertBulletChar">
                  <xsl:with-param name="character">
                    <xsl:value-of select ="a:lvl6pPr/a:buChar/@char"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl6pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl6pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl6pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl6pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl6pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl6pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl6pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'6'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
          <xsl:if test ="a:lvl6pPr/a:buAutoNum">
            <text:list-level-style-number>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:for-each select="./a:lvl6pPr">
                <xsl:call-template name="OutlineNumbering"></xsl:call-template>
              </xsl:for-each>
              <style:list-level-properties>
                <xsl:if test ="./a:lvl6pPr/@indent">
                  <xsl:attribute name ="text:min-label-width">
                    <xsl:value-of select="concat(format-number(./a:lvl6pPr/@indent div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if >
                <xsl:if test ="./a:lvl6pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl6pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl6pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'6'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-number>
          </xsl:if>
          <xsl:if test ="a:lvl6pPr/a:buBlip">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:value-of select="'●'"/>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl6pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl6pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl6pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl6pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl6pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl6pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl6pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'6'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
          <xsl:if test="not(a:lvl6pPr/a:buAutoNum) and not(a:lvl6pPr/a:buChar) and not(a:lvl6pPr/a:buBlip)">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:call-template name="InsertBulletChar">
                  <xsl:with-param name="character">
                    <xsl:value-of select ="a:lvl6pPr/a:buChar/@char"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl6pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl6pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl6pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl6pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl6pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl6pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl6pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'6'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
        </xsl:when>
        <xsl:when test="$levelNo='6'">
          <xsl:if test ="a:lvl7pPr/a:buChar">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:call-template name="InsertBulletChar">
                  <xsl:with-param name="character">
                    <xsl:value-of select ="a:lvl7pPr/a:buChar/@char"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl7pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl7pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl7pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl7pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl7pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl7pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl7pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'7'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
          <xsl:if test ="a:lvl7pPr/a:buAutoNum">
            <text:list-level-style-number>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:for-each select="./a:lvl7pPr">
                <xsl:call-template name="OutlineNumbering"></xsl:call-template>
              </xsl:for-each>
              <style:list-level-properties>
                <xsl:if test ="./a:lvl7pPr/@indent">
                  <xsl:attribute name ="text:min-label-width">
                    <xsl:value-of select="concat(format-number(./a:lvl7pPr/@indent div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if >
                <xsl:if test ="./a:lvl7pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl7pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl7pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'7'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-number>
          </xsl:if>
          <xsl:if test ="a:lvl7pPr/a:buBlip">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:value-of select="'●'"/>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl7pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl7pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl7pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl7pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl7pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl7pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl7pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'7'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
          <xsl:if test="not(a:lvl7pPr/a:buAutoNum) and not(a:lvl7pPr/a:buChar) and not(a:lvl7pPr/a:buBlip)">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:call-template name="InsertBulletChar">
                  <xsl:with-param name="character">
                    <xsl:value-of select ="a:lvl7pPr/a:buChar/@char"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl7pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl7pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl7pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl7pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl7pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl7pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl7pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'7'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
        </xsl:when>
        <xsl:when test="$levelNo='7'">
          <xsl:if test ="a:lvl8pPr/a:buChar">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:call-template name="InsertBulletChar">
                  <xsl:with-param name="character">
                    <xsl:value-of select ="a:lvl8pPr/a:buChar/@char"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl8pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl8pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl8pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl8pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl8pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl8pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl8pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'8'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
          <xsl:if test ="a:lvl8pPr/a:buAutoNum">
            <text:list-level-style-number>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:for-each select="./a:lvl8pPr">
                <xsl:call-template name="OutlineNumbering"></xsl:call-template>
              </xsl:for-each>
              <style:list-level-properties>
                <xsl:if test ="./a:lvl8pPr/@indent">
                  <xsl:attribute name ="text:min-label-width">
                    <xsl:value-of select="concat(format-number(./a:lvl8pPr/@indent div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if >
                <xsl:if test ="./a:lvl8pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl8pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl8pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'8'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-number>
          </xsl:if>
          <xsl:if test ="a:lvl8pPr/a:buBlip">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:value-of select="'●'"/>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl8pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl8pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl8pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl8pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl8pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl8pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl8pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'8'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
          <xsl:if test="not(a:lvl8pPr/a:buAutoNum) and not(a:lvl8pPr/a:buChar) and not(a:lvl8pPr/a:buBlip)">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:call-template name="InsertBulletChar">
                  <xsl:with-param name="character">
                    <xsl:value-of select ="a:lvl8pPr/a:buChar/@char"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl8pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl8pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl8pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl8pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl8pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl8pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl8pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'8'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
        </xsl:when>
        <xsl:when test="$levelNo='8'">
          <xsl:if test ="a:lvl9pPr/a:buChar">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:call-template name="InsertBulletChar">
                  <xsl:with-param name="character">
                    <xsl:value-of select ="a:lvl9pPr/a:buChar/@char"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl9pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl9pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl9pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl9pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl9pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl9pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl9pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'9'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
          <xsl:if test ="a:lvl9pPr/a:buAutoNum">
            <text:list-level-style-number>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:for-each select="./a:lvl9pPr">
                <xsl:call-template name="OutlineNumbering"></xsl:call-template>
              </xsl:for-each>
              <style:list-level-properties>
                <xsl:if test ="./a:lvl9pPr/@indent">
                  <xsl:attribute name ="text:min-label-width">
                    <xsl:value-of select="concat(format-number(./a:lvl9pPr/@indent div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if >
                <xsl:if test ="./a:lvl9pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl9pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl9pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'9'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-number>
          </xsl:if>
          <xsl:if test ="a:lvl9pPr/a:buBlip">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:value-of select="'●'"/>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl9pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl9pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl9pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl9pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl9pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl9pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl9pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'8'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
          <xsl:if test="not(a:lvl9pPr/a:buAutoNum) and not(a:lvl9pPr/a:buChar) and not(a:lvl9pPr/a:buBlip)">
            <text:list-level-style-bullet>
              <xsl:attribute name ="text:level">
                <xsl:value-of select ="$pos"/>
              </xsl:attribute>
              <xsl:attribute name ="text:bullet-char">
                <xsl:call-template name="InsertBulletChar">
                  <xsl:with-param name="character">
                    <xsl:value-of select ="a:lvl9pPr/a:buChar/@char"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute >
              <style:list-level-properties>
                <xsl:if test ="./a:lvl9pPr/@indent">
                  <xsl:choose>
                    <xsl:when test="./a:lvl9pPr/@indent &lt; 0">
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number((0 - ./a:lvl9pPr/@indent) div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name ="text:min-label-width">
                        <xsl:value-of select="concat(format-number(./a:lvl9pPr/@indent div 360000, '#.##'), 'cm')"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if >
                <xsl:if test ="./a:lvl9pPr/@marL">
                  <xsl:attribute name ="text:space-before">
                    <xsl:value-of select="concat(format-number(./a:lvl9pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:list-level-properties>
              <xsl:for-each select="./a:lvl9pPr">
                <xsl:call-template name="Outlines">
                  <xsl:with-param name="level">
                    <xsl:value-of select="'9'"/>
                  </xsl:with-param>
                  <xsl:with-param name="blnBullet">
                    <xsl:value-of select="'true'"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </text:list-level-style-bullet>
          </xsl:if>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="slideMasterOutlines">
    <xsl:param name="SlideMasterFile" />
    <xsl:param name="level"></xsl:param>

    <xsl:for-each select="document(concat('ppt/slideMasters/',$SlideMasterFile))/p:sldMaster/p:txStyles/p:bodyStyle">
      <xsl:choose>
        <xsl:when test="$level='1'">
          <style:paragraph-properties text:enable-numbering="true">

            <xsl:for-each select="./a:lvl1pPr">
              <xsl:call-template name="tmpShapeParaProp">
                <xsl:with-param name="lnSpcReduction" select="'0'"/>
              </xsl:call-template>
            </xsl:for-each>
          </style:paragraph-properties>
          <style:text-properties>
            <xsl:for-each select="a:lvl1pPr">
              <xsl:call-template name="tmpShapeTextProperty">
                <xsl:with-param name="fontType" select="'minor'"/>
              </xsl:call-template>
            </xsl:for-each>
          </style:text-properties>
        </xsl:when>
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="$level='2'">
          <style:paragraph-properties >
            <xsl:for-each select="a:lvl2pPr">
              <xsl:call-template name="tmpShapeParaProp">
                <xsl:with-param name="lnSpcReduction" select="'0'"/>
              </xsl:call-template>
            </xsl:for-each>
          </style:paragraph-properties>
          <style:text-properties>
            <xsl:for-each select="a:lvl2pPr">
              <xsl:call-template name="tmpShapeTextProperty">
                <xsl:with-param name="fontType" select="'minor'"/>
              </xsl:call-template>
            </xsl:for-each>

          </style:text-properties>
        </xsl:when>
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="$level='3'">
          <style:paragraph-properties >
            <xsl:for-each select="a:lvl3pPr">
              <xsl:call-template name="tmpShapeParaProp">
                <xsl:with-param name="lnSpcReduction" select="'0'"/>
              </xsl:call-template>
            </xsl:for-each>
          </style:paragraph-properties>
          <style:text-properties>
            <xsl:for-each select="a:lvl3pPr">
              <xsl:call-template name="tmpShapeTextProperty">
                <xsl:with-param name="fontType" select="'minor'"/>
              </xsl:call-template>
            </xsl:for-each>

          </style:text-properties>
        </xsl:when>
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="$level='4'">
          <style:paragraph-properties >
            <xsl:for-each select="a:lvl4pPr">
              <xsl:call-template name="tmpShapeParaProp">
                <xsl:with-param name="lnSpcReduction" select="'0'"/>
              </xsl:call-template>
            </xsl:for-each>
          </style:paragraph-properties>
          <style:text-properties>
            <xsl:for-each select="a:lvl4pPr">
              <xsl:call-template name="tmpShapeTextProperty">
                <xsl:with-param name="fontType" select="'minor'"/>
              </xsl:call-template>
            </xsl:for-each>
          </style:text-properties>
        </xsl:when>
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="$level='5'">
          <style:paragraph-properties >
            <xsl:for-each select="a:lvl5pPr">
              <xsl:call-template name="tmpShapeParaProp">
                <xsl:with-param name="lnSpcReduction" select="'0'"/>
              </xsl:call-template>
            </xsl:for-each>
          </style:paragraph-properties>
          <style:text-properties>
            <xsl:for-each select="a:lvl5pPr">
              <xsl:call-template name="tmpShapeTextProperty">
                <xsl:with-param name="fontType" select="'minor'"/>
              </xsl:call-template>
            </xsl:for-each>
          </style:text-properties>
        </xsl:when>
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="$level='6'">
          <style:paragraph-properties >
            <xsl:for-each select="a:lvl6pPr">
              <xsl:call-template name="tmpShapeParaProp">
                <xsl:with-param name="lnSpcReduction" select="'0'"/>
              </xsl:call-template>
            </xsl:for-each>
          </style:paragraph-properties>
          <style:text-properties>
            <xsl:for-each select="a:lvl6pPr">
              <xsl:call-template name="tmpShapeTextProperty">
                <xsl:with-param name="fontType" select="'minor'"/>
              </xsl:call-template>
            </xsl:for-each>
          </style:text-properties>
        </xsl:when>
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="$level='7'">
          <style:paragraph-properties >
            <xsl:for-each select="a:lvl7pPr">
              <xsl:call-template name="tmpShapeParaProp">
                <xsl:with-param name="lnSpcReduction" select="'0'"/>
              </xsl:call-template>
            </xsl:for-each>
          </style:paragraph-properties>
          <style:text-properties>
            <xsl:for-each select="a:lvl7pPr">
              <xsl:call-template name="tmpShapeTextProperty">
                <xsl:with-param name="fontType" select="'minor'"/>
              </xsl:call-template>
            </xsl:for-each>
          </style:text-properties>
        </xsl:when>
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="$level='8'">
          <style:paragraph-properties >
            <xsl:for-each select="a:lvl8pPr">
              <xsl:call-template name="tmpShapeParaProp">
                <xsl:with-param name="lnSpcReduction" select="'0'"/>
              </xsl:call-template>
            </xsl:for-each>
          </style:paragraph-properties>
          <style:text-properties>
            <xsl:for-each select="a:lvl8pPr">
              <xsl:call-template name="tmpShapeTextProperty">
                <xsl:with-param name="fontType" select="'minor'"/>
              </xsl:call-template>
            </xsl:for-each>
          </style:text-properties>
        </xsl:when>
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="$level='9'">
          <style:paragraph-properties >
            <xsl:for-each select="a:lvl9pPr">
              <xsl:call-template name="tmpShapeParaProp">
                <xsl:with-param name="lnSpcReduction" select="'0'"/>
              </xsl:call-template>
            </xsl:for-each>
          </style:paragraph-properties>
          <style:text-properties>
            <xsl:for-each select="a:lvl9pPr">
              <xsl:call-template name="tmpShapeTextProperty">
                <xsl:with-param name="fontType" select="'minor'"/>
              </xsl:call-template>
            </xsl:for-each>
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
  <xsl:template name="TitleStyle">
    <xsl:param name="SlideMasterFile"></xsl:param>
    <style:graphic-properties draw:stroke="none">
      <xsl:call-template name="TitleSubTitleGraphicProperty">
        <xsl:with-param name="blnSubTitle">
          <xsl:value-of select="'false'"/>
        </xsl:with-param>
      </xsl:call-template>
      <text:list-style>
        <text:list-level-style-bullet text:level="1">
        </text:list-level-style-bullet>
      </text:list-style>
    </style:graphic-properties>

    <style:paragraph-properties>
      <xsl:for-each select="document(concat('ppt/slideMasters/',$SlideMasterFile))//p:txStyles/p:titleStyle/a:lvl1pPr">
        <xsl:call-template name="tmpShapeParaProp"/>
      </xsl:for-each>
    </style:paragraph-properties>
    <xsl:for-each select="document(concat('ppt/slideMasters/',$SlideMasterFile))//p:txStyles/p:titleStyle/a:lvl1pPr">
      <style:text-properties>
        <xsl:call-template name="tmpShapeTextProperty">
          <xsl:with-param name="fontType" select="'major'"/>
        </xsl:call-template>
      </style:text-properties>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="DatePageNoFooterStyle">
    <style:graphic-properties draw:stroke="none">
      <xsl:call-template name="TitleSubTitleGraphicProperty"/>
    </style:graphic-properties>
    <style:paragraph-properties>
      <xsl:for-each select="p:txBody//a:lvl1pPr">
        <xsl:call-template name="tmpShapeParaProp"/>
      </xsl:for-each>
    </style:paragraph-properties>
    <xsl:for-each select="p:txBody//a:lvl1pPr">
      <style:text-properties>
        <xsl:call-template name="tmpShapeTextProperty">
          <xsl:with-param name="fontType" select="'minor'"/>
        </xsl:call-template>
      </style:text-properties>
    </xsl:for-each>

  </xsl:template>
  <xsl:template name="InsertBulletChar">
    <xsl:param name="character" />
    <xsl:choose>

      <xsl:when test="$character = 'à'">➔</xsl:when>
      <xsl:when test="$character = '§'">■</xsl:when>
      <xsl:when test="$character = 'q'">•</xsl:when>
      <xsl:when test="$character = '•'">•</xsl:when>
      <xsl:when test="$character = 'v'">●</xsl:when>
      <xsl:when test="$character = 'Ø'">➢</xsl:when>
      <xsl:when test="$character = 'ü'">✔</xsl:when>
      <xsl:when test="$character = 'o'">○</xsl:when>
      <xsl:when test="$character = '-'">–</xsl:when>
      <xsl:when test="$character = '»'">»</xsl:when>
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

        <style:style style:family="drawing-page">
          <xsl:attribute name="style:name">
            <xsl:value-of select="concat('dp',$currentPos)"/>
          </xsl:attribute>
          <style:drawing-page-properties presentation:background-visible="true" presentation:background-objects-visible="true" presentation:display-footer="true" presentation:display-page-number="true" presentation:display-date-time="true">
            <xsl:call-template name="getSlideMasterBGColor">
              <xsl:with-param name="slideMasterFile">
                <xsl:value-of select="$slideMasterPath"/>
              </xsl:with-param>
            </xsl:call-template>
          </style:drawing-page-properties>
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
                      <!--<xsl:for-each select ="./p:txBody/a:p">
                            <xsl:call-template name="outlineTextBox">
                                <xsl:with-param name="level">
                                <xsl:value-of select="./a:pPr/@lvl"/>
                              </xsl:with-param>
                              <xsl:with-param name="text">
                                <xsl:value-of select="./a:r/a:t/."/>
                              </xsl:with-param>
                            </xsl:call-template>
                          </xsl:for-each>-->
                    </draw:text-box>
                  </draw:frame>
                </xsl:when>
                <xsl:when test="p:nvSpPr/p:nvPr/p:ph/@type='dt'">
                  <draw:frame>
                    <xsl:attribute name ="presentation:style-name">
                      <xsl:value-of select ="concat($slideMasterName,'-DateTime')"/>
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
                    <xsl:attribute name ="presentation:style-name">
                      <xsl:value-of select ="concat($slideMasterName,'-footer')"/>
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
                    <xsl:attribute name ="presentation:style-name">
                      <xsl:value-of select ="concat($slideMasterName,'-pageno')"/>
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
                          <text:page-number></text:page-number>
                        </text:span>
                      </text:p>
                    </draw:text-box>

                  </draw:frame>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:variable  name ="GraphicId">
                    <xsl:value-of select ="concat($slideMasterName,concat('gr',position()))"/>
                  </xsl:variable>
                  <xsl:variable name ="ParaId">
                    <xsl:value-of select ="concat($slideMasterName,concat('PARA',position()))"/>
                  </xsl:variable>
                  <xsl:call-template name ="shapes">
                    <xsl:with-param name="GraphicId" select ="$GraphicId"/>
                    <xsl:with-param name ="ParaId" select="$ParaId" />
                  </xsl:call-template>
                </xsl:otherwise>
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
  <xsl:template name="SubtitleStyle">
    <xsl:param name="SlideMasterFile"></xsl:param>
    <style:graphic-properties draw:stroke="none">
      <xsl:call-template name="TitleSubTitleGraphicProperty">
        <xsl:with-param name="blnSubTitle">
          <xsl:value-of select="'true'"/>
        </xsl:with-param>
      </xsl:call-template>
      <text:list-style>
        <text:list-level-style-bullet text:level="1" text:bullet-char="?">
          <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
        </text:list-level-style-bullet>
        <text:list-level-style-bullet text:level="2" text:bullet-char="?">
          <style:list-level-properties text:space-before="0.6cm" text:min-label-width="0.6cm"/>
          <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
        </text:list-level-style-bullet>
        <text:list-level-style-bullet text:level="3" text:bullet-char="?">
          <style:list-level-properties text:space-before="1.2cm" text:min-label-width="0.6cm"/>
          <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
        </text:list-level-style-bullet>
        <text:list-level-style-bullet text:level="4" text:bullet-char="?">
          <style:list-level-properties text:space-before="1.8cm" text:min-label-width="0.6cm"/>
          <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
        </text:list-level-style-bullet>
        <text:list-level-style-bullet text:level="5" text:bullet-char="?">
          <style:list-level-properties text:space-before="2.4cm" text:min-label-width="0.6cm"/>
          <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
        </text:list-level-style-bullet>
        <text:list-level-style-bullet text:level="6" text:bullet-char="?">
          <style:list-level-properties text:space-before="3cm" text:min-label-width="0.6cm"/>
          <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
        </text:list-level-style-bullet>
        <text:list-level-style-bullet text:level="7" text:bullet-char="?">
          <style:list-level-properties text:space-before="3.6cm" text:min-label-width="0.6cm"/>
          <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
        </text:list-level-style-bullet>
        <text:list-level-style-bullet text:level="8" text:bullet-char="?">
          <style:list-level-properties text:space-before="4.2cm" text:min-label-width="0.6cm"/>
          <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
        </text:list-level-style-bullet>
        <text:list-level-style-bullet text:level="9" text:bullet-char="?">
          <style:list-level-properties text:space-before="4.8cm" text:min-label-width="0.6cm"/>
          <style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/>
        </text:list-level-style-bullet>
      </text:list-style>
    </style:graphic-properties>
    <style:paragraph-properties>
      <xsl:for-each select="document(concat('ppt/slideMasters/',$SlideMasterFile))//p:txStyles/p:bodyStyle/a:lvl1pPr">
        <xsl:call-template name="tmpShapeParaProp"/>
      </xsl:for-each>
    </style:paragraph-properties>
    <xsl:for-each select="document(concat('ppt/slideMasters/',$SlideMasterFile))//p:txStyles/p:bodyStyle/a:lvl1pPr">
      <style:text-properties>
        <xsl:call-template name="tmpShapeTextProperty">
          <xsl:with-param name="fontType" select="'minor'"/>
        </xsl:call-template>
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
    <xsl:param name="level"></xsl:param>
    <xsl:param name="paraId"></xsl:param>
    <xsl:param name="textId"></xsl:param>
    <xsl:for-each select ="p:txBody/a:p">
      <xsl:choose>
        <xsl:when test ="$level= 0">
          <text:list>
            <text:list-item>
              <text:p>
                <text:span>
                  <xsl:value-of select ="."/>
                </text:span>
              </text:p>
            </text:list-item>
          </text:list>
        </xsl:when>
        <xsl:when test ="$level = 1">
          <text:list>
            <text:list-item>
              <text:list>
                <text:list-item>
                  <text:p >
                    <text:span>
                      <xsl:value-of select ="."/>
                    </text:span>
                  </text:p>
                </text:list-item>
              </text:list>
            </text:list-item>
          </text:list>
        </xsl:when>
        <xsl:when test ="$level = 2">
          <text:list>
            <text:list-item>
              <text:list>
                <text:list-item>
                  <text:list>
                    <text:list-item>
                      <text:p >
                        <text:span>
                          <xsl:value-of select ="."/>
                        </text:span>
                      </text:p>
                    </text:list-item>
                  </text:list>
                </text:list-item>
              </text:list>
            </text:list-item>
          </text:list>
        </xsl:when>
        <xsl:when test ="$level = 3">
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
        <xsl:otherwise>
          <xsl:attribute name="draw:fill-color">
            <xsl:value-of select="'#FFFFFF'"/>
          </xsl:attribute>
          <xsl:attribute name="draw:fill">
            <xsl:value-of select="'solid'"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="Outlines">
    <xsl:param name="level"></xsl:param>
    <xsl:param name="blnBullet"></xsl:param>
    <style:paragraph-properties>
      <xsl:choose>
        <xsl:when test="./a:buChar">
          <xsl:attribute name ="text:enable-numbering">
            <xsl:value-of select ="'true'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name ="text:enable-numbering">
            <xsl:value-of select ="'false'"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
      <!--<xsl:choose>
        <xsl:when test="$level='1'">
          <xsl:attribute name ="text:enable-numbering">
            <xsl:value-of select ="'true'"/>
          </xsl:attribute>
        </xsl:when>
      </xsl:choose>-->
      <xsl:call-template name="tmpShapeParaProp"/>
    </style:paragraph-properties>
    <style:text-properties>
      <xsl:choose>
        <xsl:when test="$blnBullet='true'">
          <xsl:attribute name ="style:font-charset">
            <xsl:value-of select ="'x-symbol'"/>
          </xsl:attribute>
          <xsl:if test ="./a:buSzPct">
            <xsl:attribute name ="fo:font-size">
              <xsl:value-of select ="concat((./a:buSzPct/@val div 1000),'%')"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test ="not(./a:buSzPct)">
            <xsl:attribute name ="fo:font-size">
              <xsl:value-of select ="'100%'"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test ="./a:buClr">
            <xsl:if test ="./a:buClr/a:srgbClr">
              <xsl:variable name ="color" select ="./a:buClr/a:srgbClr/@val"/>
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="concat('#',$color)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="./a:buClr/a:schemeClr">
              <xsl:attribute name ="fo:color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="./a:buClr/a:schemeClr/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
          <xsl:if test ="not(./a:buClr)">
            <xsl:choose>
              <xsl:when test="./a:defRPr/a:solidFill/a:srgbClr">
                <xsl:attribute name ="fo:color">
                  <xsl:value-of select="concat('#',./a:defRPr/a:solidFill/a:srgbClr/@val)"></xsl:value-of>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="./a:defRPr/a:solidFill/a:schemeClr">
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
                      </xsl:choose>
                    </xsl:with-param>
                    <xsl:with-param name ="lumOff">
                      <xsl:choose>
                        <xsl:when test="./a:defRPr/a:solidFill/a:schemeClr/a:lumOff">
                          <xsl:value-of select="./a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val" />
                        </xsl:when>
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
          </xsl:if>
          <xsl:attribute name ="fo:font-family">
            <xsl:if test ="./a:buFont/@typeface">
              <xsl:variable name ="typeFaceVal" select ="./a:buFont/@typeface"/>
              <xsl:for-each select ="./a:buFont/@typeface">
                <xsl:if test ="$typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt'">
                  <xsl:call-template name="slideMasterFontName">
                    <xsl:with-param name="fontType">
                      <xsl:value-of select="'minor'"/>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:if>
                <xsl:if test ="not($typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt')">
                  <xsl:value-of select ="."/>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
            <xsl:if test ="not(./a:buFont/@typeface)">
              <xsl:value-of select ="'StarSymbol'"/>
            </xsl:if>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name ="style:font-size-asian">
            <xsl:value-of select ="concat(./a:defRPr/@sz div 100 ,'pt')"/>
          </xsl:attribute>
          <xsl:attribute name ="style:font-size-complex">
            <xsl:value-of select ="concat(./a:defRPr/@sz div 100 ,'pt')"/>
          </xsl:attribute>
          <xsl:call-template name="tmpShapeTextProperty">
            <xsl:with-param name="fontType" select="'minor'"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </style:text-properties>
  </xsl:template>
  <xsl:template name="getSPBackColor">
    <!--for getting back ground color-->
    <xsl:choose>
      <xsl:when test="./a:solidFill/a:srgbClr">
        <xsl:attribute name ="draw:fill-color">
          <xsl:value-of select ="concat('#',./a:solidFill/a:srgbClr/@val)"/>
        </xsl:attribute>
        <xsl:attribute name="draw:fill">
          <xsl:value-of select="'solid'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="./a:solidFill/a:schemeClr">
        <xsl:attribute name ="draw:fill-color">
          <xsl:call-template name ="getColorCode">
            <xsl:with-param name ="color">
              <xsl:value-of select="./a:solidFill/a:schemeClr/@val"/>
            </xsl:with-param>
            <xsl:with-param name ="lumMod">
              <xsl:if test="./a:solidFill/a:schemeClr/a:lumMod">
                <xsl:value-of select="./a:solidFill/a:schemeClr/a:lumMod/@val" />
              </xsl:if>
            </xsl:with-param>
            <xsl:with-param name ="lumOff">
              <xsl:if test="./a:solidFill/a:schemeClr/a:lumOff">
                <xsl:value-of select="./a:solidFill/a:schemeClr/a:lumOff/@val" />
              </xsl:if>
            </xsl:with-param>
          </xsl:call-template>
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
  </xsl:template>

  <!--added by vipul to insert slide master's shapes style-->
  <xsl:template name ="GraphicStyleForSlideMaster">
    <xsl:for-each select="document('ppt/presentation.xml')//p:sldMasterIdLst/p:sldMasterId">
      <xsl:variable name="sldMasterIdRelation">
        <xsl:value-of select="@r:id"></xsl:value-of>
      </xsl:variable>
      <!--Loop thru each slide master.xml-->
      <xsl:for-each select="document('ppt/_rels/presentation.xml.rels')//node()[@Id=$sldMasterIdRelation]">
        <xsl:variable name="slideMasterPath">
          <xsl:value-of select="substring-after(@Target,'/')"/>
        </xsl:variable>
        <xsl:variable name="slideMasterName">
          <xsl:value-of select="substring-before($slideMasterPath,'.xml')"/>
        </xsl:variable>

        <!--Graphic properties for shapes with p:sp nodes-->
        <xsl:for-each select="document(concat('ppt/slideMasters/',$slideMasterPath))//p:sp">
          <xsl:if test="p:nvSpPr/p:nvPr/p:ph/@type='dt'">
            <!-- style for DateTime-->
            <style:style>
              <xsl:attribute name="style:name">
                <xsl:value-of select="concat($slideMasterName,'-DateTime')"/>
              </xsl:attribute>
              <xsl:attribute name ="style:family">
                <xsl:value-of select ="'presentation'"/>
              </xsl:attribute>
              <xsl:attribute name="style:parent-style-name">
                <xsl:value-of select="concat($slideMasterName,'-backgroundobjects')"/>
              </xsl:attribute>
              <xsl:call-template name="DatePageNoFooterStyle"/>
            </style:style>
          </xsl:if>
          <xsl:if test="p:nvSpPr/p:nvPr/p:ph/@type='ftr'">
            <!-- style for footer-->
            <style:style>
              <xsl:attribute name="style:name">
                <xsl:value-of select="concat($slideMasterName,'-footer')"/>
              </xsl:attribute>
              <xsl:attribute name ="style:family">
                <xsl:value-of select ="'presentation'"/>
              </xsl:attribute>
              <xsl:attribute name="style:parent-style-name">
                <xsl:value-of select="concat($slideMasterName,'-backgroundobjects')"/>
              </xsl:attribute>
              <xsl:call-template name="DatePageNoFooterStyle"/>
            </style:style>
          </xsl:if>
          <xsl:if test="p:nvSpPr/p:nvPr/p:ph/@type='sldNum'">
            <!-- style for DateTime-->
            <style:style>
              <xsl:attribute name="style:name">
                <xsl:value-of select="concat($slideMasterName,'-pageno')"/>
              </xsl:attribute>
              <xsl:attribute name ="style:family">
                <xsl:value-of select ="'presentation'"/>
              </xsl:attribute>
              <xsl:attribute name="style:parent-style-name">
                <xsl:value-of select="concat($slideMasterName,'-backgroundobjects')"/>
              </xsl:attribute>
              <xsl:call-template name="DatePageNoFooterStyle"/>
            </style:style>
          </xsl:if>
          <!--Check for shape or texbox-->
          <xsl:if test = "not(p:nvSpPr/p:nvPr/p:ph/@type) and not(p:nvSpPr/p:nvPr/p:ph/@idx)">
            <!--Generate graphic properties ID-->
            <xsl:variable  name ="GraphicId">
              <xsl:value-of select ="concat($slideMasterName,concat('gr',position()))"/>
            </xsl:variable>
            <!--Generate Paragraph properties ID-->
            <xsl:variable  name ="ParaId">
              <xsl:value-of select ="concat($slideMasterName,concat('PARA',position()))"/>
            </xsl:variable>

            <style:style style:family="graphic" style:parent-style-name="standard">
              <xsl:attribute name ="style:name">
                <xsl:value-of select ="$GraphicId"/>
              </xsl:attribute >
              <style:graphic-properties draw:stroke="none">

                <!--FILL-->
                <xsl:call-template name ="Fill" />

                <!--LINE COLOR-->
                <xsl:call-template name ="LineColor" />

                <!--LINE STYLE-->
                <xsl:call-template name ="LineStyle"/>

                <!--TEXT ALIGNMENT-->
                <xsl:call-template name ="TextLayout" />

              </style:graphic-properties >
            </style:style>
            <style:style style:family="paragraph">
              <xsl:attribute name ="style:name">
                <xsl:value-of select ="$ParaId"/>
              </xsl:attribute >
              <style:paragraph-properties>
                <xsl:if test ="p:txBody/a:bodyPr/@vert">
                  <xsl:attribute name ="style:writing-mode">
                    <xsl:call-template name ="getTextDirection">
                      <xsl:with-param name ="vert" select ="p:txBody/a:bodyPr/@vert" />
                    </xsl:call-template>
                  </xsl:attribute>
                </xsl:if>
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
                <xsl:if test ="p:txBody/a:p/a:pPr/a:spcBef/a:spcPts/@val">
                  <xsl:attribute name ="fo:margin-top">
                    <xsl:value-of select="concat(format-number(p:txBody/a:p/a:pPr/a:spcBef/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test ="p:txBody/a:p/a:pPr/a:spcAft/a:spcPts/@val">
                  <xsl:attribute name ="fo:margin-bottom">
                    <xsl:value-of select="concat(format-number(p:txBody/a:p/a:pPr/a:spcAft/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
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
                <xsl:choose>
                  <!-- Center Alignment-->
                  <xsl:when test ="p:txBody/a:p/a:pPr/@algn ='ctr'">
                    <xsl:attribute name ="fo:text-align">
                      <xsl:value-of select ="'center'"/>
                    </xsl:attribute>
                  </xsl:when>
                  <!-- Right Alignment-->
                  <xsl:when test ="p:txBody/a:p/a:pPr/@algn ='r'">
                    <xsl:attribute name ="fo:text-align">
                      <xsl:value-of select ="'end'"/>
                    </xsl:attribute>
                  </xsl:when>
                  <!-- Left Alignment-->
                  <xsl:when test ="p:txBody/a:p/a:pPr/@algn ='l'">
                    <xsl:attribute name ="fo:text-align">
                      <xsl:value-of select ="'start'"/>
                    </xsl:attribute>
                  </xsl:when>
                  <!-- Justified Alignment-->
                  <xsl:when test ="p:txBody/a:p/a:pPr/@algn ='just'">
                    <xsl:attribute name ="fo:text-align">
                      <xsl:value-of select ="'justified'"/>
                    </xsl:attribute>
                  </xsl:when>

                  <!-- end-->
                </xsl:choose>
                <!-- If the line space is in Percentage-->
                <xsl:if test ="p:txBody/a:p/a:pPr/a:lnSpc/a:spcPct/@val">
                  <xsl:attribute name="fo:line-height">
                    <xsl:value-of select="concat(format-number(p:txBody/a:p/a:pPr/a:lnSpc/a:spcPct/@val div 1000,'###'), '%')"/>
                  </xsl:attribute>
                </xsl:if >
                <!-- If the line space is in Points-->
                <xsl:if test ="p:txBody/a:p/a:pPr/a:lnSpc/a:spcPts">
                  <xsl:attribute name="style:line-height-at-least">
                    <xsl:value-of select="concat(format-number(p:txBody/a:p/a:pPr/a:lnSpc/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </style:paragraph-properties>
              <style:text-properties>
                <xsl:attribute name ="fo:font-family">
                  <xsl:choose>
                    <xsl:when test="p:txBody/a:p/a:r/a:rPr/a:latin/@typeface">
                      <xsl:value-of select ="p:txBody/a:p/a:r/a:rPr/a:latin/@typeface"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:call-template name="slideMasterFontName">
                        <xsl:with-param name="fontType">
                          <xsl:value-of select="'minor'"/>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
                <xsl:choose>
                  <xsl:when test="p:style/a:fontRef/a:schemeClr">
                    <xsl:attribute name ="fo:color">
                      <xsl:call-template name ="getColorCode">
                        <xsl:with-param name ="color">
                          <xsl:value-of select="p:style/a:fontRef/a:schemeClr/@val"/>
                        </xsl:with-param>
                        <xsl:with-param name ="lumMod">
                          <xsl:choose>
                            <xsl:when test="p:style/a:fontRef/a:schemeClr/a:lumMod">
                              <xsl:value-of select="p:style/a:fontRef/a:schemeClr/a:lumMod/@val" />
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="''" />
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:with-param>
                        <xsl:with-param name ="lumOff">
                          <xsl:choose>
                            <xsl:when test="p:style/a:fontRef/a:schemeClr/a:lumOff">
                              <xsl:value-of select="p:style/a:fontRef/a:schemeClr/@val" />
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="''" />
                            </xsl:otherwise>
                          </xsl:choose>

                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:when test="p:txBody/a:p/a:r/a:rPr/a:solidFill/a:schemeClr">
                    <xsl:attribute name ="fo:color">
                      <xsl:call-template name ="getColorCode">
                        <xsl:with-param name ="color">
                          <xsl:value-of select="p:txBody/a:p/a:r/a:rPr/a:solidFill/a:schemeClr/@val"/>
                        </xsl:with-param>
                        <xsl:with-param name ="lumMod">
                          <xsl:choose>
                            <xsl:when test="p:txBody/a:p/a:r/a:rPr/a:solidFill/a:schemeClr/a:lumMod">
                              <xsl:value-of select="p:txBody/a:p/a:r/a:rPr/a:solidFill/a:schemeClr/a:lumMod/@val" />
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="''" />
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:with-param>
                        <xsl:with-param name ="lumOff">
                          <xsl:choose>
                            <xsl:when test="p:txBody/a:p/a:r/a:rPr/a:solidFill/a:schemeClr/a:lumOff">
                              <xsl:value-of select="p:txBody/a:p/a:r/a:rPr/a:solidFill/a:schemeClr/@val" />
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="''" />
                            </xsl:otherwise>
                          </xsl:choose>

                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:when test="p:txBody/a:p/a:r/a:rPr/a:solidFill/a:srgbClr">
                    <xsl:attribute name ="fo:color">
                      <xsl:value-of select="concat('#',p:txBody/a:p/a:r/a:rPr/a:solidFill/a:srgbClr/@val)"/>
                    </xsl:attribute>
                  </xsl:when>
                </xsl:choose>
                <xsl:attribute name ="fo:font-size">
                  <xsl:if test="p:txBody/a:p/a:r/a:rPr/@sz">
                    <xsl:value-of select ="concat(format-number(p:txBody/a:p/a:r/a:rPr/@sz div 100,'#.##'), 'pt')"/>
                  </xsl:if>
                </xsl:attribute>
              </style:text-properties>
            </style:style>
          </xsl:if >
        </xsl:for-each>
        <!--Graphic properties for shapes with p:cxnSp nodes-->
        <xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMasterPath))/p:sld/p:cSld/p:spTree/p:cxnSp">
          <!--Generate graphic properties ID-->
          <xsl:variable  name ="GraphicId">
            <xsl:value-of select ="concat($slideMasterName,concat('gr',position()))"/>
          </xsl:variable>
          <!--Generate Paragraph properties ID-->
          <xsl:variable  name ="ParaId">
            <xsl:value-of select ="concat($slideMasterName,concat('PARA',position()))"/>
          </xsl:variable>
          <style:style style:family="graphic" style:parent-style-name="standard">
            <xsl:attribute name ="style:name">
              <xsl:value-of select ="$GraphicId"/>
            </xsl:attribute >
            <style:graphic-properties draw:stroke="none">

              <!--FILL-->
              <xsl:call-template name ="Fill" />

              <!--LINE COLOR-->
              <xsl:call-template name ="LineColor" />

              <!--LINE STYLE-->
              <xsl:call-template name ="LineStyle"/>

              <!--TEXT ALIGNMENT-->
              <xsl:call-template name ="TextLayout" />
            </style:graphic-properties >
          </style:style>
          <style:style style:family="paragraph">
            <xsl:attribute name ="style:name">
              <xsl:value-of select ="$ParaId"/>
            </xsl:attribute >
            <style:paragraph-properties>
              <xsl:if test ="p:txBody/a:bodyPr/@vert">
                <xsl:attribute name ="style:writing-mode">
                  <xsl:call-template name ="getTextDirection">
                    <xsl:with-param name ="vert" select ="p:txBody/a:bodyPr/@vert" />
                  </xsl:call-template>
                </xsl:attribute>
              </xsl:if>
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
            <style:text-properties>
              <xsl:attribute name ="fo:font-family">
                <xsl:choose>
                  <xsl:when test="p:txBody/a:p/a:r/a:rPr/a:latin/@typeface">
                    <xsl:value-of select ="p:txBody/a:p/a:r/a:rPr/a:latin/@typeface"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:call-template name="slideMasterFontName">
                      <xsl:with-param name="fontType">
                        <xsl:value-of select="'minor'"/>
                      </xsl:with-param>
                    </xsl:call-template>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
              <xsl:choose>
                <xsl:when test="p:style/a:fontRef/a:schemeClr">
                  <xsl:attribute name ="fo:color">
                    <xsl:call-template name ="getColorCode">
                      <xsl:with-param name ="color">
                        <xsl:value-of select="p:style/a:fontRef/a:schemeClr/@val"/>
                      </xsl:with-param>
                      <xsl:with-param name ="lumMod">
                        <xsl:choose>
                          <xsl:when test="p:style/a:fontRef/a:schemeClr/a:lumMod">
                            <xsl:value-of select="p:style/a:fontRef/a:schemeClr/a:lumMod/@val" />
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="''" />
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:with-param>
                      <xsl:with-param name ="lumOff">
                        <xsl:choose>
                          <xsl:when test="p:style/a:fontRef/a:schemeClr/a:lumOff">
                            <xsl:value-of select="p:style/a:fontRef/a:schemeClr/@val" />
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
            </style:text-properties>
          </style:style>
        </xsl:for-each>
      </xsl:for-each>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="OutlineNumbering">
    <xsl:for-each select=".">
      <xsl:if test ="a:buAutoNum">
        <xsl:if test ="a:buAutoNum/@type[contains(.,'arabicPeriod')]">
          <xsl:attribute name ="style:num-format" >
            <xsl:value-of  select ="'1'"/>
          </xsl:attribute>
          <xsl:attribute name ="style:num-suffix">
            <xsl:value-of select ="'.'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="a:buAutoNum/@type[contains(.,'arabicParenR')]">
          <xsl:attribute name ="style:num-format" >
            <xsl:value-of  select ="'1'"/>
          </xsl:attribute >
          <xsl:attribute name ="style:num-suffix">
            <xsl:value-of select ="')'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="a:buAutoNum/@type[contains(.,'arabicParenBoth')]">
          <xsl:attribute name ="style:num-format" >
            <xsl:value-of  select ="'1'"/>
          </xsl:attribute >
          <xsl:attribute name ="style:num-suffix">
            <xsl:value-of select ="')'"/>
          </xsl:attribute>
          <xsl:attribute name ="style:num-prefix">
            <xsl:value-of select ="'('"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="a:buAutoNum/@type[contains(.,'alphaUcPeriod')]">
          <xsl:attribute name ="style:num-format" >
            <xsl:value-of  select ="'A'"/>
          </xsl:attribute >
          <xsl:attribute name ="style:num-suffix">
            <xsl:value-of select ="'.'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="a:buAutoNum/@type[contains(.,'alphaUcParenR')]">
          <xsl:attribute name ="style:num-format" >
            <xsl:value-of  select ="'A'"/>
          </xsl:attribute >
          <xsl:attribute name ="style:num-suffix">
            <xsl:value-of select ="')'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="a:buAutoNum/@type[contains(.,'alphaUcParenBoth')]">
          <xsl:attribute name ="style:num-format" >
            <xsl:value-of  select ="'A'"/>
          </xsl:attribute >
          <xsl:attribute name ="style:num-suffix">
            <xsl:value-of select ="')'"/>
          </xsl:attribute>
          <xsl:attribute name ="style:num-prefix">
            <xsl:value-of select ="'('"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="a:buAutoNum/@type[contains(.,'alphaLcPeriod')]">
          <xsl:attribute name ="style:num-format" >
            <xsl:value-of  select ="'a'"/>
          </xsl:attribute >
          <xsl:attribute name ="style:num-suffix">
            <xsl:value-of select ="'.'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="a:buAutoNum/@type[contains(.,'alphaLcParenR')]">
          <xsl:attribute name ="style:num-format" >
            <xsl:value-of  select ="'a'"/>
          </xsl:attribute >
          <xsl:attribute name ="style:num-suffix">
            <xsl:value-of select ="')'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="a:buAutoNum/@type[contains(.,'alphaLcParenBoth')]">
          <xsl:attribute name ="style:num-format" >
            <xsl:value-of  select ="'a'"/>
          </xsl:attribute >
          <xsl:attribute name ="style:num-suffix">
            <xsl:value-of select ="')'"/>
          </xsl:attribute>
          <xsl:attribute name ="style:num-prefix">
            <xsl:value-of select ="'('"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="a:buAutoNum/@type[contains(.,'romanUcPeriod')]">
          <xsl:attribute name ="style:num-format" >
            <xsl:value-of  select ="'I'"/>
          </xsl:attribute >
          <xsl:attribute name ="style:num-suffix">
            <xsl:value-of select ="'.'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="a:buAutoNum/@type[contains(.,'romanLcPeriod')]">
          <xsl:attribute name ="style:num-format" >
            <xsl:value-of  select ="'i'"/>
          </xsl:attribute >
          <xsl:attribute name ="style:num-suffix">
            <xsl:value-of select ="'.'"/>
          </xsl:attribute>
          <!-- start at value-->
          <!-- <xsl:if test ="@startAt">
              <xsl:attribute name ="text:start-value">
                <xsl:value-of select ="a:buAutoNum[@startAt]"/>
              </xsl:attribute>
            </xsl:if>-->
        </xsl:if >
        <xsl:if test="a:buAutoNum/@startAt">
          <xsl:attribute name ="text:start-value">
            <xsl:value-of select ="a:buAutoNum/@startAt"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="not(a:buAutoNum/@startAt)">
          <xsl:attribute name ="text:start-value">
            <xsl:value-of select ="1"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="TitleSubTitleGraphicProperty">
    <xsl:param name="blnSubTitle"></xsl:param>
    <!--<xsl:if test="$blnSubTitle='false'">-->
    <!--for getting back ground color-->
    <xsl:for-each select="./p:spPr">
      <xsl:call-template name="getSPBackColor"></xsl:call-template>
    </xsl:for-each>
    <!--</xsl:if>-->
    <!--for getting border color-->
    <xsl:if test="p:spPr/a:ln/a:solidFill/a:srgbClr">
      <xsl:attribute name ="draw:stroke">
        <xsl:value-of select ="concat('#',p:spPr/a:ln/a:solidFill/a:srgbClr/@val)"/>
      </xsl:attribute>
    </xsl:if>
    <!--Vertical alignment-->
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
    <xsl:if test ="p:txBody/a:bodyPr/@anchorCtr" >
      <!--Horizontal alignment-->
      <xsl:attribute name ="draw:textarea-horizontal-align">
        <xsl:if test ="p:txBody/a:bodyPr/@anchorCtr= 1">
          <xsl:value-of select ="'center'"/>
        </xsl:if>
        <xsl:if test ="p:txBody/a:bodyPr/@anchorCtr= 0">
          <xsl:value-of select ="'justify'"/>
        </xsl:if>
      </xsl:attribute>
    </xsl:if>
    <!--wrap-->
    <xsl:choose>
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
    <!--resize text to shape-->
    <xsl:choose>
      <xsl:when test ="p:txBody/a:bodyPr/a:spAutoFit">
        <xsl:attribute name ="draw:auto-grow-height" >
          <xsl:value-of select ="'true'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name ="draw:auto-grow-height" >
          <xsl:value-of select ="'false'"/>
        </xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>
    <!--Line Spacing-->
    <xsl:attribute name ="fo:min-height" >
      <xsl:value-of select ="'12.573cm'"/>
    </xsl:attribute>
    <!-- internal margin-->
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
    <xsl:if test ="p:txBody/a:bodyPr/@tIns">
      <xsl:attribute name ="fo:padding-top">
        <xsl:value-of select ="concat(format-number(p:txBody/a:bodyPr/@tIns div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="p:txBody/a:bodyPr/@bIns">
      <xsl:attribute name ="fo:padding-bottom">
        <xsl:value-of select ="concat(format-number(p:txBody/a:bodyPr/@bIns div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="p:txBody/a:bodyPr/@lIns">
      <xsl:attribute name ="fo:padding-left">
        <xsl:value-of select ="concat(format-number(p:txBody/a:bodyPr/@lIns div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="p:txBody/a:bodyPr/@rIns">
      <xsl:attribute name ="fo:padding-right">
        <xsl:value-of select ="concat(format-number(p:txBody/a:bodyPr/@rIns div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>


  </xsl:template>
  <xsl:template name="spDateTimeGraphicProperty">
    <xsl:for-each select=".">

      <!--for getting Title box back ground color-->
      <xsl:choose>

        <xsl:when test="p:txBody//a:solidFill/a:srgbClr">
          <xsl:attribute name ="draw:fill-color">
            <xsl:value-of select ="concat('#',p:txBody//a:solidFill/a:srgbClr/@val)"/>
          </xsl:attribute>
          <xsl:attribute name="draw:fill">
            <xsl:value-of select="'solid'"/>
          </xsl:attribute>
          <!--<xsl:attribute name="draw:fill">
                  <xsl:value-of select="'none'"/>
                </xsl:attribute>-->
        </xsl:when>
        <xsl:when test="p:txBody//a:solidFill/a:schemeClr">
          <xsl:attribute name ="draw:fill-color">
            <xsl:call-template name ="getColorCode">
              <xsl:with-param name ="color">
                <xsl:value-of select="p:txBody//a:solidFill/a:schemeClr/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumMod">
                <xsl:choose>
                  <xsl:when test="p:txBody//a:solidFill/a:schemeClr/a:lumMod">
                    <xsl:value-of select="p:txBody//a:solidFill/a:schemeClr/a:lumMod/@val" />
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="''" />
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:with-param>
              <xsl:with-param name ="lumOff">
                <xsl:choose>
                  <xsl:when test="p:txBody//a:solidFill/a:schemeClr/a:lumOff">
                    <xsl:value-of select="p:txBody//a:solidFill/a:schemeClr/a:lumOff/@val" />
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


    </xsl:for-each>
  </xsl:template>
  <xsl:template name="TitleSubTitleParaProperty">
    <xsl:attribute name ="fo:margin-left">
      <xsl:choose>
        <xsl:when test="./@marL">
          <xsl:value-of select ="concat(format-number(./@marL div 360000 ,'#.##'),'cm')"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select ="'0cm'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name ="fo:margin-right">
      <xsl:choose>
        <xsl:when test="./@marR">
          <xsl:value-of select ="concat(format-number(./@marR div 360000 ,'#.##'),'cm')"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select ="'0cm'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:if test ="./a:spcBef/a:spcPts/@val">
      <xsl:attribute name ="fo:margin-top">
        <xsl:value-of select="concat(format-number(./a:spcBef/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="./a:spcAft/a:spcPts/@val">
      <xsl:attribute name ="fo:margin-bottom">
        <xsl:value-of select="concat(format-number(./a:spcAft/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:attribute name ="fo:text-indent">
      <xsl:choose>
        <xsl:when test="./@indent">
          <xsl:value-of select ="concat(format-number(./@indent div 360000 ,'#.##'),'cm')"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select ="'0cm'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:choose>
      <!-- Center Alignment-->
      <xsl:when test ="./@algn ='ctr'">
        <xsl:attribute name ="fo:text-align">
          <xsl:value-of select ="'center'"/>
        </xsl:attribute>
      </xsl:when>
      <!-- Right Alignment-->
      <xsl:when test ="./@algn ='r'">
        <xsl:attribute name ="fo:text-align">
          <xsl:value-of select ="'end'"/>
        </xsl:attribute>
      </xsl:when>
      <!-- Left Alignment-->
      <xsl:when test ="./@algn ='l'">
        <xsl:attribute name ="fo:text-align">
          <xsl:value-of select ="'start'"/>
        </xsl:attribute>
      </xsl:when>
      <!-- Justified Alignment-->
      <xsl:when test ="./@algn ='just'">
        <xsl:attribute name ="fo:text-align">
          <xsl:value-of select ="'justified'"/>
        </xsl:attribute>
      </xsl:when>
      <!-- end-->
    </xsl:choose>
  </xsl:template>
  <xsl:template name="TitleSubTitleTextStyle">
    <!--for getting font color-->
    <xsl:choose>
      <xsl:when test="./a:defRPr/a:solidFill/a:srgbClr">
        <xsl:attribute name ="fo:color">
          <xsl:value-of select ="concat('#',./a:defRPr/a:solidFill/a:srgbClr/@val)"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="./a:defRPr/a:solidFill/a:schemeClr">
        <xsl:attribute name ="fo:color">
          <xsl:call-template name ="getColorCode">
            <xsl:with-param name ="color">
              <xsl:value-of select="./a:defRPr/a:solidFill/a:schemeClr/@val"/>
            </xsl:with-param>
            <xsl:with-param name ="lumMod">
              <xsl:if test="./a:defRPr/a:solidFill/a:schemeClr/a:lumMod">
                <xsl:value-of select="./a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val" />
              </xsl:if>
            </xsl:with-param>
            <xsl:with-param name ="lumOff">
              <xsl:if test="./a:defRPr/a:solidFill/a:schemeClr/a:lumOff">
                <xsl:value-of select="./a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val" />
              </xsl:if>
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
    <xsl:if test="./a:defRPr/@sz">
      <xsl:attribute name ="fo:font-size">
        <xsl:value-of select ="concat(./a:defRPr/@sz div 100 ,'pt')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:choose>
      <xsl:when test="./a:defRPr/@b='1'">
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
    <xsl:choose>
      <xsl:when test="./a:defRPr/@i='1'">
        <xsl:attribute name ="fo:font-style">
          <xsl:value-of select ="'italic'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name ="fo:font-style">
          <xsl:value-of select ="'normal'"/>
        </xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:choose>
      <xsl:when test="./a:defRPr/@u='sng'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'solid'"></xsl:value-of>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-color">
          <xsl:value-of select ="'font-color'"></xsl:value-of>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'auto'"></xsl:value-of>
        </xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'none'"/>
        </xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpShapeParaProp">
    <xsl:param name="lnSpcReduction"/>
    <!--Text alignment-->
    <xsl:if test ="@algn">
      <xsl:attribute name ="fo:text-align">
        <xsl:choose>
          <!-- Center Alignment-->
          <xsl:when test ="@algn ='ctr'">
            <xsl:value-of select ="'center'"/>
          </xsl:when>
          <!-- Right Alignment-->
          <xsl:when test ="@algn ='r'">
            <xsl:value-of select ="'end'"/>
          </xsl:when>
          <!-- Left Alignment-->
          <xsl:when test ="@algn ='l'">
            <xsl:value-of select ="'start'"/>
          </xsl:when>
          <!-- Added by lohith - for fix 1737161-->
          <xsl:when test ="@algn ='just'">
            <xsl:value-of select ="'justify'"/>
          </xsl:when>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if >
    <!-- Convert Laeft margin of the paragraph-->
    <xsl:if test ="@marL">
      <xsl:attribute name ="fo:margin-left">
        <xsl:value-of select="concat(format-number(@marL div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="@marR">
      <xsl:attribute name ="fo:margin-right">
        <xsl:value-of select="concat(format-number(@marR div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="./@indent">
      <xsl:attribute name ="fo:text-indent">
        <xsl:value-of select="concat(format-number(./@indent div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if >
    <xsl:variable name="var_marL">
      <xsl:for-each select=".">
        <xsl:value-of select="@marL"/>
      </xsl:for-each>
    </xsl:variable>
    <xsl:for-each select="./a:spcBef/a:spcPct">
      <xsl:if test ="@val">
        <xsl:attribute name ="fo:margin-top">
          <xsl:value-of select="concat(format-number( ( $var_marL div  @val ) div 360000 ,'#.###'),'cm')"/>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
    <xsl:for-each select="./a:spcBef/a:spcPts">
      <xsl:if test ="@val">
        <xsl:attribute name ="fo:margin-top">
          <xsl:value-of select="concat(format-number(@val div 2835 ,'#.##'),'cm')"/>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
    <xsl:for-each select="./a:spcAft/a:spcPct">
      <xsl:if test ="@val">
        <xsl:attribute name ="fo:margin-bottom">
          <xsl:value-of select="concat(format-number(@val div 2976 div 28.35 ,'#.##'),'cm')"/>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
    <xsl:for-each select="./a:spcAft/a:spcPts">
      <xsl:if test ="@val">
        <xsl:attribute name ="fo:margin-bottom">
          <xsl:value-of select="concat(format-number(@val div  2835 ,'#.##'),'cm')"/>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
    <xsl:for-each select="./a:lnSpc/a:spcPct">
      <xsl:if test ="a:lnSpc/a:spcPct">
        <xsl:choose>
          <xsl:when test="$lnSpcReduction='0'">
            <xsl:attribute name="fo:line-height">
              <xsl:value-of select="concat(format-number(@val div 1000,'###'), '%')"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="fo:line-height">
              <xsl:value-of select="concat(format-number((@val - $lnSpcReduction) div 1000,'###'), '%')"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if >
    </xsl:for-each>
    <xsl:for-each select="./a:lnSpc/a:spcPts">
      <xsl:if test ="a:lnSpc/a:spcPts">
        <xsl:attribute name="style:line-height-at-least">
          <xsl:value-of select="concat(format-number(@val div 2835, '#.##'), 'cm')"/>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="tmpShapeTextProperty">
    <xsl:param name="LayoutFileName"/>
    <xsl:param name="spType"/>
    <xsl:param name="index"/>
    <xsl:param name ="DefFont"/>
    <xsl:param name ="fontType"/>
    <xsl:variable name="fontscale">
      <xsl:choose>
        <xsl:when test="p:txBody/a:bodyPr/a:normAutofit/@fontScale">
          <xsl:value-of select="p:txBody/a:bodyPr/a:normAutofit/@fontScale"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'0'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:if test ="./a:defRPr/@sz">
      <xsl:attribute name ="fo:font-size"	>
        <xsl:value-of  select ="concat(format-number(./a:defRPr/@sz div 100,'#.##'),'pt')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:attribute name ="fo:font-family">
      <xsl:if test ="./a:defRPr/a:latin/@typeface">
        <xsl:variable name ="typeFaceVal" select ="./a:defRPr/a:latin/@typeface"/>
        <xsl:for-each select ="./a:defRPr/a:latin/@typeface">
          <xsl:if test ="$typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt'">
            <xsl:call-template name="slideMasterFontName">
              <xsl:with-param name="fontType">
                <xsl:value-of select="$fontType"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:if>
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
    <!-- strike style:text-line-through-style-->
    <xsl:if test ="./a:defRPr/@strike">
      <xsl:attribute name ="style:text-line-through-style">
        <xsl:value-of select ="'solid'"/>
      </xsl:attribute>
      <xsl:choose >
        <xsl:when test ="./a:defRPr/@strike='dblStrike'">
          <xsl:attribute name ="style:text-line-through-type">
            <xsl:value-of select ="'double'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test ="./a:defRPr/@strike='sngStrike'">
          <xsl:attribute name ="style:text-line-through-type">
            <xsl:value-of select ="'single'"/>
          </xsl:attribute>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
    <!-- Kening Property-->
    <xsl:if test ="./a:defRPr/@kern">
      <xsl:choose >
        <xsl:when test ="./a:defRPr/@kern = '0'">
          <xsl:attribute name ="style:letter-kerning">
            <xsl:value-of select ="'false'"/>
          </xsl:attribute >
        </xsl:when >
        <xsl:when test ="format-number(./a:defRPr/@kern,'#.##') &gt; 0">
          <xsl:attribute name ="style:letter-kerning">
            <xsl:value-of select ="'true'"/>
          </xsl:attribute >
        </xsl:when >
      </xsl:choose>
    </xsl:if >
    <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
    <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
    <xsl:if test ="./a:defRPr/a:solidFill/a:srgbClr/@val">
      <xsl:attribute name ="fo:color">
        <xsl:value-of select ="translate(concat('#',./a:defRPr/a:solidFill/a:srgbClr/@val),$ucletters,$lcletters)"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="./a:defRPr/a:solidFill/a:schemeClr/@val">
      <xsl:attribute name ="fo:color">
        <xsl:call-template name ="getColorCode">
          <xsl:with-param name ="color">
            <xsl:value-of select="./a:defRPr/a:solidFill/a:schemeClr/@val"/>
          </xsl:with-param>
          <xsl:with-param name ="lumMod">
            <xsl:value-of select="./a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
          </xsl:with-param>
          <xsl:with-param name ="lumOff">
            <xsl:value-of select="./a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:choose>
      <xsl:when test="./a:defRPr/@b='1'">
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
    <!--UnderLine-->
    <xsl:if test ="./a:defRPr/@u">
      <xsl:for-each select ="./a:defRPr">
        <xsl:call-template name="tmpUnderLine"/>
      </xsl:for-each>
    </xsl:if >
    <xsl:if test ="not(./a:defRPr/@u)">
      <xsl:attribute name ="style:text-underline-style">
        <xsl:value-of select ="'none'"/>
      </xsl:attribute>
    </xsl:if>
    <!-- Italic-->
    <xsl:attribute name ="fo:font-style">
      <xsl:if test ="@i='1'">
        <xsl:value-of select ="'italic'"/>
      </xsl:if >
      <xsl:if test ="@i='0'">
        <xsl:value-of select ="'normal'"/>
      </xsl:if >
      <xsl:if test ="not(@i)">
        <xsl:value-of select ="'normal'"/>
      </xsl:if >
    </xsl:attribute>
    <!-- Character Spacing -->
    <xsl:if test ="./a:defRPr/@spc">
      <xsl:attribute name ="fo:letter-spacing">
        <xsl:variable name="length" select="./a:defRPr/@spc" />
        <xsl:value-of select="concat(format-number($length * 2.54 div 7200,'#.###'),'cm')"/>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>
  <xsl:template name="tmpOutlineParaProp">
    <xsl:param name="lnSpcReduction"/>
    <!--Text alignment-->
    <xsl:if test ="@algn">
      <xsl:attribute name ="fo:text-align">
        <xsl:choose>
          <!-- Center Alignment-->
          <xsl:when test ="@algn ='ctr'">
            <xsl:value-of select ="'center'"/>
          </xsl:when>
          <!-- Right Alignment-->
          <xsl:when test ="@algn ='r'">
            <xsl:value-of select ="'end'"/>
          </xsl:when>
          <!-- Left Alignment-->
          <xsl:when test ="@algn ='l'">
            <xsl:value-of select ="'start'"/>
          </xsl:when>
          <!-- Added by lohith - for fix 1737161-->
          <xsl:when test ="@algn ='just'">
            <xsl:value-of select ="'justify'"/>
          </xsl:when>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if >
    <!-- Convert Laeft margin of the paragraph-->
    <xsl:if test ="@marL">
      <xsl:attribute name ="fo:margin-left">
        <xsl:value-of select="concat(format-number(@marL div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="@marR">
      <xsl:attribute name ="fo:margin-right">
        <xsl:value-of select="concat(format-number(@marR div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="@indent">
      <xsl:attribute name ="fo:text-indent">
        <xsl:value-of select="concat(format-number(@indent div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if >

    <xsl:if test ="a:spcBef/a:spcPts/@val">
      <xsl:attribute name ="fo:margin-top">
        <xsl:value-of select="concat(format-number(a:spcBef/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="a:spcAft/a:spcPts/@val">
      <xsl:attribute name ="fo:margin-bottom">
        <xsl:value-of select="concat(format-number(a:spcAft/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="a:lnSpc/a:spcPct/@val">
      <xsl:choose>
        <xsl:when test="$lnSpcReduction='0'">
          <xsl:attribute name="fo:line-height">
            <xsl:value-of select="concat(format-number(a:lnSpc/a:spcPct/@val div 1000,'###'), '%')"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="fo:line-height">
            <xsl:value-of select="concat(format-number((a:lnSpc/a:spcPct/@val - $lnSpcReduction) div 1000,'###'), '%')"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if >
    <xsl:if test ="a:lnSpc/a:spcPts/@val">
      <xsl:attribute name="style:line-height-at-least">
        <xsl:value-of select="concat(format-number(a:lnSpc/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>
  <xsl:template name ="TextParagraphProp" >
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
        <!--@ Default Font Name from PPTX -->
        <xsl:variable name ="DefFont">
          <xsl:for-each select ="document('ppt/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme
						/a:majorFont/a:latin/@typeface">
            <xsl:value-of select ="."/>
          </xsl:for-each>
        </xsl:variable>
        <xsl:variable name ="DefFontSizeTitle">
          <xsl:for-each select ="document('ppt/slideMasters/slideMaster1.xml')//p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/@sz">
            <xsl:value-of select ="."/>
          </xsl:for-each>
        </xsl:variable>
        <xsl:variable name ="DefFontSizeBody">
          <xsl:for-each select ="document('ppt/slideMasters/slideMaster1.xml')//p:txStyles/p:bodyStyle/a:lvl1pPr/a:defRPr/@sz">
            <xsl:value-of select ="."/>
          </xsl:for-each>
        </xsl:variable>
        <xsl:variable name ="DefFontSizeOther">
          <xsl:for-each select ="document('ppt/slideMasters/slideMaster1.xml')//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/@sz">
            <xsl:value-of select ="."/>
          </xsl:for-each>
        </xsl:variable>
        <!--@Modified font Names -->
        <xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMasterPath))//p:sp/p:txBody">
          <xsl:variable name ="ParaId">
            <xsl:value-of select ="concat($slideMasterName ,concat('PARA',position()))"/>
          </xsl:variable>
          <!--  by vijayeta,to get linespacing from layouts-->
          <xsl:variable name ="layoutName">
            <xsl:value-of select ="./parent::node()/p:nvSpPr/p:nvPr/p:ph/@type"/>
          </xsl:variable>
          <!--Code by Vijayeta for Bullets,set style name in case of default bullets-->
          <xsl:variable name ="listStyleName">
            <!-- Added by vijayeta, to get the text box number-->
            <xsl:variable name ="textNumber" select ="./parent::node()/p:nvSpPr/p:cNvPr/@id"/>
            <!-- Added by vijayeta, to get the text box number-->
            <xsl:choose>
              <xsl:when test ="./parent::node()/p:spPr/a:prstGeom/@prst or ./parent::node()/child::node()[1]/child::node()[1]/@name[contains(., 'TextBox')] or ./parent::node()/child::node()[1]/child::node()[1]/@name[contains(., 'Text Box')]">
                <xsl:value-of select ="concat('sm1','textboxshape_List',$textNumber)"/>
              </xsl:when>
              <xsl:otherwise >
                <xsl:value-of select ="concat('sm1','List',position())"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <!--If bullets present-->
          <xsl:if test ="(a:p/a:pPr/a:buChar) or (a:p/a:pPr/a:buAutoNum) or (a:p/a:pPr/a:buBlip) ">
            <xsl:call-template name ="insertBulletStyle">
              <xsl:with-param name ="slideRel" select ="''"/>
              <xsl:with-param name ="ParaId" select ="$ParaId"/>
              <xsl:with-param name ="listStyleName" select ="$listStyleName"/>
            </xsl:call-template>
          </xsl:if>
          <xsl:variable name ="bulletTypeBool" select ="'true'"/>
          <!-- bullets are default-->
          <xsl:if test ="not(a:p/a:pPr/a:buChar or a:p/a:pPr/a:buAutoNum or a:p/a:pPr/a:buBlip) ">
            <xsl:if test ="$bulletTypeBool='true'">
              <!-- Added by  vijayeta ,on 19th june-->
              <!--Bug fix 1739611,by vijayeta,June 21st-->
              <xsl:if test ="./parent::node()/p:nvSpPr/p:cNvPr/@name[contains(.,'Content')]
							or ./parent::node()/p:nvSpPr/p:cNvPr/@name[contains(.,'Subtitle')]                   
							or ./parent::node()/p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'subTitle')]
							or ./parent::node()/p:nvSpPr/p:cNvPr/@name[contains(.,'outline')]
							or (./parent::node()/p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'body')] and ./parent::node()/p:nvSpPr/p:cNvPr/@name[contains(.,'Rectangle')])
							or (./parent::node()/p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'body')] and ./parent::node()/p:nvSpPr/p:cNvPr/@name[contains(.,'Content')])">
                <!-- Change made by vijayeta,on 9/7/07,cosider Rectangle as content-->
                <!-- Added by  vijayeta ,on 19th june-->
                <xsl:call-template name ="insertDefaultBulletNumberStyle">
                  <xsl:with-param name ="listStyleName" select ="$listStyleName"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:if>
          </xsl:if>
          <!--End of code if bullets are default-->
          <!--End of code inserted by Vijayeta,InsertStyle For Bullets and Numbering-->
          <xsl:for-each select ="a:p">
            <!-- Code by vijayeta,to set default font size for inner levels,in case of multiple levels-->
            <xsl:variable name ="levelForDefFont">
              <xsl:if test ="$bulletTypeBool='true'">
                <xsl:if test ="a:pPr/@lvl">
                  <xsl:value-of select ="a:pPr/@lvl"/>
                </xsl:if>
                <xsl:if test ="not(a:pPr/@lvl)">
                  <xsl:value-of select ="'0'"/>
                </xsl:if>
              </xsl:if>
            </xsl:variable>
            <!--End of Code by vijayeta,to set default font size for inner levels,in case of multiple levels-->
            <style:style style:family="paragraph">
              <xsl:attribute name ="style:name">
                <xsl:value-of select ="concat($ParaId,position())"/>
              </xsl:attribute >
              <style:paragraph-properties  text:enable-numbering="false" >
                <xsl:if test ="parent::node()/a:bodyPr/@vert='vert'">
                  <xsl:attribute name ="style:writing-mode">
                    <xsl:value-of select ="'tb-rl'"/>
                  </xsl:attribute>
                </xsl:if>
                <!-- commented by pradeep -->
                <!--start-->
                <xsl:if test ="a:pPr/@algn ='ctr' or a:pPr/@algn ='r' or a:pPr/@algn ='l' or a:pPr/@algn ='just'">
                  <!--End-->
                  <xsl:attribute name ="fo:text-align">
                    <xsl:choose>
                      <!-- Center Alignment-->
                      <xsl:when test ="a:pPr/@algn ='ctr'">
                        <xsl:value-of select ="'center'"/>
                      </xsl:when>
                      <!-- Right Alignment-->
                      <xsl:when test ="a:pPr/@algn ='r'">
                        <xsl:value-of select ="'end'"/>
                      </xsl:when>
                      <!-- Left Alignment-->
                      <xsl:when test ="a:pPr/@algn ='l'">
                        <xsl:value-of select ="'start'"/>
                      </xsl:when>
                      <!-- Added by lohith - for fix 1737161-->
                      <xsl:when test ="a:pPr/@algn ='just'">
                        <xsl:value-of select ="'justify'"/>
                      </xsl:when>
                      <!-- added by pradeep-->

                    </xsl:choose>
                  </xsl:attribute>
                  <!-- commented by pradeep -->
                  <!--start-->
                </xsl:if >
                <!--End-->
                <!-- Convert Laeft margin of the paragraph-->
                <xsl:if test ="a:pPr/@marL">
                  <xsl:attribute name ="fo:margin-left">
                    <xsl:value-of select="concat(format-number(a:pPr/@marL div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test ="a:pPr/@indent">
                  <xsl:attribute name ="fo:text-indent">
                    <xsl:value-of select="concat(format-number(a:pPr/@indent div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if >
                <xsl:if test ="a:pPr/a:spcBef/a:spcPts/@val">
                  <xsl:attribute name ="fo:margin-top">
                    <xsl:value-of select="concat(format-number(a:pPr/a:spcBef/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test ="a:pPr/a:spcAft/a:spcPts/@val">
                  <xsl:attribute name ="fo:margin-bottom">
                    <xsl:value-of select="concat(format-number(a:pPr/a:spcAft/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
                <!--Code Added By Vijayeta,
              Insert Spaces between Lines in a paragraph,
              based on Percentage spacing or Point spacing as defined in an input PPTX file.-->
                <!-- If the line space is in Percentage-->
                <xsl:if test ="a:pPr/a:lnSpc/a:spcPct/@val">
                  <xsl:attribute name="fo:line-height">
                    <xsl:value-of select="concat(format-number(a:pPr/a:lnSpc/a:spcPct/@val div 1000,'###'), '%')"/>
                  </xsl:attribute>
                </xsl:if >
                <!-- If the line space is in Points-->
                <xsl:if test ="a:pPr/a:lnSpc/a:spcPts">
                  <xsl:attribute name="style:line-height-at-least">
                    <xsl:value-of select="concat(format-number(a:pPr/a:lnSpc/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
                <!-- End of Code Added By Vijayeta,to get line spacing from layout or master, dated 9-7-07-->
                <!-- Code inserted by VijayetaFor Bullets, Enable Numbering-->
                <xsl:if test ="a:pPr/a:buChar or a:pPr/a:buAutoNum or a:pPr/a:buBlip ">
                  <xsl:choose >
                    <xsl:when test ="not(a:r/a:t)">
                      <xsl:attribute name="text:enable-numbering">
                        <xsl:value-of select ="'false'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise >
                      <xsl:attribute name="text:enable-numbering">
                        <xsl:value-of select ="'true'"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if>
                <!--</xsl:if>-->
                <xsl:if test ="not(a:pPr/a:buChar) and not(a:pPr/a:buAutoNum) and not(a:pPr/a:buBlip) ">
                  <xsl:if test ="$bulletTypeBool='true'">
                    <!-- Added by  vijayeta ,on 19th june-->
                    <xsl:choose >
                      <!--Bug fix 1739611,by vijayeta,June 21st-->
                      <xsl:when test ="./parent::node()/parent::node()/p:nvSpPr/p:cNvPr/@name[contains(.,'Content')]
                   or ./parent::node()/parent::node()/p:nvSpPr/p:cNvPr/@name[contains(.,'Subtitle')]                   
                    or ./parent::node()/parent::node()/p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'subTitle')]
                    or ./parent::node()/parent::node()/p:nvSpPr/p:cNvPr/@name[contains(.,'outline')]
                              or (./parent::node()/parent::node()/p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'body')] and ./parent::node()/parent::node()/p:nvSpPr/p:cNvPr/@name[contains(.,'Rectangle')])
                              or (./parent::node()/p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'body')] and ./parent::node()/p:nvSpPr/p:cNvPr/@name[contains(.,'Content')])">
                        <!-- Change made by vijayeta,on 9-7-07, Rectangle=content-->
                        <!-- Added by vijayeta on 19th june,to enable or disable depending on buNone-->
                        <xsl:if test ="a:r/a:t">
                          <xsl:if test ="a:pPr/a:buNone">
                            <xsl:attribute name="text:enable-numbering">
                              <xsl:value-of select ="'false'"/>
                            </xsl:attribute>
                          </xsl:if>
                          <xsl:if test ="not(a:pPr/a:buNone)">
                            <xsl:attribute name="text:enable-numbering">
                              <xsl:value-of select ="'true'"/>
                            </xsl:attribute>
                          </xsl:if>
                        </xsl:if>
                        <xsl:if test ="not(a:r/a:t)">
                          <xsl:attribute name="text:enable-numbering">
                            <xsl:value-of select ="'false'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <!-- Added by vijayeta on 19th june,to enable or disable depending on buNone-->
                      </xsl:when>
                      <xsl:otherwise >
                        <xsl:attribute name="text:enable-numbering">
                          <xsl:value-of select ="'false'"/>
                        </xsl:attribute>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:if>
                </xsl:if>
                <!--End of Code inserted by VijayetaFor Bullets, Enable Numbering-->
                <!-- @@Code for Paragraph tabs Pradeep Nemadi -->
                <!-- Starts-->
                <xsl:if test ="a:pPr/a:tabLst/a:tab">
                  <xsl:call-template name ="paragraphTabstops" />
                </xsl:if>
                <!-- Ends -->
              </style:paragraph-properties >
            </style:style>
            <!-- Modified by pradeep for fix 1731885-->
            <xsl:for-each select ="node()" >
              <!-- Add here-->
              <xsl:if test ="name()='a:r'">
                <style:style style:name="{generate-id()}"  style:family="text">
                  <style:text-properties>
                    <!-- Bug 1711910 Fixed,On date 2-06-07,by Vijayeta-->
                    <xsl:attribute name ="fo:font-family">
                      <xsl:if test ="a:rPr/a:latin/@typeface">
                        <xsl:variable name ="typeFaceVal" select ="a:rPr/a:latin/@typeface"/>
                        <xsl:for-each select ="a:rPr/a:latin/@typeface">
                          <xsl:if test ="$typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt'">
                            <xsl:value-of select ="$DefFont"/>
                          </xsl:if>
                          <!-- Bug 1711910 Fixed,On date 2-06-07,by Vijayeta-->
                          <xsl:if test ="not($typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt')">
                            <xsl:value-of select ="."/>
                          </xsl:if>
                        </xsl:for-each>
                      </xsl:if>
                      <xsl:if test ="not(a:rPr/a:latin/@typeface)">
                        <!-- ADDED BY VIJAYETA,get font types for each level-->
                        <!--<xsl:for-each select ="document(concat('ppt/slideLayouts/',$layout))//p:sldLayout"-->
                        <xsl:variable name ="cnvPrId">
                          <xsl:value-of select ="parent::node()/parent::node()/parent::node()/p:nvSpPr/p:nvPr/p:ph/@type"/>
                        </xsl:variable>
                        <xsl:variable name ="Id">
                          <xsl:value-of select ="parent::node()/parent::node()/parent::node()/p:nvSpPr/p:nvPr/p:ph/@idx"/>
                        </xsl:variable>
                        <xsl:call-template name ="FontTypeFromLayout">
                          <xsl:with-param name ="prId" select ="$cnvPrId" />
                          <xsl:with-param name ="sldName" select ="''"/>
                          <xsl:with-param name ="Id" select ="$Id"/>
                          <xsl:with-param name ="levelForDefFont" select ="$levelForDefFont"/>
                          <xsl:with-param name ="DefFont" select ="$DefFont"/>
                        </xsl:call-template>
                        <!-- ADDED BY VIJAYETA,get font types for each level-->
                        <!--<xsl:value-of select ="$DefFont"/>-->
                      </xsl:if>
                    </xsl:attribute>
                    <xsl:attribute name ="style:font-family-generic"	>
                      <xsl:value-of select ="'roman'"/>
                    </xsl:attribute>
                    <xsl:attribute name ="style:font-pitch"	>
                      <xsl:value-of select ="'variable'"/>
                    </xsl:attribute>
                    <xsl:if test ="a:rPr/@sz">
                      <xsl:attribute name ="fo:font-size"	>
                        <xsl:for-each select ="a:rPr/@sz">
                          <xsl:value-of select ="concat(format-number(. div 100,'#.##'), 'pt')"/>
                        </xsl:for-each>
                      </xsl:attribute>
                    </xsl:if>
                    <xsl:attribute name ="fo:font-weight">
                      <!-- Bold Property-->
                      <xsl:if test ="a:rPr/@b">
                        <xsl:value-of select ="'bold'"/>
                      </xsl:if >
                      <xsl:if test ="not(a:rPr/@b)">
                        <xsl:value-of select ="'normal'"/>
                      </xsl:if >
                    </xsl:attribute>
                    <!-- Color -->
                    <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
                    <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
                    <xsl:if test ="a:rPr/a:solidFill/a:srgbClr/@val">
                      <xsl:attribute name ="fo:color">
                        <xsl:value-of select ="translate(concat('#',a:rPr/a:solidFill/a:srgbClr/@val),$ucletters,$lcletters)"/>
                      </xsl:attribute>
                    </xsl:if>
                    <xsl:if test ="a:rPr/a:solidFill/a:schemeClr/@val">
                      <xsl:attribute name ="fo:color">
                        <xsl:call-template name ="getColorCode">
                          <xsl:with-param name ="color">
                            <xsl:value-of select="a:rPr/a:solidFill/a:schemeClr/@val"/>
                          </xsl:with-param>
                          <xsl:with-param name ="lumMod">
                            <xsl:value-of select="a:rPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                          </xsl:with-param>
                          <xsl:with-param name ="lumOff">
                            <xsl:value-of select="a:rPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                          </xsl:with-param>
                        </xsl:call-template>
                      </xsl:attribute>
                    </xsl:if>
                    <!-- Bug fix - 1733229-->
                    <xsl:if test ="not(a:rPr/a:solidFill/a:srgbClr/@val) and not(a:rPr/a:solidFill/a:schemeClr/@val)">
                      <xsl:if test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
                        <xsl:attribute name ="fo:color">
                          <xsl:value-of select ="translate(concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val),$ucletters,$lcletters)"/>
                        </xsl:attribute>
                      </xsl:if>
                      <xsl:if test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
                        <xsl:attribute name ="fo:color">
                          <xsl:call-template name ="getColorCode">
                            <xsl:with-param name ="color">
                              <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr/@val"/>
                            </xsl:with-param>
                            <xsl:with-param name ="lumMod">
                              <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr/a:lumMod/@val"/>
                            </xsl:with-param>
                            <xsl:with-param name ="lumOff">
                              <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr/a:lumOff/@val"/>
                            </xsl:with-param>
                          </xsl:call-template>
                        </xsl:attribute>
                      </xsl:if>
                      <xsl:if test ="parent::node()/parent::node()/parent::node()/p:nvSpPr/p:cNvPr/@name[contains(.,'TextBox')]
													   and not(parent::node()/parent::node()/parent::node()/p:style/a:fontRef)">
                        <xsl:attribute name ="fo:color">
                          <xsl:value-of select ="'#000000'"/>
                        </xsl:attribute>
                      </xsl:if>

                    </xsl:if>
                    <!-- Italic-->
                    <xsl:if test ="a:rPr/@i">
                      <xsl:attribute name ="fo:font-style">
                        <xsl:value-of select ="'italic'"/>
                      </xsl:attribute>
                    </xsl:if>
                    <!-- style:text-underline-style
								style:text-underline-style="solid" style:text-underline-width="auto"-->
                    <xsl:if test ="a:rPr/@u">
                      <!-- style:text-underline-style="solid" style:text-underline-type="double"-->
                      <xsl:if test ="a:rPr/a:uFill/a:solidFill/a:srgbClr/@val">
                        <xsl:attribute name ="style:text-underline-color">
                          <xsl:value-of select ="translate(concat('#',a:rPr/a:uFill/a:solidFill/a:srgbClr/@val),$ucletters,$lcletters)"/>
                        </xsl:attribute>
                      </xsl:if>
                      <xsl:if test ="a:rPr/a:uFill/a:solidFill/a:schemeClr/@val">
                        <xsl:attribute name ="style:text-underline-color">
                          <xsl:call-template name ="getColorCode">
                            <xsl:with-param name ="color">
                              <xsl:value-of select="a:rPr/a:uFill/a:solidFill/a:schemeClr/@val"/>
                            </xsl:with-param>
                            <xsl:with-param name ="lumMod">
                              <xsl:value-of select="a:rPr/a:uFill/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                            </xsl:with-param>
                            <xsl:with-param name ="lumOff">
                              <xsl:value-of select="a:rPr/a:uFill/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                            </xsl:with-param>
                          </xsl:call-template>
                        </xsl:attribute>
                      </xsl:if>
                      <xsl:choose >
                        <xsl:when test ="a:rPr/@u='dbl'">
                          <xsl:attribute name ="style:text-underline-style">
                            <xsl:value-of select ="'solid'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="style:text-underline-type">
                            <xsl:value-of select ="'double'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="style:text-underline-width">
                            <xsl:value-of select ="'auto'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:when test ="a:rPr/@u='heavy'">
                          <xsl:attribute name ="style:text-underline-style">
                            <xsl:value-of select ="'solid'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="style:text-underline-width">
                            <xsl:value-of select ="'bold'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:when test ="a:rPr/@u='dotted'">
                          <xsl:attribute name ="style:text-underline-style">
                            <xsl:value-of select ="'dotted'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="style:text-underline-width">
                            <xsl:value-of select ="'auto'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <!-- dottedHeavy-->
                        <xsl:when test ="a:rPr/@u='dottedHeavy'">
                          <xsl:attribute name ="style:text-underline-style">
                            <xsl:value-of select ="'dotted'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="style:text-underline-width">
                            <xsl:value-of select ="'bold'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:when test ="a:rPr/@u='dash'">
                          <xsl:attribute name ="style:text-underline-style">
                            <xsl:value-of select ="'dash'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="style:text-underline-width">
                            <xsl:value-of select ="'auto'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:when test ="a:rPr/@u='dashHeavy'">
                          <xsl:attribute name ="style:text-underline-style">
                            <xsl:value-of select ="'dash'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="style:text-underline-width">
                            <xsl:value-of select ="'bold'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:when test ="a:rPr/@u='dashLong'">
                          <xsl:attribute name ="style:text-underline-style">
                            <xsl:value-of select ="'long-dash'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="style:text-underline-width">
                            <xsl:value-of select ="'auto'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:when test ="a:rPr/@u='dashLongHeavy'">
                          <xsl:attribute name ="style:text-underline-style">
                            <xsl:value-of select ="'long-dash'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="style:text-underline-width">
                            <xsl:value-of select ="'bold'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:when test ="a:rPr/@u='dotDash'">
                          <xsl:attribute name ="style:text-underline-style">
                            <xsl:value-of select ="'dot-dash'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="style:text-underline-width">
                            <xsl:value-of select ="'auto'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:when test ="a:rPr/@u='dotDashHeavy'">
                          <xsl:attribute name ="style:text-underline-style">
                            <xsl:value-of select ="'dot-dash'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="style:text-underline-width">
                            <xsl:value-of select ="'bold'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <!-- dot-dot-dash-->
                        <xsl:when test ="a:rPr/@u='dotDotDash'">
                          <xsl:attribute name ="style:text-underline-style">
                            <xsl:value-of select ="'dot-dot-dash'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="style:text-underline-width">
                            <xsl:value-of select ="'auto'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:when test ="a:rPr/@u='dotDotDashHeavy'">
                          <xsl:attribute name ="style:text-underline-style">
                            <xsl:value-of select ="'dot-dot-dash'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="style:text-underline-width">
                            <xsl:value-of select ="'bold'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <!-- Wavy and Heavy-->
                        <xsl:when test ="a:rPr/@u='wavy'">
                          <xsl:attribute name ="style:text-underline-style">
                            <xsl:value-of select ="'wave'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="style:text-underline-width">
                            <xsl:value-of select ="'auto'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:when test ="a:rPr/@u='wavyHeavy'">
                          <xsl:attribute name ="style:text-underline-style">
                            <xsl:value-of select ="'wave'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="style:text-underline-width">
                            <xsl:value-of select ="'bold'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <!-- wavyDbl-->
                        <!-- style:text-underline-style="wave" style:text-underline-type="double"-->
                        <xsl:when test ="a:rPr/@u='wavyDbl'">
                          <xsl:attribute name ="style:text-underline-style">
                            <xsl:value-of select ="'wave'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="style:text-underline-type">
                            <xsl:value-of select ="'double'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="style:text-underline-width">
                            <xsl:value-of select ="'auto'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:otherwise >
                          <xsl:attribute name ="style:text-underline-type">
                            <xsl:value-of select ="'single'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="style:text-underline-width">
                            <xsl:value-of select ="'auto'"/>
                          </xsl:attribute>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:if>
                    <!-- strike style:text-line-through-style-->
                    <xsl:if test ="a:rPr/@strike">
                      <xsl:attribute name ="style:text-line-through-style">
                        <xsl:value-of select ="'solid'"/>
                      </xsl:attribute>
                      <xsl:choose >
                        <xsl:when test ="a:rPr/@strike='dblStrike'">
                          <xsl:attribute name ="style:text-line-through-type">
                            <xsl:value-of select ="'double'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:when test ="a:rPr/@strike='sngStrike'">
                          <xsl:attribute name ="style:text-line-through-type">
                            <xsl:value-of select ="'single'"/>
                          </xsl:attribute>
                        </xsl:when>
                      </xsl:choose>
                    </xsl:if>
                    <!-- Character Spacing fo:letter-spacing Bug (1719230) fix by Sateesh -->
                    <xsl:if test ="a:rPr/@spc">
                      <xsl:attribute name ="fo:letter-spacing">
                        <xsl:variable name="length" select="a:rPr/@spc" />
                        <xsl:value-of select="concat(format-number($length * 2.54 div 7200,'#.###'),'cm')"/>
                      </xsl:attribute>
                    </xsl:if>
                    <!--Shadow fo:text-shadow-->
                    <xsl:if test ="a:rPr/a:effectLst/a:outerShdw">
                      <xsl:attribute name ="fo:text-shadow">
                        <xsl:value-of select ="'1pt 1pt'"/>
                      </xsl:attribute>
                    </xsl:if>
                    <!--Kerning true or false -->
                    <xsl:attribute name ="style:letter-kerning">
                      <xsl:choose >
                        <xsl:when test ="a:rPr/@kern = '0'">
                          <xsl:value-of select ="'false'"/>
                        </xsl:when >
                        <xsl:when test ="format-number(a:rPr/@kern,'#.##') &gt; 0">
                          <xsl:value-of select ="'true'"/>
                        </xsl:when >
                        <xsl:otherwise >
                          <xsl:value-of select ="'true'"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:attribute >
                  </style:text-properties>
                </style:style>
              </xsl:if>
              <!-- Added by lohith.ar for fix 1731885-->
              <xsl:if test ="name()='a:endParaRPr'">
                <style:style style:name="{generate-id()}"  style:family="text">
                  <style:text-properties>
                    <!-- Bug 1711910 Fixed,On date 2-06-07,by Vijayeta-->
                    <xsl:attribute name ="fo:font-family">
                      <xsl:if test ="a:endParaRPr/a:latin/@typeface">
                        <xsl:variable name ="typeFaceVal" select ="a:rPr/a:latin/@typeface"/>
                        <xsl:for-each select ="a:endParaRPr/a:latin/@typeface">
                          <xsl:if test ="$typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt'">
                            <xsl:value-of select ="$DefFont"/>
                          </xsl:if>
                          <!-- Bug 1711910 Fixed,On date 2-06-07,by Vijayeta-->
                          <xsl:if test ="not($typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt')">
                            <xsl:value-of select ="."/>
                          </xsl:if>
                        </xsl:for-each>
                      </xsl:if>
                      <xsl:if test ="not(a:endParaRPr/a:latin/@typeface)">
                        <xsl:value-of select ="$DefFont"/>
                      </xsl:if>
                    </xsl:attribute>
                    <xsl:attribute name ="style:font-family-generic"	>
                      <xsl:value-of select ="'roman'"/>
                    </xsl:attribute>
                    <xsl:attribute name ="style:font-pitch"	>
                      <xsl:value-of select ="'variable'"/>
                    </xsl:attribute>
                    <xsl:if test ="./@sz">
                      <xsl:attribute name ="fo:font-size"	>
                        <xsl:for-each select ="./@sz">
                          <xsl:value-of select ="concat(format-number(. div 100,'#.##'), 'pt')"/>
                        </xsl:for-each>
                      </xsl:attribute>
                    </xsl:if>
                  </style:text-properties>
                </style:style>
              </xsl:if>
            </xsl:for-each >
          </xsl:for-each>
        </xsl:for-each>
      </xsl:for-each>
    </xsl:for-each >

  </xsl:template>
</xsl:stylesheet>
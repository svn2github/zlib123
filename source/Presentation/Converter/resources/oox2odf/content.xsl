<?xml version="1.0" encoding="UTF-8"?>
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
<xsl:stylesheet version="1.0" 
xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
xmlns:xlink="http://www.w3.org/1999/xlink"
xmlns:dom="http://www.w3.org/2001/xml-events" 
xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
xmlns:rels="http://schemas.openxmlformats.org/package/2006/relationships"
xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
xmlns:dc="http://purl.org/dc/elements/1.1/"
xmlns:dcterms="http://purl.org/dc/terms/"
exclude-result-prefixes="p a r xlink ">
  <xsl:import href="common.xsl"/>
  <xsl:import href="shapes_reverse.xsl"/>
  <!-- Import Bullets and numbering-->
  <xsl:import href ="BulletsNumberingoox2odf.xsl"/>
  <xsl:strip-space elements="*"/>

  <!--main document-->
  <xsl:template name="content">
    <office:document-content>
      <office:automatic-styles>
        <!-- automatic styles for document body -->
        <xsl:call-template name ="textSpacingProp"/>
        <xsl:call-template name ="DateFormats"/>
        <xsl:call-template name="InsertStyles"/>
        <!-- Graphic properties for shapes -->
        <xsl:call-template name ="InsertStylesForGraphicProperties"/>
      </office:automatic-styles>
      <office:body>
        <office:presentation>
          <xsl:call-template name="InsertPresentationFooter"/>
          <xsl:call-template name="InsertDrawingPage"/>
          <!-- Added by Lohith A R : Custom Slide Show -->
          <xsl:call-template name="InsertPresentationSettings"/>
        </office:presentation>
      </office:body>
    </office:document-content>
  </xsl:template>

  <!--  generates automatic styles for paragraphs  how does it exactly work ?? -->
  <xsl:variable name ="flgFooter">
    <xsl:for-each select ="document('ppt/slides/slide1.xml')/p:sld/p:cSld/p:spTree/p:sp">
      <xsl:if test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'ftr')]" >
        <xsl:value-of select ="p:txBody/a:p/a:r/a:t"/>
      </xsl:if>
      <!--<xsl:value-of select ="."/>-->
    </xsl:for-each >
  </xsl:variable>
  <xsl:template name="InsertStyles">
    <!-- page Properties-->
    <xsl:for-each select ="document('ppt/presentation.xml')/p:presentation/p:sldIdLst/p:sldId">
      <xsl:variable name ="pageSlide">
        <xsl:value-of select ="concat(concat('ppt/slides/slide',position()),'.xml')"/>
      </xsl:variable>
      <!--added by vipul to get slide name-->
      <!--Start-->
      <xsl:variable name ="SlideFileName">
        <xsl:value-of select ="concat(concat('slide',position()),'.xml')"/>
      </xsl:variable>
      <xsl:variable name ="LayoutFileName">
        <xsl:call-template name="GetLayOutforSlide">
          <xsl:with-param name="SlideName" select="$SlideFileName">  </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <!--End-->

      <!-- Page settings like footer date slide number visible/Invisible-->
      <style:style  style:family="drawing-page">
        <xsl:attribute name ="style:name" >
          <xsl:value-of select ="concat('dp',position())"/>
        </xsl:attribute>
        <style:drawing-page-properties>
          <xsl:attribute name ="presentation:background-visible" >
            <xsl:value-of select ="'true'"/>
          </xsl:attribute>
          <xsl:attribute name ="presentation:background-objects-visible" >
            <xsl:value-of select ="'true'"/>
          </xsl:attribute>
          <xsl:attribute name ="presentation:display-footer" >
            <xsl:value-of select ="'true'"/>
          </xsl:attribute>
          <xsl:attribute name ="presentation:display-page-number" >
            <xsl:if test ="document($pageSlide)/p:sld/p:cSld/p:spTree/p:sp
						/p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'sldNum')]">
              <xsl:value-of select ="'true'"/>
            </xsl:if >
          </xsl:attribute>
          <xsl:attribute name ="presentation:display-date-time" >
            <xsl:value-of select ="'true'"/>
          </xsl:attribute>
          <!--Added by Sateesh for Background Color -->
          <xsl:variable name="Layout">
            <xsl:call-template name="GetLayOutforSlide">
              <xsl:with-param name="SlideName" select="$SlideFileName">  </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:choose>
            <!-- GET COLOR FROM SLIDES   -->
            <xsl:when test="document($pageSlide)/p:sld/p:cSld/p:bg/p:bgPr/a:solidFill/a:srgbClr/@val or document($pageSlide)/p:sld/p:cSld/p:bg/p:bgPr/a:solidFill/a:schemeClr/@val">
              <xsl:for-each select="document($pageSlide)/p:sld/p:cSld">
                <xsl:if test="p:bg/p:bgPr/a:solidFill/a:srgbClr/@val">
                  <xsl:attribute name="draw:fill-color">
                    <xsl:value-of select="concat('#',p:bg/p:bgPr/a:solidFill/a:srgbClr/@val)" />
                  </xsl:attribute>
                  <xsl:attribute name="draw:fill">
                    <xsl:value-of select="'solid'"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="p:bg/p:bgPr/a:solidFill/a:schemeClr/@val">
                  <xsl:attribute name="draw:fill-color">
                    <xsl:call-template name="getColorCode">
                      <xsl:with-param name="color">
                        <xsl:value-of select="p:bg/p:bgPr/a:solidFill/a:schemeClr/@val" />
                      </xsl:with-param>
                      <xsl:with-param name="lumMod">
                        <xsl:value-of select="p:bg/p:bgPr/a:solidFill/a:schemeClr/a:lumMod/@val" />
                      </xsl:with-param>
                      <xsl:with-param name="lumOff">
                        <xsl:value-of select="p:bg/p:bgPr/a:solidFill/a:schemeClr/a:lumOff/@val" />
                      </xsl:with-param>
                    </xsl:call-template>
                  </xsl:attribute>
                  <xsl:attribute name="draw:fill">
                    <xsl:value-of select="'solid'"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:when>
            <!-- GET COLOR FROM LAYOUT  -->
            <xsl:when test="document(concat('ppt/slideLayouts/',$Layout))//p:cSld/p:bg/p:bgPr/a:solidFill/a:srgbClr/@val or document(concat('ppt/slideLayouts/',$Layout))//p:cSld/p:bg/p:bgPr/a:solidFill/a:schemeClr/@val or document(concat('ppt/slideLayouts/',$Layout))//p:cSld/p:bg/p:bgRef/a:schemeClr/@val">
              <xsl:for-each select="document(concat('ppt/slideLayouts/',$Layout))//p:cSld">
                <xsl:choose>
                  <xsl:when test="p:bg/p:bgPr/a:solidFill/a:srgbClr/@val">
                    <xsl:attribute name="draw:fill-color">
                      <xsl:value-of select="concat('#',p:bg/p:bgPr/a:solidFill/a:srgbClr/@val)" />
                    </xsl:attribute>
                    <xsl:attribute name="draw:fill">
                      <xsl:value-of select="'solid'"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:when test="p:bg/p:bgPr/a:solidFill/a:schemeClr/@val">
                    <xsl:attribute name="draw:fill-color">
                      <xsl:call-template name="getColorCode">
                        <xsl:with-param name="color">
                          <xsl:value-of select="p:bg/p:bgPr/a:solidFill/a:schemeClr/@val" />
                        </xsl:with-param>
                        <xsl:with-param name="lumMod">
                          <xsl:value-of select="p:bg/p:bgPr/a:solidFill/a:schemeClr/a:lumMod/@val" />
                        </xsl:with-param>
                        <xsl:with-param name="lumOff">
                          <xsl:value-of select="p:bg/p:bgPr/a:solidFill/a:schemeClr/a:lumOff/@val" />
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                    <xsl:attribute name="draw:fill">
                      <xsl:value-of select="'solid'"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:when test="p:bg/p:bgRef/a:schemeClr/@val">
                    <xsl:attribute name="draw:fill-color">
                      <xsl:call-template name="getColorCode">
                        <xsl:with-param name="color">
                          <xsl:variable name="ClrMap">
                            <xsl:value-of select="p:bg/p:bgRef/a:schemeClr/@val" />
                          </xsl:variable>
                          <xsl:for-each select="parent::node()/p:clrMapOvr/a:overrideClrMapping">
                            <xsl:choose>
                              <xsl:when test="$ClrMap ='tx1'">
                                <xsl:value-of select="@tx1" />
                              </xsl:when>
                              <xsl:when test="$ClrMap ='tx2'">
                                <xsl:value-of select="@tx2" />
                              </xsl:when>
                              <xsl:when test="$ClrMap ='bg1'">
                                <xsl:value-of select="@bg1" />
                              </xsl:when>
                              <xsl:when test="$ClrMap ='bg2'">
                                <xsl:value-of select="@bg2" />
                              </xsl:when>
                            </xsl:choose>
                          </xsl:for-each>
                        </xsl:with-param>
                        <xsl:with-param name="lumMod">
                          <xsl:value-of select="p:bg/p:bgRef/a:schemeClr/a:lumMod/@val" />
                        </xsl:with-param>
                        <xsl:with-param name="lumOff">
                          <xsl:value-of select="p:bg/p:bgRef/a:schemeClr/a:lumOff/@val" />
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                    <xsl:attribute name="draw:fill">
                      <xsl:value-of select="'solid'"/>
                    </xsl:attribute>
                  </xsl:when>
                </xsl:choose>
              </xsl:for-each>
            </xsl:when>
          </xsl:choose>
          <!--End-->
        </style:drawing-page-properties>
      </style:style>
    </xsl:for-each >
    <style:style style:name="pr1" style:family="presentation" style:parent-style-name="Default-notes">
      <style:graphic-properties >
        <xsl:attribute name ="draw:fill-color" >
          <xsl:value-of select ="'#ffffff'"/>
        </xsl:attribute>
        <xsl:attribute name ="draw:auto-grow-height" >
          <xsl:value-of select ="'false'"/>
        </xsl:attribute>
        <xsl:attribute name ="fo:min-height" >
          <xsl:value-of select ="'12.573cm'"/>
        </xsl:attribute>
      </style:graphic-properties>
    </style:style>
    <xsl:call-template name ="GetStylesFromSlide"/>
  </xsl:template>
  <xsl:template name ="InsertPresentationFooter" >
    <xsl:for-each select ="document('ppt/presentation.xml')/p:presentation/p:sldIdLst/p:sldId">
      <xsl:variable name ="footerSlide">
        <xsl:value-of select ="concat(concat('ppt/slides/slide',position()),'.xml')"/>
      </xsl:variable>
      <xsl:variable name ="footerInd">
        <xsl:value-of select ="concat('ftr',position())"/>
      </xsl:variable>
      <xsl:variable name ="dateInd">
        <xsl:value-of select ="concat('dtd',position())"/>
      </xsl:variable>
      <xsl:for-each select ="document($footerSlide)/p:sld/p:cSld/p:spTree/p:sp">
        <!--concat(concat('slide',position()),'.xml')-->
        <xsl:choose >
          <xsl:when test ="not(p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'dt') or contains(.,'ftr')])">
            <!-- Do nothing-->
          </xsl:when>
          <xsl:when test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'ftr')]">
            <presentation:footer-decl >
              <xsl:attribute name ="presentation:name">
                <xsl:value-of select ="$footerInd"/>
              </xsl:attribute >
              <xsl:value-of select ="p:txBody/a:p/a:r/a:t"/>
            </presentation:footer-decl>
          </xsl:when>
          <xsl:when test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'dt')]">
            <presentation:date-time-decl >
              <xsl:attribute name ="style:data-style-name">
                <xsl:call-template name ="FooterDateFormat">
                  <xsl:with-param name ="type" select ="p:txBody/a:p/a:fld/@type" />
                </xsl:call-template>
              </xsl:attribute>
              <xsl:attribute name ="presentation:name">
                <xsl:value-of select ="$dateInd"/>
              </xsl:attribute>
              <xsl:attribute name ="presentation:source">
                <xsl:for-each select =".">
                  <xsl:if test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'dt')]" >
                    <xsl:if test ="p:txBody/a:p/a:fld">
                      <xsl:value-of select ="'current-date'"/>
                    </xsl:if>
                    <xsl:if test ="not(p:txBody/a:p/a:fld)">
                      <xsl:value-of select ="'fixed'"/>
                    </xsl:if>
                  </xsl:if>
                </xsl:for-each >
              </xsl:attribute>
              <xsl:for-each select =".">
                <xsl:if test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'dt')]" >
                  <xsl:if test ="p:txBody/a:p/a:fld">
                    <xsl:value-of select ="p:txBody/a:p/a:fld/a:t"/>
                  </xsl:if>
                  <xsl:if test ="not(p:txBody/a:p/a:fld)">
                    <xsl:value-of select ="p:txBody/a:p/a:r/a:t"/>
                  </xsl:if>
                </xsl:if>
              </xsl:for-each >
            </presentation:date-time-decl>
          </xsl:when >
        </xsl:choose>
      </xsl:for-each>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="InsertDrawingPage"  >
    <xsl:for-each select ="document('ppt/presentation.xml')/p:presentation/p:sldIdLst/p:sldId">
      <draw:page>
        <!--added by vipul to link each slides with slide Master-->
        <!--Start-->
        <xsl:attribute name="draw:master-page-name">
          <xsl:call-template name ="GetMasterFileName">
            <xsl:with-param name="slideId" select ="position()"/>
          </xsl:call-template>
        </xsl:attribute>
        <!--End-->
        <xsl:attribute name ="draw:style-name">
          <xsl:value-of select ="concat('dp',position())"/>
        </xsl:attribute>
        <xsl:attribute name ="draw:name">
          <xsl:value-of select ="concat('page',position())"/>
        </xsl:attribute>
        <xsl:attribute name ="presentation:use-footer-name">
          <xsl:value-of select ="concat('ftr',position())"/>
        </xsl:attribute>
        <xsl:attribute name ="presentation:use-date-time-name">
          <xsl:value-of select ="concat('dtd',position())"/>
        </xsl:attribute>
        <xsl:attribute name ="presentation:presentation-page-layout-name">
          <xsl:call-template name ="GetLayOutName">
            <xsl:with-param name ="slideRelName" select ="concat(concat('slide',position()),'.xml')"/>
          </xsl:call-template>
        </xsl:attribute>
        <xsl:call-template name ="DrawFrames">
          <xsl:with-param name ="SlideFile" select ="concat(concat('slide',position()),'.xml')" />
        </xsl:call-template>
      </draw:page >
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="DrawFrames">
    <xsl:param name ="SlideFile"/>
    <xsl:for-each  select="document('ppt/_rels/presentation.xml.rels')">
      <!--<xsl:for-each select ="./child::node()[1]">
				<xsl:value-of select ="."/>
			</xsl:for-each>-->
      <xsl:variable name ="slideNo">
        <xsl:value-of select ="concat('slides/',$SlideFile)"/>
      </xsl:variable>
      <xsl:variable name ="slideRel">
        <xsl:value-of select ="concat(concat('ppt/slides/_rels/',$SlideFile),'.rels')"/>
      </xsl:variable>
      <xsl:variable name ="slideLayotRel">
        <xsl:value-of select ="'ppt/slideMasters/_rels'"/>
      </xsl:variable>
      <!-- Slide Files Loop Main Loop-->
      <xsl:for-each select ="document(concat('ppt/',$slideNo))/p:sld/p:cSld/p:spTree/p:sp">
        <!-- Code for getting caps attribute from Layout-->
        <xsl:variable name ="layoutCap">
          <xsl:call-template name ="getAllCapsFromLayout">
            <xsl:with-param name ="layOutRelId" select ="$slideRel"/>
          </xsl:call-template>
        </xsl:variable>
        <!-- Code for getting caps attribute from Layout-->
        <!--Code Inserted by Vijayeta for Bullets And Numbering
            If the bullets are default,go to slide masters and get the bullets
            This part of code,checks for the type of layout and sets a falg true if condition met.
            If flag is set to true and no bullets present,then get the bullets type from slideMasters1.xml-->
        <xsl:variable name ="bulletTypeBool">
          <xsl:for-each select ="document($slideRel)//rels:Relationships/rels:Relationship">
            <xsl:if test ="document($slideRel)//rels:Relationships/rels:Relationship/@Target[contains(.,'slideLayouts')]">
              <xsl:variable name ="layout" select ="substring-after((@Target),'../slideLayouts/')"/>
              <xsl:for-each select ="document(concat('ppt/slideLayouts/',$layout))//p:sldLayout">
                <xsl:choose >
                  <!-- Changes made by vijayeta, bug fix, 1739703, date:10-7-07-->
                  <xsl:when test ="p:cSld/@name[contains(.,'Content')] or p:cSld/@name[contains(.,'Title and')] or p:cSld/@name[contains(.,'Two Content')] or  p:cSld/@name[contains(.,'Comparison')]">
                    <xsl:value-of select ="'true'"/>
                  </xsl:when>
                  <xsl:otherwise >
                    <xsl:value-of select ="'false'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:for-each>
            </xsl:if>
          </xsl:for-each>
        </xsl:variable >
        <!--End of code Inserted by Vijayeta for Bullets and Numbering,in case of default bullets-->
        <xsl:variable name ="LayoutName">
          <xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph/@type">
            <xsl:value-of select ="."/>
          </xsl:for-each>
        </xsl:variable>
        <xsl:variable name ="var_index">
          <xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph/@idx">
            <xsl:value-of select ="."/>
          </xsl:for-each>
        </xsl:variable>
        <xsl:variable name ="ParaId">
          <xsl:value-of select ="concat(substring($SlideFile,1,string-length($SlideFile)-4),concat('PARA',position()))"/>
        </xsl:variable>
        <xsl:variable name ="textLayoutId"  >
          <xsl:value-of  select ="concat(substring-before($SlideFile,'.xml'),'pr',position())"/>
        </xsl:variable>
        <!--code Inserted by Vijayeta for Bullets and Numbering,Assign a style name-->
        <xsl:variable name ="listStyleName">
          <xsl:value-of select ="concat(substring-before($SlideFile,'.xml'),'List',position())"/>
        </xsl:variable>
        <!--End of code Inserted by Vijayeta for Bullets and Numbering,Assign a style name-->
        <!-- Code for shapes start-->
        <!--<xsl:variable  name ="GraphicId">
					<xsl:value-of select ="concat('s',substring($SlideFile,6,string-length($SlideFile)-9) ,concat('gr',position()))"/>
				</xsl:variable>
				<xsl:call-template name ="shapes">
					<xsl:with-param name="GraphicId" select ="$GraphicId"/>
					<xsl:with-param name ="ParaId" select="$ParaId" />
					<xsl:with-param name ="SlideRelationId" select="$slideRel" />
				</xsl:call-template>-->
        <!-- Code for shapes end-->
        <xsl:choose >
          <xsl:when test ="not(contains(p:nvSpPr/p:cNvPr/@name,'Title')
						   or contains(p:nvSpPr/p:cNvPr/@name,'Content')
						   or contains(p:nvSpPr/p:cNvPr/@name,'Subtitle')
						  or contains(p:nvSpPr/p:cNvPr/@name,'Placeholder')
              or contains(p:nvSpPr/p:nvPr/p:ph/@type,'ctrTitle')
              or contains(p:nvSpPr/p:nvPr/p:ph/@type,'subTitle')
              or contains(p:nvSpPr/p:nvPr/p:ph/@type,'outline')
              or contains(p:nvSpPr/p:nvPr/p:ph/@type,'title')
              or contains(p:nvSpPr/p:nvPr/p:ph/@type,'body')
						  or p:nvSpPr/p:nvPr/p:ph/@idx)">
            <!-- Added ctrTitle,subtitle, outline, titel, body to fix the bug 1719280-->
          </xsl:when>
          <xsl:when test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'dt') 
							or contains(.,'ftr') or contains(.,'sldNum')]">
            <!-- Do nothing-->
            <!-- These will be covered in footer and date time -->
          </xsl:when>
          <xsl:when test ="p:spPr/a:xfrm/a:off" >
            <draw:frame draw:layer="layout" 																
										presentation:user-transformed="true">
              <xsl:attribute name ="presentation:style-name">
                <xsl:value-of select ="$textLayoutId"/>
              </xsl:attribute>
              <xsl:attribute name ="presentation:class">
                <xsl:call-template name ="LayoutType">
                  <xsl:with-param name ="LayoutStyle">
                    <xsl:value-of select ="$LayoutName"/>
                  </xsl:with-param>
                </xsl:call-template >
              </xsl:attribute>
              <xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cx]">
                <xsl:attribute name ="svg:width">
                  <xsl:call-template name="ConvertEmu">
                    <xsl:with-param name="length" select="@cx"/>
                    <xsl:with-param name="unit">cm</xsl:with-param>
                  </xsl:call-template>
                </xsl:attribute>
              </xsl:for-each>
              <xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cy]">
                <xsl:attribute name ="svg:height">
                  <xsl:call-template name="ConvertEmu">
                    <xsl:with-param name="length" select="@cy"/>
                    <xsl:with-param name="unit">cm</xsl:with-param>
                  </xsl:call-template>
                </xsl:attribute>
              </xsl:for-each>
              <xsl:for-each select ="p:spPr/a:xfrm/a:off[@x]">
                <xsl:attribute name ="svg:x">
                  <xsl:call-template name="ConvertEmu">
                    <xsl:with-param name="length" select="@x"/>
                    <xsl:with-param name="unit">cm</xsl:with-param>
                  </xsl:call-template>
                </xsl:attribute>
              </xsl:for-each>
              <xsl:for-each select ="p:spPr/a:xfrm/a:off[@y]">
                <xsl:attribute name ="svg:y">
                  <xsl:call-template name="ConvertEmu">
                    <xsl:with-param name="length" select="@y"/>
                    <xsl:with-param name="unit">cm</xsl:with-param>
                  </xsl:call-template>
                </xsl:attribute>
              </xsl:for-each>
              <draw:text-box>
                <xsl:for-each select ="p:txBody/a:p">
                  <xsl:if test ="$LayoutName != 'title'">
                    <!--Code Inserted by Vijayeta for Bullets And Numbering
                  check for levels and then depending on the condition,insert bullets,Layout or Master properties-->
                    <xsl:if test ="a:pPr/a:buChar or a:pPr/a:buAutoNum or a:pPr/a:buBlip">
                      <xsl:call-template name ="insertBulletsNumbersoox2odf">
                        <xsl:with-param name ="listStyleName" select ="concat($listStyleName,position())"/>
                        <xsl:with-param name ="ParaId" select ="$ParaId"/>
                      </xsl:call-template>
                    </xsl:if>
                    <!-- If no bullets are present or default bullets-->
                    <xsl:if test ="not(a:pPr/a:buChar) and not(a:pPr/a:buAutoNum)and not(a:pPr/a:buBlip)">
                      <!-- If bullets are default-->
                      <xsl:if test ="$bulletTypeBool='true'">
                        <xsl:call-template name ="insertBulletsNumbersoox2odf">
                          <xsl:with-param name ="listStyleName" select ="concat($listStyleName,position())"/>
                          <xsl:with-param name ="ParaId" select ="$ParaId"/>
                        </xsl:call-template>
                      </xsl:if>
                    </xsl:if>
                    <!-- If no bullets present at all-->
                    <xsl:if test ="not($bulletTypeBool='true')and not(a:pPr/a:buChar) and not(a:pPr/a:buAutoNum)and not(a:pPr/a:buBlip)">
                      <text:p >
                        <xsl:attribute name ="text:style-name">
                          <xsl:value-of select ="concat($ParaId,position())"/>
                        </xsl:attribute>
                        <xsl:for-each select ="node()">
                          <xsl:if test ="name()='a:r'">
                            <text:span text:style-name="{generate-id()}">
                              <!--<xsl:value-of select ="a:t"/>-->
                              <!--converts whitespaces sequence to text:s-->
                              <!-- 1699083 bug fix  -->
                              <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
                              <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
                              <xsl:choose >
                                <xsl:when test ="a:rPr[@cap='all'] or $layoutCap ='all'">
                                  <xsl:choose >
                                    <xsl:when test =".=''">
                                      <text:s/>
                                    </xsl:when>
                                    <xsl:when test ="not(contains(.,'  '))">
                                      <xsl:value-of select ="translate(.,$lcletters,$ucletters)"/>
                                    </xsl:when>
                                    <xsl:when test =". =' '">
                                      <text:s/>
                                    </xsl:when>
                                    <xsl:otherwise >
                                      <xsl:call-template name ="InsertWhiteSpaces">
                                        <xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
                                      </xsl:call-template>
                                    </xsl:otherwise>
                                  </xsl:choose>
                                </xsl:when>
                                <xsl:when test ="a:rPr[@cap='small'] or $layoutCap ='small'">
                                  <xsl:choose >
                                    <xsl:when test =".=''">
                                      <text:s/>
                                    </xsl:when>
                                    <xsl:when test ="not(contains(.,'  '))">
                                      <xsl:value-of select ="translate(.,$ucletters,$lcletters)"/>
                                    </xsl:when>
                                    <xsl:when test =".= ' '">
                                      <text:s/>
                                    </xsl:when>
                                    <xsl:otherwise >
                                      <xsl:call-template name ="InsertWhiteSpaces">
                                        <xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
                                      </xsl:call-template>
                                    </xsl:otherwise>
                                  </xsl:choose >
                                </xsl:when>
                                <xsl:otherwise >
                                  <xsl:choose >
                                    <xsl:when test =".=''">
                                      <text:s/>
                                    </xsl:when>
                                    <xsl:when test ="not(contains(.,'  '))">
                                      <xsl:value-of select ="."/>
                                    </xsl:when>
                                    <xsl:otherwise >
                                      <xsl:call-template name ="InsertWhiteSpaces">
                                        <xsl:with-param name ="string" select ="."/>
                                      </xsl:call-template>
                                    </xsl:otherwise >
                                  </xsl:choose>
                                </xsl:otherwise>
                              </xsl:choose>
                            </text:span>
                          </xsl:if >
                          <xsl:if test ="name()='a:br'">
                            <text:line-break/>
                          </xsl:if>
                          <!-- Added by lohith.ar for fix 1731885-->
                          <xsl:if test="name()='a:endParaRPr' and not(a:endParaRPr/a:hlinkClick)">
                            <text:span text:style-name="{generate-id()}" />
                          </xsl:if>
                        </xsl:for-each>
                      </text:p>
                    </xsl:if>
                  </xsl:if >
                  <xsl:if test ="$LayoutName = 'title'">
                    <text:p >
                      <xsl:attribute name ="text:style-name">
                        <xsl:value-of select ="concat($ParaId,position())"/>
                      </xsl:attribute>
                      <xsl:for-each select ="node()">
                        <xsl:if test ="name()='a:r'">
                          <text:span text:style-name="{generate-id()}">
                            <!--<xsl:value-of select ="a:t"/>-->
                            <!--converts whitespaces sequence to text:s-->
                            <!-- 1699083 bug fix  -->
                            <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
                            <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
                            <xsl:choose >
                              <xsl:when test ="a:rPr[@cap='all'] or $layoutCap ='all'">
                                <xsl:choose >
                                  <xsl:when test =".=''">
                                    <text:s/>
                                  </xsl:when>
                                  <xsl:when test ="not(contains(.,'  '))">
                                    <xsl:value-of select ="translate(.,$lcletters,$ucletters)"/>
                                  </xsl:when>
                                  <xsl:when test =". =' '">
                                    <text:s/>
                                  </xsl:when>
                                  <xsl:otherwise >
                                    <xsl:call-template name ="InsertWhiteSpaces">
                                      <xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
                                    </xsl:call-template>
                                  </xsl:otherwise>
                                </xsl:choose>
                              </xsl:when>
                              <xsl:when test ="a:rPr[@cap='small'] or $layoutCap ='small' ">
                                <xsl:choose >
                                  <xsl:when test =".=''">
                                    <text:s/>
                                  </xsl:when>
                                  <xsl:when test ="not(contains(.,'  '))">
                                    <xsl:value-of select ="translate(.,$ucletters,$lcletters)"/>
                                  </xsl:when>
                                  <xsl:when test =".= ' '">
                                    <text:s/>
                                  </xsl:when>
                                  <xsl:otherwise >
                                    <xsl:call-template name ="InsertWhiteSpaces">
                                      <xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
                                    </xsl:call-template>
                                  </xsl:otherwise>
                                </xsl:choose >
                              </xsl:when>
                              <xsl:otherwise >
                                <xsl:choose >
                                  <xsl:when test =".=''">
                                    <text:s/>
                                  </xsl:when>
                                  <xsl:when test ="not(contains(.,'  '))">
                                    <xsl:value-of select ="."/>
                                  </xsl:when>
                                  <xsl:otherwise >
                                    <xsl:call-template name ="InsertWhiteSpaces">
                                      <xsl:with-param name ="string" select ="."/>
                                    </xsl:call-template>
                                  </xsl:otherwise >
                                </xsl:choose>
                              </xsl:otherwise>
                            </xsl:choose>
                          </text:span>
                        </xsl:if >
                        <xsl:if test ="name()='a:br'">
                          <text:line-break/>
                        </xsl:if>
                        <!-- Added by lohith.ar for fix 1731885-->
                        <xsl:if test="name()='a:endParaRPr' and not(a:endParaRPr/a:hlinkClick)">
                          <text:span text:style-name="{generate-id()}" />
                        </xsl:if>
                      </xsl:for-each>
                    </text:p>
                  </xsl:if>
                </xsl:for-each>
              </draw:text-box >
              <!-- Added by lohith.ar - Start - Mouse click hyperlinks -->
              <office:event-listeners>
                <xsl:for-each select ="p:txBody/a:p">
                  <xsl:choose>
                    <!-- Start => Go to previous slide-->
                    <xsl:when test="a:endParaRPr/a:hlinkClick/@action[contains(.,'jump=previousslide')]">
                      <presentation:event-listener>
                        <xsl:attribute name ="script:event-name">
                          <xsl:value-of select ="'dom:click'"/>
                        </xsl:attribute>
                        <xsl:attribute name ="presentation:action">
                          <xsl:value-of select ="'previous-page'"/>
                        </xsl:attribute>
                      </presentation:event-listener>
                    </xsl:when>
                    <!-- End => Go to previous slide-->
                    <!-- Start => Go to Next slide -->
                    <xsl:when test="a:endParaRPr/a:hlinkClick/@action[contains(.,'jump=nextslide')]">
                      <presentation:event-listener>
                        <xsl:attribute name ="script:event-name">
                          <xsl:value-of select ="'dom:click'"/>
                        </xsl:attribute>
                        <xsl:attribute name ="presentation:action">
                          <xsl:value-of select ="'next-page'"/>
                        </xsl:attribute>
                      </presentation:event-listener>
                    </xsl:when>
                    <!-- End => Go to Next slide-->
                    <!-- Start => Go to First slide -->
                    <xsl:when test="a:endParaRPr/a:hlinkClick/@action[contains(.,'jump=firstslide')]">
                      <presentation:event-listener>
                        <xsl:attribute name ="script:event-name">
                          <xsl:value-of select ="'dom:click'"/>
                        </xsl:attribute>
                        <xsl:attribute name ="presentation:action">
                          <xsl:value-of select ="'first-page'"/>
                        </xsl:attribute>
                      </presentation:event-listener>
                    </xsl:when>
                    <!-- End => Go to First slide -->
                    <!-- Start => Go to Last slide -->
                    <xsl:when test="a:endParaRPr/a:hlinkClick/@action[contains(.,'jump=lastslide')]">
                      <presentation:event-listener>
                        <xsl:attribute name ="script:event-name">
                          <xsl:value-of select ="'dom:click'"/>
                        </xsl:attribute>
                        <xsl:attribute name ="presentation:action">
                          <xsl:value-of select ="'last-page'"/>
                        </xsl:attribute>
                      </presentation:event-listener>
                    </xsl:when>
                    <!-- End => Go to Last slide -->
                    <!-- Start => EndShow -->
                    <xsl:when test="a:endParaRPr/a:hlinkClick/@action[contains(.,'jump=endshow')]">
                      <presentation:event-listener>
                        <xsl:attribute name ="script:event-name">
                          <xsl:value-of select ="'dom:click'"/>
                        </xsl:attribute>
                        <xsl:attribute name ="presentation:action">
                          <xsl:value-of select ="'stop'"/>
                        </xsl:attribute>
                      </presentation:event-listener>
                    </xsl:when>
                    <!-- End => End show -->
                    <!-- Start => Go to Page or Object && Go to Other document && Run program  -->
                    <xsl:when test="a:endParaRPr/a:hlinkClick/@action[contains(.,'ppaction://hlinksldjump') or contains(.,'ppaction://hlinkfile') or contains(.,'ppaction://program') ]">
                      <presentation:event-listener>
                        <xsl:attribute name ="script:event-name">
                          <xsl:value-of select ="'dom:click'"/>
                        </xsl:attribute>
                        <xsl:variable name="RelationId">
                          <xsl:value-of select="a:endParaRPr/a:hlinkClick/@r:id"/>
                        </xsl:variable>
                        <xsl:variable name="SlideVal">
                          <xsl:value-of select="document($slideRel)/rels:Relationships/rels:Relationship[@Id=$RelationId]/@Target"/>
                        </xsl:variable>
                        <!-- Condn Go to Other page/slide-->
                        <xsl:if test="a:endParaRPr/a:hlinkClick/@action[contains(.,'ppaction://hlinksldjump')]">
                          <xsl:attribute name ="presentation:action">
                            <xsl:value-of select ="'show'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="xlink:href">
                            <xsl:value-of select ="concat('#Page',substring-before(substring-after($SlideVal,'slide'),'.xml'))"/>
                          </xsl:attribute>
                        </xsl:if>
                        <!-- Condn Go to Other document-->
                        <xsl:if test="a:endParaRPr/a:hlinkClick/@action[contains(.,'ppaction://hlinkfile')]">
                          <xsl:attribute name ="presentation:action">
                            <xsl:value-of select ="'show'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="xlink:href">
                            <xsl:if test="string-length(substring-after($SlideVal,'file:///')) > 0">
                              <xsl:value-of select ="concat('/',translate(substring-after($SlideVal,'file:///'),'\','/'))"/>
                            </xsl:if>
                            <xsl:if test="string-length(substring-after($SlideVal,'file:///')) = 0 ">
                              <xsl:value-of select ="concat('../',$SlideVal)"/>
                            </xsl:if>
                          </xsl:attribute>
                        </xsl:if>
                        <!-- Condn Go to Run program-->
                        <xsl:if test="a:endParaRPr/a:hlinkClick/@action[contains(.,'ppaction://program') ]">
                          <xsl:attribute name ="presentation:action">
                            <xsl:value-of select ="'execute'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="xlink:href">
                            <xsl:if test="string-length(substring-after($SlideVal,'file:///')) > 0">
                              <xsl:value-of select ="concat('/',translate(substring-after($SlideVal,'file:///'),'\','/'))"/>
                            </xsl:if>
                            <xsl:if test="string-length(substring-after($SlideVal,'file:///')) = 0">
                              <xsl:value-of select ="concat('../',$SlideVal)"/>
                            </xsl:if>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:attribute name ="xlink:type">
                          <xsl:value-of select ="'simple'"/>
                        </xsl:attribute>
                        <xsl:attribute name ="xlink:show">
                          <xsl:value-of select ="'new'"/>
                        </xsl:attribute>
                        <xsl:attribute name ="xlink:actuate">
                          <xsl:value-of select ="'onRequest'"/>
                        </xsl:attribute>
                      </presentation:event-listener>
                    </xsl:when>
                    <!-- End => Go to Page or Object && Go to Other document && Run program -->
                    <!-- Start => Paly sound  -->
                    <xsl:when test="a:endParaRPr/a:hlinkClick/a:snd">
                      <presentation:event-listener>
                        <xsl:attribute name ="script:event-name">
                          <xsl:value-of select ="'dom:click'"/>
                        </xsl:attribute>
                        <xsl:attribute name ="presentation:action">
                          <xsl:value-of select ="'sound'"/>
                        </xsl:attribute>
                        <presentation:sound>
                          <xsl:variable name="varMediaFileRelId">
                            <xsl:value-of select="a:endParaRPr/a:hlinkClick/a:snd/@r:embed"/>
                          </xsl:variable>
                          <xsl:variable name="varMediaFileTargetPath">
                            <xsl:value-of select="document($slideRel)/rels:Relationships/rels:Relationship[@Id=$varMediaFileRelId]/@Target"/>
                          </xsl:variable>
                          <xsl:variable name="varPptMediaFileTargetPath">
                            <xsl:value-of select="concat('ppt/',substring-after($varMediaFileTargetPath,'/'))"/>
                          </xsl:variable>
                          <xsl:variable name="varDocumentModifiedTime">
                            <xsl:value-of select="document('docProps/core.xml')/cp:coreProperties/dcterms:modified"/>
                          </xsl:variable>
                          <xsl:variable name="varDestMediaFileTargetPath">
                            <xsl:value-of select="concat(translate($varDocumentModifiedTime,':','-'),'|',$varMediaFileRelId,'.wav')"/>
                          </xsl:variable>
                          <xsl:variable name="varMediaFilePathForOdp">
                            <xsl:value-of select="concat('../_MediaFilesForOdp_',translate($varDocumentModifiedTime,':','-'),'/',$varMediaFileRelId,'.wav')"/>
                          </xsl:variable>
                          <xsl:attribute name ="xlink:href">
                            <xsl:value-of select ="$varMediaFilePathForOdp"/>
                          </xsl:attribute>
                          <xsl:attribute name ="xlink:type">
                            <xsl:value-of select ="'simple'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="xlink:show">
                            <xsl:value-of select ="'new'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="xlink:actuate">
                            <xsl:value-of select ="'onRequest'"/>
                          </xsl:attribute>
                          <pzip:extract pzip:source="{$varPptMediaFileTargetPath}" pzip:target="{$varDestMediaFileTargetPath}" />
                        </presentation:sound>
                      </presentation:event-listener>
                    </xsl:when>
                    <!-- End => Paly sound  -->
                  </xsl:choose>
                </xsl:for-each>
              </office:event-listeners>
              <!-- End - Mouse click hyperlinks-->
            </draw:frame >
          </xsl:when>
          <!-- If Slide layout files have the frame properties-->
          <xsl:when test ="not(p:nvSpPr/p:nvPr/p:ph/@type[contains(.,$LayoutName)]
								and p:spPr/a:xfrm/a:off)" >
            <xsl:variable name ="frameName">
              <xsl:if test ="p:nvSpPr/p:nvPr/p:ph/@type">
                <xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph/@type">
                  <xsl:value-of select ="."/>
                </xsl:for-each>
              </xsl:if>
              <xsl:if test="not(p:nvSpPr/p:nvPr/p:ph/@type)">
                <xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph/@idx">
                  <xsl:value-of select ="."/>
                </xsl:for-each>
              </xsl:if >
            </xsl:variable>
            <xsl:variable name ="FrameIdx">
              <xsl:if test ="p:nvSpPr/p:nvPr/p:ph/@idx">
                <xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph[@idx]">
                  <xsl:value-of select ="@idx"/>
                </xsl:for-each>
              </xsl:if>
            </xsl:variable>
            <xsl:variable name ="TextSpanNode">
              <xsl:for-each select ="p:txBody/a:p">
                <!--Code Inserted by Vijayeta for Bullets And Numbering
                  check for levels and then depending on the condition,insert bullets,If properties are defined in corresponding layouts-->
                <xsl:if test ="$LayoutName != 'title'">
                  <xsl:if test ="a:pPr/a:buChar or a:pPr/a:buAutoNum or a:pPr/a:buBlip">
                    <xsl:call-template name ="insertBulletsNumbersoox2odf">
                      <xsl:with-param name ="listStyleName" select ="concat($listStyleName,position())"/>
                      <xsl:with-param name ="ParaId" select ="$ParaId"/>
                    </xsl:call-template>
                  </xsl:if>
                  <!-- If no bullets are present or default bullets-->
                  <xsl:if test ="not(a:pPr/a:buChar) and not(a:pPr/a:buAutoNum)and not(a:pPr/a:buBlip)">
                    <!-- If bullets are default-->
                    <xsl:if test ="$bulletTypeBool='true'">
                      <xsl:call-template name ="insertBulletsNumbersoox2odf">
                        <xsl:with-param name ="listStyleName" select ="concat($listStyleName,position())"/>
                        <xsl:with-param name ="ParaId" select ="$ParaId"/>
                      </xsl:call-template>
                    </xsl:if>
                  </xsl:if>
                  <!-- If no bullets present at all-->
                  <xsl:if test ="not($bulletTypeBool='true')and not(a:pPr/a:buChar) and not(a:pPr/a:buAutoNum)and not(a:pPr/a:buBlip)">
                    <text:p >
                      <xsl:attribute name ="text:style-name">
                        <xsl:value-of select ="concat($ParaId,position())"/>
                      </xsl:attribute>
                      <xsl:for-each select ="node()">
                        <xsl:if test ="name()='a:r'">
                          <text:span text:style-name="{generate-id()}">
                            <!--<xsl:value-of select ="a:t"/>-->
                            <!--converts whitespaces sequence to text:s-->
                            <!-- 1699083 bug fix  -->
                            <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
                            <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
                            <xsl:choose >
                              <xsl:when test ="a:rPr[@cap='all'] or $layoutCap ='all' ">
                                <xsl:choose >
                                  <xsl:when test =".=''">
                                    <text:s/>
                                  </xsl:when>
                                  <xsl:when test ="not(contains(.,'  '))">
                                    <xsl:value-of select ="translate(.,$lcletters,$ucletters)"/>
                                  </xsl:when>
                                  <xsl:when test =". =' '">
                                    <text:s/>
                                  </xsl:when>
                                  <xsl:otherwise >
                                    <xsl:call-template name ="InsertWhiteSpaces">
                                      <xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
                                    </xsl:call-template>
                                  </xsl:otherwise>
                                </xsl:choose>
                              </xsl:when>
                              <xsl:when test ="a:rPr[@cap='small'] or $layoutCap ='small'">
                                <xsl:choose >
                                  <xsl:when test =".=''">
                                    <text:s/>
                                  </xsl:when>
                                  <xsl:when test ="not(contains(.,'  '))">
                                    <xsl:value-of select ="translate(.,$ucletters,$lcletters)"/>
                                  </xsl:when>
                                  <xsl:when test =".= ' '">
                                    <text:s/>
                                  </xsl:when>
                                  <xsl:otherwise >
                                    <xsl:call-template name ="InsertWhiteSpaces">
                                      <xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
                                    </xsl:call-template>
                                  </xsl:otherwise>
                                </xsl:choose >
                              </xsl:when>
                              <xsl:otherwise >
                                <xsl:choose >
                                  <xsl:when test =".=''">
                                    <text:s/>
                                  </xsl:when>
                                  <xsl:when test ="not(contains(.,'  '))">
                                    <xsl:value-of select ="."/>
                                  </xsl:when>
                                  <xsl:otherwise >
                                    <xsl:call-template name ="InsertWhiteSpaces">
                                      <xsl:with-param name ="string" select ="."/>
                                    </xsl:call-template>
                                  </xsl:otherwise >
                                </xsl:choose>
                              </xsl:otherwise>
                            </xsl:choose>
                          </text:span>
                        </xsl:if >
                        <xsl:if test ="name()='a:br'">
                          <text:line-break/>
                        </xsl:if>
                        <!-- Added by lohith.ar for fix 1731885-->
                        <xsl:if test="name()='a:endParaRPr' and not(a:endParaRPr/a:hlinkClick)">
                          <text:span text:style-name="{generate-id()}" />
                        </xsl:if>
                      </xsl:for-each>
                    </text:p>
                  </xsl:if>
                </xsl:if>
                <xsl:if test ="$LayoutName = 'title'">
                  <text:p >
                    <xsl:attribute name ="text:style-name">
                      <xsl:value-of select ="concat($ParaId,position())"/>
                    </xsl:attribute>
                    <xsl:for-each select ="node()">
                      <xsl:if test ="name()='a:r'">
                        <text:span text:style-name="{generate-id()}">
                          <!--<xsl:value-of select ="a:t"/>-->
                          <!--converts whitespaces sequence to text:s-->
                          <!-- 1699083 bug fix  -->
                          <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
                          <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
                          <xsl:choose >
                            <xsl:when test ="a:rPr[@cap='all'] or $layoutCap ='all'">
                              <xsl:choose >
                                <xsl:when test =".=''">
                                  <text:s/>
                                </xsl:when>
                                <xsl:when test ="not(contains(.,'  '))">
                                  <xsl:value-of select ="translate(.,$lcletters,$ucletters)"/>
                                </xsl:when>
                                <xsl:when test =". =' '">
                                  <text:s/>
                                </xsl:when>
                                <xsl:otherwise >
                                  <xsl:call-template name ="InsertWhiteSpaces">
                                    <xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
                                  </xsl:call-template>
                                </xsl:otherwise>
                              </xsl:choose>
                            </xsl:when>
                            <xsl:when test ="a:rPr[@cap='small'] or $layoutCap ='small'">
                              <xsl:choose >
                                <xsl:when test =".=''">
                                  <text:s/>
                                </xsl:when>
                                <xsl:when test ="not(contains(.,'  '))">
                                  <xsl:value-of select ="translate(.,$ucletters,$lcletters)"/>
                                </xsl:when>
                                <xsl:when test =".= ' '">
                                  <text:s/>
                                </xsl:when>
                                <xsl:otherwise >
                                  <xsl:call-template name ="InsertWhiteSpaces">
                                    <xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
                                  </xsl:call-template>
                                </xsl:otherwise>
                              </xsl:choose >
                            </xsl:when>
                            <xsl:otherwise >
                              <xsl:choose >
                                <xsl:when test =".=''">
                                  <text:s/>
                                </xsl:when>
                                <xsl:when test ="not(contains(.,'  '))">
                                  <xsl:value-of select ="."/>
                                </xsl:when>
                                <xsl:otherwise >
                                  <xsl:call-template name ="InsertWhiteSpaces">
                                    <xsl:with-param name ="string" select ="."/>
                                  </xsl:call-template>
                                </xsl:otherwise >
                              </xsl:choose>
                            </xsl:otherwise>
                          </xsl:choose>
                        </text:span>
                      </xsl:if >
                      <xsl:if test ="name()='a:br'">
                        <text:line-break/>
                      </xsl:if>
                      <!-- Added by lohith.ar for fix 1731885-->
                      <xsl:if test="name()='a:endParaRPr' and not(a:endParaRPr/a:hlinkClick)">
                        <text:span text:style-name="{generate-id()}" />
                      </xsl:if>
                    </xsl:for-each>
                  </text:p>
                </xsl:if>
              </xsl:for-each>
            </xsl:variable>
            <!-- Added by lohith.ar - Start - Variable for Mouse click hyperlinks -->
            <xsl:variable name="EventListnerNode">
              <!-- Link Action-->
              <office:event-listeners>
                <xsl:for-each select ="p:txBody/a:p">
                  <xsl:choose>
                    <!-- Start => Go to previous slide-->
                    <xsl:when test="a:endParaRPr/a:hlinkClick/@action[contains(.,'jump=previousslide')]">
                      <presentation:event-listener>
                        <xsl:attribute name ="script:event-name">
                          <xsl:value-of select ="'dom:click'"/>
                        </xsl:attribute>
                        <xsl:attribute name ="presentation:action">
                          <xsl:value-of select ="'previous-page'"/>
                        </xsl:attribute>
                      </presentation:event-listener>
                    </xsl:when>
                    <!-- End => Go to previous slide-->
                    <!-- Start => Go to Next slide -->
                    <xsl:when test="a:endParaRPr/a:hlinkClick/@action[contains(.,'jump=nextslide')]">
                      <presentation:event-listener>
                        <xsl:attribute name ="script:event-name">
                          <xsl:value-of select ="'dom:click'"/>
                        </xsl:attribute>
                        <xsl:attribute name ="presentation:action">
                          <xsl:value-of select ="'next-page'"/>
                        </xsl:attribute>
                      </presentation:event-listener>
                    </xsl:when>
                    <!-- End => Go to Next slide-->
                    <!-- Start => Go to First slide -->
                    <xsl:when test="a:endParaRPr/a:hlinkClick/@action[contains(.,'jump=firstslide')]">
                      <presentation:event-listener>
                        <xsl:attribute name ="script:event-name">
                          <xsl:value-of select ="'dom:click'"/>
                        </xsl:attribute>
                        <xsl:attribute name ="presentation:action">
                          <xsl:value-of select ="'first-page'"/>
                        </xsl:attribute>
                      </presentation:event-listener>
                    </xsl:when>
                    <!-- End => Go to First slide -->
                    <!-- Start => Go to Last slide -->
                    <xsl:when test="a:endParaRPr/a:hlinkClick/@action[contains(.,'jump=lastslide')]">
                      <presentation:event-listener>
                        <xsl:attribute name ="script:event-name">
                          <xsl:value-of select ="'dom:click'"/>
                        </xsl:attribute>
                        <xsl:attribute name ="presentation:action">
                          <xsl:value-of select ="'last-page'"/>
                        </xsl:attribute>
                      </presentation:event-listener>
                    </xsl:when>
                    <!-- End => Go to Last slide -->
                    <!-- Start => EndShow -->
                    <xsl:when test="a:endParaRPr/a:hlinkClick/@action[contains(.,'jump=endshow')]">
                      <presentation:event-listener>
                        <xsl:attribute name ="script:event-name">
                          <xsl:value-of select ="'dom:click'"/>
                        </xsl:attribute>
                        <xsl:attribute name ="presentation:action">
                          <xsl:value-of select ="'stop'"/>
                        </xsl:attribute>
                      </presentation:event-listener>
                    </xsl:when>
                    <!-- End => End show -->
                    <!-- Start => Go to Page or Object && Go to Other document && Run program-->
                    <xsl:when test="a:endParaRPr/a:hlinkClick/@action[contains(.,'ppaction://hlinksldjump') or contains(.,'ppaction://hlinkfile') or contains(.,'ppaction://program') ]">
                      <presentation:event-listener>
                        <xsl:attribute name ="script:event-name">
                          <xsl:value-of select ="'dom:click'"/>
                        </xsl:attribute>
                        <xsl:variable name="RelationId">
                          <xsl:value-of select="a:endParaRPr/a:hlinkClick/@r:id"/>
                        </xsl:variable>
                        <xsl:variable name="SlideVal">
                          <xsl:value-of select="document($slideRel)/rels:Relationships/rels:Relationship[@Id=$RelationId]/@Target"/>
                        </xsl:variable>
                        <!-- Condn Go to Other page/slide-->
                        <xsl:if test="a:endParaRPr/a:hlinkClick/@action[contains(.,'ppaction://hlinksldjump')]">
                          <xsl:attribute name ="presentation:action">
                            <xsl:value-of select ="'show'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="xlink:href">
                            <xsl:value-of select ="concat('#Page',substring-before(substring-after($SlideVal,'slide'),'.xml'))"/>
                          </xsl:attribute>
                        </xsl:if>
                        <!-- Condn Go to Other document-->
                        <xsl:if test="a:endParaRPr/a:hlinkClick/@action[contains(.,'ppaction://hlinkfile')]">
                          <xsl:attribute name ="presentation:action">
                            <xsl:value-of select ="'show'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="xlink:href">
                            <xsl:if test="string-length(substring-after($SlideVal,'file:///')) > 0">
                              <xsl:value-of select ="concat('/',translate(substring-after($SlideVal,'file:///'),'\','/'))"/>
                            </xsl:if>
                            <xsl:if test="string-length(substring-after($SlideVal,'file:///')) = 0">
                              <xsl:value-of select ="concat('../',$SlideVal)"/>
                            </xsl:if>
                          </xsl:attribute>
                        </xsl:if>
                        <!-- Condn Go to Run program-->
                        <xsl:if test="a:endParaRPr/a:hlinkClick/@action[contains(.,'ppaction://program') ]">
                          <xsl:attribute name ="presentation:action">
                            <xsl:value-of select ="'execute'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="xlink:href">
                            <xsl:if test="string-length(substring-after($SlideVal,'file:///')) > 0">
                              <xsl:value-of select ="concat('/',translate(substring-after($SlideVal,'file:///'),'\','/'))"/>
                            </xsl:if>
                            <xsl:if test="string-length(substring-after($SlideVal,'file:///')) = 0">
                              <xsl:value-of select ="concat('../',$SlideVal)"/>
                            </xsl:if>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:attribute name ="xlink:type">
                          <xsl:value-of select ="'simple'"/>
                        </xsl:attribute>
                        <xsl:attribute name ="xlink:show">
                          <xsl:value-of select ="'new'"/>
                        </xsl:attribute>
                        <xsl:attribute name ="xlink:actuate">
                          <xsl:value-of select ="'onRequest'"/>
                        </xsl:attribute>
                      </presentation:event-listener>
                    </xsl:when>
                    <!-- End => Go to Page or Object && Go to Other document && Run program  -->
                    <!-- Start => Paly sound  -->
                    <xsl:when test="a:endParaRPr/a:hlinkClick/a:snd">
                      <presentation:event-listener>
                        <xsl:attribute name ="script:event-name">
                          <xsl:value-of select ="'dom:click'"/>
                        </xsl:attribute>
                        <xsl:attribute name ="presentation:action">
                          <xsl:value-of select ="'sound'"/>
                        </xsl:attribute>
                        <presentation:sound>
                          <xsl:variable name="varMediaFileRelId">
                            <xsl:value-of select="a:endParaRPr/a:hlinkClick/a:snd/@r:embed"/>
                          </xsl:variable>
                          <xsl:variable name="varMediaFileTargetPath">
                            <xsl:value-of select="document($slideRel)/rels:Relationships/rels:Relationship[@Id=$varMediaFileRelId]/@Target"/>
                          </xsl:variable>
                          <xsl:variable name="varPptMediaFileTargetPath">
                            <xsl:value-of select="concat('ppt/',substring-after($varMediaFileTargetPath,'/'))"/>
                          </xsl:variable>
                          <xsl:variable name="varDocumentModifiedTime">
                            <xsl:value-of select="document('docProps/core.xml')/cp:coreProperties/dcterms:modified"/>
                          </xsl:variable>
                          <xsl:variable name="varDestMediaFileTargetPath">
                            <xsl:value-of select="concat(translate($varDocumentModifiedTime,':','-'),'|',$varMediaFileRelId,'.wav')"/>
                          </xsl:variable>
                          <xsl:variable name="varMediaFilePathForOdp">
                            <xsl:value-of select="concat('../_MediaFilesForOdp_',translate($varDocumentModifiedTime,':','-'),'/',$varMediaFileRelId,'.wav')"/>
                          </xsl:variable>
                          <xsl:attribute name ="xlink:href">
                            <xsl:value-of select ="$varMediaFilePathForOdp"/>
                          </xsl:attribute>
                          <xsl:attribute name ="xlink:type">
                            <xsl:value-of select ="'simple'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="xlink:show">
                            <xsl:value-of select ="'new'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="xlink:actuate">
                            <xsl:value-of select ="'onRequest'"/>
                          </xsl:attribute>
                          <pzip:extract pzip:source="{$varPptMediaFileTargetPath}" pzip:target="{$varDestMediaFileTargetPath}" />
                        </presentation:sound>
                      </presentation:event-listener>
                    </xsl:when>
                    <!-- End => Paly sound  -->
                  </xsl:choose>
                </xsl:for-each>
              </office:event-listeners>
              <!--<pzip:copy pzip:source="ppt/media/audio1.wav" pzip:target="media/audio1.wav"/>-->
            </xsl:variable>
            <!-- End - Variable for Mouse click hyperlinks-->
            <xsl:variable name ="LayoutFileNo">
              <xsl:for-each select ="document($slideRel)//node()/@Target[contains(.,'slideLayouts')]">
                <xsl:value-of select ="concat('ppt',substring(.,3))"/>
              </xsl:for-each>
            </xsl:variable>
            <!-- Slide Layout Files Loop Second Loop-->
            <xsl:for-each select ="document($LayoutFileNo)/p:sldLayout/p:cSld/p:spTree/p:sp">
              <xsl:variable name ="SlFrameName">
                <xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph/@type">
                  <xsl:value-of select ="."/>
                </xsl:for-each>
                <xsl:if test="not(p:nvSpPr/p:nvPr/p:ph/@type)">
                  <xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph/@idx">
                    <xsl:value-of select ="."/>
                  </xsl:for-each>
                </xsl:if >
              </xsl:variable>
              <xsl:variable name ="SlFrameNameInd">
                <xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph/@idx">
                  <xsl:value-of select ="."/>
                </xsl:for-each>
              </xsl:variable>
              <xsl:choose >
                <!-- Modified by lohith to fix the bug 1719280-->
                <xsl:when test ="not($SlFrameName = $frameName 
														or $SlFrameNameInd = $FrameIdx) ">
                  <!-- Do nothing-or $SlFrameNameInd=$FrameIdx -->
                </xsl:when>
                <!-- Commented by lohith to fix the bug 1719280-->
                <!--<xsl:when test ="not($SlFrameName = $frameName) 
														and string-length($SlFrameName) &gt; 0 
														and string-length($frameName) &gt; 0 ">
                  -->
                <!-- Do nothing-or $SlFrameNameInd=$FrameIdx -->
                <!--
                </xsl:when>-->
                <xsl:when test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'dt') 
										or contains(.,'ftr') or contains(.,'sldNum')]">
                  <!-- Do nothing-->
                  <!-- These will be covered in footer and date time -->
                </xsl:when>

                <xsl:when test ="(p:nvSpPr/p:nvPr/p:ph/@type[contains(.,$frameName)] or 
												p:nvSpPr/p:nvPr/p:ph/@idx[contains(.,$FrameIdx)])
												and p:spPr/a:xfrm/a:off" >
                  <xsl:choose >
                    <xsl:when test ="$SlFrameNameInd != $FrameIdx and
												  string-length($SlFrameName) &gt; 0 
												  and string-length($frameName) &gt; 0">
                      <!--do nothing -->
                    </xsl:when >
                    <xsl:otherwise >
                      <draw:frame draw:layer="layout" 																			
																  presentation:user-transformed="true">
                        <xsl:attribute name ="presentation:style-name">
                          <xsl:value-of select ="$textLayoutId"/>
                        </xsl:attribute>
                        <xsl:attribute name ="presentation:class">
                          <xsl:call-template name ="LayoutType">
                            <xsl:with-param name ="LayoutStyle">
                              <xsl:value-of select ="$LayoutName"/>
                            </xsl:with-param>
                          </xsl:call-template >
                        </xsl:attribute>
                        <xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cx]">
                          <xsl:attribute name ="svg:width">
                            <xsl:call-template name="ConvertEmu">
                              <xsl:with-param name="length" select="@cx"/>
                              <xsl:with-param name="unit">cm</xsl:with-param>
                            </xsl:call-template>
                          </xsl:attribute>
                        </xsl:for-each>
                        <xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cy]">
                          <xsl:attribute name ="svg:height">
                            <xsl:call-template name="ConvertEmu">
                              <xsl:with-param name="length" select="@cy"/>
                              <xsl:with-param name="unit">cm</xsl:with-param>
                            </xsl:call-template>
                          </xsl:attribute>
                        </xsl:for-each>
                        <xsl:for-each select ="p:spPr/a:xfrm/a:off[@x]">
                          <xsl:attribute name ="svg:x">
                            <xsl:call-template name="ConvertEmu">
                              <xsl:with-param name="length" select="@x"/>
                              <xsl:with-param name="unit">cm</xsl:with-param>
                            </xsl:call-template>
                          </xsl:attribute>
                        </xsl:for-each>
                        <xsl:for-each select ="p:spPr/a:xfrm/a:off[@y]">
                          <xsl:attribute name ="svg:y">
                            <xsl:call-template name="ConvertEmu">
                              <xsl:with-param name="length" select="@y"/>
                              <xsl:with-param name="unit">cm</xsl:with-param>
                            </xsl:call-template>
                          </xsl:attribute>
                        </xsl:for-each>
                        <draw:text-box>
                          <xsl:copy-of select ="$TextSpanNode"/>
                        </draw:text-box >
                        <!-- Added by lohith.ar -->
                        <xsl:copy-of select ="$EventListnerNode"/>
                      </draw:frame >

                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when >
                <xsl:when test ="not((p:nvSpPr/p:nvPr/p:ph/@type[contains(.,$frameName)] or 
											p:nvSpPr/p:nvPr/p:ph/@idx[contains(.,$FrameIdx)])
											and p:spPr/a:xfrm/a:off)">
                  <xsl:variable name ="LayoutRels">
                    <xsl:for-each select ="document(concat(concat('ppt/slideLayouts/_rels',substring($LayoutFileNo,17)),'.rels'))
													//node()/@Target[contains(.,'slideMasters')]">
                      <xsl:value-of select ="."/>
                    </xsl:for-each>
                  </xsl:variable>
                  <xsl:variable name ="MasterFileName">
                    <xsl:value-of select ="concat('ppt/slideMasters',substring($LayoutRels,16))"/>
                  </xsl:variable>
                  <!-- Slide Master Files Loop  Third Loop-->
                  <xsl:for-each select ="document($MasterFileName)/p:sldMaster/p:cSld/p:spTree/p:sp">
                    <xsl:variable name ="MstrFrameName">
                      <xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph/@type">
                        <xsl:value-of select ="."/>
                      </xsl:for-each>
                      <xsl:if test="not(p:nvSpPr/p:nvPr/p:ph/@type)">
                        <xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph/@idx">
                          <xsl:value-of select ="."/>
                        </xsl:for-each>
                      </xsl:if >
                    </xsl:variable>
                    <xsl:variable name ="MstrFrameInd">
                      <xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph/@idx">
                        <xsl:value-of select ="."/>
                      </xsl:for-each>
                    </xsl:variable>
                    <xsl:if test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,$frameName)]">
                      <xsl:value-of select ="p:nvSpPr/p:nvPr/p:ph"/>
                    </xsl:if>
                    <xsl:choose >
                      <xsl:when test ="not($MstrFrameName = $frameName
												or $MstrFrameInd =$FrameIdx)">
                        <!-- Do nothing-->
                        <!-- These will be covered in footer and date time -->
                      </xsl:when>
                      <xsl:when test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'dt') or 
																 contains(.,'ftr') or contains(.,'sldNum')]">
                        <!-- Do nothing-->
                        <!-- These will be covered in footer and date time -->
                      </xsl:when>
                      <xsl:when test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,$frameName)] 
													or p:nvSpPr/p:nvPr/p:ph/@idx[contains(.,$FrameIdx)]">
                        <draw:frame draw:layer="layout"
																		presentation:user-transformed="true">
                          <xsl:attribute name ="presentation:style-name">
                            <xsl:value-of select ="$textLayoutId"/>
                          </xsl:attribute>
                          <xsl:attribute name ="presentation:class">
                            <xsl:call-template name ="LayoutType">
                              <xsl:with-param name ="LayoutStyle">
                                <xsl:value-of select ="$LayoutName"/>
                              </xsl:with-param>
                            </xsl:call-template >
                          </xsl:attribute>
                          <xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cx]">
                            <xsl:attribute name ="svg:width">
                              <xsl:call-template name="ConvertEmu">
                                <xsl:with-param name="length" select="@cx"/>
                                <xsl:with-param name="unit">cm</xsl:with-param>
                              </xsl:call-template>
                            </xsl:attribute>
                          </xsl:for-each>
                          <xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cy]">
                            <xsl:attribute name ="svg:height">
                              <xsl:call-template name="ConvertEmu">
                                <xsl:with-param name="length" select="@cy"/>
                                <xsl:with-param name="unit">cm</xsl:with-param>
                              </xsl:call-template>
                            </xsl:attribute>
                          </xsl:for-each>
                          <xsl:for-each select ="p:spPr/a:xfrm/a:off[@x]">
                            <xsl:attribute name ="svg:x">
                              <xsl:call-template name="ConvertEmu">
                                <xsl:with-param name="length" select="@x"/>
                                <xsl:with-param name="unit">cm</xsl:with-param>
                              </xsl:call-template>
                            </xsl:attribute>
                          </xsl:for-each>
                          <xsl:for-each select ="p:spPr/a:xfrm/a:off[@y]">
                            <xsl:attribute name ="svg:y">
                              <xsl:call-template name="ConvertEmu">
                                <xsl:with-param name="length" select="@y"/>
                                <xsl:with-param name="unit">cm</xsl:with-param>
                              </xsl:call-template>
                            </xsl:attribute>
                          </xsl:for-each>
                          <draw:text-box>
                            <xsl:copy-of select ="$TextSpanNode"/>
                          </draw:text-box >
                          <!-- Added by lohith.ar -->
                          <xsl:copy-of select ="$EventListnerNode"/>
                        </draw:frame >
                      </xsl:when>
                    </xsl:choose>
                  </xsl:for-each >
                  <!-- Exit Slide Master Files Loop  Third Loop-->
                </xsl:when>
              </xsl:choose>
            </xsl:for-each>
            <!-- exit Slide layout loop second Loop-->
          </xsl:when>
        </xsl:choose >

      </xsl:for-each>
      <!-- exit Slide loop Main Loop-->

      <!--<xsl:for-each select ="document(concat('ppt/',$slideNo))/p:sld/p:cSld/p:spTree/p:cxnSp">
				-->
      <!-- Code for shapes(Line) start-->
      <!--
				<xsl:variable  name ="GraphicId">
					<xsl:value-of select ="concat('s',substring($SlideFile,6,string-length($SlideFile)-9) ,concat('grLine',position()))"/>
				</xsl:variable>
				<xsl:call-template name ="shapes">
					<xsl:with-param name="GraphicId" select ="$GraphicId"/>
					<xsl:with-param name ="SlideRelationId" select="$slideRel" />
				</xsl:call-template>

				-->
      <!-- Code for shapes(Line) end-->
      <!--
			</xsl:for-each>-->
      <xsl:for-each select ="document(concat('ppt/',$slideNo))/p:sld/p:cSld/p:spTree">
        <xsl:variable name ="SlideID">
          <xsl:value-of select ="concat('slide',substring($SlideFile,6,string-length($SlideFile)-9))"/>
        </xsl:variable>
        <xsl:call-template name ="processShapes">
          <xsl:with-param name ="SlideID"  select ="$SlideID" />
          <xsl:with-param name ="SlideRelationId" select="$slideRel" />
          <xsl:with-param name ="grID" select ="'gr'" />
          <xsl:with-param name ="prID" select ="'PARA'" />
        </xsl:call-template>
      </xsl:for-each>
      <!-- Code for shapes end-->
    </xsl:for-each>
  </xsl:template>
  <!-- Gets all caps attribute from Layout -->
  <xsl:template name ="getAllCapsFromLayout">
    <xsl:param name ="phType" select="p:nvSpPr/p:nvPr/p:ph/@type"/>
    <xsl:param name ="phIdx" select ="p:nvSpPr/p:nvPr/p:ph/@idx"/>
    <xsl:param name ="layOutRelId"/>
    <xsl:variable name ="LayoutFileNo">
      <xsl:for-each select ="document($layOutRelId)//node()/@Target[contains(.,'slideLayouts')]">
        <xsl:value-of select ="concat('ppt',substring(.,3))"/>
      </xsl:for-each>
    </xsl:variable>
    <xsl:for-each select ="document($LayoutFileNo)/p:sldLayout/p:cSld/p:spTree/p:sp">
      <xsl:if test ="p:nvSpPr/p:nvPr/p:ph/@type = $phType">
        <xsl:value-of select ="p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@cap"/>
      </xsl:if>
      <xsl:if test ="p:nvSpPr/p:nvPr/p:ph/@idx = $phIdx">
        <xsl:value-of select ="p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@cap"/>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name ="processShapes">
    <xsl:param name ="SlideID" />
    <xsl:param name ="SlideRelationId" />
    <xsl:param name ="grID" />
    <xsl:param name ="prID" />
    <xsl:for-each select ="p:sp">
      <xsl:call-template name ="drawShapes">
        <xsl:with-param name ="SlideID" select ="$SlideID" />
        <xsl:with-param name ="SlideRelationId" select ="$SlideRelationId" />
        <xsl:with-param name ="grID" select ="$grID" />
        <xsl:with-param name ="prID" select ="$prID" />
      </xsl:call-template>
    </xsl:for-each>

    <xsl:for-each select ="p:cxnSp">
      <xsl:call-template name ="drawShapes">
        <xsl:with-param name ="SlideID" select ="$SlideID" />
        <xsl:with-param name ="SlideRelationId" select ="$SlideRelationId" />
        <xsl:with-param name ="grID" select ="concat($grID,'Line')" />
      </xsl:call-template>
    </xsl:for-each>

    <xsl:for-each select ="p:grpSp">
      <draw:g>
        <xsl:call-template name ="processShapes">
          <xsl:with-param name ="SlideID"  select ="$SlideID" />
          <xsl:with-param name ="SlideRelationId" select="$SlideRelationId" />
          <xsl:with-param name ="grID" select ="concat('grp', generate-id())" />
          <xsl:with-param name ="prID" select ="concat('prp', generate-id())" />
        </xsl:call-template>
      </draw:g>
    </xsl:for-each >
  </xsl:template>
  <xsl:template name ="FrameProperties" match ="p:spPr/a:xfrm">
    <xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cx]">
      <xsl:attribute name ="svg:width">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="@cx"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:for-each>
    <xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cy]">
      <xsl:attribute name ="svg:height">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="@cy"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:for-each>
    <xsl:for-each select ="a:off[@x]">
      <xsl:attribute name ="svg:x">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="@x"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:for-each>
    <xsl:for-each select ="a:off[@y]">
      <xsl:attribute name ="svg:y">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="@y"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="WriteTextSpan">
    <xsl:param name ="file"/>
    <draw:text-box>
      <xsl:for-each select ="document(concat('ppt/',$file))/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p">
        <text:p text:style-name="P1">
          <xsl:for-each select ="a:r">
            <text:span text:style-name="T1">
              <xsl:value-of select ="a:t"/>
            </text:span>
          </xsl:for-each>
        </text:p>
      </xsl:for-each>
    </draw:text-box >
  </xsl:template>
  <xsl:template name ="GetStylesFromSlide" >
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
    <xsl:for-each select ="document('ppt/presentation.xml')/p:presentation/p:sldIdLst/p:sldId">
      <xsl:variable name ="SlideNumber">
        <xsl:value-of  select ="concat('slide',position())" />
      </xsl:variable>
      <xsl:variable name ="SlideId">
        <xsl:value-of  select ="concat(concat('slide',position()),'.xml')" />
      </xsl:variable>
      <xsl:variable name ="LayoutFileName">
        <xsl:call-template name="GetLayOutforSlide">
          <xsl:with-param name="SlideName" select="$SlideId">  </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <!--Code by Vijayeta for Bullets,set style name in case of default bullets-->
      <xsl:variable name ="slideRel">
        <xsl:value-of select ="concat(concat('ppt/slides/_rels/',$SlideNumber,'.xml'),'.rels')"/>
      </xsl:variable>
      <xsl:variable name ="layout" >
        <xsl:for-each select ="document($slideRel)//rels:Relationships/rels:Relationship">
          <xsl:if test ="document($slideRel)//rels:Relationships/rels:Relationship/@Target[contains(.,'slideLayouts')]">
            <xsl:value-of select ="substring-after((@Target),'../slideLayouts/')"/>
          </xsl:if>
        </xsl:for-each>
      </xsl:variable>
      <xsl:variable name ="bulletTypeBool">
        <!--<xsl:for-each select ="document($slideRel)//rels:Relationships/rels:Relationship">-->
        <!--<xsl:if test ="document($slideRel)//rels:Relationships/rels:Relationship/@Target[contains(.,'slideLayouts')]">-->
        <!--<xsl:variable name ="layout" select ="substring-after((@Target),'../slideLayouts/')"/>-->
        <xsl:for-each select ="document(concat('ppt/slideLayouts/',$layout))//p:sldLayout">
          <xsl:choose >
            <!-- Changes made by vijayeta, bug fix, 1739703, date:10-7-07-->
            <xsl:when test ="p:cSld/@name[contains(.,'Content')] or p:cSld/@name[contains(.,'Title and')] or p:cSld/@name[contains(.,'Two Content')] or  p:cSld/@name[contains(.,'Comparison')]">
              <xsl:value-of select ="'true'"/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select ="'false'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
        <!--</xsl:if>-->
        <!--</xsl:for-each>-->
      </xsl:variable >
      <xsl:for-each select ="document(concat('ppt/slides/',$SlideId))/p:sld/p:cSld/p:spTree/p:sp/p:txBody">
        <xsl:variable name ="ParaId">
          <xsl:value-of select ="concat($SlideNumber,concat('PARA',position()))"/>
        </xsl:variable>
        <!-- ADDED by Vipul to decide text box type i.e. Title, subtitle, outline, textbox-->
        <!--Start-->
        <xsl:variable name="var_fontScale">
          <xsl:if test="./a:bodyPr/a:normAutofit/@fontScale">
            <xsl:value-of select="./a:bodyPr/a:normAutofit/@fontScale"/>
          </xsl:if>
          <xsl:if test="not(./a:bodyPr/a:normAutofit/@fontScale)">
            <xsl:value-of select="'100000'"/>
          </xsl:if>
        </xsl:variable>
        <xsl:variable name="var_lnSpcReduction">
          <xsl:if test="./a:bodyPr/a:normAutofit/@lnSpcReduction">
            <xsl:value-of select="./a:bodyPr/a:normAutofit/@lnSpcReduction"/>
          </xsl:if>
          <xsl:if test="not(./a:bodyPr/a:normAutofit/@lnSpcReduction)">
            <xsl:value-of select="'0'"/>
          </xsl:if>
        </xsl:variable>
        <xsl:variable name="var_TextBoxType">
          <xsl:choose>
            <xsl:when test="./parent::node()/p:nvSpPr/p:nvPr/p:ph/@type='ctrTitle' or ./parent::node()/p:nvSpPr/p:nvPr/p:ph/@type='title'">
              <xsl:value-of select="'title'"/>
            </xsl:when>
            <xsl:when test="./parent::node()/p:nvSpPr/p:nvPr/p:ph/@type='subTitle'">
              <xsl:value-of select="'subtitle'"/>
            </xsl:when>
            <xsl:when test="not(./parent::node()/p:nvSpPr/p:nvPr/p:ph/@type) and ./parent::node()/p:nvSpPr/p:nvPr/p:ph/@idx">
              <xsl:value-of select="'outline'"/>
            </xsl:when>
            <xsl:when test="./parent::node()/p:nvSpPr/p:nvPr/p:ph/@type='body'">
              <xsl:value-of select="'body'"/>
            </xsl:when>
            <xsl:when test="./parent::node()/p:nvSpPr/p:nvPr/p:ph/@type='ftr'">
              <xsl:value-of select="'ftr'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="./parent::node()/p:nvSpPr/p:nvPr/p:ph/@type"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="var_index">
          <xsl:choose>
            <xsl:when test="./parent::node()/p:nvSpPr/p:nvPr/p:ph/@idx">
              <xsl:value-of select="./parent::node()/p:nvSpPr/p:nvPr/p:ph/@idx"/>
            </xsl:when>
          </xsl:choose>
        </xsl:variable>
        <!--End-->
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
					<xsl:value-of select ="concat($SlideNumber,'textboxshape_List',$textNumber)"/>
				</xsl:when>
				<xsl:otherwise >
					<xsl:value-of select ="concat($SlideNumber,'List',position())"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<!-- Added by vijayeta, on 16th july-->
		<xsl:variable name ="newLayout" >
			<xsl:for-each select ="document($slideRel)//node()/@Target[contains(.,'slideLayouts')]">
				<xsl:value-of  select ="concat('ppt/slideLayouts/_rels/',substring-after(.,'../slideLayouts/'))"/>
			</xsl:for-each>
		</xsl:variable>
		<!-- aDDED BY VIJAYEA, TO GET SLIDE MASTER, ON 15TH JULY-->
		<xsl:variable name ="slideMaster">
			<xsl:call-template name ="getmasterFromLayout">
				<xsl:with-param name ="layoutFile" select ="$newLayout"/>
			</xsl:call-template>
		</xsl:variable>
		<!-- aDDED BY VIJAYEA, TO GET SLIDE MASTER, ON 15TH JULY-->
        <!--If bullets present-->
        <xsl:if test ="(a:p/a:pPr/a:buChar) or (a:p/a:pPr/a:buAutoNum) or (a:p/a:pPr/a:buBlip) ">
          <xsl:call-template name ="insertBulletStyle">
            <xsl:with-param name ="slideRel" select ="$slideRel"/>
            <xsl:with-param name ="ParaId" select ="$ParaId"/>
            <xsl:with-param name ="listStyleName" select ="$listStyleName"/>
			  <xsl:with-param name ="slideMaster" select ="$slideMaster"/>
			  <xsl:with-param name ="var_TextBoxType" select ="$var_TextBoxType"/>
			  <xsl:with-param name ="var_index" select ="$var_index"/>
          </xsl:call-template>
        </xsl:if>
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
				  <xsl:with-param name ="slideLayout" select ="$layout"/>
				  <xsl:with-param name ="slideMaster" select ="$slideMaster"/>
				  <xsl:with-param name ="var_TextBoxType" select ="$var_TextBoxType"/>
				  <xsl:with-param name ="var_index" select ="$var_index"/>
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
              <xsl:choose>
                <xsl:when test="$var_TextBoxType='outline'">
                  <xsl:call-template name="tmpCommanOutlineParaProperty">
                    <xsl:with-param name="spType" select="$var_TextBoxType"/>
                    <xsl:with-param name="LayoutFileName" select="$LayoutFileName"/>
                    <xsl:with-param name="index" select="$var_index"/>
                    <xsl:with-param name="lnSpcReduction" select="$var_lnSpcReduction"/>
                    <xsl:with-param name="level" select="$levelForDefFont"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:call-template name="tmpCommanParaProperty">
                    <xsl:with-param name="spType" select="$var_TextBoxType"/>
                    <xsl:with-param name="LayoutFileName" select="$LayoutFileName"/>
                    <xsl:with-param name="index" select="$var_index"/>
                    <xsl:with-param name="lnSpcReduction" select="$var_lnSpcReduction"/>
                  </xsl:call-template>
                </xsl:otherwise>
              </xsl:choose>
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
                    or (./parent::node()/p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'body')] and ./parent::node()/p:nvSpPr/p:cNvPr/@name[contains(.,'Rectangle')]) 
                      or (./parent::node()/p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'body')] and ./parent::node()/p:nvSpPr/p:cNvPr/@name[contains(.,'Content')])">
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
              <!--</xsl:if>-->
              <!--End of Code inserted by VijayetaFor Bullets, Enable Numbering-->
              <!-- Ends -->
            </style:paragraph-properties >
          </style:style>
          <!-- Modified by pradeep for fix 1731885-->
          <xsl:for-each select ="node()" >
            <!-- Add here-->
            <xsl:if test ="name()='a:r'">
              <style:style style:name="{generate-id()}"  style:family="text">
                <style:text-properties>
                  <xsl:choose>
                    <xsl:when test="$var_TextBoxType='outline'">
                      <xsl:call-template name="tmpCommanOutlineTextProperty">
                        <xsl:with-param name="spType" select="$var_TextBoxType"/>
                        <xsl:with-param name="DefFont" select="$DefFont"/>
                        <xsl:with-param name="LayoutFileName" select="$LayoutFileName"/>
                        <xsl:with-param name="index" select="$var_index"/>
                        <xsl:with-param name="fontscale" select="$var_fontScale"/>
                        <xsl:with-param name="level" select="$levelForDefFont"/>
                      </xsl:call-template>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:call-template name="tmpCommanTextProperty">
                        <xsl:with-param name="spType" select="$var_TextBoxType"/>
                        <xsl:with-param name="DefFont" select="$DefFont"/>
                        <xsl:with-param name="LayoutFileName" select="$LayoutFileName"/>
                        <xsl:with-param name="index" select="$var_index"/>
                        <xsl:with-param name="fontscale" select="$var_fontScale"/>
                      </xsl:call-template>
                    </xsl:otherwise>
                  </xsl:choose>
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
                  <xsl:if test ="not(./@sz)">
                    <!--<xsl:value-of select ="parent::node()/parent::node()/parent::node()/child::node()/p:nvPr/p:ph/@type"/>-->
                    <xsl:variable name ="cnvPrId">
                      <xsl:value-of select ="parent::node()/parent::node()/parent::node()/p:nvSpPr/p:nvPr/p:ph/@type"/>
                    </xsl:variable>
                    <xsl:variable name ="Id">
                      <xsl:value-of select ="parent::node()/parent::node()/parent::node()/p:nvSpPr/p:nvPr/p:ph/@idx"/>
                    </xsl:variable>
                    <xsl:variable name ="layoutFontSize">
                      <xsl:call-template name ="getDefaultFontSize">
                        <xsl:with-param name ="prId" select ="$cnvPrId" />
                        <xsl:with-param name ="sldName" select ="$SlideNumber"/>
                        <xsl:with-param name ="Id" select ="$Id"/>
                        <xsl:with-param name ="levelForDefFont" select ="$levelForDefFont"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:variable name ="lName">
                      <xsl:for-each select ="parent::node()/parent::node()/parent::node()/child::node()/p:nvPr/p:ph/@type">
                        <xsl:value-of select ="."/>
                      </xsl:for-each>
                    </xsl:variable>
                    <!-- Added By Vijayeta,Check if textbox is present-->
                    <xsl:variable name ="isContent">
                      <xsl:if test="parent::node()/parent::node()/parent::node()/p:nvSpPr/p:nvSpPr/p:cNvPr/@name[contains(.,'Content')]">
                        <xsl:value-of select ="'true'"/>
                      </xsl:if>
                      <!--End, Added By Vijayeta,Check if textbox is present-->
                      <xsl:if test="not(parent::node()/parent::node()/parent::node()/p:nvSpPr/p:nvSpPr/p:cNvPr/@name[contains(.,'Content')])">
                        <xsl:value-of select ="'false'"/>
                      </xsl:if>
                    </xsl:variable>
                    <xsl:choose >
                      <!--<xsl:when test ="$layoutFontSize &gt; 0">
												<xsl:attribute name ="fo:font-size"	>
													<xsl:value-of select ="concat($layoutFontSize,'pt')"/>
												</xsl:attribute >
                      </xsl:when>-->
                      <xsl:when test ="$lName='title' or $lName='ctrTitle'">
                        <xsl:attribute name ="fo:font-size"	>
                          <xsl:value-of select ="concat(format-number($DefFontSizeTitle div 100,'#.##'), 'pt')"/>
                        </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$lName='title' or $lName='ctrTitle'">
                        <xsl:attribute name ="fo:font-size"	>
                          <xsl:value-of select ="concat(format-number($DefFontSizeTitle div 100,'#.##'), 'pt')"/>
                        </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="not($lName='title' or $lName='ctrTitle') and  $isContent = 'false' ">
                        <xsl:attribute name ="fo:font-size"	>
                          <xsl:value-of select ="concat(format-number($DefFontSizeOther div 100,'#.##'), 'pt')"/>
                        </xsl:attribute>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:attribute name ="fo:font-size"	>
                          <!-- Code By vijayeta,to set fontsize,if levels present-->
                          <xsl:variable name ="newFontSize">
                            <xsl:variable name ="fontSizeBody">
                              <xsl:if test ="$bulletTypeBool='true'">
                                <xsl:if test =" $levelForDefFont != '0'">
                                  <!--<xsl:call-template name ="getFontSizeForLevelsFromMaster">
																		<xsl:with-param name ="levelForDefFont" select ="$levelForDefFont"/>
                                  </xsl:call-template>-->
                                  <xsl:value-of select ="$layoutFontSize"/>

                                </xsl:if>
                                <xsl:if test =" $levelForDefFont = '0'">
                                  <xsl:value-of select ="$DefFontSizeBody"/>
                                </xsl:if>
                              </xsl:if>
                              <xsl:if test ="not($bulletTypeBool='true')">
                                <xsl:value-of select ="$DefFontSizeBody"/>
                              </xsl:if>
                            </xsl:variable >
                            <xsl:value-of select="$fontSizeBody"/>
                          </xsl:variable >
                          <!--Code Added By Vijayeta,For Line Spacing
                             If the input file has an option to Decrease the font size to fit.
                             100% being equal to 100000 and 1% to 1000,divide the scale, by which the actual font size is to be reduced, by 1000.
                             final font size=((ActualFontSize/100)* (scale/1000))/100-->
                          <xsl:if test ="parent::node()/parent::node()/a:bodyPr/a:normAutofit/@fontScale">
                            <xsl:variable name ="fontScaleReductionValue">
                              <xsl:value-of select ="parent::node()/parent::node()/a:bodyPr/a:normAutofit/@fontScale div 1000 "/>
                            </xsl:variable>
                            <xsl:value-of select ="concat(round(format-number(($newFontSize * $fontScaleReductionValue) div 10000,'#.##')), 'pt')"/>
                          </xsl:if>
                          <!-- Do not change the font size if the option is selected to be off-->
                          <xsl:if test ="not(parent::node()/parent::node()/a:bodyPr/a:normAutofit/@fontScale)">
                            <xsl:value-of select ="concat(format-number($newFontSize div 100,'#.##'), 'pt')"/>
                          </xsl:if>
                          <!--End of Code Added By Vijayeta,For Line Spacing -->
                        </xsl:attribute>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:if>
                </style:text-properties>
              </style:style>
            </xsl:if>
          </xsl:for-each >
        </xsl:for-each>
      </xsl:for-each>
    </xsl:for-each>
  </xsl:template>
  <!--@@ Pradeep Nemadi Paragraph Tab stops code is added -Start -->
  <xsl:template name ="paragraphTabstops">
    <style:tab-stops>
      <xsl:for-each select ="a:pPr/a:tabLst/a:tab">
        <style:tab-stop>
          <xsl:attribute name ="style:position">
            <xsl:value-of select ="concat(format-number(@pos div 360000,'#.##'),'cm') "/>
          </xsl:attribute>
          <xsl:if test ="@algn">
            <xsl:attribute name ="style:type">
              <xsl:choose >
                <xsl:when test ="@algn='ctr'">
                  <xsl:value-of select ="'center'"/>
                </xsl:when>
                <xsl:when test ="@algn='r'">
                  <xsl:value-of select ="'right'"/>
                </xsl:when>
                <xsl:when test ="@algn='l'">
                  <xsl:value-of select ="'left'"/>
                </xsl:when>
                <xsl:when test ="@algn='dec'">
                  <xsl:value-of select ="'char'"/>
                </xsl:when>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
        </style:tab-stop >
      </xsl:for-each>
    </style:tab-stops>
  </xsl:template>
  <!--@@ Pradeep Nemadi Paragraph Tab stops code is added -End -->
  <xsl:template name="GenerateId">
    <xsl:param name="node"/>
    <xsl:param name="nodetype"/>
    <xsl:variable name="positionInGroup">
      <xsl:for-each select="key(concat($nodetype,'s'), '')">
        <xsl:if test="generate-id($node) = generate-id(.)">
          <xsl:value-of select="position() + 1"/>
        </xsl:if>
      </xsl:for-each>
    </xsl:variable>
    <xsl:value-of select="$positionInGroup"/>
  </xsl:template>
  <xsl:template name ="LayoutType">
    <xsl:param name ="LayoutStyle"/>
    <xsl:choose >
      <xsl:when test ="$LayoutStyle='ctrTitle'">
        <xsl:value-of select ="'title'"/>
      </xsl:when>
      <xsl:when test ="$LayoutStyle='subTitle'">
        <xsl:value-of select ="'subtitle'"/>
        <!--Edited by vipul for case sensitive-->
      </xsl:when>
      <xsl:when test ="$LayoutStyle=''">
        <xsl:value-of select ="'outline'"/>
      </xsl:when>
      <xsl:when test ="$LayoutStyle='title'">
        <xsl:value-of select ="'title'"/>
      </xsl:when>
      <xsl:otherwise >
        <xsl:value-of select ="'outline'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!-- Added by Lohith A R : Custom Slide Show -->
  <xsl:template name="InsertPresentationSettings">
    <presentation:settings>
      <xsl:attribute name ="presentation:show">
        <xsl:value-of select="document('ppt/presentation.xml')/p:presentation/p:custShowLst/p:custShow/@name"/>
      </xsl:attribute>
      <xsl:for-each select="document('ppt/presentation.xml')/p:presentation/p:custShowLst/p:custShow">
        <presentation:show>
          <xsl:attribute name ="presentation:name"	>
            <xsl:value-of select ="@name"/>
          </xsl:attribute>
          <xsl:attribute name ="presentation:pages"	>
            <xsl:for-each select ="p:sldLst/p:sld">
              <xsl:variable name="SlideIdVal">
                <xsl:value-of select ="@r:id"/>
              </xsl:variable>
              <xsl:variable name="SlideVal">
                <xsl:value-of select="document('ppt/_rels/presentation.xml.rels')/rels:Relationships/rels:Relationship[@Id=$SlideIdVal]/@Target"/>
              </xsl:variable>
              <xsl:variable name="CoustomSlideList">
                <xsl:value-of select="concat('page',substring-before(substring-after($SlideVal,'slides/slide'),'.xml'),',')"/>
              </xsl:variable>
              <!--<xsl:variable name="CoustomSlideListForOdp">
                  <xsl:value-of select="substring($CoustomSlideList,1,string-length($CoustomSlideList)-1)"/>
                </xsl:variable>-->
              <xsl:value-of select="$CoustomSlideList"/>
            </xsl:for-each>
          </xsl:attribute>
        </presentation:show>
      </xsl:for-each>
    </presentation:settings>
  </xsl:template>
  <!-- Added by Lohith A R : Custom Slide Show -->
  <xsl:template name="ConvertPoints">
    <xsl:param name="length"/>
    <xsl:param name="unit"/>
    <xsl:variable name="lengthVal">
      <xsl:choose>
        <xsl:when test="contains($length,'pt')">
          <xsl:value-of select="substring-before($length,'pt')"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$length"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$lengthVal='0' or $lengthVal=''">
        <xsl:value-of select="concat(0, $unit)"/>
      </xsl:when>
      <xsl:when test="$unit = 'cm'">
        <xsl:value-of select="concat(format-number($lengthVal * 2.54 div 72,'#.###'),'cm')"/>
      </xsl:when>
      <xsl:when test="$unit = 'mm'">
        <xsl:value-of select="concat(format-number($lengthVal * 25.4 div 72,'#.###'),'mm')"/>
      </xsl:when>
      <xsl:when test="$unit = 'in'">
        <xsl:value-of select="concat(format-number($lengthVal div 72,'#.###'),'in')"/>
      </xsl:when>
      <xsl:when test="$unit = 'pt'">
        <xsl:value-of select="concat($lengthVal,'pt')"/>
      </xsl:when>
      <xsl:when test="$unit = 'pica'">
        <xsl:value-of select="concat(format-number($lengthVal div 12,'#.###'),'pica')"/>
      </xsl:when>
      <xsl:when test="$unit = 'dpt'">
        <xsl:value-of select="concat($lengthVal,'dpt')"/>
      </xsl:when>
      <xsl:when test="$unit = 'px'">
        <xsl:value-of select="concat(format-number($lengthVal * 96.19 div 72,'#.###'),'px')"/>
      </xsl:when>
      <xsl:when test="not($lengthVal)">
        <xsl:value-of select="concat(0,'cm')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$lengthVal"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name ="textSpacingProp">
    <xsl:for-each select ="document('ppt/presentation.xml')/p:presentation/p:sldIdLst/p:sldId">
      <xsl:variable name ="SlideId">
        <xsl:value-of  select ="concat(concat('slide',position()),'.xml')" />
      </xsl:variable>
      <!--added by vipul to get layout File name(i.e. slideLayout1.xml) -->
      <!--Start-->
      <xsl:variable name ="LayoutFileName">
        <xsl:call-template name="GetLayOutforSlide">
          <xsl:with-param name="SlideName" select="$SlideId">  </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <!--End-->
      <xsl:for-each select ="document(concat('ppt/slides/',$SlideId))/p:sld/p:cSld/p:spTree/p:sp">
        <xsl:variable name="var_TextBoxType">
          <xsl:choose>
            <xsl:when test="p:nvSpPr/p:nvPr/p:ph/@type='ctrTitle' or p:nvSpPr/p:nvPr/p:ph/@type='title'">
              <xsl:value-of select="'title'"/>
            </xsl:when>
            <xsl:when test="p:nvSpPr/p:nvPr/p:ph/@type='subTitle'">
              <xsl:value-of select="'subtitle'"/>
            </xsl:when>
            <xsl:when test="not(p:nvSpPr/p:nvPr/p:ph/@type) and ./parent::node()/p:nvSpPr/p:nvPr/p:ph/@idx">
              <xsl:value-of select="'outline'"/>
            </xsl:when>
            <xsl:when test="p:nvSpPr/p:nvPr/p:ph/@type='body'">
              <xsl:value-of select="'body'"/>
            </xsl:when>
            <xsl:when test="p:nvSpPr/p:nvPr/p:ph/@type='ftr'">
              <xsl:value-of select="'ftr'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="p:nvSpPr/p:nvPr/p:ph/@type"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="var_index">
          <xsl:choose >
            <xsl:when test="p:nvSpPr/p:nvPr/p:ph/@idx">
              <xsl:value-of select="p:nvSpPr/p:nvPr/p:ph/@idx"/>
            </xsl:when>
          </xsl:choose>
        </xsl:variable>

        <style:style style:family="presentation" >
          <xsl:attribute name ="style:name">
            <xsl:value-of  select ="concat(substring-before($SlideId,'.xml'),'pr',position())"/>
          </xsl:attribute>
          <!--fo:wrap-option -->
          <!--<xsl:value-of select ="."/>-->
          <style:graphic-properties draw-stroke="none">
            <xsl:call-template name="tmpCommanGraphicProperty">
              <xsl:with-param name="spType" select="$var_TextBoxType"/>
              <xsl:with-param name="LayoutFileName" select="$LayoutFileName"/>
              <xsl:with-param name="index" select="$var_index"/>
            </xsl:call-template>
          </style:graphic-properties>

        </style:style>

      </xsl:for-each >

    </xsl:for-each>
  </xsl:template>
  <!--Maps the footer date format with pptx to odp -->
  <xsl:template name ="FooterDateFormat">
    <xsl:param name ="type" />
    <xsl:choose >
      <xsl:when test ="$type ='datetime1'">
        <xsl:value-of select ="'D3'"/>
      </xsl:when>
      <xsl:when test ="$type ='datetime2'">
        <xsl:value-of select ="'D8'"/>
      </xsl:when>
      <xsl:when test ="$type ='datetime4'">
        <xsl:value-of select ="'D6'"/>
      </xsl:when>
      <xsl:when test ="$type ='datetime4'">
        <xsl:value-of select ="'D5'"/>
      </xsl:when>
      <xsl:when test ="$type ='datetime8'">
        <xsl:value-of select ="'D3T2'"/>
      </xsl:when>
      <xsl:when test ="$type ='datetime8'">
        <xsl:value-of select ="'D3T5'"/>
      </xsl:when>
      <xsl:when test ="$type ='datetime10'">
        <xsl:value-of select ="'T2'"/>
      </xsl:when>
      <xsl:when test ="$type ='datetime11'">
        <xsl:value-of select ="'T3'"/>
      </xsl:when>
      <xsl:when test ="$type ='datetime12'">
        <xsl:value-of select ="'T5'"/>
      </xsl:when>
      <xsl:when test ="$type ='datetime13'">
        <xsl:value-of select ="'T6'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select ="''"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name ="DateFormats">
    <number:date-style style:name="D8">
      <number:day-of-week number:style="long" />
      <number:text>,</number:text>
      <number:day />
      <number:text>.</number:text>
      <number:month number:style="long" number:textual="true" />
      <number:text />
      <number:year number:style="long" />
    </number:date-style>
    <number:date-style style:name="D6">
      <number:day />
      <number:text>.</number:text>
      <number:month number:style="long" number:textual="true" />
      <number:text />
      <number:year number:style="long" />
    </number:date-style>
    <number:date-style style:name="D5">
      <number:day />
      <number:text>.</number:text>
      <number:month number:textual="true" />
      <number:text />
      <number:year number:style="long" />
    </number:date-style>
    <number:date-style style:name="D3T2">
      <number:day number:style="long" />
      <number:text>.</number:text>
      <number:month number:style="long" />
      <number:text>.</number:text>
      <number:year />
      <number:text />
      <number:hours />
      <number:text>:</number:text>
      <number:minutes />
    </number:date-style>
    <number:date-style style:name="D3T5">
      <number:day number:style="long" />
      <number:text>.</number:text>
      <number:month number:style="long" />
      <number:text>.</number:text>
      <number:year />
      <number:text />
      <number:hours />
      <number:text>:</number:text>
      <number:minutes />
      <number:am-pm />
    </number:date-style>
    <number:time-style style:name="T2">
      <number:hours />
      <number:text>:</number:text>
      <number:minutes />
    </number:time-style>
    <number:time-style style:name="T3">
      <number:hours />
      <number:text>:</number:text>
      <number:minutes />
      <number:text>:</number:text>
      <number:seconds />
    </number:time-style>
    <number:time-style style:name="T5">
      <number:hours />
      <number:text>:</number:text>
      <number:minutes />
      <number:am-pm />
    </number:time-style>
    <number:time-style style:name="T6">
      <number:hours />
      <number:text>:</number:text>
      <number:minutes />
      <number:text>:</number:text>
      <number:seconds />
      <number:am-pm />
    </number:time-style>
    <number:date-style style:name="D3">
      <number:day number:style="long" />
      <number:text>.</number:text>
      <number:month number:style="long" />
      <number:text>.</number:text>
      <number:year />
    </number:date-style>
  </xsl:template>
  <xsl:template name ="getDefaultFontSize">
    <xsl:param name ="prId"  />
    <xsl:param name ="Id"  />
    <xsl:param name ="sldName" />
    <xsl:param name ="sz" />
    <xsl:param name ="levelForDefFont"/>
    <xsl:variable name ="slLtName">
      <xsl:value-of select ="concat('ppt/slides/_rels/',$sldName,'.xml.rels')"/>
    </xsl:variable>
    <xsl:variable name ="layoutName">
      <xsl:for-each  select ="document($slLtName)//node()/@Target[contains(.,'slideLayouts')]">
        <xsl:value-of select ="concat('ppt',substring(.,3))"/>
      </xsl:for-each >
    </xsl:variable>
    <!--<xsl:for-each select ="document($layoutName)/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:cNvPr[@id=$prId]">
			<xsl:value-of  select ="parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz div 100"/>
		</xsl:for-each>-->
    <!--<xsl:value-of  select ="document($layoutName)/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type=$prId]/parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz div 100"/>-->
    <xsl:choose >
      <xsl:when test ="$Id !='' and $prId != ''"	>
        <xsl:for-each select ="document($layoutName)/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@idx=$Id and @type=$prId]">
          <!--<xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz div 100,'#.##')"/>-->
          <!-- Added by vijayeta,Font Size for levels-->
          <xsl:call-template name ="getFontSizeForLevelsFromLayout">
            <xsl:with-param name ="levelForDefFont" select ="$levelForDefFont+1"/>
          </xsl:call-template>
          <!-- Added by vijayeta,Font Size for levels-->
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="$prId =''  and $Id !=''">
        <xsl:for-each select ="document($layoutName)/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@idx=$Id]">
          <!-- Added by vijayeta,Font Size for levels-->
          <xsl:call-template name ="getFontSizeForLevelsFromLayout">
            <xsl:with-param name ="levelForDefFont" select ="$levelForDefFont+1"/>
          </xsl:call-template>
          <!--<xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz div 100,'#.##')"/>-->
          <!-- Added by vijayeta,Font Size for levels-->
        </xsl:for-each>
      </xsl:when>
      <xsl:otherwise >
        <xsl:for-each select ="document($layoutName)/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type=$prId]">
          <xsl:if test ="@idx">
            <!-- Added by vijayeta,Font Size for levels-->
            <xsl:call-template name ="getFontSizeForLevelsFromLayout">
              <xsl:with-param name ="levelForDefFont" select ="$levelForDefFont+1"/>
            </xsl:call-template>
            <!-- Added by vijayeta,Font Size for levels-->
            <!--<xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz div 100,'#.##')"/>-->
          </xsl:if>
          <xsl:if  test ="not(@idx)">
            <!-- Added by vijayeta,Font Size for levels-->
            <xsl:call-template name ="getFontSizeForLevelsFromLayout">
              <xsl:with-param name ="levelForDefFont" select ="$levelForDefFont+1"/>
            </xsl:call-template>
            <!-- Added by vijayeta,Font Size for levels-->
            <!--<xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz div 100,'#.##')"/>-->
          </xsl:if>
        </xsl:for-each>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!-- Template added by vijayeta to fix bug 1739076,get font family from slide layout-->
  <xsl:template name ="FontTypeFromLayout">
    <xsl:param name ="prId"  />
    <xsl:param name ="Id"  />
    <xsl:param name ="sldName" />
    <xsl:param name ="sz" />
    <xsl:param name ="levelForDefFont"/>
    <xsl:param name ="DefFont"/>
    <xsl:param name ="AttrType"/>
    <xsl:variable name ="slLtName">
      <xsl:value-of select ="concat('ppt/slides/_rels/',$sldName,'.xml.rels')"/>
    </xsl:variable>
    <xsl:variable name ="layoutName">
      <xsl:for-each  select ="document($slLtName)//node()/@Target[contains(.,'slideLayouts')]">
        <xsl:value-of select ="concat('ppt',substring(.,3))"/>
      </xsl:for-each >
    </xsl:variable>
    <xsl:choose >
      <xsl:when test ="$Id !='' and $prId != ''"	>
        <xsl:for-each select ="document($layoutName)/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@idx=$Id and @type=$prId]">
          <xsl:call-template name ="getFontFamilyForLevelsFromLayout">
            <xsl:with-param name ="levelForDefFont">
              <xsl:value-of select="$levelForDefFont+1"/>
            </xsl:with-param>
            <xsl:with-param name ="AttrType" select ="$AttrType"/>
            <xsl:with-param name ="DefFont" select ="$DefFont"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="$prId =''  and $Id !=''">
        <xsl:for-each select ="document($layoutName)/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@idx=$Id]">
          <xsl:call-template name ="getFontFamilyForLevelsFromLayout">
            <xsl:with-param name ="levelForDefFont">
              <xsl:value-of select="$levelForDefFont+1"/>
            </xsl:with-param>

            <xsl:with-param name ="DefFont" select ="$DefFont"/>
            <xsl:with-param name ="AttrType" select ="$AttrType"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
      <xsl:otherwise >
        <xsl:for-each select ="document($layoutName)/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type=$prId]">
          <xsl:if test ="@idx">
            <xsl:call-template name ="getFontFamilyForLevelsFromLayout">
              <xsl:with-param name ="levelForDefFont">
                <xsl:value-of select="$levelForDefFont+1"/>
              </xsl:with-param>
              <xsl:with-param name ="DefFont" select ="$DefFont"/>
              <xsl:with-param name ="AttrType" select ="$AttrType"/>
            </xsl:call-template>
          </xsl:if>
          <xsl:if  test ="not(@idx)">
            <xsl:call-template name ="getFontFamilyForLevelsFromLayout">
              <xsl:with-param name ="levelForDefFont">
                <xsl:value-of select="$levelForDefFont+1"/>
              </xsl:with-param>
              <xsl:with-param name ="DefFont" select ="$DefFont"/>
              <xsl:with-param name ="AttrType" select ="$AttrType"/>
            </xsl:call-template>
          </xsl:if>
        </xsl:for-each>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!-- Template added by vijayeta to fix bug 1739076,get font family from slide layout-->
  <!--@@Get default inner margins settings from slide master -->
  <xsl:template name ="DefaultTextSpacingProp">
    <xsl:param name ="phType"/>
    <xsl:param name ="defType"/>
    <xsl:param name ="slideId"/>


    <xsl:variable name ="acPhType">
      <!--<xsl:choose >
				<xsl:when test ="$phType!=''">
					<xsl:value-of select ="$phType"/>
				</xsl:when>
				<xsl:otherwise >
					<xsl:value-of select ="'body'"/>
				</xsl:otherwise>
			</xsl:choose>-->
      <xsl:value-of  select ="./p:nvSpPr/p:nvPr/p:ph/@type"/>
    </xsl:variable>
    <xsl:variable name ="idx">
      <xsl:value-of  select ="./p:nvSpPr/p:nvPr/p:ph/@idx"/>
    </xsl:variable>

    <xsl:variable name ="slLtName">
      <xsl:value-of select ="concat('ppt/slides/_rels/',$slideId,'.rels')"/>
    </xsl:variable>
    <xsl:variable name ="layoutName">
      <xsl:for-each  select ="document($slLtName)//node()/@Target[contains(.,'slideLayouts')]">
        <xsl:value-of select ="concat('ppt',substring(.,3))"/>
      </xsl:for-each >
    </xsl:variable>
    <xsl:for-each select ="document($layoutName)//p:nvPr/p:ph[@type = $acPhType or @idx=$idx]">
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
                    <xsl:value-of  select ="'top'"/>
                  </xsl:when>
                  <!-- Middle-->
                  <xsl:when test ="p:txBody/a:bodyPr/@anchor = 'ctr' ">
                    <xsl:value-of  select ="'middle'"/>
                  </xsl:when>
                  <!-- bottom-->
                  <xsl:when test ="p:txBody/a:bodyPr/@anchor = 'b' ">
                    <xsl:value-of  select ="'bottom'"/>
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
            <xsl:value-of  select ="'0'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="GetLayOutName">
    <xsl:param name ="slideRelName"/>
    <xsl:variable name ="LtName">
      <xsl:value-of select ="concat('ppt/slides/_rels/',$slideRelName,'.rels')"/>
    </xsl:variable >
    <xsl:for-each select ="document($LtName)//node()/@Target[contains(.,'slideLayouts')]">
      <xsl:choose >
        <xsl:when test =" contains(.,'1') ">
          <xsl:value-of  select ="'AL1T0'"/>
        </xsl:when>
        <xsl:when test =" contains(.,'2') ">
          <xsl:value-of  select ="'AL2T1'"/>
        </xsl:when>
        <xsl:when test =" contains(.,'4') ">
          <xsl:value-of  select ="'AL3T3'"/>
        </xsl:when>
        <xsl:when test =" contains(.,'6') ">
          <xsl:value-of  select ="'AL1T19'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of  select ="'AL0T0'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>
  <!--Added for Layout-->
	<!-- aDDED BY VIJAYEA, TO GET SLIDE MASTER, ON 15TH JULY-->
	<xsl:template name ="getmasterFromLayout">
		<xsl:param name ="layoutFile"/>
		<xsl:variable name="slideMasterName">
			<xsl:for-each select ="document(concat($layoutFile,'.rels'))//node()/@Target[contains(.,'slideMasters')]">
				<xsl:value-of  select ="substring-after(.,'../slideMasters/')"/>
			</xsl:for-each>
		</xsl:variable>
		<xsl:value-of select ="$slideMasterName"/>
	</xsl:template>
	<!-- aDDED BY VIJAYEA, TO GET SLIDE MASTER, ON 15TH JULY-->
  <xsl:template name="GetLayOut">
    <xsl:param name="SlideNo" />
    <xsl:for-each select="document(concat('ppt/slides/_rels/','slide1.xml','.rels'))//rels:Relationships/rels:Relationship/@Target[contains(.,'slideLayouts')]">
      <xsl:value-of select="substring(.,17)" />
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="GetMasterFileName">
    <xsl:param name ="slideId"/>
    <xsl:variable name ="layoutName">
      <xsl:value-of select ="document(concat('ppt/slides/_rels/','slide',$slideId,'.xml.rels'))//node()/@Target[ contains(.,'Layout')]"/>
    </xsl:variable>
    <xsl:variable name ="lauoutReln">
      <xsl:value-of select ="concat('ppt/slideLayouts/_rels/', substring($layoutName,17),'.rels')"/>
    </xsl:variable>
    <xsl:variable name ="slideMaster">
      <xsl:value-of select ="document($lauoutReln)//node()/@Target[ contains(.,'slideMaster')]"/>
    </xsl:variable>
    <xsl:value-of select ="substring-before(substring($slideMaster,17),'.xml')"/>
  </xsl:template>
  <!-- Code By vijayeta,get font size for diff levels-->
  <xsl:template name ="getFontSizeForLevelsFromMaster">
    <xsl:param name ="levelForDefFont" />
    <xsl:for-each select ="document('ppt/slideMasters/slideMaster1.xml')//p:txStyles/p:bodyStyle">
      <xsl:choose >
        <xsl:when test ="$levelForDefFont='1'" >
          <xsl:value-of select ="a:lvl2pPr/a:defRPr/@sz"/>
        </xsl:when>
        <xsl:when test ="$levelForDefFont='2'" >
          <xsl:value-of select ="a:lvl3pPr/a:defRPr/@sz"/>
        </xsl:when>
        <xsl:when test ="$levelForDefFont='3'" >
          <xsl:value-of select ="a:lvl4pPr/a:defRPr/@sz"/>
        </xsl:when>
        <xsl:when test ="$levelForDefFont='4'" >
          <xsl:value-of select ="a:lvl5pPr/a:defRPr/@sz"/>
        </xsl:when>
        <xsl:when test ="$levelForDefFont='5'" >
          <xsl:value-of select ="a:lvl6pPr/a:defRPr/@sz"/>
        </xsl:when>
        <xsl:when test ="$levelForDefFont='6'" >
          <xsl:value-of select ="a:lvl7pPr/a:defRPr/@sz"/>
        </xsl:when>
        <xsl:when test ="$levelForDefFont='7'" >
          <xsl:value-of select ="a:lvl8pPr/a:defRPr/@sz"/>
        </xsl:when>
        <xsl:when test ="$levelForDefFont='8'" >
          <xsl:value-of select ="a:lvl9pPr/a:defRPr/@sz"/>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="getFontSizeForLevelsFromLayout">
    <xsl:param name ="levelForDefFont" />
    <xsl:choose >
      <xsl:when test ="$levelForDefFont='1'" >
        <xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz,'#.##')"/>
      </xsl:when>
      <xsl:when test ="$levelForDefFont='2'" >
        <xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:defRPr/@sz,'#.##')"/>
      </xsl:when>
      <xsl:when test ="$levelForDefFont='3'" >
        <xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:defRPr/@sz,'#.##')"/>
      </xsl:when>
      <xsl:when test ="$levelForDefFont='4'" >
        <xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:defRPr/@sz,'#.##')"/>
      </xsl:when>
      <xsl:when test ="$levelForDefFont='5'" >
        <xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:defRPr/@sz,'#.##')"/>
      </xsl:when>
      <xsl:when test ="$levelForDefFont='6'" >
        <xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:defRPr/@sz,'#.##')"/>
      </xsl:when>
      <xsl:when test ="$levelForDefFont='7'" >
        <xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:defRPr/@sz,'#.##')"/>
      </xsl:when>
      <xsl:when test ="$levelForDefFont='8'" >
        <xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:defRPr/@sz,'#.##')"/>
      </xsl:when>
      <xsl:when test ="$levelForDefFont='9'" >
        <xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:defRPr/@sz,'#.##')"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name ="getFontFamilyForLevelsFromLayout">
    <xsl:param name ="levelForDefFont" />
    <xsl:param name ="DefFont" />
    <xsl:param name ="AttrType" />
    <xsl:param name ="fc"/>

    <xsl:choose >
      <xsl:when test ="$levelForDefFont='1'" >
        <xsl:choose>
          <xsl:when test="$AttrType='Fontname'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:latin/@typeface">
              <xsl:attribute name ="fo:font-family">
                <xsl:value-of  select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontsize'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <!--<xsl:choose>
                  <xsl:when test="number($fc) =100000">-->
                <!--<xsl:value-of  select ="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz div 100,'#.##'),'pt')"/>-->
                <!--</xsl:when>
                  <xsl:otherwise>-->
                <!--<xsl:value-of  select ="concat(format-number( (parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz *($vip div 1000) )div 10000,'#.##'),'pt')"/>-->
                <xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz div 100,'#.##')"/>
                <!--</xsl:otherwise>
                </xsl:choose>-->

              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontweight'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@b='1'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'bold'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@b='0'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'normal'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Underline'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@u">
              <xsl:for-each select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr">
                <xsl:call-template name="tmpUnderLine"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='italic'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@i='1'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'italic'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@i='0'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'none'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontcolor'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr/@val) and not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/@val)">
              <xsl:choose>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
                  <xsl:attribute name ="fo:color">
                    <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val)"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
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
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <!--<xsl:when test="$AttrType='strike'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@strike">
              <xsl:attribute name ="style:text-line-through-style">
                <xsl:value-of select ="'solid'"/>
              </xsl:attribute>
              <xsl:choose >
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@strike='dblStrike'">
                  <xsl:attribute name ="style:text-line-through-type">
                    <xsl:value-of select ="'double'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@strike='sngStrike'">
                  <xsl:attribute name ="style:text-line-through-type">
                    <xsl:value-of select ="'single'"/>
                  </xsl:attribute>
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>-->
        </xsl:choose>
      </xsl:when>
      <xsl:when test ="$levelForDefFont='2'" >
        <xsl:choose>
          <xsl:when test="$AttrType='Fontname'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/a:latin/@typeface">
              <xsl:attribute name ="fo:font-family">
                <xsl:value-of  select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontsize'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <!--<xsl:choose>
                  <xsl:when test="number($fc) =100000">
                    <xsl:value-of  select ="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@sz div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>-->
                <!--<xsl:value-of  select ="concat(format-number( (parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@sz *(number($fc) div 1000) )div 10000,'#.##'),'pt')"/>-->
                <xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@sz div 100,'#.##')"/>
                <!--</xsl:otherwise>
                </xsl:choose>-->
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontweight'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@b='1'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'bold'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@b='0'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'normal'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Underline'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@u">
              <xsl:for-each select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr">
                <xsl:call-template name="tmpUnderLine"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='italic'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@i='1'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'italic'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@i='0'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'none'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <!--<xsl:when test="$AttrType='Fontcolor'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>-->
        </xsl:choose>
        <!--<xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:defRPr/a:latin/@typeface)">
					<xsl:value-of  select ="$DefFont"/>
				</xsl:if>-->
      </xsl:when>
      <xsl:when test ="$levelForDefFont='3'" >
        <xsl:choose>
          <xsl:when test="$AttrType='Fontname'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/a:latin/@typeface">
              <xsl:attribute name ="fo:font-family">
                <xsl:value-of  select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontsize'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <!--<xsl:if test="$fontscale!=''">
                  <xsl:value-of  select ="format-number(floor(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@sz * 0.00001),'#.##')"/>
                </xsl:if>
                <xsl:if test="$fontscale=''">-->
                <xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@sz div 100,'#.##')"/>
                <!--</xsl:if>-->
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontweight'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@b='1'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'bold'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@b='0'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'normal'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Underline'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@u">
              <xsl:for-each select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr">
                <xsl:call-template name="tmpUnderLine"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='italic'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@i='1'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'italic'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@i='0'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'none'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <!--<xsl:when test="$AttrType='Fontcolor'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="translate(concat('#',parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/a:solidFill/a:srgbClr/@val),$ucletters,$lcletters)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/a:solidFill/a:schemeClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/a:solidFill/a:srgbClr/@val) and not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/a:solidFill/a:schemeClr/@val)">
              <xsl:choose>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
                  <xsl:attribute name ="fo:color">
                    <xsl:value-of select ="translate(concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val),$ucletters,$lcletters)"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
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
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>-->
        </xsl:choose>
        <!--<xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:defRPr/a:latin/@typeface)">
					<xsl:value-of  select ="$DefFont"/>
				</xsl:if>-->
      </xsl:when>
      <xsl:when test ="$levelForDefFont='4'" >
        <xsl:choose>
          <xsl:when test="$AttrType='Fontname'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/a:latin/@typeface">
              <xsl:attribute name ="fo:font-family">
                <xsl:value-of  select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontsize'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@sz div 100,'#.##')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontweight'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@b='1'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'bold'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@b='0'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'normal'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Underline'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@u">
              <xsl:for-each select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr">
                <xsl:call-template name="tmpUnderLine"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='italic'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@i='1'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'italic'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@i='0'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'none'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <!--<xsl:when test="$AttrType='Fontcolor'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="translate(concat('#',parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/a:solidFill/a:srgbClr/@val),$ucletters,$lcletters)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/a:solidFill/a:schemeClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/a:solidFill/a:srgbClr/@val) and not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/a:solidFill/a:schemeClr/@val)">
              <xsl:choose>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
                  <xsl:attribute name ="fo:color">
                    <xsl:value-of select ="translate(concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val),$ucletters,$lcletters)"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
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
                </xsl:when>
              </xsl:choose>
				</xsl:if>
          </xsl:when>-->
        </xsl:choose>
        <!--<xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:defRPr/a:latin/@typeface)">
					<xsl:value-of  select ="$DefFont"/>
				</xsl:if>-->
      </xsl:when>
      <xsl:when test ="$levelForDefFont='5'" >
        <xsl:choose>
          <xsl:when test="$AttrType='Fontname'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/a:latin/@typeface">
              <xsl:attribute name ="fo:font-family">
                <xsl:value-of  select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontsize'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@sz div 100,'#.##')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontweight'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@b='1'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'bold'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@b='0'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'normal'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Underline'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@u">
              <xsl:for-each select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr">
                <xsl:call-template name="tmpUnderLine"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='italic'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@i='1'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'italic'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@i='0'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'none'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <!--<xsl:when test="$AttrType='Fontcolor'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="translate(concat('#',parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/a:solidFill/a:srgbClr/@val),$ucletters,$lcletters)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/a:solidFill/a:schemeClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/a:solidFill/a:srgbClr/@val) and not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/a:solidFill/a:schemeClr/@val)">
              <xsl:choose>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
                  <xsl:attribute name ="fo:color">
                    <xsl:value-of select ="translate(concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val),$ucletters,$lcletters)"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
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
                </xsl:when>
              </xsl:choose>
				</xsl:if>
          </xsl:when>-->
        </xsl:choose>
        <!--<xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:defRPr/a:latin/@typeface)">
					<xsl:value-of  select ="$DefFont"/>
				</xsl:if>-->
      </xsl:when>
      <xsl:when test ="$levelForDefFont='6'" >
        <xsl:choose>
          <xsl:when test="$AttrType='Fontname'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/a:latin/@typeface">
              <xsl:attribute name ="fo:font-family">
                <xsl:value-of  select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontsize'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@sz div 100,'#.##')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontweight'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@b='1'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'bold'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@b='0'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'normal'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Underline'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@u">
              <xsl:for-each select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr">
                <xsl:call-template name="tmpUnderLine"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='italic'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@i='1'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'italic'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@i='0'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'none'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <!--<xsl:when test="$AttrType='Fontcolor'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="translate(concat('#',parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/a:solidFill/a:srgbClr/@val),$ucletters,$lcletters)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/a:solidFill/a:schemeClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
				</xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/a:solidFill/a:srgbClr/@val) and not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/a:solidFill/a:schemeClr/@val)">
              <xsl:choose>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
                  <xsl:attribute name ="fo:color">
                    <xsl:value-of select ="translate(concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val),$ucletters,$lcletters)"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
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
			</xsl:when>
              </xsl:choose>
				</xsl:if>
          </xsl:when>-->
        </xsl:choose>
        <!--<xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:defRPr/a:latin/@typeface)">
					<xsl:value-of  select ="$DefFont"/>
				</xsl:if>-->
      </xsl:when>
      <xsl:when test ="$levelForDefFont='7'" >
        <xsl:choose>
          <xsl:when test="$AttrType='Fontname'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/a:latin/@typeface">
              <xsl:attribute name ="fo:font-family">
                <xsl:value-of  select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontsize'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@sz div 100,'#.##')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontweight'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@b='1'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'bold'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@b='0'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'normal'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Underline'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@u">
              <xsl:for-each select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr">
                <xsl:call-template name="tmpUnderLine"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='italic'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@i='1'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'italic'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@i='0'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'none'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <!--<xsl:when test="$AttrType='Fontcolor'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="translate(concat('#',parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/a:solidFill/a:srgbClr/@val),$ucletters,$lcletters)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/a:solidFill/a:schemeClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/a:solidFill/a:srgbClr/@val) and not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/a:solidFill/a:schemeClr/@val)">
              <xsl:choose>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
                  <xsl:attribute name ="fo:color">
                    <xsl:value-of select ="translate(concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val),$ucletters,$lcletters)"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
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
                </xsl:when>
              </xsl:choose>
				</xsl:if>
          </xsl:when>-->
        </xsl:choose>
        <!--<xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:defRPr/a:latin/@typeface)">
					<xsl:value-of  select ="$DefFont"/>
				</xsl:if>-->
      </xsl:when>
      <xsl:when test ="$levelForDefFont='8'" >
        <xsl:choose>
          <xsl:when test="$AttrType='Fontname'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/a:latin/@typeface">
              <xsl:attribute name ="fo:font-family">
                <xsl:value-of  select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontsize'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@sz div 100,'#.##')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontweight'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@b='1'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'bold'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@b='0'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'normal'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Underline'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@u">
              <xsl:for-each select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr">
                <xsl:call-template name="tmpUnderLine"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='italic'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@i='1'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'italic'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@i='0'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'none'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <!--<xsl:when test="$AttrType='Fontcolor'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="translate(concat('#',parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/a:solidFill/a:srgbClr/@val),$ucletters,$lcletters)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/a:solidFill/a:schemeClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/a:solidFill/a:srgbClr/@val) and not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/a:solidFill/a:schemeClr/@val)">
              <xsl:choose>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
                  <xsl:attribute name ="fo:color">
                    <xsl:value-of select ="translate(concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val),$ucletters,$lcletters)"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
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
                </xsl:when>
              </xsl:choose>
				</xsl:if>
          </xsl:when>-->
        </xsl:choose>
        <!--<xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:defRPr/a:latin/@typeface)">
					<xsl:value-of  select ="$DefFont"/>
				</xsl:if>-->
      </xsl:when>
      <xsl:when test ="$levelForDefFont='9'" >
        <xsl:choose>
          <xsl:when test="$AttrType='Fontname'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/a:latin/@typeface">
              <xsl:attribute name ="fo:font-family">
                <xsl:value-of  select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontsize'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <xsl:value-of  select ="format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:defRPr/@sz div 100,'#.##')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontweight'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@b='1'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'bold'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@b='0'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'normal'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Underline'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@u">
              <xsl:for-each select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr">
                <xsl:call-template name="tmpUnderLine"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='italic'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@i='1'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'italic'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@i='0'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'none'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <!--<xsl:when test="$AttrType='Fontcolor'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="translate(concat('#',parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/a:solidFill/a:srgbClr/@val),$ucletters,$lcletters)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/a:solidFill/a:schemeClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/a:solidFill/a:srgbClr/@val) and not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/a:solidFill/a:schemeClr/@val)">
              <xsl:choose>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
                  <xsl:attribute name ="fo:color">
                    <xsl:value-of select ="translate(concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val),$ucletters,$lcletters)"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
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
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>-->
        </xsl:choose>
        <!--<xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:defRPr/a:latin/@typeface)">
					<xsl:value-of  select ="$DefFont"/>
				</xsl:if>-->
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!-- added by Vipul to get back color of title, sub title, outline from slide layout-->
  <!-- start-->
  <xsl:template name="GetLayOutforSlide">
    <xsl:param name="SlideName" />
    <xsl:for-each select="document(concat('ppt/slides/_rels/',$SlideName,'.rels'))//rels:Relationships/rels:Relationship/@Target[contains(.,'slideLayouts')]">
      <xsl:value-of select="substring(.,17)" />
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="tmpSPBackColor">
    <!--for getting Title, sub title back ground color-->
    <xsl:choose>
      <xsl:when test="./p:spPr/a:solidFill/a:srgbClr/@val">
        <xsl:value-of select ="concat('#',./p:spPr/a:solidFill/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="./p:spPr/a:solidFill/a:schemeClr">
        <xsl:call-template name ="getColorCode">
          <xsl:with-param name ="color">
            <xsl:value-of select="./p:spPr/a:solidFill/a:schemeClr/@val"/>
          </xsl:with-param>
          <xsl:with-param name ="lumMod">
            <xsl:if test="./p:spPr/a:solidFill/a:schemeClr/a:lumMod">
              <xsl:value-of select="./p:spPr/a:solidFill/a:schemeClr/a:lumMod/@val" />
            </xsl:if>
          </xsl:with-param>
          <xsl:with-param name ="lumOff">
            <xsl:if test="./p:spPr/a:solidFill/a:schemeClr/a:lumOff">
              <xsl:value-of select="./p:spPr/a:solidFill/a:schemeClr/a:lumOff/@val" />
            </xsl:if>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="''"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name ="tmpDefaultFontName">
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
  <xsl:template name="tmpSlideFontFamily">
    <xsl:param name="AttrType"/>
    <xsl:param name="DefFont"/>
    <xsl:param name="fontscale"/>
    <xsl:choose>
      <xsl:when test ="$AttrType='Fontname'">
        <xsl:attribute name ="fo:font-family">
          <xsl:variable name ="typeFaceVal" select ="a:rPr/a:latin/@typeface"/>
          <xsl:for-each select ="a:rPr/a:latin/@typeface">
            <xsl:if test ="$typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt'">
              <xsl:value-of  select ="$DefFont"/>
            </xsl:if>
            <xsl:if test ="not($typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt')">
              <xsl:value-of select ="."/>
            </xsl:if>
          </xsl:for-each>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="$AttrType='Fontweight'">
        <xsl:if test ="a:rPr/@b='1'">
          <xsl:attribute name ="fo:font-weight">
            <xsl:value-of select ="'bold'"/>
          </xsl:attribute>
        </xsl:if >
        <xsl:if test ="a:rPr/@b='0'">
          <xsl:attribute name ="fo:font-weight">
            <xsl:value-of select ="'normal'"/>
          </xsl:attribute>
        </xsl:if >
      </xsl:when>
      <xsl:when test="$AttrType='Fontsize'">
        <xsl:attribute name ="fo:font-size"	>
          <xsl:for-each select ="a:rPr/@sz">
            <!--<xsl:if test="$fontscale!=''">
              <xsl:value-of select ="concat(format-number(round(. * 0.00001),'#.##'), 'pt')"/>
            </xsl:if>-->
            <!--<xsl:if test="$fontscale=''">-->
            <xsl:value-of select ="concat(format-number(. div 100,'#.##'), 'pt')"/>
            <!--</xsl:if>-->
          </xsl:for-each>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="$AttrType='Underline'">
        <xsl:for-each select ="a:rPr">
          <xsl:call-template name="tmpUnderLine"/>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="$AttrType='italic'">
        <xsl:attribute name ="fo:font-style">
          <xsl:if test ="a:rPr/@i='1'">
            <xsl:value-of select ="'italic'"/>
          </xsl:if >
          <xsl:if test ="a:rPr/@i='0'">
            <xsl:value-of select ="'none'"/>
          </xsl:if>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpLayoutFontFamily">
    <xsl:param name="DefFont"/>
    <xsl:param name="AttrType"/>
    <xsl:choose>
      <xsl:when test ="$AttrType='Fontname'">
        <xsl:variable name="varLayout_fontname">
          <xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr//a:latin/@typeface">
            <xsl:variable name ="typeFaceVal" select ="./parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr//a:latin/@typeface"/>
            <xsl:for-each select ="./parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr//a:latin/@typeface">
              <xsl:choose>
                <xsl:when test ="$typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt'">
                  <xsl:value-of  select ="$DefFont"/>
                </xsl:when>
                <xsl:when test ="not($typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt')">
                  <xsl:value-of select ="."/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="''"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:for-each>
          </xsl:if>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="$varLayout_fontname=''"/>
          <xsl:otherwise>
            <xsl:attribute name ="fo:font-family">
              <xsl:value-of select="$varLayout_fontname"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpUnderLine">
    <xsl:if test ="./a:uFill/a:solidFill/a:srgbClr/@val">
      <xsl:attribute name ="style:text-underline-color">
        <xsl:value-of select ="concat('#',./a:uFill/a:solidFill/a:srgbClr/@val)"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="./a:uFill/a:solidFill/a:schemeClr/@val">
      <xsl:attribute name ="style:text-underline-color">
        <xsl:call-template name ="getColorCode">
          <xsl:with-param name ="color">
            <xsl:value-of select="./a:uFill/a:solidFill/a:schemeClr/@val"/>
          </xsl:with-param>
          <xsl:with-param name ="lumMod">
            <xsl:value-of select="./a:uFill/a:solidFill/a:schemeClr/a:lumMod/@val"/>
          </xsl:with-param>
          <xsl:with-param name ="lumOff">
            <xsl:value-of select="./a:uFill/a:solidFill/a:schemeClr/a:lumOff/@val"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:choose >
      <xsl:when test ="./@u='dbl'">
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
      <xsl:when test ="./@u='heavy'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'solid'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'bold'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='dotted'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'dotted'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'auto'"/>
        </xsl:attribute>
      </xsl:when>
      <!-- dottedHeavy-->
      <xsl:when test ="./@u='dottedHeavy'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'dotted'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'bold'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='dash'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'dash'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'auto'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='dashHeavy'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'dash'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'bold'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='dashLong'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'long-dash'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'auto'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='dashLongHeavy'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'long-dash'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'bold'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='dotDash'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'dot-dash'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'auto'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='dotDashHeavy'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'dot-dash'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'bold'"/>
        </xsl:attribute>
      </xsl:when>
      <!-- dot-dot-dash-->
      <xsl:when test ="./@u='dotDotDash'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'dot-dot-dash'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'auto'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='dotDotDashHeavy'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'dot-dot-dash'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'bold'"/>
        </xsl:attribute>
      </xsl:when>
      <!-- Wavy and Heavy-->
      <xsl:when test ="./@u='wavy'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'wave'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'auto'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='wavyHeavy'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'wave'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'bold'"/>
        </xsl:attribute>
      </xsl:when>
      <!-- wavyDbl-->
      <!-- style:text-underline-style="wave" style:text-underline-type="double"-->
      <xsl:when test ="./@u='wavyDbl'">
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
      <xsl:when test ="./@u='sng'">
        <xsl:attribute name ="style:text-underline-type">
          <xsl:value-of select ="'single'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'auto'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='none'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'none'"/>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpCommanTextProperty">
    <xsl:param name="LayoutFileName"/>
    <xsl:param name="spType"/>
    <xsl:param name="index"/>
    <xsl:param name ="DefFont" />
    <xsl:param name ="fontscale"/>
    <xsl:if test ="a:rPr/@sz">
      <xsl:attribute name ="fo:font-size"	>
        <xsl:for-each select ="a:rPr/@sz">
          <xsl:choose>
            <xsl:when test="$fontscale ='100000'">
              <xsl:value-of  select ="concat(format-number(. div 100,'#.##'),'pt')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of  select ="concat(format-number(round((. *($fontscale div 1000) )div 10000),'#.##'),'pt')"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="not(a:rPr/@sz)">
      <xsl:variable name ="var_layoutRelation">
        <xsl:value-of select ="concat('ppt/slideLayouts/_rels/',$LayoutFileName,'.rels')"/>
      </xsl:variable>
      <xsl:variable name="var_slideMasterName">
        <xsl:for-each  select ="document($var_layoutRelation)//node()/@Target[contains(.,'slideMasters')]">
          <xsl:value-of select ="concat('ppt',substring(.,3))"/>
        </xsl:for-each>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fontscale ='100000'">
                    <xsl:value-of  select ="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz *($fontscale div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>

              </xsl:attribute>
            </xsl:if>
            <xsl:if test="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz)">
              <xsl:variable name="var_SMFontSize">
                <xsl:for-each select ="document($var_slideMasterName)//p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/@sz">
                  <xsl:value-of select ="."/>
                </xsl:for-each>
              </xsl:variable>
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fontscale ='100000'">
                    <xsl:value-of  select ="concat(format-number($var_SMFontSize div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round(($var_SMFontSize *($fontscale div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fontscale ='100000'">
                    <xsl:value-of  select ="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz *($fontscale div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>

              </xsl:attribute>
            </xsl:if>
            <xsl:if test="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz)">
              <xsl:variable name="var_SMFontSize">
                <xsl:for-each select ="document($var_slideMasterName)//p:txStyles/p:bodyStyle/a:lvl1pPr/a:defRPr/@sz">
                  <xsl:value-of select ="."/>
                </xsl:for-each>
              </xsl:variable>
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fontscale ='100000'">
                    <xsl:value-of  select ="concat(format-number($var_SMFontSize div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round(($var_SMFontSize *($fontscale div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type=$spType and @idx=$index]">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fontscale ='100000'">
                    <xsl:value-of  select ="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz *($fontscale div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz)">
              <xsl:variable name="var_SMFontSize">
                <xsl:for-each select ="document($var_slideMasterName)//p:txStyles/p:bodyStyle/a:lvl1pPr/a:defRPr/@sz">
                  <xsl:value-of select ="."/>
                </xsl:for-each>
              </xsl:variable>
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fontscale ='100000'">
                    <xsl:value-of  select ="concat(format-number($var_SMFontSize div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round(($var_SMFontSize *($fontscale div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
    <xsl:if test ="a:rPr/a:latin/@typeface">
      <xsl:call-template name="tmpSlideFontFamily">
        <xsl:with-param name="AttrType" select="'Fontname'"/>
        <xsl:with-param name="DefFont" select="$DefFont"/>
      </xsl:call-template>
    </xsl:if>
    <xsl:if test ="not(a:rPr/a:latin/@typeface)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="DefFont" select ="$DefFont"/>
              <xsl:with-param name ="AttrType" select ="'Fontname'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="DefFont" select ="$DefFont"/>
              <xsl:with-param name ="AttrType" select ="'Fontname'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="DefFont" select ="$DefFont"/>
              <xsl:with-param name ="AttrType" select ="'Fontname'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='' and $index=''">
          <xsl:attribute name ="fo:font-family">
            <xsl:variable name ="typeFaceVal" select ="a:rPr/a:latin/@typeface"/>
            <xsl:for-each select ="a:rPr/a:latin/@typeface">
              <xsl:if test ="$typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt'">
                <xsl:value-of select ="$DefFont"/>
              </xsl:if>
              <xsl:if test ="not($typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt')">
                <xsl:value-of select ="."/>
              </xsl:if>
            </xsl:for-each>
            <xsl:if test="not(a:rPr/a:latin/@typeface)">
              <xsl:value-of  select ="$DefFont"/>
            </xsl:if>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type=$spType and @idx=$index]">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="DefFont" select ="$DefFont"/>
              <xsl:with-param name ="AttrType" select ="'Fontname'"/>
            </xsl:call-template>
          </xsl:for-each>
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
    <xsl:if test ="not(a:rPr/@strike)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'stike'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'strike'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'strike'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type=$spType and @idx=$index]">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'strike'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <!-- Kening Property-->
    <xsl:if test ="a:rPr/@kern">
      <xsl:choose >
        <xsl:when test ="a:rPr/@kern = '0'">
          <xsl:attribute name ="style:letter-kerning">
            <xsl:value-of select ="'false'"/>
          </xsl:attribute >
        </xsl:when >
        <xsl:when test ="format-number(a:rPr/@kern,'#.##') &gt; 0">
          <xsl:attribute name ="style:letter-kerning">
            <xsl:value-of select ="'true'"/>
          </xsl:attribute >
        </xsl:when >
      </xsl:choose>
    </xsl:if >
    <xsl:if test ="not(a:rPr/@kern)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'Kerning'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'Kerning'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'Kerning'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type=$spType and @idx=$index]">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'Kerning'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if >
    <!-- Bold Property-->
    <xsl:if test ="a:rPr/@b">
      <xsl:call-template name="tmpSlideFontFamily">
        <xsl:with-param name="AttrType" select="'Fontweight'"/>
      </xsl:call-template>
    </xsl:if >
    <xsl:if test ="not(a:rPr/@b)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'Fontweight'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'Fontweight'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'Fontweight'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type=$spType and @idx=$index]">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'Fontweight'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if >
    <!--UnderLine-->
    <xsl:if test ="a:rPr/@u">
      <xsl:call-template name="tmpSlideFontFamily">
        <xsl:with-param name="AttrType" select="'Underline'"/>
      </xsl:call-template>
    </xsl:if >
    <xsl:if test ="not(a:rPr/@u)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'Underline'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'Underline'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'Underline'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type=$spType and @idx=$index]">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'Underline'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if >
    <!-- Italic-->
    <xsl:if test ="a:rPr/@i">
      <xsl:call-template name="tmpSlideFontFamily">
        <xsl:with-param name="AttrType" select="'italic'"/>
      </xsl:call-template>
    </xsl:if >
    <xsl:if test ="not(a:rPr/@i)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'italic'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'italic'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'italic'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type=$spType and @idx=$index]">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'italic'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if >
    <!-- Character Spacing -->
    <xsl:if test ="a:rPr/@spc">
      <xsl:attribute name ="fo:letter-spacing">
        <xsl:variable name="length" select="a:rPr/@spc" />
        <xsl:value-of select="concat(format-number($length * 2.54 div 7200,'#.###'),'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="not(a:rPr/@spc)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'charspacing'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'charspacing'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'charspacing'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type=$spType and @idx=$index]">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'charspacing'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
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
    <xsl:if test ="not(a:rPr/a:solidFill/a:srgbClr/@val) and not(a:rPr/a:solidFill/a:schemeClr/@val)">
      <xsl:choose>
        <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
          <xsl:attribute name ="fo:color">
            <xsl:value-of select ="translate(concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val),$ucletters,$lcletters)"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
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
        </xsl:when>
        <xsl:otherwise>
          <xsl:choose>
            <xsl:when test="$spType='title'">
              <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
                <xsl:call-template name ="tmpTextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontcolor'"/>
                </xsl:call-template>
              </xsl:for-each>
            </xsl:when>
            <xsl:when test="$spType='subtitle'">
              <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
                <xsl:call-template name ="tmpTextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontcolor'"/>
                </xsl:call-template>
              </xsl:for-each>
            </xsl:when>
            <xsl:when test="$spType='body'">
              <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
                <xsl:call-template name ="tmpTextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontcolor'"/>
                </xsl:call-template>
              </xsl:for-each>
            </xsl:when>
            <!--<xsl:otherwise>
              <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type=$spType and @idx=$index]">
                <xsl:call-template name ="tmpTextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontcolor'"/>
                </xsl:call-template>
              </xsl:for-each>
            </xsl:otherwise>-->
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <!--Shadow fo:text-shadow-->
    <xsl:if test ="a:rPr/a:effectLst/a:outerShdw">
      <xsl:attribute name ="fo:text-shadow">
        <xsl:value-of select ="'1pt 1pt'"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="not(a:rPr/a:effectLst/a:outerShdw)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'Textshadow'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'Textshadow'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'Textshadow'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type=$spType and @idx=$index]">
            <xsl:call-template name ="tmpTextProperty">
              <xsl:with-param name ="AttrType" select ="'Textshadow'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
  </xsl:template>
  <xsl:template name="tmpTextProperty">
    <xsl:param name ="DefFont" />
    <xsl:param name ="AttrType" />
    <xsl:param name ="fs" />
    <xsl:param name="slideMasterName"/>
    <xsl:choose>
      <xsl:when test="$AttrType='Fontname'">
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:latin/@typeface">
          <xsl:attribute name ="fo:font-family">
            <xsl:value-of  select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:latin/@typeface"/>
          </xsl:attribute>
          <xsl:attribute name ="style:font-family-generic"	>
            <xsl:value-of select ="'roman'"/>
          </xsl:attribute>
          <xsl:attribute name ="style:font-pitch"	>
            <xsl:value-of select ="'variable'"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='Fontsize'">
        <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz">
          <xsl:attribute name ="fo:font-size">
            <xsl:choose>
              <xsl:when test="$fs ='100000'">
                <xsl:value-of  select ="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz div 100,'#.##'),'pt')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of  select ="concat(format-number(round((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz *($fs div 1000) )div 10000),'#.##'),'pt')"/>
              </xsl:otherwise>
            </xsl:choose>

          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='Fontweight'">
        <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@b='1'">
          <xsl:attribute name ="fo:font-weight">
            <xsl:value-of  select ="'bold'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@b='0'">
          <xsl:attribute name ="fo:font-weight">
            <xsl:value-of  select ="'normal'"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='Kerning'">
        <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@kern">
          <xsl:choose >
            <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@kern = '0'">
              <xsl:attribute name ="style:letter-kerning">
                <xsl:value-of select ="'false'"/>
              </xsl:attribute >
            </xsl:when >
            <xsl:when test ="format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@kern,'#.##') &gt; 0">
              <xsl:attribute name ="style:letter-kerning">
                <xsl:value-of select ="'true'"/>
              </xsl:attribute >
            </xsl:when >
          </xsl:choose>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='Underline'">
        <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@u">
          <xsl:for-each select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr">
            <xsl:call-template name="tmpUnderLine"/>
          </xsl:for-each>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='italic'">
        <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@i='1'">
          <xsl:attribute name ="fo:font-style">
            <xsl:value-of  select ="'italic'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@i='0'">
          <xsl:attribute name ="fo:font-style">
            <xsl:value-of  select ="'none'"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='charspacing'">
        <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@spc">
          <xsl:attribute name ="fo:letter-spacing">
            <xsl:variable name="length" select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@spc" />
            <xsl:value-of select="concat(format-number($length * 2.54 div 7200,'#.###'),'cm')"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='Fontcolor'">
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
          <xsl:attribute name ="fo:color">
            <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/@val">
          <xsl:choose>
            <xsl:when test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/a:tint/@val='75000'">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select="'#898989'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name ="fo:color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
        <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr/@val) and not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/@val)">
          <xsl:choose>
            <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val)"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
              <xsl:choose>
                <xsl:when test="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr/a:tint/@val='75000'">
                  <xsl:attribute name ="fo:color">
                    <xsl:value-of select="'#898989'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
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
                </xsl:otherwise>

              </xsl:choose>
            </xsl:when>
          </xsl:choose>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='strike'">
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@strike">
          <xsl:attribute name ="style:text-line-through-style">
            <xsl:value-of select ="'solid'"/>
          </xsl:attribute>
          <xsl:choose >
            <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@strike='dblStrike'">
              <xsl:attribute name ="style:text-line-through-type">
                <xsl:value-of select ="'double'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@strike='sngStrike'">
              <xsl:attribute name ="style:text-line-through-type">
                <xsl:value-of select ="'single'"/>
              </xsl:attribute>
            </xsl:when>
          </xsl:choose>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='Textshadow'">
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:effectLst/a:outerShdw">
          <xsl:attribute name ="fo:text-shadow">
            <xsl:value-of select ="'1pt 1pt'"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpCommanOutlineTextProperty">
    <xsl:param name="LayoutFileName"/>
    <xsl:param name="spType"/>
    <xsl:param name="index"/>
    <xsl:param name ="DefFont" />
    <xsl:param name ="fontscale"/>
    <xsl:param name ="level"/>
    <xsl:if test ="a:rPr/@sz">
      <xsl:attribute name ="fo:font-size">
        <xsl:for-each select ="a:rPr/@sz">
          <xsl:choose>
            <xsl:when test="$fontscale ='100000'">
              <xsl:value-of  select ="concat(format-number(. div 100,'#.##'),'pt')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of  select ="concat(format-number(round((. *($fontscale div 1000) )div 10000),'#.##'),'pt')"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="not(a:rPr/@sz)">
      <xsl:variable name ="var_layoutRelation">
        <xsl:value-of select ="concat('ppt/slideLayouts/_rels/',$LayoutFileName,'.rels')"/>
      </xsl:variable>
      <xsl:variable name="var_slideMasterName">
        <xsl:for-each  select ="document($var_layoutRelation)//node()/@Target[contains(.,'slideMasters')]">
          <xsl:value-of select ="concat('ppt',substring(.,3))"/>
        </xsl:for-each>
      </xsl:variable>
      <xsl:if test="$spType='outline'">
        <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))//p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
          <xsl:if test="not(./@type) and ./@idx=$index">
            <xsl:choose>
              <xsl:when test="$level='0'">
                <xsl:call-template name ="tmpOutlineLvl1_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontsize'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                  <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
                  <xsl:with-param name ="fs" select ="$fontscale"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='1'">
                <xsl:call-template name ="tmpOutlineLvl2_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontsize'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                  <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
                  <xsl:with-param name ="fs" select ="$fontscale"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='2'">
                <xsl:call-template name ="tmpOutlineLvl3_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontsize'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                  <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
                  <xsl:with-param name ="fs" select ="$fontscale"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='3'">
                <xsl:call-template name ="tmpOutlineLvl4_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontsize'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                  <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
                  <xsl:with-param name ="fs" select ="$fontscale"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='4'">
                <xsl:call-template name ="tmpOutlineLvl5_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontsize'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                  <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
                  <xsl:with-param name ="fs" select ="$fontscale"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='5'">
                <xsl:call-template name ="tmpOutlineLvl6_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontsize'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                  <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
                  <xsl:with-param name ="fs" select ="$fontscale"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='6'">
                <xsl:call-template name ="tmpOutlineLvl7_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontsize'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                  <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
                  <xsl:with-param name ="fs" select ="$fontscale"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='7'">
                <xsl:call-template name ="tmpOutlineLvl8_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontsize'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                  <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
                  <xsl:with-param name ="fs" select ="$fontscale"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='8'">
                <xsl:call-template name ="tmpOutlineLvl9_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontsize'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                  <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
                  <xsl:with-param name ="fs" select ="$fontscale"/>
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="a:rPr/a:latin/@typeface">
      <xsl:call-template name="tmpSlideFontFamily">
        <xsl:with-param name="AttrType" select="'Fontname'"/>
        <xsl:with-param name="DefFont" select="$DefFont"/>
      </xsl:call-template>
    </xsl:if>
    <xsl:if test ="not(a:rPr/a:latin/@typeface)">
      <xsl:if test="$spType='outline'">
        <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))//p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
          <xsl:if test="not(./@type) and ./@idx=$index">
            <xsl:choose>
              <xsl:when test="$level='0'">
                <xsl:call-template name ="tmpOutlineLvl1_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontname'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='1'">
                <xsl:call-template name ="tmpOutlineLvl2_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontname'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='2'">
                <xsl:call-template name ="tmpOutlineLvl3_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontname'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='3'">
                <xsl:call-template name ="tmpOutlineLvl4_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontname'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='4'">
                <xsl:call-template name ="tmpOutlineLvl5_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontname'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='5'">
                <xsl:call-template name ="tmpOutlineLvl6_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontname'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='6'">
                <xsl:call-template name ="tmpOutlineLvl7_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontname'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='7'">
                <xsl:call-template name ="tmpOutlineLvl8_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontname'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='8'">
                <xsl:call-template name ="tmpOutlineLvl9_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontname'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
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
    <xsl:if test ="not(a:rPr/@strike)">
      <xsl:if test="$spType='outline'">
        <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))//p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
          <xsl:if test="not(./@type) and ./@idx=$index">
            <xsl:choose>
              <xsl:when test="$level='0'">
                <xsl:call-template name ="tmpOutlineLvl1_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'strike'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='1'">
                <xsl:call-template name ="tmpOutlineLvl2_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'strike'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='2'">
                <xsl:call-template name ="tmpOutlineLvl3_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'strike'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='3'">
                <xsl:call-template name ="tmpOutlineLvl4_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'strike'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='4'">
                <xsl:call-template name ="tmpOutlineLvl5_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'strike'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='5'">
                <xsl:call-template name ="tmpOutlineLvl6_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'strike'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='6'">
                <xsl:call-template name ="tmpOutlineLvl7_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'strike'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='7'">
                <xsl:call-template name ="tmpOutlineLvl8_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'strike'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='8'">
                <xsl:call-template name ="tmpOutlineLvl9_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'strike'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
    <!-- Kening Property-->
    <xsl:if test ="a:rPr/@kern">
      <xsl:choose >
        <xsl:when test ="a:rPr/@kern = '0'">
          <xsl:attribute name ="style:letter-kerning">
            <xsl:value-of select ="'false'"/>
          </xsl:attribute >
        </xsl:when >
        <xsl:when test ="format-number(a:rPr/@kern,'#.##') &gt; 0">
          <xsl:attribute name ="style:letter-kerning">
            <xsl:value-of select ="'true'"/>
          </xsl:attribute >
        </xsl:when >
      </xsl:choose>
    </xsl:if >
    <xsl:if test ="not(a:rPr/@kern)">
      <xsl:if test="$spType='outline'">
        <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))//p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
          <xsl:if test="not(./@type) and ./@idx=$index">
            <xsl:choose>
              <xsl:when test="$level='0'">
                <xsl:call-template name ="tmpOutlineLvl1_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Kerning'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='1'">
                <xsl:call-template name ="tmpOutlineLvl2_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Kerning'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='2'">
                <xsl:call-template name ="tmpOutlineLvl3_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Kerning'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='3'">
                <xsl:call-template name ="tmpOutlineLvl4_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Kerning'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='4'">
                <xsl:call-template name ="tmpOutlineLvl5_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Kerning'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='5'">
                <xsl:call-template name ="tmpOutlineLvl6_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Kerning'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='6'">
                <xsl:call-template name ="tmpOutlineLvl7_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Kerning'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='7'">
                <xsl:call-template name ="tmpOutlineLvl8_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Kerning'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='8'">
                <xsl:call-template name ="tmpOutlineLvl9_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Kerning'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:if >
    <!-- Bold Property-->
    <xsl:if test ="a:rPr/@b">
      <xsl:call-template name="tmpSlideFontFamily">
        <xsl:with-param name="AttrType" select="'Fontweight'"/>
      </xsl:call-template>
    </xsl:if >
    <xsl:if test ="not(a:rPr/@b)">
      <xsl:if test="$spType='outline'">
        <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))//p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
          <xsl:if test="not(./@type) and ./@idx=$index">
            <xsl:choose>
              <xsl:when test="$level='0'">
                <xsl:call-template name ="tmpOutlineLvl1_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontweight'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='1'">
                <xsl:call-template name ="tmpOutlineLvl2_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontweight'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='2'">
                <xsl:call-template name ="tmpOutlineLvl3_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontweight'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='3'">
                <xsl:call-template name ="tmpOutlineLvl4_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontweight'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='4'">
                <xsl:call-template name ="tmpOutlineLvl5_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontweight'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='5'">
                <xsl:call-template name ="tmpOutlineLvl6_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontweight'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='6'">
                <xsl:call-template name ="tmpOutlineLvl7_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontweight'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='7'">
                <xsl:call-template name ="tmpOutlineLvl8_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontweight'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='8'">
                <xsl:call-template name ="tmpOutlineLvl9_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Fontweight'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:if >
    <!--UnderLine-->
    <xsl:if test ="a:rPr/@u">
      <xsl:call-template name="tmpSlideFontFamily">
        <xsl:with-param name="AttrType" select="'Underline'"/>
      </xsl:call-template>
    </xsl:if >
    <xsl:if test ="not(a:rPr/@u)">
      <xsl:if test="$spType='outline'">
        <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))//p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
          <xsl:if test="not(./@type) and ./@idx=$index">
            <xsl:choose>
              <xsl:when test="$level='0'">
                <xsl:call-template name ="tmpOutlineLvl1_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Underline'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='1'">
                <xsl:call-template name ="tmpOutlineLvl2_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Underline'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='2'">
                <xsl:call-template name ="tmpOutlineLvl3_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Underline'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='3'">
                <xsl:call-template name ="tmpOutlineLvl4_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Underline'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='4'">
                <xsl:call-template name ="tmpOutlineLvl5_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Underline'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='5'">
                <xsl:call-template name ="tmpOutlineLvl6_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Underline'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='6'">
                <xsl:call-template name ="tmpOutlineLvl7_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Underline'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='7'">
                <xsl:call-template name ="tmpOutlineLvl8_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Underline'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='8'">
                <xsl:call-template name ="tmpOutlineLvl9_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Underline'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:if >
    <!-- Italic-->
    <xsl:if test ="a:rPr/@i">
      <xsl:call-template name="tmpSlideFontFamily">
        <xsl:with-param name="AttrType" select="'italic'"/>
      </xsl:call-template>
    </xsl:if >
    <xsl:if test ="not(a:rPr/@i)">
      <xsl:if test="$spType='outline'">
        <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))//p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
          <xsl:if test="not(./@type) and ./@idx=$index">
            <xsl:choose>
              <xsl:when test="$level='0'">
                <xsl:call-template name ="tmpOutlineLvl1_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'italic'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='1'">
                <xsl:call-template name ="tmpOutlineLvl2_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'italic'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='2'">
                <xsl:call-template name ="tmpOutlineLvl3_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'italic'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='3'">
                <xsl:call-template name ="tmpOutlineLvl4_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'italic'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='4'">
                <xsl:call-template name ="tmpOutlineLvl5_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'italic'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='5'">
                <xsl:call-template name ="tmpOutlineLvl6_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'italic'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='6'">
                <xsl:call-template name ="tmpOutlineLvl7_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'italic'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='7'">
                <xsl:call-template name ="tmpOutlineLvl8_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'italic'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='8'">
                <xsl:call-template name ="tmpOutlineLvl9_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'italic'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:if >
    <!-- Character Spacing -->
    <xsl:if test ="a:rPr/@spc">
      <xsl:attribute name ="fo:letter-spacing">
        <xsl:variable name="length" select="a:rPr/@spc" />
        <xsl:value-of select="concat(format-number($length * 2.54 div 7200,'#.###'),'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="not(a:rPr/@spc)">
      <xsl:if test="$spType='outline'">
        <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))//p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
          <xsl:if test="not(./@type) and ./@idx=$index">
            <xsl:choose>
              <xsl:when test="$level='0'">
                <xsl:call-template name ="tmpOutlineLvl1_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'charspacing'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='1'">
                <xsl:call-template name ="tmpOutlineLvl2_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'charspacing'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='2'">
                <xsl:call-template name ="tmpOutlineLvl3_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'charspacing'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='3'">
                <xsl:call-template name ="tmpOutlineLvl4_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'charspacing'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='4'">
                <xsl:call-template name ="tmpOutlineLvl5_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'charspacing'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='5'">
                <xsl:call-template name ="tmpOutlineLvl6_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'charspacing'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='6'">
                <xsl:call-template name ="tmpOutlineLvl7_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'charspacing'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='7'">
                <xsl:call-template name ="tmpOutlineLvl8_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'charspacing'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='8'">
                <xsl:call-template name ="tmpOutlineLvl9_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'charspacing'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
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
    <xsl:if test ="not(a:rPr/a:solidFill/a:srgbClr/@val) and not(a:rPr/a:solidFill/a:schemeClr/@val)">
      <xsl:choose>
        <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
          <xsl:attribute name ="fo:color">
            <xsl:value-of select ="translate(concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val),$ucletters,$lcletters)"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
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
        </xsl:when>
        <xsl:otherwise>
          <xsl:if test="$spType='outline'">
            <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))//p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
              <xsl:if test="not(./@type) and ./@idx=$index">
                <xsl:choose>
                  <xsl:when test="$level='0'">
                    <xsl:call-template name ="tmpOutlineLvl1_TextProperty">
                      <xsl:with-param name ="AttrType" select ="'Fontcolor'"/>
                      <xsl:with-param name ="level" select ="$level+1"/>
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:when test="$level='1'">
                    <xsl:call-template name ="tmpOutlineLvl2_TextProperty">
                      <xsl:with-param name ="AttrType" select ="'Fontcolor'"/>
                      <xsl:with-param name ="level" select ="$level+1"/>
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:when test="$level='2'">
                    <xsl:call-template name ="tmpOutlineLvl3_TextProperty">
                      <xsl:with-param name ="AttrType" select ="'Fontcolor'"/>
                      <xsl:with-param name ="level" select ="$level+1"/>
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:when test="$level='3'">
                    <xsl:call-template name ="tmpOutlineLvl4_TextProperty">
                      <xsl:with-param name ="AttrType" select ="'Fontcolor'"/>
                      <xsl:with-param name ="level" select ="$level+1"/>
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:when test="$level='4'">
                    <xsl:call-template name ="tmpOutlineLvl5_TextProperty">
                      <xsl:with-param name ="AttrType" select ="'Fontcolor'"/>
                      <xsl:with-param name ="level" select ="$level+1"/>
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:when test="$level='5'">
                    <xsl:call-template name ="tmpOutlineLvl6_TextProperty">
                      <xsl:with-param name ="AttrType" select ="'Fontcolor'"/>
                      <xsl:with-param name ="level" select ="$level+1"/>
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:when test="$level='6'">
                    <xsl:call-template name ="tmpOutlineLvl7_TextProperty">
                      <xsl:with-param name ="AttrType" select ="'Fontcolor'"/>
                      <xsl:with-param name ="level" select ="$level+1"/>
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:when test="$level='7'">
                    <xsl:call-template name ="tmpOutlineLvl8_TextProperty">
                      <xsl:with-param name ="AttrType" select ="'Fontcolor'"/>
                      <xsl:with-param name ="level" select ="$level+1"/>
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:when test="$level='8'">
                    <xsl:call-template name ="tmpOutlineLvl9_TextProperty">
                      <xsl:with-param name ="AttrType" select ="'Fontcolor'"/>
                      <xsl:with-param name ="level" select ="$level+1"/>
                    </xsl:call-template>
                  </xsl:when>
                </xsl:choose>
              </xsl:if>
            </xsl:for-each>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <xsl:if test ="a:rPr/a:effectLst/a:outerShdw">
      <xsl:attribute name ="fo:text-shadow">
        <xsl:value-of select ="'1pt 1pt'"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="not(a:rPr/a:effectLst/a:outerShdw)">
      <xsl:if test="$spType='outline'">
        <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))//p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
          <xsl:if test="not(./@type) and ./@idx=$index">
            <xsl:choose>
              <xsl:when test="$level='0'">
                <xsl:call-template name ="tmpOutlineLvl1_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Textshadow'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='1'">
                <xsl:call-template name ="tmpOutlineLvl2_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Textshadow'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='2'">
                <xsl:call-template name ="tmpOutlineLvl3_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Textshadow'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='3'">
                <xsl:call-template name ="tmpOutlineLvl4_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Textshadow'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='4'">
                <xsl:call-template name ="tmpOutlineLvl5_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Textshadow'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='5'">
                <xsl:call-template name ="tmpOutlineLvl6_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Textshadow'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='6'">
                <xsl:call-template name ="tmpOutlineLvl7_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Textshadow'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='7'">
                <xsl:call-template name ="tmpOutlineLvl8_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Textshadow'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='8'">
                <xsl:call-template name ="tmpOutlineLvl9_TextProperty">
                  <xsl:with-param name ="AttrType" select ="'Textshadow'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
  </xsl:template>
  <xsl:template name="tmpOutlineLvl1_TextProperty">
    <xsl:param name ="DefFont" />
    <xsl:param name ="AttrType" />
    <xsl:param name ="fs" />
    <xsl:param name="slideMasterName"/>
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test ="$level='1'">
        <xsl:choose>
          <xsl:when test="$AttrType='Fontname'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:latin/@typeface">
              <xsl:attribute name ="fo:font-family">
                <xsl:value-of  select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name ="style:font-family-generic"	>
                <xsl:value-of select ="'roman'"/>
              </xsl:attribute>
              <xsl:attribute name ="style:font-pitch"	>
                <xsl:value-of select ="'variable'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontsize'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fs ='100000'">
                    <xsl:value-of  select ="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz *($fs div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@sz)">
              <xsl:variable name="var_SMFontSize">
                <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl1pPr/a:defRPr/@sz">
                  <xsl:value-of select ="."/>
                </xsl:for-each>
              </xsl:variable>
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fs ='100000'">
                    <xsl:value-of  select ="concat(format-number($var_SMFontSize div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round(($var_SMFontSize *($fs div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontweight'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@b='1'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'bold'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@b='0'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'normal'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Kerning'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@kern">
              <xsl:choose >
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@kern = '0'">
                  <xsl:attribute name ="style:letter-kerning">
                    <xsl:value-of select ="'false'"/>
                  </xsl:attribute >
                </xsl:when >
                <xsl:when test ="format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@kern,'#.##') &gt; 0">
                  <xsl:attribute name ="style:letter-kerning">
                    <xsl:value-of select ="'true'"/>
                  </xsl:attribute >
                </xsl:when >
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Underline'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@u">
              <xsl:for-each select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr">
                <xsl:call-template name="tmpUnderLine"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='italic'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@i='1'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'italic'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@i='0'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'none'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='charspacing'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@spc">
              <xsl:attribute name ="fo:letter-spacing">
                <xsl:variable name="length" select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@spc" />
                <xsl:value-of select="concat(format-number($length * 2.54 div 7200,'#.###'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontcolor'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr/@val) and not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/@val)">
              <xsl:choose>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
                  <xsl:attribute name ="fo:color">
                    <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val)"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
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
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='strike'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@strike">
              <xsl:attribute name ="style:text-line-through-style">
                <xsl:value-of select ="'solid'"/>
              </xsl:attribute>
              <xsl:choose >
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@strike='dblStrike'">
                  <xsl:attribute name ="style:text-line-through-type">
                    <xsl:value-of select ="'double'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@strike='sngStrike'">
                  <xsl:attribute name ="style:text-line-through-type">
                    <xsl:value-of select ="'single'"/>
                  </xsl:attribute>
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Textshadow'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:effectLst/a:outerShdw">
              <xsl:attribute name ="fo:text-shadow">
                <xsl:value-of select ="'1pt 1pt'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpOutlineLvl3_TextProperty">
    <xsl:param name ="DefFont" />
    <xsl:param name ="AttrType" />
    <xsl:param name ="fs" />
    <xsl:param name="slideMasterName"/>
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test ="$level='3'">
        <xsl:choose>
          <xsl:when test="$AttrType='Fontname'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/a:latin/@typeface">
              <xsl:attribute name ="fo:font-family">
                <xsl:value-of  select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name ="style:font-family-generic">
                <xsl:value-of select ="'roman'"/>
              </xsl:attribute>
              <xsl:attribute name ="style:font-pitch">
                <xsl:value-of select ="'variable'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontsize'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fs ='100000'">
                    <xsl:value-of  select ="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@sz div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@sz *($fs div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@sz)">
              <xsl:variable name="var_SMFontSize">
                <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl3pPr/a:defRPr/@sz">
                  <xsl:value-of select ="."/>
                </xsl:for-each>
              </xsl:variable>
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fs ='100000'">
                    <xsl:value-of  select ="concat(format-number($var_SMFontSize div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round(($var_SMFontSize *($fs div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontweight'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@b='1'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'bold'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@b='0'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'normal'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Kerning'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@kern">
              <xsl:choose >
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@kern = '0'">
                  <xsl:attribute name ="style:letter-kerning">
                    <xsl:value-of select ="'false'"/>
                  </xsl:attribute >
                </xsl:when >
                <xsl:when test ="format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@kern,'#.##') &gt; 0">
                  <xsl:attribute name ="style:letter-kerning">
                    <xsl:value-of select ="'true'"/>
                  </xsl:attribute >
                </xsl:when >
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Underline'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@u">
              <xsl:for-each select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr">
                <xsl:call-template name="tmpUnderLine"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='italic'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@i='1'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'italic'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@i='0'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'none'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='charspacing'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@spc">
              <xsl:attribute name ="fo:letter-spacing">
                <xsl:variable name="length" select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@spc" />
                <xsl:value-of select="concat(format-number($length * 2.54 div 7200,'#.###'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontcolor'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/a:solidFill/a:schemeClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/a:solidFill/a:srgbClr/@val) and not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/a:solidFill/a:schemeClr/@val)">
              <xsl:choose>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
                  <xsl:attribute name ="fo:color">
                    <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val)"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
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
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='strike'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@strike">
              <xsl:attribute name ="style:text-line-through-style">
                <xsl:value-of select ="'solid'"/>
              </xsl:attribute>
              <xsl:choose >
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@strike='dblStrike'">
                  <xsl:attribute name ="style:text-line-through-type">
                    <xsl:value-of select ="'double'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:defRPr/@strike='sngStrike'">
                  <xsl:attribute name ="style:text-line-through-type">
                    <xsl:value-of select ="'single'"/>
                  </xsl:attribute>
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Textshadow'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:effectLst/a:outerShdw">
              <xsl:attribute name ="fo:text-shadow">
                <xsl:value-of select ="'1pt 1pt'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpOutlineLvl4_TextProperty">
    <xsl:param name ="DefFont" />
    <xsl:param name ="AttrType" />
    <xsl:param name ="fs" />
    <xsl:param name="slideMasterName"/>
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test ="$level='4'">
        <xsl:choose>
          <xsl:when test="$AttrType='Fontname'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/a:latin/@typeface">
              <xsl:attribute name ="fo:font-family">
                <xsl:value-of  select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name ="style:font-family-generic">
                <xsl:value-of select ="'roman'"/>
              </xsl:attribute>
              <xsl:attribute name ="style:font-pitch">
                <xsl:value-of select ="'variable'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontsize'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fs ='100000'">
                    <xsl:value-of  select ="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@sz div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@sz *($fs div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@sz)">
              <xsl:variable name="var_SMFontSize">
                <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl4pPr/a:defRPr/@sz">
                  <xsl:value-of select ="."/>
                </xsl:for-each>
              </xsl:variable>
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fs ='100000'">
                    <xsl:value-of  select ="concat(format-number($var_SMFontSize div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round(($var_SMFontSize *($fs div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontweight'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@b='1'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'bold'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@b='0'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'normal'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Kerning'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@kern">
              <xsl:choose >
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@kern = '0'">
                  <xsl:attribute name ="style:letter-kerning">
                    <xsl:value-of select ="'false'"/>
                  </xsl:attribute >
                </xsl:when >
                <xsl:when test ="format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@kern,'#.##') &gt; 0">
                  <xsl:attribute name ="style:letter-kerning">
                    <xsl:value-of select ="'true'"/>
                  </xsl:attribute >
                </xsl:when >
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Underline'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@u">
              <xsl:for-each select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr">
                <xsl:call-template name="tmpUnderLine"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='italic'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@i='1'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'italic'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@i='0'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'none'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='charspacing'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@spc">
              <xsl:attribute name ="fo:letter-spacing">
                <xsl:variable name="length" select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@spc" />
                <xsl:value-of select="concat(format-number($length * 2.54 div 7200,'#.###'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontcolor'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/a:solidFill/a:schemeClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/a:solidFill/a:srgbClr/@val) and not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/a:solidFill/a:schemeClr/@val)">
              <xsl:choose>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
                  <xsl:attribute name ="fo:color">
                    <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val)"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
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
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='strike'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@strike">
              <xsl:attribute name ="style:text-line-through-style">
                <xsl:value-of select ="'solid'"/>
              </xsl:attribute>
              <xsl:choose >
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@strike='dblStrike'">
                  <xsl:attribute name ="style:text-line-through-type">
                    <xsl:value-of select ="'double'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:defRPr/@strike='sngStrike'">
                  <xsl:attribute name ="style:text-line-through-type">
                    <xsl:value-of select ="'single'"/>
                  </xsl:attribute>
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Textshadow'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:effectLst/a:outerShdw">
              <xsl:attribute name ="fo:text-shadow">
                <xsl:value-of select ="'1pt 1pt'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpOutlineLvl5_TextProperty">
    <xsl:param name ="DefFont" />
    <xsl:param name ="AttrType" />
    <xsl:param name ="fs" />
    <xsl:param name="slideMasterName"/>
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test ="$level='5'">
        <xsl:choose>
          <xsl:when test="$AttrType='Fontname'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/a:latin/@typeface">
              <xsl:attribute name ="fo:font-family">
                <xsl:value-of  select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name ="style:font-family-generic">
                <xsl:value-of select ="'roman'"/>
              </xsl:attribute>
              <xsl:attribute name ="style:font-pitch">
                <xsl:value-of select ="'variable'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontsize'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fs ='100000'">
                    <xsl:value-of  select ="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@sz div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@sz *($fs div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@sz)">
              <xsl:variable name="var_SMFontSize">
                <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl5pPr/a:defRPr/@sz">
                  <xsl:value-of select ="."/>
                </xsl:for-each>
              </xsl:variable>
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fs ='100000'">
                    <xsl:value-of  select ="concat(format-number($var_SMFontSize div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round(($var_SMFontSize *($fs div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontweight'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@b='1'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'bold'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@b='0'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'normal'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Kerning'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@kern">
              <xsl:choose >
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@kern = '0'">
                  <xsl:attribute name ="style:letter-kerning">
                    <xsl:value-of select ="'false'"/>
                  </xsl:attribute >
                </xsl:when >
                <xsl:when test ="format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@kern,'#.##') &gt; 0">
                  <xsl:attribute name ="style:letter-kerning">
                    <xsl:value-of select ="'true'"/>
                  </xsl:attribute >
                </xsl:when >
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Underline'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@u">
              <xsl:for-each select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr">
                <xsl:call-template name="tmpUnderLine"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='italic'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@i='1'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'italic'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@i='0'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'none'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='charspacing'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@spc">
              <xsl:attribute name ="fo:letter-spacing">
                <xsl:variable name="length" select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@spc" />
                <xsl:value-of select="concat(format-number($length * 2.54 div 7200,'#.###'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontcolor'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/a:solidFill/a:schemeClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/a:solidFill/a:srgbClr/@val) and not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/a:solidFill/a:schemeClr/@val)">
              <xsl:choose>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
                  <xsl:attribute name ="fo:color">
                    <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val)"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
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
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='strike'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@strike">
              <xsl:attribute name ="style:text-line-through-style">
                <xsl:value-of select ="'solid'"/>
              </xsl:attribute>
              <xsl:choose >
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@strike='dblStrike'">
                  <xsl:attribute name ="style:text-line-through-type">
                    <xsl:value-of select ="'double'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:defRPr/@strike='sngStrike'">
                  <xsl:attribute name ="style:text-line-through-type">
                    <xsl:value-of select ="'single'"/>
                  </xsl:attribute>
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Textshadow'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:effectLst/a:outerShdw">
              <xsl:attribute name ="fo:text-shadow">
                <xsl:value-of select ="'1pt 1pt'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpOutlineLvl6_TextProperty">
    <xsl:param name ="DefFont" />
    <xsl:param name ="AttrType" />
    <xsl:param name ="fs" />
    <xsl:param name="slideMasterName"/>
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test ="$level='6'">
        <xsl:choose>
          <xsl:when test="$AttrType='Fontname'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/a:latin/@typeface">
              <xsl:attribute name ="fo:font-family">
                <xsl:value-of  select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name ="style:font-family-generic">
                <xsl:value-of select ="'roman'"/>
              </xsl:attribute>
              <xsl:attribute name ="style:font-pitch">
                <xsl:value-of select ="'variable'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontsize'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fs ='100000'">
                    <xsl:value-of  select ="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@sz div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@sz *($fs div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@sz)">
              <xsl:variable name="var_SMFontSize">
                <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl6pPr/a:defRPr/@sz">
                  <xsl:value-of select ="."/>
                </xsl:for-each>
              </xsl:variable>
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fs ='100000'">
                    <xsl:value-of  select ="concat(format-number($var_SMFontSize div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round(($var_SMFontSize *($fs div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontweight'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@b='1'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'bold'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@b='0'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'normal'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Kerning'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@kern">
              <xsl:choose >
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@kern = '0'">
                  <xsl:attribute name ="style:letter-kerning">
                    <xsl:value-of select ="'false'"/>
                  </xsl:attribute >
                </xsl:when >
                <xsl:when test ="format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@kern,'#.##') &gt; 0">
                  <xsl:attribute name ="style:letter-kerning">
                    <xsl:value-of select ="'true'"/>
                  </xsl:attribute >
                </xsl:when >
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Underline'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@u">
              <xsl:for-each select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr">
                <xsl:call-template name="tmpUnderLine"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='italic'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@i='1'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'italic'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@i='0'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'none'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='charspacing'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@spc">
              <xsl:attribute name ="fo:letter-spacing">
                <xsl:variable name="length" select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@spc" />
                <xsl:value-of select="concat(format-number($length * 2.54 div 7200,'#.###'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontcolor'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/a:solidFill/a:schemeClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/a:solidFill/a:srgbClr/@val) and not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/a:solidFill/a:schemeClr/@val)">
              <xsl:choose>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
                  <xsl:attribute name ="fo:color">
                    <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val)"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
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
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='strike'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@strike">
              <xsl:attribute name ="style:text-line-through-style">
                <xsl:value-of select ="'solid'"/>
              </xsl:attribute>
              <xsl:choose >
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@strike='dblStrike'">
                  <xsl:attribute name ="style:text-line-through-type">
                    <xsl:value-of select ="'double'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:defRPr/@strike='sngStrike'">
                  <xsl:attribute name ="style:text-line-through-type">
                    <xsl:value-of select ="'single'"/>
                  </xsl:attribute>
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Textshadow'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:effectLst/a:outerShdw">
              <xsl:attribute name ="fo:text-shadow">
                <xsl:value-of select ="'1pt 1pt'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpOutlineLvl2_TextProperty">
    <xsl:param name ="DefFont" />
    <xsl:param name ="AttrType" />
    <xsl:param name ="fs" />
    <xsl:param name="slideMasterName"/>
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test ="$level='2'">
        <xsl:choose>
          <xsl:when test="$AttrType='Fontname'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/a:latin/@typeface">
              <xsl:attribute name ="fo:font-family">
                <xsl:value-of  select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name ="style:font-family-generic">
                <xsl:value-of select ="'roman'"/>
              </xsl:attribute>
              <xsl:attribute name ="style:font-pitch">
                <xsl:value-of select ="'variable'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontsize'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fs ='100000'">
                    <xsl:value-of  select ="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@sz div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@sz *($fs div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@sz)">
              <xsl:variable name="var_SMFontSize">
                <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl2pPr/a:defRPr/@sz">
                  <xsl:value-of select ="."/>
                </xsl:for-each>
              </xsl:variable>
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fs ='100000'">
                    <xsl:value-of  select ="concat(format-number($var_SMFontSize div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round(($var_SMFontSize *($fs div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontweight'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@b='1'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'bold'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@b='0'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'normal'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Kerning'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@kern">
              <xsl:choose >
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@kern = '0'">
                  <xsl:attribute name ="style:letter-kerning">
                    <xsl:value-of select ="'false'"/>
                  </xsl:attribute >
                </xsl:when >
                <xsl:when test ="format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@kern,'#.##') &gt; 0">
                  <xsl:attribute name ="style:letter-kerning">
                    <xsl:value-of select ="'true'"/>
                  </xsl:attribute >
                </xsl:when >
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Underline'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@u">
              <xsl:for-each select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr">
                <xsl:call-template name="tmpUnderLine"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='italic'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@i='1'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'italic'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@i='0'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'none'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='charspacing'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@spc">
              <xsl:attribute name ="fo:letter-spacing">
                <xsl:variable name="length" select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@spc" />
                <xsl:value-of select="concat(format-number($length * 2.54 div 7200,'#.###'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontcolor'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/a:solidFill/a:schemeClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/a:solidFill/a:srgbClr/@val) and not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/a:solidFill/a:schemeClr/@val)">
              <xsl:choose>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
                  <xsl:attribute name ="fo:color">
                    <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val)"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
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
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='strike'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@strike">
              <xsl:attribute name ="style:text-line-through-style">
                <xsl:value-of select ="'solid'"/>
              </xsl:attribute>
              <xsl:choose >
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@strike='dblStrike'">
                  <xsl:attribute name ="style:text-line-through-type">
                    <xsl:value-of select ="'double'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:defRPr/@strike='sngStrike'">
                  <xsl:attribute name ="style:text-line-through-type">
                    <xsl:value-of select ="'single'"/>
                  </xsl:attribute>
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Textshadow'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:effectLst/a:outerShdw">
              <xsl:attribute name ="fo:text-shadow">
                <xsl:value-of select ="'1pt 1pt'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpOutlineLvl7_TextProperty">
    <xsl:param name ="DefFont" />
    <xsl:param name ="AttrType" />
    <xsl:param name ="fs" />
    <xsl:param name="slideMasterName"/>
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test ="$level='7'">
        <xsl:choose>
          <xsl:when test="$AttrType='Fontname'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/a:latin/@typeface">
              <xsl:attribute name ="fo:font-family">
                <xsl:value-of  select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name ="style:font-family-generic">
                <xsl:value-of select ="'roman'"/>
              </xsl:attribute>
              <xsl:attribute name ="style:font-pitch">
                <xsl:value-of select ="'variable'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontsize'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fs ='100000'">
                    <xsl:value-of  select ="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@sz div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@sz *($fs div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@sz)">
              <xsl:variable name="var_SMFontSize">
                <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl7pPr/a:defRPr/@sz">
                  <xsl:value-of select ="."/>
                </xsl:for-each>
              </xsl:variable>
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fs ='100000'">
                    <xsl:value-of  select ="concat(format-number($var_SMFontSize div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round(($var_SMFontSize *($fs div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontweight'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@b='1'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'bold'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@b='0'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'normal'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Kerning'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@kern">
              <xsl:choose >
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@kern = '0'">
                  <xsl:attribute name ="style:letter-kerning">
                    <xsl:value-of select ="'false'"/>
                  </xsl:attribute >
                </xsl:when >
                <xsl:when test ="format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@kern,'#.##') &gt; 0">
                  <xsl:attribute name ="style:letter-kerning">
                    <xsl:value-of select ="'true'"/>
                  </xsl:attribute >
                </xsl:when >
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Underline'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@u">
              <xsl:for-each select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr">
                <xsl:call-template name="tmpUnderLine"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='italic'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@i='1'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'italic'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@i='0'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'none'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='charspacing'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@spc">
              <xsl:attribute name ="fo:letter-spacing">
                <xsl:variable name="length" select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@spc" />
                <xsl:value-of select="concat(format-number($length * 2.54 div 7200,'#.###'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontcolor'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/a:solidFill/a:schemeClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/a:solidFill/a:srgbClr/@val) and not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/a:solidFill/a:schemeClr/@val)">
              <xsl:choose>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
                  <xsl:attribute name ="fo:color">
                    <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val)"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
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
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='strike'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@strike">
              <xsl:attribute name ="style:text-line-through-style">
                <xsl:value-of select ="'solid'"/>
              </xsl:attribute>
              <xsl:choose >
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@strike='dblStrike'">
                  <xsl:attribute name ="style:text-line-through-type">
                    <xsl:value-of select ="'double'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:defRPr/@strike='sngStrike'">
                  <xsl:attribute name ="style:text-line-through-type">
                    <xsl:value-of select ="'single'"/>
                  </xsl:attribute>
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Textshadow'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:effectLst/a:outerShdw">
              <xsl:attribute name ="fo:text-shadow">
                <xsl:value-of select ="'1pt 1pt'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:when>

    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpOutlineLvl8_TextProperty">
    <xsl:param name ="DefFont" />
    <xsl:param name ="AttrType" />
    <xsl:param name ="fs" />
    <xsl:param name="slideMasterName"/>
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test ="$level='8'">
        <xsl:choose>
          <xsl:when test="$AttrType='Fontname'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/a:latin/@typeface">
              <xsl:attribute name ="fo:font-family">
                <xsl:value-of  select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name ="style:font-family-generic">
                <xsl:value-of select ="'roman'"/>
              </xsl:attribute>
              <xsl:attribute name ="style:font-pitch">
                <xsl:value-of select ="'variable'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontsize'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fs ='100000'">
                    <xsl:value-of  select ="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@sz div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@sz *($fs div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@sz)">
              <xsl:variable name="var_SMFontSize">
                <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl8pPr/a:defRPr/@sz">
                  <xsl:value-of select ="."/>
                </xsl:for-each>
              </xsl:variable>
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fs ='100000'">
                    <xsl:value-of  select ="concat(format-number($var_SMFontSize div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round(($var_SMFontSize *($fs div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontweight'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@b='1'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'bold'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@b='0'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'normal'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Kerning'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@kern">
              <xsl:choose >
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@kern = '0'">
                  <xsl:attribute name ="style:letter-kerning">
                    <xsl:value-of select ="'false'"/>
                  </xsl:attribute >
                </xsl:when >
                <xsl:when test ="format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@kern,'#.##') &gt; 0">
                  <xsl:attribute name ="style:letter-kerning">
                    <xsl:value-of select ="'true'"/>
                  </xsl:attribute >
                </xsl:when >
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Underline'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@u">
              <xsl:for-each select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr">
                <xsl:call-template name="tmpUnderLine"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='italic'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@i='1'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'italic'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@i='0'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'none'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='charspacing'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@spc">
              <xsl:attribute name ="fo:letter-spacing">
                <xsl:variable name="length" select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@spc" />
                <xsl:value-of select="concat(format-number($length * 2.54 div 7200,'#.###'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontcolor'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/a:solidFill/a:schemeClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/a:solidFill/a:srgbClr/@val) and not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/a:solidFill/a:schemeClr/@val)">
              <xsl:choose>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
                  <xsl:attribute name ="fo:color">
                    <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val)"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
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
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='strike'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@strike">
              <xsl:attribute name ="style:text-line-through-style">
                <xsl:value-of select ="'solid'"/>
              </xsl:attribute>
              <xsl:choose >
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@strike='dblStrike'">
                  <xsl:attribute name ="style:text-line-through-type">
                    <xsl:value-of select ="'double'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:defRPr/@strike='sngStrike'">
                  <xsl:attribute name ="style:text-line-through-type">
                    <xsl:value-of select ="'single'"/>
                  </xsl:attribute>
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Textshadow'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:effectLst/a:outerShdw">
              <xsl:attribute name ="fo:text-shadow">
                <xsl:value-of select ="'1pt 1pt'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpOutlineLvl9_TextProperty">
    <xsl:param name ="DefFont" />
    <xsl:param name ="AttrType" />
    <xsl:param name ="fs" />
    <xsl:param name="slideMasterName"/>
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test ="$level='9'">
        <xsl:choose>
          <xsl:when test="$AttrType='Fontname'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/a:latin/@typeface">
              <xsl:attribute name ="fo:font-family">
                <xsl:value-of  select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name ="style:font-family-generic">
                <xsl:value-of select ="'roman'"/>
              </xsl:attribute>
              <xsl:attribute name ="style:font-pitch">
                <xsl:value-of select ="'variable'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontsize'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@sz">
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fs ='100000'">
                    <xsl:value-of  select ="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@sz div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@sz *($fs div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@sz)">
              <xsl:variable name="var_SMFontSize">
                <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl9pPr/a:defRPr/@sz">
                  <xsl:value-of select ="."/>
                </xsl:for-each>
              </xsl:variable>
              <xsl:attribute name ="fo:font-size">
                <xsl:choose>
                  <xsl:when test="$fs ='100000'">
                    <xsl:value-of  select ="concat(format-number($var_SMFontSize div 100,'#.##'),'pt')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of  select ="concat(format-number(round(($var_SMFontSize *($fs div 1000) )div 10000),'#.##'),'pt')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontweight'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@b='1'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'bold'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@b='0'">
              <xsl:attribute name ="fo:font-weight">
                <xsl:value-of  select ="'normal'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Kerning'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@kern">
              <xsl:choose >
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@kern = '0'">
                  <xsl:attribute name ="style:letter-kerning">
                    <xsl:value-of select ="'false'"/>
                  </xsl:attribute >
                </xsl:when >
                <xsl:when test ="format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@kern,'#.##') &gt; 0">
                  <xsl:attribute name ="style:letter-kerning">
                    <xsl:value-of select ="'true'"/>
                  </xsl:attribute >
                </xsl:when >
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Underline'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@u">
              <xsl:for-each select ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr">
                <xsl:call-template name="tmpUnderLine"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='italic'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@i='1'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'italic'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@i='0'">
              <xsl:attribute name ="fo:font-style">
                <xsl:value-of  select ="'none'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='charspacing'">
            <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@spc">
              <xsl:attribute name ="fo:letter-spacing">
                <xsl:variable name="length" select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@spc" />
                <xsl:value-of select="concat(format-number($length * 2.54 div 7200,'#.###'),'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Fontcolor'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/a:solidFill/a:schemeClr/@val">
              <xsl:attribute name ="fo:color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/a:solidFill/a:srgbClr/@val) and not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/a:solidFill/a:schemeClr/@val)">
              <xsl:choose>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
                  <xsl:attribute name ="fo:color">
                    <xsl:value-of select ="concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val)"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
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
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='strike'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@strike">
              <xsl:attribute name ="style:text-line-through-style">
                <xsl:value-of select ="'solid'"/>
              </xsl:attribute>
              <xsl:choose >
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@strike='dblStrike'">
                  <xsl:attribute name ="style:text-line-through-type">
                    <xsl:value-of select ="'double'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:defRPr/@strike='sngStrike'">
                  <xsl:attribute name ="style:text-line-through-type">
                    <xsl:value-of select ="'single'"/>
                  </xsl:attribute>
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Textshadow'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/a:effectLst/a:outerShdw">
              <xsl:attribute name ="fo:text-shadow">
                <xsl:value-of select ="'1pt 1pt'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpCommanParaProperty">
    <xsl:param name="LayoutFileName"/>
    <xsl:param name="spType"/>
    <xsl:param name="index"/>
    <xsl:param name ="DefFont" />
    <xsl:param name ="lnSpcReduction"/>
    <xsl:variable name ="var_layoutRelation">
      <xsl:value-of select ="concat('ppt/slideLayouts/_rels/',$LayoutFileName,'.rels')"/>
    </xsl:variable>
    <xsl:variable name="var_slideMasterName">
      <xsl:for-each  select ="document($var_layoutRelation)//node()/@Target[contains(.,'slideMasters')]">
        <xsl:value-of select ="concat('ppt',substring(.,3))"/>
      </xsl:for-each>
    </xsl:variable>
    <xsl:call-template name="tmpSlideParaProp">
      <xsl:with-param name="lnSpcReduction" select="$lnSpcReduction"/>
    </xsl:call-template>
    <xsl:if test ="not(a:pPr/@algn)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'TextAlignment'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'TextAlignment'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'TextAlignment'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:if>

    <xsl:if test ="not(a:pPr/@marL)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'marginLeft'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'marginLeft'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'marginLeft'"/>
              <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
            </xsl:call-template>
          </xsl:for-each>
          <xsl:if test="not(document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index])">
            <xsl:for-each select ="document($var_slideMasterName)//p:txStyles/p:bodyStyle/a:lvl1pPr/@marL">
              <xsl:if test=".">
                <xsl:attribute name ="fo:margin-left">
                  <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                </xsl:attribute>
              </xsl:if>
            </xsl:for-each>
          </xsl:if>
        </xsl:when>
      </xsl:choose>
    </xsl:if>

    <xsl:if test ="not(a:pPr/@marR)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'marginRight'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'marginRight'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">

            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'marginRight'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:if>

    <xsl:if test ="not(a:pPr/@indent)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'indent'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'indent'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'indent'"/>
              <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
            </xsl:call-template>
          </xsl:for-each>
          <xsl:if test="not(document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index])">
            <xsl:for-each select ="document($var_slideMasterName)//p:txStyles/p:bodyStyle/a:lvl1pPr/@indent">
              <xsl:if test=".">
                <xsl:attribute name ="fo:text-indent">
                  <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                </xsl:attribute>
              </xsl:if>
            </xsl:for-each>
          </xsl:if>
        </xsl:when>
      </xsl:choose>
    </xsl:if >

    <xsl:if test ="not(a:pPr/a:spcBef/a:spcPts/@val)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'marginTop'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'marginTop'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'marginTop'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:if>

    <xsl:if test ="not(a:pPr/a:spcAft/a:spcPts/@val)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'marginBottom'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'marginBottom'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'marginBottom'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
    <!-- If the line space is in Percentage-->

    <xsl:if test ="not(a:pPr/a:lnSpc/a:spcPct/@val)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'lineSpacing'"/>
              <xsl:with-param name ="lnSpaceRed" select ="$lnSpcReduction"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'lineSpacing'"/>
              <xsl:with-param name ="lnSpaceRed" select ="$lnSpcReduction"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'lineSpacing'"/>
              <xsl:with-param name ="lnSpaceRed" select ="$lnSpcReduction"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:if >
    <!-- If the line space is in Points-->

    <xsl:if test ="not(a:pPr/a:lnSpc/a:spcPts/@val)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'lineSpacing'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'lineSpacing'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'lineSpacing'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
    <xsl:if test ="not(parent::node()/a:bodyPr/@vert='vert')">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'Textdirection'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'Textdirection'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpParaProperty">
              <xsl:with-param name ="AttrType" select ="'Textdirection'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
  </xsl:template>
  <xsl:template name="tmpParaProperty">
    <xsl:param name ="DefFont" />
    <xsl:param name ="AttrType" />
    <xsl:param name ="lnSpaceRed" />
    <xsl:param name="slideMasterName"/>
    <xsl:choose>
      <xsl:when test="$AttrType='TextAlignment'">
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@algn">
          <xsl:attribute name ="fo:text-align">
            <xsl:choose>
              <!-- Center Alignment-->
              <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@algn ='ctr'">
                <xsl:value-of select ="'center'"/>
              </xsl:when>
              <!-- Right Alignment-->
              <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@algn ='r'">
                <xsl:value-of select ="'end'"/>
              </xsl:when>
              <!-- Left Alignment-->
              <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@algn ='l'">
                <xsl:value-of select ="'start'"/>
              </xsl:when>
              <!-- Justified Alignment-->
              <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@algn ='just'">
                <xsl:value-of select ="'justify'"/>
              </xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='marginLeft'">
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@marL">
          <xsl:attribute name ="fo:margin-left">
            <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@marL div 360000, '#.##'), 'cm')"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@marL)">
          <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl1pPr/@marL">
            <xsl:if test=".">
              <xsl:attribute name ="fo:margin-left">
                <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='marginRight'">
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@marR">
          <xsl:attribute name ="fo:margin-right">
            <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@marR div 360000, '#.##'), 'cm')"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='marginTop'">
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:spcBef/a:spcPts/@val">
          <xsl:attribute name ="fo:margin-top">
            <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:spcBef/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='marginBottom'">
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:spcAft/a:spcPts/@val">
          <xsl:attribute name ="fo:margin-bottom">
            <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:spcAft/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='indent'">
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@indent">
          <xsl:attribute name ="fo:text-indent">
            <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@indent div 360000, '#.##'), 'cm')"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@indent)">
          <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl1pPr/@indent">
            <xsl:if test=".">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='lineSpacing'">
        <xsl:choose>
          <xsl:when test="$lnSpaceRed='0'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:lnSpc/a:spcPct/@val">
              <xsl:attribute name="fo:line-height">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:lnSpc/a:spcPct/@val div 1000,'###'), '%')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:otherwise>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:lnSpc/a:spcPct/@val">
              <xsl:attribute name="fo:line-height">
                <xsl:value-of select="concat(format-number((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:lnSpc/a:spcPct/@val - number($lnSpaceRed)) div 1000,'###'), '%')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:otherwise>
        </xsl:choose>
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:lnSpc/a:spcPts/@val">
          <xsl:attribute name="style:line-height-at-least">
            <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:lnSpc/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='Textdirection'">
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@vert='vert'">
          <xsl:attribute name ="style:writing-mode">
            <xsl:value-of select ="'tb-rl'"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpCommanOutlineParaProperty">
    <xsl:param name="LayoutFileName"/>
    <xsl:param name="spType"/>
    <xsl:param name="index"/>
    <xsl:param name ="lnSpcReduction"/>
    <xsl:param name ="level"/>
    <xsl:variable name ="var_layoutRelation">
      <xsl:value-of select ="concat('ppt/slideLayouts/_rels/',$LayoutFileName,'.rels')"/>
    </xsl:variable>
    <xsl:variable name="var_slideMasterName">
      <xsl:for-each  select ="document($var_layoutRelation)//node()/@Target[contains(.,'slideMasters')]">
        <xsl:value-of select ="concat('ppt',substring(.,3))"/>
      </xsl:for-each>
    </xsl:variable>
    <xsl:call-template name="tmpSlideParaProp">
      <xsl:with-param name="lnSpcReduction" select="$lnSpcReduction"/>
    </xsl:call-template>
    <xsl:if test ="not(a:pPr/@algn)">
      <xsl:if test="$spType='outline'">
        <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
          <xsl:if test="not(./@type) and ./@idx=$index">
            <xsl:choose>
              <xsl:when test="$level='0'">
                <xsl:call-template name ="tmpOutlineLvl1_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'TextAlignment'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='1'">
                <xsl:call-template name ="tmpOutlineLvl2_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'TextAlignment'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='2'">
                <xsl:call-template name ="tmpOutlineLvl3_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'TextAlignment'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='3'">
                <xsl:call-template name ="tmpOutlineLvl4_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'TextAlignment'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='4'">
                <xsl:call-template name ="tmpOutlineLvl5_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'TextAlignment'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='5'">
                <xsl:call-template name ="tmpOutlineLvl6_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'TextAlignment'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='6'">
                <xsl:call-template name ="tmpOutlineLvl7_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'TextAlignment'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='7'">
                <xsl:call-template name ="tmpOutlineLvl8_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'TextAlignment'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='8'">
                <xsl:call-template name ="tmpOutlineLvl9_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'TextAlignment'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>

    <xsl:if test ="not(a:pPr/@marL)">
      <xsl:if test="$spType='outline'">
        <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
          <xsl:if test="not(./@type) and ./@idx=$index">
            <xsl:choose>
              <xsl:when test="$level='0'">
                <xsl:call-template name ="tmpOutlineLvl1_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginLeft'"/>
                  <xsl:with-param name ="spType" select ="$spType"/>
                  <xsl:with-param name ="index" select ="$index"/>
                  <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='1'">
                <xsl:call-template name ="tmpOutlineLvl2_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginLeft'"/>
                  <xsl:with-param name ="spType" select ="$spType"/>
                  <xsl:with-param name ="index" select ="$index"/>
                  <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='2'">
                <xsl:call-template name ="tmpOutlineLvl3_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginLeft'"/>
                  <xsl:with-param name ="spType" select ="$spType"/>
                  <xsl:with-param name ="index" select ="$index"/>
                  <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='3'">
                <xsl:call-template name ="tmpOutlineLvl4_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginLeft'"/>
                  <xsl:with-param name ="spType" select ="$spType"/>
                  <xsl:with-param name ="index" select ="$index"/>
                  <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='4'">
                <xsl:call-template name ="tmpOutlineLvl5_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginLeft'"/>
                  <xsl:with-param name ="spType" select ="$spType"/>
                  <xsl:with-param name ="index" select ="$index"/>
                  <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='5'">
                <xsl:call-template name ="tmpOutlineLvl6_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginLeft'"/>
                  <xsl:with-param name ="spType" select ="$spType"/>
                  <xsl:with-param name ="index" select ="$index"/>
                  <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='6'">
                <xsl:call-template name ="tmpOutlineLvl7_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginLeft'"/>
                  <xsl:with-param name ="spType" select ="$spType"/>
                  <xsl:with-param name ="index" select ="$index"/>
                  <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='7'">
                <xsl:call-template name ="tmpOutlineLvl8_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginLeft'"/>
                  <xsl:with-param name ="spType" select ="$spType"/>
                  <xsl:with-param name ="index" select ="$index"/>
                  <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='8'">
                <xsl:call-template name ="tmpOutlineLvl9_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginLeft'"/>
                  <xsl:with-param name ="spType" select ="$spType"/>
                  <xsl:with-param name ="index" select ="$index"/>
                  <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="not(a:pPr/@marR)">
      <xsl:if test="$spType='outline'">
        <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
          <xsl:if test="not(./@type) and ./@idx=$index">
            <xsl:choose>
              <xsl:when test="$level='0'">
                <xsl:call-template name ="tmpOutlineLvl1_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginRight'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='1'">
                <xsl:call-template name ="tmpOutlineLvl2_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginRight'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='2'">
                <xsl:call-template name ="tmpOutlineLvl3_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginRight'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='3'">
                <xsl:call-template name ="tmpOutlineLvl4_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginRight'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='4'">
                <xsl:call-template name ="tmpOutlineLvl5_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginRight'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='5'">
                <xsl:call-template name ="tmpOutlineLvl6_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginRight'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='6'">
                <xsl:call-template name ="tmpOutlineLvl7_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginRight'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='7'">
                <xsl:call-template name ="tmpOutlineLvl8_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginRight'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='8'">
                <xsl:call-template name ="tmpOutlineLvl9_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginRight'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>

    <xsl:if test ="not(a:pPr/@indent)">
      <xsl:if test="$spType='outline'">
        <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
          <xsl:if test="not(./@type) and ./@idx=$index">
            <xsl:choose>
              <xsl:when test="$level='0'">
                <xsl:call-template name ="tmpOutlineLvl1_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'indent'"/>
                  <xsl:with-param name ="slideMasterName" select ="$var_slideMasterName"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='1'">
                <xsl:call-template name ="tmpOutlineLvl2_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'indent'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='2'">
                <xsl:call-template name ="tmpOutlineLvl3_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'indent'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='3'">
                <xsl:call-template name ="tmpOutlineLvl4_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'indent'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='4'">
                <xsl:call-template name ="tmpOutlineLvl5_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'indent'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='5'">
                <xsl:call-template name ="tmpOutlineLvl6_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'indent'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='6'">
                <xsl:call-template name ="tmpOutlineLvl7_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'indent'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='7'">
                <xsl:call-template name ="tmpOutlineLvl8_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'indent'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='8'">
                <xsl:call-template name ="tmpOutlineLvl9_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'indent'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>

    <xsl:if test ="not(a:pPr/a:spcBef/a:spcPts/@val)">
      <xsl:if test="$spType='outline'">
        <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
          <xsl:if test="not(./@type) and ./@idx=$index">
            <xsl:choose>
              <xsl:when test="$level='0'">
                <xsl:call-template name ="tmpOutlineLvl1_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginTop'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='1'">
                <xsl:call-template name ="tmpOutlineLvl2_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginTop'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='2'">
                <xsl:call-template name ="tmpOutlineLvl3_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginTop'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='3'">
                <xsl:call-template name ="tmpOutlineLvl4_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginTop'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='4'">
                <xsl:call-template name ="tmpOutlineLvl5_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginTop'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='5'">
                <xsl:call-template name ="tmpOutlineLvl6_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginTop'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='6'">
                <xsl:call-template name ="tmpOutlineLvl7_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginTop'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='7'">
                <xsl:call-template name ="tmpOutlineLvl8_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginTop'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='8'">
                <xsl:call-template name ="tmpOutlineLvl9_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginTop'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>

    <xsl:if test ="not(a:pPr/a:spcAft/a:spcPts/@val)">
      <xsl:if test="$spType='outline'">
        <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
          <xsl:if test="not(./@type) and ./@idx=$index">
            <xsl:choose>
              <xsl:when test="$level='0'">
                <xsl:call-template name ="tmpOutlineLvl1_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginBottom'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='1'">
                <xsl:call-template name ="tmpOutlineLvl2_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginBottom'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='2'">
                <xsl:call-template name ="tmpOutlineLvl3_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginBottom'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='3'">
                <xsl:call-template name ="tmpOutlineLvl4_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginBottom'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='4'">
                <xsl:call-template name ="tmpOutlineLvl5_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginBottom'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='5'">
                <xsl:call-template name ="tmpOutlineLvl6_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginBottom'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='6'">
                <xsl:call-template name ="tmpOutlineLvl7_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginBottom'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='7'">
                <xsl:call-template name ="tmpOutlineLvl8_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginBottom'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='8'">
                <xsl:call-template name ="tmpOutlineLvl9_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'marginBottom'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
    <!-- If the line space is in Percentage-->

    <xsl:if test ="not(a:pPr/a:lnSpc/a:spcPct/@val) or not(a:pPr/a:lnSpc/a:spcPts/@val)">
      <xsl:if test="$spType='outline'">
        <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
          <xsl:if test="not(./@type) and ./@idx=$index">
            <xsl:choose>
              <xsl:when test="$level='0'">
                <xsl:call-template name ="tmpOutlineLvl1_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'lineSpacing'"/>
                  <xsl:with-param name ="lnSpaceRed" select ="$lnSpcReduction"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='1'">
                <xsl:call-template name ="tmpOutlineLvl2_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'lineSpacing'"/>
                  <xsl:with-param name ="lnSpaceRed" select ="$lnSpcReduction"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='2'">
                <xsl:call-template name ="tmpOutlineLvl3_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'lineSpacing'"/>
                  <xsl:with-param name ="lnSpaceRed" select ="$lnSpcReduction"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='3'">
                <xsl:call-template name ="tmpOutlineLvl4_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'lineSpacing'"/>
                  <xsl:with-param name ="lnSpaceRed" select ="$lnSpcReduction"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='4'">
                <xsl:call-template name ="tmpOutlineLvl5_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'lineSpacing'"/>
                  <xsl:with-param name ="lnSpaceRed" select ="$lnSpcReduction"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='5'">
                <xsl:call-template name ="tmpOutlineLvl6_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'lineSpacing'"/>
                  <xsl:with-param name ="lnSpaceRed" select ="$lnSpcReduction"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='6'">
                <xsl:call-template name ="tmpOutlineLvl7_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'lineSpacing'"/>
                  <xsl:with-param name ="lnSpaceRed" select ="$lnSpcReduction"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='7'">
                <xsl:call-template name ="tmpOutlineLvl8_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'lineSpacing'"/>
                  <xsl:with-param name ="lnSpaceRed" select ="$lnSpcReduction"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='8'">
                <xsl:call-template name ="tmpOutlineLvl9_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'lineSpacing'"/>
                  <xsl:with-param name ="lnSpaceRed" select ="$lnSpcReduction"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="not(parent::node()/a:bodyPr/@vert='vert')">
      <xsl:if test="$spType='outline'">
        <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
          <xsl:if test="not(./@type) and ./@idx=$index">
            <xsl:choose>
              <xsl:when test="$level='0'">
                <xsl:call-template name ="tmpOutlineLvl1_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'Textdirection'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='1'">
                <xsl:call-template name ="tmpOutlineLvl2_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'Textdirection'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='2'">
                <xsl:call-template name ="tmpOutlineLvl3_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'Textdirection'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='3'">
                <xsl:call-template name ="tmpOutlineLvl4_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'Textdirection'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='4'">
                <xsl:call-template name ="tmpOutlineLvl5_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'Textdirection'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='5'">
                <xsl:call-template name ="tmpOutlineLvl6_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'Textdirection'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='6'">
                <xsl:call-template name ="tmpOutlineLvl7_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'Textdirection'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='7'">
                <xsl:call-template name ="tmpOutlineLvl8_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'Textdirection'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$level='8'">
                <xsl:call-template name ="tmpOutlineLvl9_ParaProperty">
                  <xsl:with-param name ="AttrType" select ="'Textdirection'"/>
                  <xsl:with-param name ="level" select ="$level+1"/>
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
  </xsl:template>
  <xsl:template name="tmpSlideParaProp">
    <xsl:param name="lnSpcReduction"/>
    <xsl:if test ="parent::node()/a:bodyPr/@vert='vert'">
      <xsl:attribute name ="style:writing-mode">
        <xsl:value-of select ="'tb-rl'"/>
      </xsl:attribute>
    </xsl:if>
    <!--Text alignment-->
    <xsl:if test ="a:pPr/@algn">
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
    </xsl:if >
    <!-- Convert Laeft margin of the paragraph-->
    <xsl:if test ="a:pPr/@marL">
      <xsl:attribute name ="fo:margin-left">
        <xsl:value-of select="concat(format-number(a:pPr/@marL div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="a:pPr/@marR">
      <xsl:attribute name ="fo:margin-right">
        <xsl:value-of select="concat(format-number(a:pPr/@marR div 360000, '#.##'), 'cm')"/>
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
    <xsl:if test ="a:pPr/a:lnSpc/a:spcPct/@val">
      <xsl:choose>
        <xsl:when test="$lnSpcReduction='0'">
          <xsl:attribute name="fo:line-height">
            <xsl:value-of select="concat(format-number(a:pPr/a:lnSpc/a:spcPct/@val div 1000,'###'), '%')"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="fo:line-height">
            <xsl:value-of select="concat(format-number((a:pPr/a:lnSpc/a:spcPct/@val - $lnSpcReduction) div 1000,'###'), '%')"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if >
    <xsl:if test ="a:pPr/a:lnSpc/a:spcPts/@val">
      <xsl:attribute name="style:line-height-at-least">
        <xsl:value-of select="concat(format-number(a:pPr/a:lnSpc/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>
  <xsl:template name="tmpOutlineLvl1_ParaProperty">
    <xsl:param name ="AttrType" />
    <xsl:param name ="lnSpaceRed" />
    <xsl:param name="slideMasterName"/>
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test ="$level='1'">
        <xsl:choose>
          <xsl:when test="$AttrType='TextAlignment'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@algn">
              <xsl:attribute name ="fo:text-align">
                <xsl:choose>
                  <!-- Center Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@algn ='ctr'">
                    <xsl:value-of select ="'center'"/>
                  </xsl:when>
                  <!-- Right Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@algn ='r'">
                    <xsl:value-of select ="'end'"/>
                  </xsl:when>
                  <!-- Left Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@algn ='l'">
                    <xsl:value-of select ="'start'"/>
                  </xsl:when>
                  <!-- Justified Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@algn ='just'">
                    <xsl:value-of select ="'justify'"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginLeft'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@marL">
              <xsl:attribute name ="fo:margin-left">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@marL div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@marL)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl1pPr/@marL">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-left">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginRight'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@marR">
              <xsl:attribute name ="fo:margin-right">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/p:txBody//a:lvl1pPr/@marR div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@marR)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl1pPr/@marR">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-right">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginTop'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:spcBef/a:spcPts/@val">
              <xsl:attribute name ="fo:margin-top">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:spcBef/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:spcBef/a:spcPts/@val)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl1pPr/a:spcBef/a:spcPts/@val">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-top">
                    <xsl:value-of select="concat(format-number(. div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginBottom'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:spcAft/a:spcPts/@val">
              <xsl:attribute name ="fo:margin-bottom">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:spcAft/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:spcAft/a:spcPts/@val)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl1pPr/a:spcAft/a:spcPts/@val">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-bottom">
                    <xsl:value-of select="concat(format-number(. div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='indent'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@indent">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/@indent div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='lineSpacing'">
            <xsl:choose>
              <xsl:when test="$lnSpaceRed='0'">
                <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:lnSpc/a:spcPct/@val">
                  <xsl:attribute name="fo:line-height">
                    <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:lnSpc/a:spcPct/@val div 1000,'###'), '%')"/>
                    <!--<xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:lnSpc/a:spcPct/@val"/>-->
                  </xsl:attribute>
                </xsl:if>
              </xsl:when>
              <xsl:otherwise>
                <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:lnSpc/a:spcPct/@val">
                  <xsl:attribute name="fo:line-height">
                    <xsl:value-of select="concat(format-number((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:lnSpc/a:spcPct/@val - number($lnSpaceRed)) div 1000,'###'), '%')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:otherwise>
            </xsl:choose>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:lnSpc/a:spcPts/@val">
              <xsl:attribute name="style:line-height-at-least">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:lnSpc/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='Textdirection'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@vert='vert'">
              <xsl:attribute name ="style:writing-mode">
                <xsl:value-of select ="'tb-rl'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpOutlineLvl2_ParaProperty">
    <xsl:param name ="AttrType" />
    <xsl:param name ="lnSpaceRed" />
    <xsl:param name="slideMasterName"/>
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test ="$level='2'">
        <xsl:choose>
          <xsl:when test="$AttrType='Textdirection'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@vert='vert'">
              <xsl:attribute name ="style:writing-mode">
                <xsl:value-of select ="'tb-rl'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='TextAlignment'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/@algn">
              <xsl:attribute name ="fo:text-align">
                <xsl:choose>
                  <!-- Center Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/@algn ='ctr'">
                    <xsl:value-of select ="'center'"/>
                  </xsl:when>
                  <!-- Right Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/@algn ='r'">
                    <xsl:value-of select ="'end'"/>
                  </xsl:when>
                  <!-- Left Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/@algn ='l'">
                    <xsl:value-of select ="'start'"/>
                  </xsl:when>
                  <!-- Justified Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/@algn ='just'">
                    <xsl:value-of select ="'justify'"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginLeft'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/@marL">
              <xsl:attribute name ="fo:margin-left">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/@marL div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/@marL)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl2pPr/@marL">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-left">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginRight'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/@marR">
              <xsl:attribute name ="fo:margin-right">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/p:txBody//a:lvl2pPr/@marR div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/@marR)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl2pPr/@marR">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-right">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginTop'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:spcBef/a:spcPts/@val">
              <xsl:attribute name ="fo:margin-top">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:spcBef/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:spcBef/a:spcPts/@val)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl2pPr/a:spcBef/a:spcPts/@val">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-top">
                    <xsl:value-of select="concat(format-number(. div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginBottom'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:spcAft/a:spcPts/@val">
              <xsl:attribute name ="fo:margin-bottom">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:spcAft/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:spcAft/a:spcPts/@val)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl2pPr/a:spcAft/a:spcPts/@val">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-bottom">
                    <xsl:value-of select="concat(format-number(. div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='indent'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/@indent">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/@indent div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/@indent)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl2pPr/@indent">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:text-indent">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='lineSpacing'">
            <xsl:choose>
              <xsl:when test="$lnSpaceRed='0'">
                <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:lnSpc/a:spcPct/@val">
                  <xsl:attribute name="fo:line-height">
                    <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:lnSpc/a:spcPct/@val div 1000,'###'), '%')"/>
                    <!--<xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:lnSpc/a:spcPct/@val"/>-->
                  </xsl:attribute>
                </xsl:if>
              </xsl:when>
              <xsl:otherwise>
                <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:lnSpc/a:spcPct/@val">
                  <xsl:attribute name="fo:line-height">
                    <xsl:value-of select="concat(format-number((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:lnSpc/a:spcPct/@val - number($lnSpaceRed)) div 1000,'###'), '%')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:otherwise>
            </xsl:choose>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:lnSpc/a:spcPts/@val">
              <xsl:attribute name="style:line-height-at-least">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl2pPr/a:lnSpc/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpOutlineLvl3_ParaProperty">
    <xsl:param name ="AttrType" />
    <xsl:param name ="lnSpaceRed" />
    <xsl:param name="slideMasterName"/>
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test ="$level='3'">
        <xsl:choose>
          <xsl:when test="$AttrType='Textdirection'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@vert='vert'">
              <xsl:attribute name ="style:writing-mode">
                <xsl:value-of select ="'tb-rl'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='TextAlignment'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/@algn">
              <xsl:attribute name ="fo:text-align">
                <xsl:choose>
                  <!-- Center Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/@algn ='ctr'">
                    <xsl:value-of select ="'center'"/>
                  </xsl:when>
                  <!-- Right Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/@algn ='r'">
                    <xsl:value-of select ="'end'"/>
                  </xsl:when>
                  <!-- Left Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/@algn ='l'">
                    <xsl:value-of select ="'start'"/>
                  </xsl:when>
                  <!-- Justified Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/@algn ='just'">
                    <xsl:value-of select ="'justify'"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginLeft'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/@marL">
              <xsl:attribute name ="fo:margin-left">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/@marL div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/@marL)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl3pPr/@marL">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-left">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginRight'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/@marR">
              <xsl:attribute name ="fo:margin-right">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/p:txBody//a:lvl3pPr/@marR div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/@marR)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl3pPr/@marR">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-right">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginTop'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:spcBef/a:spcPts/@val">
              <xsl:attribute name ="fo:margin-top">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:spcBef/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:spcBef/a:spcPts/@val)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl3pPr/a:spcBef/a:spcPts/@val">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-top">
                    <xsl:value-of select="concat(format-number(. div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginBottom'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:spcAft/a:spcPts/@val">
              <xsl:attribute name ="fo:margin-bottom">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:spcAft/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:spcAft/a:spcPts/@val)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl3pPr/a:spcAft/a:spcPts/@val">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-bottom">
                    <xsl:value-of select="concat(format-number(. div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='indent'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/@indent">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/@indent div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/@indent)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl3pPr/@indent">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:text-indent">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='lineSpacing'">
            <xsl:choose>
              <xsl:when test="$lnSpaceRed='0'">
                <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:lnSpc/a:spcPct/@val">
                  <xsl:attribute name="fo:line-height">
                    <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:lnSpc/a:spcPct/@val div 1000,'###'), '%')"/>
                    <!--<xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:lnSpc/a:spcPct/@val"/>-->
                  </xsl:attribute>
                </xsl:if>
              </xsl:when>
              <xsl:otherwise>
                <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:lnSpc/a:spcPct/@val">
                  <xsl:attribute name="fo:line-height">
                    <xsl:value-of select="concat(format-number((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:lnSpc/a:spcPct/@val - number($lnSpaceRed)) div 1000,'###'), '%')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:otherwise>
            </xsl:choose>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:lnSpc/a:spcPts/@val">
              <xsl:attribute name="style:line-height-at-least">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl3pPr/a:lnSpc/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpOutlineLvl4_ParaProperty">
    <xsl:param name ="AttrType" />
    <xsl:param name ="lnSpaceRed" />
    <xsl:param name="slideMasterName"/>
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test ="$level='4'">
        <xsl:choose>
          <xsl:when test="$AttrType='Textdirection'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@vert='vert'">
              <xsl:attribute name ="style:writing-mode">
                <xsl:value-of select ="'tb-rl'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='TextAlignment'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/@algn">
              <xsl:attribute name ="fo:text-align">
                <xsl:choose>
                  <!-- Center Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/@algn ='ctr'">
                    <xsl:value-of select ="'center'"/>
                  </xsl:when>
                  <!-- Right Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/@algn ='r'">
                    <xsl:value-of select ="'end'"/>
                  </xsl:when>
                  <!-- Left Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/@algn ='l'">
                    <xsl:value-of select ="'start'"/>
                  </xsl:when>
                  <!-- Justified Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/@algn ='just'">
                    <xsl:value-of select ="'justify'"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginLeft'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/@marL">
              <xsl:attribute name ="fo:margin-left">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/@marL div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/@marL)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl4pPr/@marL">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-left">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginRight'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/@marR">
              <xsl:attribute name ="fo:margin-right">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/p:txBody//a:lvl4pPr/@marR div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/@marR)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl4pPr/@marR">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-right">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginTop'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:spcBef/a:spcPts/@val">
              <xsl:attribute name ="fo:margin-top">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:spcBef/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:spcBef/a:spcPts/@val)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl4pPr/a:spcBef/a:spcPts/@val">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-top">
                    <xsl:value-of select="concat(format-number(. div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginBottom'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:spcAft/a:spcPts/@val">
              <xsl:attribute name ="fo:margin-bottom">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:spcAft/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:spcAft/a:spcPts/@val)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl4pPr/a:spcAft/a:spcPts/@val">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-bottom">
                    <xsl:value-of select="concat(format-number(. div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='indent'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/@indent">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/@indent div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/@indent)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl4pPr/@indent">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:text-indent">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='lineSpacing'">
            <xsl:choose>
              <xsl:when test="$lnSpaceRed='0'">
                <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:lnSpc/a:spcPct/@val">
                  <xsl:attribute name="fo:line-height">
                    <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:lnSpc/a:spcPct/@val div 1000,'###'), '%')"/>
                    <!--<xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:lnSpc/a:spcPct/@val"/>-->
                  </xsl:attribute>
                </xsl:if>
              </xsl:when>
              <xsl:otherwise>
                <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:lnSpc/a:spcPct/@val">
                  <xsl:attribute name="fo:line-height">
                    <xsl:value-of select="concat(format-number((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:lnSpc/a:spcPct/@val - number($lnSpaceRed)) div 1000,'###'), '%')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:otherwise>
            </xsl:choose>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:lnSpc/a:spcPts/@val">
              <xsl:attribute name="style:line-height-at-least">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl4pPr/a:lnSpc/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpOutlineLvl5_ParaProperty">
    <xsl:param name ="AttrType" />
    <xsl:param name ="lnSpaceRed" />
    <xsl:param name="slideMasterName"/>
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test ="$level='5'">
        <xsl:choose>
          <xsl:when test="$AttrType='Textdirection'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@vert='vert'">
              <xsl:attribute name ="style:writing-mode">
                <xsl:value-of select ="'tb-rl'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='TextAlignment'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/@algn">
              <xsl:attribute name ="fo:text-align">
                <xsl:choose>
                  <!-- Center Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/@algn ='ctr'">
                    <xsl:value-of select ="'center'"/>
                  </xsl:when>
                  <!-- Right Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/@algn ='r'">
                    <xsl:value-of select ="'end'"/>
                  </xsl:when>
                  <!-- Left Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/@algn ='l'">
                    <xsl:value-of select ="'start'"/>
                  </xsl:when>
                  <!-- Justified Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/@algn ='just'">
                    <xsl:value-of select ="'justify'"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginLeft'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/@marL">
              <xsl:attribute name ="fo:margin-left">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/@marL div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/@marL)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl5pPr/@marL">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-left">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginRight'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/@marR">
              <xsl:attribute name ="fo:margin-right">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/p:txBody//a:lvl5pPr/@marR div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/@marR)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl5pPr/@marR">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-right">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginTop'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:spcBef/a:spcPts/@val">
              <xsl:attribute name ="fo:margin-top">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:spcBef/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:spcBef/a:spcPts/@val)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl5pPr/a:spcBef/a:spcPts/@val">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-top">
                    <xsl:value-of select="concat(format-number(. div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginBottom'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:spcAft/a:spcPts/@val">
              <xsl:attribute name ="fo:margin-bottom">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:spcAft/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:spcAft/a:spcPts/@val)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl5pPr/a:spcAft/a:spcPts/@val">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-bottom">
                    <xsl:value-of select="concat(format-number(. div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='indent'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/@indent">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/@indent div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/@indent)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl5pPr/@indent">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:text-indent">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='lineSpacing'">
            <xsl:choose>
              <xsl:when test="$lnSpaceRed='0'">
                <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:lnSpc/a:spcPct/@val">
                  <xsl:attribute name="fo:line-height">
                    <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:lnSpc/a:spcPct/@val div 1000,'###'), '%')"/>
                    <!--<xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:lnSpc/a:spcPct/@val"/>-->
                  </xsl:attribute>
                </xsl:if>
              </xsl:when>
              <xsl:otherwise>
                <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:lnSpc/a:spcPct/@val">
                  <xsl:attribute name="fo:line-height">
                    <xsl:value-of select="concat(format-number((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:lnSpc/a:spcPct/@val - number($lnSpaceRed)) div 1000,'###'), '%')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:otherwise>
            </xsl:choose>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:lnSpc/a:spcPts/@val">
              <xsl:attribute name="style:line-height-at-least">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl5pPr/a:lnSpc/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpOutlineLvl6_ParaProperty">
    <xsl:param name ="AttrType" />
    <xsl:param name ="lnSpaceRed" />
    <xsl:param name="slideMasterName"/>
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test ="$level='6'">
        <xsl:choose>
          <xsl:when test="$AttrType='Textdirection'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@vert='vert'">
              <xsl:attribute name ="style:writing-mode">
                <xsl:value-of select ="'tb-rl'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='TextAlignment'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/@algn">
              <xsl:attribute name ="fo:text-align">
                <xsl:choose>
                  <!-- Center Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/@algn ='ctr'">
                    <xsl:value-of select ="'center'"/>
                  </xsl:when>
                  <!-- Right Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/@algn ='r'">
                    <xsl:value-of select ="'end'"/>
                  </xsl:when>
                  <!-- Left Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/@algn ='l'">
                    <xsl:value-of select ="'start'"/>
                  </xsl:when>
                  <!-- Justified Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/@algn ='just'">
                    <xsl:value-of select ="'justify'"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginLeft'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/@marL">
              <xsl:attribute name ="fo:margin-left">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/@marL div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/@marL)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl6pPr/@marL">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-left">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginRight'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/@marR">
              <xsl:attribute name ="fo:margin-right">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/p:txBody//a:lvl6pPr/@marR div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/@marR)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl6pPr/@marR">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-right">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginTop'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:spcBef/a:spcPts/@val">
              <xsl:attribute name ="fo:margin-top">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:spcBef/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:spcBef/a:spcPts/@val)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl6pPr/a:spcBef/a:spcPts/@val">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-top">
                    <xsl:value-of select="concat(format-number(. div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginBottom'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:spcAft/a:spcPts/@val">
              <xsl:attribute name ="fo:margin-bottom">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:spcAft/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:spcAft/a:spcPts/@val)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl6pPr/a:spcAft/a:spcPts/@val">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-bottom">
                    <xsl:value-of select="concat(format-number(. div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='indent'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/@indent">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/@indent div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/@indent)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl6pPr/@indent">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:text-indent">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='lineSpacing'">
            <xsl:choose>
              <xsl:when test="$lnSpaceRed='0'">
                <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:lnSpc/a:spcPct/@val">
                  <xsl:attribute name="fo:line-height">
                    <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:lnSpc/a:spcPct/@val div 1000,'###'), '%')"/>
                    <!--<xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:lnSpc/a:spcPct/@val"/>-->
                  </xsl:attribute>
                </xsl:if>
              </xsl:when>
              <xsl:otherwise>
                <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:lnSpc/a:spcPct/@val">
                  <xsl:attribute name="fo:line-height">
                    <xsl:value-of select="concat(format-number((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:lnSpc/a:spcPct/@val - number($lnSpaceRed)) div 1000,'###'), '%')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:otherwise>
            </xsl:choose>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:lnSpc/a:spcPts/@val">
              <xsl:attribute name="style:line-height-at-least">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl6pPr/a:lnSpc/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpOutlineLvl7_ParaProperty">
    <xsl:param name ="AttrType" />
    <xsl:param name ="lnSpaceRed" />
    <xsl:param name="slideMasterName"/>
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test ="$level='7'">
        <xsl:choose>
          <xsl:when test="$AttrType='Textdirection'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@vert='vert'">
              <xsl:attribute name ="style:writing-mode">
                <xsl:value-of select ="'tb-rl'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='TextAlignment'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/@algn">
              <xsl:attribute name ="fo:text-align">
                <xsl:choose>
                  <!-- Center Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/@algn ='ctr'">
                    <xsl:value-of select ="'center'"/>
                  </xsl:when>
                  <!-- Right Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/@algn ='r'">
                    <xsl:value-of select ="'end'"/>
                  </xsl:when>
                  <!-- Left Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/@algn ='l'">
                    <xsl:value-of select ="'start'"/>
                  </xsl:when>
                  <!-- Justified Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/@algn ='just'">
                    <xsl:value-of select ="'justify'"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginLeft'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/@marL">
              <xsl:attribute name ="fo:margin-left">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/@marL div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/@marL)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl7pPr/@marL">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-left">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginRight'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/@marR">
              <xsl:attribute name ="fo:margin-right">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/p:txBody//a:lvl7pPr/@marR div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/@marR)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl7pPr/@marR">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-right">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginTop'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:spcBef/a:spcPts/@val">
              <xsl:attribute name ="fo:margin-top">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:spcBef/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:spcBef/a:spcPts/@val)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl7pPr/a:spcBef/a:spcPts/@val">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-top">
                    <xsl:value-of select="concat(format-number(. div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginBottom'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:spcAft/a:spcPts/@val">
              <xsl:attribute name ="fo:margin-bottom">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:spcAft/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:spcAft/a:spcPts/@val)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl7pPr/a:spcAft/a:spcPts/@val">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-bottom">
                    <xsl:value-of select="concat(format-number(. div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='indent'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/@indent">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/@indent div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/@indent)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl7pPr/@indent">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:text-indent">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='lineSpacing'">
            <xsl:choose>
              <xsl:when test="$lnSpaceRed='0'">
                <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:lnSpc/a:spcPct/@val">
                  <xsl:attribute name="fo:line-height">
                    <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:lnSpc/a:spcPct/@val div 1000,'###'), '%')"/>
                    <!--<xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:lnSpc/a:spcPct/@val"/>-->
                  </xsl:attribute>
                </xsl:if>
              </xsl:when>
              <xsl:otherwise>
                <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:lnSpc/a:spcPct/@val">
                  <xsl:attribute name="fo:line-height">
                    <xsl:value-of select="concat(format-number((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:lnSpc/a:spcPct/@val - number($lnSpaceRed)) div 1000,'###'), '%')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:otherwise>
            </xsl:choose>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:lnSpc/a:spcPts/@val">
              <xsl:attribute name="style:line-height-at-least">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl7pPr/a:lnSpc/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpOutlineLvl8_ParaProperty">
    <xsl:param name ="AttrType" />
    <xsl:param name ="lnSpaceRed" />
    <xsl:param name="slideMasterName"/>
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test ="$level='8'">
        <xsl:choose>
          <xsl:when test="$AttrType='Textdirection'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@vert='vert'">
              <xsl:attribute name ="style:writing-mode">
                <xsl:value-of select ="'tb-rl'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='TextAlignment'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/@algn">
              <xsl:attribute name ="fo:text-align">
                <xsl:choose>
                  <!-- Center Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/@algn ='ctr'">
                    <xsl:value-of select ="'center'"/>
                  </xsl:when>
                  <!-- Right Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/@algn ='r'">
                    <xsl:value-of select ="'end'"/>
                  </xsl:when>
                  <!-- Left Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/@algn ='l'">
                    <xsl:value-of select ="'start'"/>
                  </xsl:when>
                  <!-- Justified Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/@algn ='just'">
                    <xsl:value-of select ="'justify'"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginLeft'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/@marL">
              <xsl:attribute name ="fo:margin-left">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/@marL div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/@marL)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl8pPr/@marL">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-left">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginRight'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/@marR">
              <xsl:attribute name ="fo:margin-right">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/p:txBody//a:lvl8pPr/@marR div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/@marR)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl8pPr/@marR">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-right">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginTop'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:spcBef/a:spcPts/@val">
              <xsl:attribute name ="fo:margin-top">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:spcBef/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:spcBef/a:spcPts/@val)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl8pPr/a:spcBef/a:spcPts/@val">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-top">
                    <xsl:value-of select="concat(format-number(. div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginBottom'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:spcAft/a:spcPts/@val">
              <xsl:attribute name ="fo:margin-bottom">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:spcAft/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:spcAft/a:spcPts/@val)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl8pPr/a:spcAft/a:spcPts/@val">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-bottom">
                    <xsl:value-of select="concat(format-number(. div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='indent'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/@indent">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/@indent div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/@indent)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl8pPr/@indent">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:text-indent">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='lineSpacing'">
            <xsl:choose>
              <xsl:when test="$lnSpaceRed='0'">
                <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:lnSpc/a:spcPct/@val">
                  <xsl:attribute name="fo:line-height">
                    <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:lnSpc/a:spcPct/@val div 1000,'###'), '%')"/>
                    <!--<xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:lnSpc/a:spcPct/@val"/>-->
                  </xsl:attribute>
                </xsl:if>
              </xsl:when>
              <xsl:otherwise>
                <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:lnSpc/a:spcPct/@val">
                  <xsl:attribute name="fo:line-height">
                    <xsl:value-of select="concat(format-number((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:lnSpc/a:spcPct/@val - number($lnSpaceRed)) div 1000,'###'), '%')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:otherwise>
            </xsl:choose>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:lnSpc/a:spcPts/@val">
              <xsl:attribute name="style:line-height-at-least">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl8pPr/a:lnSpc/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpOutlineLvl9_ParaProperty">
    <xsl:param name ="AttrType" />
    <xsl:param name ="lnSpaceRed" />
    <xsl:param name="slideMasterName"/>
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test ="$level='9'">
        <xsl:choose>
          <xsl:when test="$AttrType='Textdirection'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@vert='vert'">
              <xsl:attribute name ="style:writing-mode">
                <xsl:value-of select ="'tb-rl'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='TextAlignment'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/@algn">
              <xsl:attribute name ="fo:text-align">
                <xsl:choose>
                  <!-- Center Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/@algn ='ctr'">
                    <xsl:value-of select ="'center'"/>
                  </xsl:when>
                  <!-- Right Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/@algn ='r'">
                    <xsl:value-of select ="'end'"/>
                  </xsl:when>
                  <!-- Left Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/@algn ='l'">
                    <xsl:value-of select ="'start'"/>
                  </xsl:when>
                  <!-- Justified Alignment-->
                  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/@algn ='just'">
                    <xsl:value-of select ="'justify'"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginLeft'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/@marL">
              <xsl:attribute name ="fo:margin-left">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/@marL div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/@marL)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl9pPr/@marL">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-left">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginRight'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/@marR">
              <xsl:attribute name ="fo:margin-right">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/p:txBody//a:lvl9pPr/@marR div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/@marR)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl9pPr/@marR">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-right">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginTop'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:spcBef/a:spcPts/@val">
              <xsl:attribute name ="fo:margin-top">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:spcBef/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:spcBef/a:spcPts/@val)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl9pPr/a:spcBef/a:spcPts/@val">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-top">
                    <xsl:value-of select="concat(format-number(. div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='marginBottom'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:spcAft/a:spcPts/@val">
              <xsl:attribute name ="fo:margin-bottom">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:spcAft/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:spcAft/a:spcPts/@val)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl9pPr/a:spcAft/a:spcPts/@val">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:margin-bottom">
                    <xsl:value-of select="concat(format-number(. div 2835, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='indent'">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/@indent">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/@indent div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/@indent)">
              <xsl:for-each select ="document($slideMasterName)//p:txStyles/p:bodyStyle/a:lvl9pPr/@indent">
                <xsl:if test=".">
                  <xsl:attribute name ="fo:text-indent">
                    <xsl:value-of select="concat(format-number(. div 360000, '#.##'), 'cm')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
          <xsl:when test="$AttrType='lineSpacing'">
            <xsl:choose>
              <xsl:when test="$lnSpaceRed='0'">
                <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:lnSpc/a:spcPct/@val">
                  <xsl:attribute name="fo:line-height">
                    <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:lnSpc/a:spcPct/@val div 1000,'###'), '%')"/>
                    <!--<xsl:value-of select="parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:lnSpc/a:spcPct/@val"/>-->
                  </xsl:attribute>
                </xsl:if>
              </xsl:when>
              <xsl:otherwise>
                <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:lnSpc/a:spcPct/@val">
                  <xsl:attribute name="fo:line-height">
                    <xsl:value-of select="concat(format-number((parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:lnSpc/a:spcPct/@val - number($lnSpaceRed)) div 1000,'###'), '%')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:otherwise>
            </xsl:choose>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:lnSpc/a:spcPts/@val">
              <xsl:attribute name="style:line-height-at-least">
                <xsl:value-of select="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody//a:lvl9pPr/a:lnSpc/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpCommanGraphicProperty">
    <xsl:param name="LayoutFileName"/>
    <xsl:param name="spType"/>
    <xsl:param name="index"/>
    <xsl:call-template name="tmpSlideGrahicProp"/>
    <xsl:if test ="not(p:spPr/a:noFill) and not(p:spPr/a:solidFill) and not(p:style/a:fillRef)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'Backcolor'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'Backcolor'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'Backcolor'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='outline'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
            <xsl:if test="not(./@type) and ./@idx=$index">
              <xsl:call-template name ="tmpLayoutGraphicProperty">
                <xsl:with-param name ="AttrType" select ="'Backcolor'"/>
              </xsl:call-template>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
    <xsl:if test ="not(p:spPr/a:ln/a:noFill) and not(p:spPr/a:ln/a:solidFill) and not(p:style/a:lnRef)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'Linecolor'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'Linecolor'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'Linecolor'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='outline'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
            <xsl:if test="not(./@type) and ./@idx=$index">
              <xsl:call-template name ="tmpLayoutGraphicProperty">
                <xsl:with-param name ="AttrType" select ="'Linecolor'"/>
              </xsl:call-template>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
    <xsl:if test ="not(p:txBody/a:bodyPr/@tIns)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'IntMargTop'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'IntMargTop'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'IntMargTop'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='outline'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
            <xsl:if test="not(./@type) and ./@idx=$index">
              <xsl:call-template name ="tmpLayoutGraphicProperty">
                <xsl:with-param name ="AttrType" select ="'IntMargTop'"/>
              </xsl:call-template>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
    <xsl:if test ="not(p:txBody/a:bodyPr/@lIns)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'IntMargLeft'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'IntMargLeft'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'IntMargLeft'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='outline'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
            <xsl:if test="not(./@type) and ./@idx=$index">
              <xsl:call-template name ="tmpLayoutGraphicProperty">
                <xsl:with-param name ="AttrType" select ="'IntMargLeft'"/>
              </xsl:call-template>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:if >
    <xsl:if test ="not(p:txBody/a:bodyPr/@bIns)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'IntMargBottom'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'IntMargBottom'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'IntMargBottom'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='outline'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
            <xsl:if test="not(./@type) and ./@idx=$index">
              <xsl:call-template name ="tmpLayoutGraphicProperty">
                <xsl:with-param name ="AttrType" select ="'IntMargBottom'"/>
              </xsl:call-template>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
    <xsl:if test ="not(p:txBody/a:bodyPr/@rIns)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'IntMargRight'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'IntMargRight'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'IntMargRight'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='outline'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
            <xsl:if test="not(./@type) and ./@idx=$index">
              <xsl:call-template name ="tmpLayoutGraphicProperty">
                <xsl:with-param name ="AttrType" select ="'IntMargRight'"/>
              </xsl:call-template>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
    <xsl:if test ="not(p:txBody/a:bodyPr/@anchor)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'VertAlign'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'VertAlign'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'VertAlign'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='outline'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
            <xsl:if test="not(./@type) and ./@idx=$index">
              <xsl:call-template name ="tmpLayoutGraphicProperty">
                <xsl:with-param name ="AttrType" select ="'VertAlign'"/>
              </xsl:call-template>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:if >
    <xsl:if test ="not(p:txBody/a:bodyPr/@anchorCtr)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'HorzAlign'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'HorzAlign'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'HorzAlign'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='outline'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
            <xsl:if test="not(./@type) and ./@idx=$index">
              <xsl:call-template name ="tmpLayoutGraphicProperty">
                <xsl:with-param name ="AttrType" select ="'HorzAlign'"/>
              </xsl:call-template>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
    <xsl:if test ="not(p:txBody/a:bodyPr/@wrap)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'Wrap'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'Wrap'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'Wrap'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='outline'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
            <xsl:if test="not(./@type) and ./@idx=$index">
              <xsl:call-template name ="tmpLayoutGraphicProperty">
                <xsl:with-param name ="AttrType" select ="'Wrap'"/>
              </xsl:call-template>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
    <xsl:if test ="not(p:txBody/a:bodyPr/spAutoFit)">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'spAutoFit'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'spAutoFit'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'spAutoFit'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='outline'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
            <xsl:if test="not(./@type) and ./@idx=$index">
              <xsl:call-template name ="tmpLayoutGraphicProperty">
                <xsl:with-param name ="AttrType" select ="'spAutoFit'"/>
              </xsl:call-template>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
    <xsl:if test ="not(p:txBody/a:bodyPr/@vert='vert')">
      <xsl:choose>
        <xsl:when test="$spType='title'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='ctrTitle' or @type= 'title']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'Textdirection'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='subtitle'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='subTitle']">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'Textdirection'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='body'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type='body' and @idx=$index]">
            <xsl:call-template name ="tmpLayoutGraphicProperty">
              <xsl:with-param name ="AttrType" select ="'Textdirection'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="$spType='outline'">
          <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
            <xsl:if test="not(./@type) and ./@idx=$index">
              <xsl:call-template name ="tmpLayoutGraphicProperty">
                <xsl:with-param name ="AttrType" select ="'Textdirection'"/>
              </xsl:call-template>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
  </xsl:template>
  <xsl:template name="tmpSlideGrahicProp">
    <!--Background Fill color-->
    <xsl:choose>
      <!-- No fill -->
      <xsl:when test ="p:spPr/a:noFill">
        <xsl:attribute name ="draw:fill">
          <xsl:value-of select="'none'" />
        </xsl:attribute>
      </xsl:when>
      <!-- Solid fill-->
      <!-- Standard color-->
      <xsl:when test ="p:spPr/a:solidFill">
        <xsl:if test ="p:spPr/a:solidFill/a:srgbClr/@val">
          <xsl:attribute name ="draw:fill">
            <xsl:value-of select="'solid'" />
          </xsl:attribute>
          <xsl:attribute name ="draw:fill-color">
            <xsl:value-of select="concat('#',p:spPr/a:solidFill/a:srgbClr/@val)"/>
          </xsl:attribute>
          <!-- Transparency percentage-->
          <xsl:if test="p:spPr/a:solidFill/a:srgbClr/a:alpha/@val">
            <xsl:variable name ="alpha">
              <xsl:value-of select ="p:spPr/a:solidFill/a:srgbClr/a:alpha/@val"/>
            </xsl:variable>
            <xsl:if test="($alpha != '') or ($alpha != 0)">
              <xsl:attribute name ="draw:opacity">
                <xsl:value-of select="concat(($alpha div 1000), '%')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
        </xsl:if>
        <!--Theme color-->
        <xsl:if test ="p:spPr/a:solidFill/a:schemeClr/@val">
          <xsl:attribute name ="draw:fill">
            <xsl:value-of select="'solid'" />
          </xsl:attribute>
          <xsl:attribute name ="draw:fill-color">
            <xsl:call-template name ="getColorCode">
              <xsl:with-param name ="color">
                <xsl:value-of select="p:spPr/a:solidFill/a:schemeClr/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumMod">
                <xsl:value-of select="p:spPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumOff">
                <xsl:value-of select="p:spPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
          <!-- Transparency percentage-->
          <xsl:if test="p:spPr/a:solidFill/a:schemeClr/a:alpha/@val">
            <xsl:variable name ="alpha">
              <xsl:value-of select ="p:spPr/a:solidFill/a:schemeClr/a:alpha/@val"/>
            </xsl:variable>
            <xsl:if test="($alpha != '') or ($alpha != 0)">
              <xsl:attribute name ="draw:opacity">
                <xsl:value-of select="concat(($alpha div 1000), '%')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
        </xsl:if>
      </xsl:when>
      <!--Fill refernce-->
      <xsl:when test ="p:style/a:fillRef">
        <!-- Standard color-->
        <xsl:if test ="p:style/a:fillRef/a:srgbClr/@val">
          <xsl:attribute name ="draw:fill">
            <xsl:value-of select="'solid'" />
          </xsl:attribute>
          <xsl:attribute name ="draw:fill-color">
            <xsl:value-of select="concat('#',p:style/a:fillRef/a:srgbClr/@val)"/>
          </xsl:attribute>
        </xsl:if>
        <!--Theme color-->
        <xsl:if test ="p:style/a:fillRef//a:schemeClr/@val">
          <xsl:attribute name ="draw:fill">
            <xsl:value-of select="'solid'" />
          </xsl:attribute>
          <xsl:attribute name ="draw:fill-color">
            <xsl:call-template name ="getColorCode">
              <xsl:with-param name ="color">
                <xsl:value-of select="p:style/a:fillRef/a:schemeClr/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumMod">
                <xsl:value-of select="p:style/a:fillRef/a:schemeClr/a:lumMod/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumOff">
                <xsl:value-of select="p:style/a:fillRef/a:schemeClr/a:lumOff/@val"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
    </xsl:choose>
    <!--Line Color-->
    <xsl:choose>
      <!-- No line-->
      <xsl:when test ="p:spPr/a:ln/a:noFill">
        <xsl:attribute name ="draw:stroke">
          <xsl:value-of select="'none'" />
        </xsl:attribute>
      </xsl:when>
      <!-- Solid line color-->
      <xsl:when test ="p:spPr/a:ln/a:solidFill">
        <xsl:attribute name ="draw:stroke">
          <xsl:value-of select="'solid'" />
        </xsl:attribute>
        <!-- Standard color for border-->
        <xsl:if test ="p:spPr/a:ln/a:solidFill/a:srgbClr/@val">
          <xsl:attribute name ="svg:stroke-color">
            <xsl:value-of select="concat('#',p:spPr/a:ln/a:solidFill/a:srgbClr/@val)"/>
          </xsl:attribute>
          <!-- Transparency percentage-->
          <xsl:if test="p:spPr/a:ln/a:solidFill/a:srgbClr/a:alpha/@val">
            <xsl:variable name ="alpha">
              <xsl:value-of select ="p:spPr/a:ln/a:solidFill/a:srgbClr/a:alpha/@val"/>
            </xsl:variable>
            <xsl:if test="($alpha != '') or ($alpha != 0)">
              <xsl:attribute name ="svg:stroke-opacity">
                <xsl:value-of select="concat(($alpha div 1000), '%')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
        </xsl:if>
        <!-- Theme color for border-->
        <xsl:if test ="p:spPr/a:ln/a:solidFill/a:schemeClr/@val">
          <xsl:attribute name ="svg:stroke-color">
            <xsl:call-template name ="getColorCode">
              <xsl:with-param name ="color">
                <xsl:value-of select="p:spPr/a:ln/a:solidFill/a:schemeClr/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumMod">
                <xsl:value-of select="p:spPr/a:ln/a:solidFill/a:schemeClr/a:lumMod/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumOff">
                <xsl:value-of select="p:spPr/a:ln/a:solidFill/a:schemeClr/a:lumOff/@val"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
          <!-- Transparency percentage-->
          <xsl:if test="p:spPr/a:ln/a:solidFill/a:schemeClr/a:alpha/@val">
            <xsl:variable name ="alpha">
              <xsl:value-of select ="p:spPr/a:ln/a:solidFill/a:schemeClr/a:alpha/@val"/>
            </xsl:variable>
            <xsl:if test="($alpha != '') or ($alpha != 0)">
              <xsl:attribute name ="svg:stroke-opacity">
                <xsl:value-of select="concat(($alpha div 1000), '%')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
        </xsl:if>
      </xsl:when>
      <xsl:when test ="p:style/a:lnRef">
        <xsl:attribute name ="draw:stroke">
          <xsl:value-of select="'solid'" />
        </xsl:attribute>
        <!--Standard color for border-->
        <xsl:if test ="p:style/a:lnRef/a:srgbClr/@val">
          <xsl:attribute name ="svg:stroke-color">
            <xsl:value-of select="concat('#',p:style/a:lnRef/a:srgbClr/@val)"/>
          </xsl:attribute>
        </xsl:if>
        <!--Theme color for border-->
        <xsl:if test ="p:style/a:lnRef/a:schemeClr/@val">
          <xsl:attribute name ="svg:stroke-color">
            <xsl:call-template name ="getColorCode">
              <xsl:with-param name ="color">
                <xsl:value-of select="p:style/a:lnRef/a:schemeClr/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumMod">
                <xsl:value-of select="p:style/a:lnRef/a:schemeClr/a:lumMod/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumOff">
                <xsl:value-of select="p:style/a:lnRef/a:schemeClr/a:lumOff/@val"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
    </xsl:choose>
    <!--Line Style-->
    <xsl:call-template name="LineStyle"/>
    <xsl:if test ="p:txBody/a:bodyPr/@tIns">
      <xsl:attribute name ="fo:padding-top">
        <xsl:value-of select ="concat(format-number(p:txBody/a:bodyPr/@tIns div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="p:txBody/a:bodyPr/@lIns">
      <xsl:attribute name ="fo:padding-left">
        <xsl:value-of select ="concat(format-number(p:txBody/a:bodyPr/@lIns div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="p:txBody/a:bodyPr/@bIns">
      <xsl:attribute name ="fo:padding-bottom">
        <xsl:value-of select ="concat(format-number(p:txBody/a:bodyPr/@bIns div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="p:txBody/a:bodyPr/@rIns">
      <xsl:attribute name ="fo:padding-right">
        <xsl:value-of select ="concat(format-number(p:txBody/a:bodyPr/@rIns div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="p:txBody/a:bodyPr/@anchor">
      <xsl:attribute name ="draw:textarea-vertical-align">
        <xsl:choose >
          <!-- Top-->
          <xsl:when test ="p:txBody/a:bodyPr/@anchor = 't' ">
            <xsl:value-of  select ="'top'"/>
          </xsl:when>
          <!-- Middle-->
          <xsl:when test ="p:txBody/a:bodyPr/@anchor = 'ctr' ">
            <xsl:value-of  select ="'middle'"/>
          </xsl:when>
          <!-- bottom-->
          <xsl:when test ="p:txBody/a:bodyPr/@anchor = 'b' ">
            <xsl:value-of  select ="'bottom'"/>
          </xsl:when>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if >
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
    <xsl:if test ="p:txBody/a:bodyPr/@wrap">
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
    </xsl:if>
    <xsl:if test ="p:txBody/a:bodyPr/a:spAutoFit">
      <xsl:attribute name ="draw:auto-grow-height" >
        <xsl:value-of select ="'true'"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:attribute name ="fo:min-height" >
      <xsl:value-of select ="'12.573cm'"/>
    </xsl:attribute>
    <xsl:if test ="p:txBody/a:bodyPr/@vert='vert'">
      <xsl:attribute name ="style:writing-mode">
        <xsl:value-of select ="'tb-rl'"/>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>
  <xsl:template name="tmpLayoutGraphicProperty">
    <xsl:param name ="AttrType" />
    <xsl:choose>
      <xsl:when test="$AttrType='Backcolor'">
        <xsl:choose>
          <!-- No fill -->
          <xsl:when test ="parent::node()/parent::node()/parent::node()/p:spPr/a:noFill">
            <xsl:attribute name ="draw:fill">
              <xsl:value-of select="'none'" />
            </xsl:attribute>
          </xsl:when>
          <!-- Solid fill-->
          <!-- Standard color-->
          <xsl:when test ="parent::node()/parent::node()/parent::node()/p:spPr/a:solidFill">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:spPr/a:solidFill/a:srgbClr/@val">
              <xsl:attribute name ="draw:fill">
                <xsl:value-of select="'solid'" />
              </xsl:attribute>
              <xsl:attribute name ="draw:fill-color">
                <xsl:value-of select="concat('#',parent::node()/parent::node()/parent::node()/p:spPr/a:solidFill/a:srgbClr/@val)"/>
              </xsl:attribute>
              <!-- Transparency percentage-->
              <xsl:if test="parent::node()/parent::node()/parent::node()/p:spPr/a:solidFill/a:srgbClr/a:alpha/@val">
                <xsl:variable name ="alpha">
                  <xsl:value-of select ="parent::node()/parent::node()/parent::node()/p:spPr/a:solidFill/a:srgbClr/a:alpha/@val"/>
                </xsl:variable>
                <xsl:if test="($alpha != '') or ($alpha != 0)">
                  <xsl:attribute name ="draw:opacity">
                    <xsl:value-of select="concat(($alpha div 1000), '%')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:if>
            </xsl:if>
            <!--Theme color-->
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:spPr/a:solidFill/a:schemeClr/@val">
              <xsl:attribute name ="draw:fill">
                <xsl:value-of select="'solid'" />
              </xsl:attribute>
              <xsl:attribute name ="draw:fill-color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:spPr/a:solidFill/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:spPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:spPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
              <!-- Transparency percentage-->
              <xsl:if test="parent::node()/parent::node()/parent::node()/p:spPr/a:solidFill/a:schemeClr/a:alpha/@val">
                <xsl:variable name ="alpha">
                  <xsl:value-of select ="parent::node()/parent::node()/parent::node()/p:spPr/a:solidFill/a:schemeClr/a:alpha/@val"/>
                </xsl:variable>
                <xsl:if test="($alpha != '') or ($alpha != 0)">
                  <xsl:attribute name ="draw:opacity">
                    <xsl:value-of select="concat(($alpha div 1000), '%')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:if>
            </xsl:if>
          </xsl:when>
          <!--Fill refernce-->
          <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fillRef">
            <!-- Standard color-->
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:style/a:fillRef/a:srgbClr/@val">
              <xsl:attribute name ="draw:fill">
                <xsl:value-of select="'solid'" />
              </xsl:attribute>
              <xsl:attribute name ="draw:fill-color">
                <xsl:value-of select="concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fillRef/a:srgbClr/@val)"/>
              </xsl:attribute>
            </xsl:if>
            <!--Theme color-->
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:style/a:fillRef//a:schemeClr/@val">
              <xsl:attribute name ="draw:fill">
                <xsl:value-of select="'solid'" />
              </xsl:attribute>
              <xsl:attribute name ="draw:fill-color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:style/a:fillRef/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:style/a:fillRef/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:style/a:fillRef/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="$AttrType='Linecolor'">
        <xsl:choose>
          <!-- No line-->
          <xsl:when test ="parent::node()/parent::node()/parent::node()/p:spPr/a:ln/a:noFill">
            <xsl:attribute name ="draw:stroke">
              <xsl:value-of select="'none'" />
            </xsl:attribute>
          </xsl:when>
          <!-- Solid line color-->
          <xsl:when test ="parent::node()/parent::node()/parent::node()/p:spPr/a:ln/a:solidFill">
            <xsl:attribute name ="draw:stroke">
              <xsl:value-of select="'solid'" />
            </xsl:attribute>
            <!-- Standard color for border-->
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:spPr/a:ln/a:solidFill/a:srgbClr/@val">
              <xsl:attribute name ="svg:stroke-color">
                <xsl:value-of select="concat('#',parent::node()/parent::node()/parent::node()/p:spPr/a:ln/a:solidFill/a:srgbClr/@val)"/>
              </xsl:attribute>
              <!-- Transparency percentage-->
              <xsl:if test="parent::node()/parent::node()/parent::node()/p:spPr/a:ln/a:solidFill/a:srgbClr/a:alpha/@val">
                <xsl:variable name ="alpha">
                  <xsl:value-of select ="parent::node()/parent::node()/parent::node()/p:spPr/a:ln/a:solidFill/a:srgbClr/a:alpha/@val"/>
                </xsl:variable>
                <xsl:if test="($alpha != '') or ($alpha != 0)">
                  <xsl:attribute name ="svg:stroke-opacity">
                    <xsl:value-of select="concat(($alpha div 1000), '%')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:if>
            </xsl:if>
            <!-- Theme color for border-->
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:spPr/a:ln/a:solidFill/a:schemeClr/@val">
              <xsl:attribute name ="svg:stroke-color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:spPr/a:ln/a:solidFill/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:spPr/a:ln/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:spPr/a:ln/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
              <!-- Transparency percentage-->
              <xsl:if test="parent::node()/parent::node()/parent::node()/p:spPr/a:ln/a:solidFill/a:schemeClr/a:alpha/@val">
                <xsl:variable name ="alpha">
                  <xsl:value-of select ="parent::node()/parent::node()/parent::node()/p:spPr/a:ln/a:solidFill/a:schemeClr/a:alpha/@val"/>
                </xsl:variable>
                <xsl:if test="($alpha != '') or ($alpha != 0)">
                  <xsl:attribute name ="svg:stroke-opacity">
                    <xsl:value-of select="concat(($alpha div 1000), '%')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:if>
            </xsl:if>
          </xsl:when>
          <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:lnRef">
            <xsl:attribute name ="draw:stroke">
              <xsl:value-of select="'solid'" />
            </xsl:attribute>
            <!--Standard color for border-->
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:style/a:lnRef/a:srgbClr/@val">
              <xsl:attribute name ="svg:stroke-color">
                <xsl:value-of select="concat('#',parent::node()/parent::node()/parent::node()/p:style/a:lnRef/a:srgbClr/@val)"/>
              </xsl:attribute>
            </xsl:if>
            <!--Theme color for border-->
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:style/a:lnRef/a:schemeClr/@val">
              <xsl:attribute name ="svg:stroke-color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:style/a:lnRef/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:style/a:lnRef/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:style/a:lnRef/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="$AttrType='IntMargRight'">
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@rIns">
          <xsl:attribute name ="fo:padding-right">
            <xsl:value-of select ="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@rIns div 360000, '#.##'), 'cm')"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='IntMargLeft'">
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@lIns">
          <xsl:attribute name ="fo:padding-left">
            <xsl:value-of select ="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@lIns div 360000, '#.##'), 'cm')"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='IntMargTop'">
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@tIns">
          <xsl:attribute name ="fo:padding-top">
            <xsl:value-of select ="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@tIns div 360000, '#.##'), 'cm')"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='IntMargBottom'">
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@bIns">
          <xsl:attribute name ="fo:padding-bottom">
            <xsl:value-of select ="concat(format-number(parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@bIns div 360000, '#.##'), 'cm')"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='Wrap'">
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@wrap">
          <xsl:choose>
            <xsl:when test="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@wrap='none'">
              <xsl:attribute name ="fo:wrap-option">
                <xsl:value-of select ="'no-wrap'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@wrap='square'">
              <xsl:attribute name ="fo:wrap-option">
                <xsl:value-of select ="'wrap'"/>
              </xsl:attribute>
            </xsl:when>
          </xsl:choose>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='spAutoFit'">
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/a:spAutoFit">
          <xsl:attribute name ="draw:auto-grow-height" >
            <xsl:value-of select ="'true'"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='VertAlign'">
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@anchor">
          <xsl:attribute name ="draw:textarea-vertical-align">
            <xsl:choose>
              <!-- Top-->
              <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@anchor = 't' ">
                <xsl:value-of  select ="'top'"/>
              </xsl:when>
              <!-- Middle-->
              <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@anchor = 'ctr' ">
                <xsl:value-of  select ="'middle'"/>
              </xsl:when>
              <!-- bottom-->
              <xsl:when test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@anchor = 'b' ">
                <xsl:value-of  select ="'bottom'"/>
              </xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='HorzAlign'">
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@anchorCtr">
          <xsl:attribute name ="draw:textarea-horizontal-align">
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@anchorCtr= 1">
              <xsl:value-of select ="'center'"/>
            </xsl:if>
            <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@anchorCtr= 0">
              <xsl:value-of select ="'justify'"/>
            </xsl:if>
          </xsl:attribute >
        </xsl:if>
      </xsl:when>
      <xsl:when test="$AttrType='Textdirection'">
        <xsl:if test ="parent::node()/parent::node()/parent::node()/p:txBody/a:bodyPr/@vert='vert'">
          <xsl:attribute name ="style:writing-mode">
            <xsl:value-of select ="'tb-rl'"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpGetCapFromLayout">
    <xsl:param name="LayoutFileName"/>
    <xsl:param name="spType"/>
    <xsl:param name="index"/>
    <xsl:for-each select="document(concat('ppt/slideLayouts/',$LayoutFileName))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type=$spType and @idx=$index]">
      <xsl:if test="parent::node()/parent::node()/parent::node()/p:txBody//a:lvl1pPr/a:defRPr/@cap">
        <xsl:choose>
          <xsl:when test ="@cap='all'">
            <xsl:value-of select ="'true'"/>
          </xsl:when>
          <xsl:when test ="@cap='small'">
            <xsl:value-of select ="'false'"/>
          </xsl:when>
        </xsl:choose>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

  <!--End-->

  <!--End-->

</xsl:stylesheet>
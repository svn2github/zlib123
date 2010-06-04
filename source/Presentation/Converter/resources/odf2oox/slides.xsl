﻿<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
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
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"   
  xmlns:odf="urn:odf"
  xmlns:fn="http://www.w3.org/2005/02/xpath-functions"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"  
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" 
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:page="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" 
  xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:dom="http://www.w3.org/2001/xml-events" 
  xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:rels="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" 
  xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0"	
  xmlns:msxsl="urn:schemas-microsoft-com:xslt"  
  
  exclude-result-prefixes="odf style text number draw page">


  <xsl:template name ="slides" match ="/office:document-content/office:body/office:presentation/draw:page" mode="slide">
    <xsl:param name ="pageNo"/>
    <xsl:message terminate="no">progress:text:p</xsl:message>
    <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
				   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
				   xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <!--Added by sateesh - Background Color-->
        <xsl:variable name="dpName">
          <xsl:value-of select="@draw:style-name" />
        </xsl:variable>
        <xsl:for-each select="//style:style[@style:name= $dpName]/style:drawing-page-properties">
          <xsl:choose>
            <xsl:when test="@draw:fill='bitmap'">
              <xsl:variable name="var_imageName" select="@draw:fill-image-name"/>
              <xsl:if test="document('styles.xml')/office:document-styles/office:styles/draw:fill-image[@draw:name=$var_imageName]">
                <p:bg>
                  <p:bgPr>
                    <a:blipFill dpi="0" rotWithShape="1">
                      <xsl:call-template name="tmpInsertBackImage">
                        <xsl:with-param name="FileName" select="concat('slide',$pageNo)" />
                        <xsl:with-param name="imageName" select="@draw:fill-image-name" />
                      </xsl:call-template>
                      <xsl:if test="@style:repeat='stretch'">
                        <a:stretch>
                          <a:fillRect />
                        </a:stretch>
                      </xsl:if>
                      <xsl:if test="@draw:fill-image-ref-point-x or @draw:fill-image-ref-point-y">
                        <a:tile tx="0" ty="0" flip="none">
                          <xsl:if test="@draw:fill-image-ref-point-x">
                            <xsl:attribute name="sx">
                              <xsl:value-of select="number(substring-before(@draw:fill-image-ref-point-x,'%')) * 1000"/>
                            </xsl:attribute>
                          </xsl:if>
                          <xsl:if test="@draw:fill-image-ref-point-y">
                            <xsl:attribute name="sy">
                              <xsl:value-of select="number(substring-before(@draw:fill-image-ref-point-y,'%')) * 1000"/>
                            </xsl:attribute>
                          </xsl:if>
                          <xsl:if test="@draw:fill-image-ref-point">
                            <xsl:attribute name="algn">
                              <xsl:choose>
                                <xsl:when test="@draw:fill-image-ref-point='top-left'">
                                  <xsl:value-of select ="'tl'"/>
                                </xsl:when>
                                <xsl:when test="@draw:fill-image-ref-point='top'">
                                  <xsl:value-of select ="'t'"/>
                                </xsl:when>
                                <xsl:when test="@draw:fill-image-ref-point='top-right'">
                                  <xsl:value-of select ="'tr'"/>
                                </xsl:when>
                                <xsl:when test="@draw:fill-image-ref-point='right'">
                                  <xsl:value-of select ="'r'"/>
                                </xsl:when>
                                <xsl:when test="@draw:fill-image-ref-point='bottom-left'">
                                  <xsl:value-of select ="'bl'"/>
                                </xsl:when>
                                <xsl:when test="@draw:fill-image-ref-point='bottom-right'">
                                  <xsl:value-of select ="'br'"/>
                                </xsl:when>
                                <xsl:when test="@draw:fill-image-ref-point='bottom'">
                                  <xsl:value-of select ="'b'"/>
                                </xsl:when>
                                <xsl:when test="@draw:fill-image-ref-point='center'">
                                  <xsl:value-of select ="'ctr'"/>
                                </xsl:when>
                              </xsl:choose>
                            </xsl:attribute>
                          </xsl:if>
                        </a:tile>
                      </xsl:if>
                    </a:blipFill>
                    <a:effectLst />
                  </p:bgPr>
                </p:bg>
              </xsl:if>
            </xsl:when>
            <xsl:when test="@draw:fill='solid'">
              <p:bg>
                <p:bgPr>
                  <a:solidFill>
                    <xsl:variable name="varSrgbVal">
                        <xsl:if test="@draw:fill-color">
                          <xsl:value-of select="substring-after(@draw:fill-color,'#')" />
                        </xsl:if>
                        <xsl:if test="not(@draw:fill-color)">
                          <xsl:value-of select="'ffffff'" />
                        </xsl:if>
                    </xsl:variable>
                    <xsl:if test="$varSrgbVal != ''">
                      <a:srgbClr>
                        <xsl:attribute name ="val">
                          <xsl:value-of select ="$varSrgbVal"/>
                      </xsl:attribute>
                    </a:srgbClr>
                    </xsl:if>
                  </a:solidFill>
                  <a:effectLst />
                </p:bgPr>
              </p:bg>
            </xsl:when>
            <xsl:when test="@draw:fill='gradient'">
              <p:bg>
                <p:bgPr>
                  <xsl:call-template name="tmpGradientFill">
                    <xsl:with-param name="gradStyleName" select="@draw:fill-gradient-name"/>
                    <xsl:with-param  name="opacity" select="substring-before(@draw:opacity,'%')"/>
                  </xsl:call-template>
                  <a:effectLst />
                </p:bgPr>
              </p:bg>
            </xsl:when>
          </xsl:choose>
        </xsl:for-each>
        <!--end-->
        <p:spTree>
          <p:nvGrpSpPr>
            <p:cNvPr id="1" name=""/>
            <p:cNvGrpSpPr/>
            <p:nvPr/>
          </p:nvGrpSpPr>
          <p:grpSpPr>
            <a:xfrm>
              <a:off x="0" y="0"/>
              <a:ext cx="0" cy="0"/>
              <a:chOff x="0" y="0"/>
              <a:chExt cx="0" cy="0"/>
            </a:xfrm>
          </p:grpSpPr>
          <xsl:variable name ="pageStyle">
            <xsl:value-of select ="@draw:style-name"/>
          </xsl:variable>
          <xsl:variable name ="footerId">
            <xsl:value-of select ="@presentation:use-footer-name"/>
          </xsl:variable>
          <xsl:variable name ="dateId">
            <xsl:value-of select ="@presentation:use-date-time-name"/>
          </xsl:variable>
          <xsl:variable name ="SMName">
            <xsl:value-of select ="@draw:master-page-name"/>
          </xsl:variable>

          <xsl:for-each select ="//style:style[@style:name=$pageStyle]">
            <xsl:if test ="style:drawing-page-properties[@presentation:display-footer='true']">
              <!--<xsl:if test ="not($footerId ='')">-->
              <!--<xsl:if test ="document('content.xml')//presentation:footer-decl[@presentation:name=$footerId]">-->
              <xsl:call-template name ="footer" >
                <xsl:with-param name ="footerId" select ="$footerId"/>
                <xsl:with-param name ="SMName" select ="$SMName"/>
              </xsl:call-template >
              <!--</xsl:if>-->
            </xsl:if>
            <xsl:if test ="style:drawing-page-properties[@presentation:display-date-time='true']
                          or style:drawing-page-properties[presentation:background-objects-visible='true'] ">
              <!--<xsl:if test ="not($dateId='')">-->
              <!--<xsl:if test ="document('content.xml')//presentation:date-time-decl[@presentation:name=$dateId]">-->
                <xsl:call-template name ="footerDate" >
                  <xsl:with-param name ="footerDateId" select ="$dateId"/>
                  <xsl:with-param name ="SMName" select ="$SMName"/>
                </xsl:call-template >
              <!--</xsl:if>-->
            </xsl:if >
            <xsl:if test ="style:drawing-page-properties[@presentation:display-page-number='true']">
              <xsl:call-template name ="slideNumber">
                <xsl:with-param name ="pageNumber" select ="$pageNo"/>
                <xsl:with-param name ="SMName" select ="$SMName"/>
              </xsl:call-template >
            </xsl:if >
          </xsl:for-each>
          <xsl:for-each select="node()">
            <xsl:choose>
              <xsl:when test="name()='draw:frame'">
                <xsl:variable name="var_pos">
                <xsl:call-template name="getShapePosTemp">
                  <xsl:with-param name="var_pos" select="position()"/>
                </xsl:call-template>
              </xsl:variable>
                <xsl:choose>
                  <xsl:when test="./table:table and ./draw:image">
                    <xsl:call-template name="tmpTables">
                      <xsl:with-param name ="pageNo" select ="$pageNo"/>
                      <xsl:with-param name ="shapeCount" select="$var_pos" />
                    </xsl:call-template>

                  </xsl:when>
                  <xsl:when test="./draw:object or ./draw:object-ole">
                    <xsl:call-template name="tmpOLEObjects">
                      <xsl:with-param name ="pageNo" select ="$pageNo"/>
                      <xsl:with-param name ="shapeCount" select="$var_pos" />
                    </xsl:call-template>

                  </xsl:when>
                  <xsl:when test="./draw:image">
                    <xsl:for-each select="./draw:image">
                      <xsl:if test ="contains(@xlink:href,'.png') or contains(@xlink:href,'.emf') or contains(@xlink:href,'.wmf') or contains(@xlink:href,'.jfif') or contains(@xlink:href,'.jpe') 
            or contains(@xlink:href,'.bmp') or contains(@xlink:href,'.dib') or contains(@xlink:href,'.rle')
            or contains(@xlink:href,'.bmz') or contains(@xlink:href,'.gfa') 
            or contains(@xlink:href,'.emz') or contains(@xlink:href,'.wmz') or contains(@xlink:href,'.pcz')
            or contains(@xlink:href,'.tif') or contains(@xlink:href,'.tiff') 
            or contains(@xlink:href,'.cdr') or contains(@xlink:href,'.cgm') or contains(@xlink:href,'.eps') 
            or contains(@xlink:href,'.pct') or contains(@xlink:href,'.pict') or contains(@xlink:href,'.wpg') 
            or contains(@xlink:href,'.jpeg') or contains(@xlink:href,'.gif') or contains(@xlink:href,'.png') or contains(@xlink:href,'.jpg')">
                        <xsl:if test="not(./@xlink:href[contains(.,'../')])">
                          <xsl:call-template name="InsertPicture">
                            <xsl:with-param name="imageNo" select="$pageNo" />
                            <xsl:with-param name="picNo" select="$var_pos" />
                            <xsl:with-param name="fileName" select="'content.xml'" />
                            <xsl:with-param name ="grStyle" >
                              <xsl:choose>
                                <xsl:when test="./parent::node()/@draw:style-name">
                                  <xsl:value-of select ="./parent::node()/@draw:style-name"/>
                                </xsl:when>
                                <xsl:when test="./parent::node()/@presentation:style-name">
                                  <xsl:value-of select ="./parent::node()/@presentation:style-name"/>
                                </xsl:when>
                              </xsl:choose>
                            </xsl:with-param >
                          </xsl:call-template>
                        </xsl:if>
                      </xsl:if>
                    </xsl:for-each>
                  </xsl:when>
                  <xsl:when test="@presentation:class[contains(.,'title')]
                                  or @presentation:class[contains(.,'subtitle')]
                                  or (@presentation:style-name and not(@presentation:class))
                                  or @presentation:class[contains(.,'outline')]">
                    <xsl:variable name ="masterPageName" select ="./parent::node()/@draw:master-page-name"/>
                    <xsl:variable name="FrameCount" select="concat('Frame',$var_pos)"/>
                    <p:sp>
                      <p:nvSpPr>
                        <xsl:choose>
                          <xsl:when test ="@presentation:class='title'">
                            <p:cNvPr name="Title 1">
                              <xsl:attribute name="id">
                                <xsl:value-of select="$var_pos+1"/>
                              </xsl:attribute>
                            </p:cNvPr>
                            <p:cNvSpPr>
                              <a:spLocks noGrp="1"/>
                            </p:cNvSpPr>
                            <p:nvPr>
                              <p:ph type="ctrTitle"/>
                            </p:nvPr>
                          </xsl:when>
                          <xsl:when test ="@presentation:class[contains(.,'outline')]">
                            <p:cNvPr name="Content Placeholder 2">
                              <xsl:attribute name="id">
                                <xsl:value-of select="$var_pos+1"/>
                              </xsl:attribute>
                            </p:cNvPr>
                            <p:cNvSpPr>
                              <a:spLocks noGrp="1"/>
                            </p:cNvSpPr>
                            <p:nvPr>
                    <!--Defect Fixed 1956085	 ODP: Outline content loss on outline view  -->

                              <p:ph type="body" idx="1"/>
                            </p:nvPr>
                          </xsl:when>
                          <xsl:when test ="@presentation:class[contains(.,'subtitle')] or @presentation:style-name ">
                            <p:cNvPr name="subTitle 1">
                              <xsl:attribute name ="id">
                                <xsl:value-of select ="$var_pos+1"/>
                              </xsl:attribute>
                            </p:cNvPr>
                            <p:cNvSpPr>
                              <a:spLocks noGrp="1"/>
                            </p:cNvSpPr>
                            <p:nvPr>
                              <p:ph type="subTitle" idx="1"/>
                            </p:nvPr>
                          </xsl:when>

                        </xsl:choose>
                      </p:nvSpPr>
                      <p:spPr>
                        <xsl:call-template name ="tmpdrawCordinates"/>
                        <a:prstGeom prst="rect">
                          <a:avLst />
                        </a:prstGeom>

                        <!--End-->
                        <!-- Solid fill color -->
                        <xsl:call-template name ="fillColor" >
                          <xsl:with-param name ="prId" select ="@presentation:style-name" />
                          <xsl:with-param name ="var_pos" select ="$var_pos" />
                        </xsl:call-template>
                        <!-- added by Vipul to insert line style-->
                        <!--start-->
                        <xsl:variable name="var_PrStyleId" select="@presentation:style-name"/>
                        <xsl:for-each select ="/office:document-content/office:automatic-styles/style:style[@style:name=$var_PrStyleId]/style:graphic-properties">
                          <xsl:if test="position()=1">
                          <xsl:call-template name ="getFillColor"/>
                            <xsl:call-template name ="getLineStyle">
                              <xsl:with-param name="parentStyle" select="../@style:parent-style-name"/>
                            </xsl:call-template>
                          </xsl:if>
                        </xsl:for-each>
                        <!--End-->
                      </p:spPr>
                      <p:txBody>
                        <xsl:call-template name ="TextAlignment" >

                          <xsl:with-param name ="prId" select ="@presentation:style-name"/>
                        </xsl:call-template >
                        <a:lstStyle/>
                        <xsl:call-template name ="processText" >
                          <xsl:with-param name ="layoutName" select ="@presentation:class"/>
                          <xsl:with-param name ="FrameCount" select ="$FrameCount"/>
                          <xsl:with-param name="ShapePos" select="$var_pos"/>
                          <!-- Paremeter added by vijayeta,get master page name, dated:13-7-07-->
                          <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                        </xsl:call-template >
                      </p:txBody>
                    </p:sp>
                  </xsl:when>
                  <xsl:when test ="(draw:text-box) and not(@presentation:style-name) and not(@presentation:class)">
                    <xsl:call-template name ="CreateShape">
                      <xsl:with-param name ="fileName" select ="'content.xml'"/>
                      <xsl:with-param name ="shapeName" select="'TextBox '" />
                      <xsl:with-param name ="shapeCount" select="$var_pos" />
                    </xsl:call-template>
                  </xsl:when>
                  <!--Sounds and Movies-->
                  <xsl:when test="./draw:plugin">
                    <xsl:for-each select="./draw:plugin">
                      <xsl:call-template name="InsertAudio">
                        <xsl:with-param name="imageNo" select="$pageNo" />
                        <xsl:with-param name="AudNo" select="$var_pos" />
                      </xsl:call-template>
                    </xsl:for-each>
                  </xsl:when>
                </xsl:choose>
          
              </xsl:when>
              <xsl:when test="name()='draw:g'">
              <xsl:variable name="var_pos">
                <xsl:call-template name="getShapePosTemp">
                  <xsl:with-param name="var_pos" select="position()"/>
                </xsl:call-template>
              </xsl:variable>
                <xsl:call-template name="tmpGroping">
                  <xsl:with-param name="pageNo" select="$pageNo"/>
                  <xsl:with-param name="startPos" select="'1'"/>
                  <xsl:with-param name="pos" select="$var_pos"/>
                  <xsl:with-param name="fileName" select="'content.xml'"/>
                  <xsl:with-param name="UniqueId" select="generate-id()"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="name()='draw:rect' or name()='draw:ellipse' or name()='draw:custom-shape'
              or name()='draw:line' or name()='draw:connector' or name()='draw:circle'">
                <!-- Code for shapes start-->
              <xsl:variable name="var_pos">
                <xsl:call-template name="getShapePosTemp">
                  <xsl:with-param name="var_pos" select="position()"/>
                  </xsl:call-template>
                </xsl:variable>
                <!--<xsl:for-each select=".">-->
                <xsl:call-template name ="shapes" >
                  <xsl:with-param name ="fileName" select ="'content.xml'"/>
                  <xsl:with-param name ="var_pos" select="$var_pos" />
                  <!-- Code for shapes End-->
                </xsl:call-template >
                <!--</xsl:for-each>-->
              </xsl:when>
            </xsl:choose>
          </xsl:for-each>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
	  <xsl:call-template name="slideTransition"/>
      <xsl:call-template name ="customAnimation"/>
    </p:sld>
  </xsl:template>
  <xsl:template name="getShapePosTemp">
    <xsl:param name="var_pos"/>
    <xsl:variable name="Cnt">
      <xsl:choose>
        <xsl:when test="preceding-sibling::draw:g">
          <xsl:for-each select="preceding-sibling::draw:g">
            <xsl:value-of select=" concat( 
                                    count(descendant::draw:frame)
                                    + count(descendant::draw:g) 
                                    + count(descendant::draw:rect) 
                                    + count(descendant::draw:ellipse) 
                                    + count(descendant::draw:line) 
                                    + count(descendant::draw:connector) 
                                    + count(descendant::draw:frame/draw:object) 
                                    + count(descendant::draw:frame/draw:object-ole) 
                                    + count(descendant::draw:circle) 
                                    + count(descendant::draw:custom-shape),':')"/>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>      
      <xsl:when test="$Cnt!=''">
        <xsl:call-template name="addVal">
          <xsl:with-param name="strVal" select="$Cnt"/>
          <xsl:with-param name="var_pos" select="$var_pos"/>
      </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$var_pos"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="addVal">
    <xsl:param name="strVal"/>
    <xsl:param name="intSum" select="0"/>
    <xsl:param name="var_pos"/>

    <xsl:choose>
      <xsl:when test="$strVal != ''">
        <xsl:choose>
          <xsl:when test="substring-before($strVal,':') = ''">
            <xsl:value-of select="number($intSum) + number($strVal) + $var_pos"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:variable name="intSubTotal">
              <xsl:value-of select="number($intSum) + number(substring-before($strVal,':'))"/>
            </xsl:variable>
            <xsl:call-template name="addVal">
              <xsl:with-param name="strVal" select="substring-after($strVal,':')"/>
              <xsl:with-param name="intSum" select="$intSubTotal"/>
              <xsl:with-param name="var_pos" select="$var_pos"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>        
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$intSum + $var_pos"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
 
  <xsl:template name ="processText" >
    <xsl:param name ="layoutName"/>
    <xsl:param name ="FrameCount"/>
    <xsl:param name ="ShapePos"/>
    <!-- Paremeter added by vijayeta,get master page name, dated:13-7-07-->
    <xsl:param name ="masterPageName"/>
    <xsl:variable name ="defFontSize">
      <xsl:value-of  select ="office:document-styles/office:styles/style:style/style:text-properties/@fo:font-size"/>
    </xsl:variable>
    <!-- Added by lohith.ar - Start - Mouse click events -->
    <xsl:variable name="PostionCount">
      <xsl:value-of select="position()"/>
    </xsl:variable>
    <!-- Start - variable to set the hyperlinks values for Frames-->

    <xsl:variable name="varRprHyperLinks">
      <xsl:if test="office:event-listeners">
        <xsl:call-template name="tmpOfficeListner">
          <xsl:with-param name="ShapeType" select="'AtchFileId'"/>
          <xsl:with-param name="PostionCount" select="generate-id()"/>
                      </xsl:call-template>
                    </xsl:if>
      </xsl:variable>
    <xsl:variable name="varFrameHyperLinks">
      <xsl:if test="office:event-listeners/presentation:event-listener[contains(@script:event-name,'dom:click')]">
            <a:endParaRPr lang="en-US" dirty="0">
          <xsl:call-template name="tmpOfficeListner">
            <xsl:with-param name="ShapeType" select="'AtchFileId'"/>
            <xsl:with-param name="PostionCount" select="generate-id()"/>
                        </xsl:call-template>
                        </a:endParaRPr>
              </xsl:if>
         </xsl:variable>
    <!-- End - variable to set the hyperlinks values for Frames-->
    <xsl:variable name ="prClsName">
      <xsl:value-of select ="@presentation:class"/>
    </xsl:variable>
    <xsl:choose >
      <xsl:when test ="draw:text-box/text:p/text:span">
        <xsl:for-each select ="draw:text-box">
          <!--<xsl:for-each select ="draw:text-box/text:p">-->
          <xsl:for-each select ="child::node()[position()]">
            <xsl:choose >
              <xsl:when test ="name()='text:p'">
                <xsl:variable name="linkCount">
                  <xsl:value-of select="position()"/>
                </xsl:variable>
                <xsl:variable name ="paraId" select ="@text:style-name"/>
                <xsl:variable name ="isNumberingEnabled">
                  <xsl:choose >
                    <xsl:when test ="//style:style[@style:name=$paraId]/style:paragraph-properties/@text:enable-numbering">
                      <xsl:value-of select ="//style:style[@style:name=$paraId]/style:paragraph-properties/@text:enable-numbering"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select ="'false'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:variable>
                <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <xsl:call-template name ="paraProperties" >
                    <xsl:with-param name ="paraId" >
                      <xsl:value-of select ="@text:style-name"/>
                    </xsl:with-param >
                    <xsl:with-param name ="isBulleted" select ="'false'"/>
                    <xsl:with-param name ="level" select ="'0'"/>
                    <xsl:with-param name ="BuImgRel" select ="concat($ShapePos,$linkCount)"/>
                    <xsl:with-param name="framePresentaionStyleId" select="parent::node()/parent::node()/./@presentation:style-name" />
                    <xsl:with-param name ="isNumberingEnabled" select ="$isNumberingEnabled"/>
                    <xsl:with-param name ="prClsName" select ="$prClsName"/>
                    <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                  </xsl:call-template >
                  <xsl:for-each select ="child::node()[position()]">
                    <xsl:choose >
                      <xsl:when test ="name()='text:span'">
                        <xsl:variable name="pos">
                          <xsl:value-of select="position()"/>
                        </xsl:variable>
                        <!-- Added by lohith - bug fix 1731885 -->
                        <xsl:if test="node()">
                          <a:r>
                            <a:rPr smtClean="0">
                              <!--Font Size -->
                              <xsl:variable name ="textId">
                                <xsl:value-of select ="@text:style-name"/>
                              </xsl:variable>
                              <xsl:if test ="not($textId ='')">
                                <xsl:call-template name ="fontStyles">
                                  <xsl:with-param name ="Tid" select ="$textId" />
                                  <xsl:with-param name ="prClassName" select ="$prClsName"/>
                                  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                                </xsl:call-template>
                              </xsl:if>
                              <xsl:choose>
                                <xsl:when test="parent::node()/parent::node()/parent::node()/office:event-listeners/presentation:event-listener">
                                  <xsl:copy-of select="$varRprHyperLinks"/>
                                </xsl:when>
                                <xsl:when test="./text:a">
                                  <xsl:call-template name="TextActionHyperLinks">
                                    <xsl:with-param name="PostionCount" select="$PostionCount"/>
                                    <xsl:with-param name="nodeType" select="'span'"/>
                                    <xsl:with-param name="linkCount" select="$linkCount"/>
                                    <xsl:with-param name="pos" select="$pos"/>
                                    <xsl:with-param name="ShapePos" select="$ShapePos"/>
                                   </xsl:call-template>
                                </xsl:when>
                              </xsl:choose>
                            </a:rPr >
                            <a:t>
                              <xsl:call-template name ="insertTab" />
                            </a:t>
                          </a:r>
                          <xsl:if test ="text:line-break">
                            <xsl:call-template name ="processBR">
                              <xsl:with-param name ="T" select ="@text:style-name" />
                              <xsl:with-param name ="prClassName" select ="$prClsName"/>
                            </xsl:call-template>
                          </xsl:if>
                        </xsl:if>
                        <xsl:if test="not(node()) and 
                                     not(parent::node()/parent::node()/parent::node()/office:event-listeners/presentation:event-listener)">
                          <a:endParaRPr dirty="0" smtClean="0">
                            <!--Font Size -->
                            <xsl:variable name ="textId">
                              <xsl:value-of select ="@text:style-name"/>
                            </xsl:variable>
                            <xsl:if test ="not($textId ='')">
                              <xsl:call-template name ="fontStyles">
                                <xsl:with-param name ="Tid" select ="$textId" />
                                <xsl:with-param name ="prClassName" select ="$prClsName"/>
                                <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                              </xsl:call-template>
                            </xsl:if>
                          </a:endParaRPr>
                        </xsl:if>
                      </xsl:when >
                      <xsl:when test ="name()='text:line-break'">
                        <xsl:call-template name ="processBR">
                          <xsl:with-param name ="T" select ="@text:style-name" />
                          <xsl:with-param name ="prClassName" select ="$prClsName"/>
                        </xsl:call-template>
                      </xsl:when>
                      <xsl:when test ="not(name()='text:span')">
                        <a:r>
                          <a:rPr smtClean="0">
                            <!--Font Size -->
                            <xsl:variable name ="textId">
                              <xsl:value-of select ="@text:style-name"/>
                            </xsl:variable>
                            <xsl:if test ="not($textId ='')">
                              <xsl:call-template name ="fontStyles">
                                <xsl:with-param name ="Tid" select ="$textId" />
                                <xsl:with-param name ="prClassName" select ="$prClsName"/>
                                <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                              </xsl:call-template>
                            </xsl:if>
                            <xsl:choose>
                              <xsl:when test="parent::node()/parent::node()/parent::node()/office:event-listeners/presentation:event-listener">
                                <xsl:copy-of select="$varRprHyperLinks"/>
                              </xsl:when>
                              <xsl:when test="./text:a">
                                <xsl:call-template name="TextActionHyperLinks">
                                  <xsl:with-param name="PostionCount" select="$PostionCount"/>
                                  <xsl:with-param name="nodeType" select="'span'"/>
                                  <xsl:with-param name="linkCount" select="$linkCount"/>
                                  <xsl:with-param name="ShapePos" select="$ShapePos"/>
                                </xsl:call-template>
                              </xsl:when>
                            </xsl:choose>
                            <!--<xsl:copy-of select="$varRprHyperLinks"/>
                            <xsl:if test="./text:a">
                              <xsl:call-template name="TextActionHyperLinks">
                                <xsl:with-param name="PostionCount" select="$PostionCount"/>
                                <xsl:with-param name="nodeType" select="'span'"/>
                                <xsl:with-param name="linkCount" select="$linkCount"/>
                                <xsl:with-param name="ShapePos" select="$ShapePos"/>
                              </xsl:call-template>
														</xsl:if>-->
                          </a:rPr >
                          <a:t>
                            <xsl:call-template name ="insertTab" />
                          </a:t>
                        </a:r>
                      </xsl:when >
                    </xsl:choose>
                  </xsl:for-each>
                  <xsl:copy-of select="$varFrameHyperLinks"/>
                </a:p >
              </xsl:when>
              <xsl:when test ="name()='text:list'">
                <xsl:variable name="linkCount">
                  <xsl:value-of select="position()"/>
                </xsl:variable>
                <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <xsl:variable name ="lvl">
                    <xsl:if test ="text:list-item/text:list">
                      <xsl:call-template name ="getListLevel">
                        <xsl:with-param name ="levelCount"/>
                      </xsl:call-template>
                    </xsl:if>
                    <xsl:if test ="not(text:list-item/text:list)">
                      <xsl:value-of select ="'0'"/>
                    </xsl:if>
                  </xsl:variable >
                  <xsl:variable name="paragraphId" >
                    <xsl:call-template name ="getParaStyleName">
                      <xsl:with-param name ="lvl" select ="$lvl"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:variable name ="isNumberingEnabled">
                    <xsl:if test ="//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering">
                      <xsl:value-of select ="//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering"/>
                    </xsl:if>
                    <!--<xsl:if test ="not(document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering)">
											<xsl:value-of select ="'true'"/>
										</xsl:if>-->
                  </xsl:variable>
                  <xsl:call-template name ="paraProperties" >
                    <xsl:with-param name ="paraId" >
                      <xsl:value-of select ="$paragraphId"/>
                    </xsl:with-param >
                    <!-- list property also included-->
                    <xsl:with-param name ="listId">
                      <xsl:value-of select ="@text:style-name"/>
                    </xsl:with-param >
                    <xsl:with-param name ="BuImgRel" select ="concat($ShapePos,$linkCount)"/>
                    <!-- Parameters added by vijayeta,Set bulleting as true/false,and set level -->
                    <xsl:with-param name ="isBulleted" select ="'true'"/>
                    <xsl:with-param name ="level" select ="$lvl"/>
                    <xsl:with-param name ="isNumberingEnabled" select ="$isNumberingEnabled"/>
                    <!-- parameter added by vijayeta, dated 11-7-07-->
                    <xsl:with-param name ="prClsName" select ="$prClsName"/>
                    <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                  </xsl:call-template >
                  <!--End of Code inserted by Vijayets for Bullets and numbering-->
                  <xsl:for-each select ="child::node()[position()]">
                    <xsl:choose >
                      <xsl:when test ="name()='text:list-item'">
                        <xsl:variable name ="currentNodeStyle">
                          <xsl:call-template name ="getTextNodeForFontStyle">
                            <xsl:with-param name ="prClassName" select ="$prClsName"/>
                            <xsl:with-param name ="lvl" select ="$lvl"/>
                            <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                          </xsl:call-template>
                        </xsl:variable>
                        <xsl:copy-of select ="$currentNodeStyle"/>
                      </xsl:when >
                    </xsl:choose>
                  </xsl:for-each>
                  <xsl:copy-of select="$varFrameHyperLinks"/>
                </a:p >
              </xsl:when>
            </xsl:choose>
          </xsl:for-each>
        </xsl:for-each >
      </xsl:when>
      <xsl:when test ="draw:text-box/text:list/text:list-item">
        <!--<xsl:when test ="draw:text-box/text:list/text:list-item/text:p/text:span">-->
        <xsl:for-each select ="draw:text-box/text:list">
          <xsl:variable name="forCount" select="position()" />
          <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <!--Code inserted by Vijayets for Bullets and numbering-->
            <!--Check if Levels are present-->
            <xsl:variable name ="lvl">
              <xsl:if test ="text:list-item/text:list">
                <xsl:call-template name ="getListLevel">
                  <xsl:with-param name ="levelCount"/>
                </xsl:call-template>
              </xsl:if>
              <xsl:if test ="not(text:list-item/text:list)">
                <xsl:value-of select ="'0'"/>
              </xsl:if>
            </xsl:variable >
            <xsl:variable name="paragraphId" >
              <xsl:call-template name ="getParaStyleName">
                <xsl:with-param name ="lvl" select ="$lvl"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name ="isNumberingEnabled">
              <xsl:choose >
                <xsl:when test ="//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering">
                  <xsl:value-of select ="//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:variable name ="styleNameFromStyles" >
                    <xsl:choose >
                      <xsl:when test ="$prClsName='subtitle' or $prClsName='title'">
                        <xsl:value-of select ="concat($masterPageName,'-',$prClsName)"/>
                      </xsl:when>
                      <xsl:when test ="$prClsName='outline'">
                        <xsl:value-of select ="concat($masterPageName,'-outline',$lvl+1)"/>
                      </xsl:when>
                    </xsl:choose>
                  </xsl:variable>
                  <xsl:choose >
                    <xsl:when test ="document('styles.xml')//office:styles/style:style[@style:name=$styleNameFromStyles]/style:paragraph-properties/@text:enable-numbering">
                      <xsl:value-of select ="document('styles.xml')//office:styles/style:style[@style:name=$styleNameFromStyles]/style:paragraph-properties/@text:enable-numbering"/>
                    </xsl:when>
                    <xsl:when test ="not(document('styles.xml')//office:styles/style:style[@style:name=$styleNameFromStyles]/style:paragraph-properties/@text:enable-numbering)">
                      <xsl:value-of select ="'true'"/>
                    </xsl:when>
                    <!--<xsl:otherwise>
                      <xsl:value-of select ="'true'"/>
                    </xsl:otherwise>-->
                  </xsl:choose>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:call-template name ="paraProperties" >
              <xsl:with-param name ="paraId" >
                <xsl:value-of select ="$paragraphId"/>
              </xsl:with-param >
              <!-- list property also included-->
              <xsl:with-param name ="listId">
                <xsl:value-of select ="@text:style-name"/>
              </xsl:with-param >
              <xsl:with-param name ="BuImgRel" select ="concat($ShapePos,$forCount)"/>
              <!-- Parameters added by vijayeta,Set bulleting as true/false,and set level -->
              <xsl:with-param name ="isBulleted" select ="'true'"/>
              <xsl:with-param name ="level" select ="$lvl"/>
              <xsl:with-param name ="isNumberingEnabled" select ="$isNumberingEnabled"/>
              <!-- parameter added by vijayeta, dated 13-7-07-->
              <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
              <xsl:with-param name ="pos" select ="$forCount"/>
              <xsl:with-param name ="FrameCount" select ="substring-after($FrameCount,'Frame')"/>
              <xsl:with-param name ="prClsName" select ="$prClsName"/>
            </xsl:call-template >
            <!--End of Code inserted by Vijayets for Bullets and numbering-->
            <xsl:for-each select ="child::node()[position()]">
              <xsl:choose >
                <xsl:when test ="name()='text:list-item'">
                  <!-- Variable "textHyperlinksForBullets" added by lohith for Hyperlinks in Bullets -->
                  <xsl:variable name="textHyperlinksForBullets">
                    <xsl:call-template name="getHyperlinksForBulltedText">
                      <xsl:with-param name="bulletLevel" select="$lvl" />
                      <xsl:with-param name="listItemCount" select="$forCount"/>
                      <xsl:with-param name="FrameCount" select="$FrameCount"/>
                      <xsl:with-param name="rootNode" select="."/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:variable name ="currentNodeStyle">
                    <xsl:call-template name ="getTextNodeForFontStyle">
                      <xsl:with-param name ="prClassName" select ="$prClsName"/>
                      <xsl:with-param name ="lvl" select ="$lvl"/>
                      <xsl:with-param name ="HyperlinksForBullets" select ="$textHyperlinksForBullets"/>
                      <!-- parameter added by vijayeta, dated 11-7-07-->
                      <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                      <xsl:with-param name ="fileName" select ="'content.xml'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:copy-of select ="$currentNodeStyle"/>
                </xsl:when >
              </xsl:choose>
            </xsl:for-each>
            <xsl:copy-of select="$varFrameHyperLinks"/>
          </a:p >
        </xsl:for-each >
      </xsl:when>
      <xsl:when test ="draw:text-box/text:list/text:list-item/text:p">
        <xsl:for-each select ="draw:text-box/text:list">
          <xsl:variable name="forCount" select="position()" />
          <a:p  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <!--Check if Levels are present-->
            <xsl:variable name ="lvl">
              <xsl:if test ="text:list-item/text:list">
                <xsl:call-template name ="getListLevel">
                  <xsl:with-param name ="levelCount"/>
                </xsl:call-template>
              </xsl:if>
              <xsl:if test ="not(text:list-item/text:list)">
                <xsl:value-of select ="'0'"/>
              </xsl:if>
            </xsl:variable >
            <xsl:variable name="paragraphId" >
              <xsl:call-template name ="getParaStyleName">
                <xsl:with-param name ="lvl" select ="$lvl"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name ="isNumberingEnabled">
              <xsl:if test ="//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering">
                <xsl:value-of select ="//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering"/>
              </xsl:if>
              <xsl:if test ="not(//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering)">
                <xsl:value-of select ="'true'"/>
              </xsl:if>
            </xsl:variable>
            <xsl:call-template name ="paraProperties" >
              <xsl:with-param name ="paraId" >
                <xsl:value-of select ="$paragraphId"/>
              </xsl:with-param >
              <!-- list property also included-->
              <xsl:with-param name ="listId">
                <xsl:value-of select ="@text:style-name"/>
              </xsl:with-param >
              <xsl:with-param name ="BuImgRel" select ="concat($ShapePos,$forCount)"/>
              <!-- Parameters added by vijayeta,Set bulleting as true/false,and set level -->
              <xsl:with-param name ="isBulleted" select ="'true'"/>
              <xsl:with-param name ="level" select ="$lvl"/>
              <xsl:with-param name ="isNumberingEnabled" select ="$isNumberingEnabled"/>
              <!-- parameter added by vijayeta, dated 11-7-07-->
              <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
              <xsl:with-param name ="prClsName" select ="$prClsName"/>
            </xsl:call-template >
            <!--End of Code inserted by Vijayets for Bullets and numbering-->
            <xsl:for-each select ="child::node()[position()]">
              <xsl:choose >
                <xsl:when test ="name()='text:list-item'">
                  <xsl:variable name ="currentNodeStyle">
                    <xsl:call-template name ="getTextNodeForFontStyle">
                      <xsl:with-param name ="prClassName" select ="$prClsName"/>
                      <xsl:with-param name ="lvl" select ="$lvl"/>
                      <!-- parameter added by vijayeta, dated 11-7-07-->
                      <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                      <xsl:with-param name ="fileName" select ="'content.xml'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:copy-of select ="$currentNodeStyle"/>
                </xsl:when >
              </xsl:choose>
            </xsl:for-each>
            <xsl:copy-of select="$varFrameHyperLinks"/>
          </a:p >
        </xsl:for-each >
      </xsl:when>
      <xsl:when test ="draw:text-box/text:p">
        <xsl:for-each select ="draw:text-box/text:p">
          <xsl:variable name="linkCount">
            <xsl:value-of select="position()"/>
          </xsl:variable>
          <xsl:variable name ="paraId" select ="@text:style-name"/>
          <xsl:variable name ="isNumberingEnabled">
            <xsl:choose >
              <xsl:when test ="//style:style[@style:name=$paraId]/style:paragraph-properties/@text:enable-numbering">
                <xsl:value-of select ="//style:style[@style:name=$paraId]/style:paragraph-properties/@text:enable-numbering"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select ="'false'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <xsl:call-template name ="paraProperties" >
              <xsl:with-param name ="paraId" >
                <xsl:value-of select ="@text:style-name"/>
              </xsl:with-param >
              <xsl:with-param name ="isBulleted" select ="'false'"/>
              <xsl:with-param name ="level" select ="'0'"/>
              <xsl:with-param name ="BuImgRel" select ="concat($ShapePos,$linkCount)"/>
              <xsl:with-param name="framePresentaionStyleId" select="parent::node()/parent::node()/./@presentation:style-name" />
              <xsl:with-param name ="isNumberingEnabled" select ="$isNumberingEnabled"/>
              <xsl:with-param name ="prClsName" select ="$prClsName"/>
              <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
            </xsl:call-template >
            <a:r>
              <a:rPr smtClean="0">
                <!--Font Size -->
                <xsl:variable name ="textId">
                  <xsl:value-of select ="@text:style-name"/>
                </xsl:variable>
                <xsl:if test ="not($textId ='')">
                  <xsl:call-template name ="fontStyles">
                    <xsl:with-param name ="Tid" select ="$textId" />
                    <xsl:with-param name ="prClassName" select ="$prClsName"/>
                    <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                  </xsl:call-template>
                </xsl:if>
                <xsl:choose>
                  <xsl:when test="parent::node()/parent::node()/office:event-listeners/presentation:event-listener">
                    <xsl:copy-of select="$varRprHyperLinks"/>
                  </xsl:when>
                  <xsl:when test="./text:a">
                    <xsl:call-template name="TextActionHyperLinks">
                      <xsl:with-param name="PostionCount" select="$PostionCount"/>
                      <xsl:with-param name="nodeType" select="'span'"/>
                      <xsl:with-param name="linkCount" select="$linkCount"/>
                      <xsl:with-param name="ShapePos" select="$ShapePos"/>
                    </xsl:call-template>
                  </xsl:when>
                </xsl:choose>
              </a:rPr >
              <a:t>
                <xsl:call-template name ="insertTab" />
              </a:t>
            </a:r>
            <xsl:if test ="text:line-break">
              <xsl:call-template name ="processBR">
                <xsl:with-param name ="T" select ="@text:style-name" />
                <xsl:with-param name ="prClassName" select ="$prClsName"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:copy-of select="$varFrameHyperLinks"/>
          </a:p >
        </xsl:for-each >
      </xsl:when>
      <xsl:otherwise >
        <xsl:variable name="linkCount">
          <xsl:value-of select="position()"/>
        </xsl:variable>
        <a:p>
          <a:r >
            <a:rPr smtClean="0">
              <a:latin charset="0" typeface="Arial" />
              <xsl:choose>
                <xsl:when test="./text:a">
                  <xsl:call-template name="TextActionHyperLinks">
                    <xsl:with-param name="PostionCount" select="$PostionCount"/>
                    <xsl:with-param name="nodeType" select="'span'"/>
                    <xsl:with-param name="linkCount" select="$linkCount"/>
                    <xsl:with-param name="ShapePos" select="$ShapePos"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:copy-of select="$varRprHyperLinks"/>
                </xsl:otherwise>
              </xsl:choose>
              <!--<xsl:copy-of select="$varRprHyperLinks"/>
              <xsl:if test="./text:a">
                <xsl:call-template name="TextActionHyperLinks">
                  <xsl:with-param name="PostionCount" select="$PostionCount"/>
                  <xsl:with-param name="nodeType" select="'span'"/>
                  <xsl:with-param name="linkCount" select="$linkCount"/>
                  <xsl:with-param name="ShapePos" select="$ShapePos"/>
                </xsl:call-template>
							</xsl:if>-->
            </a:rPr >
            <a:t>
              <xsl:call-template name ="insertTab" />
            </a:t>
          </a:r>
          <xsl:copy-of select="$varFrameHyperLinks"/>
        </a:p >
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>  
  <xsl:template name ="footer">
    <xsl:param name ="footerId"></xsl:param>
    <xsl:param name ="SMName"/>

    <xsl:for-each select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name=$SMName]/draw:frame[@presentation:class='footer']">
      <xsl:variable name="footerValue">
        <xsl:choose>
          <xsl:when test="draw:text-box/text:p//presentation:footer">
            <xsl:value-of select="'false'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'true'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:if test="$footerValue='true' or document('content.xml')//presentation:footer-decl[@presentation:name=$footerId] ">
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="16361" name="Footer Placeholder 5" />
            <p:cNvSpPr>
              <a:spLocks noGrp="1" />
            </p:cNvSpPr>
            <p:nvPr>
              <p:ph type="ftr" sz="quarter" idx="11" />
            </p:nvPr>
          </p:nvSpPr>
          <p:spPr >
            <!-- footer date layout details style:master-page style:name -->
            <xsl:if test="document('content.xml') //presentation:footer-decl[@presentation:name=$footerId]">
              <xsl:call-template name ="GetFrameDetails">
                <xsl:with-param name ="LayoutName" select ="'footer'"/>
              </xsl:call-template>
            </xsl:if>
          </p:spPr >
          <p:txBody>
            <a:bodyPr />
            <a:lstStyle />
            <a:p>
              <xsl:if test ="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name=$SMName]/draw:frame[@presentation:class='footer']">
                <xsl:for-each select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name=$SMName]/draw:frame[@presentation:class='footer']">
                  <xsl:choose>
                    <xsl:when test="document('content.xml') //presentation:footer-decl[@presentation:name=$footerId]">
                      <a:r>
                        <a:rPr dirty="0" smtClean="0"/>
                        <a:t>
                          <xsl:for-each select ="document('content.xml') //presentation:footer-decl[@presentation:name=$footerId] ">
                            <xsl:value-of select ="."/>
                          </xsl:for-each >
                        </a:t>
                      </a:r>
                      <a:endParaRPr dirty="0"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:if test="not(./draw:text-box/text:p/text:span)">
                        <a:endParaRPr/>
                      </xsl:if>
                      <xsl:if test="./draw:text-box/text:p/text:span">
                        <xsl:for-each select="./draw:text-box/text:p/text:span">
                          <a:r>
                            <a:rPr dirty="0" smtClean="0">
                              <!--Font Size-->
                              <xsl:variable name ="textId">
                                <xsl:value-of select ="@text:style-name"/>
                              </xsl:variable>
                              <xsl:if test ="not($textId ='')">
                                <xsl:call-template name ="tmpSMfontStyles">
                                  <xsl:with-param name ="TextStyleID" select ="$textId" />
                                </xsl:call-template>
                              </xsl:if>
                            </a:rPr>
                            <a:t>
                              <xsl:call-template name ="insertTab" />
                            </a:t>
                          </a:r>
                        </xsl:for-each>
                        <a:endParaRPr dirty="0"/>
                      </xsl:if>

                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:for-each>
              </xsl:if>
            </a:p>
          </p:txBody>
        </p:sp >
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="slideNumber">
    <xsl:param name ="pageNumber"/>
    <xsl:param name ="SMName"/>
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="16362" name="Slide Number Placeholder 4" />
        <p:cNvSpPr>
          <a:spLocks noGrp="1" />
        </p:cNvSpPr>
        <p:nvPr>
          <p:ph type="sldNum" sz="quarter" idx="12" />
        </p:nvPr>
      </p:nvSpPr>
      <p:spPr >
        <!-- footer slide number layout details style:master-page style:name -->
        <xsl:call-template name ="GetFrameDetails">
          <xsl:with-param name ="LayoutName" select ="'page-number'"/>
        </xsl:call-template>
      </p:spPr >
      <p:txBody>
        <a:bodyPr />
        <a:lstStyle />
        <xsl:choose>
          <xsl:when test="document('styles.xml')//style:master-page[@style:name=$SMName]/draw:frame[@presentation:class='page-number']">
        <xsl:for-each select ="document('styles.xml')//style:master-page[@style:name=$SMName]/draw:frame[@presentation:class='page-number']">
          <xsl:for-each select="./draw:text-box">
            <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:pPr/>
              <xsl:for-each select ="node()">
                <xsl:if test ="name()='text:p'" >
                  <xsl:variable name="paraId" select="@text:style-name"/>
                  <xsl:if test ="child::node()">
                    <xsl:for-each select ="node()">
                      <xsl:choose >
                        <xsl:when test ="name()='text:span'">
                          <xsl:choose>
                            <xsl:when test ="text:page-number">
                              <a:fld >
            <xsl:attribute name ="id">
                                  <xsl:value-of select ="'{763D1470-AB83-4C4C-B3B3-7F0C9DC8E8D6}'"/>
            </xsl:attribute>
                                <xsl:attribute name ="type">
                                  <xsl:value-of select ="'slidenum'"/>
                                </xsl:attribute>
                                <a:rPr dirty="0" smtClean="0">
                                  <xsl:variable name ="textId">
                                    <xsl:value-of select ="./parent::node()/@text:style-name"/>
                                  </xsl:variable>
                                  <xsl:if test ="not($textId ='')">
                                    <xsl:call-template name ="tmpSMfontStyles">
                                      <xsl:with-param name ="TextStyleID" select ="$textId" />
                                    </xsl:call-template>
                                  </xsl:if>
                                </a:rPr>
                                    <a:t>#</a:t>
                              </a:fld>
                            </xsl:when>
                            <xsl:when test="node()">
                              <a:r>
                                <a:rPr dirty="0" smtClean="0">
                                  <xsl:variable name ="textId">
                                    <xsl:value-of select ="@text:style-name"/>
                                  </xsl:variable>
                                  <xsl:if test ="not($textId ='')">
                                    <xsl:call-template name ="tmpSMfontStyles">
                                      <xsl:with-param name ="TextStyleID" select ="$textId" />
                                    </xsl:call-template>
                                  </xsl:if>
                                </a:rPr>
                                <a:t>
                                  <xsl:call-template name ="insertTab" />
                                </a:t>
                              </a:r >
                            </xsl:when>
                          </xsl:choose>
                        </xsl:when >
                        <xsl:when test ="name()='text:page-number'">
                          <a:fld >
                            <xsl:attribute name ="id">
                              <xsl:value-of select ="'{763D1470-AB83-4C4C-B3B3-7F0C9DC8E8D6}'"/>
                            </xsl:attribute>
                            <xsl:attribute name ="type">
                              <xsl:value-of select ="'slidenum'"/>
                            </xsl:attribute>
                            <a:rPr dirty="0" smtClean="0">
                              <xsl:variable name ="textId">
                                <xsl:value-of select ="./parent::node()/@text:style-name"/>
                              </xsl:variable>
                              <xsl:if test ="not($textId ='')">
                                <xsl:call-template name ="tmpSMfontStyles">
                                  <xsl:with-param name ="TextStyleID" select ="$textId" />
                                </xsl:call-template>
                              </xsl:if>
                            </a:rPr>
                                <a:t>#</a:t>
          </a:fld>
                        </xsl:when>
                      </xsl:choose>
                    </xsl:for-each>
                  </xsl:if>
                </xsl:if>
              </xsl:for-each>
        </a:p>
          </xsl:for-each>
        </xsl:for-each>
          </xsl:when>
          <xsl:otherwise>
            <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
          </xsl:otherwise>
        </xsl:choose>
     
    
      </p:txBody>
    </p:sp >
  </xsl:template>
  <xsl:template name ="footerDate">
    <xsl:param name ="footerDateId"/>
    <xsl:param name ="SMName"/>

    
      
          <xsl:choose>
            <xsl:when test="//presentation:date-time-decl[@presentation:name=$footerDateId]">
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="16363" name="Date Placeholder 3" />
        <p:cNvSpPr>
          <a:spLocks noGrp="1" />
        </p:cNvSpPr>
        <p:nvPr>
          <p:ph type="dt" sz="half" idx="10" />
        </p:nvPr>
      </p:nvSpPr>
      <p:spPr >
        <!-- footer layout details style:master-page style:name -->
        <xsl:call-template name ="GetFrameDetails">
          <xsl:with-param name ="LayoutName" select ="'date-time'"/>
          <xsl:with-param name ="SMName" select ="$SMName"/>
        </xsl:call-template>
      </p:spPr >
      <p:txBody>
        <a:bodyPr />
        <a:lstStyle />
        <a:p>
              <xsl:for-each select ="//presentation:date-time-decl[@presentation:name=$footerDateId] ">
                <xsl:choose>
                  <xsl:when test="@presentation:source='current-date'" >
                    <xsl:choose>
                      <xsl:when test="document('styles.xml')//style:master-page[@style:name=$SMName]/draw:frame[@presentation:class='date-time']/draw:text-box/text:p//presentation:date-time">
                    <a:fld >
                      <xsl:attribute name ="id">
                        <xsl:value-of select ="'{86419996-E19B-43D7-A4AA-D671C2F15715}'"/>
                      </xsl:attribute>
                      <xsl:attribute name ="type">
                        <xsl:choose >
                          <xsl:when test ="@style:data-style-name ='D3'">
                            <xsl:value-of select ="'datetime1'"/>
                          </xsl:when>
                          <xsl:when test ="@style:data-style-name ='D8'">
                            <xsl:value-of select ="'datetime2'"/>
                          </xsl:when>
                          <xsl:when test ="@style:data-style-name ='D6'">
                            <xsl:value-of select ="'datetime4'"/>
                          </xsl:when>
                          <xsl:when test ="@style:data-style-name ='D5'">
                            <xsl:value-of select ="'datetime4'"/>
                          </xsl:when>
                          <xsl:when test ="@style:data-style-name ='D3T2'">
                            <xsl:value-of select ="'datetime8'"/>
                          </xsl:when>
                          <xsl:when test ="@style:data-style-name ='D3T5'">
                            <xsl:value-of select ="'datetime8'"/>
                          </xsl:when>
                          <xsl:when test ="@style:data-style-name ='T2'">
                            <xsl:value-of select ="'datetime10'"/>
                          </xsl:when>
                          <xsl:when test ="@style:data-style-name ='T3'">
                            <xsl:value-of select ="'datetime11'"/>
                          </xsl:when>
                          <xsl:when test ="@style:data-style-name ='T5'">
                            <xsl:value-of select ="'datetime12'"/>
                          </xsl:when>
                          <xsl:when test ="@style:data-style-name ='T6'">
                            <xsl:value-of select ="'datetime13'"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select ="'datetime1'"/>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:attribute>
                      <a:rPr smtClean="0" lang="en-US"/>
                      <a:t>
                        <xsl:value-of select ="."/>
                      </a:t>
                    </a:fld>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:for-each select ="document('styles.xml')//style:master-page[@style:name=$SMName]/draw:frame[@presentation:class='date-time']/draw:text-box">
                          <xsl:for-each select ="node()">
                            <xsl:if test ="name()='text:p'" >
                              <xsl:variable name="paraId" select="@text:style-name"/>
                              <xsl:if test ="child::node()">
                                <xsl:for-each select ="node()">
                                  <xsl:choose>
                                    <xsl:when test ="name()='text:span'">
                                      <xsl:choose>
                                        <xsl:when test="node() or text:date">
                                          <a:r>
                                            <a:rPr dirty="0" smtClean="0">
                                              <xsl:variable name ="textId">
                                                <xsl:value-of select ="@text:style-name"/>
                                              </xsl:variable>
                                              <xsl:choose>
                                                <xsl:when test ="$textId !=''">
                                                <xsl:call-template name ="tmpSMfontStyles">
                                                  <xsl:with-param name ="TextStyleID" select ="$textId" />
                                                </xsl:call-template>
                                                </xsl:when>
                                                <xsl:otherwise>
                                                  <xsl:attribute name="lang">
                                                    <xsl:value-of select="'en-US'"/>
                                                  </xsl:attribute>
                                                </xsl:otherwise>
                                              </xsl:choose>
                                            </a:rPr>
                                            <a:t>
                                              <xsl:call-template name ="insertTab" />
                                            </a:t>
                                          </a:r >
                                        </xsl:when>
                                      </xsl:choose>
                                    </xsl:when>
                                    <xsl:when test="not(node())">
                                      <a:r>
                                        <a:rPr dirty="0" smtClean="0">
                                          <xsl:variable name ="textId">
                                            <xsl:value-of select ="@text:style-name"/>
                                          </xsl:variable>
                                          <xsl:choose>
                                            <xsl:when test ="$textId !=''">
                                            <xsl:call-template name ="tmpSMfontStyles">
                                              <xsl:with-param name ="TextStyleID" select ="$textId" />
                                            </xsl:call-template>
                                            </xsl:when>
                                            <xsl:otherwise>
                                              <xsl:attribute name="lang">
                                                <xsl:value-of select="'en-US'"/>
                                              </xsl:attribute>
                                            </xsl:otherwise>
                                          </xsl:choose>
                                      
                                        </a:rPr>
                                        <a:t>
                                          <xsl:call-template name ="insertTab" />
                                        </a:t>
                                      </a:r >
                                    </xsl:when>
                                  </xsl:choose>
                                </xsl:for-each>
                              </xsl:if>
                              <xsl:if test ="not(child::node())">
                                <xsl:value-of select="."/>
                              </xsl:if>
                            </xsl:if>
                          </xsl:for-each>
                        </xsl:for-each>
                      </xsl:otherwise>
                    </xsl:choose>
                
                    <a:endParaRPr lang="en-US" />
                  </xsl:when>
                  <xsl:otherwise >
                    <a:r>
                      <a:rPr lang="en-US" smtClean="0" />
                      <a:t>
                        <xsl:value-of select ="."/>
                      </a:t>
                    </a:r >
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:for-each>
            </a:p>

                </p:txBody>
              </p:sp >
            </xsl:when>
            <xsl:otherwise>
              <xsl:variable name="isDateText">
                <xsl:if test ="document('styles.xml')//style:master-page[@style:name=$SMName]/draw:frame[@presentation:class='date-time']//text:span">
                  <xsl:for-each select ="document('styles.xml')//style:master-page[@style:name=$SMName]/draw:frame[@presentation:class='date-time']//text:span">
                    <xsl:value-of select="."/>
                  </xsl:for-each>
                </xsl:if>
                <xsl:for-each select="document('styles.xml')//style:master-page[@style:name=$SMName]/draw:frame[@presentation:class='date-time']//text:p">
                  <xsl:value-of select="."/>
                </xsl:for-each>
              </xsl:variable>
              <xsl:if test="$isDateText!=''">
              <xsl:for-each select ="document('styles.xml')//style:master-page[@style:name=$SMName]/draw:frame[@presentation:class='date-time']">
                <p:sp>
                  <p:nvSpPr>
                    <p:cNvPr id="16363" name="Date Placeholder 3" />
                    <p:cNvSpPr>
                      <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                      <p:ph type="dt" sz="half" idx="10" />
                    </p:nvPr>
                  </p:nvSpPr>
                  <p:spPr >
                    <!-- footer layout details style:master-page style:name -->
                    <xsl:call-template name ="GetFrameDetails">
                      <xsl:with-param name ="LayoutName" select ="'date-time'"/>
                      <xsl:with-param name ="SMName" select ="$SMName"/>
                    </xsl:call-template>
                  </p:spPr >
                  <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                <xsl:call-template name="tmpDateTimeText"/>
                  </p:txBody>
                </p:sp >
              </xsl:for-each>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpDateTimeText">
    <xsl:for-each select="./draw:text-box">
      <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:pPr/>
        <xsl:for-each select ="node()">
          <xsl:if test ="name()='text:p'" >
            <xsl:variable name="paraId" select="@text:style-name"/>
            <xsl:if test ="child::node()">
              <xsl:for-each select ="node()">
                <xsl:choose>
                  <xsl:when test ="name()='text:span'">
                    <xsl:choose>
                      <xsl:when test ="presentation:date-time">
                    <a:fld >
                      <xsl:attribute name ="id">
                        <xsl:value-of select ="'{11859447-1FF1-4571-A7A1-75F1CAF7F5D9}'"/>
                      </xsl:attribute>
                      <xsl:attribute name ="type">
                        <xsl:value-of select ="'datetime1'"/>
                      </xsl:attribute>
                          <a:rPr dirty="0" smtClean="0">
                            <xsl:variable name ="textId">
                              <xsl:value-of select ="./parent::node()/@text:style-name"/>
                            </xsl:variable>
                            <xsl:if test ="not($textId ='')">
                              <xsl:call-template name ="tmpSMfontStyles">
                                <xsl:with-param name ="TextStyleID" select ="$textId" />
                              </xsl:call-template>
                            </xsl:if>
                          </a:rPr>
                      <a:t> </a:t>
                    </a:fld>
                     </xsl:when>
                      <xsl:when test="node() or text:date">
                    <a:r>
                          <a:rPr dirty="0" smtClean="0">
                            <xsl:variable name ="textId">
                              <xsl:value-of select ="@text:style-name"/>
                            </xsl:variable>
                            <xsl:if test ="not($textId ='')">
                              <xsl:call-template name ="tmpSMfontStyles">
                                <xsl:with-param name ="TextStyleID" select ="$textId" />
                              </xsl:call-template>
                            </xsl:if>
                          </a:rPr>
                      <a:t>
                            <xsl:call-template name ="insertTab" />
                      </a:t>
                    </a:r >
                  </xsl:when>
                    </xsl:choose>
                  </xsl:when>
                  <xsl:when test ="name()='presentation:date-time'">
                    <a:fld >
                      <xsl:attribute name ="id">
                        <xsl:value-of select ="'{11859447-1FF1-4571-A7A1-75F1CAF7F5D9}'"/>
                      </xsl:attribute>
                      <xsl:attribute name ="type">
                        <xsl:value-of select ="'datetime1'"/>
                      </xsl:attribute>
                      <a:rPr dirty="0" smtClean="0">
                        <xsl:variable name ="textId">
                          <xsl:value-of select ="./parent::node()/@text:style-name"/>
                        </xsl:variable>
                        <xsl:if test ="not($textId ='')">
                          <xsl:call-template name ="tmpSMfontStyles">
                            <xsl:with-param name ="TextStyleID" select ="$textId" />
                          </xsl:call-template>
                        </xsl:if>
                      </a:rPr>
                      <a:t> </a:t>
                    </a:fld>
                  </xsl:when>
                  <xsl:when test="not(node())">
                    <a:r>
                      <a:rPr dirty="0" smtClean="0">
                        <xsl:variable name ="textId">
                          <xsl:value-of select ="@text:style-name"/>
                        </xsl:variable>
                        <xsl:if test ="not($textId ='')">
                          <xsl:call-template name ="tmpSMfontStyles">
                            <xsl:with-param name ="TextStyleID" select ="$textId" />
                          </xsl:call-template>
                        </xsl:if>
                      </a:rPr>
                      <a:t>
                        <xsl:call-template name ="insertTab" />
                      </a:t>
                    </a:r >
                  </xsl:when>
                </xsl:choose>
              </xsl:for-each>
            </xsl:if>
            <xsl:if test ="not(child::node())">
              <xsl:value-of select="."/>
            </xsl:if>
          </xsl:if>
        </xsl:for-each>
        </a:p>
    </xsl:for-each>
  </xsl:template>
  
  <xsl:template name ="GetFrameDetails">
    <xsl:param name ="LayoutName"/>
    <xsl:param name ="SMName"/>
    <xsl:for-each select ="document('styles.xml')//style:master-page[@style:name=$SMName]/draw:frame[@presentation:class=$LayoutName]">
      <xsl:call-template name="tmpdrawCordinates"/>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="TextAlignment">
    <xsl:param name ="prId"/>
    <a:bodyPr>
      <xsl:for-each select ="draw:text-box">
        <xsl:for-each select ="text:p">
          <xsl:variable name ="ParId">
            <xsl:value-of select ="@text:style-name"/>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test="//style:style[@style:name=$ParId]/style:paragraph-properties/@style:writing-mode='tb-rl'">
              <xsl:for-each select ="//style:style[@style:name=$ParId]/style:paragraph-properties">
                <xsl:if test ="@style:writing-mode='tb-rl'">
                  <xsl:attribute name ="vert">
                    <xsl:value-of select ="'vert'"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:when>
            <xsl:when test="//style:style[@style:name=$prId]/style:paragraph-properties/@style:writing-mode='tb-rl'">
              <xsl:for-each select ="//style:style[@style:name=$prId]/style:paragraph-properties">
                <xsl:if test ="@style:writing-mode='tb-rl'">
                  <xsl:attribute name ="vert">
                    <xsl:value-of select ="'vert'"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:for-each>
            </xsl:when>
          </xsl:choose>

        </xsl:for-each>
      </xsl:for-each>
      <xsl:for-each select ="//style:style[@style:name=$prId]/style:graphic-properties">
        <xsl:call-template name ="tmpInternalPadding"/>
        <!-- Added by Vipul to Fix Bug 1794630 ODP: internal margins of textbox not retained -->
        <!--Start-->
        <xsl:variable name ="parentStyle">
          <xsl:value-of select ="parent::node()/@style:parent-style-name"/>
        </xsl:variable>
        <xsl:variable name ="default-padding-left">
          <xsl:call-template name ="getDefaultStyle">
            <xsl:with-param name ="parentStyle" select ="$parentStyle" />
            <xsl:with-param name ="attributeName" select ="'padding-left'" />
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name ="default-padding-top">
          <xsl:call-template name ="getDefaultStyle">
            <xsl:with-param name ="parentStyle" select ="$parentStyle" />
            <xsl:with-param name ="attributeName" select ="'padding-top'" />
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name ="default-padding-right">
          <xsl:call-template name ="getDefaultStyle">
            <xsl:with-param name ="parentStyle" select ="$parentStyle" />
            <xsl:with-param name ="attributeName" select ="'padding-right'" />
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name ="default-padding-bottom">
          <xsl:call-template name ="getDefaultStyle">
            <xsl:with-param name ="parentStyle" select ="$parentStyle" />
            <xsl:with-param name ="attributeName" select ="'padding-bottom'" />
          </xsl:call-template>
        </xsl:variable>
        <xsl:if test ="not(@fo:padding-left) and $parentStyle !=''">

          <xsl:attribute name ="lIns">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length">
                <xsl:call-template name="convertUnitsToCm">
                  <xsl:with-param name ="length" select ="$default-padding-left"/>
                </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>

          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="not(@fo:padding-top) and $parentStyle !=''">

          <xsl:attribute name ="tIns">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length">
                <xsl:call-template name="convertUnitsToCm">
                  <xsl:with-param name ="length" select ="$default-padding-top"/>
                </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="not(@fo:padding-right) and $parentStyle !=''">

          <xsl:attribute name ="rIns">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length">
                <xsl:call-template name="convertUnitsToCm">
                  <xsl:with-param name ="length" select ="$default-padding-right"/>
                </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="not(@fo:padding-bottom) and $parentStyle !=''">

          <xsl:attribute name ="bIns">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length">
                <xsl:call-template name="convertUnitsToCm">
                  <xsl:with-param name ="length" select ="$default-padding-bottom"/>
                </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
        <!--End-->
        <xsl:variable name ="anchorVal">
          <xsl:choose >
            <xsl:when test ="@draw:textarea-vertical-align ='top'">
              <xsl:value-of select ="'t'"/>
            </xsl:when>
            <xsl:when test ="@draw:textarea-vertical-align ='middle'">
              <xsl:value-of select ="'ctr'"/>
            </xsl:when>
            <xsl:when test ="@draw:textarea-vertical-align ='bottom'">
              <xsl:value-of select ="'b'"/>
            </xsl:when>
            <!--<xsl:otherwise >
							<xsl:value-of select ="'t'"/>
						</xsl:otherwise>-->
          </xsl:choose>
        </xsl:variable>
        <xsl:if test ="$anchorVal != ''">
          <xsl:attribute name ="anchor">
            <xsl:value-of select ="$anchorVal"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:variable name="anchorCtr">
          <xsl:choose >
            <xsl:when test ="@draw:textarea-horizontal-align ='center'">
              <xsl:value-of select ="'1'"/>
            </xsl:when>
            <xsl:when test ="@draw:textarea-horizontal-align='justify'">
              <xsl:value-of select ="'0'"/>
            </xsl:when>
          </xsl:choose>
        </xsl:variable>
        <xsl:if test ="$anchorCtr != ''">
          <xsl:attribute name ="anchorCtr">
            <xsl:value-of select ="$anchorCtr"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:call-template name="tmpWrap"/>

      </xsl:for-each>
    </a:bodyPr>
  </xsl:template>

  <xsl:template name="TextActionHyperLinks">
    <xsl:param name="PostionCount"/>
    <xsl:param name="nodeType"/>
    <xsl:param name="linkCount"/>
    <xsl:param name="ShapePos"/>
    <xsl:param name="pos"/>

    <xsl:for-each select="./text:a">
      <xsl:if test="position()=1">
        <xsl:if test="@xlink:href !=''">
        <a:hlinkClick>
          <xsl:if test="@xlink:href[ contains(.,'#Slide')]">
            <xsl:attribute name="action">
              <xsl:value-of select="'ppaction://hlinksldjump'"/>
            </xsl:attribute>
            <xsl:attribute name="r:id">
              <!--<xsl:value-of select="concat('TextHLAtchFileId',$ShapePos,'Link',$linkCount)"/>-->
              <xsl:value-of select="concat('TextHLAtchFileId',$ShapePos,'Link',$linkCount,$pos)"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
            <xsl:attribute name="r:id">
              <!--<xsl:if test="$pos !='' ">-->
              <xsl:value-of select="concat('TextHLAtchFileId',$ShapePos,'Link',$linkCount,$pos)"/>
              <!--</xsl:if>
              <xsl:if test="$pos ='' ">-->
              <!--<xsl:value-of select="concat('TextHLAtchFileId',$ShapePos,'Link',$linkCount)"/>
              </xsl:if>-->
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="not(@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and @xlink:href[ contains(.,':') or contains(.,'../')]">
            <xsl:attribute name="r:id">
              <xsl:value-of select="concat('TextHLAtchFileId',$ShapePos,'Link',$linkCount,$pos)"/>
              <!--<xsl:value-of select="concat('TextHLAtchFileId',$ShapePos,'Link',$linkCount)"/>-->
            </xsl:attribute>
          </xsl:if>
        </a:hlinkClick>
      </xsl:if>
      </xsl:if>
    </xsl:for-each>

  </xsl:template>
  <!--Get SlideLayout Type-->
  <xsl:template name ="slidesRel" match ="/office:document-content/office:body/office:presentation/draw:page">
    <xsl:param name ="slideNo"/>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <xsl:variable name ="slMasterName">
        <xsl:value-of select ="@draw:master-page-name"/>
      </xsl:variable>
      <xsl:variable name ="slMasterLink">
        <xsl:for-each select ="document('styles.xml')//style:master-page/@style:name">
          <xsl:if test ="$slMasterName = .">
            <xsl:value-of select ="position()"/>
          </xsl:if>
        </xsl:for-each>
      </xsl:variable>
      <xsl:variable name ="layoutTemplate">
        <xsl:value-of select ="substring-after(@presentation:presentation-page-layout-name,'T')"/>
      </xsl:variable >
      <xsl:variable name ="layoutNo">
        <xsl:choose >
           <xsl:when test ="$layoutTemplate ='7' or $layoutTemplate ='8' or
						   $layoutTemplate ='9' or $layoutTemplate ='10' or $layoutTemplate ='13'">
            <xsl:value-of select ="4"/>
          </xsl:when>
           <xsl:when test ="$layoutTemplate ='0'" >
            <xsl:value-of select ="1"/>
          </xsl:when>
          <xsl:when test ="$layoutTemplate ='1' or $layoutTemplate ='11'" >
            <xsl:value-of select ="2"/>
          </xsl:when>         
          <xsl:when test ="$layoutTemplate ='3'" >
            <xsl:value-of select ="4"/>
          </xsl:when>         
          <xsl:when test ="$layoutTemplate ='19'" >
            <xsl:value-of select ="6"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select ="7"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
	<!--code added by yeswanth to insert relationship ID for sound files (slide transition)-->
		<xsl:variable name="styleName">
			<xsl:value-of select="./@draw:style-name"/>
		</xsl:variable>
	
		<xsl:if test="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:family='drawing-page' and @style:name=$styleName]/style:drawing-page-properties/presentation:sound">
			<xsl:for-each select="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:family='drawing-page' and @style:name=$styleName]/style:drawing-page-properties/presentation:sound">
			<Relationship>
				<xsl:variable name="soundfileName">
          <xsl:value-of select="concat(generate-id(),'.wav')"/>
				</xsl:variable>
				<xsl:attribute name="Id">
						<xsl:value-of select="concat('strId',generate-id())"/>
				</xsl:attribute>
				<xsl:attribute name="Type">
					<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio'"/>
				</xsl:attribute>
				<xsl:attribute name="Target">
					<xsl:value-of select="concat('../media/',$soundfileName)"/>
				</xsl:attribute>
			</Relationship>
			</xsl:for-each>

		</xsl:if>		
		<!--end of code added by yeswanth-->
		
      <!-- added by vipul to insert relation for notes-->
      <!--start-->
      <xsl:if test="presentation:notes/draw:page-thumbnail">
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide">
          <xsl:attribute name ="Target">
            <xsl:value-of select ="concat('../notesSlides/notesSlide',$slideNo,'.xml')"/>
          </xsl:attribute>
        </Relationship>
      </xsl:if>
      <!--start-->
      <Relationship Id="rId1"
									Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout">
        <xsl:attribute name ="Target">
          <xsl:value-of select ="concat('../slideLayouts/slideLayout',((($slMasterLink - 1)*11)  + $layoutNo) ,'.xml')"/>
        </xsl:attribute>
      </Relationship >
      <!-- added by Vipul to set Relationship for background Image  Start-->
      <xsl:variable name="dpName">
        <xsl:value-of select="@draw:style-name" />
      </xsl:variable>
      <xsl:for-each select="//style:style[@style:name= $dpName]/style:drawing-page-properties">
        <xsl:if test="@draw:fill='bitmap'">
          <xsl:variable name="var_imageName" select="@draw:fill-image-name"/>
          <xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/draw:fill-image[@draw:name=$var_imageName]">
            <xsl:if test="position()=1">
            <xsl:if test ="contains(@xlink:href,'.png') or contains(@xlink:href,'.emf') or contains(@xlink:href,'.wmf') or contains(@xlink:href,'.jfif') or contains(@xlink:href,'.jpe') 
            or contains(@xlink:href,'.bmp') or contains(@xlink:href,'.dib') or contains(@xlink:href,'.rle')
            or contains(@xlink:href,'.bmz') or contains(@xlink:href,'.gfa') 
            or contains(@xlink:href,'.emz') or contains(@xlink:href,'.wmz') or contains(@xlink:href,'.pcz')
            or contains(@xlink:href,'.tif') or contains(@xlink:href,'.tiff') 
            or contains(@xlink:href,'.cdr') or contains(@xlink:href,'.cgm') or contains(@xlink:href,'.eps') 
            or contains(@xlink:href,'.pct') or contains(@xlink:href,'.pict') or contains(@xlink:href,'.wpg') 
            or contains(@xlink:href,'.jpeg') or contains(@xlink:href,'.gif') or contains(@xlink:href,'.png') or contains(@xlink:href,'.jpg')">
              <Relationship>
                <xsl:attribute name="Id">
                  <xsl:value-of select="concat('slide',$slideNo,'BackImg')" />
                </xsl:attribute>
                <xsl:attribute name="Type">
                  <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'" />
                </xsl:attribute>
                <xsl:attribute name="Target">
                  <xsl:value-of select="concat('../media/',substring-after(@xlink:href,'/'))" />
                  <!-- <xsl:value-of select ="concat('../media/',substring-after(@xlink:href,'/'))"/>  -->
                </xsl:attribute>
              </Relationship>
            </xsl:if>
            </xsl:if>         
          </xsl:for-each>
        </xsl:if>
      </xsl:for-each>
      <xsl:if test=".//draw:object or .//draw:object-ole">
        <xsl:choose>
          <xsl:when test="document(concat(substring-after(./child::node()[1]/@xlink:href,'./'),'/content.xml'))
                                      /office:document-content/office:body/office:drawing/draw:page/draw:frame/draw:image">

          </xsl:when>
          <xsl:when test="document(concat(substring-after(./child::node()[1]/@xlink:href,'./'),'/content.xml'))/child::node()"/>
          <xsl:when test="document(concat(translate(./child::node()[1]/@xlink:href,'/',''),'/content.xml'))/child::node()"/>
          <xsl:otherwise>
            <Relationship Id="vmldrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing">
              <xsl:attribute name="Target">
                <xsl:value-of select="concat('../drawings/vmlDrawing',$slideNo,'.vml')"/>
              </xsl:attribute>
            </Relationship>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
      <xsl:for-each select ="node()">
        <xsl:choose>
          <xsl:when test="name()='draw:frame'">
            <xsl:variable name="UniqueID">
              <xsl:value-of select="generate-id()"/>
            </xsl:variable>
            <xsl:variable name="var_pos">
              <xsl:call-template name="getShapePosTemp">
                <xsl:with-param name="var_pos" select="position()"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:for-each select=".">
              <xsl:choose>
                <xsl:when test="./draw:object or ./draw:object-ole">
                  <xsl:choose>
                    <xsl:when test="document(concat(substring-after(./child::node()[1]/@xlink:href,'./'),'/content.xml'))/child::node() or
                                      document(concat(translate(./child::node()[1]/@xlink:href,'/',''),'/content.xml'))/child::node() ">
                      <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                                 xmlns="http://schemas.openxmlformats.org/package/2006/relationships">

                        <xsl:attribute name="Id">
                          <xsl:value-of select="concat('oleObjectImage_',generate-id())"/>
                        </xsl:attribute>
                        <xsl:attribute name="Target">
                          <xsl:value-of select="concat('../media/','oleObjectImage_',generate-id(),'.png')"/>
                        </xsl:attribute>

                      </Relationship>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:call-template name="tmpOLEObjectsRel">
                        <xsl:with-param name="slideNo" select="$slideNo"/>
                      </xsl:call-template>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:when test="./table:table and ./draw:image">
                  <xsl:for-each select="./table:table/table:table-row/table:table-cell">
                  <xsl:call-template name="tmpBitmapFillRel">
                    <xsl:with-param name ="UniqueId" select="generate-id(.)" />
                    <xsl:with-param name ="FileName" select="'content.xml'" />
                    <xsl:with-param name ="prefix" select="'bitmap'" />
                  </xsl:call-template>
                    <!--To Do: Table text hyperlink-->
                    <!--<xsl:call-template name="tmpHyperLnkBuImgRel">
                      <xsl:with-param name ="var_pos" select="$var_pos" />
                      <xsl:with-param name ="shapeId" select="$var_pos" />
                      <xsl:with-param name ="UniqueId" select="generate-id(.)" />

                    </xsl:call-template>-->
                  </xsl:for-each>
                </xsl:when>
                <xsl:when test="./draw:image">
                  <xsl:for-each select="./draw:image">
                    <xsl:if test ="contains(@xlink:href,'.png') or contains(@xlink:href,'.emf') or contains(@xlink:href,'.wmf') or contains(@xlink:href,'.jfif') or contains(@xlink:href,'.jpe') 
            or contains(@xlink:href,'.bmp') or contains(@xlink:href,'.dib') or contains(@xlink:href,'.rle')
            or contains(@xlink:href,'.bmz') or contains(@xlink:href,'.gfa') 
            or contains(@xlink:href,'.emz') or contains(@xlink:href,'.wmz') or contains(@xlink:href,'.pcz')
            or contains(@xlink:href,'.tif') or contains(@xlink:href,'.tiff') 
            or contains(@xlink:href,'.cdr') or contains(@xlink:href,'.cgm') or contains(@xlink:href,'.eps') 
            or contains(@xlink:href,'.pct') or contains(@xlink:href,'.pict') or contains(@xlink:href,'.wpg') 
            or contains(@xlink:href,'.jpeg') or contains(@xlink:href,'.gif') or contains(@xlink:href,'.png') or contains(@xlink:href,'.jpg')">
                      <xsl:if test="not(./@xlink:href[contains(.,'../')])">
                        <Relationship>
                          <xsl:attribute name="Id">
                            <xsl:value-of select="concat('sl',$slideNo,'Image',$var_pos)" />
                          </xsl:attribute>
                          <xsl:attribute name="Type">
                            <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'" />
                          </xsl:attribute>
                          <xsl:attribute name="Target">
                            <xsl:value-of select="concat('../media/',substring-after(@xlink:href,'/'))" />
                          </xsl:attribute>
                        </Relationship>
                        <!--added by chhavi for picture hyperlink relationship-->
                        <xsl:if test="./following-sibling::node()[name() = 'office:event-listeners']">
                          <xsl:for-each select ="./parent::node()">
                            <xsl:call-template name="tmpOfficeListnerRelationship">
                              <xsl:with-param name="ShapeType" select="'picture'"/>
                              <xsl:with-param name="PostionCount" select="generate-id()"/>
                              <xsl:with-param name="Type" select="'FRAME'"/>
                                    </xsl:call-template>
                                      </xsl:for-each>
                        </xsl:if>
                       </xsl:if>
                    </xsl:if>
                  </xsl:for-each>
                </xsl:when>
                <xsl:when test="./draw:plugin">
                  <xsl:for-each select="./draw:plugin">
                    <Relationship>
                      <xsl:attribute name="Id">
                        <xsl:value-of  select ="concat('sl',$slideNo,'Au',$var_pos)"/>
                      </xsl:attribute>
                      <xsl:variable name="wavId">
                        <xsl:value-of  select ="concat('sl',$slideNo,'Au',$var_pos)"/>
                      </xsl:variable>
                      <xsl:attribute name="Type">
                        <xsl:if test="@xlink:href[ contains(.,'mp3')] or @xlink:href[ contains(.,'m3u')] or @xlink:href[ contains(.,'wma')] or 
          @xlink:href[ contains(.,'wax')] or @xlink:href[ contains(.,'aif')] or @xlink:href[ contains(.,'aifc')] or
          @xlink:href[ contains(.,'aiff')] or @xlink:href[ contains(.,'au')] or @xlink:href[ contains(.,'snd')] or
          @xlink:href[ contains(.,'mid')] or @xlink:href[ contains(.,'midi')] or @xlink:href[ contains(.,'rmi')] or @xlink:href[ contains(.,'wav')] or @xlink:href[ contains(.,'WAV')]">
                          <xsl:value-of select ="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio'"/>
                        </xsl:if>
                        <xsl:if test="not(@xlink:href[ contains(.,'mp3')] or @xlink:href[ contains(.,'m3u')] or @xlink:href[ contains(.,'wma')] or 
          @xlink:href[ contains(.,'wax')] or @xlink:href[ contains(.,'aif')] or @xlink:href[ contains(.,'aifc')] or
          @xlink:href[ contains(.,'aiff')] or @xlink:href[ contains(.,'au')] or @xlink:href[ contains(.,'snd')] or
          @xlink:href[ contains(.,'mid')] or @xlink:href[ contains(.,'midi')] or @xlink:href[ contains(.,'rmi')] or @xlink:href[ contains(.,'wav')] or @xlink:href[ contains(.,'WAV')])">
                          <xsl:value-of select ="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video'"/>
                        </xsl:if>
                      </xsl:attribute>
                      <xsl:attribute name="Target">
                        <xsl:choose>
                          <xsl:when test="starts-with(@xlink:href,'/') or starts-with(@xlink:href,'file:///')">
                              <xsl:value-of select="concat('file:///',translate(substring-after(@xlink:href,'/'),'/','\'))"/>
                          </xsl:when>
                          <xsl:when test="@xlink:href[contains(.,'http://')] or @xlink:href[contains(.,'https://')]">
                            <xsl:value-of select="@xlink:href"/>
                          </xsl:when>
                          <xsl:when test="@xlink:href[ contains(.,'wav')] or @xlink:href[ contains(.,'WAV')]">
                            <xsl:value-of select="concat('../media/',$wavId,'.wav')"/>
                          </xsl:when>
                          <xsl:when test="@xlink:href[ contains(.,'./')]">
                            <xsl:if test="string-length(substring-after(@xlink:href, '../')) = 0">
                              <xsl:value-of select="/"/>
                            </xsl:if>
                            <xsl:if test="not(string-length(substring-after(@xlink:href, '../')) = 0)">
                              <!--links Absolute Path-->
                              <xsl:variable name ="xlinkPath" >
                                <xsl:value-of select ="@xlink:href"/>
                              </xsl:variable>
                              <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
                            </xsl:if>
                          </xsl:when>
                          <xsl:when test="not(@xlink:href[ contains(.,'./')])">
                            <xsl:value-of select="concat('file:///',translate(substring-after(@xlink:href,'/'),'/','\'))"/>
                          </xsl:when>
                        </xsl:choose>
                      </xsl:attribute>
                      <xsl:choose>
                        <xsl:when test="starts-with(@xlink:href,'/') or starts-with(@xlink:href,'file:///')">
                          <xsl:attribute name="TargetMode">
                            <xsl:value-of select="'External'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:when test="@xlink:href[contains(.,'http://')]  or @xlink:href[contains(.,'https://')]">
                          <xsl:attribute name="TargetMode">
                            <xsl:value-of select="'External'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:when test="not(@xlink:href[ contains(.,'wav')] or @xlink:href[ contains(.,'WAV')])">
                          <xsl:attribute name="TargetMode">
                            <xsl:value-of select="'External'"/>
                          </xsl:attribute>
                        </xsl:when>
                      </xsl:choose>
                   
                    </Relationship>
                    <!-- for image link -->
                    <Relationship>
                      <xsl:attribute name ="Id">
                        <xsl:value-of  select ="concat('sl',$slideNo,'Im',$var_pos)"/>
                      </xsl:attribute>
                      <xsl:attribute name ="Type" >
                        <xsl:value-of select ="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'"/>
                      </xsl:attribute>
                      <xsl:attribute name ="Target">
                        <xsl:if test="./draw:image/@xlink:href != ''">
                          <xsl:value-of select ="concat('../media/',substring-after(./draw:image/@xlink:href,'/'))"/>
                        </xsl:if>
                        <xsl:if test="not(./draw:image/@xlink:href)">
                          <xsl:value-of select ="concat('../media/','thumbnail.png')"/>
                        </xsl:if>
                      </xsl:attribute>
                    </Relationship >
                  </xsl:for-each>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:variable name="PostionCount">
                    <xsl:value-of select="$var_pos"/>
                  </xsl:variable>
                  <xsl:choose>
                    <xsl:when test="not(@presentation:style-name) and not(@presentation:class)">
                      <xsl:for-each select="draw:text-box">
                           <xsl:variable name="shapeId">
                            <xsl:value-of select="concat('text-box',$PostionCount)"/>
                          </xsl:variable>
                        <xsl:call-template name="tmpHyperLnkBuImgRel">
                          <xsl:with-param name ="var_pos" select="$var_pos" />
                          <xsl:with-param name ="shapeId" select="$shapeId" />
                          <xsl:with-param name ="UniqueID" select="$UniqueID" />
                          
                                    </xsl:call-template>
                         </xsl:for-each>
                    </xsl:when>
                    <xsl:when test="@presentation:class or @presentation:style-name">
                      <xsl:variable name="FrameCount">
                        <xsl:value-of select="concat('Frame',$var_pos)"/>
                      </xsl:variable>
                      <xsl:for-each select="draw:text-box">
                        <xsl:for-each select="text:p">
                          <xsl:variable name="var_ParPos" select="position()"/>
                          <xsl:if test="text:a/@xlink:href !=''">
                            <Relationship>
                              <xsl:attribute name="Id">
                                <xsl:value-of select="concat('TextHLAtchFileId',$var_pos,'Link',position())"/>
                              </xsl:attribute>
                              <xsl:choose>
                                <xsl:when test="text:a/@xlink:href[contains(.,'#Slide')]">
                                  <xsl:attribute name="Type">
                                    <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
                                  </xsl:attribute>
                                  <xsl:attribute name="Target">
                                    <xsl:value-of select="concat('slide',substring-after(text:a/@xlink:href,'Slide '),'.xml')"/>
                                  </xsl:attribute>
                                </xsl:when>
                                <xsl:when test="text:a/@xlink:href[contains(.,'http://') or contains(.,'mailto:')]">
                                  <xsl:attribute name="Type">
                                    <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                                  </xsl:attribute>
                                  <xsl:attribute name="Target">
                                    <xsl:value-of select="text:a/@xlink:href"/>
                                  </xsl:attribute>
                                  <xsl:attribute name="TargetMode">
                                    <xsl:value-of select="'External'"/>
                                  </xsl:attribute>
                                </xsl:when>
                                <xsl:otherwise>
                                  <xsl:if test="text:a/@xlink:href[ contains (.,':') ]">
                                    <xsl:attribute name="Type">
                                      <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="Target">
                                      <xsl:value-of select="concat('file:///',translate(substring-after(text:a/@xlink:href,'/'),'/','\'))"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="TargetMode">
                                      <xsl:value-of select="'External'"/>
                                    </xsl:attribute>
                                  </xsl:if>
                                  <xsl:if test="not(text:a/@xlink:href[ contains (.,':') ])">
                                    <xsl:attribute name="Type">
                                      <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="Target">
                                      <!--links Absolute Path-->
                                      <xsl:variable name ="xlinkPath" >
                                        <xsl:value-of select ="text:a/@xlink:href"/>
                                      </xsl:variable>
                                      <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="TargetMode">
                                      <xsl:value-of select="'External'"/>
                                    </xsl:attribute>
                                  </xsl:if>
                                </xsl:otherwise>
                              </xsl:choose>
                            </Relationship>
                          </xsl:if>
                          <xsl:for-each select="node()">
                            <xsl:if test="name()='text:span'">
                              <xsl:variable name="tmpPos" select="position()"/>
                              <xsl:if test="text:a/@xlink:href !=''">
                                <Relationship>
                                  <xsl:attribute name="Id">
                                    <xsl:value-of select="concat('TextHLAtchFileId',$var_pos,'Link',$var_ParPos,$tmpPos)"/>
                                  </xsl:attribute>
                                  <xsl:choose>
                                    <xsl:when test="text:a/@xlink:href[contains(.,'#Slide')]">
                                      <xsl:attribute name="Type">
                                        <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
                                      </xsl:attribute>
                                      <xsl:attribute name="Target">
                                        <xsl:value-of select="concat('slide',substring-after(text:a/@xlink:href,'Slide '),'.xml')"/>
                                      </xsl:attribute>
                                    </xsl:when>
                                    <xsl:when test="text:a/@xlink:href[contains(.,'http://') or contains(.,'mailto:')]">
                                      <xsl:attribute name="Type">
                                        <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                                      </xsl:attribute>
                                      <xsl:attribute name="Target">
                                        <xsl:value-of select="text:a/@xlink:href"/>
                                      </xsl:attribute>
                                      <xsl:attribute name="TargetMode">
                                        <xsl:value-of select="'External'"/>
                                      </xsl:attribute>
                                    </xsl:when>
                                    <xsl:otherwise>
                                      <xsl:if test="text:a/@xlink:href[ contains (.,':') ]">
                                        <xsl:attribute name="Type">
                                          <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                                        </xsl:attribute>
                                        <xsl:attribute name="Target">
                                          <xsl:value-of select="concat('file:///',translate(substring-after(text:a/@xlink:href,'/'),'/','\'))"/>
                                        </xsl:attribute>
                                        <xsl:attribute name="TargetMode">
                                          <xsl:value-of select="'External'"/>
                                        </xsl:attribute>
                                      </xsl:if>
                                      <xsl:if test="not(text:a/@xlink:href[ contains (.,':') ])">
                                        <xsl:attribute name="Type">
                                          <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                                        </xsl:attribute>
                                        <xsl:attribute name="Target">
                                          <!--links Absolute Path-->
                                          <xsl:variable name ="xlinkPath" >
                                            <xsl:value-of select ="text:a/@xlink:href"/>
                                          </xsl:variable>
                                          <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
                                        </xsl:attribute>
                                        <xsl:attribute name="TargetMode">
                                          <xsl:value-of select="'External'"/>
                                        </xsl:attribute>
                                      </xsl:if>
                                    </xsl:otherwise>
                                  </xsl:choose>
                                </Relationship>
                              </xsl:if>
                            </xsl:if>
                          </xsl:for-each>
                        </xsl:for-each>
                        <xsl:for-each select="text:list">
                          <xsl:variable name="forCount" select="position()" />
                          <xsl:variable name="var_ParPos" select="position()"/>
                          <xsl:variable name="bulletlvl">
                            <xsl:call-template name ="getListLevel">
                              <xsl:with-param name ="levelCount"/>
                            </xsl:call-template>
                          </xsl:variable>
                          <xsl:variable name="blvl">
                            <xsl:if test="$bulletlvl > 0">
                              <xsl:value-of select="$bulletlvl"/>
                            </xsl:if>
                            <xsl:if test="not($bulletlvl > 0)">
                              <xsl:value-of select="'0'"/>
                            </xsl:if>
                          </xsl:variable>
                          <xsl:variable name="xhrefValue">
                            <xsl:call-template name ="getTextHyperlinksForBulltes">
                              <xsl:with-param name ="blvl" select="$blvl"/>
                            </xsl:call-template>
                          </xsl:variable>
                          <xsl:variable name ="BuImgRel" select ="concat($var_pos,$var_ParPos)"/>
                          <xsl:variable name ="listId" select ="./@text:style-name"/>
                          <xsl:variable name="paragraphId" >
                            <xsl:call-template name ="getParaStyleName">
                              <xsl:with-param name ="lvl" select ="$blvl"/>
                            </xsl:call-template>
                          </xsl:variable>
                          <xsl:variable name ="isNumberingEnabled">
                            <xsl:choose >
                              <xsl:when test ="document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering">
                                <xsl:value-of select ="document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering"/>
                              </xsl:when>
                              <xsl:otherwise>
                                <xsl:value-of select ="'true'"/>
                              </xsl:otherwise>
                            </xsl:choose>
                          </xsl:variable>
                          <xsl:if test="string-length($xhrefValue) > 0">
                            <Relationship>
                              <xsl:attribute name="Id">
                                <xsl:value-of select="concat($FrameCount,'BTHLFileId',$blvl,'Link',$forCount)"/>
                              </xsl:attribute>
                              <xsl:choose>
                                <xsl:when test="contains($xhrefValue,'#Slide')">
                                  <xsl:attribute name="Type">
                                    <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
                                  </xsl:attribute>
                                  <xsl:attribute name="Target">
                                    <xsl:value-of select="concat('slide',substring-after($xhrefValue,'Slide '),'.xml')"/>
                                  </xsl:attribute>
                                </xsl:when>
                                <xsl:when test="contains($xhrefValue,'http://') or contains($xhrefValue,'mailto:')">
                                  <xsl:attribute name="Type">
                                    <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                                  </xsl:attribute>
                                  <xsl:attribute name="Target">
                                    <xsl:value-of select="$xhrefValue"/>
                                  </xsl:attribute>
                                  <xsl:attribute name="TargetMode">
                                    <xsl:value-of select="'External'"/>
                                  </xsl:attribute>
                                </xsl:when>
                                <xsl:otherwise>
                                  <xsl:if test="contains($xhrefValue,':')">
                                    <xsl:attribute name="Type">
                                      <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="Target">
                                      <xsl:value-of select="concat('file:///',translate(substring-after($xhrefValue,'/'),'/','\'))"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="TargetMode">
                                      <xsl:value-of select="'External'"/>
                                    </xsl:attribute>
                                  </xsl:if>
                                  <xsl:if test="not(contains ($xhrefValue,':'))">
                                    <xsl:attribute name="Type">
                                      <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="Target">
                                      <!--links Absolute Path-->
                                      <xsl:variable name ="xlinkPath" >
                                        <xsl:value-of select ="$xhrefValue"/>
                                      </xsl:variable>
                                      <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="TargetMode">
                                      <xsl:value-of select="'External'"/>
                                    </xsl:attribute>
                                  </xsl:if>
                                </xsl:otherwise>
                              </xsl:choose>
                            </Relationship>
                          </xsl:if>
                          <xsl:for-each select ="document('content.xml')//text:list-style[@style:name=$listId]">
                            <xsl:if test ="text:list-level-style-image[@text:level=$blvl+1] and $isNumberingEnabled='true' and text:list-level-style-image[@text:level=$blvl+1]/@xlink:href">
                              <!--<xsl:variable name ="rId" select ="concat('buImage',$listId,$blvl+1,$forCount,$var_pos)"/>-->
                              <xsl:variable name ="rId" select ="concat('buImage',$listId,$BuImgRel,generate-id())"/>
                              <xsl:variable name ="imageName">
                                <xsl:choose>
                                  <xsl:when test="substring-after(text:list-level-style-image[@text:level=$blvl+1]/@xlink:href,'Pictures/') != ''">
                                    <xsl:value-of select="substring-after(text:list-level-style-image[@text:level=$blvl+1]/@xlink:href,'Pictures/')"/>
                                  </xsl:when>
                                  <xsl:when test="substring-after(text:list-level-style-image[@text:level=$blvl+1]/@xlink:href,'media/') != ''">
                                    <xsl:value-of select="substring-after(text:list-level-style-image[@text:level=$blvl+1]/@xlink:href,'media/')"/>
                                  </xsl:when>
                                </xsl:choose>
                              </xsl:variable>
                              <Relationship >
                                <xsl:attribute name ="Id">
                                  <xsl:value-of  select ="$rId"/>
                                </xsl:attribute>
                                <xsl:attribute name ="Type" >
                                  <xsl:value-of select ="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'"/>
                                </xsl:attribute>
                                <xsl:attribute name ="Target">
                                  <xsl:value-of select ="concat('../media/',$imageName)"/>
                                </xsl:attribute>
                              </Relationship >
                            </xsl:if>
                          </xsl:for-each>
                        </xsl:for-each>
                      </xsl:for-each>
                    </xsl:when>
                  </xsl:choose>
                  <xsl:if test="office:event-listeners">
                    <xsl:for-each select=".">
                      <xsl:variable name="PresentationClass">
                        <xsl:if test="@presentation:class">
                          <xsl:value-of select="1"/>
                        </xsl:if>
                        <xsl:if test="not(@presentation:class)">
                          <xsl:value-of select="0"/>
                        </xsl:if>
                      </xsl:variable>
                      <xsl:if test="$PresentationClass = 0">
                        <xsl:if test="count(office:event-listeners/presentation:event-listener) = 1">
                          <xsl:call-template name="tmpOfficeListnerRelationship">
                            <xsl:with-param name="ShapeType" select="'TxtBoxAtchFileId'"/>
                            <xsl:with-param name="PostionCount" select="generate-id()"/>
                            <xsl:with-param name="Type" select="'FRAME'"/>
                                    </xsl:call-template>
                                   </xsl:if>
                      </xsl:if>
                    </xsl:for-each>
                    <xsl:for-each select="./@presentation:class">
                      <xsl:if test="parent::node()/office:event-listeners/presentation:event-listener">
                        <xsl:for-each select ="parent::node()">
                        <xsl:call-template name="tmpOfficeListnerRelationship">
                          <xsl:with-param name="ShapeType" select="'AtchFileId'"/>
                          <xsl:with-param name="PostionCount" select="generate-id()"/>
                          <xsl:with-param name="Type" select="'FRAME'"/>
                                </xsl:call-template>
                              </xsl:for-each>
                                </xsl:if>
                                     </xsl:for-each>
                  </xsl:if>
                  <xsl:call-template name="tmpBitmapFillRel">
                    <xsl:with-param name ="UniqueId" select="$var_pos" />
                    <xsl:with-param name ="FileName" select="'content.xml'" />
                    <xsl:with-param name ="prefix" select="'bitmap'" />
                  </xsl:call-template>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:for-each>
          </xsl:when>
          <xsl:when test="name()='draw:custom-shape'">
            <xsl:variable name="UniqueID">
              <xsl:value-of select="generate-id()"/>
            </xsl:variable>
            <xsl:variable name="var_pos">
              <xsl:call-template name="getShapePosTemp">
                <xsl:with-param name="var_pos" select="position()"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:for-each select=".">

              <xsl:call-template name="tmpBitmapFillRel">
                <xsl:with-param name ="UniqueId" select="$var_pos" />
                <xsl:with-param name ="FileName" select="'content.xml'" />
                <xsl:with-param name ="prefix" select="'bitmap'" />
              </xsl:call-template>
              <xsl:if test="./office:event-listeners">
                <xsl:variable name="ShapePostionCount">
                  <xsl:value-of select="$var_pos"/>
                </xsl:variable>
                <xsl:call-template name="tmpOfficeListnerRelationship">
                  <xsl:with-param name="ShapeType" select="'ShapeFileId'"/>
                  <xsl:with-param name="PostionCount" select="generate-id()"/>
                  <xsl:with-param name="Type" select="'CUSTOM'"/>
                            </xsl:call-template>
                         </xsl:if>
              <xsl:variable name="ShapePostionCount">
                <xsl:value-of select="position()"/>
              </xsl:variable>
              <xsl:variable name="shapeId">
                <xsl:value-of select="concat('custom-shape',$var_pos)"/>
              </xsl:variable>
              <xsl:call-template name="tmpHyperLnkBuImgRel">
                <xsl:with-param name ="var_pos" select="$var_pos" />
                <xsl:with-param name ="shapeId" select="$shapeId" />
                <xsl:with-param name ="UniqueID" select="$UniqueID" />
                  </xsl:call-template>
                 </xsl:for-each>
          </xsl:when>
          <xsl:when test="name()='draw:rect'">
            <xsl:variable name="UniqueID">
              <xsl:value-of select="generate-id()"/>
            </xsl:variable>
            <xsl:variable name="var_pos">
              <xsl:call-template name="getShapePosTemp">
                <xsl:with-param name="var_pos" select="position()"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:for-each select=".">
              <xsl:choose>
                <xsl:when test="./office:event-listeners">
                  <xsl:variable name="ShapePostionCount">
                    <xsl:value-of select="$var_pos"/>
                  </xsl:variable>
                  <xsl:call-template name="tmpOfficeListnerRelationship">
                    <xsl:with-param name="ShapeType" select="'RectAtachFileId'"/>
                    <xsl:with-param name="PostionCount" select="generate-id()"/>
                    <xsl:with-param name="Type" select="'RECT'"/>
                            </xsl:call-template>
                         </xsl:when>
              </xsl:choose>
              <xsl:variable name="shapeId">
                <xsl:value-of select="concat('rect',$var_pos)"/>
              </xsl:variable>
              <xsl:call-template name="tmpHyperLnkBuImgRel">
                <xsl:with-param name ="var_pos" select="$var_pos" />
                <xsl:with-param name ="shapeId" select="$shapeId" />
                <xsl:with-param name ="UniqueID" select="$UniqueID" />
                  </xsl:call-template>
                <xsl:call-template name="tmpBitmapFillRel">
                <xsl:with-param name ="UniqueId" select="$var_pos" />
                <xsl:with-param name ="FileName" select="'content.xml'" />
                <xsl:with-param name ="prefix" select="'bitmap'" />
              </xsl:call-template>
            </xsl:for-each>
          </xsl:when>
          <xsl:when test="name()='draw:ellipse' or name()='draw:circle'">
            <xsl:variable name="UniqueID">
              <xsl:value-of select="generate-id()"/>
            </xsl:variable>
            <xsl:variable name="var_pos">
              <xsl:call-template name="getShapePosTemp">
                <xsl:with-param name="var_pos" select="position()"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:for-each select=".">
              <xsl:variable name="shapeId">
                <xsl:value-of select="concat('ellipse',$var_pos)"/>
              </xsl:variable>
              <xsl:call-template name="tmpHyperLnkBuImgRel">
                <xsl:with-param name ="var_pos" select="$var_pos" />
                <xsl:with-param name ="shapeId" select="$shapeId" />
                <xsl:with-param name ="UniqueID" select="$UniqueID" />
                  </xsl:call-template>
                 <xsl:call-template name="tmpBitmapFillRel">
                <xsl:with-param name ="UniqueId" select="$var_pos" />
                <xsl:with-param name ="FileName" select="'content.xml'" />
                <xsl:with-param name ="prefix" select="'bitmap'" />
              </xsl:call-template>
            </xsl:for-each>
          </xsl:when>
          <xsl:when test="name()='draw:line'">
            <xsl:variable name="UniqueID">
              <xsl:value-of select="generate-id()"/>
            </xsl:variable>
            <xsl:variable name="var_pos">
              <xsl:call-template name="getShapePosTemp">
                <xsl:with-param name="var_pos" select="position()"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:for-each select=".">
              <xsl:if test="./office:event-listeners">
                <xsl:variable name="ShapePostionCount">
                  <xsl:value-of select="$var_pos"/>
                </xsl:variable>
                <xsl:call-template name="tmpOfficeListnerRelationship">
                  <xsl:with-param name="ShapeType" select="'LineFileId'"/>
                  <xsl:with-param name="PostionCount" select="generate-id()"/>
                  <xsl:with-param name="Type" select="'LINE'"/>
                      </xsl:call-template>
                     </xsl:if>
            </xsl:for-each>
          </xsl:when>
          <xsl:when test="name()='draw:connector'">
            <xsl:variable name="UniqueID">
              <xsl:value-of select="generate-id()"/>
            </xsl:variable>
            <xsl:variable name="var_pos">
              <xsl:call-template name="getShapePosTemp">
                <xsl:with-param name="var_pos" select="position()"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:for-each select=".">
              <xsl:if test="./office:event-listener">
                <xsl:variable name="ShapePostionCount">
                  <xsl:value-of select="$var_pos"/>
                </xsl:variable>
                <xsl:call-template name="tmpOfficeListnerRelationship">
                  <xsl:with-param name="ShapeType" select="'LineFileId'"/>
                  <xsl:with-param name="PostionCount" select="generate-id()"/>
                  <xsl:with-param name="Type" select="'CONNECTOR'"/>
                      </xsl:call-template>
                          </xsl:if>
            </xsl:for-each>
          </xsl:when>
          <xsl:when test="name()='draw:g'">
            <xsl:variable name="var_pos">
              <xsl:call-template name="getShapePosTemp">
                <xsl:with-param name="var_pos" select="position()"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:call-template name="tmpGroupingRelation">
              <xsl:with-param name="slideNo" select="$slideNo"/>
              <xsl:with-param name="pos" select="$var_pos"/>
              <xsl:with-param name="startPos" select="'1'"/>
              <xsl:with-param name="FileName" select="'content.xml'"/>
            </xsl:call-template>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </Relationships>
  </xsl:template>
  <xsl:template name="GetUniqueRelationIdForWavFile">
    <xsl:param name ="FilePath"/>
    <xsl:param name ="ShapePosition"/>
    <xsl:param name="Page" />
    <xsl:param name="ShapeType" />
    <xsl:param name="isCalledFromGroup" />
    <xsl:for-each select="$Page">
      <xsl:if test="$isCalledFromGroup != 'GROUP'">
        <xsl:if test="$ShapeType = 'CUSTOM' ">
          <xsl:for-each select="./draw:custom-shape">
            <xsl:if test="position() &lt; $ShapePosition">
              <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
                <xsl:value-of select='1'/>
              </xsl:if>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
        <!--<xsl:if test="$ShapeType = 'CUSTOM' and $ShapePosition = '1'" >
          <xsl:value-of select='0'/>
        </xsl:if>-->
        <xsl:if test="$ShapeType = 'RECT'">
          <xsl:for-each select="./draw:custom-shape">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
        <xsl:if test="$ShapeType = 'LINE'">
          <xsl:for-each select="./draw:custom-shape">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)]">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:for-each>
          <xsl:for-each select="./draw:rect">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
        <xsl:if test="$ShapeType = 'FRAME'">
          <xsl:for-each select="./draw:custom-shape">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:for-each>
          <xsl:for-each select="./draw:rect">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)]">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:for-each>
          <xsl:for-each select="./draw:line">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
        <xsl:if test="$ShapeType = 'CONNECTOR'">
          <xsl:for-each select="./draw:custom-shape">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:for-each>
          <xsl:for-each select="./draw:rect">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:for-each>
          <xsl:for-each select="./draw:line">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:for-each>
          <xsl:for-each select="./draw:frame">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
      </xsl:if>
      <xsl:if test="$isCalledFromGroup = 'GROUP'">
        <xsl:if test="$ShapeType = 'CUSTOM' ">
          <xsl:for-each select="./draw:g/draw:custom-shape">
            <xsl:if test="position() &lt; $ShapePosition">
              <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
                <xsl:value-of select='1'/>
              </xsl:if>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
        <xsl:if test="$ShapeType = 'CUSTOM' and $ShapePosition = '1'" >
          <xsl:value-of select='0'/>
        </xsl:if>
        <xsl:if test="$ShapeType = 'RECT'">
          <xsl:for-each select="./draw:g/draw:custom-shape">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
        <xsl:if test="$ShapeType = 'LINE'">
          <xsl:for-each select="./draw:g/draw:custom-shape">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)]">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:for-each>
          <xsl:for-each select="./draw:g/draw:rect">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
        <xsl:if test="$ShapeType = 'FRAME'">
          <xsl:for-each select="./draw:g/draw:custom-shape">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:for-each>
          <xsl:for-each select="./draw:g/draw:rect">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)]">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:for-each>
          <xsl:for-each select="./draw:g/draw:line">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
        <xsl:if test="$ShapeType = 'CONNECTOR'">
          <xsl:for-each select="./draw:g/draw:custom-shape">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:for-each>
          <xsl:for-each select="./draw:g/draw:rect">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:for-each>
          <xsl:for-each select="./draw:g/draw:line">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:for-each>
          <xsl:for-each select="./draw:g/draw:frame">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  <!-- Templates added by vijayeta,to get paragraph style name for each of the levels in multilevelled list-->
  <xsl:template name ="getParaStyleName">
    <xsl:param name ="lvl"/>
    <xsl:choose>
      <xsl:when test ="$lvl='0'">
        <xsl:value-of select ="text:list-item/text:p/@text:style-name"/>
      </xsl:when >
      <xsl:when test ="$lvl='1'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='2'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='3'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='4'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='5'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='6'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='7'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='8'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='9'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="getHyperlinksForBulltedText">
    <xsl:param name ="bulletLevel"/>
    <xsl:param name ="listItemCount" />
    <xsl:param name ="FrameCount" />
    <xsl:param name ="rootNode" />
    <xsl:choose>
      <xsl:when test ="./text:p">
        <xsl:for-each select="./text:p">
          <xsl:if test="text:a">
            <a:hlinkClick>
              <xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
          <xsl:if test="text:span/text:a">
            <a:hlinkClick>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:p">
        <xsl:for-each select="./text:list/text:list-item/text:p">
          <xsl:if test="text:a">
            <a:hlinkClick>
              <xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
          <xsl:if test="text:span/text:a">
            <a:hlinkClick>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:if test="text:a">
            <a:hlinkClick>
              <xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
          <xsl:if test="text:span/text:a">
            <a:hlinkClick>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:if test="text:a">
            <a:hlinkClick>
              <xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
          <xsl:if test="text:span/text:a">
            <a:hlinkClick>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:if test="text:a">
            <a:hlinkClick>
              <xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
          <xsl:if test="text:span/text:a">
            <a:hlinkClick>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:if test="text:a">
            <a:hlinkClick>
              <xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
          <xsl:if test="text:span/text:a">
            <a:hlinkClick>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:if test="text:a">
            <a:hlinkClick>
              <xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
          <xsl:if test="text:span/text:a">
            <a:hlinkClick>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:if test="text:a">
            <a:hlinkClick>
              <xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
          <xsl:if test="text:span/text:a">
            <a:hlinkClick>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:if test="text:a">
            <a:hlinkClick>
              <xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
          <xsl:if test="text:span/text:a">
            <a:hlinkClick>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

	<!--added by yeswanth.s-->
	<!--slideTransition template-->
	<xsl:template name="slideTransition">
		<p:transition>
			<xsl:if test="./@draw:style-name">
				<xsl:variable name="styleName">
					<xsl:value-of select="./@draw:style-name"/>
				</xsl:variable>
				<xsl:for-each select="//style:style[@style:family='drawing-page' and @style:name=$styleName]/style:drawing-page-properties">
					
						<xsl:attribute name="spd">
							<xsl:choose>
								<xsl:when test="./@presentation:transition-speed='slow'">
									<xsl:value-of select="'slow'"/>
								</xsl:when>
								<xsl:when test="./@presentation:transition-speed='fast'">
									<xsl:value-of select="'fast'"/>
								</xsl:when>
								<xsl:otherwise>
									<xsl:value-of select="'med'"/>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					
					<xsl:if test="./@presentation:duration">
						<xsl:attribute name="advTm">
							<xsl:value-of select="'0'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="./@presentation:transition-type">
						<xsl:variable name="Minutes">
              <xsl:value-of select="substring-before(substring-after(./@presentation:duration,'H'),'M')"/>
						</xsl:variable>
						<xsl:variable name="Seconds">
              <xsl:value-of select="substring-before(substring-after(./@presentation:duration,'M'),'S')"/>
						</xsl:variable>
						<xsl:choose>
							<xsl:when test="./@presentation:transition-type='semi-automatic'">
								<xsl:attribute name="advClick">
									<xsl:value-of select="'0'"/>
								</xsl:attribute>
							</xsl:when>
							<xsl:when test="./@presentation:transition-type='automatic' and ./@presentation:duration">
                <xsl:choose>
                  <xsl:when test="$Minutes != '' and $Seconds != ''">
								<xsl:attribute name="advTm">
                    <xsl:value-of select="number(( number($Minutes) * 60 + number($Seconds ))* 1000)"/>
								</xsl:attribute>							
							</xsl:when>
                  <xsl:when test="contains(./@presentation:duration,'PT')">
                    <xsl:variable name="varSp2AdvTm">
                      <xsl:value-of select="number(number(substring-before(substring-after(./@presentation:duration,'PT'),'.'))*1000)"/>
                    </xsl:variable>
                    <xsl:if test="$varSp2AdvTm !='' and $varSp2AdvTm != 'NaN'">
                      <xsl:attribute name="advTm">
                        <xsl:value-of select="$varSp2AdvTm"/>
                      </xsl:attribute>
                    </xsl:if>
                  </xsl:when>
                  <xsl:otherwise>
								<xsl:attribute name="advTm">
                      <xsl:value-of select="'0'"/>
								</xsl:attribute>							
                  </xsl:otherwise>
          </xsl:choose>
	</xsl:when>
						</xsl:choose>
					</xsl:if>
					<xsl:choose>
						<!--barWipe-->
						<xsl:when test="@smil:type='barWipe' or @presentation:transition-style='barWipe'">
							<p:wipe>
								<xsl:choose>
									<xsl:when test="@smil:subtype='topToBottom' and @smil:direction='reverse'">
										<xsl:attribute name="dir">
											<xsl:value-of select="'u'"/>
										</xsl:attribute>
									</xsl:when>
									<xsl:when test="@smil:subtype='leftToRight' and not(@smil:direction)">
										<xsl:attribute name="dir">
											<xsl:value-of select="'r'"/>
										</xsl:attribute>
									</xsl:when>
									<xsl:when test="@smil:subtype='topToBottom' and not(@smil:direction)">
										<xsl:attribute name="dir">
											<xsl:value-of select="'d'"/>
										</xsl:attribute>
									</xsl:when>
									<xsl:otherwise>
										<!--do nothing , as no attribute is to be included-->
									</xsl:otherwise>									
								</xsl:choose>
							</p:wipe>
						</xsl:when>

            <xsl:when test="@smil:type='clockWipe' or @presentation:transition-style='clockWipe'">
              <p:wheel>
              <xsl:attribute name="spokes">
                <xsl:value-of select="'1'"/>
              </xsl:attribute>            
              </p:wheel>
            </xsl:when>
						<!--pinWheelWipe-->
						<xsl:when test="@smil:type='pinWheelWipe' or @presentation:transition-style='pinWheelWipe'">
							<p:wheel>
								<xsl:choose>
									<xsl:when test="@smil:subtype='oneBlade'">
										<xsl:attribute name="spokes">
											<xsl:value-of select="'1'"/>
										</xsl:attribute>
									</xsl:when>
									<xsl:when test="@smil:subtype='twoBladeVertical'">
										<xsl:attribute name="spokes">
											<xsl:value-of select="'2'"/>
										</xsl:attribute>
									</xsl:when>
									<xsl:when test="@smil:subtype='threeBlade'">
										<xsl:attribute name="spokes">
											<xsl:value-of select="'3'"/>
										</xsl:attribute>
									</xsl:when>
									<xsl:when test="@smil:subtype='eightBlade'">
										<xsl:attribute name="spokes">
											<xsl:value-of select="'8'"/>
										</xsl:attribute>
									</xsl:when>
									<xsl:otherwise>
										<xsl:attribute name="spokes">
											<xsl:value-of select="'4'"/>
										</xsl:attribute>
									</xsl:otherwise>
								</xsl:choose>
							</p:wheel>
						</xsl:when>

						<!--slideWipe-->
						<xsl:when test="@smil:type='slideWipe' or @presentation:transition-style='slideWipe'">
							<xsl:choose>
								<xsl:when test="@smil:direction='reverse'">
									<p:pull>
										<xsl:choose>
											<xsl:when test="@smil:subtype='fromTop'">
												<xsl:attribute name="dir">
													<xsl:value-of select="'d'"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:when test="@smil:subtype='fromLeft'">
												<xsl:attribute name="dir">
													<xsl:value-of select="'r'"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:when test="@smil:subtype='fromBottom'">
												<xsl:attribute name="dir">
													<xsl:value-of select="'u'"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:when test="@smil:subtype='fromTopRight'">
												<xsl:attribute name="dir">
													<xsl:value-of select="'ld'"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:when test="@smil:subtype='fromBottomRight'">
												<xsl:attribute name="dir">
													<xsl:value-of select="'lu'"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:when test="@smil:subtype='fromTopLeft'">
												<xsl:attribute name="dir">
													<xsl:value-of select="'rd'"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:when test="@smil:subtype='fromBottomLeft'">
												<xsl:attribute name="dir">
													<xsl:value-of select="'ru'"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:otherwise>
												<!--do nothing , as no attribute is to be included-->
											</xsl:otherwise>
										</xsl:choose>
									</p:pull>
								</xsl:when>
								<xsl:otherwise>
									<p:cover>
										<xsl:choose>
											<xsl:when test="@smil:subtype='fromTop'">
												<xsl:attribute name="dir">
													<xsl:value-of select="'d'"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:when test="@smil:subtype='fromLeft'">
												<xsl:attribute name="dir">
													<xsl:value-of select="'r'"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:when test="@smil:subtype='fromBottom'">
												<xsl:attribute name="dir">
													<xsl:value-of select="'u'"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:when test="@smil:subtype='fromTopRight'">
												<xsl:attribute name="dir">
													<xsl:value-of select="'ld'"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:when test="@smil:subtype='fromBottomRight'">
												<xsl:attribute name="dir">
													<xsl:value-of select="'lu'"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:when test="@smil:subtype='fromTopLeft'">
												<xsl:attribute name="dir">
													<xsl:value-of select="'rd'"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:when test="@smil:subtype='fromBottomLeft'">
												<xsl:attribute name="dir">
													<xsl:value-of select="'ru'"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:when test="@smil:subtype='fromRight'">
												<!--do nothing , as no attribute is to be included-->
											</xsl:when>
										</xsl:choose>
									</p:cover>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>

						
						<!--randomBarWipe-->
						<xsl:when test="@smil:type='randomBarWipe' or @presentation:transition-style='randomBarWipe'">
							<p:randomBar>
								<xsl:choose>
									<xsl:when test="@smil:subtype='vertical'">
										<xsl:attribute name="dir">
											<xsl:value-of select="'vert'"/>
										</xsl:attribute>
									</xsl:when>
									<xsl:otherwise>
										<!--do nothing , as no attribute is to be included-->
									</xsl:otherwise>
								</xsl:choose>
							</p:randomBar>
						</xsl:when>

						<!--checkerBoardWipe-->
						<xsl:when test="@smil:type='checkerBoardWipe' or @presentation:transition-style='checkerBoardWipe'">
							<p:checker>
								<xsl:choose>
									<xsl:when test="@smil:subtype='down'">
										<xsl:attribute name="dir">
											<xsl:value-of select="'vert'"/>
										</xsl:attribute>
									</xsl:when>
									<xsl:otherwise>
										<!--do nothing , as no attribute is to be included-->
									</xsl:otherwise>
								</xsl:choose>
							</p:checker>
						</xsl:when>

						<!--fourBoxWipe-->
						<xsl:when test="(@smil:type='fourBoxWipe' or @presentation:transition-style='fourBoxWipe') and @smil:subtype='cornersOut'">
							<p:plus/>
						</xsl:when>

						<!--irisWipe-->						
						<xsl:when test="@smil:type='irisWipe' or @presentation:transition-style='irisWipe'">
							<xsl:choose>
								<xsl:when test="@smil:subtype='rectangle'">
									<p:zoom>
										<xsl:if test="@smil:direction='reverse'">
											<xsl:attribute name="dir">
												<xsl:value-of select="'in'"/>
											</xsl:attribute>
										</xsl:if>
									</p:zoom>
								</xsl:when>
								<xsl:otherwise>
									<p:diamond/>
								</xsl:otherwise>
							</xsl:choose>							
						</xsl:when>

						<!--ellipseWipe-->
						<xsl:when test="@smil:type='ellipseWipe' or @presentation:transition-style='ellipseWipe'">
							<p:circle/>
						</xsl:when>

						<!--fanWipe-->
						<xsl:when test="@smil:type='fanWipe' or @presentation:transition-style='fanWipe'">
							<p:wedge/>
						</xsl:when>

						<!--blindsWipe-->
						<xsl:when test="@smil:type='blindsWipe' or @presentation:transition-style='blindsWipe'">
							<p:blinds>
								<xsl:choose>
									<xsl:when test="@smil:subtype='vertical'">
										<xsl:attribute name="dir">
											<xsl:value-of select="'vert'"/>
										</xsl:attribute>
									</xsl:when>
									<xsl:otherwise>
										<!--do nothing , as no attribute is to be included-->
									</xsl:otherwise>
								</xsl:choose>
							</p:blinds>
						</xsl:when>

						<!--fade-->
						<xsl:when test="@smil:type='fade' or @presentation:transition-style='fade'">
							<p:fade>
								<xsl:choose>
									<xsl:when test="@smil:subtype='fadeOverColor' or @smil:subtype='fadeFromColor'">
										<xsl:attribute name="thruBlk">
											<xsl:value-of select="'1'"/>
										</xsl:attribute>
									</xsl:when>
									<xsl:otherwise>
										<!--do nothing , as no attribute is to be included-->
									</xsl:otherwise>
								</xsl:choose>
							</p:fade>
						</xsl:when>


						<!--dissolve-->
						<xsl:when test="@smil:type='dissolve' or @presentation:transition-style='dissolve'">
							<p:dissolve/>
						</xsl:when>

						<!--random-->
						<xsl:when test="@smil:type='random' or @presentation:transition-style='random'">
							<p:random/>
						</xsl:when>

						<!--pushWipe-->
						<xsl:when test="@smil:type='pushWipe' or @presentation:transition-style='pushWipe'">
							<xsl:choose>
								<xsl:when test="starts-with(@smil:subtype,'comb')">
									<p:comb>
										<xsl:if test="@smil:subtype='combVertical'">
											<xsl:attribute name="dir">
												<xsl:value-of select="'vert'"/>
											</xsl:attribute>
										</xsl:if>
									</p:comb>
								</xsl:when>
								<xsl:otherwise>
									<p:push>
										<xsl:choose>
											<xsl:when test="@smil:subtype='fromTop'">
												<xsl:attribute name="dir">
													<xsl:value-of select="'d'"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:when test="@smil:subtype='fromLeft'">
												<xsl:attribute name="dir">
													<xsl:value-of select="'r'"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:when test="@smil:subtype='fromBottom'">
												<xsl:attribute name="dir">
													<xsl:value-of select="'u'"/>
												</xsl:attribute>
											</xsl:when>

											<xsl:otherwise>
												<!--do nothing , as no attribute is to be included-->
											</xsl:otherwise>
										</xsl:choose>
									</p:push>
								</xsl:otherwise>
							</xsl:choose>							
						</xsl:when>

						<!--barnDoorWipe-->
						<xsl:when test="@smil:type='barnDoorWipe' or @presentation:transition-style='barnDoorWipe'">
							<p:split>
								<xsl:choose>
									<xsl:when test="@smil:subtype='horizontal' and @smil:direction='reverse'">
										<xsl:attribute name="dir">
											<xsl:value-of select="'in'"/>
										</xsl:attribute>
									</xsl:when>
									<xsl:when test="@smil:subtype='vertical' and @smil:direction='reverse'">
										<xsl:attribute name="orient">
											<xsl:value-of select="'vert'"/>
										</xsl:attribute>
										<xsl:attribute name="dir">
											<xsl:value-of select="'in'"/>
										</xsl:attribute>
									</xsl:when>
									<xsl:when test="@smil:subtype='vertical' and not(@smil:direction)">
										<xsl:attribute name="orient">
											<xsl:value-of select="'vert'"/>
										</xsl:attribute>
									</xsl:when>
									<xsl:otherwise>
										<!--do nothing , as no attribute is to be included-->
									</xsl:otherwise>
								</xsl:choose>
							</p:split>
						</xsl:when>

						<!--waterfallWipe-->
						<xsl:when test="@smil:type='waterfallWipe' or @presentation:transition-style='waterfallWipe'">
							<p:strips>
								<xsl:choose>
									<xsl:when test="@smil:subtype='horizontalRight' and not(@smil:direction)">
										<xsl:attribute name="dir">
											<xsl:value-of select="'ld'"/>
										</xsl:attribute>
									</xsl:when>
									<xsl:when test="@smil:subtype='horizontalLeft' and not(@smil:direction)">
										<xsl:attribute name="dir">
											<xsl:value-of select="'rd'"/>
										</xsl:attribute>
									</xsl:when>
									<xsl:when test="@smil:subtype='horizontalRight' and @smil:direction='reverse'">
										<xsl:attribute name="dir">
											<xsl:value-of select="'ru'"/>
										</xsl:attribute>
									</xsl:when>
									<xsl:otherwise>
										<!--do nothing , as no attribute is to be included-->
									</xsl:otherwise>
								</xsl:choose>
							</p:strips>
						</xsl:when>
            <!--Code for SP2 Compat-->
            <xsl:when test="substring-after(@presentation:transition-style,'-') = 'checkerboard'">
              <p:checker>
                <xsl:choose>
                  <xsl:when test="substring-before(@presentation:transition-style,'-') = 'vertical'">
                    <xsl:attribute name="dir">
                      <xsl:value-of select="'vert'"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <!--do nothing , as no attribute is to be included-->
                  </xsl:otherwise>
                </xsl:choose>
              </p:checker>
            </xsl:when>

            <xsl:when test="@smil:type='starWipe' and @smil:subtype='fourPoint'">
              <p:plus/>
            </xsl:when>

            <!--randomBarWipe-->
            <xsl:when test="@smil:type='horizontal-lines' or @presentation:transition-style='horizontal-lines'">
              <p:randomBar/>                
            </xsl:when>

            <xsl:when test="@smil:type='vertical-lines' or @presentation:transition-style='vertical-lines'">
              <p:randomBar dir="vert"/>
            </xsl:when>
            
            <!--random-->
            <xsl:when test="@smil:type = 'random' or @presentation:transition-style = 'random'">
              <p:random/>
            </xsl:when>

            <!--End-->


					</xsl:choose>
					
				</xsl:for-each>
			</xsl:if>

			<!--sounds-->
			<xsl:variable name="styleName">
				<xsl:value-of select="./@draw:style-name"/>
			</xsl:variable>
			<xsl:variable name="relationIdVar">
				<xsl:value-of select="./@draw:name"/>
			</xsl:variable>

			<xsl:if test="//style:style[@style:family='drawing-page' and @style:name=$styleName]/style:drawing-page-properties/presentation:sound">
				<xsl:call-template name="createSoundNode">					
					<xsl:with-param name="styleName" select="$styleName"/>
					<xsl:with-param name="loopSound">
						<xsl:choose>
							<xsl:when test="./anim:par/anim:par/anim:audio/@smil:repeatCount">
								<xsl:value-of select="'true'"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="'false'"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:with-param>
				</xsl:call-template>				
			</xsl:if>						
		</p:transition>
	</xsl:template>

	<!--create sound node-->
	<xsl:template name="createSoundNode">	
		<xsl:param name="styleName"/>
		<xsl:param name="loopSound"/>
				<xsl:for-each select="//style:style[@style:family='drawing-page' and @style:name=$styleName]/style:drawing-page-properties/presentation:sound">
					<!--OpenXMl VAlidation error Fix-->
					<!--<p:sndAc>
						<p:stSnd>
					--><!--loop until next sound--><!--
					<xsl:if test="$loopSound='true'">
						<xsl:attribute name="loop">
							<xsl:value-of select="'1'"/>
						</xsl:attribute>
					</xsl:if>
							<p:snd>
								<xsl:variable name="soundfileName">
                                               <xsl:value-of select="concat(generate-id(),'.wav')"/>
								</xsl:variable>
								<xsl:attribute name="r:embed">
							<xsl:value-of select="concat('strId',generate-id())"/>
								</xsl:attribute>
								<xsl:attribute name="name">
									<xsl:value-of select="$soundfileName"/>
								</xsl:attribute>
								<xsl:attribute name="builtIn">
									<xsl:value-of select="'1'"/>
								</xsl:attribute>								
								<xsl:variable name="pzipsourcename">
									<xsl:if test="@xlink:href [ contains(.,'../')]">
										<xsl:value-of select="@xlink:href" />
									</xsl:if>
									<xsl:if test="not(@xlink:href [ contains(.,'../')])">
										<xsl:value-of select="substring-after(@xlink:href,'/')" />
									</xsl:if>
								</xsl:variable>
								<pzip:import pzip:source="{$pzipsourcename}" pzip:target="{concat('ppt/media/',$soundfileName)}" />
								--><!--<pzip:import pzip:source="{substring-after(@xlink:href,'/')}" pzip:target="{concat('ppt/media/',$soundfileName)}" />--><!--
								--><!--<pzip:copy pzip:source="{substring-after(@xlink:href,'/')}" pzip:target="{concat('ppt/media/','glasses.wav')}"/>--><!--
							</p:snd>
						</p:stSnd>
					</p:sndAc>-->
				</xsl:for-each>
	</xsl:template>

	<!--added by yeswanth.s-->
	<!--getting the sound-file name-->
	<xsl:template name="retString">
		<xsl:param name="string2rev"/>
		<xsl:choose>
			<xsl:when test="contains($string2rev,'/')">
				<xsl:call-template name="retString">
					<xsl:with-param name="string2rev" select="substring-after($string2rev,'/')"/>
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:choose>
					<xsl:when test="contains($string2rev,'.')">
						<xsl:value-of select="concat(substring-before($string2rev,'.'),'.wav')"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="concat($string2rev,'.wav')"/>
					</xsl:otherwise>					
				</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

</xsl:stylesheet>
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
  exclude-result-prefixes="odf style text number draw page p fo script presentation xlink svg">

  <xsl:template name ="slideMasters">
    <xsl:param name="slideMasterName"/>
    <xsl:param name ="smId" />
    <xsl:param name ="vmlPageNo" />

    <p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
		xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
		xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <xsl:for-each select="document('styles.xml')">
      <p:cSld>
        <!--Added by sateesh - Background Color-->
        <xsl:variable name="dpName">
            <xsl:value-of select="/office:document-styles/office:master-styles/style:master-page[@style:name=$slideMasterName]/@draw:style-name"/>
        </xsl:variable>
        <xsl:for-each select="//style:style[@style:name= $dpName]/style:drawing-page-properties">
          <xsl:choose>
            <xsl:when test="@draw:fill='bitmap'">
              <xsl:variable name="var_imageName" select="@draw:fill-image-name"/>
              <xsl:if test="/office:document-styles/office:styles/draw:fill-image[@draw:name=$var_imageName]">
              <p:bg>
                <p:bgPr>
                  <a:blipFill dpi="0" rotWithShape="1">
                    <xsl:call-template name="tmpInsertBackImage">
                      <xsl:with-param name="FileName" select="concat('slideMaster',$vmlPageNo)" />
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
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:if test="@draw:fill-color">
                          <xsl:value-of select="substring-after(@draw:fill-color,'#')" />
                        </xsl:if>
                        <xsl:if test="not(@draw:fill-color)">
                          <xsl:value-of select="'ffffff'" />
                        </xsl:if>
                      </xsl:attribute>
                    </a:srgbClr>
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
            <xsl:otherwise>
              <p:bg>
                <p:bgPr>
                  <a:solidFill>
                    <a:srgbClr val="ffffff"/>
                  </a:solidFill>
                  <a:effectLst />
                </p:bgPr>
              </p:bg>
            </xsl:otherwise>
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
          <xsl:for-each select="office:document-styles/office:master-styles/style:master-page[@style:name=$slideMasterName]">
            <xsl:for-each select="node()">
              <xsl:choose>
                <xsl:when test="name()='draw:frame'">
                  <xsl:variable name="var_pos">
                    <xsl:call-template name="getShapePosTemp">
                      <xsl:with-param name="var_pos" select="position()"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:for-each select=".">
                    <xsl:choose>
                      <xsl:when test="./draw:object or ./draw:object-ole">
                        <xsl:call-template name="tmpOLEObjects">
                          <xsl:with-param name ="pageNo" select ="$vmlPageNo"/>
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
                              <xsl:call-template name ="InsertPicture">
                                <xsl:with-param name ="imageNo" select ="concat('slideMaster',$vmlPageNo)"/>
                                <xsl:with-param name ="master" select ="'1'"/>
                                <xsl:with-param name="picNo" select="$var_pos" />
                                <xsl:with-param name ="fileName" select ="'styles.xml'"/>
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
                            </xsl:if >
                          </xsl:if >
                        </xsl:for-each>
                      </xsl:when>
                      <xsl:when test="@presentation:class[contains(.,'title')]
                                  or @presentation:class[contains(.,'outline')]">
                        <xsl:variable name ="masterPageName" select ="./parent::node()/@draw:master-page-name"/>
                        <xsl:variable name="FrameCount" select="concat('Frame',$var_pos)"/>
                        <p:sp>
                          <p:nvSpPr>
                            <xsl:choose>
                              <xsl:when test ="@presentation:class[contains(.,'title')]">
                                <p:cNvPr name="Title Placeholder 1">
                                  <xsl:attribute name="id">
                                    <xsl:value-of select="$var_pos +1"/>
                                  </xsl:attribute>
                                </p:cNvPr>
                                <p:cNvSpPr>
                                  <a:spLocks noGrp="1"/>
                                </p:cNvSpPr>
                                <p:nvPr>
                                  <p:ph type="title"/>
                                </p:nvPr>
                              </xsl:when>
                              <xsl:when test ="@presentation:class[contains(.,'outline')]">
                                <p:cNvPr name="Text Placeholder 2">
                                  <xsl:attribute name ="id">
                                    <xsl:value-of select="$var_pos +1"/>
                                  </xsl:attribute>
                                </p:cNvPr>
                                <p:cNvSpPr>
                                  <a:spLocks noGrp="1"/>
                                </p:cNvSpPr>
                                <p:nvPr>
                                  <p:ph type="body" idx="1"/>
                                </p:nvPr>
                              </xsl:when>
                            </xsl:choose>
                          </p:nvSpPr>
                          <p:spPr>
                            <xsl:call-template name ="tmpdrawCordinates"/>
                            <a:prstGeom prst="rect">
                              <a:avLst/>
                            </a:prstGeom>
                            <!-- Solid fill color -->
                            <xsl:variable name="prId">
                              <xsl:value-of select="@presentation:style-name"/>
                            </xsl:variable>
                            <xsl:choose>
                              <xsl:when test="/office:document-styles/office:automatic-styles/style:style[@style:name=$prId]">
                                <xsl:for-each select ="/office:document-styles/office:automatic-styles/style:style[@style:name=$prId] ">
                                  <xsl:variable name="styleName">
                                    <xsl:value-of select="@style:parent-style-name"/>
                                  </xsl:variable>
                                  <xsl:for-each select="/office:document-styles/office:styles/style:style[@style:name=$styleName]/style:graphic-properties">
                                    <xsl:call-template name="tmpSMShapeFillColor">
                                      <xsl:with-param name ="shapeCount" select ="$var_pos"/>
                                    </xsl:call-template>
                                  </xsl:for-each>

                                </xsl:for-each>
                              </xsl:when>
                              <xsl:when test="/office:document-styles/office:styles/style:style[@style:name=$prId]">
                                <xsl:for-each select="/office:document-styles/office:styles/style:style[@style:name=$prId]/style:graphic-properties">
                                  <xsl:call-template name="tmpSMShapeFillColor">
                                    <xsl:with-param name ="shapeCount" select ="$var_pos"/>
                                  </xsl:call-template>
                                </xsl:for-each>

                              </xsl:when>
                            </xsl:choose>
                            <!--Line Color-->
                            <xsl:choose>
                              <xsl:when test ="@presentation:class[contains(.,'title')]">
                                <xsl:choose>
                                  <xsl:when test="@presentation:style-name[contains(.,'title')]">
                                    <xsl:variable name="var_PrStyleId" select="@presentation:style-name"/>
                                    <xsl:for-each select ="/office:document-styles/office:styles/style:style[@style:name=$var_PrStyleId]/style:graphic-properties">
                                      <xsl:if test="position()=1">
                                      <xsl:call-template name ="getLineStyle"/>
                                      </xsl:if>
                                    </xsl:for-each>
                                    
                                  </xsl:when>
                                  <xsl:when test="not(@presentation:style-name[contains(.,'title')])">
                                    <xsl:variable name="var_titleName">
                                      <xsl:value-of select="concat($slideMasterName,'-title')"/>
                                    </xsl:variable>
                                    <xsl:for-each select ="/office:document-styles/office:styles/style:style[@style:name=$var_titleName]/style:graphic-properties">
                                      <xsl:if test="position()=1">
                                      <xsl:call-template name ="getLineStyle"/>
                                      </xsl:if>
                                    </xsl:for-each>
                                    
                                  </xsl:when>
                                </xsl:choose>
                              </xsl:when>
                              <xsl:when test ="@presentation:class[contains(.,'outline')]">
                                <xsl:choose>
                                  <xsl:when test="@presentation:style-name[contains(.,'outline')]">
                                    <xsl:variable name="var_PrStyleId" select="@presentation:style-name"/>
                                    <xsl:for-each select ="/office:document-styles/office:styles/style:style[@style:name=$var_PrStyleId]/style:graphic-properties">
                                      <xsl:if test="position()=1">
                                      <xsl:call-template name ="getLineStyle"/>
                                      </xsl:if>
                                    </xsl:for-each>
                                  </xsl:when>
                                  <xsl:when test="not(@presentation:style-name[contains(.,'outline')])">
                                    <xsl:variable name="var_outlineName">
                                      <xsl:value-of select="concat($slideMasterName,'-outline1')"/>
                                    </xsl:variable>
                                    <xsl:for-each select ="/office:document-styles/office:styles/style:style[@style:name=$var_outlineName]/style:graphic-properties">
                                      <xsl:if test="position()=1">
                                      <xsl:call-template name ="getLineStyle"/>
                                      </xsl:if>
                                    </xsl:for-each>
                                  </xsl:when>
                                </xsl:choose>
                              </xsl:when>
                            </xsl:choose>
                            <!--End-->
                          </p:spPr>
                          <p:txBody>
                            <xsl:call-template name ="SlideMasterTextAlignment" >
                              <xsl:with-param name="masterName" select="$slideMasterName"/>
                              <xsl:with-param name ="prId" select ="@presentation:style-name"/>
                            </xsl:call-template >
                            <!--<xsl:if test ="@presentation:class[contains(.,'outline')]">
                                <a:lstStyle />
                              </xsl:if>-->
                            <xsl:choose>
                              <xsl:when test ="@presentation:class[contains(.,'title')]">
                                <xsl:call-template name ="Titlebody">
                                  <xsl:with-param name="slideMasterName" select="$slideMasterName"/>
                                </xsl:call-template>
                              </xsl:when>
                              <xsl:when test ="@presentation:class[contains(.,'outline')]">
                                <xsl:call-template name ="Outlinebody">
                                  <xsl:with-param name="StyleName" select="@presentation:style-name"/>
                                  <xsl:with-param name="slideMasterName" select="$slideMasterName"/>
                                </xsl:call-template>
                              </xsl:when>
                            </xsl:choose>
                          </p:txBody>
                        </p:sp>
                      </xsl:when>
                      <!-- For Converting DateTime -->
                      <xsl:when test="@presentation:class ='date-time'">
                        <xsl:call-template name ="Datetime" >
                          <xsl:with-param name ="prId" select ="@presentation:style-name"/>
                          <xsl:with-param name="className" select="@presentation:class"/>
                        </xsl:call-template >
                      </xsl:when>
                      <!-- For Converting Footer -->
                      <xsl:when test="@presentation:class ='footer'">
                        <xsl:call-template name ="Footer" >
                          <xsl:with-param name ="prId" select ="@presentation:style-name"/>
                          <xsl:with-param name="className" select="@presentation:class"/>
                        </xsl:call-template >
                      </xsl:when>
                      <!-- For Converting PageNumber -->
                      <xsl:when test="@presentation:class ='page-number'">
                        <xsl:call-template name ="Pagenumber">
                          <xsl:with-param name ="prId" select ="@presentation:style-name"/>
                          <xsl:with-param name="className" select="@presentation:class"/>
                        </xsl:call-template >
                      </xsl:when>
                      <xsl:when test ="(draw:text-box) and not(@presentation:style-name) and not(@presentation:class)">
                        <xsl:call-template name ="CreateShape">
                          <xsl:with-param name ="fileName" select ="'styles.xml'"/>
                          <xsl:with-param name ="shapeName" select="'TextBox '" />
                          <xsl:with-param name ="shapeCount" select="$var_pos" />
                        </xsl:call-template>
                      </xsl:when>
                      <xsl:when test="./draw:plugin">
                        <xsl:for-each select="./draw:plugin">
                          <xsl:call-template name="InsertAudio">
                            <xsl:with-param name ="imageNo" select ="$slideMasterName"/>
                            <xsl:with-param name="AudNo" select="$var_pos" />
                          </xsl:call-template>
                        </xsl:for-each>
                      </xsl:when>
                    </xsl:choose>
                  </xsl:for-each>
                </xsl:when>

                <xsl:when test="name()='draw:rect' or name()='draw:ellipse' or name()='draw:custom-shape'
                     or name()='draw:line' or name()='draw:connector'">
                  <xsl:variable name="var_pos">
                    <xsl:call-template name="getShapePosTemp">
                      <xsl:with-param name="var_pos" select="position()"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <!-- Code for shapes start-->
                  <xsl:call-template name ="shapes" >
                    <xsl:with-param name ="fileName" select ="'styles.xml'"/>
                    <xsl:with-param name ="var_pos" select="$var_pos" />
                    <!-- Code for shapes End-->
                  </xsl:call-template >
                </xsl:when>
                <xsl:when test="name()='draw:g'">
                  <xsl:variable name="var_pos">
                    <xsl:call-template name="getShapePosTemp">
                      <xsl:with-param name="var_pos" select="position()"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:call-template name="tmpGroping">
                    <xsl:with-param name="startPos" select="'1'"/>
                    <xsl:with-param name="pageNo" select="$vmlPageNo"/>
                    <xsl:with-param name="pos" select="$var_pos"/>
                    <xsl:with-param name="fileName" select="'styles.xml'"/>
                    <xsl:with-param name ="master" select="'1'" />
                    <!--Fixed defect 1994866 PPTX-OLE- Grouped Embedded OLE objects are missing in Master -->
                    <xsl:with-param name ="vmlPageNo" select="$vmlPageNo" />
                  </xsl:call-template>
                </xsl:when>
              </xsl:choose>
            </xsl:for-each>
          </xsl:for-each >
       
        </p:spTree>
      </p:cSld>
      <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
      <p:sldLayoutIdLst>
        <!--<xsl:value-of select="2147483648 + (12 * (position() - 1 ))" />-->
        <p:sldLayoutId id="{2147483649 + (12 * ($smId - 1 ))}" r:id="rId1"/>
        <p:sldLayoutId id="{2147483650 + (12 * ($smId - 1 ))}" r:id="rId2"/>
        <p:sldLayoutId id="{2147483651 + (12 * ($smId - 1 ))}" r:id="rId3"/>
        <p:sldLayoutId id="{2147483652 + (12 * ($smId - 1 ))}" r:id="rId4"/>
        <p:sldLayoutId id="{2147483653 + (12 * ($smId - 1 ))}" r:id="rId5"/>
        <p:sldLayoutId id="{2147483654 + (12 * ($smId - 1 ))}" r:id="rId6"/>
        <p:sldLayoutId id="{2147483655 + (12 * ($smId - 1 ))}" r:id="rId7"/>
        <p:sldLayoutId id="{2147483656 + (12 * ($smId - 1 ))}" r:id="rId8"/>
        <p:sldLayoutId id="{2147483657 + (12 * ($smId - 1 ))}" r:id="rId9"/>
        <p:sldLayoutId id="{2147483658 + (12 * ($smId - 1 ))}" r:id="rId10"/>
        <p:sldLayoutId id="{2147483659 + (12 * ($smId - 1 ))}" r:id="rId11"/>


      </p:sldLayoutIdLst>
      <p:txStyles>
        <p:titleStyle>
          <a:lvl1pPr>
            <xsl:variable name="titleName">
              <xsl:value-of select="concat($slideMasterName,'-title')"/>
            </xsl:variable>
            <xsl:for-each select="/office:document-styles/office:styles/style:style[@style:name=$titleName]">
              <xsl:attribute name ="algn">
                <!--fo:text-align-->
                <xsl:choose >
                  <xsl:when test ="./style:paragraph-properties/@fo:text-align='center'">
                    <xsl:value-of select ="'ctr'"/>
                  </xsl:when>
                  <xsl:when test ="./style:paragraph-properties/@fo:text-align='end' or style:paragraph-properties/@fo:text-align='right'">
                    <xsl:value-of select ="'r'"/>
                  </xsl:when>
                  <xsl:when test ="./style:paragraph-properties/@fo:text-align='justify'">
                    <xsl:value-of select ="'just'"/>
                  </xsl:when>
                  <xsl:otherwise >
                    <xsl:value-of select ="'l'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
              <xsl:attribute name="defTabSz">
                <xsl:value-of select="'914400'"/>
              </xsl:attribute>
              <xsl:attribute name="rtl">
                <xsl:value-of select="'0'"/>
              </xsl:attribute>
              <xsl:attribute name="eaLnBrk">
                <xsl:value-of select="'1'"/>
              </xsl:attribute>
              <xsl:attribute name="latinLnBrk">
                <xsl:value-of select="'0'"/>
              </xsl:attribute>
              <xsl:attribute name="hangingPunct">
                <xsl:value-of select="'1'"/>
              </xsl:attribute>
              <xsl:call-template name ="tmpMarLeft"/>
              <!--Line Spacing-->
              <xsl:call-template name ="tmpLineSpacing"/>
              <xsl:call-template name ="tmpMarTop"/>
              <xsl:call-template name ="tmpMarBottom"/>

              <a:buNone/>
              <a:defRPr>
                <xsl:if test="./style:text-properties/@fo:font-size">
<!--Office 2007 Sp2-->

                  <xsl:variable name="fontSize">
                    <xsl:call-template name="point-measure">
                      <xsl:with-param name="length" select="style:text-properties/@fo:font-size"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:attribute name="sz">
                    <xsl:call-template name ="convertToPoints">
                      <xsl:with-param name ="unit" select ="'pt'"/>
                      <xsl:with-param name ="length" select ="concat($fontSize,'pt')"/>
                    </xsl:call-template>
                  </xsl:attribute>
                </xsl:if>
                <!--Font Bold attribute-->
                <xsl:if test="./style:text-properties/@fo:font-weight='bold'">
                  <xsl:attribute name ="b">
                    <xsl:value-of select ="'1'"/>
                  </xsl:attribute >
                </xsl:if >
                <!--Font Italic attribute-->
                <xsl:if test="./style:text-properties/@fo:font-style='italic'">
                  <xsl:attribute name ="i">
                    <xsl:value-of select ="'1'"/>
                  </xsl:attribute >
                </xsl:if >
                <!--Font Underline-->
                <xsl:variable name ="unLine">
                  <xsl:call-template name="Underline">
                    <xsl:with-param name="uStyle" select="./style:text-properties/@style:text-underline-style"/>
                    <xsl:with-param name="uWidth" select="./style:text-properties/@style:text-underline-width"/>
                    <xsl:with-param name="uType" select="./style:text-properties/@style:text-underline-type"/>
                  </xsl:call-template>
                </xsl:variable>
                <xsl:if test ="$unLine !=''">
                  <xsl:attribute name="u">
                    <xsl:value-of  select ="$unLine"/>
                  </xsl:attribute>
                </xsl:if>
                <!-- Kerning -->
                <xsl:if test ="./style:text-properties/@style:letter-kerning = 'true'">
                  <xsl:attribute name ="kern">
                    <xsl:value-of select="1200"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test ="./style:text-properties/@style:letter-kerning = 'false'">
                  <xsl:attribute name ="kern">
                    <xsl:value-of select="0"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test ="not(./style:text-properties/@style:letter-kerning)">
                  <xsl:attribute name ="kern">
                    <xsl:value-of select="0"/>
                  </xsl:attribute>
                </xsl:if>
                <!-- End -->

                <!-- Superscript and SubScript for Text added by Mathi on 1st Aug 2007-->
                <xsl:call-template name="SuperAndSubscripts" />

                <!--Character Spacing-->
                <xsl:call-template name="tmpCharacterSpacing" />

                <!--End-->
                <!-- Font Strike through Start-->
                <xsl:choose >
                  <xsl:when  test="style:text-properties/@style:text-line-through-type = 'solid'">
                    <xsl:attribute name ="strike">
                      <xsl:value-of select ="'sngStrike'"/>
                    </xsl:attribute >
                  </xsl:when >
                  <xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
                    <xsl:attribute name ="strike">
                      <xsl:value-of select ="'dblStrike'"/>
                    </xsl:attribute >
                  </xsl:when >
                  <!-- style:text-line-through-style-->
                  <xsl:when test="style:text-properties/@style:text-line-through-style = 'solid'">
                    <xsl:attribute name ="strike">
                      <xsl:value-of select ="'sngStrike'"/>
                    </xsl:attribute >
                  </xsl:when>
                </xsl:choose>
                <a:solidFill>
                  <a:srgbClr>
                    <xsl:attribute name="val">
                      <xsl:value-of select="substring-after(./style:text-properties/@fo:color,'#')" />
                      <xsl:if test="not(./style:text-properties/@fo:color)">
                        <xsl:value-of select="'000000'" />
                      </xsl:if>
                    </xsl:attribute>
                  </a:srgbClr>
                </a:solidFill>
                <!--Underline Color-->
                <xsl:if test ="style:text-properties/@style:text-underline-color !='font-color'">
                  <a:uFill>
                    <a:solidFill>
                      <a:srgbClr>
                        <xsl:attribute name ="val">
                          <xsl:value-of select ="substring-after(style:text-properties/@style:text-underline-color,'#')"/>
                        </xsl:attribute>
                      </a:srgbClr>
                    </a:solidFill>
                  </a:uFill>
                </xsl:if>
                <!--Shadow Color-->
                <xsl:if test ="style:text-properties/@fo:text-shadow != 'none'">
                  <a:effectLst>
                    <a:outerShdw blurRad="38100" dist="38100" dir="2700000" >
                      <a:srgbClr val="000000">
                        <a:alpha val="43137" />
                      </a:srgbClr>
                    </a:outerShdw>
                  </a:effectLst>
                </xsl:if>
                <!--End-->
                <a:latin>
                  <xsl:attribute name ="typeface" >
                    <xsl:value-of select="translate(./style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
                    <!--<xsl:value-of select="./style:text-properties/@fo:font-family"/>-->
                  </xsl:attribute>
                </a:latin>
                <a:ea typeface="+mj-ea"/>
                <a:cs typeface="+mj-cs"/>
              </a:defRPr>
            </xsl:for-each>
          </a:lvl1pPr>
        </p:titleStyle>
        <xsl:variable name="masterOutlineName">
          <xsl:value-of select="concat($slideMasterName,'-outline1')"/>
        </xsl:variable>
        <xsl:if test="//office:styles/style:style[@style:name=$masterOutlineName]">
          <p:bodyStyle>
            <a:lvl1pPr>
              <xsl:variable name="outlineName">
                <xsl:value-of select="concat($slideMasterName,'-outline1')"/>
              </xsl:variable>
              <xsl:for-each select="/office:document-styles/office:styles/style:style[@style:name=$outlineName]">
                <!--Margin-->
                <xsl:choose>
                  <xsl:when test="./style:paragraph-properties/@text:enable-numbering='true'">
                    <xsl:variable name="margin">
                      <xsl:call-template name="MarginTemplate">
                        <xsl:with-param name="masterName" select="$slideMasterName"/>
                        <xsl:with-param name="level" select="1"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="marL">
                      <xsl:value-of select="$margin"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="marL">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'cm'"/>
                        <xsl:with-param name ="length">
                          <xsl:call-template name="convertUnitsToCm">
                            <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
                          </xsl:call-template>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <!--End-->
                <!--Indent-->
                <xsl:choose>
                  <xsl:when test="./style:paragraph-properties/@text:enable-numbering='true'">
                    <xsl:variable name="indetValue">
                      <xsl:call-template name="IndentTemplate">
                        <xsl:with-param name="masterName" select="$slideMasterName"/>
                        <xsl:with-param name="level" select="1"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="indent">
                      <xsl:value-of select="$indetValue"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="indent">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'cm'"/>
                        <xsl:with-param name ="length">
                          <xsl:call-template name="convertUnitsToCm">
                            <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:text-indent"/>
                          </xsl:call-template>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <!--End-->
                <xsl:attribute name ="algn">
                  <!--fo:text-align-->
                  <xsl:choose >
                    <xsl:when test ="./style:paragraph-properties/@fo:text-align='center'">
                      <xsl:value-of select ="'ctr'"/>
                    </xsl:when>
                    <xsl:when test ="./style:paragraph-properties/@fo:text-align='end' or style:paragraph-properties/@fo:text-align='right'">
                      <xsl:value-of select ="'r'"/>
                    </xsl:when>
                    <xsl:when test ="./style:paragraph-properties/@fo:text-align='justify'">
                      <xsl:value-of select ="'just'"/>
                    </xsl:when>
                    <xsl:otherwise >
                      <xsl:value-of select ="'l'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
                <xsl:attribute name="defTabSz">
                  <xsl:value-of select="'914400'"/>
                </xsl:attribute>
                <xsl:attribute name="rtl">
                  <xsl:value-of select="'0'"/>
                </xsl:attribute>
                <xsl:attribute name="eaLnBrk">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
                <xsl:attribute name="latinLnBrk">
                  <xsl:value-of select="'0'"/>
                </xsl:attribute>
                <xsl:attribute name="hangingPunct">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
                <!--Line Spacing-->
                <xsl:call-template name ="tmpLineSpacing"/>
                <!--Space Before and After Paragraph-->
                <xsl:call-template name ="tmpMarTop"/>
                <xsl:call-template name ="tmpMarBottom"/>
                <!--End-->
                <xsl:if test="./style:paragraph-properties/@text:enable-numbering='true'">
                  <xsl:call-template name="SlideMasterBulletsNumbers">
                    <xsl:with-param name="masterName" select="$slideMasterName"/>
                    <xsl:with-param name="level" select="1"/>
                  </xsl:call-template>
                </xsl:if>
                <a:defRPr>
                  <xsl:if test="./style:text-properties/@fo:font-size">
<!--Office 2007 Sp2-->

                    <xsl:variable name="fontSize">
                      <xsl:call-template name="point-measure">
                        <xsl:with-param name="length" select="style:text-properties/@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="sz">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="concat($fontSize,'pt')"/>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:if>
                  <!--Font Bold attribute -->
                  <xsl:variable name ="boldStyle">
                    <xsl:call-template name ="BoldItalic">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="1"/>
                      <xsl:with-param name ="paramName" select ="'b'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$boldStyle ='bold'">
                    <xsl:attribute name ="b">
                      <xsl:value-of  select ="1"/>
                    </xsl:attribute >
                  </xsl:if>
                  <!-- Font Italic attribute-->
                  <xsl:variable name ="italicStyle">
                    <xsl:call-template name ="BoldItalic">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="1"/>
                      <xsl:with-param name ="paramName" select ="'i'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$italicStyle ='italic'">
                    <xsl:attribute name ="i">
                      <xsl:value-of  select ="1"/>
                    </xsl:attribute >
                  </xsl:if>
                  <!-- Font Underline-->
                  <xsl:choose>
                    <xsl:when test="./style:text-properties/@style:text-underline-style !='none'">
                      <xsl:variable name ="unLine">
                        <xsl:call-template name ="ULine">
                          <xsl:with-param name="smName" select="$slideMasterName"/>
                          <xsl:with-param name ="outlineNo"  select ="1"/>
                          <xsl:with-param name ="paramName" select ="'u'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$unLine !=''">
                        <xsl:attribute name="u">
                          <xsl:value-of  select ="$unLine"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                    <xsl:when test="./style:text-properties/@style:text-underline-style ='none'">
                      <xsl:attribute name="u">
                        <xsl:value-of select="'none'"/>
                      </xsl:attribute>
                    </xsl:when>
                  </xsl:choose>
                  <!-- Kerning -->
                  <xsl:if test ="./style:text-properties/@style:letter-kerning = 'true'">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="1200"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="./style:text-properties/@style:letter-kerning = 'false'">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="0"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="not(./style:text-properties/@style:letter-kerning)">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="0"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!-- End -->
                  <!-- Superscript and SubScript for Text added by Mathi on 1st Aug 2007-->
                  <xsl:call-template name="SuperAndSubscripts" />

                  <!--Character Spacing-->
                  <xsl:call-template name="tmpCharacterSpacing" />
                  <!--End-->
                  <!-- Font Strike through Start-->
                  <xsl:choose >
                    <xsl:when  test="style:text-properties/@style:text-line-through-type = 'solid'">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'sngStrike'"/>
                      </xsl:attribute >
                    </xsl:when >
                    <xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'dblStrike'"/>
                      </xsl:attribute >
                    </xsl:when >
                    <!-- style:text-line-through-style-->
                    <xsl:when test="style:text-properties/@style:text-line-through-style = 'solid'">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'sngStrike'"/>
                      </xsl:attribute >
                    </xsl:when>
                  </xsl:choose>
                  <a:solidFill>
                    <a:srgbClr>
                      <xsl:choose>
                        <xsl:when test="./style:text-properties/@fo:color">
                          <xsl:attribute name="val">
                            <xsl:value-of select="substring-after(./style:text-properties/@fo:color,'#')" />
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:attribute name="val">
                            <xsl:value-of select="'000000'" />
                          </xsl:attribute>
                        </xsl:otherwise>
                      </xsl:choose>
                    </a:srgbClr>
                  </a:solidFill>
                  <!--Underline Color-->
                  <xsl:if test ="style:text-properties/@style:text-underline-color !='font-color'">
                    <a:uFill>
                      <a:solidFill>
                        <a:srgbClr>
                          <xsl:attribute name ="val">
                            <xsl:value-of select ="substring-after(style:text-properties/@style:text-underline-color,'#')"/>
                          </xsl:attribute>
                        </a:srgbClr>
                      </a:solidFill>
                    </a:uFill>
                  </xsl:if>
                  <!--end-->
                  <!--Shadow Color-->
                  <xsl:if test ="style:text-properties/@fo:text-shadow != 'none'">
                    <a:effectLst>
                      <a:outerShdw blurRad="38100" dist="38100" dir="2700000" >
                        <a:srgbClr val="000000">
                          <a:alpha val="43137" />
                        </a:srgbClr>
                      </a:outerShdw>
                    </a:effectLst>
                  </xsl:if>
                  <!--End-->
                  <a:latin>
                    <xsl:attribute name="typeface">
                      <xsl:value-of select="translate(./style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
                    </xsl:attribute>
                  </a:latin>
                  <a:ea typeface="+mn-ea"/>
                  <a:cs typeface="+mn-cs"/>
                </a:defRPr>
              </xsl:for-each>
            </a:lvl1pPr>
            <a:lvl2pPr>
              <xsl:variable name="outlineName">
                <xsl:value-of select="concat($slideMasterName,'-outline1')"/>
              </xsl:variable>
              <xsl:variable name="outlineName2">
                <xsl:value-of select="concat($slideMasterName,'-outline2')"/>
              </xsl:variable>
              <xsl:for-each select="/office:document-styles/office:styles/style:style[@style:name=$outlineName2]">
                <xsl:choose>
                  <xsl:when test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                    <xsl:variable name="margin">
                      <xsl:call-template name="MarginTemplate">
                        <xsl:with-param name="masterName" select="$slideMasterName"/>
                        <xsl:with-param name="level" select="2"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="marL">
                      <xsl:value-of select="$margin"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="marL">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'cm'"/>
                        <xsl:with-param name ="length">
                          <xsl:call-template name="convertUnitsToCm">
                            <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
                          </xsl:call-template>
                        </xsl:with-param>
                      </xsl:call-template>

                    </xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <!--End-->
                <!--Indent-->
                <xsl:choose>
                  <xsl:when test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                    <xsl:variable name="indetValue">
                      <xsl:call-template name="IndentTemplate">
                        <xsl:with-param name="masterName" select="$slideMasterName"/>
                        <xsl:with-param name="level" select="2"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="indent">
                      <xsl:value-of select="$indetValue"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="indent">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'cm'"/>
                        <xsl:with-param name ="length">
                          <xsl:call-template name="convertUnitsToCm">
                            <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:text-indent"/>
                          </xsl:call-template>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <!--End-->
                <!--fo:text-align-->
                <xsl:attribute name ="algn">
                  <xsl:variable name ="alignment">
                    <xsl:call-template name ="Alignment">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="2"/>
                      <xsl:with-param name ="paramName" select ="'algn'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$alignment">
                    <xsl:value-of select="$alignment"/>
                  </xsl:if>
                </xsl:attribute>
                <!--End-->
                <xsl:attribute name="defTabSz">
                  <xsl:value-of select="'914400'"/>
                </xsl:attribute>
                <xsl:attribute name="rtl">
                  <xsl:value-of select="'0'"/>
                </xsl:attribute>
                <xsl:attribute name="eaLnBrk">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
                <xsl:attribute name="latinLnBrk">
                  <xsl:value-of select="'0'"/>
                </xsl:attribute>
                <xsl:attribute name="hangingPunct">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
                <!--Line Spacing-->
                <xsl:call-template name ="LineSpacingPoints">
                  <xsl:with-param name="smasterName" select="$slideMasterName"/>
                  <xsl:with-param name ="lineNo"  select ="2"/>
                </xsl:call-template>
                <!--End-->
                <!--Space Before and After Paragraph-->

                <xsl:call-template name ="tmpMarTop"/>
                <xsl:call-template name ="tmpMarBottom"/>
                <!--End-->
                <!--Bullets And Numbering Code-->
                <!--<xsl:if test="document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">-->
                <xsl:if test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                  <xsl:call-template name="SlideMasterBulletsNumbers">
                    <xsl:with-param name="masterName" select="$slideMasterName"/>
                    <xsl:with-param name="level" select="2"/>
                  </xsl:call-template>
                </xsl:if>
                <!--End-->
                <a:defRPr>
                  <xsl:if test="./style:text-properties/@fo:font-size">
<!--Office 2007 Sp2-->

                    <xsl:variable name="fontSize">
                      <xsl:call-template name="point-measure">
                        <xsl:with-param name="length" select="style:text-properties/@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="sz">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="concat($fontSize,'pt')"/>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:if>
                  <!--Font Bold attribute -->
                  <xsl:variable name ="boldStyle">
                    <xsl:call-template name ="BoldItalic">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="2"/>
                      <xsl:with-param name ="paramName" select ="'b'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$boldStyle ='bold'">
                    <xsl:attribute name ="b">
                      <xsl:value-of  select ="1"/>
                    </xsl:attribute >
                  </xsl:if>
                  <!-- Font Italic attribute-->
                  <xsl:variable name ="italicStyle">
                    <xsl:call-template name ="BoldItalic">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="2"/>
                      <xsl:with-param name ="paramName" select ="'i'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$italicStyle ='italic'">
                    <xsl:attribute name ="i">
                      <xsl:value-of  select ="1"/>
                    </xsl:attribute >
                  </xsl:if>
                  <!-- Font Underline-->
                  <xsl:choose>
                    <xsl:when test="./style:text-properties/@style:text-underline-style !='none'">
                      <xsl:variable name ="unLine">
                        <xsl:call-template name ="ULine">
                          <xsl:with-param name="smName" select="$slideMasterName"/>
                          <xsl:with-param name ="outlineNo"  select ="2"/>
                          <xsl:with-param name ="paramName" select ="'u'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$unLine != ''">
                        <xsl:attribute name="u">
                          <xsl:value-of  select ="$unLine"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                    <xsl:when test="./style:text-properties/@style:text-underline-style ='none'">
                      <xsl:attribute name="u">
                        <xsl:value-of select="'none'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test="not(./style:text-properties/style:text-underline-style)">
                      <xsl:variable name ="unLine">
                        <xsl:call-template name ="ULine">
                          <xsl:with-param name="smName" select="$slideMasterName"/>
                          <xsl:with-param name ="outlineNo"  select ="2"/>
                          <xsl:with-param name ="paramName" select ="'u'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$unLine !=''">
                        <xsl:attribute name="u">
                          <xsl:value-of  select ="$unLine"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                  </xsl:choose>
                  <!--End-->
                  <!-- Kerning -->
                  <xsl:if test ="./style:text-properties/@style:letter-kerning = 'true'">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="1200"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="./style:text-properties/@style:letter-kerning = 'false'">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="0"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="not(./style:text-properties/@style:letter-kerning)">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="0"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!-- End -->
                  <!-- Superscript and SubScript for Text added by Mathi on 1st Aug 2007-->
                  <xsl:call-template name="SuperAndSubscripts" />

                  <!--Character Spacing-->
                  <xsl:variable name ="space">
                    <xsl:call-template name ="Spacing">
                      <xsl:with-param name="smasterName" select="$slideMasterName"/>
                      <xsl:with-param name ="lineNo"  select ="2"/>
                      <xsl:with-param name ="parName" select ="'spc'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$space !='' ">
                    <xsl:attribute name="spc">
                      <xsl:value-of  select ="$space"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!--End-->
                  <!-- Font Strike through Start-->
                  <xsl:choose >
                    <xsl:when  test="style:text-properties/@style:text-line-through-type = 'solid'">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'sngStrike'"/>
                      </xsl:attribute >
                    </xsl:when >
                    <xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'dblStrike'"/>
                      </xsl:attribute >
                    </xsl:when >
                    <!-- style:text-line-through-style-->
                    <xsl:when test="style:text-properties/@style:text-line-through-style = 'solid'">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'sngStrike'"/>
                      </xsl:attribute >
                    </xsl:when>
                  </xsl:choose>
                  <!--Font Color-->
                  <a:solidFill>
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:variable name ="color">
                          <xsl:call-template name ="FontColor">
                            <xsl:with-param name="smasterName" select="$slideMasterName"/>
                            <xsl:with-param name ="lineNo"  select ="2"/>
                            <xsl:with-param name ="parName" select ="'val'"/>
                          </xsl:call-template>
                        </xsl:variable>
                        <xsl:if test ="$color">
                          <xsl:value-of  select ="$color"/>
                        </xsl:if>
                      </xsl:attribute>
                    </a:srgbClr>
                  </a:solidFill>
                  <!--end-->
                  <!--Underline Color-->
                  <xsl:if test ="style:text-properties/@style:text-underline-color !='font-color'">
                    <a:uFill>
                      <a:solidFill>
                        <a:srgbClr>
                          <xsl:attribute name ="val">
                            <xsl:value-of select ="substring-after(style:text-properties/@style:text-underline-color,'#')"/>
                          </xsl:attribute>
                        </a:srgbClr>
                      </a:solidFill>
                    </a:uFill>
                  </xsl:if>
                  <!--end-->
                  <!--Shadow Color-->
                  <xsl:if test ="style:text-properties/@fo:text-shadow != 'none'">
                    <a:effectLst>
                      <a:outerShdw blurRad="38100" dist="38100" dir="2700000" >
                        <a:srgbClr val="000000">
                          <a:alpha val="43137" />
                        </a:srgbClr>
                      </a:outerShdw>
                    </a:effectLst>
                  </xsl:if>
                  <!--End-->
                  <!--Font Family-->
                  <a:latin>
                    <xsl:attribute name="typeface">
                      <xsl:variable name ="foFamily">
                        <xsl:call-template name ="FontFamily">
                          <xsl:with-param name="smasterName" select="$slideMasterName"/>
                          <xsl:with-param name ="lineNo"  select ="2"/>
                          <xsl:with-param name ="parName" select ="'typeface'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$foFamily">
                        <xsl:value-of  select ="$foFamily"/>
                      </xsl:if>
                    </xsl:attribute>
                  </a:latin>
                  <a:ea typeface="+mn-ea"/>
                  <a:cs typeface="+mn-cs"/>
                </a:defRPr>
              </xsl:for-each>
            </a:lvl2pPr>
            <a:lvl3pPr>
              <xsl:variable name="outlineName">
                <xsl:value-of select="concat($slideMasterName,'-outline1')"/>
              </xsl:variable>
              <xsl:variable name="outlineName3">
                <xsl:value-of select="concat($slideMasterName,'-outline3')"/>
              </xsl:variable>
              <xsl:for-each select="/office:document-styles/office:styles/style:style[@style:name=$outlineName3]">
                <!--<xsl:attribute name="marL">
                <xsl:call-template name ="convertToPoints">
                  <xsl:with-param name ="unit" select ="'cm'"/>
                  <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
                </xsl:call-template>
              </xsl:attribute>-->
                <!--Margin-->
                <xsl:choose>
                  <xsl:when test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                    <xsl:variable name="margin">
                      <xsl:call-template name="MarginTemplate">
                        <xsl:with-param name="masterName" select="$slideMasterName"/>
                        <xsl:with-param name="level" select="3"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="marL">
                      <xsl:value-of select="$margin"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="marL">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'cm'"/>
                        <xsl:with-param name ="length">
                          <xsl:call-template name="convertUnitsToCm">
                            <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
                          </xsl:call-template>
                        </xsl:with-param>
                      </xsl:call-template>

                    </xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <!--End-->
                <!--Indent-->
                <xsl:choose>
                  <xsl:when test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                    <xsl:variable name="indetValue">
                      <xsl:call-template name="IndentTemplate">
                        <xsl:with-param name="masterName" select="$slideMasterName"/>
                        <xsl:with-param name="level" select="3"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="indent">
                      <xsl:value-of select="$indetValue"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="indent">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'cm'"/>
                        <xsl:with-param name ="length">
                          <xsl:call-template name="convertUnitsToCm">
                            <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:text-indent"/>
                          </xsl:call-template>
                        </xsl:with-param>
                      </xsl:call-template>

                    </xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <!--End-->
                <!--fo:text-align-->
                <xsl:attribute name ="algn">
                  <xsl:variable name ="alignment">
                    <xsl:call-template name ="Alignment">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="3"/>
                      <xsl:with-param name ="paramName" select ="'algn'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$alignment">
                    <xsl:value-of select="$alignment"/>
                  </xsl:if>
                </xsl:attribute>
                <!--End-->
                <xsl:attribute name="defTabSz">
                  <xsl:value-of select="'914400'"/>
                </xsl:attribute>
                <xsl:attribute name="rtl">
                  <xsl:value-of select="'0'"/>
                </xsl:attribute>
                <xsl:attribute name="eaLnBrk">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
                <xsl:attribute name="latinLnBrk">
                  <xsl:value-of select="'0'"/>
                </xsl:attribute>
                <xsl:attribute name="hangingPunct">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
                <!--Line Spacing-->
                <xsl:call-template name ="LineSpacingPoints">
                  <xsl:with-param name="smasterName" select="$slideMasterName"/>
                  <xsl:with-param name ="lineNo"  select ="3"/>
                </xsl:call-template>
                <!--End-->
                <!--Space Before and After Paragraph-->

                <xsl:call-template name ="tmpMarTop"/>
                <xsl:call-template name ="tmpMarBottom"/>
                <!--End-->
                <xsl:if test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                  <xsl:call-template name="SlideMasterBulletsNumbers">
                    <xsl:with-param name="masterName" select="$slideMasterName"/>
                    <xsl:with-param name="level" select="3"/>
                  </xsl:call-template>
                </xsl:if>
                <a:defRPr>
                  <xsl:if test="./style:text-properties/@fo:font-size">
<!--Office 2007 Sp2-->

                    <xsl:variable name="fontSize">
                      <xsl:call-template name="point-measure">
                        <xsl:with-param name="length" select="style:text-properties/@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="sz">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="concat($fontSize,'pt')"/>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:if>

                  <!--Font Bold attribute -->
                  <xsl:variable name ="boldStyle">
                    <xsl:call-template name ="BoldItalic">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="3"/>
                      <xsl:with-param name ="paramName" select ="'b'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$boldStyle ='bold'">
                    <xsl:attribute name ="b">
                      <xsl:value-of  select ="1"/>
                    </xsl:attribute >
                  </xsl:if>
                  <!-- Font Italic attribute-->
                  <xsl:variable name ="italicStyle">
                    <xsl:call-template name ="BoldItalic">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="3"/>
                      <xsl:with-param name ="paramName" select ="'i'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$italicStyle ='italic'">
                    <xsl:attribute name ="i">
                      <xsl:value-of  select ="1"/>
                    </xsl:attribute >
                  </xsl:if>
                  <!-- Font Underline-->
                  <xsl:choose>
                    <xsl:when test="./style:text-properties/@style:text-underline-style !='none'">
                      <xsl:variable name ="unLine">
                        <xsl:call-template name ="ULine">
                          <xsl:with-param name="smName" select="$slideMasterName"/>
                          <xsl:with-param name ="outlineNo"  select ="3"/>
                          <xsl:with-param name ="paramName" select ="'u'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$unLine !=''">
                        <xsl:attribute name="u">
                          <xsl:value-of  select ="$unLine"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                    <xsl:when test="./style:text-properties/@style:text-underline-style ='none'">
                      <xsl:attribute name="u">
                        <xsl:value-of select="'none'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test="not(./style:text-properties/style:text-underline-style)">
                      <xsl:variable name ="unLine">
                        <xsl:call-template name ="ULine">
                          <xsl:with-param name="smName" select="$slideMasterName"/>
                          <xsl:with-param name ="outlineNo"  select ="3"/>
                          <xsl:with-param name ="paramName" select ="'u'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$unLine !=''">
                        <xsl:attribute name="u">
                          <xsl:value-of  select ="$unLine"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                  </xsl:choose>
                  <!--end-->
                  <!-- Kerning -->
                  <xsl:if test ="./style:text-properties/@style:letter-kerning = 'true'">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="1200"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="./style:text-properties/@style:letter-kerning = 'false'">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="0"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="not(./style:text-properties/@style:letter-kerning)">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="0"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!-- End -->
                  <!-- Superscript and SubScript for Text added by Mathi on 1st Aug 2007-->
                  <xsl:call-template name="SuperAndSubscripts" />

                  <!--Character Spacing-->
                  <xsl:variable name ="space">
                    <xsl:call-template name ="Spacing">
                      <xsl:with-param name="smasterName" select="$slideMasterName"/>
                      <xsl:with-param name ="lineNo"  select ="3"/>
                      <xsl:with-param name ="parName" select ="'spc'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$space !='' ">
                    <xsl:attribute name="spc">
                      <xsl:value-of  select ="$space"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!--End-->
                  <!-- Font Strike through Start-->
                  <xsl:choose >
                    <xsl:when  test="style:text-properties/@style:text-line-through-type = 'solid'">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'sngStrike'"/>
                      </xsl:attribute >
                    </xsl:when >
                    <xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'dblStrike'"/>
                      </xsl:attribute >
                    </xsl:when >
                    <!-- style:text-line-through-style-->
                    <xsl:when test="style:text-properties/@style:text-line-through-style = 'solid'">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'sngStrike'"/>
                      </xsl:attribute >
                    </xsl:when>
                  </xsl:choose>
                  <!--Font Color-->
                  <a:solidFill>
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:variable name ="color">
                          <xsl:call-template name ="FontColor">
                            <xsl:with-param name="smasterName" select="$slideMasterName"/>
                            <xsl:with-param name ="lineNo"  select ="3"/>
                            <xsl:with-param name ="parName" select ="'val'"/>
                          </xsl:call-template>
                        </xsl:variable>
                        <xsl:if test ="$color">
                          <xsl:value-of  select ="$color"/>
                        </xsl:if>
                      </xsl:attribute>
                    </a:srgbClr>
                  </a:solidFill>
                  <!--end-->
                  <!--Underline Color-->
                  <xsl:if test ="style:text-properties/@style:text-underline-color !='font-color'">
                    <a:uFill>
                      <a:solidFill>
                        <a:srgbClr>
                          <xsl:attribute name ="val">
                            <xsl:value-of select ="substring-after(style:text-properties/@style:text-underline-color,'#')"/>
                          </xsl:attribute>
                        </a:srgbClr>
                      </a:solidFill>
                    </a:uFill>
                  </xsl:if>
                  <!--end-->
                  <!--Shadow Color-->
                  <xsl:if test ="style:text-properties/@fo:text-shadow != 'none'">
                    <a:effectLst>
                      <a:outerShdw blurRad="38100" dist="38100" dir="2700000" >
                        <a:srgbClr val="000000">
                          <a:alpha val="43137" />
                        </a:srgbClr>
                      </a:outerShdw>
                    </a:effectLst>
                  </xsl:if>
                  <!--End-->
                  <!--Font Family-->
                  <a:latin>
                    <xsl:attribute name="typeface">
                      <xsl:variable name ="foFamily">
                        <xsl:call-template name ="FontFamily">
                          <xsl:with-param name="smasterName" select="$slideMasterName"/>
                          <xsl:with-param name ="lineNo"  select ="3"/>
                          <xsl:with-param name ="parName" select ="'typeface'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$foFamily">
                        <xsl:value-of  select ="$foFamily"/>
                      </xsl:if>
                    </xsl:attribute>
                  </a:latin>
                  <a:ea typeface="+mn-ea"/>
                  <a:cs typeface="+mn-cs"/>
                </a:defRPr>
              </xsl:for-each>

            </a:lvl3pPr>
            <a:lvl4pPr>
              <xsl:variable name="outlineName">
                <xsl:value-of select="concat($slideMasterName,'-outline1')"/>
              </xsl:variable>
              <xsl:variable name="outlineName4">
                <xsl:value-of select="concat($slideMasterName,'-outline4')"/>
              </xsl:variable>
              <xsl:for-each select="/office:document-styles/office:styles/style:style[@style:name=$outlineName4]">
                <!--<xsl:attribute name="marL">
                <xsl:call-template name ="convertToPoints">
                  <xsl:with-param name ="unit" select ="'cm'"/>
                  <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
                </xsl:call-template>
              </xsl:attribute>-->
                <!--Margin-->
                <xsl:choose>
                  <xsl:when test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                    <xsl:variable name="margin">
                      <xsl:call-template name="MarginTemplate">
                        <xsl:with-param name="masterName" select="$slideMasterName"/>
                        <xsl:with-param name="level" select="4"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="marL">
                      <xsl:value-of select="$margin"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="marL">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'cm'"/>
                        <xsl:with-param name ="length">
                          <xsl:call-template name="convertUnitsToCm">
                            <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
                          </xsl:call-template>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <!--End-->
                <!--Indent-->
                <xsl:choose>
                  <xsl:when test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                    <xsl:variable name="indetValue">
                      <xsl:call-template name="IndentTemplate">
                        <xsl:with-param name="masterName" select="$slideMasterName"/>
                        <xsl:with-param name="level" select="4"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="indent">
                      <xsl:value-of select="$indetValue"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="indent">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'cm'"/>
                        <xsl:with-param name ="length">
                          <xsl:call-template name="convertUnitsToCm">
                            <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:text-indent"/>
                          </xsl:call-template>
                        </xsl:with-param>
                      </xsl:call-template>

                    </xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <!--End-->
                <!--fo:text-align-->
                <xsl:attribute name ="algn">
                  <xsl:variable name ="alignment">
                    <xsl:call-template name ="Alignment">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="4"/>
                      <xsl:with-param name ="paramName" select ="'algn'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$alignment">
                    <xsl:value-of select="$alignment"/>
                  </xsl:if>
                </xsl:attribute>
                <!--End-->
                <xsl:attribute name="defTabSz">
                  <xsl:value-of select="'914400'"/>
                </xsl:attribute>
                <xsl:attribute name="rtl">
                  <xsl:value-of select="'0'"/>
                </xsl:attribute>
                <xsl:attribute name="eaLnBrk">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
                <xsl:attribute name="latinLnBrk">
                  <xsl:value-of select="'0'"/>
                </xsl:attribute>
                <xsl:attribute name="hangingPunct">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
                <!--Line Spacing-->
                <xsl:call-template name ="LineSpacingPoints">
                  <xsl:with-param name="smasterName" select="$slideMasterName"/>
                  <xsl:with-param name ="lineNo"  select ="4"/>
                </xsl:call-template>
                <!--End-->
                <!--Space Before and After Paragraph-->
                <xsl:call-template name ="tmpMarTop"/>
                <xsl:call-template name ="tmpMarBottom"/>
                <!--End-->
                <xsl:if test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                  <xsl:call-template name="SlideMasterBulletsNumbers">
                    <xsl:with-param name="masterName" select="$slideMasterName"/>
                    <xsl:with-param name="level" select="4"/>
                  </xsl:call-template>
                </xsl:if>
                <a:defRPr>
                  <xsl:if test="./style:text-properties/@fo:font-size">
<!--Office 2007 Sp2-->

                    <xsl:variable name="fontSize">
                      <xsl:call-template name="point-measure">
                        <xsl:with-param name="length" select="style:text-properties/@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="sz">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="concat($fontSize,'pt')"/>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:if>
                  <!--Font Bold attribute -->

                  <xsl:variable name ="boldStyle">
                    <xsl:call-template name ="BoldItalic">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="4"/>
                      <xsl:with-param name ="paramName" select ="'b'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$boldStyle ='bold'">
                    <xsl:attribute name ="b">
                      <xsl:value-of  select ="1"/>
                    </xsl:attribute >
                  </xsl:if>
                  <!-- Font Italic attribute-->
                  <xsl:variable name ="italicStyle">
                    <xsl:call-template name ="BoldItalic">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="4"/>
                      <xsl:with-param name ="paramName" select ="'i'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$italicStyle ='italic'">
                    <xsl:attribute name ="i">
                      <xsl:value-of  select ="1"/>
                    </xsl:attribute >
                  </xsl:if>
                  <!-- Font Underline-->
                  <xsl:choose>
                    <xsl:when test="./style:text-properties/@style:text-underline-style !='none'">
                      <xsl:variable name ="unLine">
                        <xsl:call-template name ="ULine">
                          <xsl:with-param name="smName" select="$slideMasterName"/>
                          <xsl:with-param name ="outlineNo"  select ="4"/>
                          <xsl:with-param name ="paramName" select ="'u'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$unLine !=''">
                        <xsl:attribute name="u">
                          <xsl:value-of  select ="$unLine"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                    <xsl:when test="./style:text-properties/@style:text-underline-style ='none'">
                      <xsl:attribute name="u">
                        <xsl:value-of select="'none'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test="not(./style:text-properties/style:text-underline-style)">
                      <xsl:variable name ="unLine">
                        <xsl:call-template name ="ULine">
                          <xsl:with-param name="smName" select="$slideMasterName"/>
                          <xsl:with-param name ="outlineNo"  select ="4"/>
                          <xsl:with-param name ="paramName" select ="'u'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$unLine !=''">
                        <xsl:attribute name="u">
                          <xsl:value-of  select ="$unLine"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                  </xsl:choose>
                  <!-- Kerning -->
                  <xsl:if test ="./style:text-properties/@style:letter-kerning = 'true'">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="1200"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="./style:text-properties/@style:letter-kerning = 'false'">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="0"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="not(./style:text-properties/@style:letter-kerning)">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="0"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!-- End -->
                  <!-- Superscript and SubScript for Text added by Mathi on 1st Aug 2007-->
                  <xsl:call-template name="SuperAndSubscripts" />

                  <!--Character Spacing-->
                  <xsl:variable name ="space">
                    <xsl:call-template name ="Spacing">
                      <xsl:with-param name="smasterName" select="$slideMasterName"/>
                      <xsl:with-param name ="lineNo"  select ="4"/>
                      <xsl:with-param name ="parName" select ="'spc'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$space !='' ">
                    <xsl:attribute name="spc">
                      <xsl:value-of  select ="$space"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!--End-->
                  <!-- Font Strike through Start-->
                  <xsl:choose >
                    <xsl:when  test="style:text-properties/@style:text-line-through-type = 'solid'">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'sngStrike'"/>
                      </xsl:attribute >
                    </xsl:when >
                    <xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'dblStrike'"/>
                      </xsl:attribute >
                    </xsl:when >
                    <!-- style:text-line-through-style-->
                    <xsl:when test="style:text-properties/@style:text-line-through-style = 'solid'">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'sngStrike'"/>
                      </xsl:attribute >
                    </xsl:when>
                  </xsl:choose>
                  <!--Font Color-->
                  <a:solidFill>
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:variable name ="color">
                          <xsl:call-template name ="FontColor">
                            <xsl:with-param name="smasterName" select="$slideMasterName"/>
                            <xsl:with-param name ="lineNo"  select ="4"/>
                            <xsl:with-param name ="parName" select ="'val'"/>
                          </xsl:call-template>
                        </xsl:variable>
                        <xsl:if test ="$color">
                          <xsl:value-of  select ="$color"/>
                        </xsl:if>
                      </xsl:attribute>
                    </a:srgbClr>
                  </a:solidFill>
                  <!--end-->
                  <!--Underline Color-->
                  <xsl:if test ="style:text-properties/@style:text-underline-color !='font-color'">
                    <a:uFill>
                      <a:solidFill>
                        <a:srgbClr>
                          <xsl:attribute name ="val">
                            <xsl:value-of select ="substring-after(style:text-properties/@style:text-underline-color,'#')"/>
                          </xsl:attribute>
                        </a:srgbClr>
                      </a:solidFill>
                    </a:uFill>
                  </xsl:if>
                  <!--end-->
                  <!--Shadow Color-->
                  <xsl:if test ="style:text-properties/@fo:text-shadow != 'none'">
                    <a:effectLst>
                      <a:outerShdw blurRad="38100" dist="38100" dir="2700000" >
                        <a:srgbClr val="000000">
                          <a:alpha val="43137" />
                        </a:srgbClr>
                      </a:outerShdw>
                    </a:effectLst>
                  </xsl:if>
                  <!--End-->
                  <!--Font Family-->
                  <a:latin>
                    <xsl:attribute name="typeface">
                      <xsl:variable name ="foFamily">
                        <xsl:call-template name ="FontFamily">
                          <xsl:with-param name="smasterName" select="$slideMasterName"/>
                          <xsl:with-param name ="lineNo"  select ="4"/>
                          <xsl:with-param name ="parName" select ="'typeface'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$foFamily">
                        <xsl:value-of  select ="$foFamily"/>
                      </xsl:if>
                    </xsl:attribute>
                  </a:latin>
                  <a:ea typeface="+mn-ea"/>
                  <a:cs typeface="+mn-cs"/>
                </a:defRPr>
              </xsl:for-each>
            </a:lvl4pPr>
            <a:lvl5pPr>
              <xsl:variable name="outlineName">
                <xsl:value-of select="concat($slideMasterName,'-outline1')"/>
              </xsl:variable>
              <xsl:variable name="outlineName5">
                <xsl:value-of select="concat($slideMasterName,'-outline5')"/>
              </xsl:variable>
              <xsl:for-each select="/office:document-styles/office:styles/style:style[@style:name=$outlineName5]">
                <!--Margin-->
                <xsl:choose>
                  <xsl:when test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                    <xsl:variable name="margin">
                      <xsl:call-template name="MarginTemplate">
                        <xsl:with-param name="masterName" select="$slideMasterName"/>
                        <xsl:with-param name="level" select="5"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="marL">
                      <xsl:value-of select="$margin"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="marL">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'cm'"/>
                        <xsl:with-param name ="length">
                          <xsl:call-template name="convertUnitsToCm">
                            <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
                          </xsl:call-template>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <!--End-->
                <!--Indent-->
                <xsl:choose>
                  <xsl:when test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                    <xsl:variable name="indetValue">
                      <xsl:call-template name="IndentTemplate">
                        <xsl:with-param name="masterName" select="$slideMasterName"/>
                        <xsl:with-param name="level" select="5"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="indent">
                      <xsl:value-of select="$indetValue"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="indent">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'cm'"/>
                        <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:text-indent"/>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <!--End-->
                <!--fo:text-align-->
                <xsl:attribute name ="algn">
                  <xsl:variable name ="alignment">
                    <xsl:call-template name ="Alignment">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="5"/>
                      <xsl:with-param name ="paramName" select ="'algn'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$alignment">
                    <xsl:value-of select="$alignment"/>
                  </xsl:if>
                </xsl:attribute>
                <!--End-->
                <xsl:attribute name="defTabSz">
                  <xsl:value-of select="'914400'"/>
                </xsl:attribute>
                <xsl:attribute name="rtl">
                  <xsl:value-of select="'0'"/>
                </xsl:attribute>
                <xsl:attribute name="eaLnBrk">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
                <xsl:attribute name="latinLnBrk">
                  <xsl:value-of select="'0'"/>
                </xsl:attribute>
                <xsl:attribute name="hangingPunct">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
                <!--Line Spacing-->
                <xsl:call-template name ="LineSpacingPoints">
                  <xsl:with-param name="smasterName" select="$slideMasterName"/>
                  <xsl:with-param name ="lineNo"  select ="5"/>
                </xsl:call-template>
                <!--End-->
                <!--Space Before and After Paragraph-->
                <xsl:call-template name ="tmpMarTop"/>
                <xsl:call-template name ="tmpMarBottom"/>
                <!--End-->
                <xsl:if test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                  <xsl:call-template name="SlideMasterBulletsNumbers">
                    <xsl:with-param name="masterName" select="$slideMasterName"/>
                    <xsl:with-param name="level" select="5"/>
                  </xsl:call-template>
                </xsl:if>
                <a:defRPr>
                  <xsl:if test="./style:text-properties/@fo:font-size">
<!--Office 2007 Sp2-->

                    <xsl:variable name="fontSize">
                      <xsl:call-template name="point-measure">
                        <xsl:with-param name="length" select="style:text-properties/@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="sz">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="concat($fontSize,'pt')"/>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:if>
                  <!--Font Bold attribute -->
                  <xsl:variable name ="boldStyle">
                    <xsl:call-template name ="BoldItalic">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="5"/>
                      <xsl:with-param name ="paramName" select ="'b'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$boldStyle ='bold'">
                    <xsl:attribute name ="b">
                      <xsl:value-of  select ="1"/>
                    </xsl:attribute >
                  </xsl:if>
                  <!-- Font Italic attribute-->
                  <xsl:variable name ="italicStyle">
                    <xsl:call-template name ="BoldItalic">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="5"/>
                      <xsl:with-param name ="paramName" select ="'i'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$italicStyle ='italic'">
                    <xsl:attribute name ="i">
                      <xsl:value-of  select ="1"/>
                    </xsl:attribute >
                  </xsl:if>
                  <!-- Font Underline-->
                  <xsl:choose>
                    <xsl:when test="./style:text-properties/@style:text-underline-style !='none'">
                      <xsl:variable name ="unLine">
                        <xsl:call-template name ="ULine">
                          <xsl:with-param name="smName" select="$slideMasterName"/>
                          <xsl:with-param name ="outlineNo"  select ="5"/>
                          <xsl:with-param name ="paramName" select ="'u'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$unLine !=''">
                        <xsl:attribute name="u">
                          <xsl:value-of  select ="$unLine"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                    <xsl:when test="./style:text-properties/@style:text-underline-style ='none'">
                      <xsl:attribute name="u">
                        <xsl:value-of select="'none'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test="not(./style:text-properties/style:text-underline-style)">
                      <xsl:variable name ="unLine">
                        <xsl:call-template name ="ULine">
                          <xsl:with-param name="smName" select="$slideMasterName"/>
                          <xsl:with-param name ="outlineNo"  select ="5"/>
                          <xsl:with-param name ="paramName" select ="'u'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$unLine !=''">
                        <xsl:attribute name="u">
                          <xsl:value-of  select ="$unLine"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                  </xsl:choose>
                  <!--end-->
                  <!-- Kerning -->
                  <xsl:if test ="./style:text-properties/@style:letter-kerning = 'true'">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="1200"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="./style:text-properties/@style:letter-kerning = 'false'">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="0"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="not(./style:text-properties/@style:letter-kerning)">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="0"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!-- End -->
                  <!-- Superscript and SubScript for Text added by Mathi on 1st Aug 2007-->
                  <xsl:call-template name="SuperAndSubscripts" />

                  <!--Character Spacing-->
                  <xsl:variable name ="space">
                    <xsl:call-template name ="Spacing">
                      <xsl:with-param name="smasterName" select="$slideMasterName"/>
                      <xsl:with-param name ="lineNo"  select ="5"/>
                      <xsl:with-param name ="parName" select ="'spc'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$space !='' ">
                    <xsl:attribute name="spc">
                      <xsl:value-of  select ="$space"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!--End-->
                  <!-- Font Strike through Start-->
                  <xsl:choose >
                    <xsl:when  test="style:text-properties/@style:text-line-through-type = 'solid'">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'sngStrike'"/>
                      </xsl:attribute >
                    </xsl:when >
                    <xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'dblStrike'"/>
                      </xsl:attribute >
                    </xsl:when >
                    <!-- style:text-line-through-style-->
                    <xsl:when test="style:text-properties/@style:text-line-through-style = 'solid'">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'sngStrike'"/>
                      </xsl:attribute >
                    </xsl:when>
                  </xsl:choose>
                  <!--Font Color-->
                  <a:solidFill>
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:variable name ="color">
                          <xsl:call-template name ="FontColor">
                            <xsl:with-param name="smasterName" select="$slideMasterName"/>
                            <xsl:with-param name ="lineNo"  select ="5"/>
                            <xsl:with-param name ="parName" select ="'val'"/>
                          </xsl:call-template>
                        </xsl:variable>
                        <xsl:if test ="$color">
                          <xsl:value-of  select ="$color"/>
                        </xsl:if>
                      </xsl:attribute>
                    </a:srgbClr>
                  </a:solidFill>
                  <!--end-->
                  <!--Underline Color-->
                  <xsl:if test ="style:text-properties/@style:text-underline-color !='font-color'">
                    <a:uFill>
                      <a:solidFill>
                        <a:srgbClr>
                          <xsl:attribute name ="val">
                            <xsl:value-of select ="substring-after(style:text-properties/@style:text-underline-color,'#')"/>
                          </xsl:attribute>
                        </a:srgbClr>
                      </a:solidFill>
                    </a:uFill>
                  </xsl:if>
                  <!--end-->
                  <!--Shadow Color-->
                  <xsl:if test ="style:text-properties/@fo:text-shadow != 'none'">
                    <a:effectLst>
                      <a:outerShdw blurRad="38100" dist="38100" dir="2700000" >
                        <a:srgbClr val="000000">
                          <a:alpha val="43137" />
                        </a:srgbClr>
                      </a:outerShdw>
                    </a:effectLst>
                  </xsl:if>
                  <!--End-->
                  <!--Font Family-->
                  <a:latin>
                    <xsl:attribute name="typeface">
                      <xsl:variable name ="foFamily">
                        <xsl:call-template name ="FontFamily">
                          <xsl:with-param name="smasterName" select="$slideMasterName"/>
                          <xsl:with-param name ="lineNo"  select ="5"/>
                          <xsl:with-param name ="parName" select ="'typeface'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$foFamily">
                        <xsl:value-of  select ="$foFamily"/>
                      </xsl:if>
                    </xsl:attribute>
                  </a:latin>
                  <a:ea typeface="+mn-ea"/>
                  <a:cs typeface="+mn-cs"/>
                </a:defRPr>
              </xsl:for-each>
            </a:lvl5pPr>
            <a:lvl6pPr>
              <xsl:variable name="outlineName">
                <xsl:value-of select="concat($slideMasterName,'-outline1')"/>
              </xsl:variable>
              <xsl:variable name="outlineName6">
                <xsl:value-of select="concat($slideMasterName,'-outline6')"/>
              </xsl:variable>
              <xsl:for-each select="/office:document-styles/office:styles/style:style[@style:name=$outlineName6]">
                <!--Margin-->
                <xsl:choose>
                  <xsl:when test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                    <xsl:variable name="margin">
                      <xsl:call-template name="MarginTemplate">
                        <xsl:with-param name="masterName" select="$slideMasterName"/>
                        <xsl:with-param name="level" select="6"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="marL">
                      <xsl:value-of select="$margin"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="marL">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'cm'"/>
                        <xsl:with-param name ="length">
                          <xsl:call-template name="convertUnitsToCm">
                            <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
                          </xsl:call-template>
                        </xsl:with-param>
                      </xsl:call-template>

                    </xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <!--End-->
                <!--Indent-->
                <xsl:choose>
                  <xsl:when test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                    <xsl:variable name="indetValue">
                      <xsl:call-template name="IndentTemplate">
                        <xsl:with-param name="masterName" select="$slideMasterName"/>
                        <xsl:with-param name="level" select="6"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="indent">
                      <xsl:value-of select="$indetValue"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="indent">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'cm'"/>
                        <xsl:with-param name ="length">
                          <xsl:call-template name="convertUnitsToCm">
                            <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:text-indent"/>
                          </xsl:call-template>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <!--End-->
                <!--fo:text-align-->
                <xsl:attribute name ="algn">
                  <xsl:variable name ="alignment">
                    <xsl:call-template name ="Alignment">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="6"/>
                      <xsl:with-param name ="paramName" select ="'algn'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$alignment">
                    <xsl:value-of select="$alignment"/>
                  </xsl:if>
                </xsl:attribute>
                <!--End-->
                <xsl:attribute name="defTabSz">
                  <xsl:value-of select="'914400'"/>
                </xsl:attribute>
                <xsl:attribute name="rtl">
                  <xsl:value-of select="'0'"/>
                </xsl:attribute>
                <xsl:attribute name="eaLnBrk">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
                <xsl:attribute name="latinLnBrk">
                  <xsl:value-of select="'0'"/>
                </xsl:attribute>
                <xsl:attribute name="hangingPunct">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
                <!--Line Spacing-->
                <xsl:call-template name ="LineSpacingPoints">
                  <xsl:with-param name="smasterName" select="$slideMasterName"/>
                  <xsl:with-param name ="lineNo"  select ="6"/>
                </xsl:call-template>
                <!--End-->
                <!--Space Before and After Paragraph-->
                <xsl:call-template name ="tmpMarTop"/>
                <xsl:call-template name ="tmpMarBottom"/>
                <!--End-->
                <xsl:if test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                  <xsl:call-template name="SlideMasterBulletsNumbers">
                    <xsl:with-param name="masterName" select="$slideMasterName"/>
                    <xsl:with-param name="level" select="6"/>
                  </xsl:call-template>
                </xsl:if>
                <a:defRPr>
                  <xsl:if test="./style:text-properties/@fo:font-size">
<!--Office 2007 Sp2-->

                    <xsl:variable name="fontSize">
                      <xsl:call-template name="point-measure">
                        <xsl:with-param name="length" select="style:text-properties/@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="sz">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="concat($fontSize,'pt')"/>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:if>
                  <!--Font Bold attribute -->
                  <xsl:variable name ="boldStyle">
                    <xsl:call-template name ="BoldItalic">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="6"/>
                      <xsl:with-param name ="paramName" select ="'b'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$boldStyle ='bold'">
                    <xsl:attribute name ="b">
                      <xsl:value-of  select ="1"/>
                    </xsl:attribute >
                  </xsl:if>
                  <!-- Font Italic attribute-->
                  <xsl:variable name ="italicStyle">
                    <xsl:call-template name ="BoldItalic">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="6"/>
                      <xsl:with-param name ="paramName" select ="'i'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$italicStyle ='italic'">
                    <xsl:attribute name ="i">
                      <xsl:value-of  select ="1"/>
                    </xsl:attribute >
                  </xsl:if>
                  <!-- Font Underline-->
                  <xsl:choose>
                    <xsl:when test="./style:text-properties/@style:text-underline-style !='none'">
                      <xsl:variable name ="unLine">
                        <xsl:call-template name ="ULine">
                          <xsl:with-param name="smName" select="$slideMasterName"/>
                          <xsl:with-param name ="outlineNo"  select ="6"/>
                          <xsl:with-param name ="paramName" select ="'u'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$unLine !=''">
                        <xsl:attribute name="u">
                          <xsl:value-of  select ="$unLine"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                    <xsl:when test="./style:text-properties/@style:text-underline-style ='none'">
                      <xsl:attribute name="u">
                        <xsl:value-of select="'none'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test="not(./style:text-properties/style:text-underline-style)">
                      <xsl:variable name ="unLine">
                        <xsl:call-template name ="ULine">
                          <xsl:with-param name="smName" select="$slideMasterName"/>
                          <xsl:with-param name ="outlineNo"  select ="6"/>
                          <xsl:with-param name ="paramName" select ="'u'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$unLine !=''">
                        <xsl:attribute name="u">
                          <xsl:value-of  select ="$unLine"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                  </xsl:choose>
                  <!--end-->
                  <!-- Kerning -->
                  <xsl:if test ="./style:text-properties/@style:letter-kerning = 'true'">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="1200"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="./style:text-properties/@style:letter-kerning = 'false'">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="0"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="not(./style:text-properties/@style:letter-kerning)">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="0"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!-- End -->
                  <!-- Superscript and SubScript for Text added by Mathi on 1st Aug 2007-->
                  <xsl:call-template name="SuperAndSubscripts" />

                  <!--Character Spacing-->
                  <xsl:variable name ="space">
                    <xsl:call-template name ="Spacing">
                      <xsl:with-param name="smasterName" select="$slideMasterName"/>
                      <xsl:with-param name ="lineNo"  select ="6"/>
                      <xsl:with-param name ="parName" select ="'spc'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$space !='' ">
                    <xsl:attribute name="spc">
                      <xsl:value-of  select ="$space"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!--End-->
                  <!-- Font Strike through Start-->
                  <xsl:choose >
                    <xsl:when  test="style:text-properties/@style:text-line-through-type = 'solid'">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'sngStrike'"/>
                      </xsl:attribute >
                    </xsl:when >
                    <xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'dblStrike'"/>
                      </xsl:attribute >
                    </xsl:when >
                    <!-- style:text-line-through-style-->
                    <xsl:when test="style:text-properties/@style:text-line-through-style = 'solid'">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'sngStrike'"/>
                      </xsl:attribute >
                    </xsl:when>
                  </xsl:choose>
                  <!--Font Color-->
                  <a:solidFill>
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:variable name ="color">
                          <xsl:call-template name ="FontColor">
                            <xsl:with-param name="smasterName" select="$slideMasterName"/>
                            <xsl:with-param name ="lineNo"  select ="6"/>
                            <xsl:with-param name ="parName" select ="'val'"/>
                          </xsl:call-template>
                        </xsl:variable>
                        <xsl:if test ="$color">
                          <xsl:value-of  select ="$color"/>
                        </xsl:if>
                      </xsl:attribute>
                    </a:srgbClr>
                  </a:solidFill>
                  <!--end-->
                  <!--Underline Color-->
                  <xsl:if test ="style:text-properties/@style:text-underline-color !='font-color'">
                    <a:uFill>
                      <a:solidFill>
                        <a:srgbClr>
                          <xsl:attribute name ="val">
                            <xsl:value-of select ="substring-after(style:text-properties/@style:text-underline-color,'#')"/>
                          </xsl:attribute>
                        </a:srgbClr>
                      </a:solidFill>
                    </a:uFill>
                  </xsl:if>
                  <!--end-->
                  <!--Shadow Color-->
                  <xsl:if test ="style:text-properties/@fo:text-shadow != 'none'">
                    <a:effectLst>
                      <a:outerShdw blurRad="38100" dist="38100" dir="2700000" >
                        <a:srgbClr val="000000">
                          <a:alpha val="43137" />
                        </a:srgbClr>
                      </a:outerShdw>
                    </a:effectLst>
                  </xsl:if>
                  <!--End-->
                  <!--Font Family-->
                  <a:latin>
                    <xsl:attribute name="typeface">
                      <xsl:variable name ="foFamily">
                        <xsl:call-template name ="FontFamily">
                          <xsl:with-param name="smasterName" select="$slideMasterName"/>
                          <xsl:with-param name ="lineNo"  select ="6"/>
                          <xsl:with-param name ="parName" select ="'typeface'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$foFamily">
                        <xsl:value-of  select ="$foFamily"/>
                      </xsl:if>
                    </xsl:attribute>
                  </a:latin>
                  <a:ea typeface="+mn-ea"/>
                  <a:cs typeface="+mn-cs"/>
                </a:defRPr>
              </xsl:for-each>
            </a:lvl6pPr>
            <a:lvl7pPr>
              <xsl:variable name="outlineName">
                <xsl:value-of select="concat($slideMasterName,'-outline1')"/>
              </xsl:variable>
              <xsl:variable name="outlineName7">
                <xsl:value-of select="concat($slideMasterName,'-outline7')"/>
              </xsl:variable>
              <xsl:for-each select="/office:document-styles/office:styles/style:style[@style:name=$outlineName7]">
                <!--Margin-->
                <xsl:choose>
                  <xsl:when test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                    <xsl:variable name="margin">
                      <xsl:call-template name="MarginTemplate">
                        <xsl:with-param name="masterName" select="$slideMasterName"/>
                        <xsl:with-param name="level" select="7"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="marL">
                      <xsl:value-of select="$margin"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="marL">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'cm'"/>
                        <xsl:with-param name ="length">
                          <xsl:call-template name="convertUnitsToCm">
                            <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
                          </xsl:call-template>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <!--End-->
                <!--Indent-->
                <xsl:choose>
                  <xsl:when test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                    <xsl:variable name="indetValue">
                      <xsl:call-template name="IndentTemplate">
                        <xsl:with-param name="masterName" select="$slideMasterName"/>
                        <xsl:with-param name="level" select="7"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="indent">
                      <xsl:value-of select="$indetValue"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="indent">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'cm'"/>
                        <xsl:with-param name ="length">
                          <xsl:call-template name="convertUnitsToCm">
                            <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:text-indent"/>
                          </xsl:call-template>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <!--End-->
                <!--fo:text-align-->
                <xsl:attribute name ="algn">
                  <xsl:variable name ="alignment">
                    <xsl:call-template name ="Alignment">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="7"/>
                      <xsl:with-param name ="paramName" select ="'algn'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$alignment">
                    <xsl:value-of select="$alignment"/>
                  </xsl:if>
                </xsl:attribute>
                <!--End-->
                <xsl:attribute name="defTabSz">
                  <xsl:value-of select="'914400'"/>
                </xsl:attribute>
                <xsl:attribute name="rtl">
                  <xsl:value-of select="'0'"/>
                </xsl:attribute>
                <xsl:attribute name="eaLnBrk">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
                <xsl:attribute name="latinLnBrk">
                  <xsl:value-of select="'0'"/>
                </xsl:attribute>
                <xsl:attribute name="hangingPunct">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
                <!--Line Spacing-->
                <xsl:call-template name ="LineSpacingPoints">
                  <xsl:with-param name="smasterName" select="$slideMasterName"/>
                  <xsl:with-param name ="lineNo"  select ="7"/>
                </xsl:call-template>
                <!--End-->
                <!--Space Before and After Paragraph-->
                <xsl:call-template name ="tmpMarTop"/>
                <xsl:call-template name ="tmpMarBottom"/>
                <!--End-->
                <xsl:if test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                  <xsl:call-template name="SlideMasterBulletsNumbers">
                    <xsl:with-param name="masterName" select="$slideMasterName"/>
                    <xsl:with-param name="level" select="7"/>
                  </xsl:call-template>
                </xsl:if>
                <a:defRPr>
                  <xsl:if test="./style:text-properties/@fo:font-size">
<!--Office 2007 Sp2-->

                    <xsl:variable name="fontSize">
                      <xsl:call-template name="point-measure">
                        <xsl:with-param name="length" select="style:text-properties/@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="sz">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="concat($fontSize,'pt')"/>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:if>
                  <!--Font Bold attribute -->
                  <xsl:variable name ="boldStyle">
                    <xsl:call-template name ="BoldItalic">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="7"/>
                      <xsl:with-param name ="paramName" select ="'b'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$boldStyle ='bold'">
                    <xsl:attribute name ="b">
                      <xsl:value-of  select ="1"/>
                    </xsl:attribute >
                  </xsl:if>
                  <!-- Font Italic attribute-->
                  <xsl:variable name ="italicStyle">
                    <xsl:call-template name ="BoldItalic">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="7"/>
                      <xsl:with-param name ="paramName" select ="'i'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$italicStyle ='italic'">
                    <xsl:attribute name ="i">
                      <xsl:value-of  select ="1"/>
                    </xsl:attribute >
                  </xsl:if>
                  <!-- Font Underline-->
                  <xsl:choose>
                    <xsl:when test="./style:text-properties/@style:text-underline-style !='none'">
                      <xsl:variable name ="unLine">
                        <xsl:call-template name ="ULine">
                          <xsl:with-param name="smName" select="$slideMasterName"/>
                          <xsl:with-param name ="outlineNo"  select ="7"/>
                          <xsl:with-param name ="paramName" select ="'u'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$unLine !=''">
                        <xsl:attribute name="u">
                          <xsl:value-of  select ="$unLine"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                    <xsl:when test="./style:text-properties/@style:text-underline-style ='none'">
                      <xsl:attribute name="u">
                        <xsl:value-of select="'none'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test="not(./style:text-properties/style:text-underline-style)">
                      <xsl:variable name ="unLine">
                        <xsl:call-template name ="ULine">
                          <xsl:with-param name="smName" select="$slideMasterName"/>
                          <xsl:with-param name ="outlineNo"  select ="7"/>
                          <xsl:with-param name ="paramName" select ="'u'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$unLine !=''">
                        <xsl:attribute name="u">
                          <xsl:value-of  select ="$unLine"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                  </xsl:choose>
                  <!--end-->
                  <!-- Kerning -->
                  <xsl:if test ="./style:text-properties/@style:letter-kerning = 'true'">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="1200"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="./style:text-properties/@style:letter-kerning = 'false'">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="0"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="not(./style:text-properties/@style:letter-kerning)">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="0"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!-- End -->
                  <!-- Superscript and SubScript for Text added by Mathi on 1st Aug 2007-->
                  <xsl:call-template name="SuperAndSubscripts" />

                  <!--Character Spacing-->
                  <xsl:variable name ="space">
                    <xsl:call-template name ="Spacing">
                      <xsl:with-param name="smasterName" select="$slideMasterName"/>
                      <xsl:with-param name ="lineNo"  select ="7"/>
                      <xsl:with-param name ="parName" select ="'spc'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$space !='' ">
                    <xsl:attribute name="spc">
                      <xsl:value-of  select ="$space"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!--End-->
                  <!-- Font Strike through Start-->
                  <xsl:choose >
                    <xsl:when  test="style:text-properties/@style:text-line-through-type = 'solid'">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'sngStrike'"/>
                      </xsl:attribute >
                    </xsl:when >
                    <xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'dblStrike'"/>
                      </xsl:attribute >
                    </xsl:when >
                    <!-- style:text-line-through-style-->
                    <xsl:when test="style:text-properties/@style:text-line-through-style = 'solid'">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'sngStrike'"/>
                      </xsl:attribute >
                    </xsl:when>
                  </xsl:choose>
                  <!--Font Color-->
                  <a:solidFill>
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:variable name ="color">
                          <xsl:call-template name ="FontColor">
                            <xsl:with-param name="smasterName" select="$slideMasterName"/>
                            <xsl:with-param name ="lineNo"  select ="7"/>
                            <xsl:with-param name ="parName" select ="'val'"/>
                          </xsl:call-template>
                        </xsl:variable>
                        <xsl:if test ="$color">
                          <xsl:value-of  select ="$color"/>
                        </xsl:if>
                      </xsl:attribute>
                    </a:srgbClr>
                  </a:solidFill>
                  <!--end-->
                  <!--Underline Color-->
                  <xsl:if test ="style:text-properties/@style:text-underline-color !='font-color'">
                    <a:uFill>
                      <a:solidFill>
                        <a:srgbClr>
                          <xsl:attribute name ="val">
                            <xsl:value-of select ="substring-after(style:text-properties/@style:text-underline-color,'#')"/>
                          </xsl:attribute>
                        </a:srgbClr>
                      </a:solidFill>
                    </a:uFill>
                  </xsl:if>
                  <!--end-->
                  <!--Shadow Color-->
                  <xsl:if test ="style:text-properties/@fo:text-shadow != 'none'">
                    <a:effectLst>
                      <a:outerShdw blurRad="38100" dist="38100" dir="2700000" >
                        <a:srgbClr val="000000">
                          <a:alpha val="43137" />
                        </a:srgbClr>
                      </a:outerShdw>
                    </a:effectLst>
                  </xsl:if>
                  <!--End-->
                  <!--Font Family-->
                  <a:latin>
                    <xsl:attribute name="typeface">
                      <xsl:variable name ="foFamily">
                        <xsl:call-template name ="FontFamily">
                          <xsl:with-param name="smasterName" select="$slideMasterName"/>
                          <xsl:with-param name ="lineNo"  select ="7"/>
                          <xsl:with-param name ="parName" select ="'typeface'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$foFamily">
                        <xsl:value-of  select ="$foFamily"/>
                      </xsl:if>
                    </xsl:attribute>
                  </a:latin>
                  <a:ea typeface="+mn-ea"/>
                  <a:cs typeface="+mn-cs"/>
                </a:defRPr>
              </xsl:for-each>
            </a:lvl7pPr>
            <a:lvl8pPr>
              <xsl:variable name="outlineName">
                <xsl:value-of select="concat($slideMasterName,'-outline1')"/>
              </xsl:variable>
              <xsl:variable name="outlineName8">
                <xsl:value-of select="concat($slideMasterName,'-outline8')"/>
              </xsl:variable>
              <xsl:for-each select="/office:document-styles/office:styles/style:style[@style:name=$outlineName8]">
                <!--Margin-->
                <xsl:choose>
                  <xsl:when test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                    <xsl:variable name="margin">
                      <xsl:call-template name="MarginTemplate">
                        <xsl:with-param name="masterName" select="$slideMasterName"/>
                        <xsl:with-param name="level" select="8"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="marL">
                      <xsl:value-of select="$margin"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="marL">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'cm'"/>
                        <xsl:with-param name ="length">
                          <xsl:call-template name="convertUnitsToCm">
                            <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
                          </xsl:call-template>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <!--End-->
                <!--Indent-->
                <xsl:choose>
                  <xsl:when test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                    <xsl:variable name="indetValue">
                      <xsl:call-template name="IndentTemplate">
                        <xsl:with-param name="masterName" select="$slideMasterName"/>
                        <xsl:with-param name="level" select="8"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="indent">
                      <xsl:value-of select="$indetValue"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="indent">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'cm'"/>
                        <xsl:with-param name ="length">
                          <xsl:call-template name="convertUnitsToCm">
                            <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:text-indent"/>
                          </xsl:call-template>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <!--End-->
                <!--fo:text-align-->
                <xsl:attribute name ="algn">
                  <xsl:variable name ="alignment">
                    <xsl:call-template name ="Alignment">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="8"/>
                      <xsl:with-param name ="paramName" select ="'algn'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$alignment">
                    <xsl:value-of select="$alignment"/>
                  </xsl:if>
                </xsl:attribute>
                <!--End-->
                <xsl:attribute name="defTabSz">
                  <xsl:value-of select="'914400'"/>
                </xsl:attribute>
                <xsl:attribute name="rtl">
                  <xsl:value-of select="'0'"/>
                </xsl:attribute>
                <xsl:attribute name="eaLnBrk">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
                <xsl:attribute name="latinLnBrk">
                  <xsl:value-of select="'0'"/>
                </xsl:attribute>
                <xsl:attribute name="hangingPunct">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
                <!--Line Spacing-->
                <xsl:call-template name ="LineSpacingPoints">
                  <xsl:with-param name="smasterName" select="$slideMasterName"/>
                  <xsl:with-param name ="lineNo"  select ="8"/>
                </xsl:call-template>
                <!--End-->
                <!--Space Before and After Paragraph-->
                <xsl:call-template name ="tmpMarTop"/>
                <xsl:call-template name ="tmpMarBottom"/>
                <!--End-->
                <xsl:if test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                  <xsl:call-template name="SlideMasterBulletsNumbers">
                    <xsl:with-param name="masterName" select="$slideMasterName"/>
                    <xsl:with-param name="level" select="8"/>
                  </xsl:call-template>
                </xsl:if>
                <a:defRPr>
                  <xsl:if test="./style:text-properties/@fo:font-size">
<!--Office 2007 Sp2-->

                    <xsl:variable name="fontSize">
                      <xsl:call-template name="point-measure">
                        <xsl:with-param name="length" select="style:text-properties/@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="sz">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="concat($fontSize,'pt')"/>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:if>
                  <!--Font Bold attribute -->
                  <xsl:variable name ="boldStyle">
                    <xsl:call-template name ="BoldItalic">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="8"/>
                      <xsl:with-param name ="paramName" select ="'b'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$boldStyle ='bold'">
                    <xsl:attribute name ="b">
                      <xsl:value-of  select ="1"/>
                    </xsl:attribute >
                  </xsl:if>
                  <!-- Font Italic attribute-->
                  <xsl:variable name ="italicStyle">
                    <xsl:call-template name ="BoldItalic">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="8"/>
                      <xsl:with-param name ="paramName" select ="'i'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$italicStyle ='italic'">
                    <xsl:attribute name ="i">
                      <xsl:value-of  select ="1"/>
                    </xsl:attribute >
                  </xsl:if>
                  <!-- Font Underline-->
                  <xsl:choose>
                    <xsl:when test="./style:text-properties/@style:text-underline-style !='none'">
                      <xsl:variable name ="unLine">
                        <xsl:call-template name ="ULine">
                          <xsl:with-param name="smName" select="$slideMasterName"/>
                          <xsl:with-param name ="outlineNo"  select ="8"/>
                          <xsl:with-param name ="paramName" select ="'u'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$unLine !=''">
                        <xsl:attribute name="u">
                          <xsl:value-of  select ="$unLine"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                    <xsl:when test="./style:text-properties/@style:text-underline-style ='none'">
                      <xsl:attribute name="u">
                        <xsl:value-of select="'none'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test="not(./style:text-properties/style:text-underline-style)">
                      <xsl:variable name ="unLine">
                        <xsl:call-template name ="ULine">
                          <xsl:with-param name="smName" select="$slideMasterName"/>
                          <xsl:with-param name ="outlineNo"  select ="8"/>
                          <xsl:with-param name ="paramName" select ="'u'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$unLine !=''">
                        <xsl:attribute name="u">
                          <xsl:value-of  select ="$unLine"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                  </xsl:choose>
                  <!--end-->
                  <!-- Kerning -->
                  <xsl:if test ="./style:text-properties/@style:letter-kerning = 'true'">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="1200"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="./style:text-properties/@style:letter-kerning = 'false'">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="0"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="not(./style:text-properties/@style:letter-kerning)">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="0"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!-- End -->
                  <!-- Superscript and SubScript for Text added by Mathi on 1st Aug 2007-->
                  <xsl:call-template name="SuperAndSubscripts" />

                  <!--Character Spacing-->
                  <xsl:variable name ="space">
                    <xsl:call-template name ="Spacing">
                      <xsl:with-param name="smasterName" select="$slideMasterName"/>
                      <xsl:with-param name ="lineNo"  select ="8"/>
                      <xsl:with-param name ="parName" select ="'spc'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$space !='' ">
                    <xsl:attribute name="spc">
                      <xsl:value-of  select ="$space"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!--End-->
                  <!-- Font Strike through Start-->
                  <xsl:choose >
                    <xsl:when  test="style:text-properties/@style:text-line-through-type = 'solid'">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'sngStrike'"/>
                      </xsl:attribute >
                    </xsl:when >
                    <xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'dblStrike'"/>
                      </xsl:attribute >
                    </xsl:when >
                    <!-- style:text-line-through-style-->
                    <xsl:when test="style:text-properties/@style:text-line-through-style = 'solid'">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'sngStrike'"/>
                      </xsl:attribute >
                    </xsl:when>
                  </xsl:choose>
                  <!--Font Color-->
                  <a:solidFill>
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:variable name ="color">
                          <xsl:call-template name ="FontColor">
                            <xsl:with-param name="smasterName" select="$slideMasterName"/>
                            <xsl:with-param name ="lineNo"  select ="8"/>
                            <xsl:with-param name ="parName" select ="'val'"/>
                          </xsl:call-template>
                        </xsl:variable>
                        <xsl:if test ="$color">
                          <xsl:value-of  select ="$color"/>
                        </xsl:if>
                      </xsl:attribute>
                    </a:srgbClr>
                  </a:solidFill>
                  <!--end-->
                  <!--Underline Color-->
                  <xsl:if test ="style:text-properties/@style:text-underline-color !='font-color'">
                    <a:uFill>
                      <a:solidFill>
                        <a:srgbClr>
                          <xsl:attribute name ="val">
                            <xsl:value-of select ="substring-after(style:text-properties/@style:text-underline-color,'#')"/>
                          </xsl:attribute>
                        </a:srgbClr>
                      </a:solidFill>
                    </a:uFill>
                  </xsl:if>
                  <!--end-->
                  <!--Shadow Color-->
                  <xsl:if test ="style:text-properties/@fo:text-shadow != 'none'">
                    <a:effectLst>
                      <a:outerShdw blurRad="38100" dist="38100" dir="2700000" >
                        <a:srgbClr val="000000">
                          <a:alpha val="43137" />
                        </a:srgbClr>
                      </a:outerShdw>
                    </a:effectLst>
                  </xsl:if>
                  <!--End-->
                  <!--Font Family-->
                  <a:latin>
                    <xsl:attribute name="typeface">
                      <xsl:variable name ="foFamily">
                        <xsl:call-template name ="FontFamily">
                          <xsl:with-param name="smasterName" select="$slideMasterName"/>
                          <xsl:with-param name ="lineNo"  select ="8"/>
                          <xsl:with-param name ="parName" select ="'typeface'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$foFamily">
                        <xsl:value-of  select ="$foFamily"/>
                      </xsl:if>
                    </xsl:attribute>
                  </a:latin>
                  <a:ea typeface="+mn-ea"/>
                  <a:cs typeface="+mn-cs"/>
                </a:defRPr>
              </xsl:for-each>
            </a:lvl8pPr>
            <a:lvl9pPr>
              <xsl:variable name="outlineName">
                <xsl:value-of select="concat($slideMasterName,'-outline1')"/>
              </xsl:variable>
              <xsl:variable name="outlineName9">
                <xsl:value-of select="concat($slideMasterName,'-outline9')"/>
              </xsl:variable>
              <xsl:for-each select="/office:document-styles/office:styles/style:style[@style:name=$outlineName9]">
                <!--Margin-->
                <xsl:choose>
                  <xsl:when test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                    <xsl:variable name="margin">
                      <xsl:call-template name="MarginTemplate">
                        <xsl:with-param name="masterName" select="$slideMasterName"/>
                        <xsl:with-param name="level" select="9"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="marL">
                      <xsl:value-of select="$margin"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="marL">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'cm'"/>
                        <xsl:with-param name ="length">
                          <xsl:call-template name="convertUnitsToCm">
                            <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
                          </xsl:call-template>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <!--End-->
                <!--Indent-->
                <xsl:choose>
                  <xsl:when test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                    <xsl:variable name="indetValue">
                      <xsl:call-template name="IndentTemplate">
                        <xsl:with-param name="masterName" select="$slideMasterName"/>
                        <xsl:with-param name="level" select="9"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="indent">
                      <xsl:value-of select="$indetValue"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="indent">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'cm'"/>
                        <xsl:with-param name ="length">
                          <xsl:call-template name="convertUnitsToCm">
                            <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:text-indent"/>
                          </xsl:call-template>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <!--End-->
                <!--fo:text-align-->
                <xsl:attribute name ="algn">
                  <xsl:variable name ="alignment">
                    <xsl:call-template name ="Alignment">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="9"/>
                      <xsl:with-param name ="paramName" select ="'algn'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$alignment">
                    <xsl:value-of select="$alignment"/>
                  </xsl:if>
                </xsl:attribute>
                <!--End-->
                <xsl:attribute name="defTabSz">
                  <xsl:value-of select="'914400'"/>
                </xsl:attribute>
                <xsl:attribute name="rtl">
                  <xsl:value-of select="'0'"/>
                </xsl:attribute>
                <xsl:attribute name="eaLnBrk">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
                <xsl:attribute name="latinLnBrk">
                  <xsl:value-of select="'0'"/>
                </xsl:attribute>
                <xsl:attribute name="hangingPunct">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
                <!--Line Spacing-->
                <xsl:call-template name ="LineSpacingPoints">
                  <xsl:with-param name="smasterName" select="$slideMasterName"/>
                  <xsl:with-param name ="lineNo"  select ="9"/>
                </xsl:call-template>
                <!--End-->
                <!--Space Before and After Paragraph-->
                <xsl:call-template name ="tmpMarTop"/>
                <xsl:call-template name ="tmpMarBottom"/>
                <!--End-->
                <xsl:if test="./style:paragraph-properties/@text:enable-numbering='true' or document('styles.xml')//style:style[@style:name=$outlineName]/style:paragraph-properties/@text:enable-numbering='true'">
                  <xsl:call-template name="SlideMasterBulletsNumbers">
                    <xsl:with-param name="masterName" select="$slideMasterName"/>
                    <xsl:with-param name="level" select="9"/>
                  </xsl:call-template>
                </xsl:if>
                <a:defRPr>
                  <xsl:if test="./style:text-properties/@fo:font-size">
<!--Office 2007 Sp2-->

                    <xsl:variable name="fontSize">
                      <xsl:call-template name="point-measure">
                        <xsl:with-param name="length" select="style:text-properties/@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:variable>

                    <xsl:attribute name="sz">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="concat($fontSize,'pt')"/>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:if>
                  <!--Font Bold attribute boldStyleitalicStyleunLine-->
                  <xsl:variable name ="boldStyle">
                    <xsl:call-template name ="BoldItalic">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="9"/>
                      <xsl:with-param name ="paramName" select ="'b'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$boldStyle ='bold'">
                    <xsl:attribute name ="b">
                      <xsl:value-of  select ="1"/>
                    </xsl:attribute >
                  </xsl:if>
                  <!-- Font Italic attribute-->
                  <xsl:variable name ="italicStyle">
                    <xsl:call-template name ="BoldItalic">
                      <xsl:with-param name="smName" select="$slideMasterName"/>
                      <xsl:with-param name ="outlineNo"  select ="9"/>
                      <xsl:with-param name ="paramName" select ="'i'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$italicStyle ='italic'">
                    <xsl:attribute name ="i">
                      <xsl:value-of  select ="1"/>
                    </xsl:attribute >
                  </xsl:if>
                  <!-- Font Underline-->
                  <xsl:choose>
                    <xsl:when test="./style:text-properties/@style:text-underline-style !='none'">
                      <xsl:variable name ="unLine">
                        <xsl:call-template name ="ULine">
                          <xsl:with-param name="smName" select="$slideMasterName"/>
                          <xsl:with-param name ="outlineNo"  select ="9"/>
                          <xsl:with-param name ="paramName" select ="'u'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$unLine !=''">
                        <xsl:attribute name="u">
                          <xsl:value-of  select ="$unLine"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                    <xsl:when test="./style:text-properties/@style:text-underline-style ='none'">
                      <xsl:attribute name="u">
                        <xsl:value-of select="'none'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test="not(./style:text-properties/style:text-underline-style)">
                      <xsl:variable name ="unLine">
                        <xsl:call-template name ="ULine">
                          <xsl:with-param name="smName" select="$slideMasterName"/>
                          <xsl:with-param name ="outlineNo"  select ="9"/>
                          <xsl:with-param name ="paramName" select ="'u'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$unLine !=''">
                        <xsl:attribute name="u">
                          <xsl:value-of  select ="$unLine"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                  </xsl:choose>
                  <!--end-->
                  <!-- Kerning -->
                  <xsl:if test ="./style:text-properties/@style:letter-kerning = 'true'">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="1200"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="./style:text-properties/@style:letter-kerning = 'false'">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="0"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="not(./style:text-properties/@style:letter-kerning)">
                    <xsl:attribute name ="kern">
                      <xsl:value-of select="0"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!-- End -->
                  <!-- Superscript and SubScript for Text added by Mathi on 1st Aug 2007-->
                  <xsl:call-template name="SuperAndSubscripts" />

                  <!--Character Spacing-->

                  <xsl:variable name ="space">
                    <xsl:call-template name ="Spacing">
                      <xsl:with-param name="smasterName" select="$slideMasterName"/>
                      <xsl:with-param name ="lineNo"  select ="9"/>
                      <xsl:with-param name ="parName" select ="'spc'"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:if test ="$space !='' ">
                    <xsl:attribute name="spc">
                      <xsl:value-of  select ="$space"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!--End-->
                  <!-- Font Strike through Start-->
                  <xsl:choose >
                    <xsl:when  test="style:text-properties/@style:text-line-through-type = 'solid'">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'sngStrike'"/>
                      </xsl:attribute >
                    </xsl:when >
                    <xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'dblStrike'"/>
                      </xsl:attribute >
                    </xsl:when >
                    <!-- style:text-line-through-style-->
                    <xsl:when test="style:text-properties/@style:text-line-through-style = 'solid'">
                      <xsl:attribute name ="strike">
                        <xsl:value-of select ="'sngStrike'"/>
                      </xsl:attribute >
                    </xsl:when>
                  </xsl:choose>
                  <!--Font Color-->
                  <a:solidFill>
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:variable name ="color">
                          <xsl:call-template name ="FontColor">
                            <xsl:with-param name="smasterName" select="$slideMasterName"/>
                            <xsl:with-param name ="lineNo"  select ="9"/>
                            <xsl:with-param name ="parName" select ="'val'"/>
                          </xsl:call-template>
                        </xsl:variable>
                        <xsl:if test ="$color">
                          <xsl:value-of  select ="$color"/>
                        </xsl:if>
                      </xsl:attribute>
                    </a:srgbClr>
                  </a:solidFill>
                  <!--end-->
                  <!--Underline Color-->
                  <xsl:if test ="style:text-properties/@style:text-underline-color !='font-color'">
                    <a:uFill>
                      <a:solidFill>
                        <a:srgbClr>
                          <xsl:attribute name ="val">
                            <xsl:value-of select ="substring-after(style:text-properties/@style:text-underline-color,'#')"/>
                          </xsl:attribute>
                        </a:srgbClr>
                      </a:solidFill>
                    </a:uFill>
                  </xsl:if>
                  <!--end-->
                  <!--Shadow Color-->
                  <xsl:if test ="style:text-properties/@fo:text-shadow != 'none'">
                    <a:effectLst>
                      <a:outerShdw blurRad="38100" dist="38100" dir="2700000" >
                        <a:srgbClr val="000000">
                          <a:alpha val="43137" />
                        </a:srgbClr>
                      </a:outerShdw>
                    </a:effectLst>
                  </xsl:if>
                  <!--End-->
                  <!--Font Family-->
                  <a:latin>
                    <xsl:attribute name="typeface">
                      <xsl:variable name ="foFamily">
                        <xsl:call-template name ="FontFamily">
                          <xsl:with-param name="smasterName" select="$slideMasterName"/>
                          <xsl:with-param name ="lineNo"  select ="9"/>
                          <xsl:with-param name ="parName" select ="'typeface'"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test ="$foFamily">
                        <xsl:value-of  select ="$foFamily"/>
                      </xsl:if>
                    </xsl:attribute>
                  </a:latin>
                  <a:ea typeface="+mn-ea"/>
                  <a:cs typeface="+mn-cs"/>
                </a:defRPr>
              </xsl:for-each>
            </a:lvl9pPr>
          </p:bodyStyle>
        </xsl:if>
        <xsl:if test="not(//office:styles/style:style[@style:name=$masterOutlineName])">
          <p:bodyStyle>
            <a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
              <a:spcBef>
                <a:spcPct val="20000"/>
              </a:spcBef>
              <a:buFont typeface="Arial" pitchFamily="34" charset="0"/>
              <a:buChar char="•"/>
              <a:defRPr sz="3200" kern="1200">
                <a:solidFill>
                  <a:schemeClr val="tx1"/>
                </a:solidFill>
                <a:latin typeface="+mn-lt"/>
                <a:ea typeface="+mn-ea"/>
                <a:cs typeface="+mn-cs"/>
              </a:defRPr>
            </a:lvl1pPr>
            <a:lvl2pPr marL="742950" indent="-285750" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
              <a:spcBef>
                <a:spcPct val="20000"/>
              </a:spcBef>
              <a:buFont typeface="Arial" pitchFamily="34" charset="0"/>
              <a:buChar char="–"/>
              <a:defRPr sz="2800" kern="1200">
                <a:solidFill>
                  <a:schemeClr val="tx1"/>
                </a:solidFill>
                <a:latin typeface="+mn-lt"/>
                <a:ea typeface="+mn-ea"/>
                <a:cs typeface="+mn-cs"/>
              </a:defRPr>
            </a:lvl2pPr>
            <a:lvl3pPr marL="1143000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
              <a:spcBef>
                <a:spcPct val="20000"/>
              </a:spcBef>
              <a:buFont typeface="Arial" pitchFamily="34" charset="0"/>
              <a:buChar char="•"/>
              <a:defRPr sz="2400" kern="1200">
                <a:solidFill>
                  <a:schemeClr val="tx1"/>
                </a:solidFill>
                <a:latin typeface="+mn-lt"/>
                <a:ea typeface="+mn-ea"/>
                <a:cs typeface="+mn-cs"/>
              </a:defRPr>
            </a:lvl3pPr>
            <a:lvl4pPr marL="1600200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
              <a:spcBef>
                <a:spcPct val="20000"/>
              </a:spcBef>
              <a:buFont typeface="Arial" pitchFamily="34" charset="0"/>
              <a:buChar char="–"/>
              <a:defRPr sz="2000" kern="1200">
                <a:solidFill>
                  <a:schemeClr val="tx1"/>
                </a:solidFill>
                <a:latin typeface="+mn-lt"/>
                <a:ea typeface="+mn-ea"/>
                <a:cs typeface="+mn-cs"/>
              </a:defRPr>
            </a:lvl4pPr>
            <a:lvl5pPr marL="2057400" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
              <a:spcBef>
                <a:spcPct val="20000"/>
              </a:spcBef>
              <a:buFont typeface="Arial" pitchFamily="34" charset="0"/>
              <a:buChar char="»"/>
              <a:defRPr sz="2000" kern="1200">
                <a:solidFill>
                  <a:schemeClr val="tx1"/>
                </a:solidFill>
                <a:latin typeface="+mn-lt"/>
                <a:ea typeface="+mn-ea"/>
                <a:cs typeface="+mn-cs"/>
              </a:defRPr>
            </a:lvl5pPr>
            <a:lvl6pPr marL="2514600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
              <a:spcBef>
                <a:spcPct val="20000"/>
              </a:spcBef>
              <a:buFont typeface="Arial" pitchFamily="34" charset="0"/>
              <a:buChar char="•"/>
              <a:defRPr sz="2000" kern="1200">
                <a:solidFill>
                  <a:schemeClr val="tx1"/>
                </a:solidFill>
                <a:latin typeface="+mn-lt"/>
                <a:ea typeface="+mn-ea"/>
                <a:cs typeface="+mn-cs"/>
              </a:defRPr>
            </a:lvl6pPr>
            <a:lvl7pPr marL="2971800" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
              <a:spcBef>
                <a:spcPct val="20000"/>
              </a:spcBef>
              <a:buFont typeface="Arial" pitchFamily="34" charset="0"/>
              <a:buChar char="•"/>
              <a:defRPr sz="2000" kern="1200">
                <a:solidFill>
                  <a:schemeClr val="tx1"/>
                </a:solidFill>
                <a:latin typeface="+mn-lt"/>
                <a:ea typeface="+mn-ea"/>
                <a:cs typeface="+mn-cs"/>
              </a:defRPr>
            </a:lvl7pPr>
            <a:lvl8pPr marL="3429000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
              <a:spcBef>
                <a:spcPct val="20000"/>
              </a:spcBef>
              <a:buFont typeface="Arial" pitchFamily="34" charset="0"/>
              <a:buChar char="•"/>
              <a:defRPr sz="2000" kern="1200">
                <a:solidFill>
                  <a:schemeClr val="tx1"/>
                </a:solidFill>
                <a:latin typeface="+mn-lt"/>
                <a:ea typeface="+mn-ea"/>
                <a:cs typeface="+mn-cs"/>
              </a:defRPr>
            </a:lvl8pPr>
            <a:lvl9pPr marL="3886200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
              <a:spcBef>
                <a:spcPct val="20000"/>
              </a:spcBef>
              <a:buFont typeface="Arial" pitchFamily="34" charset="0"/>
              <a:buChar char="•"/>
              <a:defRPr sz="2000" kern="1200">
                <a:solidFill>
                  <a:schemeClr val="tx1"/>
                </a:solidFill>
                <a:latin typeface="+mn-lt"/>
                <a:ea typeface="+mn-ea"/>
                <a:cs typeface="+mn-cs"/>
              </a:defRPr>
            </a:lvl9pPr>
          </p:bodyStyle>
        </xsl:if>
        <p:otherStyle>
          <a:defPPr>
            <a:defRPr lang="en-US"/>
          </a:defPPr>
          <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
            <a:defRPr sz="1800" kern="1200">
              <a:solidFill>
                <a:schemeClr val="tx1"/>
              </a:solidFill>
              <a:latin typeface="+mn-lt"/>
              <a:ea typeface="+mn-ea"/>
              <a:cs typeface="+mn-cs"/>
            </a:defRPr>
          </a:lvl1pPr>
          <a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
            <a:defRPr sz="1800" kern="1200">
              <a:solidFill>
                <a:schemeClr val="tx1"/>
              </a:solidFill>
              <a:latin typeface="+mn-lt"/>
              <a:ea typeface="+mn-ea"/>
              <a:cs typeface="+mn-cs"/>
            </a:defRPr>
          </a:lvl2pPr>
          <a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
            <a:defRPr sz="1800" kern="1200">
              <a:solidFill>
                <a:schemeClr val="tx1"/>
              </a:solidFill>
              <a:latin typeface="+mn-lt"/>
              <a:ea typeface="+mn-ea"/>
              <a:cs typeface="+mn-cs"/>
            </a:defRPr>
          </a:lvl3pPr>
          <a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
            <a:defRPr sz="1800" kern="1200">
              <a:solidFill>
                <a:schemeClr val="tx1"/>
              </a:solidFill>
              <a:latin typeface="+mn-lt"/>
              <a:ea typeface="+mn-ea"/>
              <a:cs typeface="+mn-cs"/>
            </a:defRPr>
          </a:lvl4pPr>
          <a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
            <a:defRPr sz="1800" kern="1200">
              <a:solidFill>
                <a:schemeClr val="tx1"/>
              </a:solidFill>
              <a:latin typeface="+mn-lt"/>
              <a:ea typeface="+mn-ea"/>
              <a:cs typeface="+mn-cs"/>
            </a:defRPr>
          </a:lvl5pPr>
          <a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
            <a:defRPr sz="1800" kern="1200">
              <a:solidFill>
                <a:schemeClr val="tx1"/>
              </a:solidFill>
              <a:latin typeface="+mn-lt"/>
              <a:ea typeface="+mn-ea"/>
              <a:cs typeface="+mn-cs"/>
            </a:defRPr>
          </a:lvl6pPr>
          <a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
            <a:defRPr sz="1800" kern="1200">
              <a:solidFill>
                <a:schemeClr val="tx1"/>
              </a:solidFill>
              <a:latin typeface="+mn-lt"/>
              <a:ea typeface="+mn-ea"/>
              <a:cs typeface="+mn-cs"/>
            </a:defRPr>
          </a:lvl7pPr>
          <a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
            <a:defRPr sz="1800" kern="1200">
              <a:solidFill>
                <a:schemeClr val="tx1"/>
              </a:solidFill>
              <a:latin typeface="+mn-lt"/>
              <a:ea typeface="+mn-ea"/>
              <a:cs typeface="+mn-cs"/>
            </a:defRPr>
          </a:lvl8pPr>
          <a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
            <a:defRPr sz="1800" kern="1200">
              <a:solidFill>
                <a:schemeClr val="tx1"/>
              </a:solidFill>
              <a:latin typeface="+mn-lt"/>
              <a:ea typeface="+mn-ea"/>
              <a:cs typeface="+mn-cs"/>
            </a:defRPr>
          </a:lvl9pPr>
        </p:otherStyle>
      </p:txStyles>
      </xsl:for-each>
    </p:sldMaster>

  </xsl:template>
  
  <!--Recursive Functions-->
  <xsl:template name ="BoldItalic">
    <xsl:param name ="outlineNo" />
    <xsl:param name ="paramName"/>
    <xsl:param name="smName"/>
    <xsl:choose >
      <xsl:when test ="$outlineNo =1">
        <xsl:variable name="outline1">
          <xsl:value-of select="concat($smName,'-outline1')"/>
        </xsl:variable>
        <xsl:for-each select ="/office:document-styles/office:styles/style:style[@style:name=$outline1]/style:text-properties" >
          <xsl:choose >
            <xsl:when test ="$paramName='b' and @fo:font-weight">
              <xsl:value-of  select ="@fo:font-weight"/>
            </xsl:when>
            <xsl:when test="$paramName='i' and @fo:font-style">
              <xsl:value-of  select ="@fo:font-style"/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of  select ="'normal'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="$outlineNo &gt; 1 and $outlineNo &lt;= 9">
        <xsl:variable name="outlineStyle">
          <xsl:value-of select="concat($smName,'-outline',$outlineNo)"/>
        </xsl:variable>
        <xsl:for-each select ="/office:document-styles/office:styles/style:style[@style:name=$outlineStyle]/style:text-properties" >
          <xsl:choose >
            <xsl:when test ="$paramName='b' and @fo:font-weight">
              <xsl:value-of  select ="@fo:font-weight"/>
            </xsl:when>
            <xsl:when test ="$paramName='i' and @fo:font-style">
              <xsl:value-of  select ="@fo:font-style"/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:call-template name ="BoldItalic">
                <xsl:with-param name ="outlineNo" select ="$outlineNo - 1" />
                <xsl:with-param name="smName" select="$smName"/>
                <xsl:with-param name ="paramName" select ="$paramName"/>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name ="ULine">
    <xsl:param name ="outlineNo" />
    <xsl:param name ="paramName"/>
    <xsl:param name="smName"/>
    <xsl:choose >
      <xsl:when test ="$outlineNo =1">
        <xsl:variable name="outline1">
          <xsl:value-of select="concat($smName,'-outline1')"/>
        </xsl:variable>
        <xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outline1]/style:text-properties" >
          <xsl:choose >
            <xsl:when test="$paramName='u' and @style:text-underline-style and @style:text-underline-width or @style:text-underline-type">
              <xsl:call-template name="Underline">
                <xsl:with-param name="uStyle" select="@style:text-underline-style"/>
                <xsl:with-param name="uWidth" select="@style:text-underline-width"/>
                <xsl:with-param name="uType" select="@style:text-underline-type"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="$paramName='val' and @fo:color">
              <xsl:value-of select="substring-after(@fo:color,'#')"/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of  select ="'none'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="$outlineNo &gt; 1 and $outlineNo &lt;= 9">
        <xsl:variable name="outlineStyle">
          <xsl:value-of select="concat($smName,'-outline',$outlineNo)"/>
        </xsl:variable>
        <xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineStyle]/style:text-properties" >
          <xsl:choose >
            <xsl:when test="$paramName='u' and @style:text-underline-style and @style:text-underline-width or @style:text-underline-type">
              <xsl:call-template name="Underline">
                <xsl:with-param name="uStyle" select="@style:text-underline-style"/>
                <xsl:with-param name="uWidth" select="@style:text-underline-width"/>
                <xsl:with-param name="uType" select="@style:text-underline-type"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="$paramName='val' and @fo:color">
              <xsl:value-of select="substring-after(@fo:color,'#')"/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:call-template name ="ULine">
                <xsl:with-param name ="outlineNo" select ="$outlineNo - 1" />
                <xsl:with-param name="smName" select="$smName"/>
                <xsl:with-param name ="paramName" select ="$paramName"/>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:when>
      
    </xsl:choose>
  </xsl:template>
  <xsl:template name ="FontColor">
    <xsl:param name ="lineNo" />
    <xsl:param name ="parName"/>
    <xsl:param name="smasterName"/>
    <xsl:choose >
      <xsl:when test ="$lineNo =1">
        <xsl:variable name="outline1">
          <xsl:value-of select="concat($smasterName,'-outline1')"/>
        </xsl:variable>
        <xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outline1]/style:text-properties" >
          <xsl:choose >
            <xsl:when test="$parName='val' and @fo:color">
              <xsl:value-of select="substring-after(@fo:color,'#')"/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of  select ="'000000'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="$lineNo &gt; 1 and $lineNo &lt;= 9">
        <xsl:variable name="outlineStyle">
          <xsl:value-of select="concat($smasterName,'-outline',$lineNo)"/>
        </xsl:variable>
        <xsl:for-each select ="/office:document-styles/office:styles/style:style[@style:name=$outlineStyle]/style:text-properties" >
          <xsl:choose >
            <xsl:when test="$parName='val' and @fo:color">
              <xsl:value-of select="substring-after(@fo:color,'#')"/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:call-template name ="FontColor">
                <xsl:with-param name ="lineNo" select ="$lineNo - 1" />
                <xsl:with-param name="smasterName" select="$smasterName"/>
                <xsl:with-param name ="parName" select ="$parName"/>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name ="FontFamily">
    <xsl:param name ="lineNo" />
    <xsl:param name ="parName"/>
    <xsl:param name="smasterName"/>
    <xsl:choose >
      <xsl:when test ="$lineNo =1">
        <xsl:variable name="outline1">
          <xsl:value-of select="concat($smasterName,'-outline1')"/>
        </xsl:variable>
        <xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outline1]/style:text-properties" >
          <xsl:choose >
            <xsl:when test="$parName='typeface' and @fo:font-family">
              <xsl:value-of select="translate(@fo:font-family, &quot;'&quot;,'')"/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of  select ="'Arial'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="$lineNo &gt; 1 and $lineNo &lt;= 9">
        <xsl:variable name="outlineStyle">
          <xsl:value-of select="concat($smasterName,'-outline',$lineNo)"/>
        </xsl:variable>
        <xsl:for-each select ="/office:document-styles/office:styles/style:style[@style:name=$outlineStyle]/style:text-properties" >
          <xsl:choose >
            <xsl:when test="$parName='typeface' and @fo:font-family">
              <xsl:value-of select="translate(@fo:font-family, &quot;'&quot;,'')"/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:call-template name ="FontFamily">
                <xsl:with-param name ="lineNo" select ="$lineNo - 1" />
                <xsl:with-param name="smasterName" select="$smasterName"/>
                <xsl:with-param name ="parName" select ="$parName"/>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name ="Spacing">
    <xsl:param name ="lineNo" />
    <xsl:param name ="parName"/>
    <xsl:param name="smasterName"/>
    <xsl:variable name="outlineNo">
      <xsl:value-of select="concat($smasterName,'-outline',$lineNo)"/>
    </xsl:variable>

    <xsl:for-each select ="/office:document-styles/office:styles/style:style[@style:name=$outlineNo]/style:text-properties" >
      <xsl:choose >
        <xsl:when test="$parName='spc' and @fo:letter-spacing">
          <xsl:variable name="Unit">
            <xsl:call-template name="getConvertUnit">
              <xsl:with-param name="length" select="@fo:letter-spacing"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="letterSpacing">
            <xsl:call-template name="convertUnitsToCm">
              <xsl:with-param name ="length" select ="@fo:letter-spacing"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:if test ="substring-before(@fo:letter-spacing,$Unit)&lt; 0 ">
            <xsl:value-of select ="format-number($letterSpacing * 7200 div 2.54 ,'#')"/>
          </xsl:if >
          <xsl:if test ="substring-before(@fo:letter-spacing,$Unit)
						&gt; 0 or substring-before(@fo:letter-spacing,$Unit) = 0 ">
            <xsl:value-of select ="format-number(($letterSpacing * 72 div 2.54) *100 ,'#')"/>
          </xsl:if>
        </xsl:when>
        <xsl:otherwise >
          <xsl:choose>
            <xsl:when test="$lineNo=1">
              <xsl:value-of  select ="'0'"/>
            </xsl:when>
            <xsl:when test="$lineNo &gt; 1 and $lineNo &lt;= 9">
              <xsl:call-template name ="Spacing">
                <xsl:with-param name ="lineNo" select ="1" />
                <xsl:with-param name="smasterName" select="$smasterName"/>
                <xsl:with-param name ="parName" select ="$parName"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>

  </xsl:template>
  <xsl:template name ="LineSpacingPoints">
    <xsl:param name ="lineNo" />
    <xsl:param name="smasterName"/>
    <xsl:variable name="outlineNo">
      <xsl:value-of select="concat($smasterName,'-outline',$lineNo)"/>
    </xsl:variable>
    <xsl:for-each select ="/office:document-styles/office:styles/style:style[@style:name=$outlineNo]/style:paragraph-properties" >
      <xsl:variable name="Unit">
        <xsl:call-template name="getConvertUnit">
          <xsl:with-param name="length" select="@style:line-spacing"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="Unit1">
        <xsl:call-template name="getConvertUnit">
          <xsl:with-param name="length" select="@style:line-height-at-least"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="@fo:line-height or @style:line-spacing or @style:line-height-at-least ">
          <xsl:if test ="@fo:line-height and 
					substring-before(@fo:line-height,'%') &gt; 0 and 
					not(substring-before(@fo:line-height,'%') = 100)">
            <xsl:element name="a:lnSpc">
              <xsl:element name="a:spcPct">
                <xsl:attribute name ="val">
                  <xsl:value-of select ="format-number(substring-before(@fo:line-height,'%')* 1000,'#.##') "/>
                </xsl:attribute>
              </xsl:element>
            </xsl:element>
          </xsl:if>
          <xsl:if test ="@style:line-spacing and 
					substring-before(@style:line-spacing,$Unit) &gt; 0">
            <xsl:variable name="lineSpacing">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name ="length" select ="@style:line-spacing"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:element name="a:lnSpc">
              <xsl:element name="a:spcPts">
                <xsl:attribute name ="val">
                  <xsl:value-of select ="round($lineSpacing* 2835) "/>
                </xsl:attribute>
              </xsl:element>
            </xsl:element>
          </xsl:if>
          <xsl:if test ="@style:line-height-at-least and 
					substring-before(@style:line-height-at-least,$Unit1) &gt; 0 ">
            <xsl:variable name="lineSpacing">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name ="length" select ="@style:line-height-at-least"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:element name="a:lnSpc">
              <xsl:element name="a:spcPts">
                <xsl:attribute name ="val">
                  <xsl:value-of select ="round($lineSpacing* 2835) "/>
                </xsl:attribute>
              </xsl:element>
            </xsl:element>
          </xsl:if>
        </xsl:when>
        <xsl:otherwise >
          <xsl:if test="$lineNo !=1">
            <xsl:call-template name ="LineSpacingPoints">
              <xsl:with-param name ="lineNo" select ="1" />
              <xsl:with-param name="smasterName" select="$smasterName"/>
            </xsl:call-template>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="Alignment">
    <xsl:param name ="outlineNo" />
    <xsl:param name ="paramName"/>
    <xsl:param name="smName"/>
    <xsl:variable name="outlineStyle">
      <xsl:value-of select="concat($smName,'-outline',$outlineNo)"/>
    </xsl:variable>
    <xsl:for-each select ="/office:document-styles/office:styles/style:style[@style:name=$outlineStyle]/style:paragraph-properties" >
      <xsl:choose >
        <xsl:when test ="$paramName='algn' and @fo:text-align">
          <xsl:choose >
            <xsl:when test ="@fo:text-align='center'">
              <xsl:value-of select ="'ctr'"/>
            </xsl:when>
            <xsl:when test ="@fo:text-align='end' or @fo:text-align='right'">
              <xsl:value-of select ="'r'"/>
            </xsl:when>
            <xsl:when test ="@fo:text-align='justify'">
              <xsl:value-of select ="'just'"/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select ="'l'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise >
          <xsl:choose>
            <xsl:when test="$outlineNo=1">
              <xsl:value-of  select ="'l'"/>
            </xsl:when>
            <xsl:when test="$outlineNo &gt; 1 and $outlineNo &lt;= 9">
              <xsl:call-template name ="Alignment">
                <xsl:with-param name ="outlineNo" select ="1" />
                <xsl:with-param name="smName" select="$smName"/>
                <xsl:with-param name ="paramName" select ="$paramName"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>
  <!--End Recursive Functions-->
  <xsl:template name="Titlebody">
    <xsl:param name="slideMasterName"/>
    <xsl:variable name="prStyleName">
      <xsl:choose>
        <xsl:when test="@presentation:style-name[contains(.,'title')]">
          <xsl:value-of select="@presentation:style-name"/>
        </xsl:when>
        <xsl:when test="not(@presentation:style-name[contains(.,'outline')])">
          <xsl:value-of select="concat($slideMasterName,'-title')"/>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="lang">
      <xsl:for-each select ="/office:document-styles/office:styles/style:style[@style:name=$prStyleName]/style:text-properties">
        <xsl:if test="position()=1">
          <xsl:choose>
            <xsl:when test="@style:language-asian and @style:country-asian">
                <xsl:value-of select="concat(@style:language-asian,'-',@style:country-asian)"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="'en-US'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
      </xsl:for-each>
    </xsl:variable>
    <a:p>
      <a:r>
        <a:rPr dirty="0" smtClean="0">
        <xsl:attribute name="lang">
          <xsl:choose>
            <xsl:when test="$lang!=''">
              <xsl:value-of select="$lang"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'en-US'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        </a:rPr>
        <xsl:if test="./@presentation:class[contains(.,'title')]">
          <xsl:choose>
            <xsl:when test="./draw:text-box/text:p/text:span">
              <a:t>
                <xsl:for-each select="./draw:text-box/text:p/text:span">
                  <xsl:value-of select="."/>
                </xsl:for-each>
              </a:t>
            </xsl:when>
            <xsl:when test="./draw:text-box/text:p">
              <a:t>
                <xsl:for-each select="./draw:text-box/text:p">
                  <xsl:value-of select="."/>
                </xsl:for-each>
              </a:t>
            </xsl:when>
            <xsl:when  test="@presentation:placeholder='true'">
              <a:t>
                <xsl:value-of select="'Click to edit the title text format'"/>
              </a:t>
            </xsl:when>
            <xsl:otherwise>
              <a:t>
                <xsl:value-of select="'Click to edit the title text format'"/>
              </a:t>
            </xsl:otherwise>
          </xsl:choose>
          <!--<xsl:if test="./draw:text-box/text:p">
              <a:t>
                <xsl:for-each select="./draw:text-box/text:p">
                  <xsl:value-of select="."/>
                </xsl:for-each>
              </a:t>
            </xsl:if>
            <xsl:if test="not(./draw:text-box/text:p)">
              <a:t>Click to edit the title text format</a:t>
            </xsl:if>-->
        </xsl:if>
        <!--</a:t>-->
      </a:r>
      <a:endParaRPr lang="en-US" dirty="0"/>
    </a:p>
  </xsl:template>
  <xsl:template name="Outlinebody">
    <xsl:param name="StyleName"/>
    <xsl:param name="slideMasterName"/>
    <a:lstStyle/>
    <xsl:choose >
      <xsl:when test ="draw:text-box/text:list">
        <xsl:for-each select ="draw:text-box">
          <xsl:for-each select ="text:list">
            <a:p>
              <a:pPr >
                <xsl:attribute name ="lvl" >
                  <xsl:if test="position() &lt; 10">
                    <xsl:value-of  select ="position()-1"/>
                  </xsl:if>
                  <xsl:if test="position() &gt; 9">
                    <xsl:value-of  select ="0"/>
                  </xsl:if>
                </xsl:attribute>
              </a:pPr >
              <a:r>
                <a:rPr lang="en-US" smtClean="0"/>
                <a:t>
                  <xsl:value-of select="."/>
                </a:t>
              </a:r>
            </a:p>
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>
      <xsl:otherwise >
        <xsl:call-template name ="defaultOutline">
          <xsl:with-param name="StyleName" select="$StyleName"/>
          <xsl:with-param name="slideMasterName" select="$slideMasterName"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>
  <xsl:template name ="defaultOutline">
    <xsl:param name="StyleName"/>
    <xsl:param name="slideMasterName"/>
    <a:p>
      <a:pPr lvl="0"/>
      <a:r>
        <a:rPr smtClean="0">
          <xsl:call-template name="tmpOutlineLanguage">
            <xsl:with-param name="level" select="'1'"/>
            <xsl:with-param name="StyleName" select="$StyleName"/>
            <xsl:with-param name="slideMasterName" select="$slideMasterName"/>
          </xsl:call-template>
        </a:rPr>
        <a:t>
          <xsl:value-of select="'Click to edit the outline text format'"/>
        </a:t>
      </a:r>
    </a:p>
    <a:p>
      <a:pPr lvl="1"/>
      <a:r>
        <a:rPr smtClean="0">
          <xsl:call-template name="tmpOutlineLanguage">
            <xsl:with-param name="level" select="'2'"/>
            <xsl:with-param name="StyleName" select="$StyleName"/>
            <xsl:with-param name="slideMasterName" select="$slideMasterName"/>
          </xsl:call-template>
        </a:rPr>
        <a:t>
          <xsl:value-of select="'Second Outline Level'"/>
        </a:t>
      </a:r>
    </a:p>
    <a:p>
      <a:pPr lvl="2"/>
      <a:r>
        <a:rPr smtClean="0">
          <xsl:call-template name="tmpOutlineLanguage">
            <xsl:with-param name="level" select="'3'"/>
            <xsl:with-param name="StyleName" select="$StyleName"/>
            <xsl:with-param name="slideMasterName" select="$slideMasterName"/>
          </xsl:call-template>
        </a:rPr>
        <a:t>
          <xsl:value-of select="'Third Outline Level'"/>
        </a:t>
      </a:r>
    </a:p>
    <a:p>
      <a:pPr lvl="3"/>
      <a:r>
      <a:rPr smtClean="0">
          <xsl:call-template name="tmpOutlineLanguage">
            <xsl:with-param name="level" select="'4'"/>
            <xsl:with-param name="StyleName" select="$StyleName"/>
            <xsl:with-param name="slideMasterName" select="$slideMasterName"/>
          </xsl:call-template>
        </a:rPr>
        <a:t>
          <xsl:value-of select="'Fourth Outline Level'"/>
        </a:t>
      </a:r>
    </a:p>
    <a:p>
      <a:pPr lvl="4"/>
      <a:r>
        <a:rPr smtClean="0">
          <xsl:call-template name="tmpOutlineLanguage">
            <xsl:with-param name="level" select="'5'"/>
            <xsl:with-param name="StyleName" select="$StyleName"/>
            <xsl:with-param name="slideMasterName" select="$slideMasterName"/>
          </xsl:call-template>
        </a:rPr>
        <a:t>
          <xsl:value-of select="'Fifth Outline Level'"/>
        </a:t>
      </a:r>
    </a:p>
    <a:p>
      <a:pPr lvl="5"/>
      <a:r>
        <a:rPr smtClean="0">
          <xsl:call-template name="tmpOutlineLanguage">
            <xsl:with-param name="level" select="'6'"/>
            <xsl:with-param name="StyleName" select="$StyleName"/>
            <xsl:with-param name="slideMasterName" select="$slideMasterName"/>
          </xsl:call-template>
        </a:rPr>
        <a:t>
          <xsl:value-of select="'Sixth Outline Level'"/>
        </a:t>
      </a:r>
    </a:p>
    <a:p>
      <a:pPr lvl="6"/>
      <a:r>
        <a:rPr smtClean="0">
          <xsl:call-template name="tmpOutlineLanguage">
            <xsl:with-param name="level" select="'7'"/>
            <xsl:with-param name="StyleName" select="$StyleName"/>
            <xsl:with-param name="slideMasterName" select="$slideMasterName"/>
          </xsl:call-template>
        </a:rPr>
        <a:t>
          <xsl:value-of select="'Seventh Outline Level'"/>
        </a:t>
      </a:r>
    </a:p>
    <a:p>
      <a:pPr lvl="7"/>
      <a:r>
      <a:rPr smtClean="0">
          <xsl:call-template name="tmpOutlineLanguage">
            <xsl:with-param name="level" select="'8'"/>
            <xsl:with-param name="StyleName" select="$StyleName"/>
            <xsl:with-param name="slideMasterName" select="$slideMasterName"/>
          </xsl:call-template>
        </a:rPr>
        <a:t>
          <xsl:value-of select="'Eighth Outline Level'"/>
        </a:t>
      </a:r>
    </a:p>
    <a:p>
      <a:pPr lvl="8"/>
      <a:r>
        <a:rPr smtClean="0">
          <xsl:call-template name="tmpOutlineLanguage">
            <xsl:with-param name="level" select="'9'"/>
            <xsl:with-param name="StyleName" select="$StyleName"/>
            <xsl:with-param name="slideMasterName" select="$slideMasterName"/>
          </xsl:call-template>
        </a:rPr>
        <a:t>
          <xsl:value-of select="'Ninth Outline Level'"/>
        </a:t>
      </a:r>
      <a:endParaRPr lang="en-US"/>
    </a:p>
  </xsl:template>
  <xsl:template name="tmpOutlineLanguage">
    <xsl:param name="level"/>
    <xsl:param name="StyleName"/>
    <xsl:param name="slideMasterName"/>
    <xsl:variable name="prStyleName">
      <xsl:choose>
        <xsl:when test="$StyleName!=''">
          <xsl:value-of select="$StyleName"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="concat($slideMasterName,'-outline',$level)"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
  
      <xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$prStyleName]/style:text-properties">
        <xsl:if test="position()=1">
          <xsl:choose>
            <xsl:when test="@style:language-asian and @style:country-asian">
              <xsl:attribute name ="lang">
                <xsl:value-of select="concat(@style:language-asian,'-',@style:country-asian)"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name ="lang">
                <xsl:value-of select="'en-US'"/>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
      </xsl:for-each>
   
  </xsl:template>
  <xsl:template name ="Datetime">
    <xsl:param name="slideMasterName"/>
    <xsl:param name ="prId"/>
    <xsl:param name="className"/>
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr name="Date Placeholder 3" id="16363"/>
        <p:cNvSpPr>
          <a:spLocks noGrp="1" />
        </p:cNvSpPr>
        <p:nvPr>
          <p:ph type="dt" sz="half" idx="2"/>
        </p:nvPr>
      </p:nvSpPr>
      <p:spPr>
        <xsl:call-template name ="tmpdrawCordinates"/>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
        <!-- Solid fill color -->
        <xsl:for-each select ="document('styles.xml')/office:document-styles/office:automatic-styles/style:style[@style:name=$prId]/style:graphic-properties ">
          <xsl:call-template name="tmpSMShapeFillColor">
            <xsl:with-param name ="shapeCount" select ="'DateTime'"/>
          </xsl:call-template>
        </xsl:for-each>
        <!--Line Color-->
        <xsl:variable name="var_PrStyleId" select="@presentation:style-name"/>
        <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$var_PrStyleId]/style:graphic-properties">
          <xsl:if test="position()=1">
          <xsl:call-template name ="getLineStyle"/>
          </xsl:if>
        </xsl:for-each>
        <!--End-->
      </p:spPr>
      <p:txBody>
        <xsl:call-template name="SlideMasterTextAlignment">
          <xsl:with-param name="prId" select="$prId"/>
        </xsl:call-template>
        <xsl:call-template name="tmpDeflistProp">
          <xsl:with-param name="prId" select="$prId"/>
                  </xsl:call-template>
                       <xsl:call-template name="tmpSMDateTimeText"/>
      </p:txBody>
    </p:sp >
  </xsl:template>
  <xsl:template name="tmpSMDateTimeText">
        <xsl:for-each select="./draw:text-box">
          <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:pPr>
              <xsl:call-template name="SMParagraphStyles">
                <xsl:with-param name ="paraId"  select="text:p/@text:style-name"/>
              </xsl:call-template>
            </a:pPr>
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
                          <a:rPr  dirty="0" smtClean="0">
                            
                                <xsl:variable name ="textId">
                                  <xsl:value-of select ="./parent::node()/@text:style-name"/>
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
                              <a:t> </a:t>
                            </a:fld>
                          </xsl:when>
                          <xsl:when test="node() or text:date">
                            <a:r>
                          <a:rPr  dirty="0" smtClean="0">
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
                      <xsl:when test ="name()='presentation:date-time'">
                        <a:fld >
                          <xsl:attribute name ="id">
                            <xsl:value-of select ="'{11859447-1FF1-4571-A7A1-75F1CAF7F5D9}'"/>
                          </xsl:attribute>
                          <xsl:attribute name ="type">
                            <xsl:value-of select ="'datetime1'"/>
                          </xsl:attribute>
                      <a:rPr  dirty="0" smtClean="0">
                            <xsl:variable name ="textId">
                              <xsl:value-of select ="./parent::node()/@text:style-name"/>
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
                          <a:t> </a:t>
                        </a:fld>
                      </xsl:when>
                  <xsl:when test="not(node())">
                    <a:r>
                      <a:rPr  dirty="0" smtClean="0">
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
              </xsl:if>
            </xsl:for-each>
          </a:p>
        </xsl:for-each>
       </xsl:template>
  <xsl:template name ="Footer">
    <xsl:param name="slideMasterName"/>
    <xsl:param name ="prId"/>
    <xsl:param name="className"/>
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr name="Footer Placeholder 4" id="16361" />
        <p:cNvSpPr>
          <a:spLocks noGrp="1" />
        </p:cNvSpPr>
        <p:nvPr>
          <p:ph type="ftr" sz="quarter" idx="3"/>
        </p:nvPr>
      </p:nvSpPr>
      <p:spPr>
        <xsl:call-template name ="tmpdrawCordinates"/>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
        <!-- Solid fill color -->
        <xsl:for-each select ="document('styles.xml')/office:document-styles/office:automatic-styles/style:style[@style:name=$prId]/style:graphic-properties ">
          <xsl:call-template name="tmpSMShapeFillColor">
            <xsl:with-param name ="shapeCount" select ="'footer'"/>
          </xsl:call-template>
        </xsl:for-each>
        <!--Line Color-->
        <xsl:variable name="var_PrStyleId" select="@presentation:style-name"/>
        <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$var_PrStyleId]/style:graphic-properties">
          <xsl:if test="position()=1">
          <xsl:call-template name ="getLineStyle"/>
          </xsl:if>
        </xsl:for-each>
        <!--End-->
      </p:spPr>
      <p:txBody>
        <xsl:call-template name="SlideMasterTextAlignment">
          <xsl:with-param name="prId" select="$prId"/>
        </xsl:call-template>
        <xsl:call-template name="tmpDeflistProp">
          <xsl:with-param name="prId" select="$prId"/>
                  </xsl:call-template>
           <xsl:for-each select="./draw:text-box">
          <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:pPr>
              <xsl:call-template name="SMParagraphStyles">
                <xsl:with-param name ="paraId"  select="text:p/@text:style-name"/>
              </xsl:call-template>
            </a:pPr>
            <xsl:for-each select ="node()">
              <xsl:if test ="name()='text:p'" >
                <xsl:variable name="paraId" select="@text:style-name"/>
                <xsl:if test ="child::node()">

                  <xsl:for-each select ="node()">
                    <xsl:choose >
                      <xsl:when test ="name()='text:span'">
                        <xsl:choose>
                          <xsl:when test ="presentation:footer">
                            <a:r>
                              <a:rPr  dirty="0" smtClean="0">
                                <xsl:variable name ="textId">
                                  <xsl:value-of select ="@text:style-name"/>
                                </xsl:variable>
                                <xsl:if test ="not($textId ='')">
                                  <xsl:call-template name ="tmpSMfontStyles">
                                    <xsl:with-param name ="TextStyleID" select ="$textId" />
                                  </xsl:call-template>
                                </xsl:if>
                              </a:rPr>
                              <a:t></a:t>
                            </a:r>
                          </xsl:when>
                          <xsl:when test="node()">
                            <a:r>
                              <a:rPr  dirty="0" smtClean="0">
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
                          </xsl:when>
                        </xsl:choose>
                      </xsl:when >
                      <!--<xsl:otherwise>
                        <xsl:if test="string-length(.) != 0">
                          <a:r>
                            <a:rPr  dirty="0" smtClean="0">
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
                        </xsl:if>
                      </xsl:otherwise>-->
                    </xsl:choose>
                  </xsl:for-each>
                </xsl:if>
              </xsl:if>
            </xsl:for-each>
            <a:endParaRPr  dirty="0" smtClean="0"/>
          </a:p>
        </xsl:for-each>
      </p:txBody>
    </p:sp >
  </xsl:template>
  <xsl:template name="tmpDeflistProp">
    <xsl:param name ="prId"/>
           <a:lstStyle>
          <a:lvl1pPr>
        <!--<xsl:value-of select="'ctr'"/>-->

            <xsl:variable name ="ParId">
              <xsl:choose>
                <xsl:when test="./draw:text-box/text:p/@text:style-name">
                  <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$prId"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$ParId]">
              <xsl:attribute name="algn">
                <xsl:choose>
                  <xsl:when test ="style:paragraph-properties/@fo:text-align='center'">
                    <xsl:value-of select ="'ctr'"/>
                  </xsl:when>
                  <xsl:when test ="style:paragraph-properties/@fo:text-align='end'">
                    <xsl:value-of select ="'r'"/>
                  </xsl:when>
                  <xsl:when test ="style:paragraph-properties/@fo:text-align='justify'">
                    <xsl:value-of select ="'just'"/>
                  </xsl:when>
                  <xsl:otherwise >
                    <xsl:value-of select ="'l'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:for-each>

            <a:defRPr>

          <xsl:variable name="textId">
                <xsl:value-of select="./draw:text-box/text:p/text:span/@text:style-name"/>
              </xsl:variable>
              <xsl:choose>
            <xsl:when test="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@fo:font-size">
                  <xsl:attribute name="sz">
                <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties">
<!--Office 2007 Sp2-->

                  <xsl:variable name="fontSize">
                    <xsl:call-template name="point-measure">
                      <xsl:with-param name="length" select="@fo:font-size"/>
                    </xsl:call-template>
                  </xsl:variable>
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="concat($fontSize,'pt')"/>
                      </xsl:call-template>
                    </xsl:for-each>
                  </xsl:attribute>
                </xsl:when >
            <!--When draw:text-box/text:p/text:span is not there-->
                <xsl:when test="document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties/@fo:font-size">
                  <xsl:attribute name="sz">
                    <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties">
<!--Office 2007 Sp2-->

                      <xsl:variable name="fontSize">
                        <xsl:call-template name="point-measure">
                          <xsl:with-param name="length" select="@fo:font-size"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="concat($fontSize,'pt')"/>
                      </xsl:call-template>
                    </xsl:for-each>
                  </xsl:attribute>
                </xsl:when>
            <xsl:when test="document('styles.xml')//style:style[@style:name=$prId]/style:text-properties/@fo:font-size">
              <xsl:attribute name="sz">
                <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$prId]/style:text-properties">
<!--Office 2007 Sp2-->

                  <xsl:variable name="fontSize">
                    <xsl:call-template name="point-measure">
                      <!--changed by yeswanth.s : 6-Feb-09-->
                      <xsl:with-param name="length" select="@fo:font-size"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:call-template name ="convertToPoints">
                    <xsl:with-param name ="unit" select ="'pt'"/>
                    <xsl:with-param name ="length" select ="concat($fontSize,'pt')"/>
                  </xsl:call-template>
                </xsl:for-each>
              </xsl:attribute>
            </xsl:when >
              </xsl:choose>

          <xsl:choose>
            <xsl:when test="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@style:language-asian and document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@style:country-asian">
              <xsl:attribute name ="lang">
                <xsl:value-of select="concat(document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@style:language-asian,'-',document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@style:country-asian)"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties/@style:language-asian and document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties/@style:country-asian">
              <xsl:attribute name ="lang">
                <xsl:value-of select="concat(document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties/@style:language-asian,'-',document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties/@style:country-asian)"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="document('styles.xml')//style:style[@style:name=$prId]/style:text-properties/@style:language-asian and document('styles.xml')//style:style[@style:name=$prId]/style:text-properties/@style:country-asian">
              <xsl:attribute name ="lang">
                <xsl:value-of select="concat(document('styles.xml')//style:style[@style:name=$prId]/style:text-properties/@style:language-asian,'-',document('styles.xml')//style:style[@style:name=$prId]/style:text-properties/@style:country-asian)"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name ="lang">
                <xsl:value-of select="'en-US'"/>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
          <!--sateesh-->
              <xsl:variable name="testStyleName">
                <xsl:choose>
                  <xsl:when test="./draw:text-box/text:p/text:span/@text:style-name">
                    <xsl:value-of select="./draw:text-box/text:p/text:span/@text:style-name"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="$prId"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>
              <xsl:for-each select="document('styles.xml')//style:style[@style:name=$testStyleName]">

                <!--Font Bold attribute-->
                <xsl:if test="./style:text-properties/@fo:font-weight='bold'">
                  <xsl:attribute name ="b">
                    <xsl:value-of select ="'1'"/>
                  </xsl:attribute >
                </xsl:if>
                <!--Font Italic attribute-->
                <xsl:if test="./style:text-properties/@fo:font-style='italic'">
                  <xsl:attribute name ="i">
                    <xsl:value-of select ="'1'"/>
                  </xsl:attribute >
                </xsl:if >
                <!--Font Underline-->
                <xsl:variable name ="unLine">
                  <xsl:call-template name="Underline">
                    <xsl:with-param name="uStyle" select="./style:text-properties/@style:text-underline-style"/>
                    <xsl:with-param name="uWidth" select="./style:text-properties/@style:text-underline-width"/>
                    <xsl:with-param name="uType" select="./style:text-properties/@style:text-underline-type"/>
                  </xsl:call-template>
                </xsl:variable>
                <xsl:if test ="$unLine !=''">
                  <xsl:attribute name="u">
                    <xsl:value-of  select ="$unLine"/>
                  </xsl:attribute>
                </xsl:if>
                <!-- Kerning -->
                <xsl:if test ="./style:text-properties/@style:letter-kerning = 'true'">
                  <xsl:attribute name ="kern">
                    <xsl:value-of select="1200"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test ="./style:text-properties/@style:letter-kerning = 'false'">
                  <xsl:attribute name ="kern">
                    <xsl:value-of select="0"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test ="not(./style:text-properties/@style:letter-kerning)">
                  <xsl:attribute name ="kern">
                    <xsl:value-of select="0"/>
                  </xsl:attribute>
                </xsl:if>
                <!-- End -->
                <!--Character Spacing-->
                <xsl:call-template name ="tmpCharacterSpacing"/>

              </xsl:for-each>
              <!--End-->
              <!-- Font Strike through Start-->
              <xsl:choose >
                <xsl:when  test="style:text-properties/@style:text-line-through-type = 'solid'">
                  <xsl:attribute name ="strike">
                    <xsl:value-of select ="'sngStrike'"/>
                  </xsl:attribute >
                </xsl:when >
                <xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
                  <xsl:attribute name ="strike">
                    <xsl:value-of select ="'dblStrike'"/>
                  </xsl:attribute >
                </xsl:when >
                <!-- style:text-line-through-style-->
                <xsl:when test="style:text-properties/@style:text-line-through-style = 'solid'">
                  <xsl:attribute name ="strike">
                    <xsl:value-of select ="'sngStrike'"/>
                  </xsl:attribute >
                </xsl:when>
              </xsl:choose>
              <!-- Superscript and SubScript for Text added by Mathi on 1st Aug 2007-->
              <xsl:call-template name="SuperAndSubscripts" />

              <!--Underline Color-->
              <xsl:if test ="style:text-properties/@style:text-underline-color !='font-color'">
                <a:uFill>
                  <a:solidFill>
                    <a:srgbClr>
                      <xsl:attribute name ="val">
                        <xsl:value-of select ="substring-after(style:text-properties/@style:text-underline-color,'#')"/>
                      </xsl:attribute>
                    </a:srgbClr>
                  </a:solidFill>
                </a:uFill>
              </xsl:if>
              <!--end-->
              <a:solidFill>
                <a:srgbClr>
                    <xsl:choose>
                    <xsl:when test="document('styles.xml')//style:style[@style:name=$testStyleName]/style:text-properties/@fo:color">
                      <xsl:for-each select="document('styles.xml')//style:style[@style:name=$testStyleName]">
                        <xsl:attribute name="val">
                          <xsl:value-of select="substring-after(style:text-properties/@fo:color,'#')"/>
                        </xsl:attribute>
                      </xsl:for-each>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name="val">
                        <xsl:value-of select="'000000'"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </a:srgbClr>
              </a:solidFill>
              <xsl:if test="document('styles.xml')//style:style[@style:name=$testStyleName]/style:text-properties/@fo:text-shadow">
                <a:effectLst>
                  <a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl">
                    <a:srgbClr val="000000">
                      <a:alpha val="43137"/>
                    </a:srgbClr>
                  </a:outerShdw>
                </a:effectLst>
              </xsl:if>
              <a:latin>
                <xsl:attribute name="typeface">
                  <xsl:choose>
                <xsl:when test="translate(document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@fo:font-family, &quot;'&quot;,'')">
                  <xsl:value-of select="translate(document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
                    </xsl:when>
                    <xsl:when test="translate(document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties/@fo:font-family, &quot;'&quot;,'')">
                      <xsl:value-of select="translate(document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
                    </xsl:when>
                <xsl:when test="translate(document('styles.xml')//style:style[@style:name=$prId]/style:text-properties/@fo:font-family, &quot;'&quot;,'')">
                  <xsl:value-of select="translate(document('styles.xml')//style:style[@style:name=$prId]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
                </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="'Times New Roman'"/>
                        </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
                <xsl:attribute name="pitchFamily">
                  <xsl:value-of select="0"/>
                </xsl:attribute>
                <xsl:attribute name="charset">
                  <xsl:value-of select="0"/>
                </xsl:attribute>
              </a:latin>
            </a:defRPr>
          </a:lvl1pPr>
        </a:lstStyle>
  </xsl:template>
  <xsl:template name ="Pagenumber">
    <xsl:param name ="prId"/>
    <xsl:param name="className"/>
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr name="Slide Number Placeholder 5" id="16362" />
        <p:cNvSpPr>
          <a:spLocks noGrp="1" />
        </p:cNvSpPr>
        <p:nvPr>
          <p:ph type="sldNum" sz="quarter" idx="4"/>
        </p:nvPr>
      </p:nvSpPr>
      <p:spPr>
        <xsl:call-template name ="tmpdrawCordinates"/>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
        <!-- Solid fill color -->
        <xsl:for-each select ="/office:document-styles/office:automatic-styles/style:style[@style:name=$prId]/style:graphic-properties ">
          <xsl:call-template name="tmpSMShapeFillColor">
            <xsl:with-param name ="shapeCount" select ="'page-number'"/>
          </xsl:call-template>
        </xsl:for-each>
        <!--Line Color-->
        <xsl:variable name="var_PrStyleId" select="@presentation:style-name"/>
        <xsl:for-each select ="//style:style[@style:name=$var_PrStyleId]/style:graphic-properties">
          <xsl:if test="position()=1">
          <xsl:call-template name ="getLineStyle"/>
          </xsl:if>
        </xsl:for-each>
        <!--End-->
      </p:spPr>
      <p:txBody>
        <xsl:call-template name="SlideMasterTextAlignment">
          <xsl:with-param name="prId" select="$prId"/>
        </xsl:call-template>
        <xsl:call-template name="tmpDeflistProp">
          <xsl:with-param name="prId" select="$prId"/>
        </xsl:call-template>
        <xsl:for-each select="./draw:text-box">
          <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:pPr>
              <xsl:call-template name="SMParagraphStyles">
                <xsl:with-param name ="paraId"  select="text:p/@text:style-name"/>
              </xsl:call-template>
            </a:pPr>
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
                              <a:rPr  dirty="0" smtClean="0">
                                <xsl:variable name ="textId">
                                  <xsl:value-of select ="./parent::node()/@text:style-name"/>
                                </xsl:variable>
                                <xsl:if test ="not($textId ='')">
                                  <xsl:call-template name ="tmpSMfontStyles">
                                    <xsl:with-param name ="TextStyleID" select ="$textId" />
                                  </xsl:call-template>
                                </xsl:if>
                              </a:rPr>
                              <a:t>
                                <xsl:value-of select="."/>
                              </a:t>
                            </a:fld>
                          </xsl:when>
                          <xsl:when test="node()">
                            <a:r>
                              <a:rPr  dirty="0" smtClean="0">
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
                          <a:t>
                            <xsl:value-of select="."/>
                          </a:t>
                        </a:fld>
                      </xsl:when>
                    </xsl:choose>
                  </xsl:for-each>
                </xsl:if>
              </xsl:if>
            </xsl:for-each>
          </a:p>
        </xsl:for-each>
      </p:txBody>
    </p:sp >
  </xsl:template>

  <xsl:template name ="SlideMasterTextAlignment">
    <xsl:param name="masterName"/>
    <xsl:param name ="prId"/>
    <xsl:variable name="parentName">
      <xsl:choose>
        <xsl:when test="@presentation:class='title'">
          <xsl:value-of select="concat($masterName,'-title')"/>
        </xsl:when>
        <xsl:when test="@presentation:class='outline'">
          <xsl:value-of select="concat($masterName,'-outline1')"/>
        </xsl:when>
        <xsl:when test="@presentation:class='date-time'">
          <xsl:value-of select="$prId"/>
        </xsl:when>
        <xsl:when test="@presentation:class='footer'">
          <xsl:value-of select="$prId"/>
        </xsl:when>
        <xsl:when test="@presentation:class='page-number'">
          <xsl:value-of select="$prId"/>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>

    <a:bodyPr>
      <xsl:for-each select ="draw:text-box">
        <xsl:for-each select ="text:p">
          <xsl:variable name ="ParId">
            <xsl:value-of select ="@text:style-name"/>
          </xsl:variable>
          <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$ParId]/style:paragraph-properties">
            <xsl:if test ="@style:writing-mode='tb-rl'">
              <xsl:attribute name ="vert">
                <xsl:value-of select ="'vert'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </xsl:for-each>
      </xsl:for-each>
      <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$parentName]/style:graphic-properties">
        <xsl:call-template name ="tmpInternalPadding"/>


        <!--Added by Mathi for TextAlign in Textbox on 6th Aug 2007-->
        <xsl:for-each select ="draw:frame">
          <xsl:variable name="GrId">
            <xsl:value-of select ="@draw:style-name"/>
          </xsl:variable>
          <xsl:for-each select ="document('styles.xml')//office:automatic-styles/style:style[@style:name=$GrId]/style:graphic-properties">
            <!--<xsl:for-each select ="document($fileName)/child::node()[1]/office:automatic-styles/style:style[@style:name=$gr]/style:graphic-properties">-->
            <!-- Vertical alignment -->
            <xsl:if test="(./parent::node()/style:paragraph-properties/@style:writing-mode='tb-rl') or not(./parent::node()/style:paragraph-properties/@style:writing-mode='tb-rl')">
              <xsl:choose>
                <xsl:when test ="(@draw:textarea-vertical-align='top' and @draw:textarea-horizontal-align='left')">
                  <xsl:attribute name ="anchor">
                    <xsl:value-of select ="'b'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="anchorCtr">
                    <xsl:value-of select ="'0'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="(@draw:textarea-vertical-align='top' and @draw:textarea-horizontal-align='right')">
                  <xsl:attribute name ="anchor">
                    <xsl:value-of select ="'t'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="anchorCtr">
                    <xsl:value-of select ="'0'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="(@draw:textarea-vertical-align='top' and @draw:textarea-horizontal-align='center')">
                  <xsl:attribute name ="anchor">
                    <xsl:value-of select ="'ctr'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="anchorCtr">
                    <xsl:value-of select ="'0'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="(@draw:textarea-vertical-align='middle' and @draw:textarea-horizontal-align='right')">
                  <xsl:attribute name ="anchor">
                    <xsl:value-of select ="'t'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="anchorCtr">
                    <xsl:value-of select ="'1'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="(@draw:textarea-vertical-align='middle' and @draw:textarea-horizontal-align='center')">
                  <xsl:attribute name ="anchor">
                    <xsl:value-of select ="'ctr'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="anchorCtr">
                    <xsl:value-of select ="'1'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="(@draw:textarea-vertical-align='middle' and @draw:textarea-horizontal-align='left')">
                  <xsl:attribute name ="anchor">
                    <xsl:value-of select ="'b'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="anchorCtr">
                    <xsl:value-of select ="'1'"/>
                  </xsl:attribute>
                </xsl:when>
              </xsl:choose>
            </xsl:if>
            <!-- Horizontal alignment -->
            <xsl:if test="not(./parent::node()/style:paragraph-properties/@style:writing-mode='tb-rl')">
              <xsl:choose>
                <xsl:when test ="(@draw:textarea-vertical-align='middle' and @draw:textarea-horizontal-align='center')">
                  <xsl:attribute name ="anchor">
                    <xsl:value-of select ="'ctr'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="anchorCtr">
                    <xsl:value-of select ="'1'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="(@draw:textarea-vertical-align='bottom' and @draw:textarea-horizontal-align='center')">
                  <xsl:attribute name ="anchor">
                    <xsl:value-of select ="'b'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="anchorCtr">
                    <xsl:value-of select ="'1'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="(@draw:textarea-vertical-align='top' and @draw:textarea-horizontal-align='center')">
                  <xsl:attribute name ="anchor">
                    <xsl:value-of select ="'t'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="anchorCtr">
                    <xsl:value-of select ="'1'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="@draw:textarea-vertical-align='middle'">
                  <xsl:attribute name ="anchor">
                    <xsl:value-of select ="'ctr'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="anchorCtr">
                    <xsl:value-of select ="'0'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="@draw:textarea-vertical-align='bottom'">
                  <xsl:attribute name ="anchor">
                    <xsl:value-of select ="'b'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="anchorCtr">
                    <xsl:value-of select ="'0'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:when test ="@draw:textarea-vertical-align='top'">
                  <xsl:attribute name ="anchor">
                    <xsl:value-of select="'t'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="anchorCtr">
                    <xsl:value-of select="'0'"/>
                  </xsl:attribute>
                </xsl:when>
              </xsl:choose>
            </xsl:if>
          </xsl:for-each>
        </xsl:for-each>

        <xsl:variable name ="anchorValue">
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
          </xsl:choose>
        </xsl:variable>
        <xsl:if test ="$anchorValue != ''">
          <xsl:attribute name ="anchor">
            <xsl:value-of select ="$anchorValue"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:attribute name ="anchorCtr">
          <xsl:choose >
            <xsl:when test ="@draw:textarea-horizontal-align ='center'">
              <xsl:value-of select ="1"/>
            </xsl:when>
            <xsl:when test ="@draw:textarea-horizontal-align='justify'">
              <xsl:value-of select ="0"/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select ="0"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:call-template name="tmpWrap"/>

      </xsl:for-each>
    </a:bodyPr>
  </xsl:template>

  <!--Margin-->
  <xsl:template name ="MarginTemplate">
    <xsl:param name ="level"/>
    <xsl:param name="masterName"/>
    <xsl:variable name="listId">
      <xsl:value-of select="concat($masterName,'-outline1')"/>
    </xsl:variable>
    <xsl:for-each select ="//style:style [@style:name=$listId]">
      <xsl:choose>
        <xsl:when test="style:graphic-properties/text:list-style/child::node()[$level]">
          <xsl:call-template name ="convertToPoints">
            <xsl:with-param name ="unit" select ="'cm'"/>
            <xsl:with-param name ="length">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name ="length" select ="style:graphic-properties/text:list-style/child::node()[$level]/style:list-level-properties/@text:space-before"/>
              </xsl:call-template>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:variable name="Id">
            <xsl:value-of select="concat($masterName,'-outline',$level)"/>
          </xsl:variable>
          <xsl:for-each select ="document('styles.xml')//style:style [@style:name=$Id]">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length">
                <xsl:call-template name="convertUnitsToCm">
                  <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
                </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>
  <!--End-->
  <!--Indent-->
  <xsl:template name ="IndentTemplate">
    <xsl:param name ="level"/>
    <xsl:param name="masterName"/>
    <xsl:variable name="listId">
      <xsl:value-of select="concat($masterName,'-outline1')"/>
    </xsl:variable>
    <xsl:for-each select ="//style:style [@style:name=$listId]">
      <xsl:choose>
        <xsl:when test="style:graphic-properties/text:list-style/child::node()[$level]">
          <xsl:call-template name ="convertToPoints">
            <xsl:with-param name ="unit" select ="'cm'"/>
            <xsl:with-param name ="length">
              <xsl:call-template name="convertUnitsToCm">
                <xsl:with-param name ="length" select ="style:graphic-properties/text:list-style/child::node()[$level]/style:list-level-properties/@text:min-label-width"/>
              </xsl:call-template>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:variable name="Id">
            <xsl:value-of select="concat($masterName,'-outline',$level)"/>
          </xsl:variable>
          <xsl:for-each select ="document('styles.xml')//style:style [@style:name=$Id]">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length">
                <xsl:call-template name="convertUnitsToCm">
                  <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:text-indent"/>
                </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>
  <!--End-->
  <!--Bullets Code-->
  <xsl:template name ="SlideMasterBulletsNumbers">
    <xsl:param name ="level"/>
    <xsl:param name="masterName"/>
    <xsl:variable name="listId">
      <xsl:value-of select="concat($masterName,'-outline1')"/>
    </xsl:variable>
    <xsl:for-each select ="//style:style [@style:name=$listId]">
          <xsl:choose >
        <xsl:when test ="style:graphic-properties/text:list-style/child::node()[$level]/style:text-properties/@fo:color">
              <a:buClr>
                <a:srgbClr>
                  <xsl:attribute name ="val">
                  <xsl:value-of select ="substring-after(style:graphic-properties/text:list-style/child::node()[$level]/style:text-properties/@fo:color,'#')"/>
                  </xsl:attribute>
                </a:srgbClr>
              </a:buClr>
            </xsl:when>
        <xsl:otherwise>
          <a:buClr>
            <a:srgbClr val="000000">
            </a:srgbClr>
          </a:buClr>
        </xsl:otherwise>
      </xsl:choose>
      <xsl:if test ="(style:graphic-properties/text:list-style/child::node()[$level]/style:text-properties[@fo:font-size] and not(substring-before(style:graphic-properties/text:list-style/child::node()[$level]/style:text-properties/@fo:font-size,'%') = 100)) ">
        <xsl:choose >
          <xsl:when test ="style:graphic-properties/text:list-style/child::node()[$level]/style:text-properties/@fo:font-size">
            <a:buSzPct>
              <xsl:attribute name ="val">
                <xsl:value-of select ="format-number(substring-before(style:graphic-properties/text:list-style/child::node()[$level]/style:text-properties/@fo:font-size,'%')*1000,'#.##')"/>
              </xsl:attribute>
            </a:buSzPct>
          </xsl:when>
          <xsl:otherwise>
            <a:buSzTx/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
      <xsl:for-each select="style:graphic-properties/text:list-style/text:list-level-style-bullet[@text:level=$level]/style:text-properties">
        <xsl:if test="position()=1">
          <xsl:if test="@fo:font-family">
            <a:buFont>
              <xsl:attribute name ="typeface">
                <xsl:value-of select="@fo:font-family"/>
              </xsl:attribute>
            </a:buFont>
          </xsl:if>
        </xsl:if>
      </xsl:for-each>
      <!--<xsl:call-template name ="getBulletTypes">
						<xsl:with-param name="character" select="style:graphic-properties/text:list-style/text:list-level-style-bullet[@text:level=$level]/@text:bullet-char"/>
						<xsl:with-param name ="typeFace"/>
					</xsl:call-template>
				</xsl:attribute>-->
      <xsl:choose>
        <xsl:when test="style:graphic-properties/text:list-style/text:list-level-style-bullet[@text:level=$level]">
          <a:buChar>
            <xsl:attribute name ="char">
              <xsl:value-of select="style:graphic-properties/text:list-style/text:list-level-style-bullet[@text:level=$level]/@text:bullet-char"/>
              <!--<xsl:call-template name="bulletChar">
								<xsl:with-param name="character" select="style:graphic-properties/text:list-style/text:list-level-style-bullet[@text:level=$level]/@text:bullet-char"/>
							</xsl:call-template>-->
            </xsl:attribute>
          </a:buChar>
        </xsl:when>
        <xsl:when test="style:graphic-properties/text:list-style/text:list-level-style-number[@text:level =$level]">
          <xsl:choose>
            <xsl:when test="style:graphic-properties/text:list-level-style-number[@text:level =$level][@style:num-format='']">
              <a:buNone/>
            </xsl:when>
            <xsl:otherwise>
          <a:buAutoNum>
            <xsl:attribute name ="type">
              <xsl:call-template name="getNumFormat">
                <xsl:with-param name="format">
                  <xsl:value-of select ="style:graphic-properties/text:list-style/text:list-level-style-number[@text:level =$level]/@style:num-format"/>
                </xsl:with-param>
                <xsl:with-param name ="suff" select ="style:graphic-properties/text:list-style/text:list-level-style-number[@text:level =$level]/@style:num-suffix"/>
                <xsl:with-param name ="prefix" select ="style:graphic-properties/text:list-style/text:list-level-style-number[@text:level =$level]/@style:num-prefix"/>
              </xsl:call-template>
            </xsl:attribute>
            <xsl:if test ="style:graphic-properties/text:list-style/text:list-level-style-number[@text:level =$level]/@text:start-value">
              <xsl:attribute name ="startAt">
                <xsl:value-of select ="style:graphic-properties/text:list-style/text:list-level-style-number[@text:level =$level]/@text:start-value"/>
              </xsl:attribute>
            </xsl:if>
          </a:buAutoNum>
            </xsl:otherwise>
         </xsl:choose>
        </xsl:when>
        <xsl:when test="style:graphic-properties/text:list-style/text:list-level-style-image[@text:level=$level]">
          <xsl:if test ="style:graphic-properties/text:list-style/text:list-level-style-image[@text:level=$level]/@xlink:href">
            <a:buChar>
              <xsl:attribute name ="char">
                <xsl:call-template name="bulletChar">
                  <xsl:with-param name ="character" select ="'.'"/>
                </xsl:call-template>
              </xsl:attribute>
            </a:buChar>
          </xsl:if>
        </xsl:when>
        <xsl:when test="not(style:graphic-properties/text:list-style/text:list-level-style-bullet[@text:level=$level]) and style:paragraph-properties/@text:enable-numbering='true'">
          <a:buChar>
            <xsl:attribute name ="char">
              <xsl:call-template name="bulletChar">
                <xsl:with-param name="character" select="'.'"/>
              </xsl:call-template>
            </xsl:attribute>
          </a:buChar>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="bulletChar">
    <xsl:param name ="character"/>
    <xsl:choose>
      <!--<xsl:when test="$character = '' "></xsl:when>-->
      <xsl:when test="$character = '' ">
        <xsl:value-of select ="'q'"/>
      </xsl:when>
      <!--<xsl:when test="$character = '' "></xsl:when>-->
      <xsl:when test="$character = '☑' ">☑</xsl:when>
      <xsl:when test="$character = '•' ">•</xsl:when>
      <!--<xsl:when test="$character= '●' ">•</xsl:when>-->
      <!-- Added by vijayeta ,Fix for bug 1779341, date:23rd Aug '07-->
      <xsl:when test="$character= '●' ">
        <xsl:value-of select ="''"/>
      </xsl:when >
      <!-- Added by vijayeta ,Fix for bug 1779341, date:23rd Aug '07-->
      <!--<xsl:when test="$character = '➢' ">-->
      <xsl:when test="$character = '' or $character = '➢'">
        <xsl:value-of select ="'Ø'"/>
      </xsl:when>
      <xsl:when test="$character = '' or $character = '✔'">
        <!--<xsl:when test="$character = '✔' ">-->
        <xsl:value-of select ="'ü'"/>
      </xsl:when>
      <!--<xsl:when test="$character = '' ">
        <xsl:value-of select ="'§'"/>
      </xsl:when>-->
      <!--<xsl:when test="$character = '' ">
        <xsl:value-of select ="'§'"/>
      </xsl:when>-->
      <xsl:when test="$character = '○' ">o</xsl:when>
      <xsl:when test="$character = '➔' ">è</xsl:when>
      <xsl:when test="$character = '✗' ">✗</xsl:when>
      <xsl:when test="$character = '–' ">–</xsl:when>
      <xsl:otherwise>
        <!-- warn if Custom Bullet -->
        <xsl:message terminate="no">translation.odf2oox.bulletTypeSlideMasterCustomBullet</xsl:message>
        <xsl:value-of select ="'•'"/>
      </xsl:otherwise>
    </xsl:choose>
    <!--<xsl:choose>
			<xsl:when test="$character = '☑' ">☑</xsl:when>
			<xsl:when test="$character = '•' ">•</xsl:when>
			<xsl:when test="$character= '●' ">•</xsl:when>
			<xsl:when test="$character = '➢' ">
				<xsl:value-of select ="'Ø'"/>
			</xsl:when>
			<xsl:when test="$character = '»'">
				<xsl:value-of select ="'»'"/>
			</xsl:when>
			<xsl:when test="$character = '✔' ">
				<xsl:value-of select ="'ü'"/>
			</xsl:when>
			<xsl:when test="$character = '○' ">o</xsl:when>
			<xsl:when test="$character = '➔' ">è</xsl:when>
			<xsl:when test="$character = '✗' ">✗</xsl:when>
			<xsl:when test="$character = '–' ">–</xsl:when>
      <xsl:otherwise>
                <xsl:message terminate="no">translation.odf2oox.bulletTypeSlideMasterCustomBullet</xsl:message>
        <xsl:value-of select ="'•'"/>
      </xsl:otherwise>
		</xsl:choose>-->

  </xsl:template>

  <xsl:template name ="getBulletTypes">
    <xsl:param name ="character"/>
    <xsl:param name ="typeFace"/>
    <xsl:choose >
      <xsl:when test="$character = '' or $character= '●' or $character= '➢' or $character= '' or $character= '✔' or $character = '' or $character= '■' or $character= '' or $character= '➔'">
        <!--<xsl:when test="$character= '➢' or $character= '■' or $character= '✔' or $character= '➔'">-->
        <!-- or $character= ''  -->
        <xsl:value-of select ="'Wingdings'"/>
      </xsl:when>
      <xsl:when test="$character= 'o'">
        <xsl:value-of select ="'Courier New'"/>
      </xsl:when>
      <xsl:when test="$character= '•'">
        <xsl:value-of select ="'Arial'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select ="''"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--End-->
  <!--Relationships-->
  <xsl:template name ="slideMaster1Rel">
    <xsl:param name ="StartLayoutNo"/>
    <xsl:param name ="ThemeId"/>
    <xsl:param name="slideNo" />
    <xsl:param name ="slideMasterName"/>
    <xsl:param name="pos" />
    <xsl:param name ="vmlPageNo" />
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" >
        <xsl:attribute name ="Target">
          <xsl:value-of select ="concat('../slideLayouts/slideLayout',$StartLayoutNo,'.xml')"/>
        </xsl:attribute>
      </Relationship >

      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" >
        <xsl:attribute name ="Target">
          <xsl:value-of select ="concat('../slideLayouts/slideLayout',$StartLayoutNo +1 ,'.xml')"/>
        </xsl:attribute>
      </Relationship >
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" >
        <xsl:attribute name ="Target">
          <xsl:value-of select ="concat('../slideLayouts/slideLayout',$StartLayoutNo +2 ,'.xml')"/>
        </xsl:attribute>
      </Relationship >

      <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" >
        <xsl:attribute name ="Target">
          <xsl:value-of select ="concat('../slideLayouts/slideLayout',$StartLayoutNo+3,'.xml')"/>
        </xsl:attribute>
      </Relationship >
      <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" >
        <xsl:attribute name ="Target">
          <xsl:value-of select ="concat('../slideLayouts/slideLayout',$StartLayoutNo+4,'.xml')"/>
        </xsl:attribute>
      </Relationship >
      <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" >
        <xsl:attribute name ="Target">
          <xsl:value-of select ="concat('../slideLayouts/slideLayout',$StartLayoutNo+5,'.xml')"/>
        </xsl:attribute>
      </Relationship >
      <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" >
        <xsl:attribute name ="Target">
          <xsl:value-of select ="concat('../slideLayouts/slideLayout',$StartLayoutNo+6,'.xml')"/>
        </xsl:attribute>
      </Relationship >
      <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" >
        <xsl:attribute name ="Target">
          <xsl:value-of select ="concat('../slideLayouts/slideLayout',$StartLayoutNo+7,'.xml')"/>
        </xsl:attribute>
      </Relationship >
      <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" >
        <xsl:attribute name ="Target">
          <xsl:value-of select ="concat('../slideLayouts/slideLayout',$StartLayoutNo+8,'.xml')"/>
        </xsl:attribute>
      </Relationship >
      <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" >
        <xsl:attribute name ="Target">
          <xsl:value-of select ="concat('../slideLayouts/slideLayout',$StartLayoutNo+9,'.xml')"/>
        </xsl:attribute>
      </Relationship >
      <Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" >
        <xsl:attribute name ="Target">
          <xsl:value-of select ="concat('../slideLayouts/slideLayout',$StartLayoutNo+10,'.xml')"/>
        </xsl:attribute>
      </Relationship >
      <Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme">
        <xsl:attribute name ="Target">
          <xsl:value-of select ="concat('../theme/theme',$ThemeId,'.xml')"/>
        </xsl:attribute>
      </Relationship >
      <!-- code for picture feature start-->

      <xsl:for-each select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name=$slideMasterName]">
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
                  <xsl:value-of select="concat('../drawings/vmlDrawing',$vmlPageNo,'.vml')"/>
                </xsl:attribute>
              </Relationship>
            </xsl:otherwise>
          </xsl:choose>

        </xsl:if>
        <xsl:for-each select="node()">
          <xsl:choose>
            <xsl:when test="name()='draw:frame'">
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
                          <xsl:with-param name="slideNo" select="$vmlPageNo"/>
                        </xsl:call-template>
                      </xsl:otherwise>
                    </xsl:choose>
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
                              <xsl:value-of select="concat('sl','slideMaster',$vmlPageNo,'Image',$var_pos)" />
                            </xsl:attribute>
                            <xsl:attribute name="Type">
                              <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'" />
                            </xsl:attribute>
                            <xsl:attribute name="Target">
                              <xsl:value-of select="concat('../media/',substring-after(@xlink:href,'/'))" />
                              <!-- <xsl:value-of select ="concat('../media/',substring-after(@xlink:href,'/'))"/>  -->
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
                          <!--end here-->
                        </xsl:if>
                      </xsl:if>
                    </xsl:for-each>
                  </xsl:when>
                  <xsl:when test="./draw:plugin">
                    <xsl:for-each select="./draw:plugin">
                      <Relationship >
                        <xsl:attribute name ="Id">
                          <xsl:value-of  select ="concat('sl',$slideNo,'Au',$var_pos)"/>
                        </xsl:attribute>
                        <xsl:variable name="wavId">
                          <xsl:value-of  select ="concat('sl',$slideNo,'Au',$var_pos)"/>
                        </xsl:variable>
                        <xsl:attribute name ="Type" >
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
                        <xsl:attribute name ="Target">
                          <xsl:choose>
                            <xsl:when test="@xlink:href[ contains(.,'wav')] or @xlink:href[ contains(.,'WAV')]">
                              <xsl:value-of select="concat('../media/',$wavId,'.wav')"/>
                            </xsl:when>
                            <xsl:when test="@xlink:href[ contains(.,'./')]">
                              <xsl:if test="string-length(substring-after(@xlink:href, '../')) = 0">
                                <xsl:value-of select="/"/>
                              </xsl:if>
                              <xsl:if test="not(string-length(substring-after(@xlink:href, '../')) = 0)">
                                <xsl:value-of select="substring-after(@xlink:href, '../')"/>
                              </xsl:if>
                            </xsl:when>
                            <xsl:when test="not(@xlink:href[ contains(.,'./')])">
                              <xsl:value-of select="concat('file:///',translate(substring-after(@xlink:href,'/'),'/','\'))"/>
                            </xsl:when>
                          </xsl:choose>
                        </xsl:attribute>
                        <xsl:if test="not(@xlink:href[ contains(.,'wav')] or @xlink:href[ contains(.,'WAV')])">
                          <xsl:attribute name="TargetMode">
                            <xsl:value-of select="'External'"/>
                          </xsl:attribute>
                        </xsl:if>
                      </Relationship >
                      <!-- for image link -->
                      <Relationship>
                        <xsl:attribute name ="Id">
                          <xsl:value-of  select ="concat('sl',$slideNo,'Im',position())"/>
                        </xsl:attribute>
                        <xsl:attribute name ="Type" >
                          <xsl:value-of select ="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'"/>
                        </xsl:attribute>
                        <xsl:attribute name ="Target">
                          <xsl:if test="./draw:image/@xlink:href !=''">
                            <xsl:value-of select ="concat('../media/',substring-after(./draw:image/@xlink:href,'/'))"/>
                          </xsl:if>
                          <xsl:if test="not(./draw:image/@xlink:href)">
                            <xsl:value-of select ="concat('../media/','thumbnail.png')"/>
                          </xsl:if>
                          <!--<xsl:if test="@xlink:href">
                <xsl:value-of select ="concat('../image/',substring-after(@xlink:href,'/'))"/>  
              </xsl:if>-->
                        </xsl:attribute>
                      </Relationship >
                    </xsl:for-each>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:variable name="PostionCount">
                      <xsl:value-of select="$var_pos"/>
                    </xsl:variable>
                    <xsl:if test="not(@presentation:style-name) and not(@presentation:class)">
                      <xsl:for-each select="draw:text-box">
                        <xsl:variable name="shapeId">
                          <xsl:value-of select="concat('text-box',$PostionCount)"/>
                        </xsl:variable>
                        <xsl:for-each select="text:p">
                          <xsl:if test="text:a/@xlink:href !=''">
                            <Relationship>
                              <xsl:attribute name="Id">
                                <xsl:value-of select="concat($shapeId,'Link',position())"/>
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
                                      <xsl:value-of select="substring-after(text:a/@xlink:href,'../')"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="TargetMode">
                                      <xsl:value-of select="'External'"/>
                                    </xsl:attribute>
                                  </xsl:if>
                                </xsl:otherwise>
                              </xsl:choose>
                            </Relationship>
                          </xsl:if>
                          <xsl:if test="text:span/text:a/@xlink:href !=''">
                            <Relationship>
                              <xsl:attribute name="Id">
                                <xsl:value-of select="concat($shapeId,'Link',position())"/>
                              </xsl:attribute>
                              <xsl:choose>
                                <xsl:when test="text:span/text:a/@xlink:href[contains(.,'#Slide')]">
                                  <xsl:attribute name="Type">
                                    <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
                                  </xsl:attribute>
                                  <xsl:attribute name="Target">
                                    <xsl:value-of select="concat('slide',substring-after(text:span/text:a/@xlink:href,'Slide '),'.xml')"/>
                                  </xsl:attribute>
                                </xsl:when>
                                <xsl:when test="text:span/text:a/@xlink:href[contains(.,'http://') or contains(.,'mailto:')]">
                                  <xsl:attribute name="Type">
                                    <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                                  </xsl:attribute>
                                  <xsl:attribute name="Target">
                                    <xsl:value-of select="text:span/text:a/@xlink:href"/>
                                  </xsl:attribute>
                                  <xsl:attribute name="TargetMode">
                                    <xsl:value-of select="'External'"/>
                                  </xsl:attribute>
                                </xsl:when>
                                <xsl:otherwise>
                                  <xsl:if test="text:span/text:a/@xlink:href[ contains (.,':') ]">
                                    <xsl:attribute name="Type">
                                      <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="Target">
                                      <xsl:value-of select="concat('file:///',translate(substring-after(text:span/text:a/@xlink:href,'/'),'/','\'))"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="TargetMode">
                                      <xsl:value-of select="'External'"/>
                                    </xsl:attribute>
                                  </xsl:if>
                                  <xsl:if test="not(text:span/text:a/@xlink:href[ contains (.,':') ])">
                                    <xsl:attribute name="Type">
                                      <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="Target">
                                      <xsl:value-of select="substring-after(text:span/text:a/@xlink:href,'../')"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="TargetMode">
                                      <xsl:value-of select="'External'"/>
                                    </xsl:attribute>
                                  </xsl:if>
                                </xsl:otherwise>
                              </xsl:choose>
                            </Relationship>
                          </xsl:if>
                        </xsl:for-each>
                        <xsl:for-each select="text:list">
                          <xsl:variable name="forCount" select="position()" />
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
                          <xsl:if test="string-length($xhrefValue) > 0">
                            <xsl:call-template name="tmpShapeBulletOfficeListnerRel">
                              <xsl:with-param name="shapeId" select="$shapeId"/>
                              <xsl:with-param name="blvl" select="$blvl"/>
                              <xsl:with-param name="xhrefValue" select="$xhrefValue"/>
                              <xsl:with-param name="listItemCount" select="generate-id()"/>
                            </xsl:call-template>
                          </xsl:if>
                        </xsl:for-each>
                      </xsl:for-each>
                    </xsl:if>
                    <xsl:if test="@presentation:class">
                      <xsl:variable name="FrameCount">
                        <xsl:value-of select="concat('Frame',$var_pos)"/>
                      </xsl:variable>
                      <xsl:for-each select="draw:text-box">
                        <xsl:for-each select="text:p">
                          <xsl:if test="text:a/@xlink:href !=''">
                            <Relationship>
                              <xsl:attribute name="Id">
                                <xsl:value-of select="concat('TextHLAtchFileId',$PostionCount,'Link',position())"/>
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
                                      <xsl:value-of select="substring-after(text:a/@xlink:href,'../')"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="TargetMode">
                                      <xsl:value-of select="'External'"/>
                                    </xsl:attribute>
                                  </xsl:if>
                                </xsl:otherwise>
                              </xsl:choose>
                            </Relationship>
                          </xsl:if>
                          <xsl:if test="text:span/text:a/@xlink:href !=''">
                            <Relationship>
                              <xsl:attribute name="Id">
                                <xsl:value-of select="concat('TextHLAtchFileId',$PostionCount,'Link',position())"/>
                              </xsl:attribute>
                              <xsl:choose>
                                <xsl:when test="text:span/text:a/@xlink:href[contains(.,'#Slide')]">
                                  <xsl:attribute name="Type">
                                    <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
                                  </xsl:attribute>
                                  <xsl:attribute name="Target">
                                    <xsl:value-of select="concat('slide',substring-after(text:span/text:a/@xlink:href,'Slide '),'.xml')"/>
                                  </xsl:attribute>
                                </xsl:when>
                                <xsl:when test="text:span/text:a/@xlink:href[contains(.,'http://') or contains(.,'mailto:')]">
                                  <xsl:attribute name="Type">
                                    <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                                  </xsl:attribute>
                                  <xsl:attribute name="Target">
                                    <xsl:value-of select="text:span/text:a/@xlink:href"/>
                                  </xsl:attribute>
                                  <xsl:attribute name="TargetMode">
                                    <xsl:value-of select="'External'"/>
                                  </xsl:attribute>
                                </xsl:when>
                                <xsl:otherwise>
                                  <xsl:if test="text:span/text:a/@xlink:href[ contains (.,':') ]">
                                    <xsl:attribute name="Type">
                                      <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="Target">
                                      <xsl:value-of select="concat('file:///',translate(substring-after(text:span/text:a/@xlink:href,'/'),'/','\'))"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="TargetMode">
                                      <xsl:value-of select="'External'"/>
                                    </xsl:attribute>
                                  </xsl:if>
                                  <xsl:if test="not(text:span/text:a/@xlink:href[ contains (.,':') ])">
                                    <xsl:attribute name="Type">
                                      <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="Target">
                                      <xsl:value-of select="substring-after(text:span/text:a/@xlink:href,'../')"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="TargetMode">
                                      <xsl:value-of select="'External'"/>
                                    </xsl:attribute>
                                  </xsl:if>
                                </xsl:otherwise>
                              </xsl:choose>
                            </Relationship>
                          </xsl:if>
                        </xsl:for-each>
                        <xsl:for-each select="text:list">
                          <xsl:variable name="forCount" select="position()" />
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
                                      <xsl:value-of select="substring-after($xhrefValue,'../')"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="TargetMode">
                                      <xsl:value-of select="'External'"/>
                                    </xsl:attribute>
                                  </xsl:if>
                                </xsl:otherwise>
                              </xsl:choose>
                            </Relationship>
                          </xsl:if>
                        </xsl:for-each>
                      </xsl:for-each>
                    </xsl:if>
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
                          <xsl:if test="office:event-listeners">
                              <xsl:call-template name="tmpOfficeListnerRelationship">
                                <xsl:with-param name="ShapeType" select="'TxtBoxAtchFileId'"/>
                                <xsl:with-param name="PostionCount" select="generate-id()"/>
                                <xsl:with-param name="Type" select="'FRAME'"/>
                                    </xsl:call-template>
                                    </xsl:if>
                        </xsl:if>
                                  </xsl:for-each>
                      <xsl:for-each select="./@presentation:class">
                        <xsl:if test="parent::node()/office:event-listeners">
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
                      <xsl:with-param name ="UniqueId">
                        <xsl:choose>
                          <xsl:when test="@presentation:class='date-time'">
                            <xsl:value-of select="'DateTime'"/>
                          </xsl:when>
                          <xsl:when test="@presentation:class='page-number' or @presentation:class='footer' or  @presentation:class='header'">
                            <xsl:value-of select="@presentation:class"/>
                          </xsl:when>
                          <xsl:when test="@presentation:class='title' or @presentation:class='subtitle' or @presentation:class='outline'">
                            <xsl:value-of select="$var_pos"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="$var_pos"/>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:with-param>
                      <xsl:with-param name ="FileName" select="'styles.xml'" />
                      <xsl:with-param name ="prefix" select="'bitmap'" />
                    </xsl:call-template>

                  </xsl:otherwise>
                </xsl:choose>
              </xsl:for-each>
            </xsl:when>
            <xsl:when test="name()='draw:custom-shape'">
              <xsl:variable name="var_pos">
                <xsl:call-template name="getShapePosTemp">
                  <xsl:with-param name="var_pos" select="position()"/>
                </xsl:call-template>
              </xsl:variable>
              <xsl:for-each select=".">
                <xsl:call-template name="tmpBitmapFillRel">
                  <xsl:with-param name ="UniqueId" select="$var_pos" />
                  <xsl:with-param name ="FileName" select="'styles.xml'" />
                  <xsl:with-param name ="prefix" select="'bitmap'" />
                </xsl:call-template>
                <xsl:choose>
                  <xsl:when test="./office:event-listeners">
                    <xsl:variable name="ShapePostionCount">
                      <xsl:value-of select="$var_pos"/>
                    </xsl:variable>
                    <xsl:call-template name="tmpOfficeListnerRelationship">
                      <xsl:with-param name="ShapeType" select="'ShapeFileId'"/>
                      <xsl:with-param name="PostionCount" select="generate-id()"/>
                      <xsl:with-param name="Type" select="'CUSTOM'"/>
                                </xsl:call-template>
                                </xsl:when>
                  <xsl:otherwise>
                    <xsl:variable name="ShapePostionCount">
                      <xsl:value-of select="position()"/>
                    </xsl:variable>
                    <xsl:variable name="shapeId">
                      <xsl:value-of select="concat('custom-shape',$var_pos)"/>
                    </xsl:variable>
                    <xsl:for-each select="text:p">
                      <xsl:if test="text:a/@xlink:href !=''">
                        <Relationship>
                          <xsl:attribute name="Id">
                            <xsl:value-of select="concat($shapeId,'Link',position())"/>
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
                                  <xsl:value-of select="substring-after(text:a/@xlink:href,'../')"/>
                                </xsl:attribute>
                                <xsl:attribute name="TargetMode">
                                  <xsl:value-of select="'External'"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:otherwise>
                          </xsl:choose>
                        </Relationship>
                      </xsl:if>
                      <xsl:if test="text:span/text:a/@xlink:href !=''">
                        <Relationship>
                          <xsl:attribute name="Id">
                            <xsl:value-of select="concat($shapeId,'Link',position())"/>
                          </xsl:attribute>
                          <xsl:choose>
                            <xsl:when test="text:span/text:a/@xlink:href[contains(.,'#Slide')]">
                              <xsl:attribute name="Type">
                                <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
                              </xsl:attribute>
                              <xsl:attribute name="Target">
                                <xsl:value-of select="concat('slide',substring-after(text:span/text:a/@xlink:href,'Slide '),'.xml')"/>
                              </xsl:attribute>
                            </xsl:when>
                            <xsl:when test="text:span/text:a/@xlink:href[contains(.,'http://') or contains(.,'mailto:')]">
                              <xsl:attribute name="Type">
                                <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                              </xsl:attribute>
                              <xsl:attribute name="Target">
                                <xsl:value-of select="text:span/text:a/@xlink:href"/>
                              </xsl:attribute>
                              <xsl:attribute name="TargetMode">
                                <xsl:value-of select="'External'"/>
                              </xsl:attribute>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:if test="text:span/text:a/@xlink:href[ contains (.,':') ]">
                                <xsl:attribute name="Type">
                                  <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                                </xsl:attribute>
                                <xsl:attribute name="Target">
                                  <xsl:value-of select="concat('file:///',translate(substring-after(text:span/text:a/@xlink:href,'/'),'/','\'))"/>
                                </xsl:attribute>
                                <xsl:attribute name="TargetMode">
                                  <xsl:value-of select="'External'"/>
                                </xsl:attribute>
                              </xsl:if>
                              <xsl:if test="not(text:span/text:a/@xlink:href[ contains (.,':') ])">
                                <xsl:attribute name="Type">
                                  <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                                </xsl:attribute>
                                <xsl:attribute name="Target">
                                  <xsl:value-of select="substring-after(text:span/text:a/@xlink:href,'../')"/>
                                </xsl:attribute>
                                <xsl:attribute name="TargetMode">
                                  <xsl:value-of select="'External'"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:otherwise>
                          </xsl:choose>
                        </Relationship>
                      </xsl:if>
                    </xsl:for-each>
                    <xsl:for-each select="text:list">
                      <xsl:variable name="forCount" select="position()" />
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
                      <xsl:if test="string-length($xhrefValue) > 0">


                      </xsl:if>
                    </xsl:for-each>

                  </xsl:otherwise>
                </xsl:choose>
              </xsl:for-each>
            </xsl:when>
            <xsl:when test="name()='draw:rect'">
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
                  <xsl:otherwise>
                    <xsl:variable name="shapeId">
                      <xsl:value-of select="concat('rect',$var_pos)"/>
                    </xsl:variable>
                    <xsl:for-each select="text:p">
                      <xsl:if test="text:a/@xlink:href !=''">
                        <Relationship>
                          <xsl:attribute name="Id">
                            <xsl:value-of select="concat($shapeId,'Link',position())"/>
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
                                  <xsl:value-of select="substring-after(text:a/@xlink:href,'../')"/>
                                </xsl:attribute>
                                <xsl:attribute name="TargetMode">
                                  <xsl:value-of select="'External'"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:otherwise>
                          </xsl:choose>
                        </Relationship>
                      </xsl:if>
                      <xsl:if test="text:span/text:a/@xlink:href !=''">
                        <Relationship>
                          <xsl:attribute name="Id">
                            <xsl:value-of select="concat($shapeId,'Link',position())"/>
                          </xsl:attribute>
                          <xsl:choose>
                            <xsl:when test="text:span/text:a/@xlink:href[contains(.,'#Slide')]">
                              <xsl:attribute name="Type">
                                <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
                              </xsl:attribute>
                              <xsl:attribute name="Target">
                                <xsl:value-of select="concat('slide',substring-after(text:span/text:a/@xlink:href,'Slide '),'.xml')"/>
                              </xsl:attribute>
                            </xsl:when>
                            <xsl:when test="text:span/text:a/@xlink:href[contains(.,'http://') or contains(.,'mailto:')]">
                              <xsl:attribute name="Type">
                                <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                              </xsl:attribute>
                              <xsl:attribute name="Target">
                                <xsl:value-of select="text:span/text:a/@xlink:href"/>
                              </xsl:attribute>
                              <xsl:attribute name="TargetMode">
                                <xsl:value-of select="'External'"/>
                              </xsl:attribute>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:if test="text:span/text:a/@xlink:href[ contains (.,':') ]">
                                <xsl:attribute name="Type">
                                  <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                                </xsl:attribute>
                                <xsl:attribute name="Target">
                                  <xsl:value-of select="concat('file:///',translate(substring-after(text:span/text:a/@xlink:href,'/'),'/','\'))"/>
                                </xsl:attribute>
                                <xsl:attribute name="TargetMode">
                                  <xsl:value-of select="'External'"/>
                                </xsl:attribute>
                              </xsl:if>
                              <xsl:if test="not(text:span/text:a/@xlink:href[ contains (.,':') ])">
                                <xsl:attribute name="Type">
                                  <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                                </xsl:attribute>
                                <xsl:attribute name="Target">
                                  <xsl:value-of select="substring-after(text:span/text:a/@xlink:href,'../')"/>
                                </xsl:attribute>
                                <xsl:attribute name="TargetMode">
                                  <xsl:value-of select="'External'"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:otherwise>
                          </xsl:choose>
                        </Relationship>
                      </xsl:if>
                    </xsl:for-each>
                    <xsl:for-each select="text:list">
                      <xsl:variable name="forCount" select="position()" />
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
                      <xsl:if test="string-length($xhrefValue) > 0">
                        <xsl:call-template name="tmpShapeBulletOfficeListnerRel">
                          <xsl:with-param name="shapeId" select="$shapeId"/>
                          <xsl:with-param name="blvl" select="$blvl"/>
                          <xsl:with-param name="xhrefValue" select="$xhrefValue"/>
                          <xsl:with-param name="listItemCount" select="generate-id()"/>
                        </xsl:call-template>
                      </xsl:if>
                    </xsl:for-each>

                    <xsl:call-template name="tmpBitmapFillRel">
                      <xsl:with-param name ="UniqueId" select="$var_pos" />
                      <xsl:with-param name ="FileName" select="'styles.xml'" />
                      <xsl:with-param name ="prefix" select="'bitmap'" />
                    </xsl:call-template>

                  </xsl:otherwise>
                </xsl:choose>
              </xsl:for-each>
            </xsl:when>
            <xsl:when test="name()='draw:ellipse'">
              <xsl:variable name="var_pos">
                <xsl:call-template name="getShapePosTemp">
                  <xsl:with-param name="var_pos" select="position()"/>
                </xsl:call-template>
              </xsl:variable>
              <xsl:for-each select=".">
                <xsl:variable name="shapeId">
                  <xsl:value-of select="concat('ellipse',$var_pos)"/>
                </xsl:variable>
                <xsl:for-each select="text:p">
                  <xsl:if test="text:a/@xlink:href !=''">
                    <Relationship>
                      <xsl:attribute name="Id">
                        <xsl:value-of select="concat($shapeId,'Link',position())"/>
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
                              <xsl:value-of select="substring-after(text:a/@xlink:href,'../')"/>
                            </xsl:attribute>
                            <xsl:attribute name="TargetMode">
                              <xsl:value-of select="'External'"/>
                            </xsl:attribute>
                          </xsl:if>
                        </xsl:otherwise>
                      </xsl:choose>
                    </Relationship>
                  </xsl:if>
                  <xsl:if test="text:span/text:a/@xlink:href !=''">
                    <Relationship>
                      <xsl:attribute name="Id">
                        <xsl:value-of select="concat($shapeId,'Link',position())"/>
                      </xsl:attribute>
                      <xsl:choose>
                        <xsl:when test="text:span/text:a/@xlink:href[contains(.,'#Slide')]">
                          <xsl:attribute name="Type">
                            <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
                          </xsl:attribute>
                          <xsl:attribute name="Target">
                            <xsl:value-of select="concat('slide',substring-after(text:span/text:a/@xlink:href,'Slide '),'.xml')"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:when test="text:span/text:a/@xlink:href[contains(.,'http://') or contains(.,'mailto:')]">
                          <xsl:attribute name="Type">
                            <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                          </xsl:attribute>
                          <xsl:attribute name="Target">
                            <xsl:value-of select="text:span/text:a/@xlink:href"/>
                          </xsl:attribute>
                          <xsl:attribute name="TargetMode">
                            <xsl:value-of select="'External'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:if test="text:span/text:a/@xlink:href[ contains (.,':') ]">
                            <xsl:attribute name="Type">
                              <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                            </xsl:attribute>
                            <xsl:attribute name="Target">
                              <xsl:value-of select="concat('file:///',translate(substring-after(text:span/text:a/@xlink:href,'/'),'/','\'))"/>
                            </xsl:attribute>
                            <xsl:attribute name="TargetMode">
                              <xsl:value-of select="'External'"/>
                            </xsl:attribute>
                          </xsl:if>
                          <xsl:if test="not(text:span/text:a/@xlink:href[ contains (.,':') ])">
                            <xsl:attribute name="Type">
                              <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                            </xsl:attribute>
                            <xsl:attribute name="Target">
                              <xsl:value-of select="substring-after(text:span/text:a/@xlink:href,'../')"/>
                            </xsl:attribute>
                            <xsl:attribute name="TargetMode">
                              <xsl:value-of select="'External'"/>
                            </xsl:attribute>
                          </xsl:if>
                        </xsl:otherwise>
                      </xsl:choose>
                    </Relationship>
                  </xsl:if>
                </xsl:for-each>
                <xsl:for-each select="text:list">
                  <xsl:variable name="forCount" select="position()" />
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
                  <xsl:if test="string-length($xhrefValue) > 0">

                    <xsl:call-template name="tmpShapeBulletOfficeListnerRel">
                      <xsl:with-param name="shapeId" select="$shapeId"/>
                      <xsl:with-param name="blvl" select="$blvl"/>
                      <xsl:with-param name="xhrefValue" select="$xhrefValue"/>
                      <xsl:with-param name="listItemCount" select="generate-id()"/>
                    </xsl:call-template>
                  </xsl:if>
                </xsl:for-each>

                <xsl:call-template name="tmpBitmapFillRel">
                  <xsl:with-param name ="UniqueId" select="$var_pos" />
                  <xsl:with-param name ="FileName" select="'styles.xml'" />
                  <xsl:with-param name ="prefix" select="'bitmap'" />
                </xsl:call-template>

              </xsl:for-each>
            </xsl:when>
            <xsl:when test="name()='draw:line'">
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
                <xsl:with-param name="slideNo" select="concat('slideMaster',$vmlPageNo)"/>
                <xsl:with-param name="pos" select="$var_pos"/>
                <xsl:with-param name="startPos" select="'1'"/>
                <xsl:with-param name="FileName" select="'styles.xml'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:for-each>
      </xsl:for-each>
      <!-- added by Vipul to set Relationship for background Image  Start-->
      <xsl:variable name="dpName">
        <xsl:value-of select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name=$slideMasterName]/@draw:style-name"/>
      </xsl:variable>
      <xsl:for-each select="document('styles.xml')/office:document-styles/office:automatic-styles/style:style[@style:name= $dpName]/style:drawing-page-properties">
        <xsl:if test="@draw:fill='bitmap'">
          <xsl:variable name="var_imageName" select="@draw:fill-image-name"/>
          <xsl:for-each select="./parent::node()/parent::node()/parent::node()/office:styles/draw:fill-image[@draw:name=$var_imageName]">
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
                  <xsl:value-of select="concat('slideMaster',$vmlPageNo,'BackImg')" />
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
          </xsl:for-each>
        </xsl:if>

      </xsl:for-each>

      <!-- added by Vipul to set Relationship for background Image  End-->
    </Relationships>
  </xsl:template >
  <!-- Superscript and SubScript for Text modified by Mathi on 23rd Aug 2007-->
  <xsl:template name="SuperAndSubscripts">
    <xsl:if test="./style:text-properties/@style:text-position">
      <xsl:choose>
        <xsl:when test="(substring-before(./style:text-properties/@style:text-position,' ') = 'super')">
          <xsl:attribute name="baseline">
            <xsl:variable name="blsuper">
              <xsl:value-of select="substring-before(substring-after(./style:text-properties/@style:text-position,'super '),'%')"/>
            </xsl:variable>
            <xsl:value-of select="($blsuper * 1000)"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="(substring-before(./style:text-properties/@style:text-position,' ') = 'sub')">
          <xsl:attribute name="baseline">
            <xsl:variable name="blsub">
              <xsl:value-of select="substring-before(substring-after(./style:text-properties/@style:text-position,'sub '),'%')"/>
            </xsl:variable>
            <xsl:value-of select="($blsub * (-1000))"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="contains(substring-before(./style:text-properties/@style:text-position,' '),'%') = '%'">
          <xsl:variable name="val">
            <xsl:value-of select="substring-before(substring-before(./style:text-properties/@style:text-position,' '),'%')"/>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test="($val &gt; 0)">
              <xsl:attribute name="baseline">
                <xsl:variable name="bspsb">
                  <xsl:value-of select="substring-before(substring-after(./style:text-properties/@style:text-position,' '),'%')"/>
                </xsl:variable>
                <xsl:value-of select="($bspsb * 1000)"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="($val &lt; 0)">
              <xsl:attribute name="baseline">
                <xsl:variable name="bspsb">
                  <xsl:value-of select="substring-before(substring-after(./style:text-properties/@style:text-position,' '),'%')"/>
                </xsl:variable>
                <xsl:value-of select="($bspsb * (-1000))"/>
              </xsl:attribute>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
  </xsl:template>

</xsl:stylesheet>

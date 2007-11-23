<?xml version="1.0" encoding="utf-8" standalone="yes" ?>
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
  
  <xsl:template name ="handOutMasters">
    <xsl:param name="handOutMasterName"/>
    <xsl:param name ="hoId" />
    <xsl:param name ="headerName"/>
    <xsl:param name ="footerName"/>
    <xsl:param name ="dateTimeName"/>
    <!-- warn no resizing thumbnails representing slides -->
    <xsl:message terminate="no">translation.odf2oox.handOutMasterTypeThumbNail</xsl:message>
    <p:handoutMaster
     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
     xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:bg>
          <p:bgPr>
            <a:solidFill>
              <a:srgbClr>
                <xsl:variable name="dpName">
                  <xsl:value-of select="document('styles.xml')/office:document-styles/office:master-styles/style:handout-master[@style:name=$handOutMasterName]/@draw:style-name"/>
                </xsl:variable>
                <xsl:choose>
                  <xsl:when test="document('styles.xml')//style:style[@style:name= $dpName]/style:drawing-page-properties/@draw:fill='solid'">
                    <xsl:attribute name="val">
                      <xsl:value-of select="substring-after(document('styles.xml')//style:style[@style:name= $dpName]/style:drawing-page-properties/@draw:fill-color,'#')" />
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:if test="not(@draw:fill-color)">
                      <xsl:attribute name="val">
                        <xsl:value-of select="'ffffff'" />
                      </xsl:attribute>
                    </xsl:if>
                  </xsl:otherwise>
                </xsl:choose>
              </a:srgbClr>
            </a:solidFill>
            <a:effectLst/>
          </p:bgPr>
        </p:bg>
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
          <xsl:for-each select="document('styles.xml')/office:document-styles/office:master-styles/style:handout-master/draw:frame">
            <xsl:choose >
              <xsl:when test ="./@presentation:class ='header'">
                <xsl:variable name="dpName">
                  <xsl:value-of select="./parent::node()/@draw:style-name"/>
                </xsl:variable>
                <xsl:variable name ="headerText" >
                  <xsl:if test ="not(./draw:text-box/text:p/presentation:header) or not (./draw:text-box/text:p/text:span/presentation:header)">
                    <xsl:for-each select ="./draw:text-box/text:p">
                      <xsl:if test ="./text:span">
                        <xsl:value-of select ="./text:span"/>
                      </xsl:if>
                      <xsl:if test ="not(./text:span)">
                        <xsl:value-of select ="."/>
                      </xsl:if>
                    </xsl:for-each>
                  </xsl:if>
                  <xsl:if test ="./draw:text-box/text:p/presentation:header or ./draw:text-box/text:p/text:span/presentation:header">
                    <xsl:for-each select ="document('content.xml')//office:presentation/presentation:header-decl[@presentation:name=$headerName]">
                      <xsl:value-of select ="."/>
                    </xsl:for-each>
                  </xsl:if>
                </xsl:variable>
                <xsl:if test="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name= $dpName]/style:drawing-page-properties[@presentation:display-header='true'] or not(./parent::node()/parent::node()/office:automatic-styles/style:style[@style:name= $dpName]/style:drawing-page-properties/@presentation:display-header)">
                  <p:sp>
                    <p:nvSpPr>
                      <p:cNvPr name="Header Placeholder 1">
                        <xsl:attribute name="id">
                          <xsl:value-of select="position()+1"/>
                        </xsl:attribute>
                      </p:cNvPr>
                      <!-- End-->
                      <p:cNvSpPr>
                        <a:spLocks noGrp="1"/>
                      </p:cNvSpPr>
                      <p:nvPr>
                        <p:ph type="hdr" sz="quarter" />
                      </p:nvPr>
                    </p:nvSpPr>
                    <p:spPr>
                      <a:xfrm>
                        <a:off x="304800" y="304800" />
                        <a:ext cx="3368675" cy="503238" />
                      </a:xfrm>
                      <!--<a:xfrm>
                        <xsl:call-template name ="writeCo-ordinates"/>
                      </a:xfrm>-->
                      <a:prstGeom prst="rect">
                        <a:avLst/>
                      </a:prstGeom>
                      <!-- Solid fill color -->
                      <!--<xsl:variable name="prId">
                        <xsl:value-of select="@presentation:style-name"/>
                      </xsl:variable>-->
                      <xsl:variable name="styleName">
                        <xsl:value-of select="@draw:style-name"/>
                      </xsl:variable>
                      <xsl:if test="document('styles.xml')//style:style[@style:name=$styleName]/style:graphic-properties">
                        <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$styleName]/style:graphic-properties ">
                          <xsl:call-template name="tmpSMShapeFillColor"/>
                              </xsl:for-each>
                                           </xsl:if>
                    </p:spPr>
                    <p:txBody>
                      <xsl:call-template name ="handOutTextAndAlignment" >
                        <xsl:with-param name ="Id" select ="@draw:text-style-name"/>
                        <xsl:with-param name ="headerFooterText" select ="$headerText"/>
                      </xsl:call-template >
                    </p:txBody>
                  </p:sp>
                </xsl:if>
              </xsl:when>
              <xsl:when test="./@presentation:class ='date-time'">
                <xsl:variable name="dpName">
                  <xsl:value-of select="./parent::node()/@draw:style-name"/>
                </xsl:variable>
                <xsl:if test="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name= $dpName]/style:drawing-page-properties[@presentation:display-date-time='true'] or not(./parent::node()/parent::node()/office:automatic-styles/style:style[@style:name= $dpName]/style:drawing-page-properties/@presentation:display-date-time)">
                  <p:sp>
                    <p:nvSpPr>
                      <p:cNvPr id="3" name="Date Placeholder 2" />
                      <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                      </p:cNvSpPr>
                      <p:nvPr>
                        <p:ph type="dt" sz="quarter" idx="1" />
                      </p:nvPr>
                    </p:nvSpPr>
                    <p:spPr>
                      <a:xfrm>
                        <a:off x="4114800" y="304800" />
                        <a:ext cx="3368675" cy="503238" />
                      </a:xfrm>
                      <!--<a:xfrm>
                        <xsl:call-template name ="writeCo-ordinates"/>
                      </a:xfrm>-->
                      <a:prstGeom prst="rect">
                        <a:avLst/>
                      </a:prstGeom>
                      <xsl:variable name="styleName">
                        <xsl:value-of select="@draw:style-name"/>
                      </xsl:variable>
                      <xsl:if test="document('styles.xml')//style:style[@style:name=$styleName]/style:graphic-properties">
                        <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$styleName]/style:graphic-properties ">
                          <xsl:call-template name="tmpSMShapeFillColor"/>
                              </xsl:for-each>
                                    </xsl:if>
                    </p:spPr>
                    <p:txBody>
                      <xsl:variable name ="dateId">
                        <xsl:value-of select ="./parent::node()/@presentation:use-date-time-name"/>
                      </xsl:variable>
                      <xsl:call-template name ="handOutDatetime" >
                        <xsl:with-param name ="Id" select ="@draw:text-style-name"/>
                        <xsl:with-param name ="dateId" select ="$dateId"/>
                      </xsl:call-template >
                    </p:txBody>
                  </p:sp>
                </xsl:if>
              </xsl:when>
              <xsl:when test="./@presentation:class ='footer'">
                <xsl:variable name="dpName">
                  <xsl:value-of select="./parent::node()/@draw:style-name"/>
                </xsl:variable>
                <xsl:variable name ="footerText" >
                  <xsl:if test ="not(./draw:text-box/text:p/presentation:footer) or not (./draw:text-box/text:p/text:span/presentation:footer)">
                    <xsl:for-each select ="./draw:text-box/text:p">
                      <xsl:if test ="./text:span">
                        <xsl:value-of select ="./text:span"/>
                      </xsl:if>
                      <xsl:if test ="not(./text:span)">
                        <xsl:value-of select ="."/>
                      </xsl:if>
                    </xsl:for-each>
                  </xsl:if>
                  <xsl:if test ="./draw:text-box/text:p/presentation:footer or ./draw:text-box/text:p/text:span/presentation:footer">
                    <xsl:for-each select ="document('content.xml')//office:presentation/presentation:footer-decl[@presentation:name=$footerName]">
                      <xsl:value-of select ="."/>
                    </xsl:for-each>
                  </xsl:if>
                </xsl:variable>
                <xsl:if test="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name= $dpName]/style:drawing-page-properties[@presentation:display-footer='true'] or not(./parent::node()/parent::node()/office:automatic-styles/style:style[@style:name= $dpName]/style:drawing-page-properties/@presentation:display-footer)">
                  <p:sp>
                    <p:nvSpPr>
                      <p:cNvPr id="4" name="Footer Placeholder 3" />
                      <!-- End-->
                      <p:cNvSpPr>
                        <a:spLocks noGrp="1"/>
                      </p:cNvSpPr>
                      <p:nvPr>
                        <p:ph type="ftr" sz="quarter" idx="2" />
                      </p:nvPr>
                    </p:nvSpPr>
                    <p:spPr>
                      <a:xfrm>
                        <a:off x="228600" y="9555162" />
                        <a:ext cx="3368675" cy="503238" />
                      </a:xfrm>
                      <!--<a:xfrm>
                        <xsl:call-template name ="writeCo-ordinates"/>
                      </a:xfrm>-->
                      <a:prstGeom prst="rect">
                        <a:avLst/>
                      </a:prstGeom>
                      <!-- Solid fill color -->
                      <xsl:variable name="styleName">
                        <xsl:value-of select="@draw:style-name"/>
                      </xsl:variable>
                      <xsl:if test="document('styles.xml')//style:style[@style:name=$styleName]/style:graphic-properties">
                        <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$styleName]/style:graphic-properties ">
                          <xsl:call-template name="tmpSMShapeFillColor"/>
                              </xsl:for-each>
                                        </xsl:if>
                    </p:spPr>
                    <p:txBody>
                      <xsl:call-template name ="handOutTextAndAlignment" >
                        <xsl:with-param name ="Id" select ="@draw:text-style-name"/>
                        <xsl:with-param name ="headerFooterText" select ="$footerText"/>
                      </xsl:call-template >
                    </p:txBody>
                  </p:sp>
                </xsl:if>
              </xsl:when>
              <xsl:when test="./@presentation:class ='page-number'">
                <xsl:variable name="dpName">
                  <xsl:value-of select="./parent::node()/@draw:style-name"/>
                </xsl:variable>
                <xsl:if test="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name= $dpName]/style:drawing-page-properties[@presentation:display-page-number='true'] or not(./parent::node()/parent::node()/office:automatic-styles/style:style[@style:name= $dpName]/style:drawing-page-properties/@presentation:display-page-number)">
                  <p:sp>
                    <p:nvSpPr>
                      <p:cNvPr id="5" name="Slide Number Placeholder 4" />
                      <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                      </p:cNvSpPr>
                      <p:nvPr>
                        <p:ph type="sldNum" sz="quarter" idx="3" />
                      </p:nvPr>
                    </p:nvSpPr>
                    <p:spPr>
                      <a:xfrm>
                        <a:off x="4114800" y="9555162" />
                        <a:ext cx="3368675" cy="503238" />
                      </a:xfrm>
                      <!--<a:xfrm>
                        <xsl:call-template name ="writeCo-ordinates"/>
                      </a:xfrm>-->
                      <a:prstGeom prst="rect">
                        <a:avLst/>
                      </a:prstGeom>
                      <!-- Solid fill color -->
                      <xsl:variable name="styleName">
                        <xsl:value-of select="@draw:style-name"/>
                      </xsl:variable>
                      <xsl:if test="document('styles.xml')//style:style[@style:name=$styleName]/style:graphic-properties">
                        <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$styleName]/style:graphic-properties ">
                          <xsl:call-template name="tmpSMShapeFillColor"/>
                              </xsl:for-each>
                                             </xsl:if>
                    </p:spPr>
                    <p:txBody>
                      <xsl:call-template name ="handoutPagenumber" >
                        <xsl:with-param name ="Id" select ="@draw:text-style-name"/>
                      </xsl:call-template >
                    </p:txBody>
                  </p:sp>
                </xsl:if>
              </xsl:when>
            </xsl:choose>
          </xsl:for-each >
          <!--Code for shapes start-->
          <xsl:call-template name ="shapes" >
            <xsl:with-param name ="fileName" select ="'styles.xml'"/>
          </xsl:call-template >
          <!--Pictures/Images-->
          <!--<xsl:for-each select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name=$slideMasterName]/draw:frame/draw:image">
            <xsl:call-template name ="InsertPicture">
              <xsl:with-param name ="imageNo" select ="1"/>
              <xsl:with-param name ="master" select ="1"/>
            </xsl:call-template>
          </xsl:for-each>-->
          <!--Pictures/Images-->
        </p:spTree>
      </p:cSld>
      <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3"
                accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink" />
    </p:handoutMaster>
  </xsl:template>

  <xsl:template name ="handOutDatetime">
    <xsl:param name ="Id"/>
    <xsl:param name ="dateId"/>
    <xsl:variable name ="datetimeText" >
      <xsl:if test ="not(./draw:text-box/text:p/presentation:date-time or ./draw:text-box/text:p/text:span/presentation:date-time)">
        <xsl:for-each select ="./draw:text-box/text:p">
          <xsl:if test ="./text:span">
            <xsl:value-of select ="./draw:text-box/text:p/text:span"/>
          </xsl:if>
          <xsl:if test ="not(./text:span)">
            <xsl:value-of select ="./draw:text-box/text:p"/>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
      <xsl:if test ="./draw:text-box/text:p/presentation:date-time or ./draw:text-box/text:p/text:span/presentation:date-time">
        <xsl:value-of select ="'0'"/>
      </xsl:if>
    </xsl:variable>
    <a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" >
      <!--<xsl:attribute name ="anchor">
        <xsl:value-of select ="$anchorValue"/>
      </xsl:attribute>-->
    </a:bodyPr>
    <a:lstStyle>
      <a:lvl1pPr>
        <!--<xsl:value-of select="'ctr'"/>-->
        <xsl:if test ="./draw:text-box/text:p">
          <xsl:variable name ="ParId">
            <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
          </xsl:variable>
          <xsl:for-each select ="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$ParId]">
            <xsl:attribute name="algn">
              <xsl:choose >
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
        </xsl:if>
        <a:defRPr>
          <xsl:variable name ="ParId">
            <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
          </xsl:variable>
          <xsl:variable name="textId">
            <xsl:value-of select="./draw:text-box/text:p/text:span/@text:style-name"/>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$textId]/style:text-properties/@fo:font-size">
              <xsl:attribute name="sz">
                <xsl:for-each select ="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$textId]/style:text-properties">
                  <xsl:call-template name ="convertToPoints">
                    <xsl:with-param name ="unit" select ="'pt'"/>
                    <xsl:with-param name ="length" select ="@fo:font-size"/>
                  </xsl:call-template>
                </xsl:for-each>
              </xsl:attribute>
            </xsl:when>
            <!--When draw:text-box/text:p/text:span is not there-->
            <xsl:otherwise>
              <xsl:attribute name="sz">
                <xsl:for-each select ="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$ParId]/style:text-properties">
                  <xsl:call-template name ="convertToPoints">
                    <xsl:with-param name ="unit" select ="'pt'"/>
                    <xsl:with-param name ="length" select ="@fo:font-size"/>
                  </xsl:call-template>
                </xsl:for-each>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
          <!--sateesh-->
          <xsl:for-each select="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$textId]">
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
            <!--Character Spacing-->
            <xsl:if test ="style:text-properties/@fo:letter-spacing [contains(.,'cm')]">
              <xsl:attribute name ="spc">
                <xsl:if test ="substring-before(style:text-properties/@fo:letter-spacing,'cm')&lt; 0 ">
                  <xsl:value-of select ="format-number(substring-before(style:text-properties/@fo:letter-spacing,'cm') * 7200 div 2.54 ,'#')"/>
                </xsl:if >
                <xsl:if test ="substring-before(style:text-properties/@fo:letter-spacing,'cm')
								&gt; 0 or substring-before(style:text-properties/@fo:letter-spacing,'cm') = 0 ">
                  <xsl:value-of select ="format-number((substring-before(style:text-properties/@fo:letter-spacing,'cm') * 72 div 2.54) *100 ,'#')"/>
                </xsl:if>
              </xsl:attribute>
            </xsl:if >
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
                <xsl:when test="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$textId]/style:text-properties/@fo:color">
                  <xsl:attribute name="val">
                    <xsl:value-of select="substring-after(./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$textId]/style:text-properties/@fo:color,'#')"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:attribute name="val">
                    <xsl:value-of select="'000000'"/>
                  </xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
            </a:srgbClr>
          </a:solidFill>
          <xsl:if test="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$textId]/style:text-properties/@fo:text-shadow">
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
                <xsl:when test="translate(./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$ParId]/style:text-properties/@fo:font-family, &quot;'&quot;,'')">
                  <xsl:value-of select="translate(./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$ParId]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:variable name ="graphicStyleName">
                    <xsl:value-of select ="./@draw:style-name"/>
                  </xsl:variable>
                  <xsl:variable name ="parentStyleName">
                    <xsl:value-of select ="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$graphicStyleName]/@style:parent-style-name"/>
                  </xsl:variable>
                  <xsl:variable name ="defFontFamily">
                    <xsl:value-of select ="translate(./parent::node()/parent::node()/parent::node()/office:styles/style:style[@style:name=$parentStyleName]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
                  </xsl:variable>
                  <xsl:value-of select="$defFontFamily"/>
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
    <a:p>
      <a:pPr>
        <xsl:variable name ="ParId">
          <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
        </xsl:variable>
        <xsl:for-each select ="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$ParId]">
          <xsl:call-template name ="tmpHMParaProperties"/>
        </xsl:for-each >
      </a:pPr>
      <xsl:if test ="$datetimeText != '0'">
        <a:r>
          <a:rPr lang="en-US" smtClean="0" />
          <a:t>
            <xsl:value-of select ="$datetimeText"/>
          </a:t>
        </a:r >
      </xsl:if>
      <xsl:if test ="$datetimeText = '0'">
        <xsl:for-each select ="document('content.xml')//office:body/office:presentation/presentation:date-time-decl">
          <xsl:choose>
            <xsl:when test="./@presentation:source='current-date' and (./@presentation:name=$dateId or not($dateId))" >
              <a:fld >
                <xsl:attribute name ="id">
                  <xsl:value-of select ="'{86419996-E19B-43D7-A4AA-D671C2F15715}'"/>
                </xsl:attribute>
                <xsl:attribute name ="type">
                  <xsl:choose >
                    <xsl:when test ="./presentation:date-time-decl/@style:data-style-name ='D3'">
                      <xsl:value-of select ="'datetime1'"/>
                    </xsl:when>
                    <xsl:when test ="./presentation:date-time-decl/@style:data-style-name ='D8'">
                      <xsl:value-of select ="'datetime2'"/>
                    </xsl:when>
                    <xsl:when test ="./presentation:date-time-decl/@style:data-style-name ='D6'">
                      <xsl:value-of select ="'datetime4'"/>
                    </xsl:when>
                    <xsl:when test ="./presentation:date-time-decl/@style:data-style-name ='D5'">
                      <xsl:value-of select ="'datetime4'"/>
                    </xsl:when>
                    <xsl:when test ="./presentation:date-time-decl/@style:data-style-name ='D3T2'">
                      <xsl:value-of select ="'datetime8'"/>
                    </xsl:when>
                    <xsl:when test ="./presentation:date-time-decl/@style:data-style-name ='D3T5'">
                      <xsl:value-of select ="'datetime8'"/>
                    </xsl:when>
                    <xsl:when test ="./presentation:date-time-decl/@style:data-style-name ='T2'">
                      <xsl:value-of select ="'datetime10'"/>
                    </xsl:when>
                    <xsl:when test ="./presentation:date-time-decl/@style:data-style-name ='T3'">
                      <xsl:value-of select ="'datetime11'"/>
                    </xsl:when>
                    <xsl:when test ="./presentation:date-time-decl/@style:data-style-name ='T5'">
                      <xsl:value-of select ="'datetime12'"/>
                    </xsl:when>
                    <xsl:when test ="./presentation:date-time-decl/@style:data-style-name ='T6'">
                      <xsl:value-of select ="'datetime13'"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select ="'datetime1'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
                <a:rPr lang="en-US" smtClean="0"/>
                <a:t>4/5/2007</a:t>
              </a:fld>
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
      </xsl:if>
    </a:p>
  </xsl:template>
  <xsl:template name ="handoutPagenumber">
    <xsl:param name ="Id"/>
    <a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" />
    <a:lstStyle>
      <a:lvl1pPr>
        <xsl:if test ="./draw:text-box/text:p">
          <xsl:variable name ="ParId">
            <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
          </xsl:variable>
          <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$ParId]">
            <xsl:attribute name="algn">
              <xsl:choose >
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
        </xsl:if>
        <a:defRPr>
          <xsl:variable name ="ParId">
            <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
          </xsl:variable>
          <xsl:variable name="textId">
            <xsl:value-of select="./draw:text-box/text:p/text:span/@text:style-name"/>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$textId]/style:text-properties/@fo:font-size">
              <xsl:attribute name="sz">
                <xsl:for-each select ="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$textId]/style:text-properties">
                  <xsl:call-template name ="convertToPoints">
                    <xsl:with-param name ="unit" select ="'pt'"/>
                    <xsl:with-param name ="length" select ="@fo:font-size"/>
                  </xsl:call-template>
                </xsl:for-each>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="sz">
                <xsl:for-each select ="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$ParId]/style:text-properties">
                  <xsl:call-template name ="convertToPoints">
                    <xsl:with-param name ="unit" select ="'pt'"/>
                    <xsl:with-param name ="length" select ="@fo:font-size"/>
                  </xsl:call-template>
                </xsl:for-each>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
          <xsl:for-each select="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$textId]">
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
            <!--Character Spacing-->
            <xsl:if test ="style:text-properties/@fo:letter-spacing [contains(.,'cm')]">
              <xsl:attribute name ="spc">
                <xsl:if test ="substring-before(style:text-properties/@fo:letter-spacing,'cm')&lt; 0 ">
                  <xsl:value-of select ="format-number(substring-before(style:text-properties/@fo:letter-spacing,'cm') * 7200 div 2.54 ,'#')"/>
                </xsl:if >
                <xsl:if test ="substring-before(style:text-properties/@fo:letter-spacing,'cm')
								&gt; 0 or substring-before(style:text-properties/@fo:letter-spacing,'cm') = 0 ">
                  <xsl:value-of select ="format-number((substring-before(style:text-properties/@fo:letter-spacing,'cm') * 72 div 2.54) *100 ,'#')"/>
                </xsl:if>
              </xsl:attribute>
            </xsl:if >
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
              <!--<xsl:variable name="textId">
                <xsl:value-of select="./draw:text-box/text:p/text:span/@text:style-name"/>
              </xsl:variable>-->
              <xsl:choose>
                <xsl:when test="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$textId]/style:text-properties/@fo:color">

                  <xsl:attribute name="val">
                    <xsl:value-of select="substring-after(./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$textId]/style:text-properties/@fo:color,'#')"/>

                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:attribute name="val">
                    <xsl:value-of select="'000000'"/>
                  </xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
            </a:srgbClr>
          </a:solidFill>
          <xsl:if test="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$textId]/style:text-properties/@fo:text-shadow">
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
                <xsl:when test="translate(./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$ParId]/style:text-properties/@fo:font-family, &quot;'&quot;,'')">
                  <xsl:value-of select="translate(./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$ParId]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:variable name ="graphicStyleName">
                    <xsl:value-of select ="./@draw:style-name"/>
                  </xsl:variable>
                  <xsl:variable name ="parentStyleName">
                    <xsl:value-of select ="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$graphicStyleName]/@style:parent-style-name"/>
                  </xsl:variable>
                  <xsl:variable name ="defFontFamily">
                    <xsl:value-of select ="translate(./parent::node()/parent::node()/parent::node()/office:styles/style:style[@style:name=$parentStyleName]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
                  </xsl:variable>
                  <xsl:value-of select="$defFontFamily"/>
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
    <a:p>
      <a:pPr>
        <xsl:variable name ="ParId">
          <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
        </xsl:variable>
        <xsl:for-each select ="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$ParId]">
          <xsl:call-template name ="tmpHMParaProperties"/>
        </xsl:for-each >
      </a:pPr>
      <xsl:choose>
        <xsl:when test="./draw:text-box/text:p/text:page-number" >
          <a:fld >
            <xsl:attribute name ="id">
              <xsl:value-of select ="'{763D1470-AB83-4C4C-B3B3-7F0C9DC8E8D6}'"/>
            </xsl:attribute>
            <xsl:attribute name ="type">
              <xsl:value-of select ="'slidenum'"/>
            </xsl:attribute>
            <a:rPr lang="en-US" smtClean="0" />
            <a:t>
              <xsl:value-of select="."/>
            </a:t>
          </a:fld>
          <a:endParaRPr lang="en-US" />
        </xsl:when>
        <xsl:when test="./draw:text-box/text:p/text:span/text:page-number">
          <a:r>
            <a:rPr lang="en-US" smtClean="0" />
            <a:t>‹#›</a:t>
          </a:r >
        </xsl:when>
        <xsl:otherwise >
          <a:r>
            <a:rPr lang="en-US" smtClean="0" />
            <a:t>
              <xsl:for-each select="./draw:text-box/text:p/text:span">
                <xsl:value-of select="."/>
              </xsl:for-each>
            </a:t>
          </a:r >
        </xsl:otherwise>
      </xsl:choose>
    </a:p>
  </xsl:template>
  <xsl:template name ="handOutTextAndAlignment" >
    <xsl:param name ="Id"/>
    <xsl:param name ="headerFooterText"/>
    <a:bodyPr vert="horz" >
      <!--lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"-->
      <!--<xsl:attribute name ="anchor">
        <xsl:value-of select ="$anchorValue"/>
      </xsl:attribute>-->
    </a:bodyPr>
    <a:lstStyle>
      <a:lvl1pPr>
        <!--<xsl:value-of select="'ctr'"/>-->
        <xsl:if test ="./draw:text-box/text:p">
          <xsl:variable name ="ParId">
            <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
          </xsl:variable>
          <xsl:for-each select ="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$ParId]">
            <xsl:attribute name="algn">
              <xsl:choose >
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
        </xsl:if>
        <a:defRPr>
          <xsl:variable name ="ParId">
            <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
          </xsl:variable>
          <xsl:variable name="textId">
            <xsl:value-of select="./draw:text-box/text:p/text:span/@text:style-name"/>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$textId]/style:text-properties/@fo:font-size">
              <xsl:attribute name="sz">
                <xsl:for-each select ="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$textId]/style:text-properties">
                  <xsl:call-template name ="convertToPoints">
                    <xsl:with-param name ="unit" select ="'pt'"/>
                    <xsl:with-param name ="length" select ="@fo:font-size"/>
                  </xsl:call-template>
                </xsl:for-each>
              </xsl:attribute>
            </xsl:when>
            <!--When draw:text-box/text:p/text:span is not there-->
            <xsl:otherwise>
              <xsl:attribute name="sz">
                <xsl:for-each select ="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$ParId]/style:text-properties">
                  <xsl:call-template name ="convertToPoints">
                    <xsl:with-param name ="unit" select ="'pt'"/>
                    <xsl:with-param name ="length" select ="@fo:font-size"/>
                  </xsl:call-template>
                </xsl:for-each>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
          <!--sateesh-->
          <xsl:for-each select="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$textId]">
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
            <!--Character Spacing-->
            <xsl:if test ="style:text-properties/@fo:letter-spacing [contains(.,'cm')]">
              <xsl:attribute name ="spc">
                <xsl:if test ="substring-before(style:text-properties/@fo:letter-spacing,'cm')&lt; 0 ">
                  <xsl:value-of select ="format-number(substring-before(style:text-properties/@fo:letter-spacing,'cm') * 7200 div 2.54 ,'#')"/>
                </xsl:if >
                <xsl:if test ="substring-before(style:text-properties/@fo:letter-spacing,'cm')
								&gt; 0 or substring-before(style:text-properties/@fo:letter-spacing,'cm') = 0 ">
                  <xsl:value-of select ="format-number((substring-before(style:text-properties/@fo:letter-spacing,'cm') * 72 div 2.54) *100 ,'#')"/>
                </xsl:if>
              </xsl:attribute>
            </xsl:if >
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
                <xsl:when test="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$textId]/style:text-properties/@fo:color">
                  <xsl:attribute name="val">
                    <xsl:value-of select="substring-after(./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$textId]/style:text-properties/@fo:color,'#')"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:attribute name="val">
                    <xsl:value-of select="'000000'"/>
                  </xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
            </a:srgbClr>
          </a:solidFill>
          <xsl:if test="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$textId]/style:text-properties/@fo:text-shadow">
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
                <xsl:when test="translate(./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$ParId]/style:text-properties/@fo:font-family, &quot;'&quot;,'')">
                  <xsl:value-of select="translate(./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$ParId]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:variable name ="graphicStyleName">
                    <xsl:value-of select ="./@draw:style-name"/>
                  </xsl:variable>
                  <xsl:variable name ="parentStyleName">
                    <xsl:value-of select ="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$graphicStyleName]/@style:parent-style-name"/>
                  </xsl:variable>
                  <xsl:variable name ="defFontFamily">
                    <xsl:value-of select ="translate(./parent::node()/parent::node()/parent::node()/office:styles/style:style[@style:name=$parentStyleName]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
                  </xsl:variable>
                  <xsl:value-of select="$defFontFamily"/>
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
    <a:p>
      <a:pPr>
        <xsl:variable name ="ParId">
          <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
        </xsl:variable>
        <xsl:for-each select ="./parent::node()/parent::node()/parent::node()/office:automatic-styles/style:style[@style:name=$ParId]">
          <xsl:call-template name ="tmpHMParaProperties"/>
        </xsl:for-each >
      </a:pPr>
      <a:r>
        <a:rPr lang="en-US" dirty="0" smtClean="0"/>
        <a:t>
          <xsl:value-of select="$headerFooterText"/>
        </a:t>
      </a:r>
      <a:endParaRPr lang="en-US"/>
    </a:p>
  </xsl:template >
  <xsl:template name="Underline">
    <!-- Font underline-->
    <xsl:param name="uStyle"/>
    <xsl:param name="uWidth"/>
    <xsl:param name="uType"/>
    <xsl:choose >
      <xsl:when test="$uStyle='solid' and
								$uType='double'">
        <xsl:value-of select ="'dbl'"/>
      </xsl:when>
      <xsl:when test="$uStyle='solid' and	$uWidth='bold'">
        <xsl:value-of select ="'heavy'"/>
      </xsl:when>
      <xsl:when test="$uStyle='solid' and $uWidth='auto'">
        <xsl:value-of select ="'sng'"/>
      </xsl:when>
      <!-- Dotted lean and dotted bold under line -->
      <xsl:when test="$uStyle='dotted' and	$uWidth='auto'">
        <xsl:value-of select ="'dotted'"/>
      </xsl:when>
      <xsl:when test="$uStyle='dotted' and	$uWidth='bold'">
        <xsl:value-of select ="'dottedHeavy'"/>
      </xsl:when>
      <!-- Dash lean and dash bold underline -->
      <xsl:when test="$uStyle='dash' and
								$uWidth='auto'">
        <xsl:value-of select ="'dash'"/>
      </xsl:when>
      <xsl:when test="$uStyle='dash' and
								$uWidth='bold'">
        <xsl:value-of select ="'dashHeavy'"/>
      </xsl:when>
      <!-- Dash long and dash long bold -->
      <xsl:when test="$uStyle='long-dash' and
								$uWidth='auto'">
        <xsl:value-of select ="'dashLong'"/>
      </xsl:when>
      <xsl:when test="$uStyle='long-dash' and
								$uWidth='bold'">
        <xsl:value-of select ="'dashLongHeavy'"/>
      </xsl:when>
      <!-- dot Dash and dot dash bold -->
      <xsl:when test="$uStyle='dot-dash' and
								$uWidth='auto'">
        <xsl:value-of select ="'dotDashLong'"/>
      </xsl:when>
      <xsl:when test="$uStyle='dot-dash' and
								$uWidth='bold'">
        <xsl:value-of select ="'dotDashHeavy'"/>
      </xsl:when>
      <!-- dot-dot-dash-->
      <xsl:when test="$uStyle='dot-dot-dash' and
								$uWidth='auto'">
        <xsl:value-of select ="'dotDotDash'"/>
      </xsl:when>
      <xsl:when test="$uStyle='dot-dot-dash' and
								$uWidth='bold'">
        <xsl:value-of select ="'dotDotDashHeavy'"/>
      </xsl:when>
      <!-- double Wavy -->
      <xsl:when test="$uStyle='wave' and
								$uType='double'">
        <xsl:value-of select ="'wavyDbl'"/>
      </xsl:when>
      <!-- Wavy and Wavy Heavy-->
      <xsl:when test="$uStyle='wave' and
								$uWidth='auto'">
        <xsl:value-of select ="'wavy'"/>
      </xsl:when>
      <xsl:when test="$uStyle='wave' and
								$uWidth='bold'">
        <xsl:value-of select ="'wavyHeavy'"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!-- Handout paragraph properties-->
  <xsl:template name ="tmpHMParaProperties">
    <!-- Code inserted by Vijayeta for Bullets and numbering,For bullet properties-->
    <!--<xsl:if test ="not($level='0')">
        <xsl:attribute name ="lvl">
          <xsl:value-of select ="$level"/>
        </xsl:attribute>
      </xsl:if>-->

    <!--marL="first line indent property"-->
    <xsl:if test ="style:paragraph-properties/@fo:text-indent 
							and substring-before(style:paragraph-properties/@fo:text-indent,'cm') != 0">
      <xsl:attribute name ="indent">
        <!--fo:text-indent-->
        <xsl:call-template name ="convertToPoints">
          <xsl:with-param name="length"  select ="style:paragraph-properties/@fo:text-indent"/>
          <xsl:with-param name ="unit" select ="'cm'"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if >
    <xsl:if test ="style:paragraph-properties/@fo:text-align">
      <xsl:attribute name ="algn">
        <!--fo:text-align-->
        <xsl:choose >
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
    </xsl:if >
    <xsl:if test ="style:paragraph-properties/@fo:margin-left and 
							   substring-before(style:paragraph-properties/@fo:margin-left,'cm') &gt; 0">
      <xsl:attribute name ="marL">
        <!--fo:margin-left-->
        <xsl:call-template name ="convertToPoints">
          <xsl:with-param name="length"  select ="style:paragraph-properties/@fo:margin-left"/>
          <xsl:with-param name ="unit" select ="'cm'"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if >
    <xsl:if test ="style:paragraph-properties/@fo:line-height and 
					substring-before(style:paragraph-properties/@fo:line-height,'%') &gt; 0 and 
					not(substring-before(style:paragraph-properties/@fo:line-height,'%') = 100)">
      <a:lnSpc>
        <a:spcPct>
          <xsl:attribute name ="val">
            <xsl:value-of select ="format-number(substring-before(style:paragraph-properties/@fo:line-height,'%')* 1000,'#.##') "/>
          </xsl:attribute>
        </a:spcPct>
      </a:lnSpc>
    </xsl:if>
    <xsl:if test ="style:paragraph-properties/@style:line-spacing and 
					substring-before(style:paragraph-properties/@style:line-spacing,'cm') &gt; 0">
      <a:lnSpc>
        <a:spcPts>
          <xsl:attribute name ="val">
            <xsl:value-of select ="round(substring-before(style:paragraph-properties/@style:line-spacing,'cm')* 2835) "/>
          </xsl:attribute>
        </a:spcPts>
      </a:lnSpc>
    </xsl:if>
    <xsl:if test ="style:paragraph-properties/@style:line-height-at-least and 
					substring-before(style:paragraph-properties/@style:line-height-at-least,'cm') &gt; 0 ">
      <a:lnSpc>
        <a:spcPts>
          <xsl:attribute name ="val">
            <xsl:value-of select ="round(substring-before(style:paragraph-properties/@style:line-height-at-least,'cm')* 2835) "/>
          </xsl:attribute>
        </a:spcPts>
      </a:lnSpc>
    </xsl:if>
    <xsl:if test ="style:paragraph-properties/@fo:margin-top and 
						substring-before(style:paragraph-properties/@fo:margin-top,'cm') &gt; 0 ">
      <a:spcBef>
        <a:spcPts>
          <xsl:attribute name ="val">
            <!--fo:margin-top-->
            <xsl:value-of select ="round(substring-before(style:paragraph-properties/@fo:margin-top,'cm')* 2835) "/>
          </xsl:attribute>
        </a:spcPts>
      </a:spcBef >
    </xsl:if>
    <xsl:if test ="style:paragraph-properties/@fo:margin-bottom and 
					    substring-before(style:paragraph-properties/@fo:margin-bottom,'cm') &gt; 0 ">
      <a:spcAft>
        <a:spcPts>
          <xsl:attribute name ="val">
            <!--fo:margin-bottom-->
            <xsl:value-of select ="round(substring-before(style:paragraph-properties/@fo:margin-bottom,'cm')* 2835) "/>
          </xsl:attribute>
        </a:spcPts>
      </a:spcAft>
    </xsl:if >
    <!-- Code Added by Vijayeta,for Paragraph Spacing, Before and After-->
    <!--<xsl:if test ="isBulleted='false'">
				<a:buNone/>
        </xsl:if>-->
    <!--<xsl:if test ="$isNumberingEnabled='false'">
        <a:buNone/>
      </xsl:if>
      <xsl:if test ="$isBulleted='true'">
        <xsl:if test ="$isNumberingEnabled='true'">
          <xsl:call-template name ="insertBulletsNumbers" >
            <xsl:with-param name ="listId" select ="$listId"/>
            <xsl:with-param name ="level" select ="$level+1"/>
            -->
    <!-- parameter added by vijayeta, dated 13-7-07-->
    <!--
            <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
          </xsl:call-template>
        </xsl:if>
      </xsl:if>-->
    <xsl:call-template name ="paragraphTabstops"/>
  </xsl:template>
  <!--Relationships-->
  <xsl:template name ="handoutMaster1Rel">
    <xsl:param name ="ThemeId"/>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme">
        <xsl:attribute name ="Target">
          <xsl:value-of select ="concat('../theme/theme',$ThemeId,'.xml')"/>
        </xsl:attribute>
      </Relationship >
      <!-- code for picture feature start-->
      <!--<xsl:for-each select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name=$slideMasterName]/draw:frame/draw:image">
        <Relationship>
          <xsl:attribute name="Id">
            <xsl:value-of select="concat('sl',$slideNo,'Image',position())" />
          </xsl:attribute>
          <xsl:attribute name="Type">
            <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'" />
          </xsl:attribute>
          <xsl:attribute name="Target">
            <xsl:value-of select="concat('../media/',substring-after(@xlink:href,'/'))" />
            -->
      <!-- <xsl:value-of select ="concat('../media/',substring-after(@xlink:href,'/'))"/>  -->
      <!--
          </xsl:attribute>
        </Relationship>
      </xsl:for-each>-->
      <!-- code for picture feature end -->
    </Relationships>
  </xsl:template >
  <xsl:template name ="writeCo-ordinates">
    <xsl:variable name ="angle">
      <xsl:value-of select="substring-after(substring-before(substring-before(@draw:transform,'translate'),')'),'(')" />
    </xsl:variable>
    <xsl:variable name ="x2">
      <xsl:value-of select="substring-before(substring-before(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
    </xsl:variable>
    <xsl:variable name ="y2">
      <xsl:value-of select="substring-before(substring-after(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
    </xsl:variable>
    <xsl:if test="@draw:transform">
      <xsl:attribute name ="rot">
        <xsl:value-of select ="concat('draw-transform:ROT:',substring-before(@svg:width,'cm'), ':',
																   substring-before(@svg:height,'cm'), ':', 
																   $x2, ':',  $y2, ':', $angle)"/>
      </xsl:attribute>
    </xsl:if>
    <a:off >
      <xsl:if test="not(@draw:transform)">
        <xsl:attribute name ="x">
          <!--<xsl:value-of select ="@svg:x"/>-->
          <xsl:call-template name ="convertToPoints">
            <xsl:with-param name ="unit" select ="'cm'"/>
            <xsl:with-param name ="length" select ="@svg:x"/>
          </xsl:call-template>
        </xsl:attribute>
        <xsl:attribute name ="y">
          <!--<xsl:value-of select ="@svg:y"/>-->
          <xsl:call-template name ="convertToPoints">
            <xsl:with-param name ="unit" select ="'cm'"/>
            <xsl:with-param name ="length" select ="@svg:y"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if >
      <xsl:if test="@draw:transform">
        <xsl:attribute name ="x">
          <xsl:value-of select ="concat('draw-transform:X:',substring-before(@svg:width,'cm'), ':',
																   substring-before(@svg:height,'cm'), ':', 
																   $x2, ':', $y2, ':', $angle)"/>
        </xsl:attribute>
        <xsl:attribute name ="y">
          <xsl:value-of select ="concat('draw-transform:Y:',substring-before(@svg:width,'cm'), ':',
																   substring-before(@svg:height,'cm'), ':', 
																   $x2, ':',$y2, ':', $angle)"/>
        </xsl:attribute>
      </xsl:if>
    </a:off>
    <!--<a:ext cx="7772400" cy="1600200" />-->
    <a:ext>
      <xsl:attribute name ="cx">
        <!--<xsl:value-of select ="@svg:width"/>-->
        <xsl:call-template name ="convertToPoints">
          <xsl:with-param name ="unit" select ="'cm'"/>
          <xsl:with-param name ="length" select ="@svg:width"/>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name ="cy">
        <!--<xsl:value-of select ="@svg:height"/>-->
        <xsl:call-template name ="convertToPoints">
          <xsl:with-param name ="unit" select ="'cm'"/>
          <xsl:with-param name ="length" select ="@svg:height"/>
        </xsl:call-template>
      </xsl:attribute>
    </a:ext>
  </xsl:template>
</xsl:stylesheet>

<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
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
  <xsl:template name ="NotesMasters">
    <!-- warn,notes master -->
    <xsl:message terminate="no">translation.odf2oox.notesMasterMultipleToSingle</xsl:message>
    <p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:bg>
          <p:bgRef idx="1001">
            <a:schemeClr val="bg1"/>
          </p:bgRef>
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
          <xsl:for-each select="presentation:notes">
            <!-- For Converting header -->
            <xsl:for-each select ="draw:frame[@presentation:class='header'] ">
              <xsl:call-template name="tmpNMHeader"/>
            </xsl:for-each>
            <!-- For Converting DateTime -->
            <xsl:for-each select ="draw:frame[@presentation:class='date-time'] ">
              <xsl:call-template name="tmpNMDatetime"/>
            </xsl:for-each>
            <!-- For Converting SldImg -->
            <xsl:for-each select ="draw:page-thumbnail[@presentation:class='page'] ">
              <xsl:call-template name="Page"/>
            </xsl:for-each>
            <!-- for convertion notes-->
            <xsl:for-each select ="draw:frame[@presentation:class='notes'] ">
              <xsl:call-template name="Notes"/>
            </xsl:for-each>
            <!-- For Converting Footer -->
            <xsl:for-each select ="draw:frame[@presentation:class='footer'] ">
              <xsl:call-template name="tmpNMFooter"/>
            </xsl:for-each>
            <!-- For Converting PageNumber -->
            <xsl:for-each select ="draw:frame[@presentation:class='page-number'] ">
              <xsl:call-template name="tmpNMPagenumber"/>
            </xsl:for-each>
          </xsl:for-each>
        </p:spTree>
      </p:cSld>
      <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
      <p:notesStyle>
        <a:lvl1pPr>
          <xsl:variable name="slideMasterName">
            <xsl:value-of select="@style:name"/>
          </xsl:variable>
          <xsl:variable name="outlineName">
            <xsl:value-of select="concat($slideMasterName,'-notes')"/>
          </xsl:variable>
          <xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName]">
            <!--Margin-->
            <xsl:attribute name="marL">
              <xsl:call-template name ="convertToPoints">
                <xsl:with-param name ="unit" select ="'cm'"/>
                <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
              </xsl:call-template>
            </xsl:attribute>
            <!--End-->
            <!--Indent-->
            <xsl:attribute name="indent">
              <xsl:call-template name ="convertToPoints">
                <xsl:with-param name ="unit" select ="'cm'"/>
                <xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:text-indent"/>
              </xsl:call-template>
            </xsl:attribute>
            <!--End-->
            <xsl:attribute name ="algn">
              <!--fo:text-align-->
              <xsl:choose >
                <xsl:when test ="./style:paragraph-properties/@fo:text-align='center'">
                  <xsl:value-of select ="'ctr'"/>
                </xsl:when>
                <xsl:when test ="./style:paragraph-properties/@fo:text-align='end'">
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
            <!--Code for Line Spacing-->
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
            <!--If the line spacing is in terms of Points,multiply the value with 2835-->
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
            <!--End-->
            <!--Space Before and After Paragraph-->
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
            <!--End-->
            <a:defRPr>
              <xsl:attribute name="sz">
                <xsl:call-template name ="convertToPoints">
                  <xsl:with-param name ="unit" select ="'pt'"/>
                  <xsl:with-param name ="length" select ="./style:text-properties/@fo:font-size"/>
                </xsl:call-template>
              </xsl:attribute>
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
              <!-- Font Underline-->
              <xsl:variable name ="unLine">
                <xsl:call-template name="tmpNMUnderline">
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
        <!--<a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1200" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl1pPr>-->
        <a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1200" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl2pPr>
        <a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1200" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl3pPr>
        <a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1200" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl4pPr>
        <a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1200" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl5pPr>
        <a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1200" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl6pPr>
        <a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1200" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl7pPr>
        <a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1200" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl8pPr>
        <a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1200" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl9pPr>
      </p:notesStyle>
    </p:notesMaster>
  </xsl:template>
  <xsl:template name="tmpNMHeader">
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="2" name="Header Placeholder 1"/>
        <p:cNvSpPr>
          <a:spLocks noGrp="1"/>
        </p:cNvSpPr>
        <p:nvPr>
          <p:ph type="hdr" sz="quarter"/>
        </p:nvPr>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <xsl:call-template name ="tmpNMwriteCo-ordinates"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
        <xsl:variable name="presentationId">
          <xsl:value-of select="@presentation:style-name"/>
        </xsl:variable>
        <!--Background Color-->
        <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$presentationId] ">
          <xsl:if test="./style:graphic-properties/@draw:fill='solid'">
            <a:solidFill>
              <a:srgbClr>
                <xsl:attribute name="val">
                  <xsl:value-of select="substring-after(./style:graphic-properties/@draw:fill-color,'#')" />
                </xsl:attribute>
              </a:srgbClr>
            </a:solidFill>
          </xsl:if>
          <!--End-->
          <!--Line Style-->
          <xsl:if test="not(./style:graphic-properties/@draw:stroke='none')">
            <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$presentationId]/style:graphic-properties">
              <xsl:call-template name ="tmpNMgetFillColor"/>
              <xsl:call-template name ="tmpNMLineStyle">
                <xsl:with-param name="parentStyle" select="$presentationId"/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:if>
          <!--End-->
        </xsl:for-each>
      </p:spPr>
      <p:txBody>
        <xsl:call-template name ="tmpNMTextAlignment" >
          <xsl:with-param name ="prId" select ="@presentation:style-name"/>
        </xsl:call-template >
        <a:lstStyle>
          <!--<a:lvl1pPr algn="l">
            <a:defRPr sz="1200"/>
          </a:lvl1pPr>-->
          <a:lvl1pPr>
            <!--<xsl:value-of select="'ctr'"/>-->
            <xsl:variable name="ParId">
              <xsl:choose>
                <xsl:when test="./draw:text-box/text:p/@text:style-name">
                  <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="@presentation:style-name"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <!--<xsl:variable name ="ParId">
              <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
            </xsl:variable>-->
            <xsl:if test ="./draw:text-box/text:p">
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
              <!--<xsl:variable name ="ParId">
                <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
              </xsl:variable>-->
              <xsl:variable name="textId">
                <xsl:choose>
                  <xsl:when test="./draw:text-box/text:p/text:span/@text:style-name">
                    <xsl:value-of select ="./draw:text-box/text:p/text:span/@text:style-name"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="@presentation:style-name"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>
              <!--<xsl:variable name="textId">
                <xsl:value-of select="./draw:text-box/text:p/text:span/@text:style-name"/>
              </xsl:variable>-->
              <xsl:choose>
                <xsl:when test="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@fo:font-size">
                  <xsl:attribute name="sz">
                    <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:for-each>
                  </xsl:attribute>
                </xsl:when>
                <!--When draw:text-box/text:p/text:span is not there-->
                <xsl:when test="document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties/@fo:font-size">
                  <xsl:attribute name="sz">
                    <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:for-each>
                  </xsl:attribute>
                </xsl:when>
              </xsl:choose>
              <!--sateesh-->
              <xsl:for-each select="document('styles.xml')//style:style[@style:name=$textId]">
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
                  <xsl:call-template name="tmpNMUnderline">
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
                    <xsl:when test="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@fo:color">
                      <xsl:for-each select="document('styles.xml')//style:style[@style:name=$textId]">
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
              <xsl:if test="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@fo:text-shadow">
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
                    <xsl:when test="translate(document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties/@fo:font-family, &quot;'&quot;,'')">
                      <xsl:value-of select="translate(document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
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
        <a:p>
          <!--Default-->
          <xsl:if test="not(./draw:text-box/text:p/text:span)">
            <a:endParaRPr lang="en-US"/>
          </xsl:if>
          <!--End-->
          <xsl:if test="./draw:text-box/text:p/text:span">
            <a:r>
              <a:rPr lang="en-US" dirty="0" err="1" smtClean="0"/>
              <a:t>
                <xsl:for-each select="draw:text-box/text:p/text:span">
                  <xsl:value-of select="."/>
                </xsl:for-each>
              </a:t>
            </a:r>
            <a:endParaRPr lang="en-US" dirty="0"/>
          </xsl:if>
        </a:p>
      </p:txBody>
    </p:sp>
  </xsl:template>
  
  <xsl:template name="tmpNMDatetime">
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="3" name="Date Placeholder 2"/>
        <p:cNvSpPr>
          <a:spLocks noGrp="1"/>
        </p:cNvSpPr>
        <p:nvPr>
          <p:ph type="dt" idx="1"/>
        </p:nvPr>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <xsl:call-template name ="tmpNMwriteCo-ordinates"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
        <xsl:variable name="presentationId">
          <xsl:value-of select="@presentation:style-name"/>
        </xsl:variable>
        <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$presentationId] ">
          <xsl:if test="./style:graphic-properties/@draw:fill='solid'">
            <a:solidFill>
              <a:srgbClr>
                <xsl:attribute name="val">
                  <xsl:value-of select="substring-after(./style:graphic-properties/@draw:fill-color,'#')" />
                </xsl:attribute>
              </a:srgbClr>
            </a:solidFill>
          </xsl:if>
          <!--Line Style-->
          <xsl:if test="not(./style:graphic-properties/@draw:stroke='none')">
            <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$presentationId]/style:graphic-properties">
              <xsl:call-template name ="tmpNMgetFillColor"/>
              <xsl:call-template name ="tmpNMLineStyle">
                <xsl:with-param name="parentStyle" select="$presentationId"/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:if>
          <!--End-->
        </xsl:for-each>
      </p:spPr>
      <p:txBody>
        <xsl:call-template name ="tmpNMTextAlignment" >
          <xsl:with-param name ="prId" select ="@presentation:style-name"/>
        </xsl:call-template >
        <a:lstStyle>
          <!--<a:lvl1pPr algn="r">
            <a:defRPr sz="1200"/>
          </a:lvl1pPr>-->
          <a:lvl1pPr>
            <!--<xsl:value-of select="'ctr'"/>-->
            <!--<xsl:variable name ="ParId">
              <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
            </xsl:variable>-->
            <xsl:variable name="ParId">
              <xsl:choose>
                <xsl:when test="./draw:text-box/text:p/@text:style-name">
                  <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="@presentation:style-name"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:if test ="./draw:text-box/text:p">
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
              <!--<xsl:variable name ="ParId">
                <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
              </xsl:variable>-->
              <!--<xsl:variable name="textId">
                <xsl:value-of select="./draw:text-box/text:p/text:span/@text:style-name"/>
              </xsl:variable>-->
              <xsl:variable name="textId">
                <xsl:choose>
                  <xsl:when test="./draw:text-box/text:p/text:span/@text:style-name">
                    <xsl:value-of select ="./draw:text-box/text:p/text:span/@text:style-name"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="@presentation:style-name"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>
              <xsl:choose>
                <xsl:when test="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@fo:font-size">
                  <xsl:attribute name="sz">
                    <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:for-each>
                  </xsl:attribute>
                </xsl:when>
                <!--When draw:text-box/text:p/text:span is not there-->
                <xsl:when test="document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties/@fo:font-size">
                  <xsl:attribute name="sz">
                    <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:for-each>
                  </xsl:attribute>
                </xsl:when>
              </xsl:choose>
              <!--sateesh-->
              <xsl:for-each select="document('styles.xml')//style:style[@style:name=$textId]">
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
                  <xsl:call-template name="tmpNMUnderline">
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
                    <xsl:when test="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@fo:color">
                      <xsl:for-each select="document('styles.xml')//style:style[@style:name=$textId]">
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
              <xsl:if test="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@fo:text-shadow">
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
                    <xsl:when test="translate(document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties/@fo:font-family, &quot;'&quot;,'')">
                      <xsl:value-of select="translate(document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
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
        <a:p>
          <!--default-->
          <xsl:if test="not(./draw:text-box/text:p/text:span)">
            <a:fld>
              <xsl:attribute name ="id">
                <xsl:value-of select ="'{0DF055E1-9ED0-4568-B092-CB4520033461}'"/>
              </xsl:attribute>
              <xsl:attribute name ="type">
                <xsl:value-of select ="'datetime1'"/>
              </xsl:attribute>
              <a:rPr lang="en-US" smtClean="0"/>
              <a:t></a:t>
            </a:fld>
            <a:endParaRPr lang="en-US"/>
          </xsl:if>
          <!--end-->
          <xsl:if test="./draw:text-box/text:p/text:span">
            <a:r>
              <a:rPr lang="en-US" dirty="0" smtClean="0"/>
              <a:t>
                <xsl:for-each select="draw:text-box/text:p/text:span">
                  <xsl:value-of select="."/>
                </xsl:for-each>
              </a:t>
            </a:r>
            <a:endParaRPr lang="en-US" dirty="0"/>
          </xsl:if>
        </a:p>
      </p:txBody>
    </p:sp>
  </xsl:template>

  <xsl:template name="Page">
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="4" name="Slide Image Placeholder 3"/>
        <p:cNvSpPr>
          <a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/>
        </p:cNvSpPr>
        <p:nvPr>
          <p:ph type="sldImg" idx="2"/>
        </p:nvPr>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <xsl:call-template name ="tmpNMwriteCo-ordinates"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
        <a:noFill />
         <a:ln w="12700">
          <a:solidFill>
            <a:prstClr val="black" />
          </a:solidFill>
        </a:ln>
      </p:spPr>
      <p:txBody>
        <xsl:call-template name ="tmpNMTextAlignment" >
          <xsl:with-param name ="prId" select ="@presentation:style-name"/>
        </xsl:call-template >
        <a:lstStyle/>
        <a:p>
          <a:endParaRPr lang="en-US"/>
        </a:p>
      </p:txBody>
    </p:sp>
  </xsl:template>
  
  <xsl:template name="Notes">
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="5" name="Notes Placeholder 4"/>
        <p:cNvSpPr>
          <a:spLocks noGrp="1"/>
        </p:cNvSpPr>
        <p:nvPr>
          <p:ph type="body" sz="quarter" idx="3"/>
        </p:nvPr>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <xsl:call-template name ="tmpNMwriteCo-ordinates"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </p:spPr>
      <p:txBody>
        <xsl:call-template name ="tmpNMTextAlignment" >
          <xsl:with-param name ="prId" select ="@presentation:style-name"/>
        </xsl:call-template >
        <a:lstStyle/>
        <a:p>
          <a:pPr lvl="0"/>
          <a:r>
            <a:rPr lang="en-US" smtClean="0"/>
            <a:t>Click to edit the notes format</a:t>
          </a:r>
        </a:p>
      </p:txBody>
    </p:sp>
  </xsl:template>
  
  <xsl:template name="tmpNMFooter">
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="6" name="Footer Placeholder 5"/>
        <p:cNvSpPr>
          <a:spLocks noGrp="1"/>
        </p:cNvSpPr>
        <p:nvPr>
          <p:ph type="ftr" sz="quarter" idx="4"/>
        </p:nvPr>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <xsl:call-template name ="tmpNMwriteCo-ordinates"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
        <xsl:variable name="presentationId">
          <xsl:value-of select="@presentation:style-name"/>
        </xsl:variable>
        <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$presentationId] ">
          <xsl:if test="./style:graphic-properties/@draw:fill='solid'">
            <a:solidFill>
              <a:srgbClr>
                <xsl:attribute name="val">
                  <xsl:value-of select="substring-after(./style:graphic-properties/@draw:fill-color,'#')" />
                </xsl:attribute>
              </a:srgbClr>
            </a:solidFill>
          </xsl:if>
          <!--Line Style-->
          <xsl:if test="not(./style:graphic-properties/@draw:stroke='none')">
            <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$presentationId]/style:graphic-properties">
              <xsl:call-template name ="tmpNMgetFillColor"/>
              <xsl:call-template name ="tmpNMLineStyle">
                <xsl:with-param name="parentStyle" select="$presentationId"/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:if>
          <!--End-->
        </xsl:for-each>
      </p:spPr>
      <p:txBody>
        <xsl:call-template name ="tmpNMTextAlignment" >
          <xsl:with-param name ="prId" select ="@presentation:style-name"/>
        </xsl:call-template >
        <a:lstStyle>
          <!--<a:lvl1pPr algn="l">
            <a:defRPr sz="1200"/>
          </a:lvl1pPr>-->
          <a:lvl1pPr>
            <!--<xsl:value-of select="'ctr'"/>-->
            <!--<xsl:variable name ="ParId">
              <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
            </xsl:variable>-->
            <xsl:variable name="ParId">
              <xsl:choose>
                <xsl:when test="./draw:text-box/text:p/@text:style-name">
                  <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="@presentation:style-name"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:if test ="./draw:text-box/text:p">
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
              <!--<xsl:variable name ="ParId">
                <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
              </xsl:variable>-->
              <!--<xsl:variable name="textId">
                <xsl:value-of select="./draw:text-box/text:p/text:span/@text:style-name"/>
              </xsl:variable>-->
              <xsl:variable name="textId">
                <xsl:choose>
                  <xsl:when test="./draw:text-box/text:p/text:span/@text:style-name">
                    <xsl:value-of select ="./draw:text-box/text:p/text:span/@text:style-name"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="@presentation:style-name"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>
              <xsl:choose>
                <xsl:when test="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@fo:font-size">
                  <xsl:attribute name="sz">
                    <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:for-each>
                  </xsl:attribute>
                </xsl:when>
                <!--When draw:text-box/text:p/text:span is not there-->
                <xsl:when test="document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties/@fo:font-size">
                  <xsl:attribute name="sz">
                    <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:for-each>
                  </xsl:attribute>
                </xsl:when>
              </xsl:choose>
              <!--sateesh-->
              <xsl:for-each select="document('styles.xml')//style:style[@style:name=$textId]">
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
                  <xsl:call-template name="tmpNMUnderline">
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
                    <xsl:when test="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@fo:color">
                      <xsl:for-each select="document('styles.xml')//style:style[@style:name=$textId]">
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
              <xsl:if test="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@fo:text-shadow">
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
                    <xsl:when test="translate(document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties/@fo:font-family, &quot;'&quot;,'')">
                      <xsl:value-of select="translate(document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
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
        <a:p>
          <!--default-->
          <xsl:if test="not(./draw:text-box/text:p/text:span)">
            <a:endParaRPr lang="en-US"/>
          </xsl:if>
          <!--end-->
          <xsl:if test="./draw:text-box/text:p/text:span">
            <a:r>
              <a:rPr lang="en-US" dirty="0" smtClean="0"/>
              <a:t>
                <xsl:for-each select="draw:text-box/text:p/text:span">
                  <xsl:value-of select="."/>
                </xsl:for-each>
              </a:t>
            </a:r>
            <a:endParaRPr lang="en-US" dirty="0"/>
          </xsl:if>
        </a:p>
      </p:txBody>
    </p:sp>
  </xsl:template>

  <xsl:template name="tmpNMPagenumber">
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="7" name="Slide Number Placeholder 6"/>
        <p:cNvSpPr>
          <a:spLocks noGrp="1"/>
        </p:cNvSpPr>
        <p:nvPr>
          <p:ph type="sldNum" sz="quarter" idx="5"/>
        </p:nvPr>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <xsl:call-template name ="tmpNMwriteCo-ordinates"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
        <xsl:variable name="presentationId">
          <xsl:value-of select="@presentation:style-name"/>
        </xsl:variable>
        <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$presentationId] ">
          <xsl:if test="./style:graphic-properties/@draw:fill='solid'">
            <a:solidFill>
              <a:srgbClr>
                <xsl:attribute name="val">
                  <xsl:value-of select="substring-after(./style:graphic-properties/@draw:fill-color,'#')" />
                </xsl:attribute>
              </a:srgbClr>
            </a:solidFill>
          </xsl:if>
          <!--Line Style-->
          <xsl:if test="not(./style:graphic-properties/@draw:stroke='none')">
            <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$presentationId]/style:graphic-properties">
              <xsl:call-template name ="tmpNMgetFillColor"/>
              <xsl:call-template name ="tmpNMLineStyle">
                <xsl:with-param name="parentStyle" select="$presentationId"/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:if>
          <!--End-->
        </xsl:for-each>
      </p:spPr>
      <p:txBody>
        <xsl:call-template name ="tmpNMTextAlignment" >
          <xsl:with-param name ="prId" select ="@presentation:style-name"/>
        </xsl:call-template >
        <a:lstStyle>
          <!--<a:lvl1pPr algn="r">
            <a:defRPr sz="1200"/>
          </a:lvl1pPr>-->
          <a:lvl1pPr>
            <!--<xsl:value-of select="'ctr'"/>-->
            <!--<xsl:variable name ="ParId">
              <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
            </xsl:variable>-->
            <xsl:variable name="ParId">
              <xsl:choose>
                <xsl:when test="./draw:text-box/text:p/@text:style-name">
                  <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="@presentation:style-name"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:if test ="./draw:text-box/text:p">
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
              <!--<xsl:variable name ="ParId">
                <xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
              </xsl:variable>-->
              <!--<xsl:variable name="textId">
                <xsl:value-of select="./draw:text-box/text:p/text:span/@text:style-name"/>
              </xsl:variable>-->
              <xsl:variable name="textId">
                <xsl:choose>
                  <xsl:when test="./draw:text-box/text:p/text:span/@text:style-name">
                    <xsl:value-of select ="./draw:text-box/text:p/text:span/@text:style-name"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="@presentation:style-name"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>
              <xsl:choose>
                <xsl:when test="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@fo:font-size">
                  <xsl:attribute name="sz">
                    <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:for-each>
                  </xsl:attribute>
                </xsl:when>
                <!--When draw:text-box/text:p/text:span is not there-->
                <xsl:when test="document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties/@fo:font-size">
                  <xsl:attribute name="sz">
                    <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties">
                      <xsl:call-template name ="convertToPoints">
                        <xsl:with-param name ="unit" select ="'pt'"/>
                        <xsl:with-param name ="length" select ="@fo:font-size"/>
                      </xsl:call-template>
                    </xsl:for-each>
                  </xsl:attribute>
                </xsl:when>
              </xsl:choose>
              <!--sateesh-->
              <xsl:for-each select="document('styles.xml')//style:style[@style:name=$textId]">
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
                  <xsl:call-template name="tmpNMUnderline">
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
                    <xsl:when test="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@fo:color">
                      <xsl:for-each select="document('styles.xml')//style:style[@style:name=$textId]">
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
              <xsl:if test="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@fo:text-shadow">
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
                    <xsl:when test="translate(document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties/@fo:font-family, &quot;'&quot;,'')">
                      <xsl:value-of select="translate(document('styles.xml')//style:style[@style:name=$ParId]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
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
        <a:p>
          <!--default-->
          <xsl:if test="not(./draw:text-box/text:p/text:span)">
            <a:fld type="slidenum">
              <xsl:attribute name ="id">
                <xsl:value-of select ="'{2AB35230-451B-423A-AB56-944A40C3FA51}'"/>
              </xsl:attribute>
              <a:rPr lang="en-US" smtClean="0"/>
              <a:t>‹#›</a:t>
            </a:fld>
            <a:endParaRPr lang="en-US"/>
          </xsl:if>
          <!--end-->
          <xsl:if test="./draw:text-box/text:p/text:span">
            <a:r>
              <a:rPr lang="en-US" dirty="0" smtClean="0"/>
              <a:t>
                <xsl:for-each select="draw:text-box/text:p/text:span">
                  <xsl:value-of select="."/>
                </xsl:for-each>
              </a:t>
            </a:r>
            <a:endParaRPr lang="en-US" dirty="0"/>
          </xsl:if>
        </a:p>
      </p:txBody>
    </p:sp>
  </xsl:template>
  
  <xsl:template name ="tmpNMwriteCo-ordinates">
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

  <xsl:template name ="tmpNMTextAlignment">
    <xsl:param name="masterName"/>
    <xsl:param name ="prId"/>
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
      <xsl:for-each select ="document('styles.xml')//style:style[@style:name=$prId]/style:graphic-properties">
      <!--<xsl:for-each select ="document('styles.xml')//style:style[@style:name=$parentName]/style:graphic-properties">-->
        <xsl:attribute name ="tIns">
          <xsl:choose>
            <xsl:when test ="@fo:padding-top and
					substring-before(@fo:padding-top,'cm') &gt; 0">
              <!--<xsl:if test ="@fo:padding-top">-->
              <xsl:value-of select ="format-number(substring-before(@fo:padding-top,'cm')  *   360000 ,'#.##')"/>
            </xsl:when>
            <!-- Added by Lohith - or @fo:padding-top = '0cm'-->
            <xsl:when test ="not(@fo:padding-top) or @fo:padding-top = '0cm'">
              <xsl:value-of select ="0"/>
            </xsl:when >
            <xsl:otherwise>
              <xsl:value-of select ="0"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name ="bIns">
          <xsl:choose>
            <xsl:when test="@fo:padding-bottom and
					substring-before(@fo:padding-bottom,'cm') &gt; 0">
              <!--<xsl:if test ="@fo:padding-bottom">-->
              <xsl:value-of select ="format-number(substring-before(@fo:padding-bottom,'cm')  *   360000 ,'#.##')"/>
            </xsl:when>
            <!-- Added by Lohith - or @fo:padding-bottom = '0cm'-->
            <xsl:when test ="not(@fo:padding-bottom) or @fo:padding-bottom = '0cm'">
              <xsl:value-of select ="0"/>
            </xsl:when >
            <xsl:otherwise>
              <xsl:value-of select ="0"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name ="lIns">
          <xsl:choose>
            <xsl:when test="@fo:padding-left and
					substring-before(@fo:padding-left,'cm') &gt; 0">
              <!--<xsl:if test ="@fo:padding-left">-->
              <xsl:value-of select ="format-number(substring-before(@fo:padding-left,'cm')  *   360000 ,'#.##')"/>
            </xsl:when>
            <!-- Added by Lohith - or @fo:padding-left = '0cm'-->
            <xsl:when test ="not(@fo:padding-left) or @fo:padding-left = '0cm'">
              <xsl:value-of select ="0"/>
            </xsl:when >
            <xsl:otherwise>
              <xsl:value-of select ="0"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name ="rIns">
          <xsl:choose>
            <xsl:when test="@fo:padding-right and
					substring-before(@fo:padding-right,'cm') &gt; 0">
              <!--<xsl:if test ="@fo:padding-right">-->
              <xsl:value-of select ="format-number(substring-before(@fo:padding-right,'cm')  *   360000 ,'#.##')"/>
            </xsl:when>
            <!-- Added by Lohith - or @fo:padding-right = '0cm'-->
            <xsl:when test ="not(@fo:padding-right) or @fo:padding-right = '0cm'">
              <xsl:value-of select ="0"/>
            </xsl:when >
            <xsl:otherwise>
              <xsl:value-of select ="0"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
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
        <xsl:attribute name ="wrap">
          <xsl:choose >
            <!--<xsl:when test ="@fo:wrap-option ='no-wrap'">
							<xsl:value-of select ="'none'"/>
						</xsl:when>
						<xsl:when test ="@fo:wrap-option ='wrap'">
							<xsl:value-of select ="'square'"/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select ="'square'"/>
						</xsl:otherwise>-->
            <xsl:when test ="((@draw:auto-grow-height = 'false') and (@draw:auto-grow-width = 'false')) or (@fo:wrap-option='wrap')">
              <xsl:value-of select ="'none'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select ="'square'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>

      </xsl:for-each>
    </a:bodyPr>
  </xsl:template>

  <xsl:template name ="tmpNMLineStyle">
    <xsl:param name ="parentStyle" />
    <a:ln>
      <!-- Border width -->
      <xsl:choose>
        <xsl:when test ="@svg:stroke-width">
          <xsl:attribute name ="w">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name="length"  select ="@svg:stroke-width"/>
              <xsl:with-param name ="unit" select ="'cm'"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:when>
        <!-- Default border width from styles.xml-->
        <xsl:when test ="($parentStyle != '')">
          <xsl:for-each select ="document('styles.xml')//style:style[@style:name = $parentStyle]/style:graphic-properties">
            <xsl:if test ="@svg:stroke-width">
              <xsl:attribute name ="w">
                <xsl:call-template name ="convertToPoints">
                  <xsl:with-param name="length"  select ="@svg:stroke-width"/>
                  <xsl:with-param name ="unit" select ="'cm'"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if >
            <!-- Code for the Bug 1746356 -->
            <xsl:if test ="not(@svg:stroke-width)">
              <xsl:attribute name ="w">
                <xsl:value-of select ="'0'"/>
              </xsl:attribute>
            </xsl:if >
            <!-- Code for the Bug 1746356 -->
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
      <!-- Cap type-->
      <xsl:if test ="@draw:stroke-dash">
        <xsl:variable name ="dash" select ="@draw:stroke-dash" />
        <xsl:if test ="document('styles.xml')/office:document-styles/office:styles/draw:stroke-dash[@draw:name=$dash]">
          <xsl:if test ="document('styles.xml')/office:document-styles/office:styles/draw:stroke-dash[@draw:name=$dash]/@draw:style='round'">
            <xsl:attribute name ="cap">
              <xsl:value-of select ="'rnd'"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>
      </xsl:if>

      <!-- Line color -->
      <xsl:variable name ="lineOpacity"	>
        <xsl:value-of select ="@svg:stroke-opacity"/>
      </xsl:variable>
      <xsl:choose>
        <!-- Invisible line-->
        <xsl:when test ="@draw:stroke='none'">
          <a:noFill />
        </xsl:when>
        <!-- Solid color -->
        <xsl:when test ="@svg:stroke-color">
          <xsl:call-template name ="tmpNMgetFillColor">
            <xsl:with-param name ="fill-color" select ="@svg:stroke-color" />
            <xsl:with-param name ="opacity" select ="@svg:stroke-opacity" />
          </xsl:call-template>
        </xsl:when>
        <xsl:when test ="($parentStyle != '')">
          <xsl:for-each select ="document('styles.xml')//style:style[@style:name = $parentStyle]/style:graphic-properties">
            <xsl:choose>
              <!-- Invisible line-->
              <xsl:when test ="@draw:stroke='none'">
                <a:noFill />
              </xsl:when>
              <!-- Solid color -->
              <xsl:when test ="@svg:stroke-color">
                <xsl:call-template name ="tmpNMgetFillColor">
                  <xsl:with-param name ="fill-color" select ="@svg:stroke-color" />
                  <xsl:with-param name ="opacity" select ="$lineOpacity" />
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>

      <!-- Dash type-->
      <xsl:if test ="(@draw:stroke='dash')">
        <a:prstDash>
          <xsl:attribute name ="val">
            <xsl:call-template name ="getDashType">
              <xsl:with-param name ="stroke-dash" select ="@draw:stroke-dash" />
            </xsl:call-template>
          </xsl:attribute>
        </a:prstDash>
      </xsl:if>

      <!-- Line join type-->
      <xsl:if test ="@draw:stroke-linejoin">
        <xsl:call-template name ="getJoinType">
          <xsl:with-param name ="stroke-linejoin" select ="@draw:stroke-linejoin" />
        </xsl:call-template>
      </xsl:if>

      <!--Arrow type-->
      <xsl:if test="(@draw:marker-start) and (@draw:marker-start != '')">
        <a:headEnd>
          <xsl:attribute name ="type">
            <xsl:call-template name ="getArrowType">
              <xsl:with-param name ="ArrowType" select ="@draw:marker-start" />
            </xsl:call-template>
          </xsl:attribute>
          <xsl:if test ="@draw:marker-start-width">
            <xsl:call-template name ="setArrowSize">
              <xsl:with-param name ="size" select ="substring-before(@draw:marker-start-width,'cm')" />
            </xsl:call-template >
          </xsl:if>
        </a:headEnd>
      </xsl:if>

      <xsl:if test="(@draw:marker-end) and (@draw:marker-end != '')">
        <a:tailEnd>
          <xsl:attribute name ="type">
            <xsl:call-template name ="getArrowType">
              <xsl:with-param name ="ArrowType" select ="@draw:marker-end" />
            </xsl:call-template>
          </xsl:attribute>

          <xsl:if test ="@draw:marker-end-width">
            <xsl:call-template name ="setArrowSize">
              <xsl:with-param name ="size" select ="substring-before(@draw:marker-end-width,'cm')" />
            </xsl:call-template >
          </xsl:if>
        </a:tailEnd>
      </xsl:if>
    </a:ln>
  </xsl:template>

  <xsl:template name ="tmpNMgetFillColor">
    <xsl:param name ="fill-color" />
    <xsl:param name ="opacity" />
    <xsl:if test ="$fill-color != ''">
      <a:solidFill>
        <a:srgbClr>
          <xsl:attribute name ="val">
            <xsl:value-of select ="substring-after($fill-color,'#')"/>
          </xsl:attribute>
          <xsl:if test ="$opacity != ''">
            <a:alpha>
              <xsl:variable name ="alpha" select ="substring-before($opacity,'%')" />
              <xsl:attribute name ="val">
                <xsl:if test ="$alpha = 0">
                  <xsl:value-of select ="0000"/>
                </xsl:if>
                <xsl:if test ="$alpha != 0">
                  <xsl:value-of select ="$alpha * 1000"/>
                </xsl:if>
              </xsl:attribute>
            </a:alpha>
          </xsl:if>
        </a:srgbClr>
      </a:solidFill>
    </xsl:if>
  </xsl:template>

  <xsl:template name="tmpNMUnderline">
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
      <xsl:when test="$uType = 'single'">
          <xsl:value-of select ="'sng'"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name ="notesMasterRel" match ="/office:document-content/office:body/office:presentation/draw:page">
    <xsl:param name ="ThemeId"/>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <!--<Relationship Id="nmasterId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
                    Target="../theme/themenmaster1.xml"/>-->
      <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme">
        <xsl:attribute name="Id">
          <xsl:value-of select="'nmThemeId'"/>
        </xsl:attribute>
        <xsl:attribute name ="Target">
          <xsl:value-of select ="concat('../theme/theme',$ThemeId,'.xml')"/>
        </xsl:attribute>
      </Relationship >
    </Relationships>
  </xsl:template>
</xsl:stylesheet>

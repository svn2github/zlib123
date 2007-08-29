<?xml version="1.0" encoding="utf-8" ?>
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
<!--<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
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
  exclude-result-prefixes="odf style text number draw page svg presentation fo script xlink">
  <xsl:import href ="common.xsl"/>
  <xsl:import href ="shapes_direct.xsl"/>
  <!-- Templated for Slide Layouts - Added by lohith.ar -->
  <xsl:template name="InsertSlideLayout1">
    <xsl:param name ="MasterName" />
    <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
				   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
				   xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
				   type="title" preserve="1">
      <p:cSld name="Title Slide">
        <xsl:variable name="subtileName">
          <xsl:value-of select="concat($MasterName,'-subtitle')"/>
        </xsl:variable>
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
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Title 1"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph type="ctrTitle"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="685800" y="2130425" />
                <a:ext cx="7772400" cy="1470025" />
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master title style</a:t>
                </a:r>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="3" name="Subtitle 2"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph type="subTitle" idx="1"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="1371600" y="3886200" />
                <a:ext cx="6400800" cy="1752600" />
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr>
               <xsl:variable name ="anchorValue">
                  <xsl:choose >
                    <xsl:when test ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$subtileName]/style:graphic-properties/@draw:textarea-vertical-align ='top'">
                      <xsl:value-of select ="'t'"/>
                    </xsl:when>
                    <xsl:when test ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$subtileName]/style:graphic-properties/@draw:textarea-vertical-align ='middle'">
                      <xsl:value-of select ="'ctr'"/>
                    </xsl:when>
                    <xsl:when test ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$subtileName]/style:graphic-properties/@draw:textarea-vertical-align ='bottom'">
                      <xsl:value-of select ="'b'"/>
                    </xsl:when>
                  </xsl:choose>
                </xsl:variable>
                <xsl:if test ="$anchorValue != ''">
                  <xsl:attribute name ="anchor">
                    <xsl:value-of select ="$anchorValue"/>
                  </xsl:attribute>
                </xsl:if>
              </a:bodyPr>
              <a:lstStyle>
                <a:lvl1pPr>
                  <xsl:variable name ="alignmentValue">
                    <xsl:choose >
                      <xsl:when test ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$subtileName]/style:paragraph-properties/@fo:text-align ='center'">
                        <xsl:value-of select ="'ctr'"/>
                      </xsl:when>
                      <xsl:when test ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$subtileName]/style:paragraph-properties/@fo:text-align ='end'">
                        <xsl:value-of select ="'r'"/>
                      </xsl:when>
                      <xsl:when test ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$subtileName]/style:paragraph-properties/@fo:text-align ='justify'">
                        <xsl:value-of select ="'just'"/>
                      </xsl:when>
                      <xsl:when test ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$subtileName]/style:paragraph-properties/@fo:text-align ='start'">
                        <xsl:value-of select ="'l'"/>
                      </xsl:when>
                    </xsl:choose>
                  </xsl:variable>
                  <xsl:if test ="$alignmentValue != ''">
                    <xsl:attribute name ="algn">
                      <xsl:value-of select ="$alignmentValue"/>
                    </xsl:attribute>
                  </xsl:if>
                  <a:buNone/>
                  <a:defRPr>
                    <!--<a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>-->
                  </a:defRPr>
                </a:lvl1pPr>
                <a:lvl2pPr>
                  <a:buNone/>
                  <a:defRPr>
                    <!--<a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>-->
                  </a:defRPr>
                </a:lvl2pPr>
                <a:lvl3pPr>
                  <a:buNone/>
                  <a:defRPr>
                    <!--<a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>-->
                  </a:defRPr>
                </a:lvl3pPr>
                <a:lvl4pPr>
                  <a:buNone/>
                  <a:defRPr>
                    <!--<a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>-->
                  </a:defRPr>
                </a:lvl4pPr>
                <a:lvl5pPr>
                  <a:buNone/>
                  <a:defRPr>
                    <!--<a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>-->
                  </a:defRPr>
                </a:lvl5pPr>
                <a:lvl6pPr>
                  <a:buNone/>
                  <a:defRPr>
                    <!--<a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>-->
                  </a:defRPr>
                </a:lvl6pPr>
                <a:lvl7pPr>
                  <a:buNone/>
                  <a:defRPr>
                    <!--<a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>-->
                  </a:defRPr>
                </a:lvl7pPr>
                <a:lvl8pPr>
                  <a:buNone/>
                  <a:defRPr>
                    <!--<a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>-->
                  </a:defRPr>
                </a:lvl8pPr>
                <a:lvl9pPr>
                  <a:buNone/>
                  <a:defRPr>
                    <!--<a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>-->
                  </a:defRPr>
                </a:lvl9pPr>
              </a:lstStyle>
              <a:p>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master subtitle style</a:t>
                </a:r>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <xsl:call-template name ="masterFooter">
            <xsl:with-param name="smasterName" select="$MasterName"/>
            <xsl:with-param name ="startNo" select ="4"/>
          </xsl:call-template>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>

  <xsl:template name="InsertSlideLayout2">
    <xsl:param name ="MasterName" />
    <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
					xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
					xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="obj" preserve="1">
      <p:cSld name="Title and Content">
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
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Title 1"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph type="title"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="457200" y="274638" />
                <a:ext cx="8229600" cy="1143000" />
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master title style</a:t>
                </a:r>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="3" name="Content Placeholder 2"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph idx="1"/>
              </p:nvPr>
            </p:nvSpPr>
            <!--<p:spPr/>-->
            <p:spPr>
              <a:xfrm>
                <a:off x="457200" y="1600200" />
                <a:ext cx="8229600" cy="4525963" />
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr lvl="0"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master text styles</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="1"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Second level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="2"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Third level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="3"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Fourth level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="4"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Fifth level</a:t>
                </a:r>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <xsl:call-template name ="masterFooter">
            <xsl:with-param name="smasterName" select="$MasterName"/>
            <xsl:with-param name ="startNo" select ="4"/>
          </xsl:call-template>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>

  <xsl:template name="InsertSlideLayout3">
    <xsl:param name ="MasterName" />
    <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
		  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
		  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="secHead" preserve="1">
      <p:cSld name="Section Header">
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
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Title 1"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph type="title"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="722313" y="4406900"/>
                <a:ext cx="7772400" cy="1362075"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <!--anchor="t"-->
              <a:lstStyle>
                <a:lvl1pPr>
                  <a:defRPr/>
                </a:lvl1pPr>
              </a:lstStyle>
              <a:p>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master title style</a:t>
                </a:r>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="3" name="Text Placeholder 2"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph type="body" idx="1"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="722313" y="2906713"/>
                <a:ext cx="7772400" cy="1500187"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle>
                <a:lvl1pPr>
                  <a:buNone/>
                  <a:defRPr>
                    <!--<a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>-->
                  </a:defRPr>
                </a:lvl1pPr>
                <a:lvl2pPr>
                  <a:buNone/>
                  <a:defRPr>
                    <!--<a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>-->
                  </a:defRPr>
                </a:lvl2pPr>
                <a:lvl3pPr>
                  <a:buNone/>
                  <a:defRPr>
                    <!--<a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>-->
                  </a:defRPr>
                </a:lvl3pPr>
                <a:lvl4pPr>
                  <a:buNone/>
                  <a:defRPr>
                    <!--<a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>-->
                  </a:defRPr>
                </a:lvl4pPr>
                <a:lvl5pPr>
                  <a:buNone/>
                  <a:defRPr>
                    <!--<a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>-->
                  </a:defRPr>
                </a:lvl5pPr>
                <a:lvl6pPr>
                  <a:buNone/>
                  <a:defRPr>
                    <!--<a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>-->
                  </a:defRPr>
                </a:lvl6pPr>
                <a:lvl7pPr>
                  <a:buNone/>
                  <a:defRPr sz="1400">
                    <!--<a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>-->
                  </a:defRPr>
                </a:lvl7pPr>
                <a:lvl8pPr>
                  <a:buNone/>
                  <a:defRPr>
                    <!--<a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>-->
                  </a:defRPr>
                </a:lvl8pPr>
                <a:lvl9pPr>
                  <a:buNone/>
                  <a:defRPr>
                    <!--<a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>-->
                  </a:defRPr>
                </a:lvl9pPr>
              </a:lstStyle>
              <a:p>
                <a:pPr lvl="0"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master text styles</a:t>
                </a:r>
              </a:p>
            </p:txBody>
          </p:sp>
          <xsl:call-template name ="masterFooter">
            <xsl:with-param name="smasterName" select="$MasterName"/>
            <xsl:with-param name ="startNo" select ="4"/>
          </xsl:call-template>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>

  <xsl:template name="InsertSlideLayout4">
    <xsl:param name ="MasterName" />
    <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
					 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
					 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="twoObj" preserve="1">
      <p:cSld name="Two Content">
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
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Title 1"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph type="title"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="457200" y="274638"/>
                <a:ext cx="8229600" cy="1143000"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master title style</a:t>
                </a:r>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="3" name="Content Placeholder 2"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph sz="half" idx="1"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="457200" y="1600200"/>
                <a:ext cx="4038600" cy="4525963"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle>
                <a:lvl1pPr>
                  <a:defRPr/>
                </a:lvl1pPr>
                <a:lvl2pPr>
                  <a:defRPr />
                </a:lvl2pPr>
                <a:lvl3pPr>
                  <a:defRPr/>
                </a:lvl3pPr>
                <a:lvl4pPr>
                  <a:defRPr/>
                </a:lvl4pPr>
                <a:lvl5pPr>
                  <a:defRPr/>
                </a:lvl5pPr>
                <a:lvl6pPr>
                  <a:defRPr/>
                </a:lvl6pPr>
                <a:lvl7pPr>
                  <a:defRPr/>
                </a:lvl7pPr>
                <a:lvl8pPr>
                  <a:defRPr/>
                </a:lvl8pPr>
                <a:lvl9pPr>
                  <a:defRPr/>
                </a:lvl9pPr>
              </a:lstStyle>
              <a:p>
                <a:pPr lvl="0"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master text styles</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="1"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Second level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="2"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Third level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="3"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Fourth level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="4"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Fifth level</a:t>
                </a:r>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="4" name="Content Placeholder 3"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph sz="half" idx="2"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="4648200" y="1600200"/>
                <a:ext cx="4038600" cy="4525963"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle>
                <a:lvl1pPr>
                  <a:defRPr/>
                </a:lvl1pPr>
                <a:lvl2pPr>
                  <a:defRPr/>
                </a:lvl2pPr>
                <a:lvl3pPr>
                  <a:defRPr/>
                </a:lvl3pPr>
                <a:lvl4pPr>
                  <a:defRPr/>
                </a:lvl4pPr>
                <a:lvl5pPr>
                  <a:defRPr/>
                </a:lvl5pPr>
                <a:lvl6pPr>
                  <a:defRPr/>
                </a:lvl6pPr>
                <a:lvl7pPr>
                  <a:defRPr/>
                </a:lvl7pPr>
                <a:lvl8pPr>
                  <a:defRPr/>
                </a:lvl8pPr>
                <a:lvl9pPr>
                  <a:defRPr/>
                </a:lvl9pPr>
              </a:lstStyle>
              <a:p>
                <a:pPr lvl="0"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master text styles</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="1"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Second level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="2"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Third level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="3"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Fourth level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="4"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Fifth level</a:t>
                </a:r>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <xsl:call-template name ="masterFooter">
            <xsl:with-param name="smasterName" select="$MasterName"/>
            <xsl:with-param name ="startNo" select ="5"/>
          </xsl:call-template>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>

  <xsl:template name="InsertSlideLayout5">
    <xsl:param name ="MasterName" />
    <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
					 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
					 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
					 type="twoTxTwoObj" preserve="1">
      <p:cSld name="Comparison">
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
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Title 1"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph type="title"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="457200" y="274638"/>
                <a:ext cx="8229600" cy="1143000"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle>
                <a:lvl1pPr>
                  <a:defRPr/>
                </a:lvl1pPr>
              </a:lstStyle>
              <a:p>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master title style</a:t>
                </a:r>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="3" name="Text Placeholder 2"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph type="body" idx="1"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="457200" y="1535113"/>
                <a:ext cx="4040188" cy="639762"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr anchor="b"/>
              <a:lstStyle>
                <a:lvl1pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl1pPr>
                <a:lvl2pPr >
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl2pPr>
                <a:lvl3pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl3pPr>
                <a:lvl4pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl4pPr>
                <a:lvl5pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl5pPr>
                <a:lvl6pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl6pPr>
                <a:lvl7pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl7pPr>
                <a:lvl8pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl8pPr>
                <a:lvl9pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl9pPr>
              </a:lstStyle>
              <a:p>
                <a:pPr lvl="0"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master text styles</a:t>
                </a:r>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="4" name="Content Placeholder 3"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph sz="half" idx="2"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="457200" y="2174875"/>
                <a:ext cx="4040188" cy="3951288"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle>
                <a:lvl1pPr>
                  <a:defRPr/>
                </a:lvl1pPr>
                <a:lvl2pPr>
                  <a:defRPr/>
                </a:lvl2pPr>
                <a:lvl3pPr>
                  <a:defRPr />
                </a:lvl3pPr>
                <a:lvl4pPr>
                  <a:defRPr/>
                </a:lvl4pPr>
                <a:lvl5pPr>
                  <a:defRPr/>
                </a:lvl5pPr>
                <a:lvl6pPr>
                  <a:defRPr/>
                </a:lvl6pPr>
                <a:lvl7pPr>
                  <a:defRPr/>
                </a:lvl7pPr>
                <a:lvl8pPr>
                  <a:defRPr/>
                </a:lvl8pPr>
                <a:lvl9pPr>
                  <a:defRPr/>
                </a:lvl9pPr>
              </a:lstStyle>
              <a:p>
                <a:pPr lvl="0"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master text styles</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="1"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Second level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="2"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Third level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="3"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Fourth level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="4"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Fifth level</a:t>
                </a:r>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="5" name="Text Placeholder 4"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph type="body" sz="quarter" idx="3"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="4645025" y="1535113"/>
                <a:ext cx="4041775" cy="639762"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle>
                <a:lvl1pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl1pPr>
                <a:lvl2pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl2pPr>
                <a:lvl3pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl3pPr>
                <a:lvl4pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl4pPr>
                <a:lvl5pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl5pPr>
                <a:lvl6pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl6pPr>
                <a:lvl7pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl7pPr>
                <a:lvl8pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl8pPr>
                <a:lvl9pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl9pPr>
              </a:lstStyle>
              <a:p>
                <a:pPr lvl="0"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master text styles</a:t>
                </a:r>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="6" name="Content Placeholder 5"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph sz="quarter" idx="4"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="4645025" y="2174875"/>
                <a:ext cx="4041775" cy="3951288"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle>
                <a:lvl1pPr>
                  <a:defRPr/>
                </a:lvl1pPr>
                <a:lvl2pPr>
                  <a:defRPr/>
                </a:lvl2pPr>
                <a:lvl3pPr>
                  <a:defRPr/>
                </a:lvl3pPr>
                <a:lvl4pPr>
                  <a:defRPr/>
                </a:lvl4pPr>
                <a:lvl5pPr>
                  <a:defRPr/>
                </a:lvl5pPr>
                <a:lvl6pPr>
                  <a:defRPr/>
                </a:lvl6pPr>
                <a:lvl7pPr>
                  <a:defRPr/>
                </a:lvl7pPr>
                <a:lvl8pPr>
                  <a:defRPr/>
                </a:lvl8pPr>
                <a:lvl9pPr>
                  <a:defRPr/>
                </a:lvl9pPr>
              </a:lstStyle>
              <a:p>
                <a:pPr lvl="0"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master text styles</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="1"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Second level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="2"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Third level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="3"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Fourth level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="4"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Fifth level</a:t>
                </a:r>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <xsl:call-template name ="masterFooter">
            <xsl:with-param name="smasterName" select="$MasterName"/>
            <xsl:with-param name ="startNo" select ="7"/>
          </xsl:call-template>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>

  <xsl:template name="InsertSlideLayout6">
    <xsl:param name ="MasterName" />
    <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
					 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
					 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
					 type="titleOnly" preserve="1">
      <p:cSld name="Title Only">
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
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Title 1"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph type="title"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="457200" y="274638" />
                <a:ext cx="8229600" cy="1143000" />
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master title style</a:t>
                </a:r>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <xsl:call-template name ="masterFooter">
            <xsl:with-param name="smasterName" select="$MasterName"/>
            <xsl:with-param name ="startNo" select ="3"/>
          </xsl:call-template>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>

  <xsl:template name="InsertSlideLayout7">
    <xsl:param name ="MasterName" />
    <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
					 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
					 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
					 type="blank" preserve="1">
      <p:cSld name="Blank">
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
          <xsl:call-template name ="masterFooter">
            <xsl:with-param name="smasterName" select="$MasterName"/>
            <xsl:with-param name ="startNo" select ="2"/>
          </xsl:call-template>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>

  <xsl:template name="InsertSlideLayout8">
    <xsl:param name ="MasterName" />
    <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
					 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
					 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
					 type="objTx" preserve="1">
      <p:cSld name="Content with Caption">
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
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Title 1"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph type="title"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>

              <a:xfrm>
                <a:off x="457200" y="273050"/>
                <a:ext cx="3008313" cy="1162050"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle>
                <a:lvl1pPr>
                  <a:defRPr/>
                </a:lvl1pPr>
              </a:lstStyle>
              <a:p>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master title style</a:t>
                </a:r>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="3" name="Content Placeholder 2"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph idx="1"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="3575050" y="273050"/>
                <a:ext cx="5111750" cy="5853113"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle>
                <a:lvl1pPr>
                  <a:defRPr/>
                </a:lvl1pPr>
                <a:lvl2pPr>
                  <a:defRPr/>
                </a:lvl2pPr>
                <a:lvl3pPr>
                  <a:defRPr/>
                </a:lvl3pPr>
                <a:lvl4pPr>
                  <a:defRPr/>
                </a:lvl4pPr>
                <a:lvl5pPr>
                  <a:defRPr/>
                </a:lvl5pPr>
                <a:lvl6pPr>
                  <a:defRPr/>
                </a:lvl6pPr>
                <a:lvl7pPr>
                  <a:defRPr/>
                </a:lvl7pPr>
                <a:lvl8pPr>
                  <a:defRPr/>
                </a:lvl8pPr>
                <a:lvl9pPr>
                  <a:defRPr/>
                </a:lvl9pPr>
              </a:lstStyle>
              <a:p>
                <a:pPr lvl="0"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master text styles</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="1"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Second level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="2"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Third level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="3"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Fourth level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="4"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Fifth level</a:t>
                </a:r>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="4" name="Text Placeholder 3"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph type="body" sz="half" idx="2"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="457200" y="1435100"/>
                <a:ext cx="3008313" cy="4691063"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle>
                <a:lvl1pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl1pPr>
                <a:lvl2pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl2pPr>
                <a:lvl3pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl3pPr>
                <a:lvl4pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl4pPr>
                <a:lvl5pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl5pPr>
                <a:lvl6pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl6pPr>
                <a:lvl7pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl7pPr>
                <a:lvl8pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl8pPr>
                <a:lvl9pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl9pPr>
              </a:lstStyle>
              <a:p>
                <a:pPr lvl="0"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master text styles</a:t>
                </a:r>
              </a:p>
            </p:txBody>
          </p:sp>
          <xsl:call-template name ="masterFooter">
            <xsl:with-param name="smasterName" select="$MasterName"/>
            <xsl:with-param name ="startNo" select ="5"/>
          </xsl:call-template>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>

  <xsl:template name="InsertSlideLayout9">
    <xsl:param name ="MasterName" />
    <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
					 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
					 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="picTx" preserve="1">
      <p:cSld name="Picture with Caption">
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
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Title 1"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph type="title"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>

              <a:xfrm>
                <a:off x="1792288" y="4800600"/>
                <a:ext cx="5486400" cy="566738"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle>
                <a:lvl1pPr>
                  <a:defRPr/>
                </a:lvl1pPr>
              </a:lstStyle>
              <a:p>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master title style</a:t>
                </a:r>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="3" name="Picture Placeholder 2"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph type="pic" idx="1"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="1792288" y="612775"/>
                <a:ext cx="5486400" cy="4114800"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle>
                <a:lvl1pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl1pPr>
                <a:lvl2pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl2pPr>
                <a:lvl3pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl3pPr>
                <a:lvl4pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl4pPr>
                <a:lvl5pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl5pPr>
                <a:lvl6pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl6pPr>
                <a:lvl7pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl7pPr>
                <a:lvl8pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl8pPr>
                <a:lvl9pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl9pPr>
              </a:lstStyle>
              <a:p>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="4" name="Text Placeholder 3"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph type="body" sz="half" idx="2"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="1792288" y="5367338"/>
                <a:ext cx="5486400" cy="804862"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle>
                <a:lvl1pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl1pPr>
                <a:lvl2pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl2pPr>
                <a:lvl3pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl3pPr>
                <a:lvl4pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl4pPr>
                <a:lvl5pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl5pPr>
                <a:lvl6pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl6pPr>
                <a:lvl7pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl7pPr>
                <a:lvl8pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl8pPr>
                <a:lvl9pPr>
                  <a:buNone/>
                  <a:defRPr/>
                </a:lvl9pPr>
              </a:lstStyle>
              <a:p>
                <a:pPr lvl="0"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master text styles</a:t>
                </a:r>
              </a:p>
            </p:txBody>
          </p:sp>
          <xsl:call-template name ="masterFooter">
            <xsl:with-param name="smasterName" select="$MasterName"/>
            <xsl:with-param name ="startNo" select ="5"/>
          </xsl:call-template>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>

  <xsl:template name="InsertSlideLayout10">
    <xsl:param name ="MasterName" />
    <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
					 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
					 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
					 type="vertTx" preserve="1">
      <p:cSld name="Title and Vertical Text">
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
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Title 1"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph type="title"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="457200" y="274638"/>
                <a:ext cx="8229600" cy="1143000"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master title style</a:t>
                </a:r>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="3" name="Vertical Text Placeholder 2"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph type="body" orient="vert" idx="1"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="457200" y="274638"/>
                <a:ext cx="6019800" cy="5851525"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr vert="eaVert"/>
              <a:lstStyle/>
              <a:p>
                <a:pPr lvl="0"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master text styles</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="1"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Second level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="2"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Third level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="3"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Fourth level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="4"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Fifth level</a:t>
                </a:r>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <xsl:call-template name ="masterFooter">
            <xsl:with-param name="smasterName" select="$MasterName"/>
            <xsl:with-param name ="startNo" select ="4"/>
          </xsl:call-template>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>

  <xsl:template name="InsertSlideLayout11">
    <xsl:param name ="MasterName" />
    <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
					 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
					 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
					 type="vertTitleAndTx" preserve="1">
      <p:cSld name="Vertical Title and Text">
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
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Vertical Title 1"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph type="title" orient="vert"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="6629400" y="274638"/>
                <a:ext cx="2057400" cy="5851525"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr vert="eaVert"/>
              <a:lstStyle/>
              <a:p>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master title style</a:t>
                </a:r>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="3" name="Vertical Text Placeholder 2"/>
              <p:cNvSpPr>
                <a:spLocks noGrp="1"/>
              </p:cNvSpPr>
              <p:nvPr>
                <p:ph type="body" orient="vert" idx="1"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="457200" y="274638"/>
                <a:ext cx="6019800" cy="5851525"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr vert="eaVert"/>
              <a:lstStyle/>
              <a:p>
                <a:pPr lvl="0"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Click to edit Master text styles</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="1"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Second level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="2"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Third level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="3"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Fourth level</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr lvl="4"/>
                <a:r>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>Fifth level</a:t>
                </a:r>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <xsl:call-template name ="masterFooter">
            <xsl:with-param name="smasterName" select="$MasterName"/>
            <xsl:with-param name ="startNo" select ="4"/>
          </xsl:call-template>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>
  <!-- Templated for Slide Layouts - Added by lohith.ar-->

  <!-- Templates for Slide Layout relations -->

  <xsl:template name="InsertLayoutRel">
    <xsl:param name ="slideMaasterNo"/>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1"
			  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster">
        <xsl:attribute name ="Target">
          <xsl:value-of select ="concat('../slideMasters/slideMaster',$slideMaasterNo,'.xml')"/>
        </xsl:attribute>
      </Relationship >
    </Relationships>
  </xsl:template>

  <!-- Templates for Slide Layout relations -->

  <xsl:template name ="masterFooter">
    <xsl:param name ="smasterName" />
    <xsl:param name ="startNo"/>
    <xsl:if test="draw:frame/@presentation:class='date-time'">
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr >
            <xsl:attribute name="id">
              <xsl:value-of select ="$startNo"/>
            </xsl:attribute>
            <xsl:attribute name="name">
              <xsl:value-of select ="concat('Date Placeholder ',$startNo -1 )"/>
            </xsl:attribute>
          </p:cNvPr >
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="dt" sz="half" idx="10"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <xsl:for-each select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name=$smasterName]/draw:frame[@presentation:class='date-time']">
              <xsl:choose>
                <xsl:when test="./draw:text-box/text:p/presentation:date-time or ./draw:text-box/text:p/text:span/presentation:date-time" >
                  <a:fld >
                    <xsl:attribute name ="id">
                      <xsl:value-of select ="'{11859447-1FF1-4571-A7A1-75F1CAF7F5D9}'"/>
                    </xsl:attribute>
                    <xsl:attribute name ="type">
                      <xsl:value-of select ="'datetime1'"/>
                    </xsl:attribute>
                    <a:rPr lang="en-US" smtClean="0"/>
                    <a:t> </a:t>
                  </a:fld>
                  <a:endParaRPr lang="en-US"/>
                </xsl:when>
                <xsl:when test="./draw:text-box/text:p/text:date">
                  <a:r>
                    <a:t>
                      <xsl:for-each select="./draw:text-box/text:p/text:date">
                        <xsl:value-of select="."/>
                      </xsl:for-each>
                    </a:t>
                  </a:r >
                </xsl:when>
                <xsl:otherwise >
                  <a:r>
                    <a:t>
                      <xsl:for-each select="./draw:text-box/text:p/text:span">
                        <xsl:value-of select="."/>
                      </xsl:for-each>
                    </a:t>
                  </a:r >
                </xsl:otherwise>
              </xsl:choose>
            </xsl:for-each>
            <xsl:if test ="not(document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name=$smasterName]/draw:frame[@presentation:class='date-time'])">
              <a:endParaRPr/>
            </xsl:if>
          </a:p>
        </p:txBody>
      </p:sp>
    </xsl:if>
    <xsl:if test="draw:frame/@presentation:class='footer'">
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr >
            <xsl:attribute name="id">
              <xsl:value-of select ="$startNo + 1"/>
            </xsl:attribute>
            <xsl:attribute name="name">
              <xsl:value-of select ="concat('Footer Placeholder ',$startNo )"/>
            </xsl:attribute>
          </p:cNvPr >
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="ftr" sz="quarter" idx="11"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <xsl:if test ="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name=$smasterName]/draw:frame[@presentation:class='footer']">
              <a:r>
                <a:rPr lang="en-US" dirty="0" smtClean="0"/>
                <a:t>
                  <xsl:for-each select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name=$smasterName]/draw:frame[@presentation:class='footer']">
                    <xsl:for-each select="./draw:text-box/text:p/text:span">
                      <xsl:value-of select="."/>
                    </xsl:for-each>
                  </xsl:for-each>
                </a:t>
              </a:r>
              <a:endParaRPr lang="en-US"/>
            </xsl:if>
            <xsl:if test ="not(document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name=$smasterName]/draw:frame[@presentation:class='footer'])">
              <a:endParaRPr/>
            </xsl:if>
          </a:p>
        </p:txBody>
      </p:sp>
    </xsl:if>
    <xsl:if test="draw:frame/@presentation:class='page-number'">
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr>
            <xsl:attribute name="id">
              <xsl:value-of select ="$startNo+2"/>
            </xsl:attribute>
            <xsl:attribute name="name">
              <xsl:value-of select ="concat('Slide Number Placeholder ',$startNo + 1 )"/>
            </xsl:attribute>
          </p:cNvPr >
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="sldNum" sz="quarter" idx="12"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <xsl:for-each select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name=$smasterName]/draw:frame[@presentation:class='page-number']">
              <xsl:choose>
                <xsl:when test="./draw:text-box/text:p/text:page-number" >
                  <a:fld >
                    <xsl:attribute name ="id">
                      <xsl:value-of select ="'{763D1470-AB83-4C4C-B3B3-7F0C9DC8E8D6}'"/>
                    </xsl:attribute>
                    <xsl:attribute name ="type">
                      <xsl:value-of select ="'slidenum'"/>
                    </xsl:attribute>
                    <a:rPr lang="en-US" smtClean="0"/>
                    <a:t>
                      <xsl:value-of select="."/>
                    </a:t>
                  </a:fld>
                  <a:endParaRPr lang="en-US"/>
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
            </xsl:for-each>
            <xsl:if test ="not(document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name=$smasterName]/draw:frame[@presentation:class='page-number'])">
              <a:endParaRPr/>
            </xsl:if>
          </a:p>
        </p:txBody>
      </p:sp>
    </xsl:if>
  </xsl:template>
</xsl:stylesheet>

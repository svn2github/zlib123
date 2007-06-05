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
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <!-- Templated for Slide Layouts - Added by lohith.ar -->
  <xsl:template name="InsertSlideLayout1">
    <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
                 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
                 type="title" preserve="1">
      <p:cSld name="Title Slide">
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
                <a:off x="685800" y="2130425"/>
                <a:ext cx="7772400" cy="1470025"/>
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
                <a:off x="1371600" y="3886200"/>
                <a:ext cx="6400800" cy="1752600"/>
              </a:xfrm>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle>
                <a:lvl1pPr marL="0" indent="0" algn="ctr">
                  <a:buNone/>
                  <a:defRPr>
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
                  </a:defRPr>
                </a:lvl1pPr>
                <a:lvl2pPr marL="457200" indent="0" algn="ctr">
                  <a:buNone/>
                  <a:defRPr>
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
                  </a:defRPr>
                </a:lvl2pPr>
                <a:lvl3pPr marL="914400" indent="0" algn="ctr">
                  <a:buNone/>
                  <a:defRPr>
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
                  </a:defRPr>
                </a:lvl3pPr>
                <a:lvl4pPr marL="1371600" indent="0" algn="ctr">
                  <a:buNone/>
                  <a:defRPr>
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
                  </a:defRPr>
                </a:lvl4pPr>
                <a:lvl5pPr marL="1828800" indent="0" algn="ctr">
                  <a:buNone/>
                  <a:defRPr>
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
                  </a:defRPr>
                </a:lvl5pPr>
                <a:lvl6pPr marL="2286000" indent="0" algn="ctr">
                  <a:buNone/>
                  <a:defRPr>
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
                  </a:defRPr>
                </a:lvl6pPr>
                <a:lvl7pPr marL="2743200" indent="0" algn="ctr">
                  <a:buNone/>
                  <a:defRPr>
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
                  </a:defRPr>
                </a:lvl7pPr>
                <a:lvl8pPr marL="3200400" indent="0" algn="ctr">
                  <a:buNone/>
                  <a:defRPr>
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
                  </a:defRPr>
                </a:lvl8pPr>
                <a:lvl9pPr marL="3657600" indent="0" algn="ctr">
                  <a:buNone/>
                  <a:defRPr>
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
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
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="4" name="Date Placeholder 3"/>
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
                <a:fld type="datetimeFigureOut">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{11859447-1FF1-4571-A7A1-75F1CAF7F5D9}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>4/9/2007</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="5" name="Footer Placeholder 4"/>
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
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="6" name="Slide Number Placeholder 5"/>
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
                <a:fld type="slidenum">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{519AE663-802C-4E82-9DD4-183A9F5CB787}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>‹#›</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>

  <xsl:template name="InsertSlideLayout2">
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
            <p:spPr/>
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
            <p:spPr/>
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
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="4" name="Date Placeholder 3"/>
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
                <a:fld type="datetimeFigureOut">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{11859447-1FF1-4571-A7A1-75F1CAF7F5D9}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>4/9/2007</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="5" name="Footer Placeholder 4"/>
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
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="6" name="Slide Number Placeholder 5"/>
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
                <a:fld type="slidenum">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{519AE663-802C-4E82-9DD4-183A9F5CB787}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>‹#›</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>

  <xsl:template name="InsertSlideLayout3">
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
              <a:bodyPr anchor="t"/>
              <a:lstStyle>
                <a:lvl1pPr algn="l">
                  <a:defRPr sz="4000" b="1" cap="all"/>
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
              <a:bodyPr anchor="b"/>
              <a:lstStyle>
                <a:lvl1pPr marL="0" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="2000">
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
                  </a:defRPr>
                </a:lvl1pPr>
                <a:lvl2pPr marL="457200" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1800">
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
                  </a:defRPr>
                </a:lvl2pPr>
                <a:lvl3pPr marL="914400" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1600">
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
                  </a:defRPr>
                </a:lvl3pPr>
                <a:lvl4pPr marL="1371600" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1400">
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
                  </a:defRPr>
                </a:lvl4pPr>
                <a:lvl5pPr marL="1828800" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1400">
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
                  </a:defRPr>
                </a:lvl5pPr>
                <a:lvl6pPr marL="2286000" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1400">
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
                  </a:defRPr>
                </a:lvl6pPr>
                <a:lvl7pPr marL="2743200" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1400">
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
                  </a:defRPr>
                </a:lvl7pPr>
                <a:lvl8pPr marL="3200400" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1400">
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
                  </a:defRPr>
                </a:lvl8pPr>
                <a:lvl9pPr marL="3657600" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1400">
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
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
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="4" name="Date Placeholder 3"/>
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
                <a:fld type="datetimeFigureOut">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{11859447-1FF1-4571-A7A1-75F1CAF7F5D9}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>4/9/2007</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="5" name="Footer Placeholder 4"/>
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
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="6" name="Slide Number Placeholder 5"/>
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
                <a:fld type="slidenum">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{519AE663-802C-4E82-9DD4-183A9F5CB787}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>‹#›</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>

  <xsl:template name="InsertSlideLayout4">
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
            <p:spPr/>
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
                  <a:defRPr sz="2800"/>
                </a:lvl1pPr>
                <a:lvl2pPr>
                  <a:defRPr sz="2400"/>
                </a:lvl2pPr>
                <a:lvl3pPr>
                  <a:defRPr sz="2000"/>
                </a:lvl3pPr>
                <a:lvl4pPr>
                  <a:defRPr sz="1800"/>
                </a:lvl4pPr>
                <a:lvl5pPr>
                  <a:defRPr sz="1800"/>
                </a:lvl5pPr>
                <a:lvl6pPr>
                  <a:defRPr sz="1800"/>
                </a:lvl6pPr>
                <a:lvl7pPr>
                  <a:defRPr sz="1800"/>
                </a:lvl7pPr>
                <a:lvl8pPr>
                  <a:defRPr sz="1800"/>
                </a:lvl8pPr>
                <a:lvl9pPr>
                  <a:defRPr sz="1800"/>
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
                  <a:defRPr sz="2800"/>
                </a:lvl1pPr>
                <a:lvl2pPr>
                  <a:defRPr sz="2400"/>
                </a:lvl2pPr>
                <a:lvl3pPr>
                  <a:defRPr sz="2000"/>
                </a:lvl3pPr>
                <a:lvl4pPr>
                  <a:defRPr sz="1800"/>
                </a:lvl4pPr>
                <a:lvl5pPr>
                  <a:defRPr sz="1800"/>
                </a:lvl5pPr>
                <a:lvl6pPr>
                  <a:defRPr sz="1800"/>
                </a:lvl6pPr>
                <a:lvl7pPr>
                  <a:defRPr sz="1800"/>
                </a:lvl7pPr>
                <a:lvl8pPr>
                  <a:defRPr sz="1800"/>
                </a:lvl8pPr>
                <a:lvl9pPr>
                  <a:defRPr sz="1800"/>
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
              <p:cNvPr id="5" name="Date Placeholder 4"/>
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
                <a:fld type="datetimeFigureOut">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{11859447-1FF1-4571-A7A1-75F1CAF7F5D9}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>4/9/2007</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="6" name="Footer Placeholder 5"/>
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
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="7" name="Slide Number Placeholder 6"/>
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
                <a:fld type="slidenum">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{519AE663-802C-4E82-9DD4-183A9F5CB787}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>‹#›</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>

  <xsl:template name="InsertSlideLayout5">
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
            <p:spPr/>
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
                <a:lvl1pPr marL="0" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="2400" b="1"/>
                </a:lvl1pPr>
                <a:lvl2pPr marL="457200" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="2000" b="1"/>
                </a:lvl2pPr>
                <a:lvl3pPr marL="914400" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1800" b="1"/>
                </a:lvl3pPr>
                <a:lvl4pPr marL="1371600" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1600" b="1"/>
                </a:lvl4pPr>
                <a:lvl5pPr marL="1828800" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1600" b="1"/>
                </a:lvl5pPr>
                <a:lvl6pPr marL="2286000" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1600" b="1"/>
                </a:lvl6pPr>
                <a:lvl7pPr marL="2743200" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1600" b="1"/>
                </a:lvl7pPr>
                <a:lvl8pPr marL="3200400" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1600" b="1"/>
                </a:lvl8pPr>
                <a:lvl9pPr marL="3657600" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1600" b="1"/>
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
                  <a:defRPr sz="2400"/>
                </a:lvl1pPr>
                <a:lvl2pPr>
                  <a:defRPr sz="2000"/>
                </a:lvl2pPr>
                <a:lvl3pPr>
                  <a:defRPr sz="1800"/>
                </a:lvl3pPr>
                <a:lvl4pPr>
                  <a:defRPr sz="1600"/>
                </a:lvl4pPr>
                <a:lvl5pPr>
                  <a:defRPr sz="1600"/>
                </a:lvl5pPr>
                <a:lvl6pPr>
                  <a:defRPr sz="1600"/>
                </a:lvl6pPr>
                <a:lvl7pPr>
                  <a:defRPr sz="1600"/>
                </a:lvl7pPr>
                <a:lvl8pPr>
                  <a:defRPr sz="1600"/>
                </a:lvl8pPr>
                <a:lvl9pPr>
                  <a:defRPr sz="1600"/>
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
              <a:bodyPr anchor="b"/>
              <a:lstStyle>
                <a:lvl1pPr marL="0" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="2400" b="1"/>
                </a:lvl1pPr>
                <a:lvl2pPr marL="457200" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="2000" b="1"/>
                </a:lvl2pPr>
                <a:lvl3pPr marL="914400" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1800" b="1"/>
                </a:lvl3pPr>
                <a:lvl4pPr marL="1371600" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1600" b="1"/>
                </a:lvl4pPr>
                <a:lvl5pPr marL="1828800" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1600" b="1"/>
                </a:lvl5pPr>
                <a:lvl6pPr marL="2286000" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1600" b="1"/>
                </a:lvl6pPr>
                <a:lvl7pPr marL="2743200" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1600" b="1"/>
                </a:lvl7pPr>
                <a:lvl8pPr marL="3200400" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1600" b="1"/>
                </a:lvl8pPr>
                <a:lvl9pPr marL="3657600" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1600" b="1"/>
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
                  <a:defRPr sz="2400"/>
                </a:lvl1pPr>
                <a:lvl2pPr>
                  <a:defRPr sz="2000"/>
                </a:lvl2pPr>
                <a:lvl3pPr>
                  <a:defRPr sz="1800"/>
                </a:lvl3pPr>
                <a:lvl4pPr>
                  <a:defRPr sz="1600"/>
                </a:lvl4pPr>
                <a:lvl5pPr>
                  <a:defRPr sz="1600"/>
                </a:lvl5pPr>
                <a:lvl6pPr>
                  <a:defRPr sz="1600"/>
                </a:lvl6pPr>
                <a:lvl7pPr>
                  <a:defRPr sz="1600"/>
                </a:lvl7pPr>
                <a:lvl8pPr>
                  <a:defRPr sz="1600"/>
                </a:lvl8pPr>
                <a:lvl9pPr>
                  <a:defRPr sz="1600"/>
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
              <p:cNvPr id="7" name="Date Placeholder 6"/>
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
                <a:fld type="datetimeFigureOut">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{11859447-1FF1-4571-A7A1-75F1CAF7F5D9}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>4/9/2007</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="8" name="Footer Placeholder 7"/>
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
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="9" name="Slide Number Placeholder 8"/>
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
                <a:fld type="slidenum">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{519AE663-802C-4E82-9DD4-183A9F5CB787}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>‹#›</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>

  <xsl:template name="InsertSlideLayout6">
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
            <p:spPr/>
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
              <p:cNvPr id="3" name="Date Placeholder 2"/>
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
                <a:fld type="datetimeFigureOut">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{11859447-1FF1-4571-A7A1-75F1CAF7F5D9}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>4/9/2007</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="4" name="Footer Placeholder 3"/>
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
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="5" name="Slide Number Placeholder 4"/>
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
                <a:fld type="slidenum">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{519AE663-802C-4E82-9DD4-183A9F5CB787}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>‹#›</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>

  <xsl:template name="InsertSlideLayout7">
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
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Date Placeholder 1"/>
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
                <a:fld type="datetimeFigureOut">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{11859447-1FF1-4571-A7A1-75F1CAF7F5D9}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>4/9/2007</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="3" name="Footer Placeholder 2"/>
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
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="4" name="Slide Number Placeholder 3"/>
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
                <a:fld type="slidenum">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{519AE663-802C-4E82-9DD4-183A9F5CB787}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>‹#›</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>

  <xsl:template name="InsertSlideLayout8">
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
              <a:bodyPr anchor="b"/>
              <a:lstStyle>
                <a:lvl1pPr algn="l">
                  <a:defRPr sz="2000" b="1"/>
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
                  <a:defRPr sz="3200"/>
                </a:lvl1pPr>
                <a:lvl2pPr>
                  <a:defRPr sz="2800"/>
                </a:lvl2pPr>
                <a:lvl3pPr>
                  <a:defRPr sz="2400"/>
                </a:lvl3pPr>
                <a:lvl4pPr>
                  <a:defRPr sz="2000"/>
                </a:lvl4pPr>
                <a:lvl5pPr>
                  <a:defRPr sz="2000"/>
                </a:lvl5pPr>
                <a:lvl6pPr>
                  <a:defRPr sz="2000"/>
                </a:lvl6pPr>
                <a:lvl7pPr>
                  <a:defRPr sz="2000"/>
                </a:lvl7pPr>
                <a:lvl8pPr>
                  <a:defRPr sz="2000"/>
                </a:lvl8pPr>
                <a:lvl9pPr>
                  <a:defRPr sz="2000"/>
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
                <a:lvl1pPr marL="0" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1400"/>
                </a:lvl1pPr>
                <a:lvl2pPr marL="457200" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1200"/>
                </a:lvl2pPr>
                <a:lvl3pPr marL="914400" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1000"/>
                </a:lvl3pPr>
                <a:lvl4pPr marL="1371600" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="900"/>
                </a:lvl4pPr>
                <a:lvl5pPr marL="1828800" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="900"/>
                </a:lvl5pPr>
                <a:lvl6pPr marL="2286000" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="900"/>
                </a:lvl6pPr>
                <a:lvl7pPr marL="2743200" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="900"/>
                </a:lvl7pPr>
                <a:lvl8pPr marL="3200400" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="900"/>
                </a:lvl8pPr>
                <a:lvl9pPr marL="3657600" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="900"/>
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
              <p:cNvPr id="5" name="Date Placeholder 4"/>
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
                <a:fld type="datetimeFigureOut">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{11859447-1FF1-4571-A7A1-75F1CAF7F5D9}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>4/9/2007</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="6" name="Footer Placeholder 5"/>
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
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="7" name="Slide Number Placeholder 6"/>
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
                <a:fld type="slidenum">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{519AE663-802C-4E82-9DD4-183A9F5CB787}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>‹#›</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>

  <xsl:template name="InsertSlideLayout9">
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
              <a:bodyPr anchor="b"/>
              <a:lstStyle>
                <a:lvl1pPr algn="l">
                  <a:defRPr sz="2000" b="1"/>
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
                <a:lvl1pPr marL="0" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="3200"/>
                </a:lvl1pPr>
                <a:lvl2pPr marL="457200" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="2800"/>
                </a:lvl2pPr>
                <a:lvl3pPr marL="914400" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="2400"/>
                </a:lvl3pPr>
                <a:lvl4pPr marL="1371600" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="2000"/>
                </a:lvl4pPr>
                <a:lvl5pPr marL="1828800" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="2000"/>
                </a:lvl5pPr>
                <a:lvl6pPr marL="2286000" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="2000"/>
                </a:lvl6pPr>
                <a:lvl7pPr marL="2743200" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="2000"/>
                </a:lvl7pPr>
                <a:lvl8pPr marL="3200400" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="2000"/>
                </a:lvl8pPr>
                <a:lvl9pPr marL="3657600" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="2000"/>
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
                <a:lvl1pPr marL="0" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1400"/>
                </a:lvl1pPr>
                <a:lvl2pPr marL="457200" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1200"/>
                </a:lvl2pPr>
                <a:lvl3pPr marL="914400" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="1000"/>
                </a:lvl3pPr>
                <a:lvl4pPr marL="1371600" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="900"/>
                </a:lvl4pPr>
                <a:lvl5pPr marL="1828800" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="900"/>
                </a:lvl5pPr>
                <a:lvl6pPr marL="2286000" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="900"/>
                </a:lvl6pPr>
                <a:lvl7pPr marL="2743200" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="900"/>
                </a:lvl7pPr>
                <a:lvl8pPr marL="3200400" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="900"/>
                </a:lvl8pPr>
                <a:lvl9pPr marL="3657600" indent="0">
                  <a:buNone/>
                  <a:defRPr sz="900"/>
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
              <p:cNvPr id="5" name="Date Placeholder 4"/>
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
                <a:fld type="datetimeFigureOut">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{11859447-1FF1-4571-A7A1-75F1CAF7F5D9}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>4/9/2007</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="6" name="Footer Placeholder 5"/>
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
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="7" name="Slide Number Placeholder 6"/>
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
                <a:fld type="slidenum">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{519AE663-802C-4E82-9DD4-183A9F5CB787}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>‹#›</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>

  <xsl:template name="InsertSlideLayout10">
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
            <p:spPr/>
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
            <p:spPr/>
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
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="4" name="Date Placeholder 3"/>
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
                <a:fld type="datetimeFigureOut">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{11859447-1FF1-4571-A7A1-75F1CAF7F5D9}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>4/9/2007</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="5" name="Footer Placeholder 4"/>
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
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="6" name="Slide Number Placeholder 5"/>
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
                <a:fld type="slidenum">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{519AE663-802C-4E82-9DD4-183A9F5CB787}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>‹#›</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>

  <xsl:template name="InsertSlideLayout11">
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
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="4" name="Date Placeholder 3"/>
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
                <a:fld type="datetimeFigureOut">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{11859447-1FF1-4571-A7A1-75F1CAF7F5D9}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>4/9/2007</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="5" name="Footer Placeholder 4"/>
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
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="6" name="Slide Number Placeholder 5"/>
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
                <a:fld type="slidenum">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{519AE663-802C-4E82-9DD4-183A9F5CB787}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:t>‹#›</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:masterClrMapping/>
      </p:clrMapOvr>
    </p:sldLayout>
  </xsl:template>
  <!-- Templated for Slide Layouts - Added by lohith.ar-->

  <!-- Templates for Slide Layout relations -->

  <xsl:template name="InsertSlideLayout1DotXmlRels">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
    </Relationships>
  </xsl:template>

  <xsl:template name="InsertSlideLayout2DotXmlRels">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
    </Relationships>
  </xsl:template>

  <xsl:template name="InsertSlideLayout3DotXmlRels">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
    </Relationships>
  </xsl:template>

  <xsl:template name="InsertSlideLayout4DotXmlRels">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
    </Relationships>
  </xsl:template>

  <xsl:template name="InsertSlideLayout5DotXmlRels">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
    </Relationships>
  </xsl:template>

  <xsl:template name="InsertSlideLayout6DotXmlRels">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
    </Relationships>
  </xsl:template>

  <xsl:template name="InsertSlideLayout7DotXmlRels">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
    </Relationships>
  </xsl:template>

  <xsl:template name="InsertSlideLayout8DotXmlRels">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
    </Relationships>
  </xsl:template>

  <xsl:template name="InsertSlideLayout9DotXmlRels">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
    </Relationships>
  </xsl:template>

  <xsl:template name="InsertSlideLayout10DotXmlRels">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
    </Relationships>
  </xsl:template>
  <xsl:template name="InsertSlideLayout11DotXmlRels">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
    </Relationships>
  </xsl:template>

  <!-- Templates for Slide Layout relations -->

</xsl:stylesheet>

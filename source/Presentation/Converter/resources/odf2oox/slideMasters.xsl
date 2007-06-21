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
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

  <xsl:template name ="slideMaster1">
    <p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
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
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Title Placeholder 1"/>
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
              <a:prstGeom prst="rect">
                <a:avLst/>
              </a:prstGeom>
            </p:spPr>
            <p:txBody>
              <a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr">
                <a:normAutofit/>
              </a:bodyPr>
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
                <a:off x="457200" y="1600200"/>
                <a:ext cx="8229600" cy="4525963"/>
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst/>
              </a:prstGeom>
            </p:spPr>
            <p:txBody>
              <a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0">
                <a:normAutofit/>
              </a:bodyPr>
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
                <p:ph type="dt" sz="half" idx="2"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="457200" y="6356350"/>
                <a:ext cx="2133600" cy="365125"/>
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst/>
              </a:prstGeom>
            </p:spPr>
            <p:txBody>
              <a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/>
              <a:lstStyle>
                <a:lvl1pPr algn="l">
                  <a:defRPr sz="1200">
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
                  </a:defRPr>
                </a:lvl1pPr>
              </a:lstStyle>
              <a:p>
                <a:fld type="datetimeFigureOut">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{1D8BD707-D9CF-40AE-B4C6-C98DA3205C09}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:pPr/>
                  <a:t>4/5/2007</a:t>
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
                <p:ph type="ftr" sz="quarter" idx="3"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="3124200" y="6356350"/>
                <a:ext cx="2895600" cy="365125"/>
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst/>
              </a:prstGeom>
            </p:spPr>
            <p:txBody>
              <a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/>
              <a:lstStyle>
                <a:lvl1pPr algn="ctr">
                  <a:defRPr sz="1200">
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
                  </a:defRPr>
                </a:lvl1pPr>
              </a:lstStyle>
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
                <p:ph type="sldNum" sz="quarter" idx="4"/>
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="6553200" y="6356350"/>
                <a:ext cx="2133600" cy="365125"/>
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst/>
              </a:prstGeom>
            </p:spPr>
            <p:txBody>
              <a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/>
              <a:lstStyle>
                <a:lvl1pPr algn="r">
                  <a:defRPr sz="1200">
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:tint val="75000"/>
                      </a:schemeClr>
                    </a:solidFill>
                  </a:defRPr>
                </a:lvl1pPr>
              </a:lstStyle>
              <a:p>
                <a:fld  type="slidenum">
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="'{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}'"/>
                  </xsl:attribute>
                  <a:rPr lang="en-US" smtClean="0"/>
                  <a:pPr/>
                  <a:t>‹#›</a:t>
                </a:fld>
                <a:endParaRPr lang="en-US"/>
              </a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
      <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
      <p:sldLayoutIdLst>
        <p:sldLayoutId id="2147483649" r:id="rId1"/>
        <p:sldLayoutId id="2147483650" r:id="rId2"/>
        <p:sldLayoutId id="2147483651" r:id="rId3"/>
        <p:sldLayoutId id="2147483652" r:id="rId4"/>
        <p:sldLayoutId id="2147483653" r:id="rId5"/>
        <p:sldLayoutId id="2147483654" r:id="rId6"/>
        <p:sldLayoutId id="2147483655" r:id="rId7"/>
        <p:sldLayoutId id="2147483656" r:id="rId8"/>
        <p:sldLayoutId id="2147483657" r:id="rId9"/>
        <p:sldLayoutId id="2147483658" r:id="rId10"/>
        <p:sldLayoutId id="2147483659" r:id="rId11"/>
      </p:sldLayoutIdLst>
      <p:txStyles>
        <p:titleStyle>
          <a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
            <a:spcBef>
              <a:spcPct val="0"/>
            </a:spcBef>
            <a:buNone/>
            <a:defRPr sz="4400" kern="1200">
              <a:solidFill>
                <a:schemeClr val="tx1"/>
              </a:solidFill>
              <a:latin typeface="+mj-lt"/>
              <a:ea typeface="+mj-ea"/>
              <a:cs typeface="+mj-cs"/>
            </a:defRPr>
          </a:lvl1pPr>
        </p:titleStyle>
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
    </p:sldMaster>
  </xsl:template>
  <xsl:template name ="slideMaster1Rel">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout8.xml"/>
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout3.xml"/>
      <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout7.xml"/>
      <Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
      <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout6.xml"/>
      <Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout11.xml"/>
      <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout5.xml"/>
      <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout10.xml"/>
      <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout4.xml"/>
      <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout9.xml"/>
    </Relationships>
  </xsl:template >
</xsl:stylesheet>
<?xml version="1.0" encoding="UTF-8"?>
<!--
  * Copyright (c) 2006, Clever Age
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
  *     * Neither the name of Clever Age nor the names of its contributors 
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
  * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
-->
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">

  <xsl:template name="CreateChartFile">
    <xsl:param name="sheetNum"/>

    <xsl:for-each
      select="descendant::draw:frame/draw:object[document(concat(translate(@xlink:href,'./',''),'/content.xml'))/office:document-content/office:body/office:chart][1]">
      <pzip:entry pzip:target="{concat('xl/charts/chart',$sheetNum,'_',position(),'.xml')}">

        <!-- example chart file-->

        <c:chartSpace>
          <c:lang val="pl-PL"/>
          <c:chart>
            <c:title>
              <c:tx>
                <c:rich>
                  <a:bodyPr/>
                  <a:lstStyle/>
                  <a:p>
                    <a:pPr>
                      <a:defRPr sz="1760" b="0" i="0" u="none" strike="noStrike" baseline="0">
                        <a:solidFill>
                          <a:srgbClr val="000000"/>
                        </a:solidFill>
                        <a:latin typeface="Arial"/>
                        <a:ea typeface="Arial"/>
                        <a:cs typeface="Arial"/>
                      </a:defRPr>
                    </a:pPr>
                    <a:r>
                      <a:rPr lang="pl-PL"/>
                      <a:t>Main Title</a:t>
                    </a:r>
                  </a:p>
                </c:rich>
              </c:tx>
              <c:layout>
                <c:manualLayout>
                  <c:xMode val="edge"/>
                  <c:yMode val="edge"/>
                  <c:x val="0.4150333420032648"/>
                  <c:y val="3.0952452921440013E-2"/>
                </c:manualLayout>
              </c:layout>
              <c:spPr>
                <a:noFill/>
                <a:ln w="25400">
                  <a:noFill/>
                </a:ln>
              </c:spPr>
            </c:title>
            <c:plotArea>
              <c:layout>
                <c:manualLayout>
                  <c:layoutTarget val="inner"/>
                  <c:xMode val="edge"/>
                  <c:yMode val="edge"/>
                  <c:x val="0.11111128841032286"/>
                  <c:y val="0.19285759127974164"/>
                  <c:w val="0.72712534327343636"/>
                  <c:h val="0.63571576384803719"/>
                </c:manualLayout>
              </c:layout>
              <c:lineChart>
                <c:grouping val="standard"/>
                <c:ser>
                  <c:idx val="0"/>
                  <c:order val="0"/>
                  <c:spPr>
                    <a:ln w="3175">
                      <a:solidFill>
                        <a:srgbClr val="9999FF"/>
                      </a:solidFill>
                      <a:prstDash val="solid"/>
                    </a:ln>
                  </c:spPr>
                  <c:marker>
                    <c:symbol val="diamond"/>
                    <c:size val="6"/>
                    <c:spPr>
                      <a:solidFill>
                        <a:srgbClr val="9999FF"/>
                      </a:solidFill>
                      <a:ln>
                        <a:solidFill>
                          <a:srgbClr val="000000"/>
                        </a:solidFill>
                        <a:prstDash val="solid"/>
                      </a:ln>
                    </c:spPr>
                  </c:marker>
                  <c:val>
                    <c:numRef>
                      <c:f>
                        <xsl:text>Sheet</xsl:text>
                        <xsl:value-of select="$sheetNum"/>
                        <xsl:text>!$C$2:$G$2</xsl:text>
                      </c:f>
                      <c:numCache>
                        <c:formatCode>General</c:formatCode>
                        <c:ptCount val="5"/>
                        <c:pt idx="0">
                          <c:v>1</c:v>
                        </c:pt>
                        <c:pt idx="1">
                          <c:v>2</c:v>
                        </c:pt>
                        <c:pt idx="2">
                          <c:v>3</c:v>
                        </c:pt>
                        <c:pt idx="3">
                          <c:v>4</c:v>
                        </c:pt>
                        <c:pt idx="4">
                          <c:v>5</c:v>
                        </c:pt>
                      </c:numCache>
                    </c:numRef>
                  </c:val>
                </c:ser>
                <c:ser>
                  <c:idx val="1"/>
                  <c:order val="1"/>
                  <c:spPr>
                    <a:ln w="3175">
                      <a:solidFill>
                        <a:srgbClr val="993366"/>
                      </a:solidFill>
                      <a:prstDash val="solid"/>
                    </a:ln>
                  </c:spPr>
                  <c:marker>
                    <c:symbol val="square"/>
                    <c:size val="6"/>
                    <c:spPr>
                      <a:solidFill>
                        <a:srgbClr val="993366"/>
                      </a:solidFill>
                      <a:ln>
                        <a:solidFill>
                          <a:srgbClr val="000000"/>
                        </a:solidFill>
                        <a:prstDash val="solid"/>
                      </a:ln>
                    </c:spPr>
                  </c:marker>
                  <c:val>
                    <c:numRef>
                      <c:f>
                        <xsl:text>Sheet</xsl:text>
                        <xsl:value-of select="$sheetNum"/>
                        <xsl:text>!$C$2:$G$2</xsl:text>                        
                      </c:f>
                      <c:numCache>
                        <c:formatCode>General</c:formatCode>
                        <c:ptCount val="5"/>
                        <c:pt idx="0">
                          <c:v>13</c:v>
                        </c:pt>
                        <c:pt idx="1">
                          <c:v>21</c:v>
                        </c:pt>
                        <c:pt idx="2">
                          <c:v>22</c:v>
                        </c:pt>
                        <c:pt idx="3">
                          <c:v>21</c:v>
                        </c:pt>
                        <c:pt idx="4">
                          <c:v>11</c:v>
                        </c:pt>
                      </c:numCache>
                    </c:numRef>
                  </c:val>
                </c:ser>
                <c:marker val="1"/>
                <c:axId val="121059200"/>
                <c:axId val="121065856"/>
              </c:lineChart>
              <c:catAx>
                <c:axId val="121059200"/>
                <c:scaling>
                  <c:orientation val="minMax"/>
                </c:scaling>
                <c:axPos val="b"/>
                <c:majorGridlines>
                  <c:spPr>
                    <a:ln w="3175">
                      <a:solidFill>
                        <a:srgbClr val="000000"/>
                      </a:solidFill>
                      <a:prstDash val="solid"/>
                    </a:ln>
                  </c:spPr>
                </c:majorGridlines>
                <c:title>
                  <c:tx>
                    <c:rich>
                      <a:bodyPr/>
                      <a:lstStyle/>
                      <a:p>
                        <a:pPr>
                          <a:defRPr sz="1215" b="0" i="0" u="none" strike="noStrike" baseline="0">
                            <a:solidFill>
                              <a:srgbClr val="000000"/>
                            </a:solidFill>
                            <a:latin typeface="Arial"/>
                            <a:ea typeface="Arial"/>
                            <a:cs typeface="Arial"/>
                          </a:defRPr>
                        </a:pPr>
                        <a:r>
                          <a:rPr lang="pl-PL"/>
                          <a:t>X axis title</a:t>
                        </a:r>
                      </a:p>
                    </c:rich>
                  </c:tx>
                  <c:layout>
                    <c:manualLayout>
                      <c:xMode val="edge"/>
                      <c:yMode val="edge"/>
                      <c:x val="0.41176536293237292"/>
                      <c:y val="0.90238305055582813"/>
                    </c:manualLayout>
                  </c:layout>
                  <c:spPr>
                    <a:noFill/>
                    <a:ln w="25400">
                      <a:noFill/>
                    </a:ln>
                  </c:spPr>
                </c:title>
                <c:numFmt formatCode="General" sourceLinked="1"/>
                <c:tickLblPos val="low"/>
                <c:spPr>
                  <a:ln w="9525">
                    <a:noFill/>
                  </a:ln>
                </c:spPr>
                <c:txPr>
                  <a:bodyPr rot="0" vert="horz"/>
                  <a:lstStyle/>
                  <a:p>
                    <a:pPr>
                      <a:defRPr sz="940" b="0" i="0" u="none" strike="noStrike" baseline="0">
                        <a:solidFill>
                          <a:srgbClr val="000000"/>
                        </a:solidFill>
                        <a:latin typeface="Arial"/>
                        <a:ea typeface="Arial"/>
                        <a:cs typeface="Arial"/>
                      </a:defRPr>
                    </a:pPr>
                    <a:endParaRPr lang="pl-PL"/>
                  </a:p>
                </c:txPr>
                <c:crossAx val="121065856"/>
                <c:crossesAt val="0"/>
                <c:auto val="1"/>
                <c:lblAlgn val="ctr"/>
                <c:lblOffset val="100"/>
                <c:tickLblSkip val="1"/>
                <c:tickMarkSkip val="1"/>
              </c:catAx>
              <c:valAx>
                <c:axId val="121065856"/>
                <c:scaling>
                  <c:orientation val="minMax"/>
                </c:scaling>
                <c:axPos val="l"/>
                <c:majorGridlines>
                  <c:spPr>
                    <a:ln w="3175">
                      <a:solidFill>
                        <a:srgbClr val="000000"/>
                      </a:solidFill>
                      <a:prstDash val="solid"/>
                    </a:ln>
                  </c:spPr>
                </c:majorGridlines>
                <c:title>
                  <c:tx>
                    <c:rich>
                      <a:bodyPr/>
                      <a:lstStyle/>
                      <a:p>
                        <a:pPr>
                          <a:defRPr sz="1215" b="0" i="0" u="none" strike="noStrike" baseline="0">
                            <a:solidFill>
                              <a:srgbClr val="000000"/>
                            </a:solidFill>
                            <a:latin typeface="Arial"/>
                            <a:ea typeface="Arial"/>
                            <a:cs typeface="Arial"/>
                          </a:defRPr>
                        </a:pPr>
                        <a:r>
                          <a:rPr lang="pl-PL"/>
                          <a:t>Y axis title</a:t>
                        </a:r>
                      </a:p>
                    </c:rich>
                  </c:tx>
                  <c:layout>
                    <c:manualLayout>
                      <c:xMode val="edge"/>
                      <c:yMode val="edge"/>
                      <c:x val="2.6143832567134789E-2"/>
                      <c:y val="0.42142955131499099"/>
                    </c:manualLayout>
                  </c:layout>
                  <c:spPr>
                    <a:noFill/>
                    <a:ln w="25400">
                      <a:noFill/>
                    </a:ln>
                  </c:spPr>
                </c:title>
                <c:numFmt formatCode="General" sourceLinked="0"/>
                <c:tickLblPos val="low"/>
                <c:spPr>
                  <a:ln w="9525">
                    <a:noFill/>
                  </a:ln>
                </c:spPr>
                <c:txPr>
                  <a:bodyPr rot="0" vert="horz"/>
                  <a:lstStyle/>
                  <a:p>
                    <a:pPr>
                      <a:defRPr sz="940" b="0" i="0" u="none" strike="noStrike" baseline="0">
                        <a:solidFill>
                          <a:srgbClr val="000000"/>
                        </a:solidFill>
                        <a:latin typeface="Arial"/>
                        <a:ea typeface="Arial"/>
                        <a:cs typeface="Arial"/>
                      </a:defRPr>
                    </a:pPr>
                    <a:endParaRPr lang="pl-PL"/>
                  </a:p>
                </c:txPr>
                <c:crossAx val="121059200"/>
                <c:crosses val="autoZero"/>
                <c:crossBetween val="midCat"/>
              </c:valAx>
              <c:spPr>
                <a:noFill/>
                <a:ln w="3175">
                  <a:solidFill>
                    <a:srgbClr val="000000"/>
                  </a:solidFill>
                  <a:prstDash val="lgDashDot"/>
                </a:ln>
              </c:spPr>
            </c:plotArea>
            <c:legend>
              <c:legendPos val="r"/>
              <c:layout>
                <c:manualLayout>
                  <c:xMode val="edge"/>
                  <c:yMode val="edge"/>
                  <c:x val="0.86438046425089399"/>
                  <c:y val="0.46428679382160021"/>
                  <c:w val="0.12254921515844433"/>
                  <c:h val="9.2857358764320039E-2"/>
                </c:manualLayout>
              </c:layout>
              <c:spPr>
                <a:solidFill>
                  <a:srgbClr val="D9D9D9"/>
                </a:solidFill>
                <a:ln w="3175">
                  <a:solidFill>
                    <a:srgbClr val="000000"/>
                  </a:solidFill>
                  <a:prstDash val="solid"/>
                </a:ln>
              </c:spPr>
              <c:txPr>
                <a:bodyPr/>
                <a:lstStyle/>
                <a:p>
                  <a:pPr>
                    <a:defRPr sz="740" b="0" i="0" u="none" strike="noStrike" baseline="0">
                      <a:solidFill>
                        <a:srgbClr val="000000"/>
                      </a:solidFill>
                      <a:latin typeface="Arial"/>
                      <a:ea typeface="Arial"/>
                      <a:cs typeface="Arial"/>
                    </a:defRPr>
                  </a:pPr>
                  <a:endParaRPr lang="pl-PL"/>
                </a:p>
              </c:txPr>
            </c:legend>
            <c:dispBlanksAs val="gap"/>
          </c:chart>
          <c:spPr>
            <a:solidFill>
              <a:srgbClr val="FFFFFF"/>
            </a:solidFill>
            <a:ln w="3175">
              <a:solidFill>
                <a:srgbClr val="000000"/>
              </a:solidFill>
              <a:prstDash val="solid"/>
            </a:ln>
          </c:spPr>
          <c:txPr>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
              <a:pPr>
                <a:defRPr sz="1100" b="0" i="0" u="none" strike="noStrike" baseline="0">
                  <a:solidFill>
                    <a:srgbClr val="000000"/>
                  </a:solidFill>
                  <a:latin typeface="Calibri"/>
                  <a:ea typeface="Calibri"/>
                  <a:cs typeface="Calibri"/>
                </a:defRPr>
              </a:pPr>
              <a:endParaRPr lang="pl-PL"/>
            </a:p>
          </c:txPr>
          <c:printSettings>
            <c:headerFooter alignWithMargins="0"/>
            <c:pageMargins b="1" l="0.75" r="0.75" t="1" header="0.4921259845" footer="0.4921259845"/>
            <c:pageSetup/>
          </c:printSettings>
        </c:chartSpace>

      </pzip:entry>
    </xsl:for-each>
  </xsl:template>

</xsl:stylesheet>

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
  xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">

  <!-- @Filename: chart.xsl -->
  <!-- @Description: This stylesheet is used for charts conversion -->
  <!-- @Created: 2007-05-24 -->

  <xsl:key name="rows" match="table:table-rows" use="''"/>
  <xsl:key name="header" match="table:table-header-rows" use="''"/>
  <xsl:key name="series" match="chart:series" use="''"/>
  <xsl:key name="style" match="style:style" use="@style:name"/>
  <xsl:key name="chart" match="chart:chart" use="''"/>

  <xsl:template name="CreateChartFile">
    <!-- @Description: Searches for all charts within sheet and creates output chart files. -->
    <!-- @Context: table:table -->

    <xsl:param name="sheetNum"/>
    <!-- (number) sheet number -->

    <xsl:variable name="chart">
      <xsl:for-each select="descendant::draw:frame/draw:object">
        <xsl:choose>
          <xsl:when test="not(document(concat(translate(@xlink:href,'./',''),'/settings.xml')))">
            <xsl:for-each select="document(concat(translate(@xlink:href,'./',''),'/content.xml'))">
              <xsl:choose>
                <xsl:when test="office:document-content/office:body/office:chart">
                  <xsl:text>true</xsl:text>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:text>false</xsl:text>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:for-each>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>false</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
    </xsl:variable>

    <xsl:if test="contains($chart, 'true')">
      <xsl:for-each
        select="descendant::draw:frame/draw:object[document(concat(translate(@xlink:href,'./',''),'/content.xml'))/office:document-content/office:body/office:chart]">

        <xsl:variable name="chartDirectory">
          <xsl:value-of select="translate(@xlink:href,'./','')"/>
        </xsl:variable>

        <xsl:variable name="chartNum">
          <xsl:value-of select="position()"/>
        </xsl:variable>

        <!-- insert chart file content -->
        <xsl:for-each
          select="document(concat(translate(@xlink:href,'./',''),'/content.xml'))/office:document-content/office:body/office:chart">
          <pzip:entry pzip:target="{concat('xl/charts/chart',$sheetNum,'_',$chartNum,'.xml')}">

            <c:chartSpace>
              <c:lang val="pl-PL"/>

              <xsl:for-each select="chart:chart">
                <xsl:call-template name="InsertChart">
                  <xsl:with-param name="chartDirectory" select="$chartDirectory"/>
                </xsl:call-template>
                <!-- chart area properties -->
                <xsl:call-template name="InsertSpPr">
                  <xsl:with-param name="chartDirectory" select="$chartDirectory"/>
                </xsl:call-template>
              </xsl:for-each>

              <c:txPr>
                <a:bodyPr/>
                <a:lstStyle/>
                <a:p>
                  <a:pPr>
                    <a:defRPr sz="1000" b="0" i="0" u="none" strike="noStrike" baseline="0">
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

              <c:printSettings>
                <c:headerFooter alignWithMargins="0"/>
                <c:pageMargins b="1" l="0.75000000000000044" r="0.75000000000000044" t="1"
                  header="0.49212598450000022" footer="0.49212598450000022"/>
                <c:pageSetup/>
              </c:printSettings>
            </c:chartSpace>
          </pzip:entry>

          <xsl:call-template name="CreateChartRelationships">
            <xsl:with-param name="chartDirectory" select="$chartDirectory"/>
            <xsl:with-param name="chartFile">
              <xsl:value-of select="concat('chart',$sheetNum,'_',$chartNum)"/>
            </xsl:with-param>
          </xsl:call-template>

        </xsl:for-each>


      </xsl:for-each>
    </xsl:if>
  </xsl:template>

  <xsl:template name="CreateChartRelationships">
    <xsl:param name="chartFile"/>
    <xsl:param name="chartDirectory"/>

    <!-- check if there is bitmap fill -->
    <xsl:if
      test="/office:document-content/office:automatic-styles/style:style/style:graphic-properties[@draw:fill = 'bitmap' ]">
      <pzip:entry pzip:target="{concat('xl/charts/_rels/',$chartFile,'.xml.rels')}">
        <xsl:call-template name="InsertChartRels">
          <xsl:with-param name="chartDirectory" select="$chartDirectory"/>
        </xsl:call-template>
      </pzip:entry>
    </xsl:if>

  </xsl:template>

  <xsl:template name="InsertChart">
    <!-- @Description: Writes chart definition to output chart file. -->
    <!-- @Context: chart:chart -->
    <xsl:param name="chartDirectory"/>

    <xsl:variable name="chartWidth">
      <xsl:value-of select="substring-before(@svg:width,'cm')"/>
    </xsl:variable>

    <xsl:variable name="chartHeight">
      <xsl:value-of select="substring-before(@svg:height,'cm')"/>
    </xsl:variable>

    <c:chart>
      <xsl:for-each select="chart:title">
        <xsl:call-template name="InsertTitle">
          <xsl:with-param name="chartWidth" select="$chartWidth"/>
          <xsl:with-param name="chartHeight" select="$chartHeight"/>
          <xsl:with-param name="chartDirectory" select="$chartDirectory"/>
        </xsl:call-template>
      </xsl:for-each>

      <xsl:call-template name="InsertView3D"/>

      <xsl:call-template name="InsertPlotArea">
        <xsl:with-param name="chartWidth" select="$chartWidth"/>
        <xsl:with-param name="chartHeight" select="$chartHeight"/>
        <xsl:with-param name="chartDirectory" select="$chartDirectory"/>
      </xsl:call-template>

      <xsl:call-template name="InsertLegend">
        <xsl:with-param name="chartWidth" select="$chartWidth"/>
        <xsl:with-param name="chartHeight" select="$chartHeight"/>
      </xsl:call-template>

      <xsl:call-template name="InsertAdditionalProperties"/>
    </c:chart>
  </xsl:template>

  <xsl:template name="InsertTitle">
    <!-- @Description: Inserts chart title -->
    <!-- @Context: chart:chart -->
    <xsl:param name="chartWidth"/>
    <xsl:param name="chartHeight"/>
    <xsl:param name="chartDirectory"/>

    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr>
            <xsl:for-each select="key('style',@chart:style-name)/style:chart-properties">

              <!-- rotation -->
              <xsl:if test="@style:rotation-angle">
                <xsl:attribute name="rot">
                  <xsl:choose>
                    <!-- 0 deg -->
                    <xsl:when test="@style:rotation-angle = 0">
                      <xsl:text>0</xsl:text>
                    </xsl:when>
                    <!-- (0 ; 180> deg -->
                    <xsl:when test="@style:rotation-angle &lt; 90 and @style:rotation-angle &lt;= 180">
                      <xsl:text>-</xsl:text>
                      <xsl:value-of select="@style:rotation-angle * 60000"/>
                    </xsl:when>
                    <!-- (180 ; 360) deg -->
                    <xsl:otherwise>
                      <xsl:value-of select="(360 - @style:rotation-angle) * 60000"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
              </xsl:if>
              
              <!-- vertically stacked -->
              <xsl:if test="@style:direction = 'ttb' ">
                <xsl:attribute name="vert">
                  <xsl:text>wordArtVert</xsl:text>
                </xsl:attribute>
              </xsl:if>
            </xsl:for-each>
          </a:bodyPr>
          <a:lstStyle/>

          <a:p>
            <a:pPr>
              <a:defRPr sz="1300" b="0" i="0" u="none" strike="noStrike" baseline="0">

                <xsl:for-each select="key('style',@chart:style-name)/style:text-properties">
                  <xsl:call-template name="InsertRunProperties"/>
                </xsl:for-each>

              </a:defRPr>
            </a:pPr>
            <a:r>
              <a:rPr lang="pl-PL"/>
              <a:t>
                <xsl:value-of xml:space="preserve" select="text:p"/>
              </a:t>
            </a:r>
          </a:p>
        </c:rich>
      </c:tx>
      
      <!-- best results are when position is default -->
      
      <!--c:layout>
        <c:manualLayout>
          <c:xMode val="edge"/>
          <c:yMode val="edge"/>

          <xsl:call-template name="InsideChartPosition">
            <xsl:with-param name="chartWidth" select="$chartWidth"/>
            <xsl:with-param name="chartHeight" select="$chartHeight"/>
          </xsl:call-template>

        </c:manualLayout>
      </c:layout-->

      <xsl:call-template name="InsertSpPr">
        <xsl:with-param name="chartDirectory" select="$chartDirectory"/>
      </xsl:call-template>

    </c:title>
  </xsl:template>

  <xsl:template name="InsertPlotArea">
    <!-- @Description: Outputs chart plot area -->
    <!-- @Context: chart:chart -->
    <xsl:param name="chartWidth"/>
    <xsl:param name="chartHeight"/>
    <xsl:param name="chartDirectory"/>

    <xsl:for-each select="chart:plot-area">

      <xsl:variable name="plotXOffset">
        <xsl:value-of select="substring-before(@svg:x,'cm' ) div $chartWidth"/>
      </xsl:variable>

      <xsl:variable name="plotYOffset">
        <xsl:value-of select="substring-before(@svg:y,'cm' ) div $chartHeight "/>
      </xsl:variable>

      <c:plotArea>
        <!-- best results are when position is default -->
        
        <!--c:layout>
          <c:manualLayout>
            <c:layoutTarget val="inner"/>
            <c:xMode val="edge"/>
            <c:yMode val="edge"/>

            <xsl:call-template name="InsideChartPosition">
              <xsl:with-param name="chartWidth" select="$chartWidth"/>
              <xsl:with-param name="chartHeight" select="$chartHeight"/>
            </xsl:call-template-->

            <!--c:x val="0.10452961672473868"/>
          <c:y val="0.24334600760456274"/>
          <c:w val="0.68641114982578355"/>
          <c:h val="0.60836501901140683"/-->
          <!--/c:manualLayout>
        </c:layout-->

        <xsl:for-each select="parent::node()">
          <xsl:call-template name="InsertChartType"/>
        </xsl:for-each>

        <xsl:if test="key('series','')[@chart:class]">
          <xsl:for-each select="parent::node()">
            <xsl:call-template name="InsertSecondaryLineChart"/>
          </xsl:for-each>
        </xsl:if>

        <xsl:if test="not(//chart:chart[attribute::chart:class='chart:ring'])">
          <xsl:if test="not(//chart:chart[attribute::chart:class='chart:circle'])">

            <!-- TO DO secondary axis case -->
            <xsl:for-each select="chart:axis[@chart:name = 'primary-x' ]">
              <xsl:call-template name="InsertAxisX">
                <xsl:with-param name="chartWidth" select="$chartWidth"/>
                <xsl:with-param name="chartHeight" select="$chartHeight"/>
                <xsl:with-param name="chartDirectory" select="$chartDirectory"/>
              </xsl:call-template>
            </xsl:for-each>

            <!-- TO DO secondary axis case -->
            <xsl:for-each select="chart:axis[@chart:name='primary-y']">
              <xsl:call-template name="InsertAxisY">
                <xsl:with-param name="chartWidth" select="$chartWidth"/>
                <xsl:with-param name="chartHeight" select="$chartHeight"/>
                <xsl:with-param name="chartDirectory" select="$chartDirectory"/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:if>
        </xsl:if>
        <!-- for the Radar Chart -->
        <xsl:if
          test="//chart:chart[attribute::chart:class='chart:radar'] and not(boolean(//chart:axis/chart:categories))">
          <xsl:for-each select="chart:axis[@chart:dimension = 'y' ]">
            <xsl:call-template name="InsertCatAx">
              <xsl:with-param name="chartWidth" select="$chartWidth+20"/>
              <xsl:with-param name="chartHeight" select="$chartHeight"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:if>

        <!-- plot area graphic properties -->
        <xsl:for-each select="chart:wall">
          <xsl:call-template name="InsertSpPr">
            <xsl:with-param name="chartDirectory" select="$chartDirectory"/>
          </xsl:call-template>
        </xsl:for-each>

      </c:plotArea>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="InsertLegend">
    <!-- @Description: Inserts chart legend. -->
    <!-- @Context: chart:chart -->
    <xsl:param name="chartWidth"/>
    <xsl:param name="chartHeight"/>

    <xsl:for-each select="chart:legend">
      <c:legend>
        <c:legendPos val="r"/>
        <c:layout>
          <c:manualLayout>
            <c:xMode val="edge"/>
            <c:yMode val="edge"/>

            <xsl:call-template name="InsideChartPosition">
              <xsl:with-param name="chartWidth" select="$chartWidth"/>
              <xsl:with-param name="chartHeight" select="$chartHeight"/>
            </xsl:call-template>

          </c:manualLayout>
        </c:layout>
        <c:spPr>
          <a:noFill/>
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
              <a:defRPr sz="550" b="0" i="0" u="none" strike="noStrike" baseline="0">
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
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="InsertAdditionalProperties">

    <c:dispBlanksAs val="gap"/>
  </xsl:template>

  <xsl:template name="InsertChartType">
    <!-- @Description: Inserts appriopiate type of chart -->
    <!-- @Context: chart:chart -->

    <xsl:choose>

      <xsl:when
        test="@chart:class='chart:bar' and key('style',chart:plot-area/@chart:style-name)/style:chart-properties/@chart:three-dimensional = 'true' ">
        <c:bar3DChart>

          <!-- bar or column chart -->
          <xsl:choose>
            <xsl:when
              test="key('style',chart:plot-area/@chart:style-name)/style:chart-properties/@chart:vertical = 'false' ">
              <c:barDir val="col"/>
            </xsl:when>
            <xsl:otherwise>
              <c:barDir val="bar"/>
            </xsl:otherwise>
          </xsl:choose>

          <c:grouping val="clustered">
            <xsl:call-template name="SetDataGroupingAtribute"/>
          </c:grouping>

          <xsl:call-template name="InsertChartContent"/>

          <!-- set shape -->
          <c:shape val="box">
            <xsl:if
              test="key('style',chart:plot-area/@chart:style-name)/style:chart-properties/@chart:solid-type != 'cuboid' ">
              <xsl:attribute name="val">
                <xsl:value-of
                  select="key('style',chart:plot-area/@chart:style-name)/style:chart-properties/@chart:solid-type"
                />
              </xsl:attribute>
            </xsl:if>
          </c:shape>

        </c:bar3DChart>
      </xsl:when>

      <xsl:when test="@chart:class='chart:bar' ">
        <c:barChart>

          <!-- bar or column chart -->
          <xsl:choose>
            <xsl:when
              test="key('style',chart:plot-area/@chart:style-name)/style:chart-properties/@chart:vertical = 'false' ">
              <c:barDir val="col"/>
            </xsl:when>
            <xsl:otherwise>
              <c:barDir val="bar"/>
            </xsl:otherwise>
          </xsl:choose>

          <c:grouping val="clustered">
            <xsl:call-template name="SetDataGroupingAtribute"/>
          </c:grouping>

          <xsl:call-template name="InsertChartContent"/>

          <!-- set overlap for stacked data charts -->
          <xsl:for-each
            select="key('style',chart:plot-area/@chart:style-name)/style:chart-properties">
            <xsl:if test="@chart:stacked = 'true' or @chart:percentage = 'true' ">
              <c:overlap val="100"/>
            </xsl:if>
          </xsl:for-each>

        </c:barChart>
      </xsl:when>

      <xsl:when
        test="@chart:class='chart:line' and key('style',chart:plot-area/@chart:style-name)/style:chart-properties/@chart:three-dimensional = 'true' ">
        <c:line3DChart>
          <c:grouping val="standard"/>

          <xsl:call-template name="InsertChartContent"/>
        </c:line3DChart>
      </xsl:when>

      <xsl:when test="@chart:class='chart:line' ">
        <c:lineChart>

          <c:grouping val="standard">
            <xsl:call-template name="SetDataGroupingAtribute"/>
          </c:grouping>

          <xsl:call-template name="InsertChartContent"/>
        </c:lineChart>
      </xsl:when>

      <xsl:when
        test="@chart:class='chart:area' and key('style',chart:plot-area/@chart:style-name)/style:chart-properties/@chart:three-dimensional = 'true' ">
        <c:area3DChart>

          <c:grouping val="standard">
            <xsl:call-template name="SetDataGroupingAtribute"/>
          </c:grouping>

          <xsl:call-template name="InsertChartContent"/>
        </c:area3DChart>
      </xsl:when>

      <xsl:when test="@chart:class='chart:area' ">
        <c:areaChart>

          <c:grouping val="standard">
            <xsl:call-template name="SetDataGroupingAtribute"/>
          </c:grouping>

          <xsl:call-template name="InsertChartContent"/>
        </c:areaChart>
      </xsl:when>

      <xsl:when
        test="@chart:class='chart:circle' and key('style',chart:plot-area/@chart:style-name)/style:chart-properties/@chart:three-dimensional = 'true' ">
        <c:pie3DChart>
          <c:varyColors val="1"/>
          <xsl:call-template name="InsertChartContent"/>
        </c:pie3DChart>
      </xsl:when>

      <xsl:when test="@chart:class='chart:circle' ">
        <c:pieChart>
          <c:varyColors val="1"/>
          <xsl:call-template name="InsertChartContent"/>
        </c:pieChart>
      </xsl:when>

      <xsl:when test="@chart:class='chart:ring' ">
        <c:doughnutChart>
          <c:varyColors val="1"/>
          <xsl:call-template name="InsertChartContent"/>
          <c:firstSliceAng val="0"/>
          <c:holeSize val="50"/>
        </c:doughnutChart>
      </xsl:when>

      <xsl:when test="@chart:class='chart:radar' ">
        <c:radarChart>
          <c:radarStyle val="marker"/>
          <xsl:call-template name="InsertChartContent"/>
        </c:radarChart>
      </xsl:when>

      <!-- making problems at this time -->
      <!--<xsl:when test="@chart:class='chart:scatter' ">
        <c:scatterChart>
          <c:scatterStyle val="lineMarker"/>
          <xsl:call-template name="InsertChartContent"/>
        </c:scatterChart>
      </xsl:when>
      
      <xsl:when test="@chart:class='chart:stock' ">
        <c:stockChart>
          <xsl:call-template name="InsertChartContent"/>
        </c:stockChart>
      </xsl:when>-->

      <!-- temporary otherwise eventually none -->
      <xsl:otherwise>
        <c:barChart>
          <c:barDir val="col"/>
          <c:grouping val="clustered"/>
          <xsl:call-template name="InsertChartContent"/>
        </c:barChart>
      </xsl:otherwise>

    </xsl:choose>

  </xsl:template>

  <xsl:template name="InsertChartContent">
    <!-- @Description: Inserts chart content -->
    <!-- @Context: chart:chart -->

    <xsl:variable name="numSeries">
      <!-- (number) number of series inside chart -->
      <xsl:value-of select="count(key('series','')[not(@chart:class='chart:line')])"/>
    </xsl:variable>

    <xsl:variable name="numPoints">
      <!-- (number) maximum number of data point -->
      <xsl:value-of select="count(key('rows','')/table:table-row)"/>
    </xsl:variable>

    <xsl:variable name="reverseCategories">
      <xsl:for-each select="key('style',chart:plot-area/@chart:style-name)/style:chart-properties">
        <xsl:choose>
          <!-- reverse categories for: (pie charts) or (ring charts) or (2D bar charts stacked or percentage) -->
          <xsl:when
            test="(key('chart','')/@chart:class = 'chart:circle' ) or (key('chart','')/@chart:class = 'chart:ring' ) or 
            (@chart:vertical = 'true' and @chart:three-dimensional = 'false' )">
            <xsl:text>true</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>false</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
    </xsl:variable>

    <xsl:variable name="reverseSeries">
      <xsl:for-each select="key('style',chart:plot-area/@chart:style-name)/style:chart-properties">
        <xsl:choose>
          <!-- reverse series for: (2D normal bar chart) or (2D normal area chart) or (ring:chart) -->
          <xsl:when
            test="(@chart:vertical = 'true' and not(@chart:stacked = 'true' or @chart:percentage = 'true' ) and @chart:three-dimensional = 'false' ) or 
            (key('chart','')/@chart:class = 'chart:area' and not(@chart:stacked = 'true' or @chart:percentage = 'true' ) and @chart:three-dimensional = 'false' ) or
            (key('chart','')/@chart:class = 'chart:ring')">
            <xsl:text>true</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>false</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
    </xsl:variable>

    <xsl:for-each select="key('rows','')">
      <xsl:call-template name="InsertSeries">
        <xsl:with-param name="numSeries" select="$numSeries"/>
        <xsl:with-param name="numPoints" select="$numPoints"/>
        <xsl:with-param name="reverseCategories" select="$reverseCategories"/>
        <xsl:with-param name="reverseSeries" select="$reverseSeries"/>
      </xsl:call-template>
    </xsl:for-each>


    <c:axId val="110226048"/>
    <c:axId val="110498176"/>

  </xsl:template>

  <xsl:template name="InsertAxisX">
    <!-- @Description: Inserts X-Axis -->
    <!-- @Context: chart:chart-axis -->
    <xsl:param name="chartWidth"/>
    <xsl:param name="chartHeight"/>
    <xsl:param name="chartDirectory"/>

    <c:catAx>
      <c:axId val="110226048"/>
      <c:scaling>
        <c:orientation val="minMax"/>
      </c:scaling>
      <c:axPos val="b"/>

      <xsl:if
        test="key('style',@chart:style-name)/style:chart-properties/@chart:display-label = 'true' ">
        <xsl:for-each select="chart:title">
          <xsl:call-template name="InsertTitle">
            <xsl:with-param name="chartWidth" select="$chartWidth"/>
            <xsl:with-param name="chartHeight" select="$chartHeight"/>
            <xsl:with-param name="chartDirectory" select="$chartDirectory"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:if>

      <c:numFmt formatCode="General" sourceLinked="1"/>

      <c:tickLblPos val="low">
        <xsl:if
          test="key('style',@chart:style-name)/style:chart-properties/@chart:display-label = 'false' ">
          <xsl:attribute name="val">
            <xsl:text>none</xsl:text>
          </xsl:attribute>
        </xsl:if>
      </c:tickLblPos>

      <xsl:if
        test="key('style',@chart:style-name)/style:chart-properties/@chart:display-label = 'true' ">
        <c:spPr>
          <a:ln w="3175">
            <a:solidFill>
              <a:srgbClr val="000000"/>
            </a:solidFill>
            <a:prstDash val="solid"/>
          </a:ln>
        </c:spPr>

        <!--xsl:for-each select="chart:title">
          <xsl:call-template name="InsertSpPr">
            <xsl:with-param name="chartDirectory" select="$chartDirectory"/>
          </xsl:call-template>
        </xsl:for-each-->

        <c:txPr>
          <a:bodyPr rot="0" vert="horz"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr>
              <a:defRPr>

                <xsl:for-each select="key('style',@chart:style-name)/style:text-properties">
                  <!-- template common with text-box-->
                  <xsl:call-template name="InsertRunProperties"/>
                </xsl:for-each>

              </a:defRPr>
            </a:pPr>
            <a:endParaRPr lang="pl-PL"/>
          </a:p>
        </c:txPr>
      </xsl:if>

      <c:crossAx val="110498176"/>
      <c:crossesAt val="0"/>
      <c:auto val="1"/>
      <c:lblAlgn val="ctr"/>
      <c:lblOffset val="100"/>
      <c:tickLblSkip val="1"/>
      <c:tickMarkSkip val="1"/>
    </c:catAx>
  </xsl:template>

  <xsl:template name="InsertAxisY">
    <!-- @Description: Inserts Y-Axis -->
    <!-- @Context: chart:chart -->
    <xsl:param name="chartWidth"/>
    <xsl:param name="chartHeight"/>
    <xsl:param name="chartDirectory"/>

    <c:valAx>
      <c:axId val="110498176"/>
      <c:scaling>
        <c:orientation val="minMax"/>
      </c:scaling>

      <c:axPos val="l"/>

      <xsl:if
        test="key('style',@chart:style-name)/style:chart-properties/@chart:display-label = 'true' ">
        <xsl:for-each select="chart:title">
          <xsl:call-template name="InsertTitle">
            <xsl:with-param name="chartWidth" select="$chartWidth"/>
            <xsl:with-param name="chartHeight" select="$chartHeight"/>
            <xsl:with-param name="chartDirectory" select="$chartDirectory"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:if>

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

      <!--xsl:call-template name="InsertTitle">
        <xsl:with-param name="chartWidth" select="$chartWidth"/>
        <xsl:with-param name="chartHeight" select="$chartHeight"/>
      </xsl:call-template-->

      <c:numFmt formatCode="General" sourceLinked="0"/>

      <c:tickLblPos val="low">
        <xsl:if
          test="key('style',@chart:style-name)/style:chart-properties/@chart:display-label = 'false' ">
          <xsl:attribute name="val">
            <xsl:text>none</xsl:text>
          </xsl:attribute>
        </xsl:if>
      </c:tickLblPos>

      <c:spPr>
        <a:ln w="3175">
          <a:solidFill>
            <a:srgbClr val="000000"/>
          </a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
      </c:spPr>
      <c:txPr>
        <a:bodyPr rot="0" vert="horz"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr>
            <a:defRPr sz="700" b="0" i="0" u="none" strike="noStrike" baseline="0">

              <xsl:for-each select="key('style',@chart:style-name)/style:text-properties">
                <!-- template common with text-box-->
                <xsl:call-template name="InsertRunProperties"/>
              </xsl:for-each>

            </a:defRPr>
          </a:pPr>
          <a:endParaRPr lang="pl-PL"/>
        </a:p>
      </c:txPr>
      <c:crossAx val="110226048"/>
      <c:crosses val="autoZero"/>

      <!-- cross type -->
      <xsl:choose>
        <xsl:when
          test="key('chart','')/@chart:class='chart:area' or key('chart','')/@chart:class='chart:line' ">
          <c:crossBetween val="midCat"/>
        </xsl:when>
        <xsl:otherwise>
          <c:crossBetween val="between"/>
        </xsl:otherwise>
      </xsl:choose>


    </c:valAx>
  </xsl:template>

  <xsl:template name="InsertSeries">
    <!-- @Description: Outputs chart series and their values -->
    <!-- @Context: table:table-rows -->

    <xsl:param name="numSeries"/>
    <!-- (number) number of series inside chart -->
    <xsl:param name="numPoints"/>
    <!-- (number) maximum number of data point -->
    <xsl:param name="count" select="0"/>
    <!-- (number) loop counter -->
    <xsl:param name="reverseSeries"/>
    <!-- (string) is chart vertically aligned -->
    <xsl:param name="reverseCategories"/>

    <xsl:variable name="number">
      <xsl:choose>
        <xsl:when test="$reverseSeries = 'true' ">
          <xsl:value-of select="$numSeries - $count"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$count + 1"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="$count &lt; $numSeries">
        <xsl:choose>
          <!-- if therw is secondary axis ommit its series for now -->
          <xsl:when test="key('series','')[position() = $number and @chart:attached-axis]">
            <xsl:if
              test="key('series','')[position() = $number and @chart:attached-axis = 'primary-y']">
              <xsl:call-template name="Series">
                <xsl:with-param name="numSeries" select="$numSeries"/>
                <xsl:with-param name="numPoints" select="$numPoints"/>
                <xsl:with-param name="count" select="$count"/>
                <xsl:with-param name="reverseSeries" select="$reverseSeries"/>
                <xsl:with-param name="reverseCategories" select="$reverseCategories"/>
              </xsl:call-template>
            </xsl:if>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="Series">
              <xsl:with-param name="numSeries" select="$numSeries"/>
              <xsl:with-param name="numPoints" select="$numPoints"/>
              <xsl:with-param name="count" select="$count"/>
              <xsl:with-param name="reverseSeries" select="$reverseSeries"/>
              <xsl:with-param name="reverseCategories" select="$reverseCategories"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>

        <xsl:call-template name="InsertSeries">
          <xsl:with-param name="numSeries" select="$numSeries"/>
          <xsl:with-param name="numPoints" select="$numPoints"/>
          <xsl:with-param name="count" select="$count + 1"/>
          <xsl:with-param name="reverseSeries" select="$reverseSeries"/>
          <xsl:with-param name="reverseCategories" select="$reverseCategories"/>
        </xsl:call-template>
      </xsl:when>
    </xsl:choose>

  </xsl:template>

  <xsl:template name="Series">
    <xsl:param name="numSeries"/>
    <xsl:param name="numPoints"/>
    <xsl:param name="count"/>
    <xsl:param name="reverseSeries"/>
    <xsl:param name="reverseCategories"/>

    <!-- count series from backwards if reverseSeries is = "true" -->
    <xsl:variable name="number">
      <xsl:choose>
        <xsl:when test="$reverseSeries = 'true' ">
          <xsl:value-of select="$numSeries - $count"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$count + 1"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="chartType">
      <xsl:value-of select="ancestor::chart:chart/@chart:class"/>
    </xsl:variable>

    <xsl:variable name="styleName">
      <!-- (string) series style name -->
      <xsl:value-of select="key('series','')[position() = $number]/@chart:style-name"/>
    </xsl:variable>

    <c:ser>
      <c:idx val="{$count}"/>
      <c:order val="{$count}"/>

      <!-- series name -->
      <c:tx>
        <c:v>
          <xsl:choose>
            <xsl:when test="$reverseSeries = 'true' ">
              <xsl:value-of
                select="key('header','')/table:table-row/table:table-cell[$numSeries + 1 - $count]"
              />
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="key('header','')/table:table-row/table:table-cell[$count + 2]"/>
            </xsl:otherwise>
          </xsl:choose>
        </c:v>
      </c:tx>

      <!-- insert shape properties -->
      <xsl:if test="$chartType != 'chart:ring' ">
        <xsl:call-template name="InsertShapeProperties">
          <xsl:with-param name="styleName" select="$styleName"/>
        </xsl:call-template>
      </xsl:if>

      <!-- insert this series data points shape properties -->
      <xsl:choose>
        <xsl:when test="$chartType = 'chart:ring' ">
          <xsl:for-each select="key('series','')[last()]">
            <xsl:call-template name="InsertRingPointsShapeProperties">
              <xsl:with-param name="totalPoints" select="$numPoints"/>
              <xsl:with-param name="series" select="$number"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>

        <xsl:when test="$reverseCategories = 'true' ">
          <xsl:for-each select="key('series','')[position() = $number]/child::node()[last()]">
            <xsl:call-template name="InsertDataPointsShapeProperties">
              <xsl:with-param name="parentStyleName" select="$styleName"/>
              <xsl:with-param name="reverse" select=" 'true' "/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>

        <xsl:otherwise>
          <xsl:for-each select="key('series','')[position() = $number]/child::node()[1]">
            <xsl:call-template name="InsertDataPointsShapeProperties">
              <xsl:with-param name="parentStyleName" select="$styleName"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>

      <!-- marker type -->
      <!-- if line chart or radar chart or bar chart with lines -->
      <xsl:if
        test="$chartType = 'chart:line' or $chartType = 'chart:radar' or ancestor::chart:chart/chart:plot-area/chart:series[position() = $number]/@chart:class = 'chart:line'">
        <xsl:for-each select="ancestor::chart:chart/chart:plot-area">
          <xsl:choose>
            <!-- when plot-area has 'no symbol' property -->
            <xsl:when
              test="key('style',@chart:style-name )/style:chart-properties/@chart:symbol-type = 'none' ">
              <c:marker>
                <c:symbol val="none"/>
              </c:marker>
            </xsl:when>
            <xsl:otherwise>
              <!-- if this series has 'no-symbol' property -->
              <xsl:if
                test="key('style',chart:series[position() = $number]/@chart:style-name )/style:chart-properties/@chart:symbol-type = 'none' ">
                <c:marker>
                  <c:symbol val="none"/>
                </c:marker>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:if>

      <!-- insert data categories -->
      <c:cat>
        <c:strLit>
          <c:ptCount val="{$numPoints}"/>
          <xsl:choose>
            <xsl:when test="$reverseCategories = 'true' ">
              <xsl:for-each select="key('rows','')">
                <xsl:call-template name="InsertCategoriesReverse">
                  <xsl:with-param name="numCategories" select="$numPoints"/>
                </xsl:call-template>
              </xsl:for-each>
            </xsl:when>

            <xsl:otherwise>
              <xsl:for-each select="key('rows','')">
                <xsl:call-template name="InsertCategories">
                  <xsl:with-param name="numCategories" select="$numPoints"/>
                </xsl:call-template>
              </xsl:for-each>
            </xsl:otherwise>
          </xsl:choose>
        </c:strLit>
      </c:cat>

      <!-- series values -->
      <c:val>
        <c:numRef>
          <!-- TO DO: reference to sheet cell -->
          <!-- i.e. <c:f>Sheet1!$D$3:$D$4</c:f> -->
          <c:numCache>
            <c:formatCode>General</c:formatCode>
            <c:ptCount val="{$numPoints}"/>

            <!-- number of this series -->
            <xsl:variable name="thisSeries">
              <xsl:choose>
                <!-- when $reverseSeries = 'true' then count backwards -->
                <xsl:when test="$reverseSeries = 'true' ">
                  <xsl:value-of select="$numSeries - 1 - $count"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$count"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>

            <xsl:for-each select="ancestor::chart:chart">
              <xsl:choose>
                <xsl:when test="$reverseCategories = 'true' ">
                  <xsl:for-each select="key('rows','')">
                    <xsl:call-template name="InsertPointsReverse">
                      <xsl:with-param name="series" select="$thisSeries"/>
                      <xsl:with-param name="numCategories" select="$numPoints"/>
                    </xsl:call-template>
                  </xsl:for-each>
                </xsl:when>

                <xsl:otherwise>
                  <xsl:for-each select="key('rows','')">
                    <xsl:call-template name="InsertPoints">
                      <xsl:with-param name="series" select="$thisSeries"/>
                    </xsl:call-template>
                  </xsl:for-each>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:for-each>
          </c:numCache>
        </c:numRef>
      </c:val>

      <!-- smooth line -->
      <xsl:for-each select="ancestor::chart:chart">
        <xsl:if test="@chart:class = 'chart:line' ">
          <xsl:for-each
            select="key('style',chart:plot-area/@chart:style-name)/style:chart-properties">
            <xsl:if test="@chart:interpolation != 'none' ">
              <c:smooth val="1"/>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
      </xsl:for-each>
    </c:ser>
  </xsl:template>

  <xsl:template name="InsertPoints">
    <!-- @Description: Outputs series data points -->
    <!-- @Context: table:table-rows -->

    <xsl:param name="series"/>

    <xsl:for-each select="table:table-row">
      <xsl:if test="table:table-cell[$series + 2]/text:p != '1.#NAN' ">
        <c:pt idx="{position() - 1}">
          <c:v>
            <!-- $ series + 2 because position starts with 1 and we skip first cell -->
            <xsl:value-of select="table:table-cell[$series + 2]/text:p"/>
          </c:v>
        </c:pt>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="InsertPointsReverse">
    <!-- @Description: Outputs series data points -->
    <!-- @Context: table:table-rows -->

    <xsl:param name="series"/>
    <xsl:param name="numCategories"/>
    <xsl:param name="count" select="0"/>

    <xsl:choose>
      <xsl:when test="$count &lt; $numCategories">
        <xsl:for-each select="table:table-row[$numCategories - $count]">
          <xsl:if test="table:table-cell[$series + 2]/text:p != '1.#NAN' ">
            <c:pt idx="{$count}">
              <c:v>
                <!-- $ series + 2 because position starts with 1 and we skip first cell -->
                <xsl:value-of select="table:table-cell[$series + 2]/text:p"/>
              </c:v>
            </c:pt>
          </xsl:if>
        </xsl:for-each>

        <xsl:call-template name="InsertPointsReverse">
          <xsl:with-param name="series" select="$series"/>
          <xsl:with-param name="numCategories" select="$numCategories"/>
          <xsl:with-param name="count" select="$count + 1"/>
        </xsl:call-template>
      </xsl:when>
    </xsl:choose>

  </xsl:template>

  <xsl:template name="InsertCategories">
    <!-- @Description: Outputs categories names-->
    <!-- @Context: table:table-rows -->

    <xsl:param name="numCategories"/>

    <!-- categories names -->
    <xsl:for-each select="table:table-row">
      <c:pt idx="{position() - 1}">
        <c:v>
          <xsl:value-of select="table:table-cell[1]/text:p"/>
        </c:v>
      </c:pt>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="InsertCategoriesReverse">
    <!-- @Description: Outputs categories names-->
    <!-- @Context: table:table-rows -->

    <xsl:param name="numCategories"/>
    <xsl:param name="count" select="0"/>

    <!-- categories names -->
    <xsl:choose>
      <xsl:when test="$count &lt; $numCategories">
        <xsl:for-each select="table:table-row[$numCategories - $count]">
          <c:pt idx="{$count}">
            <c:v>
              <xsl:value-of select="table:table-cell[1]/text:p"/>
            </c:v>
          </c:pt>
        </xsl:for-each>

        <xsl:call-template name="InsertCategoriesReverse">
          <xsl:with-param name="numCategories" select="$numCategories"/>
          <xsl:with-param name="count" select="$count + 1"/>
        </xsl:call-template>
      </xsl:when>
    </xsl:choose>

  </xsl:template>

  <xsl:template name="SetDataGroupingAtribute">
    <!-- @Description: Sets data grouping type -->
    <!-- @Context: chart:chart -->

    <!-- choose data grouping -->
    <xsl:for-each select="key('style',chart:plot-area/@chart:style-name)/style:chart-properties">
      <xsl:choose>
        <xsl:when test="@chart:stacked = 'true' ">
          <xsl:attribute name="val">
            <xsl:text>stacked</xsl:text>
          </xsl:attribute>
        </xsl:when>

        <xsl:when test="@chart:percentage = 'true' ">
          <xsl:attribute name="val">
            <xsl:text>percentStacked</xsl:text>
          </xsl:attribute>
        </xsl:when>

        <xsl:when test="@chart:deep = 'true' ">
          <xsl:attribute name="val">
            <xsl:text>standard</xsl:text>
          </xsl:attribute>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="InsertSpPr">
    <xsl:param name="chartDirectory"/>

    <xsl:for-each select="key('style', @chart:style-name)/style:graphic-properties">
      <c:spPr>
        <xsl:call-template name="InsertDrawingFill">
          <xsl:with-param name="chartDirectory" select="$chartDirectory"/>
        </xsl:call-template>
        <xsl:call-template name="InsertDrawingBorder"/>
      </c:spPr>
    </xsl:for-each>

  </xsl:template>

  <xsl:template name="InsertShapeProperties">
    <xsl:param name="styleName"/>
    <xsl:param name="parentStyleName"/>

    <!-- series shape property -->
    <c:spPr>

      <!-- fill color -->
      <xsl:if
        test="key('style',$styleName)/style:graphic-properties/@draw:fill-color or key('style',$parentStyleName)/style:graphic-properties/@draw:fill-color">
        <a:solidFill>
          <a:srgbClr val="9999FF">
            <xsl:attribute name="val">
              <xsl:choose>
                <xsl:when test="key('style',$styleName)/style:graphic-properties/@draw:fill-color">
                  <xsl:value-of
                    select="substring(key('style',$styleName)/style:graphic-properties/@draw:fill-color,2)"
                  />
                </xsl:when>
                <xsl:when
                  test="key('style',$parentStyleName)/style:graphic-properties/@draw:fill-color">
                  <xsl:value-of
                    select="substring(key('style',$parentStyleName)/style:graphic-properties/@draw:fill-color,2)"
                  />
                </xsl:when>
              </xsl:choose>
            </xsl:attribute>
          </a:srgbClr>
        </a:solidFill>
      </xsl:if>

      <!-- line color -->
      <xsl:if test="not(key('style',$styleName)/style:graphic-properties/@draw:stroke = 'none')">
        <a:ln w="3175">
          <a:solidFill>
            <a:srgbClr val="000000">
              <xsl:if
                test="(key('style',$styleName)/style:graphic-properties/@svg:stroke-color or key('style',$parentStyleName)/style:graphic-properties/@svg:stroke-color)">
                <xsl:attribute name="val">
                  <xsl:choose>
                    <xsl:when
                      test="key('style',$styleName)/style:graphic-properties/@svg:stroke-color">
                      <xsl:value-of
                        select="substring(key('style',$styleName)/style:graphic-properties/@svg:stroke-color,2)"
                      />
                    </xsl:when>
                    <xsl:when
                      test="key('style',$parentStyleName)/style:graphic-properties/@svg:stroke-color">
                      <xsl:value-of
                        select="substring(key('style',$parentStyleName)/style:graphic-properties/@svg:stroke-color,2)"
                      />
                    </xsl:when>
                  </xsl:choose>
                </xsl:attribute>
              </xsl:if>
            </a:srgbClr>
          </a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
      </xsl:if>
    </c:spPr>

  </xsl:template>

  <xsl:template name="InsertDataPointsShapeProperties">
    <!-- @Description: Sets data grouping type -->
    <!-- @Context: chart:data-point -->

    <xsl:param name="parentStyleName"/>
    <xsl:param name="reverse"/>
    <xsl:param name="count" select="0"/>

    <xsl:variable name="points">
      <xsl:choose>
        <xsl:when test="@chart:repeated">
          <xsl:value-of select="@chart:repeated"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:text>1</xsl:text>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <!-- only fill and stroke color for now -->
    <xsl:if test="@chart:style-name">
      <c:dPt>
        <c:idx val="{$count}"/>
        <xsl:call-template name="InsertShapeProperties">
          <xsl:with-param name="parentStyleName" select="$parentStyleName"/>
          <xsl:with-param name="styleName" select="@chart:style-name"/>
        </xsl:call-template>
      </c:dPt>
    </xsl:if>

    <!-- get next data point -->
    <xsl:choose>
      <!-- previous if categories are aligned in reverse order -->
      <xsl:when test="$reverse = 'true' ">
        <xsl:if test="preceding-sibling::node()[1]">
          <xsl:for-each select="preceding-sibling::node()[1]">
            <xsl:call-template name="InsertDataPointsShapeProperties">
              <xsl:with-param name="parentStyleName" select="$parentStyleName"/>
              <xsl:with-param name="reverse" select="$reverse"/>
              <xsl:with-param name="count" select="$count + $points"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:if>
      </xsl:when>
      <!-- folowing data point -->
      <xsl:otherwise>
        <xsl:if test="following-sibling::node()[1]">
          <xsl:for-each select="following-sibling::node()[1]">
            <xsl:call-template name="InsertDataPointsShapeProperties">
              <xsl:with-param name="parentStyleName" select="$parentStyleName"/>
              <xsl:with-param name="count" select="$count + $points"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <xsl:template name="InsertSecondaryLineChart">
    <c:lineChart>
      <c:grouping val="standard"/>

      <xsl:call-template name="InsertSecondaryChartContent"/>

    </c:lineChart>
  </xsl:template>

  <xsl:template name="InsertSecondaryChartContent">
    <!-- @Description: Inserts chart content -->
    <!-- @Context: chart:chart -->

    <xsl:variable name="numSeries">
      <!-- (number) number of series inside chart -->
      <xsl:value-of select="count(key('series','')[@chart:class='chart:line'])"/>
    </xsl:variable>

    <xsl:variable name="numPoints">
      <!-- (number) maximum number of data point -->
      <xsl:value-of select="count(key('rows','')/table:table-row)"/>
    </xsl:variable>

    <xsl:variable name="primarySeries">
      <xsl:value-of select="count(key('series','')[not(@chart:class='chart:line')])"/>
    </xsl:variable>

    <xsl:for-each select="key('rows','')">
      <xsl:call-template name="InsertSecondaryChartSeries">
        <xsl:with-param name="numSeries" select="$numSeries"/>
        <xsl:with-param name="numPoints" select="$numPoints"/>
        <xsl:with-param name="primarySeries" select="$primarySeries"/>
      </xsl:call-template>
    </xsl:for-each>

    <c:axId val="110226048"/>
    <c:axId val="110498176"/>

  </xsl:template>

  <xsl:template name="InsertSecondaryChartSeries">
    <!-- @Description: Outputs chart series and their values -->
    <!-- @Context: table:table-rows -->

    <xsl:param name="numSeries"/>
    <!-- (number) number of series inside chart -->
    <xsl:param name="numPoints"/>
    <!-- (number) maximum number of data point -->
    <xsl:param name="primarySeries"/>
    <xsl:param name="count" select="0"/>
    <!-- (number) loop counter -->

    <!-- count series from backwards if reverseSeries is = "true" -->
    <xsl:variable name="styleName">
      <!-- (string) series style name -->
      <xsl:value-of
        select="key('series','')[position() = $primarySeries + 1 + $count]/@chart:style-name"/>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="$count &lt; $numSeries">
        <c:ser>
          <c:idx val="{$primarySeries + $count}"/>
          <c:order val="{$primarySeries + $count}"/>

          <!-- series name -->
          <c:tx>
            <c:v>
              <xsl:value-of
                select="key('header','')/table:table-row/table:table-cell[$primarySeries + $count + 2]"
              />
            </c:v>
          </c:tx>

          <xsl:call-template name="InsertShapeProperties">
            <xsl:with-param name="styleName" select="$styleName"/>
          </xsl:call-template>

          <!-- marker type -->
          <xsl:if
            test="key('chart','' )/@chart:class = 'chart:line' or ancestor::chart:chart/chart:plot-area/chart:series[position() = $primarySeries + 1 + $count]/@chart:class = 'chart:line'">
            <xsl:for-each
              select="ancestor::chart:chart/chart:plot-area/chart:series[position() = $primarySeries + 1 + $count]">
              <xsl:choose>
                <xsl:when
                  test="key('style',@chart:style-name )/style:chart-properties/@chart:symbol-type = 'none' ">
                  <c:marker>
                    <c:symbol val="none"/>
                  </c:marker>
                </xsl:when>
              </xsl:choose>
            </xsl:for-each>
          </xsl:if>

          <!-- insert data categories -->
          <c:cat>
            <c:strLit>
              <c:ptCount val="{$numPoints}"/>
              <xsl:for-each select="key('rows','')">
                <xsl:call-template name="InsertCategories">
                  <xsl:with-param name="numCategories" select="$numPoints"/>
                </xsl:call-template>
              </xsl:for-each>
            </c:strLit>
          </c:cat>

          <!-- series values -->
          <c:val>
            <c:numRef>
              <!-- TO DO: reference to sheet cell -->
              <!-- i.e. <c:f>Sheet1!$D$3:$D$4</c:f> -->
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val="{$numPoints}"/>

                <!-- number of this series -->
                <xsl:variable name="thisSeries">
                  <xsl:value-of select="$primarySeries + $count"/>
                </xsl:variable>

                <xsl:for-each select="ancestor::chart:chart">
                  <xsl:for-each select="key('rows','')">
                    <xsl:call-template name="InsertPoints">
                      <xsl:with-param name="series" select="$thisSeries"/>
                    </xsl:call-template>
                  </xsl:for-each>
                </xsl:for-each>
              </c:numCache>
            </c:numRef>
          </c:val>

          <!-- smooth line -->
          <xsl:for-each select="ancestor::chart:chart">
            <xsl:if test="@chart:class = 'chart:line' ">
              <xsl:for-each
                select="key('style',chart:plot-area/@chart:style-name)/style:chart-properties">
                <xsl:if test="@chart:interpolation != 'none' ">
                  <c:smooth val="1"/>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:for-each>
        </c:ser>

        <xsl:call-template name="InsertSecondaryChartSeries">
          <xsl:with-param name="numSeries" select="$numSeries"/>
          <xsl:with-param name="numPoints" select="$numPoints"/>
          <xsl:with-param name="primarySeries" select="$primarySeries"/>
          <xsl:with-param name="count" select="$count + 1"/>
        </xsl:call-template>
      </xsl:when>
    </xsl:choose>

  </xsl:template>

  <xsl:template name="InsertView3D">
    <!-- @Description: Sets 3D view definition -->
    <!-- @Context: chart:chart -->

    <xsl:if
      test="@chart:class = 'chart:circle' and key('style',chart:plot-area/@chart:style-name)/style:chart-properties/@chart:three-dimensional = 'true' ">
      <c:view3D>
        <c:rotY val="90"/>
      </c:view3D>
    </xsl:if>

  </xsl:template>

  <xsl:template name="InsertRingPointsShapeProperties">
    <!-- @Description: Inserts ring chart data point shape property-->
    <!-- @Context: chart:series[last()] -->
    <xsl:param name="totalPoints"/>
    <xsl:param name="series"/>
    <xsl:param name="count" select="0"/>

    <xsl:variable name="dataPointStyle">
      <xsl:for-each
        select="parent::node()/chart:series[position() = $totalPoints - $count]/chart:data-point[position() = 1]">
        <xsl:call-template name="GetDataPointStyleName">
          <xsl:with-param name="point" select="$series"/>
        </xsl:call-template>
      </xsl:for-each>
    </xsl:variable>

    <xsl:variable name="styleName">
      <xsl:choose>
        <xsl:when test="$dataPointStyle != '' ">
          <xsl:value-of select="$dataPointStyle"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="@chart:style-name"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <c:dPt>
      <c:idx val="{$count}"/>
      <xsl:call-template name="InsertShapeProperties">
        <xsl:with-param name="styleName" select="$styleName"/>
      </xsl:call-template>
    </c:dPt>

    <xsl:for-each select="preceding-sibling::node()[name() = 'chart:series' ][1]">
      <xsl:call-template name="InsertRingPointsShapeProperties">
        <xsl:with-param name="totalPoints" select="$totalPoints"/>
        <xsl:with-param name="series" select="$series"/>
        <xsl:with-param name="count" select="$count + 1"/>
      </xsl:call-template>
    </xsl:for-each>

  </xsl:template>

  <xsl:template name="GetDataPointStyleName">
    <!-- @Description: Gets data point style name -->
    <!-- @Context: chart:data-point[1] -->

    <xsl:param name="point"/>
    <xsl:param name="count" select="0"/>

    <xsl:variable name="points">
      <xsl:choose>
        <xsl:when test="@chart:repeated">
          <xsl:value-of select="@chart:repeated"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:text>1</xsl:text>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="$count + $points &gt;= $point ">
        <xsl:value-of select="@chart:style-name"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="following-sibling::chart:data-point">
          <xsl:for-each select="following-sibling::chart:data-point[1]">
            <xsl:call-template name="GetDataPointStyleName">
              <xsl:with-param name="point" select="$point"/>
              <xsl:with-param name="count" select="$count + $points"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsideChartPosition">
    <xsl:param name="chartWidth"/>
    <xsl:param name="chartHeight"/>

    <xsl:variable name="xOffset">
      <xsl:value-of select="substring-before(@svg:x,'cm' )"/>
    </xsl:variable>

    <xsl:variable name="yOffset">
      <xsl:value-of select="substring-before(@svg:y,'cm' )"/>
    </xsl:variable>

    <xsl:variable name="width">
      <xsl:value-of select="substring-before(@svg:width,'cm' )"/>
    </xsl:variable>

    <xsl:variable name="height">
      <xsl:value-of select="substring-before(@svg:height,'cm' )"/>
    </xsl:variable>

    <!-- ugly hack for plot area height increase when width is too wide -->
    <xsl:variable name="plotWidthCorrection">
      <xsl:text>0.025</xsl:text>
    </xsl:variable>

    <xsl:variable name="plotHeightCorrection">
      <xsl:text>0.05</xsl:text>
    </xsl:variable>

    <!-- ugly hack for axis Y title influence on plot-area display -->
    <xsl:variable name="plotXCorrection">
      <xsl:text>0.1</xsl:text>
    </xsl:variable>

    <xsl:if test="$xOffset div $chartWidth">
      <c:x>

        <xsl:attribute name="val">
          <xsl:variable name="correction">
            <xsl:choose>
              <xsl:when
                test="chart:axis[@chart:dimension = 'y' ]/chart:title and name() = 'chart:plot-area' ">
                <xsl:value-of select="$plotXCorrection"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="0"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:value-of select="$xOffset div $chartWidth + $correction"/>
        </xsl:attribute>

      </c:x>
    </xsl:if>


    <xsl:choose>
      <!-- when axis Y title -->
      <xsl:when test="parent::node()/@chart:dimension = 'y' ">
        <xsl:if test="1 - $yOffset div $chartHeight">
          <c:y>
            <xsl:attribute name="val">
              <xsl:value-of select="1 - $yOffset div $chartHeight"/>
            </xsl:attribute>
          </c:y>
        </xsl:if>
      </xsl:when>
      <xsl:when test="//chart:chart[@chart:class='chart:radar']">
        <c:y>
          <xsl:attribute name="val">
            <xsl:value-of select="20"/>
          </xsl:attribute>
        </c:y>
      </xsl:when>
      <xsl:when test="$yOffset div $chartHeight">
        <c:y>
          <xsl:attribute name="val">
            <xsl:value-of select="$yOffset div $chartHeight"/>
          </xsl:attribute>
        </c:y>
      </xsl:when>
      <xsl:otherwise/>
    </xsl:choose>

    <xsl:if test="@svg:width">
      <c:w>
        <xsl:attribute name="val">

          <xsl:variable name="correction">
            <xsl:choose>
              <!-- ugly hack for axis title influence on plot-area display -->
              <!--xsl:when
                test="chart:axis[@chart:dimension = 'y' ]/chart:title and name() = 'chart:plot-area' ">
                <xsl:value-of select="$plotWidthCorrection + $plotXCorrection"/>
              </xsl:when-->
              <xsl:when test="name() = 'chart:plot-area' ">
                <xsl:value-of select="$plotWidthCorrection"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="0"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>

          <xsl:value-of select="$width div $chartWidth - $correction"/>
        </xsl:attribute>

      </c:w>
    </xsl:if>

    <xsl:if test="@svg:height">
      <c:h>
        <xsl:attribute name="val">
          <xsl:variable name="correction">
            <xsl:choose>
              <!-- ugly hack for axis title influence on plot-area display -->
              <xsl:when
                test="chart:axis[@chart:dimension = 'x' ]/chart:title and name() = 'chart:plot-area' ">
                <xsl:value-of select="$plotHeightCorrection + $plotXCorrection"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="0"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>

          <!--xsl:value-of select="$height div $chartHeight - $correction"/-->
          <xsl:value-of select="$height div $chartHeight"/>
        </xsl:attribute>

      </c:h>
    </xsl:if>

  </xsl:template>

  <xsl:template name="InsertCatAx">
    <!-- @Description: Inserts Category Axis -->
    <!-- @Context: chart:chart -->
    <xsl:param name="chartWidth"/>
    <xsl:param name="chartHeight"/>

    <c:catAx>
      <c:axId val="110226048"/>
      <c:scaling>
        <c:orientation val="minMax"/>
      </c:scaling>

      <c:axPos val="b"/>
      <xsl:call-template name="InsertTitle">
        <xsl:with-param name="chartWidth" select="$chartWidth"/>
        <xsl:with-param name="chartHeight" select="$chartHeight"/>
      </xsl:call-template>
      <c:majorGridlines/>

      <xsl:call-template name="InsertTitle">
        <xsl:with-param name="chartWidth" select="$chartWidth"/>
        <xsl:with-param name="chartHeight" select="$chartHeight"/>
      </xsl:call-template>

      <c:tickLblPos val="low">
        <xsl:if
          test="key('style',@chart:style-name)/style:chart-properties/@chart:display-label = 'false' ">
          <xsl:attribute name="val">
            <xsl:text>none</xsl:text>
          </xsl:attribute>
        </xsl:if>
      </c:tickLblPos>

      <c:spPr>
        <a:ln w="3175">
          <a:solidFill>
            <a:srgbClr val="000000"/>
          </a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
      </c:spPr>
      <c:txPr>
        <a:bodyPr rot="0" vert="horz"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr>
            <a:defRPr sz="700" b="0" i="0" u="none" strike="noStrike" baseline="0">

              <xsl:for-each select="key('style',@chart:style-name)/style:text-properties">
                <!-- template common with text-box-->
                <xsl:call-template name="InsertRunProperties"/>
              </xsl:for-each>

            </a:defRPr>
          </a:pPr>
          <a:endParaRPr lang="pl-PL"/>
        </a:p>
      </c:txPr>
      <c:crossAx val="110498176"/>
      <c:crosses val="autoZero"/>

    </c:catAx>
  </xsl:template>

</xsl:stylesheet>

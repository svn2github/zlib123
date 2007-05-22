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
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main">

  <xsl:import href="relationships.xsl"/>

  <xsl:template name="CreateCharts">

    <!-- get all sheet Id's -->
    <xsl:for-each select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet">

      <xsl:variable name="sheet">
        <xsl:call-template name="GetTarget">
          <xsl:with-param name="id">
            <xsl:value-of select="@r:id"/>
          </xsl:with-param>
          <xsl:with-param name="document">xl/workbook.xml</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <!--i.e. $sheet = worksheets/sheet1.xml -->

      <xsl:variable name="sheetNum">
        <xsl:value-of select="position()"/>
      </xsl:variable>

      <!-- go to sheet file and search for drawing -->
      <xsl:for-each select="document(concat('xl/',$sheet))/e:worksheet/e:drawing">

        <xsl:variable name="drawing">
          <xsl:call-template name="GetTarget">
            <xsl:with-param name="id">
              <xsl:value-of select="@r:id"/>
            </xsl:with-param>
            <xsl:with-param name="document">
              <xsl:value-of select="concat('xl/',$sheet)"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <!-- i.e. $drawing = ../drawings/drawing2.xml -->

        <!-- go to sheet drawing file and search for charts -->
        <xsl:for-each
          select="document(concat('xl/',substring-after($drawing,'/')))/xdr:wsDr/xdr:twoCellAnchor/xdr:graphicFrame/a:graphic/a:graphicData/c:chart">

          <xsl:variable name="chart">
            <xsl:call-template name="GetTarget">
              <xsl:with-param name="id">
                <xsl:value-of select="@r:id"/>
              </xsl:with-param>
              <xsl:with-param name="document">
                <xsl:value-of select="concat('xl/',substring-after($drawing,'/'))"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <!-- i.e. $chart = ../charts/chart1.xml -->

          <xsl:variable name="chartId">
            <xsl:value-of select="generate-id(.)"/>
          </xsl:variable>

          <!-- finally go to a chart file -->
          <xsl:for-each select="document(concat('xl/',substring-after($chart,'/')))">
            
            <xsl:call-template name="InsertChart">
              <xsl:with-param name="chartId" select="$chartId"/>
            </xsl:call-template>
            
          </xsl:for-each>
        </xsl:for-each>
      </xsl:for-each>
    </xsl:for-each>

  </xsl:template>

  <xsl:template name="InsertChart">
    <xsl:param name="chartId"/>

    <xsl:call-template name="InsertChartContent">
      <xsl:with-param name="chartId" select="$chartId"/>
    </xsl:call-template>

    <xsl:call-template name="InsertChartStyles">
      <xsl:with-param name="chartId" select="$chartId"/>
    </xsl:call-template>
  </xsl:template>

  <xsl:template name="InsertChartContent">
    <xsl:param name="chartId"/>

    <pzip:entry pzip:target="{concat('Object ',$chartId,'/content.xml')}">
      <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
        xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
        xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
        xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
        xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
        xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/"
        xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
        xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
        xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
        xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
        xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"
        xmlns:math="http://www.w3.org/1998/Math/MathML"
        xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"
        xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
        xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer"
        xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events"
        xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema"
        xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" office:version="1.0">
        <office:automatic-styles>
          <number:number-style style:name="N0">
            <number:number number:min-integer-digits="1"/>
          </number:number-style>
          <style:style style:name="ch1" style:family="chart">
            <style:graphic-properties draw:stroke="none" draw:fill-color="#ffffff"/>
          </style:style>
          <style:style style:name="ch2" style:family="chart">
            <style:chart-properties style:direction="ltr"/>
            <style:graphic-properties draw:stroke="none" draw:fill="none"/>
            <style:text-properties fo:font-family="Arial" style:font-family-generic="swiss"
              style:font-pitch="variable" fo:font-size="13pt"
              style:font-family-asian="&apos;MS Gothic&apos;"
              style:font-family-generic-asian="system" style:font-pitch-asian="variable"
              style:font-size-asian="13pt" style:font-family-complex="Tahoma"
              style:font-family-generic-complex="system" style:font-pitch-complex="variable"
              style:font-size-complex="13pt"/>
          </style:style>
          <style:style style:name="ch3" style:family="chart">
            <style:graphic-properties draw:fill="none"/>
            <style:text-properties fo:font-family="Arial" style:font-family-generic="swiss"
              style:font-pitch="variable" fo:font-size="6pt"
              style:font-family-asian="&apos;MS Gothic&apos;"
              style:font-family-generic-asian="system" style:font-pitch-asian="variable"
              style:font-size-asian="6pt" style:font-family-complex="Tahoma"
              style:font-family-generic-complex="system" style:font-pitch-complex="variable"
              style:font-size-complex="6pt"/>
          </style:style>
          <style:style style:name="ch4" style:family="chart">
            <style:chart-properties chart:japanese-candle-stick="false" chart:stock-with-volume="false"
              chart:three-dimensional="false" chart:deep="false" chart:lines="false"
              chart:interpolation="none" chart:symbol-type="none" chart:vertical="false"
              chart:lines-used="0" chart:connect-bars="false" chart:series-source="columns"
              chart:mean-value="false" chart:error-margin="0" chart:error-lower-limit="0"
              chart:error-upper-limit="0" chart:error-category="none" chart:error-percentage="0"
              chart:regression-type="none" chart:data-label-number="none" chart:data-label-text="false"
              chart:data-label-symbol="false"/>
          </style:style>
          <style:style style:name="ch5" style:family="chart" style:data-style-name="N0">
            <style:chart-properties chart:display-label="true" chart:tick-marks-major-inner="false"
              chart:tick-marks-major-outer="true" chart:logarithmic="false" chart:text-overlap="false"
              text:line-break="true" chart:label-arrangement="side-by-side" chart:visible="true"
              style:direction="ltr"/>
            <style:graphic-properties draw:stroke="solid" svg:stroke-width="0cm"
              svg:stroke-color="#000000"/>
            <style:text-properties fo:font-family="Arial" style:font-family-generic="swiss"
              style:font-pitch="variable" fo:font-size="7pt"
              style:font-family-asian="&apos;MS Gothic&apos;"
              style:font-family-generic-asian="system" style:font-pitch-asian="variable"
              style:font-size-asian="7pt" style:font-family-complex="Tahoma"
              style:font-family-generic-complex="system" style:font-pitch-complex="variable"
              style:font-size-complex="7pt"/>
          </style:style>
          <style:style style:name="ch6" style:family="chart" style:data-style-name="N0">
            <style:chart-properties chart:display-label="true" chart:tick-marks-major-inner="false"
              chart:tick-marks-major-outer="true" chart:logarithmic="false" chart:origin="0"
              chart:gap-width="100" chart:overlap="0" chart:text-overlap="false" text:line-break="false"
              chart:label-arrangement="side-by-side" chart:visible="true" style:direction="ltr"/>
            <style:graphic-properties draw:stroke="solid" svg:stroke-width="0cm"
              svg:stroke-color="#000000"/>
            <style:text-properties fo:font-family="Arial" style:font-family-generic="swiss"
              style:font-pitch="variable" fo:font-size="7pt"
              style:font-family-asian="&apos;MS Gothic&apos;"
              style:font-family-generic-asian="system" style:font-pitch-asian="variable"
              style:font-size-asian="7pt" style:font-family-complex="Tahoma"
              style:font-family-generic-complex="system" style:font-pitch-complex="variable"
              style:font-size-complex="7pt"/>
          </style:style>
          <style:style style:name="ch7" style:family="chart">
            <style:graphic-properties draw:fill-color="#9999ff"/>
            <style:text-properties fo:font-family="Arial" style:font-family-generic="swiss"
              style:font-pitch="variable" fo:font-size="6pt"
              style:font-family-asian="&apos;MS Gothic&apos;"
              style:font-family-generic-asian="system" style:font-pitch-asian="variable"
              style:font-size-asian="6pt" style:font-family-complex="Tahoma"
              style:font-family-generic-complex="system" style:font-pitch-complex="variable"
              style:font-size-complex="6pt"/>
          </style:style>
          <style:style style:name="ch8" style:family="chart">
            <style:graphic-properties draw:fill-color="#993366"/>
            <style:text-properties fo:font-family="Arial" style:font-family-generic="swiss"
              style:font-pitch="variable" fo:font-size="6pt"
              style:font-family-asian="&apos;MS Gothic&apos;"
              style:font-family-generic-asian="system" style:font-pitch-asian="variable"
              style:font-size-asian="6pt" style:font-family-complex="Tahoma"
              style:font-family-generic-complex="system" style:font-pitch-complex="variable"
              style:font-size-complex="6pt"/>
          </style:style>
          <style:style style:name="ch9" style:family="chart">
            <style:graphic-properties draw:fill-color="#ffffcc"/>
            <style:text-properties fo:font-family="Arial" style:font-family-generic="swiss"
              style:font-pitch="variable" fo:font-size="6pt"
              style:font-family-asian="&apos;MS Gothic&apos;"
              style:font-family-generic-asian="system" style:font-pitch-asian="variable"
              style:font-size-asian="6pt" style:font-family-complex="Tahoma"
              style:font-family-generic-complex="system" style:font-pitch-complex="variable"
              style:font-size-complex="6pt"/>
          </style:style>
          <style:style style:name="ch10" style:family="chart">
            <style:graphic-properties draw:stroke="none" draw:fill="none"/>
          </style:style>
          <style:style style:name="ch11" style:family="chart">
            <style:graphic-properties draw:stroke="none" draw:fill-color="#999999"/>
          </style:style>
        </office:automatic-styles>
        <office:body>
          <office:chart>
            <chart:chart svg:width="8cm" svg:height="7cm" chart:class="chart:bar" chart:style-name="ch1">
              <chart:title svg:x="3cm" svg:y="0.14cm" chart:style-name="ch2">
                <text:p>Main Title</text:p>
              </chart:title>
              <chart:legend chart:legend-position="end" svg:x="6.459cm" svg:y="3.005cm"
                chart:style-name="ch3"/>
              <chart:plot-area chart:style-name="ch4" table:cell-range-address="Sheet1.$D$3:.$F$3"
                chart:table-number-list="0" svg:x="0.16cm" svg:y="0.937cm" svg:width="5.98cm"
                svg:height="5.923cm">
                <chart:axis chart:dimension="x" chart:name="primary-x" chart:style-name="ch5">
                  <chart:categories table:cell-range-address="local-table.A2:.A2"/>
                </chart:axis>
                <chart:axis chart:dimension="y" chart:name="primary-y" chart:style-name="ch6">
                  <chart:grid chart:class="major"/>
                </chart:axis>
                <chart:series chart:style-name="ch7">
                  <chart:data-point/>
                </chart:series>
                <chart:series chart:style-name="ch8">
                  <chart:data-point/>
                </chart:series>
                <chart:series chart:style-name="ch9">
                  <chart:data-point/>
                </chart:series>
                <chart:wall chart:style-name="ch10"/>
                <chart:floor chart:style-name="ch11"/>
              </chart:plot-area>
              <table:table table:name="local-table">
                <table:table-header-columns>
                  <table:table-column/>
                </table:table-header-columns>
                <table:table-columns>
                  <table:table-column table:number-columns-repeated="3"/>
                </table:table-columns>
                <table:table-header-rows>
                  <table:table-row>
                    <table:table-cell>
                      <text:p/>
                    </table:table-cell>
                    <table:table-cell office:value-type="string">
                      <text:p>Column D</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="string">
                      <text:p>Column E</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="string">
                      <text:p>Column F</text:p>
                    </table:table-cell>
                  </table:table-row>
                </table:table-header-rows>
                <table:table-rows>
                  <table:table-row>
                    <table:table-cell office:value-type="string">
                      <text:p>Row 3</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="float" office:value="3">
                      <text:p>3</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="float" office:value="6">
                      <text:p>6</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="float" office:value="4">
                      <text:p>4</text:p>
                    </table:table-cell>
                  </table:table-row>
                </table:table-rows>
              </table:table>
            </chart:chart>
          </office:chart>
        </office:body>
      </office:document-content>
      
    </pzip:entry>
  </xsl:template>

  <xsl:template name="InsertChartStyles">
    <xsl:param name="chartId"/>

    <pzip:entry pzip:target="{concat('Object ',$chartId,'/styles.xml')}">
      <office:document-styles>
        <xsl:attribute name="zzzzz">
          <xsl:value-of select="$chartId"/>
        </xsl:attribute>
        <office:styles/>
      </office:document-styles>
    </pzip:entry>
  </xsl:template>
</xsl:stylesheet>

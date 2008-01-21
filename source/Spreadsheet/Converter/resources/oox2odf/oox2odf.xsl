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
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:oox="urn:oox"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
  exclude-result-prefixes="oox e r xdr c">

  <xsl:import href="content.xsl"/>
  <xsl:import href="common-meta.xsl"/>
  <xsl:import href="relationships.xsl"/>
  <xsl:import href="settings.xsl"/>
  <xsl:import href="common.xsl"/>
  <xsl:import href="styles.xsl"/>
  <xsl:import href="merge_cell.xsl"/>
  <xsl:import href="chart.xsl"/>

  <xsl:param name="outputFile"/>
  <xsl:output method="xml" encoding="UTF-8"/>

  <xsl:key name="Part" match="/oox:package/oox:part" use="@oox:name"/>

  <!-- App version number -->
  <xsl:variable name="app-version">1.0.0</xsl:variable>

  <xsl:template match="/oox:package">
    <pzip:archive pzip:target="{$outputFile}">

      <!-- Manifest -->
      <pzip:entry pzip:target="META-INF/manifest.xml">
        <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
          <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"
            manifest:full-path="/"/>
          <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
          <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml"/>
          <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="meta.xml"/>
          <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="settings.xml"/>
          <xsl:call-template name="InsertChartEntries"/>
        </manifest:manifest>
      </pzip:entry>

      <!-- main content -->
      <pzip:entry pzip:target="content.xml">
        <xsl:call-template name="content"/>
      </pzip:entry>

      <!-- styles  -->
      <pzip:entry pzip:target="styles.xml">
        <xsl:call-template name="styles"/>
      </pzip:entry>

      <!-- meta -->
      <pzip:entry pzip:target="meta.xml">
        <xsl:call-template name="meta"/>
      </pzip:entry>

      <!-- settings  -->
      <pzip:entry pzip:target="settings.xml">
        <xsl:call-template name="InsertSettings"/>
      </pzip:entry>

      <!-- charts created as new object -->
      <xsl:call-template name="CreateObjectCharts"/>
      <!-- charts created as new sheet -->
      <xsl:call-template name="CreateSheetCharts"/>

    </pzip:archive>
  </xsl:template>

  <xsl:template name="InsertChartEntries">
    <!-- get all sheet Id's -->
    <xsl:for-each select="key('Part', 'xl/workbook.xml')/e:workbook/e:sheets/e:sheet">

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

      <!-- go to worksheet file and search for drawing -->
      <xsl:for-each select="key('Part', concat('xl/',$sheet))/e:worksheet/e:drawing">

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

        <!-- finally insert entry for each chart -->
        <xsl:for-each
          select="key('Part', concat('xl/',substring-after($drawing,'/')))/xdr:wsDr/xdr:twoCellAnchor/xdr:graphicFrame/a:graphic/a:graphicData/c:chart">
          <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.chart"
            manifest:full-path="{concat('Object ',generate-id(.),'/')}"/>
        </xsl:for-each>
      </xsl:for-each>

      <!-- go to chartsheet file and search for drawing -->
      <xsl:for-each select="key('Part', concat('xl/',$sheet))/e:chartsheet/e:drawing">

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

        <!-- finally insert entry for each chart -->
        <xsl:for-each
          select="key('Part', concat('xl/',substring-after($drawing,'/')))/xdr:wsDr/xdr:absoluteAnchor/xdr:graphicFrame/a:graphic/a:graphicData/c:chart">
          <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.chart"
            manifest:full-path="{concat('Object ',generate-id(.),'/')}"/>
        </xsl:for-each>
      </xsl:for-each>

    </xsl:for-each>
  </xsl:template>
</xsl:stylesheet>

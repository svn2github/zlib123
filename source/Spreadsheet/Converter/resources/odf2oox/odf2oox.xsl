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
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:odf="urn:odf"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns:math="http://www.w3.org/1998/Math/MathML"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:x="urn:schemas-microsoft-com:office:excel" exclude-result-prefixes="odf style text number">

  <xsl:import href="workbook.xsl"/>
  <xsl:import href="sharedStrings.xsl"/>
  <xsl:import href="odf2oox-compute-size.xsl"/>
  <xsl:import href="contentTypes.xsl"/>
  <xsl:import href="package_relationships.xsl"/>
  <xsl:import href="common-meta.xsl"/>
  <xsl:import href="part_relationships.xsl"/>
  <xsl:import href="common.xsl"/>
  <xsl:import href="merge_cell.xsl"/>
  <xsl:import href="styles.xsl"/>
  <xsl:import href="number.xsl"/>
  <xsl:import href="comments.xsl"/>
  <xsl:import href="date.xsl"/>
  <xsl:import href="chart.xsl"/>
  <xsl:import href="drawing.xsl"/>
  <xsl:import href="connections.xsl"/>

  <xsl:strip-space elements="*"/>
  <xsl:preserve-space elements="text:p text:span number:text"/>

  <xsl:param name="outputFile"/>
  <xsl:output method="xml" encoding="UTF-8"/>

  <!-- App version number -->
  <!-- WARNING: it has to be of type xx.yy -->
  <!-- (otherwise Word cannot open the doc) -->
  <xsl:variable name="app-version">1.00</xsl:variable>

  <!-- existence of docProps/custom.xml file -->
  <xsl:variable name="docprops-custom-file">
    <xsl:choose>
      <xsl:when test="document('meta.xml')/office:document-meta/office:meta/meta:user-defined != ''">
        <xsl:text>1</xsl:text>      
      </xsl:when>
      <xsl:otherwise>
        <xsl:text>0</xsl:text>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:variable>
    

  <xsl:key name="chart" match="office:chart" use="''"/>

  <xsl:template match="/odf:source">
    <xsl:processing-instruction name="mso-application">progid="Word.Document"</xsl:processing-instruction>


    <pzip:archive pzip:target="{$outputFile}">

      <pzip:entry pzip:target="[Content_Types].xml">
        <xsl:call-template name="ContentTypes"/>
      </pzip:entry>
      <!-- CHANGE -->
      <!-- styles -->
      <pzip:entry pzip:target="xl/styles.xml">
        <xsl:call-template name="styles"/>
      </pzip:entry>
      <!-- input: styles.xml -->

      <!-- main content -->
      <pzip:entry pzip:target="xl/workbook.xml">
        <xsl:call-template name="InsertWorkbook"/>
      </pzip:entry>
      
      <!-- main content -->
      <xsl:if
        test="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table/table:table-row/table:table-cell/table:cell-range-source">
        <pzip:entry pzip:target="xl/connections.xml">
          <xsl:call-template name="InsertConnections"/>
        </pzip:entry>
      </xsl:if>

      <!-- shared strings -->
      <pzip:entry pzip:target="xl/sharedStrings.xml">
        <xsl:call-template name="InsertSharedStrings"/>
      </pzip:entry>
      <!-- input: content.xml -->
      <!-- output: xl/sharedStrings.xml -->


      <!-- insert sheets -->
      <xsl:call-template name="InsertSheets"/>
      <!-- input: content.xml -->
      <!-- output:  xl/worksheets/sheet_N_.xml -->
      <xsl:if test="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table/draw:frame/draw:object">
      <pzip:entry pzip:target="xl/media/image1.emf">
        <empty/>
      </pzip:entry>
      </xsl:if>

      <xsl:for-each
        select="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table">

        <xsl:variable name="comment">
          <xsl:choose>
            <xsl:when test="descendant::office:annotation">
              <xsl:text>true</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>false</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:variable name="picture">
          <xsl:choose>
            <xsl:when
              test="descendant::draw:frame/draw:image[not(starts-with(@xlink:href,'./ObjectReplacements')) and not(name(parent::node()/parent::node()) = 'draw:g' )]">
              <xsl:text>true</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>false</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:variable name="hyperlink">
          <xsl:choose>
            <xsl:when test="descendant::text:a[not(ancestor::draw:custom-shape) and not(ancestor::office:annotation)]">
              <xsl:text>true</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>false</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:variable name="textBox">
          <xsl:choose>
            <xsl:when test="descendant::draw:frame/draw:text-box">
              <xsl:text>true</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>false</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:variable name="chart">
          <xsl:for-each select="descendant::draw:frame/draw:object">
            <xsl:choose>
              <xsl:when test="not(document(concat(translate(@xlink:href,'./',''),'/settings.xml')))">
                <xsl:for-each
                  select="document(concat(translate(@xlink:href,'./',''),'/content.xml'))">
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

        <xsl:variable name="oleObject">
          <xsl:choose>
            <xsl:when test="descendant::draw:frame/draw:object">
              <xsl:text>true</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>false</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>


        <!-- insert comments -->
        <xsl:if test="$comment = 'true' ">
          <xsl:call-template name="InsertComments">
            <xsl:with-param name="sheetId" select="position()"/>
          </xsl:call-template>
        </xsl:if>

        <!-- create VmlDrawing.xml file for comment`s-->
        <xsl:if test="$comment = 'true'">
        <xsl:call-template name="CreateVmlDrawing"/>
        </xsl:if>

        <!-- create VmlDrawing.xml file -->
        <xsl:if test="$oleObject = 'true'">
            <xsl:call-template name="CreateVmlDrawing"/>
          <pzip:entry pzip:target="xl/media/image1.emf">
            <empty/>
          </pzip:entry>       
        </xsl:if>
     
        <xsl:if test="contains($chart,'true') or $picture='true' or $textBox = 'true' ">
          <xsl:call-template name="CreateDrawing"/>

          <xsl:if test="contains($chart,'true') or $picture='true' or $oleObject='true'">
            <xsl:call-template name="CreateDrawingRelationships"/>
          </xsl:if>

          <xsl:if test="contains($chart,'true')">
            <xsl:call-template name="CreateChartFile">
              <xsl:with-param name="sheetNum" select="position()"/>
            </xsl:call-template>
          </xsl:if>

        </xsl:if>

        <!-- insert relationships -->
        <xsl:call-template name="CreateSheetRelationships">
          <xsl:with-param name="sheetNum" select="position()"/>
          <xsl:with-param name="comment" select="$comment"/>
          <xsl:with-param name="chart" select="$chart"/>
          <xsl:with-param name="picture" select="$picture"/>
          <xsl:with-param name="hyperlink" select="$hyperlink"/>
          <xsl:with-param name="textBox" select="$textBox"/>
        </xsl:call-template>
      </xsl:for-each>

      <!-- insert pivotTables -->
      <!-- <xsl:call-template name="InsertPivotTables"/> -->
      <!-- input: content.xml -->
      <!-- output:  xl/pivotTable/pivotTable_N_.xml 
                                 xl/pivotTable/_rels/pivotTable_N_.xml.rels 
                                 xl/pivotCacheDefinition/pivotCacheDefinition_N_.xml
                                 xl/pivotCacheDefinition/pivotCacheRecords_N_.xml
                                 xl/pivotCacheDefinition/_rels/pivotCacheDefinition_N_.xml.rels                
            -->

      <!-- insert revisions -->
      <!-- <xsl:call-template name="InsertRevisions"/> -->
      <!-- input: content.xml -->
      <!-- output: xl/revisions/revisionHeaders.xml
                                xl/revisions/revisionLog_N_.xml
                                xl/revisions/userNames.xml
                                xl/revisions/_rels/revisionHeaders.xml.rels
            -->

      <!-- connections -->
      <!--<pzip:entry pzip:target="xl/connections.xml">
                <xsl:call-template name="InsertConnections"/>
            </pzip:entry>-->
      <!-- input: content.xml -->

      <!-- queryTables -->
      <!--<xsl:call-template name="InsertQueryTables"/>-->
      <!-- input: content.xml -->
      <!-- output:  xl/queryTables/queryTableN.xml -->

      <!-- settings  -->
      <!-- <pzip:entry pzip:target="xl/settings.xml">
                <xsl:call-template name="InsertSettings"/>
            </pzip:entry>-->
      <!-- input: content.xml -->

      <!-- insert drawings -->
      <!-- input: content.xml 
        Object_N_/content.xml
        Pictures/_imageFile_
      -->
      <!-- output:  xl/drawings/drawing_N_.xml
        xl/drawings/_rels/drawing_N_.xml.rels
        xl/charts/chart_N_.xml 
        xl/media/_imageFile_ -->

      <!-- /CHANGE -->

      <!-- part relationship item -->
      <pzip:entry pzip:target="xl/_rels/workbook.xml.rels">
        <xsl:call-template name="InsertPartRelationships"/>
      </pzip:entry>

      <!-- Document core properties -->
      <pzip:entry pzip:target="docProps/core.xml">
        <xsl:call-template name="docprops-core"/>
      </pzip:entry>

      <!-- Document app properties -->
      <pzip:entry pzip:target="docProps/app.xml">
        <xsl:call-template name="docprops-app"/>
      </pzip:entry>

      <!-- Document custom properties -->
      <xsl:if test="$docprops-custom-file > 0">
        <pzip:entry pzip:target="docProps/custom.xml">
          <xsl:call-template name="docprops-custom"/>
        </pzip:entry>
      </xsl:if>

      <!-- package relationship item -->
      <pzip:entry pzip:target="_rels/.rels">
        <xsl:call-template name="package-relationships"/>
      </pzip:entry>

      <xsl:call-template name="InsertLinkExternalRels"/>
      
      <!-- Insert Change Tracking -->
      <xsl:if test="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:tracked-changes">
        <xsl:call-template name="CreateRevisionHeadersRels"/>
        <xsl:call-template name="CreateRevisionFiles"/>
        <xsl:call-template name="revisionHeaders"/>
        <xsl:call-template name="userName"/>
      </xsl:if>

    </pzip:archive>
  </xsl:template>

  <xsl:template name="docprops-app">
    <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
      xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
      xmlns:dc="http://purl.org/dc/elements/1.1/">
      <xsl:call-template name="GetDocSecurityExtendedProperty"/>
      <xsl:call-template name="GetApplicationExtendedProperty"/>
      <xsl:for-each select="document('meta.xml')/office:document-meta/office:meta">
        <xsl:apply-templates select="meta:editing-duration"/>
      </xsl:for-each>
    </Properties>
  </xsl:template>

  <xsl:template name="InsertComments">
    <xsl:param name="sheetId"/>
    <pzip:entry
      pzip:target="{concat(&quot;xl/comments&quot;,$sheetId,&quot;.xml&quot;)}">
      <xsl:call-template name="comments">
        <xsl:with-param name="sheetNum" select="$sheetId"/>
      </xsl:call-template>
    </pzip:entry>
  </xsl:template>

  <xsl:template name="CreateSheetRelationships">
    <xsl:param name="comment"/>
    <xsl:param name="picture"/>
    <xsl:param name="hyperlink"/>
    <xsl:param name="chart"/>
    <xsl:param name="textBox"/>

    <!--      <xsl:if
      test="$comment = 'true' or $picture != 'true' or $hyperlink = 'true' or contains($chart,'true')">-->

    <xsl:variable name="OLEObject">
      <xsl:choose>
        <xsl:when
          test="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table/table:shapes/draw:frame/draw:object">
          <xsl:text>true</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:text>false</xsl:text>
        </xsl:otherwise>
      </xsl:choose>

    </xsl:variable>
    <xsl:if
      test="$comment = 'true' or $hyperlink='true' or contains($chart,'true') or $picture = 'true' or $textBox = 'true' or document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table/table:table-row/table:table-cell/table:cell-range-source or $OLEObject = 'true'">
      <!-- package relationship item -->
      <pzip:entry pzip:target="{concat('xl/worksheets/_rels/sheet',position(),'.xml.rels')}">
        <xsl:call-template name="InsertWorksheetsRels">
          <xsl:with-param name="sheetNum" select="position()"/>
          <xsl:with-param name="comment" select="$comment"/>
          <xsl:with-param name="picture" select="$picture"/>
          <xsl:with-param name="hyperlink" select="$hyperlink"/>
          <xsl:with-param name="chart" select="$chart"/>
          <xsl:with-param name="OLEObject">
            <xsl:value-of select="$OLEObject"/>
          </xsl:with-param>
          <xsl:with-param name="textBox" select="$textBox"/>
        </xsl:call-template>
      </pzip:entry>

    </xsl:if>
    <xsl:if test="$OLEObject = 'true'">
      <xsl:call-template name="InsertOLE_rels"/>
      <xsl:call-template name="InsertOLEexternalLinks"/>
      <xsl:call-template name="OLEexternalLinks_rels"/>
    </xsl:if>

  </xsl:template>

  <xsl:template name="CreateDrawingRelationships">
    <!-- package relationship item -->
    <pzip:entry pzip:target="{concat('xl/drawings/_rels/drawing',position(),'.xml.rels')}">
      <xsl:call-template name="InsertDrawingRels">
        <xsl:with-param name="sheetNum" select="position()"/>
      </xsl:call-template>
    </pzip:entry>
  </xsl:template>

  <xsl:template name="CreateVmlDrawing">

    <xsl:variable name="sheetId" select="position()"/>
    <pzip:entry pzip:target="{concat('xl/drawings/vmlDrawing',position(),'.vml')}">
      <pxsi:dummyContainer xmlns:pxsi="urn:cleverage:xmlns:post-processings:comments">
        <o:shapelayout v:ext="edit">
          <o:idmap v:ext="edit" data="1"/>
        </o:shapelayout>
        <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"
          path="m,l,21600r21600,l21600,xe">
          <v:stroke joinstyle="miter"/>
          <v:path gradientshapeok="t" o:connecttype="rect"/>
        </v:shapetype>
        <xsl:for-each
          select="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table[$sheetId]/descendant::table:table-row/table:table-cell/office:annotation">
          <xsl:call-template name="InsertTextBox"/>
        </xsl:for-each>

        <xsl:for-each
          select="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table/table:shapes/draw:frame">

          <xsl:variable name="width">
            <xsl:call-template name="point-measure">
              <xsl:with-param name="length" select="@svg:width"/>
            </xsl:call-template>
          </xsl:variable>

          <xsl:variable name="height">
            <xsl:call-template name="point-measure">
              <xsl:with-param name="length" select="@svg:height"/>
            </xsl:call-template>
          </xsl:variable>

          <xsl:variable name="z-index">
            <xsl:value-of select="position()"/>
          </xsl:variable>

          <xsl:variable name="margin-left">
            <xsl:call-template name="point-measure">
              <xsl:with-param name="length" select="@svg:x"/>
            </xsl:call-template>
          </xsl:variable>

          <xsl:variable name="margin-top">
            <xsl:call-template name="point-measure">
              <xsl:with-param name="length" select="@svg:y"/>
            </xsl:call-template>
          </xsl:variable>

          <v:shape type="#_x0000_t75"
            style="position:absolute;
            margin-left:{$margin-left}pt;margin-top:{$margin-top}pt;width:{$width}pt;height:{$height}pt;z-index:{$z-index}"
            filled="t" fillcolor="window [65]" stroked="t" strokecolor="windowText [64]"
            o:insetmode="auto">

            <xsl:attribute name="id">
              <xsl:value-of select="concat('_x0000_s', 1025 + position())"/>
            </xsl:attribute>
            <v:fill color2="window [65]"/>
            <v:imagedata o:relid="rId1" o:title=""/>
            <x:ClientData ObjectType="Pict">
              <x:SizeWithCells/>
              <x:CF>Pict</x:CF>
              <x:AutoPict/>
              <x:DDE/>
            </x:ClientData>
          </v:shape>
        </xsl:for-each>
      </pxsi:dummyContainer>
    </pzip:entry>
  </xsl:template>

  <xsl:template name="CreateDrawing">
    <pzip:entry pzip:target="{concat('xl/drawings/drawing',position(),'.xml')}">
      <xsl:call-template name="InsertDrawing"/>
    </pzip:entry>
  </xsl:template>
  
  <xsl:template name="CreateRevisionHeadersRels">
    <pzip:entry pzip:target="xl/revisions/_rels/revisionHeaders.xml.rels">
      <xsl:call-template name="InsertRevisionsRels"/>
    </pzip:entry>
  </xsl:template>
  
  
</xsl:stylesheet>

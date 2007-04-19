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
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:x="urn:schemas-microsoft-com:office:excel"
  exclude-result-prefixes="odf style text number">

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

  <xsl:strip-space elements="*"/>
  <xsl:preserve-space elements="text:p text:span number:text"/>

  <xsl:param name="outputFile"/>
  <xsl:output method="xml" encoding="UTF-8"/>

  <!-- App version number -->
  <!-- WARNING: it has to be of type xx.yy -->
  <!-- (otherwise Word cannot open the doc) -->
  <xsl:variable name="app-version">1.00</xsl:variable>

  <!-- existence of docProps/custom.xml file -->
  <xsl:variable name="docprops-custom-file"
    select="count(document('meta.xml')/office:document-meta/office:meta/meta:user-defined)"
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"/>

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
    
      <!-- shared strings (ewentualny postprocessing)-->
      <pzip:entry pzip:target="xl/sharedStrings.xml">
        <xsl:call-template name="InsertSharedStrings"/>
      </pzip:entry>
      <!-- input: content.xml -->
      <!-- output: xl/sharedStrings.xml -->

      
      <!-- insert sheets -->
      <xsl:call-template name="InsertSheets"/>
      <!-- input: content.xml -->
      <!-- output:  xl/worksheets/sheet_N_.xml -->
      
      <!--insert comments-->
      <xsl:for-each select="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table">
        <xsl:if test="descendant::office:annotation">
          <xsl:call-template name="InsertComments">
            <xsl:with-param name="sheetId" select="position()"/>
          </xsl:call-template>
          
          <!-- package relationship item -->
          <pzip:entry pzip:target="{concat('xl/worksheets/_rels/sheet',position(),'.xml.rels')}">
         
            <xsl:call-template name="InsertWorksheetsRels">
                <xsl:with-param name="sheetNum" select="position()"/>
            </xsl:call-template>
          </pzip:entry>
          
        </xsl:if>
       </xsl:for-each>
      
      <xsl:for-each select="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table">
        <xsl:if test="descendant::office:annotation">
          <xsl:variable name="sheetId" select="position()"/>
          <pzip:entry pzip:target="{concat('xl/drawings/vmlDrawing',position(),'.vml')}">
          <pxsi:dummyContainer  xmlns:pxsi="urn:cleverage:xmlns:post-processings:comments">
            <o:shapelayout v:ext="edit">
              <o:idmap v:ext="edit" data="1"/>
            </o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"
              path="m,l,21600r21600,l21600,xe">
              <v:stroke joinstyle="miter"/>
              <v:path gradientshapeok="t" o:connecttype="rect"/>
            </v:shapetype>
            <xsl:for-each select="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table[$sheetId]/table:table-row/table:table-cell/office:annotation">
              <xsl:call-template name="InsertTextBox"/>
            </xsl:for-each>
          </pxsi:dummyContainer>
          </pzip:entry>
        </xsl:if>
      </xsl:for-each>
      
      <!-- insert drawings -->
      <!--<xsl:call-template name="InsertDrawings"/>-->
      <!-- input: content.xml 
                              Object_N_/content.xml
                              Pictures/_imageFile_
            -->
      <!-- output:  xl/drawings/drawing_N_.xml
                                 xl/drawings/_rels/drawing_N_.xml.rels
                                 xl/charts/chart_N_.xml 
                                 xl/media/_imageFile_ -->

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

      <!-- part relationship item -->
      <!--<pzip:entry pzip:target="xl/_rels/document.xml.rels">
                <xsl:call-template name="InsertPartRelationships"/>
            </pzip:entry> -->
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
      
    </pzip:archive>
  </xsl:template>

 <xsl:template name="docprops-app">
   <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
     xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
     xmlns:dc="http://purl.org/dc/elements/1.1/">
   <xsl:call-template name="GetDocSecurityExtendedProperty"/>
   <xsl:call-template name="GetApplicationExtendedProperty"/>
  </Properties>
</xsl:template>
  
  <xsl:template name="InsertComments">
  <xsl:param name="sheetId"/>
    <pzip:entry pzip:target='{concat("xl/comments",$sheetId,".xml")}'>
      <xsl:call-template name="comments">
          <xsl:with-param name="sheetNum" select="$sheetId"/>
      </xsl:call-template>
    </pzip:entry>
   </xsl:template>
</xsl:stylesheet>

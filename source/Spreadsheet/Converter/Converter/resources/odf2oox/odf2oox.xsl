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
    xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
    exclude-result-prefixes="odf style text number">
    
    <xsl:import href="workbook.xsl"/>
    <xsl:import href="sharedStrings.xsl"/>
    <xsl:import href="odf2oox-compute-size.xsl"/>
    
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
   
 <!-- CHANGE -->          
            <!-- styles -->
            <!--<pzip:entry pzip:target="xl/styles.xml">
                <xsl:call-template name="styles"/>
            </pzip:entry>-->
            <!-- input: styles.xml -->
            
            <!-- main content -->
            <pzip:entry pzip:target="xl/workbook.xml">
                <xsl:call-template name="InsertWorkbook"/>
            </pzip:entry>
            <!-- input: content.xml -->
            
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
            
        </pzip:archive>
    </xsl:template>
    
    
</xsl:stylesheet>

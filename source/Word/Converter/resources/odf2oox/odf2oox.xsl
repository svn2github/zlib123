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

  <xsl:import href="measures.xsl"/>
  <xsl:import href="2oox-common.xsl"/>
  <xsl:import href="2oox-docprops.xsl"/>
  <xsl:import href="2oox-document.xsl"/>
  <xsl:import href="2oox-numbering.xsl"/>
  <xsl:import href="2oox-footnotes.xsl"/>
  <xsl:import href="2oox-endnotes.xsl"/>
  <xsl:import href="2oox-header-footer.xsl"/>
  <xsl:import href="2oox-fonts.xsl"/>
  <xsl:import href="2oox-styles.xsl"/>
  <xsl:import href="2oox-dataStyles.xsl"/>
  <xsl:import href="2oox-frames.xsl"/>
  <xsl:import href="2oox-ole.xsl"/>
  <xsl:import href="2oox-settings.xsl"/>
  <xsl:import href="2oox-sections.xsl"/>
  <xsl:import href="2oox-part_relationships.xsl"/>
  <xsl:import href="2oox-package_relationships.xsl"/>
  <xsl:import href="2oox-contentTypes.xsl"/>
  <xsl:import href="2oox-comments.xsl"/>
  <xsl:import href="2oox-change-tracking.xsl"/>
 
  <xsl:strip-space elements="*"/>
  <xsl:preserve-space elements="text:p text:span number:text"/>
  
  <xsl:param name="outputFile"/>
  <xsl:output method="xml" encoding="UTF-8"/>

  <!-- App version number -->
  <!-- WARNING: it has to be of type xx.yy -->
  <!-- (otherwise Word cannot open the doc) -->
  <xsl:variable name="app-version">2.00</xsl:variable>

  <!-- existence of docProps/custom.xml file -->
  <xsl:variable name="docprops-custom-file"
    select="count(document('meta.xml')/office:document-meta/office:meta/meta:user-defined)"
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"/>

  <xsl:template match="/odf:source">
    <xsl:processing-instruction name="mso-application">progid="Word.Document"</xsl:processing-instruction>

   
    <pzip:archive pzip:target="{$outputFile}">

      <!-- sections preformatting -->
      <xsl:call-template name="sectionsPreProcessing"/>
      
      
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

      <!-- styles -->
      <pzip:entry pzip:target="word/styles.xml">
        <xsl:call-template name="styles"/>
      </pzip:entry>

      <!-- main content -->
      <pzip:entry pzip:target="word/document.xml">
        <xsl:call-template name="document"/>
      </pzip:entry>

      <!-- numbering (lists) -->
      <pzip:entry pzip:target="word/numbering.xml">
        <xsl:call-template name="numbering"/>
      </pzip:entry>

      <!-- numbering relationships item-->
      <pzip:entry pzip:target="word/_rels/numbering.xml.rels">
        <xsl:call-template name="InsertNumberingInternalRelationships"/>
      </pzip:entry>
      
      <!-- footnotes -->
      <pzip:entry pzip:target="word/footnotes.xml">
        <xsl:call-template name="footnotes"/>
      </pzip:entry>

      <!-- footnotes part relationships -->
      <pzip:entry pzip:target="word/_rels/footnotes.xml.rels">
        <xsl:call-template name="InsertFootnotesInternalRelationships"/>
      </pzip:entry>

      <!-- endnotes -->
      <pzip:entry pzip:target="word/endnotes.xml">
        <xsl:call-template name="endnotes"/>
      </pzip:entry>

      <!-- endnotes part relationships -->
      <pzip:entry pzip:target="word/_rels/endnotes.xml.rels">
        <xsl:call-template name="InsertEndnotesInternalRelationships"/>
      </pzip:entry>

      <!-- Comment   -->
      <pzip:entry pzip:target="word/comments.xml">
        <xsl:call-template name="comments"/>
      </pzip:entry>

      <!-- headers and footers -->
      <xsl:call-template name="InsertHeaderFooterParts"/>

      <!-- fonts declaration -->
      <pzip:entry pzip:target="word/fontTable.xml">
        <xsl:call-template name="fonts"/>
      </pzip:entry>

      <!-- settings  -->
      <pzip:entry pzip:target="word/settings.xml">
        <xsl:call-template name="InsertSettings"/>
      </pzip:entry>

      <!-- part relationship item -->
      <pzip:entry pzip:target="word/_rels/document.xml.rels">
        <xsl:call-template name="InsertPartRelationships"/>
      </pzip:entry>

      <!-- content types -->
      <pzip:entry pzip:target="[Content_Types].xml">
        <xsl:call-template name="contentTypes"/>
      </pzip:entry>

      <!-- package relationship item -->
      <pzip:entry pzip:target="_rels/.rels">
        <xsl:call-template name="package-relationships"/>
      </pzip:entry>

    </pzip:archive>
  </xsl:template>


</xsl:stylesheet>

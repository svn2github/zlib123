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
<xsl:stylesheet version="1.0" 
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
  xmlns:oox="urn:oox"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns:rels="http://schemas.openxmlformats.org/package/2006/relationships"
  exclude-result-prefixes="oox rels">

  <xsl:import href="common.xsl"/>
  <xsl:import href="content.xsl"/>
  <xsl:import href="pictures.xsl"/>
  <xsl:import href="styles.xsl"/>
  <xsl:import href="meta.xsl"/>
  <xsl:import href="settings.xsl"/>
  <xsl:import href="relationships.xsl"/>
  <xsl:import href="footnotes.xsl"/>
  <xsl:import href="sections.xsl"/>
  <xsl:import href="comments.xsl"/>

  <xsl:param name="outputFile"/>
  <xsl:output method="xml" encoding="UTF-8" indent="yes"/>

  
  <!-- packages relationships -->
  <!--
  <xsl:variable name="package-rels" select="document('_rels/.rels')"/>
  <xsl:variable name="officeDocument"
    select="string($package-rels/rels:Relationships/rels:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument']/@Target)"/>
  <xsl:variable name="core-properties"
    select="string($package-rels/rels:Relationships/rels:Relationship[@Type='http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties']/@Target)"/>
  <xsl:variable name="extended-properties"
    select="string($package-rels/rels:Relationships/rels:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties']/@Target)"/>
  <xsl:variable name="custom-properties"
    select="string($package-rels/rels:Relationships/rels:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties']/@Target)"/>
  
  <xsl:variable name="document-path" select="concat(substring-before($officeDocument, '/'), '/')"/>
    -->
  <!-- part relationships -->
  <!-- TODO multilevel /.../.../ -->
  <!--
  <xsl:variable name="part-relationships"
    select="concat(concat($document-path, '_rels/'), concat(substring-after($officeDocument, '/'), '.rels'))"/>
  <xsl:variable name="numbering"
    select="concat($document-path, document($part-relationships)/rels:Relationships/rels:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering']/@Target)"/>
  <xsl:variable name="styles"
    select="concat($document-path, document($part-relationships)/rels:Relationships/rels:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles']/@Target)"/>
  <xsl:variable name="fontTable"
    select="concat($document-path, document($part-relationships)/rels:Relationships/rels:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable']/@Target)"/>
  <xsl:variable name="settings"
    select="concat($document-path, document($part-relationships)/rels:Relationships/rels:Relationship[@rType='http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings']/@Target)"/>
  <xsl:variable name="footnotes"
    select="concat($document-path, document($part-relationships)/rels:Relationships/rels:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes']/@Target)"/>
  <xsl:variable name="endnotes"
    select="concat($document-path, document($part-relationships)/rels:Relationships/rels:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes']/@Target)"/>
  <xsl:variable name="comments"
    select="concat($document-path, document($part-relationships)/rels:Relationships/rels:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments']/@Target)"/>
  -->


  <!-- App version number -->
  <!-- WARNING: it has to be of type xx.yy -->
  <!-- (otherwise Word cannot open the doc) -->
  <xsl:variable name="app-version">0.3</xsl:variable>

  <xsl:template match="/oox:source">

    <pzip:archive pzip:target="{$outputFile}">

      <!-- Manifest -->
      <pzip:entry pzip:target="META-INF/manifest.xml">
        <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
          <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.text"
            manifest:full-path="/"/>
          <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
          <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml"/>
        </manifest:manifest>
      </pzip:entry>

      <!-- main content -->
      <pzip:entry pzip:target="content.xml">
        <xsl:call-template name="content"/>
      </pzip:entry>

      <!-- styles -->
      <pzip:entry pzip:target="styles.xml">
        <xsl:call-template name="styles"/>
      </pzip:entry>

      <!-- meta -->
      <pzip:entry pzip:target="meta.xml">
        <xsl:call-template name="meta"/>
      </pzip:entry>

      <!-- settings -->
      <pzip:entry pzip:target="settings.xml">
        <xsl:call-template name="settings"/>
      </pzip:entry>

    </pzip:archive>
  </xsl:template>


</xsl:stylesheet>

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
  xmlns:rels="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  exclude-result-prefixes="oox rels">

  <xsl:import href="measures.xsl"/>
  <xsl:import href="common-meta.xsl"/>
  <xsl:import href="utils.xsl"/>
  <xsl:import href="2odf-common.xsl"/>
  <xsl:import href="2odf-content.xsl"/>
  <xsl:import href="2odf-pictures.xsl"/>
  <xsl:import href="2odf-styles.xsl"/>
  <xsl:import href="2odf-frames.xsl"/>
  <xsl:import href="2odf-shapes.xsl"/>
  <xsl:import href="2odf-ole.xsl"/>
  <xsl:import href="2odf-settings.xsl"/>
  <xsl:import href="2odf-relationships.xsl"/>
  <xsl:import href="2odf-footnotes.xsl"/>
  <xsl:import href="2odf-sections.xsl"/>
  <xsl:import href="2odf-comments.xsl"/>
  <xsl:import href="2odf-track.xsl"/>

  <xsl:param name="outputFile"/>

  <!-- a string containing detailed information on environment and
       converter version to be added to the document's meta data -->
  <xsl:param name="generator"/>
  
  <xsl:output method="xml" encoding="UTF-8"/>

  <xsl:key name="Part" match="/oox:package/oox:part" use="@oox:name"/>
  
  <!-- packages relationships -->
  <!--
  <xsl:variable name="package-rels" select="/oox:package/oox:part[@oox:name='_rels/.rels']"/>
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
    select="concat($document-path, document($part-relationships)/rels:Relationships/rels:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings']/@Target)"/>
  <xsl:variable name="footnotes"
    select="concat($document-path, document($part-relationships)/rels:Relationships/rels:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes']/@Target)"/>
  <xsl:variable name="endnotes"
    select="concat($document-path, document($part-relationships)/rels:Relationships/rels:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes']/@Target)"/>
  <xsl:variable name="comments"
    select="concat($document-path, document($part-relationships)/rels:Relationships/rels:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments']/@Target)"/>
  -->

  <!-- App version number -->
  <xsl:variable name="app-version">2.0.0</xsl:variable>
  
  <xsl:template match="/oox:package">

    <pzip:archive pzip:target="{$outputFile}">

      <!-- Manifest -->
      <pzip:entry pzip:target="META-INF/manifest.xml">
        <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
          <manifest:file-entry manifest:full-path="/">
            <xsl:attribute name="manifest:media-type">
              <xsl:choose>
                <xsl:when test="substring($outputFile, string-length($outputFile) - 2) = 'ott'">
                  <xsl:value-of select="'application/vnd.oasis.opendocument.text-template'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'application/vnd.oasis.opendocument.text'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </manifest:file-entry>
          <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
          <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml"/>
          <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="meta.xml"/>
          <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="settings.xml"/>
          
          <!-- copy relationships from document -->
          <xsl:for-each select="key('Part', 'word/_rels/document.xml.rels')/rels:Relationships/rels:Relationship[substring-before(@Target,'/')='media' or substring-before(@Target,'/')='embeddings']">
            <xsl:call-template name="InsertManifestFileEntry"/>
          </xsl:for-each>
          <!-- divo TODO: does this work for more than one footer/header (i.e. if we have footer2.xml.rel etc) ??? -->
          <xsl:for-each select="key('Part', 'word/_rels/footer1.xml.rels')/rels:Relationships/rels:Relationship[substring-before(@Target,'/')='media' or substring-before(@Target,'/')='embeddings']">
            <xsl:call-template name="InsertManifestFileEntry"/>
          </xsl:for-each>
          <xsl:for-each select="key('Part', 'word/_rels/header1.xml.rels')/rels:Relationships/rels:Relationship[substring-before(@Target,'/')='media' or substring-before(@Target,'/')='embeddings']">
            <xsl:call-template name="InsertManifestFileEntry"/>
          </xsl:for-each>
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
        <xsl:call-template name="meta">
          <xsl:with-param name="app-version" select="$app-version" />
          <xsl:with-param name="generator" select="$generator" />
        </xsl:call-template>
      </pzip:entry>

      <!-- settings -->
      <pzip:entry pzip:target="settings.xml">
        <xsl:call-template name="settings"/>
      </pzip:entry>

    </pzip:archive>
  </xsl:template>

  
  <!--
  Summary: Inserts the manifest entry for a single reference
  Author: Clever Age
  Modified: makz (DIaLOGIKa)
  Date: 14.11.2007
  -->
  <xsl:template name="InsertManifestFileEntry">
    <manifest:file-entry>

      <xsl:variable name="thisId" select="@Id" />
      <xsl:variable name="suffix" select="translate(substring-after(@Target,'.'), 
                    'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />

      <!-- write the media type -->
      <xsl:choose>
        <!-- GIF -->
        <xsl:when test="$suffix='gif'">
          <xsl:attribute name="manifest:media-type">
            <xsl:text>image/gif</xsl:text>
          </xsl:attribute>
        </xsl:when>
        <!-- JPEG -->
        <xsl:when test="$suffix='jpg' or $suffix='jpeg' or $suffix='jpe' or $suffix='jfif'">
          <xsl:attribute name="manifest:media-type">
            <xsl:text>image/jpeg</xsl:text>
          </xsl:attribute>
        </xsl:when>
        <!-- TIFF -->
        <xsl:when test="$suffix='tif' or $suffix='tiff'">
          <xsl:attribute name="manifest:media-type">
            <xsl:text>image/tiff</xsl:text>
          </xsl:attribute>
        </xsl:when>
        <!-- PNG -->
        <xsl:when test="$suffix='png'">
          <xsl:attribute name="manifest:media-type">
            <xsl:text>image/png</xsl:text>
          </xsl:attribute>
        </xsl:when>
        <!-- EMF -->
        <xsl:when test="$suffix='emf'">
          <xsl:attribute name="manifest:media-type">
            <xsl:text>application/x-openoffice-wmf;windows_formatname="Image EMF"</xsl:text>
          </xsl:attribute>
        </xsl:when>
        <!-- WMF -->
        <xsl:when test="$suffix='wmf'">
          <xsl:attribute name="manifest:media-type">
            <xsl:text>application/x-openoffice-wmf;windows_formatname="Image WMF"</xsl:text>
          </xsl:attribute>
        </xsl:when>
        <!-- Binaries OLE -->
        <xsl:when test="$suffix='bin' or $suffix='xls' or $suffix='doc' or $suffix='ppt'">
          <xsl:attribute name="manifest:media-type">
            <xsl:text>application/vnd.sun.star.oleobject</xsl:text>
          </xsl:attribute>
        </xsl:when>
      </xsl:choose>

      <!-- write the path -->
      <xsl:choose>
        <!-- the ref is a ole -->
        <xsl:when test="@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject'">
          <xsl:attribute name="manifest:full-path">
            <xsl:value-of select="substring-before(substring-after(@Target,'/'), '.')"/>
          </xsl:attribute>
        </xsl:when>
        <!-- the ref is ole picture -->
        <xsl:when test="key('Part', 'word/document.xml')/w:document/w:body//v:shape/v:imagedata[./@r:id=$thisId]">
          <xsl:attribute name="manifest:full-path">

            <xsl:text>ObjectReplacements/</xsl:text>
            <xsl:call-template name="InsertOlePreviewName">
              <xsl:with-param name="thisId" select="$thisId" />
              <xsl:with-param name="relFile" select="ancestor::oox:part/@oox:name" />
            </xsl:call-template>

          </xsl:attribute>
        </xsl:when>
        <!-- the ref is a normal picture -->
        <xsl:otherwise>
          <xsl:attribute name="manifest:full-path">
            <xsl:text>Pictures/</xsl:text>
            <xsl:value-of select="substring-after(@Target,'/')"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>

    </manifest:file-entry>
  </xsl:template>

</xsl:stylesheet>

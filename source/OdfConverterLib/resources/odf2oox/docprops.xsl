<?xml version="1.0" encoding="UTF-8" ?>
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
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" exclude-result-prefixes="office meta">

  <xsl:template name="docprops-core">
    <cp:coreProperties
      xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
      xmlns:dcmitype="http://purl.org/dc/dcmitype/">
      <xsl:apply-templates select="document('meta.xml')/office:document-meta/office:meta"
        mode="core"/>
    </cp:coreProperties>
  </xsl:template>

  <xsl:template name="docprops-app">
    <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
      xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
      <xsl:apply-templates select="document('meta.xml')/office:document-meta/office:meta" mode="app"
      />
    </Properties>
  </xsl:template>

  <xsl:template name="docprops-custom">
    <xsl:apply-templates select="document('meta.xml')/office:document-meta/office:meta"
      mode="custom"/>
  </xsl:template>

  <xsl:template match="/office:document-meta/office:meta" mode="core">
    <!-- creation date -->
    <xsl:if test="meta:creation-date">
      <dcterms:created xmlns:dcterms="http://purl.org/dc/terms/"
        xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="dcterms:W3CDTF">
        <xsl:value-of select="meta:creation-date"/>
      </dcterms:created>
    </xsl:if>
    <!-- creator -->
    <xsl:if test="meta:initial-creator">
      <dc:creator xmlns:dc="http://purl.org/dc/elements/1.1/">
        <xsl:value-of select="meta:initial-creator"/>
      </dc:creator>
    </xsl:if>
    <!-- description -->
    <xsl:if test="dc:description">
      <dc:description xmlns:dc="http://purl.org/dc/elements/1.1/">
        <xsl:value-of select="dc:description"/>
      </dc:description>
    </xsl:if>
    <!-- identifier -->
    <xsl:if test="dc:identifier">
      <dc:identifier xmlns:dc="http://purl.org/dc/elements/1.1/">
        <xsl:value-of select="dc:identifier"/>
      </dc:identifier>
    </xsl:if>
    <!-- keywords -->
    <xsl:if test="meta:keyword">
      <cp:keywords
        xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
        <xsl:apply-templates select="meta:keyword"/>
      </cp:keywords>
    </xsl:if>
    <!-- language -->
    <xsl:if test="dc:language">
      <dc:language xmlns:dc="http://purl.org/dc/elements/1.1/">
        <xsl:value-of select="dc:language"/>
      </dc:language>
    </xsl:if>
    <!-- last modification author -->
    <xsl:if test="dc:creator">
      <cp:lastModifiedBy
        xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
        <xsl:value-of select="dc:creator"/>
      </cp:lastModifiedBy>
    </xsl:if>
    <!-- last printing -->
    <xsl:if test="meta:printed-date">
      <cp:lastPrinted
        xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
        <xsl:value-of select="meta:printed-date"/>
      </cp:lastPrinted>
    </xsl:if>
    <!-- last modification date -->
    <xsl:if test="dc:date">
      <dcterms:modified xmlns:dcterms="http://purl.org/dc/terms/"
        xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="dcterms:W3CDTF">
        <xsl:value-of select="dc:date"/>
      </dcterms:modified>
    </xsl:if>
    <!-- number of times it was saved -->
    <xsl:if test="meta:editing-cycles">
      <cp:revision
        xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
        <xsl:value-of select="meta:editing-cycles"/>
      </cp:revision>
    </xsl:if>
    <!-- topic -->
    <xsl:if test="dc:subject">
      <dc:subject xmlns:dc="http://purl.org/dc/elements/1.1/">
        <xsl:value-of select="dc:subject"/>
      </dc:subject>
    </xsl:if>
    <!-- title -->
    <xsl:if test="dc:title">
      <dc:title xmlns:dc="http://purl.org/dc/elements/1.1/">
        <xsl:value-of select="dc:title"/>
      </dc:title>
    </xsl:if>
  </xsl:template>

  <xsl:template match="/office:document-meta/office:meta/meta:keyword">
    <xsl:if test="not(position() = 1)">
      <xsl:text> </xsl:text>
    </xsl:if>
    <xsl:value-of select="."/>
  </xsl:template>

  <xsl:template match="/office:document-meta/office:meta" mode="app">
    <xsl:if test="meta:document-statistic/@meta:page-count">
      <Pages xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
        <xsl:value-of select="meta:document-statistic/@meta:page-count"/>
      </Pages>
    </xsl:if>
    <xsl:if test="meta:document-statistic/@meta:word-count">
      <Words xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
        <xsl:value-of select="meta:document-statistic/@meta:word-count"/>
      </Words>
    </xsl:if>
    <Application xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">ODF Converter</Application>
    <DocSecurity xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">0</DocSecurity>
    <xsl:if test="meta:document-statistic/@meta:paragraph-count">
      <Paragraphs xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
        <xsl:value-of select="meta:document-statistic/@meta:paragraph-count"/>
      </Paragraphs>
    </xsl:if>
    <!-- editing duration -->
    <xsl:if test="meta:editing-duration">
      <xsl:variable name="hours">
        <xsl:choose>
          <xsl:when test="contains(meta:editing-duration, 'H')">
            <xsl:value-of
              select="substring-before(substring-after(meta:editing-duration, 'T'), 'H')"/>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:variable name="minutes">
        <xsl:choose>
          <xsl:when test="contains(meta:editing-duration, 'H')">
            <xsl:value-of
              select="substring-before(substring-after(meta:editing-duration, 'H'), 'M')"/>
          </xsl:when>
          <xsl:when test="contains(meta:editing-duration, 'M')">
            <xsl:value-of
              select="substring-before(substring-after(meta:editing-duration, 'T'), 'M')"/>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:variable name="seconds">
        <xsl:choose>
          <xsl:when test="contains(meta:editing-duration, 'M')">
            <xsl:value-of
              select="substring-before(substring-after(meta:editing-duration, 'M'), 'S')"/>
          </xsl:when>
          <xsl:when test="contains(meta:editing-duration, 'S')">
            <xsl:value-of
              select="substring-before(substring-after(meta:editing-duration, 'T'), 'S')"/>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:if test="number($hours) and number($minutes) and number($seconds)">
        <TotalTime xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
          <xsl:choose>
            <xsl:when test="number($seconds) &gt; 30">
              <xsl:value-of select="60 * number($hours) + number($minutes) + 1"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="60 * number($hours) + number($minutes)"/>
            </xsl:otherwise>
          </xsl:choose>
        </TotalTime>
      </xsl:if>
    </xsl:if>
    <ScaleCrop xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">false</ScaleCrop>
    <LinksUpToDate xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">false</LinksUpToDate>
    <xsl:if test="meta:document-statistic/@meta:character-count">
      <CharactersWithSpaces xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
        <xsl:value-of select="meta:document-statistic/@meta:character-count"/>
      </CharactersWithSpaces>
    </xsl:if>
    <SharedDoc xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">false</SharedDoc>
    <HyperlinksChanged xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">false</HyperlinksChanged>
    <AppVersion xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
      <xsl:value-of select="$app-version"/>
    </AppVersion>
  </xsl:template>

  <xsl:template match="/office:document-meta/office:meta" mode="custom">
    <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
      xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
      <xsl:apply-templates select="meta:user-defined"/>
    </Properties>
  </xsl:template>

  <xsl:template match="/office:document-meta/office:meta/meta:user-defined">
    <property xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
      xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
      <xsl:attribute name="fmtid">{D5CDD505-2E9C-101B-9397-08002B2CF9AE}</xsl:attribute>
      <xsl:attribute name="pid">
        <xsl:value-of select="position() + 1"/>
      </xsl:attribute>
      <xsl:attribute name="name">
        <xsl:value-of select="@meta:name"/>
      </xsl:attribute>
      <vt:lpwstr>
        <xsl:value-of select="text()"/>
      </vt:lpwstr>
    </property>
  </xsl:template>

</xsl:stylesheet>

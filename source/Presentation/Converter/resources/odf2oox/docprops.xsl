<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
<!--
Copyright (c) 2007, Sonata Software Limited
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
*     * Neither the name of Sonata Software Limited nor the names of its contributors
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
* SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" exclude-result-prefixes="office meta">

  <xsl:template name="docprops-core">
    <!--<cp:coreProperties
      xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
      xmlns:dcmitype="http://purl.org/dc/dcmitype/">
      <xsl:apply-templates select="document('meta.xml')/office:document-meta/office:meta"
        mode="core"/>
		
    </cp:coreProperties-->>	
	  <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
		  xmlns:dc="http://purl.org/dc/elements/1.1/" 
		  xmlns:dcterms="http://purl.org/dc/terms/" 
		  xmlns:dcmitype="http://purl.org/dc/dcmitype/" 
		  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
		  <dc:title>Title</dc:title>
		  <dc:creator></dc:creator>
		  <cp:lastModifiedBy>pradeep.n</cp:lastModifiedBy>
		  <cp:revision>1</cp:revision>
		  <dcterms:created xsi:type="dcterms:W3CDTF">2006-08-16T00:00:00Z</dcterms:created>
		  <dcterms:modified xsi:type="dcterms:W3CDTF">2007-04-05T10:53:16Z</dcterms:modified>
	  </cp:coreProperties>
  </xsl:template>

  <xsl:template name="docprops-app">
    <!--<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
      xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
      <xsl:apply-templates select="document('meta.xml')/office:document-meta/office:meta" mode="app"
      />
    </Properties>-->
	  <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" 
		  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
		  xmlns:dc="http://purl.org/dc/elements/1.1/">
		  <TotalTime>0</TotalTime>
		  <Words>3</Words>
		  <Application>Microsoft Office PowerPoint</Application>
		  <PresentationFormat>On-screen Show (4:3)</PresentationFormat>
		  <Paragraphs>2</Paragraphs>
		  <Slides>1</Slides>
		  <Notes>0</Notes>
		  <HiddenSlides>0</HiddenSlides>
		  <MMClips>0</MMClips>
		  <ScaleCrop>false</ScaleCrop>
		  <HeadingPairs>
			  <vt:vector size="4" baseType="variant">
				  <vt:variant>
					  <vt:lpstr>Theme</vt:lpstr>
				  </vt:variant>
				  <vt:variant>
					  <vt:i4>1</vt:i4>
				  </vt:variant>
				  <vt:variant>
					  <vt:lpstr>Slide Titles</vt:lpstr>
				  </vt:variant>
				  <vt:variant>
					  <vt:i4>1</vt:i4>
				  </vt:variant>
			  </vt:vector>
		  </HeadingPairs>
		  <TitlesOfParts>
			  <vt:vector size="2" baseType="lpstr">
				  <vt:lpstr>Office Theme</vt:lpstr>
				  <vt:lpstr>Title</vt:lpstr>
			  </vt:vector>
		  </TitlesOfParts>
		  <Company></Company>
		  <LinksUpToDate>false</LinksUpToDate>
		  <SharedDoc>false</SharedDoc>
		  <HyperlinksChanged>false</HyperlinksChanged>
		  <AppVersion>12.0000</AppVersion>
	  </Properties>
  </xsl:template>

  <xsl:template name="docprops-custom">
    <xsl:apply-templates select="document('meta.xml')/office:document-meta/office:meta"
      mode="custom"/>
  </xsl:template>

  <xsl:template match="/office:document-meta/office:meta" mode="core">
    <!-- report lost properties -->
    <xsl:if test="meta:auto-reload">
      <xsl:message terminate="no">translation.odf2oox.internetProperties</xsl:message>
    </xsl:if>
    <xsl:if test="meta:hyperlink-behaviour">
      <xsl:message terminate="no">translation.odf2oox.internetProperties</xsl:message>
    </xsl:if>
    <!-- creation date -->
    <xsl:if test="meta:creation-date">
      <xsl:variable name="dateIsValid">
        <xsl:call-template name="validateDate">
          <xsl:with-param name="date">
            <xsl:value-of select="meta:creation-date"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$dateIsValid != 'false' ">
          <dcterms:created xmlns:dcterms="http://purl.org/dc/terms/"
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="dcterms:W3CDTF">
            <xsl:value-of select="$dateIsValid"/>
          </dcterms:created>
        </xsl:when>
        <xsl:otherwise>
          <xsl:message terminate="no">translation.odf2oox.documentCreationDate</xsl:message>
        </xsl:otherwise>
      </xsl:choose>
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
      <xsl:variable name="dateIsValid">
        <xsl:call-template name="validateDate">
          <xsl:with-param name="date">
            <xsl:value-of select="dc:date"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$dateIsValid != 'false' ">
          <dcterms:modified xmlns:dcterms="http://purl.org/dc/terms/"
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="dcterms:W3CDTF">
            <xsl:value-of select="$dateIsValid"/>
          </dcterms:modified>
        </xsl:when>
        <xsl:otherwise>
          <xsl:message terminate="no">translation.odf2oox.documentModificationDate</xsl:message>
        </xsl:otherwise>
      </xsl:choose>
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
    <Application xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
      >ODF Converter</Application>
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
    <ScaleCrop xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
      >false</ScaleCrop>
    <LinksUpToDate xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
      >false</LinksUpToDate>
    <xsl:if test="meta:document-statistic/@meta:character-count">
      <CharactersWithSpaces
        xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
        <xsl:value-of select="meta:document-statistic/@meta:character-count"/>
      </CharactersWithSpaces>
    </xsl:if>
    <xsl:if test="meta:document-statistic/@meta:non-whitespace-character-count">
      <Characters xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
        <xsl:value-of select="meta:document-statistic/@meta:non-whitespace-character-count"/>
      </Characters>
    </xsl:if>
    <SharedDoc xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
      >false</SharedDoc>
    <HyperlinksChanged
      xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">false</HyperlinksChanged>
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

  <!-- check if a date is valid regarding W3C format -->
  <xsl:template name="validateDate">
    <xsl:param name="date"/>
    <!-- year -->
    <xsl:variable name="Y">
      <xsl:choose>
        <xsl:when test="contains($date, '-')">
          <xsl:value-of select="substring-before($date, '-')"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$date"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <!-- month -->
    <xsl:variable name="M">
      <xsl:variable name="date_Y" select="substring-after($date, concat($Y, '-'))"/>
      <xsl:choose>
        <xsl:when test="contains($date_Y, '-')">
          <xsl:value-of select="substring-before($date_Y, '-')"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$date_Y"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <!-- day -->
    <xsl:variable name="D">
      <xsl:variable name="date_Y_m" select="substring-after($date, concat($Y, '-', $M, '-'))"/>
      <xsl:choose>
        <xsl:when test="contains($date_Y_m, 'T')">
          <xsl:value-of select="substring-before($date_Y_m, 'T')"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$date_Y_m"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <!-- hour -->
    <xsl:variable name="h">
      <xsl:if test="contains($date, 'T')">
        <xsl:value-of select="substring-before(substring-after($date, 'T'), ':')"/>
      </xsl:if>
    </xsl:variable>
    <!-- minutes -->
    <xsl:variable name="m">
      <xsl:if test="contains($date, 'T')">
        <xsl:variable name="time_h" select="substring-after(substring-after($date, 'T'), ':')"/>
        <xsl:choose>
          <xsl:when test="contains($time_h, ':')">
            <xsl:value-of select="substring-before($time_h, ':')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$time_h"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
    </xsl:variable>
    <!-- seconds (with our without decimal) -->
    <xsl:variable name="s">
      <xsl:if test="contains($date, 'T')">
        <xsl:variable name="time_h_m"
          select="substring-after(substring-after($date, 'T'), concat($h, ':', $m, ':'))"/>
        <xsl:choose>
          <xsl:when test="contains($time_h_m, 'Z')">
            <xsl:value-of select="substring-before($time_h_m, 'Z')"/>
          </xsl:when>
          <xsl:when test="contains($time_h_m, '+')">
            <xsl:value-of select="substring-before($time_h_m, '+')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$time_h_m"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
    </xsl:variable>

    <!-- validate format -->
    <xsl:choose>
      <xsl:when test="string-length($Y) != 4">false</xsl:when>
      <xsl:when test="$M != '' and string-length($M) != 2">false</xsl:when>
      <xsl:when test="$D != '' and string-length($D) != 2">false</xsl:when>
      <xsl:when test="$h != '' and string-length($h) != 2">false</xsl:when>
      <xsl:when test="$m != '' and string-length($m) != 2">false</xsl:when>
      <xsl:when test="$s != '' and string-length($s) &gt; 4">false</xsl:when>
      <xsl:when test="string-length(substring-after($date, 'Z')) &gt; 0">false</xsl:when>
      <xsl:when test="$M != '' and number($M) &gt; 12">false</xsl:when>
      <xsl:when test="$M != '' and number($M) &lt; 1">false</xsl:when>
      <xsl:when test="$D != '' and number($D) &gt; 31">false</xsl:when>
      <xsl:when test="$D != '' and number($D) &lt; 1">false</xsl:when>
      <xsl:when test="$h != '' and number($h) &gt; 24">false</xsl:when>
      <xsl:when test="$h != '' and number($h) &lt; 1">false</xsl:when>
      <xsl:when test="$m != '' and number($m) &gt; 60">false</xsl:when>
      <xsl:when test="$m != '' and number($m) &lt; 1">false</xsl:when>
      <xsl:when test="$s != '' and number($s) &gt; 61">false</xsl:when>
      <xsl:when test="$s != '' and number($s) &lt; 1">false</xsl:when>
      <xsl:when test="contains($date, 'T') and not(contains($date, 'Z') or contains($date, '+'))">
        <xsl:value-of select="concat($date, 'Z')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$date"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

	<xsl:template name ="presProps">

		<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
			xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
			xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>
	</xsl:template>
	<xsl:template name ="tabStyles">
		<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" >
			<xsl:attribute name ="def">
				<xsl:value-of select ="'{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'"/>
			</xsl:attribute>
		</a:tblStyleLst >
	</xsl:template>
	<xsl:template name ="viewProps">		
		<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
			xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
			xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
			<p:normalViewPr showOutlineIcons="0">
				<p:restoredLeft sz="15620"/>
				<p:restoredTop sz="94660"/>
			</p:normalViewPr>
			<p:slideViewPr>
				<p:cSldViewPr>
					<p:cViewPr varScale="1">
						<p:scale>
							<a:sx n="66" d="100"/>
							<a:sy n="66" d="100"/>
						</p:scale>
						<p:origin x="-600" y="-96"/>
					</p:cViewPr>
					<p:guideLst>
						<p:guide orient="horz" pos="2160"/>
						<p:guide pos="2880"/>
					</p:guideLst>
				</p:cSldViewPr>
			</p:slideViewPr>
			<p:notesTextViewPr>
				<p:cViewPr>
					<p:scale>
						<a:sx n="100" d="100"/>
						<a:sy n="100" d="100"/>
					</p:scale>
					<p:origin x="0" y="0"/>
				</p:cViewPr>
			</p:notesTextViewPr>
			<p:gridSpacing cx="78028800" cy="78028800"/>
		</p:viewPr>
	</xsl:template>
</xsl:stylesheet>

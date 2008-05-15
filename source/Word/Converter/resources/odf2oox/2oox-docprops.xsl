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
  xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
  exclude-result-prefixes="office meta">

  <xsl:import href="common-meta.xsl"/>  

  <xsl:template name="docprops-app">
    <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
      xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
		<xsl:for-each select="document('meta.xml')/office:document-meta/office:meta">
			<!-- page count -->
			<xsl:apply-templates select="meta:document-statistic/@meta:page-count"/>
			<!-- word count -->
			<xsl:apply-templates select="meta:document-statistic/@meta:word-count"/>
			<!-- application property -->
			<xsl:call-template name="GetApplicationExtendedProperty"/>
			<!-- doc security -->
			<xsl:call-template name="GetDocSecurityExtendedProperty"/>
			<!-- paragraph count -->
			<xsl:apply-templates select="meta:document-statistic/@meta:paragraph-count"/>
			<!-- editing duration -->
			<xsl:apply-templates select="meta:editing-duration"/>
			<ScaleCrop>false</ScaleCrop>
			<LinksUpToDate>false</LinksUpToDate>
			<!-- character statistics -->
			<xsl:apply-templates select="meta:document-statistic/@meta:character-count"/>
			<!-- non white space character statistics -->
			<xsl:apply-templates select="meta:document-statistic/@meta:non-whitespace-character-count"/>
			<SharedDoc>false</SharedDoc>
			<HyperlinksChanged>false</HyperlinksChanged>
			<AppVersion>
				<xsl:value-of select="$app-version"/>
			</AppVersion>
			<xsl:for-each select="meta:user-defined[@meta:name = 'Company']">
				<Company>
					<xsl:value-of select="text()"/>
				</Company>
			</xsl:for-each>
			<xsl:for-each select="meta:user-defined[@meta:name = 'Manager']">
				<Manager>
					<xsl:value-of select="text()"/>
				</Manager>
			</xsl:for-each>
			<xsl:for-each select="meta:user-defined[@meta:name = 'HyperlinkBase']">
				<HyperlinkBase>
					<xsl:value-of select="text()"/>
				</HyperlinkBase>
			</xsl:for-each>		
		</xsl:for-each>
    </Properties>
  </xsl:template>  

</xsl:stylesheet>

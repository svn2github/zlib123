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
*     * Neither the name of the University of California, Berkeley nor the
*       names of its contributors may be used to endorse or promote products
*       derived from this software without specific prior written permission.
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
	xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/3/main"
	xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
	xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
	xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
	exclude-result-prefixes="office style svg">
	
	<xsl:template name="fonts">
		<w:fonts>
			<xsl:apply-templates select="document('content.xml')/office:document-content/office:font-face-decls" mode="fonts"/>
			<xsl:apply-templates select="document('styles.xml')/office:document-styles/office:font-face-decls" mode="fonts"/>
		</w:fonts>
	</xsl:template>

	<!-- Make sure we manage all cases -->	
	<xsl:template match="style:font-face" mode="fonts">
		<w:font>
			<xsl:choose>
				<xsl:when test="@svg:font-family = 'StarSymbol'">
					<xsl:attribute name="w:name">Symbol</xsl:attribute>
				</xsl:when>
				<xsl:otherwise>
					<!--<xsl:attribute name="w:name"><xsl:value-of select="@style:name"/></xsl:attribute>-->
					<xsl:attribute name="w:name"><xsl:value-of select="@svg:font-family"/></xsl:attribute>
				</xsl:otherwise>
			</xsl:choose>
			<xsl:if test="@style:font-family-generic">
				<w:family w:val="{@style:font-family-generic}"/>
			</xsl:if>
			<!--
			<xsl:if test="@style:font-charset">
				<w:charset w:val="{@style:font-charset}"/>
			</xsl:if>
			-->
			<xsl:if test="@style:font-pitch">
				<w:pitch w:val="{@style:font-pitch}"/>
			</xsl:if>
		</w:font>
	</xsl:template>
	
	<!-- ignored -->
	<xsl:template match="text()" mode="fonts"/>
	
</xsl:stylesheet>
	
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
		xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
                	xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
		xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
		xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/3/main"  
		exclude-result-prefixes="office text style">
	
	<xsl:key name="list-style" match="text:list-style" use="@style:name"/>
	
	<xsl:template name="numbering">
		<w:numbering>
			<xsl:apply-templates select="document('content.xml')/office:document-content/office:automatic-styles/text:list-style" mode="numbering"/>
			<xsl:apply-templates select="document('content.xml')/office:document-content/office:automatic-styles/text:list-style" mode="num"/>
		</w:numbering>
	</xsl:template>
	
	<xsl:template match="text:list-style" mode="numbering">
		<w:abstractNum w:abstractNumId="{count(preceding-sibling::text:list-style)+1}">
			<xsl:apply-templates select="text:list-level-style-number|text:list-level-style-bullet|list-level-style-image" mode="numbering"/>
		</w:abstractNum>
	</xsl:template>

	<xsl:template match="text:list-style" mode="num">
		<w:num w:numId="{count(preceding-sibling::text:list-style)+1}">
			<w:abstractNumId w:val="{count(preceding-sibling::text:list-style)+1}"/>
		</w:num>
	</xsl:template>
	
	<xsl:template match="text:list-level-style-bullet" mode="numbering">
		<xsl:if test="number(@text:level) &lt; 10">
			<w:lvl w:ilvl="{number(@text:level) - 1}">
			    	<w:numFmt w:val="bullet"/>
			    	<w:lvlText w:val="•"/>
				<w:lvlJc w:val="left"/>
				<w:pPr>
					<w:ind>
						<xsl:if test="style:list-level-properties/@text:space-before">
							<xsl:attribute name="w:left">
								<xsl:call-template name="twips-measure">
									<xsl:with-param name="length" select="style:list-level-properties/@text:space-before"/>
								</xsl:call-template>
							</xsl:attribute>
						</xsl:if>
						<xsl:if test="style:list-level-properties/@text:min-label-width">
							<xsl:attribute name="w:hanging">
								<xsl:call-template name="twips-measure">
									<xsl:with-param name="length" select="style:list-level-properties/@text:min-label-width"/>
								</xsl:call-template>
							</xsl:attribute>
						</xsl:if>
					</w:ind>
				</w:pPr>
			</w:lvl>
		</xsl:if>
	</xsl:template>
	
	<xsl:template match="text:list-level-style-number" mode="numbering">
		<xsl:if test="number(@text:level) &lt; 10">
			<w:lvl w:ilvl="{number(@text:level) - 1}">
				<xsl:choose>
					<xsl:when test="@text:start-value">
						<w:start w:val="{@text:start-value}"/>
					</xsl:when>
					<xsl:otherwise>
						<w:start w:val="1"/>
					</xsl:otherwise>
				</xsl:choose>
				<w:numFmt>
					<xsl:attribute name="w:val">
						<xsl:call-template name="num-format">
							<xsl:with-param name="format"><xsl:value-of select="@style:num-format"/></xsl:with-param>
						</xsl:call-template>
					</xsl:attribute>
				</w:numFmt>
				<w:lvlText w:val="{concat(@style:num-prefix, concat('%', @text:level), @style:num-suffix)}"/>
				<w:lvlJc w:val="left"/>
				<w:pPr>
					<w:ind>
						<xsl:if test="style:list-level-properties/@text:space-before">
							<xsl:attribute name="w:left">
								<xsl:call-template name="twips-measure">
									<xsl:with-param name="length" select="style:list-level-properties/@text:space-before"/>
								</xsl:call-template>
							</xsl:attribute>
						</xsl:if>
						<xsl:if test="style:list-level-properties/@text:min-label-width">
							<xsl:attribute name="w:hanging">
								<xsl:call-template name="twips-measure">
									<xsl:with-param name="length" select="style:list-level-properties/@text:min-label-width"/>
								</xsl:call-template>
							</xsl:attribute>
						</xsl:if>
					</w:ind>
				</w:pPr>
			</w:lvl>
		</xsl:if>
	</xsl:template>

</xsl:stylesheet>

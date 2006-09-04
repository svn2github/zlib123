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
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/3/main"  
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
	xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
    exclude-result-prefixes="style text office">

	<xsl:preserve-space elements="number:text"/>

	<xsl:template match="text:page-number" mode="paragraph">
		<w:fldSimple w:instr=" PAGE ">
			<xsl:apply-templates mode="paragraph"/>
		</w:fldSimple>
	</xsl:template>

	<xsl:template match="text:page-count" mode="paragraph">
		<w:fldSimple w:instr=" NUMPAGES ">
			<xsl:apply-templates mode="paragraph"/>
		</w:fldSimple>
	</xsl:template>

	<!-- Date Fields -->
	<xsl:template match="text:date" mode="paragraph">
		<xsl:choose>
			<xsl:when test="@text:fixed='true'">
				<w:r><w:t><xsl:value-of select="text()"/></w:t></w:r>
			</xsl:when>
			<xsl:otherwise>
				<w:fldSimple>
					<xsl:variable name="curStyle" select="@style:data-style-name"/>
					<xsl:variable name="dataStyle">
						<xsl:apply-templates select="/*/office:automatic-styles/number:date-style[@style:name=$curStyle]" mode="dataStyle"/>
					</xsl:variable>
					<xsl:attribute name="w:instr"><xsl:value-of select="concat('DATE \@ &quot;',$dataStyle,'&quot;')"/></xsl:attribute>
					<w:r>
	                    <w:rPr>
	                        <w:noProof/>
	                    </w:rPr>
	                    <w:t><xsl:value-of select="text()"/></w:t>
	                </w:r>
				</w:fldSimple>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<xsl:template match="number:text" mode="dataStyle"><xsl:value-of select="."/></xsl:template>

	<xsl:template match="number:day-of-week" mode="dataStyle">
		<xsl:choose>
			<xsl:when test="@number:style='long'">dddd</xsl:when>
			<xsl:otherwise>ddd</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<xsl:template match="number:day" mode="dataStyle">
		<xsl:choose>
			<xsl:when test="@number:style='long'">dd</xsl:when>
			<xsl:otherwise>d</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<xsl:template match="number:month" mode="dataStyle">
		<xsl:if test="@number:textual='true'">MM</xsl:if>
		<xsl:choose>
			<xsl:when test="@number:style='long'">MM</xsl:when>
			<xsl:otherwise>M</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<xsl:template match="number:year" mode="dataStyle">
		<xsl:choose>
			<xsl:when test="@number:style='long'">yyyy</xsl:when>
			<xsl:otherwise>yy</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- Time Fields -->
	<xsl:template match="text:time" mode="paragraph">
		<xsl:choose>
			<xsl:when test="@text:fixed='true'">
				<w:r><w:t><xsl:value-of select="text()"/></w:t></w:r>
			</xsl:when>
			<xsl:otherwise>
				<w:fldSimple>
					<xsl:variable name="curStyle" select="@style:data-style-name"/>
					<xsl:variable name="dataStyle">
						<xsl:apply-templates select="/*/office:automatic-styles/number:time-style[@style:name=$curStyle]" mode="dataStyle"/>
					</xsl:variable>
					<xsl:attribute name="w:instr"><xsl:value-of select="concat('TIME \@ &quot;',$dataStyle,'&quot;')"/></xsl:attribute>
					<w:r>
	                    <w:rPr>
	                        <w:noProof/>
	                    </w:rPr>
	                    <w:t><xsl:value-of select="text()"/></w:t>
	                </w:r>
				</w:fldSimple>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<xsl:template match="number:hours" mode="dataStyle">
		<xsl:choose>
			<xsl:when test="parent::node()/number:am-pm">
				<xsl:choose>
					<xsl:when test="@number:style='long'">hh</xsl:when>
					<xsl:otherwise>h</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<xsl:choose>
					<xsl:when test="@number:style='long'">HH</xsl:when>
					<xsl:otherwise>H</xsl:otherwise>
				</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<xsl:template match="number:minutes" mode="dataStyle">
		<xsl:choose>
			<xsl:when test="@number:style='long'">mm</xsl:when>
			<xsl:otherwise>m</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<xsl:template match="number:seconds" mode="dataStyle">
		<xsl:choose>
			<xsl:when test="@number:style='long'">ss</xsl:when>
			<xsl:otherwise>s</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<xsl:template match="number:am-pm" mode="dataStyle">am/pm</xsl:template>

	<!-- Author Fields -->
	<!-- TODO : comment csv file -->
	<xsl:template match="text:author-name[not(@text:fixed='true')]" mode="paragraph">
		<w:fldSimple w:instr=" USERNAME \* MERGEFORMAT ">
			<xsl:apply-templates mode="paragraph"/>
		</w:fldSimple>
	</xsl:template>

	<!-- User Fields -->
	<!-- TODO : comment csv file -->
	<xsl:template match="text:author-initials[not(@text:fixed='true')]" mode="paragraph">
		<w:fldSimple w:instr=" USERINITIALS \* Upper  \* MERGEFORMAT ">
			<xsl:apply-templates mode="paragraph"/>
		</w:fldSimple>
	</xsl:template>

	<xsl:template match="text:initial-creator[not(@text:fixed='true')]" mode="paragraph">
		<w:fldSimple w:instr=" AUTHOR ">
			<xsl:apply-templates mode="paragraph"/>
		</w:fldSimple>
	</xsl:template>
	
	<xsl:template match="text:subject[not(@text:fixed='true')]" mode="paragraph">
		<w:fldSimple w:instr=" SUBJECT ">
			<xsl:apply-templates mode="paragraph"/>
		</w:fldSimple>
	</xsl:template>
	
	<xsl:template match="text:title[not(@text:fixed='true')]" mode="paragraph">
		<w:fldSimple w:instr=" TITLE ">
			<xsl:apply-templates mode="paragraph"/>
		</w:fldSimple>
	</xsl:template>
</xsl:stylesheet>


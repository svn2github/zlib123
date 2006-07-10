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
                xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
				xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/3/main"  
				exclude-result-prefixes="style">

	<!--
	function    : num-format
	param       : format (string)
	description : convert the ODF numFormat to OOX numFormat
	-->
	<xsl:template name="num-format">
		<xsl:param name="format"/>
		<xsl:choose>
			<xsl:when test="$format='1'">decimal</xsl:when>
			<xsl:when test="$format='i'">lowerRoman</xsl:when>
			<xsl:when test="$format='I'">upperRoman</xsl:when>
			<xsl:when test="$format='a'">lowerLetter</xsl:when>
			<xsl:when test="$format='A'">upperLetter</xsl:when>
			<xsl:otherwise>decimal</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	
	<!--
		U n i t s
		
		1 pt = 20 twip
		1 in = 72 pt = 1440 twip
		1 cm = 1440 / 2.54 twip
		1 pica = 12 pt
		1 dpt (didot point) = 1/72 in (almost the same as 1 pt)
		1 px = 0.0264cm at 96dpi (Windows default)
	-->
	
	
	<!-- 
		Convert various length units to twips (twentieths of a point)
	-->
	<xsl:template name="twips-measure">
		<xsl:param name="length"/>
		<xsl:choose>
			<xsl:when test="contains($length, 'cm')">
				<xsl:value-of select="round(number(substring-before($length, 'cm')) * 1440 div 2.54)"/>
			</xsl:when>
			<xsl:when test="contains($length, 'in')">
				<xsl:value-of select="round(number(substring-before($length, 'in')) * 1440)"/>
			</xsl:when>
			<xsl:when test="contains($length, 'pt')">
				<xsl:value-of select="round(number(substring-before($length, 'pt')) * 20)"/>
			</xsl:when>
			<xsl:when test="contains($length, 'twip')">
				<xsl:value-of select="substring-before($length, 'twip')"/>
			</xsl:when>
			<xsl:when test="contains($length, 'pica')">
				<xsl:value-of select="round(number(substring-before($length, 'pica')) * 240)"/>
			</xsl:when>
			<xsl:when test="contains($length, 'dpt')">
				<xsl:value-of select="round(number(substring-before($length, 'dpt')) * 20)"/>
			</xsl:when>
			<xsl:when test="contains($length, 'px')">
				<xsl:value-of select="round(number(substring-before($length, 'px')) * 1440 div 96.19)"/>
			</xsl:when>
			<xsl:when test="not($length)">0</xsl:when>
			<xsl:otherwise><xsl:value-of select="$length"/></xsl:otherwise>
		</xsl:choose>	
	</xsl:template> 
	
	<!-- 
		Convert to emu
		1cm = 360000 emu
	-->
	<xsl:template name="emu-measure">
		<xsl:param name="length"/>
		<xsl:choose>
			<xsl:when test="contains($length, 'cm')">
				<xsl:value-of select="round(number(substring-before($length, 'cm')) * 360000)"/>
			</xsl:when>
			<xsl:when test="contains($length, 'in')">
				<xsl:value-of select="round(number(substring-before($length, 'in')) * 360000 * 2.54)"/>
			</xsl:when>
			<xsl:when test="contains($length, 'pt')">
				<xsl:value-of select="round(number(substring-before($length, 'pt')) * 360000 * 2.54 * 72)"/>
			</xsl:when>
			<xsl:when test="contains($length, 'px')">
				<xsl:value-of select="round(number(substring-before($length, 'px')) * 360000 div 37.87)"/>
			</xsl:when>
			<xsl:otherwise><xsl:value-of select="$length"/></xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	
  
</xsl:stylesheet>

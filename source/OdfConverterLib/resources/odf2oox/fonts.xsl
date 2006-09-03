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
<xsl:stylesheet version="1.0"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/3/main"
	xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
	xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
	xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
	exclude-result-prefixes="office style svg">

  <xsl:key name="fonts"
		match="office:font-face-decls/style:font-face"
		use="@style:name"/>
	
	<xsl:template name="fonts">
		<w:fonts>
			<!-- We suppose that the fonts declared in content.xml and styles.xml are the same. -->
			<xsl:apply-templates select="document('content.xml')/office:document-content/office:font-face-decls/style:font-face" mode="fonts"/>
		</w:fonts>
	</xsl:template>

	<!-- Make sure we manage all cases -->	
	<xsl:template match="style:font-face" mode="fonts">
    <!-- We do not take into consideration fonts that have a name that does not match the family name. -->
    <xsl:if test='not(@svg:font-family) or (@style:name=@svg:font-family) or (concat("&apos;",@style:name,"&apos;")=@svg:font-family)'>
      <w:font>
        <!-- Make sur the 'x-symbol' charset is always '02' and the asian and complex charset are not control -->
        <xsl:choose>
          <xsl:when test="@svg:font-family = 'StarSymbol'">
            <xsl:attribute name="w:name">Symbol</xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="w:name">
              <xsl:value-of select="@style:name"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
        <xsl:choose>
          <xsl:when test="@style:font-charset = 'x-symbol'">
            <w:charset w:val="02"/>
          </xsl:when>
          <xsl:otherwise>
            <w:charset w:val="00"/>
          </xsl:otherwise>
        </xsl:choose>
        <xsl:choose>

          <!-- open xml don't know the attribute filed system : replace with auto -->
          <xsl:when test="@style:font-family-generic='system'">
            <w:family w:val="auto"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:if test="@style:font-family-generic">
              <w:family w:val="{@style:font-family-generic}"/>
            </xsl:if>
          </xsl:otherwise>
        </xsl:choose>
        <xsl:if test="@style:font-pitch">
          <w:pitch w:val="{@style:font-pitch}"/>
        </xsl:if>
      </w:font>
    </xsl:if>
  </xsl:template>

  <!-- Map font types -->
  <xsl:template name="ComputeFontName">
    <xsl:param name="fontName"/>
    <xsl:choose>
      <xsl:when test="$fontName = 'StarSymbol'">Symbol</xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <!--xsl:when test="document('styles.xml')//style:font-face[@style:name=$fontName]/@svg:font-family"-->
          <xsl:when test="key('fonts',$fontName)/@svg:font-family">
              <xsl:value-of select='translate(key("fonts",$fontName)[1]/@svg:font-family,"&apos;","")'/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$fontName"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
	
	<!-- ignored -->
	<xsl:template match="text()" mode="fonts"/>
	
</xsl:stylesheet>
	

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
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/3/main"
  exclude-result-prefixes="w">

  <xsl:import href="tables.xsl"/>
  
  <xsl:template name="content">
    <office:document-content>
      <office:scripts/>
      <office:font-face-decls/>
      <office:automatic-styles/>
      <office:body>
        <office:text>
          <xsl:apply-templates select="document('word/document.xml')/w:document/w:body"/>
        </office:text>
      </office:body>
    </office:document-content>
  </xsl:template>
  
  <xsl:template name="w:p">
    <xsl:apply-templates/>
  </xsl:template>
  
  <xsl:template match="w:r">
    <xsl:apply-templates/>
  </xsl:template>  
  
  <xsl:template match="w:t">
    
    <xsl:message terminate="no">progress:w:t</xsl:message>
    <xsl:variable name="outline">
      <xsl:value-of select="ancestor::w:p/w:pPr/w:pStyle/@w:val"/>
    </xsl:variable>
    <xsl:choose>
      <xsl:when  test="document('word/styles.xml')/w:styles/w:style[@w:styleId=$outline]/w:pPr/w:outlineLvl/@w:val">
        <text:h>
          <xsl:attribute name="text:outline-level">
            <xsl:value-of select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$outline]/w:pPr/w:outlineLvl/@w:val+1"/>
          </xsl:attribute>
          <xsl:attribute name="text:style-name">
            <xsl:value-of select="."/>
          </xsl:attribute>
          <xsl:value-of select="."/>
        </text:h>
      </xsl:when>
      <xsl:when test="ancestor::w:p/w:pPr/w:rPr/w:rStyle/@w:val">
        <text:p>
          <xsl:attribute name="text:style-name">
            <xsl:value-of select="ancestor::w:p/w:pPr/w:rPr/w:rStyle/@w:val"/>
          </xsl:attribute>
          <xsl:value-of select="."/>
        </text:p>
      </xsl:when>
      <xsl:otherwise>
        <text:p text:style-name="Standard">
          <xsl:value-of select="."/>
        </text:p>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>

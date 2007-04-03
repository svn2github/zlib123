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
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  exclude-result-prefixes="e">
  
  <!-- insert  number format style -->
  
  <xsl:template match="e:numFmt" mode="automaticstyles">
    <xsl:choose>
      
      <!-- when there are different formats for positive and negative numbers -->
      <xsl:when test="contains(@formatCode,';')">
        <number:number-style style:name="{concat(generate-id(.),'P0')}">
          <xsl:call-template name="InsertNumberFormatting">
            <xsl:with-param name="formatCode">
              <xsl:value-of select="substring-before(@formatCode,';')"/>
            </xsl:with-param>
          </xsl:call-template>
        </number:number-style>
        <number:number-style style:name="{generate-id(.)}">
          <xsl:call-template name="InsertNumberFormatting">
            <xsl:with-param name="formatCode">
              <xsl:value-of select="substring-after(@formatCode,';')"/>
            </xsl:with-param>
          </xsl:call-template>
          <style:map style:condition="value()&gt;=0" style:apply-style-name="{concat(generate-id(.),'P0')}"/>
        </number:number-style>
      </xsl:when>
      
      <xsl:otherwise>
        <number:number-style style:name="{generate-id(.)}">
          <xsl:call-template name="InsertNumberFormatting">
            <xsl:with-param name="formatCode">
              <xsl:value-of select="@formatCode"/>
            </xsl:with-param>
          </xsl:call-template>
        </number:number-style>
      </xsl:otherwise>
      
    </xsl:choose>
  </xsl:template>
  
  <!-- template to create number format -->
  
  <xsl:template name="InsertNumberFormatting">
    <xsl:param name="formatCode"/>
    
    <!-- handle red negative numbers -->
    <xsl:if test="contains($formatCode,'Red')">
      <style:text-properties fo:color="#ff0000"/>
    </xsl:if>
    
    <!-- add '-' at the beginning -->
    <xsl:if test="contains($formatCode,'-')">
      <number:text>-</number:text>
    </xsl:if>
    
    <number:number>
      
      <xsl:variable name="formatCodeWithoutComma">
        <xsl:choose>
          <xsl:when test="contains($formatCode,'.')">
            <xsl:value-of select="substring-before($formatCode,'.')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$formatCode"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      
      <!-- decimal places -->
      <xsl:attribute name="number:decimal-places">
        <xsl:choose>
          <xsl:when test="contains($formatCode,'.')">
            <xsl:call-template name="InsertDecimalPlaces">
              <xsl:with-param name="code">
                <xsl:value-of select="substring-after($formatCode,'.')"/>
              </xsl:with-param>
              <xsl:with-param name="value">0</xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      
      <!-- min integer digits -->
      
      <xsl:if test="contains($formatCodeWithoutComma,'0')">
        <xsl:attribute name="number:min-integer-digits">
          <xsl:call-template name="InsertMinIntegerDigits">
            <xsl:with-param name="code">
              <xsl:value-of select="substring-before($formatCodeWithoutComma,'0')"/>
            </xsl:with-param>
            <xsl:with-param name="value">1</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      
      <!-- grouping -->
      <xsl:if test="contains($formatCode,',')">
        <xsl:attribute name="number:grouping">true</xsl:attribute>
      </xsl:if>
      
    </number:number>
  </xsl:template>
  
  <!-- template which inserts min integer digits -->
  
  <xsl:template name="InsertMinIntegerDigits">
    <xsl:param name="code"/>
    <xsl:param name="value"/>
    <xsl:choose>
      <xsl:when test="contains($code,'0')">
        <xsl:call-template name="InsertMinIntegerDigits">
          <xsl:with-param name="code">
            <xsl:value-of select="substring($code,0,string-length($code)-1)"/>
          </xsl:with-param>
          <xsl:with-param name="value">
            <xsl:value-of select="$value+1"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$value"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- template which inserts decimal places -->
  
  <xsl:template name="InsertDecimalPlaces">
    <xsl:param name="code"/>
    <xsl:param name="value"/>
    <xsl:choose>
      <xsl:when test="contains($code,'0') or contains($code,'#')">
        <xsl:call-template name="InsertDecimalPlaces">
          <xsl:with-param name="code">
            <xsl:value-of select="substring($code,2)"/>
          </xsl:with-param>
          <xsl:with-param name="value">
            <xsl:value-of select="$value+1"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$value"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
</xsl:stylesheet>

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
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  exclude-result-prefixes="number style fo">
  
  <!-- template to insert date formats -->
  
  <xsl:template match="number:date-style" mode="numFormat">
    <xsl:param name="numId"/>
    <numFmt numFmtId="{$numId}">
      <xsl:attribute name="formatCode">
        <xsl:apply-templates mode="date"/>
      </xsl:attribute>
    </numFmt>
    <xsl:choose>
      <xsl:when test="following-sibling::number:date-style">
        <xsl:apply-templates select="following-sibling::number:date-style[1]" mode="numFormat">
          <xsl:with-param name="numId">
            <xsl:value-of select="$numId + 1"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <!-- insert year format -->
  
  <xsl:template match="number:year" mode="date">
    <xsl:choose>
      <xsl:when test="@number:style='long'">yyyy</xsl:when>
      <xsl:otherwise>yy</xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- insert month format -->
  
  <xsl:template match="number:month" mode="date">
    <xsl:choose>
      <xsl:when test="@number:style='long' and @number:textual='true'">mmmm</xsl:when>
      <xsl:when test="@number:textual='true'">mmm</xsl:when>
      <xsl:when test="@number:style='long'">mm</xsl:when>
      <xsl:otherwise>m</xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- insert day of week format -->
  
  <xsl:template match="number:day-of-week" mode="date">
    <xsl:choose>
    <xsl:when test="@number:style='long'">dddd</xsl:when>
    <xsl:otherwise>ddd</xsl:otherwise>
  </xsl:choose></xsl:template>
  
  <!-- insert day format -->
  
  <xsl:template match="number:day" mode="date">
    <xsl:choose>
      <xsl:when test="@number:style='long'">dd</xsl:when>
      <xsl:otherwise>d</xsl:otherwise>
    </xsl:choose>
  </xsl:template> 
  
  <!-- insert hour format -->
  
  <xsl:template match="number:hours" mode="date">
    <xsl:choose>
      <xsl:when test="@number:style='long'">hh</xsl:when>
      <xsl:otherwise>h</xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- insert minute format -->
  
  <xsl:template match="number:minutes" mode="date">
    <xsl:choose>
      <xsl:when test="@number:style='long'">mm</xsl:when>
      <xsl:otherwise>m</xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- add am or pm -->
  
  <xsl:template match="number:am-pm" mode="date">AM/PM</xsl:template>
  
  <!-- insert second format -->
  
  <xsl:template match="number:seconds" mode="date">
    <xsl:choose>
      <xsl:when test="@number:style='long'">ss</xsl:when>
      <xsl:otherwise>s</xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- add all formatting characters -->
  
  <xsl:template match="number:text" mode="date">
    <xsl:value-of select="."/>
  </xsl:template>
  
</xsl:stylesheet>
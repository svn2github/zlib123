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
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/3/main"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  exclude-result-prefixes="w">



  <!--
		U n i t s
		
		1 pt = 20 twip
		1 in = 72 pt = 1440 twip
		1 cm = 1440 / 2.54 twip
		1 pica = 12 pt
		1 dpt (didot point) = 1/72 in (almost the same as 1 pt)
		1 px = 0.0264cm at 96dpi (Windows default)
		1 milimeter(mm) = 0.1cm
               1cm = 360000 emu
  -->

  <!-- Convert a measure in twips to a 'unit' measure -->
  <xsl:template name="ConvertTwips">
    <xsl:param name="length"/>
    <xsl:param name="unit"/>
    <xsl:choose>
      <xsl:when test="$unit = 'cm'">
        <xsl:value-of select="concat(format-number($length * 2.54 div 1440,'#.###'),'cm')"/>
      </xsl:when>
      <xsl:when test="$unit = 'mm'">
        <xsl:value-of select="concat(format-number($length * 25.4 div 1440,'#.###'),'mm')"/>
      </xsl:when>
      <xsl:when test="$unit= 'in'">
        <xsl:value-of select="concat(format-number($length div 1440,'#.###'),'in')"/>
      </xsl:when>
      <xsl:when test="$unit = 'pt'">
        <xsl:value-of select="concat(format-number($length div 20,'#.###'),'pt')"/>
      </xsl:when>
      <xsl:when test="$unit = 'twip'">
        <xsl:value-of select="concat($length,'twip')"/>
      </xsl:when>
      <xsl:when test="$unit = 'pica'">
        <xsl:value-of select="concat(format-number($length div 240,'#.###'),'pica')"/>
      </xsl:when>
      <xsl:when test="$unit = 'dpt'">
        <xsl:value-of select="concat(format-number($length div 20,'#.###'),'dpt')"/>
      </xsl:when>
      <xsl:when test="$unit = 'px'">
        <xsl:value-of select="concat(format-number($length * 96.19 div 1440,'#.###'),'px')"/>
      </xsl:when>
      <xsl:when test="not($length)">
        <xsl:value-of select="concat(0,'cm')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$length"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Convert a measure in points to a 'unit' measure -->
  <xsl:template name="ConvertPoints">
    <xsl:param name="length"/>
    <xsl:param name="unit"/>
    <xsl:choose>
      <xsl:when test="$unit = 'cm'">
        <xsl:value-of select="concat(format-number($length * 2.54 div 72,'#.###'),'cm')"/>
      </xsl:when>
      <xsl:when test="$unit = 'mm'">
        <xsl:value-of select="concat(format-number($length * 25.4 div 72,'#.###'),'mm')"/>
      </xsl:when>
      <xsl:when test="$unit = 'in'">
        <xsl:value-of select="concat(format-number($length div 72,'#.###'),'in')"/>
      </xsl:when>
      <xsl:when test="$unit = 'pt'">
        <xsl:value-of select="concat($length,'pt')"/>
      </xsl:when>
      <xsl:when test="$unit = 'pica'">
        <xsl:value-of select="concat(format-number($length div 12,'#.###'),'pica')"/>
      </xsl:when>
      <xsl:when test="$unit = 'dpt'">
        <xsl:value-of select="concat($length,'dpt')"/>
      </xsl:when>
      <xsl:when test="$unit = 'px'">
        <xsl:value-of select="concat(format-number($length * 96.19 div 72,'#.###'),'px')"/>
      </xsl:when>
      <xsl:when test="not($length)">
        <xsl:value-of select="concat(0,'cm')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$length"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Convert a measure in half points to a 'unit' measure -->
  <xsl:template name="ConvertHalfPoints">
    <xsl:param name="length"/>
    <xsl:param name="unit"/>
    <xsl:choose>
      <xsl:when test="$unit = 'cm'">
        <xsl:value-of select="concat(format-number($length * 2.54 div 144,'#.###'),'cm')"/>
      </xsl:when>
      <xsl:when test="$unit = 'mm'">
        <xsl:value-of select="concat(format-number($length * 25.4 div 144,'#.###'),'mm')"/>
      </xsl:when>
      <xsl:when test="$unit = 'in'">
        <xsl:value-of select="concat(format-number($length div 144,'#.###'),'in')"/>
      </xsl:when>
      <xsl:when test="$unit = 'pt'">
        <xsl:value-of select="concat($length div 2,'pt')"/>
      </xsl:when>
      <xsl:when test="$unit = 'pica'">
        <xsl:value-of select="concat(format-number($length div 144,'#.###'),'pica')"/>
      </xsl:when>
      <xsl:when test="$unit = 'dpt'">
        <xsl:value-of select="concat($length div 2,'dpt')"/>
      </xsl:when>
      <xsl:when test="$unit = 'px'">
        <xsl:value-of select="concat(format-number($length * 96.19 div 144,'#.###'),'px')"/>
      </xsl:when>
      <xsl:when test="not($length)">
        <xsl:value-of select="concat(0,'cm')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$length"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Convert a measure in eigths of a point to a 'unit' measure -->
  <xsl:template name="ConvertEighthsPoints">
    <xsl:param name="length"/>
    <xsl:param name="unit"/>
    <xsl:choose>
      <xsl:when test="$unit = 'cm'">
        <xsl:value-of select="concat(format-number($length * 2.54 div 576,'#.###'),'cm')"/>
      </xsl:when>
      <xsl:when test="$unit = 'mm'">
        <xsl:value-of select="concat(format-number($length * 25.4 div 576,'#.###'),'mm')"/>
      </xsl:when>
      <xsl:when test="$unit = 'in'">
        <xsl:value-of select="concat(format-number($length div 576,'#.###'),'in')"/>
      </xsl:when>
      <xsl:when test="$unit = 'pt'">
        <xsl:value-of select="concat(format-number($length div 8,'#.###'),'pt')"/>
      </xsl:when>
      <xsl:when test="$unit = 'pica'">
        <xsl:value-of select="concat(format-number($length div 96,'#.###'),'pica')"/>
      </xsl:when>
      <xsl:when test="$unit = 'dpt'">
        <xsl:value-of select="concat(format-number($length div 8,'#.###'),'dpt')"/>
      </xsl:when>
      <xsl:when test="$unit = 'px'">
        <xsl:value-of select="concat(format-number($length * 96.19 div 576,'#.###'),'px')"/>
      </xsl:when>
      <xsl:when test="not($length)">
        <xsl:value-of select="concat(0,'cm')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$length"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="emu-measure">
    <xsl:param name="length"/>
    <xsl:param name="unit"/>
    <xsl:choose>
      <xsl:when test="$unit = 'cm'">
        <xsl:value-of select="concat(format-number($length div 360000, '#.##'), 'cm')"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

</xsl:stylesheet>

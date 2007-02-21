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
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main" exclude-result-prefixes="e r">
  
  
  <!-- gets a column number from cell coordinates -->
  <xsl:template name="GetColNum">
    <xsl:param name="cell"/>
    <xsl:param name="columnId"/>
    
    <xsl:choose>
      <!-- when whole literal column id has been extracted than convert alphabetic index to number -->
      <xsl:when test="number(substring($cell,1,1))">
        <xsl:call-template name="GetAlphabeticPosition">
          <xsl:with-param name="literal" select="$columnId"/>
        </xsl:call-template>
      </xsl:when>
      <!--  recursively extract literal column id (i.e if $cell='GH15' it will return 'GH') -->
      <xsl:otherwise>
        <xsl:call-template name="GetColNum">
          <xsl:with-param name="cell" select="substring-after($cell,substring($cell,1,1))"/>
          <xsl:with-param name="columnId" select="concat($columnId,substring($cell,1,1))"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- translates literal index to number -->
  <xsl:template name="GetAlphabeticPosition">
    <xsl:param name="literal"/>
    <xsl:param name="number" select="0"/>
    <xsl:param name="level" select="0"/>
    
    <xsl:variable name="lastCharacter">
      <xsl:value-of select="substring($literal,string-length($literal))"/>
    </xsl:variable>
    
    <xsl:variable name="lastCharacterPosition">
      <xsl:call-template name="CharacterToPosition">
        <xsl:with-param name="character" select="$lastCharacter"/>
      </xsl:call-template>
    </xsl:variable>
    
    <xsl:variable name="power">
      <xsl:call-template name="Power">
        <xsl:with-param name="base" select="26"/>
        <xsl:with-param name="exponent" select="$level"/>
      </xsl:call-template>
    </xsl:variable>
    
    <xsl:choose>
      <xsl:when test="string-length($literal)>1">
        <xsl:call-template name="GetAlphabeticPosition">
          <xsl:with-param name="literal" select="substring-before($literal,$lastCharacter)"/>
          <xsl:with-param name="level" select="$level+1"/>
          <xsl:with-param name="number">
            <xsl:value-of select="$lastCharacterPosition*$power + $number"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$lastCharacterPosition*$power + $number"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- returns position in alphabet of a single character-->
  <xsl:template name="CharacterToPosition">
    <xsl:param name="character"/>
    
    <xsl:choose>
      <xsl:when test="$character='A'">1</xsl:when>
      <xsl:when test="$character='B'">2</xsl:when>
      <xsl:when test="$character='C'">3</xsl:when>
      <xsl:when test="$character='D'">4</xsl:when>
      <xsl:when test="$character='E'">5</xsl:when>
      <xsl:when test="$character='F'">6</xsl:when>
      <xsl:when test="$character='G'">7</xsl:when>
      <xsl:when test="$character='H'">8</xsl:when>
      <xsl:when test="$character='I'">9</xsl:when>
      <xsl:when test="$character='J'">10</xsl:when>
      <xsl:when test="$character='K'">11</xsl:when>
      <xsl:when test="$character='L'">12</xsl:when>
      <xsl:when test="$character='M'">13</xsl:when>
      <xsl:when test="$character='N'">14</xsl:when>
      <xsl:when test="$character='O'">15</xsl:when>
      <xsl:when test="$character='P'">16</xsl:when>
      <xsl:when test="$character='Q'">17</xsl:when>
      <xsl:when test="$character='R'">18</xsl:when>
      <xsl:when test="$character='S'">19</xsl:when>
      <xsl:when test="$character='T'">20</xsl:when>
      <xsl:when test="$character='U'">21</xsl:when>
      <xsl:when test="$character='V'">22</xsl:when>
      <xsl:when test="$character='W'">23</xsl:when>
      <xsl:when test="$character='X'">24</xsl:when>
      <xsl:when test="$character='Y'">25</xsl:when>
      <xsl:when test="$character='Z'">26</xsl:when>
    </xsl:choose>
  </xsl:template>
  
  
  </xsl:stylesheet>
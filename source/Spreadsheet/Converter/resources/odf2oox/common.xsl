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
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">
  
  
  <!-- template to convert cell type -->
  <xsl:template name="ConvertTypes">
    <xsl:param name="type"/>
    <xsl:choose>
      <xsl:when test="$type = 'boolean'">b</xsl:when>
      <xsl:when test="$type = 'currency' or $type='float' or $type='percentage'">n</xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <!-- converting number of column to char name of column -->
  <xsl:template name="NumbersToChars">
    <xsl:param name="num"/>
    <xsl:choose>
      <xsl:when test="$num = 0">A</xsl:when>
      <xsl:when test="$num = 1">B</xsl:when>
      <xsl:when test="$num = 2">C</xsl:when>
      <xsl:when test="$num = 3">D</xsl:when>
      <xsl:when test="$num = 4">E</xsl:when>
      <xsl:when test="$num = 5">F</xsl:when>
      <xsl:when test="$num = 6">G</xsl:when>
      <xsl:when test="$num = 7">H</xsl:when>
      <xsl:when test="$num = 8">I</xsl:when>
      <xsl:when test="$num = 9">J</xsl:when>
      <xsl:when test="$num = 10">K</xsl:when>
      <xsl:when test="$num = 11">L</xsl:when>
      <xsl:when test="$num = 12">M</xsl:when>
      <xsl:when test="$num = 13">N</xsl:when>
      <xsl:when test="$num = 14">O</xsl:when>
      <xsl:when test="$num = 15">P</xsl:when>
      <xsl:when test="$num = 16">Q</xsl:when>
      <xsl:when test="$num = 17">R</xsl:when>
      <xsl:when test="$num = 18">S</xsl:when>
      <xsl:when test="$num = 19">T</xsl:when>
      <xsl:when test="$num = 20">U</xsl:when>
      <xsl:when test="$num = 21">V</xsl:when>
      <xsl:when test="$num = 22">W</xsl:when>
      <xsl:when test="$num = 23">X</xsl:when>
      <xsl:when test="$num = 24">Y</xsl:when>
      <xsl:when test="$num = 25">Z</xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="NumbersToChars">
          <xsl:with-param name="num">
            <xsl:value-of select="floor($num div 26)-1"/>
          </xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="NumbersToChars">
          <xsl:with-param name="num">
            <xsl:value-of select="$num - 26*floor($num div 26)"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
</xsl:stylesheet>

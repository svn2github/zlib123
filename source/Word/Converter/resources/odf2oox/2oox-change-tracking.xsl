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
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:pct="urn:cleverage:xmlns:post-processings:change-tracking" 
  exclude-result-prefixes="office text dc">

  <xsl:key name="changed-regions" match="text:tracked-changes/text:changed-region" use="@text:id"/>

  <!-- change starts -->
  
  <xsl:template match="text:change-start">
    <xsl:call-template name="InsertChangeStart"/>
  </xsl:template>

  <xsl:template match="text:change-start" mode="paragraph">
    <xsl:call-template name="InsertChangeStart"/>
  </xsl:template>

  <!-- Insert a change start item -->
  <xsl:template name="InsertChangeStart">
    <xsl:variable name="changed-region" select="key('changed-regions', @text:change-id)"/>
    <xsl:choose>
      <xsl:when test="$changed-region/text:insertion">
        <xsl:call-template name="InsertStartInsertion">
          <xsl:with-param name="changed-region" select="$changed-region"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$changed-region/text:format-change">
        <xsl:call-template name="InsertStartFormatChange">
          <xsl:with-param name="changed-region" select="$changed-region"/>
        </xsl:call-template>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <!-- change ends -->

  <xsl:template match="text:change-end">
    <xsl:call-template name="InsertChangeEnd"/>
  </xsl:template>

  <xsl:template match="text:change-end" mode="paragraph">
    <xsl:call-template name="InsertChangeEnd"/>
  </xsl:template>

  <!-- Insert a change end item -->
  <xsl:template name="InsertChangeEnd">
    <xsl:variable name="changed-region" select="key('changed-regions', @text:change-id)"/>
    <xsl:choose>
      <xsl:when test="$changed-region/text:insertion">
        <xsl:call-template name="InsertEndInsertion"/>
      </xsl:when>
      <xsl:when test="$changed-region/text:format-change">
        <xsl:call-template name="InsertEndFormatChange"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- changes -->

  <xsl:template match="text:change">
    <xsl:call-template name="InsertChange"/>
  </xsl:template>

  <xsl:template match="text:change" mode="paragraph">
    <xsl:call-template name="InsertChange"/>
  </xsl:template>

  <!-- Insert a change item -->
  <xsl:template name="InsertChange">
    <xsl:variable name="changed-region" select="key('changed-regions', @text:change-id)"/>
    <xsl:choose>
      <xsl:when test="$changed-region/text:deletion">
        <xsl:call-template name="InsertDeletion">
          <xsl:with-param name="changed-region" select="$changed-region"/>
        </xsl:call-template>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- insertions -->
  
  <!-- Insert start insertion item -->
  <xsl:template name="InsertStartInsertion">
    <xsl:param name="changed-region"/>
    <pct:start-insert pct:id="{@text:change-id}" pct:creator="{$changed-region/text:insertion/office:change-info/dc:creator}" pct:date="{$changed-region/text:insertion/office:change-info/dc:date}"/>
  </xsl:template>

  <!-- Insert end insertion item -->
  <xsl:template name="InsertEndInsertion">
    <pct:end-insert pct:id="{@text:change-id}"/>
  </xsl:template>
  
  <!-- deletions -->

  <!-- Insert a deletion -->
  <xsl:template name="InsertDeletion">
    <xsl:param name="changed-region"/>
    <pct:deletion pct:creator="{$changed-region/text:deletion/office:change-info/dc:creator}" pct:date="{$changed-region/text:deletion/office:change-info/dc:date}">
      <xsl:apply-templates select="$changed-region/text:deletion"/>
    </pct:deletion>
  </xsl:template>
  
  <!-- format changes -->

  <!-- Insert start format change item -->
  <xsl:template name="InsertStartFormatChange">
    <xsl:param name="changed-region"/>
    <pct:start-format-change pct:id="{@text:change-id}" pct:creator="{$changed-region/text:format-change/office:change-info/dc:creator}" pct:date="{$changed-region/text:format-change/office:change-info/dc:date}"/>
  </xsl:template>

  <!-- Insert end format change item -->
  <xsl:template name="InsertEndFormatChange">
    <pct:end-format-change pct:id="{@text:change-id}"/>
  </xsl:template>
  
  
</xsl:stylesheet>

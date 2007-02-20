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
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  exclude-result-prefixes="e a pic r w">

  <xsl:template name="GetTarget">
    <xsl:param name="document"/>
    <xsl:param name="file" select="$document"/>
    <xsl:param name="id"/>

    <xsl:choose>
      <xsl:when test="contains($file,'/')">
        <xsl:call-template name="GetTarget">
          <xsl:with-param name="file">
            <xsl:value-of select="substring-after($file,'/')"/>
          </xsl:with-param>
          <xsl:with-param name="id" select="$id"/>
          <xsl:with-param name="document" select="$document"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="document(concat(substring-before($document,$file),'_rels/',$file,'.rels'))">
          <xsl:for-each
            select="document(concat(substring-before($document,$file),'_rels/',$file,'.rels'))//node()[name()='Relationship']">
            <xsl:if test="./@Id=$id">
              <xsl:value-of select="./@Target"/>
              </xsl:if>
          </xsl:for-each>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>

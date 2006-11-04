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
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  exclude-result-prefixes="w r office style">

  <!-- content types -->
  <xsl:template name="contentTypes">
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="jpeg" ContentType="image/jpeg"/>
      <Default Extension="jpg" ContentType="image/jpeg"/>
      <Default Extension="gif" ContentType="image/gif"/>
      <Default Extension="png" ContentType="image/png"/>
      <Default Extension="wmf" ContentType="image/x-emf"/>
      <Default Extension="emf" ContentType="image/x-emf"/>
      <Default Extension="tif" ContentType="image/tiff"/>
      <Default Extension="tiff" ContentType="image/tiff"/>
      <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
      <Default Extension="rels"
        ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml"
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/numbering.xml"
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
      <Override PartName="/docProps/core.xml"
        ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
      <Override PartName="/docProps/app.xml"
        ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
      <xsl:if test="$docprops-custom-file > 0">
        <Override PartName="/docProps/custom.xml"
          ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
      </xsl:if>
      <Override PartName="/word/styles.xml"
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
      <Override PartName="/word/fontTable.xml"
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
      <Override PartName="/word/settings.xml"
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
      <Override PartName="/word/footnotes.xml"
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
      <Override PartName="/word/endnotes.xml"
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
      <Override PartName="/word/comments.xml"
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>

     <xsl:call-template name="InsertHeaderFooterContentTypes"/>

    </Types>
  </xsl:template>
  
  <!-- Headers / Footers content types -->
  <xsl:template name="InsertHeaderFooterContentTypes">
    <xsl:variable name="masterPages"
      select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page"/>
    
    <xsl:for-each select="$masterPages/style:header | $masterPages/style:header-left">
      <Override xmlns="http://schemas.openxmlformats.org/package/2006/content-types"
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml">
        <xsl:attribute name="PartName">/word/header<xsl:value-of select="position()"/>.xml</xsl:attribute>
      </Override>
    </xsl:for-each>
    <xsl:for-each select="$masterPages/style:footer | $masterPages/style:footer-left">
      <Override xmlns="http://schemas.openxmlformats.org/package/2006/content-types"
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml">
        <xsl:attribute name="PartName">/word/footer<xsl:value-of select="position()"/>.xml</xsl:attribute>
      </Override>
    </xsl:for-each>
  </xsl:template>

</xsl:stylesheet>
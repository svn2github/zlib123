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
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main" exclude-result-prefixes="e r">

  
  <xsl:template name="InsertOLEObjects">
    <xsl:if test="e:oleObject[not(@r:id)]">
    <table:shapes>
        <xsl:for-each select="e:oleObject ">
            <xsl:call-template name="InsertOLEObjectsLinks"/>    
        </xsl:for-each>
    </table:shapes>  
    </xsl:if>
  </xsl:template>
  
  <xsl:template name="InsertOLEObjectsLinks">
    
    <xsl:variable name="ProgId">
      <xsl:value-of select="@progId"/>
    </xsl:variable>
    
    <xsl:variable name="NumberExternalLink">
      <xsl:value-of select="@link"/>
    </xsl:variable>
    
    <xsl:variable name="ExternalLink">
      <xsl:value-of select="concat('xl/externalLinks/externalLink', substring-after(substring-before($NumberExternalLink, ']'), '['), '.xml')"/>
    </xsl:variable>
    
    <xsl:variable name="ExternalLinkRels">
      <xsl:value-of select="concat('xl/externalLinks/_rels/externalLink', substring-after(substring-before($NumberExternalLink, ']'), '['), '.xml', '.rels')"/>
    </xsl:variable>
    
    <xsl:variable name="ExternalLinkId">
      <xsl:for-each select="document($ExternalLink)">
        <xsl:value-of select="e:externalLink/e:oleLink[@progId = $ProgId]/@r:id"/>
      </xsl:for-each>
    </xsl:variable>
    
    <xsl:variable name="XlinkOLEObject">
    
      <xsl:for-each
        select="document($ExternalLinkRels)//node()[name()='Relationship']">
        <xsl:if test="./@Id = $ExternalLinkId">
          <xsl:value-of select="translate(substring-after(./@Target, 'file://'), '\', '/')"/>
        </xsl:if>
      </xsl:for-each>
    </xsl:variable>
    
    <xsl:if test="$XlinkOLEObject != ''">
    
    <draw:frame draw:z-index="0" svg:width="2.137cm" svg:height="0.452cm"
      svg:x="2.258cm" svg:y="2.259cm">
    <draw:object xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad">
      <xsl:attribute name="xlink:href">    
        <xsl:value-of select="$XlinkOLEObject"/>
      </xsl:attribute>
    </draw:object>
    </draw:frame>
      
    </xsl:if>
    
  </xsl:template>

</xsl:stylesheet>

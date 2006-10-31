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
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
  xmlns:uri="http://schemas.openxmlformats.org/drawingml/2006/picture"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink" 
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
  exclude-result-prefixes="w">

  <!-- Pictures conversion needs copy of image files in zipEnrty to work correctly (but id does't crash  -->
 
  <xsl:template match="wp:inline | wp:anchor">
    
    <xsl:variable name="document1">
      <xsl:call-template name="GetDocumentName">
        <xsl:with-param name="rootId">
          <xsl:value-of select="generate-id(/node())"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    
    <xsl:variable name="document">
    <xsl:choose>
      <xsl:when test="$document1 != ''">
        <xsl:value-of select="$document1"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:text>document.xml</xsl:text>
      </xsl:otherwise>
    </xsl:choose>
    </xsl:variable>
    
    <xsl:call-template name="CopyPictures">
      <xsl:with-param name="document">
        <xsl:value-of select="$document"/>
      </xsl:with-param>
    </xsl:call-template>
      
    <draw:frame text:anchor-type="paragraph">
      <!-- TODO: @draw:style-name @text:anchor-type -->
      

      <xsl:attribute name="draw:name">
        <xsl:value-of select="wp:docPr/@name"/>
      </xsl:attribute>
      <xsl:call-template name="SetSize"/>
      <xsl:if test="self::wp:anchor">
        <xsl:call-template name="SetPosition"/>
      </xsl:if>
      <draw:image xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad">
        <xsl:variable name="id">
          <xsl:value-of select="a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed"/>
        </xsl:variable>
        <xsl:if test="document(concat('word/_rels/',$document,'.rels'))">
          <xsl:for-each
            select="document(concat('word/_rels/',$document,'.rels'))//node()[name() = 'Relationship']">
            
            <xsl:if test="./@Id=$id">
              <xsl:variable name="pzipsource">
                <xsl:value-of select="./@Target"/>
              </xsl:variable>
              <xsl:variable name="pziptarget">
                <xsl:value-of select="substring-after($pzipsource,'/')"/>
              </xsl:variable>
              
              <xsl:attribute name="xlink:href">
                <xsl:value-of select="concat('pictures/', $pziptarget)"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
      </draw:image>
    </draw:frame>
  </xsl:template>

  <xsl:template name="SetSize">
    <xsl:attribute name="svg:height">
      <xsl:call-template name="emu-measure">
        <xsl:with-param name="length" select="wp:extent/@cy"/>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="svg:width">
      <xsl:call-template name="emu-measure">
        <xsl:with-param name="length" select="wp:extent/@cx"/>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>

  <xsl:template name="SetPosition">
    <xsl:attribute name="svg:x">
      <xsl:call-template name="emu-measure">
        <xsl:with-param name="length" select="wp:positionH/wp:posOffset"/>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="svg:y">
      <xsl:call-template name="emu-measure">
        <xsl:with-param name="length" select="wp:positionV/wp:posOffset"/>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>

</xsl:stylesheet>

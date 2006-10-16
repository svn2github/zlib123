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
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  exclude-result-prefixes="w table">

  <xsl:template match="w:tbl">
    <table:table>
      <xsl:attribute name="table:style-name">
        <xsl:value-of select="generate-id(self::w:tbl)"/>
      </xsl:attribute>
      <!--TODO: @table:style-name -->
      <xsl:apply-templates select="w:tblGrid/w:gridCol"/>
      <xsl:apply-templates select="w:tr"/>
    </table:table>
  </xsl:template>
  
  <xsl:template match="w:gridCol">
    <table:table-column>
      <xsl:attribute name="table:style-name">
        <xsl:value-of select="generate-id(self::w:gridCol)"/>
      </xsl:attribute>
      <!--TODO: @table:style-name -->
    </table:table-column>
  </xsl:template>
  
  <xsl:template match="w:tr">
    <xsl:choose>
      <xsl:when test="child::w:trPr/w:tblHeader">
        <table:table-header-rows>
          <table:table-row>
            <xsl:apply-templates select="w:tc"/>      
          </table:table-row>
        </table:table-header-rows>
      </xsl:when>
      <xsl:otherwise>
        <table:table-row>
          <xsl:apply-templates select="w:tc"/>      
        </table:table-row>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template match="w:tc">
    <table:table-cell>
      <xsl:attribute name="table:style-name">
        <xsl:value-of select="generate-id(self::w:tc)"/>
      </xsl:attribute>
      <!--TODO: @table:style-name -->
      <xsl:apply-templates/>
    </table:table-cell>
  </xsl:template>

  <xsl:template match="w:tblPr" mode="automaticstyles">
    <style:style style:name="{generate-id(parent::w:tbl)}" style:family="table">
      <xsl:if test="w:tblStyle">
        <xsl:attribute name="style:parent-style-name">
          <xsl:value-of select="w:tblStyle/@w:val"/>
        </xsl:attribute>
      </xsl:if>
      <style:table-properties>
        <xsl:call-template name="InsertTableProperties"/>
      </style:table-properties>
    </style:style>
  </xsl:template>
  
  
  <xsl:template match="w:tcPr" mode="automaticstyles">
    <style:style style:name="{generate-id(parent::w:tc)}" style:family="table-cell">
      <xsl:if test="w:tcStyle">
        <xsl:attribute name="style:parent-style-name">
          <xsl:value-of select="w:tcStyle/@w:val"/>
        </xsl:attribute>
      </xsl:if>
      <style:table-cell-properties>
        <xsl:call-template name="InsertCellsProperties"/>
      </style:table-cell-properties>
    </style:style>
  </xsl:template>
  
  
  <xsl:template match="w:gridCol" mode="automaticstyles">
    <style:style style:name="{generate-id(self::w:gridCol)}" style:family="table-column">
      <!--<xsl:if test="w:tblStyle">
        <xsl:attribute name="style:parent-style-name">
          <xsl:value-of select="w:tblStyle/@w:val"/>
        </xsl:attribute>
      </xsl:if>-->
      <style:table-column-properties>
        <xsl:call-template name="InsertColumnProperties"/>
      </style:table-column-properties>
    </style:style>
  </xsl:template>
  
  <xsl:template name="InsertTableProperties">
    <xsl:if test="w:tblW/@w:type = 'dxa'">
      <xsl:attribute name="style:width">
        <xsl:call-template name="ConvertTwips">
          <xsl:with-param name="length">
            <xsl:value-of select="w:tblW/@w:w"/>
          </xsl:with-param>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:choose>
        <xsl:when test="w:jc">
          <xsl:attribute name="table:align">
            <xsl:value-of select="w:jc/@w:val"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="table:align">left</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <xsl:if test="w:tblW/@w:type = 'auto'">
      <xsl:attribute name="style:width">
        <xsl:call-template name="ConvertTwips">
          <xsl:with-param name="length">
            <xsl:value-of select="sum(ancestor::w:tbl/w:tblGrid/w:gridCol/@w:w)"/>
          </xsl:with-param>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:choose>
        <xsl:when test="w:jc">
          <xsl:attribute name="table:align">
            <xsl:value-of select="w:jc/@w:val"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="table:align">left</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <xsl:if test="w:tblInd">
      <xsl:attribute name="fo:margin-left">
        <xsl:call-template name="ConvertTwips">
          <xsl:with-param name="length">
            <xsl:value-of select="w:tblInd/@w:w"/>
          </xsl:with-param>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>
  
  <xsl:template name="InsertColumnProperties">
    <xsl:attribute name="style:column-width">
      <xsl:call-template name="ConvertTwips">
        <xsl:with-param name="length">
          <xsl:value-of select="./@w:w"/>
        </xsl:with-param>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>
  
  <xsl:template name="InsertCellsProperties"></xsl:template>
</xsl:stylesheet>

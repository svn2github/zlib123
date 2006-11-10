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
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" exclude-result-prefixes="w table">

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
            <xsl:attribute name="table:style-name">
              <xsl:value-of select="generate-id(self::w:tr)"/>
            </xsl:attribute>
            <xsl:apply-templates select="w:tc"/>
          </table:table-row>
        </table:table-header-rows>
      </xsl:when>
      <xsl:otherwise>
        <table:table-row>
          <xsl:attribute name="table:style-name">
            <xsl:value-of select="generate-id(self::w:tr)"/>
          </xsl:attribute>
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

      <!-- column-span -->
      <xsl:if test="w:tcPr/w:gridSpan">
        <xsl:attribute name="table:number-columns-spanned">
          <xsl:value-of select="w:tcPr/w:gridSpan/@w:val"/>
        </xsl:attribute>
      </xsl:if>

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
      <style:table-column-properties>
        <xsl:call-template name="InsertColumnProperties"/>
      </style:table-column-properties>
    </style:style>
  </xsl:template>

  <xsl:template match="w:trPr" mode="automaticstyles">
    <style:style style:name="{generate-id(parent::w:tr)}" style:family="table-row">
      <style:table-row-properties>
        <xsl:call-template name="InsertRowProperties"/>
      </style:table-row-properties>
    </style:style>
  </xsl:template>

  <!--  insert table properties: width, align, indent-->
  <xsl:template name="InsertTableProperties">
    <xsl:choose>
      <xsl:when test="w:tblW/@w:type='pct'">
        <xsl:attribute name="style:rel-width">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:value-of select="w:tblW/@w:w"/>
            </xsl:with-param>
            <xsl:with-param name="unit">pct</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
        <xsl:attribute name="table:align">
          <xsl:call-template name="InsertTableAlign"/>
        </xsl:attribute>       
      </xsl:when>
      <xsl:when test="w:tblW/@w:type = 'dxa'">
        <xsl:attribute name="style:width">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:value-of select="w:tblW/@w:w"/>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
        <xsl:attribute name="table:align">
          <xsl:call-template name="InsertTableAlign"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="w:tblW/@w:type = 'dxa'">
        <xsl:attribute name="style:width">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:value-of select="sum(ancestor::w:tbl/w:tblGrid/w:gridCol/@w:w)"/>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
        <xsl:attribute name="table:align">
          <xsl:call-template name="InsertTableAlign"/>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
    
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

  <xsl:template name="InsertTableAlign">
    <xsl:choose>
      <xsl:when test="w:jc">
        <xsl:value-of select="w:jc/@w:val"/>
      </xsl:when>
      <xsl:when test="w:tblpPr/@w:tblpXSpec">
        <xsl:value-of select="w:tblpPr/@w:tblpXSpec"/>
      </xsl:when>
      <xsl:otherwise>left</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--  insert column properties: width-->
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

  <!--insert cells properties: vertical align, margins, borders  -->
  <xsl:template name="InsertCellsProperties">

    <!--    vertical align-->
    <xsl:if test="w:vAlign">
      <xsl:choose>
        <xsl:when test="w:vAlign/@w:val = 'center'">
          <xsl:attribute name="style:vertical-align">
            <xsl:value-of select="'middle'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="style:vertical-align">
            <xsl:value-of select="w:vAlign/@w:val"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <!--    margins-->
    <xsl:variable name="mstyleId">
      <xsl:value-of select="ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val"/>
    </xsl:variable>
    <xsl:if test="document('word/styles.xml')//w:styles/w:style/@w:styleId = $mstyleId">
      <xsl:variable name="mstyle"
        select="document('word/styles.xml')//w:styles/w:style/w:tblPr/w:tblCellMar"/>
      <xsl:call-template name="InsertCellMargins">
        <xsl:with-param name="tcMar" select="w:tcMar/w:bottom"/>
        <xsl:with-param name="tblMar" select="ancestor::w:tbl/w:tblPr/w:tblCellMar/w:bottom"/>
        <xsl:with-param name="tblDefMar" select="$mstyle/w:bottom"/>
        <xsl:with-param name="attribute">fo:padding-bottom</xsl:with-param>
      </xsl:call-template>
      <xsl:call-template name="InsertCellMargins">
        <xsl:with-param name="tcMar" select="w:tcMar/w:left"/>
        <xsl:with-param name="tblMar" select="ancestor::w:tbl/w:tblPr/w:tblCellMar/w:left"/>
        <xsl:with-param name="tblDefMar" select="$mstyle/w:left"/>
        <xsl:with-param name="attribute">fo:padding-left</xsl:with-param>
      </xsl:call-template>
      <xsl:call-template name="InsertCellMargins">
        <xsl:with-param name="tcMar" select="w:tcMar/w:right"/>
        <xsl:with-param name="tblMar" select="ancestor::w:tbl/w:tblPr/w:tblCellMar/w:right"/>
        <xsl:with-param name="tblDefMar" select="$mstyle/w:right"/>
        <xsl:with-param name="attribute">fo:padding-right</xsl:with-param>
      </xsl:call-template>
      <xsl:call-template name="InsertCellMargins">
        <xsl:with-param name="tcMar" select="w:tcMar/w:top"/>
        <xsl:with-param name="tblMar" select="ancestor::w:tbl/w:tblPr/w:tblCellMar/w:top"/>
        <xsl:with-param name="tblDefMar" select="$mstyle/w:top"/>
        <xsl:with-param name="attribute">fo:padding-top</xsl:with-param>
      </xsl:call-template>
    </xsl:if>

    <!--    borders-->
    <xsl:choose>
      <xsl:when test="w:tcBorders">
        <xsl:variable name="styleId">
          <xsl:value-of select="ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val"/>
        </xsl:variable>
        <xsl:variable name="style1"
          select="document('word/styles.xml')//w:styles/w:style[@w:styleId = $styleId]/w:tblPr/w:tblBorders"/>
        <xsl:variable name="style2"
          select="document('word/styles.xml')//w:styles/w:style[@w:default = 1 and @w:type = 'table']/w:tblPr/w:tblBorders"/>
        <xsl:variable name="style" select="$style1|$style2"/>
        <xsl:call-template name="InsertCellBorder">
          <xsl:with-param name="tcBorder" select="w:tcBorders/w:bottom"/>
          <xsl:with-param name="tblBorder" select="ancestor::w:tbl/w:tblPr/w:tblBorders/w:bottom"/>
          <xsl:with-param name="tblDefBorder" select="$style/w:bottom"/>
          <xsl:with-param name="attribute">fo:border-bottom</xsl:with-param>
          <xsl:with-param name="attribute2">style:border-line-width-bottom</xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="InsertCellBorder">
          <xsl:with-param name="tcBorder" select="w:tcBorders/w:right"/>
          <xsl:with-param name="tblBorder" select="ancestor::w:tbl/w:tblPr/w:tblBorders/w:right"/>
          <xsl:with-param name="tblDefBorder" select="$style/w:right"/>
          <xsl:with-param name="attribute">fo:border-right</xsl:with-param>
          <xsl:with-param name="attribute2">style:border-line-width-right</xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="InsertCellBorder">
          <xsl:with-param name="tcBorder" select="w:tcBorders/w:left"/>
          <xsl:with-param name="tblBorder" select="ancestor::w:tbl/w:tblPr/w:tblBorders/w:left"/>
          <xsl:with-param name="tblDefBorder" select="$style/w:left"/>
          <xsl:with-param name="attribute">fo:border-left</xsl:with-param>
          <xsl:with-param name="attribute2">style:border-line-width-left</xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="InsertCellBorder">
          <xsl:with-param name="tcBorder" select="w:tcBorders/w:top"/>
          <xsl:with-param name="tblBorder" select="ancestor::w:tbl/w:tblPr/w:tblBorders/w:top"/>
          <xsl:with-param name="tblDefBorder" select="$style/w:top"/>
          <xsl:with-param name="attribute">fo:border-top</xsl:with-param>
          <xsl:with-param name="attribute2">style:border-line-width-top</xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="styleId">
          <xsl:value-of select="ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val"/>
        </xsl:variable>
        <!--<xsl:if test="document('word/styles.xml')//w:styles/w:style/@w:styleId = $styleId">
          <xsl:variable name="style" select="document('word/styles.xml')//w:styles/w:style/w:tblPr/w:tblBorders"/>-->

        <xsl:variable name="style1"
          select="document('word/styles.xml')//w:styles/w:style[@w:styleId = $styleId]/w:tblPr/w:tblBorders"/>
        <xsl:variable name="style2"
          select="document('word/styles.xml')//w:styles/w:style[@w:default = 1 and @w:type = 'table']/w:tblPr/w:tblBorders"/>
        <xsl:variable name="style" select="$style1|$style2"/>
        <xsl:call-template name="InsertTableBorder">
          <xsl:with-param name="tblBorder" select="ancestor::w:tbl/w:tblPr/w:tblBorders/w:bottom"/>
          <xsl:with-param name="tblDefBorder" select="$style/w:bottom"/>
          <xsl:with-param name="attribute">fo:border-bottom</xsl:with-param>
          <xsl:with-param name="attribute2">style:border-line-width-bottom</xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="InsertTableBorder">
          <xsl:with-param name="tblBorder" select="ancestor::w:tbl/w:tblPr/w:tblBorders/w:right"/>
          <xsl:with-param name="tblDefBorder" select="$style/w:right"/>
          <xsl:with-param name="attribute">fo:border-right</xsl:with-param>
          <xsl:with-param name="attribute2">style:border-line-width-right</xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="InsertTableBorder">
          <xsl:with-param name="tblBorder" select="ancestor::w:tbl/w:tblPr/w:tblBorders/w:left"/>
          <xsl:with-param name="tblDefBorder" select="$style/w:left"/>
          <xsl:with-param name="attribute">fo:border-left</xsl:with-param>
          <xsl:with-param name="attribute2">style:border-line-width-left</xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="InsertTableBorder">
          <xsl:with-param name="tblBorder" select="ancestor::w:tbl/w:tblPr/w:tblBorders/w:top"/>
          <xsl:with-param name="tblDefBorder" select="$style/w:top"/>
          <xsl:with-param name="attribute">fo:border-top</xsl:with-param>
          <xsl:with-param name="attribute2">style:border-line-width-top</xsl:with-param>
        </xsl:call-template>
       <!--        </xsl:if>-->
      </xsl:otherwise>
    </xsl:choose>
    
<!-- Background color -->
    <xsl:if test="w:shd">
      <xsl:attribute name="fo:background-color">
        <xsl:value-of select="concat('#', w:shd/@w:fill)"/>
      </xsl:attribute>
    </xsl:if>

  </xsl:template>

  <!--  insert cell margins-->
  <xsl:template name="InsertCellMargins">
    <xsl:param name="tcMar"/>
    <xsl:param name="tblMar"/>
    <xsl:param name="tblDefMar"/>
    <xsl:param name="attribute"/>
    <xsl:choose>
      <xsl:when test="$tcMar">
        <xsl:attribute name="{$attribute}">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:value-of select="$tcMar/@w:w"/>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="$tblMar">
        <xsl:attribute name="{$attribute}">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:value-of select="$tblMar/@w:w"/>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="{$attribute}">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:value-of select="$tblDefMar/@w:w"/>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--  insert cell borders from cell properties-->
  <xsl:template name="InsertCellBorder">
    <xsl:param name="tcBorder"/>
    <xsl:param name="tblBorder"/>
    <xsl:param name="tblDefBorder"/>
    <xsl:param name="attribute"/>
    <xsl:param name="attribute2"/>
    <xsl:choose>
      <xsl:when test="$tcBorder">
        <xsl:choose>
          <xsl:when test="$tcBorder/@w:val = 'nil'">
            <xsl:attribute name="{$attribute}">
              <xsl:value-of select="'none'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:variable name="width">
              <xsl:call-template name="ConvertTwips">
                <xsl:with-param name="length">
                  <xsl:choose>
                    <xsl:when test="$tcBorder/@w:sz != 0">
                      <xsl:value-of select="$tcBorder/@w:sz"/>
                    </xsl:when>
                    <xsl:otherwise>2</xsl:otherwise>
                  </xsl:choose>
                </xsl:with-param>
                <xsl:with-param name="unit">cm</xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="style">
              <xsl:choose>
                <xsl:when test="$tcBorder/@w:val = 'single'">
                  <xsl:value-of select="'solid'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$tcBorder/@w:val"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:if test="$tcBorder/@w:val = 'double'">
              <xsl:variable name="width2">
                <xsl:call-template name="ConvertTwips">
                  <xsl:with-param name="length">
                    <xsl:value-of select="$tcBorder/@w:sz div 3"/>
                  </xsl:with-param>
                  <xsl:with-param name="unit">cm</xsl:with-param>
                </xsl:call-template>
              </xsl:variable>
              <xsl:attribute name="{$attribute2}">
                <xsl:value-of select="concat($width2,' ',$width2,' ',$width2)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:variable name="color">
              <xsl:choose>
                <xsl:when test="$tcBorder/@w:color = 'auto'">
                  <xsl:choose>
                    <xsl:when test="$tblBorder/@w:color = 'auto' or not($tblBorder/@w:color)">
                      <xsl:variable name="styleId">
                        <xsl:value-of select="ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val"/>
                      </xsl:variable>
                      <xsl:choose>
                        <xsl:when
                          test="document('word/styles.xml')//w:styles/w:style/@w:styleId = $styleId">
                          <xsl:value-of select="$tblDefBorder/@w:color"/>
                        </xsl:when>
                        <xsl:otherwise>000000</xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$tblBorder/@w:color"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$tcBorder/@w:color"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:attribute name="{$attribute}">
              <xsl:value-of select="concat($width,' ',$style,' #',$color)"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="InsertTableBorder">
          <xsl:with-param name="tblBorder" select="$tblBorder"/>
          <xsl:with-param name="tblDefBorder" select="$tblDefBorder"/>
          <xsl:with-param name="attribute">
            <xsl:value-of select="$attribute"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--  insert cell borders from table properties -->
  <xsl:template name="InsertTableBorder">
    <xsl:param name="tblBorder"/>
    <xsl:param name="tblDefBorder"/>
    <xsl:param name="attribute"/>
    <xsl:param name="attribute2"/>
    <xsl:choose>
      <xsl:when test="ancestor::w:tbl/w:tblPr/w:tblBorders">
        <xsl:variable name="width">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:value-of select="$tblBorder/@w:sz"/>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="style">
          <xsl:choose>
            <xsl:when test="$tblBorder/@w:val = 'single'">
              <xsl:value-of select="'solid'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$tblBorder/@w:val"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:if test="$tblBorder/@w:val = 'double'">
          <xsl:variable name="width2">
            <xsl:call-template name="ConvertTwips">
              <xsl:with-param name="length">
                <xsl:value-of select="$tblBorder/@w:sz div 3"/>
              </xsl:with-param>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:attribute name="{$attribute2}">
            <xsl:value-of select="concat($width2,' ',$width2,' ',$width2)"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:variable name="color">
          <xsl:choose>
            <xsl:when test="$tblBorder/@w:color = 'auto' or not($tblBorder/@w:color)">
              <xsl:variable name="styleId">
                <xsl:value-of select="ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val"/>
              </xsl:variable>
              <xsl:choose>
                <xsl:when test="document('word/styles.xml')//w:styles/w:style/@w:styleId = $styleId">
                  <xsl:value-of select="$tblDefBorder/@w:color"/>
                </xsl:when>
                <xsl:otherwise>000000</xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$tblBorder/@w:color"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:attribute name="{$attribute}">
          <xsl:value-of select="concat($width,' ',$style,' #',$color)"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="styleId">
          <xsl:value-of select="ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val"/>
        </xsl:variable>
        <xsl:if test="document('word/styles.xml')//w:styles/w:style/@w:styleId = $styleId">
          <xsl:variable name="width">
            <xsl:call-template name="ConvertTwips">
              <xsl:with-param name="length">
                <xsl:value-of select="$tblDefBorder/@w:sz"/>
              </xsl:with-param>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="style">
            <xsl:choose>
              <xsl:when test="$tblDefBorder/@w:val = 'single'">
                <xsl:value-of select="'solid'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$tblDefBorder/@w:val"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:if test="$tblDefBorder/@w:val = 'double'">
            <xsl:variable name="width2">
              <xsl:call-template name="ConvertTwips">
                <xsl:with-param name="length">
                  <xsl:value-of select="$tblDefBorder/@w:sz div 3"/>
                </xsl:with-param>
                <xsl:with-param name="unit">cm</xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:attribute name="{$attribute2}">
              <xsl:value-of select="concat($width2,' ',$width2,' ',$width2)"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:variable name="color">
            <xsl:value-of select="$tblDefBorder/@w:color"/>
          </xsl:variable>
          <xsl:attribute name="{$attribute}">
            <xsl:value-of select="concat($width,' ',$style,' #',$color)"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--  insert row properties: height-->
  <xsl:template name="InsertRowProperties">
    <xsl:if
      test="(w:trHeight/@w:hRule and w:trHeight/@w:hRule!='auto' ) or not(w:trHeight/@w:hRule)">
      <xsl:attribute name="style:min-row-height">
        <xsl:call-template name="ConvertTwips">
          <xsl:with-param name="length">
            <xsl:value-of select="w:trHeight/@w:val"/>
          </xsl:with-param>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

</xsl:stylesheet>

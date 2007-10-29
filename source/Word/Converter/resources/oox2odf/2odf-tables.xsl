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
  xmlns:oox="urn:oox"
  exclude-result-prefixes="w table oox">

  <!-- 
  *************************************************************************
  MATCHING TEMPLATES
  *************************************************************************
  -->
  
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
            <xsl:apply-templates />
          </table:table-row>
        </table:table-header-rows>
      </xsl:when>
      <xsl:otherwise>
        <table:table-row>
          <xsl:attribute name="table:style-name">
            <xsl:value-of select="generate-id(self::w:tr)"/>
          </xsl:attribute>
          <xsl:apply-templates />
        </table:table-row>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="w:tc">
    <xsl:choose>
      <!-- for w:vMerge="continuous" cells, create a table:covered-table-cell element --> 
      <xsl:when test="w:tcPr/w:vMerge and not(w:tcPr/w:vMerge/@w:val = 'restart')">
        <table:covered-table-cell table:style-name="{generate-id()}">
          <xsl:apply-templates/>
        </table:covered-table-cell>
      </xsl:when>
      <xsl:otherwise>
        <!-- create normal a table cell -->
        <table:table-cell table:style-name="{generate-id()}">
          <!-- column-span -->
          <xsl:if test="w:tcPr/w:gridSpan and w:tcPr/w:gridSpan/@w:val != '0'">
            <xsl:attribute name="table:number-columns-spanned">
              <xsl:value-of select="w:tcPr/w:gridSpan/@w:val"/>
            </xsl:attribute>
          </xsl:if>
          <!-- new row-span (w:vMerge="restart") -->
          <xsl:if test="w:tcPr/w:vMerge and w:tcPr/w:vMerge/@w:val = 'restart' ">
            <xsl:variable name="number-rows-spanned">
              <xsl:call-template name="ComputeNumberRowsSpanned">
                <xsl:with-param name="cellPosition">
                  <xsl:call-template name="GetCellPosition"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:attribute name="table:number-rows-spanned">
              <xsl:choose>
                <!-- patch: can't have number-rows-spanned < 2 -->
                <xsl:when test="$number-rows-spanned &lt; 2">2</xsl:when>
                <xsl:otherwise><xsl:value-of select="$number-rows-spanned"/></xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
          <xsl:apply-templates/>
        </table:table-cell>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="w:tblPr" mode="automaticstyles">
    <style:style style:name="{generate-id(parent::w:tbl)}" style:family="table">
      <xsl:if test="w:tblStyle">
        <xsl:attribute name="style:parent-style-name">
          <xsl:value-of select="w:tblStyle/@w:val"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:call-template name="MasterPageName"/>
      <style:table-properties table:border-model="collapsing">
        <xsl:call-template name="InsertTableProperties">
          <xsl:with-param name="Default">StyleTableProperties</xsl:with-param>
        </xsl:call-template>
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

  <!-- 
  *************************************************************************
  CALLED TEMPLATES
  *************************************************************************
  -->

  <!-- compute the number of rows that are spanned by context cell -->
  <xsl:template name="ComputeNumberRowsSpanned">
    <xsl:param name="cellPosition"/>
    <xsl:param name="rowCount" select="1"/>

    <xsl:choose>
      <xsl:when test="ancestor::w:tr[1]/following-sibling::w:tr[1]">
        <xsl:for-each select="ancestor::w:tr[1]/following-sibling::w:tr[1]/w:tc[1]">
          <xsl:call-template name="ComputeNumberRowsSpannedUsingCells">
            <xsl:with-param name="cellPosition" select="$cellPosition"/>
            <xsl:with-param name="rowCount" select="$rowCount"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$rowCount"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="ComputeNumberRowsSpannedUsingCells">
    <xsl:param name="cellPosition"/>
    <xsl:param name="rowCount"/>

    <xsl:variable name="currentPosition">
      <xsl:call-template name="GetCellPosition"/>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$currentPosition = $cellPosition">
        <xsl:choose>
          <!-- go to the following row -->
          <xsl:when test="w:tcPr/w:vMerge and not(w:tcPr/w:vMerge/@w:val = 'restart')">
             <xsl:choose>
               <xsl:when test="not(parent::w:tr/following-sibling::w:tr[1]/w:tc[1])">
                 <!-- Bug #1694962 -->
                 <xsl:value-of select="$rowCount +2"/>
               </xsl:when>
               <xsl:otherwise>
                 <xsl:for-each select="parent::w:tr/following-sibling::w:tr[1]/w:tc[1]">
                   <xsl:call-template name="ComputeNumberRowsSpannedUsingCells">
                     <xsl:with-param name="cellPosition" select="$cellPosition"/>
                     <xsl:with-param name="rowCount" select="$rowCount+1"/>
                   </xsl:call-template>
                 </xsl:for-each>
               </xsl:otherwise>
             </xsl:choose>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$rowCount"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <!-- go to the following cell -->
      <xsl:when test="$currentPosition &lt; $cellPosition and following-sibling::w:tc[1]">
        <xsl:for-each select="following-sibling::w:tc[1]">
          <xsl:call-template name="ComputeNumberRowsSpannedUsingCells">
            <xsl:with-param name="cellPosition" select="$cellPosition"/>
            <xsl:with-param name="rowCount" select="$rowCount"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$rowCount"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- find the position of context cell -->
  <xsl:template name="GetCellPosition">
    <xsl:param name="currentPosition" select="1"/>
    <xsl:choose>
      <xsl:when test="preceding-sibling::w:tc">
        <xsl:for-each select="preceding-sibling::w:tc[1]">
          <xsl:choose>
            <xsl:when test="w:tcPr/w:gridSpan and w:tcPr/w:gridSpan/@w:val != '0' ">
              <xsl:call-template name="GetCellPosition">
                <xsl:with-param name="currentPosition"
                  select="$currentPosition+number(w:tcPr/w:gridSpan/@w:val)"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="GetCellPosition">
                <xsl:with-param name="currentPosition" select="$currentPosition+1"/>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$currentPosition"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--  insert table properties: width, align, indent-->
  <xsl:template name="InsertTableProperties">
    <xsl:param name="Default"/>
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
      <xsl:when test="w:tblW/@w:type = 'auto'">
        <xsl:attribute name="style:width">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:value-of select="sum(ancestor::w:tbl[1]/w:tblGrid/w:gridCol/@w:w)"/>
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

    <xsl:if
      test="parent::w:tbl/descendant::w:pageBreakBefore and 
      not(generate-id(parent::w:tbl) = generate-id(ancestor::w:body/child::node()[1]))">
      <xsl:choose>
        <xsl:when test="not(parent::w:tbl/preceding::w:p[1]/w:pPr/w:sectPr)">
          <xsl:attribute name="fo:break-before">
            <xsl:text>page</xsl:text>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:variable name="precSectPr" select="preceding::w:p[1]/w:pPr/w:sectPr"/>
          <xsl:variable name="followingSectPr" select="following::w:sectPr"/>
          <xsl:if
            test="($precSectPr/w:pgSz/@w:w = $followingSectPr/w:pgSz/@w:w
            and $precSectPr/w:pgSz/@w:h = $followingSectPr/w:pgSz/@w:h
            and $precSectPr/w:pgSz/@w:orient = $followingSectPr/w:pgSz/@w:orient
            and $precSectPr/w:pgMar/@w:top = $followingSectPr/w:pgMar/@w:top
            and $precSectPr/w:pgMar/@w:left = $followingSectPr/w:pgMar/@w:left
            and $precSectPr/w:pgMar/@w:right = $followingSectPr/w:pgMar/@w:right
            and $precSectPr/w:pgMar/@w:bottom = $followingSectPr/w:pgMar/@w:bottom
            and $precSectPr/w:pgMar/@w:header = $followingSectPr/w:pgMar/@w:header
            and $precSectPr/w:pgMar/@w:footer = $followingSectPr/w:pgMar/@w:footer)">
            <xsl:attribute name="fo:break-before">
              <xsl:text>page</xsl:text>
            </xsl:attribute>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="$Default='StyleTableProperties'">
      <xsl:choose>
        <xsl:when test="w:tblpPr/@w:bottomFromText">
          <xsl:attribute name="fo:margin-bottom">
            <xsl:call-template name="ConvertTwips">
              <xsl:with-param name="length">
                <xsl:value-of select="w:tblpPr/@w:bottomFromText"/>
              </xsl:with-param>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="fo:margin-bottom">0cm</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
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

  <!--
  Summary: insert cells properties: vertical align, margins, borders  
  Author: Clever Age
  Modified: makz (DIaLOGIKa)
  Date: 26.10.2007
  -->
  <xsl:template name="InsertCellsProperties">

    <!-- vertical align -->
    <xsl:call-template name="InsertCellVAlign"/>

    <!-- margins-->
    <xsl:variable name="mstyleId">
      <xsl:value-of select="ancestor::w:tbl[1]/w:tblPr/w:tblStyle/@w:val"/>
    </xsl:variable>
    
    <xsl:choose>
      <xsl:when test="key('Part', 'word/styles.xml')/w:styles/w:style[@w:styleId=$mstyleId or @w:styleId=concat('CONTENT_',$mstyleId)]">
        <xsl:variable name="mstyle" select="key('Part', 'word/styles.xml')/w:styles/w:style[@w:styleId = $mstyleId or @w:styleId = concat('CONTENT_',$mstyleId)]/w:tblPr/w:tblCellMar"/>
        <!-- margin is specified in styles.xml -->
        <xsl:call-template name="InsertCellMargins">
          <xsl:with-param name="tcMar" select="w:tcMar/w:bottom"/>
          <xsl:with-param name="tblMar" select="ancestor::w:tbl[1]/w:tblPr/w:tblCellMar/w:bottom"/>
          <xsl:with-param name="tblDefMar" select="$mstyle/w:bottom"/>
          <xsl:with-param name="attribute">fo:padding-bottom</xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="InsertCellMargins">
          <xsl:with-param name="tcMar" select="w:tcMar/w:left"/>
          <xsl:with-param name="tblMar" select="ancestor::w:tbl[1]/w:tblPr/w:tblCellMar/w:left"/>
          <xsl:with-param name="tblDefMar" select="$mstyle/w:left"/>
          <xsl:with-param name="attribute">fo:padding-left</xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="InsertCellMargins">
          <xsl:with-param name="tcMar" select="w:tcMar/w:right"/>
          <xsl:with-param name="tblMar" select="ancestor::w:tbl[1]/w:tblPr/w:tblCellMar/w:right"/>
          <xsl:with-param name="tblDefMar" select="$mstyle/w:right"/>
          <xsl:with-param name="attribute">fo:padding-right</xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="InsertCellMargins">
          <xsl:with-param name="tcMar" select="w:tcMar/w:top"/>
          <xsl:with-param name="tblMar" select="ancestor::w:tbl[1]/w:tblPr/w:tblCellMar/w:top"/>
          <xsl:with-param name="tblDefMar" select="$mstyle/w:top"/>
          <xsl:with-param name="attribute">fo:padding-top</xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="w:tcMar or ancestor::w:tbl[1]/w:tblPr/w:tblCellMar">
        <!-- margin is specified direct in the table -->
        <xsl:call-template name="InsertCellMargins">
          <xsl:with-param name="tcMar" select="w:tcMar/w:bottom"/>
          <xsl:with-param name="tblMar" select="ancestor::w:tbl[1]/w:tblPr/w:tblCellMar/w:bottom"/>
          <xsl:with-param name="attribute">fo:padding-bottom</xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="InsertCellMargins">
          <xsl:with-param name="tcMar" select="w:tcMar/w:left"/>
          <xsl:with-param name="tblMar" select="ancestor::w:tbl[1]/w:tblPr/w:tblCellMar/w:left"/>
          <xsl:with-param name="attribute">fo:padding-left</xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="InsertCellMargins">
          <xsl:with-param name="tcMar" select="w:tcMar/w:right"/>
          <xsl:with-param name="tblMar" select="ancestor::w:tbl[1]/w:tblPr/w:tblCellMar/w:right"/>
          <xsl:with-param name="attribute">fo:padding-right</xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="InsertCellMargins">
          <xsl:with-param name="tcMar" select="w:tcMar/w:top"/>
          <xsl:with-param name="tblMar" select="ancestor::w:tbl[1]/w:tblPr/w:tblCellMar/w:top"/>
          <xsl:with-param name="attribute">fo:padding-top</xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <!-- use application dafaults (no margin specified) -->
        <xsl:attribute name="fo:padding-top">
          <xsl:text>0cm</xsl:text>
        </xsl:attribute>
        <xsl:attribute name="fo:padding-bottom">
          <xsl:text>0cm</xsl:text>
        </xsl:attribute>
        <xsl:attribute name="fo:padding-left">
          <xsl:text>0.19cm</xsl:text>
        </xsl:attribute>
        <xsl:attribute name="fo:padding-right">
          <xsl:text>0.19cm</xsl:text>
        </xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>
    
    <!--    borders-->
    <xsl:call-template name="InsertCellBorders"/>

    <!-- Background color -->
    <xsl:if test="w:shd">
      <xsl:attribute name="fo:background-color">
        <xsl:for-each select="w:shd">
          <xsl:call-template name="ComputeShading"/>
          </xsl:for-each>
      </xsl:attribute>
    </xsl:if>
    
  </xsl:template>
 
  <!-- insert table cell content vertical alignment -->
  <xsl:template name="InsertCellVAlign">
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
      <xsl:when test="$tblDefMar">
        <xsl:attribute name="{$attribute}">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:value-of select="$tblDefMar/@w:w"/>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- insert cell borders using cell properties -->
  <xsl:template name="InsertCellBorders">
    <xsl:call-template name="InsertTopBottomBorderPropertiesUsingSide">
      <xsl:with-param name="side">bottom</xsl:with-param>
    </xsl:call-template>
    <xsl:call-template name="InsertLeftRightBorderPropertiesUsingSide">
      <xsl:with-param name="side">right</xsl:with-param>
    </xsl:call-template>
    <xsl:call-template name="InsertLeftRightBorderPropertiesUsingSide">
      <xsl:with-param name="side">left</xsl:with-param>
    </xsl:call-template>
    <xsl:call-template name="InsertTopBottomBorderPropertiesUsingSide">
      <xsl:with-param name="side">top</xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <!-- find border properties element, for top and bottom sides -->
  <xsl:template name="InsertTopBottomBorderPropertiesUsingSide">
    <xsl:param name="side"/>
    <xsl:param name="contextCell" select="parent::w:tc"/>
    <xsl:param name="styleId" select="ancestor::w:tbl[1]/w:tblPr/w:tblStyle/@w:val"/>
    <xsl:param name="tcBorders" select="w:tcBorders"/>
    <xsl:param name="tblBorders" select="ancestor::w:tbl[1]/w:tblPr/w:tblBorders"/>

    <xsl:choose>
      <!-- use cell definition -->
      <xsl:when test="$side = 'top' and $tcBorders/w:top">
        <xsl:call-template name="InsertCellBorder">
          <xsl:with-param name="border" select="$tcBorders/w:top"/>
          <xsl:with-param name="side" select="$side"/>
          <xsl:with-param name="styleId" select="$styleId"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$side = 'top' and $tcBorders/w:insideV">
        <xsl:call-template name="InsertCellBorder">
          <xsl:with-param name="border" select="$tcBorders/w:insideV"/>
          <xsl:with-param name="side" select="$side"/>
          <xsl:with-param name="styleId" select="$styleId"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$side = 'bottom' and $tcBorders/w:bottom">
        <xsl:call-template name="InsertCellBorder">
          <xsl:with-param name="border" select="$tcBorders/w:bottom"/>
          <xsl:with-param name="side" select="$side"/>
          <xsl:with-param name="styleId" select="$styleId"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$side = 'bottom' and $tcBorders/w:insideV">
        <xsl:call-template name="InsertCellBorder">
          <xsl:with-param name="border" select="$tcBorders/w:insideV"/>
          <xsl:with-param name="side" select="$side"/>
          <xsl:with-param name="styleId" select="$styleId"/>
        </xsl:call-template>
      </xsl:when>
      <!-- use table definition -->
      <xsl:when
        test="$side = 'top' and $tblBorders/w:top and not($contextCell/parent::w:tr/preceding-sibling::w:tr)">
        <xsl:call-template name="InsertCellBorder">
          <xsl:with-param name="border" select="$tblBorders/w:top"/>
          <xsl:with-param name="side" select="$side"/>
          <xsl:with-param name="styleId" select="$styleId"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when
        test="$side = 'top' and $tblBorders/w:insideV and $contextCell/parent::w:tr/preceding-sibling::w:tr">
        <xsl:call-template name="InsertCellBorder">
          <xsl:with-param name="border" select="$tblBorders/w:insideV"/>
          <xsl:with-param name="side" select="$side"/>
          <xsl:with-param name="styleId" select="$styleId"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when
        test="$side = 'bottom' and $tblBorders/w:bottom and not($contextCell/parent::w:tr/following-sibling::w:tr)">
        <xsl:call-template name="InsertCellBorder">
          <xsl:with-param name="border" select="$tblBorders/w:bottom"/>
          <xsl:with-param name="side" select="$side"/>
          <xsl:with-param name="styleId" select="$styleId"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when
        test="$side = 'bottom' and $tblBorders/w:insideV and $contextCell/parent::w:tr/following-sibling::w:tr">
        <xsl:call-template name="InsertCellBorder">
          <xsl:with-param name="border" select="$tblBorders/w:insideV"/>
          <xsl:with-param name="side" select="$side"/>
          <xsl:with-param name="styleId" select="$styleId"/>
        </xsl:call-template>
      </xsl:when>
      <!-- climb style hierarchy -->
      <xsl:otherwise>
        <xsl:for-each select="key('Part', 'word/styles.xml')">
          <xsl:choose>
            <xsl:when test="key('StyleId', $styleId)">
              <xsl:call-template name="InsertTopBottomBorderPropertiesUsingSide">
                <xsl:with-param name="side" select="$side"/>
                <xsl:with-param name="contextCell" select="$contextCell"/>
                <xsl:with-param name="styleId"
                  select="key('StyleId', $styleId)/w:tblPr/w:tblStyle/@w:val"/>
                <xsl:with-param name="tcBorders"
                  select="key('StyleId', $styleId)/w:tcPr/w:tcBorders"/>
                <xsl:with-param name="tblBorders"
                  select="key('StyleId', $styleId)/w:tblPr/w:tblBorders"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="key('default-styles', 'table')/w:tblPr/w:tblStyle/@w:val != $styleId">
                <xsl:call-template name="InsertTopBottomBorderPropertiesUsingSide">
                  <xsl:with-param name="side" select="$side"/>
                  <xsl:with-param name="contextCell" select="$contextCell"/>
                  <xsl:with-param name="styleId"
                    select="key('default-styles', 'table')/w:tblPr/w:tblStyle/@w:val"/>
                  <xsl:with-param name="tcBorders"
                    select="key('default-styles', 'table')/w:tcPr/w:tcBorders"/>
                  <xsl:with-param name="tblBorders"
                    select="key('default-styles', 'table')/w:tblPr/w:tblBorders"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- find border properties element, for right and left sides -->
  <xsl:template name="InsertLeftRightBorderPropertiesUsingSide">
    <xsl:param name="side"/>
    <xsl:param name="contextCell" select="parent::w:tc"/>
    <xsl:param name="styleId" select="ancestor::w:tbl[1]/w:tblPr/w:tblStyle/@w:val"/>
    <xsl:param name="tcBorders" select="w:tcBorders"/>
    <xsl:param name="tblBorders" select="ancestor::w:tbl[1]/w:tblPr/w:tblBorders"/>

    <xsl:choose>
      <!-- use cell definition -->
      <xsl:when test="$side = 'left' and $tcBorders/w:left">
        <xsl:call-template name="InsertCellBorder">
          <xsl:with-param name="border" select="$tcBorders/w:left"/>
          <xsl:with-param name="side" select="$side"/>
          <xsl:with-param name="styleId" select="$styleId"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$side = 'left' and $tcBorders/w:insideH">
        <xsl:call-template name="InsertCellBorder">
          <xsl:with-param name="border" select="$tcBorders/w:insideH"/>
          <xsl:with-param name="side" select="$side"/>
          <xsl:with-param name="styleId" select="$styleId"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$side = 'right' and $tcBorders/w:right">
        <xsl:call-template name="InsertCellBorder">
          <xsl:with-param name="border" select="$tcBorders/w:right"/>
          <xsl:with-param name="side" select="$side"/>
          <xsl:with-param name="styleId" select="$styleId"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$side = 'right' and $tcBorders/w:insideH">
        <xsl:call-template name="InsertCellBorder">
          <xsl:with-param name="border" select="$tcBorders/w:insideH"/>
          <xsl:with-param name="side" select="$side"/>
          <xsl:with-param name="styleId" select="$styleId"/>
        </xsl:call-template>
      </xsl:when>
      <!-- use table definition -->
      <xsl:when
        test="$side = 'left' and $tblBorders/w:left and not($contextCell/preceding-sibling::w:tc)">
        <xsl:call-template name="InsertCellBorder">
          <xsl:with-param name="border" select="$tblBorders/w:left"/>
          <xsl:with-param name="side" select="$side"/>
          <xsl:with-param name="styleId" select="$styleId"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when
        test="$side = 'left' and $tblBorders/w:insideH and $contextCell/preceding-sibling::w:tc">
        <xsl:call-template name="InsertCellBorder">
          <xsl:with-param name="border" select="$tblBorders/w:insideH"/>
          <xsl:with-param name="side" select="$side"/>
          <xsl:with-param name="styleId" select="$styleId"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when
        test="$side = 'right' and $tblBorders/w:right and not($contextCell/following-sibling::w:tc)">
        <xsl:call-template name="InsertCellBorder">
          <xsl:with-param name="border" select="$tblBorders/w:right"/>
          <xsl:with-param name="side" select="$side"/>
          <xsl:with-param name="styleId" select="$styleId"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when
        test="$side = 'right' and $tblBorders/w:insideH and $contextCell/following-sibling::w:tc">
        <xsl:call-template name="InsertCellBorder">
          <xsl:with-param name="border" select="$tblBorders/w:insideH"/>
          <xsl:with-param name="side" select="$side"/>
          <xsl:with-param name="styleId" select="$styleId"/>
        </xsl:call-template>
      </xsl:when>
      <!-- climb style hierarchy -->
      <xsl:otherwise>
        <xsl:for-each select="key('Part', 'word/styles.xml')">
          <xsl:choose>
            <xsl:when test="key('StyleId', $styleId)">
              <xsl:call-template name="InsertLeftRightBorderPropertiesUsingSide">
                <xsl:with-param name="side" select="$side"/>
                <xsl:with-param name="contextCell" select="$contextCell"/>
                <xsl:with-param name="styleId"
                  select="key('StyleId', $styleId)/w:tblPr/w:tblStyle/@w:val"/>
                <xsl:with-param name="tcBorders"
                  select="key('StyleId', $styleId)/w:tcPr/w:tcBorders"/>
                <xsl:with-param name="tblBorders"
                  select="key('StyleId', $styleId)/w:tblPr/w:tblBorders"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="key('default-styles', 'table')/w:tblPr/w:tblStyle/@w:val != $styleId">
                <xsl:call-template name="InsertLeftRightBorderPropertiesUsingSide">
                  <xsl:with-param name="side" select="$side"/>
                  <xsl:with-param name="contextCell" select="$contextCell"/>
                  <xsl:with-param name="styleId"
                    select="key('default-styles', 'table')/w:tblPr/w:tblStyle/@w:val"/>
                  <xsl:with-param name="tcBorders"
                    select="key('default-styles', 'table')/w:tcPr/w:tcBorders"/>
                  <xsl:with-param name="tblBorders"
                    select="key('default-styles', 'table')/w:tblPr/w:tblBorders"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--  insert cell borders from cell properties-->
  <xsl:template name="InsertCellBorder">
    <xsl:param name="side"/>
    <xsl:param name="border"/>
    <xsl:param name="styleId"/>

    <xsl:if test="$border">
      <xsl:choose>
        <xsl:when test="$border/@w:val = 'nil' or $border/@w:val = 'none' ">
          <xsl:attribute name="{concat('fo:border-', $side)}">
            <xsl:value-of select="'none'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <!-- compute width -->
          <xsl:variable name="width">
            <xsl:choose>
              <xsl:when test="$border/@w:sz != 0">
                <xsl:call-template name="ConvertEighthsPoints">
                  <xsl:with-param name="length" select="$border/@w:sz"/>
                  <xsl:with-param name="unit">cm</xsl:with-param>
                </xsl:call-template>
              </xsl:when>
              <!-- default value (arbitrary) -->
              <xsl:otherwise>0.002cm</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <!-- compute style -->
          <xsl:variable name="style">
            <xsl:choose>
              <xsl:when test="$border/@w:val = 'single'">
                <xsl:value-of select="'solid'"/>
              </xsl:when>
              <xsl:otherwise>
                <!-- Styles of Table Border to do (Open Office Bug)-->
                <xsl:value-of select="'solid'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <!-- compute color -->
          <xsl:variable name="color">
            <xsl:choose>
              <xsl:when test="$border/@w:color = 'auto' ">
                <xsl:choose>
                  <xsl:when test="$side = 'top' or $side = 'bottom' ">
                    <xsl:call-template name="GetTopBottomCellBorderColor">
                      <xsl:with-param name="side" select="$side"/>
                      <xsl:with-param name="styleId" select="$styleId"/>
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:when test="$side = 'left' or $side = 'right' ">
                    <xsl:call-template name="GetLeftRightCellBorderColor">
                      <xsl:with-param name="side" select="$side"/>
                      <xsl:with-param name="styleId" select="$styleId"/>
                    </xsl:call-template>
                  </xsl:when>
                </xsl:choose>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$border/@w:color"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <!-- write fo:border attribute -->
          <xsl:attribute name="{concat('fo:border-', $side)}">
            <xsl:value-of select="concat($width,' ',$style,' #',$color)"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
  </xsl:template>

  <!-- find the color in case it is set to automatic -->
  <xsl:template name="GetTopBottomCellBorderColor">
    <xsl:param name="side"/>
    <xsl:param name="styleId"/>
    <xsl:param name="contextCell" select="parent::w:tc"/>
    <xsl:param name="tcBorders" select="w:tcBorders"/>
    <xsl:param name="tblBorders" select="ancestor::w:tbl[1]/w:tblPr/w:tblBorders"/>

    <xsl:choose>
      <!-- use cell definition -->
      <xsl:when
        test="$side = 'top' and $tcBorders/w:top/@color != 'auto' and not($contextCell/parent::w:tr/preceding-sibling::w:tr)">
        <xsl:value-of select="$tcBorders/w:top/@color"/>
      </xsl:when>
      <xsl:when
        test="$side = 'top' and $tcBorders/w:insideH/@color != 'auto' and $contextCell/parent::w:tr/preceding-sibling::w:tr">
        <xsl:value-of select="$tcBorders/w:insideH/@color"/>
      </xsl:when>
      <xsl:when
        test="$side = 'bottom' and $tcBorders/w:bottom/@color != 'auto' and not($contextCell/parent::w:tr/following-sibling::w:tr)">
        <xsl:value-of select="$tcBorders/w:bottom/@color"/>
      </xsl:when>
      <xsl:when
        test="$side = 'bottom' and $tcBorders/w:insideH/@color != 'auto' and $contextCell/parent::w:tr/following-sibling::w:tr">
        <xsl:value-of select="$tcBorders/w:insideH/@color"/>
      </xsl:when>
      <!-- use table definition -->
      <xsl:when
        test="$side = 'top' and $tblBorders/w:left/@color != 'auto' and not($contextCell/parent::w:tr/preceding-sibling::w:tr)">
        <xsl:value-of select="$tblBorders/w:top/@color"/>
      </xsl:when>
      <xsl:when
        test="$side = 'top' and $tblBorders/w:insideH/@color != 'auto' and $contextCell/parent::w:tr/preceding-sibling::w:tr">
        <xsl:value-of select="$tblBorders/w:insideH/@color"/>
      </xsl:when>
      <xsl:when
        test="$side = 'bottom' and $tblBorders/w:bottom/@color != 'auto' and not($contextCell/parent::w:tr/following-sibling::w:tr)">
        <xsl:value-of select="$tblBorders/w:right/@color"/>
      </xsl:when>
      <xsl:when
        test="$side = 'bottom' and $tblBorders/w:insideH/@color != 'auto' and $contextCell/parent::w:tr/following-sibling::w:tr">
        <xsl:value-of select="$tblBorders/w:insideH/@color"/>
      </xsl:when>
      <!-- climb style hierarchy -->
      <xsl:otherwise>
        <xsl:for-each select="key('Part', 'word/styles.xml')">
          <xsl:choose>
            <xsl:when test="key('StyleId', $styleId)">
              <xsl:call-template name="GetTopBottomCellBorderColor">
                <xsl:with-param name="side" select="$side"/>
                <xsl:with-param name="contextCell" select="$contextCell"/>
                <xsl:with-param name="styleId"
                  select="key('StyleId', $styleId)/w:tblPr/w:tblStyle/@w:val"/>
                <xsl:with-param name="tcBorders"
                  select="key('StyleId', $styleId)/w:tcPr/w:tcBorders"/>
                <xsl:with-param name="tblBorders"
                  select="key('StyleId', $styleId)/w:tblPr/w:tblBorders"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when
                  test="key('default-styles', 'table')/w:tblPr/w:tblStyle/@w:val != $styleId">
                  <xsl:call-template name="GetTopBottomCellBorderColor">
                    <xsl:with-param name="side" select="$side"/>
                    <xsl:with-param name="contextCell" select="$contextCell"/>
                    <xsl:with-param name="styleId"
                      select="key('default-styles', 'table')/w:tblPr/w:tblStyle/@w:val"/>
                    <xsl:with-param name="tcBorders"
                      select="key('default-styles', 'table')/w:tcPr/w:tcBorders"/>
                    <xsl:with-param name="tblBorders"
                      select="key('default-styles', 'table')/w:tblPr/w:tblBorders"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>000000</xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- find the color in case it is set to automatic -->
  <xsl:template name="GetLeftRightCellBorderColor">
    <xsl:param name="side"/>
    <xsl:param name="styleId"/>
    <xsl:param name="contextCell" select="parent::w:tc"/>
    <xsl:param name="tcBorders" select="w:tcBorders"/>
    <xsl:param name="tblBorders" select="ancestor::w:tbl[1]/w:tblPr/w:tblBorders"/>

    <xsl:choose>
      <!-- use cell definition -->
      <xsl:when
        test="$side = 'left' and $tcBorders/w:left/@color != 'auto' and not($contextCell/preceding-sibling::w:tc)">
        <xsl:value-of select="$tcBorders/w:left/@color"/>
      </xsl:when>
      <xsl:when
        test="$side = 'left' and $tcBorders/w:insideH/@color != 'auto' and $contextCell/preceding-sibling::w:tc">
        <xsl:value-of select="$tcBorders/w:insideH/@color"/>
      </xsl:when>
      <xsl:when
        test="$side = 'right' and $tcBorders/w:bottom/@color != 'auto' and not($contextCell/following-sibling::w:tc)">
        <xsl:value-of select="$tcBorders/w:right/@color"/>
      </xsl:when>
      <xsl:when
        test="$side = 'right' and $tcBorders/w:insideH/@color != 'auto' and $contextCell/following-sibling::w:tc">
        <xsl:value-of select="$tcBorders/w:insideH/@color"/>
      </xsl:when>
      <!-- use table definition -->
      <xsl:when
        test="$side = 'left' and $tblBorders/w:left/@color != 'auto' and not($contextCell/preceding-sibling::w:tc)">
        <xsl:value-of select="$tblBorders/w:left/@color"/>
      </xsl:when>
      <xsl:when
        test="$side = 'left' and $tblBorders/w:insideH/@color != 'auto' and $contextCell/preceding-sibling::w:tc">
        <xsl:value-of select="$tblBorders/w:insideH/@color"/>
      </xsl:when>
      <xsl:when
        test="$side = 'right' and $tblBorders/w:right/@color != 'auto' and not($contextCell/following-sibling::w:tc)">
        <xsl:value-of select="$tblBorders/w:right/@color"/>
      </xsl:when>
      <xsl:when
        test="$side = 'right' and $tblBorders/w:insideH/@color != 'auto' and $contextCell/following-sibling::w:tc">
        <xsl:value-of select="$tblBorders/w:insideH/@color"/>
      </xsl:when>
      <!-- climb style hierarchy -->
      <xsl:otherwise>
        <xsl:for-each select="key('Part', 'word/styles.xml')">
          <xsl:choose>
            <xsl:when test="key('StyleId', $styleId)">
              <xsl:call-template name="GetLeftRightCellBorderColor">
                <xsl:with-param name="side" select="$side"/>
                <xsl:with-param name="contextCell" select="$contextCell"/>
                <xsl:with-param name="styleId"
                  select="key('StyleId', $styleId)/w:tblPr/w:tblStyle/@w:val"/>
                <xsl:with-param name="tcBorders"
                  select="key('StyleId', $styleId)/w:tcPr/w:tcBorders"/>
                <xsl:with-param name="tblBorders"
                  select="key('StyleId', $styleId)/w:tblPr/w:tblBorders"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when
                  test="key('default-styles', 'table')/w:tblPr/w:tblStyle/@w:val != $styleId">
                  <xsl:call-template name="GetLeftRightCellBorderColor">
                    <xsl:with-param name="side" select="$side"/>
                    <xsl:with-param name="contextCell" select="$contextCell"/>
                    <xsl:with-param name="styleId"
                      select="key('default-styles', 'table')/w:tblPr/w:tblStyle/@w:val"/>
                    <xsl:with-param name="tcBorders"
                      select="key('default-styles', 'table')/w:tcPr/w:tcBorders"/>
                    <xsl:with-param name="tblBorders"
                      select="key('default-styles', 'table')/w:tblPr/w:tblBorders"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>000000</xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--  insert row properties: height-->
  <xsl:template name="InsertRowProperties">
    <xsl:choose>
      <xsl:when test="w:trHeight/@w:hRule='exact'">
        <xsl:attribute name="style:row-height">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:value-of select="w:trHeight/@w:val"/>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="style:min-row-height">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:value-of select="w:trHeight/@w:val"/>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:attribute name="style:keep-together">
      <xsl:choose>
        <xsl:when test="w:cantSplit">
          <xsl:text>false</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:text>true</xsl:text>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>

</xsl:stylesheet>

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
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">
  
  <xsl:import href="measures.xsl"/>
  <xsl:import href="pixel-measure.xsl"/>
  <xsl:key name="style" match="style:style" use="@style:name"/>
  
  <!-- table is converted into sheet -->
  <xsl:template match="table:table" mode="sheet">
    <xsl:param name="cellNumber"/>
    <xsl:param name="sheetId"/>
    <pzip:entry pzip:target="{concat(concat('xl/worksheets/sheet',$sheetId),'.xml')}">
      <xsl:call-template name="InsertWorksheet">
        <xsl:with-param name="cellNumber" select="$cellNumber"/>
      </xsl:call-template>
    </pzip:entry>
    
    <!-- convert next table -->
      <xsl:apply-templates select="following-sibling::table:table[1]" mode="sheet">
        <xsl:with-param name="cellNumber">
          <xsl:value-of select="$cellNumber + count(table:table-row/table:table-cell[text:p])"/>
        </xsl:with-param>
        <xsl:with-param name="sheetId">
          <xsl:value-of select="$sheetId + 1"/>
        </xsl:with-param>
      </xsl:apply-templates>
    
  </xsl:template>
  
  <!-- insert sheet -->
  <xsl:template name="InsertWorksheet">
    <xsl:param name="cellNumber"/>
    <worksheet>
      <xsl:if test="table:table-column">
        <cols>
          
          <!-- insert first column -->
          <xsl:apply-templates select="table:table-column[1]" mode="sheet">
            <xsl:with-param name="colNumber">1</xsl:with-param>
          </xsl:apply-templates>
          
        </cols>
      </xsl:if>
      <sheetData>
        
        <!-- insert first row -->
        <xsl:apply-templates select="table:table-row[1]" mode="sheet">
          <xsl:with-param name="rowNumber">1</xsl:with-param>
          <xsl:with-param name="cellNumber" select="$cellNumber"/>
        </xsl:apply-templates>
        
      </sheetData>
    </worksheet>
  </xsl:template>
  
  <!-- insert column properties into sheet -->
  <xsl:template match="table:table-column" mode="sheet">
    <xsl:param name="colNumber"/>
    <col min="{$colNumber}">
      <xsl:attribute name="max">
        <xsl:choose>
          <xsl:when test="@table:number-columns-repeated">
            <xsl:value-of select="$colNumber+@table:number-columns-repeated - 1"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$colNumber"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      
      <!-- insert column width -->
      <xsl:attribute name="width">
        <xsl:variable name="pixelWidth">
          <xsl:call-template name="pixel-measure">
            <xsl:with-param name="length">
          <xsl:value-of select="key('style',@table:style-name)/style:table-column-properties/@style:column-width"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="fontSize">
          <xsl:call-template name="pixel-measure">
            <xsl:with-param name="length">11pt</xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:value-of select="($pixelWidth+5) div (2 div 3 * $fontSize)"/>
      </xsl:attribute>
      <xsl:attribute name="customWidth">1</xsl:attribute>
      
      <xsl:if test="@table:visibility = 'collapse'">
        <xsl:attribute name="hidden">1</xsl:attribute>
      </xsl:if>
    </col>    
    
    <!-- insert next column -->
    <xsl:if test="following-sibling::table:table-column">
      <xsl:apply-templates select="following-sibling::table:table-column[1]" mode="sheet">
        <xsl:with-param name="colNumber">
          <xsl:choose>
            <xsl:when test="@table:number-columns-repeated">
              <xsl:value-of select="$colNumber+@table:number-columns-repeated"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$colNumber+1"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
      </xsl:apply-templates>
    </xsl:if>
    
  </xsl:template>
  
  <!-- insert row into sheet -->
  <xsl:template match="table:table-row" mode="sheet">
    <xsl:param name="rowNumber"/>
    <xsl:param name="cellNumber"/>
    <row r="{$rowNumber}">
      
      <!-- insert row height -->
      <xsl:attribute name="ht">
        <xsl:call-template name="ConvertMeasure">
          <xsl:with-param name="length">
            <xsl:value-of select="key('style',@table:style-name)/style:table-row-properties/@style:row-height"/>
          </xsl:with-param>
          <xsl:with-param name="unit">point</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="customHeight">1</xsl:attribute>
      
      <xsl:if test="@table:visibility = 'collapse' or @table:visibility = 'filter'">
        <xsl:attribute name="hidden">1</xsl:attribute>
      </xsl:if>
      <!-- insert first cell -->
      <xsl:apply-templates select="table:table-cell[1]" mode="sheet">
        <xsl:with-param name="colNumber">0</xsl:with-param>
        <xsl:with-param name="rowNumber" select="$rowNumber"/>
        <xsl:with-param name="cellNumber" select="$cellNumber"/>
      </xsl:apply-templates>
      
    </row>
    
    <!-- insert next row -->
    <xsl:if test="following-sibling::table:table-row">
      <xsl:apply-templates select="following-sibling::table:table-row[1]" mode="sheet">
        <xsl:with-param name="rowNumber">
          <xsl:choose>
            <xsl:when test="@table:number-rows-repeated">
              <xsl:value-of select="$rowNumber+@table:number-rows-repeated"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$rowNumber+1"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
        <xsl:with-param name="cellNumber">
          <xsl:value-of select="$cellNumber + count(child::table:table-cell[text:p])"/>
        </xsl:with-param>
      </xsl:apply-templates>
    </xsl:if>
    
  </xsl:template>
  
  <!-- insert cell into row -->
  <xsl:template match="table:table-cell" mode="sheet">
    <xsl:param name="colNumber"/>
    <xsl:param name="rowNumber"/>
    <xsl:param name="cellNumber"/>
    <xsl:if test="child::text:p">
      <c>
        <xsl:attribute name="r">
          <xsl:variable name="colChar">
            <xsl:call-template name="NumbersToChars">
              <xsl:with-param name="num" select="$colNumber"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:value-of select="concat($colChar,$rowNumber)"/>
        </xsl:attribute>
        <xsl:attribute name="t">s</xsl:attribute>
        <v>
          <xsl:value-of select="number($cellNumber)"/>
        </v>
      </c>
    </xsl:if>
    
    <!-- insert next cell --> 
    <xsl:if test="following-sibling::table:table-cell">
      <xsl:apply-templates select="following-sibling::table:table-cell[1]" mode="sheet">
        <xsl:with-param name="colNumber">
          <xsl:choose>
            <xsl:when test="@table:number-columns-repeated">
              <xsl:value-of select="$colNumber+@table:number-columns-repeated"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$colNumber+1"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
        <xsl:with-param name="rowNumber" select="$rowNumber"/>
        <xsl:with-param name="cellNumber">
          <xsl:choose>
            <xsl:when test="child::text:p">
              <xsl:value-of select="$cellNumber + 1"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$cellNumber"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
      </xsl:apply-templates>
    </xsl:if>
    
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

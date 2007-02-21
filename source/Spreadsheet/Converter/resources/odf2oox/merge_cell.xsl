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
  
  <!-- Insert Merge Cell -->
  <xsl:template name="CheckMergeCell">
    <xsl:if test="table:table-row/table:table-cell[@table:number-columns-spanned]">
      <mergeCells>
        <xsl:variable name="rowNumber">
          <xsl:choose>
            <xsl:when test="table:table-row[1]/@table:number-rows-repeated">
              <xsl:value-of select="table:table-row[1]/@table:number-rows-repeated"/>
            </xsl:when>
            <xsl:otherwise>1</xsl:otherwise>
          </xsl:choose>  
        </xsl:variable>
        <xsl:apply-templates select="table:table-row[1]" mode="merge">
          <xsl:with-param name="rowNumber">
            <xsl:value-of select="$rowNumber"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </mergeCells>
    </xsl:if>
  </xsl:template>
  
  
  <xsl:template match="table:table-row" mode="merge">
    <xsl:param name="rowNumber"/>
    
    <!-- Check if atribute table:number-columns-spanned exist -->
    <xsl:if test="table:table-cell[@table:number-columns-spanned]">        
      <xsl:for-each select="table:table-cell[1]">
        <xsl:call-template name="MergeRow">
          <xsl:with-param name="rowNumber">
            <xsl:value-of select="$rowNumber"/>
          </xsl:with-param>
          <xsl:with-param name="colNumber">1</xsl:with-param>
        </xsl:call-template>
      </xsl:for-each>
    </xsl:if>
    
    <!-- Check if next row exist -->
    <xsl:if test="following-sibling::table:table-row">
      <xsl:apply-templates select="following-sibling::table:table-row[1]" mode="merge">
        <xsl:with-param name="rowNumber">
          <xsl:choose>
            <xsl:when test="following-sibling::table:table-row[1]/@table:number-rows-repeated">
              <xsl:value-of select="$rowNumber + following-sibling::table:table-row[1]/@table:number-rows-repeated"/>    
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$rowNumber + 1"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
      </xsl:apply-templates>
    </xsl:if>
  </xsl:template>
  
  
  <!-- to insert merge cell -->
  <xsl:template name="MergeRow">
    <xsl:param name="rowNumber"/>
    <xsl:param name="colNumber"/>
    
    
    <!-- Merge Cell-->
    
    <xsl:choose>
      <xsl:when test="@table:number-columns-spanned">
        <mergeCell>
          
          <xsl:variable name="CollStartChar">
            <xsl:call-template name="NumbersToChars">
              <xsl:with-param name="num">
                <xsl:value-of select="$colNumber - 1"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          
          <xsl:variable name="CollEndChar">
            <xsl:call-template name="NumbersToChars">
              <xsl:with-param name="num">
                <xsl:value-of select="@table:number-columns-spanned + $colNumber - 2"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          
          <xsl:attribute name="ref">
            <xsl:value-of select="concat(concat(concat($CollStartChar,$rowNumber),':'), concat($CollEndChar, $rowNumber + @table:number-rows-spanned - 1))"/>
          </xsl:attribute>
          
        </mergeCell>
      </xsl:when>
      <xsl:otherwise/>
    </xsl:choose>
    
    <!-- Check if next cell exist -->
    
    <xsl:if test="following-sibling::table:table-cell">
      
      <xsl:variable name="CoveredTable">
        <xsl:choose>
          <xsl:when test="following-sibling::node()[1][name()='table:covered-table-cell']/@table:number-columns-repeated">
            <xsl:value-of select="following-sibling::table:covered-table-cell[1]/@table:number-columns-repeated"/>
          </xsl:when>
          <xsl:when test="following-sibling::node()[1][name()='table:covered-table-cell']">1</xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>        
      </xsl:variable>
      
      <xsl:variable name="ColumnsRepeated">
        <xsl:choose>          
          <xsl:when test="@table:number-columns-spanned">
            <xsl:value-of select="$colNumber + @table:number-columns-spanned"/>
          </xsl:when>
          <xsl:when test="@table:number-columns-repeated">
            <xsl:value-of select="$colNumber + @table:number-columns-repeated + $CoveredTable"/>
          </xsl:when>        
          <xsl:otherwise>
            <xsl:value-of select="$colNumber + 1 + $CoveredTable"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      
      <xsl:for-each select="following-sibling::table:table-cell[1]">
        <xsl:call-template name="MergeRow">
          <xsl:with-param name="rowNumber">
            <xsl:value-of select="$rowNumber"/>
          </xsl:with-param>
          <xsl:with-param name="colNumber">
            <xsl:value-of select="$ColumnsRepeated"/>            
          </xsl:with-param>
        </xsl:call-template>
      </xsl:for-each>
      
    </xsl:if>
  </xsl:template>
  
</xsl:stylesheet>

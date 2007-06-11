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
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:xlink="http://www.w3.org/1999/xlink"
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main" exclude-result-prefixes="e r">
    
    <!-- Get cell with note -->
    
    <xsl:template name="NoteCell">
        <xsl:param name="sheetNr"/>
        <xsl:apply-templates select="document(concat('xl/comments',$sheetNr,'.xml'))/e:comments" mode="note-cell"/>
    </xsl:template>
    
    <xsl:template match="e:comments" mode="note-cell">
        <xsl:apply-templates select="e:commentList/e:comment" mode="note-cell"/>
    </xsl:template>
    
    <xsl:template match="e:comment" mode="note-cell">
        
        <xsl:variable name="numCol">
            <xsl:call-template name="GetColNum">
                <xsl:with-param name="cell">
                    <xsl:value-of select="@ref"/>
                </xsl:with-param>
            </xsl:call-template>
        </xsl:variable>
        
        <xsl:variable name="numRow">
            <xsl:call-template name="GetRowNum">
                <xsl:with-param name="cell">
                    <xsl:value-of select="@ref"/>
                </xsl:with-param>
            </xsl:call-template>
        </xsl:variable>
        <xsl:value-of select="concat($numRow,':',$numCol,';')"/>
    </xsl:template>
   
    <!-- Get Row with Note -->
   <xsl:template name="NoteRow">
    <xsl:param name="NoteCell"/>
    <xsl:param name="Result"/>
    <xsl:choose>
      <xsl:when test="$NoteCell != ''">
        <xsl:call-template name="NoteRow">
          <xsl:with-param name="NoteCell">
            <xsl:value-of select="substring-after($NoteCell, ';')"/>
          </xsl:with-param>
          <xsl:with-param name="Result">
            <xsl:value-of
              select="concat($Result,  concat(substring-before($NoteCell, ':'), ';'))"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$Result"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- Insert all picture betwen two rows -->
  <xsl:template name="InsertPictureAndNoteBetwenTwoRows">
    <xsl:param name="StartRow"/>
    <xsl:param name="EndRow"/>
    <xsl:param name="PictureRow"/>
    <xsl:param name="PictureCell"/>
    <xsl:param name="NoteCell"></xsl:param>
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    
    <xsl:variable name="GetMinRowWithPicture">
      <xsl:call-template name="GetMinRowWithPicture">
        <xsl:with-param name="PictureOrNoteRow">
          <xsl:value-of select="$PictureRow"/>
        </xsl:with-param>
        <xsl:with-param name="AfterRow">
          <xsl:value-of select="$StartRow"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    
    <xsl:choose>
      <!-- Insert empty rows before -->
      <xsl:when
        test="$GetMinRowWithPicture != '' and $GetMinRowWithPicture &gt;= $StartRow and $GetMinRowWithPicture &lt; $EndRow">
        <xsl:if test="$GetMinRowWithPicture - $StartRow &gt; 0">
          <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
            table:number-rows-repeated="{$GetMinRowWithPicture - $StartRow}">
            <table:table-cell table:number-columns-repeated="256"/>
          </table:table-row>
        </xsl:if>
        <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}">
          
          <xsl:variable name="PictureColl">
            <xsl:call-template name="GetCollsWithPicture">
              <xsl:with-param name="rowNumber">
                <xsl:value-of select="$GetMinRowWithPicture"/>
              </xsl:with-param>
              <xsl:with-param name="PictureCell">
                <xsl:value-of select="concat(';', $PictureCell)"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          
          <xsl:call-template name="InsertPictureAndNoteBetwenTwoColl">
            <xsl:with-param name="sheet">
              <xsl:value-of select="$sheet"/>
            </xsl:with-param>
            <xsl:with-param name="NameSheet">
              <xsl:value-of select="$NameSheet"/>
            </xsl:with-param>
            <xsl:with-param name="rowNum">
              <xsl:value-of select="$GetMinRowWithPicture"/>
            </xsl:with-param>
            <xsl:with-param name="PictureColl">
              <xsl:value-of select="$PictureColl"/>
            </xsl:with-param>
            <xsl:with-param name="StartColl">
              <xsl:text>0</xsl:text>
            </xsl:with-param>
            <xsl:with-param name="EndColl">
              
              <xsl:text>256</xsl:text>
            </xsl:with-param>
          </xsl:call-template>
          
        </table:table-row>
        
        <xsl:call-template name="InsertPictureAndNoteBetwenTwoRows">
          <xsl:with-param name="StartRow">
            <xsl:value-of select="$GetMinRowWithPicture + 1"/>
          </xsl:with-param>
          <xsl:with-param name="EndRow">
            <xsl:value-of select="$EndRow"/>
          </xsl:with-param>
          <xsl:with-param name="PictureRow">
            <xsl:value-of select="$PictureRow"/>
          </xsl:with-param>
          <xsl:with-param name="PictureCell">
            <xsl:value-of select="$PictureCell"/>
          </xsl:with-param>
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="NameSheet">
            <xsl:value-of select="$NameSheet"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      
      <xsl:otherwise>
        <xsl:if test="$EndRow - $StartRow - 1 &gt; 0">
          <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
            table:number-rows-repeated="{$EndRow - $StartRow - 1}">
            <table:table-cell table:number-columns-repeated="256"/>
          </table:table-row>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
    
  </xsl:template>
  
  <!-- Insert all picture betwen two cell -->
  <xsl:template name="InsertPictureAndNoteBetwenTwoColl">
    <xsl:param name="StartColl"/>
    <xsl:param name="EndColl"/>
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="rowNum"/>
    <xsl:param name="PictureColl"/>
    <xsl:param name="document"/>
    <xsl:param name="sheetNr"/>
    
    <xsl:variable name="GetMinCollWithPicture">
      <xsl:call-template name="GetMinRowWithPicture">
        <xsl:with-param name="PictureRow">
          <xsl:value-of select="$PictureColl"/>
        </xsl:with-param>
        <xsl:with-param name="AfterRow">
          <xsl:value-of select="$StartColl - 1"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    
    
    <xsl:choose>
      <!-- Insert empty rows before -->
      <xsl:when
        test="$GetMinCollWithPicture != '' and $GetMinCollWithPicture &gt;= $StartColl and $GetMinCollWithPicture &lt; $EndColl">
        
        <xsl:if test="$GetMinCollWithPicture - $StartColl &gt; 0">
          <table:table-cell table:number-columns-repeated="{$GetMinCollWithPicture - $StartColl}"/>
        </xsl:if>
        
        <xsl:for-each select="ancestor::e:worksheet/e:drawing">
          <xsl:variable name="thisCellCol">
            <xsl:call-template name="NumbersToChars">
              <xsl:with-param name="num">
                <xsl:value-of select="$PictureColl -1"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="thisCell">
            <xsl:value-of select="concat($thisCellCol,$rowNum)"/>
          </xsl:variable>
          <xsl:apply-templates select="document(concat('xl/comments',$sheetNr,'.xml'))/e:comments/e:commentList/e:comment[@ref=$thisCell]">
            <xsl:with-param name="number" select="$sheetNr"/>
          </xsl:apply-templates>
          <xsl:variable name="Target">
            <xsl:call-template name="GetTargetPicture">
              <xsl:with-param name="sheet">
                <xsl:value-of select="substring-after($sheet, '/')"/>
              </xsl:with-param>
              <xsl:with-param name="id">
                <xsl:value-of select="@r:id"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          
          <table:table-cell>
            
            
            <xsl:call-template name="InsertPictureInThisCell">
              <xsl:with-param name="sheet">
                <xsl:value-of select="$sheet"/>
              </xsl:with-param>
              <xsl:with-param name="NameSheet">
                <xsl:value-of select="$NameSheet"/>
              </xsl:with-param>
              <xsl:with-param name="collNum">
                <xsl:value-of select="$GetMinCollWithPicture"/>
              </xsl:with-param>
              <xsl:with-param name="rowNum">
                <xsl:value-of select="$rowNum"/>
              </xsl:with-param>
              <xsl:with-param name="Target">
                <xsl:value-of select="$Target"/>
              </xsl:with-param>
            </xsl:call-template>
          </table:table-cell>
          
        </xsl:for-each>
        
        <xsl:call-template name="InsertPictureAndNoteBetwenTwoColl">
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="NameSheet">
            <xsl:value-of select="$NameSheet"/>
          </xsl:with-param>
          <xsl:with-param name="rowNum">
            <xsl:value-of select="$rowNum"/>
          </xsl:with-param>
          <xsl:with-param name="PictureColl">
            <xsl:value-of select="$PictureColl"/>
          </xsl:with-param>
          <xsl:with-param name="StartColl">
            <xsl:choose>
              <xsl:when test="$document = 'worksheet'">
                <xsl:value-of select="$GetMinCollWithPicture + 1"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$GetMinCollWithPicture + 2"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="EndColl">
            <xsl:value-of select="$EndColl"/>
          </xsl:with-param>
          <xsl:with-param name="document">
            <xsl:value-of select="$document"/>
          </xsl:with-param>
        </xsl:call-template>
        
      </xsl:when>
      
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$StartColl = 0">
            <table:table-cell table:number-columns-repeated="{$EndColl - $StartColl}"/>
          </xsl:when>
          <xsl:otherwise>
            <table:table-cell table:number-columns-repeated="{$EndColl - $StartColl - 1}"/>
          </xsl:otherwise>
          
        </xsl:choose>
        
        
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- Insert all note betwen two rows -->
  <xsl:template name="InsertNoteBetwenTwoRows">
    <xsl:param name="StartRow"/>
    <xsl:param name="EndRow"/>
    <xsl:param name="NoteRow"/>
    <xsl:param name="NoteCell"/>
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="sheetNr"/>
    
    <xsl:variable name="GetMinRowWithNote">
      <xsl:call-template name="GetMinRowWithPicture">
        <xsl:with-param name="PictureOrNoteRow">
          <xsl:value-of select="$NoteRow"/>
        </xsl:with-param>
        <xsl:with-param name="AfterRow">
          <xsl:value-of select="$StartRow"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    
    <xsl:choose>
      <!-- Insert empty rows before -->
      <xsl:when
        test="$GetMinRowWithNote != '' and $GetMinRowWithNote &gt;= $StartRow and $GetMinRowWithNote &lt; $EndRow">
        <xsl:if test="$GetMinRowWithNote - $StartRow &gt; 0">
          <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
            table:number-rows-repeated="{$GetMinRowWithNote - $StartRow}">
            <table:table-cell table:number-columns-repeated="256"/>
          </table:table-row>
        </xsl:if>
        <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}">
          
          <xsl:variable name="NoteColl">
            <xsl:call-template name="GetCollsWithPicture">
              <xsl:with-param name="rowNumber">
                <xsl:value-of select="$GetMinRowWithNote"/>
              </xsl:with-param>
              <xsl:with-param name="PictureCell">
                <xsl:value-of select="concat(';', $NoteCell)"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          
          <xsl:call-template name="InsertNoteBetwenTwoColl">
            <xsl:with-param name="sheet">
              <xsl:value-of select="$sheet"/>
            </xsl:with-param>
            <xsl:with-param name="NameSheet">
              <xsl:value-of select="$NameSheet"/>
            </xsl:with-param>
            <xsl:with-param name="rowNum">
              <xsl:value-of select="$GetMinRowWithNote"/>
            </xsl:with-param>
            <xsl:with-param name="NoteColl">
              <xsl:value-of select="$NoteColl"/>
            </xsl:with-param>
            <xsl:with-param name="StartColl">
              <xsl:text>0</xsl:text>
            </xsl:with-param>
            <xsl:with-param name="EndColl">
              <xsl:text>256</xsl:text>
            </xsl:with-param>
            <xsl:with-param name="sheetNr" select="$sheetNr"/>
          </xsl:call-template>
          
        </table:table-row>
        
        <xsl:call-template name="InsertNoteBetwenTwoRows">
          <xsl:with-param name="StartRow">
            <xsl:value-of select="$GetMinRowWithNote + 1"/>
          </xsl:with-param>
          <xsl:with-param name="EndRow">
            <xsl:value-of select="$EndRow"/>
          </xsl:with-param>
          <xsl:with-param name="NoteRow">
            <xsl:value-of select="$NoteRow"/>
          </xsl:with-param>
          <xsl:with-param name="NoteCell">
            <xsl:value-of select="$NoteCell"/>
          </xsl:with-param>
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="NameSheet">
            <xsl:value-of select="$NameSheet"/>
          </xsl:with-param>
          <xsl:with-param name="sheetNr" select="$sheetNr"/>
        </xsl:call-template>
      </xsl:when>
      
      <xsl:otherwise>
        <xsl:if test="$EndRow - $StartRow - 1 &gt; 0">
          <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
            table:number-rows-repeated="{$EndRow - $StartRow - 1}">
            <table:table-cell table:number-columns-repeated="256"/>
          </table:table-row>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
    
  </xsl:template>
  
 
  
  <!-- Insert all picture betwen two cell -->
  <xsl:template name="InsertNoteBetwenTwoColl">
    <xsl:param name="StartColl"/>
    <xsl:param name="EndColl"/>
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="rowNum"/>
    <xsl:param name="NoteColl"/>
    <xsl:param name="document"/>
    <xsl:param name="sheetNr"/>
    
    <xsl:variable name="GetMinCollWithNote">
      <xsl:call-template name="GetMinRowWithPicture">
        <xsl:with-param name="PictureRow">
          <xsl:value-of select="$NoteColl"/>
        </xsl:with-param>
        <xsl:with-param name="AfterRow">
          <xsl:value-of select="$StartColl - 1"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
   
    
    <xsl:choose>
      <!-- Insert empty rows before -->
      <xsl:when
        test="$GetMinCollWithNote != '' and $GetMinCollWithNote &gt;= $StartColl and $GetMinCollWithNote &lt; $EndColl">
        
        <xsl:if test="$GetMinCollWithNote - $StartColl &gt; 0">
          <table:table-cell table:number-columns-repeated="{$GetMinCollWithNote - $StartColl - 1}"/>
        </xsl:if>
        
          
          <table:table-cell>
           <xsl:call-template name="InsertNoteInThisCell">
               <xsl:with-param name="sheetNr">
                   <xsl:value-of select="$sheetNr"/>
               </xsl:with-param>
               <xsl:with-param name="colNum">
                   <xsl:value-of select="$GetMinCollWithNote"/>
               </xsl:with-param>
               <xsl:with-param name="rowNum">
                   <xsl:value-of select="$rowNum"/>
               </xsl:with-param>
           </xsl:call-template>
          </table:table-cell>

        <xsl:call-template name="InsertNoteBetwenTwoColl">
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="NameSheet">
            <xsl:value-of select="$NameSheet"/>
          </xsl:with-param>
          <xsl:with-param name="rowNum">
            <xsl:value-of select="$rowNum"/>
          </xsl:with-param>
          <xsl:with-param name="PictureColl">
            <xsl:value-of select="$NoteColl"/>
          </xsl:with-param>
          <xsl:with-param name="StartColl">
           <xsl:value-of select="$GetMinCollWithNote"/>
          </xsl:with-param>
          <xsl:with-param name="EndColl">
            <xsl:value-of select="$EndColl"/>
          </xsl:with-param>
          <xsl:with-param name="document">
            <xsl:value-of select="$document"/>
          </xsl:with-param>
            <xsl:with-param name="sheetNr">
                <xsl:value-of select="$sheetNr"/>
            </xsl:with-param>
        </xsl:call-template>
        
      </xsl:when>
      
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$StartColl = 0">
            <table:table-cell table:number-columns-repeated="{$EndColl - $StartColl}"/>
          </xsl:when>
          <xsl:otherwise>
            <table:table-cell table:number-columns-repeated="{$EndColl - $StartColl - 1}"/>
          </xsl:otherwise>
          
        </xsl:choose>
        
        
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
    
    <!-- Insert Note in This Cell -->
    <xsl:template name="InsertNoteInThisCell">
        <xsl:param name="rowNum"/>
        <xsl:param name="colNum"/>
        <xsl:param name="sheetNr"/>
        
        <xsl:variable name="thisCellCol">
            <xsl:call-template name="NumbersToChars">
                <xsl:with-param name="num">
                    <xsl:value-of select="$colNum -1"/>
                </xsl:with-param>
            </xsl:call-template>
        </xsl:variable>
        
        <xsl:variable name="thisCell">
            <xsl:value-of select="concat($thisCellCol,$rowNum)"/>
        </xsl:variable>
        
        <xsl:apply-templates select="document(concat('xl/comments',$sheetNr,'.xml'))/e:comments/e:commentList/e:comment[@ref=$thisCell]">
            <xsl:with-param name="number" select="$sheetNr"/>
        </xsl:apply-templates>
    
    </xsl:template>
    
    
    <!-- Insert Comment -->
    <xsl:template match="e:comment">
        
        <!--@Description: adds a note -->
        <!--@context: none -->
        
        <xsl:param name="number"/>
        
        <!--(int) number of comments file -->
        <xsl:variable name="numberOfComment">
            <xsl:value-of select="count(preceding-sibling::e:comment)+1"/>
        </xsl:variable>
        
        <office:annotation>
            <xsl:apply-templates select="document(concat('xl/drawings/vmlDrawing',$number,'.vml'))/xml/v:shape[position()=$numberOfComment]" mode="drawing">
                <xsl:with-param name="text" select="e:text"/>
            </xsl:apply-templates>
            <text:p text:style-name="{generate-id(e:text)}">
                <xsl:apply-templates select="e:text/e:r"/>
            </text:p>
        </office:annotation>
    </xsl:template>
    
</xsl:stylesheet>

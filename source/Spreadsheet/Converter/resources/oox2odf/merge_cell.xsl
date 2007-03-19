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
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main" exclude-result-prefixes="e r">


  <xsl:template name="CheckIfMerge">
    <xsl:param name="colNum"/>
    <xsl:param name="rowNum"/>

    <xsl:for-each select="ancestor::e:worksheet">

<!-- Checks if Merge Cells exist  in the document-->
      <xsl:choose>
        <xsl:when test="e:mergeCells/e:mergeCell">
          <xsl:apply-templates select="e:mergeCells/e:mergeCell[1]" mode="merge">
            <xsl:with-param name="colNum">
              <xsl:value-of select="$colNum"/>
            </xsl:with-param>
            <xsl:with-param name="rowNum">
              <xsl:value-of select="$rowNum"/>
            </xsl:with-param>
          </xsl:apply-templates>
        </xsl:when>
        <xsl:otherwise>
          <xsl:text>false</xsl:text>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>
  
  <xsl:template name="CheckIfBigMerge">
    <xsl:param name="colNum"/>
    <xsl:param name="BigMergeCell"/>
    <xsl:param name="Result"/>
    
    
    <xsl:choose>
      <xsl:when test="$BigMergeCell != ''">
        
        <!-- Checks if Merge Cells exist  in the document-->
        <xsl:variable name="StartBigMerge">
          <xsl:value-of select="substring-before($BigMergeCell, ':')"/>
        </xsl:variable>
        <xsl:variable name="EndBigMerge">
          <xsl:value-of select="substring-after(substring-before($BigMergeCell, ';'), ':')"/>
        </xsl:variable>
        
        <xsl:choose>
          <xsl:when test="$colNum = $StartBigMerge ">
            <xsl:value-of select="concat(concat($StartBigMerge, ':'), $EndBigMerge)"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="CheckIfBigMerge">
              <xsl:with-param name="colNum">
                <xsl:value-of select="$colNum"/>
              </xsl:with-param>
              <xsl:with-param name="BigMergeCell">
                <xsl:value-of select="substring-after($BigMergeCell, ';')"/>
              </xsl:with-param>
              <xsl:with-param name="Result">
                <xsl:choose>
                  <xsl:when test="$Result = '' and $StartBigMerge &gt; $colNum">
                    <xsl:value-of select="concat(concat($StartBigMerge, ':'), $EndBigMerge)"/>                    
                  </xsl:when>
                  <xsl:when test="number(substring-before($Result, ':')) &gt; $StartBigMerge and $StartBigMerge &gt; $colNum">
                    <xsl:value-of select="concat(concat($StartBigMerge, ':'), $EndBigMerge)"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="$Result"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>        
        <xsl:value-of select="$Result"/>        
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="CheckIfBigMergeRow">
    <xsl:param name="RowNum"/>
    <xsl:param name="BigMergeCellRow"/>
    <xsl:param name="Result"/>
    
    
    <xsl:choose>
      <xsl:when test="$BigMergeCellRow != ''">
        
        <!-- Checks if Merge Cells exist  in the document-->
        <xsl:variable name="StartBigMerge">
          <xsl:value-of select="substring-before($BigMergeCellRow, ':')"/>
        </xsl:variable>
        <xsl:variable name="EndBigMerge">
          <xsl:value-of select="substring-after(substring-before($BigMergeCellRow, ';'), ':')"/>
        </xsl:variable>
        
        <xsl:choose>
          <xsl:when test="$RowNum = $StartBigMerge ">
            <xsl:value-of select="concat(concat($StartBigMerge, ':'), $EndBigMerge)"/>
          </xsl:when>         
          <xsl:otherwise>
            <xsl:call-template name="CheckIfBigMergeRow">
                <xsl:with-param name="RowNum">
                  <xsl:value-of select="$RowNum"/>
                </xsl:with-param>
              <xsl:with-param name="BigMergeCellRow">
                <xsl:value-of select="substring-after($BigMergeCellRow, ';')"/>
              </xsl:with-param>
              <xsl:with-param name="Result">
                <xsl:choose>
                  <xsl:when test="$EndBigMerge &gt; $RowNum and $RowNum &gt; $StartBigMerge">
                    <xsl:text>true</xsl:text>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="$Result"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>        
      </xsl:when>      
      <xsl:otherwise>        
        <xsl:value-of select="$Result"/>        
      </xsl:otherwise>
    </xsl:choose>
    
  </xsl:template>

  <xsl:template match="e:mergeCell" mode="merge">
    <xsl:param name="colNum"/>
    <xsl:param name="rowNum"/>

    <xsl:variable name="StartMergeCell">
      <xsl:value-of select="substring-before(@ref, ':')"/>
    </xsl:variable>

    <xsl:variable name="EndMergeCell">
      <xsl:value-of select="substring-after(@ref, ':')"/>
    </xsl:variable>

    <xsl:variable name="StartColNum">
      <xsl:call-template name="GetColNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="$StartMergeCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="StartRowNum">
      <xsl:call-template name="GetRowNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="$StartMergeCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="EndColNum">
      <xsl:call-template name="GetColNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="$EndMergeCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="EndRowNum">
      <xsl:call-template name="GetRowNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="$EndMergeCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:choose>

      <!-- Checks if "Merge Cell" is starting in this cell -->
      <xsl:when test="$colNum = $StartColNum and $rowNum = $StartRowNum">
        <xsl:value-of
          select="concat(number($EndRowNum - $StartRowNum + 1), concat(':',number($EndColNum - $StartColNum+1)))"
        />
      </xsl:when>

      <!-- Checks if this cell is in "Merge Cell" -->
      <xsl:when
        test="$colNum = $StartColNum and $EndColNum &gt;= $colNum and $rowNum &gt;= $StartRowNum and $EndRowNum &gt;= $rowNum">
        <xsl:value-of select="concat('true:', number($EndColNum - $StartColNum +1 ))"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="following-sibling::e:mergeCell">
            <xsl:apply-templates select="following-sibling::e:mergeCell[1]" mode="merge">
              <xsl:with-param name="colNum">
                <xsl:value-of select="$colNum"/>
              </xsl:with-param>
              <xsl:with-param name="rowNum">
                <xsl:value-of select="$rowNum"/>
              </xsl:with-param>
            </xsl:apply-templates>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>false</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>
  

  
  <xsl:template match="e:mergeCell" mode="BigMergeColl">
    <xsl:param name="sheet"/>
    <xsl:param name="RefValue"/>
    
    <xsl:variable name="StartMergeCell">
      <xsl:value-of select="substring-before(@ref, ':')"/>
    </xsl:variable>
    
    <xsl:variable name="EndMergeCell">
      <xsl:value-of select="substring-after(@ref, ':')"/>
    </xsl:variable>
    
    <xsl:variable name="StartColNum">
      <xsl:call-template name="GetColNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="$StartMergeCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    
    <xsl:variable name="EndColNum">
      <xsl:call-template name="GetColNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="$EndMergeCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>    
    
    <xsl:variable name="StartRowNum">
      <xsl:call-template name="GetRowNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="$StartMergeCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    
    <xsl:variable name="EndRowNum">
      <xsl:call-template name="GetRowNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="$EndMergeCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    
     <xsl:choose>
      <xsl:when test="following-sibling::e:mergeCell">
        <xsl:apply-templates select="following-sibling::e:mergeCell[1]" mode="BigMerge">
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="RefValue">
            <xsl:choose>
              <xsl:when test="$StartRowNum = 1 and $EndRowNum = 1048576">                
                <xsl:value-of select="concat($RefValue, concat($StartColNum, concat(':', concat($EndColNum, ';'))))"/>    
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$RefValue"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$StartRowNum = 1 and $EndRowNum = 1048576">
            <xsl:value-of select="concat($RefValue, concat($StartColNum, concat(':', concat($EndColNum, ';'))))"/>
          </xsl:when>
          <xsl:otherwise>            
            <xsl:value-of select="$RefValue"/>
          </xsl:otherwise>
        </xsl:choose>        
      </xsl:otherwise>      
    </xsl:choose>
  </xsl:template>
  
  
  <xsl:template match="e:mergeCell" mode="BigMergeRow">
    <xsl:param name="sheet"/>
    <xsl:param name="RefValue"/>
    
    
    <xsl:variable name="StartMergeCell">
      <xsl:value-of select="substring-before(@ref, ':')"/>
    </xsl:variable>
    
    <xsl:variable name="EndMergeCell">
      <xsl:value-of select="substring-after(@ref, ':')"/>
    </xsl:variable>
    
    <xsl:variable name="StartColNum">
      <xsl:call-template name="GetColNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="$StartMergeCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    
    <xsl:variable name="EndColNum">
      <xsl:call-template name="GetColNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="$EndMergeCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>    
    
    <xsl:variable name="StartRowNum">
      <xsl:call-template name="GetRowNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="$StartMergeCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    
    <xsl:variable name="EndRowNum">
      <xsl:call-template name="GetRowNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="$EndMergeCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    
    <xsl:choose>
      <xsl:when test="following-sibling::e:mergeCell">
        <xsl:apply-templates select="following-sibling::e:mergeCell[1]" mode="BigMergeRow">
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="RefValue">
            <xsl:choose>
              <xsl:when test="$StartColNum = 1 and $EndColNum = 16384">                
                <xsl:value-of select="concat($RefValue, concat($StartRowNum, concat(':', concat($EndRowNum, ';'))))"/>    
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$RefValue"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>        
        </xsl:apply-templates>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$StartColNum = 1 and $EndColNum = 16384">
            <xsl:value-of select="concat($RefValue, concat($StartRowNum, concat(':', concat($EndRowNum, ';'))))"/>
          </xsl:when>
          <xsl:otherwise>            
            <xsl:value-of select="$RefValue"/>
          </xsl:otherwise>
        </xsl:choose>        
      </xsl:otherwise>      
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="InsertColumnsBigMergeRow">
    <xsl:param name="RowNumber"/>
    <xsl:param name="Repeat"/>
    <xsl:param name="BigMergeCell"/>
    
    <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}">     
      <xsl:call-template name="InsertColumnsBigMergeColl">
        <xsl:with-param name="BigMergeCell">
          <xsl:value-of select="$BigMergeCell"/>
        </xsl:with-param>
        <xsl:with-param name="RowNumber">
          <xsl:value-of select="$RowNumber"/>
        </xsl:with-param>
      </xsl:call-template>
    </table:table-row>
    
    <xsl:if test="$Repeat &gt; $RowNumber">
      <xsl:call-template name="InsertColumnsBigMergeRow">
        <xsl:with-param name="RowNumber">
          <xsl:value-of select="$RowNumber + 1"/>
        </xsl:with-param>
        <xsl:with-param name="Repeat">
          <xsl:value-of select="$Repeat"/>
        </xsl:with-param>     
        <xsl:with-param name="BigMergeCell">
          <xsl:value-of select="$BigMergeCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if>    
  </xsl:template>
  
  <xsl:template name="InsertColumnsBigMergeColl">
    <xsl:param name="CollNumber" select="1"/> 
    <xsl:param name="BigMergeCell"/>
    <xsl:param name="RowNumber"/>
    
    
    <xsl:if test="$CollNumber &lt; 257">
      <xsl:choose>
        <xsl:when test="$BigMergeCell != ''">      
          
          <xsl:variable name="MergeCell">            
                <xsl:call-template name="CheckIfBigMerge">
                  <xsl:with-param name="colNum">
                    <xsl:value-of select="$CollNumber"/>
                  </xsl:with-param>
                  <xsl:with-param name="BigMergeCell">
                    <xsl:value-of select="$BigMergeCell"/>
                  </xsl:with-param>
                </xsl:call-template>
          </xsl:variable>
          
          <xsl:choose>     
            <xsl:when test="number(substring-before($MergeCell, ':')) &gt; $CollNumber">        
              <table:table-cell>          
                <xsl:attribute name="table:number-columns-repeated">                  
                  <xsl:choose>
                    <xsl:when test="number(substring-before($MergeCell, ':')) &gt; 256">
                      <xsl:value-of select="256 - number($CollNumber)"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="number(substring-before($MergeCell, ':')) - number($CollNumber)"/>    
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
              </table:table-cell>
              <xsl:call-template name="InsertColumnsBigMergeColl">
                <xsl:with-param name="CollNumber">
                  <xsl:value-of select="number(substring-before($MergeCell, ':'))"/>
                </xsl:with-param>
                <xsl:with-param name="BigMergeCell">
                  <xsl:value-of select="$BigMergeCell"/>
                </xsl:with-param>
                <xsl:with-param name="RowNumber">
                  <xsl:value-of select="$RowNumber"/>
                </xsl:with-param>          
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when test="$RowNumber = '1'">
                  <xsl:variable name="ColumnsSpanned">
                    <xsl:value-of select="number(substring-after($MergeCell, ':')) - number(substring-before($MergeCell, ':')) + 1"/>              
                  </xsl:variable>
                  <xsl:if test="$ColumnsSpanned != 'NaN'">            
                    <table:table-cell>
                      <xsl:attribute name="table:number-rows-spanned">
                        <xsl:value-of select="65536"/>
                      </xsl:attribute>
                      <xsl:attribute name="table:number-columns-spanned">
                        <xsl:choose>
                          <xsl:when test="number(substring-after($MergeCell, ':')) &gt; 256">
                            <xsl:value-of select="256 - number(substring-before($MergeCell, ':')) + 1"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="number(substring-after($MergeCell, ':')) - number(substring-before($MergeCell, ':')) + 1"/> 
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:attribute>
                    </table:table-cell>
                  </xsl:if>
                  <xsl:choose>
                    <xsl:when test="number(substring-after($MergeCell, ':')) - number(substring-before($MergeCell, ':')) = 1">
                      <table:covered-table-cell/>   
                      <xsl:call-template name="InsertColumnsBigMergeColl">
                        <xsl:with-param name="CollNumber">          
                          <xsl:value-of select="$CollNumber + 2"/>
                        </xsl:with-param>
                        <xsl:with-param name="BigMergeCell">
                          <xsl:value-of select="concat(substring-before($BigMergeCell, $MergeCell), substring-after($BigMergeCell, concat($MergeCell, ';')))"/>
                        </xsl:with-param>
                        <xsl:with-param name="RowNumber">
                          <xsl:value-of select="$RowNumber"/>
                        </xsl:with-param>
                      </xsl:call-template>                
                    </xsl:when>
                    <xsl:when test="number(substring-after($MergeCell, ':')) - number(substring-before($MergeCell, ':')) &gt; 1">
                      <table:covered-table-cell>        
                        <xsl:choose>
                          <xsl:when test="number(substring-after($MergeCell, ':')) &gt; number(substring-before($MergeCell, ':'))">
                            <xsl:attribute name="table:number-columns-repeated">
                              <xsl:choose>
                                <xsl:when test="number(substring-after($MergeCell, ':')) &gt; 256">
                                  <xsl:value-of select="256 - number(substring-before($MergeCell, ':'))"/>    
                                </xsl:when>
                                <xsl:otherwise>
                                  <xsl:value-of select="number(substring-after($MergeCell, ':')) - number(substring-before($MergeCell, ':'))"/>    
                                </xsl:otherwise>
                              </xsl:choose>
                            </xsl:attribute>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:call-template name="InsertColumnsBigMergeColl">
                              <xsl:with-param name="CollNumber">          
                                <xsl:value-of select="$CollNumber + 1"/>
                              </xsl:with-param>
                              <xsl:with-param name="BigMergeCell">
                                <xsl:value-of select="concat(substring-before($BigMergeCell, $MergeCell), substring-after($BigMergeCell, concat($MergeCell, ';')))"/>
                              </xsl:with-param>
                              <xsl:with-param name="RowNumber">
                                <xsl:value-of select="$RowNumber"/>
                              </xsl:with-param>
                            </xsl:call-template>
                          </xsl:otherwise>
                        </xsl:choose>                  
                      </table:covered-table-cell>
                      <xsl:call-template name="InsertColumnsBigMergeColl">
                        <xsl:with-param name="CollNumber">          
                          <xsl:value-of select="$CollNumber + number(substring-after($MergeCell, ':')) - number(substring-before($MergeCell, ':')) + 1"/>
                        </xsl:with-param>
                        <xsl:with-param name="BigMergeCell">
                          <xsl:value-of select="concat(substring-before($BigMergeCell, $MergeCell), substring-after($BigMergeCell, concat($MergeCell, ';')))"/>
                        </xsl:with-param>
                        <xsl:with-param name="RowNumber">
                          <xsl:value-of select="$RowNumber"/>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:call-template name="InsertColumnsBigMergeColl">
                        <xsl:with-param name="CollNumber">
                          <xsl:variable name="NextCell">
                            <xsl:value-of select="$CollNumber + number(substring-after($MergeCell, ':')) - number(substring-before($MergeCell, ':')) + 1"/>
                          </xsl:variable>
                          <xsl:choose>
                            <xsl:when test="contains($BigMergeCell, concat($NextCell, ':'))">
                              <xsl:value-of select="$NextCell"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="$CollNumber + number(substring-after($MergeCell, ':')) - number(substring-before($MergeCell, ':')) + 1"/>    
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:with-param>
                        <xsl:with-param name="BigMergeCell">
                          <xsl:value-of select="concat(substring-before($BigMergeCell, $MergeCell), substring-after($BigMergeCell, concat($MergeCell, ';')))"/>
                        </xsl:with-param>
                        <xsl:with-param name="RowNumber">
                          <xsl:value-of select="$RowNumber"/>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:otherwise>
                  <table:covered-table-cell>        
                    <xsl:choose>
                      <xsl:when test="number(substring-after($MergeCell, ':')) &gt; number(substring-before($MergeCell, ':'))">
                        <xsl:attribute name="table:number-columns-repeated">                    
                          <xsl:choose>
                            <xsl:when test="number(substring-after($MergeCell, ':')) &gt; 256">
                              <xsl:value-of select="256 - number(substring-before($MergeCell, ':')) + 1"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="number(substring-after($MergeCell, ':')) - number(substring-before($MergeCell, ':')) + 1"/>    
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
                      </xsl:when>
                      <xsl:otherwise/>
                    </xsl:choose>             
                  </table:covered-table-cell>            
                </xsl:otherwise>
              </xsl:choose>
              
              <xsl:call-template name="InsertColumnsBigMergeColl">
                <xsl:with-param name="CollNumber">          
                  <xsl:value-of select="$CollNumber + number(substring-after($MergeCell, ':')) - number(substring-before($MergeCell, ':')) + 1"/>
                </xsl:with-param>
                <xsl:with-param name="BigMergeCell">
                  <xsl:value-of select="concat(substring-before($BigMergeCell, $MergeCell), substring-after($BigMergeCell, concat($MergeCell, ';')))"/>
                </xsl:with-param>
                <xsl:with-param name="RowNumber">
                  <xsl:value-of select="$RowNumber"/>
                </xsl:with-param>
              </xsl:call-template>
              
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:if test="$CollNumber &lt; 256">
            <table:table-cell>
              <xsl:choose>
                <xsl:when test="256 - $CollNumber &gt; 1">
                  <xsl:attribute name="table:number-columns-repeated">
                    <xsl:value-of select="256 - $CollNumber + 1"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise/>
              </xsl:choose>         
            </table:table-cell>  
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>     
  </xsl:template>

</xsl:stylesheet>

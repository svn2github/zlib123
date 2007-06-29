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
  xmlns:v="urn:schemas-microsoft-com:vml" xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main" exclude-result-prefixes="e r">

  <xsl:import href="relationships.xsl"/>

  <xsl:key name="hyperlinkPosition" match="e:c" use="'@r'"/>
  <xsl:key name="ref" match="e:hyperlink" use="@ref"/>

  <!-- Insert sheet without text -->
  <xsl:template name="InsertEmptySheet">
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="BigMergeCell"/>
    <xsl:param name="BigMergeRow"/>
    <xsl:param name="RowNumber"/>
    <xsl:param name="PictureCell"/>
    <xsl:param name="PictureRow"/>
    <xsl:param name="NoteCell"/>  
    <xsl:param name="NoteRow"/>
    <xsl:param name="ConditionalCell"/>
    <xsl:param name="ConditionalCellStyle"/>
    <xsl:param name="ConditionalRow"/>
    <xsl:param name="sheetNr"/>

    <xsl:choose>
      <!-- when sheet is empty  -->
      <xsl:when
        test="not(e:worksheet/e:sheetData/e:row/e:c/e:v) and $BigMergeCell = '' and $BigMergeRow = '' and $PictureCell = '' and $NoteCell = '' and $ConditionalCell = ''">
        <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
          table:number-rows-repeated="65536">
          <table:table-cell table:number-columns-repeated="256"/>
        </table:table-row>
      </xsl:when>
      <!-- when there are only picture, conditional and note in sheet  -->
      <xsl:when
        test="not(e:worksheet/e:sheetData/e:row/e:c/e:v) and $BigMergeCell = '' and $BigMergeRow = '' and ($PictureCell != '' or $NoteCell != '' or $ConditionalCell != '')">

       <xsl:call-template name="InsertEmptySheetWithElements">
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="NameSheet">
            <xsl:value-of select="$NameSheet"/>
          </xsl:with-param>
          <xsl:with-param name="PictureCell">
            <xsl:value-of select="$PictureCell"/>
          </xsl:with-param>
          <xsl:with-param name="PictureRow">
            <xsl:value-of select="$PictureRow"/>
          </xsl:with-param>
          <xsl:with-param name="NoteCell">
            <xsl:value-of select="$NoteCell"/>
          </xsl:with-param>
          <xsl:with-param name="NoteRow">
            <xsl:value-of select="$NoteRow"/>
          </xsl:with-param>
          <xsl:with-param name="ConditionalCell">
            <xsl:value-of select="$ConditionalCell"/>
          </xsl:with-param>
          <xsl:with-param name="ConditionalCellStyle">
            <xsl:value-of select="$ConditionalCellStyle"/>
          </xsl:with-param>
         <xsl:with-param name="ConditionalRow">
           <xsl:value-of select="$ConditionalRow"/>
         </xsl:with-param>
         <xsl:with-param name="sheetNr" select="$sheetNr"/>
        </xsl:call-template>

      </xsl:when>
      <xsl:when test="$BigMergeRow != '' and e:worksheet/e:sheetData/e:row/e:c">
        <table:table-row table:style-name="ro1">
          <table:covered-table-cell table:number-columns-repeated="256"/>
        </table:table-row>
      </xsl:when>
      <xsl:otherwise>
        <!-- it is necessary when sheet has different default row height -->
        <xsl:if
          test="65536 - e:worksheet/e:sheetData/e:row[last()]/@r &gt; 0 or $BigMergeCell != ''">
          <xsl:choose>
            <xsl:when test="$BigMergeCell != ''">
              <xsl:call-template name="InsertColumnsBigMergeRow">
                <xsl:with-param name="Repeat">
                  <xsl:choose>
                    <xsl:when test="e:worksheet/e:sheetData/e:row[last()]/@r">
                      <xsl:value-of select="65536 - e:worksheet/e:sheetData/e:row[last()]/@r"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="65536"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:with-param>
                <xsl:with-param name="RowNumber">
                  <xsl:choose>
                    <xsl:when test="$RowNumber != ''">
                      <xsl:value-of select="$RowNumber"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:text>1</xsl:text>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:with-param>
                <xsl:with-param name="BigMergeCell">
                  <xsl:value-of select="$BigMergeCell"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
                table:number-rows-repeated="{65536 - e:worksheet/e:sheetData/e:row[last()]/@r}">
                <table:table-cell table:number-columns-repeated="256"/>
              </table:table-row>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <xsl:template name="InsertThisRow">
    <xsl:param name="BigMergeCell"/>
    <xsl:param name="BigMergeRow"/>
    <xsl:param name="lastCellColumnNumber"/>
    <xsl:param name="CheckIfBigMerge"/>
    <xsl:param name="this"/>
    <xsl:param name="PictureCell"/>
    <xsl:param name="PictureRow"/>
    <xsl:param name="NoteCell"/>
    <xsl:param name="NoteRow"/>
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="sheetNr"/>
    <xsl:param name="ConditionalCell"/>
    <xsl:param name="ConditionalCellStyle"/>


    <xsl:variable name="GetMinRowWithElements">
      <xsl:call-template name="GetMinRowWithPicture">
         <xsl:with-param name="PictureRow">
          <xsl:value-of select="concat($PictureRow, $NoteRow)"/>
        </xsl:with-param>
        <xsl:with-param name="AfterRow">
          <xsl:value-of select="$lastCellColumnNumber"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="PictureColl">
      <xsl:call-template name="GetCollsWithElement">
        <xsl:with-param name="rowNumber">
          <xsl:value-of select="@r"/>
        </xsl:with-param>
        <xsl:with-param name="ElementCell">
          <xsl:value-of select="concat(';', $PictureCell)"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="NoteColl">
      <xsl:call-template name="GetCollsWithElement">
        <xsl:with-param name="rowNumber">
          <xsl:value-of select="@r"/>
        </xsl:with-param>
        <xsl:with-param name="ElementCell">
          <xsl:value-of select="concat(';', $NoteCell)"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    
    <xsl:choose>

      <!-- Insert if this row is merged with another row -->
      <xsl:when test="contains($CheckIfBigMerge,'true')">

        <table:table-row table:style-name="ro1">
          <table:covered-table-cell table:number-columns-repeated="256"/>
        </table:table-row>

      </xsl:when>

      <xsl:otherwise>

        <table:table-row>
          <xsl:attribute name="table:style-name">
            <xsl:choose>
              <xsl:when test="@ht">
                <xsl:value-of select="generate-id(.)"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>

          <xsl:if test="@hidden=1">
            <xsl:attribute name="table:visibility">
              <xsl:text>collapse</xsl:text>
            </xsl:attribute>
          </xsl:if>


          <!-- Insert First Cell in Row  -->
          <xsl:if test="$CheckIfBigMerge = ''">
            <xsl:apply-templates select="e:c[1]">
              <xsl:with-param name="BigMergeCell">
                <xsl:value-of select="$BigMergeCell"/>
              </xsl:with-param>
              <xsl:with-param name="PictureRow">
                <xsl:value-of select="$PictureRow"/>
              </xsl:with-param>
              <xsl:with-param name="PictureCell">
                <xsl:value-of select="$PictureCell"/>
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
              <xsl:with-param name="ConditionalCell">
                <xsl:value-of select="$ConditionalCell"/>
              </xsl:with-param>
              <xsl:with-param name="ConditionalCellStyle">
                <xsl:value-of select="$ConditionalCellStyle"/>
              </xsl:with-param>
            </xsl:apply-templates>
          </xsl:if>

          <xsl:variable name="CheckIfBigMergeAfter">
            <xsl:call-template name="CheckIfBigMergeAfter">
              <xsl:with-param name="colNum">
                <xsl:value-of select="$lastCellColumnNumber"/>
              </xsl:with-param>
              <xsl:with-param name="rowNum">1</xsl:with-param>
              <xsl:with-param name="BigMergeCell">
                <xsl:value-of select="$BigMergeCell"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>

          <!-- complete row with empty cells if last cell number < 256 -->
          <xsl:choose>
            <xsl:when
              test="$lastCellColumnNumber &lt; 256 and $CheckIfBigMerge = '' and $CheckIfBigMergeAfter != 'true' and $PictureCell = ''">
              <table:table-cell table:number-columns-repeated="{256 - $lastCellColumnNumber}">
                <!-- if there is a default cell style for the row -->
                <xsl:if test="@s">
                  <xsl:attribute name="table:style-name">
                    <xsl:for-each select="document('xl/styles.xml')">
                      <xsl:value-of select="generate-id(key('Xf', '')[position() = $this/@s + 1])"/>
                    </xsl:for-each>
                  </xsl:attribute>
                </xsl:if>
              </table:table-cell>
            </xsl:when>
            <xsl:when
              test="$lastCellColumnNumber &lt; 256 and $CheckIfBigMerge != '' and not(e:c) and $PictureCell = ''">
              <table:table-cell table:number-columns-repeated="{256 - $lastCellColumnNumber}">
                <!-- if there is a default cell style for the row -->
                <xsl:if test="@s">
                  <xsl:attribute name="table:style-name">
                    <xsl:for-each select="document('xl/styles.xml')">
                      <xsl:value-of select="generate-id(key('Xf', '')[position() = $this/@s + 1])"/>
                    </xsl:for-each>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="$CheckIfBigMerge != ''">
                  <xsl:attribute name="table:number-rows-spanned">
                    <xsl:choose>
                      <xsl:when test="number(substring-after($CheckIfBigMerge, ':')) &gt; 65536">
                        <xsl:value-of
                          select="65536 - number(substring-before($CheckIfBigMerge, ':')) + 1"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:choose>
                          <xsl:when
                            test="number(substring-after($CheckIfBigMerge, ':')) &gt; 256">
                            <xsl:value-of
                              select="256 - number(substring-before($CheckIfBigMerge, ':'))"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of
                              select="number(substring-after($CheckIfBigMerge, ':')) - number(substring-before($CheckIfBigMerge, ':')) + 1"
                            />
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
                  <xsl:attribute name="table:number-columns-spanned">256</xsl:attribute>
                </xsl:if>
              </table:table-cell>
            </xsl:when>
            <xsl:when test="$PictureCell != '' and $GetMinRowWithElements=@r and not(e:c)">  
              <xsl:call-template name="InsertElementsBetweenTwoColl">
                <xsl:with-param name="sheet">
                  <xsl:value-of select="$sheet"/>
                </xsl:with-param>
                <xsl:with-param name="NameSheet">
                  <xsl:value-of select="$NameSheet"/>
                </xsl:with-param>
                <xsl:with-param name="PictureCell">
                  <xsl:value-of select="$PictureCell"/>
                </xsl:with-param>
                <xsl:with-param name="PictureRow">
                  <xsl:value-of select="$PictureRow"/>
                </xsl:with-param>
                <xsl:with-param name="NoteCell">
                  <xsl:value-of select="$NoteCell"/>
                </xsl:with-param>
                <xsl:with-param name="NoteRow">
                  <xsl:value-of select="$NoteRow"/>
                </xsl:with-param>
                <!--xsl:with-param name="ConditionalCell">
                  <xsl:value-of select="$ConditionalCell"/>
                  </xsl:with-param>
                  <xsl:with-param name="ConditionalCellStyle">
                  <xsl:value-of select="$ConditionalCellStyle"/>
                  </xsl:with-param-->
                <xsl:with-param name="ElementsColl">
                  <xsl:value-of select="concat($PictureColl, $NoteColl)"/>
                </xsl:with-param>
                <xsl:with-param name="rowNum">
                  <xsl:value-of select="@r"/>
                </xsl:with-param>
                <xsl:with-param name="prevColl">
                  <xsl:text>0</xsl:text>
                </xsl:with-param>
                <xsl:with-param name="sheetNr" select="$sheetNr"/>
                <xsl:with-param name="EndColl">
                  <xsl:text>256</xsl:text>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="$lastCellColumnNumber &lt; 256 and $CheckIfBigMerge != ''">
              <xsl:apply-templates select="e:c[1]">
                <xsl:with-param name="BigMergeCell">
                  <xsl:value-of select="$BigMergeCell"/>
                </xsl:with-param>
              </xsl:apply-templates>
            </xsl:when>
            <xsl:otherwise>
              <xsl:message terminate="no">translation.oox2odf.ColNumber</xsl:message>
            </xsl:otherwise>
          </xsl:choose>
        </table:table-row>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <xsl:template name="InsertThisColumn">
    <xsl:param name="DefaultCellStyleName"/>
    <xsl:param name="headerColsStart"/>
    <xsl:param name="headerColsEnd"/>
    <xsl:param name="beforeHeader" select="'false'"/>
    <xsl:param name="afterHeader" select="'false'"/>

    <table:table-column table:style-name="{generate-id(.)}">

      <xsl:choose>
        <!-- when this is the rest of a column range after header -->
        <xsl:when test="$afterHeader = 'true' ">
          <xsl:attribute name="table:number-columns-repeated">
            <xsl:value-of select="@max - $headerColsEnd"/>
          </xsl:attribute>
        </xsl:when>
        <!-- when this is the part of a column range before header -->
        <xsl:when test="$beforeHeader = 'true'">
          <xsl:attribute name="table:number-columns-repeated">
            <xsl:value-of select="$headerColsStart - @min"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="@min = $headerColsEnd"/>
        <!-- when this column range starts before header and ends after -->
        <xsl:when
          test="$headerColsStart != '' and (@min &lt; $headerColsStart and @max &gt; $headerColsEnd)">
          <xsl:attribute name="table:number-columns-repeated">
            <xsl:value-of select="$headerColsEnd - $headerColsStart + 1"/>
          </xsl:attribute>
        </xsl:when>
        <!-- when this column range starts before header and ends inside -->
        <xsl:when
          test="$headerColsStart != '' and @min &lt; $headerColsStart and @max &gt;= $headerColsStart">
          <xsl:attribute name="table:number-columns-repeated">
            <xsl:value-of select="@max - $headerColsStart + 1"/>
          </xsl:attribute>
        </xsl:when>
        <!-- when this column range starts inside header and ends outside -->
        <xsl:when
          test="$headerColsEnd != '' and @min &lt; $headerColsEnd and @max &gt; $headerColsEnd">
          <xsl:attribute name="table:number-columns-repeated">
            <xsl:value-of select="$headerColsEnd - @min + 1"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:if test="not(@min = @max)">
            <xsl:attribute name="table:number-columns-repeated">
              <xsl:value-of select="@max - @min + 1"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:attribute name="table:default-cell-style-name">
        <xsl:value-of select="$DefaultCellStyleName"/>
      </xsl:attribute>
      <xsl:if test="@hidden=1">
        <xsl:attribute name="table:visibility">
          <xsl:text>collapse</xsl:text>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@style">
        <xsl:variable name="position">
          <xsl:value-of select="@style + 1"/>
        </xsl:variable>
        <xsl:attribute name="table:default-cell-style-name">
          <xsl:for-each select="document('xl/styles.xml')">
            <xsl:value-of select="generate-id(key('Xf', '')[position() = $position])"/>
          </xsl:for-each>
        </xsl:attribute>
      </xsl:if>
    </table:table-column>
  </xsl:template>

  <xsl:template name="InsertEmptyCell">
    <xsl:param name="BeforeMerge"/>
    <xsl:param name="prevCellCol"/>
    <xsl:param name="BigMergeCell"/>
    <xsl:param name="this"/>
    <xsl:param name="colNum"/>
    <xsl:param name="rowNum"/>
    <xsl:param name="CheckIfMerge"/>
    <xsl:param name="PictureCell"/>
    <xsl:param name="PictureRow"/>
    <xsl:param name="NoteCell"/>
    <xsl:param name="NoteRow"/>
    <xsl:param name="PictureColl"/>
    <xsl:param name="NoteColl"/>
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="sheetNr"/>

    <xsl:variable name="CheckIfBigMergeBefore">
      <xsl:call-template name="CheckIfBigMergeBefore">
        <xsl:with-param name="prevCellCol">
          <xsl:value-of select="$prevCellCol"/>
        </xsl:with-param>
        <xsl:with-param name="colNum">
          <xsl:value-of select="$colNum"/>
        </xsl:with-param>
        <xsl:with-param name="rowNum">
          <xsl:value-of select="$rowNum"/>
        </xsl:with-param>
        <xsl:with-param name="BigMergeCell">
          <xsl:value-of select="$BigMergeCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    
    <xsl:variable name="ElementsColl">
      <xsl:value-of select="concat($PictureColl, $NoteColl)"/>
    </xsl:variable>
    
    <xsl:variable name="GetMinCollWithElement">
      <xsl:call-template name="GetMinRowWithPicture">
        <xsl:with-param name="PictureRow">
          <xsl:value-of select="concat(';', $ElementsColl)"/>
        </xsl:with-param>
        <xsl:with-param name="AfterRow">
          <xsl:value-of select="$prevCellCol"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!-- if there were empty cells in a row before this one then insert empty cells-->
    <xsl:choose>
      <!-- when this cell is the first one in a row but not in column A and there aren't pectures before this one -->
      <xsl:when
        test="$prevCellCol = '' and $colNum>1 and $BeforeMerge != 'true' and $CheckIfBigMergeBefore != 'true'  and ($GetMinCollWithElement = '' or ($GetMinCollWithElement != '' and $GetMinCollWithElement &gt;= $colNum))">
        <table:table-cell>
          <xsl:attribute name="table:number-columns-repeated">
            <xsl:value-of select="$colNum - 1"/>
          </xsl:attribute>
          <!-- if there is a default cell style for the row -->
          <xsl:if test="parent::node()/@s">
            <xsl:variable name="position">
              <xsl:value-of select="$this/parent::node()/@s + 1"/>
            </xsl:variable>
            <xsl:attribute name="table:style-name">
              <xsl:for-each select="document('xl/styles.xml')">
                <xsl:value-of select="generate-id(key('Xf', '')[position() = $position])"/>
              </xsl:for-each>
            </xsl:attribute>
          </xsl:if>
        </table:table-cell>
      </xsl:when>
      <xsl:when
        test="position() = 1 and $colNum > 1 and not(preceding-sibling::e:c) and $BeforeMerge != 'true' and $CheckIfBigMergeBefore != 'true' and $GetMinCollWithElement != '' and $GetMinCollWithElement &lt; $colNum">
        
        <!-- Insert picture before this col -->
        <xsl:call-template name="InsertElementsBetweenTwoColl">
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="NameSheet">
            <xsl:value-of select="$NameSheet"/>
          </xsl:with-param>
          <xsl:with-param name="PictureCell">
            <xsl:value-of select="$PictureCell"/>
          </xsl:with-param>
          <xsl:with-param name="PictureRow">
            <xsl:value-of select="$PictureRow"/>
          </xsl:with-param>
          <xsl:with-param name="NoteCell">
            <xsl:value-of select="$NoteCell"/>
          </xsl:with-param>
          <xsl:with-param name="NoteRow">
            <xsl:value-of select="$NoteRow"/>
          </xsl:with-param>
          <!--xsl:with-param name="ConditionalCell">
            <xsl:value-of select="$ConditionalCell"/>
            </xsl:with-param>
            <xsl:with-param name="ConditionalCellStyle">
            <xsl:value-of select="$ConditionalCellStyle"/>
            </xsl:with-param-->
          <xsl:with-param name="ElementsColl">
            <xsl:value-of select="$ElementsColl"/>
          </xsl:with-param>
          <xsl:with-param name="rowNum">
            <xsl:value-of select="$rowNum"/>
          </xsl:with-param>
          <xsl:with-param name="prevColl">
            <xsl:text>0</xsl:text>
          </xsl:with-param>
          <xsl:with-param name="sheetNr" select="$sheetNr"/>
          <xsl:with-param name="EndColl">
            <xsl:value-of select="$colNum"/>
          </xsl:with-param>
        </xsl:call-template>
        
      </xsl:when>
      
      <xsl:when
        test="$colNum>1 and $BeforeMerge != 'true' and $CheckIfBigMergeBefore != 'true' and $GetMinCollWithElement != '' and $GetMinCollWithElement &lt; $colNum">

        <!-- Insert picture before this col -->
        <xsl:call-template name="InsertElementsBetweenTwoColl">
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="NameSheet">
            <xsl:value-of select="$NameSheet"/>
          </xsl:with-param>
          <xsl:with-param name="PictureCell">
            <xsl:value-of select="$PictureCell"/>
          </xsl:with-param>
          <xsl:with-param name="PictureRow">
            <xsl:value-of select="$PictureRow"/>
          </xsl:with-param>
          <xsl:with-param name="NoteCell">
            <xsl:value-of select="$NoteCell"/>
          </xsl:with-param>
          <xsl:with-param name="NoteRow">
            <xsl:value-of select="$NoteRow"/>
          </xsl:with-param>
          <!--xsl:with-param name="ConditionalCell">
            <xsl:value-of select="$ConditionalCell"/>
          </xsl:with-param>
          <xsl:with-param name="ConditionalCellStyle">
            <xsl:value-of select="$ConditionalCellStyle"/>
          </xsl:with-param-->
          <xsl:with-param name="ElementsColl">
            <xsl:value-of select="$ElementsColl"/>
          </xsl:with-param>
          <xsl:with-param name="rowNum">
            <xsl:value-of select="$rowNum"/>
          </xsl:with-param>
          <xsl:with-param name="prevColl">
            <xsl:value-of select="$prevCellCol"/>
          </xsl:with-param>
          <xsl:with-param name="sheetNr" select="$sheetNr"/>
          <xsl:with-param name="EndColl">
            <xsl:value-of select="$colNum"/>
          </xsl:with-param>
        </xsl:call-template>

      </xsl:when>

      <!-- when this cell is not first one in a row and there were empty cells after previous non-empty cell -->
      <xsl:when test="$prevCellCol != '' ">

        <xsl:variable name="prevCellColNum">
          <xsl:choose>
            <!-- if previous column was specified by number -->
            <xsl:when test="$prevCellCol > -1">
              <xsl:value-of select="$prevCellCol"/>
            </xsl:when>
            <!-- if previous column was specified by prewious cell position (i.e. H4) -->
            <xsl:when test="$prevCellCol != ''">
              <xsl:call-template name="GetColNum">
                <xsl:with-param name="cell">
                  <xsl:value-of select="$prevCellCol"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:when>
            <!-- if there wasn't previous cell-->
            <xsl:otherwise>-1</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:if test="$colNum>$prevCellColNum+1">
          <xsl:choose>
            <xsl:when test="$BigMergeCell = ''">
              <table:table-cell>
                <xsl:attribute name="table:number-columns-repeated">
                  <xsl:value-of select="$colNum - $prevCellColNum - 1"/>
                </xsl:attribute>
                <!-- if there is a default cell style for the row -->
                <xsl:if test="parent::node()/@s">
                  <xsl:variable name="position">
                    <xsl:value-of select="$this/parent::node()/@s + 1"/>
                  </xsl:variable>
                  <xsl:attribute name="table:style-name">
                    <xsl:for-each select="document('xl/styles.xml')">
                      <xsl:value-of select="generate-id(key('Xf', '')[position() = $position])"/>
                    </xsl:for-each>
                  </xsl:attribute>
                </xsl:if>
                <xsl:variable name="thisCellCol">
                  <xsl:call-template name="NumbersToChars">
                    <xsl:with-param name="num">
                      <xsl:value-of select="$colNum -1"/>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:variable>
                <xsl:variable name="thisCell">
                  <xsl:value-of select="concat($thisCellCol,$rowNum -1)"/>
                </xsl:variable>
                <xsl:apply-templates
                  select="document(concat('xl/comments',$sheetNr,'.xml'))/e:comments/e:commentList/e:comment[@ref=$thisCell]">
                  <xsl:with-param name="number" select="$sheetNr"/>
                </xsl:apply-templates>
              </table:table-cell>
            </xsl:when>
            <xsl:when test="$CheckIfBigMergeBefore = 'true'">
              <xsl:call-template name="InsertEmptyColumnsSheetNotEmpty">
                <xsl:with-param name="BigMergeCell">
                  <xsl:value-of select="$BigMergeCell"/>
                </xsl:with-param>
                <xsl:with-param name="prevCellCol">
                  <xsl:value-of select="$prevCellCol"/>
                </xsl:with-param>
                <xsl:with-param name="colNum">
                  <xsl:value-of select="$colNum"/>
                </xsl:with-param>
                <xsl:with-param name="rowNum">
                  <xsl:value-of select="$rowNum"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="$colNum - $prevCellCol -1 &gt; 0">
              <table:table-cell>
                <xsl:attribute name="table:number-columns-repeated">
                  <xsl:value-of select="$colNum - $prevCellCol - 1"/>
                </xsl:attribute>
                <!-- if there is a default cell style for the row -->
                <xsl:if test="parent::node()/@s">
                  <xsl:variable name="position">
                    <xsl:value-of select="$this/parent::node()/@s + 1"/>
                  </xsl:variable>
                  <xsl:attribute name="table:style-name">
                    <xsl:for-each select="document('xl/styles.xml')">
                      <xsl:value-of select="generate-id(key('Xf', '')[position() = $position])"/>
                    </xsl:for-each>
                  </xsl:attribute>
                </xsl:if>
                <xsl:variable name="thisCellCol">
                  <xsl:call-template name="NumbersToChars">
                    <xsl:with-param name="num">
                      <xsl:value-of select="$colNum"/>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:variable>
                <xsl:variable name="thisCell">
                  <xsl:value-of select="concat($thisCellCol,$rowNum - 1)"/>
                </xsl:variable>
                <xsl:apply-templates
                  select="document(concat('xl/comments',$sheetNr,'.xml'))/e:comments/e:commentList/e:comment[@ref=$thisCell]">
                  <xsl:with-param name="number" select="$sheetNr"/>
                </xsl:apply-templates>
              </table:table-cell>
            </xsl:when>
          </xsl:choose>
        </xsl:if>
      </xsl:when>
      <xsl:when test="$CheckIfBigMergeBefore = 'true'">
        <xsl:call-template name="InsertEmptyColumnsSheetNotEmpty">
          <xsl:with-param name="BigMergeCell">
            <xsl:value-of select="$BigMergeCell"/>
          </xsl:with-param>
          <xsl:with-param name="prevCellCol">
            <xsl:value-of select="$prevCellCol"/>
          </xsl:with-param>
          <xsl:with-param name="colNum">
            <xsl:value-of select="$colNum"/>
          </xsl:with-param>
          <xsl:with-param name="rowNum">
            <xsl:value-of select="$rowNum"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise> </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <xsl:template name="InsertThisCell">
    <xsl:param name="BeforeMerge"/>
    <xsl:param name="prevCellCol"/>
    <xsl:param name="BigMergeCell"/>
    <xsl:param name="this"/>
    <xsl:param name="colNum"/>
    <xsl:param name="rowNum"/>
    <xsl:param name="CheckIfMerge"/>
    <xsl:param name="PictureCell"/>
    <xsl:param name="PictureRow"/>
    <xsl:param name="PictureColl"/>
    <xsl:param name="NoteCell"/>
    <xsl:param name="NoteRow"/>
    <xsl:param name="NoteColl"/>
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="sheetNr"/>
    <xsl:param name="ConditionalCell"/>
    <xsl:param name="ConditionalCellStyle"/>

    <xsl:message terminate="no">progress:c</xsl:message>

    <xsl:choose>

      <!-- Insert covered cell if this is Merge Cell -->
      <xsl:when test="contains($CheckIfMerge,'true')">
        
        <xsl:call-template name="InsertCoveredTableCell">
          <xsl:with-param name="BeforeMerge">
            <xsl:value-of select="$BeforeMerge"/>
          </xsl:with-param>
          <xsl:with-param name="prevCellCol">
            <xsl:value-of select="$prevCellCol"/>
          </xsl:with-param>
          <xsl:with-param name="BigMergeCell">
            <xsl:value-of select="$BigMergeCell"/>
          </xsl:with-param>
          <xsl:with-param name="this">
            <xsl:value-of select="$this"/>
          </xsl:with-param>
          <xsl:with-param name="colNum">
            <xsl:value-of select="$colNum"/>
          </xsl:with-param>
          <xsl:with-param name="rowNum">
            <xsl:value-of select="$rowNum"/>
          </xsl:with-param>
          <xsl:with-param name="CheckIfMerge">
            <xsl:value-of select="$CheckIfMerge"/>
          </xsl:with-param>
          <xsl:with-param name="PictureCell">
            <xsl:value-of select="$PictureCell"/>
          </xsl:with-param>
          <xsl:with-param name="PictureRow">
            <xsl:value-of select="$PictureRow"/>
          </xsl:with-param>
          <xsl:with-param name="PictureColl">
            <xsl:value-of select="$PictureColl"/>
          </xsl:with-param>
          <xsl:with-param name="NoteCell">
            <xsl:value-of select="$NoteCell"/>
          </xsl:with-param>
          <xsl:with-param name="NoteRow">
            <xsl:value-of select="$NoteRow"/>
          </xsl:with-param>
          <xsl:with-param name="NoteColl">
            <xsl:value-of select="$NoteColl"/>
          </xsl:with-param>
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="NameSheet">
            <xsl:value-of select="$NameSheet"/>
          </xsl:with-param>
          <xsl:with-param name="sheetNr">
            <xsl:value-of select="$sheetNr"/>
          </xsl:with-param>
          <xsl:with-param name="ConditionalCell">
            <xsl:value-of select="$ConditionalCell"/>
          </xsl:with-param>
          <xsl:with-param name="ConditionalCellStyle">
            <xsl:value-of select="$ConditionalCellStyle"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:otherwise>
        <!-- insert this one cell-->

        <table:table-cell>
          <xsl:variable name="position">
            <xsl:value-of select="$this/@s + 1"/>
          </xsl:variable>
          <xsl:variable name="CheckIfPicture">
            <xsl:choose>
              <xsl:when
                test="contains(concat(';', $PictureCell), concat(';', $rowNum, ':', $colNum, ';'))">
                <xsl:text>true</xsl:text>
              </xsl:when>
              <xsl:otherwise>
                <xsl:text>false</xsl:text>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>

          <xsl:variable name="CheckIfNote">
            <xsl:choose>
              <xsl:when
                test="contains(concat(';', $NoteCell), concat(';', $rowNum, ':', $colNum, ';'))">
                <xsl:text>true</xsl:text>
              </xsl:when>
              <xsl:otherwise>
                <xsl:text>false</xsl:text>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>


          <!-- Insert "Merge Cell" if "Merge Cell" is starting in this cell -->
          <xsl:if test="$CheckIfMerge != 'false'">
            <xsl:attribute name="table:number-rows-spanned">
              <xsl:choose>
                <xsl:when test="number(substring-before($CheckIfMerge, ':')) &gt; 65536">65536</xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="substring-before($CheckIfMerge, ':')"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name="table:number-columns-spanned">
              <xsl:choose>
                <xsl:when test="$colNum + number(substring-after($CheckIfMerge, ':')) &gt; 256">
                  <xsl:value-of select="256 - $colNum + 1"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="substring-after($CheckIfMerge, ':')"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>

          <xsl:if
            test="@s or contains(concat(';', $ConditionalCell), concat(';', $rowNum, ':', $colNum, ';'))">
            <xsl:choose>
              
             <xsl:when test="@s and contains(concat(';', $ConditionalCell), concat(';', $rowNum, ':', $colNum, ';'))">
                <xsl:variable name="CellStyleId">
                  <xsl:for-each select="document('xl/styles.xml')">
                    <xsl:value-of select="generate-id(key('Xf', '')[position() = $position])"/>
                  </xsl:for-each>
                </xsl:variable>
                <xsl:variable name="ConditionalStyleId">
                  <xsl:value-of
                    select="generate-id(key('ConditionalFormatting', '')[position() = substring-before(substring-after(concat(';', $ConditionalCellStyle), concat(';', $rowNum, ':', $colNum, ';-')), ';') + 1])"
                  />
                </xsl:variable>
                <xsl:attribute name="table:style-name">
                  <xsl:for-each select="document('xl/styles.xml')">
                    <xsl:value-of select="concat($CellStyleId, $ConditionalStyleId)"/>
                  </xsl:for-each>
                </xsl:attribute>
              </xsl:when>
              
              <xsl:when test="@s">
                <xsl:attribute name="table:style-name">
                  <xsl:for-each select="document('xl/styles.xml')">
                    <xsl:value-of select="generate-id(key('Xf', '')[position() = $position])"/>
                  </xsl:for-each>
                </xsl:attribute>
              </xsl:when>
              
              <xsl:otherwise>
                <xsl:attribute name="table:style-name">
                  <xsl:value-of
                    select="generate-id(key('ConditionalFormatting', '')[position() = substring-before(substring-after(concat(';', $ConditionalCellStyle), concat(';', $rowNum, ':', $colNum, ';-')), ';') + 1])"
                  />
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>


            <xsl:variable name="horizontal">
              <xsl:for-each select="document('xl/styles.xml')">
                <xsl:value-of select="key('Xf', '')[position() = $position]/e:alignment/@horizontal"
                />
              </xsl:for-each>
            </xsl:variable>
            <xsl:if test="$horizontal = 'centerContinuous' and e:v">
              <xsl:variable name="continous">
                <xsl:call-template name="CountContinuous"/>
              </xsl:variable>
              <xsl:attribute name="table:number-columns-spanned">
                <xsl:value-of select="$continous"/>
              </xsl:attribute>
              <xsl:attribute name="table:number-rows-spanned">
                <xsl:text>1</xsl:text>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>

          <xsl:if test="e:v">
            <xsl:call-template name="InsertText">
              <xsl:with-param name="position">
                <xsl:value-of select="$position"/>
              </xsl:with-param>
              <xsl:with-param name="colNum" select="$colNum"/>
              <xsl:with-param name="rowNum" select="$rowNum"/>
              <xsl:with-param name="sheetNr" select="$sheetNr"/>
            </xsl:call-template>
          </xsl:if>

          <xsl:if test="$CheckIfPicture = 'true'">

            <xsl:variable name="Target">
              <xsl:for-each select="ancestor::e:worksheet/e:drawing">
                <xsl:call-template name="GetTargetPicture">
                  <xsl:with-param name="sheet">
                    <xsl:value-of select="substring-after($sheet, '/')"/>
                  </xsl:with-param>
                  <xsl:with-param name="id">
                    <xsl:value-of select="@r:id"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>
            </xsl:variable>

            <xsl:call-template name="InsertPictureInThisCell">
              <xsl:with-param name="sheet">
                <xsl:value-of select="$sheet"/>
              </xsl:with-param>
              <xsl:with-param name="NameSheet">
                <xsl:value-of select="$NameSheet"/>
              </xsl:with-param>
              <xsl:with-param name="collNum">
                <xsl:value-of select="$colNum"/>
              </xsl:with-param>
              <xsl:with-param name="rowNum">
                <xsl:value-of select="$rowNum"/>
              </xsl:with-param>
              <xsl:with-param name="Target">
                <xsl:value-of select="$Target"/>
              </xsl:with-param>
            </xsl:call-template>

          </xsl:if>

          <xsl:if test="$CheckIfNote = 'true'">
            <xsl:call-template name="InsertNoteInThisCell">
              <xsl:with-param name="rowNum">
                <xsl:value-of select="$rowNum"/>
              </xsl:with-param>
              <xsl:with-param name="colNum">
                <xsl:value-of select="$colNum"/>
              </xsl:with-param>
              <xsl:with-param name="sheetNr">
                <xsl:value-of select="$sheetNr"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:if>

        </table:table-cell>

        <!-- Insert covered cell if Merge Cell is starting-->
       
        <xsl:choose>
          <xsl:when
            test="$CheckIfMerge != 'false' and substring-after($CheckIfMerge, ':') &gt; 1">
            <xsl:call-template name="InsertCoveredTableCell">
              <xsl:with-param name="BeforeMerge">
                <xsl:value-of select="$BeforeMerge"/>
              </xsl:with-param>
              <xsl:with-param name="prevCellCol">
                <xsl:value-of select="$prevCellCol"/>
              </xsl:with-param>
              <xsl:with-param name="BigMergeCell">
                <xsl:value-of select="$BigMergeCell"/>
              </xsl:with-param>
              <xsl:with-param name="this">
                <xsl:value-of select="$this"/>
              </xsl:with-param>
              <xsl:with-param name="colNum">
                <xsl:value-of select="$colNum + 1"/>
              </xsl:with-param>
              <xsl:with-param name="rowNum">
                <xsl:value-of select="$rowNum"/>
              </xsl:with-param>
              <xsl:with-param name="CheckIfMerge">
                <xsl:value-of select="$CheckIfMerge"/>
              </xsl:with-param>
              <xsl:with-param name="PictureCell">
                <xsl:value-of select="$PictureCell"/>
              </xsl:with-param>
              <xsl:with-param name="PictureRow">
                <xsl:value-of select="$PictureRow"/>
              </xsl:with-param>
              <xsl:with-param name="PictureColl">
                <xsl:value-of select="$PictureColl"/>
              </xsl:with-param>
              <xsl:with-param name="NoteCell">
                <xsl:value-of select="$NoteCell"/>
              </xsl:with-param>
              <xsl:with-param name="NoteRow">
                <xsl:value-of select="$NoteRow"/>
              </xsl:with-param>
              <xsl:with-param name="NoteColl">
                <xsl:value-of select="$NoteColl"/>
              </xsl:with-param>
              <xsl:with-param name="sheet">
                <xsl:value-of select="$sheet"/>
              </xsl:with-param>
              <xsl:with-param name="NameSheet">
                <xsl:value-of select="$NameSheet"/>
              </xsl:with-param>
              <xsl:with-param name="sheetNr">
                <xsl:value-of select="$sheetNr"/>
              </xsl:with-param>
              <xsl:with-param name="ConditionalCell">
                <xsl:value-of select="$ConditionalCell"/>
              </xsl:with-param>
              <xsl:with-param name="ConditionalCellStyle">
                <xsl:value-of select="$ConditionalCellStyle"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <!-- when cell had 'centerContinuous' horizontal alignment -->
          <xsl:when test="@s and e:v">
            <xsl:variable name="position">
              <xsl:value-of select="$this/@s + 1"/>
            </xsl:variable>
            <xsl:variable name="horizontal">
              <xsl:for-each select="document('xl/styles.xml')">
                <xsl:value-of select="key('Xf', '')[position() = $position]/e:alignment/@horizontal"
                />
              </xsl:for-each>
            </xsl:variable>
            <xsl:variable name="continuous">
              <xsl:call-template name="CountContinuous"/>
            </xsl:variable>
            <xsl:if test="$horizontal = 'centerContinuous' ">
              <table:covered-table-cell>
                <xsl:attribute name="table:number-columns-repeated">
                  <xsl:value-of select="$continuous - 1"/>
                </xsl:attribute>
                <xsl:attribute name="table:style-name">
                  <xsl:for-each select="document('xl/styles.xml')">
                    <xsl:value-of select="generate-id(key('Xf', '')[position() = 1])"/>
                  </xsl:for-each>
                </xsl:attribute>
              </table:covered-table-cell>
            </xsl:if>
          </xsl:when>
        </xsl:choose>

      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <!-- change  '%20' to space  after conversion-->
  <xsl:template name="Change20PercentToSpace">
    <xsl:param name="string"/>

    <xsl:choose>
      <xsl:when test="contains($string,'%20')">
        <xsl:choose>
          <xsl:when test="substring-before($string,'%20') =''">
            <xsl:call-template name="Change20PercentToSpace">
              <xsl:with-param name="string" select="concat(' ',substring-after($string,'%20'))"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="substring-before($string,'%20') !=''">
            <xsl:call-template name="Change20PercentToSpace">
              <xsl:with-param name="string"
                select="concat(substring-before($string,'%20'),' ',substring-after($string,'%20'))"
              />
            </xsl:call-template>
          </xsl:when>
        </xsl:choose>
      </xsl:when>

      <!--xsl:when test="contains($slash,'..\..\..\')">
        <xsl:choose>
          <xsl:when test="substring-before($slash,'..\..\..\') =''">
            <xsl:call-template name="Change20PercentToSpace">
              <xsl:with-param name="slash"
                select="concat('../../../',substring-after($slash,'..\..\..\'))"/>
            </xsl:call-template>
          </xsl:when>
        </xsl:choose>
      </xsl:when-->

      <xsl:when test="contains($string,'..\..\')">
        <xsl:choose>
          <xsl:when test="substring-before($string,'..\..\') =''">

            <xsl:call-template name="Change20PercentToSpace">
              <xsl:with-param name="string"
                select="concat('../../../',substring-after($string,'..\..\'))"/>
            </xsl:call-template>
          </xsl:when>
        </xsl:choose>
      </xsl:when>

      <xsl:when test="contains($string,'..\')">
        <xsl:choose>
          <xsl:when test="substring-before($string,'..\') =''">

            <xsl:call-template name="Change20PercentToSpace">
              <xsl:with-param name="string" select="concat('../../',substring-after($string,'..\'))"
              />
            </xsl:call-template>
          </xsl:when>
        </xsl:choose>
      </xsl:when>


      <xsl:when test="not(contains($string,'../')) and not(contains($string,'..\..\'))">

        <xsl:value-of select="concat('../',$string)">
          <!--xsl:call-template name="Change20PercentToSpace">
              <xsl:with-param name="slash" select="concat('..\',substring-after($string,''))"/>
            </xsl:call-template-->
        </xsl:value-of>

      </xsl:when>

      <xsl:otherwise>
        <xsl:value-of select="$string"/>
        <!--xsl:value-of select="translate($string,'\','/')"/-->
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <xsl:template name="InsertText">
    <xsl:param name="position"/>
    <xsl:param name="colNum"/>
    <xsl:param name="rowNum"/>
    <xsl:param name="sheetNr"/>
    
    <xsl:choose>
      <xsl:when test="@t='s' ">
        <xsl:attribute name="office:value-type">
          <xsl:text>string</xsl:text>
        </xsl:attribute>
        <xsl:variable name="id">
          <xsl:value-of select="e:v"/>
        </xsl:variable>
        <text:p>
          <xsl:choose>
            <xsl:when test="key('ref',@r)">
              <text:a>
                <xsl:attribute name="xlink:href">
                  <xsl:variable name="target">
                    <!-- path to sheet file from xl/ catalog (i.e. $sheet = worksheets/sheet1.xml) -->
                    <xsl:for-each select="key('ref',@r)">
                      <xsl:call-template name="GetTarget">
                        <xsl:with-param name="id">
                          <xsl:value-of select="@r:id"/>
                        </xsl:with-param>
                          <xsl:with-param name="document">
                            <xsl:value-of select="concat('xl/worksheets/sheet', $sheetNr, '.xml')"/>
                          </xsl:with-param>
                      </xsl:call-template>
                    </xsl:for-each>
                  </xsl:variable>
                  <xsl:choose>
                    <!-- when hyperlink leads to a file in network -->
                    <xsl:when test="starts-with($target,'file:///\\')">
                      <xsl:value-of select="translate(substring-after($target,'file:///'),'\','/')"/>
                    </xsl:when>
                    <!--when hyperlink leads to www or mailto -->
                    <xsl:when test="contains($target,':')">
                      <xsl:value-of select="$target"/>
                    </xsl:when>
                    <!-- when hyperlink leads to another place in workbook -->
                    <xsl:when test="key('ref',@r)/@location != '' ">
                      <xsl:for-each select="key('ref',@r)">

                        <xsl:variable name="apos">
                          <xsl:text>&apos;</xsl:text>
                        </xsl:variable>

                        <xsl:variable name="sheetName">
                          <xsl:choose>
                            <xsl:when test="starts-with(@location,$apos)">
                              <xsl:value-of select="$apos"/>
                              <xsl:value-of
                                select="substring-before(substring-after(@location,$apos),$apos)"/>
                              <xsl:value-of select="$apos"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="substring-before(@location,'!')"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:variable>

                        <xsl:variable name="invalidChars">
                          <xsl:text>&apos;!$-()</xsl:text>
                        </xsl:variable>

                        <xsl:variable name="checkedName">
                          <xsl:for-each select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet[@name = translate($sheetName,$apos,'')]">
                            <xsl:call-template name="CheckSheetName">
                              <xsl:with-param name="sheetNumber">
                                <xsl:for-each
                                  select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet[@name = translate($sheetName,$apos,'')]">
                                  <xsl:value-of select="count(preceding-sibling::e:sheet) + 1"/>
                                </xsl:for-each>
                              </xsl:with-param>
                              <xsl:with-param name="name">
                                <xsl:value-of select="translate($sheetName,$invalidChars,'')"/>
                              </xsl:with-param>
                            </xsl:call-template>
                          </xsl:for-each>
                        </xsl:variable>

                        <xsl:text>#</xsl:text>
                        <xsl:value-of select="$checkedName"/>
                        <xsl:text>.</xsl:text>
                        <xsl:value-of select="substring-after(@location,concat($sheetName,'!'))"/>

                      </xsl:for-each>
                    </xsl:when>
                    <!--when hyperlink leads to a document -->
                    <xsl:otherwise>
                      <xsl:call-template name="Change20PercentToSpace">
                        <xsl:with-param name="string" select="$target"/>
                      </xsl:call-template>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>

                <!-- a postprocessor puts here strings from sharedstrings -->
                <pxsi:v xmlns:pxsi="urn:cleverage:xmlns:post-processings:shared-strings">
                  <xsl:value-of select="e:v"/>
                </pxsi:v>

              </text:a>
            </xsl:when>

            <xsl:otherwise>
              <!-- a postprocessor puts here strings from sharedstrings -->
              <pxsi:v xmlns:pxsi="urn:cleverage:xmlns:post-processings:shared-strings">
                <xsl:value-of select="e:v"/>
              </pxsi:v>
            </xsl:otherwise>
          </xsl:choose>
        </text:p>
      </xsl:when>
      <xsl:when test="@t = 'e' ">
        <xsl:attribute name="office:value-type">
          <xsl:text>string</xsl:text>
        </xsl:attribute>
        <text:p>
          <xsl:value-of select="e:v"/>
        </text:p>
      </xsl:when>
      <xsl:when test="@t = 'str' ">
        <xsl:attribute name="office:value-type">
          <xsl:choose>
            <xsl:when test="number(e:v)">
              <xsl:text>float</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>string</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <text:p>
          <xsl:choose>
            <xsl:when test="number(e:v)">
              <xsl:call-template name="FormatNumber">
                <xsl:with-param name="value">
                  <xsl:value-of select="e:v"/>
                </xsl:with-param>
                <xsl:with-param name="numStyle">
                  <xsl:for-each select="document('xl/styles.xml')">
                    <xsl:value-of
                      select="key('numFmtId',key('Xf','')[position()=$position]/@numFmtId)/@formatCode"
                    />
                  </xsl:for-each>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="e:v"/>
            </xsl:otherwise>
          </xsl:choose>
        </text:p>
      </xsl:when>
      <xsl:when test="@t = 'n'">
        <xsl:attribute name="office:value-type">
          <xsl:text>float</xsl:text>
        </xsl:attribute>
        <xsl:attribute name="office:value">
          <xsl:value-of select="e:v"/>
        </xsl:attribute>
        <text:p>
          <xsl:call-template name="FormatNumber">
            <xsl:with-param name="value">
              <xsl:value-of select="e:v"/>
            </xsl:with-param>
            <xsl:with-param name="numStyle">
              <xsl:for-each select="document('xl/styles.xml')">
                <xsl:value-of
                  select="key('numFmtId',key('Xf','')[position()=$position]/@numFmtId)/@formatCode"
                />
              </xsl:for-each>
            </xsl:with-param>
          </xsl:call-template>
        </text:p>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="numStyle">
          <xsl:for-each select="document('xl/styles.xml')">
            <xsl:value-of
              select="key('numFmtId',key('Xf','')[position()=$position]/@numFmtId)/@formatCode"/>
          </xsl:for-each>
        </xsl:variable>
        <xsl:variable name="numId">
          <xsl:for-each select="document('xl/styles.xml')">
            <xsl:value-of select="key('Xf','')[position()=$position]/@numFmtId"/>
          </xsl:for-each>
        </xsl:variable>
        <xsl:attribute name="office:value-type">
          <xsl:choose>
            <xsl:when
              test="contains($numStyle,'%') or ((not($numStyle) or $numStyle = '')  and ($numId = 9 or $numId = 10))">
              <xsl:text>percentage</xsl:text>
            </xsl:when>
            <xsl:when
              test="(contains($numStyle,'y') or (contains($numStyle,'m') and not(contains($numStyle,'h') or contains($numStyle,'s'))) or (contains($numStyle,'d') and not(contains($numStyle,'Red'))) or ($numId &gt; 13 and $numId &lt; 18) or $numId = 22)">
              <xsl:text>date</xsl:text>
            </xsl:when>

            <!--'and' at the end is for Latvian currency -->
            <xsl:when
              test="contains($numStyle,'h') or contains($numStyle,'s') and not(contains($numStyle,'[$Ls-426]'))">
              <xsl:text>time</xsl:text>
            </xsl:when>

            <xsl:when
              test="contains($numStyle,'zł') or contains($numStyle,'$') or contains($numStyle,'£') or contains($numStyle,'€')">
              <xsl:text>currency</xsl:text>
            </xsl:when>
            <xsl:when test="$numId = 49">
              <xsl:text>string</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>float</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:choose>

          <xsl:when
            test="(contains($numStyle,'y') or (contains($numStyle,'m') and not(contains($numStyle,'h') or contains($numStyle,'s'))) or (contains($numStyle,'d') and not(contains($numStyle,'Red'))) or ($numId &gt; 13 and $numId &lt; 18) or $numId = 22)">
            <xsl:attribute name="office:date-value">
              <xsl:call-template name="NumberToDate">
                <xsl:with-param name="value">
                  <xsl:value-of select="e:v"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:when>

          <!--'and' at the end is for Latvian currency -->
          <xsl:when
            test="contains($numStyle,'h') or contains($numStyle,'s') and not(contains($numStyle,'[$Ls-426]'))">
            <xsl:attribute name="office:time-value">
              <xsl:call-template name="NumberToTime">
                <xsl:with-param name="value">
                  <xsl:value-of select="e:v"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="office:value">
              <xsl:value-of select="e:v"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
        <text:p>
          <xsl:choose>

            <xsl:when
              test="(contains($numStyle,'y') or (contains($numStyle,'m') and not(contains($numStyle,'h') or contains($numStyle,'s'))) or (contains($numStyle,'d') and not(contains($numStyle,'Red'))) or ($numId &gt; 13 and $numId &lt; 18) or $numId = 22)">
              <xsl:call-template name="FormatDate">
                <xsl:with-param name="value">
                  <xsl:call-template name="NumberToDate">
                    <xsl:with-param name="value">
                      <xsl:value-of select="e:v"/>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:with-param>
                <xsl:with-param name="format">
                  <xsl:choose>
                    <xsl:when test="contains($numStyle,']')">
                      <xsl:value-of select="substring-after($numStyle,']')"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$numStyle"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:with-param>
                <xsl:with-param name="numId">
                  <xsl:value-of select="$numId"/>
                </xsl:with-param>
                <xsl:with-param name="processedFormat">
                  <xsl:choose>
                    <xsl:when test="contains($numStyle,']')">
                      <xsl:value-of select="substring-after($numStyle,']')"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$numStyle"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:with-param>
                <xsl:with-param name="numValue">
                  <xsl:value-of select="e:v"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:when>

            <!--'and' at the end is for Latvian currency -->
            <xsl:when
              test="contains($numStyle,'h') or contains($numStyle,'s') and not(contains($numStyle,'[$Ls-426]'))">
              <xsl:call-template name="FormatTime">
                <xsl:with-param name="value">
                  <xsl:call-template name="NumberToTime">
                    <xsl:with-param name="value">
                      <xsl:value-of select="e:v"/>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:with-param>
                <xsl:with-param name="format">
                  <xsl:choose>
                    <xsl:when test="contains($numStyle,']') and not(contains($numStyle,'[h'))">
                      <xsl:value-of select="substring-after($numStyle,']')"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$numStyle"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:with-param>
                <xsl:with-param name="numId">
                  <xsl:value-of select="$numId"/>
                </xsl:with-param>
                <xsl:with-param name="processedFormat">
                  <xsl:choose>
                    <xsl:when test="contains($numStyle,']') and not(contains($numStyle,'[h'))">
                      <xsl:value-of select="substring-after($numStyle,']')"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$numStyle"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:with-param>
                <xsl:with-param name="numValue">
                  <xsl:value-of select="e:v"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:when>

            <xsl:otherwise>
              <xsl:call-template name="FormatNumber">
                <xsl:with-param name="value">
                  <xsl:value-of select="e:v"/>
                </xsl:with-param>
                <xsl:with-param name="numStyle">
                  <xsl:value-of select="$numStyle"/>
                </xsl:with-param>
                <xsl:with-param name="numId">
                  <xsl:value-of select="$numId"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </text:p>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertNextCell">
    <xsl:param name="BeforeMerge"/>
    <xsl:param name="prevCellCol"/>
    <xsl:param name="BigMergeCell"/>
    <xsl:param name="this"/>
    <xsl:param name="colNum"/>
    <xsl:param name="rowNum"/>
    <xsl:param name="CheckIfMerge"/>
    <xsl:param name="PictureCell"/>
    <xsl:param name="PictureRow"/>
    <xsl:param name="PictureColl"/>
    <xsl:param name="NoteCell"/>
    <xsl:param name="NoteRow"/>
    <xsl:param name="NoteColl"/>
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="GetMinCollWithElement"/>
    <xsl:param name="sheetNr"/>
    <xsl:param name="ConditionalCell"/>
    <xsl:param name="ConditionalCellStyle"/>



    
    <xsl:variable name="CheckIfBigMergeBefore">
      <xsl:call-template name="CheckIfBigMergeBefore">
        <xsl:with-param name="prevCellCol">
          <xsl:value-of select="$prevCellCol"/>
        </xsl:with-param>
        <xsl:with-param name="colNum">
          <xsl:value-of select="$colNum"/>
        </xsl:with-param>
        <xsl:with-param name="rowNum">
          <xsl:value-of select="$rowNum"/>
        </xsl:with-param>
        <xsl:with-param name="BigMergeCell">
          <xsl:value-of select="$BigMergeCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="CheckIfBigMergeAfter">
      <xsl:call-template name="CheckIfBigMergeAfter">
        <xsl:with-param name="colNum">
          <xsl:value-of select="$colNum"/>
        </xsl:with-param>
        <xsl:with-param name="rowNum">
          <xsl:value-of select="$rowNum"/>
        </xsl:with-param>
        <xsl:with-param name="BigMergeCell">
          <xsl:value-of select="$BigMergeCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>


    <xsl:variable name="countContinuous">
      <xsl:choose>
        <xsl:when test="@s and e:v">
          <xsl:variable name="position">
            <xsl:value-of select="$this/@s + 1"/>
          </xsl:variable>
          <xsl:variable name="horizontal">
            <xsl:for-each select="document('xl/styles.xml')">
              <xsl:value-of select="key('Xf', '')[position() = $position]/e:alignment/@horizontal"/>
            </xsl:for-each>
          </xsl:variable>

          <xsl:choose>
            <xsl:when test="$horizontal = 'centerContinuous' ">
              <xsl:call-template name="CountContinuous"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>0</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:text>0</xsl:text>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    
    <xsl:variable name="ElementsColl">
      <xsl:value-of select="concat($PictureColl, $NoteColl)"/>
    </xsl:variable>
    
 
    <xsl:choose>

      <!-- calc supports only 256 columns -->
      <xsl:when test="$colNum &gt; 255">
        <xsl:message terminate="no">translation.oox2odf.ColNumber</xsl:message>
      </xsl:when>

      <!-- Skips empty coll (in Merge Cell) -->

      <xsl:when test="$CheckIfMerge != 'false' and following-sibling::e:c">
        <xsl:choose>
          <!-- if in this cell big merge coll is starting -->
          <xsl:when test="$CheckIfBigMergeBefore = 'start' and following-sibling::e:c">
            <xsl:apply-templates select="following-sibling::e:c[1]">
              <xsl:with-param name="BeforeMerge">
                <xsl:text>true</xsl:text>
              </xsl:with-param>
              <xsl:with-param name="prevCellCol">
                <xsl:value-of select="$colNum + number(substring-after($CheckIfMerge, ':')) - 1"/>
              </xsl:with-param>
              <xsl:with-param name="BigMergeCell">
                <xsl:value-of select="$BigMergeCell"/>
              </xsl:with-param>
              <xsl:with-param name="PictureRow">
                <xsl:value-of select="$PictureRow"/>
              </xsl:with-param>
              <xsl:with-param name="PictureCell">
                <xsl:value-of select="$PictureCell"/>
              </xsl:with-param>
              <xsl:with-param name="NoteCell">
                <xsl:value-of select="$NoteCell"/>
              </xsl:with-param>
              <xsl:with-param name="NoteRow">
                <xsl:value-of select="$NoteRow"/>
              </xsl:with-param>
              <xsl:with-param name="sheet">
                <xsl:value-of select="$sheet"/>
              </xsl:with-param>
              <xsl:with-param name="NameSheet">
                <xsl:value-of select="$NameSheet"/>
              </xsl:with-param>
              <xsl:with-param name="sheetNr" select="$sheetNr"/>
              <xsl:with-param name="ConditionalCell">
                <xsl:value-of select="$ConditionalCell"/>
              </xsl:with-param>
              <xsl:with-param name="ConditionalCellStyle">
                <xsl:value-of select="$ConditionalCellStyle"/>
              </xsl:with-param>
            </xsl:apply-templates>
          </xsl:when>
          <!-- if this cell is inside row of merged cells ($CheckIfMerged is true:number_of_cols_spaned) -->
          <xsl:when
            test="contains($CheckIfMerge,'true') and substring-after($CheckIfMerge, ':') &gt; 1">
            <xsl:if test="following-sibling::e:c[number(substring-after($CheckIfMerge, ':')) - 1]">
              <xsl:apply-templates
                select="following-sibling::e:c[number(substring-after($CheckIfMerge, ':')) - 1]">
                <xsl:with-param name="BeforeMerge">
                  <xsl:text>true</xsl:text>
                </xsl:with-param>
                <xsl:with-param name="prevCellCol">
                  <xsl:value-of select="$colNum + number(substring-after($CheckIfMerge, ':')) - 1"/>
                </xsl:with-param>
                <xsl:with-param name="BigMergeCell">
                  <xsl:value-of select="$BigMergeCell"/>
                </xsl:with-param>
                <xsl:with-param name="PictureRow">
                  <xsl:value-of select="$PictureRow"/>
                </xsl:with-param>
                <xsl:with-param name="PictureCell">
                  <xsl:value-of select="$PictureCell"/>
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
                <xsl:with-param name="ConditionalCell">
                  <xsl:value-of select="$ConditionalCell"/>
                </xsl:with-param>
                <xsl:with-param name="ConditionalCellStyle">
                  <xsl:value-of select="$ConditionalCellStyle"/>
                </xsl:with-param>
              </xsl:apply-templates>
            </xsl:if>
          </xsl:when>

          <!-- if this cell starts row of merged cells ($CheckIfMerged has dimensions of merged cell) -->
          <xsl:otherwise>

            <xsl:if test="following-sibling::e:c[number(substring-after($CheckIfMerge, ':'))]">
              <xsl:apply-templates
                select="following-sibling::e:c[number(substring-after($CheckIfMerge, ':'))]">
                <xsl:with-param name="BeforeMerge">
                  <xsl:text>true</xsl:text>
                </xsl:with-param>
                <xsl:with-param name="prevCellCol">
                  <xsl:value-of select="$colNum + number(substring-after($CheckIfMerge, ':')) - 1"/>
                </xsl:with-param>
                <xsl:with-param name="BigMergeCell">
                  <xsl:value-of select="$BigMergeCell"/>
                </xsl:with-param>
                <xsl:with-param name="PictureRow">
                  <xsl:value-of select="$PictureRow"/>
                </xsl:with-param>
                <xsl:with-param name="PictureCell">
                  <xsl:value-of select="$PictureCell"/>
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
                <xsl:with-param name="ConditionalCell">
                  <xsl:value-of select="$ConditionalCell"/>
                </xsl:with-param>
                <xsl:with-param name="ConditionalCellStyle">
                  <xsl:value-of select="$ConditionalCellStyle"/>
                </xsl:with-param>
              </xsl:apply-templates>
            </xsl:if>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <!-- Insert Empty Big Merge Coll after this cell -->
      <xsl:when test="not(following-sibling::e:c) and $CheckIfBigMergeAfter = 'true'">
        <xsl:call-template name="InsertEmptyColumnsSheetNotEmpty">
          <xsl:with-param name="BigMergeCell">
            <xsl:value-of select="$BigMergeCell"/>
          </xsl:with-param>
          <xsl:with-param name="prevCellCol">
            <xsl:value-of select="$colNum"/>
          </xsl:with-param>
          <xsl:with-param name="colNum">256</xsl:with-param>
          <xsl:with-param name="rowNum">
            <xsl:value-of select="$rowNum"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <!-- skip merged cells when horizontal alignment = 'centerContinuous' -->
      <xsl:when test="$countContinuous != 0">

        <xsl:if test="following-sibling::e:c[$countContinuous]">
          <xsl:apply-templates select="following-sibling::e:c[position() = $countContinuous]">
            <xsl:with-param name="BeforeMerge">
              <xsl:text>true</xsl:text>
            </xsl:with-param>
            <xsl:with-param name="prevCellCol">
              <xsl:value-of select="$colNum + $countContinuous - 1"/>
            </xsl:with-param>
            <xsl:with-param name="PictureRow">
              <xsl:value-of select="$PictureRow"/>
            </xsl:with-param>
            <xsl:with-param name="PictureCell">
              <xsl:value-of select="$PictureCell"/>
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
            <xsl:with-param name="ConditionalCell">
              <xsl:value-of select="$ConditionalCell"/>
            </xsl:with-param>
            <xsl:with-param name="ConditionalCellStyle">
              <xsl:value-of select="$ConditionalCellStyle"/>
            </xsl:with-param>
          </xsl:apply-templates>
        </xsl:if>
      </xsl:when>

     <xsl:when
       test="not(following-sibling::e:c) and ($PictureCell != '' or $NoteCell != '') and $GetMinCollWithElement &gt; $colNum">

       <!-- Insert picture before this col -->
       <xsl:call-template name="InsertElementsBetweenTwoColl">
         <xsl:with-param name="sheet">
           <xsl:value-of select="$sheet"/>
         </xsl:with-param>
         <xsl:with-param name="NameSheet">
           <xsl:value-of select="$NameSheet"/>
         </xsl:with-param>
         <xsl:with-param name="PictureCell">
           <xsl:value-of select="$PictureCell"/>
         </xsl:with-param>
         <xsl:with-param name="PictureRow">
           <xsl:value-of select="$PictureRow"/>
         </xsl:with-param>
         <xsl:with-param name="NoteCell">
           <xsl:value-of select="$NoteCell"/>
         </xsl:with-param>
         <xsl:with-param name="NoteRow">
           <xsl:value-of select="$NoteRow"/>
         </xsl:with-param>
         <!--xsl:with-param name="ConditionalCell">
           <xsl:value-of select="$ConditionalCell"/>
           </xsl:with-param>
           <xsl:with-param name="ConditionalCellStyle">
           <xsl:value-of select="$ConditionalCellStyle"/>
           </xsl:with-param-->
         <xsl:with-param name="ElementsColl">
           <xsl:value-of select="$ElementsColl"/>
           </xsl:with-param>
         <xsl:with-param name="rowNum">
           <xsl:value-of select="$rowNum"/>
         </xsl:with-param>
         <xsl:with-param name="prevColl">
           <xsl:value-of select="$prevCellCol"/>
         </xsl:with-param>
         <xsl:with-param name="sheetNr" select="$sheetNr"/>
         <xsl:with-param name="EndColl">
           <xsl:text>256</xsl:text>
         </xsl:with-param>
       </xsl:call-template>
    
       
       

      </xsl:when>

      <xsl:otherwise>

        <xsl:if test="following-sibling::e:c">
          <xsl:apply-templates select="following-sibling::e:c[1]">
            <xsl:with-param name="prevCellCol">
              <xsl:call-template name="GetColNum">
                <xsl:with-param name="cell">
                  <xsl:value-of select="@r"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:with-param>
            <xsl:with-param name="BigMergeCell">
              <xsl:value-of select="$BigMergeCell"/>
            </xsl:with-param>
            <xsl:with-param name="PictureRow">
              <xsl:value-of select="$PictureRow"/>
            </xsl:with-param>
            <xsl:with-param name="PictureCell">
              <xsl:value-of select="$PictureCell"/>
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
            <xsl:with-param name="sheetNr">
              <xsl:value-of select="$sheetNr"/>
            </xsl:with-param>
            <xsl:with-param name="ConditionalCell">
              <xsl:value-of select="$ConditionalCell"/>
            </xsl:with-param>
            <xsl:with-param name="ConditionalCellStyle">
              <xsl:value-of select="$ConditionalCellStyle"/>
            </xsl:with-param>
          </xsl:apply-templates>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <!--  convert multiple white spaces  -->
  <xsl:template name="InsertWhiteSpaces">
    <xsl:param name="string" select="."/>
    <xsl:param name="length" select="string-length(.)"/>
    <!-- string which doesn't contain whitespaces-->
    <xsl:choose>
      <xsl:when test="not(contains($string,' '))">
        <xsl:value-of select="$string"/>
      </xsl:when>
      <!-- convert white spaces  -->
      <xsl:otherwise>
        <xsl:variable name="before">
          <xsl:value-of select="substring-before($string,' ')"/>
        </xsl:variable>
        <xsl:variable name="after">
          <xsl:call-template name="CutStartSpaces">
            <xsl:with-param name="cuted">
              <xsl:value-of select="substring-after($string,' ')"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:if test="$before != '' ">
          <xsl:value-of select="concat($before,' ')"/>
        </xsl:if>
        <!--add remaining whitespaces as text:s if there are any-->
        <xsl:if test="string-length(concat($before,' ', $after)) &lt; $length ">
          <xsl:choose>
            <xsl:when test="($length - string-length(concat($before, $after))) = 1">
              <text:s/>
            </xsl:when>
            <xsl:otherwise>
              <text:s>
                <xsl:attribute name="text:c">
                  <xsl:choose>
                    <xsl:when test="$before = ''">
                      <xsl:value-of select="$length - string-length($after)"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$length - string-length(concat($before,' ', $after))"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
              </text:s>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
        <!--repeat it for substring which has whitespaces-->
        <xsl:if test="contains($string,' ') and $length &gt; 0">
          <xsl:call-template name="InsertWhiteSpaces">
            <xsl:with-param name="string">
              <xsl:value-of select="$after"/>
            </xsl:with-param>
            <xsl:with-param name="length">
              <xsl:value-of select="string-length($after)"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--  cut start spaces -->
  <xsl:template name="CutStartSpaces">
    <xsl:param name="cuted"/>
    <xsl:choose>
      <xsl:when test="starts-with($cuted,' ')">
        <xsl:call-template name="CutStartSpaces">
          <xsl:with-param name="cuted">
            <xsl:value-of select="substring-after($cuted,' ')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$cuted"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertColumns">
    <xsl:param name="sheet"/>

    <xsl:variable name="DefaultCellStyleName">
      <xsl:for-each select="document('xl/styles.xml')">
        <xsl:value-of select="generate-id(key('Xf', '')[1])"/>
      </xsl:for-each>
    </xsl:variable>

    <xsl:variable name="sheetName">
      <xsl:value-of select="@name"/>
    </xsl:variable>

    <xsl:variable name="apos">
      <xsl:text>&apos;</xsl:text>
    </xsl:variable>

    <xsl:variable name="charHeaderColsStart">
      <xsl:choose>
        <!-- if sheet name in range definition is in apostrophes -->
        <xsl:when
          test="document('xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($apos,$sheetName,$apos))]">
          <xsl:for-each
            select="document('xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($apos,$sheetName,$apos))]">
            <xsl:choose>
              <!-- when header columns are present -->
              <xsl:when test="contains(text(),',')">
                <xsl:value-of
                  select="substring-before(substring-after(substring-after(substring-before(text(),','),$apos),concat($apos,'!$')),':')"
                />
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of
                  select="substring-before(substring-after(substring-after(text(),$apos),concat($apos,'!$')),':')"
                />
              </xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </xsl:when>

        <xsl:when
          test="document('xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($sheetName,'!'))]">
          <xsl:for-each
            select="document('xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($sheetName,'!'))]">
            <xsl:choose>
              <!-- when header columns are present -->
              <xsl:when test="contains(text(),',')">
                <xsl:value-of
                  select="substring-before(substring-after(substring-before(text(),','),'$'),':')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="substring-before(substring-after(text(),'$'),':')"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="charHeaderColsEnd">
      <xsl:choose>
        <!-- if sheet name in range definition is in apostrophes -->
        <xsl:when
          test="document('xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($apos,$sheetName,$apos))]">
          <xsl:for-each
            select="document('xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($apos,$sheetName,$apos))]">
            <xsl:choose>
              <!-- when header columns are present -->
              <xsl:when test="contains(text(),',')">
                <xsl:value-of
                  select="substring-after(substring-after(substring-before(text(),','),':'),'$')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="substring-after(substring-after(text(),':'),'$')"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </xsl:when>

        <xsl:when
          test="document('xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($sheetName,'!'))]">
          <xsl:for-each
            select="document('xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($sheetName,'!'))]">
            <xsl:choose>
              <!-- when header columns are present -->
              <xsl:when test="contains(text(),',')">
                <xsl:value-of
                  select="substring-after(substring-after(substring-before(text(),','),':'),'$')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="substring-after(substring-after(text(),':'),'$')"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="headerColsStart">
      <xsl:if test="$charHeaderColsStart != '' and not(number($charHeaderColsStart))">
        <xsl:call-template name="GetAlphabeticPosition">
          <xsl:with-param name="literal" select="$charHeaderColsStart"/>
        </xsl:call-template>
      </xsl:if>
    </xsl:variable>

    <xsl:variable name="headerColsEnd">
      <xsl:if test="$charHeaderColsEnd != '' and not(number($charHeaderColsEnd))">
        <xsl:call-template name="GetAlphabeticPosition">
          <xsl:with-param name="literal" select="$charHeaderColsEnd"/>
        </xsl:call-template>
      </xsl:if>
    </xsl:variable>

    <xsl:for-each select="document(concat('xl/',$sheet))/e:worksheet">

      <xsl:apply-templates select="e:cols/e:col[1]">
        <xsl:with-param name="number">1</xsl:with-param>
        <xsl:with-param name="sheet" select="$sheet"/>
        <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
        <xsl:with-param name="headerColsStart" select="$headerColsStart"/>
        <xsl:with-param name="headerColsEnd" select="$headerColsEnd"/>
      </xsl:apply-templates>

      <!-- if all columns are default (there aren't any e:col tags) and there is a header -->
      <xsl:if test="not(e:cols/e:col) and $headerColsStart != '' ">

        <!-- insert columns before header -->
        <xsl:if test="$headerColsStart &gt; 1">
          <table:table-column table:style-name="{generate-id(key('SheetFormatPr', ''))}"
            table:number-columns-repeated="{$headerColsStart - 1}">
            <xsl:attribute name="table:default-cell-style-name">
              <xsl:value-of select="$DefaultCellStyleName"/>
            </xsl:attribute>
          </table:table-column>
        </xsl:if>

        <!-- insert header columns -->
        <table:table-header-columns>
          <table:table-column table:style-name="{generate-id(key('SheetFormatPr', ''))}"
            table:number-columns-repeated="{$headerColsEnd - $headerColsStart + 1}">
            <xsl:attribute name="table:default-cell-style-name">
              <xsl:value-of select="$DefaultCellStyleName"/>
            </xsl:attribute>
          </table:table-column>
        </table:table-header-columns>
      </xsl:if>
    </xsl:for-each>

    <!-- apply default column style for last columns which style wasn't changed -->
    <xsl:for-each select="document(concat('xl/',$sheet))">
      <xsl:choose>
        <xsl:when
          test="$headerColsStart != '' and not(key('Col', '')[@max &gt; $headerColsEnd])">
          <table:table-column table:style-name="{generate-id(key('SheetFormatPr', ''))}"
            table:number-columns-repeated="{256 - $headerColsEnd}">
            <xsl:attribute name="table:default-cell-style-name">
              <xsl:value-of select="$DefaultCellStyleName"/>
            </xsl:attribute>
          </table:table-column>
        </xsl:when>
        <xsl:when test="not(key('Col', ''))">
          <table:table-column table:style-name="{generate-id(key('SheetFormatPr', ''))}"
            table:number-columns-repeated="256">
            <xsl:attribute name="table:default-cell-style-name">
              <xsl:value-of select="$DefaultCellStyleName"/>
            </xsl:attribute>
          </table:table-column>
        </xsl:when>
        <xsl:otherwise>
          <xsl:if test="key('Col', '')[last()]/@max &lt; 256">
            <table:table-column table:style-name="{generate-id(key('SheetFormatPr', ''))}"
              table:number-columns-repeated="{256 - key('Col', '')[last()]/@max}">
              <xsl:attribute name="table:default-cell-style-name">
                <xsl:value-of select="$DefaultCellStyleName"/>
              </xsl:attribute>
            </table:table-column>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="e:col">
    <xsl:param name="number"/>
    <xsl:param name="sheet"/>
    <xsl:param name="DefaultCellStyleName"/>
    <xsl:param name="headerColsStart"/>
    <xsl:param name="headerColsEnd"/>

    <!-- if there were columns with default properties before this column then insert default columns-->
    <xsl:choose>
      <!-- when this column is the first non-default one but it's not the column A -->
      <xsl:when
        test="$number = 1 and @min &gt; 1 and ($headerColsStart= '' or $headerColsStart &gt; 1)">
        <table:table-column>
          <xsl:attribute name="table:style-name">
            <xsl:for-each select="document(concat('xl/',$sheet))">
              <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
            </xsl:for-each>
          </xsl:attribute>

          <xsl:attribute name="table:default-cell-style-name">
            <xsl:value-of select="$DefaultCellStyleName"/>
          </xsl:attribute>

          <xsl:attribute name="table:number-columns-repeated">
            <xsl:choose>
              <!-- when there is a header -->
              <xsl:when test="$headerColsStart != '' ">
                <xsl:choose>
                  <xsl:when test="@min &lt; $headerColsStart">
                    <xsl:value-of select="@min - 1"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="$headerColsStart - 1"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>

              <xsl:otherwise>
                <xsl:value-of select="@min - 1"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>

          <!-- Possible are nesesary code -->
          <!--xsl:if test="@style">            
            <xsl:variable name="position">
            <xsl:value-of select="$this/@style + 1"/>
            </xsl:variable>
            
            <xsl:attribute name="table:default-cell-style-name">
              <xsl:for-each select="document('xl/styles.xml')">
                <xsl:value-of select="generate-id(key('Xf', '')[position() = $position])"/>
              </xsl:for-each>
            </xsl:attribute>
            </xsl:if-->

        </table:table-column>

      </xsl:when>
      <!-- when this column is not first non-default one and there were default columns after previous non-default column (if there was a gap between this and previous column)-->
      <xsl:when test="preceding-sibling::e:col[1]/@max &lt; @min - 1">
        <xsl:choose>
          <!-- when there is a header -->
          <xsl:when test="$headerColsStart != '' ">
            <xsl:choose>
              <xsl:when test="preceding-sibling::e:col/@max + 1 = $headerColsStart"/>
              <!-- insert empty columns before header -->
              <xsl:when
                test="@min &gt; $headerColsStart and preceding-sibling::e:col[1]/@max &lt; $headerColsStart">
                <table:table-column>
                  <xsl:attribute name="table:style-name">
                    <xsl:for-each select="document(concat('xl/',$sheet))">
                      <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
                    </xsl:for-each>
                  </xsl:attribute>

                  <xsl:attribute name="table:default-cell-style-name">
                    <xsl:value-of select="$DefaultCellStyleName"/>
                  </xsl:attribute>

                  <xsl:if test="$headerColsStart - preceding::e:col[1]/@max - 1 &gt; 1">
                    <xsl:attribute name="table:number-columns-repeated">
                      <xsl:value-of select="$headerColsStart - preceding::e:col[1]/@max - 1"/>
                    </xsl:attribute>
                  </xsl:if>
                </table:table-column>
              </xsl:when>
              <xsl:when test="@min = $headerColsEnd + 1"/>
              <!-- insert empty columns after header -->
              <xsl:when
                test="@min &gt; $headerColsEnd and preceding-sibling::e:col[1]/@max &lt; $headerColsEnd">
                <table:table-column>
                  <xsl:attribute name="table:style-name">
                    <xsl:for-each select="document(concat('xl/',$sheet))">
                      <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
                    </xsl:for-each>
                  </xsl:attribute>

                  <xsl:attribute name="table:default-cell-style-name">
                    <xsl:value-of select="$DefaultCellStyleName"/>
                  </xsl:attribute>

                  <xsl:if test="@min - $headerColsEnd - 1 &gt; 1">
                    <xsl:attribute name="table:number-columns-repeated">
                      <xsl:value-of select="@min - $headerColsEnd - 1"/>
                    </xsl:attribute>
                  </xsl:if>
                </table:table-column>
              </xsl:when>
              <!-- insert simple empty rows -->
              <xsl:otherwise>
                <table:table-column>
                  <xsl:attribute name="table:style-name">
                    <xsl:for-each select="document(concat('xl/',$sheet))">
                      <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
                    </xsl:for-each>
                  </xsl:attribute>

                  <xsl:attribute name="table:default-cell-style-name">
                    <xsl:value-of select="$DefaultCellStyleName"/>
                  </xsl:attribute>

                  <xsl:if test="@min - preceding::e:col[1]/@max - 1 &gt; 1">
                    <xsl:attribute name="table:number-columns-repeated">
                      <xsl:value-of select="@min - preceding::e:col[1]/@max - 1"/>
                    </xsl:attribute>
                  </xsl:if>
                </table:table-column>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>

          <xsl:otherwise>
            <table:table-column>
              <xsl:attribute name="table:style-name">
                <xsl:for-each select="document(concat('xl/',$sheet))">
                  <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
                </xsl:for-each>
              </xsl:attribute>

              <xsl:attribute name="table:number-columns-repeated">
                <xsl:value-of select="@min - preceding::e:col[1]/@max - 1"/>
              </xsl:attribute>

              <xsl:attribute name="table:default-cell-style-name">
                <xsl:value-of select="$DefaultCellStyleName"/>
              </xsl:attribute>
            </table:table-column>
          </xsl:otherwise>
        </xsl:choose>

      </xsl:when>
    </xsl:choose>

    <!-- insert this column -->
    <xsl:choose>
      <!-- if this is the first column after beginning of the header -->
      <xsl:when
        test="$headerColsStart != '' and (@max &gt;= $headerColsStart and not(preceding-sibling::e:col[1]/@max &gt;= $headerColsStart))">

        <!-- insert part of a column range (range: @max > @min) before header-->
        <xsl:if test="@min &lt; $headerColsStart and @max &gt;= $headerColsStart">
          <xsl:call-template name="InsertThisColumn">
            <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
            <xsl:with-param name="headerColsStart" select="$headerColsStart"/>
            <xsl:with-param name="beforeHeader" select="'true'"/>
          </xsl:call-template>
        </xsl:if>

        <table:table-header-columns>

          <!-- insert column settings at the beginning of the header -->
          <xsl:choose>
            <!-- insert previous column settings -->
            <xsl:when test="preceding-sibling::e:col[1][@max &gt;= $headerColsStart]">
              <xsl:variable name="preceding">
                <xsl:value-of select="preceding-sibling::node()[1]"/>
              </xsl:variable>

              <table:table-column table:style-name="{generate-id(preceding-sibling::node()[1])}">
                <xsl:attribute name="table:number-columns-repeated">
                  <xsl:value-of select="preceding-sibling::node()[1]/@max - $headerColsStart - 1"/>
                </xsl:attribute>

                <xsl:attribute name="table:default-cell-style-name">
                  <xsl:value-of select="$DefaultCellStyleName"/>
                </xsl:attribute>

                <xsl:if test="preceding-sibling::e:col[1]/@hidden=1">
                  <xsl:attribute name="table:visibility">
                    <xsl:text>collapse</xsl:text>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="preceding-sibling::e:col[1]/@style">
                  <xsl:variable name="position">
                    <xsl:value-of select="preceding-sibling::e:col[1]/@style + 1"/>
                  </xsl:variable>
                  <xsl:attribute name="table:default-cell-style-name">
                    <xsl:for-each select="document('xl/styles.xml')">
                      <xsl:value-of select="generate-id(key('Xf', '')[position() = $position])"/>
                    </xsl:for-each>
                  </xsl:attribute>
                </xsl:if>
              </table:table-column>
            </xsl:when>
            <xsl:when test="@min &gt; $headerColsStart">
              <table:table-column table:style-name="{generate-id(key('SheetFormatPr', ''))}">
                <xsl:if test="@min - $headerColsStart &gt; 0">
                  <xsl:attribute name="table:number-columns-repeated">
                    <xsl:choose>
                      <xsl:when test="@min &lt;= $headerColsEnd">
                        <xsl:value-of select="@min - $headerColsStart"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="$headerColsEnd - $headerColsStart + 1"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
                </xsl:if>

                <xsl:attribute name="table:default-cell-style-name">
                  <xsl:value-of select="$DefaultCellStyleName"/>
                </xsl:attribute>
              </table:table-column>
            </xsl:when>
          </xsl:choose>

          <xsl:if test="@min &lt;= $headerColsEnd">
            <xsl:call-template name="InsertThisColumn">
              <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
              <xsl:with-param name="headerColsStart" select="$headerColsStart"/>
              <xsl:with-param name="headerColsEnd" select="$headerColsEnd"/>
            </xsl:call-template>
          </xsl:if>

          <!-- insert next header-column -->
          <xsl:if
            test="following-sibling::e:col[@min &lt;= $headerColsEnd and @min &lt; 256][1]">
            <xsl:apply-templates
              select="following-sibling::e:col[@min &lt;= $headerColsEnd and @min &lt; 256][1]"
              mode="header">
              <xsl:with-param name="number" select="@max + 1"/>
              <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
              <xsl:with-param name="sheet" select="$sheet"/>
              <xsl:with-param name="headerColsStart" select="$headerColsStart"/>
              <xsl:with-param name="headerColsEnd" select="$headerColsEnd"/>
            </xsl:apply-templates>
          </xsl:if>

          <!-- insert empty columns at the end of the header -->
          <xsl:for-each
            select="parent::node()/e:col[@min &lt;= 256 and @max &gt;= $headerColsStart and @min &lt;= $headerColsEnd][last()]">
            <xsl:if test="@max &lt; $headerColsEnd">
              <table:table-column table:style-name="{generate-id(key('SheetFormatPr', ''))}">

                <xsl:attribute name="table:number-columns-repeated">
                  <xsl:value-of select="$headerColsEnd - @max"/>
                </xsl:attribute>

                <xsl:attribute name="table:default-cell-style-name">
                  <xsl:value-of select="$DefaultCellStyleName"/>
                </xsl:attribute>
              </table:table-column>
            </xsl:if>
          </xsl:for-each>

        </table:table-header-columns>

        <!-- if this column range starts inside header, but ends outside write rest of the columns outside header -->
        <xsl:if test="@min &lt;= $headerColsEnd and @max &gt; $headerColsEnd">
          <xsl:call-template name="InsertThisColumn">
            <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
            <xsl:with-param name="headerColsEnd" select="$headerColsEnd"/>
            <xsl:with-param name="afterHeader" select="'true'"/>
          </xsl:call-template>
        </xsl:if>

        <!-- if further there is column range starts inside header, but ends outside write rest of the columns outside header -->
        <xsl:for-each
          select="following-sibling::e:col[@min &lt;= $headerColsEnd and @max &gt; $headerColsEnd]">
          <xsl:call-template name="InsertThisColumn">
            <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
            <xsl:with-param name="headerColsEnd" select="$headerColsEnd"/>
            <xsl:with-param name="afterHeader" select="'true'"/>
          </xsl:call-template>
        </xsl:for-each>

        <!-- if header is empty -->
        <xsl:if test="@min &gt; $headerColsEnd">

          <xsl:choose>
            <!-- first row after start of the header is right after end of the header -->
            <xsl:when test="@min = $headerColsEnd + 1">
              <xsl:call-template name="InsertThisColumn">
                <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <!-- insert default columns between header and column -->
              <table:table-column table:style-name="{generate-id(key('SheetFormatPr', ''))}">

                <xsl:if test="@min &gt; $headerColsEnd + 2">
                  <xsl:attribute name="table:number-columns-repeated">
                    <xsl:value-of select="@min - $headerColsEnd - 1"/>
                  </xsl:attribute>
                </xsl:if>

                <xsl:attribute name="table:default-cell-style-name">
                  <xsl:value-of select="$DefaultCellStyleName"/>
                </xsl:attribute>
              </table:table-column>

              <xsl:call-template name="InsertThisColumn">
                <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>

      </xsl:when>
      <!-- if header is after the last column -->
      <xsl:when test="not(following-sibling::e:col) and @max &lt; $headerColsStart">

        <xsl:call-template name="InsertThisColumn">
          <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
        </xsl:call-template>

        <!-- insert default columns before header -->
        <xsl:if test="$headerColsStart &gt; @max + 1">
          <table:table-column table:style-name="{generate-id(key('SheetFormatPr', ''))}">

            <xsl:if test="$headerColsStart &gt; @max + 2">
              <xsl:attribute name="table:number-columns-repeated">
                <xsl:value-of select="$headerColsStart - @max - 1"/>
              </xsl:attribute>
            </xsl:if>

            <xsl:attribute name="table:default-cell-style-name">
              <xsl:value-of select="$DefaultCellStyleName"/>
            </xsl:attribute>
          </table:table-column>
        </xsl:if>

        <table:table-header-columns>
          <table:table-column table:style-name="{generate-id(key('SheetFormatPr', ''))}">

            <xsl:if test="$headerColsEnd  - $headerColsStart &gt; 1">
              <xsl:attribute name="table:number-columns-repeated">
                <xsl:value-of select="$headerColsEnd  - $headerColsStart + 1"/>
              </xsl:attribute>
            </xsl:if>

            <xsl:attribute name="table:default-cell-style-name">
              <xsl:value-of select="$DefaultCellStyleName"/>
            </xsl:attribute>
          </table:table-column>
        </table:table-header-columns>

      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="InsertThisColumn">
          <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>

    <xsl:choose>

      <!-- calc supports only 256 columns -->
      <xsl:when test="$number &gt; 255">
        <xsl:message terminate="no">translation.oox2odf.ColNumber</xsl:message>
      </xsl:when>

      <xsl:when test="$headerColsStart = '' ">
        <xsl:apply-templates select="following-sibling::e:col[1]">
          <xsl:with-param name="number" select="@max + 1"/>
          <xsl:with-param name="sheet" select="$sheet"/>
          <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
          <xsl:with-param name="headerColsStart" select="$headerColsStart"/>
          <xsl:with-param name="headerColsEnd" select="$headerColsEnd"/>
        </xsl:apply-templates>
      </xsl:when>

      <!-- if next is the first header column -->
      <xsl:when
        test="@max &lt; $headerColsStart and following-sibling::e:col[1][@min &lt;= $headerColsEnd and @max &gt;= $headerColsStart]">
        <xsl:apply-templates select="following-sibling::e:col[1]">
          <xsl:with-param name="number" select="@max + 1"/>
          <xsl:with-param name="sheet" select="$sheet"/>
          <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
          <xsl:with-param name="headerColsStart" select="$headerColsStart"/>
          <xsl:with-param name="headerColsEnd" select="$headerColsEnd"/>
        </xsl:apply-templates>
      </xsl:when>

      <!-- jump over the header -->
      <xsl:when
        test="following-sibling::e:col[@min &lt; $headerColsStart or @min &gt; $headerColsEnd]">
        <xsl:apply-templates
          select="following-sibling::e:col[@min &lt; $headerColsStart or @min &gt; $headerColsEnd][1]">
          <xsl:with-param name="number" select="@max + 1"/>
          <xsl:with-param name="sheet" select="$sheet"/>
          <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
          <xsl:with-param name="headerColsStart" select="$headerColsStart"/>
          <xsl:with-param name="headerColsEnd" select="$headerColsEnd"/>
        </xsl:apply-templates>
      </xsl:when>
    </xsl:choose>

  </xsl:template>

  <xsl:template match="e:col" mode="header">
    <xsl:param name="number"/>
    <xsl:param name="sheet"/>
    <xsl:param name="DefaultCellStyleName"/>
    <xsl:param name="headerColsStart"/>
    <xsl:param name="headerColsEnd"/>

    <!-- if there were columns with default properties before this column then insert default columns-->
    <xsl:choose>
      <!-- when this column is the first non-default one but it's not the column A -->
      <xsl:when
        test="$number = 1 and @min &gt; 1 and ($headerColsStart= '' or $headerColsStart &gt; 1)">
        <table:table-column>
          <xsl:attribute name="table:style-name">
            <xsl:for-each select="document(concat('xl/',$sheet))">
              <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
            </xsl:for-each>
          </xsl:attribute>

          <xsl:attribute name="table:default-cell-style-name">
            <xsl:value-of select="$DefaultCellStyleName"/>
          </xsl:attribute>

          <xsl:attribute name="table:number-columns-repeated">
            <xsl:choose>
              <!-- when there is a header -->
              <xsl:when test="$headerColsStart != '' ">
                <xsl:choose>
                  <xsl:when test="@min &lt; $headerColsStart">
                    <xsl:value-of select="@min - 1"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="$headerColsStart - 1"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>

              <xsl:otherwise>
                <xsl:value-of select="@min - 1"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>

          <!-- Possible are nesesary code -->
          <!--xsl:if test="@style">            
            <xsl:variable name="position">
            <xsl:value-of select="$this/@style + 1"/>
            </xsl:variable>
            
            <xsl:attribute name="table:default-cell-style-name">
            <xsl:for-each select="document('xl/styles.xml')">
            <xsl:value-of select="generate-id(key('Xf', '')[position() = $position])"/>
            </xsl:for-each>
            </xsl:attribute>
            </xsl:if-->

        </table:table-column>

      </xsl:when>
      <!-- when this column is not first non-default one and there were default columns after previous non-default column (if there was a gap between this and previous column)-->
      <xsl:when test="preceding-sibling::e:col[1]/@max &lt; @min - 1 ">
        <table:table-column>
          <xsl:attribute name="table:style-name">
            <xsl:for-each select="document(concat('xl/',$sheet))">
              <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
            </xsl:for-each>
          </xsl:attribute>

          <xsl:attribute name="table:number-columns-repeated">
            <xsl:choose>
              <!-- when there is a header -->
              <xsl:when test="$headerColsStart != '' ">
                <xsl:choose>
                  <!-- insert empty columns before header -->
                  <xsl:when
                    test="@min &gt; $headerColsStart and preceding-sibling::e:col[1]/@max &lt; $headerColsStart">
                    <xsl:value-of select="$headerColsStart - preceding::e:col[1]/@max - 1"/>
                  </xsl:when>
                  <!-- insert empty columns after header -->
                  <xsl:when
                    test="@min &gt; $headerColsEnd and preceding-sibling::e:col[1]/@max &lt; $headerColsEnd">
                    <xsl:value-of select="@max - $headerColsEnd - 1"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="@min - preceding::e:col[1]/@max - 1"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>

              <xsl:otherwise>
                <xsl:value-of select="@min - preceding::e:col[1]/@max - 1"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>

          <xsl:attribute name="table:default-cell-style-name">
            <xsl:value-of select="$DefaultCellStyleName"/>
          </xsl:attribute>

        </table:table-column>
      </xsl:when>
    </xsl:choose>

    <!-- insert this column -->
    <xsl:call-template name="InsertThisColumn">
      <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
      <xsl:with-param name="headerColsEnd" select="$headerColsEnd"/>
    </xsl:call-template>

    <!-- insert next column -->
    <xsl:choose>
      <!-- calc supports only 256 columns -->
      <xsl:when test="$number &gt; 255">
        <xsl:message terminate="no">translation.oox2odf.ColNumber</xsl:message>
      </xsl:when>

      <xsl:when
        test="following-sibling::e:col[1][@min &lt;= $headerColsEnd and @min &lt; 256]">
        <xsl:apply-templates select="following-sibling::e:col[1]" mode="header">
          <xsl:with-param name="number" select="@max + 1"/>
          <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
          <xsl:with-param name="sheet" select="$sheet"/>
          <xsl:with-param name="headerColsStart" select="$headerColsStart"/>
          <xsl:with-param name="headerColsEnd" select="$headerColsEnd"/>
        </xsl:apply-templates>
      </xsl:when>

    </xsl:choose>

  </xsl:template>

  <xsl:template name="FormatNumber">

    <!-- @Descripition: inserts number to cell in a correct format -->
    <!-- @Context: None -->

    <xsl:param name="value"/>
    <!-- (float) number value -->
    <xsl:param name="numStyle"/>
    <!-- (string) number format -->
    <xsl:param name="numId"/>
    <!-- (int) number format ID -->
    <xsl:variable name="formatCode">
      <xsl:choose>
        <xsl:when test="contains($numStyle,';')">
          <xsl:value-of select="substring-before($numStyle,';')"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$numStyle"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="outputValue">
      <xsl:choose>
        <xsl:when test="contains($value,'.') and $numStyle and $numStyle!=''">
          <xsl:call-template name="FormatAfterComma">
            <xsl:with-param name="valueAfterComma">
              <xsl:value-of select="substring-after($value,'.')"/>
            </xsl:with-param>
            <xsl:with-param name="valueBeforeComma">
              <xsl:value-of select="substring-before($value,'.')"/>
            </xsl:with-param>
            <xsl:with-param name="format">
              <xsl:choose>
                <xsl:when test="contains($formatCode,'_')">
                  <xsl:value-of select="substring-before($formatCode,'_')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$formatCode"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="contains($value,'.') and $numId = 10">
          <xsl:call-template name="FormatAfterComma">
            <xsl:with-param name="valueAfterComma">
              <xsl:value-of select="substring-after($value,'.')"/>
            </xsl:with-param>
            <xsl:with-param name="valueBeforeComma">
              <xsl:value-of select="substring-before($value,'.')"/>
            </xsl:with-param>
            <xsl:with-param name="format">0.00%</xsl:with-param>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="$numId = 9">
          <xsl:value-of select="format-number($value,'0%')"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$value"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    
    <xsl:choose>

      <!-- add '%' if it's percentage format-->
      <xsl:when test="contains($formatCode,'%') and not(contains($outputValue,'%'))">
        <xsl:value-of select="concat($outputValue,'%')"/>
      </xsl:when>

      <!-- add currency symbol if there is one-->
      <xsl:when test="contains($formatCode,'[$') or contains($formatCode,'zł')">
        <xsl:variable name="currency">
          <xsl:choose>
            <xsl:when test="contains($formatCode,'zł')">zł</xsl:when>
            <xsl:when test="contains($formatCode,'Red')">
              <xsl:variable name="tempFormat">
                <xsl:value-of
                  select="substring-after(substring-before(substring-after($formatCode,'Red]'),']'),'[$')"
                />
              </xsl:variable>
              <xsl:choose>
                <xsl:when test="contains($tempFormat,'-')">
                  <xsl:value-of select="substring-before($tempFormat,'-')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$tempFormat"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
              <xsl:variable name="tempFormat2">
                <xsl:value-of select="substring-after(substring-before($formatCode,']'),'[$')"/>
              </xsl:variable>
              <xsl:choose>
                <xsl:when test="contains($tempFormat2,'-')">
                  <xsl:value-of select="substring-before($tempFormat2,'-')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$tempFormat2"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:choose>
          <xsl:when
            test="contains(substring-before($formatCode,$currency),'0') or contains(substring-before($formatCode,$currency),'#')">
            <xsl:value-of select="concat($outputValue,$currency)"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat($currency,$outputValue)"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <xsl:otherwise>
        <xsl:value-of select="$outputValue"/>
      </xsl:otherwise>

    </xsl:choose>
  </xsl:template>

  <xsl:template name="FormatAfterComma">

    <!-- @Descripition: formats number after comma -->
    <!-- @Context: None -->

    <xsl:param name="valueAfterComma"/>
    <!-- (int) number value after comma -->
    <xsl:param name="valueBeforeComma"/>
    <!-- (int) number value before comma -->
    <xsl:param name="format"/>
    <!-- (string) format code -->
    <xsl:variable name="plainFormat">
      <xsl:choose>
        <xsl:when test="contains(substring-after($format,'.'),'\')">
          <xsl:value-of
            select="concat(concat(substring-before($format,'.'),'.'),substring-before(substring-after($format,'.'),'\'))"
          />
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$format"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="not(contains($format,'.'))">
        <xsl:value-of select="$valueBeforeComma"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of
          select="concat($valueBeforeComma,format-number(concat('.',$valueAfterComma),concat('.',substring-after($plainFormat,'.'))))"
        />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="CountContinuous">
    <xsl:param name="count" select="0"/>

    <xsl:variable name="carryOn">
      <xsl:choose>
        <xsl:when test="following-sibling::e:c[1]/@s and not(following-sibling::e:c[1]/e:v)">
          <xsl:variable name="position">
            <xsl:value-of select="following-sibling::e:c[1]/@s + 1"/>
          </xsl:variable>
          <xsl:variable name="horizontal">
            <xsl:for-each select="document('xl/styles.xml')">
              <xsl:value-of select="key('Xf', '')[position() = $position]/e:alignment/@horizontal"/>
            </xsl:for-each>
          </xsl:variable>

          <xsl:choose>
            <xsl:when test="$horizontal = 'centerContinuous' ">
              <xsl:text>true</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>false</xsl:text>
            </xsl:otherwise>
          </xsl:choose>

        </xsl:when>
        <xsl:otherwise>
          <xsl:text>false</xsl:text>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="$carryOn = 'true'">
        <xsl:for-each select="following-sibling::e:c[1]">
          <xsl:call-template name="CountContinuous">
            <xsl:with-param name="count" select="$count + 1"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
      <xsl:otherwise>
        <!-- number of following cells plus the starting one -->
        <xsl:value-of select="$count + 1"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="FormatDate">

    <!-- @Descripition: inserts date to cell in a correct format -->
    <!-- @Context: None -->

    <xsl:param name="value"/>
    <!-- (dateTime) input date value -->
    <xsl:param name="format"/>
    <!-- (string) format code -->
    <xsl:param name="numId"/>
    <!-- (int) format ID -->
    <xsl:param name="processedFormat"/>
    <!-- (string) part of format code which is being processed -->
    <xsl:param name="outputValue"/>
    <!-- (dateTime) output date value -->
    <xsl:param name="numValue"/>
    <!-- (float) date value as a number  -->
    <xsl:choose>

      <!-- year -->
      <xsl:when test="starts-with($processedFormat,'y')">
        <xsl:choose>
          <xsl:when test="starts-with(substring-after($processedFormat,'y'),'yyy')">
            <xsl:call-template name="FormatDate">
              <xsl:with-param name="value" select="$value"/>
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="numId" select="$numId"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'yyyy')"/>
              </xsl:with-param>
              <xsl:with-param name="outputValue">
                <xsl:value-of select="concat($outputValue,substring-before($value,'-'))"/>
              </xsl:with-param>
              <xsl:with-param name="numValue" select="$numValue"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="FormatDate">
              <xsl:with-param name="value" select="$value"/>
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="numId" select="$numId"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'yy')"/>
              </xsl:with-param>
              <xsl:with-param name="outputValue">
                <xsl:value-of
                  select="concat($outputValue,substring(substring-before($value,'-'),3))"/>
              </xsl:with-param>
              <xsl:with-param name="numValue" select="$numValue"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <xsl:when test="starts-with($processedFormat,'m')">
        <xsl:choose>

          <!-- minutes -->
          <xsl:when test="contains(substring-before($format,'m'),'h:')">
            <xsl:choose>
              <xsl:when test="starts-with(substring-after($processedFormat,'m'),'m')">
                <xsl:call-template name="FormatDate">
                  <xsl:with-param name="value" select="$value"/>
                  <xsl:with-param name="format" select="$format"/>
                  <xsl:with-param name="numId" select="$numId"/>
                  <xsl:with-param name="processedFormat">
                    <xsl:value-of select="substring-after($processedFormat,'mm')"/>
                  </xsl:with-param>
                  <xsl:with-param name="outputValue">
                    <xsl:value-of
                      select="concat($outputValue,substring-before(substring-after(substring-after($value,'T'),':'),':'))"
                    />
                  </xsl:with-param>
                  <xsl:with-param name="numValue" select="$numValue"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:otherwise>
                <xsl:call-template name="FormatDate">
                  <xsl:with-param name="value" select="$value"/>
                  <xsl:with-param name="format" select="$format"/>
                  <xsl:with-param name="numId" select="$numId"/>
                  <xsl:with-param name="processedFormat">
                    <xsl:value-of select="substring-after($processedFormat,'m')"/>
                  </xsl:with-param>
                  <xsl:with-param name="outputValue">
                    <xsl:value-of
                      select="concat($outputValue,number(substring-before(substring-after(substring-after($value,'T'),':'),':')))"
                    />
                  </xsl:with-param>
                  <xsl:with-param name="numValue" select="$numValue"/>
                </xsl:call-template>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>

          <!-- month -->
          <xsl:otherwise>
            <xsl:choose>
              <xsl:when test="starts-with(substring-after($processedFormat,'m'),'mmm')">
                <xsl:call-template name="FormatDate">
                  <xsl:with-param name="value" select="$value"/>
                  <xsl:with-param name="format" select="$format"/>
                  <xsl:with-param name="numId" select="$numId"/>
                  <xsl:with-param name="processedFormat">
                    <xsl:value-of select="substring-after($processedFormat,'mmmm')"/>
                  </xsl:with-param>
                  <xsl:with-param name="outputValue">
                    <xsl:variable name="monthName">
                      <xsl:call-template name="ConvertMonthToName">
                        <xsl:with-param name="month">
                          <xsl:value-of select="substring-before(substring-after($value,'-'),'-')"/>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:value-of select="concat($outputValue,$monthName)"/>
                  </xsl:with-param>
                  <xsl:with-param name="numValue" select="$numValue"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="starts-with(substring-after($processedFormat,'m'),'mm')">
                <xsl:call-template name="FormatDate">
                  <xsl:with-param name="value" select="$value"/>
                  <xsl:with-param name="format" select="$format"/>
                  <xsl:with-param name="numId" select="$numId"/>
                  <xsl:with-param name="processedFormat">
                    <xsl:value-of select="substring-after($processedFormat,'mmm')"/>
                  </xsl:with-param>
                  <xsl:with-param name="outputValue">
                    <xsl:variable name="monthName">
                      <xsl:call-template name="ConvertMonthToShortName">
                        <xsl:with-param name="month">
                          <xsl:value-of select="substring-before(substring-after($value,'-'),'-')"/>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:value-of select="concat($outputValue,$monthName)"/>
                  </xsl:with-param>
                  <xsl:with-param name="numValue" select="$numValue"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="starts-with(substring-after($processedFormat,'m'),'m')">
                <xsl:call-template name="FormatDate">
                  <xsl:with-param name="value" select="$value"/>
                  <xsl:with-param name="format" select="$format"/>
                  <xsl:with-param name="numId" select="$numId"/>
                  <xsl:with-param name="processedFormat">
                    <xsl:value-of select="substring-after($processedFormat,'mm')"/>
                  </xsl:with-param>
                  <xsl:with-param name="outputValue">
                    <xsl:value-of
                      select="concat($outputValue,substring-before(substring-after($value,'-'),'-'))"
                    />
                  </xsl:with-param>
                  <xsl:with-param name="numValue" select="$numValue"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:otherwise>
                <xsl:call-template name="FormatDate">
                  <xsl:with-param name="value" select="$value"/>
                  <xsl:with-param name="format" select="$format"/>
                  <xsl:with-param name="numId" select="$numId"/>
                  <xsl:with-param name="processedFormat">
                    <xsl:value-of select="substring-after($processedFormat,'m')"/>
                  </xsl:with-param>
                  <xsl:with-param name="outputValue">
                    <xsl:value-of
                      select="concat($outputValue,number(substring-before(substring-after($value,'-'),'-')))"
                    />
                  </xsl:with-param>
                  <xsl:with-param name="numValue" select="$numValue"/>
                </xsl:call-template>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <!-- day -->
      <xsl:when test="starts-with($processedFormat,'d')">
        <xsl:choose>
          <xsl:when test="starts-with(substring-after($processedFormat,'d'),'ddd')">
            <xsl:call-template name="FormatDate">
              <xsl:with-param name="value" select="$value"/>
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="numId" select="$numId"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'dddd')"/>
              </xsl:with-param>
              <xsl:with-param name="outputValue">
                <xsl:variable name="dayOfWeek">
                  <xsl:call-template name="ConvertDayToName">
                    <xsl:with-param name="day">
                      <xsl:value-of select="$numValue"/>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:variable>
                <xsl:value-of select="concat($outputValue,$dayOfWeek)"/>
              </xsl:with-param>
              <xsl:with-param name="numValue" select="$numValue"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="starts-with(substring-after($processedFormat,'d'),'dd')">
            <xsl:call-template name="FormatDate">
              <xsl:with-param name="value" select="$value"/>
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="numId" select="$numId"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'dddd')"/>
              </xsl:with-param>
              <xsl:with-param name="outputValue">
                <xsl:variable name="dayOfWeek">
                  <xsl:call-template name="ConvertDayToShortName">
                    <xsl:with-param name="day">
                      <xsl:value-of select="$numValue"/>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:variable>
                <xsl:value-of select="concat($outputValue,$dayOfWeek)"/>
              </xsl:with-param>
              <xsl:with-param name="numValue" select="$numValue"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="starts-with(substring-after($processedFormat,'d'),'d')">
            <xsl:call-template name="FormatDate">
              <xsl:with-param name="value" select="$value"/>
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="numId" select="$numId"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'dd')"/>
              </xsl:with-param>
              <xsl:with-param name="outputValue">
                <xsl:value-of
                  select="concat($outputValue,substring-before(substring-after(substring-after($value,'-'),'-'),'T'))"
                />
              </xsl:with-param>
              <xsl:with-param name="numValue" select="$numValue"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="FormatDate">
              <xsl:with-param name="value" select="$value"/>
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="numId" select="$numId"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'d')"/>
              </xsl:with-param>
              <xsl:with-param name="outputValue">
                <xsl:value-of
                  select="concat($outputValue,number(substring-before(substring-after(substring-after($value,'-'),'-'),'T')))"
                />
              </xsl:with-param>
              <xsl:with-param name="numValue" select="$numValue"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <!-- hours -->
      <xsl:when test="starts-with($processedFormat,'h')">
        <xsl:choose>
          <xsl:when test="starts-with(substring-after($processedFormat,'h'),'h')">
            <xsl:call-template name="FormatDate">
              <xsl:with-param name="value" select="$value"/>
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="numId" select="$numId"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'hh')"/>
              </xsl:with-param>
              <xsl:with-param name="outputValue">
                <xsl:value-of
                  select="concat($outputValue,substring-before(substring-after($value,'T'),':'))"/>
              </xsl:with-param>
              <xsl:with-param name="numValue" select="$numValue"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="FormatDate">
              <xsl:with-param name="value" select="$value"/>
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="numId" select="$numId"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'h')"/>
              </xsl:with-param>
              <xsl:with-param name="outputValue">
                <xsl:value-of
                  select="concat($outputValue,number(substring-before(substring-after($value,'T'),':')))"
                />
              </xsl:with-param>
              <xsl:with-param name="numValue" select="$numValue"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <!-- seconds -->
      <xsl:when test="starts-with($processedFormat,'s')">
        <xsl:choose>
          <xsl:when test="starts-with(substring-after($processedFormat,'s'),'s')">
            <xsl:call-template name="FormatDate">
              <xsl:with-param name="value" select="$value"/>
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="numId" select="$numId"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'ss')"/>
              </xsl:with-param>
              <xsl:with-param name="outputValue">
                <xsl:value-of
                  select="concat($outputValue,substring-after(substring-after(substring-after($value,'T'),':'),':'))"
                />
              </xsl:with-param>
              <xsl:with-param name="numValue" select="$numValue"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="FormatDate">
              <xsl:with-param name="value" select="$value"/>
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="numId" select="$numId"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'s')"/>
              </xsl:with-param>
              <xsl:with-param name="outputValue">
                <xsl:value-of
                  select="concat($outputValue,number(substring-after(substring-after(substring-after($value,'T'),':'),':')))"
                />
              </xsl:with-param>
              <xsl:with-param name="numValue" select="$numValue"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <xsl:when
        test="starts-with($processedFormat,'\') or starts-with($processedFormat,'@') or starts-with($processedFormat,';')">
        <xsl:call-template name="FormatDate">
          <xsl:with-param name="value" select="$value"/>
          <xsl:with-param name="format" select="$format"/>
          <xsl:with-param name="numId" select="$numId"/>
          <xsl:with-param name="processedFormat">
            <xsl:value-of select="substring($processedFormat,2)"/>
          </xsl:with-param>
          <xsl:with-param name="outputValue" select="$outputValue"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="string-length($processedFormat) = 0">
        <xsl:value-of select="$outputValue"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="FormatDate">
          <xsl:with-param name="value" select="$value"/>
          <xsl:with-param name="format" select="$format"/>
          <xsl:with-param name="numId" select="$numId"/>
          <xsl:with-param name="processedFormat">
            <xsl:value-of select="substring($processedFormat,2)"/>
          </xsl:with-param>
          <xsl:with-param name="outputValue">
            <xsl:value-of select="concat($outputValue,substring($processedFormat,0,2))"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="FormatTime">

    <!-- @Descripition: inserts time to cell in a correct format -->
    <!-- @Context: None -->

    <xsl:param name="value"/>
    <!-- (time) input time value -->
    <xsl:param name="format"/>
    <!-- (string) format code -->
    <xsl:param name="numId"/>
    <!-- (int) format ID -->
    <xsl:param name="processedFormat"/>
    <!-- (string) part of format code which is being processed -->
    <xsl:param name="outputValue"/>
    <!-- (time) output time value -->
    <xsl:choose>
      <xsl:when test="starts-with($processedFormat,'h')">
        <xsl:choose>
          <xsl:when test="starts-with(substring-after($processedFormat,'h'),'h')">
            <xsl:call-template name="FormatTime">
              <xsl:with-param name="value" select="$value"/>
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="numId" select="$numId"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'hh')"/>
              </xsl:with-param>
              <xsl:with-param name="outputValue">
                <xsl:value-of
                  select="concat($outputValue,substring-before(substring-after($value,'PT'),'H'))"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="FormatTime">
              <xsl:with-param name="value" select="$value"/>
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="numId" select="$numId"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'h')"/>
              </xsl:with-param>
              <xsl:with-param name="outputValue">
                <xsl:value-of
                  select="concat($outputValue,number(substring-before(substring-after($value,'PT'),'H')))"
                />
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="starts-with($processedFormat,'m')">
        <xsl:choose>
          <xsl:when test="starts-with(substring-after($processedFormat,'m'),'m')">
            <xsl:call-template name="FormatTime">
              <xsl:with-param name="value" select="$value"/>
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="numId" select="$numId"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'mm')"/>
              </xsl:with-param>
              <xsl:with-param name="outputValue">
                <xsl:value-of
                  select="concat($outputValue,substring-before(substring-after($value,'H'),'M'))"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="FormatTime">
              <xsl:with-param name="value" select="$value"/>
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="numId" select="$numId"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'m')"/>
              </xsl:with-param>
              <xsl:with-param name="outputValue">
                <xsl:value-of
                  select="concat($outputValue,number(substring-before(substring-after($value,'H'),'M')))"
                />
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="starts-with($processedFormat,'s')">
        <xsl:choose>
          <xsl:when test="starts-with(substring-after($processedFormat,'s'),'s')">
            <xsl:call-template name="FormatTime">
              <xsl:with-param name="value" select="$value"/>
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="numId" select="$numId"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'ss')"/>
              </xsl:with-param>
              <xsl:with-param name="outputValue">
                <xsl:value-of
                  select="concat($outputValue,substring-before(substring-after($value,'M'),'S'))"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="FormatTime">
              <xsl:with-param name="value" select="$value"/>
              <xsl:with-param name="format" select="$format"/>
              <xsl:with-param name="numId" select="$numId"/>
              <xsl:with-param name="processedFormat">
                <xsl:value-of select="substring-after($processedFormat,'s')"/>
              </xsl:with-param>
              <xsl:with-param name="outputValue">
                <xsl:value-of
                  select="concat($outputValue,number(substring-before(substring-after($value,'M'),'S')))"
                />
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when
        test="starts-with($processedFormat,'\') or starts-with($processedFormat,'@') or starts-with($processedFormat,';') or starts-with($processedFormat,'[') or starts-with($processedFormat,']')">
        <xsl:call-template name="FormatTime">
          <xsl:with-param name="value" select="$value"/>
          <xsl:with-param name="format" select="$format"/>
          <xsl:with-param name="numId" select="$numId"/>
          <xsl:with-param name="processedFormat">
            <xsl:value-of select="substring($processedFormat,2)"/>
          </xsl:with-param>
          <xsl:with-param name="outputValue" select="$outputValue"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="string-length($processedFormat) = 0">
        <xsl:value-of select="$outputValue"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="FormatTime">
          <xsl:with-param name="value" select="$value"/>
          <xsl:with-param name="format" select="$format"/>
          <xsl:with-param name="numId" select="$numId"/>
          <xsl:with-param name="processedFormat">
            <xsl:value-of select="substring($processedFormat,2)"/>
          </xsl:with-param>
          <xsl:with-param name="outputValue">
            <xsl:value-of select="concat($outputValue,substring($processedFormat,0,2))"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="ConvertMonthToName">

    <!-- @Descripition: gets month name from its number -->
    <!-- @Context: None -->

    <xsl:param name="month"/>
    <!-- (int) month number -->
    <xsl:choose>
      <xsl:when test="number($month) = 1">January</xsl:when>
      <xsl:when test="number($month) = 2">February</xsl:when>
      <xsl:when test="number($month) = 3">March</xsl:when>
      <xsl:when test="number($month) = 4">April</xsl:when>
      <xsl:when test="number($month) = 5">May</xsl:when>
      <xsl:when test="number($month) = 6">June</xsl:when>
      <xsl:when test="number($month) = 7">July</xsl:when>
      <xsl:when test="number($month) = 8">August</xsl:when>
      <xsl:when test="number($month) = 9">September</xsl:when>
      <xsl:when test="number($month) = 10">October</xsl:when>
      <xsl:when test="number($month) = 11">November</xsl:when>
      <xsl:when test="number($month) = 12">December</xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="ConvertMonthToShortName">

    <!-- @Descripition: gets shortcut of month name from its number -->
    <!-- @Context: None -->

    <xsl:param name="month"/>
    <!-- (int) month number -->
    <xsl:choose>
      <xsl:when test="number($month) = 1">Jan</xsl:when>
      <xsl:when test="number($month) = 2">Feb</xsl:when>
      <xsl:when test="number($month) = 3">Mar</xsl:when>
      <xsl:when test="number($month) = 4">Apr</xsl:when>
      <xsl:when test="number($month) = 5">May</xsl:when>
      <xsl:when test="number($month) = 6">Jun</xsl:when>
      <xsl:when test="number($month) = 7">Jul</xsl:when>
      <xsl:when test="number($month) = 8">Aug</xsl:when>
      <xsl:when test="number($month) = 9">Sep</xsl:when>
      <xsl:when test="number($month) = 10">Oct</xsl:when>
      <xsl:when test="number($month) = 11">Nov</xsl:when>
      <xsl:when test="number($month) = 12">Dec</xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="ConvertDayToName">

    <!-- @Descripition: gets day of week name from number of days after 1.01.1900 -->
    <!-- @Context: None -->

    <xsl:param name="day"/>
    <!-- (int) number of days -->
    <xsl:variable name="dayOfWeek">
      <xsl:value-of select="$day - 7 * floor($day div 7)"/>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$dayOfWeek = 0">Monday</xsl:when>
      <xsl:when test="$dayOfWeek = 1">Tuesday</xsl:when>
      <xsl:when test="$dayOfWeek = 2">Wednesday</xsl:when>
      <xsl:when test="$dayOfWeek = 3">Thursday</xsl:when>
      <xsl:when test="$dayOfWeek = 4">Friday</xsl:when>
      <xsl:when test="$dayOfWeek = 5">Saturday</xsl:when>
      <xsl:when test="$dayOfWeek = 6">Sunday</xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="ConvertDayToShortName">

    <!-- @Descripition: gets shortcut of day of week name from number of days after 1.01.1900 -->
    <!-- @Context: None -->

    <xsl:param name="day"/>
    <!-- (int) number of days -->
    <xsl:variable name="dayOfWeek">
      <xsl:value-of select="$day - 7 * floor($day div 7)"/>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$dayOfWeek = 0">Mon</xsl:when>
      <xsl:when test="$dayOfWeek = 1">Tue</xsl:when>
      <xsl:when test="$dayOfWeek = 2">Wed</xsl:when>
      <xsl:when test="$dayOfWeek = 3">Thu</xsl:when>
      <xsl:when test="$dayOfWeek = 4">Fri</xsl:when>
      <xsl:when test="$dayOfWeek = 5">Sat</xsl:when>
      <xsl:when test="$dayOfWeek = 6">Sun</xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="v:shape" mode="drawing">

    <!--@Description: adds shape to put note into -->
    <!--@Context: none-->

    <xsl:param name="text"/>
    <!-- (node)note text node -->
    <xsl:attribute name="draw:style-name">
      <xsl:value-of select="generate-id(.)"/>
    </xsl:attribute>
    <xsl:attribute name="draw:text-style-name">
      <xsl:value-of select="generate-id($text)"/>
    </xsl:attribute>
    <xsl:attribute name="svg:width">
      <xsl:call-template name="ConvertPoints">
        <xsl:with-param name="length">
          <xsl:value-of select="substring-before(substring-after(@style,'width:'),';')"/>
        </xsl:with-param>
        <xsl:with-param name="unit">in</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="svg:height">
      <xsl:call-template name="ConvertPoints">
        <xsl:with-param name="length">
          <xsl:value-of select="substring-before(substring-after(@style,'height:'),';')"/>
        </xsl:with-param>
        <xsl:with-param name="unit">in</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="svg:x">
      <xsl:call-template name="ConvertPoints">
        <xsl:with-param name="length">
          <xsl:value-of select="substring-before(substring-after(@style,'margin-left:'),';')"/>
        </xsl:with-param>
        <xsl:with-param name="unit">in</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="svg:y">
      <xsl:call-template name="ConvertPoints">
        <xsl:with-param name="length">
          <xsl:value-of select="substring-before(substring-after(@style,'margin-top:'),';')"/>
        </xsl:with-param>
        <xsl:with-param name="unit">in</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="draw:caption-point-x">-0.2402in</xsl:attribute>
    <xsl:attribute name="draw:caption-point-y">0in</xsl:attribute>
  </xsl:template>




</xsl:stylesheet>

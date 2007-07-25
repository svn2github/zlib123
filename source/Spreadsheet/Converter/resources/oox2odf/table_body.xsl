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
  <xsl:import href="insert_cols.xsl"/>
  <xsl:import href="date_time.xsl"/>
  <xsl:import href="insert_text.xsl"/>
  
  <xsl:key name="hyperlinkPosition" match="e:c" use="'@r'"/>
  <xsl:key name="ref" match="e:hyperlink" use="@ref"/>
  <!--xsl:key name="outlineLevelRow" match="e:sheetFormatPr" use="@outlineLevelRow"/-->
  <!--xsl:key name="outlineLevel" match="e:row" use="@outlineLevel"/-->

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
    <xsl:param name="ValidationCell"/>     
    <xsl:param name="ValidationRow"/>
    <xsl:param name="ValidationCellStyle"/>    
    <xsl:param name="sheetNr"/>

    <xsl:choose>
      <!-- when sheet is empty  -->
      <xsl:when
        test="not(e:worksheet/e:sheetData/e:row/e:c/e:v) and $BigMergeCell = '' and $BigMergeRow = '' and $PictureCell = '' and $NoteCell = '' and $ConditionalCell = '' and $ValidationCell = ''">
        <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
          table:number-rows-repeated="65536">
          <table:table-cell table:number-columns-repeated="256"/>
        </table:table-row>
      </xsl:when>
      <!-- when there are only picture, conditional and note in sheet  -->
      <xsl:when
        test="not(e:worksheet/e:sheetData/e:row/e:c/e:v) and $BigMergeCell = '' and $BigMergeRow = '' and ($PictureCell != '' or $NoteCell != '' or $ConditionalCell != '' or $ValidationCell != '')">

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
          <xsl:with-param name="ValidationCell">
            <xsl:value-of select="$ValidationCell"/>
          </xsl:with-param>
          <xsl:with-param name="ValidationRow">
            <xsl:value-of select="$ValidationRow"/>
          </xsl:with-param>
          <xsl:with-param name="ValidationCellStyle">
            <xsl:value-of select="$ValidationCellStyle"/>
          </xsl:with-param>
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
    <xsl:param name="removeFilter"/>
    <xsl:param name="ValidationRow"/>
    <xsl:param name="ValidationCell"/>
    <xsl:param name="ValidationCellStyle"/>
    <xsl:param name="ConnectionsCell"/>
    <xsl:param name="outlineLevel"/>


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

          <!-- if row is fidden and fiter is not being removed -->
          <xsl:if
            test="@hidden=1 and not(@r &gt;= substring-before($removeFilter,':') and @r &lt;= substring-after($removeFilter,':'))">
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
              <xsl:with-param name="ValidationCell">
                <xsl:value-of select="$ValidationCell"/>
              </xsl:with-param>
              <xsl:with-param name="ValidationRow">
                <xsl:value-of select="$ValidationRow"/>
              </xsl:with-param>
              <xsl:with-param name="ValidationCellStyle">
                <xsl:value-of select="$ValidationCellStyle"/>
              </xsl:with-param>
              <xsl:with-param name="ConnectionsCell">
                <xsl:value-of select="$ConnectionsCell"/>
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
              test="$lastCellColumnNumber &lt; 256 and $CheckIfBigMerge = '' and $CheckIfBigMergeAfter != 'true' and $PictureCell = '' and $ConditionalCell = '' and $NoteCell = ''">
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
              test="$lastCellColumnNumber &lt; 256 and $CheckIfBigMerge != '' and not(e:c) and $PictureCell = '' and $ConditionalCell = '' and $NoteCell = ''">
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
            <xsl:when
              test="($PictureCell != '' or $NoteCell != '' or $ConditionalCell != '') and $GetMinRowWithElements=@r and not(e:c)">
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
                <xsl:with-param name="ConditionalCell">
                  <xsl:value-of select="$ConditionalCell"/>
                </xsl:with-param>
                <xsl:with-param name="ConditionalCellStyle">
                  <xsl:value-of select="$ConditionalCellStyle"/>
                </xsl:with-param>
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
                <xsl:with-param name="ValidationCell">
                  <xsl:value-of select="$ValidationCell"/>
                </xsl:with-param>
                <xsl:with-param name="ValidationRow">
                  <xsl:value-of select="$ValidationRow"/>
                </xsl:with-param>
                <xsl:with-param name="ValidationCellStyle">
                  <xsl:value-of select="$ValidationCellStyle"/>
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
    <xsl:param name="ValidationCell"/>     
    <xsl:param name="ValidationRow"/>
    <xsl:param name="ValidationCellStyle"/>
    <xsl:param name="ValidationColl"/>
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
      <xsl:value-of select="concat($PictureColl, $NoteColl, $ValidationColl)"/>
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
          <xsl:with-param name="ValidationCell">
            <xsl:value-of select="$ValidationCell"/>
          </xsl:with-param>
          <xsl:with-param name="ValidationRow">
            <xsl:value-of select="$ValidationRow"/>
          </xsl:with-param>
          <xsl:with-param name="ValidationCellStyle">
            <xsl:value-of select="$ValidationCellStyle"/>
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
          <xsl:with-param name="ValidationCell">
            <xsl:value-of select="$ValidationCell"/>
          </xsl:with-param>
          <xsl:with-param name="ValidationRow">
            <xsl:value-of select="$ValidationRow"/>
          </xsl:with-param>
          <xsl:with-param name="ValidationCellStyle">
            <xsl:value-of select="$ValidationCellStyle"/>
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
    <xsl:param name="ValidationCell"/>
    <xsl:param name="ValidationRow"/>
    <xsl:param name="ValidationCellStyle"/>
    <xsl:param name="ConnectionsCell"/>

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
          <xsl:with-param name="ValidationCell">
            <xsl:value-of select="$ValidationCell"/>
          </xsl:with-param>
          <xsl:with-param name="ValidationRow">
            <xsl:value-of select="$ValidationRow"/>
          </xsl:with-param>
          <xsl:with-param name="ValidationCellStyle">
            <xsl:value-of select="$ValidationCellStyle"/>
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


          <!-- check if conditional -->
          <xsl:if
            test="@s or contains(concat(';', $ConditionalCell), concat(';', $rowNum, ':', $colNum, ';'))">
            <xsl:choose>

              <xsl:when
                test="@s and contains(concat(';', $ConditionalCell), concat(';', $rowNum, ':', $colNum, ';'))">
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

          <!-- chceck if DataValidation -->

          <xsl:if
            test="contains(concat(';', $ValidationCell), concat(';', $rowNum, ':', $colNum, ';'))">
            <xsl:attribute name="table:content-validation-name">
              <xsl:value-of select="(concat('val', $sheetNr, (. + 1)))"/>
            </xsl:attribute>
            
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

          <!-- Insert Connections Cell  -->
          <xsl:if
            test="contains(concat(';', $ConnectionsCell), concat(';', $rowNum, ':', $colNum, '-'))">
            <xsl:call-template name="InsertConnections">
              <xsl:with-param name="rowNum">
                <xsl:value-of select="$rowNum"/>
              </xsl:with-param>
              <xsl:with-param name="colNum">
                <xsl:value-of select="$colNum"/>
              </xsl:with-param>
              <xsl:with-param name="ConnecionsCell">
                <xsl:value-of select="$ConnectionsCell"/>
              </xsl:with-param>
              <xsl:with-param name="sheetNr">
                <xsl:value-of select="$sheetNr"/>
              </xsl:with-param>
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
              <xsl:with-param name="ValidationCell">
                <xsl:value-of select="$ValidationCell"/>
              </xsl:with-param>
              <xsl:with-param name="ValidationRow">
                <xsl:value-of select="$ValidationRow"/>
              </xsl:with-param>
              <xsl:with-param name="ValidationCellStyle">
                <xsl:value-of select="$ValidationCellStyle"/>
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
    <xsl:param name="ConnectionsCell"/>
    <xsl:param name="ValidationCell"/>     
    <xsl:param name="ValidationRow"/>
    <xsl:param name="ValidationCellStyle"/>
    <xsl:param name="ValidationColl"/>


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
      <xsl:value-of select="concat($PictureColl, $NoteColl, $ValidationColl)"/>
    </xsl:variable>
    
    <xsl:variable name="GetMinCollAfterThisCell">
      <xsl:call-template name="GetMinRowWithPicture">
        <xsl:with-param name="PictureRow">
          <xsl:value-of select="concat($PictureColl, $NoteColl, $ValidationColl)"/>
        </xsl:with-param>
        <xsl:with-param name="AfterRow">
          <xsl:value-of select="$colNum + 1"/>
        </xsl:with-param>
      </xsl:call-template>
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
              <xsl:with-param name="ConnecionsCell">
                <xsl:value-of select="$ConnectionsCell"/>
              </xsl:with-param>
              <xsl:with-param name="ValidationCell">
                <xsl:value-of select="$ValidationCell"/>
              </xsl:with-param>
              <xsl:with-param name="ValidationRow">
                <xsl:value-of select="$ValidationRow"/>
              </xsl:with-param>
              <xsl:with-param name="ValidationCellStyle">
                <xsl:value-of select="$ValidationCellStyle"/>
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
                <xsl:with-param name="ConnecionsCell">
                  <xsl:value-of select="$ConnectionsCell"/>
                </xsl:with-param>
                <xsl:with-param name="ValidationCell">
                  <xsl:value-of select="$ValidationCell"/>
                </xsl:with-param>
                <xsl:with-param name="ValidationRow">
                  <xsl:value-of select="$ValidationRow"/>
                </xsl:with-param>
                <xsl:with-param name="ValidationCellStyle">
                  <xsl:value-of select="$ValidationCellStyle"/>
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
                <xsl:with-param name="ConnecionsCell">
                  <xsl:value-of select="$ConnectionsCell"/>
                </xsl:with-param>
                <xsl:with-param name="ValidationCell">
                  <xsl:value-of select="$ValidationCell"/>
                </xsl:with-param>
                <xsl:with-param name="ValidationRow">
                  <xsl:value-of select="$ValidationRow"/>
                </xsl:with-param>
                <xsl:with-param name="ValidationCellStyle">
                  <xsl:value-of select="$ValidationCellStyle"/>
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
            <xsl:with-param name="ConnecionsCell">
              <xsl:value-of select="$ConnectionsCell"/>
            </xsl:with-param>
            <xsl:with-param name="ValidationCell">
              <xsl:value-of select="$ValidationCell"/>
            </xsl:with-param>
            <xsl:with-param name="ValidationRow">
              <xsl:value-of select="$ValidationRow"/>
            </xsl:with-param>
            <xsl:with-param name="ValidationCellStyle">
              <xsl:value-of select="$ValidationCellStyle"/>
            </xsl:with-param>
          </xsl:apply-templates>
        </xsl:if>
      </xsl:when>

      <xsl:when
        test="not(following-sibling::e:c) and ($PictureCell != '' or $NoteCell != '' or $ValidationCell != '') and $GetMinCollAfterThisCell &gt; $colNum">

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
          <xsl:with-param name="ConnecionsCell">
            <xsl:value-of select="$ConnectionsCell"/>
          </xsl:with-param>
          <xsl:with-param name="ValidationCell">
            <xsl:value-of select="$ValidationCell"/>
          </xsl:with-param>
          <xsl:with-param name="ValidationRow">
            <xsl:value-of select="$ValidationRow"/>
          </xsl:with-param>
          <xsl:with-param name="ValidationCellStyle">
            <xsl:value-of select="$ValidationCellStyle"/>
          </xsl:with-param>
          <xsl:with-param name="ValidationColl">
            <xsl:value-of select="$ValidationColl"/>
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
            <xsl:with-param name="ConnectionsCell">
              <xsl:value-of select="$ConnectionsCell"/>
            </xsl:with-param>
            <xsl:with-param name="ValidationCell">
              <xsl:value-of select="$ValidationCell"/>
            </xsl:with-param>
            <xsl:with-param name="ValidationRow">
              <xsl:value-of select="$ValidationRow"/>
            </xsl:with-param>
            <xsl:with-param name="ValidationCellStyle">
              <xsl:value-of select="$ValidationCellStyle"/>
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

    <xsl:variable name="formatingMarks">
      <xsl:call-template name="StripText">
        <xsl:with-param name="formatCode" select="$formatCode"/>
      </xsl:call-template>
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
                <xsl:when test="contains($formatingMarks,'_')">
                  <xsl:value-of select="substring-before($formatingMarks,'_')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$formatingMarks"/>
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
      <xsl:when test="contains($formatingMarks,'[$') or contains($formatingMarks,'zł')">
        <xsl:variable name="currency">
          <xsl:choose>
            <xsl:when test="contains($formatCode,'zł')">zł</xsl:when>
            <xsl:when test="contains($formatCode,'Red')">
              <xsl:variable name="tempFormat">
                <xsl:value-of
                  select="substring-after(substring-before(substring-after($formatingMarks,'Red]'),']'),'[$')"
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
                <xsl:value-of select="substring-after(substring-before($formatingMarks,']'),'[$')"/>
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
            test="contains(substring-before($formatingMarks,$currency),'0') or contains(substring-before($formatingMarks,$currency),'#')">
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

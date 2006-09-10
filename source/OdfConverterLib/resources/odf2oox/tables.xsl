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
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/3/main"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  exclude-result-prefixes="office text table fo style draw xlink v svg"
  xmlns:w10="urn:schemas-microsoft-com:office:word">


  <!-- tables -->
  <xsl:template match="table:table">
    <xsl:variable name="styleName">
      <xsl:value-of select="@table:style-name"/>
    </xsl:variable>
    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="{@table:style-name}"/>
        <xsl:variable name="tableProp"
          select="key('automatic-styles', @table:style-name)/style:table-properties"/>
        <w:tblW w:type="{$type}">
          <xsl:attribute name="w:w">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length"
                select="key('automatic-styles', @table:style-name)/style:table-properties/@style:width"
              />
            </xsl:call-template>
          </xsl:attribute>
        </w:tblW>
        <xsl:if
          test="key('automatic-styles', @table:style-name)/style:table-properties/@table:align">
          <xsl:choose>
            <xsl:when
              test="key('automatic-styles', @table:style-name)/style:table-properties/@table:align = 'margins'">
              <w:jc w:val="left"/>
              <!--User agents that do not support the "margins" value, may treat this value as "left".-->
            </xsl:when>
            <xsl:otherwise>
              <w:jc
                w:val="{key('automatic-styles', @table:style-name)/style:table-properties/@table:align}"
              />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
        <xsl:if
          test="key('automatic-styles', @table:style-name)/style:table-properties/@fo:margin-left != '' ">
          <w:tblInd w:type="{$type}">
            <xsl:attribute name="w:w">
              <xsl:call-template name="twips-measure">
                <xsl:with-param name="length"
                  select="key('automatic-styles', @table:style-name)/style:table-properties/@fo:margin-left"
                />
              </xsl:call-template>
            </xsl:attribute>
          </w:tblInd>
        </xsl:if>
        <!-- Default layout algorithm in ODF is "fixed". -->
        <w:tblLayout w:type="fixed"/>

        <!--table background-->
        <xsl:if test="$tableProp/@fo:background-color">
          <xsl:choose>
            <xsl:when test="$tableProp/@fo:background-color != 'transparent' ">
              <w:shd w:val="clear" w:color="auto"
                w:fill="{substring($tableProp/@fo:background-color, 2, string-length($tableProp/@fo:background-color) -1)}"
              />
            </xsl:when>
          </xsl:choose>
        </xsl:if>

      </w:tblPr>
      <xsl:call-template name="InsertTblGrid"/>
      <xsl:apply-templates select="table:table-rows|table:table-header-rows|table:table-row"/>
    </w:tbl>




    <!-- Section detection  : 3 cases -->
    <xsl:if test="not(ancestor::table:table)">
      <!-- 1 - Following neighbour's (ie paragraph, heading or table) master style  -->
      <xsl:variable name="followings"
        select="following-sibling::text:p | following-sibling::text:h | following-sibling::table:table"/>
      <xsl:variable name="masterPageStarts"
        select="boolean(key('master-based-styles', $followings[1]/@text:style-name | $followings[1]/@table:style-name))"/>


      <!-- section creation -->
      <xsl:if
        test="($masterPageStarts = 'true') and not(ancestor::text:note-body) and not(ancestor::table:table)">
        <w:p>
          <w:pPr>
            <w:sectPr>

              <!-- Determine the master style that rules this section -->
              <xsl:variable name="currentMasterStyle"
                select="key('master-based-styles', @text:style-name)"/>

              <xsl:choose>
                <xsl:when test="boolean($currentMasterStyle)">
                  <!-- current element style is tied to a master page -->
                  <xsl:for-each select="document('styles.xml')">
                    <xsl:apply-templates
                      select="key('page-layouts', key('master-pages', $currentMasterStyle[1]/@style:master-page-name)/@style:page-layout-name)/style:page-layout-properties"
                      mode="master-page"/>
                  </xsl:for-each>
                </xsl:when>
                <xsl:otherwise>
                  <!-- current style is not tied to a master page, find the preceding one -->
                  <xsl:variable name="precedings"
                    select="preceding::text:p[key('master-based-styles', @text:style-name)] | preceding::text:h[key('master-based-styles', @text:style-name)] | preceding::table:table[key('master-based-styles', @table:style-name)]"/>
                  <xsl:variable name="precedingMasterStyle"
                    select="key('master-based-styles', $precedings[last()]/@text:style-name | $precedings[last()]/@table:style-name)"/>
                  <xsl:choose>
                    <xsl:when test="boolean($precedingMasterStyle)">
                      <xsl:for-each select="document('styles.xml')">
                        <xsl:apply-templates
                          select="key('page-layouts', key('master-pages', $precedingMasterStyle[1]/@style:master-page-name)/@style:page-layout-name)/style:page-layout-properties"
                          mode="master-page"/>
                      </xsl:for-each>
                    </xsl:when>
                    <xsl:otherwise>
                      <!-- otherwise, apply the default master style -->
                      <xsl:for-each select="document('styles.xml')">
                        <xsl:apply-templates
                          select="key('page-layouts', /office:document-styles/office:master-styles/style:master-page[1]/@style:page-layout-name)/style:page-layout-properties"
                          mode="master-page"/>
                      </xsl:for-each>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:otherwise>
              </xsl:choose>

            </w:sectPr>
          </w:pPr>
        </w:p>
      </xsl:if>
    </xsl:if>



  </xsl:template>

  <!-- COMMENT: please rename the template and add a description -->
  <xsl:template name="subtable">
    <xsl:param name="node"/>
    <xsl:for-each select="$node/table:table-cell">
      <xsl:call-template name="table-cell"/>
    </xsl:for-each>
  </xsl:template>

  <!-- COMMENT: please rename the template and add a description -->
  <xsl:template name="merged-rows">
    <xsl:param name="i" select="0"/>
    <xsl:param name="iterator"/>
    <!-- COMMENT: is this variable really necessary? -->
    <xsl:variable name="test">
      <xsl:if test="$i > 0">
        <xsl:text>true</xsl:text>
      </xsl:if>
    </xsl:variable>
    <xsl:if test="$test='true'">
      <w:tr>
        <xsl:for-each select="table:table-cell">
          <xsl:choose>
            <xsl:when test="table:table[@table:is-sub-table='true']">
              <!-- table to process -->
              <xsl:call-template name="subtable">
                <xsl:with-param name="node" select="table:table/child::table:table-row[$iterator]"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="@table:number-columns-spanned">
              <xsl:choose>
                <xsl:when test="$iterator = 1">
                  <xsl:call-template name="table-cell">
                    <xsl:with-param name="grid"
                      select="round(number(@table:number-columns-spanned))"/>
                    <xsl:with-param name="merge" select="1"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:call-template name="table-cell">
                    <xsl:with-param name="grid"
                      select="round(number(@table:number-columns-spanned))"/>
                    <xsl:with-param name="merge" select="2"/>
                  </xsl:call-template>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when test="$iterator = 1">
                  <xsl:call-template name="table-cell">
                    <xsl:with-param name="merge" select="1"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:call-template name="table-cell">
                    <xsl:with-param name="merge" select="2"/>
                  </xsl:call-template>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </w:tr>
      <xsl:call-template name="merged-rows">
        <xsl:with-param name="i" select="$i  -1"/>
        <xsl:with-param name="iterator" select="$iterator +1"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

  <!--************************
      ** Compute table grid **
      ************************-->

  <!-- Inserts a table grid -->
  <xsl:template name="InsertTblGrid">
    <xsl:variable name="unit">
      <xsl:call-template name="GetUnit">
        <xsl:with-param name="length"
          select="key('automatic-styles',@table:style-name)/style:table-properties/@style:width"/>
      </xsl:call-template>
    </xsl:variable>
    <w:tblGrid>
      <xsl:call-template name="InsertGridCols">
        <xsl:with-param name="lastWidth">0</xsl:with-param>
        <xsl:with-param name="unit" select="$unit"/>
      </xsl:call-template>
    </w:tblGrid>
  </xsl:template>

  <!-- Inserts grid columns -->
  <xsl:template name="InsertGridCols">
    <xsl:param name="unit"/>
    <xsl:param name="lastWidth"/>
    <xsl:variable name="tableWidth">
      <xsl:value-of
        select="substring-before(key('automatic-styles',@table:style-name)/style:table-properties/@style:width,$unit)"
      />
    </xsl:variable>
    <!-- We continue until we've reached end of table (with a small incertitude of 0.005) -->
    <xsl:if test="$tableWidth - $lastWidth > 0.005 or $tableWidth - $lastWidth &lt; -0.005">
      <xsl:variable name="nextWidth">
        <xsl:choose>
          <xsl:when test="table:table-header-rows/table:table-row">
            <xsl:apply-templates select="table:table-header-rows/table:table-row[1]"
              mode="findNextWidth">
              <xsl:with-param name="unit" select="$unit"/>
              <xsl:with-param name="lastWidth" select="$lastWidth"/>
              <xsl:with-param name="nextWidth" select="$tableWidth"/>
              <xsl:with-param name="offset">0</xsl:with-param>
            </xsl:apply-templates>
          </xsl:when>
          <xsl:otherwise>
            <xsl:apply-templates select="table:table-row[1]" mode="findNextWidth">
              <xsl:with-param name="unit" select="$unit"/>
              <xsl:with-param name="lastWidth" select="$lastWidth"/>
              <xsl:with-param name="nextWidth" select="$tableWidth"/>
              <xsl:with-param name="offset">0</xsl:with-param>
            </xsl:apply-templates>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <w:gridCol>
        <xsl:attribute name="w:w">
          <xsl:call-template name="twips-measure">
            <xsl:with-param name="length" select="concat($nextWidth - $lastWidth,$unit)"/>
          </xsl:call-template>
        </xsl:attribute>
      </w:gridCol>
      <xsl:call-template name="InsertGridCols">
        <xsl:with-param name="unit" select="$unit"/>
        <xsl:with-param name="lastWidth" select="$nextWidth"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

  <!-- iterate over every row to find the next width -->
  <!-- the next width is the smallest value that matches the following conditions:
        - it is bigger than the last width
        - a cell is starting at this value -->
  <xsl:template match="table:table-row" mode="findNextWidth">
    <xsl:param name="unit"/>
    <xsl:param name="lastWidth"/>
    <xsl:param name="nextWidth"/>
    <xsl:param name="offset"/>
    <!-- compute next width for this row -->
    <xsl:variable name="rowNextWidth">
      <xsl:apply-templates select="table:table-cell[1]" mode="findNextWidth">
        <xsl:with-param name="unit" select="$unit"/>
        <xsl:with-param name="lastWidth" select="$lastWidth"/>
        <xsl:with-param name="currentWidth" select="$offset"/>
      </xsl:apply-templates>
    </xsl:variable>
    <!-- take the smaller new width between previous rows' and current row's -->
    <xsl:variable name="computedNextWidth">
      <xsl:choose>
        <xsl:when test="$rowNextWidth &lt; $nextWidth">
          <xsl:value-of select="$rowNextWidth"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$nextWidth"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <!-- iterate over next row -->
      <xsl:when test="following-sibling::table:table-row">
        <xsl:apply-templates select="following-sibling::table:table-row[1]" mode="findNextWidth">
          <xsl:with-param name="unit" select="$unit"/>
          <xsl:with-param name="lastWidth" select="$lastWidth"/>
          <xsl:with-param name="nextWidth" select="$computedNextWidth"/>
          <xsl:with-param name="offset" select="$offset"/>
        </xsl:apply-templates>
      </xsl:when>
      <!-- if we are in table header rows, go to normal rows -->
      <xsl:when test="parent::table:table-header-rows and ../../table:table-row">
        <xsl:apply-templates select="../../table:table-row[1]" mode="findNextWidth">
          <xsl:with-param name="unit" select="$unit"/>
          <xsl:with-param name="lastWidth" select="$lastWidth"/>
          <xsl:with-param name="nextWidth" select="$computedNextWidth"/>
          <xsl:with-param name="offset" select="$offset"/>
        </xsl:apply-templates>
      </xsl:when>
      <!-- end of table, use the computed value -->
      <xsl:otherwise>
        <xsl:value-of select="$computedNextWidth"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- iterate over every cell to find the next width -->
  <!-- the next width is the smallest value that matches the following conditions:
        - it is bigger than the last width
        - a cell is starting at this value -->
  <xsl:template match="table:table-cell" mode="findNextWidth">
    <xsl:param name="unit"/>
    <xsl:param name="lastWidth"/>
    <xsl:param name="currentWidth"/>
    <xsl:variable name="cellWidth">
      <xsl:call-template name="GetCellWidth">
        <xsl:with-param name="unit" select="$unit"/>
      </xsl:call-template>
    </xsl:variable>
    <!-- $nextWidth: value at which the current cell ends -->
    <!-- we have to do this round(x*10000) div 10000 to avoid decimal artifacts -->
    <xsl:variable name="nextWidth" select="round(($currentWidth + $cellWidth)*10000) div 10000"/>
    <xsl:choose>
      <!-- current cell is starting before last width, iterate over the next cell -->
      <xsl:when test="$nextWidth &lt;= $lastWidth">
        <xsl:apply-templates select="following-sibling::table:table-cell[1]" mode="findNextWidth">
          <xsl:with-param name="unit" select="$unit"/>
          <xsl:with-param name="lastWidth" select="$lastWidth"/>
          <xsl:with-param name="currentWidth" select="$nextWidth"/>
        </xsl:apply-templates>
      </xsl:when>
      <xsl:otherwise>
        <!-- current cell is starting after last width -->
        <xsl:choose>
          <!-- if there is a subtable, iterate over every row of it -->
          <xsl:when test="child::table:table[@table:is-sub-table = 'true']">
            <xsl:choose>
              <xsl:when test="child::table:table/table:table-header-rows/table:table-row">
                <xsl:apply-templates
                  select="child::table:table/table:table-header-rows/table:table-row[1]"
                  mode="findNextWidth">
                  <xsl:with-param name="unit" select="$unit"/>
                  <xsl:with-param name="lastWidth" select="$lastWidth"/>
                  <xsl:with-param name="nextWidth" select="$nextWidth"/>
                  <xsl:with-param name="offset" select="$currentWidth"/>
                </xsl:apply-templates>
              </xsl:when>
              <xsl:otherwise>
                <xsl:apply-templates select="child::table:table/table:table-row[1]"
                  mode="findNextWidth">
                  <xsl:with-param name="unit" select="$unit"/>
                  <xsl:with-param name="lastWidth" select="$lastWidth"/>
                  <xsl:with-param name="nextWidth" select="$nextWidth"/>
                  <xsl:with-param name="offset" select="$currentWidth"/>
                </xsl:apply-templates>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <!-- no subtable, use this cell's ending width -->
          <xsl:otherwise>
            <xsl:value-of select="$nextWidth"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Gets the width of the current cell -->
  <xsl:template name="GetCellWidth">
    <xsl:param name="unit">
      <xsl:call-template name="GetUnit">
        <xsl:with-param name="length"
          select="key('automatic-styles',ancestor::table:table[1]/@table:style-name)/style:table-properties/@style:width"
        />
      </xsl:call-template>
    </xsl:param>
    <xsl:param name="colNumber">
      <xsl:call-template name="GetColumnNumber"/>
    </xsl:param>
    <xsl:param name="colSpan">
      <xsl:choose>
        <xsl:when test="@table:number-columns-spanned">
          <xsl:value-of select="@table:number-columns-spanned"/>
        </xsl:when>
        <xsl:otherwise>1</xsl:otherwise>
      </xsl:choose>
    </xsl:param>
    <xsl:param name="currentColNumber">1</xsl:param>
    <xsl:param name="currentPosInColumns">1</xsl:param>
    <xsl:variable name="rangeColNumber">
      <xsl:choose>
        <xsl:when
          test="ancestor::table:table[1]/table:table-column[position() = $currentPosInColumns]/@table:number-columns-repeated">
          <xsl:value-of
            select="ancestor::table:table[1]/table:table-column[position() = $currentPosInColumns]/@table:number-columns-repeated"
          />
        </xsl:when>
        <xsl:otherwise>1</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$colNumber &lt; ($currentColNumber + $rangeColNumber)">
        <xsl:variable name="currentColumnWidth">
          <xsl:value-of
            select="substring-before(key('automatic-styles',ancestor::table:table[1]/table:table-column[position() = $currentPosInColumns]/@table:style-name)/style:table-column-properties/@style:column-width,$unit)"
          />
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="($colNumber + $colSpan) &lt;= ($currentColNumber + $rangeColNumber)">
            <xsl:value-of select="$currentColumnWidth * $colSpan"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:variable name="remainingCellWidth">
              <xsl:call-template name="GetCellWidth">
                <xsl:with-param name="unit" select="$unit"/>
                <xsl:with-param name="colNumber" select="$currentColNumber + $rangeColNumber"/>
                <xsl:with-param name="colSpan"
                  select="$colNumber + $colSpan - $currentColNumber - $rangeColNumber"/>
                <xsl:with-param name="currentColNumber" select="$currentColNumber + $rangeColNumber"/>
                <xsl:with-param name="currentPosInColumns" select="$currentPosInColumns + 1"/>
              </xsl:call-template>
            </xsl:variable>
            <!-- we have to do this round(x*10000) div 10000 to avoid decimal artifacts -->
            <xsl:value-of
              select="round(($remainingCellWidth + ($currentColNumber + $rangeColNumber - $colNumber) * $currentColumnWidth) * 10000) div 10000"
            />
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="GetCellWidth">
          <xsl:with-param name="unit" select="$unit"/>
          <xsl:with-param name="colNumber" select="$colNumber"/>
          <xsl:with-param name="colSpan" select="$colSpan"/>
          <xsl:with-param name="currentColNumber" select="$currentColNumber + $rangeColNumber"/>
          <xsl:with-param name="currentPosInColumns" select="$currentPosInColumns + 1"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Gets the column number of the current cell -->
  <xsl:template name="GetColumnNumber">
    <xsl:param name="currentCellNumber">1</xsl:param>
    <xsl:choose>
      <xsl:when test="$currentCellNumber = count(preceding-sibling::table:table-cell) + 1">1</xsl:when>
      <xsl:otherwise>
        <xsl:variable name="currentColSpan">
          <xsl:choose>
            <xsl:when
              test="preceding-sibling::table:table-cell[position() = $currentCellNumber]/@table:number-columns-spanned">
              <xsl:value-of
                select="preceding-sibling::table:table-cell[position() = $currentCellNumber]/@table:number-columns-spanned"
              />
            </xsl:when>
            <xsl:otherwise>1</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="remainingCols">
          <xsl:call-template name="GetColumnNumber">
            <xsl:with-param name="currentCellNumber" select="$currentCellNumber + 1"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:value-of select="$currentColSpan + $remainingCols"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- table rows -->
  <xsl:template match="table:table-row|table:table-header-row">
    <xsl:choose>
      <xsl:when test="table:table-cell/table:table/@table:is-sub-table='true'">
        <!-- merged cells -->
        <xsl:variable name="total_rows"
          select="count(table:table-cell/table:table[@table:is-sub-table='true']/table:table-row)"/>
        <xsl:variable name="subtables"
          select="count(table:table-cell/table:table[@table:is-sub-table='true'])"/>
        <xsl:call-template name="merged-rows">
          <xsl:with-param name="i" select="$total_rows div $subtables"/>
          <xsl:with-param name="iterator" select="1"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <w:tr>
          <xsl:if
            test="key('automatic-styles',child::table:table-cell/@table:style-name)/style:table-cell-properties/@fo:wrap-option='no-wrap'">
            <!-- Override layout algorithm -->
            <w:tblPrEx>
              <w:tblLayout w:type="auto"/>
            </w:tblPrEx>
          </xsl:if>
          <w:trPr>
            <xsl:if test="parent::table:table-header-rows">
              <w:tblHeader/>
            </xsl:if>
            <xsl:if test="@table:style-name">
              <xsl:variable name="rowStyle" select="@table:style-name"/>
              <xsl:variable name="widthType"
                select="key('automatic-styles', $rowStyle)/style:table-row-properties"/>

              <xsl:variable name="widthRow">
                <xsl:choose>
                  <xsl:when test="$widthType[@style:row-height]">
                    <xsl:value-of select="$widthType/@style:row-height"/>
                  </xsl:when>
                  <xsl:when test="$widthType[@style:min-row-height]">
                    <xsl:value-of select="$widthType/@style:min-row-height"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="0"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>

              <w:trHeight>
                <xsl:attribute name="w:val">
                  <xsl:call-template name="twips-measure">
                    <xsl:with-param name="length" select="$widthRow"/>
                  </xsl:call-template>
                </xsl:attribute>
                <xsl:attribute name="w:hRule">
                  <xsl:choose>
                    <xsl:when test="$widthType[@style:row-height]">
                      <xsl:value-of select="'exact'"/>
                    </xsl:when>
                    <xsl:when test="$widthType[@style:min-row-height]">
                      <xsl:value-of select="'atLeast'"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="'auto'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
              </w:trHeight>
            </xsl:if>
            <!-- row styles -->

            <!--keep together-->
            <xsl:if
              test="key('automatic-styles', @table:style-name)/style:table-row-properties/@style:keep-together = 'false'
              or key('automatic-styles', ancestor::table:table[@table:style-name][1]/@table:style-name)/style:table-properties/@style:may-break-between-rows='false'">
              <w:cantSplit/>
            </xsl:if>

          </w:trPr>
          <xsl:apply-templates
            select="table:table-cell[position() &lt; 64]|table:covered-table-cell">
            <xsl:with-param name="colsNumber" select="count(table:table-cell)"/>
          </xsl:apply-templates>
        </w:tr>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- table cells -->
  <xsl:template match="table:table-cell" name="table-cell">
    <xsl:param name="colsNumber"/>
    <xsl:param name="grid" select="0"/>
    <xsl:param name="merge" select="0"/>
    <w:tc>
      <w:tcPr>
        <!-- point on the cell style properties -->
        <xsl:variable name="cellProp"
          select="key('automatic-styles', @table:style-name)/style:table-cell-properties"/>
        <xsl:variable name="tableStyle"
          select="ancestor::table:table[@table:style-name][1]/@table:style-name"/>
        <xsl:variable name="rowStyle" select="../@table:style-name"/>
        <xsl:variable name="tableProp"
          select="key('automatic-styles', $tableStyle)/style:table-properties"/>
        <xsl:variable name="rowProp"
          select="key('automatic-styles', $rowStyle)/style:table-row-properties"/>

        <!-- cell width -->
        <w:tcW w:type="{$type}">
          <xsl:attribute name="w:w">
            <xsl:variable name="unit">
              <xsl:call-template name="GetUnit">
                <xsl:with-param name="length" select="$tableProp/@style:width"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="length">
              <xsl:call-template name="GetCellWidth">
                <xsl:with-param name="unit" select="$unit"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length" select="concat($length,$unit)"/>
            </xsl:call-template>
          </xsl:attribute>
        </w:tcW>

        <!-- cell span -->
        <xsl:choose>
          <xsl:when test="$merge = 1">
            <w:gridSpan w:val="{$grid}"/>
            <w:vmerge w:val="restart"/>
          </xsl:when>
          <xsl:when test="$merge = 2">
            <w:gridSpan w:val="{$grid}"/>
            <w:vmerge w:val="continue"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:variable name="numberOfColumns">
              <xsl:choose>
                <xsl:when
                  test="ancestor::table:table[1]/descendant::table:table-row[table:table-cell/table:table[@table:is-sub-table='true']]">
                  <!-- TODO : Finish compute the span value of the current cell -->
                  <xsl:call-template name="ComputeSpanValue"/>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:choose>
              <xsl:when test="@table:number-columns-spanned">
                <w:gridSpan w:val="{@table:number-columns-spanned}"/>
              </xsl:when>
              <xsl:otherwise>
                <w:gridSpan w:val="{$numberOfColumns}"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>

        <w:tcBorders>
          <xsl:choose>
            <xsl:when test="$cellProp[@fo:border and @fo:border!='none' ]">
              <xsl:variable name="border" select="$cellProp/@fo:border"/>
              <!-- fo:border = "0.002cm solid #000000" -->
              <xsl:variable name="border-color" select="substring-after($border, '#')"/>
              <xsl:variable name="border-size">
                <xsl:call-template name="eightspoint-measure">
                  <xsl:with-param name="length" select="substring-before($border,' ')"/>
                </xsl:call-template>
              </xsl:variable>
              <w:top w:val="single" w:color="{$border-color}" w:sz="{$border-size}"/>
              <w:left w:val="single" w:color="{$border-color}" w:sz="{$border-size}"/>
              <w:bottom w:val="single" w:color="{$border-color}" w:sz="{$border-size}"/>
              <w:right w:val="single" w:color="{$border-color}" w:sz="{$border-size}"/>
            </xsl:when>

            <xsl:otherwise>
              <xsl:if test="$cellProp[@fo:border-top and @fo:border-top != 'none']">
                <xsl:variable name="border" select="$cellProp/@fo:border-top"/>
                <w:top w:val="single" w:color="{substring-after($border, '#')}">
                  <xsl:attribute name="w:sz">
                    <xsl:call-template name="eightspoint-measure">
                      <xsl:with-param name="length" select="substring-before($border, ' ')"/>
                    </xsl:call-template>
                  </xsl:attribute>
                </w:top>
              </xsl:if>
              <xsl:if test="$cellProp[@fo:border-left and @fo:border-left != 'none']">
                <xsl:variable name="border" select="$cellProp/@fo:border-left"/>
                <w:left w:val="single" w:color="{substring-after($border, '#')}">
                  <xsl:attribute name="w:sz">
                    <xsl:call-template name="eightspoint-measure">
                      <xsl:with-param name="length" select="substring-before($border, ' ')"/>
                    </xsl:call-template>
                  </xsl:attribute>
                </w:left>
              </xsl:if>
              <xsl:if test="$cellProp[@fo:border-bottom and @fo:border-bottom != 'none']">
                <xsl:variable name="border" select="$cellProp/@fo:border-bottom"/>
                <w:bottom w:val="single" w:color="{substring-after($border, '#')}">
                  <xsl:attribute name="w:sz">
                    <xsl:call-template name="eightspoint-measure">
                      <xsl:with-param name="length" select="substring-before($border, ' ')"/>
                    </xsl:call-template>
                  </xsl:attribute>
                </w:bottom>
              </xsl:if>
              <xsl:if
                test="$cellProp[(@fo:border-right and @fo:border-right != 'none')] or (position() &lt; $colsNumber and position() = 63)">
                <xsl:variable name="border">
                  <xsl:choose>
                    <xsl:when test="position() &lt; $colsNumber and position() = 63">
                      <xsl:value-of select="$cellProp/@fo:border-left"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$cellProp/@fo:border-right"/>
                    </xsl:otherwise>
                  </xsl:choose>

                </xsl:variable>
                <w:right w:val="single" w:color="{substring-after($border, '#')}">
                  <xsl:attribute name="w:sz">
                    <xsl:call-template name="eightspoint-measure">
                      <xsl:with-param name="length" select="substring-before($border, ' ')"/>
                    </xsl:call-template>
                  </xsl:attribute>
                </w:right>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
        </w:tcBorders>

        <!--cell background color-->
        <xsl:choose>
          <xsl:when
            test="$cellProp/@fo:background-color and $cellProp/@fo:background-color != 'transparent' ">
            <w:shd w:val="clear" w:color="auto"
              w:fill="{substring($cellProp/@fo:background-color, 2, string-length($cellProp/@fo:background-color) -1)}"
            />
          </xsl:when>

          <xsl:otherwise>
            <xsl:choose>
              <xsl:when
                test="$rowProp/@fo:background-color and $rowProp/@fo:background-color != 'transparent' ">
                <w:shd w:val="clear" w:color="auto"
                  w:fill="{substring($rowProp/@fo:background-color, 2, string-length($rowProp/@fo:background-color) -1)}"
                />
              </xsl:when>
              <xsl:when
                test="not($rowProp/@fo:background-color) and $tableProp/@fo:background-color != 'transparent'  and $tableProp/@fo:background-color">
                <w:shd w:val="clear" w:color="auto"
                  w:fill="{substring($tableProp/@fo:background-color, 2, string-length($tableProp/@fo:background-color) -1)}"
                />
              </xsl:when>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>

        <w:tcMar>
          <xsl:choose>
            <xsl:when test="$cellProp[@fo:padding and @fo:padding != 'none']">
              <xsl:variable name="padding">
                <xsl:call-template name="twips-measure">
                  <xsl:with-param name="length" select="$cellProp/@fo:padding"/>
                </xsl:call-template>
              </xsl:variable>
              <w:top w:w="{$padding}" w:type="{$type}"/>
              <w:left w:w="{$padding}" w:type="{$type}"/>
              <w:bottom w:w="{$padding}" w:type="{$type}"/>
              <w:right w:w="{$padding}" w:type="{$type}"/>
            </xsl:when>
            <xsl:otherwise>
              <w:top w:type="{$type}">
                <xsl:attribute name="w:w">
                  <xsl:call-template name="twips-measure">
                    <xsl:with-param name="length" select="$cellProp/@fo:padding-top"/>
                  </xsl:call-template>
                </xsl:attribute>
              </w:top>
              <w:left w:type="{$type}">
                <xsl:attribute name="w:w">
                  <xsl:call-template name="twips-measure">
                    <xsl:with-param name="length" select="$cellProp/@fo:padding-left"/>
                  </xsl:call-template>
                </xsl:attribute>
              </w:left>
              <w:bottom w:type="{$type}">
                <xsl:attribute name="w:w">
                  <xsl:call-template name="twips-measure">
                    <xsl:with-param name="length" select="$cellProp/@fo:padding-bottom"/>
                  </xsl:call-template>
                </xsl:attribute>
              </w:bottom>
              <w:right w:type="{$type}">
                <xsl:attribute name="w:w">
                  <xsl:call-template name="twips-measure">
                    <xsl:with-param name="length" select="$cellProp/@fo:padding-right"/>
                  </xsl:call-template>
                </xsl:attribute>
              </w:right>
            </xsl:otherwise>
          </xsl:choose>
        </w:tcMar>
        <xsl:if test="$cellProp/@style:writing-mode">
          <xsl:choose>
            <xsl:when test="$cellProp[@style:writing-mode = 'tb-rl']">
              <w:textDirection w:val="tbRl"/>
            </xsl:when>
            <xsl:when test="$cellProp[@style:writing-mode = 'lr-tb']">
              <w:textDirection w:val="lrTb"/>
            </xsl:when>
          </xsl:choose>
        </xsl:if>
        <xsl:if test="$cellProp[@style:vertical-align and @style:vertical-align!='']">
          <xsl:choose>
            <xsl:when test="$cellProp/@style:vertical-align = 'middle'">
              <w:vAlign w:val="center"/>
            </xsl:when>
            <xsl:otherwise>
              <w:vAlign w:val="{$cellProp/@style:vertical-align}"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>

      </w:tcPr>

      <xsl:choose>
        <xsl:when test="not(child::table:table[@table:is-sub-table='true']) and $merge &lt; 2">
          <xsl:apply-templates/>
          <!-- must precede a w:tc, otherwise it crashes. Xml schema validation does not check this. -->
        </xsl:when>
        <xsl:otherwise>
          <w:p/>
        </xsl:otherwise>
      </xsl:choose>
    </w:tc>
  </xsl:template>

  <!-- Computes span value -->
  <xsl:template name="ComputeSpanValue">1</xsl:template>


</xsl:stylesheet>

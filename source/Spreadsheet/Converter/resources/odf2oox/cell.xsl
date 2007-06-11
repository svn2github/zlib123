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
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" exclude-result-prefixes="table r">

  <!-- insert column properties into sheet -->
  <xsl:template match="table:table-column" mode="sheet">
    <xsl:param name="colNumber"/>
    <xsl:param name="defaultFontSize"/>

    <!-- tableId required for CheckIfColumnStyle template -->
    <xsl:variable name="tableId">
      <xsl:value-of select="generate-id(ancestor::table:table)"/>
    </xsl:variable>

    <xsl:variable name="max">
      <xsl:choose>
        <xsl:when test="@table:number-columns-repeated">
          <xsl:value-of select="$colNumber+@table:number-columns-repeated - 1"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$colNumber"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:if test="not($max = '256' and @table:visibility = 'collapse')">
      <col min="{$colNumber}">
        <xsl:attribute name="max">
          <xsl:value-of select="$max"/>
        </xsl:attribute>

        <!-- insert column width -->
        <xsl:if
          test="key('style',@table:style-name)/style:table-column-properties/@style:column-width">
          <xsl:attribute name="width">
            <xsl:call-template name="ConvertToCharacters">
              <xsl:with-param name="width">
                <xsl:value-of
                  select="key('style',@table:style-name)/style:table-column-properties/@style:column-width"
                />
              </xsl:with-param>
              <xsl:with-param name="defaultFontSize" select="$defaultFontSize"/>
            </xsl:call-template>
          </xsl:attribute>
          <xsl:attribute name="customWidth">1</xsl:attribute>
        </xsl:if>

        <xsl:if test="@table:visibility = 'collapse'">
          <xsl:attribute name="hidden">1</xsl:attribute>
        </xsl:if>

        <xsl:if test="@table:default-cell-style-name != 'Default' ">
          <!-- column style is when in all posible rows there is a cell in this column -->
          <xsl:variable name="checkColumnStyle">
            <xsl:for-each select="ancestor::table:table/descendant::table:table-row[1]">
              <xsl:call-template name="CheckIfColumnStyle">
                <xsl:with-param name="number" select="$colNumber"/>
                <xsl:with-param name="table" select="$tableId"/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:variable>

          <xsl:if test="$checkColumnStyle = 'true' ">
            <xsl:for-each select="key('style',@table:default-cell-style-name)">
              <xsl:attribute name="style">
                <xsl:number count="style:style[@style:family='table-cell']" level="any"/>
              </xsl:attribute>
            </xsl:for-each>
          </xsl:if>
        </xsl:if>
      </col>
    </xsl:if>

    <!-- insert next column -->
    <xsl:choose>
      <!-- next column is sibling of this one -->
      <xsl:when test="following-sibling::node()[1][name() = 'table:table-column' ]">
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
          <xsl:with-param name="defaultFontSize" select="$defaultFontSize"/>
        </xsl:apply-templates>
      </xsl:when>
      <!-- this is the last column before header  -->
      <xsl:when test="following-sibling::node()[1][name() = 'table:table-header-columns' ]">
        <xsl:apply-templates select="following-sibling::node()[1]/table:table-column[1]"
          mode="sheet">
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
          <xsl:with-param name="defaultFontSize" select="$defaultFontSize"/>
        </xsl:apply-templates>
      </xsl:when>
      <!-- this is the last column inside header -->
      <xsl:when
        test="not(following-sibling::node()[1][name() = 'table:table-column' ]) and parent::node()[name() = 'table:table-header-columns' ] and parent::node()/following-sibling::table:table-column[1]">
        <xsl:apply-templates select="parent::node()/following-sibling::table:table-column[1]"
          mode="sheet">
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
          <xsl:with-param name="defaultFontSize" select="$defaultFontSize"/>
        </xsl:apply-templates>
      </xsl:when>
    </xsl:choose>

  </xsl:template>

  <!-- Check if 256 column are hidden -->

  <xsl:template match="table:table-column" mode="defaultColWidth">
    <xsl:param name="colNumber"/>
    <xsl:param name="result"/>

    <xsl:variable name="max">
      <xsl:choose>
        <xsl:when test="@table:number-columns-repeated">
          <xsl:value-of select="$colNumber+@table:number-columns-repeated - 1"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$colNumber"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="not($max = '256' and @table:visibility = 'collapse')">
        <!-- insert next column -->
        <xsl:choose>
          <!-- next column is a sibling of this one -->
          <xsl:when test="following-sibling::node()[1][name() = 'table:table-column' ]">
            <xsl:apply-templates select="following-sibling::table:table-column[1]"
              mode="defaultColWidth">
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
              <xsl:with-param name="result">
                <xsl:text>false</xsl:text>
              </xsl:with-param>
            </xsl:apply-templates>
          </xsl:when>
          <!-- this is the last column before header -->
          <xsl:when test="following-sibling::node()[1][name() = 'table:table-header-columns' ]">
            <xsl:apply-templates select="following-sibling::node()[1]/table:table-column[1]"
              mode="defaultColWidth">
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
              <xsl:with-param name="result">
                <xsl:text>false</xsl:text>
              </xsl:with-param>
            </xsl:apply-templates>
          </xsl:when>
          <!-- this is the last column inside header -->
          <xsl:when
            test="not(following-sibling::node()[1][name() = 'table:table-column' ]) and parent::node()[name() = 'table:table-header-columns' ]">
            <xsl:apply-templates select="parent::node()/following-sibling::table:table-column[1]"
              mode="defaultColWidth">
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
              <xsl:with-param name="result">
                <xsl:text>false</xsl:text>
              </xsl:with-param>
            </xsl:apply-templates>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$result"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:text>true</xsl:text>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>


  <!-- Create Variable with "Default Style Name" -->
  <xsl:template name="CreateDefaultTag">
    <xsl:param name="colNumber"/>
    <xsl:param name="TagNumber"/>

    <xsl:param name="Tag"/>
    <xsl:param name="repeat"/>

    <xsl:choose>
      <xsl:when test="@table:number-columns-repeated &gt; $repeat">
        <xsl:call-template name="CreateDefaultTag">
          <xsl:with-param name="colNumber">
            <xsl:value-of select="$colNumber + 1"/>
          </xsl:with-param>
          <xsl:with-param name="TagNumber">
            <xsl:value-of select="$TagNumber"/>
          </xsl:with-param>
          <xsl:with-param name="Tag">
            <xsl:value-of
              select="concat(concat(concat($colNumber, @table:default-cell-style-name),';'), $Tag)"
            />
          </xsl:with-param>
          <xsl:with-param name="repeat">
            <xsl:value-of select="$repeat + 1"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of
          select="concat(concat(concat($colNumber, @table:default-cell-style-name),';'), $Tag)"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="CreateStyleTags">
    <xsl:param name="repeat"/>
    <xsl:param name="colNumber"/>
    <xsl:param name="TagNumber"/>

    <xsl:param name="Tag"/>
    <xsl:param name="count" select="1"/>

    <xsl:choose>
      <xsl:when test="$repeat &gt; $count">
        <xsl:call-template name="CreateStyleTags">
          <xsl:with-param name="repeat" select="$repeat"/>
          <xsl:with-param name="colNumber" select="$colNumber + 1"/>
          <xsl:with-param name="TagNumber" select="$TagNumber"/>
          <xsl:with-param name="Tag">
            <xsl:value-of
              select="concat(concat(concat ('K', (concat($colNumber, ':')), @table:default-cell-style-name),';'), $Tag)"
            />
          </xsl:with-param>
          <xsl:with-param name="count" select="$count + 1"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of
          select="concat(concat(concat ('K', (concat($colNumber, ':')), @table:default-cell-style-name),';'), $Tag)"
        />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="table:table-column" mode="tag">
    <xsl:param name="colNumber"/>
    <xsl:param name="Tag"/>

    <xsl:variable name="TagNumber">
      <xsl:value-of select="concat('Tag', position())"/>
    </xsl:variable>

    <xsl:variable name="NextColl">
      <xsl:choose>
        <xsl:when test="@table:number-columns-repeated">
          <xsl:value-of select="$colNumber+@table:number-columns-repeated"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$colNumber+1"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <!-- style char for this table:table-column tag -->
    <xsl:variable name="TagChar">
      <xsl:choose>
        <!-- when this tag describes group of columns with changed default-cell-style-name -->
        <xsl:when
          test="@table:number-columns-repeated and @table:default-cell-style-name != 'Default' ">
          <xsl:call-template name="CreateStyleTags">
            <xsl:with-param name="colNumber">
              <xsl:value-of select="$colNumber"/>
            </xsl:with-param>
            <xsl:with-param name="TagNumber">
              <xsl:value-of select="$TagNumber"/>
            </xsl:with-param>
            <xsl:with-param name="Tag">
              <xsl:value-of select="$Tag"/>
            </xsl:with-param>
            <xsl:with-param name="repeat">
              <xsl:value-of select="@table:number-columns-repeated"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:when>
        <!-- when this tag describes column(s) with default default-cell-style-name -->
        <xsl:when test="@table:default-cell-style-name = 'Default' ">
          <xsl:call-template name="CreateDefaultTag">
            <xsl:with-param name="colNumber">
              <xsl:value-of select="$colNumber"/>
            </xsl:with-param>
            <xsl:with-param name="TagNumber">
              <xsl:value-of select="$TagNumber"/>
            </xsl:with-param>
            <xsl:with-param name="Tag">
              <xsl:value-of select="$Tag"/>
            </xsl:with-param>
            <xsl:with-param name="repeat">1</xsl:with-param>
          </xsl:call-template>
        </xsl:when>
        <!-- when this tag describes single column with changed default-cell-style-name -->
        <xsl:otherwise>
          <xsl:value-of
            select="concat(concat(concat ('K', (concat($colNumber, ':')), @table:default-cell-style-name),';'), $Tag)"
          />
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <!--check next column -->
    <xsl:choose>
      <!-- next column is sibling of this one -->
      <xsl:when test="following-sibling::node()[1][name() = 'table:table-column' ]">
        <xsl:apply-templates select="following-sibling::table:table-column[1]" mode="tag">
          <xsl:with-param name="colNumber">
            <xsl:value-of select="$NextColl"/>
          </xsl:with-param>
          <xsl:with-param name="Tag">
            <xsl:value-of select="$TagChar"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>
      <!-- this is the last column before header  -->
      <xsl:when test="following-sibling::node()[1][name() = 'table:table-header-columns' ]">
        <xsl:apply-templates select="following-sibling::node()[1]/table:table-column[1]" mode="tag">
          <xsl:with-param name="colNumber">
            <xsl:value-of select="$NextColl"/>
          </xsl:with-param>
          <xsl:with-param name="Tag">
            <xsl:value-of select="$TagChar"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>
      <!-- this is the last column inside header  -->
      <xsl:when
        test="not(following-sibling::node()[1][name() = 'table:table-column' ]) and parent::node()[name() = 'table:table-header-columns' ] and parent::node()/following-sibling::table:table-column[1]">
        <xsl:apply-templates select="parent::node()/following-sibling::table:table-column[1]"
          mode="tag">
          <xsl:with-param name="colNumber">
            <xsl:value-of select="$NextColl"/>
          </xsl:with-param>
          <xsl:with-param name="Tag">
            <xsl:value-of select="$TagChar"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$TagChar"/>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <!-- Check if 65536 rows are hidden -->

  <xsl:template match="table:table-row" mode="zeroHeight">
    <xsl:param name="rowNumber"/>
    <xsl:param name="collapse"/>

    <xsl:choose>
      <xsl:when
        test="not(following-sibling::table:table-row or following-sibling::table:table-header-rows) and @table:visibility='collapse'">
        <xsl:variable name="CheckNumber">
          <xsl:choose>
            <xsl:when test="@table:number-rows-repeated">
              <xsl:value-of select="$rowNumber+@table:number-rows-repeated"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$rowNumber+1"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="$CheckNumber = '65536'">
            <xsl:text>true</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>false</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when
        test="following-sibling::table:table-row[@table:visibility='collapse'] or parent::node()/following-sibling::table:table-row[@table:visibility='collapse']">
        <xsl:choose>
          <xsl:when test="name(following-sibling::node()[1]) = 'table:table-row'">
            <xsl:apply-templates select="following-sibling::table:table-row[1]" mode="zeroHeight">
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
            </xsl:apply-templates>
          </xsl:when>
          <xsl:when test="name(following-sibling::node()[1]) = 'table:table-header-rows'">
            <xsl:apply-templates
              select="following-sibling::table:table-header-rows/table:table-row[1]"
              mode="zeroHeight">
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
            </xsl:apply-templates>
          </xsl:when>
          <xsl:when
            test="name(parent::node()) = 'table:table-header-rows' and parent::node()/following-sibling::table:table-row">
            <xsl:apply-templates select="parent::node()/following-sibling::table:table-row[1]"
              mode="zeroHeight">
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
            </xsl:apply-templates>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:text>false</xsl:text>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <!-- insert row into sheet -->
  <xsl:template match="table:table-row" mode="sheet">
    <xsl:param name="rowNumber"/>
    <xsl:param name="cellNumber"/>
    <xsl:param name="defaultRowHeight"/>
    <xsl:param name="TableColumnTagNum"/>
    <xsl:param name="MergeCell"/>
    <xsl:param name="MergeCellStyle"/>
    <xsl:param name="CheckRowHidden"/>
    <xsl:param name="CheckIfDefaultBorder"/>

    <xsl:variable name="height">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length">
          <xsl:value-of
            select="key('style',@table:style-name)/style:table-row-properties/@style:row-height"/>
        </xsl:with-param>
        <xsl:with-param name="round">false</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if
      test="table:table-cell or @table:visibility='collapse' or  @table:visibility='filter' or ($height != $defaultRowHeight and following-sibling::table:table-row/table:table-cell/text:p|text:span) or table:covered-table-cell">

      <row r="{$rowNumber}">

        <!-- insert row height -->
        <xsl:if test="$height">
          <xsl:attribute name="ht">
            <xsl:value-of select="$height"/>
          </xsl:attribute>
          <xsl:if
            test="not(key('style',@table:style-name)/style:table-row-properties/@style:use-optimal-row-height = 'true' ) or table:covered-table-cell">
            <xsl:attribute name="customHeight">1</xsl:attribute>
          </xsl:if>
        </xsl:if>

        <xsl:if test="@table:visibility = 'collapse' or @table:visibility = 'filter'">
          <xsl:attribute name="hidden">1</xsl:attribute>
        </xsl:if>

        <!-- insert first cell -->
        <xsl:apply-templates select="node()[1]" mode="sheet">
          <xsl:with-param name="colNumber">0</xsl:with-param>
          <xsl:with-param name="rowNumber" select="$rowNumber"/>
          <xsl:with-param name="cellNumber" select="$cellNumber"/>
          <xsl:with-param name="TableColumnTagNum">
            <xsl:value-of select="$TableColumnTagNum"/>
          </xsl:with-param>
          <xsl:with-param name="MergeCell">
            <xsl:value-of select="$MergeCell"/>
          </xsl:with-param>
          <xsl:with-param name="MergeCellStyle">
            <xsl:value-of select="$MergeCellStyle"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </row>

      <!-- insert repeated rows -->
      <xsl:if test="@table:number-rows-repeated">
        <xsl:if
          test="not(@table:visibility='collapse') or not(following-sibling::table:table-row[@table:visibility = 'collapse']) or following-sibling::table:table-row/table:table-cell/text:p or $CheckRowHidden != 'true' or contains($CheckIfDefaultBorder, 'true')">
          <xsl:call-template name="InsertRepeatedRows">
            <xsl:with-param name="numberRowsRepeated">
              <xsl:value-of select="@table:number-rows-repeated"/>
            </xsl:with-param>
            <xsl:with-param name="numberOfAllRowsRepeated">
              <xsl:value-of select="@table:number-rows-repeated"/>
            </xsl:with-param>
            <xsl:with-param name="rowNumber">
              <xsl:value-of select="$rowNumber + 1"/>
            </xsl:with-param>
            <xsl:with-param name="cellNumber">
              <xsl:value-of select="$cellNumber"/>
            </xsl:with-param>
            <xsl:with-param name="height">
              <xsl:value-of select="$height"/>
            </xsl:with-param>
            <xsl:with-param name="defaultRowHeight">
              <xsl:value-of select="$defaultRowHeight"/>
            </xsl:with-param>
            <xsl:with-param name="TableColumnTagNum">
              <xsl:value-of select="$TableColumnTagNum"/>
            </xsl:with-param>
            <xsl:with-param name="MergeCell">
              <xsl:value-of select="$MergeCell"/>
            </xsl:with-param>
            <xsl:with-param name="MergeCellStyle">
              <xsl:value-of select="$MergeCellStyle"/>
            </xsl:with-param>
            <xsl:with-param name="CheckIfDefaultBorder">
              <xsl:value-of select="$CheckIfDefaultBorder"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:if>
      </xsl:if>

    </xsl:if>

    <!-- insert next row -->
    <xsl:choose>
      <!-- next row is a sibling -->
      <xsl:when test="following-sibling::node()[1][name() = 'table:table-row' ]">
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
            <!-- last or is for cells with error -->
            <xsl:value-of
              select="$cellNumber + count(child::table:table-cell[text:p and not(@office:value-type='float') and (@office:value-type='string' or @office:value-type='boolean' or not((number(text:p) or text:p = 0 or contains(text:p,',') or contains(text:p,'%') or @office:value-type='currency' or @office:value-type='date' or @office:value-type='time')))])"
            />
          </xsl:with-param>
          <xsl:with-param name="defaultRowHeight" select="$defaultRowHeight"/>
          <xsl:with-param name="TableColumnTagNum">
            <xsl:value-of select="$TableColumnTagNum"/>
          </xsl:with-param>
          <xsl:with-param name="MergeCell">
            <xsl:value-of select="$MergeCell"/>
          </xsl:with-param>
          <xsl:with-param name="MergeCellStyle">
            <xsl:value-of select="$MergeCellStyle"/>
          </xsl:with-param>
          <xsl:with-param name="CheckRowHidden">
            <xsl:value-of select="$CheckRowHidden"/>
          </xsl:with-param>
          <xsl:with-param name="CheckIfDefaultBorder">
            <xsl:value-of select="$CheckIfDefaultBorder"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>
      <!-- next row is inside header rows -->
      <xsl:when test="following-sibling::node()[1][name() = 'table:table-header-rows' ]">
        <xsl:apply-templates select="following-sibling::table:table-header-rows/table:table-row[1]"
          mode="sheet">
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
            <!-- last or is for cells with error -->
            <xsl:value-of
              select="$cellNumber + count(child::table:table-cell[text:p and not(@office:value-type='float') and (@office:value-type='string' or @office:value-type='boolean' or not((number(text:p) or text:p = 0 or contains(text:p,',') or contains(text:p,'%') or @office:value-type='currency' or @office:value-type='date'  or @office:value-type='time')))])"
            />
          </xsl:with-param>
          <xsl:with-param name="defaultRowHeight" select="$defaultRowHeight"/>
          <xsl:with-param name="TableColumnTagNum">
            <xsl:value-of select="$TableColumnTagNum"/>
          </xsl:with-param>
          <xsl:with-param name="MergeCell">
            <xsl:value-of select="$MergeCell"/>
          </xsl:with-param>
          <xsl:with-param name="MergeCellStyle">
            <xsl:value-of select="$MergeCellStyle"/>
          </xsl:with-param>
          <xsl:with-param name="CheckRowHidden">
            <xsl:value-of select="$CheckRowHidden"/>
          </xsl:with-param>
          <xsl:with-param name="CheckIfDefaultBorder">
            <xsl:value-of select="$CheckIfDefaultBorder"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>
      <!-- this is last row inside header rows, next row is outside -->
      <xsl:when
        test="parent::node()[name()='table:table-header-rows'] and not(following-sibling::node()[1][name() = 'table:table-row' ])">
        <xsl:apply-templates select="parent::node()/following-sibling::table:table-row[1]"
          mode="sheet">
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
            <!-- last or is for cells with error -->
            <xsl:value-of
              select="$cellNumber + count(child::table:table-cell[text:p and not(@office:value-type='float') and (@office:value-type='string' or @office:value-type='boolean' or not((number(text:p) or text:p = 0 or contains(text:p,',') or contains(text:p,'%') or @office:value-type='currency' or @office:value-type='date' or @office:value-type='time')))])"
            />
          </xsl:with-param>
          <xsl:with-param name="defaultRowHeight" select="$defaultRowHeight"/>
          <xsl:with-param name="TableColumnTagNum">
            <xsl:value-of select="$TableColumnTagNum"/>
          </xsl:with-param>
          <xsl:with-param name="MergeCell">
            <xsl:value-of select="$MergeCell"/>
          </xsl:with-param>
          <xsl:with-param name="MergeCellStyle">
            <xsl:value-of select="$MergeCellStyle"/>
          </xsl:with-param>
          <xsl:with-param name="CheckRowHidden">
            <xsl:value-of select="$CheckRowHidden"/>
          </xsl:with-param>
          <xsl:with-param name="CheckIfDefaultBorder">
            <xsl:value-of select="$CheckIfDefaultBorder"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- template which inserts repeated rows -->
  <xsl:template name="InsertRepeatedRows">
    <xsl:param name="numberRowsRepeated"/>
    <xsl:param name="numberOfAllRowsRepeated"/>
    <xsl:param name="rowNumber"/>
    <xsl:param name="cellNumber"/>
    <xsl:param name="height"/>
    <xsl:param name="defaultRowHeight"/>
    <xsl:param name="TableColumnTagNum"/>
    <xsl:param name="MergeCell"/>
    <xsl:param name="MergeCellStyle"/>
    <xsl:param name="CheckIfDefaultBorder"/>

    <xsl:if
      test="table:table-cell/text:p or @table:visibility='collapse' or  @table:visibility='filter' or ($height != $defaultRowHeight and following-sibling::table:table-row/table:table-cell/text:p|text:span) or contains($CheckIfDefaultBorder, 'true')">

      <xsl:choose>
        <xsl:when test="$numberRowsRepeated &gt; 1">
          <row>
            <xsl:attribute name="r">
              <xsl:value-of select="$rowNumber"/>
            </xsl:attribute>

            <!-- insert row height -->
            <xsl:if test="$height">
              <xsl:attribute name="ht">
                <xsl:value-of select="$height"/>
              </xsl:attribute>
              <xsl:attribute name="customHeight">1</xsl:attribute>
            </xsl:if>

            <xsl:if test="@table:visibility = 'collapse' or @table:visibility = 'filter'">
              <xsl:attribute name="hidden">1</xsl:attribute>
            </xsl:if>

            <!-- insert first cell -->
            <xsl:apply-templates select="node()[1]" mode="sheet">
              <xsl:with-param name="colNumber">0</xsl:with-param>
              <xsl:with-param name="rowNumber" select="$rowNumber"/>
              <xsl:with-param name="cellNumber" select="$cellNumber"/>
              <xsl:with-param name="TableColumnTagNum">
                <xsl:value-of select="$TableColumnTagNum"/>
              </xsl:with-param>
              <xsl:with-param name="MergeCell">
                <xsl:value-of select="$MergeCell"/>
              </xsl:with-param>
              <xsl:with-param name="MergeCellStyle">
                <xsl:value-of select="$MergeCellStyle"/>
              </xsl:with-param>
            </xsl:apply-templates>
          </row>


          <!-- insert repeated rows -->
          <xsl:if test="@table:number-rows-repeated">
            <xsl:call-template name="InsertRepeatedRows">
              <xsl:with-param name="numberRowsRepeated">
                <xsl:choose>
                  <xsl:when test="$numberRowsRepeated - 1 &gt; 180">
                    <xsl:text>49</xsl:text>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="$numberRowsRepeated - 1"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:with-param>
              <xsl:with-param name="numberOfAllRowsRepeated">
                <xsl:value-of select="@table:number-rows-repeated"/>
              </xsl:with-param>
              <xsl:with-param name="rowNumber">
                <xsl:value-of select="$rowNumber + 1"/>
              </xsl:with-param>
              <xsl:with-param name="cellNumber">
                <xsl:value-of select="$cellNumber"/>
              </xsl:with-param>
              <xsl:with-param name="height">
                <xsl:value-of select="$height"/>
              </xsl:with-param>
              <xsl:with-param name="defaultRowHeight">
                <xsl:value-of select="$defaultRowHeight"/>
              </xsl:with-param>
              <xsl:with-param name="TableColumnTagNum">
                <xsl:value-of select="$TableColumnTagNum"/>
              </xsl:with-param>
              <xsl:with-param name="MergeCell">
                <xsl:value-of select="$MergeCell"/>
              </xsl:with-param>
              <xsl:with-param name="MergeCellStyle">
                <xsl:value-of select="$MergeCellStyle"/>
              </xsl:with-param>
              <xsl:with-param name="CheckIfDefaultBorder">
                <xsl:value-of select="$CheckIfDefaultBorder"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:if>
        </xsl:when>
      </xsl:choose>
    </xsl:if>

  </xsl:template>

  <!-- insert cell into row -->
  <xsl:template match="table:table-cell|table:covered-table-cell" mode="sheet">
    <xsl:param name="colNumber"/>
    <xsl:param name="rowNumber"/>
    <xsl:param name="cellNumber"/>
    <xsl:param name="TableColumnTagNum"/>
    <xsl:param name="MergeCell"/>
    <xsl:param name="MergeCellStyle"/>
    <xsl:message terminate="no">progress:table:table-cell</xsl:message>
    <xsl:variable name="CountStyleTableCell">
      <xsl:value-of
        select="count(document('content.xml')/office:document-content/office:automatic-styles/style:style[@style:family='table-cell'])"
      />
    </xsl:variable>

    <xsl:call-template name="InsertConvertCell">
      <xsl:with-param name="colNumber">
        <xsl:value-of select="$colNumber"/>
      </xsl:with-param>
      <xsl:with-param name="rowNumber">
        <xsl:value-of select="$rowNumber"/>
      </xsl:with-param>
      <xsl:with-param name="cellNumber">
        <xsl:value-of select="$cellNumber"/>
      </xsl:with-param>
      <xsl:with-param name="ColumnRepeated">1</xsl:with-param>
      <xsl:with-param name="TableColumnTagNum">
        <xsl:value-of select="$TableColumnTagNum"/>
      </xsl:with-param>
      <xsl:with-param name="MergeCell">
        <xsl:value-of select="$MergeCell"/>
      </xsl:with-param>
      <xsl:with-param name="MergeCellStyle">
        <xsl:value-of select="$MergeCellStyle"/>
      </xsl:with-param>
      <xsl:with-param name="CountStyleTableCell">
        <xsl:value-of select="$CountStyleTableCell"/>
      </xsl:with-param>
    </xsl:call-template>

    <xsl:apply-templates mode="cell">
      <xsl:with-param name="colNumber">
        <xsl:value-of select="$colNumber"/>
      </xsl:with-param>
      <xsl:with-param name="rowNumber">
        <xsl:value-of select="$rowNumber"/>
      </xsl:with-param>
    </xsl:apply-templates>

  </xsl:template>

  <xsl:template match="office:annotation" mode="cell">
    <xsl:param name="colNumber"/>
    <xsl:param name="rowNumber"/>
    <xsl:variable name="colChar">
      <xsl:call-template name="NumbersToChars">
        <xsl:with-param name="num" select="$colNumber"/>
      </xsl:call-template>
    </xsl:variable>
    <pxsi:commentmark xmlns:pxsi="urn:cleverage:xmlns:post-processings:comments"
      ref="{concat($colChar,$rowNumber)}" noteId="{count(preceding::office:annotation)+1}"/>

    <pxsi:commentDrawingMark xmlns:x="urn:schemas-microsoft-com:office:excel"
      xmlns:pxsi="urn:cleverage:xmlns:post-processings:drawings"
      noteId="{count(preceding::office:annotation)+1}">
      <x:Row>
        <xsl:value-of select="$rowNumber - 1"/>
      </x:Row>
      <x:Column>
        <xsl:value-of select="$colNumber"/>
      </x:Column>
    </pxsi:commentDrawingMark>

  </xsl:template>

  <!-- insert cell -->
  <xsl:template name="InsertNextCell">
    <xsl:param name="colNumber"/>
    <xsl:param name="rowNumber"/>
    <xsl:param name="cellNumber"/>
    <xsl:param name="TableColumnTagNum"/>
    <xsl:param name="MergeCell"/>
    <xsl:param name="MergeCellStyle"/>

    <xsl:choose>
      <!-- Checks if  next index is table:table-cell-->
      <xsl:when test="following-sibling::node()[1][name()='table:table-cell']">
        <xsl:apply-templates select="following-sibling::table:table-cell[1]" mode="sheet">
          <xsl:with-param name="colNumber">
            <xsl:value-of select="$colNumber+1"/>
          </xsl:with-param>
          <xsl:with-param name="rowNumber" select="$rowNumber"/>
          <!-- if this node was table:table-cell with string than increase cellNumber-->
          <xsl:with-param name="cellNumber">
            <xsl:choose>
              <xsl:when
                test="name()='table:table-cell' and child::text:p and not(@office:value-type='float') and (@office:value-type='string' or @office:value-type='boolean' or not((number(text:p) or text:p = 0 or contains(text:p,',') or contains(text:p,'%') or @office:value-type='currency' or @office:value-type='date' or @office:value-type='time')))">
                <xsl:value-of select="$cellNumber + 1"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$cellNumber"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="TableColumnTagNum">
            <xsl:value-of select="$TableColumnTagNum"/>
          </xsl:with-param>
          <xsl:with-param name="MergeCell">
            <xsl:value-of select="$MergeCell"/>
          </xsl:with-param>
          <xsl:with-param name="MergeCellStyle">
            <xsl:value-of select="$MergeCellStyle"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>
      <!--  Checks if next index is table:covered-table-cell-->
      <xsl:when test="following-sibling::node()[1][name()='table:covered-table-cell']">
        <xsl:apply-templates select="following-sibling::table:covered-table-cell[1]" mode="sheet">
          <xsl:with-param name="colNumber">
            <xsl:value-of select="$colNumber+1"/>
          </xsl:with-param>
          <xsl:with-param name="rowNumber" select="$rowNumber"/>
          <!-- if this node was table:table-cell with string than increase cellNumber-->
          <xsl:with-param name="cellNumber">
            <xsl:choose>
              <xsl:when
                test="name()='table:table-cell' and child::text:p and (@office:value-type='string' or @office:value-type='boolean' or not((number(text:p) or text:p = 0 or contains(text:p,',') or contains(text:p,'%') or @office:value-type='currency' or @office:value-type='date' or @office:value-type='time')))">
                <xsl:value-of select="$cellNumber + 1"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$cellNumber"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="ColumnRepeated">1</xsl:with-param>
          <xsl:with-param name="TableColumnTagNum">
            <xsl:value-of select="$TableColumnTagNum"/>
          </xsl:with-param>
          <xsl:with-param name="MergeCell">
            <xsl:value-of select="$MergeCell"/>
          </xsl:with-param>
          <xsl:with-param name="MergeCellStyle">
            <xsl:value-of select="$MergeCellStyle"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>
    </xsl:choose>

  </xsl:template>


  <!-- Insert Cell for "@table:number-columns-repeated"-->
  <xsl:template name="InsertConvertCell">
    <xsl:param name="colNumber"/>
    <xsl:param name="rowNumber"/>
    <xsl:param name="cellNumber"/>
    <xsl:param name="ColumnRepeated"/>
    <xsl:param name="TableColumnTagNum"/>
    <xsl:param name="MergeCell"/>
    <xsl:param name="MergeCellStyle"/>
    <xsl:param name="CountStyleTableCell"/>

    <!-- do not show covered cells content -->

    <xsl:call-template name="InsertCell">
      <xsl:with-param name="colNumber">
        <xsl:value-of select="$colNumber"/>
      </xsl:with-param>
      <xsl:with-param name="rowNumber">
        <xsl:value-of select="$rowNumber"/>
      </xsl:with-param>
      <xsl:with-param name="cellNumber">
        <xsl:value-of select="$cellNumber"/>
      </xsl:with-param>
      <xsl:with-param name="TableColumnTagNum">
        <xsl:value-of select="$TableColumnTagNum"/>
      </xsl:with-param>
      <xsl:with-param name="MergeCell">
        <xsl:value-of select="$MergeCell"/>
      </xsl:with-param>
      <xsl:with-param name="MergeCellStyle">
        <xsl:value-of select="$MergeCellStyle"/>
      </xsl:with-param>
      <xsl:with-param name="CountStyleTableCell">
        <xsl:value-of select="$CountStyleTableCell"/>
      </xsl:with-param>
    </xsl:call-template>



    <!-- Insert cells if "@table:number-columns-repeated"  > 1 -->
    <xsl:choose>
      <!-- do not output unnecessary last cells in a row (when last cell is table:table-cell with default style, repeated columns and without content)  -->
      <!-- explenation of last 'and':
                  If column has got changed default-cell-style-name then in $TableColumnTagNum string there is entry 'K' col_number ':' default-cell-style-name ';' if not there is no 'K'. 
                  So if inside cell range defined by table:table-cell with repeated columns attribute there is column that has changed default-cell-style-name then before ';'col_number 
                  in $TableColumnTagNum there should be 'K' ($TableColumnTagNum contains listed default-cell-style-name from backward) -->
      <xsl:when
        test="@table:number-columns-repeated and not(following-sibling::node()[1]) and name() = 'table:table-cell' and not(text:p) and not(@table:table-style) and 
        not(contains(substring-before($TableColumnTagNum,';$colNumber:'),'K') or contains($TableColumnTagNum,concat('K',$colNumber)))"> </xsl:when>

      <xsl:when
        test="@table:number-columns-repeated and number(@table:number-columns-repeated) &gt; $ColumnRepeated">

        <xsl:call-template name="InsertConvertCell">
          <xsl:with-param name="colNumber">
            <xsl:value-of select="$colNumber+1"/>
          </xsl:with-param>
          <xsl:with-param name="rowNumber" select="$rowNumber"/>
          <xsl:with-param name="cellNumber">
            <xsl:value-of select="$cellNumber"/>
          </xsl:with-param>
          <xsl:with-param name="ColumnRepeated">
            <xsl:value-of select="$ColumnRepeated + 1"/>
          </xsl:with-param>
          <xsl:with-param name="TableColumnTagNum">
            <xsl:value-of select="$TableColumnTagNum"/>
          </xsl:with-param>
          <xsl:with-param name="MergeCell">
            <xsl:value-of select="$MergeCell"/>
          </xsl:with-param>
          <xsl:with-param name="MergeCellStyle">
            <xsl:value-of select="$MergeCellStyle"/>
          </xsl:with-param>
          <xsl:with-param name="CountStyleTableCell">
            <xsl:value-of select="$CountStyleTableCell"/>
          </xsl:with-param>
        </xsl:call-template>

      </xsl:when>
      <xsl:otherwise>

        <!-- insert next cell -->
        <xsl:call-template name="InsertNextCell">
          <xsl:with-param name="colNumber">
            <xsl:value-of select="$colNumber"/>
          </xsl:with-param>
          <xsl:with-param name="rowNumber">
            <xsl:value-of select="$rowNumber"/>
          </xsl:with-param>
          <xsl:with-param name="cellNumber">
            <xsl:value-of select="$cellNumber"/>
          </xsl:with-param>
          <xsl:with-param name="TableColumnTagNum">
            <xsl:value-of select="$TableColumnTagNum"/>
          </xsl:with-param>
          <xsl:with-param name="MergeCell">
            <xsl:value-of select="$MergeCell"/>
          </xsl:with-param>
          <xsl:with-param name="MergeCellStyle">
            <xsl:value-of select="$MergeCellStyle"/>
          </xsl:with-param>
        </xsl:call-template>

      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <xsl:template name="InsertCell">
    <xsl:param name="colNumber"/>
    <xsl:param name="rowNumber"/>
    <xsl:param name="cellNumber"/>
    <xsl:param name="TableColumnTagNum"/>
    <xsl:param name="MergeCell"/>
    <xsl:param name="MergeCellStyle"/>
    <xsl:param name="CountStyleTableCell"/>

    <xsl:variable name="columnCellStyle">
      <xsl:call-template name="GetColumnCellStyle">
        <xsl:with-param name="colNum">
          <xsl:value-of select="$colNumber + 1"/>
        </xsl:with-param>
        <xsl:with-param name="TableColumnTagNum">
          <xsl:value-of select="$TableColumnTagNum"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="CheckIfMerge">
      <xsl:call-template name="CheckIfMergeColl">
        <xsl:with-param name="MergeCell">
          <xsl:value-of select="$MergeCell"/>
        </xsl:with-param>
        <xsl:with-param name="MergeCellStyle">
          <xsl:value-of select="$MergeCellStyle"/>
        </xsl:with-param>
        <xsl:with-param name="colNumber">
          <xsl:value-of select="$colNumber"/>
        </xsl:with-param>
        <xsl:with-param name="rowNumber">
          <xsl:value-of select="$rowNumber"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="MergeStart">
      <xsl:choose>
        <xsl:when test="name() = 'table:covered-table-cell'">
          <xsl:call-template name="CheckIfInMerge">
            <xsl:with-param name="MergeCell">
              <xsl:value-of select="$MergeCell"/>
            </xsl:with-param>
            <xsl:with-param name="colNumber">
              <xsl:value-of select="$colNumber + 1"/>
            </xsl:with-param>
            <xsl:with-param name="rowNumber">
              <xsl:value-of select="$rowNumber"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:text>false</xsl:text>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="cellFormats">
      <xsl:value-of
        select="count(document('content.xml')/office:document-content/office:automatic-styles/style:style[@style:family='table-cell']) + 1"
      />
    </xsl:variable>

    <xsl:variable name="cellStyles">
      <xsl:value-of
        select="count(document('styles.xml')/office:document-styles/office:styles/style:style[@style:family='table-cell'])"
      />
    </xsl:variable>

    <xsl:if
      test="child::text:p or $columnCellStyle != '' or name() = 'table:covered-table-cell' or $CheckIfMerge = 'start' or @table:style-name != ''">

      <c>
        <xsl:attribute name="r">
          <xsl:variable name="colChar">
            <xsl:call-template name="NumbersToChars">
              <xsl:with-param name="num" select="$colNumber"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:value-of select="concat($colChar,$rowNumber)"/>
        </xsl:attribute>

        <!-- insert cell style number -->
        <xsl:choose>
          <!-- if it is a multiline cell -->
          <xsl:when test="text:p[2]">

            <xsl:variable name="multilineNumber">
              <xsl:for-each select="text:p[2]">
                <xsl:number count="table:table-cell[text:p[2]]" level="any"/>
              </xsl:for-each>
            </xsl:variable>

            <xsl:attribute name="s">
              <xsl:value-of select="$cellFormats + $cellStyles + $multilineNumber - 1"/>
            </xsl:attribute>
          </xsl:when>

          <!-- if it is a hyperlink  in the cell-->
          <xsl:when
            test="descendant::text:a[not(ancestor::table:table-row-group or ancestor::table:covered-table-cell)]">

            <xsl:variable name="multilines">
              <xsl:for-each
                select="document('content.xml')/office:document-content/office:body/office:spreadsheet">
                <xsl:value-of select="count(descendant::table:table-cell/text:p[2])"/>
              </xsl:for-each>
            </xsl:variable>

            <xsl:attribute name="s">
              <xsl:choose>
                <!-- if there is no 'Hyperlink' style -->
                <xsl:when
                  test="not(document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = 'Hyperlink' ])">
                  <xsl:value-of select="$cellFormats + $cellStyles + $multilines"/>
                </xsl:when>
                
                <xsl:otherwise>
                  <xsl:variable name="hyperlinkStyle">
                    <xsl:for-each select="document('styles.xml')">
                      <xsl:for-each select="key('style','Hyperlink')">
                        <xsl:number count="style:style[@style:family='table-cell']" level="any"/>
                      </xsl:for-each>
                    </xsl:for-each>
                  </xsl:variable>
                  <xsl:value-of select="$CountStyleTableCell + $hyperlinkStyle"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </xsl:when>

          <xsl:otherwise>
            <xsl:choose>
              <xsl:when
                test="name() = 'table:covered-table-cell' and substring-before(substring-after($MergeCellStyle, concat($MergeStart, ':')), ';') != ''">
                <xsl:variable name="style">
                  <xsl:value-of
                    select="substring-before(substring-after($MergeCellStyle, concat($MergeStart, ':')), ';')"
                  />
                </xsl:variable>
                <xsl:choose>
                  <xsl:when test="key('style', $style)">
                    <xsl:for-each select="key('style', $style)">
                      <xsl:attribute name="s">
                        <xsl:number count="style:style[@style:family='table-cell']" level="any"/>
                      </xsl:attribute>
                    </xsl:for-each>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:variable name="TableStyleName">
                      <xsl:value-of select="@table:style-name"/>
                    </xsl:variable>
                    <xsl:for-each select="document('styles.xml')">
                      <xsl:for-each select="key('style',$TableStyleName)">
                        <xsl:variable name="CountTableCell">
                          <xsl:number count="style:style[@style:family='table-cell']" level="any"/>
                        </xsl:variable>
                        <xsl:attribute name="s">
                          <xsl:value-of select="$CountStyleTableCell+$CountTableCell"/>
                        </xsl:attribute>
                      </xsl:for-each>
                    </xsl:for-each>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>
              <!-- when style is specified in cell -->
              <xsl:when test="@table:style-name and not(name() ='table:covered-table-cell')">
                <xsl:choose>
                  <xsl:when test="key('style',@table:style-name)">
                    <xsl:for-each select="key('style',@table:style-name)">
                      <xsl:attribute name="s">
                        <xsl:number count="style:style[@style:family='table-cell']" level="any"/>
                      </xsl:attribute>
                    </xsl:for-each>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:variable name="TableStyleName">
                      <xsl:value-of select="@table:style-name"/>
                    </xsl:variable>
                    <xsl:for-each select="document('styles.xml')">
                      <xsl:for-each select="key('style',$TableStyleName)">
                        <xsl:variable name="CountTableCell">
                          <xsl:number count="style:style[@style:family='table-cell']" level="any"/>
                        </xsl:variable>
                        <xsl:attribute name="s">
                          <xsl:value-of select="$CountStyleTableCell+$CountTableCell"/>
                        </xsl:attribute>
                      </xsl:for-each>
                    </xsl:for-each>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>
              <!-- when style is specified in column -->
              <xsl:when test="$columnCellStyle != '' ">
                <xsl:choose>
                  <xsl:when test="key('style',$columnCellStyle)">
                    <xsl:for-each select="key('style',$columnCellStyle)">
                      <xsl:attribute name="s">
                        <xsl:number count="style:style[@style:family='table-cell']" level="any"/>
                      </xsl:attribute>
                    </xsl:for-each>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:for-each select="document('styles.xml')">
                      <xsl:for-each select="key('style',$columnCellStyle)">
                        <xsl:variable name="CountTableCell">
                          <xsl:number count="style:style[@style:family='table-cell']" level="any"/>
                        </xsl:variable>
                        <xsl:attribute name="s">
                          <xsl:value-of select="$CountStyleTableCell+$CountTableCell"/>
                        </xsl:attribute>
                      </xsl:for-each>
                    </xsl:for-each>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>

            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
        <!-- convert cell type -->
        <xsl:if test="child::text:p and not(name() = 'table:covered-table-cell')">
          <xsl:choose>
            <xsl:when
              test="@office:value-type='float' or (@office:value-type!='string' and @office:value-type != 'percentage' and @office:value-type != 'currency' and @office:value-type != 'date' and @office:value-type != 'time' and @office:value-type!='boolean' and (number(text:p) or text:p = 0 or contains(text:p,',') or contains(text:p,'%')))">
              <xsl:variable name="Type">
                <xsl:call-template name="ConvertTypes">
                  <xsl:with-param name="type">
                    <xsl:value-of select="@office:value-type"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:variable>

              <xsl:attribute name="t">
                <xsl:value-of select="$Type"/>
              </xsl:attribute>

              <v>
                <xsl:choose>
                  <xsl:when test="$Type = 'n'">
                    <xsl:choose>
                      <xsl:when test="@office:value != ''">
                        <xsl:value-of select="@office:value"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="translate(text:p,',','.')"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="text:p"/>
                  </xsl:otherwise>
                </xsl:choose>
              </v>
            </xsl:when>

            <!-- percentage -->
            <xsl:when test="@office:value-type = 'percentage'">
              <v>
                <xsl:choose>
                  <xsl:when test="@office:value">
                    <xsl:value-of select="@office:value"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="translate(substring-before(text:p, '%'), ',', '.')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </v>
            </xsl:when>

            <!-- currency -->
            <xsl:when test="@office:value-type = 'currency'">
              <v>
                <xsl:value-of select="@office:value"/>
              </v>
            </xsl:when>

            <!-- date -->
            <xsl:when test="@office:value-type='date'">
              <v>
                <xsl:call-template name="DateToNumber">
                  <xsl:with-param name="value">
                    <xsl:value-of select="@office:date-value"/>
                  </xsl:with-param>
                </xsl:call-template>
              </v>
            </xsl:when>

            <!-- time-->
            <xsl:when test="@office:value-type = 'time'">
              <v>
                <xsl:call-template name="TimeToNumber">
                  <xsl:with-param name="value">
                    <xsl:value-of select="@office:time-value"/>
                  </xsl:with-param>
                </xsl:call-template>
              </v>
            </xsl:when>

            <!-- last or when number cell has error -->
            <xsl:when
              test="not(@office:value-type='float') and @office:value-type = 'string' or @office:value-type = 'boolean' or not((number(text:p) or text:p = 0 or contains(text:p,',') or contains(text:p,'%') or @office:value-type='currency'))">
              <xsl:attribute name="t">s</xsl:attribute>
              <v>
                <xsl:value-of select="number($cellNumber)"/>
              </v>
            </xsl:when>
          </xsl:choose>
        </xsl:if>
      </c>
    </xsl:if>
  </xsl:template>

  <xsl:template name="GetColumnCellStyle">
    <xsl:param name="colNum"/>
    <xsl:param name="table-column_tagNum" select="1"/>
    <xsl:param name="column" select="1"/>
    <xsl:param name="TableColumnTagNum"/>
    <xsl:value-of
      select="substring-before(substring-after($TableColumnTagNum, concat(concat('K', $colNum), ':')), ';')"
    />
  </xsl:template>

  <xsl:template name="GetColNumber">
    <xsl:param name="position"/>
    <xsl:param name="count" select="1"/>
    <xsl:param name="value" select="1"/>

    <xsl:choose>
      <xsl:when test="$count &lt; $position">
        <xsl:variable name="columns">
          <xsl:choose>
            <xsl:when test="@table:number-columns-repeated">
              <xsl:value-of select="@table:number-columns-repeated"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>1</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:for-each select="following-sibling::node()[name() = 'table:table-cell' or name() = 'table:covered-table-cell' ][1]">
          <xsl:call-template name="GetColNumber">
            <xsl:with-param name="position">
              <xsl:value-of select="$position"/>
            </xsl:with-param>
            <xsl:with-param name="count">
              <xsl:value-of select="$count + 1"/>
            </xsl:with-param>
            <xsl:with-param name="value">
              <xsl:value-of select="$value + $columns"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$value"/>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <xsl:template name="GetRowNumber">
    <xsl:param name="rowId"/>
    <xsl:param name="tableId"/>
    <xsl:param name="value" select="1"/>

    <xsl:choose>
      <xsl:when test="generate-id(.) != $rowId">
        <xsl:variable name="rows">
          <xsl:choose>
            <xsl:when test="@table:number-rows-repeated">
              <xsl:value-of select="@table:number-rows-repeated"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>1</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:for-each
          select="following::table:table-row[generate-id(ancestor::table:table) = $tableId][1]">
          <xsl:call-template name="GetRowNumber">
            <xsl:with-param name="rowId" select="$rowId"/>
            <xsl:with-param name="tableId" select="$tableId"/>
            <xsl:with-param name="value">
              <xsl:value-of select="$value + $rows"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:for-each>

      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$value"/>
      </xsl:otherwise>
    </xsl:choose>
    <!--<xsl:value-of select="$rows"/>-->
  </xsl:template>

  <xsl:template name="GetLinkNumber">
    <xsl:param name="position"/>
    <xsl:param name="count" select="1"/>
    <xsl:param name="value" select="1"/>
    <xsl:choose>
      <xsl:when test="$count &lt; $position">
        <xsl:variable name="linkNum"/>

        <xsl:for-each select="following-sibling::text:a[1]">
          <xsl:call-template name="GetLinkNumber">
            <xsl:with-param name="position">
              <xsl:value-of select="$position"/>
            </xsl:with-param>
            <xsl:with-param name="count">
              <xsl:value-of select="$count + 1"/>
            </xsl:with-param>
            <xsl:with-param name="value">
              <xsl:value-of select="$value + $linkNum"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
    </xsl:choose>

  </xsl:template>

  <xsl:template name="CheckIfColumnStyle">
    <xsl:param name="number"/>
    <xsl:param name="table"/>
    <xsl:param name="count" select="0"/>
    <xsl:param name="value"/>

    <xsl:variable name="newValue">
      <xsl:for-each select="table:table-cell[1]">
        <xsl:call-template name="CheckIfCellNonEmpty">
          <xsl:with-param name="number">
            <xsl:value-of select="$number"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:for-each>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="$newValue != 'end' ">

        <xsl:variable name="rows">
          <xsl:choose>
            <xsl:when test="@table:number-rows-repeated">
              <xsl:value-of select="@table:number-rows-repeated"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>1</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:variable name="table2">
          <xsl:for-each select="following::table:table-row[1]">
            <xsl:value-of select="generate-id(ancestor::table:table)"/>
          </xsl:for-each>
        </xsl:variable>

        <!-- go to the next row in the same table-->
        <xsl:choose>
          <xsl:when test="$table = $table2">
            <xsl:for-each select="following::table:table-row[1]">
              <xsl:call-template name="CheckIfColumnStyle">
                <xsl:with-param name="number" select="$number"/>
                <xsl:with-param name="table" select="$table"/>
                <xsl:with-param name="count" select="$count + $rows"/>
                <xsl:with-param name="value" select="$newValue"/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:when>
          <xsl:otherwise>
            <xsl:choose>
              <!-- when last possible row was reached -->
              <xsl:when test="$count + $rows = 65536">
                <xsl:text>true</xsl:text>
              </xsl:when>
              <xsl:otherwise>
                <xsl:text>false</xsl:text>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <!-- when cell in column $number didn't occur in the row-->
      <xsl:otherwise>
        <xsl:text>false</xsl:text>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- "true" when cell is non-empty, "false" when empty, "end" when after the end of a row -->
  <xsl:template name="CheckIfCellNonEmpty">
    <xsl:param name="number"/>
    <xsl:param name="count" select="1"/>
    <xsl:param name="value"/>

    <xsl:choose>
      <xsl:when test="$count &lt; $number">
        <xsl:variable name="columns">
          <xsl:choose>
            <xsl:when test="@table:number-columns-repeated">
              <xsl:value-of select="@table:number-columns-repeated"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>1</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:variable name="newValue">
          <xsl:choose>
            <xsl:when test="text:p">
              <xsl:text>true</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>false</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:choose>
          <xsl:when test="following-sibling::table:table-cell[1]">
            <xsl:for-each select="following-sibling::table:table-cell[1]">
              <xsl:call-template name="CheckIfCellNonEmpty">
                <xsl:with-param name="number">
                  <xsl:value-of select="$number"/>
                </xsl:with-param>
                <xsl:with-param name="count">
                  <xsl:value-of select="$count + $columns"/>
                </xsl:with-param>
                <xsl:with-param name="value">
                  <xsl:value-of select="$newValue"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:when>
          <!-- when cell is in last table:table-cell tag and it has columns-repeated attribute -->
          <xsl:when test="$columns &gt; 1 and $count + $columns &gt;= $number">
            <xsl:value-of select="$newValue"/>
          </xsl:when>
          <!-- when cell is after the end of the row -->
          <xsl:otherwise>
            <xsl:text>end</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$count = $number">
            <xsl:variable name="newValue">
              <xsl:choose>
                <xsl:when test="text:p">
                  <xsl:text>true</xsl:text>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:text>false</xsl:text>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:value-of select="$newValue"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$value"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template match="text()" mode="cell"/>
</xsl:stylesheet>

<?xml version="1.0" encoding="UTF-8"?>
<!--
  * Copyright (c) 2007, Clever Age
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
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:pxsi="urn:cleverage:xmlns:post-processings:pivotTable"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  exclude-result-prefixes="text">

  <xsl:template name="InsertPivotTable">
    <!-- @Context: table:data-pilot-table -->

    <pivotTableDefinition applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0"
      applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1"
      dataCaption="Data" updatedVersion="3" minRefreshableVersion="3" showCalcMbrs="0"
      useAutoFormatting="1" itemPrintTitles="1" createdVersion="3" indent="0" outline="1"
      outlineData="1" multipleFieldFilters="0">

      <xsl:attribute name="name">
        <xsl:value-of select="@table:name"/>
      </xsl:attribute>

      <xsl:if test="@table:grand-total = 'none' or @table:grand-total = 'column' ">
        <xsl:attribute name="rowGrandTotals">
          <xsl:text>0</xsl:text>
        </xsl:attribute>
      </xsl:if>

      <xsl:if test="@table:grand-total = 'none' or @table:grand-total = 'row' ">
        <xsl:attribute name="colGrandTotals">
          <xsl:text>0</xsl:text>
        </xsl:attribute>
      </xsl:if>

      <xsl:if
        test="table:data-pilot-field[@table:source-field-name = '' and @table:orientation = 'row' ]">
        <xsl:attribute name="dataOnRows">
          <xsl:text>1</xsl:text>
        </xsl:attribute>
      </xsl:if>

      <xsl:variable name="source">
        <xsl:value-of select="table:source-cell-range/@table:cell-range-address"/>
      </xsl:variable>

      <xsl:attribute name="cacheId">
        <!-- first pilot table on this source id -->
        <xsl:variable name="pivotSource">
          <xsl:value-of select="table:source-cell-range/@table:cell-range-address"/>
        </xsl:variable>

        <xsl:for-each select="key('pivot','')">
          <!-- do not duplicate the same source range cache -->
          <xsl:if
            test="table:source-cell-range/@table:cell-range-address = $pivotSource and 
            not(preceding-sibling::table:data-pilot-table[table:source-cell-range/@table:cell-range-address = $pivotSource])">
            <xsl:value-of select="count(preceding-sibling::table:data-pilot-table)"/>
          </xsl:if>
        </xsl:for-each>
      </xsl:attribute>

      <location firstHeaderRow="1" firstDataRow="2" firstDataCol="1">

        <xsl:variable name="pageFields">
          <xsl:value-of select="count(table:data-pilot-field[@table:orientation = 'page' ])"/>
        </xsl:variable>

        <xsl:if test="table:data-pilot-field[@table:orientation = 'page' ]">
          <xsl:attribute name="rowPageCount">
            <xsl:value-of select="$pageFields"/>
          </xsl:attribute>
          <xsl:attribute name="colPageCount">
            <xsl:text>1</xsl:text>
          </xsl:attribute>
        </xsl:if>

        <xsl:variable name="startRow">
          <xsl:call-template name="GetRowNum">
            <xsl:with-param name="cell">
              <xsl:value-of
                select="substring-after(substring-before(@table:target-range-address,':'),'.')"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>

        <xsl:choose>
          <xsl:when test="@table:show-filter-button = 'false' ">
            <xsl:attribute name="ref">
              <xsl:value-of
                select="substring-before(substring-after(substring-before(@table:target-range-address,':'),'.'),$startRow)"/>
              <xsl:value-of select="$startRow + 1 + $pageFields"/>
              <xsl:text>:</xsl:text>
              <xsl:value-of
                select="substring-after(substring-after(@table:target-range-address,':'),'.')"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>

            <xsl:attribute name="ref">
              <!-- column-->
              <xsl:value-of
                select="substring-before(substring-after(substring-before(@table:target-range-address,':'),'.'),$startRow)"/>
              <xsl:value-of select="$startRow + 2 + $pageFields"/>
              <xsl:text>:</xsl:text>
              <xsl:value-of
                select="substring-after(substring-after(@table:target-range-address,':'),'.')"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>

      </location>

      <pivotFields>

        <!--xsl:attribute name="count">

          <xsl:variable name="startCol">
            <xsl:call-template name="GetColNum">
              <xsl:with-param name="cell">
                <xsl:value-of
                  select="substring-after(substring-before(@table:target-range-address,':'),'.')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>

          <xsl:variable name="endCol">
            <xsl:call-template name="GetColNum">
              <xsl:with-param name="cell">
                <xsl:value-of
                  select="substring-after(substring-after(@table:target-range-address,':'),'.')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>

          <xsl:value-of select="$endCol - $startCol + 1"/>
        </xsl:attribute-->

        <xsl:variable name="names">
          <xsl:for-each select="table:data-pilot-field[@table:source-field-name != '' ]">
            <xsl:if test="position() &gt; 1 ">
              <xsl:text>~</xsl:text>
            </xsl:if>
            <xsl:value-of select="@table:source-field-name"/>
          </xsl:for-each>
        </xsl:variable>

        <xsl:variable name="axes">
          <xsl:for-each select="table:data-pilot-field[@table:source-field-name != '' ]">
            <xsl:if test="position() &gt; 1 ">
              <xsl:text>~</xsl:text>
            </xsl:if>
            <xsl:choose>
              <xsl:when test="@table:orientation = 'page' ">
                <xsl:text>axisPage</xsl:text>
              </xsl:when>
              <xsl:when test="@table:orientation = 'column' ">
                <xsl:text>axisCol</xsl:text>
              </xsl:when>
              <xsl:when test="@table:orientation = 'row' ">
                <xsl:text>axisRow</xsl:text>
              </xsl:when>
            </xsl:choose>
          </xsl:for-each>
        </xsl:variable>

        <xsl:variable name="sort">
          <xsl:for-each select="table:data-pilot-field[@table:source-field-name != '' ]">
            <xsl:if test="position() &gt; 1 ">
              <xsl:text>~</xsl:text>
            </xsl:if>
            <xsl:choose>
              <xsl:when
                test="table:data-pilot-level/table:data-pilot-sort-info[@table:order = 'descending' and @table:sort-mode != 'manual' and @table:sort-mode != 'none']">
                <xsl:text>descending</xsl:text>
              </xsl:when>
              <xsl:when
                test="table:data-pilot-level/table:data-pilot-sort-info[@table:order = 'ascending' and @table:sort-mode != 'manual' and @table:sort-mode != 'none'] ">
                <xsl:text>ascending</xsl:text>
              </xsl:when>
            </xsl:choose>
          </xsl:for-each>
        </xsl:variable>

        <xsl:variable name="hide">
          <xsl:for-each select="table:data-pilot-field[@table:source-field-name != '' ]">
            <xsl:if test="position() &gt; 1 ">
              <xsl:text>~</xsl:text>
            </xsl:if>
            <xsl:if
              test="table:data-pilot-level/table:data-pilot-members/table:data-pilot-member[@table:display = 'false' ]">
              <xsl:text>;</xsl:text>
              <xsl:for-each
                select="table:data-pilot-level/table:data-pilot-members/table:data-pilot-member[@table:display = 'false' ]">
                <xsl:value-of select="@table:name"/>
                <xsl:text>;</xsl:text>
              </xsl:for-each>
            </xsl:if>
          </xsl:for-each>
        </xsl:variable>

        <xsl:variable name="subtotals">
          <xsl:for-each select="table:data-pilot-field[@table:source-field-name != '' ]">
            <xsl:if test="position() &gt; 1 ">
              <xsl:text>~</xsl:text>
            </xsl:if>

            <xsl:if
              test="table:data-pilot-level/table:data-pilot-members/table:data-pilot-member[@table:display = 'false' ]">
              <xsl:text>;</xsl:text>

              <xsl:for-each
                select="table:data-pilot-level/table:data-pilot-subtotals/table:data-pilot-subtotal[@table:function]">

                <xsl:choose>
                  <xsl:when test="@table:function = 'sum' ">
                    <xsl:text>sum</xsl:text>
                  </xsl:when>
                  <xsl:when test="@table:function = 'count' ">
                    <xsl:text>count</xsl:text>
                  </xsl:when>
                  <xsl:when test="@table:function = 'average' ">
                    <xsl:text>average</xsl:text>
                  </xsl:when>
                  <xsl:when test="@table:function = 'max' ">
                    <xsl:text>max</xsl:text>
                  </xsl:when>
                  <xsl:when test="@table:function = 'min' ">
                    <xsl:text>min</xsl:text>
                  </xsl:when>
                  <xsl:when test="@table:function = 'product' ">
                    <xsl:text>productSubtotal</xsl:text>
                  </xsl:when>
                  <xsl:when test="@table:function = 'countnums' ">
                    <xsl:text>countnums</xsl:text>
                  </xsl:when>
                  <xsl:when test="@table:function = 'stdev' ">
                    <xsl:text>stdev</xsl:text>
                  </xsl:when>
                  <xsl:when test="@table:function = 'stdevp' ">
                    <xsl:text>stdevp</xsl:text>
                  </xsl:when>
                  <xsl:when test="@table:function = 'var' ">
                    <xsl:text>var</xsl:text>
                  </xsl:when>
                  <xsl:when test="@table:function = 'varp' ">
                    <xsl:text>varp</xsl:text>
                  </xsl:when>
                </xsl:choose>
                <xsl:text>;</xsl:text>

              </xsl:for-each>
            </xsl:if>
          </xsl:for-each>
        </xsl:variable>

        <pxsi:pivotFields>
          <xsl:attribute name="pxsi:names">
            <xsl:value-of select="$names"/>
          </xsl:attribute>

          <xsl:attribute name="pxsi:axes">
            <xsl:value-of select="$axes"/>
          </xsl:attribute>

          <xsl:attribute name="pxsi:sort">
            <xsl:value-of select="$sort"/>
          </xsl:attribute>

          <xsl:attribute name="pxsi:hide">
            <xsl:value-of select="$hide"/>
          </xsl:attribute>

          <xsl:attribute name="pxsi:subtotals">
            <xsl:value-of select="$subtotals"/>
          </xsl:attribute>
        </pxsi:pivotFields>

        <!-- pivotFields have names, the one field without name (definition?) can occur not only on the first place -->
        <!--xsl:for-each select="table:data-pilot-field[@table:source-field-name != '' ]">

          <pivotField showAll="0"-->

        <!-- field orientation -->
        <!--xsl:choose>
              <xsl:when test="@table:orientation = 'page' ">
                <xsl:attribute name="axis">
                  <xsl:text>axisPage</xsl:text>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="@table:orientation = 'column' ">
                <xsl:attribute name="axis">
                  <xsl:text>axisCol</xsl:text>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="@table:orientation = 'row' ">
                <xsl:attribute name="axis">
                  <xsl:text>axisRow</xsl:text>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="@table:orientation = 'data' ">
                <xsl:attribute name="dataField">
                  <xsl:text>1</xsl:text>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>


            <xsl:for-each
              select="table:data-pilot-level/table:data-pilot-subtotals/table:data-pilot-subtotal[@table:function]">

              <xsl:if test="@table:function = 'sum' ">
                <xsl:attribute name="sumSubtotal">
                  <xsl:text>1</xsl:text>
                </xsl:attribute>
              </xsl:if>

              <xsl:if test="@table:function = 'count' ">
                <xsl:attribute name="countASubtotal">
                  <xsl:text>1</xsl:text>
                </xsl:attribute>
              </xsl:if>

              <xsl:if test="@table:function = 'average' ">
                <xsl:attribute name="avgSubtotal">
                  <xsl:text>1</xsl:text>
                </xsl:attribute>
              </xsl:if>

              <xsl:if test="@table:function = 'max' ">
                <xsl:attribute name="maxSubtotal">
                  <xsl:text>1</xsl:text>
                </xsl:attribute>
              </xsl:if>

              <xsl:if test="@table:function = 'min' ">
                <xsl:attribute name="minSubtotal">
                  <xsl:text>1</xsl:text>
                </xsl:attribute>
              </xsl:if>

              <xsl:if test="@table:function = 'product' ">
                <xsl:attribute name="productSubtotal">
                  <xsl:text>1</xsl:text>
                </xsl:attribute>
              </xsl:if>


              <xsl:if test="@table:function = 'countnums' ">
                <xsl:attribute name="countSubtotal">
                  <xsl:text>1</xsl:text>
                </xsl:attribute>
              </xsl:if>

              <xsl:if test="@table:function = 'stdev' ">
                <xsl:attribute name="stdDevSubtotal">
                  <xsl:text>1</xsl:text>
                </xsl:attribute>
              </xsl:if>

              <xsl:if test="@table:function = 'stdevp' ">
                <xsl:attribute name="stdDevPSubtotal">
                  <xsl:text>1</xsl:text>
                </xsl:attribute>
              </xsl:if>

              <xsl:if test="@table:function = 'var' ">
                <xsl:attribute name="varSubtotal">
                  <xsl:text>1</xsl:text>
                </xsl:attribute>
              </xsl:if>

              <xsl:if test="@table:function = 'varp' ">
                <xsl:attribute name="varPSubtotal">
                  <xsl:text>1</xsl:text>
                </xsl:attribute>
              </xsl:if>

            </xsl:for-each>

            <xsl:choose>
              <xsl:when
                test="table:data-pilot-level/table:data-pilot-sort-info[@table:order = 'descending' and @table:sort-mode != 'manual' and @table:sort-mode != 'none']">
                <xsl:attribute name="sortType">
                  <xsl:text>descending</xsl:text>
                </xsl:attribute>
              </xsl:when>
              <xsl:when
                test="table:data-pilot-level/table:data-pilot-sort-info[@table:order = 'ascending' and @table:sort-mode != 'manual' and @table:sort-mode != 'none'] ">
                <xsl:attribute name="sortType">
                  <xsl:text>ascending</xsl:text>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>

            <xsl:if test="not(@table:orientation = 'data')">
              <items-->
        <!-- w tym miejscu trzeba dodać do postprocesora, ze jesli wystepuje 
                        table:data-pilot-level/table:data-pilot-subtotals/table:data-pilot-subtotal[@table:function]
                  to dla każdego @table:function stworzyć <item t='wartosc_funkcji'>-->
        <!-- 
                  z   'sum'             na     'sum'
                  z   'count'           na     'countA'
                  z   'average'      na     'avg'
                  z   'max'             na     'max'
                  z   'min'              na     'min'
                  z   'product'       na     'product'
                  z   'countnums'    na     'count'
                  z   'stdev'           na     'stdDev'
                  z   'stdevp'       na   'stdDevP'
                  z   'var'             na     'var'
                  z   'varp'           na     'varP'
                -->
        <!--pxsi:items pxsi:field="{@table:source-field-name}">
                  <xsl:if
                    test="table:data-pilot-level/table:data-pilot-members/table:data-pilot-member[@table:display = 'false' ]">
                    <xsl:attribute name="pxsi:hide">
                      <xsl:text>;</xsl:text>
                      <xsl:for-each
                        select="table:data-pilot-level/table:data-pilot-members/table:data-pilot-member[@table:display = 'false' ]">
                        <xsl:value-of select="@table:name"/>
                        <xsl:text>;</xsl:text>
                      </xsl:for-each>
                    </xsl:attribute>
                  </xsl:if>
                </pxsi:items>
                <item t="default"/>
              </items>
            </xsl:if>

          </pivotField>
        </xsl:for-each-->

        <xsl:variable name="sourceStartCol">
          <xsl:call-template name="GetColNum">
            <xsl:with-param name="cell">
              <xsl:value-of
                select="substring-after(substring-before(table:source-cell-range/@table:cell-range-address,':'),'.')"
              />
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>

        <xsl:variable name="sourceEndCol">
          <xsl:call-template name="GetColNum">
            <xsl:with-param name="cell">
              <xsl:value-of
                select="substring-after(substring-after(table:source-cell-range/@table:cell-range-address,':'),'.')"
              />
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>

      </pivotFields>

      <xsl:variable name="axedFields">
        <xsl:value-of
          select="count(table:data-pilot-field[@table:source-field-name != '' and @table:orientation != 'data' ])"
        />
      </xsl:variable>

      <xsl:variable name="dataFields">
        <xsl:value-of
          select="count(table:data-pilot-field[@table:source-field-name != '' and @table:orientation = 'data' ])"
        />
      </xsl:variable>

      <xsl:if
        test="table:data-pilot-field[@table:source-field-name != '' and @table:orientation = 'row' ] or $dataFields &gt;= 2">
        <rowFields count="{count(table:data-pilot-field[@table:orientation = 'row'])}">
          <xsl:for-each select="table:data-pilot-field[@table:orientation = 'row']">

            <xsl:choose>
              <!-- "Values" field -->
              <xsl:when test="@table:source-field-name = '' and $dataFields &gt;= 2">
                <field x="-2"/>
              </xsl:when>
              <xsl:when test="@table:source-field-name != '' ">
                <field>
                  <pxsi:fieldPos pxsi:name="{@table:source-field-name}" pxsi:attribute="x"/>
                </field>
              </xsl:when>
            </xsl:choose>
          </xsl:for-each>

        </rowFields>
      </xsl:if>

      <!-- TO DO <rowItems> -->

      <xsl:if test="table:data-pilot-field[@table:orientation = 'column' ]">
        <colFields
          count="{count(table:data-pilot-field[@table:source-field-name != '' and @table:orientation = 'column'])}">

          <xsl:for-each select="table:data-pilot-field[@table:orientation = 'column']">

            <xsl:choose>
              <!-- "Values" field -->
              <xsl:when test="@table:source-field-name = '' ">
                <field x="-2"/>
              </xsl:when>
              <xsl:when test="@table:source-field-name != '' ">
                <field>
                  <pxsi:fieldPos pxsi:name="{@table:source-field-name}" pxsi:attribute="x"/>
                </field>
              </xsl:when>
            </xsl:choose>
          </xsl:for-each>
        </colFields>
      </xsl:if>

      <!-- page fields -->
      <xsl:if test="table:data-pilot-field[@table:orientation = 'page' ]">
        <pageFields>
          <xsl:for-each select="table:data-pilot-field[@table:orientation = 'page' ]">
            <pageField hier="-1">

              <pxsi:fieldPos pxsi:name="{@table:source-field-name}" pxsi:attribute="fld"/>

              <xsl:if test="@table:selected-page">
                <pxsi:pageItem>
                  <xsl:attribute name="pxsi:pageField">
                    <xsl:value-of select="@table:source-field-name"/>
                  </xsl:attribute>
                  <xsl:attribute name="pxsi:pageItem">
                    <xsl:value-of select="@table:selected-page"/>
                  </xsl:attribute>
                </pxsi:pageItem>
              </xsl:if>

            </pageField>
          </xsl:for-each>
        </pageFields>
      </xsl:if>

      <!-- date fields -->
      <xsl:if
        test="table:data-pilot-field[@table:source-field-name != '' and @table:orientation = 'data' ]">
        <dataFields
          count="{count(table:data-pilot-field[@table:source-field-name != '' and @table:orientation = 'data'])}">
          <xsl:for-each
            select="table:data-pilot-field[@table:source-field-name != '' and @table:orientation = 'data']">
            <dataField>
              <xsl:attribute name="name">
                <xsl:value-of select="@table:source-field-name"/>
              </xsl:attribute>

              <xsl:if test="@table:function != 'sum' ">

                <xsl:attribute name="subtotal">
                  <xsl:choose>
                    <xsl:when test="@table:function = 'countnums' ">
                      <xsl:text>countNums</xsl:text>
                    </xsl:when>
                    <xsl:when test="@table:function = 'stdev' ">
                      <xsl:text>stdDev</xsl:text>
                    </xsl:when>
                    <xsl:when test="@table:function = 'stdevp' ">
                      <xsl:text>stdDevp</xsl:text>
                    </xsl:when>
                    <xsl:when test="@table:function != 'sum' ">
                      <xsl:value-of select="@table:function"/>
                    </xsl:when>
                  </xsl:choose>
                </xsl:attribute>

              </xsl:if>

              <!-- show data format as -->
              <xsl:for-each select="table:data-pilot-field-reference[@table:type]">

                <xsl:attribute name="showDataAs">

                  <xsl:choose>
                    <xsl:when test="@table:type= 'member-difference' ">
                      <xsl:text>difference</xsl:text>
                    </xsl:when>

                    <xsl:when test="@table:type= 'member-percentage' ">
                      <xsl:text>percent</xsl:text>
                    </xsl:when>

                    <xsl:when test="@table:type= 'member-percentage-difference' ">
                      <xsl:text>percentDiff</xsl:text>
                    </xsl:when>

                    <xsl:when test="@table:type= 'running-total' ">
                      <xsl:text>runTotal</xsl:text>
                    </xsl:when>

                    <xsl:when test="@table:type= 'row-percentage' ">
                      <xsl:text>percentOfRow</xsl:text>
                    </xsl:when>

                    <xsl:when test="@table:type= 'column-percentage' ">
                      <xsl:text>percentOfColumn</xsl:text>
                    </xsl:when>

                    <xsl:when test="@table:type= 'total-percentage' ">
                      <xsl:text>percentOfTotal</xsl:text>
                    </xsl:when>

                    <xsl:when test="@table:type= 'index' ">
                      <xsl:text>index</xsl:text>
                    </xsl:when>

                  </xsl:choose>
                </xsl:attribute>
              </xsl:for-each>

              <xsl:attribute name="baseField">
                <xsl:text>0</xsl:text>
              </xsl:attribute>

              <xsl:attribute name="baseItem">
                <xsl:text>0</xsl:text>
              </xsl:attribute>
              
              <pxsi:fieldPos pxsi:name="{@table:source-field-name}" pxsi:attribute="fld"/>
              
            </dataField>
          </xsl:for-each>
        </dataFields>
      </xsl:if>

      <pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1"
        showRowStripes="0" showColStripes="0" showLastColumn="1"/>
    </pivotTableDefinition>
  </xsl:template>

  <xsl:template name="InsertCacheDefinition">
    <!-- @Context: table:data-pilot-table -->
    <pivotCacheDefinition r:id="{generate-id(.)}" refreshedBy="Author"
      refreshedDate="39328.480206134256" createdVersion="3" refreshedVersion="3"
      minRefreshableVersion="3" recordCount="32">

      <cacheSource type="worksheet">
        <worksheetSource ref="E1:G33" sheet="Sheet2">
          <xsl:attribute name="ref">
            <xsl:value-of
              select="substring-after(substring-before(table:source-cell-range/@table:cell-range-address,':'),'.')"/>
            <xsl:text>:</xsl:text>
            <xsl:value-of
              select="substring-after(substring-after(table:source-cell-range/@table:cell-range-address,':'),'.')"
            />
          </xsl:attribute>

          <xsl:attribute name="sheet">
            <xsl:value-of
              select="substring-before(table:source-cell-range/@table:cell-range-address,'.')"/>
          </xsl:attribute>
        </worksheetSource>
      </cacheSource>

      <cacheFields>

        <pxsi:cacheFields/>

      </cacheFields>

    </pivotCacheDefinition>
  </xsl:template>

  <xsl:template name="InsertCacheRecords">
    <!-- @Context: table:data-pilot-table -->

    <pivotCacheRecords>

      <pxsi:cacheRecords/>

    </pivotCacheRecords>

  </xsl:template>

  <xsl:template name="GetCacheRows">
    <xsl:param name="rowStart"/>
    <xsl:param name="rowEnd"/>
    <xsl:param name="colStart"/>
    <xsl:param name="colEnd"/>
    <xsl:param name="rowNum" select="1"/>
    <xsl:param name="sheetName"/>

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

    <xsl:choose>
      <!-- skip rows before start of the source -->
      <xsl:when test="$rowNum + $rows &lt;= $rowStart">
        <xsl:for-each
          select="following::table:table-row[ancestor::table:table/@table:name = $sheetName][1]">
          <xsl:call-template name="GetCacheRows">
            <xsl:with-param name="rowStart" select="$rowStart"/>
            <xsl:with-param name="rowEnd" select="$rowEnd"/>
            <xsl:with-param name="colStart" select="$colStart"/>
            <xsl:with-param name="colEnd" select="$colEnd"/>
            <xsl:with-param name="rowNum" select="$rowNum + $rows"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test="$rowNum &lt;= $rowEnd">

        <xsl:choose>
          <xsl:when test="@table:number-rows-repeated &gt; 1">
            <xsl:call-template name="InsertRepeatedCacheRow">
              <xsl:with-param name="colStart" select="$colStart"/>
              <xsl:with-param name="colEnd" select="$colEnd"/>
              <xsl:with-param name="repeat" select="$rows"/>
              <xsl:with-param name="rowNum" select="$rowNum"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="InsertCacheRow">
              <xsl:with-param name="colStart" select="$colStart"/>
              <xsl:with-param name="colEnd" select="$colEnd"/>
              <xsl:with-param name="rowNum" select="$rowNum"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>

        <xsl:for-each
          select="following::table:table-row[ancestor::table:table/@table:name = $sheetName][1]">
          <xsl:call-template name="GetCacheRows">
            <xsl:with-param name="rowStart" select="$rowStart"/>
            <xsl:with-param name="rowEnd" select="$rowEnd"/>
            <xsl:with-param name="colStart" select="$colStart"/>
            <xsl:with-param name="colEnd" select="$colEnd"/>
            <xsl:with-param name="sheetName" select="$sheetName"/>
            <xsl:with-param name="rowNum" select="$rowNum + $rows"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
    </xsl:choose>

  </xsl:template>

  <xsl:template name="InsertRepeatedCacheRow">
    <xsl:param name="colStart"/>
    <xsl:param name="colEnd"/>
    <xsl:param name="repeat"/>
    <xsl:param name="count" select="1"/>
    <!-- temporary parameter -->
    <xsl:param name="rowNum"/>

    <xsl:if test="$count &lt;= $repeat">
      <xsl:call-template name="InsertCacheRow">
        <xsl:with-param name="colStart" select="$colStart"/>
        <xsl:with-param name="colEnd" select="$colEnd"/>
        <xsl:with-param name="rowNum" select="$rowNum"/>
      </xsl:call-template>

      <xsl:call-template name="InsertRepeatedCacheRow">
        <xsl:with-param name="colStart" select="$colStart"/>
        <xsl:with-param name="colEnd" select="$colEnd"/>
        <xsl:with-param name="repeat" select="$repeat"/>
        <xsl:with-param name="count" select="$count + 1"/>
        <xsl:with-param name="rowNum" select="$rowNum + 1"/>
      </xsl:call-template>
    </xsl:if>

  </xsl:template>

  <xsl:template name="InsertCacheRow">
    <xsl:param name="colStart"/>
    <xsl:param name="colEnd"/>
    <!-- temporary parameter -->
    <xsl:param name="rowNum"/>

    <r>
      <xsl:attribute name="c">
        <xsl:value-of select="$rowNum"/>
      </xsl:attribute>
      <xsl:call-template name="InsertCacheRowFields">
        <xsl:with-param name="colStart" select="$colStart"/>
        <xsl:with-param name="colEnd" select="$colEnd"/>
      </xsl:call-template>
    </r>

  </xsl:template>

  <xsl:template name="InsertCacheRowFields">
    <xsl:param name="colStart"/>
    <xsl:param name="colEnd"/>
    <xsl:param name="colNum" select="1"/>

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

    <xsl:choose>
      <!-- skip cells before start of the source -->
      <xsl:when test="$colNum &lt; $colStart and $colNum + $columns - 1 &lt; $colStart">
        <xsl:for-each select="following-sibling::node()[1]">
          <xsl:call-template name="InsertCacheRowFields">
            <xsl:with-param name="colStart" select="$colStart"/>
            <xsl:with-param name="colEnd" select="$colEnd"/>
            <xsl:with-param name="colNum" select="$colNum + $columns"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>

      <!-- when single column is inside source -->
      <xsl:when test="$columns = 1 and $colNum &lt;= $colEnd">

        <xsl:call-template name="InsertCacheRowField">
          <xsl:with-param name="colNum" select="$colNum"/>
        </xsl:call-template>

        <xsl:for-each select="following-sibling::node()[1]">
          <xsl:call-template name="InsertCacheRowFields">
            <xsl:with-param name="colStart" select="$colStart"/>
            <xsl:with-param name="colEnd" select="$colEnd"/>
            <xsl:with-param name="colNum" select="$colNum + $columns"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>

      <!-- when whole repeated column is inside source -->
      <xsl:when
        test="$columns &gt; 1 and $colNum &gt;= $colStart and $colNum + $columns - 1 &lt;= $colEnd">

        <xsl:call-template name="InsertRepeatedCacheRowFields">
          <xsl:with-param name="repeat" select="$columns"/>
          <xsl:with-param name="colNum" select="$colNum"/>
        </xsl:call-template>

        <xsl:for-each select="following-sibling::node()[1]">
          <xsl:call-template name="InsertCacheRowFields">
            <xsl:with-param name="colStart" select="$colStart"/>
            <xsl:with-param name="colEnd" select="$colEnd"/>
            <xsl:with-param name="colNum" select="$colNum + $columns"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>

      <!-- when whole source is inside repeated columns -->
      <xsl:when test="$colNum &lt;= $colStart and $colNum + $columns - 1 &gt;= $colEnd">

        <xsl:call-template name="InsertRepeatedCacheRowFields">
          <xsl:with-param name="repeat" select="$colEnd - $colStart + 1"/>
          <xsl:with-param name="colNum" select="$colNum"/>
        </xsl:call-template>

      </xsl:when>

      <!-- when source starts inside repeated columns 
        (if $colNum + $columns - 1 is also lower than $colEnd than it will be covered by the first when )-->
      <xsl:when test="$colNum &lt; $colStart and $colNum + $columns - 1 &lt;= $colEnd">

        <xsl:call-template name="InsertRepeatedCacheRowFields">
          <xsl:with-param name="repeat" select="$columns - ($colStart - $colNum)"/>
          <xsl:with-param name="colNum" select="$colNum"/>
        </xsl:call-template>

        <xsl:for-each select="following-sibling::node()[1]">
          <xsl:call-template name="InsertCacheRowFields">
            <xsl:with-param name="colStart" select="$colStart"/>
            <xsl:with-param name="colEnd" select="$colEnd"/>
            <xsl:with-param name="colNum" select="$colNum + $columns"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>

      <!-- when source ends inside repeated columns -->
      <xsl:when test="$colNum &lt;= $colEnd and $colNum + $columns - 1 &gt; $colEnd">

        <xsl:call-template name="InsertRepeatedCacheRowFields">
          <xsl:with-param name="repeat" select="$colEnd - $colNum + 1"/>
          <xsl:with-param name="colNum" select="$colNum"/>
        </xsl:call-template>

      </xsl:when>
    </xsl:choose>

  </xsl:template>

  <xsl:template name="InsertRepeatedCacheRowFields">
    <xsl:param name="repeat"/>
    <xsl:param name="count" select="1"/>
    <!-- temporary parameter -->
    <xsl:param name="colNum"/>

    <xsl:if test="$count &lt;= $repeat">
      <xsl:call-template name="InsertCacheRowField">
        <xsl:with-param name="colNum" select="$colNum"/>
      </xsl:call-template>

      <xsl:call-template name="InsertRepeatedCacheRowFields">
        <xsl:with-param name="repeat" select="$repeat"/>
        <xsl:with-param name="count" select="$count + 1"/>
        <xsl:with-param name="colNum" select="$colNum + 1"/>
      </xsl:call-template>
    </xsl:if>

  </xsl:template>

  <xsl:template name="InsertCacheRowField">
    <!-- temporary parameter -->
    <xsl:param name="colNum"/>

    <c>
      <xsl:value-of select="$colNum"/>
    </c>
  </xsl:template>

  <xsl:template name="WritePivotCells">

    <xsl:variable name="sheetName">
      <xsl:value-of select="@table:name"/>
    </xsl:variable>
    <xsl:text>;</xsl:text>

    <!-- do not process twice duplicated pivot source-->
    <xsl:for-each
      select="key('pivot','')[not(preceding-sibling::table:data-pilot-table/table:source-cell-range/@table:cell-range-address = table:source-cell-range/@table:cell-range-address)]">

      <!-- check only pivot sources on this sheet -->
      <xsl:if
        test="substring-before(table:source-cell-range/@table:cell-range-address, '.') = $sheetName ">

        <xsl:variable name="rowStart">
          <xsl:for-each select="table:source-cell-range">
            <xsl:call-template name="GetRowNum">
              <xsl:with-param name="cell">
                <xsl:value-of
                  select="substring-after(substring-before(@table:cell-range-address,':'),'.')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:variable>

        <xsl:variable name="colStart">
          <xsl:for-each select="table:source-cell-range">
            <xsl:call-template name="GetColNum">
              <xsl:with-param name="cell">
                <xsl:value-of
                  select="substring-after(substring-before(@table:cell-range-address,':'),'.')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:variable>

        <xsl:variable name="rowEnd">
          <xsl:for-each select="table:source-cell-range">
            <xsl:call-template name="GetRowNum">
              <xsl:with-param name="cell">
                <xsl:value-of
                  select="substring-after(substring-after(@table:cell-range-address,':'),'.')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:variable>

        <xsl:variable name="colEnd">
          <xsl:for-each select="table:source-cell-range">
            <xsl:call-template name="GetColNum">
              <xsl:with-param name="cell">
                <xsl:value-of
                  select="substring-after(substring-after(@table:cell-range-address,':'),'.')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:variable>

        <xsl:call-template name="ListOutCellsInRange">
          <xsl:with-param name="rowStart" select="$rowStart"/>
          <xsl:with-param name="rowEnd" select="$rowEnd"/>
          <xsl:with-param name="colStart" select="$colStart"/>
          <xsl:with-param name="colEnd" select="$colEnd"/>
        </xsl:call-template>
      </xsl:if>
    </xsl:for-each>

  </xsl:template>

  <xsl:template name="ListOutCellsInRange">
    <xsl:param name="rowStart"/>
    <xsl:param name="rowEnd"/>
    <xsl:param name="colStart"/>
    <xsl:param name="colEnd"/>
    <xsl:param name="row" select="$rowStart"/>

    <xsl:call-template name="ListOutCellsInRow">
      <xsl:with-param name="row" select="$row"/>
      <xsl:with-param name="colStart" select="$colStart"/>
      <xsl:with-param name="colEnd" select="$colEnd"/>
    </xsl:call-template>

    <xsl:if test="$row != $rowEnd">
      <xsl:call-template name="ListOutCellsInRange">
        <xsl:with-param name="rowStart" select="$rowStart"/>
        <xsl:with-param name="rowEnd" select="$rowEnd"/>
        <xsl:with-param name="colStart" select="$colStart"/>
        <xsl:with-param name="colEnd" select="$colEnd"/>
        <xsl:with-param name="row" select="$row + 1"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

  <xsl:template name="ListOutCellsInRow">
    <xsl:param name="row"/>
    <xsl:param name="colStart"/>
    <xsl:param name="colEnd"/>
    <xsl:param name="column" select="$colStart"/>

    <xsl:value-of select="$column"/>
    <xsl:text>:</xsl:text>
    <xsl:value-of select="$row"/>
    <xsl:text>;</xsl:text>

    <xsl:if test="$column != $colEnd">
      <xsl:call-template name="ListOutCellsInRow">
        <xsl:with-param name="colStart" select="$colStart"/>
        <xsl:with-param name="colEnd" select="$colEnd"/>
        <xsl:with-param name="row" select="$row"/>
        <xsl:with-param name="column" select="$column + 1"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

  <xsl:template match="text:p" mode="pivot">
    <xsl:if test="preceding-sibling::text:p[1]">
      <xsl:text> </xsl:text>
    </xsl:if>
    <xsl:apply-templates mode="pivot" xml:space="preserve"/>
  </xsl:template>

  <xsl:template match="text:span" mode="pivot">
    <xsl:apply-templates mode="pivot" xml:space="preserve"/>
  </xsl:template>

  <xsl:template match="text:a" mode="pivot">
    <xsl:apply-templates mode="pivot" xml:space="preserve"/>
  </xsl:template>

  <xsl:template match="text:s" mode="pivot">
    <pxs:s xmlns:pxs="urn:cleverage:xmlns:post-processings:extra-spaces">
      <xsl:attribute name="pxs:c">
        <xsl:choose>
          <xsl:when test="@text:c">
            <xsl:value-of select="@text:c"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>1</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </pxs:s>
  </xsl:template>

  <xsl:template match="text()" mode="pivot">
    <xsl:value-of select="."/>
  </xsl:template>

</xsl:stylesheet>

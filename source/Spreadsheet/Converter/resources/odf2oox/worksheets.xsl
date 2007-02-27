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
  xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">

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
        <xsl:with-param name="sheetId" select="$sheetId"/>
      </xsl:call-template>
    </pzip:entry>

    <!-- convert next table -->
    <xsl:apply-templates select="following-sibling::table:table[1]" mode="sheet">
      <xsl:with-param name="cellNumber">
        <xsl:value-of
          select="$cellNumber + count(table:table-row/table:table-cell[text:p and @office:value-type='string'])"
        />
      </xsl:with-param>
      <xsl:with-param name="sheetId">
        <xsl:value-of select="$sheetId + 1"/>
      </xsl:with-param>
    </xsl:apply-templates>

  </xsl:template>

  <!-- insert sheet -->
  <xsl:template name="InsertWorksheet">
    <xsl:param name="cellNumber"/>
    <xsl:param name="sheetId"/>
    <worksheet>

      <!-- compute default row height -->
      <xsl:variable name="defaultRowHeight">
        <xsl:choose>
          <xsl:when test="table:table-row[@table:number-rows-repeated > 32768]">
            <xsl:for-each select="table:table-row[@table:number-rows-repeated > 32768]">
              <xsl:call-template name="ConvertMeasure">
                <xsl:with-param name="length">
                  <xsl:value-of
                    select="key('style',@table:style-name)/style:table-row-properties/@style:row-height"
                  />
                </xsl:with-param>
                <xsl:with-param name="unit">point</xsl:with-param>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:when>
          <xsl:otherwise>13</xsl:otherwise>
        </xsl:choose>
      </xsl:variable>

      <!-- insert cursor position -->
      <sheetViews>
        <sheetView  workbookViewId="0">
          
          <xsl:variable name="ActiveTable">
            <xsl:for-each select="document('settings.xml')">
              <xsl:value-of select="key('ConfigItem', 'ActiveTable')"/>
            </xsl:for-each>
          </xsl:variable>
          
          <xsl:variable name="ActiveTableNumber">
            <xsl:for-each select="office:spreadsheet/table:table[@table:name=$ActiveTable]">
              <xsl:value-of select="count(preceding-sibling::table:table)"/>        
            </xsl:for-each>
          </xsl:variable>

          <xsl:if test="$sheetId = $ActiveTableNumber">
            <xsl:attribute name="activeTab"><xsl:text>1</xsl:text></xsl:attribute>
          </xsl:if>
          
          <selection>
          <xsl:variable name="col">
            <xsl:call-template name="NumbersToChars">
              <xsl:with-param name="num">
              <xsl:choose>
                <xsl:when test="document('settings.xml')/office:document-settings/office:settings/config:config-item-set[@config:name = 'ooo:view-settings']/config:config-item-map-indexed[@config:name = 'Views']/config:config-item-map-entry/config:config-item-map-named[@config:name='Tables']/config:config-item-map-entry[position() = $sheetId]/config:config-item[@config:name='CursorPositionX']">
                  <xsl:value-of select="document('settings.xml')/office:document-settings/office:settings/config:config-item-set[@config:name = 'ooo:view-settings']/config:config-item-map-indexed[@config:name = 'Views']/config:config-item-map-entry/config:config-item-map-named[@config:name='Tables']/config:config-item-map-entry[position() = $sheetId]/config:config-item[@config:name='CursorPositionX']"/>      
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
              </xsl:with-param>
           </xsl:call-template>
          </xsl:variable>
            <xsl:variable name="row">
              <xsl:choose>
                <xsl:when test="document('settings.xml')/office:document-settings/office:settings/config:config-item-set[@config:name = 'ooo:view-settings']/config:config-item-map-indexed[@config:name = 'Views']/config:config-item-map-entry/config:config-item-map-named[@config:name='Tables']/config:config-item-map-entry[position() = $sheetId]/config:config-item[@config:name='CursorPositionY']">
                  <xsl:value-of select="document('settings.xml')/office:document-settings/office:settings/config:config-item-set[@config:name = 'ooo:view-settings']/config:config-item-map-indexed[@config:name = 'Views']/config:config-item-map-entry/config:config-item-map-named[@config:name='Tables']/config:config-item-map-entry[position() = $sheetId]/config:config-item[@config:name='CursorPositionY']"/>    
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <!-- activeCell row value cannot be 0 -->
            <xsl:variable name="checkedRow">
              <xsl:choose>
                <xsl:when test="$row = 0">1</xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$row + 1"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>

            <xsl:attribute name="activeCell">
              <xsl:value-of select="concat($col,$checkedRow)"/>
            </xsl:attribute>
            <xsl:attribute name="sqref">
              <xsl:value-of select="concat($col,$checkedRow)"/>
            </xsl:attribute>
          </selection>
        </sheetView>
      </sheetViews>
      
      <sheetFormatPr defaultColWidth="12.40909090909091" defaultRowHeight="{$defaultRowHeight}"
        customHeight="true"/>
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
          <xsl:with-param name="defaultRowHeight" select="$defaultRowHeight"/>
        </xsl:apply-templates>

      </sheetData>

      <!-- Insert Merge Cells -->
      <xsl:call-template name="CheckMergeCell"/>

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
      <xsl:if
        test="key('style',@table:style-name)/style:table-column-properties/@style:column-width">
        <xsl:attribute name="width">
          <xsl:variable name="pixelWidth">
            <xsl:call-template name="pixel-measure">
              <xsl:with-param name="length">
                <xsl:value-of
                  select="key('style',@table:style-name)/style:table-column-properties/@style:column-width"
                />
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="fontSize">
            <xsl:text>11</xsl:text>
          </xsl:variable>
          <xsl:value-of select="($pixelWidth+5) div (2 div 3 * $fontSize)"/>
        </xsl:attribute>
        <xsl:attribute name="customWidth">1</xsl:attribute>
      </xsl:if>

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
    <xsl:param name="defaultRowHeight"/>
    <xsl:variable name="height">
      <xsl:call-template name="ConvertMeasure">
        <xsl:with-param name="length">
          <xsl:value-of
            select="key('style',@table:style-name)/style:table-row-properties/@style:row-height"/>
        </xsl:with-param>
        <xsl:with-param name="unit">point</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:if
      test="table:table-cell or @table:visibility='collapse' or  @table:visibility='filter' or $height != $defaultRowHeight or table:covered-table-cell">
      <row r="{$rowNumber}">

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
        <xsl:choose>
          <xsl:when test="node()[1][name()='table:table-cell']">
          <xsl:apply-templates select="table:table-cell[1]" mode="sheet">
          <xsl:with-param name="colNumber">0</xsl:with-param>
          <xsl:with-param name="rowNumber" select="$rowNumber"/>
          <xsl:with-param name="cellNumber" select="$cellNumber"/>
          </xsl:apply-templates>
          </xsl:when>
          <xsl:when test="node()[1][name()='table:covered-table-cell']">            
          </xsl:when>
          <xsl:otherwise/>
        </xsl:choose>
        
      </row>

      <!-- insert repeated rows -->
      <xsl:if test="@table:number-rows-repeated">
        <xsl:call-template name="InsertRepeatedRows">
          <xsl:with-param name="numberRowsRepeated">
            <xsl:value-of select="@table:number-rows-repeated"/>
          </xsl:with-param>
          <xsl:with-param name="numberOfAllRowsRepeated">
            <xsl:value-of select="@table:number-rows-repeated"/>
          </xsl:with-param>
          <xsl:with-param name="rowNumber" select="$rowNumber"/>
          <xsl:with-param name="height" select="$height"/>
          <xsl:with-param name="defaultRowHeight" select="$defaultRowHeight"/>
        </xsl:call-template>
      </xsl:if>
      
    </xsl:if>

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
          <xsl:value-of
            select="$cellNumber + count(child::table:table-cell[text:p and @office:value-type='string'])"
          />
        </xsl:with-param>
        <xsl:with-param name="defaultRowHeight" select="$defaultRowHeight"/>
      </xsl:apply-templates>
    </xsl:if>

  </xsl:template>

  <!-- template which inserts repeated rows -->
  <xsl:template name="InsertRepeatedRows">
    <xsl:param name="numberRowsRepeated"/>
    <xsl:param name="numberOfAllRowsRepeated"/>
    <xsl:param name="rowNumber"/>
    <xsl:param name="height"/>
    <xsl:param name="defaultRowHeight"/>
    <xsl:choose>
      <xsl:when test="$numberRowsRepeated &gt; 1">
        <xsl:if
          test="table:table-cell/text:p or @table:visibility='collapse' or  @table:visibility='filter' or ($height != $defaultRowHeight)">
          <row r="{$rowNumber + 1 + $numberOfAllRowsRepeated - $numberRowsRepeated}">

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
          </row>
        </xsl:if>

        <!-- insert repeated rows -->
        <xsl:if test="@table:number-rows-repeated">
          <xsl:call-template name="InsertRepeatedRows">
            <xsl:with-param name="numberRowsRepeated">
              <xsl:value-of select="$numberRowsRepeated - 1"/>
            </xsl:with-param>
            <xsl:with-param name="numberOfAllRowsRepeated" select="$numberOfAllRowsRepeated"/>
            <xsl:with-param name="rowNumber" select="$rowNumber"/>
            <xsl:with-param name="height" select="$height"/>
            <xsl:with-param name="defaultRowHeight" select="$defaultRowHeight"/>
          </xsl:call-template>
        </xsl:if>

      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- insert cell into row -->
 <xsl:template match="table:table-cell" mode="sheet">
    <xsl:param name="colNumber"/>
    <xsl:param name="rowNumber"/>
    <xsl:param name="cellNumber"/>
    <xsl:param name="ColumnRepeated"/>

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
    </xsl:call-template>

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
    </xsl:call-template>
   
  </xsl:template>
  
  <!-- insert cell -->
  <xsl:template name="InsertNextCell">
    <xsl:param name="colNumber"/>
    <xsl:param name="rowNumber"/>
    <xsl:param name="cellNumber"/>
   
    <xsl:if test="following-sibling::table:table-cell or following-sibling::table:covered-table-cell">
      <xsl:choose>
        <!-- Checks if  next index is table:table-cell-->
        <xsl:when test="following-sibling::node()[1][name()='table:table-cell']">          
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
                <xsl:when test="child::text:p and @office:value-type='string'">
                  <xsl:value-of select="$cellNumber + 1"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$cellNumber"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
          </xsl:apply-templates>
        </xsl:when>
        <!--  Checks if next index is table:covered-table-cell-->
        <xsl:when test="following-sibling::node()[1][name()='table:covered-table-cell']">          
          <xsl:apply-templates select="following-sibling::table:covered-table-cell[1]" mode="sheet">
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
                <xsl:when test="child::text:p and @office:value-type='string'">
                  <xsl:value-of select="$cellNumber + 1"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$cellNumber"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
            <xsl:with-param name="ColumnRepeated">1</xsl:with-param>
          </xsl:apply-templates>
        </xsl:when>  
      </xsl:choose>
      </xsl:if>
  </xsl:template>
  
    <xsl:template match="table:covered-table-cell" mode="sheet">
    <xsl:param name="colNumber"/>
    <xsl:param name="rowNumber"/>
    <xsl:param name="cellNumber"/>
    <xsl:param name="ColumnRepeated"/>
    
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
      <xsl:with-param name="ColumnRepeated">
        <xsl:value-of select="$ColumnRepeated"/>
      </xsl:with-param>
    </xsl:call-template>
    
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
    </xsl:call-template>    
  </xsl:template>
  
  <!-- Insert empty Cell for "table:covered-table-cell"-->
  <xsl:template name="InsertConvertCell">
    <xsl:param name="colNumber"/>
    <xsl:param name="rowNumber"/>
    <xsl:param name="cellNumber"/>
    <xsl:param name="ColumnRepeated"/>
    
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
    </xsl:call-template>
    
    <!-- Insert cells if "@table:number-columns-repeated"  > 1 -->   
    <xsl:if test="@table:number-columns-repeated and number(@table:number-columns-repeated) &gt; $ColumnRepeated">      
      <xsl:call-template name="InsertConvertCell">
        <xsl:with-param name="colNumber">
          <xsl:value-of select="$colNumber+1"/>
        </xsl:with-param>
        <xsl:with-param name="rowNumber" select="$rowNumber"/>
        <xsl:with-param name="cellNumber">
          <xsl:choose>
            <xsl:when test="child::text:p and @office:value-type='string'">
              <xsl:value-of select="$cellNumber + 1"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$cellNumber"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
        <xsl:with-param name="ColumnRepeated"><xsl:value-of select="$ColumnRepeated + 1"/></xsl:with-param>
      </xsl:call-template>
    </xsl:if> 
  </xsl:template>
  
  <xsl:template name="InsertCell">
    <xsl:param name="colNumber"/>
    <xsl:param name="rowNumber"/>
    <xsl:param name="cellNumber"/>    
    <c>
      <xsl:attribute name="r">
        <xsl:variable name="colChar">
          <xsl:call-template name="NumbersToChars">
            <xsl:with-param name="num" select="$colNumber"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:value-of select="concat($colChar,$rowNumber)"/>
      </xsl:attribute>
      
      <!-- insert cell style number-->
      <xsl:if test="@table:style-name">
        <xsl:for-each select="key('style',@table:style-name)">
          <xsl:attribute name="s">
            <xsl:number count="style:style[@style:family='table-cell']"
              level="any"/>
          </xsl:attribute>
        </xsl:for-each>
      </xsl:if>
      
      <!-- convert cell type -->
      <xsl:if test="child::text:p">
        <xsl:choose>
          <xsl:when test="@office:value-type!='string'">
            <xsl:attribute name="t">
              <xsl:call-template name="ConvertTypes">
                <xsl:with-param name="type">
                  <xsl:value-of select="@office:value-type"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
            <v>
              <xsl:value-of select="text:p"/>
            </v>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="t">s</xsl:attribute>
            <v>
              <xsl:value-of select="number($cellNumber)"/>
            </v>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
    </c>
  </xsl:template>

</xsl:stylesheet>

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
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" exclude-result-prefixes="table">

  <xsl:import href="measures.xsl"/>
  <xsl:import href="pixel-measure.xsl"/>
  <xsl:import href="page.xsl"/>
  <xsl:import href="border.xsl"/>

  <xsl:key name="table-row" match="table:table-row" use=" '' "/>
  <xsl:key name="StyleFamily" match="style:style" use="@style:family"/>
  <xsl:key name="tableMasterPage" match="style:style[@style:family='table']" use="@style:name"/>
  <xsl:key name="ConfigItem"
    match="office:document-settings/office:settings/config:config-item-set[@config:name = 'ooo:view-settings']/config:config-item-map-indexed[@config:name = 'Views']/config:config-item-map-entry/config:config-item"
    use="@config:name"/>

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
        <!-- last 'or' for cells with error -->
        <!-- cellNumber + number string cells inside simple rows + number string cells inside header rows -->
        <xsl:value-of
          select="$cellNumber + count(table:table-row/table:table-cell[text:p and not(@office:value-type='float') and (@office:value-type='string' or @office:value-type='boolean' or not((number(text:p) or text:p = 0 or contains(text:p,',') or contains(text:p,'%') or @office:value-type='currency' or @office:value-type='date' or @office:value-type='time')))]) +
          count(table:table-header-rows/table:table-row/table:table-cell[text:p and not(@office:value-type='float') and (@office:value-type='string' or @office:value-type='boolean' or not((number(text:p) or text:p = 0 or contains(text:p,',') or contains(text:p,'%') or @office:value-type='currency' or @office:value-type='date' or @office:value-type='time')))])"
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

      <xsl:variable name="MergeCell">
        <xsl:call-template name="WriteMergeCell"/>
      </xsl:variable>

      <xsl:variable name="MergeCellStyle">
        <xsl:call-template name="WriteMergeStyle"/>
      </xsl:variable>

      <xsl:variable name="masterPage">
        <xsl:value-of select="key('tableMasterPage',@table:style-name)/@style:master-page-name"/>
      </xsl:variable>

      <xsl:variable name="pageStyle">
        <xsl:value-of
          select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name =  $masterPage]/@style:page-layout-name"
        />
      </xsl:variable>

      <!-- if property 'fit print range(s) to width/height' is being used -->
      <xsl:for-each select="document('styles.xml')">
        <xsl:if
          test="key('pageStyle',$pageStyle)/style:page-layout-properties/@style:scale-to-X or key('pageStyle',$pageStyle)/style:page-layout-properties/@style:scale-to-Y">
          <sheetPr>
            <pageSetUpPr fitToPage="1"/>
          </sheetPr>
        </xsl:if>
      </xsl:for-each>

      <xsl:call-template name="InsertViewSettings">
        <xsl:with-param name="sheetId" select="$sheetId"/>
        <xsl:with-param name="MergeCell">
          <xsl:value-of select="$MergeCell"/>
        </xsl:with-param>
        <xsl:with-param name="MergeCellStyle">
          <xsl:value-of select="$MergeCellStyle"/>
        </xsl:with-param>
      </xsl:call-template>

      <xsl:call-template name="InsertSheetContent">
        <xsl:with-param name="sheetId" select="$sheetId"/>
        <xsl:with-param name="cellNumber" select="$cellNumber"/>
        <xsl:with-param name="MergeCell">
          <xsl:value-of select="$MergeCell"/>
        </xsl:with-param>
        <xsl:with-param name="MergeCellStyle">
          <xsl:value-of select="$MergeCellStyle"/>
        </xsl:with-param>
      </xsl:call-template>

      <!-- Insert Merge Cells -->
      <xsl:call-template name="InsertMergeCells">
        <xsl:with-param name="MergeCell">
          <xsl:value-of select="$MergeCell"/>
        </xsl:with-param>
        <xsl:with-param name="MergeCellStyle">
          <xsl:value-of select="$MergeCellStyle"/>
        </xsl:with-param>
      </xsl:call-template>

      <xsl:call-template name="InsertPageProperties">
        <xsl:with-param name="pageStyle" select="$pageStyle"/>
      </xsl:call-template>

      <xsl:call-template name="InsertHeaderFooter"/>
      
      <xsl:variable name="picture">
        <xsl:choose>
          <xsl:when test="table:table-row/table:table-cell/draw:frame/draw:image">
            <xsl:text>true</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>false</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      
      <xsl:variable name="chart">
        <xsl:for-each select="descendant::draw:frame/draw:object">
          <xsl:for-each select="document(concat(translate(@xlink:href,'./',''),'/content.xml'))">
            <xsl:choose>
              <xsl:when test="office:document-content/office:body/office:chart">
                <xsl:text>true</xsl:text>
              </xsl:when>
              <xsl:otherwise>
                <xsl:text>false</xsl:text>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </xsl:for-each>
      </xsl:variable>
      
      <xsl:if test="contains($chart,'true') or $picture = 'true'">
        <drawing r:id="{concat('d_rId',$sheetId)}"/>
      </xsl:if>      
      
      <xsl:if test="descendant::office:annotation">
        <legacyDrawing r:id="{concat('v_rId',$sheetId)}"/>
      </xsl:if>
      
    </worksheet>
  </xsl:template>

  <xsl:template name="InsertViewSettings">
    <xsl:param name="sheetId"/>
    <xsl:param name="MergeCell"/>
    <xsl:param name="MergeCellStyle"/>

    <sheetViews>
      <sheetView workbookViewId="0">

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

        <xsl:variable name="pageBreakView">
          <xsl:for-each select="document('settings.xml')">
            <xsl:choose>
              <xsl:when test="key('ConfigItem', 'ShowPageBreakPreview')">

                <xsl:value-of select="key('ConfigItem', 'ShowPageBreakPreview')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:text>false</xsl:text>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </xsl:variable>

        <xsl:for-each select="document('settings.xml')">

          <xsl:variable name="hasColumnRowHeaders">
            <xsl:choose>
              <xsl:when test="key('ConfigItem', 'HasColumnRowHeaders')">

                <xsl:value-of select="key('ConfigItem', 'HasColumnRowHeaders')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:text>true</xsl:text>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>

          <xsl:if test="$hasColumnRowHeaders='false'">
            <xsl:attribute name="showRowColHeaders">
              <xsl:text>0</xsl:text>
            </xsl:attribute>
          </xsl:if>

          <!-- if it is normal view than take zoom from ZoomValue; if it's a PageBreakView then from PageViewZoomValue -->
          <xsl:variable name="zoom">
            <xsl:choose>
              <!-- normal view-->
              <xsl:when test="$pageBreakView = 'false' ">
                <xsl:choose>
                  <xsl:when test="key('ConfigItem', 'ZoomValue')">
                    <xsl:value-of select="key('ConfigItem', 'ZoomValue')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:text>100</xsl:text>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>
              <!-- PageBreakView -->
              <xsl:otherwise>
                <!-- take zoom value from PageViewZoomValue -->
                <xsl:choose>
                  <xsl:when test="key('ConfigItem', 'PageViewZoomValue')">
                    <xsl:value-of select="key('ConfigItem', 'PageViewZoomValue')"/>
                  </xsl:when>
                  <!-- if there isn't PageViewZoomValue take zoom value from ZoomValue -->
                  <xsl:when test="key('ConfigItem', 'ZoomValue')">
                    <xsl:value-of select="key('ConfigItem', 'ZoomValue')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:text>100</xsl:text>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>

          <xsl:if test="$zoom">
            <xsl:attribute name="zoomScale">
              <xsl:value-of select="$zoom"/>
            </xsl:attribute>
          </xsl:if>

        </xsl:for-each>

        <xsl:if test="$sheetId = $ActiveTableNumber">
          <xsl:attribute name="activeTab">
            <xsl:text>1</xsl:text>
          </xsl:attribute>
        </xsl:if>

        <xsl:if test="$pageBreakView = 'true'">
          <xsl:attribute name="view">
            <xsl:text>pageBreakPreview</xsl:text>
          </xsl:attribute>
        </xsl:if>

        <selection>

          <xsl:variable name="col">
            <xsl:call-template name="NumbersToChars">
              <xsl:with-param name="num">
                <xsl:choose>
                  <xsl:when
                    test="document('settings.xml')/office:document-settings/office:settings/config:config-item-set[@config:name = 'ooo:view-settings']/config:config-item-map-indexed[@config:name = 'Views']/config:config-item-map-entry/config:config-item-map-named[@config:name='Tables']/config:config-item-map-entry[position() = $sheetId]/config:config-item[@config:name='CursorPositionX']">
                    <xsl:value-of
                      select="document('settings.xml')/office:document-settings/office:settings/config:config-item-set[@config:name = 'ooo:view-settings']/config:config-item-map-indexed[@config:name = 'Views']/config:config-item-map-entry/config:config-item-map-named[@config:name='Tables']/config:config-item-map-entry[position() = $sheetId]/config:config-item[@config:name='CursorPositionX']"
                    />
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:text>1</xsl:text>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="row">
            <xsl:choose>
              <xsl:when
                test="document('settings.xml')/office:document-settings/office:settings/config:config-item-set[@config:name = 'ooo:view-settings']/config:config-item-map-indexed[@config:name = 'Views']/config:config-item-map-entry/config:config-item-map-named[@config:name='Tables']/config:config-item-map-entry[position() = $sheetId]/config:config-item[@config:name='CursorPositionY']">
                <xsl:value-of
                  select="document('settings.xml')/office:document-settings/office:settings/config:config-item-set[@config:name = 'ooo:view-settings']/config:config-item-map-indexed[@config:name = 'Views']/config:config-item-map-entry/config:config-item-map-named[@config:name='Tables']/config:config-item-map-entry[position() = $sheetId]/config:config-item[@config:name='CursorPositionY']"
                />
              </xsl:when>
              <xsl:otherwise>
                <xsl:text>1</xsl:text>
              </xsl:otherwise>
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
  </xsl:template>

  <xsl:template name="InsertSheetContent">
    <xsl:param name="sheetId"/>
    <xsl:param name="cellNumber"/>
    <xsl:param name="MergeCell"/>
    <xsl:param name="MergeCellStyle"/>

    <!-- compute default row height -->
    <xsl:variable name="defaultRowHeight">
      <xsl:choose>
        <xsl:when test="descendant::table:table-row[@table:number-rows-repeated > 32768]">
          <xsl:for-each select="descendant::table:table-row[@table:number-rows-repeated > 32768]">
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

    <!-- get default font size -->
    <xsl:variable name="baseFontSize">
      <xsl:for-each select="document('styles.xml')">
        <xsl:choose>
          <xsl:when
            test="office:document-styles/office:styles/style:style[@style:name='Default' and @style:family = 'table-cell']/style:text-properties/@fo:font-size">
            <xsl:value-of
              select="office:document-styles/office:styles/style:style[@style:name='Default' and @style:family = 'table-cell']/style:text-properties/@fo:font-size"
            />
          </xsl:when>
          <xsl:otherwise>10</xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
    </xsl:variable>

    <!-- Check if 256 column are hidden -->
    <xsl:variable name="CheckCollHidden">
      <xsl:apply-templates select="descendant::table:table-column[1]" mode="defaultColWidth">
        <xsl:with-param name="colNumber">
          <xsl:text>1</xsl:text>
        </xsl:with-param>
        <xsl:with-param name="result">
          <xsl:text>false</xsl:text>
        </xsl:with-param>
      </xsl:apply-templates>
    </xsl:variable>

    <!-- Check if default border areexisted in default column-->
    <xsl:variable name="CheckIfDefaultBorder">
      <xsl:apply-templates select="descendant::table:table-column[1]" mode="DefaultBorder"/>
    </xsl:variable>

    <!-- Check if 65536 rows are hidden -->
    <xsl:variable name="CheckRowHidden">
      <xsl:choose>
        <xsl:when test="table:table-row[@table:visibility='collapse']">
          <xsl:apply-templates select="descendant::table:table-row[1]" mode="zeroHeight">
            <xsl:with-param name="rowNumber">
              <xsl:text>0</xsl:text>
            </xsl:with-param>
          </xsl:apply-templates>
        </xsl:when>
        <xsl:otherwise>
          <xsl:text>false</xsl:text>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="defaultFontSize">
      <xsl:choose>
        <xsl:when test="contains($baseFontSize,'pt')">
          <xsl:value-of select="substring-before($baseFontSize,'pt')"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$baseFontSize"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <!-- compute default column width -->
    <xsl:variable name="defaultColWidth">
      <xsl:choose>
        <xsl:when test="$CheckCollHidden != 'true' ">
          <xsl:call-template name="ConvertToCharacters">
            <xsl:with-param name="width">
              <xsl:value-of select="concat('0.8925','in')"/>
            </xsl:with-param>
            <xsl:with-param name="defaultFontSize" select="$defaultFontSize"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:text>0</xsl:text>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <sheetFormatPr>

      <xsl:attribute name="defaultColWidth">
        <xsl:value-of select="$defaultColWidth"/>
      </xsl:attribute>

      <xsl:attribute name="defaultRowHeight">
        <xsl:value-of select="$defaultRowHeight"/>
      </xsl:attribute>

      <xsl:attribute name="customHeight">
        <xsl:text>true</xsl:text>
      </xsl:attribute>

      <xsl:if test="$CheckRowHidden = 'true'">
        <xsl:attribute name="zeroHeight">
          <xsl:text>1</xsl:text>
        </xsl:attribute>
      </xsl:if>

    </sheetFormatPr>

    <xsl:if test="descendant::table:table-column[1]">
      <cols>

        <!-- insert first column -->
        <xsl:apply-templates select="descendant::table:table-column[1]" mode="sheet">
          <xsl:with-param name="colNumber">1</xsl:with-param>
          <xsl:with-param name="defaultFontSize" select="$defaultFontSize"/>
        </xsl:apply-templates>

      </cols>
    </xsl:if>
    <sheetData>

      <xsl:variable name="ColumnTagNum">
        <xsl:apply-templates select="descendant::table:table-column[1]" mode="tag">
          <xsl:with-param name="colNumber">1</xsl:with-param>
          <xsl:with-param name="defaultFontSize" select="$defaultFontSize"/>
        </xsl:apply-templates>
      </xsl:variable>

      <!-- insert first row -->
      <xsl:apply-templates select="descendant::table:table-row[1]" mode="sheet">
        <xsl:with-param name="rowNumber">1</xsl:with-param>
        <xsl:with-param name="cellNumber" select="$cellNumber"/>
        <xsl:with-param name="defaultRowHeight" select="$defaultRowHeight"/>
        <xsl:with-param name="TableColumnTagNum">
          <xsl:value-of select="$ColumnTagNum"/>
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

    </sheetData>
  </xsl:template>

  <xsl:template name="InsertHeaderFooter">
    <xsl:if
      test="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name = 'Default']/style:header[not(@style:display = 'false' )]/child::node() or
      document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name = 'Default']/style:footer[not(@style:display = 'false' )]/child::node()">
      <headerFooter>
        <xsl:for-each
          select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name = 'Default']">
          <xsl:if
            test="not(style:header-left/@style:display = 'false' ) or not(style:footer-left/@style:display = 'false' )">
            <xsl:attribute name="differentOddEven">
              <xsl:text>1</xsl:text>
            </xsl:attribute>
          </xsl:if>
          <xsl:if
            test="not(style:header-left/@style:display = 'false') or style:header-left/child::node[1]/text() != '' or not(style:footer-left/@style:display = 'false') or style:footer-left/child::node[1]/text() != '' or not(style:header/@style:display = 'false') or style:header/child::node[1]/text() != '' or not(style:footer/@style:display = 'false') or style:footer/child::node[1]/text() != '' ">
            <xsl:message terminate="no">translation.odf2oox.HeaderFooterCharNumber</xsl:message>
          </xsl:if>
        </xsl:for-each>
        <xsl:for-each select="document('styles.xml')/office:document-styles">
          <xsl:call-template name="OddHeaderFooter"/>
          <xsl:call-template name="EvenHeaderFooter"/>
        </xsl:for-each>
      </headerFooter>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertPageProperties">
    <xsl:param name="pageStyle"/>

    <xsl:for-each select="document('styles.xml')">
      <xsl:for-each select="key('pageStyle',$pageStyle)/style:page-layout-properties">
        <xsl:if test="@style:table-centering">
          <printOptions>
            <!-- table horizontal alignment -->
            <xsl:if test="@style:table-centering = 'horizontal' or @style:table-centering = 'both' ">
              <xsl:attribute name="horizontalCentered">
                <xsl:text>1</xsl:text>
              </xsl:attribute>
            </xsl:if>

            <!-- table vertical alignment -->
            <xsl:if test="@style:table-centering = 'vertical' or @style:table-centering = 'both' ">
              <xsl:attribute name="verticalCentered">
                <xsl:text>1</xsl:text>
              </xsl:attribute>
            </xsl:if>

            <!-- headings -->
            <xsl:if test="@style:print and contains(@style:print,'headers' )">
              <xsl:attribute name="headings">
                <xsl:text>1</xsl:text>
              </xsl:attribute>
            </xsl:if>

            <!-- headings -->
            <xsl:if test="@style:print and contains(@style:print,'grid' )">
              <xsl:attribute name="gridLines">
                <xsl:text>1</xsl:text>
              </xsl:attribute>
            </xsl:if>

          </printOptions>
        </xsl:if>

        <!-- page margins -->
        <xsl:call-template name="InsertMargins">
          <xsl:with-param name="pageStyle" select="$pageStyle"/>
        </xsl:call-template>

        <xsl:if
          test="(@fo:page-width and @fo:page-height) or @style:scale-to or @style:first-page-number or @style:scale-to-X or style:scale-to-Y or @style:print-page-order or @style:print-orientation or @style:print-page-order">
          <pageSetup>
            <!-- paper size -->
            <xsl:if test="@fo:page-width and @fo:page-height">
              <xsl:attribute name="paperSize">
                <xsl:call-template name="TranslatePaperSize">
                  <xsl:with-param name="width">
                    <xsl:value-of
                      select="format-number(substring-before(@fo:page-width,'cm'),'#.##')"/>
                  </xsl:with-param>
                  <xsl:with-param name="height">
                    <xsl:value-of
                      select="format-number(substring-before(@fo:page-height,'cm'),'#.##')"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>

            <!-- scale by factor -->
            <xsl:if test="@style:scale-to">
              <xsl:attribute name="scale">
                <xsl:choose>
                  <xsl:when test="contains(@style:scale-to,'%')">
                    <xsl:value-of select="substring-before(@style:scale-to,'%')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="@style:scale-to"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>

            <!-- first page number -->
            <xsl:if test="@style:first-page-number and number(@style:first-page-number)">
              <xsl:attribute name="firstPageNumber">
                <xsl:value-of select="@style:first-page-number"/>
              </xsl:attribute>
              <xsl:attribute name="useFirstPageNumber">
                <xsl:text>1</xsl:text>
              </xsl:attribute>
            </xsl:if>

            <!-- fit print range(s) to width/height -->
            <xsl:if test="@style:scale-to-X">
              <xsl:attribute name="fitToWidth">
                <xsl:value-of select="@style:scale-to-X"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="@style:scale-to-Y">
              <xsl:attribute name="fitToHeight">
                <xsl:value-of select="@style:scale-to-Y"/>
              </xsl:attribute>
            </xsl:if>

            <!-- page order-->
            <xsl:if test="@style:print-page-order = 'ltr' ">
              <xsl:attribute name="pageOrder">
                <xsl:text>overThenDown</xsl:text>
              </xsl:attribute>
            </xsl:if>

            <!-- paper orientation -->
            <xsl:if test="@style:print-orientation">
              <xsl:attribute name="orientation">
                <xsl:value-of select="@style:print-orientation"/>
              </xsl:attribute>
            </xsl:if>

            <!-- notes -->
            <xsl:if test="@style:print and contains(@style:print,'annotations' )">
              <xsl:attribute name="cellComments">
                <xsl:text>atEnd</xsl:text>
              </xsl:attribute>
            </xsl:if>
          </pageSetup>
        </xsl:if>
      </xsl:for-each>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="InsertMargins">
    <xsl:param name="pageStyle"/>

    <xsl:if test="@fo:margin-top or @fo:margin-bottom or @fo:margin-right or @fo:margin-left">

      <pageMargins left="0.78740157480314965" right="0.70866141732283472" top="0.74803149606299213"
        bottom="0.74803149606299213" header="0.31496062992125984" footer="0.31496062992125984">
        <xsl:if test="@fo:margin-left">
          <xsl:attribute name="left">
            <!-- 1 inch = 1440 twips -->
            <xsl:variable name="twips">
              <xsl:call-template name="ConvertMeasure">
                <xsl:with-param name="length" select="@fo:margin-left"/>
                <xsl:with-param name="unit">
                  <xsl:text>twips</xsl:text>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:value-of select="$twips div 1440"/>
          </xsl:attribute>
        </xsl:if>

        <xsl:if test="@fo:margin-right">
          <xsl:attribute name="right">
            <!-- 1 inch = 1440 twips -->
            <xsl:variable name="twips">
              <xsl:call-template name="ConvertMeasure">
                <xsl:with-param name="length" select="@fo:margin-right"/>
                <xsl:with-param name="unit">
                  <xsl:text>twips</xsl:text>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:value-of select="$twips div 1440"/>
          </xsl:attribute>
        </xsl:if>

        <xsl:if test="@fo:margin-top">
          <xsl:attribute name="top">
            <!-- 1 inch = 1440 twips -->
            <xsl:variable name="twips">
              <xsl:call-template name="ConvertMeasure">
                <xsl:with-param name="length" select="@fo:margin-top"/>
                <xsl:with-param name="unit">
                  <xsl:text>twips</xsl:text>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:value-of select="$twips div 1440"/>
          </xsl:attribute>
        </xsl:if>

        <xsl:if test="@fo:margin-bottom">
          <xsl:attribute name="bottom">
            <!-- 1 inch = 1440 twips -->
            <xsl:variable name="twips">
              <xsl:call-template name="ConvertMeasure">
                <xsl:with-param name="length" select="@fo:margin-bottom"/>
                <xsl:with-param name="unit">
                  <xsl:text>twips</xsl:text>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:value-of select="$twips div 1440"/>
          </xsl:attribute>
        </xsl:if>
      </pageMargins>
    </xsl:if>
  </xsl:template>

  <xsl:template name="TranslatePaperSize">
    <xsl:param name="height"/>
    <xsl:param name="width"/>

    <xsl:choose>
      <!-- A3 -->
      <xsl:when test="($width = 42 and $height = 29.7) or ($width = 29.7 and $height = 42)">
        <xsl:text>8</xsl:text>
      </xsl:when>
      <!-- A4 -->
      <xsl:when test="($width = 29.7 and $height = 21) or ($width = 21 and $height = 29.7)">
        <xsl:text>9</xsl:text>
      </xsl:when>
      <!-- A5 -->
      <xsl:when test="($width = 21 and $height = 14.8) or ($width = 14.8 and $height = 21)">
        <xsl:text>11</xsl:text>
      </xsl:when>
      <!-- B4 (JIS) -->
      <xsl:when test="($width = 36.4 and $height = 25.7) or ($width = 25.7 and $height = 36.4)">
        <xsl:text>12</xsl:text>
      </xsl:when>
      <!-- B5 (JIS) -->
      <xsl:when test="($width = 25.7 and $height = 18.2) or ($width = 18.2 and $height = 25.7)">
        <xsl:text>13</xsl:text>
      </xsl:when>
      <!-- Letter -->
      <xsl:when test="($width = 27.94 and $height = 21.59) or ($width = 21.59 and $height = 27.94)">
        <xsl:text>1</xsl:text>
      </xsl:when>
      <!-- Tabloid -->
      <xsl:when test="($width = 43.13 and $height = 27.96) or ($width = 27.96 and $height = 43.13)">
        <xsl:text>3</xsl:text>
      </xsl:when>
      <!-- Legal -->
      <xsl:when test="($width = 35.57 and $height = 21.59) or ($width = 21.59 and $height = 35.57)">
        <xsl:text>5</xsl:text>
      </xsl:when>
      <!-- Japanese Postcard -->
      <xsl:when test="($width = 14.8 and $height = 10) or ($width = 10 and $height = 14.8)">
        <xsl:text>43</xsl:text>
      </xsl:when>
      <!-- A4 as default -->
      <xsl:otherwise>
        <xsl:text>9</xsl:text>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <xsl:template match="table:table-column" mode="DefaultBorder">
    <xsl:choose>
      <xsl:when
        test="key('style', @table:default-cell-style-name)/style:table-cell-properties/@fo:border or key('style', @table:default-cell-style-name)/style:table-cell-properties/@fo:border-top or
        key('style', @table:default-cell-style-name)/style:table-cell-properties/@fo:border-bottom or key('style', @table:default-cell-style-name)/style:table-cell-properties/@fo:border-left or
        key('style', @table:default-cell-style-name)/style:table-cell-properties/@fo:border-right">
        <xsl:text>true</xsl:text>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <!-- next column is sibling of this one -->
          <xsl:when test="following-sibling::node()[1][name() = 'table:table-column' ]">
            <xsl:apply-templates select="following-sibling::table:table-column[1]"
              mode="DefaultBorder"/>
          </xsl:when>
          <!-- this is the last column before header  -->
          <xsl:when test="following-sibling::node()[1][name() = 'table:table-header-columns' ]">
            <xsl:apply-templates select="following-sibling::node()[1]/table:table-column[1]"
              mode="DefaultBorder"/>
          </xsl:when>
          <!-- this is the last column inside header  -->
          <xsl:when
            test="not(following-sibling::node()[1][name() = 'table:table-column' ]) and parent::node()[name() = 'table:table-header-columns' ] and parent::node()/following-sibling::table:table-column[1]">
            <xsl:apply-templates select="parent::node()/following-sibling::table:table-column[1]"
              mode="DefaultBorder"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>false</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="CountRows">
    <xsl:param name="value" select="0"/>

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
      <xsl:when test="following-sibling::table:table-row[1]">
        <xsl:for-each select="following-sibling::table:table-row[1]">
          <xsl:call-template name="CountRows">
            <xsl:with-param name="value" select="number($value)+number($rows)"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$value + $rows"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="CountCols">
    <xsl:param name="value" select="0"/>

    <xsl:variable name="cols">
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
      <xsl:when test="following-sibling::table:table-column[1]">
        <xsl:for-each select="following-sibling::table:table-column[1]">
          <xsl:call-template name="CountCols">
            <xsl:with-param name="value" select="number($value)+number($cols)"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$value + $cols"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

</xsl:stylesheet>

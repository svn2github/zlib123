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
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" exclude-result-prefixes="table">

  <xsl:import href="measures.xsl"/>
  <xsl:import href="pixel-measure.xsl"/>
  <xsl:import href="page.xsl"/>
  <xsl:key name="StyleFamily" match="style:style" use="@style:family"/>
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
        <xsl:value-of
          select="$cellNumber + count(table:table-row/table:table-cell[text:p and (@office:value-type='string' or not((number(text:p) or text:p = 0)))])"
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
      <xsl:call-template name="InsertViewSettings">
        <xsl:with-param name="sheetId" select="$sheetId"/>
      </xsl:call-template>

      <xsl:call-template name="InsertSheetContent">
        <xsl:with-param name="sheetId" select="$sheetId"/>
        <xsl:with-param name="cellNumber" select="$cellNumber"/>
      </xsl:call-template>

      <!-- Insert Merge Cells -->
      <xsl:call-template name="CheckMergeCell"/>

      <xsl:call-template name="InsertPageProperties"/>
      <xsl:call-template name="InsertHeaderFooter"/>
    </worksheet>
  </xsl:template>

  <xsl:template name="InsertViewSettings">
    <xsl:param name="sheetId"/>

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
                <xsl:for-each select="document('settings.xml')">
                  <xsl:choose>
                    <xsl:when
                      test="office:document-settings/office:settings/config:config-item-set[@config:name = 'ooo:view-settings']/config:config-item-map-indexed[@config:name = 'Views']/config:config-item-map-entry/config:config-item-map-named[@config:name='Tables']/config:config-item-map-entry[position() = $sheetId]/config:config-item[@config:name='CursorPositionX']">
                      <xsl:value-of
                        select="office:document-settings/office:settings/config:config-item-set[@config:name = 'ooo:view-settings']/config:config-item-map-indexed[@config:name = 'Views']/config:config-item-map-entry/config:config-item-map-named[@config:name='Tables']/config:config-item-map-entry[position() = $sheetId]/config:config-item[@config:name='CursorPositionX']"
                      />
                    </xsl:when>
                    <xsl:otherwise>0</xsl:otherwise>
                  </xsl:choose>
                </xsl:for-each>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="row">
            <xsl:for-each select="document('settings.xml')">
              <xsl:choose>
                <xsl:when
                  test="office:document-settings/office:settings/config:config-item-set[@config:name = 'ooo:view-settings']/config:config-item-map-indexed[@config:name = 'Views']/config:config-item-map-entry/config:config-item-map-named[@config:name='Tables']/config:config-item-map-entry[position() = $sheetId]/config:config-item[@config:name='CursorPositionY']">
                  <xsl:value-of
                    select="office:document-settings/office:settings/config:config-item-set[@config:name = 'ooo:view-settings']/config:config-item-map-indexed[@config:name = 'Views']/config:config-item-map-entry/config:config-item-map-named[@config:name='Tables']/config:config-item-map-entry[position() = $sheetId]/config:config-item[@config:name='CursorPositionY']"
                  />
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:for-each>
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
      <xsl:call-template name="ConvertToCharacters">
        <xsl:with-param name="width">
          <xsl:value-of select="concat('0.8925','in')"/>
        </xsl:with-param>
        <xsl:with-param name="defaultFontSize" select="$defaultFontSize"/>
      </xsl:call-template>
    </xsl:variable>

    <sheetFormatPr defaultColWidth="{$defaultColWidth}" defaultRowHeight="{$defaultRowHeight}"
      customHeight="true"/>
    <xsl:if test="table:table-column">
      <cols>

        <!-- insert first column -->
        <xsl:apply-templates select="table:table-column[1]" mode="sheet">
          <xsl:with-param name="colNumber">1</xsl:with-param>
          <xsl:with-param name="defaultFontSize" select="$defaultFontSize"/>
        </xsl:apply-templates>

      </cols>
    </xsl:if>
    <sheetData>

      <xsl:variable name="ColumnTagNum">
        <xsl:apply-templates select="table:table-column[1]" mode="tag">
          <xsl:with-param name="colNumber">1</xsl:with-param>
          <xsl:with-param name="defaultFontSize" select="$defaultFontSize"/>
        </xsl:apply-templates>
      </xsl:variable>

      <!-- insert first row -->
      <xsl:apply-templates select="table:table-row[1]" mode="sheet">
        <xsl:with-param name="rowNumber">1</xsl:with-param>
        <xsl:with-param name="cellNumber" select="$cellNumber"/>
        <xsl:with-param name="defaultRowHeight" select="$defaultRowHeight"/>
        <xsl:with-param name="TableColumnTagNum">
          <xsl:value-of select="$ColumnTagNum"/>
        </xsl:with-param>
      </xsl:apply-templates>

    </sheetData>
  </xsl:template>

  <xsl:template name="InsertHeaderFooter">
    <xsl:if
      test="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name = 'Default']/style:header/child::node() or
    document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name = 'Default']/style:footer/child::node()">
      <headerFooter>
        <xsl:for-each
          select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name = 'Default']/style:header-left">
          <xsl:if test="not(@style:display = 'false' )">
            <xsl:attribute name="differentOddEven">
              <xsl:text>1</xsl:text>
            </xsl:attribute>
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

    <!-- default pageStyle name-->
    <xsl:variable name="pageStyle">
      <xsl:value-of
        select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name = 'Default' ]/@style:page-layout-name"
      />
    </xsl:variable>

    <xsl:for-each select="document('styles.xml')">
      <xsl:for-each select="key('pageStyle',$pageStyle)">
        <xsl:if test="style:page-layout-properties/@style:table-centering">
          <printOptions>
            <xsl:if
              test="style:page-layout-properties/@style:table-centering = 'horizontal' or style:page-layout-properties/@style:table-centering = 'both' ">
              <xsl:attribute name="horizontalCentered">
                <xsl:text>1</xsl:text>
              </xsl:attribute>
            </xsl:if>
            <xsl:if
              test="style:page-layout-properties/@style:table-centering = 'vertical' or style:page-layout-properties/@style:table-centering = 'both' ">
              <xsl:attribute name="verticalCentered">
                <xsl:text>1</xsl:text>
              </xsl:attribute>
            </xsl:if>
          </printOptions>
        </xsl:if>

        <xsl:call-template name="InsertMargins">
          <xsl:with-param name="pageStyle" select="$pageStyle"/>
        </xsl:call-template>

        <xsl:if test="style:page-layout-properties/@style:print-orientation">
          <pageSetup>
            <xsl:if test="style:page-layout-properties/@style:print-orientation">
              <xsl:attribute name="orientation">
                <xsl:value-of select="style:page-layout-properties/@style:print-orientation"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if
              test="style:page-layout-properties/@fo:page-width and style:page-layout-properties/@fo:page-height">
              <xsl:attribute name="paperSize">
                <xsl:call-template name="TranslatePaperSize">
                  <xsl:with-param name="width" select="style:page-layout-properties/@fo:page-width"/>
                  <xsl:with-param name="height"
                    select="style:page-layout-properties/@fo:page-height"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
          </pageSetup>
        </xsl:if>
      </xsl:for-each>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="InsertMargins">
    <xsl:param name="pageStyle"/>

    <xsl:if
      test="style:page-layout-properties/@fo:margin-top or style:page-layout-properties/@fo:margin-bottom or 
      style:page-layout-properties/@fo:margin-right or style:page-layout-properties/@fo:margin-left">

      <pageMargins left="0.78740157480314965" right="0.70866141732283472" top="0.74803149606299213"
        bottom="0.74803149606299213" header="0.31496062992125984" footer="0.31496062992125984">
        <xsl:if test="style:page-layout-properties/@fo:margin-left">
          <xsl:attribute name="left">
            <!-- 1 inch = 1440 twips -->
            <xsl:variable name="twips">
              <xsl:call-template name="ConvertMeasure">
                <xsl:with-param name="length" select="style:page-layout-properties/@fo:margin-left"/>
                <xsl:with-param name="unit">
                  <xsl:text>twips</xsl:text>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:value-of select="$twips div 1440"/>
          </xsl:attribute>
        </xsl:if>

        <xsl:if test="style:page-layout-properties/@fo:margin-right">
          <xsl:attribute name="right">
            <!-- 1 inch = 1440 twips -->
            <xsl:variable name="twips">
              <xsl:call-template name="ConvertMeasure">
                <xsl:with-param name="length" select="style:page-layout-properties/@fo:margin-right"/>
                <xsl:with-param name="unit">
                  <xsl:text>twips</xsl:text>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:value-of select="$twips div 1440"/>
          </xsl:attribute>
        </xsl:if>

        <xsl:if test="style:page-layout-properties/@fo:margin-top">
          <xsl:attribute name="top">
            <!-- 1 inch = 1440 twips -->
            <xsl:variable name="twips">
              <xsl:call-template name="ConvertMeasure">
                <xsl:with-param name="length" select="style:page-layout-properties/@fo:margin-top"/>
                <xsl:with-param name="unit">
                  <xsl:text>twips</xsl:text>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:value-of select="$twips div 1440"/>
          </xsl:attribute>
        </xsl:if>

        <xsl:if test="style:page-layout-properties/@fo:margin-bottom">
          <xsl:attribute name="bottom">
            <!-- 1 inch = 1440 twips -->
            <xsl:variable name="twips">
              <xsl:call-template name="ConvertMeasure">
                <xsl:with-param name="length"
                  select="style:page-layout-properties/@fo:margin-bottom"/>
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
      <xsl:when test="$width='42cm' and $height='29.699cm' ">
        <xsl:text>8</xsl:text>
      </xsl:when>
      <!-- A4 -->
      <xsl:when test="$width='29.699cm' and $height='20.999cm' ">
        <xsl:text>9</xsl:text>
      </xsl:when>
      <!-- A5 -->
      <xsl:when test="$width='20.999cm' and $height='14.799cm' ">
        <xsl:text>11</xsl:text>
      </xsl:when>
      <!-- B4 (JIS) -->
      <xsl:when test="$width='36.4cm' and $height='25.7cm' ">
        <xsl:text>12</xsl:text>
      </xsl:when>
      <!-- B5 (JIS) -->
      <xsl:when test="$width='25.7cm' and $height='18.2cm' ">
        <xsl:text>13</xsl:text>
      </xsl:when>
      <!-- Letter -->
      <xsl:when test="$width='27.94cm' and $height='21.59cm' ">
        <xsl:text>1</xsl:text>
      </xsl:when>
      <!-- Tabloid -->
      <xsl:when test="$width='43.127cm' and $height='27.958cm' ">
        <xsl:text>3</xsl:text>
      </xsl:when>
      <!-- Legal -->
      <xsl:when test="$width='35.565cm' and $height='21.59cm' ">
        <xsl:text>5</xsl:text>
      </xsl:when>
      <!-- A4 as default -->
      <xsl:otherwise>
        <xsl:text>9</xsl:text>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

</xsl:stylesheet>

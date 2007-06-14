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
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  exclude-result-prefixes="office style table config text">

  <xsl:import href="worksheets.xsl"/>
  <xsl:import href="cell.xsl"/>
  <xsl:import href="sharedStrings.xsl"/>
  <xsl:import href="common.xsl"/>

  <xsl:key name="ConfigItem" match="config:config-item" use="@config:name"/>
  <xsl:key name="style" match="style:style" use="@style:name"/>
  <xsl:key name="pageStyle" match="style:page-layout" use="@style:name"/>

  <!-- main workbook template-->
  <xsl:template name="InsertWorkbook">
    <xsl:apply-templates select="document('content.xml')/office:document-content"/>
  </xsl:template>

  <!-- workbook body template -->
  <xsl:template match="office:body">
    <workbook>
      <xsl:call-template name="WorkbookView"/>
      <xsl:apply-templates select="office:spreadsheet"/>
    </workbook>
  </xsl:template>

  <!-- workbook  view template-->
  <xsl:template name="WorkbookView">
    <bookViews>
      <workbookView>

        <!-- Insert firstSheet attribute when first sheet is hidden -->
        <xsl:if
          test="key('style',office:spreadsheet/table:table[position()=1]/@table:style-name)/style:table-properties/@table:display = 'false'">
          <xsl:attribute name="firstSheet">
            <xsl:variable name="TableStyleName">
              <xsl:value-of
                select="office:spreadsheet/table:table[key('style',@table:style-name)/style:table-properties/@table:display != 'false']/@table:style-name"
              />
            </xsl:variable>
            <xsl:for-each
              select="office:spreadsheet/table:table[@table:style-name=$TableStyleName][position()=1]">
              <xsl:value-of select="count(preceding-sibling::table:table)"/>
            </xsl:for-each>
          </xsl:attribute>
        </xsl:if>

        <!-- Insert activeTab (Active sheet after open the file) -->
        <xsl:attribute name="activeTab">
          <xsl:variable name="ActiveTable">
            <xsl:for-each select="document('settings.xml')">
              <xsl:value-of select="key('ConfigItem', 'ActiveTable')"/>
            </xsl:for-each>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test="office:spreadsheet/table:table[@table:name=$ActiveTable]">
              <xsl:for-each select="office:spreadsheet/table:table[@table:name=$ActiveTable]">
                <xsl:value-of select="count(preceding-sibling::table:table)"/>
              </xsl:for-each>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>

      </workbookView>
    </bookViews>
  </xsl:template>

  <!-- insert references to all sheets -->
  <xsl:template match="office:spreadsheet">
    <sheets>
      <xsl:for-each select="table:table">
        <sheet>
          <!-- characters "*\/[];'?" can not occur in sheet name -->
          <xsl:attribute name="name">
            <!-- if there is a shheet with the same name modify name -->
            <xsl:call-template name="CheckSheetName">
              <xsl:with-param name="sheetNumber">
                <xsl:value-of select="position()"/>
              </xsl:with-param>
              <xsl:with-param name="name">
                <xsl:value-of
                  select="substring(translate(@table:name,&quot;*\/[]:&apos;?&quot;,&quot;&quot;),1,31)"
                />
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
          <xsl:attribute name="sheetId">
            <xsl:value-of select="position()"/>
          </xsl:attribute>
          <xsl:attribute name="r:id">
            <xsl:value-of select="generate-id(.)"/>
          </xsl:attribute>
          <xsl:if
            test="key('style',@table:style-name)/style:table-properties/@table:display = 'false'">
            <xsl:attribute name="state">
              <xsl:text>hidden</xsl:text>
            </xsl:attribute>
          </xsl:if>
        </sheet>
      </xsl:for-each>
    </sheets>
    <definedNames>
      <xsl:call-template name="InsertPrintRanges"/>
      <xsl:call-template name="InsertHeaders"/>
    </definedNames>

  </xsl:template>

  <!-- insert all sheets -->
  <xsl:template name="InsertSheets">
    <!-- convert first table -->
    <xsl:apply-templates
      select="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table[1]"
      mode="sheet">
      <xsl:with-param name="cellNumber">0</xsl:with-param>
      <xsl:with-param name="sheetId">1</xsl:with-param>
    </xsl:apply-templates>
  </xsl:template>

  <xsl:template name="CheckSheetName">
    <xsl:param name="name"/>
    <xsl:param name="sheetNumber"/>

    <xsl:choose>
      <!-- when there are at least 2 sheets with the same name after removal of forbidden characters and cutting to 31 characters (name correction) -->
      <xsl:when
        test="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table[substring(translate(@table:name,&quot;*\/[]:&apos;?&quot;,&quot;&quot;),1,31) = $name][2]">
        <xsl:variable name="nameConflictsBefore">
          <xsl:for-each
            select="document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table[position() = $sheetNumber]">
            <!-- count sheets before this one whose name (after correction) collide with this sheet name (after correction) -->
            <xsl:value-of
              select="count(document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table[substring(translate(@table:name,&quot;*\/[]:&apos;?&quot;,&quot;&quot;),1,31) = $name and position() &lt; $sheetNumber])"
            />
          </xsl:for-each>
        </xsl:variable>
        <!-- cut name and add "(N)" at the end where N is seqential number of duplicated name -->
        <xsl:value-of
          select="concat(substring(translate(@table:name,&quot;*\/[]:&apos;?&quot;,&quot;&quot;),1,31 - 2 - string-length($nameConflictsBefore + 1)),'_',$nameConflictsBefore + 1)"
        />
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$name"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertPrintRanges">
    <xsl:for-each select="table:table">
      <xsl:if test="@table:print-ranges">
        <definedName name="_xlnm.Print_Area">
          <xsl:attribute name="localSheetId">
            <xsl:value-of select="position() - 1"/>
          </xsl:attribute>
          <xsl:call-template name="InsertRange">
            <xsl:with-param name="range" select="@table:print-ranges"/>
          </xsl:call-template>
        </definedName>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="InsertRange">
    <xsl:param name="range"/>

    <xsl:variable name="row1">
      <xsl:call-template name="GetRowNum">
        <xsl:with-param name="cell"
          select="concat(substring-after($range,'.'),substring-before($range,':'))"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="col1">
      <xsl:value-of select="substring-before(substring-after($range,'.'),$row1)"/>
    </xsl:variable>

    <xsl:variable name="row2">
      <xsl:variable name="endCell">
        <xsl:choose>
          <!-- when there is next range there is a space before -->
          <xsl:when test="contains(substring-after(substring-after($range,'.'),'.'),' ')">
            <xsl:value-of
              select="substring-before(substring-after(substring-after($range,'.'),'.'),' ')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="substring-after(substring-after($range,'.'),'.')"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>

      <xsl:call-template name="GetRowNum">
        <xsl:with-param name="cell" select="$endCell"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="col2">
      <xsl:value-of
        select="substring-before(substring-after(substring-after($range,'.'),'.'),$row2)"/>
    </xsl:variable>

    <!-- if sheet name contains space then name has to be inside apostrophes -->
    <xsl:variable name="sheetName">
      <xsl:text>'</xsl:text>
      <xsl:value-of select="@table:name"/>
      <xsl:text>'</xsl:text>
    </xsl:variable>

    <xsl:value-of
      select="concat($sheetName,'!$',substring-before(substring-after($range,'.'),$row1),'$', $row1,':','$',$col2,'$', $row2)"/>

    <!-- if there is next range there is a space before -->
    <xsl:if test="contains(substring-after(substring-after($range,'.'),'.'),' ')">
      <xsl:text>,</xsl:text>
      <xsl:call-template name="InsertRange">
        <xsl:with-param name="range">
          <xsl:value-of
            select="substring-after(substring-after(substring-after($range,'.'),'.'),' ')"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertHeaders">

    <xsl:for-each select="table:table">
      <xsl:if test="table:table-header-rows or table:table-header-columns">
        <definedName name="_xlnm.Print_Titles">
          <xsl:attribute name="localSheetId">
            <xsl:value-of select="position() - 1"/>
          </xsl:attribute>

          <xsl:choose>
            <!-- when there are only header columns -->
            <xsl:when test="table:table-header-columns and not(table:table-header-rows)">
              <xsl:call-template name="InsertHeaderCols"/>
            </xsl:when>
            <!-- when there are only header rows -->
            <xsl:when test="table:table-header-rows and not(table:table-header-columns)">
              <xsl:call-template name="InsertHeaderRows"/>
            </xsl:when>
            <!-- when there are both: header rows and header columns -->
            <xsl:when test="table:table-header-rows and table:table-header-columns">
              <xsl:call-template name="InsertHeaderCols"/>
              <xsl:text>,</xsl:text>
              <xsl:call-template name="InsertHeaderRows"/>
            </xsl:when>
          </xsl:choose>

        </definedName>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="InsertHeaderRows">

    <xsl:variable name="headerRowStart">
      <xsl:variable name="count">
        <xsl:for-each select="descendant::table:table-row[1]">
          <xsl:call-template name="CountHeaderRowsStart"/>
        </xsl:for-each>
      </xsl:variable>
      
      <xsl:choose>
        <xsl:when test="$count = '' ">
          <xsl:text>1</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$count"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="headerRows">
      <xsl:for-each select="table:table-header-rows/table:table-row[1]">
        <xsl:call-template name="CountRows"/>
      </xsl:for-each>
    </xsl:variable>

    <xsl:variable name="sheetName">
      <xsl:choose>
        <xsl:when test="contains(@table:name,' ') or contains(@table:name,'!') or contains(@table:name,'$')">
          <xsl:text>'</xsl:text>
          <xsl:value-of select="@table:name"/>
          <xsl:text>'</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="@table:name"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:value-of
      select="concat($sheetName,'!$',$headerRowStart,':$',$headerRowStart + $headerRows  - 1)"/>

  </xsl:template>

  <xsl:template name="InsertHeaderCols">

    <xsl:variable name="headerColStart">
      <xsl:for-each select="child::node()[name() != 'office:forms' ][1]">
        <xsl:call-template name="CountHeaderColsStart"/>
      </xsl:for-each>
    </xsl:variable>

    <xsl:variable name="headerCols">
      <xsl:for-each select="table:table-header-columns/table:table-column[1]">
        <xsl:call-template name="CountCols"/>
      </xsl:for-each>
    </xsl:variable>

    <xsl:variable name="charHeaderColStart">
      <xsl:call-template name="NumbersToChars">
        <xsl:with-param name="num" select="$headerColStart - 1"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="charHeaderColEnd">
      <xsl:call-template name="NumbersToChars">
        <xsl:with-param name="num" select="$headerColStart + $headerCols  - 2"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="sheetName">
      <xsl:choose>
        <xsl:when test="contains(@table:name,' ') or contains(@table:name,'!') or contains(@table:name,'$')">
          <xsl:text>'</xsl:text>
          <xsl:value-of select="@table:name"/>
          <xsl:text>'</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="@table:name"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:value-of select="concat($sheetName,'!$',$charHeaderColStart,':$',$charHeaderColEnd)"/>

  </xsl:template>

  <xsl:template name="CountHeaderRowsStart">
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
      <xsl:when test="name() = 'table:table-row' ">
        <xsl:for-each select="following-sibling::node()[1]">
          <xsl:call-template name="CountHeaderRowsStart">
            <xsl:with-param name="value" select="$value + $rows"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$value + 1"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="CountHeaderColsStart">
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
      <xsl:when test="name() = 'table:table-column' ">
        <xsl:for-each select="following-sibling::node()[1]">
          <xsl:call-template name="CountHeaderColsStart">
            <xsl:with-param name="value" select="$value + $cols"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$value + 1"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

</xsl:stylesheet>

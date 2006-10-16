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
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  exclude-result-prefixes="office text table fo style draw">


  <!-- Specifies that measurement of table properties are interpreted as twentieths of a point -->
  <xsl:variable name="type">dxa</xsl:variable>

  <!-- Tables -->
  <xsl:template match="table:table">
    <w:tbl>
      <xsl:call-template name="MarkMasterPage"/>
      <w:tblPr>
        <xsl:choose>
          <xsl:when test="@table:is-sub-table='true'">
            <xsl:call-template name="InsertSubTableProperties"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="InsertTableProperties"/>
          </xsl:otherwise>
        </xsl:choose>
      </w:tblPr>
      <xsl:call-template name="InsertTblGrid"/>
      <!-- Header rows not handled the same way in OOX -->
      <xsl:if test="table:table-header-rows/table:table-row">
        <xsl:message terminate="no">feedback:Table header repeated on every page</xsl:message>
      </xsl:if>
      <xsl:if test="count(table:table-header-rows/table:table-row) &gt; 1">
        <xsl:message terminate="no">feedback:Table header rows</xsl:message>
      </xsl:if>
      <xsl:apply-templates select="table:table-header-rows/table:table-row | table:table-row"/>
    </w:tbl>
    <xsl:call-template name="ManageSectionsInTable"/>
  </xsl:template>

  <!-- Inserts table properties -->
  <xsl:template name="InsertTableProperties">
    <w:tblStyle w:val="{@table:style-name}"/>
    <xsl:variable name="tableProp"
      select="key('automatic-styles', @table:style-name)/style:table-properties"/>

    <!-- report lost attributes -->
    <xsl:if test="$tableProp/@fo:keep-with-next">
      <xsl:message terminate="no">feedback:Table together with next paragraph</xsl:message>
    </xsl:if>
    <xsl:if test="not($tableProp/@style:may-break-between-rows='true')">
      <xsl:message terminate="no">feedback:Unsplitable table</xsl:message>
    </xsl:if>
    <xsl:if test="$tableProp/@style:background-image">
      <xsl:message terminate="no">feedback:Table background image</xsl:message>
    </xsl:if>
    <xsl:if test="$tableProp/@style:shadow">
      <xsl:message terminate="no">feedback:Table shadow</xsl:message>
    </xsl:if>
    
    <w:tblW w:type="{$type}">
      <xsl:attribute name="w:w">
        <xsl:call-template name="twips-measure">
          <xsl:with-param name="length" select="$tableProp/@style:width"/>
        </xsl:call-template>
      </xsl:attribute>
    </w:tblW>
    <xsl:if test="$tableProp/@table:align">
      <xsl:choose>
        <xsl:when test="$tableProp/@table:align = 'margins'">
          <!--User agents that do not support the "margins" value, may treat this value as "left".-->
          <xsl:message terminate="no">feedback:Manual alignment of table</xsl:message>
          <w:jc w:val="left"/>
        </xsl:when>
        <xsl:otherwise>
          <w:jc w:val="{$tableProp/@table:align}"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <xsl:call-template name="InsertTableIndent"/>

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

    <!-- Default layout algorithm in ODF is "fixed". -->
    <w:tblLayout w:type="fixed"/>
  </xsl:template>

  <!-- Insert the table indent if margin-left defined, or if cell-padding greater than 0. -->
  <xsl:template name="InsertTableIndent">
    <xsl:variable name="marginLeft">
      <xsl:choose>
        <xsl:when
          test="key('automatic-styles', @table:style-name)/style:table-properties/@fo:margin-left != '' ">
          <xsl:call-template name="twips-measure">
            <xsl:with-param name="length"
              select="key('automatic-styles', @table:style-name)/style:table-properties/@fo:margin-left"
            />
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <!-- get left padding from cells (ignore subcells padding) -->
    <xsl:variable name="padding">
      <xsl:choose>
        <xsl:when
          test="key('automatic-styles', descendant::table:table-cell[not(@table:is-sub-table = 'true')]/@table:style-name)/style:table-cell-properties[@fo:padding and @fo:padding != 'none']">
          <xsl:call-template name="twips-measure">
            <xsl:with-param name="length">
              <xsl:value-of
                select="key('automatic-styles', descendant::table:table-cell[not(@table:is-sub-table = 'true')]/@table:style-name)/style:table-cell-properties/@fo:padding"
              />
            </xsl:with-param>
          </xsl:call-template>
        </xsl:when>
        <xsl:when
          test="key('automatic-styles', descendant::table:table-cell[not(@table:is-sub-table = 'true')]/@table:style-name)/style:table-cell-properties[@fo:padding-left and @fo:padding-left != 'none']">
          <xsl:call-template name="twips-measure">
            <xsl:with-param name="length">
              <xsl:value-of
                select="key('automatic-styles', descendant::table:table-cell[not(@table:is-sub-table = 'true')]/@table:style-name)/style:table-cell-properties/@fo:padding-left"
              />
            </xsl:with-param>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <w:tblInd w:type="{$type}">
      <xsl:attribute name="w:w">
        <xsl:value-of select="$marginLeft + $padding"/>
      </xsl:attribute>
    </w:tblInd>
  </xsl:template>

  <!-- In case the table is a subtable. Unherits a few properties of table it belongs to. -->
  <xsl:template name="InsertSubTableProperties">
    <xsl:variable name="tableStyleName">
      <xsl:value-of
        select="ancestor::table:table[1][not(@table:is-sub-table='true')]/@table:style-name"/>
    </xsl:variable>
    <w:tblStyle w:val="{$tableStyleName}"/>
    <xsl:variable name="tableProp"
      select="key('automatic-styles', $tableStyleName)/style:table-properties"/>

    <!-- report lost attributes -->
    <xsl:if test="$tableProp/@fo:keep-with-next">
      <xsl:message terminate="no">feedback:Table together with next paragraph</xsl:message>
    </xsl:if>
    <xsl:if test="not($tableProp/@style:may-break-between-rows='true')">
      <xsl:message terminate="no">feedback:Unsplitable table</xsl:message>
    </xsl:if>
    <xsl:if test="$tableProp/@style:background-image">
      <xsl:message terminate="no">feedback:Table background image</xsl:message>
    </xsl:if>
    <xsl:if test="$tableProp/@style:shadow">
      <xsl:message terminate="no">feedback:Table shadow</xsl:message>
    </xsl:if>
    
    <w:tblW w:type="{$type}">
      <xsl:attribute name="w:w">
        <xsl:call-template name="twips-measure">
          <xsl:with-param name="length">
            <xsl:for-each select="ancestor::table:table-cell[1]">
              <xsl:call-template name="GetCellWidth"/>
            </xsl:for-each>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </w:tblW>

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

    <!-- Default layout algorithm in ODF is "fixed". -->
    <w:tblLayout w:type="fixed"/>
  </xsl:template>

  <!-- Inserts a table grid -->
  <xsl:template name="InsertTblGrid">
    <w:tblGrid>
      <xsl:apply-templates select="table:table-column" mode="gridCol"/>
    </w:tblGrid>
  </xsl:template>

  <!-- match columns for gridCols -->
  <xsl:template match="table:table-column" mode="gridCol">
    <xsl:call-template name="InsertGridCol">
      <xsl:with-param name="width"
        select="key('automatic-styles',@table:style-name)/style:table-column-properties/@style:column-width"/>
      <xsl:with-param name="number" select="@table:number-columns-repeated"/>
    </xsl:call-template>
  </xsl:template>

  <!-- Inserts a gridCol -->
  <xsl:template name="InsertGridCol">
    <xsl:param name="width"/>
    <xsl:param name="number"/>
    <w:gridCol>
      <xsl:attribute name="w:w">
        <xsl:call-template name="twips-measure">
          <xsl:with-param name="length" select="$width"/>
        </xsl:call-template>
      </xsl:attribute>
    </w:gridCol>
    <xsl:if test="$number > 1">
      <xsl:call-template name="InsertGridCol">
        <xsl:with-param name="width" select="$width"/>
        <xsl:with-param name="number" select="$number - 1"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>



  <!-- table rows -->
  <xsl:template match="table:table-row">
    <w:tr>
      <xsl:if
        test="key('automatic-styles',child::table:table-cell/@table:style-name)/style:table-cell-properties/@fo:wrap-option='no-wrap'">
        <!-- Override layout algorithm -->
        <w:tblPrEx>
          <w:tblLayout w:type="auto"/>
        </w:tblPrEx>
      </xsl:if>
      <w:trPr>
        <xsl:call-template name="InsertRowProperties"/>
      </w:trPr>
      <xsl:apply-templates select="*[position() &lt; 64]"/>
      <xsl:if test="*[position() &gt;= 64]">
        <xsl:message terminate="no">feedback: table with more than 64 columns</xsl:message>
      </xsl:if>
    </w:tr>
  </xsl:template>

  <!-- Inserts row properties -->
  <xsl:template name="InsertRowProperties">
    <!-- report lost attributes -->
    <xsl:if
      test="key('automatic-styles',@table:style-name)/style:table-row-properties/@style:background-image">
      <xsl:message terminate="no">feedback:Row background image</xsl:message>
    </xsl:if>

    <xsl:call-template name="InsertRowHeaderMark"/>
    <xsl:call-template name="InsertRowHeight"/>
    <xsl:call-template name="InsertRowKeepTogether"/>
  </xsl:template>

  <!-- Inserts row header mark -->
  <xsl:template name="InsertRowHeaderMark">
    <xsl:if test="parent::table:table-header-rows">
      <w:tblHeader/>
    </xsl:if>
  </xsl:template>

  <!-- Inserts row height -->
  <xsl:template name="InsertRowHeight">
    <xsl:if test="@table:style-name">
      <xsl:variable name="rowProp"
        select="key('automatic-styles',@table:style-name)/style:table-row-properties"/>
      <xsl:variable name="widthRow">
        <xsl:choose>
          <xsl:when test="$rowProp[@style:row-height]">
            <xsl:value-of select="$rowProp/@style:row-height"/>
          </xsl:when>
          <xsl:when test="$rowProp[@style:min-row-height]">
            <xsl:value-of select="$rowProp/@style:min-row-height"/>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
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
            <xsl:when test="$rowProp[@style:row-height]">
              <xsl:value-of select="'exact'"/>
            </xsl:when>
            <xsl:when test="$rowProp[@style:min-row-height]">
              <xsl:value-of select="'atLeast'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'auto'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </w:trHeight>
    </xsl:if>
  </xsl:template>

  <!--Inserts keep together-->
  <xsl:template name="InsertRowKeepTogether">
    <xsl:if
      test="key('automatic-styles', @table:style-name)/style:table-row-properties/@style:keep-together = 'false'
    or key('automatic-styles', ancestor::table:table[@table:style-name][1]/@table:style-name)/style:table-properties/@style:may-break-between-rows='false'">
      <w:cantSplit/>
    </xsl:if>
  </xsl:template>

  <!-- table cells -->
  <xsl:template match="table:table-cell">
    <w:tc>
      <w:tcPr>
        <xsl:call-template name="InsertCellProperties"/>
      </w:tcPr>
      <xsl:apply-templates/>
      <xsl:if test="child::table:table/@table:is-sub-table = 'true'">
        <xsl:call-template name="InsertEmptyParagraph"/>
      </xsl:if>
    </w:tc>
  </xsl:template>

  <xsl:template name="InsertCellProperties">
    <!-- pointers on the cell style properties -->
    <xsl:variable name="cellProp"
      select="key('automatic-styles', @table:style-name)/style:table-cell-properties"/>
    <xsl:variable name="rowProp"
      select="key('automatic-styles', ../@table:style-name)/style:table-row-properties"/>
    <xsl:variable name="tableProp"
      select="key('automatic-styles', ancestor::table:table[@table:style-name][1]/@table:style-name)/style:table-properties"/>

    <!-- report lost attributes -->
    <xsl:if test="$cellProp/@style:cell-protect">
      <xsl:message terminate="no">feedback:Protected cell</xsl:message>
    </xsl:if>
    <xsl:if test="$cellProp/@style:background-image">
      <xsl:message terminate="no">feedback:Cell background image</xsl:message>
    </xsl:if>
    <xsl:if test="$tableProp/@style:shadow">
      <xsl:message terminate="no">feedback:Cell shadow</xsl:message>
    </xsl:if>
    
    <xsl:call-template name="InsertCellWidth"/>
    <xsl:call-template name="InsertCellSpan">
      <xsl:with-param name="vmerge"/>
    </xsl:call-template>
    <xsl:call-template name="InsertCellBorders">
      <xsl:with-param name="cellProp" select="$cellProp"/>
    </xsl:call-template>
    <xsl:call-template name="InsertCellBgColor">
      <xsl:with-param name="cellProp" select="$cellProp"/>
      <xsl:with-param name="rowProp" select="$rowProp"/>
      <xsl:with-param name="tableProp" select="$tableProp"/>
    </xsl:call-template>
    <xsl:call-template name="InsertCellMargins">
      <xsl:with-param name="cellProp" select="$cellProp"/>
    </xsl:call-template>
    <xsl:call-template name="InsertCellWritingMode">
      <xsl:with-param name="cellProp" select="$cellProp"/>
    </xsl:call-template>
    <xsl:call-template name="InsertCellValign">
      <xsl:with-param name="cellProp" select="$cellProp"/>
    </xsl:call-template>
  </xsl:template>

  <!-- Inserts the cell width -->
  <xsl:template name="InsertCellWidth">
    <w:tcW w:type="{$type}">
      <xsl:attribute name="w:w">
        <xsl:call-template name="GetCellWidth"/>
      </xsl:attribute>
    </w:tcW>
  </xsl:template>

  <!-- Inserts the cell span -->
  <xsl:template name="InsertCellSpan">
    <xsl:param name="vmerge"/>
    <!--xsl:choose>
      <xsl:when test="$vmerge != ''">
        <w:gridSpan w:val="{$grid}"/>
        <w:vMerge w:val="{$vmerge}"/>
      </xsl:when>
      <xsl:otherwise-->
    <xsl:if test="@table:number-columns-spanned">
      <w:gridSpan w:val="{@table:number-columns-spanned}"/>
    </xsl:if>
    <!--/xsl:otherwise>
    </xsl:choose-->
  </xsl:template>

  <!-- Inserts the cell borders -->
  <xsl:template name="InsertCellBorders">
    <xsl:param name="cellProp"/>
    <xsl:variable name="colsNumber" select="count(parent::table:table-row/*)"/>
    <w:tcBorders>
      <xsl:choose>
        <!-- if subCell, do not put border on the outer cells of the subtable it belongs -->
        <xsl:when test="ancestor::table:table[@table:is-sub-table='true']">
          <xsl:call-template name="GetSubCellBorders">
            <xsl:with-param name="colsNumber" select="$colsNumber"/>
            <xsl:with-param name="cellProp" select="$cellProp"/>
          </xsl:call-template>
        </xsl:when>
        <!-- if subtable, get the border from subcells -->
        <xsl:when test="table:table[@table:is-sub-table='true']">
          <xsl:call-template name="GetBordersFromSubCells">
            <xsl:with-param name="colsNumber" select="$colsNumber"/>
          </xsl:call-template>
        </xsl:when>
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
  </xsl:template>

  <!-- Write borders attribute by extracting them from subcells -->
  <xsl:template name="GetBordersFromSubCells">
    <xsl:param name="colsNumber"/>

    <xsl:variable name="SubCellsStyleName"
      select="descendant::node()[self::table:table-cell[parent::table:table-row[1 or last()]] or self::table:table-cell[1 or last()]]/@table:style-name"/>
    <xsl:variable name="subCellsProps"
      select="key('automatic-styles', $SubCellsStyleName)/style:table-cell-properties"/>

    <xsl:choose>
      <xsl:when test="$subCellsProps[@fo:border and @fo:border!='none' ]">
        <xsl:variable name="border">
          <!-- NB : value-of takes the first subCell properties only (not the whole node set) -->
          <xsl:value-of select="$subCellsProps[@fo:border and @fo:border != 'none']/@fo:border"/>
        </xsl:variable>
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
        <xsl:variable name="topCellsProps"
          select="key('automatic-styles', descendant::table:table-row[1]/table:table-cell/@table:style-name)/style:table-cell-properties"/>
        <xsl:variable name="rightCellsProps"
          select="key('automatic-styles', descendant::table:table-row/table:table-cell[last()]/@table:style-name)/style:table-cell-properties"/>
        <xsl:variable name="leftCellsProps"
          select="key('automatic-styles', descendant::table:table-row/table:table-cell[1]/@table:style-name)/style:table-cell-properties"/>
        <xsl:variable name="bottomCellsProps"
          select="key('automatic-styles', descendant::table:table-row[last()]/table:table-cell/@table:style-name)/style:table-cell-properties"/>

        <xsl:if test="$topCellsProps[@fo:border-top and @fo:border-top != 'none']">
          <xsl:variable name="border">
            <!-- NB : value-of takes the first subCell properties only (not the whole node set) -->
            <xsl:value-of
              select="$topCellsProps[@fo:border-top and @fo:border-top != 'none']/@fo:border-top"/>
          </xsl:variable>
          <w:top w:val="single" w:color="{substring-after($border, '#')}">
            <xsl:attribute name="w:sz">
              <xsl:call-template name="eightspoint-measure">
                <xsl:with-param name="length" select="substring-before($border, ' ')"/>
              </xsl:call-template>
            </xsl:attribute>
          </w:top>
        </xsl:if>
        <xsl:if test="$leftCellsProps[@fo:border-left and @fo:border-left != 'none']">
          <xsl:variable name="border">
            <!-- NB : value-of takes the first subCell properties only (not the whole node set) -->
            <xsl:value-of
              select="$leftCellsProps[@fo:border-left and @fo:border-left != 'none']/@fo:border-left"
            />
          </xsl:variable>
          <w:left w:val="single" w:color="{substring-after($border, '#')}">
            <xsl:attribute name="w:sz">
              <xsl:call-template name="eightspoint-measure">
                <xsl:with-param name="length" select="substring-before($border, ' ')"/>
              </xsl:call-template>
            </xsl:attribute>
          </w:left>
        </xsl:if>
        <xsl:if test="$bottomCellsProps[@fo:border-bottom and @fo:border-bottom != 'none']">
          <xsl:variable name="border">
            <!-- NB : value-of takes the first subCell properties only (not the whole node set) -->
            <xsl:value-of
              select="$leftCellsProps[@fo:border-bottom and @fo:border-bottom != 'none']/@fo:border-bottom"
            />
          </xsl:variable>
          <w:bottom w:val="single" w:color="{substring-after($border, '#')}">
            <xsl:attribute name="w:sz">
              <xsl:call-template name="eightspoint-measure">
                <xsl:with-param name="length" select="substring-before($border, ' ')"/>
              </xsl:call-template>
            </xsl:attribute>
          </w:bottom>
        </xsl:if>
        <xsl:if
          test="$rightCellsProps[(@fo:border-right and @fo:border-right != 'none')] or (position() &lt; $colsNumber and position() = 63)">
          <xsl:variable name="border">
            <xsl:choose>
              <xsl:when test="position() &lt; $colsNumber and position() = 63">
                <xsl:value-of
                  select="$subCellsProps[(@fo:border-left and @fo:border-left != 'none')]/@fo:border-left"
                />
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of
                  select="$rightCellsProps[(@fo:border-right and @fo:border-right != 'none')]/@fo:border-right"
                />
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
  </xsl:template>

  <!-- Write border attribute except if cell is on the border of the subtable -->
  <xsl:template name="GetSubCellBorders">
    <xsl:param name="cellProp"/>
    <xsl:param name="colsNumber"/>

    <xsl:choose>
      <!-- if subtable, get the border from subcells -->
      <xsl:when test="table:table[@table:is-sub-table='true']">
        <xsl:call-template name="GetBordersFromSubCells">
          <xsl:with-param name="colsNumber" select="$colsNumber"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$cellProp[@fo:border and @fo:border!='none' ]">
        <xsl:variable name="border" select="$cellProp/@fo:border"/>
        <!-- fo:border = "0.002cm solid #000000" -->
        <xsl:variable name="border-color" select="substring-after($border, '#')"/>
        <xsl:variable name="border-size">
          <xsl:call-template name="eightspoint-measure">
            <xsl:with-param name="length" select="substring-before($border,' ')"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:if test="parent::table:table-row/preceding-sibling::table:table-row">
          <w:top w:val="single" w:color="{$border-color}" w:sz="{$border-size}"/>
        </xsl:if>
        <xsl:if test="not(position()=1)">
          <w:left w:val="single" w:color="{$border-color}" w:sz="{$border-size}"/>
        </xsl:if>
        <xsl:if test="parent::table:table-row/following-sibling::table:table-row">
          <w:bottom w:val="single" w:color="{$border-color}" w:sz="{$border-size}"/>
        </xsl:if>
        <xsl:if test="not(position()=last())">
          <w:right w:val="single" w:color="{$border-color}" w:sz="{$border-size}"/>
        </xsl:if>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if
          test="$cellProp[@fo:border-top and @fo:border-top != 'none'] and parent::table:table-row/preceding-sibling::table:table-row">
          <xsl:variable name="border" select="$cellProp/@fo:border-top"/>
          <w:top w:val="single" w:color="{substring-after($border, '#')}">
            <xsl:attribute name="w:sz">
              <xsl:call-template name="eightspoint-measure">
                <xsl:with-param name="length" select="substring-before($border, ' ')"/>
              </xsl:call-template>
            </xsl:attribute>
          </w:top>
        </xsl:if>
        <xsl:if
          test="$cellProp[@fo:border-left and @fo:border-left != 'none'] and not(position()=1)">
          <xsl:variable name="border" select="$cellProp/@fo:border-left"/>
          <w:left w:val="single" w:color="{substring-after($border, '#')}">
            <xsl:attribute name="w:sz">
              <xsl:call-template name="eightspoint-measure">
                <xsl:with-param name="length" select="substring-before($border, ' ')"/>
              </xsl:call-template>
            </xsl:attribute>
          </w:left>
        </xsl:if>
        <xsl:if
          test="$cellProp[@fo:border-bottom and @fo:border-bottom != 'none'] and parent::table:table-row/following-sibling::table:table-row">
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
          test="($cellProp[(@fo:border-right and @fo:border-right != 'none')] or (position() &lt; $colsNumber and position() = 63)) and not(position()=last())">
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
  </xsl:template>

  <!-- Inserts the cell boackground color -->
  <xsl:template name="InsertCellBgColor">
    <xsl:param name="cellProp"/>
    <xsl:param name="rowProp"/>
    <xsl:param name="tableProp"/>
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
            test="$tableProp/@fo:background-color and $tableProp/@fo:background-color != 'transparent' ">
            <w:shd w:val="clear" w:color="auto"
              w:fill="{substring($tableProp/@fo:background-color, 2, string-length($tableProp/@fo:background-color) -1)}"
            />
          </xsl:when>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Inserts the cell margins -->
  <xsl:template name="InsertCellMargins">
    <xsl:param name="cellProp"/>
    <xsl:choose>
      <xsl:when test="table:table[@table:is-sub-table='true']">
        <w:tcMar>
          <w:top w:w="0" w:type="{$type}"/>
          <w:left w:w="0" w:type="{$type}"/>
          <w:bottom w:w="0" w:type="{$type}"/>
          <w:right w:w="0" w:type="{$type}"/>
        </w:tcMar>
      </xsl:when>
      <xsl:otherwise>
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
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Inserts the cell writing mode -->
  <xsl:template name="InsertCellWritingMode">
    <xsl:param name="cellProp"/>
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
  </xsl:template>

  <!-- Inserts the cell vertical alignment -->
  <xsl:template name="InsertCellValign">
    <xsl:param name="cellProp"/>
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
  </xsl:template>

  <!-- Inserts an empty paragraph -->
  <xsl:template name="InsertEmptyParagraph">
    <w:p>
      <w:pPr>
        <w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="text" w:y="1"/>
        <w:suppressOverlap/>
      </w:pPr>
    </w:p>
  </xsl:template>

  <!-- Gets the width of the current cell -->
  <xsl:template name="GetCellWidth">
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
          <xsl:call-template name="twips-measure">
            <xsl:with-param name="length"
              select="key('automatic-styles',ancestor::table:table[1]/table:table-column[position() = $currentPosInColumns]/@table:style-name)/style:table-column-properties/@style:column-width"
            />
          </xsl:call-template>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="($colNumber + $colSpan) &lt;= ($currentColNumber + $rangeColNumber)">
            <xsl:value-of select="$currentColumnWidth * $colSpan"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:variable name="remainingCellWidth">
              <xsl:call-template name="GetCellWidth">
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
    <xsl:value-of
      select="count(preceding-sibling::table:table-cell) + count(preceding-sibling::table:covered-table-cell) + 1"
    />
  </xsl:template>


  <!-- Inserts a page break before if needed -->
  <xsl:template name="InsertPageBreakBefore">
    <!-- in the first paragraph of a table -->
    <xsl:if test="not(preceding-sibling::text:p) and ancestor-or-self::table:table">
      <!-- if first cell -->
      <xsl:if test="parent::table:table-cell[not(preceding-sibling::node())]">
        <!-- if first row -->
        <xsl:if
          test="ancestor::table:table-row[preceding-sibling::node()[1][name()='table:table-column' or name()='table:table-columns'] or not(preceding-sibling::node())]">
          <!-- if associated style has a pgBreakBefore property -->
          <xsl:if
            test="key('automatic-styles', ancestor::table:table[1]/@table:style-name)/style:table-properties/@fo:break-before = 'page'
        or key('automatic-styles', ancestor::table:table[1]/@table:style-name)/@style:master-page-name != ''">
            <w:pageBreakBefore/>
          </xsl:if>
        </xsl:if>
      </xsl:if>
    </xsl:if>
  </xsl:template>


</xsl:stylesheet>

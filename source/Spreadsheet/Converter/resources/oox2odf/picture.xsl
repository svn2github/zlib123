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
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  exclude-result-prefixes="e r c xdr draw xlink">

  <xsl:import href="relationships.xsl"/>
  <xsl:import href="chart.xsl"/>
  
  <xsl:key name="drawing" match="e:drawing" use="''"/>

  <!-- We check cell when the picture is starting and ending -->
  <xsl:template name="PictureCell">
    <xsl:param name="sheet"/>
    <xsl:apply-templates select="e:worksheet/e:drawing">
      <xsl:with-param name="sheet">
        <xsl:value-of select="$sheet"/>
      </xsl:with-param>
    </xsl:apply-templates>
  </xsl:template>

  <!-- Get Row with Picture -->
  <xsl:template name="PictureRow">
    <xsl:param name="PictureCell"/>
    <xsl:param name="Result"/>
    <xsl:choose>
      <xsl:when test="$PictureCell != ''">
        <xsl:call-template name="PictureRow">
          <xsl:with-param name="PictureCell">
            <xsl:value-of select="substring-after($PictureCell, ';')"/>
          </xsl:with-param>
          <xsl:with-param name="Result">
            <xsl:value-of
              select="concat($Result,  concat(substring-before($PictureCell, ':'), ';'))"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$Result"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="e:drawing">
    <xsl:param name="sheet"/>

    <xsl:variable name="Target">
      <xsl:call-template name="GetTargetPicture">
        <xsl:with-param name="sheet">
          <xsl:value-of select="$sheet"/>
        </xsl:with-param>
        <xsl:with-param name="id">
          <xsl:value-of select="@r:id"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:for-each select="document(concat('xl/', substring-after($Target, '/')))">
      <xsl:if test="xdr:wsDr/xdr:twoCellAnchor">
        <xsl:apply-templates select="xdr:wsDr/xdr:twoCellAnchor[1]"/>
      </xsl:if>
    </xsl:for-each>

  </xsl:template>

  <!-- We check drawing's file -->
  <xsl:template name="GetTargetPicture">
    <xsl:param name="id"/>
    <xsl:param name="sheet"/>
    <xsl:if
      test="document(concat(concat('xl/worksheets/_rels/', $sheet), '.rels'))//node()[name()='Relationship']">
      <xsl:for-each
        select="document(concat(concat('xl/worksheets/_rels/', $sheet), '.rels'))//node()[name()='Relationship']">
        <xsl:if test="./@Id=$id">
          <xsl:value-of select="@Target"/>
        </xsl:if>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>


  <!-- We check cell when the picture is starting and ending -->

  <xsl:template match="xdr:twoCellAnchor">
    <xsl:param name="PictureCell"/>

    <xsl:variable name="PictureColStart">
      <xsl:value-of select="xdr:from/xdr:col"/>
    </xsl:variable>

    <xsl:variable name="PictureRowStart">
      <xsl:value-of select="xdr:from/xdr:row"/>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="following-sibling::xdr:twoCellAnchor">
        <xsl:apply-templates select="following-sibling::xdr:twoCellAnchor[1]">
          <xsl:with-param name="PictureCell">
            <xsl:value-of
              select="concat(concat(concat(concat($PictureCell, $PictureRowStart), ':'), $PictureColStart), ';')"
            />
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of
          select="concat(concat(concat(concat($PictureCell, $PictureRowStart), ':'), $PictureColStart), ';')"
        />
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <!-- Insert Empty Rows before picture -->
  <xsl:template name="InsertEmptyRows">
    <xsl:param name="repeat"/>
    <xsl:param name="sheet"/>
    <xsl:if test="$repeat &gt; 0">
      <xsl:for-each select="document(concat('xl/',$sheet))">
        <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
          table:number-rows-repeated="{$repeat}">
          <table:table-cell table:number-columns-repeated="256"/>
        </table:table-row>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>

  <!-- Insert Empty Cols before picture -->
  <xsl:template name="InsertEmptyColls">
    <xsl:param name="repeat"/>
    <table:table-cell table:number-columns-repeated="{$repeat}"/>
  </xsl:template>

  <!-- Insert picture -->

  <xsl:template name="InsertPictureEmptySheet">
    <xsl:param name="rowNum"/>
    <xsl:param name="collNum"/>
    <xsl:param name="PictureCell"/>
    <xsl:param name="CollsWithPicture"/>
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="Drawing"/>

    <xsl:if test="$CollsWithPicture != ''">

      <xsl:variable name="CollStart">
        <xsl:value-of select="number(xdr:from/xdr:col) - 1"/>
      </xsl:variable>

      <xsl:variable name="id">
        <xsl:for-each select="document(substring-after($sheet, '/'))">
          <xsl:value-of select="key('drawing', '')/@r:id"/>
        </xsl:for-each>
      </xsl:variable>

      <xsl:if test="$collNum != number(substring-before($CollsWithPicture, ';'))">
        <xsl:call-template name="InsertEmptyColls">
          <xsl:with-param name="repeat">
            <xsl:value-of select="number(substring-before($CollsWithPicture, ';')) - $collNum"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:if>
      <table:table-cell>
        <xsl:for-each
          select="document(concat(concat('xl/worksheets/_rels/', substring-after($sheet, '/')), '.rels'))//node()[name()='Relationship']">
          <xsl:call-template name="CopyPictures">
            <xsl:with-param name="document">
              <xsl:value-of select="concat($Drawing, '.rels')"/>
            </xsl:with-param>
            <xsl:with-param name="targetName">
              <xsl:text>Pictures</xsl:text>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:for-each>

        <xsl:for-each select="xdr:wsDr/xdr:twoCellAnchor">
          <xsl:if
            test="xdr:from/xdr:col = number(substring-before($CollsWithPicture, ';')) and xdr:from/xdr:row = $rowNum">
            <xsl:call-template name="InsertPicture">
              <xsl:with-param name="NameSheet">
                <xsl:value-of select="$NameSheet"/>
              </xsl:with-param>
              <xsl:with-param name="sheet">
                <xsl:value-of select="$sheet"/>
              </xsl:with-param>
              <xsl:with-param name="Drawing">
                <xsl:value-of select="$Drawing"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:if>
        </xsl:for-each>

      </table:table-cell>

      <!-- Insert Next Picture in this row (if exist) -->

      <xsl:if test="substring-after($CollsWithPicture, ';') != '' and $CollsWithPicture != ''">
        <xsl:call-template name="InsertPictureEmptySheet">
          <xsl:with-param name="rowNum">
            <xsl:value-of select="$rowNum"/>
          </xsl:with-param>
          <xsl:with-param name="collNum">
            <xsl:value-of select="substring-before($CollsWithPicture, ';') + 1"/>
          </xsl:with-param>
          <xsl:with-param name="PictureCell">
            <xsl:value-of select="$PictureCell"/>
          </xsl:with-param>
          <xsl:with-param name="CollsWithPicture">
            <xsl:value-of select="substring-after($CollsWithPicture, ';')"/>
          </xsl:with-param>
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="Drawing">
            <xsl:value-of select="$Drawing"/>
          </xsl:with-param>
          <xsl:with-param name="NameSheet">
            <xsl:value-of select="$NameSheet"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:if>
    </xsl:if>

  </xsl:template>

  <xsl:template name="InsertPicture">
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="Drawing"/>

    <xsl:choose>
      <xsl:when test="xdr:pic">
        <draw:frame draw:z-index="0" draw:name="Graphics 1" draw:text-style-name="P1">
          <!--style name-->
          <xsl:attribute name="draw:style-name">
            <xsl:value-of select="generate-id(.)"/>
          </xsl:attribute>

          <!-- position -->
          <xsl:call-template name="SetPosition">
            <xsl:with-param name="sheet">
              <xsl:value-of select="$sheet"/>
            </xsl:with-param>
            <xsl:with-param name="NameSheet">
              <xsl:value-of select="$NameSheet"/>
            </xsl:with-param>
          </xsl:call-template>

          <xsl:call-template name="InsertImageFlip">
            <xsl:with-param name="atribute">
              <xsl:text>draw</xsl:text>
            </xsl:with-param>
          </xsl:call-template>

          <draw:image xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad">

            <xsl:call-template name="InsertImageHref">
              <xsl:with-param name="document">
                <xsl:value-of select="concat($Drawing, '.rels')"/>
              </xsl:with-param>
            </xsl:call-template>

            <text:p/>
          </draw:image>

        </draw:frame>
      </xsl:when>

      <xsl:when test="xdr:graphicFrame/a:graphic/a:graphicData/c:chart">
        <draw:frame draw:z-index="0">
          <!-- position -->
          <xsl:call-template name="SetPosition">
            <xsl:with-param name="sheet">
              <xsl:value-of select="$sheet"/>
            </xsl:with-param>
            <xsl:with-param name="NameSheet">
              <xsl:value-of select="$NameSheet"/>
            </xsl:with-param>
          </xsl:call-template>

          <draw:object xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad">
            <xsl:attribute name="xlink:href">
              <xsl:value-of select="concat('./Object ',generate-id(xdr:graphicFrame/a:graphic/a:graphicData/c:chart))"/>
            </xsl:attribute>
<!--            <draw:object draw:notify-on-update-of-ranges="Sheet1.D6:Sheet1.D6"
              xlink:href="./Object 1"/>-->
          </draw:object>

        </draw:frame>
      </xsl:when>
    </xsl:choose>
  </xsl:template>


  <!-- Get min. row number with picture -->
  <xsl:template name="GetMinRowWithPicture">
    <xsl:param name="min"/>
    <xsl:param name="PictureRow"/>
    <xsl:param name="AfterRow"/>

    <xsl:variable name="numRow">
      <xsl:value-of select="substring-before($PictureRow, ';')"/>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="$AfterRow != ''">

        <xsl:choose>
          <xsl:when test="$PictureRow = ''">
            <xsl:value-of select="$min"/>
          </xsl:when>
          <xsl:when test="$min = ''">
            <xsl:call-template name="GetMinRowWithPicture">
              <xsl:with-param name="min">
                <xsl:if test="$AfterRow &lt;= substring-before($PictureRow, ';')">
                  <xsl:value-of select="substring-before($PictureRow, ';')"/>
                </xsl:if>
              </xsl:with-param>
              <xsl:with-param name="PictureRow">
                <xsl:value-of select="substring-after($PictureRow, ';')"/>
              </xsl:with-param>
              <xsl:with-param name="AfterRow">
                <xsl:value-of select="$AfterRow"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:when
            test="$min &gt; substring-before($PictureRow, ';') and $AfterRow &lt;= substring-before($PictureRow, ';')">
            <xsl:call-template name="GetMinRowWithPicture">
              <xsl:with-param name="min">
                <xsl:value-of select="substring-before($PictureRow, ';')"/>
              </xsl:with-param>
              <xsl:with-param name="PictureRow">
                <xsl:value-of select="substring-after($PictureRow, ';')"/>
              </xsl:with-param>
              <xsl:with-param name="AfterRow">
                <xsl:value-of select="$AfterRow"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="GetMinRowWithPicture">
              <xsl:with-param name="min">
                <xsl:value-of select="$min"/>
              </xsl:with-param>
              <xsl:with-param name="PictureRow">
                <xsl:value-of select="substring-after($PictureRow, ';')"/>
              </xsl:with-param>
              <xsl:with-param name="AfterRow">
                <xsl:value-of select="$AfterRow"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>

      </xsl:when>
      <xsl:otherwise>

        <xsl:choose>
          <xsl:when test="$PictureRow = ''">
            <xsl:value-of select="$min"/>
          </xsl:when>
          <xsl:when test="$min = ''">
            <xsl:call-template name="GetMinRowWithPicture">
              <xsl:with-param name="min">
                <xsl:value-of select="substring-before($PictureRow, ';')"/>
              </xsl:with-param>
              <xsl:with-param name="PictureRow">
                <xsl:value-of select="substring-after($PictureRow, ';')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="$min &gt; substring-before($PictureRow, ';')">
            <xsl:call-template name="GetMinRowWithPicture">
              <xsl:with-param name="min">
                <xsl:value-of select="substring-before($PictureRow, ';')"/>
              </xsl:with-param>
              <xsl:with-param name="PictureRow">
                <xsl:value-of select="substring-after($PictureRow, ';')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="GetMinRowWithPicture">
              <xsl:with-param name="min">
                <xsl:value-of select="$min"/>
              </xsl:with-param>
              <xsl:with-param name="PictureRow">
                <xsl:value-of select="substring-after($PictureRow, ';')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>

      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <!-- Get colls with picture from this row  -->

  <xsl:template name="GetCollsWithPicture">
    <xsl:param name="rowNumber"/>
    <xsl:param name="PictureColl"/>
    <xsl:param name="PictureCell"/>

    <xsl:choose>
      <xsl:when test="contains($PictureCell, concat(concat(';',$rowNumber),':'))">
        <xsl:call-template name="GetCollsWithPicture">
          <xsl:with-param name="rowNumber">
            <xsl:value-of select="$rowNumber"/>
          </xsl:with-param>
          <xsl:with-param name="PictureColl">
            <xsl:value-of
              select="concat($PictureColl, concat(substring-before(substring-after($PictureCell, concat(';', concat($rowNumber, ':'))),';'), ';'))"
            />
          </xsl:with-param>
          <xsl:with-param name="PictureCell">
            <xsl:value-of
              select="concat(';', substring-after(substring-after($PictureCell, concat(';', concat($rowNumber, ':'))),';'))"
            />
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$PictureColl"/>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <!-- Insert Table Body -->

  <xsl:template name="TwoCellAnchor">
    <xsl:param name="prevRow" select="0"/>
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="PictureCell"/>
    <xsl:param name="PictureRow"/>
    <xsl:param name="Drawing"/>

    <xsl:if test="$PictureRow != ''">

      <xsl:variable name="TableStyleName">
        <xsl:for-each select="document(concat('xl/',$sheet))">
          <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
        </xsl:for-each>
      </xsl:variable>

      <xsl:variable name="GetMinRowWithPicture">
        <xsl:call-template name="GetMinRowWithPicture">
          <xsl:with-param name="PictureRow">
            <xsl:value-of select="$PictureRow"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>

      <xsl:call-template name="InsertEmptyRows">
        <xsl:with-param name="repeat">
          <xsl:value-of select="number($GetMinRowWithPicture) - number($prevRow)"/>
        </xsl:with-param>
        <xsl:with-param name="sheet">
          <xsl:value-of select="$sheet"/>
        </xsl:with-param>
      </xsl:call-template>

      <xsl:variable name="CollsWithPicture">
        <xsl:call-template name="GetCollsWithPicture">
          <xsl:with-param name="rowNumber">
            <xsl:value-of select="$GetMinRowWithPicture"/>
          </xsl:with-param>
          <xsl:with-param name="PictureCell">
            <xsl:value-of select="concat(';', $PictureCell)"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>

      <!-- Insert Empty Row with picture (pictures) -->
      <table:table-row table:style-name="{$TableStyleName}">
        <xsl:call-template name="InsertPictureEmptySheet">
          <xsl:with-param name="collNum">
            <xsl:text>0</xsl:text>
          </xsl:with-param>
          <xsl:with-param name="rowNum">
            <xsl:value-of select="$GetMinRowWithPicture"/>
          </xsl:with-param>
          <xsl:with-param name="PictureCell">
            <xsl:value-of select="$PictureCell"/>
          </xsl:with-param>
          <xsl:with-param name="CollsWithPicture">
            <xsl:value-of select="$CollsWithPicture"/>
          </xsl:with-param>
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="Drawing">
            <xsl:value-of select="$Drawing"/>
          </xsl:with-param>
          <xsl:with-param name="NameSheet">
            <xsl:value-of select="$NameSheet"/>
          </xsl:with-param>
        </xsl:call-template>
      </table:table-row>


      <xsl:if
        test="concat(substring-after($PictureRow, concat($GetMinRowWithPicture,';')), substring-before($PictureRow, concat($GetMinRowWithPicture,';'))) != ''">
        <xsl:call-template name="TwoCellAnchor">
          <xsl:with-param name="prevRow">
            <xsl:value-of select="$GetMinRowWithPicture + 1"/>
          </xsl:with-param>
          <xsl:with-param name="PictureRow">
            <xsl:call-template name="DeleteRow">
              <xsl:with-param name="GetMinRowWithPicture">
                <xsl:value-of select="$GetMinRowWithPicture"/>
              </xsl:with-param>
              <xsl:with-param name="PictureRow">
                <xsl:value-of
                  select="concat(substring-after($PictureRow, concat($GetMinRowWithPicture,';')), substring-before($PictureRow, concat($GetMinRowWithPicture,';')))"
                />
              </xsl:with-param>
            </xsl:call-template>
          </xsl:with-param>
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="PictureCell">
            <xsl:value-of select="$PictureCell"/>
          </xsl:with-param>
          <xsl:with-param name="Drawing">
            <xsl:value-of select="$Drawing"/>
          </xsl:with-param>
          <xsl:with-param name="NameSheet">
            <xsl:value-of select="$NameSheet"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:if>
    </xsl:if>


  </xsl:template>

  <!-- delete cell with picture which are inserted -->

  <xsl:template name="DeleteRow">
    <xsl:param name="PictureRow"/>
    <xsl:param name="GetMinRowWithPicture"/>
    <xsl:choose>
      <xsl:when
        test="contains($PictureRow, concat(';', concat($GetMinRowWithPicture,';'))) and $PictureRow != ''">
        <xsl:call-template name="DeleteRow">
          <xsl:with-param name="GetMinRowWithPicture">
            <xsl:value-of select="$GetMinRowWithPicture"/>
          </xsl:with-param>
          <xsl:with-param name="PictureRow">
            <xsl:value-of
              select="concat(substring-after($PictureRow, concat($GetMinRowWithPicture,';')), substring-before($PictureRow, concat($GetMinRowWithPicture,';')))"
            />
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$PictureRow"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <!-- Insert Empty Sheet with picture -->

  <xsl:template name="InsertEmptySheetWithPicture">
    <xsl:param name="PictureCell"/>
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>

    <xsl:variable name="PictureRow">
      <xsl:call-template name="PictureRow">
        <xsl:with-param name="PictureCell">
          <xsl:value-of select="$PictureCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="id">
      <xsl:value-of select="key('drawing', '')/@r:id"/>
    </xsl:variable>

    <xsl:for-each
      select="document(concat(concat('xl/worksheets/_rels/', substring-after($sheet, '/')), '.rels'))//node()[name()='Relationship']">
      <xsl:variable name="Drawing">
        <xsl:value-of select="substring-after(substring-after(@Target, '/'), '/')"/>
      </xsl:variable>
      <xsl:if test="./@Id=$id">
        <xsl:for-each select="document(concat('xl/', substring-after(@Target, '/')))">
          <xsl:if test="xdr:wsDr/xdr:twoCellAnchor">
            <xsl:call-template name="TwoCellAnchor">
              <xsl:with-param name="PictureRow">
                <xsl:value-of select="$PictureRow"/>
              </xsl:with-param>
              <xsl:with-param name="sheet">
                <xsl:value-of select="$sheet"/>
              </xsl:with-param>
              <xsl:with-param name="PictureCell">
                <xsl:value-of select="$PictureCell"/>
              </xsl:with-param>
              <xsl:with-param name="Drawing">
                <xsl:value-of select="$Drawing"/>
              </xsl:with-param>
              <xsl:with-param name="NameSheet">
                <xsl:value-of select="$NameSheet"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:for-each>

  </xsl:template>


  <!-- inserts image href from relationships -->
  <xsl:template name="InsertImageHref">
    <xsl:param name="document"/>
    <xsl:param name="rId"/>
    <xsl:param name="targetName"/>
    <xsl:param name="srcFolder" select="'Pictures'"/>

    <xsl:variable name="id">
      <xsl:choose>
        <xsl:when test="$rId != ''">
          <xsl:value-of select="$rId"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="xdr:pic/xdr:blipFill/a:blip/@r:embed"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:for-each
      select="document(concat('xl/drawings/_rels/',$document))//node()[name() = 'Relationship']">
      <xsl:if test="./@Id=$id">
        <xsl:variable name="targetmode">
          <xsl:value-of select="./@TargetMode"/>
        </xsl:variable>
        <xsl:variable name="pzipsource">
          <xsl:value-of select="./@Target"/>
        </xsl:variable>
        <xsl:variable name="pziptarget">
          <xsl:choose>
            <xsl:when test="$targetName != ''">
              <xsl:value-of select="$targetName"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-after(substring-after($pzipsource,'/'), '/')"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:attribute name="xlink:href">
          <xsl:choose>
            <xsl:when test="$targetmode='External'">
              <xsl:value-of select="$pziptarget"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="concat($srcFolder,'/', $pziptarget)"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

  <!-- Set Position -->
  <xsl:template name="SetPosition">
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:attribute name="table:end-cell-address">
      <xsl:variable name="ColEnd">
        <xsl:call-template name="NumbersToChars">
          <xsl:with-param name="num">
            <xsl:value-of select="xdr:to/xdr:col"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="RowEnd">
        <xsl:value-of select="xdr:to/xdr:row + 1"/>
      </xsl:variable>
      <xsl:value-of select="concat($NameSheet, '.', $ColEnd, $RowEnd)"/>
    </xsl:attribute>
    <xsl:attribute name="svg:x">
      <xsl:call-template name="ConvertEmu">
        <xsl:with-param name="length" select="xdr:from/xdr:colOff"/>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="svg:y">
      <xsl:call-template name="ConvertEmu">
        <xsl:with-param name="length" select="xdr:from/xdr:rowOff"/>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="table:end-x">
      <xsl:call-template name="ConvertEmu">
        <xsl:with-param name="length" select="xdr:to/xdr:colOff"/>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="table:end-y">
      <xsl:call-template name="ConvertEmu">
        <xsl:with-param name="length" select="xdr:to/xdr:rowOff"/>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>

  <xsl:template name="InsertPictureProperties">

    <style:style style:name="{generate-id(.)}" style:family="graphic">

      <!--in Word there are no parent style for image - make default Graphics in OO -->
      <xsl:attribute name="style:parent-style-name">
        <xsl:text>Graphics</xsl:text>
        <!--xsl:value-of select="w:tblStyle/@w:val"/-->
      </xsl:attribute>
      <style:graphic-properties>
        <!--xsl:call-template name="InsertImagePosH"/>
          <xsl:call-template name="InsertImagePosV"/-->
        <!--xsl:call-template name="InsertImageCrop"/>
          <xsl:call-template name="InsertImageWrap"/>
          <xsl:call-template name="InsertImageMargins"/>
            <xsl:call-template name="InsertImageFlowWithtText"/-->
        <xsl:call-template name="InsertImageFlip">
          <xsl:with-param name="atribute">
            <xsl:text>style</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
        <xsl:call-template name="InsertImageBorder"/>
      </style:graphic-properties>
    </style:style>

  </xsl:template>

  <xsl:template match="e:sheet" mode="PictureStyle">
    <xsl:param name="number"/>

    <xsl:variable name="Id">
      <xsl:call-template name="GetTarget">
        <xsl:with-param name="id">
          <xsl:value-of select="@r:id"/>
        </xsl:with-param>
        <xsl:with-param name="document">xl/workbook.xml</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="sheetName">
      <xsl:value-of select="@name"/>
    </xsl:variable>

    <xsl:for-each select="document(concat('xl/',$Id))">
      <xsl:apply-templates select="e:worksheet/e:drawing" mode="PictureStyle">
        <xsl:with-param name="sheet">
          <xsl:value-of select="$Id"/>
        </xsl:with-param>
      </xsl:apply-templates>
    </xsl:for-each>

    <!-- Insert next Table -->

    <xsl:apply-templates select="following-sibling::e:sheet[1]" mode="PictureStyle">
      <xsl:with-param name="number">
        <xsl:value-of select="$number + 1"/>
      </xsl:with-param>
    </xsl:apply-templates>

  </xsl:template>

  <xsl:template match="e:drawing" mode="PictureStyle">
    <xsl:param name="sheet"/>
    <xsl:variable name="Target">
      <xsl:call-template name="GetTargetPicture">
        <xsl:with-param name="sheet">
          <xsl:value-of select="substring-after($sheet, '/')"/>
        </xsl:with-param>
        <xsl:with-param name="id">
          <xsl:value-of select="@r:id"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:for-each select="document(concat('xl/', substring-after($Target, '/')))">
      <xsl:if test="xdr:wsDr/xdr:twoCellAnchor">
        <xsl:apply-templates select="xdr:wsDr/xdr:twoCellAnchor[1]" mode="PictureStyle"/>
      </xsl:if>
    </xsl:for-each>

  </xsl:template>

  <xsl:template match="xdr:twoCellAnchor" mode="PictureStyle">
    <xsl:param name="PictureCell"/>

    <xsl:variable name="PictureColStart">
      <xsl:value-of select="xdr:from/xdr:col"/>
    </xsl:variable>

    <xsl:variable name="PictureRowStart">
      <xsl:value-of select="xdr:from/xdr:row"/>
    </xsl:variable>

    <xsl:call-template name="InsertPictureProperties"/>

    <xsl:apply-templates select="following-sibling::xdr:twoCellAnchor[1]" mode="PictureStyle">
      <xsl:with-param name="PictureCell">
        <xsl:value-of
          select="concat(concat(concat(concat($PictureCell, $PictureRowStart), ':'), $PictureColStart), ';')"
        />
      </xsl:with-param>
    </xsl:apply-templates>

  </xsl:template>

  <!-- To do insert border -->

  <xsl:template name="InsertImageBorder">
    <!--xsl:if
      test="xdr:pic/xdr:spPr/a:ln[not(a:noFill)]">

      <xsl:variable name="width">
        <xsl:call-template name="ConvertEmu3">
          <xsl:with-param name="length">
            <xsl:value-of
              select="xdr:pic/xdr:spPr/a:ln/@w"
            />
          </xsl:with-param>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="type">
        <xsl:choose>
          <xsl:when
            test="xdr:pic/xdr:spPr/a:ln/a:prstDash/@val = 'solid'">
            <xsl:text>solid</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>solid</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:variable name="color">
        <xsl:choose>
          <xsl:when
            test="xdr:pic/xdr:spPr/a:ln/a:solidFill/a:srgbClr">
            <xsl:value-of
              select="xdr:pic/xdr:spPr/a:ln/a:solidFill/a:srgbClr/@val"
            />
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>000000</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:attribute name="fo:border">
        <xsl:value-of select="concat($width,' ',$type,' #',$color)"/>
      </xsl:attribute>
    </xsl:if-->
  </xsl:template>

  <!-- Insert Flip properties -->
  <xsl:template name="InsertImageFlip">
    <xsl:param name="atribute"/>
    <!--  picture flip (vertical, horizontal)-->
    <xsl:if test="xdr:pic/xdr:spPr/a:xfrm/attribute::node()">
      <xsl:choose>
        <!-- to do flip vertical and flip vertical-horizontal -->
        <!--xsl:when test="xdr:pic/xdr:spPr/a:xfrm/@flipV = '1' and xdr:pic/xdr:spPr/a:xfrm/@flipH = '1' and $atribute != 'style'">
             <xsl:attribute name="draw:transform">
              <xsl:text>rotate (3.1415926535892) translate (2.253cm 0.217cm)</xsl:text>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="xdr:pic/xdr:spPr/a:xfrm/@flipV = '1' and $atribute != 'style'">
              <xsl:attribute name="draw:transform">
              <xsl:text>rotate (3.1415926535892) translate (2.256cm 0.108cm)</xsl:text>
              </xsl:attribute>
            </xsl:when-->
        <xsl:when test="xdr:pic/xdr:spPr/a:xfrm/@flipH = '1'  and $atribute = 'style'">
          <xsl:attribute name="style:mirror">
            <xsl:text>horizontal</xsl:text>
          </xsl:attribute>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
  </xsl:template>

  <!-- Insert all picture betwen two rows -->
  <xsl:template name="InsertPictureBetwenTwoRows">
    <xsl:param name="StartRow"/>
    <xsl:param name="EndRow"/>
    <xsl:param name="PictureRow"/>
    <xsl:param name="PictureCell"/>
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>

    <xsl:variable name="GetMinRowWithPicture">
      <xsl:call-template name="GetMinRowWithPicture">
        <xsl:with-param name="PictureRow">
          <xsl:value-of select="$PictureRow"/>
        </xsl:with-param>
        <xsl:with-param name="AfterRow">
          <xsl:value-of select="$StartRow"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:choose>
      <!-- Insert empty rows before -->
      <xsl:when
        test="$GetMinRowWithPicture != '' and $GetMinRowWithPicture &gt;= $StartRow and $GetMinRowWithPicture &lt; $EndRow">
        <xsl:if test="$GetMinRowWithPicture - $StartRow &gt; 0">
          <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
            table:number-rows-repeated="{$GetMinRowWithPicture - $StartRow}">
            <table:table-cell table:number-columns-repeated="256"/>
          </table:table-row>
        </xsl:if>
        <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}">

          <xsl:variable name="PictureColl">
            <xsl:call-template name="GetCollsWithPicture">
              <xsl:with-param name="rowNumber">
                <xsl:value-of select="$GetMinRowWithPicture"/>
              </xsl:with-param>
              <xsl:with-param name="PictureCell">
                <xsl:value-of select="concat(';', $PictureCell)"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>

          <xsl:call-template name="InsertPictureBetwenTwoColl">
            <xsl:with-param name="sheet">
              <xsl:value-of select="$sheet"/>
            </xsl:with-param>
            <xsl:with-param name="NameSheet">
              <xsl:value-of select="$NameSheet"/>
            </xsl:with-param>
            <xsl:with-param name="rowNum">
              <xsl:value-of select="$GetMinRowWithPicture"/>
            </xsl:with-param>
            <xsl:with-param name="PictureColl">
              <xsl:value-of select="$PictureColl"/>
            </xsl:with-param>
            <xsl:with-param name="StartColl">
              <xsl:text>0</xsl:text>
            </xsl:with-param>
            <xsl:with-param name="EndColl">
              <xsl:text>256</xsl:text>
            </xsl:with-param>
          </xsl:call-template>

        </table:table-row>

        <xsl:call-template name="InsertPictureBetwenTwoRows">
          <xsl:with-param name="StartRow">
            <xsl:value-of select="$GetMinRowWithPicture + 1"/>
          </xsl:with-param>
          <xsl:with-param name="EndRow">
            <xsl:value-of select="$EndRow"/>
          </xsl:with-param>
          <xsl:with-param name="PictureRow">
            <xsl:value-of select="$PictureRow"/>
          </xsl:with-param>
          <xsl:with-param name="PictureCell">
            <xsl:value-of select="$PictureCell"/>
          </xsl:with-param>
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="NameSheet">
            <xsl:value-of select="$NameSheet"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <xsl:otherwise>
        <xsl:if test="$EndRow - $StartRow - 1 &gt; 0">
          <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
            table:number-rows-repeated="{$EndRow - $StartRow - 1}">
            <table:table-cell table:number-columns-repeated="256"/>
          </table:table-row>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <!-- Insert all picture betwen two cell -->
  <xsl:template name="InsertPictureBetwenTwoColl">
    <xsl:param name="StartColl"/>
    <xsl:param name="EndColl"/>
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="rowNum"/>
    <xsl:param name="PictureColl"/>
    <xsl:param name="document"/>



    <xsl:variable name="GetMinCollWithPicture">
      <xsl:call-template name="GetMinRowWithPicture">
        <xsl:with-param name="PictureRow">
          <xsl:value-of select="$PictureColl"/>
        </xsl:with-param>
        <xsl:with-param name="AfterRow">
          <xsl:value-of select="$StartColl - 1"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>


    <xsl:choose>
      <!-- Insert empty rows before -->
      <xsl:when
        test="$GetMinCollWithPicture != '' and $GetMinCollWithPicture &gt;= $StartColl and $GetMinCollWithPicture &lt; $EndColl">

        <xsl:if test="$GetMinCollWithPicture - $StartColl &gt; 0">
          <table:table-cell table:number-columns-repeated="{$GetMinCollWithPicture - $StartColl}"/>
        </xsl:if>

        <xsl:for-each select="ancestor::e:worksheet/e:drawing">

          <xsl:variable name="Target">
            <xsl:call-template name="GetTargetPicture">
              <xsl:with-param name="sheet">
                <xsl:value-of select="substring-after($sheet, '/')"/>
              </xsl:with-param>
              <xsl:with-param name="id">
                <xsl:value-of select="@r:id"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>

          <table:table-cell>


            <xsl:call-template name="InsertPictureInThisCell">
              <xsl:with-param name="sheet">
                <xsl:value-of select="$sheet"/>
              </xsl:with-param>
              <xsl:with-param name="NameSheet">
                <xsl:value-of select="$NameSheet"/>
              </xsl:with-param>
              <xsl:with-param name="collNum">
                <xsl:value-of select="$GetMinCollWithPicture"/>
              </xsl:with-param>
              <xsl:with-param name="rowNum">
                <xsl:value-of select="$rowNum"/>
              </xsl:with-param>
              <xsl:with-param name="Target">
                <xsl:value-of select="$Target"/>
              </xsl:with-param>
            </xsl:call-template>
          </table:table-cell>

        </xsl:for-each>

        <xsl:call-template name="InsertPictureBetwenTwoColl">
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="NameSheet">
            <xsl:value-of select="$NameSheet"/>
          </xsl:with-param>
          <xsl:with-param name="rowNum">
            <xsl:value-of select="$rowNum"/>
          </xsl:with-param>
          <xsl:with-param name="PictureColl">
            <xsl:value-of select="$PictureColl"/>
          </xsl:with-param>
          <xsl:with-param name="StartColl">
            <xsl:choose>
              <xsl:when test="$document = 'worksheet'">
                <xsl:value-of select="$GetMinCollWithPicture + 1"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$GetMinCollWithPicture + 2"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="EndColl">
            <xsl:value-of select="$EndColl"/>
          </xsl:with-param>
          <xsl:with-param name="document">
            <xsl:value-of select="$document"/>
          </xsl:with-param>
        </xsl:call-template>

      </xsl:when>

      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$StartColl = 0">
            <table:table-cell table:number-columns-repeated="{$EndColl - $StartColl}"/>
          </xsl:when>
          <xsl:otherwise>
            <table:table-cell table:number-columns-repeated="{$EndColl - $StartColl - 1}"/>
          </xsl:otherwise>

        </xsl:choose>


      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Insert cell with picture -->
  <xsl:template name="InsertPictureInThisCell">
    <xsl:param name="rowNum"/>
    <xsl:param name="collNum"/>
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="Target"/>

    <xsl:for-each select="document(concat('xl/', substring-after($Target, '/')))">
      <xsl:if test="xdr:wsDr/xdr:twoCellAnchor">

        <xsl:for-each
          select="document(concat(concat('xl/worksheets/_rels/', substring-after($sheet, '/')), '.rels'))//node()[name()='Relationship']">
          <xsl:call-template name="CopyPictures">
            <xsl:with-param name="document">
              <xsl:value-of
                select="concat(substring-after(substring-after($Target, '/'), '/'), '.rels')"/>
            </xsl:with-param>
            <xsl:with-param name="targetName">
              <xsl:text>Pictures</xsl:text>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:for-each>

        <xsl:for-each select="xdr:wsDr/xdr:twoCellAnchor">
          <xsl:if test="xdr:from/xdr:col = $collNum and xdr:from/xdr:row = $rowNum">
            <xsl:call-template name="InsertPicture">
              <xsl:with-param name="NameSheet">
                <xsl:value-of select="$NameSheet"/>
              </xsl:with-param>
              <xsl:with-param name="sheet">
                <xsl:value-of select="$sheet"/>
              </xsl:with-param>
              <xsl:with-param name="Drawing">
                <xsl:value-of select="substring-after(substring-after($Target, '/'), '/')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:if>
        </xsl:for-each>

      </xsl:if>
    </xsl:for-each>

  </xsl:template>

</xsl:stylesheet>

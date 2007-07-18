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
  xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"
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

  <!-- Get cell with picture -->
  <xsl:template name="PictureCell">
    <xsl:param name="sheet"/>
    <xsl:param name="mode"/>

    <xsl:choose>
      <!-- when this is a chartsheet (chartsheets/sheet1.xml)-->
      <xsl:when test="$mode = 'chartsheets' ">
        <xsl:apply-templates select="e:chartsheet/e:drawing" mode="chartsheet">
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates select="e:worksheet/e:drawing">
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:otherwise>
    </xsl:choose>
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

  <xsl:template match="e:drawing" mode="chartsheet">
    <xsl:param name="sheet"/>

    <xsl:variable name="id">
      <xsl:value-of select="@r:id"/>
    </xsl:variable>

    <xsl:variable name="Target">
      <xsl:if
        test="document(concat(concat('xl/chartsheets/_rels/', $sheet), '.rels'))//node()[name()='Relationship']">
        <xsl:for-each
          select="document(concat(concat('xl/chartsheets/_rels/', $sheet), '.rels'))//node()[name()='Relationship']">
          <xsl:if test="./@Id=$id">
            <xsl:value-of select="@Target"/>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:variable>

    <xsl:for-each select="document(concat('xl/', substring-after($Target, '/')))">
      <xsl:if test="xdr:wsDr/xdr:absoluteAnchor">
        <xsl:text>1:1;</xsl:text>
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
      <xsl:choose>
        <xsl:when test="xdr:pic/xdr:spPr/a:xfrm/@flipV = 1">
          <xsl:value-of select="xdr:to/xdr:col + 1"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="xdr:from/xdr:col + 1"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="PictureRowStart">
      <xsl:choose>
        <xsl:when test="xdr:pic/xdr:spPr/a:xfrm/@flipV = 1">
          <xsl:value-of select="xdr:to/xdr:row + 1"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="xdr:from/xdr:row + 1"/>
        </xsl:otherwise>
      </xsl:choose>
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

  <!-- Insert picture -->

  <xsl:template name="InsertPicture">
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="Drawing"/>

    <xsl:choose>
      <!-- insert picture -->
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

          <xsl:for-each select="xdr:pic/xdr:spPr">
            <xsl:call-template name="InsertImageFlip">
              <xsl:with-param name="atribute">
                <xsl:text>draw</xsl:text>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:for-each>

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

      <!-- insert chart -->
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
              <xsl:value-of
                select="concat('./Object ',generate-id(xdr:graphicFrame/a:graphic/a:graphicData/c:chart))"
              />
            </xsl:attribute>
          </draw:object>

        </draw:frame>
      </xsl:when>

      <!-- insert text-box -->
      <xsl:when test="xdr:sp/xdr:nvSpPr/xdr:cNvSpPr/@txBox = 1">
        <draw:frame draw:z-index="0">

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

          <draw:text-box>
            <xsl:apply-templates select="xdr:sp/xdr:txBody"/>
          </draw:text-box>

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
            <xsl:choose>
              <xsl:when test="xdr:pic/xdr:spPr/a:xfrm/@flipV = 1">
                <xsl:value-of select="xdr:to/xdr:col + (xdr:to/xdr:col - xdr:from/xdr:col) - 1"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="xdr:to/xdr:col"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="RowEnd">
        <xsl:choose>
          <xsl:when test="xdr:pic/xdr:spPr/a:xfrm/@flipV = 1">
            <xsl:value-of select="xdr:to/xdr:row + (xdr:to/xdr:row - xdr:from/xdr:row) + 1"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="xdr:to/xdr:row + 1"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>


      <xsl:variable name="apos">
        <xsl:text>&apos;</xsl:text>
      </xsl:variable>

      <xsl:variable name="invalidChars">
        <xsl:text>&apos;!$-()</xsl:text>
      </xsl:variable>

      <xsl:variable name="checkedName">
        <xsl:for-each
          select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet[@name = translate($NameSheet,$apos,'')]">
          <xsl:call-template name="CheckSheetName">
            <xsl:with-param name="sheetNumber">
              <xsl:for-each
                select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet[@name = translate($NameSheet,$apos,'')]">
                <xsl:value-of select="count(preceding-sibling::e:sheet) + 1"/>
              </xsl:for-each>
            </xsl:with-param>
            <xsl:with-param name="name">
              <xsl:value-of select="translate($NameSheet,$invalidChars,'')"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:variable>

      <xsl:value-of select="concat($apos,$checkedName,$apos, '.', $ColEnd, $RowEnd)"/>
    </xsl:attribute>

    <xsl:attribute name="svg:x">
      <xsl:choose>
        <xsl:when test="xdr:pic/xdr:spPr/a:xfrm/@flipV = 1">
          <xsl:call-template name="ConvertEmu">
            <xsl:with-param name="length" select="xdr:to/xdr:colOff"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="ConvertEmu">
            <xsl:with-param name="length" select="xdr:from/xdr:colOff"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="svg:y">
      <xsl:call-template name="ConvertEmu">
        <xsl:with-param name="length" select="xdr:from/xdr:rowOff"/>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="table:end-x">
      <xsl:choose>
        <xsl:when test="xdr:pic/xdr:spPr/a:xfrm/@flipV = 1">
          <xsl:call-template name="ConvertEmu">
            <xsl:with-param name="length" select="xdr:from/xdr:colOff"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="ConvertEmu">
            <xsl:with-param name="length" select="xdr:to/xdr:colOff"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="table:end-y">
      <xsl:call-template name="ConvertEmu">
        <xsl:with-param name="length" select="xdr:to/xdr:rowOff"/>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>

  <!-- Insert Picture Style -->

  <xsl:template name="InsertGraphicProperties">

    <xsl:call-template name="InsertImageFlip">
      <xsl:with-param name="atribute">
        <xsl:text>style</xsl:text>
      </xsl:with-param>
    </xsl:call-template>

    <!--<xsl:call-template name="InsertGraphicBorder"/>-->

    <xsl:call-template name="InsertFill"/>
    <xsl:call-template name="InsertLineColor"/>
    <xsl:call-template name="InsertLineStyle"/>

    <xsl:for-each select="parent::node()/xdr:txBody">
      <xsl:call-template name="InsertTextLayout"/>
    </xsl:for-each>

    <xsl:attribute name="fo:min-height">
      <xsl:variable name="border">
        <xsl:choose>
          <xsl:when test="a:ln/@w">

            <xsl:variable name="width">
              <xsl:call-template name="ConvertEmu3">
                <xsl:with-param name="length">
                  <xsl:value-of select="a:ln/@w"/>
                </xsl:with-param>
                <xsl:with-param name="unit">cm</xsl:with-param>
              </xsl:call-template>
            </xsl:variable>

            <xsl:choose>
              <xsl:when test="substring($width,1,1) = '.' ">
                <xsl:value-of select="concat('0',substring-before($width,'cm'))"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="substring-before($width,'cm')"/>
              </xsl:otherwise>
            </xsl:choose>

          </xsl:when>
          <xsl:otherwise>
            <xsl:text>0</xsl:text>
          </xsl:otherwise>
        </xsl:choose>

      </xsl:variable>
      <xsl:value-of select="concat(a:xfrm/a:ext/@cy div 360000 - $border,'cm')"/>
    </xsl:attribute>

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

    <style:style style:name="{generate-id(.)}" style:family="graphic">
      <style:graphic-properties>

        <!-- insert graphic properties -->
        <xsl:choose>
          <xsl:when test="xdr:pic">
            <xsl:for-each select="xdr:pic/xdr:spPr">
              <xsl:call-template name="InsertGraphicProperties"/>
            </xsl:for-each>
          </xsl:when>
          <xsl:when test="xdr:sp/xdr:spPr">
            <xsl:for-each select="xdr:sp/xdr:spPr">
              <xsl:call-template name="InsertGraphicProperties"/>
            </xsl:for-each>
          </xsl:when>
        </xsl:choose>
      </style:graphic-properties>
    </style:style>

    <!--xsl:call-template name="InsertGraphicProperties"/-->

    <xsl:apply-templates select="following-sibling::xdr:twoCellAnchor[1]" mode="PictureStyle">
      <xsl:with-param name="PictureCell">
        <xsl:value-of
          select="concat(concat(concat(concat($PictureCell, $PictureRowStart), ':'), $PictureColStart), ';')"
        />
      </xsl:with-param>
    </xsl:apply-templates>

  </xsl:template>

  <!-- To do insert border -->

  <xsl:template name="InsertGraphicBorder">
    <xsl:if test="a:ln[not(a:noFill)]">

      <xsl:variable name="width">
        <xsl:call-template name="ConvertEmu3">
          <xsl:with-param name="length">
            <xsl:value-of select="a:ln/@w"/>
          </xsl:with-param>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>

      <xsl:variable name="type">
        <xsl:choose>
          <xsl:when test="a:ln/a:prstDash/@val = 'solid'">
            <xsl:text>solid</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>solid</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>

      <xsl:variable name="color">
        <xsl:choose>
          <xsl:when test="a:ln/a:solidFill/a:srgbClr/@val != ''">
            <xsl:value-of select="a:ln/a:solidFill/a:srgbClr/@val"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>000000</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>

      <xsl:attribute name="draw:stroke">
        <xsl:value-of select="$type"/>
      </xsl:attribute>

      <xsl:attribute name="draw:stroke-dash">
        <xsl:text>Line_20_with_20_Fine_20_Dots</xsl:text>
      </xsl:attribute>

      <xsl:attribute name="svg:stroke-width">
        <xsl:value-of select="$width"/>
      </xsl:attribute>

      <xsl:attribute name="svg:stroke-color">
        <xsl:value-of select="concat('#', $color)"/>
      </xsl:attribute>

    </xsl:if>
  </xsl:template>

  <!-- Insert Flip properties -->
  <xsl:template name="InsertImageFlip">
    <xsl:param name="atribute"/>
    <!--  picture flip (vertical, horizontal)-->
    <xsl:if test="a:xfrm/attribute::node()">
      <xsl:choose>
        <!-- TO DO Vertical  -->
        <xsl:when test="a:xfrm/@flipV = '1' and $atribute != 'style'">
          <xsl:attribute name="draw:transform">
            <xsl:text>rotate (3.1415926535892) translate (2.064cm 0.425cm)</xsl:text>
          </xsl:attribute>
        </xsl:when>
        <!-- horizontal -->
        <xsl:when test="a:xfrm/@flipH = '1'  and $atribute = 'style'">
          <xsl:attribute name="style:mirror">
            <xsl:text>horizontal</xsl:text>
          </xsl:attribute>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
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

          <xsl:choose>
            <xsl:when test="xdr:pic/xdr:spPr/a:xfrm/@flipV = 1">
              <xsl:if test="(xdr:to/xdr:col + 1) = $collNum and (xdr:to/xdr:row + 1) = $rowNum">

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
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="(xdr:from/xdr:col + 1) = $collNum and (xdr:from/xdr:row + 1) = $rowNum">
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
            </xsl:otherwise>
          </xsl:choose>

        </xsl:for-each>

      </xsl:if>
    </xsl:for-each>

  </xsl:template>

  <xsl:template name="InsertChartsheet">
    <xsl:param name="number"/>
    <xsl:param name="sheet"/>

    <table:table>

      <!-- Insert Table (Sheet) Name -->

      <xsl:variable name="checkedName">
        <xsl:call-template name="CheckSheetName">
          <xsl:with-param name="sheetNumber">
            <xsl:value-of select="$number"/>
          </xsl:with-param>
          <xsl:with-param name="name">
            <xsl:value-of select="translate(@name,'!$-()','')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>

      <xsl:attribute name="table:name">
        <!--        <xsl:value-of select="@name"/>-->
        <xsl:value-of select="$checkedName"/>
      </xsl:attribute>

      <!-- Insert Table Style Name (style:table-properties) -->

      <xsl:attribute name="table:style-name">
        <xsl:value-of select="generate-id()"/>
      </xsl:attribute>

      <!-- drawing.xml file path -->
      <xsl:variable name="Target">
        <xsl:for-each select="document(concat('xl/',$sheet))/e:chartsheet/e:drawing">
          <xsl:variable name="id">
            <xsl:value-of select="@r:id"/>
          </xsl:variable>
          <xsl:if
            test="document(concat(concat('xl/chartsheets/_rels/', substring-after($sheet,'/')), '.rels'))//node()[name()='Relationship']">
            <xsl:for-each
              select="document(concat(concat('xl/chartsheets/_rels/', substring-after($sheet,'/')), '.rels'))//node()[name()='Relationship']">
              <xsl:if test="./@Id=$id">
                <xsl:value-of select="@Target"/>
              </xsl:if>
            </xsl:for-each>
          </xsl:if>
        </xsl:for-each>
      </xsl:variable>

      <office:forms form:automatic-focus="false" form:apply-design-mode="false"/>
      <table:shapes>

        <xsl:for-each select="document(concat('xl/', substring-after($Target, '/')))">
          <xsl:for-each select="xdr:wsDr/xdr:absoluteAnchor">

            <draw:frame draw:z-index="0" svg:width="7.999cm" svg:height="6.999cm" svg:x="0cm"
              svg:y="0cm">

              <xsl:call-template name="InsertAbsolutePosition"/>
              <xsl:call-template name="InsertAbsoluteSize"/>

              <draw:object xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad">
                <xsl:attribute name="xlink:href">
                  <xsl:value-of
                    select="concat('./Object ',generate-id(xdr:graphicFrame/a:graphic/a:graphicData/c:chart))"
                  />
                </xsl:attribute>
              </draw:object>

            </draw:frame>
          </xsl:for-each>
        </xsl:for-each>

      </table:shapes>
      <table:table-column
        table:style-name="{generate-id(document('xl/worksheets/sheet1.xml')/e:worksheet/e:sheetFormatPr)}"/>
      <table:table-row
        table:style-name="{generate-id(document('xl/worksheets/sheet1.xml')/e:worksheet/e:sheetFormatPr)}">
        <table:table-cell/>
      </table:table-row>
    </table:table>
  </xsl:template>

  <xsl:template name="InsertAbsolutePosition">

    <xsl:if test="xdr:pos">
      <xsl:attribute name="svg:x">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="xdr:pos/@x"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="svg:y">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="xdr:pos/@y"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertAbsoluteSize">

    <xsl:if test="xdr:ext">
      <xsl:attribute name="svg:width">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="xdr:ext/@cx"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="svg:height">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="xdr:ext/@cy"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

  <xsl:template match="a:p">
    <text:p>
      <xsl:if test="a:pPr">
        <xsl:attribute name="text:style-name">
          <xsl:value-of select="generate-id(.)"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:apply-templates/>
    </text:p>
  </xsl:template>

  <xsl:template match="a:r">
    <text:span>
      <xsl:if test="a:rPr">
        <xsl:attribute name="text:style-name">
          <xsl:value-of select="generate-id(.)"/>
        </xsl:attribute>
      </xsl:if>

      <xsl:apply-templates/>
    </text:span>
  </xsl:template>

  <xsl:template match="a:t">
    <xsl:choose>
      <!--check whether string contains  whitespace sequence-->
      <xsl:when test="not(contains(., '  '))">
        <xsl:choose>
          <!-- single space before case -->
          <xsl:when test="substring(text(),1,1) = ' ' ">
            <text:s/>
            <xsl:value-of select="substring-after(text(),' ')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="."/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <!--converts whitespaces sequence to text:s -->
        <!-- inside "if" when text starts with a single space -->
        <xsl:if test="substring(text(),1,1) = ' ' and substring(text(),2,1) != ' ' ">
          <text:s/>
        </xsl:if>
        <xsl:call-template name="InsertWhiteSpaces"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertFill">

    <xsl:choose>
      <!-- No fill -->
      <xsl:when test="a:noFill">

        <xsl:attribute name="draw:fill">
          <xsl:value-of select="'none'"/>
        </xsl:attribute>

        <xsl:attribute name="draw:fill-color">
          <xsl:value-of select="'#ffffff'"/>
        </xsl:attribute>
      </xsl:when>

      <!-- Solid fill-->
      <xsl:when test="a:solidFill">
        <xsl:attribute name="draw:fill">
          <xsl:value-of select="'solid'"/>
        </xsl:attribute>

        <!-- Standard color-->
        <xsl:if test="a:solidFill/a:srgbClr/@val">
          <xsl:attribute name="draw:fill-color">
            <xsl:value-of select="concat('#',a:solidFill/a:srgbClr/@val)"/>
          </xsl:attribute>

          <!-- Transparency percentage-->
          <xsl:if test="a:solidFill/a:srgbClr/a:alpha/@val">
            <xsl:variable name="alpha">
              <xsl:value-of select="a:solidFill/a:srgbClr/a:alpha/@val"/>
            </xsl:variable>

            <xsl:if test="($alpha != '') or ($alpha != 0)">
              <xsl:attribute name="draw:opacity">
                <xsl:value-of select="concat(($alpha div 1000), '%')"/>
              </xsl:attribute>
            </xsl:if>

          </xsl:if>
        </xsl:if>

        <!--Theme color-->
        <xsl:if test="a:solidFill/a:schemeClr/@val">

          <xsl:attribute name="draw:fill-color">
            <xsl:call-template name="getColorCode">
              <xsl:with-param name="color">
                <xsl:value-of select="a:solidFill/a:schemeClr/@val"/>
              </xsl:with-param>
              <xsl:with-param name="lumMod">
                <xsl:value-of select="a:solidFill/a:schemeClr/a:lumMod/@val"/>
              </xsl:with-param>
              <xsl:with-param name="lumOff">
                <xsl:value-of select="a:solidFill/a:schemeClr/a:lumOff/@val"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>

          <!-- Transparency percentage-->
          <xsl:if test="a:solidFill/a:schemeClr/a:alpha/@val">
            <xsl:variable name="alpha">
              <xsl:value-of select="a:solidFill/a:schemeClr/a:alpha/@val"/>
            </xsl:variable>

            <xsl:if test="($alpha != '') or ($alpha != 0)">
              <xsl:attribute name="draw:opacity">
                <xsl:value-of select="concat(($alpha div 1000), '%')"/>
              </xsl:attribute>
            </xsl:if>

          </xsl:if>
        </xsl:if>
      </xsl:when>

      <!-- fill from style -->
      <xsl:when test="parent::node()/xdr:style/a:fillRef">
        <xsl:attribute name="draw:fill">
          <xsl:value-of select="'solid'"/>
        </xsl:attribute>

        <!-- Standard color-->
        <xsl:if test="parent::node()/xdr:style/a:fillRef/a:srgbClr/@val">
          <xsl:attribute name="draw:fill-color">
            <xsl:value-of select="concat('#',parent::node()/xdr:style/a:fillRef/a:srgbClr/@val)"/>
          </xsl:attribute>

          <!-- Shade percentage-->
          <!--<xsl:if test="xdr:style/a:fillRef/a:srgbClr/a:shade/@val">
              <xsl:variable name ="shade">
              <xsl:value-of select ="a:solidFill/a:srgbClr/a:shade/@val"/>
              </xsl:variable>
              <xsl:if test="($shade != '') or ($shade != 0)">
              <xsl:attribute name ="draw:shadow-opacity">
              <xsl:value-of select="concat(($shade div 1000), '%')"/>
              </xsl:attribute>
              </xsl:if>
              </xsl:if>-->
        </xsl:if>

        <!--Theme color-->
        <xsl:if test="parent::node()/xdr:style/a:fillRef//a:schemeClr/@val">
          <xsl:attribute name="draw:fill-color">
            <xsl:call-template name="getColorCode">
              <xsl:with-param name="color">
                <xsl:value-of select="parent::node()/xdr:style/a:fillRef/a:schemeClr/@val"/>
              </xsl:with-param>
              <xsl:with-param name="lumMod">
                <xsl:value-of select="parent::node()/xdr:style/a:fillRef/a:schemeClr/a:lumMod/@val"
                />
              </xsl:with-param>
              <xsl:with-param name="lumOff">
                <xsl:value-of select="parent::node()/xdr:style/a:fillRef/a:schemeClr/a:lumOff/@val"
                />
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>

          <!-- Shade percentage-->
          <!--<xsl:if test="a:solidFill/a:schemeClr/a:shade/@val">
              <xsl:variable name ="shade">
              <xsl:value-of select ="a:solidFill/a:schemeClr/a:shade/@val"/>
              </xsl:variable>
              <xsl:if test="($shade != '') or ($shade != 0)">
              <xsl:attribute name ="draw:shadow-opacity">
              <xsl:value-of select="concat(($shade div 1000), '%')"/>
              </xsl:attribute>
              </xsl:if>
              </xsl:if>-->
        </xsl:if>

      </xsl:when>
      <!-- default fill for textbox -->
        <xsl:otherwise>
        <xsl:if test="parent::node()/xdr:nvSpPr/xdr:cNvSpPr/@txBox = 1">
          <xsl:attribute name="draw:fill">
            <xsl:text>none</xsl:text>
          </xsl:attribute>
        </xsl:if>
        </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertLineColor">

    <xsl:choose>
      <!-- No line-->
      <xsl:when test="a:ln/a:noFill">
        <xsl:attribute name="draw:stroke">
          <xsl:value-of select="'none'"/>
        </xsl:attribute>
      </xsl:when>

      <!-- Solid line color-->
      <xsl:when test="a:ln/a:solidFill">
        <xsl:attribute name="draw:stroke">
          <xsl:value-of select="'solid'"/>
        </xsl:attribute>

        <!-- Standard color for border-->
        <xsl:if test="a:ln/a:solidFill/a:srgbClr/@val">
          <xsl:attribute name="svg:stroke-color">
            <xsl:value-of select="concat('#',a:ln/a:solidFill/a:srgbClr/@val)"/>
          </xsl:attribute>

          <!-- Transparency percentage-->
          <xsl:if test="a:ln/a:solidFill/a:srgbClr/a:alpha/@val">
            <xsl:variable name="alpha">
              <xsl:value-of select="a:ln/a:solidFill/a:srgbClr/a:alpha/@val"/>
            </xsl:variable>

            <xsl:if test="($alpha != '') or ($alpha != 0)">
              <xsl:attribute name="svg:stroke-opacity">
                <xsl:value-of select="concat(($alpha div 1000), '%')"/>
              </xsl:attribute>
            </xsl:if>

          </xsl:if>
        </xsl:if>

        <!-- Theme color for border-->
        <xsl:if test="a:ln/a:solidFill/a:schemeClr/@val">
          <xsl:attribute name="svg:stroke-color">
            <xsl:call-template name="getColorCode">
              <xsl:with-param name="color">
                <xsl:value-of select="a:ln/a:solidFill/a:schemeClr/@val"/>
              </xsl:with-param>
              <xsl:with-param name="lumMod">
                <xsl:value-of select="a:ln/a:solidFill/a:schemeClr/a:lumMod/@val"/>
              </xsl:with-param>
              <xsl:with-param name="lumOff">
                <xsl:value-of select="a:ln/a:solidFill/a:schemeClr/a:lumOff/@val"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>

          <!-- Transparency percentage-->
          <xsl:if test="a:ln/a:solidFill/a:schemeClr/a:alpha/@val">
            <xsl:variable name="alpha">
              <xsl:value-of select="a:ln/a:solidFill/a:schemeClr/a:alpha/@val"/>
            </xsl:variable>

            <xsl:if test="($alpha != '') or ($alpha != 0)">
              <xsl:attribute name="svg:stroke-opacity">
                <xsl:value-of select="concat(($alpha div 1000), '%')"/>
              </xsl:attribute>
            </xsl:if>

          </xsl:if>
        </xsl:if>
      </xsl:when>

      <xsl:otherwise>
        <!--Line reference-->
        <xsl:if
          test="not( (a:prstGeom/@prst='flowChartInternalStorage') or
          (a:prstGeom/@prst='flowChartPredefinedProcess') or
          (a:prstGeom/@prst='flowChartSummingJunction') or
          (a:prstGeom/@prst='flowChartOr') )">
          <xsl:if test="parent::node()/xdr:style/a:lnRef">
            <xsl:attribute name="draw:stroke">
              <xsl:value-of select="'solid'"/>
            </xsl:attribute>

            <!--Standard color for border-->
            <xsl:if test="parent::node()/xdr:style/a:lnRef/a:srgbClr/@val">
              <xsl:attribute name="svg:stroke-color">
                <xsl:value-of select="concat('#',parent::node()/xdr:style/a:lnRef/a:srgbClr/@val)"/>
              </xsl:attribute>

              <!--Shade percentage-->
              <!--
                <xsl:if test="xdr:style/a:lnRef/a:srgbClr/a:shade/@val">
                <xsl:variable name ="shade">
                <xsl:value-of select ="xdr:style/a:lnRef/a:srgbClr/a:shade/@val"/>
                </xsl:variable>
              -->
              <!--<xsl:if test="($shade != '') or ($shade != 0)">
                <xsl:attribute name ="svg:stroke-opacity">
                <xsl:value-of select="concat(($shade div 1000), '%')"/>
                </xsl:attribute>
                </xsl:if>-->
              <!--
                </xsl:if>-->
            </xsl:if>

            <!--Theme color for border-->
            <xsl:if test="parent::node()/xdr:style/a:lnRef/a:schemeClr/@val">

              <xsl:attribute name="svg:stroke-color">
                <xsl:call-template name="getColorCode">
                  <xsl:with-param name="color">
                    <xsl:value-of select="parent::node()/xdr:style/a:lnRef/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name="lumMod">
                    <xsl:value-of
                      select="parent::node()/xdr:style/a:lnRef/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name="lumOff">
                    <xsl:value-of
                      select="parent::node()/xdr:style/a:lnRef/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>

              <!--Shade percentage -->
              <!--<xsl:if test="xdr:style/a:lnRef/a:schemeClr/a:shade/@val">
                <xsl:variable name ="shade">
                <xsl:value-of select ="xdr:style/a:lnRef/a:schemeClr/a:shade/@val"/>
                </xsl:variable>
              -->
              <!--<xsl:if test="($shade != '') or ($shade != 0)">
                <xsl:attribute name ="svg:stroke-opacity">
                <xsl:value-of select="concat(($shade div 1000), '%')"/>
                </xsl:attribute>
                </xsl:if>-->
              <!--
                </xsl:if>-->
            </xsl:if>

          </xsl:if>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertLineStyle">

    <!-- Line width-->
    <xsl:choose>
      <xsl:when test="a:ln/@w">
        <xsl:attribute name="svg:stroke-width">
          <xsl:call-template name="ConvertEmu">
            <xsl:with-param name="length" select="a:ln/@w"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:when>

      <!-- line width from theme -->
      <xsl:when test="parent::node()/xdr:style/a:lnRef/@idx">
        <xsl:variable name="index">
          <xsl:value-of select="parent::node()/xdr:style/a:lnRef/@idx + 1"/>
        </xsl:variable>
        <xsl:attribute name="svg:stroke-width">
          <xsl:call-template name="ConvertEmu">
            <xsl:with-param name="length"
              select="document('xl/theme/theme1.xml')/a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln[position() = $index]/@w"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:when>

      <!-- TextBox default border -->
      <xsl:when test="not(a:ln/@w) and (parent::node()/xdr:nvSpPr/xdr:cNvSpPr/@txBox = 1)">
        <xsl:attribute name="draw:stroke">
          <xsl:value-of select="'none'"/>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>


    <!-- Line Dash property-->
    <xsl:choose>
      <xsl:when test="not(a:ln/a:noFill) and a:ln/a:prstDash">
        <xsl:for-each select="a:ln">
          <xsl:call-template name="InsertLineDash"/>
        </xsl:for-each>
      </xsl:when>

      <!-- line dash from theme -->
      <xsl:when test="parent::node()/xdr:style/a:lnRef/@idx">
        <xsl:variable name="index">
          <xsl:value-of select="parent::node()/xdr:style/a:lnRef/@idx + 1"/>
        </xsl:variable>

        <xsl:for-each
          select="document('xl/theme/theme1.xml')/a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln[position() = $index]">
          <xsl:call-template name="InsertLineDash"/>
        </xsl:for-each>
      </xsl:when>
    </xsl:choose>

    <!-- Line join property -->
    <xsl:choose>
      <xsl:when test="a:ln/a:miter">
        <xsl:attribute name="draw:stroke-linejoin">
          <xsl:text>miter</xsl:text>
        </xsl:attribute>
      </xsl:when>

      <xsl:when test="a:ln/a:bevel">
        <xsl:attribute name="draw:stroke-linejoin">
          <xsl:text>bevel</xsl:text>
        </xsl:attribute>
      </xsl:when>

      <xsl:when test="a:ln/a:round">
        <xsl:attribute name="draw:stroke-linejoin">
          <xsl:text>round</xsl:text>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>

    <!-- Line Arrow -->
    <!-- Head End-->
    <xsl:for-each select="a:ln/a:headEnd">
      <xsl:if test="@type">
        <xsl:attribute name="draw:marker-start">
          <xsl:value-of select="@type"/>
        </xsl:attribute>

        <xsl:attribute name="draw:marker-start-width">
          <xsl:call-template name="getArrowSize">
            <xsl:with-param name="w" select="@w"/>
            <xsl:with-param name="len" select="@len"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>

    <!-- Tail End-->
    <xsl:for-each select="a:ln/a:tailEnd">
      <xsl:if test="@type">
        <xsl:attribute name="draw:marker-end">
          <xsl:value-of select="@type"/>
        </xsl:attribute>

        <xsl:attribute name="draw:marker-end-width">
          <xsl:call-template name="getArrowSize">
            <xsl:with-param name="w" select="@w"/>
            <xsl:with-param name="len" select="@len"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="getArrowSize">
    <xsl:param name="w"/>
    <xsl:param name="len"/>

    <xsl:choose>
      <xsl:when test="($w = 'sm' ) and ($len = 'sm' )">
        <xsl:value-of select="concat(0.15,'cm' )"/>
      </xsl:when>

      <xsl:when test="($w = 'sm' ) and ($len = 'med' )">
        <xsl:value-of select="concat(0.18,'cm' )"/>
      </xsl:when>

      <xsl:when test="($w = 'sm' ) and ($len = 'lg' )">
        <xsl:value-of select="concat(0.2,'cm' )"/>
      </xsl:when>

      <xsl:when test="($w = 'med' ) and ($len = 'sm' )">
        <xsl:value-of select="concat(0.21,'cm' )"/>
      </xsl:when>

      <xsl:when test="($w = 'med' ) and ($len = 'lg' )">
        <xsl:value-of select="concat(0.3,'cm' )"/>
      </xsl:when>

      <xsl:when test="($w = 'lg' ) and ($len = 'sm' )">
        <xsl:value-of select="concat(0.31,'cm' )"/>
      </xsl:when>

      <xsl:when test="($w = 'lg' ) and ($len = 'med' )">
        <xsl:value-of select="concat(0.35,'cm' )"/>
      </xsl:when>

      <xsl:when test="($w = 'lg' ) and ($len = 'lg' )">
        <xsl:value-of select="concat(0.4,'cm' )"/>
      </xsl:when>

      <xsl:otherwise>
        <xsl:value-of select="concat(0.25,'cm' )"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertLineDash">

    <xsl:if test="a:prstDash/@val">
      <xsl:attribute name="draw:stroke">
        <xsl:text>solid</xsl:text>
      </xsl:attribute>
    </xsl:if>

    <!--xsl:choose>
      <xsl:when test="(a:prstDash/@val = 'solid' ) or not(a:prstDash/@val)">
        <xsl:attribute name="draw:stroke">
          <xsl:text>solid</xsl:text>
        </xsl:attribute>
      </xsl:when>

      <xsl:otherwise>
        <xsl:attribute name="draw:stroke">
          <xsl:text>dash</xsl:text>
        </xsl:attribute>

        <xsl:attribute name="draw:stroke-dash">

          <xsl:choose>
            <xsl:when test="(a:prstDash/@val = 'sysDot' ) and (@cap = 'rnd' )">
              <xsl:text>sysDotRound</xsl:text>
            </xsl:when>

            <xsl:when test="a:prstDash/@val = 'sysDot' ">
              <xsl:text>sysDot</xsl:text>
            </xsl:when>

            <xsl:when test="(a:prstDash/@val = 'sysDash' ) and (@cap = 'rnd' )">
              <xsl:text>sysDashRound</xsl:text>
            </xsl:when>

            <xsl:when test="a:prstDash/@val = 'sysDash' ">
              <xsl:text>sysDash</xsl:text>
            </xsl:when>

            <xsl:when test="(a:prstDash/@val = 'dash' ) and (@cap = 'rnd' )">
              <xsl:text>dashRound</xsl:text>
            </xsl:when>

            <xsl:when test="a:prstDash/@val = 'dash' ">
              <xsl:text>dash</xsl:text>
            </xsl:when>

            <xsl:when test="(a:prstDash/@val = 'dashDot' ) and (@cap = 'rnd' )">
              <xsl:text>dashDotRound</xsl:text>
            </xsl:when>

            <xsl:when test="a:prstDash/@val = 'dashDot' ">
              <xsl:text>dashDot</xsl:text>
            </xsl:when>

            <xsl:when test="(a:prstDash/@val = 'lgDash' ) and (@cap = 'rnd' )">
              <xsl:text>lgDashRound</xsl:text>
            </xsl:when>

            <xsl:when test="a:prstDash/@val = 'lgDash' ">
              <xsl:text>lgDash</xsl:text>
            </xsl:when>

            <xsl:when test="(a:prstDash/@val = 'lgDashDot' ) and (@cap = 'rnd' )">
              <xsl:text>lgDashDotRound</xsl:text>
            </xsl:when>

            <xsl:when test="a:prstDash/@val = 'lgDashDot' ">
              <xsl:text>lgDashDot</xsl:text>
            </xsl:when>

            <xsl:when test="(a:prstDash/@val = 'lgDashDotDot' ) and (@cap = 'rnd' )">
              <xsl:text>lgDashDotDotRound</xsl:text>
            </xsl:when>

            <xsl:when test="a:prstDash/@val = 'lgDashDotDot' ">
              <xsl:text>lgDashDotDot</xsl:text>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>

      </xsl:otherwise>
    </xsl:choose-->
  </xsl:template>

  <xsl:template name="InsertTextLayout">

    <xsl:if test="a:bodyPr/@lIns">
      <xsl:attribute name="fo:padding-left">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="a:bodyPr/@lIns"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="not(a:bodyPr/@lIns)">
      <xsl:attribute name="fo:padding-left">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="'91440'"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="a:bodyPr/@tIns">
      <xsl:attribute name="fo:padding-top">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="a:bodyPr/@tIns"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="not(a:bodyPr/@tIns)">
      <xsl:attribute name="fo:padding-top">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="'45720'"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="a:bodyPr/@rIns">
      <xsl:attribute name="fo:padding-right">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="a:bodyPr/@rIns"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="not(a:bodyPr/@rIns)">
      <xsl:attribute name="fo:padding-right">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="'91440'"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="a:bodyPr/@bIns">
      <xsl:attribute name="fo:padding-bottom">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="a:bodyPr/@bIns"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="not(a:bodyPr/@bIns)">
      <xsl:attribute name="fo:padding-bottom">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="'45720'"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="( (a:bodyPr/a:spAutoFit) or (a:bodyPr/@wrap='square') )">
      <xsl:attribute name="draw:auto-grow-height">
        <xsl:value-of select="'true'"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="not(a:bodyPr/a:spAutoFit)">
      <xsl:attribute name="draw:auto-grow-height">
        <xsl:value-of select="'false'"/>
      </xsl:attribute>
      <xsl:attribute name="draw:auto-grow-width">
        <xsl:value-of select="'false'"/>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertTextBoxTextStyles">
    <!-- get all sheet Id's -->
    <xsl:for-each select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet">

      <xsl:variable name="sheet">
        <xsl:call-template name="GetTarget">
          <xsl:with-param name="id">
            <xsl:value-of select="@r:id"/>
          </xsl:with-param>
          <xsl:with-param name="document">xl/workbook.xml</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <!--i.e. $sheet = worksheets/sheet1.xml -->

      <!-- go to worksheet file and search for drawing -->
      <xsl:for-each select="document(concat('xl/',$sheet))/e:worksheet/e:drawing">

        <xsl:variable name="drawing">
          <xsl:call-template name="GetTarget">
            <xsl:with-param name="id">
              <xsl:value-of select="@r:id"/>
            </xsl:with-param>
            <xsl:with-param name="document">
              <xsl:value-of select="concat('xl/',$sheet)"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <!-- i.e. $drawing = ../drawings/drawing2.xml -->

        <!-- finally insert entry for each chart -->
        <xsl:for-each
          select="document(concat('xl/',substring-after($drawing,'/')))/xdr:wsDr/xdr:twoCellAnchor/xdr:sp/xdr:txBody">

          <xsl:for-each select="a:p/a:pPr">
            <xsl:call-template name="InsertTextBoxParagraphStyle"/>
          </xsl:for-each>

          <xsl:for-each select="a:p/a:r/a:rPr">
            <xsl:call-template name="InsertTextBoxRunStyle"/>
          </xsl:for-each>

        </xsl:for-each>
      </xsl:for-each>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="InsertTextBoxParagraphStyle">
    <style:style style:name="{generate-id(parent::node())}" style:family="paragraph">
      <style:paragraph-properties>

        <!-- horizontal alignment -->
        <xsl:if test="@algn ='ctr' or @algn ='r' or @algn ='l' or @algn ='just' ">
          <xsl:attribute name="fo:text-align">
            <xsl:choose>
              <!-- Center Alignment-->
              <xsl:when test="@algn ='ctr' ">
                <xsl:text>center</xsl:text>
              </xsl:when>
              <!-- Right Alignment-->
              <xsl:when test="@algn ='r' ">
                <xsl:text>end</xsl:text>
              </xsl:when>
              <!-- Left Alignment-->
              <xsl:when test="@algn ='l' ">
                <xsl:text>start</xsl:text>
              </xsl:when>
              <xsl:when test="@algn ='just' ">
                <xsl:text>justify</xsl:text>
              </xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>

        <!-- left margin -->
        <xsl:if test="@marL">
          <xsl:attribute name="fo:margin-left">
            <xsl:value-of select="concat(format-number(@marL div 360000, '0.##'), 'cm')"/>
          </xsl:attribute>
        </xsl:if>

        <!-- first line indent -->
        <xsl:if test="@indent">
          <xsl:attribute name="fo:text-indent">
            <xsl:choose>
              <xsl:when test="@indent &gt; 0">
                <xsl:value-of select="concat(format-number(@indent div 360000, '0.##'), 'cm')"/>
              </xsl:when>
              <xsl:when test="@indent = 0">
                <xsl:text>0</xsl:text>
              </xsl:when>
              <xsl:when
                test="(@indent &lt; 0) and number(substring-after(@indent,'-')) &lt;= @marL">
                <xsl:value-of select="concat(format-number(@indent div 360000, '0.##'), 'cm')"/>
              </xsl:when>
              <xsl:when
                test="(@indent &lt; 0) and number(substring-after(@indent,'-')) &gt; @marL">
                <xsl:value-of select="concat(format-number(@marL div 360000, '0.##'), 'cm')"/>
              </xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>

        <!-- spacing before -->
        <xsl:if test="a:spcBef/a:spcPts/@val">
          <xsl:attribute name="fo:margin-top">
            <xsl:value-of
              select="concat(format-number(a:spcBef/a:spcPts/@val div 2835, '0.##'), 'cm')"/>
          </xsl:attribute>
        </xsl:if>

        <!-- spacig after-->
        <xsl:if test="a:spcAft/a:spcPts/@val">
          <xsl:attribute name="fo:margin-bottom">
            <xsl:value-of
              select="concat(format-number(a:spcAft/a:spcPts/@val div 2835, '0.##'), 'cm')"/>
          </xsl:attribute>
        </xsl:if>

        <!-- If the line space is in Percentage-->
        <xsl:if test="a:lnSpc/a:spcPct/@val">
          <xsl:attribute name="fo:line-height">
            <xsl:value-of select="concat(format-number(a:lnSpc/a:spcPct/@val div 1000,'###'), '%')"
            />
          </xsl:attribute>
        </xsl:if>

        <!-- If the line space is in Points-->
        <xsl:if test="a:lnSpc/a:spcPts">
          <xsl:attribute name="style:line-height-at-least">
            <xsl:value-of
              select="concat(format-number(a:lnSpc/a:spcPts/@val div 2835, '0.##'), 'cm')"/>
          </xsl:attribute>
        </xsl:if>
      </style:paragraph-properties>
    </style:style>
  </xsl:template>

  <xsl:template name="InsertTextBoxRunStyle">

    <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
    <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>

    <style:style style:name="{generate-id(parent::node())}" style:family="text">
      <style:text-properties fo:font-family="Calibri" fo:font-size="11pt" fo:font-weight="normal"
        fo:color="#000000">

        <xsl:choose>
          <xsl:when test="a:latin/@typeface != '+mn-lt' and a:latin/@typeface != '+mj-lt' ">
            <xsl:attribute name="fo:font-family">
              <xsl:value-of select="a:latin/@typeface"/>
            </xsl:attribute>
          </xsl:when>
          <!--xsl:when test="not(a:latin) and a:cs">
            <xsl:attribute name="fo:font-family">
              <xsl:value-of select="a:cs/@typeface"/>
            </xsl:attribute>
          </xsl:when-->
        </xsl:choose>

        <!-- font size -->
        <xsl:if test="@sz">
          <xsl:attribute name="fo:font-size">
            <xsl:value-of select="concat(format-number(@sz div 100,'0.##'), 'pt')"/>
          </xsl:attribute>
        </xsl:if>

        <!-- bold -->
        <xsl:if test="@b = 1">
          <xsl:attribute name="fo:font-weight">
            <xsl:value-of select="'bold'"/>
          </xsl:attribute>
        </xsl:if>

        <!-- Italic-->
        <xsl:if test="@i = 1">
          <xsl:attribute name="fo:font-style">
            <xsl:value-of select="'italic'"/>
          </xsl:attribute>
        </xsl:if>

        <!-- font color -->
        <xsl:choose>
          <xsl:when test="a:solidFill/a:srgbClr/@val">
            <xsl:attribute name="fo:color">
              <xsl:value-of select="concat('#',a:solidFill/a:srgbClr/@val)"/>
            </xsl:attribute>
          </xsl:when>

          <!-- scheme color -->
          <xsl:when test="a:solidFill/a:schemeClr/@val">
            <xsl:attribute name="fo:color">
              <xsl:call-template name="getColorCode">
                <xsl:with-param name="color">
                  <xsl:value-of select="a:solidFill/a:schemeClr/@val"/>
                </xsl:with-param>
                <xsl:with-param name="lumMod">
                  <xsl:value-of select="a:solidFill/a:schemeClr/a:lumMod/@val"/>
                </xsl:with-param>
                <xsl:with-param name="lumOff">
                  <xsl:value-of select="a:solidFill/a:schemeClr/a:lumOff/@val"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:when>

          <!-- style color -->
          <xsl:when
            test="parent::node()/parent::node()/parent::node()/parent::node()/xdr:style/a:fontRef/a:srgbClr">
            <xsl:attribute name="fo:color">
              <xsl:value-of
                select="concat('#',parent::node()/parent::node()/parent::node()/parent::node()/xdr:style/a:fontRef/a:srgbClr/@val)"
              />
            </xsl:attribute>
          </xsl:when>

          <!-- style scheme color -->
          <xsl:when
            test="parent::node()/parent::node()/parent::node()/parent::node()/xdr:style/a:fontRef/a:schemeClr">
            <xsl:for-each
              select="parent::node()/parent::node()/parent::node()/parent::node()/xdr:style/a:fontRef/a:schemeClr">
              <xsl:attribute name="fo:color">
                <xsl:call-template name="getColorCode">
                  <xsl:with-param name="color">
                    <xsl:value-of select="@val"/>
                  </xsl:with-param>
                  <xsl:with-param name="lumMod">
                    <xsl:value-of select="a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name="lumOff">
                    <xsl:value-of select="a:lumOff/@val"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:for-each>
          </xsl:when>
        </xsl:choose>

        <!-- style:text-underline-style -->
        <xsl:if test="@u">
          <!-- underline style -->
          <xsl:choose>
            <xsl:when test="@u = 'sng' ">
              <xsl:attribute name="style:text-underline-type">
                <xsl:text>single</xsl:text>
              </xsl:attribute>
              <xsl:attribute name="style:text-underline-width">
                <xsl:text>auto</xsl:text>
              </xsl:attribute>
            </xsl:when>

            <!-- double -->
            <xsl:when test="@u = 'dbl' ">
              <xsl:attribute name="style:text-underline-style">
                <xsl:text>solid</xsl:text>
              </xsl:attribute>
              <xsl:attribute name="style:text-underline-type">
                <xsl:text>double</xsl:text>
              </xsl:attribute>
              <xsl:attribute name="style:text-underline-width">
                <xsl:text>auto</xsl:text>
              </xsl:attribute>
            </xsl:when>

            <!-- bold -->
            <xsl:when test="@u= 'heavy' ">
              <xsl:attribute name="style:text-underline-style">
                <xsl:text>solid</xsl:text>
              </xsl:attribute>
              <xsl:attribute name="style:text-underline-width">
                <xsl:text>bold</xsl:text>
              </xsl:attribute>
            </xsl:when>

            <!-- dotted-->
            <xsl:when test="@u= 'dotted' ">
              <xsl:attribute name="style:text-underline-style">
                <xsl:text>dotted</xsl:text>
              </xsl:attribute>
              <xsl:attribute name="style:text-underline-width">
                <xsl:text>auto</xsl:text>
              </xsl:attribute>
            </xsl:when>

            <!-- bold dotted -->
            <xsl:when test="@u = 'dottedHeavy' ">
              <xsl:attribute name="style:text-underline-style">
                <xsl:text>dotted</xsl:text>
              </xsl:attribute>
              <xsl:attribute name="style:text-underline-width">
                <xsl:text>bold</xsl:text>
              </xsl:attribute>
            </xsl:when>

            <!-- dash -->
            <xsl:when test="@u = 'dash' ">
              <xsl:attribute name="style:text-underline-style">
                <xsl:text>dash</xsl:text>
              </xsl:attribute>
              <xsl:attribute name="style:text-underline-width">
                <xsl:text>auto</xsl:text>
              </xsl:attribute>
            </xsl:when>

            <!-- bold dash -->
            <xsl:when test="@u = 'dashHeavy' ">
              <xsl:attribute name="style:text-underline-style">
                <xsl:text>dash</xsl:text>
              </xsl:attribute>
              <xsl:attribute name="style:text-underline-width">
                <xsl:text>bold</xsl:text>
              </xsl:attribute>
            </xsl:when>

            <!-- long dash -->
            <xsl:when test="@u = 'dashLong' ">
              <xsl:attribute name="style:text-underline-style">
                <xsl:text>long-dash</xsl:text>
              </xsl:attribute>
              <xsl:attribute name="style:text-underline-width">
                <xsl:text>auto</xsl:text>
              </xsl:attribute>
            </xsl:when>

            <!-- bold long dash -->
            <xsl:when test="@u = 'dashLongHeavy' ">
              <xsl:attribute name="style:text-underline-style">
                <xsl:text>long-dash</xsl:text>
              </xsl:attribute>
              <xsl:attribute name="style:text-underline-width">
                <xsl:text>bold</xsl:text>
              </xsl:attribute>
            </xsl:when>

            <!-- dot-dash -->
            <xsl:when test="@u = 'dotDash' ">
              <xsl:attribute name="style:text-underline-style">
                <xsl:text>dot-dash</xsl:text>
              </xsl:attribute>
              <xsl:attribute name="style:text-underline-width">
                <xsl:text>auto</xsl:text>
              </xsl:attribute>
            </xsl:when>

            <!-- bold dot-dash -->
            <xsl:when test="@u = 'dotDashHeavy' ">
              <xsl:attribute name="style:text-underline-style">
                <xsl:text>dot-dash</xsl:text>
              </xsl:attribute>
              <xsl:attribute name="style:text-underline-width">
                <xsl:text>bold</xsl:text>
              </xsl:attribute>
            </xsl:when>

            <!-- dot-dot-dash -->
            <xsl:when test="@u = 'dotDotDash' ">
              <xsl:attribute name="style:text-underline-style">
                <xsl:text>dot-dot-dash</xsl:text>
              </xsl:attribute>
              <xsl:attribute name="style:text-underline-width">
                <xsl:text>auto</xsl:text>
              </xsl:attribute>
            </xsl:when>

            <!-- bold dot-dot-dash -->
            <xsl:when test="@u = 'dotDotDashHeavy' ">
              <xsl:attribute name="style:text-underline-style">
                <xsl:text>dot-dot-dash</xsl:text>
              </xsl:attribute>
              <xsl:attribute name="style:text-underline-width">
                <xsl:text>bold</xsl:text>
              </xsl:attribute>
            </xsl:when>

            <!-- wavy -->
            <xsl:when test="@u = 'wavy' ">
              <xsl:attribute name="style:text-underline-style">
                <xsl:text>wave</xsl:text>
              </xsl:attribute>
              <xsl:attribute name="style:text-underline-width">
                <xsl:text>auto</xsl:text>
              </xsl:attribute>
            </xsl:when>

            <!-- bold wave -->
            <xsl:when test="@u = 'wavyHeavy' ">
              <xsl:attribute name="style:text-underline-style">
                <xsl:text>wave</xsl:text>
              </xsl:attribute>
              <xsl:attribute name="style:text-underline-width">
                <xsl:text>bold</xsl:text>
              </xsl:attribute>
            </xsl:when>

            <!-- double wave -->
            <xsl:when test="@u = 'wavyDbl' ">
              <xsl:attribute name="style:text-underline-style">
                <xsl:text>wave</xsl:text>
              </xsl:attribute>
              <xsl:attribute name="style:text-underline-type">
                <xsl:text>double</xsl:text>
              </xsl:attribute>
              <xsl:attribute name="style:text-underline-width">
                <xsl:text>auto</xsl:text>
              </xsl:attribute>
            </xsl:when>

            <xsl:otherwise>
              <xsl:attribute name="style:text-underline-type">
                <xsl:text>single</xsl:text>
              </xsl:attribute>
              <xsl:attribute name="style:text-underline-width">
                <xsl:text>auto</xsl:text>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>

          <!-- underline color -->
          <xsl:if test="a:uFill/a:solidFill/a:srgbClr/@val">
            <xsl:attribute name="style:text-underline-color">
              <xsl:value-of select="concat('#',a:uFill/a:solidFill/a:srgbClr/@val)"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="a:uFill/a:solidFill/a:schemeClr/@val">
            <xsl:attribute name="style:text-underline-color">
              <xsl:call-template name="getColorCode">
                <xsl:with-param name="color">
                  <xsl:value-of select="a:uFill/a:solidFill/a:schemeClr/@val"/>
                </xsl:with-param>
                <xsl:with-param name="lumMod">
                  <xsl:value-of select="a:uFill/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                </xsl:with-param>
                <xsl:with-param name="lumOff">
                  <xsl:value-of select="a:uFill/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>

        <!-- strike-through -->
        <xsl:if test="@strike">
          <xsl:attribute name="style:text-line-through-style">
            <xsl:text>solid</xsl:text>
          </xsl:attribute>

          <xsl:choose>
            <xsl:when test="@strike = 'sngStrike' ">
              <xsl:attribute name="style:text-line-through-type">
                <xsl:value-of select="'single'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="@strike = 'dblStrike' ">
              <xsl:attribute name="style:text-line-through-type">
                <xsl:text>double</xsl:text>
              </xsl:attribute>
            </xsl:when>
          </xsl:choose>
        </xsl:if>

        <!-- text position -->
        <xsl:if test="@baseline != 0">
          <xsl:attribute name="style:text-position">
            <xsl:value-of select="@baseline div 1000"/>
            <xsl:text> 75%</xsl:text>
          </xsl:attribute>
        </xsl:if>

        <!-- character spacing -->
        <xsl:if test="@spc">
          <xsl:attribute name="fo:letter-spacing">
            <xsl:variable name="length" select="@spc"/>
            <xsl:value-of select="concat(format-number($length * 2.54 div 7200,'0.###'),'cm')"/>
          </xsl:attribute>
        </xsl:if>

        <!--Kerning true or false -->
        <xsl:if test="@kern != 0">
          <xsl:attribute name="style:letter-kerning">
            <xsl:text>true</xsl:text>
          </xsl:attribute>
        </xsl:if>

        <!--Shadow fo:text-shadow-->
        <xsl:if test="a:effectLst/a:outerShdw">
          <xsl:attribute name="fo:text-shadow">
            <xsl:text>1pt 1pt</xsl:text>
          </xsl:attribute>
        </xsl:if>

      </style:text-properties>

    </style:style>
  </xsl:template>

</xsl:stylesheet>

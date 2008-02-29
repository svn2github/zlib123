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
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
  xmlns:oox="urn:oox"
  exclude-result-prefixes="e oox r">

  <xsl:import href="common.xsl"/>
  <xsl:import href="relationships.xsl"/>
  <xsl:import href="border.xsl"/>
  <xsl:import href="styles.xsl"/>

  <!-- Get cell when the conditional is starting and ending -->
  <xsl:template name="ConditionalCell">
    <xsl:param name="sheet"/>
    <xsl:param name="document"/>

    <!--xsl:apply-templates select="e:worksheet/e:conditionalFormatting[1]">
      <xsl:with-param name="document">
        <xsl:value-of select="$document"/>
      </xsl:with-param>
    </xsl:apply-templates-->
  </xsl:template>

  <!-- Get Row with Conditional -->
  <xsl:template name="ConditionalRow">
    <xsl:param name="ConditionalCell"/>
    <xsl:param name="Result"/>
    <xsl:choose>
      <xsl:when test="$ConditionalCell != ''">
        <xsl:call-template name="ConditionalRow">
          <xsl:with-param name="ConditionalCell">
            <xsl:value-of select="substring-after($ConditionalCell, ';')"/>
          </xsl:with-param>
          <xsl:with-param name="Result">
            <xsl:value-of
              select="concat($Result,  concat(substring-before($ConditionalCell, ':'), ';'))"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$Result"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

 

  <xsl:template match="e:conditionalFormatting">
    <xsl:param name="ConditionalCell"/>
    <xsl:param name="document"/>
	  <xsl:choose>
		  <!--Defect Id: 1803593, file 'ilas_EVE_Database_V0.02.xlsx'
		      Changes Done by: Vijayeta
			  Date: 10th Jan '08
			  Databar, colourscales and iconsets not supported in Open Office, hence do not execute this part of code-->
		  <xsl:when test ="./e:cfRule and not(./e:cfRule[@type= 'dataBar'] or ./e:cfRule[@type= 'colorScale'] or ./e:cfRule[@type= 'iconSet'])">
    <xsl:variable name="colNum">
      <xsl:call-template name="GetColNum">
        <xsl:with-param name="cell">
          <xsl:choose>
            <xsl:when test="contains(@sqref, ':')">
              <xsl:value-of select="substring-before(@sqref, ':')"/>
            </xsl:when>
            <xsl:when test="contains(@sqref, ' ')">
              <xsl:value-of select="substring-before(@sqref, ' ')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="@sqref"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="rowNum">
      <xsl:call-template name="GetRowNum">
        <xsl:with-param name="cell">
          <xsl:choose>
            <xsl:when test="contains(@sqref, ':')">
              <xsl:value-of select="substring-before(@sqref, ':')"/>
            </xsl:when>
            <xsl:when test="contains(@sqref, ' ')">
              <xsl:value-of select="substring-before(@sqref, ' ')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="@sqref"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="dxfIdStyle">
      <xsl:value-of select="count(preceding-sibling::e:conditionalFormatting)"/>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="contains(@sqref, ':')">
        <xsl:choose>
          <xsl:when test="following-sibling::e:conditionalFormatting">
            <xsl:apply-templates select="following-sibling::e:conditionalFormatting[1]">
              <xsl:with-param name="ConditionalCell">
                <xsl:call-template name="InsertConditionalCell">
                  <xsl:with-param name="ConditionalCell">
                    <xsl:value-of select="$ConditionalCell"/>
                  </xsl:with-param>
                  <xsl:with-param name="StartCell">
                    <xsl:value-of select="substring-before(@sqref, ':')"/>
                  </xsl:with-param>
                  <xsl:with-param name="EndCell">
                    <xsl:value-of select="substring-after(@sqref, ':')"/>
                  </xsl:with-param>
                  <xsl:with-param name="document">
                    <xsl:value-of select="$document"/>
                  </xsl:with-param>
                  <xsl:with-param name="dxfIdStyle">
                    <xsl:value-of select="$dxfIdStyle"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:with-param>
              <xsl:with-param name="document">
                <xsl:value-of select="$document"/>
              </xsl:with-param>
            </xsl:apply-templates>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="InsertConditionalCell">
              <xsl:with-param name="ConditionalCell">
                <xsl:value-of select="$ConditionalCell"/>
              </xsl:with-param>
              <xsl:with-param name="StartCell">
                <xsl:value-of select="substring-before(@sqref, ':')"/>
              </xsl:with-param>
              <xsl:with-param name="EndCell">
                <xsl:value-of select="substring-after(@sqref, ':')"/>
              </xsl:with-param>
              <xsl:with-param name="document">
                <xsl:value-of select="$document"/>
              </xsl:with-param>
              <xsl:with-param name="dxfIdStyle">
                <xsl:value-of select="$dxfIdStyle"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>

      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="following-sibling::e:conditionalFormatting">
            <xsl:apply-templates select="following-sibling::e:conditionalFormatting[1]">
              <xsl:with-param name="ConditionalCell">
                <xsl:choose>
                  <xsl:when test="$document='style'">
                    <xsl:value-of
                      select="concat($rowNum, ':', $colNum, ';', '-', $dxfIdStyle, ';', $ConditionalCell)"
                    />
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="concat($rowNum, ':', $colNum, ';', $ConditionalCell)"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:with-param>
              <xsl:with-param name="document">
                <xsl:value-of select="$document"/>
              </xsl:with-param>
            </xsl:apply-templates>
          </xsl:when>
          <xsl:otherwise>
            <xsl:choose>
              <xsl:when test="$document='style'">
                <xsl:value-of
                  select="concat($rowNum, ':', $colNum, ';', '-', $dxfIdStyle, ';', $ConditionalCell)"
                />
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="concat($rowNum, ':', $colNum, ';', $ConditionalCell)"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
		  </xsl:when >

		  </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertConditionalCell">
    <xsl:param name="ConditionalCell"/>
    <xsl:param name="StartCell"/>
    <xsl:param name="EndCell"/>
    <xsl:param name="document"/>
    <xsl:param name="dxfIdStyle"/>

    <xsl:variable name="StartColNum">
      <xsl:call-template name="GetColNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="$StartCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="StartRowNum">
      <xsl:call-template name="GetRowNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="$StartCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="EndColNum">
      <xsl:call-template name="GetColNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="$EndCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="EndRowNum">
      <xsl:call-template name="GetRowNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="$EndCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="RepeatRowConditional">
      <xsl:call-template name="RepeatRowConditional">
        <xsl:with-param name="StartColNum">
          <xsl:value-of select="$StartColNum"/>
        </xsl:with-param>
        <xsl:with-param name="EndColNum">
          <xsl:value-of select="$EndColNum"/>
        </xsl:with-param>
        <xsl:with-param name="StartRowNum">
          <xsl:value-of select="$StartRowNum"/>
        </xsl:with-param>
        <xsl:with-param name="EndRowNum">
          <xsl:value-of select="$EndRowNum"/>
        </xsl:with-param>
        <xsl:with-param name="document">
          <xsl:value-of select="$document"/>
        </xsl:with-param>
        <xsl:with-param name="dxfIdStyle">
          <xsl:value-of select="$dxfIdStyle"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:value-of select="concat($ConditionalCell, $RepeatRowConditional)"/>

  </xsl:template>

  <xsl:template name="RepeatRowConditional">
    <xsl:param name="ConditionalCell"/>
    <xsl:param name="StartColNum"/>
    <xsl:param name="StartRowNum"/>
    <xsl:param name="EndColNum"/>
    <xsl:param name="EndRowNum"/>
    <xsl:param name="document"/>
    <xsl:param name="dxfIdStyle"/>

    <xsl:choose>
      <xsl:when test="$StartRowNum &lt;= $EndRowNum">

        <xsl:call-template name="RepeatColConditional">
          <xsl:with-param name="StartColNum">
            <xsl:value-of select="$StartColNum"/>
          </xsl:with-param>
          <xsl:with-param name="EndColNum">
            <xsl:value-of select="$EndColNum"/>
          </xsl:with-param>
          <xsl:with-param name="StartRowNum">
            <xsl:value-of select="$StartRowNum"/>
          </xsl:with-param>
          <xsl:with-param name="EndRowNum">
            <xsl:value-of select="$EndRowNum"/>
          </xsl:with-param>
          <xsl:with-param name="document">
            <xsl:value-of select="$document"/>
          </xsl:with-param>
          <xsl:with-param name="dxfIdStyle">
            <xsl:value-of select="$dxfIdStyle"/>
          </xsl:with-param>
        </xsl:call-template>

        <xsl:call-template name="RepeatRowConditional">
          <xsl:with-param name="ConditionalCell">
            <xsl:choose>
              <xsl:when test="$document='style'">
                <xsl:value-of
                  select="concat($StartRowNum, ':', $StartColNum, ';', '-', $dxfIdStyle, ';',$ConditionalCell)"
                />
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of
                  select="concat($StartRowNum, ':', $StartColNum, ';', $ConditionalCell)"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="StartRowNum">
            <xsl:value-of select="$StartRowNum + 1"/>
          </xsl:with-param>
          <xsl:with-param name="StartColNum">
            <xsl:value-of select="$StartColNum"/>
          </xsl:with-param>
          <xsl:with-param name="EndRowNum">
            <xsl:value-of select="$EndRowNum"/>
          </xsl:with-param>
          <xsl:with-param name="EndColNum">
            <xsl:value-of select="$EndColNum"/>
          </xsl:with-param>
          <xsl:with-param name="document">
            <xsl:value-of select="$document"/>
          </xsl:with-param>
          <xsl:with-param name="dxfIdStyle">
            <xsl:value-of select="$dxfIdStyle"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="RepeatColConditional">
    <xsl:param name="ConditionalCell"/>
    <xsl:param name="StartColNum"/>
    <xsl:param name="StartRowNum"/>
    <xsl:param name="EndColNum"/>
    <xsl:param name="EndRowNum"/>
    <xsl:param name="document"/>
    <xsl:param name="dxfIdStyle"/>

    <xsl:choose>
      <xsl:when test="$StartColNum != $EndColNum">
        <xsl:call-template name="RepeatColConditional">
          <xsl:with-param name="ConditionalCell">
            <xsl:choose>
              <xsl:when test="$document='style'">
                <xsl:value-of
                  select="concat($StartRowNum, ':', $StartColNum, ';', '-', $dxfIdStyle, ';',$ConditionalCell)"
                />
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of
                  select="concat($StartRowNum, ':', $StartColNum, ';', $ConditionalCell)"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="StartRowNum">
            <xsl:value-of select="$StartRowNum "/>
          </xsl:with-param>
          <xsl:with-param name="StartColNum">
            <xsl:value-of select="$StartColNum + 1"/>
          </xsl:with-param>
          <xsl:with-param name="EndRowNum">
            <xsl:value-of select="$EndRowNum"/>
          </xsl:with-param>
          <xsl:with-param name="EndColNum">
            <xsl:value-of select="$EndColNum"/>
          </xsl:with-param>
          <xsl:with-param name="document">
            <xsl:value-of select="$document"/>
          </xsl:with-param>
          <xsl:with-param name="dxfIdStyle">
            <xsl:value-of select="$dxfIdStyle"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$document='style'">
            <xsl:value-of
              select="concat($StartRowNum, ':', $StartColNum, ';', '-', $dxfIdStyle, ';',$ConditionalCell)"
            />
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat($StartRowNum, ':', $StartColNum, ';', $ConditionalCell)"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="e:sheet" mode="ConditionalStyle">
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

    <!-- Check If Conditionals are in this sheet -->

    <xsl:variable name="ConditionalCell">
      <!--xsl:for-each select="key('Part', concat('xl/',$Id))">
        <xsl:call-template name="ConditionalCell"/>
      </xsl:for-each-->
    </xsl:variable>

    <xsl:variable name="ConditionalCellStyle">
      <!--xsl:for-each select="key('Part', concat('xl/',$Id))">
        <xsl:call-template name="ConditionalCell">
          <xsl:with-param name="document">
            <xsl:text>style</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:for-each-->
    </xsl:variable>

    <xsl:for-each select="key('Part', concat('xl/',$Id))">

      <xsl:apply-templates select="e:worksheet/e:conditionalFormatting" mode="ConditionalStyle">
        <xsl:with-param name="sheet">
          <xsl:value-of select="$Id"/>
        </xsl:with-param>
      </xsl:apply-templates>

      <xsl:apply-templates select="e:worksheet/e:sheetData/e:row/e:c" mode="ConditionalAndCellStyle">
        <xsl:with-param name="ConditionalCell">
          <xsl:value-of select="$ConditionalCell"/>
        </xsl:with-param>
        <xsl:with-param name="ConditionalCellStyle">
          <xsl:value-of select="$ConditionalCellStyle"/>
        </xsl:with-param>
      </xsl:apply-templates>

    </xsl:for-each>

    <!-- Insert next Table -->

    <xsl:apply-templates select="following-sibling::e:sheet[1]" mode="ConditionalStyle">
      <xsl:with-param name="number">
        <xsl:value-of select="$number + 1"/>
      </xsl:with-param>
    </xsl:apply-templates>

  </xsl:template>

  <xsl:template match="e:conditionalFormatting" mode="ConditionalStyle">
    <xsl:call-template name="InsertConditionalProperties"/>
  </xsl:template>

  <xsl:template name="InsertConditionalProperties">
    
    <style:style style:name="{generate-id(.)}" style:family="table-cell"
      style:parent-style-name="Default">
      <xsl:call-template name="InsertConditional"/>
    </style:style>
  </xsl:template>

  <!-- Insert Conditional -->
  <xsl:template name="InsertConditional">
    <xsl:for-each select="e:cfRule">
      <xsl:sort select="@priority"/>
      <style:map>
        <xsl:attribute name="style:apply-style-name">
          <xsl:variable name="PositionStyle">
            <xsl:value-of select="@dxfId"/>
          </xsl:variable>
          <xsl:choose>
            <!-- if there is a specified style for cells fullfilling certain condition -->
            <xsl:when test="@dxfId != ''">
              <xsl:value-of select="generate-id(key('Dxf', @dxfId))"/>
            </xsl:when>
            <!-- default style -->
            <xsl:otherwise>
              <xsl:text>Default</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:choose>
          <xsl:when test="@operator='equal'">
            <xsl:attribute name="style:condition">
              <xsl:text>cell-content()=</xsl:text>
              <xsl:value-of select="e:formula"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="@operator='lessThanOrEqual'">
            <xsl:attribute name="style:condition">
              <xsl:text>cell-content()&lt;=</xsl:text>
              <xsl:value-of select="e:formula"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="@operator='lessThan'">
            <xsl:attribute name="style:condition">
              <xsl:text>cell-content()&lt;</xsl:text>
              <xsl:value-of select="e:formula"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="@operator='greaterThan'">
            <xsl:attribute name="style:condition">
              <xsl:text>cell-content()&gt;</xsl:text>
              <xsl:value-of select="e:formula"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="@operator='greaterThanOrEqual'">
            <xsl:attribute name="style:condition">
              <xsl:text>cell-content()&gt;=</xsl:text>
              <xsl:value-of select="e:formula"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="@operator='notEqual'">
            <xsl:attribute name="style:condition">
              <xsl:text>cell-content()!=</xsl:text>
              <xsl:value-of select="e:formula"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="@operator='between'">
            <xsl:attribute name="style:condition">
              <xsl:value-of
                select="concat('cell-content-is-between(', e:formula, ',', e:formula[2], ')') "/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="@operator='notBetween'">
            <xsl:attribute name="style:condition">
              <xsl:value-of
                select="concat('cell-content-is-not-between(', e:formula, ',', e:formula[2], ')') "
              />
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="style:condition">
              <xsl:text>is-true-formula(FALSE)</xsl:text>
              <!--for cfRules not supported by converter;-->
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </style:map>
    </xsl:for-each>
  </xsl:template>

  <!-- Insert Coditional Styles -->

  <xsl:template name="InsertConditionalStyles">
    <xsl:for-each select="key('Part', 'xl/styles.xml')">
      <xsl:apply-templates select="e:styleSheet/e:dxfs/e:dxf"/>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="e:dxf">
    <style:style style:name="{generate-id(.)}" style:family="table-cell">
      <style:text-properties>
        <xsl:apply-templates select="e:font[1]" mode="style"/>
      </style:text-properties>
      <style:table-cell-properties>
        <xsl:variable name="this" select="."/>
        <xsl:apply-templates select="e:fill" mode="style"/>
        <xsl:call-template name="InsertBorder"/>
      </style:table-cell-properties>
    </style:style>
  </xsl:template>

  <!-- Insert Cell Style and Conditional Style -->
  <xsl:template match="e:c" mode="ConditionalAndCellStyle">
    <xsl:param name="ConditionalCell"/>
    <xsl:param name="ConditionalCellStyle"/>

    <xsl:variable name="colNum">
      <xsl:call-template name="GetColNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="@r"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="rowNum">
      <xsl:call-template name="GetRowNum">
        <xsl:with-param name="cell">
          <xsl:value-of select="@r"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if
      test="contains(concat(';', $ConditionalCellStyle), concat(';', $rowNum, ':', $colNum, ';')) and @s != ''">

      <xsl:variable name="CellStyleNumber">
        <xsl:value-of select="@s"/>
      </xsl:variable>

      <style:style style:family="table-cell">

        <xsl:attribute name="style:name">
          <xsl:variable name="CellStyleId">
            <xsl:for-each select="key('Part', 'xl/styles.xml')/e:styleSheet">
              <xsl:for-each select="key('Xf', $CellStyleNumber)">
                <xsl:value-of select="generate-id(.)"/>
              </xsl:for-each>
            </xsl:for-each>
          </xsl:variable>
          <xsl:variable name="ConditionalStyleId">
            <xsl:for-each
              select="key('ConditionalFormatting', ancestor::e:worksheet/@oox:part)[@oox:id = substring-before(substring-after(concat(';', $ConditionalCellStyle), concat(';', $rowNum, ':', $colNum, ';-')), ';')]">
              <xsl:value-of select="generate-id(.)"/>
            </xsl:for-each>
          </xsl:variable>
          <xsl:value-of select="concat($CellStyleId, $ConditionalStyleId)"/>
        </xsl:attribute>

        <xsl:for-each select="key('Xf', $CellStyleNumber)">
          <xsl:call-template name="InsertCellFormat"/>
        </xsl:for-each>

        <xsl:for-each
          select="key('ConditionalFormatting', ancestor::e:worksheet/@oox:part)[@oox:id = substring-before(substring-after(concat(';', $ConditionalCellStyle), concat(';', $rowNum, ':', $colNum, ';-')), ';')]">
          <xsl:call-template name="InsertConditional"/>
        </xsl:for-each>
      </style:style>

    </xsl:if>

  </xsl:template>
  
  
  <!-- Get cell when the conditional is starting and ending -->
  <xsl:template name="ConditionalCellCol">
    <xsl:param name="sheet"/>
    <xsl:param name="document"/>
    
    <xsl:apply-templates select="e:worksheet/e:conditionalFormatting[1]" mode="Col">
      <xsl:with-param name="document">
        <xsl:value-of select="$document"/>
      </xsl:with-param>
    </xsl:apply-templates>
    
  </xsl:template>
  
  <xsl:template match="e:conditionalFormatting" mode="Col">
    <xsl:param name="ConditionalCellCol"/>
    <xsl:param name="document"/>
    
    <xsl:choose>
      
      <xsl:when test="following-sibling::e:conditionalFormatting">
        
        <xsl:apply-templates select="following-sibling::e:conditionalFormatting[1]" mode="Col">
          <xsl:with-param name="ConditionalCellCol">
            <xsl:call-template name="InsertConditionalCol">
              <xsl:with-param name="result">
                <xsl:value-of select="$ConditionalCellCol"/>
              </xsl:with-param>
              <xsl:with-param name="sqref">
                <xsl:value-of select="@sqref"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:with-param>
          <xsl:with-param name="document">
            <xsl:value-of select="$document"/>
          </xsl:with-param>
        </xsl:apply-templates>
        
      </xsl:when>
      
      <xsl:otherwise>
        
        <xsl:call-template name="InsertConditionalCol">
          <xsl:with-param name="result">
            <xsl:value-of select="$ConditionalCellCol"/>
          </xsl:with-param>
          <xsl:with-param name="sqref">
            <xsl:value-of select="@sqref"/>
          </xsl:with-param>
        </xsl:call-template>  
        
      </xsl:otherwise>
      
    </xsl:choose>
    
    
      
    
    
  </xsl:template>
  
  <xsl:template name="InsertConditionalCol">
    <xsl:param name="sqref"/>
    <xsl:param name="result"/>
 
    <xsl:choose>
      <xsl:when test="substring-before($sqref, ' ') != ''">
        <xsl:call-template name="InsertConditionalCol">
          <xsl:with-param name="sqref">
            <xsl:value-of select="substring-after($sqref, ' ')"/>
          </xsl:with-param>
          <xsl:with-param name="result">
            <xsl:call-template name="SingleCol">
              <xsl:with-param name="result">
                <xsl:value-of select="$result"/>
              </xsl:with-param>
              <xsl:with-param name="sqref">
                <xsl:value-of select="substring-before($sqref, ' ')"/>
              </xsl:with-param>
            </xsl:call-template>            
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="substring-before($sqref, ' ') = '' and $sqref != ''">
        <xsl:call-template name="SingleCol">
          <xsl:with-param name="result">
            <xsl:value-of select="$result"/>
          </xsl:with-param>
          <xsl:with-param name="sqref">
            <xsl:value-of select="$sqref"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$result"/>
      </xsl:otherwise>
    </xsl:choose>
    
    
  </xsl:template>
  
  <xsl:template name="SingleCol">
    <xsl:param name="result"/>
    <xsl:param name="sqref"/>
    
    <xsl:choose>
      <xsl:when test="contains($sqref, ':')">
        
        <xsl:variable name="StartColNum">
          <xsl:call-template name="GetColNum">
            <xsl:with-param name="cell">
              <xsl:value-of select="substring-before($sqref, ':')"/>
             </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>        
        
        
        <xsl:variable name="EndColNum">
          <xsl:call-template name="GetColNum">
            <xsl:with-param name="cell">
              <xsl:value-of select="substring-after($sqref, ':')"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        
        <xsl:choose>
          <xsl:when test="$StartColNum = $EndColNum and $StartColNum &lt; 257">            
           
            <xsl:choose>
              <xsl:when test="not(contains(concat(':', $result), concat(':', $StartColNum, ':')))">           
                <xsl:value-of select="concat($result, $StartColNum, ':')"/>    
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$result"/>
              </xsl:otherwise>
            </xsl:choose>
            
          </xsl:when>
          <xsl:otherwise>
         
            <xsl:call-template name="RepeatCol">
              <xsl:with-param name="result">
                <xsl:value-of select="$result"/>
              </xsl:with-param>
              <xsl:with-param name="start">
                <xsl:value-of select="$StartColNum"/>
              </xsl:with-param>
              <xsl:with-param name="end">
                <xsl:value-of select="$EndColNum"/>
              </xsl:with-param>
            </xsl:call-template>
            
          </xsl:otherwise>
        </xsl:choose>
        
      </xsl:when>
      
      <xsl:otherwise>
        
        <xsl:variable name="ColNr">
        <xsl:call-template name="GetColNum">
          <xsl:with-param name="cell">
            <xsl:value-of select="$sqref"/>
          </xsl:with-param>
        </xsl:call-template>
        </xsl:variable>
        
        <xsl:choose>
          <xsl:when test="not(contains(concat(':', $result), concat(':', $ColNr, ':'))) and $ColNr &lt; 257">
            <xsl:value-of select="concat($result, $ColNr, ':')"/>    
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$result"/>
          </xsl:otherwise>
        </xsl:choose>
        
        
        
      </xsl:otherwise>
      
    </xsl:choose>
    
  </xsl:template>
  
  <xsl:template name="RepeatCol">
    <xsl:param name="result"/>
    <xsl:param name="start"/>
    <xsl:param name="end"/>
    
    <xsl:choose>
      <xsl:when test="$start=$end and $start &lt; 257">        
        <xsl:value-of select="concat($result, $start, ':')"/>    
      </xsl:when>
      
      <xsl:when test="$start &lt; $end and $start &lt; 257">  
        
        <xsl:call-template name="RepeatCol">
          <xsl:with-param name="result">
            
            <xsl:choose>
              <xsl:when test="not(contains(concat(':', $result), concat(':', $start, ':'))) and $start &lt; 257"> 
                <xsl:value-of select="concat($result, $start, ':')"/>    
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$result"/>
              </xsl:otherwise>
            </xsl:choose>            
            
          </xsl:with-param>
          <xsl:with-param name="start">
            <xsl:value-of select="$start+1"/>
          </xsl:with-param>
          <xsl:with-param name="end">
            <xsl:value-of select="$end"/>
          </xsl:with-param>
        </xsl:call-template>
        
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$result"/>
      </xsl:otherwise>
    </xsl:choose>
    
  </xsl:template>



  <xsl:template name="ConditionalCellAll">
    <xsl:param name="sheet"/>
    <xsl:param name="document"/>
    
    <xsl:apply-templates select="e:worksheet/e:conditionalFormatting[1]" mode="All">
      <xsl:with-param name="document">
        <xsl:value-of select="$document"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalBefore">
        <xsl:text>0</xsl:text>
      </xsl:with-param>
    </xsl:apply-templates>
    
  </xsl:template>
  
  <xsl:template match="e:conditionalFormatting" mode="All">
    <xsl:param name="ConditionalCellAll"/>
    <xsl:param name="document"/>
    <xsl:param name="ConditionalBefore"/>
    
    
    <xsl:choose>
      
      <xsl:when test="following-sibling::e:conditionalFormatting">
        
        <xsl:apply-templates select="following-sibling::e:conditionalFormatting[1]" mode="All">
          <xsl:with-param name="ConditionalCellAll">
            <xsl:choose>
              <xsl:when test="$ConditionalCellAll != '' and @sqref != ''">
                <xsl:value-of select="concat($ConditionalCellAll, ' ', @sqref, ' -', $ConditionalBefore)"/>    
              </xsl:when>
              <xsl:when test="@sqref != ''">
                <xsl:value-of select="concat(@sqref, ' -', $ConditionalBefore)"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$ConditionalCellAll"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="document">
            <xsl:value-of select="$document"/>
          </xsl:with-param>
          <xsl:with-param name="ConditionalBefore">
            <xsl:value-of select="$ConditionalBefore + 1"/>
          </xsl:with-param>
        </xsl:apply-templates>
        
      </xsl:when>
      
      <xsl:otherwise>
        <xsl:choose>
        <xsl:when test="$ConditionalCellAll != '' and @sqref != ''">
          <xsl:value-of select="concat($ConditionalCellAll, ' ', @sqref, ' -', $ConditionalBefore)"/>    
        </xsl:when>
        <xsl:when test="@sqref != ''">
          <xsl:value-of select="concat(@sqref, ' -', $ConditionalBefore)"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$ConditionalCellAll"/>
        </xsl:otherwise>
      </xsl:choose>        
      </xsl:otherwise>
      
    </xsl:choose>
    
  </xsl:template>
  
  <xsl:template name="ConditionalCellSingle">
    <xsl:param name="result"/>
    <xsl:param name="sqref"/>
    
    <xsl:variable name="ConditionalBefore">
      <xsl:choose>
        <xsl:when test="substring-before(substring-after($sqref, '-'), ' ') != ''">
          <xsl:value-of select="substring-before(substring-after($sqref, '-'), ' ')"/>     
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="substring-after($sqref, '-')"/>
        </xsl:otherwise>
      </xsl:choose>
      </xsl:variable>
    
    <xsl:choose>
      <xsl:when test="substring-before($sqref, ' ') != ''">
                  <xsl:call-template name="ConditionalCellSingle">
                    <xsl:with-param name="result">
                      <xsl:choose>
                        <xsl:when test="not(contains(substring-before($sqref, ' '), ':')) and not(contains(substring-before($sqref, ' '), '-'))">
                          <xsl:value-of select="concat(substring-before($sqref, ' '), '-', $ConditionalBefore, ':', $result)"/>    
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$result"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:with-param>
                    <xsl:with-param name="sqref">
                      <xsl:value-of select="substring-after($sqref, ' ')"/>
                    </xsl:with-param>
                  </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="not(contains($sqref, ':')) and $sqref != ''  and not(contains($sqref, '-'))">
            <xsl:value-of select="concat($sqref, ':', $result)"/>    
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$result"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
    
    
  </xsl:template>
  
  <xsl:template name="ConditionalCellMultiple">
    <xsl:param name="result"/>
    <xsl:param name="sqref"/>
    
    <xsl:variable name="ConditionalBefore">
      <xsl:choose>
        <xsl:when test="substring-before(substring-after($sqref, '-'), ' ') != ''">
          <xsl:value-of select="substring-before(substring-after($sqref, '-'), ' ')"/>     
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="substring-after($sqref, '-')"/>
        </xsl:otherwise>
      </xsl:choose>
      
     
    </xsl:variable>
    
    <xsl:choose>
      <xsl:when test="substring-before($sqref, ' ') != ''">
        <xsl:call-template name="ConditionalCellMultiple">
          <xsl:with-param name="result">
            <xsl:choose>
              <xsl:when test="contains(substring-before($sqref, ' '), ':')  and not(contains(substring-before($sqref, ' '), '-'))">
                <xsl:value-of select="concat(substring-before($sqref, ' '),  '-', $ConditionalBefore,';', $result)"/>    
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$result"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="sqref">
            <xsl:value-of select="substring-after($sqref, ' ')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="contains($sqref, ':') and $sqref != ''  and not(contains(substring-before($sqref, ' '), '-'))">
            <xsl:value-of select="concat($sqref,  '-', $ConditionalBefore, ';', $result)"/>    
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$result"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
    
    
  </xsl:template>
  
  <xsl:template name="CheckIfConditionalInThisCell">
    <xsl:param name="rowNum"/>
    <xsl:param name="colNum"/>
    <xsl:param name="ConditionalCellMultiple"/>

  <xsl:choose>
    <xsl:when test="$ConditionalCellMultiple != ''">
      
      
      
      <xsl:variable name="CellMultiple">
        <xsl:value-of select="substring-before($ConditionalCellMultiple, ';')"/>
      </xsl:variable>
      
      <xsl:variable name="StartColNum">
        <xsl:call-template name="GetColNum">
          <xsl:with-param name="cell">
            <xsl:value-of select="substring-before($CellMultiple, ':')"/>
          </xsl:with-param>
        </xsl:call-template>        
      </xsl:variable>
      
      <xsl:variable name="StartRowNum">
        <xsl:call-template name="GetRowNum">
          <xsl:with-param name="cell">
            <xsl:value-of select="substring-before($CellMultiple, ':')"/>
          </xsl:with-param>
        </xsl:call-template>    
      </xsl:variable>
      
      <xsl:variable name="EndColNum">
        <xsl:call-template name="GetColNum">
          <xsl:with-param name="cell">
            <xsl:value-of select="substring-before(substring-after($CellMultiple, ':'), '-')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      
      <xsl:variable name="EndRowNum">
        <xsl:call-template name="GetRowNum">
          <xsl:with-param name="cell">
            <xsl:value-of select="substring-before(substring-after($CellMultiple, ':'), '-')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      
      <!--xsl:value-of select="concat($StartColNum, '-', $colNum, '-', $StartRowNum, '-', $rowNum, '-', $colNum, '-', $EndColNum, '-', $rowNum, '-', $EndRowNum)"/-->
      
      <xsl:choose>
        <xsl:when test="($StartColNum - 1) &lt; $colNum and ($StartRowNum - 1) &lt; $rowNum and $colNum &lt; ($EndColNum + 1) and $rowNum &lt; ($EndRowNum + 1)">
          <xsl:value-of select="$CellMultiple"/>
        </xsl:when>
        <xsl:otherwise>
        
          <xsl:call-template name="CheckIfConditionalInThisCell">
            <xsl:with-param name="rowNum">
              <xsl:value-of select="$rowNum"/>
            </xsl:with-param>
            <xsl:with-param name="colNum">
              <xsl:value-of select="$colNum"/>
            </xsl:with-param>
            <xsl:with-param name="ConditionalCellMultiple">
              <xsl:value-of select="substring-after($ConditionalCellMultiple, ';')"/>
            </xsl:with-param>
          </xsl:call-template>
          
        </xsl:otherwise>
      </xsl:choose>
      
    </xsl:when>
    
    <xsl:otherwise>
      <xsl:text>false</xsl:text>
    </xsl:otherwise>
    
  </xsl:choose>
  
  </xsl:template>
  
  <xsl:template name="GetMinRowWithConditional">
    <xsl:param name="ConditionalCellSingle"/>
    <xsl:param name="ConditionalCellMultiple"/>
    <xsl:param name="AfterRow"/>
    
    <xsl:call-template name="GetMinRowWithConditionalSingle">
      <xsl:with-param name="ConditionalCellSingle">
        <xsl:value-of select="$ConditionalCellSingle"/>
      </xsl:with-param>
    </xsl:call-template>
    
  </xsl:template>
  
  <xsl:template name="GetMinRowWithConditionalSingle">
    <xsl:param name="ConditionalCellSingle"/>
    <xsl:param name="AfterRow"/>
    
  </xsl:template>

</xsl:stylesheet>

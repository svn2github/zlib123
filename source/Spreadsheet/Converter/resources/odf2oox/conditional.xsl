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
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" exclude-result-prefixes="table r">


  <!-- search coditional -->
  <xsl:template match="table:table-row" mode="conditional">
    <xsl:param name="rowNumber"/>
    <xsl:param name="cellNumber"/>
    <xsl:param name="TableColumnTagNum"/>


    <xsl:apply-templates select="table:table-cell[1]" mode="conditional">
      <xsl:with-param name="colNumber">
        <xsl:text>0</xsl:text>
      </xsl:with-param>
      <xsl:with-param name="rowNumber" select="$rowNumber"/>
      <xsl:with-param name="TableColumnTagNum" select="$TableColumnTagNum"/>
    </xsl:apply-templates>

    <!-- check next row -->
    <xsl:choose>
      <!-- next row is a sibling -->
      <xsl:when test="following-sibling::node()[1][name() = 'table:table-row' ]">
        <xsl:apply-templates select="following-sibling::table:table-row[1]" mode="conditional">
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
            <xsl:text>0</xsl:text>
          </xsl:with-param>
          <xsl:with-param name="TableColumnTagNum" select="$TableColumnTagNum"/>
        </xsl:apply-templates>
      </xsl:when>
      <!-- next row is inside header rows -->
      <xsl:when test="following-sibling::node()[1][name() = 'table:table-header-rows' ]">
        <xsl:apply-templates select="following-sibling::table:table-header-rows/table:table-row[1]"
          mode="conditional">
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
            <xsl:text>0</xsl:text>
          </xsl:with-param>
          <xsl:with-param name="TableColumnTagNum" select="$TableColumnTagNum"/>
        </xsl:apply-templates>
      </xsl:when>
      <!-- this is last row inside header rows, next row is outside -->
      <xsl:when
        test="parent::node()[name()='table:table-header-rows'] and not(following-sibling::node()[1][name() = 'table:table-row' ])">
        <xsl:apply-templates select="parent::node()/following-sibling::table:table-row[1]"
          mode="conditional">
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
            <xsl:text>0</xsl:text>
          </xsl:with-param>
          <xsl:with-param name="TableColumnTagNum" select="$TableColumnTagNum"/>
        </xsl:apply-templates>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- insert coditional -->
  <xsl:template match="table:table-cell|table:covered-table-cell" mode="conditional">
    <xsl:param name="colNumber"/>
    <xsl:param name="rowNumber"/>
    <xsl:param name="TableColumnTagNum"/>

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

    <xsl:variable name="styleName">
      <xsl:choose>
        <xsl:when test="@table:style-name != '' ">
          <xsl:value-of select="@table:style-name"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$columnCellStyle"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:if test="key('style',$styleName)/style:map/@style:condition != '' ">
      <conditionalFormatting>
        <xsl:variable name="ColChar">
          <xsl:call-template name="NumbersToChars">
            <xsl:with-param name="num">
              <xsl:value-of select="$colNumber"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:attribute name="sqref">
          <xsl:value-of select="concat($ColChar, $rowNumber)"/>
        </xsl:attribute>
        <xsl:for-each select="key('style', $styleName)/style:map">
          <cfRule type="cellIs" priority="{position()}">
            <xsl:if test="contains(@style:condition,'is-true-formula')">
              <xsl:attribute name="type">
                <xsl:text>expression</xsl:text>
              </xsl:attribute>
            </xsl:if>
            <xsl:attribute name="dxfId">
              <xsl:value-of select="count(preceding::style:map) + 1"/>
            </xsl:attribute>
            <xsl:if test="not(contains(@style:condition,'is-true-formula'))">
              <xsl:attribute name="operator">
                <xsl:choose>
                  <xsl:when test="contains(@style:condition, '&lt;=')">
                    <xsl:text>lessThanOrEqual</xsl:text>
                  </xsl:when>
                  <xsl:when test="contains(@style:condition, '&lt;')">
                    <xsl:text>lessThan</xsl:text>
                  </xsl:when>
                  <xsl:when test="contains(@style:condition, '&gt;=')">
                    <xsl:text>greaterThanOrEqual</xsl:text>
                  </xsl:when>
                  <xsl:when test="contains(@style:condition, '&gt;')">
                    <xsl:text>greaterThan</xsl:text>
                  </xsl:when>
                  <xsl:when test="contains(@style:condition, '!=')">
                    <xsl:text>notEqual</xsl:text>
                  </xsl:when>
                  <xsl:when test="contains(@style:condition, 'cell-content-is-between')">
                    <xsl:text>between</xsl:text>
                  </xsl:when>
                  <xsl:when test="contains(@style:condition, 'cell-content-is-not-between')">
                    <xsl:text>notBetween</xsl:text>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:text>equal</xsl:text>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
            <xsl:call-template name="InsertCoditionalFormula"/>
          </cfRule>
        </xsl:for-each>
      </conditionalFormatting>
    </xsl:if>

    <xsl:if test="following-sibling::table:table-cell">
      <xsl:apply-templates
        select="following-sibling::table:table-cell[1]|following-sibling::table:covered-table-cell[1]"
        mode="conditional">
        <xsl:with-param name="colNumber">
          <xsl:choose>
            <xsl:when test="@table:number-columns-repeated != ''">
              <xsl:value-of select="number($colNumber) + number(@table:number-columns-repeated)"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$colNumber + 1"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
        <xsl:with-param name="rowNumber">
          <xsl:value-of select="$rowNumber"/>
        </xsl:with-param>
        <xsl:with-param name="TableColumnTagNum" select="$TableColumnTagNum"/>
      </xsl:apply-templates>
    </xsl:if>

  </xsl:template>

  <!-- Coditional Format -->

  <xsl:template name="InsertConditionalFormat">
    <dxfs count="3">
      <dxf>
        <font>
          <condense val="0"/>
          <extend val="0"/>
          <color rgb="FF9C0006"/>
        </font>
        <fill>
          <patternFill>
            <bgColor rgb="FFFFC7CE"/>
          </patternFill>
        </fill>
      </dxf>
      <xsl:for-each
        select="document('content.xml')/office:document-content/office:automatic-styles/style:style/style:map[@style:condition != '']">
        <xsl:variable name="StyleApplyStyleName">
          <xsl:value-of select="@style:apply-style-name"/>
        </xsl:variable>
        <dxf>
          <xsl:for-each
            select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$StyleApplyStyleName]">
            <font>
              <xsl:call-template name="InsertTextProperties">
                <xsl:with-param name="mode">default</xsl:with-param>
              </xsl:call-template>
            </font>
          </xsl:for-each>
          <!-- style:table-cell-properties fo:background-color       -->
          <xsl:for-each
            select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$StyleApplyStyleName]">
            <xsl:apply-templates select="style:table-cell-properties" mode="background-color">
              <xsl:with-param name="Object">
                <xsl:text>conditional</xsl:text>
              </xsl:with-param>
            </xsl:apply-templates>
            <xsl:apply-templates select="style:table-cell-properties" mode="border"/>
          </xsl:for-each>

        </dxf>
      </xsl:for-each>

    </dxfs>
  </xsl:template>

  <!-- Insert Formula of Coditional -->

  <xsl:template name="InsertCoditionalFormula">
    <xsl:choose>
      <xsl:when test="contains(@style:condition, '.$#REF!$#REF!')">
        <formula/>
      </xsl:when>
      <xsl:when test="contains(@style:condition, '&lt;=')">
        <formula>
          <xsl:choose>
            <xsl:when test="contains(@style:condition, '[')">
              <xsl:value-of
                select="substring-after(substring-before(substring-after(@style:condition, '['), ']'), '.')"
              />
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-after(@style:condition, '=')"/>
            </xsl:otherwise>
          </xsl:choose>
        </formula>
      </xsl:when>
      <xsl:when test="contains(@style:condition, '&lt;')">
        <formula>
          <xsl:choose>
            <xsl:when test="contains(@style:condition, '[')">
              <xsl:value-of
                select="substring-after(substring-before(substring-after(@style:condition, '['), ']'), '.')"
              />
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-after(@style:condition, '&lt;')"/>
            </xsl:otherwise>
          </xsl:choose>
        </formula>
      </xsl:when>
      <xsl:when test="contains(@style:condition, '&gt;=')">
        <formula>
          <xsl:choose>
            <xsl:when test="contains(@style:condition, '[')">
              <xsl:value-of
                select="substring-after(substring-before(substring-after(@style:condition, '['), ']'), '.')"
              />
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-after(@style:condition, '=')"/>
            </xsl:otherwise>
          </xsl:choose>
        </formula>
      </xsl:when>
      <xsl:when test="contains(@style:condition, '&gt;')">
        <formula>
          <xsl:choose>
            <xsl:when test="contains(@style:condition, '[')">
              <xsl:value-of
                select="substring-after(substring-before(substring-after(@style:condition, '['), ']'), '.')"
              />
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-after(@style:condition, '&gt;')"/>
            </xsl:otherwise>
          </xsl:choose>
        </formula>
      </xsl:when>
      <xsl:when test="contains(@style:condition, '!=')">
        <formula>
          <xsl:choose>
            <xsl:when test="contains(@style:condition, '[')">
              <xsl:value-of
                select="substring-after(substring-before(substring-after(@style:condition, '['), ']'), '.')"
              />
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-after(@style:condition, '=')"/>
            </xsl:otherwise>
          </xsl:choose>
        </formula>
      </xsl:when>
      <xsl:when test="contains(@style:condition, 'cell-content-is-between')">
        <formula>
          <xsl:choose>
            <xsl:when test="contains(@style:condition, '[')">
              <xsl:value-of
                select="substring-after(substring-before(substring-after(@style:condition, '['), ']'), '.')"
              />
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-before(substring-after(@style:condition, '('), ',')"/>
            </xsl:otherwise>
          </xsl:choose>
        </formula>
        <formula>
          <xsl:choose>
            <xsl:when test="contains(@style:condition, '[')">
              <xsl:value-of
                select="substring-after(substring-before(substring-after(@style:condition, '['), ']'), '.')"
              />
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-before(substring-after(@style:condition, ','), ')')"/>
            </xsl:otherwise>
          </xsl:choose>
        </formula>
      </xsl:when>
      <xsl:when test="contains(@style:condition, 'cell-content-is-not-between')">
        <formula>
          <xsl:choose>
            <xsl:when test="contains(@style:condition, '[')">
              <xsl:value-of
                select="substring-after(substring-before(substring-after(@style:condition, '['), ']'), '.')"
              />
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-before(substring-after(@style:condition, '('), ',')"/>
            </xsl:otherwise>
          </xsl:choose>
        </formula>
        <formula>
          <xsl:choose>
            <xsl:when test="contains(@style:condition, '[')">
              <xsl:value-of
                select="substring-after(substring-before(substring-after(@style:condition, '['), ']'), '.')"
              />              
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-before(substring-after(@style:condition, ','), ')')"/>
            </xsl:otherwise>
          </xsl:choose>
        </formula>
      </xsl:when>
      <xsl:when test="contains(@style:condition, '=')">
        <formula>
          <xsl:choose>
            <xsl:when test="contains(@style:condition, '[')">
              <xsl:value-of
                select="substring-after(substring-before(substring-after(@style:condition, '['), ']'), '.')"
              />
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-after(@style:condition, '=')"/>
            </xsl:otherwise>
          </xsl:choose>
        </formula>
      </xsl:when>
      <xsl:when test="contains(@style:condition, 'is-true-formula')">
        <formula>
          <!-- string inside is-true-formula(...) -->
          <xsl:value-of
            select="substring-after(substring(@style:condition,1,string-length(@style:condition) -1), 'is-true-formula(' )"
          />
        </formula>
      </xsl:when>
      <xsl:otherwise>
        <formula>
          <xsl:choose>
            <xsl:when test="contains(@style:condition, '[')">
              <xsl:value-of
                select="substring-after(substring-before(substring-after(@style:condition, '['), ']'), '.')"
              />
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-after(@style:condition, '=')"/>
            </xsl:otherwise>
          </xsl:choose>
        </formula>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>


</xsl:stylesheet>

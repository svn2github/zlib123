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
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:dc="http://purl.org/dc/elements/1.1/" exclude-result-prefixes="table r text style fo dc">

  <xsl:import href="measures.xsl"/>

  <!-- template which inserts sharedstringscontent -->
  <xsl:template name="InsertSharedStrings">
    <sst>
      <xsl:variable name="Count">
        <!-- string can be in a simple row or a header row -->
        <xsl:value-of
          select="count(document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table/table:table-row/table:table-cell[text:p and not(@office:value-type='float') and (@office:value-type='string' or @office:value-type='boolean' or not((number(text:p) or text:p = 0 or contains(text:p,',') or contains(text:p,'%') or @office:value-type='currency' or @office:value-type='date' or @office:value-type='time')))] |
          document('content.xml')/office:document-content/office:body/office:spreadsheet/table:table/table:table-header-rows/table:table-row/table:table-cell[text:p and not(@office:value-type='float') and (@office:value-type='string' or @office:value-type='boolean' or not((number(text:p) or text:p = 0 or contains(text:p,',') or contains(text:p,'%') or @office:value-type='currency' or @office:value-type='date' or @office:value-type='time')))])"
        />
      </xsl:variable>
      <!-- possibly unnecessary attribute-->
      <xsl:attribute name="count">
        <xsl:value-of select="$Count"/>
      </xsl:attribute>
      <xsl:attribute name="uniqueCount">
        <xsl:value-of select="$Count"/>
      </xsl:attribute>
      <xsl:call-template name="InsertString"/>
    </sst>
  </xsl:template>

  <!-- template which inserts a string into sharedstrings -->
  <xsl:template name="InsertString">
    <!-- string can be in a simple row or a header row -->
    <!--this is shortened version for grouping-->
    <xsl:for-each
      select="document('content.xml')/office:document-content/office:body/office:spreadsheet/descendant::table:table-cell[text:p and not(@office:value-type='float') and (@office:value-type='string' or @office:value-type='boolean' or not((number(text:p) or text:p = 0 or contains(text:p,',') or contains(text:p,'%') or @office:value-type='currency' or @office:value-type='date' or @office:value-type='time')))]
      ">
      <si>
        <xsl:choose>
          <xsl:when test="text:span|text:p/text:span">
            <xsl:apply-templates mode="run" select="text:p"/>
          </xsl:when>
          <xsl:otherwise>
            <t xml:space="preserve"><xsl:apply-templates mode="text" select="text:p"/></t>
          </xsl:otherwise>
        </xsl:choose>
      </si>
    </xsl:for-each>
  </xsl:template>

  <!-- text:span conversion -->
  <xsl:template match="text:span" mode="run">
    <xsl:variable name="tekst">
      <xsl:value-of select="."/>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$tekst != '' or text:s">
        <r>
          <xsl:apply-templates select="key('style',@text:style-name)" mode="textstyles">
            <xsl:with-param name="parentCellStyleName">
              <xsl:value-of select="ancestor::table:table-cell/@table:style-name"/>
            </xsl:with-param>
            <xsl:with-param name="defaultCellStyleName">
              <xsl:value-of select="ancestor::table:table-column/@table:default-cell-style-name"/>
            </xsl:with-param>
          </xsl:apply-templates>
          <t xml:space="preserve"><xsl:apply-templates mode="text"/></t>
        </r>
      </xsl:when>
      <xsl:otherwise/>
    </xsl:choose>
  </xsl:template>

  <!-- when there is formatted text in a string, all texts must be in runs -->
  <xsl:template match="text()" mode="run">
    <r>
      <xsl:apply-templates select="key('style',ancestor::table:table-cell/@table:style-name)"
        mode="textstyles">
        <xsl:with-param name="defaultCellStyleName">
          <xsl:value-of select="ancestor::table:table-column/@table:default-cell-style-name"/>
        </xsl:with-param>
      </xsl:apply-templates>
      <xsl:variable name="value">
        <xsl:value-of select="."/>
      </xsl:variable>
      <!-- caution with 'Enters' because they can result with additional space in output text -->
      <t xml:space="preserve"><xsl:choose>
        <xsl:when test="not(contains($value, '_x'))"><xsl:value-of select="$value"/></xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="HexaDecimalValue">
            <xsl:with-param name="value">
              <xsl:value-of select="$value"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose></t>
    </r>
  </xsl:template>

  <xsl:template match="text()[parent::dc:date]" mode="text"/>
  <xsl:template match="text()[parent::dc:date]" mode="run"/>

  <xsl:template match="text()" mode="text">
    <xsl:variable name="value">
      <xsl:value-of select="."/>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="not(contains($value, '_x'))">
        <xsl:value-of select="$value"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="HexaDecimalValue">
          <xsl:with-param name="value">
            <xsl:value-of select="$value"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="text:s" mode="text">
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

  <!-- when there are more than one line of text, enter must be added -->
  <xsl:template match="text:p[preceding-sibling::text:p]" mode="run">
    <r>
      <!-- set text formating of first span -->
      <xsl:for-each select="text:span[1]">
        <xsl:apply-templates select="key('style',@text:style-name)" mode="textstyles">
          <xsl:with-param name="parentCellStyleName">
            <xsl:value-of select="ancestor::table:table-cell/@table:style-name"/>
          </xsl:with-param>
          <xsl:with-param name="defaultCellStyleName">
            <xsl:value-of select="ancestor::table:table-column/@table:default-cell-style-name"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:for-each>
      <t xml:space="preserve"><xsl:value-of select="'&#xD;'"/></t>
    </r>
    <xsl:if test="text() or text:span/text()">
      <xsl:apply-templates mode="run"/>
    </xsl:if>
  </xsl:template>

  <xsl:template match="text:p[preceding-sibling::text:p]" mode="text">
    <xsl:value-of select="'&#xD;'"/>
    <xsl:apply-templates mode="text"/>
  </xsl:template>

  <!-- when there are HaxaDecimal value (_x...._), must be added _x005F -->
  <xsl:template name="HexaDecimalValue">
    <xsl:param name="value"/>
    <xsl:param name="result"/>
    <xsl:choose>
      <xsl:when test="substring(substring-after($value, '_x'), 5, 1) = '_'">
        <xsl:variable name="CheckIfHexadecimal">
          <xsl:call-template name="CheckIfHexadecimal">
            <xsl:with-param name="value">
              <xsl:value-of select="substring-before(substring-after($value, '_x'), '_')"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="$CheckIfHexadecimal = 'true'">
            <xsl:call-template name="HexaDecimalValue">
              <xsl:with-param name="value">
                <xsl:value-of select="substring-after(substring-after($value, '_x'), '_')"/>
              </xsl:with-param>
              <xsl:with-param name="result">
                <xsl:value-of
                  select="concat($result, concat(concat(concat(substring-before($value, '_x'), '_x005F_x'), substring-before(substring-after($value, '_x'), '_')), '_'))"
                />
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="HexaDecimalValue">
              <xsl:with-param name="value">
                <xsl:value-of select="substring-after(substring-after($value, '_x'), '_')"/>
              </xsl:with-param>
              <xsl:with-param name="result">
                <xsl:value-of
                  select="concat($result, substring-before($value, substring-after(substring-after($value, '_x'), '_')))"
                />
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="concat($result, $value)"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

</xsl:stylesheet>

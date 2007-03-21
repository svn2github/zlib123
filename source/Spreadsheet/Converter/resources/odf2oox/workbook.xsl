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

  <xsl:key name="ConfigItem" match="config:config-item" use="@config:name"/>
  <xsl:key name="style" match="style:style" use="@style:name"/>

  <!-- main workbook template-->
  <xsl:template name="InsertWorkbook">
    <xsl:apply-templates select="document('content.xml')/office:document-content"/>
  </xsl:template>

  <!-- workbook body template -->
  <xsl:template match="office:body">
    <workbook>
      <xsl:call-template name="workbookView"/>
      <xsl:apply-templates select="office:spreadsheet"/>
    </workbook>
  </xsl:template>

  <!-- workbook  view template-->
  <xsl:template name="workbookView">
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
          select="concat(substring(translate(@table:name,&quot;*\/[]:&apos;?&quot;,&quot;&quot;),1,31 - 2 - string-length($nameConflictsBefore + 1)),'(',$nameConflictsBefore + 1,')')"
        />
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$name"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

</xsl:stylesheet>

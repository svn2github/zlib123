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
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:v="urn:schemas-microsoft-com:vml" xmlns:xlink="http://www.w3.org/1999/xlink"
  exclude-result-prefixes="a e r number">

  <xsl:import href="relationships.xsl"/>
  <xsl:import href="border.xsl"/>
  <xsl:import href="headers.xsl"/>
  <xsl:import href="insert_cols.xsl"/>
  <xsl:import href="note.xsl"/>

  <xsl:key name="Border" match="e:borders" use="''"/>
  <xsl:key name="CellStylesId" match="e:cellStyle" use="@xfId"/>

  <xsl:template name="styles">

    <xsl:variable name="activeTab">
      <xsl:for-each select="document('xl/workbook.xml')">
        <xsl:choose>
          <xsl:when test="e:workbook/e:bookViews/e:workbookView/@activeTab">
            <xsl:value-of select="e:workbook/e:bookViews/e:workbookView/@activeTab"/>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
    </xsl:variable>

    <office:document-styles>
      <office:font-face-decls>
        <xsl:call-template name="InsertFonts"/>
        <xsl:call-template name="InsertHeaderFooterFonts">
          <xsl:with-param name="activeTab" select="$activeTab"/>
        </xsl:call-template>
      </office:font-face-decls>
      <office:styles>
        <xsl:call-template name="InsertDefaultTableCellStyle"/>
        <xsl:call-template name="InsertCellStyle"/>
        <!-- Insert Conditional Styles-->
        <xsl:call-template name="InsertConditionalStyles"/>
      </office:styles>
      <office:automatic-styles>
        <xsl:call-template name="InsertPageLayout">
          <xsl:with-param name="activeTab" select="$activeTab"/>
        </xsl:call-template>
        <xsl:call-template name="InsertHeaderFooterStyles">
          <xsl:with-param name="activeTab" select="$activeTab"/>
        </xsl:call-template>
      </office:automatic-styles>
      <office:master-styles>
        <style:master-page style:name="Default" style:page-layout-name="pm1">
          <xsl:call-template name="InsertHeaderFooter">
            <xsl:with-param name="activeTab" select="$activeTab"/>
          </xsl:call-template>
        </style:master-page>
      </office:master-styles>
    </office:document-styles>
  </xsl:template>

  <xsl:template name="InsertFonts">
    <xsl:for-each select="document('xl/styles.xml')/e:styleSheet/e:fonts/e:font[e:name]">
      <style:font-face style:name="{e:name/@val}" svg:font-family="{e:name/@val}"/>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="InsertDefaultTableCellStyle">
    <style:style style:name="Default" style:family="table-cell">
      <style:text-properties>
        <xsl:apply-templates select="document('xl/styles.xml')/e:styleSheet/e:fonts/e:font[1]"
          mode="style"/>
      </style:text-properties>
    </style:style>
  </xsl:template>

  <!-- insert column styles from all sheets -->
  <xsl:template name="InsertColumnStyles">
    <xsl:for-each select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet">
      <xsl:call-template name="InsertSheetColumnStyles">
        <xsl:with-param name="sheet">
          <xsl:call-template name="GetTarget">
            <xsl:with-param name="id">
              <xsl:value-of select="@r:id"/>
            </xsl:with-param>
            <xsl:with-param name="document">xl/workbook.xml</xsl:with-param>
          </xsl:call-template>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:for-each>
  </xsl:template>

  <!-- insert column styles from selected sheet -->
  <xsl:template name="InsertSheetColumnStyles">
    <xsl:param name="sheet"/>

    <xsl:variable name="ManualColBreaks">
      <xsl:for-each select="document(concat('xl/',$sheet))/e:worksheet/e:colBreaks/e:brk">
        <xsl:value-of select="concat(@id,';')"/>
      </xsl:for-each>
    </xsl:variable>

    <!-- default & column Breake style -->
    <xsl:choose>
      <xsl:when test="document(concat('xl/',$sheet))/e:worksheet/e:colBreaks">
        <style:style
          style:name="{generate-id(document(concat('xl/',$sheet))/e:worksheet/e:colBreaks)}"
          style:family="table-column">
          <style:table-column-properties fo:break-before="page">
            <xsl:attribute name="style:column-width">
              <xsl:choose>
                <xsl:when
                  test="document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr/@defaultColWidth">
                  <xsl:call-template name="ConvertFromCharacters">
                    <xsl:with-param name="value">
                      <xsl:value-of
                        select="document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr/@defaultColWidth"
                      />
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <!-- Excel application default-->
                  <xsl:call-template name="ConvertFromCharacters">
                    <xsl:with-param name="value" select="'8.43'"/>
                  </xsl:call-template>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </style:table-column-properties>
        </style:style>
      </xsl:when>
      <xsl:otherwise>
        <style:style
          style:name="{generate-id(document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr)}"
          style:family="table-column">
          <style:table-column-properties fo:break-before="auto">
            <xsl:attribute name="style:column-width">
              <xsl:choose>
                <xsl:when
                  test="document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr/@defaultColWidth">
                  <xsl:call-template name="ConvertFromCharacters">
                    <xsl:with-param name="value">
                      <xsl:value-of
                        select="document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr/@defaultColWidth"
                      />
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <!-- Excel application default-->
                  <xsl:call-template name="ConvertFromCharacters">
                    <xsl:with-param name="value" select="'8.43'"/>
                  </xsl:call-template>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </style:table-column-properties>
        </style:style>
      </xsl:otherwise>
    </xsl:choose>
    
    <xsl:apply-templates select="document(concat('xl/',$sheet))/e:worksheet/e:cols/e:col[1]"
      mode="automaticstyles">
      <xsl:with-param name="manualBreakes" select="$ManualColBreaks"/>
    </xsl:apply-templates>
  </xsl:template>

  <xsl:template name="colBreakId">
    <xsl:variable name="sheet">
      <xsl:call-template name="GetTarget">
        <xsl:with-param name="id">
          <xsl:value-of select="@r:id"/>
        </xsl:with-param>
        <xsl:with-param name="document">xl/workbook.xml</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:attribute name="table:style-name">
      <xsl:value-of select="generate-id(document(concat('xl/',$sheet))/e:worksheet/e:colBreaks)"/>
    </xsl:attribute>
  </xsl:template>

  <xsl:template match="e:col" mode="automaticstyles">
    <xsl:param name="manualBreakes"/>

    <xsl:variable name="breakes">
      <xsl:choose>
        <xsl:when test="substring-before($manualBreakes,';') &lt;= @max">
          <xsl:call-template name="CutBreaks">
            <xsl:with-param name="breakes" select="$manualBreakes"/>
            <xsl:with-param name="max" select="@max"/>
            <xsl:with-param name="mode">
              <xsl:text>style</xsl:text>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$manualBreakes"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <!-- insert column style with break -->
    <xsl:if test="substring-before($manualBreakes,';') &lt;= @max">
      <style:style style:name="{concat(generate-id(.),'-break')}" style:family="table-column">
        <style:table-column-properties fo:break-before="page">
          <xsl:if test="@width">
            <xsl:attribute name="style:column-width">
              <xsl:call-template name="ConvertFromCharacters">
                <xsl:with-param name="value" select="@width"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
        </style:table-column-properties>
      </style:style>
    </xsl:if>

    <!-- TO DO: do not create column style without break when this is e:col for single column and it has break -->
    <!--xsl:if test="not(@min != @max and substring-before($manualBreakes,';') &lt;= @max)"-->
      <style:style style:name="{generate-id(.)}" style:family="table-column">
        <style:table-column-properties>
          <xsl:if test="@width">
            <xsl:attribute name="style:column-width">
              <xsl:call-template name="ConvertFromCharacters">
                <xsl:with-param name="value" select="@width"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
        </style:table-column-properties>
      </style:style>
    <!--/xsl:if-->

    <xsl:if test="following-sibling::e:col">
      <xsl:apply-templates select="following-sibling::e:col[1]" mode="automaticstyles">
        <xsl:with-param name="manualBreakes" select="$breakes"/>
      </xsl:apply-templates>
    </xsl:if>

  </xsl:template>

  <!-- insert row styles from all sheets -->
  <xsl:template name="InsertRowStyles">
    <xsl:for-each select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet">
      <xsl:call-template name="InsertSheetRowStyles">
        <xsl:with-param name="sheet">
          <xsl:call-template name="GetTarget">
            <xsl:with-param name="id">
              <xsl:value-of select="@r:id"/>
            </xsl:with-param>
            <xsl:with-param name="document">xl/workbook.xml</xsl:with-param>
          </xsl:call-template>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:for-each>
  </xsl:template>

  <!-- insert row styles from selected sheet -->
  <xsl:template name="InsertSheetRowStyles">
    <xsl:param name="sheet"/>

    <!-- default & row Break style -->
    
    <xsl:choose>
      <xsl:when test="document(concat('xl/',$sheet))/e:worksheet/e:rowBreaks">
        <style:style
          style:name="{generate-id(document(concat('xl/',$sheet))/e:worksheet/e:rowBreaks)}"
          style:family="table-row">
          <style:table-row-properties fo:break-before="page">
            <xsl:attribute name="style:row-height">
              <xsl:choose>
                <xsl:when
                  test="document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr/@defaultRowHeight">
                  <xsl:call-template name="ConvertToCentimeters">
                    <xsl:with-param name="length">
                      <xsl:value-of
                        select="concat(document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr/@defaultRowHeight,'pt')"
                      />
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <!-- Excel application default-->
                  <xsl:call-template name="ConvertToCentimeters">
                    <xsl:with-param name="length" select="'20px'"/>
                  </xsl:call-template>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </style:table-row-properties>
        </style:style>
      </xsl:when>
      <xsl:otherwise>
        <style:style
          style:name="{generate-id(document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr)}"
          style:family="table-row">
          <style:table-row-properties fo:break-before="auto">
            <xsl:attribute name="style:row-height">
              <xsl:choose>
                <xsl:when
                  test="document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr/@defaultRowHeight">
                  <xsl:call-template name="ConvertToCentimeters">
                    <xsl:with-param name="length">
                      <xsl:value-of
                        select="concat(document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr/@defaultRowHeight,'pt')"
                      />
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <!-- Excel application default-->
                  <xsl:call-template name="ConvertToCentimeters">
                    <xsl:with-param name="length" select="'20px'"/>
                  </xsl:call-template>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </style:table-row-properties>
        </style:style>
      </xsl:otherwise>
    </xsl:choose>
    
    <xsl:apply-templates select="document(concat('xl/',$sheet))/e:worksheet/e:sheetData"
      mode="automaticstyles"/>
  </xsl:template>

  <xsl:template match="e:row" mode="automaticstyles">
    <xsl:if test="@ht">
      <style:style style:name="{generate-id(.)}" style:family="table-row">
        <style:table-row-properties fo:break-before="auto">
          <xsl:attribute name="style:row-height">
            <xsl:call-template name="ConvertToCentimeters">
              <xsl:with-param name="length" select="concat(@ht,'pt')"/>
            </xsl:call-template>
          </xsl:attribute>
          <xsl:attribute name="style:use-optimal-row-height">
            <xsl:choose>
              <xsl:when test="@customHeight">
                <xsl:text>false</xsl:text>
              </xsl:when>
              <xsl:when test="@thickTop">
                <xsl:text>false</xsl:text>
              </xsl:when>
              <xsl:when test="@thickBot">
                <xsl:text>false</xsl:text>
              </xsl:when>
              <xsl:otherwise>
                <xsl:text>true</xsl:text>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </style:table-row-properties>
      </style:style>
    </xsl:if>
  </xsl:template>

  <!--  Insert Table Properties -->
  <xsl:template name="InsertStyleTableProperties">
    <xsl:for-each select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet">
      <style:style>
        <xsl:attribute name="style:name">
          <xsl:value-of select="generate-id()"/>
        </xsl:attribute>
        <xsl:attribute name="style:family">
          <xsl:text>table</xsl:text>
        </xsl:attribute>
        <xsl:attribute name="style:master-page-name">
          <xsl:text>Default</xsl:text>
        </xsl:attribute>
        <style:table-properties>
          <xsl:if test="@state='hidden'">
            <xsl:attribute name="table:display">
              <xsl:text>false</xsl:text>
            </xsl:attribute>
          </xsl:if>

          <xsl:variable name="sheet">
            <xsl:call-template name="GetTarget">
              <xsl:with-param name="id">
                <xsl:value-of select="@r:id"/>
              </xsl:with-param>
              <xsl:with-param name="document">xl/workbook.xml</xsl:with-param>
            </xsl:call-template>
          </xsl:variable>

          <xsl:for-each select="document(concat('xl/',$sheet))/e:worksheet/e:sheetViews/e:sheetView">
            <xsl:if test="@rightToLeft = 1">
              <xsl:attribute name="style:writing-mode">
                <xsl:value-of select="'rl-tb'"/>
              </xsl:attribute>
              <xsl:attribute name="sheet">
                <xsl:value-of select="$sheet"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>

        </style:table-properties>
      </style:style>
    </xsl:for-each>
  </xsl:template>

  <!-- insert number styles-->

  <xsl:template name="InsertNumberStyles">

    <xsl:apply-templates select="document('xl/styles.xml')/e:styleSheet/e:cellXfs"
      mode="fixedNumFormat"/>
    <xsl:apply-templates select="document('xl/styles.xml')/e:styleSheet/e:numFmts"
      mode="automaticstyles"/>
  </xsl:template>

  <xsl:template name="InsertCellStyles">
    <xsl:apply-templates select="document('xl/styles.xml')/e:styleSheet/e:cellXfs"
      mode="automaticstyles"/>
  </xsl:template>

  <!-- cell formats -->
  <xsl:template match="e:xf" mode="automaticstyles">

    <style:style style:name="{generate-id(.)}" style:family="table-cell">
      <xsl:call-template name="InsertCellFormat"/>
    </style:style>

  </xsl:template>

  <xsl:template name="InsertCellFormat">
    <xsl:choose>

      <!-- existing number format -->
      <xsl:when test="key('numFmtId',@numFmtId)">
        <xsl:attribute name="style:data-style-name">
          <xsl:value-of select="generate-id(key('numFmtId',@numFmtId))"/>
        </xsl:attribute>
      </xsl:when>

      <!-- fixed number format -->
      <xsl:when test="@numFmtId &gt; 0">
        <xsl:attribute name="style:data-style-name">
          <xsl:value-of select="concat('N',@numFmtId)"/>
        </xsl:attribute>
      </xsl:when>

    </xsl:choose>

    <xsl:attribute name="style:parent-style-name">
      <xsl:text>Default</xsl:text>
    </xsl:attribute>

    <xsl:if
      test="@applyAlignment = 1 or @applyBorder = 1 or (@applyProtection=1) or  @borderId != '0' or @fillId!='0' or @applyFill= 1">
      <style:table-cell-properties>
        <xsl:if test="@applyAlignment = 1">
          <!-- vertical-align -->
          <xsl:attribute name="style:vertical-align">
            <xsl:choose>
              <xsl:when test="e:alignment/@vertical = 'center' ">
                <xsl:text>middle</xsl:text>
              </xsl:when>
              <xsl:when test="not(e:alignment/@vertical)">
                <xsl:text>bottom</xsl:text>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="e:alignment/@vertical"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="e:alignment/@horizontal = 'fill' ">
          <xsl:attribute name="style:repeat-content">
            <xsl:text>true</xsl:text>
          </xsl:attribute>
        </xsl:if>
        <xsl:if
          test="e:alignment/@horizontal = 'distributed' or e:alignment/@vertical = 'distributed' ">
          <xsl:attribute name="fo:wrap-option">
            <xsl:text>wrap</xsl:text>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="e:alignment/@wrapText">
          <xsl:attribute name="fo:wrap-option">
            <xsl:text>wrap</xsl:text>
          </xsl:attribute>
        </xsl:if>

        <!-- text orientation -->
        <xsl:if test="e:alignment/@textRotation">
          <xsl:choose>
            <xsl:when test="e:alignment/@textRotation = 255">
              <xsl:attribute name="style:direction">
                <xsl:text>ttb</xsl:text>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="style:rotation-angle">
                <xsl:choose>
                  <xsl:when
                    test="e:alignment/@textRotation &lt; 90 or e:alignment/@textRotation = 90">
                    <xsl:value-of select="e:alignment/@textRotation"/>
                  </xsl:when>
                  <xsl:when test="e:alignment/@textRotation &gt; 90">
                    <xsl:value-of select="450 - e:alignment/@textRotation"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
          <xsl:attribute name="style:rotation-align">
            <xsl:text>none</xsl:text>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="@applyBorder = 1 or  @borderId != '0'">
          <xsl:call-template name="InsertBorder"/>
        </xsl:if>

        <xsl:if test="@applyProtection=1 and e:protection">
          <xsl:attribute name="style:cell-protect">
            <xsl:choose>
              <xsl:when test="e:protection/@locked=0 and e:protection/@hidden=1">
                <xsl:text>formula-hidden</xsl:text>
              </xsl:when>
              <xsl:when test="e:protection/@hidden= 1">
                <xsl:text>protected formula-hidden</xsl:text>
              </xsl:when>
              <xsl:otherwise>
                <xsl:text>protected</xsl:text>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>

        <xsl:if test="@applyFill=1 or @fillId!='0'">
          <xsl:variable name="this" select="."/>
          <xsl:apply-templates
            select="ancestor::e:styleSheet/e:fills/e:fill[position() = $this/@fillId + 1]"
            mode="style"/>
        </xsl:if>

      </style:table-cell-properties>

      <!-- default horizontal alignment when text has angle orientation  -->
      <xsl:if test="not(e:alignment/@horizontal) and e:alignment/@textRotation">
        <style:paragraph-properties>
          <xsl:attribute name="fo:text-align">
            <xsl:choose>
              <xsl:when
                test="e:alignment/@textRotation &lt; 90 or e:alignment/@textRotation = 180">
                <xsl:text>start</xsl:text>
              </xsl:when>
              <xsl:when
                test="e:alignment/@textRotation &gt; 89 and e:alignment/@textRotation &lt; 180">
                <xsl:text>end</xsl:text>
              </xsl:when>
              <xsl:when test="e:alignment/@textRotation = 255">
                <xsl:text>center</xsl:text>
              </xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </style:paragraph-properties>
      </xsl:if>

      <xsl:if test="e:alignment/@horizontal">
        <style:paragraph-properties>
          <!-- horizontal-align -->
          <xsl:attribute name="fo:text-align">
            <xsl:choose>
              <xsl:when test="e:alignment/@horizontal = 'left' ">
                <xsl:text>start</xsl:text>
              </xsl:when>
              <xsl:when test="e:alignment/@horizontal = 'right' ">
                <xsl:text>end</xsl:text>
              </xsl:when>
              <xsl:when test="e:alignment/@horizontal = 'centerContinuous' ">
                <xsl:text>center</xsl:text>
              </xsl:when>
              <xsl:when test="e:alignment/@horizontal = 'distributed' ">
                <xsl:text>center</xsl:text>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="e:alignment/@horizontal"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>

          <xsl:if test="e:alignment/@indent and e:alignment/@indent != '0'">
            <xsl:attribute name="fo:margin-left">
              <xsl:variable name="indennt">
                <xsl:value-of select="e:alignment/@indent"/>
              </xsl:variable>
              <xsl:value-of select="format-number($indennt*0.353,'##.###cm') "/>
            </xsl:attribute>
          </xsl:if>

          <xsl:if test="e:alignment/@indent and e:alignment/@indent != '0'">
            <xsl:attribute name="fo:margin-right">
              <xsl:variable name="indennnt">
                <xsl:value-of select="e:alignment/@indent"/>
              </xsl:variable>
              <xsl:value-of select="format-number($indennnt*0.353,'##.###cm') "/>
            </xsl:attribute>
          </xsl:if>

        </style:paragraph-properties>
      </xsl:if>
    </xsl:if>

    <xsl:if test="@applyFont = 1 or  @fontId != ''">
      <style:text-properties>
        <xsl:variable name="this" select="."/>
        <xsl:apply-templates
          select="ancestor::e:styleSheet/e:fonts/e:font[position() = $this/@fontId + 1]"
          mode="style"/>
      </style:text-properties>
    </xsl:if>
  </xsl:template>

  <!-- convert font name-->
  <xsl:template match="e:rFont" mode="style">
    <xsl:attribute name="style:font-name">
      <xsl:value-of select="@val"/>
    </xsl:attribute>
  </xsl:template>

  <!-- bold -->
  <xsl:template match="e:b" mode="style">
    <xsl:attribute name="fo:font-weight">
      <xsl:text>bold</xsl:text>
    </xsl:attribute>
  </xsl:template>

  <!-- italic -->
  <xsl:template match="e:i" mode="style">
    <xsl:attribute name="fo:font-style">
      <xsl:text>italic</xsl:text>
    </xsl:attribute>
  </xsl:template>

  <!-- insert underline -->
  <xsl:template match="e:u" mode="style">
    <xsl:call-template name="InsertUnderline"/>
  </xsl:template>

  <!-- superscript/subscript-->
  <xsl:template match="e:vertAlign" mode="style">
    <xsl:attribute name="style:text-position">
      <xsl:choose>
        <xsl:when test="@val = 'superscript' ">
          <xsl:text>super 58%</xsl:text>
        </xsl:when>
        <xsl:when test="@val = 'subscript' ">
          <xsl:text>sub 58%</xsl:text>
        </xsl:when>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>

  <xsl:template name="InsertTextStyles">

    <!-- @Description: inserts text styles -->
    <!-- @Context: none -->

    <xsl:apply-templates select="document('xl/sharedStrings.xml')/e:sst/e:si/e:r[e:rPr]"
      mode="automaticstyles"/>
    
    <xsl:for-each select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet">
      <xsl:apply-templates select="document(concat('xl/comments', position(), '.xml'))/e:comments">
        <xsl:with-param name="number">
          <xsl:value-of select="position()"/>
        </xsl:with-param>
      </xsl:apply-templates>  
    </xsl:for-each>

  </xsl:template>

  <xsl:template match="e:comments">

    <!-- @Description: inserts text styles in notes -->
    <!-- @Context: none -->

    <xsl:param name="number"/>
    <!--(int) number of sheet/comments file -->
    <xsl:apply-templates select="e:commentList/e:comment/e:text/e:r[e:rPr]" mode="automaticstyles"/>
    <xsl:apply-templates select="document(concat('xl/comments',$number + 1,'.xml'))">
      <xsl:with-param name="number">
        <xsl:value-of select="$number+1"/>
      </xsl:with-param>
    </xsl:apply-templates>
  </xsl:template>

  <!-- convert run properties into span style -->
  <xsl:template match="e:r" mode="automaticstyles">
    <style:style style:name="{generate-id(.)}" style:family="text">
      <style:text-properties fo:font-weight="normal" fo:font-style="normal"
        style:text-underline-type="none" style:text-position="0% 100%">
        <xsl:apply-templates select="e:rPr" mode="style"/>
      </style:text-properties>
    </style:style>
  </xsl:template>

  <xsl:template name="InsertUnderline">
    <xsl:choose>
      <xsl:when test="@val = 'double'">
        <xsl:attribute name="style:text-underline-type">double</xsl:attribute>
        <xsl:attribute name="style:text-underline-style">solid</xsl:attribute>
        <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
      </xsl:when>
      <xsl:when test="@val = 'single'">
        <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
        <xsl:attribute name="style:text-underline-style">solid</xsl:attribute>
        <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
      </xsl:when>
      <xsl:when test="@val = 'singleAccounting'">
        <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
        <xsl:attribute name="style:text-underline-style">accounting</xsl:attribute>
        <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
      </xsl:when>
      <xsl:when test="@val = 'doubleAccounting'">
        <xsl:attribute name="style:text-underline-type">double</xsl:attribute>
        <xsl:attribute name="style:text-underline-style">accounting</xsl:attribute>
        <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
      </xsl:when>
      <xsl:when test="@val = 'none'">
        <xsl:attribute name="style:text-underline-type">none</xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
        <xsl:attribute name="style:text-underline-style">solid</xsl:attribute>
        <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:if test="@color">
      <xsl:attribute name="style:text-underline-color">
        <xsl:choose>
          <xsl:when test="@color = 'auto'">font-color</xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat('#',@color)"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

  <!-- convert font size -->
  <xsl:template match="e:sz" mode="style">
    <xsl:attribute name="fo:font-size">
      <xsl:value-of select="round(@val)"/>
    </xsl:attribute>
  </xsl:template>

  <xsl:template match="e:name" mode="style">
    <xsl:attribute name="style:font-name">
      <xsl:value-of select="@val"/>
    </xsl:attribute>
  </xsl:template>

  <!-- insert strike-through -->
  <xsl:template match="e:strike" mode="style">
    <xsl:attribute name="style:text-line-through-style">
      <xsl:text>solid</xsl:text>
    </xsl:attribute>
  </xsl:template>

  <!-- cell color fill -->
  <xsl:template match="e:fgColor" mode="style">
    <xsl:attribute name="fo:background-color">
      <xsl:call-template name="InsertColor"/>
    </xsl:attribute>
  </xsl:template>

  <!-- cell color fill in conditional -->
  <xsl:template match="e:bgColor" mode="style">
    <xsl:if test="ancestor::e:dxf">
      <xsl:attribute name="fo:background-color">
        <xsl:call-template name="InsertColor"/>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

  <xsl:template match="e:color" mode="style">

    <xsl:attribute name="fo:color">
      <xsl:call-template name="InsertColor"/>
    </xsl:attribute>
  </xsl:template>

  <xsl:template name="InsertPageLayout">
    <xsl:variable name="ActiveTabNumber">
      <xsl:for-each select="document('xl/workbook.xml')">
        <xsl:choose>
          <xsl:when test="e:workbook/e:bookViews/e:workbookView/@activeTab">
            <xsl:value-of select="e:workbook/e:bookViews/e:workbookView/@activeTab"/>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
    </xsl:variable>

    <xsl:for-each
      select="document(concat('xl/worksheets/sheet', $ActiveTabNumber + 1,'.xml'))/e:worksheet">
      <style:page-layout style:name="pm1">
        <style:page-layout-properties>

          <!-- paper size -->
          <xsl:for-each select="e:pageSetup">
            <xsl:choose>
              <xsl:when test="@paperSize">
                <xsl:call-template name="InsertPaperSize">
                  <xsl:with-param name="paperSize" select="@paperSize"/>
                  <xsl:with-param name="orientation">
                    <xsl:choose>
                      <xsl:when test="@orientation = 'landscape' ">
                        <xsl:text>landscape</xsl:text>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:text>portrait</xsl:text>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:when>
              <!-- Letter -->
              <xsl:otherwise>
                <xsl:attribute name="fo:page-width">
                  <xsl:text>27.94cm</xsl:text>
                </xsl:attribute>
                <xsl:attribute name="fo:page-height">
                  <xsl:text>21.59cm</xsl:text>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>

            <!-- paper orientation -->
            <xsl:if test="@orientation">
              <xsl:attribute name="style:print-orientation">
                <xsl:value-of select="@orientation"/>
              </xsl:attribute>
            </xsl:if>

            <!-- page order -->
            <xsl:if test="@pageOrder = 'overThenDown' ">
              <xsl:attribute name="style:print-page-order">
                <xsl:text>ltr</xsl:text>
              </xsl:attribute>
            </xsl:if>

            <!-- first page number -->
            <xsl:if test="@useFirstPageNumber = 1 and @firstPageNumber">
              <xsl:attribute name="style:first-page-number">
                <xsl:value-of select="@firstPageNumber"/>
              </xsl:attribute>
            </xsl:if>

            <!-- fit on pages wide / pages tall -->
            <xsl:if test="parent::node()/e:sheetPr/e:pageSetUpPr/@fitToPage = 1">
              <xsl:attribute name="style:scale-to-X">
                <xsl:choose>
                  <xsl:when test="@fitToWidth">
                    <xsl:value-of select="@fitToWidth"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:text>1</xsl:text>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>

              <xsl:attribute name="style:scale-to-Y">
                <xsl:choose>
                  <xsl:when test="@fitToHeight">
                    <xsl:value-of select="@fitToHeight"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:text>1</xsl:text>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>

            <!-- scale printout -->
            <xsl:if test="@scale">
              <xsl:attribute name="style:scale-to">
                <xsl:value-of select="@scale"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>

          <xsl:call-template name="InsertPageMargins"/>

          <xsl:for-each select="e:printOptions">
            <!-- table alignment -->
            <xsl:if test="@horizontalCentered = 1 or @verticalCentered = 1">
              <xsl:attribute name="style:table-centering">
                <xsl:choose>
                  <xsl:when test="@horizontalCentered = 1 and @verticalCentered = 1">
                    <xsl:text>both</xsl:text>
                  </xsl:when>
                  <xsl:when test="@horizontalCentered = 1">
                    <xsl:text>horizontal</xsl:text>
                  </xsl:when>
                  <xsl:when test="@verticalCentered = 1">
                    <xsl:text>vertical</xsl:text>
                  </xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>

            <!-- print -->
            <xsl:if
              test="@headings = 1 or @gridLines = 1 or (parent::node()/e:pageSetup/@cellComments and parent::node()/e:pageSetup/@cellComments != 'none') ">
              <!-- notes -->
              <xsl:variable name="printNotes">
                <xsl:choose>
                  <xsl:when test="parent::node()/e:pageSetup/@cellComments != 'none' ">
                    <xsl:text>annotations charts drawings objects zero-values</xsl:text>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:text>charts drawings objects zero-values</xsl:text>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>

              <!-- gridlines-->
              <xsl:variable name="printGrid">
                <xsl:choose>
                  <xsl:when test="@gridLines = 1">
                    <xsl:value-of select="concat($printNotes,' grid')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="$printNotes"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>

              <!-- row and column headings -->
              <xsl:variable name="printHeaders">
                <xsl:choose>
                  <xsl:when test="@headings = 1">
                    <xsl:value-of select="concat($printGrid,' headers')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="$printGrid"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>

              <xsl:attribute name="style:print">
                <xsl:value-of select="$printHeaders"/>
              </xsl:attribute>
            </xsl:if>

          </xsl:for-each>

        </style:page-layout-properties>
        <style:header-style>
          <style:header-footer-properties fo:min-height="0.751cm" fo:margin-left="0cm"
            fo:margin-right="0cm" fo:margin-bottom="0.25cm"/>
        </style:header-style>
        <style:footer-style>
          <style:header-footer-properties fo:min-height="0.751cm" fo:margin-left="0cm"
            fo:margin-right="0cm" fo:margin-top="0.25cm"/>
        </style:footer-style>
      </style:page-layout>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="InsertPageMargins">
    <xsl:for-each select="e:pageMargins">
      <xsl:if test="@top">
        <xsl:attribute name="fo:margin-top">
          <xsl:call-template name="ConvertToCentimeters">
            <xsl:with-param name="length" select="concat(@top,'in')"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@bottom">
        <xsl:attribute name="fo:margin-bottom">
          <xsl:call-template name="ConvertToCentimeters">
            <xsl:with-param name="length" select="concat(@bottom,'in')"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@left">
        <xsl:attribute name="fo:margin-left">
          <xsl:call-template name="ConvertToCentimeters">
            <xsl:with-param name="length" select="concat(@left,'in')"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@right">
        <xsl:attribute name="fo:margin-right">
          <xsl:call-template name="ConvertToCentimeters">
            <xsl:with-param name="length" select="concat(@right,'in')"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>


  <xsl:template name="InsertCellStyle">
    <xsl:for-each select="document('xl/styles.xml')/e:styleSheet/e:cellXfs">
      <xsl:apply-templates select="e:xf" mode="stylesandformating"/>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="e:xf" mode="stylesandformating">
    <xsl:variable name="Xfid">
      <xsl:value-of select="@xfId"/>
    </xsl:variable>
    <xsl:if
      test="$Xfid != '0' and not(preceding-sibling::e:xf/@xfId = $Xfid) and key('CellStylesId', @xfId)/@name != 'Default'">
      <style:style>
        <xsl:attribute name="style:name">
          <xsl:value-of select="key('CellStylesId', @xfId)/@name"/>
        </xsl:attribute>
        <xsl:attribute name="style:family">
          <xsl:text>table-cell</xsl:text>
        </xsl:attribute>
        <xsl:call-template name="InsertCellFormat"/>
      </style:style>
    </xsl:if>
  </xsl:template>

</xsl:stylesheet>

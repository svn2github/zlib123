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
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  exclude-result-prefixes="a e r">

  <xsl:import href="relationships.xsl"/>

  <xsl:template name="styles">
    <office:document-styles>
      <office:font-face-decls>
        <xsl:call-template name="InsertFonts"/>
      </office:font-face-decls>
      <office:styles>
        <xsl:call-template name="InsertDefaultTableCellStyle"/>
      </office:styles>
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
        <xsl:apply-templates select="document('xl/styles.xml')/e:styleSheet/e:fonts/e:font[1]" mode="style"/>
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

    <!-- default style -->
    <style:style style:name="{generate-id(document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr)}"
      style:family="table-column">
      <style:table-column-properties fo:break-before="auto">
        <xsl:attribute name="style:column-width">
          <xsl:choose>
            <xsl:when test="document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr/@defaultColWidth">
              <xsl:call-template name="ConvertFromCharacters">
                <xsl:with-param name="value">
                  <xsl:value-of
                    select="document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr/@defaultColWidth"/>
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

    <xsl:apply-templates select="document(concat('xl/',$sheet))/e:worksheet/e:cols" mode="automaticstyles"/>
  </xsl:template>

  <xsl:template match="e:col" mode="automaticstyles">
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

    <!-- default style -->
    <style:style style:name="{generate-id(document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr)}"
      style:family="table-row">
      <style:table-row-properties fo:break-before="auto">
        <xsl:attribute name="style:row-height">
          <xsl:choose>
            <xsl:when test="document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr/@defaultRowHeight">
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
        <style:table-properties>
          <xsl:if test="@state='hidden'">
            <xsl:attribute name="table:display">
              <xsl:text>false</xsl:text>
            </xsl:attribute>
          </xsl:if>
        </style:table-properties>
      </style:style>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="InsertCellStyles">
    <xsl:apply-templates select="document('xl/styles.xml')/e:styleSheet/e:cellXfs" mode="automaticstyles"/>
  </xsl:template>

  <!-- cell formats -->
  <xsl:template match="e:xf" mode="automaticstyles">
    <style:style style:name="{generate-id(.)}" style:family="table-cell">
      <xsl:if test="@applyAlignment = 1">
        <style:table-cell-properties>
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
          <xsl:if test="e:alignment/@horizontal = 'fill' ">
            <xsl:attribute name="style:repeat-content">
              <xsl:text>true</xsl:text>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="e:alignment/@horizontal = 'distributed' or e:alignment/@vertical = 'distributed' ">
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
                    <xsl:when test="e:alignment/@textRotation &lt; 90 or e:alignment/@textRotation = 90">
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
        </style:table-cell-properties>

        <!-- default horizontal alignment when text has angle orientation  -->
        <xsl:if test="not(e:alignment/@horizontal) and e:alignment/@textRotation">
          <style:paragraph-properties>
            <xsl:attribute name="fo:text-align">
              <xsl:choose>
                <xsl:when test="e:alignment/@textRotation &lt; 90 or e:alignment/@textRotation = 180">
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
          </style:paragraph-properties>
        </xsl:if>
      </xsl:if>

      <xsl:if test="@applyFont = 1">
        <style:text-properties>
          <xsl:variable name="this" select="."/>
          <xsl:apply-templates select="ancestor::e:styleSheet/e:fonts/e:font[position() = $this/@fontId + 1]"
            mode="style"/>
        </style:text-properties>
      </xsl:if>
    </style:style>
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

  <!-- insert text styles -->
  <xsl:template name="InsertTextStyles">
    <xsl:apply-templates select="document('xl/sharedStrings.xml')/e:sst/e:si/e:r[e:rPr]"
      mode="automaticstyles"/>
  </xsl:template>

  <!-- convert run properties into span style -->
  <xsl:template match="e:r" mode="automaticstyles">
    <style:style style:name="{generate-id(.)}" style:family="text">
      <style:text-properties fo:font-weight="normal" fo:font-style="normal" style:text-underline-type="none">
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
      <xsl:value-of select="@val"/>
    </xsl:attribute>
  </xsl:template>

  <xsl:template match="e:name" mode="style">
    <xsl:attribute name="style:font-name">
      <xsl:value-of select="@val"/>
    </xsl:attribute>
  </xsl:template>

  <!-- insert text-line-through -->
  <xsl:template match="e:strike" mode="style">
    <xsl:attribute name="style:text-line-through-style">
      <xsl:text>solid</xsl:text>
    </xsl:attribute>
  </xsl:template>

  <xsl:template match="e:color" mode="style">

    <xsl:attribute name="fo:color">
      <xsl:variable name="this" select="."/>
      <xsl:choose>
        <xsl:when test="@rgb">
          <xsl:value-of select="concat('#',substring(@rgb,3,9))"/>
        </xsl:when>
        <xsl:when test="@indexed">
          <xsl:variable name="color">
            <xsl:value-of select="document('xl/styles.xml')/e:styleSheet/e:colors/e:indexedColors/e:rgbColor[$this/@indexed + 1]/@rgb"/>
          </xsl:variable>
          <xsl:value-of select="concat('#',substring($color,3,9))"/>
        </xsl:when>
        <xsl:when test="@theme">
          <xsl:variable name="color">
            <xsl:choose>
              <!-- if color is in 'sysClr' node -->
              <xsl:when
                test="document('xl/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/child::node()[position() = $this/@theme]/child::node()/@lastClr">
                <xsl:value-of
                  select="document('xl/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/child::node()[position() = $this/@theme]/child::node()/@lastClr"
                />
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of
                  select="document('xl/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/child::node()[position() = $this/@theme]/child::node()/@val"
                />
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:value-of select="concat('#',$color)"/>
        </xsl:when>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>

</xsl:stylesheet>

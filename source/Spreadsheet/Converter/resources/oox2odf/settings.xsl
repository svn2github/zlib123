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
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ooo="http://openoffice.org/2004/office" office:version="1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  exclude-result-prefixes="e r w">

  <xsl:import href="relationships.xsl"/>
  <xsl:import href="common.xsl"/>

  <xsl:template name="InsertSettings">
    <office:document-settings>
      <office:settings>
        <config:config-item-set config:name="ooo:view-settings">
          <config:config-item-map-indexed config:name="Views">
            <config:config-item-map-entry>

              <!-- settings based on default active table -->
              <xsl:for-each select="document('xl/workbook.xml')">
                <xsl:variable name="ActiveTabNumber">
                  <xsl:choose>
                    <xsl:when test="e:workbook/e:bookViews/e:workbookView/@activeTab">
                      <xsl:value-of select="e:workbook/e:bookViews/e:workbookView/@activeTab"/>
                    </xsl:when>
                    <xsl:otherwise>0</xsl:otherwise>
                  </xsl:choose>
                </xsl:variable>

                <!-- Set default Active Table (Sheet) -->
                <config:config-item config:name="ActiveTable" config:type="string">
                  <xsl:for-each
                    select="e:workbook/e:sheets/e:sheet[position() = $ActiveTabNumber + 1]/@name">
                    <xsl:value-of select="."/>
                  </xsl:for-each>
                </config:config-item>

                <config:config-item config:name="ZoomValue" config:type="int">
                  <xsl:for-each
                    select="document(concat('xl/worksheets/sheet', $ActiveTabNumber + 1,'.xml'))/e:worksheet/e:sheetViews/e:sheetView">
                    <xsl:choose>
                      <xsl:when test="not(@view = 'pageBreakPreview') and @zoomScale">
                        <xsl:value-of select="@zoomScale"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:text>100</xsl:text>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:for-each>
                </config:config-item>

                <config:config-item config:name="PageViewZoomValue" config:type="int">
                  <xsl:for-each
                    select="document(concat('xl/worksheets/sheet', $ActiveTabNumber + 1,'.xml'))/e:worksheet/e:sheetViews/e:sheetView">
                    <xsl:choose>
                      <xsl:when test="@view = 'pageBreakPreview' and @zoomScale">
                        <xsl:value-of select="@zoomScale"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:text>100</xsl:text>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:for-each>
                </config:config-item>

                <config:config-item config:name="ShowPageBreakPreview" config:type="boolean">
                  <xsl:for-each
                    select="document(concat('xl/worksheets/sheet', $ActiveTabNumber + 1,'.xml'))/e:worksheet/e:sheetViews/e:sheetView">
                    <xsl:choose>
                      <xsl:when test="@view = 'pageBreakPreview' ">
                        <xsl:text>true</xsl:text>
                      </xsl:when>
                      <xsl:otherwise>false</xsl:otherwise>
                    </xsl:choose>
                  </xsl:for-each>
                </config:config-item>
              </xsl:for-each>

              <config:config-item-map-named config:name="Tables">
                <xsl:for-each select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet">
                  <xsl:call-template name="InsertCursorPosition">
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
              </config:config-item-map-named>
            </config:config-item-map-entry>
          </config:config-item-map-indexed>
        </config:config-item-set>
      </office:settings>
    </office:document-settings>
  </xsl:template>

  <xsl:template name="InsertCursorPosition">
    <xsl:param name="sheet"/>
    <config:config-item-map-entry config:name="{@name}">
      <config:config-item config:name="CursorPositionX" config:type="int">
        <xsl:choose>
          <xsl:when
            test="document(concat('xl/',$sheet))/e:worksheet/e:sheetViews/e:sheetView/e:selection/@activeCell">
            <xsl:variable name="col">
              <xsl:call-template name="GetColNum">
                <xsl:with-param name="cell">
                  <xsl:value-of
                    select="document(concat('xl/',$sheet))/e:worksheet/e:sheetViews/e:sheetView/e:selection/@activeCell"
                  />
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:value-of select="$col - 1"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>0</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </config:config-item>
      <config:config-item config:name="CursorPositionY" config:type="int">
        <xsl:choose>
          <xsl:when
            test="document(concat('xl/',$sheet))/e:worksheet/e:sheetViews/e:sheetView/e:selection/@activeCell">
            <xsl:variable name="row">
              <xsl:call-template name="GetRowNum">
                <xsl:with-param name="cell">
                  <xsl:value-of
                    select="document(concat('xl/',$sheet))/e:worksheet/e:sheetViews/e:sheetView/e:selection/@activeCell"
                  />
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:value-of select="$row - 1"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>0</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </config:config-item>
    </config:config-item-map-entry>
  </xsl:template>
</xsl:stylesheet>

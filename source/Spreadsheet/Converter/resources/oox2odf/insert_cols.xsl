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
  xmlns:v="urn:schemas-microsoft-com:vml" xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main" exclude-result-prefixes="e r">





  <xsl:template match="e:col" mode="header">
    <xsl:param name="number"/>
    <xsl:param name="sheet"/>
    <xsl:param name="DefaultCellStyleName"/>
    <xsl:param name="headerColsStart"/>
    <xsl:param name="headerColsEnd"/>
    <xsl:param name="GroupCell"/>
    <xsl:param name="ManualColBreaks"/>


    <!-- if there were columns with default properties before this column then insert default columns-->
    <xsl:choose>
      <!-- when this column is the first non-default one but it's not the column A -->
      <xsl:when
        test="$number = 1 and @min &gt; 1 and ($headerColsStart= '' or $headerColsStart &gt; 1)">
        <table:table-column>
          <xsl:attribute name="table:style-name">
            <xsl:for-each select="document(concat('xl/',$sheet))">
              <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
            </xsl:for-each>
          </xsl:attribute>

          <xsl:attribute name="table:default-cell-style-name">
            <xsl:value-of select="$DefaultCellStyleName"/>
          </xsl:attribute>

          <xsl:attribute name="table:number-columns-repeated">
            <xsl:choose>
              <!-- when there is a header -->
              <xsl:when test="$headerColsStart != '' ">
                <xsl:choose>
                  <xsl:when test="@min &lt; $headerColsStart">
                    <xsl:value-of select="@min - 1"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="$headerColsStart - 1"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>

              <xsl:otherwise>
                <xsl:value-of select="@min - 1"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>

          <!-- Possible are nesesary code -->
          <!--xsl:if test="@style">            
                      <xsl:variable name="position">
                      <xsl:value-of select="$this/@style + 1"/>
                      </xsl:variable>
                      
                      <xsl:attribute name="table:default-cell-style-name">
                      <xsl:for-each select="document('xl/styles.xml')">
                      <xsl:value-of select="generate-id(key('Xf', '')[position() = $position])"/>
                      </xsl:for-each>
                      </xsl:attribute>
                      </xsl:if-->

        </table:table-column>

      </xsl:when>
      <!-- when this column is not first non-default one and there were default columns after previous non-default column (if there was a gap between this and previous column)-->
      <xsl:when test="preceding-sibling::e:col[1]/@max &lt; @min - 1 ">
        <table:table-column>
          <xsl:attribute name="table:style-name">
            <xsl:for-each select="document(concat('xl/',$sheet))">
              <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
            </xsl:for-each>
          </xsl:attribute>

          <xsl:attribute name="table:number-columns-repeated">
            <xsl:choose>
              <!-- when there is a header -->
              <xsl:when test="$headerColsStart != '' ">
                <xsl:choose>
                  <!-- insert empty columns before header -->
                  <xsl:when
                    test="@min &gt; $headerColsStart and preceding-sibling::e:col[1]/@max &lt; $headerColsStart">
                    <xsl:value-of select="$headerColsStart - preceding::e:col[1]/@max - 1"/>
                  </xsl:when>
                  <!-- insert empty columns after header -->
                  <xsl:when
                    test="@min &gt; $headerColsEnd and preceding-sibling::e:col[1]/@max &lt; $headerColsEnd">
                    <xsl:value-of select="@max - $headerColsEnd - 1"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="@min - preceding::e:col[1]/@max - 1"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>

              <xsl:otherwise>
                <xsl:value-of select="@min - preceding::e:col[1]/@max - 1"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>

          <xsl:attribute name="table:default-cell-style-name">
            <xsl:value-of select="$DefaultCellStyleName"/>
          </xsl:attribute>

        </table:table-column>
      </xsl:when>
    </xsl:choose>

    <!-- insert this column -->
    <xsl:call-template name="InsertThisColumn">
      <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
      <xsl:with-param name="headerColsEnd" select="$headerColsEnd"/>
      <xsl:with-param name="GroupCell">
        <xsl:value-of select="$GroupCell"/>
      </xsl:with-param>
      <xsl:with-param name="ManualColBreaks">
        <xsl:value-of select="$ManualColBreaks"/>
      </xsl:with-param>
    </xsl:call-template>

    <!-- insert next column -->
    <xsl:choose>
      <!-- calc supports only 256 columns -->
      <xsl:when test="$number &gt; 255">
        <xsl:message terminate="no">translation.oox2odf.ColNumber</xsl:message>
      </xsl:when>

      <xsl:when
        test="following-sibling::e:col[1][@min &lt;= $headerColsEnd and @min &lt; 256]">
        <xsl:apply-templates select="following-sibling::e:col[1]" mode="header">
          <xsl:with-param name="number" select="@max + 1"/>
          <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
          <xsl:with-param name="sheet" select="$sheet"/>
          <xsl:with-param name="headerColsStart" select="$headerColsStart"/>
          <xsl:with-param name="headerColsEnd" select="$headerColsEnd"/>
          <xsl:with-param name="GroupCell">
            <xsl:value-of select="$GroupCell"/>
          </xsl:with-param>
          <xsl:with-param name="ManualColBreaks">
            <xsl:value-of select="$ManualColBreaks"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>

    </xsl:choose>

  </xsl:template>




  <xsl:template match="e:col">
    <xsl:param name="number"/>
    <xsl:param name="sheet"/>
    <xsl:param name="DefaultCellStyleName"/>
    <xsl:param name="headerColsStart"/>
    <xsl:param name="headerColsEnd"/>
    <xsl:param name="GroupCell"/>
    <xsl:param name="ManualColBreaks"/>
    <xsl:param name="prevManualBreak" select="0"/>
    <xsl:param name="AfterRow"/>

    <xsl:variable name="GetMinManualColBreak">
      <xsl:call-template name="GetMinRowWithPicture">
        <xsl:with-param name="PictureRow">
          <xsl:value-of select="$ManualColBreaks"/>
        </xsl:with-param>
        <xsl:with-param name="AfterRow">
          <xsl:choose>
            <xsl:when test="preceding-sibling::e:col/@max">
              <xsl:value-of select="preceding-sibling::e:col/@max"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>1</xsl:text>
            </xsl:otherwise>
          </xsl:choose>

        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>


    <!-- if there were columns with default properties before this column then insert default columns-->
    <xsl:choose>
      <!-- when this column is the first non-default one but it's not the column A -->
      <xsl:when
        test="$number = 1 and @min &gt; 1 and ($headerColsStart= '' or $headerColsStart &gt; 1)">


        <xsl:call-template name="InsertColBreak">
          <xsl:with-param name="ManualColBreaks">
            <xsl:value-of select="$ManualColBreaks"/>
          </xsl:with-param>
          <xsl:with-param name="GetMinManualColBreak">
            <xsl:value-of select="$GetMinManualColBreak"/>
          </xsl:with-param>
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="DefaultCellStyleName">
            <xsl:value-of select="$DefaultCellStyleName"/>
          </xsl:with-param>
        </xsl:call-template>

        <xsl:variable name="ManualCol">
            <xsl:call-template name="GetMaxValueBetweenTwoValues">
              <xsl:with-param name="min">
                <xsl:value-of select="preceding-sibling::e:col/@max +1"/>
              </xsl:with-param>
              <xsl:with-param name="max">
                <xsl:value-of select="@min"/>
              </xsl:with-param>
              <xsl:with-param name="value">
                <xsl:value-of select="$ManualColBreaks"/>                
              </xsl:with-param>
            </xsl:call-template>
        </xsl:variable>
        
        <table:table-column>
          <xsl:attribute name="table:style-name">
            <xsl:for-each select="document(concat('xl/',$sheet))">
              <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
            </xsl:for-each>
          </xsl:attribute>

          <xsl:attribute name="table:default-cell-style-name">
            <xsl:value-of select="$DefaultCellStyleName"/>
          </xsl:attribute>
          
          <xsl:attribute name="table:number-columns-repeated">
            <xsl:choose>
              <!-- when there is a header -->
              <xsl:when test="$headerColsStart != '' ">
                <xsl:choose>
                  <xsl:when test="@min &lt; $headerColsStart">
                    <xsl:value-of select="@min - 1"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="$headerColsStart - 1"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>

              <xsl:otherwise>
                <xsl:choose>
                  <xsl:when test="@min &gt; $GetMinManualColBreak">
                    <xsl:value-of select="@min - ($ManualCol + 1) - 1"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="@min - 1"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>

          <!-- Possible are nesesary code -->
          <!--xsl:if test="@style">            
            <xsl:variable name="position">
            <xsl:value-of select="$this/@style + 1"/>
            </xsl:variable>
            
            <xsl:attribute name="table:default-cell-style-name">
              <xsl:for-each select="document('xl/styles.xml')">
                <xsl:value-of select="generate-id(key('Xf', '')[position() = $position])"/>
              </xsl:for-each>
            </xsl:attribute>
            </xsl:if-->
        </table:table-column>

      </xsl:when>
      <!-- when this column is not first non-default one and there were default columns after previous non-default column (if there was a gap between this and previous column)-->
      <xsl:when test="preceding-sibling::e:col[1]/@max &lt; @min - 1">
        
        <xsl:call-template name="InsertColBreak">
          <xsl:with-param name="ManualColBreaks">
            <xsl:value-of select="$ManualColBreaks"/>
          </xsl:with-param>
          <xsl:with-param name="GetMinManualColBreak">
            <xsl:value-of select="$GetMinManualColBreak"/>
          </xsl:with-param>
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="DefaultCellStyleName">
            <xsl:value-of select="$DefaultCellStyleName"/>
          </xsl:with-param>
        </xsl:call-template>
        
        <xsl:variable name="ManualCol">
          <xsl:call-template name="GetMaxValueBetweenTwoValues">
            <xsl:with-param name="min">
              <xsl:text>1</xsl:text>
            </xsl:with-param>
            <xsl:with-param name="max">
              <xsl:value-of select="@min"/>
            </xsl:with-param>
            <xsl:with-param name="value">
              <xsl:value-of select="$ManualColBreaks"/>                
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        
        <xsl:choose>
          <!-- when there is a header -->
          <xsl:when test="$headerColsStart != '' ">
            <xsl:choose>
              <xsl:when test="preceding-sibling::e:col/@max + 1 = $headerColsStart"/>
              <!-- insert empty columns before header -->
              <xsl:when
                test="@min &gt; $headerColsStart and preceding-sibling::e:col[1]/@max &lt; $headerColsStart">
                <table:table-column>
                  <xsl:attribute name="table:style-name">
                    <xsl:for-each select="document(concat('xl/',$sheet))">
                      <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
                    </xsl:for-each>
                  </xsl:attribute>

                  <xsl:attribute name="table:default-cell-style-name">
                    <xsl:value-of select="$DefaultCellStyleName"/>
                  </xsl:attribute>

                  <xsl:if test="$headerColsStart - preceding::e:col[1]/@max - 1 &gt; 1">
                    <xsl:attribute name="table:number-columns-repeated">
                      <xsl:value-of select="$headerColsStart - preceding::e:col[1]/@max - 1"/>
                    </xsl:attribute>
                  </xsl:if>
                </table:table-column>
              </xsl:when>
              <xsl:when test="@min = $headerColsEnd + 1"/>
              <!-- insert empty columns after header -->
              <xsl:when
                test="@min &gt; $headerColsEnd and preceding-sibling::e:col[1]/@max &lt; $headerColsEnd">
                <table:table-column>
                  <xsl:attribute name="table:style-name">
                    <xsl:for-each select="document(concat('xl/',$sheet))">
                      <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
                    </xsl:for-each>
                  </xsl:attribute>

                  <xsl:attribute name="table:default-cell-style-name">
                    <xsl:value-of select="$DefaultCellStyleName"/>
                  </xsl:attribute>

                  <xsl:if test="@min - $headerColsEnd - 1 &gt; 1">
                    <xsl:attribute name="table:number-columns-repeated">
                      <xsl:value-of select="@min - $headerColsEnd - 1"/>
                    </xsl:attribute>
                  </xsl:if>
                </table:table-column>
              </xsl:when>
              <!-- insert simple empty rows -->
              <xsl:otherwise>
                <table:table-column>
                  <xsl:attribute name="table:style-name">
                    <xsl:for-each select="document(concat('xl/',$sheet))">
                      <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
                    </xsl:for-each>
                  </xsl:attribute>

                  <xsl:attribute name="table:default-cell-style-name">
                    <xsl:value-of select="$DefaultCellStyleName"/>
                  </xsl:attribute>

                  <xsl:if test="@min - preceding::e:col[1]/@max - 1 &gt; 1">
                    <xsl:attribute name="table:number-columns-repeated">
                      <xsl:value-of select="@min - preceding::e:col[1]/@max - 1"/>
                    </xsl:attribute>
                  </xsl:if>
                </table:table-column>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>

          <xsl:otherwise>
            <table:table-column>
              <xsl:attribute name="table:style-name">
                <xsl:for-each select="document(concat('xl/',$sheet))">
                  <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
                </xsl:for-each>
              </xsl:attribute>

              <xsl:attribute name="table:number-columns-repeated">
                <xsl:choose>
                  <xsl:when test="@min &gt; $GetMinManualColBreak">
                    <xsl:value-of select="@min - ($ManualCol + 1) - 1"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="@min - preceding::e:col[1]/@max - 1"/>
                  </xsl:otherwise>
                </xsl:choose>                
              </xsl:attribute>

              <xsl:attribute name="table:default-cell-style-name">
                <xsl:value-of select="$DefaultCellStyleName"/>
              </xsl:attribute>
            </table:table-column>
          </xsl:otherwise>
        </xsl:choose>

      </xsl:when>
    </xsl:choose>

    <!-- insert this column -->
    <xsl:choose>
      <!-- if this is the first column after beginning of the header -->
      <xsl:when
        test="$headerColsStart != '' and (@max &gt;= $headerColsStart and not(preceding-sibling::e:col[1]/@max &gt;= $headerColsStart))">

        <!-- insert part of a column range (range: @max > @min) before header-->
        <xsl:if test="@min &lt; $headerColsStart and @max &gt;= $headerColsStart">
          <xsl:call-template name="InsertThisColumn">
            <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
            <xsl:with-param name="headerColsStart" select="$headerColsStart"/>
            <xsl:with-param name="beforeHeader" select="'true'"/>
            <xsl:with-param name="GroupCell">
              <xsl:value-of select="$GroupCell"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:if>

        <table:table-header-columns>

          <!-- insert column settings at the beginning of the header -->
          <xsl:choose>
            <!-- insert previous column settings -->
            <xsl:when test="preceding-sibling::e:col[1][@max &gt;= $headerColsStart]">
              <xsl:variable name="preceding">
                <xsl:value-of select="preceding-sibling::node()[1]"/>
              </xsl:variable>

              <table:table-column table:style-name="{generate-id(preceding-sibling::node()[1])}">
                <xsl:attribute name="table:number-columns-repeated">
                  <xsl:value-of select="preceding-sibling::node()[1]/@max - $headerColsStart - 1"/>
                </xsl:attribute>

                <xsl:attribute name="table:default-cell-style-name">
                  <xsl:value-of select="$DefaultCellStyleName"/>
                </xsl:attribute>

                <xsl:if test="preceding-sibling::e:col[1]/@hidden=1">
                  <xsl:attribute name="table:visibility">
                    <xsl:text>collapse</xsl:text>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="preceding-sibling::e:col[1]/@style">
                  <xsl:variable name="position">
                    <xsl:value-of select="preceding-sibling::e:col[1]/@style + 1"/>
                  </xsl:variable>
                  <xsl:attribute name="table:default-cell-style-name">
                    <xsl:for-each select="document('xl/styles.xml')">
                      <xsl:value-of select="generate-id(key('Xf', '')[position() = $position])"/>
                    </xsl:for-each>
                  </xsl:attribute>
                </xsl:if>
              </table:table-column>
            </xsl:when>
            <xsl:when test="@min &gt; $headerColsStart">
              <table:table-column table:style-name="{generate-id(key('SheetFormatPr', ''))}">
                <xsl:if test="@min - $headerColsStart &gt; 0">
                  <xsl:attribute name="table:number-columns-repeated">
                    <xsl:choose>
                      <xsl:when test="@min &lt;= $headerColsEnd">
                        <xsl:value-of select="@min - $headerColsStart"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="$headerColsEnd - $headerColsStart + 1"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
                </xsl:if>

                <xsl:attribute name="table:default-cell-style-name">
                  <xsl:value-of select="$DefaultCellStyleName"/>
                </xsl:attribute>
              </table:table-column>
            </xsl:when>
          </xsl:choose>

          <xsl:if test="@min &lt;= $headerColsEnd">
            <xsl:call-template name="InsertThisColumn">
              <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
              <xsl:with-param name="headerColsStart" select="$headerColsStart"/>
              <xsl:with-param name="headerColsEnd" select="$headerColsEnd"/>
              <xsl:with-param name="GroupCell">
                <xsl:value-of select="$GroupCell"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:if>

          <!-- insert next header-column -->
          <xsl:if
            test="following-sibling::e:col[@min &lt;= $headerColsEnd and @min &lt; 256][1]">
            <xsl:apply-templates
              select="following-sibling::e:col[@min &lt;= $headerColsEnd and @min &lt; 256][1]"
              mode="header">
              <xsl:with-param name="number" select="@max + 1"/>
              <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
              <xsl:with-param name="sheet" select="$sheet"/>
              <xsl:with-param name="headerColsStart" select="$headerColsStart"/>
              <xsl:with-param name="headerColsEnd" select="$headerColsEnd"/>
              <xsl:with-param name="GroupCell">
                <xsl:value-of select="$GroupCell"/>
              </xsl:with-param>
            </xsl:apply-templates>
          </xsl:if>

          <!-- insert empty columns at the end of the header -->
          <xsl:for-each
            select="parent::node()/e:col[@min &lt;= 256 and @max &gt;= $headerColsStart and @min &lt;= $headerColsEnd][last()]">
            <xsl:if test="@max &lt; $headerColsEnd">
              <table:table-column table:style-name="{generate-id(key('SheetFormatPr', ''))}">

                <xsl:attribute name="table:number-columns-repeated">
                  <xsl:value-of select="$headerColsEnd - @max"/>
                </xsl:attribute>

                <xsl:attribute name="table:default-cell-style-name">
                  <xsl:value-of select="$DefaultCellStyleName"/>
                </xsl:attribute>
              </table:table-column>
            </xsl:if>
          </xsl:for-each>

        </table:table-header-columns>

        <!-- if this column range starts inside header, but ends outside write rest of the columns outside header -->
        <xsl:if test="@min &lt;= $headerColsEnd and @max &gt; $headerColsEnd">
          <xsl:call-template name="InsertThisColumn">
            <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
            <xsl:with-param name="headerColsEnd" select="$headerColsEnd"/>
            <xsl:with-param name="afterHeader" select="'true'"/>
            <xsl:with-param name="GroupCell">
              <xsl:value-of select="$GroupCell"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:if>

        <!-- if further there is column range starts inside header, but ends outside write rest of the columns outside header -->
        <xsl:for-each
          select="following-sibling::e:col[@min &lt;= $headerColsEnd and @max &gt; $headerColsEnd]">
          <xsl:call-template name="InsertThisColumn">
            <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
            <xsl:with-param name="headerColsEnd" select="$headerColsEnd"/>
            <xsl:with-param name="afterHeader" select="'true'"/>
            <xsl:with-param name="GroupCell">
              <xsl:value-of select="$GroupCell"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:for-each>

        <!-- if header is empty -->
        <xsl:if test="@min &gt; $headerColsEnd">

          <xsl:choose>
            <!-- first row after start of the header is right after end of the header -->
            <xsl:when test="@min = $headerColsEnd + 1">
              <xsl:call-template name="InsertThisColumn">
                <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
                <xsl:with-param name="GroupCell">
                  <xsl:value-of select="$GroupCell"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <!-- insert default columns between header and column -->
              <table:table-column table:style-name="{generate-id(key('SheetFormatPr', ''))}">

                <xsl:if test="@min &gt; $headerColsEnd + 2">
                  <xsl:attribute name="table:number-columns-repeated">
                    <xsl:value-of select="@min - $headerColsEnd - 1"/>
                  </xsl:attribute>
                </xsl:if>

                <xsl:attribute name="table:default-cell-style-name">
                  <xsl:value-of select="$DefaultCellStyleName"/>
                </xsl:attribute>
              </table:table-column>

              <xsl:call-template name="InsertThisColumn">
                <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
                <xsl:with-param name="GroupCell">
                  <xsl:value-of select="$GroupCell"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>

      </xsl:when>
      <!-- if header is after the last column -->
      <xsl:when test="not(following-sibling::e:col) and @max &lt; $headerColsStart">

        <xsl:call-template name="InsertThisColumn">
          <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
          <xsl:with-param name="GroupCell">
            <xsl:value-of select="$GroupCell"/>
          </xsl:with-param>
        </xsl:call-template>

        <!-- insert default columns before header -->
        <xsl:if test="$headerColsStart &gt; @max + 1">
          <table:table-column table:style-name="{generate-id(key('SheetFormatPr', ''))}">

            <xsl:if test="$headerColsStart &gt; @max + 2">
              <xsl:attribute name="table:number-columns-repeated">
                <xsl:value-of select="$headerColsStart - @max - 1"/>
              </xsl:attribute>
            </xsl:if>

            <xsl:attribute name="table:default-cell-style-name">
              <xsl:value-of select="$DefaultCellStyleName"/>
            </xsl:attribute>
          </table:table-column>
        </xsl:if>

        <table:table-header-columns>
          <table:table-column table:style-name="{generate-id(key('SheetFormatPr', ''))}">

            <xsl:if test="$headerColsEnd  - $headerColsStart &gt; 1">
              <xsl:attribute name="table:number-columns-repeated">
                <xsl:value-of select="$headerColsEnd  - $headerColsStart + 1"/>
              </xsl:attribute>
            </xsl:if>

            <xsl:attribute name="table:default-cell-style-name">
              <xsl:value-of select="$DefaultCellStyleName"/>
            </xsl:attribute>
          </table:table-column>
        </table:table-header-columns>

      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="InsertThisColumn">
          <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
          <xsl:with-param name="GroupCell">
            <xsl:value-of select="$GroupCell"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>

    <xsl:choose>

      <!-- calc supports only 256 columns -->
      <xsl:when test="$number &gt; 255">
        <xsl:message terminate="no">translation.oox2odf.ColNumber</xsl:message>
      </xsl:when>

      <xsl:when test="$headerColsStart = '' ">
        <xsl:apply-templates select="following-sibling::e:col[1]">
          <xsl:with-param name="number" select="@max + 1"/>
          <xsl:with-param name="sheet" select="$sheet"/>
          <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
          <xsl:with-param name="headerColsStart" select="$headerColsStart"/>
          <xsl:with-param name="headerColsEnd" select="$headerColsEnd"/>
          <xsl:with-param name="GroupCell">
            <xsl:value-of select="$GroupCell"/>
          </xsl:with-param>
          <xsl:with-param name="ManualColBreaks">
            <xsl:value-of select="$ManualColBreaks"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>

      <!-- if next is the first header column -->
      <xsl:when
        test="@max &lt; $headerColsStart and following-sibling::e:col[1][@min &lt;= $headerColsEnd and @max &gt;= $headerColsStart]">
        <xsl:apply-templates select="following-sibling::e:col[1]">
          <xsl:with-param name="number" select="@max + 1"/>
          <xsl:with-param name="sheet" select="$sheet"/>
          <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
          <xsl:with-param name="headerColsStart" select="$headerColsStart"/>
          <xsl:with-param name="headerColsEnd" select="$headerColsEnd"/>
          <xsl:with-param name="GroupCell">
            <xsl:value-of select="$GroupCell"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>

      <!-- jump over the header -->
      <xsl:when
        test="following-sibling::e:col[@min &lt; $headerColsStart or @min &gt; $headerColsEnd]">
        <xsl:apply-templates
          select="following-sibling::e:col[@min &lt; $headerColsStart or @min &gt; $headerColsEnd][1]">
          <xsl:with-param name="number" select="@max + 1"/>
          <xsl:with-param name="sheet" select="$sheet"/>
          <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
          <xsl:with-param name="headerColsStart" select="$headerColsStart"/>
          <xsl:with-param name="headerColsEnd" select="$headerColsEnd"/>
          <xsl:with-param name="GroupCell">
            <xsl:value-of select="$GroupCell"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  
  <xsl:template name="InsertColBreak">
    <xsl:param name="GetMinManualColBreak"/>
    <xsl:param name="ManualColBreaks"/>
    <xsl:param name="sheet"/>
    <xsl:param name="DefaultCellStyleName"/>
    <xsl:param name="prevManualBreak" select="0"/>
    <xsl:choose>
      <xsl:when test="@min &gt; $GetMinManualColBreak and not(preceding-sibling::e:col)">

        <table:table-column>
          <xsl:attribute name="table:style-name">
            <xsl:for-each select="document(concat('xl/',$sheet))">
              <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
            </xsl:for-each>
          </xsl:attribute>

          <xsl:attribute name="table:default-cell-style-name">
            <xsl:value-of select="$DefaultCellStyleName"/>
          </xsl:attribute>

          <xsl:attribute name="table:number-columns-repeated">
            <xsl:value-of select="$GetMinManualColBreak - $prevManualBreak"/>
          </xsl:attribute>
        </table:table-column>

        <table:table-column
          table:style-name="{generate-id(document(concat('xl/',$sheet))/e:worksheet/e:colBreaks)}">
          <xsl:attribute name="table:default-cell-style-name">
            <xsl:value-of select="$DefaultCellStyleName"/>
          </xsl:attribute>
        </table:table-column>

        <xsl:variable name="GetTemplate_GetMinRowWithPicture">
          <xsl:call-template name="GetMinRowWithPicture">
            <xsl:with-param name="PictureRow">
              <xsl:value-of select="$ManualColBreaks"/>
            </xsl:with-param>
            <xsl:with-param name="AfterRow">
              <xsl:value-of select="$GetMinManualColBreak + 1"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>

        <xsl:call-template name="InsertColBreak">
          <xsl:with-param name="ManualColBreaks">
            <xsl:value-of select="$ManualColBreaks"/>
          </xsl:with-param>
          <xsl:with-param name="GetMinManualColBreak">
            <xsl:value-of select="$GetTemplate_GetMinRowWithPicture"/>
          </xsl:with-param>
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="DefaultCellStyleName">
            <xsl:value-of select="$DefaultCellStyleName"/>
          </xsl:with-param>
          <xsl:with-param name="prevManualBreak">
            <xsl:value-of select="$GetMinManualColBreak + 1"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      <xsl:when test="@min &gt; $GetMinManualColBreak and preceding-sibling::e:col">
        <table:table-column>
          <xsl:attribute name="table:style-name">
            <xsl:for-each select="document(concat('xl/',$sheet))">
              <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
            </xsl:for-each>
          </xsl:attribute>
          
          <xsl:attribute name="table:default-cell-style-name">
            <xsl:value-of select="$DefaultCellStyleName"/>
          </xsl:attribute>
          
          <xsl:attribute name="table:number-columns-repeated">
            <xsl:choose>
              <xsl:when test="$prevManualBreak &gt; 0">
                <xsl:value-of select="$GetMinManualColBreak - $prevManualBreak"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$GetMinManualColBreak - preceding::e:col[1]/@max"/>
              </xsl:otherwise>
            </xsl:choose>
            

            
<!-- 
              GetMinManualColBreak<xsl:value-of select="$GetMinManualColBreak"/>
              prevManualBreak<xsl:value-of select="$prevManualBreak"/>
              ManualColBreaks<xsl:value-of select="$ManualColBreaks"/>
              preceding::e:col[1]/@min<xsl:value-of select="preceding::e:col[1]/@min"/>
              preceding::e:col[1]/@max<xsl:value-of select="preceding::e:col[1]/@max"/>
              preceding-sibling::e:col[1]/@min<xsl:value-of select="preceding-sibling::e:col[1]/@min"/>
              preceding-sibling::e:col[1]/@max<xsl:value-of select="preceding-sibling::e:col[1]/@max"/>
-->
            
          </xsl:attribute>

        </table:table-column>
        
        <table:table-column
          table:style-name="{generate-id(document(concat('xl/',$sheet))/e:worksheet/e:colBreaks)}">
          <xsl:attribute name="table:default-cell-style-name">
            <xsl:value-of select="$DefaultCellStyleName"/>
          </xsl:attribute>
        </table:table-column>
        
        <xsl:variable name="GetTemplate_GetMinRowWithPicture">
          <xsl:call-template name="GetMinRowWithPicture">
            <xsl:with-param name="PictureRow">
              <xsl:value-of select="$ManualColBreaks"/>
            </xsl:with-param>
            <xsl:with-param name="AfterRow">
              <xsl:value-of select="$GetMinManualColBreak + 1"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        
        <xsl:call-template name="InsertColBreak">
          <xsl:with-param name="ManualColBreaks">
            <xsl:value-of select="$ManualColBreaks"/>
          </xsl:with-param>
          <xsl:with-param name="GetMinManualColBreak">
            <xsl:value-of select="$GetTemplate_GetMinRowWithPicture"/>
          </xsl:with-param>
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="DefaultCellStyleName">
            <xsl:value-of select="$DefaultCellStyleName"/>
          </xsl:with-param>
          <xsl:with-param name="prevManualBreak">
            <xsl:value-of select="$GetMinManualColBreak + 1"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertColumns">
    <xsl:param name="sheet"/>
    <xsl:param name="GroupCell"/>
    <xsl:param name="ManualColBreaks"/>

    <xsl:variable name="DefaultCellStyleName">
      <xsl:for-each select="document('xl/styles.xml')">
        <xsl:value-of select="generate-id(key('Xf', '')[1])"/>
      </xsl:for-each>
    </xsl:variable>

    <xsl:variable name="sheetName">
      <xsl:value-of select="@name"/>
    </xsl:variable>

    <xsl:variable name="apos">
      <xsl:text>&apos;</xsl:text>
    </xsl:variable>

    <xsl:variable name="charHeaderColsStart">
      <xsl:choose>
        <!-- if sheet name in range definition is in apostrophes -->
        <xsl:when
          test="document('xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($apos,$sheetName,$apos))]">
          <xsl:for-each
            select="document('xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($apos,$sheetName,$apos))]">
            <xsl:choose>
              <!-- when header columns are present -->
              <xsl:when test="contains(text(),',')">
                <xsl:value-of
                  select="substring-before(substring-after(substring-after(substring-before(text(),','),$apos),concat($apos,'!$')),':')"
                />
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of
                  select="substring-before(substring-after(substring-after(text(),$apos),concat($apos,'!$')),':')"
                />
              </xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </xsl:when>

        <xsl:when
          test="document('xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($sheetName,'!'))]">
          <xsl:for-each
            select="document('xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($sheetName,'!'))]">
            <xsl:choose>
              <!-- when header columns are present -->
              <xsl:when test="contains(text(),',')">
                <xsl:value-of
                  select="substring-before(substring-after(substring-before(text(),','),'$'),':')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="substring-before(substring-after(text(),'$'),':')"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="charHeaderColsEnd">
      <xsl:choose>
        <!-- if sheet name in range definition is in apostrophes -->
        <xsl:when
          test="document('xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($apos,$sheetName,$apos))]">
          <xsl:for-each
            select="document('xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($apos,$sheetName,$apos))]">
            <xsl:choose>
              <!-- when header columns are present -->
              <xsl:when test="contains(text(),',')">
                <xsl:value-of
                  select="substring-after(substring-after(substring-before(text(),','),':'),'$')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="substring-after(substring-after(text(),':'),'$')"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </xsl:when>

        <xsl:when
          test="document('xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($sheetName,'!'))]">
          <xsl:for-each
            select="document('xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($sheetName,'!'))]">
            <xsl:choose>
              <!-- when header columns are present -->
              <xsl:when test="contains(text(),',')">
                <xsl:value-of
                  select="substring-after(substring-after(substring-before(text(),','),':'),'$')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="substring-after(substring-after(text(),':'),'$')"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="headerColsStart">
      <xsl:if test="$charHeaderColsStart != '' and not(number($charHeaderColsStart))">
        <xsl:call-template name="GetAlphabeticPosition">
          <xsl:with-param name="literal" select="$charHeaderColsStart"/>
        </xsl:call-template>
      </xsl:if>
    </xsl:variable>

    <xsl:variable name="headerColsEnd">
      <xsl:if test="$charHeaderColsEnd != '' and not(number($charHeaderColsEnd))">
        <xsl:call-template name="GetAlphabeticPosition">
          <xsl:with-param name="literal" select="$charHeaderColsEnd"/>
        </xsl:call-template>
      </xsl:if>
    </xsl:variable>

    <xsl:for-each select="document(concat('xl/',$sheet))/e:worksheet">

      <xsl:apply-templates select="e:cols/e:col[1]">
        <xsl:with-param name="number">1</xsl:with-param>
        <xsl:with-param name="sheet" select="$sheet"/>
        <xsl:with-param name="DefaultCellStyleName" select="$DefaultCellStyleName"/>
        <xsl:with-param name="headerColsStart" select="$headerColsStart"/>
        <xsl:with-param name="headerColsEnd" select="$headerColsEnd"/>
        <xsl:with-param name="GroupCell">
          <xsl:value-of select="$GroupCell"/>
        </xsl:with-param>
        <xsl:with-param name="ManualColBreaks">
          <xsl:value-of select="$ManualColBreaks"/>
        </xsl:with-param>
      </xsl:apply-templates>

      <!-- if all columns are default (there aren't any e:col tags) and there is a header -->
      <xsl:if test="not(e:cols/e:col) and $headerColsStart != '' ">

        <!-- insert columns before header -->
        <xsl:if test="$headerColsStart &gt; 1">
          <table:table-column table:style-name="{generate-id(key('SheetFormatPr', ''))}"
            table:number-columns-repeated="{$headerColsStart - 1}">
            <xsl:attribute name="table:default-cell-style-name">
              <xsl:value-of select="$DefaultCellStyleName"/>
            </xsl:attribute>
          </table:table-column>
        </xsl:if>

        <!-- insert header columns -->
        <table:table-header-columns>
          <table:table-column table:style-name="{generate-id(key('SheetFormatPr', ''))}"
            table:number-columns-repeated="{$headerColsEnd - $headerColsStart + 1}">
            <xsl:attribute name="table:default-cell-style-name">
              <xsl:value-of select="$DefaultCellStyleName"/>
            </xsl:attribute>
          </table:table-column>
        </table:table-header-columns>
      </xsl:if>
    </xsl:for-each>

    <!-- apply default column style for last columns which style wasn't changed -->
    <xsl:for-each select="document(concat('xl/',$sheet))">
      <xsl:choose>
        <xsl:when
          test="$headerColsStart != '' and not(key('Col', '')[@max &gt; $headerColsEnd])">
          <table:table-column table:style-name="{generate-id(key('SheetFormatPr', ''))}"
            table:number-columns-repeated="{256 - $headerColsEnd}">
            <xsl:attribute name="table:default-cell-style-name">
              <xsl:value-of select="$DefaultCellStyleName"/>
            </xsl:attribute>
          </table:table-column>
        </xsl:when>
        <xsl:when test="not(key('Col', ''))">
          <table:table-column table:style-name="{generate-id(key('SheetFormatPr', ''))}"
            table:number-columns-repeated="256">
            <xsl:attribute name="table:default-cell-style-name">
              <xsl:value-of select="$DefaultCellStyleName"/>
            </xsl:attribute>
          </table:table-column>
        </xsl:when>
        <xsl:otherwise>
          <xsl:if test="key('Col', '')[last()]/@max &lt; 256">
            <table:table-column table:style-name="{generate-id(key('SheetFormatPr', ''))}"
              table:number-columns-repeated="{256 - key('Col', '')[last()]/@max}">
              <xsl:attribute name="table:default-cell-style-name">
                <xsl:value-of select="$DefaultCellStyleName"/>
              </xsl:attribute>
            </table:table-column>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="InsertThisColumn">
    <xsl:param name="DefaultCellStyleName"/>
    <xsl:param name="headerColsStart"/>
    <xsl:param name="headerColsEnd"/>
    <xsl:param name="beforeHeader" select="'false'"/>
    <xsl:param name="afterHeader" select="'false'"/>
    <xsl:param name="outlineLevel"/>
    <xsl:param name="GroupCell"/>
    <xsl:param name="ManualColBreaks"/>


    <!-- Insert Group Start -->
    <xsl:if test="contains(concat(';', $GroupCell), concat(';', @min, ':'))">
      <xsl:call-template name="InsertColumnGroupStart">
        <xsl:with-param name="GroupCell">
          <xsl:value-of select="$GroupCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if>

    <table:table-column table:style-name="{generate-id(.)}">

      <xsl:choose>
        <!-- when this is the rest of a column range after header -->
        <xsl:when test="$afterHeader = 'true' ">
          <xsl:attribute name="table:number-columns-repeated">
            <xsl:value-of select="@max - $headerColsEnd"/>
          </xsl:attribute>
        </xsl:when>
        <!-- when this is the part of a column range before header -->
        <xsl:when test="$beforeHeader = 'true'">
          <xsl:attribute name="table:number-columns-repeated">
            <xsl:value-of select="$headerColsStart - @min"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="@min = $headerColsEnd"/>
        <!-- when this column range starts before header and ends after -->
        <xsl:when
          test="$headerColsStart != '' and (@min &lt; $headerColsStart and @max &gt; $headerColsEnd)">
          <xsl:attribute name="table:number-columns-repeated">
            <xsl:value-of select="$headerColsEnd - $headerColsStart + 1"/>
          </xsl:attribute>
        </xsl:when>
        <!-- when this column range starts before header and ends inside -->
        <xsl:when
          test="$headerColsStart != '' and @min &lt; $headerColsStart and @max &gt;= $headerColsStart">
          <xsl:attribute name="table:number-columns-repeated">
            <xsl:value-of select="@max - $headerColsStart + 1"/>
          </xsl:attribute>
        </xsl:when>
        <!-- when this column range starts inside header and ends outside -->
        <xsl:when
          test="$headerColsEnd != '' and @min &lt; $headerColsEnd and @max &gt; $headerColsEnd">
          <xsl:attribute name="table:number-columns-repeated">
            <xsl:value-of select="$headerColsEnd - @min + 1"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:if test="not(@min = @max)">
            <xsl:attribute name="table:number-columns-repeated">
              <xsl:choose>
                <xsl:when test="@max &gt; 256">
                  <xsl:value-of select="256 - @min + 1"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="@max - @min + 1"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:attribute name="table:default-cell-style-name">
        <xsl:value-of select="$DefaultCellStyleName"/>
      </xsl:attribute>
      <xsl:if test="@hidden=1">
        <xsl:attribute name="table:visibility">
          <xsl:text>collapse</xsl:text>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@style">
        <xsl:variable name="position">
          <xsl:value-of select="@style + 1"/>
        </xsl:variable>
        <xsl:attribute name="table:default-cell-style-name">
          <xsl:for-each select="document('xl/styles.xml')">
            <xsl:value-of select="generate-id(key('Xf', '')[position() = $position])"/>
          </xsl:for-each>
        </xsl:attribute>
      </xsl:if>
    </table:table-column>
    
    <!-- Insert Group End -->
    
    <xsl:if test="contains(concat(';', $GroupCell), concat(':', @max, ';'))">
      <xsl:call-template name="InsertColumnGroupEnd">
        <xsl:with-param name="GroupCell">
          <xsl:value-of select="$GroupCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if>
    
  </xsl:template>
  
  <xsl:template name="GetMaxValueBetweenTwoValues">
    <xsl:param name="max"/>
    <xsl:param name="min"/>
    <xsl:param name="value"/>
    <xsl:param name="result" select="0"/>

    <xsl:variable name="FirstValue">
      <xsl:value-of select="substring-before($value, ';')"/>
    </xsl:variable>
    
    <xsl:choose>
      <xsl:when test="$value = ''">
        <xsl:value-of select="$result"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="GetMaxValueBetweenTwoValues">
          <xsl:with-param name="min">
            <xsl:value-of select="$min"/>
          </xsl:with-param>
          <xsl:with-param name="max">
            <xsl:value-of select="$max"/>
          </xsl:with-param>
          <xsl:with-param name="value">
            <xsl:value-of select="substring-after($value, ';')"/>
          </xsl:with-param>
          <xsl:with-param name="result">
            <xsl:choose>
              <xsl:when test="($FirstValue &gt; $result) and ($FirstValue &gt;= $min) and ($FirstValue &lt;= $max)">
                <xsl:value-of select="$FirstValue"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$result"/>
              </xsl:otherwise>
            </xsl:choose>    
          </xsl:with-param>              
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
    
    
    
    
    
    
  </xsl:template>

</xsl:stylesheet>

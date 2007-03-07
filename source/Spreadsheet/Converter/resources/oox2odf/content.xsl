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
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main" exclude-result-prefixes="e r">

  <xsl:import href="relationships.xsl"/>
  <xsl:import href="measures.xsl"/>
  <xsl:import href="styles.xsl"/>
  
  <xsl:key name="Xf" match="e:styleSheet/e:cellXfs/e:xf" use="''"/>
  <xsl:key name="Sst" match="e:si" use="''"/>
  <xsl:key name="SheetFormatPr" match="e:sheetFormatPr" use="''"/>
  <xsl:key name="Col" match="e:col" use="''"/>

  <xsl:template name="content">
    <office:document-content>
      <office:scripts/>
      <office:font-face-decls>
        <xsl:call-template name="InsertFonts"/>
      </office:font-face-decls>
      <office:automatic-styles>
        <xsl:call-template name="InsertColumnStyles"/>
        <xsl:call-template name="InsertRowStyles"/>
        <xsl:call-template name="InsertCellStyles"/>
        <xsl:call-template name="InsertStyleTableProperties"/>
        <xsl:call-template name="InsertTextStyles"/>
      </office:automatic-styles>
      <xsl:call-template name="InsertSheets"/>
    </office:document-content>
  </xsl:template>

  <xsl:template name="InsertSheets">

    <office:body>
      <office:spreadsheet>
        <xsl:for-each select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet">
          <table:table>

            <!-- Insert Table (Sheet) Name -->

            <xsl:attribute name="table:name">
              <xsl:value-of select="@name"/>
            </xsl:attribute>

            <!-- Insert Table Style Name (style:table-properties) -->

            <xsl:attribute name="table:style-name">
              <xsl:value-of select="generate-id()"/>
            </xsl:attribute>

            <xsl:call-template name="InsertSheetContent">
              <xsl:with-param name="sheet">
                <xsl:call-template name="GetTarget">
                  <xsl:with-param name="id">
                    <xsl:value-of select="@r:id"/>
                  </xsl:with-param>
                  <xsl:with-param name="document">xl/workbook.xml</xsl:with-param>
                </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </table:table>
        </xsl:for-each>
      </office:spreadsheet>
    </office:body>
  </xsl:template>

  <xsl:template name="InsertSheetContent">
    <xsl:param name="sheet"/>    
    

    <xsl:call-template name="InsertColumns">
      <xsl:with-param name="sheet" select="$sheet"/>
    </xsl:call-template>

    <xsl:for-each select="document(concat('xl/',$sheet))">     
      

      <xsl:apply-templates select="e:worksheet/e:sheetData/e:row"/>
    
    <xsl:choose>
      <!-- when sheet is empty -->
      <xsl:when test="not(e:worksheet/e:sheetData/e:row/e:c/e:v)">
        <table:table-row
          table:style-name="{generate-id(key('SheetFormatPr', ''))}"
          table:number-rows-repeated="65536">
          <table:table-cell/>
        </table:table-row>
      </xsl:when>
      <xsl:otherwise>
        <!-- it is necessary when sheet has different default row height -->
        <xsl:if
          test="65536 - e:worksheet/e:sheetData/e:row[last()]/@r &gt; 0">
          <table:table-row
            table:style-name="{generate-id(key('SheetFormatPr', ''))}"
            table:number-rows-repeated="{65536 - e:worksheet/e:sheetData/e:row[last()]/@r}">
            <table:table-cell table:number-columns-repeated="256"/>
          </table:table-row>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
    </xsl:for-each>
    
  </xsl:template>

  <xsl:template match="e:row">

    <xsl:variable name="this" select="."/>

    <xsl:variable name="lastCellColumnNumber">
      <xsl:choose>
        <xsl:when test="e:c[last()]/@r">
          <xsl:call-template name="GetColNum">
            <xsl:with-param name="cell">
              <xsl:value-of select="e:c[last()]/@r"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>1</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <!-- if there were empty rows before this one then insert empty rows -->
    <xsl:choose>
      <!-- when this row is the first non-empty one but not row 1 -->
      <xsl:when test="position()=1 and @r>1">
        <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
          table:number-rows-repeated="{@r - 1}">
          <table:table-cell table:number-columns-repeated="256"/>
        </table:table-row>
      </xsl:when>
      <xsl:otherwise>
        <!-- when this row is not first one and there were empty rows after previous non-empty row -->
        <xsl:if test="preceding::e:row[1]/@r &lt;  @r - 1">
          <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}">
            <xsl:attribute name="table:number-rows-repeated">
              <xsl:value-of select="@r -1 - preceding::e:row[1]/@r"/>
            </xsl:attribute>
            <table:table-cell table:number-columns-repeated="256"/>
          </table:table-row>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>

    <!-- insert this row -->
    <table:table-row>
      <xsl:attribute name="table:style-name">
        <xsl:choose>
          <xsl:when test="@ht">
            <xsl:value-of select="generate-id(.)"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:if test="@hidden=1">
        <xsl:attribute name="table:visibility">
          <xsl:text>collapse</xsl:text>
        </xsl:attribute>
      </xsl:if>

      <!-- Insert First Cell in Row  -->
      <xsl:apply-templates select="e:c[1]"/>        
      

      <!-- complete row with empty cells if last cell number < 256 -->
      <xsl:if test="$lastCellColumnNumber &lt; 256">
        <table:table-cell table:number-columns-repeated="{256 - $lastCellColumnNumber}">
          <!-- if there is a default cell style for the row -->
          <xsl:if test="@s">
            <xsl:attribute name="table:style-name">
              <xsl:for-each select="document('xl/styles.xml')">
              <xsl:value-of
                select="generate-id(key('Xf', '')[position() = $this/@s + 1])"
              />
              </xsl:for-each>
            </xsl:attribute>
          </xsl:if>
        </table:table-cell>
      </xsl:if>
    </table:table-row>
  </xsl:template>

  <xsl:template match="e:c">
    <xsl:param name="BeforeMerge"/>
    <xsl:param name="prevCellCol"/>

    <xsl:variable name="this" select="."/>

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

    <xsl:variable name="CheckIfMerge">
      <xsl:call-template name="CheckIfMerge">
        <xsl:with-param name="colNum">
          <xsl:value-of select="$colNum"/>
        </xsl:with-param>
        <xsl:with-param name="rowNum">
          <xsl:value-of select="$rowNum"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!-- if there were empty cells in a row before this one then insert empty cells-->
    <xsl:choose>
      <!-- when this cell is the first one in a row but not in column A -->
      <xsl:when test="$prevCellCol = '' and $colNum>1 and $BeforeMerge != 'true'">
        <table:table-cell>
          <xsl:attribute name="table:number-columns-repeated">
            <xsl:value-of select="$colNum - 1"/>
          </xsl:attribute>
          <!-- if there is a default cell style for the row -->
          <xsl:if test="parent::node()/@s">
            <xsl:variable name="position">
              <xsl:value-of select="$this/parent::node()/@s + 1"/>
            </xsl:variable>
            <xsl:attribute name="table:style-name">
              <xsl:for-each select="document('xl/styles.xml')">
              <xsl:value-of
                select="generate-id(key('Xf', '')[position() = $position])"
              />
              </xsl:for-each>
            </xsl:attribute>
          </xsl:if>
        </table:table-cell>
      </xsl:when>
      <!-- when this cell is not first one in a row and there were empty cells after previous non-empty cell -->
      <xsl:when test="$prevCellCol != ''">
        <xsl:variable name="prevCellColNum">
          <xsl:choose>
          <xsl:when
          test="$prevCellCol != ''">
          <xsl:call-template name="GetColNum">
          <xsl:with-param name="cell">
          <xsl:value-of select="$prevCellCol"/>
          </xsl:with-param>
          </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>-1</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:if test="$colNum>$prevCellColNum+1">
        <table:table-cell>
          <xsl:attribute name="table:number-columns-repeated">
            <xsl:value-of select="$colNum - $prevCellColNum - 1"/>
          </xsl:attribute>
          <!-- if there is a default cell style for the row -->
          <xsl:if test="parent::node()/@s">
            <xsl:variable name="position">
              <xsl:value-of select="$this/parent::node()/@s + 1"/>
            </xsl:variable>
            <xsl:attribute name="table:style-name">
              <xsl:for-each select="document('xl/styles.xml')">
                <xsl:value-of
                  select="generate-id(key('Xf', '')[position() = $position])"
                />
              </xsl:for-each>
            </xsl:attribute>
          </xsl:if>
        </table:table-cell>
        </xsl:if>
      </xsl:when>
    </xsl:choose>

    <xsl:choose>
      <!-- Insert covered cell if this is Merge Cell -->
      <xsl:when test="contains($CheckIfMerge,'true')">
        <xsl:choose>
          <xsl:when test="number(substring-after($CheckIfMerge, ':')) &gt; 1">
            <table:covered-table-cell>
              <xsl:attribute name="table:number-columns-repeated">
                <xsl:value-of select="number(substring-after($CheckIfMerge, ':')) - 1"/>
              </xsl:attribute>
            </table:covered-table-cell>
          </xsl:when>
          <xsl:otherwise>
            <table:covered-table-cell/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <xsl:otherwise>

        <!-- insert this one cell-->
        <table:table-cell>

          <!-- Insert "Merge Cell" if "Merge Cell" is starting in this cell -->
          <xsl:if test="$CheckIfMerge != 'false'">
            <xsl:attribute name="table:number-rows-spanned">
              <xsl:value-of select="substring-before($CheckIfMerge, ':')"/>
            </xsl:attribute>
            <xsl:attribute name="table:number-columns-spanned">
              <xsl:value-of select="substring-after($CheckIfMerge, ':')"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="@s">
            <xsl:variable name="position">
              <xsl:value-of select="$this/@s + 1"/>
            </xsl:variable>
            <xsl:attribute name="table:style-name">
              <xsl:for-each select="document('xl/styles.xml')">
                <xsl:value-of
                  select="generate-id(key('Xf', '')[position() = $position])"
                />
              </xsl:for-each>              
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="e:v and not(e:v = '#REF!' or e:v='#DIV/0!' )">
            <xsl:choose>
              <xsl:when test="@t='s'">
                <xsl:attribute name="office:value-type">
                  <xsl:text>string</xsl:text>
                </xsl:attribute>
                <xsl:variable name="id">
                  <xsl:value-of select="e:v"/>
                </xsl:variable>
                <text:p>
                  <xsl:for-each select="document('xl/sharedStrings.xml')/e:sst/e:si[position()=$id+1]">
                    <xsl:apply-templates/>
                  </xsl:for-each>
                </text:p>
              </xsl:when>
              <xsl:otherwise>
                <xsl:attribute name="office:value-type">
                  <xsl:text>float</xsl:text>
                </xsl:attribute>
                <xsl:attribute name="office:value">
                  <xsl:value-of select="e:v"/>
                </xsl:attribute>
                <text:p>
                  <xsl:value-of select="e:v"/>
                </text:p>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:if>
        </table:table-cell>

        <!-- Insert covered cell if Merge Cell is starting-->
        <xsl:if test="$CheckIfMerge != 'false' and substring-after($CheckIfMerge, ':') &gt; 1">
          <table:covered-table-cell>
            <xsl:attribute name="table:number-columns-repeated">
              <xsl:value-of select="number(substring-after($CheckIfMerge, ':')) - 1"/>
            </xsl:attribute>
          </table:covered-table-cell>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>

    <!-- Insert next coll -->

    <xsl:choose>

      <!-- Skips empty coll (in Merge Cell) -->

      <xsl:when test="$CheckIfMerge != 'false'">
        <xsl:choose>
          <xsl:when test="contains($CheckIfMerge,'true') and substring-after($CheckIfMerge, ':') &gt; 1">
            <xsl:if test="following-sibling::e:c[number(substring-after($CheckIfMerge, ':')) - 1]">
              <xsl:apply-templates
                select="following-sibling::e:c[number(substring-after($CheckIfMerge, ':')) - 1]">
                <xsl:with-param name="BeforeMerge">
                  <xsl:text>true</xsl:text>
                </xsl:with-param>
                <xsl:with-param name="prevCellCol">
                  <xsl:value-of select="@r"/>
                </xsl:with-param>
              </xsl:apply-templates>
            </xsl:if>
          </xsl:when>
          <xsl:otherwise>
            <xsl:if test="following-sibling::e:c[number(substring-after($CheckIfMerge, ':'))]">
              <xsl:apply-templates
                select="following-sibling::e:c[number(substring-after($CheckIfMerge, ':'))]">
                <xsl:with-param name="BeforeMerge">
                  <xsl:text>true</xsl:text>
                </xsl:with-param>
                <xsl:with-param name="prevCellCol">
                  <xsl:value-of select="@r"/>
                </xsl:with-param>
              </xsl:apply-templates>
            </xsl:if>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <xsl:otherwise>
        <xsl:if test="following-sibling::e:c">
          <xsl:apply-templates select="following-sibling::e:c[1]">
            <xsl:with-param name="prevCellCol">
              <xsl:value-of select="@r"/>
            </xsl:with-param>
          </xsl:apply-templates>
        </xsl:if>
      </xsl:otherwise>

    </xsl:choose>

  </xsl:template>

  <!-- convert run into span -->
  <xsl:template match="e:r">
    <text:span>
      <xsl:if test="e:rPr">
        <xsl:attribute name="text:style-name">
          <xsl:value-of select="generate-id(.)"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:apply-templates/>
    </text:span>
  </xsl:template>

  <xsl:template match="e:t">
    <xsl:choose>
      <!--check whether string contains  whitespace sequence-->
      <xsl:when test="not(contains(., '  '))">
        <xsl:value-of select="."/>
      </xsl:when>
      <xsl:otherwise>
        <!--converts whitespaces sequence to text:s-->
        <xsl:call-template name="InsertWhiteSpaces"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--  convert multiple white spaces  -->
  <xsl:template name="InsertWhiteSpaces">
    <xsl:param name="string" select="."/>
    <xsl:param name="length" select="string-length(.)"/>
    <!-- string which doesn't contain whitespaces-->
    <xsl:choose>
      <xsl:when test="not(contains($string,' '))">
        <xsl:value-of select="$string"/>
      </xsl:when>
      <!-- convert white spaces  -->
      <xsl:otherwise>
        <xsl:variable name="before">
          <xsl:value-of select="substring-before($string,' ')"/>
        </xsl:variable>
        <xsl:variable name="after">
          <xsl:call-template name="CutStartSpaces">
            <xsl:with-param name="cuted">
              <xsl:value-of select="substring-after($string,' ')"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:if test="$before != '' ">
          <xsl:value-of select="concat($before,' ')"/>
        </xsl:if>
        <!--add remaining whitespaces as text:s if there are any-->
        <xsl:if test="string-length(concat($before,' ', $after)) &lt; $length ">
          <xsl:choose>
            <xsl:when test="($length - string-length(concat($before, $after))) = 1">
              <text:s/>
            </xsl:when>
            <xsl:otherwise>
              <text:s>
                <xsl:attribute name="text:c">
                  <xsl:choose>
                    <xsl:when test="$before = ''">
                      <xsl:value-of select="$length - string-length($after)"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$length - string-length(concat($before,' ', $after))"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
              </text:s>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
        <!--repeat it for substring which has whitespaces-->
        <xsl:if test="contains($string,' ') and $length &gt; 0">
          <xsl:call-template name="InsertWhiteSpaces">
            <xsl:with-param name="string">
              <xsl:value-of select="$after"/>
            </xsl:with-param>
            <xsl:with-param name="length">
              <xsl:value-of select="string-length($after)"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--  cut start spaces -->
  <xsl:template name="CutStartSpaces">
    <xsl:param name="cuted"/>
    <xsl:choose>
      <xsl:when test="starts-with($cuted,' ')">
        <xsl:call-template name="CutStartSpaces">
          <xsl:with-param name="cuted">
            <xsl:value-of select="substring-after($cuted,' ')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$cuted"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertColumns">
    <xsl:param name="sheet"/>

    <xsl:variable name="DefaultCellStyleName">
      <xsl:for-each select="document('xl/styles.xml')">
        <xsl:value-of select="generate-id(key('Xf', '')[1])"/>
      </xsl:for-each>
    </xsl:variable>
    
    
    <xsl:for-each select="document(concat('xl/',$sheet))/e:worksheet/e:cols/e:col">
      <xsl:variable name="this" select="."/>
      
      <!-- if there were columns with default properties before this column then insert default columns-->
      <xsl:choose>
        <!-- when this column is the first non-default one but it's not the column A -->
        <xsl:when test="position()=1 and @min>1">
          <table:table-column>
            
            <xsl:attribute name="table:style-name">
             <xsl:for-each select="document(concat('xl/',$sheet))">
               <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
             </xsl:for-each>
            </xsl:attribute>
            
            <xsl:attribute name="table:number-columns-repeated">
              <xsl:value-of select="@min - 1"/>
            </xsl:attribute>
            
            <xsl:attribute name="table:default-cell-style-name">
              <xsl:value-of select="$DefaultCellStyleName"/>
            </xsl:attribute>
            <xsl:if test="@style">
              <xsl:variable name="position">
                <xsl:value-of select="$this/@style + 1"/>
              </xsl:variable>
              <xsl:attribute name="table:default-cell-style-name">
                <xsl:for-each select="document('xl/styles.xml')">
                <xsl:value-of
                  select="generate-id(key('Xf', '')[position() = $position])"
                />
               </xsl:for-each>
              </xsl:attribute>
            </xsl:if>
          </table:table-column>
          
        </xsl:when>
        <!-- when this column is not first non-default one and there were default columns after previous non-default column (if there was a gap between this and previous column)-->
        <xsl:when test="preceding::e:col[1]/@max &lt; @min - 1">
          <table:table-column>
            
            <xsl:attribute name="table:style-name">
              <xsl:for-each select="document(concat('xl/',$sheet))">
              <xsl:value-of select="generate-id(key('SheetFormatPr', ''))"/>
              </xsl:for-each>
            </xsl:attribute>
            
            <xsl:attribute name="table:number-columns-repeated">
              <xsl:value-of select="@min - preceding::e:col[1]/@max - 1"/>
            </xsl:attribute>
            
            <xsl:attribute name="table:default-cell-style-name">
              <xsl:value-of select="$DefaultCellStyleName"/>
            </xsl:attribute>
            
          </table:table-column>
        </xsl:when>
      </xsl:choose>

      <!-- insert this column -->
      <table:table-column table:style-name="{generate-id(.)}">
        <xsl:if test="not(@min = @max)">
          <xsl:attribute name="table:number-columns-repeated">
            <xsl:value-of select="@max - @min + 1"/>
          </xsl:attribute>
        </xsl:if>
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
            <xsl:value-of select="$this/@style + 1"/>
          </xsl:variable>
          <xsl:attribute name="table:default-cell-style-name">
            <xsl:for-each select="document('xl/styles.xml')">
            <xsl:value-of
              select="generate-id(key('Xf', '')[position() = $position])"
            />
            </xsl:for-each>
          </xsl:attribute>
        </xsl:if>
      </table:table-column>
    </xsl:for-each>

    <!-- apply default column style for last columns which style wasn't changed -->    
    <xsl:for-each select="document(concat('xl/',$sheet))">
    <xsl:choose>      
      <xsl:when test="not(key('Col', ''))">
        <table:table-column
          table:style-name="{generate-id(key('SheetFormatPr', ''))}"
          table:number-columns-repeated="256">
          <xsl:attribute name="table:default-cell-style-name">
            <xsl:value-of select="$DefaultCellStyleName"/>
          </xsl:attribute>
        </table:table-column>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="key('Col', '')[last()]/@max &lt; 256">
          <table:table-column
            table:style-name="{generate-id(key('SheetFormatPr', ''))}"
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

  <xsl:template name="ConvertFromCharacters">
    <xsl:param name="value"/>

    <!-- strange but true: the best result is when you WON'T convert average digit width from pt to px-->
    <xsl:variable name="defaultFontSize">
      <xsl:for-each select="document('xl/styles.xml')">
      <xsl:choose>
        <xsl:when test="e:styleSheet/e:fonts/e:font">
          <xsl:value-of select="e:styleSheet/e:fonts/e:font[1]/e:sz/@val"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:text>11</xsl:text>
        </xsl:otherwise>
      </xsl:choose>
      </xsl:for-each>
    </xsl:variable>

    <!-- for proportional fonts average digit width is 2/3 of font size-->
    <xsl:variable name="avgDigitWidth">
      <xsl:value-of select="round($defaultFontSize * 2 div 3)"/>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="$avgDigitWidth * $value = 0">
        <xsl:text>0cm</xsl:text>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="ConvertToCentimeters">
          <xsl:with-param name="length">
            <xsl:value-of select="concat(($avgDigitWidth * $value),'px')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

</xsl:stylesheet>

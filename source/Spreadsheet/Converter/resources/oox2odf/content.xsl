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
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main" exclude-result-prefixes="e r">

  <xsl:import href="relationships.xsl"/>
  <xsl:import href="measures.xsl"/>
  <xsl:import href="styles.xsl"/>
  <xsl:import href="table_body.xsl"/>

  <xsl:key name="numFmtId" match="e:styleSheet/e:numFmts/e:numFmt" use="@numFmtId"/>
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
        <xsl:call-template name="InsertNumberStyles"/>
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
        
        <!-- insert strings from sharedStrings to be moved later by post-processor-->
        <xsl:for-each select="document('xl/sharedStrings.xml')/e:sst">
          <pxsi:sst xmlns:pxsi="urn:cleverage:xmlns:post-processings:shared-strings">
            <xsl:apply-templates select="e:si"/>
          </pxsi:sst>
        </xsl:for-each>
    

        <xsl:apply-templates select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet[1]">
          <xsl:with-param name="number">1</xsl:with-param>
        </xsl:apply-templates>
        
      </office:spreadsheet>
    </office:body>
  </xsl:template>
  
  <xsl:template match="e:sheet">
    <xsl:param name="number"/>

       <xsl:variable name="Id">
            <xsl:call-template name="GetTarget">
              <xsl:with-param name="id">
                <xsl:value-of select="@r:id"/>
              </xsl:with-param>
              <xsl:with-param name="document">xl/workbook.xml</xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          
          <xsl:variable name="BigMergeCell">
            <xsl:for-each select="document(concat('xl/',$Id))/e:worksheet/e:mergeCells">
              <xsl:apply-templates select="e:mergeCell[1]" mode="BigMergeColl"/>
            </xsl:for-each>           
          </xsl:variable>
    
        <xsl:variable name="BigMergeRow">
          <xsl:for-each select="document(concat('xl/',$Id))/e:worksheet/e:mergeCells">
            <xsl:apply-templates select="e:mergeCell[1]" mode="BigMergeRow"/>
          </xsl:for-each>           
        </xsl:variable>

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
                <xsl:value-of select="$Id"/>
         </xsl:with-param>
              <xsl:with-param name="BigMergeCell">
                <xsl:value-of select="$BigMergeCell"/>
              </xsl:with-param>
        <xsl:with-param name="BigMergeRow">
          <xsl:value-of select="$BigMergeRow"/>
        </xsl:with-param>
      </xsl:call-template>
      
    </table:table>
    
    <!-- Insert next Table -->
    <xsl:choose>
      <xsl:when test="$number &gt; 255">
        <xsl:message terminate="no">translation.oox2odf.SheetNumber</xsl:message>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates select="following-sibling::e:sheet[1]">
          <xsl:with-param name="number">
            <xsl:value-of select="$number + 1"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:otherwise>
    </xsl:choose>
    
  </xsl:template>

  <!-- insert string -->
  <xsl:template match="e:si">
    <pxsi:si  pxsi:number="{count(preceding-sibling::e:si)}" xmlns:pxsi="urn:cleverage:xmlns:post-processings:shared-strings">
      <xsl:apply-templates/>
    </pxsi:si>
  </xsl:template>
  
  <xsl:template name="InsertSheetContent">
    <xsl:param name="sheet"/>
    <xsl:param name="BigMergeCell"/>
    <xsl:param name="BigMergeRow"/>


    <xsl:call-template name="InsertColumns">
      <xsl:with-param name="sheet" select="$sheet"/>
    </xsl:call-template>

    <xsl:for-each select="document(concat('xl/',$sheet))">

      <!-- Insert Row  -->
      <xsl:apply-templates select="e:worksheet/e:sheetData/e:row[@r &lt; 65537]">
        <xsl:with-param name="BigMergeCell">
          <xsl:value-of select="$BigMergeCell"/>
        </xsl:with-param>
        <xsl:with-param name="BigMergeRow">
          <xsl:value-of select="$BigMergeRow"/>
        </xsl:with-param>
      </xsl:apply-templates>
      
      <!-- OpenOffice calc supports only 65536 rows -->  
      <xsl:if test="e:worksheet/e:sheetData/e:row[@r &gt; 65536]">
        <xsl:message terminate="no">translation.oox2odf.RowNumber</xsl:message>
      </xsl:if>
      
      <!-- Insert sheet without text -->
      <xsl:call-template name="InsertEmptySheet">
        <xsl:with-param name="sheet">
          <xsl:value-of select="$sheet"/>
        </xsl:with-param>
        <xsl:with-param name="BigMergeCell">
          <xsl:value-of select="$BigMergeCell"/>
        </xsl:with-param>
        <xsl:with-param name="BigMergeRow">
          <xsl:value-of select="$BigMergeRow"/>
        </xsl:with-param>
        <xsl:with-param name="RowNumber">
          <xsl:value-of select="e:worksheet/e:sheetData/e:row[position() = last()]/@r"/>    
        </xsl:with-param>
      </xsl:call-template>
      
    </xsl:for-each>

  </xsl:template>

  

  <xsl:template match="e:row">
    <xsl:param name="BigMergeCell"/>
    <xsl:param name="BigMergeRow"/>

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
    
    <xsl:variable name="CheckIfBigMerge">
      <xsl:call-template name="CheckIfBigMergeRow">
        <xsl:with-param name="RowNum">
          <xsl:value-of select="@r"/>
        </xsl:with-param>
        <xsl:with-param name="BigMergeCellRow">
          <xsl:value-of select="$BigMergeRow"/>
        </xsl:with-param>        
      </xsl:call-template>
    </xsl:variable>


      <!-- if there were empty rows before this one then insert empty rows -->
    <xsl:choose>
      <!-- when this row is the first non-empty one but not row 1 and there aren't Big Merge Coll-->
      <xsl:when test="position()=1 and @r>1 and $BigMergeCell = ''">
            <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
              table:number-rows-repeated="{@r - 1}">
              <table:table-cell table:number-columns-repeated="256"/>
            </table:table-row>
      </xsl:when>

      <!-- when this row is the first non-empty one but not row 1 and there aren't Big Merge Coll-->
      <xsl:when test="position()=1 and @r>1 and $BigMergeCell != ''">
       
        <xsl:call-template name="InsertColumnsBigMergeRow">
          <xsl:with-param name="RowNumber">
            <xsl:value-of select="1"/>
          </xsl:with-param>
          <xsl:with-param name="Repeat">
            <xsl:value-of select="@r - 1"/>
          </xsl:with-param>     
          <xsl:with-param name="BigMergeCell">
            <xsl:value-of select="$BigMergeCell"/>
          </xsl:with-param>
        </xsl:call-template>
        
      </xsl:when>
      <xsl:otherwise>
        <!-- when this row is not first one and there were empty rows after previous non-empty row -->
        <xsl:if test="preceding::e:row[1]/@r &lt;  @r - 1">
              <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
                table:number-rows-repeated="{@r -1 - preceding::e:row[1]/@r}">
                <table:table-cell table:number-columns-repeated="256"/>
              </table:table-row>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
    
    <!-- insert this row -->
    <xsl:call-template name="InsertThisRow">
      <xsl:with-param name="BigMergeCell">
        <xsl:value-of select="$BigMergeCell"/>
      </xsl:with-param>
      <xsl:with-param name="BigMergeRow">
        <xsl:value-of select="$BigMergeRow"/>
      </xsl:with-param>
      <xsl:with-param name="lastCellColumnNumber">
        <xsl:value-of select="$lastCellColumnNumber"/>
      </xsl:with-param>
      <xsl:with-param name="CheckIfBigMerge">
        <xsl:value-of select="$CheckIfBigMerge"/>
      </xsl:with-param>
      <xsl:with-param name="this" select="$this"/>
    </xsl:call-template>

</xsl:template>
  
  <xsl:template match="e:c">
    <xsl:param name="BeforeMerge"/>
    <xsl:param name="prevCellCol"/>
    <xsl:param name="BigMergeCell"/>

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
    
    
   <!-- Insert Empty Cell -->
    <xsl:call-template name="InsertEmptyCell">
      <xsl:with-param name="BeforeMerge">
        <xsl:value-of select="$BeforeMerge"/>
      </xsl:with-param>
      <xsl:with-param name="BigMergeCell">
        <xsl:value-of select="$BigMergeCell"/>
      </xsl:with-param>
      <xsl:with-param name="prevCellCol">
        <xsl:value-of select="$prevCellCol"/>
      </xsl:with-param>
      <xsl:with-param name="this" select="$this"/>
      <xsl:with-param name="colNum">
        <xsl:value-of select="$colNum"/>
      </xsl:with-param>
      <xsl:with-param name="rowNum">
        <xsl:value-of select="$rowNum"/>
      </xsl:with-param>
      <xsl:with-param name="CheckIfMerge">
        <xsl:value-of select="$CheckIfMerge"/>
      </xsl:with-param>
    </xsl:call-template>
      
    <!-- Insert this cell or covered cell if this one is Merge Cell -->
    <xsl:call-template name="InsertThisCell">
      <xsl:with-param name="BeforeMerge">
        <xsl:value-of select="$BeforeMerge"/>
      </xsl:with-param>
      <xsl:with-param name="BigMergeCell">
        <xsl:value-of select="$BigMergeCell"/>
      </xsl:with-param>
      <xsl:with-param name="prevCellCol">
        <xsl:value-of select="$prevCellCol"/>
      </xsl:with-param>
      <xsl:with-param name="this" select="$this"/>
      <xsl:with-param name="colNum">
        <xsl:value-of select="$colNum"/>
      </xsl:with-param>
      <xsl:with-param name="rowNum">
        <xsl:value-of select="$rowNum"/>
      </xsl:with-param>
      <xsl:with-param name="CheckIfMerge">
        <xsl:value-of select="$CheckIfMerge"/>
      </xsl:with-param>
    </xsl:call-template>
    

    <!-- Insert next coll -->
    <xsl:call-template name="InsertNextCell">
      <xsl:with-param name="BeforeMerge">
        <xsl:value-of select="$BeforeMerge"/>
      </xsl:with-param>
      <xsl:with-param name="BigMergeCell">
        <xsl:value-of select="$BigMergeCell"/>
      </xsl:with-param>
      <xsl:with-param name="prevCellCol">
        <xsl:value-of select="$prevCellCol"/>
      </xsl:with-param>
      <xsl:with-param name="this" select="$this"/>
      <xsl:with-param name="colNum">
        <xsl:value-of select="$colNum"/>
      </xsl:with-param>
      <xsl:with-param name="rowNum">
        <xsl:value-of select="$rowNum"/>
      </xsl:with-param>
      <xsl:with-param name="CheckIfMerge">
        <xsl:value-of select="$CheckIfMerge"/>
      </xsl:with-param>
    </xsl:call-template>
   
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
  
  <xsl:template match="e:col">
    <xsl:param name="number"/>
    <xsl:param name="sheet"/>
    <xsl:param name="DefaultCellStyleName"/>
    <xsl:variable name="this" select="."/>
    
    <!-- Insert Column Style   -->    
    <xsl:call-template name="InsertColumnStyle">
      <xsl:with-param name="sheet">
        <xsl:value-of select="$sheet"/>
      </xsl:with-param>
      <xsl:with-param name="DefaultCellStyleName">
        <xsl:value-of select="$DefaultCellStyleName"/>
      </xsl:with-param>
      <xsl:with-param name="number">
        <xsl:value-of select="$number"/>
      </xsl:with-param>
      <xsl:with-param name="this" select="$this"/>
    </xsl:call-template>
  
  </xsl:template>
  

  
</xsl:stylesheet>

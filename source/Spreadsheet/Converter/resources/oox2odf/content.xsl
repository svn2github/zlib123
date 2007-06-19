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
  <xsl:import href="number.xsl"/>
  <xsl:import href="picture.xsl"/>
  <xsl:import href="note.xsl"/>
  <xsl:import href="conditional.xsl"/>
  
  <xsl:key name="numFmtId" match="e:styleSheet/e:numFmts/e:numFmt" use="@numFmtId"/>
  <xsl:key name="Xf" match="e:styleSheet/e:cellXfs/e:xf" use="''"/>
  <xsl:key name="Dxf" match="e:styleSheet/e:dxfs/e:dxf" use="''"/>
  <xsl:key name="Sst" match="e:si" use="''"/>
  <xsl:key name="SheetFormatPr" match="e:sheetFormatPr" use="''"/>
  <xsl:key name="Col" match="e:col" use="''"/>
  <xsl:key name="ConditionalFormatting" match="e:conditionalFormatting" use="''"/>

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
        <!-- Insert Picture properties -->
        <xsl:apply-templates select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet[1]"
          mode="PictureStyle">
          <xsl:with-param name="number">1</xsl:with-param>
        </xsl:apply-templates>
        <!-- Insert Conditional Properties -->
        <xsl:apply-templates select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet[1]" mode="ConditionalStyle">
          <xsl:with-param name="number">1</xsl:with-param>
        </xsl:apply-templates>
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
    

    
    <!-- Check If Picture are in this sheet  -->
    <xsl:variable name="PictureCell">
      <xsl:for-each select="document(concat('xl/',$Id))">
      <xsl:call-template name="PictureCell">
        <xsl:with-param name="sheet">
          <xsl:value-of select="substring-after($Id, '/')"/>
        </xsl:with-param>
      </xsl:call-template>
      </xsl:for-each>
    </xsl:variable>
    
    <!-- Check if notes are in this sheet -->
    <xsl:variable name="NoteCell">
      <xsl:for-each select="document(concat('xl/',$Id))">
      <xsl:call-template name="NoteCell">
        <xsl:with-param name="sheetNr" select="$number"/>
      </xsl:call-template>
     </xsl:for-each>
    </xsl:variable>
    
    <!-- Check If Conditionals are in this sheet -->
    
    <xsl:variable name="ConditionalCell">
      <xsl:for-each select="document(concat('xl/',$Id))">
      <xsl:call-template name="ConditionalCell"/>
      </xsl:for-each>
    </xsl:variable>
    
    <xsl:variable name="ConditionalCellStyle">
      <xsl:for-each select="document(concat('xl/',$Id))">
      <xsl:call-template name="ConditionalCell">
        <xsl:with-param name="document">
          <xsl:text>style</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
      </xsl:for-each>
    </xsl:variable>
    
    <xsl:variable name="PictureAndNoteCell">
      <xsl:value-of select="concat($PictureCell,$NoteCell)"/>
    </xsl:variable>
    
    <xsl:variable name="PictureRow">
      <xsl:for-each select="document(concat('xl/',$Id))">
      <xsl:call-template name="PictureRow">
        <xsl:with-param name="PictureCell">
          <xsl:value-of select="$PictureCell"/>
        </xsl:with-param>
      </xsl:call-template>
       </xsl:for-each>
    </xsl:variable>
    
    <xsl:variable name="NoteRow">
      <xsl:for-each select="document(concat('xl/',$Id))">
      <xsl:call-template name="NoteRow">
        <xsl:with-param name="NoteCell">
          <xsl:value-of select="$NoteCell"/>
        </xsl:with-param>
      </xsl:call-template>
      </xsl:for-each>
    </xsl:variable>
    
    <xsl:variable name="PictureAndNoteRow">
      <xsl:value-of select="concat($PictureRow,$NoteRow)"/>
    </xsl:variable>



    <table:table>

      <!-- Insert Table (Sheet) Name -->

      <xsl:variable name="checkedName">
        <xsl:call-template name="CheckSheetName">
          <xsl:with-param name="sheetNumber">
            <xsl:value-of select="$number"/>
          </xsl:with-param>
          <xsl:with-param name="name">
            <xsl:value-of select="translate(@name,'!$-()','')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>

      <xsl:attribute name="table:name">
        <!--        <xsl:value-of select="@name"/>-->
        <xsl:value-of select="$checkedName"/>
      </xsl:attribute>

      <!-- Insert Table Style Name (style:table-properties) -->

      <xsl:attribute name="table:style-name">
        <xsl:value-of select="generate-id()"/>
      </xsl:attribute>

      <xsl:apply-templates
        select="document('xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[1]"
        mode="PrintArea">
        <xsl:with-param name="name">
          <xsl:value-of select="@name"/>
        </xsl:with-param>
        <xsl:with-param name="checkedName">
          <xsl:value-of select="$checkedName"/>
        </xsl:with-param>
      </xsl:apply-templates>

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
        <xsl:with-param name="NameSheet">
          <xsl:value-of select="@name"/>
        </xsl:with-param>
        <xsl:with-param name="sheetNr">
          <xsl:value-of select="$number"/>
        </xsl:with-param>
        <xsl:with-param name="ConditionalCell">
          <xsl:value-of select="$ConditionalCell"/>
        </xsl:with-param>
        <xsl:with-param name="NoteCell">
          <xsl:value-of select="$NoteCell"/>
        </xsl:with-param>
        <xsl:with-param name="NoteRow">
          <xsl:value-of select="$NoteRow"/>
        </xsl:with-param>
        <xsl:with-param name="PictureAndNoteCell">
          <xsl:value-of select="$PictureAndNoteCell"/>
        </xsl:with-param>
        <xsl:with-param name="PictureAndNoteRow">
          <xsl:value-of select="$PictureAndNoteRow"/>
        </xsl:with-param>
        <xsl:with-param name="ConditionalCellStyle">
          <xsl:value-of select="$ConditionalCellStyle"/>
        </xsl:with-param>
        <xsl:with-param name="PictureCell">
          <xsl:value-of select="$PictureCell"/>
        </xsl:with-param>
        <xsl:with-param name="PictureRow">
          <xsl:value-of select="$PictureRow"/>
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

  <xsl:template match="e:definedName" mode="PrintArea">
    <xsl:param name="name"/>
    <xsl:param name="checkedName"/>

    <xsl:variable name="value">
      <xsl:value-of select="."/>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="contains($value, $name) and @name = '_xlnm.Print_Area' ">

        <!--         
        <xsl:variable name="apos">
          <xsl:text>&apos;</xsl:text>
        </xsl:variable>
        
        <xsl:variable name="sheetName">
          <xsl:choose>
            <xsl:when test="starts-with(text(),$apos)">
              <xsl:value-of select="$apos"/>
              <xsl:value-of select="substring-before(substring-after(text(),$apos),$apos)"/>
              <xsl:value-of select="$apos"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-before(text(),'!')"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
     
        <xsl:variable name="printArea">
          <xsl:value-of select="$sheetName"/>
          <xsl:value-of select="substring-after(text(),$sheetName)"/>
        </xsl:variable>
-->
        <xsl:attribute name="table:print-ranges">
          <xsl:call-template name="InsertRanges">
            <xsl:with-param name="ranges" select="text()"/>
            <xsl:with-param name="mode" select="substring-after(text(),',')"/>
            <xsl:with-param name="checkedName" select="$checkedName"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:when>

      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="following-sibling::e:definedName">
            <xsl:apply-templates select="following-sibling::e:definedName[1]" mode="PrintArea">
              <xsl:with-param name="name">
                <xsl:value-of select="$name"/>
              </xsl:with-param>
              <xsl:with-param name="checkedName">
                <xsl:value-of select="$checkedName"/>
              </xsl:with-param>
            </xsl:apply-templates>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="table:print">
              <xsl:text>false</xsl:text>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertRanges">
    <xsl:param name="ranges"/>
    <xsl:param name="mode"/>
    <xsl:param name="checkedName"/>

    <xsl:variable name="apos">
      <xsl:text>&apos;</xsl:text>
    </xsl:variable>

    <!-- take sheet name from <definedName> (can be inside apostrophes and can be distinct from $checkedName) 
           it is needed for <definedName> processing -->
    <xsl:variable name="sheetName">
      <xsl:choose>
        <xsl:when test="starts-with($ranges,$apos)">
          <xsl:value-of select="$apos"/>
          <xsl:value-of select="substring-before(substring-after($ranges,$apos),$apos)"/>
          <xsl:value-of select="$apos"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="substring-before($ranges,'!')"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:choose>
      <!-- when there are more than one range -->
      <xsl:when test="$mode != '' ">

        <!-- single-cell range can be defined either as Sheet1!$A$2:$A$2 or as Sheet1!$A$2-->
        <xsl:variable name="startRange">
          <xsl:choose>
            <xsl:when
              test="contains(substring-before(substring-after($ranges, concat($sheetName,'!') ),','), ':' )">
              <xsl:value-of
                select="substring-before(substring-after($ranges,concat($sheetName,'!')),':' )"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of
                select="substring-before(substring-after($ranges,concat($sheetName,'!')),',' )"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <!-- single-cell range can be defined either as Sheet1!$A$2:$A$2 or as Sheet1!$A$2-->
        <xsl:variable name="endRange">
          <xsl:choose>
            <xsl:when
              test="contains(substring-before(substring-after($ranges, concat($sheetName,'!') ),','), ':' )">
              <xsl:value-of select="substring-before(substring-after($ranges,':'),',' )"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of
                select="substring-before(substring-after($ranges,concat($sheetName,'!')),',' )"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:variable name="start">
          <xsl:choose>
            <!-- when print range is defined for whole rows -->
            <xsl:when test="number(translate($startRange,'$',''))">
              <xsl:value-of select="concat('A',translate($startRange,'$',''))"/>
            </xsl:when>
            <!-- when print range is defined for whole columns -->
            <xsl:when test="$startRange = translate($startRange,'1234567890','')">
              <xsl:value-of select="concat(translate($startRange,'$',''),'1')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="translate($startRange,'$','')"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:variable name="end">
          <xsl:choose>
            <!-- when print range is defined for whole rows -->
            <xsl:when test="number(translate($endRange,'$',''))">
              <xsl:value-of select="concat('IV',translate($endRange,'$',''))"/>
            </xsl:when>
            <!-- when print range is defined for whole columns -->
            <xsl:when test="$endRange = translate($endRange,'1234567890','')">
              <xsl:value-of select="concat(translate($endRange,'$',''),'65536')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="translate($endRange,'$','')"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:value-of
          select="concat($apos,$checkedName,$apos,'.',$start,':',$apos,$checkedName,$apos,'.',$end)"/>
        <xsl:text> </xsl:text>

        <xsl:call-template name="InsertRanges">
          <xsl:with-param name="ranges" select="substring-after($ranges,',')"/>
          <xsl:with-param name="mode" select="substring-after(substring-after($ranges,','),',')"/>
          <xsl:with-param name="checkedName" select="$checkedName"/>
        </xsl:call-template>
      </xsl:when>

      <!-- this is the last range -->
      <xsl:otherwise>
        <!-- single-cell range can be defined either as Sheet1!$A$2:$A$2 or as Sheet1!$A$2-->
        <xsl:variable name="startRange">
          <xsl:choose>
            <xsl:when test="contains(substring-after($ranges, concat($sheetName,'!') ), ':' )">
              <xsl:value-of
                select="substring-before(substring-after($ranges,concat($sheetName,'!')),':' )"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-after($ranges,concat($sheetName,'!'))"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <!-- single-cell range can be defined either as Sheet1!$A$2:$A$2 or as Sheet1!$A$2-->
        <xsl:variable name="endRange">
          <xsl:choose>
            <xsl:when test="contains(substring-after($ranges, concat($sheetName,'!') ), ':' )">
              <xsl:value-of select="substring-after($ranges,':')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-after($ranges,concat($sheetName,'!'))"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:variable name="start">
          <xsl:choose>
            <!-- when print range is defined for whole rows -->
            <xsl:when test="number(translate($startRange,'$',''))">
              <xsl:value-of select="concat('A',translate($startRange,'$',''))"/>
            </xsl:when>
            <!-- when print range is defined for whole columns -->
            <xsl:when test="$startRange = translate($startRange,'1234567890','')">
              <xsl:value-of select="concat(translate($startRange,'$',''),'1')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="translate($startRange,'$','')"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:variable name="end">
          <xsl:choose>
            <!-- when print range is defined for whole rows -->
            <xsl:when test="number(translate($endRange,'$',''))">
              <xsl:value-of select="concat('IV',translate($endRange,'$',''))"/>
            </xsl:when>
            <!-- when print range is defined for whole columns -->
            <xsl:when test="$endRange = translate($endRange,'1234567890','')">
              <xsl:value-of select="concat(translate($endRange,'$',''),'65536')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="translate($endRange,'$','')"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:value-of
          select="concat($apos,$checkedName,$apos,'.',$start,':',$apos,$checkedName,$apos,'.',$end)"/>

      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- insert string -->
  <xsl:template match="e:si">
    <pxsi:si pxsi:number="{count(preceding-sibling::e:si)}"
      xmlns:pxsi="urn:cleverage:xmlns:post-processings:shared-strings">
      <xsl:apply-templates/>
    </pxsi:si>
  </xsl:template>

  <xsl:template name="InsertSheetContent">
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="BigMergeCell"/>
    <xsl:param name="BigMergeRow"/>
    <xsl:param name="sheetNr"/>
    <xsl:param name="PictureCell"/>
    <xsl:param name="NoteCell"/>
    <xsl:param name="ConditionalCell"/>
    <xsl:param name="ConditionalCellStyle"/>
    <xsl:param name="PictureAndNoteCell"/>
    <xsl:param name="PictureRow"/>
    <xsl:param name="NoteRow"/>  
    <xsl:param name="PictureAndNoteRow"/>
    

    <xsl:call-template name="InsertColumns">
      <xsl:with-param name="sheet" select="$sheet"/>
    </xsl:call-template>

    <xsl:variable name="sheetName">
      <xsl:value-of select="@name"/>
    </xsl:variable>

    <xsl:for-each select="document(concat('xl/',$sheet))">

      <xsl:variable name="apos">
        <xsl:text>&apos;</xsl:text>
      </xsl:variable>

      <xsl:variable name="headerRowsStart">
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
                    select="substring-before(substring-after(substring-after(substring-after(text(),','),$apos),concat($apos,'!$')),':')"
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
                    select="substring-before(substring-after(substring-after(text(),','),'$'),':')"
                  />
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="substring-before(substring-after(text(),'$'),':')"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:for-each>
          </xsl:when>
        </xsl:choose>

      </xsl:variable>

      <xsl:variable name="headerRowsEnd">
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
                    select="substring-after(substring-after(substring-after(text(),','),':'),'$')"/>
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
                    select="substring-after(substring-after(substring-after(text(),','),':'),'$')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="substring-after(substring-after(text(),':'),'$')"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:for-each>
          </xsl:when>
        </xsl:choose>
      </xsl:variable>

      <!-- Insert Row  -->
      <xsl:choose>

        <!-- if there are header rows -->
        <xsl:when test="$headerRowsStart != '' and number($headerRowsStart)">

          <!-- insert rows before header rows -->
          <xsl:apply-templates
            select="e:worksheet/e:sheetData/e:row[@r &lt; $headerRowsStart and @r &lt; 65537]"
            mode="headers">
            <xsl:with-param name="BigMergeCell">
              <xsl:value-of select="$BigMergeCell"/>
            </xsl:with-param>
            <xsl:with-param name="BigMergeRow">
              <xsl:value-of select="$BigMergeRow"/>
            </xsl:with-param>
            <xsl:with-param name="headerRowsStart" select="$headerRowsStart"/>
            <xsl:with-param name="headerRowsEnd" select="$headerRowsEnd"/>
            <xsl:with-param name="sheetNr" select="$sheetNr"/>
          </xsl:apply-templates>

          <!-- insert empty rows before header -->
          <xsl:if
            test="$headerRowsStart &gt; 1 and not(e:worksheet/e:sheetData/e:row[@r = $headerRowsStart - 1 and @r &lt; 65537])">
            <xsl:choose>
              <!-- when there aren't any rows before at all -->
              <xsl:when
                test="not(e:worksheet/e:sheetData/e:row[@r &lt; $headerRowsStart and @r &lt; 65537])">
                <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
                  table:number-rows-repeated="{$headerRowsStart - 1}">
                  <table:table-cell table:number-columns-repeated="256"/>
                </table:table-row>
              </xsl:when>
              <!-- if there was a row before header -->
              <xsl:otherwise>
                <xsl:for-each
                  select="e:worksheet/e:sheetData/e:row[@r &lt; $headerRowsStart and @r &lt; 65537][last()]">
                  <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
                    table:number-rows-repeated="{$headerRowsStart - @r - 1}">
                    <table:table-cell table:number-columns-repeated="256"/>
                  </table:table-row>
                </xsl:for-each>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:if>

          <!-- insert header rows -->
          <table:table-header-rows>
            <xsl:apply-templates
              select="e:worksheet/e:sheetData/e:row[@r &gt;= $headerRowsStart and @r &lt;= $headerRowsEnd and @r &lt; 65537]"
              mode="headers">
              <xsl:with-param name="BigMergeCell">
                <xsl:value-of select="$BigMergeCell"/>
              </xsl:with-param>
              <xsl:with-param name="BigMergeRow">
                <xsl:value-of select="$BigMergeRow"/>
              </xsl:with-param>
              <xsl:with-param name="headerRowsStart" select="$headerRowsStart"/>
              <xsl:with-param name="headerRowsEnd" select="$headerRowsEnd"/>
              <xsl:with-param name="sheetNr" select="$sheetNr"/>
            </xsl:apply-templates>

            <!-- if header is empty -->
            <xsl:choose>
              <xsl:when test="not(e:worksheet/e:sheetData/e:row/e:c/e:v) and $BigMergeCell != '' ">
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
                  <xsl:with-param name="PictureCell">
                    <xsl:value-of select="$PictureCell"/>
                  </xsl:with-param>
                  <xsl:with-param name="NoteCell">
                    <xsl:value-of select="$NoteCell"/>
                  </xsl:with-param>
                  <xsl:with-param name="PictureAndNoteCell">
                    <xsl:value-of select="$PictureAndNoteCell"/>
                  </xsl:with-param>
                  <xsl:with-param name="NameSheet">
                    <xsl:value-of select="$sheetName"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:when>

              <xsl:when
                test="not(e:worksheet/e:sheetData/e:row[@r &gt;= $headerRowsStart and @r &lt;= $headerRowsEnd and @r &lt; 65537])">
                <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
                  table:number-rows-repeated="{$headerRowsEnd - $headerRowsStart + 1}">
                  <table:table-cell table:number-columns-repeated="256"/>
                </table:table-row>
              </xsl:when>
            </xsl:choose>

            <!-- if there are empty rows at the end of the header -->
            <xsl:for-each
              select="e:worksheet/e:sheetData/e:row[@r &gt;= $headerRowsStart and @r &lt;= $headerRowsEnd and @r &lt; 65537][last()]">
              <xsl:if test="@r &lt; $headerRowsEnd">
                <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
                  table:number-rows-repeated="{$headerRowsEnd - @r}">
                  <table:table-cell table:number-columns-repeated="256"/>
                </table:table-row>
              </xsl:if>
            </xsl:for-each>
          </table:table-header-rows>

          <!-- insert rows after header rows -->
          <xsl:apply-templates
            select="e:worksheet/e:sheetData/e:row[@r &gt; $headerRowsEnd and @r &lt; 65537]"
            mode="headers">
            <xsl:with-param name="BigMergeCell">
              <xsl:value-of select="$BigMergeCell"/>
            </xsl:with-param>
            <xsl:with-param name="BigMergeRow">
              <xsl:value-of select="$BigMergeRow"/>
            </xsl:with-param>
            <xsl:with-param name="headerRowsStart" select="$headerRowsStart"/>
            <xsl:with-param name="headerRowsEnd" select="$headerRowsEnd"/>
          </xsl:apply-templates>
        </xsl:when>

        <!-- if there aren't any header rows -->
        <xsl:otherwise>

          <xsl:choose>
            <xsl:when test="e:worksheet/e:sheetData/e:row/e:c/e:v">
              <xsl:apply-templates select="e:worksheet/e:sheetData/e:row[@r &lt; 65537]">
                <xsl:with-param name="BigMergeCell">
                  <xsl:value-of select="$BigMergeCell"/>
                </xsl:with-param>
                <xsl:with-param name="BigMergeRow">
                  <xsl:value-of select="$BigMergeRow"/>
                </xsl:with-param>
                <xsl:with-param name="PictureCell">
                  <xsl:value-of select="$PictureCell"/>
                </xsl:with-param>
                <xsl:with-param name="PictureRow">
                  <xsl:value-of select="$PictureRow"/>
                </xsl:with-param>
                <xsl:with-param name="NoteCell">
                  <xsl:value-of select="$NoteCell"/>
                </xsl:with-param>
                <xsl:with-param name="NoteRow">
                  <xsl:value-of select="$NoteRow"/>
                </xsl:with-param>
                <xsl:with-param name="PictureAndNoteCell">
                  <xsl:value-of select="$PictureAndNoteCell"/>
                </xsl:with-param>
                <xsl:with-param name="PictureAndNoteRow">
                  <xsl:value-of select="$PictureAndNoteRow"/>
                </xsl:with-param>
                <xsl:with-param name="sheet">
                  <xsl:value-of select="$sheet"/>
                </xsl:with-param>
                <xsl:with-param name="NameSheet">
                  <xsl:value-of select="$NameSheet"/>
                </xsl:with-param>
                <xsl:with-param name="sheetNr" select="$sheetNr"/>
                <xsl:with-param name="ConditionalCell">
                  <xsl:value-of select="$ConditionalCell"/>
                </xsl:with-param>
                <xsl:with-param name="ConditionalCellStyle">
                  <xsl:value-of select="$ConditionalCellStyle"/>
                </xsl:with-param>
              </xsl:apply-templates>
            </xsl:when>

            <xsl:otherwise>
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
                <xsl:with-param name="PictureCell">
                  <xsl:value-of select="$PictureCell"/>
                </xsl:with-param>
                <xsl:with-param name="NoteCell">
                  <xsl:value-of select="$NoteCell"/>
                </xsl:with-param>
                <xsl:with-param name="PictureAndNoteCell">
                  <xsl:value-of select="$PictureAndNoteCell"/>
                </xsl:with-param>
                <xsl:with-param name="NameSheet">
                  <xsl:value-of select="$sheetName"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>

        </xsl:otherwise>
      </xsl:choose>

      <!-- OpenOffice calc supports only 65536 rows -->
      <xsl:if test="e:worksheet/e:sheetData/e:row[@r &gt; 65536]">
        <xsl:message terminate="no">translation.oox2odf.RowNumber</xsl:message>
      </xsl:if>

    </xsl:for-each>

  </xsl:template>

  <xsl:template match="e:row">
    <xsl:param name="BigMergeCell"/>
    <xsl:param name="BigMergeRow"/>
    <xsl:param name="PictureCell"/>
    <xsl:param name="PictureRow"/>
    <xsl:param name="NoteCell"/>
    <xsl:param name="NoteRow"/>
    <xsl:param name="PictureAndNoteCell"/>
    <xsl:param name="PictureAndNoteRow"/>
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="sheetNr"/>
    <xsl:param name="ConditionalCell"/>
    <xsl:param name="ConditionalCellStyle"/>

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

    <xsl:variable name="GetMinRowWithPicture">
      <xsl:call-template name="GetMinRowWithPicture">
        <xsl:with-param name="PictureRow">
          <xsl:value-of select="concat($PictureRow, $NoteRow)"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>


    <!-- if there were empty rows before this one then insert empty rows -->
    <xsl:choose>
      <!-- when this row is the first non-empty one but not row 1 and there aren't Big Merge Coll and Pictures-->
      <xsl:when
        test="position()=1 and @r>1 and $BigMergeCell = '' and ($GetMinRowWithPicture = '' or ($GetMinRowWithPicture + 1) &gt;= @r)">

        <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
          table:number-rows-repeated="{@r - 1}">
          <table:table-cell table:number-columns-repeated="256"/>
        </table:table-row>

      </xsl:when>

      <!-- when this row is the first non-empty one but not row 1 and there aren't Big Merge Coll, and there are Pictures before this row-->
      <xsl:when
        test="position()=1 and @r>1 and $BigMergeCell = '' and $PictureRow != '' and ($GetMinRowWithPicture + 1) &lt; @r">

        <xsl:call-template name="InsertPictureBetwenTwoRows">
          <xsl:with-param name="StartRow">
            <xsl:text>0</xsl:text>
          </xsl:with-param>
          <xsl:with-param name="EndRow">
            <xsl:value-of select="@r"/>
          </xsl:with-param>
          <xsl:with-param name="PictureRow">
            <xsl:value-of select="$PictureRow"/>
          </xsl:with-param>
          <xsl:with-param name="PictureCell">
            <xsl:value-of select="$PictureCell"/>
          </xsl:with-param>
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="NameSheet">
            <xsl:value-of select="$NameSheet"/>
          </xsl:with-param>
          <xsl:with-param name="sheetNr" select="$sheetNr"/>
        </xsl:call-template>

      </xsl:when>



      <!-- when this row is the first non-empty one but not row 1 and there aren't Big Merge Coll-->

      <xsl:when test="position()=1 and @r>1 and $BigMergeCell != ''">

        <xsl:call-template name="InsertBigMergeFirstRowEmpty">
          <xsl:with-param name="RowNumber">
            <xsl:value-of select="1"/>
          </xsl:with-param>
          <xsl:with-param name="BigMergeCell">
            <xsl:value-of select="$BigMergeCell"/>
          </xsl:with-param>
        </xsl:call-template>

        <xsl:if test="@r - 2 &gt; 0"/>
        <xsl:call-template name="InsertBigMergeRowEmpty">
          <xsl:with-param name="RowNumber">
            <xsl:value-of select="2"/>
          </xsl:with-param>
          <xsl:with-param name="BigMergeCell">
            <xsl:value-of select="$BigMergeCell"/>
          </xsl:with-param>
          <xsl:with-param name="Repeat">
            <xsl:value-of select="@r - 2"/>
          </xsl:with-param>

        </xsl:call-template>

      </xsl:when>

      <!-- when this row is not first one and there were pictures rows after previous non-empty row-->
      <xsl:when
        test="preceding::e:row[1]/@r &lt;  @r - 1 and ($GetMinRowWithPicture + 1) &gt; preceding::e:row[1]/@r and ($GetMinRowWithPicture + 1) &lt; @r - 1">

        <xsl:call-template name="InsertPictureBetwenTwoRows">
          <xsl:with-param name="StartRow">
            <xsl:value-of select="preceding::e:row[1]/@r"/>
          </xsl:with-param>
          <xsl:with-param name="EndRow">
            <xsl:value-of select="@r"/>
          </xsl:with-param>
          <xsl:with-param name="PictureRow">
            <xsl:value-of select="$PictureRow"/>
          </xsl:with-param>
          <xsl:with-param name="PictureCell">
            <xsl:value-of select="$PictureCell"/>
          </xsl:with-param>
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="NameSheet">
            <xsl:value-of select="$NameSheet"/>
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
      <xsl:with-param name="PictureRow">
        <xsl:value-of select="$PictureRow"/>
      </xsl:with-param>
      <xsl:with-param name="PictureCell">
        <xsl:value-of select="$PictureCell"/>
      </xsl:with-param>
      <xsl:with-param name="NoteRow">
        <xsl:value-of select="$NoteRow"/>
      </xsl:with-param>
      <xsl:with-param name="NoteCell">
        <xsl:value-of select="$NoteCell"/>
      </xsl:with-param>
      <xsl:with-param name="PictureAndNoteCell">
        <xsl:value-of select="$PictureAndNoteCell"/>
      </xsl:with-param>
      <xsl:with-param name="sheet">
        <xsl:value-of select="$sheet"/>
      </xsl:with-param>
      <xsl:with-param name="NameSheet">
        <xsl:value-of select="$NameSheet"/>
      </xsl:with-param>
      <xsl:with-param name="sheetNr" select="$sheetNr"/>
      <xsl:with-param name="ConditionalCell">
        <xsl:value-of select="$ConditionalCell"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellStyle">
        <xsl:value-of select="$ConditionalCellStyle"/>
      </xsl:with-param>
    </xsl:call-template>

    <xsl:if
      test="not(following-sibling::e:row) and $PictureRow != '' and ($GetMinRowWithPicture +1 &gt; @r or (not(e:c) and $GetMinRowWithPicture + 1 = @r))">

      <xsl:call-template name="InsertPictureBetwenTwoRows">
        <xsl:with-param name="StartRow">
          <xsl:value-of select="@r"/>
        </xsl:with-param>
        <xsl:with-param name="EndRow">
          <xsl:value-of select="65536"/>
        </xsl:with-param>
        <xsl:with-param name="PictureRow">
          <xsl:value-of select="$PictureRow"/>
        </xsl:with-param>
        <xsl:with-param name="PictureCell">
          <xsl:value-of select="$PictureCell"/>
        </xsl:with-param>
        <xsl:with-param name="sheet">
          <xsl:value-of select="$sheet"/>
        </xsl:with-param>
        <xsl:with-param name="NameSheet">
          <xsl:value-of select="$NameSheet"/>
        </xsl:with-param>
        <xsl:with-param name="sheetNr">
          <xsl:value-of select="$sheetNr"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if>

    <xsl:if
      test="not(following-sibling::e:row) and $NoteRow != '' and ($GetMinRowWithPicture &gt; @r or (not(e:c) and $GetMinRowWithPicture = @r))">

      <xsl:call-template name="InsertNoteBetwenTwoRows">
        <xsl:with-param name="StartRow">
          <xsl:value-of select="@r"/>
        </xsl:with-param>
        <xsl:with-param name="EndRow">
          <xsl:value-of select="65536"/>
        </xsl:with-param>
        <xsl:with-param name="NoteRow">
          <xsl:value-of select="$NoteRow"/>
        </xsl:with-param>
        <xsl:with-param name="NoteCell">
          <xsl:value-of select="$NoteCell"/>
        </xsl:with-param>
        <xsl:with-param name="sheet">
          <xsl:value-of select="$sheet"/>
        </xsl:with-param>
        <xsl:with-param name="NameSheet">
          <xsl:value-of select="$NameSheet"/>
        </xsl:with-param>
        <xsl:with-param name="sheetNr">
          <xsl:value-of select="$sheetNr"/>
        </xsl:with-param>
      </xsl:call-template>

    </xsl:if>

  </xsl:template>

  <xsl:template match="e:row" mode="headers">
    <xsl:param name="BigMergeCell"/>
    <xsl:param name="BigMergeRow"/>
    <xsl:param name="headerRowsStart"/>
    <xsl:param name="headerRowsEnd"/>
    <xsl:param name="PictureRow"/>
    <xsl:param name="sheetNr"/>

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

    <xsl:choose>
      <!-- when this row is the first non-empty one before header but not row 1 and there aren't Big Merge Coll -->
      <xsl:when
        test="position()=1 and @r > 1 and @r &lt; $headerRowsStart and $BigMergeCell = '' ">
        <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
          table:number-rows-repeated="{@r - 1}">
          <table:table-cell table:number-columns-repeated="256"/>
        </table:table-row>
      </xsl:when>

      <!-- if this is a header row -->
      <xsl:when
        test="$headerRowsStart != ''  and @r &gt;= $headerRowsStart and @r &lt;= $headerRowsEnd">
        <xsl:choose>
          <!-- when this row is the first non-empty one but not row 1 and there are Big Merge Coll-->
          <xsl:when test="position()=1 and @r>1 and $BigMergeCell != '' ">
            <xsl:call-template name="InsertBigMergeFirstRowEmpty">
              <xsl:with-param name="RowNumber">
                <xsl:value-of select="1"/>
              </xsl:with-param>
              <xsl:with-param name="BigMergeCell">
                <xsl:value-of select="$BigMergeCell"/>
              </xsl:with-param>
            </xsl:call-template>
            <xsl:if test="@r - 2 &gt; 0"/>
            <xsl:call-template name="InsertBigMergeRowEmpty">
              <xsl:with-param name="RowNumber">
                <xsl:value-of select="2"/>
              </xsl:with-param>
              <xsl:with-param name="BigMergeCell">
                <xsl:value-of select="$BigMergeCell"/>
              </xsl:with-param>
              <xsl:with-param name="Repeat">
                <xsl:value-of select="@r - 2"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>

          <!-- if the first non-empty header row isn't the first header row -->
          <xsl:when test="position() = 1 and @r &gt; $headerRowsStart">
            <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
              table:number-rows-repeated="{@r - $headerRowsStart}">
              <table:table-cell table:number-columns-repeated="256"/>
            </table:table-row>
          </xsl:when>

          <!-- if there are empty header rows before this one header row-->
          <xsl:when test="position() &gt; 1 and @r - 1 != preceding-sibling::e:row[1]/@r">
            <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
              table:number-rows-repeated="{@r - preceding-sibling::e:row[1]/@r - 1}">
              <table:table-cell table:number-columns-repeated="256"/>
            </table:table-row>
          </xsl:when>
        </xsl:choose>
      </xsl:when>

      <!-- when this row is the first non-empty one after header rows and there aren't Big Merge Coll -->
      <xsl:when test="position()=1 and @r &gt; $headerRowsEnd + 1 and $BigMergeCell = '' ">
        <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
          table:number-rows-repeated="{@r - $headerRowsEnd - 1}">
          <table:table-cell table:number-columns-repeated="256"/>
        </table:table-row>
      </xsl:when>

      <!-- when this row is the first non-empty one but not row 1 and there are Big Merge Coll (this is not a header row) -->
      <xsl:when test="position()=1 and @r>1 and $BigMergeCell != '' ">
        <xsl:call-template name="InsertBigMergeFirstRowEmpty">
          <xsl:with-param name="RowNumber">
            <xsl:value-of select="1"/>
          </xsl:with-param>
          <xsl:with-param name="BigMergeCell">
            <xsl:value-of select="$BigMergeCell"/>
          </xsl:with-param>
        </xsl:call-template>
        <xsl:if test="@r - 2 &gt; 0"/>
        <xsl:call-template name="InsertBigMergeRowEmpty">
          <xsl:with-param name="RowNumber">
            <xsl:value-of select="2"/>
          </xsl:with-param>
          <xsl:with-param name="BigMergeCell">
            <xsl:value-of select="$BigMergeCell"/>
          </xsl:with-param>
          <xsl:with-param name="Repeat">
            <xsl:value-of select="@r - 2"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <!-- when this row is not first one and there were empty rows after previous non-empty row -->
      <xsl:when test="position() != 1 and @r != preceding-sibling::e:row[1]/@r + 1">
        <table:table-row table:style-name="{generate-id(key('SheetFormatPr', ''))}"
          table:number-rows-repeated="{@r -1 - preceding::e:row[1]/@r}">
          <table:table-cell table:number-columns-repeated="256"/>
        </table:table-row>
      </xsl:when>
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
      <xsl:with-param name="headerRowsStart" select="$headerRowsStart"/>
      <xsl:with-param name="sheetNr" select="$sheetNr"/>
    </xsl:call-template>

  </xsl:template>

  <xsl:template match="e:c">
    <xsl:param name="BeforeMerge"/>
    <xsl:param name="prevCellCol"/>
    <xsl:param name="BigMergeCell"/>
    <xsl:param name="PictureCell"/>
    <xsl:param name="PictureRow"/>
    <xsl:param name="NoteCell"/>
    <xsl:param name="NoteRow"/>
    <xsl:param name="PictureAndNoteCell"/>
    <xsl:param name="PictureAndNoteRow"/>
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="sheetNr"/>
    <xsl:param name="ConditionalCell"/>
    <xsl:param name="ConditionalCellStyle"/>

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

    <xsl:variable name="PictureColl">
      <xsl:call-template name="GetCollsWithPicture">
        <xsl:with-param name="rowNumber">
          <xsl:value-of select="$rowNum - 1"/>
        </xsl:with-param>
        <xsl:with-param name="PictureCell">
          <xsl:value-of select="concat(';', $PictureCell)"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="NoteColl">
      <xsl:call-template name="GetCollsWithPicture">
        <xsl:with-param name="rowNumber">
          <xsl:value-of select="$rowNum"/>
        </xsl:with-param>
        <xsl:with-param name="PictureCell">
          <xsl:value-of select="concat(';', $NoteCell)"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="GetMinCollWithPicture">
      <xsl:call-template name="GetMinRowWithPicture">
        <xsl:with-param name="PictureRow">
          <xsl:value-of select="concat($PictureColl, $NoteColl)"/>
        </xsl:with-param>
        <xsl:with-param name="AfterRow">
          <xsl:choose>
            <xsl:when test="$prevCellCol = ''">
              <xsl:text>0</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$prevCellCol - 1"/>
            </xsl:otherwise>
          </xsl:choose>
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
      <xsl:with-param name="PictureRow">
        <xsl:value-of select="$PictureRow"/>
      </xsl:with-param>
      <xsl:with-param name="PictureCell">
        <xsl:value-of select="$PictureCell"/>
      </xsl:with-param>
      <xsl:with-param name="NoteRow">
        <xsl:value-of select="$NoteRow"/>
      </xsl:with-param>
      <xsl:with-param name="NoteCell">
        <xsl:value-of select="$NoteCell"/>
      </xsl:with-param>
      <xsl:with-param name="PictureAndNoteRow">
        <xsl:value-of select="$PictureAndNoteRow"/>
      </xsl:with-param>
      <xsl:with-param name="sheet">
        <xsl:value-of select="$sheet"/>
      </xsl:with-param>
      <xsl:with-param name="NameSheet">
        <xsl:value-of select="$NameSheet"/>
      </xsl:with-param>
      <xsl:with-param name="PictureColl">
        <xsl:value-of select="$PictureColl"/>
      </xsl:with-param>
      <xsl:with-param name="NoteColl">
        <xsl:value-of select="$NoteColl"/>
      </xsl:with-param>
      <xsl:with-param name="GetMinCollWithPicture">
        <xsl:value-of select="$GetMinCollWithPicture"/>
      </xsl:with-param>
      <xsl:with-param name="sheetNr" select="$sheetNr"/>
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
      <xsl:with-param name="PictureRow">
        <xsl:value-of select="$PictureRow"/>
      </xsl:with-param>
      <xsl:with-param name="PictureCell">
        <xsl:value-of select="$PictureCell"/>
      </xsl:with-param>
      <xsl:with-param name="NoteRow">
        <xsl:value-of select="$NoteRow"/>
      </xsl:with-param>
      <xsl:with-param name="NoteCell">
        <xsl:value-of select="$NoteCell"/>
      </xsl:with-param>
      <xsl:with-param name="PictureAndNoteRow">
        <xsl:value-of select="$PictureAndNoteRow"/>
      </xsl:with-param>
      <xsl:with-param name="PictureAndNoteCell">
        <xsl:value-of select="$PictureAndNoteCell"/>
      </xsl:with-param>
      <xsl:with-param name="sheet">
        <xsl:value-of select="$sheet"/>
      </xsl:with-param>
      <xsl:with-param name="NameSheet">
        <xsl:value-of select="$NameSheet"/>
      </xsl:with-param>
      <xsl:with-param name="GetMinCollWithPicture">
        <xsl:value-of select="$GetMinCollWithPicture"/>
      </xsl:with-param>
      <xsl:with-param name="sheetNr" select="$sheetNr"/>
      <xsl:with-param name="ConditionalCell">
        <xsl:value-of select="$ConditionalCell"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellStyle">
        <xsl:value-of select="$ConditionalCellStyle"/>
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
        <xsl:value-of select="$colNum"/>
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
      <xsl:with-param name="PictureRow">
        <xsl:value-of select="$PictureRow"/>
      </xsl:with-param>
      <xsl:with-param name="PictureCell">
        <xsl:value-of select="$PictureCell"/>
      </xsl:with-param>
      <xsl:with-param name="PictureColl">
        <xsl:value-of select="$PictureColl"/>
      </xsl:with-param>
      <xsl:with-param name="NoteRow">
        <xsl:value-of select="$NoteRow"/>
      </xsl:with-param>
      <xsl:with-param name="NoteCell">
        <xsl:value-of select="$NoteCell"/>
      </xsl:with-param>
      <xsl:with-param name="NoteColl">
        <xsl:value-of select="$NoteColl"/>
      </xsl:with-param>
      <xsl:with-param name="sheet">
        <xsl:value-of select="$sheet"/>
      </xsl:with-param>
      <xsl:with-param name="NameSheet">
        <xsl:value-of select="$NameSheet"/>
      </xsl:with-param>   
      <xsl:with-param name="sheetNr" select="$sheetNr"/>
      <xsl:with-param name="ConditionalCell">
        <xsl:value-of select="$ConditionalCell"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellStyle">
        <xsl:value-of select="$ConditionalCellStyle"/>
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

  <xsl:template name="CheckSheetName">
    <xsl:param name="name"/>
    <xsl:param name="sheetNumber"/>

    <xsl:choose>
      <!-- when there are at least 2 sheets with the same name after removal of forbidden characters and cutting to 31 characters (name correction) -->
      <xsl:when test="parent::node()/e:sheet[translate(@name,'!$-()','') = $name][2]">
        <xsl:variable name="nameConflictsBefore">
          <!-- count sheets before this one whose name (after correction) collide with this sheet name (after correction) -->
          <xsl:value-of
            select="count(parent::node()/e:sheet[translate(@name,'!$-()','') = $name and position() &lt; $sheetNumber])"
          />
        </xsl:variable>
        <!-- cut name and add "(N)" at the end where N is seqential number of duplicated name -->
        <xsl:value-of select="concat(translate(@name,'!$-()',''),'_',$nameConflictsBefore + 1)"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$name"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

</xsl:stylesheet>

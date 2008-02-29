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
<!--
Modification Log
LogNo. |Date       |ModifiedBy   |BugNo.   |Modification                                                      |
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
RefNo-1 22-Jan-2008 Sandeep S     1833074   Changes for fixing Cell Content missing and 1832335 New line inserted in note content after roundtrip conversions                                              
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:pxsi="urn:cleverage:xmlns:post-processings:shared-strings"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
  xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
  xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"
  xmlns:math="http://www.w3.org/1998/Math/MathML"
  xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"
  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
  xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer"
  xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events"
  xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" office:version="1.0"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:oox="urn:oox"
  exclude-result-prefixes="e r pxsi oox">

  <xsl:import href="relationships.xsl"/>
  <xsl:import href="database-ranges.xsl"/>
  <xsl:import href="styles.xsl"/>
  <xsl:import href="table_body.xsl"/>
  <xsl:import href="number.xsl"/>
  <xsl:import href="picture.xsl"/>
  <xsl:import href="note.xsl"/>
  <xsl:import href="conditional.xsl"/>
  <xsl:import href="validation.xsl"/>
  <xsl:import href="elements.xsl"/>
  <xsl:import href="measures.xsl"/>
  <xsl:import href="ole_objects.xsl"/>
  <xsl:import href="connections.xsl"/>
  <xsl:import href="groups.xsl"/>
  <xsl:import href="scenario.xsl"/>
  <xsl:import href="change_tracking.xsl"/>
  <xsl:import href="pivot_tables.xsl"/>


  <!--xsl:key name="Sst" match="e:si" use="''"/-->
  <xsl:key name="SheetFormatPr" match="e:sheetFormatPr" use="@oox:part"/>
  <xsl:key name="Col" match="e:col" use="@oox:part"/>
  <xsl:key name="ConditionalFormatting" match="e:worksheet/e:conditionalFormatting" use="@oox:part"/>

  <!-- recursive search and replace -->
  <xsl:template name="recursive">
    <xsl:param name="oldString"/>
    <xsl:param name="newString"/>
    <xsl:param name="wholeText"/>
    <xsl:choose>
      <xsl:when test="contains($wholeText, $oldString)">
        <xsl:value-of select="concat(substring-before($wholeText, $oldString), $newString)"/>
        <xsl:call-template name="recursive">
          <xsl:with-param name="oldString" select="$oldString"/>
          <xsl:with-param name="newString" select="$newString"/>
          <xsl:with-param name="wholeText" select="substring-after($wholeText, $oldString)"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$wholeText"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="content">
    <office:document-content>
      <office:scripts/>
      <office:font-face-decls>
        <xsl:call-template name="InsertFonts"/>
      </office:font-face-decls>
      <office:automatic-styles>
        <xsl:call-template name="InsertColumnStyles"/>

        <!--xsl:comment>Row Styles</xsl:comment-->
        <xsl:call-template name="InsertRowStyles"/>

        <!--xsl:comment>Number Styles</xsl:comment-->
        <xsl:call-template name="InsertNumberStyles"/>

        <!--xsl:comment>Cell Styles</xsl:comment-->
        <xsl:call-template name="InsertCellStyles"/>

        <!--xsl:comment>Merged Cell Styles</xsl:comment-->
        <xsl:call-template name="InsertMergeCellStyles"/>

        <!--xsl:comment>Horizontal Cell Styles</xsl:comment-->
        <xsl:if test="key('Part', 'xl/styles.xml')/e:styleSheet/e:cellXfs/e:xf/e:alignment/@horizontal = 'centerContinuous'">
          <xsl:call-template name="InsertHorizontalCellStyles"/>
        </xsl:if>

        <!--xsl:comment>Table Styles</xsl:comment-->
        <xsl:call-template name="InsertStyleTableProperties"/>

        <!--xsl:comment>Text Styles</xsl:comment-->
        <xsl:call-template name="InsertTextStyles"/>

        <!--xsl:comment>Text Box Styles</xsl:comment-->
        <xsl:call-template name="InsertTextBoxTextStyles"/>
        <!-- Insert Picture properties -->

        <!--xsl:comment>Picture Styles</xsl:comment-->
        <xsl:apply-templates select="key('Part', 'xl/workbook.xml')/e:workbook/e:sheets/e:sheet[1]" mode="PictureStyle">
          <xsl:with-param name="number">1</xsl:with-param>
        </xsl:apply-templates>

        <!-- Insert Conditional Properties -->
        <!--xsl:comment>Conditional Styles</xsl:comment-->
        <xsl:apply-templates select="key('Part', 'xl/workbook.xml')/e:workbook/e:sheets/e:sheet[1]" mode="ConditionalStyle">
          <xsl:with-param name="number">1</xsl:with-param>
        </xsl:apply-templates>

        <!-- Insert Scenario properties -->
        <!--xsl:comment>Scenario Styles</xsl:comment-->
        <xsl:call-template name="InsertScenarioStyles"/>
        <!-- Insert Note Shape properties -->
        <!--xsl:comment>Note Styles</xsl:comment-->
        <xsl:for-each select="key('Part', 'xl/workbook.xml')/e:workbook/e:sheets/e:sheet">
          <xsl:call-template name="InsertNoteStyles">
            <xsl:with-param name="sheetNr">
              <xsl:value-of select="position()"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:for-each>
      </office:automatic-styles>
      <xsl:call-template name="InsertSheets"/>
    </office:document-content>
  </xsl:template>

  <xsl:template name="InsertSheets">

    <office:body>
      <office:spreadsheet>

        <!--Insert Change Tracking -->
        <xsl:call-template name="InsertChangeTracking"/>

        <xsl:variable name="rSheredStrings">
          <xsl:call-template name="rSheredStrings"/>
        </xsl:variable>

        <xsl:apply-templates select="key('Part', 'xl/workbook.xml')/e:workbook/e:sheets/e:sheet[1]"
          mode="Validation">
          <xsl:with-param name="number">1</xsl:with-param>
          <xsl:with-param name="rSheredStrings">
            <xsl:value-of select="$rSheredStrings"/>
          </xsl:with-param>
        </xsl:apply-templates>

        <!-- insert strings from sharedStrings to be moved later by post-processor-->
        <xsl:for-each select="key('Part', 'xl/sharedStrings.xml')/e:sst">
          <pxsi:sst xmlns:pxsi="urn:cleverage:xmlns:post-processings:shared-strings">
            <xsl:for-each select="e:si">
              <xsl:call-template name="e:si"/>
            </xsl:for-each>
          </pxsi:sst>
        </xsl:for-each>

        <!--xsl:for-each select="key('Part', 'xl/workbook.xml')/e:workbook/e:sheets/e:sheet">
          <xsl:call-template name="gaga" >
            <xsl:with-param name="number">1</xsl:with-param>
            <xsl:with-param name="rSheredStrings" select="$rSheredStrings"/>
          </xsl:call-template>
        </xsl:for-each-->

        <xsl:apply-templates select="key('Part', 'xl/workbook.xml')/e:workbook/e:sheets/e:sheet[1]">
          <xsl:with-param name="number">1</xsl:with-param>
          <xsl:with-param name="rSheredStrings">
            <xsl:value-of select="$rSheredStrings"/>
          </xsl:with-param>
        </xsl:apply-templates>

        <table:database-ranges>
          <xsl:apply-templates select="key('Part', 'xl/workbook.xml')/e:workbook/e:sheets/e:sheet[1]">
            <xsl:with-param name="number">1</xsl:with-param>
            <xsl:with-param name="mode" select="'database'"/>
            <xsl:with-param name="rSheredStrings">
              <xsl:value-of select="$rSheredStrings"/>
            </xsl:with-param>
          </xsl:apply-templates>
        </table:database-ranges>

        <!--xsl:variable name="pivotTables"-->
        <xsl:call-template name="InsertPilotTables"/>
        <!--/xsl:variable-->

      </office:spreadsheet>
    </office:body>
  </xsl:template>

  <xsl:template match="e:sheet">
    <xsl:param name="number"/>
    <xsl:param name="mode"/>
    <xsl:param name="rSheredStrings"/>

    <xsl:variable name="target">
      <xsl:call-template name="GetTarget">
        <xsl:with-param name="id" select="@r:id"/>
        <xsl:with-param name="document">
          <xsl:text>xl/workbook.xml</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:choose>
      <!-- insert tables -->
      <xsl:when test="$mode = '' ">

        <xsl:choose>
          <!-- when sheet is a chartsheet -->
          <xsl:when test="starts-with($target,'chartsheets/')">
            <xsl:call-template name="InsertChartsheet">
              <xsl:with-param name="number" select="$number"/>
              <xsl:with-param name="sheet" select="$target"/>
            </xsl:call-template>
          </xsl:when>
          <!-- when sheet is a worksheet -->
          <xsl:otherwise>
            <xsl:call-template name="InsertWorksheet">
              <xsl:with-param name="number" select="$number"/>
              <xsl:with-param name="rSheredStrings">
                <xsl:value-of select="$rSheredStrings"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <!-- insert database ranges -->
      <xsl:otherwise>

        <!-- if there is filter or sort -->
        <xsl:if
          test="key('Part', concat('xl/',$target))/e:worksheet/e:autoFilter or key('Part', concat('xl/',$target))/e:worksheet/e:sortState/e:sortCondition[not(@customList)]">

          <xsl:variable name="checkedName">
            <xsl:call-template name="CheckSheetName">
              <xsl:with-param name="sheetNumber">
                <xsl:value-of select="$number"/>
              </xsl:with-param>
              <xsl:with-param name="name">
                <xsl:value-of select="translate(@name,'!-$#:(),.+','')"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>

          <xsl:for-each select="key('Part', concat('xl/',$target))/e:worksheet">
            <xsl:call-template name="InsertDatabaseRange">
              <xsl:with-param name="number" select="$number"/>
              <xsl:with-param name="checkedName" select="$checkedName"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:if>

      </xsl:otherwise>
    </xsl:choose>

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
          <xsl:with-param name="mode" select="$mode"/>
          <xsl:with-param name="rSheredStrings">
            <xsl:value-of select="$rSheredStrings"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <xsl:template name="InsertWorksheet">
    <xsl:param name="number"/>
    <xsl:param name="rSheredStrings"/>

    <xsl:variable name="Id">
      <xsl:call-template name="GetTarget">
        <xsl:with-param name="id">
          <xsl:value-of select="@r:id"/>
        </xsl:with-param>
        <xsl:with-param name="document">xl/workbook.xml</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

	  <!-- Defect :1803593, file '03706191.CONFIDENTIAL.xlsx 
		   Changes by: Vijayeta
		   Desc:‘-‘ is removed from the list of symbols to get the value of ‘checkedName’.,line 342 
		   This is done because some of the sheets in the file ‘03706191.CONFIDENTIAL.xlsx’ have names such as ‘E3-SITES’, N2-L, and so on
      -->
    <xsl:variable name="checkedName">
      <xsl:call-template name="CheckSheetName">
        <xsl:with-param name="sheetNumber">
          <xsl:value-of select="$number"/>
        </xsl:with-param>
        <xsl:with-param name="name">
					<xsl:value-of select="translate(@name,'!$#():,.+','')"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="BigMergeCell">
      <xsl:for-each select="key('Part', concat('xl/',$Id))/e:worksheet/e:mergeCells">
        <xsl:apply-templates select="e:mergeCell[1]" mode="BigMergeColl"/>
      </xsl:for-each>
    </xsl:variable>

    <xsl:variable name="BigMergeRow">
      <xsl:for-each select="key('Part', concat('xl/',$Id))/e:worksheet/e:mergeCells">
        <xsl:apply-templates select="e:mergeCell[1]" mode="BigMergeRow"/>
      </xsl:for-each>
    </xsl:variable>

    <!-- variable with merge cell -->
    <xsl:variable name="MergeCell">
      <xsl:for-each select="key('Part', concat('xl/',$Id))/e:worksheet/e:mergeCells">
        <xsl:apply-templates select="e:mergeCell[1]" mode="merge"/>
      </xsl:for-each>
    </xsl:variable>


    <!-- Check If Picture are in this sheet  -->
    <xsl:variable name="PictureCell">
      <xsl:for-each select="key('Part', concat('xl/',$Id))">
        <xsl:call-template name="PictureCell">
          <xsl:with-param name="sheet">
            <xsl:value-of select="substring-after($Id, '/')"/>
          </xsl:with-param>
          <xsl:with-param name="mode">
            <xsl:value-of select="substring-before($Id, '/')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:for-each>
    </xsl:variable>

    <!-- Check if notes are in this sheet -->
    <xsl:variable name="NoteCell">
      <xsl:for-each select="key('Part', concat('xl/',$Id))">
        <xsl:call-template name="NoteCell">
          <xsl:with-param name="sheetNr" select="$number"/>
        </xsl:call-template>
      </xsl:for-each>
    </xsl:variable>

    <!-- Check If Conditionals are in this sheet -->

    <xsl:variable name="ConditionalCell">
      <!--xsl:for-each select="key('Part', concat('xl/',$Id))">
        <xsl:call-template name="ConditionalCell"/>
      </xsl:for-each-->
    </xsl:variable>

    <xsl:variable name="ConditionalRow">
      <!--xsl:call-template name="ConditionalRow">
        <xsl:with-param name="ConditionalCell">
          <xsl:value-of select="$ConditionalCell"/>
        </xsl:with-param>
      </xsl:call-template-->
    </xsl:variable>

    <xsl:variable name="ConditionalCellStyle">
      <!--xsl:for-each select="key('Part', concat('xl/',$Id))">
        <xsl:call-template name="ConditionalCell">
          <xsl:with-param name="document">
            <xsl:text>style</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:for-each-->
    </xsl:variable>
    
    <xsl:variable name="ConditionalCellCol">
      <xsl:for-each select="key('Part', concat('xl/',$Id))">
          <xsl:call-template name="ConditionalCellCol">        
          </xsl:call-template>
       </xsl:for-each>
    </xsl:variable>
    
    <xsl:variable name="ConditionalCellAll">
      <xsl:for-each select="key('Part', concat('xl/',$Id))">
        <xsl:call-template name="ConditionalCellAll">          
        </xsl:call-template>
      </xsl:for-each>
    </xsl:variable>
    
    <xsl:variable name="ConditionalCellSingle">
      <xsl:call-template name="ConditionalCellSingle">
        <xsl:with-param name="sqref">
          <xsl:value-of select="$ConditionalCellAll"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    
    
    <xsl:variable name="ConditionalCellMultiple">
      <xsl:call-template name="ConditionalCellMultiple">
        <xsl:with-param name="sqref">
          <xsl:value-of select="$ConditionalCellAll"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    

    <!-- Check If Data Validation are in this sheet -->
    <xsl:variable name="ValidationCell">
      <xsl:for-each select="key('Part', concat('xl/',$Id))">
        <xsl:call-template name="ValidationCell"/>
      </xsl:for-each>
    </xsl:variable>

    <xsl:variable name="ValidationRow">
      <xsl:call-template name="ValidationRow">
        <xsl:with-param name="ValidationCell">
          <xsl:value-of select="$ValidationCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="ValidationCellStyle">
      <xsl:for-each select="key('Part', concat('xl/',$Id))">
        <xsl:call-template name="ValidationCell">
          <xsl:with-param name="document">
            <xsl:text>style</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:for-each>
    </xsl:variable>

    <!-- Check if Scenario are in this sheet -->
    <xsl:variable name="ScenarioCell">
      <xsl:for-each select="key('Part', concat('xl/',$Id))">
        <xsl:call-template name="ScenarioCell"/>
      </xsl:for-each>
    </xsl:variable>

    <xsl:variable name="PictureRow">
      <xsl:for-each select="key('Part', concat('xl/',$Id))">
        <xsl:call-template name="PictureRow">
          <xsl:with-param name="PictureCell">
            <xsl:value-of select="$PictureCell"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:for-each>
    </xsl:variable>

    <xsl:variable name="NoteRow">
      <xsl:for-each select="key('Part', concat('xl/',$Id))">
        <xsl:call-template name="NoteRow">
          <xsl:with-param name="NoteCell">
            <xsl:value-of select="$NoteCell"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:for-each>
    </xsl:variable>

    <xsl:variable name="removeFilter">
      <xsl:if test="key('Part', concat('xl/',$Id))/e:worksheet/e:autoFilter">

        <xsl:for-each select="key('Part', concat('xl/',$Id))/e:worksheet/e:autoFilter">
          <xsl:variable name="filtersNum">
            <xsl:value-of select="count(e:filterColumn/e:filters)"/>
          </xsl:variable>
          <xsl:variable name="customFiltersNum">
            <xsl:value-of select="count(e:filterColumn/e:customFilters)"/>
          </xsl:variable>
          <xsl:variable name="topFiltersNum">
            <xsl:value-of select="count(e:filterColumn/e:top10)"/>
          </xsl:variable>

          <xsl:if
            test="e:filterColumn/e:filters/e:filter[position() = 2] and $filtersNum + $customFiltersNum + $topFiltersNum &gt; 1">
            <xsl:call-template name="GetRowNum">
              <xsl:with-param name="cell" select="substring-before(@ref,':')"/>
            </xsl:call-template>
            <xsl:text>:</xsl:text>
            <xsl:call-template name="GetRowNum">
              <xsl:with-param name="cell" select="substring-after(@ref,':')"/>
            </xsl:call-template>
          </xsl:if>

        </xsl:for-each>
      </xsl:if>
    </xsl:variable>

    <xsl:variable name="ConnectionCell">
      <xsl:call-template name="ConnectionsCell">
        <xsl:with-param name="number">
          <xsl:value-of select="$number - 1"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!-- variable with values of all manual row breaks-->
    <xsl:variable name="AllRowBreakes">
      <xsl:for-each select="key('Part', concat('xl/',$Id))/e:worksheet/e:rowBreaks/e:brk">
        <xsl:value-of select="concat(@id + 1,';')"/>
      </xsl:for-each>
    </xsl:variable>

    <table:table>

      <xsl:variable name="GroupCell">
        <xsl:apply-templates select="key('Part', concat('xl/',$Id))/e:worksheet/e:cols/e:col[1]"
          mode="groupTag"/>
      </xsl:variable>


      <!-- Insert Table (Sheet) Name -->
      <xsl:attribute name="table:name">
        <!--        <xsl:value-of select="@name"/>-->
        <xsl:value-of select="$checkedName"/>
      </xsl:attribute>

      <!-- Insert Table Style Name (style:table-properties) -->

      <xsl:attribute name="table:style-name">
        <xsl:value-of select="generate-id()"/>
      </xsl:attribute>

      <!-- insert Print Range -->
      <xsl:variable name="apostrof">
        <xsl:text>&apos;</xsl:text>
      </xsl:variable>
      <xsl:for-each select="key('Part', 'xl/workbook.xml')/e:workbook/e:definedNames/e:definedName">
        <!-- for the current sheet -->
        <xsl:if test="not(contains(self::node(),'#REF'))">
          <!-- if print range with apostrophes -->
          <!--
		         Defect :1803593, file '03706191.CONFIDENTIAL.xlsx 
		         Changes by: Vijayeta
		         Desc:variable 'sheetNamePrntRnge', to prevent, a situation where a sheet name is E3 and there might be another sheet as E3-GS   
        -->
          <xsl:variable name ="sheetNamePrntRnge">
            <xsl:value-of select ="substring-before(./self::node(),'!')"/>
          </xsl:variable>
          <!--<xsl:if test="contains(./self::node(), concat($apostrof, $checkedName)) and (@name = '_xlnm.Print_Area' or @name = '_xlnm.Print_Titles')">-->
          <xsl:if test="(($sheetNamePrntRnge = concat($apostrof, $checkedName,$apostrof)) and (@name = '_xlnm.Print_Area'))">
            <!-- one print range with apostrophes -->
            <xsl:if test="not(contains(./self::node(),concat(',', $apostrof, $checkedName)))">
              <xsl:variable name ="prntRnge">
                <xsl:call-template name="recursive">
                  <xsl:with-param name="oldString" select="':'"/>
                  <xsl:with-param name="newString" select="concat(':', $apostrof, $checkedName, $apostrof, '.')"/>
                  <xsl:with-param name="wholeText" select="translate(./self::node(), '!', '.')"/>
                </xsl:call-template>
              </xsl:variable>
              <xsl:variable name ="PartOne">
                <xsl:value-of select ="substring-before($prntRnge,':')"/>
              </xsl:variable>
              <xsl:variable name ="PartTwo">
                <xsl:value-of select ="substring-after($prntRnge,':')"/>
              </xsl:variable>
              <xsl:variable name ="beginSheet">
                <xsl:value-of select ="substring-before($PartOne,'.')"/>
              </xsl:variable>
              <xsl:variable name ="endSheet">
                <xsl:value-of select ="substring-before($PartTwo,'.')"/>
              </xsl:variable>
              <xsl:variable name ="beginCol">
                <xsl:choose >
                  <xsl:when test ="contains(substring-after($PartOne,'$'),'$')">
                    <xsl:value-of select ="substring-before(substring-after($PartOne,'$'),'$')"/>
                  </xsl:when>
                  <xsl:otherwise >
                    <xsl:value-of select ="substring-after($PartOne,'$')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>
              <xsl:variable name ="endCol">
                <xsl:choose >
                  <xsl:when test ="contains(substring-after($PartTwo,'$'),'$')">
                    <xsl:value-of select ="substring-before(substring-after($PartTwo,'$'),'$')"/>
                  </xsl:when>
                  <xsl:otherwise >
                    <xsl:value-of select ="substring-after($PartTwo,'$')"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>
              <xsl:variable name ="beginRow">
                <xsl:choose >
                  <xsl:when test ="contains(substring-after($PartOne,'$'),'$')">
                    <xsl:value-of select ="substring-after(substring-after($PartOne,'$'),'$')"/>
                  </xsl:when>
                  <xsl:otherwise >
                    <xsl:value-of select ="1"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>
              <xsl:variable name ="endRow">
                <xsl:choose >
                  <xsl:when test ="contains(substring-after($PartTwo,'$'),'$')">
                    <xsl:value-of select ="substring-after(substring-after($PartTwo,'$'),'$')"/>
                  </xsl:when>
                  <xsl:otherwise >
                    <xsl:value-of select ="65536"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>
              <xsl:variable name ="newPrntRnge">
                <xsl:if test ="$endSheet!=''">
                  <xsl:value-of select ="concat($beginSheet,'.',$beginCol,$beginRow,':',$endSheet,'.',$endCol,$endRow)"/>
                </xsl:if>
                <xsl:if test ="$endSheet=''">
                  <xsl:value-of select ="concat($beginSheet,'.',$beginCol,$beginRow,':',$beginSheet,'.',$endCol,$endRow)"/>
                </xsl:if>
              </xsl:variable>
              <xsl:attribute name="table:print-ranges">
                <xsl:value-of select ="$newPrntRnge"/>
              </xsl:attribute>
            </xsl:if>
            <!-- multiple print ranges with apostrophes -->
            <!-- Bug: 1803593, for the file U S Extreme Temperature, on round trip
				         Fixed By: Vijayeta
				         Desc: Print range in the input xlsx file is not in a proper format, hence in the converted ods file print range does not appear and hence,
				            the xlsx file after round trip crashes. This part of code fixes this problem
					   -->
            <xsl:if test="contains(./self::node(),concat(',', $apostrof, $checkedName))">
              <xsl:variable name ="GetPrintRange">
                <xsl:value-of select="./self::node()"/>
              </xsl:variable>
              <xsl:variable name ="newRange">
                <xsl:value-of select ="translate(translate(translate($GetPrintRange,'!','.'),'$',''),',',' ')"/>
              </xsl:variable>
              <xsl:attribute name="table:print-ranges">
                <xsl:value-of select ="$newRange"/>
              </xsl:attribute>
            </xsl:if>
            <!-- End of fix for Bug 1803593, for the file U S Extreme Temperature, on round trip-->
          </xsl:if>
        </xsl:if>
      </xsl:for-each>
      <xsl:for-each select="key('Part', 'xl/workbook.xml')/e:workbook/e:definedNames/e:definedName">
        <!-- for the current sheet -->
        <!-- if the print range is without apostrophes -->
        <xsl:if test="not(contains(self::node(),'#REF'))">
        <!-- 
		      Defect :1803593, file '03706191.CONFIDENTIAL.xlsx 
		      Changes by: Vijayeta
		      Desc:‘-‘ is removed from the list of symbols,line 653 
		            This is done because some of the sheets in the file ‘03706191.CONFIDENTIAL.xlsx’ have names such as ‘E3-SITES’, N2-L, and so on
        -->
          <xsl:if test="string($checkedName) = translate(substring-before(self::node(), '!'), '!$#:(),.+','') and (@name = '_xlnm.Print_Area' or @name = '_xlnm.Print_Titles')">
            <!-- one print range without apostrophes -->
            <xsl:if test="not(contains(./self::node(), concat(',', $checkedName)))">
              <xsl:variable name="temporary">
                <xsl:call-template name="recursive">
                  <xsl:with-param name="wholeText" select="translate(self::node(), '!', '.')"/>
                  <xsl:with-param name="newString" select="concat(':', $checkedName, '.')"/>
                  <xsl:with-param name="oldString" select="':'"/>
                </xsl:call-template>
              </xsl:variable>
              <xsl:attribute name="table:print-ranges">
                <xsl:value-of select ="translate(translate($temporary, '!', '.'),'$','')"/>
              </xsl:attribute>
            </xsl:if>
            <!-- multiple print ranges without apostrophes -->
            <xsl:if test="contains(./self::node(), concat(',', $checkedName))">
              <xsl:attribute name="table:print-ranges">
                <xsl:call-template name="recursive">
                  <xsl:with-param name="newString" select="concat(':', $checkedName, '.')"/>
								<xsl:with-param name="wholeText" select="translate(translate(translate(./self::node(), '!', '.'), ',', ' '),'$','')"/>
                  <xsl:with-param name="oldString" select="':'"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
        </xsl:if>
      </xsl:for-each>
      <xsl:variable name="GroupRowStart">
        <xsl:call-template name="GroupRow">
          <xsl:with-param name="Id">
            <xsl:value-of select="$Id"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>

      <xsl:variable name="GroupRowEnd">
        <xsl:call-template name="GroupRowEnd">
          <xsl:with-param name="Id">
            <xsl:value-of select="$Id"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>

      <xsl:apply-templates
        select="key('Part', 'xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[1]"
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
        <xsl:with-param name="MergeCell">
          <xsl:value-of select="$MergeCell"/>
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
        <xsl:with-param name="ConditionalCellStyle">
          <xsl:value-of select="$ConditionalCellStyle"/>
        </xsl:with-param>
        <xsl:with-param name="ConditionalRow">
          <xsl:value-of select="$ConditionalRow"/>
        </xsl:with-param>
        <xsl:with-param name="ConditionalCellCol">
          <xsl:value-of select="$ConditionalCellCol"/>
        </xsl:with-param>
        <xsl:with-param name="ConditionalCellAll">
          <xsl:value-of select="$ConditionalCellAll"/>
        </xsl:with-param>
        <xsl:with-param name="ConditionalCellSingle">
          <xsl:value-of select="$ConditionalCellSingle"/>
        </xsl:with-param>
        <xsl:with-param name="ConditionalCellMultiple">
          <xsl:value-of select="$ConditionalCellMultiple"/>
        </xsl:with-param>
        <xsl:with-param name="PictureCell">
          <xsl:value-of select="$PictureCell"/>
        </xsl:with-param>
        <xsl:with-param name="PictureRow">
          <xsl:value-of select="$PictureRow"/>
        </xsl:with-param>
        <xsl:with-param name="removeFilter" select="$removeFilter"/>
        <xsl:with-param name="ValidationCell">
          <xsl:value-of select="$ValidationCell"/>
        </xsl:with-param>
        <xsl:with-param name="ValidationRow">
          <xsl:value-of select="$ValidationRow"/>
        </xsl:with-param>
        <xsl:with-param name="ValidationCellStyle">
          <xsl:value-of select="$ValidationCellStyle"/>
        </xsl:with-param>
        <xsl:with-param name="ConnectionsCell">
          <xsl:value-of select="$ConnectionCell"/>
        </xsl:with-param>
        <xsl:with-param name="GroupCell">
          <xsl:value-of select="$GroupCell"/>
        </xsl:with-param>
        <xsl:with-param name="GroupRowStart">
          <xsl:value-of select="$GroupRowStart"/>
        </xsl:with-param>
        <xsl:with-param name="GroupRowEnd">
          <xsl:value-of select="$GroupRowEnd"/>
        </xsl:with-param>
        <xsl:with-param name="AllRowBreakes">
          <xsl:value-of select="$AllRowBreakes"/>
        </xsl:with-param>
        <xsl:with-param name="rSheredStrings">
          <xsl:value-of select="$rSheredStrings"/>
        </xsl:with-param>
      </xsl:call-template>

    </table:table>

    <xsl:for-each select="key('Part', concat('xl/',$Id))/e:worksheet/e:scenarios">
      <xsl:call-template name="Scenarios"/>
    </xsl:for-each>

    <xsl:if test="key('Part', concat('xl/',$Id))/e:worksheet/e:dataConsolidate">
      <xsl:message terminate="no">translation.oox2odf.DataConsolidation</xsl:message>
    </xsl:if>

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

        <xsl:call-template name="InsertRanges">
          <xsl:with-param name="ranges" select="text()"/>
          <xsl:with-param name="mode" select="substring-after(text(),',')"/>
          <xsl:with-param name="checkedName" select="$checkedName"/>
        </xsl:call-template>
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
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertRanges">
    <xsl:param name="ranges"/>
    <xsl:param name="mode"/>
    <xsl:param name="checkedName"/>
    <!-- if print ranges attribute does not contain '#REF' -->
    <xsl:for-each select="key('Part', 'xl/workbook.xml')">
      <xsl:if
        test="not(contains(//workbook/definedNames/definedName[attribute::name = '_xlnm.Print_Area' ],'#REF'))">
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
                    select="substring-before(substring-after($ranges,concat($sheetName,'!')),':' )"
                  />
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of
                    select="substring-before(substring-after($ranges,concat($sheetName,'!')),',' )"
                  />
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
                    select="substring-before(substring-after($ranges,concat($sheetName,'!')),',' )"
                  />
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
            <!--  
           <xsl:value-of
              select="concat($apos,$checkedName,$apos,'.',$start,':',$apos,$checkedName,$apos,'.',$end)"/>
            <xsl:text> </xsl:text>
-->
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
                    select="substring-before(substring-after($ranges,concat($sheetName,'!')),':' )"
                  />
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

            <!--          <xsl:value-of
              select="concat($apos,$checkedName,$apos,'.',$start,':',$apos,$checkedName,$apos,'.',$end)"/>
-->
          </xsl:otherwise>
        </xsl:choose>

      </xsl:if>
    </xsl:for-each>
    <!-- if print ranges attribute contains '#REF' then there should be 'table:print' attribute put instead with 'false' value -->
    <xsl:for-each select="key('Part', 'xl/workbook.xml')">
      <xsl:if
        test="contains(//workbook/definedNames/definedName[attribute::name = '_xlnm.Print_Area' ],'#REF')">
        <xsl:attribute name="table:print">
          <xsl:value-of select="'false'"/>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>


  <!-- insert string -->
  <xsl:template name="e:si">
    <xsl:if test="not(e:r)">
      <pxsi:si pxsi:number="{position() - 1}">
        <xsl:apply-templates/>
      </pxsi:si>
    </xsl:if>

  </xsl:template>

  <xsl:template name="InsertSheetContent">
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="BigMergeCell"/>
    <xsl:param name="BigMergeRow"/>
    <xsl:param name="MergeCell"/>
    <xsl:param name="sheetNr"/>
    <xsl:param name="PictureCell"/>
    <xsl:param name="PictureRow"/>
    <xsl:param name="NoteCell"/>
    <xsl:param name="NoteRow"/>
    <xsl:param name="ConditionalCell"/>
    <xsl:param name="ConditionalCellStyle"/>
    <xsl:param name="ConditionalRow"/>
    <xsl:param name="ConditionalCellCol"/>
    <xsl:param name="ConditionalCellAll"/>
    <xsl:param name="ConditionalCellSingle"/>
    <xsl:param name="ConditionalCellMultiple"/>
    <xsl:param name="removeFilter"/>
    <xsl:param name="ValidationCell"/>
    <xsl:param name="ValidationRow"/>
    <xsl:param name="ValidationCellStyle"/>
    <xsl:param name="ConnectionsCell"/>
    <xsl:param name="GroupCell"/>
    <xsl:param name="GroupRowStart"/>
    <xsl:param name="GroupRowEnd"/>
    <xsl:param name="AllRowBreakes"/>
    <xsl:param name="rSheredStrings"/>

    <xsl:for-each select="key('Part', concat('xl/',$sheet))/e:worksheet/e:oleObjects">
      <xsl:call-template name="InsertOLEObjects"/>
    </xsl:for-each>


    <xsl:variable name="ManualColBreaks">
      <xsl:for-each select="key('Part', concat('xl/',$sheet))/e:worksheet/e:colBreaks/e:brk">
        <xsl:value-of select="concat(@id,';')"/>
      </xsl:for-each>
    </xsl:variable>


    <xsl:call-template name="InsertColumns">
      <xsl:with-param name="sheet" select="$sheet"/>
      <xsl:with-param name="GroupCell">
        <xsl:value-of select="$GroupCell"/>
      </xsl:with-param>
      <xsl:with-param name="ManualColBreaks">
        <xsl:value-of select="$ManualColBreaks"/>
      </xsl:with-param>
    </xsl:call-template>

    <xsl:variable name="sheetName">
      <xsl:value-of select="@name"/>
    </xsl:variable>

    <xsl:for-each select="key('Part', concat('xl/',$sheet))">

      <xsl:variable name="apos">
        <xsl:text>&apos;</xsl:text>
      </xsl:variable>

      <xsl:variable name="headerRowsStart">
        <xsl:choose>
          <!-- if sheet name in range definition is in apostrophes -->
          <xsl:when
            test="key('Part', 'xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($apos,$sheetName,$apos))]">
            <xsl:for-each
              select="key('Part', 'xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($apos,$sheetName,$apos))]">
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
            test="key('Part', 'xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($sheetName,'!'))]">
            <xsl:for-each
              select="key('Part', 'xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($sheetName,'!'))]">
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
            test="key('Part', 'xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($apos,$sheetName,$apos))]">
            <xsl:for-each
              select="key('Part', 'xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($apos,$sheetName,$apos))]">
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
            test="key('Part', 'xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($sheetName,'!'))]">
            <xsl:for-each
              select="key('Part', 'xl/workbook.xml')/e:workbook/e:definedNames/e:definedName[@name= '_xlnm.Print_Titles' and starts-with(text(),concat($sheetName,'!'))]">
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
            <xsl:with-param name="MergeCell">
              <xsl:value-of select="$MergeCell"/>
            </xsl:with-param>
            <xsl:with-param name="headerRowsStart" select="$headerRowsStart"/>
            <xsl:with-param name="headerRowsEnd" select="$headerRowsEnd"/>
            <xsl:with-param name="sheetNr" select="$sheetNr"/>
            <xsl:with-param name="removeFilter" select="$removeFilter"/>
            <xsl:with-param name="rSheredStrings">
              <xsl:value-of select="$rSheredStrings"/>
            </xsl:with-param>
          </xsl:apply-templates>

          <!-- insert empty rows before header -->
          <xsl:if
            test="$headerRowsStart &gt; 1 and not(e:worksheet/e:sheetData/e:row[@r = $headerRowsStart - 1 and @r &lt; 65537])">
            <xsl:choose>
              <!-- when there aren't any rows before at all -->
              <xsl:when
                test="not(e:worksheet/e:sheetData/e:row[@r &lt; $headerRowsStart and @r &lt; 65537])">
                <table:table-row table:style-name="{concat('ro', key('Part', concat('xl/',$sheet))/e:worksheet/@oox:part)}"
                  table:number-rows-repeated="{$headerRowsStart - 1}">
                  <table:table-cell table:number-columns-repeated="256"/>
                </table:table-row>
              </xsl:when>
              <!-- if there was a row before header -->
              <xsl:otherwise>
                <xsl:for-each
                  select="e:worksheet/e:sheetData/e:row[@r &lt; $headerRowsStart and @r &lt; 65537][last()]">
                  <table:table-row table:style-name="{concat('ro', key('Part', concat('xl/',$sheet))/e:worksheet/@oox:part)}"
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
              <xsl:with-param name="MergeCell">
                <xsl:value-of select="$MergeCell"/>
              </xsl:with-param>
              <xsl:with-param name="headerRowsStart" select="$headerRowsStart"/>
              <xsl:with-param name="headerRowsEnd" select="$headerRowsEnd"/>
              <xsl:with-param name="sheetNr" select="$sheetNr"/>
              <xsl:with-param name="removeFilter" select="$removeFilter"/>
              <xsl:with-param name="rSheredStrings">
                <xsl:value-of select="$rSheredStrings"/>
              </xsl:with-param>
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
                  <xsl:with-param name="MergeCell">
                    <xsl:value-of select="$MergeCell"/>
                  </xsl:with-param>
                  <xsl:with-param name="RowNumber">
                    <xsl:value-of select="e:worksheet/e:sheetData/e:row[position() = last()]/@r"/>
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
                  <xsl:with-param name="NameSheet">
                    <xsl:value-of select="$sheetName"/>
                  </xsl:with-param>
                  <xsl:with-param name="ConditionalCell">
                    <xsl:value-of select="$ConditionalCell"/>
                  </xsl:with-param>
                  <xsl:with-param name="ConditionalCellStyle">
                    <xsl:value-of select="$ConditionalCellStyle"/>
                  </xsl:with-param>
                  <xsl:with-param name="ConditionalRow">
                    <xsl:value-of select="$ConditionalRow"/>
                  </xsl:with-param>
                  <xsl:with-param name="ConditionalCellCol">
                    <xsl:value-of select="$ConditionalCellCol"/>
                  </xsl:with-param>
                  <xsl:with-param name="ConditionalCellAll">
                    <xsl:value-of select="$ConditionalCellAll"/>
                  </xsl:with-param>
                  <xsl:with-param name="ConditionalCellSingle">
                    <xsl:value-of select="$ConditionalCellSingle"/>
                  </xsl:with-param>
                  <xsl:with-param name="ConditionalCellMultiple">
                    <xsl:value-of select="$ConditionalCellMultiple"/>
                  </xsl:with-param>
                  <xsl:with-param name="sheetNr" select="$sheetNr"/>
                  <xsl:with-param name="ValidationCell">
                    <xsl:value-of select="$ValidationCell"/>
                  </xsl:with-param>
                  <xsl:with-param name="ValidationRow">
                    <xsl:value-of select="$ValidationRow"/>
                  </xsl:with-param>
                  <xsl:with-param name="ValidationCellStyle">
                    <xsl:value-of select="$ValidationCellStyle"/>
                  </xsl:with-param>
                  <xsl:with-param name="AllRowBreakes">
                    <xsl:value-of select="$AllRowBreakes"/>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:when>

              <xsl:when
                test="not(e:worksheet/e:sheetData/e:row[@r &gt;= $headerRowsStart and @r &lt;= $headerRowsEnd and @r &lt; 65537])">
                <table:table-row table:style-name="{concat('ro', key('Part', concat('xl/',$sheet))/e:worksheet/@oox:part)}"
                  table:number-rows-repeated="{$headerRowsEnd - $headerRowsStart + 1}">
                  <table:table-cell table:number-columns-repeated="256"/>
                </table:table-row>
              </xsl:when>
            </xsl:choose>

            <!-- if there are empty rows at the end of the header -->
            <xsl:for-each
              select="e:worksheet/e:sheetData/e:row[@r &gt;= $headerRowsStart and @r &lt;= $headerRowsEnd and @r &lt; 65537][last()]">
              <xsl:if test="@r &lt; $headerRowsEnd">
                <table:table-row table:style-name="{concat('ro', key('Part', concat('xl/',$sheet))/e:worksheet/@oox:part)}"
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
            <xsl:with-param name="MergeCell">
              <xsl:value-of select="$MergeCell"/>
            </xsl:with-param>
            <!-- Code Added By Sateesh Reddy Date:01-Feb-2008 -->
            <xsl:with-param name="sheetNr" select="$sheetNr"/>
            <!-- End -->
            <xsl:with-param name="headerRowsStart" select="$headerRowsStart"/>
            <xsl:with-param name="headerRowsEnd" select="$headerRowsEnd"/>
            <xsl:with-param name="removeFilter" select="$removeFilter"/>
            <xsl:with-param name="rSheredStrings">
              <xsl:value-of select="$rSheredStrings"/>
            </xsl:with-param>
          </xsl:apply-templates>
        </xsl:when>

        <!-- if there aren't any header rows -->
        <xsl:otherwise>

          <xsl:choose>
            <xsl:when test="e:worksheet/e:sheetData/e:row">
              <xsl:apply-templates select="e:worksheet/e:sheetData/e:row[@r &lt; 65537]">
                <xsl:with-param name="BigMergeCell">
                  <xsl:value-of select="$BigMergeCell"/>
                </xsl:with-param>
                <xsl:with-param name="BigMergeRow">
                  <xsl:value-of select="$BigMergeRow"/>
                </xsl:with-param>
                <xsl:with-param name="MergeCell">
                  <xsl:value-of select="$MergeCell"/>
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
                <xsl:with-param name="ConditionalRow">
                  <xsl:value-of select="$ConditionalRow"/>
                </xsl:with-param>
                <xsl:with-param name="ConditionalCellStyle">
                  <xsl:value-of select="$ConditionalCellStyle"/>
                </xsl:with-param>
                <xsl:with-param name="ConditionalCellCol">
                  <xsl:value-of select="$ConditionalCellCol"/>
                </xsl:with-param>
                <xsl:with-param name="ConditionalCellAll">
                  <xsl:value-of select="$ConditionalCellAll"/>
                </xsl:with-param>
                <xsl:with-param name="ConditionalCellSingle">
                  <xsl:value-of select="$ConditionalCellSingle"/>
                </xsl:with-param>
                <xsl:with-param name="ConditionalCellMultiple">
                  <xsl:value-of select="$ConditionalCellMultiple"/>
                </xsl:with-param>
                <xsl:with-param name="removeFilter" select="$removeFilter"/>
                <xsl:with-param name="ValidationCell">
                  <xsl:value-of select="$ValidationCell"/>
                </xsl:with-param>
                <xsl:with-param name="ValidationRow">
                  <xsl:value-of select="$ValidationRow"/>
                </xsl:with-param>
                <xsl:with-param name="ValidationCellStyle">
                  <xsl:value-of select="$ValidationCellStyle"/>
                </xsl:with-param>
                <xsl:with-param name="ConnectionsCell">
                  <xsl:value-of select="$ConnectionsCell"/>
                </xsl:with-param>
                <xsl:with-param name="GroupRowStart">
                  <xsl:value-of select="$GroupRowStart"/>
                </xsl:with-param>
                <xsl:with-param name="GroupRowEnd">
                  <xsl:value-of select="$GroupRowEnd"/>
                </xsl:with-param>
                <xsl:with-param name="AllRowBreakes">
                  <xsl:value-of select="$AllRowBreakes"/>
                </xsl:with-param>
                <xsl:with-param name="rSheredStrings">
                  <xsl:value-of select="$rSheredStrings"/>
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
                <xsl:with-param name="MergeCell">
                  <xsl:value-of select="$MergeCell"/>
                </xsl:with-param>
                <xsl:with-param name="RowNumber">
                  <xsl:value-of select="e:worksheet/e:sheetData/e:row[position() = last()]/@r"/>
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
                <xsl:with-param name="NameSheet">
                  <xsl:value-of select="$sheetName"/>
                </xsl:with-param>
                <xsl:with-param name="ConditionalCell">
                  <xsl:value-of select="$ConditionalCell"/>
                </xsl:with-param>
                <xsl:with-param name="ConditionalCellStyle">
                  <xsl:value-of select="$ConditionalCellStyle"/>
                </xsl:with-param>
                <xsl:with-param name="ConditionalRow">
                  <xsl:value-of select="$ConditionalRow"/>
                </xsl:with-param>
                <xsl:with-param name="ConditionalCellCol">
                  <xsl:value-of select="$ConditionalCellCol"/>
                </xsl:with-param>
                <xsl:with-param name="ConditionalCellAll">
                  <xsl:value-of select="$ConditionalCellAll"/>
                </xsl:with-param>
                <xsl:with-param name="ConditionalCellSingle">
                  <xsl:value-of select="$ConditionalCellSingle"/>
                </xsl:with-param>
                <xsl:with-param name="ConditionalCellMultiple">
                  <xsl:value-of select="$ConditionalCellMultiple"/>
                </xsl:with-param>
                <xsl:with-param name="sheetNr" select="$sheetNr"/>
                <xsl:with-param name="ValidationCell">
                  <xsl:value-of select="$ValidationCell"/>
                </xsl:with-param>
                <xsl:with-param name="ValidationRow">
                  <xsl:value-of select="$ValidationRow"/>
                </xsl:with-param>
                <xsl:with-param name="ValidationCellStyle">
                  <xsl:value-of select="$ValidationCellStyle"/>
                </xsl:with-param>
                <xsl:with-param name="AllRowBreakes">
                  <xsl:value-of select="$AllRowBreakes"/>
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
    <xsl:param name="MergeCell"/>
    <xsl:param name="PictureCell"/>
    <xsl:param name="PictureRow"/>
    <xsl:param name="NoteCell"/>
    <xsl:param name="NoteRow"/>
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="sheetNr"/>
    <xsl:param name="ConditionalCell"/>
    <xsl:param name="ConditionalCellStyle"/>
    <xsl:param name="ConditionalRow"/>
    <xsl:param name="ConditionalCellCol"/>
    <xsl:param name="ConditionalCellAll"/>
    <xsl:param name="ConditionalCellSingle"/>
    <xsl:param name="ConditionalCellMultiple"/>
    <xsl:param name="removeFilter"/>
    <xsl:param name="ValidationCell"/>
    <xsl:param name="ValidationRow"/>
    <xsl:param name="ValidationCellStyle"/>
    <xsl:param name="ConnectionsCell"/>
    <xsl:param name="outlineLevel"/>
    <xsl:param name="GroupRowStart"/>
    <xsl:param name="GroupRowEnd"/>
    <xsl:param name="AllRowBreakes"/>
    <xsl:param name="rSheredStrings"/>

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
        <xsl:otherwise>0</xsl:otherwise>
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

    <xsl:variable name="GetMinRowWithElement">
      <xsl:call-template name="GetMinRowWithPicture">
        <xsl:with-param name="PictureRow">
          <xsl:value-of
            select="concat($PictureRow, $NoteRow, $ConditionalRow, $ValidationRow, $AllRowBreakes)"
          />
        </xsl:with-param>
        <xsl:with-param name="AfterRow">
          <xsl:choose>
            <xsl:when test="preceding-sibling::e:row[1]/@r = ''">
              <xsl:text>0</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="preceding-sibling::e:row[1]/@r"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!-- if there were empty rows before this one then insert empty rows -->
    <xsl:choose>
      <!-- when this row is the first non-empty one but not row 1 and there aren't Big Merge Coll and Pictures-->
      <xsl:when
        test="position()=1 and @r>1 and $BigMergeCell = '' and ($GetMinRowWithElement = '' or ($GetMinRowWithElement) &gt;= @r)">

        <table:table-row table:style-name="{concat('ro', key('Part', concat('xl/',$sheet))/e:worksheet/@oox:part)}"
          table:number-rows-repeated="{@r - 1}">
          <table:table-cell table:number-columns-repeated="256"/>
        </table:table-row>

      </xsl:when>

      <!-- when this row is the first non-empty one but not row 1 and there aren't Big Merge Coll, and there are Pictures before this row-->
      <xsl:when
        test="position()=1 and @r>1 and $BigMergeCell = '' and $GetMinRowWithElement &lt; @r">

        <xsl:call-template name="InsertElementsBetwenTwoRows">
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="NameSheet">
            <xsl:value-of select="$NameSheet"/>
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
          <xsl:with-param name="ConditionalCell">
            <xsl:value-of select="$ConditionalCell"/>
          </xsl:with-param>
          <xsl:with-param name="ConditionalCellStyle">
            <xsl:value-of select="$ConditionalCellStyle"/>
          </xsl:with-param>
          <xsl:with-param name="ConditionalRow">
            <xsl:value-of select="$ConditionalRow"/>
          </xsl:with-param>
          <xsl:with-param name="ConditionalCellCol">
            <xsl:value-of select="$ConditionalCellCol"/>
          </xsl:with-param>
          <xsl:with-param name="ConditionalCellAll">
            <xsl:value-of select="$ConditionalCellAll"/>
          </xsl:with-param>
          <xsl:with-param name="ConditionalCellSingle">
            <xsl:value-of select="$ConditionalCellSingle"/>
          </xsl:with-param>
          <xsl:with-param name="ConditionalCellMultiple">
            <xsl:value-of select="$ConditionalCellMultiple"/>
          </xsl:with-param>
          <xsl:with-param name="sheetNr" select="$sheetNr"/>
          <xsl:with-param name="EndRow">
            <xsl:value-of select="@r - 1"/>
          </xsl:with-param>
          <xsl:with-param name="ValidationCell">
            <xsl:value-of select="$ValidationCell"/>
          </xsl:with-param>
          <xsl:with-param name="ValidationRow">
            <xsl:value-of select="$ValidationRow"/>
          </xsl:with-param>
          <xsl:with-param name="ValidationCellStyle">
            <xsl:value-of select="$ValidationCellStyle"/>
          </xsl:with-param>
          <xsl:with-param name="AllRowBreakes">
            <xsl:value-of select="$AllRowBreakes"/>
          </xsl:with-param>
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
          <xsl:with-param name="MergeCell">
            <xsl:value-of select="$MergeCell"/>
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
          <xsl:with-param name="MergeCell">
            <xsl:value-of select="$MergeCell"/>
          </xsl:with-param>
          <xsl:with-param name="Repeat">
            <xsl:value-of select="@r - 2"/>
          </xsl:with-param>

        </xsl:call-template>

      </xsl:when>

      <!-- when this row is not first one and there were pictures rows after previous non-empty row-->
      <xsl:when
        test="preceding-sibling::e:row[1]/@r &lt;  @r - 1 and $GetMinRowWithElement &gt; preceding-sibling::e:row[1]/@r and $GetMinRowWithElement &lt; @r - 1">

        <xsl:call-template name="InsertElementsBetwenTwoRows">
          <xsl:with-param name="sheet">
            <xsl:value-of select="$sheet"/>
          </xsl:with-param>
          <xsl:with-param name="NameSheet">
            <xsl:value-of select="$NameSheet"/>
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
          <xsl:with-param name="ConditionalCell">
            <xsl:value-of select="$ConditionalCell"/>
          </xsl:with-param>
          <xsl:with-param name="ConditionalCellStyle">
            <xsl:value-of select="$ConditionalCellStyle"/>
          </xsl:with-param>
          <xsl:with-param name="ConditionalRow">
            <xsl:value-of select="$ConditionalRow"/>
          </xsl:with-param>
          <xsl:with-param name="ConditionalCellCol">
            <xsl:value-of select="$ConditionalCellCol"/>
          </xsl:with-param>
          <xsl:with-param name="ConditionalCellAll">
            <xsl:value-of select="$ConditionalCellAll"/>
          </xsl:with-param>
          <xsl:with-param name="ConditionalCellSingle">
            <xsl:value-of select="$ConditionalCellSingle"/>
          </xsl:with-param>
          <xsl:with-param name="ConditionalCellMultiple">
            <xsl:value-of select="$ConditionalCellMultiple"/>
          </xsl:with-param>
          <xsl:with-param name="sheetNr" select="$sheetNr"/>
          <xsl:with-param name="EndRow">
            <xsl:value-of select="@r - 1"/>
          </xsl:with-param>
          <xsl:with-param name="prevRow">
            <xsl:value-of select="preceding-sibling::e:row[1]/@r"/>
          </xsl:with-param>
          <xsl:with-param name="ValidationCell">
            <xsl:value-of select="$ValidationCell"/>
          </xsl:with-param>
          <xsl:with-param name="ValidationRow">
            <xsl:value-of select="$ValidationRow"/>
          </xsl:with-param>
          <xsl:with-param name="ValidationCellStyle">
            <xsl:value-of select="$ValidationCellStyle"/>
          </xsl:with-param>
          <xsl:with-param name="AllRowBreakes">
            <xsl:value-of select="$AllRowBreakes"/>
          </xsl:with-param>
        </xsl:call-template>

      </xsl:when>

      <xsl:otherwise>

        <!-- when this row is not first one and there were empty rows after previous non-empty row -->
        <xsl:if test="preceding-sibling::e:row[1]/@r &lt;  @r - 1">
          <table:table-row table:style-name="{concat('ro', key('Part', concat('xl/',$sheet))/e:worksheet/@oox:part)}"
            table:number-rows-repeated="{@r -1 - preceding-sibling::e:row[1]/@r}">
            <table:table-cell table:number-columns-repeated="256"/>
          </table:table-row>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>

    <xsl:if test="contains(concat(':', $GroupRowStart), concat(':', @r, ':'))">
      <xsl:call-template name="InsertRowGroupStart">
        <xsl:with-param name="GroupCell">
          <xsl:value-of select="$GroupRowStart"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if>

    <!-- insert this row -->
    <xsl:call-template name="InsertThisRow">
      <xsl:with-param name="BigMergeCell">
        <xsl:value-of select="$BigMergeCell"/>
      </xsl:with-param>
      <xsl:with-param name="BigMergeRow">
        <xsl:value-of select="$BigMergeRow"/>
      </xsl:with-param>
      <xsl:with-param name="MergeCell">
        <xsl:value-of select="$MergeCell"/>
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
      <xsl:with-param name="ConditionalCellCol">
        <xsl:value-of select="$ConditionalCellCol"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellAll">
        <xsl:value-of select="$ConditionalCellAll"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellSingle">
        <xsl:value-of select="$ConditionalCellSingle"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellMultiple">
        <xsl:value-of select="$ConditionalCellMultiple"/>
      </xsl:with-param>
      <xsl:with-param name="removeFilter" select="$removeFilter"/>
      <xsl:with-param name="ValidationCell">
        <xsl:value-of select="$ValidationCell"/>
      </xsl:with-param>
      <xsl:with-param name="ValidationRow">
        <xsl:value-of select="$ValidationRow"/>
      </xsl:with-param>
      <xsl:with-param name="ValidationCellStyle">
        <xsl:value-of select="$ValidationCellStyle"/>
      </xsl:with-param>
      <xsl:with-param name="ConnectionsCell">
        <xsl:value-of select="$ConnectionsCell"/>
      </xsl:with-param>
      <xsl:with-param name="outlineLevel"/>
      <xsl:with-param name="AllRowBreakes">
        <xsl:value-of select="$AllRowBreakes"/>
      </xsl:with-param>
      <xsl:with-param name="rSheredStrings">
        <xsl:value-of select="$rSheredStrings"/>
      </xsl:with-param>
    </xsl:call-template>

    <xsl:if test="contains(concat(':', $GroupRowEnd), concat(':', @r, ':'))">
      <xsl:call-template name="InsertRowGroupEnd">
        <xsl:with-param name="GroupCell">
          <xsl:value-of select="$GroupRowEnd"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if>

    <xsl:if
      test="not(following-sibling::e:row) and ($PictureRow != '' or $NoteRow != '' or $ConditionalRow != '' or $ValidationRow != '' or $AllRowBreakes!='' )">

      <xsl:variable name="GetMinRowWithElementAfterLastRow">
        <xsl:call-template name="GetMinRowWithPicture">
          <xsl:with-param name="PictureRow">
            <xsl:value-of select="concat($PictureRow, $NoteRow, $ConditionalRow, $AllRowBreakes)"/>
          </xsl:with-param>
          <xsl:with-param name="AfterRow">
            <xsl:value-of select="@r"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>

      <xsl:call-template name="InsertElementsBetwenTwoRows">
        <xsl:with-param name="sheet">
          <xsl:value-of select="$sheet"/>
        </xsl:with-param>
        <xsl:with-param name="NameSheet">
          <xsl:value-of select="$NameSheet"/>
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
        <xsl:with-param name="ConditionalCell">
          <xsl:value-of select="$ConditionalCell"/>
        </xsl:with-param>
        <xsl:with-param name="ConditionalCellStyle">
          <xsl:value-of select="$ConditionalCellStyle"/>
        </xsl:with-param>
        <xsl:with-param name="ConditionalRow">
          <xsl:value-of select="$ConditionalRow"/>
        </xsl:with-param>
        <xsl:with-param name="ConditionalCellCol">
          <xsl:value-of select="$ConditionalCellCol"/>
        </xsl:with-param>
        <xsl:with-param name="ConditionalCellAll">
          <xsl:value-of select="$ConditionalCellAll"/>
        </xsl:with-param>
        <xsl:with-param name="ConditionalCellSingle">
          <xsl:value-of select="$ConditionalCellSingle"/>
        </xsl:with-param>
        <xsl:with-param name="ConditionalCellMultiple">
          <xsl:value-of select="$ConditionalCellMultiple"/>
        </xsl:with-param>
        <xsl:with-param name="sheetNr" select="$sheetNr"/>
        <xsl:with-param name="EndRow">
          <xsl:value-of select="65535"/>
        </xsl:with-param>
        <xsl:with-param name="prevRow">
          <xsl:value-of select="@r"/>
        </xsl:with-param>
        <xsl:with-param name="ValidationCell">
          <xsl:value-of select="$ValidationCell"/>
        </xsl:with-param>
        <xsl:with-param name="ValidationRow">
          <xsl:value-of select="$ValidationRow"/>
        </xsl:with-param>
        <xsl:with-param name="ValidationCellStyle">
          <xsl:value-of select="$ValidationCellStyle"/>
        </xsl:with-param>
        <xsl:with-param name="AllRowBreakes">
          <xsl:value-of select="$AllRowBreakes"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

  <xsl:template match="e:row" mode="headers">
    <xsl:param name="BigMergeCell"/>
    <xsl:param name="BigMergeRow"/>
    <xsl:param name="MergeCell"/>
    <xsl:param name="headerRowsStart"/>
    <xsl:param name="headerRowsEnd"/>
    <xsl:param name="PictureRow"/>
    <xsl:param name="sheetNr"/>
    <xsl:param name="removeFilter"/>
    <xsl:param name="rSheredStrings"/>

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
        <table:table-row table:style-name="{concat('ro', ancestor::e:worksheet/@oox:part)}"
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
              <xsl:with-param name="MergeCell">
                <xsl:value-of select="$MergeCell"/>
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
              <xsl:with-param name="MergeCell">
                <xsl:value-of select="$MergeCell"/>
              </xsl:with-param>
              <xsl:with-param name="Repeat">
                <xsl:value-of select="@r - 2"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>

          <!-- if the first non-empty header row isn't the first header row -->
          <xsl:when test="position() = 1 and @r &gt; $headerRowsStart">
            <table:table-row table:style-name="{concat('ro', ancestor::e:worksheet/@oox:part)}"
              table:number-rows-repeated="{@r - $headerRowsStart}">
              <table:table-cell table:number-columns-repeated="256"/>
            </table:table-row>
          </xsl:when>

          <!-- if there are empty header rows before this one header row-->
          <xsl:when test="position() &gt; 1 and @r - 1 != preceding-sibling::e:row[1]/@r">
            <table:table-row table:style-name="{concat('ro', ancestor::e:worksheet/@oox:part)}"
              table:number-rows-repeated="{@r - preceding-sibling::e:row[1]/@r - 1}">
              <table:table-cell table:number-columns-repeated="256"/>
            </table:table-row>
          </xsl:when>
        </xsl:choose>
      </xsl:when>

      <!-- when this row is the first non-empty one after header rows and there aren't Big Merge Coll -->
      <xsl:when test="position()=1 and @r &gt; $headerRowsEnd + 1 and $BigMergeCell = '' ">
        <table:table-row table:style-name="{concat('ro', ancestor::e:worksheet/@oox:part)}"
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
          <xsl:with-param name="MergeCell">
            <xsl:value-of select="$MergeCell"/>
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
          <xsl:with-param name="MergeCell">
            <xsl:value-of select="$MergeCell"/>
          </xsl:with-param>
          <xsl:with-param name="Repeat">
            <xsl:value-of select="@r - 2"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <!-- when this row is not first one and there were empty rows after previous non-empty row -->
      <xsl:when test="position() != 1 and @r != preceding-sibling::e:row[1]/@r + 1">
        <table:table-row table:style-name="{concat('ro', ancestor::e:worksheet/@oox:part)}"
          table:number-rows-repeated="{@r -1 - preceding-sibling::e:row[1]/@r}">
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
      <xsl:with-param name="MergeCell">
        <xsl:value-of select="$MergeCell"/>
      </xsl:with-param>
      <xsl:with-param name="lastCellColumnNumber">
        <xsl:value-of select="$lastCellColumnNumber"/>
      </xsl:with-param>
      <xsl:with-param name="CheckIfBigMerge">
        <xsl:value-of select="$CheckIfBigMerge"/>
      </xsl:with-param>
      <xsl:with-param name="rSheredStrings">
        <xsl:value-of select="$rSheredStrings"/>
      </xsl:with-param>

      <xsl:with-param name="this" select="$this"/>
      <xsl:with-param name="headerRowsStart" select="$headerRowsStart"/>
      <xsl:with-param name="sheetNr" select="$sheetNr"/>
      <xsl:with-param name="removeFilter" select="$removeFilter"/>
    </xsl:call-template>

  </xsl:template>

  <xsl:template name="ConvertCell" match="e:c">
    <xsl:param name="BeforeMerge"/>
    <xsl:param name="prevCellCol"/>
    <xsl:param name="BigMergeCell"/>
    <xsl:param name="MergeCell"/>
    <xsl:param name="PictureCell"/>
    <xsl:param name="PictureRow"/>
    <xsl:param name="NoteCell"/>
    <xsl:param name="NoteRow"/>
    <xsl:param name="sheet"/>
    <xsl:param name="NameSheet"/>
    <xsl:param name="sheetNr"/>
    <xsl:param name="ConditionalCell"/>
    <xsl:param name="ConditionalCellStyle"/>
    <xsl:param name="ConditionalCellCol"/>
    <xsl:param name="ConditionalCellAll"/>
    <xsl:param name="ConditionalCellSingle"/>
    <xsl:param name="ConditionalCellMultiple"/>
    <xsl:param name="ValidationCell"/>
    <xsl:param name="ValidationRow"/>
    <xsl:param name="ValidationCellStyle"/>
    <xsl:param name="ConnectionsCell"/>
    <xsl:param name="rSheredStrings"/>

    <xsl:variable name="this" select="."/>

    <xsl:variable name="colNum" >
      <xsl:call-template name="GetColNum2">
        <xsl:with-param name="cell">
          <xsl:value-of select="@oox:p"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="rowNum" >
      <xsl:call-template name="GetRowNum2">
        <xsl:with-param name="cell">
          <xsl:value-of select="@oox:p"/>
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
        <xsl:with-param name="MergeCell">
          <xsl:value-of select="$MergeCell"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="PictureColl">
      <xsl:call-template name="GetCollsWithElement">
        <xsl:with-param name="rowNumber">
          <xsl:value-of select="$rowNum"/>
        </xsl:with-param>
        <xsl:with-param name="ElementCell">
          <xsl:value-of select="concat(';', $PictureCell)"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="NoteColl">
      <xsl:call-template name="GetCollsWithElement">
        <xsl:with-param name="rowNumber">
          <xsl:value-of select="$rowNum"/>
        </xsl:with-param>
        <xsl:with-param name="ElementCell">
          <xsl:value-of select="concat(';', $NoteCell)"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="ConditionalColl">
      <xsl:call-template name="GetCollsWithElement">
        <xsl:with-param name="rowNumber">
          <xsl:value-of select="$rowNum"/>
        </xsl:with-param>
        <xsl:with-param name="ElementCell">
          <xsl:value-of select="concat(';', $ConditionalCell)"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="ValidationColl">
      <xsl:call-template name="GetCollsWithElement">
        <xsl:with-param name="rowNumber">
          <xsl:value-of select="$rowNum"/>
        </xsl:with-param>
        <xsl:with-param name="ElementCell">
          <xsl:value-of select="concat(';', $ValidationCell)"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="GetMinCollWithElement">
      <xsl:call-template name="GetMinRowWithPicture">
        <xsl:with-param name="PictureRow">
          <xsl:value-of select="concat($PictureColl, $NoteColl, $ValidationColl, $ConditionalColl)"/>
        </xsl:with-param>
        <xsl:with-param name="AfterRow">
          <xsl:choose>
            <xsl:when test="$prevCellCol = ''">
              <xsl:text>0</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$prevCellCol"/>
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
      <xsl:with-param name="MergeCell">
        <xsl:value-of select="$MergeCell"/>
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
      <xsl:with-param name="sheetNr" select="$sheetNr"/>
      <xsl:with-param name="ConditionalCell">
        <xsl:value-of select="$ConditionalCell"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellStyle">
        <xsl:value-of select="$ConditionalCellStyle"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCol">
        <xsl:value-of select="$ConditionalColl"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellCol">
        <xsl:value-of select="$ConditionalCellCol"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellAll">
        <xsl:value-of select="$ConditionalCellAll"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellSingle">
        <xsl:value-of select="$ConditionalCellSingle"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellMultiple">
        <xsl:value-of select="$ConditionalCellMultiple"/>
      </xsl:with-param>
      <xsl:with-param name="ValidationCell">
        <xsl:value-of select="$ValidationCell"/>
      </xsl:with-param>
      <xsl:with-param name="ValidationRow">
        <xsl:value-of select="$ValidationRow"/>
      </xsl:with-param>
      <xsl:with-param name="ValidationCellStyle">
        <xsl:value-of select="$ValidationCellStyle"/>
      </xsl:with-param>
      <xsl:with-param name="ValidationColl">
        <xsl:value-of select="$ValidationColl"/>
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
      <xsl:with-param name="MergeCell">
        <xsl:value-of select="$MergeCell"/>
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
      <xsl:with-param name="GetMinCollWithPicture">
        <xsl:value-of select="$GetMinCollWithElement"/>
      </xsl:with-param>
      <xsl:with-param name="sheetNr" select="$sheetNr"/>
      <xsl:with-param name="ConditionalCell">
        <xsl:value-of select="$ConditionalCell"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellStyle">
        <xsl:value-of select="$ConditionalCellStyle"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellCol">
        <xsl:value-of select="$ConditionalCellCol"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellAll">
        <xsl:value-of select="$ConditionalCellAll"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellSingle">
        <xsl:value-of select="$ConditionalCellSingle"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellMultiple">
        <xsl:value-of select="$ConditionalCellMultiple"/>
      </xsl:with-param>
      <xsl:with-param name="ValidationCell">
        <xsl:value-of select="$ValidationCell"/>
      </xsl:with-param>
      <xsl:with-param name="ValidationRow">
        <xsl:value-of select="$ValidationRow"/>
      </xsl:with-param>
      <xsl:with-param name="ValidationCellStyle">
        <xsl:value-of select="$ValidationCellStyle"/>
      </xsl:with-param>
      <xsl:with-param name="ConnectionsCell">
        <xsl:value-of select="$ConnectionsCell"/>
      </xsl:with-param>
      <xsl:with-param name="rSheredStrings">
        <xsl:value-of select="$rSheredStrings"/>
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
      <xsl:with-param name="MergeCell">
        <xsl:value-of select="$MergeCell"/>
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
      <xsl:with-param name="GetMinCollWithElement">
        <xsl:value-of select="$GetMinCollWithElement"/>
      </xsl:with-param>
      <xsl:with-param name="sheetNr" select="$sheetNr"/>
      <xsl:with-param name="ConditionalCell">
        <xsl:value-of select="$ConditionalCell"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellStyle">
        <xsl:value-of select="$ConditionalCellStyle"/>
      </xsl:with-param>
      <xsl:with-param name="ConnectionsCell">
        <xsl:value-of select="$ConnectionsCell"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellCol">
        <xsl:value-of select="$ConditionalCellCol"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellAll">
        <xsl:value-of select="$ConditionalCellAll"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellSingle">
        <xsl:value-of select="$ConditionalCellSingle"/>
      </xsl:with-param>
      <xsl:with-param name="ConditionalCellMultiple">
        <xsl:value-of select="$ConditionalCellMultiple"/>
      </xsl:with-param>
      <xsl:with-param name="ValidationCell">
        <xsl:value-of select="$ValidationCell"/>
      </xsl:with-param>
      <xsl:with-param name="ValidationRow">
        <xsl:value-of select="$ValidationRow"/>
      </xsl:with-param>
      <xsl:with-param name="ValidationCellStyle">
        <xsl:value-of select="$ValidationCellStyle"/>
      </xsl:with-param>
      <xsl:with-param name="ValidationColl">
        <xsl:value-of select="$ValidationColl"/>
      </xsl:with-param>
      <xsl:with-param name="rSheredStrings">
        <xsl:value-of select="$rSheredStrings"/>
      </xsl:with-param>
    </xsl:call-template>

  </xsl:template>


  <!-- convert run into span -->
  <xsl:template match="e:r">
    <!--RefNo-1: Changes for fixing 1833074 XLSX: Cell Content missing-->
    <!--<xsl:param name ="textp"/>-->
    <!-- 
	        Bug No.          :1805599
		    Code Modified By:Vijayeta
			Date            :6th Nov '07
			Description     :New Line to be added, for which the comment is checked for new lines.
			                If present Post Processor in the CS file 'OdfSharedStringsPostProcessor.cs' is called
	   -->
    <xsl:variable name="textContent">
      <xsl:value-of select="./e:t"/>
    </xsl:variable>
    <xsl:choose >
      <xsl:when test ="contains($textContent, '&#xA;')">
        <xsl:variable name ="Id" >
          <xsl:value-of select ="generate-id(.)"/>
        </xsl:variable>
        <xsl:value-of select ="concat('SonataAnnotation|',$textContent,'|',$Id)"/>
      </xsl:when>
      <!--Start of RefNo-1-->
      <!--<xsl:when test ="$textp='T'">
        <text:span>
          <xsl:if test="e:rPr">
            <xsl:attribute name="text:style-name">
              <xsl:value-of select="generate-id(.)"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:apply-templates/>
        </text:span>
      </xsl:when>-->
      <!--End of RefNo-1-->
      <xsl:otherwise >
        <!--RefNo-1:Commented
        <text:p text:style-name="{generate-id(./parent::node())}">-->
        <text:span>
          <xsl:if test="e:rPr">
            <xsl:attribute name="text:style-name">
              <xsl:value-of select="generate-id(.)"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:apply-templates/>
        </text:span>
        <!--</text:p >-->
      </xsl:otherwise>
    </xsl:choose>
    <!-- End of modification for the bug 1805599-->

    <!--<text:span>
      <xsl:if test="e:rPr">
        <xsl:attribute name="text:style-name">
          <xsl:value-of select="generate-id(.)"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:apply-templates/>
    </text:span>-->
  </xsl:template>

  <!-- convert run into span in hyperlink-->
  <xsl:template match="e:r" mode="hyperlink">
    <xsl:param name="XlinkHref"/>
    <xsl:param name="position"/>

    <xsl:choose>
      <xsl:when test="$position = '1'">
        <text:span>
          <xsl:if test="e:rPr">
            <xsl:attribute name="text:style-name">
              <xsl:value-of select="generate-id(.)"/>
            </xsl:attribute>
          </xsl:if>
          <text:a>
            <xsl:attribute name="xlink:href">
              <xsl:value-of select="$XlinkHref"/>
            </xsl:attribute>
            <xsl:apply-templates/>
            <xsl:if test="following-sibling::e:r">
              <xsl:apply-templates select="following-sibling::e:r[1]" mode="hyperlink">
                <xsl:with-param name="XlinkHref">
                  <xsl:value-of select="$XlinkHref"/>
                </xsl:with-param>
                <xsl:with-param name="position">
                  <xsl:value-of select="$position+1"/>
                </xsl:with-param>
              </xsl:apply-templates>
            </xsl:if>
          </text:a>
        </text:span>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates/>
        <xsl:if test="following-sibling::e:r">
          <xsl:apply-templates select="following-sibling::e:r[1]" mode="hyperlink">
            <xsl:with-param name="XlinkHref">
              <xsl:value-of select="$XlinkHref"/>
            </xsl:with-param>
            <xsl:with-param name="position">
              <xsl:value-of select="$position+1"/>
            </xsl:with-param>
          </xsl:apply-templates>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>


  </xsl:template>

  <xsl:template name="CheckSheetName">
    <xsl:param name="name"/>
    <xsl:param name="sheetNumber"/>

    <xsl:choose>
      <!-- when there are at least 2 sheets with the same name after removal of forbidden characters and cutting to 31 characters (name correction) -->
      <xsl:when test="parent::node()/e:sheet[translate(@name,'!$-():#,.+','') = $name][2]">
        <xsl:variable name="nameConflictsBefore">
          <!-- count sheets before this one whose name (after correction) collide with this sheet name (after correction) -->
          <xsl:value-of
            select="count(parent::node()/e:sheet[translate(@name,'!$-():#,.+','') = $name and position() &lt; $sheetNumber])"
          />
        </xsl:variable>
        <!-- cut name and add "(N)" at the end where N is seqential number of duplicated name -->
        <xsl:value-of select="concat(translate(@name,'!$-()#:,.+',''),'_',$nameConflictsBefore + 1)"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$name"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="rSheredStrings">
    <xsl:param name="result"/>

    <xsl:for-each select="key('Part', 'xl/sharedStrings.xml')/e:sst">

      <xsl:apply-templates select="e:si[1]" mode="r">
        <xsl:with-param name="result">
          <xsl:choose>
            <xsl:when test="e:r">
              <xsl:value-of select="concat($result, '0', ';')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$result"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
        <xsl:with-param name="position">
          <xsl:text>0</xsl:text>
        </xsl:with-param>
      </xsl:apply-templates>

    </xsl:for-each>

  </xsl:template>


  <xsl:template match="e:si" mode="r">
    <xsl:param name="result"/>
    <xsl:param name="position"/>

    <xsl:choose>
      <xsl:when test="following-sibling::e:si/e:r">
        <xsl:apply-templates select="following-sibling::e:si[1]" mode="r">
          <xsl:with-param name="result">
            <xsl:choose>
              <xsl:when test="e:r">
                <xsl:value-of select="concat($result, $position, ';')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$result"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="position">
            <xsl:value-of select="$position + 1"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="e:r">
            <xsl:value-of select="concat($result, $position, ';')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$result"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

</xsl:stylesheet>

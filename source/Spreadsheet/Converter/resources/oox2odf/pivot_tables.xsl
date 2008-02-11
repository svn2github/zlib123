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
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main" exclude-result-prefixes="e r">

  <xsl:import href="relationships.xsl"/>
  <xsl:import href="common.xsl"/>
  <xsl:import href="measures.xsl"/>


  <xsl:template name="InsertPilotTables">

    <xsl:variable name="pivotSheets">
      <xsl:call-template name="PivotSheets"/>
    </xsl:variable>

    <xsl:if test="$pivotSheets !=''">
      <table:data-pilot-tables>
        <xsl:call-template name="CreatePilotTables">
          <xsl:with-param name="pivotSheets" select="$pivotSheets"/>
        </xsl:call-template>
      </table:data-pilot-tables>
    </xsl:if>

  </xsl:template>

  <xsl:template name="PivotSheets">
    <!-- @Description: Searches for all Pivot tables  within workbook and starts conversion. -->
    <!-- @Context: None -->

    <!-- get all sheet Id's -->
    <xsl:for-each select="key('Part', 'xl/workbook.xml')/e:workbook/e:sheets/e:sheet">

      <xsl:variable name="sheet">
        <!-- path to sheet file from xl/ catalog (i.e. $sheet = worksheets/sheet1.xml) -->
        <xsl:call-template name="GetTarget">
          <xsl:with-param name="id" select="@r:id"/>
          <xsl:with-param name="document">
            <xsl:text>xl/workbook.xml</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>

      <xsl:variable name="fileName">
        <xsl:value-of select="substring-after(substring-before($sheet, '.'),'/')"/>
      </xsl:variable>

      <xsl:if
        test="key('Part', concat('xl/worksheets/_rels/',$fileName,'.xml.rels'))//node()[name()='Relationship' and contains(@Type,'pivotTable' )]/@Target">
        <xsl:value-of select="position()"/>
        <xsl:text>,</xsl:text>
      </xsl:if>

    </xsl:for-each>
  </xsl:template>

  <xsl:template name="CreatePilotTables">
    <xsl:param name="pivotSheets"/>

    <xsl:if test="$pivotSheets != '' ">
      <xsl:call-template name="InsertSheetPilotTables">
        <xsl:with-param name="sheetNum" select="substring-before($pivotSheets,',')"/>
      </xsl:call-template>
      <xsl:call-template name="CreatePilotTables">
        <xsl:with-param name="pivotSheets" select="substring-after($pivotSheets,',')"/>
      </xsl:call-template>
    </xsl:if>

  </xsl:template>

  <!-- search  target Pivot file-->
  <xsl:template name="InsertSheetPilotTables">
    <xsl:param name="sheetNum"/>

    <xsl:for-each
      select="key('Part', 'xl/workbook.xml')/e:workbook/e:sheets/e:sheet[position() = $sheetNum]">

      <xsl:variable name="sheetName">
        <xsl:call-template name="CheckSheetName">
          <xsl:with-param name="sheetNumber">
            <xsl:value-of select="$sheetNum"/>
          </xsl:with-param>
          <xsl:with-param name="name">
            <xsl:value-of select="translate(@name,'!-$#:(),.+','')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>

      <xsl:variable name="sheet">
        <!-- path to sheet file from xl/ catalog (i.e. $sheet = worksheets/sheet1.xml) -->
        <xsl:call-template name="GetTarget">
          <xsl:with-param name="id" select="@r:id"/>
          <xsl:with-param name="document">
            <xsl:text>xl/workbook.xml</xsl:text>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>

      <xsl:variable name="fileName">
        <xsl:value-of select="substring-after(substring-before($sheet, '.'),'/')"/>
      </xsl:variable>

      <xsl:for-each
        select="key('Part', concat('xl/worksheets/_rels/',$fileName,'.xml.rels'))//node()[name()='Relationship' and contains(@Type,'pivotTable' )]">

        <xsl:variable name="TargetPilotFile">
          <xsl:value-of select="substring-after(@Target,'../')"/>
        </xsl:variable>
		  <xsl:variable name="cacheFile">
			  <xsl:for-each
				select="key('Part', concat('xl/pivotTables/_rels/',substring-after($TargetPilotFile,'/'),'.rels'))//node()[name()='Relationship' and contains(@Type,'pivotCacheDefinition' )]">
				  <xsl:value-of select="substring-after(@Target,'../')"/>
			  </xsl:for-each>
		  </xsl:variable>
		  <!-- Vijayeta, 1803593,sales.xlsx,konto2006.xlsx-->
		  <xsl:variable name ="OneMoreValidation">
			  <xsl:call-template name ="oneMoreValidation">
				  <xsl:with-param name ="cacheFile">
					  <xsl:value-of select ="$cacheFile"/>
				  </xsl:with-param>
			  </xsl:call-template>
		  </xsl:variable>
		  <!-- Vijayeta, 1803593,sales.xlsx-->
        <xsl:if test="key('Part', concat('xl/',$TargetPilotFile))/e:pivotTableDefinition">
          <xsl:for-each select="key('Part', concat('xl/',$TargetPilotFile))/e:pivotTableDefinition">

            <xsl:variable name="name">
              <xsl:value-of select="@name"/>
            </xsl:variable>

            <xsl:variable name="numberOfAxisPage">
              <xsl:for-each select="e:pivotFields">
                <xsl:value-of select="count(e:pivotField[@axis ='axisPage'])"/>
              </xsl:for-each>
            </xsl:variable>
				  <!-- Vijayeta, 1803593,sales.xlsx,konto2006.xlsx-->
				  <xsl:if test ="$OneMoreValidation='true'">
            <table:data-pilot-table table:name="{$name}">

              <xsl:for-each select="e:location">

                <xsl:variable name="firstTargetAdress">
                  <xsl:value-of select="substring-before(@ref,':')"/>
                </xsl:variable>

                <xsl:variable name="lastTargetAdress">
                  <xsl:value-of select="substring-after(@ref,':')"/>
                </xsl:variable>

                <xsl:variable name="numRow">
                  <xsl:call-template name="GetRowNum">
                    <xsl:with-param name="cell">
                      <xsl:value-of select="$firstTargetAdress"/>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:variable>

                <xsl:variable name="ifPageRow">
                  <xsl:choose>
                    <xsl:when test="parent::node()/e:pivotFields/e:pivotField[@axis ='axisPage']">
                      <xsl:value-of select="$numRow - 1 - $numberOfAxisPage"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$numRow"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:variable>

                <xsl:variable name="numCol">
                  <xsl:call-template name="GetColNum">
                    <xsl:with-param name="cell">
                      <xsl:value-of select="$firstTargetAdress"/>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:variable>

                <xsl:variable name="firstButtonCellCol">
                  <xsl:call-template name="NumbersToChars">
                    <xsl:with-param name="num">
                      <xsl:value-of select="$numCol - 1"/>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:variable>

                <xsl:variable name="nextButtonCellCol">
                  <xsl:call-template name="NumbersToChars">
                    <xsl:with-param name="num">
                      <xsl:value-of select="$numCol"/>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:variable>

                <xsl:variable name="apos">
                  <xsl:text>&apos;</xsl:text>
                </xsl:variable>

                <xsl:attribute name="table:target-range-address">
                  <xsl:value-of
                    select="concat($apos,$sheetName,$apos,'.',$firstButtonCellCol,$ifPageRow,':',$apos,$sheetName, $apos,'.',$lastTargetAdress)"
                  />
                </xsl:attribute>

                <!--xsl:variable name="pageButton">
                  <xsl:if test="parent::node()/e:pivotFields/e:pivotField[@axis ='axisPage']">
                  <xsl:text> 1 </xsl:text>
                  </xsl:if>
                  </xsl:variable>
                  
                  <xsl:variable name="rowButton">
                  <xsl:if test="parent::node()/e:pivotFields/e:pivotField[@axis ='axisRow']">
                  <xsl:text> 1 </xsl:text>
                  </xsl:if>
                  </xsl:variable>
                  
                  <xsl:variable name="colButton">
                  <xsl:if test="parent::node()/e:pivotFields/e:pivotField[@axis ='axisCol']">
                  <xsl:text> 1 </xsl:text>
                  </xsl:if>
                  </xsl:variable-->


                <xsl:attribute name="table:buttons">

                  <xsl:value-of
                    select="concat($apos,$sheetName,$apos,'.',$firstButtonCellCol,$numRow,' ',$apos,$sheetName,$apos,'.',$nextButtonCellCol,$numRow)"/>

                </xsl:attribute>

                <xsl:attribute name="table:show-filter-button">
                  <xsl:text>false</xsl:text>
                </xsl:attribute>

              </xsl:for-each>

              <!-- locate cache for this pivot table -->
						  <!--<xsl:variable name="cacheFile">
                <xsl:for-each
                  select="key('Part', concat('xl/pivotTables/_rels/',substring-after($TargetPilotFile,'/'),'.rels'))//node()[name()='Relationship' and contains(@Type,'pivotCacheDefinition' )]">

                  <xsl:value-of select="substring-after(@Target,'../')"/>

                </xsl:for-each>
              </xsl:variable>-->


              <!-- insert pilot table source range -->
              <!--Defect Id       :1803578
			        * Code Changed by :Sateesh Reddy
			        * Date            :05th Jan '08
			        * Description     :This part of code was added because when a PivotCacheDefinition contains external link file is crashing in round trip.
                                 To avoid crash I added(xsl:if condition) this code.
				      -->
              <xsl:if test="not(key('Part',concat('xl/pivotCache/_rels/',substring-after($cacheFile,'/'),'.rels'))//node()[name()='Relationship' and contains(@Type,'externalLinkPath') and @TargetMode])">">
               
              
              <xsl:for-each select="key('Part', concat('xl/',$cacheFile))/e:pivotCacheDefinition">
                <xsl:for-each select="e:cacheSource/e:worksheetSource">
                  <xsl:variable name="apos">
                    <xsl:text>&apos;</xsl:text>
                  </xsl:variable>

                  <xsl:variable name="sheetSourceName">
                    <xsl:value-of select="@sheet"/>
                  </xsl:variable>

                  <xsl:variable name="firstSourceAdress">
                    <xsl:value-of select="substring-before(@ref,':')"/>
                  </xsl:variable>

                  <xsl:variable name="lastSourceAdress">
                    <xsl:value-of select="substring-after(@ref,':')"/>
                  </xsl:variable>

                  <table:source-cell-range>

                    <xsl:attribute name="table:cell-range-address">
                      <xsl:value-of
                        select="concat($apos,$sheetSourceName,$apos,'.',$firstSourceAdress,':',$apos,$sheetSourceName,$apos,'.',$lastSourceAdress)"
                      />
                    </xsl:attribute>

                    <!-- Insert Filters-->
                    <xsl:for-each
                      select="key('Part', concat('xl/',$TargetPilotFile))/e:pivotTableDefinition/e:filters">

                      <table:filter table:condition-source-range-address="">
                        <table:filter-or>

                          <xsl:for-each
                            select="key('Part', concat('xl/',$TargetPilotFile))/e:pivotTableDefinition/e:filters/e:filter">
                            <xsl:variable name="labelFilterNum">
                              <xsl:value-of select="@fld"/>
                            </xsl:variable>

                            <xsl:choose>
                              <xsl:when
                                test="e:autoFilter/e:filterColumn/e:customFilters/e:customFilter">
                                <xsl:for-each
                                  select="e:autoFilter/e:filterColumn/e:customFilters/e:customFilter">
                                  <table:filter-condition>
                                    <xsl:attribute name="table:operator">
                                      <xsl:call-template name="TranslateFilterOperator"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="table:field-number">
                                      <xsl:value-of select="$labelFilterNum"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="table:value">
                                      <xsl:call-template name="TranslateFilterValue"/>
                                    </xsl:attribute>
                                  </table:filter-condition>
                                </xsl:for-each>
                              </xsl:when>

                              <xsl:when test="e:autoFilter/e:filterColumn/e:top10">
                                <xsl:message terminate="no"
                                >translation.oox2odf.PivotFilter</xsl:message>
                              </xsl:when>

                              <xsl:when test="count(e:filters/e:filter) &gt; 1">
                                <table:filter-or>
                                  <xsl:for-each select="e:filters/e:filter">
                                    <table:filter-condition table:operator="=">
                                      <xsl:attribute name="table:field-number">
                                        <xsl:value-of select="$labelFilterNum"/>
                                      </xsl:attribute>
                                      <xsl:attribute name="table:value">
                                        <xsl:value-of select="@val"/>
                                      </xsl:attribute>
                                    </table:filter-condition>
                                  </xsl:for-each>
                                </table:filter-or>
                              </xsl:when>

                              <xsl:when test="e:autoFilter/e:filterColumn/e:customFilters/@and = 1">
                                <table:filter-and>
                                  <xsl:for-each
                                    select="e:autoFilter/e:filterColumn/e:customFilters/e:customFilter">
                                    <table:filter-condition>
                                      <xsl:attribute name="table:operator">
                                        <xsl:call-template name="TranslateFilterOperator"/>
                                      </xsl:attribute>
                                      <xsl:attribute name="table:field-number">
                                        <xsl:value-of select="$labelFilterNum"/>
                                      </xsl:attribute>
                                      <xsl:attribute name="table:value">
                                        <xsl:call-template name="TranslateFilterValue"/>
                                      </xsl:attribute>
                                    </table:filter-condition>
                                  </xsl:for-each>
                                </table:filter-and>
                              </xsl:when>

                              <xsl:when
                                test="e:autoFilter/e:filterColumn/e:customFilters/e:customFilter">
                                <xsl:for-each
                                  select="e:autoFilter/e:filterColumn/e:customFilters/e:customFilter">
                                  <table:filter-condition>
                                    <xsl:attribute name="table:operator">
                                      <xsl:call-template name="TranslateFilterOperator"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="table:field-number">
                                      <xsl:value-of select="$labelFilterNum"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="table:value">
                                      <xsl:call-template name="TranslateFilterValue"/>
                                    </xsl:attribute>
                                  </table:filter-condition>
                                </xsl:for-each>
                              </xsl:when>

                            </xsl:choose>
                          </xsl:for-each>

                        </table:filter-or>
                      </table:filter>
                    </xsl:for-each>

                  </table:source-cell-range>
                </xsl:for-each>
              </xsl:for-each>

              </xsl:if>
              <!-- insert pilot table fields -->
              <xsl:for-each select="e:pageFields/e:pageField">
                <xsl:choose>
                  <xsl:when test="@fld !='-2' ">
                    <xsl:call-template name="InsertDataPilotField">
                      <xsl:with-param name="cacheFile" select="$cacheFile"/>
                      <xsl:with-param name="orientation">
                        <xsl:text>page</xsl:text>
                      </xsl:with-param>
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:when test="@fld ='-2' ">
                    <xsl:call-template name="InsertEmptyDataPilotField">
                      <xsl:with-param name="orientation">
                        <xsl:text>page</xsl:text>
                      </xsl:with-param>
                    </xsl:call-template>
                  </xsl:when>
                </xsl:choose>
              </xsl:for-each>

              <xsl:for-each select="e:colFields/e:field">
                <xsl:choose>
                  <xsl:when test="@x !='-2' ">
                    <xsl:call-template name="InsertDataPilotField">
                      <xsl:with-param name="cacheFile" select="$cacheFile"/>
                      <xsl:with-param name="orientation">
                        <xsl:text>column</xsl:text>
                      </xsl:with-param>
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:when test="@x ='-2'">
                    <xsl:call-template name="InsertEmptyDataPilotField">
                      <xsl:with-param name="orientation">
                        <xsl:text>column</xsl:text>
                      </xsl:with-param>
                    </xsl:call-template>
                  </xsl:when>
                </xsl:choose>
              </xsl:for-each>

              <xsl:for-each select="e:rowFields/e:field">
                <xsl:choose>
                  <xsl:when test="@x !='-2' ">
                    <xsl:call-template name="InsertDataPilotField">
                      <xsl:with-param name="cacheFile" select="$cacheFile"/>
                      <xsl:with-param name="orientation">
                        <xsl:text>row</xsl:text>
                      </xsl:with-param>
                    </xsl:call-template>
                  </xsl:when>

                  <xsl:when test="@x ='-2'">
                    <xsl:call-template name="InsertEmptyDataPilotField">
                      <xsl:with-param name="orientation">
                        <xsl:text>row</xsl:text>
                      </xsl:with-param>
                    </xsl:call-template>
                  </xsl:when>
                </xsl:choose>
              </xsl:for-each>

              <xsl:for-each select="e:dataFields/e:dataField">
                <xsl:call-template name="InsertDataPilotField">
                  <xsl:with-param name="cacheFile" select="$cacheFile"/>
                  <xsl:with-param name="orientation">
                    <xsl:text>data</xsl:text>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:for-each>

            </table:data-pilot-table>
				  </xsl:if>
				  <!--End of Vijayeta, 1803593,sales.xlsx,konto2006.xlsx-->
          </xsl:for-each>
        </xsl:if>
      </xsl:for-each>

    </xsl:for-each>

  </xsl:template>

  <xsl:template name="InsertDataPilotField">
    <xsl:param name="cacheFile"/>
    <xsl:param name="orientation"/>

    <xsl:variable name="fieldNum">
      <xsl:choose>

        <xsl:when test="@x">
          <xsl:value-of select="@x"/>
        </xsl:when>

        <xsl:when test="@fld">
          <xsl:value-of select="@fld"/>
        </xsl:when>

      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="checkIfDataField">
      <xsl:for-each select="e:dataFields/e:dataField">
        <xsl:value-of select="@fld"/>
      </xsl:for-each>
    </xsl:variable>

    <!-- do not converted group fields and formula fields -->
    <xsl:if
      test="key('Part', concat('xl/',$cacheFile))/e:pivotCacheDefinition/e:cacheFields/e:cacheField[position() = $fieldNum + 1 and not(e:fieldGroup) and not(@formula)]">

      <table:data-pilot-field>

        <xsl:for-each
          select="key('Part', concat('xl/',$cacheFile))/e:pivotCacheDefinition/e:cacheFields/e:cacheField[position() = $fieldNum + 1]">

          <xsl:attribute name="table:source-field-name">

            <xsl:choose>
              <xsl:when test="number(translate(@name, ',' , '.' ))">

                <xsl:variable name="replaceDecimal">
                  <xsl:value-of select="format-number(translate(@name, ',' , '.' ),'0.##')"/>
                </xsl:variable>

                <xsl:value-of select="translate($replaceDecimal, '.' , ',' )"/>
              </xsl:when>

              <xsl:otherwise>
                <xsl:value-of select="@name"/>
              </xsl:otherwise>
            </xsl:choose>

          </xsl:attribute>
        </xsl:for-each>

        <xsl:attribute name="table:used-hierarchy">
          <xsl:text>0</xsl:text>
        </xsl:attribute>

        <xsl:attribute name="table:orientation">

          <xsl:value-of select="$orientation"/>
          <!--xsl:for-each
          select="parent::node()/parent::node()/e:pivotFields/e:pivotField[position() = $fieldNum + 1]">
          <xsl:choose>
            <xsl:when test="@axis = 'axisPage' ">
              <xsl:text>page</xsl:text>
            </xsl:when>
            <xsl:when test="@axis = 'axisRow' ">
              <xsl:text>row</xsl:text>
            </xsl:when>
            <xsl:when test="@axis = 'axisCol' ">
              <xsl:text>column</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>data</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each-->
        </xsl:attribute>

        <xsl:attribute name="table:function">

          <xsl:choose>
            <xsl:when
              test="parent::node()[name() = 'pageFields' or name() = 'rowFields'  or name() = 'colFields']">
              <xsl:text>auto</xsl:text>
            </xsl:when>

            <xsl:otherwise>
              <xsl:choose>

                <xsl:when test="@subtotal = 'count'">
                  <xsl:text>count</xsl:text>
                </xsl:when>

                <xsl:when test="@subtotal = 'average'">
                  <xsl:text>average</xsl:text>
                </xsl:when>

                <xsl:when test="@subtotal = 'max'">
                  <xsl:text>max</xsl:text>
                </xsl:when>

                <xsl:when test="@subtotal = 'min'">
                  <xsl:text>min</xsl:text>
                </xsl:when>

                <xsl:when test="@subtotal = 'product'">
                  <xsl:text>product</xsl:text>
                </xsl:when>

                <xsl:when test="@subtotal = 'countNums'">
                  <xsl:text>countnums</xsl:text>
                </xsl:when>

                <xsl:when test="@subtotal = 'stdDev'">
                  <xsl:text>stddev</xsl:text>
                </xsl:when>

                <xsl:when test="@subtotal = 'stdDevp'">
                  <xsl:text>stddevp</xsl:text>
                </xsl:when>

                <xsl:when test="@subtotal = 'varp'">
                  <xsl:text>varp</xsl:text>
                </xsl:when>

                <xsl:when test="@subtotal = 'var'">
                  <xsl:text>var</xsl:text>
                </xsl:when>

                <xsl:otherwise>
                  <xsl:text>sum</xsl:text>
                </xsl:otherwise>

              </xsl:choose>
            </xsl:otherwise>

          </xsl:choose>
        </xsl:attribute>

        <!-- if this field is a data field-->
        <xsl:if test="parent::e:dataFields">

          <xsl:variable name="base">
            <xsl:value-of select="@baseField"/>
          </xsl:variable>

          <xsl:variable name="member">
            <xsl:value-of select="@baseItem"/>
          </xsl:variable>

          <xsl:variable name="baseCacheNum">
            <xsl:for-each
              select="parent::node()/parent::node()/e:pivotFields/e:pivotField[position()=$base + 1]/e:items">
              <xsl:value-of select="child::node()[$member + 1]/@x"/>
            </xsl:for-each>
          </xsl:variable>

          <xsl:variable name="baseCacheValue">
            <xsl:for-each
              select="key('Part', concat('xl/',$cacheFile))/e:pivotCacheDefinition/e:cacheFields/e:cacheField[position() = $base + 1]/e:sharedItems">
              <xsl:value-of select="child::node()[$baseCacheNum + 1]/@v"/>
            </xsl:for-each>
          </xsl:variable>

          <xsl:if test="@showDataAs">

            <table:data-pilot-field-reference>

              <xsl:attribute name="table:type">
                <xsl:choose>

                  <xsl:when test="@showDataAs = 'difference' ">
                    <xsl:text>member-difference</xsl:text>
                  </xsl:when>

                  <xsl:when test="@showDataAs = 'percent' ">
                    <xsl:text>member-percentage</xsl:text>
                  </xsl:when>

                  <xsl:when test="@showDataAs = 'percentDiff' ">
                    <xsl:text>member-percentage-difference</xsl:text>
                  </xsl:when>

                  <xsl:when test="@showDataAs = 'runTotal' ">
                    <xsl:text>running-total</xsl:text>
                  </xsl:when>

                  <xsl:when test="@showDataAs = 'percentOfRow' ">
                    <xsl:text>row-percentage</xsl:text>
                  </xsl:when>

                  <xsl:when test="@showDataAs = 'percentOfColumn' ">
                    <xsl:text>column-percentage</xsl:text>
                  </xsl:when>

                  <xsl:when test="@showDataAs = 'percentOfTotal' ">
                    <xsl:text>total-percentage</xsl:text>
                  </xsl:when>

                  <xsl:when test="@showDataAs = 'index' ">
                    <xsl:text>index</xsl:text>
                  </xsl:when>

                </xsl:choose>
              </xsl:attribute>

              <xsl:for-each
                select="key('Part', concat('xl/',$cacheFile))/e:pivotCacheDefinition/e:cacheFields/e:cacheField[$base + 1]">

                <xsl:attribute name="table:field-name">
                  <xsl:value-of select="@name"/>
                </xsl:attribute>

                <xsl:for-each select="e:sharedItems">
                  <xsl:attribute name="table:member-name">
                    <xsl:value-of select="$baseCacheValue"/>
                  </xsl:attribute>
                </xsl:for-each>

              </xsl:for-each>

              <xsl:attribute name="table:member-type">
                <xsl:text>named</xsl:text>
              </xsl:attribute>

            </table:data-pilot-field-reference>
          </xsl:if>

        </xsl:if>

        <table:data-pilot-level>

          <xsl:attribute name="table:show-empty">
            <xsl:text>false</xsl:text>
          </xsl:attribute>


          <xsl:if test="e:items/child::node()/@h">

            <table:data-pilot-members>

              <xsl:for-each select="e:items/e:item[@x and @h='1']">

                <xsl:variable name="hiddenFieldNum">
                  <xsl:value-of select="@x"/>
                </xsl:variable>

                <xsl:variable name="hiddenCacheValue">
                  <xsl:for-each
                    select="key('Part', concat('xl/',$cacheFile))/e:pivotCacheDefinition/e:cacheFields/e:cacheField[position() = $fieldNum]/e:sharedItems">
                    <xsl:value-of select="child::node()[position() = $hiddenFieldNum + 1]/@v"/>
                  </xsl:for-each>
                </xsl:variable>

                <table:data-pilot-member>

                  <xsl:attribute name="table:name">
                    <xsl:value-of select="$hiddenCacheValue"/>
                  </xsl:attribute>

                  <xsl:attribute name="table:display">
                    <xsl:text>false</xsl:text>
                  </xsl:attribute>

                  <xsl:attribute name="table:show-details">
                    <xsl:text>true</xsl:text>
                  </xsl:attribute>

                </table:data-pilot-member>

              </xsl:for-each>

            </table:data-pilot-members>

          </xsl:if>

          <xsl:for-each
            select="parent::node()/parent::node()/e:pivotFields/e:pivotField[position() = $fieldNum + 1][@axis and @sumSubtotal or @countSubtotal or @avgSubtotal or @maxSubtotal or @minSubtotal or @productSubtotal or @countSubtotal or @stdDevSubtotal or @stdDevPSubtotal or @varSubtotal or @varPSubtotal]">

            <table:data-pilot-subtotals>

              <xsl:if test="@sumSubtotal">
                <table:data-pilot-subtotal>
                  <xsl:attribute name="table:function">
                    <xsl:text>sum</xsl:text>
                  </xsl:attribute>
                </table:data-pilot-subtotal>
              </xsl:if>

              <xsl:if test="@countASubtotal">
                <table:data-pilot-subtotal>
                  <xsl:attribute name="table:function">
                    <xsl:text>count</xsl:text>
                  </xsl:attribute>
                </table:data-pilot-subtotal>
              </xsl:if>

              <xsl:if test="@avgSubtotal">
                <table:data-pilot-subtotal>
                  <xsl:attribute name="table:function">
                    <xsl:text>average</xsl:text>
                  </xsl:attribute>
                </table:data-pilot-subtotal>
              </xsl:if>

              <xsl:if test="@maxSubtotal">
                <table:data-pilot-subtotal>
                  <xsl:attribute name="table:function">
                    <xsl:text>max</xsl:text>
                  </xsl:attribute>
                </table:data-pilot-subtotal>
              </xsl:if>

              <xsl:if test="@minSubtotal">
                <table:data-pilot-subtotal>
                  <xsl:attribute name="table:function">
                    <xsl:text>min</xsl:text>
                  </xsl:attribute>
                </table:data-pilot-subtotal>
              </xsl:if>

              <xsl:if test="@productSubtotal">
                <table:data-pilot-subtotal>
                  <xsl:attribute name="table:function">
                    <xsl:text>product</xsl:text>
                  </xsl:attribute>
                </table:data-pilot-subtotal>
              </xsl:if>

              <xsl:if test="@countSubtotal">
                <table:data-pilot-subtotal>
                  <xsl:attribute name="table:function">
                    <xsl:text>countnums</xsl:text>
                  </xsl:attribute>
                </table:data-pilot-subtotal>
              </xsl:if>

              <xsl:if test="@stdDevSubtotal">
                <table:data-pilot-subtotal>
                  <xsl:attribute name="table:function">
                    <xsl:text>stdev</xsl:text>
                  </xsl:attribute>
                </table:data-pilot-subtotal>
              </xsl:if>

              <xsl:if test="@stdDevPSubtotal">
                <table:data-pilot-subtotal>
                  <xsl:attribute name="table:function">
                    <xsl:text>stdevp</xsl:text>
                  </xsl:attribute>
                </table:data-pilot-subtotal>
              </xsl:if>

              <xsl:if test="@varSubtotal">
                <table:data-pilot-subtotal>
                  <xsl:attribute name="table:function">
                    <xsl:text>var</xsl:text>
                  </xsl:attribute>
                </table:data-pilot-subtotal>
              </xsl:if>

              <xsl:if test="@varPSubtotal">
                <table:data-pilot-subtotal>
                  <xsl:attribute name="table:function">
                    <xsl:text>varp</xsl:text>
                  </xsl:attribute>
                </table:data-pilot-subtotal>
              </xsl:if>

            </table:data-pilot-subtotals>

          </xsl:for-each>

          <table:data-pilot-display-info>

            <xsl:attribute name="table:enabled">
              <xsl:text>false</xsl:text>
            </xsl:attribute>

            <xsl:attribute name="table:display-member-mode">
              <xsl:text>from-top</xsl:text>
            </xsl:attribute>

          </table:data-pilot-display-info>

          <table:data-pilot-sort-info>

            <xsl:attribute name="table:order">
              <xsl:for-each
                select="parent::node()/parent::node()/e:pivotFields/e:pivotField[position() = $fieldNum + 1]">
                <xsl:choose>
                  <xsl:when test="@sortType">
                    <xsl:value-of select="@sortType"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:text>manual</xsl:text>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:for-each>
            </xsl:attribute>

            <xsl:attribute name="table:sort-mode">
              <xsl:for-each
                select="parent::node()/parent::node()/e:pivotFields/e:pivotField[position() = $fieldNum + 1]">
                <xsl:choose>
                  <xsl:when test="@sortType">
                    <xsl:text>name</xsl:text>
                  </xsl:when>
                  <xsl:when test="@axis">
                    <xsl:text>manual</xsl:text>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:text>none</xsl:text>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:for-each>
            </xsl:attribute>
          </table:data-pilot-sort-info>

          <table:data-pilot-layout-info>

            <xsl:attribute name="table:add-empty-lines">
              <xsl:choose>
                <xsl:when
                  test="parent::node()/parent::node()/e:pivotFields/e:pivotField/@insertBlankRow = '1' ">
                  <xsl:text>true</xsl:text>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:text>false</xsl:text>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>

            <xsl:attribute name="table:layout-mode">

              <xsl:choose>
                <xsl:when
                  test="parent::node()/parent::node()/e:pivotFields/e:pivotField/@outline = '0' ">
                  <xsl:text>tabular-layout</xsl:text>
                </xsl:when>

                <xsl:otherwise>
                  <xsl:choose>
                    <xsl:when
                      test="parent::node()/parent::node()/e:pivotFields/e:pivotField/@subtotalTop = '0' ">
                      <xsl:text>outline-subtotals-bottom</xsl:text>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:text>outline-subtotals-top</xsl:text>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:otherwise>

              </xsl:choose>

            </xsl:attribute>

          </table:data-pilot-layout-info>

        </table:data-pilot-level>

      </table:data-pilot-field>
    </xsl:if>

    <!--xsl:if test="@countSubtotal">
      <table:data-pilot-field>

        <xsl:attribute name="table:source-field-name"> </xsl:attribute>

        <xsl:attribute name="table:orientation"> </xsl:attribute>

        <xsl:attribute name="table:used-hierarchy">
          <xsl:text>-1</xsl:text>
        </xsl:attribute>

        <xsl:attribute name="table:function"> </xsl:attribute>

      </table:data-pilot-field>
    </xsl:if-->

  </xsl:template>

  <!-- put empty field after Row Label or Column Label-->
  <xsl:template name="InsertEmptyDataPilotField">
    <xsl:param name="orientation"/>

    <table:data-pilot-field table:source-field-name="">

      <xsl:attribute name="table:is-data-layout-field">
        <xsl:text>true</xsl:text>
      </xsl:attribute>

      <xsl:attribute name="table:orientation">

        <xsl:value-of select="$orientation"/>
        <!--xsl:choose>

          <xsl:when test="@axis = 'axisPage' ">
            <xsl:text>row</xsl:text>
          </xsl:when>

          <xsl:when test="@axis = 'axisRow' ">
            <xsl:text>row</xsl:text>
          </xsl:when>

          <xsl:when test="@axis = 'axisCol' ">
            <xsl:text>column</xsl:text>
          </xsl:when>

        </xsl:choose-->
      </xsl:attribute>

      <xsl:attribute name="table:used-hierarchy">
        <xsl:text>-1</xsl:text>
      </xsl:attribute>

      <xsl:attribute name="table:function">
        <xsl:text>auto</xsl:text>
      </xsl:attribute>

      <table:data-pilot-level table:show-empty="true"/>
    </table:data-pilot-field>

  </xsl:template>

  <xsl:template name="TranslateFilterOperator">

    <xsl:choose>
      <xsl:when test="substring(@val,1,1) = '*' or substring(@val,string-length(@val),1) = '*' ">
        <xsl:if test="@operator = 'notEqual' ">
          <xsl:text>!</xsl:text>
        </xsl:if>
        <xsl:text>match</xsl:text>
      </xsl:when>
      <xsl:when test="@operator = 'notEqual' ">
        <xsl:text>!=</xsl:text>
      </xsl:when>
      <xsl:when test="@operator = 'greaterThan' ">
        <xsl:text>&gt;</xsl:text>
      </xsl:when>
      <xsl:when test="@operator = 'greaterThanOrEqual' ">
        <xsl:text>&gt;=</xsl:text>
      </xsl:when>
      <xsl:when test="@operator = 'lessThan' ">
        <xsl:text>&lt;</xsl:text>
      </xsl:when>
      <xsl:when test="@operator = 'lessThanOrEqual' ">
        <xsl:text>&lt;=</xsl:text>
      </xsl:when>
      <xsl:when test="not(@operator)">
        <xsl:text>=</xsl:text>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="TranslateFilterValue">
    <xsl:choose>
      <!-- contains -->
      <xsl:when test="substring(@val,1,1) = '*' and substring(@val,string-length(@val),1) = '*' ">
        <xsl:text>.</xsl:text>
        <xsl:value-of select="substring(@val,1,string-length(@val)-1)"/>
        <xsl:text>.*</xsl:text>
      </xsl:when>
      <!-- begins with -->
      <xsl:when test="substring(@val,string-length(@val),1) = '*' ">
        <xsl:text>^</xsl:text>
        <xsl:value-of select="substring(@val,1,string-length(@val)-1)"/>
        <xsl:text>.*</xsl:text>
      </xsl:when>
      <!-- ends with -->
      <xsl:when test="substring(@val,1,1) = '*' ">
        <xsl:text>.</xsl:text>
        <xsl:value-of select="@val"/>
        <xsl:text>$</xsl:text>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="@val"/>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>
<!-- Vijayeta, 1803593,sales.xlsx,konto2006.xlsx-->
<xsl:template name ="oneMoreValidation">
	<xsl:param name ="cacheFile"/>
	<xsl:variable name ="validate">
		<xsl:for-each select="key('Part', concat('xl/',$cacheFile))/e:pivotCacheDefinition">
			<xsl:for-each select="e:cacheSource/e:worksheetSource">
				<xsl:variable name="apos">
					<xsl:text>&apos;</xsl:text>
				</xsl:variable>

				<xsl:variable name="sheetSourceName">
					<xsl:value-of select="@sheet"/>
				</xsl:variable>

				<xsl:variable name="firstSourceAdress">
					<xsl:value-of select="substring-before(@ref,':')"/>
				</xsl:variable>

				<xsl:variable name="lastSourceAdress">
					<xsl:value-of select="substring-after(@ref,':')"/>
				</xsl:variable>
				<xsl:if test ="$sheetSourceName != '' and $firstSourceAdress!='' and $lastSourceAdress!='' ">
					<xsl:value-of select ="'true'"/>
				</xsl:if>
			</xsl:for-each>
		</xsl:for-each >
	</xsl:variable>
	<xsl:if test ="$validate='true'">
		<xsl:value-of select ="'true'"/>
	</xsl:if>
	<xsl:if test ="$validate !='true'">
		<xsl:value-of select ="'false'"/>
	</xsl:if>
</xsl:template>
<!--End of Vijayeta, 1803593,sales.xlsx,konto2006.xlsx-->
</xsl:stylesheet>

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
    xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    exclude-result-prefixes="e r">

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
        <xsl:for-each select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet">

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
                test="document(concat('xl/worksheets/_rels/',$fileName,'.xml.rels'))//node()[name()='Relationship' and contains(@Type,'pivotTable' )]/@Target">
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

    <xsl:template name="InsertSheetPilotTables">
        <xsl:param name="sheetNum"/>

        <xsl:for-each
            select="document(concat('xl/worksheets/_rels/sheet',$sheetNum,'.xml.rels'))//node()[name()='Relationship' and contains(@Type,'pivotTable' )]">

            <xsl:variable name="TargetPilotFile">
                <xsl:value-of select="@Target"/>
            </xsl:variable>

            <xsl:variable name="sheetName">
                <xsl:for-each select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet">
                    <xsl:if test="position() = $sheetNum">
                        <xsl:value-of select="@name"/>
                    </xsl:if>
                </xsl:for-each>
            </xsl:variable>

            <xsl:if
                test="document(concat('xl/',substring-after($TargetPilotFile,'../')))/e:pivotTableDefinition">

                <xsl:for-each
                    select="document(concat('xl/',substring-after($TargetPilotFile,'../')))/e:pivotTableDefinition">

                    <xsl:variable name="name">
                        <xsl:value-of select="@name"/>
                    </xsl:variable>

                    <table:data-pilot-table table:name="{$name}">

                        <xsl:for-each select="e:location">

                            <xsl:variable name="firstTargetAdress">
                                <xsl:value-of select="substring-before(@ref,':')"/>
                            </xsl:variable>

                            <xsl:variable name="lastTargetAdress">
                                <xsl:value-of select="substring-after(@ref,':')"/>
                            </xsl:variable>


                            <xsl:attribute name="table:target-range-address">
                                <xsl:value-of
                                    select="concat($sheetName,'.',$firstTargetAdress,':',$sheetName,'.',$lastTargetAdress)"
                                />
                            </xsl:attribute>

                            <xsl:attribute name="table:buttons">

                                <xsl:variable name="numRow">
                                    <xsl:call-template name="GetRowNum">
                                        <xsl:with-param name="cell">
                                            <xsl:value-of select="$firstTargetAdress"/>
                                        </xsl:with-param>
                                    </xsl:call-template>
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

                                <xsl:value-of
                                    select="concat($sheetName,'.',$firstButtonCellCol,$numRow + 1,' ',$sheetName,'.',$nextButtonCellCol,$numRow)"/>

                            </xsl:attribute>

                            <xsl:attribute name="table:show-filter-button">
                                <xsl:text>false</xsl:text>
                            </xsl:attribute>

                        </xsl:for-each>

                        <xsl:for-each
                            select="document(concat('xl/pivotTables/_rels/pivotTable',$sheetNum,'.xml.rels'))//node()[name()='Relationship' and contains(@Type,'pivotCacheDefinition' )]">

                            <xsl:variable name="TargetPilotCacheFile">
                                <xsl:value-of select="@Target"/>
                            </xsl:variable>

                            <xsl:for-each
                                select="document(concat('xl/',substring-after($TargetPilotCacheFile,'../')))/e:pivotCacheDefinition">

                                <xsl:for-each select="e:cacheSource/e:worksheetSource">

                                    <xsl:variable name="firstSourceAdress">
                                        <xsl:value-of select="substring-before(@ref,':')"/>
                                    </xsl:variable>

                                    <xsl:variable name="lastSourceAdress">
                                        <xsl:value-of select="substring-after(@ref,':')"/>
                                    </xsl:variable>

                                    <table:source-cell-range>

                                        <xsl:attribute name="table:cell-range-address">
                                            <xsl:value-of
                                                select="concat($sheetName,'.',$firstSourceAdress,':',$sheetName,'.',$lastSourceAdress)"
                                            />
                                        </xsl:attribute>

                                    </table:source-cell-range>

                                    <table:data-pilot-field table:source-field-name="">

                                        <xsl:attribute name="table:is-data-layout-field">
                                            <xsl:text>true</xsl:text>
                                        </xsl:attribute>

                                        <xsl:attribute name="table:orientation">
                                            <xsl:text>column</xsl:text>
                                        </xsl:attribute>

                                        <xsl:attribute name="table:used-hierarchy">
                                            <xsl:text>-1</xsl:text>
                                        </xsl:attribute>

                                        <xsl:attribute name="table:function">
                                            <xsl:text>auto</xsl:text>
                                        </xsl:attribute>

                                        <table:data-pilot-level table:show-empty="true"/>
                                    </table:data-pilot-field>

                                </xsl:for-each>


                                <xsl:for-each
                                    select="document(concat('xl/',substring-after($TargetPilotFile,'../')))/e:pivotTableDefinition/e:pivotFields/e:pivotField[@axis or @dataField]">

                                    <xsl:variable name="fieldNum">
                                        <xsl:value-of select="position()"/>
                                    </xsl:variable>

                                    <table:data-pilot-field>

                                        <xsl:for-each
                                            select="document(concat('xl/',substring-after($TargetPilotCacheFile,'../')))/e:pivotCacheDefinition/e:cacheFields/e:cacheField[position() = $fieldNum]">

                                            <xsl:attribute name="table:source-field-name">

                                                <xsl:choose>
                                                  <xsl:when
                                                  test="number(translate(@name, ',' , '.' ))">

                                                  <xsl:variable name="replaceDecimal">
                                                  <xsl:value-of
                                                  select="format-number(translate(@name, ',' , '.' ),'0.##')"
                                                  />
                                                  </xsl:variable>

                                                  <xsl:value-of
                                                  select="translate($replaceDecimal, '.' , ',' )"
                                                  />
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

                                        <xsl:for-each
                                            select="document(concat('xl/',substring-after($TargetPilotFile,'../')))/e:pivotTableDefinition/e:pivotFields/e:pivotField[position() = $fieldNum]">

                                            <xsl:attribute name="table:orientation">
                                                <xsl:choose>
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
                                            </xsl:attribute>

                                            <xsl:attribute name="table:function">

                                                <xsl:choose>
                                                  <xsl:when
                                                  test="@axis = 'axisRow' or @axis = 'axisCol'">
                                                  <xsl:text>auto</xsl:text>
                                                  </xsl:when>

                                                  <xsl:otherwise>

                                                  <xsl:for-each
                                                  select="parent::node()/parent::node()/e:dataFields/e:dataField[@fld=$fieldNum - 1]">

                                                  <xsl:choose>
                                                  <xsl:when test="@subtotal = 'count'">
                                                  <xsl:text>count</xsl:text>
                                                  </xsl:when>

                                                  <xsl:when
                                                  test="@subtotal = 'average'">
                                                  <xsl:text>average</xsl:text>
                                                  </xsl:when>

                                                  <xsl:when test="@subtotal = 'max'">
                                                  <xsl:text>max</xsl:text>
                                                  </xsl:when>

                                                  <xsl:when test="@subtotal = 'min'">
                                                  <xsl:text>min</xsl:text>
                                                  </xsl:when>

                                                  <xsl:when
                                                  test="@subtotal = 'product'">
                                                  <xsl:text>product</xsl:text>
                                                  </xsl:when>

                                                  <xsl:when
                                                  test="@subtotal = 'countNums'">
                                                  <xsl:text>countnums</xsl:text>
                                                  </xsl:when>

                                                  <xsl:when
                                                  test="@subtotal = 'stdDev'">
                                                  <xsl:text>stddev</xsl:text>
                                                  </xsl:when>

                                                  <xsl:when
                                                  test="@subtotal = 'stdDevp'">
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

                                                  </xsl:for-each>

                                                  </xsl:otherwise>

                                                </xsl:choose>
                                            </xsl:attribute>

                                        </xsl:for-each>

                                        <table:data-pilot-level>

                                            <xsl:attribute name="table:show-empty">
                                                <xsl:text>false</xsl:text>
                                            </xsl:attribute>

                                            <table:data-pilot-display-info>

                                                <xsl:attribute name="table:enabled">
                                                  <xsl:text>false</xsl:text>
                                                </xsl:attribute>

                                                <xsl:attribute name="table:display-member-mode">
                                                  <xsl:text>from-top</xsl:text>
                                                </xsl:attribute>

                                                <!--xsl:attribute name="table:member-count">
                                                    
                                                </xsl:attribute>
                                            
                                                <xsl:attribute name="table:data-field">
                                                    
                                                </xsl:attribute-->

                                            </table:data-pilot-display-info>

                                            <table:data-pilot-sort-info>

                                                <xsl:if
                                                  test="parent::node()/e:pivotField[1]/@sortType">
                                                  <xsl:attribute name="table:order">
                                                      <xsl:value-of select="parent::node()/e:pivotField[1]/@sortType"/>
                                                  </xsl:attribute>
                                                </xsl:if>
                                                
                                                <xsl:attribute name="table:sort-mode">
                                                  <xsl:text>manual</xsl:text>
                                                </xsl:attribute>

                                            </table:data-pilot-sort-info>

                                            <table:data-pilot-layout-info>

                                                <xsl:attribute name="table:add-empty-lines">
                                                  <xsl:text>false</xsl:text>
                                                </xsl:attribute>

                                                <xsl:attribute name="table:layout-mode">
                                                  <xsl:text>tabular-layout</xsl:text>
                                                </xsl:attribute>

                                            </table:data-pilot-layout-info>

                                        </table:data-pilot-level>

                                    </table:data-pilot-field>

                                </xsl:for-each>

                            </xsl:for-each>
                        </xsl:for-each>

                    </table:data-pilot-table>
                </xsl:for-each>
            </xsl:if>
        </xsl:for-each>
    </xsl:template>

</xsl:stylesheet>

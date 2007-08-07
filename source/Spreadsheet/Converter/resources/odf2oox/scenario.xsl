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
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
    xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:x="urn:schemas-microsoft-com:office:excel"
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" exclude-result-prefixes="table">

    <xsl:import href="common.xsl"/>
    <xsl:import href="cell.xsl"/>



    <xsl:template name="InsertScenario">
        <xsl:if test="following-sibling::table:table[1]/table:scenario">

            <scenarios>

                <xsl:apply-templates select="following-sibling::table:table[1]" mode="scenario"/>

                <!--xsl:call-template name="SearchScenarioCells">
                    <xsl:with-param name="colNum"/>
                    <xsl:with-param name="rows"/>
                </xsl:call-template-->

            </scenarios>

        </xsl:if>
    </xsl:template>


    <xsl:template match="table:table" mode="scenario">
        <xsl:param name="count" select="1"/>

        <scenario>

            <xsl:for-each select="table:scenario">
                <xsl:attribute name="name">
                    <xsl:for-each select="ancestor::table:table[@table:name !='']">
                        <xsl:value-of select="@table:name"/>
                    </xsl:for-each>
                </xsl:attribute>
            </xsl:for-each>

            <xsl:variable name="ScenarioRanges">
                <xsl:value-of select="table:scenario/@table:scenario-ranges"/>
            </xsl:variable>


            <xsl:attribute name="comment">
                <xsl:if test="table:scenario/@table:comment">
                    <xsl:value-of select="table:scenario/@table:comment"/>
                </xsl:if>
            </xsl:attribute>

            <xsl:if test="table:scenario/@table:protected">
                <xsl:attribute name="locked">
                    <xsl:text>1</xsl:text>
                </xsl:attribute>
            </xsl:if>

            <xsl:variable name="row1">
                <xsl:call-template name="GetRowNum">
                    <xsl:with-param name="cell"
                        select="substring-after(substring-before($ScenarioRanges,':'),'.')"/>
                </xsl:call-template>
            </xsl:variable>

            <xsl:variable name="row2">
                <xsl:call-template name="GetRowNum">
                    <xsl:with-param name="cell"
                        select="substring-after(substring-after($ScenarioRanges,':'),'.')"/>
                </xsl:call-template>
            </xsl:variable>

            <xsl:variable name="col1">
                <!-- substring-before rowNum in cell coordinates-->
                <xsl:value-of
                    select="substring-before(substring-after(substring-before($ScenarioRanges,':'),'.'),$row1)"
                />
            </xsl:variable>

            <xsl:variable name="col2">
                <xsl:value-of
                    select="substring-before(substring-after(substring-after($ScenarioRanges,':'),'.'),$row2)"
                />
            </xsl:variable>
            <xsl:variable name="startCol">
                <xsl:choose>
                    <xsl:when test="$ScenarioRanges">
                        <!-- A equals 0 -->
                        <xsl:call-template name="GetAlphabeticPosition">
                            <xsl:with-param name="literal" select="$col1"/>
                        </xsl:call-template>
                    </xsl:when>
                </xsl:choose>
            </xsl:variable>

            <xsl:variable name="endCol">
                <xsl:choose>
                    <xsl:when test="$ScenarioRanges">
                        <!-- A equals 0 -->
                        <xsl:call-template name="GetAlphabeticPosition">
                            <xsl:with-param name="literal" select="$col2"/>
                        </xsl:call-template>
                    </xsl:when>
                </xsl:choose>
            </xsl:variable>

            <xsl:if test="@table:scenario-ranges"/>

            <xsl:apply-templates select="table:table-row[1]" mode="scenario">
                <xsl:with-param name="rowNumber">
                    <xsl:text>1</xsl:text>
                </xsl:with-param>
                <xsl:with-param name="colNumber">
                    <xsl:text>1</xsl:text>
                </xsl:with-param>
                <xsl:with-param name="row1">
                    <xsl:value-of select="$row1"/>
                </xsl:with-param>
                <xsl:with-param name="row2">
                    <xsl:value-of select="$row2"/>
                </xsl:with-param>
                <xsl:with-param name="startCol">
                    <xsl:value-of select="$startCol"/>
                </xsl:with-param>
                <xsl:with-param name="endCol">
                    <xsl:value-of select="$endCol"/>
                </xsl:with-param>
            </xsl:apply-templates>

        </scenario>

        <xsl:if test="following-sibling::table:table[1]/table:scenario">
            <xsl:choose>
                <xsl:when test="$count &lt; 31">
                    <xsl:apply-templates select="following-sibling::table:table[1]" mode="scenario">
                        <xsl:with-param name="count">
                            <xsl:value-of select="$count + 1"/>
                        </xsl:with-param>
                    </xsl:apply-templates>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:message terminate="no">translation.odf2oox.ScenarioNumber</xsl:message>
                </xsl:otherwise>
            </xsl:choose>


        </xsl:if>

    </xsl:template>

    <xsl:template name="SearchScenarioCells">
        <xsl:for-each select="/office:document-content/office:body/office:spreadsheet/table:table">

            <!--xsl:call-template name="InsertScenario">
                <xsl:with-param name="startCol"/>
                <xsl:with-param name="endCol"/>
            </xsl:call-template-->

            <xsl:for-each select="table:table-row/table:table-cell/text:p">
                <xsl:variable name="colPosition">
                    <xsl:for-each select="ancestor::table:table-cell">
                        <xsl:value-of
                            select="count(preceding-sibling::table:table-cell) + count(preceding-sibling::table:covered-table-cell) + 1"
                        />
                    </xsl:for-each>
                </xsl:variable>

                <xsl:variable name="rowPosition">
                    <xsl:value-of select="generate-id(ancestor::table:table-row)"/>
                </xsl:variable>

                <!-- real column number -->
                <xsl:variable name="colNum">
                    <xsl:for-each select="ancestor::table:table-row/table:table-cell[1]">
                        <xsl:call-template name="GetColNumber">
                            <xsl:with-param name="position">
                                <xsl:value-of select="$colPosition"/>
                            </xsl:with-param>
                        </xsl:call-template>
                    </xsl:for-each>
                </xsl:variable>

                <xsl:variable name="rows">
                    <xsl:for-each select="ancestor::table:table/descendant::table:table-row[1]">
                        <xsl:call-template name="GetRowNumber">
                            <xsl:with-param name="rowId" select="$rowPosition"/>
                            <xsl:with-param name="tableId"
                                select="generate-id(ancestor::table:table)"/>
                        </xsl:call-template>
                    </xsl:for-each>
                </xsl:variable>

                <!--xsl:variable name="colChar">
                    <xsl:call-template name="NumbersToChars">
                        <xsl:with-param name="num" select="$colNum -1"/>
                    </xsl:call-template>
                    </xsl:variable-->

                <!--inputCells r="{concat($colNum,':',$rows)}"/-->
            </xsl:for-each>
        </xsl:for-each>
    </xsl:template>

    <!-- search scenario -->
    <xsl:template match="table:table-row" mode="scenario">
        <xsl:param name="rowNumber"/>
        <xsl:param name="colNumber"/>
        <xsl:param name="row1"/>
        <xsl:param name="row2"/>
        <xsl:param name="startCol"/>
        <xsl:param name="endCol"/>


        <!-- Insert row in scenario -->
        <xsl:if test="$rowNumber &gt;= $row1 and $rowNumber &lt;=$row2">

            <xsl:call-template name="InsertRowsScenario">
                <xsl:with-param name="rowNumber">
                    <xsl:value-of select="$rowNumber"/>
                </xsl:with-param>
                <xsl:with-param name="Repeat">
                    <xsl:choose>
                        <xsl:when test="@table:number-rows-repeated">
                            <xsl:value-of select="@table:number-rows-repeated"/>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:text>1</xsl:text>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:with-param>
                <xsl:with-param name="RepeatNumber">
                    <xsl:text>1</xsl:text>
                </xsl:with-param>
                <xsl:with-param name="row1">
                    <xsl:value-of select="$row1"/>
                </xsl:with-param>
                <xsl:with-param name="row2">
                    <xsl:value-of select="$row2"/>
                </xsl:with-param>
                <xsl:with-param name="startCol">
                    <xsl:value-of select="$startCol"/>
                </xsl:with-param>
                <xsl:with-param name="endCol">
                    <xsl:value-of select="$endCol"/>
                </xsl:with-param>
            </xsl:call-template>


        </xsl:if>

        <!-- check next row -->
        <xsl:choose>
            <!-- next row is a sibling -->
            <xsl:when test="following-sibling::node()[1][name() = 'table:table-row' ]">
                <xsl:apply-templates select="following-sibling::table:table-row[1]" mode="scenario">
                    <xsl:with-param name="rowNumber">
                        <xsl:choose>
                            <xsl:when test="@table:number-rows-repeated">
                                <xsl:value-of select="$rowNumber+@table:number-rows-repeated"/>
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:value-of select="$rowNumber+1"/>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:with-param>
                    <xsl:with-param name="cellNumber">
                        <xsl:text>0</xsl:text>
                    </xsl:with-param>
                    <xsl:with-param name="row1">
                        <xsl:value-of select="$row1"/>
                    </xsl:with-param>
                    <xsl:with-param name="row2">
                        <xsl:value-of select="$row2"/>
                    </xsl:with-param>
                    <xsl:with-param name="startCol">
                        <xsl:value-of select="$startCol"/>
                    </xsl:with-param>
                    <xsl:with-param name="endCol">
                        <xsl:value-of select="$endCol"/>
                    </xsl:with-param>
                </xsl:apply-templates>
            </xsl:when>
        </xsl:choose>
    </xsl:template>

    <xsl:template name="InsertRowsScenario">
        <xsl:param name="rowNumber"/>
        <xsl:param name="Repeat"/>
        <xsl:param name="RepeatNumber"/>
        <xsl:param name="row1"/>
        <xsl:param name="row2"/>
        <xsl:param name="startCol"/>
        <xsl:param name="endCol"/>

        <xsl:apply-templates select="child::node()[name() = 'table:table-cell'][1]" mode="scenario">
            <xsl:with-param name="colNumber">
                <xsl:text>1</xsl:text>
            </xsl:with-param>
            <xsl:with-param name="rowNumber" select="$rowNumber"/>
            <xsl:with-param name="row1">
                <xsl:value-of select="$row1"/>
            </xsl:with-param>
            <xsl:with-param name="row2">
                <xsl:value-of select="$row2"/>
            </xsl:with-param>
            <xsl:with-param name="startCol">
                <xsl:value-of select="$startCol"/>
            </xsl:with-param>
            <xsl:with-param name="endCol">
                <xsl:value-of select="$endCol"/>
            </xsl:with-param>
        </xsl:apply-templates>

        <xsl:if test="$RepeatNumber &lt; $Repeat">
            <xsl:call-template name="InsertRowsScenario">
                <xsl:with-param name="rowNumber">
                    <xsl:value-of select="$rowNumber + 1"/>
                </xsl:with-param>
                <xsl:with-param name="RepeatNumber">
                    <xsl:value-of select="$RepeatNumber + 1"/>
                </xsl:with-param>
                <xsl:with-param name="Repeat">
                    <xsl:value-of select="$Repeat"/>
                </xsl:with-param>
                <xsl:with-param name="startCol">
                    <xsl:value-of select="$startCol"/>
                </xsl:with-param>
                <xsl:with-param name="endCol">
                    <xsl:value-of select="$endCol"/>
                </xsl:with-param>
            </xsl:call-template>
        </xsl:if>


    </xsl:template>

    <!-- insert scenario -->
    <xsl:template match="table:table-cell" mode="scenario">
        <xsl:param name="colNumber"/>
        <xsl:param name="rowNumber"/>
        <xsl:param name="row1"/>
        <xsl:param name="row2"/>
        <xsl:param name="startCol"/>
        <xsl:param name="endCol"/>


        <xsl:variable name="colChar">
            <xsl:call-template name="NumbersToChars">
                <xsl:with-param name="num" select="$colNumber - 1 "/>
            </xsl:call-template>
        </xsl:variable>

        <xsl:choose>

            <xsl:when test="text:p">
                <inputCells>
                    <xsl:attribute name="r">
                        <xsl:value-of select="concat($colChar, $rowNumber)"/>
                    </xsl:attribute>
                    <xsl:attribute name="val">
                        <xsl:value-of select="text:p"/>
                    </xsl:attribute>
                </inputCells>
            </xsl:when>

            <xsl:when
                test="$colNumber &lt; $startCol and $colNumber +  @table:number-columns-repeated &gt; $endCol">
                <xsl:call-template name="InsertInputCels">
                    <xsl:with-param name="start">
                        <xsl:value-of select="$startCol"/>
                    </xsl:with-param>
                    <xsl:with-param name="end">
                        <xsl:value-of select="$endCol"/>
                    </xsl:with-param>
                    <xsl:with-param name="rowNumber">
                        <xsl:value-of select="$rowNumber"/>
                    </xsl:with-param>
                </xsl:call-template>
            </xsl:when>

            <xsl:when
                test="$colNumber &lt; $startCol and $startCol &lt;= $colNumber + @table:number-columns-repeated and $colNumber + @table:number-columns-repeated &lt; $endCol">
                <xsl:call-template name="InsertInputCels">
                    <xsl:with-param name="start">
                        <xsl:value-of select="$startCol"/>
                    </xsl:with-param>
                    <xsl:with-param name="end">
                        <xsl:value-of select="$colNumber + @table:number-columns-repeated - 1"/>
                    </xsl:with-param>
                    <xsl:with-param name="rowNumber">
                        <xsl:value-of select="$rowNumber"/>
                    </xsl:with-param>
                </xsl:call-template>
            </xsl:when>

            <xsl:when
                test="$startCol &lt; $colNumber and $endCol &gt;= $colNumber and $colNumber + @table:number-columns-repeated &gt; $endCol">

                <xsl:call-template name="InsertInputCels">
                    <xsl:with-param name="start">
                        <xsl:value-of select="$colNumber"/>
                    </xsl:with-param>
                    <xsl:with-param name="end">
                        <xsl:value-of select="$endCol"/>
                    </xsl:with-param>
                    <xsl:with-param name="rowNumber">
                        <xsl:value-of select="$rowNumber"/>
                    </xsl:with-param>
                </xsl:call-template>
            </xsl:when>

            <xsl:when
                test="$colNumber &gt;= $startCol and $colNumber +  @table:number-columns-repeated &lt;= $endCol">

                <xsl:call-template name="InsertInputCels">
                    <xsl:with-param name="start">
                        <xsl:value-of select="$colNumber"/>
                    </xsl:with-param>
                    <xsl:with-param name="end">
                        <xsl:value-of select="$colNumber +  @table:number-columns-repeated - 1"/>
                    </xsl:with-param>
                    <xsl:with-param name="rowNumber">
                        <xsl:value-of select="$rowNumber"/>
                    </xsl:with-param>
                </xsl:call-template>

            </xsl:when>
        </xsl:choose>


        <xsl:choose>
            <xsl:when test="following-sibling::table:table-cell">

                <xsl:apply-templates
                    select="following-sibling::node()[name() = 'table:table-cell' or name() = 'table:covered-table-cell'][1]"
                    mode="scenario">
                    <xsl:with-param name="colNumber">
                        <xsl:choose>
                            <xsl:when test="@table:number-columns-repeated != ''">
                                <xsl:value-of
                                    select="number($colNumber) + number(@table:number-columns-repeated)"
                                />
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:value-of select="$colNumber + 1"/>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:with-param>
                    <xsl:with-param name="rowNumber">
                        <xsl:value-of select="$rowNumber"/>
                    </xsl:with-param>
                    <xsl:with-param name="row1">
                        <xsl:value-of select="$row1"/>
                    </xsl:with-param>
                    <xsl:with-param name="row2">
                        <xsl:value-of select="$row2"/>
                    </xsl:with-param>
                    <xsl:with-param name="startCol">
                        <xsl:value-of select="$startCol"/>
                    </xsl:with-param>
                    <xsl:with-param name="endCol">
                        <xsl:value-of select="$endCol"/>
                    </xsl:with-param>
                </xsl:apply-templates>

            </xsl:when>

        </xsl:choose>

    </xsl:template>

    <xsl:template name="InsertInputCels">
        <xsl:param name="start"/>
        <xsl:param name="end"/>
        <xsl:param name="rowNumber"/>

        <xsl:variable name="colChar">
            <xsl:call-template name="NumbersToChars">
                <xsl:with-param name="num" select="$start -1"/>
            </xsl:call-template>
        </xsl:variable>

        <xsl:if test="$start &lt;= $end">
            <inputCells>
                <xsl:attribute name="r">
                    <xsl:value-of select="concat($colChar, $rowNumber)"/>
                </xsl:attribute>
                <xsl:attribute name="val">
                    <xsl:value-of select="text:p"/>
                </xsl:attribute>
            </inputCells>

            <xsl:call-template name="InsertInputCels">
                <xsl:with-param name="start">
                    <xsl:value-of select="$start + 1"/>
                </xsl:with-param>
                <xsl:with-param name="end">
                    <xsl:value-of select="$end"/>
                </xsl:with-param>
                <xsl:with-param name="rowNumber">
                    <xsl:value-of select="$rowNumber"/>
                </xsl:with-param>
            </xsl:call-template>

        </xsl:if>

    </xsl:template>

</xsl:stylesheet>

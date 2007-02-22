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
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
    exclude-result-prefixes="svg">

    <xsl:import href="measures.xsl"/>
    <xsl:key name="font" match="style:font-face" use="@style:name"/>

    <xsl:template name="styles">
        <styleSheet>
            <xsl:call-template name="InsertFonts"/>
            <xsl:call-template name="InsertFills"/>
            <xsl:call-template name="InsertBorders"/>
            <xsl:call-template name="InsertFormatingRecords"/>
            <xsl:call-template name="InsertCellFormats"/>
            <xsl:call-template name="InsertCellStyles"/>
            <xsl:call-template name="InsertFormats"/>
            <xsl:call-template name="InsertTableStyles"/>
        </styleSheet>
    </xsl:template>

    <xsl:template name="InsertFonts">
        <fonts>
            <xsl:attribute name="count">
                <xsl:value-of
                    select="count(document('content.xml')/office:document-content/office:automatic-styles/style:style[@style:family='table-cell']/style:text-properties) + 1"
                />
            </xsl:attribute>

            <!-- default font-->
            <xsl:choose>
                <xsl:when
                    test="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name='Default' and @style:family='table-cell']/style:text-properties">
                    <xsl:apply-templates
                        select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name='Default' and @style:family='table-cell']/style:text-properties"
                        mode="fonts"/>
                </xsl:when>
                <!-- application default-->
                <xsl:otherwise>
                    <font>
                        <sz val="10"/>
                        <name val="Arial"/>
                    </font>
                </xsl:otherwise>
            </xsl:choose>

            <xsl:apply-templates
                select="document('content.xml')/office:document-content/office:automatic-styles"
                mode="fonts"/>
        </fonts>
    </xsl:template>

    <xsl:template name="InsertFills">
        <fills count="2">
            <fill>
                <patternFill patternType="none"/>
            </fill>
            <fill>
                <patternFill patternType="gray125"/>
            </fill>
        </fills>
    </xsl:template>

    <xsl:template name="InsertBorders">
        <borders count="1">
            <border>
                <left/>
                <right/>
                <top/>
                <bottom/>
                <diagonal/>
            </border>
        </borders>
    </xsl:template>

    <xsl:template name="InsertFormatingRecords">
        <cellStyleXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        </cellStyleXfs>
    </xsl:template>

    <xsl:template name="InsertCellFormats">
        <cellXfs>
            <xsl:attribute name="count">
                <xsl:value-of
                    select="count(document('content.xml')/office:document-content/office:automatic-styles/style:style[@style:family='table-cell']/style:text-properties) + 1"
                />
            </xsl:attribute>

            <!-- default style -->
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>

            <xsl:apply-templates
                select="document('content.xml')/office:document-content/office:automatic-styles"
                mode="applyFont"/>
        </cellXfs>
    </xsl:template>

    <xsl:template name="InsertCellStyles">
        <cellStyles count="1">
            <cellStyle name="Normal" xfId="0" builtinId="0"/>
        </cellStyles>
    </xsl:template>

    <xsl:template name="InsertFormats">
        <dxfs count="0"/>
    </xsl:template>

    <xsl:template name="InsertTableStyles">
        <tableStyles count="0" defaultTableStyle="TableStyleMedium9"
            defaultPivotStyle="PivotStyleLight16"/>
    </xsl:template>

    <xsl:template match="style:text-properties[parent::node()[@style:family='table-cell' or @style:family='text']]"
        mode="fonts">
        <font>
            <xsl:call-template name="InsertTextProperties">
                <xsl:with-param name="mode">fonts</xsl:with-param>
            </xsl:call-template>
        </font>
    </xsl:template>

    <xsl:template name="InsertUnderline">
        <xsl:if test="@style:text-underline-style != 'none' ">
            <u>
                <xsl:attribute name="val">
                    <xsl:choose>
                        <xsl:when test="@style:text-underline-style = 'accounting' ">
                            <xsl:choose>
                                <xsl:when test="@style:text-underline-type = 'double' "
                                    >doubleAccounting</xsl:when>
                                <xsl:otherwise>singleAccounting</xsl:otherwise>
                            </xsl:choose>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:choose>
                                <xsl:when test="@style:text-underline-type = 'double' ">double</xsl:when>
                                <xsl:otherwise>single</xsl:otherwise>
                            </xsl:choose>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:attribute>
            </u>
        </xsl:if>
    </xsl:template>

    <xsl:template match="style:text-properties[parent::node()/@style:family='table-cell']"
        mode="applyFont">
        <xf numFmtId="0" fillId="0" borderId="0" xfId="0" applyFont="1">
            <xsl:attribute name="fontId">
                <xsl:number count="style:text-properties[parent::node()/@style:family='table-cell']"
                    level="any"/>
            </xsl:attribute>
        </xf>
    </xsl:template>

    <!-- insert run properties -->
    <xsl:template match="style:style" mode="textstyles">
        <xsl:if test="style:text-properties">
            <rPr>
                <xsl:apply-templates select="style:text-properties" mode="textstyles"/>
            </rPr>
        </xsl:if>
    </xsl:template>

    <!-- convert text properties -->
    <xsl:template match="style:text-properties" mode="textstyles">
        <xsl:call-template name="InsertTextProperties">
            <xsl:with-param name="mode">textstyles</xsl:with-param>
        </xsl:call-template>
    </xsl:template>

    <!-- insert text properties -->
    <xsl:template name="InsertTextProperties">
        <xsl:param name="mode"/>
        <xsl:if test="@fo:font-weight='bold'">
            <b/>
        </xsl:if>
        <xsl:if test="@fo:font-style='italic'">
            <i/>
        </xsl:if>
        <xsl:if test="@style:text-underline-style">
            <xsl:call-template name="InsertUnderline"/>
        </xsl:if>
        <xsl:if test="@fo:font-size">
            <sz>
                <xsl:attribute name="val">
                    <xsl:call-template name="point-measure">
                        <xsl:with-param name="length">
                            <xsl:value-of select="@fo:font-size"/>
                        </xsl:with-param>
                    </xsl:call-template>
                </xsl:attribute>
            </sz>
        </xsl:if>
        <xsl:if test="@style:text-line-through-style and @style:text-line-through-style != 'none' ">
            <strike/>
        </xsl:if>
        <xsl:if test="@fo:color">
            <color rgb="{concat('FF',substring-after(@fo:color,'#'))}"/>
        </xsl:if>
        <xsl:choose>
            <xsl:when test="$mode = 'textstyles'">
                <xsl:if test="@style:font-name">
                    <rFont val="{@style:font-name}"/>
                </xsl:if>
            </xsl:when>
            <xsl:when test="$mode = 'fonts'">
              <xsl:if test="key('font',@style:font-name)/@svg:font-family">
                <name>
                    <xsl:attribute name="val">
                        <xsl:choose>
                            <xsl:when
                                test="not(translate(key('font',@style:font-name)/@svg:font-family,&quot;&apos;&quot;,&quot;&quot;) = '' )">
                                <xsl:value-of
                                    select="translate(key('font',@style:font-name)/@svg:font-family,&quot;&apos;&quot;,&quot;&quot;)"
                                />
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:value-of select="key('font',@style:font-name)/@svg:font-family"
                                />
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:attribute>
                </name>
              </xsl:if>
            </xsl:when>
        </xsl:choose>
    </xsl:template>

</xsl:stylesheet>

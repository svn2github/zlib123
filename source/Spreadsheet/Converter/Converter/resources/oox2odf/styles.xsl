<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    exclude-result-prefixes="e">

    <xsl:import href="relationships.xsl"/>
    
    <xsl:template name="InsertColumnStyles">
        <xsl:for-each select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet">
            <xsl:call-template name="InsertSheetColumnStyles">
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
    </xsl:template>
    
    <xsl:template name="InsertSheetColumnStyles">
        <xsl:param name="sheet"/>
        
        <!-- default style -->
        <style:style
            style:name="{generate-id(document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr)}"
            style:family="table-column">
            <style:table-column-properties fo:break-before="auto">
                <xsl:attribute name="style:column-width">
                    <xsl:choose>
                        <xsl:when
                            test="document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr/@defaultColWidth">
                            <xsl:call-template name="ConvertFromCharacters">
                                <xsl:with-param name="value">
                                    <xsl:value-of
                                        select="document(concat('xl/',$sheet))/e:worksheet/e:sheetFormatPr/@defaultColWidth"
                                    />
                                </xsl:with-param>
                            </xsl:call-template>
                        </xsl:when>
                        <xsl:otherwise>
                            <!-- Excel application default-->
                            <xsl:call-template name="ConvertToCentimeters">
                                <xsl:with-param name="length" select="'64px'"/>
                            </xsl:call-template>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:attribute>
            </style:table-column-properties>
        </style:style>
        
        <xsl:apply-templates select="document(concat('xl/',$sheet))/e:worksheet/e:cols"
            mode="automaticstyles"/>
    </xsl:template>
    
    <xsl:template match="e:col" mode="automaticstyles">
        <style:style style:name="{generate-id(.)}" style:family="table-column">
            <style:table-column-properties>
                <xsl:if test="@width">
                    <xsl:attribute name="style:column-width">
                        <xsl:call-template name="ConvertFromCharacters">
                            <xsl:with-param name="value" select="@width"/>
                        </xsl:call-template>
                    </xsl:attribute>
                </xsl:if>
            </style:table-column-properties>
        </style:style>
    </xsl:template>
    
    <xsl:template name="InsertRowStyles">
        <xsl:for-each select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet">
            <xsl:call-template name="InsertSheetRowStyles">
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
    </xsl:template>
    
    <xsl:template name="InsertSheetRowStyles">
        <xsl:param name="sheet"/>
        
    </xsl:template>
    
    <!--  Insert Table Properties -->
    <xsl:template name="InsertStyleTableProperties">
        <xsl:for-each select="document('xl/workbook.xml')/e:workbook/e:sheets/e:sheet">
            <style:style>
                <xsl:attribute name="style:name">
                    <xsl:value-of select="generate-id()"/>
                </xsl:attribute>
                <xsl:attribute name="style:family">
                    <xsl:text>table</xsl:text>
                </xsl:attribute>
                <style:table-properties>
                    <xsl:if test="@state='hidden'">
                        <xsl:attribute name="table:display">
                            <xsl:text>false</xsl:text>
                        </xsl:attribute>
                    </xsl:if>
                </style:table-properties>
            </style:style>
        </xsl:for-each>
    </xsl:template>
    
    <xsl:template name="InsertCellStyles">
        <xsl:apply-templates select="document('xl/styles.xml')/e:styleSheet/e:cellXfs"
            mode="automaticstyles"/>
    </xsl:template>
    
    <xsl:template match="e:xf" mode="automaticstyles">
        <style:style style:name="{generate-id(.)}" style:family="table-cell">
            <style:text-properties>
                <xsl:variable name="this" select="."/>
                <xsl:apply-templates
                    select="ancestor::e:styleSheet/e:fonts/e:font[position() = $this/@fontId + 1]"
                    mode="style"/>
            </style:text-properties>
        </style:style>
    </xsl:template>
    
    <xsl:template match="e:b" mode="style">
        <xsl:attribute name="fo:font-weight">
            <xsl:text>bold</xsl:text>
        </xsl:attribute>
    </xsl:template>

    <xsl:template match="e:i" mode="style">
        <xsl:attribute name="fo:font-style">
            <xsl:text>italic</xsl:text>
        </xsl:attribute>
    </xsl:template>

    <xsl:template match="e:u" mode="style">
        <xsl:call-template name="InsertUnderline"/>
    </xsl:template>

    <xsl:template name="InsertUnderline">
        <xsl:choose>
            <xsl:when test="@e:val = 'dash'">
                <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
                <xsl:attribute name="style:text-underline-style">dash</xsl:attribute>
                <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
            </xsl:when>
            <xsl:when test="@e:val = 'dashDotDotHeavy'">
                <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
                <xsl:attribute name="style:text-underline-style">dot-dot-dash</xsl:attribute>
                <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
            </xsl:when>
            <xsl:when test="@e:val = 'dashDotHeavy'">
                <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
                <xsl:attribute name="style:text-underline-style">dot-dash</xsl:attribute>
                <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
            </xsl:when>
            <xsl:when test="@e:val = 'dashedHeavy'">
                <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
                <xsl:attribute name="style:text-underline-style">dash</xsl:attribute>
                <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
            </xsl:when>
            <xsl:when test="@e:val = 'dashLong'">
                <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
                <xsl:attribute name="style:text-underline-style">long-dash</xsl:attribute>
                <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
            </xsl:when>
            <xsl:when test="@e:val = 'dashLongHeavy'">
                <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
                <xsl:attribute name="style:text-underline-style">long-dash</xsl:attribute>
                <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
            </xsl:when>
            <xsl:when test="@e:val = 'dotDash'">
                <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
                <xsl:attribute name="style:text-underline-style">dot-dash</xsl:attribute>
                <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
            </xsl:when>
            <xsl:when test="@e:val = 'dotDotDash'">
                <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
                <xsl:attribute name="style:text-underline-style">dot-dot-dash</xsl:attribute>
                <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
            </xsl:when>
            <xsl:when test="@e:val = 'dotted'">
                <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
                <xsl:attribute name="style:text-underline-style">dotted</xsl:attribute>
                <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
            </xsl:when>
            <xsl:when test="@e:val = 'dottedHeavy'">
                <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
                <xsl:attribute name="style:text-underline-style">dotted</xsl:attribute>
                <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
            </xsl:when>
            <xsl:when test="@e:val = 'double'">
                <xsl:attribute name="style:text-underline-type">double</xsl:attribute>
                <xsl:attribute name="style:text-underline-style">solid</xsl:attribute>
                <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
            </xsl:when>
            <xsl:when test="@e:val = 'single'">
                <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
                <xsl:attribute name="style:text-underline-style">solid</xsl:attribute>
                <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
            </xsl:when>
            <xsl:when test="@e:val = 'thick'">
                <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
                <xsl:attribute name="style:text-underline-style">solid</xsl:attribute>
                <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
            </xsl:when>
            <xsl:when test="@e:val = 'wave'">
                <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
                <xsl:attribute name="style:text-underline-style">wave</xsl:attribute>
                <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
            </xsl:when>
            <xsl:when test="@e:val = 'wavyDouble'">
                <xsl:attribute name="style:text-underline-type">double</xsl:attribute>
                <xsl:attribute name="style:text-underline-style">wave</xsl:attribute>
                <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
            </xsl:when>
            <xsl:when test="@e:val = 'wavyHeavy'">
                <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
                <xsl:attribute name="style:text-underline-style">wave</xsl:attribute>
                <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
            </xsl:when>
            <xsl:when test="@e:val = 'words'">
                <xsl:attribute name="style:text-underline-mode">skip-white-space</xsl:attribute>
                <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
                <xsl:attribute name="style:text-underline-style">solid</xsl:attribute>
                <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
            </xsl:when>
            <xsl:when test="@e:val = 'singleAccounting'">
                <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
                <xsl:attribute name="style:text-underline-style">accounting</xsl:attribute>
                <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
            </xsl:when>
            <xsl:when test="@e:val = 'doubleAccounting'">
                <xsl:attribute name="style:text-underline-type">double</xsl:attribute>
                <xsl:attribute name="style:text-underline-style">accounting</xsl:attribute>
                <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
            </xsl:when>
            <xsl:when test="@e:val = 'none'">
                <xsl:attribute name="style:text-underline-type">none</xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
                <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
                <xsl:attribute name="style:text-underline-style">solid</xsl:attribute>
                <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
            </xsl:otherwise>
        </xsl:choose>
        <xsl:if test="@e:color">
            <xsl:attribute name="style:text-underline-color">
                <xsl:choose>
                    <xsl:when test="@e:color = 'auto'">font-color</xsl:when>
                    <xsl:otherwise>
                        <xsl:value-of select="concat('#',@e:color)"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:attribute>
        </xsl:if>
    </xsl:template>
    
    <xsl:template match="e:sz" mode="style">
        <!-- do not insert this property into drop cap text style -->
        <xsl:attribute name="fo:font-size">
            <xsl:value-of select="@val"/>
        </xsl:attribute>
    </xsl:template>
    
    <xsl:template match="e:name" mode="style">
        <xsl:attribute name="style:font-name">
            <xsl:value-of select="@val"/>
        </xsl:attribute>
    </xsl:template>
    
</xsl:stylesheet>

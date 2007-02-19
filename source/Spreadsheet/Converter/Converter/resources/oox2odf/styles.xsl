<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main" exclude-result-prefixes="e">

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

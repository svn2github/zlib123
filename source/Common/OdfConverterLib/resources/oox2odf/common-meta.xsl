<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:cust-p="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/" 
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
    xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
    exclude-result-prefixes="vt cust-p cp">

    <xsl:template name="meta">
        <xsl:param name="app-version"/>
        <office:document-meta>
            <office:meta>
                <!-- generator -->
                <meta:generator>ODF Converter <xsl:if test="$app-version">v<xsl:value-of
                            select="$app-version"/></xsl:if></meta:generator>
                <!-- title -->
                <xsl:apply-templates
                    select="document('docProps/core.xml')/cp:coreProperties/dc:title" mode="meta"/>
                <!-- description -->
                <xsl:apply-templates
                    select="document('docProps/core.xml')/cp:coreProperties/dc:description" mode="meta"/>
                <!-- creator -->
                <xsl:apply-templates
                    select="document('docProps/core.xml')/cp:coreProperties/dc:creator" mode="meta"/>
                <!-- creation date -->
                <xsl:apply-templates
                    select="document('docProps/core.xml')/cp:coreProperties/dcterms:created" mode="meta"/>
                <!-- last modifier -->
                <xsl:apply-templates
                    select="document('docProps/core.xml')/cp:coreProperties/cp:lastModifiedBy" mode="meta"/>
                <!-- last modification date -->
                <xsl:apply-templates
                    select="document('docProps/core.xml')/cp:coreProperties/dcterms:modified" mode="meta"/>
                <!-- last print date -->
                <xsl:apply-templates
                    select="document('docProps/core.xml')/cp:coreProperties/cp:lastPrinted" mode="meta"/>
                <!-- subject -->
                <xsl:apply-templates
                    select="document('docProps/core.xml')/cp:coreProperties/dc:subject" mode="meta"/>
                <!-- editing time -->
                <xsl:apply-templates select="document('docProps/core.xml')/Properties/TotalTime" mode="meta"/>
                <!-- revision number -->
                <xsl:apply-templates
                    select="document('docProps/core.xml')/cp:coreProperties/cp:revision" mode="meta"/>
                <!-- custom properties-->
                <xsl:apply-templates
                    select="document('docProps/custom.xml')/cust-p:Properties/cust-p:property" mode="meta"/>
                <!-- keywords -->
                <xsl:apply-templates
                    select="document('docProps/core.xml')/cp:coreProperties/cp:keywords" mode="meta"/>
                <!-- language -->
                <xsl:apply-templates
                    select="document('docProps/core.xml')/cp:coreProperties/dc:language" mode="meta"/>

                <!-- document statistics -->
                <xsl:call-template name="GetDocumentStatistics"/>
            </office:meta>
        </office:document-meta>
    </xsl:template>

    <!-- title -->
    <xsl:template match="dc:title" mode="meta">
        <dc:title>
            <xsl:value-of select="."/>
        </dc:title>
    </xsl:template>

    <!-- description -->
    <xsl:template match="dc:description" mode="meta">
        <dc:description>
            <xsl:value-of select="."/>
        </dc:description>
    </xsl:template>

    <!-- creator -->
    <xsl:template match="dc:creator" mode="meta">
        <meta:initial-creator>
            <xsl:value-of select="."/>
        </meta:initial-creator>
    </xsl:template>

    <!-- creation date -->
    <xsl:template match="dcterms:created" mode="meta">
        <meta:creation-date>
            <xsl:value-of select="."/>
        </meta:creation-date>
    </xsl:template>

    <!-- last modifier -->
    <xsl:template match="cp:lastModifiedBy" mode="meta">
        <dc:creator>
            <xsl:value-of select="."/>
        </dc:creator>
    </xsl:template>

    <!-- last modification date -->
    <xsl:template match="dcterms:modified" mode="meta">
        <dc:date>
            <xsl:value-of select="."/>
        </dc:date>
    </xsl:template>

    <!-- last print date -->
    <xsl:template match="cp:lastPrinted" mode="meta">
        <meta:print-date>
            <xsl:value-of select="."/>
        </meta:print-date>
    </xsl:template>

    <!-- subject -->
    <xsl:template match="dc:subject" mode="meta">
        <dc:subject>
            <xsl:value-of select="."/>
        </dc:subject>
    </xsl:template>

    <!-- editing time -->
    <xsl:template match="TotalTime" mode="meta">
        <meta:editing-duration>
            <xsl:choose>
                <xsl:when test=". &gt; 60">
                    <xsl:text>PT</xsl:text>
                    <xsl:variable name="hours">
                        <xsl:value-of select="round(.) div 60"/>
                    </xsl:variable>
                    <xsl:value-of select="$hours"/>
                    <xsl:text>H</xsl:text>
                    <xsl:value-of select="number(.) - (number($hours) * 60)"/>
                    <xsl:text>M</xsl:text>
                    <xsl:text>0S</xsl:text>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:text>PT</xsl:text>
                    <xsl:text>0H</xsl:text>
                    <xsl:value-of select="."/>
                    <xsl:text>M</xsl:text>
                    <xsl:text>0S</xsl:text>
                </xsl:otherwise>
            </xsl:choose>
        </meta:editing-duration>
    </xsl:template>


    <!-- revision number -->
    <xsl:template match="cp:revision" mode="meta">
        <meta:editing-cycles>
            <xsl:value-of select="."/>
        </meta:editing-cycles>
    </xsl:template>

    <!-- custom properties-->
    <xsl:template match="cust-p:property" mode="meta">
        <xsl:for-each select=".">
            <meta:user-defined meta:name="{@name}">
                <xsl:attribute name="meta:type">
                    <xsl:choose>
                        <xsl:when test="vt:bool">
                            <xsl:text>boolean</xsl:text>
                        </xsl:when>
                        <xsl:when test="vt:filetime">
                            <xsl:text>date</xsl:text>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:text>string</xsl:text>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:attribute>
                <xsl:value-of select="child::node()/text()"/>
            </meta:user-defined>
        </xsl:for-each>
    </xsl:template>

    <!-- keywords -->
    <xsl:template match="cp:keywords" mode="meta">
        <meta:keyword>
            <xsl:value-of select="."/>
        </meta:keyword>
    </xsl:template>

    <!-- language -->
    <xsl:template match="dc:language" mode="meta">
        <dc:language>
            <xsl:value-of select="."/>
        </dc:language>
    </xsl:template>

    <xsl:template name="GetDocumentStatistics">
        <meta:document-statistic>
            <xsl:for-each select="document('docProps/app.xml')/Properties">
                <!-- word count -->
                <xsl:if test="Word">
                    <xsl:attribute name="meta:word-count">
                        <xsl:value-of select="Word"/>
                    </xsl:attribute>
                </xsl:if>
                <!-- page count -->
                <xsl:if test="Pages">
                    <xsl:attribute name="meta:page-count">
                        <xsl:value-of select="Pages"/>
                    </xsl:attribute>
                </xsl:if>
                <!-- paragraph count -->
                <xsl:if test="Paragraphs">
                    <xsl:attribute name="meta:paragraph-count">
                        <xsl:value-of select="Paragraphs"/>
                    </xsl:attribute>
                </xsl:if>
                <!-- character count -->
                <xsl:if test="Characters">
                    <xsl:attribute name="meta:non-whitespace-character-count">
                        <xsl:value-of select="Characters"/>
                    </xsl:attribute>
                </xsl:if>
                <!-- character count -->
                <xsl:if test="CharactersWithSpaces">
                    <xsl:attribute name="meta:character-count">
                        <xsl:value-of select="CharactersWithSpaces"/>
                    </xsl:attribute>
                </xsl:if>
            </xsl:for-each>
        </meta:document-statistic>
    </xsl:template>
    <!-- other properties to fit :
        core : category, contentStatus, contentType
        app : Application, AppVersion, Company, DigSig, DocSecurity, HeadingPairs, HiddenSlides, HLinks, HyperlinkBase, HyperlinksChanged,
        Lines, LinksUpToDate, Manager, MMClips, Notes, PresentationFormat, Properties, ScaleCrop, SharedDoc, Slides, Template, TitlesOfParts.
    -->


</xsl:stylesheet>

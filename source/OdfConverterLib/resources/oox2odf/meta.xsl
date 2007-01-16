<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:cust-p="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:dcmitype="http://purl.org/dc/dcmitype/"
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:xlink="http://www.w3.org/1999/xlink"
    xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
    xmlns:ooo="http://openoffice.org/2004/office" office:version="1.0"
    xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <xsl:template name="meta">
        <office:document-meta>
            <office:meta>
                <!-- title -->
                <xsl:if test="document('docProps/core.xml')/cp:coreProperties/dc:title">
                    <dc:title>
                        <xsl:value-of
                            select="document('docProps/core.xml')/cp:coreProperties/dc:title"/>
                    </dc:title>
                </xsl:if>
                <!-- description -->
                <xsl:if test="document('docProps/core.xml')/cp:coreProperties/dc:description">
                    <dc:description>
                        <xsl:value-of
                            select="document('docProps/core.xml')/cp:coreProperties/dc:description"
                        />
                    </dc:description>
                </xsl:if>
                <!-- creator -->
                <xsl:if test="document('docProps/core.xml')/cp:coreProperties/dc:creator">
                    <meta:initial-creator>
                        <xsl:value-of
                            select="document('docProps/core.xml')/cp:coreProperties/dc:creator"/>
                    </meta:initial-creator>
                </xsl:if>
                <!-- creation date -->
                <xsl:if test="document('docProps/core.xml')/cp:coreProperties/dcterms:created">
                    <meta:creation-date>
                        <xsl:value-of
                            select="document('docProps/core.xml')/cp:coreProperties/dcterms:created"
                        />
                    </meta:creation-date>
                </xsl:if>
                <!-- last modifier -->
                <xsl:if test="document('docProps/core.xml')/cp:coreProperties/cp:lastModifiedBy">
                    <dc:creator>
                        <xsl:value-of
                            select="document('docProps/core.xml')/cp:coreProperties/lastModifiedBy"
                        />
                    </dc:creator>
                </xsl:if>
                <!-- last modification date -->
                <xsl:if test="document('docProps/core.xml')/cp:coreProperties/dcterms:modified">
                    <dc:date>
                        <xsl:value-of
                            select="document('docProps/core.xml')/cp:coreProperties/dcterms:modified"
                        />
                    </dc:date>
                </xsl:if>
                <!-- last print date -->
                <xsl:if test="document('docProps/core.xml')/cp:coreProperties/cp:lastPrinted">
                    <meta:print-date>
                        <xsl:value-of
                            select="document('docProps/core.xml')/cp:coreProperties/lastPrinted"/>
                    </meta:print-date>
                </xsl:if>
                <!-- subject -->
                <xsl:if test="document('docProps/core.xml')/cp:coreProperties/dc:subject">
                    <dc:subject>
                        <xsl:value-of
                            select="document('docProps/core.xml')/cp:coreProperties/dc:subject"/>
                    </dc:subject>
                </xsl:if>
                <!-- editing time -->
                <xsl:if test="document('docProps/core.xml')/Properties/TotalTime">
                    <meta:editing-duration>
                        <xsl:choose>
                            <xsl:when
                                test="document('docProps/core.xml')/Properties/TotalTime &gt; 60">
                                <xsl:text>PT</xsl:text>
                                <xsl:variable name="hours">
                                    <xsl:value-of
                                        select="round(number(document('docProps/core.xml')/Properties/TotalTime) div 60)"
                                    />
                                </xsl:variable>
                                <xsl:value-of select="$hours"/>
                                <xsl:text>H</xsl:text>
                                <xsl:value-of
                                    select="number(document('docProps/core.xml')/Properties/TotalTime) - (number($hours) * 60)"/>
                                <xsl:text>M</xsl:text>
                                <xsl:text>0S</xsl:text>
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:text>PT</xsl:text>
                                <xsl:text>0H</xsl:text>
                                <xsl:value-of
                                    select="document('docProps/core.xml')/Properties/TotalTime"/>
                                <xsl:text>M</xsl:text>
                                <xsl:text>0S</xsl:text>
                            </xsl:otherwise>
                        </xsl:choose>
                    </meta:editing-duration>
                </xsl:if>
                <!-- revision number -->
                <xsl:if test="document('docProps/core.xml')/cp:coreProperties/cp:revision">
                    <meta:editing-cycles>
                        <xsl:value-of
                            select="document('docProps/core.xml')/cp:coreProperties/cp:modified"/>
                    </meta:editing-cycles>
                </xsl:if>
                <!-- custom properties-->
                <xsl:if test="document('docProps/custom.xml')/cust-p:Properties/cust-p:property">
                    <xsl:for-each
                        select="document('docProps/custom.xml')/cust-p:Properties/cust-p:property">
                        <xsl:variable name="numer">
                            <xsl:number/>
                        </xsl:variable>
                        <xsl:if test="$numer &lt; 5">
                            <meta:user-defined meta:name="Info {$numer}">
                                <xsl:value-of select="@name"/>
                                <xsl:text>: </xsl:text>
                                <xsl:value-of select="vt:lpwstr/text()"/>
                            </meta:user-defined>
                        </xsl:if>
                    </xsl:for-each>
                </xsl:if>
                <!-- keywords -->
                <xsl:if test="document('docProps/core.xml')/cp:coreProperties/cp:keywords">
                    <meta:keyword>
                        <xsl:value-of
                            select="document('docProps/core.xml')/cp:coreProperties/cp:keywords"/>
                    </meta:keyword>
                </xsl:if>
                <!-- language -->
                <xsl:if test="document('docProps/core.xml')/cp:coreProperties/dc:language">
                    <dc:language>
                        <xsl:value-of
                            select="document('docProps/core.xml')/cp:coreProperties/dc:language"/>
                    </dc:language>
                </xsl:if>
                <!-- document statistics -->
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
                <!-- other properties to fit :
                    core : category, contentStatus, contentType
                    app : Application, AppVersion, Company, DigSig, DocSecurity, HeadingPairs, HiddenSlides, HLinks, HyperlinkBase, HyperlinksChanged,
                    Lines, LinksUpToDate, Manager, MMClips, Notes, PresentationFormat, Properties, ScaleCrop, SharedDoc, Slides, Template, TitlesOfParts.
                -->
            </office:meta>
        </office:document-meta>
    </xsl:template>
</xsl:stylesheet>

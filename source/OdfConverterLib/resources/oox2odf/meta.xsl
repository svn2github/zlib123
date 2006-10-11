<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet 
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/" 
    xmlns:dcterms="http://purl.org/dc/terms/"
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
                <dc:title>
                    <xsl:value-of select="document('docProps/core.xml')/cp:coreProperties/dc:title"/>
                </dc:title>
                <dc:description>
                    <xsl:value-of select="document('docProps/core.xml')/cp:coreProperties/dc:description"/>
                </dc:description>
                <dc:subject>
                    <xsl:value-of select="document('docProps/core.xml')/cp:coreProperties/dc:subject"/>
                </dc:subject>
                <meta:creation-date>
                    <xsl:value-of select="document('docProps/core.xml')/cp:coreProperties/dcterms:created"/>
                </meta:creation-date>
                <dc:date>
                    <xsl:value-of select="document('docProps/core.xml')/cp:coreProperties/dcterms:modified"/>
                </dc:date>
                <meta:keyword>
                    <xsl:value-of select="document('docProps/core.xml')/cp:coreProperties/cp:keywords"/>
                </meta:keyword>
                 <!--TODO:<meta:document-statistic>
                    <xsl:attribute name="meta:page-count">
                         <xsl:value-of select="document('docProps/app.xml')/Properties/Pages"/>
                    </xsl:attribute>
                </meta:document-statistic>-->
            </office:meta>
        </office:document-meta>
    </xsl:template>
</xsl:stylesheet>

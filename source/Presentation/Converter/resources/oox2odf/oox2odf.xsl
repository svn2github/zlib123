<?xml version="1.0" encoding="UTF-8"?>


<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:oox="urn:oox"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns:rels="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
  exclude-result-prefixes="oox rels">
	<xsl:import href="content.xsl"/>
	<xsl:import href="styles.xsl"/>
	<xsl:import href="settings.xsl"/>
	<xsl:import href="meta.xsl"/>
	

	<xsl:param name="outputFile"/>
	<xsl:output method="xml" encoding="UTF-8"/>

	<!-- App version number -->
	<xsl:variable name="app-version">1.0.0</xsl:variable>
	<xsl:template match="/oox:source">
		<pzip:archive pzip:target="{$outputFile}">
			<!-- Manifest -->
			<pzip:entry pzip:target="META-INF/manifest.xml">
				<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
					<manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.text"
					  manifest:full-path="/"/>
					<manifest:file-entry manifest:media-type="" manifest:full-path="Configurations2/statusbar/"/>
					<manifest:file-entry manifest:media-type="" manifest:full-path="Configurations2/accelerator/current.xml"/>
					<manifest:file-entry manifest:media-type="" manifest:full-path="Configurations2/accelerator/"/>
					<manifest:file-entry manifest:media-type="" manifest:full-path="Configurations2/floater/"/>
					<manifest:file-entry manifest:media-type="" manifest:full-path="Configurations2/popupmenu/"/>
					<manifest:file-entry manifest:media-type="" manifest:full-path="Configurations2/progressbar/"/>
					<manifest:file-entry manifest:media-type="" manifest:full-path="Configurations2/menubar/"/>
					<manifest:file-entry manifest:media-type="" manifest:full-path="Configurations2/toolbar/"/>
					<manifest:file-entry manifest:media-type="" manifest:full-path="Configurations2/images/Bitmaps/"/>
					<manifest:file-entry manifest:media-type="" manifest:full-path="Configurations2/images/"/>
					<manifest:file-entry manifest:media-type="application/vnd.sun.xml.ui.configuration" manifest:full-path="Configurations2/"/>
					<manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
					<manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml"/>
					<manifest:file-entry manifest:media-type="text/xml" manifest:full-path="meta.xml"/>
					<manifest:file-entry manifest:media-type="" manifest:full-path="Thumbnails/"/>
					<manifest:file-entry manifest:media-type="" manifest:full-path="Thumbnails/thumbnail.png"/>
					<manifest:file-entry manifest:media-type="text/xml" manifest:full-path="settings.xml"/>
					<xsl:for-each
					  select="document('ppt/presentation.xml')//node()[name() = 'Relationship'][substring-before(@Target,'/') = 'media']">
						<xsl:call-template name="InsertManifestFileEntry"/>
					</xsl:for-each >
				</manifest:manifest>
			</pzip:entry>
			<pzip:entry pzip:target="content.xml">
				<xsl:call-template name="content" />
			</pzip:entry >
			<pzip:entry pzip:target="styles.xml">
				<xsl:call-template name="styles" />
			</pzip:entry>
			<pzip:entry pzip:target="settings.xml">
				<xsl:call-template name="settings" />
			</pzip:entry>
			<pzip:entry pzip:target="meta.xml">
				<xsl:call-template name="meta" />
			</pzip:entry>
			<pzip:entry pzip:target="Configurations2/accelerator/current.xml">
				<xsl:call-template name="InsertManifestFileEntry"/>
			</pzip:entry>

		</pzip:archive>
	</xsl:template>
	<xsl:template name="InsertManifestFileEntry">
		<manifest:file-entry>
			<xsl:attribute name="manifest:media-type">
				<xsl:if test="substring-after(@Target,'.') = 'gif'">
					<xsl:text>image/gif</xsl:text>
				</xsl:if>
				<xsl:if
				  test="substring-after(@Target,'.') = 'jpg' or substring-after(@Target,'.') = 'jpeg'  or substring-after(@Target,'.') = 'jpe' or substring-after(@Target,'.') = 'jfif' ">
					<xsl:text>image/jpeg</xsl:text>
				</xsl:if>
				<xsl:if test="substring-after(@Target,'.') = 'tif' or substring-after(@Target,'.') = 'tiff'">
					<xsl:text>image/tiff</xsl:text>
				</xsl:if>
				<xsl:if test="substring-after(@Target,'.') = 'png'">
					<xsl:text>image/png</xsl:text>
				</xsl:if>
			</xsl:attribute>
			<xsl:attribute name="manifest:full-path">
				<xsl:text>Pictures/</xsl:text>
				<xsl:value-of select="substring-after(@Target,'/')"/>
			</xsl:attribute>
		</manifest:file-entry>
	</xsl:template>
</xsl:stylesheet>
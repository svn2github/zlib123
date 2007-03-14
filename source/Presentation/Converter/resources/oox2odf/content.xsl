<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:pchar="urn:cleverage:xmlns:post-processings:characters"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" 
  xmlns:text="urn:oasis:names:tc:presentation:xmlns:text:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink" 
  xmlns:dc="http://purl.org/dc/elements/1.1/"  
  xmlns:p="http://schemas.openxmlformats.org/presentation/2006/main" 
  xmlns:r="http://schemas.openxmlformats.org/presentation/2006/relationships"
  xmlns:a="http://schemas.openxmlformats.org/presentation/2006/relationships"
  xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"  
  xmlns:v="urn:schemas-microsoft-com:vml" exclude-result-prefixes="p r a xlink number">

	<!--xsl:import href="tables.xsl"/-->
	<!--xsl:import href="common.xsl"/-->
	<!-- 
	<xsl:import href="lists.xsl"/>
	<xsl:import href="fonts.xsl"/>
	<xsl:import href="fields.xsl"/>
	<xsl:import href="footnotes.xsl"/>
	<xsl:import href="indexes.xsl"/>
	<xsl:import href="track.xsl"/>
	<xsl:import href="frames.xsl"/>
	<xsl:import href="sections.xsl"/>
	<xsl:import href="comments.xsl"/>
    -->
	<xsl:strip-space elements="*"/>
	<xsl:preserve-space elements="p:p"/>
	<xsl:preserve-space elements="p:r"/>

	<xsl:key name="InstrText" match="p:instrText" use="''"/>
	<xsl:key name="bookmarkStart" match="p:bookmarkStart" use="@p:id"/>
	<xsl:key name="pPr" match="p:pPr" use="''"/>
	<xsl:key name="sectPr" match="p:sectPr" use="''"/>

	<!--main document-->
	<xsl:template name="content">
		<office:document-content>
			<office:scripts/>
			<office:font-face-decls>
				<xsl:apply-templates select="document('ppt/fontTable.xml')/p:fonts"/>
			</office:font-face-decls>
			<office:automatic-styles>
			<!-- automatic styles for document body -->
			<xsl:call-template name="InsertBodyStyles"/>
			<xsl:call-template name="InsertListStyles"/>
			<!--  
			<xsl:call-template name="InsertBodyStyles"/>
			<xsl:call-template name="InsertListStyles"/>
			<xsl:call-template name="InsertSectionsStyles"/>
			<xsl:call-template name="InsertFootnoteStyles"/>
			<xsl:call-template name="InsertEndnoteStyles"/>
			<xsl:call-template name="InsertFrameStyle"/> -->
			</office:automatic-styles>
			<office:body>
				<office:text>
				<!-- <xsl:call-template name="TrackChanges"/>
				<xsl:call-template name="InsertDocumentBody"/> -->
				</office:text>
			</office:body>
		</office:document-content>
	</xsl:template>

	<!--  generates automatic styles for paragraphs  ho w does it exactly work ?? -->
	<xsl:template name="InsertBodyStyles">
		<xsl:apply-templates select="document('ppt/slide1.xml')//p:txbody/a:p"  mode="automaticstyles"/>
	</xsl:template>

	<xsl:template name="InsertListStyles">
		<!-- document with lists-->
		<xsl:for-each select="document('ppt/presentation.xml')">
			<xsl:choose>
				<xsl:when test="key('pPr', '')/p:numPr/p:numId">
					<!-- automatic list styles with empty num format for elements which has non-existent w:num attached -->
					<xsl:apply-templates
					  select="key('pPr', '')/p:numPr/p:numId[not(document('ppt/numbering.xml')/p:numbering/p:num/@p:numId = @p:val)][1]"
					  mode="automaticstyles"/>
					<!-- automatic list styles-->
					<xsl:apply-templates select="document('ppt/numbering.xml')/p:numbering/p:num"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:for-each select="document('ppt/styles.xml')">
						<xsl:if test="key('pPr', '')/p:numPr/p:numId">
							<!-- automatic list styles-->
							<xsl:apply-templates select="document('ppt/numbering.xml')/p:numbering/p:num"/>
						</xsl:if>
					</xsl:for-each>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:for-each>
	</xsl:template>
	
</xsl:stylesheet>
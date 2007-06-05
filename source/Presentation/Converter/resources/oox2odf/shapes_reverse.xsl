<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet version="1.0" 
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink" 
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  exclude-result-prefixes="p a dc xlink draw">

	<xsl:import href="common.xsl"/>
	<!--<xsl:import href ="trignm.xsl"/>-->

	<!-- Template for Shapes -->
	<xsl:template  name="shapes">
		<xsl:param name="GraphicId" />
		<xsl:param name="ParaId" />
		<xsl:choose>
			<xsl:when test ="p:nvSpPr/p:cNvPr/@name[contains(., 'Rectangle Custom')]">
				<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="rectangle" draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N"/>
				</draw:custom-shape>
			</xsl:when>
			<xsl:when test ="p:nvSpPr/p:cNvPr/@name[contains(., 'Rectangle')]">
				<draw:rect draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				</draw:rect>
			</xsl:when>
			<xsl:when test ="p:nvSpPr/p:cNvPr/@name[contains(., 'TextBox Custom')]">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="mso-spt202" draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N"/>
				</draw:custom-shape>
			</xsl:when>
			<xsl:when test ="p:nvSpPr/p:cNvPr/@name[contains(., 'TextBox')] or p:nvSpPr/p:cNvPr/@name[contains(., 'Text Box')]">
				<draw:frame draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				</draw:frame>
			</xsl:when>
			<!--<xsl:when test ="p:nvCxnSpPr/p:cNvPr/@name[contains(., 'Straight Connector')]">
				<draw:line draw:layer="layout">
					<xsl:call-template name ="DrawLine">
						<xsl:with-param name="grID" select ="$GraphicId"/>
					</xsl:call-template>
				</draw:line>
			</xsl:when>-->
		</xsl:choose>
	</xsl:template>

	<!--  Draw Shape reading values from pptx p:spPr-->
	<xsl:template name ="CreateShape">
		<xsl:param name ="grID" />
		<xsl:param name ="prID" />

		<xsl:attribute name ="draw:style-name">
			<xsl:value-of select ="$grID"/>
		</xsl:attribute>
		<xsl:attribute name ="draw:text-style-name">
			<xsl:value-of select ="$prID"/>
		</xsl:attribute>

		<xsl:for-each select ="p:spPr/a:xfrm">
			<xsl:attribute name ="svg:width">
				<xsl:call-template name="ConvertEmu">
					<xsl:with-param name="length" select="a:ext/@cx"/>
					<xsl:with-param name="unit">cm</xsl:with-param>
				</xsl:call-template>
			</xsl:attribute>

			<xsl:attribute name ="svg:height">
				<xsl:call-template name="ConvertEmu">
					<xsl:with-param name="length" select="a:ext/@cy"/>
					<xsl:with-param name="unit">cm</xsl:with-param>
				</xsl:call-template>
			</xsl:attribute>

			<xsl:attribute name ="svg:x">
				<xsl:call-template name="ConvertEmu">
					<xsl:with-param name="length" select="a:off/@x"/>
					<xsl:with-param name="unit">cm</xsl:with-param>
				</xsl:call-template>
			</xsl:attribute>

			<xsl:attribute name ="svg:y">
				<xsl:call-template name="ConvertEmu">
					<xsl:with-param name="length" select="a:off/@y"/>
					<xsl:with-param name="unit">cm</xsl:with-param>
				</xsl:call-template>
			</xsl:attribute>
		</xsl:for-each>
		<xsl:choose>
			<xsl:when test ="(p:nvSpPr/p:cNvPr/@name[contains(., 'TextBox')] or p:nvSpPr/p:cNvPr/@name[contains(., 'Text Box')]) and not(p:nvSpPr/p:cNvPr/@name[contains(., 'TextBox Custom')])">
				<draw:text-box>
					<xsl:call-template name ="AddShapeText">
						<xsl:with-param name ="prID" select ="$prID" />
					</xsl:call-template>
				</draw:text-box>
			</xsl:when>
			<xsl:otherwise>
				<xsl:call-template name ="AddShapeText">
					<xsl:with-param name ="prID" select ="$prID" />
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<xsl:template name ="DrawLine">
		<xsl:param name ="grID" />

		<xsl:attribute name ="draw:style-name">
			<xsl:value-of select ="$grID"/>
		</xsl:attribute>

		<xsl:for-each select ="p:spPr/a:xfrm">
			<xsl:if test ="not(@flipV)">
				<xsl:attribute name ="svg:x1">
					<xsl:call-template name="ConvertEmu">
						<xsl:with-param name="length" select="a:off/@x"/>
						<xsl:with-param name="unit">cm</xsl:with-param>
					</xsl:call-template>
				</xsl:attribute>
				<xsl:attribute name ="svg:y1">
					<xsl:call-template name="ConvertEmu">
						<xsl:with-param name="length" select="a:off/@y"/>
						<xsl:with-param name="unit">cm</xsl:with-param>
					</xsl:call-template>
				</xsl:attribute>
				<xsl:attribute name ="svg:x2">
					<xsl:call-template name="ConvertEmu">
						<xsl:with-param name="length" select="a:off/@x + a:ext/@cx"/>
						<xsl:with-param name="unit">cm</xsl:with-param>
					</xsl:call-template>
				</xsl:attribute>
				<xsl:attribute name ="svg:y2">
					<xsl:call-template name="ConvertEmu">
						<xsl:with-param name="length" select="a:off/@y + a:ext/@cy"/>
						<xsl:with-param name="unit">cm</xsl:with-param>
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>
			<xsl:if test ="@flipV">
				<xsl:attribute name ="svg:x1">
					<xsl:call-template name="ConvertEmu">
						<xsl:with-param name="length" select="a:off/@x"/>
						<xsl:with-param name="unit">cm</xsl:with-param>
					</xsl:call-template>
				</xsl:attribute>
				<xsl:attribute name ="svg:y1">
					<xsl:call-template name="ConvertEmu">
						<xsl:with-param name="length" select="a:off/@y + a:ext/@cy"/>
						<xsl:with-param name="unit">cm</xsl:with-param>
					</xsl:call-template>
				</xsl:attribute>
				<xsl:attribute name ="svg:x2">
					<xsl:call-template name="ConvertEmu">
						<xsl:with-param name="length" select="a:off/@x + a:ext/@cx"/>
						<xsl:with-param name="unit">cm</xsl:with-param>
					</xsl:call-template>
				</xsl:attribute>
				<xsl:attribute name ="svg:y2">
					<xsl:call-template name="ConvertEmu">
						<xsl:with-param name="length" select="a:off/@y"/>
						<xsl:with-param name="unit">cm</xsl:with-param>
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>
			<!--<xsl:variable name ="xCenter">
				<xsl:value-of select ="(a:off/@x + a:ext/@cx) div 2"/>
			</xsl:variable>
			<xsl:variable name ="yCenter">
				<xsl:value-of select ="(a:off/@y + a:ext/@cy) div 2"/>
			</xsl:variable>
			<xsl:variable name ="angle">
				<xsl:value-of select ="(@rot div 60000) * ((22 div 7) div 180)"/>
			</xsl:variable>
			<xsl:variable name ="cxBy2">
				<xsl:if test ="@flipH">
					<xsl:value-of select ="(-1 * a:ext/@cx) div 2"/>
				</xsl:if>
				<xsl:if test ="not(@flipH)">
					<xsl:value-of select ="a:ext/@cx div 2"/>
				</xsl:if>
			</xsl:variable>
			<xsl:variable name ="cyBy2">
				<xsl:if test ="@flipV">
					<xsl:value-of select ="(-1 * a:ext/@cy) div 2"/>
				</xsl:if>
				<xsl:if test ="not(@flipV)">
					<xsl:value-of select ="a:ext/@cy div 2"/>
				</xsl:if>
			</xsl:variable>
			<xsl:attribute name ="svg:x1">
				<xsl:value-of select ="$xCenter - m:cos($angle) * $cxBy2 + m:sin($angle) * $cyBy2"/>
			</xsl:attribute>
			<xsl:attribute name ="svg:y1">
				<xsl:value-of select ="$yCenter - m:sin($angle) * $cxBy2 - m:cos($angle) * $cyBy2"/>
			</xsl:attribute>
			<xsl:attribute name ="svg:x2">
				<xsl:value-of select ="$xCenter + m:cos($angle) * $cxBy2 - m:sin($angle) * $cyBy2"/>
			</xsl:attribute>
			<xsl:attribute name ="svg:y2">
				<xsl:value-of select ="$yCenter + m:sin($angle) * $cxBy2 + m:cos($angle) * $cyBy2"/>
			</xsl:attribute>-->

		</xsl:for-each>

	</xsl:template>

	<!--  Add text to the shape -->
	<xsl:template name ="AddShapeText">
		<xsl:param name ="prID" />

		<xsl:for-each select ="p:txBody/a:p">
			<text:p >
				<xsl:attribute name ="text:style-name">
					<xsl:value-of select ="concat($prID,position())"/>
				</xsl:attribute>
				<xsl:for-each select ="node()">
					<xsl:if test ="name()='a:r'">
						<text:span text:style-name="{generate-id()}">
							<!--<xsl:value-of select ="a:t"/>-->
							<!--converts whitespaces sequence to text:s-->
							<!-- 1699083 bug fix  -->
							<xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
							<xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
							<xsl:choose >
								<xsl:when test ="a:rPr[@cap='all']">
									<xsl:choose >
										<xsl:when test =".=''">
											<text:s/>
										</xsl:when>
										<xsl:when test ="not(contains(.,'  '))">
											<xsl:value-of select ="translate(.,$lcletters,$ucletters)"/>
										</xsl:when>
										<xsl:otherwise >
											<xsl:call-template name ="InsertWhiteSpaces">
												<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
											</xsl:call-template>
										</xsl:otherwise>
									</xsl:choose>
								</xsl:when>
								<xsl:when test ="a:rPr[@cap='small']">
									<xsl:choose >
										<xsl:when test =".=''">
											<text:s/>
										</xsl:when>
										<xsl:when test ="not(contains(.,'  '))">
											<xsl:value-of select ="translate(.,$ucletters,$lcletters)"/>
										</xsl:when>
										<xsl:otherwise >
											<xsl:call-template name ="InsertWhiteSpaces">
												<xsl:with-param name ="string" select ="translate(.,$ucletters,$lcletters)"/>
											</xsl:call-template>
										</xsl:otherwise>
									</xsl:choose >
								</xsl:when>
								<xsl:otherwise >
									<xsl:choose >
										<xsl:when test =".=''">
											<text:s/>
										</xsl:when>
										<xsl:when test ="not(contains(.,'  '))">
											<xsl:value-of select ="."/>
										</xsl:when>
										<xsl:otherwise >
											<xsl:call-template name ="InsertWhiteSpaces">
												<xsl:with-param name ="string" select ="."/>
											</xsl:call-template>
										</xsl:otherwise >
									</xsl:choose>
								</xsl:otherwise>
							</xsl:choose>
						</text:span>
					</xsl:if >
					<xsl:if test ="name()='a:br'">
						<text:line-break/>
					</xsl:if>
				</xsl:for-each>
			</text:p>
		</xsl:for-each>

	</xsl:template>

	<!-- Generate autometic styles in contet.xsl for graphic properties-->
	<xsl:template name ="InsertStylesForGraphicProperties"  >
		<xsl:for-each select ="document('ppt/presentation.xml')/p:presentation/p:sldIdLst/p:sldId">
			<xsl:variable name ="SlideId">
				<xsl:value-of  select ="concat(concat('slide',position()),'.xml')" />
			</xsl:variable>

			<xsl:for-each select ="document(concat('ppt/slides/',$SlideId))/p:sld/p:cSld/p:spTree/p:sp">
				<!--Check for shape or texbox -->
				<xsl:if test = "p:style or p:nvSpPr/p:cNvPr/@name[contains(., 'TextBox')] or p:style or p:nvSpPr/p:cNvPr/@name[contains(., 'Text Box')] or p:nvSpPr/p:cNvPr/@name[contains(., 'Rectangle')]">
					<!-- Generate graphic properties ID-->
					<xsl:variable  name ="GraphicId">
						<xsl:value-of select ="concat('s',substring($SlideId,6,string-length($SlideId)-9) ,concat('gr',position()))"/>
					</xsl:variable>

					<style:style style:family="graphic" style:parent-style-name="standard">
						<xsl:attribute name ="style:name">
							<xsl:value-of select ="$GraphicId"/>
						</xsl:attribute >
						<style:graphic-properties>

							<!-- FILL -->
							<xsl:call-template name ="Fill" />

							<!-- LINE COLOR -->
							<xsl:call-template name ="LineColor" />

							<!-- LINE STYLE -->
							<xsl:call-template name ="LineStyle"/>

							<!-- TEXT ALIGNMENT -->
							<xsl:call-template name ="TextLayout" />

						</style:graphic-properties >
						<xsl:if test ="p:txBody/a:bodyPr/@vert">
							<style:paragraph-properties>
								<xsl:attribute name ="style:writing-mode">
									<xsl:call-template name ="getTextDirection">
										<xsl:with-param name ="vert" select ="p:txBody/a:bodyPr/@vert" />
									</xsl:call-template>
								</xsl:attribute>
							</style:paragraph-properties>
						</xsl:if>
					</style:style>
				</xsl:if >
			</xsl:for-each>

			<xsl:for-each select ="document(concat('ppt/slides/',$SlideId))/p:sld/p:cSld/p:spTree/p:cxnSp">
				<!-- Generate graphic properties ID-->
				<xsl:variable  name ="GraphicId">
					<xsl:value-of select ="concat('s',substring($SlideId,6,string-length($SlideId)-9) ,concat('grLine',position()))"/>
				</xsl:variable>

				<style:style style:family="graphic">
					<xsl:attribute name ="style:name">
						<xsl:value-of select ="$GraphicId"/>
					</xsl:attribute >
					<style:graphic-properties>

						<!-- LINE COLOR -->
						<xsl:call-template name ="LineColor" />

						<!-- LINE STYLE -->
						<xsl:call-template name ="LineStyle"/>

					</style:graphic-properties >
				</style:style>

			</xsl:for-each>
		</xsl:for-each>
	</xsl:template>

	<!-- Get fill color for shape-->
	<xsl:template name="Fill">
		<xsl:choose>
			<!-- No fill -->
			<xsl:when test ="p:spPr/a:noFill">
				<xsl:attribute name ="draw:fill">
					<xsl:value-of select="'none'" />
				</xsl:attribute>
				<xsl:attribute name ="draw:fill-color">
					<xsl:value-of select="'#ffffff'"/>
				</xsl:attribute>
			</xsl:when>

			<!-- Solid fill-->
			<xsl:when test ="p:spPr/a:solidFill">
				<xsl:attribute name ="draw:fill">
					<xsl:value-of select="'solid'" />
				</xsl:attribute>
				<!-- Standard color-->
				<xsl:if test ="p:spPr/a:solidFill/a:srgbClr[@val]">
					<xsl:attribute name ="draw:fill-color">
						<xsl:value-of select="concat('#',p:spPr/a:solidFill/a:srgbClr/@val)"/>
					</xsl:attribute>
					<!-- Transparency percentage-->
					<xsl:if test="p:spPr/a:solidFill/a:srgbClr/a:alpha/@val">
						<xsl:variable name ="alpha">
							<xsl:value-of select ="p:spPr/a:solidFill/a:srgbClr/a:alpha/@val"/>
						</xsl:variable>
						<xsl:if test="($alpha != '') or ($alpha != 0)">
							<xsl:attribute name ="draw:opacity">
								<xsl:value-of select="concat(($alpha div 1000), '%')"/>
							</xsl:attribute>
						</xsl:if>
					</xsl:if>
				</xsl:if>

				<!--Theme color-->
				<xsl:if test ="p:spPr/a:solidFill/a:schemeClr[@val]">
					<xsl:attribute name ="draw:fill-color">
						<xsl:call-template name ="getColorCode">
							<xsl:with-param name ="color">
								<xsl:value-of select="p:spPr/a:solidFill/a:schemeClr/@val"/>
							</xsl:with-param>
							<xsl:with-param name ="lumMod">
								<xsl:value-of select="p:spPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
							</xsl:with-param>
							<xsl:with-param name ="lumOff">
								<xsl:value-of select="p:spPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
							</xsl:with-param>
						</xsl:call-template>
					</xsl:attribute>
					<!-- Transparency percentage-->
					<xsl:if test="p:spPr/a:solidFill/a:schemeClr/a:alpha/@val">
						<xsl:variable name ="alpha">
							<xsl:value-of select ="p:spPr/a:solidFill/a:schemeClr/a:alpha/@val"/>
						</xsl:variable>
						<xsl:if test="($alpha != '') or ($alpha != 0)">
							<xsl:attribute name ="draw:opacity">
								<xsl:value-of select="concat(($alpha div 1000), '%')"/>
							</xsl:attribute>
						</xsl:if>
					</xsl:if>
				</xsl:if>
			</xsl:when>

			<!--Fill refernce-->
			<xsl:when test ="p:style/a:fillRef">
				<xsl:attribute name ="draw:fill">
					<xsl:value-of select="'solid'" />
				</xsl:attribute>
				<!-- Standard color-->
				<xsl:if test ="p:style/a:fillRef/a:srgbClr[@val]">
					<xsl:attribute name ="draw:fill-color">
						<xsl:value-of select="concat('#',p:style/a:fillRef/a:srgbClr/@val)"/>
					</xsl:attribute>
					<!-- Shade percentage-->
					<!--<xsl:if test="p:style/a:fillRef/a:srgbClr/a:shade/@val">
						<xsl:variable name ="shade">
							<xsl:value-of select ="a:solidFill/a:srgbClr/a:shade/@val"/>
						</xsl:variable>
						<xsl:if test="($shade != '') or ($shade != 0)">
							<xsl:attribute name ="draw:shadow-opacity">
								<xsl:value-of select="concat(($shade div 1000), '%')"/>
							</xsl:attribute>
						</xsl:if>
					</xsl:if>-->
				</xsl:if>

				<!--Theme color-->
				<xsl:if test ="p:style/a:fillRef//a:schemeClr[@val]">
					<xsl:attribute name ="draw:fill-color">
						<xsl:call-template name ="getColorCode">
							<xsl:with-param name ="color">
								<xsl:value-of select="p:style/a:fillRef/a:schemeClr/@val"/>
							</xsl:with-param>
							<xsl:with-param name ="lumMod">
								<xsl:value-of select="p:style/a:fillRef/a:schemeClr/a:lumMod/@val"/>
							</xsl:with-param>
							<xsl:with-param name ="lumOff">
								<xsl:value-of select="p:style/a:fillRef/a:schemeClr/a:lumOff/@val"/>
							</xsl:with-param>
						</xsl:call-template>
					</xsl:attribute>
					<!-- Shade percentage-->
					<!--<xsl:if test="a:solidFill/a:schemeClr/a:shade/@val">
						<xsl:variable name ="shade">
							<xsl:value-of select ="a:solidFill/a:schemeClr/a:shade/@val"/>
						</xsl:variable>
						<xsl:if test="($shade != '') or ($shade != 0)">
							<xsl:attribute name ="draw:shadow-opacity">
								<xsl:value-of select="concat(($shade div 1000), '%')"/>
							</xsl:attribute>
						</xsl:if>
					</xsl:if>-->
				</xsl:if>
			</xsl:when>
		</xsl:choose>
	</xsl:template>

	<!-- Get border color for shape -->
	<xsl:template name ="LineColor">

		<xsl:choose>
			<!-- No line-->
			<xsl:when test ="p:spPr/a:ln/a:noFill">
				<xsl:attribute name ="draw:stroke">
					<xsl:value-of select="'none'" />
				</xsl:attribute>
			</xsl:when>

			<!-- Solid line color-->
			<xsl:when test ="p:spPr/a:ln/a:solidFill">
				<xsl:attribute name ="draw:stroke">
					<xsl:value-of select="'solid'" />
				</xsl:attribute>
				<!-- Standard color for border-->
				<xsl:if test ="p:spPr/a:ln/a:solidFill/a:srgbClr[@val]">
					<xsl:attribute name ="svg:stroke-color">
						<xsl:value-of select="concat('#',p:spPr/a:ln/a:solidFill/a:srgbClr/@val)"/>
					</xsl:attribute>
					<!-- Transparency percentage-->
					<xsl:if test="p:spPr/a:ln/a:solidFill/a:srgbClr/a:alpha/@val">
						<xsl:variable name ="alpha">
							<xsl:value-of select ="p:spPr/a:ln/a:solidFill/a:srgbClr/a:alpha/@val"/>
						</xsl:variable>
						<xsl:if test="($alpha != '') or ($alpha != 0)">
							<xsl:attribute name ="svg:stroke-opacity">
								<xsl:value-of select="concat(($alpha div 1000), '%')"/>
							</xsl:attribute>
						</xsl:if>
					</xsl:if>
				</xsl:if>
				<!-- Theme color for border-->
				<xsl:if test ="p:spPr/a:ln/a:solidFill/a:schemeClr[@val]">
					<xsl:attribute name ="svg:stroke-color">
						<xsl:call-template name ="getColorCode">
							<xsl:with-param name ="color">
								<xsl:value-of select="p:spPr/a:ln/a:solidFill/a:schemeClr/@val"/>
							</xsl:with-param>
							<xsl:with-param name ="lumMod">
								<xsl:value-of select="p:spPr/a:ln/a:solidFill/a:schemeClr/a:lumMod/@val"/>
							</xsl:with-param>
							<xsl:with-param name ="lumOff">
								<xsl:value-of select="p:spPr/a:ln/a:solidFill/a:schemeClr/a:lumOff/@val"/>
							</xsl:with-param>
						</xsl:call-template>
					</xsl:attribute>
					<!-- Transparency percentage-->
					<xsl:if test="p:spPr/a:ln/a:solidFill/a:schemeClr/a:alpha/@val">
						<xsl:variable name ="alpha">
							<xsl:value-of select ="p:spPr/a:ln/a:solidFill/a:schemeClr/a:alpha/@val"/>
						</xsl:variable>
						<xsl:if test="($alpha != '') or ($alpha != 0)">
							<xsl:attribute name ="svg:stroke-opacity">
								<xsl:value-of select="concat(($alpha div 1000), '%')"/>
							</xsl:attribute>
						</xsl:if>
					</xsl:if>
				</xsl:if>
			</xsl:when>

			<!--Line reference-->
			<xsl:when test ="p:style/a:lnRef">
				<xsl:attribute name ="draw:stroke">
					<xsl:value-of select="'solid'" />
				</xsl:attribute>
				<!--Standard color for border-->
				<xsl:if test ="p:style/a:lnRef/a:srgbClr[@val]">
					<!--<xsl:attribute name ="svg:stroke-color">
						<xsl:value-of select="concat('#',p:style/a:lnRef/a:srgbClr/@val)"/>
					</xsl:attribute>
					-->
					<!--Shade percentage-->
					<!--
					<xsl:if test="p:style/a:lnRef/a:srgbClr/a:shade/@val">
						<xsl:variable name ="shade">
							<xsl:value-of select ="p:style/a:lnRef/a:srgbClr/a:shade/@val"/>
						</xsl:variable>
						-->
					<!--<xsl:if test="($shade != '') or ($shade != 0)">
							<xsl:attribute name ="svg:stroke-opacity">
								<xsl:value-of select="concat(($shade div 1000), '%')"/>
							</xsl:attribute>
						</xsl:if>-->
					<!--
					</xsl:if>-->
				</xsl:if>
				<!--Theme color for border-->
				<xsl:if test ="p:style/a:lnRef/a:schemeClr[@val]">
					<!--<xsl:attribute name ="svg:stroke-color">
						<xsl:call-template name ="getColorCode">
							<xsl:with-param name ="color">
								<xsl:value-of select="p:style/a:lnRef/a:schemeClr/@val"/>
							</xsl:with-param>
							<xsl:with-param name ="lumMod">
								<xsl:value-of select="p:style/a:lnRef/a:schemeClr/a:lumMod/@val"/>
							</xsl:with-param>
							<xsl:with-param name ="lumOff">
								<xsl:value-of select="p:style/a:lnRef/a:schemeClr/a:lumOff/@val"/>
							</xsl:with-param>
						</xsl:call-template>
					</xsl:attribute>-->
					<!--Shade percentage -->
					<!--<xsl:if test="p:style/a:lnRef/a:schemeClr/a:shade/@val">
						<xsl:variable name ="shade">
							<xsl:value-of select ="p:style/a:lnRef/a:schemeClr/a:shade/@val"/>
						</xsl:variable>
						-->
					<!--<xsl:if test="($shade != '') or ($shade != 0)">
							<xsl:attribute name ="svg:stroke-opacity">
								<xsl:value-of select="concat(($shade div 1000), '%')"/>
							</xsl:attribute>
						</xsl:if>-->
					<!--
					</xsl:if>-->
				</xsl:if>
			</xsl:when>
			<xsl:otherwise>
				<xsl:attribute name ="draw:stroke">
					<xsl:value-of select="'none'" />
				</xsl:attribute>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- Get line styles for shape -->
	<xsl:template name ="LineStyle">
		<!-- Line width-->
		<xsl:for-each select ="p:spPr">
			<xsl:if test ="a:ln[@w]">
				<xsl:for-each select ="a:ln[@w]">
					<xsl:attribute name ="svg:stroke-width">
						<xsl:call-template name="ConvertEmu">
							<xsl:with-param name="length" select="@w"/>
							<xsl:with-param name="unit">cm</xsl:with-param>
						</xsl:call-template>
					</xsl:attribute>
				</xsl:for-each>
			</xsl:if>
		</xsl:for-each>

		<!-- Line Dash property-->
		<xsl:for-each select ="p:spPr/a:ln">
			<xsl:if test ="not(a:noFill)">
				<xsl:choose>
					<xsl:when test ="(a:prstDash/@val='solid') or not(a:prstDash/@val)">
						<xsl:attribute name ="draw:stroke">
							<xsl:value-of select ="'solid'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:otherwise>
						<xsl:attribute name ="draw:stroke">
							<xsl:value-of select ="'dash'"/>
						</xsl:attribute>
						<xsl:attribute name ="draw:stroke-dash">
							<xsl:choose>
								<xsl:when test="(a:prstDash/@val='sysDot') and (@cap='rnd')">
									<xsl:value-of select ="'sysDotRound'"/>
								</xsl:when>
								<xsl:when test="a:prstDash/@val='sysDot'">
									<xsl:value-of select ="'sysDot'"/>
								</xsl:when>
								<xsl:when test ="(a:prstDash/@val='sysDash') and (@cap='rnd')">
									<xsl:value-of select ="'sysDashRound'"/>
								</xsl:when>
								<xsl:when test ="a:prstDash/@val='sysDash'">
									<xsl:value-of select ="'sysDash'"/>
								</xsl:when>
								<xsl:when test ="(a:prstDash/@val='dash') and (@cap='rnd')">
									<xsl:value-of select ="'dashRound'"/>
								</xsl:when>
								<xsl:when test ="a:prstDash/@val='dash'">
									<xsl:value-of select ="'dash'"/>
								</xsl:when>
								<xsl:when test ="(a:prstDash/@val='dashDot') and (@cap='rnd')">
									<xsl:value-of select ="'dashDotRound'"/>
								</xsl:when>
								<xsl:when test ="a:prstDash/@val='dashDot'">
									<xsl:value-of select ="'dashDot'"/>
								</xsl:when>
								<xsl:when test ="(a:prstDash/@val='lgDash') and (@cap='rnd')">
									<xsl:value-of select ="'lgDashRound'"/>
								</xsl:when>
								<xsl:when test ="a:prstDash/@val='lgDash'">
									<xsl:value-of select ="'lgDash'"/>
								</xsl:when>
								<xsl:when test ="(a:prstDash/@val='lgDashDot') and (@cap='rnd')">
									<xsl:value-of select ="'lgDashDotRound'"/>
								</xsl:when>
								<xsl:when test ="a:prstDash/@val='lgDashDot'">
									<xsl:value-of select ="'lgDashDot'"/>
								</xsl:when>
								<xsl:when test ="(a:prstDash/@val='lgDashDotDot') and (@cap='rnd')">
									<xsl:value-of select ="'lgDashDotDotRound'"/>
								</xsl:when>
								<xsl:when test ="a:prstDash/@val='lgDashDotDot'">
									<xsl:value-of select ="'lgDashDotDot'"/>
								</xsl:when>
							</xsl:choose>
						</xsl:attribute>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:if>
		</xsl:for-each >

		<!-- Line join property -->
		<xsl:choose>
			<xsl:when test ="p:spPr/a:ln/a:miter">
				<xsl:attribute name ="draw:stroke-linejoin">
					<xsl:value-of select ="'miter'"/>
				</xsl:attribute>
			</xsl:when>
			<xsl:when test ="p:spPr/a:ln/a:bevel">
				<xsl:attribute name ="draw:stroke-linejoin">
					<xsl:value-of select ="'bevel'"/>
				</xsl:attribute>
			</xsl:when>
			<xsl:when test ="p:spPr/a:ln/a:round">
				<xsl:attribute name ="draw:stroke-linejoin">
					<xsl:value-of select ="'round'"/>
				</xsl:attribute>
			</xsl:when>
		</xsl:choose>

		<!-- Line Arrow -->
		<!-- Head End-->
		<xsl:for-each select ="p:spPr/a:ln/a:headEnd">
			<xsl:attribute name ="draw:marker-start">
				<xsl:value-of select ="@type"/>
			</xsl:attribute>
			<xsl:attribute name ="draw:marker-start-width">
				<xsl:call-template name ="getArrowSize">
					<xsl:with-param name ="w" select ="@w" />
					<xsl:with-param name ="len" select ="@len" />
				</xsl:call-template>
			</xsl:attribute>
		</xsl:for-each>

		<!-- Tail End-->
		<xsl:for-each select ="p:spPr/a:ln/a:tailEnd">
			<xsl:attribute name ="draw:marker-end">
				<xsl:value-of select ="@type"/>
			</xsl:attribute>
			<xsl:attribute name ="draw:marker-end-width">
				<xsl:call-template name ="getArrowSize">
					<xsl:with-param name ="w" select ="@w" />
					<xsl:with-param name ="len" select ="@len" />
				</xsl:call-template>
			</xsl:attribute>
		</xsl:for-each>
	</xsl:template>

	<xsl:template name ="getArrowSize">
		<xsl:param name ="w" />
		<xsl:param name ="len" />

		<xsl:choose>
			<xsl:when test ="($w = 'sm') and ($len = 'sm')">
				<xsl:value-of select ="'0.15cm'"/>
			</xsl:when>
			<xsl:when test ="($w = 'sm') and ($len = 'med')">
				<xsl:value-of select ="'0.18cm'"/>
			</xsl:when>
			<xsl:when test ="($w = 'sm') and ($len = 'lg')">
				<xsl:value-of select ="'0.2cm'"/>
			</xsl:when>
			<xsl:when test ="($w = 'med') and ($len = 'sm')">
				<xsl:value-of select ="'0.21cm'" />
			</xsl:when>
			<xsl:when test ="($w = 'med') and ($len = 'lg')">
				<xsl:value-of select ="'0.3cm'" />
			</xsl:when>
			<xsl:when test ="($w = 'lg') and ($len = 'sm')">
				<xsl:value-of select ="'0.31cm'" />
			</xsl:when>
			<xsl:when test ="($w = 'lg') and ($len = 'med')">
				<xsl:value-of select ="'0.35cm'" />
			</xsl:when>
			<xsl:when test ="($w = 'lg') and ($len = 'lg')">
				<xsl:value-of select ="'0.4cm'" />
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select ="'0.25cm'"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- Get text layout for shapes -->
	<xsl:template name ="TextLayout">
		<xsl:for-each select="p:txBody">
			<xsl:attribute name ="draw:textarea-horizontal-align">
				<xsl:choose>
					<!--BugFixed-1709909-->
					<xsl:when test ="p:txBody/a:bodyPr/@anchorCtr= 1">
						<xsl:value-of select ="'center'"/>
					</xsl:when>
					<xsl:when test ="p:txBody/a:bodyPr/@anchorCtr= 0">
						<xsl:value-of select ="'justify'"/>
					</xsl:when>
					<!--Justify alignment-->
					<xsl:otherwise>
						<xsl:value-of select ="'justify'"/>
					</xsl:otherwise>
				</xsl:choose>
				<!-- Center alignment
					<xsl:when test ="(a:p/a:pPr/@algn ='ctr') or (a:bodyPr/@anchorCtr = '1')">
						<xsl:value-of select ="'center'"/>
					</xsl:when>-->
				<!-- Right alignment
					<xsl:when test ="a:p/a:pPr/@algn ='r'">
						<xsl:value-of select ="'right'"/>
					</xsl:when>-->
				<!--Justify alignment
					<xsl:when test ="(a:p/a:pPr/@algn ='just') or (a:bodyPr/@anchorCtr = '0')">
						<xsl:value-of select ="'justify'"/>
					</xsl:when>-->
				<!-- Left alignment
					<xsl:otherwise>
						<xsl:value-of select ="'left'"/>
					</xsl:otherwise>-->

			</xsl:attribute>
			<xsl:attribute name ="draw:textarea-vertical-align">
				<xsl:choose>
					<!-- Middle alignment-->
					<xsl:when test ="a:bodyPr/@anchor ='ctr'">
						<xsl:value-of select ="'middle'"/>
					</xsl:when>
					<!-- Bottom alignment-->
					<xsl:when test ="a:bodyPr/@anchor ='b'">
						<xsl:value-of select ="'bottom'"/>
					</xsl:when>
					<!-- Top alignment -->
					<xsl:otherwise>
						<xsl:value-of select ="'top'"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:attribute>

			<!--fo:padding-->
			<xsl:if test ="a:bodyPr/@lIns or (a:bodyPr/@wrap='square') or (a:p/a:pPr/@fontAlgn='auto')">
				<xsl:attribute name ="fo:padding-left">
					<xsl:call-template name="getPadding">
						<xsl:with-param name="length" select="a:bodyPr/@lIns"/>
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>

			<xsl:if test ="a:bodyPr/@tIns or (a:bodyPr/@wrap='square') or (a:p/a:pPr/@fontAlgn='auto')">
				<xsl:attribute name ="fo:padding-top">
					<xsl:call-template name="getPadding">
						<xsl:with-param name="length" select="a:bodyPr/@tIns"/>
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>

			<xsl:if test ="a:bodyPr/@rIns or (a:bodyPr/@wrap='square') or (a:p/a:pPr/@fontAlgn='auto')">
				<xsl:attribute name ="fo:padding-right">
					<xsl:call-template name="getPadding">
						<xsl:with-param name="length" select="a:bodyPr/@rIns"/>
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>

			<xsl:if test ="a:bodyPr/@bIns or (a:bodyPr/@wrap='square') or (a:p/a:pPr/@fontAlgn='auto')">
				<xsl:attribute name ="fo:padding-bottom">
					<xsl:call-template name="getPadding">
						<xsl:with-param name="length" select="a:bodyPr/@bIns"/>
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>

			<!--Wrap text in shape -->
			<xsl:choose>
				<xsl:when test="(a:bodyPr/@wrap='none')">
					<xsl:attribute name ="fo:wrap-option">
						<xsl:value-of select ="'wrap'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="(a:bodyPr/@wrap='square') or (a:p/a:pPr/@fontAlgn='auto')">
					<xsl:attribute name ="fo:wrap-option">
						<xsl:value-of select ="'no-wrap'"/>
					</xsl:attribute>
				</xsl:when>
			</xsl:choose>

			<xsl:if test ="a:bodyPr/a:spAutoFit">
				<xsl:attribute name="draw:auto-grow-height">
					<xsl:value-of select ="'true'"/>
				</xsl:attribute>
			</xsl:if>
			<xsl:if test ="not(a:bodyPr/a:spAutoFit)">
				<xsl:attribute name="draw:auto-grow-height">
					<xsl:value-of select ="'false'"/>
				</xsl:attribute>
				<xsl:attribute name="draw:auto-grow-width">
					<xsl:value-of select ="'false'"/>
				</xsl:attribute>
			</xsl:if>
		</xsl:for-each>
	</xsl:template>

	<!-- Get text padding-->
	<xsl:template name ="getPadding">
		<xsl:param name ="length" />
		<xsl:choose>
			<xsl:when test ="($length != '')">
				<xsl:call-template name="ConvertEmu">
					<xsl:with-param name="length" select="$length"/>
					<xsl:with-param name="unit">cm</xsl:with-param>
				</xsl:call-template>
			</xsl:when>
			<xsl:when test ="a:bodyPr/@wrap='square'">
				<xsl:value-of select ="'0cm'"/>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<!--Text direction-->
	<xsl:template name ="getTextDirection">
		<xsl:param name ="vert" />
		<xsl:choose>
			<xsl:when test ="$vert = 'vert'">
				<xsl:value-of select ="'tb-rl'"/>
			</xsl:when>
			<!--<xsl:when test ="$vert = 'vert270'">
					<xsl:value-of select ="'tb-lr'"/>
				</xsl:when>
				<xsl:when test ="$vert = 'wordArtVert'">
					<xsl:value-of select ="'tb'"/>
				</xsl:when>
				<xsl:when test ="$vert = 'eaVert'">
					<xsl:value-of select ="'lr'" />
				</xsl:when>
				<xsl:when test ="$vert = 'mongolianVert'">
					<xsl:value-of select ="'lr'" />
				</xsl:when>
				<xsl:when test ="$vert = 'wordArtVertRtl'">
					<xsl:value-of select ="'lr'" />
				</xsl:when>-->
		</xsl:choose>
	</xsl:template>
</xsl:stylesheet >
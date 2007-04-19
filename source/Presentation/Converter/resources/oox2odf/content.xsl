<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" 
xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
xmlns:xlink="http://www.w3.org/1999/xlink"
xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
xmlns:rels="http://schemas.openxmlformats.org/package/2006/relationships"
exclude-result-prefixes="p a r xlink number">
	<xsl:import href="common.xsl"/>
	<xsl:import href="shapes_reverse.xsl"/>

	<xsl:strip-space elements="*"/>

	<!--main document-->
	<xsl:template name="content">
		<office:document-content>
			<office:automatic-styles>
				<!-- automatic styles for document body -->
				<xsl:call-template name="InsertStyles"/>
				<!-- Graphic properties for shapes -->
				<xsl:call-template name ="InsertStylesForGraphicProperties"/>
			</office:automatic-styles>
			<office:body>
				<office:presentation>
					<xsl:call-template name="InsertPresentationFooter"/>
					<xsl:call-template name="InsertDrawingPage"/>
					<!-- Added by Lohith A R : Custom Slide Show -->
					<xsl:call-template name="InsertPresentationSettings"/>
				</office:presentation>
			</office:body>
		</office:document-content>
	</xsl:template>

	<!--  generates automatic styles for paragraphs  how does it exactly work ?? -->
	<xsl:variable name ="flgFooter">
		<xsl:for-each select ="document('ppt/slides/slide1.xml')/p:sld/p:cSld/p:spTree/p:sp">
			<xsl:if test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'ftr')]" >
				<xsl:value-of select ="p:txBody/a:p/a:r/a:t"/>
			</xsl:if>
			<!--<xsl:value-of select ="."/>-->
		</xsl:for-each >
	</xsl:variable>
	<xsl:template name="InsertStyles">
		<!-- page Properties-->
		<xsl:for-each select ="document('ppt/presentation.xml')/p:presentation/p:sldIdLst/p:sldId">
			<xsl:variable name ="pageSlide">
				<xsl:value-of select ="concat(concat('ppt/slides/slide',position()),'.xml')"/>
			</xsl:variable>
			<!-- Page settings like footer date slide number visible/Invisible-->
			<style:style  style:family="drawing-page">
				<xsl:attribute name ="style:name" >
					<xsl:value-of select ="concat('dp',position())"/>
				</xsl:attribute>
				<style:drawing-page-properties>
					<xsl:attribute name ="presentation:background-visible" >
						<xsl:value-of select ="'true'"/>
					</xsl:attribute>
					<xsl:attribute name ="presentation:background-objects-visible" >
						<xsl:value-of select ="'true'"/>
					</xsl:attribute>
					<xsl:attribute name ="presentation:display-footer" >
						<xsl:value-of select ="'true'"/>
					</xsl:attribute>
					<xsl:attribute name ="presentation:display-page-number" >
						<xsl:if test ="document($pageSlide)/p:sld/p:cSld/p:spTree/p:sp
						/p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'sldNum')]">
							<xsl:value-of select ="'true'"/>
						</xsl:if >
					</xsl:attribute>
					<xsl:attribute name ="presentation:display-date-time" >
						<xsl:value-of select ="'true'"/>
					</xsl:attribute>
				</style:drawing-page-properties>
			</style:style>
		</xsl:for-each >


		<style:style style:name="pr1" style:family="presentation" style:parent-style-name="Default-notes">
			<style:graphic-properties >
				<xsl:attribute name ="draw:fill-color" >
					<xsl:value-of select ="'#ffffff'"/>
				</xsl:attribute>
				<xsl:attribute name ="draw:auto-grow-height" >
					<xsl:value-of select ="'false'"/>
				</xsl:attribute>
				<xsl:attribute name ="fo:min-height" >
					<xsl:value-of select ="'12.573cm'"/>
				</xsl:attribute>
			</style:graphic-properties>
		</style:style>
		<xsl:call-template name ="GetStylesFromSlide"/>
	</xsl:template>
	<xsl:template name ="InsertPresentationFooter" >
		<xsl:for-each select ="document('ppt/presentation.xml')/p:presentation/p:sldIdLst/p:sldId">
			<xsl:variable name ="footerSlide">
				<xsl:value-of select ="concat(concat('ppt/slides/slide',position()),'.xml')"/>
			</xsl:variable>
			<xsl:variable name ="footerInd">
				<xsl:value-of select ="concat('ftr',position())"/>
			</xsl:variable>
			<xsl:variable name ="dateInd">
				<xsl:value-of select ="concat('dtd',position())"/>
			</xsl:variable>
			<xsl:for-each select ="document($footerSlide)/p:sld/p:cSld/p:spTree/p:sp">
				<!--concat(concat('slide',position()),'.xml')-->
				<xsl:choose >
					<xsl:when test ="not(p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'dt') or contains(.,'ftr')])">
						<!-- Do nothing-->
					</xsl:when>
					<xsl:when test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'ftr')]">
						<presentation:footer-decl >
							<xsl:attribute name ="presentation:name">
								<xsl:value-of select ="$footerInd"/>
							</xsl:attribute >
							<xsl:value-of select ="p:txBody/a:p/a:r/a:t"/>
						</presentation:footer-decl>
					</xsl:when>
					<xsl:when test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'dt')]">
						<presentation:date-time-decl style:data-style-name="D3">
							<xsl:attribute name ="presentation:name">
								<xsl:value-of select ="$dateInd"/>
							</xsl:attribute>
							<xsl:attribute name ="presentation:source">
								<xsl:for-each select =".">
									<xsl:if test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'dt')]" >
										<xsl:if test ="p:txBody/a:p/a:fld">
											<xsl:value-of select ="'current-date'"/>
										</xsl:if>
										<xsl:if test ="not(p:txBody/a:p/a:fld)">
											<xsl:value-of select ="'fixed'"/>
										</xsl:if>
									</xsl:if>
								</xsl:for-each >
							</xsl:attribute>
							<xsl:for-each select =".">
								<xsl:if test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'dt')]" >
									<xsl:if test ="p:txBody/a:p/a:fld">
										<xsl:value-of select ="p:txBody/a:p/a:fld/a:t"/>
									</xsl:if>
									<xsl:if test ="not(p:txBody/a:p/a:fld)">
										<xsl:value-of select ="p:txBody/a:p/a:r/a:t"/>
									</xsl:if>
								</xsl:if>
							</xsl:for-each >
						</presentation:date-time-decl>
					</xsl:when >
				</xsl:choose>
			</xsl:for-each>
		</xsl:for-each>
	</xsl:template>
	<xsl:template name ="InsertDrawingPage"  >
		<xsl:for-each select ="document('ppt/presentation.xml')/p:presentation/p:sldIdLst/p:sldId">
			<xsl:value-of select ="."/>
			<draw:page  
				draw:master-page-name="Default">
				<xsl:attribute name ="draw:style-name">
					<xsl:value-of select ="concat('dp',position())"/>
				</xsl:attribute>
				<xsl:attribute name ="draw:name">
					<xsl:value-of select ="concat('Page',position())"/>
				</xsl:attribute>
				<xsl:attribute name ="presentation:use-footer-name">
					<xsl:value-of select ="concat('ftr',position())"/>
				</xsl:attribute>
				<xsl:attribute name ="presentation:use-date-time-name">
					<xsl:value-of select ="concat('dtd',position())"/>
				</xsl:attribute>
				<xsl:call-template name ="DrawFrames">
					<xsl:with-param name ="SlideFile" select ="concat(concat('slide',position()),'.xml')" />
				</xsl:call-template>
			</draw:page >
		</xsl:for-each>
	</xsl:template>
	<xsl:template name ="DrawFrames">
		<xsl:param name ="SlideFile"/>
		<xsl:for-each  select="document('ppt/_rels/presentation.xml.rels')">
			<!--<xsl:for-each select ="./child::node()[1]">
				<xsl:value-of select ="."/>
			</xsl:for-each>-->
			<xsl:variable name ="slideNo">
				<xsl:value-of select ="concat('slides/',$SlideFile)"/>
			</xsl:variable>
			<xsl:variable name ="slideRel">
				<xsl:value-of select ="concat(concat('ppt/slides/_rels/',$SlideFile),'.rels')"/>
			</xsl:variable>
			<xsl:variable name ="slideLayotRel">
				<xsl:value-of select ="'ppt/slideMasters/_rels'"/>
			</xsl:variable>
			<!-- Slide Files Loop Main Loop-->
			<xsl:for-each select ="document(concat('ppt/',$slideNo))/p:sld/p:cSld/p:spTree/p:sp">
				<xsl:variable name ="LayoutName">
					<xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph/@type">
						<xsl:value-of select ="."/>
					</xsl:for-each>
				</xsl:variable>
				<xsl:variable name ="ParaId">
					<xsl:value-of select ="concat(substring($SlideFile,1,string-length($SlideFile)-4),concat('PARA',position()))"/>
				</xsl:variable>
				<xsl:choose >

					<xsl:when test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'dt') 
					or contains(.,'ftr') or contains(.,'sldNum')]">
						<!-- Do nothing-->
						<!-- These will be covered in footer and date time -->
					</xsl:when>
					<!-- Code for shapes start-->
					<xsl:when test ="p:style or p:nvSpPr/p:cNvPr/@name[contains(., 'TextBox')]">
						<xsl:variable  name ="GraphicId">
							<xsl:value-of select ="concat('s',substring($SlideFile,6,string-length($SlideFile)-9) ,concat('gr',position()))"/>
						</xsl:variable>
						<xsl:call-template name ="shapes">
							<xsl:with-param name="GraphicId" select ="$GraphicId"/>
							<xsl:with-param name ="ParaId" select="$ParaId" />
						</xsl:call-template>
					</xsl:when>
					<!-- Code for shapes end-->

					<xsl:when test ="p:spPr/a:xfrm/a:off" >
						<draw:frame presentation:style-name="pr1" draw:layer="layout" 																
						    presentation:user-transformed="true">
							<xsl:attribute name ="presentation:class">
								<xsl:call-template name ="LayoutType">
									<xsl:with-param name ="LayoutStyle">
										<xsl:value-of select ="$LayoutName"/>
									</xsl:with-param>
								</xsl:call-template >
							</xsl:attribute>
							<xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cx]">
								<xsl:attribute name ="svg:width">
									<xsl:call-template name="ConvertEmu">
										<xsl:with-param name="length" select="@cx"/>
										<xsl:with-param name="unit">cm</xsl:with-param>
									</xsl:call-template>
								</xsl:attribute>
							</xsl:for-each>
							<xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cy]">
								<xsl:attribute name ="svg:height">
									<xsl:call-template name="ConvertEmu">
										<xsl:with-param name="length" select="@cy"/>
										<xsl:with-param name="unit">cm</xsl:with-param>
									</xsl:call-template>
								</xsl:attribute>
							</xsl:for-each>
							<xsl:for-each select ="p:spPr/a:xfrm/a:off[@x]">
								<xsl:attribute name ="svg:x">
									<xsl:call-template name="ConvertEmu">
										<xsl:with-param name="length" select="@x"/>
										<xsl:with-param name="unit">cm</xsl:with-param>
									</xsl:call-template>
								</xsl:attribute>
							</xsl:for-each>
							<xsl:for-each select ="p:spPr/a:xfrm/a:off[@y]">
								<xsl:attribute name ="svg:y">
									<xsl:call-template name="ConvertEmu">
										<xsl:with-param name="length" select="@y"/>
										<xsl:with-param name="unit">cm</xsl:with-param>
									</xsl:call-template>
								</xsl:attribute>
							</xsl:for-each>
							<draw:text-box>
								<xsl:for-each select ="p:txBody/a:p">
									<text:p >
										<xsl:attribute name ="text:style-name">
											<xsl:value-of select ="concat($ParaId,position())"/>
										</xsl:attribute>
										<xsl:for-each select ="a:r">
											<text:span text:style-sname="{generate-id()}">
												<xsl:value-of select ="a:t"/>
											</text:span>
										</xsl:for-each>
									</text:p>
								</xsl:for-each>
							</draw:text-box >
						</draw:frame >
					</xsl:when>
					<!-- If Slide layout files have the frame properties-->
					<xsl:when test ="not(p:nvSpPr/p:nvPr/p:ph/@type[contains(.,$LayoutName)]
						and p:spPr/a:xfrm/a:off)" >
						<xsl:variable name ="frameName">
							<xsl:if test ="p:nvSpPr/p:nvPr/p:ph/@type">
								<xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph/@type">
									<xsl:value-of select ="."/>
								</xsl:for-each>
							</xsl:if>
							<xsl:if test="not(p:nvSpPr/p:nvPr/p:ph/@type)">
								<xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph/@idx">
									<xsl:value-of select ="."/>
								</xsl:for-each>
							</xsl:if >
						</xsl:variable>
						<xsl:variable name ="FrameIdx">
							<xsl:if test ="p:nvSpPr/p:nvPr/p:ph/@idx">
								<xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph[@idx]">
									<xsl:value-of select ="@idx"/>
								</xsl:for-each>
							</xsl:if>
						</xsl:variable>
						<xsl:variable name ="TextSpanNode">
							<xsl:for-each select ="p:txBody/a:p">
								<text:p>
									<xsl:attribute name ="text:style-name">
										<xsl:value-of select ="concat($ParaId,position())"/>
									</xsl:attribute >
									<xsl:for-each select ="a:r">
										<text:span text:style-name="{generate-id()}">
											<xsl:value-of select ="a:t"/>
										</text:span>
									</xsl:for-each>
								</text:p>
							</xsl:for-each>
						</xsl:variable>
						<xsl:variable name ="LayoutFileNo">
							<xsl:for-each select ="document($slideRel)//node()/@Target[contains(.,'slideLayouts')]">
								<xsl:value-of select ="concat('ppt',substring(.,3))"/>
							</xsl:for-each>
						</xsl:variable>
						<!-- Slide Layout Files Loop Second Loop-->
						<xsl:for-each select ="document($LayoutFileNo)/p:sldLayout/p:cSld/p:spTree/p:sp">
							<xsl:variable name ="SlFrameName">
								<xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph/@type">
									<xsl:value-of select ="."/>
								</xsl:for-each>
								<xsl:if test="not(p:nvSpPr/p:nvPr/p:ph/@type)">
									<xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph/@idx">
										<xsl:value-of select ="."/>
									</xsl:for-each>
								</xsl:if >
							</xsl:variable>
							<xsl:variable name ="SlFrameNameInd">
								<xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph/@idx">
									<xsl:value-of select ="."/>
								</xsl:for-each>
							</xsl:variable>
							<xsl:choose >
								<xsl:when test ="not($SlFrameName = $frameName or $SlFrameNameInd =$FrameIdx )">
									<!-- Do nothing-or $SlFrameNameInd=$FrameIdx
									 These will be covered in footer and date time -->
								</xsl:when>
								<xsl:when test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'dt') 
								or contains(.,'ftr') or contains(.,'sldNum')]">
									<!-- Do nothing-->
									<!-- These will be covered in footer and date time -->
								</xsl:when>

								<xsl:when test ="(p:nvSpPr/p:nvPr/p:ph/@type[contains(.,$frameName)] or 
										p:nvSpPr/p:nvPr/p:ph/@idx[contains(.,$FrameIdx)])
										and p:spPr/a:xfrm/a:off" >
									<draw:frame presentation:style-name="pr1" draw:layer="layout" 																			
										presentation:user-transformed="true">
										<xsl:attribute name ="presentation:class">
											<xsl:call-template name ="LayoutType">
												<xsl:with-param name ="LayoutStyle">
													<xsl:value-of select ="$LayoutName"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:attribute>
										<xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cx]">
											<xsl:attribute name ="svg:width">
												<xsl:call-template name="ConvertEmu">
													<xsl:with-param name="length" select="@cx"/>
													<xsl:with-param name="unit">cm</xsl:with-param>
												</xsl:call-template>
											</xsl:attribute>
										</xsl:for-each>
										<xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cy]">
											<xsl:attribute name ="svg:height">
												<xsl:call-template name="ConvertEmu">
													<xsl:with-param name="length" select="@cy"/>
													<xsl:with-param name="unit">cm</xsl:with-param>
												</xsl:call-template>
											</xsl:attribute>
										</xsl:for-each>
										<xsl:for-each select ="p:spPr/a:xfrm/a:off[@x]">
											<xsl:attribute name ="svg:x">
												<xsl:call-template name="ConvertEmu">
													<xsl:with-param name="length" select="@x"/>
													<xsl:with-param name="unit">cm</xsl:with-param>
												</xsl:call-template>
											</xsl:attribute>
										</xsl:for-each>
										<xsl:for-each select ="p:spPr/a:xfrm/a:off[@y]">
											<xsl:attribute name ="svg:y">
												<xsl:call-template name="ConvertEmu">
													<xsl:with-param name="length" select="@y"/>
													<xsl:with-param name="unit">cm</xsl:with-param>
												</xsl:call-template>
											</xsl:attribute>
										</xsl:for-each>
										<draw:text-box>
											<xsl:copy-of select ="$TextSpanNode"/>
										</draw:text-box >
									</draw:frame >
								</xsl:when >
								<xsl:when test ="not((p:nvSpPr/p:nvPr/p:ph/@type[contains(.,$frameName)] or 
									p:nvSpPr/p:nvPr/p:ph/@idx[contains(.,$FrameIdx)])
									and p:spPr/a:xfrm/a:off)">
									<xsl:variable name ="LayoutRels">
										<xsl:for-each select ="document(concat(concat('ppt/slideLayouts/_rels',substring($LayoutFileNo,17)),'.rels'))
											//node()/@Target[contains(.,'slideMasters')]">
											<xsl:value-of select ="."/>
										</xsl:for-each>
									</xsl:variable>
									<xsl:variable name ="MasterFileName">
										<xsl:value-of select ="concat('ppt/slideMasters',substring($LayoutRels,16))"/>
									</xsl:variable>
									<!-- Slide Master Files Loop  Third Loop-->
									<xsl:for-each select ="document($MasterFileName)/p:sldMaster/p:cSld/p:spTree/p:sp">
										<xsl:variable name ="MstrFrameName">
											<xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph/@type">
												<xsl:value-of select ="."/>
											</xsl:for-each>
											<xsl:if test="not(p:nvSpPr/p:nvPr/p:ph/@type)">
												<xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph/@idx">
													<xsl:value-of select ="."/>
												</xsl:for-each>
											</xsl:if >
										</xsl:variable>
										<xsl:variable name ="MstrFrameInd">
											<xsl:for-each select ="p:nvSpPr/p:nvPr/p:ph/@idx">
												<xsl:value-of select ="."/>
											</xsl:for-each>
										</xsl:variable>
										<xsl:if test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,$frameName)]">
											<xsl:value-of select ="p:nvSpPr/p:nvPr/p:ph"/>
										</xsl:if>
										<xsl:choose >
											<xsl:when test ="not($MstrFrameName = $frameName
										or $MstrFrameInd =$FrameIdx)">
												<!-- Do nothing-->
												<!-- These will be covered in footer and date time -->
											</xsl:when>
											<xsl:when test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,'dt') or 
														 contains(.,'ftr') or contains(.,'sldNum')]">
												<!-- Do nothing-->
												<!-- These will be covered in footer and date time -->
											</xsl:when>
											<xsl:when test ="p:nvSpPr/p:nvPr/p:ph/@type[contains(.,$frameName)] 
											or p:nvSpPr/p:nvPr/p:ph/@idx[contains(.,$FrameIdx)]">
												<draw:frame presentation:style-name="pr1" draw:layer="layout"
												presentation:user-transformed="true">
													<xsl:attribute name ="presentation:class">
														<xsl:call-template name ="LayoutType">
															<xsl:with-param name ="LayoutStyle">
																<xsl:value-of select ="$LayoutName"/>
															</xsl:with-param>
														</xsl:call-template >
													</xsl:attribute>
													<xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cx]">
														<xsl:attribute name ="svg:width">
															<xsl:call-template name="ConvertEmu">
																<xsl:with-param name="length" select="@cx"/>
																<xsl:with-param name="unit">cm</xsl:with-param>
															</xsl:call-template>
														</xsl:attribute>
													</xsl:for-each>
													<xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cy]">
														<xsl:attribute name ="svg:height">
															<xsl:call-template name="ConvertEmu">
																<xsl:with-param name="length" select="@cy"/>
																<xsl:with-param name="unit">cm</xsl:with-param>
															</xsl:call-template>
														</xsl:attribute>
													</xsl:for-each>
													<xsl:for-each select ="p:spPr/a:xfrm/a:off[@x]">
														<xsl:attribute name ="svg:x">
															<xsl:call-template name="ConvertEmu">
																<xsl:with-param name="length" select="@x"/>
																<xsl:with-param name="unit">cm</xsl:with-param>
															</xsl:call-template>
														</xsl:attribute>
													</xsl:for-each>
													<xsl:for-each select ="p:spPr/a:xfrm/a:off[@y]">
														<xsl:attribute name ="svg:y">
															<xsl:call-template name="ConvertEmu">
																<xsl:with-param name="length" select="@y"/>
																<xsl:with-param name="unit">cm</xsl:with-param>
															</xsl:call-template>
														</xsl:attribute>
													</xsl:for-each>
													<draw:text-box>
														<xsl:copy-of select ="$TextSpanNode"/>
													</draw:text-box >
												</draw:frame >
											</xsl:when>
										</xsl:choose>
									</xsl:for-each >
									<!-- Exit Slide Master Files Loop  Third Loop-->
								</xsl:when>
							</xsl:choose>
						</xsl:for-each>
						<!-- exit Slide layout loop second Loop-->
					</xsl:when>
				</xsl:choose >
			</xsl:for-each>
			<!-- exit Slide loop Main Loop-->

			<xsl:for-each select ="document(concat('ppt/',$slideNo))/p:sld/p:cSld/p:spTree/p:cxnSp">
				<!-- Code for shapes start-->
				<xsl:if test ="p:style">
					<xsl:variable  name ="GraphicId">
						<xsl:value-of select ="concat('s',substring($SlideFile,6,string-length($SlideFile)-9) ,concat('gr',position()))"/>
					</xsl:variable>
					<xsl:call-template name ="shapes">
						<xsl:with-param name="GraphicId" select ="$GraphicId"/>
					</xsl:call-template>
				</xsl:if>
				<!-- Code for shapes end-->
			</xsl:for-each>
		</xsl:for-each >
	</xsl:template>
	<xsl:template name ="FrameProperties" match ="p:spPr/a:xfrm">
		<xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cx]">
			<xsl:attribute name ="svg:width">
				<xsl:call-template name="ConvertEmu">
					<xsl:with-param name="length" select="@cx"/>
					<xsl:with-param name="unit">cm</xsl:with-param>
				</xsl:call-template>
			</xsl:attribute>
		</xsl:for-each>
		<xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cy]">
			<xsl:attribute name ="svg:height">
				<xsl:call-template name="ConvertEmu">
					<xsl:with-param name="length" select="@cy"/>
					<xsl:with-param name="unit">cm</xsl:with-param>
				</xsl:call-template>
			</xsl:attribute>
		</xsl:for-each>
		<xsl:for-each select ="a:off[@x]">
			<xsl:attribute name ="svg:x">
				<xsl:call-template name="ConvertEmu">
					<xsl:with-param name="length" select="@x"/>
					<xsl:with-param name="unit">cm</xsl:with-param>
				</xsl:call-template>
			</xsl:attribute>
		</xsl:for-each>
		<xsl:for-each select ="a:off[@y]">
			<xsl:attribute name ="svg:y">
				<xsl:call-template name="ConvertEmu">
					<xsl:with-param name="length" select="@y"/>
					<xsl:with-param name="unit">cm</xsl:with-param>
				</xsl:call-template>
			</xsl:attribute>
		</xsl:for-each>
	</xsl:template>
	<xsl:template name ="WriteTextSpan">
		<xsl:param name ="file"/>
		<draw:text-box>
			<xsl:for-each select ="document(concat('ppt/',$file))/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p">
				<text:p text:style-name="P1">
					<xsl:for-each select ="a:r">
						<text:span text:style-name="T1">
							<xsl:value-of select ="a:t"/>
						</text:span>
					</xsl:for-each>
				</text:p>
			</xsl:for-each>
		</draw:text-box >
	</xsl:template>

	<xsl:template name ="GetStylesFromSlide" >
		<!--@ Default Font Name from PPTX -->
		<xsl:variable name ="DefFont">
			<xsl:for-each select ="document('ppt/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme
						/a:majorFont/a:latin/@typeface">
				<xsl:value-of select ="."/>
			</xsl:for-each>
		</xsl:variable>

		<!--@Modified font Names -->
		<xsl:for-each select ="document('ppt/presentation.xml')/p:presentation/p:sldIdLst/p:sldId">
			<xsl:variable name ="SlideNumber">
				<xsl:value-of  select ="concat('slide',position())" />
			</xsl:variable>
			<xsl:variable name ="SlideId">
				<xsl:value-of  select ="concat(concat('slide',position()),'.xml')" />
			</xsl:variable>
			<xsl:for-each select ="document(concat('ppt/slides/',$SlideId))/p:sld/p:cSld/p:spTree/p:sp/p:txBody">
				<xsl:variable name ="ParaId">
					<xsl:value-of select ="concat($SlideNumber,concat('PARA',position()))"/>
				</xsl:variable>
				<xsl:for-each select ="a:p">
					<style:style style:family="paragraph">
						<xsl:attribute name ="style:name">
							<xsl:value-of select ="concat($ParaId,position())"/>
						</xsl:attribute >
						<style:paragraph-properties >
							<xsl:attribute name ="fo:text-align">
								<xsl:choose>
									<!-- Center Alignment-->
									<xsl:when test ="a:pPr/@algn ='ctr'">
										<xsl:value-of select ="'center'"/>
									</xsl:when>
									<!-- Right Alignment-->
									<xsl:when test ="a:pPr/@algn ='r'">
										<xsl:value-of select ="'end'"/>
									</xsl:when>
									<!-- Left Alignment-->
									<xsl:when test ="a:pPr/@algn ='l' or a:pPr/@algn">
										<xsl:value-of select ="'start'"/>
									</xsl:when>
								</xsl:choose>
							</xsl:attribute>
							<!-- Convert Laeft margin of the paragraph-->
							<xsl:attribute name ="fo:margin-left">
								<xsl:value-of select="concat(format-number(a:pPr/@marL div 360000, '#.##'), 'cm')"/>
							</xsl:attribute>
							<xsl:attribute name ="fo:text-indent">
								<xsl:value-of select="concat(format-number(a:pPr/@indent div 360000, '#.##'), 'cm')"/>
							</xsl:attribute>
							<xsl:attribute name ="fo:margin-top">
								<xsl:value-of select="concat(format-number(a:pPr/a:spcBef/a:spcPts/@val div 1000, '#.##'), 'cm')"/>
							</xsl:attribute>
							<xsl:attribute name ="fo:margin-bottom">
								<xsl:value-of select="concat(format-number(a:pPr/a:spcAft/a:spcPts/@val div 1000, '#.##'), 'cm')"/>
							</xsl:attribute>
							<xsl:if test ="a:pPr/a:lnSpc/a:spcPct/@val">
								<xsl:attribute name="fo:line-height">
									<xsl:value-of select="concat(format-number(a:pPr/a:lnSpc/a:spcPct/@val div 1000, '#.##'), '%')"/>
								</xsl:attribute>
							</xsl:if >
						</style:paragraph-properties >
					</style:style>
					<xsl:for-each select ="a:r" >
						<style:style style:name="{generate-id()}"  style:family="text">
							<style:text-properties>
								<xsl:attribute name ="fo:font-family">
									<xsl:if test ="a:rPr/a:latin/@typeface">
										<xsl:for-each select ="a:rPr/a:latin/@typeface">
											<xsl:value-of select ="."/>
										</xsl:for-each>
									</xsl:if>
									<xsl:if test ="not(a:rPr/a:latin/@typeface)">
										<xsl:value-of select ="$DefFont"/>
									</xsl:if>
								</xsl:attribute>
								<xsl:attribute name ="style:font-family-generic"	>
									<xsl:value-of select ="'roman'"/>
								</xsl:attribute>
								<xsl:attribute name ="style:font-pitch"	>
									<xsl:value-of select ="'variable'"/>
								</xsl:attribute>
								<xsl:if test ="a:rPr/@sz">
									<xsl:attribute name ="fo:font-size"	>
										<xsl:for-each select ="a:rPr/@sz">
											<xsl:value-of select ="concat(format-number(. div 100,'#.##'), 'pt')"/>
										</xsl:for-each>
									</xsl:attribute>
								</xsl:if>
								<xsl:attribute name ="fo:font-weight">
									<!-- Bold Property-->
									<xsl:if test ="a:rPr/@b">
										<xsl:value-of select ="'bold'"/>
									</xsl:if >
									<xsl:if test ="not(a:rPr/@b)">
										<xsl:value-of select ="'normal'"/>
									</xsl:if >
								</xsl:attribute>
								<!-- Color -->
								<xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
								<xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
								<xsl:if test ="a:rPr/a:solidFill/a:srgbClr/@val">
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="translate(concat('#',a:rPr/a:solidFill/a:srgbClr/@val),$ucletters,$lcletters)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="a:rPr/a:solidFill/a:schemeClr/@val">
									<xsl:variable name ="ThemeColor">
										<xsl:value-of select ="concat('a:',a:rPr/a:solidFill/a:schemeClr/@val)"/>
									</xsl:variable>
									<xsl:attribute name ="fo:color">
										<xsl:for-each  select ="document('ppt/theme/theme1.xml')
														/a:theme/a:themeElements/a:clrScheme/node()[name()=$ThemeColor]" >
											<xsl:value-of select ="translate(concat('#',a:srgbClr/@val),$ucletters,$lcletters)"/>
										</xsl:for-each>
									</xsl:attribute>
								</xsl:if>
								<!-- Italic-->
								<xsl:if test ="a:rPr/@i">
									<xsl:attribute name ="fo:font-style">
										<xsl:value-of select ="'italic'"/>
									</xsl:attribute>
								</xsl:if>
								<!-- style:text-underline-style
								style:text-underline-style="solid" style:text-underline-width="auto"-->
								<xsl:if test ="a:rPr/@u">
									<xsl:attribute name ="style:text-underline-style">
										<xsl:value-of select ="'solid'"/>
									</xsl:attribute>
									<xsl:attribute name ="style:text-underline-width">
										<xsl:value-of select ="'auto'"/>
									</xsl:attribute>
								</xsl:if>
								<!-- strike style:text-line-through-style-->
								<xsl:if test ="a:rPr/@strike">
									<xsl:attribute name ="style:text-line-through-style">
										<xsl:value-of select ="'solid'"/>
									</xsl:attribute>
								</xsl:if>
								<!-- Charector Spacing fo:letter-spacing-->
								<xsl:if test ="a:rPr/@spc">
									<xsl:attribute name ="fo:letter-spacing">
										<xsl:value-of select ="a:rPr/@spc"/>
									</xsl:attribute>
								</xsl:if>
								<!--Shadow fo:text-shadow-->
								<xsl:if test ="a:rPr/a:effectLst/a:outerShdw">
									<xsl:attribute name ="fo:text-shadow">
										<xsl:value-of select ="'1pt 1pt'"/>
									</xsl:attribute>
								</xsl:if>
								<!--Kerning true or false -->
								<xsl:if test ="a:rPr/@kern">
									<xsl:attribute name ="style:letter-kerning">
										<xsl:value-of select ="'true'"/>
									</xsl:attribute>
								</xsl:if>
							</style:text-properties>
						</style:style>
					</xsl:for-each >
				</xsl:for-each>
			</xsl:for-each>
		</xsl:for-each>
	</xsl:template>
	<xsl:template name="GenerateId">
		<xsl:param name="node"/>
		<xsl:param name="nodetype"/>
		<xsl:variable name="positionInGroup">
			<xsl:for-each select="key(concat($nodetype,'s'), '')">
				<xsl:if test="generate-id($node) = generate-id(.)">
					<xsl:value-of select="position() + 1"/>
				</xsl:if>
			</xsl:for-each>
		</xsl:variable>
		<xsl:value-of select="$positionInGroup"/>
	</xsl:template>


	<xsl:template name ="LayoutType">
		<xsl:param name ="LayoutStyle"/>
		<xsl:choose >
			<xsl:when test ="$LayoutStyle='ctrTitle'">
				<xsl:value-of select ="'title'"/>
			</xsl:when>
			<xsl:when test ="$LayoutStyle='subTitle'">
				<xsl:value-of select ="'subTitle'"/>
			</xsl:when>
			<xsl:when test ="$LayoutStyle=''">
				<xsl:value-of select ="'outline'"/>
			</xsl:when>
			<xsl:when test ="$LayoutStyle='title'">
				<xsl:value-of select ="'title'"/>
			</xsl:when>
			<xsl:otherwise >
				<xsl:value-of select ="'outline'"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- Added by Lohith A R : Custom Slide Show -->
	<xsl:template name="InsertPresentationSettings">
		<presentation:settings>
			<xsl:attribute name ="presentation:show">
				<xsl:value-of select="document('ppt/presentation.xml')/p:presentation/p:custShowLst/p:custShow/@name"/>
			</xsl:attribute>
			<xsl:for-each select="document('ppt/presentation.xml')/p:presentation/p:custShowLst/p:custShow">
				<presentation:show>
					<xsl:attribute name ="presentation:name"	>
						<xsl:value-of select ="@name"/>
					</xsl:attribute>
					<xsl:attribute name ="presentation:pages"	>
						<xsl:for-each select ="p:sldLst/p:sld">
							<xsl:variable name="SlideIdVal">
								<xsl:value-of select ="@r:id"/>
							</xsl:variable>
							<xsl:variable name="SlideVal">
								<xsl:value-of select="document('ppt/_rels/presentation.xml.rels')/rels:Relationships/rels:Relationship[@Id=$SlideIdVal]/@Target"/>
							</xsl:variable>
							<xsl:variable name="CoustomSlideList">
								<xsl:value-of select="concat('Page',substring-before(substring-after($SlideVal,'slides/slide'),'.xml'),',')"/>
							</xsl:variable>
							<!--<xsl:variable name="CoustomSlideListForOdp">
                  <xsl:value-of select="substring($CoustomSlideList,1,string-length($CoustomSlideList)-1)"/>
                </xsl:variable>-->
							<xsl:value-of select="$CoustomSlideList"/>
						</xsl:for-each>
					</xsl:attribute>
				</presentation:show>
			</xsl:for-each>
		</presentation:settings>
	</xsl:template>
	<!-- Added by Lohith A R : Custom Slide Show -->
	<xsl:template name="ConvertPoints">
		<xsl:param name="length"/>
		<xsl:param name="unit"/>
		<xsl:variable name="lengthVal">
			<xsl:choose>
				<xsl:when test="contains($length,'pt')">
					<xsl:value-of select="substring-before($length,'pt')"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="$length"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="$lengthVal='0' or $lengthVal=''">
				<xsl:value-of select="concat(0, $unit)"/>
			</xsl:when>
			<xsl:when test="$unit = 'cm'">
				<xsl:value-of select="concat(format-number($lengthVal * 2.54 div 72,'#.###'),'cm')"/>
			</xsl:when>
			<xsl:when test="$unit = 'mm'">
				<xsl:value-of select="concat(format-number($lengthVal * 25.4 div 72,'#.###'),'mm')"/>
			</xsl:when>
			<xsl:when test="$unit = 'in'">
				<xsl:value-of select="concat(format-number($lengthVal div 72,'#.###'),'in')"/>
			</xsl:when>
			<xsl:when test="$unit = 'pt'">
				<xsl:value-of select="concat($lengthVal,'pt')"/>
			</xsl:when>
			<xsl:when test="$unit = 'pica'">
				<xsl:value-of select="concat(format-number($lengthVal div 12,'#.###'),'pica')"/>
			</xsl:when>
			<xsl:when test="$unit = 'dpt'">
				<xsl:value-of select="concat($lengthVal,'dpt')"/>
			</xsl:when>
			<xsl:when test="$unit = 'px'">
				<xsl:value-of select="concat(format-number($lengthVal * 96.19 div 72,'#.###'),'px')"/>
			</xsl:when>
			<xsl:when test="not($lengthVal)">
				<xsl:value-of select="concat(0,'cm')"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$lengthVal"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
</xsl:stylesheet>
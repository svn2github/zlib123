<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
<!-- 
 Pradeep Nemadi
-->
<xsl:stylesheet version="1.0" 
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"   
  xmlns:odf="urn:odf"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"  
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" 
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:page="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" 
  xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  exclude-result-prefixes="odf style text number draw page">
	<xsl:import href ="common.xsl"/>
	<xsl:import href ="shapes_direct.xsl"/>
	<xsl:template name ="slides" match ="/office:document-content/office:body/office:presentation/draw:page" mode="slide">
		<xsl:param name ="pageNo"/>
		<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
			xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
			xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">			
			<p:cSld>
				<p:spTree>
					<p:nvGrpSpPr>
						<p:cNvPr id="1" name=""/>
						<p:cNvGrpSpPr/>
						<p:nvPr/>
					</p:nvGrpSpPr>
					<p:grpSpPr>
						<a:xfrm>
							<a:off x="0" y="0"/>
							<a:ext cx="0" cy="0"/>
							<a:chOff x="0" y="0"/>
							<a:chExt cx="0" cy="0"/>
						</a:xfrm>
					</p:grpSpPr>
					<xsl:for-each select ="draw:frame[@presentation:class[contains(.,'title') or contains(.,'outline')]]">						
						<xsl:choose >
							<!-- for convertion title-->							
							<xsl:when test ="@presentation:class ='title' 
							or @presentation:class ='subtitle'" >
								<p:sp>
									<p:nvSpPr>
										<p:cNvPr id="2" name="Title 1"/>
										<p:cNvSpPr>
											<a:spLocks noGrp="1"/>
										</p:cNvSpPr>
										<p:nvPr>
											<xsl:if test ="@presentation:class ='title'">
												<p:ph type="ctrTitle"/>
											</xsl:if>
											<xsl:if test ="@presentation:class ='subtitle'">
												<p:ph type="subTitle"/>
											</xsl:if>
										</p:nvPr>
									</p:nvSpPr>
									<p:spPr>										
										<a:xfrm>											
											 <a:off >
												 <xsl:attribute name ="x">												 
													 <!--<xsl:value-of select ="@svg:x"/>-->
													 <xsl:call-template name ="convertToPoints">
														 <xsl:with-param name ="unit" select ="'cm'"/>
														 <xsl:with-param name ="length" select ="@svg:x"/>
													 </xsl:call-template>
												 </xsl:attribute>
												 <xsl:attribute name ="y">
													 <!--<xsl:value-of select ="@svg:y"/>-->
													 <xsl:call-template name ="convertToPoints">
														 <xsl:with-param name ="unit" select ="'cm'"/>
														 <xsl:with-param name ="length" select ="@svg:y"/>
													 </xsl:call-template>
												 </xsl:attribute>
											 </a:off>
											<!--<a:ext cx="7772400" cy="1600200" />-->
											 <a:ext>
											 <xsl:attribute name ="cx">
												 <!--<xsl:value-of select ="@svg:width"/>-->
												 <xsl:call-template name ="convertToPoints">
													 <xsl:with-param name ="unit" select ="'cm'"/>
													 <xsl:with-param name ="length" select ="@svg:width"/>
												 </xsl:call-template>
											 </xsl:attribute>
											 <xsl:attribute name ="cy">
												 <!--<xsl:value-of select ="@svg:height"/>-->
												 <xsl:call-template name ="convertToPoints">
													 <xsl:with-param name ="unit" select ="'cm'"/>
													 <xsl:with-param name ="length" select ="@svg:height"/>
												 </xsl:call-template>
											 </xsl:attribute>										
										</a:ext>
										 </a:xfrm>
										<!-- Solid fill color -->
										<xsl:call-template name ="fillColor" >
											<xsl:with-param name ="prId" select ="@presentation:style-name" />
										</xsl:call-template>
									</p:spPr>
									<p:txBody>
										<a:bodyPr/>
										<a:lstStyle/>										
										<xsl:call-template name ="processText" >
											<xsl:with-param name ="layoutName" select ="@presentation:class"/>
										</xsl:call-template >																				
									</p:txBody>
								</p:sp>
							</xsl:when>
							<!-- For Converting content -->
							<xsl:when test ="@presentation:class ='outline'">
								<p:sp>
									<p:nvSpPr>
										<p:cNvPr id="3" name="Content Placeholder 2"/>
										<p:cNvSpPr>
											<a:spLocks noGrp="1"/>
										</p:cNvSpPr>
										<p:nvPr>
											<p:ph idx="1"/>
										</p:nvPr>
									</p:nvSpPr>
									<p:spPr>										
										<a:xfrm>											
											<a:off >
												<xsl:attribute name ="x">
													<!--<xsl:value-of select ="@svg:x"/>-->
													<xsl:call-template name ="convertToPoints">
														<xsl:with-param name ="unit" select ="'cm'"/>
														<xsl:with-param name ="length" select ="@svg:x"/>
													</xsl:call-template>
												</xsl:attribute>
												<xsl:attribute name ="y">
													<!--<xsl:value-of select ="@svg:y"/>-->
													<xsl:call-template name ="convertToPoints">
														<xsl:with-param name ="unit" select ="'cm'"/>
														<xsl:with-param name ="length" select ="@svg:y"/>
													</xsl:call-template>
												</xsl:attribute>
											</a:off>											
											<a:ext>
												<xsl:attribute name ="cx">
													<!--<xsl:value-of select ="@svg:width"/>-->
													<xsl:call-template name ="convertToPoints">
														<xsl:with-param name ="unit" select ="'cm'"/>
														<xsl:with-param name ="length" select ="@svg:width"/>
													</xsl:call-template>
												</xsl:attribute>
												<xsl:attribute name ="cy">
													<!--<xsl:value-of select ="@svg:height"/>-->
													<xsl:call-template name ="convertToPoints">
														<xsl:with-param name ="unit" select ="'cm'"/>
														<xsl:with-param name ="length" select ="@svg:height"/>
													</xsl:call-template>
												</xsl:attribute>
											</a:ext>
										</a:xfrm>
										<!-- Solid fill color -->
										<xsl:call-template name ="fillColor" >
											<xsl:with-param name ="prId" select ="@presentation:style-name" />
										</xsl:call-template>
									</p:spPr>
									<p:txBody>
										<a:bodyPr/>
										<a:lstStyle/>
										<xsl:call-template name ="processText" >
											<xsl:with-param name ="layoutName" select ="@presentation:class"></xsl:with-param>
										</xsl:call-template >
									</p:txBody>
								</p:sp>
							</xsl:when>
						</xsl:choose>						
					</xsl:for-each >
					
					<!-- Code for shapes start-->
					<xsl:call-template name ="shapes" />
					<!-- Code for shapes end-->
					
					<!-- Code for footer , slide number and date time control -->
					<xsl:variable name ="pageStyle">
						<xsl:value-of select ="@draw:style-name"/>
					</xsl:variable>
					<xsl:variable name ="footerId">
						<xsl:value-of select ="@presentation:use-footer-name"/>
					</xsl:variable>
					<xsl:variable name ="dateId">
						<xsl:value-of select ="@presentation:use-date-time-name"/>
					</xsl:variable>
					<xsl:for-each select ="document('content.xml')//style:style[@style:name=$pageStyle]">
						<xsl:if test ="style:drawing-page-properties[@presentation:display-footer='true']">
							<xsl:if test ="not($footerId ='')">
								<xsl:call-template name ="footer" >
									<xsl:with-param name ="footerId" select ="$footerId"/>
								</xsl:call-template >
							</xsl:if>
						</xsl:if>						
						<xsl:if test ="style:drawing-page-properties[@presentation:display-date-time='true']">
							<xsl:if test ="not($dateId='')">
								<xsl:call-template name ="footerDate" >
									<xsl:with-param name ="footerDateId" select ="$dateId"/>
								</xsl:call-template >
							</xsl:if>
						</xsl:if >
						<xsl:if test ="style:drawing-page-properties[@presentation:display-page-number='true']">
							<xsl:call-template name ="slideNumber">
								<xsl:with-param name ="pageNumber" select ="$pageNo"/>
							</xsl:call-template >
						</xsl:if >
					</xsl:for-each>
					
					
				</p:spTree>
			</p:cSld>
			<p:clrMapOvr>
				<a:masterClrMapping/>
			</p:clrMapOvr>
		</p:sld>
	</xsl:template>

<xsl:template name ="processText" >
		<xsl:param name ="layoutName"/>
		<xsl:variable name ="defFontSize">
			<xsl:value-of  select ="office:document-styles/office:styles/style:style/style:text-properties/@fo:font-size"/>
		</xsl:variable>		
		<xsl:choose >
			<xsl:when test ="$layoutName='title' or $layoutName='subtitle'">
				<xsl:choose>
					<xsl:when test ="draw:text-box/text:p/text:span">
						<xsl:for-each select ="draw:text-box/text:p">
							<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
								<!--<xsl:value-of select ="child::node()[1]"/>-->
								<xsl:for-each select ="child::node()[position()]">
									<!--<xsl:value-of select ="name()"/>-->
									<xsl:if test ="name()='text:span'">
										<xsl:for-each  select =".">
											<a:r >
												<a:rPr lang="en-US" smtClean="0">
													<!--Font Size -->
													<xsl:variable name ="textId">
														<xsl:value-of select ="@text:style-name"/>
													</xsl:variable>
													<xsl:if test ="not($textId ='')">
														<xsl:call-template name ="fontStyles">
															<xsl:with-param name ="Tid" select ="$textId" />
														</xsl:call-template>
													</xsl:if>
												</a:rPr >
												<a:t>
													<xsl:value-of select ="."/>
												</a:t>
											</a:r>
										</xsl:for-each>
									</xsl:if>
									<xsl:if test ="not(name() ='text:span')">
										<a:r >
											<a:rPr lang="en-US" smtClean="0">											
											</a:rPr >
											<a:t>
												<xsl:value-of select ="."/>
											</a:t>
										</a:r>
									</xsl:if >
								</xsl:for-each >
							</a:p>
						</xsl:for-each >
					</xsl:when>
					<xsl:when test ="draw:text-box/text:p">
						<xsl:for-each select ="draw:text-box/text:p">
							<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
								<a:r>
									<a:rPr lang="en-US" smtClean="0">										
									</a:rPr >
									<a:t>
										<xsl:value-of select ="."/>
									</a:t>
								</a:r>
							</a:p >
						</xsl:for-each >
					</xsl:when>
				</xsl:choose>
			</xsl:when>
			<xsl:when test ="not($layoutName='title' or $layoutName='subtitle')">
				<xsl:choose >
					<xsl:when test ="draw:text-box/text:p/text:span">
						<xsl:for-each select ="draw:text-box/text:p/text:span">
							<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
								<xsl:call-template name ="paraProperties" >
									<xsl:with-param name ="paraId" >
										<xsl:value-of select ="parent::text:p/@text:style-name"/>
									</xsl:with-param >
								</xsl:call-template >
								<a:r>
								<a:rPr lang="en-US" smtClean="0">
									<!--Font Size -->
									<xsl:variable name ="textId">
										<xsl:value-of select ="@text:style-name"/>
									</xsl:variable>
									<xsl:if test ="not($textId ='')">
										<xsl:call-template name ="fontStyles">
											<xsl:with-param name ="Tid" select ="$textId" />
										</xsl:call-template>
									</xsl:if>
								</a:rPr >
								<a:t>
									<xsl:value-of select ="."/>
								</a:t>
								</a:r>
							</a:p>	
						</xsl:for-each >
					</xsl:when>
					<xsl:when test ="draw:text-box/text:list/text:list-item/text:p/text:span">
						<xsl:for-each select ="draw:text-box/text:list/text:list-item/text:p">
							<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
								<xsl:call-template name ="paraProperties" >
									<xsl:with-param name ="paraId" >
										<xsl:value-of select ="@text:style-name"/>
									</xsl:with-param >
								</xsl:call-template >
								<xsl:for-each select ="child::node()[position()]">
									<xsl:choose >
										<xsl:when test ="name()='text:span'">
											<a:r>
												<a:rPr lang="en-US" smtClean="0">
													<!--Font Size -->
													<xsl:variable name ="textId">
														<xsl:value-of select ="@text:style-name"/>
													</xsl:variable>
													<xsl:if test ="not($textId ='')">
														<xsl:call-template name ="fontStyles">
															<xsl:with-param name ="Tid" select ="$textId" />
														</xsl:call-template>
													</xsl:if>
												</a:rPr >
												<a:t>
													<xsl:value-of select ="."/>
												</a:t>
											</a:r>
										</xsl:when >
										<xsl:when test ="not(name()='text:span')">
											<a:r>
												<a:rPr lang="en-US" smtClean="0">
													<!--Font Size -->
													<xsl:variable name ="textId">
														<xsl:value-of select ="@text:style-name"/>
													</xsl:variable>
													<xsl:if test ="not($textId ='')">
														<xsl:call-template name ="fontStyles">
															<xsl:with-param name ="Tid" select ="$textId" />
														</xsl:call-template>
													</xsl:if>
												</a:rPr >
												<a:t>
													<xsl:value-of select ="."/>
												</a:t>
											</a:r>
										</xsl:when >
									</xsl:choose>
								</xsl:for-each>
							</a:p >	
						</xsl:for-each >
					</xsl:when>
					<xsl:when test ="draw:text-box/text:list/text:list-item/text:p">
						<xsl:for-each select ="draw:text-box/text:list/text:list-item/text:p">
							<a:p  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
								<xsl:call-template name ="paraProperties" >
									<xsl:with-param name ="paraId" >
										<xsl:value-of select ="@text:style-name"/>
									</xsl:with-param >
								</xsl:call-template >
								<a:r >
									<a:rPr lang="en-US" smtClean="0">									
									</a:rPr >
									<a:t>
										<xsl:value-of select ="."/>
									</a:t>
								</a:r>
							</a:p>
						</xsl:for-each >
					</xsl:when>
					<xsl:when test ="draw:text-box/text:p">
					<xsl:for-each select ="draw:text-box/text:p">
						<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
							<xsl:call-template name ="paraProperties" >
								<xsl:with-param name ="paraId" >
									<xsl:value-of select ="@text:style-name"/>
								</xsl:with-param >
							</xsl:call-template >
							<a:r >
								<a:rPr lang="en-US" smtClean="0">									
								</a:rPr >
								<a:t>
									<xsl:value-of select ="."/>
								</a:t>
							</a:r>
						</a:p >
					</xsl:for-each >
				</xsl:when>
				</xsl:choose>
			</xsl:when >
	</xsl:choose>
</xsl:template>
<xsl:template name ="footer">
		<xsl:param name ="footerId"></xsl:param>
		<p:sp>
			<p:nvSpPr>
				<p:cNvPr id="6" name="Footer Placeholder 5" />
				<p:cNvSpPr>
					<a:spLocks noGrp="1" />
				</p:cNvSpPr>
				<p:nvPr>
					<p:ph type="ftr" sz="quarter" idx="11" />
				</p:nvPr>
			</p:nvSpPr>
			<p:spPr >
				<!-- footer date layout details style:master-page style:name -->
				<xsl:call-template name ="GetFrameDetails">
					<xsl:with-param name ="LayoutName" select ="'footer'"/>
				</xsl:call-template>				
			</p:spPr >
			<p:txBody>
				<a:bodyPr />
				<a:lstStyle />
				<a:p>
					<a:r>
						<a:rPr lang="en-US" smtClean="0" />
						<a:t>
							<xsl:for-each select ="document('content.xml') //presentation:footer-decl[@presentation:name=$footerId] ">
								<xsl:value-of select ="."/>
							</xsl:for-each >
						</a:t>
					</a:r>
					<a:endParaRPr lang="en-US" />
				</a:p>
			</p:txBody>
		</p:sp >
	</xsl:template>
<xsl:template name ="slideNumber">
		<xsl:param name ="pageNumber"/>
			<p:sp>
				<p:nvSpPr>
					<p:cNvPr id="5" name="Slide Number Placeholder 4" />
					<p:cNvSpPr>
						<a:spLocks noGrp="1" />
					</p:cNvSpPr>
					<p:nvPr>
						<p:ph type="sldNum" sz="quarter" idx="12" />
					</p:nvPr>
				</p:nvSpPr>
				<p:spPr >
					<!-- footer slide number layout details style:master-page style:name -->
					<xsl:call-template name ="GetFrameDetails">
						<xsl:with-param name ="LayoutName" select ="'page-number'"/>
					</xsl:call-template>
				</p:spPr >
				<p:txBody>
					<a:bodyPr />
					<a:lstStyle />
					<a:p>
						<a:fld type="slidenum">
							<xsl:attribute name ="id">
								<xsl:value-of select ="'{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}'"/>
							</xsl:attribute>
							<a:rPr lang="en-US" smtClean="0" />
							<a:pPr />
							<a:t>
								<xsl:value-of select ="$pageNumber"/>
							</a:t>
						</a:fld>
						<a:endParaRPr lang="en-US" />
					</a:p>
				</p:txBody>
			</p:sp >
	</xsl:template>
<xsl:template name ="footerDate">
		<xsl:param name ="footerDateId"/>
		<p:sp>
			<p:nvSpPr>
				<p:cNvPr id="4" name="Date Placeholder 3" />
				<p:cNvSpPr>
					<a:spLocks noGrp="1" />
				</p:cNvSpPr>
				<p:nvPr>
					<p:ph type="dt" sz="half" idx="10" />
				</p:nvPr>
			</p:nvSpPr>			
			<p:spPr >
				<!-- footer layout details style:master-page style:name -->
				<xsl:call-template name ="GetFrameDetails">
					<xsl:with-param name ="LayoutName" select ="'date-time'"/>
				</xsl:call-template>
			</p:spPr >
			<p:txBody>
				<a:bodyPr />
				<a:lstStyle />
				<a:p>
					<xsl:for-each select ="document('content.xml') 
					//presentation:date-time-decl[@presentation:name=$footerDateId] ">
						<xsl:choose>
							<xsl:when test="@presentation:source='current-date'" >
								<a:fld type="datetime1">
									<xsl:attribute name ="id">
										<xsl:value-of select ="'{86419996-E19B-43D7-A4AA-D671C2F15715}'"/>
									</xsl:attribute>
									<a:rPr lang="en-US" smtClean="0" />
									<a:t>
										<xsl:value-of select ="."/>	
									</a:t>
								</a:fld>
								<a:endParaRPr lang="en-US" />
							</xsl:when>
							<xsl:otherwise >
								<a:r>
									<a:rPr lang="en-US" smtClean="0" />
									<a:t>
										<xsl:value-of select ="."/>
									</a:t>
									<a:endParaRPr lang="en-US" />
								</a:r >
							</xsl:otherwise>
						</xsl:choose>
					</xsl:for-each >					
					
				</a:p>
			</p:txBody>
		</p:sp >
	</xsl:template>
<xsl:template name ="GetFrameDetails">
		<xsl:param name ="LayoutName"/>
		<xsl:for-each select ="document('styles.xml')//style:master-page[@style:name='Default']/draw:frame[@presentation:class=$LayoutName]">
			<!--<xsl:value-of select ="."/>-->
<a:xfrm>
			<a:off >
				<xsl:attribute name ="x">
					<!--<xsl:value-of select ="@svg:x"/>-->
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'cm'"/>
						<xsl:with-param name ="length" select ="@svg:x"/>
					</xsl:call-template>
				</xsl:attribute>
				<xsl:attribute name ="y">
					<!--<xsl:value-of select ="@svg:y"/>-->
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'cm'"/>
						<xsl:with-param name ="length" select ="@svg:y"/>
					</xsl:call-template>
				</xsl:attribute>
			</a:off >
			<a:ext>
				<xsl:attribute name ="cx">
					<!--<xsl:value-of select ="@svg:width"/>-->
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'cm'"/>
						<xsl:with-param name ="length" select ="@svg:width"/>
					</xsl:call-template>
				</xsl:attribute>
				<xsl:attribute name ="cy">
					<!--<xsl:value-of select ="@svg:height"/>-->
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'cm'"/>
						<xsl:with-param name ="length" select ="@svg:height"/>
					</xsl:call-template>
				</xsl:attribute>
			</a:ext >						
</a:xfrm>
		</xsl:for-each>
	</xsl:template>
<xsl:template name ="slidesRel">				
		<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
			<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
		</Relationships>
</xsl:template>

</xsl:stylesheet>
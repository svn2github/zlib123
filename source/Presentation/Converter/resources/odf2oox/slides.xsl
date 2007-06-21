<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
<!-- 
Copyright (c) 2007, Sonata Software Limited
* All rights reserved.
*
* Redistribution and use in source and binary forms, with or without
* modification, are permitted provided that the following conditions are met:
*
*     * Redistributions of source code must retain the above copyright
*       notice, this list of conditions and the following disclaimer.
*     * Redistributions in binary form must reproduce the above copyright
*       notice, this list of conditions and the following disclaimer in the
*       documentation and/or other materials provided with the distribution.
*     * Neither the name of Sonata Software Limited nor the names of its contributors
*       may be used to endorse or promote products derived from this software
*       without specific prior written permission.
*
* THIS SOFTWARE IS PROVIDED BY THE REGENTS AND CONTRIBUTORS ``AS IS'' AND ANY
* EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
* WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
* DISCLAIMED. IN NO EVENT SHALL THE REGENTS AND CONTRIBUTORS BE LIABLE FOR ANY
* DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
* (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
* LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
* ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
* (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
* SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE
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
  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  exclude-result-prefixes="odf style text number draw page">
	<xsl:import href ="common.xsl"/>
	<xsl:import href ="shapes_direct.xsl"/>
	<xsl:template name ="slides" match ="/office:document-content/office:body/office:presentation/draw:page" mode="slide">
	<xsl:param name ="pageNo"/>
	<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
		   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
		   xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
	<p:cSld>
	<!--Added by sateesh - Background Color-->
		<p:bg>
			<p:bgPr>
				<a:solidFill>
					<a:srgbClr>
						<xsl:variable name="dpName">
							<xsl:value-of select="@draw:style-name" />
						</xsl:variable>
						<xsl:variable name="SMName">
							<xsl:value-of select="@draw:master-page-name" />
						</xsl:variable>
						<xsl:variable name="SMdpName">
							<xsl:value-of select="document('styles.xml')//style:master-page[@style:name= $SMName]/@draw:style-name"/>
						</xsl:variable>
						<xsl:attribute name="val">
							<xsl:choose>
								<xsl:when test="document('content.xml')//style:style[@style:name= $dpName]/style:drawing-page-properties/@draw:fill-color">
									<xsl:if test="document('content.xml')//style:style[@style:name= $dpName]/style:drawing-page-properties">
										<xsl:value-of select="substring-after(document('content.xml')//style:style[@style:name= $dpName]/style:drawing-page-properties/@draw:fill-color,'#')" />
									</xsl:if>
								</xsl:when>
								<xsl:when test="document('styles.xml')//style:style[@style:name= $SMdpName]/style:drawing-page-properties/@draw:fill-color">
									<!--<xsl:when test="document('styles.xml')//style:master-page[@style:name= $SMName]">-->
									<!--<xsl:variable name="SMdpName">
                    <xsl:value-of select="document('styles.xml')//style:master-page[@style:name= $SMName]/@draw:style-name"/>
                  </xsl:variable>-->
									<xsl:if test="document('styles.xml')//style:style[@style:name= $SMdpName]/style:drawing-page-properties">
										<xsl:value-of select="substring-after(document('styles.xml')//style:style[@style:name= $SMdpName]/style:drawing-page-properties/@draw:fill-color,'#')" />
									</xsl:if>
								</xsl:when>
								<xsl:otherwise>
									<xsl:if test="not(@draw:fill-color)">
										<xsl:value-of select="'ffffff'" />
									</xsl:if>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</a:srgbClr>
				</a:solidFill>
				<a:effectLst />
			</p:bgPr>
		</p:bg>
	 <!--end-->
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
					<xsl:for-each select ="draw:frame[@presentation:class[contains(.,'title') or contains(.,'subTitle') or contains(.,'outline')]]">
					  <xsl:choose >
						<!-- for convertion title-->
						<xsl:when test ="@presentation:class ='title' or @presentation:class ='subTitle' or @presentation:class ='subtitle'" >
						  <p:sp>
							  <p:nvSpPr>
                    <!--Added by Venkatesh,Bug 1714218,File Unreadable in PP2003-->
                    <p:cNvPr name="Title 1">
                      <xsl:attribute name="id">
                        <xsl:value-of select="position()+1"/>
                      </xsl:attribute>
                    </p:cNvPr>
                    <!-- End-->               
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
										<xsl:call-template name ="TextAlignment" >
											<xsl:with-param name ="prId" select ="@presentation:style-name"/>
										</xsl:call-template >
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
                    <!-- Added by Venkatesh-->
                    <p:cNvPr name="Content Placeholder 2">
                      <xsl:attribute name ="id">
                        <xsl:value-of select ="position()+1"/>
                      </xsl:attribute>
                    </p:cNvPr>
                    <!--end by Venkatesh-->
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
										<xsl:call-template name ="TextAlignment" >
											<xsl:with-param name ="prId" select ="@presentation:style-name"/>
										</xsl:call-template >
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
							<!--<xsl:if test ="not($footerId ='')">-->
							<xsl:if test ="document('content.xml')//presentation:footer-decl[@presentation:name=$footerId]">
								<xsl:call-template name ="footer" >
									<xsl:with-param name ="footerId" select ="$footerId"/>
								</xsl:call-template >
							</xsl:if>
						</xsl:if>
						<xsl:if test ="style:drawing-page-properties[@presentation:display-date-time='true']">
							<!--<xsl:if test ="not($dateId='')">-->
							<xsl:if test ="document('content.xml')//presentation:date-time-decl[@presentation:name=$dateId]">
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
		<!-- Added by lohith.ar - Start - Mouse click events -->
		<xsl:variable name="PostionCount">
			<xsl:value-of select="position()"/>
		</xsl:variable>
		<!-- Start - variable to set the hyperlinks values for Frames-->
		<xsl:variable name="varFrameHyperLinks">
			<xsl:choose>
				<xsl:when test="office:event-listeners">
					<xsl:for-each select ="office:event-listeners/presentation:event-listener">
						<a:endParaRPr lang="en-US" dirty="0">
							<xsl:if test="@script:event-name[contains(.,'dom:click')]">
								<a:hlinkClick>
									<xsl:choose>
										<!-- Go to previous slide-->
										<xsl:when test="@presentation:action[ contains(.,'previous-page')]">
											<xsl:attribute name="action">
												<xsl:value-of select="'ppaction://hlinkshowjump?jump=previousslide'"/>
											</xsl:attribute>
										</xsl:when>
										<!-- Go to Next slide-->
										<xsl:when test="@presentation:action[ contains(.,'next-page')]">
											<xsl:attribute name="action">
												<xsl:value-of select="'ppaction://hlinkshowjump?jump=nextslide'"/>
											</xsl:attribute>
										</xsl:when>
										<!-- Go to First slide-->
										<xsl:when test="@presentation:action[ contains(.,'first-page')]">
											<xsl:attribute name="action">
												<xsl:value-of select="'ppaction://hlinkshowjump?jump=firstslide'"/>
											</xsl:attribute>
										</xsl:when>
										<!-- Go to Last slide-->
										<xsl:when test="@presentation:action[ contains(.,'last-page')]">
											<xsl:attribute name="action">
												<xsl:value-of select="'ppaction://hlinkshowjump?jump=lastslide'"/>
											</xsl:attribute>
										</xsl:when>
                    <!-- End Presentation -->
                    <xsl:when test="@presentation:action[ contains(.,'stop')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinkshowjump?jump=endshow'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- paly Sound-->
                    <xsl:when test="@presentation:action[ contains(.,'sound')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://noaction'"/>
                      </xsl:attribute>
                    </xsl:when>
										<!-- Run program-->
                    <xsl:when test="@xlink:href and @presentation:action[ contains(.,'execute')]">
											<xsl:attribute name="action">
												<xsl:value-of select="'ppaction://program'"/>
											</xsl:attribute>
										</xsl:when>
										<!-- Go to slide -->
										<xsl:when test="@xlink:href[ contains(.,'#page')]">
											<xsl:attribute name="action">
												<xsl:value-of select="'ppaction://hlinksldjump'"/>
											</xsl:attribute>
										</xsl:when>
										<!-- Go to document -->
										<xsl:when test="@xlink:href and @presentation:action[ contains(.,'show')] ">
											<xsl:if test="not(@xlink:href[ contains(.,'#page')])">
												<xsl:attribute name="action">
													<xsl:value-of select="'ppaction://hlinkfile'"/>
												</xsl:attribute>
											</xsl:if>
										</xsl:when>
									</xsl:choose>

									<!-- set value for attribute r:id-->
									<xsl:choose>
										<!-- For jump to next,previous,first,last-->
                    <xsl:when test="@presentation:action[ contains(.,'page') or contains(.,'stop')  or contains(.,'sound')]">
											<xsl:attribute name="r:id">
												<xsl:value-of select="''"/>
											</xsl:attribute>
										</xsl:when>
										<!-- For Run program & got to slide-->
                    <xsl:when test="@xlink:href[contains(.,'#page')]">
											<xsl:attribute name="r:id">
												<xsl:value-of select="concat('AtchFileId',$PostionCount)"/>
											</xsl:attribute>
										</xsl:when>
										<!-- For Go to document -->
										<xsl:when test="@xlink:href and @presentation:action[ contains(.,'show')] ">
											<xsl:if test="not(@xlink:href[ contains(.,'#page')])">
												<xsl:attribute name="r:id">
													<xsl:value-of select="concat('AtchFileId',$PostionCount)"/>
												</xsl:attribute>
											</xsl:if>
										</xsl:when>
                    <!-- For Go to Run Programs -->
                    <xsl:when test="@xlink:href and @presentation:action[ contains(.,'execute')]">
                      <xsl:attribute name="r:id">
                        <xsl:value-of select="concat('AtchFileId',$PostionCount)"/>
                      </xsl:attribute>
                    </xsl:when>
									</xsl:choose>
                  <!-- Play Sound-->
                  <xsl:if test="@presentation:action[ contains(.,'sound')]">
                    <a:snd>
                      <xsl:variable name="varMediaFilePath">
                        <xsl:if test="presentation:sound/@xlink:href [ contains(.,'../')]">
                          <xsl:value-of select="presentation:sound/@xlink:href" />
                        </xsl:if>
                        <xsl:if test="not(presentation:sound/@xlink:href [ contains(.,'../')])">
                          <xsl:value-of select="substring-after(presentation:sound/@xlink:href,'/')" />
                        </xsl:if>
                      </xsl:variable>
                      <xsl:variable name="varFileRelId">
                        <xsl:value-of select="translate(translate(translate(translate(translate($varMediaFilePath,'/','_'),'..','_'),'.','_'),':','_'),'%20D','_')"/>
                      </xsl:variable>
                      <xsl:attribute name="r:embed">
                        <xsl:value-of select="$varFileRelId"/>
                      </xsl:attribute>
                      <xsl:attribute name="name">
                        <xsl:value-of select="concat('Sound_',$PostionCount)"/>
                      </xsl:attribute>
                      <xsl:attribute name="builtIn">
                        <xsl:value-of select='1'/>
                      </xsl:attribute>
                      <pzip:extract pzip:source="{$varMediaFilePath}" pzip:target="{concat('ppt/media/',$varFileRelId,'.wav')}" />
                    </a:snd>
                  </xsl:if>
								</a:hlinkClick>
							</xsl:if>
						</a:endParaRPr>
					</xsl:for-each>
				</xsl:when>
			</xsl:choose>
		</xsl:variable>
		<!-- End - variable to set the hyperlinks values for Frames-->
		<!-- Start - variable to set the hyperlinks values for Text inside Frame-->
		<xsl:variable name="varTextHyperLinks">
			<xsl:choose>
				<xsl:when test="office:event-listeners">
					<xsl:for-each select ="office:event-listeners/presentation:event-listener">
						<xsl:if test="@script:event-name[contains(.,'dom:click')]">
							<a:hlinkClick>
								<xsl:choose>
									<!-- Go to previous slide-->
									<xsl:when test="@presentation:action[ contains(.,'previous-page')]">
										<xsl:attribute name="action">
											<xsl:value-of select="'ppaction://hlinkshowjump?jump=previousslide'"/>
										</xsl:attribute>
									</xsl:when>
									<!-- Go to Next slide-->
									<xsl:when test="@presentation:action[ contains(.,'next-page')]">
										<xsl:attribute name="action">
											<xsl:value-of select="'ppaction://hlinkshowjump?jump=nextslide'"/>
										</xsl:attribute>
									</xsl:when>
									<!-- Go to First slide-->
									<xsl:when test="@presentation:action[ contains(.,'first-page')]">
										<xsl:attribute name="action">
											<xsl:value-of select="'ppaction://hlinkshowjump?jump=firstslide'"/>
										</xsl:attribute>
									</xsl:when>
									<!-- Go to Last slide-->
									<xsl:when test="@presentation:action[ contains(.,'last-page')]">
										<xsl:attribute name="action">
											<xsl:value-of select="'ppaction://hlinkshowjump?jump=lastslide'"/>
										</xsl:attribute>
									</xsl:when>
                  <!-- End Presentation -->
                  <xsl:when test="@presentation:action[ contains(.,'stop')]">
                    <xsl:attribute name="action">
                      <xsl:value-of select="'ppaction://hlinkshowjump?jump=endshow'"/>
                    </xsl:attribute>
                  </xsl:when>
                  <!-- Play Sound -->
                  <xsl:when test="@presentation:action[ contains(.,'sound')]">
                    <xsl:attribute name="action">
                      <xsl:value-of select="'ppaction://noaction'"/>
                    </xsl:attribute>
                  </xsl:when>
									<!-- Run program-->
                  <xsl:when test="@xlink:href and @presentation:action[ contains(.,'execute')]">
										<xsl:attribute name="action">
											<xsl:value-of select="'ppaction://program'"/>
										</xsl:attribute>
									</xsl:when>
									<!-- Go to slide -->
									<xsl:when test="@xlink:href[ contains(.,'#page')]">
										<xsl:attribute name="action">
											<xsl:value-of select="'ppaction://hlinksldjump'"/>
										</xsl:attribute>
									</xsl:when>
									<!-- Go to document -->
									<xsl:when test="@xlink:href and @presentation:action[ contains(.,'show')] ">
										<xsl:if test="not(@xlink:href[ contains(.,'#page')])">
											<xsl:attribute name="action">
												<xsl:value-of select="'ppaction://hlinkfile'"/>
											</xsl:attribute>
										</xsl:if>
									</xsl:when>
								</xsl:choose>
								<!-- set value for attribute r:id-->
								<xsl:choose>
									<!-- For jump to next,previous,first,last-->
                  <xsl:when test="@presentation:action[ contains(.,'page') or contains(.,'stop') or contains(.,'sound')]">
										<xsl:attribute name="r:id">
											<xsl:value-of select="''"/>
										</xsl:attribute>
									</xsl:when>
									<!-- For Run program & got to slide-->
                  <xsl:when test="@xlink:href[contains(.,'#page')]">
										<xsl:attribute name="r:id">
											<xsl:value-of select="concat('AtchFileId',$PostionCount)"/>
										</xsl:attribute>
									</xsl:when>
									<!-- For Go to document -->
									<xsl:when test="@xlink:href and @presentation:action[ contains(.,'show')] ">
										<xsl:if test="not(@xlink:href[ contains(.,'#page')])">
											<xsl:attribute name="r:id">
												<xsl:value-of select="concat('AtchFileId',$PostionCount)"/>
											</xsl:attribute>
										</xsl:if>
									</xsl:when>
                  <!-- For Go to Run Programs -->
                  <xsl:when test="@xlink:href and @presentation:action[ contains(.,'execute')]">
                    <xsl:attribute name="r:id">
                      <xsl:value-of select="concat('AtchFileId',$PostionCount)"/>
                    </xsl:attribute>
                  </xsl:when>
								</xsl:choose>
                <!-- Play Sound-->
                <xsl:if test="@presentation:action[ contains(.,'sound')]">
                  <a:snd>
                    <xsl:variable name="varMediaFilePath">
                      <xsl:if test="presentation:sound/@xlink:href [ contains(.,'../')]">
                        <xsl:value-of select="presentation:sound/@xlink:href" />
                      </xsl:if>
                      <xsl:if test="not(presentation:sound/@xlink:href [ contains(.,'../')])">
                        <xsl:value-of select="substring-after(presentation:sound/@xlink:href,'/')" />
                      </xsl:if>
                    </xsl:variable>
                    <xsl:variable name="varFileRelId">
                      <xsl:value-of select="translate(translate(translate(translate(translate($varMediaFilePath,'/','_'),'..','_'),'.','_'),':','_'),'%20D','_')"/>
                    </xsl:variable>
                    <xsl:attribute name="r:embed">
                      <xsl:value-of select="$varFileRelId"/>
                    </xsl:attribute>
                    <xsl:attribute name="name">
                      <xsl:value-of select="concat('Sound_',$PostionCount)"/>
                    </xsl:attribute>
                    <xsl:attribute name="builtIn">
                      <xsl:value-of select='1'/>
                    </xsl:attribute>
                    <pzip:extract pzip:source="{$varMediaFilePath}" pzip:target="{concat('ppt/media/',$varFileRelId,'.wav')}" />
                  </a:snd>
                </xsl:if>
							</a:hlinkClick>
						</xsl:if>
					</xsl:for-each>
				</xsl:when>
			</xsl:choose>
		</xsl:variable>
		<!-- End - variable to set the hyperlinks values for Text inside Frame -->
		<xsl:variable name ="prClsName">
			<xsl:value-of select ="@presentation:class"/>
		</xsl:variable>
		<xsl:choose >
			<xsl:when test ="draw:text-box/text:p/text:span">
				<xsl:for-each select ="draw:text-box/text:p">
					<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
						<xsl:call-template name ="paraProperties" >
							<xsl:with-param name ="paraId" >
								<xsl:value-of select ="@text:style-name"/>								
							</xsl:with-param >
              <xsl:with-param name ="isBulleted" select ="'false'"/>
              <xsl:with-param name ="level" select ="'0'"/>
              <xsl:with-param name="framePresentaionStyleId" select="parent::node()/parent::node()/./@presentation:style-name" />
              <xsl:with-param name ="isNumberingEnabled" select ="'false'"/>
						</xsl:call-template >
						<xsl:for-each select ="child::node()[position()]">
							<xsl:choose >
								<xsl:when test ="name()='text:span'">
                  <!-- Added by lohith - bug fix 1731885 -->
                  <xsl:if test="node()">
									<a:r>
										<a:rPr lang="en-US" smtClean="0">
											<!--Font Size -->
											<xsl:variable name ="textId">
												<xsl:value-of select ="@text:style-name"/>
											</xsl:variable>
											<xsl:if test ="not($textId ='')">
												<xsl:call-template name ="fontStyles">
													<xsl:with-param name ="Tid" select ="$textId" />
													<xsl:with-param name ="prClassName" select ="$prClsName"/>
												</xsl:call-template>
											</xsl:if>
											<xsl:copy-of select="$varTextHyperLinks"/>
										</a:rPr >
										<a:t>
											<xsl:call-template name ="insertTab" />
										</a:t>
									</a:r>
									<xsl:if test ="text:line-break">
										<xsl:call-template name ="processBR">
											<xsl:with-param name ="T" select ="@text:style-name" />
											<xsl:with-param name ="prClassName" select ="$prClsName"/>
										</xsl:call-template>
									</xsl:if>
                  </xsl:if>
                  <!-- Added by lohith - bug fix 1731885 -->
                  <xsl:if test="not(node())">
                    <a:endParaRPr lang="en-US" smtClean="0">
                      <!--Font Size -->
                      <xsl:variable name ="textId">
                        <xsl:value-of select ="@text:style-name"/>
                      </xsl:variable>
                      <xsl:if test ="not($textId ='')">
                        <xsl:call-template name ="fontStyles">
                          <xsl:with-param name ="Tid" select ="$textId" />
                          <xsl:with-param name ="prClassName" select ="$prClsName"/>
                        </xsl:call-template>
                      </xsl:if>
                    </a:endParaRPr>
                  </xsl:if>

								</xsl:when >
								<xsl:when test ="name()='text:line-break'">
									<xsl:call-template name ="processBR">
										<xsl:with-param name ="T" select ="@text:style-name" />
										<xsl:with-param name ="prClassName" select ="$prClsName"/>
									</xsl:call-template>
								</xsl:when>
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
													<xsl:with-param name ="prClassName" select ="$prClsName"/>
												</xsl:call-template>
											</xsl:if>
											<xsl:copy-of select="$varTextHyperLinks"/>
										</a:rPr >
										<a:t>
											<xsl:call-template name ="insertTab" />
										</a:t>
									</a:r>
								</xsl:when >
							</xsl:choose>
						</xsl:for-each>
						<xsl:copy-of select="$varFrameHyperLinks"/>
					</a:p>
				</xsl:for-each >
			</xsl:when>
			<xsl:when test ="draw:text-box/text:list/text:list-item/text:p/text:span">
				<xsl:for-each select ="draw:text-box/text:list">
					<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <!--Code inserted by Vijayets for Bullets and numbering-->
            <!--Check if Levels are present-->
            <xsl:variable name ="lvl">
              <xsl:if test ="text:list-item/text:list">
                <xsl:call-template name ="getListLevel">
                  <xsl:with-param name ="levelCount"/>
                </xsl:call-template>
              </xsl:if>
              <xsl:if test ="not(text:list-item/text:list)">
                <xsl:value-of select ="'0'"/>
              </xsl:if>
            </xsl:variable >
            <xsl:variable name="paragraphId" >
              <xsl:call-template name ="getParaStyleName">
                <xsl:with-param name ="lvl" select ="$lvl"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name ="isNumberingEnabled">              
              <xsl:if test ="document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering">
                <xsl:value-of select ="document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering"/>
              </xsl:if>
              <xsl:if test ="not(document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering)">
                <xsl:value-of select ="'true'"/>
              </xsl:if>
            </xsl:variable>
						<xsl:call-template name ="paraProperties" >
							<xsl:with-param name ="paraId" >                
                  <xsl:value-of select ="$paragraphId"/>                             
              </xsl:with-param >
              <!-- list property also included-->
              <xsl:with-param name ="listId">
								<xsl:value-of select ="@text:style-name"/>
							</xsl:with-param >
              <!-- Parameters added by vijayeta,Set bulleting as true/false,and set level -->
              <xsl:with-param name ="isBulleted" select ="'true'"/>
              <xsl:with-param name ="level" select ="$lvl"/>
              <xsl:with-param name ="isNumberingEnabled" select ="$isNumberingEnabled"/>
						</xsl:call-template >
            <!--End of Code inserted by Vijayets for Bullets and numbering-->
						<xsl:for-each select ="child::node()[position()]">
							<xsl:choose >
                <xsl:when test ="name()='text:list-item'">
                  <xsl:variable name ="currentNodeStyle">
                    <xsl:call-template name ="getTextNodeForFontStyle">
                      <xsl:with-param name ="this"/>
                    </xsl:call-template>
                  </xsl:variable>
									<a:r>
										<a:rPr lang="en-US" smtClean="0">
											<!--Font Size -->
											<xsl:variable name ="textId">
												<xsl:value-of select ="$currentNodeStyle"/>
											</xsl:variable>
											<xsl:if test ="not($textId ='')">
												<xsl:call-template name ="fontStyles">
													<xsl:with-param name ="Tid" select ="$textId" />
													<xsl:with-param name ="prClassName" select ="$prClsName"/>
                          <xsl:with-param name ="lvl" select ="$lvl"/>
												</xsl:call-template>
											</xsl:if>
											<xsl:copy-of select="$varTextHyperLinks"/>
										</a:rPr >
										<a:t>
											<xsl:call-template name ="insertTab" />
										</a:t>
									</a:r>
								</xsl:when >							
							</xsl:choose>
						</xsl:for-each>
						<xsl:copy-of select="$varFrameHyperLinks"/>
					</a:p >
				</xsl:for-each >
			</xsl:when>
			<xsl:when test ="draw:text-box/text:list/text:list-item/text:p">
				<xsl:for-each select ="draw:text-box/text:list">
					<a:p  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <!--Check if Levels are present-->
            <xsl:variable name ="lvl">
              <xsl:if test ="text:list-item/text:list">
                <xsl:call-template name ="getListLevel">
                  <xsl:with-param name ="levelCount"/>
                </xsl:call-template>
              </xsl:if>
              <xsl:if test ="not(text:list-item/text:list)">
                <xsl:value-of select ="'0'"/>
              </xsl:if>
            </xsl:variable >
            <xsl:variable name="paragraphId" >
              <xsl:call-template name ="getParaStyleName">
                <xsl:with-param name ="lvl" select ="$lvl"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name ="isNumberingEnabled">
              <xsl:if test ="document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering">
                <xsl:value-of select ="document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering"/>
              </xsl:if>
              <xsl:if test ="not(document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering)">
                <xsl:value-of select ="'true'"/>
              </xsl:if>
            </xsl:variable>
						<xsl:call-template name ="paraProperties" >
              <xsl:with-param name ="paraId" >
                <xsl:value-of select ="$paragraphId"/>
              </xsl:with-param >
              <!-- list property also included-->
              <xsl:with-param name ="listId">
								<xsl:value-of select ="@text:style-name"/>
							</xsl:with-param >

              <!-- Parameters added by vijayeta,Set bulleting as true/false,and set level -->
              <xsl:with-param name ="isBulleted" select ="'true'"/>
              <xsl:with-param name ="level" select ="$lvl"/>
              <xsl:with-param name ="isNumberingEnabled" select ="$isNumberingEnabled"/>
						</xsl:call-template >
            <!--End of Code inserted by Vijayets for Bullets and numbering-->
            <xsl:for-each select ="child::node()[position()]">
              <xsl:choose >
                <xsl:when test ="name()='text:list-item'">
                  <xsl:variable name ="currentNodeStyle">
                    <xsl:call-template name ="getTextNodeForFontStyle">
                      <xsl:with-param name ="this"/>
                    </xsl:call-template>
                  </xsl:variable>                
						<a:r >
							<a:rPr lang="en-US" smtClean="0">
                <!--Font Size -->
                <xsl:variable name ="textId">
                  <xsl:value-of select ="$currentNodeStyle"/>
                </xsl:variable>
                <xsl:if test ="not($textId ='')">
                  <xsl:call-template name ="fontStyles">
                    <xsl:with-param name ="Tid" select ="$textId" />
                    <xsl:with-param name ="prClassName" select ="$prClsName"/>
                    <xsl:with-param name ="lvl" select ="$lvl"/>
                  </xsl:call-template>
                </xsl:if>
                <!--End of code, Applying text styles--> 												
							</a:rPr >
							<a:t>
								<xsl:call-template name ="insertTab" />
							</a:t>
						</a:r>
                </xsl:when>
              </xsl:choose>
            </xsl:for-each >
						<xsl:copy-of select="$varFrameHyperLinks"/>
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
              <xsl:with-param name ="isBulleted" select ="'false'"/>
              <xsl:with-param name ="level" select ="'0'"/>
              <xsl:with-param name="framePresentaionStyleId" select="parent::node()/parent::node()/./@presentation:style-name" />
              <xsl:with-param name ="isNumberingEnabled" select ="'false'"/>
						</xsl:call-template >
						<a:r >
							<a:rPr lang="en-US" smtClean="0">
								<xsl:variable name ="DefFontSize">
									<xsl:call-template name ="getDefaultFontSize">
										<xsl:with-param name ="className" select ="$prClsName"/>
									</xsl:call-template >
								</xsl:variable>
								<xsl:if  test ="$DefFontSize!=''">
									<xsl:attribute name ="sz">										
										<xsl:call-template name ="convertToPoints">
											<xsl:with-param name ="unit" select ="'pt'"/>
											<xsl:with-param name ="length" select ="$DefFontSize"/>
										</xsl:call-template>
									</xsl:attribute>
								</xsl:if>
                <a:latin charset="0" typeface="Arial" />
								<xsl:copy-of select="$varTextHyperLinks"/>								
							</a:rPr >
							<a:t>
								<xsl:call-template name ="insertTab" />
							</a:t>
						</a:r>
						<xsl:if test ="text:span/text:line-break">
							<xsl:call-template name ="processBR">
								<xsl:with-param name ="T" select ="text:span/@text:style-name" />
								<xsl:with-param name ="prClassName" select ="$prClsName"/>
							</xsl:call-template>
						</xsl:if>
						<xsl:copy-of select="$varFrameHyperLinks"/>
					</a:p >
				</xsl:for-each >
			</xsl:when>
			<xsl:otherwise >
				<a:p>
					<a:r >
						<a:rPr lang="en-US" smtClean="0">
							<a:latin charset="0" typeface="Arial" />
							<xsl:copy-of select="$varTextHyperLinks"/>
						</a:rPr >
						<a:t>							
							<xsl:call-template name ="insertTab" />
						</a:t>
					</a:r>
					<xsl:copy-of select="$varFrameHyperLinks"/>
				</a:p >
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name ="processBR">
		<xsl:param name ="T" />
		<a:br>
			<a:rPr lang="en-US" smtClean="0">
				<!--Font Size -->
				<xsl:if test ="$T !=''">
					<xsl:call-template name ="fontStyles">
						<xsl:with-param name ="Tid" select ="$T" />
					</xsl:call-template>
				</xsl:if>
				<xsl:if test ="$T =''">
					<a:latin charset="0" typeface="Arial" />
				</xsl:if >
			</a:rPr >
		</a:br>
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
					<!--<a:endParaRPr lang="en-US" />-->
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
					<!--<a:endParaRPr lang="en-US" />-->
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
								<a:fld >									
									<xsl:attribute name ="id">
										<xsl:value-of select ="'{86419996-E19B-43D7-A4AA-D671C2F15715}'"/>
									</xsl:attribute>
									<xsl:attribute name ="type">
										<xsl:choose >
											<xsl:when test ="@style:data-style-name ='D3'">
												<xsl:value-of select ="'datetime1'"/>
											</xsl:when>											
											<xsl:when test ="@style:data-style-name ='D8'">
												<xsl:value-of select ="'datetime2'"/>
											</xsl:when>
											<xsl:when test ="@style:data-style-name ='D6'">
												<xsl:value-of select ="'datetime4'"/>
											</xsl:when>
											<xsl:when test ="@style:data-style-name ='D5'">
												<xsl:value-of select ="'datetime4'"/>
											</xsl:when>
											<xsl:when test ="@style:data-style-name ='D3T2'">
												<xsl:value-of select ="'datetime8'"/>
											</xsl:when>
											<xsl:when test ="@style:data-style-name ='D3T5'">
												<xsl:value-of select ="'datetime8'"/>
											</xsl:when>
											<xsl:when test ="@style:data-style-name ='T2'">
												<xsl:value-of select ="'datetime10'"/>
											</xsl:when>
											<xsl:when test ="@style:data-style-name ='T3'">
												<xsl:value-of select ="'datetime11'"/>
											</xsl:when>
											<xsl:when test ="@style:data-style-name ='T5'">
												<xsl:value-of select ="'datetime12'"/>
											</xsl:when>
											<xsl:when test ="@style:data-style-name ='T6'">
												<xsl:value-of select ="'datetime13'"/>
											</xsl:when>
											<xsl:otherwise>
												<xsl:value-of select ="'datetime1'"/>												
											</xsl:otherwise>
										</xsl:choose>
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
								</a:r >
							</xsl:otherwise>
						</xsl:choose>
					</xsl:for-each>
				</a:p>
			</p:txBody>
		</p:sp >
	</xsl:template>
	<xsl:template name ="GetFrameDetails">
		<xsl:param name ="LayoutName"/>
		<xsl:for-each select ="document('styles.xml')//style:master-page[@style:name='Default']/draw:frame[@presentation:class=$LayoutName]">			
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
	<xsl:template name ="TextAlignment">
		<xsl:param name ="prId"/>
		<a:bodyPr>			
			<xsl:for-each select ="draw:text-box">
				<xsl:for-each select ="text:p">
					<xsl:variable name ="ParId">
						<xsl:value-of select ="@text:style-name"/>
					</xsl:variable>
					<xsl:for-each select ="document('content.xml')//style:style[@style:name=$ParId]/style:paragraph-properties">
						<xsl:if test ="@style:writing-mode='tb-rl'">
							<xsl:attribute name ="vert">
								<xsl:value-of select ="'vert'"/>
							</xsl:attribute>
						</xsl:if>
					</xsl:for-each>
				</xsl:for-each>
			</xsl:for-each>
			<xsl:for-each select ="document('content.xml')//style:style[@style:name=$prId]/style:graphic-properties">				
				<xsl:attribute name ="tIns">
					<xsl:if test ="@fo:padding-top">
						<xsl:value-of select ="format-number(substring-before(@fo:padding-top,'cm')  *   360000 ,'#.##')"/>
					</xsl:if>
          <!-- Added by Lohith - or @fo:padding-top = '0cm'-->
					<xsl:if test ="not(@fo:padding-top) or @fo:padding-top = '0cm'">
						<xsl:value-of select ="0"/>
					</xsl:if >
				</xsl:attribute>
				<xsl:attribute name ="bIns">
					<xsl:if test ="@fo:padding-bottom">
						<xsl:value-of select ="format-number(substring-before(@fo:padding-bottom,'cm')  *   360000 ,'#.##')"/>
					</xsl:if>
          <!-- Added by Lohith - or @fo:padding-bottom = '0cm'-->
					<xsl:if test ="not(@fo:padding-bottom) or @fo:padding-bottom = '0cm'">
						<xsl:value-of select ="0"/>
					</xsl:if >
				</xsl:attribute>
				<xsl:attribute name ="lIns">
					<xsl:if test ="@fo:padding-left">
						<xsl:value-of select ="format-number(substring-before(@fo:padding-left,'cm')  *   360000 ,'#.##')"/>
					</xsl:if>
          <!-- Added by Lohith - or @fo:padding-left = '0cm'-->
					<xsl:if test ="not(@fo:padding-left) or @fo:padding-left = '0cm'">
						<xsl:value-of select ="0"/>
					</xsl:if >
				</xsl:attribute>
				<xsl:attribute name ="rIns">
					<xsl:if test ="@fo:padding-right">
						<xsl:value-of select ="format-number(substring-before(@fo:padding-right,'cm')  *   360000 ,'#.##')"/>
					</xsl:if>
          <!-- Added by Lohith - or @fo:padding-right = '0cm'-->
					<xsl:if test ="not(@fo:padding-right) or @fo:padding-right = '0cm'">
						<xsl:value-of select ="0"/>
					</xsl:if >
				</xsl:attribute>
				<xsl:attribute name ="anchor">
					<xsl:choose >
						<xsl:when test ="draw:textarea-vertical-align ='top'">
							<xsl:value-of select ="'t'"/>
						</xsl:when>
						<xsl:when test ="draw:textarea-vertical-align ='middle'">
							<xsl:value-of select ="'ctr'"/>
						</xsl:when>
						<xsl:when test ="draw:textarea-vertical-align ='bottom'">
							<xsl:value-of select ="'b'"/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select ="'ctr'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
				<xsl:attribute name ="anchorCtr">
					<xsl:choose >
						<xsl:when test ="draw:textarea-horizontal-align ='center'">
							<xsl:value-of select ="1"/>
						</xsl:when>
						<xsl:when test ="draw:textarea-horizontal-align='justify'">
							<xsl:value-of select ="0"/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select ="0"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
				<xsl:attribute name ="wrap">
					<xsl:choose >
						<xsl:when test ="fo:wrap-option ='no-wrap'">
							<xsl:value-of select ="'none'"/>
						</xsl:when>
						<xsl:when test ="fo:wrap-option ='wrap'">
							<xsl:value-of select ="'square'"/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select ="'square'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>

			</xsl:for-each>
		</a:bodyPr>
	</xsl:template>
	<!--<xsl:template name ="getDefaultFontSize">
		<xsl:param name ="prClassName" />
		<xsl:variable name ="LayoutName">
			<xsl:call-template name ="getClassName" >
				<xsl:with-param name ="clsName" select="$prClassName"/>
			</xsl:call-template>
		</xsl:variable>
		
	</xsl:template>-->	
	<xsl:template name ="slidesRel" match ="/office:document-content/office:body/office:presentation/draw:page">
		<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
			<xsl:variable name ="ctrTitle">
				<xsl:for-each select ="draw:frame/@presentation:class">
					<xsl:if test =". = 'title'">
						<xsl:value-of select ="."/>
					</xsl:if>
				</xsl:for-each>
			</xsl:variable >

			<xsl:variable name ="ctrSubTitle">
				<xsl:for-each select ="draw:frame/@presentation:class">
					<xsl:if test =". = 'subtitle'">
						<xsl:value-of select ="."/>
					</xsl:if>
				</xsl:for-each>
			</xsl:variable >

			<xsl:variable name ="ctrOutline">
				<xsl:for-each select ="draw:frame/@presentation:class">
					<xsl:if test =". = 'outline'">
						<xsl:value-of select ="."/>
					</xsl:if>
				</xsl:for-each>
			</xsl:variable >

			<xsl:variable name ="ctrObject">
				<xsl:for-each select ="draw:frame/@presentation:class">
					<xsl:if test =". = 'object'">
						<xsl:value-of select ="."/>
					</xsl:if>
				</xsl:for-each>
			</xsl:variable>
			<xsl:variable name ="ctrChart">
				<xsl:for-each select ="draw:frame/@presentation:class">
					<xsl:if test =". = 'object'">
						<xsl:value-of select ="."/>
					</xsl:if>
				</xsl:for-each>
			</xsl:variable >
			<xsl:variable name ="ctrGraphic">
				<xsl:for-each select ="draw:frame/@presentation:class">
					<xsl:if test =". = 'object'">
						<xsl:value-of select ="."/>
					</xsl:if>
				</xsl:for-each>
			</xsl:variable>
			<xsl:choose >
				<xsl:when test ="@presentation:presentation-page-layout-name='AL0T0'" >
					<Relationship Id="rId1" 
					Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" 
					Target="../slideLayouts/slideLayout7.xml"/>
				</xsl:when>
				<xsl:when test ="$ctrTitle='title' 
								and $ctrSubTitle ='' 
								and $ctrOutline=''
								and $ctrObject=''
								and $ctrChart=''
								and $ctrGraphic=''">
					<Relationship Id="rId1" 
						Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" 
						Target="../slideLayouts/slideLayout6.xml"/>
				</xsl:when>
				<xsl:when test ="$ctrTitle='title' 
								and $ctrSubTitle ='subtitle' 
								and $ctrOutline=''
								and $ctrObject=''
								and $ctrChart=''
								and $ctrGraphic=''">
					<Relationship Id="rId1" 
						Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" 
						Target="../slideLayouts/slideLayout1.xml"/>
				</xsl:when>
				<xsl:when test ="$ctrTitle='title' 
								and $ctrSubTitle ='' 
								and $ctrOutline='outline'
								and $ctrObject=''
								and $ctrChart=''
								and $ctrGraphic=''">
					<Relationship Id="rId1" 
						Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" 
						Target="../slideLayouts/slideLayout2.xml"/>
				</xsl:when>
				<xsl:when test ="$ctrTitle='title' 
								and $ctrSubTitle ='' 
								and $ctrOutline=''
								and $ctrObject='object'
								and $ctrChart=''
								and $ctrGraphic=''">
					<Relationship Id="rId1" 
						Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" 
						Target="../slideLayouts/slideLayout2.xml"/>
				</xsl:when>

				<xsl:when test ="$ctrTitle='title' 
								and $ctrSubTitle ='' 
								and $ctrOutline=''
								and $ctrObject=''
								and $ctrChart='chart'
								and $ctrGraphic=''">
					<Relationship Id="rId1" 
						Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" 
						Target="../slideLayouts/slideLayout2.xml"/>
				</xsl:when >

				<xsl:when test ="$ctrTitle='title' 
								and $ctrSubTitle ='' 
								and $ctrOutline=''
								and $ctrObject=''
								and $ctrChart=''
								and $ctrGraphic='graphic'">
					<Relationship Id="rId1" 
						Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" 
						Target="../slideLayouts/slideLayout2.xml"/>
				</xsl:when>
				<!-- Layout codition starts here here-->
				<xsl:when test ="$ctrTitle='title' 
								and $ctrSubTitle ='' 
								and $ctrOutline='outline'
								and $ctrObject='object'
								and $ctrChart=''
								and $ctrGraphic=''">
					<Relationship Id="rId1" 
						Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" 
						Target="../slideLayouts/slideLayout4.xml"/>
				</xsl:when>
				<xsl:when test ="$ctrTitle='title' 
								and $ctrSubTitle ='' 
								and $ctrOutline='outline'
								and $ctrObject=''
								and $ctrChart='chart'
								and $ctrGraphic=''">
					<Relationship Id="rId1" 
						Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" 
						Target="../slideLayouts/slideLayout4.xml"/>
				</xsl:when>

				<xsl:when test ="$ctrTitle='title' 
								and $ctrSubTitle ='' 
								and $ctrOutline='outline'
								and $ctrObject=''
								and $ctrChart=''
								and $ctrGraphic='graphic'">
					<Relationship Id="rId1" 
						Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" 
						Target="../slideLayouts/slideLayout4.xml"/>
				</xsl:when>

				<xsl:when test ="$ctrTitle='title' 
								and $ctrSubTitle ='' 
								and $ctrOutline=''
								and $ctrObject='object'
								and $ctrChart='chart'
								and $ctrGraphic=''">
					<Relationship Id="rId1" 
						Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" 
						Target="../slideLayouts/slideLayout4.xml"/>
				</xsl:when>
				<xsl:when test ="$ctrTitle='title' 
								and $ctrSubTitle ='' 
								and $ctrOutline=''
								and $ctrObject='object'
								and $ctrChart=''
								and $ctrGraphic='graphic'">
					<Relationship Id="rId1" 
						Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" 
						Target="../slideLayouts/slideLayout4.xml"/>
				</xsl:when>
				<xsl:when test ="$ctrTitle='title' 
								and $ctrSubTitle ='' 
								and $ctrOutline=''
								and $ctrObject=''
								and $ctrChart='chart'
								and $ctrGraphic='graphic'">
					<Relationship Id="rId1" 
						Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" 
						Target="../slideLayouts/slideLayout4.xml"/>
				</xsl:when>
				<xsl:when test ="$ctrTitle='title' and
								 count(draw:frame/@presentation:class[contains(.,'outline')]) =  2 ">
					<Relationship Id="rId1" 
						Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" 
						Target="../slideLayouts/slideLayout4.xml"/>
				</xsl:when >
				<!-- Layout 4 codition ends here-->
				<xsl:otherwise >
					<Relationship Id="rId1" 
						Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" 
						Target="../slideLayouts/slideLayout7.xml"/>
				</xsl:otherwise>
			</xsl:choose>

      <!-- Added by lohith - Start - Add relation files of hyperlinks-->
      <xsl:if test="draw:custom-shape/office:event-listeners">
        <xsl:for-each select="draw:custom-shape">
          <xsl:variable name="ShapePostionCount">
            <xsl:value-of select="position()"/>
          </xsl:variable>
          <xsl:for-each select ="office:event-listeners/presentation:event-listener">
            <xsl:if test="@xlink:href">
              <Relationship>
                <xsl:attribute name="Id">
                  <xsl:value-of select="concat('ShapeFileId',$ShapePostionCount)"/>
                </xsl:attribute>
                <xsl:choose>
                  <xsl:when test="@xlink:href[contains(.,'#page')]">
                    <xsl:attribute name="Type">
                      <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
                    </xsl:attribute>
                    <xsl:attribute name="Target">
                      <xsl:value-of select="concat('slide',substring-after(@xlink:href,'page'),'.xml')"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:when test="@xlink:href">
                    <xsl:attribute name="Type">
                      <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                    </xsl:attribute>
                    <xsl:attribute name="Target">
                      <xsl:if test="@xlink:href[ contains(.,'./')]">
                        <xsl:if test="string-length(substring-after(@xlink:href, '../')) = 0">
                          <xsl:value-of select="/"/>
                        </xsl:if>
                        <xsl:if test="not(string-length(substring-after(@xlink:href, '../')) = 0)">
                          <xsl:value-of select="substring-after(@xlink:href, '../')"/>
                        </xsl:if>
                      </xsl:if>
                      <xsl:if test="not(@xlink:href[ contains(.,'./')])">
                        <xsl:value-of select="concat('file://',@xlink:href)"/>
                      </xsl:if>
                    </xsl:attribute>
                    <xsl:attribute name="TargetMode">
                      <xsl:value-of select="'External'"/>
                    </xsl:attribute>
                  </xsl:when>
                </xsl:choose>
              </Relationship>
            </xsl:if>
            <xsl:if test="presentation:sound">
              <xsl:variable name="varblnDuplicateRelation">
                <xsl:call-template name="GetUniqueRelationIdForWavFile">
                  <xsl:with-param name="FilePath" select="presentation:sound/@xlink:href" />
                  <xsl:with-param name="ShapePosition" select="$ShapePostionCount" />
                  <xsl:with-param name="ShapeType" select="'CUSTOM'" />
                  <xsl:with-param name="Page" select="parent::node()/parent::node()/parent::node()" />
                </xsl:call-template>
              </xsl:variable>
              <xsl:if test="$varblnDuplicateRelation != 1">
                <Relationship>
                  <xsl:variable name="varMediaFilePath">
                    <xsl:if test="presentation:sound/@xlink:href [ contains(.,'../')]">
                      <xsl:value-of select="presentation:sound/@xlink:href" />
                    </xsl:if>
                    <xsl:if test="not(presentation:sound/@xlink:href [ contains(.,'../')])">
                      <xsl:value-of select="substring-after(presentation:sound/@xlink:href,'/')" />
                    </xsl:if>
                  </xsl:variable>
                  <xsl:variable name="varFileRelId">
                    <xsl:value-of select="translate(translate(translate(translate(translate($varMediaFilePath,'/','_'),'..','_'),'.','_'),':','_'),'%20D','_')"/>
                  </xsl:variable>
                  <xsl:attribute name="Id">
                    <xsl:value-of select="$varFileRelId"/>
                  </xsl:attribute>
                  <!--<xsl:attribute name="Id">
                  <xsl:value-of select="concat('ShapeFileId',$ShapePostionCount)"/>
                </xsl:attribute>-->
                  <xsl:attribute name="Type">
                    <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio'"/>
                  </xsl:attribute>
                  <xsl:attribute name="Target">
                    <xsl:value-of select="concat('../media/',$varFileRelId,'.wav')"/>
                  </xsl:attribute>
                </Relationship>
              </xsl:if>
            </xsl:if>
          </xsl:for-each>
        </xsl:for-each>
      </xsl:if>
      <xsl:if test="draw:rect/office:event-listeners">
        <xsl:for-each select="draw:rect">
          <xsl:variable name="ShapePostionCount">
            <xsl:value-of select="position()"/>
          </xsl:variable>
          <xsl:for-each select ="office:event-listeners/presentation:event-listener">
            <xsl:if test="@xlink:href">
              <Relationship>
                <xsl:attribute name="Id">
                  <xsl:value-of select="concat('RectAtachFileId',$ShapePostionCount)"/>
                </xsl:attribute>
                <xsl:choose>
                  <xsl:when test="@xlink:href[contains(.,'#page')]">
                    <xsl:attribute name="Type">
                      <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
                    </xsl:attribute>
                    <xsl:attribute name="Target">
                      <xsl:value-of select="concat('slide',substring-after(@xlink:href,'page'),'.xml')"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:when test="@xlink:href">
                    <xsl:attribute name="Type">
                      <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                    </xsl:attribute>
                    <xsl:attribute name="Target">
                      <xsl:if test="@xlink:href[ contains(.,'./')]">
                        <xsl:if test="string-length(substring-after(@xlink:href, '../')) = 0">
                          <xsl:value-of select="/"/>
                        </xsl:if>
                        <xsl:if test="not(string-length(substring-after(@xlink:href, '../')) = 0)">
                          <xsl:value-of select="substring-after(@xlink:href, '../')"/>
                        </xsl:if>
                      </xsl:if>
                      <xsl:if test="not(@xlink:href[ contains(.,'./')])">
                        <xsl:value-of select="concat('file://',@xlink:href)"/>
                      </xsl:if>
                    </xsl:attribute>
                    <xsl:attribute name="TargetMode">
                      <xsl:value-of select="'External'"/>
                    </xsl:attribute>
                  </xsl:when>
                </xsl:choose>
              </Relationship>
            </xsl:if>
            <xsl:if test="presentation:sound">
              <xsl:variable name="varblnDuplicateRelation">
                <xsl:call-template name="GetUniqueRelationIdForWavFile">
                  <xsl:with-param name="FilePath" select="presentation:sound/@xlink:href" />
                  <xsl:with-param name="ShapePosition" select="$ShapePostionCount" />
                  <xsl:with-param name="ShapeType" select="'RECT'" />
                  <xsl:with-param name="Page" select="parent::node()/parent::node()/parent::node()" />
                </xsl:call-template>
              </xsl:variable>
              <xsl:if test="$varblnDuplicateRelation != 1">
                <Relationship>
                  <xsl:variable name="varMediaFilePath">
                    <xsl:if test="presentation:sound/@xlink:href [ contains(.,'../')]">
                      <xsl:value-of select="presentation:sound/@xlink:href" />
                    </xsl:if>
                    <xsl:if test="not(presentation:sound/@xlink:href [ contains(.,'../')])">
                      <xsl:value-of select="substring-after(presentation:sound/@xlink:href,'/')" />
                    </xsl:if>
                  </xsl:variable>
                  <xsl:variable name="varFileRelId">
                    <xsl:value-of select="translate(translate(translate(translate(translate($varMediaFilePath,'/','_'),'..','_'),'.','_'),':','_'),'%20D','_')"/>
                  </xsl:variable>
                  <xsl:attribute name="Id">
                    <xsl:value-of select="$varFileRelId"/>
                  </xsl:attribute>
                  <!--<xsl:attribute name="Id">
                  <xsl:value-of select="concat('RectAtachFileId',$ShapePostionCount)"/>
                </xsl:attribute>-->
                  <xsl:attribute name="Type">
                    <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio'"/>
                  </xsl:attribute>
                  <xsl:attribute name="Target">
                    <xsl:value-of select="concat('../media/',$varFileRelId,'.wav')"/>
                  </xsl:attribute>
                </Relationship>
              </xsl:if>
            </xsl:if>
          </xsl:for-each>
        </xsl:for-each>
      </xsl:if>
      <xsl:if test="draw:line/office:event-listeners">
        <xsl:for-each select="draw:line">
          <xsl:variable name="ShapePostionCount">
            <xsl:value-of select="position()"/>
          </xsl:variable>
          <xsl:for-each select ="office:event-listeners/presentation:event-listener">
            <xsl:if test="@xlink:href">
              <Relationship>
                <xsl:attribute name="Id">
                  <xsl:value-of select="concat('LineFileId',$ShapePostionCount)"/>
                </xsl:attribute>
                <xsl:choose>
                  <xsl:when test="@xlink:href[contains(.,'#page')]">
                    <xsl:attribute name="Type">
                      <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
                    </xsl:attribute>
                    <xsl:attribute name="Target">
                      <xsl:value-of select="concat('slide',substring-after(@xlink:href,'page'),'.xml')"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:when test="@xlink:href">
                    <xsl:attribute name="Type">
                      <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                    </xsl:attribute>
                    <xsl:attribute name="Target">
                      <xsl:if test="@xlink:href[ contains(.,'./')]">
                        <xsl:if test="string-length(substring-after(@xlink:href, '../')) = 0">
                          <xsl:value-of select="/"/>
                        </xsl:if>
                        <xsl:if test="not(string-length(substring-after(@xlink:href, '../')) = 0)">
                          <xsl:value-of select="substring-after(@xlink:href, '../')"/>
                        </xsl:if>
                      </xsl:if>
                      <xsl:if test="not(@xlink:href[ contains(.,'./')])">
                        <xsl:value-of select="concat('file://',@xlink:href)"/>
                      </xsl:if>
                    </xsl:attribute>
                    <xsl:attribute name="TargetMode">
                      <xsl:value-of select="'External'"/>
                    </xsl:attribute>
                  </xsl:when>
                </xsl:choose>
              </Relationship>
            </xsl:if>
            <xsl:if test="presentation:sound">
              <xsl:variable name="varblnDuplicateRelation">
                <xsl:call-template name="GetUniqueRelationIdForWavFile">
                  <xsl:with-param name="FilePath" select="presentation:sound/@xlink:href" />
                  <xsl:with-param name="ShapePosition" select="$ShapePostionCount" />
                  <xsl:with-param name="ShapeType" select="'LINE'" />
                  <xsl:with-param name="Page" select="parent::node()/parent::node()/parent::node()" />
                </xsl:call-template>
              </xsl:variable>
              <xsl:if test="$varblnDuplicateRelation != 1">
                <Relationship>
                  <xsl:variable name="varMediaFilePath">
                    <xsl:if test="presentation:sound/@xlink:href [ contains(.,'../')]">
                      <xsl:value-of select="presentation:sound/@xlink:href" />
                    </xsl:if>
                    <xsl:if test="not(presentation:sound/@xlink:href [ contains(.,'../')])">
                      <xsl:value-of select="substring-after(presentation:sound/@xlink:href,'/')" />
                    </xsl:if>
                  </xsl:variable>
                  <xsl:variable name="varFileRelId">
                    <xsl:value-of select="translate(translate(translate(translate(translate($varMediaFilePath,'/','_'),'..','_'),'.','_'),':','_'),'%20D','_')"/>
                  </xsl:variable>
                  <xsl:attribute name="Id">
                    <xsl:value-of select="$varFileRelId"/>
                  </xsl:attribute>
                  <!--<xsl:attribute name="Id">
                  <xsl:value-of select="concat('LineFileId',$ShapePostionCount)"/>
                </xsl:attribute>-->
                  <xsl:attribute name="Type">
                    <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio'"/>
                  </xsl:attribute>
                  <xsl:attribute name="Target">
                    <xsl:value-of select="concat('../media/',$varFileRelId,'.wav')"/>
                  </xsl:attribute>
                </Relationship>
              </xsl:if>
            </xsl:if>
          </xsl:for-each>
        </xsl:for-each>
      </xsl:if>
      <xsl:if test="draw:frame/office:event-listeners">
        <!-- For TextBoxes-->
        <xsl:for-each select="draw:frame">
          <xsl:variable name="PresentationClass">
            <xsl:if test="@presentation:class">
              <xsl:value-of select="1"/>
            </xsl:if>
            <xsl:if test="not(@presentation:class)">
              <xsl:value-of select="0"/>
            </xsl:if>
          </xsl:variable>
          <xsl:if test="$PresentationClass = 0">
            <xsl:variable name="PostionCount">
              <xsl:value-of select="position()"/>
            </xsl:variable>
            <xsl:for-each select ="office:event-listeners/presentation:event-listener">
              <xsl:if test="@xlink:href">
                <Relationship>
                  <xsl:attribute name="Id">
                    <xsl:value-of select="concat('TxtBoxAtchFileId',$PostionCount)"/>
                  </xsl:attribute>
                  <xsl:choose>
                    <xsl:when test="@xlink:href[contains(.,'#page')]">
                      <xsl:attribute name="Type">
                        <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
                      </xsl:attribute>
                      <xsl:attribute name="Target">
                        <xsl:value-of select="concat('slide',substring-after(@xlink:href,'page'),'.xml')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test="@xlink:href">
                      <xsl:attribute name="Type">
                        <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                      </xsl:attribute>
                      <xsl:attribute name="Target">
                        <xsl:if test="@xlink:href[ contains(.,'./')]">
                          <xsl:if test="string-length(substring-after(@xlink:href, '../')) = 0">
                            <xsl:value-of select="/"/>
                          </xsl:if>
                          <xsl:if test="not(string-length(substring-after(@xlink:href, '../')) = 0)">
                            <xsl:value-of select="substring-after(@xlink:href, '../')"/>
                          </xsl:if>
                        </xsl:if>
                        <xsl:if test="not(@xlink:href[ contains(.,'./')])">
                          <xsl:value-of select="concat('file://',@xlink:href)"/>
                        </xsl:if>
                      </xsl:attribute>
                      <xsl:attribute name="TargetMode">
                        <xsl:value-of select="'External'"/>
                      </xsl:attribute>
                    </xsl:when>
                  </xsl:choose>
                </Relationship>
              </xsl:if>
              <xsl:if test="presentation:sound">
                <xsl:variable name="varblnDuplicateRelation">
                  <xsl:call-template name="GetUniqueRelationIdForWavFile">
                    <xsl:with-param name="FilePath" select="presentation:sound/@xlink:href" />
                    <xsl:with-param name="ShapePosition" select="$PostionCount" />
                    <xsl:with-param name="ShapeType" select="'FRAME'" />
                    <xsl:with-param name="Page" select="parent::node()/parent::node()/parent::node()" />
                  </xsl:call-template>
                </xsl:variable>
                <xsl:if test="$varblnDuplicateRelation != 1">
                  <Relationship>
                    <xsl:variable name="varMediaFilePath">
                      <xsl:if test="presentation:sound/@xlink:href [ contains(.,'../')]">
                        <xsl:value-of select="presentation:sound/@xlink:href" />
                      </xsl:if>
                      <xsl:if test="not(presentation:sound/@xlink:href [ contains(.,'../')])">
                        <xsl:value-of select="substring-after(presentation:sound/@xlink:href,'/')" />
                      </xsl:if>
                    </xsl:variable>
                    <xsl:variable name="varFileRelId">
                      <xsl:value-of select="translate(translate(translate(translate(translate($varMediaFilePath,'/','_'),'..','_'),'.','_'),':','_'),'%20D','_')"/>
                    </xsl:variable>
                    <xsl:attribute name="Id">
                      <xsl:value-of select="$varFileRelId"/>
                    </xsl:attribute>
                    <!--<xsl:attribute name="Id">
                    <xsl:value-of select="concat('TxtBoxAtchFileId',$PostionCount)"/>
                  </xsl:attribute>-->
                    <xsl:attribute name="Type">
                      <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio'"/>
                    </xsl:attribute>
                    <xsl:attribute name="Target">
                      <xsl:value-of select="concat('../media/',$varFileRelId,'.wav')"/>
                    </xsl:attribute>
                  </Relationship>
                </xsl:if>
              </xsl:if>
            </xsl:for-each>
          </xsl:if>
        </xsl:for-each>
        <!-- For frames-->
        <xsl:for-each select="draw:frame/@presentation:class">
          <xsl:variable name="PostionCount">
            <xsl:value-of select="position()"/>
          </xsl:variable>
          <xsl:for-each select ="parent::node()/office:event-listeners/presentation:event-listener">
            <xsl:if test="@xlink:href">
              <Relationship>
                <xsl:attribute name="Id">
                  <xsl:value-of select="concat('AtchFileId',$PostionCount)"/>
                </xsl:attribute>
                <xsl:choose>
                  <xsl:when test="@xlink:href[contains(.,'#page')]">
                    <xsl:attribute name="Type">
                      <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
                    </xsl:attribute>
                    <xsl:attribute name="Target">
                      <xsl:value-of select="concat('slide',substring-after(@xlink:href,'page'),'.xml')"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:when test="@xlink:href">
                    <xsl:attribute name="Type">
                      <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                    </xsl:attribute>
                    <xsl:attribute name="Target">
                      <xsl:if test="@xlink:href[ contains(.,'./')]">
                        <xsl:if test="string-length(substring-after(@xlink:href, '../')) = 0">
                          <xsl:value-of select="/"/>
                        </xsl:if>
                        <xsl:if test="not(string-length(substring-after(@xlink:href, '../')) = 0)">
                          <xsl:value-of select="substring-after(@xlink:href, '../')"/>
                        </xsl:if>
                      </xsl:if>
                      <xsl:if test="not(@xlink:href[ contains(.,'./')])">
                        <xsl:value-of select="concat('file://',@xlink:href)"/>
                      </xsl:if>
                    </xsl:attribute>
                    <xsl:attribute name="TargetMode">
                      <xsl:value-of select="'External'"/>
                    </xsl:attribute>
                  </xsl:when>
                </xsl:choose>
              </Relationship>
            </xsl:if>
            <xsl:if test="presentation:sound">
              <xsl:variable name="varblnDuplicateRelation">
                <xsl:call-template name="GetUniqueRelationIdForWavFile">
                  <xsl:with-param name="FilePath" select="presentation:sound/@xlink:href" />
                  <xsl:with-param name="ShapePosition" select="$PostionCount" />
                  <xsl:with-param name="ShapeType" select="'FRAME'" />
                  <xsl:with-param name="Page" select="parent::node()/parent::node()/parent::node()" />
                </xsl:call-template>
              </xsl:variable>
              <xsl:if test="$varblnDuplicateRelation != 1">
                <Relationship>
                  <xsl:variable name="varMediaFilePath">
                    <xsl:if test="presentation:sound/@xlink:href [ contains(.,'../')]">
                      <xsl:value-of select="presentation:sound/@xlink:href" />
                    </xsl:if>
                    <xsl:if test="not(presentation:sound/@xlink:href [ contains(.,'../')])">
                      <xsl:value-of select="substring-after(presentation:sound/@xlink:href,'/')" />
                    </xsl:if>
                  </xsl:variable>
                  <xsl:variable name="varFileRelId">
                    <xsl:value-of select="translate(translate(translate(translate(translate($varMediaFilePath,'/','_'),'..','_'),'.','_'),':','_'),'%20D','_')"/>
                  </xsl:variable>
                  <xsl:attribute name="Id">
                    <xsl:value-of select="$varFileRelId"/>
                  </xsl:attribute>
                  <xsl:attribute name="Type">
                    <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio'"/>
                  </xsl:attribute>
                  <xsl:attribute name="Target">
                    <xsl:value-of select="concat('../media/',$varFileRelId,'.wav')"/>
                  </xsl:attribute>
                </Relationship>
              </xsl:if>
            </xsl:if>
          </xsl:for-each>
        </xsl:for-each>
      </xsl:if>
      <xsl:if test="draw:connector/office:event-listeners">
        <xsl:for-each select="draw:connector">
          <xsl:variable name="ShapePostionCount">
            <xsl:value-of select="position()"/>
          </xsl:variable>
          <xsl:for-each select ="office:event-listeners/presentation:event-listener">
            <xsl:if test="@xlink:href">
              <Relationship>
                <xsl:attribute name="Id">
                  <xsl:value-of select="concat('LineFileId',$ShapePostionCount)"/>
                </xsl:attribute>
                <xsl:choose>
                  <xsl:when test="@xlink:href[contains(.,'#page')]">
                    <xsl:attribute name="Type">
                      <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
                    </xsl:attribute>
                    <xsl:attribute name="Target">
                      <xsl:value-of select="concat('slide',substring-after(@xlink:href,'page'),'.xml')"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:when test="@xlink:href">
                    <xsl:attribute name="Type">
                      <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
                    </xsl:attribute>
                    <xsl:attribute name="Target">
                      <xsl:if test="@xlink:href[ contains(.,'./')]">
                        <xsl:if test="string-length(substring-after(@xlink:href, '../')) = 0">
                          <xsl:value-of select="/"/>
                        </xsl:if>
                        <xsl:if test="not(string-length(substring-after(@xlink:href, '../')) = 0)">
                          <xsl:value-of select="substring-after(@xlink:href, '../')"/>
                        </xsl:if>
                      </xsl:if>
                      <xsl:if test="not(@xlink:href[ contains(.,'./')])">
                        <xsl:value-of select="concat('file://',@xlink:href)"/>
                      </xsl:if>
                    </xsl:attribute>
                    <xsl:attribute name="TargetMode">
                      <xsl:value-of select="'External'"/>
                    </xsl:attribute>
                  </xsl:when>
                </xsl:choose>
              </Relationship>
            </xsl:if>
            <xsl:if test="presentation:sound">
              <xsl:variable name="varblnDuplicateRelation">
                <xsl:call-template name="GetUniqueRelationIdForWavFile">
                  <xsl:with-param name="FilePath" select="presentation:sound/@xlink:href" />
                  <xsl:with-param name="ShapePosition" select="$ShapePostionCount" />
                  <xsl:with-param name="ShapeType" select="'CONNECTOR'" />
                  <xsl:with-param name="Page" select="parent::node()/parent::node()/parent::node()" />
                </xsl:call-template>
              </xsl:variable>
              <xsl:if test="$varblnDuplicateRelation != 1">
                <Relationship>
                  <xsl:variable name="varMediaFilePath">
                    <xsl:if test="presentation:sound/@xlink:href [ contains(.,'../')]">
                      <xsl:value-of select="presentation:sound/@xlink:href" />
                    </xsl:if>
                    <xsl:if test="not(presentation:sound/@xlink:href [ contains(.,'../')])">
                      <xsl:value-of select="substring-after(presentation:sound/@xlink:href,'/')" />
                    </xsl:if>
                  </xsl:variable>
                  <xsl:variable name="varFileRelId">
                    <xsl:value-of select="translate(translate(translate(translate(translate($varMediaFilePath,'/','_'),'..','_'),'.','_'),':','_'),'%20D','_')"/>
                  </xsl:variable>
                  <xsl:attribute name="Id">
                    <xsl:value-of select="$varFileRelId"/>
                  </xsl:attribute>
                  <!--<xsl:attribute name="Id">
                  <xsl:value-of select="concat('LineFileId',$ShapePostionCount)"/>
                </xsl:attribute>-->
                  <xsl:attribute name="Type">
                    <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio'"/>
                  </xsl:attribute>
                  <xsl:attribute name="Target">
                    <xsl:value-of select="concat('../media/',$varFileRelId,'.wav')"/>
                  </xsl:attribute>
                </Relationship>
              </xsl:if>
            </xsl:if>
          </xsl:for-each>
        </xsl:for-each>
      </xsl:if>
      <!-- End - Add relation files of hyperlinks-->
      <!--Code inserted by Vijayeta for Bullets and numbering,set relatioship for pictures-->
      <!--<xsl:for-each select ="document('content.xml')//office:document-content/office:automatic-styles/text:list-style/text:list-level-style-image">
        <xsl:if test ="@text:level='1'">
          <Relationship>
            <xsl:attribute name="Id">
              <xsl:value-of select="$rId"/>
            </xsl:attribute>
            <xsl:attribute name="Type">
              <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'"/>
            </xsl:attribute>
            <xsl:attribute name="Target">
              <xsl:value-of select="concat('../media/',substring-after(@xlink:href,'Pictures/'))"/>
            </xsl:attribute>
          </Relationship>
        </xsl:if>
      </xsl:for-each>-->
      <!--End of Code inserted by Vijayeta for Bullets and numbering,set relatioship for pictures-->
		</Relationships>
	</xsl:template>


  <xsl:template name="GetUniqueRelationIdForWavFile">
    <xsl:param name ="FilePath"/>
    <xsl:param name ="ShapePosition"/>
    <xsl:param name="Page" />
    <xsl:param name="ShapeType" />
    <xsl:for-each select="$Page">
      <xsl:if test="$ShapeType = 'CUSTOM' ">
        <xsl:for-each select="./draw:custom-shape">
          <xsl:if test="position() &lt; $ShapePosition">
            <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
              <xsl:value-of select='1'/>
            </xsl:if>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
      <xsl:if test="$ShapeType = 'CUSTOM' and $ShapePosition = '1'" >
        <xsl:value-of select='0'/>
      </xsl:if>
      <xsl:if test="$ShapeType = 'RECT'">
        <xsl:for-each select="./draw:custom-shape">
          <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
            <xsl:value-of select='1'/>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
      <xsl:if test="$ShapeType = 'LINE'">
        <xsl:for-each select="./draw:custom-shape">
          <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)]">
            <xsl:value-of select='1'/>
          </xsl:if>
        </xsl:for-each>
        <xsl:for-each select="./draw:rect">
          <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
            <xsl:value-of select='1'/>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
      <xsl:if test="$ShapeType = 'FRAME'">
        <xsl:for-each select="./draw:custom-shape">
          <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
            <xsl:value-of select='1'/>
          </xsl:if>
        </xsl:for-each>
        <xsl:for-each select="./draw:rect">
          <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)]">
            <xsl:value-of select='1'/>
          </xsl:if>
        </xsl:for-each>
        <xsl:for-each select="./draw:line">
          <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
            <xsl:value-of select='1'/>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
      <xsl:if test="$ShapeType = 'CONNECTOR'">
        <xsl:for-each select="./draw:custom-shape">
          <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
            <xsl:value-of select='1'/>
          </xsl:if>
        </xsl:for-each>
        <xsl:for-each select="./draw:rect">
          <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
            <xsl:value-of select='1'/>
          </xsl:if>
        </xsl:for-each>
        <xsl:for-each select="./draw:line">
          <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
            <xsl:value-of select='1'/>
          </xsl:if>
        </xsl:for-each>
        <xsl:for-each select="./draw:frame">
          <xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
            <xsl:value-of select='1'/>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  <!-- Templates added by vijayeta,to get paragraph style name for each of the levels in multilevelled list-->
  <xsl:template name ="getParaStyleName">
    <xsl:param name ="lvl"/>
    <xsl:choose>
      <xsl:when test ="$lvl='0'">
        <xsl:value-of select ="text:list-item/text:p/@text:style-name"/>
      </xsl:when >
      <xsl:when test ="$lvl='1'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='2'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='3'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='4'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='5'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='6'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='7'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='8'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='9'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name ="getTextNodeForFontStyle">
    <xsl:param name ="this"/>
    <xsl:choose>
      <xsl:when test ="./text:p">
        <xsl:choose >
          <xsl:when test ="./text:p/text:span">
            <xsl:value-of select ="./text:p/text:span/@text:style-name"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select ="./text:p/@text:style-name"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:p">
        <xsl:choose >
          <xsl:when test ="./text:list/text:list-item/text:p/text:span">
            <xsl:value-of select ="./text:list/text:list-item/text:p/text:span/@text:style-name"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select ="./text:list/text:list-item/text:p/@text:style-name"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:choose >
          <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:p/text:span">
            <xsl:value-of select ="./text:list/text:list-item/text:list/text:list-item/text:p/text:span/@text:style-name"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select ="./text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:choose >
          <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/text:span">
            <xsl:value-of select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/text:span/@text:style-name"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:choose >
          <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/text:span">
            <xsl:value-of select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/text:span/@text:style-name"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:choose >
          <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/text:span">
            <xsl:value-of select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/text:span/@text:style-name"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:choose >
          <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/text:span">
            <xsl:value-of select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/text:span/@text:style-name"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:choose >
          <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/text:span">
            <xsl:value-of select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/text:span/@text:style-name"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:choose >
          <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/text:span">
            <xsl:value-of select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/text:span/@text:style-name"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!--End,to get paragraph style name for each of the levels in multilevelled list-->

</xsl:stylesheet>
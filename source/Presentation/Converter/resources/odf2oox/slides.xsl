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
	<xsl:import href ="picture.xsl"/>
	<xsl:import href ="customAnimation.xsl"/>
	<xsl:template name ="slides" match ="/office:document-content/office:body/office:presentation/draw:page" mode="slide">
		<xsl:param name ="pageNo"/>
		<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
				   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
				   xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
			<p:cSld>
				<!--Added by sateesh - Background Color-->
				<xsl:variable name="dpName">
					<xsl:value-of select="@draw:style-name" />
				</xsl:variable>
				<xsl:for-each select="document('content.xml')//style:style[@style:name= $dpName]/style:drawing-page-properties">
					<xsl:choose>
						<xsl:when test="@draw:fill='bitmap'">
							<p:bg>
								<p:bgPr>
									<a:blipFill dpi="0" rotWithShape="1">
										<xsl:call-template name="tmpInsertBackImage">
											<xsl:with-param name="FileName" select="concat('slide',$pageNo)" />
											<xsl:with-param name="imageName" select="@draw:fill-image-name" />
										</xsl:call-template>
										<xsl:if test="@style:repeat='stretch'">
											<a:stretch>
												<a:fillRect />
											</a:stretch>
										</xsl:if>
										<xsl:if test="@draw:fill-image-ref-point-x or @draw:fill-image-ref-point-y">
											<a:tile tx="0" ty="0" flip="none">
												<xsl:if test="@draw:fill-image-ref-point-x">
													<xsl:attribute name="sx">
														<xsl:value-of select="number(substring-before(@draw:fill-image-ref-point-x,'%')) * 1000"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test="@draw:fill-image-ref-point-y">
													<xsl:attribute name="sy">
														<xsl:value-of select="number(substring-before(@draw:fill-image-ref-point-y,'%')) * 1000"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test="@draw:fill-image-ref-point">
													<xsl:attribute name="algn">
														<xsl:choose>
															<xsl:when test="@draw:fill-image-ref-point='top-left'">
																<xsl:value-of select ="'tl'"/>
															</xsl:when>
															<xsl:when test="@draw:fill-image-ref-point='top'">
																<xsl:value-of select ="'t'"/>
															</xsl:when>
															<xsl:when test="@draw:fill-image-ref-point='top-right'">
																<xsl:value-of select ="'tr'"/>
															</xsl:when>
															<xsl:when test="@draw:fill-image-ref-point='right'">
																<xsl:value-of select ="'r'"/>
															</xsl:when>
															<xsl:when test="@draw:fill-image-ref-point='bottom-left'">
																<xsl:value-of select ="'bl'"/>
															</xsl:when>
															<xsl:when test="@draw:fill-image-ref-point='bottom-right'">
																<xsl:value-of select ="'br'"/>
															</xsl:when>
															<xsl:when test="@draw:fill-image-ref-point='bottom'">
																<xsl:value-of select ="'b'"/>
															</xsl:when>
															<xsl:when test="@draw:fill-image-ref-point='center'">
																<xsl:value-of select ="'ctr'"/>
															</xsl:when>
														</xsl:choose>
													</xsl:attribute>
												</xsl:if>
											</a:tile>
										</xsl:if>
									</a:blipFill>
									<a:effectLst />
								</p:bgPr>
							</p:bg>
						</xsl:when>
						<xsl:when test="@draw:fill='solid'">
							<p:bg>
								<p:bgPr>
									<a:solidFill>
										<a:srgbClr>
											<xsl:attribute name="val">
												<xsl:if test="@draw:fill-color">
													<xsl:value-of select="substring-after(@draw:fill-color,'#')" />
												</xsl:if>
												<xsl:if test="not(@draw:fill-color)">
													<xsl:value-of select="'ffffff'" />
												</xsl:if>
											</xsl:attribute>
										</a:srgbClr>
									</a:solidFill>
									<a:effectLst />
								</p:bgPr>
							</p:bg>
						</xsl:when>
            <xsl:when test="@draw:fill='gradient'">
              <xsl:variable name="gradStyleName" select="@draw:fill-gradient-name"/>
              <p:bg>
                <p:bgPr>
                  <a:solidFill>
                    <a:srgbClr>
                      <xsl:for-each select="document('styles.xml')//draw:gradient[@draw:name= $gradStyleName]">
                        <xsl:attribute name="val">
                          <xsl:if test="@draw:start-color">
                            <xsl:value-of select="substring-after(@draw:start-color,'#')" />
                          </xsl:if>
                          <xsl:if test="not(@draw:start-color)">
                            <xsl:value-of select="'ffffff'" />
                          </xsl:if>
                        </xsl:attribute>
                      </xsl:for-each>
                    </a:srgbClr>
                  </a:solidFill>
                  <a:effectLst />
                </p:bgPr>
              </p:bg>
            </xsl:when>
					</xsl:choose>
				</xsl:for-each>
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
					<xsl:variable name ="pageStyle">
						<xsl:value-of select ="@draw:style-name"/>
					</xsl:variable>
					<xsl:variable name ="footerId">
						<xsl:value-of select ="@presentation:use-footer-name"/>
					</xsl:variable>
					<xsl:variable name ="dateId">
						<xsl:value-of select ="@presentation:use-date-time-name"/>
					</xsl:variable>
          <xsl:variable name ="SMName">
            <xsl:value-of select ="@draw:master-page-name"/>
          </xsl:variable>
          
					<xsl:for-each select ="document('content.xml')//style:style[@style:name=$pageStyle]">
						<xsl:if test ="style:drawing-page-properties[@presentation:display-footer='true']">
							<!--<xsl:if test ="not($footerId ='')">-->
							<!--<xsl:if test ="document('content.xml')//presentation:footer-decl[@presentation:name=$footerId]">-->
								<xsl:call-template name ="footer" >
									<xsl:with-param name ="footerId" select ="$footerId"/>
                  <xsl:with-param name ="SMName" select ="$SMName"/>
								</xsl:call-template >
							<!--</xsl:if>-->
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
					<xsl:for-each select="node()">
						<xsl:choose>
							<xsl:when test="name()='draw:frame'">
								<xsl:variable name="var_pos" select="position()"/>
              
								<!--<xsl:for-each select=".">-->
									<xsl:choose>
										<xsl:when test="./draw:image">
											<xsl:for-each select="./draw:image">
												<xsl:if test ="contains(@xlink:href,'.png') or contains(@xlink:href,'.emf') or contains(@xlink:href,'.wmf') or contains(@xlink:href,'.jfif') or contains(@xlink:href,'.jpe') 
            or contains(@xlink:href,'.bmp') or contains(@xlink:href,'.dib') or contains(@xlink:href,'.rle')
            or contains(@xlink:href,'.bmz') or contains(@xlink:href,'.gfa') 
            or contains(@xlink:href,'.emz') or contains(@xlink:href,'.wmz') or contains(@xlink:href,'.pcz')
            or contains(@xlink:href,'.tif') or contains(@xlink:href,'.tiff') 
            or contains(@xlink:href,'.cdr') or contains(@xlink:href,'.cgm') or contains(@xlink:href,'.eps') 
            or contains(@xlink:href,'.pct') or contains(@xlink:href,'.pict') or contains(@xlink:href,'.wpg') 
            or contains(@xlink:href,'.jpeg') or contains(@xlink:href,'.gif') or contains(@xlink:href,'.png') or contains(@xlink:href,'.jpg')">
                         <xsl:if test="not(./@xlink:href[contains(.,'../')])">
													<xsl:call-template name="InsertPicture">
														<xsl:with-param name="imageNo" select="$pageNo" />
														<xsl:with-param name="picNo" select="$var_pos" />
													</xsl:call-template>
												</xsl:if>
											</xsl:if>
											</xsl:for-each>
										</xsl:when>
										<xsl:when test="@presentation:class[contains(.,'title')]
                                  or @presentation:class[contains(.,'subtitle')]
                                  or @presentation:class[contains(.,'outline')]">
											<xsl:variable name ="masterPageName" select ="./parent::node()/@draw:master-page-name"/>
											<xsl:variable name="FrameCount" select="concat('Frame',$var_pos)"/>
											<p:sp>
												<p:nvSpPr>
													<xsl:choose>
														<xsl:when test ="@presentation:class[contains(.,'subtitle')]">
															<p:cNvPr name="subTitle 1">
																<xsl:attribute name="id">
																	<xsl:value-of select="$var_pos+1"/>
																</xsl:attribute>
															</p:cNvPr>
															<p:cNvSpPr>
																<a:spLocks noGrp="1"/>
															</p:cNvSpPr>
															<p:nvPr>
																<p:ph type="subTitle" idx="1"/>
															</p:nvPr>
														</xsl:when>
														<xsl:when test ="@presentation:class[contains(.,'title')]">
															<p:cNvPr name="Title 1">
																<xsl:attribute name="id">
																	<xsl:value-of select="$var_pos+1"/>
																</xsl:attribute>
															</p:cNvPr>
															<p:cNvSpPr>
																<a:spLocks noGrp="1"/>
															</p:cNvSpPr>
															<p:nvPr>
																<p:ph type="ctrTitle"/>
															</p:nvPr>
														</xsl:when>
														<xsl:when test ="@presentation:class[contains(.,'outline')]">
															<p:cNvPr name="Content Placeholder 2">
																<xsl:attribute name ="id">
																	<xsl:value-of select ="$var_pos+1"/>
																</xsl:attribute>
															</p:cNvPr>
															<p:cNvSpPr>
																<a:spLocks noGrp="1"/>
															</p:cNvSpPr>
															<p:nvPr>
																<p:ph idx="1"/>
															</p:nvPr>
														</xsl:when>
													</xsl:choose>
												</p:nvSpPr>
												<p:spPr>
													<!--Edited by Vipul to add Rotation-->
													<!--Start-->
													<a:xfrm>
														<xsl:variable name ="angle">
															<xsl:value-of select="substring-after(substring-before(substring-before(@draw:transform,'translate'),')'),'(')" />
														</xsl:variable>
														<xsl:variable name ="x2">
															<xsl:value-of select="substring-before(substring-before(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
														</xsl:variable>
														<xsl:variable name ="y2">
															<xsl:value-of select="substring-before(substring-after(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
														</xsl:variable>
														<xsl:if test="@draw:transform">
															<xsl:attribute name ="rot">
																<xsl:value-of select ="concat('draw-transform:ROT:',substring-before(@svg:width,'cm'), ':',
																   substring-before(@svg:height,'cm'), ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
															</xsl:attribute>
														</xsl:if>
														<a:off>
															<xsl:if test="@draw:transform">
																<xsl:attribute name ="x">
																	<xsl:value-of select ="concat('draw-transform:X:',substring-before(@svg:width,'cm'), ':',
																   substring-before(@svg:height,'cm'), ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
																</xsl:attribute>

																<xsl:attribute name ="y">
																	<xsl:value-of select ="concat('draw-transform:Y:',substring-before(@svg:width,'cm'), ':',
																   substring-before(@svg:height,'cm'), ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
																</xsl:attribute>
															</xsl:if>
															<xsl:if test="not(@draw:transform)">
																<xsl:attribute name ="x">
																	<xsl:call-template name ="convertToPoints">
																		<xsl:with-param name ="unit" select ="'cm'"/>
																		<xsl:with-param name ="length" select ="@svg:x"/>
																	</xsl:call-template>
																</xsl:attribute>
																<xsl:attribute name ="y">
																	<xsl:call-template name ="convertToPoints">
																		<xsl:with-param name ="unit" select ="'cm'"/>
																		<xsl:with-param name ="length" select ="@svg:y"/>
																	</xsl:call-template>
																</xsl:attribute>
															</xsl:if>
														</a:off>
														<a:ext>
															<xsl:attribute name ="cx">
																<xsl:call-template name ="convertToPoints">
																	<xsl:with-param name ="unit" select ="'cm'"/>
																	<xsl:with-param name ="length" select ="@svg:width"/>
																</xsl:call-template>
															</xsl:attribute>
															<xsl:attribute name ="cy">
																<xsl:call-template name ="convertToPoints">
																	<xsl:with-param name ="unit" select ="'cm'"/>
																	<xsl:with-param name ="length" select ="@svg:height"/>
																</xsl:call-template>
															</xsl:attribute>
														</a:ext>
													</a:xfrm>
													<!--End-->
													<!-- Solid fill color -->
													<xsl:call-template name ="fillColor" >
														<xsl:with-param name ="prId" select ="@presentation:style-name" />
													</xsl:call-template>
													<!-- added by Vipul to insert line style-->
													<!--start-->
													<xsl:variable name="var_PrStyleId" select="@presentation:style-name"/>
													<xsl:for-each select ="document('content.xml')/office:document-content/office:automatic-styles/style:style[@style:name=$var_PrStyleId]/style:graphic-properties">
														<xsl:call-template name ="getFillColor"/>
														<xsl:call-template name ="getLineStyle"/>
													</xsl:for-each>
													<!--End-->
												</p:spPr>
												<p:txBody>
													<xsl:call-template name ="TextAlignment" >
														<xsl:with-param name ="prId" select ="@presentation:style-name"/>
													</xsl:call-template >
													<a:lstStyle/>
													<xsl:call-template name ="processText" >
														<xsl:with-param name ="layoutName" select ="@presentation:class"/>
														<xsl:with-param name ="FrameCount" select ="$FrameCount"/>
														<!-- Paremeter added by vijayeta,get master page name, dated:13-7-07-->
														<xsl:with-param name ="masterPageName" select ="$masterPageName"/>
													</xsl:call-template >
												</p:txBody>
											</p:sp>
										</xsl:when>
										<xsl:when test ="(draw:text-box) and not(@presentation:style-name) and not(@presentation:class)">
											<xsl:call-template name ="CreateShape">
												<xsl:with-param name ="fileName" select ="'content.xml'"/>
												<xsl:with-param name ="shapeName" select="'TextBox '" />
												<xsl:with-param name ="shapeCount" select="$var_pos" />
											</xsl:call-template>
										</xsl:when>
										<xsl:when test="./draw:plugin">
											<xsl:for-each select="./draw:plugin">
												<xsl:call-template name="InsertAudio">
													<xsl:with-param name="imageNo" select="$pageNo" />
													<xsl:with-param name="AudNo" select="$var_pos" />
												</xsl:call-template>
											</xsl:for-each>
										</xsl:when>
									</xsl:choose>
								<!--</xsl:for-each>-->
							</xsl:when>
              <xsl:when test="name()='draw:g'">
                <xsl:variable name="var_pos" select="position()"/>
                <!--<xsl:variable name="var_pos" select="number"/>-->
                <!--<xsl:for-each select=".">-->
                <xsl:call-template name="tmpGroping">
                  <xsl:with-param name="pageNo" select="$pageNo"/>
                  <xsl:with-param name="pos" select="$var_pos"/>
              </xsl:call-template>
                <!--</xsl:for-each>-->
							</xsl:when>
              <xsl:when test="name()='draw:rect' or name()='draw:ellipse' or name()='draw:custom-shape'
              or name()='draw:line' or name()='draw:connector'">
								<!-- Code for shapes start-->
								<xsl:variable name="var_pos" select="position()"/>
								<!--<xsl:for-each select=".">-->
								<xsl:call-template name ="shapes" >
									<xsl:with-param name ="fileName" select ="'content.xml'"/>
									<xsl:with-param name ="var_pos" select="$var_pos" />
									<!-- Code for shapes End-->
								</xsl:call-template >
								<!--</xsl:for-each>-->
              </xsl:when>
						</xsl:choose>
					</xsl:for-each>
				</p:spTree>
			</p:cSld>
			<p:clrMapOvr>
				<a:masterClrMapping/>
			</p:clrMapOvr>
			<xsl:call-template name ="customAnimation"/>
		</p:sld>
	</xsl:template>
	<xsl:template name ="processText" >
		<xsl:param name ="layoutName"/>
		<xsl:param name ="FrameCount"/>
		<!-- Paremeter added by vijayeta,get master page name, dated:13-7-07-->
		<xsl:param name ="masterPageName"/>
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
										<!--<xsl:when test="@xlink:href[ contains(.,'#page')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinksldjump'"/>
                      </xsl:attribute>
                    </xsl:when>-->
										<xsl:when test="@xlink:href[ contains(.,'#')] and string-length(substring-before(@xlink:href,'#')) = 0 ">
											<xsl:variable name="pageID">
												<xsl:call-template name="getThePageId">
													<xsl:with-param name="PageName" select="substring-after(@xlink:href,'#')"/>
												</xsl:call-template>
											</xsl:variable>
											<xsl:if test="$pageID > 0">
												<xsl:attribute name="action">
													<xsl:value-of select="'ppaction://hlinksldjump'"/>
												</xsl:attribute>
												<xsl:attribute name="r:id">
													<xsl:value-of select="concat('AtchFileId',position())"/>
												</xsl:attribute>
											</xsl:if>
										</xsl:when>
										<!-- Go to document -->
										<xsl:when test="@xlink:href and @presentation:action[ contains(.,'show')] ">
											<xsl:if test="not(@xlink:href[ contains(.,'#')])">
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
										<!--<xsl:when test="@xlink:href[contains(.,'#page')]">
                      <xsl:attribute name="r:id">
                        <xsl:value-of select="concat('AtchFileId',$PostionCount)"/>
                      </xsl:attribute>
                    </xsl:when>-->
										<!-- For Go to document -->
										<xsl:when test="@xlink:href and @presentation:action[ contains(.,'show')] ">
											<xsl:if test="not(@xlink:href[ contains(.,'#')])">
												<xsl:attribute name="r:id">
													<xsl:value-of select="concat('AtchFileId',position())"/>
												</xsl:attribute>
											</xsl:if>
										</xsl:when>
										<!-- For Go to Run Programs -->
										<xsl:when test="@xlink:href and @presentation:action[ contains(.,'execute')]">
											<xsl:attribute name="r:id">
												<xsl:value-of select="concat('AtchFileId',position())"/>
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
												<xsl:value-of select="concat('Sound_',position())"/>
											</xsl:attribute>
											<xsl:attribute name="builtIn">
												<xsl:value-of select='1'/>
											</xsl:attribute>
											<pzip:import pzip:source="{$varMediaFilePath}" pzip:target="{concat('ppt/media/',$varFileRelId,'.wav')}" />
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
		<xsl:variable name ="prClsName">
			<xsl:value-of select ="@presentation:class"/>
		</xsl:variable>
		<xsl:choose >
			<xsl:when test ="draw:text-box/text:p/text:span">
				<xsl:for-each select ="node()">
					<!--<xsl:for-each select ="draw:text-box/text:p">-->
					<xsl:for-each select ="child::node()[position()]">
						<xsl:choose >
							<xsl:when test ="name()='text:p'">
                <xsl:variable name="linkCount">
                  <xsl:value-of select="position()"/>
                </xsl:variable>
                <xsl:variable name ="paraId" select ="@text:style-name"/>
                <xsl:variable name ="isNumberingEnabled">
                  <xsl:choose >
                    <xsl:when test ="document('content.xml')//style:style[@style:name=$paraId]/style:paragraph-properties/@text:enable-numbering">
                      <xsl:value-of select ="document('content.xml')//style:style[@style:name=$paraId]/style:paragraph-properties/@text:enable-numbering"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select ="'false'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:variable>
								<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
									<xsl:call-template name ="paraProperties" >
										<xsl:with-param name ="paraId" >
											<xsl:value-of select ="@text:style-name"/>
										</xsl:with-param >
										<xsl:with-param name ="isBulleted" select ="'false'"/>
										<xsl:with-param name ="level" select ="'0'"/>
										<xsl:with-param name="framePresentaionStyleId" select="parent::node()/parent::node()/./@presentation:style-name" />
                    <xsl:with-param name ="isNumberingEnabled" select ="$isNumberingEnabled"/>
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
															<xsl:if test="./text:a">
                                <xsl:call-template name="TextActionHyperLinks">
                                    <xsl:with-param name="PostionCount" select="$PostionCount"/>
                                    <xsl:with-param name="nodeType" select="'span'"/>
                                  <xsl:with-param name="linkCount" select="$linkCount"/>
                                  </xsl:call-template>
															</xsl:if>
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
														<xsl:if test="./text:a">
                                <xsl:call-template name="TextActionHyperLinks">
                                  <xsl:with-param name="PostionCount" select="$PostionCount"/>
                                  <xsl:with-param name="nodeType" select="'span'"/>
                                  <xsl:with-param name="linkCount" select="$linkCount"/>
                                </xsl:call-template>
														</xsl:if>
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
							</xsl:when>
							<xsl:when test ="name()='text:list'">
								<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
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
										<!--<xsl:if test ="not(document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering)">
											<xsl:value-of select ="'true'"/>
										</xsl:if>-->
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
										<!-- parameter added by vijayeta, dated 11-7-07-->
										<xsl:with-param name ="masterPageName" select ="$masterPageName"/>
									</xsl:call-template >
									<!--End of Code inserted by Vijayets for Bullets and numbering-->
									<xsl:for-each select ="child::node()[position()]">
										<xsl:choose >
											<xsl:when test ="name()='text:list-item'">
												<xsl:variable name ="currentNodeStyle">
													<xsl:call-template name ="getTextNodeForFontStyle">
														<xsl:with-param name ="prClassName" select ="$prClsName"/>
														<xsl:with-param name ="lvl" select ="$lvl"/>
														<xsl:with-param name ="masterPageName" select ="$masterPageName"/>
													</xsl:call-template>
												</xsl:variable>
												<xsl:copy-of select ="$currentNodeStyle"/>
											</xsl:when >
										</xsl:choose>
									</xsl:for-each>
									<xsl:copy-of select="$varFrameHyperLinks"/>
								</a:p >
							</xsl:when>
						</xsl:choose>
					</xsl:for-each>
				</xsl:for-each >
			</xsl:when>
			<xsl:when test ="draw:text-box/text:list/text:list-item">
				<!--<xsl:when test ="draw:text-box/text:list/text:list-item/text:p/text:span">-->
				<xsl:for-each select ="draw:text-box/text:list">
					<xsl:variable name="forCount" select="position()" />
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
							<xsl:choose >
								<xsl:when test ="document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering">
									<xsl:value-of select ="document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering"/>
								</xsl:when>
								<xsl:otherwise>
									<xsl:variable name ="styleNameFromStyles" >
										<xsl:choose >
											<xsl:when test ="$prClsName='subtitle' or $prClsName='title'">
												<xsl:value-of select ="concat($prClsName,'-',$masterPageName)"/>
											</xsl:when>
											<xsl:when test ="$prClsName='outline'">
												<xsl:value-of select ="concat($masterPageName,'-outline',$lvl+1)"/>
											</xsl:when>
										</xsl:choose>
									</xsl:variable>
									<xsl:choose >
										<xsl:when test ="document('styles.xml')//office:styles/style:style[@style:name=$styleNameFromStyles]/style:paragraph-properties/@text:enable-numbering">
											<xsl:value-of select ="document('styles.xml')//office:styles/style:style[@style:name=$styleNameFromStyles]/style:paragraph-properties/@text:enable-numbering"/>
										</xsl:when>
                    <!--<xsl:when test ="not(document('styles.xml')//office:styles/style:style[@style:name=$styleNameFromStyles]/style:paragraph-properties/@text:enable-numbering)">
                      <xsl:value-of select ="'true'"/>
                    </xsl:when>-->
										<!--<xsl:otherwise>
                      <xsl:value-of select ="'true'"/>
                    </xsl:otherwise>-->
									</xsl:choose>
								</xsl:otherwise>
							</xsl:choose>
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
							<!-- parameter added by vijayeta, dated 13-7-07-->
							<xsl:with-param name ="masterPageName" select ="$masterPageName"/>
              <xsl:with-param name ="pos" select ="$forCount"/>
						</xsl:call-template >
						<!--End of Code inserted by Vijayets for Bullets and numbering-->
						<xsl:for-each select ="child::node()[position()]">
							<xsl:choose >
								<xsl:when test ="name()='text:list-item'">
									<!-- Variable "textHyperlinksForBullets" added by lohith for Hyperlinks in Bullets -->
									<xsl:variable name="textHyperlinksForBullets">
										<xsl:call-template name="getHyperlinksForBulltedText">
											<xsl:with-param name="bulletLevel" select="$lvl" />
											<xsl:with-param name="listItemCount" select="$forCount"/>
											<xsl:with-param name="FrameCount" select="$FrameCount"/>
											<xsl:with-param name="rootNode" select="."/>
										</xsl:call-template>
									</xsl:variable>
									<xsl:variable name ="currentNodeStyle">
										<xsl:call-template name ="getTextNodeForFontStyle">
											<xsl:with-param name ="prClassName" select ="$prClsName"/>
											<xsl:with-param name ="lvl" select ="$lvl"/>
											<xsl:with-param name ="HyperlinksForBullets" select ="$textHyperlinksForBullets"/>
											<!-- parameter added by vijayeta, dated 11-7-07-->
											<xsl:with-param name ="masterPageName" select ="$masterPageName"/>
											<xsl:with-param name ="fileName" select ="'content.xml'"/>
										</xsl:call-template>
									</xsl:variable>
									<xsl:copy-of select ="$currentNodeStyle"/>
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
							<!-- parameter added by vijayeta, dated 11-7-07-->
							<xsl:with-param name ="masterPageName" select ="$masterPageName"/>
						</xsl:call-template >
						<!--End of Code inserted by Vijayets for Bullets and numbering-->
						<xsl:for-each select ="child::node()[position()]">
							<xsl:choose >
								<xsl:when test ="name()='text:list-item'">
									<xsl:variable name ="currentNodeStyle">
										<xsl:call-template name ="getTextNodeForFontStyle">
											<xsl:with-param name ="prClassName" select ="$prClsName"/>
											<xsl:with-param name ="lvl" select ="$lvl"/>
											<!-- parameter added by vijayeta, dated 11-7-07-->
											<xsl:with-param name ="masterPageName" select ="$masterPageName"/>
											<xsl:with-param name ="fileName" select ="'content.xml'"/>
										</xsl:call-template>
									</xsl:variable>
									<xsl:copy-of select ="$currentNodeStyle"/>
								</xsl:when >
							</xsl:choose>
						</xsl:for-each>
						<xsl:copy-of select="$varFrameHyperLinks"/>
					</a:p >
				</xsl:for-each >
			</xsl:when>
			<xsl:when test ="draw:text-box/text:p">
				<xsl:for-each select ="draw:text-box/text:p">
          <xsl:variable name="linkCount">
            <xsl:value-of select="position()"/>
          </xsl:variable>
          <xsl:variable name ="paraId" select ="@text:style-name"/>
          <xsl:variable name ="isNumberingEnabled">
            <xsl:choose >
              <xsl:when test ="document('content.xml')//style:style[@style:name=$paraId]/style:paragraph-properties/@text:enable-numbering">
                <xsl:value-of select ="document('content.xml')//style:style[@style:name=$paraId]/style:paragraph-properties/@text:enable-numbering"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select ="'false'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
					<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
						<xsl:call-template name ="paraProperties" >
							<xsl:with-param name ="paraId" >
								<xsl:value-of select ="@text:style-name"/>
							</xsl:with-param >
							<xsl:with-param name ="isBulleted" select ="'false'"/>
							<xsl:with-param name ="level" select ="'0'"/>
							<xsl:with-param name="framePresentaionStyleId" select="parent::node()/parent::node()/./@presentation:style-name" />
							<xsl:with-param name ="isNumberingEnabled" select ="$isNumberingEnabled"/>
						</xsl:call-template >
						<a:r >
							<a:rPr lang="en-US" smtClean="0">
								<!-- edited by vipul font size is not needed for title, outline and subtitle-->
								<xsl:if test="$prClsName != 'outline' and $prClsName != 'title' and $prClsName != 'subtitle'">
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
								</xsl:if>
								<xsl:if test="./text:a">
                  <xsl:call-template name="TextActionHyperLinks">
                    <xsl:with-param name="PostionCount" select="$PostionCount"/>
                    <xsl:with-param name="nodeType" select="'span'"/>
                    <xsl:with-param name="linkCount" select="$linkCount"/>
                  </xsl:call-template>
								</xsl:if>
							</a:rPr>
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
        <xsl:variable name="linkCount">
          <xsl:value-of select="position()"/>
        </xsl:variable>
				<a:p>
					<a:r >
						<a:rPr lang="en-US" smtClean="0">
							<a:latin charset="0" typeface="Arial" />
							<xsl:if test="./text:a">
                <xsl:call-template name="TextActionHyperLinks">
                  <xsl:with-param name="PostionCount" select="$PostionCount"/>
                  <xsl:with-param name="nodeType" select="'span'"/>
                  <xsl:with-param name="linkCount" select="$linkCount"/>
                </xsl:call-template>
							</xsl:if>
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
    <xsl:param name ="SMName"/>
    
    <xsl:for-each select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name=$SMName]/draw:frame[@presentation:class='footer']">
      <xsl:variable name="footerValue">
        <xsl:for-each select ="draw:text-box/text:p">
          <xsl:if test="./text:span/presentation:footer or ./presentation:footer">
            <xsl:value-of select="'false'"/>
          </xsl:if>
          <xsl:if test="not(text:span/presentation:footer) or not(./presentation:footer)">
            <xsl:value-of select="'true'"/>
          </xsl:if>
        </xsl:for-each>
      </xsl:variable>
      <xsl:if test="$footerValue='true' or document('content.xml')//presentation:footer-decl[@presentation:name=$footerId] ">
		<p:sp>
			<p:nvSpPr>
				<p:cNvPr id="16361" name="Footer Placeholder 5" />
				<p:cNvSpPr>
					<a:spLocks noGrp="1" />
				</p:cNvSpPr>
				<p:nvPr>
					<p:ph type="ftr" sz="quarter" idx="11" />
				</p:nvPr>
			</p:nvSpPr>
			<p:spPr >
				<!-- footer date layout details style:master-page style:name -->
        <xsl:if test="document('content.xml') //presentation:footer-decl[@presentation:name=$footerId]">
          <xsl:call-template name ="GetFrameDetails">
            <xsl:with-param name ="LayoutName" select ="'footer'"/>
          </xsl:call-template>
        </xsl:if>
			</p:spPr >
			<p:txBody>
				<a:bodyPr />
				<a:lstStyle />
        <a:p>
          <xsl:if test ="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name=$SMName]/draw:frame[@presentation:class='footer']">
            <xsl:for-each select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name=$SMName]/draw:frame[@presentation:class='footer']">
              
                      <xsl:choose>
                        <xsl:when test="document('content.xml') //presentation:footer-decl[@presentation:name=$footerId]">
                          <a:r>
                            <a:rPr lang="en-US" dirty="0" smtClean="0"/>
                            <a:t>
                          <xsl:for-each select ="document('content.xml') //presentation:footer-decl[@presentation:name=$footerId] ">
                            <xsl:value-of select ="."/>
                          </xsl:for-each >
                            </a:t>
                          </a:r>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:if test="not(./draw:text-box/text:p/text:span)">
                        <a:endParaRPr lang="en-US"/>
              </xsl:if>
              <xsl:if test="./draw:text-box/text:p/text:span">

                <xsl:for-each select="./draw:text-box/text:p/text:span">
                  <a:r>
                    <a:rPr lang="en-US" dirty="0" smtClean="0">
                      <!--Font Size-->
                      <xsl:variable name ="textId">
                        <xsl:value-of select ="@text:style-name"/>
                      </xsl:variable>
                      <xsl:if test ="not($textId ='')">
                        <xsl:call-template name ="tmpSMfontStyles">
                          <xsl:with-param name ="TextStyleID" select ="$textId" />
                        </xsl:call-template>
                      </xsl:if>
                    </a:rPr>
                    <a:t>
                          <xsl:call-template name ="insertTab" />
                    </a:t>
                  </a:r>
                </xsl:for-each>
              </xsl:if>
                          <a:endParaRPr lang="en-US" dirty="0"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:for-each>
          </xsl:if>
      		</a:p>
			</p:txBody>
		</p:sp >
    </xsl:if>
    </xsl:for-each>
	</xsl:template>
	<xsl:template name ="slideNumber">
		<xsl:param name ="pageNumber"/>
		<p:sp>
			<p:nvSpPr>
				<p:cNvPr id="16362" name="Slide Number Placeholder 4" />
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
				<p:cNvPr id="16363" name="Date Placeholder 3" />
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
				<xsl:variable name ="angle">
					<xsl:value-of select="substring-after(substring-before(substring-before(@draw:transform,'translate'),')'),'(')" />
				</xsl:variable>
				<xsl:variable name ="x2">
					<xsl:value-of select="substring-before(substring-before(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
				</xsl:variable>
				<xsl:variable name ="y2">
					<xsl:value-of select="substring-before(substring-after(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
				</xsl:variable>
				<xsl:if test="@draw:transform">
					<xsl:attribute name ="rot">
						<xsl:value-of select ="concat('draw-transform:ROT:',substring-before(@svg:width,'cm'), ':',
																   substring-before(@svg:height,'cm'), ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
					</xsl:attribute>
				</xsl:if>
				<a:off>
					<xsl:if test="@draw:transform">
						<xsl:attribute name ="x">
							<xsl:value-of select ="concat('draw-transform:X:',substring-before(@svg:width,'cm'), ':',
																   substring-before(@svg:height,'cm'), ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
						</xsl:attribute>

						<xsl:attribute name ="y">
							<xsl:value-of select ="concat('draw-transform:Y:',substring-before(@svg:width,'cm'), ':',
																   substring-before(@svg:height,'cm'), ':', 
																   $x2, ':', 
                                   $y2, ':', 
																   $angle)"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="not(@draw:transform)">
						<xsl:attribute name ="x">
							<xsl:call-template name ="convertToPoints">
								<xsl:with-param name ="unit" select ="'cm'"/>
								<xsl:with-param name ="length" select ="@svg:x"/>
							</xsl:call-template>
						</xsl:attribute>
						<xsl:attribute name ="y">
							<xsl:call-template name ="convertToPoints">
								<xsl:with-param name ="unit" select ="'cm'"/>
								<xsl:with-param name ="length" select ="@svg:y"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
				</a:off>
				<a:ext>
					<xsl:attribute name ="cx">
						<xsl:call-template name ="convertToPoints">
							<xsl:with-param name ="unit" select ="'cm'"/>
							<xsl:with-param name ="length" select ="@svg:width"/>
						</xsl:call-template>
					</xsl:attribute>
					<xsl:attribute name ="cy">
						<xsl:call-template name ="convertToPoints">
							<xsl:with-param name ="unit" select ="'cm'"/>
							<xsl:with-param name ="length" select ="@svg:height"/>
						</xsl:call-template>
					</xsl:attribute>
				</a:ext>
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

        <xsl:if test ="@fo:padding-top and @fo:padding-top!='cm'" >
          <xsl:attribute name ="tIns">
						<xsl:value-of select ="format-number(substring-before(@fo:padding-top,'cm')  *   360000 ,'#.##')"/>
          </xsl:attribute>
					</xsl:if>
					<!-- Added by Lohith - or @fo:padding-top = '0cm'-->
					<xsl:if test ="@fo:padding-top = '0cm' or @fo:padding-top='cm'">
            <xsl:attribute name ="tIns">
						<xsl:value-of select ="0"/>
            </xsl:attribute>
          </xsl:if >
          <xsl:if test ="@fo:padding-bottom and @fo:padding-bottom != 'cm'">
          <xsl:attribute name ="bIns">
						<xsl:value-of select ="format-number(substring-before(@fo:padding-bottom,'cm')  *   360000 ,'#.##')"/>
          </xsl:attribute>
        </xsl:if>
					<!-- Added by Lohith - or @fo:padding-bottom = '0cm'-->
					<xsl:if test ="@fo:padding-bottom = '0cm' or @fo:padding-bottom='cm'">
            <xsl:attribute name ="bIns">
						<xsl:value-of select ="0"/>
            </xsl:attribute>
          </xsl:if >
          <xsl:if test ="@fo:padding-left and @fo:padding-left != 'cm'">
          <xsl:attribute name ="lIns">
						<xsl:value-of select ="format-number(substring-before(@fo:padding-left,'cm')  *   360000 ,'#.##')"/>
        </xsl:attribute>
      </xsl:if>
      <!-- Added by Lohith - or @fo:padding-left = '0cm'-->
      <xsl:if test ="@fo:padding-left = '0cm' or @fo:padding-left='cm'">
        <xsl:attribute name ="lIns">
          <xsl:value-of select ="0"/>
            </xsl:attribute>
					</xsl:if >
					<xsl:if test ="@fo:padding-right and @fo:padding-right != 'cm'">
            <xsl:attribute name ="rIns">
						<xsl:value-of select ="format-number(substring-before(@fo:padding-right,'cm')  *   360000 ,'#.##')"/>
            </xsl:attribute>				
              </xsl:if>
					<!-- Added by Lohith - or @fo:padding-right = '0cm'-->
					<xsl:if test ="@fo:padding-right = '0cm' or @fo:padding-right='cm'">
            <xsl:attribute name ="rIns">
              <xsl:value-of select ="0"/>
            </xsl:attribute>
					</xsl:if >
				<xsl:variable name ="anchorVal">
					<xsl:choose >
						<xsl:when test ="@draw:textarea-vertical-align ='top'">
							<xsl:value-of select ="'t'"/>
						</xsl:when>
						<xsl:when test ="@draw:textarea-vertical-align ='middle'">
							<xsl:value-of select ="'ctr'"/>
						</xsl:when>
						<xsl:when test ="@draw:textarea-vertical-align ='bottom'">
							<xsl:value-of select ="'b'"/>
						</xsl:when>
						<!--<xsl:otherwise >
							<xsl:value-of select ="'t'"/>
						</xsl:otherwise>-->
					</xsl:choose>
				</xsl:variable>
				<xsl:if test ="$anchorVal != ''">
					<xsl:attribute name ="anchor">
						<xsl:value-of select ="$anchorVal"/>
					</xsl:attribute>
				</xsl:if>
				<xsl:attribute name ="anchorCtr">
					<xsl:choose >
						<xsl:when test ="@draw:textarea-horizontal-align ='center'">
							<xsl:value-of select ="1"/>
						</xsl:when>
						<xsl:when test ="@draw:textarea-horizontal-align='justify'">
							<xsl:value-of select ="0"/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select ="0"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
				<xsl:attribute name ="wrap">
					<xsl:choose >
						<!--<xsl:when test ="@fo:wrap-option ='no-wrap'">
							<xsl:value-of select ="'none'"/>
						</xsl:when>
						<xsl:when test ="@fo:wrap-option ='wrap'">
							<xsl:value-of select ="'square'"/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select ="'square'"/>
						</xsl:otherwise>-->
						<xsl:when test ="((@draw:auto-grow-height = 'false') and (@draw:auto-grow-width = 'false'))">
							<xsl:value-of select ="'none'"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select ="'square'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>

			</xsl:for-each>
		</a:bodyPr>
	</xsl:template>
  
  <xsl:template name="TextActionHyperLinks">
    <xsl:param name="PostionCount"/>
    <xsl:param name="nodeType"/>
    <xsl:param name="linkCount"/>
    <xsl:if test="text:a">
      <xsl:for-each select="text:a">
        <a:hlinkClick>
          <xsl:if test="@xlink:href[ contains(.,'#Slide')]">
            <xsl:attribute name="action">
              <xsl:value-of select="'ppaction://hlinksldjump'"/>
            </xsl:attribute>
            <xsl:attribute name="r:id">
              <xsl:value-of select="concat('TextHLAtchFileId',$linkCount,'Link',$linkCount)"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
            <xsl:attribute name="r:id">
              <xsl:value-of select="concat('TextHLAtchFileId',$linkCount,'Link',$linkCount)"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="not(@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and @xlink:href[ contains(.,':') or contains(.,'../')]">
            <xsl:attribute name="r:id">
              <xsl:value-of select="concat('TextHLAtchFileId',$linkCount,'Link',$linkCount)"/>
            </xsl:attribute>
          </xsl:if>
        </a:hlinkClick>
      </xsl:for-each>
    </xsl:if>
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
		<xsl:param name ="slideNo"/>
		<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">


			<xsl:variable name ="slMasterName">
				<xsl:value-of select ="@draw:master-page-name"/>
			</xsl:variable>
			<xsl:variable name ="slMasterLink">
				<xsl:for-each select ="document('styles.xml')//style:master-page/@style:name">
					<xsl:if test ="$slMasterName = .">
						<xsl:value-of select ="position()"/>
					</xsl:if>
				</xsl:for-each>
			</xsl:variable>
			<xsl:variable name ="layoutTemplate">
				<xsl:value-of select ="substring-after(@presentation:presentation-page-layout-name,'T')"/>
			</xsl:variable >

			<xsl:variable name ="layoutNo">
				<xsl:choose >
					<!-- Odp Layout 9,10,11,12,13 is mapped to layout 4-->
					<xsl:when test ="$layoutTemplate ='7' or $layoutTemplate ='8' or
						   $layoutTemplate ='9' or $layoutTemplate ='10' or $layoutTemplate ='13'">
						<xsl:value-of select ="4"/>
					</xsl:when>
					<!-- Odp Layout 2 is mapped to layout 1-->
					<xsl:when test ="$layoutTemplate ='0'" >
						<xsl:value-of select ="1"/>
					</xsl:when>
					<!-- Odp Layout 5 is mapped to layout 6-->
					<xsl:when test ="$layoutTemplate ='19'" >
						<xsl:value-of select ="6"/>
					</xsl:when>
					<!-- Odp Layout 3 is mapped to layout 2-->
					<xsl:when test ="$layoutTemplate ='1'" >
						<xsl:value-of select ="2"/>
					</xsl:when>
					<!-- Odp Layout 4 is mapped to layout 4-->
					<xsl:when test ="$layoutTemplate ='3'" >
						<xsl:value-of select ="4"/>
					</xsl:when>
					<!-- Odp Layout 4,5,6,7,8,14,16,17,18,19,20 is mapped to layout 7-->
					<xsl:otherwise >
						<xsl:value-of select ="7"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:variable>
			<!-- added by vipul to insert relation for notes-->
			<!--start-->
			<xsl:if test="presentation:notes/draw:page-thumbnail">
				<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide">
					<xsl:attribute name ="Target">
						<xsl:value-of select ="concat('../notesSlides/notesSlide',$slideNo,'.xml')"/>
					</xsl:attribute>
				</Relationship>
			</xsl:if>
			<!--start-->
			<Relationship Id="rId1"
									Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout">
				<xsl:attribute name ="Target">
					<xsl:value-of select ="concat('../slideLayouts/slideLayout',((($slMasterLink - 1)*11)  + $layoutNo) ,'.xml')"/>
				</xsl:attribute>
			</Relationship >
			<!-- added by Vipul to set Relationship for background Image  Start-->
			<xsl:variable name="dpName">
				<xsl:value-of select="@draw:style-name" />
			</xsl:variable>
			<xsl:for-each select="document('content.xml')//style:style[@style:name= $dpName]/style:drawing-page-properties">
				<xsl:if test="@draw:fill='bitmap'">
					<xsl:variable name="var_imageName" select="@draw:fill-image-name"/>
					<xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/draw:fill-image[@draw:name=$var_imageName]">
						<xsl:if test ="contains(@xlink:href,'.png') or contains(@xlink:href,'.emf') or contains(@xlink:href,'.wmf') or contains(@xlink:href,'.jfif') or contains(@xlink:href,'.jpe') 
            or contains(@xlink:href,'.bmp') or contains(@xlink:href,'.dib') or contains(@xlink:href,'.rle')
            or contains(@xlink:href,'.bmz') or contains(@xlink:href,'.gfa') 
            or contains(@xlink:href,'.emz') or contains(@xlink:href,'.wmz') or contains(@xlink:href,'.pcz')
            or contains(@xlink:href,'.tif') or contains(@xlink:href,'.tiff') 
            or contains(@xlink:href,'.cdr') or contains(@xlink:href,'.cgm') or contains(@xlink:href,'.eps') 
            or contains(@xlink:href,'.pct') or contains(@xlink:href,'.pict') or contains(@xlink:href,'.wpg') 
            or contains(@xlink:href,'.jpeg') or contains(@xlink:href,'.gif') or contains(@xlink:href,'.png') or contains(@xlink:href,'.jpg')">
							<Relationship>
								<xsl:attribute name="Id">
									<xsl:value-of select="concat('slide',$slideNo,'BackImg')" />
								</xsl:attribute>
								<xsl:attribute name="Type">
									<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'" />
								</xsl:attribute>
								<xsl:attribute name="Target">
									<xsl:value-of select="concat('../media/',substring-after(@xlink:href,'/'))" />
									<!-- <xsl:value-of select ="concat('../media/',substring-after(@xlink:href,'/'))"/>  -->
								</xsl:attribute>
							</Relationship>
						</xsl:if>
					</xsl:for-each>
				</xsl:if>
			</xsl:for-each>
			<xsl:for-each select ="node()">
				<xsl:choose>
					<xsl:when test="name()='draw:frame'">
						<xsl:variable name="var_pos" select="position()"/>
						<xsl:for-each select=".">
							<xsl:choose>
								<xsl:when test="./draw:image">
									<xsl:for-each select="./draw:image">
										<xsl:if test ="contains(@xlink:href,'.png') or contains(@xlink:href,'.emf') or contains(@xlink:href,'.wmf') or contains(@xlink:href,'.jfif') or contains(@xlink:href,'.jpe') 
            or contains(@xlink:href,'.bmp') or contains(@xlink:href,'.dib') or contains(@xlink:href,'.rle')
            or contains(@xlink:href,'.bmz') or contains(@xlink:href,'.gfa') 
            or contains(@xlink:href,'.emz') or contains(@xlink:href,'.wmz') or contains(@xlink:href,'.pcz')
            or contains(@xlink:href,'.tif') or contains(@xlink:href,'.tiff') 
            or contains(@xlink:href,'.cdr') or contains(@xlink:href,'.cgm') or contains(@xlink:href,'.eps') 
            or contains(@xlink:href,'.pct') or contains(@xlink:href,'.pict') or contains(@xlink:href,'.wpg') 
            or contains(@xlink:href,'.jpeg') or contains(@xlink:href,'.gif') or contains(@xlink:href,'.png') or contains(@xlink:href,'.jpg')">
                                                                                  <xsl:if test="not(./@xlink:href[contains(.,'../')])">
											<Relationship>
												<xsl:attribute name="Id">
													<xsl:value-of select="concat('sl',$slideNo,'Image',$var_pos)" />
												</xsl:attribute>
												<xsl:attribute name="Type">
													<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'" />
												</xsl:attribute>
												<xsl:attribute name="Target">
													<xsl:value-of select="concat('../media/',substring-after(@xlink:href,'/'))" />
													<!-- <xsl:value-of select ="concat('../media/',substring-after(@xlink:href,'/'))"/>  -->
												</xsl:attribute>
											</Relationship>
										</xsl:if>
									</xsl:if>
									</xsl:for-each>
								</xsl:when>
								<xsl:when test="./draw:plugin">
									<xsl:for-each select="./draw:plugin">
										<Relationship>
											<xsl:attribute name="Id">
												<xsl:value-of  select ="concat('sl',$slideNo,'Au',$var_pos)"/>
											</xsl:attribute>
											<xsl:variable name="wavId">
												<xsl:value-of  select ="concat('sl',$slideNo,'Au',$var_pos)"/>
											</xsl:variable>
											<xsl:attribute name="Type">
												<xsl:if test="@xlink:href[ contains(.,'mp3')] or @xlink:href[ contains(.,'m3u')] or @xlink:href[ contains(.,'wma')] or 
          @xlink:href[ contains(.,'wax')] or @xlink:href[ contains(.,'aif')] or @xlink:href[ contains(.,'aifc')] or
          @xlink:href[ contains(.,'aiff')] or @xlink:href[ contains(.,'au')] or @xlink:href[ contains(.,'snd')] or
          @xlink:href[ contains(.,'mid')] or @xlink:href[ contains(.,'midi')] or @xlink:href[ contains(.,'rmi')] or @xlink:href[ contains(.,'wav')] or @xlink:href[ contains(.,'WAV')]">
													<xsl:value-of select ="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio'"/>
												</xsl:if>
												<xsl:if test="not(@xlink:href[ contains(.,'mp3')] or @xlink:href[ contains(.,'m3u')] or @xlink:href[ contains(.,'wma')] or 
          @xlink:href[ contains(.,'wax')] or @xlink:href[ contains(.,'aif')] or @xlink:href[ contains(.,'aifc')] or
          @xlink:href[ contains(.,'aiff')] or @xlink:href[ contains(.,'au')] or @xlink:href[ contains(.,'snd')] or
          @xlink:href[ contains(.,'mid')] or @xlink:href[ contains(.,'midi')] or @xlink:href[ contains(.,'rmi')] or @xlink:href[ contains(.,'wav')] or @xlink:href[ contains(.,'WAV')])">
													<xsl:value-of select ="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video'"/>
												</xsl:if>
											</xsl:attribute>
											<xsl:attribute name="Target">
												<xsl:choose>
													<xsl:when test="@xlink:href[ contains(.,'wav')] or @xlink:href[ contains(.,'WAV')]">
														<xsl:value-of select="concat('../media/',$wavId,'.wav')"/>
													</xsl:when>
													<xsl:when test="@xlink:href[ contains(.,'./')]">
														<xsl:if test="string-length(substring-after(@xlink:href, '../')) = 0">
															<xsl:value-of select="/"/>
														</xsl:if>
														<xsl:if test="not(string-length(substring-after(@xlink:href, '../')) = 0)">
                              <!--links Absolute Path-->
                               <xsl:variable name ="xlinkPath" >
                                  <xsl:value-of select ="@xlink:href"/>
                               </xsl:variable>
                               <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/> 
														</xsl:if>
													</xsl:when>
													<xsl:when test="not(@xlink:href[ contains(.,'./')])">
														<xsl:value-of select="concat('file:///',translate(substring-after(@xlink:href,'/'),'/','\'))"/>
													</xsl:when>
												</xsl:choose>
											</xsl:attribute>
											<xsl:if test="not(@xlink:href[ contains(.,'wav')] or @xlink:href[ contains(.,'WAV')])">
												<xsl:attribute name="TargetMode">
													<xsl:value-of select="'External'"/>
												</xsl:attribute>
											</xsl:if>
										</Relationship>
										<!-- for image link -->
										<Relationship>
											<xsl:attribute name ="Id">
												<xsl:value-of  select ="concat('sl',$slideNo,'Im',$var_pos)"/>
											</xsl:attribute>
											<xsl:attribute name ="Type" >
												<xsl:value-of select ="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'"/>
											</xsl:attribute>
											<xsl:attribute name ="Target">
												<xsl:if test="./draw:image/@xlink:href">
													<xsl:value-of select ="concat('../media/',substring-after(./draw:image/@xlink:href,'/'))"/>
												</xsl:if>
												<xsl:if test="not(./draw:image/@xlink:href)">
													<xsl:value-of select ="concat('../media/','thumbnail.png')"/>
												</xsl:if>
											</xsl:attribute>
										</Relationship >
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:variable name="PostionCount">
										<xsl:value-of select="$var_pos"/>
									</xsl:variable>
									<xsl:if test="not(@presentation:style-name) and not(@presentation:class)">
										<xsl:for-each select="draw:text-box">
											<xsl:variable name="shapeId">
												<xsl:value-of select="concat('text-box',$PostionCount)"/>
											</xsl:variable>
											<xsl:for-each select="text:p">
												<xsl:if test="text:a/@xlink:href">
													<Relationship>
														<xsl:attribute name="Id">
															<xsl:value-of select="concat($shapeId,'Link',position())"/>
														</xsl:attribute>
														<xsl:choose>
															<xsl:when test="text:a/@xlink:href[contains(.,'#Slide')]">
																<xsl:attribute name="Type">
																	<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
																</xsl:attribute>
																<xsl:attribute name="Target">
																	<xsl:value-of select="concat('slide',substring-after(text:a/@xlink:href,'Slide '),'.xml')"/>
																</xsl:attribute>
															</xsl:when>
															<xsl:when test="text:a/@xlink:href[contains(.,'http://') or contains(.,'mailto:')]">
																<xsl:attribute name="Type">
																	<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
																</xsl:attribute>
																<xsl:attribute name="Target">
																	<xsl:value-of select="text:a/@xlink:href"/>
																</xsl:attribute>
																<xsl:attribute name="TargetMode">
																	<xsl:value-of select="'External'"/>
																</xsl:attribute>
															</xsl:when>
															<xsl:otherwise>
																<xsl:if test="text:a/@xlink:href[ contains (.,':') ]">
																	<xsl:attribute name="Type">
																		<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
																	</xsl:attribute>
																	<xsl:attribute name="Target">
																		<xsl:value-of select="concat('file:///',translate(substring-after(text:a/@xlink:href,'/'),'/','\'))"/>
																	</xsl:attribute>
																	<xsl:attribute name="TargetMode">
																		<xsl:value-of select="'External'"/>
																	</xsl:attribute>
																</xsl:if>
																<xsl:if test="not(text:a/@xlink:href[ contains (.,':') ])">
																	<xsl:attribute name="Type">
																		<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
																	</xsl:attribute>
																	<xsl:attribute name="Target">
                                    <!--links Absolute Path-->
                                    <xsl:variable name ="xlinkPath" >
                                      <xsl:value-of select ="text:a/@xlink:href"/>
                                    </xsl:variable>
                                    <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
																	</xsl:attribute>
																	<xsl:attribute name="TargetMode">
																		<xsl:value-of select="'External'"/>
																	</xsl:attribute>
																</xsl:if>
															</xsl:otherwise>
														</xsl:choose>
													</Relationship>
												</xsl:if>
												<xsl:if test="text:span/text:a/@xlink:href">
													<Relationship>
														<xsl:attribute name="Id">
															<xsl:value-of select="concat($shapeId,'Link',position())"/>
														</xsl:attribute>
														<xsl:choose>
															<xsl:when test="text:span/text:a/@xlink:href[contains(.,'#Slide')]">
																<xsl:attribute name="Type">
																	<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
																</xsl:attribute>
																<xsl:attribute name="Target">
																	<xsl:value-of select="concat('slide',substring-after(text:span/text:a/@xlink:href,'Slide '),'.xml')"/>
																</xsl:attribute>
															</xsl:when>
															<xsl:when test="text:span/text:a/@xlink:href[contains(.,'http://') or contains(.,'mailto:')]">
																<xsl:attribute name="Type">
																	<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
																</xsl:attribute>
																<xsl:attribute name="Target">
																	<xsl:value-of select="text:span/text:a/@xlink:href"/>
																</xsl:attribute>
																<xsl:attribute name="TargetMode">
																	<xsl:value-of select="'External'"/>
																</xsl:attribute>
															</xsl:when>
															<xsl:otherwise>
																<xsl:if test="text:span/text:a/@xlink:href[ contains (.,':') ]">
																	<xsl:attribute name="Type">
																		<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
																	</xsl:attribute>
																	<xsl:attribute name="Target">
																		<xsl:value-of select="concat('file:///',translate(substring-after(text:span/text:a/@xlink:href,'/'),'/','\'))"/>
																	</xsl:attribute>
																	<xsl:attribute name="TargetMode">
																		<xsl:value-of select="'External'"/>
																	</xsl:attribute>
																</xsl:if>
																<xsl:if test="not(text:span/text:a/@xlink:href[ contains (.,':') ])">
																	<xsl:attribute name="Type">
																		<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
																	</xsl:attribute>
																	<xsl:attribute name="Target">
                                    <!--links Absolute Path-->
                                    <xsl:variable name ="xlinkPath" >
                                      <xsl:value-of select ="text:span/text:a/@xlink:href"/>
                                    </xsl:variable>
                                    <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/> 
																	</xsl:attribute>
																	<xsl:attribute name="TargetMode">
																		<xsl:value-of select="'External'"/>
																	</xsl:attribute>
																</xsl:if>
															</xsl:otherwise>
														</xsl:choose>
													</Relationship>
												</xsl:if>
											</xsl:for-each>
											<xsl:for-each select="text:list">
												<xsl:variable name="forCount" select="position()" />
												<xsl:variable name="bulletlvl">
													<xsl:call-template name ="getListLevel">
														<xsl:with-param name ="levelCount"/>
													</xsl:call-template>
												</xsl:variable>
												<xsl:variable name="blvl">
													<xsl:if test="$bulletlvl > 0">
														<xsl:value-of select="$bulletlvl"/>
													</xsl:if>
													<xsl:if test="not($bulletlvl > 0)">
														<xsl:value-of select="'0'"/>
													</xsl:if>
												</xsl:variable>
                        <xsl:variable name ="listId" select ="./@text:style-name"/>
												<xsl:variable name="xhrefValue">
													<xsl:call-template name ="getTextHyperlinksForBulltes">
														<xsl:with-param name ="blvl" select="$blvl"/>
													</xsl:call-template>
												</xsl:variable>
												<xsl:if test="string-length($xhrefValue) > 0">
													<Relationship>
														<xsl:attribute name="Id">
															<xsl:value-of select="concat($shapeId,'BLVL',$blvl,'Link',$forCount)"/>
														</xsl:attribute>
														<xsl:choose>
															<xsl:when test="contains($xhrefValue,'#Slide')">
																<xsl:attribute name="Type">
																	<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
																</xsl:attribute>
																<xsl:attribute name="Target">
																	<xsl:value-of select="concat('slide',substring-after($xhrefValue,'Slide '),'.xml')"/>
																</xsl:attribute>
															</xsl:when>
															<xsl:when test="contains($xhrefValue,'http://') or contains($xhrefValue,'mailto:')">
																<xsl:attribute name="Type">
																	<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
																</xsl:attribute>
																<xsl:attribute name="Target">
																	<xsl:value-of select="$xhrefValue"/>
																</xsl:attribute>
																<xsl:attribute name="TargetMode">
																	<xsl:value-of select="'External'"/>
																</xsl:attribute>
															</xsl:when>
															<xsl:otherwise>
																<xsl:if test="contains($xhrefValue,':')">
																	<xsl:attribute name="Type">
																		<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
																	</xsl:attribute>
																	<xsl:attribute name="Target">
																		<xsl:value-of select="concat('file:///',translate(substring-after($xhrefValue,'/'),'/','\'))"/>
																	</xsl:attribute>
																	<xsl:attribute name="TargetMode">
																		<xsl:value-of select="'External'"/>
																	</xsl:attribute>
																</xsl:if>
																<xsl:if test="not(contains ($xhrefValue,':'))">
																	<xsl:attribute name="Type">
																		<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
																	</xsl:attribute>
																	<xsl:attribute name="Target">
                                    <!--links Absolute Path-->
                                    <xsl:variable name ="xlinkPath" >
                                      <xsl:value-of select ="$xhrefValue"/>
                                    </xsl:variable>
                                    <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>  
																	</xsl:attribute>
																	<xsl:attribute name="TargetMode">
																		<xsl:value-of select="'External'"/>
																	</xsl:attribute>
																</xsl:if>
															</xsl:otherwise>
														</xsl:choose>
													</Relationship>
												</xsl:if>
                        <xsl:for-each select ="document('content.xml')//text:list-style[@style:name=$listId]">
                          <xsl:if test ="text:list-level-style-image[@text:level=$blvl+1]">
                            <xsl:variable name ="rId" select ="concat('buImage',$listId,$blvl+1,$forCount)"/>
                            <xsl:variable name ="imageName" select ="substring-after(text:list-level-style-image[@text:level=$blvl+1]/@xlink:href,'Pictures/')"/>
                            <Relationship >
                              <xsl:attribute name ="Id">
                                <xsl:value-of  select ="$rId"/>
                              </xsl:attribute>
                              <xsl:attribute name ="Type" >
                                <xsl:value-of select ="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'"/>
                              </xsl:attribute>
                              <xsl:attribute name ="Target">
                                <xsl:value-of select ="concat('../media/',$imageName)"/>
                              </xsl:attribute>
                            </Relationship >
                          </xsl:if>
                        </xsl:for-each>
											</xsl:for-each>
										</xsl:for-each>
									</xsl:if>
									<xsl:if test="@presentation:class">
										<xsl:variable name="FrameCount">
											<xsl:value-of select="concat('Frame',$var_pos)"/>
										</xsl:variable>
										<xsl:for-each select="draw:text-box">
											<xsl:for-each select="text:p">
												<xsl:if test="text:a/@xlink:href">
													<Relationship>
														<xsl:attribute name="Id">
															<xsl:value-of select="concat('TextHLAtchFileId',position(),'Link',position())"/>
														</xsl:attribute>
														<xsl:choose>
															<xsl:when test="text:a/@xlink:href[contains(.,'#Slide')]">
																<xsl:attribute name="Type">
																	<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
																</xsl:attribute>
																<xsl:attribute name="Target">
																	<xsl:value-of select="concat('slide',substring-after(text:a/@xlink:href,'Slide '),'.xml')"/>
																</xsl:attribute>
															</xsl:when>
															<xsl:when test="text:a/@xlink:href[contains(.,'http://') or contains(.,'mailto:')]">
																<xsl:attribute name="Type">
																	<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
																</xsl:attribute>
																<xsl:attribute name="Target">
																	<xsl:value-of select="text:a/@xlink:href"/>
																</xsl:attribute>
																<xsl:attribute name="TargetMode">
																	<xsl:value-of select="'External'"/>
																</xsl:attribute>
															</xsl:when>
															<xsl:otherwise>
																<xsl:if test="text:a/@xlink:href[ contains (.,':') ]">
																	<xsl:attribute name="Type">
																		<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
																	</xsl:attribute>
																	<xsl:attribute name="Target">
																		<xsl:value-of select="concat('file:///',translate(substring-after(text:a/@xlink:href,'/'),'/','\'))"/>
																	</xsl:attribute>
																	<xsl:attribute name="TargetMode">
																		<xsl:value-of select="'External'"/>
																	</xsl:attribute>
																</xsl:if>
																<xsl:if test="not(text:a/@xlink:href[ contains (.,':') ])">
																	<xsl:attribute name="Type">
																		<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
																	</xsl:attribute>
																	<xsl:attribute name="Target">
                                    <!--links Absolute Path-->
                                    <xsl:variable name ="xlinkPath" >
                                      <xsl:value-of select ="text:a/@xlink:href"/>
                                    </xsl:variable>
                                    <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>  
																	</xsl:attribute>
																	<xsl:attribute name="TargetMode">
																		<xsl:value-of select="'External'"/>
																	</xsl:attribute>
																</xsl:if>
															</xsl:otherwise>
														</xsl:choose>
													</Relationship>
												</xsl:if>
												<xsl:if test="text:span/text:a/@xlink:href">
													<Relationship>
														<xsl:attribute name="Id">
															<xsl:value-of select="concat('TextHLAtchFileId',position(),'Link',position())"/>
														</xsl:attribute>
														<xsl:choose>
															<xsl:when test="text:span/text:a/@xlink:href[contains(.,'#Slide')]">
																<xsl:attribute name="Type">
																	<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
																</xsl:attribute>
																<xsl:attribute name="Target">
																	<xsl:value-of select="concat('slide',substring-after(text:span/text:a/@xlink:href,'Slide '),'.xml')"/>
																</xsl:attribute>
															</xsl:when>
															<xsl:when test="text:span/text:a/@xlink:href[contains(.,'http://') or contains(.,'mailto:')]">
																<xsl:attribute name="Type">
																	<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
																</xsl:attribute>
																<xsl:attribute name="Target">
																	<xsl:value-of select="text:span/text:a/@xlink:href"/>
																</xsl:attribute>
																<xsl:attribute name="TargetMode">
																	<xsl:value-of select="'External'"/>
																</xsl:attribute>
															</xsl:when>
															<xsl:otherwise>
																<xsl:if test="text:span/text:a/@xlink:href[ contains (.,':') ]">
																	<xsl:attribute name="Type">
																		<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
																	</xsl:attribute>
																	<xsl:attribute name="Target">
																		<xsl:value-of select="concat('file:///',translate(substring-after(text:span/text:a/@xlink:href,'/'),'/','\'))"/>
																	</xsl:attribute>
																	<xsl:attribute name="TargetMode">
																		<xsl:value-of select="'External'"/>
																	</xsl:attribute>
																</xsl:if>
																<xsl:if test="not(text:span/text:a/@xlink:href[ contains (.,':') ])">
																	<xsl:attribute name="Type">
																		<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
																	</xsl:attribute>
																	<xsl:attribute name="Target">
                                    <!--links Absolute Path-->
                                    <xsl:variable name ="xlinkPath" >
                                      <xsl:value-of select ="text:span/text:a/@xlink:href"/>
                                    </xsl:variable>
                                    <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
																	</xsl:attribute>
																	<xsl:attribute name="TargetMode">
																		<xsl:value-of select="'External'"/>
																	</xsl:attribute>
																</xsl:if>
															</xsl:otherwise>
														</xsl:choose>
													</Relationship>
												</xsl:if>
											</xsl:for-each>
											<xsl:for-each select="text:list">
												<xsl:variable name="forCount" select="position()" />
												<xsl:variable name="bulletlvl">
													<xsl:call-template name ="getListLevel">
														<xsl:with-param name ="levelCount"/>
													</xsl:call-template>
												</xsl:variable>
												<xsl:variable name="blvl">
													<xsl:if test="$bulletlvl > 0">
														<xsl:value-of select="$bulletlvl"/>
													</xsl:if>
													<xsl:if test="not($bulletlvl > 0)">
														<xsl:value-of select="'0'"/>
													</xsl:if>
												</xsl:variable>
												<xsl:variable name="xhrefValue">
													<xsl:call-template name ="getTextHyperlinksForBulltes">
														<xsl:with-param name ="blvl" select="$blvl"/>
													</xsl:call-template>
												</xsl:variable>
                        <xsl:variable name ="listId" select ="./@text:style-name"/>
												<xsl:if test="string-length($xhrefValue) > 0">
													<Relationship>
														<xsl:attribute name="Id">
															<xsl:value-of select="concat($FrameCount,'BTHLFileId',$blvl,'Link',$forCount)"/>
														</xsl:attribute>
														<xsl:choose>
															<xsl:when test="contains($xhrefValue,'#Slide')">
																<xsl:attribute name="Type">
																	<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
																</xsl:attribute>
																<xsl:attribute name="Target">
																	<xsl:value-of select="concat('slide',substring-after($xhrefValue,'Slide '),'.xml')"/>
																</xsl:attribute>
															</xsl:when>
															<xsl:when test="contains($xhrefValue,'http://') or contains($xhrefValue,'mailto:')">
																<xsl:attribute name="Type">
																	<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
																</xsl:attribute>
																<xsl:attribute name="Target">
																	<xsl:value-of select="$xhrefValue"/>
																</xsl:attribute>
																<xsl:attribute name="TargetMode">
																	<xsl:value-of select="'External'"/>
																</xsl:attribute>
															</xsl:when>
															<xsl:otherwise>
																<xsl:if test="contains($xhrefValue,':')">
																	<xsl:attribute name="Type">
																		<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
																	</xsl:attribute>
																	<xsl:attribute name="Target">
																		<xsl:value-of select="concat('file:///',translate(substring-after($xhrefValue,'/'),'/','\'))"/>
																	</xsl:attribute>
																	<xsl:attribute name="TargetMode">
																		<xsl:value-of select="'External'"/>
																	</xsl:attribute>
																</xsl:if>
																<xsl:if test="not(contains ($xhrefValue,':'))">
																	<xsl:attribute name="Type">
																		<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
																	</xsl:attribute>
																	<xsl:attribute name="Target">
                                    <!--links Absolute Path-->
                                    <xsl:variable name ="xlinkPath" >
                                      <xsl:value-of select ="$xhrefValue"/>
                                    </xsl:variable>
                                    <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
																	</xsl:attribute>
																	<xsl:attribute name="TargetMode">
																		<xsl:value-of select="'External'"/>
																	</xsl:attribute>
																</xsl:if>
															</xsl:otherwise>
														</xsl:choose>
													</Relationship>
												</xsl:if>
                        <xsl:for-each select ="document('content.xml')//text:list-style[@style:name=$listId]">
                          <xsl:if test ="text:list-level-style-image[@text:level=$blvl+1]">
                            <xsl:variable name ="rId" select ="concat('buImage',$listId,$blvl+1,$forCount)"/>
                            <xsl:variable name ="imageName" select ="substring-after(text:list-level-style-image[@text:level=$blvl+1]/@xlink:href,'Pictures/')"/>
                            <Relationship >
                              <xsl:attribute name ="Id">
                                <xsl:value-of  select ="$rId"/>
                              </xsl:attribute>
                              <xsl:attribute name ="Type" >
                                <xsl:value-of select ="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'"/>
                              </xsl:attribute>
                              <xsl:attribute name ="Target">
                                <xsl:value-of select ="concat('../media/',$imageName)"/>
                              </xsl:attribute>
                            </Relationship >
                          </xsl:if>
                        </xsl:for-each>
											</xsl:for-each>
										</xsl:for-each>
									</xsl:if>
									<xsl:if test="office:event-listeners">
										<xsl:for-each select=".">
											<xsl:variable name="PresentationClass">
												<xsl:if test="@presentation:class">
													<xsl:value-of select="1"/>
												</xsl:if>
												<xsl:if test="not(@presentation:class)">
													<xsl:value-of select="0"/>
												</xsl:if>
											</xsl:variable>
											<xsl:if test="$PresentationClass = 0">
												<xsl:if test="count(office:event-listeners/presentation:event-listener) = 1">
													<xsl:for-each select ="office:event-listeners/presentation:event-listener">
														<xsl:if test="@xlink:href">
															<Relationship>
																<xsl:choose>
																	<xsl:when test="@xlink:href[ contains(.,'#')] and string-length(substring-before(@xlink:href,'#')) = 0 ">
																		<xsl:variable name="pageID">
																			<xsl:call-template name="getThePageId">
																				<xsl:with-param name="PageName" select="substring-after(@xlink:href,'#')"/>
																			</xsl:call-template>
																		</xsl:variable>
																		<xsl:if test="$pageID > 0">
																			<xsl:attribute name="Id">
																				<xsl:value-of select="concat('ShapeFileId',$PostionCount)"/>
																				<!--<xsl:value-of select="concat('TxtBoxAtchFileId',$PostionCount)"/>-->
																			</xsl:attribute>
																			<xsl:attribute name="Type">
																				<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
																			</xsl:attribute>
																			<xsl:attribute name="Target">
																				<xsl:value-of select="concat('slide',$pageID,'.xml')"/>
																			</xsl:attribute>
																		</xsl:if>
																	</xsl:when>
																	<xsl:when test="@xlink:href">
																		<xsl:attribute name="Id">
																			<xsl:value-of select="concat('TxtBoxAtchFileId',$PostionCount)"/>
																			<!--<xsl:value-of select="concat('TxtBoxAtchFileId',$PostionCount)"/>-->
																		</xsl:attribute>
																		<xsl:attribute name="Type">
																			<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
																		</xsl:attribute>
																		<xsl:attribute name="Target">
																			<xsl:if test="@xlink:href[ contains(.,'./')]">
																				<xsl:if test="string-length(substring-after(@xlink:href, '../')) = 0">
																					<xsl:value-of select="/"/>
																				</xsl:if>
																				<xsl:if test="not(string-length(substring-after(@xlink:href, '../')) = 0)">
                                          <!--links Absolute Path-->
                                          <xsl:variable name ="xlinkPath" >
                                            <xsl:value-of select ="@xlink:href"/>
                                          </xsl:variable>
                                          <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
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
											</xsl:if>
										</xsl:for-each>
										<xsl:for-each select="./@presentation:class">
											<xsl:for-each select ="parent::node()/office:event-listeners/presentation:event-listener">
												<xsl:if test="@xlink:href">
													<Relationship>
														<xsl:choose>
															<xsl:when test="@xlink:href[ contains(.,'#')] and string-length(substring-before(@xlink:href,'#')) = 0 ">
																<xsl:variable name="pageID">
																	<xsl:call-template name="getThePageId">
																		<xsl:with-param name="PageName" select="substring-after(@xlink:href,'#')"/>
																	</xsl:call-template>
																</xsl:variable>
																<xsl:if test="$pageID > 0">
																	<xsl:attribute name="Id">
																		<xsl:value-of select="concat('AtchFileId',position())"/>
																		<!--<xsl:value-of select="concat('AtchFileId',$PostionCount)"/>-->
																	</xsl:attribute>
																	<xsl:attribute name="Type">
																		<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
																	</xsl:attribute>
																	<xsl:attribute name="Target">
																		<xsl:value-of select="concat('slide',$pageID,'.xml')"/>
																	</xsl:attribute>
																</xsl:if>
															</xsl:when>
															<xsl:when test="@xlink:href">
																<xsl:attribute name="Id">
																	<xsl:value-of select="concat('TxtBoxAtchFileId',$PostionCount)"/>
																	<!--<xsl:value-of select="concat('AtchFileId',$PostionCount)"/>-->
																</xsl:attribute>
																<xsl:attribute name="Type">
																	<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
																</xsl:attribute>
																<xsl:attribute name="Target">
																	<xsl:if test="@xlink:href[ contains(.,'./')]">
																		<xsl:if test="string-length(substring-after(@xlink:href, '../')) = 0">
																			<xsl:value-of select="/"/>
																		</xsl:if>
																		<xsl:if test="not(string-length(substring-after(@xlink:href, '../')) = 0)">
                                      <!--links Absolute Path-->
                                      <xsl:variable name ="xlinkPath" >
                                        <xsl:value-of select ="@xlink:href"/>
                                      </xsl:variable>
                                      <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
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
								</xsl:otherwise>
							</xsl:choose>
						</xsl:for-each>
					</xsl:when>
					<xsl:when test="name()='draw:custom-shape'">
						<xsl:variable name="var_pos" select="position()"/>
						<xsl:for-each select=".">
							<xsl:choose>
								<xsl:when test="./office:event-listeners">
									<xsl:variable name="ShapePostionCount">
										<xsl:value-of select="$var_pos"/>
									</xsl:variable>
									<xsl:for-each select ="office:event-listeners/presentation:event-listener">
										<xsl:if test="@xlink:href">
											<Relationship>
												<xsl:choose>
													<xsl:when test="@xlink:href[ contains(.,'#')] and string-length(substring-before(@xlink:href,'#')) = 0 ">
														<xsl:variable name="pageID">
															<xsl:call-template name="getThePageId">
																<xsl:with-param name="PageName" select="substring-after(@xlink:href,'#')"/>
															</xsl:call-template>
														</xsl:variable>
														<xsl:if test="$pageID > 0">
															<xsl:attribute name="Id">
																<xsl:value-of select="concat('ShapeFileId',$ShapePostionCount)"/>
															</xsl:attribute>
															<xsl:attribute name="Type">
																<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
															</xsl:attribute>
															<xsl:attribute name="Target">
																<xsl:value-of select="concat('slide',$pageID,'.xml')"/>
															</xsl:attribute>
														</xsl:if>
													</xsl:when>
													<xsl:when test="@xlink:href">
														<xsl:attribute name="Id">
															<xsl:value-of select="concat('ShapeFileId',$ShapePostionCount)"/>
														</xsl:attribute>
														<xsl:attribute name="Type">
															<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
														</xsl:attribute>
														<xsl:attribute name="Target">
															<xsl:if test="@xlink:href[ contains(.,'./')]">
																<xsl:if test="string-length(substring-after(@xlink:href, '../')) = 0">
																	<xsl:value-of select="/"/>
																</xsl:if>
																<xsl:if test="not(string-length(substring-after(@xlink:href, '../')) = 0)">
                                  <!--links Absolute Path-->
                                  <xsl:variable name ="xlinkPath" >
                                    <xsl:value-of select ="@xlink:href"/>
                                  </xsl:variable>
                                  <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
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
								</xsl:when>
								<xsl:otherwise>
									<xsl:variable name="ShapePostionCount">
										<xsl:value-of select="position()"/>
									</xsl:variable>
									<xsl:variable name="shapeId">
										<xsl:value-of select="concat('custom-shape',$var_pos)"/>
									</xsl:variable>
									<xsl:for-each select="text:p">
										<xsl:if test="text:a/@xlink:href">
											<Relationship>
												<xsl:attribute name="Id">
													<xsl:value-of select="concat($shapeId,'Link',position())"/>
												</xsl:attribute>
												<xsl:choose>
													<xsl:when test="text:a/@xlink:href[contains(.,'#Slide')]">
														<xsl:attribute name="Type">
															<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
														</xsl:attribute>
														<xsl:attribute name="Target">
															<xsl:value-of select="concat('slide',substring-after(text:a/@xlink:href,'Slide '),'.xml')"/>
														</xsl:attribute>
													</xsl:when>
													<xsl:when test="text:a/@xlink:href[contains(.,'http://') or contains(.,'mailto:')]">
														<xsl:attribute name="Type">
															<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
														</xsl:attribute>
														<xsl:attribute name="Target">
															<xsl:value-of select="text:a/@xlink:href"/>
														</xsl:attribute>
														<xsl:attribute name="TargetMode">
															<xsl:value-of select="'External'"/>
														</xsl:attribute>
													</xsl:when>
													<xsl:otherwise>
														<xsl:if test="text:a/@xlink:href[ contains (.,':') ]">
															<xsl:attribute name="Type">
																<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
															</xsl:attribute>
															<xsl:attribute name="Target">
																<xsl:value-of select="concat('file:///',translate(substring-after(text:a/@xlink:href,'/'),'/','\'))"/>
															</xsl:attribute>
															<xsl:attribute name="TargetMode">
																<xsl:value-of select="'External'"/>
															</xsl:attribute>
														</xsl:if>
														<xsl:if test="not(text:a/@xlink:href[ contains (.,':') ])">
															<xsl:attribute name="Type">
																<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
															</xsl:attribute>
															<xsl:attribute name="Target">
                                <!--links Absolute Path-->
                                <xsl:variable name ="xlinkPath" >
                                  <xsl:value-of select ="text:a/@xlink:href"/>
                                </xsl:variable>
                                <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
															</xsl:attribute>
															<xsl:attribute name="TargetMode">
																<xsl:value-of select="'External'"/>
															</xsl:attribute>
														</xsl:if>
													</xsl:otherwise>
												</xsl:choose>
											</Relationship>
										</xsl:if>
										<xsl:if test="text:span/text:a/@xlink:href">
											<Relationship>
												<xsl:attribute name="Id">
													<xsl:value-of select="concat($shapeId,'Link',position())"/>
												</xsl:attribute>
												<xsl:choose>
													<xsl:when test="text:span/text:a/@xlink:href[contains(.,'#Slide')]">
														<xsl:attribute name="Type">
															<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
														</xsl:attribute>
														<xsl:attribute name="Target">
															<xsl:value-of select="concat('slide',substring-after(text:span/text:a/@xlink:href,'Slide '),'.xml')"/>
														</xsl:attribute>
													</xsl:when>
													<xsl:when test="text:span/text:a/@xlink:href[contains(.,'http://') or contains(.,'mailto:')]">
														<xsl:attribute name="Type">
															<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
														</xsl:attribute>
														<xsl:attribute name="Target">
															<xsl:value-of select="text:span/text:a/@xlink:href"/>
														</xsl:attribute>
														<xsl:attribute name="TargetMode">
															<xsl:value-of select="'External'"/>
														</xsl:attribute>
													</xsl:when>
													<xsl:otherwise>
														<xsl:if test="text:span/text:a/@xlink:href[ contains (.,':') ]">
															<xsl:attribute name="Type">
																<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
															</xsl:attribute>
															<xsl:attribute name="Target">
																<xsl:value-of select="concat('file:///',translate(substring-after(text:span/text:a/@xlink:href,'/'),'/','\'))"/>
															</xsl:attribute>
															<xsl:attribute name="TargetMode">
																<xsl:value-of select="'External'"/>
															</xsl:attribute>
														</xsl:if>
														<xsl:if test="not(text:span/text:a/@xlink:href[ contains (.,':') ])">
															<xsl:attribute name="Type">
																<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
															</xsl:attribute>
															<xsl:attribute name="Target">
                                <!--links Absolute Path-->
                                <xsl:variable name ="xlinkPath" >
                                  <xsl:value-of select ="text:span/text:a/@xlink:href"/>
                                </xsl:variable>
                                <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
															</xsl:attribute>
															<xsl:attribute name="TargetMode">
																<xsl:value-of select="'External'"/>
															</xsl:attribute>
														</xsl:if>
													</xsl:otherwise>
												</xsl:choose>
											</Relationship>
										</xsl:if>
									</xsl:for-each>
									<xsl:for-each select="text:list">
										<xsl:variable name="forCount" select="position()" />
										<xsl:variable name="bulletlvl">
											<xsl:call-template name ="getListLevel">
												<xsl:with-param name ="levelCount"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:variable name="blvl">
											<xsl:if test="$bulletlvl > 0">
												<xsl:value-of select="$bulletlvl"/>
											</xsl:if>
											<xsl:if test="not($bulletlvl > 0)">
												<xsl:value-of select="'0'"/>
											</xsl:if>
										</xsl:variable>
										<xsl:variable name="xhrefValue">
											<xsl:call-template name ="getTextHyperlinksForBulltes">
												<xsl:with-param name ="blvl" select="$blvl"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:if test="string-length($xhrefValue) > 0">
											<Relationship>
												<xsl:attribute name="Id">
													<xsl:value-of select="concat($shapeId,'BLVL',$blvl,'Link',$forCount)"/>
												</xsl:attribute>
												<xsl:choose>
													<xsl:when test="contains($xhrefValue,'#Slide')">
														<xsl:attribute name="Type">
															<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
														</xsl:attribute>
														<xsl:attribute name="Target">
															<xsl:value-of select="concat('slide',substring-after($xhrefValue,'Slide '),'.xml')"/>
														</xsl:attribute>
													</xsl:when>
													<xsl:when test="contains($xhrefValue,'http://') or contains($xhrefValue,'mailto:')">
														<xsl:attribute name="Type">
															<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
														</xsl:attribute>
														<xsl:attribute name="Target">
															<xsl:value-of select="$xhrefValue"/>
														</xsl:attribute>
														<xsl:attribute name="TargetMode">
															<xsl:value-of select="'External'"/>
														</xsl:attribute>
													</xsl:when>
													<xsl:otherwise>
														<xsl:if test="contains($xhrefValue,':')">
															<xsl:attribute name="Type">
																<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
															</xsl:attribute>
															<xsl:attribute name="Target">
																<xsl:value-of select="concat('file:///',translate(substring-after($xhrefValue,'/'),'/','\'))"/>
															</xsl:attribute>
															<xsl:attribute name="TargetMode">
																<xsl:value-of select="'External'"/>
															</xsl:attribute>
														</xsl:if>
														<xsl:if test="not(contains ($xhrefValue,':'))">
															<xsl:attribute name="Type">
																<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
															</xsl:attribute>
															<xsl:attribute name="Target">
                                <!--links Absolute Path-->
                                <xsl:variable name ="xlinkPath" >
                                  <xsl:value-of select ="$xhrefValue"/>
                                </xsl:variable>
                                <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
															</xsl:attribute>
															<xsl:attribute name="TargetMode">
																<xsl:value-of select="'External'"/>
															</xsl:attribute>
														</xsl:if>
													</xsl:otherwise>
												</xsl:choose>
											</Relationship>
										</xsl:if>
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:for-each>
					</xsl:when>
					<xsl:when test="name()='draw:rect'">
						<xsl:variable name="var_pos" select="position()"/>
						<xsl:for-each select=".">
							<xsl:choose>
								<xsl:when test="./office:event-listeners">
									<xsl:variable name="ShapePostionCount">
										<xsl:value-of select="$var_pos"/>
									</xsl:variable>
									<xsl:for-each select ="office:event-listeners/presentation:event-listener">
										<xsl:if test="@xlink:href">
											<Relationship>
												<xsl:choose>
													<!--<xsl:when test="@xlink:href[contains(.,'#page')]">
                    <xsl:attribute name="Type">
                      <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
                    </xsl:attribute>
                    <xsl:attribute name="Target">
                      <xsl:value-of select="concat('slide',substring-after(@xlink:href,'page'),'.xml')"/>
                    </xsl:attribute>
                  </xsl:when>-->
													<xsl:when test="@xlink:href[ contains(.,'#')] and string-length(substring-before(@xlink:href,'#')) = 0 ">
														<xsl:variable name="pageID">
															<xsl:call-template name="getThePageId">
																<xsl:with-param name="PageName" select="substring-after(@xlink:href,'#')"/>
															</xsl:call-template>
														</xsl:variable>
														<xsl:if test="$pageID > 0">
															<xsl:attribute name="Id">
																<xsl:value-of select="concat('RectAtachFileId',$ShapePostionCount)"/>
															</xsl:attribute>
															<xsl:attribute name="Type">
																<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
															</xsl:attribute>
															<xsl:attribute name="Target">
																<xsl:value-of select="concat('slide',$pageID,'.xml')"/>
															</xsl:attribute>
														</xsl:if>
													</xsl:when>
													<xsl:when test="@xlink:href">
														<xsl:attribute name="Id">
															<xsl:value-of select="concat('RectAtachFileId',$ShapePostionCount)"/>
														</xsl:attribute>
														<xsl:attribute name="Type">
															<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
														</xsl:attribute>
														<xsl:attribute name="Target">
															<xsl:if test="@xlink:href[ contains(.,'./')]">
																<xsl:if test="string-length(substring-after(@xlink:href, '../')) = 0">
																	<xsl:value-of select="/"/>
																</xsl:if>
																<xsl:if test="not(string-length(substring-after(@xlink:href, '../')) = 0)">
                                  <!--links Absolute Path-->
                                  <xsl:variable name ="xlinkPath" >
                                    <xsl:value-of select ="@xlink:href"/>
                                  </xsl:variable>
                                  <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
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
								</xsl:when>
								<xsl:otherwise>
									<xsl:variable name="shapeId">
										<xsl:value-of select="concat('rect',$var_pos)"/>
									</xsl:variable>
									<xsl:for-each select="text:p">
										<xsl:if test="text:a/@xlink:href">
											<Relationship>
												<xsl:attribute name="Id">
													<xsl:value-of select="concat($shapeId,'Link',position())"/>
												</xsl:attribute>
												<xsl:choose>
													<xsl:when test="text:a/@xlink:href[contains(.,'#Slide')]">
														<xsl:attribute name="Type">
															<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
														</xsl:attribute>
														<xsl:attribute name="Target">
															<xsl:value-of select="concat('slide',substring-after(text:a/@xlink:href,'Slide '),'.xml')"/>
														</xsl:attribute>
													</xsl:when>
													<xsl:when test="text:a/@xlink:href[contains(.,'http://') or contains(.,'mailto:')]">
														<xsl:attribute name="Type">
															<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
														</xsl:attribute>
														<xsl:attribute name="Target">
															<xsl:value-of select="text:a/@xlink:href"/>
														</xsl:attribute>
														<xsl:attribute name="TargetMode">
															<xsl:value-of select="'External'"/>
														</xsl:attribute>
													</xsl:when>
													<xsl:otherwise>
														<xsl:if test="text:a/@xlink:href[ contains (.,':') ]">
															<xsl:attribute name="Type">
																<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
															</xsl:attribute>
															<xsl:attribute name="Target">
																<xsl:value-of select="concat('file:///',translate(substring-after(text:a/@xlink:href,'/'),'/','\'))"/>
															</xsl:attribute>
															<xsl:attribute name="TargetMode">
																<xsl:value-of select="'External'"/>
															</xsl:attribute>
														</xsl:if>
														<xsl:if test="not(text:a/@xlink:href[ contains (.,':') ])">
															<xsl:attribute name="Type">
																<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
															</xsl:attribute>
															<xsl:attribute name="Target">
                                <!--links Absolute Path-->
                                <xsl:variable name ="xlinkPath" >
                                  <xsl:value-of select ="text:a/@xlink:href"/>
                                </xsl:variable>
                                <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
															</xsl:attribute>
															<xsl:attribute name="TargetMode">
																<xsl:value-of select="'External'"/>
															</xsl:attribute>
														</xsl:if>
													</xsl:otherwise>
												</xsl:choose>
											</Relationship>
										</xsl:if>
										<xsl:if test="text:span/text:a/@xlink:href">
											<Relationship>
												<xsl:attribute name="Id">
													<xsl:value-of select="concat($shapeId,'Link',position())"/>
												</xsl:attribute>
												<xsl:choose>
													<xsl:when test="text:span/text:a/@xlink:href[contains(.,'#Slide')]">
														<xsl:attribute name="Type">
															<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
														</xsl:attribute>
														<xsl:attribute name="Target">
															<xsl:value-of select="concat('slide',substring-after(text:span/text:a/@xlink:href,'Slide '),'.xml')"/>
														</xsl:attribute>
													</xsl:when>
													<xsl:when test="text:span/text:a/@xlink:href[contains(.,'http://') or contains(.,'mailto:')]">
														<xsl:attribute name="Type">
															<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
														</xsl:attribute>
														<xsl:attribute name="Target">
															<xsl:value-of select="text:span/text:a/@xlink:href"/>
														</xsl:attribute>
														<xsl:attribute name="TargetMode">
															<xsl:value-of select="'External'"/>
														</xsl:attribute>
													</xsl:when>
													<xsl:otherwise>
														<xsl:if test="text:span/text:a/@xlink:href[ contains (.,':') ]">
															<xsl:attribute name="Type">
																<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
															</xsl:attribute>
															<xsl:attribute name="Target">
																<xsl:value-of select="concat('file:///',translate(substring-after(text:span/text:a/@xlink:href,'/'),'/','\'))"/>
															</xsl:attribute>
															<xsl:attribute name="TargetMode">
																<xsl:value-of select="'External'"/>
															</xsl:attribute>
														</xsl:if>
														<xsl:if test="not(text:span/text:a/@xlink:href[ contains (.,':') ])">
															<xsl:attribute name="Type">
																<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
															</xsl:attribute>
															<xsl:attribute name="Target">
                                <!--links Absolute Path-->
                                <xsl:variable name ="xlinkPath" >
                                  <xsl:value-of select ="text:span/text:a/@xlink:href"/>
                                </xsl:variable>
                                <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
															</xsl:attribute>
															<xsl:attribute name="TargetMode">
																<xsl:value-of select="'External'"/>
															</xsl:attribute>
														</xsl:if>
													</xsl:otherwise>
												</xsl:choose>
											</Relationship>
										</xsl:if>
									</xsl:for-each>
									<xsl:for-each select="text:list">
										<xsl:variable name="forCount" select="position()" />
										<xsl:variable name="bulletlvl">
											<xsl:call-template name ="getListLevel">
												<xsl:with-param name ="levelCount"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:variable name="blvl">
											<xsl:if test="$bulletlvl > 0">
												<xsl:value-of select="$bulletlvl"/>
											</xsl:if>
											<xsl:if test="not($bulletlvl > 0)">
												<xsl:value-of select="'0'"/>
											</xsl:if>
										</xsl:variable>
										<xsl:variable name="xhrefValue">
											<xsl:call-template name ="getTextHyperlinksForBulltes">
												<xsl:with-param name ="blvl" select="$blvl"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:if test="string-length($xhrefValue) > 0">
											<Relationship>
												<xsl:attribute name="Id">
													<xsl:value-of select="concat($shapeId,'BLVL',$blvl,'Link',$forCount)"/>
												</xsl:attribute>
												<xsl:choose>
													<xsl:when test="contains($xhrefValue,'#Slide')">
														<xsl:attribute name="Type">
															<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
														</xsl:attribute>
														<xsl:attribute name="Target">
															<xsl:value-of select="concat('slide',substring-after($xhrefValue,'Slide '),'.xml')"/>
														</xsl:attribute>
													</xsl:when>
													<xsl:when test="contains($xhrefValue,'http://') or contains($xhrefValue,'mailto:')">
														<xsl:attribute name="Type">
															<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
														</xsl:attribute>
														<xsl:attribute name="Target">
															<xsl:value-of select="$xhrefValue"/>
														</xsl:attribute>
														<xsl:attribute name="TargetMode">
															<xsl:value-of select="'External'"/>
														</xsl:attribute>
													</xsl:when>
													<xsl:otherwise>
														<xsl:if test="contains($xhrefValue,':')">
															<xsl:attribute name="Type">
																<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
															</xsl:attribute>
															<xsl:attribute name="Target">
																<xsl:value-of select="concat('file:///',translate(substring-after($xhrefValue,'/'),'/','\'))"/>
															</xsl:attribute>
															<xsl:attribute name="TargetMode">
																<xsl:value-of select="'External'"/>
															</xsl:attribute>
														</xsl:if>
														<xsl:if test="not(contains ($xhrefValue,':'))">
															<xsl:attribute name="Type">
																<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
															</xsl:attribute>
															<xsl:attribute name="Target">
                                <!--links Absolute Path-->
                                <xsl:variable name ="xlinkPath" >
                                  <xsl:value-of select ="$xhrefValue"/>
                                </xsl:variable>
                                <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
															</xsl:attribute>
															<xsl:attribute name="TargetMode">
																<xsl:value-of select="'External'"/>
															</xsl:attribute>
														</xsl:if>
													</xsl:otherwise>
												</xsl:choose>
											</Relationship>
										</xsl:if>
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:for-each>
					</xsl:when>
					<xsl:when test="name()='draw:ellipse'">
						<xsl:variable name="var_pos" select="position()"/>
						<xsl:for-each select=".">
							<xsl:variable name="shapeId">
								<xsl:value-of select="concat('ellipse',$var_pos)"/>
							</xsl:variable>
							<xsl:for-each select="text:p">
								<xsl:if test="text:a/@xlink:href">
									<Relationship>
										<xsl:attribute name="Id">
											<xsl:value-of select="concat($shapeId,'Link',position())"/>
										</xsl:attribute>
										<xsl:choose>
											<xsl:when test="text:a/@xlink:href[contains(.,'#Slide')]">
												<xsl:attribute name="Type">
													<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
												</xsl:attribute>
												<xsl:attribute name="Target">
													<xsl:value-of select="concat('slide',substring-after(text:a/@xlink:href,'Slide '),'.xml')"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:when test="text:a/@xlink:href[contains(.,'http://') or contains(.,'mailto:')]">
												<xsl:attribute name="Type">
													<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
												</xsl:attribute>
												<xsl:attribute name="Target">
													<xsl:value-of select="text:a/@xlink:href"/>
												</xsl:attribute>
												<xsl:attribute name="TargetMode">
													<xsl:value-of select="'External'"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:otherwise>
												<xsl:if test="text:a/@xlink:href[ contains (.,':') ]">
													<xsl:attribute name="Type">
														<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
													</xsl:attribute>
													<xsl:attribute name="Target">
														<xsl:value-of select="concat('file:///',translate(substring-after(text:a/@xlink:href,'/'),'/','\'))"/>
													</xsl:attribute>
													<xsl:attribute name="TargetMode">
														<xsl:value-of select="'External'"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test="not(text:a/@xlink:href[ contains (.,':') ])">
													<xsl:attribute name="Type">
														<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
													</xsl:attribute>
													<xsl:attribute name="Target">
                            <!--links Absolute Path-->
                            <xsl:variable name ="xlinkPath" >
                              <xsl:value-of select ="text:a/@xlink:href"/>
                            </xsl:variable>
                            <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
													</xsl:attribute>
													<xsl:attribute name="TargetMode">
														<xsl:value-of select="'External'"/>
													</xsl:attribute>
												</xsl:if>
											</xsl:otherwise>
										</xsl:choose>
									</Relationship>
								</xsl:if>
								<xsl:if test="text:span/text:a/@xlink:href">
									<Relationship>
										<xsl:attribute name="Id">
											<xsl:value-of select="concat($shapeId,'Link',position())"/>
										</xsl:attribute>
										<xsl:choose>
											<xsl:when test="text:span/text:a/@xlink:href[contains(.,'#Slide')]">
												<xsl:attribute name="Type">
													<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
												</xsl:attribute>
												<xsl:attribute name="Target">
													<xsl:value-of select="concat('slide',substring-after(text:span/text:a/@xlink:href,'Slide '),'.xml')"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:when test="text:span/text:a/@xlink:href[contains(.,'http://') or contains(.,'mailto:')]">
												<xsl:attribute name="Type">
													<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
												</xsl:attribute>
												<xsl:attribute name="Target">
													<xsl:value-of select="text:span/text:a/@xlink:href"/>
												</xsl:attribute>
												<xsl:attribute name="TargetMode">
													<xsl:value-of select="'External'"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:otherwise>
												<xsl:if test="text:span/text:a/@xlink:href[ contains (.,':') ]">
													<xsl:attribute name="Type">
														<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
													</xsl:attribute>
													<xsl:attribute name="Target">
														<xsl:value-of select="concat('file:///',translate(substring-after(text:span/text:a/@xlink:href,'/'),'/','\'))"/>
													</xsl:attribute>
													<xsl:attribute name="TargetMode">
														<xsl:value-of select="'External'"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test="not(text:span/text:a/@xlink:href[ contains (.,':') ])">
													<xsl:attribute name="Type">
														<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
													</xsl:attribute>
													<xsl:attribute name="Target">
                            <!--links Absolute Path-->
                            <xsl:variable name ="xlinkPath" >
                              <xsl:value-of select ="text:span/text:a/@xlink:href"/>
                            </xsl:variable>
                            <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
													</xsl:attribute>
													<xsl:attribute name="TargetMode">
														<xsl:value-of select="'External'"/>
													</xsl:attribute>
												</xsl:if>
											</xsl:otherwise>
										</xsl:choose>
									</Relationship>
								</xsl:if>
							</xsl:for-each>
							<xsl:for-each select="text:list">
								<xsl:variable name="forCount" select="position()" />
								<xsl:variable name="bulletlvl">
									<xsl:call-template name ="getListLevel">
										<xsl:with-param name ="levelCount"/>
									</xsl:call-template>
								</xsl:variable>
								<xsl:variable name="blvl">
									<xsl:if test="$bulletlvl > 0">
										<xsl:value-of select="$bulletlvl"/>
									</xsl:if>
									<xsl:if test="not($bulletlvl > 0)">
										<xsl:value-of select="'0'"/>
									</xsl:if>
								</xsl:variable>
								<xsl:variable name="xhrefValue">
									<xsl:call-template name ="getTextHyperlinksForBulltes">
										<xsl:with-param name ="blvl" select="$blvl"/>
									</xsl:call-template>
								</xsl:variable>
								<xsl:if test="string-length($xhrefValue) > 0">
									<Relationship>
										<xsl:attribute name="Id">
											<xsl:value-of select="concat($shapeId,'BLVL',$blvl,'Link',$forCount)"/>
										</xsl:attribute>
										<xsl:choose>
											<xsl:when test="contains($xhrefValue,'#Slide')">
												<xsl:attribute name="Type">
													<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
												</xsl:attribute>
												<xsl:attribute name="Target">
													<xsl:value-of select="concat('slide',substring-after($xhrefValue,'Slide '),'.xml')"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:when test="contains($xhrefValue,'http://') or contains($xhrefValue,'mailto:')">
												<xsl:attribute name="Type">
													<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
												</xsl:attribute>
												<xsl:attribute name="Target">
													<xsl:value-of select="$xhrefValue"/>
												</xsl:attribute>
												<xsl:attribute name="TargetMode">
													<xsl:value-of select="'External'"/>
												</xsl:attribute>
											</xsl:when>
											<xsl:otherwise>
												<xsl:if test="contains($xhrefValue,':')">
													<xsl:attribute name="Type">
														<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
													</xsl:attribute>
													<xsl:attribute name="Target">
														<xsl:value-of select="concat('file:///',translate(substring-after($xhrefValue,'/'),'/','\'))"/>
													</xsl:attribute>
													<xsl:attribute name="TargetMode">
														<xsl:value-of select="'External'"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test="not(contains ($xhrefValue,':'))">
													<xsl:attribute name="Type">
														<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
													</xsl:attribute>
													<xsl:attribute name="Target">
                            <!--links Absolute Path-->
                            <xsl:variable name ="xlinkPath" >
                              <xsl:value-of select ="$xhrefValue"/>
                            </xsl:variable>
                            <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
													</xsl:attribute>
													<xsl:attribute name="TargetMode">
														<xsl:value-of select="'External'"/>
													</xsl:attribute>
												</xsl:if>
											</xsl:otherwise>
										</xsl:choose>
									</Relationship>
								</xsl:if>
							</xsl:for-each>
						</xsl:for-each>
					</xsl:when>
					<xsl:when test="name()='draw:line'">
						<xsl:variable name="var_pos" select="position()"/>
						<xsl:for-each select=".">
							<xsl:if test="./office:event-listeners">
								<xsl:variable name="ShapePostionCount">
									<xsl:value-of select="$var_pos"/>
								</xsl:variable>
								<xsl:for-each select ="office:event-listeners/presentation:event-listener">
									<xsl:if test="@xlink:href">
										<Relationship>
											<xsl:choose>
												<xsl:when test="@xlink:href[ contains(.,'#')] and string-length(substring-before(@xlink:href,'#')) = 0 ">
													<xsl:variable name="pageID">
														<xsl:call-template name="getThePageId">
															<xsl:with-param name="PageName" select="substring-after(@xlink:href,'#')"/>
														</xsl:call-template>
													</xsl:variable>
													<xsl:if test="$pageID > 0">
														<xsl:attribute name="Id">
															<xsl:value-of select="concat('LineFileId',$ShapePostionCount)"/>
														</xsl:attribute>
														<xsl:attribute name="Type">
															<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
														</xsl:attribute>
														<xsl:attribute name="Target">
															<xsl:value-of select="concat('slide',$pageID,'.xml')"/>
														</xsl:attribute>
													</xsl:if>
												</xsl:when>
												<xsl:when test="@xlink:href">
													<xsl:attribute name="Id">
														<xsl:value-of select="concat('LineFileId',$ShapePostionCount)"/>
													</xsl:attribute>
													<xsl:attribute name="Type">
														<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
													</xsl:attribute>
													<xsl:attribute name="Target">
														<xsl:if test="@xlink:href[ contains(.,'./')]">
															<xsl:if test="string-length(substring-after(@xlink:href, '../')) = 0">
																<xsl:value-of select="/"/>
															</xsl:if>
															<xsl:if test="not(string-length(substring-after(@xlink:href, '../')) = 0)">
                                <!--links Absolute Path-->
                                <xsl:variable name ="xlinkPath" >
                                  <xsl:value-of select ="@xlink:href"/>
                                </xsl:variable>
                                <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
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
							</xsl:if>
						</xsl:for-each>
					</xsl:when>
					<xsl:when test="name()='draw:connector'">
						<xsl:variable name="var_pos" select="position()"/>
						<xsl:for-each select=".">
							<xsl:if test="./office:event-listener">
								<xsl:variable name="ShapePostionCount">
									<xsl:value-of select="$var_pos"/>
								</xsl:variable>
								<xsl:for-each select ="office:event-listeners/presentation:event-listener">
									<xsl:if test="@xlink:href">
										<Relationship>
											<xsl:choose>
												<xsl:when test="@xlink:href[ contains(.,'#')] and string-length(substring-before(@xlink:href,'#')) = 0 ">
													<xsl:variable name="pageID">
														<xsl:call-template name="getThePageId">
															<xsl:with-param name="PageName" select="substring-after(@xlink:href,'#')"/>
														</xsl:call-template>
													</xsl:variable>
													<xsl:if test="$pageID > 0">
														<xsl:attribute name="Id">
															<xsl:value-of select="concat('LineFileId',$ShapePostionCount)"/>
														</xsl:attribute>
														<xsl:attribute name="Type">
															<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'"/>
														</xsl:attribute>
														<xsl:attribute name="Target">
															<xsl:value-of select="concat('slide',$pageID,'.xml')"/>
														</xsl:attribute>
													</xsl:if>
												</xsl:when>
												<xsl:when test="@xlink:href">
													<xsl:attribute name="Id">
														<xsl:value-of select="concat('LineFileId',$ShapePostionCount)"/>
													</xsl:attribute>
													<xsl:attribute name="Type">
														<xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'"/>
													</xsl:attribute>
													<xsl:attribute name="Target">
														<xsl:if test="@xlink:href[ contains(.,'./')]">
															<xsl:if test="string-length(substring-after(@xlink:href, '../')) = 0">
																<xsl:value-of select="/"/>
															</xsl:if>
															<xsl:if test="not(string-length(substring-after(@xlink:href, '../')) = 0)">
                                <!--links Absolute Path-->
                                <xsl:variable name ="xlinkPath" >
                                  <xsl:value-of select ="@xlink:href"/>
                                </xsl:variable>
                                <xsl:value-of select ="concat('hyperlink-path:',$xlinkPath)"/>
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
							</xsl:if>
						</xsl:for-each>
					</xsl:when>
					<xsl:when test="name()='draw:g'">
            <xsl:variable name="var_pos" select="position()"/>
            <xsl:for-each select ="node()">
              <xsl:variable name="var_num_1">
                <xsl:value-of select="position()"/>
              </xsl:variable>
              <xsl:variable name="var_num_2">
                <xsl:number level="any"/>
              </xsl:variable>
              <xsl:variable name="NvPrId" select="number(concat($var_pos,$var_num_1,$var_num_2))"/>
              <xsl:choose>
                <xsl:when test="name()='draw:frame'">
                  <!--<xsl:variable name="var_pos" select="position()"/>-->
                  <xsl:for-each select=".">
                    <xsl:choose>
                      <xsl:when test="./draw:image">
                        <xsl:for-each select="./draw:image">
                          <xsl:if test ="contains(@xlink:href,'.png') or contains(@xlink:href,'.emf') or contains(@xlink:href,'.wmf') or contains(@xlink:href,'.jfif') or contains(@xlink:href,'.jpe') 
            or contains(@xlink:href,'.bmp') or contains(@xlink:href,'.dib') or contains(@xlink:href,'.rle')
            or contains(@xlink:href,'.bmz') or contains(@xlink:href,'.gfa') 
            or contains(@xlink:href,'.emz') or contains(@xlink:href,'.wmz') or contains(@xlink:href,'.pcz')
            or contains(@xlink:href,'.tif') or contains(@xlink:href,'.tiff') 
            or contains(@xlink:href,'.cdr') or contains(@xlink:href,'.cgm') or contains(@xlink:href,'.eps') 
            or contains(@xlink:href,'.pct') or contains(@xlink:href,'.pict') or contains(@xlink:href,'.wpg') 
            or contains(@xlink:href,'.jpeg') or contains(@xlink:href,'.gif') or contains(@xlink:href,'.png') or contains(@xlink:href,'.jpg')">
                            <xsl:if test="not(./@xlink:href[contains(.,'../')])">
                            <Relationship>
                              <xsl:attribute name="Id">
                                <xsl:value-of select="concat('sl',$slideNo,'Image',$NvPrId)" />
                              </xsl:attribute>
                              <xsl:attribute name="Type">
                                <xsl:value-of select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'" />
                              </xsl:attribute>
                              <xsl:attribute name="Target">
                                <xsl:value-of select="concat('../media/',substring-after(@xlink:href,'/'))" />
                                <!-- <xsl:value-of select ="concat('../media/',substring-after(@xlink:href,'/'))"/>  -->
                              </xsl:attribute>
                            </Relationship>
                          </xsl:if>
                          </xsl:if>
                        </xsl:for-each>
                      </xsl:when>
                      <xsl:when test="./draw:plugin">
                      </xsl:when>
                    </xsl:choose>
                  </xsl:for-each>
                </xsl:when>
                <xsl:when test="name()='draw:custom-shape'">
                </xsl:when>
                <xsl:when test="name()='draw:rect'">
                </xsl:when>
                <xsl:when test="name()='draw:ellipse'">
                </xsl:when>
                <xsl:when test="name()='draw:line'">
                </xsl:when>
                <xsl:when test="name()='draw:connector'">
                </xsl:when>
                <xsl:when test="name()='draw:g'">
                </xsl:when>
              </xsl:choose>
            </xsl:for-each>
					</xsl:when>
				</xsl:choose>
			</xsl:for-each>
     
		</Relationships>
	</xsl:template>
	<xsl:template name="GetUniqueRelationIdForWavFile">
		<xsl:param name ="FilePath"/>
		<xsl:param name ="ShapePosition"/>
		<xsl:param name="Page" />
		<xsl:param name="ShapeType" />
		<xsl:param name="isCalledFromGroup" />
		<xsl:for-each select="$Page">
			<xsl:if test="$isCalledFromGroup != 'GROUP'">
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
			</xsl:if>
			<xsl:if test="$isCalledFromGroup = 'GROUP'">
				<xsl:if test="$ShapeType = 'CUSTOM' ">
					<xsl:for-each select="./draw:g/draw:custom-shape">
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
					<xsl:for-each select="./draw:g/draw:custom-shape">
						<xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
							<xsl:value-of select='1'/>
						</xsl:if>
					</xsl:for-each>
				</xsl:if>
				<xsl:if test="$ShapeType = 'LINE'">
					<xsl:for-each select="./draw:g/draw:custom-shape">
						<xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)]">
							<xsl:value-of select='1'/>
						</xsl:if>
					</xsl:for-each>
					<xsl:for-each select="./draw:g/draw:rect">
						<xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
							<xsl:value-of select='1'/>
						</xsl:if>
					</xsl:for-each>
				</xsl:if>
				<xsl:if test="$ShapeType = 'FRAME'">
					<xsl:for-each select="./draw:g/draw:custom-shape">
						<xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
							<xsl:value-of select='1'/>
						</xsl:if>
					</xsl:for-each>
					<xsl:for-each select="./draw:g/draw:rect">
						<xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)]">
							<xsl:value-of select='1'/>
						</xsl:if>
					</xsl:for-each>
					<xsl:for-each select="./draw:g/draw:line">
						<xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
							<xsl:value-of select='1'/>
						</xsl:if>
					</xsl:for-each>
				</xsl:if>
				<xsl:if test="$ShapeType = 'CONNECTOR'">
					<xsl:for-each select="./draw:g/draw:custom-shape">
						<xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
							<xsl:value-of select='1'/>
						</xsl:if>
					</xsl:for-each>
					<xsl:for-each select="./draw:g/draw:rect">
						<xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
							<xsl:value-of select='1'/>
						</xsl:if>
					</xsl:for-each>
					<xsl:for-each select="./draw:g/draw:line">
						<xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
							<xsl:value-of select='1'/>
						</xsl:if>
					</xsl:for-each>
					<xsl:for-each select="./draw:g/draw:frame">
						<xsl:if test="office:event-listeners/presentation:event-listener/presentation:sound/@xlink:href [ contains(.,$FilePath)] ">
							<xsl:value-of select='1'/>
						</xsl:if>
					</xsl:for-each>
				</xsl:if>
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
	<xsl:template name="getHyperlinksForBulltedText">
		<xsl:param name ="bulletLevel"/>
		<xsl:param name ="listItemCount" />
		<xsl:param name ="FrameCount" />
		<xsl:param name ="rootNode" />
		<xsl:choose>
			<xsl:when test ="./text:p">
				<xsl:for-each select="./text:p">
					<xsl:if test="text:a">
						<a:hlinkClick>
							<xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
								<xsl:attribute name="action">
									<xsl:value-of select="'ppaction://hlinksldjump'"/>
								</xsl:attribute>
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
						</a:hlinkClick>
					</xsl:if>
					<xsl:if test="text:span/text:a">
						<a:hlinkClick>
							<xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
								<xsl:attribute name="action">
									<xsl:value-of select="'ppaction://hlinksldjump'"/>
								</xsl:attribute>
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
						</a:hlinkClick>
					</xsl:if>
				</xsl:for-each>
			</xsl:when>
			<xsl:when test ="./text:list/text:list-item/text:p">
				<xsl:for-each select="./text:list/text:list-item/text:p">
					<xsl:if test="text:a">
						<a:hlinkClick>
							<xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
								<xsl:attribute name="action">
									<xsl:value-of select="'ppaction://hlinksldjump'"/>
								</xsl:attribute>
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
						</a:hlinkClick>
					</xsl:if>
					<xsl:if test="text:span/text:a">
						<a:hlinkClick>
							<xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
								<xsl:attribute name="action">
									<xsl:value-of select="'ppaction://hlinksldjump'"/>
								</xsl:attribute>
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
						</a:hlinkClick>
					</xsl:if>
				</xsl:for-each>
			</xsl:when>
			<xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:p">
				<xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:p">
					<xsl:if test="text:a">
						<a:hlinkClick>
							<xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
								<xsl:attribute name="action">
									<xsl:value-of select="'ppaction://hlinksldjump'"/>
								</xsl:attribute>
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
						</a:hlinkClick>
					</xsl:if>
					<xsl:if test="text:span/text:a">
						<a:hlinkClick>
							<xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
								<xsl:attribute name="action">
									<xsl:value-of select="'ppaction://hlinksldjump'"/>
								</xsl:attribute>
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
						</a:hlinkClick>
					</xsl:if>
				</xsl:for-each>
			</xsl:when>
			<xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
				<xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
					<xsl:if test="text:a">
						<a:hlinkClick>
							<xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
								<xsl:attribute name="action">
									<xsl:value-of select="'ppaction://hlinksldjump'"/>
								</xsl:attribute>
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
						</a:hlinkClick>
					</xsl:if>
					<xsl:if test="text:span/text:a">
						<a:hlinkClick>
							<xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
								<xsl:attribute name="action">
									<xsl:value-of select="'ppaction://hlinksldjump'"/>
								</xsl:attribute>
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
						</a:hlinkClick>
					</xsl:if>
				</xsl:for-each>
			</xsl:when>
			<xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
				<xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
					<xsl:if test="text:a">
						<a:hlinkClick>
							<xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
								<xsl:attribute name="action">
									<xsl:value-of select="'ppaction://hlinksldjump'"/>
								</xsl:attribute>
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
						</a:hlinkClick>
					</xsl:if>
					<xsl:if test="text:span/text:a">
						<a:hlinkClick>
							<xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
								<xsl:attribute name="action">
									<xsl:value-of select="'ppaction://hlinksldjump'"/>
								</xsl:attribute>
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
						</a:hlinkClick>
					</xsl:if>
				</xsl:for-each>
			</xsl:when>
			<xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
				<xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
					<xsl:if test="text:a">
						<a:hlinkClick>
							<xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
								<xsl:attribute name="action">
									<xsl:value-of select="'ppaction://hlinksldjump'"/>
								</xsl:attribute>
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
						</a:hlinkClick>
					</xsl:if>
					<xsl:if test="text:span/text:a">
						<a:hlinkClick>
							<xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
								<xsl:attribute name="action">
									<xsl:value-of select="'ppaction://hlinksldjump'"/>
								</xsl:attribute>
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
						</a:hlinkClick>
					</xsl:if>
				</xsl:for-each>
			</xsl:when>
			<xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
				<xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
					<xsl:if test="text:a">
						<a:hlinkClick>
							<xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
								<xsl:attribute name="action">
									<xsl:value-of select="'ppaction://hlinksldjump'"/>
								</xsl:attribute>
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
						</a:hlinkClick>
					</xsl:if>
					<xsl:if test="text:span/text:a">
						<a:hlinkClick>
							<xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
								<xsl:attribute name="action">
									<xsl:value-of select="'ppaction://hlinksldjump'"/>
								</xsl:attribute>
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
						</a:hlinkClick>
					</xsl:if>
				</xsl:for-each>
			</xsl:when>
			<xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
				<xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
					<xsl:if test="text:a">
						<a:hlinkClick>
							<xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
								<xsl:attribute name="action">
									<xsl:value-of select="'ppaction://hlinksldjump'"/>
								</xsl:attribute>
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
						</a:hlinkClick>
					</xsl:if>
					<xsl:if test="text:span/text:a">
						<a:hlinkClick>
							<xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
								<xsl:attribute name="action">
									<xsl:value-of select="'ppaction://hlinksldjump'"/>
								</xsl:attribute>
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
						</a:hlinkClick>
					</xsl:if>
				</xsl:for-each>
			</xsl:when>
			<xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
				<xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
					<xsl:if test="text:a">
						<a:hlinkClick>
							<xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
								<xsl:attribute name="action">
									<xsl:value-of select="'ppaction://hlinksldjump'"/>
								</xsl:attribute>
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
						</a:hlinkClick>
					</xsl:if>
					<xsl:if test="text:span/text:a">
						<a:hlinkClick>
							<xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
								<xsl:attribute name="action">
									<xsl:value-of select="'ppaction://hlinksldjump'"/>
								</xsl:attribute>
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat($FrameCount,'BTHLFileId',$bulletLevel,'Link',$listItemCount)"/>
								</xsl:attribute>
							</xsl:if>
						</a:hlinkClick>
					</xsl:if>
				</xsl:for-each>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
  <xsl:template name="tmpGroping">
    <xsl:param name="pos"/>
    <xsl:param name="pageNo"/>
    <p:grpSp>
      <p:nvGrpSpPr>
        <p:cNvPr name="Title 1">
          <xsl:attribute name="name">
            <xsl:value-of select="concat('Group ',$pos+1)"/>
          </xsl:attribute>
          <xsl:attribute name="id">
            <xsl:value-of select="$pos+1"/>
          </xsl:attribute>
        </p:cNvPr>
        <p:cNvGrpSpPr>
          <a:grpSpLocks/>
        </p:cNvGrpSpPr>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <xsl:call-template name="tmpgroupingCordinates"/>
      <xsl:for-each select="node()">
        <xsl:variable name="var_num_1">
          <xsl:value-of select="position()"/>
        </xsl:variable>
        <xsl:variable name="var_num_2">
          <xsl:number level="any"/>
        </xsl:variable>
        <xsl:variable name="NvPrId" select="number(concat($pos,$var_num_1,$var_num_2))"/>
        <xsl:choose>
          <!--<xsl:when test="name()='draw:g'">-->
          <!--<xsl:for-each select=".">
              <xsl:call-template name="tmpGroping">
              <xsl:with-param name ="pos" select ="$pos"/>
              <xsl:with-param name ="pageNo" select ="$pageNo"/>
            </xsl:call-template>
            </xsl:for-each>-->
          <!--</xsl:when>-->
          <xsl:when test="name()='draw:frame'">
            <xsl:variable name="var_pos" select="position()"/>

            <!--<xsl:for-each select=".">-->
            <xsl:choose>
              <xsl:when test="./draw:image">
                <xsl:for-each select="./draw:image">
                  <xsl:if test ="contains(@xlink:href,'.png') or contains(@xlink:href,'.emf') or contains(@xlink:href,'.wmf') or contains(@xlink:href,'.jfif') or contains(@xlink:href,'.jpe') 
            or contains(@xlink:href,'.bmp') or contains(@xlink:href,'.dib') or contains(@xlink:href,'.rle')
            or contains(@xlink:href,'.bmz') or contains(@xlink:href,'.gfa') 
            or contains(@xlink:href,'.emz') or contains(@xlink:href,'.wmz') or contains(@xlink:href,'.pcz')
            or contains(@xlink:href,'.tif') or contains(@xlink:href,'.tiff') 
            or contains(@xlink:href,'.cdr') or contains(@xlink:href,'.cgm') or contains(@xlink:href,'.eps') 
            or contains(@xlink:href,'.pct') or contains(@xlink:href,'.pict') or contains(@xlink:href,'.wpg') 
            or contains(@xlink:href,'.jpeg') or contains(@xlink:href,'.gif') or contains(@xlink:href,'.png') or contains(@xlink:href,'.jpg')">
                    <xsl:if test="not(./@xlink:href[contains(.,'../')])">
                      <xsl:call-template name="InsertPicture">
                        <xsl:with-param name="imageNo" select="$pageNo" />
                        <xsl:with-param name="picNo" select="$NvPrId" />
                        <xsl:with-param name ="grpFlag" select="'true'" />
                      </xsl:call-template>
                    </xsl:if>
                  </xsl:if>
                </xsl:for-each>
              </xsl:when>
              <xsl:when test="@presentation:class[contains(.,'title')]
                                  or @presentation:class[contains(.,'subtitle')]
                                  or @presentation:class[contains(.,'outline')]">
              </xsl:when>
              <xsl:when test ="(draw:text-box) and not(@presentation:style-name) and not(@presentation:class)">
                <xsl:call-template name ="CreateShape">
                  <xsl:with-param name ="fileName" select ="'content.xml'"/>
                  <xsl:with-param name ="shapeName" select="'TextBox '" />
                  <xsl:with-param name ="shapeCount" select="$NvPrId" />
                  <xsl:with-param name ="grpFlag" select="'true'" />
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="./draw:plugin">
                <!--<xsl:for-each select="./draw:plugin">
                    <xsl:call-template name="InsertAudio">
                      <xsl:with-param name="imageNo" select="$pageNo" />
                      <xsl:with-param name="AudNo" select="$var_pos" />
                    <xsl:with-param name ="NvPrId" select="$pos + $NvPrId" />
                    </xsl:call-template>
                  </xsl:for-each>-->
              </xsl:when>
            </xsl:choose>
            <!--</xsl:for-each>-->
          </xsl:when>
          <xsl:when test="name()='draw:rect' or name()='draw:ellipse' or name()='draw:custom-shape' 
                or name()='draw:line' or name()='draw:connector'">
            <xsl:variable name="var_pos" select="position()"/>
            <xsl:call-template name ="shapes" >
              <xsl:with-param name ="fileName" select ="'content.xml'"/>
              <xsl:with-param name ="var_pos" select="$NvPrId" />
              <xsl:with-param name ="grpFlag" select="'true'" />
            </xsl:call-template >
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </p:grpSp>
  </xsl:template>
</xsl:stylesheet>
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
    xmlns:msxsl="urn:schemas-microsoft-com:xslt"
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
  xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" 
  xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0"
  exclude-result-prefixes="odf style text number draw page smil anim">

	<xsl:template name ="customAnimation">
		<xsl:param name ="slideId"/>
		<!--<xsl:if test ="anim:par/anim:seq/anim:par">-->

		  <xsl:variable name ="animationVal">
			<xsl:for-each select ="anim:par/anim:seq/anim:par">
				<!-- Added by s1th -->
				<xsl:for-each select="./anim:par">
					
					<xsl:for-each select="./node()">
						<xsl:if test="name()='anim:par' or name()='anim:iterate'">

				<xsl:variable name ="validateAnimation">
					<xsl:call-template name ="validateAnimation"/>
				</xsl:variable>
				<xsl:variable name ="animationType">
					<xsl:call-template name ="animationType"/>
				</xsl:variable>
				<xsl:variable name ="animationId">
					<xsl:call-template name ="animationId">
						<xsl:with-param name ="animationType" select ="$animationType"/>
					</xsl:call-template>
				</xsl:variable>
				<xsl:if test ="$validateAnimation!='false'">
					<xsl:if test ="$animationId!=''">
						<p:par>
							<p:cTn id="3" fill="hold">
								<p:stCondLst>
									<p:cond delay="indefinite"/>
								</p:stCondLst>
								<p:childTnLst>
								
									<p:par>
										<p:cTn id="4" fill="hold">
											<p:stCondLst>
												<p:cond delay="0" />
											</p:stCondLst>
											<p:childTnLst>
												<p:par>
                          <p:cTn id="5" fill="hold">
                            <!-- added by yeswanth , Fix for Animation Start type -->
                            <xsl:attribute name="nodeType">
                              <xsl:choose>
			<xsl:when test="./@presentation:node-type='after-previous'">
                              <xsl:value-of select="'afterEffect'"/>
                                </xsl:when>
			<xsl:when test="./@presentation:node-type='with-previous'">
                                <xsl:value-of select="'withEffect'"/>
                                </xsl:when>
                        <xsl:when test="./@presentation:node-type='on-click'">
                                <xsl:value-of select="'clickEffect'"/>
                                </xsl:when>
                                <xsl:otherwise>
                                  <xsl:value-of select="'clickEffect'"/>
                                </xsl:otherwise>
                              </xsl:choose>
                            </xsl:attribute>
                            <!-- End -->
														<xsl:attribute name ="presetClass">
															<xsl:value-of select ="$animationType"/>
														</xsl:attribute>
														<xsl:attribute name ="presetID">
															<xsl:value-of select ="$animationId"/>
														</xsl:attribute>
														<xsl:variable name ="animationSubId">
															<xsl:call-template name ="animationSubId">
																<xsl:with-param name ="animationType" select="$animationType"/>
																<xsl:with-param name ="animationId" select="$animationId"/>
															</xsl:call-template>
														</xsl:variable>
														<xsl:attribute name ="presetSubtype">
															<xsl:value-of select ="$animationSubId"/>
														</xsl:attribute>
														<p:stCondLst>
                              <p:cond>
                                <xsl:attribute name ="delay">
                                  <xsl:choose >
                                    <!-- commented by chhavi-->
                                    <!--<xsl:when test ="substring-before(@smil:begin,'s') &gt; 0">
                                      <xsl:value-of select ="round(substring-before(@smil:begin,'s')* 1000)"/>-->
                                    <!-- ending here-->
                                    <!-- added by chhavi for delay in custom animation-->
				    <xsl:when test ="substring-before(./@smil:begin,'s') &gt; 0">
				    <xsl:value-of select ="round(substring-before(./@smil:begin,'s')* 1000)"/>
                                      <!--ending here-->
                                    </xsl:when>
                                    <xsl:otherwise>
                                      <xsl:value-of select ="'0'"/>
                                    </xsl:otherwise>
                                  </xsl:choose>
                                </xsl:attribute>
                              </p:cond>
														</p:stCondLst>
																<xsl:if test ="name()='anim:iterate'">
															<p:iterate >
																<xsl:attribute name ="type">
																	<xsl:call-template name ="interateType" >
																				<xsl:with-param name ="itType" select ="./@anim:iterate-type"/>
																	</xsl:call-template>
																</xsl:attribute>
																<p:tmPct>
																	<xsl:attribute name ="val">
																		<xsl:choose >
																					<xsl:when test ="substring-before(./@anim:iterate-interval,'s') &gt; 0 ">
																						<xsl:value-of select ="substring-before(./@anim:iterate-interval,'s') * 100000"/>
																			</xsl:when>
																			<xsl:otherwise >
																				<xsl:value-of select ="'0'"/>
																			</xsl:otherwise>
																		</xsl:choose>
																	</xsl:attribute>
																</p:tmPct >
															</p:iterate>
														</xsl:if>
														<p:childTnLst>
															<xsl:call-template name ="processAnim">
															</xsl:call-template>
														</p:childTnLst>
													</p:cTn>
												</p:par>
											</p:childTnLst>
										</p:cTn>
									</p:par>

								</p:childTnLst>
							</p:cTn>
						</p:par>
										
					</xsl:if>
				</xsl:if>
					
						</xsl:if>
						
					</xsl:for-each>

		
				</xsl:for-each>
				<!-- commented by s1th -->
			</xsl:for-each>
		</xsl:variable>
		<xsl:if test ="msxsl:node-set($animationVal)/p:par">
			
			<!--<xsl:if test =" $animationVal !='' or ($animationVal)">-->
			<p:timing>
				<p:tnLst>
					<p:par>
						<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
							<p:childTnLst>
								<p:seq concurrent="1" nextAc="seek">
									<p:cTn id="2" dur="indefinite" nodeType="mainSeq">
										<p:childTnLst>
											<xsl:copy-of select ="$animationVal"/>
										</p:childTnLst>
									</p:cTn >
									<p:prevCondLst>
										<p:cond evt="onPrev" delay="0">
											<p:tgtEl>
												<p:sldTgt/>
											</p:tgtEl>
										</p:cond>
									</p:prevCondLst>
									<p:nextCondLst>
										<p:cond evt="onNext" delay="0">
											<p:tgtEl>
												<p:sldTgt/>
											</p:tgtEl>
										</p:cond>
									</p:nextCondLst>
								</p:seq >
							</p:childTnLst>
						</p:cTn>
					</p:par >
				</p:tnLst>
			</p:timing>
		</xsl:if>
		<!--</xsl:if>-->
	</xsl:template>
	<xsl:template name ="processAnim">
		<xsl:for-each select =".">
			<!-- change made for emphasis -->
			<!--<xsl:for-each select ="./parent::node()/parent::node()">-->

			<xsl:for-each select ="./node()">
        <xsl:variable name ="varspid">
          <xsl:call-template name ="tmpspTarget"/>
        </xsl:variable >
        <xsl:if test ="$varspid!=''">
          <xsl:choose >
                    <xsl:when test ="name(.)='anim:set'">
              <p:set>
							<p:cBhvr>
								<p:cTn id="6" fill="hold">
									<xsl:attribute name ="dur">
										<xsl:choose >
											<xsl:when test ="./@smil:dur='indefinite'">
												<xsl:value-of select ="'indefinite'"/>
											</xsl:when>
                        <xsl:otherwise>
                          <!-- Added by chhavi for duration -->
                          <xsl:choose >
                            <xsl:when test ="number(substring-before(./@smil:dur,'s')) &gt; 0 ">
												<xsl:value-of select ="round(substring-before(./@smil:dur,'s') * 1000)"/>
                      </xsl:when>
											<xsl:otherwise>
												<xsl:value-of select ="'0'"/>
											</xsl:otherwise>
										</xsl:choose>
                          <!-- ending here-->
                        </xsl:otherwise>
                      </xsl:choose>
									</xsl:attribute>
									<xsl:call-template name ="smilBegin"/>
								</p:cTn>
								<p:tgtEl>
									<p:spTgt>
										<xsl:call-template name ="spTarget"/>
									</p:spTgt>
								</p:tgtEl>
								<p:attrNameLst>
									<p:attrName>
										<xsl:call-template name ="attributeNameList">
										</xsl:call-template>
									</p:attrName>
								</p:attrNameLst>
							</p:cBhvr>
							<p:to>
								<xsl:choose>
									<xsl:when test ="./@smil:attributeName ='color' 
									or ./@smil:attributeName='fill-color'">
										<p:clrVal>
											<xsl:call-template name ="attributeNameValue"/>
										</p:clrVal>
									</xsl:when>
									<xsl:otherwise >
										<xsl:call-template name ="attributeNameValue"/>
									</xsl:otherwise>
								</xsl:choose>
							</p:to>
						</p:set>
					</xsl:when >
					<xsl:when test ="name(.)='anim:animate'">
						<p:anim valueType="num">
							<xsl:if test ="@smil:calcMode">
								<xsl:attribute name ="calcmode">
									<xsl:value-of select ="@smil:calcMode"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="@smil:by">
								<xsl:attribute name ="by">
									<xsl:call-template name ="mapCoordinates">
										<xsl:with-param name ="strVal" select ="@smil:by"/>
									</xsl:call-template>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test =" @smil:attributeName = 'font-size' ">
								<xsl:attribute name ="to">
									<xsl:value-of select ="substring-before(@smil:to,'pt')"/>
								</xsl:attribute>
							</xsl:if >
							<xsl:if test ="@smil:from">
								<xsl:attribute name ="from">
									<xsl:call-template name ="mapCoordinates">
										<xsl:with-param name ="strVal" select ="@smil:from"/>
									</xsl:call-template>
								</xsl:attribute>
								<xsl:attribute name ="to">
									<xsl:call-template name ="mapCoordinates">
										<xsl:with-param name ="strVal" select ="@smil:to"/>
									</xsl:call-template>
								</xsl:attribute>
							</xsl:if>
							<p:cBhvr additive="base">
								<p:cTn id="31" fill="hold" >
									<xsl:attribute name ="dur">
										<xsl:choose >
											<xsl:when test ="substring-before(./@smil:dur,'s') &gt; 0 ">
												<xsl:value-of select ="round(substring-before(./@smil:dur,'s') * 1000)"/>
											</xsl:when>
											<xsl:otherwise >
												<xsl:value-of select ="'0'"/>
											</xsl:otherwise>
										</xsl:choose>
									</xsl:attribute>
									<xsl:if test ="@smil:keySplines">
										<xsl:attribute name ="tmFilter">
											<xsl:value-of select ="@smil:keySplines"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="smil:autoReverse='true' ">
										<xsl:attribute name ="autoRev">
											<xsl:value-of select ="'1'"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="@smil:accelerate">
										<xsl:attribute name ="accel">
											<xsl:value-of select ="round(@smil:accelerate * 100000)"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:call-template name ="smilBegin"/>
								</p:cTn >
								<p:tgtEl>
									<p:spTgt>
										<xsl:call-template name ="spTarget"/>
									</p:spTgt>
								</p:tgtEl>
								<p:attrNameLst>
									<p:attrName>
										<xsl:call-template name ="tavAttributeNameValue">
										</xsl:call-template>
									</p:attrName>
								</p:attrNameLst>
							</p:cBhvr>

							<xsl:if test ="@smil:values or @smil:to">

								<xsl:if test =" @smil:attributeName!='font-size' ">
									<xsl:if test ="not(@smil:from)">
										<xsl:choose>
											<xsl:when test ="@smil:attributeName='color'
												or @smil:attributeName='fill-color' ">
												<xsl:call-template name ="smilValuesRGB"/>
											</xsl:when>
											<xsl:otherwise >
												<p:tavLst>
													<xsl:call-template name ="tavListValues">
													</xsl:call-template>
												</p:tavLst>
											</xsl:otherwise>
										</xsl:choose>
									</xsl:if>
								</xsl:if>
							</xsl:if>
						</p:anim>
					</xsl:when >
					<xsl:when test ="name(.)='anim:transitionFilter'">
						<p:animEffect transition="in" >
							<xsl:variable name ="smilFilter">
								<xsl:call-template name ="smilFilter"/>
							</xsl:variable>
							<xsl:if test ="$smilFilter!=''">
								<xsl:attribute name ="filter">
									<xsl:value-of select ="$smilFilter"/>
								</xsl:attribute>
							</xsl:if>
							<p:cBhvr>
								<p:cTn id="11" >
									<xsl:attribute name ="dur">
										<xsl:choose >
											<xsl:when test ="substring-before(./@smil:dur,'s') &gt; 0 ">
												<xsl:value-of select ="round(substring-before(./@smil:dur,'s') * 1000)"/>
											</xsl:when>
											<xsl:otherwise >
												<xsl:value-of select ="'0'"/>
											</xsl:otherwise>
										</xsl:choose>
									</xsl:attribute>
									<xsl:call-template name ="smilBegin"/>
								</p:cTn >
								<p:tgtEl>
									<p:spTgt>
										<xsl:call-template name ="spTarget"/>
									</p:spTgt>
								</p:tgtEl>
							</p:cBhvr>
						</p:animEffect>
					</xsl:when >
					<xsl:when test ="name(.)='anim:animateColor'">
						<p:animClr>
							<xsl:attribute name ="clrSpc">
								<xsl:value-of select ="@anim:color-interpolation"/>
							</xsl:attribute>
							<p:cBhvr >
								<p:cTn id="10"  fill="hold">
									<xsl:attribute name ="dur">
										<xsl:choose >
											<xsl:when test ="substring-before(./@smil:dur,'s') &gt; 0 ">
												<xsl:value-of select ="round(substring-before(./@smil:dur,'s') * 1000)"/>
											</xsl:when>
											<xsl:otherwise >
												<xsl:value-of select ="'0'"/>
											</xsl:otherwise>
										</xsl:choose>
									</xsl:attribute>
									<xsl:call-template name ="smilBegin"/>
								</p:cTn >
								<p:tgtEl>
									<p:spTgt>
										<xsl:call-template name ="spTarget"/>
									</p:spTgt>
								</p:tgtEl>
								<p:attrNameLst>
									<p:attrName>
										<xsl:call-template name ="attributeNameList"/>
									</p:attrName>
								</p:attrNameLst>
							</p:cBhvr>
							<xsl:if test ="@smil:to">
								<p:to>
									<xsl:call-template name ="attributeNameValue"/>
								</p:to>
							</xsl:if>
							<xsl:if test ="@smil:by">
								<p:by>
									<p:hsl>
										<xsl:variable name ="hslH">
											<xsl:value-of select ="substring-after(@smil:by,'hsl(')"/>
										</xsl:variable>
										<xsl:attribute name ="h">
											<xsl:value-of select ="round(substring-before($hslH,',') * 60000)"/>
										</xsl:attribute>
										<xsl:variable name ="hsls">
											<xsl:value-of select ="substring-after($hslH,',')"/>
										</xsl:variable>
										<xsl:attribute name ="s">
											<xsl:value-of select ="round(substring-before($hsls,'%') * 1000) "/>
										</xsl:attribute>
										<xsl:variable name ="hsll">
											<xsl:value-of select ="substring-after($hsls,',')"/>
										</xsl:variable>
										<xsl:attribute name ="l">
											<xsl:value-of select ="round(substring-before($hsll,'%') * 1000)"/>
										</xsl:attribute>
									</p:hsl >
								</p:by>
							</xsl:if>
						</p:animClr>
					</xsl:when>
					<xsl:when test ="name(.)='anim:animateTransform'">
						<xsl:if test ="@svg:type='scale'">
							<p:animScale>
								<p:cBhvr>
									<p:cTn id="14" fill="hold">
										<xsl:attribute name ="dur">
											<xsl:choose >
												<xsl:when test ="substring-before(./@smil:dur,'s') &gt; 0 ">
													<xsl:value-of select ="round(substring-before(./@smil:dur,'s') * 1000)"/>
												</xsl:when>
												<xsl:otherwise >
													<xsl:value-of select ="'0'"/>
												</xsl:otherwise>
											</xsl:choose>
										</xsl:attribute>
										<xsl:call-template name ="smilBegin"/>
									</p:cTn >
									<p:tgtEl>
										<p:spTgt>
											<xsl:call-template name ="spTarget"/>
										</p:spTgt>
									</p:tgtEl>
								</p:cBhvr>
								<xsl:choose>
									<xsl:when test ="@smil:from">
										<p:from>
											<xsl:attribute name ="x">
												<xsl:value-of select ="round(substring-before(@smil:from,',') * 100000)"/>
											</xsl:attribute>
											<xsl:attribute name ="y">
												<xsl:value-of select ="round(substring-after(@smil:from,',') * 100000)"/>
											</xsl:attribute>
										</p:from >
										<p:to>
											<xsl:attribute name ="x">
												<xsl:value-of select ="round(substring-before(@smil:to,',') * 100000)"/>
											</xsl:attribute>
											<xsl:attribute name ="y">
												<xsl:value-of select ="round(substring-after(@smil:to,',') * 100000)"/>
											</xsl:attribute>
										</p:to >
									</xsl:when>
									<xsl:otherwise >
										<!--  Added by vijayeta, bug number 1775269,
                        and the 'if' condition to test if '@smil:by' is present, is added, which is the bug fix for 1775523, 
                        date: 20th Aug '07, flash bulb type , both the bugs were related to roundtrip crash in output pptx-->
										<xsl:if test ="@smil:by">
											<p:by>
												<xsl:attribute name="x">
													<xsl:value-of select="round(substring-before(@smil:by,',') * 100000)" />
												</xsl:attribute>
												<xsl:attribute name="y">
													<xsl:value-of select="round(substring-after(@smil:by,',') * 100000)" />
												</xsl:attribute>
											</p:by>
										</xsl:if>
										<!--  Added by vijayeta, bug number 1775269, date: 20th Aug '07, flash bulb type -->
										<!--<p:by>
										<xsl:attribute name ="x">
											<xsl:value-of select ="round(substring-before(@smil:to,',') * 100000)"/>
										</xsl:attribute>
										<xsl:attribute name ="y">
											<xsl:value-of select ="round(substring-after(@smil:to,',') * 100000)"/>
										</xsl:attribute>
									</p:by >-->
									</xsl:otherwise>
								</xsl:choose>
							</p:animScale>
						</xsl:if>
						<xsl:if test ="@svg:type='rotate'">
							<p:animRot>
								<xsl:attribute name="by">
									<xsl:value-of select ="round(@smil:by * 60000)"/>
								</xsl:attribute>
								<p:cBhvr>
									<p:cTn id="18"  fill="hold">
										<xsl:attribute name ="dur">
											<xsl:choose >
												<xsl:when test ="substring-before(./@smil:dur,'s') &gt; 0 ">
													<xsl:value-of select ="round(substring-before(./@smil:dur,'s') * 1000)"/>
												</xsl:when>
												<xsl:otherwise >
													<xsl:value-of select ="'0'"/>
												</xsl:otherwise>
											</xsl:choose>
										</xsl:attribute>
										<xsl:call-template name ="smilBegin"/>
									</p:cTn >
									<p:tgtEl>
										<p:spTgt>
											<xsl:call-template name ="spTarget"/>
										</p:spTgt>
									</p:tgtEl>
									<p:attrNameLst>
										<p:attrName>r</p:attrName>
									</p:attrNameLst>
								</p:cBhvr>
							</p:animRot>
						</xsl:if >
					</xsl:when>
			  <!-- previously  ./anim:par/node() was there in name()-->
					<xsl:when test ="name(.)='anim:animateMotion'">
						<p:animMotion origin="layout" pathEditMode="relative" ptsTypes="">
							<xsl:attribute name ="path">
								<xsl:value-of select ="@svg:path"/>
							</xsl:attribute>
							<p:cBhvr>
								<p:cTn id="108" fill="hold">
									<xsl:attribute name ="dur">
										<xsl:choose >
											<xsl:when test ="substring-before(./@smil:dur,'s') &gt; 0 ">
												<xsl:value-of select ="round(substring-before(./@smil:dur,'s') * 1000)"/>
											</xsl:when>
											<xsl:otherwise >
												<xsl:value-of select ="'0'"/>
											</xsl:otherwise>
										</xsl:choose>
									</xsl:attribute>
									<xsl:attribute name ="decel">
										<xsl:choose >
											<xsl:when test ="@smil:decelerate &gt; 0 ">
												<xsl:value-of select ="round(@smil:decelerate * 100000)"/>
											</xsl:when>
											<xsl:otherwise >
												<xsl:value-of select ="'0'"/>
											</xsl:otherwise>
										</xsl:choose>
									</xsl:attribute>
									<xsl:call-template name ="smilBegin"/>
								</p:cTn>
								<p:tgtEl>
									<p:spTgt>
										<xsl:call-template name ="spTarget"/>
									</p:spTgt>
								</p:tgtEl>
								<p:attrNameLst>
									<p:attrName>ppt_x</p:attrName>
									<p:attrName>ppt_y</p:attrName>
								</p:attrNameLst>
							</p:cBhvr>
						</p:animMotion>
					</xsl:when>
				</xsl:choose >
        </xsl:if>
			</xsl:for-each >
		</xsl:for-each >
	</xsl:template>
	<xsl:template name ="smilValuesRGB">
		<p:tavLst>
			<p:tav>
				<xsl:attribute name ="tm">
					<!--<xsl:value-of select ="round(@smil:keyTimes * 1000)"/>-->
					<xsl:choose >
						<xsl:when test ="substring-before(@smil:keyTimes,';') &gt; 0">
							<xsl:value-of select ="round(substring-before(@smil:keyTimes,';') * 1000)"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select ="'0'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
				<p:val>
					<p:clrVal>
						<a:srgbClr >
							<xsl:attribute name ="val">
								<xsl:value-of select ="substring-after(substring-before(@smil:values,';'),'#')"/>
							</xsl:attribute>
						</a:srgbClr >
					</p:clrVal>
				</p:val>
			</p:tav>
			<p:tav >
				<xsl:attribute name ="tm">
					<xsl:choose >
						<xsl:when test ="substring-after(@smil:keyTimes,';') &gt; 0">
							<xsl:value-of select ="round(substring-after(@smil:keyTimes,';') * 1000)"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select ="'0'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
				<p:val>
					<p:clrVal>
						<a:srgbClr >
							<xsl:attribute name ="val">
								<xsl:value-of select ="substring-after(substring-after(@smil:values,';'),'#')"/>
							</xsl:attribute >
						</a:srgbClr >
					</p:clrVal>
				</p:val>
			</p:tav>
		</p:tavLst>
	</xsl:template>
	<!-- changed here for ca -->
	<xsl:template name ="smilBegin">
		<xsl:if test ="./@smil:begin">
			<!--<xsl:if test ="./parent::node()/@smil:begin">-->
			<p:stCondLst>
				<p:cond >
					<xsl:attribute name ="delay">
						<xsl:choose >
							<xsl:when test ="substring-before(./@smil:begin,'s') &gt; 0">
								<!--<xsl:when test ="substring-before(./parent::node()/@smil:begin,'s') &gt; 0">-->
								<xsl:value-of select ="round(substring-before(./@smil:begin,'s')* 1000)"/>
								<!--<xsl:value-of select ="round(substring-before(./parent::node()/@smil:begin,'s')* 1000)"/>-->
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select ="'0'"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
				</p:cond >
			</p:stCondLst>
		</xsl:if>
	</xsl:template>
	<xsl:template name ="validateAnimation">

		<xsl:variable name ="nvPrId">
			<xsl:call-template name ="getnvPrIdval">
				<xsl:with-param name ="spId">
					<xsl:choose >
						<xsl:when test ="./child::node()[1]/@smil:targetElement">
							<xsl:value-of select ="./child::node()[1]/@smil:targetElement"/>
						</xsl:when>
						<xsl:when test ="./@smil:targetElement">
							<xsl:value-of select ="./@smil:targetElement"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="'id1'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:with-param >
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name ="nvPrIdPara">
			<xsl:call-template name ="getParaIdval">
				<xsl:with-param name ="spId">
					<xsl:choose >
						<xsl:when test ="child::node()[1]/@smil:targetElement">
							<xsl:value-of select ="child::node()[1]/@smil:targetElement"/>
						</xsl:when>
						<xsl:when test ="./@smil:targetElement">
							<xsl:value-of select ="./@smil:targetElement"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="'id1'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:with-param>
			</xsl:call-template>
		</xsl:variable>
		<xsl:if test ="$nvPrId='' and $nvPrIdPara=''">
			<xsl:value-of select ="'false'"/>
		</xsl:if>
	</xsl:template>
	<xsl:template name ="spTarget">
		<xsl:attribute name ="spid">
			<xsl:call-template name ="getnvPrId">
				<xsl:with-param name ="spId">
					<xsl:choose >
						<xsl:when test ="./@smil:targetElement">
							<xsl:value-of select ="./@smil:targetElement"/>
						</xsl:when>
						<xsl:when test ="parent::node()/@smil:targetElement">
							<xsl:value-of select ="parent::node()/@smil:targetElement"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="'id1'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:with-param >
			</xsl:call-template>
		</xsl:attribute>
		<xsl:variable name ="spId">
			<xsl:call-template name ="getParaId">
				<xsl:with-param name ="spId">
					<xsl:choose >
						<xsl:when test ="./@smil:targetElement">
							<xsl:value-of select ="./@smil:targetElement"/>
						</xsl:when>
						<xsl:when test ="parent::node()/@smil:targetElement">
							<xsl:value-of select ="parent::node()/@smil:targetElement"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="'id1'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:with-param>
			</xsl:call-template>
		</xsl:variable>
		<xsl:if test ="$spId!=''">
			<p:txEl>
				<p:pRg  >
					<xsl:attribute name ="st">
						<xsl:value-of select ="$spId"/>
					</xsl:attribute>
					<xsl:attribute name ="end">
						<xsl:value-of select ="$spId"/>
					</xsl:attribute>
				</p:pRg >
			</p:txEl>
		</xsl:if>
	</xsl:template>
  <xsl:template name ="tmpspTarget">
    <xsl:variable name ="spId">
      <xsl:call-template name ="getnvPrId">
        <xsl:with-param name ="spId">
          <xsl:choose >
            <xsl:when test ="./parent::node()/@smil:targetElement">
              <xsl:value-of select ="./parent::node()/@smil:targetElement"/>
            </xsl:when>
            <xsl:when test ="./@smil:targetElement">
              <xsl:value-of select ="./@smil:targetElement"/>
            </xsl:when>
	    <xsl:otherwise>
	      <xsl:value-of select="'id1'"/>
	    </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param >
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select ="$spId"/>
  </xsl:template>
	<xsl:template name ="tavListValues">
		<xsl:call-template name ="addTavListNode">
			<xsl:with-param name ="string" select ="./@smil:values"/>
			<xsl:with-param name ="smilVal" select ="./@smil:keyTimes"/>
		</xsl:call-template >
	</xsl:template>

	<xsl:template name ="tavAttributeNameValue">
		<xsl:choose>
			<xsl:when test ="./@smil:attributeName ='x'">
				<xsl:value-of select ="'ppt_x'"/>
			</xsl:when>
			<xsl:when test ="./@smil:attributeName ='y'">
				<xsl:value-of select ="'ppt_y'"/>
			</xsl:when>
			<xsl:when test ="./@smil:attributeName ='width'">
				<xsl:value-of select ="'ppt_w'"/>
			</xsl:when>
			<xsl:when test ="./@smil:attributeName ='height'">
				<xsl:value-of select ="'ppt_h'"/>
			</xsl:when>
			<xsl:otherwise >
				<xsl:call-template name ="attributeNameList" />
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name ="attributeNameList">
		<xsl:choose >
			<xsl:when test ="./@smil:attributeName ='x'">
				<xsl:value-of select ="'ppt_x'"/>
			</xsl:when>
			<xsl:when test ="./@smil:attributeName ='y'">
				<xsl:value-of select ="'ppt_y'"/>
			</xsl:when>
			<xsl:when test ="./@smil:attributeName ='width'">
				<xsl:value-of select ="'ppt_w'"/>
			</xsl:when>
			<xsl:when test ="./@smil:attributeName ='height'">
				<xsl:value-of select ="'ppt_h'"/>
			</xsl:when>
			<xsl:when test ="./@smil:attributeName ='visibility'">
				<xsl:value-of select ="'style.visibility'"/>
			</xsl:when>
			<xsl:when test ="./@smil:attributeName='font-family'">
				<xsl:value-of select ="'style.fontFamily'"/>
			</xsl:when>
			<xsl:when test ="./@smil:attributeName='color'">
				<xsl:value-of select ="'style.color'"/>
			</xsl:when>
			<xsl:when test ="./@smil:attributeName='fill'">
				<xsl:value-of select ="'fill.type'"/>
			</xsl:when>
			<xsl:when test ="./@smil:attributeName='fill-color'">
				<xsl:value-of select ="'fillcolor'"/>
			</xsl:when>
			<xsl:when test ="./@smil:attributeName='font-weight'">
				<xsl:value-of select ="'style.fontWeight'"/>
			</xsl:when>
			<xsl:when test ="./@smil:attributeName='text-underline'">
				<xsl:value-of select ="'style.textDecorationUnderline'"/>
			</xsl:when>
			<xsl:when test ="./@smil:attributeName='opacity'">
				<xsl:value-of select ="'style.opacity'"/>
			</xsl:when>
			<xsl:when test ="./@smil:attributeName='stroke-color'">
				<xsl:value-of select ="'stroke.color'"/>
			</xsl:when>
			<xsl:when test ="./@smil:attributeName='font-style'">
				<xsl:value-of select ="'style.fontStyle'"/>
			</xsl:when>
			<xsl:when test ="./@smil:attributeName='font-size'">
				<xsl:value-of select ="'style.fontSize'"/>
			</xsl:when>

		</xsl:choose>
	</xsl:template>

	<xsl:template name ="attributeNameValue">
		<xsl:choose >
			<xsl:when test ="./@smil:to='hidden'">
				<p:strVal val="hidden"/>
			</xsl:when>
			<xsl:when test ="./@smil:to='visible'">
				<p:strVal val="visible"/>
			</xsl:when>
			<xsl:when  test ="./@anim:color-interpolation='rgb'">
				<a:srgbClr val="{substring-after(./@smil:to,'#')}" />
			</xsl:when>
			<xsl:when  test ="./@smil:to='opacity'">
				<p:strVal val="'./@smil:to'" />
			</xsl:when>
			<xsl:when test ="./@smil:to='font-weight'">
				<p:strVal val="'style.fontWeight"/>
			</xsl:when>
			<xsl:when test ="./@smil:attributeName='color' or ./@smil:attributeName='fill-color'">
				<a:srgbClr val="{substring-after(./@smil:to,'#')}" />
			</xsl:when>
			<xsl:when test ="./@smil:attributeName='text-underline'">
				<p:strVal val="true"/>
			</xsl:when >
			<xsl:otherwise >
				<p:strVal val="{./@smil:to}"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<xsl:template name ="getnvPrId">
		<xsl:param name ="spId"/>
		<xsl:variable name ="varSpid">
			<xsl:for-each select ="./parent::node()/parent::node()/parent::node()/parent::node()/parent::node()/parent::node()">
        <xsl:if test="position()=1">
        <xsl:for-each select ="node()">
          <xsl:if test ="name()='draw:rect' or name()='draw:ellipse'
                  or name()='draw:custom-shape' or  name()='draw:circle'
                  or name()='draw:g' or name()='draw:frame'">
            <xsl:choose>
              <xsl:when test="name()='draw:frame'">
                <xsl:choose>
                  <xsl:when test="./draw:image">
                    <xsl:if test ="contains(./draw:image/@xlink:href,'.png') or contains(./draw:image/@xlink:href,'.emf') 
                      or contains(./draw:image/@xlink:href,'.wmf') or contains(./draw:image/@xlink:href,'.jfif') or contains(./draw:image/@xlink:href,'.jpe') 
            or contains(./draw:image/@xlink:href,'.bmp') or contains(./draw:image/@xlink:href,'.dib')
            or contains(./draw:image/@xlink:href,'.rle')
            or contains(./draw:image/@xlink:href,'.bmz') or contains(./draw:image/@xlink:href,'.gfa') 
            or contains(./draw:image/@xlink:href,'.emz') or contains(./draw:image/@xlink:href,'.wmz') 
            or contains(./draw:image/@xlink:href,'.pcz')
            or contains(./draw:image/@xlink:href,'.tif') or contains(./draw:image/@xlink:href,'.tiff') 
            or contains(./draw:image/@xlink:href,'.cdr') or contains(./draw:image/@xlink:href,'.cgm')
            or contains(./draw:image/@xlink:href,'.eps') 
            or contains(./draw:image/@xlink:href,'.pct') or contains(./draw:image/@xlink:href,'.pict') 
            or contains(./draw:image/@xlink:href,'.wpg') 
            or contains(./draw:image/@xlink:href,'.jpeg') or contains(./draw:image/@xlink:href,'.gif') 
            or contains(./draw:image/@xlink:href,'.png') or contains(./draw:image/@xlink:href,'.jpg')">
                      <xsl:if test="not(./draw:image/@xlink:href[contains(.,'../')])">
                        <xsl:call-template name="tmpgetNvPrID">
                          <xsl:with-param name="spId" select="$spId"/>
                        </xsl:call-template>
                      </xsl:if>
                    </xsl:if>
                  </xsl:when>
                  <xsl:when test="@presentation:class[contains(.,'title')]
                                  or @presentation:class[contains(.,'subtitle')]
                                  or @presentation:class[contains(.,'outline')]">
                    <xsl:call-template name="tmpgetNvPrID">
                      <xsl:with-param name="spId" select="$spId"/>
                    </xsl:call-template>

                  </xsl:when>
                  <xsl:when test ="(draw:text-box) and not(@presentation:class)">
                    <xsl:call-template name="tmpgetNvPrID">
                      <xsl:with-param name="spId" select="$spId"/>
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:when test="./draw:plugin">
                  </xsl:when>
                </xsl:choose>
              </xsl:when>
              <xsl:when test="name()='draw:custom-shape'">
                <xsl:variable name ="shapeName">
                  <xsl:call-template name="tmpgetCustShapeType"/>
                </xsl:variable>
                <xsl:if test="$shapeName != ''">
                  <xsl:call-template name="tmpgetNvPrID">
                    <xsl:with-param name="spId" select="$spId"/>
                  </xsl:call-template>
                </xsl:if>
              </xsl:when>
              <xsl:when test="name()='draw:rect'">
                <xsl:call-template name="tmpgetNvPrID">
                  <xsl:with-param name="spId" select="$spId"/>
                </xsl:call-template>
              </xsl:when>
				<xsl:when test="name()='draw:g'">
					<xsl:call-template name="tmpgetNvPrID">
						<xsl:with-param name="spId" select="$spId"/>
					</xsl:call-template>
				</xsl:when>
              <xsl:when test="name()='draw:ellipse' or name()='draw:circle'">
                <xsl:if test="not(@draw:kind)">
                <xsl:call-template name="tmpgetNvPrID">
                  <xsl:with-param name="spId" select="$spId"/>
                </xsl:call-template>
                </xsl:if>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
          
        </xsl:for-each>
        </xsl:if>
	</xsl:for-each >
		</xsl:variable>
		<xsl:value-of select ="$varSpid"/>
	</xsl:template>
	<xsl:template name ="getParaId">
		<xsl:param name ="spId"/>
		<xsl:variable name ="varSpid">
			<xsl:for-each select ="./parent::node()/parent::node()/parent::node()/parent::node()/parent::node()/parent::node()">
				<xsl:for-each select ="node()">
					<xsl:variable name ="nvPrId">
						<xsl:value-of select ="position()"/>
					</xsl:variable>
					<xsl:variable name ="paraId">
						<xsl:for-each select =".//text:p">
							<xsl:if test ="$spId =@text:id">
								<xsl:value-of select ="position()"/>
							</xsl:if>
						</xsl:for-each>
					</xsl:variable>
					<xsl:if test ="$paraId!=''">
						<xsl:value-of select ="$paraId -1"/>
					</xsl:if>
				</xsl:for-each>
			</xsl:for-each>
		</xsl:variable>
		<xsl:value-of select ="$varSpid"/>
	</xsl:template>

	<xsl:template name="addTavListNode">
		<xsl:param name="string"/>
		<xsl:param name ="smilVal"/>

		<xsl:variable name="first" select="substring-before($string,';')"/>
		<xsl:variable name="rest" select="substring-after($string,';')"/>

		<xsl:variable name="smilfirst" select="substring-before($smilVal,';')"/>
		<xsl:variable name="smilrest" select="substring-after($smilVal,';')"/>

		<xsl:if test='$first'>
			<p:tav>
				<xsl:attribute name ="tm">
					<xsl:choose >
						<xsl:when test ="$smilfirst &gt; 0 ">
							<xsl:value-of select ="round($smilfirst * 100000)"/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select ="'0'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
				<xsl:if test ="@anim:formula">
					<xsl:attribute name ="fmla">
						<xsl:call-template name ="mapCoordinates">
							<xsl:with-param name ="strVal" select ="@anim:formula"/>
						</xsl:call-template>
					</xsl:attribute>
				</xsl:if>
				<p:val>
					<p:strVal>
						<xsl:attribute name ="val">
							<xsl:call-template name ="mapCoordinates">
								<xsl:with-param name ="strVal" select ="$first"/>
							</xsl:call-template>
						</xsl:attribute>
					</p:strVal>
				</p:val>
			</p:tav>
		</xsl:if>
		<xsl:if test='$rest'>
			<xsl:call-template name="addTavListNode">
				<xsl:with-param name="string" select="$rest"/>
				<xsl:with-param name ="smilVal" select ="$smilrest"/>
			</xsl:call-template>
		</xsl:if>
		<xsl:if test='not($rest)'>
			<p:tav >
				<xsl:attribute name ="tm">
					<xsl:choose >
						<xsl:when test ="$smilVal &gt; 0 ">
							<xsl:value-of select ="round($smilVal * 100000)"/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select ="'0'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
				<p:val>
					<p:strVal>
						<xsl:attribute name ="val">
							<xsl:call-template name ="mapCoordinates">
								<xsl:with-param name ="strVal" select ="$string"/>
							</xsl:call-template>
						</xsl:attribute >
					</p:strVal>
				</p:val>
			</p:tav>
		</xsl:if>
	</xsl:template>
	<xsl:template name ="animationType">
		<xsl:choose>
			<xsl:when test ="./@presentation:preset-class='entrance' or ./@presentation:preset-class='entrance'">
				<xsl:value-of select ="'entr'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-class='exit' or ./@presentation:preset-class='exit'">
				<xsl:value-of select ="'exit'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-class='emphasis' or ./@presentation:preset-class='emphasis'">
				<xsl:value-of select ="'emph'"/>
			</xsl:when>

		</xsl:choose>
	</xsl:template>

	<xsl:template name ="animationId">
		<xsl:param name ="animationType"/>
		<xsl:choose >
			<xsl:when test ="$animationType ='entr'">
				<xsl:choose >
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-appear'">
						<xsl:value-of select ="1"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-fly-in'">
						<xsl:value-of select ="'2'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-venetian-blinds'">
						<xsl:value-of select ="'3'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-box'">
						<xsl:value-of select ="'4'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-checkerboard'">
						<xsl:value-of select ="'5'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-circle'">
						<xsl:value-of select ="'6'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-fly-in-slow'">
						<xsl:value-of select ="'7'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-diamond'">
						<xsl:value-of select ="'8'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-dissolve-in'">
						<xsl:value-of select ="'9'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-fade-in'">
						<xsl:value-of select ="'10'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-flash-once'">
						<xsl:value-of select ="'11'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-peek-in'">
						<xsl:value-of select ="'12'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-plus'">
						<xsl:value-of select ="'13'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-random-bars'">
						<xsl:value-of select ="'14'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-random'">
						<xsl:value-of select ="'24'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-split'">
						<xsl:value-of select ="'16'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-diagonal-squares'">
						<xsl:value-of select ="'18'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-wedge'">
						<xsl:value-of select ="'20'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-wheel'">
						<xsl:value-of select ="'21'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-wipe'">
						<xsl:value-of select ="'22'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-expand'">
						<xsl:value-of select ="'55'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-fade-in-and-swivel'">
						<xsl:value-of select ="'45'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-fade-in-and-zoom'">
						<xsl:value-of select ="'53'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-ascend'">
						<xsl:value-of select ="'42'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-center-revolve'">
						<xsl:value-of select ="'43'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-colored-lettering'">
						<xsl:value-of select ="'27'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-compress'">
						<xsl:value-of select ="'50'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-descend'">
						<xsl:value-of select ="'47'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-ease-in'">
						<xsl:value-of select ="'29'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-turn-and-grow'">
						<xsl:value-of select ="'31'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-rise-up'">
						<xsl:value-of select ="'37'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-spin-in'">
						<xsl:value-of select ="'49'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-stretchy'">
						<xsl:value-of select ="'17'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-swivel'">
						<xsl:value-of select ="'19'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-unfold'">
						<xsl:value-of select ="'40'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-zoom'">
						<xsl:value-of select ="'23'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-boomerang'">
						<xsl:value-of select ="'25'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-bounce'">
						<xsl:value-of select ="'26'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-movie-credits'">
						<xsl:value-of select ="'28'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-curve-up'">
						<xsl:value-of select ="'52'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-flip'">
						<xsl:value-of select ="'56'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-float'">
						<xsl:value-of select ="'30'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-fold'">
						<xsl:value-of select ="'58'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-glide'">
						<xsl:value-of select ="'54'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-breaks'">
						<xsl:value-of select ="'34'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-magnify'">
						<xsl:value-of select ="'51'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-pinwheel'">
						<xsl:value-of select ="'35'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-sling'">
						<xsl:value-of select ="'48'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-spiral-in'">
						<xsl:value-of select ="'15'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-falling-in'">
						<xsl:value-of select ="'38'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-thread'">
						<xsl:value-of select ="'39'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-entrance-whip'">
						<xsl:value-of select ="'41'"/>
					</xsl:when>
				</xsl:choose>
			</xsl:when>
			<xsl:when test ="$animationType ='exit'">
				<xsl:choose >
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-venetian-blinds'">
						<xsl:value-of select ="'3'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-box'">
						<xsl:value-of select ="'4'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-checkerboard'">
						<xsl:value-of select ="'5'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-circle'">
						<xsl:value-of select ="'6'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-crawl-out'">
						<xsl:value-of select ="'7'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-diamond'">
						<xsl:value-of select ="'8'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-disappear'">
						<xsl:value-of select ="'1'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-dissolve'">
						<xsl:value-of select ="'9'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-flash-once'">
						<xsl:value-of select ="'11'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-fly-out'">
						<xsl:value-of select ="'2'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-peek-out'">
						<xsl:value-of select ="'12'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-plus'">
						<xsl:value-of select ="'13'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-random-bars'">
						<xsl:value-of select ="'14'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-random'">
						<xsl:value-of select ="'24'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-split'">
						<xsl:value-of select ="'16'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-diagonal-squares'">
						<xsl:value-of select ="'18'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-wheel'">
						<xsl:value-of select ="'21'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-wipe'">
						<xsl:value-of select ="'22'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-contract'">
						<xsl:value-of select ="'55'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-fade-out'">
						<xsl:value-of select ="'10'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-fade-out-and-swivel'">
						<xsl:value-of select ="'45'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-fade-out-and-zoom'">
						<xsl:value-of select ="'53'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-ascend'">
						<xsl:value-of select ="'47'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-center-revolve'">
						<xsl:value-of select ="'43'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-collapse'">
						<xsl:value-of select ="'17'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-colored-lettering'">
						<xsl:value-of select ="'27'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-descend'">
						<xsl:value-of select ="'42'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-ease-out'">
						<xsl:value-of select ="'29'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-turn-and-grow'">
						<xsl:value-of select ="'31'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-sink-down'">
						<xsl:value-of select ="'37'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-spin-out'">
						<xsl:value-of select ="'49'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-stretchy'">
						<xsl:value-of select ="'50'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-unfold'">
						<xsl:value-of select ="'40'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-zoom'">
						<xsl:value-of select ="'23'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-boomerang'">
						<xsl:value-of select ="'25'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-bounce'">
						<xsl:value-of select ="'26'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-movie-credits'">
						<xsl:value-of select ="'28'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-curve-down'">
						<xsl:value-of select ="'52'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-flip'">
						<xsl:value-of select ="'56'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-float'">
						<xsl:value-of select ="'30'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-fold'">
						<xsl:value-of select ="'58'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-glide'">
						<xsl:value-of select ="'54'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-breaks'">
						<xsl:value-of select ="'34'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-magnify'">
						<xsl:value-of select ="'51'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-pinwheel'">
						<xsl:value-of select ="'35'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-sling'">
						<xsl:value-of select ="'48'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-spiral-out'">
						<xsl:value-of select ="'15'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-swish'">
						<xsl:value-of select ="'38'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-swivel'">
						<xsl:value-of select ="'19'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-thread'">
						<xsl:value-of select ="'39'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-whip'">
						<xsl:value-of select ="'41'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-exit-wedge'">
						<xsl:value-of select ="'20'"/>
					</xsl:when>
				</xsl:choose>
			</xsl:when>
			<xsl:when test ="$animationType ='emph'">
				<xsl:choose >
          <xsl:when test ="./@presentation:preset-id ='ooo-emphasis-fill-color'">
            <xsl:value-of select ="'1'"/>
          </xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-font'">
						<xsl:value-of select ="'2'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-font-color'">
						<xsl:value-of select ="'3'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-font-size'">
						<xsl:value-of select ="'4'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-font-style'">
						<xsl:value-of select ="'5'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-grow-and-shrink'">
						<xsl:value-of select ="'6'"/>
					</xsl:when>
          <xsl:when test ="./@presentation:preset-id ='ooo-emphasis-line-color'">
            <xsl:value-of select ="'7'"/>
          </xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-spin'">
						<xsl:value-of select ="'8'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-transparency'">
						<xsl:value-of select ="'9'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-bold-flash'">
						<xsl:value-of select ="'10'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-color-over-by-word'">
						<xsl:value-of select ="'16'"/>
					</xsl:when>					
					<!--  Added by vijayeta, bug number 1775269, date: 20th Aug '07-->
					<!-- slide 1 of input file 'animation partial.odp' -->
					<xsl:when test="./@presentation:preset-id ='ooo-emphasis-color-blend'">
						<xsl:value-of select="'19'" />
					</xsl:when>
					<!-- slide 2 of input file 'animation partial.odp' -->
					<xsl:when test="./@presentation:preset-id ='ooo-emphasis-color-over-by-letter'">
						<xsl:value-of select="'20'" />
					</xsl:when>
					<!-- slide 10 of input file 'animation partial.odp'-->
					<xsl:when test="./@presentation:preset-id ='ooo-emphasis-reveal-underline'">
						<xsl:value-of select="'18'" />
					</xsl:when>
					<!--  Added by vijayeta, bug number 1775269, date: 20th Aug '07 -->

					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-color-over-by-letter'">
						<xsl:value-of select ="'20'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-complementary-color'">
						<xsl:value-of select ="'21'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-complementary-color-2'">
						<xsl:value-of select ="'22'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-contrasting-color'">
						<xsl:value-of select ="'23'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-darken'">
						<xsl:value-of select ="'24'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-desaturate'">
						<xsl:value-of select ="'25'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-flash-bulb'">
						<xsl:value-of select ="'26'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-lighten'">
						<xsl:value-of select ="'30'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-vertical-highlight'">
						<xsl:value-of select ="'33'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-flicker'">
						<xsl:value-of select ="'27'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-grow-with-color'">
						<xsl:value-of select ="'28'"/>
					</xsl:when>
          <xsl:when test ="./@presentation:preset-id ='ooo-emphasis-shimmer'">
						<xsl:value-of select ="'36'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-teeter'">
						<xsl:value-of select ="'32'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-blast'">
						<xsl:value-of select ="'14'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-blink'">
						<xsl:value-of select ="'35'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-bold-reveal'">
						<xsl:value-of select ="'15'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-style-emphasis'">
						<xsl:value-of select ="'31'"/>
					</xsl:when>
					<xsl:when test ="./@presentation:preset-id ='ooo-emphasis-wave'">
						<xsl:value-of select ="'34'"/>
					</xsl:when>
				</xsl:choose>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<xsl:template name ="animationSubId">
		<xsl:param name ="animationType"/>
		<xsl:param name ="animationId"/>
		<xsl:choose>
			<xsl:when test ="./@presentation:preset-sub-type ='from-top'">
				<xsl:value-of select ="'1'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='1'">
				<xsl:value-of select ="'1'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='from-right'">
				<xsl:value-of select ="'2'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='2'">
				<xsl:value-of select ="'2'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='from-top-right'">
				<xsl:value-of select ="'3'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='right-to-top'">
				<xsl:value-of select ="'3'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='3'">
				<xsl:value-of select ="'3'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='from-bottom'">
				<xsl:value-of select ="'4'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='4'">
				<xsl:value-of select ="'4'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='horizontal'">
				<xsl:value-of select ="'10'"/>
				<!--<xsl:value-of select ="'5'"/>-->
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='downward'">
				<xsl:value-of select ="'5'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='from-bottom-right'">
				<xsl:value-of select ="'6'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='right-to-bottom'">
				<xsl:value-of select ="'6'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='from-left'">
				<xsl:value-of select ="'8'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='8'">
				<xsl:value-of select ="'8'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='20'">
				<xsl:value-of select ="'20'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='544'">
				<xsl:value-of select ="'544'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='from-top-left'">
				<xsl:value-of select ="'9'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='left-to-top'">
				<xsl:value-of select ="'9'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='vertical'">
				<xsl:value-of select ="'5'"/>
				<!--<xsl:value-of select ="'10'"/>-->
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='across'">
				<xsl:value-of select ="'10'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='from-bottom-left'">
				<xsl:value-of select ="'12'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='left-to-bottom'">
				<xsl:value-of select ="'12'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='in'">
				<xsl:value-of select ="'16'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='vertical-in'">
				<xsl:value-of select ="'21'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='horizontal-in'">
				<xsl:value-of select ="'26'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='out'">
				<xsl:value-of select ="'32'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='vertical-out'">
				<xsl:value-of select ="'37'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='horizontal-out'">
				<xsl:value-of select ="'42'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='in-from-screen-center'">
				<xsl:value-of select ="'528'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='in-slightly'">
				<xsl:value-of select ="'272'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='out-from-screen-center'">
				<xsl:value-of select ="'36'"/>
			</xsl:when>
			<xsl:when test ="./@presentation:preset-sub-type ='out-slightly'">
				<xsl:value-of select ="'288'"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select ="'0'"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name ="smilFilter">
		<xsl:choose >
			<xsl:when test ="@smil:type = 'blindsWipe' and @smil:subtype ='Horizontal' ">
				<xsl:value-of  select ="'blinds(horizontal)'"/>
			</xsl:when>
			<xsl:when test ="@smil:type = 'blindsWipe' and @smil:subtype ='Vertical' ">
				<xsl:value-of  select ="'blinds(vertical)'"/>
			</xsl:when>
			<xsl:when test ="@smil:type = 'irisWipe' and @smil:subtype ='rectangle' and @smil:direction='reverse'">
				<xsl:value-of  select ="'box(in)'"/>
			</xsl:when>
			<xsl:when test ="@smil:type = 'irisWipe' and @smil:subtype ='rectangle' ">
				<xsl:value-of  select ="'box(out)'"/>
			</xsl:when>
			<xsl:when test ="@smil:type = 'checkerBoardWipe' and @smil:subtype ='across' ">
				<xsl:value-of  select ="'checkerboard(across)'"/>
			</xsl:when>
			<xsl:when test ="@smil:type = 'checkerBoardWipe' and @smil:subtype ='down' ">
				<xsl:value-of  select ="'checkerboard(down)'"/>
			</xsl:when>
			<xsl:when test ="@smil:type = 'ellipseWipe' and @smil:subtype ='horizontal' and @smil:direction='reverse' ">
				<xsl:value-of  select ="'circle(in)'"/>
			</xsl:when>
			<xsl:when test ="@smil:type = 'ellipseWipe' and @smil:subtype ='horizontal' ">
				<xsl:value-of  select ="'circle(out)'"/>
			</xsl:when>
			<xsl:when test ="@smil:type = 'irisWipe' and @smil:subtype ='diamond' and @smil:direction='reverse'">
				<xsl:value-of  select ="'diamond(in)'"/>
			</xsl:when>
			<xsl:when test ="@smil:type = 'irisWipe' and @smil:subtype ='diamond' ">
				<xsl:value-of  select ="'diamond(out)'"/>
			</xsl:when>
			<xsl:when test ="@smil:type = 'slideWipe' and @smil:subtype ='fromBottom' ">
				<xsl:value-of  select ="'slide(fromBottom)'"/>
			</xsl:when>
			<xsl:when test ="@smil:type = 'slideWipe' and @smil:subtype ='fromLeft' ">
				<xsl:value-of  select ="'slide(fromLeft)'"/>
			</xsl:when>
			<xsl:when test ="@smil:type = 'fromRight' and @smil:subtype ='slideWipe' ">
				<xsl:value-of  select ="'slide(fromRight)'"/>
			</xsl:when>
			<xsl:when test ="@smil:type = 'slideWipe' and @smil:subtype ='fromTop' ">
				<xsl:value-of  select ="'slide(fromTop)'"/>
			</xsl:when>
			<xsl:when test ="@smil:type = 'fourBoxWipe' and @smil:subtype ='cornersIn' ">
				<xsl:value-of  select ="'plus(in)'"/>
			</xsl:when>
			<xsl:when test ="@smil:type = 'fourBoxWipe' and @smil:subtype ='cornersIn' and @smil:direction='reverse' ">
				<xsl:value-of  select ="'plus(out)'"/>
			</xsl:when>
			<xsl:when test ="@smil:type = 'randomBarWipe' and @smil:subtype ='horizontal' ">
				<xsl:value-of  select ="'randombar(horizontal)'"/>
			</xsl:when>
			<xsl:when test ="@smil:type = 'randomBarWipe' and @smil:subtype ='vertical' ">
				<xsl:value-of  select ="'randombar(vertical)'"/>
			</xsl:when>
			<!--Start of RefNo-1-->
			<xsl:when test ="@smil:type = 'fanWipe' and @smil:subtype ='centerTop' ">
				<xsl:value-of  select ="'wedge'"/>
			</xsl:when>

			<xsl:when test ="@smil:type = 'pinWheelWipe'">
				<xsl:choose>
					<xsl:when test="@smil:subtype ='oneBlade'">
						<xsl:value-of  select ="'wheel(1)'"/>
					</xsl:when>
					<xsl:when test="@smil:subtype ='twoBladeVertical'">
						<xsl:value-of  select ="'wheel(2)'"/>
					</xsl:when>
					<xsl:when test="@smil:subtype ='threeBlade'">
						<xsl:value-of  select ="'wheel(3)'"/>
					</xsl:when>
					<xsl:when test="@smil:subtype ='fourBlade'">
						<xsl:value-of  select ="'wheel(4)'"/>
					</xsl:when>
					<xsl:when test="@smil:subtype ='eightBlade'">
						<xsl:value-of  select ="'wheel(8)'"/>
					</xsl:when>
				</xsl:choose>
			</xsl:when>

			<xsl:when test ="@smil:type = 'barWipe'">
				<xsl:choose>
					<xsl:when test="@smil:subtype ='topToBottom' and @smil:direction='reverse'">
						<xsl:value-of  select ="'wipe(down)'"/>
					</xsl:when>
					<xsl:when test="@smil:subtype ='topToBottom'">
						<xsl:value-of  select ="'wipe(up)'"/>
					</xsl:when>
					<xsl:when test="@smil:subtype ='leftToRight' and @smil:direction='reverse'">
						<xsl:value-of  select ="'wipe(right)'"/>
					</xsl:when>
					<xsl:when test="@smil:subtype ='leftToRight' ">
						<xsl:value-of  select ="'wipe(left)'"/>
					</xsl:when>
				</xsl:choose>
			</xsl:when>

			<xsl:when test ="@smil:type = 'fade' and @smil:subtype ='crossfade' ">
				<xsl:value-of  select ="'fade'"/>
			</xsl:when>

			<xsl:when test ="@smil:type = 'barnDoorWipe'">
				<xsl:choose>
					<xsl:when test="@smil:subtype ='horizontal' and @smil:direction='reverse'">
						<xsl:value-of  select ="'barn(inHorizontal)'"/>
					</xsl:when>
					<xsl:when test="@smil:subtype ='horizontal'">
						<xsl:value-of  select ="'barn(outHorizontal)'"/>
					</xsl:when>
					<xsl:when test="@smil:subtype ='vertical' and @smil:direction='reverse'">
						<xsl:value-of  select ="'barn(inVertical)'"/>
					</xsl:when>
					<xsl:when test="@smil:subtype ='vertical' ">
						<xsl:value-of  select ="'barn(outVertical)'"/>
					</xsl:when>
				</xsl:choose>
			</xsl:when>

			<xsl:when test ="@smil:type = 'waterfallWipe'">
				<xsl:choose>
					<xsl:when test="@smil:subtype ='horizontalRight' and @smil:direction='reverse'">
						<xsl:value-of  select ="'strips(upRight)'"/>
					</xsl:when>
					<xsl:when test="@smil:subtype ='horizontalRight'">
						<xsl:value-of  select ="'strips(downLeft)'"/>
					</xsl:when>
					<xsl:when test="@smil:subtype ='horizontalLeft' and @smil:direction='reverse'">
						<xsl:value-of  select ="'strips(upLeft)'"/>
					</xsl:when>
					<xsl:when test="@smil:subtype ='horizontalLeft' ">
						<xsl:value-of  select ="'strips(downRight)'"/>
					</xsl:when>
				</xsl:choose>
			</xsl:when>

			<xsl:when test ="@smil:type = 'dissolve'">
				<xsl:value-of  select ="'dissolve'"/>
			</xsl:when>

			<!--End of RefNo-1-->
		</xsl:choose>

	</xsl:template>
	<xsl:template name="stringReplace">
		<xsl:param name="text"/>
		<xsl:param name="replace"/>
		<xsl:param name="with"/>
		<xsl:choose>
			<xsl:when test="contains($text,$replace)">
				<xsl:value-of select="substring-before($text,$replace)"/>
				<xsl:value-of select="$with"/>
				<xsl:call-template name="stringReplace">
					<xsl:with-param name="text"	select="substring-after($text,$replace)"/>
					<xsl:with-param name="replace" select="$replace"/>
					<xsl:with-param name="with" select="$with"/>
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$text"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name ="mapCoordinates">
		<xsl:param name ="strVal"/>
		<xsl:variable name ="strVal1">
			<xsl:value-of select ="$strVal"/>
		</xsl:variable>
		<xsl:variable name ="strValppt_x">
			<xsl:call-template name ="stringReplace">
				<xsl:with-param name ="text" select ="$strVal1"/>
				<xsl:with-param name ="replace" select ="'x'" />
				<xsl:with-param name ="with" select ="'#ppt_x'"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name ="strValppt_y">
			<xsl:call-template name ="stringReplace">
				<xsl:with-param name ="text" select ="$strValppt_x"/>
				<xsl:with-param name ="replace" select ="'y'" />
				<xsl:with-param name ="with" select ="'#ppt_y'"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name ="strValppt_h">
			<xsl:call-template name ="stringReplace">
				<xsl:with-param name ="text" select ="$strValppt_y"/>
				<xsl:with-param name ="replace" select ="'height'" />
				<xsl:with-param name ="with" select ="'#ppt_h'"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name ="strValppt_w">
			<xsl:call-template name ="stringReplace">
				<xsl:with-param name ="text" select ="$strValppt_h"/>
				<xsl:with-param name ="replace" select ="'width'" />
				<xsl:with-param name ="with" select ="'#ppt_w'"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:value-of select ="$strValppt_w"/>
	</xsl:template>
	<xsl:template name ="interateType" >
		<xsl:param name ="itType" select ="@anim:iterate-type"/>
		<xsl:choose >
			<xsl:when test ="$itType ='by-letter'">
				<xsl:value-of select ="'lt'"/>
			</xsl:when>
			<xsl:otherwise >
				<xsl:value-of select ="'lt'"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<xsl:template name ="getnvPrIdval">
		<xsl:param name ="spId"/>
		<xsl:variable name ="varSpid">
			<xsl:for-each select ="./parent::node()/parent::node()/parent::node()/parent::node()/parent::node()">
        
				<xsl:for-each select ="node()">
          <xsl:if test ="name()='draw:rect' or name()='draw:ellipse'
                  or name()='draw:custom-shape' 
                  or name()='draw:g' or name()='draw:frame'">
            <xsl:choose>
              <xsl:when test="name()='draw:frame'">
                <xsl:choose>
                  <xsl:when test="./draw:image">
                      <xsl:if test ="contains(./draw:image/@xlink:href,'.png') or contains(./draw:image/@xlink:href,'.emf') 
                      or contains(./draw:image/@xlink:href,'.wmf') or contains(./draw:image/@xlink:href,'.jfif') or contains(./draw:image/@xlink:href,'.jpe') 
            or contains(./draw:image/@xlink:href,'.bmp') or contains(./draw:image/@xlink:href,'.dib')
            or contains(./draw:image/@xlink:href,'.rle')
            or contains(./draw:image/@xlink:href,'.bmz') or contains(./draw:image/@xlink:href,'.gfa') 
            or contains(./draw:image/@xlink:href,'.emz') or contains(./draw:image/@xlink:href,'.wmz') 
            or contains(./draw:image/@xlink:href,'.pcz')
            or contains(./draw:image/@xlink:href,'.tif') or contains(./draw:image/@xlink:href,'.tiff') 
            or contains(./draw:image/@xlink:href,'.cdr') or contains(./draw:image/@xlink:href,'.cgm')
            or contains(./draw:image/@xlink:href,'.eps') 
            or contains(./draw:image/@xlink:href,'.pct') or contains(./draw:image/@xlink:href,'.pict') 
            or contains(./draw:image/@xlink:href,'.wpg') 
            or contains(./draw:image/@xlink:href,'.jpeg') or contains(./draw:image/@xlink:href,'.gif') 
            or contains(./draw:image/@xlink:href,'.png') or contains(./draw:image/@xlink:href,'.jpg')">
                        <xsl:if test="not(./draw:image/@xlink:href[contains(.,'../')])">
                          <xsl:call-template name="tmpgetNvPrID">
                            <xsl:with-param name="spId" select="$spId"/>
                          </xsl:call-template>
                        </xsl:if>
                      </xsl:if>
                    </xsl:when>
                    <xsl:when test="@presentation:class[contains(.,'title')]
                                  or @presentation:class[contains(.,'subtitle')]
                                  or @presentation:class[contains(.,'outline')]">
                      <xsl:call-template name="tmpgetNvPrID">
                        <xsl:with-param name="spId" select="$spId"/>
                      </xsl:call-template>

                    </xsl:when>
                    <xsl:when test ="(draw:text-box) and not(@presentation:class)">
                      <xsl:call-template name="tmpgetNvPrID">
                        <xsl:with-param name="spId" select="$spId"/>
                      </xsl:call-template>
                    </xsl:when>
                    <xsl:when test="./draw:plugin">
                    </xsl:when>
                  </xsl:choose>
                </xsl:when>
               <xsl:when test="name()='draw:custom-shape'">
                  <xsl:variable name ="shapeName">
                    <xsl:call-template name="tmpgetCustShapeType"/>
                  </xsl:variable>
                  <xsl:if test="$shapeName != ''">
                    <xsl:call-template name="tmpgetNvPrID">
                      <xsl:with-param name="spId" select="$spId"/>
                    </xsl:call-template>
                </xsl:if>
              </xsl:when>
				<xsl:when test="name()='draw:g'">
					<xsl:call-template name="tmpgetNvPrID">
						<xsl:with-param name="spId" select="$spId"/>
					</xsl:call-template>
				</xsl:when>
              <xsl:when test="name()='draw:rect' or name()='draw:ellipse'
                          or name()='draw:line' or name()='draw:connector'">
                <xsl:call-template name="tmpgetNvPrID">
                  <xsl:with-param name="spId" select="$spId"/>
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>

      </xsl:for-each >
    </xsl:variable>
    <xsl:value-of select ="$varSpid"/>
  </xsl:template>
  <xsl:template name ="getParaIdval">
    <xsl:param name ="spId"/>
    <xsl:variable name ="varSpid">
      <xsl:for-each select ="./parent::node()/parent::node()/parent::node()/parent::node()/parent::node()">
        <xsl:for-each select ="node()">
          <xsl:variable name ="nvPrId">
            <xsl:value-of select ="position()"/>
          </xsl:variable>
          <xsl:variable name ="paraId">
            <xsl:for-each select =".//text:p">
              <xsl:if test ="$spId =@text:id">
                <xsl:value-of select ="position()"/>
              </xsl:if>
            </xsl:for-each>
					</xsl:variable>
					<xsl:if test ="$paraId!=''">
						<xsl:value-of select ="$paraId -1"/>
					</xsl:if>
				</xsl:for-each>
			</xsl:for-each>
		</xsl:variable>
		<xsl:value-of select ="$varSpid"/>
	</xsl:template>
  <xsl:template name="tmpgetNvPrID">
    <xsl:param name ="spId"/>
    <xsl:variable name ="nvPrId">
      <xsl:value-of select ="position()"/>
    </xsl:variable>
    <xsl:variable name ="drawId">
      <xsl:if test ="$spId =@draw:id">
        <xsl:value-of select ="position()"/>
      </xsl:if >
    </xsl:variable >
    <xsl:variable name ="paraId">
      <xsl:for-each select =".//text:p">
        <xsl:if test ="$spId =@text:id">
          <xsl:value-of select ="position()"/>
        </xsl:if>
      </xsl:for-each>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test ="$paraId!=''">
        <xsl:value-of select ="$nvPrId +1 "/>
      </xsl:when>
      <xsl:when test ="$paraId='' and $drawId !=''">
        <xsl:value-of select ="$nvPrId +1 "/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet >

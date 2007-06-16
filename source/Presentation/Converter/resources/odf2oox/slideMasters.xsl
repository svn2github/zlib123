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
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
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
	
	<!--<xsl:import href ="slides.xsl"/>-->
	<!--[@style:name=$titleName]-->
	<xsl:template name ="slideMasters">
		<xsl:param name="slideMasterName"/>
		<!--<xsl:param name ="pageNo"/>-->
		<!--<xsl:for-each select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name=$slideMasterName]">-->
		<!--match ="/office:document-styles/office:master-styles/style:master-page" mode="slideMaster"-->
		<!--<xsl:template name ="slideMaster1">-->
		<!--<xsl:for-each select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page">-->
		<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
					 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
					 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
			<p:cSld>
				<p:bg>
					<p:bgPr>
						<a:solidFill>
							<a:srgbClr>
								<xsl:variable name="dpName">
									<xsl:value-of select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name=$slideMasterName]/@draw:style-name"/>
								</xsl:variable>
								<xsl:attribute name="val">
									<xsl:choose>
										<xsl:when test="document('styles.xml')//style:style[@style:name= $dpName]/style:drawing-page-properties/@draw:fill-color">
											<xsl:for-each select="document('styles.xml')//style:style[@style:name= $dpName]/style:drawing-page-properties">
												<xsl:value-of select="substring-after(@draw:fill-color,'#')" />
											</xsl:for-each>
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
					</p:bgPr>
				</p:bg>
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
					<xsl:for-each select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[@style:name=$slideMasterName]/draw:frame">
						<xsl:if test="@presentation:class='title' or @presentation:class='outline' or @presentation:class='footer' or @presentation:class='date-time' or @presentation:class='page-number'">
							<p:sp>
								<p:nvSpPr>
									<p:cNvPr>
										<xsl:attribute name="id">
											<xsl:value-of select="position()+1"/>
										</xsl:attribute>
										<xsl:attribute name="name">
											<xsl:choose>
												<xsl:when test="@presentation:class='title'">
													<xsl:value-of select="'Title Placeholder 1'"/>
												</xsl:when>
												<xsl:when test="@presentation:class='outline'">
													<xsl:value-of select="'Text Placeholder 2'"/>
												</xsl:when>
												<xsl:when test="@presentation:class='date-time'">
													<xsl:value-of select="'Date Placeholder 3'"/>
												</xsl:when>
												<xsl:when test="@presentation:class='footer'">
													<xsl:value-of select="'Footer Placeholder 4'"/>
												</xsl:when>
												<xsl:when test="@presentation:class='page-number'">
													<xsl:value-of select="'Slide Number Placeholder 5'"/>
												</xsl:when>
											</xsl:choose>
										</xsl:attribute>

									</p:cNvPr>
									<p:cNvSpPr>
										<a:spLocks noGrp="1"/>
									</p:cNvSpPr>
									<p:nvPr>
										<p:ph>
											<xsl:attribute name="type">
												<xsl:choose>
													<xsl:when test ="@presentation:class ='title'">
														<xsl:value-of select="'title'"/>
													</xsl:when>
													<xsl:when test="@presentation:class ='outline'">
														<xsl:value-of select="'body'"/>
													</xsl:when>
													<xsl:when test ="@presentation:class ='date-time'">
														<xsl:value-of select="'dt'"/>
													</xsl:when>
													<xsl:when test="@presentation:class ='footer'">
														<xsl:value-of select="'ftr'"/>
													</xsl:when>
													<xsl:when test ="@presentation:class ='page-number'">
														<xsl:value-of select="'sldNum'"/>
													</xsl:when>
												</xsl:choose>
											</xsl:attribute>
											<xsl:if test="not(@presentation:class ='title')">
												<xsl:attribute name="idx">
													<xsl:choose>
														<xsl:when test="@presentation:class ='outline'">
															<xsl:value-of select="'1'"/>
														</xsl:when>
														<xsl:when test ="@presentation:class ='date-time'">
															<xsl:value-of select="'2'"/>
														</xsl:when>
														<xsl:when test="@presentation:class ='footer'">
															<xsl:value-of select="'3'"/>
														</xsl:when>
														<xsl:when test ="@presentation:class ='page-number'">
															<xsl:value-of select="'4'"/>
														</xsl:when>
													</xsl:choose>
												</xsl:attribute>
												<xsl:if test="not(@presentation:class ='outline')">
												<xsl:attribute name="sz">
												<xsl:choose>
														<xsl:when test ="@presentation:class ='date-time'">
															<xsl:value-of select="'half'"/>
														</xsl:when>
														<xsl:when test="@presentation:class ='footer'">
															<xsl:value-of select="'quarter'"/>
														</xsl:when>
														<xsl:when test ="@presentation:class ='page-number'">
															<xsl:value-of select="'quarter'"/>
														</xsl:when>
													</xsl:choose>
												</xsl:attribute>
												</xsl:if>
											</xsl:if>
										</p:ph>
									</p:nvPr>
								</p:nvSpPr>
								<p:spPr>
									<a:xfrm>
										<a:off>
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
									<a:prstGeom prst="rect">
										<a:avLst/>
									</a:prstGeom>
								</p:spPr>
								<p:txBody>
							<!--<a:normAutofit/>-->
									<xsl:call-template name ="TextAlignment" >
										<xsl:with-param name ="prId" select ="@presentation:style-name"/>
									</xsl:call-template >
									<xsl:if test="@presentation:class ='title'">
										<xsl:call-template name="Titlebody"/>
									</xsl:if>
									<xsl:if test="@presentation:class ='outline'">
										<xsl:call-template name="Outlinebody"/>
									</xsl:if>
									<xsl:if test="@presentation:class ='date-time'">
										<xsl:call-template name="Datetimebody"/>
									</xsl:if>
										<xsl:if test="@presentation:class ='footer'">
										<xsl:call-template name="Footerbody"/>
									</xsl:if>
									<xsl:if test="@presentation:class ='page-number'">
										<xsl:call-template name="Pagenumberbody"/>
									</xsl:if>
								</p:txBody>
							</p:sp>
						</xsl:if>
					</xsl:for-each>

				</p:spTree>
			</p:cSld>
			<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
			<p:sldLayoutIdLst>
				<p:sldLayoutId id="2147483649" r:id="rId1"/>
				<p:sldLayoutId id="2147483650" r:id="rId2"/>
				<p:sldLayoutId id="2147483651" r:id="rId3"/>
				<p:sldLayoutId id="2147483652" r:id="rId4"/>
				<p:sldLayoutId id="2147483653" r:id="rId5"/>
				<p:sldLayoutId id="2147483654" r:id="rId6"/>
				<p:sldLayoutId id="2147483655" r:id="rId7"/>
				<p:sldLayoutId id="2147483656" r:id="rId8"/>
				<p:sldLayoutId id="2147483657" r:id="rId9"/>
				<p:sldLayoutId id="2147483658" r:id="rId10"/>
				<p:sldLayoutId id="2147483659" r:id="rId11"/>
			</p:sldLayoutIdLst>
			
			<p:txStyles>
				<p:titleStyle>
					<a:lvl1pPr>
						<xsl:variable name="titleName">
							<xsl:value-of select="concat($slideMasterName,'-title')"/>
						</xsl:variable>
						<xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$titleName]">
							<xsl:attribute name="algn">
								<!--<xsl:value-of select="./style:paragraph-properties/@fo:text-align"/>-->
								<!--<xsl:value-of select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$titleName]/style:paragraph-properties/@fo:text-align"/>-->
								<xsl:choose >
									<xsl:when test ="./style:paragraph-properties/@fo:text-align='center'">
										<xsl:value-of select ="'ctr'"/>
									</xsl:when>
									<xsl:when test ="./style:paragraph-properties/@fo:text-align='end'">
										<xsl:value-of select ="'r'"/>
									</xsl:when>
									<xsl:otherwise >
										<xsl:value-of select ="'l'"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
							<xsl:attribute name="defTabSz">
								<xsl:value-of select="'914400'"/>
							</xsl:attribute>
							<xsl:attribute name="rtl">
								<xsl:value-of select="'0'"/>
							</xsl:attribute>
							<xsl:attribute name="eaLnBrk">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
							<xsl:attribute name="latinLnBrk">
								<xsl:value-of select="'0'"/>
							</xsl:attribute>
							<xsl:attribute name="hangingPunct">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
							<a:spcBef>
								<a:spcPct>
									<xsl:attribute name="val">
										<xsl:value-of select="'0'"/>
									</xsl:attribute>
								</a:spcPct>
							</a:spcBef>
							<a:buNone/>
							<a:defRPr>
								<xsl:attribute name="sz">
									<!--<xsl:value-of select ="@fo:font-size"/>-->
									<xsl:call-template name ="convertToPoints">
										<xsl:with-param name ="unit" select ="'pt'"/>
										<xsl:with-param name ="length" select ="./style:text-properties/@fo:font-size"/>
									</xsl:call-template>
								</xsl:attribute>
								<!--Font Bold attribute--> 
								<xsl:if test="./style:text-properties/@fo:font-weight='bold'">
									<xsl:attribute name ="b">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
								</xsl:if >
								 <!--Font Italic attribute-->
								<xsl:if test="./style:text-properties/@fo:font-style='italic'">
									<xsl:attribute name ="i">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
								</xsl:if >
								 <!--Font Underline-->
								<!--<xsl:if test="not(./style:text-properties/style:text-underline-style='solid')">
									<xsl:attribute name="u">
										<xsl:call-template name="Underline"/>
									</xsl:attribute>
								</xsl:if>-->
								
								<xsl:attribute name="kern">
									<xsl:value-of select="'0'"/>
								</xsl:attribute>
								<a:solidFill>
									<a:srgbClr>
										<xsl:attribute name="val">
											<xsl:value-of select="substring-after(./style:text-properties/@fo:color,'#')" />
											<!--<xsl:value-of select="//style:style[@style:name=$titleName]/child::node()[3]/@fo:color"/>-->
											<xsl:if test="not(./style:text-properties/@fo:color)">
												<xsl:value-of select="'000000'" />
											</xsl:if>
										</xsl:attribute>
									</a:srgbClr>
								</a:solidFill>
								<a:latin>
									<xsl:attribute name ="typeface" >
										<xsl:value-of select="translate(./style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
										<!--<xsl:value-of select="./style:text-properties/@fo:font-family"/>-->
									</xsl:attribute>
								</a:latin>
								<a:ea typeface="+mj-ea"/>
								<a:cs typeface="+mj-cs"/>
							</a:defRPr>
						</xsl:for-each>
					</a:lvl1pPr>
				</p:titleStyle>
				<p:bodyStyle>
					<a:lvl1pPr>
						<xsl:variable name="outlineName">
							 <xsl:value-of select="concat($slideMasterName,'-outline1')"/>
						</xsl:variable>
						<xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName]">
							<xsl:attribute name="marL">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name ="unit" select ="'cm'"/>
									<xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:attribute name="indent">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name ="unit" select ="'cm'"/>
									<xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:text-indent"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:attribute name="algn">
								
								<!--<xsl:value-of select="./@fo:text-align"/>-->
								<xsl:choose >
									<xsl:when test ="./style:paragraph-properties/@fo:text-align='center'">
										<xsl:value-of select ="'ctr'"/>
									</xsl:when>
									<xsl:when test ="./style:paragraph-properties/@fo:text-align='end'">
										<xsl:value-of select ="'r'"/>
									</xsl:when>
									<xsl:otherwise >
										<xsl:value-of select ="'l'"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
							<xsl:attribute name="defTabSz">
								<xsl:value-of select="'914400'"/>
							</xsl:attribute>
							<xsl:attribute name="rtl">
								<xsl:value-of select="'0'"/>
							</xsl:attribute>
							<xsl:attribute name="eaLnBrk">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
							<xsl:attribute name="latinLnBrk">
								<xsl:value-of select="'0'"/>
							</xsl:attribute>
							<xsl:attribute name="hangingPunct">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
							<a:spcBef>
								<a:spcPct>
									<xsl:attribute name="val">
										<xsl:value-of select="'0'"/>
										<!--<xsl:call-template name ="convertToPoints">
											<xsl:with-param name ="unit" select ="'cm'"/>
											<xsl:with-param name ="length" select ="(./style:graphic-properties/text:list-style/text:list-level-style-bullet)[1]/child::node()[1]/@text:space-before"/>
										</xsl:call-template>-->
										<!--<xsl:value-of select="(/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[1]/child::node()[1]/@text:space-before"/>-->
									</xsl:attribute>
								</a:spcPct>
							</a:spcBef>
							<a:buFont>
								<xsl:attribute name="typeface">
									<xsl:value-of select="'StarSymbol'"/>
									<!--<xsl:value-of select="(/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[1]/child::node()[2]/@fo:font-family"/>-->
								</xsl:attribute>
								<xsl:attribute name="pitchFamily">
									<xsl:value-of select="'0'"/>
								</xsl:attribute>
								<xsl:attribute name="charset">
									<xsl:value-of select="'1'"/>
								</xsl:attribute>
							</a:buFont>
							<a:buChar>
								<xsl:attribute name="char">
									<!--<xsl:value-of select="'•'"/>-->
									<xsl:value-of select="(./style:graphic-properties/text:list-style/text:list-level-style-bullet/@text:bullet-char)[1]"/>
								</xsl:attribute>
							</a:buChar>
							<a:defRPr>
								<xsl:attribute name="sz">
									<xsl:call-template name ="convertToPoints">
										<xsl:with-param name ="unit" select ="'pt'"/>
										<xsl:with-param name ="length" select ="./style:text-properties/@fo:font-size"/>
									</xsl:call-template>
								</xsl:attribute>
								<!--Font Bold attribute -->
								<xsl:if test="./style:text-properties/@fo:font-weight[contains(.,'bold')]">
									<xsl:attribute name ="b">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
								</xsl:if >
								<!-- Font Italic attribute-->
								<xsl:if test="./style:text-properties/@fo:font-style[contains(.,'italic')]">
									<xsl:attribute name ="i">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
								</xsl:if >
								<!-- Font Underline-->
								<!--<xsl:if test="not(./style:text-properties/style:text-underline-style[contains(.,'none')])">
									<xsl:attribute name="u">
										<xsl:call-template name="Underline"/>
									</xsl:attribute>
								</xsl:if>-->
								<xsl:attribute name="kern">
									<xsl:value-of select="'0'"/>
								</xsl:attribute>
								<!--<xsl:attribute name="strike">
									<xsl:choose >
										<xsl:when  test="style:text-properties/@style:text-line-through-type[contains(.,'solid')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'sngStrike'"/>
											</xsl:attribute >
										</xsl:when >
										<xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'dblStrike'"/>
											</xsl:attribute >
										</xsl:when >
										--><!-- style:text-line-through-style--><!--
										<xsl:when test="style:text-properties/@style:text-line-through-style[contains(.,'solid')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'sngStrike'"/>
											</xsl:attribute >
										</xsl:when>
									</xsl:choose>
								</xsl:attribute>-->
								<a:solidFill>
									<a:srgbClr>
										<xsl:attribute name="val">
											<xsl:value-of select="substring-after(./style:text-properties/@fo:color,'#')" />
											<xsl:if test="not(./style:text-properties/@fo:color)">
												<xsl:value-of select="'000000'" />
											</xsl:if>
										</xsl:attribute>
									</a:srgbClr>
								</a:solidFill>
								<a:effectLst/>
								<a:latin>
									<xsl:attribute name="typeface">
										<xsl:value-of select="translate(./style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
									</xsl:attribute>
								</a:latin>
								<a:ea typeface="+mn-ea"/>
								<a:cs typeface="+mn-cs"/>
							</a:defRPr>
						</xsl:for-each>
					</a:lvl1pPr>
					<a:lvl2pPr>
						<xsl:variable name="outlineName">
							<xsl:value-of select="concat($slideMasterName,'-outline1')"/>
						</xsl:variable>
						<xsl:variable name="outlineName2">
							<xsl:value-of select="concat($slideMasterName,'-outline2')"/>
							<!--<xsl:value-of select="concat(document('styles.xml')/office:document-styles/office:master-styles/style:master-page/@style:name,'-outline1')"/>-->
						</xsl:variable>
						<!--<xsl:variable name="outlineName2">
							<xsl:value-of select="concat($slideMasterName,'-outline2')"/>
						</xsl:variable>-->
						<xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName2]">
							<xsl:attribute name="marL">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name ="unit" select ="'cm'"/>
									<xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:attribute name="indent">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name ="unit" select ="'cm'"/>
									<xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:text-indent"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:attribute name="algn">
								<xsl:choose >
									<xsl:when test ="./style:paragraph-properties/@fo:text-align='center'">
										<xsl:value-of select ="'ctr'"/>
									</xsl:when>
									<xsl:when test ="./style:paragraph-properties/@fo:text-align='end'">
										<xsl:value-of select ="'r'"/>
									</xsl:when>
									<xsl:otherwise >
										<xsl:value-of select ="'l'"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
							<xsl:attribute name="defTabSz">
								<xsl:value-of select="'914400'"/>
							</xsl:attribute>
							<xsl:attribute name="rtl">
								<xsl:value-of select="'0'"/>
							</xsl:attribute>
							<xsl:attribute name="eaLnBrk">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
							<xsl:attribute name="latinLnBrk">
								<xsl:value-of select="'0'"/>
							</xsl:attribute>
							<xsl:attribute name="hangingPunct">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
							<a:spcBef>
								<a:spcPct>
									<xsl:attribute name="val">
										<xsl:value-of select="'0'"/>
										<!--<xsl:call-template name ="convertToPoints">
											<xsl:with-param name ="unit" select ="'cm'"/>
											<xsl:with-param name ="length" select ="(/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[2]/child::node()[1]/@text:space-before"/>
										</xsl:call-template>
										<xsl:value-of select="(/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[1]/child::node()[1]/@text:space-before"/>-->
									</xsl:attribute>
								</a:spcPct>
							</a:spcBef>
							<a:buFont>
								<xsl:attribute name="typeface">
									<xsl:value-of select="'StarSymbol'"/>
									<!--<xsl:value-of select="(document/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[1]/child::node()[2]/@fo:font-family"/>-->
								</xsl:attribute>
								<xsl:attribute name="pitchFamily">
									<xsl:value-of select="'0'"/>
								</xsl:attribute>
								<xsl:attribute name="charset">
									<xsl:value-of select="'1'"/>
								</xsl:attribute>
							</a:buFont>
							<a:buChar>
								<xsl:attribute name="char">
									<!--<xsl:value-of select="'–'"/>-->
									<xsl:value-of select="(document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet/@text:bullet-char)[2]"/>
								</xsl:attribute>
							</a:buChar>
							<a:defRPr>
								<xsl:attribute name="sz">
									<xsl:call-template name ="convertToPoints">
										<xsl:with-param name ="unit" select ="'pt'"/>
										<xsl:with-param name ="length" select ="./style:text-properties/@fo:font-size"/>
									</xsl:call-template>
								</xsl:attribute>
							<!--Font Bold attribute -->
								<xsl:if test="./style:text-properties/@fo:font-weight[contains(.,'bold')]">
									<xsl:attribute name ="b">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
								</xsl:if >
							<!-- Font Italic attribute-->
								<xsl:if test="./style:text-properties/@fo:font-style[contains(.,'italic')]">
									<xsl:attribute name ="i">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
								</xsl:if >
								<!-- Font Underline-->
								<!--<xsl:if test="not(./style:text-properties/style:text-underline-style[contains(.,'none')])">
									<xsl:attribute name="u">
										<xsl:call-template name="Underline"/>
									</xsl:attribute>
								</xsl:if>-->
								<xsl:attribute name="kern">
									<xsl:value-of select="'0'"/>
								</xsl:attribute>
								<!--<xsl:attribute name="strike">
									<xsl:choose >
										<xsl:when  test="style:text-properties/@style:text-line-through-type[contains(.,'solid')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'sngStrike'"/>
											</xsl:attribute >
										</xsl:when >
										<xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'dblStrike'"/>
											</xsl:attribute >
										</xsl:when >
										<xsl:when test="style:text-properties/@style:text-line-through-style[contains(.,'solid')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'sngStrike'"/>
											</xsl:attribute >
										</xsl:when>
									</xsl:choose>
								</xsl:attribute>-->
								<a:solidFill>
									<a:srgbClr>
										<xsl:attribute name="val">
											<xsl:value-of select="substring-after(./style:text-properties/@fo:color,'#')" />
											<xsl:if test="not(./style:text-properties/@fo:color)">
												<xsl:value-of select="'000000'" />
											</xsl:if>
										</xsl:attribute>
									</a:srgbClr>
								</a:solidFill>
								<a:effectLst/>
								<a:latin>
									<xsl:attribute name="typeface">
										<xsl:choose>
											<xsl:when test="translate(./style:text-properties/@fo:font-family, &quot;'&quot;,'')">
												<xsl:value-of select="translate(./style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
											</xsl:when>
										<xsl:otherwise>
												<xsl:value-of select="'Arial'"/>
												<!--<xsl:value-of select="translate(document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>-->
										</xsl:otherwise>
										</xsl:choose>
									</xsl:attribute>
								</a:latin>
								<a:ea typeface="+mn-ea"/>
								<a:cs typeface="+mn-cs"/>
							</a:defRPr>
						</xsl:for-each>
					</a:lvl2pPr>
					<a:lvl3pPr>
						<xsl:variable name="outlineName">
							<xsl:value-of select="concat($slideMasterName,'-outline1')"/>
						</xsl:variable>
						<xsl:variable name="outlineName3">
							<xsl:value-of select="concat($slideMasterName,'-outline3')"/>
							<!--<xsl:value-of select="concat(document('styles.xml')/office:document-styles/office:master-styles/style:master-page/@style:name,'-outline1')"/>-->
						</xsl:variable>
						<!--<xsl:variable name="outlineName2">
							<xsl:value-of select="concat($slideMasterName,'-outline2')"/>
						</xsl:variable>-->
						<xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName3]">
							<xsl:attribute name="marL">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name ="unit" select ="'cm'"/>
									<xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:attribute name="indent">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name ="unit" select ="'cm'"/>
									<xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:text-indent"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:attribute name="algn">
								<xsl:choose >
									<xsl:when test ="./style:paragraph-properties/@fo:text-align='center'">
										<xsl:value-of select ="'ctr'"/>
									</xsl:when>
									<xsl:when test ="./style:paragraph-properties/@fo:text-align='end'">
										<xsl:value-of select ="'r'"/>
									</xsl:when>
									<xsl:otherwise >
										<xsl:value-of select ="'l'"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
							<xsl:attribute name="defTabSz">
								<xsl:value-of select="'914400'"/>
							</xsl:attribute>
							<xsl:attribute name="rtl">
								<xsl:value-of select="'0'"/>
							</xsl:attribute>
							<xsl:attribute name="eaLnBrk">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
							<xsl:attribute name="latinLnBrk">
								<xsl:value-of select="'0'"/>
							</xsl:attribute>
							<xsl:attribute name="hangingPunct">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
							<a:spcBef>
								<a:spcPct>
									<xsl:attribute name="val">
										<xsl:value-of select="'0'"/>
										<!--<xsl:call-template name ="convertToPoints">
											<xsl:with-param name ="unit" select ="'cm'"/>
											<xsl:with-param name ="length" select ="(/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[2]/child::node()[1]/@text:space-before"/>
										</xsl:call-template>
										<xsl:value-of select="(/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[1]/child::node()[1]/@text:space-before"/>-->
									</xsl:attribute>
								</a:spcPct>
							</a:spcBef>
							<a:buFont>
								<xsl:attribute name="typeface">
									<xsl:value-of select="'StarSymbol'"/>
									<!--<xsl:value-of select="(document/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[1]/child::node()[2]/@fo:font-family"/>-->
								</xsl:attribute>
								<xsl:attribute name="pitchFamily">
									<xsl:value-of select="'0'"/>
								</xsl:attribute>
								<xsl:attribute name="charset">
									<xsl:value-of select="'1'"/>
								</xsl:attribute>
							</a:buFont>
							<a:buChar>
								<xsl:attribute name="char">
									<!--<xsl:value-of select="'●'"/>-->
									<xsl:value-of select="(document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet/@text:bullet-char)[3]"/>
								</xsl:attribute>
							</a:buChar>
							<a:defRPr>
								<xsl:attribute name="sz">
									<xsl:call-template name ="convertToPoints">
										<xsl:with-param name ="unit" select ="'pt'"/>
										<xsl:with-param name ="length" select ="./style:text-properties/@fo:font-size"/>
									</xsl:call-template>
								</xsl:attribute>
								<!--Font Bold attribute -->
								<xsl:if test="./style:text-properties/@fo:font-weight[contains(.,'bold')]">
									<xsl:attribute name ="b">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
								</xsl:if >
								<!-- Font Italic attribute-->
								<xsl:if test="./style:text-properties/@fo:font-style[contains(.,'italic')]">
									<xsl:attribute name ="i">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
								</xsl:if >
								<!-- Font Underline-->
								<!--<xsl:if test="not(./style:text-properties/style:text-underline-style[contains(.,'none')])">
									<xsl:attribute name="u">
										<xsl:call-template name="Underline"/>
									</xsl:attribute>
								</xsl:if>-->
								<xsl:attribute name="kern">
									<xsl:value-of select="'0'"/>
								</xsl:attribute>
								<!--<xsl:attribute name="strike">
									<xsl:choose >
										<xsl:when  test="style:text-properties/@style:text-line-through-type[contains(.,'solid')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'sngStrike'"/>
											</xsl:attribute >
										</xsl:when >
										<xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'dblStrike'"/>
											</xsl:attribute >
										</xsl:when >
										<xsl:when test="style:text-properties/@style:text-line-through-style[contains(.,'solid')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'sngStrike'"/>
											</xsl:attribute >
										</xsl:when>
									</xsl:choose>
								</xsl:attribute>-->
								<a:solidFill>
									<a:srgbClr>
										<xsl:attribute name="val">
											<xsl:value-of select="substring-after(./style:text-properties/@fo:color,'#')" />
											<xsl:if test="not(./style:text-properties/@fo:color)">
												<xsl:value-of select="'000000'" />
											</xsl:if>
										</xsl:attribute>
									</a:srgbClr>
								</a:solidFill>
								<a:effectLst/>
								<a:latin>
									<xsl:attribute name="typeface">
										<xsl:choose>
											<xsl:when test="translate(./style:text-properties/@fo:font-family, &quot;'&quot;,'')">
												<xsl:value-of select="translate(./style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
											</xsl:when>
											<xsl:otherwise>
												<xsl:value-of select="'Arial'"/>
												<!--<xsl:value-of select="translate(document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>-->
											</xsl:otherwise>
										</xsl:choose>
									</xsl:attribute>
								</a:latin>
								<a:ea typeface="+mn-ea"/>
								<a:cs typeface="+mn-cs"/>
							</a:defRPr>
						</xsl:for-each>

					</a:lvl3pPr>
					<a:lvl4pPr>
						<xsl:variable name="outlineName">
							<xsl:value-of select="concat($slideMasterName,'-outline1')"/>
						</xsl:variable>
						<xsl:variable name="outlineName4">
							<xsl:value-of select="concat($slideMasterName,'-outline4')"/>
							<!--<xsl:value-of select="concat(document('styles.xml')/office:document-styles/office:master-styles/style:master-page/@style:name,'-outline1')"/>-->
						</xsl:variable>
						<!--<xsl:variable name="outlineName2">
							<xsl:value-of select="concat($slideMasterName,'-outline2')"/>
						</xsl:variable>-->
						<xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName4]">
							<xsl:attribute name="marL">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name ="unit" select ="'cm'"/>
									<xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:attribute name="indent">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name ="unit" select ="'cm'"/>
									<xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:text-indent"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:attribute name="algn">
								<xsl:choose >
									<xsl:when test ="./style:paragraph-properties/@fo:text-align='center'">
										<xsl:value-of select ="'ctr'"/>
									</xsl:when>
									<xsl:when test ="./style:paragraph-properties/@fo:text-align='end'">
										<xsl:value-of select ="'r'"/>
									</xsl:when>
									<xsl:otherwise >
										<xsl:value-of select ="'l'"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
							<xsl:attribute name="defTabSz">
								<xsl:value-of select="'914400'"/>
							</xsl:attribute>
							<xsl:attribute name="rtl">
								<xsl:value-of select="'0'"/>
							</xsl:attribute>
							<xsl:attribute name="eaLnBrk">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
							<xsl:attribute name="latinLnBrk">
								<xsl:value-of select="'0'"/>
							</xsl:attribute>
							<xsl:attribute name="hangingPunct">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
							<a:spcBef>
								<a:spcPct>
									<xsl:attribute name="val">
										<xsl:value-of select="'0'"/>
										<!--<xsl:call-template name ="convertToPoints">
											<xsl:with-param name ="unit" select ="'cm'"/>
											<xsl:with-param name ="length" select ="(/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[2]/child::node()[1]/@text:space-before"/>
										</xsl:call-template>
										<xsl:value-of select="(/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[1]/child::node()[1]/@text:space-before"/>-->
									</xsl:attribute>
								</a:spcPct>
							</a:spcBef>
							<a:buFont>
								<xsl:attribute name="typeface">
									<xsl:value-of select="'StarSymbol'"/>
									<!--<xsl:value-of select="(document/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[1]/child::node()[2]/@fo:font-family"/>-->
								</xsl:attribute>
								<xsl:attribute name="pitchFamily">
									<xsl:value-of select="'0'"/>
								</xsl:attribute>
								<xsl:attribute name="charset">
									<xsl:value-of select="'1'"/>
								</xsl:attribute>
							</a:buFont>
							<a:buChar>
								<xsl:attribute name="char">
									<!--<xsl:value-of select="'–'"/>-->
									<xsl:value-of select="(document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet/@text:bullet-char)[4]"/>
								</xsl:attribute>
							</a:buChar>
							<a:defRPr>
								<xsl:attribute name="sz">
									<xsl:call-template name ="convertToPoints">
										<xsl:with-param name ="unit" select ="'pt'"/>
										<xsl:with-param name ="length" select ="./style:text-properties/@fo:font-size"/>
									</xsl:call-template>
								</xsl:attribute>
								<!--Font Bold attribute -->
								<xsl:if test="./style:text-properties/@fo:font-weight[contains(.,'bold')]">
									<xsl:attribute name ="b">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
								</xsl:if >
								<!-- Font Italic attribute-->
								<xsl:if test="./style:text-properties/@fo:font-style[contains(.,'italic')]">
									<xsl:attribute name ="i">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
								</xsl:if >
								<!-- Font Underline-->
								<!--<xsl:if test="not(./style:text-properties/style:text-underline-style[contains(.,'none')])">
									<xsl:attribute name="u">
										<xsl:call-template name="Underline"/>
									</xsl:attribute>
								</xsl:if>-->
								<xsl:attribute name="kern">
									<xsl:value-of select="'0'"/>
								</xsl:attribute>
								<!--<xsl:attribute name="strike">
									<xsl:choose >
										<xsl:when  test="style:text-properties/@style:text-line-through-type[contains(.,'solid')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'sngStrike'"/>
											</xsl:attribute >
										</xsl:when >
										<xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'dblStrike'"/>
											</xsl:attribute >
										</xsl:when >
										<xsl:when test="style:text-properties/@style:text-line-through-style[contains(.,'solid')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'sngStrike'"/>
											</xsl:attribute >
										</xsl:when>
									</xsl:choose>
								</xsl:attribute>-->
								<a:solidFill>
									<a:srgbClr>
										<xsl:attribute name="val">
											<xsl:value-of select="substring-after(./style:text-properties/@fo:color,'#')" />
											<xsl:if test="not(./style:text-properties/@fo:color)">
												<xsl:value-of select="'000000'" />
											</xsl:if>
										</xsl:attribute>
									</a:srgbClr>
								</a:solidFill>
								<a:effectLst/>
								<a:latin>
									<xsl:attribute name="typeface">
										<xsl:choose>
											<xsl:when test="translate(./style:text-properties/@fo:font-family, &quot;'&quot;,'')">
												<xsl:value-of select="translate(./style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
											</xsl:when>
											<xsl:otherwise>
												<xsl:value-of select="'Arial'"/>
												<!--<xsl:value-of select="translate(document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>-->
											</xsl:otherwise>
										</xsl:choose>
									</xsl:attribute>
								</a:latin>
								<a:ea typeface="+mn-ea"/>
								<a:cs typeface="+mn-cs"/>
							</a:defRPr>
						</xsl:for-each>
					</a:lvl4pPr>
					<a:lvl5pPr>
						<xsl:variable name="outlineName">
							<xsl:value-of select="concat($slideMasterName,'-outline1')"/>
						</xsl:variable>
						<xsl:variable name="outlineName5">
							<xsl:value-of select="concat($slideMasterName,'-outline5')"/>
							<!--<xsl:value-of select="concat(document('styles.xml')/office:document-styles/office:master-styles/style:master-page/@style:name,'-outline1')"/>-->
						</xsl:variable>
						<!--<xsl:variable name="outlineName2">
							<xsl:value-of select="concat($slideMasterName,'-outline2')"/>
						</xsl:variable>-->
						<xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName5]">
							<xsl:attribute name="marL">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name ="unit" select ="'cm'"/>
									<xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:attribute name="indent">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name ="unit" select ="'cm'"/>
									<xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:text-indent"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:attribute name="algn">
								<xsl:choose >
									<xsl:when test ="./style:paragraph-properties/@fo:text-align='center'">
										<xsl:value-of select ="'ctr'"/>
									</xsl:when>
									<xsl:when test ="./style:paragraph-properties/@fo:text-align='end'">
										<xsl:value-of select ="'r'"/>
									</xsl:when>
									<xsl:otherwise >
										<xsl:value-of select ="'l'"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
							<xsl:attribute name="defTabSz">
								<xsl:value-of select="'914400'"/>
							</xsl:attribute>
							<xsl:attribute name="rtl">
								<xsl:value-of select="'0'"/>
							</xsl:attribute>
							<xsl:attribute name="eaLnBrk">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
							<xsl:attribute name="latinLnBrk">
								<xsl:value-of select="'0'"/>
							</xsl:attribute>
							<xsl:attribute name="hangingPunct">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
							<a:spcBef>
								<a:spcPct>
									<xsl:attribute name="val">
										<xsl:value-of select="'0'"/>
										<!--<xsl:call-template name ="convertToPoints">
											<xsl:with-param name ="unit" select ="'cm'"/>
											<xsl:with-param name ="length" select ="(/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[2]/child::node()[1]/@text:space-before"/>
										</xsl:call-template>
										<xsl:value-of select="(/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[1]/child::node()[1]/@text:space-before"/>-->
									</xsl:attribute>
								</a:spcPct>
							</a:spcBef>
							<a:buFont>
								<xsl:attribute name="typeface">
									<xsl:value-of select="'StarSymbol'"/>
									<!--<xsl:value-of select="(document/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[1]/child::node()[2]/@fo:font-family"/>-->
								</xsl:attribute>
								<xsl:attribute name="pitchFamily">
									<xsl:value-of select="'0'"/>
								</xsl:attribute>
								<xsl:attribute name="charset">
									<xsl:value-of select="'1'"/>
								</xsl:attribute>
							</a:buFont>
							<a:buChar>
								<xsl:attribute name="char">
									<!--<xsl:value-of select="'●'"/>-->
									<xsl:value-of select="(document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet/@text:bullet-char)[5]"/>
								</xsl:attribute>
							</a:buChar>
							<a:defRPr>
								<xsl:attribute name="sz">
									<xsl:call-template name ="convertToPoints">
										<xsl:with-param name ="unit" select ="'pt'"/>
										<xsl:with-param name ="length" select ="./style:text-properties/@fo:font-size"/>
									</xsl:call-template>
								</xsl:attribute>
								<!--Font Bold attribute -->
								<xsl:if test="./style:text-properties/@fo:font-weight[contains(.,'bold')]">
									<xsl:attribute name ="b">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
								</xsl:if >
								<!-- Font Italic attribute-->
								<xsl:if test="./style:text-properties/@fo:font-style[contains(.,'italic')]">
									<xsl:attribute name ="i">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
								</xsl:if >
								<!-- Font Underline-->
								<!--<xsl:if test="not(./style:text-properties/style:text-underline-style[contains(.,'none')])">
									<xsl:attribute name="u">
										<xsl:call-template name="Underline"/>
									</xsl:attribute>
								</xsl:if>-->
								<xsl:attribute name="kern">
									<xsl:value-of select="'0'"/>
								</xsl:attribute>
								<!--<xsl:attribute name="strike">
									<xsl:choose >
										<xsl:when  test="style:text-properties/@style:text-line-through-type[contains(.,'solid')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'sngStrike'"/>
											</xsl:attribute >
										</xsl:when >
										<xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'dblStrike'"/>
											</xsl:attribute >
										</xsl:when >
										<xsl:when test="style:text-properties/@style:text-line-through-style[contains(.,'solid')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'sngStrike'"/>
											</xsl:attribute >
										</xsl:when>
									</xsl:choose>
								</xsl:attribute>-->
								<a:solidFill>
									<a:srgbClr>
										<xsl:attribute name="val">
											<xsl:value-of select="substring-after(./style:text-properties/@fo:color,'#')" />
											<xsl:if test="not(./style:text-properties/@fo:color)">
												<xsl:value-of select="'000000'" />
											</xsl:if>
										</xsl:attribute>
									</a:srgbClr>
								</a:solidFill>
								<a:effectLst/>
								<a:latin>
									<xsl:attribute name="typeface">
										<xsl:choose>
											<xsl:when test="translate(./style:text-properties/@fo:font-family, &quot;'&quot;,'')">
												<xsl:value-of select="translate(./style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
											</xsl:when>
											<xsl:otherwise>
												<xsl:value-of select="'Arial'"/>
												<!--<xsl:value-of select="translate(document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>-->
											</xsl:otherwise>
										</xsl:choose>
									</xsl:attribute>
								</a:latin>
								<a:ea typeface="+mn-ea"/>
								<a:cs typeface="+mn-cs"/>
							</a:defRPr>
						</xsl:for-each>
					</a:lvl5pPr>
					<a:lvl6pPr>
						<xsl:variable name="outlineName">
							<xsl:value-of select="concat($slideMasterName,'-outline1')"/>
						</xsl:variable>
						<xsl:variable name="outlineName6">
							<xsl:value-of select="concat($slideMasterName,'-outline6')"/>
							<!--<xsl:value-of select="concat(document('styles.xml')/office:document-styles/office:master-styles/style:master-page/@style:name,'-outline1')"/>-->
						</xsl:variable>
						<!--<xsl:variable name="outlineName2">
							<xsl:value-of select="concat($slideMasterName,'-outline2')"/>
						</xsl:variable>-->
						<xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName6]">
							<xsl:attribute name="marL">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name ="unit" select ="'cm'"/>
									<xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:attribute name="indent">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name ="unit" select ="'cm'"/>
									<xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:text-indent"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:attribute name="algn">
								<xsl:choose >
									<xsl:when test ="./style:paragraph-properties/@fo:text-align='center'">
										<xsl:value-of select ="'ctr'"/>
									</xsl:when>
									<xsl:when test ="./style:paragraph-properties/@fo:text-align='end'">
										<xsl:value-of select ="'r'"/>
									</xsl:when>
									<xsl:otherwise >
										<xsl:value-of select ="'l'"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
							<xsl:attribute name="defTabSz">
								<xsl:value-of select="'914400'"/>
							</xsl:attribute>
							<xsl:attribute name="rtl">
								<xsl:value-of select="'0'"/>
							</xsl:attribute>
							<xsl:attribute name="eaLnBrk">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
							<xsl:attribute name="latinLnBrk">
								<xsl:value-of select="'0'"/>
							</xsl:attribute>
							<xsl:attribute name="hangingPunct">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
							<a:spcBef>
								<a:spcPct>
									<xsl:attribute name="val">
										<xsl:value-of select="'0'"/>
										<!--<xsl:call-template name ="convertToPoints">
											<xsl:with-param name ="unit" select ="'cm'"/>
											<xsl:with-param name ="length" select ="(/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[2]/child::node()[1]/@text:space-before"/>
										</xsl:call-template>
										<xsl:value-of select="(/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[1]/child::node()[1]/@text:space-before"/>-->
									</xsl:attribute>
								</a:spcPct>
							</a:spcBef>
							<a:buFont>
								<xsl:attribute name="typeface">
									<xsl:value-of select="'StarSymbol'"/>
									<!--<xsl:value-of select="(document/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[1]/child::node()[2]/@fo:font-family"/>-->
								</xsl:attribute>
								<xsl:attribute name="pitchFamily">
									<xsl:value-of select="'0'"/>
								</xsl:attribute>
								<xsl:attribute name="charset">
									<xsl:value-of select="'1'"/>
								</xsl:attribute>
							</a:buFont>
							<a:buChar>
								<xsl:attribute name="char">
									<!--<xsl:value-of select="'●'"/>-->
									<xsl:value-of select="(document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet/@text:bullet-char)[6]"/>
								</xsl:attribute>
							</a:buChar>
							<a:defRPr>
								<xsl:attribute name="sz">
									<xsl:call-template name ="convertToPoints">
										<xsl:with-param name ="unit" select ="'pt'"/>
										<xsl:with-param name ="length" select ="./style:text-properties/@fo:font-size"/>
									</xsl:call-template>
								</xsl:attribute>
								<!--Font Bold attribute -->
								<xsl:if test="./style:text-properties/@fo:font-weight[contains(.,'bold')]">
									<xsl:attribute name ="b">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
								</xsl:if >
								<!-- Font Italic attribute-->
								<xsl:if test="./style:text-properties/@fo:font-style[contains(.,'italic')]">
									<xsl:attribute name ="i">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
								</xsl:if >
								<!-- Font Underline-->
								<!--<xsl:if test="not(./style:text-properties/style:text-underline-style[contains(.,'none')])">
									<xsl:attribute name="u">
										<xsl:call-template name="Underline"/>
									</xsl:attribute>
								</xsl:if>-->
								<xsl:attribute name="kern">
									<xsl:value-of select="'0'"/>
								</xsl:attribute>
								<!--<xsl:attribute name="strike">
									<xsl:choose >
										<xsl:when  test="style:text-properties/@style:text-line-through-type[contains(.,'solid')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'sngStrike'"/>
											</xsl:attribute >
										</xsl:when >
										<xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'dblStrike'"/>
											</xsl:attribute >
										</xsl:when >
										<xsl:when test="style:text-properties/@style:text-line-through-style[contains(.,'solid')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'sngStrike'"/>
											</xsl:attribute >
										</xsl:when>
									</xsl:choose>
								</xsl:attribute>-->
								<a:solidFill>
									<a:srgbClr>
										<xsl:attribute name="val">
											<xsl:value-of select="substring-after(./style:text-properties/@fo:color,'#')" />
											<xsl:if test="not(./style:text-properties/@fo:color)">
												<xsl:value-of select="'000000'" />
											</xsl:if>
										</xsl:attribute>
									</a:srgbClr>
								</a:solidFill>
								<a:effectLst/>
								<a:latin>
									<xsl:attribute name="typeface">
										<xsl:choose>
											<xsl:when test="translate(./style:text-properties/@fo:font-family, &quot;'&quot;,'')">
												<xsl:value-of select="translate(./style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
											</xsl:when>
											<xsl:otherwise>
												<xsl:value-of select="'Arial'"/>
												<!--<xsl:value-of select="translate(document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>-->
											</xsl:otherwise>
										</xsl:choose>
									</xsl:attribute>
								</a:latin>
								<a:ea typeface="+mn-ea"/>
								<a:cs typeface="+mn-cs"/>
							</a:defRPr>
						</xsl:for-each>
					</a:lvl6pPr>
					<a:lvl7pPr>
						<xsl:variable name="outlineName">
							<xsl:value-of select="concat($slideMasterName,'-outline1')"/>
						</xsl:variable>
						<xsl:variable name="outlineName7">
							<xsl:value-of select="concat($slideMasterName,'-outline7')"/>
							<!--<xsl:value-of select="concat(document('styles.xml')/office:document-styles/office:master-styles/style:master-page/@style:name,'-outline1')"/>-->
						</xsl:variable>
						<!--<xsl:variable name="outlineName2">
							<xsl:value-of select="concat($slideMasterName,'-outline2')"/>
						</xsl:variable>-->
						<xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName7]">
							<xsl:attribute name="marL">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name ="unit" select ="'cm'"/>
									<xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:attribute name="indent">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name ="unit" select ="'cm'"/>
									<xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:text-indent"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:attribute name="algn">
								<xsl:choose >
									<xsl:when test ="./style:paragraph-properties/@fo:text-align='center'">
										<xsl:value-of select ="'ctr'"/>
									</xsl:when>
									<xsl:when test ="./style:paragraph-properties/@fo:text-align='end'">
										<xsl:value-of select ="'r'"/>
									</xsl:when>
									<xsl:otherwise >
										<xsl:value-of select ="'l'"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
							<xsl:attribute name="defTabSz">
								<xsl:value-of select="'914400'"/>
							</xsl:attribute>
							<xsl:attribute name="rtl">
								<xsl:value-of select="'0'"/>
							</xsl:attribute>
							<xsl:attribute name="eaLnBrk">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
							<xsl:attribute name="latinLnBrk">
								<xsl:value-of select="'0'"/>
							</xsl:attribute>
							<xsl:attribute name="hangingPunct">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
							<a:spcBef>
								<a:spcPct>
									<xsl:attribute name="val">
										<xsl:value-of select="'0'"/>
										<!--<xsl:call-template name ="convertToPoints">
											<xsl:with-param name ="unit" select ="'cm'"/>
											<xsl:with-param name ="length" select ="(/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[2]/child::node()[1]/@text:space-before"/>
										</xsl:call-template>
										<xsl:value-of select="(/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[1]/child::node()[1]/@text:space-before"/>-->
									</xsl:attribute>
								</a:spcPct>
							</a:spcBef>
							<a:buFont>
								<xsl:attribute name="typeface">
									<xsl:value-of select="'StarSymbol'"/>
									<!--<xsl:value-of select="(document/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[1]/child::node()[2]/@fo:font-family"/>-->
								</xsl:attribute>
								<xsl:attribute name="pitchFamily">
									<xsl:value-of select="'0'"/>
								</xsl:attribute>
								<xsl:attribute name="charset">
									<xsl:value-of select="'1'"/>
								</xsl:attribute>
							</a:buFont>
							<a:buChar>
								<xsl:attribute name="char">
									<!--<xsl:value-of select="'●'"/>-->
									<xsl:value-of select="(document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet/@text:bullet-char)[7]"/>
								</xsl:attribute>
							</a:buChar>
							<a:defRPr>
								<xsl:attribute name="sz">
									<xsl:call-template name ="convertToPoints">
										<xsl:with-param name ="unit" select ="'pt'"/>
										<xsl:with-param name ="length" select ="./style:text-properties/@fo:font-size"/>
									</xsl:call-template>
								</xsl:attribute>
							<!--Font Bold attribute -->
								<xsl:if test="./style:text-properties/@fo:font-weight[contains(.,'bold')]">
									<xsl:attribute name ="b">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
								</xsl:if >
							<!-- Font Italic attribute-->
								<xsl:if test="./style:text-properties/@fo:font-style[contains(.,'italic')]">
									<xsl:attribute name ="i">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
								</xsl:if >
								<!-- Font Underline-->
								<!--<xsl:if test="not(./style:text-properties/style:text-underline-style[contains(.,'none')])">
									<xsl:attribute name="u">
										<xsl:call-template name="Underline"/>
									</xsl:attribute>
								</xsl:if>-->
								<xsl:attribute name="kern">
									<xsl:value-of select="'0'"/>
								</xsl:attribute>
								<!--<xsl:attribute name="strike">
									<xsl:choose >
										<xsl:when  test="style:text-properties/@style:text-line-through-type[contains(.,'solid')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'sngStrike'"/>
											</xsl:attribute >
										</xsl:when >
										<xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'dblStrike'"/>
											</xsl:attribute >
										</xsl:when >
										<xsl:when test="style:text-properties/@style:text-line-through-style[contains(.,'solid')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'sngStrike'"/>
											</xsl:attribute >
										</xsl:when>
									</xsl:choose>
								</xsl:attribute>-->
								<a:solidFill>
									<a:srgbClr>
										<xsl:attribute name="val">
											<xsl:value-of select="substring-after(./style:text-properties/@fo:color,'#')" />
											<xsl:if test="not(./style:text-properties/@fo:color)">
												<xsl:value-of select="'000000'" />
											</xsl:if>
										</xsl:attribute>
									</a:srgbClr>
								</a:solidFill>
								<a:effectLst/>
								<a:latin>
									<xsl:attribute name="typeface">
										<xsl:choose>
											<xsl:when test="translate(./style:text-properties/@fo:font-family, &quot;'&quot;,'')">
												<xsl:value-of select="translate(./style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
											</xsl:when>
										<xsl:otherwise>
												<xsl:value-of select="'Arial'"/>
												<!--<xsl:value-of select="translate(document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>-->
										</xsl:otherwise>
										</xsl:choose>
									</xsl:attribute>
								</a:latin>
								<a:ea typeface="+mn-ea"/>
								<a:cs typeface="+mn-cs"/>
							</a:defRPr>
						</xsl:for-each>

					</a:lvl7pPr>
					<a:lvl8pPr>
						<xsl:variable name="outlineName">
							<xsl:value-of select="concat($slideMasterName,'-outline1')"/>
						</xsl:variable>
						<xsl:variable name="outlineName8">
							<xsl:value-of select="concat($slideMasterName,'-outline8')"/>
							<!--<xsl:value-of select="concat(document('styles.xml')/office:document-styles/office:master-styles/style:master-page/@style:name,'-outline1')"/>-->
						</xsl:variable>
						<!--<xsl:variable name="outlineName2">
							<xsl:value-of select="concat($slideMasterName,'-outline2')"/>
						</xsl:variable>-->
						<xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName8]">
							<xsl:attribute name="marL">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name ="unit" select ="'cm'"/>
									<xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:attribute name="indent">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name ="unit" select ="'cm'"/>
									<xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:text-indent"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:attribute name="algn">
								<xsl:choose >
									<xsl:when test ="./style:paragraph-properties/@fo:text-align='center'">
										<xsl:value-of select ="'ctr'"/>
									</xsl:when>
									<xsl:when test ="./style:paragraph-properties/@fo:text-align='end'">
										<xsl:value-of select ="'r'"/>
									</xsl:when>
									<xsl:otherwise >
										<xsl:value-of select ="'l'"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
							<xsl:attribute name="defTabSz">
								<xsl:value-of select="'914400'"/>
							</xsl:attribute>
							<xsl:attribute name="rtl">
								<xsl:value-of select="'0'"/>
							</xsl:attribute>
							<xsl:attribute name="eaLnBrk">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
							<xsl:attribute name="latinLnBrk">
								<xsl:value-of select="'0'"/>
							</xsl:attribute>
							<xsl:attribute name="hangingPunct">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
							<a:spcBef>
								<a:spcPct>
									<xsl:attribute name="val">
										<xsl:value-of select="'0'"/>
										<!--<xsl:call-template name ="convertToPoints">
											<xsl:with-param name ="unit" select ="'cm'"/>
											<xsl:with-param name ="length" select ="(/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[2]/child::node()[1]/@text:space-before"/>
										</xsl:call-template>
										<xsl:value-of select="(/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[1]/child::node()[1]/@text:space-before"/>-->
									</xsl:attribute>
								</a:spcPct>
							</a:spcBef>
							<a:buFont>
								<xsl:attribute name="typeface">
									<xsl:value-of select="'StarSymbol'"/>
									<!--<xsl:value-of select="(document/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[1]/child::node()[2]/@fo:font-family"/>-->
								</xsl:attribute>
								<xsl:attribute name="pitchFamily">
									<xsl:value-of select="'0'"/>
								</xsl:attribute>
								<xsl:attribute name="charset">
									<xsl:value-of select="'1'"/>
								</xsl:attribute>
							</a:buFont>
							<a:buChar>
								<xsl:attribute name="char">
									<!--<xsl:value-of select="'●'"/>-->
									<xsl:value-of select="(document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet/@text:bullet-char)[8]"/>
								</xsl:attribute>
							</a:buChar>
							<a:defRPr>
								<xsl:attribute name="sz">
									<xsl:call-template name ="convertToPoints">
										<xsl:with-param name ="unit" select ="'pt'"/>
										<xsl:with-param name ="length" select ="./style:text-properties/@fo:font-size"/>
									</xsl:call-template>
								</xsl:attribute>
								<!--Font Bold attribute -->
								<xsl:if test="./style:text-properties/@fo:font-weight[contains(.,'bold')]">
									<xsl:attribute name ="b">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
								</xsl:if >
								<!-- Font Italic attribute-->
								<xsl:if test="./style:text-properties/@fo:font-style[contains(.,'italic')]">
									<xsl:attribute name ="i">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
								</xsl:if >
								<!-- Font Underline-->
								<!--<xsl:if test="not(./style:text-properties/style:text-underline-style[contains(.,'none')])">
									<xsl:attribute name="u">
										<xsl:call-template name="Underline"/>
									</xsl:attribute>
								</xsl:if>-->
								<xsl:attribute name="kern">
									<xsl:value-of select="'0'"/>
								</xsl:attribute>
								<!--<xsl:attribute name="strike">
									<xsl:choose >
										<xsl:when  test="style:text-properties/@style:text-line-through-type[contains(.,'solid')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'sngStrike'"/>
											</xsl:attribute >
										</xsl:when >
										<xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'dblStrike'"/>
											</xsl:attribute >
										</xsl:when >
										<xsl:when test="style:text-properties/@style:text-line-through-style[contains(.,'solid')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'sngStrike'"/>
											</xsl:attribute >
										</xsl:when>
									</xsl:choose>
								</xsl:attribute>-->
								<a:solidFill>
									<a:srgbClr>
										<xsl:attribute name="val">
											<xsl:value-of select="substring-after(./style:text-properties/@fo:color,'#')" />
											<xsl:if test="not(./style:text-properties/@fo:color)">
												<xsl:value-of select="'000000'" />
											</xsl:if>
										</xsl:attribute>
									</a:srgbClr>
								</a:solidFill>
								<a:effectLst/>
								<a:latin>
									<xsl:attribute name="typeface">
										<xsl:choose>
											<xsl:when test="translate(./style:text-properties/@fo:font-family, &quot;'&quot;,'')">
												<xsl:value-of select="translate(./style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
											</xsl:when>
											<xsl:otherwise>
												<xsl:value-of select="'Arial'"/>
												<!--<xsl:value-of select="translate(document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>-->
											</xsl:otherwise>
										</xsl:choose>
									</xsl:attribute>
								</a:latin>
								<a:ea typeface="+mn-ea"/>
								<a:cs typeface="+mn-cs"/>
							</a:defRPr>
						</xsl:for-each>
					</a:lvl8pPr>
					<a:lvl9pPr>
						<xsl:variable name="outlineName">
							<xsl:value-of select="concat($slideMasterName,'-outline1')"/>
						</xsl:variable>
						<xsl:variable name="outlineName9">
							<xsl:value-of select="concat($slideMasterName,'-outline9')"/>
							<!--<xsl:value-of select="concat(document('styles.xml')/office:document-styles/office:master-styles/style:master-page/@style:name,'-outline1')"/>-->
						</xsl:variable>
						<!--<xsl:variable name="outlineName2">
							<xsl:value-of select="concat($slideMasterName,'-outline2')"/>
						</xsl:variable>-->
						<xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName9]">
							<xsl:attribute name="marL">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name ="unit" select ="'cm'"/>
									<xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:margin-left"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:attribute name="indent">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name ="unit" select ="'cm'"/>
									<xsl:with-param name ="length" select ="./style:paragraph-properties/@fo:text-indent"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:attribute name="algn">
								<xsl:choose >
									<xsl:when test ="./style:paragraph-properties/@fo:text-align='center'">
										<xsl:value-of select ="'ctr'"/>
									</xsl:when>
									<xsl:when test ="./style:paragraph-properties/@fo:text-align='end'">
										<xsl:value-of select ="'r'"/>
									</xsl:when>
									<xsl:otherwise >
										<xsl:value-of select ="'l'"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
							<xsl:attribute name="defTabSz">
								<xsl:value-of select="'914400'"/>
							</xsl:attribute>
							<xsl:attribute name="rtl">
								<xsl:value-of select="'0'"/>
							</xsl:attribute>
							<xsl:attribute name="eaLnBrk">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
							<xsl:attribute name="latinLnBrk">
								<xsl:value-of select="'0'"/>
							</xsl:attribute>
							<xsl:attribute name="hangingPunct">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
							<a:spcBef>
								<a:spcPct>
									<xsl:attribute name="val">
										<xsl:value-of select="'0'"/>
										<!--<xsl:call-template name ="convertToPoints">
											<xsl:with-param name ="unit" select ="'cm'"/>
											<xsl:with-param name ="length" select ="(/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[2]/child::node()[1]/@text:space-before"/>
										</xsl:call-template>
										<xsl:value-of select="(/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[1]/child::node()[1]/@text:space-before"/>-->
									</xsl:attribute>
								</a:spcPct>
							</a:spcBef>
							<a:buFont>
								<xsl:attribute name="typeface">
									<xsl:value-of select="'StarSymbol'"/>
									<!--<xsl:value-of select="(document/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet)[1]/child::node()[2]/@fo:font-family"/>-->
								</xsl:attribute>
								<xsl:attribute name="pitchFamily">
									<xsl:value-of select="'0'"/>
								</xsl:attribute>
								<xsl:attribute name="charset">
									<xsl:value-of select="'1'"/>
								</xsl:attribute>
							</a:buFont>
							<a:buChar>
								<xsl:attribute name="char">
									<!--<xsl:value-of select="'●'"/>-->
									<xsl:value-of select="(document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:graphic-properties/text:list-style/text:list-level-style-bullet/@text:bullet-char)[9]"/>
								</xsl:attribute>
							</a:buChar>
							<a:defRPr>
								<xsl:attribute name="sz">
									<xsl:call-template name ="convertToPoints">
										<xsl:with-param name ="unit" select ="'pt'"/>
										<xsl:with-param name ="length" select ="./style:text-properties/@fo:font-size"/>
									</xsl:call-template>
								</xsl:attribute>
								<!--Font Bold attribute -->
								<xsl:if test="./style:text-properties/@fo:font-weight[contains(.,'bold')]">
									<xsl:attribute name ="b">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
								</xsl:if >
								<!-- Font Italic attribute-->
								<xsl:if test="./style:text-properties/@fo:font-style[contains(.,'italic')]">
									<xsl:attribute name ="i">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
								</xsl:if >
								<!-- Font Underline-->
								<!--<xsl:if test="not(./style:text-properties/style:text-underline-style[contains(.,'none')])">
									<xsl:attribute name="u">
										<xsl:call-template name="Underline"/>
									</xsl:attribute>
								</xsl:if>-->
								<xsl:attribute name="kern">
									<xsl:value-of select="'0'"/>
								</xsl:attribute>
								<!--<xsl:attribute name="strike">
									<xsl:choose >
										<xsl:when  test="style:text-properties/@style:text-line-through-type[contains(.,'solid')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'sngStrike'"/>
											</xsl:attribute >
										</xsl:when >
										<xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'dblStrike'"/>
											</xsl:attribute >
										</xsl:when >
										<xsl:when test="style:text-properties/@style:text-line-through-style[contains(.,'solid')]">
											<xsl:attribute name ="strike">
												<xsl:value-of select ="'sngStrike'"/>
											</xsl:attribute >
										</xsl:when>
									</xsl:choose>
								</xsl:attribute>-->
								<a:solidFill>
									<a:srgbClr>
										<xsl:attribute name="val">
											<xsl:value-of select="substring-after(./style:text-properties/@fo:color,'#')" />
											<xsl:if test="not(./style:text-properties/@fo:color)">
												<xsl:value-of select="'000000'" />
											</xsl:if>
										</xsl:attribute>
									</a:srgbClr>
								</a:solidFill>
								<a:effectLst/>
								<a:latin>
									<xsl:attribute name="typeface">
										<xsl:choose>
											<xsl:when test="translate(./style:text-properties/@fo:font-family, &quot;'&quot;,'')">
												<xsl:value-of select="translate(./style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>
											</xsl:when>
											<xsl:otherwise>
												<xsl:value-of select="'Arial'"/>
												<!--<xsl:value-of select="translate(document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$outlineName]/style:text-properties/@fo:font-family, &quot;'&quot;,'')"/>-->
											</xsl:otherwise>
										</xsl:choose>
									</xsl:attribute>
								</a:latin>
								<a:ea typeface="+mn-ea"/>
								<a:cs typeface="+mn-cs"/>
							</a:defRPr>
						</xsl:for-each>
					</a:lvl9pPr>
				</p:bodyStyle>
				<p:otherStyle>
					<a:defPPr>
						<a:defRPr lang="en-US"/>
					</a:defPPr>
					<a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
						<a:defRPr sz="1800" kern="1200">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:ea typeface="+mn-ea"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl1pPr>
					<a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
						<a:defRPr sz="1800" kern="1200">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:ea typeface="+mn-ea"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl2pPr>
					<a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
						<a:defRPr sz="1800" kern="1200">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:ea typeface="+mn-ea"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl3pPr>
					<a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
						<a:defRPr sz="1800" kern="1200">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:ea typeface="+mn-ea"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl4pPr>
					<a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
						<a:defRPr sz="1800" kern="1200">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:ea typeface="+mn-ea"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl5pPr>
					<a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
						<a:defRPr sz="1800" kern="1200">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:ea typeface="+mn-ea"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl6pPr>
					<a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
						<a:defRPr sz="1800" kern="1200">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:ea typeface="+mn-ea"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl7pPr>
					<a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
						<a:defRPr sz="1800" kern="1200">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:ea typeface="+mn-ea"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl8pPr>
					<a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
						<a:defRPr sz="1800" kern="1200">
							<a:solidFill>
								<a:schemeClr val="tx1"/>
							</a:solidFill>
							<a:latin typeface="+mn-lt"/>
							<a:ea typeface="+mn-ea"/>
							<a:cs typeface="+mn-cs"/>
						</a:defRPr>
					</a:lvl9pPr>
				</p:otherStyle>
			</p:txStyles>
		</p:sldMaster>
	</xsl:template>

	<xsl:template name="Titlebody">
		<a:lstStyle/>
		<a:p>
			<a:r>
				<a:rPr lang="en-US" dirty="0" smtClean="0"/>
				<a:t>
					<!--<xsl:value-of select="(/office:document-styles/office:master-styles/style:master-page/draw:frame/)[3]/draw:text-box/text:p"/>-->
					<!--<xsl:value-of select="/office:document-styles/office:master-styles/style:master-page/draw:frame/draw:text-box/text:p"/>-->
					
					<xsl:if test="./@presentation:class[contains(.,'title')]">
						<xsl:for-each select="./draw:text-box/text:p">
							<xsl:value-of select="."/>
						</xsl:for-each>
					</xsl:if>
				</a:t>
			</a:r>
			<a:endParaRPr lang="en-US" dirty="0"/>
		</a:p>
	</xsl:template>
	<xsl:template name="Outlinebody">
		<a:lstStyle/>
		<a:p>
			<a:pPr lvl="0"/>
			<a:r>
				<a:rPr  lang="en-US" dirty="0" smtClean="0"/>
				<a:t>
					<xsl:value-of select="(draw:text-box//child::node()/text:p)[1]"/>
				</a:t>
			</a:r>
		</a:p>
		<a:p>
			<a:pPr lvl="1"/>
			<a:r>
				<a:rPr lang="en-US" dirty="0" smtClean="0"/>
				<a:t>
					<xsl:value-of select="(draw:text-box//child::node()/text:p)[2]"/>
				</a:t>
			</a:r>
		</a:p>
		<a:p>
			<a:pPr lvl="2"/>
			<a:r>
				<a:rPr lang="en-US" dirty="0" smtClean="0"/>
				<a:t>
					<xsl:value-of select="(draw:text-box//child::node()/text:p)[3]"/>
				</a:t>
			</a:r>
		</a:p>
		<a:p>
			<a:pPr lvl="3"/>
			<a:r>
				<a:rPr lang="en-US" dirty="0" smtClean="0"/>
				<a:t>
					<xsl:value-of select="(draw:text-box//child::node()/text:p)[4]"/>
				</a:t>
			</a:r>
		</a:p>
		<a:p>
			<a:pPr lvl="4"/>
			<a:r>
				<a:rPr lang="en-US" dirty="0" smtClean="0"/>
				<a:t>
					<xsl:value-of select="(draw:text-box//child::node()/text:p)[5]"/>
				</a:t>
			</a:r>
		</a:p>
		<a:p>
			<a:pPr lvl="5"/>
			<a:r>
				<a:rPr lang="en-US" dirty="0" smtClean="0"/>
				<a:t>
					<xsl:value-of select="(draw:text-box//child::node()/text:p)[6]"/>
				</a:t>
			</a:r>
		</a:p>
		<a:p>
			<a:pPr lvl="6"/>
			<a:r>
				<a:rPr lang="en-US" dirty="0" smtClean="0"/>
				<a:t>
					<xsl:value-of select="(draw:text-box//child::node()/text:p)[7]"/>
				</a:t>
			</a:r>
		</a:p>
		<a:p>
			<a:pPr lvl="7"/>
			<a:r>
				<a:rPr lang="en-US" dirty="0" smtClean="0"/>
				<a:t>
					<xsl:value-of select="(draw:text-box//child::node()/text:p)[8]"/>
				</a:t>
			</a:r>
		</a:p>
		<a:p>
			<a:pPr lvl="8"/>
			<a:r>
				<a:rPr lang="en-US" smtClean="0"/>
				<a:t>
					<xsl:value-of select="(draw:text-box//child::node()/text:p)[9]"/>
				</a:t>
			</a:r>
			
			<a:endParaRPr lang="en-US"/>
		</a:p>
	</xsl:template>
	<xsl:template name="Datetimebody">
		<a:lstStyle>
			<a:lvl1pPr>
				<xsl:attribute name="algn">
					<!--<xsl:value-of select="'ctr'"/>-->
						<xsl:if test ="./draw:text-box/text:p">
							<xsl:variable name ="ParId">
								<xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
							</xsl:variable>
							<xsl:for-each select ="document('styles.xml')//style:style[@style:name=$ParId]/style:paragraph-properties">
								<xsl:choose >
									<xsl:when test ="./@fo:text-align='center'">
										<xsl:value-of select ="'ctr'"/>
									</xsl:when>
									<xsl:when test ="./@fo:text-align='end'">
										<xsl:value-of select ="'r'"/>
									</xsl:when>
									<xsl:otherwise >
										<xsl:value-of select ="'l'"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:for-each>
						</xsl:if>
				</xsl:attribute>
				<a:defRPr>
						<!--<xsl:value-of select="'0'"/>-->
								<xsl:variable name ="textId">
									<xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
								</xsl:variable>
						<xsl:if test="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@fo:font-size">
							<xsl:attribute name="sz">
								<xsl:for-each select ="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties">
									<xsl:call-template name ="convertToPoints">
										<xsl:with-param name ="unit" select ="'pt'"/>
										<xsl:with-param name ="length" select ="@fo:font-size"/>
									</xsl:call-template>
								</xsl:for-each>
						</xsl:attribute>
					</xsl:if>
					<a:solidFill>
						<a:srgbClr>
							<xsl:attribute name="val">
								<xsl:value-of select="'000000'"/>
							</xsl:attribute>
						</a:srgbClr>
					</a:solidFill>
				</a:defRPr>
			</a:lvl1pPr>
		</a:lstStyle>
		<a:p>
			<a:fld type="datetimeFigureOut">
				<xsl:attribute name ="id">
					<xsl:value-of select ="'{11859447-1FF1-4571-A7A1-75F1CAF7F5D9}'"/>
				</xsl:attribute>
				<a:rPr lang="en-US" smtClean="0"/>
				<a:t>
					<!--<xsl:value-of select="(/office:document-styles/office:master-styles/style:master-page/draw:frame/draw:text-box/text:p/text:span)[1]/child::node()[1]"/>-->
				</a:t>
			</a:fld>
			<a:endParaRPr lang="en-US" dirty="0"/>
		</a:p>
	</xsl:template>
	<xsl:template name="Footerbody">
		<a:lstStyle>
			<a:lvl1pPr>
				<xsl:attribute name="algn">
					<!--<xsl:value-of select="'ctr'"/>-->
					<xsl:if test ="./draw:text-box/text:p">
						<xsl:variable name ="ParId">
							<xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
						</xsl:variable>
						<xsl:for-each select ="document('styles.xml')//style:style[@style:name=$ParId]/style:paragraph-properties">
							<xsl:choose >
								<xsl:when test ="./@fo:text-align='center'">
									<xsl:value-of select ="'ctr'"/>
								</xsl:when>
								<xsl:when test ="./@fo:text-align='end'">
									<xsl:value-of select ="'r'"/>
								</xsl:when>
								<xsl:otherwise >
									<xsl:value-of select ="'l'"/>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:for-each>
					</xsl:if>
				</xsl:attribute>
				<a:defRPr>
						<!--<xsl:value-of select="'0'"/>-->
						<xsl:variable name ="textId">
							<xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
						</xsl:variable>
						<xsl:if test="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@fo:font-size">
						<xsl:attribute name="sz">
							<xsl:for-each select ="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name ="unit" select ="'pt'"/>
									<xsl:with-param name ="length" select ="@fo:font-size"/>
								</xsl:call-template>
							</xsl:for-each>
						</xsl:attribute>
						</xsl:if>
					<a:solidFill>
						<a:srgbClr>
							<xsl:attribute name="val">
								<xsl:value-of select="'000000'"/>
							</xsl:attribute>
						</a:srgbClr>
					</a:solidFill>
				</a:defRPr>
			</a:lvl1pPr>
		</a:lstStyle>
		<a:p>
			<!--<a:rPr lang="en-US" smtClean="0"/>
			<a:t>
				--><!--<xsl:value-of select="(/office:document-styles/office:master-styles/style:master-page/draw:frame/draw:text-box/text:p/text:span)[2]/child::node()[1]"/>--><!--
			</a:t>-->			
			<a:endParaRPr lang="en-US" dirty="0"/>
		</a:p>
	</xsl:template>
	<xsl:template name="Pagenumberbody">
		<a:lstStyle>
			<a:lvl1pPr>
				<xsl:attribute name="algn">
					<!--<xsl:value-of select="'ctr'"/>-->
					<xsl:if test ="./draw:text-box/text:p">
						<xsl:variable name ="ParId">
							<xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
						</xsl:variable>
						<xsl:for-each select ="document('styles.xml')//style:style[@style:name=$ParId]/style:paragraph-properties">
							<xsl:choose >
								<xsl:when test ="./@fo:text-align='center'">
									<xsl:value-of select ="'ctr'"/>
								</xsl:when>
								<xsl:when test ="./@fo:text-align='end'">
									<xsl:value-of select ="'r'"/>
								</xsl:when>
								<xsl:otherwise >
									<xsl:value-of select ="'l'"/>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:for-each>
					</xsl:if>
				</xsl:attribute>
				<a:defRPr>
						<!--<xsl:value-of select="'0'"/>-->
						<xsl:variable name ="textId">
							<xsl:value-of select ="./draw:text-box/text:p/@text:style-name"/>
						</xsl:variable>
						<xsl:if test="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties/@fo:font-size">
						<xsl:attribute name="sz">
							<xsl:for-each select ="document('styles.xml')//style:style[@style:name=$textId]/style:text-properties">
								<xsl:call-template name ="convertToPoints">
									<xsl:with-param name ="unit" select ="'pt'"/>
									<xsl:with-param name ="length" select ="@fo:font-size"/>
								</xsl:call-template>
							</xsl:for-each>
						</xsl:attribute>
						</xsl:if>
					<a:solidFill>
						<a:srgbClr>
							<xsl:attribute name="val">
								<xsl:value-of select="'000000'"/>
							</xsl:attribute>
						</a:srgbClr>
					</a:solidFill>
				</a:defRPr>
			</a:lvl1pPr>
		</a:lstStyle>
		<a:p>
			<a:fld type="slidenum">
				<xsl:attribute name ="id">
					<xsl:value-of select ="'{763D1470-AB83-4C4C-B3B3-7F0C9DC8E8D6}'"/>
					</xsl:attribute>
					<a:rPr lang="en-US" smtClean="0" />
					<!--<a:pPr />-->
					<a:t>
						<!--<xsl:value-of select="(/office:document-styles/office:master-styles/style:master-page/draw:frame/draw:text-box/text:p/text:span)[3]/child::node()[1]"/>-->
					</a:t>
				</a:fld>
				<a:endParaRPr lang="en-US" dirty="0" />
			</a:p>
	</xsl:template>
	<xsl:template name="Underline">
		<!-- Font underline-->
		<xsl:param name ="prClassName"/>
		<xsl:variable name="titleName">
			<xsl:value-of select="concat(document('styles.xml')/office:document-styles/office:master-styles/style:master-page/@style:name,'-title')"/>
		</xsl:variable>
		<xsl:for-each select="/office:document-styles/office:styles/style:style[@style:name=$titleName]">
			<xsl:choose >
				<xsl:when test="style:text-properties/@style:text-underline-style[contains(.,'solid')] and
								style:text-properties/@style:text-underline-type[contains(.,'double')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dbl'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style[contains(.,'solid')] and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'heavy'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style[contains(.,'solid')] and
							style:text-properties/@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'sng'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- Dotted lean and dotted bold under line -->
				<xsl:when test="style:text-properties/@style:text-underline-style[contains(.,'dotted')] and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dotted'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style[contains(.,'dotted')] and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dottedHeavy'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- Dash lean and dash bold underline -->
				<xsl:when test="style:text-properties/@style:text-underline-style[contains(.,'dash')] and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dash'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style[contains(.,'dash')] and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dashHeavy'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- Dash long and dash long bold -->
				<xsl:when test="style:text-properties/@style:text-underline-style[contains(.,'long-dash')] and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dashLong'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style[contains(.,'long-dash')] and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dashLongHeavy'"/>
					</xsl:attribute >
				</xsl:when>

				<!-- dot Dash and dot dash bold -->
				<xsl:when test="style:text-properties/@style:text-underline-style[contains(.,'dot-dash')] and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dotDashLong'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style[contains(.,'dot-dash')] and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dotDashHeavy'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- dot-dot-dash-->
				<xsl:when test="style:text-properties/@style:text-underline-style[contains(.,'dot-dot-dash')] and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dotDotDash'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style[contains(.,'dot-dot-dash')] and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dotDotDashHeavy'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- double Wavy -->
				<xsl:when test="style:text-properties/@style:text-underline-style[contains(.,'wave')] and
								style:text-properties/@style:text-underline-type[contains(.,'double')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'wavyDbl'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- Wavy and Wavy Heavy-->
				<xsl:when test="style:text-properties/@style:text-underline-style[contains(.,'wave')] and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'wavy'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style[contains(.,'wave')] and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'wavyHeavy'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:otherwise >
					<xsl:call-template name ="getUnderlineFromStyles">
						<xsl:with-param name ="className" select ="$prClassName"/>
					</xsl:call-template>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:for-each>
	</xsl:template>
	<xsl:template name ="TextAlignment">
		<xsl:param name="slideMasterName"/>
		<xsl:param name ="prId"/>
		<a:bodyPr>
				<xsl:variable name ="ParId">
						<xsl:value-of select ="draw:text-box/text:p/@text:style-name"/>
					</xsl:variable>
					<xsl:for-each select ="document('styles.xml')//style:style[@style:name=$ParId]/style:paragraph-properties">
							<xsl:if test ="./style:paragraph-properties/@style:writing-mode='tb-rl'">
								<xsl:attribute name ="rtlCol">
									<xsl:value-of select ="'0'"/>
								</xsl:attribute>
								<xsl:attribute name ="vert">
									<xsl:value-of select ="'horz'"/>
								</xsl:attribute>
							</xsl:if>
						</xsl:for-each>
						<xsl:for-each select="document('styles.xml')//style:style[@style:name=$prId]/style:graphic-properties">
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
							<!--<xsl:attribute name ="anchor">
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
							</xsl:attribute>-->
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
	<xsl:template name="slideMaster1Rel">
		<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
			<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout8.xml" />
			<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout3.xml" />
			<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout7.xml" />
			<Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml" />
			<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml" />
			<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml" />
			<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout6.xml" />
			<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout11.xml" />
			<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout5.xml" />
			<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout10.xml" />
			<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout4.xml" />
			<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout9.xml" />
		</Relationships>
	</xsl:template>
</xsl:stylesheet>
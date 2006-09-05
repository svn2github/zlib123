<?xml version="1.0" encoding="UTF-8"?>
<!--
 * Copyright (c) 2006, Clever Age
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
 *     * Neither the name of Clever Age nor the names of its contributors 
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
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
  -->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/3/main"
	xmlns:v="urn:schemas-microsoft-com:vml"
	xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
	xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
	xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
	xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
	xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
	xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
	xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
	xmlns:xlink="http://www.w3.org/1999/xlink"
	exclude-result-prefixes="office text table fo style draw xlink v svg"
	xmlns:w10="urn:schemas-microsoft-com:office:word">

	<xsl:strip-space elements="*"/>
	<xsl:preserve-space elements="text:p"/>
	<xsl:preserve-space elements="text:span"/>

	<xsl:key name="annotations" match="//office:annotation" use="''"/>
	<xsl:key name="automatic-styles" match="office:automatic-styles/style:style" use="@style:name"/>
	<xsl:key name="hyperlinks" match="text:a" use="''"/>
	<xsl:key name="headers" match="text:h" use="''"/>
	<xsl:key name="images" match="draw:frame[not(./draw:object-ole) and ./draw:image/@xlink:href]"
		use="''"/>
	<xsl:key name="ole-objects" match="draw:frame[./draw:object-ole] " use="''"/>
	<xsl:key name="master-pages" match="style:master-page" use="@style:name"/>
	<xsl:key name="page-layouts" match="style:page-layout" use="@style:name"/>
	<xsl:key name="master-page-based-styles" match="style:style[@style:master-page-name]"
		use="@style:name"/>


	<!-- COMMENT: what is this variable for? -->
	<xsl:variable name="type">dxa</xsl:variable>

	<!-- main document -->
	<xsl:template name="document">
		<w:document>
			<!-- COMMENT: how are we sure this is the correct background? -->
			<!-- COMMENT: See if we cannot use a key -->
			<xsl:if
				test="document('styles.xml')//style:page-layout[@style:name=//office:master-styles/style:master-page/@style:page-layout-name]/style:page-layout-properties/@fo:background-color">
				<w:background>
					<xsl:attribute name="w:color">
						<xsl:value-of
							select="translate(substring-after(document('styles.xml')//style:page-layout[@style:name=//office:master-styles/style:master-page/@style:page-layout-name]/style:page-layout-properties/@fo:background-color,'#'),'f','F')"
						/>
					</xsl:attribute>
				</w:background>
			</xsl:if>
			<xsl:apply-templates select="document('content.xml')/office:document-content"/>
		</w:document>
	</xsl:template>

	<!-- document body -->
	<xsl:template match="office:body">
		<w:body>
			<xsl:apply-templates/>
			<w:sectPr>
				<!-- Header/Footer configuration -->
				<xsl:call-template name="HeaderFooter"/>
				<!-- Footnotes and endnotes configuration -->
				<xsl:call-template name="footnotes-configuration"/>
				<xsl:call-template name="endnotes-configuration"/>
				<!-- Page layout properties -->
				<!--- all the paragraphs tied to a master style -->
				<xsl:variable name="mp-paragraphs"
					select=".//text:p[key('master-page-based-styles', @text:style-name)]"/>
				<!-- the last one -->
				<xsl:variable name="last-mp-paragraph" select="$mp-paragraphs[last()]"/>
				<!-- the master page name it is related to -->
				<xsl:variable name="master-page-name"
					select="key('master-page-based-styles', $last-mp-paragraph/@text:style-name)/@style:master-page-name"/>

				<xsl:for-each select="document('styles.xml')">
					<xsl:choose>
						<!-- if we use a master-page based style -->
						<xsl:when test="$master-page-name">
							<!-- get the associated page layout -->
							<xsl:variable name="page-layout-name"
								select="key('master-pages', $master-page-name)/@style:page-layout-name"/>
							<!-- apply the layout properties -->
							<xsl:apply-templates
								select="key('page-layouts', $page-layout-name)/style:page-layout-properties"
								mode="master-page"/>
						</xsl:when>
						<xsl:otherwise>
							<!-- use default master page -->
							<!-- TODO : find a more reliable way to spot the default master page -->
							<xsl:apply-templates select="key('page-layouts', 'pm1')" mode="master-page"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:for-each>
				<!-- Shall the header and footer be different on the first page -->
				<xsl:call-template name="TitlePg"/>
			</w:sectPr>
		</w:body>
	</xsl:template>

	<!-- paragraphs and headings -->
	<xsl:template match="text:p | text:h">
		<xsl:param name="level" select="0"/>
		<xsl:message terminate="no">progress:text:p</xsl:message>

		<w:p>
			<xsl:call-template name="InsertParagraphProperties">
				<xsl:with-param name="level" select="$level"/>
			</xsl:call-template>

			<!-- TOC id (used for headings only) -->
			<xsl:variable name="tocId">
				<xsl:if test="self::text:h">
					<xsl:value-of select="number(count(preceding::text:h)+1)"/>
				</xsl:if>
			</xsl:variable>

			<xsl:if test="self::text:h">
				<w:bookmarkStart w:id="{$tocId}" w:name="{concat('_Toc',$tocId)}"/>
			</xsl:if>

			<!-- footnotes or endnotes: insert the mark in the first paragraph -->
			<xsl:if test="parent::text:note-body and position() = 1">
				<xsl:apply-templates select="../../text:note-citation" mode="note"/>
			</xsl:if>

			<xsl:choose>

				<!-- we are in table of contents -->
				<xsl:when test="parent::text:index-body">
					<xsl:call-template name="InsertTocEntry"/>
				</xsl:when>

				<!-- ignore draw:frame/draw:text-box if it's embedded in another draw:frame/draw:text-box becouse word doesn't support it -->
				<xsl:when test="self::node()[ancestor::draw:text-box and descendant::draw:text-box]"/>

				<xsl:otherwise>
					<xsl:apply-templates mode="paragraph"/>
				</xsl:otherwise>

			</xsl:choose>

			<!-- selfstanding image before paragraph-->
			<xsl:if
				test="name(preceding-sibling::*[1]) = 'draw:frame' and preceding-sibling::*[1]/draw:image">
				<xsl:apply-templates select="preceding-sibling::*[1]" mode="paragraph"/>
			</xsl:if>

			<!-- If there is a page-break-after in the paragraph style -->
			<xsl:if
				test="key('automatic-styles',@text:style-name)/style:paragraph-properties/@fo:break-after='page'">
				<w:r>
					<w:br w:type="page"/>
				</w:r>
			</xsl:if>

			<xsl:if test="self::text:h">
				<w:bookmarkEnd w:id="{$tocId}"/>
			</xsl:if>

		</w:p>
	</xsl:template>

	<!-- Inserts the paragraph properties -->
	<xsl:template name="InsertParagraphProperties">
		<xsl:param name="level"/>

		<w:pPr>

			<!-- insert paragraph style -->
			<xsl:call-template name="InsertParagraphStyle">
				<xsl:with-param name="styleName">
					<xsl:choose>
						<xsl:when test="count(draw:frame) = 1 and count(child::text()) = 0">
							<xsl:value-of select="draw:frame/@draw:style-name"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="@text:style-name"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:with-param>
			</xsl:call-template>

			<!-- insert heading outline level -->
			<xsl:call-template name="InsertOutlineLevel">
				<xsl:with-param name="node" select="."/>
			</xsl:call-template>

			<!-- insert indentation -->
			<xsl:call-template name="InsertIndent">
				<xsl:with-param name="level" select="$level"/>
			</xsl:call-template>

			<!-- insert page break before table when required -->
			<xsl:call-template name="InsertPageBreakBefore"/>

		</w:pPr>

		<!-- if we are in an annotation, we may have to insert annotation reference -->
		<xsl:call-template name="InsertAnnotationReference"/>

	</xsl:template>

	<!-- Inserts the style of a paragraph -->
	<xsl:template name="InsertParagraphStyle">
		<xsl:param name="styleName"/>
		<xsl:variable name="prefixedStyleName">
			<xsl:call-template name="GetPrefixedStyleName">
				<xsl:with-param name="styleName" select="$styleName"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:if test="$prefixedStyleName != ''">
			<w:pStyle w:val="{$prefixedStyleName}"/>
		</xsl:if>
	</xsl:template>

	<!-- Inserts the outline level of a heading if needed -->
	<xsl:template name="InsertOutlineLevel">
		<xsl:param name="node"/>
		<xsl:if test="$node[self::text:h]">
			<xsl:choose>
				<xsl:when test="not($node/@text:outline-level)">
					<w:outlineLvl w:val="0"/>
				</xsl:when>
				<xsl:when test="$node/@text:outline-level &lt;= 9">
					<!-- COMMENT: See if we cannot use a key -->
					<xsl:if
						test="document('styles.xml')//office:document-styles/office:styles/text:outline-style/text:outline-level-style/@style:num-format !=''">
						<w:numPr>
							<w:ilvl w:val="{$node/@text:outline-level - 1}"/>
							<w:numId w:val="1"/>
						</w:numPr>
					</xsl:if>
					<w:outlineLvl w:val="{$node/@text:outline-level}"/>
				</xsl:when>
				<xsl:otherwise>
					<w:outlineLvl w:val="9"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:if>
	</xsl:template>

	<!-- Inserts paragraph indentation -->
	<!-- COMMENT: please try to split this template into smaller ones -->
	<xsl:template name="InsertIndent">
		<xsl:param name="level" select="0"/>
		<xsl:variable name="styleName">
			<xsl:call-template name="GetStyleName"/>
		</xsl:variable>
		<xsl:variable name="parentStyleName"
			select="key('automatic-styles',$styleName)/@style:parent-style-name"/>
		<xsl:variable name="listStyleName"
			select="key('automatic-styles',$styleName)/@style:list-style-name"/>
		<xsl:variable name="paragraphMargin">
			<xsl:choose>
				<xsl:when
					test="key('automatic-styles',$styleName)/style:paragraph-properties/@fo:margin-left">
					<xsl:call-template name="twips-measure">
						<xsl:with-param name="length"
							select="key('automatic-styles',$styleName)/style:paragraph-properties/@fo:margin-left"
						/>
					</xsl:call-template>
				</xsl:when>
				<xsl:otherwise>
					<xsl:for-each select="document('styles.xml')">
						<xsl:choose>
							<xsl:when
								test="key('styles',$parentStyleName)/style:paragraph-properties/@fo:margin-left">
								<xsl:call-template name="twips-measure">
									<xsl:with-param name="length"
										select="key('styles',$parentStyleName)/style:paragraph-properties/@fo:margin-left"
									/>
								</xsl:call-template>
							</xsl:when>
							<xsl:otherwise>0</xsl:otherwise>
						</xsl:choose>
					</xsl:for-each>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>

		<xsl:choose>
			<xsl:when test="ancestor-or-self::text:list">
				<xsl:variable name="minLabelWidthTwip">
					<xsl:choose>
						<xsl:when
							test="document('content.xml')//text:list-style[@style:name = $listStyleName]/*[@text:level = $level+1]/style:list-level-properties/@text:min-label-width">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length"
									select="document('content.xml')//text:list-style[@style:name = $listStyleName]/*[@text:level = $level+1]/style:list-level-properties/@text:min-label-width"
								/>
							</xsl:call-template>
						</xsl:when>
						<xsl:when
							test="document('styles.xml')//text:list-style[@style:name = $listStyleName]/*[@text:level = $level+1]/style:list-level-properties/@text:min-label-width">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length"
									select="document('styles.xml')//text:list-style[@style:name = $listStyleName]/*[@text:level = $level+1]/style:list-level-properties/@text:min-label-width"
								/>
							</xsl:call-template>
						</xsl:when>
						<xsl:otherwise>0</xsl:otherwise>
					</xsl:choose>
				</xsl:variable>
				<xsl:variable name="spaceBeforeTwip">
					<xsl:choose>
						<xsl:when
							test="document('content.xml')//text:list-style[@style:name = $listStyleName]/*[@text:level = $level+1]/style:list-level-properties/@text:space-before">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length"
									select="document('content.xml')//text:list-style[@style:name = $listStyleName]/*[@text:level = $level+1]/style:list-level-properties/@text:space-before"
								/>
							</xsl:call-template>
						</xsl:when>
						<xsl:when
							test="document('styles.xml')//text:list-style[@style:name = $listStyleName]/*[@text:level = $level+1]/style:list-level-properties/@text:space-before">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length"
									select="document('styles.xml')//text:list-style[@style:name = $listStyleName]/*[@text:level = $level+1]/style:list-level-properties/@text:space-before"
								/>
							</xsl:call-template>
						</xsl:when>
						<xsl:otherwise>0</xsl:otherwise>
					</xsl:choose>
				</xsl:variable>
				<xsl:variable name="minLabelDistanceTwip">
					<xsl:choose>
						<xsl:when
							test="document('content.xml')//text:list-style[@style:name = $listStyleName]/*[@text:level = $level+1]/style:list-level-properties/@text:min-label-distance">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length"
									select="document('content.xml')//text:list-style[@style:name = $listStyleName]/*[@text:level = $level+1]/style:list-level-properties/@text:min-label-distance"
								/>
							</xsl:call-template>
						</xsl:when>
						<xsl:when
							test="document('styles.xml')//text:list-style[@style:name = $listStyleName]/*[@text:level = $level+1]/style:list-level-properties/@text:min-label-distance">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length"
									select="document('styles.xml')//text:list-style[@style:name = $listStyleName]/*[@text:level = $level+1]/style:list-level-properties/@text:min-label-distance"
								/>
							</xsl:call-template>
						</xsl:when>
						<xsl:when test="$listStyleName=''">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length"
									select="document('styles.xml')//text:outline-style/*[@text:level = $level+1]/style:list-level-properties/@text:min-label-distance"
								/>
							</xsl:call-template>
						</xsl:when>
						<xsl:otherwise>0</xsl:otherwise>
					</xsl:choose>
				</xsl:variable>

				<xsl:choose>
					<xsl:when
						test="not(ancestor-or-self::text:list-header) and (self::text:list-item or not(preceding-sibling::node()))">
						<xsl:if test="$paragraphMargin != 0">
							<xsl:variable name="textIndent">
								<xsl:choose>
									<xsl:when
										test="key('automatic-styles',$styleName)/style:paragraph-properties/@fo:text-indent">
										<xsl:call-template name="twips-measure">
											<xsl:with-param name="length"
												select="key('automatic-styles',$styleName)/style:paragraph-properties/@fo:text-indent"
											/>
										</xsl:call-template>
									</xsl:when>
									<xsl:otherwise>
										<xsl:for-each select="document('styles.xml')">
											<xsl:choose>
												<xsl:when
													test="key('styles',$parentStyleName)/style:paragraph-properties/@fo:text-indent">
													<xsl:call-template name="twips-measure">
														<xsl:with-param name="length"
															select="key('styles',$parentStyleName)/style:paragraph-properties/@fo:text-indent"
														/>
													</xsl:call-template>
												</xsl:when>
												<xsl:otherwise>0</xsl:otherwise>
											</xsl:choose>
										</xsl:for-each>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:variable>
							<xsl:choose>
								<xsl:when test="$textIndent != 0">
									<w:tabs>
										<w:tab>
											<xsl:attribute name="w:val">clear</xsl:attribute>
											<xsl:attribute name="w:pos">
												<xsl:choose>
													<xsl:when
														test="document('content.xml')//text:list-style[@style:name = $listStyleName]/*[@text:level = $level+1]/style:list-level-style-number/@text:display-levels">
														<xsl:value-of select="$spaceBeforeTwip + $minLabelDistanceTwip"/>
													</xsl:when>
													<xsl:otherwise>
														<xsl:choose>
															<xsl:when test="$minLabelWidthTwip &lt; $minLabelDistanceTwip">
																<xsl:value-of
																	select="$spaceBeforeTwip + $minLabelWidthTwip + $minLabelDistanceTwip"
																/>
															</xsl:when>
															<xsl:otherwise>
																<xsl:value-of select="$spaceBeforeTwip + $minLabelWidthTwip"/>
															</xsl:otherwise>
														</xsl:choose>
													</xsl:otherwise>
												</xsl:choose>
											</xsl:attribute>
										</w:tab>
										<w:tab>
											<xsl:attribute name="w:val">num</xsl:attribute>
											<xsl:attribute name="w:pos">
												<xsl:value-of
													select="$minLabelDistanceTwip + $paragraphMargin + $textIndent"/>
											</xsl:attribute>
										</w:tab>
									</w:tabs>
									<w:ind>
										<xsl:attribute name="w:left">
											<xsl:value-of select="$paragraphMargin + $spaceBeforeTwip"/>
										</xsl:attribute>
										<xsl:if
											test="not(ancestor-or-self::text:list-header) and (self::text:list-item or not(preceding-sibling::node()))">
											<xsl:attribute name="w:firstLine">
												<xsl:value-of select="$textIndent"/>
											</xsl:attribute>
										</xsl:if>
									</w:ind>
								</xsl:when>
								<xsl:otherwise>
									<w:tabs>
										<w:tab>
											<xsl:attribute name="w:val">clear</xsl:attribute>
											<xsl:attribute name="w:pos">
												<xsl:choose>
													<xsl:when
														test="document('content.xml')//text:list-style[@style:name = $listStyleName]/*[@text:level = $level+1]/style:list-level-style-number/@text:display-levels">
														<xsl:value-of select="$spaceBeforeTwip + $minLabelDistanceTwip"/>
													</xsl:when>
													<xsl:otherwise>
														<xsl:choose>
															<xsl:when test="$minLabelWidthTwip &lt; $minLabelDistanceTwip">
																<xsl:value-of
																	select="$spaceBeforeTwip + $minLabelWidthTwip + $minLabelDistanceTwip"
																/>
															</xsl:when>
															<xsl:otherwise>
																<xsl:value-of select="$spaceBeforeTwip + $minLabelWidthTwip"/>
															</xsl:otherwise>
														</xsl:choose>
													</xsl:otherwise>
												</xsl:choose>
											</xsl:attribute>
										</w:tab>
										<w:tab>
											<xsl:attribute name="w:val">num</xsl:attribute>
											<xsl:attribute name="w:pos">
												<xsl:choose>
													<xsl:when
														test="document('content.xml')//text:list-style[@style:name = $listStyleName]/*[@text:level = $level+1]/style:list-level-style-number/@text:display-levels">
														<xsl:value-of
															select="$paragraphMargin + $spaceBeforeTwip + $minLabelDistanceTwip"/>
													</xsl:when>
													<xsl:otherwise>
														<xsl:choose>
															<xsl:when test="$minLabelWidthTwip &lt; $minLabelDistanceTwip">
																<xsl:value-of
																	select="$paragraphMargin + $spaceBeforeTwip + $minLabelWidthTwip + $minLabelDistanceTwip"
																/>
															</xsl:when>
															<xsl:otherwise>
																<xsl:value-of
																	select="$paragraphMargin + $spaceBeforeTwip + $minLabelWidthTwip"
																/>
															</xsl:otherwise>
														</xsl:choose>
													</xsl:otherwise>
												</xsl:choose>
											</xsl:attribute>
										</w:tab>
									</w:tabs>
									<w:ind>
										<xsl:attribute name="w:left">
											<xsl:value-of
												select="$paragraphMargin  + $spaceBeforeTwip + $minLabelWidthTwip"/>
										</xsl:attribute>
										<xsl:if
											test="not(ancestor-or-self::text:list-header) and (self::text:list-item or not(preceding-sibling::node()))">
											<xsl:attribute name="w:hanging">
												<xsl:value-of select="$minLabelWidthTwip"/>
											</xsl:attribute>
										</xsl:if>
									</w:ind>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:if>
					</xsl:when>
					<xsl:otherwise>
						<w:ind>
							<xsl:attribute name="w:left">
								<xsl:value-of select="$paragraphMargin  + $spaceBeforeTwip + $minLabelWidthTwip"/>
							</xsl:attribute>
							<xsl:if
								test="not(ancestor-or-self::text:list-header) and (self::text:list-item or not(preceding-sibling::node()))">
								<xsl:attribute name="w:hanging">
									<xsl:value-of select="$minLabelWidthTwip"/>
								</xsl:attribute>
							</xsl:if>
						</w:ind>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<xsl:if test="$paragraphMargin != 0">
					<w:ind>
						<xsl:attribute name="w:left">
							<xsl:value-of select="$paragraphMargin"/>
						</xsl:attribute>
					</w:ind>
				</xsl:if>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- Computes the style name to be used be InsertIndent template -->
	<!-- COMMENT: verify that all cases are martched (I just added self::text:h
       Why not simply match text:list-item and if not everything else? -->
	<!-- COMMENT: see if we cannot reuse this template to factorise all the Set*Styles -->
	<xsl:template name="GetStyleName">
		<xsl:if test="self::text:list-item">
			<xsl:value-of select="*[1][self::text:p]/@text:style-name"/>
		</xsl:if>
		<xsl:if test="parent::text:list-header|self::text:p|self::text:h">
			<xsl:value-of select="@text:style-name"/>
		</xsl:if>
	</xsl:template>

	<!-- Inserts a table of content entry -->
	<xsl:template name="InsertTocEntry">
		<xsl:variable name="num">
			<xsl:choose>
				<xsl:when test="ancestor::text:table-index">
					<xsl:value-of select="count(preceding-sibling::text:p)+count( key('headers',''))+1"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:number/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>

		<xsl:variable name="ile">
			<xsl:number/>
		</xsl:variable>

		<xsl:if test="$ile = 1">
			<w:r>
				<w:fldChar w:fldCharType="begin"/>
			</w:r>
			<w:r>
				<xsl:choose>
					<xsl:when test="ancestor::text:table-of-content">
						<w:instrText xml:space="preserve"> TOC \o "1-<xsl:choose><xsl:when test="parent::text:index-body/preceding-sibling::text:table-of-content-source/@text:outline-level=10">9</xsl:when><xsl:otherwise><xsl:value-of select="parent::text:index-body/preceding-sibling::text:table-of-content-source/@text:outline-level"/></xsl:otherwise></xsl:choose>"<xsl:if test="text:a"> \h </xsl:if></w:instrText>
					</xsl:when>
					<xsl:when test="ancestor::text:illustration-index">
						<w:instrText xml:space="preserve"> TOC  \c "<xsl:value-of select="parent::text:index-body/preceding-sibling::text:illustration-index-source/@text:caption-sequence-name"/>" </w:instrText>
					</xsl:when>
					<xsl:when test="ancestor::text:alphabetical-index">
						<w:instrText xml:space="preserve"> INDEX \e "" \c "<xsl:choose><xsl:when test="key('automatic-styles',ancestor::text:alphabetical-index/@text:style-name)/style:section-properties/style:columns/@fo:column-count=0">1</xsl:when><xsl:otherwise><xsl:value-of select="key('automatic-styles',ancestor::text:alphabetical-index/@text:style-name)/style:section-properties/style:columns/@fo:column-count"/></xsl:otherwise></xsl:choose>" \z "1045" </w:instrText>
					</xsl:when>
					<xsl:otherwise>
						<w:instrText xml:space="preserve"> TOC  \c "<xsl:value-of select="parent::text:index-body/preceding-sibling::text:table-index-source/@text:caption-sequence-name"/>" </w:instrText>
					</xsl:otherwise>
				</xsl:choose>
			</w:r>
			<w:r>
				<w:fldChar w:fldCharType="separate"/>
			</w:r>
		</xsl:if>
		<xsl:choose>
			<!-- COMMENT: duplicate with text:a matching? -->
			<xsl:when test="text:a">
				<w:hyperlink w:history="1">
					<xsl:attribute name="w:anchor">
						<xsl:value-of select="concat('_Toc',$num)"/>
					</xsl:attribute>
					<xsl:call-template name="tableContent">
						<xsl:with-param name="num" select="$num"/>
						<xsl:with-param name="test">1</xsl:with-param>
					</xsl:call-template>
				</w:hyperlink>
			</xsl:when>
			<xsl:otherwise>
				<xsl:call-template name="tableContent">
					<xsl:with-param name="num" select="$num"/>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- Inserts a page break before if needed -->
	<xsl:template name="InsertPageBreakBefore">
		<xsl:if
			test="parent::node()[name()='table:table-cell' and position()=1] and ancestor::node()[name()='table:table-row' and not(preceding-sibling::node())] and key('automatic-styles',ancestor::table:table/@table:style-name)/style:table-properties/@fo:break-before='page'">
			<w:pageBreakBefore/>
		</xsl:if>
	</xsl:template>

	<!-- Inserts an annotation reference if needed -->
	<xsl:template name="InsertAnnotationReference">
		<xsl:if test="ancestor::office:annotation and position() = 1">
			<w:r>
				<w:annotationRef/>
			</w:r>
		</xsl:if>
	</xsl:template>

	<!-- note marks -->
	<xsl:template match="text:note-citation" mode="note">
		<w:r>
			<w:rPr>
				<w:rStyle w:val="{concat(../@text:note-class, 'Reference')}"/>
			</w:rPr>
			<xsl:choose>
				<xsl:when test="../text:note-citation/@text:label">
					<w:t>
						<xsl:value-of select="../text:note-citation"/>
					</w:t>
				</xsl:when>
				<xsl:otherwise>
					<xsl:choose>
						<xsl:when test="../@text:note-class = 'footnote'">
							<w:footnoteRef/>
						</xsl:when>
						<xsl:when test="../@text:note-class = 'endnote' ">
							<w:endnoteRef/>
						</xsl:when>
					</xsl:choose>
				</xsl:otherwise>
			</xsl:choose>
		</w:r>
		<!-- add an extra tab -->
		<w:r>
			<w:tab/>
		</w:r>
	</xsl:template>

	<!-- annotations -->
	<xsl:template match="office:annotation" mode="paragraph">
		<xsl:variable name="id">
			<xsl:call-template name="GenerateId">
				<xsl:with-param name="node" select="."/>
				<xsl:with-param name="nodetype" select="'annotation'"/>
			</xsl:call-template>
		</xsl:variable>
		<w:r>
			<xsl:call-template name="InsertRunProperties"/>
			<w:commentReference>
				<xsl:attribute name="w:id">
					<xsl:value-of select="$id"/>
				</xsl:attribute>
			</w:commentReference>
		</w:r>
	</xsl:template>

	<!-- links -->
	<xsl:template match="text:a" mode="paragraph">
		<xsl:choose>
			<!-- COMMENT: duplicate with TOC handling within paragraphs? -->
			<xsl:when test="ancestor::text:index-body">
				<xsl:variable name="num" select="count(parent::*/preceding-sibling::*)"/>
				<w:hyperlink w:history="1">
					<xsl:attribute name="w:anchor">
						<xsl:value-of select="concat('_Toc',$num)"/>
					</xsl:attribute>
					<xsl:apply-templates mode="paragraph"/>
				</w:hyperlink>
			</xsl:when>
			<xsl:otherwise>
				<w:hyperlink r:id="{generate-id()}">
					<xsl:apply-templates mode="paragraph"/>
				</w:hyperlink>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- COMMENT: what is this template used for? -->
	<xsl:template name="ComputeMarginX">
		<xsl:param name="parent"/>
		<xsl:choose>
			<xsl:when test="$parent">
				<xsl:variable name="recursive_result">
					<xsl:call-template name="point-measure">
						<xsl:with-param name="length">
							<xsl:call-template name="ComputeMarginX">
								<xsl:with-param name="parent" select="$parent[position()>1]"/>
							</xsl:call-template>
						</xsl:with-param>
					</xsl:call-template>
				</xsl:variable>
				<xsl:variable name="svgx">
					<xsl:choose>
						<xsl:when test="$parent[1]/@svg:x">
							<xsl:call-template name="point-measure">
								<xsl:with-param name="length">
									<xsl:value-of select="$parent[1]/@svg:x"/>
								</xsl:with-param>
							</xsl:call-template>
						</xsl:when>
						<xsl:otherwise>0</xsl:otherwise>
					</xsl:choose>

				</xsl:variable>
				<xsl:value-of select="$svgx+$recursive_result"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="0"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- COMMENT: what is this template used for? -->
	<xsl:template name="ComputeMarginY">
		<xsl:param name="parent"/>
		<xsl:choose>
			<xsl:when test="$parent">
				<xsl:variable name="recursive_result">
					<xsl:call-template name="point-measure">
						<xsl:with-param name="length">
							<xsl:call-template name="ComputeMarginY">
								<xsl:with-param name="parent" select="$parent[position()>1]"/>
							</xsl:call-template>
						</xsl:with-param>
					</xsl:call-template>
				</xsl:variable>
				<xsl:variable name="svgy">
					<xsl:choose>
						<xsl:when test="$parent[1]/@svg:y">
							<xsl:call-template name="point-measure">
								<xsl:with-param name="length">
									<xsl:value-of select="$parent[1]/@svg:y"/>
								</xsl:with-param>
							</xsl:call-template>
						</xsl:when>
						<xsl:otherwise>0</xsl:otherwise>
					</xsl:choose>
				</xsl:variable>
				<xsl:value-of select="$svgy+$recursive_result"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="0"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- text boxes -->
	<!-- COMMENT: horrible mess, please try to refactor and split into smaller parts -->
	<xsl:template match="draw:text-box" mode="paragraph">
		<w:r>
			<w:rPr>
				<xsl:call-template name="InsertTextBoxStyle"/>
			</w:rPr>
			<w:pict>
				<v:shapetype/>
				<v:shape type="#_x0000_t202">
					<xsl:variable name="styleGraphicProperties"
						select="key('automatic-styles',parent::draw:frame/@draw:style-name)/style:graphic-properties"/>

					<xsl:variable name="frameW">
						<xsl:call-template name="point-measure">
							<xsl:with-param name="length"
								select="parent::draw:frame/@svg:width|parent::draw:frame/@fo:min-width"/>
						</xsl:call-template>
					</xsl:variable>
					<xsl:variable name="frameH">
						<xsl:call-template name="point-measure">
							<xsl:with-param name="length" select="@fo:min-height|parent::draw:frame/@svg:height"/>
						</xsl:call-template>
					</xsl:variable>
					<xsl:variable name="posL">
						<xsl:if test="parent::draw:frame/@svg:x">
							<xsl:variable name="leftM">
								<xsl:call-template name="ComputeMarginX">
									<xsl:with-param name="parent" select="ancestor::draw:frame"/>
								</xsl:call-template>
							</xsl:variable>
							<xsl:value-of select="$leftM"/>
						</xsl:if>
					</xsl:variable>
					<xsl:variable name="posT">
						<xsl:if test="parent::draw:frame/@svg:y">
							<xsl:variable name="topM">
								<xsl:call-template name="ComputeMarginY">
									<xsl:with-param name="parent" select="ancestor::draw:frame"/>
								</xsl:call-template>
							</xsl:variable>
							<xsl:value-of select="$topM"/>
						</xsl:if>
					</xsl:variable>
					<xsl:variable name="marginL">
						<xsl:call-template name="point-measure">
							<xsl:with-param name="length" select="$styleGraphicProperties/@fo:margin-left"/>
						</xsl:call-template>
					</xsl:variable>

					<xsl:variable name="marginT">
						<xsl:call-template name="point-measure">
							<xsl:with-param name="length" select="$styleGraphicProperties/@fo:margin-top"/>
						</xsl:call-template>
					</xsl:variable>

					<xsl:variable name="marginR">
						<xsl:call-template name="point-measure">
							<xsl:with-param name="length" select="$styleGraphicProperties/@fo:margin-right"/>
						</xsl:call-template>
					</xsl:variable>

					<xsl:variable name="marginB">
						<xsl:call-template name="point-measure">
							<xsl:with-param name="length" select="$styleGraphicProperties/@fo:margin-bottom"/>
						</xsl:call-template>
					</xsl:variable>

					<xsl:variable name="zIndex">
						<xsl:value-of select="parent::draw:frame/@draw:z-index"/>
					</xsl:variable>

					<xsl:variable name="frameWrap"
						select="key('automatic-styles', parent::draw:frame/@draw:style-name)/style:graphic-properties/@style:wrap"/>
					<xsl:variable name="relWidth"
						select="substring-before(parent::draw:frame/@style:rel-width,'%')"/>
					<xsl:variable name="relHeight"
						select="substring-before(parent::draw:frame/@style:rel-height,'%')"/>

					<xsl:attribute name="style">
						<xsl:if test="$frameWrap != 'none' ">
							<xsl:value-of select="'position:absolute;'"/>
						</xsl:if>
						<xsl:if test="not($frameWrap)">
							<xsl:value-of select="'position:absolute;'"/>
						</xsl:if>

						<xsl:value-of select="concat('width:',$frameW,'pt;')"/>
						<xsl:value-of select="concat('height:',$frameH,'pt;')"/>

						<xsl:if test="$relWidth">
							<xsl:value-of select="concat('mso-width-percent:',$relWidth,'0;')"/>
						</xsl:if>
						<xsl:if test="$relHeight">
							<xsl:value-of select="concat('mso-height-percent:',$relHeight,'0;')"/>
						</xsl:if>

						<xsl:value-of select="concat('z-index:', $zIndex, ';')"/>
						<xsl:if test="parent::draw:frame/@svg:x">
							<xsl:value-of select="concat('margin-left:',$posL,'pt;')"/>
						</xsl:if>
						<xsl:if test="parent::draw:frame/@svg:y">
							<xsl:value-of select="concat('margin-top:',$posT,'pt;')"/>
						</xsl:if>
						<xsl:if test="parent::draw:frame/@text:anchor-type = 'page'">
							<xsl:value-of
								select="concat('mso-position-horizontal-relative:',parent::draw:frame/@text:anchor-type,';')"/>
							<xsl:value-of
								select="concat('mso-position-vertical-relative:',parent::draw:frame/@text:anchor-type,';')"
							/>
						</xsl:if>

						<!-- The same style defined in styles.xsl  TODO manage horizontal-rel-->
						<xsl:if
							test="key('automatic-styles', parent::draw:frame/@draw:style-name)/style:graphic-properties/@style:horizontal-pos">
							<xsl:choose>
								<xsl:when
									test="key('automatic-styles', parent::draw:frame/@draw:style-name)/style:graphic-properties/@style:horizontal-pos = 'center'">
									<xsl:value-of select="concat('mso-position-horizontal:', 'center',';')"/>
								</xsl:when>
								<xsl:when
									test="key('automatic-styles', parent::draw:frame/@draw:style-name)/style:graphic-properties/@style:horizontal-pos='left'">
									<xsl:value-of select="concat('mso-position-horizontal:', 'left',';')"/>
								</xsl:when>
								<xsl:when
									test="key('automatic-styles', parent::draw:frame/@draw:style-name)/style:graphic-properties/@style:horizontal-pos='right'">
									<xsl:value-of select="concat('mso-position-horizontal:', 'right',';')"/>
								</xsl:when>
								<!-- <xsl:otherwise><xsl:value-of select="concat('mso-position-horizontal:', 'center',';')"/></xsl:otherwise> -->
							</xsl:choose>
						</xsl:if>
						<xsl:if test="parent::draw:frame/@fo:min-width">
							<xsl:value-of select="'mso-wrap-style:none;'"/>
						</xsl:if>
						<xsl:if
							test="key('automatic-styles', parent::draw:frame/@draw:style-name)/style:graphic-properties/@fo:margin-left">
							<xsl:value-of select="concat('mso-wrap-distance-left:', $marginL,'pt;')"/>
						</xsl:if>
						<xsl:if
							test="key('automatic-styles', parent::draw:frame/@draw:style-name)/style:graphic-properties/@fo:margin-top">
							<xsl:value-of select="concat('mso-wrap-distance-top:', $marginT,'pt;')"/>
						</xsl:if>
						<xsl:if
							test="key('automatic-styles', parent::draw:frame/@draw:style-name)/style:graphic-properties/@fo:margin-right">
							<xsl:value-of select="concat('mso-wrap-distance-right:', $marginR,'pt;')"/>
						</xsl:if>
						<xsl:if
							test="key('automatic-styles', parent::draw:frame/@draw:style-name)/style:graphic-properties/@fo:margin-bottom">
							<xsl:value-of select="concat('mso-wrap-distance-bottom:', $marginB,'pt;')"/>
						</xsl:if>

					</xsl:attribute>
					<xsl:if
						test="key('automatic-styles', parent::draw:frame/@draw:style-name)/style:graphic-properties/@fo:background-color">
						<xsl:attribute name="fillcolor">
							<xsl:value-of
								select="key('automatic-styles', parent::draw:frame/@draw:style-name)/style:graphic-properties/@fo:background-color"
							/>
						</xsl:attribute>
					</xsl:if>

					<xsl:variable name="opacity"
						select="100 - substring-before(key('automatic-styles', parent::draw:frame/@draw:style-name)/style:graphic-properties/@style:background-transparency,'%')"/>

					<xsl:if
						test="key('automatic-styles', parent::draw:frame/@draw:style-name)/style:graphic-properties/@style:background-transparency">
						<v:fill>
							<xsl:attribute name="opacity">
								<xsl:value-of select="concat($opacity,'%')"/>
							</xsl:attribute>
						</v:fill>
					</xsl:if>

					<v:textbox>
						<xsl:attribute name="style">
							<xsl:if test="@fo:min-height">
								<xsl:value-of select="'mso-fit-shape-to-text:t'"/>
							</xsl:if>
						</xsl:attribute>
						<xsl:attribute name="inset">

							<xsl:choose>
								<xsl:when
									test="key('automatic-styles', parent::draw:frame/@draw:style-name)/style:graphic-properties/@fo:padding or key('automatic-styles', parent::draw:frame/@draw:style-name)/style:graphic-properties/@fo:padding-top">
									<xsl:call-template name="padding">
										<xsl:with-param name="graphicProperties"
											select="key('automatic-styles', parent::draw:frame/@draw:style-name)/style:graphic-properties"
										/>
									</xsl:call-template>
								</xsl:when>
								<xsl:otherwise>
									<xsl:variable name="parentStyleName">
										<xsl:value-of
											select="key('automatic-styles', parent::draw:frame/@draw:style-name)/@style:parent-style-name"
										/>
									</xsl:variable>
									<xsl:call-template name="padding">
										<xsl:with-param name="graphicProperties"
											select="document('styles.xml')//office:document-styles/office:styles/style:style[@style:name = $parentStyleName]/style:graphic-properties"
										/>
									</xsl:call-template>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
						<w:txbxContent>
							<xsl:for-each select="child::node()">
								<xsl:apply-templates select="."/>
							</xsl:for-each>
						</w:txbxContent>

						<xsl:message>
							<xsl:value-of select="$frameWrap"/>
						</xsl:message>

						<!--frame wrap-->
						<xsl:choose>
							<xsl:when test="$frameWrap = 'none' ">
								<w10:wrap type="none"/>
								<w10:anchorlock/>
							</xsl:when>
							<xsl:when test="$frameWrap = 'left' ">
								<w10:wrap type="square" side="left"/>
							</xsl:when>
							<xsl:when test="$frameWrap = 'right' ">
								<w10:wrap type="square" side="right"/>
							</xsl:when>
							<xsl:when test="not($frameWrap)">
								<w10:wrap type="square"/>
							</xsl:when>
							<xsl:when test="$frameWrap = 'parallel' ">
								<w10:wrap type="square"/>
							</xsl:when>
							<xsl:when test="$frameWrap = 'dynamic' ">
								<w10:wrap type="square" side="largest"/>
							</xsl:when>
						</xsl:choose>



					</v:textbox>
				</v:shape>
			</w:pict>
		</w:r>
	</xsl:template>
	<!--converts padding into inset-->
	<xsl:template name="padding">
		<xsl:param name="graphicProperties"/>
		<xsl:choose>
			<xsl:when test="$graphicProperties/@fo:padding">
				<xsl:variable name="padding">
					<xsl:call-template name="milimeter-measure">
						<xsl:with-param name="length" select="$graphicProperties/@fo:padding"/>
					</xsl:call-template>
				</xsl:variable>
				<xsl:value-of select="concat($padding,'mm,',$padding,'mm,',$padding,'mm,',$padding,'mm')"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:variable name="padding-top">
					<xsl:if test="$graphicProperties/@fo:padding-top">
						<xsl:call-template name="milimeter-measure">
							<xsl:with-param name="length"
								select="key('automatic-styles', parent::draw:frame/@draw:style-name)/style:graphic-properties/@fo:padding-top"
							/>
						</xsl:call-template>
					</xsl:if>
				</xsl:variable>
				<xsl:variable name="padding-right">
					<xsl:if test="$graphicProperties/@fo:padding-right">
						<xsl:call-template name="milimeter-measure">
							<xsl:with-param name="length" select="$graphicProperties/@fo:padding-right"/>
						</xsl:call-template>
					</xsl:if>
				</xsl:variable>
				<xsl:variable name="padding-bottom">
					<xsl:if test="$graphicProperties/@fo:padding-bottom">
						<xsl:call-template name="milimeter-measure">
							<xsl:with-param name="length" select="$graphicProperties/@fo:padding-bottom"/>
						</xsl:call-template>
					</xsl:if>
				</xsl:variable>
				<xsl:variable name="padding-left">
					<xsl:if test="$graphicProperties/@fo:padding-left">
						<xsl:call-template name="milimeter-measure">
							<xsl:with-param name="length" select="$graphicProperties/@fo:padding-left"/>
						</xsl:call-template>
					</xsl:if>
				</xsl:variable>
				<xsl:if test="$graphicProperties/@fo:padding-top">
					<xsl:value-of
						select="concat($padding-left,'mm,',$padding-top,'mm,',$padding-right,'mm,',$padding-bottom,'mm')"
					/>
				</xsl:if>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!-- Inserts the style of a text box -->
	<xsl:template name="InsertTextBoxStyle">
		<xsl:variable name="prefixedStyleName">
			<xsl:call-template name="GetPrefixedStyleName">
				<xsl:with-param name="styleName" select="parent::draw:frame/@draw:style-name"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:if test="$prefixedStyleName!=''">
			<w:rStyle w:val="{$prefixedStyleName}"/>
		</xsl:if>
	</xsl:template>

	<!-- @TODO  positioning text-boxes -->
	<xsl:template match="draw:frame" mode="paragraph">
				<xsl:call-template name="InsertEmbeddedTextboxes"/>
	</xsl:template>
	
	<xsl:template match="draw:frame">
		<w:p>
			<xsl:call-template name="InsertEmbeddedTextboxes"/>
		</w:p>
	</xsl:template>
    
	<!-- inserts textboxes which are embedded in odf as one after another in word -->
	<xsl:template name="InsertEmbeddedTextboxes">
		<xsl:for-each select="descendant::draw:text-box">
			<xsl:apply-templates mode="paragraph" select="."/>
		</xsl:for-each>
	</xsl:template>

	<!-- lists -->
	<xsl:template match="text:list">
		<xsl:param name="level" select="-1"/>
		<xsl:apply-templates>
			<xsl:with-param name="level" select="$level+1"/>
		</xsl:apply-templates>
	</xsl:template>

	<!-- list headers -->
	<xsl:template match="text:list-header">
		<xsl:param name="level"/>
		<xsl:apply-templates>
			<xsl:with-param name="level" select="$level"/>
		</xsl:apply-templates>
	</xsl:template>

	<!-- list items -->
	<xsl:template match="text:list-item">
		<xsl:param name="level"/>
		<xsl:choose>
			<xsl:when test="*[1][self::text:p or self::text:h]">
				<w:p>
					<w:pPr>

						<!-- insert style -->
						<xsl:call-template name="InsertParagraphStyle">
							<xsl:with-param name="styleName" select="*[1]/@text:style-name"/>
						</xsl:call-template>

						<!-- insert number -->
						<xsl:call-template name="InsertListItemNumber">
							<xsl:with-param name="level" select="$level"/>
						</xsl:call-template>

						<!-- override abstract num indent and tab if paragraph has margin defined -->
						<xsl:call-template name="InsertIndent">
							<xsl:with-param name="level" select="$level"/>
						</xsl:call-template>

						<!-- insert heading outline level -->
						<xsl:call-template name="InsertOutlineLevel">
							<xsl:with-param name="node" select="*[1]"/>
						</xsl:call-template>

						<!-- insert page break before table when required -->
						<xsl:call-template name="InsertPageBreakBefore"/>
					</w:pPr>

					<!-- if we are in an annotation, we may have to insert annotation reference -->
					<xsl:call-template name="InsertAnnotationReference"/>

					<!-- footnote or endnote - Include the mark to the first paragraph only when first child of 
          text:note-body is not paragraph -->
					<xsl:if
						test="ancestor::text:note and not(ancestor::text:note-body/child::*[1][self::text:p | self::text:h]) and position() = 1">
						<xsl:apply-templates select="ancestor::text:note/text:note-citation" mode="note"/>
					</xsl:if>

					<xsl:apply-templates mode="paragraph"/>

				</w:p>

				<!-- others (text:p or text:list) -->
				<xsl:apply-templates select="*[position() != 1]">
					<xsl:with-param name="level" select="$level"/>
				</xsl:apply-templates>
			</xsl:when>
			<xsl:otherwise>
				<xsl:apply-templates>
					<xsl:with-param name="level" select="$level"/>
				</xsl:apply-templates>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- Inserts the number of a list item -->
	<xsl:template name="InsertListItemNumber">
		<xsl:param name="level"/>
		<w:numPr>
			<w:ilvl w:val="{$level}"/>
			<w:numId>
				<xsl:attribute name="w:val">
					<xsl:call-template name="numberingId">
						<xsl:with-param name="styleName" select="ancestor::text:list/@text:style-name"/>
					</xsl:call-template>
				</xsl:attribute>
			</w:numId>
		</w:numPr>
	</xsl:template>

	<!-- COMMENT: please be more explicit about the goal of this template and find a more explicit name -->
	<!-- table of contents -->
	<xsl:template name="tableContent">
		<xsl:param name="num"/>
		<!-- COMMENT: what is this "test" param for??? Please use more significant name -->
		<xsl:param name="test" select="0"/>
		<w:r>
			<xsl:choose>
				<xsl:when test="$test=1">
					<w:t>
						<xsl:for-each select="text:a/child::text()[position() &lt; last()]">
							<xsl:value-of select="."/>
						</xsl:for-each>
					</w:t>
				</xsl:when>
				<xsl:otherwise>
					<w:t>
						<xsl:choose>
							<xsl:when test="number(child::text()[last()])">
								<xsl:for-each select="child::node()[position() &lt; last()]">
									
									<xsl:choose>
										<xsl:when test="self::text()">
											<xsl:value-of select="."/> 
										</xsl:when>
										<xsl:otherwise>
											<xsl:apply-templates select="."/>
										</xsl:otherwise>
									</xsl:choose>
								</xsl:for-each>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="child::text()[last()]"/>
							</xsl:otherwise>
						</xsl:choose>
					</w:t>
					<!--<xsl:apply-templates select="child::text()[1]" mode="text"/>-->
				</xsl:otherwise>
			</xsl:choose>
		</w:r>
		<xsl:apply-templates select="text:tab|text:a/text:tab" mode="paragraph"/>
		<xsl:if test="not(ancestor::text:alphabetical-index)">
			<w:r>
				<w:rPr/>
				<w:fldChar w:fldCharType="begin">
					<w:fldData xml:space="preserve">CNDJ6nn5us4RjIIAqgBLqQsCAAAACAAAAA4AAABfAFQAbwBjADEANAAxADgAMwA5ADIANwA2AAAA</w:fldData>
				</w:fldChar>
			</w:r>
			<w:r>
				<w:rPr/>
				<w:instrText xml:space="preserve"><xsl:value-of select="concat('PAGEREF _Toc', $num, ' \h')"/></w:instrText>
			</w:r>
			<w:r>
				<w:rPr/>
				<w:fldChar w:fldCharType="separate"/>
			</w:r>
		</xsl:if>
		<w:r>
			<xsl:choose>
				<xsl:when test="$test=1">
					<xsl:apply-templates select="text:a/child::text()[last()]" mode="text"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:apply-templates select="child::text()[last()]" mode="text"/>
					<!--<xsl:apply-templates select="child::text()[1]" mode="text"/>-->
				</xsl:otherwise>
			</xsl:choose>
		</w:r>
		<xsl:if test="not(ancestor::text:alphabetical-index)">
			<w:r>
				<w:rPr/>
				<w:fldChar w:fldCharType="end"/>
			</w:r>
		</xsl:if>
	</xsl:template>

	<!-- indexes -->
	<xsl:template match="text:table-index|text:alphabetical-index|text:illustration-index">
		<xsl:if test="text:index-body/text:index-title/text:p">
			<xsl:apply-templates select="text:index-body/text:index-title/text:p"/>
		</xsl:if>

		<xsl:for-each select="text:index-body/child::text:p">
			<xsl:variable name="num">
				<xsl:value-of select="position()+count( key('headers',''))+1"/>
			</xsl:variable>
			<xsl:apply-templates select="."/>
		</xsl:for-each>
		<w:p>
			<w:r>
				<w:rPr/>
				<w:fldChar w:fldCharType="end"/>
			</w:r>
		</w:p>
	</xsl:template>

	<!-- table of content -->
	<xsl:template match="text:table-of-content">
		<w:sdt>
			<w:sdtPr>
				<w:docPartObj>
					<w:docPartType w:val="'Table of Contents'"/>
				</w:docPartObj>
			</w:sdtPr>
			<w:sdtContent>
				<xsl:if test="text:index-body/text:index-title/text:p">
					<xsl:apply-templates select="text:index-body/text:index-title/text:p"/>
				</xsl:if>
				<xsl:for-each select="text:index-body/child::text:p">
					<xsl:variable name="num">
						<xsl:number/>
					</xsl:variable>
					<xsl:apply-templates select="."/>
				</xsl:for-each>
				<w:p>
					<w:r>
						<w:rPr/>
						<w:fldChar w:fldCharType="end"/>
					</w:r>
				</w:p>
			</w:sdtContent>
		</w:sdt>
	</xsl:template>

	<!-- tables -->
	<xsl:template match="table:table">
		<xsl:variable name="styleName">
			<xsl:value-of select="@table:style-name"/>
		</xsl:variable>
		<w:tbl>
			<w:tblPr>
				<w:tblStyle w:val="{@table:style-name}"/>
				<xsl:variable name="tableProp"
					select="key('automatic-styles', @table:style-name)/style:table-properties"/>
				<w:tblW w:type="{$type}">
					<xsl:attribute name="w:w">
						<xsl:call-template name="twips-measure">
							<xsl:with-param name="length"
								select="key('automatic-styles', @table:style-name)/style:table-properties/@style:width"
							/>
						</xsl:call-template>
					</xsl:attribute>
				</w:tblW>
				<xsl:if
					test="key('automatic-styles', @table:style-name)/style:table-properties/@table:align">
					<xsl:choose>
						<xsl:when
							test="key('automatic-styles', @table:style-name)/style:table-properties/@table:align = 'margins'">
							<w:jc w:val="left"/>
							<!--User agents that do not support the "margins" value, may treat this value as "left".-->
						</xsl:when>
						<xsl:otherwise>
							<w:jc
								w:val="{key('automatic-styles', @table:style-name)/style:table-properties/@table:align}"
							/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:if>
				<xsl:if
					test="key('automatic-styles', @table:style-name)/style:table-properties/@fo:margin-left != '' ">
					<w:tblInd w:type="{$type}">
						<xsl:attribute name="w:w">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length"
									select="key('automatic-styles', @table:style-name)/style:table-properties/@fo:margin-left"
								/>
							</xsl:call-template>
						</xsl:attribute>
					</w:tblInd>
				</xsl:if>
				<!-- Default layout algorithm in ODF is "fixed". -->
				<w:tblLayout w:type="fixed"/>

				<!--table background-->
				<xsl:if test="$tableProp/@fo:background-color">
					<xsl:choose>
						<xsl:when test="$tableProp/@fo:background-color != 'transparent' ">
							<w:shd w:val="clear" w:color="auto"
								w:fill="{substring($tableProp/@fo:background-color, 2, string-length($tableProp/@fo:background-color) -1)}"
							/>
						</xsl:when>
					</xsl:choose>
				</xsl:if>

			</w:tblPr>
			<w:tblGrid>
				<xsl:apply-templates select="table:table-column"/>
			</w:tblGrid>
			<xsl:apply-templates
				select="table:table-rows|table:table-header-rows|table:table-row|table:table-header-row"/>
		</w:tbl>
	</xsl:template>

	<!-- COMMENT: please rename the template and add a description -->
	<xsl:template name="subtable">
		<xsl:param name="node"/>
		<xsl:for-each select="$node/table:table-cell">
			<xsl:call-template name="table-cell"/>
		</xsl:for-each>
	</xsl:template>

	<!-- COMMENT: please rename the template and add a description -->
	<xsl:template name="merged-rows">
		<xsl:param name="i" select="0"/>
		<xsl:param name="iterator"/>
		<!-- COMMENT: is this variable really necessary? -->
		<xsl:variable name="test">
			<xsl:if test="$i > 0">
				<xsl:text>true</xsl:text>
			</xsl:if>
		</xsl:variable>
		<xsl:if test="$test='true'">
			<w:tr>
				<xsl:for-each select="table:table-cell">
					<xsl:choose>
						<xsl:when test="table:table[@table:is-sub-table='true']">
							<!-- table to process -->
							<xsl:call-template name="subtable">
								<xsl:with-param name="node" select="table:table/child::table:table-row[$iterator]"/>
							</xsl:call-template>
						</xsl:when>
						<xsl:when test="@table:number-columns-spanned">
							<xsl:choose>
								<xsl:when test="$iterator = 1">
									<xsl:call-template name="table-cell">
										<xsl:with-param name="grid"
											select="round(number(@table:number-columns-spanned))"/>
										<xsl:with-param name="merge" select="1"/>
									</xsl:call-template>
								</xsl:when>
								<xsl:otherwise>
									<xsl:call-template name="table-cell">
										<xsl:with-param name="grid"
											select="round(number(@table:number-columns-spanned))"/>
										<xsl:with-param name="merge" select="2"/>
									</xsl:call-template>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:otherwise>
							<xsl:choose>
								<xsl:when test="$iterator = 1">
									<xsl:call-template name="table-cell">
										<xsl:with-param name="merge" select="1"/>
									</xsl:call-template>
								</xsl:when>
								<xsl:otherwise>
									<xsl:call-template name="table-cell">
										<xsl:with-param name="merge" select="2"/>
									</xsl:call-template>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:for-each>
			</w:tr>
			<xsl:call-template name="merged-rows">
				<xsl:with-param name="i" select="$i  -1"/>
				<xsl:with-param name="iterator" select="$iterator +1"/>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>

	<!-- table columns -->
	<xsl:template match="table:table-column">
		<xsl:param name="repeat" select="1"/>

		<xsl:variable name="columnNumber">
			<xsl:call-template name="CountTableColumns">
				<xsl:with-param name="nodeList" select="preceding-sibling::table:table-column"/>
			</xsl:call-template>
		</xsl:variable>

		<xsl:if test="$columnNumber &lt; 63">
			<!-- relative width not supported yet -->
			<w:gridCol>
				<xsl:attribute name="w:w">
					<xsl:call-template name="twips-measure">
						<xsl:with-param name="length"
							select="key('automatic-styles', @table:style-name)/style:table-column-properties/@style:column-width"
						/>
					</xsl:call-template>
				</xsl:attribute>
			</w:gridCol>
			<xsl:if test="@table:number-columns-repeated ">
				<xsl:if test="@table:number-columns-repeated &gt; $repeat">
					<xsl:apply-templates select=".">
						<xsl:with-param name="repeat" select="$repeat + 1"/>
					</xsl:apply-templates>
				</xsl:if>
			</xsl:if>
		</xsl:if>
	</xsl:template>

	<!-- Counts table:table-column (so the max number of cols is 63) -->
	<xsl:template name="CountTableColumns">
		<xsl:param name="nodeList"/>
		<xsl:choose>
			<xsl:when test="$nodeList">
				<xsl:variable name="recursive_result">
					<xsl:call-template name="CountTableColumns">
						<xsl:with-param name="nodeList" select="$nodeList[position() > 1]"/>
					</xsl:call-template>
				</xsl:variable>
				<xsl:choose>
					<xsl:when test="$nodeList[1]/@table:number-columns-repeated ">
						<xsl:value-of
							select="number($nodeList[1]/@table:number-columns-repeated) + $recursive_result"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="1 + $recursive_result"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="0"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- table rows -->
	<xsl:template match="table:table-row|table:table-header-row">
		<xsl:choose>
			<xsl:when test="table:table-cell/table:table/@table:is-sub-table='true'">
				<!-- merged cells -->
				<xsl:variable name="total_rows"
					select="count(table:table-cell/table:table[@table:is-sub-table='true']/table:table-row)"/>
				<xsl:variable name="subtables"
					select="count(table:table-cell/table:table[@table:is-sub-table='true'])"/>
				<xsl:call-template name="merged-rows">
					<xsl:with-param name="i" select="$total_rows div $subtables"/>
					<xsl:with-param name="iterator" select="1"/>
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<w:tr>
					<xsl:if
						test="key('automatic-styles',child::table:table-cell/@table:style-name)/style:table-cell-properties/@fo:wrap-option='no-wrap'">
						<!-- Override layout algorithm -->
						<w:tblPrEx>
							<w:tblLayout w:type="auto"/>
						</w:tblPrEx>
					</xsl:if>
					<w:trPr>
						<xsl:if test="name(parent::*) = 'table:table-header-rows'">
							<w:tblHeader/>
						</xsl:if>
						<xsl:if test="@table:style-name">
							<xsl:variable name="rowStyle" select="@table:style-name"/>
							<xsl:variable name="widthType"
								select="key('automatic-styles', $rowStyle)/style:table-row-properties"/>

							<xsl:variable name="widthRow">
								<xsl:choose>
									<xsl:when test="$widthType[@style:row-height]">
										<xsl:value-of select="$widthType/@style:row-height"/>
									</xsl:when>
									<xsl:when test="$widthType[@style:min-row-height]">
										<xsl:value-of select="$widthType/@style:min-row-height"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:value-of select="0"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:variable>

							<w:trHeight>
								<xsl:attribute name="w:val">
									<xsl:call-template name="twips-measure">
										<xsl:with-param name="length" select="$widthRow"/>
									</xsl:call-template>
								</xsl:attribute>
								<xsl:attribute name="w:hRule">
									<xsl:choose>
										<xsl:when test="$widthType[@style:row-height]">
											<xsl:value-of select="'exact'"/>
										</xsl:when>
										<xsl:when test="$widthType[@style:min-row-height]">
											<xsl:value-of select="'atLeast'"/>
										</xsl:when>
										<xsl:otherwise>
											<xsl:value-of select="'auto'"/>
										</xsl:otherwise>
									</xsl:choose>
								</xsl:attribute>
							</w:trHeight>
						</xsl:if>
						<!-- row styles -->

						<!--keep together-->
						<xsl:if test="key('automatic-styles', @table:style-name)/style:table-row-properties/@style:keep-together = 'false'
								or key('automatic-styles', substring-before(@table:style-name,'.'))/style:table-properties/@style:may-break-between-rows='false'">
								<w:cantSplit/>
						</xsl:if>
						
					</w:trPr>
					<xsl:apply-templates
						select="table:table-cell[position() &lt; 64]|table:covered-table-cell">
						<xsl:with-param name="colsNumber" select="count(table:table-cell)"/>
					</xsl:apply-templates>
				</w:tr>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- Sums up the widths of a column in case of merged columns -->
	<xsl:template name="ComputeCellWidth">
		<xsl:param name="cellSpanned"/>
		<xsl:param name="col"/>
		<xsl:variable name="table" select="ancestor::table:table[1]"/>
		<xsl:choose>
			<xsl:when test="$cellSpanned &gt; 1">
				<xsl:variable name="recursive_result">
					<xsl:call-template name="ComputeCellWidth">
						<xsl:with-param name="cellSpanned" select="$cellSpanned -1"/>
						<xsl:with-param name="col" select="$col+1"/>
					</xsl:call-template>
				</xsl:variable>
				<xsl:variable name="colStyle"
					select="$table/table:table-column[position() = $col]/@table:style-name"/>
				<xsl:variable name="width">
					<xsl:call-template name="twips-measure">
						<xsl:with-param name="length"
							select="key('automatic-styles', $colStyle)/style:table-column-properties/@style:column-width"
						/>
					</xsl:call-template>
				</xsl:variable>
				<xsl:value-of select="$width + $recursive_result"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:variable name="colStyle"
					select="$table/table:table-column[position() = $col]/@table:style-name"/>
				<xsl:call-template name="twips-measure">
					<xsl:with-param name="length"
						select="key('automatic-styles', $colStyle)/style:table-column-properties/@style:column-width"
					/>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- COMMENT: please rename the template and add a description -->
	<xsl:template name="cellWidth">
		<xsl:param name="col"/>
		<xsl:variable name="cellSpanned">
			<xsl:choose>
				<xsl:when test="@table:number-columns-spanned">
					<xsl:value-of select="@table:number-columns-spanned"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="1"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>

		<xsl:variable name="width">
			<xsl:call-template name="ComputeCellWidth">
				<xsl:with-param name="cellSpanned" select="$cellSpanned"/>
				<xsl:with-param name="col" select="$col"/>
			</xsl:call-template>
		</xsl:variable>

		<w:tcW w:type="{$type}">
			<xsl:attribute name="w:w">
				<xsl:value-of select="$width"/>
			</xsl:attribute>
		</w:tcW>
	</xsl:template>

	<!-- COMMENT: please rename the template and add a description -->
	<xsl:template name="loopCell">
		<xsl:param name="cellPosition"/>
		<xsl:param name="colNum"/>
		<xsl:param name="iterator"/>
		<xsl:param name="repeat"/>

		<xsl:choose>
			<xsl:when test="$iterator = $cellPosition">
				<xsl:call-template name="cellWidth">
					<xsl:with-param name="col">
						<xsl:value-of select="$colNum"/>
					</xsl:with-param>
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:variable name="table" select="ancestor::table:table[1]"/>
				<xsl:variable name="column" select="$table/table:table-column[position() = $colNum]"/>
				<xsl:variable name="repeatVal">
					<xsl:choose>
						<xsl:when test="$repeat &gt; 0">
							<xsl:value-of select="$repeat"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="$column/@table:number-columns-repeated"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:variable>
				<xsl:choose>
					<xsl:when test="$repeatVal and $repeatVal &gt; 1">
						<xsl:call-template name="loopCell">
							<xsl:with-param name="iterator">
								<xsl:value-of select="number($iterator)+1"/>
							</xsl:with-param>
							<xsl:with-param name="cellPosition">
								<xsl:value-of select="$cellPosition"/>
							</xsl:with-param>
							<xsl:with-param name="colNum">
								<xsl:value-of select="$colNum"/>
							</xsl:with-param>
							<xsl:with-param name="repeat">
								<xsl:value-of select="number($repeatVal)-1"/>
							</xsl:with-param>
						</xsl:call-template>
					</xsl:when>
					<xsl:otherwise>
						<xsl:call-template name="loopCell">
							<xsl:with-param name="iterator">
								<xsl:value-of select="number($iterator)+1"/>
							</xsl:with-param>
							<xsl:with-param name="cellPosition">
								<xsl:value-of select="$cellPosition"/>
							</xsl:with-param>
							<xsl:with-param name="colNum">
								<xsl:value-of select="number($colNum)+1"/>
							</xsl:with-param>
							<xsl:with-param name="repeat">-1</xsl:with-param>
						</xsl:call-template>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- table cells -->
	<xsl:template match="table:table-cell" name="table-cell">
		<xsl:param name="colsNumber"/>
		<xsl:param name="grid" select="0"/>
		<xsl:param name="merge" select="0"/>
		<w:tc>
			<w:tcPr>
				<!-- point on the cell style properties -->
				<xsl:variable name="cellProp"
					select="key('automatic-styles', @table:style-name)/style:table-cell-properties"/>
				<xsl:variable name="tableStyle" select="substring-before(@table:style-name, '.')"/>
				<xsl:variable name="rowStyle" select="../@table:style-name"/>
				<xsl:variable name="tableProp"
					select="key('automatic-styles', $tableStyle)/style:table-properties"/>
				<xsl:variable name="rowProp"
					select="key('automatic-styles', $rowStyle)/style:table-row-properties"/>

				<!-- width of the cell -  @TODO handle problem of merged columns or rows-->
				<xsl:call-template name="loopCell">
					<xsl:with-param name="iterator">1</xsl:with-param>
					<xsl:with-param name="cellPosition" select="position()"/>
					<xsl:with-param name="colNum">1</xsl:with-param>
					<xsl:with-param name="repeat">-1</xsl:with-param>
				</xsl:call-template>

				<xsl:choose>
					<xsl:when test="$merge = 1">
						<w:gridSpan w:val="{$grid}"/>
						<w:vmerge w:val="restart"/>
					</xsl:when>
					<xsl:when test="$merge = 2">
						<w:gridSpan w:val="{$grid}"/>
						<w:vmerge w:val="continue"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:if test="@table:number-columns-spanned">
							<w:gridSpan w:val="{@table:number-columns-spanned}"/>
						</xsl:if>
					</xsl:otherwise>
				</xsl:choose>
				<w:tcBorders>
					<xsl:choose>
						<xsl:when test="$cellProp[@fo:border and @fo:border!='none' ]">
							<xsl:variable name="border" select="$cellProp/@fo:border"/>
							<!-- fo:border = "0.002cm solid #000000" -->
							<xsl:variable name="border-color" select="substring-after($border, '#')"/>
							<xsl:variable name="border-size">
								<xsl:call-template name="eightspoint-measure">
									<xsl:with-param name="length" select="substring-before($border,' ')"/>
								</xsl:call-template>
							</xsl:variable>
							<w:top w:val="single" w:color="{$border-color}" w:sz="{$border-size}"/>
							<w:left w:val="single" w:color="{$border-color}" w:sz="{$border-size}"/>
							<w:bottom w:val="single" w:color="{$border-color}" w:sz="{$border-size}"/>
							<w:right w:val="single" w:color="{$border-color}" w:sz="{$border-size}"/>
						</xsl:when>

						<xsl:otherwise>
							<xsl:if test="$cellProp[@fo:border-top and @fo:border-top != 'none']">
								<xsl:variable name="border" select="$cellProp/@fo:border-top"/>
								<w:top w:val="single" w:color="{substring-after($border, '#')}">
									<xsl:attribute name="w:sz">
										<xsl:call-template name="eightspoint-measure">
											<xsl:with-param name="length" select="substring-before($border, ' ')"/>
										</xsl:call-template>
									</xsl:attribute>
								</w:top>
							</xsl:if>
							<xsl:if test="$cellProp[@fo:border-left and @fo:border-left != 'none']">
								<xsl:variable name="border" select="$cellProp/@fo:border-left"/>
								<w:left w:val="single" w:color="{substring-after($border, '#')}">
									<xsl:attribute name="w:sz">
										<xsl:call-template name="eightspoint-measure">
											<xsl:with-param name="length" select="substring-before($border, ' ')"/>
										</xsl:call-template>
									</xsl:attribute>
								</w:left>
							</xsl:if>
							<xsl:if test="$cellProp[@fo:border-bottom and @fo:border-bottom != 'none']">
								<xsl:variable name="border" select="$cellProp/@fo:border-bottom"/>
								<w:bottom w:val="single" w:color="{substring-after($border, '#')}">
									<xsl:attribute name="w:sz">
										<xsl:call-template name="eightspoint-measure">
											<xsl:with-param name="length" select="substring-before($border, ' ')"/>
										</xsl:call-template>
									</xsl:attribute>
								</w:bottom>
							</xsl:if>
							<xsl:if
								test="$cellProp[(@fo:border-right and @fo:border-right != 'none')] or (position() &lt; $colsNumber and position() = 63)">
								<xsl:variable name="border">
									<xsl:choose>
										<xsl:when test="position() &lt; $colsNumber and position() = 63">
											<xsl:value-of select="$cellProp/@fo:border-left"/>
										</xsl:when>
										<xsl:otherwise>
											<xsl:value-of select="$cellProp/@fo:border-right"/>
										</xsl:otherwise>
									</xsl:choose>

								</xsl:variable>
								<w:right w:val="single" w:color="{substring-after($border, '#')}">
									<xsl:attribute name="w:sz">
										<xsl:call-template name="eightspoint-measure">
											<xsl:with-param name="length" select="substring-before($border, ' ')"/>
										</xsl:call-template>
									</xsl:attribute>
								</w:right>
							</xsl:if>
						</xsl:otherwise>
					</xsl:choose>
				</w:tcBorders>

				<!--cell background color-->
				<xsl:choose>
					<xsl:when
						test="$cellProp/@fo:background-color and $cellProp/@fo:background-color != 'transparent' ">
						<w:shd w:val="clear" w:color="auto"
							w:fill="{substring($cellProp/@fo:background-color, 2, string-length($cellProp/@fo:background-color) -1)}"
						/>
					</xsl:when>

					<xsl:otherwise>
						<xsl:choose>
							<xsl:when
								test="$rowProp/@fo:background-color and $rowProp/@fo:background-color != 'transparent' ">
								<w:shd w:val="clear" w:color="auto"
									w:fill="{substring($rowProp/@fo:background-color, 2, string-length($rowProp/@fo:background-color) -1)}"
								/>
							</xsl:when>
							<xsl:when
								test="not($rowProp/@fo:background-color) and $tableProp/@fo:background-color != 'transparent'  and $tableProp/@fo:background-color">
								<w:shd w:val="clear" w:color="auto"
									w:fill="{substring($tableProp/@fo:background-color, 2, string-length($tableProp/@fo:background-color) -1)}"
								/>
							</xsl:when>
						</xsl:choose>
					</xsl:otherwise>
				</xsl:choose>

				<w:tcMar>
					<xsl:choose>
						<xsl:when test="$cellProp[@fo:padding and @fo:padding != 'none']">
							<xsl:variable name="padding">
								<xsl:call-template name="twips-measure">
									<xsl:with-param name="length" select="$cellProp/@fo:padding"/>
								</xsl:call-template>
							</xsl:variable>
							<w:top w:w="{$padding}" w:type="{$type}"/>
							<w:left w:w="{$padding}" w:type="{$type}"/>
							<w:bottom w:w="{$padding}" w:type="{$type}"/>
							<w:right w:w="{$padding}" w:type="{$type}"/>
						</xsl:when>
						<xsl:otherwise>
							<w:top w:type="{$type}">
								<xsl:attribute name="w:w">
									<xsl:call-template name="twips-measure">
										<xsl:with-param name="length" select="$cellProp/@fo:padding-top"/>
									</xsl:call-template>
								</xsl:attribute>
							</w:top>
							<w:left w:type="{$type}">
								<xsl:attribute name="w:w">
									<xsl:call-template name="twips-measure">
										<xsl:with-param name="length" select="$cellProp/@fo:padding-left"/>
									</xsl:call-template>
								</xsl:attribute>
							</w:left>
							<w:bottom w:type="{$type}">
								<xsl:attribute name="w:w">
									<xsl:call-template name="twips-measure">
										<xsl:with-param name="length" select="$cellProp/@fo:padding-bottom"/>
									</xsl:call-template>
								</xsl:attribute>
							</w:bottom>
							<w:right w:type="{$type}">
								<xsl:attribute name="w:w">
									<xsl:call-template name="twips-measure">
										<xsl:with-param name="length" select="$cellProp/@fo:padding-right"/>
									</xsl:call-template>
								</xsl:attribute>
							</w:right>
						</xsl:otherwise>
					</xsl:choose>
				</w:tcMar>
				<xsl:if test="$cellProp/@style:writing-mode">
					<xsl:choose>
						<xsl:when test="$cellProp[@style:writing-mode = 'tb-rl']">
							<w:textDirection w:val="tbRl"/>
						</xsl:when>
						<xsl:when test="$cellProp[@style:writing-mode = 'lr-tb']">
							<w:textDirection w:val="lrTb"/>
						</xsl:when>
					</xsl:choose>
				</xsl:if>
				<xsl:if test="$cellProp[@style:vertical-align and @style:vertical-align!='']">
					<xsl:choose>
						<xsl:when test="$cellProp/@style:vertical-align = 'middle'">
							<w:vAlign w:val="center"/>
						</xsl:when>
						<xsl:otherwise>
							<w:vAlign w:val="{$cellProp/@style:vertical-align}"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:if>

			</w:tcPr>
			<xsl:choose>
				<xsl:when test="not(child::table:table) and $merge &lt; 2">
					<xsl:apply-templates/>
					<!-- must precede a w:tc, otherwise it crashes. Xml schema validation does not check this. -->
				</xsl:when>
				<xsl:otherwise>
					<w:p/>
				</xsl:otherwise>
			</xsl:choose>
		</w:tc>
	</xsl:template>

	<!-- text and spaces -->
	<xsl:template match="text()|text:s" mode="paragraph">
		<w:r>
			<xsl:call-template name="InsertRunProperties"/>
			<xsl:apply-templates select="." mode="text"/>
		</w:r>
	</xsl:template>

	<!-- Inserts the Run properties -->
	<xsl:template name="InsertRunProperties">
		<!-- apply text properties if needed -->
		<xsl:if test="ancestor::text:span">
			<w:rPr>
				<xsl:call-template name="InsertRunStyle"/>
				<xsl:call-template name="OverrideToggleProperties">
					<xsl:with-param name="styleName" select="ancestor::text:span[1]/@text:style-name"/>
				</xsl:call-template>
			</w:rPr>
		</xsl:if>
	</xsl:template>

	<!-- Inserts the style of a run -->
	<xsl:template name="InsertRunStyle">
		<xsl:variable name="prefixedStyleName">
			<xsl:call-template name="GetPrefixedStyleName">
				<xsl:with-param name="styleName" select="ancestor::text:span[1]/@text:style-name"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:if test="$prefixedStyleName!=''">
			<w:rStyle w:val="{$prefixedStyleName}"/>
		</xsl:if>
	</xsl:template>

	<!-- Overrides toggle properties -->
	<xsl:template name="OverrideToggleProperties">
		<xsl:param name="styleName"/>
		<xsl:choose>
			<xsl:when test="key('automatic-styles',$styleName)">
				<!-- recursive call on parent style (not very clean) -->
				<xsl:if test="key('automatic-styles',$styleName)/@style:parent-style-name">
					<xsl:call-template name="OverrideToggleProperties">
						<xsl:with-param name="styleName"
							select="key('automatic-styles',$styleName)/@style:parent-style-name"/>
					</xsl:call-template>
				</xsl:if>
				<xsl:apply-templates select="key('automatic-styles',$styleName)/style:text-properties"
					mode="rPr">
					<xsl:with-param name="onlyToggle" select="'true'"/>
				</xsl:apply-templates>
			</xsl:when>
			<xsl:otherwise>
				<xsl:for-each select="document('styles.xml')">
					<!-- recursive call on parent style (not very clean) -->
					<xsl:if test="key('styles',$styleName)/@style:parent-style-name">
						<xsl:call-template name="OverrideToggleProperties">
							<xsl:with-param name="styleName"
								select="key('styles',$styleName)/@style:parent-style-name"/>
						</xsl:call-template>
					</xsl:if>
					<xsl:apply-templates select="key('styles',$styleName)/style:text-properties" mode="rPr">
						<xsl:with-param name="onlyToggle" select="'true'"/>
					</xsl:apply-templates>
				</xsl:for-each>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- spaces (within a text flow) -->
	<xsl:template match="text:s" mode="text">
		<w:t xml:space="preserve"><xsl:call-template name="extra-spaces"><xsl:with-param name="spaces" select="@text:c"/></xsl:call-template></w:t>
	</xsl:template>

	<!-- simple text (within a text flow) -->
	<xsl:template match="text()" mode="text">

		<xsl:choose>
			<xsl:when test="preceding-sibling::text:tab">
				<w:t xml:space="preserve"><xsl:value-of select="."/></w:t>
			</xsl:when>
			<xsl:when test="not(following-sibling::text:tab)">
				<w:t xml:space="preserve"><xsl:value-of select="."/></w:t>
			</xsl:when>

			<xsl:otherwise> </xsl:otherwise>
		</xsl:choose>


	</xsl:template>

	<!-- tab stops -->
	<xsl:template match="text:tab-stop" mode="paragraph">
		<w:r>
			<w:tab/>
			<w:t/>
		</w:r>
	</xsl:template>

	<!-- tabs -->
	<xsl:template match="text:tab" mode="paragraph">
		<w:r>
			<w:tab/>
		</w:r>
	</xsl:template>

	<!-- line breaks -->
	<xsl:template match="text:line-break" mode="paragraph">
		<w:r>
			<w:br/>
			<w:t/>
		</w:r>
	</xsl:template>

	<!-- line breaks (within the text flow) -->
	<xsl:template match="text:line-break" mode="text">
		<w:br/>
	</xsl:template>

	<!-- notes (footnotes or endnotes) -->
	<xsl:template match="text:note" mode="paragraph">
		<w:r>
			<w:rPr>
				<w:rStyle w:val="{concat(@text:note-class, 'Reference')}"/>
			</w:rPr>
			<xsl:apply-templates select="." mode="text"/>
		</w:r>
	</xsl:template>

	<!-- footnotes -->
	<xsl:template match="text:note[@text:note-class='footnote']" mode="text">
		<w:footnoteReference>
			<xsl:attribute name="w:id">
				<xsl:call-template name="GenerateId">
					<xsl:with-param name="node" select="."/>
					<xsl:with-param name="nodetype" select="@text:note-class"/>
				</xsl:call-template>
			</xsl:attribute>
		</w:footnoteReference>
		<xsl:if test="text:note-citation/@text:label">
			<w:t>
				<xsl:value-of select="text:note-citation"/>
			</w:t>
		</xsl:if>
	</xsl:template>

	<!-- endnotes -->
	<xsl:template match="text:note[@text:note-class='endnote']" mode="text">
		<w:endnoteReference>
			<xsl:attribute name="w:id">
				<xsl:call-template name="GenerateId">
					<xsl:with-param name="node" select="."/>
					<xsl:with-param name="nodetype" select="@text:note-class"/>
				</xsl:call-template>
			</xsl:attribute>
		</w:endnoteReference>
		<xsl:if test="text:note-citation/@text:label">
			<w:t>
				<xsl:value-of select="text:note-citation"/>
			</w:t>
		</xsl:if>
	</xsl:template>

	<!-- alphabetical indexes -->
	<xsl:template match="text:alphabetical-index-mark-end" mode="paragraph">
    	<w:r>
      		<w:fldChar w:fldCharType="begin"/>
    	</w:r>
    	<w:r>
      		<w:instrText xml:space="preserve"> XE "</w:instrText>
    	</w:r>
    	<w:r>
      		<w:instrText>        
        			<xsl:for-each select="preceding-sibling::node()">
         				 <xsl:choose>
           					 <xsl:when test="self::text()">
              						  <xsl:value-of select="."/>
            					</xsl:when>        
           					 <xsl:otherwise>
                						  <xsl:apply-templates select="."></xsl:apply-templates> 
            
            					</xsl:otherwise>
          				</xsl:choose>     
        			</xsl:for-each>
      		</w:instrText>        
    	</w:r>
   	 <w:r>
    	  	<w:instrText xml:space="preserve">" </w:instrText>
   	 </w:r>
    	<w:r>
      	<w:fldChar w:fldCharType="end"/>
    	</w:r>
   	 <!-- <xsl:apply-templates select="text:s" mode="text"></xsl:apply-templates> -->
  	</xsl:template>
 	<!-- spaces -->
  	<xsl:template match="text:s">  
    		  <xsl:call-template name="extra-spaces"><xsl:with-param name="spaces" select="@text:c"/></xsl:call-template>
 	 </xsl:template> 
	<!-- sequences -->
	<xsl:template match="text:sequence" mode="paragraph">
		<xsl:variable name="id">
			<xsl:value-of select="number(count(preceding::text:sequence)+count( key('headers','')))+1"/>
			<!--<xsl:value-of select="count( key('headers',''))"/>-->
		</xsl:variable>
		<w:fldSimple>
			<xsl:variable name="numType">
				<xsl:choose>
					<xsl:when test="@style:num-format = 'i'">\* roman</xsl:when>
					<xsl:when test="@style:num-format = 'I'">\* Roman</xsl:when>
					<xsl:when test="@style:num-format = 'a'">\* alphabetic</xsl:when>
					<xsl:when test="@style:num-format = 'A'">\* ALPHABETIC</xsl:when>
					<xsl:otherwise>\* arabic</xsl:otherwise>
				</xsl:choose>
			</xsl:variable>
			<xsl:attribute name="w:instr">
				<xsl:value-of select="concat('SEQ ', @text:name,' ', $numType)"/>
			</xsl:attribute>
			<w:bookmarkStart w:id="{$id}" w:name="{concat('_Toc',$id)}"/>
			<w:r>
				<w:t>
					<xsl:value-of select="."/>
				</w:t>
			</w:r>
			<w:bookmarkEnd w:id="{$id}"/>
		</w:fldSimple>
	</xsl:template>

	<!-- sections -->
	<xsl:template match="text:section">
		<xsl:choose>
			<xsl:when test="@text:display='none'"> </xsl:when>
			<xsl:otherwise>
				<xsl:apply-templates/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- Extra spaces management, courtesy of J. David Eisenberg -->
	<xsl:variable name="spaces" xml:space="preserve">                                       </xsl:variable>

	<xsl:template name="extra-spaces">
		<xsl:param name="spaces"/>
		<xsl:choose>
			<xsl:when test="$spaces">
				<xsl:call-template name="insert-spaces">
					<xsl:with-param name="n" select="$spaces"/>
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:text> </xsl:text>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<xsl:template name="insert-spaces">
		<xsl:param name="n"/>
		<xsl:choose>
			<xsl:when test="$n &lt;= string-length($spaces)">
				<xsl:value-of select="substring($spaces, 1, $n)" xml:space="preserve"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$spaces"/>
				<xsl:call-template name="insert-spaces">
					<xsl:with-param name="n">
						<xsl:value-of select="$n - string-length($spaces)"/>
					</xsl:with-param>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- ignored -->
	<xsl:template match="text()"/>
	<xsl:template match="text:tracked-changes"/>

</xsl:stylesheet>

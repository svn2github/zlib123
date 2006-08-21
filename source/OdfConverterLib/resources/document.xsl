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
	exclude-result-prefixes="office text table fo style draw xlink">

	<xsl:strip-space elements="*"/>
	<xsl:preserve-space elements="text:p"/>
	<xsl:preserve-space elements="text:span"/>

	<xsl:key name="style"
		match="office:automatic-styles/style:style|office:automatic-styles/style:style"
		use="@style:name"/>
	<xsl:key name="hyperlinks" match="text:a" use="''"/>
	<xsl:key name="images" match="draw:frame[not(./draw:object-ole) and ./draw:image/@xlink:href]"
		use="''"/>

	<xsl:variable name="type">dxa</xsl:variable>

	<xsl:template name="document">
		<w:document>
			<xsl:apply-templates select="document('content.xml')/office:document-content"/>
		</w:document>
	</xsl:template>

	<!--xsl:template match="text()" mode="document">
	</xsl:template-->

	<xsl:template match="office:body">
		<w:body>
			<xsl:apply-templates/>
			<xsl:call-template name="headerFooter"/>

		</w:body>
	</xsl:template>

	<xsl:template match="office:text">
		<xsl:apply-templates/>
	</xsl:template>

	<!-- headings -->

	<xsl:template match="text:h">
		<w:p>
			<w:pPr>
				<xsl:variable name="headerStyle">
					<xsl:call-template name="headerStyleName">
						<xsl:with-param name="styleName" select="@text:style-name"/>
					</xsl:call-template>
				</xsl:variable>

				<xsl:variable name="nameStyle">
					<xsl:value-of select="concat($headerStyle,@text:style-name)"/>
				</xsl:variable>

				<w:pStyle>
					<xsl:attribute name="w:val">
						<xsl:value-of select="$nameStyle"/>
					</xsl:attribute>
				</w:pStyle>
				<!-- <w:pStyle w:val="{@text:style-name}"/> -->
				<xsl:choose>
					<xsl:when test="not(@text:outline-level)">
						<w:outlineLvl w:val="0"/>
					</xsl:when>
					<xsl:when test="@text:outline-level &lt;= 9">
						<xsl:if
							test="document('styles.xml')//office:document-styles/office:styles/text:outline-style/text:outline-level-style/@style:num-format !=''">
							<w:numPr>
								<w:ilvl w:val="{@text:outline-level - 1}"/>
								<w:numId w:val="1"/>
							</w:numPr>
						</xsl:if>
						<w:outlineLvl w:val="{@text:outline-level}"/>
					</xsl:when>
					<xsl:otherwise>
						<w:outlineLvl w:val="9"/>
					</xsl:otherwise>
				</xsl:choose>
			</w:pPr>
			<xsl:variable name="id">
				<xsl:number/>
			</xsl:variable>

			<w:bookmarkStart w:id="{$id}" w:name="{concat('_Toc',$id)}"/>
			<xsl:apply-templates mode="paragraph"/>
			<w:bookmarkEnd w:id="{$id}"/>

		</w:p>
	</xsl:template>

	<!-- paragraphs -->

	<xsl:template match="text:p">
		<xsl:param name="level" select="0"/>
		<xsl:message terminate="no">progress:text:p</xsl:message>

		<xsl:choose>
			<xsl:when test="$level = 0">
				<w:p>
					<xsl:choose>
						<!-- we are in a footnote or endnote -->
						<xsl:when test="parent::text:note-body">
							<xsl:variable name="note" select="ancestor::text:note"/>
							<w:pPr>
								<xsl:choose>
									<xsl:when
										test="count(draw:frame) = 1 and count(child::text()) = 0">
										<w:pStyle w:val="{draw:frame/@draw:style-name}"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:variable name="headerStyle">
											<xsl:call-template name="headerStyleName"/>
										</xsl:variable>
										<xsl:variable name="nameStyle">
											<xsl:value-of
												select="concat($headerStyle,@text:style-name)"/>
										</xsl:variable>
										<w:pStyle>
											<xsl:attribute name="w:val">
												<xsl:value-of select="$nameStyle"/>
											</xsl:attribute>
										</w:pStyle>
										<!-- <w:pStyle w:val="{@text:style-name}"/> -->
										<!--w:pStyle w:val="{concat($note/@text:note-class, 'Text')}"/-->
									</xsl:otherwise>
								</xsl:choose>
							</w:pPr>

							<xsl:if test="position() = 1 and parent::text:note-body">

								<!-- Include the mark to the first paragraph -->
								<xsl:apply-templates select="../../text:note-citation"/>

							</xsl:if>
							<xsl:apply-templates mode="paragraph"/>
						</xsl:when>
						<!-- we are in table of contents -->
						<xsl:when test="parent::text:index-body and ancestor::text:table-of-content">
							<w:pPr>
								<xsl:variable name="headerStyle">
									<xsl:call-template name="headerStyleName">
										<xsl:with-param name="styleName" select="@text:style-name"/>
									</xsl:call-template>
								</xsl:variable>

								<xsl:variable name="nameStyle">
									<xsl:value-of select="concat($headerStyle,@text:style-name)"/>
								</xsl:variable>

								<w:pStyle>
									<xsl:attribute name="w:val">
										<xsl:value-of select="$nameStyle"/>
									</xsl:attribute>
								</w:pStyle>
							</w:pPr>
							<xsl:variable name="num">
								<xsl:number/>
							</xsl:variable>
							<xsl:if test="$num=1">
								<w:r>
									<w:fldChar w:fldCharType="begin"/>
								</w:r>
								<w:r>
									<w:instrText xml:space="preserve"> TOC \o "1-<xsl:choose><xsl:when test="parent::text:index-body/preceding-sibling::text:table-of-content-source/@text:outline-level=10">9</xsl:when><xsl:otherwise><xsl:value-of select="parent::text:index-body/preceding-sibling::text:table-of-content-source/@text:outline-level"/></xsl:otherwise></xsl:choose>"<xsl:if test="text:a"> \h </xsl:if></w:instrText>
								</w:r>
								<w:r>
									<w:fldChar w:fldCharType="separate"/>
								</w:r>
							</xsl:if>
							<xsl:choose>
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
						</xsl:when>
						<!-- main scenario -->
						<xsl:otherwise>
							<w:pPr>
								<xsl:variable name="headerStyle">
									<xsl:call-template name="headerStyleName">
										<xsl:with-param name="styleName" select="@text:style-name"/>
									</xsl:call-template>
								</xsl:variable>
								<xsl:variable name="nameStyle">
									<xsl:value-of select="concat($headerStyle,@text:style-name)"/>
								</xsl:variable>

								<w:pStyle>
									<xsl:attribute name="w:val">
										<xsl:value-of select="$nameStyle"/>
									</xsl:attribute>
								</w:pStyle>
								<xsl:call-template name="indent">
									<xsl:with-param name="level" select="$level"/>
								</xsl:call-template>
							</w:pPr>
							<xsl:choose>
								<xsl:when
									test="child::draw:frame and not(parent::draw:text-box) and child::draw:frame/child::draw:text-box">
									<xsl:apply-templates select="draw:frame"/>
								</xsl:when>
								<xsl:otherwise>
									<xsl:apply-templates mode="paragraph"/>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:otherwise>
					</xsl:choose>
					<!-- selfstanding image before paragraph-->
					<xsl:if
						test="name(preceding-sibling::*[1]) = 'draw:frame' and preceding-sibling::*[1]/draw:image">
						<xsl:apply-templates select="preceding-sibling::*[1]" mode="paragraph"/>
					</xsl:if>
				</w:p>
			</xsl:when>
			<xsl:otherwise>
				<!-- We are in a list -->
				<w:p>
					<w:pPr>
						<xsl:variable name="headerStyle">
							<xsl:call-template name="headerStyleName">
								<xsl:with-param name="styleName" select="@text:style-name"/>
							</xsl:call-template>
						</xsl:variable>

						<xsl:variable name="nameStyle">
							<xsl:value-of select="concat($headerStyle,@text:style-name)"/>
						</xsl:variable>

						<w:pStyle>
							<xsl:attribute name="w:val">
								<xsl:value-of select="$nameStyle"/>
							</xsl:attribute>
						</w:pStyle>
						<!-- <w:pStyle w:val="{@text:style-name}"/> -->

						<xsl:call-template name="indent">
							<xsl:with-param name="level" select="$level"/>
						</xsl:call-template>
					</w:pPr>
					<xsl:apply-templates mode="paragraph"/>
				</w:p>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<xsl:template match="text:note-citation">
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

			<!-- extra space between footnote/endnote mark and text-->
			<w:t xml:space="preserve"> </w:t>
		</w:r>
	</xsl:template>

	<!-- links -->

	<!--xsl:template match="text:a" mode="paragraph">
		<w:hyperlink r:id='{generate-id()}'>
			<w:r>
				<w:rPr>
					<w:rStyle w:val="{@text:style-name}"/>
				</w:rPr>
				<w:t>
					<xsl:apply-templates mode="hyperlink"/>
				</w:t>
			</w:r>
		</w:hyperlink>
	</xsl:template-->

	<xsl:template match="text:a" mode="paragraph">
		<xsl:choose>
			<xsl:when test="ancestor::text:index-body">
				<xsl:variable name="num" select="count(parent::*/preceding-sibling::*)"> </xsl:variable>
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

	<!-- TODO : find the best way to avoid code duplication -->
	<xsl:template match="text:tab-stop" mode="hyperlink">
		<w:tab/>
	</xsl:template>

	<xsl:template match="text:line-break" mode="hyperlink">
		<w:br/>
	</xsl:template>

	<xsl:template match="text()" mode="hyperlink">
		<xsl:value-of select="."/>
	</xsl:template>

	<xsl:template match="text:s" mode="hyperlink">
		<xsl:call-template name="extra-spaces"/>
	</xsl:template>

	<xsl:template match="text:span" mode="hyperlink">
		<w:rPr>
			<xsl:variable name="headerStyle">
				<xsl:call-template name="headerStyleName">
					<xsl:with-param name="styleName" select="@text:style-name"/>
				</xsl:call-template>
			</xsl:variable>

			<xsl:variable name="nameStyle">
				<xsl:value-of select="concat($headerStyle,@text:style-name)"/>
			</xsl:variable>

			<w:rStyle>
				<xsl:attribute name="w:val">
					<xsl:value-of select="$nameStyle"/>
				</xsl:attribute>
			</w:rStyle>
		</w:rPr>
		<w:t>
			<xsl:attribute name="xml:space">preserve</xsl:attribute>
			<xsl:apply-templates mode="hyperlink"/>
		</w:t>
	</xsl:template>

	<xsl:template match="draw:text-box">
		<w:r>
			<w:rPr>
				<w:rStyle w:val="{parent::draw:frame/@draw:style-name}"/>
			</w:rPr>
			<w:pict>
				<v:shapetype/>
				<v:shape type="#_x0000_t202">
					<xsl:variable name="frameW">
						<xsl:call-template name="point-measure">
							<xsl:with-param name="length"
								select="parent::draw:frame/@svg:width|parent::draw:frame/@fo:min-width"
							/>
						</xsl:call-template>
					</xsl:variable>
					<xsl:variable name="frameH">
						<xsl:call-template name="point-measure">
							<xsl:with-param name="length" select="@fo:min-height|parent::draw:frame/@svg:height"/>
						</xsl:call-template>
					</xsl:variable>
					<xsl:variable name="marginL">
						<xsl:call-template name="point-measure">
							<xsl:with-param name="length" select="parent::draw:frame/@svg:x"/>
						</xsl:call-template>
					</xsl:variable>
					<xsl:variable name="marginT">
						<xsl:call-template name="point-measure">
							<xsl:with-param name="length" select="parent::draw:frame/@svg:y"/>
						</xsl:call-template>
					</xsl:variable>
					<xsl:variable name="zIndex">
						<xsl:value-of select="parent::draw:frame/@draw:z-index"/>
					</xsl:variable>
					<xsl:attribute name="style">
						<xsl:value-of select="'position:absolute;'"/>
						<xsl:value-of select="concat('width:',$frameW,'pt;')"/>
						<xsl:value-of select="concat('height:',$frameH,'pt;')"/>
						<xsl:value-of select="concat('z-index:', $zIndex, ';')"/>
						<xsl:value-of select="concat('margin-left:',$marginL,'pt;')"/>
						<xsl:value-of select="concat('margin-top:',$marginT,'pt;')"/>
						<xsl:if
							test="key('style', parent::draw:frame/@draw:style-name)/style:graphic-properties/@style:horizontal-pos">
							<xsl:value-of
								select="concat('mso-position-horizontal:', key('style', parent::draw:frame/@draw:style-name)/style:graphic-properties/@style:horizontal-pos,';')"
							/>
						</xsl:if>
					</xsl:attribute>
					<xsl:if
						test="key('style', parent::draw:frame/@draw:style-name)/style:graphic-properties/@fo:background-color">
						<xsl:attribute name="fillcolor">
							<xsl:value-of
								select="key('style', parent::draw:frame/@draw:style-name)/style:graphic-properties/@fo:background-color"
							/>
						</xsl:attribute>
					</xsl:if>
					<v:textbox>
						<xsl:attribute name="style">
							<xsl:if test="@fo:min-height">
								<xsl:value-of select="'mso-fit-shape-to-text:t'"/>
							</xsl:if>
						</xsl:attribute>
						<w:txbxContent>
							<xsl:for-each select="child::node()">
								<xsl:apply-templates select="."/>
							</xsl:for-each>
						</w:txbxContent>
					</v:textbox>
				</v:shape>
			</w:pict>
		</w:r>
	</xsl:template>


	<!-- @TODO  positioning text-boxes -->
	<xsl:template match="draw:frame">
		<xsl:choose>
			<xsl:when test="not(parent::text:p)">
				<w:p>
					<xsl:for-each select="descendant::draw:text-box">
						<xsl:apply-templates select="."/>
					</xsl:for-each>
				</w:p>
			</xsl:when>
			<xsl:otherwise>
				<xsl:for-each select="descendant::draw:text-box">
					<xsl:apply-templates select="."/>
				</xsl:for-each>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- lists -->

	<xsl:template match="text:list">
		<xsl:param name="level" select="-1"/>
		<xsl:apply-templates>
			<xsl:with-param name="level" select="$level+1"/>
		</xsl:apply-templates>
	</xsl:template>

	<xsl:template match="text:list-item|text:list-header">
		<xsl:param name="level"/>
		<xsl:if test="self::text:list-item">
			<xsl:choose>
				<xsl:when test="*[1][self::text:p]">
					<w:p>
						<w:pPr>
							<xsl:variable name="headerStyle">
								<xsl:call-template name="headerStyleName">
									<xsl:with-param name="styleName"
										select="*[1][self::text:p]/@text:style-name"/>
								</xsl:call-template>
							</xsl:variable>

							<xsl:variable name="nameStyle">
								<xsl:value-of
									select="concat($headerStyle,*[1][self::text:p]/@text:style-name)"
								/>
							</xsl:variable>

							<w:pStyle>
								<xsl:attribute name="w:val">
									<xsl:value-of select="$nameStyle"/>
								</xsl:attribute>
							</w:pStyle>
							<!--<w:pStyle w:val="{*[1][self::text:p]/@text:style-name}"/> -->
							<w:numPr>
								<w:ilvl w:val="{$level}"/>
								<w:numId>
									<xsl:attribute name="w:val">
										<xsl:call-template name="numberingId">
											<xsl:with-param name="styleName"
												select="ancestor::text:list/@text:style-name"/>
										</xsl:call-template>
									</xsl:attribute>
								</w:numId>
							</w:numPr>

							<!-- override abstract num indent  if paragraph has margin defined -->
							<xsl:call-template name="indent">
								<xsl:with-param name="level" select="$level"/>
							</xsl:call-template>
						</w:pPr>

						<!--footnote or endnote - Include the mark to the first paragraph only when first child of text:note-body is not paragraph -->
						<xsl:if
							test="ancestor::text:note and not(ancestor::text:note-body/child::*[position()=1]/text:p) and position() = 1">
							<xsl:apply-templates select="ancestor::text:note/text:note-citation"/>
						</xsl:if>

						<!-- first paragraph -->
						<xsl:apply-templates select="*[1][self::text:p]" mode="paragraph"/>
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
		</xsl:if>
		<xsl:if test="self::text:list-header">
			<xsl:apply-templates>
				<xsl:with-param name="level" select="$level"/>
			</xsl:apply-templates>
		</xsl:if>
	</xsl:template>

	<xsl:template name="indent">
		<xsl:param name="level" select="0"/>
		<xsl:variable name="styleName">
			<xsl:call-template name="styleName"/>
		</xsl:variable>
		<xsl:variable name="parentStyleName"
			select="key('style',$styleName)/@style:parent-style-name"/>
		<xsl:variable name="listStyleName" select="key('style',$styleName)/@style:list-style-name"/>
		<xsl:variable name="paragraphMargin">
			<xsl:choose>
				<xsl:when test="key('style',$styleName)/paragraph-properties/@fo:left-margin">
					<xsl:call-template name="twips-measure">
						<xsl:with-param name="length"
							select="key('style',$styleName)/style:paragraph-properties/@fo:left-margin"
						/>
					</xsl:call-template>
				</xsl:when>
				<xsl:when
					test="document('styles.xml')//style:style[@style:name = $parentStyleName]/style:paragraph-properties/@fo:margin-left">
					<xsl:call-template name="twips-measure">
						<xsl:with-param name="length"
							select="document('styles.xml')//style:style[@style:name = $parentStyleName]/style:paragraph-properties/@fo:margin-left"
						/>
					</xsl:call-template>
				</xsl:when>
				<xsl:otherwise>0</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:variable name="minLabelWidthTwip">
			<xsl:call-template name="twips-measure">
				<xsl:with-param name="length"
					select="document('content.xml')//text:list-style[@style:name = $listStyleName]/*[@text:level = $level+1]/style:list-level-properties/@text:min-label-width"
				/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name="spaceBeforeTwip">
			<xsl:call-template name="twips-measure">
				<xsl:with-param name="length"
					select="document('content.xml')//text:list-style[@style:name = $listStyleName]/*[@text:level = $level+1]/style:list-level-properties/@text:space-before"
				/>
			</xsl:call-template>
		</xsl:variable>

		<xsl:if test="parent::text:list-header|self::text:p">
			<w:ind>
				<xsl:attribute name="w:left">
					<xsl:value-of select="$minLabelWidthTwip + $paragraphMargin  + $spaceBeforeTwip"
					/>
				</xsl:attribute>
				<xsl:if test="not(ancestor-or-self::text:list)">
					<xsl:attribute name="w:hanging">
						<xsl:value-of select="$spaceBeforeTwip"/>
					</xsl:attribute>
				</xsl:if>
			</w:ind>
		</xsl:if>
	</xsl:template>

	<xsl:template name="styleName">
		<xsl:if test="self::text:list-item">
			<xsl:value-of select="*[1][self::text:p]/@text:style-name"/>
		</xsl:if>
		<xsl:if test="parent::text:list-header|self::text:p">
			<xsl:value-of select="@text:style-name"/>
		</xsl:if>
	</xsl:template>

	<!-- table of contents -->

	<xsl:template name="tableContent">
		<xsl:param name="num"/>
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
						<xsl:for-each select="child::text()[position() &lt; last()]">
							<xsl:value-of select="."/>
						</xsl:for-each>
					</w:t>
					<!--<xsl:apply-templates select="child::text()[1]" mode="text"/>-->
				</xsl:otherwise>
			</xsl:choose>
		</w:r>
		<xsl:apply-templates select="text:tab|text:a/text:tab" mode="paragraph"/>
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
		<w:r>
			<w:rPr/>
			<w:fldChar w:fldCharType="end"/>
		</w:r>
	</xsl:template>

	<xsl:template match="text:table-of-content">
		<w:sdt>
			<w:sdtPr>
				<w:docPartObj>
					<w:docPartType>
						<xsl:attribute name="w:val">
							<xsl:value-of select="'Table of Contents'"/>
						</xsl:attribute>
					</w:docPartType>
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

	<xsl:template name="subtable">
		<xsl:param name="node"/>
		<xsl:for-each select="$node/table:table-cell">
			<xsl:call-template name="table-cell"/>
		</xsl:for-each>
	</xsl:template>

	<xsl:template name="merged-rows">
		<xsl:param name="i" select="0"/>
		<xsl:param name="iterator"/>
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
								<xsl:with-param name="node"
									select="table:table/child::table:table-row[$iterator]"/>
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

	<xsl:template match="table:table">
		<xsl:variable name="styleName">
			<xsl:value-of select="@table:style-name"/>
		</xsl:variable>
		<xsl:if
			test="//office:automatic-styles/style:style[@style:name=$styleName]/style:table-properties/@fo:break-before='page'">
			<w:p>
				<w:r>
					<w:br w:type="page"/>
				</w:r>
			</w:p>
		</xsl:if>
		<w:tbl>
			<w:tblPr>
				<w:tblStyle w:val="{@table:style-name}"/>
				<xsl:variable name="tableProp"
					select="key('style', @table:style-name)/style:table-properties"/>
				<w:tblW w:type="{$type}">
					<xsl:attribute name="w:w">
						<xsl:call-template name="twips-measure">
							<xsl:with-param name="length"
								select="key('style', @table:style-name)/style:table-properties/@style:width"
							/>
						</xsl:call-template>
					</xsl:attribute>
				</w:tblW>
				<xsl:if test="key('style', @table:style-name)/style:table-properties/@table:align">
					<xsl:choose>
						<xsl:when
							test="key('style', @table:style-name)/style:table-properties/@table:align = 'margins'">
							<w:jc w:val="left"/>
							<!--User agents that do not support the "margins" value, may treat this value as "left".-->
						</xsl:when>
						<xsl:otherwise>
							<w:jc
								w:val="{key('style', @table:style-name)/style:table-properties/@table:align}"
							/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:if>

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
				select="table:table-rows|table:table-header-rows|table:table-row|table:table-header-row"
			/>
		</w:tbl>
	</xsl:template>

	<!-- this fragment count table:table-column so the max number of cols was 63 -->
	<xsl:template name="colCount">
		<xsl:param name="nodeList"/>
		<xsl:choose>
			<xsl:when test="$nodeList">
				<xsl:variable name="recursive_result">
					<xsl:call-template name="colCount">
						<xsl:with-param name="productList" select="$nodeList[position() > 1]"/>
					</xsl:call-template>
				</xsl:variable>
				<xsl:choose>
					<xsl:when test="$nodeList[1]/@table:number-columns-repeated ">
						<xsl:value-of
							select="number($nodeList[1]/@table:number-columns-repeated) + $recursive_result"
						/>
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

	<xsl:template match="table:table-column">
		<xsl:param name="repeat" select="1"/>

		<xsl:variable name="columnNumber">
			<xsl:call-template name="colCount">
				<xsl:with-param name="nodeList" select="preceding-sibling::table:table-column"/>
			</xsl:call-template>
		</xsl:variable>

		<xsl:if test="$columnNumber &lt; 63">
			<!-- relative width not supported yet -->
			<w:gridCol>
				<xsl:attribute name="w:w">
					<xsl:call-template name="twips-measure">
						<xsl:with-param name="length"
							select="key('style', @table:style-name)/style:table-column-properties/@style:column-width"
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

	<xsl:template match="table:table-rows|table:table-header-rows">
		<xsl:apply-templates select="table:table-row|table:table-header-row"/>
	</xsl:template>

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
					<w:trPr>
						<xsl:if test="name(parent::*) = 'table:table-header-rows'">
							<w:tblHeader/>
						</xsl:if>
						<xsl:if test="@table:style-name">
							<xsl:variable name="rowStyle" select="@table:style-name"/>
							<xsl:variable name="widthType"
								select="key('style', $rowStyle)/style:table-row-properties"/>

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
						<xsl:if
							test="key('style', @table:style-name)/style:table-row-properties/@style:keep-together = 'false'">
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



	<!-- this fragment sum up width of a column in case of merged columns -->
	<xsl:template name="cellWidthCount">
		<xsl:param name="cellSpanned"/>
		<xsl:param name="col"/>
		<xsl:variable name="table" select="ancestor::table:table[1]"/>
		<xsl:choose>
			<xsl:when test="$cellSpanned &gt; 1">
				<xsl:variable name="recursive_result">
					<xsl:call-template name="cellWidthCount">
						<xsl:with-param name="cellSpanned" select="$cellSpanned -1"/>
						<xsl:with-param name="col" select="$col+1"/>
					</xsl:call-template>
				</xsl:variable>
				<xsl:variable name="colStyle"
					select="$table/table:table-column[position() = $col]/@table:style-name"/>
				<xsl:variable name="width">
					<xsl:call-template name="twips-measure">
						<xsl:with-param name="length"
							select="key('style', $colStyle)/style:table-column-properties/@style:column-width"
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
						select="key('style', $colStyle)/style:table-column-properties/@style:column-width"
					/>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

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
			<xsl:call-template name="cellWidthCount">
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

	<xsl:template match="table:table-cell" name="table-cell">
		<xsl:param name="colsNumber"/>
		<xsl:param name="grid" select="0"/>
		<xsl:param name="merge" select="0"/>
		<w:tc>
			<w:tcPr>
				<!-- point on the cell style properties -->
				<xsl:variable name="cellProp"
					select="key('style', @table:style-name)/style:table-cell-properties"/>
				<xsl:variable name="tableStyle" select="substring-before(@table:style-name, '.')"/>
				<xsl:variable name="rowStyle" select="../@table:style-name"/>
				<xsl:variable name="tableProp"
					select="key('style', $tableStyle)/style:table-properties"/>
				<xsl:variable name="rowProp"
					select="key('style', $rowStyle)/style:table-row-properties"/>

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
									<xsl:with-param name="length"
										select="substring-before($border,' ')"/>
								</xsl:call-template>
							</xsl:variable>
							<w:top w:val="single" w:color="{$border-color}" w:sz="{$border-size}"/>
							<w:left w:val="single" w:color="{$border-color}" w:sz="{$border-size}"/>
							<w:bottom w:val="single" w:color="{$border-color}" w:sz="{$border-size}"/>
							<w:right w:val="single" w:color="{$border-color}" w:sz="{$border-size}"
							/>
						</xsl:when>

						<xsl:otherwise>
							<xsl:if test="$cellProp[@fo:border-top and @fo:border-top != 'none']">
								<xsl:variable name="border" select="$cellProp/@fo:border-top"/>
								<w:top w:val="single" w:color="{substring-after($border, '#')}">
									<xsl:attribute name="w:sz">
										<xsl:call-template name="eightspoint-measure">
											<xsl:with-param name="length"
												select="substring-before($border, ' ')"/>
										</xsl:call-template>
									</xsl:attribute>
								</w:top>
							</xsl:if>
							<xsl:if test="$cellProp[@fo:border-left and @fo:border-left != 'none']">
								<xsl:variable name="border" select="$cellProp/@fo:border-left"/>
								<w:left w:val="single" w:color="{substring-after($border, '#')}">
									<xsl:attribute name="w:sz">
										<xsl:call-template name="eightspoint-measure">
											<xsl:with-param name="length"
												select="substring-before($border, ' ')"/>
										</xsl:call-template>
									</xsl:attribute>
								</w:left>
							</xsl:if>
							<xsl:if
								test="$cellProp[@fo:border-bottom and @fo:border-bottom != 'none']">
								<xsl:variable name="border" select="$cellProp/@fo:border-bottom"/>
								<w:bottom w:val="single" w:color="{substring-after($border, '#')}">
									<xsl:attribute name="w:sz">
										<xsl:call-template name="eightspoint-measure">
											<xsl:with-param name="length"
												select="substring-before($border, ' ')"/>
										</xsl:call-template>
									</xsl:attribute>
								</w:bottom>
							</xsl:if>
							<xsl:if
								test="$cellProp[(@fo:border-right and @fo:border-right != 'none')] or (position() &lt; $colsNumber and position() = 63)">
								<xsl:variable name="border">
									<xsl:choose>
										<xsl:when
											test="position() &lt; $colsNumber and position() = 63">
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
											<xsl:with-param name="length"
												select="substring-before($border, ' ')"/>
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
										<xsl:with-param name="length"
											select="$cellProp/@fo:padding-top"/>
									</xsl:call-template>
								</xsl:attribute>
							</w:top>
							<w:left w:type="{$type}">
								<xsl:attribute name="w:w">
									<xsl:call-template name="twips-measure">
										<xsl:with-param name="length"
											select="$cellProp/@fo:padding-left"/>
									</xsl:call-template>
								</xsl:attribute>
							</w:left>
							<w:bottom w:type="{$type}">
								<xsl:attribute name="w:w">
									<xsl:call-template name="twips-measure">
										<xsl:with-param name="length"
											select="$cellProp/@fo:padding-bottom"/>
									</xsl:call-template>
								</xsl:attribute>
							</w:bottom>
							<w:right w:type="{$type}">
								<xsl:attribute name="w:w">
									<xsl:call-template name="twips-measure">
										<xsl:with-param name="length"
											select="$cellProp/@fo:padding-right"/>
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
		<xsl:param name="styleName"/>
		<w:r>
			<w:rPr>
				<xsl:apply-templates select="key('style', $styleName)/style:text-properties"
					mode="toggle"/>
			</w:rPr>
			<xsl:apply-templates select="." mode="text"/>

		</w:r>
	</xsl:template>

	<xsl:template name="text" match="text()|text:s" mode="text">

		<xsl:choose>

			<xsl:when test="name()='text:s'">
				<w:t xml:space="preserve"><xsl:call-template name="extra-spaces"><xsl:with-param name="spaces" select="@text:c"/></xsl:call-template></w:t>
			</xsl:when>
			<xsl:otherwise>
				<w:t>
					<xsl:attribute name="xml:space">preserve</xsl:attribute>
					<xsl:value-of select="."/>

				</w:t>
			</xsl:otherwise>
		</xsl:choose>


	</xsl:template>


	<xsl:template match="text:span" mode="paragraph">
		<w:r>
			<w:rPr>
				<xsl:variable name="headerStyle">
					<xsl:call-template name="headerStyleName">
						<xsl:with-param name="styleName" select="@text:style-name"/>
					</xsl:call-template>
				</xsl:variable>

				<xsl:variable name="nameStyle">
					<xsl:value-of select="concat($headerStyle,@text:style-name)"/>
				</xsl:variable>
				<!-- <w:rStyle w:val="{@text:style-name}"/> -->
				<w:rStyle>
					<xsl:attribute name="w:val">
						<xsl:value-of select="$nameStyle"/>
					</xsl:attribute>
				</w:rStyle>

				<xsl:variable name="fontSize">
					<xsl:choose>
						<xsl:when
							test="key('style',parent::*/@text:style-name)/child::style:text-properties/@fo:font-size">
							<xsl:value-of
								select="key('style',parent::*/@text:style-name)/child::style:text-properties/@fo:font-size"
							/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of
								select="document('styles.xml')/office:document-styles/office:styles/style:default-style[@style:family='paragraph']/style:text-properties/@fo:font-size"
							/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:variable>
				<xsl:apply-templates select="key('style', @text:style-name)/style:text-properties"
					mode="toggle">
					<xsl:with-param name="fontSize" select="$fontSize"/>
				</xsl:apply-templates>
			</w:rPr>
			<xsl:apply-templates mode="text"/>
		</w:r>
	</xsl:template>

	<xsl:template match="text:span[child::*]" mode="paragraph">
		<xsl:apply-templates mode="paragraph">
			<xsl:with-param name="styleName" select="@text:style-name"/>
		</xsl:apply-templates>
	</xsl:template>

	<xsl:template match="text:tab-stop" mode="paragraph">
		<w:r>
			<w:tab/>
			<w:t/>
		</w:r>
	</xsl:template>


	<xsl:template match="text:tab" mode="paragraph">
		<w:r>
			<w:tab/>
		</w:r>
	</xsl:template>


	<xsl:template match="text:line-break" mode="paragraph">
		<w:r>
			<w:br/>
			<w:t/>
		</w:r>
	</xsl:template>

	<xsl:template match="text:line-break" mode="text">
		<w:br/>
	</xsl:template>

	<!-- footnotes and endnotes -->
	<xsl:template match="text:note" mode="paragraph">
		<w:r>
			<w:rPr>
				<w:rStyle w:val="{concat(@text:note-class, 'Reference')}"/>
			</w:rPr>
			<xsl:apply-templates select="." mode="text"/>
		</w:r>
	</xsl:template>

	<xsl:template match="text:note[@text:note-class='footnote']" mode="text">
		<w:footnoteReference>
			<xsl:attribute name="w:id">
				<xsl:call-template name="noteId">
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

	<xsl:template match="text:note[@text:note-class='endnote']" mode="text">
		<w:endnoteReference>
			<xsl:attribute name="w:id">
				<xsl:call-template name="noteId">
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

	<!-- 
		Generate a decimal identifier based on the position of the current 
		footenote/endnote among all the indexed footnotes/endnotes.
	-->
	<xsl:template name="noteId">
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

	<!--<xsl:template match="text:page-number" mode="paragraph">
		<w:fldSimple w:instr=" PAGE   \* MERGEFORMAT ">
			<w:r>
				<w:rPr>
					<w:rStyle w:val="{../@text:style-name}"/>
				</w:rPr>
				<xsl:apply-templates mode="text"/>
			</w:r>
		</w:fldSimple>
	</xsl:template>-->

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

	<!-- odt section -->
	<xsl:template match="text:section">
		<xsl:choose>
			<xsl:when test="@text:display='none'"> </xsl:when>
			<xsl:otherwise>
				<xsl:apply-templates/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

</xsl:stylesheet>

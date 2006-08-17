<?xml version="1.0" encoding="UTF-8" ?>
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
	xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
	xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
	xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
	xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
	xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
	xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
	exclude-result-prefixes="office style fo text draw number">

	<xsl:variable name="asianLayoutId">1</xsl:variable>

	<xsl:template name="styles">
		<w:styles>
			<w:docDefaults>
				<!-- Default text properties -->
				<xsl:variable name="paragraphDefaultStyle"
					select="document('styles.xml')/office:document-styles/office:styles/style:default-style[@style:family='paragraph']"/>
				<w:rPrDefault>
					<xsl:apply-templates select="$paragraphDefaultStyle/style:text-properties"
						mode="styles"/>
				</w:rPrDefault>
				<!-- Default paragraph properties -->
				<w:pPrDefault>
					<xsl:apply-templates select="$paragraphDefaultStyle/style:paragraph-properties"
						mode="styles"/>
				</w:pPrDefault>
			</w:docDefaults>
			<xsl:apply-templates
				select="document('styles.xml')/office:document-styles/office:styles" mode="styles"/>
			<xsl:apply-templates
				select="document('content.xml')/office:document-content/office:automatic-styles"
				mode="styles"/>
			<xsl:apply-templates
				select="document('styles.xml')/office:document-styles/office:automatic-styles"
				mode="styles"/>
		</w:styles>
	</xsl:template>



	<xsl:template match="style:style" mode="styles">
		<w:style w:styleId="{@style:name}">

			<xsl:choose>
				<xsl:when test="@style:family = 'text' ">
					<xsl:attribute name="w:type">character</xsl:attribute>
				</xsl:when>
				<xsl:when test="@style:family = 'graphic' ">
					<xsl:attribute name="w:type">paragraph</xsl:attribute>
				</xsl:when>
				<xsl:when test="@style:family = 'section' ">
					<xsl:attribute name="w:type">paragraph</xsl:attribute>
				</xsl:when>
				<xsl:when test="@style:family = 'paragraph' ">
					<xsl:attribute name="w:type">paragraph</xsl:attribute>
				</xsl:when>

				<xsl:when test="@style:family = 'table' ">
					<xsl:attribute name="w:type">table</xsl:attribute>
				</xsl:when>
				<xsl:when test="@style:family = 'table-cell' ">
					<xsl:attribute name="w:type">table</xsl:attribute>
				</xsl:when>
				<xsl:when test="@style:family = 'table-row' ">
					<xsl:attribute name="w:type">table</xsl:attribute>
				</xsl:when>
				<xsl:when test="@style:family = 'table-column' ">
					<xsl:attribute name="w:type">table</xsl:attribute>
				</xsl:when>

				<!--
				<xsl:when test="@style:family">
					<xsl:attribute name="w:type">
						<xsl:value-of select="@style:family"/>
					</xsl:attribute>
				</xsl:when>
				-->
				<xsl:otherwise>
					<xsl:attribute name="w:type">character</xsl:attribute>
				</xsl:otherwise>
			</xsl:choose>

			<!-- Nested elements-->

			<xsl:choose>
				<xsl:when test="@style:display-name">
					<w:name w:val="{@style:display-name}"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:if test="@style:name">
						<w:name w:val="{@style:name}"/>
					</xsl:if>
				</xsl:otherwise>
			</xsl:choose>


			<xsl:if test="@style:parent-style-name">
				<w:basedOn w:val="{@style:parent-style-name}"/>
			</xsl:if>
			<xsl:if test="@style:next-style-name">
				<w:next w:val="{@style:next-style-name}"/>
			</xsl:if>

			<xsl:if test=" name(parent::*) = 'office:automatic-styles' ">
				<!-- Automatic-styles are not displayed -->
				<w:hidden/>
			</xsl:if>

			<w:qFormat/>

			<xsl:if
				test="@style:master-page-name and not(style:paragraph-properties) and not(style:paragraph-properties = '')">
				<w:pPr>
					<w:pageBreakBefore/>
				</w:pPr>
			</xsl:if>

			<xsl:apply-templates mode="styles"/>



		</w:style>
	</xsl:template>

	<xsl:template match="style:default-style" mode="styles"> </xsl:template>

	<xsl:template
		match="style:paragraph-properties[parent::style:style  or parent::style:default-style]"
		mode="styles">
		<w:pPr>
			<xsl:if test="@fo:keep-with-next='always'">
				<w:keepNext/>
			</xsl:if>

			<xsl:if test="@fo:keep-together='always'">
				<w:keepLines/>
			</xsl:if>

			<xsl:if test="@fo:break-before='page'">
				<w:pageBreakBefore/>
			</xsl:if>
			<xsl:if
				test="parent::node()/@style:master-page-name and not(parent::node()/@style:master-page-name = '')">
				<w:pageBreakBefore/>
			</xsl:if>

			<xsl:if test="@fo:widows or @fo:orphans">
				<w:widowControl>
					<xsl:attribute name="w:val">off</xsl:attribute>
				</w:widowControl>
			</xsl:if>

			<!-- border color + padding  -->
			<xsl:if test="@fo:border and @fo:border != 'none' ">
				<w:pBdr>
					<w:top>
						<xsl:choose>
							<xsl:when test="contains(@fo:border, 'double')">
								<xsl:attribute name="w:val">double</xsl:attribute>
							</xsl:when>
							<xsl:otherwise>
								<xsl:attribute name="w:val">single</xsl:attribute>
							</xsl:otherwise>
						</xsl:choose>
						<xsl:attribute name="w:sz">
							<xsl:call-template name="eightspoint-measure">
								<xsl:with-param name="length"
									select="substring-before(@fo:border,  ' ')"/>
							</xsl:call-template>
						</xsl:attribute>
						<xsl:if test="@fo:padding">
							<xsl:attribute name="w:space">
								<xsl:call-template name="point-measure">
									<xsl:with-param name="length" select="@fo:padding"/>
								</xsl:call-template>
							</xsl:attribute>
						</xsl:if>
						<xsl:if test="@fo:padding-top">
							<xsl:attribute name="w:space">
								<xsl:call-template name="point-measure">
									<xsl:with-param name="length" select="@fo:padding-top"/>
								</xsl:call-template>
							</xsl:attribute>
						</xsl:if>

						<xsl:variable name="color"
							select="substring(@fo:border, string-length(@fo:border) -5, 6)"/>
						<xsl:if test="$color != 'none' ">
							<xsl:attribute name="w:color">
								<xsl:value-of select="$color"/>
							</xsl:attribute>
						</xsl:if>

					</w:top>
					<w:left>
						<xsl:choose>
							<xsl:when test="contains(@fo:border, 'double')">
								<xsl:attribute name="w:val">double</xsl:attribute>
							</xsl:when>
							<xsl:otherwise>
								<xsl:attribute name="w:val">single</xsl:attribute>
							</xsl:otherwise>
						</xsl:choose>
						<xsl:attribute name="w:sz">
							<xsl:call-template name="eightspoint-measure">
								<xsl:with-param name="length"
									select="substring-before(@fo:border,  ' ')"/>
							</xsl:call-template>
						</xsl:attribute>
						<xsl:if test="@fo:padding">
							<xsl:attribute name="w:space">
								<xsl:call-template name="point-measure">
									<xsl:with-param name="length" select="@fo:padding"/>
								</xsl:call-template>
							</xsl:attribute>
						</xsl:if>
						<xsl:if test="@fo:padding-left">
							<xsl:attribute name="w:space">
								<xsl:call-template name="point-measure">
									<xsl:with-param name="length" select="@fo:padding-left"/>
								</xsl:call-template>
							</xsl:attribute>
						</xsl:if>
						<xsl:attribute name="w:color">
							<xsl:value-of
								select="substring(@fo:border, string-length(@fo:border) -5, 6)"/>
						</xsl:attribute>
					</w:left>
					<w:bottom>
						<xsl:choose>
							<xsl:when test="contains(@fo:border, 'double')">
								<xsl:attribute name="w:val">double</xsl:attribute>
							</xsl:when>
							<xsl:otherwise>
								<xsl:attribute name="w:val">single</xsl:attribute>
							</xsl:otherwise>
						</xsl:choose>
						<xsl:attribute name="w:sz">
							<xsl:call-template name="eightspoint-measure">
								<xsl:with-param name="length"
									select="substring-before(@fo:border,  ' ')"/>
							</xsl:call-template>
						</xsl:attribute>
						<xsl:if test="@fo:padding">
							<xsl:attribute name="w:space">
								<xsl:call-template name="point-measure">
									<xsl:with-param name="length" select="@fo:padding"/>
								</xsl:call-template>
							</xsl:attribute>
						</xsl:if>
						<xsl:if test="@fo:padding-bottom">
							<xsl:attribute name="w:space">
								<xsl:call-template name="point-measure">
									<xsl:with-param name="length" select="@fo:padding-bottom"/>
								</xsl:call-template>
							</xsl:attribute>
						</xsl:if>
						<xsl:attribute name="w:color">
							<xsl:value-of
								select="substring(@fo:border, string-length(@fo:border) -5, 6)"/>
						</xsl:attribute>
					</w:bottom>
					<w:right>
						<xsl:choose>
							<xsl:when test="contains(@fo:border, 'double')">
								<xsl:attribute name="w:val">double</xsl:attribute>
							</xsl:when>
							<xsl:otherwise>
								<xsl:attribute name="w:val">single</xsl:attribute>
							</xsl:otherwise>
						</xsl:choose>
						<xsl:attribute name="w:sz">
							<xsl:call-template name="eightspoint-measure">
								<xsl:with-param name="length"
									select="substring-before(@fo:border,  ' ')"/>
							</xsl:call-template>
						</xsl:attribute>
						<xsl:if test="@fo:padding">
							<xsl:attribute name="w:space">
								<xsl:call-template name="point-measure">
									<xsl:with-param name="length" select="@fo:padding"/>
								</xsl:call-template>
							</xsl:attribute>
						</xsl:if>
						<xsl:if test="@fo:padding-right">
							<xsl:attribute name="w:space">
								<xsl:call-template name="point-measure">
									<xsl:with-param name="length" select="@fo:padding-right"/>
								</xsl:call-template>
							</xsl:attribute>
						</xsl:if>
						<xsl:attribute name="w:color">
							<xsl:value-of
								select="substring(@fo:border, string-length(@fo:border) -5, 6)"/>
						</xsl:attribute>
					</w:right>
				</w:pBdr>
			</xsl:if>
			<xsl:if
				test="@fo:border-top or @fo:border-left or @fo:border-bottom or @fo:border-right">
				<w:pBdr>
					<xsl:if test="@fo:border-top and (@fo:border-top != 'none')">
						<w:top>
							<xsl:choose>
								<xsl:when test="contains(@fo:border-top, 'double')">
									<xsl:attribute name="w:val">double</xsl:attribute>
								</xsl:when>
								<xsl:otherwise>
									<xsl:attribute name="w:val">single</xsl:attribute>
								</xsl:otherwise>
							</xsl:choose>
							<xsl:attribute name="w:sz">
								<xsl:call-template name="eightspoint-measure">
									<xsl:with-param name="length"
										select="substring-before(@fo:border-top,  ' ')"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:if test="@fo:padding-top">
								<xsl:attribute name="w:space">
									<xsl:call-template name="point-measure">
										<xsl:with-param name="length" select="@fo:padding-top"/>
									</xsl:call-template>
								</xsl:attribute>
							</xsl:if>
							<xsl:attribute name="w:color">
								<xsl:value-of
									select="substring(@fo:border-top, string-length(@fo:border-top) -5, 6)"
								/>
							</xsl:attribute>
						</w:top>
					</xsl:if>
					<xsl:if test="@fo:border-left and (@fo:border-left != 'none')">
						<w:left>
							<xsl:choose>
								<xsl:when test="contains(@fo:border-left, 'double')">
									<xsl:attribute name="w:val">double</xsl:attribute>
								</xsl:when>
								<xsl:otherwise>
									<xsl:attribute name="w:val">single</xsl:attribute>
								</xsl:otherwise>
							</xsl:choose>
							<xsl:attribute name="w:sz">
								<xsl:call-template name="eightspoint-measure">
									<xsl:with-param name="length"
										select="substring-before(@fo:border-left,  ' ')"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:if test="@fo:padding-left">
								<xsl:attribute name="w:space">
									<xsl:call-template name="point-measure">
										<xsl:with-param name="length" select="@fo:padding-left"/>
									</xsl:call-template>
								</xsl:attribute>
							</xsl:if>
							<xsl:attribute name="w:color">
								<xsl:value-of
									select="substring(@fo:border-left, string-length(@fo:border-left) -5, 6)"
								/>
							</xsl:attribute>
						</w:left>
					</xsl:if>
					<xsl:if test="@fo:border-bottom and (@fo:border-bottom != 'none')">
						<w:bottom>
							<xsl:choose>
								<xsl:when test="contains(@fo:border-bottom, 'double')">
									<xsl:attribute name="w:val">double</xsl:attribute>
								</xsl:when>
								<xsl:otherwise>
									<xsl:attribute name="w:val">single</xsl:attribute>
								</xsl:otherwise>
							</xsl:choose>
							<xsl:attribute name="w:sz">
								<xsl:call-template name="eightspoint-measure">
									<xsl:with-param name="length"
										select="substring-before(@fo:border-bottom,  ' ')"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:if test="@fo:padding-bottom">
								<xsl:attribute name="w:space">
									<xsl:call-template name="point-measure">
										<xsl:with-param name="length" select="@fo:padding-bottom"/>
									</xsl:call-template>
								</xsl:attribute>
							</xsl:if>
							<xsl:attribute name="w:color">
								<xsl:value-of
									select="substring(@fo:border-bottom, string-length(@fo:border-bottom) -5, 6)"
								/>
							</xsl:attribute>
						</w:bottom>
					</xsl:if>
					<xsl:if test="@fo:border-right and (@fo:border-right != 'none')">
						<w:right>
							<xsl:choose>
								<xsl:when test="contains(@fo:border-right, 'double')">
									<xsl:attribute name="w:val">double</xsl:attribute>
								</xsl:when>
								<xsl:otherwise>
									<xsl:attribute name="w:val">single</xsl:attribute>
								</xsl:otherwise>
							</xsl:choose>
							<xsl:attribute name="w:sz">
								<xsl:call-template name="eightspoint-measure">
									<xsl:with-param name="length"
										select="substring-before(@fo:border-right,  ' ')"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:if test="@fo:padding-right">
								<xsl:attribute name="w:space">
									<xsl:call-template name="point-measure">
										<xsl:with-param name="length" select="@fo:padding-right"/>
									</xsl:call-template>
								</xsl:attribute>
							</xsl:if>
							<xsl:attribute name="w:color">
								<xsl:value-of
									select="substring(@fo:border-right, string-length(@fo:border-right) -5, 6)"
								/>
							</xsl:attribute>
						</w:right>
					</xsl:if>
				</w:pBdr>
			</xsl:if>

			<!-- background color -->
			<xsl:if test="@fo:background-color and (@fo:background-color != 'transparent')">
				<w:shd>
					<xsl:attribute name="w:val">clear</xsl:attribute>
					<xsl:attribute name="w:color">auto</xsl:attribute>
					<xsl:attribute name="w:fill">
						<xsl:value-of
							select="substring(@fo:background-color, 2, string-length(@fo:background-color) -1)"
						/>
					</xsl:attribute>
				</w:shd>
			</xsl:if>

			<!-- tabs -->
			<xsl:if test="style:tab-stops/style:tab-stop">
				<w:tabs>
					<xsl:for-each select="style:tab-stops/style:tab-stop">
						<xsl:call-template name="tabStop"/>
					</xsl:for-each>
					<xsl:variable name="styleName">
						<xsl:value-of select="parent::node()/@style:name"/>
					</xsl:variable>
					<xsl:variable name="parentstyleName">
						<xsl:value-of select="parent::node()/@style:parent-style-name"/>
					</xsl:variable>
					<xsl:for-each select="document('styles.xml')//style:style[@style:name=$parentstyleName]//style:tab-stop">
						<xsl:if test="not(@style:position=document('content.xml')//style:style[@style:name=$styleName]//style:tab-stop/@style:position)">
							<xsl:call-template name="tabStop">
								<xsl:with-param name="styleType">clear</xsl:with-param>
							</xsl:call-template>
						</xsl:if>
					</xsl:for-each>
				</w:tabs>
			</xsl:if>

			<xsl:if test="@style:text-autospace">
				<w:autoSpaceDN>
					<xsl:choose>
						<xsl:when test="@style:text-autospace='none'">
							<xsl:attribute name="w:val">off</xsl:attribute>
						</xsl:when>
						<xsl:when test="@style:text-autospace='ideograph-alpha'">
							<xsl:attribute name="w:val">on</xsl:attribute>
						</xsl:when>
					</xsl:choose>
				</w:autoSpaceDN>
			</xsl:if>

			<xsl:if
				test="@style:line-height-at-least or @fo:line-height or @fo:margin-bottom or @fo:margin-top or @style:line-spacing">
				<w:spacing>
					<xsl:if test="@style:line-height-at-least">
						<xsl:attribute name="w:lineRule">atLeast</xsl:attribute>
						<xsl:attribute name="w:line">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length" select="@style:line-height-at-least"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="@style:line-spacing">
						<xsl:attribute name="w:lineRule">auto</xsl:attribute>
						<xsl:attribute name="w:line">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length" select="@style:line-spacing"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="contains(@fo:line-height, '%')">
						<xsl:attribute name="w:lineRule">auto</xsl:attribute>
						<xsl:attribute name="w:line">
							<xsl:value-of
								select="round(number(substring-before(@fo:line-height, '%')) * 240 div 100)"
							/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="contains(@fo:line-height, 'cm')">
						<xsl:attribute name="w:lineRule">exact</xsl:attribute>
						<xsl:attribute name="w:line">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length" select="@fo:line-height"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="@fo:margin-bottom">
						<xsl:attribute name="w:after">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length" select="@fo:margin-bottom"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="@fo:margin-top">
						<xsl:attribute name="w:before">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length" select="@fo:margin-top"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
				</w:spacing>
			</xsl:if>

			<xsl:if
				test="@fo:margin-left or @fo:margin-right or @fo:text-indent or @text:space-before">
				<w:ind>
					<xsl:if test="@fo:margin-left">
						<xsl:attribute name="w:left">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length" select="@fo:margin-left"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="@text:space-before">
						<xsl:attribute name="w:left">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length" select="@text:space-before"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="@fo:margin-right">
						<xsl:attribute name="w:right">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length" select="@fo:margin-right"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="@fo:text-indent and (@style:auto-text-indent!='true')">
						<xsl:attribute name="w:firstLine">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length" select="@fo:text-indent"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
					<!-- Trick to replace automatic indent for OOX -->
					<xsl:if test="@style:auto-text-indent='true'">
						<xsl:attribute name="w:firstLine">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length"
									select="parent::style:style/style:text-properties/@fo:font-size"
								/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
				</w:ind>
			</xsl:if>

			<!-- TODO this should be modified when we will manage bidi properties -->
			<xsl:if test="@fo:text-align">
				<w:jc>
					<xsl:choose>
						<xsl:when test="@fo:text-align='center'">
							<xsl:attribute name="w:val">center</xsl:attribute>
						</xsl:when>
						<xsl:when test="@fo:text-align='start'">
							<xsl:attribute name="w:val">left</xsl:attribute>
						</xsl:when>
						<xsl:when test="@fo:text-align='end'">
							<xsl:attribute name="w:val">right</xsl:attribute>
						</xsl:when>
						<xsl:when test="@fo:text-align='justify'">
							<xsl:attribute name="w:val">both</xsl:attribute>
						</xsl:when>
						<xsl:otherwise>
							<xsl:attribute name="w:val">start</xsl:attribute>
						</xsl:otherwise>
					</xsl:choose>

				</w:jc>
			</xsl:if>

			<xsl:if test="@style:vertical-align">
				<w:textAlignment>
					<xsl:choose>
						<xsl:when test="@style:vertical-align='bottom'">
							<xsl:attribute name="w:val">bottom</xsl:attribute>
						</xsl:when>
						<xsl:when test="@style:vertical-align='top'">
							<xsl:attribute name="w:val">top</xsl:attribute>
						</xsl:when>
						<xsl:when test="@style:vertical-align='middle'">
							<xsl:attribute name="w:val">center</xsl:attribute>
						</xsl:when>
						<xsl:otherwise>
							<xsl:attribute name="w:val">auto</xsl:attribute>
						</xsl:otherwise>
					</xsl:choose>
				</w:textAlignment>
			</xsl:if>
		</w:pPr>
	</xsl:template>

	<xsl:template match="style:text-properties[parent::style:style  or parent::style:default-style]"
		mode="styles">
		<w:rPr>
			<xsl:if test="@style:font-name">
				<w:rFonts>
					<xsl:if test="@style:font-name">
						<xsl:attribute name="w:ascii">
							<xsl:call-template name="computeFontName">
								<xsl:with-param name="fontName" select="@style:font-name"/>
							</xsl:call-template>
						</xsl:attribute>
						<xsl:attribute name="w:hAnsi">
							<xsl:call-template name="computeFontName">
								<xsl:with-param name="fontName" select="@style:font-name"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="@style:font-name-complex">
						<xsl:attribute name="w:cs">
							<xsl:call-template name="computeFontName">
								<xsl:with-param name="fontName" select="@style:font-name-complex"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="@style:font-name-asian">
						<xsl:attribute name="w:eastAsia">
							<xsl:call-template name="computeFontName">
								<xsl:with-param name="fontName" select="@style:font-name-asian"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
				</w:rFonts>
			</xsl:if>

			<xsl:if test="@fo:font-weight or @fo:font-weight-asian">
				<xsl:choose>
					<xsl:when test="@fo:font-weight = 'bold' or @fo:font-weight-asian = 'bold'">
						<w:b w:val="on"/>
					</xsl:when>
					<xsl:when test="@fo:font-weight = 'normal' or @fo:font-weight-asian = 'normal'">
						<w:b w:val="off"/>
					</xsl:when>
				</xsl:choose>
			</xsl:if>

			<xsl:if test="@fo:font-weight-complex">
				<xsl:choose>
					<xsl:when test="@style:font-weight-complex = 'bold'">
						<w:bCs w:val="on"/>
					</xsl:when>
					<xsl:when test="@style:font-weight-complex = 'normal'">
						<w:bCs w:val="off"/>
					</xsl:when>
				</xsl:choose>
			</xsl:if>

			<!-- TODO determine all the different font styles-->
			<xsl:if test="@fo:font-style or @style:font-style-asian or @style:font-style-complex">
				<xsl:choose>
					<xsl:when
						test="@fo:font-style = 'italic' or @style:font-style-asian = 'italic' or @style:font-style-complex = 'italic'">
						<w:i w:val="on"/>
					</xsl:when>
					<xsl:when
						test="@fo:font-style = 'oblique' or @style:font-style-asian = 'oblique' or @style:font-style-complex = 'oblique'">
						<w:i w:val="on"/>
					</xsl:when>
					<xsl:otherwise>
						<w:i w:val="off"/>
					</xsl:otherwise>
					<!-- It could be also oblique in fo DTD, but it is not possible to set it via Ooo interface -->
				</xsl:choose>
			</xsl:if>

			<xsl:if test="@fo:text-transform">
				<xsl:choose>
					<xsl:when test="@fo:text-transform = 'uppercase'">
						<w:caps w:val="on"/>
					</xsl:when>
					<xsl:when test="@fo:text-transform = 'none'">
						<w:caps w:val="off"/>
					</xsl:when>
					<!-- It could be also lowercase or capitalize in fo DTD, but it is not possible to set it via word 2007 interface -->
				</xsl:choose>
			</xsl:if>

			<xsl:if test="@fo:font-variant">
				<xsl:choose>
					<xsl:when test="@fo:font-variant = 'small-caps'">
						<w:smallCaps w:val="on"/>
					</xsl:when>
					<xsl:when test="@fo:font-weight = 'normal'">
						<w:smallCaps w:val="off"/>
					</xsl:when>
				</xsl:choose>
			</xsl:if>

			<xsl:if test="@style:text-line-through-style != 'none'">
				<xsl:choose>
					<xsl:when test="@style:text-line-through-type = 'double'">
						<w:dstrike w:val="on"/>
					</xsl:when>
					<xsl:when test="@style:text-line-through-type = 'none'">
						<w:strike w:val="off"/>
					</xsl:when>
					<xsl:otherwise>
						<w:strike w:val="on"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:if>

			<xsl:if test="@style:text-outline">
				<xsl:choose>
					<xsl:when test="@style:text-outline = 'true'">
						<w:outline w:val="on"/>
					</xsl:when>
					<xsl:when test="@style:text-outline = 'false'">
						<w:outline w:val="off"/>
					</xsl:when>
				</xsl:choose>
			</xsl:if>

			<xsl:if test="(@fo:text-shadow)">
				<xsl:choose>
					<xsl:when test="@fo:text-shadow = 'none'">
						<w:shadow w:val="off"/>
					</xsl:when>
					<xsl:otherwise>
						<w:shadow w:val="on"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:if>

			<xsl:if test="(@style:font-relief)">
				<xsl:choose>
					<xsl:when test="@style:font-relief = 'embossed'">
						<w:emboss w:val="on"/>
					</xsl:when>
					<xsl:when test="@style:font-relief = 'engraved'">
						<w:imprint w:val="on"/>
					</xsl:when>
					<xsl:otherwise>
						<w:emboss w:val="off"/>
						<w:imprint w:val="off"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:if>

			<xsl:if test="@text:display">
				<xsl:choose>
					<xsl:when test="@text:display = 'true'">
						<w:vanish w:val="on"/>
					</xsl:when>
					<xsl:when test="@text:display = 'none'">
						<w:vanish w:val="off"/>
					</xsl:when>
				</xsl:choose>
			</xsl:if>

			<xsl:if test="@fo:color">
				<w:color>
					<xsl:attribute name="w:val">
						<xsl:value-of select="substring(@fo:color, 2, string-length(@fo:color) -1)"
						/>
					</xsl:attribute>
				</w:color>
			</xsl:if>

			<xsl:if test="@fo:letter-spacing">
				<w:spacing>
					<xsl:attribute name="w:val">
						<xsl:choose>
							<xsl:when test="@fo:letter-spacing='normal'">
								<xsl:value-of select="0"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:call-template name="twips-measure">
									<xsl:with-param name="length" select="@fo:letter-spacing"/>
								</xsl:call-template>		
							</xsl:otherwise>						
						</xsl:choose>
						
					</xsl:attribute>
				</w:spacing>
			</xsl:if>

			<xsl:if test="@style:text-scale">
				<w:w>
					<xsl:attribute name="w:val">
						<xsl:variable name="scale" select="substring(@style:text-scale, 1, string-length(@style:text-scale)-1)"/>
						<xsl:choose>
							<xsl:when test="number($scale) &lt; 600"><xsl:value-of select="$scale"/></xsl:when>
							<xsl:otherwise><xsl:message terminate="no">feedback: Text scale can't exceed 600%</xsl:message>
								600</xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
				</w:w>
			</xsl:if>

			<xsl:if test="@style:letter-kerning = 'true'">
				<w:kern w:val="16"/>
			</xsl:if>

			<xsl:if test="@fo:font-size">
				<xsl:variable name="sz">
					<xsl:call-template name="computeSize">
						<xsl:with-param name="node" select="current()"/>
					</xsl:call-template>
				</xsl:variable>
				<xsl:if test="number($sz)">
					<w:sz w:val="{$sz}"/>
				</xsl:if>
			</xsl:if>

			<xsl:if test="@fo:font-size-complex">
				<w:szCs>
					<xsl:attribute name="w:val">
						<xsl:value-of
							select="number(substring-before(@fo:font-size-complex, 'pt')) * 2"/>
					</xsl:attribute>
				</w:szCs>
			</xsl:if>

			<xsl:if test="@style:text-underline-style != 'none' ">
				<w:u>
					<xsl:attribute name="w:val">
						<xsl:choose>
							<xsl:when test="@style:text-underline-style = 'dotted'">
								<xsl:choose>
									<xsl:when test="@style:text-underline-width = 'thick' "
										>dottedHeavy</xsl:when>
									<xsl:when test="@style:text-underline-width = 'bold' "
										>dottedHeavy</xsl:when>
									<xsl:otherwise>dotted</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="@style:text-underline-style = 'dash' ">
								<xsl:choose>
									<xsl:when test="@style:text-underline-width = 'thick' "
										>dashedHeavy</xsl:when>
									<xsl:when test="@style:text-underline-width = 'bold' "
										>dashedHeavy</xsl:when>
									<xsl:otherwise>dash</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="@style:text-underline-style = 'long-dash'">
								<xsl:choose>
									<xsl:when test="@style:text-underline-width = 'thick' "
										>dashLongHeavy</xsl:when>
									<xsl:when test="@style:text-underline-width = 'bold' "
										>dashLongHeavy</xsl:when>
									<xsl:otherwise>dashLong</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="@style:text-underline-style = 'dot-dash' ">
								<xsl:choose>
									<xsl:when test="@style:text-underline-width = 'thick' "
										>dashDotHeavy</xsl:when>
									<xsl:when test="@style:text-underline-width = 'bold' "
										>dashDotHeavy</xsl:when>
									<xsl:otherwise>dotDash</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="@style:text-underline-style = 'dot-dot-dash' ">
								<xsl:choose>
									<xsl:when test="@style:text-underline-width = 'thick' "
										>dashDotDotHeavy</xsl:when>
									<xsl:when test="@style:text-underline-width = 'bold' "
										>dashDotDotHeavy</xsl:when>
									<xsl:otherwise>dotDotDash</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="@style:text-underline-style = 'wave' ">
								<xsl:choose>
									<xsl:when test="@style:text-underline-type = 'double' "
										>wavyDouble</xsl:when>
									<xsl:when test="@style:text-underline-width = 'thick' "
										>wavyHeavy</xsl:when>
									<xsl:when test="@style:text-underline-width = 'bold' ">wavyHeavy</xsl:when>
									<xsl:otherwise>wave</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:otherwise>
								<xsl:choose>
									<xsl:when test="@style:text-underline-type = 'double' ">double</xsl:when>
									<xsl:when test="@style:text-underline-width = 'thick' ">thick</xsl:when>
									<xsl:when test="@style:text-underline-width = 'bold' ">thick</xsl:when>
									<xsl:when
										test="@style:text-underline-mode = 'skip-white-space' "
										>words</xsl:when>
									<xsl:otherwise>single</xsl:otherwise>
								</xsl:choose>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
					<xsl:if test="@style:text-underline-color">
						<xsl:attribute name="w:color">
							<xsl:choose>
								<xsl:when test="@style:text-underline-color = 'font-color'">auto</xsl:when>
								<xsl:otherwise>
									<xsl:value-of
										select="substring(@style:text-underline-color, 2, string-length(@style:text-underline-color)-1)"
									/>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</xsl:if>
				</w:u>
			</xsl:if>

			<xsl:if test="@fo:background-color">
				<w:shd w:val="clear" w:color="auto">
					<xsl:attribute name="w:fill">
						<xsl:choose>
							<xsl:when test="@fo:background-color = 'transparent'">auto</xsl:when>
							<xsl:otherwise>
								<xsl:value-of
									select="substring(@fo:background-color, 2, string-length(@fo:background-color)-1)"
								/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
				</w:shd>
			</xsl:if>

			<xsl:if test="@style:text-emphasize">
				<w:em>
					<xsl:attribute name="w:val">
						<xsl:choose>
							<xsl:when test="@style:text-emphasize = 'accent above'">comma</xsl:when>
							<xsl:when test="@style:text-emphasize = 'dot above'">dot</xsl:when>
							<xsl:when test="@style:text-emphasize = 'circle above'">circle</xsl:when>
							<xsl:when test="@style:text-emphasize = 'dot below'">underDot</xsl:when>
							<xsl:otherwise>none</xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
				</w:em>
			</xsl:if>

			<xsl:if test="@style:text-position">
				<xsl:choose>
					<xsl:when test="substring(@style:text-position, 1, 3) = 'sub'">
						<w:vertAlign w:val="subscript"/>
					</xsl:when>
					<xsl:when test="substring(@style:text-position, 1, 5) = 'super'">
						<w:vertAlign w:val="superscript"/>
					</xsl:when>
				</xsl:choose>
			</xsl:if>

			<xsl:if
				test="(@fo:language and @fo:country) or (@style:language-asian and @style:country-asian) or (@style:language-complex and @style:country-complex)">
				<w:lang>
					<xsl:if test="(@fo:language and @fo:country)">
						<xsl:attribute name="w:val"><xsl:value-of select="@fo:language"
								/>-<xsl:value-of select="@fo:country"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="(@style:language-asian and @style:country-asian)">
						<xsl:attribute name="w:eastAsia"><xsl:value-of
								select="@style:language-asian"/>-<xsl:value-of
								select="@style:country-asian"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="(@style:language-complex and @style:country-complex)">
						<xsl:attribute name="w:bidi"><xsl:value-of select="@style:language-complex"
								/>-<xsl:value-of select="@style:country-complex"/></xsl:attribute>
					</xsl:if>
				</w:lang>
			</xsl:if>

			<xsl:if test="@style:text-combine">
				<xsl:choose>
					<xsl:when test="@style:text-combine = 'lines'">
						<w:eastAsianLayout>
							<xsl:attribute name="w:id">
								<xsl:number count="style:style"/>
							</xsl:attribute>
							<xsl:attribute name="w:combine">lines</xsl:attribute>
							<xsl:choose>
								<xsl:when
									test="@style:text-combine-start-char = '&lt;' and @style:text-combine-end-char = '&gt;'">
									<xsl:attribute name="w:combineBrackets">angle</xsl:attribute>
								</xsl:when>
								<xsl:when
									test="@style:text-combine-start-char = '{' and @style:text-combine-end-char = '}'">
									<xsl:attribute name="w:combineBrackets">curly</xsl:attribute>
								</xsl:when>
								<xsl:when
									test="@style:text-combine-start-char = '(' and @style:text-combine-end-char = ')'">
									<xsl:attribute name="w:combineBrackets">round</xsl:attribute>
								</xsl:when>
								<xsl:when
									test="@style:text-combine-start-char = '[' and @style:text-combine-end-char = ']'">
									<xsl:attribute name="w:combineBrackets">square</xsl:attribute>
								</xsl:when>
								<xsl:otherwise>
									<xsl:attribute name="w:combineBrackets">none</xsl:attribute>
								</xsl:otherwise>
							</xsl:choose>
						</w:eastAsianLayout>
					</xsl:when>
				</xsl:choose>
			</xsl:if>

		</w:rPr>
	</xsl:template>

	<!-- toggle properties : this text properties must be applied into the template "text:span" mode="paragraph" 
	to erase the paragraph style -->
	<xsl:template match="style:text-properties[parent::style:style  or parent::style:default-style]" mode="toggle">
		<xsl:if test="@fo:font-weight or @fo:font-weight-asian">
			<xsl:choose>
				<xsl:when test="@fo:font-weight = 'bold' or @fo:font-weight-asian = 'bold'">
					<w:b w:val="on"/>
				</xsl:when>
				<xsl:when test="@fo:font-weight = 'normal' or @fo:font-weight-asian = 'normal'">
					<w:b w:val="off"/>
				</xsl:when>
			</xsl:choose>
		</xsl:if>
		
		<xsl:if test="@fo:font-weight-complex">
			<xsl:choose>
				<xsl:when test="@style:font-weight-complex = 'bold'">
					<w:bCs w:val="on"/>
				</xsl:when>
				<xsl:when test="@style:font-weight-complex = 'normal'">
					<w:bCs w:val="off"/>
				</xsl:when>
			</xsl:choose>
		</xsl:if>
		
		<!-- TODO determine all the different font styles-->
		<xsl:if test="@fo:font-style or @style:font-style-asian or @style:font-style-complex">
			<xsl:choose>
				<xsl:when
					test="@fo:font-style = 'italic' or @style:font-style-asian = 'italic' or @style:font-style-complex = 'italic'">
					<w:i w:val="on"/>
				</xsl:when>
				<xsl:when
					test="@fo:font-style = 'oblique' or @style:font-style-asian = 'oblique' or @style:font-style-complex = 'oblique'">
					<w:i w:val="on"/>
				</xsl:when>
				<xsl:otherwise>
					<w:i w:val="off"/>
				</xsl:otherwise>
				<!-- It could be also oblique in fo DTD, but it is not possible to set it via Ooo interface -->
			</xsl:choose>
		</xsl:if>
		
		<xsl:if test="@fo:text-transform">
			<xsl:choose>
				<xsl:when test="@fo:text-transform = 'uppercase'">
					<w:caps w:val="on"/>
				</xsl:when>
				<xsl:when test="@fo:text-transform = 'none'">
					<w:caps w:val="off"/>
				</xsl:when>
				<!-- It could be also lowercase or capitalize in fo DTD, but it is not possible to set it via word 2007 interface -->
			</xsl:choose>
		</xsl:if>
		
		<xsl:if test="@fo:font-variant">
			<xsl:choose>
				<xsl:when test="@fo:font-variant = 'small-caps'">
					<w:smallCaps w:val="on"/>
				</xsl:when>
				<xsl:when test="@fo:font-weight = 'normal'">
					<w:smallCaps w:val="off"/>
				</xsl:when>
			</xsl:choose>
		</xsl:if>
		
		<xsl:if test="@style:text-line-through-style">
			<xsl:choose>
				<xsl:when test="@style:text-line-through-type = 'double'">
					<w:dstrike w:val="on"/>
				</xsl:when>
				<xsl:when test="@style:text-line-through-type = 'none'">
					<w:strike w:val="off"/>
				</xsl:when>
				<xsl:otherwise>
					<w:strike w:val="on"/>
				</xsl:otherwise>
			</xsl:choose>
	
		</xsl:if>
		
		<xsl:if test="@style:text-outline">
			<xsl:choose>
				<xsl:when test="@style:text-outline = 'true'">
					<w:outline w:val="on"/>
				</xsl:when>
				<xsl:when test="@style:text-outline = 'false'">
					<w:outline w:val="off"/>
				</xsl:when>
			</xsl:choose>
		</xsl:if>
		
		<xsl:if test="(@fo:text-shadow)">
			<xsl:choose>
				<xsl:when test="@fo:text-shadow = 'none'">
					<w:shadow w:val="off"/>
				</xsl:when>
				<xsl:otherwise>
					<w:shadow w:val="on"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:if>
		
		<xsl:if test="(@style:font-relief)">
			<xsl:choose>
				<xsl:when test="@style:font-relief = 'embossed'">
					<w:emboss w:val="on"/>
				</xsl:when>
				<xsl:when test="@style:font-relief = 'engraved'">
					<w:imprint w:val="on"/>
				</xsl:when>
				<xsl:otherwise>
					<w:emboss w:val="off"/>
					<w:imprint w:val="off"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:if>
		
		<xsl:if test="@text:display">
			<xsl:choose>
				<xsl:when test="@text:display = 'true'">
					<w:vanish w:val="on"/>
				</xsl:when>
				<xsl:when test="@text:display = 'none'">
					<w:vanish w:val="off"/>
				</xsl:when>
			</xsl:choose>
		</xsl:if>
		
		<xsl:if test="@style:text-underline-style">
			<xsl:choose>
		<xsl:when test="@style:text-underline-style != 'none' ">
			<w:u>
				<xsl:attribute name="w:val">
					<xsl:choose>
						<xsl:when test="@style:text-underline-style = 'dotted'">
							<xsl:choose>
								<xsl:when test="@style:text-underline-width = 'thick' "
									>dottedHeavy</xsl:when>
								<xsl:when test="@style:text-underline-width = 'bold' "
									>dottedHeavy</xsl:when>
								<xsl:otherwise>dotted</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="@style:text-underline-style = 'dash' ">
							<xsl:choose>
								<xsl:when test="@style:text-underline-width = 'thick' "
									>dashedHeavy</xsl:when>
								<xsl:when test="@style:text-underline-width = 'bold' "
									>dashedHeavy</xsl:when>
								<xsl:otherwise>dash</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="@style:text-underline-style = 'long-dash'">
							<xsl:choose>
								<xsl:when test="@style:text-underline-width = 'thick' "
									>dashLongHeavy</xsl:when>
								<xsl:when test="@style:text-underline-width = 'bold' "
									>dashLongHeavy</xsl:when>
								<xsl:otherwise>dashLong</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="@style:text-underline-style = 'dot-dash' ">
							<xsl:choose>
								<xsl:when test="@style:text-underline-width = 'thick' "
									>dashDotHeavy</xsl:when>
								<xsl:when test="@style:text-underline-width = 'bold' "
									>dashDotHeavy</xsl:when>
								<xsl:otherwise>dotDash</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="@style:text-underline-style = 'dot-dot-dash' ">
							<xsl:choose>
								<xsl:when test="@style:text-underline-width = 'thick' "
									>dashDotDotHeavy</xsl:when>
								<xsl:when test="@style:text-underline-width = 'bold' "
									>dashDotDotHeavy</xsl:when>
								<xsl:otherwise>dotDotDash</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:when test="@style:text-underline-style = 'wave' ">
							<xsl:choose>
								<xsl:when test="@style:text-underline-type = 'double' "
									>wavyDouble</xsl:when>
								<xsl:when test="@style:text-underline-width = 'thick' "
									>wavyHeavy</xsl:when>
								<xsl:when test="@style:text-underline-width = 'bold' ">wavyHeavy</xsl:when>
								<xsl:otherwise>wave</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:otherwise>
							<xsl:choose>
								<xsl:when test="@style:text-underline-type = 'double' ">double</xsl:when>
								<xsl:when test="@style:text-underline-width = 'thick' ">thick</xsl:when>
								<xsl:when test="@style:text-underline-width = 'bold' ">thick</xsl:when>
								<xsl:when
									test="@style:text-underline-mode = 'skip-white-space' "
									>words</xsl:when>
								<xsl:otherwise>single</xsl:otherwise>
							</xsl:choose>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
				<xsl:if test="@style:text-underline-color">
					<xsl:attribute name="w:color">
						<xsl:choose>
							<xsl:when test="@style:text-underline-color = 'font-color'">auto</xsl:when>
							<xsl:otherwise>
								<xsl:value-of
									select="substring(@style:text-underline-color, 2, string-length(@style:text-underline-color)-1)"
								/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
				</xsl:if>
			</w:u>
			</xsl:when>
			<xsl:otherwise>
				<w:u w:val="none"/>
			</xsl:otherwise>					
			</xsl:choose>
		</xsl:if>
		<xsl:if test="@style:text-emphasize">
			<w:em>
				<xsl:attribute name="w:val">
					<xsl:choose>
						<xsl:when test="@style:text-emphasize = 'accent above'">comma</xsl:when>
						<xsl:when test="@style:text-emphasize = 'dot above'">dot</xsl:when>
						<xsl:when test="@style:text-emphasize = 'circle above'">circle</xsl:when>
						<xsl:when test="@style:text-emphasize = 'dot below'">underDot</xsl:when>
						<xsl:otherwise>none</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
			</w:em>
		</xsl:if>
	</xsl:template>
		
	<xsl:template match="style:graphic-properties[parent::style:style]" mode="styles">
		<w:pPr>
			<xsl:if test="@style:horizontal-pos">
				<w:jc>
					<xsl:attribute name="w:val">
						<xsl:choose>
							<xsl:when test="@style:horizontal-pos = 'center'">center</xsl:when>
							<xsl:when test="@style:horizontal-pos='left'">left</xsl:when>
							<xsl:when test="@style:horizontal-pos='right'">right</xsl:when>
							<xsl:otherwise>center</xsl:otherwise>
							<!--
							<value>from-left</value>
							<value>inside</value>
							<value>outside</value>
							<value>from-inside</value>
							-->
							<!-- TODO manage horizontal-rel -->
						</xsl:choose>
					</xsl:attribute>
				</w:jc>
			</xsl:if>
		</w:pPr>
	</xsl:template>


	<!-- footnote and endnote reference and text styles -->
	<xsl:template
		match="text:notes-configuration[@text:note-class='footnote'] | text:notes-configuration[@text:note-class='endnote']"
		mode="styles">
		<w:style w:styleId="{concat(@text:note-class, 'Reference')}" w:type="character">
			<w:name w:val="note reference"/>
			<w:basedOn w:val="{@text:citation-body-style-name}"/>
			<w:semiHidden/>
			<w:unhideWhenUsed/>
		</w:style>
		<w:style w:styleId="{concat(@text:note-class, 'Text')}" w:type="paragraph">
			<w:name w:val="note text"/>
			<w:basedOn w:val="{@text:citation-style-name}"/>
			<!--w:link w:val="FootnoteTextChar"/-->
			<w:semiHidden/>
			<w:unhideWhenUsed/>
		</w:style>
	</xsl:template>


	<xsl:template name="computeSize">
		<xsl:param name="node"/>
		<xsl:if test="contains($node/@fo:font-size, 'pt')">
			<xsl:value-of select="round(number(substring-before($node/@fo:font-size, 'pt')) * 2)"/>
		</xsl:if>
		<xsl:if test="contains($node/@fo:font-size, '%')">
			<xsl:variable name="parentStyleName" select="$node/../@style:parent-style-name"/>
			<xsl:variable name="value">
				<xsl:call-template name="computeSize">
					<!-- should we look for @style:name in styles.xml, otherwise in content.xml ? -->
					<xsl:with-param name="node"
						select="/office:document-styles/office:styles/style:style[@style:name = $parentStyleName]/style:text-properties"
					/>
				</xsl:call-template>
			</xsl:variable>
			<xsl:choose>
				<xsl:when test="number($value)">
					<xsl:value-of
						select="round(number(substring-before($node/@fo:font-size, '%')) div 100 * number($value))"
					/>
				</xsl:when>
				<xsl:otherwise>
					<!-- fetch the default font size for this style family -->
					<xsl:variable name="defaultProps"
						select="/office:document-styles/office:styles/style:default-style[@style:family=$node/../@style:family]/style:text-properties"/>
					<xsl:variable name="defaultValue"
						select="number(substring-before($defaultProps/@fo:font-size, 'pt'))*2"/>
					<xsl:value-of
						select="round(number(substring-before($node/@fo:font-size, '%')) div 100 * number($defaultValue))"
					/>
				</xsl:otherwise>
			</xsl:choose>

		</xsl:if>
	</xsl:template>

	<xsl:template name="tabStop">
		<xsl:param name="styleType"/>
		<w:tab>
			<xsl:attribute name="w:pos">
				<xsl:variable name="position">
					<xsl:call-template name="twips-measure">
						<xsl:with-param name="length" select="ancestor::style:paragraph-properties/@fo:margin-left"/>
					</xsl:call-template>
				</xsl:variable>
				<xsl:variable name="position2">
					<xsl:call-template name="twips-measure">
						<xsl:with-param name="length" select="@style:position"/>
					</xsl:call-template>
				</xsl:variable>
				<xsl:value-of select="$position+$position2"/>
			</xsl:attribute>
			<xsl:attribute name="w:val">
				<xsl:choose>
					<xsl:when test="@style:type">	
						<xsl:value-of select="@style:type"/>
					</xsl:when>
					<xsl:when test="$styleType">
						<xsl:value-of select="$styleType"/>
					</xsl:when>
					<xsl:otherwise>left</xsl:otherwise>
				</xsl:choose>
			</xsl:attribute>
			<xsl:attribute name="w:leader">
				<xsl:choose>
					<xsl:when test="@style:leader-text">
						<xsl:choose>
							<xsl:when test="@style:leader-text='.'">
								<xsl:value-of select="'dot'"/>
							</xsl:when>
							<xsl:when test="@style:leader-text='-'">
								<xsl:value-of select="'hyphen'"/>
							</xsl:when>
							<xsl:when test="@style:leader-text='_'">
								<xsl:value-of select="'heavy'"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="'none'"/>
							</xsl:otherwise>
						</xsl:choose>
						
					</xsl:when>
					<xsl:when test="@style:leader-style and not(@style:leader-text)">
						<xsl:choose>
							<xsl:when test="@style:leader-style='dotted'">
								<xsl:value-of select="'dot'"/>
							</xsl:when>
							<xsl:when test="@style:leader-style='solid'">
								<xsl:value-of select="'heavy'"/>
							</xsl:when>
							<xsl:when test="@style:leader-style='dash'">
								<xsl:value-of select="'hyphen'"/>
							</xsl:when>
							<xsl:when test="@style:leader-style='dotted'">
								<xsl:value-of select="'dot'"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="'none'"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="'none'"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:attribute>
		</w:tab>		
	</xsl:template>



	<!-- ignored -->
	<xsl:template match="text()" mode="styles"/>
	<!--xsl:template match="draw:marker" mode="styles"/>
	<xsl:template match="text:outline-style" mode="styles"/>
	<xsl:template match="text:list-style" mode="styles"/>
	<xsl:template match="text:footnotes-configuration" mode="styles"/>
	<xsl:template match="text:endnotes-configuration" mode="styles"/>
	<xsl:template match="text:linenumbering-configuration" mode="styles"/>
	<xsl:template match="style:page-master" mode="styles"/>
	<xsl:template match="number:date-style" mode="styles"/-->

	<!-- Map font types -->
	<xsl:template name="computeFontName">
		<xsl:param name="fontName"/>
		<xsl:choose>
			<xsl:when test="$fontName = 'StarSymbol'">Symbol</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$fontName"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

</xsl:stylesheet>

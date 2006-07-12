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
<xsl:stylesheet version="1.0"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/3/main"
	xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
	xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
	xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
	xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
	xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
	xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
	exclude-result-prefixes="office style fo text draw number">
		
	
	<xsl:template name="styles">
		<w:styles>
			<w:docDefaults>
				<!-- Default text properties -->
				<xsl:variable name="paragraphDefaultStyle" select="document('styles.xml')/office:document-styles/office:styles/style:default-style[@style:family='paragraph']"/>
				<w:rPrDefault>
					<xsl:apply-templates select="$paragraphDefaultStyle/style:text-properties" mode="styles"/>
				</w:rPrDefault>
				<!-- Default paragraph properties -->
				<w:pPrDefault>
					<xsl:apply-templates select="$paragraphDefaultStyle/style:paragraph-properties" mode="styles"/>
				</w:pPrDefault>	
			</w:docDefaults>
			<xsl:apply-templates select="document('styles.xml')/office:document-styles/office:styles" mode="styles"/>
			<xsl:apply-templates select="document('content.xml')/office:document-content/office:automatic-styles" mode="styles"/>
		</w:styles>
	</xsl:template>
	
	
	
	<xsl:template match="style:style" mode="styles">
		<w:style  w:styleId="{@style:name}" w:customStyle="1">
		
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
				
				<xsl:when test="@style-family = 'table' ">
					<xsl:attribute name="w:type">table</xsl:attribute>
				</xsl:when>
				<xsl:when test="@style-family = 'table-cell' ">
					<xsl:attribute name="w:type">table</xsl:attribute>
				</xsl:when>
				<xsl:when test="@style-family = 'table-row' ">
					<xsl:attribute name="w:type">table</xsl:attribute>
				</xsl:when>
				<xsl:when test="@style-family = 'table-column' ">
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
			
			<xsl:choose>
				<xsl:when test="@style:name">
					<w:name w:val="{@style:name}"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:if test="@style:display-name">
						<w:name w:val="{@style:display-name}"/>
					</xsl:if>
				</xsl:otherwise>
			</xsl:choose>
			
			<!-- Nested elements-->

			<!--xsl:if test="@style:display-name">
				<w:name w:val="{@style:display-name}"/>
			</xsl:if>
			
			<xsl:if test="@style:name">
				<w:name w:val="{@style:name}"/>
			</xsl:if-->
			
			
			
			<xsl:if test="@style:parent-style-name">
				<w:basedOn w:val="{@style:parent-style-name}" />
			</xsl:if>
			<xsl:if test="@style:next-style-name">
				<w:next w:val="{@style:next-style-name}" />
			</xsl:if>
			<w:qFormat/>
			<xsl:if test="name(parent::*) = 'office:automatic-styles'">
				<!--w:semiHidden/-->
				<!--w:hidden/-->
			</xsl:if>
			<xsl:apply-templates mode="styles"/>
		</w:style>
	</xsl:template>
	
	<xsl:template match="style:default-style" mode="styles">
	</xsl:template>
	
	<xsl:template match="style:paragraph-properties[parent::style:style  or parent::style:default-style]" mode="styles">
		<w:pPr>
			<!-- background color -->
			<xsl:if test="@fo:background-color and (@fo:background-color != 'transparent')">
				<w:shd>
					<xsl:attribute name="w:val">solid</xsl:attribute>
					<xsl:attribute name="w:fill"><xsl:value-of select="substring(@fo:background-color, 2, string-length(@fo:background-color) -1)"/></xsl:attribute>
				</w:shd>
			</xsl:if>
			<!-- tabs -->
			<xsl:if test="style:tab-stops/style:tab-stop">
				<w:tabs>
				<xsl:for-each select="style:tab-stops/style:tab-stop">
					<w:tab>
						<xsl:attribute name="w:pos">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length" select="@style:position"/>
							</xsl:call-template>
						</xsl:attribute>					
						<xsl:if test="@style:type">
							<xsl:attribute name="w:val"><xsl:value-of select="@style:type"/></xsl:attribute>
						</xsl:if>
						<!-- Default value -->
						<xsl:if test="not(@style:type)">
							<xsl:attribute name="w:val">clear</xsl:attribute>
						</xsl:if>
					</w:tab>
				</xsl:for-each>
				</w:tabs>
			</xsl:if>
			<xsl:if test="@style:line-height-at-least or fo:line-height or @fo:margin-bottom or @fo:margin-top">
				<w:spacing>
					<xsl:if test="@style:line-height-at-least">
						<xsl:attribute name="w:lineRule">atLeast</xsl:attribute>
						<xsl:attribute name="w:line">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length" select="@style:line-height-at-least"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="contains(@fo:line-height, '%')">
						<xsl:attribute name="w:lineRule">auto</xsl:attribute>
						<xsl:attribute name="w:line"><xsl:value-of select="round(number(substring-before(@fo:line-height, '%')) * 240 div 100)"/></xsl:attribute>						
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
			<xsl:if test="@fo:margin-left or @fo:margin-right or @fo:text-indent or @text:space-before">
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
					<xsl:if test="@fo:text-indent">
						<xsl:attribute name="w:firstLine">
							<xsl:call-template name="twips-measure">
								<xsl:with-param name="length" select="@fo:text-indent"/>
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
			
			<!--
			<xsl:if test="@fo:break-before='page'">
				<w:pageBreakBefore val="on"/>
			</xsl:if>
			-->
				
		<!-- TODO Manage Bidi but it is not prioritary-->
		</w:pPr>
	</xsl:template>
	
	<xsl:template match="style:text-properties[parent::style:style  or parent::style:default-style]" mode="styles">
		<w:rPr>
		
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
		
			<!-- TODO make sure this is the right method -->
			<!-- TODO manage asian and others fonts -->
			<xsl:if test="@style:font-name">
				<w:rFonts>
					<xsl:if test="@style:font-name">
						<xsl:attribute name="w:ascii"><xsl:call-template name="computeFontName"><xsl:with-param name="fontName" select="@style:font-name"/></xsl:call-template></xsl:attribute>
						<xsl:attribute name="w:hAnsi"><xsl:call-template name="computeFontName"><xsl:with-param name="fontName" select="@style:font-name"/></xsl:call-template></xsl:attribute>
					</xsl:if>
					<xsl:if test="@style:font-name-complex">
						<xsl:attribute name="w:cs"><xsl:call-template name="computeFontName"><xsl:with-param name="fontName" select="@style:font-name-complex"/></xsl:call-template></xsl:attribute>
					</xsl:if>
					<xsl:if test="@style:font-name-asian">
						<xsl:attribute name="w:eastAsia"><xsl:call-template name="computeFontName"><xsl:with-param name="fontName" select="@style:font-name-asian"/></xsl:call-template></xsl:attribute>
					</xsl:if>
				</w:rFonts>
			</xsl:if>
		
			
			<xsl:if test="@fo:font-size-complex">
				<w:sz-cs>
					<xsl:attribute name="w:val"><xsl:value-of select="number(substring-before(@fo:font-size-complex, 'pt')) * 2"/></xsl:attribute>
				</w:sz-cs>
			</xsl:if>
			
			<!-- TODO determine all the different font styles-->
			<xsl:if test="@fo:font-style">
				<xsl:choose>
					<xsl:when test="@fo:font-style = 'italic'">
						<w:i w:val="on"/>						
					</xsl:when>
					<xsl:when test="@fo:font-style = 'none'">
						<w:i w:val="off"/>
					</xsl:when>
					<xsl:otherwise>
						<w:i w:val="off"/>
					</xsl:otherwise>
					<!-- It could be also oblique in fo DTD, but it is not possible to set it via Ooo interface -->
				</xsl:choose>
			</xsl:if>
			<xsl:if test="@fo:font-style-complex">
				<xsl:choose>
					<xsl:when test="@fo:font-style-complex = 'italic'">
						<w:i-cs/>						
					</xsl:when>
				</xsl:choose>
			</xsl:if>
			<xsl:if test="@fo:font-weight">
				<xsl:choose>
					<xsl:when test="@fo:font-weight = 'bold'">
						<w:b w:val="on"/>						
					</xsl:when>
					<xsl:when test="@fo:font-weight = 'normal'">
						<w:b w:val="off"/>
					</xsl:when>
				</xsl:choose>
			</xsl:if>
			<xsl:if test="@fo:font-weight-complex">
				<xsl:choose>
					<xsl:when test="@fo:font-weight-complex = 'bold'">
						<w:b-cs/>
					</xsl:when>
				</xsl:choose>
			</xsl:if>
			<xsl:if test="@fo:color">
				<w:color>
					<xsl:attribute name="w:val"><xsl:value-of select="substring(@fo:color, 2, string-length(@fo:color) -1)"/></xsl:attribute>
				</w:color>
			</xsl:if>
			
			<xsl:if test="@style:text-underline-style != 'none' ">
				<w:u>
					<xsl:attribute name="w:val">
						<xsl:choose>
							<xsl:when test="@style:text-underline-style = 'dotted'">
								<xsl:choose>
									<xsl:when test="@style:text-underline-width = 'thick' ">dottedHeavy</xsl:when>
									<xsl:when test="@style:text-underline-width = 'bold' ">dottedHeavy</xsl:when>
									<xsl:otherwise>dotted</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="@style:text-underline-style = 'dash' ">
								<xsl:choose>
									<xsl:when test="@style:text-underline-width = 'thick' ">dashedHeavy</xsl:when>
									<xsl:when test="@style:text-underline-width = 'bold' ">dashedHeavy</xsl:when>
									<xsl:otherwise>dash</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="@style:text-underline-style = 'long-dash'">
								<xsl:choose>
									<xsl:when test="@style:text-underline-width = 'thick' ">dashLongHeavy</xsl:when>
									<xsl:when test="@style:text-underline-width = 'bold' ">dashLongHeavy</xsl:when>
									<xsl:otherwise>dashLong</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="@style:text-underline-style = 'dot-dash' ">
								<xsl:choose>
									<xsl:when test="@style:text-underline-width = 'thick' ">dashDotHeavy</xsl:when>
									<xsl:when test="@style:text-underline-width = 'bold' ">dashDotHeavy</xsl:when>
									<xsl:otherwise>dotDash</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="@style:text-underline-style = 'dot-dot-dash' ">
								<xsl:choose>
									<xsl:when test="@style:text-underline-width = 'thick' ">dashDotDotHeavy</xsl:when>
									<xsl:when test="@style:text-underline-width = 'bold' ">dashDotDotHeavy</xsl:when>
									<xsl:otherwise>dotDotDash</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:when test="@style:text-underline-style = 'wave' ">
								<xsl:choose>
									<xsl:when test="@style:text-underline-type = 'double' ">wavyDouble</xsl:when>
									<xsl:when test="@style:text-underline-width = 'thick' ">wavyHeavy</xsl:when>
									<xsl:when test="@style:text-underline-width = 'bold' ">wavyHeavy</xsl:when>
									<xsl:otherwise>wave</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:otherwise>
								<xsl:choose>
									<xsl:when test="@style:text-underline-type = 'double' ">double</xsl:when>
									<xsl:when test="@style:text-underline-width = 'thick' ">thick</xsl:when>
									<xsl:when test="@style:text-underline-width = 'bold' ">thick</xsl:when>
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
									<xsl:value-of select="substring(@style:text-underline-color, 2, string-length(@style:text-underline-color)-1)"/>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</xsl:if>
				</w:u>
			</xsl:if>
		</w:rPr>		
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
					<xsl:with-param name="node" select="/office:document-styles/office:styles/style:style[@style:name = $parentStyleName]/style:text-properties"/>
				</xsl:call-template>
			</xsl:variable>
			<xsl:choose>
				<xsl:when test="number($value)">
					<xsl:value-of select="round(number(substring-before($node/@fo:font-size, '%')) div 100 * number($value))"/>
				</xsl:when>
				<xsl:otherwise> <!-- fetch the default font size for this style family -->
					<xsl:variable name="defaultProps" select="/office:document-styles/office:styles/style:default-style[@style:family=$node/../@style:family]/style:text-properties"/>
					<xsl:variable name="defaultValue" select="number(substring-before($defaultProps/@fo:font-size, 'pt'))*2"/>
					<xsl:value-of select="round(number(substring-before($node/@fo:font-size, '%')) div 100 * number($defaultValue))"/>
				</xsl:otherwise>
			</xsl:choose>
			
		</xsl:if>
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

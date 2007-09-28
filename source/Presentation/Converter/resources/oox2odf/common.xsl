<?xml version="1.0" encoding="utf-8"?>
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
xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
xmlns:xlink="http://www.w3.org/1999/xlink"
xmlns:dom="http://www.w3.org/2001/xml-events" 
xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
xmlns:rels="http://schemas.openxmlformats.org/package/2006/relationships"
xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
xmlns:dc="http://purl.org/dc/elements/1.1/"
xmlns:dcterms="http://purl.org/dc/terms/"
exclude-result-prefixes="p a r xlink ">

	<!-- Shape constants-->
	<!-- Arrow size -->
	<xsl:variable name="sm-sm">
		<xsl:value-of select ="'0.14'"/>
	</xsl:variable>
	<xsl:variable name="sm-med">
		<xsl:value-of select ="'0.245'"/>
	</xsl:variable>
	<!--<xsl:variable name="sm-lg">
		<xsl:value-of select ="'0.2'"/>
	</xsl:variable>-->
	<xsl:variable name="med-sm">
		<xsl:value-of select ="'0.234'" />
	</xsl:variable>
	<xsl:variable name="med-med">
		<xsl:value-of select ="'0.351'"/>
	</xsl:variable>
	<!--<xsl:variable name="med-lg">
		<xsl:value-of select ="'0.3'" />
	</xsl:variable>-->
	<xsl:variable name="lg-sm">
		<xsl:value-of select ="'0.31'" />
	</xsl:variable>
	<xsl:variable name="lg-med">
		<xsl:value-of select ="'0.35'" />
	</xsl:variable>
	<!--<xsl:variable name="lg-lg">
		<xsl:value-of select ="'0.4'" />
	</xsl:variable>-->
	<!-- Get line styles for shape -->
	<xsl:template name ="tmpLineStyle">
    <xsl:param name="ThemeName"/>
    <xsl:message terminate="no">progress:a:p</xsl:message>
		<!-- Line width-->
		<xsl:for-each select ="p:spPr">
			<xsl:if test ="a:ln/@w">
				<xsl:attribute name ="svg:stroke-width">
					<xsl:call-template name="ConvertEmu">
						<xsl:with-param name="length" select="a:ln/@w"/>
						<xsl:with-param name="unit">cm</xsl:with-param>
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>
			<!--<xsl:if test ="not(a:ln/@w) and (parent::node()/p:nvSpPr/p:cNvPr/@name[contains(., 'Text')])">
				<xsl:attribute name ="draw:stroke">
					<xsl:value-of select ="'none'"/>
				</xsl:attribute>
			</xsl:if>-->
			<!--Bug fix for default BlueBorder in textbox by Mathi on 26thAug 2007-->
			<xsl:if test="(not(a:ln) and not(parent::node()/p:nvSpPr/p:cNvSpPr/@txBox) and not(parent::node()/p:nvSpPr/p:nvPr/p:ph) and not(parent::node()/p:style)) or 
		                  (not(a:ln/@w) and (parent::node()/p:nvSpPr/p:cNvPr/@name[contains(., 'Content Placeholder')]) and not(parent::node()/p:nvSpPr/p:nvPr/p:ph) and not(parent::node()/p:style)) or
						  ((parent::node()/p:nvSpPr/p:cNvPr/@name[contains(., 'Title ')]) and (parent::node()/p:nvSpPr/p:cNvSpPr/@txBox) and not(parent::node()/p:nvSpPr/p:nvPr/p:ph) and not(parent::node()/p:style))">
				<xsl:attribute name ="draw:stroke">
					<xsl:value-of select ="'none'"/>
				</xsl:attribute>
			</xsl:if>
			<!--End of Code-->
		</xsl:for-each>

    <xsl:choose>
      <xsl:when test ="p:style/a:lnRef/@idx and not(p:spPr/a:ln/@w)">
        <xsl:variable name="idx" select="p:style/a:lnRef/@idx"/>
        <xsl:for-each select ="document($ThemeName)//a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln">
          <xsl:if test="position()=$idx">
            <xsl:attribute name ="svg:stroke-width">
              <xsl:value-of select="concat(format-number(./@w div 360000, '#.#####'), 'cm')"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="not(p:spPr/a:ln/@w) and p:spPr/a:ln">
        <xsl:for-each select ="document($ThemeName)//a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln">
          <xsl:if test="position()=1">
            <xsl:attribute name ="svg:stroke-width">
              <xsl:value-of select="concat(format-number(./@w div 360000, '#.#####'), 'cm')"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
    </xsl:choose>
		<!-- Line Dash property-->
		<xsl:for-each select ="p:spPr/a:ln">
			<xsl:if test ="not(a:noFill)">
				<xsl:choose>
					<xsl:when test ="(a:prstDash/@val='solid') or not(a:prstDash/@val)">
						<!--<xsl:attribute name ="draw:stroke">
							<xsl:value-of select ="'solid'"/>
						</xsl:attribute>-->
					</xsl:when>
					<xsl:otherwise>
						<xsl:attribute name ="draw:stroke">
							<xsl:value-of select ="'dash'"/>
						</xsl:attribute>
						<xsl:attribute name ="draw:stroke-dash">
							<xsl:choose>
								<xsl:when test="(a:prstDash/@val='sysDot') and (@cap='rnd')">
									<xsl:value-of select ="'sysDotRound'"/>
								</xsl:when>
								<xsl:when test="a:prstDash/@val='sysDot'">
									<xsl:value-of select ="'sysDot'"/>
								</xsl:when>
								<xsl:when test ="(a:prstDash/@val='sysDash') and (@cap='rnd')">
									<xsl:value-of select ="'sysDashRound'"/>
								</xsl:when>
								<xsl:when test ="a:prstDash/@val='sysDash'">
									<xsl:value-of select ="'sysDash'"/>
								</xsl:when>
								<xsl:when test ="(a:prstDash/@val='dash') and (@cap='rnd')">
									<xsl:value-of select ="'dashRound'"/>
								</xsl:when>
								<xsl:when test ="a:prstDash/@val='dash'">
									<xsl:value-of select ="'dash'"/>
								</xsl:when>
								<xsl:when test ="(a:prstDash/@val='dashDot') and (@cap='rnd')">
									<xsl:value-of select ="'dashDotRound'"/>
								</xsl:when>
								<xsl:when test ="a:prstDash/@val='dashDot'">
									<xsl:value-of select ="'dashDot'"/>
								</xsl:when>
								<xsl:when test ="(a:prstDash/@val='lgDash') and (@cap='rnd')">
									<xsl:value-of select ="'lgDashRound'"/>
								</xsl:when>
								<xsl:when test ="a:prstDash/@val='lgDash'">
									<xsl:value-of select ="'lgDash'"/>
								</xsl:when>
								<xsl:when test ="(a:prstDash/@val='lgDashDot') and (@cap='rnd')">
									<xsl:value-of select ="'lgDashDotRound'"/>
								</xsl:when>
								<xsl:when test ="a:prstDash/@val='lgDashDot'">
									<xsl:value-of select ="'lgDashDot'"/>
								</xsl:when>
								<xsl:when test ="(a:prstDash/@val='lgDashDotDot') and (@cap='rnd')">
									<xsl:value-of select ="'lgDashDotDotRound'"/>
								</xsl:when>
								<xsl:when test ="a:prstDash/@val='lgDashDotDot'">
									<xsl:value-of select ="'lgDashDotDot'"/>
								</xsl:when>
							</xsl:choose>
						</xsl:attribute>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:if>
		</xsl:for-each >

		<!-- Line join property -->
		<xsl:choose>
			<xsl:when test ="p:spPr/a:ln/a:miter">
				<xsl:attribute name ="draw:stroke-linejoin">
					<xsl:value-of select ="'miter'"/>
				</xsl:attribute>
			</xsl:when>
			<xsl:when test ="p:spPr/a:ln/a:bevel">
				<xsl:attribute name ="draw:stroke-linejoin">
					<xsl:value-of select ="'bevel'"/>
				</xsl:attribute>
			</xsl:when>
			<xsl:when test ="p:spPr/a:ln/a:round">
				<xsl:attribute name ="draw:stroke-linejoin">
					<xsl:value-of select ="'round'"/>
				</xsl:attribute>
			</xsl:when>
		</xsl:choose>

		<!-- Line Arrow -->
		<!-- Head End-->
		<xsl:for-each select ="p:spPr/a:ln/a:headEnd">
			<xsl:if test ="@type">
				<xsl:attribute name ="draw:marker-start">
					<xsl:value-of select ="@type"/>
				</xsl:attribute>
				<xsl:variable name="lnw">
					<xsl:call-template name="ConvertEmu">
						<xsl:with-param name="length" select="parent::node()/parent::node()/a:ln/@w"/>
						<xsl:with-param name="unit">cm</xsl:with-param>
					</xsl:call-template>
				</xsl:variable>
				<xsl:attribute name ="draw:marker-start-width">
					<xsl:call-template name ="getArrowSize">
						<xsl:with-param name ="pptlw" select ="parent::node()/parent::node()/a:ln/@w" />
						<xsl:with-param name ="lw" select ="substring-before($lnw,'cm')" />
						<xsl:with-param name ="type" select ="@type" />
						<xsl:with-param name ="w" select ="@w" />
						<xsl:with-param name ="len" select ="@len" />
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>
		</xsl:for-each>


		<!-- Tail End-->
		<xsl:for-each select ="p:spPr/a:ln/a:tailEnd">
			<xsl:if test ="@type">
				<xsl:attribute name ="draw:marker-end">
					<xsl:value-of select ="@type"/>
				</xsl:attribute>
				<xsl:variable name="lnw">
					<xsl:call-template name="ConvertEmu">
						<xsl:with-param name="length" select="parent::node()/parent::node()/a:ln/@w"/>
						<xsl:with-param name="unit">cm</xsl:with-param>
					</xsl:call-template>
				</xsl:variable>
				<xsl:attribute name ="draw:marker-end-width">
					<xsl:call-template name ="getArrowSize">
						<xsl:with-param name ="pptlw" select ="parent::node()/parent::node()/a:ln/@w" />
						<xsl:with-param name ="lw" select ="substring-before($lnw,'cm')" />
						<xsl:with-param name ="type" select ="@type" />
						<xsl:with-param name ="w" select ="@w" />
						<xsl:with-param name ="len" select ="@len" />
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>
		</xsl:for-each>
	</xsl:template>

	<!-- Get arrow size -->
	<xsl:template name ="getArrowSize">
		<xsl:param name ="pptlw" />
		<xsl:param name ="lw" />
		<xsl:param name ="type" />
		<xsl:param name ="w" />
		<xsl:param name ="len" />
		<xsl:choose>
			<!-- selection for (top row arrow type) and (arrow size as 'arrow') -->
			<xsl:when test ="(($pptlw &gt; '24765') and ($type = 'arrow') and ($w='sm') and ($len='med')) or (($pptlw &gt; '24765') and ($type = 'arrow') and ($w='sm') and ($len='sm')) or (($pptlw &gt; '24765') and ($type = 'arrow') and ($w='sm') and ($len='lg'))">
				<xsl:value-of select ="concat(($lw) * (3.5),'cm')"/>
			</xsl:when>
			<xsl:when test ="(($pptlw &lt; '24766')  and ($type = 'arrow') and ($w='sm') and ($len='med')) or (($pptlw &lt; '24766')  and ($type = 'arrow') and ($w='sm') and ($len='sm')) or (($pptlw &lt; '24766')  and ($type = 'arrow') and ($w='sm') and ($len='lg'))">
				<xsl:value-of select ="concat($sm-med,'cm')"/>
			</xsl:when>
			<!-- selection for (top row arrow type) and non-selection of (arrow size as 'arrow') -->
			<xsl:when test ="(($pptlw &gt; '25400') and ($w='sm') and ($len='med')) or (($pptlw &gt; '25400') and ($w='sm') and ($len='sm')) or (($pptlw &gt; '25400') and ($w='sm') and ($len='lg'))">
				<xsl:value-of select ="concat(($lw) * (2),'cm')"/>
			</xsl:when>
			<xsl:when test ="(($pptlw &lt; '25401')  and ($w='sm') and ($len='med')) or (($pptlw &lt; '25401')  and ($w='sm') and ($len='sm')) or (($pptlw &lt; '25401')  and ($w='sm') and ($len='lg'))">
				<xsl:value-of select ="concat($sm-sm,'cm')"/>
			</xsl:when>

			<!-- selection for (middle row arrow type) and (arrow size as 'arrow') -->
			<xsl:when test ="(($pptlw &gt; '28575') and ($type = 'arrow')) or (($pptlw &gt; '28575') and ($type = 'arrow') and ($w='med') and ($len='sm')) or (($pptlw &gt; '28575') and ($type = 'arrow') and ($w='med') and ($len='med')) or (($pptlw &gt; '28575') and ($type = 'arrow') and ($w='med') and ($len='lg'))">
				<xsl:value-of select ="concat(($lw) * (4.5),'cm')"/>
			</xsl:when>
			<xsl:when test ="(($pptlw &lt; '28576') and ($type = 'arrow')) or (($pptlw &lt; '28576')  and ($type = 'arrow') and ($w='med') and ($len='sm')) or (($pptlw &lt; '28576')  and ($type = 'arrow') and ($w='med') and ($len='med')) or (($pptlw &lt; '28576')  and ($type = 'arrow') and ($w='med') and ($len='lg'))">
				<xsl:value-of select ="concat($med-med,'cm')"/>
			</xsl:when>
			<!-- selection for (middle row arrow type) and non-selection of (arrow size as 'arrow') -->
			<xsl:when test ="(($pptlw &gt; '28575')) or (($pptlw &gt; '28575') and ($w='med') and ($len='sm')) or (($pptlw &gt; '28575') and ($w='med') and ($len='med')) or (($pptlw &gt; '28575') and ($w='med') and ($len='lg'))">
				<xsl:value-of select ="concat(($lw) * (3),'cm')"/>
			</xsl:when>
			<xsl:when test ="(($pptlw &lt; '28576')) or (($pptlw &lt; '28576') and ($w='med') and ($len='sm')) or (($pptlw &lt; '28576') and ($w='med') and ($len='med')) or (($pptlw &lt; '28576') and ($w='med') and ($len='lg'))">
				<xsl:value-of select ="concat($med-sm,'cm')"/>
			</xsl:when>

			<!-- selection for (bottom row arrow type) and (arrow size as 'arrow') -->
			<xsl:when test ="(($pptlw &gt; '24765') and ($type = 'arrow') and ($w='lg') and ($len='med')) or  (($pptlw &gt; '24765') and ($type = 'arrow') and ($w='lg') and ($len='sm')) or (($pptlw &gt; '24765') and ($type = 'arrow') and ($w='lg') and ($len='lg'))">
				<xsl:value-of select ="concat(($lw) * (6),'cm')"/>
			</xsl:when>
			<xsl:when test ="(($pptlw &lt; '24766')  and ($type = 'arrow') and ($w='lg') and ($len='med')) or (($pptlw &lt; '24766') and ($type = 'arrow') and ($w='lg') and ($len='sm')) or (($pptlw &lt; '24766')  and ($type = 'arrow') and ($w='lg') and ($len='lg'))">
				<xsl:value-of select ="concat($lg-med,'cm')"/>
			</xsl:when>
			<!-- selection for (bottom row arrow type) and non-selection of (arrow size as 'arrow') -->
			<xsl:when test ="(($pptlw &gt; '25400') and ($w='lg') and ($len='med')) or (($pptlw &gt; '25400') and ($w='lg') and ($len='sm')) or (($pptlw &gt; '25400') and ($w='lg') and ($len='lg'))">
				<xsl:value-of select ="concat(($lw) * (5),'cm')"/>
			</xsl:when>
			<xsl:when test ="(($pptlw &lt; '25401')  and ($w='lg') and ($len='med')) or (($pptlw &lt; '25401')  and ($w='lg') and ($len='sm')) or (($pptlw &lt; '25401')  and ($w='lg') and ($len='lg'))">
				<xsl:value-of select ="concat($lg-sm,'cm')"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select ="concat($med-med,'cm')"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	
	<xsl:template name="getColorCode">
		<xsl:param name="color"/>
		<xsl:param name ="lumMod"/>
		<xsl:param name ="lumOff"/>
    <xsl:message terminate="no">progress:a:p</xsl:message>
		<xsl:message terminate="no">progress:p:cSld</xsl:message>
		<xsl:variable name ="ThemeColor">
			<xsl:for-each select ="document('ppt/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme">
				<xsl:for-each select ="node()">
					<xsl:if test ="name() =concat('a:',$color)">
						<xsl:choose >
							<xsl:when test ="contains(node()/@val,'window') ">
								<xsl:value-of select ="node()/@lastClr"/>
							</xsl:when>
							<xsl:otherwise >
								<xsl:value-of select ="node()/@val"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:if>
				</xsl:for-each>
			</xsl:for-each>			
		</xsl:variable>
		<xsl:variable name ="BgTxColors">
			<xsl:if test ="$color ='bg2'">
				<xsl:value-of select ="document('ppt/theme/theme1.xml') //a:lt2/a:srgbClr/@val"/>
			</xsl:if>
			<xsl:if test ="$color ='bg1'">
				<!--<xsl:value-of select ="document('ppt/theme/theme1.xml') //a:lt1/node()/@lastClr" />-->
				<xsl:choose>
					<xsl:when test ="document('ppt/theme/theme1.xml') //a:lt1/a:srgbClr/@val">
						<xsl:value-of select ="document('ppt/theme/theme1.xml') //a:lt1/a:srgbClr/@val" />
					</xsl:when>
					<xsl:when test="document('ppt/theme/theme1.xml') //a:lt1/node()/@lastClr">
						<xsl:value-of select ="document('ppt/theme/theme1.xml') //a:lt1/node()/@lastClr" />
					</xsl:when>
				</xsl:choose>
			</xsl:if>
			<!--<xsl:if test ="$color ='tx1'">
				<xsl:value-of select ="document('ppt/theme/theme1.xml') //a:dk1/node()/@lastClr"/>
			</xsl:if>-->
			<xsl:if test ="$color ='tx1'">
				<xsl:choose>
					<xsl:when test ="document('ppt/theme/theme1.xml') //a:dk1/node()/@lastClr">
					<xsl:value-of select ="document('ppt/theme/theme1.xml') //a:dk1/node()/@lastClr"/>
					</xsl:when>
				<xsl:when test ="document('ppt/theme/theme1.xml') //a:dk1/node()/@val">
					<xsl:value-of select ="document('ppt/theme/theme1.xml') //a:dk1/node()/@val"/>
				</xsl:when>
				</xsl:choose>
			</xsl:if>
			<xsl:if test ="$color ='tx2'">
				<xsl:value-of select ="document('ppt/theme/theme1.xml') //a:dk2/a:srgbClr/@val"/>
			</xsl:if>			
		</xsl:variable>
		<xsl:variable name ="NewColor">
			<xsl:if test ="$ThemeColor != ''">
				<xsl:value-of select ="$ThemeColor"/>
			</xsl:if>
			<xsl:if test ="$BgTxColors !=''">
				<xsl:value-of select ="$BgTxColors"/>
			</xsl:if>
		</xsl:variable>
		<xsl:call-template name ="ConverThemeColor">
			<xsl:with-param name="color" select="$NewColor" />
			<xsl:with-param name ="lumMod" select ="$lumMod"/>
			<xsl:with-param name ="lumOff" select ="$lumOff"/>
		</xsl:call-template>
	</xsl:template >
	<xsl:template name="ConvertEmu">
		<xsl:param name="length" />
		<xsl:param name="unit" />
		<xsl:message terminate="no">progress:p:cSld</xsl:message>
		<xsl:choose>
			<xsl:when test="$length = '' or not($length) or $length = 0 or format-number($length div 360000, '#.##') = ''">
				<xsl:value-of select="concat(0,'cm')" />
			</xsl:when>
			<xsl:when test="$unit = 'cm'">
				<xsl:value-of select="concat(format-number($length div 360000, '#.##'), 'cm')" />
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="InsertWhiteSpaces">
		<xsl:param name="string" select="."/>
		<xsl:param name="length" select="string-length(.)"/>
		<xsl:message terminate="no">progress:p:cSld</xsl:message>
		<!-- string which doesn't contain whitespaces-->
		<xsl:choose>
			<xsl:when test="not(contains($string,' '))">
				<xsl:value-of select="$string"/>
			</xsl:when>
			<!-- convert white spaces  -->
			<xsl:otherwise>
				<xsl:variable name="before">
					<xsl:value-of select="substring-before($string,' ')"/>
				</xsl:variable>
				<xsl:variable name="after">
					<xsl:call-template name="CutStartSpaces">
						<xsl:with-param name="cuted">
							<xsl:value-of select="substring-after($string,' ')"/>
						</xsl:with-param>
					</xsl:call-template>
				</xsl:variable>
				<xsl:if test="$before != '' ">
					<xsl:value-of select="concat($before,' ')"/>
				</xsl:if>
				<!--add remaining whitespaces as text:s if there are any-->
				<xsl:if test="string-length(concat($before,' ', $after)) &lt; $length ">
					<xsl:choose>
						<xsl:when test="($length - string-length(concat($before, $after))) = 1">
							<text:s/>
						</xsl:when>
						<xsl:otherwise>
							<text:s>
								<xsl:attribute name="text:c">
									<xsl:choose>
										<xsl:when test="$before = ''">
											<xsl:value-of select="$length - string-length($after)"/>
										</xsl:when>
										<xsl:otherwise>
											<xsl:value-of select="$length - string-length(concat($before,' ', $after))"/>
										</xsl:otherwise>
									</xsl:choose>
								</xsl:attribute>
							</text:s>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:if>
				<!--repeat it for substring which has whitespaces-->
				<xsl:if test="contains($string,' ') and $length &gt; 0">
					<xsl:call-template name="InsertWhiteSpaces">
						<xsl:with-param name="string">
							<xsl:value-of select="$after"/>
						</xsl:with-param>
						<xsl:with-param name="length">
							<xsl:value-of select="string-length($after)"/>
						</xsl:with-param>
					</xsl:call-template>
				</xsl:if>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="CutStartSpaces">
		<xsl:param name="cuted"/>
		<xsl:choose>
			<xsl:when test="starts-with($cuted,' ')">
				<xsl:call-template name="CutStartSpaces">
					<xsl:with-param name="cuted">
						<xsl:value-of select="substring-after($cuted,' ')"/>
					</xsl:with-param>
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$cuted"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="ConverThemeColor">
		<xsl:param name="color"/>
		<xsl:param name ="lumMod"/>
		<xsl:param name ="lumOff"/>
		<xsl:message terminate="no">progress:p:cSld</xsl:message>
		<xsl:variable name ="Red">
			<xsl:call-template name ="HexToDec">
				<xsl:with-param name ="number" select ="substring($color,1,2)"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name ="Green">
			<xsl:call-template name ="HexToDec">
				<xsl:with-param name ="number" select ="substring($color,3,2)"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name ="Blue">
			<xsl:call-template name ="HexToDec">
				<xsl:with-param name ="number" select ="substring($color,5,2)"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:choose >
			<xsl:when test ="$lumOff = '' and $lumMod != '' ">
				<xsl:variable name ="NewRed">
					<xsl:value-of  select =" floor($Red * $lumMod div 100000) "/>
				</xsl:variable>
				<xsl:variable name ="NewGreen">
					<xsl:value-of  select =" floor($Green * $lumMod div 100000)"/>
				</xsl:variable>
				<xsl:variable name ="NewBlue">
					<xsl:value-of  select =" floor($Blue * $lumMod div 100000)"/>
				</xsl:variable>
				<xsl:call-template name ="CreateRGBColor">
					<xsl:with-param name ="Red" select ="$NewRed"/>
					<xsl:with-param name ="Green" select ="$NewGreen"/>
					<xsl:with-param name ="Blue" select ="$NewBlue"/>
				</xsl:call-template>
			</xsl:when>			
			<xsl:when test ="$lumMod = '' and $lumOff != ''">
				<!-- TBD Not sure whether this condition will occure-->
			</xsl:when>
			<xsl:when test ="$lumMod = '' and $lumOff =''">
				<xsl:value-of  select ="concat('#',$color)"/>
			</xsl:when>
			<xsl:when test ="$lumOff != '' and $lumMod!= '' ">
				<xsl:variable name ="NewRed">
					<xsl:value-of select ="floor(((255 - $Red) * (1 - ($lumMod  div 100000)))+ $Red )"/>
				</xsl:variable>
				<xsl:variable name ="NewGreen">
					<xsl:value-of select ="floor(((255 - $Green) * ($lumOff  div 100000)) + $Green )"/>
				</xsl:variable>
				<xsl:variable name ="NewBlue">
					<xsl:value-of select ="floor(((255 - $Blue) * ($lumOff div 100000)) + $Blue) "/>
				</xsl:variable>
				<xsl:call-template name ="CreateRGBColor">
					<xsl:with-param name ="Red" select ="$NewRed"/>
					<xsl:with-param name ="Green" select ="$NewGreen"/>
					<xsl:with-param name ="Blue" select ="$NewBlue"/>
				</xsl:call-template>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<!-- Converts Decimal values to RGB color -->
	<xsl:template name ="CreateRGBColor">
		<xsl:param name ="Red"/>
		<xsl:param name ="Green"/>
		<xsl:param name ="Blue"/>
		<xsl:message terminate="no">progress:p:cSld</xsl:message>
		<xsl:variable name ="NewRed">
			<xsl:call-template name ="DecToHex">
				<xsl:with-param name ="number" select ="$Red"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name ="NewGreen">
			<xsl:call-template name ="DecToHex">
				<xsl:with-param name ="number" select ="$Green"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name ="NewBlue">
			<xsl:call-template name ="DecToHex">
				<xsl:with-param name ="number" select ="$Blue"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
		<xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
		<xsl:value-of  select ="translate(concat('#',$NewRed,$NewGreen,$NewBlue),ucletters,lcletters)"/>
	</xsl:template>
	<!-- Converts Hexa Decimal Value to Decimal-->
	<xsl:template name="HexToDec">
		<!-- @Description: This is a recurive algorithm converting a hex to decimal -->
		<!-- @Context: None -->
		<xsl:param name="number"/>
		<!-- (string|number) The hex number to convert -->
		<xsl:param name="step" select="0"/>
		<!-- (number) The exponent (only used during convertion)-->
		<xsl:param name="value" select="0"/>
		<!-- (number) The result from the previous digit's convertion (only used during convertion) -->
		<xsl:variable name="number1">
			<!-- translates all letters to lower case -->
			<xsl:value-of select="translate($number,'ABCDEF','abcdef')"/>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="string-length($number1) &gt; 0">
				<xsl:variable name="one">
					<!-- The last digit in the hex number -->
					<xsl:choose>
						<xsl:when test="substring($number1,string-length($number1) ) = 'a'">
							<xsl:text>10</xsl:text>
						</xsl:when>
						<xsl:when test="substring($number1,string-length($number1)) = 'b'">
							<xsl:text>11</xsl:text>
						</xsl:when>
						<xsl:when test="substring($number1,string-length($number1)) = 'c'">
							<xsl:text>12</xsl:text>
						</xsl:when>
						<xsl:when test="substring($number1,string-length($number1)) = 'd'">
							<xsl:text>13</xsl:text>
						</xsl:when>
						<xsl:when test="substring($number1,string-length($number1)) = 'e'">
							<xsl:text>14</xsl:text>
						</xsl:when>
						<xsl:when test="substring($number1,string-length($number1)) = 'f'">
							<xsl:text>15</xsl:text>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="substring($number1,string-length($number1))"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:variable>
				<xsl:variable name="power">
					<!-- The result of the exponent calculation -->
					<xsl:call-template name="Power">
						<xsl:with-param name="base">16</xsl:with-param>
						<xsl:with-param name="exponent">
							<xsl:value-of select="number($step)"/>
						</xsl:with-param>
						<xsl:with-param name="value1">16</xsl:with-param>
					</xsl:call-template>
				</xsl:variable>
				<xsl:choose>
					<xsl:when test="string-length($number1) = 1">
						<xsl:value-of select="($one * $power )+ number($value)"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:call-template name="HexToDec">
							<xsl:with-param name="number">
								<xsl:value-of select="substring($number1,1,string-length($number1) - 1)"/>
							</xsl:with-param>
							<xsl:with-param name="step">
								<xsl:value-of select="number($step) + 1"/>
							</xsl:with-param>
							<xsl:with-param name="value">
								<xsl:value-of select="($one * $power) + number($value)"/>
							</xsl:with-param>
						</xsl:call-template>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<!-- Converts Decimal Value to Hexadecimal-->
	<xsl:template name="DecToHex">
		<xsl:param name="number"/>
		<xsl:variable name="high">
			<xsl:call-template name="HexMap">
				<xsl:with-param name="value">
					<xsl:value-of select="floor($number div 16)"/>
				</xsl:with-param>
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name="low">
			<xsl:call-template name="HexMap">
				<xsl:with-param name="value">
					<xsl:value-of select="$number mod 16"/>
				</xsl:with-param>
			</xsl:call-template>
		</xsl:variable>
		<xsl:value-of select="concat($high,$low)"/>
	</xsl:template>
	<!-- calculates power function -->
	<xsl:template name="HexMap">
		<xsl:param name="value"/>
		<xsl:choose>
			<xsl:when test="$value = 10">
				<xsl:text>A</xsl:text>
			</xsl:when>
			<xsl:when test="$value = 11">
				<xsl:text>B</xsl:text>
			</xsl:when>
			<xsl:when test="$value = 12">
				<xsl:text>C</xsl:text>
			</xsl:when>
			<xsl:when test="$value = 13">
				<xsl:text>D</xsl:text>
			</xsl:when>
			<xsl:when test="$value = 14">
				<xsl:text>E</xsl:text>
			</xsl:when>
			<xsl:when test="$value = 15">
				<xsl:text>F</xsl:text>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$value"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="Power">
		<xsl:param name="base"/>
		<xsl:param name="exponent"/>
		<xsl:param name="value1" select="$base"/>
		<xsl:choose>
			<xsl:when test="$exponent = 0">
				<xsl:text>1</xsl:text>
			</xsl:when>
			<xsl:otherwise>
				<xsl:choose>
					<xsl:when test="$exponent &gt; 1">
						<xsl:call-template name="Power">
							<xsl:with-param name="base">
								<xsl:value-of select="$base"/>
							</xsl:with-param>
							<xsl:with-param name="exponent">
								<xsl:value-of select="$exponent -1"/>
							</xsl:with-param>
							<xsl:with-param name="value1">
								<xsl:value-of select="$value1 * $base"/>
							</xsl:with-param>
						</xsl:call-template>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="$value1"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
  <!-- Added by Vipul-->
  <!--Start-->
  <xsl:template name="tmpWriteCordinates">
    <xsl:message terminate="no">progress:a:p</xsl:message>
	<xsl:message terminate="no">progress:p:cSld</xsl:message>
    <xsl:for-each select ="p:spPr/a:xfrm">
      <xsl:choose>
        <xsl:when test="@rot">
          <xsl:variable name ="xCenter">
            <xsl:value-of select ="a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name ="yCenter">
            <xsl:value-of select ="a:ext/@cy "/>
          </xsl:variable>
          <xsl:variable name ="angle">
            <xsl:if test ="not(@rot)">
              <xsl:value-of select="'0'" />
            </xsl:if>
            <xsl:if test ="@rot">
              <xsl:value-of select ="@rot"/>
            </xsl:if>
          </xsl:variable>
          <xsl:variable name ="xCord">
            <xsl:if test ="a:off/@x">
              <xsl:value-of select ="a:off/@x"/>
            </xsl:if>
            <xsl:if test ="not(a:off/@x) ">
              <xsl:value-of select ="'0'"/>
            </xsl:if>
          </xsl:variable>
          <xsl:variable name ="yCord">
            <xsl:if test ="a:off/@y">
              <xsl:value-of select ="a:off/@y"/>
            </xsl:if>
            <xsl:if test ="not(a:off/@y)">
              <xsl:value-of select ="'0'"/>
            </xsl:if>
          </xsl:variable>
          <xsl:variable name ="var_flipH">
            <xsl:if test ="@flipH">
              <xsl:value-of select ="@flipH"/>
            </xsl:if>
            <xsl:if test ="not(@flipH) ">
              <xsl:value-of select ="'0'"/>
            </xsl:if>
          </xsl:variable>
          <xsl:variable name ="var_flipV">
            <xsl:if test ="@flipV">
              <xsl:value-of select ="@flipV"/>
            </xsl:if>
            <xsl:if test ="not(@flipV)">
              <xsl:value-of select ="'0'"/>
            </xsl:if>
          </xsl:variable>
          <xsl:attribute name ="draw:transform">
            <xsl:value-of select ="concat('draw-transform:',$xCord, ':',$yCord, ':',$xCenter, ':', $yCenter, ':', $var_flipH, ':', $var_flipV, ':', $angle)"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name ="svg:x">
            <xsl:call-template name="ConvertEmu">
              <xsl:with-param name="length" select="a:off/@x"/>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
          <xsl:attribute name ="svg:y">
            <xsl:call-template name="ConvertEmu">
              <xsl:with-param name="length" select="a:off/@y"/>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
      <xsl:attribute name ="svg:width">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="a:ext/@cx"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name ="svg:height">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="a:ext/@cy"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="tmpGropingWriteCordinates">
    <xsl:message terminate="no">progress:p:cSld</xsl:message>
     
    <xsl:for-each select ="p:spPr/a:xfrm">
      <xsl:attribute name ="svg:width">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="((parent::node()/parent::node()/parent::node()/p:grpSpPr/a:xfrm/a:ext/@cx - 
											parent::node()/parent::node()/parent::node()/p:grpSpPr/a:xfrm/a:chExt/@cx)
											+ a:ext/@cx)"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name ="svg:height">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="((parent::node()/parent::node()/parent::node()/p:grpSpPr/a:xfrm/a:ext/@cy -
                  parent::node()/parent::node()/parent::node()/p:grpSpPr/a:xfrm/a:chExt/@cy)+ a:ext/@cy)"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:choose>
        <xsl:when test="@rot">
          <xsl:variable name ="xCenter">
            <xsl:value-of select ="a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name ="yCenter">
            <xsl:value-of select ="a:ext/@cy "/>
          </xsl:variable>
          <xsl:variable name ="angle">
            <xsl:if test ="not(@rot)">
              <xsl:value-of select="'0'" />
            </xsl:if>
            <xsl:if test ="@rot">
              <xsl:value-of select ="@rot"/>
            </xsl:if>
          </xsl:variable>
          <xsl:variable name ="xCord">
            <xsl:if test ="a:off/@x">
              <xsl:value-of select ="a:off/@x"/>
            </xsl:if>
            <xsl:if test ="not(a:off/@x) ">
              <xsl:value-of select ="'0'"/>
            </xsl:if>
          </xsl:variable>
          <xsl:variable name ="yCord">
            <xsl:if test ="a:off/@y">
              <xsl:value-of select ="a:off/@y"/>
            </xsl:if>
            <xsl:if test ="not(a:off/@y)">
              <xsl:value-of select ="'0'"/>
            </xsl:if>
          </xsl:variable>
          <xsl:variable name ="var_flipH">
            <xsl:if test ="@flipH">
              <xsl:value-of select ="@flipH"/>
            </xsl:if>
            <xsl:if test ="not(@flipH) ">
              <xsl:value-of select ="'0'"/>
            </xsl:if>
          </xsl:variable>
          <xsl:variable name ="var_flipV">
            <xsl:if test ="@flipV">
              <xsl:value-of select ="@flipV"/>
            </xsl:if>
            <xsl:if test ="not(@flipV)">
              <xsl:value-of select ="'0'"/>
            </xsl:if>
          </xsl:variable>
          <xsl:attribute name ="draw:transform">
            <xsl:value-of select ="concat('draw-transform:',$xCord, ':',$yCord, ':',$xCenter, ':', $yCenter, ':', $var_flipH, ':', $var_flipV, ':', $angle)"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name ="svg:x">
            <xsl:call-template name="ConvertEmu">
              <xsl:with-param name="length" select="(parent::node()/parent::node()/parent::node()/p:grpSpPr/a:xfrm/a:off/@x - 
											parent::node()/parent::node()/parent::node()/p:grpSpPr/a:xfrm/a:chOff/@x 
											+ a:off/@x)"/>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
      </xsl:attribute>
          <xsl:attribute name ="svg:y">
            <xsl:call-template name="ConvertEmu">
              <xsl:with-param name="length" select="((parent::node()/parent::node()/parent::node()/p:grpSpPr/a:xfrm/a:off/@y - 
											parent::node()/parent::node()/parent::node()/p:grpSpPr/a:xfrm/a:chOff/@y)
											+ a:off/@y)"/>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
     
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="tmpSlideParagraphStyle">
    <xsl:param name="lnSpcReduction"/>
    <xsl:param name="blnSlide"/>
	  <xsl:message terminate="no">progress:p:cSld</xsl:message>
    <!--Text alignment-->
    <xsl:if test ="a:pPr/@algn">
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
          <xsl:when test ="a:pPr/@algn ='l'">
            <xsl:value-of select ="'start'"/>
          </xsl:when>
          <!-- Added by lohith - for fix 1737161-->
          <xsl:when test ="a:pPr/@algn ='just'">
            <xsl:value-of select ="'justify'"/>
          </xsl:when>
          <!-- added by pradeep-->

        </xsl:choose>
      </xsl:attribute>
    </xsl:if >
    <!-- Convert Laeft margin of the paragraph-->
    <xsl:if test ="a:pPr/@marL">
      <xsl:attribute name ="fo:margin-left">
        <xsl:value-of select="concat(format-number(a:pPr/@marL div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="a:pPr/@marR">
      <xsl:attribute name ="fo:margin-right">
        <xsl:value-of select="concat(format-number(a:pPr/@marR div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="a:pPr/@indent">
      <xsl:attribute name ="fo:text-indent">
        <xsl:value-of select="concat(format-number(a:pPr/@indent div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if >
    <xsl:if test ="a:pPr/a:spcBef/a:spcPts/@val">
      <xsl:attribute name ="fo:margin-top">
        <xsl:value-of select="concat(format-number(a:pPr/a:spcBef/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="a:pPr/a:spcAft/a:spcPts/@val">
      <xsl:attribute name ="fo:margin-bottom">
        <xsl:value-of select="concat(format-number(a:pPr/a:spcAft/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <!-- added by Vipul to calculate space before and space after if space pts is in percentage-->
    <!--Start-->
    <xsl:if test ="a:pPr/a:spcBef/a:spcPct/@val">
      <xsl:variable name="var_MaxFntSize">
        <xsl:for-each select="./a:r/a:rPr/@sz">
          <xsl:sort data-type="number" order="descending"/>
          <xsl:if test="position()=1">
            <xsl:value-of select="."/>
          </xsl:if>
        </xsl:for-each>
      </xsl:variable>
      <xsl:if test="$var_MaxFntSize !=''">
      <xsl:attribute name ="fo:margin-top">
        <xsl:value-of select="concat(format-number( (($var_MaxFntSize * ( a:pPr/a:spcBef/a:spcPct/@val div 100000 ) ) div 2835 ) * 1.2, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    </xsl:if>
    <xsl:if test ="a:pPr/a:spcAft/a:spcPct/@val">
      <xsl:variable name="var_MaxFontSize">
        <xsl:for-each select="./a:r/a:rPr/@sz">
          <xsl:sort data-type="number" order="descending"/>
          <xsl:if test="position()=1">
            <xsl:value-of select="."/>
          </xsl:if>
        </xsl:for-each>
      </xsl:variable>
      <xsl:attribute name ="fo:margin-bottom">
        <xsl:value-of select="concat(format-number(  ( ($var_MaxFontSize * ( a:pPr/a:spcAft/a:spcPct/@val div 100000 ) ) div 2835 ) * 1.2, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <!--End-->
    <xsl:if test ="a:pPr/a:lnSpc/a:spcPct/@val">
      <xsl:choose>
        <xsl:when test="$lnSpcReduction='0'">
          <xsl:attribute name="fo:line-height">
            <xsl:value-of select="concat(format-number(a:pPr/a:lnSpc/a:spcPct/@val div 1000,'###'), '%')"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="fo:line-height">
            <xsl:value-of select="concat(format-number((a:pPr/a:lnSpc/a:spcPct/@val - $lnSpcReduction) div 1000,'###'), '%')"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if >
    <xsl:if test ="a:pPr/a:lnSpc/a:spcPts/@val">
      <xsl:attribute name="style:line-height-at-least">
        <xsl:value-of select="concat(format-number(a:pPr/a:lnSpc/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>
  <xsl:template name="tmpSMParagraphStyle">
    <xsl:param name="lnSpcReduction"/>
   
    <!--Text alignment-->
    <xsl:if test ="@algn">
      <xsl:attribute name ="fo:text-align">
        <xsl:choose>
          <!-- Center Alignment-->
          <xsl:when test ="@algn ='ctr'">
            <xsl:value-of select ="'center'"/>
          </xsl:when>
          <!-- Right Alignment-->
          <xsl:when test ="@algn ='r'">
            <xsl:value-of select ="'end'"/>
          </xsl:when>
          <!-- Left Alignment-->
          <xsl:when test ="@algn ='l'">
            <xsl:value-of select ="'start'"/>
          </xsl:when>
          <!-- Added by lohith - for fix 1737161-->
          <xsl:when test ="@algn ='just'">
            <xsl:value-of select ="'justify'"/>
          </xsl:when>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if >
    <!-- Convert Laeft margin of the paragraph-->
    <xsl:if test ="@marL">
      <xsl:attribute name ="fo:margin-left">
        <xsl:value-of select="concat(format-number( @marL div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="@marR">
      <xsl:attribute name ="fo:margin-right">
        <xsl:value-of select="concat(format-number(@marR div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="./@indent">
      <xsl:attribute name ="fo:text-indent">
        <xsl:value-of select="concat(format-number(./@indent div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if >
    <xsl:if test ="./a:defRPr/@sz">
      <xsl:variable name="var_fontsize">
        <xsl:value-of select="./a:defRPr/@sz"/>
      </xsl:variable>
      <xsl:for-each select="./a:spcBef/a:spcPct">
        <xsl:if test ="@val">
          <xsl:attribute name ="fo:margin-top">
            <xsl:value-of select="concat(format-number( (( $var_fontsize * ( @val div 100000 ) ) div 2835) * 1.2  ,'#.###'),'cm')"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:for-each>
      <xsl:for-each select="./a:spcAft/a:spcPct">
        <xsl:if test ="@val">
          <xsl:attribute name ="fo:margin-bottom">
            <xsl:value-of select="concat(format-number( (( $var_fontsize * ( @val div 100000 ) ) div 2835 ) * 1.2 ,'#.###'),'cm')"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:for-each>
    </xsl:if>
    <xsl:for-each select="./a:spcAft/a:spcPts">
      <xsl:if test ="@val">
        <xsl:attribute name ="fo:margin-bottom">
          <xsl:value-of select="concat(format-number(@val div  2835 ,'#.##'),'cm')"/>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
    <xsl:for-each select="./a:spcBef/a:spcPts">
      <xsl:if test ="@val">
        <xsl:attribute name ="fo:margin-top">
          <xsl:value-of select="concat(format-number(@val div  2835 ,'#.##'),'cm')"/>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
    <xsl:for-each select="./a:lnSpc/a:spcPct">
      <xsl:if test ="a:lnSpc/a:spcPct">
        <xsl:choose>
          <xsl:when test="$lnSpcReduction='0'">
            <xsl:attribute name="fo:line-height">
              <xsl:value-of select="concat(format-number(@val div 1000,'###'), '%')"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="fo:line-height">
              <xsl:value-of select="concat(format-number((@val - $lnSpcReduction) div 1000,'###'), '%')"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if >
    </xsl:for-each>
    <xsl:for-each select="./a:lnSpc/a:spcPts">
      <xsl:if test ="a:lnSpc/a:spcPts">
        <xsl:attribute name="style:line-height-at-least">
          <xsl:value-of select="concat(format-number(@val div 2835, '#.##'), 'cm')"/>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="tmpSlideTextProperty">
    <xsl:param name ="DefFont" />
    <xsl:param name ="fontscale"/>
    <xsl:param name="index"/>
	<xsl:message terminate="no">progress:p:cSld</xsl:message>
    <xsl:if test ="a:rPr/@lang">
      <xsl:attribute name ="style:language-asian">
        <xsl:value-of select="substring-before(a:rPr/@lang,'-')"/>
      </xsl:attribute>
      <xsl:attribute name ="style:country-asian">
        <xsl:value-of select="substring-after(a:rPr/@lang,'-')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="a:rPr/@sz">
      <xsl:attribute name ="fo:font-size"	>
        <xsl:for-each select ="a:rPr/@sz">
          <xsl:choose>
            <xsl:when test="$fontscale ='100000'">
              <xsl:value-of  select ="concat(format-number(. div 100,'#.##'),'pt')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of  select ="concat(format-number(round((. *($fontscale div 1000) )div 10000),'#.##'),'pt')"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:attribute>
      <xsl:attribute name ="style:font-size-asian">
        <xsl:for-each select ="a:rPr/@sz">
          <xsl:choose>
            <xsl:when test="$fontscale ='100000'">
              <xsl:value-of  select ="concat(format-number(. div 100,'#.##'),'pt')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of  select ="concat(format-number(round((. *($fontscale div 1000) )div 10000),'#.##'),'pt')"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:attribute>
    </xsl:if>
    <xsl:choose>
      <xsl:when test ="a:rPr/a:latin/@typeface">
        <xsl:attribute name ="fo:font-family">
          <xsl:variable name ="typeFaceVal" select ="a:rPr/a:latin/@typeface"/>
          <xsl:for-each select ="a:rPr/a:latin/@typeface">
            <xsl:if test ="$typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt'">
              <xsl:value-of  select ="$DefFont"/>
            </xsl:if>
            <xsl:if test ="not($typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt')">
              <xsl:value-of select ="."/>
            </xsl:if>
          </xsl:for-each>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="a:rPr/a:cs/@typeface">
        <xsl:attribute name ="fo:font-family">
          <xsl:variable name ="typeFaceVal" select ="a:rPr/a:cs/@typeface"/>
          <xsl:for-each select ="a:rPr/a:cs/@typeface">
            <xsl:if test ="$typeFaceVal='+mn-cs' or $typeFaceVal='+mj-cs'">
              <xsl:value-of  select ="$DefFont"/>
            </xsl:if>
            <xsl:if test ="not($typeFaceVal='+mn-cs' or $typeFaceVal='+mj-cs')">
              <xsl:value-of select ="."/>
            </xsl:if>
          </xsl:for-each>
        </xsl:attribute >
      </xsl:when>
      <xsl:when test ="a:rPr/a:sym/@typeface">
        <xsl:attribute name ="fo:font-family">
          <xsl:variable name ="typeFaceVal" select ="a:rPr/a:sym/@typeface"/>
          <xsl:for-each select ="a:rPr/a:sym/@typeface">
            <xsl:if test ="$typeFaceVal='+mn-sym' or $typeFaceVal='+mj-sym'">
              <xsl:value-of  select ="$DefFont"/>
            </xsl:if>
            <xsl:if test ="not($typeFaceVal='+mn-sym' or $typeFaceVal='+mj-sym')">
              <xsl:value-of select ="."/>
            </xsl:if>
          </xsl:for-each>
        </xsl:attribute >
      </xsl:when>
    </xsl:choose>

    <!-- bug fix 1777569 -->
	
	  <!-- strike style:text-line-through-style-->
    <xsl:if test ="a:rPr/@strike">
		<xsl:if test ="a:rPr/@strike!='noStrike'">
			<xsl:attribute name ="style:text-line-through-style">
				<xsl:value-of select ="'solid'"/>
			</xsl:attribute>
			<xsl:choose >
				<xsl:when test ="a:rPr/@strike='dblStrike'">
					<xsl:attribute name ="style:text-line-through-type">
						<xsl:value-of select ="'double'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test ="a:rPr/@strike='sngStrike'">
					<xsl:attribute name ="style:text-line-through-type">
						<xsl:value-of select ="'single'"/>
					</xsl:attribute>
				</xsl:when>
			</xsl:choose>
		</xsl:if>
    </xsl:if>
    <!-- Kening Property-->
    <xsl:if test ="a:rPr/@kern">
      <xsl:choose >
        <xsl:when test ="a:rPr/@kern = '0'">
          <xsl:attribute name ="style:letter-kerning">
            <xsl:value-of select ="'false'"/>
          </xsl:attribute >
        </xsl:when >
        <xsl:when test ="format-number(a:rPr/@kern,'#.##') &gt; 0">
          <xsl:attribute name ="style:letter-kerning">
            <xsl:value-of select ="'true'"/>
          </xsl:attribute >
        </xsl:when >
      </xsl:choose>
    </xsl:if >
    <!-- Bold Property-->
    <xsl:if test ="a:rPr/@b">
      <xsl:if test ="a:rPr/@b='1'">
        <xsl:attribute name ="fo:font-weight">
          <xsl:value-of select ="'bold'"/>
        </xsl:attribute>
      </xsl:if >
      <xsl:if test ="a:rPr/@b='0'">
        <xsl:attribute name ="fo:font-weight">
          <xsl:value-of select ="'normal'"/>
        </xsl:attribute>
      </xsl:if >
    </xsl:if >
    
    <!--UnderLine-->
    <xsl:if test ="a:rPr/@u">
      <xsl:for-each select ="a:rPr">
        <xsl:call-template name="tmpUnderLine"/>
      </xsl:for-each>
    </xsl:if >
    <!-- Italic-->
    <!-- Fix for the bug 1780908, by vijayeta, date 24th Aug '07-->
    <xsl:if test ="a:rPr/@i">
      <xsl:attribute name ="fo:font-style">
        <xsl:if test ="a:rPr/@i='1'">
          <xsl:value-of select ="'italic'"/>
        </xsl:if >
        <xsl:if test ="a:rPr/@i='0'">
          <!--<xsl:value-of select ="'none'"/>-->
          <xsl:value-of select ="'normal'"/>
        </xsl:if>
      </xsl:attribute>
    </xsl:if >
    <!-- Fix for the bug 1780908, by vijayeta, date 24th Aug '07-->
    <!-- Character Spacing -->
    <xsl:if test ="a:rPr/@spc">
      <xsl:attribute name ="fo:letter-spacing">
        <xsl:variable name="length" select="a:rPr/@spc" />
        <xsl:value-of select="concat(format-number($length * 2.54 div 7200,'#.###'),'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
    <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
    <xsl:if test ="a:rPr/a:solidFill/a:srgbClr/@val">
      <xsl:attribute name ="fo:color">
        <xsl:value-of select ="translate(concat('#',a:rPr/a:solidFill/a:srgbClr/@val),$ucletters,$lcletters)"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="a:rPr/a:solidFill/a:schemeClr/@val">
      <xsl:attribute name ="fo:color">
        <xsl:call-template name ="getColorCode">
          <xsl:with-param name ="color">
            <xsl:value-of select="a:rPr/a:solidFill/a:schemeClr/@val"/>
          </xsl:with-param>
          <xsl:with-param name ="lumMod">
            <xsl:value-of select="a:rPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
          </xsl:with-param>
          <xsl:with-param name ="lumOff">
            <xsl:value-of select="a:rPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="not(a:rPr/a:solidFill/a:srgbClr/@val) and not(a:rPr/a:solidFill/a:schemeClr/@val)">
      <xsl:choose>
        <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
          <xsl:attribute name ="fo:color">
            <xsl:value-of select ="translate(concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val),$ucletters,$lcletters)"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
          <xsl:attribute name ="fo:color">
            <xsl:call-template name ="getColorCode">
              <xsl:with-param name ="color">
                <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumMod">
                <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr/a:lumMod/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumOff">
                <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr/a:lumOff/@val"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
    <!--Shadow fo:text-shadow-->
    <xsl:if test ="a:rPr/a:effectLst/a:outerShdw">
      <xsl:attribute name ="fo:text-shadow">
        <xsl:value-of select ="'1pt 1pt'"/>
      </xsl:attribute>
    </xsl:if>
	  
  <!--SuperScript and SubScript for Text added by Mathi on 31st Jul 2007-->
	  <xsl:if test="(a:rPr/@baseline)">
		  <xsl:variable name="baseData">
			  <xsl:value-of select="a:rPr/@baseline"/>
		  </xsl:variable>
		  <xsl:choose>
			  <xsl:when test="(a:rPr/@baseline &gt; 0)">
				  <xsl:variable name="superCont">
					  <xsl:value-of select="concat('super ',format-number($baseData div 1000,'#.###'),'%')"/>
				  </xsl:variable>
				  <xsl:attribute name="style:text-position">
					  <xsl:value-of select="$superCont"/>
				  </xsl:attribute>
			  </xsl:when>
			  <xsl:when test="(a:rPr/@baseline &lt; 0)">
				  <xsl:variable name="subCont">
					  <xsl:value-of select="concat('sub ',format-number(substring-after($baseData,'-') div 1000,'#.###'),'%')"/>
				  </xsl:variable>
				  <xsl:attribute name="style:text-position">
					  <xsl:value-of select="$subCont"/>
				  </xsl:attribute>
			  </xsl:when>
		  </xsl:choose>
	  </xsl:if>
	  
      </xsl:template>
  <xsl:template name="tmpSlideGrahicProp">
    <xsl:param name="ThemeName"/>
	  <xsl:message terminate="no">progress:p:cSld</xsl:message>
    <xsl:message terminate="no">progress:a:p</xsl:message>
    <!--Background Fill color-->
    <xsl:choose>
      <!-- No fill -->
      <xsl:when test ="p:spPr/a:noFill">
        <xsl:attribute name ="draw:fill">
          <xsl:value-of select="'none'" />
        </xsl:attribute>
      </xsl:when>
      <!-- Solid fill-->
      <!-- Standard color-->
      <xsl:when test ="p:spPr/a:solidFill">
        <xsl:if test ="p:spPr/a:solidFill/a:srgbClr/@val">
          <xsl:attribute name ="draw:fill">
            <xsl:value-of select="'solid'" />
          </xsl:attribute>
          <xsl:attribute name ="draw:fill-color">
            <xsl:value-of select="concat('#',p:spPr/a:solidFill/a:srgbClr/@val)"/>
          </xsl:attribute>
          <!-- Transparency percentage-->
          <xsl:if test="p:spPr/a:solidFill/a:srgbClr/a:alpha/@val">
            <xsl:variable name ="alpha">
              <xsl:value-of select ="p:spPr/a:solidFill/a:srgbClr/a:alpha/@val"/>
            </xsl:variable>
            <xsl:if test="($alpha != '') or ($alpha != 0)">
              <xsl:attribute name ="draw:opacity">
                <xsl:value-of select="concat(($alpha div 1000), '%')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
        </xsl:if>
        <!--Theme color-->
        <xsl:if test ="p:spPr/a:solidFill/a:schemeClr/@val">
          <xsl:attribute name ="draw:fill">
            <xsl:value-of select="'solid'" />
          </xsl:attribute>
          <xsl:attribute name ="draw:fill-color">
            <xsl:call-template name ="getColorCode">
              <xsl:with-param name ="color">
                <xsl:value-of select="p:spPr/a:solidFill/a:schemeClr/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumMod">
                <xsl:value-of select="p:spPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumOff">
                <xsl:value-of select="p:spPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
          <!-- Transparency percentage-->
          <xsl:if test="p:spPr/a:solidFill/a:schemeClr/a:alpha/@val">
            <xsl:variable name ="alpha">
              <xsl:value-of select ="p:spPr/a:solidFill/a:schemeClr/a:alpha/@val"/>
            </xsl:variable>
            <xsl:if test="($alpha != '') or ($alpha != 0)">
              <xsl:attribute name ="draw:opacity">
                <xsl:value-of select="concat(($alpha div 1000), '%')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
        </xsl:if>
      </xsl:when>
      <xsl:when test="p:spPr/a:gradFill">
        <xsl:for-each select="p:spPr/a:gradFill/a:gsLst/child::node()[1]">
          <xsl:if test="name()='a:gs'">
            <xsl:choose>
              <xsl:when test="a:srgbClr/@val">
                <xsl:attribute name="draw:fill-color">
                  <xsl:value-of select="concat('#',a:srgbClr/@val)" />
                </xsl:attribute>
                <xsl:attribute name="draw:fill">
                  <xsl:value-of select="'solid'"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="a:schemeClr/@val">
                <xsl:attribute name="draw:fill-color">
                  <xsl:call-template name="getColorCode">
                    <xsl:with-param name="color">
                      <xsl:value-of select="a:schemeClr/@val" />
                    </xsl:with-param>
                    <xsl:with-param name="lumMod">
                      <xsl:value-of select="a:schemeClr/a:lumMod/@val" />
                    </xsl:with-param>
                    <xsl:with-param name="lumOff">
                      <xsl:value-of select="a:schemeClr/a:lumOff/@val" />
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:attribute>
                <xsl:attribute name="draw:fill">
                  <xsl:value-of select="'solid'"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:choose>
              <xsl:when test="a:srgbClr/a:alpha/@val">
                <xsl:variable name ="opacity">
                  <xsl:value-of select ="a:srgbClr/a:alpha/@val"/>
                </xsl:variable>
                <xsl:if test="($opacity != '') or ($opacity != 0)">
                  <xsl:attribute name ="draw:opacity">
                    <xsl:value-of select="concat(($opacity div 1000), '%')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:when>
              <xsl:when test="a:schemeClr/a:alpha/@val">
                <xsl:variable name ="opacity">
                  <xsl:value-of select ="a:schemeClr/a:alpha/@val"/>
                </xsl:variable>
                <xsl:if test="($opacity != '') or ($opacity != 0)">
                  <xsl:attribute name ="draw:opacity">
                    <xsl:value-of select="concat(($opacity div 1000), '%')"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <!--Fill refernce-->
      <xsl:when test ="p:style/a:fillRef">
        <!-- Standard color-->
        <xsl:if test ="p:style/a:fillRef/a:srgbClr/@val">
          <xsl:attribute name ="draw:fill">
            <xsl:value-of select="'solid'" />
          </xsl:attribute>
          <xsl:attribute name ="draw:fill-color">
            <xsl:value-of select="concat('#',p:style/a:fillRef/a:srgbClr/@val)"/>
          </xsl:attribute>
        </xsl:if>
        <!--Theme color-->
        <xsl:if test ="p:style/a:fillRef//a:schemeClr/@val">
          <xsl:attribute name ="draw:fill">
            <xsl:value-of select="'solid'" />
          </xsl:attribute>
          <xsl:attribute name ="draw:fill-color">
            <xsl:call-template name ="getColorCode">
              <xsl:with-param name ="color">
                <xsl:value-of select="p:style/a:fillRef/a:schemeClr/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumMod">
                <xsl:value-of select="p:style/a:fillRef/a:schemeClr/a:lumMod/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumOff">
                <xsl:value-of select="p:style/a:fillRef/a:schemeClr/a:lumOff/@val"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
    </xsl:choose>
    <!-- added by vipul for fill Tranperancy-->
    <xsl:choose>
      <xsl:when test="p:spPr/a:solidFill/a:srgbClr/a:alpha/@val">
        <xsl:variable name ="opacity">
          <xsl:value-of select ="p:spPr/a:solidFill/a:srgbClr/a:alpha/@val"/>
        </xsl:variable>
        <xsl:if test="($opacity != '') or ($opacity != 0)">
          <xsl:attribute name ="draw:opacity">
            <xsl:value-of select="concat(($opacity div 1000), '%')"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <xsl:when test="p:spPr/a:solidFill/a:schemeClr/a:alpha/@val">
        <xsl:variable name ="opacity">
          <xsl:value-of select ="p:spPr/a:solidFill/a:schemeClr/a:alpha/@val"/>
        </xsl:variable>
        <xsl:if test="($opacity != '') or ($opacity != 0)">
          <xsl:attribute name ="draw:opacity">
            <xsl:value-of select="concat(($opacity div 1000), '%')"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
    </xsl:choose>
    
    <!--Line Color-->
    <xsl:choose>
      <!-- No line-->
      <xsl:when test ="p:spPr/a:ln/a:noFill">
        <xsl:attribute name ="draw:stroke">
          <xsl:value-of select="'none'" />
        </xsl:attribute>
      </xsl:when>
      <!-- Solid line color-->
      <xsl:when test ="p:spPr/a:ln/a:solidFill">
        <xsl:attribute name ="draw:stroke">
          <xsl:value-of select="'solid'" />
        </xsl:attribute>
        <!-- Standard color for border-->
        <xsl:if test ="p:spPr/a:ln/a:solidFill/a:srgbClr/@val">
          <xsl:attribute name ="svg:stroke-color">
            <xsl:value-of select="concat('#',p:spPr/a:ln/a:solidFill/a:srgbClr/@val)"/>
          </xsl:attribute>
          <!-- Transparency percentage-->
          <xsl:if test="p:spPr/a:ln/a:solidFill/a:srgbClr/a:alpha/@val">
            <xsl:variable name ="alpha">
              <xsl:value-of select ="p:spPr/a:ln/a:solidFill/a:srgbClr/a:alpha/@val"/>
            </xsl:variable>
            <xsl:if test="($alpha != '') or ($alpha != 0)">
              <xsl:attribute name ="svg:stroke-opacity">
                <xsl:value-of select="concat(($alpha div 1000), '%')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
        </xsl:if>
        <!-- Theme color for border-->
        <xsl:if test ="p:spPr/a:ln/a:solidFill/a:schemeClr/@val">
          <xsl:attribute name ="svg:stroke-color">
            <xsl:call-template name ="getColorCode">
              <xsl:with-param name ="color">
                <xsl:value-of select="p:spPr/a:ln/a:solidFill/a:schemeClr/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumMod">
                <xsl:value-of select="p:spPr/a:ln/a:solidFill/a:schemeClr/a:lumMod/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumOff">
                <xsl:value-of select="p:spPr/a:ln/a:solidFill/a:schemeClr/a:lumOff/@val"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
          <!-- Transparency percentage-->
          <xsl:if test="p:spPr/a:ln/a:solidFill/a:schemeClr/a:alpha/@val">
            <xsl:variable name ="alpha">
              <xsl:value-of select ="p:spPr/a:ln/a:solidFill/a:schemeClr/a:alpha/@val"/>
            </xsl:variable>
            <xsl:if test="($alpha != '') or ($alpha != 0)">
              <xsl:attribute name ="svg:stroke-opacity">
                <xsl:value-of select="concat(($alpha div 1000), '%')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
        </xsl:if>
      </xsl:when>
      <xsl:when test ="p:style/a:lnRef">
        <xsl:attribute name ="draw:stroke">
          <xsl:value-of select="'solid'" />
        </xsl:attribute>
        <!--Standard color for border-->
        <xsl:if test ="p:style/a:lnRef/a:srgbClr/@val">
          <xsl:attribute name ="svg:stroke-color">
            <xsl:value-of select="concat('#',p:style/a:lnRef/a:srgbClr/@val)"/>
          </xsl:attribute>
        </xsl:if>
        <!--Theme color for border-->
        <xsl:if test ="p:style/a:lnRef/a:schemeClr/@val">
          <xsl:attribute name ="svg:stroke-color">
            <xsl:call-template name ="getColorCode">
              <xsl:with-param name ="color">
                <xsl:value-of select="p:style/a:lnRef/a:schemeClr/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumMod">
                <xsl:value-of select="p:style/a:lnRef/a:schemeClr/a:lumMod/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumOff">
                <xsl:value-of select="p:style/a:lnRef/a:schemeClr/a:lumOff/@val"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
    </xsl:choose>
    <!--Line Style-->
    <xsl:call-template name="tmpLineStyle">
      <xsl:with-param name="ThemeName" select="$ThemeName"/>
    </xsl:call-template>
	<xsl:message terminate="no">progress:p:cSld</xsl:message>
    <xsl:if test ="p:txBody/a:bodyPr/@tIns">
      <xsl:attribute name ="fo:padding-top">
        <xsl:value-of select ="concat(format-number(number(p:txBody/a:bodyPr/@tIns div 360000), '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="p:txBody/a:bodyPr/@lIns">
      <xsl:attribute name ="fo:padding-left">
        <xsl:value-of select ="concat(format-number(number(p:txBody/a:bodyPr/@lIns div 360000), '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="p:txBody/a:bodyPr/@bIns">
      <xsl:attribute name ="fo:padding-bottom">
        <xsl:value-of select ="concat(format-number(number(p:txBody/a:bodyPr/@bIns div 360000), '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="p:txBody/a:bodyPr/@rIns">
      <xsl:attribute name ="fo:padding-right">
        <xsl:value-of select ="concat(format-number(number(p:txBody/a:bodyPr/@rIns div 360000), '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="p:txBody/a:bodyPr/@anchor">
      <xsl:attribute name ="draw:textarea-vertical-align">
        <xsl:choose >
          <!-- Top-->
          <xsl:when test ="p:txBody/a:bodyPr/@anchor = 't' ">
            <xsl:value-of  select ="'top'"/>
          </xsl:when>
          <!-- Middle-->
          <xsl:when test ="p:txBody/a:bodyPr/@anchor = 'ctr' ">
            <xsl:value-of  select ="'middle'"/>
          </xsl:when>
          <!-- bottom-->
          <xsl:when test ="p:txBody/a:bodyPr/@anchor = 'b' ">
            <xsl:value-of  select ="'bottom'"/>
          </xsl:when>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if >
    <xsl:if test ="p:txBody/a:bodyPr/@anchorCtr" >
      <xsl:attribute name ="draw:textarea-horizontal-align">
        <xsl:if test ="p:txBody/a:bodyPr/@anchorCtr= 1">
          <xsl:value-of select ="'center'"/>
        </xsl:if>
        <xsl:if test ="p:txBody/a:bodyPr/@anchorCtr= 0">
          <xsl:value-of select ="'justify'"/>
        </xsl:if>
      </xsl:attribute >
    </xsl:if>
    <xsl:if test ="p:txBody/a:bodyPr/@wrap">
      <xsl:choose>
        <xsl:when test="p:txBody/a:bodyPr/@wrap='none'">
          <xsl:attribute name ="fo:wrap-option">
            <xsl:value-of select ="'no-wrap'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="p:txBody/a:bodyPr/@wrap='square'">
          <xsl:attribute name ="fo:wrap-option">
            <xsl:value-of select ="'wrap'"/>
          </xsl:attribute>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
    <xsl:if test ="p:txBody/a:bodyPr/a:spAutoFit">
      <xsl:attribute name ="draw:auto-grow-height" >
        <xsl:value-of select ="'true'"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:for-each select="./p:txBody">
      <xsl:if test="a:bodyPr/@vert='vert'">
        <xsl:choose>
          <xsl:when test="(a:bodyPr/@anchor='b' and a:bodyPr/@anchorCtr='0') or (a:bodyPr/@anchor='b' and not(a:bodyPr/@anchorCtr))">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'left'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'top'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="(a:bodyPr/@anchor='t' and a:bodyPr/@anchorCtr='0') or (a:bodyPr/@anchor='t' and not(a:bodyPr/@anchorCtr))">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'right'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'top'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="(a:bodyPr/@anchor='ctr' and a:bodyPr/@anchorCtr='0') or (a:bodyPr/@anchor='ctr' and not(a:bodyPr/@anchorCtr))">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'center'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'top'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="a:bodyPr/@anchor='t' and a:bodyPr/@anchorCtr='1'">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'right'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'middle'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="a:bodyPr/@anchor='ctr' and a:bodyPr/@anchorCtr='1'">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'center'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'middle'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="a:bodyPr/@anchor='b' and a:bodyPr/@anchorCtr='1'">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'left'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'middle'"/>
            </xsl:attribute>
          </xsl:when>
        </xsl:choose>
      </xsl:if>
      <xsl:if test="a:bodyPr/@vert='horz' or not(a:bodyPr/@vert) ">
        <xsl:choose>
          <xsl:when test="(a:bodyPr/@anchor='t' and a:bodyPr/@anchorCtr='0') or (a:bodyPr/@anchor='t' and not(a:bodyPr/@anchorCtr))">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'justify'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'top'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="(a:bodyPr/@anchor='ctr' and a:bodyPr/@anchorCtr='0') or (a:bodyPr/@anchor='ctr' and not(a:bodyPr/@anchorCtr))  ">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'justify'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'middle'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="a:bodyPr/@anchor='t' and a:bodyPr/@anchorCtr='1'">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'center'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'top'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="(a:bodyPr/@anchor='b' and a:bodyPr/@anchorCtr='0') or (a:bodyPr/@anchor='b' and not(a:bodyPr/@anchorCtr))">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'justify'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'bottom'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="a:bodyPr/@anchor='ctr' and a:bodyPr/@anchorCtr='1'">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'center'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'middle'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="a:bodyPr/@anchor='b' and a:bodyPr/@anchorCtr='1'">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'center'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'bottom'"/>
            </xsl:attribute>
          </xsl:when>
        </xsl:choose>
      </xsl:if>
</xsl:for-each>
  </xsl:template>
  <!-- Start - Template to Add hyperlinks defined at the Text level - added by lohith-->
  <xsl:template name="AddTextHyperlinks">
    <xsl:param name="nodeAColonR" />
    <xsl:param name="slideRelationId" />
    <xsl:param name="slideId" />
	<xsl:message terminate="no">progress:p:cSld</xsl:message>
    <!-- varible to get the slide number ( Eg: 1,2 etc )-->
    <xsl:variable name="slidePostion">
      <xsl:value-of select="format-number(substring-after($slideId,'slide'),'#')"/>
    </xsl:variable>
    <!-- varible to get the numer of slides in the persentation -->
    <xsl:variable name="slideCount">
      <xsl:value-of select="count(document('ppt/presentation.xml')/p:presentation/p:sldIdLst/p:sldId)"/>
    </xsl:variable>
    <xsl:if test="$nodeAColonR/a:hlinkClick/@action[contains(.,'slide')] and string-length($nodeAColonR/a:hlinkClick/@r:id) = 0 ">
      <xsl:choose>
        <xsl:when test="$nodeAColonR/a:hlinkClick/@action[ contains(.,'jump=previousslide')]">
          <xsl:if test="$slidePostion > 1">
            <xsl:attribute name="xlink:href">
              <xsl:value-of select="concat('#Slide ',($slidePostion - 1))"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="$slidePostion = 1">
            <xsl:value-of select="concat('#Slide ',1)"/>
          </xsl:if>
        </xsl:when>
        <xsl:when test="$nodeAColonR/a:hlinkClick/@action[ contains(.,'jump=nextslide')]">
          <xsl:if test="$slidePostion = $slideCount">
            <xsl:attribute name="xlink:href">
              <xsl:value-of select="concat('#Slide ',$slidePostion)"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="$slidePostion &lt; $slideCount">
            <xsl:attribute name="xlink:href">
              <xsl:value-of select="concat('#Slide ',($slidePostion + 1))"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:when>
        <xsl:when test="$nodeAColonR/a:hlinkClick/@action[ contains(.,'jump=firstslide')]">
          <xsl:attribute name="xlink:href">
            <xsl:value-of select="concat('#Slide ',1)"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="$nodeAColonR/a:hlinkClick/@action[ contains(.,'jump=lastslide')]">
          <xsl:attribute name="xlink:href">
            <xsl:value-of select="concat('#Slide ',$slideCount)"/>
          </xsl:attribute>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
    <xsl:if test="string-length($nodeAColonR/a:hlinkClick/@r:id) > 0">
      <xsl:variable name="RelationId">
        <xsl:value-of select="$nodeAColonR/a:hlinkClick/@r:id"/>
      </xsl:variable>
      <xsl:variable name="Target">
        <xsl:value-of select="document($slideRelationId)/rels:Relationships/rels:Relationship[@Id=$RelationId]/@Target"/>
      </xsl:variable>
      <xsl:variable name="type">
        <xsl:value-of select="document($slideRelationId)/rels:Relationships/rels:Relationship[@Id=$RelationId]/@Type"/>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="contains($Target,'mailto:') or contains($Target,'http:') or contains($Target,'https:')">
          <xsl:attribute name="xlink:href">
            <xsl:value-of select="$Target"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="contains($Target,'slide')">
          <xsl:attribute name="xlink:href">
            <xsl:value-of select="concat('#Slide ',substring-before(substring-after($Target,'slide'),'.xml'))"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="contains($Target,'file:///')">
          <xsl:attribute name="xlink:href">
            <xsl:value-of select="concat('file:///',translate(substring-after($Target,'file:///'),'\','/'))"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="contains($type,'hyperlink') and not(contains($Target,'http:')) and not(contains($Target,'https:'))">
          <!-- warn if hyperlink Path  -->
          <xsl:message terminate="no">translation.oox2odf.hyperlinkTypeRelativePath</xsl:message>
          <xsl:attribute name="xlink:href">
            <!--Link Absolute Path-->
            <xsl:value-of select ="concat('hyperlink-path:',$Target)"/>
            <!--End-->
          </xsl:attribute>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
  </xsl:template>
  <!-- End - Template to Add hyperlinks defined at the Text level -->
  <!-- Template to generate GUID - Folder to copy the sound files -->
  <xsl:template name="GenerateGUIDForFolderName">
    <xsl:param name="RootNode"/>
    <xsl:variable name="GUIDOne">
      <xsl:value-of select="generate-id($RootNode)"/>
    </xsl:variable>
    <xsl:variable name="GUIDTwo">
      <xsl:value-of select="generate-id(child::node())"/>
    </xsl:variable>
    <xsl:variable name="GUIDThree">
      <xsl:value-of select="generate-id(.)"/>
    </xsl:variable>
    <xsl:value-of select="concat($GUIDOne,'-',$GUIDTwo,'-',$GUIDThree)"/>
  </xsl:template>
  <!-- End of template - GenerateGUIDForFolderName -->
  <xsl:template name="tmpUnderLine">
	<xsl:message terminate="no">progress:p:cSld</xsl:message>
    <xsl:if test ="./a:uFill/a:solidFill/a:srgbClr/@val">
      <xsl:attribute name ="style:text-underline-color">
        <xsl:value-of select ="concat('#',./a:uFill/a:solidFill/a:srgbClr/@val)"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="./a:uFill/a:solidFill/a:schemeClr/@val">
      <xsl:attribute name ="style:text-underline-color">
        <xsl:call-template name ="getColorCode">
          <xsl:with-param name ="color">
            <xsl:value-of select="./a:uFill/a:solidFill/a:schemeClr/@val"/>
          </xsl:with-param>
          <xsl:with-param name ="lumMod">
            <xsl:value-of select="./a:uFill/a:solidFill/a:schemeClr/a:lumMod/@val"/>
          </xsl:with-param>
          <xsl:with-param name ="lumOff">
            <xsl:value-of select="./a:uFill/a:solidFill/a:schemeClr/a:lumOff/@val"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:choose >
      <xsl:when test ="./@u='dbl'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'solid'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-type">
          <xsl:value-of select ="'double'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'auto'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='heavy'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'solid'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'bold'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='dotted'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'dotted'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'auto'"/>
        </xsl:attribute>
      </xsl:when>
      <!-- dottedHeavy-->
      <xsl:when test ="./@u='dottedHeavy'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'dotted'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'bold'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='dash'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'dash'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'auto'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='dashHeavy'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'dash'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'bold'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='dashLong'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'long-dash'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'auto'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='dashLongHeavy'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'long-dash'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'bold'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='dotDash'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'dot-dash'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'auto'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='dotDashHeavy'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'dot-dash'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'bold'"/>
        </xsl:attribute>
      </xsl:when>
      <!-- dot-dot-dash-->
      <xsl:when test ="./@u='dotDotDash'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'dot-dot-dash'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'auto'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='dotDotDashHeavy'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'dot-dot-dash'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'bold'"/>
        </xsl:attribute>
      </xsl:when>
      <!-- Wavy and Heavy-->
      <xsl:when test ="./@u='wavy'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'wave'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'auto'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='wavyHeavy'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'wave'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'bold'"/>
        </xsl:attribute>
      </xsl:when>
      <!-- wavyDbl-->
      <!-- style:text-underline-style="wave" style:text-underline-type="double"-->
      <xsl:when test ="./@u='wavyDbl'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'wave'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-type">
          <xsl:value-of select ="'double'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'auto'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='sng'">
        <xsl:attribute name ="style:text-underline-type">
          <xsl:value-of select ="'single'"/>
        </xsl:attribute>
        <xsl:attribute name ="style:text-underline-width">
          <xsl:value-of select ="'auto'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="./@u='none'">
        <xsl:attribute name ="style:text-underline-style">
          <xsl:value-of select ="'none'"/>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name ="tmpShapeTextProcess">
    <xsl:param name="ParaId"/>
    <xsl:param name="TypeId"/>
    <xsl:param name="SMName"/>
    <xsl:param name="DefFont"/>
    
    <xsl:variable name ="slideRel" select ="concat('ppt/slides/_rels/',$TypeId,'.xml.rels')"/>
	<xsl:message terminate="no">progress:p:cSld</xsl:message>
    <xsl:message terminate="no">progress:a:p</xsl:message>
    <xsl:for-each select ="./p:txBody">
      <xsl:variable name="var_fontScale">
        <xsl:if test="./a:bodyPr/a:normAutofit/@fontScale">
          <xsl:value-of select="./a:bodyPr/a:normAutofit/@fontScale"/>
        </xsl:if>
        <xsl:if test="not(./a:bodyPr/a:normAutofit/@fontScale)">
          <xsl:value-of select="'100000'"/>
        </xsl:if>
      </xsl:variable>
      <xsl:variable name="var_lnSpcReduction">
        <xsl:if test="./a:bodyPr/a:normAutofit/@lnSpcReduction">
          <xsl:value-of select="./a:bodyPr/a:normAutofit/@lnSpcReduction"/>
        </xsl:if>
        <xsl:if test="not(./a:bodyPr/a:normAutofit/@lnSpcReduction)">
          <xsl:value-of select="'0'"/>
        </xsl:if>
      </xsl:variable>
      <xsl:variable name ="listStyleName">
        <!-- Added by vijayeta, to get the text box number-->
        <xsl:variable name ="textNumber" select ="./parent::node()/p:nvSpPr/p:cNvPr/@id"/>
        <!-- Added by vijayeta, to get the text box number-->
        <xsl:choose>
          <xsl:when test ="./parent::node()/p:spPr/a:prstGeom/@prst or ./parent::node()/child::node()[1]/child::node()[1]/@name[contains(., 'TextBox')] or ./parent::node()/child::node()[1]/child::node()[1]/@name[contains(., 'Text Box')]">
            <xsl:value-of select ="concat($TypeId,'textboxshape_List',$textNumber)"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select ="concat($TypeId,'List',position())"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:if test ="(a:p/a:pPr/a:buChar) or (a:p/a:pPr/a:buAutoNum) or (a:p/a:pPr/a:buBlip) ">
        <xsl:call-template name ="insertBulletStyle">          
          <xsl:with-param name ="ParaId" select ="$ParaId"/>
          <xsl:with-param name ="listStyleName" select ="$listStyleName"/>
          <xsl:with-param name ="slideRel" select ="$slideRel"/>
        </xsl:call-template>
      </xsl:if>
      <xsl:for-each select ="a:p">
        <xsl:variable name ="levelForDefFont">
          <!--<xsl:if test ="$bulletTypeBool='true'">-->
          <xsl:if test ="a:pPr/@lvl">
            <xsl:value-of select ="a:pPr/@lvl"/>
          </xsl:if>
          <xsl:if test ="not(a:pPr/@lvl)">
            <xsl:value-of select ="'0'"/>
          </xsl:if>
          <!--</xsl:if>-->
        </xsl:variable>
        <style:style style:family="paragraph">
          <xsl:attribute name ="style:name">
            <xsl:value-of select ="concat($ParaId,position())"/>
          </xsl:attribute >
          <style:paragraph-properties  text:enable-numbering="false" >
            <xsl:if test ="parent::node()/a:bodyPr/@vert='vert'">
              <xsl:attribute name ="style:writing-mode">
                <xsl:value-of select ="'tb-rl'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:call-template name="tmpSlideParagraphStyle">
              <xsl:with-param name="lnSpcReduction" select="$var_lnSpcReduction"/>
            </xsl:call-template>
            <!-- added by vipul to get Paragraph propreties from Presentation.xml-->
              <xsl:call-template name="tmpPresentationDefaultParagraphStyle">
              <xsl:with-param name="lnSpcReduction" select="$var_lnSpcReduction"/>
              <xsl:with-param name="level" select="$levelForDefFont+1"/>
            </xsl:call-template>
            <xsl:if test ="a:pPr/a:buChar or a:pPr/a:buAutoNum or a:pPr/a:buBlip ">
              <xsl:choose >
                <xsl:when test ="not(a:r/a:t)">
                  <xsl:attribute name="text:enable-numbering">
                    <xsl:value-of select ="'false'"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise >
                  <xsl:attribute name="text:enable-numbering">
                    <xsl:value-of select ="'true'"/>
                  </xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:if>
            <!-- @@Code for Paragraph tabs Pradeep Nemadi -->
            <!-- Starts-->
            <xsl:if test ="a:pPr/a:tabLst/a:tab">
              <xsl:call-template name ="paragraphTabstops" />
            </xsl:if>
            <!-- Ends -->
          </style:paragraph-properties >
        </style:style>
        <xsl:for-each select ="node()" >
          <!-- Add here-->
          <xsl:if test ="name()='a:r'">
            <style:style style:family="text">
              <xsl:attribute name="style:name">
                <xsl:value-of select="concat($TypeId,generate-id())"/>
              </xsl:attribute>
              <style:text-properties>
                <xsl:call-template name="tmpSlideTextProperty">
                  <xsl:with-param name="DefFont" select="$DefFont"/>
                  <xsl:with-param name="fontscale" select="$var_fontScale"/>
                </xsl:call-template>
                <!-- added by vipul to get text propreties from Presentation.xml-->
                <xsl:call-template name="tmpPresentationDefaultTextProp">
                  <xsl:with-param name="DefFont" select="$DefFont"/>
                  <xsl:with-param name="level" select="$levelForDefFont+1"/>
                  <xsl:with-param name="fontscale" select="$var_fontScale"/>
                  <xsl:with-param name="SMName" select="$SMName"/>
                </xsl:call-template>
                
              </style:text-properties>
            </style:style>
          </xsl:if>
          <!-- Added by lohith.ar for fix 1731885-->
          <xsl:if test ="name()='a:endParaRPr'">
            <style:style style:family="text">
              <xsl:attribute name="style:name">
                <xsl:value-of select="concat($TypeId,generate-id())"/>
              </xsl:attribute>
              <style:text-properties>
                <!-- Bug 1711910 Fixed,On date 2-06-07,by Vijayeta-->
                <xsl:choose>
                  <xsl:when test ="a:latin/@typeface">
                    <xsl:attribute name ="fo:font-family">
                    <xsl:variable name ="typeFaceVal" select ="a:latin/@typeface"/>
                    <xsl:for-each select ="a:latin/@typeface">
                      <xsl:if test ="$typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt'">
                        <xsl:value-of select ="$DefFont"/>
                      </xsl:if>
                      <!-- Bug 1711910 Fixed,On date 2-06-07,by Vijayeta-->
                      <xsl:if test ="not($typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt')">
                        <xsl:value-of select ="."/>
                      </xsl:if>
                    </xsl:for-each>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:when test ="a:cs/@typeface">
                    <xsl:attribute name ="fo:font-family">
                      <xsl:variable name ="typeFaceVal" select ="a:cs/@typeface"/>
                      <xsl:for-each select ="a:cs/@typeface">
                        <xsl:if test ="$typeFaceVal='+mn-cs' or $typeFaceVal='+mj-cs'">
                          <xsl:value-of select ="$DefFont"/>
                        </xsl:if>
                        <xsl:if test ="not($typeFaceVal='+mn-cs' or $typeFaceVal='+mj-cs')">
                          <xsl:value-of select ="."/>
                        </xsl:if>
                      </xsl:for-each>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:when test ="a:sym/@typeface">
                    <xsl:attribute name ="fo:font-family">
                      <xsl:variable name ="typeFaceVal" select ="a:sym/@typeface"/>
                      <xsl:for-each select ="a:sym/@typeface">
                        <xsl:if test ="$typeFaceVal='+mn-sym' or $typeFaceVal='+mj-sym'">
                          <xsl:value-of select ="$DefFont"/>
                        </xsl:if>
                        <xsl:if test ="not($typeFaceVal='+mn-sym' or $typeFaceVal='+mj-sym')">
                          <xsl:value-of select ="."/>
                        </xsl:if>
                      </xsl:for-each>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:when test ="not(a:latin/@typeface) and not(a:cs/@typeface) and not(a:sym/@typeface) ">
                    <xsl:attribute name ="fo:font-family">
                      <xsl:value-of select ="$DefFont"/>
                    </xsl:attribute>
                  </xsl:when>
                </xsl:choose>

                <xsl:attribute name ="style:font-family-generic"	>
                  <xsl:value-of select ="'roman'"/>
                </xsl:attribute>
                <xsl:attribute name ="style:font-pitch"	>
                  <xsl:value-of select ="'variable'"/>
                </xsl:attribute>
                <xsl:attribute name ="style:font-family-generic"	>
                  <xsl:value-of select ="'roman'"/>
                </xsl:attribute>
                <xsl:attribute name ="style:font-pitch"	>
                  <xsl:value-of select ="'variable'"/>
                </xsl:attribute>
                <xsl:if test ="./@sz">
                  <xsl:attribute name ="fo:font-size"	>
                    <xsl:for-each select ="./@sz">
                      <xsl:value-of select ="concat(format-number(. div 100,'#.##'), 'pt')"/>
                    </xsl:for-each>
                  </xsl:attribute>
                </xsl:if>
                <!--Empty lines font size in BVT and bug 1764385, date 13th aug-->
                <!--it was '<xsl:if test ="not(a:pPr/@sz)">' earlier-->
                <xsl:if test ="not(./@sz)">
                  <xsl:if test="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr/a:defRPr/@sz">
                    <xsl:for-each select="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr/a:defRPr">
                      <xsl:attribute name ="fo:font-size">
                        <xsl:value-of  select ="concat(format-number(@sz div 100,'#.##'),'pt')"/>
                      </xsl:attribute>
                    </xsl:for-each>
                  </xsl:if>
                </xsl:if>
              </style:text-properties>
            </style:style>
          </xsl:if>
        </xsl:for-each >
      </xsl:for-each>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="paragraphTabstops">
    <style:tab-stops>
      <xsl:for-each select ="a:pPr/a:tabLst/a:tab">
        <style:tab-stop>
          <xsl:attribute name ="style:position">
            <xsl:value-of select ="concat(format-number(@pos div 360000,'#.##'),'cm') "/>
          </xsl:attribute>
          <xsl:if test ="@algn">
            <xsl:attribute name ="style:type">
              <xsl:choose >
                <xsl:when test ="@algn='ctr'">
                  <xsl:value-of select ="'center'"/>
                </xsl:when>
                <xsl:when test ="@algn='r'">
                  <xsl:value-of select ="'right'"/>
                </xsl:when>
                <xsl:when test ="@algn='l'">
                  <xsl:value-of select ="'left'"/>
                </xsl:when>
                <xsl:when test ="@algn='dec'">
                  <xsl:value-of select ="'char'"/>
                </xsl:when>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
        </style:tab-stop >
      </xsl:for-each>
    </style:tab-stops>
  </xsl:template>
  <!-- Hand out Common Templates-->
  <xsl:template name="tmpHMParagraphStyle">
    <!--Text alignment-->
    <xsl:if test ="./p:txBody/a:p/a:pPr/@algn">
      <xsl:attribute name ="fo:text-align">
        <xsl:choose>
          <!-- Center Alignment-->
          <xsl:when test ="./p:txBody/a:p/a:pPr/@algn ='ctr'">
            <xsl:value-of select ="'center'"/>
          </xsl:when>
          <!-- Right Alignment-->
          <xsl:when test ="./p:txBody/a:p/a:pPr/@algn ='r'">
            <xsl:value-of select ="'end'"/>
          </xsl:when>
          <!-- Left Alignment-->
          <xsl:when test ="./p:txBody/a:p/a:pPr/@algn ='l'">
            <xsl:value-of select ="'start'"/>
          </xsl:when>
          <!-- Added by lohith - for fix 1737161-->
          <xsl:when test ="./p:txBody/a:p/a:pPr/@algn ='just'">
            <xsl:value-of select ="'justify'"/>
          </xsl:when>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if >
    <!-- Convert Laeft margin of the paragraph-->
    <xsl:if test ="./p:txBody/a:p/a:pPr/@marL">
      <xsl:attribute name ="fo:margin-left">
        <xsl:value-of select="concat(format-number( ./p:txBody/a:p/a:pPr/@marL div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="./p:txBody/a:p/a:pPr/@marR">
      <xsl:attribute name ="fo:margin-right">
        <xsl:value-of select="concat(format-number(./p:txBody/a:p/a:pPr/@marR div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <!--<xsl:if test ="./@indent">
      <xsl:attribute name ="fo:text-indent">
        <xsl:value-of select="concat(format-number(./@indent div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if >-->
    <xsl:if test ="./p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz">
      <xsl:variable name="var_fontsize">
        <xsl:value-of select="./p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz"/>
      </xsl:variable>
      <xsl:for-each select="./p:txBody/a:p/a:pPr/a:spcBef/a:spcPct">
        <xsl:if test ="@val">
          <xsl:attribute name ="fo:margin-top">
            <xsl:value-of select="concat(format-number( (( $var_fontsize * ( @val div 100000 ) ) div 2835) * 1.2  ,'#.###'),'cm')"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:for-each>
      <xsl:for-each select="./p:txBody/a:p/a:pPr/a:spcAft/a:spcPct">
        <xsl:if test ="@val">
          <xsl:attribute name ="fo:margin-bottom">
            <xsl:value-of select="concat(format-number( (( $var_fontsize * ( @val div 100000 ) ) div 2835 ) * 1.2 ,'#.###'),'cm')"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:for-each>
    </xsl:if>
    <xsl:for-each select="./p:txBody/a:p/a:pPr/a:spcAft/a:spcPts">
      <xsl:if test ="@val">
        <xsl:attribute name ="fo:margin-bottom">
          <xsl:value-of select="concat(format-number(@val div  2835 ,'#.##'),'cm')"/>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
    <xsl:for-each select="./p:txBody/a:p/a:pPr/a:spcBef/a:spcPts">
      <xsl:if test ="@val">
        <xsl:attribute name ="fo:margin-top">
          <xsl:value-of select="concat(format-number(@val div  2835 ,'#.##'),'cm')"/>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
    <xsl:for-each select="./p:txBody/a:p/a:pPr/a:lnSpc">
      <xsl:if test ="./a:spcPct">
        <!--<xsl:choose>
          <xsl:when test="$lnSpcReduction='0'">-->
        <xsl:attribute name="fo:line-height">
          <xsl:value-of select="concat(format-number(./a:spcPct/@val div 1000,'###'), '%')"/>
        </xsl:attribute>
        <!--</xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="fo:line-height">
              <xsl:value-of select="concat(format-number((./a:spcPct/@val - $lnSpcReduction) div 1000,'###'), '%')"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>-->
      </xsl:if >
      <xsl:if test ="./a:spcPts">
        <xsl:attribute name="style:line-height-at-least">
          <xsl:value-of select="concat(format-number(./a:spcPts/@val div 2835, '#.##'), 'cm')"/>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="tmpHandOutTextProperty">
    <xsl:param name ="DefFont" />
    <xsl:param name ="fontscale"/>
    <xsl:if test ="a:rPr/@sz">
      <xsl:attribute name ="fo:font-size"	>
        <xsl:for-each select ="a:rPr/@sz">
          <xsl:choose>
            <xsl:when test="$fontscale ='100000'">
              <xsl:value-of  select ="concat(format-number(. div 100,'#.##'),'pt')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of  select ="concat(format-number(round((. *($fontscale div 1000) )div 10000),'#.##'),'pt')"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="a:rPr/a:latin/@typeface">
      <xsl:attribute name ="fo:font-family">
        <xsl:variable name ="typeFaceVal" select ="a:rPr/a:latin/@typeface"/>
        <xsl:for-each select ="a:rPr/a:latin/@typeface">
          <xsl:if test ="$typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt'">
            <xsl:value-of  select ="$DefFont"/>
          </xsl:if>
          <xsl:if test ="not($typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt')">
            <xsl:value-of select ="."/>
          </xsl:if>
        </xsl:for-each>
      </xsl:attribute>
    </xsl:if>
    <!-- strike style:text-line-through-style-->
    <xsl:if test ="a:rPr/@strike">
      <xsl:attribute name ="style:text-line-through-style">
        <xsl:value-of select ="'solid'"/>
      </xsl:attribute>
      <xsl:choose >
        <xsl:when test ="a:rPr/@strike='dblStrike'">
          <xsl:attribute name ="style:text-line-through-type">
            <xsl:value-of select ="'double'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test ="a:rPr/@strike='sngStrike'">
          <xsl:attribute name ="style:text-line-through-type">
            <xsl:value-of select ="'single'"/>
          </xsl:attribute>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
    <!-- Kening Property-->
    <xsl:if test ="a:rPr/@kern">
      <xsl:choose >
        <xsl:when test ="a:rPr/@kern = '0'">
          <xsl:attribute name ="style:letter-kerning">
            <xsl:value-of select ="'false'"/>
          </xsl:attribute >
        </xsl:when >
        <xsl:when test ="format-number(a:rPr/@kern,'#.##') &gt; 0">
          <xsl:attribute name ="style:letter-kerning">
            <xsl:value-of select ="'true'"/>
          </xsl:attribute >
        </xsl:when >
      </xsl:choose>
    </xsl:if >
    <!-- Bold Property-->
    <xsl:if test ="a:rPr/@b">
      <xsl:if test ="a:rPr/@b='1'">
        <xsl:attribute name ="fo:font-weight">
          <xsl:value-of select ="'bold'"/>
        </xsl:attribute>
      </xsl:if >
      <xsl:if test ="a:rPr/@b='0'">
        <xsl:attribute name ="fo:font-weight">
          <xsl:value-of select ="'normal'"/>
        </xsl:attribute>
      </xsl:if >
    </xsl:if >
    <!--UnderLine-->
    <xsl:if test ="a:rPr/@u">
      <xsl:for-each select ="a:rPr">
        <xsl:call-template name="tmpUnderLine"/>
      </xsl:for-each>
    </xsl:if >
    <!-- Italic-->
    <xsl:if test ="a:rPr/@i">
      <xsl:attribute name ="fo:font-style">
        <xsl:if test ="a:rPr/@i='1'">
          <xsl:value-of select ="'italic'"/>
        </xsl:if >
        <xsl:if test ="a:rPr/@i='0'">
          <xsl:value-of select ="'none'"/>
        </xsl:if>
      </xsl:attribute>
    </xsl:if >
    <!-- Character Spacing -->
    <xsl:if test ="a:rPr/@spc">
      <xsl:attribute name ="fo:letter-spacing">
        <xsl:variable name="length" select="a:rPr/@spc" />
        <xsl:value-of select="concat(format-number($length * 2.54 div 7200,'#.###'),'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
    <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
    <xsl:if test ="a:rPr/a:solidFill/a:srgbClr/@val">
      <xsl:attribute name ="fo:color">
        <xsl:value-of select ="translate(concat('#',a:rPr/a:solidFill/a:srgbClr/@val),$ucletters,$lcletters)"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="a:rPr/a:solidFill/a:schemeClr/@val">
      <xsl:attribute name ="fo:color">
        <xsl:call-template name ="getColorCode">
          <xsl:with-param name ="color">
            <xsl:value-of select="a:rPr/a:solidFill/a:schemeClr/@val"/>
          </xsl:with-param>
          <xsl:with-param name ="lumMod">
            <xsl:value-of select="a:rPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
          </xsl:with-param>
          <xsl:with-param name ="lumOff">
            <xsl:value-of select="a:rPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
   
    <!--Shadow fo:text-shadow-->
    <xsl:if test ="a:rPr/a:effectLst/a:outerShdw">
      <xsl:attribute name ="fo:text-shadow">
        <xsl:value-of select ="'1pt 1pt'"/>
      </xsl:attribute>
    </xsl:if>
    <!--SuperScript and SubScript for Text added by Mathi on 31st Jul 2007-->
    <xsl:if test="(a:rPr/@baseline)">
      <xsl:variable name="baseData">
        <xsl:value-of select="a:rPr/@baseline"/>
      </xsl:variable>
      <xsl:choose>
		  <xsl:when test="(a:rPr/@baseline &gt; 0)">
          <xsl:variable name="superCont">
            <xsl:value-of select="concat('super ',format-number($baseData div 1000,'#.###'),'%')"/>
          </xsl:variable>
          <xsl:attribute name="style:text-position">
            <xsl:value-of select="$superCont"/>
          </xsl:attribute>
        </xsl:when>
		  <xsl:when test="(a:rPr/@baseline &lt; 0)">
          <xsl:variable name="subCont">
            <xsl:value-of select="concat('sub ',format-number(substring-after($baseData,'-') div 1000,'#.###'),'%')"/>
          </xsl:variable>
          <xsl:attribute name="style:text-position">
            <xsl:value-of select="$subCont"/>
          </xsl:attribute>
		  </xsl:when>
      </xsl:choose>
    </xsl:if>
  </xsl:template>
  <xsl:template name ="tmpPresentationDefaultTextProp">
    <xsl:param name="level"/>
    <xsl:param name="DefFont"/>
    <xsl:param name="SMName"/>
    <xsl:message terminate="no">progress:p:cSld</xsl:message>
    <xsl:message terminate="no">progress:a:p</xsl:message>
    <xsl:variable name ="nodeName">
      <xsl:value-of select ="concat('a:lvl',$level,'pPr')"/>
    </xsl:variable>
    <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
    <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
    <xsl:if test ="not(a:rPr/@sz)">
      <xsl:if test="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/@sz">
        <xsl:for-each select="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr">
          <xsl:attribute name ="fo:font-size">
            <xsl:value-of  select ="concat(format-number(@sz div 100,'#.##'),'pt')"/>
          </xsl:attribute>
          <xsl:attribute name ="style:font-size-asian">
            <xsl:value-of  select ="concat(format-number(@sz div 100,'#.##'),'pt')"/>
          </xsl:attribute>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="not(a:rPr/a:latin/@typeface) and not(a:rPr/a:cs/@typeface) and not(a:rPr/a:sym/@typeface)">
      <xsl:if test="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/a:latin/@typeface">
        <xsl:attribute name ="fo:font-family">
          <xsl:variable name ="typeFaceVal" select ="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/a:latin/@typeface"/>
          <xsl:for-each select ="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/a:latin/@typeface">
            <xsl:if test ="$typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt'">
              <xsl:value-of  select ="$DefFont"/>
            </xsl:if>
            <xsl:if test ="not($typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt')">
              <xsl:value-of select ="."/>
            </xsl:if>
          </xsl:for-each>
        </xsl:attribute>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="not(a:rPr/@strike)">
      <xsl:if test="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/@strike">
        <xsl:for-each select="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr">
          <xsl:attribute name ="style:text-line-through-style">
            <xsl:value-of select ="'solid'"/>
          </xsl:attribute>
          <xsl:choose >
            <xsl:when test ="@strike='dblStrike'">
              <xsl:attribute name ="style:text-line-through-type">
                <xsl:value-of select ="'double'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test ="@strike='sngStrike'">
              <xsl:attribute name ="style:text-line-through-type">
                <xsl:value-of select ="'single'"/>
              </xsl:attribute>
            </xsl:when>
          </xsl:choose>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="not(a:rPr/@kern)">
      <xsl:if test="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/@kern">
        <xsl:for-each select="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/a:defRPr">
          <xsl:choose >
            <xsl:when test ="@kern = '0'">
              <xsl:attribute name ="style:letter-kerning">
                <xsl:value-of select ="'false'"/>
              </xsl:attribute >
            </xsl:when >
            <xsl:when test ="format-number(@kern,'#.##') &gt; 0">
              <xsl:attribute name ="style:letter-kerning">
                <xsl:value-of select ="'true'"/>
              </xsl:attribute >
            </xsl:when >
          </xsl:choose>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="not(a:rPr/@b)">
      <xsl:if test="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/@b">
        <xsl:for-each select="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr">
          <xsl:if test ="@b='1'">
            <xsl:attribute name ="fo:font-weight">
              <xsl:value-of select ="'bold'"/>
            </xsl:attribute>
          </xsl:if >
          <xsl:if test ="@b='0'">
            <xsl:attribute name ="fo:font-weight">
              <xsl:value-of select ="'normal'"/>
            </xsl:attribute>
          </xsl:if >
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="not(a:rPr/@u)">
      <xsl:if test="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/@u">
        <xsl:for-each select="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr">
          <xsl:call-template name="tmpUnderLine"/>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="not(a:rPr/@i)">
      <xsl:if test="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/@i">
        <xsl:for-each select="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr">
          <xsl:attribute name ="fo:font-style">
            <xsl:if test ="@i='1'">
              <xsl:value-of select ="'italic'"/>
            </xsl:if >
            <xsl:if test ="@i='0'">
              <xsl:value-of select ="'none'"/>
            </xsl:if>
          </xsl:attribute>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="not(a:rPr/@spc)">
      <xsl:if test="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/@spc">
        <xsl:for-each select="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr">
          <xsl:attribute name ="fo:letter-spacing">
            <xsl:variable name="length" select="@spc" />
            <xsl:value-of select="concat(format-number($length * 2.54 div 7200,'#.###'),'cm')"/>
          </xsl:attribute>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
	  <xsl:if test ="not(a:rPr/a:solidFill/a:srgbClr/@val or a:rPr/a:solidFill/a:schemeClr/@val)">
		  <xsl:choose >
			  <xsl:when test ="not(a:rPr/a:solidFill/a:srgbClr/@val)">
				  <xsl:if test="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
					  <xsl:for-each select="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr">
						  <xsl:attribute name ="fo:color">
							  <xsl:value-of select ="translate(concat('#',a:solidFill/a:srgbClr/@val),$ucletters,$lcletters)"/>
						  </xsl:attribute>
					  </xsl:for-each>
				  </xsl:if>
			  </xsl:when>
			  <xsl:when test ="not(a:rPr/a:solidFill/a:schemeClr/@val)">
				  <xsl:if test="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/@val">
					  <xsl:for-each select="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr">
						  <xsl:attribute name ="fo:color">
							  <xsl:call-template name ="getColorCode">
								  <xsl:with-param name ="color">
									  <xsl:value-of select="a:solidFill/a:schemeClr/@val"/>
								  </xsl:with-param>
								  <xsl:with-param name ="lumMod">
									  <xsl:value-of select="a:solidFill/a:schemeClr/a:lumMod/@val"/>
								  </xsl:with-param>
								  <xsl:with-param name ="lumOff">
									  <xsl:value-of select="a:solidFill/a:schemeClr/a:lumOff/@val"/>
								  </xsl:with-param>
							  </xsl:call-template>
						  </xsl:attribute>
					  </xsl:for-each>
				  </xsl:if>
			  </xsl:when>
			  <xsl:otherwise >
				  <xsl:if test ="not(a:rPr/a:solidFill/a:srgbClr/@val) and not(a:rPr/a:solidFill/a:schemeClr/@val)">
					  <xsl:choose>
						  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
							  <xsl:attribute name ="fo:color">
								  <xsl:value-of select ="translate(concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val),$ucletters,$lcletters)"/>
							  </xsl:attribute>
						  </xsl:when>
						  <xsl:when test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
							  <xsl:attribute name ="fo:color">
								  <xsl:call-template name ="getColorCode">
									  <xsl:with-param name ="color">
										  <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr/@val"/>
									  </xsl:with-param>
									  <xsl:with-param name ="lumMod">
										  <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr/a:lumMod/@val"/>
									  </xsl:with-param>
									  <xsl:with-param name ="lumOff">
										  <xsl:value-of select="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr/a:lumOff/@val"/>
									  </xsl:with-param>
								  </xsl:call-template>
							  </xsl:attribute>
						  </xsl:when>
					  </xsl:choose>
				  </xsl:if>
			  </xsl:otherwise>
		  </xsl:choose>
	  </xsl:if>
    <xsl:if test ="not(a:effectLst/a:outerShdw)">
      <xsl:if test="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/a:effectLst/a:outerShdw">
        <xsl:for-each select="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr">
          <xsl:attribute name ="fo:text-shadow">
            <xsl:value-of select ="'1pt 1pt'"/>
          </xsl:attribute>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
	<xsl:if test="not(a:rPr/@baseline)">
		  <xsl:if test="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/@baseline">
			  <xsl:for-each select="document(concat('ppt/slideMasters/',$SMName))//p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr">
				  <xsl:variable name="baseData">
					  <xsl:value-of select="@baseline"/>
				  </xsl:variable>
				  <xsl:choose>
					  <xsl:when test="(@baseline &gt; 0)">
						  <xsl:variable name="superCont">
							  <xsl:value-of select="concat('super ',format-number($baseData div 1000,'#.###'),'%')"/>
						  </xsl:variable>
						  <xsl:attribute name="style:text-position">
							  <xsl:value-of select="$superCont"/>
						  </xsl:attribute>
					  </xsl:when>
					  <xsl:when test="(@baseline &lt; 0)">
						  <xsl:variable name="subCont">
							  <xsl:value-of select="concat('sub ',format-number(substring-after($baseData,'-') div 1000,'#.###'),'%')"/>
						  </xsl:variable>
						  <xsl:attribute name="style:text-position">
							  <xsl:value-of select="$subCont"/>
						  </xsl:attribute>
					  </xsl:when>
				  </xsl:choose>
			  </xsl:for-each>
		  </xsl:if>
	  </xsl:if>
  </xsl:template>
  <xsl:template name="tmpPresentationDefaultParagraphStyle">
    <xsl:param name="lnSpcReduction"/>
    <xsl:param name="level"/>
    <xsl:variable name ="nodeName">
      <xsl:value-of select ="concat('a:lvl',$level,'pPr')"/>
    </xsl:variable>
  
    <xsl:if test ="a:pPr/a:spcBef/a:spcPct/@val">
      <xsl:variable name="var_MaxFntSize">
        <xsl:choose>
          <xsl:when test ="not(./a:r/a:rPr/@sz)">
            <xsl:if test="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr/a:defRPr/@sz">
              <xsl:for-each select="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr/a:defRPr">
                  <xsl:value-of  select ="@sz"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:variable>
      <xsl:if test="$var_MaxFntSize!=''">
      <xsl:attribute name ="fo:margin-top">
        <xsl:value-of select="concat(format-number( (($var_MaxFntSize * ( a:pPr/a:spcBef/a:spcPct/@val div 100000 ) ) div 2835 ) * 1.2, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    </xsl:if>
    <xsl:if test ="a:pPr/a:spcAft/a:spcPct/@val">
      <xsl:variable name="var_MaxFntSize">
        <xsl:choose>
          <xsl:when test ="not(./a:r/a:rPr/@sz)">
            <xsl:if test="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr/a:defRPr/@sz">
              <xsl:for-each select="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr/a:defRPr">
                <xsl:value-of  select ="@sz"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:variable>
      <xsl:if test="$var_MaxFntSize!=''">
      <xsl:attribute name ="fo:margin-bottom">
        <xsl:value-of select="concat(format-number( (($var_MaxFntSize * ( a:pPr/a:spcAft/a:spcPct/@val div 100000 ) ) div 2835 ) * 1.2, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    </xsl:if>
    <xsl:if test ="not(./a:pPr/@algn)">
      <xsl:if test="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr/@algn">
        <xsl:for-each select="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr">
          <xsl:attribute name ="fo:text-align">
            <xsl:choose>
              <!-- Center Alignment-->
              <xsl:when test ="@algn ='ctr'">
                <xsl:value-of select ="'center'"/>
              </xsl:when>
              <!-- Right Alignment-->
              <xsl:when test ="@algn ='r'">
                <xsl:value-of select ="'end'"/>
              </xsl:when>
              <!-- Left Alignment-->
              <xsl:when test ="@algn ='l'">
                <xsl:value-of select ="'start'"/>
              </xsl:when>
              <!-- Added by lohith - for fix 1737161-->
              <xsl:when test ="@algn ='just'">
                <xsl:value-of select ="'justify'"/>
              </xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="not(./a:pPr/@marL)">
      <xsl:if test="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr/@marL">
        <xsl:for-each select="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr">
          <xsl:attribute name ="fo:margin-left">
            <xsl:value-of select="concat(format-number( @marL div 360000, '#.##'), 'cm')"/>
          </xsl:attribute>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="not(./a:pPr/@marR)">
      <xsl:if test="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr/@marR">
        <xsl:for-each select="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr">
          <xsl:attribute name ="fo:margin-right">
            <xsl:value-of select="concat(format-number(@marR div 360000, '#.##'), 'cm')"/>
          </xsl:attribute>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="not(./a:pPr/@indent)">
      <xsl:if test="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr/@indent">
        <xsl:for-each select="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr">
          <xsl:attribute name ="fo:text-indent">
            <xsl:value-of select="concat(format-number(@indent div 360000, '#.##'), 'cm')"/>
          </xsl:attribute>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="not(./a:spcAft/a:spcPts)">
      <xsl:if test="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr/a:spcAft/a:spcPts">
        <xsl:for-each select="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr/a:spcAft/a:spcPts">
          <xsl:if test ="@val">
            <xsl:attribute name ="fo:margin-bottom">
              <xsl:value-of select="concat(format-number(@val div  2835 ,'#.##'),'cm')"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="not(./a:spcBef/a:spcPts)">
      <xsl:if test="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr/a:spcBef/a:spcPts">
        <xsl:for-each select="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr/a:spcBef/a:spcPts">
          <xsl:if test ="@val">
            <xsl:attribute name ="fo:margin-top">
              <xsl:value-of select="concat(format-number(@val div  2835 ,'#.##'),'cm')"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="not(./a:lnSpc/a:spcPct)">
      <xsl:if test="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr/a:lnSpc/a:spcPct">
        <xsl:for-each select="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr/a:lnSpc/a:spcPct">
          <xsl:choose>
            <xsl:when test="$lnSpcReduction='0'">
              <xsl:attribute name="fo:line-height">
                <xsl:value-of select="concat(format-number(@val div 1000,'###'), '%')"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="fo:line-height">
                <xsl:value-of select="concat(format-number((@val - $lnSpcReduction) div 1000,'###'), '%')"/>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="not(./a:lnSpc/a:spcPts)">
      <xsl:if test="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr/a:lnSpc/a:spcPts">
        <xsl:for-each select="document('ppt/presentation.xml')/p:presentation/p:defaultTextStyle/a:lvl1pPr/a:lnSpc/a:spcPts">
          <xsl:attribute name="style:line-height-at-least">
            <xsl:value-of select="concat(format-number(@val div 2835, '#.##'), 'cm')"/>
          </xsl:attribute>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
  </xsl:template>
  <!--End-->
</xsl:stylesheet>

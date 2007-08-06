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
	<xsl:template name="getColorCode">
		<xsl:param name="color"/>
		<xsl:param name ="lumMod"/>
		<xsl:param name ="lumOff"/>
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
        <xsl:value-of select ="concat(format-number(a:ext/@cx div 360000 ,'#.###'),'cm')"/>
      </xsl:attribute>
      <xsl:attribute name ="svg:height">
        <xsl:value-of select ="concat(format-number(a:ext/@cy div 360000 ,'#.###'),'cm')"/>
      </xsl:attribute>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="tmpSlideParagraphStyle">
    <xsl:param name="lnSpcReduction"/>
    <xsl:param name="blnSlide"/>
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
      <xsl:attribute name ="fo:margin-top">
        <xsl:value-of select="concat(format-number( (($var_MaxFntSize * ( a:pPr/a:spcBef/a:spcPct/@val div 100000 ) ) div 2835 ) * 1.2, '#.##'), 'cm')"/>
      </xsl:attribute>
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
  </xsl:template>
  <xsl:template name="tmpSlideGrahicProp">
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
    <!--<xsl:call-template name="LineStyle"/>-->
    <xsl:if test ="p:txBody/a:bodyPr/@tIns">
      <xsl:attribute name ="fo:padding-top">
        <xsl:value-of select ="concat(format-number(p:txBody/a:bodyPr/@tIns div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="p:txBody/a:bodyPr/@lIns">
      <xsl:attribute name ="fo:padding-left">
        <xsl:value-of select ="concat(format-number(p:txBody/a:bodyPr/@lIns div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="p:txBody/a:bodyPr/@bIns">
      <xsl:attribute name ="fo:padding-bottom">
        <xsl:value-of select ="concat(format-number(p:txBody/a:bodyPr/@bIns div 360000, '#.##'), 'cm')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test ="p:txBody/a:bodyPr/@rIns">
      <xsl:attribute name ="fo:padding-right">
        <xsl:value-of select ="concat(format-number(p:txBody/a:bodyPr/@rIns div 360000, '#.##'), 'cm')"/>
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
    <!--<xsl:attribute name ="fo:min-height" >
      <xsl:value-of select ="'12.573cm'"/>
    </xsl:attribute>-->
   
  </xsl:template>
  <xsl:template name="tmpUnderLine">
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
    <!--@ Default Font Name from PPTX -->
    <xsl:variable name ="DefFont">
      <xsl:for-each select ="document('ppt/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme
						/a:majorFont/a:latin/@typeface">
        <xsl:value-of select ="."/>
      </xsl:for-each>
    </xsl:variable>
    <!--@Modified font Names -->
    <xsl:for-each select ="./p:txBody">
      <!--  by vijayeta,to get linespacing from layouts-->
      <xsl:variable name ="layoutName">
        <xsl:value-of select ="./parent::node()/p:nvSpPr/p:nvPr/p:ph/@type"/>
      </xsl:variable>
      <!--Code by Vijayeta for Bullets,set style name in case of default bullets-->
      <xsl:variable name ="listStyleName">
        <!-- Added by vijayeta, to get the text box number-->
        <xsl:variable name ="textNumber" select ="./parent::node()/p:nvSpPr/p:cNvPr/@id"/>
        <!-- Added by vijayeta, to get the text box number-->
        <xsl:choose>
          <xsl:when test ="./parent::node()/p:spPr/a:prstGeom/@prst or ./parent::node()/child::node()[1]/child::node()[1]/@name[contains(., 'TextBox')] or ./parent::node()/child::node()[1]/child::node()[1]/@name[contains(., 'Text Box')]">
            <xsl:value-of select ="concat('sm1','textboxshape_List',$textNumber)"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select ="concat('sm1','List',position())"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
       <xsl:for-each select ="a:p">
        <!-- Code by vijayeta,to set default font size for inner levels,in case of multiple levels-->
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
        <!--End of Code by vijayeta,to set default font size for inner levels,in case of multiple levels-->
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
            <!-- commented by pradeep -->
            <!--start-->
            <xsl:if test ="a:pPr/@algn ='ctr' or a:pPr/@algn ='r' or a:pPr/@algn ='l' or a:pPr/@algn ='just'">
              <!--End-->
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
              <!-- commented by pradeep -->
              <!--start-->
            </xsl:if >
            <!--End-->
            <!-- Convert Laeft margin of the paragraph-->
            <xsl:if test ="a:pPr/@marL">
              <xsl:attribute name ="fo:margin-left">
                <xsl:value-of select="concat(format-number(a:pPr/@marL div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="a:pPr/@indent">
              <xsl:attribute name ="fo:text-indent">
                <xsl:value-of select="concat(format-number(a:pPr/@indent div 360000, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if >
            <xsl:variable name="var_MaxFontSize">
              <xsl:for-each select="./a:r/a:rPr/@sz">
                <xsl:sort data-type="number" order="descending"/>
                <xsl:if test="position()=1">
                  <xsl:value-of select="."/>
                </xsl:if>
              </xsl:for-each>
            </xsl:variable>
            <xsl:if test ="a:pPr/a:spcBef/a:spcPts/@val">
              <xsl:attribute name ="fo:margin-top">
                <xsl:value-of select="concat(format-number(a:pPr/a:spcBef/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="a:pPr/a:spcBef/a:spcPct/@val">
              <xsl:attribute name ="fo:margin-top">
                <xsl:value-of select="concat(format-number( ( ($var_MaxFontSize * ( a:pPr/a:spcBef/a:spcPct/@val div 100000 ) ) div 2835 ) * 1.2, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="a:pPr/a:spcAft/a:spcPts/@val">
              <xsl:attribute name ="fo:margin-bottom">
                <xsl:value-of select="concat(format-number(a:pPr/a:spcAft/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="a:pPr/a:spcAft/a:spcPct/@val">
              <xsl:attribute name ="fo:margin-bottom">
                <xsl:value-of select="concat(format-number( ( ($var_MaxFontSize * ( a:pPr/a:spcAft/a:spcPct/@val div 100000 ) ) div 2835 ) * 1.2, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <!-- If the line space is in Percentage-->
            <xsl:if test ="a:pPr/a:lnSpc/a:spcPct/@val">
              <xsl:attribute name="fo:line-height">
                <xsl:value-of select="concat(format-number(a:pPr/a:lnSpc/a:spcPct/@val div 1000,'###'), '%')"/>
              </xsl:attribute>
            </xsl:if >
            <!-- If the line space is in Points-->
            <xsl:if test ="a:pPr/a:lnSpc/a:spcPts">
              <xsl:attribute name="style:line-height-at-least">
                <xsl:value-of select="concat(format-number(a:pPr/a:lnSpc/a:spcPts/@val div 2835, '#.##'), 'cm')"/>
              </xsl:attribute>
            </xsl:if>
            <!-- End of Code Added By Vijayeta,to get line spacing from layout or master, dated 9-7-07-->
            <!-- Code inserted by VijayetaFor Bullets, Enable Numbering-->
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
        <!-- Modified by pradeep for fix 1731885-->
        <xsl:for-each select ="node()" >
          <!-- Add here-->
          <xsl:if test ="name()='a:r'">
            <style:style style:name="{generate-id()}"  style:family="text">
              <style:text-properties>
                <!-- Bug 1711910 Fixed,On date 2-06-07,by Vijayeta-->
                <xsl:attribute name ="fo:font-family">
                  <xsl:if test ="a:rPr/a:latin/@typeface">
                    <xsl:variable name ="typeFaceVal" select ="a:rPr/a:latin/@typeface"/>
                    <xsl:for-each select ="a:rPr/a:latin/@typeface">
                      <xsl:if test ="$typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt'">
                        <xsl:value-of select ="$DefFont"/>
                      </xsl:if>
                      <!-- Bug 1711910 Fixed,On date 2-06-07,by Vijayeta-->
                      <xsl:if test ="not($typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt')">
                        <xsl:value-of select ="."/>
                      </xsl:if>
                    </xsl:for-each>
                  </xsl:if>
                  <xsl:if test ="not(a:rPr/a:latin/@typeface)">
                    <!-- ADDED BY VIJAYETA,get font types for each level-->
                    <!--<xsl:for-each select ="document(concat('ppt/slideLayouts/',$layout))//p:sldLayout"-->
                    <xsl:variable name ="cnvPrId">
                      <xsl:value-of select ="parent::node()/parent::node()/parent::node()/p:nvSpPr/p:nvPr/p:ph/@type"/>
                    </xsl:variable>
                    <xsl:variable name ="Id">
                      <xsl:value-of select ="parent::node()/parent::node()/parent::node()/p:nvSpPr/p:nvPr/p:ph/@idx"/>
                    </xsl:variable>
                    <xsl:call-template name ="FontTypeFromLayout">
                      <xsl:with-param name ="prId" select ="$cnvPrId" />
                      <xsl:with-param name ="sldName" select ="''"/>
                      <xsl:with-param name ="Id" select ="$Id"/>
                      <xsl:with-param name ="levelForDefFont" select ="$levelForDefFont"/>
                      <xsl:with-param name ="DefFont" select ="$DefFont"/>
                    </xsl:call-template>
                    <!-- ADDED BY VIJAYETA,get font types for each level-->
                    <!--<xsl:value-of select ="$DefFont"/>-->
                  </xsl:if>
                </xsl:attribute>
                <xsl:attribute name ="style:font-family-generic"	>
                  <xsl:value-of select ="'roman'"/>
                </xsl:attribute>
                <xsl:attribute name ="style:font-pitch"	>
                  <xsl:value-of select ="'variable'"/>
                </xsl:attribute>
                <xsl:if test ="a:rPr/@sz">
                  <xsl:attribute name ="fo:font-size"	>
                    <xsl:for-each select ="a:rPr/@sz">
                      <xsl:value-of select ="concat(format-number(. div 100,'#.##'), 'pt')"/>
                    </xsl:for-each>
                  </xsl:attribute>
                </xsl:if>
                <xsl:attribute name ="fo:font-weight">
                  <!-- Bold Property-->
                  <xsl:if test ="a:rPr/@b">
                    <xsl:value-of select ="'bold'"/>
                  </xsl:if >
                  <xsl:if test ="not(a:rPr/@b)">
                    <xsl:value-of select ="'normal'"/>
                  </xsl:if >
                </xsl:attribute>
                <!-- Color -->
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
                <!-- Bug fix - 1733229-->
                <xsl:if test ="not(a:rPr/a:solidFill/a:srgbClr/@val) and not(a:rPr/a:solidFill/a:schemeClr/@val)">
                  <xsl:if test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr">
                    <xsl:attribute name ="fo:color">
                      <xsl:value-of select ="translate(concat('#',parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:srgbClr/@val),$ucletters,$lcletters)"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="parent::node()/parent::node()/parent::node()/p:style/a:fontRef/a:schemeClr">
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
                  </xsl:if>
                  <xsl:if test ="parent::node()/parent::node()/parent::node()/p:nvSpPr/p:cNvPr/@name[contains(.,'TextBox')]
													   and not(parent::node()/parent::node()/parent::node()/p:style/a:fontRef)">
                    <xsl:attribute name ="fo:color">
                      <xsl:value-of select ="'#000000'"/>
                    </xsl:attribute>
                  </xsl:if>

                </xsl:if>
                <!-- Italic-->
                <xsl:if test ="a:rPr/@i">
                  <xsl:attribute name ="fo:font-style">
                    <xsl:value-of select ="'italic'"/>
                  </xsl:attribute>
                </xsl:if>
                <!-- style:text-underline-style
					        			style:text-underline-style="solid" style:text-underline-width="auto"-->
                <xsl:if test ="a:rPr/@u">
                  <!-- style:text-underline-style="solid" style:text-underline-type="double"-->
                  <xsl:if test ="a:rPr/a:uFill/a:solidFill/a:srgbClr/@val">
                    <xsl:attribute name ="style:text-underline-color">
                      <xsl:value-of select ="translate(concat('#',a:rPr/a:uFill/a:solidFill/a:srgbClr/@val),$ucletters,$lcletters)"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="a:rPr/a:uFill/a:solidFill/a:schemeClr/@val">
                    <xsl:attribute name ="style:text-underline-color">
                      <xsl:call-template name ="getColorCode">
                        <xsl:with-param name ="color">
                          <xsl:value-of select="a:rPr/a:uFill/a:solidFill/a:schemeClr/@val"/>
                        </xsl:with-param>
                        <xsl:with-param name ="lumMod">
                          <xsl:value-of select="a:rPr/a:uFill/a:solidFill/a:schemeClr/a:lumMod/@val"/>
                        </xsl:with-param>
                        <xsl:with-param name ="lumOff">
                          <xsl:value-of select="a:rPr/a:uFill/a:solidFill/a:schemeClr/a:lumOff/@val"/>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:choose >
                    <xsl:when test ="a:rPr/@u='dbl'">
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
                    <xsl:when test ="a:rPr/@u='heavy'">
                      <xsl:attribute name ="style:text-underline-style">
                        <xsl:value-of select ="'solid'"/>
                      </xsl:attribute>
                      <xsl:attribute name ="style:text-underline-width">
                        <xsl:value-of select ="'bold'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test ="a:rPr/@u='dotted'">
                      <xsl:attribute name ="style:text-underline-style">
                        <xsl:value-of select ="'dotted'"/>
                      </xsl:attribute>
                      <xsl:attribute name ="style:text-underline-width">
                        <xsl:value-of select ="'auto'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- dottedHeavy-->
                    <xsl:when test ="a:rPr/@u='dottedHeavy'">
                      <xsl:attribute name ="style:text-underline-style">
                        <xsl:value-of select ="'dotted'"/>
                      </xsl:attribute>
                      <xsl:attribute name ="style:text-underline-width">
                        <xsl:value-of select ="'bold'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test ="a:rPr/@u='dash'">
                      <xsl:attribute name ="style:text-underline-style">
                        <xsl:value-of select ="'dash'"/>
                      </xsl:attribute>
                      <xsl:attribute name ="style:text-underline-width">
                        <xsl:value-of select ="'auto'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test ="a:rPr/@u='dashHeavy'">
                      <xsl:attribute name ="style:text-underline-style">
                        <xsl:value-of select ="'dash'"/>
                      </xsl:attribute>
                      <xsl:attribute name ="style:text-underline-width">
                        <xsl:value-of select ="'bold'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test ="a:rPr/@u='dashLong'">
                      <xsl:attribute name ="style:text-underline-style">
                        <xsl:value-of select ="'long-dash'"/>
                      </xsl:attribute>
                      <xsl:attribute name ="style:text-underline-width">
                        <xsl:value-of select ="'auto'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test ="a:rPr/@u='dashLongHeavy'">
                      <xsl:attribute name ="style:text-underline-style">
                        <xsl:value-of select ="'long-dash'"/>
                      </xsl:attribute>
                      <xsl:attribute name ="style:text-underline-width">
                        <xsl:value-of select ="'bold'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test ="a:rPr/@u='dotDash'">
                      <xsl:attribute name ="style:text-underline-style">
                        <xsl:value-of select ="'dot-dash'"/>
                      </xsl:attribute>
                      <xsl:attribute name ="style:text-underline-width">
                        <xsl:value-of select ="'auto'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test ="a:rPr/@u='dotDashHeavy'">
                      <xsl:attribute name ="style:text-underline-style">
                        <xsl:value-of select ="'dot-dash'"/>
                      </xsl:attribute>
                      <xsl:attribute name ="style:text-underline-width">
                        <xsl:value-of select ="'bold'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- dot-dot-dash-->
                    <xsl:when test ="a:rPr/@u='dotDotDash'">
                      <xsl:attribute name ="style:text-underline-style">
                        <xsl:value-of select ="'dot-dot-dash'"/>
                      </xsl:attribute>
                      <xsl:attribute name ="style:text-underline-width">
                        <xsl:value-of select ="'auto'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test ="a:rPr/@u='dotDotDashHeavy'">
                      <xsl:attribute name ="style:text-underline-style">
                        <xsl:value-of select ="'dot-dot-dash'"/>
                      </xsl:attribute>
                      <xsl:attribute name ="style:text-underline-width">
                        <xsl:value-of select ="'bold'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- Wavy and Heavy-->
                    <xsl:when test ="a:rPr/@u='wavy'">
                      <xsl:attribute name ="style:text-underline-style">
                        <xsl:value-of select ="'wave'"/>
                      </xsl:attribute>
                      <xsl:attribute name ="style:text-underline-width">
                        <xsl:value-of select ="'auto'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test ="a:rPr/@u='wavyHeavy'">
                      <xsl:attribute name ="style:text-underline-style">
                        <xsl:value-of select ="'wave'"/>
                      </xsl:attribute>
                      <xsl:attribute name ="style:text-underline-width">
                        <xsl:value-of select ="'bold'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- wavyDbl-->
                    <!-- style:text-underline-style="wave" style:text-underline-type="double"-->
                    <xsl:when test ="a:rPr/@u='wavyDbl'">
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
                    <xsl:otherwise >
                      <xsl:attribute name ="style:text-underline-type">
                        <xsl:value-of select ="'single'"/>
                      </xsl:attribute>
                      <xsl:attribute name ="style:text-underline-width">
                        <xsl:value-of select ="'auto'"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
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
                <!-- Character Spacing fo:letter-spacing Bug (1719230) fix by Sateesh -->
                <xsl:if test ="a:rPr/@spc">
                  <xsl:attribute name ="fo:letter-spacing">
                    <xsl:variable name="length" select="a:rPr/@spc" />
                    <xsl:value-of select="concat(format-number($length * 2.54 div 7200,'#.###'),'cm')"/>
                  </xsl:attribute>
                </xsl:if>
                <!--Shadow fo:text-shadow-->
                <xsl:if test ="a:rPr/a:effectLst/a:outerShdw">
                  <xsl:attribute name ="fo:text-shadow">
                    <xsl:value-of select ="'1pt 1pt'"/>
                  </xsl:attribute>
                </xsl:if>
                <!--Kerning true or false -->
                <xsl:attribute name ="style:letter-kerning">
                  <xsl:choose >
                    <xsl:when test ="a:rPr/@kern = '0'">
                      <xsl:value-of select ="'false'"/>
                    </xsl:when >
                    <xsl:when test ="format-number(a:rPr/@kern,'#.##') &gt; 0">
                      <xsl:value-of select ="'true'"/>
                    </xsl:when >
                    <xsl:otherwise >
                      <xsl:value-of select ="'true'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute >
              </style:text-properties>
            </style:style>
          </xsl:if>
          <!-- Added by lohith.ar for fix 1731885-->
          <xsl:if test ="name()='a:endParaRPr'">
            <style:style style:name="{generate-id()}"  style:family="text">
              <style:text-properties>
                <!-- Bug 1711910 Fixed,On date 2-06-07,by Vijayeta-->
                <xsl:attribute name ="fo:font-family">
                  <xsl:if test ="a:endParaRPr/a:latin/@typeface">
                    <xsl:variable name ="typeFaceVal" select ="a:rPr/a:latin/@typeface"/>
                    <xsl:for-each select ="a:endParaRPr/a:latin/@typeface">
                      <xsl:if test ="$typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt'">
                        <xsl:value-of select ="$DefFont"/>
                      </xsl:if>
                      <!-- Bug 1711910 Fixed,On date 2-06-07,by Vijayeta-->
                      <xsl:if test ="not($typeFaceVal='+mn-lt' or $typeFaceVal='+mj-lt')">
                        <xsl:value-of select ="."/>
                      </xsl:if>
                    </xsl:for-each>
                  </xsl:if>
                  <xsl:if test ="not(a:endParaRPr/a:latin/@typeface)">
                    <xsl:value-of select ="$DefFont"/>
                  </xsl:if>
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
              </style:text-properties>
            </style:style>
          </xsl:if>
        </xsl:for-each >
      </xsl:for-each>
    </xsl:for-each>
  </xsl:template>
  <!--End-->
</xsl:stylesheet>

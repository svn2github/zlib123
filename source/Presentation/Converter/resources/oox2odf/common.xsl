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
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"  
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  exclude-result-prefixes="a style svg fo">
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

</xsl:stylesheet>

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" 
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"  
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" 
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  exclude-result-prefixes="style draw fo text">
	<xsl:import href ="BulletsNumbering.xsl"/>
	<xsl:template name ="convertToPoints">
		<xsl:param name ="unit"/>
		<xsl:param name ="length"/>
		<xsl:message terminate="no">progress:text:p</xsl:message>
		<xsl:variable name="lengthVal">
			<xsl:choose>
				<xsl:when test="contains($length,'cm')">
					<xsl:value-of select="substring-before($length,'cm')"/>
				</xsl:when>
				<xsl:when test="contains($length,'pt')">
					<xsl:value-of select="substring-before($length,'pt')"/>
				</xsl:when>
				<xsl:when test="contains($length,'in')">
					<xsl:choose>
						<xsl:when test ="$unit='cm'" >
							<xsl:value-of select="substring-before($length,'in') * 2.54 "/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select="substring-before($length,'in')" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<!--mm to cm -->
				<xsl:when test="contains($length,'mm')">
					<xsl:choose>
						<xsl:when test ="$unit='cm'" >
							<xsl:value-of select="substring-before($length,'mm') div 10 "/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select="substring-before($length,'mm')" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<!-- m to cm -->
				<xsl:when test="contains($length,'m')">
					<xsl:choose>
						<xsl:when test ="$unit='cm'" >
							<xsl:value-of select="substring-before($length,'m') * 100 "/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select="substring-before($length,'m')" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<!-- km to cm -->
				<xsl:when test="contains($length,'km')">
					<xsl:choose>
						<xsl:when test ="$unit='cm'" >
							<xsl:value-of select="substring-before($length,'km') * 100000  "/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select="substring-before($length,'km')" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<!-- mi to cm -->
				<xsl:when test="contains($length,'mi')">
					<xsl:choose>
						<xsl:when test ="$unit='cm'" >
							<xsl:value-of select="substring-before($length,'mi') * 160934.4"/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select="substring-before($length,'mi')" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<!-- ft to cm -->
				<xsl:when test="contains($length,'ft')">
					<xsl:choose>
						<xsl:when test ="$unit='cm'" >
							<xsl:value-of select="substring-before($length,'ft') * 30.48 "/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select="substring-before($length,'ft')" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<!-- em to cm -->
				<xsl:when test="contains($length,'em')">
					<xsl:choose>
						<xsl:when test ="$unit='cm'" >
							<xsl:value-of select="round(substring-before($length,'em') * .4233) "/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select="substring-before($length,'em')" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<!-- px to cm -->
				<xsl:when test="contains($length,'px')">
					<xsl:choose>
						<xsl:when test ="$unit='cm'" >
							<xsl:value-of select="round(substring-before($length,'px') div 35.43307) "/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select="substring-before($length,'px')" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<!-- pc to cm -->
				<xsl:when test="contains($length,'pc')">
					<xsl:choose>
						<xsl:when test ="$unit='cm'" >
							<xsl:value-of select="round(substring-before($length,'pc') div 2.362) "/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select="substring-before($length,'pc')" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<!-- ex to cm 1 ex to 6 px-->
				<xsl:when test="contains($length,'ex')">
					<xsl:choose>
						<xsl:when test ="$unit='cm'" >
							<xsl:value-of select="round((substring-before($length,'ex') div 35.43307)* 6) "/>
						</xsl:when>
						<xsl:otherwise >
							<xsl:value-of select="substring-before($length,'ex')" />
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="$length"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test ="contains($lengthVal,'NaN')">
				<xsl:value-of select ="0"/>
			</xsl:when>
			<xsl:when test="$lengthVal='0' or $lengthVal='' or ( ($lengthVal &lt; 0) and ($unit != 'cm')) ">
				<xsl:value-of select="0"/>
			</xsl:when>
			<xsl:when test="$unit = 'cm'">
				<xsl:value-of select="concat(format-number($lengthVal * 360000,'#'),'')"/>
			</xsl:when>
			<xsl:when test="$unit = 'mm'">
				<xsl:value-of select="concat(format-number($lengthVal * 25.4 div 72,'#.###'),'mm')"/>
			</xsl:when>
			<xsl:when test="$unit = 'in'">
				<xsl:value-of select="concat(format-number($lengthVal div 72,'#.###'),'in')"/>
			</xsl:when>
			<xsl:when test="$unit = 'pt'">
				<xsl:value-of select="concat(format-number($lengthVal,'#') * 100 ,'')"/>
				<!--Added by lohith - format-number($lengthVal,'#') to make sure that pt will be a int not a real value-->
			</xsl:when>
			<xsl:when test="$unit = 'pica'">
				<xsl:value-of select="concat(format-number($lengthVal div 12,'#.###'),'pica')"/>
			</xsl:when>
			<xsl:when test="$unit = 'dpt'">
				<xsl:value-of select="concat($lengthVal,'dpt')"/>
			</xsl:when>
			<xsl:when test="$unit = 'px'">
				<xsl:value-of select="concat(format-number($lengthVal * 96.19 div 72,'#.###'),'px')"/>
			</xsl:when>
			<xsl:when test="not($lengthVal)">
				<xsl:value-of select="concat(0,'cm')"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$lengthVal"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template >
	<xsl:template name ="fontStyles">
		<xsl:param name ="Tid"/>
		<xsl:param name ="prClassName"/>
		<xsl:param name ="lvl"/>
		<!-- Parameter Added by Vijayeta, on 13-7-07-->
		<xsl:param name ="masterPageName"/>
		<xsl:param name="slideMaster" />
		<xsl:message terminate="no">progress:text:p</xsl:message>
		<xsl:variable name ="fileName">
			<xsl:if test ="$slideMaster !=''">
				<xsl:value-of select ="$slideMaster"/>
			</xsl:if>
			<xsl:if test ="$slideMaster =''">
				<xsl:value-of select ="'content.xml'"/>
			</xsl:if >
		</xsl:variable >

		<xsl:for-each  select ="document($fileName)//office:automatic-styles/style:style[@style:name =$Tid ]">
			<xsl:message terminate="no">progress:text:p</xsl:message>
			<!-- Added by lohith :substring-before(style:text-properties/@fo:font-size,'pt')&gt; 0  because sz(font size) shouldnt be zero - 16filesbug-->
			<xsl:if test="style:text-properties/@fo:font-size and substring-before(style:text-properties/@fo:font-size,'pt')&gt; 0 ">
				<xsl:attribute name ="sz">
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'pt'"/>
						<xsl:with-param name ="length" select ="style:text-properties/@fo:font-size"/>
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>
			<!--Superscript and SubScript for Text added by Mathi on 31st Jul 2007-->
			<xsl:if test="style:text-properties/@style:text-position">
				<xsl:choose>
					<xsl:when test="(substring-before(style:text-properties/@style:text-position,' ') = 'super')">
						<xsl:attribute name="baseline">
							<xsl:variable name="blsuper">
								<xsl:value-of select="substring-before(substring-after(style:text-properties/@style:text-position,'super '),'%')"/>
							</xsl:variable>
							<xsl:value-of select="($blsuper * 1000)"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test="(substring-before(style:text-properties/@style:text-position,' ') = 'sub')">
						<xsl:attribute name="baseline">
							<xsl:variable name="blsub">
								<xsl:value-of select="substring-before(substring-after(style:text-properties/@style:text-position,'sub '),'%')"/>
							</xsl:variable>
							<xsl:value-of select="($blsub * (-1000))"/>
						</xsl:attribute>
					</xsl:when>
				</xsl:choose>
			</xsl:if>

			<xsl:if test ="not(style:text-properties/@fo:font-size)">
        <!-- added by vipul no check for default font for title, outline, and subtitle-->
        <xsl:if test="$prClassName !='outline' and $prClassName !='subtitle' and $prClassName !='title'">
				<!-- Coded by vijayeta,set font size depending on levels-->
				<xsl:variable name="GetDefaultFontSizeForShape">
					<xsl:call-template name ="getDefaultFontSize">
						<xsl:with-param name ="className" select ="$prClassName"/>
						<!-- Node added by vijayeta,to insert font sizes to inner levels-->
						<xsl:with-param name ="lvl" select ="$lvl+1"/>
					</xsl:call-template >
				</xsl:variable>
				<xsl:if test ="substring-before($GetDefaultFontSizeForShape, 'pt') > 0">
					<xsl:attribute name ="sz">
						<xsl:call-template name ="convertToPoints">
							<xsl:with-param name ="unit" select ="'pt'"/>
							<xsl:with-param name ="length" select ="$GetDefaultFontSizeForShape"/>
						</xsl:call-template>
					</xsl:attribute>
				</xsl:if>
			</xsl:if>
      </xsl:if>
			<!--Font bold attribute -->
			<xsl:if test="style:text-properties/@fo:font-weight[contains(.,'bold')]">
				<xsl:attribute name ="b">
					<xsl:value-of select ="'1'"/>
				</xsl:attribute >
			</xsl:if >
		
			<!-- Kerning - Added by lohith.ar -->
			<!-- Start -->
			<xsl:if test ="style:text-properties/@style:letter-kerning = 'true'">
				<xsl:attribute name ="kern">
					<xsl:value-of select="1200"/>
				</xsl:attribute>
			</xsl:if>
			<xsl:if test ="style:text-properties/@style:letter-kerning = 'false'">
				<xsl:attribute name ="kern">
					<xsl:value-of select="0"/>
				</xsl:attribute>
			</xsl:if>
			<!--<xsl:if test ="not(style:text-properties/@style:letter-kerning)">
        <xsl:attribute name ="kern">
          <xsl:value-of select="0"/>
        </xsl:attribute>
      </xsl:if>-->
			<!-- End -->
			<!-- Font Inclined-->
			<xsl:if test="style:text-properties/@fo:font-style[contains(.,'italic')]">
				<xsl:attribute name ="i">
					<xsl:value-of select ="'1'"/>
				</xsl:attribute >
			</xsl:if >
			
			<!-- Font underline-->
			<!-- Font underline-->

			<xsl:choose >
				<!-- Added by lohith for fix - 1744082 - Start-->
				<xsl:when test="style:text-properties/@style:text-underline-type = 'single'">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'sng'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- Fix - 1744082 - End-->
				<xsl:when test="style:text-properties/@style:text-underline-style = 'solid' and
								style:text-properties/@style:text-underline-type[contains(.,'double')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dbl'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style  = 'solid' and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'heavy'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style = 'solid' and
							style:text-properties/@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'sng'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- Dotted lean and dotted bold under line -->
				<xsl:when test="style:text-properties/@style:text-underline-style = 'dotted' and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dotted'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style = 'dotted' and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dottedHeavy'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- Dash lean and dash bold underline -->
				<xsl:when test="style:text-properties/@style:text-underline-style = 'dash' and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dash'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style = 'dash' and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dashHeavy'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- Dash long and dash long bold -->
				<xsl:when test="style:text-properties/@style:text-underline-style = 'long-dash' and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dashLong'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style = 'long-dash' and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dashLongHeavy'"/>
					</xsl:attribute >
				</xsl:when>

				<!-- dot Dash and dot dash bold -->
				<xsl:when test="style:text-properties/@style:text-underline-style = 'dot-dash' and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dotDash'"/>
						<!-- Modified by lohith for fix 1739785 - dotDashLong to dotDash-->
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style = 'dot-dash' and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dotDashHeavy'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- dot-dot-dash-->
				<xsl:when test="style:text-properties/@style:text-underline-style= 'dot-dot-dash' and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dotDotDash'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style= 'dot-dot-dash' and
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
			<!-- Font Strike through Start-->
			<xsl:choose >
				<xsl:when  test="style:text-properties/@style:text-line-through-type = 'solid'">
					<xsl:attribute name ="strike">
						<xsl:value-of select ="'sngStrike'"/>
					</xsl:attribute >
				</xsl:when >
				<xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
					<xsl:attribute name ="strike">
						<xsl:value-of select ="'dblStrike'"/>
					</xsl:attribute >
				</xsl:when >
				<!-- style:text-line-through-style-->
				<xsl:when test="style:text-properties/@style:text-line-through-style = 'solid'">
					<xsl:attribute name ="strike">
						<xsl:value-of select ="'sngStrike'"/>
					</xsl:attribute >
				</xsl:when>
			</xsl:choose>
			<!-- Font Strike through end-->
			<!--Charector spacing -->
			<!-- Modfied by lohith - @fo:letter-spacing will have a text value 'normal' when no change is required -->
			<xsl:if test ="style:text-properties/@fo:letter-spacing [contains(.,'cm')]">
				<!-- Modfied by lohith - "spc" should be a int value, '#.##'has been replaced by '#'   -->
				<xsl:attribute name ="spc">
					<xsl:if test ="substring-before(style:text-properties/@fo:letter-spacing,'cm')&lt; 0 ">
						<!--<xsl:value-of select ="format-number(substring-before(style:text-properties/@fo:letter-spacing,'cm') * 3.5 *1000 ,'#')"/>-->
						<xsl:value-of select ="format-number(substring-before(style:text-properties/@fo:letter-spacing,'cm') * 7200 div 2.54 ,'#')"/>
					</xsl:if >
					<xsl:if test ="substring-before(style:text-properties/@fo:letter-spacing,'cm')
								&gt; 0 or substring-before(style:text-properties/@fo:letter-spacing,'cm') = 0 ">
						<!--<xsl:value-of select ="format-number((substring-before(style:text-properties/@fo:letter-spacing,'cm') div .035) *100 ,'#')"/>-->
						<xsl:value-of select ="format-number((substring-before(style:text-properties/@fo:letter-spacing,'cm') * 72 div 2.54) *100 ,'#')"/>
					</xsl:if>
				</xsl:attribute>
			</xsl:if >
			<!--Color Node set as standard colors -->
			<xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
			<xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
			<xsl:if test ="style:text-properties/@fo:color">
				<a:solidFill>
					<a:srgbClr  >
						<xsl:attribute name ="val">
							<!--<xsl:value-of   select ="substring-after(style:text-properties/@fo:color,'#')"/>-->
							<xsl:value-of select ="translate(substring-after(style:text-properties/@fo:color,'#'),$lcletters,$ucletters)"/>
						</xsl:attribute>
					</a:srgbClr >
				</a:solidFill>
			</xsl:if>

			<!-- Text Shadow fix -->
			<xsl:if test ="style:text-properties/@fo:text-shadow != 'none'">
				<a:effectLst>
					<a:outerShdw blurRad="38100" dist="38100" dir="2700000" >
						<a:srgbClr val="000000">
							<a:alpha val="43137" />
						</a:srgbClr>
					</a:outerShdw>
				</a:effectLst>
			</xsl:if>

			<xsl:if test ="style:text-properties/@fo:font-family">
				<a:latin charset="0" >
					<xsl:attribute name ="typeface" >
						<!-- fo:font-family-->
						<xsl:value-of select ="translate(style:text-properties/@fo:font-family, &quot;'&quot;,'')" />
					</xsl:attribute>
				</a:latin >
			</xsl:if>
			<!--Commented by Vipul As no need to assign defualt font as Arial-->
			<!--Start-->
			<xsl:if test ="not(style:text-properties/@fo:font-family) ">
        <!-- added by vipul no check for default font for title, outline, and subtitle-->
        <xsl:if test="$prClassName !='outline' and $prClassName !='subtitle' and $prClassName !='title'">
				<xsl:if test ="$prClassName='standard'">
					<a:latin charset="0" >
						<xsl:attribute name ="typeface" >
							<xsl:call-template name ="getDefaultFonaName">
								<xsl:with-param name ="className" select ="$prClassName"/>
								<!--Node added by vijayeta,to insert font sizes to inner levels-->
								<xsl:with-param name ="lvl" select ="$lvl+1"/>
							</xsl:call-template >
						</xsl:attribute >
					</a:latin >
				</xsl:if>
        </xsl:if>
			</xsl:if>
			<!--End-->
			<!-- Underline color -->
			<xsl:if test ="style:text-properties/style:text-underline-color">
				<a:uFill>
					<a:solidFill>
						<a:srgbClr>
							<xsl:attribute name ="val">
								<xsl:value-of select ="substring-after(style:text-properties/style:text-underline-color,'#')"/>
							</xsl:attribute>
						</a:srgbClr>
					</a:solidFill>
				</a:uFill>
			</xsl:if>
		</xsl:for-each >
	</xsl:template>
	<xsl:template name ="paraProperties" >
		<!--- Code inserted by Vijayeta for Bullets and numbering,For bullet properties-->
		<xsl:param name ="paraId" />
		<xsl:param name ="listId"/>
		<xsl:param name ="isBulleted" />
		<xsl:param name ="level"/>
		<xsl:param name ="isNumberingEnabled" />
		<xsl:param name ="framePresentaionStyleId"/>
		<!-- parameter added by vijayeta, dated 13-7-07-->
		<xsl:param name ="masterPageName"/>
		<xsl:param name="slideMaster" />
		<xsl:message terminate="no">progress:text:p</xsl:message>
		<xsl:variable name ="fileName">
			<xsl:if test ="$slideMaster !=''">
				<xsl:value-of select ="$slideMaster"/>
			</xsl:if>
			<xsl:if test ="$slideMaster =''">
				<xsl:value-of select ="'content.xml'"/>
			</xsl:if >
		</xsl:variable >
		<xsl:for-each select ="document($fileName)//style:style[@style:name=$paraId]">
			<xsl:message terminate="no">progress:text:p</xsl:message>
			<a:pPr>
				<!-- Code inserted by Vijayeta for Bullets and numbering,For bullet properties-->
				<xsl:if test ="not($level='0')">
					<xsl:attribute name ="lvl">
						<xsl:value-of select ="$level"/>
					</xsl:attribute>
				</xsl:if>
				<!-- Added by vijayeta for teh fix 1739081-->
				<xsl:if test ="$isNumberingEnabled='true'">
					<xsl:variable name ="marL">
						<xsl:call-template name="MarginTemplateForSlide">
							<xsl:with-param name="level" select="$level"/>
							<xsl:with-param name ="listId" select ="$listId"/>
							<xsl:with-param name ="fileName" select ="$fileName"/>
						</xsl:call-template >
					</xsl:variable>
					<xsl:if test ="$marL!=''">
						<xsl:attribute name ="marL">
							<xsl:value-of select ="$marL"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:variable name ="indetValue">
						<xsl:call-template name="IndentTemplateForSlide">
							<xsl:with-param name="level" select="$level"/>
							<xsl:with-param name ="listId" select ="$listId"/>
							<xsl:with-param name ="fileName" select ="$fileName"/>
						</xsl:call-template >
					</xsl:variable>
					<xsl:if test ="$indetValue!=''">
						<xsl:attribute name="indent">
							<xsl:value-of select="$indetValue"/>
						</xsl:attribute>
					</xsl:if>
				</xsl:if>
				<!-- Added by vijayeta for teh fix 1739081-->
				<!--marL="first line indent property"-->
				<xsl:if test ="style:paragraph-properties/@fo:text-indent 
							and substring-before(style:paragraph-properties/@fo:text-indent,'cm') != 0">

					<!--fo:text-indent-->
					<xsl:variable name ="varIndent">
						<xsl:call-template name ="convertToPoints">
							<xsl:with-param name="length"  select ="style:paragraph-properties/@fo:text-indent"/>
							<xsl:with-param name ="unit" select ="'cm'"/>
						</xsl:call-template>
					</xsl:variable>
					<xsl:if test ="$varIndent!=''">
						<xsl:attribute name ="indent">
							<xsl:value-of select ="$varIndent"/>
						</xsl:attribute>
					</xsl:if>
				</xsl:if >
				<xsl:if test ="style:paragraph-properties/@fo:text-align">
					<xsl:attribute name ="algn">
						<!--fo:text-align-->
						<xsl:choose >
							<xsl:when test ="style:paragraph-properties/@fo:text-align='center'">
								<xsl:value-of select ="'ctr'"/>
							</xsl:when>
							<xsl:when test ="style:paragraph-properties/@fo:text-align='end'">
								<xsl:value-of select ="'r'"/>
							</xsl:when>
							<xsl:when test ="style:paragraph-properties/@fo:text-align='justify'">
								<xsl:value-of select ="'just'"/>
							</xsl:when>
							<xsl:otherwise >
								<xsl:value-of select ="'l'"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
				</xsl:if >
				<!-- Added by Lohith - to set the text alignment using frame properties-->
				<xsl:if test ="not(style:paragraph-properties/@fo:text-align)">
					<xsl:for-each select ="document('content.xml')//style:style[@style:name=$framePresentaionStyleId]">
						<xsl:if test="style:graphic-properties/@draw:textarea-horizontal-align = 'left'">
							<xsl:attribute name ="algn">
								<xsl:value-of select ="'l'"/>
							</xsl:attribute>
						</xsl:if>
						<xsl:if test="style:graphic-properties/@draw:textarea-horizontal-align = 'right'">
							<xsl:attribute name ="algn">
								<xsl:value-of select ="'r'"/>
							</xsl:attribute>
						</xsl:if>
						<xsl:if test="style:graphic-properties/@draw:textarea-horizontal-align = 'center'">
							<xsl:attribute name ="algn">
								<xsl:value-of select ="'ctr'"/>
							</xsl:attribute>
						</xsl:if>
					</xsl:for-each>
				</xsl:if >

				<xsl:if test ="style:paragraph-properties/@fo:margin-left and 
							   substring-before(style:paragraph-properties/@fo:margin-left,'cm') &gt; 0">

					<!--fo:margin-left-->
					<xsl:variable name ="varMarginLeft">
						<xsl:call-template name ="convertToPoints">
							<xsl:with-param name="length"  select ="style:paragraph-properties/@fo:margin-left"/>
							<xsl:with-param name ="unit" select ="'cm'"/>
						</xsl:call-template>
					</xsl:variable>
					<xsl:if test ="$varMarginLeft!=''">
						<xsl:attribute name ="marL">
							<xsl:value-of select ="$varMarginLeft"/>
						</xsl:attribute>
					</xsl:if>
				</xsl:if >

				<!--Code inserted by Vijayeta For Line Spacing,
            If the line spacing is in terms of Percentage, multiply the value with 1000-->
				<xsl:if test ="style:paragraph-properties/@fo:line-height and 
					substring-before(style:paragraph-properties/@fo:line-height,'%') &gt; 0 and 
					not(substring-before(style:paragraph-properties/@fo:line-height,'%') = 100)">
					<a:lnSpc>
						<a:spcPct>
							<xsl:attribute name ="val">
								<xsl:value-of select ="format-number(substring-before(style:paragraph-properties/@fo:line-height,'%')* 1000,'#.##') "/>
							</xsl:attribute>
						</a:spcPct>
					</a:lnSpc>
				</xsl:if>
				<!--If the line spacing is in terms of Points,multiply the value with 2835-->
				<xsl:if test ="style:paragraph-properties/@style:line-spacing and 
					substring-before(style:paragraph-properties/@style:line-spacing,'cm') &gt; 0">
					<a:lnSpc>
						<a:spcPts>
							<xsl:attribute name ="val">
								<xsl:value-of select ="round(substring-before(style:paragraph-properties/@style:line-spacing,'cm')* 2835) "/>
							</xsl:attribute>
						</a:spcPts>
					</a:lnSpc>
				</xsl:if>
				<xsl:if test ="style:paragraph-properties/@style:line-height-at-least and 
					substring-before(style:paragraph-properties/@style:line-height-at-least,'cm') &gt; 0 ">
					<a:lnSpc>
						<a:spcPts>
							<xsl:attribute name ="val">
								<xsl:value-of select ="round(substring-before(style:paragraph-properties/@style:line-height-at-least,'cm')* 2835) "/>
							</xsl:attribute>
						</a:spcPts>
					</a:lnSpc>
				</xsl:if>
				<!--End of Code inserted by Vijayeta For Line Spacing -->
				<!-- Code Added by Vijayeta,for Paragraph Spacing, Before and After
             Multiply the value in cm with 2835
			 date: on 01-06-07-->
				<xsl:if test ="style:paragraph-properties/@fo:margin-top and 
						substring-before(style:paragraph-properties/@fo:margin-top,'cm') &gt; 0 ">
					<a:spcBef>
						<a:spcPts>
							<xsl:attribute name ="val">
								<!--fo:margin-top-->
								<xsl:value-of select ="round(substring-before(style:paragraph-properties/@fo:margin-top,'cm')* 2835) "/>
							</xsl:attribute>
						</a:spcPts>
					</a:spcBef >
				</xsl:if>
				<xsl:if test ="style:paragraph-properties/@fo:margin-bottom and 
					    substring-before(style:paragraph-properties/@fo:margin-bottom,'cm') &gt; 0 ">
					<a:spcAft>
						<a:spcPts>
							<xsl:attribute name ="val">
								<!--fo:margin-bottom-->
								<xsl:value-of select ="round(substring-before(style:paragraph-properties/@fo:margin-bottom,'cm')* 2835) "/>
							</xsl:attribute>
						</a:spcPts>
					</a:spcAft>
				</xsl:if >
				<!-- Code Added by Vijayeta,for Paragraph Spacing, Before and After-->
				<!--<xsl:if test ="isBulleted='false'">
				<a:buNone/>
        </xsl:if>-->
				<xsl:if test ="$isNumberingEnabled='false'">
					<a:buNone/>
				</xsl:if>
				<xsl:if test ="$isBulleted='true'">
					<xsl:if test ="$isNumberingEnabled='true'">
						<xsl:call-template name ="insertBulletsNumbers" >
							<xsl:with-param name ="listId" select ="$listId"/>
							<xsl:with-param name ="level" select ="$level+1"/>
							<!-- parameter added by vijayeta, dated 13-7-07-->
							<xsl:with-param name ="masterPageName" select ="$masterPageName"/>
						</xsl:call-template>
					</xsl:if>
				</xsl:if>

				<!--Code Inserted by vijayeta,For Bullets and Numbering,Set Level if present-->
				<!-- @@ Code for paragraph tabs -Start-->
				<xsl:call-template name ="paragraphTabstops"/>
				<!-- @@ Code for paragraph tabs -End-->
			</a:pPr>
		</xsl:for-each >
	</xsl:template>
	<xsl:template name ="fillColor">
		<xsl:param name ="prId"/>
		<xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
		<xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
		<xsl:for-each select ="document('content.xml')/office:document-content/office:automatic-styles/style:style[@style:name=$prId] ">
			<!--test="not(style:graphic-properties/@draw:fill = 'none' - Added by lohith.ar for invalid fill color for textboxes - Fill type should be given priority on fill color-->
			<xsl:if test ="style:graphic-properties/@draw:fill-color">
				<a:solidFill>
					<a:srgbClr  >
						<xsl:attribute name ="val">
							<xsl:value-of select ="translate(substring-after(style:graphic-properties/@draw:fill-color,'#'),$lcletters,$ucletters)"/>
						</xsl:attribute>
						<xsl:if test ="not(style:graphic-properties/@draw:fill)">
							<a:alpha val="0" />
						</xsl:if>
						<xsl:if test ="style:graphic-properties/@draw:fill ='none'">
							<a:alpha val="0" />
						</xsl:if>
					</a:srgbClr >
				</a:solidFill>
			</xsl:if>
		</xsl:for-each>
	</xsl:template>
	<xsl:template name ="getClassName">
		<xsl:param name ="clsName"/>
		<!-- Node added by vijayeta,to insert font sizes to inner levels-->
		<xsl:param name ="lvl"/>
		<xsl:param name ="masterPageName"/>
		<xsl:choose >
			<xsl:when test ="$clsName='title'">
				<xsl:choose>
					<xsl:when test="$masterPageName='Standard'">
						<xsl:value-of select ="'Standard-title'"/>
					</xsl:when>
					<xsl:when test="$masterPageName='Default'">
						<xsl:value-of select ="'Default-title'"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select ="concat($masterPageName,'-title')"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:when test ="$clsName='subtitle'">
				<xsl:choose>
					<xsl:when test="$masterPageName='Standard'">
						<xsl:value-of select ="'Standard-subtitle'"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select ="concat($masterPageName,'-subtitle')"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<!-- By vijayeta class name s in stylea.xml for differant levels-->
			<xsl:when test ="$clsName='outline'">
				<xsl:choose>
					<xsl:when test="$masterPageName='Standard'">
						<xsl:value-of select ="concat('Standard-outline',$lvl+1)"/>
					</xsl:when>
					<xsl:when test="$masterPageName='Default'">
						<xsl:value-of select ="concat('Default-outline',$lvl+11)"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select ="concat($masterPageName,'-outline',$lvl+1)"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when >
			<xsl:when test="$clsName='standard'">
				<xsl:value-of select ="'standard'"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select ="$clsName"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name ="getDefaultFonaName">
		<xsl:param name ="className"/>
		<xsl:param name ="lvl"/>
		<xsl:param name ="masterPageName" />
		<xsl:variable name ="defaultClsName">
			<xsl:call-template name ="getClassName">
				<xsl:with-param name ="clsName" select="$className"/>
				<!-- Node added by vijayeta,to insert font sizes to inner levels-->
				<xsl:with-param name ="masterPageName" select ="$masterPageName"/>
				<xsl:with-param name ="lvl" select ="$lvl"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:choose >
			<xsl:when test ="document('styles.xml')//style:style[@style:name = $defaultClsName]/style:paragraph-properties/@fo:font-family">
				<xsl:variable name ="FontName">
					<xsl:value-of select ="document('styles.xml')//style:style[@style:name = $defaultClsName]/style:paragraph-properties/@fo:font-family"/>
				</xsl:variable>
				<xsl:value-of select ="translate($FontName, &quot;'&quot;,'')" />
			</xsl:when>
			<xsl:when test ="document('styles.xml')//style:style[@style:name = $defaultClsName]/style:text-properties/@fo:font-family">
				<xsl:variable name ="FontName">
					<xsl:value-of select ="document('styles.xml')//style:style[@style:name = $defaultClsName]/style:text-properties/@fo:font-family"/>
				</xsl:variable >
				<xsl:value-of select ="translate($FontName, &quot;'&quot;,'')" />
			</xsl:when>
			<!-- Added by lohith - to access default Font family-->
			<xsl:when test ="$defaultClsName='standard'">
				<xsl:variable name ="shapeFontName">
					<xsl:value-of select ="document('styles.xml')//style:style[@style:name = $defaultClsName]/style:text-properties/@fo:font-family"/>
				</xsl:variable>
				<xsl:value-of select ="translate($shapeFontName, &quot;'&quot;,'')" />
			</xsl:when>
			<xsl:when test ="not(document('styles.xml')//style:style[@style:name = $defaultClsName]/style:text-properties/@fo:font-family)">
				<xsl:variable name ="parentFontName">
					<xsl:value-of select ="document('styles.xml')//style:style[@style:name = $defaultClsName]/@style:parent-style-name"/>
				</xsl:variable>
				<xsl:variable name ="shapeFontName">
					<xsl:value-of select ="document('styles.xml')//style:style[@style:name = $parentFontName]/style:text-properties/@fo:font-family"/>
				</xsl:variable>
				<xsl:if test ="$shapeFontName !=''">
					<xsl:value-of select ="translate($shapeFontName, &quot;'&quot;,'')" />
				</xsl:if>
				<xsl:if test ="$shapeFontName =''">
					<xsl:value-of select ="'Arial'"/>
				</xsl:if>
			</xsl:when>
			<xsl:otherwise >
				<xsl:value-of select ="'Arial'"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name ="getDefaultFontSize">
		<xsl:param name ="className"/>
		<xsl:param name ="lvl"/>
		<xsl:param name ="masterPageName" />
		<xsl:message terminate="no">progress:text:p</xsl:message>
		<xsl:variable name ="defaultClsName">
			<xsl:call-template name ="getClassName">
				<xsl:with-param name ="clsName" select="$className"/>
				<xsl:with-param name ="masterPageName" select ="$masterPageName"/>
				<xsl:with-param name ="lvl" select ="$lvl"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:choose >
			<xsl:when test ="document('styles.xml')//style:style[@style:name = $defaultClsName]/style:text-properties/@fo:font-size">
				<xsl:value-of select ="document('styles.xml')//style:style[@style:name = $defaultClsName]/style:text-properties/@fo:font-size" />
			</xsl:when>
			<xsl:when test ="document('styles.xml')//style:style[@style:name = concat('Standard-',$className)]/style:text-properties/@fo:font-size">
				<xsl:value-of select ="document('styles.xml')//style:style[@style:name = concat('Standard-',$className)]/style:text-properties/@fo:font-size" />
			</xsl:when>
			<xsl:when test="document('styles.xml')//style:style[@style:name = 'Standard-outline1']/style:text-properties/@style:font-size">
				<xsl:value-of select ="document('styles.xml')//style:style[@style:name = 'Standard-outline1']/style:text-properties/@fo:font-size" />
			</xsl:when>
			<xsl:when test="document('styles.xml')//style:style[@style:name = 'Standard-outline1']/style:text-properties/@style:font-size-asian">
				<xsl:value-of select ="document('styles.xml')//style:style[@style:name = 'Standard-outline1']/style:text-properties/@style:font-size-asian" />
			</xsl:when>
			<xsl:when test="document('styles.xml')//style:style[@style:name = 'Standard-outline1']/style:text-properties/@style:font-size-complex">
				<xsl:value-of select ="document('styles.xml')//style:style[@style:name = 'Standard-outline1']/style:text-properties/@style:font-size-complex" />
			</xsl:when>
			<!-- Added by lohith : sz(font size) shouldnt be zero - 16filesbug-->
			<xsl:otherwise>
				<xsl:choose>
					<xsl:when test="document('styles.xml')//style:style[@style:name = 'standard']/style:text-properties/@fo:font-size">
						<xsl:value-of select ="document('styles.xml')//style:style[@style:name = 'standard']/style:text-properties/@fo:font-size" />
					</xsl:when>
					<xsl:when test="document('styles.xml')//style:style[@style:name = 'standard']/style:text-properties/@style:font-size-asian">
						<xsl:value-of select ="document('styles.xml')//style:style[@style:name = 'standard']/style:text-properties/@style:font-size-asian" />
					</xsl:when>
					<xsl:when test="document('styles.xml')//style:style[@style:name = 'standard']/style:text-properties/@style:font-size-complex">
						<xsl:value-of select ="document('styles.xml')//style:style[@style:name = 'standard']/style:text-properties/@style:font-size-complex" />
					</xsl:when>
				</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name ="getUnderlineFromStyles" >
		<xsl:param name ="className"/>
		<xsl:message terminate="no">progress:text:p</xsl:message>
		<xsl:variable name ="defaultClsName">
			<xsl:call-template name ="getClassName">
				<xsl:with-param name ="clsName" select="$className"/>
			</xsl:call-template>
		</xsl:variable>
		<xsl:for-each select ="document('styles.xml')//style:style[@style:name = $defaultClsName ]">
			<xsl:message terminate="no">progress:text:p</xsl:message>
			<xsl:choose >
				<xsl:when test="style:text-properties/@style:text-underline-style  = 'solid' and
								style:text-properties/@style:text-underline-type[contains(.,'double')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dbl'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style  = 'solid' and
							style:text-properties/@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'sng'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style  = 'solid' and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'heavy'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- Dotted lean and dotted bold under line -->
				<xsl:when test="style:text-properties/@style:text-underline-style = 'dotted' and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dotted'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style = 'dotted' and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dottedHeavy'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- Dash lean and dash bold underline -->
				<xsl:when test="style:text-properties/@style:text-underline-style = 'dash' and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dash'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style = 'dash' and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dashHeavy'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- Dash long and dash long bold -->
				<xsl:when test="style:text-properties/@style:text-underline-style = 'long-dash' and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dashLong'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style = 'long-dash' and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dashLongHeavy'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- dot Dash and dot dash bold -->
				<xsl:when test="style:text-properties/@style:text-underline-style= 'dot-dash' and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dotDash'"/>
						<!-- Modified by lohith for fix 1739785 - dotDashLong to dotDash-->
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style= 'dot-dash' and
								style:text-properties/@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dotDashHeavy'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- dot-dot-dash-->
				<xsl:when test="style:text-properties/@style:text-underline-style= 'dot-dot-dash' and
								style:text-properties/@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dotDotDash'"/>
					</xsl:attribute >
				</xsl:when>
				<xsl:when test="style:text-properties/@style:text-underline-style= 'dot-dot-dash' and
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

			</xsl:choose >
			<!-- Stroke decoration code -->
			<xsl:choose >
				<xsl:when  test="style:text-properties/@style:text-line-through-type = 'solid'">
					<xsl:attribute name ="strike">
						<xsl:value-of select ="'sngStrike'"/>
					</xsl:attribute >
				</xsl:when >
				<xsl:when test="style:text-properties/@style:text-line-through-type[contains(.,'double')]">
					<xsl:attribute name ="strike">
						<xsl:value-of select ="'dblStrike'"/>
					</xsl:attribute >
				</xsl:when >
				<!-- style:text-line-through-style-->
				<xsl:when test="style:text-properties/@style:text-line-through-style = 'solid'">
					<xsl:attribute name ="strike">
						<xsl:value-of select ="'sngStrike'"/>
					</xsl:attribute >
				</xsl:when>
			</xsl:choose>
		</xsl:for-each>

	</xsl:template>
	<xsl:template name ="insertTab">
		<xsl:for-each select ="node()">
			<xsl:choose >
				<xsl:when test ="name()=''">
					<xsl:value-of select ="."/>
				</xsl:when>
				<xsl:when test ="name()='text:tab'">
					<xsl:value-of select ="'&#09;'"/>
				</xsl:when >
				<xsl:when test ="name()='text:s'">
					<xsl:call-template name ="insertSpace">
						<xsl:with-param name ="spaceVal" select ="@text:c"/>
					</xsl:call-template>
				</xsl:when >
				<xsl:when test =".='' and child::node()">
					<xsl:value-of select ="' '"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select ="."/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:for-each>
		<xsl:if test ="not(node())">
			<xsl:value-of select ="."/>
		</xsl:if>
	</xsl:template>
	<xsl:template name ="insertSpace">
		<xsl:param name ="spaceVal"/>
		<xsl:choose>
			<xsl:when test ="$spaceVal=1">
				<xsl:value-of  select ="'&#32;'"/>
			</xsl:when>
			<xsl:when test ="$spaceVal=2">
				<xsl:value-of  select ="'&#32;&#32;'"/>
			</xsl:when>
			<xsl:when test ="$spaceVal=3">
				<xsl:value-of  select ="'&#32;&#32;&#32;'"/>
			</xsl:when>
			<xsl:when test ="$spaceVal=4">
				<xsl:value-of  select ="'&#32;&#32;&#32;&#32;'"/>
			</xsl:when>
			<xsl:when test ="$spaceVal=5">
				<xsl:value-of  select ="'&#32;&#32;&#32;&#32;&#32;'"/>
			</xsl:when>
			<xsl:when test ="$spaceVal=6">
				<xsl:value-of  select ="'&#32;&#32;&#32;&#32;&#32;&#32;'"/>
			</xsl:when>
			<xsl:when test ="$spaceVal=7">
				<xsl:value-of  select ="'&#32;&#32;&#32;&#32;&#32;&#32;&#32;'"/>
			</xsl:when>
			<xsl:when test ="$spaceVal=8">
				<xsl:value-of  select ="'&#32;&#32;&#32;&#32;&#32;&#32;&#32;&#32;'"/>
			</xsl:when>
			<xsl:when test ="$spaceVal=9">
				<xsl:value-of  select ="'&#32;&#32;&#32;&#32;&#32;&#32;&#32;&#32;&#32;'"/>
			</xsl:when>
			<xsl:when test ="$spaceVal=10">
				<xsl:value-of  select ="'&#32;&#32;&#32;&#32;&#32;&#32;&#32;&#32;&#32;&#32;'"/>
			</xsl:when>
			<xsl:when test ="$spaceVal &gt; 10">
				<xsl:call-template name ="insertSpace" >
					<xsl:with-param name ="spaceVal" select ="$spaceVal -10 "/>
				</xsl:call-template>
				<xsl:value-of  select ="'&#32;&#32;&#32;&#32;&#32;&#32;&#32;&#32;&#32;&#32;'"/>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<xsl:template name ="paragraphTabstops">
		<a:tabLst>
			<xsl:for-each select ="style:paragraph-properties/style:tab-stops/style:tab-stop">
				<a:tab >
					<xsl:attribute name ="pos">
						<xsl:value-of select ="round(substring-before(@style:position,'cm') * 360000)"/>
					</xsl:attribute>
					<xsl:attribute name ="algn">
						<xsl:choose >
							<xsl:when test ="@style:type ='center'">
								<xsl:value-of select ="'ctr'"/>
							</xsl:when>
							<xsl:when test ="@style:type ='left'">
								<xsl:value-of select ="'l'"/>
							</xsl:when>
							<xsl:when test ="@style:type ='right'">
								<xsl:value-of select ="'r'"/>
							</xsl:when>
							<xsl:when test ="@style:type ='char'">
								<xsl:value-of select ="'dec'"/>
							</xsl:when>
							<xsl:otherwise >
								<xsl:value-of select ="'l'"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
				</a:tab >
			</xsl:for-each>
		</a:tabLst >
	</xsl:template>
	
	<!-- Template added by lohith - to get the page id -->
	<xsl:template name="getThePageId">
		<xsl:param name="PageName"/>
		<xsl:variable name="customPage">
			<xsl:for-each select="document('content.xml')/office:document-content/office:body/office:presentation/draw:page">
				<xsl:if test="@draw:name = $PageName">
					<xsl:value-of select="position()" />
				</xsl:if>
			</xsl:for-each>
		</xsl:variable>
		<xsl:if test="$customPage > 0 ">
			<xsl:value-of select="$customPage" />
		</xsl:if>
		<!-- Added for expectional case where #page is used to link the slides - Eg: Uni_animation.odp -->
		<xsl:if test="not($customPage > 0)">
			<xsl:if test="format-number(substring-after($PageName,'page'),'#') > 0 ">
				<xsl:value-of select="substring-after($PageName,'page')" />
			</xsl:if>
		</xsl:if>
	</xsl:template>
	<xsl:template name ="MarginTemplateForSlide">
		<xsl:param name ="level"/>
		<xsl:param name="listId"/>
		<xsl:param name="fileName"/>
		<xsl:for-each select ="document($fileName)//text:list-style[@style:name=$listId]">
			<xsl:choose >
				<xsl:when test ="./child::node()[$level]/style:list-level-properties/@text:space-before">
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'cm'"/>
						<xsl:with-param name ="length" select ="./child::node()[$level]/style:list-level-properties/@text:space-before"/>
					</xsl:call-template>
				</xsl:when>
				<xsl:otherwise >
          <!--edited by vipul Margin Left cant be 0cm-->
					<xsl:value-of select ="'0'"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:for-each >
	</xsl:template>
	<xsl:template name="IndentTemplateForSlide">
		<xsl:param name="level"/>
		<xsl:param name ="listId" />
		<xsl:param name ="fileName"/>
		<xsl:for-each select ="document($fileName)//text:list-style[@style:name=$listId]">
			<xsl:choose >
				<xsl:when test ="./child::node()[$level]/style:list-level-properties/@text:space-before">
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'cm'"/>
						<xsl:with-param name ="length" select ="./child::node()[$level]/style:list-level-properties/@text:min-label-width"/>
					</xsl:call-template>
				</xsl:when>
				<xsl:otherwise >
          <!--edited by vipul indent cant be 0cm-->
					<xsl:value-of select ="'0'"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:for-each >
	</xsl:template >
  <!-- Font styles for bulleted text, added by vijayeta-->
  <xsl:template name ="getTextNodeForFontStyle">
    <xsl:param name ="prClassName"/>
    <xsl:param name ="lvl" />
    <xsl:param name ="HyperlinksForBullets" />
    <xsl:param name ="masterPageName"/>
    <xsl:param name ="fileName"/>
	  <xsl:choose>
		  <!-- level 0 or no level in pptx,level 1 in odp-->
		  <xsl:when test ="./text:p">
			  <xsl:for-each select ="./text:p">
				  <xsl:call-template name ="addFontStyleTextLevels">
					  <xsl:with-param name ="prClassName" select ="$prClassName"/>
					  <xsl:with-param name ="lvl" select ="$lvl"/>
					  <xsl:with-param name ="HyperlinksForBullets" select ="$HyperlinksForBullets"/>
					  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
					  <xsl:with-param name ="fileName" select ="$fileName"/>
				  </xsl:call-template>
			  </xsl:for-each>
		  </xsl:when>
		  <!-- level 1 in pptx,level 2 in odp-->
		  <xsl:when test ="./text:list/text:list-item/text:p">
			  <xsl:for-each select ="./text:list/text:list-item/text:p">
				  <xsl:call-template name ="addFontStyleTextLevels">
					  <xsl:with-param name ="prClassName" select ="$prClassName"/>
					  <xsl:with-param name ="lvl" select ="$lvl"/>
					  <xsl:with-param name ="HyperlinksForBullets" select ="$HyperlinksForBullets"/>
					  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
					  <xsl:with-param name ="fileName" select ="$fileName"/>
				  </xsl:call-template>
			  </xsl:for-each>
		  </xsl:when>
		  <!-- level 2 in pptx,level 3 in odp-->
		  <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:p">
			  <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:p">
				  <xsl:call-template name ="addFontStyleTextLevels">
					  <xsl:with-param name ="prClassName" select ="$prClassName"/>
					  <xsl:with-param name ="lvl" select ="$lvl"/>
					  <xsl:with-param name ="HyperlinksForBullets" select ="$HyperlinksForBullets"/>
					  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
					  <xsl:with-param name ="fileName" select ="$fileName"/>
				  </xsl:call-template>
			  </xsl:for-each>
		  </xsl:when>
		  <!-- level 3 in pptx,level 4 in odp-->
		  <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
			  <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
				  <xsl:call-template name ="addFontStyleTextLevels">
					  <xsl:with-param name ="prClassName" select ="$prClassName"/>
					  <xsl:with-param name ="lvl" select ="$lvl"/>
					  <xsl:with-param name ="HyperlinksForBullets" select ="$HyperlinksForBullets"/>
					  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
					  <xsl:with-param name ="fileName" select ="$fileName"/>
				  </xsl:call-template>
			  </xsl:for-each>
		  </xsl:when>
		  <!-- level 4 in pptx,level 5 in odp-->
		  <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
			  <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
				  <xsl:call-template name ="addFontStyleTextLevels">
					  <xsl:with-param name ="prClassName" select ="$prClassName"/>
					  <xsl:with-param name ="lvl" select ="$lvl"/>
					  <xsl:with-param name ="HyperlinksForBullets" select ="$HyperlinksForBullets"/>
					  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
					  <xsl:with-param name ="fileName" select ="$fileName"/>
				  </xsl:call-template>
			  </xsl:for-each>
		  </xsl:when>
		  <!-- level 5 in pptx,level 6 in odp-->
		  <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
			  <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
				  <xsl:call-template name ="addFontStyleTextLevels">
					  <xsl:with-param name ="prClassName" select ="$prClassName"/>
					  <xsl:with-param name ="lvl" select ="$lvl"/>
					  <xsl:with-param name ="HyperlinksForBullets" select ="$HyperlinksForBullets"/>
					  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
					  <xsl:with-param name ="fileName" select ="$fileName"/>
				  </xsl:call-template>
			  </xsl:for-each>
		  </xsl:when>
		  <!-- level 6 in pptx,level 7 in odp-->
		  <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
			  <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
				  <xsl:call-template name ="addFontStyleTextLevels">
					  <xsl:with-param name ="prClassName" select ="$prClassName"/>
					  <xsl:with-param name ="lvl" select ="$lvl"/>
					  <xsl:with-param name ="HyperlinksForBullets" select ="$HyperlinksForBullets"/>
					  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
					  <xsl:with-param name ="fileName" select ="$fileName"/>
				  </xsl:call-template>
			  </xsl:for-each>
		  </xsl:when>
		  <!-- level 7 in pptx,level 8 in odp-->
		  <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
			  <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
				  <xsl:call-template name ="addFontStyleTextLevels">
					  <xsl:with-param name ="prClassName" select ="$prClassName"/>
					  <xsl:with-param name ="lvl" select ="$lvl"/>
					  <xsl:with-param name ="HyperlinksForBullets" select ="$HyperlinksForBullets"/>
					  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
					  <xsl:with-param name ="fileName" select ="$fileName"/>
				  </xsl:call-template>
			  </xsl:for-each>
		  </xsl:when>
		  <!-- level 8 in pptx,level 9 in odp-->
		  <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
			  <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
				  <xsl:call-template name ="addFontStyleTextLevels">
					  <xsl:with-param name ="prClassName" select ="$prClassName"/>
					  <xsl:with-param name ="lvl" select ="$lvl"/>
					  <xsl:with-param name ="HyperlinksForBullets" select ="$HyperlinksForBullets"/>
					  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
					  <xsl:with-param name ="fileName" select ="$fileName"/>
				  </xsl:call-template>
			  </xsl:for-each>
		  </xsl:when>
		  <!-- level 9 in pptx,level 10 in odp in case ther's 10th level in input odp, for example, DRM.odp
			Since pptx does not support 10th level, any text that is of 10th level in odp, 
			is set back to 9th level in pptx.
			date:14th Aug, '07-->
		  <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
			  <xsl:for-each select ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
				  <xsl:call-template name ="addFontStyleTextLevels">
					  <xsl:with-param name ="prClassName" select ="$prClassName"/>
					  <xsl:with-param name ="lvl" select ="$lvl"/>
					  <xsl:with-param name ="HyperlinksForBullets" select ="$HyperlinksForBullets"/>
					  <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
					  <xsl:with-param name ="fileName" select ="$fileName"/>
				  </xsl:call-template>
			  </xsl:for-each>
		  </xsl:when>
	  </xsl:choose>
  </xsl:template>
  <!--End,to get paragraph style name for each of the levels in multilevelled list-->
  <xsl:template name ="addFontStyleTextLevels">
    <xsl:param name ="prClassName"/>
    <xsl:param name ="lvl"/>
    <xsl:param name ="HyperlinksForBullets" />
    <xsl:param name ="masterPageName" />
    <xsl:param name ="fileName"/>
    <xsl:for-each select ="child::node()[position()]">
      <xsl:if test ="name()='text:span'">
        <xsl:if test ="not(./text:line-break)">
			<xsl:if test ="child::node()">
				<a:r>
					<a:rPr lang="en-US" smtClean="0">
						<!--Font Size -->
						<xsl:variable name ="textId">
							<xsl:value-of select ="@text:style-name"/>
						</xsl:variable>
						<xsl:if test ="not($textId ='')">
							<xsl:call-template name ="fontStyles">
								<xsl:with-param name ="Tid" select ="$textId" />
								<xsl:with-param name ="prClassName" select ="$prClassName"/>
								<xsl:with-param name ="lvl" select ="$lvl"/>
								<xsl:with-param name ="masterPageName" select ="$masterPageName"/>
								<xsl:with-param name ="fileName" select ="$fileName"/>
							</xsl:call-template>
						</xsl:if>
						<xsl:if test="./text:a">
							<xsl:copy-of select="$HyperlinksForBullets"/>
						</xsl:if>
					</a:rPr >
					<a:t>
						<xsl:call-template name ="insertTab" />
					</a:t>
				</a:r>
			</xsl:if>
			<xsl:if test ="not(child::node())">
				<a:endParaRPr lang="en-US" dirty="0" smtClean="0" >
					<xsl:if test ="not(@text:style-name ='')">
						<xsl:call-template name ="getFontSizeFamilyFromContentEndPara">
							<xsl:with-param name ="Tid" select ="@text:style-name"/>							
						</xsl:call-template>
					</xsl:if >
					<xsl:if test ="@text:style-name =''">
						<a:endParaRPr lang="en-US" dirty="0" smtClean="0"/>
					</xsl:if>
				</a:endParaRPr>
			</xsl:if>
		</xsl:if>
        <xsl:if test ="./text:line-break">
          <xsl:call-template name ="processBR">
            <xsl:with-param name ="T" select ="@text:style-name" />
            <xsl:with-param name ="prClassName" select ="$prClassName"/>
          </xsl:call-template>
        </xsl:if>
      </xsl:if>
      <xsl:if test ="name()='text:line-break'">
        <xsl:call-template name ="processBR">
          <xsl:with-param name ="T" select ="@text:style-name" />
          <xsl:with-param name ="prClassName" select ="$prClassName"/>
        </xsl:call-template>
      </xsl:if>
      <xsl:if test ="not(name()='text:span' or name()='text:line-break')">
        <a:r>
          <a:rPr lang="en-US" smtClean="0">
            <!--Font Size -->
            <xsl:variable name ="textId">
              <xsl:value-of select ="./parent::node()/@text:style-name"/>
            </xsl:variable>
            <xsl:if test ="not($textId ='')">
              <xsl:call-template name ="fontStyles">
                <xsl:with-param name ="Tid" select ="$textId" />
                <xsl:with-param name ="prClassName" select ="$prClassName"/>
                <xsl:with-param name ="lvl" select ="$lvl"/>
                <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                <xsl:with-param name ="fileName" select ="$fileName"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test="./text:a">
              <xsl:copy-of select="$HyperlinksForBullets"/>
            </xsl:if>
          </a:rPr >
          <a:t>
            <xsl:call-template name ="insertTab" />
          </a:t>
        </a:r>
      </xsl:if >
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="processBR">
    <xsl:param name ="T" />
    <a:br>
      <a:rPr lang="en-US" smtClean="0">
        <!--Font Size -->
        <xsl:if test ="not($T ='')">
          <xsl:call-template name ="fontStyles">
            <xsl:with-param name ="Tid" select ="$T" />
          </xsl:call-template>
        </xsl:if>
      </a:rPr >

    </a:br>
  </xsl:template>
	<xsl:template name ="getFontSizeFamilyFromContentEndPara">
		<xsl:param name ="Tid"/>
		<xsl:for-each  select ="document('content.xml')//office:automatic-styles/style:style[@style:name =$Tid ]">
			<xsl:if test="style:text-properties/@fo:font-size and substring-before(style:text-properties/@fo:font-size,'pt')&gt; 0 ">
				<xsl:attribute name ="sz">
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'pt'"/>
						<xsl:with-param name ="length" select ="style:text-properties/@fo:font-size"/>
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>
			<xsl:if test ="style:text-properties/@fo:font-family">
				<a:latin charset="0" >
					<xsl:attribute name ="typeface" >
						<!-- fo:font-family-->
						<xsl:value-of select ="translate(style:text-properties/@fo:font-family, &quot;'&quot;,'')" />
					</xsl:attribute>
				</a:latin >
			</xsl:if>
		</xsl:for-each>
		</xsl:template>
</xsl:stylesheet>

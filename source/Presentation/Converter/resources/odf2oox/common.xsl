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
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
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
  exclude-result-prefixes="odf style text number draw page r presentation fo script xlink svg">
	
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
    <xsl:param name ="flagPresentationClass"/>
		<xsl:param name ="lvl"/>
    <xsl:param name="parentStyleName"/>
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
        <xsl:call-template name="tmpSuperSubScriptForward"/>
			</xsl:if>

			<xsl:if test ="not(style:text-properties/@fo:font-size) and $flagPresentationClass='No'">
        <xsl:call-template name="tmpgetDefualtTextProp">
          <xsl:with-param name="parentStyleName" select="$parentStyleName"/>
          <xsl:with-param name="attrName" select="'Fontsize'"/>
					</xsl:call-template >
			      </xsl:if>
			<!--Font bold attribute -->
			<xsl:if test="style:text-properties/@fo:font-weight[contains(.,'bold')]">
				<xsl:attribute name ="b">
					<xsl:value-of select ="'1'"/>
				</xsl:attribute >
			</xsl:if >
      <xsl:if test ="not(style:text-properties/@fo:font-weight[contains(.,'bold')]) and $flagPresentationClass='No'">
        <xsl:call-template name="tmpgetDefualtTextProp">
          <xsl:with-param name="parentStyleName" select="$parentStyleName"/>
          <xsl:with-param name="attrName" select="'Bold'"/>
        </xsl:call-template>
      </xsl:if>
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
      <xsl:if test ="not(style:text-properties/@style:letter-kerning) and  $flagPresentationClass='No'">
        <xsl:call-template name="tmpgetDefualtTextProp">
          <xsl:with-param name="parentStyleName" select="$parentStyleName"/>
          <xsl:with-param name="attrName" select="'kerning'"/>
        </xsl:call-template>
      </xsl:if>
			
			<!-- End -->
			<!-- Font Inclined-->
			<xsl:if test="style:text-properties/@fo:font-style[contains(.,'italic')]">
				<xsl:attribute name ="i">
					<xsl:value-of select ="'1'"/>
				</xsl:attribute >
			</xsl:if >
      <xsl:if test ="not(style:text-properties/@fo:font-style[contains(.,'italic')]) and $flagPresentationClass='No'">
        <xsl:call-template name="tmpgetDefualtTextProp">
          <xsl:with-param name="parentStyleName" select="$parentStyleName"/>
          <xsl:with-param name="attrName" select="'italic'"/>
        </xsl:call-template>
      </xsl:if>
			
			<!-- Font underline-->
      <xsl:for-each select="style:text-properties">
        <xsl:call-template name="tmpUnderLineStyle"/>
      </xsl:for-each>

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
        <xsl:variable name ="spc">
					<xsl:if test ="substring-before(style:text-properties/@fo:letter-spacing,'cm')&lt; 0 ">
						<!--<xsl:value-of select ="format-number(substring-before(style:text-properties/@fo:letter-spacing,'cm') * 3.5 *1000 ,'#')"/>-->
						<xsl:value-of select ="format-number(substring-before(style:text-properties/@fo:letter-spacing,'cm') * 7200 div 2.54 ,'#')"/>
					</xsl:if >
					<xsl:if test ="substring-before(style:text-properties/@fo:letter-spacing,'cm')
								&gt; 0 or substring-before(style:text-properties/@fo:letter-spacing,'cm') = 0 ">
						<!--<xsl:value-of select ="format-number((substring-before(style:text-properties/@fo:letter-spacing,'cm') div .035) *100 ,'#')"/>-->
						<xsl:value-of select ="format-number((substring-before(style:text-properties/@fo:letter-spacing,'cm') * 72 div 2.54) *100 ,'#')"/>
					</xsl:if>
				</xsl:variable>
        <xsl:if test ="$spc!=''">
          <xsl:attribute name ="spc">
            <xsl:value-of select ="$spc"/>
          </xsl:attribute>
        </xsl:if>
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
      <xsl:if test ="not(style:text-properties/@fo:color) and $flagPresentationClass='No'">
        <xsl:call-template name="tmpgetDefualtTextProp">
          <xsl:with-param name="parentStyleName" select="$parentStyleName"/>
          <xsl:with-param name="attrName" select="'Fontcolor'"/>
        </xsl:call-template>
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
      <xsl:if test ="not(style:text-properties/@fo:text-shadow) and $flagPresentationClass='No'">
        <xsl:call-template name="tmpgetDefualtTextProp">
          <xsl:with-param name="parentStyleName" select="$parentStyleName"/>
          <xsl:with-param name="attrName" select="'Textshadow'"/>
        </xsl:call-template>
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
      <xsl:if test ="not(style:text-properties/@fo:font-family) and $flagPresentationClass='No'">
        <xsl:call-template name="tmpgetDefualtTextProp">
          <xsl:with-param name="parentStyleName" select="$parentStyleName"/>
          <xsl:with-param name="attrName" select="'Fontname'"/>
							</xsl:call-template >
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
    <xsl:param name="pos" />
    <xsl:param name ="shapeCount" />
    <xsl:param name ="FrameCount"/>
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
        <!-- Condition which checks if text indent is greater than 0 is removed, as the input has an indent of value 0
             bug number 1779336,by vijayeta, date:23rd aug '07-->
        <xsl:if test ="style:paragraph-properties/@fo:text-indent">
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
        <!-- End,bug number 1779336 vijayeta, date:23rd aug '07-->
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
        <!-- Condition which checks if text indent is greater than 0 is removed, as the input has a marginn left of value 0
             bug number 1779336 by vijayeta, date:23rd aug '07-->
				<xsl:if test ="style:paragraph-properties/@fo:margin-left">
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
        <xsl:if test ="style:paragraph-properties/@fo:margin-right">
          <!-- warn if indent after text-->
          <xsl:message terminate="no">translation.odf2oox.paragraphIndentTypeAfterText</xsl:message>
        </xsl:if>
        <!-- End,bug number 1779336 vijayeta, date:23rd aug '07-->
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
        <xsl:if test ="substring-before(style:paragraph-properties/@style:line-spacing,'cm') > 0">
					<a:lnSpc>
						<a:spcPts>
              <xsl:attribute name ="val">
                <xsl:call-template name ="convertToPointsLineSpacing">
                  <xsl:with-param name="length"  select ="style:paragraph-properties/@style:line-spacing"/>
                  <xsl:with-param name ="unit" select ="'cm'"/>
                </xsl:call-template>
                <!--<xsl:value-of select ="round(substring-before(style:paragraph-properties/@style:line-spacing,'cm')* 2835) "/>-->
              </xsl:attribute>
						</a:spcPts>
					</a:lnSpc>
				</xsl:if>
        <xsl:if test ="substring-before(style:paragraph-properties/@style:line-height-at-least,'cm') > 0 ">
					<a:lnSpc>
						<a:spcPts>
              <xsl:attribute name ="val">
                <xsl:call-template name ="convertToPointsLineSpacing">
                  <xsl:with-param name="length"  select ="style:paragraph-properties/@style:line-height-at-least"/>
                  <xsl:with-param name ="unit" select ="'cm'"/>
                </xsl:call-template>
                <!--<xsl:value-of select ="round(substring-before(style:paragraph-properties/@style:line-height-at-least,'cm')* 2835) "/>-->
              </xsl:attribute>
						</a:spcPts>
					</a:lnSpc>
				</xsl:if>
				<!--End of Code inserted by Vijayeta For Line Spacing -->
				<!-- Code Added by Vijayeta,for Paragraph Spacing, Before and After
             Multiply the value in cm with 2835
			 date: on 01-06-07-->
				<xsl:if test ="style:paragraph-properties/@fo:margin-top">
					<a:spcBef>
						<a:spcPts>
              <xsl:attribute name ="val">
                <!--fo:margin-top-->
                <xsl:call-template name ="convertToPointsLineSpacing">
                  <xsl:with-param name="length"  select ="style:paragraph-properties/@fo:margin-top"/>
                  <xsl:with-param name ="unit" select ="'cm'"/>
                </xsl:call-template>
                <!--<xsl:value-of select ="round(substring-before(style:paragraph-properties/@fo:margin-top,'cm')* 2835) "/>-->
              </xsl:attribute>
						</a:spcPts>
					</a:spcBef >
				</xsl:if>
				<xsl:if test ="style:paragraph-properties/@fo:margin-bottom">
					<a:spcAft>
						<a:spcPts>
              <xsl:attribute name ="val">
                <!--fo:margin-bottom-->
                <xsl:call-template name ="convertToPointsLineSpacing">
                  <xsl:with-param name="length"  select ="style:paragraph-properties/@fo:margin-bottom"/>
                  <xsl:with-param name ="unit" select ="'cm'"/>
                </xsl:call-template>
                <!--<xsl:value-of select ="round(substring-before(style:paragraph-properties/@fo:margin-bottom,'cm')* 2835) "/>-->
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
              <xsl:with-param name ="pos" select ="$pos"/>
              <xsl:with-param name ="shapeCount" select ="$shapeCount" />
              <xsl:with-param name ="FrameCount" select ="$FrameCount"/>
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
      <xsl:choose>
        <xsl:when test ="style:graphic-properties/@draw:fill='solid'">
          <a:solidFill>
            <a:srgbClr  >
              <xsl:attribute name ="val">
                <xsl:value-of select ="translate(substring-after(style:graphic-properties/@draw:fill-color,'#'),$lcletters,$ucletters)"/>
              </xsl:attribute>
              <xsl:if test ="style:graphic-properties/@draw:opacity">
                <xsl:variable name="tranparency" select="substring-before(style:graphic-properties/@draw:opacity,'%')"/>
                <xsl:choose>
                  <xsl:when test="$tranparency !=''">
                    <a:alpha>
                      <xsl:attribute name="val">
                        <xsl:value-of select="format-number($tranparency * 1000,'#')" />
                      </xsl:attribute>
                    </a:alpha>
                  </xsl:when>
                </xsl:choose>               
              </xsl:if>
            </a:srgbClr >
          </a:solidFill>
        </xsl:when>
        <xsl:when test ="style:graphic-properties/@draw:fill='none'">
          <a:noFill/>
          </xsl:when>
        <xsl:when test ="style:graphic-properties/@draw:fill='gradient'">
          <xsl:call-template name="tmpGradientFill">
            <xsl:with-param name="gradStyleName" select="style:graphic-properties/@draw:fill-gradient-name"/>
            <xsl:with-param  name="opacity" select="substring-before(style:graphic-properties/@draw:opacity,'%')"/>
          </xsl:call-template>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="tmpGradientFill">
    <xsl:param name="gradStyleName"/>
    <xsl:param name="opacity"/>
    <a:gradFill flip="none" rotWithShape="1">

                <xsl:for-each select="document('styles.xml')//draw:gradient[@draw:name= $gradStyleName]">
        <a:gsLst>
          <a:gs pos="0">
            <a:srgbClr>
                  <xsl:attribute name="val">
                    <xsl:if test="@draw:start-color">
                      <xsl:value-of select="substring-after(@draw:start-color,'#')" />
                    </xsl:if>
                    <xsl:if test="not(@draw:start-color)">
                      <xsl:value-of select="'ffffff'" />
                    </xsl:if>
                  </xsl:attribute>
              <xsl:if test ="$opacity !=''">
                    <a:alpha>
                      <xsl:attribute name="val">
                    <xsl:value-of select="format-number($opacity * 1000,'#')" />
                      </xsl:attribute>
                    </a:alpha>
              </xsl:if>
            </a:srgbClr >
          </a:gs>
          <a:gs pos="100000">
            <a:srgbClr>
              <xsl:attribute name="val">
                <xsl:if test="@draw:end-color">
                  <xsl:value-of select="substring-after(@draw:end-color,'#')" />
                </xsl:if>
                <xsl:if test="not(@draw:end-color)">
                  <xsl:value-of select="'ffffff'" />
                </xsl:if>
              </xsl:attribute>
              <xsl:if test ="$opacity !=''">
                <a:alpha>
                  <xsl:attribute name="val">
                    <xsl:value-of select="format-number($opacity * 1000,'#')" />
                  </xsl:attribute>
                </a:alpha>
              </xsl:if>
            </a:srgbClr>
          </a:gs>
        </a:gsLst>
      </xsl:for-each>
      <xsl:choose>
        <xsl:when test="@draw:style='radial'">
          <a:path path="circle">
            <a:fillToRect/>
          </a:path>
        </xsl:when>
        <xsl:when test="@draw:style='linear'">
          <a:lin ang="0" scaled="1"/>
          <a:tileRect/>
        </xsl:when>
        <xsl:when test="@draw:style='rectangular' or @draw:style='square'">
          <a:path path="rect">
            <a:fillToRect/>
            <a:tileRect/>
          </a:path>
        </xsl:when>
        <xsl:otherwise>
          <a:lin ang="0" scaled="1"/>
          <a:tileRect/>
        </xsl:otherwise>
        </xsl:choose>
     

    </a:gradFill>
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
  <xsl:template name ="getDefaultFontColor">
    <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
    <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
    <!--<xsl:for-each select="document('styles.xml')//style:style[@style:name = 'standard']//style:text-properties">
      <xsl:if test="position()=1">-->
     
      <xsl:choose >
        <xsl:when test ="document('styles.xml')//style:style[@style:name = 'standard']//style:text-properties/@fo:color">
          <a:solidFill>
            <a:srgbClr  >
          <xsl:attribute name ="val">
          <xsl:value-of select ="translate(substring-after(document('styles.xml')//style:style[@style:name = 'standard']//style:text-properties/@fo:color,'#'),$lcletters,$ucletters)"/>
          </xsl:attribute>
            </a:srgbClr >
          </a:solidFill>
        </xsl:when>
        <xsl:otherwise>
          <xsl:if test="document('styles.xml')//style:style[@style:name = 'standard']//style:text-properties/@style:use-window-font-color='true'">
            <a:solidFill>
                <a:sysClr val="windowText"/>
            </a:solidFill>
          </xsl:if>
          
        </xsl:otherwise>
      </xsl:choose>
  
         <!--</xsl:if>
    </xsl:for-each>-->

  
  </xsl:template>
  <xsl:template name="tmpgetDefualtTextProp">
    <xsl:param name="parentStyleName"/>
    <xsl:param name="attrName"/>
    <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
    <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
    <xsl:for-each select="document('styles.xml')//style:style[@style:name = $parentStyleName]/style:text-properties">
      <xsl:variable name="prStyleName" select="./parent::node()/@style:parent-style-name"/>
      <xsl:choose>
        <xsl:when test="$attrName='Fontcolor'">
          <xsl:choose>
            <xsl:when test="@fo:color">
              <a:solidFill>
                <a:srgbClr  >
                  <xsl:attribute name ="val">
                    <xsl:value-of select ="translate(substring-after(@fo:color,'#'),$lcletters,$ucletters)"/>
                  </xsl:attribute>
                </a:srgbClr >
              </a:solidFill>
            </xsl:when>
            <xsl:when test="@style:use-window-font-color='true'">
              <a:solidFill>
                <a:sysClr val="windowText"/>
              </a:solidFill>
            </xsl:when>
            <xsl:when test="not(@fo:color) and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtTextProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'Fontcolor'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
  
        <xsl:when test="$attrName='Fontsize'">
          <xsl:choose>
            <xsl:when test="substring-before(@fo:font-size,'pt') > 0">
              <xsl:attribute name ="sz">
                <xsl:call-template name ="convertToPoints">
                  <xsl:with-param name ="unit" select ="'pt'"/>
                  <xsl:with-param name ="length" select ="@fo:font-size"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="substring-before(@style:font-size-asian,'pt') > 0">
              <xsl:attribute name ="sz">
                <xsl:call-template name ="convertToPoints">
                  <xsl:with-param name ="unit" select ="'pt'"/>
                  <xsl:with-param name ="length" select ="@style:font-size-asian"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="substring-before(@style:font-size-complex,'pt') > 0">
              <xsl:attribute name ="sz">
                <xsl:call-template name ="convertToPoints">
                  <xsl:with-param name ="unit" select ="'pt'"/>
                  <xsl:with-param name ="length" select ="@style:font-size-complex"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="not(@fo:font-size) and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtTextProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'Fontsize'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="$attrName='Fontname'">
          <xsl:choose>
            <xsl:when test="@fo:font-family">
              <a:latin charset="0" >
                <xsl:attribute name ="typeface" >
                  <xsl:value-of select ="translate(@fo:font-family, &quot;'&quot;,'')" />
                </xsl:attribute>
              </a:latin>
            </xsl:when>
            <xsl:when test="not(@fo:font-family) and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtTextProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'Fontname'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="$attrName='Bold'">
          <xsl:choose>
            <xsl:when test="@fo:font-weight[contains(.,'bold')]">
              <xsl:attribute name ="b">
                <xsl:value-of select ="'1'"/>
              </xsl:attribute >
            </xsl:when>
            <xsl:when test="not(@fo:font-weight[contains(.,'bold')]) and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtTextProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'Bold'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="$attrName='italic'">
          <xsl:choose>
            <xsl:when test="@fo:font-style[contains(.,'italic')]">
              <xsl:attribute name ="i">
                <xsl:value-of select ="'1'"/>
              </xsl:attribute >
            </xsl:when>
            <xsl:when test="not(@fo:font-style[contains(.,'italic')]) and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtTextProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'italic'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="$attrName='kerning'">
          <xsl:choose>
            <xsl:when test ="@style:letter-kerning = 'true'">
              <xsl:attribute name ="kern">
                <xsl:value-of select="1200"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test ="@style:letter-kerning = 'false'">
              <xsl:attribute name ="kern">
                <xsl:value-of select="0"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="not(@style:letter-kerning) and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtTextProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'kerning'"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="not(@style:letter-kerning)">
              <xsl:attribute name ="kern">
                <xsl:value-of select="0"/>
              </xsl:attribute>
            </xsl:when>
          </xsl:choose>
        
        </xsl:when>
        <xsl:when test="$attrName='Textshadow'">
          <xsl:choose>
            <xsl:when test ="@fo:text-shadow != 'none'">
              <a:effectLst>
                <a:outerShdw blurRad="38100" dist="38100" dir="2700000" >
                  <a:srgbClr val="000000">
                    <a:alpha val="43137" />
                  </a:srgbClr>
                </a:outerShdw>
              </a:effectLst>
            </xsl:when>
            <xsl:when test="not(@fo:text-shadow) and $prStyleName !=''">
              <xsl:call-template name="tmpgetDefualtTextProp">
                <xsl:with-param name="parentStyleName" select="$prStyleName"/>
                <xsl:with-param name="attrName" select="'Textshadow'"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
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
    <xsl:for-each select ="document('styles.xml')//style:style[@style:name = $defaultClsName]/style:text-properties">
			<xsl:message terminate="no">progress:text:p</xsl:message>
      <xsl:call-template name="tmpUnderLineStyle"/>
      <!-- Stroke decoration code -->
			<xsl:choose >
        <xsl:when  test="@style:text-line-through-type = 'solid'">
          <xsl:attribute name ="strike">
            <xsl:value-of select ="'sngStrike'"/>
          </xsl:attribute >
        </xsl:when >
        <xsl:when test="@style:text-line-through-type[contains(.,'double')]">
          <xsl:attribute name ="strike">
            <xsl:value-of select ="'dblStrike'"/>
          </xsl:attribute >
        </xsl:when >
        <!-- style:text-line-through-style-->
        <xsl:when test="@style:text-line-through-style = 'solid'">
          <xsl:attribute name ="strike">
            <xsl:value-of select ="'sngStrike'"/>
					</xsl:attribute >
				</xsl:when>
      </xsl:choose>
    </xsl:for-each>

  </xsl:template>
  <xsl:template name="tmpUnderLineStyle">
    <xsl:choose >
      <!-- Added by lohith for fix - 1744082 - Start-->
      <xsl:when test="@style:text-underline-type = 'single'">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'sng'"/>
					</xsl:attribute >
				</xsl:when>
      <!-- Fix - 1744082 - End-->
      <xsl:when test="@style:text-underline-style = 'solid' and
								@style:text-underline-type[contains(.,'double')]">
        <xsl:attribute name ="u">
          <xsl:value-of select ="'dbl'"/>
        </xsl:attribute >
      </xsl:when>
      <xsl:when test="@style:text-underline-style  = 'solid' and
								@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'heavy'"/>
					</xsl:attribute >
				</xsl:when>
      <xsl:when test="@style:text-underline-style = 'solid' and
							@style:text-underline-width[contains(.,'auto')]">
        <xsl:attribute name ="u">
          <xsl:value-of select ="'sng'"/>
        </xsl:attribute >
      </xsl:when>
				<!-- Dotted lean and dotted bold under line -->
      <xsl:when test="@style:text-underline-style = 'dotted' and
								@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dotted'"/>
					</xsl:attribute >
				</xsl:when>
      <xsl:when test="@style:text-underline-style = 'dotted' and
								@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dottedHeavy'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- Dash lean and dash bold underline -->
      <xsl:when test="@style:text-underline-style = 'dash' and
								@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dash'"/>
					</xsl:attribute >
				</xsl:when>
      <xsl:when test="@style:text-underline-style = 'dash' and
								@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dashHeavy'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- Dash long and dash long bold -->
      <xsl:when test="@style:text-underline-style = 'long-dash' and
								@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dashLong'"/>
					</xsl:attribute >
				</xsl:when>
      <xsl:when test="@style:text-underline-style = 'long-dash' and
								@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dashLongHeavy'"/>
					</xsl:attribute >
				</xsl:when>

				<!-- dot Dash and dot dash bold -->
      <xsl:when test="@style:text-underline-style = 'dot-dash' and
								@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dotDash'"/>
						<!-- Modified by lohith for fix 1739785 - dotDashLong to dotDash-->
					</xsl:attribute >
				</xsl:when>
      <xsl:when test="@style:text-underline-style = 'dot-dash' and
								@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dotDashHeavy'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- dot-dot-dash-->
      <xsl:when test="@style:text-underline-style= 'dot-dot-dash' and
								@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dotDotDash'"/>
					</xsl:attribute >
				</xsl:when>
      <xsl:when test="@style:text-underline-style= 'dot-dot-dash' and
								@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'dotDotDashHeavy'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- double Wavy -->
      <xsl:when test="@style:text-underline-style[contains(.,'wave')] and
								@style:text-underline-type[contains(.,'double')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'wavyDbl'"/>
					</xsl:attribute >
				</xsl:when>
				<!-- Wavy and Wavy Heavy-->
      <xsl:when test="@style:text-underline-style[contains(.,'wave')] and
								@style:text-underline-width[contains(.,'auto')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'wavy'"/>
					</xsl:attribute >
				</xsl:when>
      <xsl:when test="@style:text-underline-style[contains(.,'wave')] and
								@style:text-underline-width[contains(.,'bold')]">
					<xsl:attribute name ="u">
						<xsl:value-of select ="'wavyHeavy'"/>
					</xsl:attribute >
				</xsl:when>
      <xsl:otherwise >
        <!--<xsl:call-template name ="getUnderlineFromStyles">
          <xsl:with-param name ="className" select ="$prClassName"/>
        </xsl:call-template>-->
      </xsl:otherwise>
			</xsl:choose>
		
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
          <xsl:choose>
            <xsl:when test="@style:position='cm' or @style:position='NaNcm'">
              <xsl:attribute name ="pos">
              <xsl:value-of select="'0'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
					<xsl:attribute name ="pos">
						<xsl:value-of select ="round(substring-before(@style:position,'cm') * 360000)"/>
					</xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
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
    <xsl:param name ="flagPresentationClass"/>
    
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
            <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
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
            <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
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
            <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
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
            <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
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
            <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
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
            <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
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
            <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
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
            <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
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
            <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
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
            <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
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
    <xsl:param name ="flagPresentationClass"/>

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
                    <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
                  </xsl:call-template>
                </xsl:if>
              <xsl:if test="name()='text:a' or ./text:a">
                  <xsl:copy-of select="$HyperlinksForBullets"/>
                </xsl:if>
              </a:rPr >
              <a:t>
                <xsl:call-template name ="insertTab" />
              </a:t>
            </a:r>
          </xsl:if>
          <!-- Bug 1744106 fixed by vijayeta, date 16th Aug '07, add a new templaet to set font size and family in endPara-->
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
          <!--End, Bug 1744106 fixed by vijayeta, date 16th Aug '07, add a new templaet to set font size and family in endPara-->
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
                <xsl:with-param name ="flagPresentationClass" select ="$flagPresentationClass"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test="name()='text:a' or ./text:a">
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
  <!-- Bug 1744106 fixed by vijayeta, date 16th Aug '07, add a new templaet to set font size and family in endPara-->
  <xsl:template name="getFontSizeFamilyFromContentEndPara">
    <xsl:param name="Tid" />
    <xsl:for-each select="document('content.xml')//office:automatic-styles/style:style[@style:name =$Tid ]">
      <xsl:if test="style:text-properties/@fo:font-size and substring-before(style:text-properties/@fo:font-size,'pt')> 0">
        <xsl:attribute name="sz">
          <xsl:call-template name="convertToPoints">
            <xsl:with-param name="unit" select="'pt'" />
            <xsl:with-param name="length" select="style:text-properties/@fo:font-size" />
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="style:text-properties/@fo:font-family">
        <a:latin charset="0">
          <xsl:attribute name="typeface">
            <!--  fo:font-family-->
            <xsl:value-of select="translate(style:text-properties/@fo:font-family,&quot;'&quot;,'')" />
          </xsl:attribute>
        </a:latin>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="tmpSMfontStyles">
    <xsl:param name ="TextStyleID"/>
    <xsl:param name ="prClassName"/>
    <xsl:param name ="lvl"/>
    <xsl:param name ="masterPageName"/>
    <xsl:message terminate="no">progress:text:p</xsl:message>

    <xsl:for-each  select ="document('styles.xml')//style:style[@style:name =$TextStyleID ]">
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
        <xsl:call-template name="tmpSuperSubScriptForward"/>
      </xsl:if>
      <!--Font bold attribute -->
      <xsl:if test="style:text-properties/@fo:font-weight[contains(.,'bold')]">
        <xsl:attribute name ="b">
          <xsl:value-of select ="'1'"/>
        </xsl:attribute >
      </xsl:if >
      <!-- Kerning -->
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
      <!-- Font Inclined-->
      <xsl:if test="style:text-properties/@fo:font-style[contains(.,'italic')]">
        <xsl:attribute name ="i">
          <xsl:value-of select ="'1'"/>
        </xsl:attribute >
      </xsl:if >

      <!-- Font underline-->
      <xsl:call-template name="Underline"/>

      <!-- Font Strike through -->
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
      <!--Charector spacing -->
      <xsl:if test ="style:text-properties/@fo:letter-spacing [contains(.,'cm')]">
        <!-- Modfied by lohith - "spc" should be a int value, '#.##'has been replaced by '#'   -->
        <xsl:variable name ="spc">
          <xsl:if test ="substring-before(style:text-properties/@fo:letter-spacing,'cm')&lt; 0 ">
            <!--<xsl:value-of select ="format-number(substring-before(style:text-properties/@fo:letter-spacing,'cm') * 3.5 *1000 ,'#')"/>-->
            <xsl:value-of select ="format-number(substring-before(style:text-properties/@fo:letter-spacing,'cm') * 7200 div 2.54 ,'#')"/>
          </xsl:if >
          <xsl:if test ="substring-before(style:text-properties/@fo:letter-spacing,'cm')
								&gt; 0 or substring-before(style:text-properties/@fo:letter-spacing,'cm') = 0 ">
            <!--<xsl:value-of select ="format-number((substring-before(style:text-properties/@fo:letter-spacing,'cm') div .035) *100 ,'#')"/>-->
            <xsl:value-of select ="format-number((substring-before(style:text-properties/@fo:letter-spacing,'cm') * 72 div 2.54) *100 ,'#')"/>
          </xsl:if>
        </xsl:variable>
        <xsl:if test ="$spc!=''">
          <xsl:attribute name ="spc">
            <xsl:value-of select ="$spc"/>
          </xsl:attribute>
        </xsl:if>
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
  <xsl:template name="Underline">
    <!-- Font underline-->
    <xsl:param name="uStyle"/>
    <xsl:param name="uWidth"/>
    <xsl:param name="uType"/>

    <xsl:choose >
      <xsl:when test="$uStyle='solid' and
								$uType='double'">
        <xsl:value-of select ="'dbl'"/>
      </xsl:when>
      <xsl:when test="$uStyle='solid' and	$uWidth='bold'">
        <xsl:value-of select ="'heavy'"/>
      </xsl:when>
      <xsl:when test="$uStyle='solid' and $uWidth='auto'">
        <xsl:value-of select ="'sng'"/>
      </xsl:when>
      <!-- Dotted lean and dotted bold under line -->
      <xsl:when test="$uStyle='dotted' and	$uWidth='auto'">
        <xsl:value-of select ="'dotted'"/>
      </xsl:when>
      <xsl:when test="$uStyle='dotted' and	$uWidth='bold'">
        <xsl:value-of select ="'dottedHeavy'"/>
      </xsl:when>
      <!-- Dash lean and dash bold underline -->
      <xsl:when test="$uStyle='dash' and
								$uWidth='auto'">
        <xsl:value-of select ="'dash'"/>
      </xsl:when>
      <xsl:when test="$uStyle='dash' and
								$uWidth='bold'">
        <xsl:value-of select ="'dashHeavy'"/>
      </xsl:when>
      <!-- Dash long and dash long bold -->
      <xsl:when test="$uStyle='long-dash' and
								$uWidth='auto'">
        <xsl:value-of select ="'dashLong'"/>
      </xsl:when>
      <xsl:when test="$uStyle='long-dash' and
								$uWidth='bold'">
        <xsl:value-of select ="'dashLongHeavy'"/>
      </xsl:when>
      <!-- dot Dash and dot dash bold -->
      <xsl:when test="$uStyle='dot-dash' and
								$uWidth='auto'">
        <xsl:value-of select ="'dotDashLong'"/>
      </xsl:when>
      <xsl:when test="$uStyle='dot-dash' and
								$uWidth='bold'">
        <xsl:value-of select ="'dotDashHeavy'"/>
      </xsl:when>
      <!-- dot-dot-dash-->
      <xsl:when test="$uStyle='dot-dot-dash' and
								$uWidth='auto'">
        <xsl:value-of select ="'dotDotDash'"/>
      </xsl:when>
      <xsl:when test="$uStyle='dot-dot-dash' and
								$uWidth='bold'">
        <xsl:value-of select ="'dotDotDashHeavy'"/>
      </xsl:when>
      <!-- double Wavy -->
      <xsl:when test="$uStyle='wave' and
								$uType='double'">
        <xsl:value-of select ="'wavyDbl'"/>
      </xsl:when>
      <!-- Wavy and Wavy Heavy-->
      <xsl:when test="$uStyle='wave' and
								$uWidth='auto'">
        <xsl:value-of select ="'wavy'"/>
      </xsl:when>
      <xsl:when test="$uStyle='wave' and
								$uWidth='bold'">
        <xsl:value-of select ="'wavyHeavy'"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name ="convertToPointsLineSpacing">
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
      <xsl:when test="not(contains($length,'pt'))">
        <xsl:value-of select="concat(format-number($lengthVal * 2835,'#'),'')"/>
      </xsl:when>
      <xsl:when test ="contains($length,'pt')">
        <xsl:value-of select="concat(format-number($lengthVal,'#') * 100 ,'')"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template >
  <!--End of snippet for Bug 1744106, fixed by vijayeta, date 16th Aug '07, add a new template to set font size and family in endPara-->
 
  <xsl:template name="tmpgroupingCordinates">

    <p:grpSpPr>
      <a:xfrm>
        <xsl:variable name="grpWidth">
          <xsl:for-each select="node()">
            <!--<xsl:if test="name()='draw:g'">-->
            <xsl:choose>
              <!--<xsl:when test="name()='draw:g'">
                <xsl:call-template name="tmpgroupingCordinates"/>
              </xsl:when>-->
              <xsl:when test="name()='draw:frame'">
                <xsl:choose>
                  <xsl:when test="./draw:image">
                    <xsl:if test ="contains(./draw:image/@xlink:href,'.png') or contains(./draw:image/@xlink:href,'.emf') or contains(./draw:image/@xlink:href,'.wmf')
                    or contains(./draw:image/@xlink:href,'.jfif') or contains(./draw:image/@xlink:href,'.jpe') 
                    or contains(./draw:image/@xlink:href,'.bmp') or contains(./draw:image/@xlink:href,'.dib') 
                    or contains(./draw:image/@xlink:href,'.rle')
                    or contains(./draw:image/@xlink:href,'.bmz') or contains(./draw:image/@xlink:href,'.gfa') 
                    or contains(./draw:image/@xlink:href,'.emz') or contains(./draw:image/@xlink:href,'.wmz')
                    or contains(./draw:image/@xlink:href,'.pcz')
            or contains(./draw:image/@xlink:href,'.tif') or contains(./draw:image/@xlink:href,'.tiff') 
            or contains(./draw:image/@xlink:href,'.cdr') or contains(./draw:image/@xlink:href,'.cgm') or contains(./draw:image/@xlink:href,'.eps') 
            or contains(./draw:image/@xlink:href,'.pct') or contains(./draw:image/@xlink:href,'.pict') or contains(./draw:image/@xlink:href,'.wpg') 
            or contains(./draw:image/@xlink:href,'.jpeg') or contains(./draw:image/@xlink:href,'.gif') or contains(./draw:image/@xlink:href,'.png')
            or contains(./draw:image/@xlink:href,'.jpg')">
                      <xsl:if test="not(./draw:image/@xlink:href[contains(.,'../')])">
                        <xsl:choose>
                          <xsl:when test="@svg:x and @svg:width">
                            <xsl:value-of select="concat(number(substring-before(@svg:x,'cm')) + number(substring-before(@svg:width,'cm')) , ':')"/>
                          </xsl:when>
                          <xsl:when test="@draw:transform">
                            <xsl:variable name ="x">
                              <xsl:value-of select="substring-before(substring-before(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
                            </xsl:variable>
                            <xsl:value-of select="concat(number($x) + number(substring-before(@svg:width,'cm')) , ':')"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="'0:'"/>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:if>
                    </xsl:if>
                  </xsl:when>
                  <xsl:when test="@presentation:class[contains(.,'title')]
                                  or @presentation:class[contains(.,'subtitle')]
                                  or @presentation:class[contains(.,'outline')]"></xsl:when>
                  <xsl:when test ="(draw:text-box) and not(@presentation:style-name) and not(@presentation:class)">
                    <xsl:choose>
                      <xsl:when test="@svg:x and @svg:width">
                        <xsl:value-of select="concat(number(substring-before(@svg:x,'cm')) + number(substring-before(@svg:width,'cm')) , ':')"/>
                      </xsl:when>
                      <xsl:when test="@draw:transform">
                        <xsl:variable name ="x">
                          <xsl:value-of select="substring-before(substring-before(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
                        </xsl:variable>
                        <xsl:value-of select="concat(number($x) + number(substring-before(@svg:width,'cm')) , ':')"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="'0:'"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:when>
                </xsl:choose>
              </xsl:when>
              <xsl:when test="name()='draw:line' or  name()='draw:connector'">
                <xsl:choose>
                  <xsl:when test="@svg:x1 and @svg:x2">
                    <xsl:if test="svg:x2 >= svg:x1">
                    <xsl:value-of select="concat(number(substring-before(@svg:x1,'cm')) + ( number(substring-before(@svg:x2,'cm')) - number(substring-before(@svg:x1,'cm')) ), ':')"/>
                    </xsl:if>
                    <xsl:if test="svg:x1 >= svg:x2">
                      <xsl:value-of select="concat(number(substring-before(@svg:x1,'cm')) + ( number(substring-before(@svg:x1,'cm')) - number(substring-before(@svg:x2,'cm')) ), ':')"/>
                    </xsl:if>
                  </xsl:when>

                  <xsl:otherwise>
                    <xsl:value-of select="'0:'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>

              <xsl:when test="name()='draw:rect' or name()='draw:ellipse' or name()='draw:custom-shape'">
                <xsl:choose>
                  <xsl:when test="@svg:x and @svg:width">
                    <xsl:value-of select="concat(number(substring-before(@svg:x,'cm')) + number(substring-before(@svg:width,'cm')) , ':')"/>
                  </xsl:when>
                  <xsl:when test="@draw:transform">
                    <xsl:variable name ="x">
                      <xsl:value-of select="substring-before(substring-before(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
                    </xsl:variable>
                    <xsl:value-of select="concat(number($x) + number(substring-before(@svg:width,'cm')) , ':')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'0:'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>
            </xsl:choose>


          </xsl:for-each>
        </xsl:variable>
        <xsl:variable name="grpHeight">
          <xsl:for-each select="node()">
            <xsl:choose>
              <!--<xsl:when test="name()='draw:g'">
                    <xsl:call-template name="tmpgroupingCordinates"/>
                  </xsl:when>-->
              <xsl:when test="name()='draw:frame'">
                <xsl:choose>
                  <xsl:when test="./draw:image">
                    <xsl:if test ="contains(./draw:image/@xlink:href,'.png') or contains(./draw:image/@xlink:href,'.emf') or contains(./draw:image/@xlink:href,'.wmf')
                    or contains(./draw:image/@xlink:href,'.jfif') or contains(./draw:image/@xlink:href,'.jpe') 
                    or contains(./draw:image/@xlink:href,'.bmp') or contains(./draw:image/@xlink:href,'.dib') 
                    or contains(./draw:image/@xlink:href,'.rle')
                    or contains(./draw:image/@xlink:href,'.bmz') or contains(./draw:image/@xlink:href,'.gfa') 
                    or contains(./draw:image/@xlink:href,'.emz') or contains(./draw:image/@xlink:href,'.wmz')
                    or contains(./draw:image/@xlink:href,'.pcz')
            or contains(./draw:image/@xlink:href,'.tif') or contains(./draw:image/@xlink:href,'.tiff') 
            or contains(./draw:image/@xlink:href,'.cdr') or contains(./draw:image/@xlink:href,'.cgm') or contains(./draw:image/@xlink:href,'.eps') 
            or contains(./draw:image/@xlink:href,'.pct') or contains(./draw:image/@xlink:href,'.pict') or contains(./draw:image/@xlink:href,'.wpg') 
            or contains(./draw:image/@xlink:href,'.jpeg') or contains(./draw:image/@xlink:href,'.gif') or contains(./draw:image/@xlink:href,'.png')
            or contains(./draw:image/@xlink:href,'.jpg')">
                      <xsl:if test="not(./draw:image/@xlink:href[contains(.,'../')])">
                        <xsl:choose>
                          <xsl:when test="@svg:y and @svg:height">
                            <xsl:value-of select="concat(number(substring-before(@svg:y,'cm')) + number(substring-before(@svg:width,'cm')) , ':')"/>
                          </xsl:when>
                          <xsl:when test="@draw:transform">
                            <xsl:variable name ="y">
                              <xsl:value-of select="substring-before(substring-after(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
                            </xsl:variable>
                            <xsl:value-of select="concat(number($y) + number(substring-before(@svg:height,'cm')) , ':')"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="'0:'"/>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:if>
                    </xsl:if>
                  </xsl:when>
                  <xsl:when test="@presentation:class[contains(.,'title')]
                                  or @presentation:class[contains(.,'subtitle')]
                                  or @presentation:class[contains(.,'outline')]"></xsl:when>
                  <xsl:when test ="(draw:text-box) and not(@presentation:style-name) and not(@presentation:class)">
                    <xsl:choose>
                      <xsl:when test="@svg:y and @svg:height">
                        <xsl:value-of select="concat(number(substring-before(@svg:y,'cm')) + number(substring-before(@svg:height,'cm')) , ':')"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="'0:'"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:when>
                </xsl:choose>
              </xsl:when>
              <xsl:when test="name()='draw:line' or  name()='draw:connector'">
                <xsl:choose>
                  <xsl:when test="@svg:y1 and @svg:y2">
                    <xsl:if test="svg:y2 >= svg:y1">
                    <xsl:value-of select="concat(number(substring-before(@svg:y1,'cm')) + ( number(substring-before(@svg:y2,'cm')) - number(substring-before(@svg:y1,'cm')) ), ':')"/>
                    </xsl:if>
                    <xsl:if test="svg:y1 >= svg:y2">
                      <xsl:value-of select="concat(number(substring-before(@svg:y1,'cm')) + ( number(substring-before(@svg:y1,'cm')) - number(substring-before(@svg:y2,'cm')) ), ':')"/>
                    </xsl:if>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'0:'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>

              <xsl:when test="name()='draw:rect' or name()='draw:ellipse' or name()='draw:custom-shape'">
                <xsl:choose>
                  <xsl:when test="@svg:y and @svg:height">
                    <xsl:value-of select="concat(number(substring-before(@svg:y,'cm')) + number(substring-before(@svg:height,'cm')) , ':')"/>
                  </xsl:when>
                  <xsl:when test="@draw:transform">
                    <xsl:variable name ="y">
                      <xsl:value-of select="substring-before(substring-after(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
                    </xsl:variable>
                    <xsl:value-of select="concat(number($y) + number(substring-before(@svg:height,'cm')) , ':')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'0:'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>
            </xsl:choose>
          </xsl:for-each>
        </xsl:variable>
        <xsl:variable name="grpMinX">
          <xsl:for-each select="node()">
            <xsl:choose>
              <!--<xsl:when test="name()='draw:g'">
                    <xsl:call-template name="tmpgroupingCordinates"/>
                  </xsl:when>-->
              <xsl:when test="name()='draw:frame'">
                <xsl:choose>
                  <xsl:when test="./draw:image">
                    <xsl:if test ="contains(./draw:image/@xlink:href,'.png') or contains(./draw:image/@xlink:href,'.emf') or contains(./draw:image/@xlink:href,'.wmf')
                    or contains(./draw:image/@xlink:href,'.jfif') or contains(./draw:image/@xlink:href,'.jpe') 
                    or contains(./draw:image/@xlink:href,'.bmp') or contains(./draw:image/@xlink:href,'.dib') 
                    or contains(./draw:image/@xlink:href,'.rle')
                    or contains(./draw:image/@xlink:href,'.bmz') or contains(./draw:image/@xlink:href,'.gfa') 
                    or contains(./draw:image/@xlink:href,'.emz') or contains(./draw:image/@xlink:href,'.wmz')
                    or contains(./draw:image/@xlink:href,'.pcz')
            or contains(./draw:image/@xlink:href,'.tif') or contains(./draw:image/@xlink:href,'.tiff') 
            or contains(./draw:image/@xlink:href,'.cdr') or contains(./draw:image/@xlink:href,'.cgm') or contains(./draw:image/@xlink:href,'.eps') 
            or contains(./draw:image/@xlink:href,'.pct') or contains(./draw:image/@xlink:href,'.pict') or contains(./draw:image/@xlink:href,'.wpg') 
            or contains(./draw:image/@xlink:href,'.jpeg') or contains(./draw:image/@xlink:href,'.gif') or contains(./draw:image/@xlink:href,'.png')
            or contains(./draw:image/@xlink:href,'.jpg')">
                      <xsl:if test="not(./draw:image/@xlink:href[contains(.,'../')])">
                        <xsl:choose>
                          <xsl:when test="@svg:x">
                            <xsl:value-of select="concat(substring-before(@svg:x,'cm'), ':')"/>
                          </xsl:when>
                          <xsl:when test="@draw:transform">
                            <xsl:variable name ="x">
                              <xsl:value-of select="substring-before(substring-before(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
                            </xsl:variable>
                            <xsl:value-of select="concat($x, ':')"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="'0:'"/>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:if>
                    </xsl:if>
                  </xsl:when>
                  <xsl:when test="@presentation:class[contains(.,'title')]
                                  or @presentation:class[contains(.,'subtitle')]
                                  or @presentation:class[contains(.,'outline')]"></xsl:when>
                  <xsl:when test ="(draw:text-box) and not(@presentation:style-name) and not(@presentation:class)">
                    <xsl:choose>
                      <xsl:when test="@svg:x">
                        <xsl:value-of select="concat(substring-before(@svg:x,'cm'), ':')"/>
                      </xsl:when>
                      <xsl:when test="@draw:transform">
                        <xsl:variable name ="x">
                          <xsl:value-of select="substring-before(substring-before(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
                        </xsl:variable>
                        <xsl:value-of select="concat($x, ':')"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="'0:'"/>
                      </xsl:otherwise>
                    </xsl:choose>

                  </xsl:when>
                </xsl:choose>
              </xsl:when>
              <xsl:when test="name()='draw:line' or  name()='draw:connector'">
                <xsl:choose>
                  <xsl:when test="@svg:x1">
                    <xsl:value-of select="concat(substring-before(@svg:x1,'cm'), ':')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'0:'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>

              <xsl:when test="name()='draw:rect' or name()='draw:ellipse' or name()='draw:custom-shape'">

                <xsl:choose>
                  <xsl:when test="@svg:x">
                    <xsl:value-of select="concat(substring-before(@svg:x,'cm'), ':')"/>
                  </xsl:when>
                  <xsl:when test="@draw:transform">
                    <xsl:variable name ="x">
                      <xsl:value-of select="substring-before(substring-before(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
                    </xsl:variable>
                    <xsl:value-of select="concat($x, ':')"/>
                  </xsl:when>
                  <xsl:when test="@draw:transform">
                    <xsl:variable name ="x">
                      <xsl:value-of select="substring-before(substring-before(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
                    </xsl:variable>
                    <xsl:value-of select="concat($x, ':')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'0:'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>
            </xsl:choose>
          </xsl:for-each>
        </xsl:variable>
        <xsl:variable name="grpMinY">
          <xsl:for-each select="node()">
            <xsl:choose>
              <!--<xsl:when test="name()='draw:g'">
                    <xsl:call-template name="tmpgroupingCordinates"/>
                  </xsl:when>-->
              <xsl:when test="name()='draw:frame'">
                <xsl:choose>
                  <xsl:when test="./draw:image">
                    <xsl:if test ="contains(./draw:image/@xlink:href,'.png') or contains(./draw:image/@xlink:href,'.emf') or contains(./draw:image/@xlink:href,'.wmf')
                    or contains(./draw:image/@xlink:href,'.jfif') or contains(./draw:image/@xlink:href,'.jpe') 
                    or contains(./draw:image/@xlink:href,'.bmp') or contains(./draw:image/@xlink:href,'.dib') 
                    or contains(./draw:image/@xlink:href,'.rle')
                    or contains(./draw:image/@xlink:href,'.bmz') or contains(./draw:image/@xlink:href,'.gfa') 
                    or contains(./draw:image/@xlink:href,'.emz') or contains(./draw:image/@xlink:href,'.wmz')
                    or contains(./draw:image/@xlink:href,'.pcz')
            or contains(./draw:image/@xlink:href,'.tif') or contains(./draw:image/@xlink:href,'.tiff') 
            or contains(./draw:image/@xlink:href,'.cdr') or contains(./draw:image/@xlink:href,'.cgm') or contains(./draw:image/@xlink:href,'.eps') 
            or contains(./draw:image/@xlink:href,'.pct') or contains(./draw:image/@xlink:href,'.pict') or contains(./draw:image/@xlink:href,'.wpg') 
            or contains(./draw:image/@xlink:href,'.jpeg') or contains(./draw:image/@xlink:href,'.gif') or contains(./draw:image/@xlink:href,'.png')
            or contains(./draw:image/@xlink:href,'.jpg')">
                      <xsl:if test="not(./draw:image/@xlink:href[contains(.,'../')])">
                        <xsl:choose>
                          <xsl:when test="@svg:y">
                            <xsl:value-of select="concat(substring-before(@svg:y,'cm'), ':')"/>
                          </xsl:when>
                          <xsl:when test="@draw:transform">
                            <xsl:variable name ="y">
                              <xsl:value-of select="substring-before(substring-after(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
                            </xsl:variable>
                            <xsl:value-of select="concat($y, ':')"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="'0:'"/>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:if>
                    </xsl:if>
                  </xsl:when>
                  <xsl:when test="@presentation:class[contains(.,'title')]
                                  or @presentation:class[contains(.,'subtitle')]
                                  or @presentation:class[contains(.,'outline')]"></xsl:when>
                  <xsl:when test ="(draw:text-box) and not(@presentation:style-name) and not(@presentation:class)">
                    <xsl:choose>
                      <xsl:when test="@svg:y">
                        <xsl:value-of select="concat(substring-before(@svg:y,'cm'), ':')"/>
                      </xsl:when>
                      <xsl:when test="@draw:transform">
                        <xsl:variable name ="y">
                          <xsl:value-of select="substring-before(substring-after(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
                        </xsl:variable>
                        <xsl:value-of select="concat($y, ':')"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="'0:'"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:when>
                </xsl:choose>
              </xsl:when>
              <xsl:when test="name()='draw:line' or  name()='draw:connector'">
                <xsl:choose>
                  <xsl:when test="@svg:y1">
                    <xsl:value-of select="concat(substring-before(@svg:y1,'cm'), ':')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'0:'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>
              <xsl:when test="name()='draw:rect' or name()='draw:ellipse' or name()='draw:custom-shape'">
                <xsl:choose>
                  <xsl:when test="@svg:y">
                    <xsl:value-of select="concat(substring-before(@svg:y,'cm'), ':')"/>
                  </xsl:when>
                  <xsl:when test="@draw:transform">
                    <xsl:variable name ="y">
                      <xsl:value-of select="substring-before(substring-after(substring-before(substring-after(substring-after(@draw:transform,'translate'),'('),')'),' '),'cm')" />
                    </xsl:variable>
                    <xsl:value-of select="concat($y, ':')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'0:'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>
            </xsl:choose>
          </xsl:for-each>
        </xsl:variable>
        <xsl:variable name="a-off-x">
          <xsl:value-of select="concat('group-svgXYWidthHeight:onlyX@',$grpMinX)"/>
        </xsl:variable>
        <xsl:variable name="a-off-y">
          <xsl:value-of select="concat('group-svgXYWidthHeight:onlyY@',$grpMinY)"/>
        </xsl:variable>
        <a:off>
          <xsl:attribute name ="x">
            <xsl:value-of select="$a-off-x"/>
          </xsl:attribute>
          <xsl:attribute name ="y">
            <xsl:value-of select="$a-off-y"/>
          </xsl:attribute>
        </a:off>
        <xsl:variable name="a-ext-cx">
          <xsl:value-of select="concat('group-svgXYWidthHeight:CX@',$grpWidth,'$',$grpMinX)"/>
        </xsl:variable>
        <xsl:variable name="a-ext-cy">
          <xsl:value-of select="concat('group-svgXYWidthHeight:CY@',$grpHeight,'$',$grpMinY)"/>
        </xsl:variable>
        <a:ext>
          <xsl:attribute name ="cx">
            <xsl:value-of select="$a-ext-cx"/>
          </xsl:attribute>
          <xsl:attribute name ="cy">
            <xsl:value-of select="$a-ext-cy"/>
          </xsl:attribute>
        </a:ext>
        <a:chOff>
          <xsl:attribute name ="x">
            <xsl:value-of select="$a-off-x"/>
          </xsl:attribute>
          <xsl:attribute name ="y">
            <xsl:value-of select="$a-off-y"/>
          </xsl:attribute>
        </a:chOff>
        <a:chExt>
          <xsl:attribute name ="cx">
            <xsl:value-of select="$a-ext-cx"/>
          </xsl:attribute>
          <xsl:attribute name ="cy">
            <xsl:value-of select="$a-ext-cy"/>
          </xsl:attribute>
        </a:chExt>

      </a:xfrm>
    </p:grpSpPr>

  </xsl:template>
  <xsl:template name="tmpgetCustShapeType">
    <xsl:choose>
      <!-- Text Box -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt202') and (draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N')">
        <xsl:value-of select ="'TextBox Custom '"/>
      </xsl:when>
      <!-- Oval -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='ellipse')">
        <xsl:value-of select ="'Oval Custom '"/>
      </xsl:when>
      <!-- U-Turn Arrow (Added by A.Mathi as on 23/07/2007) -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt100') and (draw:enhanced-geometry/@draw:enhanced-path='M 0 877824 L 0 384048 W 0 0 768096 768096 0 384048 384049 0 L 393192 0 W 9144 0 777240 768096 393192 0 777240 384049 L 777240 438912 L 886968 438912 L 667512 658368 L 448056 438912 L 557784 438912 L 557784 384048 A 228600 219456 557784 548640 557784 384048 393192 219456 L 384048 219456 A 219456 219456 548640 548640 384048 219456 219456 384048 L 219456 877824 Z N')">
        <xsl:value-of select ="'U-Turn Arrow'"/>
      </xsl:when>
      <!-- Isosceles Triangle -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='isosceles-triangle')">
        <xsl:value-of select ="'Isosceles Triangle '"/>
      </xsl:when>
      <!-- Right Arrow (Added by Mathi as on 4/7/2007 -->
      <xsl:when test = "(draw:enhanced-geometry/@draw:type='right-arrow')">
        <xsl:value-of select ="'Right Arrow '"/>
      </xsl:when>
      <!-- Left Arrow (Added by Mathi as on 5/7/2007) -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='left-arrow')">
        <xsl:value-of select ="'Left Arrow '"/>
      </xsl:when>
      <!-- Up Arrow (Added by Mathi as on 19/7/2007)-->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='up-arrow')">
        <xsl:value-of select ="'Up Arrow '"/>
      </xsl:when>
      <!-- Down Arrow (Added by Mathi as on 19/7/2007)-->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='down-arrow')">
        <xsl:value-of select ="'Down Arrow '"/>
      </xsl:when>
      <!-- Left-Right Arrow (Added by Mathi as on 19/7/2007)-->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='left-right-arrow')">
        <xsl:value-of select ="'Left-Right Arrow '"/>
      </xsl:when>
      <!-- Up-Down Arrow (Added by Mathi as on 19/7/2007)-->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='up-down-arrow')">
        <xsl:value-of select ="'Up-Down Arrow '"/>
      </xsl:when>
      <!-- Right Triangle -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='right-triangle')">
        <xsl:value-of select ="'Right Triangle '"/>
      </xsl:when>
      <!-- Parallelogram -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='parallelogram')">
        <xsl:value-of select ="'Parallelogram '"/>
      </xsl:when>
      <!-- Trapezoid (Added by A.Mathi as on 24/07/2007) -->
      <xsl:when test="(draw:enhanced-geometry/@draw:type='mso-spt100') and 
									(draw:enhanced-geometry/@draw:enhanced-path='M 0 1216152 L 228600 0 L 685800 0 L 914400 1216152 Z N')">
        <xsl:value-of select ="'Trapezoid '"/>
      </xsl:when>
      <xsl:when test="(draw:enhanced-geometry/@draw:type='trapezoid')">
        <xsl:value-of select ="'flowchart-manual-operation '"/>
      </xsl:when>
      <!-- Diamond -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='diamond')and
									 (draw:enhanced-geometry/@draw:enhanced-path='M 10800 0 L 21600 10800 10800 21600 0 10800 10800 0 Z N')">
        <xsl:value-of select ="'Diamond '"/>
      </xsl:when>
      <!-- Regular Pentagon -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='pentagon') and
									 (draw:enhanced-geometry/@draw:enhanced-path='M 10800 0 L 0 8260 4230 21600 17370 21600 21600 8260 10800 0 Z N')">
        <xsl:value-of select ="'Regular Pentagon '"/>
      </xsl:when>
      <!-- Hexagon -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='hexagon')">
        <xsl:value-of select ="'Hexagon '"/>
      </xsl:when>
      <!-- Octagon -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='octagon')">
        <xsl:value-of select ="'Octagon '"/>
      </xsl:when>
      <!-- Cube -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='cube')">
        <xsl:value-of select ="'Cube '"/>
      </xsl:when>
      <!-- Can -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='can')">
        <xsl:value-of select ="'Can '"/>
      </xsl:when>
      <!-- Circular Arrow -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='circular-arrow')">
        <xsl:value-of select ="'circular-arrow'"/>
      </xsl:when>
      <!-- Chord -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt100') and 
                       (draw:enhanced-geometry/@draw:enhanced-path='M 780489 780489 W 0 0 914400 914400 780489 780489 457200 0 Z N')">
        <xsl:value-of select ="'Chord'"/>
      </xsl:when>
      <!-- LeftUp Arrow -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt89') and
								 (draw:enhanced-geometry/@draw:enhanced-path='M 0 ?f5 L ?f2 ?f0 ?f2 ?f7 ?f7 ?f7 ?f7 ?f2 ?f0 ?f2 ?f5 0 21600 ?f2 ?f1 ?f2 ?f1 ?f1 ?f2 ?f1 ?f2 21600 Z N')">
        <xsl:value-of select ="'leftUpArrow'"/>
      </xsl:when>
      <!-- BentUp Arrow -->
      <!-- Fix for the bug 24, Internal Defects.xls, date 9th Aug '07, by vijayeta-->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt100') and
								 (draw:enhanced-geometry/@draw:enhanced-path='M 0 1428750 L 2562225 1428750 L 2562225 476250 L 2324100 476250 L 2800350 0 L 3276600 476250 L 3038475 476250 L 3038475 1905000 L 0 1905000 Z N')">
        <xsl:value-of select ="'bentUpArrow'"/>
      </xsl:when>
      <!--End of Fix for the bug 24, Internal Defects.xls, date 9th Aug '07, by vijayeta-->
      <!-- Cross (Added by Mathi on 19/7/2007)-->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='cross')">
        <xsl:value-of select ="'plus '"/>
      </xsl:when>
      <!-- "No Symbol" (Added by A.Mathi as on 19/07/2007)-->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='forbidden')">
        <xsl:value-of select ="'noSmoking '"/>
      </xsl:when>
      <!-- Bent Arrow (Added by A.Mathi as on 23/07/2007) -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt100') and (draw:enhanced-geometry/@draw:enhanced-path='M 0 868680 L 0 457772 W 0 101727 712090 813817 0 457772 356046 101727 L 610362 101727 L 610362 0 L 813816 203454 L 610362 406908 L 610362 305181 L 356045 305181 A 203454 305181 508636 610363 356045 305181 203454 457772 L 203454 868680 Z N')">
        <xsl:value-of select ="'bentArrow '"/>
      </xsl:when>
      <!--  Folded Corner (Added by A.Mathi as on 19/07/2007) -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='paper')">
        <xsl:value-of select ="'foldedCorner '"/>
      </xsl:when>
      <!--  Lightning Bolt (Added by A.Mathi as on 20/07/2007) -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:enhanced-path='M 640 233 L 221 293 506 12 367 0 29 406 431 347 145 645 99 520 0 861 326 765 209 711 640 233 640 233 Z N')">
        <xsl:value-of select ="'lightningBolt '"/>
      </xsl:when>
      <!--  Explosion 1 (Added by A.Mathi as on 20/07/2007) -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='bang')">
        <xsl:value-of select ="'irregularSeal1 '"/>
      </xsl:when>
      <!-- Left Bracket (Added by A.Mathi as on 20/07/2007) -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='left-bracket')">
        <xsl:value-of select ="'Left Bracket '"/>
      </xsl:when>
      <!-- Right Bracket (Added by A.Mathi as on 20/07/2007) -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='right-bracket')">
        <xsl:value-of select ="'Right Bracket '"/>
      </xsl:when>
      <!-- Left Brace (Added by A.Mathi as on 20/07/2007) -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='left-brace')">
        <xsl:value-of select ="'Left Brace '"/>
      </xsl:when>
      <!-- Right Brace (Added by A.Mathi as on 23/07/2007) -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='right-brace')">
        <xsl:value-of select ="'Right Brace '"/>
      </xsl:when>
      <!-- Rectangular Callout (Added by A.Mathi as on 23/07/2007) -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='rectangular-callout')">
        <xsl:value-of select ="'Rectangular Callout '"/>
      </xsl:when>
      <!-- Rounded Rectangular Callout (Added by A.Mathi as on 23/07/2007) -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='round-rectangular-callout')">
        <xsl:value-of select ="'wedgeRoundRectCallout '"/>
      </xsl:when>
      <!-- Oval Callout (Added by A.Mathi as on 23/07/2007) -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='round-callout')">
        <xsl:value-of select ="'wedgeEllipseCallout '"/>
      </xsl:when>
      <!-- Cloud Callout (Added by A.Mathi as on 23/07/2007) -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='cloud-callout')">
        <xsl:value-of select ="'Cloud Callout '"/>
      </xsl:when>
      <!-- Rectangle -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='rectangle') and (draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N')">
        <xsl:value-of select ="'Rectangle Custom '"/>
      </xsl:when>
      <!-- Rounded Rectangle -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='round-rectangle')">
        <xsl:value-of select ="'Rounded Rectangle '"/>
      </xsl:when>
      <!-- Snip Single Corner Rectangle -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-card') and 
									 (draw:enhanced-geometry/@draw:mirror-horizontal='true') and
									 (draw:enhanced-geometry/@draw:enhanced-path='M 4300 0 L 21600 0 21600 21600 0 21600 0 4300 4300 0 Z N')">
        <xsl:value-of select ="'Snip Single Corner Rectangle '"/>
      </xsl:when>
      <!-- Snip Same Side Corner Rectangle -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt100') and 
									 (draw:enhanced-geometry/@draw:enhanced-path='M 50801 0 L 863599 0 L 914400 50801 L 914400 304800 L 914400 304800 L 0 304800 L 0 304800 L 0 50801 Z N')">
        <xsl:value-of select ="'Snip Same Side Corner Rectangle '"/>
      </xsl:when>
      <!-- Snip Diagonal Corner Rectangle -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt100') and
									 (draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 876299 0 L 914400 38101 L 914400 228600 L 914400 228600 L 38101 228600 L 0 190499 L 0 0 Z N')">
        <xsl:value-of select ="'Snip Diagonal Corner Rectangle '"/>
      </xsl:when>
      <!-- Snip and Round Single Corner Rectangle -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt100') and
									 (draw:enhanced-geometry/@draw:enhanced-path='M 38101 0 L 876299 0 L 914400 38101 L 914400 228600 L 0 228600 L 0 38101 W 0 0 76202 76202 0 38101 38101 0 Z N')">
        <xsl:value-of select ="'Snip and Round Single Corner Rectangle '"/>
      </xsl:when>
      <!-- Round Single Corner Rectangle -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt100') and
									 (draw:enhanced-geometry/@draw:enhanced-path='M 0 0 L 876299 0 W 838198 0 914400 76202 876299 0 914400 38101 L 914400 228600 L 0 228600 Z N')">
        <xsl:value-of select ="'Round Single Corner Rectangle '"/>
      </xsl:when>
      <!-- Round Same Side Corner Rectangle -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt100') and
									 (draw:enhanced-geometry/@draw:enhanced-path='M 152403 0 L 761997 0 W 609594 0 914400 304806 761997 0 914400 152403 L 914400 914400 L 0 914400 L 0 152403 W 0 0 304806 304806 0 152403 152403 0 Z N')">
        <xsl:value-of select ="'Round Same Side Corner Rectangle '"/>
      </xsl:when>
      <!-- Round Diagonal Corner Rectangle -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='mso-spt100') and
									 (draw:enhanced-geometry/@draw:enhanced-path='M 254005 0 L 2286000 0 L 2286000 1269995 W 1777990 1015990 2286000 1524000 2286000 1269995 2031995 1524000 L 0 1524000 L 0 254005 W 0 0 508010 508010 0 254005 254005 0 Z N')">
        <xsl:value-of select ="'Round Diagonal Corner Rectangle '"/>
      </xsl:when>

      <!-- Flow Chart: Process -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-process')">
        <xsl:value-of select ="'Flowchart: Process '"/>
      </xsl:when>
      <!-- Flow Chart: Alternate Process -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-alternate-process')">
        <xsl:value-of select ="'Flowchart: Alternate Process '"/>
      </xsl:when>
      <!-- Flow Chart: Decision -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-decision')">
        <xsl:value-of select ="'Flowchart: Decision '"/>
      </xsl:when>
      <!-- Flow Chart: Data -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-data')">
        <xsl:value-of select ="'Flowchart: Data '"/>
      </xsl:when>
      <!-- Flow Chart: Predefined-process -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-predefined-process')">
        <xsl:value-of select ="'Flowchart: Predefined Process '"/>
      </xsl:when>
      <!-- Flow Chart: Internal-storage -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-internal-storage')">
        <xsl:value-of select ="'Flowchart: Internal Storage '"/>
      </xsl:when>
      <!-- Flow Chart: Document -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-document')">
        <xsl:value-of select ="'Flowchart: Document '"/>
      </xsl:when>
      <!-- Flow Chart: Multidocument -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-multidocument')">
        <xsl:value-of select ="'Flowchart: Multi document '"/>
      </xsl:when>
      <!-- Flow Chart: Terminator -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-terminator')">
        <xsl:value-of select ="'Flowchart: Terminator '"/>
      </xsl:when>
      <!-- Flow Chart: Preparation -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-preparation')">
        <xsl:value-of select ="'Flowchart: Preparation '"/>
      </xsl:when>
      <!-- Flow Chart: Manual-input -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-manual-input')">
        <xsl:value-of select ="'Flowchart: Manual Input '"/>
      </xsl:when>
      <!-- Flow Chart: Manual-operation -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-manual-operation')">
        <xsl:value-of select ="'Flowchart: Manual Operation '"/>
      </xsl:when>
      <!-- Flow Chart: Connector -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-connector')">
        <xsl:value-of select ="'Flowchart: Connector '"/>
      </xsl:when>
      <!-- Flow Chart: Off-page-connector -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-off-page-connector')">
        <xsl:value-of select ="'Flowchart: Off-page Connector '"/>
      </xsl:when>
      <!-- Flow Chart: Card -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-card')">
        <xsl:value-of select ="'Flowchart: Card '"/>
      </xsl:when>
      <!-- Flow Chart: Punched-tape -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-punched-tape')">
        <xsl:value-of select ="'Flowchart: Punched Tape '"/>
      </xsl:when>
      <!-- Flow Chart: Summing-junction -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-summing-junction')">
        <xsl:value-of select ="'Flowchart: Summing Junction '"/>
      </xsl:when>
      <!-- Flow Chart: Or -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-or')">
        <xsl:value-of select ="'Flowchart: Or '"/>
      </xsl:when>
      <!-- Flow Chart: Collate -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-collate')">
        <xsl:value-of select ="'Flowchart: Collate '"/>
      </xsl:when>
      <!-- Flow Chart: Sort -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-sort')">
        <xsl:value-of select ="'Flowchart: Sort '"/>
      </xsl:when>
      <!-- Flow Chart: Extract -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-extract')">
        <xsl:value-of select ="'Flowchart: Extract '"/>
      </xsl:when>
      <!-- Flow Chart: Merge -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-merge')">
        <xsl:value-of select ="'Flowchart: Merge '"/>
      </xsl:when>
      <!-- Flow Chart: Stored-data -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-stored-data')">
        <xsl:value-of select ="'Flowchart: Stored Data '"/>
      </xsl:when>
      <!-- Flow Chart: Delay-->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-delay')">
        <xsl:value-of select ="'Flowchart: Delay '"/>
      </xsl:when>
      <!-- Flow Chart: Sequential-access -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-sequential-access')">
        <xsl:value-of select ="'Flowchart: Sequential Access Storage '"/>
      </xsl:when>
      <!-- Flow Chart: Direct-access-storage -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-direct-access-storage')">
        <xsl:value-of select ="'Flowchart: Direct Access Storage'"/>
      </xsl:when>
      <!-- Flow Chart: Magnetic-disk --> 
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-magnetic-disk')">
        <xsl:value-of select ="'Flowchart: Magnetic Disk '"/>
      </xsl:when>
      <!-- Flow Chart: Display -->
      <xsl:when test ="(draw:enhanced-geometry/@draw:type='flowchart-display')">
        <xsl:value-of select ="'Flowchart: Display '"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
 <xsl:template name="tmpSuperSubScriptForward">
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
      <xsl:when test="contains(substring-before(style:text-properties/@style:text-position,' '),'%') = '%'">
        <xsl:variable name="val">
          <xsl:value-of select="number(substring-before(substring-before(style:text-properties/@style:text-position,' '),'%'))"/>
        </xsl:variable>
        <xsl:attribute name="baseline">
          <xsl:value-of select="$val * 1000"/>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>

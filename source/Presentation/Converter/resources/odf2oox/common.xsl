<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xsl:stylesheet version="1.0" 
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" 
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"  
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" 
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  exclude-result-prefixes="style draw fo">
	
	<xsl:template name ="convertToPoints">
		<xsl:param name ="unit"/>
		<xsl:param name ="length"/>
		<xsl:variable name="lengthVal">
			<xsl:choose>
				<xsl:when test="contains($length,'cm')">
					<xsl:value-of select="substring-before($length,'cm')"/>
				</xsl:when>
				<xsl:when test="contains($length,'pt')">
					<xsl:value-of select="substring-before($length,'pt')"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="$length"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:choose>
			<xsl:when test="$lengthVal='0' or $lengthVal='' or $lengthVal &lt; 0 ">
				<xsl:value-of select="0"/>
			</xsl:when>
			<xsl:when test="$unit = 'cm'">
				<xsl:value-of select="concat(format-number($lengthVal * 360000,'#.###'),'')"/>
			</xsl:when>
			<xsl:when test="$unit = 'mm'">
				<xsl:value-of select="concat(format-number($lengthVal * 25.4 div 72,'#.###'),'mm')"/>
			</xsl:when>
			<xsl:when test="$unit = 'in'">
				<xsl:value-of select="concat(format-number($lengthVal div 72,'#.###'),'in')"/>
			</xsl:when>
			<xsl:when test="$unit = 'pt'">
				<xsl:value-of select="concat($lengthVal * 100 ,'')"/>
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
	</xsl:template>

	<xsl:template name ="fontStyles">
		<xsl:param name ="Tid"/>
		<xsl:for-each  select ="document('content.xml')/office:document-content/office:automatic-styles/style:style[@style:name =$Tid ]">
			<xsl:if test="style:text-properties/@fo:font-size">
				<xsl:attribute name ="sz">
					<xsl:call-template name ="convertToPoints">
						<xsl:with-param name ="unit" select ="'pt'"/>
						<xsl:with-param name ="length" select ="style:text-properties/@fo:font-size"/>
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>
			<!--Font bold attribute -->
			<xsl:if test="style:text-properties/@fo:font-weight[contains(.,'bold')]">
				<xsl:attribute name ="b">
					<xsl:value-of select ="'1'"/>
				</xsl:attribute >
			</xsl:if >
			<!-- Font Inclined-->
			<xsl:if test="style:text-properties/@fo:font-style[contains(.,'italic')]">
				<xsl:attribute name ="i">
					<xsl:value-of select ="'1'"/>
				</xsl:attribute >
			</xsl:if >
			<!-- Font underline-->
			<xsl:if test="style:text-properties/@style:text-underline-style[contains(.,'solid')]">
				<xsl:attribute name ="u">
					<xsl:value-of select ="'sng'"/>
				</xsl:attribute >
			</xsl:if >
			<!-- Font Strike through-->
			<xsl:if test="style:text-properties/@style:text-line-through-style[contains(.,'solid')]">
				<xsl:attribute name ="strike">
					<xsl:value-of select ="'sngStrike'"/>
				</xsl:attribute >
			</xsl:if >
			<!--Charector spacing -->
			<xsl:if test ="style:text-properties/@fo:letter-spacing">
				<xsl:attribute name ="spc">
					<xsl:if test ="substring-before(style:text-properties/@fo:letter-spacing,'cm')&lt; 0 ">
						<xsl:value-of select ="format-number(substring-before(style:text-properties/@fo:letter-spacing,'cm') * 3.5 *1000 ,'#.##')"/>
					</xsl:if >
					<xsl:if test ="substring-before(style:text-properties/@fo:letter-spacing,'cm')
				&gt; 0 or substring-before(style:text-properties/@fo:letter-spacing,'cm') = 0 ">
						<xsl:value-of select ="format-number((substring-before(style:text-properties/@fo:letter-spacing,'cm') div .035) *100 ,'#.##')"/>
					</xsl:if>
				</xsl:attribute>
			</xsl:if >
			<!-- Text Shadow -->
			<xsl:if test="style:text-properties/@fo:text-shadow[contains(.,'pt')]">
				<a:effectLst>
					<a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl">
						<a:srgbClr val="000000">
							<a:alpha val="43137" />
						</a:srgbClr>
					</a:outerShdw>
				</a:effectLst>
			</xsl:if >
			<!--Color Node set as standard colors -->
			<xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
			<xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
			<a:solidFill>
				<a:srgbClr  >
					<xsl:attribute name ="val">
						<!--<xsl:value-of   select ="substring-after(style:text-properties/@fo:color,'#')"/>-->
						<xsl:value-of select ="translate(substring-after(style:text-properties/@fo:color,'#'),$lcletters,$ucletters)"/>
					</xsl:attribute>
				</a:srgbClr >
			</a:solidFill>
		</xsl:for-each >
	</xsl:template>

	<xsl:template name ="paraProperties" >
		<xsl:param name ="paraId" />
		<xsl:for-each select ="document('content.xml')//style:style[@style:name=$paraId]">
			<a:pPr>
				<!--marL="first line indent property"-->
				<xsl:if test ="style:paragraph-properties/@fo:text-indent">
					<xsl:attribute name ="indent">
						<!--fo:text-indent-->
						<xsl:call-template name ="convertToPoints">
							<xsl:with-param name="length"  select ="style:paragraph-properties/@fo:text-indent"/>
							<xsl:with-param name ="unit" select ="'cm'"/>
						</xsl:call-template>
					</xsl:attribute>
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
							<xsl:otherwise >
								<xsl:value-of select ="l"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
				</xsl:if >
				<xsl:if test ="style:paragraph-properties/@fo:margin-left">
					<xsl:attribute name ="marL">
						<!--fo:margin-left-->
						<xsl:call-template name ="convertToPoints">
							<xsl:with-param name="length"  select ="style:paragraph-properties/@fo:margin-left"/>
							<xsl:with-param name ="unit" select ="'cm'"/>
						</xsl:call-template>
					</xsl:attribute>
				</xsl:if >
				<xsl:if test ="style:paragraph-properties/@fo:margin-top">
					<a:spcBef>
						<a:spcPts>
							<xsl:attribute name ="val">
								<!--fo:margin-top-->
								<xsl:value-of select ="format-number(substring-before(style:paragraph-properties/@fo:margin-top,'cm')* 1000,'#.##') "/>
							</xsl:attribute>
						</a:spcPts>
					</a:spcBef >
				</xsl:if>
				<xsl:if test ="style:paragraph-properties/@fo:margin-bottom">
					<a:spcAft>
						<a:spcPts>
							<xsl:attribute name ="val">
								<!--fo:margin-bottom-->
								<xsl:value-of select ="format-number(substring-before(style:paragraph-properties/@fo:margin-bottom,'cm')* 1000,'#.##') "/>
							</xsl:attribute>
						</a:spcPts>
					</a:spcAft>
				</xsl:if >
				<xsl:if test ="style:paragraph-properties/@fo:line-height">
					<a:lnSpc>
						<a:spcPts>
							<xsl:attribute name ="val">
								<!--fo:line-height-->
								<xsl:value-of select ="format-number(substring-before(style:paragraph-properties/@fo:line-height,'%')* 1000,'#.##') "/>
							</xsl:attribute>
						</a:spcPts>
					</a:lnSpc>
				</xsl:if>
			</a:pPr>
		</xsl:for-each >
	</xsl:template>
	
	<xsl:template name ="fillColor">
		<xsl:param name ="prId"/>
		<xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
		<xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
		<xsl:for-each select ="document('content.xml')/office:document-content/office:automatic-styles/style:style[@style:name=$prId] ">
			<xsl:if test ="style:graphic-properties/@draw:fill-color">
				<a:solidFill>
					<a:srgbClr  >
						<xsl:attribute name ="val">
							<xsl:value-of select ="translate(substring-after(style:graphic-properties/@draw:fill-color,'#'),$lcletters,$ucletters)"/>
						</xsl:attribute>
					</a:srgbClr >
				</a:solidFill>
			</xsl:if>
		</xsl:for-each>
	</xsl:template>


</xsl:stylesheet>

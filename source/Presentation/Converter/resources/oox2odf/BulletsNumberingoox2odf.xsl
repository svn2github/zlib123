<?xml version="1.0" encoding="utf-8" standalone="yes" ?>
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
  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:rels="http://schemas.openxmlformats.org/package/2006/relationships"
xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  exclude-result-prefixes="odf style text number draw page">

	<xsl:import href="common.xsl"/>
	<xsl:template name ="insertBulletsNumbersoox2odf">
		<xsl:param name ="listStyleName" />
		<xsl:param  name ="ParaId"/>
		<!--condition,If Levels Present-->
		<xsl:if test ="./a:pPr/@lvl">
			<xsl:call-template name ="insertMultipleLevels">
				<xsl:with-param name ="levelCount" select ="./a:pPr/@lvl"/>
				<xsl:with-param name ="ParaId" select ="$ParaId"/>
				<xsl:with-param name ="listStyleName" select ="$listStyleName"/>
			</xsl:call-template>
		</xsl:if>
		<!-- Fix by vijayeta,on 22nd june,bug number1739817,also test if explicit lvl=0-->
		<xsl:if test ="not(./a:pPr/@lvl) or ./a:pPr/@lvl='0'">
			<text:list>
				<xsl:attribute name ="text:style-name">
					<xsl:value-of select ="$listStyleName"/>
				</xsl:attribute>
				<text:list-item>
					<text:p >
						<xsl:attribute name ="text:style-name">
							<xsl:value-of select ="concat($ParaId,position())"/>
						</xsl:attribute>
						<xsl:for-each select ="node()">
							<xsl:if test ="name()='a:r'">
								<text:span text:style-name="{generate-id()}">
									<!--<xsl:value-of select ="a:t"/>-->
									<!--converts whitespaces sequence to text:s-->
									<!-- 1699083 bug fix  -->
									<xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
									<xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
									<xsl:choose >
										<xsl:when test ="a:rPr[@cap='all']">
											<xsl:choose >
												<xsl:when test =".=''">
													<text:s/>
												</xsl:when>
												<xsl:when test ="not(contains(.,'  '))">
													<xsl:value-of select ="translate(.,$lcletters,$ucletters)"/>
												</xsl:when>
												<xsl:when test =". =' '">
													<text:s/>
												</xsl:when>
												<xsl:otherwise >
													<xsl:call-template name ="InsertWhiteSpaces">
														<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
													</xsl:call-template>
												</xsl:otherwise>
											</xsl:choose>
										</xsl:when>
										<xsl:when test ="a:rPr[@cap='small']">
											<xsl:choose >
												<xsl:when test =".=''">
													<text:s/>
												</xsl:when>
												<xsl:when test ="not(contains(.,'  '))">
													<xsl:value-of select ="translate(.,$ucletters,$lcletters)"/>
												</xsl:when>
												<xsl:when test =".= ' '">
													<text:s/>
												</xsl:when>
												<xsl:otherwise >
													<xsl:call-template name ="InsertWhiteSpaces">
														<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
													</xsl:call-template>
												</xsl:otherwise>
											</xsl:choose >
										</xsl:when>
										<xsl:otherwise >
											<xsl:choose >
												<xsl:when test =".=''">
													<text:s/>
												</xsl:when>
												<xsl:when test ="not(contains(.,'  '))">
													<xsl:value-of select ="."/>
												</xsl:when>
												<xsl:otherwise >
													<xsl:call-template name ="InsertWhiteSpaces">
														<xsl:with-param name ="string" select ="."/>
													</xsl:call-template>
												</xsl:otherwise >
											</xsl:choose>
										</xsl:otherwise>
									</xsl:choose>
								</text:span>
							</xsl:if >
							<xsl:if test ="name()='a:br'">
								<text:line-break/>
							</xsl:if>
						</xsl:for-each>
					</text:p>
				</text:list-item>
			</text:list>
		</xsl:if>
		<!--End of condition,If Levels Present-->
	</xsl:template>
	<xsl:template name ="insertBulletStyle">
		<xsl:param name ="slideRel" />
		<xsl:param name ="ParaId" />
		<xsl:param name ="listStyleName"/>
		<xsl:param name ="slideMaster" />
		<xsl:param name ="var_TextBoxType" />
		<xsl:param name ="var_index" />
		<xsl:param name ="layoutName"/>
		<xsl:if test ="./a:p/a:pPr/@lvl">
			<xsl:for-each select ="./a:p">
				<!-- Addded by vijayeta on 18th june,to check for bunone-->
				<xsl:if test ="not(a:pPr/a:buNone)">
					<!-- Addded by vijayeta on 18th june,to check for bunone-->
					<xsl:if test ="a:pPr/@lvl">
						<xsl:variable name ="levelStyle">
							<xsl:value-of select ="concat($listStyleName,position(),'lvl',a:pPr/@lvl)"/>
						</xsl:variable>
						<xsl:variable name ="textLevel" select ="a:pPr/@lvl"/>
						<xsl:variable name ="newTextLvl" select ="$textLevel+1"/>
						<text:list-style>
							<xsl:attribute name ="style:name">
								<xsl:value-of select ="$levelStyle"/>
							</xsl:attribute >
							<xsl:choose >
								<xsl:when test ="a:pPr/a:buChar">
									<text:list-level-style-bullet>
										<xsl:attribute name ="text:level">
											<xsl:value-of select ="$textLevel + 1"/>
										</xsl:attribute >
										<xsl:attribute name ="text:bullet-char">
											<xsl:call-template name ="insertBulletCharacter">
												<xsl:with-param name ="character" select ="a:pPr/a:buChar/@char" />
											</xsl:call-template>
										</xsl:attribute >
										<style:list-level-properties text:min-label-width="0.8cm">
											<xsl:attribute name ="text:space-before">
												<xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
													<xsl:value-of select="concat(format-number(child::node()[$newTextLvl]/@marL div 360000,'#.##'),'cm')"/>
												</xsl:for-each>
											</xsl:attribute>
										</style:list-level-properties>
										<style:text-properties style:font-charset="x-symbol">
											<xsl:if test ="a:pPr/a:buFont/@typeface">
												<xsl:if test ="a:pPr/a:buFont[@typeface='Arial']">
													<xsl:attribute name ="fo:font-family">
														<xsl:value-of select ="'StarSymbol'"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test ="not(a:pPr/a:buFont[@typeface='Arial'])">
													<xsl:attribute name ="fo:font-family">
														<xsl:value-of select ="a:pPr/a:buFont/@typeface"/>
													</xsl:attribute>
												</xsl:if>
											</xsl:if>
											<xsl:if test ="not(a:pPr/a:buFont/@typeface)">
												<xsl:attribute name ="fo:font-family">
													<xsl:value-of select ="'StarSymbol'"/>
												</xsl:attribute>
											</xsl:if>
											<xsl:if test ="a:pPr/a:buSzPct">
												<xsl:attribute name ="fo:font-size">
													<xsl:value-of select ="concat((a:pPr/a:buSzPct/@val div 1000),'%')"/>
												</xsl:attribute>
											</xsl:if>
											<xsl:if test ="not(a:pPr/a:buSzPct)">
												<xsl:attribute name ="fo:font-size">
													<xsl:value-of select ="'100%'"/>
												</xsl:attribute>
											</xsl:if>
											<xsl:if test ="a:pPr/a:buClr">
												<xsl:if test ="a:pPr/a:buClr/a:srgbClr">
													<xsl:variable name ="color" select ="a:pPr/a:buClr/a:srgbClr/@val"/>
													<xsl:attribute name ="fo:color">
														<xsl:value-of select ="concat('#',$color)"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test ="a:pPr/a:buClr/a:schemeClr">
													<xsl:call-template name ="getColorCode">
														<xsl:with-param name ="color">
															<xsl:value-of select="a:pPr/a:buClr/a:schemeClr/@val"/>
														</xsl:with-param>
													</xsl:call-template >
												</xsl:if>
											</xsl:if>
											<!-- Code added by vijayeta, bug fix 1746350-->
											<xsl:if test ="not(a:pPr/a:buClr)">
												<xsl:if test ="a:r/a:rPr/a:solidFill">
													<xsl:if test ="a:r/a:rPr/a:solidFill/a:srgbClr">
														<xsl:attribute name ="fo:color">
															<xsl:value-of select ="concat('#',a:r/a:rPr/a:solidFill/a:srgbClr/@val)"/>
														</xsl:attribute>
													</xsl:if>
													<xsl:if test ="a:r/a:rPr/a:solidFill/a:schemeClr">
														<xsl:attribute name ="fo:color">
															<xsl:call-template name ="getColorCode">
																<xsl:with-param name ="color">
																	<xsl:value-of select="a:r/a:rPr/a:solidFill/a:schemeClr/@val"/>
																</xsl:with-param>
															</xsl:call-template >
														</xsl:attribute>
													</xsl:if>
												</xsl:if>
												<xsl:if test ="not(a:r/a:rPr/a:solidFill)">
													<xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
														<xsl:variable name ="levelColor" >
															<xsl:call-template name ="getLevelColor">
																<xsl:with-param name ="levelNum" select ="$newTextLvl"/>
															</xsl:call-template>
														</xsl:variable>
														<xsl:attribute name ="fo:color">
															<xsl:value-of select ="concat('#',$levelColor)"/>
														</xsl:attribute>
													</xsl:for-each>
												</xsl:if>
											</xsl:if>
											<!--End of Code added by vijayeta, bug fix 1746350-->
										</style:text-properties >
									</text:list-level-style-bullet>
								</xsl:when>
								<xsl:when test ="a:pPr/a:buAutoNum">
									<text:list-level-style-number>
										<xsl:attribute name ="text:level">
											<xsl:value-of select ="$textLevel + 1"/>
										</xsl:attribute >
										<xsl:variable name ="startAt">
											<xsl:if test ="a:pPr/a:buAutoNum/@startAt">
												<xsl:value-of select ="a:pPr/a:buAutoNum/@startAt" />
											</xsl:if>
											<xsl:if test ="not(a:pPr/a:buAutoNum/@startAt)">
												<xsl:value-of select ="'1'" />
											</xsl:if>
										</xsl:variable>
										<xsl:call-template name ="insertNumber">
											<xsl:with-param name ="number" select ="a:pPr/a:buAutoNum/@type"/>
											<xsl:with-param name ="startAt" select ="$startAt"/>
										</xsl:call-template>
										<style:list-level-properties text:min-label-width="0.952cm" />
										<style:text-properties style:font-family-generic="swiss" style:font-pitch="variable">
											<xsl:if test ="a:pPr/a:buFont/@typeface">
												<xsl:attribute name ="fo:font-family">
													<xsl:value-of select ="a:pPr/a:buFont/@typeface"/>
												</xsl:attribute>
											</xsl:if>
											<xsl:if test ="not(a:pPr/a:buFont/@typeface)">
												<xsl:attribute name ="fo:font-family">
													<xsl:value-of select ="'Arial'"/>
												</xsl:attribute>
											</xsl:if>
											<xsl:if test ="not(a:pPr/a:buSzPct)">
												<xsl:attribute name ="fo:font-size">
													<xsl:value-of select ="'100%'"/>
												</xsl:attribute>
											</xsl:if>
											<xsl:if test ="a:pPr/a:buClr">
												<xsl:if test ="a:p/a:pPr/a:buClr/a:srgbClr">
													<xsl:attribute name ="fo:color">
														<xsl:value-of select ="concat('#',a:p/a:pPr/a:buClr/a:srgbClr/@val)"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test ="a:pPr/a:buClr/a:schemeClr">
													<xsl:call-template name ="getColorCode">
														<xsl:with-param name ="color">
															<xsl:value-of select="a:pPr/a:buClr/a:schemeClr/@val"/>
														</xsl:with-param>
													</xsl:call-template >
												</xsl:if>
											</xsl:if>
											<!-- Code added by vijayeta, bug fix 1746350-->
											<xsl:if test ="not(a:pPr/a:buClr)">
												<xsl:if test ="a:r/a:rPr/a:solidFill">
													<xsl:if test ="a:r/a:rPr/a:solidFill/a:srgbClr">
														<xsl:attribute name ="fo:color">
															<xsl:value-of select ="concat('#',a:r/a:rPr/a:solidFill/a:srgbClr/@val)"/>
														</xsl:attribute>
													</xsl:if>
													<xsl:if test ="a:r/a:rPr/a:solidFill/a:schemeClr">
														<xsl:attribute name ="fo:color">
															<xsl:call-template name ="getColorCode">
																<xsl:with-param name ="color">
																	<xsl:value-of select="a:r/a:rPr/a:solidFill/a:schemeClr/@val"/>
																</xsl:with-param>
															</xsl:call-template >
														</xsl:attribute>
													</xsl:if>
												</xsl:if>
												<xsl:if test ="not(a:r/a:rPr/a:solidFill)">
													<xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
														<xsl:variable name ="levelColor" >
															<xsl:call-template name ="getLevelColor">
																<xsl:with-param name ="levelNum" select ="$newTextLvl"/>
															</xsl:call-template>
														</xsl:variable>
														<xsl:attribute name ="fo:color">
															<xsl:value-of select ="concat('#',$levelColor)"/>
														</xsl:attribute>
													</xsl:for-each>
												</xsl:if>
											</xsl:if>
											<!--End of Code added by vijayeta, bug fix 1746350-->
										</style:text-properties>
									</text:list-level-style-number>
								</xsl:when>
								<xsl:when test ="a:pPr/a:buBlip">
									<xsl:variable name ="rId" select ="a:pPr/a:buBlip/a:blip/@r:embed"/>
									<xsl:variable name="XlinkHref">
										<xsl:variable name="pzipsource">
											<xsl:value-of select="document($slideRel)//rels:Relationships/rels:Relationship[@Id=$rId]/@Target"/>
										</xsl:variable>
										<xsl:value-of select="concat('Pictures/', substring-after($pzipsource,'../media/'))"/>
									</xsl:variable>
									<xsl:call-template name="copyPictures">
										<xsl:with-param name="document">
											<xsl:value-of select="$slideRel"/>
										</xsl:with-param>
										<xsl:with-param name="rId">
											<xsl:value-of select ="$rId"/>
										</xsl:with-param>
									</xsl:call-template>
									<text:list-style>
										<xsl:attribute name ="style:name">
											<xsl:value-of select ="$listStyleName"/>
										</xsl:attribute >
										<text:list-level-style-image text:level="1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad">
											<xsl:attribute name="xlink:href">
												<xsl:value-of select="$XlinkHref"/>
											</xsl:attribute>
											<style:list-level-properties style:vertical-pos="middle" style:vertical-rel="line" fo:width="0.318cm" fo:height="0.318cm" />
										</text:list-level-style-image>
									</text:list-style>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle/child::node()[1]">
										<xsl:if test ="a:buChar">
											<text:list-level-style-bullet>
												<xsl:attribute name ="text:level">
													<xsl:value-of select ="$textLevel + 1"/>
												</xsl:attribute >
												<xsl:attribute name ="text:bullet-char">
													<xsl:choose>
														<xsl:when test="a:buChar/@char= '•'">
															<xsl:value-of select ="'•'"/>
														</xsl:when>
														<xsl:when test="a:buChar/@char= 'Ø'">
															<xsl:value-of select ="'➢'"/>
														</xsl:when>
														<xsl:when test="a:buChar/@char= 'o'">
															<xsl:value-of select ="'○'"/>
														</xsl:when>
														<!--<xsl:when test="a:buChar/@char= '§'">
                              <xsl:value-of select ="'■'"/>
                            </xsl:when>
                            <xsl:when test="a:buChar/@char= 'q'">
                              <xsl:value-of select ="''"/>
                            </xsl:when>-->
														<xsl:when test="a:buChar/@char= 'ü'">
															<xsl:value-of select ="'✔'"/>
														</xsl:when>
														<xsl:when test="a:buChar/@char = '-'">
															<xsl:value-of select ="'-'"/>
														</xsl:when>
														<xsl:when test="a:buChar/@char = '»'">
															<xsl:value-of select ="'»'"/>
														</xsl:when>
														<!-- Added by  vijayeta ,on 19th june-->
														<xsl:when test="a:buChar/@char = 'è'">
															<xsl:value-of select ="'➔'"/>
														</xsl:when>
														<!-- Added by  vijayeta ,on 19th june-->
														<xsl:otherwise>•</xsl:otherwise>
													</xsl:choose >
												</xsl:attribute >
												<style:list-level-properties text:min-label-width="0.952cm" />
												<style:text-properties style:font-family-generic="swiss" style:font-pitch="variable">
													<xsl:if test ="a:pPr/a:buFont/@typeface">
														<xsl:if test ="a:pPr/a:buFont[@typeface='Arial']">
															<xsl:attribute name ="fo:font-family">
																<xsl:value-of select ="'StarSymbol'"/>
															</xsl:attribute>
														</xsl:if>
														<xsl:if test ="not(a:pPr/a:buFont[@typeface='Arial'])">
															<xsl:attribute name ="fo:font-family">
																<xsl:value-of select ="a:pPr/a:buFont/@typeface"/>
															</xsl:attribute>
														</xsl:if>
													</xsl:if>
													<xsl:if test ="not(a:pPr/a:buFont/@typeface)">
														<xsl:attribute name ="fo:font-family">
															<xsl:value-of select ="'StarSymbol'"/>
														</xsl:attribute>
													</xsl:if>
													<xsl:if test ="a:buSzPct">
														<xsl:attribute name ="fo:font-size">
															<xsl:value-of select ="concat((a:buSzPct/@val div 1000),'%')"/>
														</xsl:attribute>
													</xsl:if>
													<xsl:if test ="not(a:buSzPct)">
														<xsl:attribute name ="fo:font-size">
															<xsl:value-of select ="'100%'"/>
														</xsl:attribute>
													</xsl:if>
													<xsl:if test ="a:buClr">
														<xsl:if test ="a:buClr/a:srgbClr">
															<xsl:variable name ="color" select ="a:buClr/a:srgbClr/@val"/>
															<xsl:attribute name ="fo:color">
																<xsl:value-of select ="concat('#',$color)"/>
															</xsl:attribute>
														</xsl:if>
														<xsl:if test ="a:buClr/a:schemeClr">
															<xsl:call-template name ="getColorCode">
																<xsl:with-param name ="color">
																	<xsl:value-of select="a:buClr/a:schemeClr/@val"/>
																</xsl:with-param>
															</xsl:call-template >
														</xsl:if>
													</xsl:if>
												</style:text-properties >
											</text:list-level-style-bullet>
										</xsl:if>
										<xsl:if test ="a:buAutoNum">
											<text:list-level-style-number>
												<xsl:attribute name ="text:level">
													<xsl:value-of select ="$textLevel + 1"/>
												</xsl:attribute >
												<xsl:if test ="a:buAutoNum/@type[contains(.,'arabicPeriod')]">
													<xsl:attribute name ="style:num-format" >
														<xsl:value-of  select ="'1'"/>
													</xsl:attribute >
													<xsl:attribute name ="style:num-suffix">
														<xsl:value-of select ="'.'"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test ="a:buAutoNum/@type[contains(.,'arabicParenR')]">
													<xsl:attribute name ="style:num-format" >
														<xsl:value-of  select ="'1'"/>
													</xsl:attribute >
													<xsl:attribute name ="style:num-suffix">
														<xsl:value-of select ="')'"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test ="a:buAutoNum/@type[contains(.,'arabicParenBoth')]">
													<xsl:attribute name ="style:num-format" >
														<xsl:value-of  select ="'1'"/>
													</xsl:attribute >
													<xsl:attribute name ="style:num-suffix">
														<xsl:value-of select ="')'"/>
													</xsl:attribute>
													<xsl:attribute name ="style:num-prefix">
														<xsl:value-of select ="'('"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test ="a:buAutoNum/@type[contains(.,'alphaUcPeriod')]">
													<xsl:attribute name ="style:num-format" >
														<xsl:value-of  select ="'A'"/>
													</xsl:attribute >
													<xsl:attribute name ="style:num-suffix">
														<xsl:value-of select ="'.'"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test ="a:buAutoNum/@type[contains(.,'alphaUcParenR')]">
													<xsl:attribute name ="style:num-format" >
														<xsl:value-of  select ="'A'"/>
													</xsl:attribute >
													<xsl:attribute name ="style:num-suffix">
														<xsl:value-of select ="')'"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test ="a:buAutoNum/@type[contains(.,'alphaUcParenBoth')]">
													<xsl:attribute name ="style:num-format" >
														<xsl:value-of  select ="'A'"/>
													</xsl:attribute >
													<xsl:attribute name ="style:num-suffix">
														<xsl:value-of select ="')'"/>
													</xsl:attribute>
													<xsl:attribute name ="style:num-prefix">
														<xsl:value-of select ="'('"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test ="a:buAutoNum/@type[contains(.,'alphaLcPeriod')]">
													<xsl:attribute name ="style:num-format" >
														<xsl:value-of  select ="'a'"/>
													</xsl:attribute >
													<xsl:attribute name ="style:num-suffix">
														<xsl:value-of select ="'.'"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test ="a:buAutoNum/@type[contains(.,'alphaLcParenR')]">
													<xsl:attribute name ="style:num-format" >
														<xsl:value-of  select ="'a'"/>
													</xsl:attribute >
													<xsl:attribute name ="style:num-suffix">
														<xsl:value-of select ="')'"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test ="a:buAutoNum/@type[contains(.,'alphaLcParenBoth')]">
													<xsl:attribute name ="style:num-format" >
														<xsl:value-of  select ="'a'"/>
													</xsl:attribute >
													<xsl:attribute name ="style:num-suffix">
														<xsl:value-of select ="')'"/>
													</xsl:attribute>
													<xsl:attribute name ="style:num-prefix">
														<xsl:value-of select ="'('"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test ="a:buAutoNum/@type[contains(.,'romanUcPeriod')]">
													<xsl:attribute name ="style:num-format" >
														<xsl:value-of  select ="'I'"/>
													</xsl:attribute >
													<xsl:attribute name ="style:num-suffix">
														<xsl:value-of select ="'.'"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test ="a:buAutoNum/@type[contains(.,'romanLcPeriod')]">
													<xsl:attribute name ="style:num-format" >
														<xsl:value-of  select ="'i'"/>
													</xsl:attribute >
													<xsl:attribute name ="style:num-suffix">
														<xsl:value-of select ="'.'"/>
													</xsl:attribute>
												</xsl:if >
												<!-- start at value-->
												<xsl:if test ="a:buAutoNum/@startAt">
													<xsl:attribute name ="text:start-value">
														<xsl:value-of select ="a:buAutoNum/@startAt"/>
													</xsl:attribute>
												</xsl:if>
												<style:list-level-properties text:min-label-width="0.952cm" />
												<style:text-properties style:font-family-generic="swiss" style:font-pitch="variable" fo:color="#000000" fo:font-size="100%">
													<xsl:if test ="a:buFont/@typeface">
														<xsl:attribute name ="fo:font-familyr">
															<xsl:value-of select ="a:buFont/@typeface"/>
														</xsl:attribute>
													</xsl:if>
													<xsl:if test ="not(a:buFont/@typeface)">
														<xsl:attribute name ="fo:font-familyr">
															<xsl:value-of select ="'Arial'"/>
														</xsl:attribute>
													</xsl:if>
													<xsl:if test ="a:buSzPct">
														<xsl:attribute name ="fo:font-size">
															<xsl:value-of select ="concat((a:buSzPct/@val div 1000),'%')"/>
														</xsl:attribute>
													</xsl:if>
													<xsl:if test ="not(a:buSzPct)">
														<xsl:attribute name ="fo:font-size">
															<xsl:value-of select ="'100%'"/>
														</xsl:attribute>
													</xsl:if>
													<xsl:if test ="a:buClr">
														<xsl:if test ="a:buClr/a:srgbClr">
															<xsl:variable name ="color" select ="a:buClr/a:srgbClr/@val"/>
															<xsl:attribute name ="fo:color">
																<xsl:value-of select ="concat('#',$color)"/>
															</xsl:attribute>
														</xsl:if>
														<xsl:if test ="a:buClr/a:schemeClr">
															<xsl:call-template name ="getColorCode">
																<xsl:with-param name ="color">
																	<xsl:value-of select="a:buClr/a:schemeClr/@val"/>
																</xsl:with-param>
															</xsl:call-template >
														</xsl:if>
													</xsl:if>
												</style:text-properties>
											</text:list-level-style-number>
										</xsl:if>
										<xsl:if test ="a:buBlip">
											<xsl:variable name ="rId" select ="a:buBlip/a:blip/@r:embed"/>
											<xsl:variable name="XlinkHref">
												<xsl:variable name="pzipsource">
													<xsl:value-of select="document('ppt/slideMasters/_rels/slideMaster1.xml.rels')//rels:Relationships/rels:Relationship[@Id=$rId]/@Target"/>
												</xsl:variable>
												<xsl:value-of select="concat('Pictures/', substring-after($pzipsource,'../media/'))"/>
											</xsl:variable>
											<xsl:call-template name="copyPictures">
												<xsl:with-param name="document">
													<xsl:value-of select="document('ppt/slideMasters/_rels/slideMaster1.xml.rels')"/>
												</xsl:with-param>
												<xsl:with-param name="rId">
													<xsl:value-of select ="$rId"/>
												</xsl:with-param>
											</xsl:call-template>
											<text:list-level-style-image text:level="1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad">
												<xsl:attribute name="xlink:href">
													<xsl:value-of select="$XlinkHref"/>
												</xsl:attribute>
												<style:list-level-properties style:vertical-pos="middle" style:vertical-rel="line" fo:width="0.318cm" fo:height="0.318cm" />
											</text:list-level-style-image>
										</xsl:if>
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</text:list-style>
					</xsl:if>
					<!-- Fix by vijayeta,on 22nd june,bug number1739817,also test if explicit lvl=0-->
					<xsl:if test ="not(a:pPr/@lvl)or a:pPr/@lvl='0' ">
						<text:list-style>
							<xsl:attribute name ="style:name">
								<xsl:value-of select ="concat($listStyleName,position())"/>
							</xsl:attribute >
							<xsl:if test ="a:pPr/a:buChar">
								<text:list-level-style-bullet>
									<xsl:attribute name ="text:level">
										<xsl:value-of select ="1"/>
									</xsl:attribute >
									<xsl:attribute name ="text:bullet-char">
										<xsl:call-template name ="insertBulletCharacter">
											<xsl:with-param name ="character" select ="a:pPr/a:buChar/@char" />
										</xsl:call-template>
									</xsl:attribute >
									<style:list-level-properties text:min-label-width="0.8cm">
										<xsl:attribute name ="text:space-before">
											<xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
												<xsl:value-of select="concat(format-number(child::node()[1]/@marL div 360000,'#.##'),'cm')"/>
											</xsl:for-each>
										</xsl:attribute>
									</style:list-level-properties>
									<style:text-properties style:font-charset="x-symbol" >
										<xsl:if test ="a:pPr/a:buFont/@typeface">
											<xsl:if test ="a:pPr/a:buFont[@typeface='Arial']">
												<xsl:attribute name ="fo:font-family">
													<xsl:value-of select ="'StarSymbol'"/>
												</xsl:attribute>
											</xsl:if>
											<xsl:if test ="not(a:pPr/a:buFont[@typeface='Arial'])">
												<xsl:attribute name ="fo:font-family">
													<xsl:value-of select ="a:pPr/a:buFont/@typeface"/>
												</xsl:attribute>
											</xsl:if>
										</xsl:if>
										<xsl:if test ="not(a:pPr/a:buFont/@typeface)">
											<xsl:attribute name ="fo:font-family">
												<xsl:value-of select ="'StarSymbol'"/>
											</xsl:attribute>
										</xsl:if>
										<xsl:if test ="a:pPr/a:buSzPct">
											<xsl:attribute name ="fo:font-size">
												<xsl:value-of select ="concat((a:pPr/a:buSzPct/@val div 1000),'%')"/>
											</xsl:attribute>
										</xsl:if>
										<xsl:if test ="not(a:pPr/a:buSzPct)">
											<xsl:attribute name ="fo:font-size">
												<xsl:value-of select ="'100%'"/>
											</xsl:attribute>
										</xsl:if>
										<xsl:if test ="a:pPr/a:buClr">
											<xsl:if test ="a:pPr/a:buClr/a:srgbClr">
												<xsl:variable name ="color" select ="a:pPr/a:buClr/a:srgbClr/@val"/>
												<xsl:attribute name ="fo:color">
													<xsl:value-of select ="concat('#',$color)"/>
												</xsl:attribute>
											</xsl:if>
											<xsl:if test ="a:pPr/a:buClr/a:schemeClr">
												<xsl:call-template name ="getColorCode">
													<xsl:with-param name ="color">
														<xsl:value-of select="a:pPr/a:buClr/a:schemeClr/@val"/>
													</xsl:with-param>
												</xsl:call-template >
											</xsl:if>
										</xsl:if>
										<xsl:if test ="not(a:pPr/a:buClr)">
											<xsl:if test ="a:r/a:rPr/a:solidFill">
												<xsl:if test ="a:r/a:rPr/a:solidFill/a:srgbClr">
													<xsl:attribute name ="fo:color">
														<xsl:value-of select ="concat('#',a:r/a:rPr/a:solidFill/a:srgbClr/@val)"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test ="a:r/a:rPr/a:solidFill/a:schemeClr">
													<xsl:attribute name ="fo:color">
														<xsl:call-template name ="getColorCode">
															<xsl:with-param name ="color">
																<xsl:value-of select="a:r/a:rPr/a:solidFill/a:schemeClr/@val"/>
															</xsl:with-param>
														</xsl:call-template >
													</xsl:attribute>
												</xsl:if>
											</xsl:if>
											<xsl:if test ="not(a:r/a:rPr/a:solidFill)">
												<!-- Code added by vijayeta, bug fix 1746350-->
												<xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
													<xsl:if test ="a:lvl1pPr/a:buClr/a:srgbClr">
														<xsl:attribute name ="fo:color">
															<xsl:value-of select ="a:lvl1pPr/a:buClr/a:srgbClr/@val"/>
														</xsl:attribute>
													</xsl:if>
													<xsl:if test ="a:lvl1pPr/a:buClr/a:schemeClr">
														<xsl:attribute name ="fo:color">
															<xsl:call-template name ="getColorCode">
																<xsl:with-param name ="color">
																	<xsl:value-of select="a:lvl1pPr/a:buClr/a:schemeClr/@val"/>
																</xsl:with-param>
															</xsl:call-template >
														</xsl:attribute>
													</xsl:if>
												</xsl:for-each>
											</xsl:if>
										</xsl:if>
										<!--End of Code added by vijayeta, bug fix 1746350-->
									</style:text-properties >
								</text:list-level-style-bullet>
							</xsl:if>
							<xsl:if test ="a:pPr/a:buAutoNum">
								<text:list-level-style-number>
									<xsl:attribute name ="text:level">
										<xsl:value-of select ="'1'"/>
									</xsl:attribute >
									<xsl:variable name ="startAt">
										<xsl:if test ="a:pPr/a:buAutoNum/@startAt">
											<xsl:value-of select ="a:pPr/a:buAutoNum/@startAt" />
										</xsl:if>
										<xsl:if test ="not(a:pPr/a:buAutoNum/@startAt)">
											<xsl:value-of select ="'1'" />
										</xsl:if>
									</xsl:variable>
									<xsl:call-template name ="insertNumber">
										<xsl:with-param name ="number" select ="a:pPr/a:buAutoNum/@type"/>
										<xsl:with-param name ="startAt" select ="$startAt"/>
									</xsl:call-template>
									<style:list-level-properties text:min-label-width="0.952cm" />
									<style:text-properties style:font-family-generic="swiss" style:font-pitch="variable">
										<xsl:if test ="a:pPr/a:buFont/@typeface">
											<xsl:attribute name ="fo:font-family">
												<xsl:value-of select ="a:pPr/a:buFont/@typeface"/>
											</xsl:attribute>
										</xsl:if>
										<xsl:if test ="not(a:pPr/a:buFont/@typeface)">
											<xsl:attribute name ="fo:font-family">
												<xsl:value-of select ="'Arial'"/>
											</xsl:attribute>
										</xsl:if>
										<xsl:if test ="a:pPr/a:buSzPct">
											<xsl:attribute name ="fo:font-size">
												<xsl:value-of select ="concat((a:pPr/a:buSzPct/@val div 1000),'%')"/>
											</xsl:attribute>
										</xsl:if>
										<xsl:if test ="not(a:pPr/a:buSzPct)">
											<xsl:attribute name ="fo:font-size">
												<xsl:value-of select ="'100%'"/>
											</xsl:attribute>
										</xsl:if>
										<xsl:if test ="a:pPr/a:buClr">
											<xsl:if test ="a:p/a:pPr/a:buClr/a:srgbClr">
												<xsl:attribute name ="fo:color">
													<xsl:value-of select ="concat('#',a:p/a:pPr/a:buClr/a:srgbClr/@val)"/>
												</xsl:attribute>
											</xsl:if>
											<xsl:if test ="a:pPr/a:buClr/a:schemeClr">
												<xsl:call-template name ="getColorCode">
													<xsl:with-param name ="color">
														<xsl:value-of select="a:pPr/a:buClr/a:schemeClr/@val"/>
													</xsl:with-param>
												</xsl:call-template >
											</xsl:if>
										</xsl:if>
										<!-- Code added by vijayeta, bug fix 1746350-->
										<xsl:if test ="not(a:pPr/a:buClr)">
											<xsl:if test ="a:r/a:rPr/a:solidFill">
												<xsl:if test ="a:r/a:rPr/a:solidFill/a:srgbClr">
													<xsl:attribute name ="fo:color">
														<xsl:value-of select ="concat('#',a:r/a:rPr/a:solidFill/a:srgbClr/@val)"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test ="a:r/a:rPr/a:solidFill/a:schemeClr">
													<xsl:attribute name ="fo:color">
														<xsl:call-template name ="getColorCode">
															<xsl:with-param name ="color">
																<xsl:value-of select="a:r/a:rPr/a:solidFill/a:schemeClr/@val"/>
															</xsl:with-param>
														</xsl:call-template >
													</xsl:attribute>
												</xsl:if>
											</xsl:if>
											<xsl:if test ="not(a:r/a:rPr/a:solidFill)">
												<xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
													<xsl:if test ="a:lvl1pPr/a:buClr/a:srgbClr">
														<xsl:attribute name ="fo:color">
															<xsl:value-of select ="a:lvl1pPr/a:buClr/a:srgbClr/@val"/>
														</xsl:attribute>
													</xsl:if>
													<xsl:if test ="a:lvl1pPr/a:buClr/a:schemeClr">
														<xsl:attribute name ="fo:color">
															<xsl:call-template name ="getColorCode">
																<xsl:with-param name ="color">
																	<xsl:value-of select="a:lvl1pPr/a:buClr/a:schemeClr/@val"/>
																</xsl:with-param>
															</xsl:call-template >
														</xsl:attribute>
													</xsl:if>
												</xsl:for-each>
											</xsl:if>
										</xsl:if >
										<!--End of Code added by vijayeta, bug fix 1746350-->
									</style:text-properties>
								</text:list-level-style-number>
							</xsl:if>
							<xsl:if test ="a:pPr/a:buBlip">
								<xsl:variable name ="rId" select ="a:pPr/a:buBlip/a:blip/@r:embed"/>
								<xsl:variable name="XlinkHref">
									<xsl:variable name="pzipsource">
										<xsl:value-of select="document($slideRel)//rels:Relationships/rels:Relationship[@Id=$rId]/@Target"/>
									</xsl:variable>
									<xsl:value-of select="concat('Pictures/', substring-after($pzipsource,'../media/'))"/>
								</xsl:variable>
								<xsl:call-template name="copyPictures">
									<xsl:with-param name="document">
										<xsl:value-of select="$slideRel"/>
									</xsl:with-param>
									<xsl:with-param name="rId">
										<xsl:value-of select ="$rId"/>
									</xsl:with-param>
								</xsl:call-template>
								<text:list-style>
									<xsl:attribute name ="style:name">
										<xsl:value-of select ="$listStyleName"/>
									</xsl:attribute >
									<text:list-level-style-image text:level="1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad">
										<xsl:attribute name="xlink:href">
											<xsl:value-of select="$XlinkHref"/>
										</xsl:attribute>
										<style:list-level-properties style:vertical-pos="middle" style:vertical-rel="line" fo:width="0.318cm" fo:height="0.318cm" />
									</text:list-level-style-image>
								</text:list-style>
							</xsl:if>
						</text:list-style>
					</xsl:if>
					<!-- Addded by vijayeta on 18th june,to check for bunone-->
				</xsl:if>
				<!-- Addded by vijayeta on 18th june,to check for bunone-->
			</xsl:for-each>
		</xsl:if>
		<!--End of Condition if levels are present-->
		<!--Condition if levels are not present-->
		<!-- Fix by vijayeta,on 22nd june,bug number1739817,also test if explicit lvl=0-->
		<xsl:if test ="not(./a:p/a:pPr/@lvl) or ./a:p/a:pPr/@lvl='0'">
			<xsl:for-each select ="./a:p">
				<!-- Addded by vijayeta on 18th june,to check for bunone-->
				<xsl:if test ="not(a:pPr/a:buNone)">
					<!-- Addded by vijayeta on 18th june,to check for bunone-->
					<xsl:if test ="a:pPr/a:buChar">
						<text:list-style>
							<xsl:attribute name ="style:name">
								<xsl:value-of select ="concat($listStyleName,position())"/>
							</xsl:attribute >
							<text:list-level-style-bullet>
								<xsl:attribute name ="text:level">
									<xsl:value-of select ="1"/>
								</xsl:attribute >
								<xsl:attribute name ="text:bullet-char">
									<xsl:call-template name ="insertBulletCharacter">
										<xsl:with-param name ="character" select ="a:pPr/a:buChar/@char" />
									</xsl:call-template>
								</xsl:attribute >
								<style:list-level-properties text:min-label-width="0.952cm" />
								<style:text-properties style:font-charset="x-symbol">
									<xsl:if test ="a:pPr/a:buFont/@typeface">
										<xsl:if test ="a:pPr/a:buFont[@typeface='Arial']">
											<xsl:attribute name ="fo:font-family">
												<xsl:value-of select ="'StarSymbol'"/>
											</xsl:attribute>
										</xsl:if>
										<xsl:if test ="not(a:pPr/a:buFont[@typeface='Arial'])">
											<xsl:attribute name ="fo:font-family">
												<xsl:value-of select ="a:pPr/a:buFont/@typeface"/>
											</xsl:attribute>
										</xsl:if>
									</xsl:if>
									<xsl:if test ="not(a:pPr/a:buFont/@typeface)">
										<xsl:attribute name ="fo:font-family">
											<xsl:value-of select ="'StarSymbol'"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="a:pPr/a:buSzPct">
										<xsl:attribute name ="fo:font-size">
											<xsl:value-of select ="concat((a:pPr/a:buSzPct/@val div 1000),'%')"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="not(a:pPr/a:buSzPct)">
										<xsl:attribute name ="fo:font-size">
											<xsl:value-of select ="'100%'"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="a:p/a:pPr/a:buClr">
										<xsl:if test ="a:pPr/a:buClr/a:srgbClr">
											<xsl:variable name ="color" select ="/a:pPr/a:buClr/a:srgbClr/@val"/>
											<xsl:attribute name ="fo:color">
												<xsl:value-of select ="concat('#',$color)"/>
											</xsl:attribute>
										</xsl:if>
										<xsl:if test ="a:pPr/a:buClr/a:schemeClr">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="a:pPr/a:buClr/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:if>
									</xsl:if>
									<!-- Code added by vijayeta, bug fix 1746350-->

									<xsl:if test ="not(a:pPr/a:buClr)">
										<xsl:if test ="a:r/a:rPr/a:solidFill">
											<xsl:if test ="a:r/a:rPr/a:solidFill/a:srgbClr">
												<xsl:attribute name ="fo:color">
													<xsl:value-of select ="concat('#',a:r/a:rPr/a:solidFill/a:srgbClr/@val)"/>
												</xsl:attribute>
											</xsl:if>
											<xsl:if test ="a:r/a:rPr/a:solidFill/a:schemeClr">
												<xsl:attribute name ="fo:color">
													<xsl:call-template name ="getColorCode">
														<xsl:with-param name ="color">
															<xsl:value-of select="a:r/a:rPr/a:solidFill/a:schemeClr/@val"/>
														</xsl:with-param>
													</xsl:call-template >
												</xsl:attribute>
											</xsl:if>
										</xsl:if>
										<xsl:if test ="not(a:r/a:rPr/a:solidFill)">
											<xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
												<xsl:if test ="a:lvl1pPr/a:buClr/a:srgbClr">
													<xsl:attribute name ="fo:color">
														<xsl:value-of select ="a:lvl1pPr/a:buClr/a:srgbClr/@val"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test ="a:lvl1pPr/a:buClr/a:schemeClr">
													<xsl:attribute name ="fo:color">
														<xsl:call-template name ="getColorCode">
															<xsl:with-param name ="color">
																<xsl:value-of select="a:lvl1pPr/a:buClr/a:schemeClr/@val"/>
															</xsl:with-param>
														</xsl:call-template >
													</xsl:attribute>
												</xsl:if>
											</xsl:for-each>
										</xsl:if>
									</xsl:if>
									<!--End of Code added by vijayeta, bug fix 1746350-->
								</style:text-properties >
							</text:list-level-style-bullet>
						</text:list-style>
					</xsl:if>
					<xsl:if test ="a:pPr/a:buAutoNum">
						<text:list-style>
							<xsl:attribute name ="style:name">
								<xsl:value-of select ="concat($listStyleName,position())"/>
							</xsl:attribute >
							<text:list-level-style-number>
								<xsl:attribute name ="text:level">
									<xsl:value-of select ="'1'"/>
								</xsl:attribute >
								<xsl:variable name ="startAt">
									<xsl:if test ="a:pPr/a:buAutoNum/@startAt">
										<xsl:value-of select ="a:pPr/a:buAutoNum/@startAt" />
									</xsl:if>
									<xsl:if test ="not(a:pPr/a:buAutoNum/@startAt)">
										<xsl:value-of select ="'1'" />
									</xsl:if>
								</xsl:variable>
								<xsl:call-template name ="insertNumber">
									<xsl:with-param name ="number" select ="a:pPr/a:buAutoNum/@type"/>
									<xsl:with-param name ="startAt" select ="$startAt"/>
								</xsl:call-template>
								<style:list-level-properties text:min-label-width="0.952cm" />
								<style:text-properties style:font-family-generic="swiss" style:font-pitch="variable">
									<xsl:if test ="a:pPr/a:buFont/@typeface">
										<xsl:attribute name ="fo:font-family">
											<xsl:value-of select ="a:pPr/a:buFont/@typeface"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="not(a:pPr/a:buFont/@typeface)">
										<xsl:attribute name ="fo:font-family">
											<xsl:value-of select ="'Arial'"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="a:pPr/a:buSzPct">
										<xsl:attribute name ="fo:font-size">
											<xsl:value-of select ="concat((/a:pPr/a:buSzPct/@val div 1000),'%')"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="not(a:pPr/a:buSzPct)">
										<xsl:attribute name ="fo:font-size">
											<xsl:value-of select ="'100%'"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="a:pPr/a:buClr">
										<xsl:if test ="a:pPr/a:buClr/a:srgbClr">
											<xsl:attribute name ="fo:color">
												<xsl:value-of select ="concat('#',a:pPr/a:buClr/a:srgbClr/@val)"/>
											</xsl:attribute>
										</xsl:if>
										<xsl:if test ="a:pPr/a:buClr/a:schemeClr">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="a:pPr/a:buClr/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:if>
									</xsl:if>
									<!-- Code added by vijayeta, bug fix 1746350-->
									<xsl:if test ="not(a:pPr/a:buClr)">
										<xsl:if test ="a:r/a:rPr/a:solidFill">
											<xsl:if test ="a:r/a:rPr/a:solidFill/a:srgbClr">
												<xsl:attribute name ="fo:color">
													<xsl:value-of select ="concat('#',a:r/a:rPr/a:solidFill/a:srgbClr/@val)"/>
												</xsl:attribute>
											</xsl:if>
											<xsl:if test ="a:r/a:rPr/a:solidFill/a:schemeClr">
												<xsl:attribute name ="fo:color">
													<xsl:call-template name ="getColorCode">
														<xsl:with-param name ="color">
															<xsl:value-of select="a:r/a:rPr/a:solidFill/a:schemeClr/@val"/>
														</xsl:with-param>
													</xsl:call-template >
												</xsl:attribute>
											</xsl:if>
										</xsl:if>
										<xsl:if test ="not(a:r/a:rPr/a:solidFill)">
											<xsl:variable name ="isColorInLayoutMaster">
												<xsl:if test ="(document(concat('ppt/slideLayouts/',$layoutName)))//p:sldLayout/p:cSld/p:spTree/p:sp/p:txBody/a:lstStyle/a:lvl1pPr/a:spcBef/a:spcPts">
													<xsl:value-of select ="'false'"/>
												</xsl:if>
												<xsl:if test ="not(document(concat('ppt/slideLayouts/',$layoutName))//p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@type=$layoutName]/parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:spcBef/a:spcPts)">
													<xsl:value-of select ="'true'"/>
												</xsl:if>
											</xsl:variable >
											<xsl:if test ="$isColorInLayoutMaster='false'">
											</xsl:if>
											<xsl:if test ="$isColorInLayoutMaster='true'">
												<xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
													<xsl:if test ="a:lvl1pPr/a:buClr/a:srgbClr">
														<xsl:attribute name ="fo:color">
															<xsl:value-of select ="a:lvl1pPr/a:buClr/a:srgbClr/@val"/>
														</xsl:attribute>
													</xsl:if>
													<xsl:if test ="a:lvl1pPr/a:buClr/a:schemeClr">
														<xsl:attribute name ="fo:color">
															<xsl:call-template name ="getColorCode">
																<xsl:with-param name ="color">
																	<xsl:value-of select="a:lvl1pPr/a:buClr/a:schemeClr/@val"/>
																</xsl:with-param>
															</xsl:call-template >
														</xsl:attribute>
													</xsl:if>
												</xsl:for-each>
											</xsl:if>
										</xsl:if>
									</xsl:if >
									<!--End of Code added by vijayeta, bug fix 1746350-->
								</style:text-properties>
							</text:list-level-style-number>
						</text:list-style>
					</xsl:if>
					<xsl:if test ="a:pPr/a:buBlip">
						<xsl:variable name ="rId" select ="./a:p/a:pPr/a:buBlip/a:blip/@r:embed"/>
						<xsl:variable name="XlinkHref">
							<xsl:variable name="pzipsource">
								<xsl:value-of select="document($slideRel)//rels:Relationships/rels:Relationship[@Id=$rId]/@Target"/>
							</xsl:variable>
							<xsl:value-of select="concat('Pictures/', substring-after($pzipsource,'../media/'))"/>
						</xsl:variable>
						<xsl:call-template name="copyPictures">
							<xsl:with-param name="document">
								<xsl:value-of select="$slideRel"/>
							</xsl:with-param>
							<xsl:with-param name="rId">
								<xsl:value-of select ="$rId"/>
							</xsl:with-param>
						</xsl:call-template>
						<text:list-style>
							<xsl:attribute name ="style:name">
								<xsl:value-of select ="$listStyleName"/>
							</xsl:attribute >
							<text:list-level-style-image text:level="1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad">
								<xsl:attribute name="xlink:href">
									<xsl:value-of select="$XlinkHref"/>
								</xsl:attribute>
								<style:list-level-properties style:vertical-pos="middle" style:vertical-rel="line" fo:width="0.318cm" fo:height="0.318cm" />
							</text:list-level-style-image>
						</text:list-style>
					</xsl:if>
					<!-- Addded by vijayeta on 18th june,to check for bunone-->
				</xsl:if >
				<!-- Addded by vijayeta on 18th june,to check for bunone-->
			</xsl:for-each>
		</xsl:if>
		<!--End  of Condition if levels are not present-->
	</xsl:template>
	<xsl:template name ="insertDefaultBulletNumberStyle">
		<xsl:param name ="listStyleName"/>
		<xsl:param name ="slideLayout" />
		<xsl:param name ="slideMaster" />
		<xsl:param name ="var_TextBoxType"/>
		<xsl:param name ="var_index"/>
		<xsl:if test ="./a:p/a:pPr/@lvl">
			<xsl:for-each select ="./a:p">
				<xsl:variable name ="position" select ="position()"/>
				<xsl:if test ="not(a:pPr/a:buNone)">
					<!--Condition if levels are present,and bullets are default-->
					<xsl:if test ="./a:pPr/@lvl and ./a:pPr/@lvl !='0'">
						<xsl:variable name ="textLevel" select ="a:pPr/@lvl"/>
						<!--<xsl:if test ="not(./a:pPr/a:buNone)">-->
						<xsl:variable name ="levelStyle">
							<xsl:value-of select ="concat($listStyleName,$position,'lvl',$textLevel)"/>
						</xsl:variable>
						<xsl:variable name ="newTextLvl" select ="$textLevel+1"/>
						<!-- Check if layout define bullets-->
						<xsl:for-each select ="document(concat('ppt/slideLayouts/',$slideLayout))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
							<xsl:if test="not(./@type) and ./@idx=$var_index">
								<xsl:call-template name ="getBulletForLevelsLayout">
									<xsl:with-param name ="slideMaster" select ="$slideMaster"/>
									<xsl:with-param name ="newTextLvl" select ="$newTextLvl"/>
									<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
								</xsl:call-template>
							</xsl:if>
						</xsl:for-each>
						<!-- Check if layout define bullets-->
						<!--</xsl:if>-->
					</xsl:if>
					<!-- Fix by vijayeta,on 22nd june,bug number1739817,also test if explicit lvl=0-->
					<xsl:if test ="not(./a:pPr/@lvl) or ./a:pPr/@lvl='0'">
						<xsl:if test ="not(./a:pPr/a:buNone)">
							<xsl:variable name ="levelStyle" select =" concat($listStyleName,$position)"/>
							<xsl:for-each select ="document(concat('ppt/slideLayouts/',$slideLayout))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
								<xsl:if test="not(./@type) and ./@idx=$var_index">
									<xsl:call-template name ="getBulletForLevelsLayout">
										<xsl:with-param name ="slideMaster" select ="$slideMaster"/>
										<xsl:with-param name ="newTextLvl" select ="'1'"/>
										<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
									</xsl:call-template>
								</xsl:if>
							</xsl:for-each>
						</xsl:if>
					</xsl:if>
				</xsl:if>
			</xsl:for-each>
		</xsl:if>
		<!--Condition if levels are present,and bullets are default-->
		<!--<xsl:if test ="not(./a:pPr/@lvl)">-->
		<!-- Fix by vijayeta,on 22nd june,bug number1739817,also test if explicit lvl=0-->
		<xsl:if test ="not(./a:p/a:pPr/@lvl) or ./a:p/a:pPr/@lvl='0'">
			<xsl:if test ="not(./a:p/a:pPr/a:buNone)">
				<xsl:for-each select ="./a:p">
					<xsl:variable name ="position" select ="position()"/>
					<xsl:variable name ="levelStyle" select =" concat($listStyleName,$position)"/>
					<xsl:for-each select ="document(concat('ppt/slideLayouts/',$slideLayout))/p:sldLayout/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph">
						<xsl:if test="not(./@type) and ./@idx=$var_index">
							<xsl:call-template name ="getBulletForLevelsLayout">
								<xsl:with-param name ="slideMaster" select ="$slideMaster"/>
								<xsl:with-param name ="newTextLvl" select ="'1'"/>
								<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
							</xsl:call-template>
						</xsl:if>
					</xsl:for-each>
				</xsl:for-each>
			</xsl:if>
		</xsl:if >
		<!--</xsl:if>-->

	</xsl:template>
	<!-- Code By vijayeta,insertMultipleLevels in OdpFiles-->
	<xsl:template name ="insertMultipleLevels">
		<xsl:param name ="levelCount"/>
		<xsl:param name ="ParaId"/>
		<xsl:param name ="listStyleName"/>
		<xsl:choose>
			<xsl:when test ="$levelCount='1'">
				<text:list>
					<xsl:attribute name ="text:style-name">
						<xsl:value-of select ="concat($listStyleName,'lvl',$levelCount)"/>
					</xsl:attribute>
					<text:list-item>
						<text:list>
							<text:list-item>
								<text:p >
									<xsl:attribute name ="text:style-name">
										<xsl:value-of select ="concat($ParaId,position())"/>
									</xsl:attribute>
									<xsl:for-each select ="node()">
										<xsl:if test ="name()='a:r'">
											<text:span text:style-name="{generate-id()}">
												<!--<xsl:value-of select ="a:t"/>-->
												<!--converts whitespaces sequence to text:s-->
												<!-- 1699083 bug fix  -->
												<xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
												<xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
												<xsl:choose >
													<xsl:when test ="a:rPr[@cap='all']">
														<xsl:choose >
															<xsl:when test =".=''">
																<text:s/>
															</xsl:when>
															<xsl:when test ="not(contains(.,'  '))">
																<xsl:value-of select ="translate(.,$lcletters,$ucletters)"/>
															</xsl:when>
															<xsl:when test =". =' '">
																<text:s/>
															</xsl:when>
															<xsl:otherwise >
																<xsl:call-template name ="InsertWhiteSpaces">
																	<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
																</xsl:call-template>
															</xsl:otherwise>
														</xsl:choose>
													</xsl:when>
													<xsl:when test ="a:rPr[@cap='small']">
														<xsl:choose >
															<xsl:when test =".=''">
																<text:s/>
															</xsl:when>
															<xsl:when test ="not(contains(.,'  '))">
																<xsl:value-of select ="translate(.,$ucletters,$lcletters)"/>
															</xsl:when>
															<xsl:when test =".= ' '">
																<text:s/>
															</xsl:when>
															<xsl:otherwise >
																<xsl:call-template name ="InsertWhiteSpaces">
																	<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
																</xsl:call-template>
															</xsl:otherwise>
														</xsl:choose >
													</xsl:when>
													<xsl:otherwise >
														<xsl:choose >
															<xsl:when test =".=''">
																<text:s/>
															</xsl:when>
															<xsl:when test ="not(contains(.,'  '))">
																<xsl:value-of select ="."/>
															</xsl:when>
															<xsl:otherwise >
																<xsl:call-template name ="InsertWhiteSpaces">
																	<xsl:with-param name ="string" select ="."/>
																</xsl:call-template>
															</xsl:otherwise >
														</xsl:choose>
													</xsl:otherwise>
												</xsl:choose>
											</text:span>
										</xsl:if >
										<xsl:if test ="name()='a:br'">
											<text:line-break/>
										</xsl:if>
									</xsl:for-each>
								</text:p>
							</text:list-item>
						</text:list>
					</text:list-item>
				</text:list>
			</xsl:when>
			<xsl:when test ="$levelCount='2'">
				<text:list>
					<xsl:attribute name ="text:style-name">
						<xsl:value-of select ="concat($listStyleName,'lvl',$levelCount)"/>
					</xsl:attribute>
					<text:list-item>
						<text:list>
							<text:list-item>
								<text:list>
									<text:list-item>
										<text:p >
											<xsl:attribute name ="text:style-name">
												<xsl:value-of select ="concat($ParaId,position())"/>
											</xsl:attribute>
											<xsl:for-each select ="node()">
												<xsl:if test ="name()='a:r'">
													<text:span text:style-name="{generate-id()}">
														<!--<xsl:value-of select ="a:t"/>-->
														<!--converts whitespaces sequence to text:s-->
														<!-- 1699083 bug fix  -->
														<xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
														<xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
														<xsl:choose >
															<xsl:when test ="a:rPr[@cap='all']">
																<xsl:choose >
																	<xsl:when test =".=''">
																		<text:s/>
																	</xsl:when>
																	<xsl:when test ="not(contains(.,'  '))">
																		<xsl:value-of select ="translate(.,$lcletters,$ucletters)"/>
																	</xsl:when>
																	<xsl:when test =". =' '">
																		<text:s/>
																	</xsl:when>
																	<xsl:otherwise >
																		<xsl:call-template name ="InsertWhiteSpaces">
																			<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
																		</xsl:call-template>
																	</xsl:otherwise>
																</xsl:choose>
															</xsl:when>
															<xsl:when test ="a:rPr[@cap='small']">
																<xsl:choose >
																	<xsl:when test =".=''">
																		<text:s/>
																	</xsl:when>
																	<xsl:when test ="not(contains(.,'  '))">
																		<xsl:value-of select ="translate(.,$ucletters,$lcletters)"/>
																	</xsl:when>
																	<xsl:when test =".= ' '">
																		<text:s/>
																	</xsl:when>
																	<xsl:otherwise >
																		<xsl:call-template name ="InsertWhiteSpaces">
																			<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
																		</xsl:call-template>
																	</xsl:otherwise>
																</xsl:choose >
															</xsl:when>
															<xsl:otherwise >
																<xsl:choose >
																	<xsl:when test =".=''">
																		<text:s/>
																	</xsl:when>
																	<xsl:when test ="not(contains(.,'  '))">
																		<xsl:value-of select ="."/>
																	</xsl:when>
																	<xsl:otherwise >
																		<xsl:call-template name ="InsertWhiteSpaces">
																			<xsl:with-param name ="string" select ="."/>
																		</xsl:call-template>
																	</xsl:otherwise >
																</xsl:choose>
															</xsl:otherwise>
														</xsl:choose>
													</text:span>
												</xsl:if >
												<xsl:if test ="name()='a:br'">
													<text:line-break/>
												</xsl:if>
											</xsl:for-each>
										</text:p>
									</text:list-item>
								</text:list>
							</text:list-item>
						</text:list>
					</text:list-item>
				</text:list>
			</xsl:when>
			<xsl:when test ="$levelCount='3'">
				<text:list>
					<xsl:attribute name ="text:style-name">
						<xsl:value-of select ="concat($listStyleName,'lvl',$levelCount)"/>
					</xsl:attribute>
					<text:list-item>
						<text:list>
							<text:list-item>
								<text:list>
									<text:list-item>
										<text:list>
											<text:list-item>
												<text:p >
													<xsl:attribute name ="text:style-name">
														<xsl:value-of select ="concat($ParaId,position())"/>
													</xsl:attribute>
													<xsl:for-each select ="node()">
														<xsl:if test ="name()='a:r'">
															<text:span text:style-name="{generate-id()}">
																<!--<xsl:value-of select ="a:t"/>-->
																<!--converts whitespaces sequence to text:s-->
																<!-- 1699083 bug fix  -->
																<xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
																<xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
																<xsl:choose >
																	<xsl:when test ="a:rPr[@cap='all']">
																		<xsl:choose >
																			<xsl:when test =".=''">
																				<text:s/>
																			</xsl:when>
																			<xsl:when test ="not(contains(.,'  '))">
																				<xsl:value-of select ="translate(.,$lcletters,$ucletters)"/>
																			</xsl:when>
																			<xsl:when test =". =' '">
																				<text:s/>
																			</xsl:when>
																			<xsl:otherwise >
																				<xsl:call-template name ="InsertWhiteSpaces">
																					<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
																				</xsl:call-template>
																			</xsl:otherwise>
																		</xsl:choose>
																	</xsl:when>
																	<xsl:when test ="a:rPr[@cap='small']">
																		<xsl:choose >
																			<xsl:when test =".=''">
																				<text:s/>
																			</xsl:when>
																			<xsl:when test ="not(contains(.,'  '))">
																				<xsl:value-of select ="translate(.,$ucletters,$lcletters)"/>
																			</xsl:when>
																			<xsl:when test =".= ' '">
																				<text:s/>
																			</xsl:when>
																			<xsl:otherwise >
																				<xsl:call-template name ="InsertWhiteSpaces">
																					<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
																				</xsl:call-template>
																			</xsl:otherwise>
																		</xsl:choose >
																	</xsl:when>
																	<xsl:otherwise >
																		<xsl:choose >
																			<xsl:when test =".=''">
																				<text:s/>
																			</xsl:when>
																			<xsl:when test ="not(contains(.,'  '))">
																				<xsl:value-of select ="."/>
																			</xsl:when>
																			<xsl:otherwise >
																				<xsl:call-template name ="InsertWhiteSpaces">
																					<xsl:with-param name ="string" select ="."/>
																				</xsl:call-template>
																			</xsl:otherwise >
																		</xsl:choose>
																	</xsl:otherwise>
																</xsl:choose>
															</text:span>
														</xsl:if >
														<xsl:if test ="name()='a:br'">
															<text:line-break/>
														</xsl:if>
													</xsl:for-each>
												</text:p>
											</text:list-item>
										</text:list>
									</text:list-item>
								</text:list>
							</text:list-item>
						</text:list>
					</text:list-item>
				</text:list>
			</xsl:when>
			<xsl:when test ="$levelCount='4'">
				<text:list>
					<xsl:attribute name ="text:style-name">
						<xsl:value-of select ="concat($listStyleName,'lvl',$levelCount)"/>
					</xsl:attribute>
					<text:list-item>
						<text:list>
							<text:list-item>
								<text:list>
									<text:list-item>
										<text:list>
											<text:list-item>
												<text:list>
													<text:list-item>
														<text:p >
															<xsl:attribute name ="text:style-name">
																<xsl:value-of select ="concat($ParaId,position())"/>
															</xsl:attribute>
															<xsl:for-each select ="node()">
																<xsl:if test ="name()='a:r'">
																	<text:span text:style-name="{generate-id()}">
																		<!--<xsl:value-of select ="a:t"/>-->
																		<!--converts whitespaces sequence to text:s-->
																		<!-- 1699083 bug fix  -->
																		<xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
																		<xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
																		<xsl:choose >
																			<xsl:when test ="a:rPr[@cap='all']">
																				<xsl:choose >
																					<xsl:when test =".=''">
																						<text:s/>
																					</xsl:when>
																					<xsl:when test ="not(contains(.,'  '))">
																						<xsl:value-of select ="translate(.,$lcletters,$ucletters)"/>
																					</xsl:when>
																					<xsl:when test =". =' '">
																						<text:s/>
																					</xsl:when>
																					<xsl:otherwise >
																						<xsl:call-template name ="InsertWhiteSpaces">
																							<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
																						</xsl:call-template>
																					</xsl:otherwise>
																				</xsl:choose>
																			</xsl:when>
																			<xsl:when test ="a:rPr[@cap='small']">
																				<xsl:choose >
																					<xsl:when test =".=''">
																						<text:s/>
																					</xsl:when>
																					<xsl:when test ="not(contains(.,'  '))">
																						<xsl:value-of select ="translate(.,$ucletters,$lcletters)"/>
																					</xsl:when>
																					<xsl:when test =".= ' '">
																						<text:s/>
																					</xsl:when>
																					<xsl:otherwise >
																						<xsl:call-template name ="InsertWhiteSpaces">
																							<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
																						</xsl:call-template>
																					</xsl:otherwise>
																				</xsl:choose >
																			</xsl:when>
																			<xsl:otherwise >
																				<xsl:choose >
																					<xsl:when test =".=''">
																						<text:s/>
																					</xsl:when>
																					<xsl:when test ="not(contains(.,'  '))">
																						<xsl:value-of select ="."/>
																					</xsl:when>
																					<xsl:otherwise >
																						<xsl:call-template name ="InsertWhiteSpaces">
																							<xsl:with-param name ="string" select ="."/>
																						</xsl:call-template>
																					</xsl:otherwise >
																				</xsl:choose>
																			</xsl:otherwise>
																		</xsl:choose>
																	</text:span>
																</xsl:if >
																<xsl:if test ="name()='a:br'">
																	<text:line-break/>
																</xsl:if>
															</xsl:for-each>
														</text:p>
													</text:list-item>
												</text:list>
											</text:list-item>
										</text:list>
									</text:list-item>
								</text:list>
							</text:list-item>
						</text:list>
					</text:list-item>
				</text:list>
			</xsl:when>
			<xsl:when test ="$levelCount='5'">
				<text:list>
					<xsl:attribute name ="text:style-name">
						<xsl:value-of select ="concat($listStyleName,'lvl',$levelCount)"/>
					</xsl:attribute>
					<text:list-item>
						<text:list>
							<text:list-item>
								<text:list>
									<text:list-item>
										<text:list>
											<text:list-item>
												<text:list>
													<text:list-item>
														<text:list>
															<text:list-item>
																<text:p >
																	<xsl:attribute name ="text:style-name">
																		<xsl:value-of select ="concat($ParaId,position())"/>
																	</xsl:attribute>
																	<xsl:for-each select ="node()">
																		<xsl:if test ="name()='a:r'">
																			<text:span text:style-name="{generate-id()}">
																				<!--<xsl:value-of select ="a:t"/>-->
																				<!--converts whitespaces sequence to text:s-->
																				<!-- 1699083 bug fix  -->
																				<xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
																				<xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
																				<xsl:choose >
																					<xsl:when test ="a:rPr[@cap='all']">
																						<xsl:choose >
																							<xsl:when test =".=''">
																								<text:s/>
																							</xsl:when>
																							<xsl:when test ="not(contains(.,'  '))">
																								<xsl:value-of select ="translate(.,$lcletters,$ucletters)"/>
																							</xsl:when>
																							<xsl:when test =". =' '">
																								<text:s/>
																							</xsl:when>
																							<xsl:otherwise >
																								<xsl:call-template name ="InsertWhiteSpaces">
																									<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
																								</xsl:call-template>
																							</xsl:otherwise>
																						</xsl:choose>
																					</xsl:when>
																					<xsl:when test ="a:rPr[@cap='small']">
																						<xsl:choose >
																							<xsl:when test =".=''">
																								<text:s/>
																							</xsl:when>
																							<xsl:when test ="not(contains(.,'  '))">
																								<xsl:value-of select ="translate(.,$ucletters,$lcletters)"/>
																							</xsl:when>
																							<xsl:when test =".= ' '">
																								<text:s/>
																							</xsl:when>
																							<xsl:otherwise >
																								<xsl:call-template name ="InsertWhiteSpaces">
																									<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
																								</xsl:call-template>
																							</xsl:otherwise>
																						</xsl:choose >
																					</xsl:when>
																					<xsl:otherwise >
																						<xsl:choose >
																							<xsl:when test =".=''">
																								<text:s/>
																							</xsl:when>
																							<xsl:when test ="not(contains(.,'  '))">
																								<xsl:value-of select ="."/>
																							</xsl:when>
																							<xsl:otherwise >
																								<xsl:call-template name ="InsertWhiteSpaces">
																									<xsl:with-param name ="string" select ="."/>
																								</xsl:call-template>
																							</xsl:otherwise >
																						</xsl:choose>
																					</xsl:otherwise>
																				</xsl:choose>
																			</text:span>
																		</xsl:if >
																		<xsl:if test ="name()='a:br'">
																			<text:line-break/>
																		</xsl:if>
																	</xsl:for-each>
																</text:p>
															</text:list-item>
														</text:list>
													</text:list-item>
												</text:list>
											</text:list-item>
										</text:list>
									</text:list-item>
								</text:list>
							</text:list-item>
						</text:list>
					</text:list-item>
				</text:list>
			</xsl:when>
			<xsl:when test ="$levelCount='6'">
				<text:list>
					<xsl:attribute name ="text:style-name">
						<xsl:value-of select ="concat($listStyleName,'lvl',$levelCount)"/>
					</xsl:attribute>
					<text:list-item>
						<text:list>
							<text:list-item>
								<text:list>
									<text:list-item>
										<text:list>
											<text:list-item>
												<text:list>
													<text:list-item>
														<text:list>
															<text:list-item>
																<text:list>
																	<text:list-item>
																		<text:p >
																			<xsl:attribute name ="text:style-name">
																				<xsl:value-of select ="concat($ParaId,position())"/>
																			</xsl:attribute>
																			<xsl:for-each select ="node()">
																				<xsl:if test ="name()='a:r'">
																					<text:span text:style-name="{generate-id()}">
																						<!--<xsl:value-of select ="a:t"/>-->
																						<!--converts whitespaces sequence to text:s-->
																						<!-- 1699083 bug fix  -->
																						<xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
																						<xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
																						<xsl:choose >
																							<xsl:when test ="a:rPr[@cap='all']">
																								<xsl:choose >
																									<xsl:when test =".=''">
																										<text:s/>
																									</xsl:when>
																									<xsl:when test ="not(contains(.,'  '))">
																										<xsl:value-of select ="translate(.,$lcletters,$ucletters)"/>
																									</xsl:when>
																									<xsl:when test =". =' '">
																										<text:s/>
																									</xsl:when>
																									<xsl:otherwise >
																										<xsl:call-template name ="InsertWhiteSpaces">
																											<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
																										</xsl:call-template>
																									</xsl:otherwise>
																								</xsl:choose>
																							</xsl:when>
																							<xsl:when test ="a:rPr[@cap='small']">
																								<xsl:choose >
																									<xsl:when test =".=''">
																										<text:s/>
																									</xsl:when>
																									<xsl:when test ="not(contains(.,'  '))">
																										<xsl:value-of select ="translate(.,$ucletters,$lcletters)"/>
																									</xsl:when>
																									<xsl:when test =".= ' '">
																										<text:s/>
																									</xsl:when>
																									<xsl:otherwise >
																										<xsl:call-template name ="InsertWhiteSpaces">
																											<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
																										</xsl:call-template>
																									</xsl:otherwise>
																								</xsl:choose >
																							</xsl:when>
																							<xsl:otherwise >
																								<xsl:choose >
																									<xsl:when test =".=''">
																										<text:s/>
																									</xsl:when>
																									<xsl:when test ="not(contains(.,'  '))">
																										<xsl:value-of select ="."/>
																									</xsl:when>
																									<xsl:otherwise >
																										<xsl:call-template name ="InsertWhiteSpaces">
																											<xsl:with-param name ="string" select ="."/>
																										</xsl:call-template>
																									</xsl:otherwise >
																								</xsl:choose>
																							</xsl:otherwise>
																						</xsl:choose>
																					</text:span>
																				</xsl:if >
																				<xsl:if test ="name()='a:br'">
																					<text:line-break/>
																				</xsl:if>
																			</xsl:for-each>
																		</text:p>
																	</text:list-item>
																</text:list>
															</text:list-item>
														</text:list>
													</text:list-item>
												</text:list>
											</text:list-item>
										</text:list>
									</text:list-item>
								</text:list>
							</text:list-item>
						</text:list>
					</text:list-item>
				</text:list>
			</xsl:when>
			<xsl:when test ="$levelCount='7'">
				<text:list>
					<xsl:attribute name ="text:style-name">
						<xsl:value-of select ="concat($listStyleName,'lvl',$levelCount)"/>
					</xsl:attribute>
					<text:list-item>
						<text:list>
							<text:list-item>
								<text:list>
									<text:list-item>
										<text:list>
											<text:list-item>
												<text:list>
													<text:list-item>
														<text:list>
															<text:list-item>
																<text:list>
																	<text:list-item>
																		<text:list>
																			<text:list-item>
																				<text:p >
																					<xsl:attribute name ="text:style-name">
																						<xsl:value-of select ="concat($ParaId,position())"/>
																					</xsl:attribute>
																					<xsl:for-each select ="node()">
																						<xsl:if test ="name()='a:r'">
																							<text:span text:style-name="{generate-id()}">
																								<!--<xsl:value-of select ="a:t"/>-->
																								<!--converts whitespaces sequence to text:s-->
																								<!-- 1699083 bug fix  -->
																								<xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
																								<xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
																								<xsl:choose >
																									<xsl:when test ="a:rPr[@cap='all']">
																										<xsl:choose >
																											<xsl:when test =".=''">
																												<text:s/>
																											</xsl:when>
																											<xsl:when test ="not(contains(.,'  '))">
																												<xsl:value-of select ="translate(.,$lcletters,$ucletters)"/>
																											</xsl:when>
																											<xsl:when test =". =' '">
																												<text:s/>
																											</xsl:when>
																											<xsl:otherwise >
																												<xsl:call-template name ="InsertWhiteSpaces">
																													<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
																												</xsl:call-template>
																											</xsl:otherwise>
																										</xsl:choose>
																									</xsl:when>
																									<xsl:when test ="a:rPr[@cap='small']">
																										<xsl:choose >
																											<xsl:when test =".=''">
																												<text:s/>
																											</xsl:when>
																											<xsl:when test ="not(contains(.,'  '))">
																												<xsl:value-of select ="translate(.,$ucletters,$lcletters)"/>
																											</xsl:when>
																											<xsl:when test =".= ' '">
																												<text:s/>
																											</xsl:when>
																											<xsl:otherwise >
																												<xsl:call-template name ="InsertWhiteSpaces">
																													<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
																												</xsl:call-template>
																											</xsl:otherwise>
																										</xsl:choose >
																									</xsl:when>
																									<xsl:otherwise >
																										<xsl:choose >
																											<xsl:when test =".=''">
																												<text:s/>
																											</xsl:when>
																											<xsl:when test ="not(contains(.,'  '))">
																												<xsl:value-of select ="."/>
																											</xsl:when>
																											<xsl:otherwise >
																												<xsl:call-template name ="InsertWhiteSpaces">
																													<xsl:with-param name ="string" select ="."/>
																												</xsl:call-template>
																											</xsl:otherwise >
																										</xsl:choose>
																									</xsl:otherwise>
																								</xsl:choose>
																							</text:span>
																						</xsl:if >
																						<xsl:if test ="name()='a:br'">
																							<text:line-break/>
																						</xsl:if>
																					</xsl:for-each>
																				</text:p>
																			</text:list-item>
																		</text:list>
																	</text:list-item>
																</text:list>
															</text:list-item>
														</text:list>
													</text:list-item>
												</text:list>
											</text:list-item>
										</text:list>
									</text:list-item>
								</text:list>
							</text:list-item>
						</text:list>
					</text:list-item>
				</text:list>
			</xsl:when>
			<xsl:when test ="$levelCount='8'">
				<text:list>
					<xsl:attribute name ="text:style-name">
						<xsl:value-of select ="concat($listStyleName,'lvl',$levelCount)"/>
					</xsl:attribute>
					<text:list-item>
						<text:list>
							<text:list-item>
								<text:list>
									<text:list-item>
										<text:list>
											<text:list-item>
												<text:list>
													<text:list-item>
														<text:list>
															<text:list-item>
																<text:list>
																	<text:list-item>
																		<text:list>
																			<text:list-item>
																				<text:list>
																					<text:list-item>
																						<text:p >
																							<xsl:attribute name ="text:style-name">
																								<xsl:value-of select ="concat($ParaId,position())"/>
																							</xsl:attribute>
																							<xsl:for-each select ="node()">
																								<xsl:if test ="name()='a:r'">
																									<text:span text:style-name="{generate-id()}">
																										<!--<xsl:value-of select ="a:t"/>-->
																										<!--converts whitespaces sequence to text:s-->
																										<!-- 1699083 bug fix  -->
																										<xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
																										<xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
																										<xsl:choose >
																											<xsl:when test ="a:rPr[@cap='all']">
																												<xsl:choose >
																													<xsl:when test =".=''">
																														<text:s/>
																													</xsl:when>
																													<xsl:when test ="not(contains(.,'  '))">
																														<xsl:value-of select ="translate(.,$lcletters,$ucletters)"/>
																													</xsl:when>
																													<xsl:when test =". =' '">
																														<text:s/>
																													</xsl:when>
																													<xsl:otherwise >
																														<xsl:call-template name ="InsertWhiteSpaces">
																															<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
																														</xsl:call-template>
																													</xsl:otherwise>
																												</xsl:choose>
																											</xsl:when>
																											<xsl:when test ="a:rPr[@cap='small']">
																												<xsl:choose >
																													<xsl:when test =".=''">
																														<text:s/>
																													</xsl:when>
																													<xsl:when test ="not(contains(.,'  '))">
																														<xsl:value-of select ="translate(.,$ucletters,$lcletters)"/>
																													</xsl:when>
																													<xsl:when test =".= ' '">
																														<text:s/>
																													</xsl:when>
																													<xsl:otherwise >
																														<xsl:call-template name ="InsertWhiteSpaces">
																															<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
																														</xsl:call-template>
																													</xsl:otherwise>
																												</xsl:choose >
																											</xsl:when>
																											<xsl:otherwise >
																												<xsl:choose >
																													<xsl:when test =".=''">
																														<text:s/>
																													</xsl:when>
																													<xsl:when test ="not(contains(.,'  '))">
																														<xsl:value-of select ="."/>
																													</xsl:when>
																													<xsl:otherwise >
																														<xsl:call-template name ="InsertWhiteSpaces">
																															<xsl:with-param name ="string" select ="."/>
																														</xsl:call-template>
																													</xsl:otherwise >
																												</xsl:choose>
																											</xsl:otherwise>
																										</xsl:choose>
																									</text:span>
																								</xsl:if >
																								<xsl:if test ="name()='a:br'">
																									<text:line-break/>
																								</xsl:if>
																							</xsl:for-each>
																						</text:p>
																					</text:list-item>
																				</text:list>
																			</text:list-item>
																		</text:list>
																	</text:list-item>
																</text:list>
															</text:list-item>
														</text:list>
													</text:list-item>
												</text:list>
											</text:list-item>
										</text:list>
									</text:list-item>
								</text:list>
							</text:list-item>
						</text:list>
					</text:list-item>
				</text:list>
			</xsl:when>
			<xsl:when test ="$levelCount='9'">
				<text:list>
					<xsl:attribute name ="text:style-name">
						<xsl:value-of select ="concat($listStyleName,'lvl',$levelCount)"/>
					</xsl:attribute>
					<text:list-item>
						<text:list>
							<text:list-item>
								<text:list>
									<text:list-item>
										<text:list>
											<text:list-item>
												<text:list>
													<text:list-item>
														<text:list>
															<text:list-item>
																<text:list>
																	<text:list-item>
																		<text:list>
																			<text:list-item>
																				<text:list>
																					<text:list-item>
																						<text:list>
																							<text:list-item>
																								<text:p >
																									<xsl:attribute name ="text:style-name">
																										<xsl:value-of select ="concat($ParaId,position())"/>
																									</xsl:attribute>
																									<xsl:for-each select ="node()">
																										<xsl:if test ="name()='a:r'">
																											<text:span text:style-name="{generate-id()}">
																												<!--<xsl:value-of select ="a:t"/>-->
																												<!--converts whitespaces sequence to text:s-->
																												<!-- 1699083 bug fix  -->
																												<xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
																												<xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
																												<xsl:choose >
																													<xsl:when test ="a:rPr[@cap='all']">
																														<xsl:choose >
																															<xsl:when test =".=''">
																																<text:s/>
																															</xsl:when>
																															<xsl:when test ="not(contains(.,'  '))">
																																<xsl:value-of select ="translate(.,$lcletters,$ucletters)"/>
																															</xsl:when>
																															<xsl:when test =". =' '">
																																<text:s/>
																															</xsl:when>
																															<xsl:otherwise >
																																<xsl:call-template name ="InsertWhiteSpaces">
																																	<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
																																</xsl:call-template>
																															</xsl:otherwise>
																														</xsl:choose>
																													</xsl:when>
																													<xsl:when test ="a:rPr[@cap='small']">
																														<xsl:choose >
																															<xsl:when test =".=''">
																																<text:s/>
																															</xsl:when>
																															<xsl:when test ="not(contains(.,'  '))">
																																<xsl:value-of select ="translate(.,$ucletters,$lcletters)"/>
																															</xsl:when>
																															<xsl:when test =".= ' '">
																																<text:s/>
																															</xsl:when>
																															<xsl:otherwise >
																																<xsl:call-template name ="InsertWhiteSpaces">
																																	<xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
																																</xsl:call-template>
																															</xsl:otherwise>
																														</xsl:choose >
																													</xsl:when>
																													<xsl:otherwise >
																														<xsl:choose >
																															<xsl:when test =".=''">
																																<text:s/>
																															</xsl:when>
																															<xsl:when test ="not(contains(.,'  '))">
																																<xsl:value-of select ="."/>
																															</xsl:when>
																															<xsl:otherwise >
																																<xsl:call-template name ="InsertWhiteSpaces">
																																	<xsl:with-param name ="string" select ="."/>
																																</xsl:call-template>
																															</xsl:otherwise >
																														</xsl:choose>
																													</xsl:otherwise>
																												</xsl:choose>
																											</text:span>
																										</xsl:if >
																										<xsl:if test ="name()='a:br'">
																											<text:line-break/>
																										</xsl:if>
																									</xsl:for-each>
																								</text:p>
																							</text:list-item>
																						</text:list>
																					</text:list-item>
																				</text:list>
																			</text:list-item>
																		</text:list>
																	</text:list-item>
																</text:list>
															</text:list-item>
														</text:list>
													</text:list-item>
												</text:list>
											</text:list-item>
										</text:list>
									</text:list-item>
								</text:list>
							</text:list-item>
						</text:list>
					</text:list-item>
				</text:list>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="copyPictures">
		<xsl:param name="document"/>
		<xsl:param name="rId"/>
		<xsl:param name="targetName"/>
		<xsl:param name="destFolder" select="'Pictures'"/>
		<!--  Copy Pictures Files to the picture catalog -->
		<xsl:variable name="id">
			<xsl:choose>
				<xsl:when test="$rId != ''">
					<xsl:value-of select="$rId"/>
				</xsl:when>
			</xsl:choose>
		</xsl:variable>
		<xsl:for-each
		  select="document($document)//rels:Relationships/rels:Relationship">
			<xsl:if test="./@Id=$id">
				<xsl:variable name="targetmode">
					<xsl:value-of select="./@TargetMode"/>
				</xsl:variable>
				<xsl:variable name="pzipsource">
					<xsl:value-of select="substring-after(./@Target,'../media/')"/>
					<!-- image1.gif-->
				</xsl:variable>
				<xsl:variable name="pziptarget">
					<xsl:choose>
						<xsl:when test="$targetName != ''">
							<xsl:value-of select="$targetName"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="$pzipsource"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:variable>
				<xsl:choose>
					<xsl:when test="$targetmode='External'"/>
					<xsl:when test="$destFolder = '.'">
						<pzip:copy pzip:source="{concat('ppt/media/',$pzipsource)}" pzip:target="{$pziptarget}"/>
					</xsl:when>
					<xsl:otherwise>
						<pzip:copy pzip:source="{concat('ppt/media/',$pzipsource)}" pzip:target="{concat($destFolder,'/',$pzipsource)}"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:if>
		</xsl:for-each>

	</xsl:template>
	<!-- level colour-->
	<xsl:template name ="getLevelColor">
		<xsl:param name ="levelNum"/>
		<xsl:choose>
			<xsl:when test ="$levelNum='1'">
				<xsl:if test ="a:lvl1pPr/a:buClr/a:srgbClr">
					<xsl:value-of select ="a:lvl1pPr/a:buClr/a:srgbClr/@val"/>
				</xsl:if>
				<xsl:if test ="a:lvl1pPr/a:buClr/a:schemeClr">
					<xsl:call-template name ="getColorCode">
						<xsl:with-param name ="color">
							<xsl:value-of select="a:lvl1pPr/a:buClr/a:schemeClr/@val"/>
						</xsl:with-param>
					</xsl:call-template>
				</xsl:if >
			</xsl:when>
			<xsl:when test ="$levelNum='2'">
				<xsl:if test ="a:lvl2pPr/a:buClr/a:srgbClr">
					<xsl:value-of select ="a:lvl2pPr/a:buClr/a:srgbClr/@val"/>
				</xsl:if>
				<xsl:if test ="a:lvl2pPr/a:buClr/a:schemeClr">
					<xsl:call-template name ="getColorCode">
						<xsl:with-param name ="color">
							<xsl:value-of select="a:lvl2pPr/a:buClr/a:schemeClr/@val"/>
						</xsl:with-param>
					</xsl:call-template>
				</xsl:if >
			</xsl:when>
			<xsl:when test ="$levelNum='3'">
				<xsl:if test ="a:lvl3pPr/a:buClr/a:srgbClr">
					<xsl:value-of select ="a:lvl3pPr/a:buClr/a:srgbClr/@val"/>
				</xsl:if>
				<xsl:if test ="a:lvl3pPr/a:buClr/a:schemeClr">
					<xsl:call-template name ="getColorCode">
						<xsl:with-param name ="color">
							<xsl:value-of select="a:lvl3pPr/a:buClr/a:schemeClr/@val"/>
						</xsl:with-param>
					</xsl:call-template>
				</xsl:if >
			</xsl:when>
			<xsl:when test ="$levelNum='4'">
				<xsl:if test ="a:lvl4pPr/a:buClr/a:srgbClr">
					<xsl:value-of select ="a:lvl4pPr/a:buClr/a:srgbClr/@val"/>
				</xsl:if>
				<xsl:if test ="a:lvl4pPr/a:buClr/a:schemeClr">
					<xsl:call-template name ="getColorCode">
						<xsl:with-param name ="color">
							<xsl:value-of select="a:lvl4pPr/a:buClr/a:schemeClr/@val"/>
						</xsl:with-param>
					</xsl:call-template>
				</xsl:if >
			</xsl:when>
			<xsl:when test ="$levelNum='5'">
				<xsl:if test ="a:lvl5pPr/a:buClr/a:srgbClr">
					<xsl:value-of select ="a:lvl5pPr/a:buClr/a:srgbClr/@val"/>
				</xsl:if>
				<xsl:if test ="a:lvl5pPr/a:buClr/a:schemeClr">
					<xsl:call-template name ="getColorCode">
						<xsl:with-param name ="color">
							<xsl:value-of select="a:lvl5pPr/a:buClr/a:schemeClr/@val"/>
						</xsl:with-param>
					</xsl:call-template>
				</xsl:if >
			</xsl:when>
			<xsl:when test ="$levelNum='6'">
				<xsl:if test ="a:lvl6pPr/a:buClr/a:srgbClr">
					<xsl:value-of select ="a:lvl6pPr/a:buClr/a:srgbClr/@val"/>
				</xsl:if>
				<xsl:if test ="a:lvl6pPr/a:buClr/a:schemeClr">
					<xsl:call-template name ="getColorCode">
						<xsl:with-param name ="color">
							<xsl:value-of select="a:lvl6pPr/a:buClr/a:schemeClr/@val"/>
						</xsl:with-param>
					</xsl:call-template>
				</xsl:if >
			</xsl:when>
			<xsl:when test ="$levelNum='7'">
				<xsl:if test ="a:lvl7pPr/a:buClr/a:srgbClr">
					<xsl:value-of select ="a:lvl7pPr/a:buClr/a:srgbClr/@val"/>
				</xsl:if>
				<xsl:if test ="a:lvl7pPr/a:buClr/a:schemeClr">
					<xsl:call-template name ="getColorCode">
						<xsl:with-param name ="color">
							<xsl:value-of select="a:lvl7pPr/a:buClr/a:schemeClr/@val"/>
						</xsl:with-param>
					</xsl:call-template>
				</xsl:if >
			</xsl:when>
			<xsl:when test ="$levelNum='8'">
				<xsl:if test ="a:lvl8pPr/a:buClr/a:srgbClr">
					<xsl:value-of select ="a:lvl8pPr/a:buClr/a:srgbClr/@val"/>
				</xsl:if>
				<xsl:if test ="a:lvl8pPr/a:buClr/a:schemeClr">
					<xsl:call-template name ="getColorCode">
						<xsl:with-param name ="color">
							<xsl:value-of select="a:lvl8pPr/a:buClr/a:schemeClr/@val"/>
						</xsl:with-param>
					</xsl:call-template>
				</xsl:if >
			</xsl:when>
			<xsl:when test ="$levelNum='9'">
				<xsl:if test ="a:lvl9pPr/a:buClr/a:srgbClr">
					<xsl:value-of select ="a:lvl9pPr/a:buClr/a:srgbClr/@val"/>
				</xsl:if>
				<xsl:if test ="a:lvl9pPr/a:buClr/a:schemeClr">
					<xsl:call-template name ="getColorCode">
						<xsl:with-param name ="color">
							<xsl:value-of select="a:lvl9pPr/a:buClr/a:schemeClr/@val"/>
						</xsl:with-param>
					</xsl:call-template>
				</xsl:if >
			</xsl:when>
		</xsl:choose>

	</xsl:template>
	<!-- Template insertBulletCharacter-->
	<xsl:template name ="insertBulletCharacter">
		<xsl:param name ="character"/>
		<xsl:choose>
			<xsl:when test="$character= '•'">
				<xsl:value-of select ="'•'"/>
			</xsl:when>
			<xsl:when test="$character= 'Ø'">
				<xsl:value-of select ="'➢'"/>
			</xsl:when>
			<xsl:when test="$character= 'o'">
				<xsl:value-of select ="'○'"/>
			</xsl:when>
			<!--<xsl:when test="a:pPr/a:buChar/@char= '§'">
              <xsl:value-of select ="'■'"/>
          </xsl:when>
          <xsl:when test="a:pPr/a:buChar/@char= 'q'">
              <xsl:value-of select ="''"/>
          </xsl:when>-->
			<xsl:when test="$character= 'ü'">
				<xsl:value-of select ="'✔'"/>
			</xsl:when>
			<xsl:when test="$character = '–'">
				<xsl:value-of select ="'–'"/>
			</xsl:when>
			<xsl:when test="$character = '»'">
				<xsl:value-of select ="'»'"/>
			</xsl:when>
			<!-- Added by  vijayeta ,on 19th june-->
			<xsl:when test="$character = 'è'">
				<xsl:value-of select ="'➔'"/>
			</xsl:when>
			<!-- Added by  vijayeta ,on 19th june-->
			<xsl:otherwise>•</xsl:otherwise>
		</xsl:choose >
	</xsl:template>
	<xsl:template name ="insertNumber">
		<xsl:param name ="number"/>
		<xsl:param name ="startAt"/>
		<xsl:if test ="$number='arabicPeriod'">
			<xsl:attribute name ="style:num-format" >
				<xsl:value-of  select ="'1'"/>
			</xsl:attribute >
			<xsl:attribute name ="style:num-suffix">
				<xsl:value-of select ="'.'"/>
			</xsl:attribute>
		</xsl:if>
		<xsl:if test ="$number='arabicParenR'">
			<xsl:attribute name ="style:num-format" >
				<xsl:value-of  select ="'1'"/>
			</xsl:attribute >
			<xsl:attribute name ="style:num-suffix">
				<xsl:value-of select ="')'"/>
			</xsl:attribute>
		</xsl:if>
		<xsl:if test ="$number='arabicParenBoth'">
			<xsl:attribute name ="style:num-format" >
				<xsl:value-of  select ="'1'"/>
			</xsl:attribute >
			<xsl:attribute name ="style:num-suffix">
				<xsl:value-of select ="')'"/>
			</xsl:attribute>
			<xsl:attribute name ="style:num-prefix">
				<xsl:value-of select ="'('"/>
			</xsl:attribute>
		</xsl:if>
		<xsl:if test ="$number='alphaUcPeriod'">
			<xsl:attribute name ="style:num-format" >
				<xsl:value-of  select ="'A'"/>
			</xsl:attribute >
			<xsl:attribute name ="style:num-suffix">
				<xsl:value-of select ="'.'"/>
			</xsl:attribute>
		</xsl:if>
		<xsl:if test ="$number='alphaUcParenR'">
			<xsl:attribute name ="style:num-format" >
				<xsl:value-of  select ="'A'"/>
			</xsl:attribute >
			<xsl:attribute name ="style:num-suffix">
				<xsl:value-of select ="')'"/>
			</xsl:attribute>
		</xsl:if>
		<xsl:if test ="$number='alphaUcParenBoth'">
			<xsl:attribute name ="style:num-format" >
				<xsl:value-of  select ="'A'"/>
			</xsl:attribute >
			<xsl:attribute name ="style:num-suffix">
				<xsl:value-of select ="')'"/>
			</xsl:attribute>
			<xsl:attribute name ="style:num-prefix">
				<xsl:value-of select ="'('"/>
			</xsl:attribute>
		</xsl:if>
		<xsl:if test ="$number='alphaLcPeriod'">
			<xsl:attribute name ="style:num-format" >
				<xsl:value-of  select ="'a'"/>
			</xsl:attribute >
			<xsl:attribute name ="style:num-suffix">
				<xsl:value-of select ="'.'"/>
			</xsl:attribute>
		</xsl:if>
		<xsl:if test ="$number='alphaLcParenR'">
			<xsl:attribute name ="style:num-format" >
				<xsl:value-of  select ="'a'"/>
			</xsl:attribute >
			<xsl:attribute name ="style:num-suffix">
				<xsl:value-of select ="')'"/>
			</xsl:attribute>
		</xsl:if>
		<xsl:if test ="$number='alphaLcParenBoth'">
			<xsl:attribute name ="style:num-format" >
				<xsl:value-of  select ="'a'"/>
			</xsl:attribute >
			<xsl:attribute name ="style:num-suffix">
				<xsl:value-of select ="')'"/>
			</xsl:attribute>
			<xsl:attribute name ="style:num-prefix">
				<xsl:value-of select ="'('"/>
			</xsl:attribute>
		</xsl:if>
		<xsl:if test ="$number='romanUcPeriod'">
			<xsl:attribute name ="style:num-format" >
				<xsl:value-of  select ="'I'"/>
			</xsl:attribute >
			<xsl:attribute name ="style:num-suffix">
				<xsl:value-of select ="'.'"/>
			</xsl:attribute>
		</xsl:if>
		<xsl:if test ="$number='romanLcPeriod'">
			<xsl:attribute name ="style:num-format" >
				<xsl:value-of  select ="'i'"/>
			</xsl:attribute >
			<xsl:attribute name ="style:num-suffix">
				<xsl:value-of select ="'.'"/>
			</xsl:attribute>
		</xsl:if >
		<!-- start at value-->
		<xsl:attribute name ="text:start-value">
			<xsl:value-of select ="$startAt"/>
		</xsl:attribute>
	</xsl:template>
	<xsl:template name ="getBulletForLevelsLayout">
		<xsl:param name ="slideMaster"/>
		<xsl:param name ="newTextLvl"/>
		<xsl:param name ="levelStyle"/>
		<xsl:choose>
			<xsl:when test ="$newTextLvl='1'">
				<xsl:call-template name ="getBullet_Level1">
					<xsl:with-param name ="slideMaster"  select ="$slideMaster"/>
					<xsl:with-param name ="newTextLvl" select ="$newTextLvl"  />
					<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
				</xsl:call-template>
			</xsl:when>
			<xsl:when test ="$newTextLvl='2'">
				<xsl:call-template name ="getBullet_Level2">
					<xsl:with-param name ="slideMaster"  select ="$slideMaster"/>
					<xsl:with-param name ="newTextLvl" select ="$newTextLvl"  />
					<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
				</xsl:call-template>
			</xsl:when>
			<xsl:when test ="$newTextLvl='3'">
				<xsl:call-template name ="getBullet_Level3">
					<xsl:with-param name ="slideMaster"  select ="$slideMaster"/>
					<xsl:with-param name ="newTextLvl" select ="$newTextLvl"  />
					<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
				</xsl:call-template>
			</xsl:when>
			<xsl:when test ="$newTextLvl='4'">
				<xsl:call-template name ="getBullet_Level4">
					<xsl:with-param name ="slideMaster"  select ="$slideMaster"/>
					<xsl:with-param name ="newTextLvl" select ="$newTextLvl"  />
					<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
				</xsl:call-template>

			</xsl:when>
			<xsl:when test ="$newTextLvl='5'">
				<xsl:call-template name ="getBullet_Level5">
					<xsl:with-param name ="slideMaster"  select ="$slideMaster"/>
					<xsl:with-param name ="newTextLvl" select ="$newTextLvl"  />
					<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
				</xsl:call-template>
			</xsl:when>
			<xsl:when test ="$newTextLvl='6'">
				<xsl:call-template name ="getBullet_Level6">
					<xsl:with-param name ="slideMaster"  select ="$slideMaster"/>
					<xsl:with-param name ="newTextLvl" select ="$newTextLvl"  />
					<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
				</xsl:call-template>
			</xsl:when>
			<xsl:when test ="$newTextLvl='7'">
				<xsl:call-template name ="getBullet_Level7">
					<xsl:with-param name ="slideMaster"  select ="$slideMaster"/>
					<xsl:with-param name ="newTextLvl" select ="$newTextLvl"  />
					<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
				</xsl:call-template>
			</xsl:when>
			<xsl:when test ="$newTextLvl='8'">
				<xsl:call-template name ="getBullet_Level8">
					<xsl:with-param name ="slideMaster"  select ="$slideMaster"/>
					<xsl:with-param name ="newTextLvl" select ="$newTextLvl"  />
					<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
				</xsl:call-template>
			</xsl:when>
			<xsl:when test ="$newTextLvl='9'">
				<xsl:call-template name ="getBullet_Level9">
					<xsl:with-param name ="slideMaster"  select ="$slideMaster"/>
					<xsl:with-param name ="newTextLvl" select ="$newTextLvl"  />
					<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
				</xsl:call-template>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<xsl:template name ="getBulletsFromSlideMaster">
		<xsl:param name ="slideMaster" />
		<xsl:param name ="newTextLvl"  />
		<xsl:param name ="levelStyle"/>
		<xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle/child::node()[$newTextLvl]">
			<xsl:if test ="a:buChar">
				<text:list-style>
					<xsl:attribute name ="style:name">
						<xsl:value-of select ="$levelStyle"/>
					</xsl:attribute >
					<text:list-level-style-bullet>
						<xsl:attribute name ="text:level">
							<xsl:value-of select ="$newTextLvl"/>
						</xsl:attribute >
						<xsl:attribute name ="text:bullet-char">
							<xsl:call-template name ="insertBulletCharacter">
								<xsl:with-param name ="character" select ="a:buChar/@char" />
							</xsl:call-template>
						</xsl:attribute >
						<style:list-level-properties text:min-label-width="0.8cm">
							<xsl:attribute name ="text:space-before">
								<xsl:value-of select="concat(format-number(@marL div 360000,'#.##'),'cm')"/>
							</xsl:attribute>
						</style:list-level-properties>
						<style:text-properties fo:font-family="StarSymbol" style:font-charset="x-symbol">
							<xsl:if test ="a:buFont/@typeface">
								<xsl:if test ="a:buFont[@typeface='Arial']">
									<xsl:attribute name ="fo:font-family">
										<xsl:value-of select ="'StarSymbol'"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="not(a:buFont[@typeface='Arial'])">
									<xsl:attribute name ="fo:font-family">
										<xsl:value-of select ="a:buFont/@typeface"/>
									</xsl:attribute>
								</xsl:if>
							</xsl:if>
							<xsl:if test ="not(a:buFont/@typeface)">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="'StarSymbol'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="a:buSzPct">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="concat((a:buSzPct/@val div 1000),'%')"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(a:buSzPct)">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="'100%'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="a:buClr">
								<xsl:if test ="a:buClr/a:srgbClr">
									<xsl:variable name ="color" select ="a:buClr/a:srgbClr/@val"/>
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',$color)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="a:buClr/a:schemeClr">
									<xsl:call-template name ="getColorCode">
										<xsl:with-param name ="color">
											<xsl:value-of select="a:buClr/a:schemeClr/@val"/>
										</xsl:with-param>
									</xsl:call-template >
								</xsl:if>
							</xsl:if>
							<xsl:if test ="not(a:buClr)">
								<xsl:if test ="a:buClrTx or ./a:defRPr/a:solidFill">
									<xsl:if test ="./a:defRPr/a:solidFill/a:srgbClr">
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',./a:defRPr/a:solidFill/a:srgbClr/@val)"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="./a:defRPr/a:solidFill/a:schemeClr">
										<xsl:attribute name ="fo:color">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="./a:defRPr/a:solidFill/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:attribute>
									</xsl:if>
								</xsl:if>
							</xsl:if>
						</style:text-properties >
					</text:list-level-style-bullet>
				</text:list-style>
			</xsl:if>
			<xsl:if test ="a:buAutoNum">
				<text:list-style>
					<xsl:attribute name ="style:name">
						<xsl:value-of select ="$levelStyle"/>
					</xsl:attribute >
					<text:list-level-style-number>
						<xsl:attribute name ="text:level">
							<xsl:value-of select ="$newTextLvl"/>
						</xsl:attribute >
						<xsl:variable name ="startAt">
							<xsl:if test ="a:buAutoNum/@startAt">
								<xsl:value-of select ="a:buAutoNum/@startAt" />
							</xsl:if>
							<xsl:if test ="not(a:buAutoNum/@startAt)">
								<xsl:value-of select ="'1'" />
							</xsl:if>
						</xsl:variable>
						<xsl:call-template name ="insertNumber">
							<xsl:with-param name ="number" select ="a:buAutoNum/@type"/>
							<xsl:with-param name ="startAt" select ="$startAt"/>
						</xsl:call-template>
						<style:list-level-properties text:min-label-width="0.952cm" />
						<style:text-properties style:font-family-generic="swiss" style:font-pitch="variable" fo:font-size="100%">
							<xsl:if test ="a:buFont/@typeface">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="a:buFont/@typeface"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(a:buFont/@typeface)">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="'Arial'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="a:buSzPct">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="concat((a:buSzPct/@val div 1000),'%')"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(a:buSzPct)">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="'100%'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="a:buClr">
								<xsl:if test ="a:buClr/a:srgbClr">
									<xsl:variable name ="color" select ="a:buClr/a:srgbClr/@val"/>
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',$color)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="a:buClr/a:schemeClr">
									<xsl:call-template name ="getColorCode">
										<xsl:with-param name ="color">
											<xsl:value-of select="a:buClr/a:schemeClr/@val"/>
										</xsl:with-param>
									</xsl:call-template >
								</xsl:if>
							</xsl:if>
							<xsl:if test ="not(a:buClr)">
								<xsl:if test ="a:buClrTx or ./a:defRPr/a:solidFill">
									<xsl:if test ="./a:defRPr/a:solidFill/a:srgbClr">
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',./a:defRPr/a:solidFill/a:srgbClr/@val)"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="./a:defRPr/a:solidFill/a:schemeClr">
										<xsl:attribute name ="fo:color">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="./a:defRPr/a:solidFill/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:attribute>
									</xsl:if>
								</xsl:if>
							</xsl:if>
						</style:text-properties>
					</text:list-level-style-number>
				</text:list-style>
			</xsl:if>
			<xsl:if test ="a:buBlip">
				<xsl:variable name ="rId" select ="a:buBlip/a:blip/@r:embed"/>
				<xsl:variable name="XlinkHref">
					<xsl:variable name="pzipsource">
						<xsl:value-of select="document('ppt/slideMasters/_rels/slideMaster1.xml.rels')//rels:Relationships/rels:Relationship[@Id=$rId]/@Target"/>
					</xsl:variable>
					<xsl:value-of select="concat('Pictures/', substring-after($pzipsource,'../media/'))"/>
				</xsl:variable>
				<xsl:call-template name="copyPictures">
					<xsl:with-param name="document">
						<xsl:value-of select="document('ppt/slideMasters/_rels/slideMaster1.xml.rels')"/>
					</xsl:with-param>
					<xsl:with-param name="rId">
						<xsl:value-of select ="$rId"/>
					</xsl:with-param>
				</xsl:call-template>
				<text:list-level-style-image text:level="1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad">
					<xsl:attribute name="xlink:href">
						<xsl:value-of select="$XlinkHref"/>
					</xsl:attribute>
					<style:list-level-properties style:vertical-pos="middle" style:vertical-rel="line" fo:width="0.318cm" fo:height="0.318cm" />
				</text:list-level-style-image>
			</xsl:if>
			<!--</text:list-style>-->
		</xsl:for-each>
	</xsl:template>

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
	<xsl:template name ="getBullet_Level1">				
		<xsl:param name ="slideMaster" />
		<xsl:param name ="newTextLvl" />
		<xsl:param name ="levelStyle"/>
		<xsl:choose >
			<xsl:when test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buChar">
				<text:list-style>
					<xsl:attribute name ="style:name">
						<xsl:value-of select ="$levelStyle"/>
					</xsl:attribute >
					<text:list-level-style-bullet>
						<xsl:attribute name ="text:level">
							<xsl:value-of select ="$newTextLvl"/>
						</xsl:attribute >
						<xsl:attribute name ="text:bullet-char">
							<xsl:call-template name ="insertBulletCharacter">
								<xsl:with-param name ="character" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buChar/@char" />
							</xsl:call-template>
						</xsl:attribute >
						<style:list-level-properties text:min-label-width="0.8cm"/>
						<style:text-properties style:font-charset="x-symbol">
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buFont/@typeface">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buFont[@typeface='Arial']">
									<xsl:attribute name ="fo:font-family">
										<xsl:value-of select ="'StarSymbol'"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buFont[@typeface='Arial'])">
									<xsl:attribute name ="fo:font-family">
										<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buFont/@typeface"/>
									</xsl:attribute>
								</xsl:if>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buFont/@typeface)">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="'StarSymbol'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buSzPct">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="concat((./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buSzPct/@val div 1000),'%')"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buSzPct)">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="'100%'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buClr">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buClr/a:srgbClr">
									<xsl:variable name ="color" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buClr/a:srgbClr/@val"/>
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',$color)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buClr/a:schemeClr">
									<xsl:attribute name ="fo:color">
										<xsl:call-template name ="getColorCode">
											<xsl:with-param name ="color">
												<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buClr/a:schemeClr/@val"/>
											</xsl:with-param>
										</xsl:call-template >
									</xsl:attribute>
								</xsl:if>
							</xsl:if>
							<!-- Code added by vijayeta, bug fix 1746350-->
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buClr)">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill">
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr">
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr">
										<xsl:attribute name ="fo:color">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:attribute>
									</xsl:if>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill)">
									<xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
										<xsl:variable name ="levelColor" >
											<xsl:call-template name ="getLevelColor">
												<xsl:with-param name ="levelNum" select ="$newTextLvl"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',$levelColor)"/>
										</xsl:attribute>
									</xsl:for-each>
								</xsl:if>
							</xsl:if>
							<!--End of Code added by vijayeta, bug fix 1746350-->
						</style:text-properties >
					</text:list-level-style-bullet>
				</text:list-style>
			</xsl:when>
			<xsl:when test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buAutoNum">
				<xsl:variable name ="Number" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buAutoNum/@type"/>
				<text:list-level-style-number>
					<xsl:attribute name ="text:level">
						<xsl:value-of select ="$newTextLvl"/>
					</xsl:attribute >
					<xsl:variable name ="startAt">
						<xsl:if test ="a:pPr/a:buAutoNum/@startAt">
							<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buAutoNum/@startAt" />
						</xsl:if>
						<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buAutoNum/@startAt)">
							<xsl:value-of select ="'1'" />
						</xsl:if>
					</xsl:variable>
					<xsl:call-template name ="insertNumber">
						<xsl:with-param name ="number" select ="$Number"/>
						<xsl:with-param name ="startAt" select ="$startAt"/>
					</xsl:call-template>
					<style:list-level-properties text:min-label-width="0.952cm" />
					<style:text-properties style:font-family-generic="swiss" style:font-pitch="variable">
						<xsl:if test ="a:pPr/a:buFont/@typeface">
							<xsl:attribute name ="fo:font-family">
								<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buFont/@typeface"/>
							</xsl:attribute>
						</xsl:if>
						<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buFont/@typeface)">
							<xsl:attribute name ="fo:font-family">
								<xsl:value-of select ="'Arial'"/>
							</xsl:attribute>
						</xsl:if>
						<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buSzPct)">
							<xsl:attribute name ="fo:font-size">
								<xsl:value-of select ="'100%'"/>
							</xsl:attribute>
						</xsl:if>
						<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buClr">
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buClr/a:srgbClr">
								<xsl:attribute name ="fo:color">
									<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buClr/a:srgbClr/@val)"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buClr/a:schemeClr">
								<xsl:attribute name ="fo:color">
									<xsl:call-template name ="getColorCode">
										<xsl:with-param name ="color">
											<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buClr/a:schemeClr/@val"/>
										</xsl:with-param>
									</xsl:call-template >
								</xsl:attribute >
							</xsl:if>
						</xsl:if>
						<!--Code added by vijayeta, bug fix 1746350-->
						<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:buClr)">
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill">
								<xsl:if test ="a:r/a:rPr/a:solidFill/a:srgbClr">
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr">
									<xsl:attribute name ="fo:color">
										<xsl:call-template name ="getColorCode">
											<xsl:with-param name ="color">
												<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
											</xsl:with-param>
										</xsl:call-template >
									</xsl:attribute>
								</xsl:if>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill)">
								<xsl:for-each select ="document(concat('ppt/slideMasters',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
									<xsl:variable name ="levelColor" >
										<xsl:call-template name ="getLevelColor">
											<xsl:with-param name ="levelNum" select ="$newTextLvl"/>
										</xsl:call-template>
									</xsl:variable>
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',$levelColor)"/>
									</xsl:attribute>
								</xsl:for-each>
							</xsl:if>
						</xsl:if>
						<!--End of Code added by vijayeta, bug fix 1746350-->
					</style:text-properties>
				</text:list-level-style-number>
			</xsl:when>			
			<xsl:otherwise >
			  <xsl:call-template name ="getBulletsFromSlideMaster">
					<xsl:with-param name ="slideMaster" select ="$slideMaster"/>
					<xsl:with-param name ="newTextLvl" select ="$newTextLvl" />
					<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name ="getBullet_Level2">
		<xsl:param name ="slideMaster" />
		<xsl:param name ="newTextLvl" />
		<xsl:param name ="levelStyle"/>
		<xsl:choose >
			<xsl:when test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buChar">
				<text:list-style>
					<xsl:attribute name ="style:name">
						<xsl:value-of select ="$levelStyle"/>
					</xsl:attribute >
					<text:list-level-style-bullet>
						<xsl:attribute name ="text:level">
							<xsl:value-of select ="$newTextLvl"/>
						</xsl:attribute >
						<xsl:attribute name ="text:bullet-char">
							<xsl:call-template name ="insertBulletCharacter">
								<xsl:with-param name ="character" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buChar/@char" />
							</xsl:call-template>
						</xsl:attribute >
						<style:list-level-properties text:min-label-width="0.8cm"/>
						<style:text-properties style:font-charset="x-symbol">
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buFont/@typeface">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buFont[@typeface='Arial']">
									<xsl:attribute name ="fo:font-family">
										<xsl:value-of select ="'StarSymbol'"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buFont[@typeface='Arial'])">
									<xsl:attribute name ="fo:font-family">
										<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buFont/@typeface"/>
									</xsl:attribute>
								</xsl:if>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buFont/@typeface)">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="'StarSymbol'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buSzPct">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="concat((./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buSzPct/@val div 1000),'%')"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buSzPct)">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="'100%'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buClr">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buClr/a:srgbClr">
									<xsl:variable name ="color" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buClr/a:srgbClr/@val"/>
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',$color)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buClr/a:schemeClr">
									<xsl:attribute name ="fo:color">
										<xsl:call-template name ="getColorCode">
											<xsl:with-param name ="color">
												<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buClr/a:schemeClr/@val"/>
											</xsl:with-param>
										</xsl:call-template >
									</xsl:attribute >
								</xsl:if>
							</xsl:if>
							<!-- Code added by vijayeta, bug fix 1746350-->
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buClr)">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:defRPr/a:solidFill">
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:defRPr/a:solidFill/a:srgbClr">
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:defRPr/a:solidFill/a:schemeClr">
										<xsl:attribute name ="fo:color">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:attribute>
									</xsl:if>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:defRPr/a:solidFill)">
									<xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
										<xsl:variable name ="levelColor" >
											<xsl:call-template name ="getLevelColor">
												<xsl:with-param name ="levelNum" select ="$newTextLvl"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',$levelColor)"/>
										</xsl:attribute>
									</xsl:for-each>
								</xsl:if>
							</xsl:if>
							<!--End of Code added by vijayeta, bug fix 1746350-->
						</style:text-properties >
					</text:list-level-style-bullet>
				</text:list-style>
			</xsl:when>
			<xsl:when test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buAutoNum">
				<xsl:variable name ="Number" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buAutoNum/@type"/>
				<text:list-style>
					<xsl:attribute name ="style:name">
						<xsl:value-of select ="$levelStyle"/>
					</xsl:attribute >
					<text:list-level-style-number>
						<xsl:attribute name ="text:level">
							<xsl:value-of select ="$newTextLvl"/>
						</xsl:attribute >
						<xsl:variable name ="startAt">
							<xsl:if test ="a:pPr/a:buAutoNum/@startAt">
								<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buAutoNum/@startAt" />
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buAutoNum/@startAt)">
								<xsl:value-of select ="'1'" />
							</xsl:if>
						</xsl:variable>
						<xsl:call-template name ="insertNumber">
							<xsl:with-param name ="number" select ="$Number"/>
							<xsl:with-param name ="startAt" select ="$startAt"/>
						</xsl:call-template>
						<style:list-level-properties text:min-label-width="0.952cm" />
						<style:text-properties style:font-family-generic="swiss" style:font-pitch="variable">
							<xsl:if test ="a:pPr/a:buFont/@typeface">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buFont/@typeface"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buFont/@typeface)">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="'Arial'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buSzPct)">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="'100%'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buClr">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buClr/a:srgbClr">
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buClr/a:srgbClr/@val)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buClr/a:schemeClr">
									<xsl:attribute name ="fo:color">
										<xsl:call-template name ="getColorCode">
											<xsl:with-param name ="color">
												<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buClr/a:schemeClr/@val"/>
											</xsl:with-param>
										</xsl:call-template >
									</xsl:attribute >
								</xsl:if>
							</xsl:if>
							<!--Code added by vijayeta, bug fix 1746350-->
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:buClr)">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:defRPr/a:solidFill">
									<xsl:if test ="a:r/a:rPr/a:solidFill/a:srgbClr">
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:defRPr/a:solidFill/a:schemeClr">
										<xsl:attribute name ="fo:color">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:attribute>
									</xsl:if>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl2pPr/a:defRPr/a:solidFill)">
									<xsl:for-each select ="document(concat('ppt/slideMasters',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
										<xsl:variable name ="levelColor" >
											<xsl:call-template name ="getLevelColor">
												<xsl:with-param name ="levelNum" select ="$newTextLvl"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',$levelColor)"/>
										</xsl:attribute>
									</xsl:for-each>
								</xsl:if>
							</xsl:if>
							<!--End of Code added by vijayeta, bug fix 1746350-->
						</style:text-properties>
					</text:list-level-style-number>
				</text:list-style>
			</xsl:when>			
			<xsl:otherwise >
				<xsl:call-template name ="getBulletsFromSlideMaster">
					<xsl:with-param name ="slideMaster" select ="$slideMaster"/>
					<xsl:with-param name ="newTextLvl" select ="$newTextLvl" />
					<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template >
	<xsl:template name ="getBullet_Level3">
		<xsl:param name ="slideMaster" />
		<xsl:param name ="newTextLvl" />
		<xsl:param name ="levelStyle"/>
		<xsl:choose >
			<xsl:when test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl32pPr/a:buChar">
				<text:list-style>
					<xsl:attribute name ="style:name">
						<xsl:value-of select ="$levelStyle"/>
					</xsl:attribute >
					<text:list-level-style-bullet>
						<xsl:attribute name ="text:level">
							<xsl:value-of select ="$newTextLvl"/>
						</xsl:attribute >
						<xsl:attribute name ="text:bullet-char">
							<xsl:call-template name ="insertBulletCharacter">
								<xsl:with-param name ="character" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buChar/@char" />
							</xsl:call-template>
						</xsl:attribute >
						<style:list-level-properties text:min-label-width="0.8cm"/>

						<style:text-properties style:font-charset="x-symbol">
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buFont/@typeface">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buFont[@typeface='Arial']">
									<xsl:attribute name ="fo:font-family">
										<xsl:value-of select ="'StarSymbol'"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buFont[@typeface='Arial'])">
									<xsl:attribute name ="fo:font-family">
										<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buFont/@typeface"/>
									</xsl:attribute>
								</xsl:if>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buFont/@typeface)">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="'StarSymbol'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buSzPct">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="concat((./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buSzPct/@val div 1000),'%')"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buSzPct)">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="'100%'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buClr">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buClr/a:srgbClr">
									<xsl:variable name ="color" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buClr/a:srgbClr/@val"/>
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',$color)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buClr/a:schemeClr">
									<xsl:call-template name ="getColorCode">
										<xsl:with-param name ="color">
											<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buClr/a:schemeClr/@val"/>
										</xsl:with-param>
									</xsl:call-template >
								</xsl:if>
							</xsl:if>

							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buClr)">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:defRPr/a:solidFill">
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:defRPr/a:solidFill/a:srgbClr">
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:defRPr/a:solidFill/a:schemeClr">
										<xsl:attribute name ="fo:color">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:attribute>
									</xsl:if>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:defRPr/a:solidFill)">
									<xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
										<xsl:variable name ="levelColor" >
											<xsl:call-template name ="getLevelColor">
												<xsl:with-param name ="levelNum" select ="$newTextLvl"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',$levelColor)"/>
										</xsl:attribute>
									</xsl:for-each>
								</xsl:if>
							</xsl:if>

						</style:text-properties >
					</text:list-level-style-bullet>
				</text:list-style>
			</xsl:when>
			<xsl:when test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buAutoNum">
				<xsl:variable name ="Number" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buAutoNum/@type"/>
				<text:list-style>
					<xsl:attribute name ="style:name">
						<xsl:value-of select ="$levelStyle"/>
					</xsl:attribute >
					<text:list-level-style-number>
						<xsl:attribute name ="text:level">
							<xsl:value-of select ="$newTextLvl"/>
						</xsl:attribute >
						<xsl:variable name ="startAt">
							<xsl:if test ="a:pPr/a:buAutoNum/@startAt">
								<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buAutoNum/@startAt" />
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buAutoNum/@startAt)">
								<xsl:value-of select ="'1'" />
							</xsl:if>
						</xsl:variable>
						<xsl:call-template name ="insertNumber">
							<xsl:with-param name ="number" select ="$Number"/>
							<xsl:with-param name ="startAt" select ="$startAt"/>
						</xsl:call-template>
						<style:list-level-properties text:min-label-width="0.952cm" />
						<style:text-properties style:font-family-generic="swiss" style:font-pitch="variable">
							<xsl:if test ="a:pPr/a:buFont/@typeface">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buFont/@typeface"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buFont/@typeface)">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="'Arial'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buSzPct)">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="'100%'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buClr">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buClr/a:srgbClr">
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buClr/a:srgbClr/@val)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buClr/a:schemeClr">
									<xsl:call-template name ="getColorCode">
										<xsl:with-param name ="color">
											<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buClr/a:schemeClr/@val"/>
										</xsl:with-param>
									</xsl:call-template >
								</xsl:if>
							</xsl:if>

							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:buClr)">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:defRPr/a:solidFill">
									<xsl:if test ="a:r/a:rPr/a:solidFill/a:srgbClr">
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:defRPr/a:solidFill/a:schemeClr">
										<xsl:attribute name ="fo:color">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:attribute>
									</xsl:if>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl3pPr/a:defRPr/a:solidFill)">
									<xsl:for-each select ="document(concat('ppt/slideMasters',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
										<xsl:variable name ="levelColor" >
											<xsl:call-template name ="getLevelColor">
												<xsl:with-param name ="levelNum" select ="$newTextLvl"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',$levelColor)"/>
										</xsl:attribute>
									</xsl:for-each>
								</xsl:if>
							</xsl:if>
						</style:text-properties>
					</text:list-level-style-number>
				</text:list-style>
			</xsl:when>
			<xsl:otherwise >
				<xsl:call-template name ="getBulletsFromSlideMaster">
					<xsl:with-param name ="slideMaster" select ="$slideMaster"/>
					<xsl:with-param name ="newTextLvl" select ="$newTextLvl" />
					<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template >
	<xsl:template name ="getBullet_Level4">
		<xsl:param name ="slideMaster" />
		<xsl:param name ="newTextLvl" />
		<xsl:param name ="levelStyle"/>
		<xsl:choose >
			<xsl:when test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buChar">
				<text:list-style>
					<xsl:attribute name ="style:name">
						<xsl:value-of select ="$levelStyle"/>
					</xsl:attribute >
					<text:list-level-style-bullet>
						<xsl:attribute name ="text:level">
							<xsl:value-of select ="$newTextLvl"/>
						</xsl:attribute >
						<xsl:attribute name ="text:bullet-char">
							<xsl:call-template name ="insertBulletCharacter">
								<xsl:with-param name ="character" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buChar/@char" />
							</xsl:call-template>
						</xsl:attribute >
						<style:list-level-properties text:min-label-width="0.8cm"/>

						<style:text-properties style:font-charset="x-symbol">
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buFont/@typeface">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buFont[@typeface='Arial']">
									<xsl:attribute name ="fo:font-family">
										<xsl:value-of select ="'StarSymbol'"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buFont[@typeface='Arial'])">
									<xsl:attribute name ="fo:font-family">
										<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buFont/@typeface"/>
									</xsl:attribute>
								</xsl:if>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buFont/@typeface)">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="'StarSymbol'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buSzPct">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="concat((./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buSzPct/@val div 1000),'%')"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buSzPct)">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="'100%'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buClr">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buClr/a:srgbClr">
									<xsl:variable name ="color" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buClr/a:srgbClr/@val"/>
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',$color)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buClr/a:schemeClr">
									<xsl:attribute name ="fo:color">
										<xsl:call-template name ="getColorCode">
											<xsl:with-param name ="color">
												<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buClr/a:schemeClr/@val"/>
											</xsl:with-param>
										</xsl:call-template >
									</xsl:attribute >
								</xsl:if>
							</xsl:if>

							<!--Code added by vijayeta, bug fix 1746450-->

							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buClr)">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:defRPr/a:solidFill">
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:defRPr/a:solidFill/a:srgbClr">
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:defRPr/a:solidFill/a:schemeClr">
										<xsl:attribute name ="fo:color">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:attribute>
									</xsl:if>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:defRPr/a:solidFill)">
									<xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
										<xsl:variable name ="levelColor" >
											<xsl:call-template name ="getLevelColor">
												<xsl:with-param name ="levelNum" select ="$newTextLvl"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',$levelColor)"/>
										</xsl:attribute>
									</xsl:for-each>
								</xsl:if>
							</xsl:if>

							<!--End of Code added by vijayeta, bug fix 1746450-->

						</style:text-properties >
					</text:list-level-style-bullet>
				</text:list-style>
			</xsl:when>
			<xsl:when test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buAutoNum">
				<xsl:variable name ="Number" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buAutoNum/@type"/>
				<text:list-style>
					<xsl:attribute name ="style:name">
						<xsl:value-of select ="$levelStyle"/>
					</xsl:attribute >
					<text:list-level-style-number>
						<xsl:attribute name ="text:level">
							<xsl:value-of select ="$newTextLvl"/>
						</xsl:attribute >
						<xsl:variable name ="startAt">
							<xsl:if test ="a:pPr/a:buAutoNum/@startAt">
								<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buAutoNum/@startAt" />
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buAutoNum/@startAt)">
								<xsl:value-of select ="'1'" />
							</xsl:if>
						</xsl:variable>
						<xsl:call-template name ="insertNumber">
							<xsl:with-param name ="number" select ="$Number"/>
							<xsl:with-param name ="startAt" select ="$startAt"/>
						</xsl:call-template>
						<style:list-level-properties text:min-label-width="0.952cm" />
						<style:text-properties style:font-family-generic="swiss" style:font-pitch="variable">
							<xsl:if test ="a:pPr/a:buFont/@typeface">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buFont/@typeface"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buFont/@typeface)">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="'Arial'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buSzPct)">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="'100%'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buClr">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buClr/a:srgbClr">
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buClr/a:srgbClr/@val)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buClr/a:schemeClr">
									<xsl:call-template name ="getColorCode">
										<xsl:with-param name ="color">
											<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buClr/a:schemeClr/@val"/>
										</xsl:with-param>
									</xsl:call-template >
								</xsl:if>
							</xsl:if>

							<!--Code added by vijayeta, bug fix 1746450-->

							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:buClr)">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:defRPr/a:solidFill">
									<xsl:if test ="a:r/a:rPr/a:solidFill/a:srgbClr">
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:defRPr/a:solidFill/a:schemeClr">
										<xsl:attribute name ="fo:color">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:attribute>
									</xsl:if>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl4pPr/a:defRPr/a:solidFill)">
									<xsl:for-each select ="document(concat('ppt/slideMasters',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
										<xsl:variable name ="levelColor" >
											<xsl:call-template name ="getLevelColor">
												<xsl:with-param name ="levelNum" select ="$newTextLvl"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',$levelColor)"/>
										</xsl:attribute>
									</xsl:for-each>
								</xsl:if>
							</xsl:if>

							<!--End of Code added by vijayeta, bug fix 1746450-->

						</style:text-properties>
					</text:list-level-style-number>
				</text:list-style>
			</xsl:when>
			<xsl:otherwise >
				<xsl:call-template name ="getBulletsFromSlideMaster">
					<xsl:with-param name ="slideMaster" select ="$slideMaster"/>
					<xsl:with-param name ="newTextLvl" select ="$newTextLvl" />
					<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template >
	<xsl:template name ="getBullet_Level5">
		<xsl:param name ="slideMaster" />
		<xsl:param name ="newTextLvl" />
		<xsl:param name ="levelStyle"/>
		<xsl:choose >
			<xsl:when test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buChar">
				<text:list-style>
					<xsl:attribute name ="style:name">
						<xsl:value-of select ="$levelStyle"/>
					</xsl:attribute >
					<text:list-level-style-bullet>
						<xsl:attribute name ="text:level">
							<xsl:value-of select ="$newTextLvl"/>
						</xsl:attribute >
						<xsl:attribute name ="text:bullet-char">
							<xsl:call-template name ="insertBulletCharacter">
								<xsl:with-param name ="character" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buChar/@char" />
							</xsl:call-template>
						</xsl:attribute >
						<style:list-level-properties text:min-label-width="0.8cm"/>

						<style:text-properties style:font-charset="x-symbol">
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buFont/@typeface">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buFont[@typeface='Arial']">
									<xsl:attribute name ="fo:font-family">
										<xsl:value-of select ="'StarSymbol'"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buFont[@typeface='Arial'])">
									<xsl:attribute name ="fo:font-family">
										<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buFont/@typeface"/>
									</xsl:attribute>
								</xsl:if>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buFont/@typeface)">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="'StarSymbol'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buSzPct">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="concat((./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buSzPct/@val div 1000),'%')"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buSzPct)">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="'100%'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buClr">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buClr/a:srgbClr">
									<xsl:variable name ="color" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buClr/a:srgbClr/@val"/>
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',$color)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buClr/a:schemeClr">
									<xsl:attribute name ="fo:color">
										<xsl:call-template name ="getColorCode">
											<xsl:with-param name ="color">
												<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buClr/a:schemeClr/@val"/>
											</xsl:with-param>
										</xsl:call-template >
									</xsl:attribute >
								</xsl:if>
							</xsl:if>

							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buClr)">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:defRPr/a:solidFill">
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:defRPr/a:solidFill/a:srgbClr">
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:defRPr/a:solidFill/a:schemeClr">
										<xsl:attribute name ="fo:color">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:attribute>
									</xsl:if>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:defRPr/a:solidFill)">
									<xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
										<xsl:variable name ="levelColor" >
											<xsl:call-template name ="getLevelColor">
												<xsl:with-param name ="levelNum" select ="$newTextLvl"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',$levelColor)"/>
										</xsl:attribute>
									</xsl:for-each>
								</xsl:if>
							</xsl:if>



						</style:text-properties >
					</text:list-level-style-bullet>
				</text:list-style>
			</xsl:when>
			<xsl:when test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buAutoNum">
				<xsl:variable name ="Number" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buAutoNum/@type"/>
				<text:list-style>
					<xsl:attribute name ="style:name">
						<xsl:value-of select ="$levelStyle"/>
					</xsl:attribute >
					<text:list-level-style-number>
						<xsl:attribute name ="text:level">
							<xsl:value-of select ="$newTextLvl"/>
						</xsl:attribute >
						<xsl:variable name ="startAt">
							<xsl:if test ="a:pPr/a:buAutoNum/@startAt">
								<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buAutoNum/@startAt" />
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buAutoNum/@startAt)">
								<xsl:value-of select ="'1'" />
							</xsl:if>
						</xsl:variable>
						<xsl:call-template name ="insertNumber">
							<xsl:with-param name ="number" select ="$Number"/>
							<xsl:with-param name ="startAt" select ="$startAt"/>
						</xsl:call-template>
						<style:list-level-properties text:min-label-width="0.952cm" />
						<style:text-properties style:font-family-generic="swiss" style:font-pitch="variable">
							<xsl:if test ="a:pPr/a:buFont/@typeface">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buFont/@typeface"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buFont/@typeface)">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="'Arial'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buSzPct)">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="'100%'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buClr">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buClr/a:srgbClr">
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buClr/a:srgbClr/@val)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buClr/a:schemeClr">
									<xsl:attribute name ="fo:color">
										<xsl:call-template name ="getColorCode">
											<xsl:with-param name ="color">
												<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buClr/a:schemeClr/@val"/>
											</xsl:with-param>
										</xsl:call-template >
									</xsl:attribute>
								</xsl:if>
							</xsl:if>



							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:buClr)">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:defRPr/a:solidFill">
									<xsl:if test ="a:r/a:rPr/a:solidFill/a:srgbClr">
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:defRPr/a:solidFill/a:schemeClr">
										<xsl:attribute name ="fo:color">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:attribute>
									</xsl:if>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl5pPr/a:defRPr/a:solidFill)">
									<xsl:for-each select ="document(concat('ppt/slideMasters',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
										<xsl:variable name ="levelColor" >
											<xsl:call-template name ="getLevelColor">
												<xsl:with-param name ="levelNum" select ="$newTextLvl"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',$levelColor)"/>
										</xsl:attribute>
									</xsl:for-each>
								</xsl:if>
							</xsl:if>



						</style:text-properties>
					</text:list-level-style-number>
				</text:list-style>
			</xsl:when>			
			<xsl:otherwise >
				<xsl:call-template name ="getBulletsFromSlideMaster">
					<xsl:with-param name ="slideMaster" select ="$slideMaster"/>
					<xsl:with-param name ="newTextLvl" select ="$newTextLvl" />
					<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template >
	<xsl:template name ="getBullet_Level6">
		<xsl:param name ="slideMaster" />
		<xsl:param name ="newTextLvl" />
		<xsl:param name ="levelStyle"/>
		<xsl:choose >
			<xsl:when test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buChar">
				<text:list-style>
					<xsl:attribute name ="style:name">
						<xsl:value-of select ="$levelStyle"/>
					</xsl:attribute >
					<text:list-level-style-bullet>
						<xsl:attribute name ="text:level">
							<xsl:value-of select ="$newTextLvl"/>
						</xsl:attribute >
						<xsl:attribute name ="text:bullet-char">
							<xsl:call-template name ="insertBulletCharacter">
								<xsl:with-param name ="character" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buChar/@char" />
							</xsl:call-template>
						</xsl:attribute >
						<style:list-level-properties text:min-label-width="0.8cm"/>

						<style:text-properties style:font-charset="x-symbol">
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buFont/@typeface">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buFont[@typeface='Arial']">
									<xsl:attribute name ="fo:font-family">
										<xsl:value-of select ="'StarSymbol'"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buFont[@typeface='Arial'])">
									<xsl:attribute name ="fo:font-family">
										<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buFont/@typeface"/>
									</xsl:attribute>
								</xsl:if>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buFont/@typeface)">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="'StarSymbol'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buSzPct">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="concat((./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buSzPct/@val div 1000),'%')"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buSzPct)">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="'100%'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buClr">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buClr/a:srgbClr">
									<xsl:variable name ="color" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buClr/a:srgbClr/@val"/>
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',$color)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buClr/a:schemeClr">
									<xsl:call-template name ="getColorCode">
										<xsl:with-param name ="color">
											<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buClr/a:schemeClr/@val"/>
										</xsl:with-param>
									</xsl:call-template >
								</xsl:if>
							</xsl:if>



							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buClr)">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:defRPr/a:solidFill">
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:defRPr/a:solidFill/a:srgbClr">
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:defRPr/a:solidFill/a:schemeClr">
										<xsl:attribute name ="fo:color">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:attribute>
									</xsl:if>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:defRPr/a:solidFill)">
									<xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
										<xsl:variable name ="levelColor" >
											<xsl:call-template name ="getLevelColor">
												<xsl:with-param name ="levelNum" select ="$newTextLvl"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',$levelColor)"/>
										</xsl:attribute>
									</xsl:for-each>
								</xsl:if>
							</xsl:if>



						</style:text-properties >
					</text:list-level-style-bullet>
				</text:list-style>
			</xsl:when>
			<xsl:when test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buAutoNum">
				<xsl:variable name ="Number" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buAutoNum/@type"/>
				<text:list-style>
					<xsl:attribute name ="style:name">
						<xsl:value-of select ="$levelStyle"/>
					</xsl:attribute >
					<text:list-level-style-number>
						<xsl:attribute name ="text:level">
							<xsl:value-of select ="$newTextLvl"/>
						</xsl:attribute >
						<xsl:variable name ="startAt">
							<xsl:if test ="a:pPr/a:buAutoNum/@startAt">
								<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buAutoNum/@startAt" />
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buAutoNum/@startAt)">
								<xsl:value-of select ="'1'" />
							</xsl:if>
						</xsl:variable>
						<xsl:call-template name ="insertNumber">
							<xsl:with-param name ="number" select ="$Number"/>
							<xsl:with-param name ="startAt" select ="$startAt"/>
						</xsl:call-template>
						<style:list-level-properties text:min-label-width="0.962cm" />
						<style:text-properties style:font-family-generic="swiss" style:font-pitch="variable">
							<xsl:if test ="a:pPr/a:buFont/@typeface">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buFont/@typeface"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buFont/@typeface)">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="'Arial'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buSzPct)">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="'100%'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buClr">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buClr/a:srgbClr">
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buClr/a:srgbClr/@val)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buClr/a:schemeClr">
									<xsl:attribute name ="fo:color">
										<xsl:call-template name ="getColorCode">
											<xsl:with-param name ="color">
												<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buClr/a:schemeClr/@val"/>
											</xsl:with-param>
										</xsl:call-template >
									</xsl:attribute >
								</xsl:if>
							</xsl:if>



							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:buClr)">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:defRPr/a:solidFill">
									<xsl:if test ="a:r/a:rPr/a:solidFill/a:srgbClr">
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:defRPr/a:solidFill/a:schemeClr">
										<xsl:attribute name ="fo:color">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:attribute>
									</xsl:if>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl6pPr/a:defRPr/a:solidFill)">
									<xsl:for-each select ="document(concat('ppt/slideMasters',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
										<xsl:variable name ="levelColor" >
											<xsl:call-template name ="getLevelColor">
												<xsl:with-param name ="levelNum" select ="$newTextLvl"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',$levelColor)"/>
										</xsl:attribute>
									</xsl:for-each>
								</xsl:if>
							</xsl:if>



						</style:text-properties>
					</text:list-level-style-number>
				</text:list-style>
			</xsl:when>      
			<xsl:otherwise >
				<xsl:call-template name ="getBulletsFromSlideMaster">
					<xsl:with-param name ="slideMaster" select ="$slideMaster"/>
					<xsl:with-param name ="newTextLvl" select ="$newTextLvl" />
					<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template >
	<xsl:template name ="getBullet_Level7">
		<xsl:param name ="slideMaster" />
		<xsl:param name ="newTextLvl" />
		<xsl:param name ="levelStyle"/>
		<xsl:choose >
			<xsl:when test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buChar">
				<text:list-style>
					<xsl:attribute name ="style:name">
						<xsl:value-of select ="$levelStyle"/>
					</xsl:attribute >
					<text:list-level-style-bullet>
						<xsl:attribute name ="text:level">
							<xsl:value-of select ="$newTextLvl"/>
						</xsl:attribute >
						<xsl:attribute name ="text:bullet-char">
							<xsl:call-template name ="insertBulletCharacter">
								<xsl:with-param name ="character" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buChar/@char" />
							</xsl:call-template>
						</xsl:attribute >
						<style:list-level-properties text:min-label-width="0.8cm"/>
						<style:text-properties style:font-charset="x-symbol">
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buFont/@typeface">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buFont[@typeface='Arial']">
									<xsl:attribute name ="fo:font-family">
										<xsl:value-of select ="'StarSymbol'"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buFont[@typeface='Arial'])">
									<xsl:attribute name ="fo:font-family">
										<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buFont/@typeface"/>
									</xsl:attribute>
								</xsl:if>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buFont/@typeface)">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="'StarSymbol'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buSzPct">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="concat((./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buSzPct/@val div 1000),'%')"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buSzPct)">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="'100%'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buClr">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buClr/a:srgbClr">
									<xsl:variable name ="color" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buClr/a:srgbClr/@val"/>
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',$color)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buClr/a:schemeClr">
									<xsl:attribute name ="fo:color">
										<xsl:call-template name ="getColorCode">
											<xsl:with-param name ="color">
												<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buClr/a:schemeClr/@val"/>
											</xsl:with-param>
										</xsl:call-template >
									</xsl:attribute >
								</xsl:if>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buClr)">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:defRPr/a:solidFill">
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:defRPr/a:solidFill/a:srgbClr">
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:defRPr/a:solidFill/a:schemeClr">
										<xsl:attribute name ="fo:color">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:attribute>
									</xsl:if>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:defRPr/a:solidFill)">
									<xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
										<xsl:variable name ="levelColor" >
											<xsl:call-template name ="getLevelColor">
												<xsl:with-param name ="levelNum" select ="$newTextLvl"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',$levelColor)"/>
										</xsl:attribute>
									</xsl:for-each>
								</xsl:if>
							</xsl:if>
						</style:text-properties >
					</text:list-level-style-bullet>
				</text:list-style>
			</xsl:when>
			<xsl:when test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buAutoNum">
				<xsl:variable name ="Number" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buAutoNum/@type"/>
				<text:list-style>
					<xsl:attribute name ="style:name">
						<xsl:value-of select ="$levelStyle"/>
					</xsl:attribute >
					<text:list-level-style-number>
						<xsl:attribute name ="text:level">
							<xsl:value-of select ="$newTextLvl"/>
						</xsl:attribute >
						<xsl:variable name ="startAt">
							<xsl:if test ="a:pPr/a:buAutoNum/@startAt">
								<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buAutoNum/@startAt" />
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buAutoNum/@startAt)">
								<xsl:value-of select ="'1'" />
							</xsl:if>
						</xsl:variable>
						<xsl:call-template name ="insertNumber">
							<xsl:with-param name ="number" select ="$Number"/>
							<xsl:with-param name ="startAt" select ="$startAt"/>
						</xsl:call-template>
						<style:list-level-properties text:min-label-width="0.972cm" />
						<style:text-properties style:font-family-generic="swiss" style:font-pitch="variable">
							<xsl:if test ="a:pPr/a:buFont/@typeface">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buFont/@typeface"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buFont/@typeface)">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="'Arial'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buSzPct)">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="'100%'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buClr">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buClr/a:srgbClr">
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buClr/a:srgbClr/@val)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buClr/a:schemeClr">
									<xsl:attribute name ="fo:color">
										<xsl:call-template name ="getColorCode">
											<xsl:with-param name ="color">
												<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buClr/a:schemeClr/@val"/>
											</xsl:with-param>
										</xsl:call-template >
									</xsl:attribute >
								</xsl:if>
							</xsl:if>



							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:buClr)">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:defRPr/a:solidFill">
									<xsl:if test ="a:r/a:rPr/a:solidFill/a:srgbClr">
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:defRPr/a:solidFill/a:schemeClr">
										<xsl:attribute name ="fo:color">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:attribute>
									</xsl:if>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl7pPr/a:defRPr/a:solidFill)">
									<xsl:for-each select ="document(concat('ppt/slideMasters',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
										<xsl:variable name ="levelColor" >
											<xsl:call-template name ="getLevelColor">
												<xsl:with-param name ="levelNum" select ="$newTextLvl"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',$levelColor)"/>
										</xsl:attribute>
									</xsl:for-each>
								</xsl:if>
							</xsl:if>
						</style:text-properties>
					</text:list-level-style-number>
				</text:list-style>
			</xsl:when>			
			<xsl:otherwise >
				<xsl:call-template name ="getBulletsFromSlideMaster">
					<xsl:with-param name ="slideMaster" select ="$slideMaster"/>
					<xsl:with-param name ="newTextLvl" select ="$newTextLvl" />
					<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template >
	<xsl:template name ="getBullet_Level8">
		<xsl:param name ="slideMaster" />
		<xsl:param name ="newTextLvl" />
		<xsl:param name ="levelStyle"/>
		<xsl:choose >
			<xsl:when test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buChar">
				<text:list-style>
					<xsl:attribute name ="style:name">
						<xsl:value-of select ="$levelStyle"/>
					</xsl:attribute >
					<text:list-level-style-bullet>
						<xsl:attribute name ="text:level">
							<xsl:value-of select ="$newTextLvl"/>
						</xsl:attribute >
						<xsl:attribute name ="text:bullet-char">
							<xsl:call-template name ="insertBulletCharacter">
								<xsl:with-param name ="character" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buChar/@char" />
							</xsl:call-template>
						</xsl:attribute >
						<style:list-level-properties text:min-label-width="0.8cm"/>

						<style:text-properties style:font-charset="x-symbol">
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buFont/@typeface">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buFont[@typeface='Arial']">
									<xsl:attribute name ="fo:font-family">
										<xsl:value-of select ="'StarSymbol'"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buFont[@typeface='Arial'])">
									<xsl:attribute name ="fo:font-family">
										<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buFont/@typeface"/>
									</xsl:attribute>
								</xsl:if>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buFont/@typeface)">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="'StarSymbol'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buSzPct">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="concat((./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buSzPct/@val div 1000),'%')"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buSzPct)">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="'100%'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buClr">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buClr/a:srgbClr">
									<xsl:variable name ="color" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buClr/a:srgbClr/@val"/>
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',$color)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buClr/a:schemeClr">
									<xsl:attribute name ="fo:color">
										<xsl:call-template name ="getColorCode">
											<xsl:with-param name ="color">
												<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buClr/a:schemeClr/@val"/>
											</xsl:with-param>
										</xsl:call-template >
									</xsl:attribute >
								</xsl:if>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buClr)">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:defRPr/a:solidFill">
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:defRPr/a:solidFill/a:srgbClr">
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:defRPr/a:solidFill/a:schemeClr">
										<xsl:attribute name ="fo:color">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:attribute>
									</xsl:if>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:defRPr/a:solidFill)">
									<xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
										<xsl:variable name ="levelColor" >
											<xsl:call-template name ="getLevelColor">
												<xsl:with-param name ="levelNum" select ="$newTextLvl"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',$levelColor)"/>
										</xsl:attribute>
									</xsl:for-each>
								</xsl:if>
							</xsl:if>



						</style:text-properties >
					</text:list-level-style-bullet>
				</text:list-style>
			</xsl:when>
			<xsl:when test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buAutoNum">
				<xsl:variable name ="Number" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buAutoNum/@type"/>
				<text:list-style>
					<xsl:attribute name ="style:name">
						<xsl:value-of select ="$levelStyle"/>
					</xsl:attribute >
					<text:list-level-style-number>
						<xsl:attribute name ="text:level">
							<xsl:value-of select ="$newTextLvl"/>
						</xsl:attribute >
						<xsl:variable name ="startAt">
							<xsl:if test ="a:pPr/a:buAutoNum/@startAt">
								<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buAutoNum/@startAt" />
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buAutoNum/@startAt)">
								<xsl:value-of select ="'1'" />
							</xsl:if>
						</xsl:variable>
						<xsl:call-template name ="insertNumber">
							<xsl:with-param name ="number" select ="$Number"/>
							<xsl:with-param name ="startAt" select ="$startAt"/>
						</xsl:call-template>
						<style:list-level-properties text:min-label-width="0.982cm" />
						<style:text-properties style:font-family-generic="swiss" style:font-pitch="variable">
							<xsl:if test ="a:pPr/a:buFont/@typeface">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buFont/@typeface"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buFont/@typeface)">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="'Arial'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buSzPct)">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="'100%'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buClr">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buClr/a:srgbClr">
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buClr/a:srgbClr/@val)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buClr/a:schemeClr">
									<xsl:attribute name ="fo:color">
										<xsl:call-template name ="getColorCode">
											<xsl:with-param name ="color">
												<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buClr/a:schemeClr/@val"/>
											</xsl:with-param>
										</xsl:call-template >
									</xsl:attribute >
								</xsl:if>
							</xsl:if>



							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:buClr)">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:defRPr/a:solidFill">
									<xsl:if test ="a:r/a:rPr/a:solidFill/a:srgbClr">
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:defRPr/a:solidFill/a:schemeClr">
										<xsl:attribute name ="fo:color">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:attribute>
									</xsl:if>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl8pPr/a:defRPr/a:solidFill)">
									<xsl:for-each select ="document(concat('ppt/slideMasters',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
										<xsl:variable name ="levelColor" >
											<xsl:call-template name ="getLevelColor">
												<xsl:with-param name ="levelNum" select ="$newTextLvl"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',$levelColor)"/>
										</xsl:attribute>
									</xsl:for-each>
								</xsl:if>
							</xsl:if>
						</style:text-properties>
					</text:list-level-style-number>
				</text:list-style>
			</xsl:when>		
			<xsl:otherwise >
				<xsl:call-template name ="getBulletsFromSlideMaster">
					<xsl:with-param name ="slideMaster" select ="$slideMaster"/>
					<xsl:with-param name ="newTextLvl" select ="$newTextLvl" />
					<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template >
	<xsl:template name ="getBullet_Level9">
		<xsl:param name ="slideMaster" />
		<xsl:param name ="newTextLvl" />
		<xsl:param name ="levelStyle"/>
		<xsl:choose >
			<xsl:when test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buChar">
				<text:list-style>
					<xsl:attribute name ="style:name">
						<xsl:value-of select ="$levelStyle"/>
					</xsl:attribute >
					<text:list-level-style-bullet>
						<xsl:attribute name ="text:level">
							<xsl:value-of select ="$newTextLvl"/>
						</xsl:attribute >
						<xsl:attribute name ="text:bullet-char">
							<xsl:call-template name ="insertBulletCharacter">
								<xsl:with-param name ="character" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buChar/@char" />
							</xsl:call-template>
						</xsl:attribute >
						<style:list-level-properties text:min-label-width="0.9cm"/>

						<style:text-properties style:font-charset="x-symbol">
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buFont/@typeface">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buFont[@typeface='Arial']">
									<xsl:attribute name ="fo:font-family">
										<xsl:value-of select ="'StarSymbol'"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buFont[@typeface='Arial'])">
									<xsl:attribute name ="fo:font-family">
										<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buFont/@typeface"/>
									</xsl:attribute>
								</xsl:if>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buFont/@typeface)">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="'StarSymbol'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buSzPct">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="concat((./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buSzPct/@val div 1000),'%')"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buSzPct)">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="'100%'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buClr">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buClr/a:srgbClr">
									<xsl:variable name ="color" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buClr/a:srgbClr/@val"/>
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',$color)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buClr/a:schemeClr">
									<xsl:attribute name ="fo:color">
										<xsl:call-template name ="getColorCode">
											<xsl:with-param name ="color">
												<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buClr/a:schemeClr/@val"/>
											</xsl:with-param>
										</xsl:call-template >
									</xsl:attribute >
								</xsl:if>
							</xsl:if>

							<!--Code added by vijayeta, bug fix 1999990-->
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buClr)">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:defRPr/a:solidFill">
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:defRPr/a:solidFill/a:srgbClr">
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:defRPr/a:solidFill/a:schemeClr">
										<xsl:attribute name ="fo:color">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:attribute>
									</xsl:if>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:defRPr/a:solidFill)">
									<xsl:for-each select ="document(concat('ppt/slideMasters/',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
										<xsl:variable name ="levelColor" >
											<xsl:call-template name ="getLevelColor">
												<xsl:with-param name ="levelNum" select ="$newTextLvl"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',$levelColor)"/>
										</xsl:attribute>
									</xsl:for-each>
								</xsl:if>
							</xsl:if>

							<!--End of Code added by vijayeta, bug fix 1999990-->

						</style:text-properties >
					</text:list-level-style-bullet>
				</text:list-style>
			</xsl:when>
			<xsl:when test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buAutoNum">
				<xsl:variable name ="Number" select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buAutoNum/@type"/>
				<text:list-style>
					<xsl:attribute name ="style:name">
						<xsl:value-of select ="$levelStyle"/>
					</xsl:attribute >
					<text:list-level-style-number>
						<xsl:attribute name ="text:level">
							<xsl:value-of select ="$newTextLvl"/>
						</xsl:attribute >
						<xsl:variable name ="startAt">
							<xsl:if test ="a:pPr/a:buAutoNum/@startAt">
								<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buAutoNum/@startAt" />
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buAutoNum/@startAt)">
								<xsl:value-of select ="'1'" />
							</xsl:if>
						</xsl:variable>
						<xsl:call-template name ="insertNumber">
							<xsl:with-param name ="number" select ="$Number"/>
							<xsl:with-param name ="startAt" select ="$startAt"/>
						</xsl:call-template>
						<style:list-level-properties text:min-label-width="0.992cm" />
						<style:text-properties style:font-family-generic="swiss" style:font-pitch="variable">
							<xsl:if test ="a:pPr/a:buFont/@typeface">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buFont/@typeface"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buFont/@typeface)">
								<xsl:attribute name ="fo:font-family">
									<xsl:value-of select ="'Arial'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buSzPct)">
								<xsl:attribute name ="fo:font-size">
									<xsl:value-of select ="'100%'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buClr">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buClr/a:srgbClr">
									<xsl:attribute name ="fo:color">
										<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buClr/a:srgbClr/@val)"/>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buClr/a:schemeClr">
									<xsl:call-template name ="getColorCode">
										<xsl:with-param name ="color">
											<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buClr/a:schemeClr/@val"/>
										</xsl:with-param>
									</xsl:call-template >
								</xsl:if>
							</xsl:if>
							<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:buClr)">
								<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:defRPr/a:solidFill">
									<xsl:if test ="a:r/a:rPr/a:solidFill/a:srgbClr">
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
										</xsl:attribute>
									</xsl:if>
									<xsl:if test ="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:defRPr/a:solidFill/a:schemeClr">
										<xsl:attribute name ="fo:color">
											<xsl:call-template name ="getColorCode">
												<xsl:with-param name ="color">
													<xsl:value-of select="./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:defRPr/a:solidFill/a:schemeClr/@val"/>
												</xsl:with-param>
											</xsl:call-template >
										</xsl:attribute>
									</xsl:if>
								</xsl:if>
								<xsl:if test ="not(./parent::node()/parent::node()/parent::node()/p:txBody/a:lstStyle/a:lvl9pPr/a:defRPr/a:solidFill)">
									<xsl:for-each select ="document(concat('ppt/slideMasters',$slideMaster))//p:sldMaster/p:txStyles/p:bodyStyle">
										<xsl:variable name ="levelColor" >
											<xsl:call-template name ="getLevelColor">
												<xsl:with-param name ="levelNum" select ="$newTextLvl"/>
											</xsl:call-template>
										</xsl:variable>
										<xsl:attribute name ="fo:color">
											<xsl:value-of select ="concat('#',$levelColor)"/>
										</xsl:attribute>
									</xsl:for-each>
								</xsl:if>
							</xsl:if>



						</style:text-properties>
					</text:list-level-style-number>
				</text:list-style>
			</xsl:when>			
			<xsl:otherwise >
				<xsl:call-template name ="getBulletsFromSlideMaster">
					<xsl:with-param name ="slideMaster" select ="$slideMaster"/>
					<xsl:with-param name ="newTextLvl" select ="$newTextLvl" />
					<xsl:with-param name ="levelStyle" select ="$levelStyle"/>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template >
</xsl:stylesheet>

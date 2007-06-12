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
    <xsl:if test ="not(./a:pPr/@lvl)">
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
    <xsl:if test ="./a:p/a:pPr/@lvl">
      <xsl:for-each select ="./a:p">
        <xsl:if test ="a:pPr/@lvl">
          <xsl:variable name ="levelStyle">
            <xsl:value-of select ="concat($listStyleName,'lvl',a:pPr/@lvl)"/>
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
                    <xsl:choose>
                      <xsl:when test="a:pPr/a:buChar/@char= 'Ø'">
                        <xsl:value-of select ="'➢'"/>
                      </xsl:when>
                      <xsl:when test="a:pPr/a:buChar/@char= 'o'">
                        <xsl:value-of select ="'○'"/>
                      </xsl:when>
                      <xsl:when test="a:pPr/a:buChar/@char= '§'">
                        <xsl:value-of select ="'■'"/>
                      </xsl:when>
                      <xsl:when test="a:pPr/a:buChar/@char= 'q'">
                        <xsl:value-of select ="''"/>
                      </xsl:when>
                      <xsl:when test="a:pPr/a:buChar/@char= 'ü'">
                        <xsl:value-of select ="'✔'"/>
                      </xsl:when>
                      <xsl:otherwise>•</xsl:otherwise>
                    </xsl:choose >
                  </xsl:attribute >
                  <style:list-level-properties text:min-label-width="0.8cm">
                    <xsl:attribute name ="text:space-before">
                      <xsl:for-each select ="document('ppt/slideMasters/slideMaster1.xml')//p:sldMaster/p:txStyles/p:bodyStyle">
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
                        <xsl:value-of select ="'45%'"/>
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
                  </style:text-properties >
                </text:list-level-style-bullet>
              </xsl:when>
              <xsl:when test ="a:pPr/a:buAutoNum">
                <text:list-level-style-number>
                  <xsl:attribute name ="text:level">
                    <xsl:value-of select ="$textLevel + 1"/>
                  </xsl:attribute >
                  <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'arabicPeriod')]">
                    <xsl:attribute name ="style:num-format" >
                      <xsl:value-of  select ="'1'"/>
                    </xsl:attribute >
                    <xsl:attribute name ="style:num-suffix">
                      <xsl:value-of select ="'.'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'arabicParenR')]">
                    <xsl:attribute name ="style:num-format" >
                      <xsl:value-of  select ="'1'"/>
                    </xsl:attribute >
                    <xsl:attribute name ="style:num-suffix">
                      <xsl:value-of select ="')'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'arabicParenBoth')]">
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
                  <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'alphaUcPeriod')]">
                    <xsl:attribute name ="style:num-format" >
                      <xsl:value-of  select ="'A'"/>
                    </xsl:attribute >
                    <xsl:attribute name ="style:num-suffix">
                      <xsl:value-of select ="'.'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'alphaUcParenR')]">
                    <xsl:attribute name ="style:num-format" >
                      <xsl:value-of  select ="'A'"/>
                    </xsl:attribute >
                    <xsl:attribute name ="style:num-suffix">
                      <xsl:value-of select ="')'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'alphaUcParenBoth')]">
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
                  <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'alphaLcPeriod')]">
                    <xsl:attribute name ="style:num-format" >
                      <xsl:value-of  select ="'a'"/>
                    </xsl:attribute >
                    <xsl:attribute name ="style:num-suffix">
                      <xsl:value-of select ="'.'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'alphaLcParenR')]">
                    <xsl:attribute name ="style:num-format" >
                      <xsl:value-of  select ="'a'"/>
                    </xsl:attribute >
                    <xsl:attribute name ="style:num-suffix">
                      <xsl:value-of select ="')'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'alphaLcParenBoth')]">
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
                  <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'romanUcPeriod')]">
                    <xsl:attribute name ="style:num-format" >
                      <xsl:value-of  select ="'I'"/>
                    </xsl:attribute >
                    <xsl:attribute name ="style:num-suffix">
                      <xsl:value-of select ="'.'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'romanLcPeriod')]">
                    <xsl:attribute name ="style:num-format" >
                      <xsl:value-of  select ="'i'"/>
                    </xsl:attribute >
                    <xsl:attribute name ="style:num-suffix">
                      <xsl:value-of select ="'.'"/>
                    </xsl:attribute>
                    <!-- start at value-->
                    <xsl:if test ="@startAt">
                      <xsl:attribute name ="text:start-value">
                        <xsl:value-of select ="a:pPr/a:buAutoNum[@startAt]"/>
                      </xsl:attribute>
                    </xsl:if>
                  </xsl:if >
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
                        <xsl:value-of select ="'45%'"/>
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
                <xsl:for-each select ="document('ppt/slideMasters/slideMaster1.xml')//p:sldMaster/p:txStyles/p:bodyStyle/child::node()[1]">
                  <xsl:if test ="a:buChar">
                    <text:list-level-style-bullet>
                      <xsl:attribute name ="text:level">
                        <xsl:value-of select ="$textLevel + 1"/>
                      </xsl:attribute >
                      <xsl:attribute name ="text:bullet-char">
                        <xsl:choose>
                          <xsl:when test="a:buChar/@char= 'Ø'">
                            <xsl:value-of select ="'➢'"/>
                          </xsl:when>
                          <xsl:when test="a:buChar/@char= 'o'">
                            <xsl:value-of select ="'○'"/>
                          </xsl:when>
                          <xsl:when test="a:buChar/@char= '§'">
                            <xsl:value-of select ="'■'"/>
                          </xsl:when>
                          <xsl:when test="a:buChar/@char= 'q'">
                            <xsl:value-of select ="''"/>
                          </xsl:when>
                          <xsl:when test="a:buChar/@char= 'ü'">
                            <xsl:value-of select ="'✔'"/>
                          </xsl:when>
                          <xsl:when test="a:buChar/@char = '-'">
                            <xsl:value-of select ="'–'"/>
                          </xsl:when>
                          <xsl:when test="a:buChar/@char = '»'">
                            <xsl:value-of select ="'»'"/>
                          </xsl:when>
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
                            <xsl:value-of select ="'45%'"/>
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
                        <!-- start at value-->
                        <xsl:if test ="@startAt">
                          <xsl:attribute name ="text:start-value">
                            <xsl:value-of select ="a:buAutoNum[@startAt]"/>
                          </xsl:attribute>
                        </xsl:if>
                      </xsl:if >
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
                            <xsl:value-of select ="'45%'"/>
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
        <xsl:if test ="not(a:pPr/@lvl)">
          <text:list-style>
            <xsl:attribute name ="style:name">
              <xsl:value-of select ="$listStyleName"/>
            </xsl:attribute >
            <xsl:if test ="a:pPr/a:buChar">
              <text:list-level-style-bullet>
                <xsl:attribute name ="text:level">
                  <xsl:value-of select ="1"/>
                </xsl:attribute >
                <xsl:attribute name ="text:bullet-char">
                  <xsl:choose>
                    <xsl:when test="a:pPr/a:buChar/@char= 'Ø'">
                      <xsl:value-of select ="'➢'"/>
                    </xsl:when>
                    <xsl:when test="a:pPr/a:buChar/@char= 'o'">
                      <xsl:value-of select ="'○'"/>
                    </xsl:when>
                    <xsl:when test="a:pPr/a:buChar/@char= '§'">
                      <xsl:value-of select ="'■'"/>
                    </xsl:when>
                    <xsl:when test="a:pPr/a:buChar/@char= 'q'">
                      <xsl:value-of select ="''"/>
                    </xsl:when>
                    <xsl:when test="a:pPr/a:buChar/@char= 'ü'">
                      <xsl:value-of select ="'✔'"/>
                    </xsl:when>
                    <xsl:otherwise>•</xsl:otherwise>
                  </xsl:choose >
                </xsl:attribute >
                <style:list-level-properties text:min-label-width="0.8cm">
                  <xsl:attribute name ="text:space-before">
                    <xsl:for-each select ="document('ppt/slideMasters/slideMaster1.xml')//p:sldMaster/p:txStyles/p:bodyStyle">
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
                </style:text-properties >
              </text:list-level-style-bullet>
            </xsl:if>
            <xsl:if test ="a:pPr/a:buAutoNum">
              <text:list-level-style-number>
                <xsl:attribute name ="text:level">
                  <xsl:value-of select ="'1'"/>
                </xsl:attribute >
                <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'arabicPeriod')]">
                  <xsl:attribute name ="style:num-format" >
                    <xsl:value-of  select ="'1'"/>
                  </xsl:attribute >
                  <xsl:attribute name ="style:num-suffix">
                    <xsl:value-of select ="'.'"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'arabicParenR')]">
                  <xsl:attribute name ="style:num-format" >
                    <xsl:value-of  select ="'1'"/>
                  </xsl:attribute >
                  <xsl:attribute name ="style:num-suffix">
                    <xsl:value-of select ="')'"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'arabicParenBoth')]">
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
                <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'alphaUcPeriod')]">
                  <xsl:attribute name ="style:num-format" >
                    <xsl:value-of  select ="'A'"/>
                  </xsl:attribute >
                  <xsl:attribute name ="style:num-suffix">
                    <xsl:value-of select ="'.'"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'alphaUcParenR')]">
                  <xsl:attribute name ="style:num-format" >
                    <xsl:value-of  select ="'A'"/>
                  </xsl:attribute >
                  <xsl:attribute name ="style:num-suffix">
                    <xsl:value-of select ="')'"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'alphaUcParenBoth')]">
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
                <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'alphaLcPeriod')]">
                  <xsl:attribute name ="style:num-format" >
                    <xsl:value-of  select ="'a'"/>
                  </xsl:attribute >
                  <xsl:attribute name ="style:num-suffix">
                    <xsl:value-of select ="'.'"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'alphaLcParenR')]">
                  <xsl:attribute name ="style:num-format" >
                    <xsl:value-of  select ="'a'"/>
                  </xsl:attribute >
                  <xsl:attribute name ="style:num-suffix">
                    <xsl:value-of select ="')'"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'alphaLcParenBoth')]">
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
                <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'romanUcPeriod')]">
                  <xsl:attribute name ="style:num-format" >
                    <xsl:value-of  select ="'I'"/>
                  </xsl:attribute >
                  <xsl:attribute name ="style:num-suffix">
                    <xsl:value-of select ="'.'"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test ="a:pPr/a:buAutoNum/@type[contains(.,'romanLcPeriod')]">
                  <xsl:attribute name ="style:num-format" >
                    <xsl:value-of  select ="'i'"/>
                  </xsl:attribute >
                  <xsl:attribute name ="style:num-suffix">
                    <xsl:value-of select ="'.'"/>
                  </xsl:attribute>
                  <!-- start at value-->
                  <xsl:if test ="@startAt">
                    <xsl:attribute name ="text:start-value">
                      <xsl:value-of select ="a:pPr/a:buAutoNum[@startAt]"/>
                    </xsl:attribute>
                  </xsl:if>
                </xsl:if >
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
      </xsl:for-each>
    </xsl:if>
    <!--End of Condition if levels are present-->
    <!--Condition if levels are not present-->
    <xsl:if test ="not(./a:p/a:pPr/@lvl)">
      <xsl:if test ="./a:p/a:pPr/a:buChar">
        <text:list-style>
          <xsl:attribute name ="style:name">
            <xsl:value-of select ="$listStyleName"/>
          </xsl:attribute >
          <text:list-level-style-bullet>
            <xsl:attribute name ="text:level">
              <xsl:value-of select ="1"/>
            </xsl:attribute >
            <xsl:attribute name ="text:bullet-char">
              <xsl:choose>
                <xsl:when test="./a:p/a:pPr/a:buChar/@char= 'Ø'">
                  <xsl:value-of select ="'➢'"/>
                </xsl:when>
                <xsl:when test="./a:p/a:pPr/a:buChar/@char= 'o'">
                  <xsl:value-of select ="'○'"/>
                </xsl:when>
                <xsl:when test="./a:p/a:pPr/a:buChar/@char= '§'">
                  <xsl:value-of select ="'■'"/>
                </xsl:when>
                <xsl:when test="./a:p/a:pPr/a:buChar/@char= 'q'">
                  <xsl:value-of select ="''"/>
                </xsl:when>
                <xsl:when test="./a:p/a:pPr/a:buChar/@char= 'ü'">
                  <xsl:value-of select ="'✔'"/>
                </xsl:when>
                <xsl:otherwise>•</xsl:otherwise>
              </xsl:choose >
            </xsl:attribute >
            <style:list-level-properties text:min-label-width="0.952cm" />
            <style:text-properties style:font-charset="x-symbol">
              <xsl:if test ="./a:p/a:pPr/a:buFont/@typeface">
                <xsl:if test ="./a:p/a:pPr/a:buFont[@typeface='Arial']">
                  <xsl:attribute name ="fo:font-family">
                    <xsl:value-of select ="'StarSymbol'"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test ="not(./a:p/a:pPr/a:buFont[@typeface='Arial'])">
                  <xsl:attribute name ="fo:font-family">
                    <xsl:value-of select ="./a:p/a:pPr/a:buFont/@typeface"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:if>
              <xsl:if test ="not(./a:p/a:pPr/a:buFont/@typeface)">
                <xsl:attribute name ="fo:font-family">
                  <xsl:value-of select ="'StarSymbol'"/>
                </xsl:attribute>
              </xsl:if>              
              <xsl:if test ="./a:p/a:pPr/a:buSzPct">
                <xsl:attribute name ="fo:font-size">
                  <xsl:value-of select ="concat((./a:p/a:pPr/a:buSzPct/@val div 1000),'%')"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test ="not(./a:p/a:pPr/a:buSzPct)">
                <xsl:attribute name ="fo:font-size">
                  <xsl:value-of select ="'100%'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test ="./a:p/a:p/a:pPr/a:buClr">
                <xsl:if test ="./a:p/a:pPr/a:buClr/a:srgbClr">
                  <xsl:variable name ="color" select ="./a:p/a:pPr/a:buClr/a:srgbClr/@val"/>
                  <xsl:attribute name ="fo:color">
                    <xsl:value-of select ="concat('#',$color)"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test ="./a:p/a:pPr/a:buClr/a:schemeClr">
                  <xsl:call-template name ="getColorCode">
                    <xsl:with-param name ="color">
                      <xsl:value-of select="./a:p/a:pPr/a:buClr/a:schemeClr/@val"/>
                    </xsl:with-param>
                  </xsl:call-template >
                </xsl:if>
              </xsl:if>
            </style:text-properties >
          </text:list-level-style-bullet>
        </text:list-style>
      </xsl:if>
      <xsl:if test ="./a:p/a:pPr/a:buAutoNum">
        <text:list-style>
          <xsl:attribute name ="style:name">
            <xsl:value-of select ="$listStyleName"/>
          </xsl:attribute >
          <text:list-level-style-number>
            <xsl:attribute name ="text:level">
              <xsl:value-of select ="'1'"/>
            </xsl:attribute >
            <xsl:if test ="./a:p/a:pPr/a:buAutoNum/@type[contains(.,'arabicPeriod')]">
              <xsl:attribute name ="style:num-format" >
                <xsl:value-of  select ="'1'"/>
              </xsl:attribute >
              <xsl:attribute name ="style:num-suffix">
                <xsl:value-of select ="'.'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="./a:p/a:pPr/a:buAutoNum/@type[contains(.,'arabicParenR')]">
              <xsl:attribute name ="style:num-format" >
                <xsl:value-of  select ="'1'"/>
              </xsl:attribute >
              <xsl:attribute name ="style:num-suffix">
                <xsl:value-of select ="')'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="./a:p/a:pPr/a:buAutoNum/@type[contains(.,'arabicParenBoth')]">
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
            <xsl:if test ="./a:p/a:pPr/a:buAutoNum/@type[contains(.,'alphaUcPeriod')]">
              <xsl:attribute name ="style:num-format" >
                <xsl:value-of  select ="'A'"/>
              </xsl:attribute >
              <xsl:attribute name ="style:num-suffix">
                <xsl:value-of select ="'.'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="./a:p/a:pPr/a:buAutoNum/@type[contains(.,'alphaUcParenR')]">
              <xsl:attribute name ="style:num-format" >
                <xsl:value-of  select ="'A'"/>
              </xsl:attribute >
              <xsl:attribute name ="style:num-suffix">
                <xsl:value-of select ="')'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="./a:p/a:pPr/a:buAutoNum/@type[contains(.,'alphaUcParenBoth')]">
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
            <xsl:if test ="./a:p/a:pPr/a:buAutoNum/@type[contains(.,'alphaLcPeriod')]">
              <xsl:attribute name ="style:num-format" >
                <xsl:value-of  select ="'a'"/>
              </xsl:attribute >
              <xsl:attribute name ="style:num-suffix">
                <xsl:value-of select ="'.'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="./a:p/a:pPr/a:buAutoNum/@type[contains(.,'alphaLcParenR')]">
              <xsl:attribute name ="style:num-format" >
                <xsl:value-of  select ="'a'"/>
              </xsl:attribute >
              <xsl:attribute name ="style:num-suffix">
                <xsl:value-of select ="')'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="./a:p/a:pPr/a:buAutoNum/@type[contains(.,'alphaLcParenBoth')]">
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
            <xsl:if test ="./a:p/a:pPr/a:buAutoNum/@type[contains(.,'romanUcPeriod')]">
              <xsl:attribute name ="style:num-format" >
                <xsl:value-of  select ="'I'"/>
              </xsl:attribute >
              <xsl:attribute name ="style:num-suffix">
                <xsl:value-of select ="'.'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="./a:p/a:pPr/a:buAutoNum/@type[contains(.,'romanLcPeriod')]">
              <xsl:attribute name ="style:num-format" >
                <xsl:value-of  select ="'i'"/>
              </xsl:attribute >
              <xsl:attribute name ="style:num-suffix">
                <xsl:value-of select ="'.'"/>
              </xsl:attribute>
              <!-- start at value-->
              <xsl:if test ="@startAt">
                <xsl:attribute name ="text:start-value">
                  <xsl:value-of select ="./a:p/a:pPr/a:buAutoNum[@startAt]"/>
                </xsl:attribute>
              </xsl:if>
            </xsl:if >
            <style:list-level-properties text:min-label-width="0.952cm" />
            <style:text-properties style:font-family-generic="swiss" style:font-pitch="variable">
              <xsl:if test ="./a:p/a:pPr/a:buFont/@typeface">
                <xsl:attribute name ="fo:font-family">
                  <xsl:value-of select ="./a:p/a:pPr/a:buFont/@typeface"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test ="not(./a:p/a:pPr/a:buFont/@typeface)">
                <xsl:attribute name ="fo:font-family">
                  <xsl:value-of select ="'Arial'"/>
                </xsl:attribute>
              </xsl:if>             
              <xsl:if test ="./a:p/a:pPr/a:buSzPct">
                <xsl:attribute name ="fo:font-size">
                  <xsl:value-of select ="concat((./a:p/a:pPr/a:buSzPct/@val div 1000),'%')"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test ="not(./a:p/a:pPr/a:buSzPct)">
                <xsl:attribute name ="fo:font-size">
                  <xsl:value-of select ="'100%'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test ="./a:p/a:pPr/a:buClr">
                <xsl:if test ="./a:p/a:pPr/a:buClr/a:srgbClr">
                  <xsl:attribute name ="fo:color">
                    <xsl:value-of select ="concat('#',./a:p/a:pPr/a:buClr/a:srgbClr/@val)"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test ="a:p/a:pPr/a:buClr/a:schemeClr">
                  <xsl:call-template name ="getColorCode">
                    <xsl:with-param name ="color">
                      <xsl:value-of select="a:pPr/a:buClr/a:schemeClr/@val"/>
                    </xsl:with-param>
                  </xsl:call-template >
                </xsl:if>
              </xsl:if>
            </style:text-properties>
          </text:list-level-style-number>
        </text:list-style>
      </xsl:if>
      <xsl:if test ="./a:p/a:pPr/a:buBlip">
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
    </xsl:if>
    <!--End  of Condition if levels are not present-->
  </xsl:template>

  <xsl:template name ="insertDefaultBulletNumberStyle">
    <xsl:param name ="listStyleName"/>
    <xsl:if test ="./a:p/a:pPr/@lvl">
      <xsl:for-each select ="./a:p">
        <!--Condition if levels are present,and bullets are default-->
        <xsl:if test ="./a:pPr/@lvl">
          <xsl:if test ="not(./a:pPr/a:buNone)">
            <xsl:variable name ="levelStyle">
              <xsl:value-of select ="concat($listStyleName,'lvl',a:pPr/@lvl)"/>
            </xsl:variable>
            <xsl:variable name ="textLevel" select ="a:pPr/@lvl"/>
            <xsl:variable name ="newTextLvl" select ="$textLevel+1"/>
            <xsl:for-each select ="document('ppt/slideMasters/slideMaster1.xml')//p:sldMaster/p:txStyles/p:bodyStyle">
              <text:list-style>
                <xsl:attribute name ="style:name">
                  <xsl:value-of select ="$levelStyle"/>
                </xsl:attribute >
                <xsl:if test ="child::node()[$newTextLvl]/a:buChar">
                  <text:list-level-style-bullet>
                    <xsl:attribute name ="text:level">
                      <xsl:value-of select ="$textLevel+1"/>
                    </xsl:attribute >
                    <xsl:attribute name ="text:bullet-char">
                      <xsl:choose>
                        <xsl:when test="child::node()[$newTextLvl]/a:buChar/@char= 'Ø'">
                          <xsl:value-of select ="'➢'"/>
                        </xsl:when>
                        <xsl:when test="child::node()[$newTextLvl]/a:buChar/@char= 'o'">
                          <xsl:value-of select ="'○'"/>
                        </xsl:when>
                        <xsl:when test="child::node()[$newTextLvl]/a:buChar/@char= '§'">
                          <xsl:value-of select ="'■'"/>
                        </xsl:when>
                        <xsl:when test="child::node()[$newTextLvl]/a:buChar/@char= 'q'">
                          <xsl:value-of select ="''"/>
                        </xsl:when>
                        <xsl:when test="child::node()[$newTextLvl]/a:buChar/@char= 'ü'">
                          <xsl:value-of select ="'✔'"/>
                        </xsl:when>
                        <xsl:when test="child::node()[$newTextLvl]/a:buChar/@char = '–'">
                          <xsl:value-of select ="'–'"/>
                        </xsl:when>
                        <xsl:when test="child::node()[$newTextLvl]/a:buChar/@char = '»'">
                          <xsl:value-of select ="'»'"/>
                        </xsl:when>
                        <xsl:otherwise>•</xsl:otherwise>
                      </xsl:choose >
                    </xsl:attribute >
                    <style:list-level-properties text:min-label-width="0.8cm">
                      <xsl:attribute name ="text:space-before">
                        <xsl:value-of select="concat(format-number(child::node()[$newTextLvl]/@marL div 360000,'#.##'),'cm')"/>
                      </xsl:attribute>
                    </style:list-level-properties>
                    <style:text-properties fo:font-family="StarSymbol" style:font-charset="x-symbol">
                      <xsl:if test ="child::node()[$newTextLvl]/a:buFont/@typeface">
                        <xsl:if test ="child::node()[$newTextLvl]/a:buFont[@typeface='Arial']">
                          <xsl:attribute name ="fo:font-family">
                            <xsl:value-of select ="'StarSymbol'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test ="not(child::node()[$newTextLvl]/a:buFont[@typeface='Arial'])">
                          <xsl:attribute name ="fo:font-family">
                            <xsl:value-of select ="child::node()[$newTextLvl]/a:buFont/@typeface"/>
                          </xsl:attribute>
                        </xsl:if>
                      </xsl:if>
                      <xsl:if test ="not(child::node()[$newTextLvl]/a:buFont/@typeface)">
                        <xsl:attribute name ="fo:font-family">
                          <xsl:value-of select ="'StarSymbol'"/>
                        </xsl:attribute>
                      </xsl:if>                     
                      <xsl:if test ="child::node()[$newTextLvl]/a:buSzPct">
                        <xsl:attribute name ="fo:font-size">
                          <xsl:value-of select ="concat((child::node()[$newTextLvl]/a:buSzPct/@val div 1000),'%')"/>
                        </xsl:attribute>
                      </xsl:if>
                      <xsl:if test ="not(child::node()[$newTextLvl]/a:buSzPct)">
                        <xsl:attribute name ="fo:font-size">
                          <xsl:value-of select ="'45%'"/>
                        </xsl:attribute>
                      </xsl:if>
                      <xsl:if test ="child::node()[$newTextLvl]/a:buClr">
                        <xsl:if test ="child::node()[$newTextLvl]/a:buClr/a:srgbClr">
                          <xsl:variable name ="color" select ="child::node()[$newTextLvl]/a:buClr/a:srgbClr/@val"/>
                          <xsl:attribute name ="fo:color">
                            <xsl:value-of select ="concat('#',$color)"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test ="child::node()[$newTextLvl]/a:buClr/a:schemeClr">
                          <xsl:call-template name ="getColorCode">
                            <xsl:with-param name ="color">
                              <xsl:value-of select="child::node()[$newTextLvl]/a:buClr/a:schemeClr/@val"/>
                            </xsl:with-param>
                          </xsl:call-template >
                        </xsl:if>
                      </xsl:if>
                    </style:text-properties >
                  </text:list-level-style-bullet>
                </xsl:if>
                <xsl:if test ="child::node()[$newTextLvl]/a:buAutoNum">
                  <text:list-level-style-number>
                    <xsl:attribute name ="text:level">
                      <xsl:value-of select ="$textLevel+1"/>
                    </xsl:attribute >
                    <xsl:if test ="child::node()[$newTextLvl]/a:buAutoNum/@type[contains(.,'arabicPeriod')]">
                      <xsl:attribute name ="style:num-format" >
                        <xsl:value-of  select ="'1'"/>
                      </xsl:attribute >
                      <xsl:attribute name ="style:num-suffix">
                        <xsl:value-of select ="'.'"/>
                      </xsl:attribute>
                    </xsl:if>
                    <xsl:if test ="child::node()[$newTextLvl]/a:buAutoNum/@type[contains(.,'arabicParenR')]">
                      <xsl:attribute name ="style:num-format" >
                        <xsl:value-of  select ="'1'"/>
                      </xsl:attribute >
                      <xsl:attribute name ="style:num-suffix">
                        <xsl:value-of select ="')'"/>
                      </xsl:attribute>
                    </xsl:if>
                    <xsl:if test ="child::node()[$newTextLvl]/a:buAutoNum/@type[contains(.,'arabicParenBoth')]">
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
                    <xsl:if test ="child::node()[$newTextLvl]/a:buAutoNum/@type[contains(.,'alphaUcPeriod')]">
                      <xsl:attribute name ="style:num-format" >
                        <xsl:value-of  select ="'A'"/>
                      </xsl:attribute >
                      <xsl:attribute name ="style:num-suffix">
                        <xsl:value-of select ="'.'"/>
                      </xsl:attribute>
                    </xsl:if>
                    <xsl:if test ="child::node()[$newTextLvl]/a:buAutoNum/@type[contains(.,'alphaUcParenR')]">
                      <xsl:attribute name ="style:num-format" >
                        <xsl:value-of  select ="'A'"/>
                      </xsl:attribute >
                      <xsl:attribute name ="style:num-suffix">
                        <xsl:value-of select ="')'"/>
                      </xsl:attribute>
                    </xsl:if>
                    <xsl:if test ="child::node()[$newTextLvl]/a:buAutoNum/@type[contains(.,'alphaUcParenBoth')]">
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
                    <xsl:if test ="child::node()[$newTextLvl]/a:buAutoNum/@type[contains(.,'alphaLcPeriod')]">
                      <xsl:attribute name ="style:num-format" >
                        <xsl:value-of  select ="'a'"/>
                      </xsl:attribute >
                      <xsl:attribute name ="style:num-suffix">
                        <xsl:value-of select ="'.'"/>
                      </xsl:attribute>
                    </xsl:if>
                    <xsl:if test ="child::node()[$newTextLvl]/a:buAutoNum/@type[contains(.,'alphaLcParenR')]">
                      <xsl:attribute name ="style:num-format" >
                        <xsl:value-of  select ="'a'"/>
                      </xsl:attribute >
                      <xsl:attribute name ="style:num-suffix">
                        <xsl:value-of select ="')'"/>
                      </xsl:attribute>
                    </xsl:if>
                    <xsl:if test ="child::node()[$newTextLvl]/a:buAutoNum/@type[contains(.,'alphaLcParenBoth')]">
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
                    <xsl:if test ="child::node()[$newTextLvl]/a:buAutoNum/@type[contains(.,'romanUcPeriod')]">
                      <xsl:attribute name ="style:num-format" >
                        <xsl:value-of  select ="'I'"/>
                      </xsl:attribute >
                      <xsl:attribute name ="style:num-suffix">
                        <xsl:value-of select ="'.'"/>
                      </xsl:attribute>
                    </xsl:if>
                    <xsl:if test ="child::node()[$newTextLvl]/a:buAutoNum/@type[contains(.,'romanLcPeriod')]">
                      <xsl:attribute name ="style:num-format" >
                        <xsl:value-of  select ="'i'"/>
                      </xsl:attribute >
                      <xsl:attribute name ="style:num-suffix">
                        <xsl:value-of select ="'.'"/>
                      </xsl:attribute>
                      <!-- start at value-->
                      <xsl:if test ="@startAt">
                        <xsl:attribute name ="text:start-value">
                          <xsl:value-of select ="child::node()[$newTextLvl]/a:buAutoNum[@startAt]"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:if >
                    <style:list-level-properties text:min-label-width="0.952cm" />
                    <style:text-properties style:font-family-generic="swiss" style:font-pitch="variable" fo:color="#000000" fo:font-size="100%">
                      <xsl:if test ="child::node()[$newTextLvl]/a:buFont/@typeface">
                        <xsl:attribute name ="fo:font-family">
                          <xsl:value-of select ="child::node()[$newTextLvl]/a:buFont/@typeface"/>
                        </xsl:attribute>
                      </xsl:if>
                      <xsl:if test ="not(child::node()[$newTextLvl]/a:buFont/@typeface)">
                        <xsl:attribute name ="fo:font-family">
                          <xsl:value-of select ="'Arial'"/>
                        </xsl:attribute>
                      </xsl:if>                      
                      <xsl:if test ="child::node()[$newTextLvl]/a:buSzPct">
                        <xsl:attribute name ="fo:font-size">
                          <xsl:value-of select ="concat((child::node()[$newTextLvl]/a:buSzPct/@val div 1000),'%')"/>
                        </xsl:attribute>
                      </xsl:if>
                      <xsl:if test ="not(child::node()[$newTextLvl]/a:buSzPct)">
                        <xsl:attribute name ="fo:font-size">
                          <xsl:value-of select ="'45%'"/>
                        </xsl:attribute>
                      </xsl:if>
                      <xsl:if test ="child::node()[$newTextLvl]/a:buClr">
                        <xsl:if test ="child::node()[$newTextLvl]/a:buClr/a:srgbClr">
                          <xsl:variable name ="color" select ="child::node()[$newTextLvl]/a:buClr/a:srgbClr/@val"/>
                          <xsl:attribute name ="fo:color">
                            <xsl:value-of select ="concat('#',$color)"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test ="child::node()[$newTextLvl]/a:buClr/a:schemeClr">
                          <xsl:call-template name ="getColorCode">
                            <xsl:with-param name ="color">
                              <xsl:value-of select="child::node()[$newTextLvl]/a:buClr/a:schemeClr/@val"/>
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
              </text:list-style>
            </xsl:for-each>
          </xsl:if>
        </xsl:if>
        <!--Condition if levels are present,and bullets are default-->
        <xsl:if test ="not(./a:pPr/@lvl)">
          <xsl:if test ="not(./a:pPr/a:buNone)">
            <xsl:for-each select ="document('ppt/slideMasters/slideMaster1.xml')//p:sldMaster/p:txStyles/p:bodyStyle/child::node()[1]">
              <text:list-style>
                <xsl:attribute name ="style:name">
                  <xsl:value-of select ="$listStyleName"/>
                </xsl:attribute >
                <xsl:if test ="a:buChar">
                  <text:list-level-style-bullet>
                    <xsl:attribute name ="text:level">
                      <xsl:value-of select ="1"/>
                    </xsl:attribute >
                    <xsl:attribute name ="text:bullet-char">
                      <xsl:choose>
                        <xsl:when test="a:buChar/@char= 'Ø'">
                          <xsl:value-of select ="'➢'"/>
                        </xsl:when>
                        <xsl:when test="a:buChar/@char= 'o'">
                          <xsl:value-of select ="'○'"/>
                        </xsl:when>
                        <xsl:when test="a:buChar/@char= '§'">
                          <xsl:value-of select ="'■'"/>
                        </xsl:when>
                        <xsl:when test="a:buChar/@char= 'q'">
                          <xsl:value-of select ="''"/>
                        </xsl:when>
                        <xsl:when test="a:buChar/@char= 'ü'">
                          <xsl:value-of select ="'✔'"/>
                        </xsl:when>
                        <xsl:otherwise>•</xsl:otherwise>
                      </xsl:choose >
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
                    </style:text-properties >
                  </text:list-level-style-bullet>
                </xsl:if>
                <xsl:if test ="a:buAutoNum">
                  <text:list-level-style-number>
                    <xsl:attribute name ="text:level">
                      <xsl:value-of select ="'1'"/>
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
                      <!-- start at value-->
                      <xsl:if test ="@startAt">
                        <xsl:attribute name ="text:start-value">
                          <xsl:value-of select ="a:buAutoNum[@startAt]"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:if >
                    <style:list-level-properties text:min-label-width="0.952cm" />
                    <style:text-properties style:font-family-generic="swiss" style:font-pitch="variable" fo:color="#000000" fo:font-size="100%">
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
                          <xsl:value-of select ="'45%'"/>
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
              </text:list-style>
            </xsl:for-each>
          </xsl:if>          
        </xsl:if>
      </xsl:for-each>
    </xsl:if>
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


  <!-- Code By vijayeta,insertMultipleLevels in OdpFiles-->

  <!-- Vijayeta,Bullets-->
  <xsl:template name="InsertBulletChar">
    <xsl:param name ="character"/>
    <xsl:choose>
      <xsl:when test="$character= 'o'">O</xsl:when>
      <xsl:when test="$character = '§'">■</xsl:when>
      <xsl:when test="$character = 'q'"></xsl:when>
      <xsl:when test="$character = '•'">•</xsl:when>
      <xsl:when test="$character = 'v'">●</xsl:when>
      <xsl:when test="$character = 'Ø'">➢</xsl:when>
      <xsl:when test="$character = 'ü'">✔</xsl:when>
      <xsl:when test="$character = ''">■</xsl:when>
      <xsl:when test="$character = 'o'">○</xsl:when>
      <xsl:when test="$character = ''">➔</xsl:when>
      <xsl:when test="$character = ''">✗</xsl:when>
      <xsl:when test="$character = '-' ">–</xsl:when>
      <xsl:when test="$character = '–'">–</xsl:when>
      <xsl:when test="$character = ''">–</xsl:when>
      <xsl:otherwise>•</xsl:otherwise>
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
  <!-- delete this-->
  <xsl:template name="getColorCode">
    <xsl:param name="color"/>
    <xsl:param name ="lumMod"/>
    <xsl:param name ="lumOff"/>
    <xsl:variable name ="ThemeColor">
      <xsl:for-each select ="document('ppt/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme">
        <xsl:for-each select ="node()">
          <xsl:if test ="name() =concat('a:',$color)">
            <xsl:value-of select ="a:srgbClr/@val"/>
          </xsl:if>
        </xsl:for-each>
      </xsl:for-each>
    </xsl:variable>
    <xsl:variable name ="BgTxColors">
      <xsl:if test ="$color ='bg2'">
        <xsl:value-of select ="document('ppt/theme/theme1.xml') //a:lt2/a:srgbClr/@val"/>
      </xsl:if>
      <xsl:if test ="$color ='bg1'">
        <xsl:value-of select ="document('ppt/theme/theme1.xml') //a:lt1/a:srgbClr/@lastClr"/>
      </xsl:if>
      <xsl:if test ="$color ='tx1'">
        <xsl:value-of select ="document('ppt/theme/theme1.xml') //a:dk1/a:srgbClr/@lastClr"/>
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
</xsl:stylesheet>
<!--<xsl:if test ="a:pPr/@lvl">
  <xsl:call-template name ="insertMultipleLevels">
    <xsl:with-param name ="levelCount" select ="a:pPr/@lvl"/>
    <xsl:with-param name ="ParaId" select ="$ParaId"/>
    <xsl:with-param name ="listStyleName" select ="$listStyleName"/>
  </xsl:call-template>
</xsl:if>
<xsl:if test ="not(a:pPr/@lvl)">
  <text:list>
    <xsl:attribute name ="text:style-name">
      <xsl:value-of select ="$listStyleName"/>
    </xsl:attribute>
    <text:list-item>
      <text:p>
        <xsl:attribute name ="text:style-name">
          <xsl:value-of select ="concat($ParaId,position())"/>
        </xsl:attribute >
        <xsl:for-each select ="node()">
          <xsl:if test ="name()='a:r'">
            <text:span text:style-name="{generate-id()}">
              <xsl:value-of select ="a:t"/>
              -->
<!--converts whitespaces sequence to text:s 1699083 bug fix-->
<!--

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
    </text:list-item >
  </text:list >
</xsl:if>-->


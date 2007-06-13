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
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"  
  exclude-result-prefixes="odf style text number draw page">

  <xsl:variable name ="rId" select ="generate-id(.)"/>
  <xsl:template name ="insertBulletsNumbers">
    <xsl:param name ="listId"/>
    <xsl:param name ="level" />
    <!--<xsl:variable name ="newLevel" select ="$level+1"/>-->
    <xsl:for-each select ="document('content.xml')//text:list-style [@style:name=$listId]">
      <xsl:if test ="./child::node()[1]/style:text-properties[@fo:color]">
        <a:buClr>
          <a:srgbClr>
            <xsl:attribute name ="val">
              <xsl:value-of select ="substring-after(./child::node()[1]/style:text-properties/@fo:color,'#')"/>
            </xsl:attribute>
          </a:srgbClr>
        </a:buClr>
      </xsl:if>
      <xsl:if test ="(./child::node()[1]/style:text-properties[@fo:font-size] and not(substring-before(./child::node()[1]/style:text-properties/@fo:font-size,'%') = 100)) ">
        <a:buSzPct>
          <xsl:attribute name ="val">
            <xsl:value-of select ="format-number(substring-before(./child::node()[1]/style:text-properties/@fo:font-size,'%')*1000,'#.##')"/>
          </xsl:attribute>
        </a:buSzPct>
      </xsl:if>
      <a:buFont>
        <xsl:attribute name ="typeface">
          <!--<xsl:call-template name ="getBulletType">
            <xsl:with-param name ="type" select ="@text:bullet-char"/>
          </xsl:call-template>-->
        </xsl:attribute>
      </a:buFont>
      <xsl:if test ="text:list-level-style-bullet">
        <!--<xsl:for-each select ="./child::node()[1]">-->
          <xsl:if test ="text:list-level-style-bullet[@text:level=$level]">
          <a:buChar>
            <xsl:attribute name ="char">
              <xsl:call-template name="insertBulletChar">
                <xsl:with-param name="character" select="./child::node()[$level]/@text:bullet-char"/>
              </xsl:call-template>
            </xsl:attribute>
          </a:buChar>
          </xsl:if>        
        <!--</xsl:for-each >-->
      </xsl:if>
      <xsl:if test="text:list-level-style-number">
        <!--<xsl:for-each select ="./child::node()[1]">-->
        <xsl:if test ="text:list-level-style-number[@text:level =$level]">
          <a:buAutoNum>
            <xsl:attribute name ="type">
                <xsl:call-template name="getNumFormat">
                  <xsl:with-param name="format">
                    <xsl:value-of select="./child::node()[$level]/@style:num-format"/>
                  </xsl:with-param>
                  <xsl:with-param name ="suff" select ="./child::node()[$level]/@style:num-suffix"/>
                  <xsl:with-param name ="prefix" select ="./child::node()[$level]/@style:num-prefix"/>
                </xsl:call-template>
              </xsl:attribute>
              <xsl:if test ="@text:start-value">
                <xsl:attribute name ="startAt">
                  <xsl:value-of select ="@text:start-value"/>
                </xsl:attribute>
              </xsl:if>
            </a:buAutoNum>
          </xsl:if>
        <!--</xsl:for-each>-->
      </xsl:if>
      <xsl:if test ="text:list-level-style-image">
        <!--<xsl:if test ="text:list-level-style-image[@text:level=$level]">-->
          <a:buBlip>
            <a:blip>
                <xsl:attribute name ="r:embed">
                  <xsl:value-of select ="$rId"/>
                </xsl:attribute>
              </a:blip>
            </a:buBlip>
          </xsl:if>        
        <xsl:call-template name="copyPictures"/>
      <!--</xsl:if>-->
    </xsl:for-each>
    <!--End of condition,If Levels Present-->
  </xsl:template>
  <!-- Code by Vijayeta,Types of Bullets and Numbering formats-->
  <xsl:template name="insertBulletChar">
    <xsl:param name ="character"/>
    <xsl:choose>
      <xsl:when test="$character = '' "></xsl:when>
      <xsl:when test="$character = '' "></xsl:when>
      <xsl:when test="$character = '☑' ">☑</xsl:when>
      <xsl:when test="$character = '•' ">•</xsl:when>
      <xsl:when test="$character= '●' ">●</xsl:when>
      <xsl:when test="$character = '➢' ">➢</xsl:when>
      <xsl:when test="$character = '✔' ">✔</xsl:when>
      <xsl:when test="$character = '■' ">■</xsl:when>
      <xsl:when test="$character = '○' ">o</xsl:when>
      <xsl:when test="$character = '➔' ">➔</xsl:when>
      <xsl:when test="$character = '✗' ">✗</xsl:when>
      <xsl:when test="$character = '–' ">–</xsl:when>
      <xsl:otherwise>•</xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="getNumFormat">
    <xsl:param name="format"/>
    <xsl:param name="suff"/>
    <xsl:param name ="prefix"/>
    <xsl:choose>
      <xsl:when test="$format= '1' and $suff=')' and not($prefix) ">arabicParenR</xsl:when>
      <xsl:when test="$format= '1' and $suff=')' and $prefix='(' ">arabicParenBoth</xsl:when>
      <xsl:when test="$format= '1' and $suff='.'">arabicPeriod</xsl:when>
      <xsl:when test="$format= 'A' and $suff='.' ">alphaUcPeriod</xsl:when>
      <xsl:when test="$format= 'A' and $suff=')' and not($prefix) ">alphaUcParenR</xsl:when>
      <xsl:when test="$format= 'A' and $suff=')' and $prefix='(' ">alphaUcParenBoth</xsl:when>
      <xsl:when test="$format= 'a' and $suff='.' ">alphaLcPeriod</xsl:when>
      <xsl:when test="$format= 'a' and $suff=')' and not($prefix) ">alphaLcParenR</xsl:when>
      <xsl:when test="$format= 'a' and $suff=')' and $prefix='(' ">alphaLcParenBoth</xsl:when>
      <xsl:when test="$format= 'i' and $suff='.' ">romanLcPeriod</xsl:when>
      <xsl:when test="$format= 'I' and $suff='.' ">romanUcPeriod</xsl:when>
      <xsl:otherwise>decimal</xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!-- If bullets are pictures-->
  <xsl:template name="copyPictures">
    <xsl:param name ="sourceFolder" select ="'Pictures'"/>
    <xsl:param name="destFolder" select="'ppt/media/'"/>
    <!--  Copy Pictures Files to the picture catalog -->
    <xsl:for-each select="document('content.xml')//office:document-content/office:automatic-styles/text:list-style/text:list-level-style-image">
      <xsl:if test ="@text:level='1'">
        <xsl:variable name="pzipsource">
          <xsl:value-of select="substring-after(@xlink:href,'Pictures/')"/>
          <!-- image1.gif-->
        </xsl:variable>
        <pzip:copy pzip:source="{concat($sourceFolder,'/',$pzipsource)}" pzip:target="{concat($destFolder,'/',$pzipsource)}"/>
      </xsl:if >
    </xsl:for-each>
  </xsl:template>
  <!--End of Code by Vijayeta,Types of Bullets and Numbering formats-->
  <xsl:template name ="getListLevel">
    <xsl:param name ="levelCount"/>
    <xsl:choose>
      <xsl:when test ="./text:list-item/text:list/text:list-item/text:p">
        <xsl:value-of select ="'1'"/>
      </xsl:when>
      <xsl:when test ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:value-of select ="'2'"/>
      </xsl:when>
      <xsl:when test ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:value-of select ="'3'"/>
      </xsl:when>
      <xsl:when test ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:value-of select ="'4'"/>
      </xsl:when>
      <xsl:when test ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:value-of select ="'5'"/>
      </xsl:when>
      <xsl:when test ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:value-of select ="'6'"/>
      </xsl:when>
      <xsl:when test ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:value-of select ="'7'"/>
      </xsl:when>
      <xsl:when test ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:value-of select ="'8'"/>
      </xsl:when>
      <xsl:when test ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:value-of select ="'9'"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>
<?xml version="1.0" encoding="UTF-8"?>
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
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  exclude-result-prefixes="office text style fo">

  <xsl:output method="xml" encoding="UTF-8"/>

  <xsl:key name="list-style" match="text:list-style" use="@style:name"/>
  <xsl:key name="list-content" match="text:list" use="@text:style-name"/>

  <xsl:variable name="automaticListStylesCount"
    select="count(document('content.xml')/office:document-content/office:automatic-styles/text:list-style)"/>
  <xsl:variable name="stylesListStyleCount"
    select="count(document('styles.xml')/office:document-styles/office:styles/text:list-style|document('styles.xml')/office:document-styles/office:automatic-styles/text:list-style)"/>

  <!--generate numbering definitions: abstract numbering w:abstractNumbering and numbering instances w:num -->
  <xsl:template name="numbering">
    <w:numbering>
      <xsl:apply-templates
        select="document('styles.xml')/office:document-styles/office:styles/text:outline-style"
        mode="numbering"/>
      <xsl:apply-templates
        select="document('content.xml')/office:document-content/office:automatic-styles/text:list-style"
        mode="numbering"/>
      <xsl:apply-templates
        select="document('styles.xml')/office:document-styles/office:styles/text:list-style"
        mode="numbering">
        <xsl:with-param name="offset" select="$automaticListStylesCount"/>
      </xsl:apply-templates>
      <xsl:apply-templates
        select="document('styles.xml')/office:document-styles/office:automatic-styles/text:list-style"
        mode="numbering">
        <xsl:with-param name="offset" select="$automaticListStylesCount"/>
      </xsl:apply-templates>
      <xsl:for-each select="document('content.xml')">
        <xsl:apply-templates select="key('restarting-lists','')" mode="numbering">
          <xsl:with-param name="offset" select="$automaticListStylesCount + $stylesListStyleCount"/>
        </xsl:apply-templates>
      </xsl:for-each>

      <xsl:apply-templates
        select="document('styles.xml')/office:document-styles/office:styles/text:outline-style"
        mode="num"/>
      <xsl:apply-templates
        select="document('content.xml')/office:document-content/office:automatic-styles/text:list-style"
        mode="num"/>
      <xsl:apply-templates
        select="document('styles.xml')/office:document-styles/office:styles/text:list-style"
        mode="num">
        <xsl:with-param name="offset" select="$automaticListStylesCount"/>
      </xsl:apply-templates>
      <xsl:apply-templates
        select="document('styles.xml')/office:document-styles/office:automatic-styles/text:list-style"
        mode="num">
        <xsl:with-param name="offset" select="$automaticListStylesCount"/>
      </xsl:apply-templates>
      <xsl:for-each select="document('content.xml')">
        <xsl:apply-templates select="key('restarting-lists','')" mode="num">
          <xsl:with-param name="offset" select="$automaticListStylesCount + $stylesListStyleCount"/>
        </xsl:apply-templates>
      </xsl:for-each>
    </w:numbering>
  </xsl:template>

  <xsl:template match="text:list-style" mode="numbering">
    <xsl:param name="offset" select="0"/>
    <w:abstractNum w:abstractNumId="{count(preceding-sibling::text:list-style)+2+$offset}">
      <xsl:apply-templates
        select="text:list-level-style-number|text:list-level-style-bullet|list-level-style-image"
        mode="numbering"/>
    </w:abstractNum>
  </xsl:template>

  <xsl:template match="text:list-style" mode="num">
    <xsl:param name="offset" select="0"/>
    <w:num w:numId="{count(preceding-sibling::text:list-style)+2+$offset}">
      <w:abstractNumId w:val="{count(preceding-sibling::text:list-style)+2+$offset}"/>
    </w:num>
  </xsl:template>

  <!--bullet or numbered lists properties-->
  <xsl:template match="text:list-level-style-bullet|text:list-level-style-number" mode="numbering">

    <!-- odf supports list level up to 10-->
    <xsl:if test="number(@text:level) &lt; 10">
      <w:lvl w:ilvl="{number(@text:level) - 1}">

        <!--bullet style list-->
        <xsl:if test="name() = 'text:list-level-style-bullet' ">
          <w:numFmt w:val="bullet"/>
          <w:lvlText w:val="{@text:bullet-char}">
            <xsl:attribute name="w:val">
              <xsl:call-template name="InsertBulletChar"/>
            </xsl:attribute>
          </w:lvlText>
        </xsl:if>

        <!--numbered list-->
        <xsl:if test="name() = 'text:list-level-style-number' ">
          <xsl:call-template name="InsertNumberedListProperties"/>
        </xsl:if>

        <!--list justification-->
        <xsl:call-template name="InsertListJustification"/>

        <!--list paragraph properties like tab, indent-->
        <w:pPr>
          <xsl:call-template name="InsertListParagraphProperties"/>
        </w:pPr>

        <!-- bullet type for bullet list-->
        <xsl:choose>
          <xsl:when test="name() = 'text:list-level-style-bullet' ">
            <xsl:call-template name="InsertBulletType">
              <xsl:with-param name="char">
                <xsl:call-template name="InsertBulletChar"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <!-- run properties for numbered list-->
          <xsl:otherwise>
            <xsl:call-template name="InsertRunProperties"/>
          </xsl:otherwise>
        </xsl:choose>

      </w:lvl>
    </xsl:if>

  </xsl:template>

  
  <xsl:template name="InsertListParagraphProperties">
    <xsl:variable name="spaceBeforeTwip">
      <xsl:call-template name="twips-measure">
        <xsl:with-param name="length" select="style:list-level-properties/@text:space-before"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="minLabelWidthTwip">
      <xsl:call-template name="twips-measure">
        <xsl:with-param name="length" select="style:list-level-properties/@text:min-label-width"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="minLabelDistanceTwip">
      <xsl:call-template name="twips-measure">
        <xsl:with-param name="length" select="style:list-level-properties/@text:min-label-distance"
        />
      </xsl:call-template>
    </xsl:variable>

    <!--If props/@style:display-levels is defined and greater than 1, the tabs may not be well converted.-->
    <w:tabs>
      <w:tab>
        <xsl:attribute name="w:val">num</xsl:attribute>
        <xsl:attribute name="w:pos">
          <xsl:choose>
            <xsl:when test="$spaceBeforeTwip &lt; 0">
              <xsl:choose>
                <xsl:when test="$minLabelWidthTwip &lt; $minLabelDistanceTwip">
                  <xsl:value-of
                    select="$minLabelWidthTwip - $spaceBeforeTwip + $minLabelDistanceTwip"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$minLabelWidthTwip - $spaceBeforeTwip"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when test="$minLabelWidthTwip &lt; $minLabelDistanceTwip">
                  <xsl:value-of
                    select="$spaceBeforeTwip + $minLabelWidthTwip + $minLabelDistanceTwip"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$spaceBeforeTwip + $minLabelWidthTwip"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </w:tab>
    </w:tabs>

    <!-- disable hyphenation -->
    <xsl:for-each select="document('content.xml')">
      <xsl:choose>
        <xsl:when
          test="key('automatic-styles',@text:style-name)/style:text-properties/@fo:hyphenate='true'"/>
        <xsl:otherwise>
          <xsl:variable name="styleName">
            <xsl:value-of select="@text:style-name"/>
          </xsl:variable>
          <xsl:for-each select="document('styles.xml')">
            <xsl:choose>
              <xsl:when test="key('styles',$styleName)/style:text-properties/@fo:hyphenate='true'">
                <w:suppressAutoHyphens w:val="false"/>
              </xsl:when>
              <xsl:otherwise>
                <w:suppressAutoHyphens w:val="true"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>

    <w:ind>
      <xsl:attribute name="w:left">
        <xsl:value-of select="$minLabelWidthTwip + $spaceBeforeTwip "/>
      </xsl:attribute>
      <xsl:choose>
        <xsl:when test="$spaceBeforeTwip &lt; 0">
          <xsl:attribute name="w:firstLine">
            <xsl:value-of select="$minLabelWidthTwip"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="w:hanging">
            <xsl:value-of select="$minLabelWidthTwip"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </w:ind>
  </xsl:template>

  <!-- list justification -->
  <xsl:template name="InsertListJustification">
    <w:lvlJc>
      <xsl:attribute name="w:val">
        <xsl:call-template name="InsertTextAlign">
          <xsl:with-param name="align" select="style:list-level-properties/@fo:text-align"/>
        </xsl:call-template>
      </xsl:attribute>
    </w:lvlJc>
  </xsl:template>

  <xsl:template name="InsertBulletChar">
  <xsl:choose>
      <xsl:when test="@text:bullet-char = '' "></xsl:when>
      <xsl:when test="@text:bullet-char = '' "></xsl:when>
      <xsl:when test="@text:bullet-char = '☑' "></xsl:when>
      <xsl:when test="@text:bullet-char = '•' ">•</xsl:when>
      <xsl:when test="@text:bullet-char = '●' "></xsl:when>
      <xsl:when test="@text:bullet-char = '➢' "></xsl:when>
      <xsl:when test="@text:bullet-char = '✔' "></xsl:when>
      <xsl:when test="@text:bullet-char = '■' "></xsl:when>
      <xsl:when test="@text:bullet-char = '○' ">o</xsl:when>
      <xsl:when test="@text:bullet-char = '➔' "></xsl:when>
      <xsl:when test="@text:bullet-char = '✗' "></xsl:when>
      <xsl:when test="@text:bullet-char = '–' ">–</xsl:when>
      <xsl:otherwise>•</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertBulletType">
    <xsl:param name="char"/>
    <xsl:choose>
      <xsl:when
        test="$char = ''  or  $char = '' or  $char = ''  or  $char = ' ✗'  or $char=''  or $char='' ">
        <w:rPr>
          <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
          <xsl:call-template name="GetBulletSize">
            <xsl:with-param name="fontName">
              <xsl:value-of select="current()/style:text-properties/@style:font-name"/>
            </xsl:with-param>
          </xsl:call-template>
        </w:rPr>
      </xsl:when>

      <xsl:when test="$char = '' ">
        <w:rPr>
          <w:rFonts w:ascii="Wingdings 3" w:hAnsi="Wingdings 3" w:hint="default"/>
          <xsl:call-template name="GetBulletSize">
            <xsl:with-param name="fontName">
              <xsl:value-of select="current()/style:text-properties/@style:font-name"/>
            </xsl:with-param>
          </xsl:call-template>
        </w:rPr>
      </xsl:when>

      <xsl:when test="$char = '' ">
        <w:rPr>
          <w:rFonts w:ascii="Wingdings 2" w:hAnsi="Wingdings 2" w:hint="default"/>
          <xsl:call-template name="GetBulletSize">
            <xsl:with-param name="fontName">
              <xsl:value-of select="current()/style:text-properties/@style:font-name"/>
            </xsl:with-param>
          </xsl:call-template>
       </w:rPr>
      </xsl:when>

      <xsl:when test="$char  = 'o'  or $char = '•' or $char = '–' ">
        <w:rPr>
          <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:hint="default"/>
          <xsl:call-template name="GetBulletSize">
            <xsl:with-param name="fontName">
              <xsl:value-of select="current()/style:text-properties/@style:font-name"/>
            </xsl:with-param>
          </xsl:call-template>
       </w:rPr>
      </xsl:when>
      <xsl:when test="$char  = '' ">
        <w:rPr>
          <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>
          <xsl:call-template name="GetBulletSize">
            <xsl:with-param name="fontName">
              <xsl:value-of select="current()/style:text-properties/@style:font-name"/>
            </xsl:with-param>
          </xsl:call-template>
       </w:rPr>
      </xsl:when>
     <xsl:otherwise/>
    </xsl:choose>
  </xsl:template>

  <!-- numbered list properties like: numbering format, start value -->
  <xsl:template name="InsertNumberedListProperties">
    <xsl:choose>
      <xsl:when test="@text:start-value">
        <w:start w:val="{@text:start-value}"/>
      </xsl:when>
      <xsl:otherwise>
        <w:start w:val="1"/>
      </xsl:otherwise>
    </xsl:choose>
    <w:numFmt>
      <xsl:attribute name="w:val">
        <xsl:call-template name="GetNumFormat">
          <xsl:with-param name="format">
            <xsl:value-of select="@style:num-format"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </w:numFmt>
    <xsl:choose>
      <xsl:when test="ancestor::text:outline-style">
        <w:lvlText>
            <xsl:attribute name="w:val">
            <xsl:value-of select="@style:num-prefix"/>
            <xsl:choose>
              <xsl:when test="@style:num-format='1'">
                <xsl:call-template name="GetNumberingLevelText">
                  <xsl:with-param name="level" select="@text:level"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="@style:num-format='a'">
                <xsl:value-of select="concat('%',string(@text:level))"/>
              </xsl:when>
              <xsl:when test="@style:num-format='A'">
                <xsl:value-of select="concat('%',string(@text:level))"/>
              </xsl:when>
              <xsl:otherwise/>
            </xsl:choose>
            <xsl:value-of select="@style:num-suffix"/>
          </xsl:attribute>
        </w:lvlText>
      </xsl:when>
      <xsl:otherwise>
        <w:lvlText>
          <xsl:attribute name="w:val">
            <xsl:choose>
              <xsl:when test="@style:num-format = '' "/>
              <xsl:otherwise>
                <xsl:value-of select="@style:num-prefix"/>
                <xsl:choose>
                  <xsl:when test="@text:display-levels">
                    <xsl:call-template name="GetLevelText">
                      <xsl:with-param name="displayLevels" select="@text:display-levels"/>
                      <xsl:with-param name="level" select="@text:level"/>
                      <xsl:with-param name="listStyleName"
                        select="ancestor::text:list-style/@style:name"/>
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:otherwise>%<xsl:value-of select="@text:level"/></xsl:otherwise>
                </xsl:choose>
                <xsl:value-of select="@style:num-suffix"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </w:lvlText>
      </xsl:otherwise>
    </xsl:choose>
   </xsl:template>

  <xsl:template name="GetBulletSize">
    <xsl:param name="fontName"/>
    <xsl:if
      test="document('styles.xml')/office:document-styles/office:styles/style:style/style:text-properties[@style:font-name = $fontName]">
      <xsl:variable name="sz">
        <xsl:call-template name="computeSize">
          <xsl:with-param name="node"
            select="document('styles.xml')/office:document-styles/office:styles/style:style/style:text-properties[@style:font-name = $fontName]"
          />
        </xsl:call-template>
      </xsl:variable>
      <xsl:if test="number($sz)">
        <w:sz w:val="{$sz}"/>
      </xsl:if>
    </xsl:if>
  </xsl:template>
  <!--
		function    : num-format
		param       : format (string)
		description : convert the ODF numFormat to OOX numFormat
	-->
  <xsl:template name="GetNumFormat">
    <xsl:param name="format"/>
    <xsl:choose>
      <xsl:when test="$format= '1' ">decimal</xsl:when>
      <xsl:when test="$format= 'i' ">lowerRoman</xsl:when>
      <xsl:when test="$format= 'I' ">upperRoman</xsl:when>
      <xsl:when test="$format= 'a' ">lowerLetter</xsl:when>
      <xsl:when test="$format= 'A' ">upperLetter</xsl:when>
      <xsl:otherwise>decimal</xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <xsl:template name="InsertTextAlign">
    <xsl:param name="align"/>
    <xsl:choose>
      <xsl:when test="left">left</xsl:when>
      <xsl:when test="right">right</xsl:when>
      <xsl:when test="center">center</xsl:when>
      <xsl:otherwise>left</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- outline style is applied to headings : generate numbering instance -->
  <xsl:template match="text:outline-style" mode="num">
    <w:num w:numId="1">
      <w:abstractNumId w:val="1"/>
    </w:num>
  </xsl:template>

  <!-- outline style applied to headings : generate abstract num definition-->
  <xsl:template match="text:outline-style" mode="numbering">
    <w:abstractNum w:abstractNumId="1">
      <xsl:for-each select="text:outline-level-style[@text:level &lt; 10]">
        <xsl:if test="@text:level">
          <w:lvl w:ilvl="{number(@text:level)-1}">

            <!--numbered list properties: num start, level, num type-->
            <xsl:call-template name="InsertNumberedListProperties"/>

            <!--justification-->
            <w:lvlJc w:val="left"/>

            <!--  list paragraph properties: indent -->
            <w:pPr>
              <xsl:call-template name="InsertOutlineParagraphProperties"/>
            </w:pPr>
          </w:lvl>
        </xsl:if>
      </xsl:for-each>
    </w:abstractNum>
  </xsl:template>

  <xsl:template name="InsertOutlineParagraphProperties">
    <xsl:variable name="spaceBeforeTwip">
      <xsl:call-template name="twips-measure">
        <xsl:with-param name="length" select="style:list-level-properties/@text:space-before"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="minLabelWidthTwip">
      <xsl:call-template name="twips-measure">
        <xsl:with-param name="length" select="style:list-level-properties/@text:min-label-width"/>
      </xsl:call-template>
    </xsl:variable>

    <w:ind>
      <xsl:choose>
        <xsl:when test="number($spaceBeforeTwip) &lt; 0">
          <xsl:choose>
            <xsl:when test="style:list-level-properties/@text:min-label-width">
              <xsl:attribute name="w:left">
                <xsl:value-of select="$minLabelWidthTwip + $spaceBeforeTwip"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="w:left">
                <xsl:value-of select="$spaceBeforeTwip"/>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
          <xsl:choose>
            <xsl:when test="style:list-level-properties/@text:space-before">
              <xsl:attribute name="w:firstLine">
                <xsl:value-of select="$minLabelWidthTwip"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="w:hanging">
                <xsl:value-of select="$minLabelWidthTwip"/>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:choose>
            <xsl:when test="style:list-level-properties/@text:min-label-width">
              <xsl:attribute name="w:left">
                <xsl:value-of select="$minLabelWidthTwip + $spaceBeforeTwip"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="w:left">
                <xsl:value-of select="$spaceBeforeTwip"/>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
          <xsl:choose>
            <xsl:when test="style:list-level-properties/@text:space-before">
              <xsl:attribute name="w:hanging">
                <xsl:value-of select="$minLabelWidthTwip"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="w:hanging">
                <xsl:value-of select="$minLabelWidthTwip"/>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </w:ind>
  </xsl:template>

  <xsl:template name="GetNumberingLevelText">
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test="$level = 9">%1.%2.%3.%4.%5.%6.%7.%8.%9</xsl:when>
      <xsl:when test="$level = 8">%1.%2.%3.%4.%5.%6.%7.%8</xsl:when>
      <xsl:when test="$level = 7">%1.%2.%3.%4.%5.%6.%7</xsl:when>
      <xsl:when test="$level = 6">%1.%2.%3.%4.%5.%6</xsl:when>
      <xsl:when test="$level = 5">%1.%2.%3.%4.%5</xsl:when>
      <xsl:when test="$level = 4">%1.%2.%3.%4</xsl:when>
      <xsl:when test="$level = 3">%1.%2.%3</xsl:when>
      <xsl:when test="$level = 2">%1.%2</xsl:when>
      <xsl:when test="$level = 1">%1</xsl:when>
      <xsl:otherwise>%1</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Generate the format string for multiple levels -->
  <!--  level : current level -->
  <!-- displayLevels : number of preceding levels displayed at this level -->
  <xsl:template name="GetLevelText">
    <xsl:param name="displayLevels"/>
    <xsl:param name="level"/>
    <xsl:param name="listStyleName"/>
    
    <xsl:if test="$displayLevels &gt; 0">
      <xsl:call-template name="GetLevelText">
        <xsl:with-param name="displayLevels" select="$displayLevels - 1"/>
        <xsl:with-param name="level" select="$level - 1"/>
        <xsl:with-param name="listStyleName" select="$listStyleName"/>
      </xsl:call-template>
      <!-- levels with no formatting are not displayed -->
      <xsl:if
        test="key('list-style', $listStyleName)/text:list-level-style-number[@text:level = $level]/@style:num-format != '' ">
        <xsl:if test="$displayLevels &gt; 1">.</xsl:if>%<xsl:value-of select="$level"/>
      </xsl:if>
    </xsl:if>
  </xsl:template>
  
  
  
  <!-- appearance and behaviour for set of list paragraphs -->
  <xsl:template match="text:list" mode="numbering">
    <xsl:param name="offset" select="0"/>
    <xsl:variable name="listStyleName" select="@text:style-name"/>
    <w:abstractNum
      w:abstractNumId="{count(preceding-sibling::text:list[text:list-item/@text:start-value])+2+$offset}">
        <xsl:choose>
          
          <!--  look for list style in content.xml -->
          <xsl:when test="key('list-style', $listStyleName)">
            <xsl:for-each select="key('list-style', $listStyleName)">
              <xsl:apply-templates
                select="text:list-level-style-number|text:list-level-style-bullet|list-level-style-image"
                mode="numbering"/>
            </xsl:for-each>
          </xsl:when>
          
          <xsl:otherwise>
            <!-- look for list style in styles.xml -->
            <xsl:for-each select="document('styles.xml')">
              <xsl:choose>
                
                <xsl:when test="key('list-style', $listStyleName)">
                  <xsl:for-each select="key('list-style', $listStyleName)">
                    <xsl:apply-templates
                      select="text:list-level-style-number|text:list-level-style-bullet|list-level-style-image"
                      mode="numbering"/>
                  </xsl:for-each>
                </xsl:when>
                
                <xsl:otherwise> </xsl:otherwise>
              
              </xsl:choose>
            </xsl:for-each>
          </xsl:otherwise>
        </xsl:choose>
    </w:abstractNum>
  </xsl:template>

 
  <!-- numbering definition which references particular abstractNum and can be applied to a list paragraph -->
  <xsl:template match="text:list" mode="num">
    <xsl:param name="offset" select="0"/>
    <w:num
      w:numId="{count(preceding-sibling::text:list[text:list-item/@text:start-value])+2+$offset}">
      <w:abstractNumId
        w:val="{count(preceding-sibling::text:list[text:list-item/@text:start-value])+2+$offset}"/>
    </w:num>
  </xsl:template>

  <!-- numbering id put in paragraph which reference to numbering definition -->
  <xsl:template name="GetNumberingId">
    <xsl:param name="styleName"/>
    <xsl:choose>

      <!-- first, if list is a restarting special overriding num-->
      <xsl:when test="parent::text:list/text:list-item/@text:start-value">
        <xsl:value-of
          select="count(parent::text:list/preceding-sibling::text:list[text:list-item/@text:start-value])+2+$stylesListStyleCount + $automaticListStylesCount"
        />
      </xsl:when>

      <!-- look for this list style into content.xml -->
      <xsl:when
        test="key('list-style', $styleName) and not(ancestor::style:header) and not(ancestor::style:footer)">
        <xsl:value-of
          select="count(key('list-style', $styleName)/preceding-sibling::text:list-style)+2"/>
      </xsl:when>

      <!-- otherwise, look into styles.xml (add the offset) -->
      <xsl:otherwise>
        <xsl:for-each select="document('styles.xml')">
          <xsl:value-of
            select="count(key('list-style', $styleName)/preceding-sibling::text:list-style)+2+$automaticListStylesCount"
          />
        </xsl:for-each>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- computes numbering indent which is applied in list paragraph properties -->
  <xsl:template name="ComputeNumberingIndent">
    <xsl:param name="attribute"/>
    <xsl:param name="level"/>
    <xsl:param name="listStyleName"/>

     <xsl:call-template name="twips-measure">
       <xsl:with-param name="length">
         <xsl:choose>
           <xsl:when test="$listStyleName='' and $attribute='text:min-label-distance'">
             <!--   text:outline-style -->
             <xsl:value-of
               select="document('styles.xml')//text:outline-style/*[@text:level = $level+1]/style:list-level-properties/@text:min-label-distance"
             />
           </xsl:when>
           <!-- look into content.xml -->
           <xsl:when
             test="key('list-style', $listStyleName)[1]/*[@text:level = $level+1]/style:list-level-properties/@*[name()=$attribute]">
             <xsl:value-of
               select="key('list-style', $listStyleName)[1]/*[@text:level = $level+1]/style:list-level-properties/@*[name()=$attribute]"/>
           </xsl:when>
           <xsl:otherwise>
             <!-- look into styles.xml -->
             <xsl:for-each select="document('styles.xml')">
               <xsl:choose>
               <xsl:when
                 test="key('list-style', $listStyleName)[1]/*[@text:level = $level+1]/style:list-level-properties/@*[name()=$attribute]">
                 <xsl:value-of
                   select="key('list-style', $listStyleName)[1]/*[@text:level = $level+1]/style:list-level-properties/@*[name()=$attribute]"/>
               </xsl:when>
               <xsl:otherwise>0</xsl:otherwise>
               </xsl:choose>
             </xsl:for-each>
           </xsl:otherwise>
         </xsl:choose>
       
       </xsl:with-param>
     </xsl:call-template>
  </xsl:template>

  
  <!-- Inserts paragraph indentation -->
  <xsl:template name="InsertIndent">
    <xsl:param name="level" select="0"/>
    
    <xsl:if test="ancestor-or-self::text:list">
      <xsl:variable name="styleName">
        <xsl:call-template name="GetStyleName"/>
      </xsl:variable>
      <xsl:variable name="parentStyleName"
        select="key('automatic-styles',$styleName)/@style:parent-style-name"/>
      <xsl:variable name="listStyleName"
        select="key('automatic-styles',$styleName)/@style:list-style-name"/>
      
      <!-- Indent to add to numbering values. -->
      <xsl:variable name="addLeftIndent">
        <xsl:call-template name="ComputeAdditionalIndent">
          <xsl:with-param name="side" select="'left'"/>
          <xsl:with-param name="style" select="key('automatic-styles', $styleName)[1]"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="addRightIndent">
        <xsl:call-template name="ComputeAdditionalIndent">
          <xsl:with-param name="side" select="'right'"/>
          <xsl:with-param name="style" select="key('automatic-styles', $styleName)[1]"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="firstLineIndent">
        <xsl:call-template name="GetFirstLineIndent">
          <xsl:with-param name="style" select="key('automatic-styles', $styleName)[1]"/>
        </xsl:call-template>
      </xsl:variable>
      
      <!-- Indent extracted from numbering properties -->
      <xsl:variable name="minLabelWidthTwip">
        <xsl:call-template name="ComputeNumberingIndent">
          <xsl:with-param name="attribute" select="'text:min-label-width'"/>
          <xsl:with-param name="level" select="$level"/>
          <xsl:with-param name="listStyleName" select="$listStyleName"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="spaceBeforeTwip">
        <xsl:call-template name="ComputeNumberingIndent">
          <xsl:with-param name="attribute" select="'text:space-before'"/>
          <xsl:with-param name="level" select="$level"/>
          <xsl:with-param name="listStyleName" select="$listStyleName"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="minLabelDistanceTwip">
        <xsl:call-template name="ComputeNumberingIndent">
          <xsl:with-param name="attribute" select="'text:min-label-distance'"/>
          <xsl:with-param name="level" select="$level"/>
          <xsl:with-param name="listStyleName" select="$listStyleName"/>
        </xsl:call-template>
      </xsl:variable>
      
      <!-- Override tabs if margin defined. -->
      <xsl:if
        test="not(ancestor-or-self::text:list-header) and (self::text:list-item or not(preceding-sibling::node())) and $addLeftIndent != 0">
        <w:tabs>
          <w:tab>
            <xsl:attribute name="w:val">clear</xsl:attribute>
            <xsl:attribute name="w:pos">
              <xsl:choose>
                <xsl:when
                  test="document('content.xml')//text:list-style[@style:name = $listStyleName]/*[@text:level = $level+1]/style:list-level-style-number/@text:display-levels">
                  <xsl:value-of select="$spaceBeforeTwip + $minLabelDistanceTwip"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:choose>
                    <xsl:when test="$minLabelWidthTwip &lt; $minLabelDistanceTwip">
                      <xsl:value-of
                        select="$spaceBeforeTwip + $minLabelWidthTwip + $minLabelDistanceTwip"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$spaceBeforeTwip + $minLabelWidthTwip"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </w:tab>
          <xsl:choose>
            <xsl:when test="$firstLineIndent != 0">
              <w:tab>
                <xsl:attribute name="w:val">num</xsl:attribute>
                <xsl:attribute name="w:pos">
                  <xsl:value-of select="$minLabelDistanceTwip + $addLeftIndent + $firstLineIndent"/>
                </xsl:attribute>
              </w:tab>
            </xsl:when>
            <xsl:otherwise>
              <w:tab>
                <xsl:attribute name="w:val">num</xsl:attribute>
                <xsl:attribute name="w:pos">
                  <xsl:choose>
                    <xsl:when
                      test="document('content.xml')//text:list-style[@style:name = $listStyleName]/*[@text:level = $level+1]/style:list-level-style-number/@text:display-levels">
                      <xsl:value-of
                        select="$addLeftIndent + $spaceBeforeTwip + $minLabelDistanceTwip"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:choose>
                        <xsl:when test="$minLabelWidthTwip &lt; $minLabelDistanceTwip">
                          <xsl:value-of
                            select="$addLeftIndent + $spaceBeforeTwip + $minLabelWidthTwip + $minLabelDistanceTwip"
                          />
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of
                            select="$addLeftIndent + $spaceBeforeTwip + $minLabelWidthTwip"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
              </w:tab>
            </xsl:otherwise>
          </xsl:choose>
        </w:tabs>
      </xsl:if>
      
      <!-- insert indent with paragraph and numbering properties
        if any value affects the numbering style defined, or if list element has no numberig -->
      <xsl:choose>
        <!-- List element with numbering -->
        <xsl:when
          test="not(ancestor-or-self::text:list-header) and (self::text:list-item or not(preceding-sibling::node()))">
          <xsl:if test="$addLeftIndent != 0 or $addRightIndent !=  0 or $firstLineIndent != 0">
            <w:ind>
              <xsl:attribute name="w:left">
                <xsl:value-of select="$addLeftIndent + $spaceBeforeTwip + $minLabelWidthTwip"/>
              </xsl:attribute>
              <xsl:attribute name="w:right">
                <xsl:value-of select="$addRightIndent"/>
              </xsl:attribute>
              <xsl:choose>
                <xsl:when test="$firstLineIndent !=0">
                  <xsl:attribute name="w:firstLine">
                    <xsl:value-of select="$firstLineIndent - $minLabelWidthTwip"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:choose>
                    <xsl:when test="$spaceBeforeTwip &lt; 0">
                      <xsl:attribute name="w:firstLine">
                        <xsl:value-of select="$minLabelWidthTwip"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name="w:hanging">
                        <xsl:value-of select="$minLabelWidthTwip"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:otherwise>
              </xsl:choose>
            </w:ind>
          </xsl:if>
        </xsl:when>
        <!-- Othe list element -->
        <xsl:otherwise>
          <w:ind>
            <xsl:attribute name="w:left">
              <xsl:value-of select="$addLeftIndent + $spaceBeforeTwip + $minLabelWidthTwip"/>
            </xsl:attribute>
            <xsl:attribute name="w:right">
              <xsl:value-of select="$addRightIndent"/>
            </xsl:attribute>
          </w:ind>
        </xsl:otherwise>
      </xsl:choose>
      
    </xsl:if>
  </xsl:template>
  
  
  
</xsl:stylesheet>

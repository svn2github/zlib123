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
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/3/main"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  exclude-result-prefixes="office text style fo">

  <xsl:output method="xml" encoding="UTF-8"/>

  <xsl:key name="list-style" match="text:list-style" use="@style:name"/>
  <xsl:key name="list-content" match="text:list" use="@text:style-name"/>

  <xsl:variable name="automaticListStylesCount"
    select="count(document('content.xml')/office:document-content/office:automatic-styles/text:list-style)"/>
  <xsl:variable name="stylesListStyleCount"
    select="count(document('styles.xml')/office:document-styles/office:styles/text:list-style)"/>
  

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

  <xsl:template match="text:list-level-style-bullet|text:list-level-style-number" mode="numbering">
    <xsl:if test="number(@text:level) &lt; 10">
      <w:lvl w:ilvl="{number(@text:level) - 1}">
        <xsl:if test="name() = 'text:list-level-style-bullet' ">
          <w:numFmt w:val="bullet"/>
          <w:lvlText w:val="{@text:bullet-char}">
            <xsl:attribute name="w:val">
              <xsl:call-template name="textChar"/>
            </xsl:attribute>
          </w:lvlText>
        </xsl:if>
        <xsl:if test="name() = 'text:list-level-style-number' ">
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
              <xsl:call-template name="num-format">
                <xsl:with-param name="format">
                  <xsl:value-of select="@style:num-format"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
          </w:numFmt>
          <!-- xsl:if test="ancestor::text:list-style/@text:consecutive-numbering = 'true' ">
						<w:lvlRestart w:val="1"/>
						</xsl:if -->
          <w:lvlText>
            <xsl:attribute name="w:val">
              <xsl:choose>
                <xsl:when test="@style:num-format = '' "/>
                <xsl:otherwise>
                  <xsl:value-of select="@style:num-prefix"/>
                  <xsl:choose>
                    <xsl:when test="@text:display-levels">
                      <xsl:call-template name="display-levels">
                        <xsl:with-param name="dlvl" select="@text:display-levels"/>
                        <xsl:with-param name="lvl" select="@text:level"/>
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
        </xsl:if>

        <w:lvlJc>
          <xsl:attribute name="w:val">
            <xsl:call-template name="text-align">
              <xsl:with-param name="align" select="style:list-level-properties/@fo:text-align"/>
            </xsl:call-template>
          </xsl:attribute>
        </w:lvlJc>
        <w:pPr>
          <xsl:variable name="spaceBeforeTwip">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length" select="style:list-level-properties/@text:space-before"
              />
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="minLabelWidthTwip">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length"
                select="style:list-level-properties/@text:min-label-width"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="minLabelDistanceTwip">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length"
                select="style:list-level-properties/@text:min-label-distance"/>
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
                        <xsl:value-of select="$minLabelWidthTwip + $minLabelDistanceTwip"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="$minLabelWidthTwip"/>
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
                    <xsl:when
                      test="key('styles',$styleName)/style:text-properties/@fo:hyphenate='true'">
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
        </w:pPr>
        <xsl:choose>
          <xsl:when test="name() = 'text:list-level-style-bullet' ">
            <xsl:call-template name="bulletType">
              <xsl:with-param name="char">
                <xsl:call-template name="textChar"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="InsertRunProperties"/>
          </xsl:otherwise>
        </xsl:choose>
      </w:lvl>
    </xsl:if>

  </xsl:template>

  <xsl:template name="textChar">
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

  <xsl:template name="bulletType">
    <xsl:param name="char"/>
    <xsl:choose>
      <xsl:when
        test="$char = ''  or  $char = '' or  $char = ''  or  $char = ' ✗'  or $char=''  or $char='' ">
        <w:rPr>
          <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
          <xsl:call-template name="bulletSize">
            <xsl:with-param name="fontName">
              <xsl:value-of select="current()/style:text-properties/@style:font-name"/>
            </xsl:with-param>
          </xsl:call-template>


        </w:rPr>
      </xsl:when>

      <xsl:when test="$char = '' ">
        <w:rPr>
          <w:rFonts w:ascii="Wingdings 3" w:hAnsi="Wingdings 3" w:hint="default"/>
          <xsl:call-template name="bulletSize">
            <xsl:with-param name="fontName">
              <xsl:value-of select="current()/style:text-properties/@style:font-name"/>
            </xsl:with-param>
          </xsl:call-template>


        </w:rPr>
      </xsl:when>

      <xsl:when test="$char = '' ">
        <w:rPr>
          <w:rFonts w:ascii="Wingdings 2" w:hAnsi="Wingdings 2" w:hint="default"/>
          <xsl:call-template name="bulletSize">
            <xsl:with-param name="fontName">
              <xsl:value-of select="current()/style:text-properties/@style:font-name"/>
            </xsl:with-param>
          </xsl:call-template>

        </w:rPr>
      </xsl:when>

      <xsl:when test="$char  = 'o'  or $char = '•' or $char = '–' ">
        <w:rPr>
          <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:hint="default"/>
          <xsl:call-template name="bulletSize">
            <xsl:with-param name="fontName">
              <xsl:value-of select="current()/style:text-properties/@style:font-name"/>
            </xsl:with-param>
          </xsl:call-template>


        </w:rPr>
      </xsl:when>
      <xsl:when test="$char  = '' ">
        <w:rPr>
          <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>
          <xsl:call-template name="bulletSize">
            <xsl:with-param name="fontName">
              <xsl:value-of select="current()/style:text-properties/@style:font-name"/>
            </xsl:with-param>
          </xsl:call-template>

        </w:rPr>
      </xsl:when>

      <xsl:otherwise/>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="bulletSize">
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
  <xsl:template name="num-format">
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


  <xsl:template name="text-align">
    <xsl:param name="align"/>
    <xsl:choose>
      <xsl:when test="left">left</xsl:when>
      <xsl:when test="right">right</xsl:when>
      <xsl:when test="center">center</xsl:when>
      <xsl:otherwise>left</xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <!-- Generate the format string for multiple levels -->
  <xsl:template name="display-levels">
    <xsl:param name="dlvl"/>
    <xsl:param name="lvl"/>
    <xsl:param name="listStyleName"/>
    <xsl:if test="$dlvl &gt; 0">
      <xsl:call-template name="display-levels">
        <xsl:with-param name="dlvl" select="$dlvl - 1"/>
        <xsl:with-param name="lvl" select="$lvl - 1"/>
        <xsl:with-param name="listStyleName" select="$listStyleName"/>
      </xsl:call-template>
      <!-- levels with no formatting are not displayed -->
      <xsl:if
        test="key('list-style', $listStyleName)/text:list-level-style-number[@text:level = $lvl]/@style:num-format != '' ">
        <xsl:if test="$dlvl &gt; 1">.</xsl:if>%<xsl:value-of select="$lvl"/>
      </xsl:if>
    </xsl:if>
  </xsl:template>
  <xsl:template match="text:outline-style" mode="num">
    <w:num w:numId="1">
      <w:abstractNumId w:val="1"/>
    </w:num>
  </xsl:template>

  <xsl:template match="text:outline-style" mode="numbering">
    <w:abstractNum w:abstractNumId="1">
      <xsl:for-each select="text:outline-level-style[@text:level &lt; 10]">
        <xsl:if test="@text:level">
          <w:lvl w:ilvl="{number(@text:level)-1}">

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
                <xsl:call-template name="num-format">
                  <xsl:with-param name="format" select="@style:num-format"/>
                </xsl:call-template>
              </xsl:attribute>
            </w:numFmt>
            <!-- <w:pStyle w:val="{concat('Heading',number(@text:level)-1)}"/> -->
            <w:lvlText>

              <xsl:attribute name="w:val">
                <xsl:value-of select="@style:num-prefix"/>
                <xsl:choose>
                  <xsl:when test="@style:num-format='1'">
                    <xsl:call-template name="text-format">
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
            <w:lvlJc w:val="left"/>
            <w:pPr>
              <xsl:variable name="spaceBeforeTwip">
                <xsl:call-template name="twips-measure">
                  <xsl:with-param name="length"
                    select="style:list-level-properties/@text:space-before"/>
                </xsl:call-template>
              </xsl:variable>
              <xsl:variable name="minLabelWidthTwip">
                <xsl:call-template name="twips-measure">
                  <xsl:with-param name="length"
                    select="style:list-level-properties/@text:min-label-width"/>
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
            </w:pPr>
          </w:lvl>
        </xsl:if>
      </xsl:for-each>


    </w:abstractNum>
  </xsl:template>
  <!--	<xsl:template match="text:outline-level-style">
		
	</xsl:template> -->

  <xsl:template name="text-format">
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
  
  <xsl:template match="text:list" mode="numbering">
    <xsl:param name="offset" select="0"/>
    <xsl:variable name="listStyleName" select="@text:style-name"/>
    <w:abstractNum w:abstractNumId="{count(preceding-sibling::text:list[text:list-item/@text:start-value])+2+$offset}">
      <xsl:choose>
        <xsl:when
          test="document('content.xml')/office:document-content/office:automatic-styles/text:list-style/@style:name=$listStyleName">
          <xsl:for-each
            select="document('content.xml')/office:document-content/office:automatic-styles/text:list-style[@style:name=$listStyleName]">
            <xsl:apply-templates
              select="text:list-level-style-number|text:list-level-style-bullet|list-level-style-image"
              mode="numbering"/>
          </xsl:for-each>
        </xsl:when>
        <xsl:when
          test="document('styles.xml')/office:document-styles/office:styles/text:list-style/@style:name=$listStyleName">
          <xsl:for-each
            select="document('styles.xml')/office:document-styles/office:styles/text:list-style[@style:name=$listStyleName]">
            <xsl:apply-templates
              select="text:list-level-style-number|text:list-level-style-bullet|list-level-style-image"
              mode="numbering"/>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
        </xsl:otherwise>
      </xsl:choose>
    </w:abstractNum>
  </xsl:template>
  
  <xsl:template match="text:list" mode="num">
    <xsl:param name="offset" select="0"/>
    <w:num w:numId="{count(preceding-sibling::text:list[text:list-item/@text:start-value])+2+$offset}">
      <w:abstractNumId w:val="{count(preceding-sibling::text:list[text:list-item/@text:start-value])+2+$offset}"/>
    </w:num>
  </xsl:template>
    
  <xsl:template name="numberingId">
    <xsl:param name="styleName"/>
    <xsl:choose>
      <!-- first, if list is a restarting special overriding num-->
      <xsl:when test="parent::text:list/text:list-item/@text:start-value">
        <xsl:value-of
          select="count(parent::text:list/preceding-sibling::text:list[text:list-item/@text:start-value])+2+$stylesListStyleCount + $automaticListStylesCount"
        />
      </xsl:when>
      <!-- look for this list style into content.xml -->
      <xsl:when test="key('list-style', $styleName)">
        <xsl:value-of
          select="count(key('list-style',$styleName)/preceding-sibling::text:list-style)+2+$stylesListStyleCount"
        />
      </xsl:when>
      <!-- otherwise, look into styles.xml (add the offset) -->
      <xsl:otherwise>
        <xsl:for-each select="document('styles.xml')">
          <xsl:value-of
            select="count(key('list-style',$styleName)/preceding-sibling::text:list-style)+2+$automaticListStylesCount"
          />
        </xsl:for-each>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
</xsl:stylesheet>

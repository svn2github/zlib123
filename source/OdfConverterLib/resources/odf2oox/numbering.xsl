﻿<?xml version="1.0" encoding="UTF-8"?>
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
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  exclude-result-prefixes="office text style fo xlink">

  <xsl:output method="xml" encoding="UTF-8"/>

  <xsl:key name="list-style" match="text:list-style" use="@style:name"/>
  <xsl:key name="list-content" match="text:list" use="@text:style-name"/>
  <xsl:key name="bullets" match="text:list-level-style-image" use="''"/>
  <xsl:key name="outlined-styles"
    match="style:style[@style:default-outline-level and @style:list-style-name]" use="''"/>

  <xsl:variable name="automaticListStylesCount"
    select="count(document('content.xml')/office:document-content/office:automatic-styles/text:list-style)"/>
  <xsl:variable name="stylesListStyleCount"
    select="count(document('styles.xml')/office:document-styles/office:styles/text:list-style|document('styles.xml')/office:document-styles/office:automatic-styles/text:list-style)"/>
  <xsl:variable name="chapterListCount">
    <!-- if there is a list for chapter defined, 1. Else 0 -->
    <!-- look in content.xml or styles.xml -->
    <xsl:for-each select="document('content.xml')">
      <xsl:choose>
        <xsl:when test="key('outlined-styles','')">1</xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="document('styles.xml')">
            <xsl:choose>
              <xsl:when test="key('outlined-styles','')">1</xsl:when>
              <xsl:otherwise>0</xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:variable>

  <!--generate numbering definitions: abstract numbering w:abstractNumbering and numbering instances w:num -->
  <xsl:template name="numbering">
    <w:numbering xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006"
      xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:o12="http://schemas.microsoft.com/office/2004/7/core"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
      xmlns:v="urn:schemas-microsoft-com:vml"
      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:w10="urn:schemas-microsoft-com:office:word"
      xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">

      <xsl:for-each select="document('content.xml')">
        <xsl:for-each select="key('bullets','')">
          <xsl:call-template name="numPicBullet"/>
        </xsl:for-each>
      </xsl:for-each>
      <xsl:for-each select="document('styles.xml')">
        <xsl:for-each select="key('bullets','')">
          <xsl:call-template name="numPicBullet"/>
        </xsl:for-each>
      </xsl:for-each>
      <!-- default list for chapters -->
      <xsl:apply-templates
        select="document('styles.xml')/office:document-styles/office:styles/text:outline-style"
        mode="numbering"/>
      <!-- list for chapter numbering -->
      <xsl:if test="$chapterListCount = 1">
        <xsl:call-template name="BuildChapterList"/>
      </xsl:if>
      <!-- other lists -->
      <xsl:apply-templates
        select="document('content.xml')/office:document-content/office:automatic-styles/text:list-style"
        mode="numbering"/>
      <xsl:apply-templates
        select="document('styles.xml')/office:document-styles/office:styles/text:list-style"
        mode="numbering">
        <xsl:with-param name="offset" select="$automaticListStylesCount + $chapterListCount"/>
      </xsl:apply-templates>
      <xsl:apply-templates
        select="document('styles.xml')/office:document-styles/office:automatic-styles/text:list-style"
        mode="numbering">
        <xsl:with-param name="offset" select="$automaticListStylesCount + $chapterListCount"/>
      </xsl:apply-templates>
      <xsl:for-each select="document('content.xml')">
        <xsl:apply-templates select="key('restarting-lists','')" mode="numbering">
          <xsl:with-param name="offset"
            select="$automaticListStylesCount + $stylesListStyleCount + $chapterListCount"/>
        </xsl:apply-templates>
      </xsl:for-each>

      <xsl:apply-templates
        select="document('styles.xml')/office:document-styles/office:styles/text:outline-style"
        mode="num"/>
      <xsl:if test="$chapterListCount = 1">
        <xsl:call-template name="BuildChapterListNum"/>
      </xsl:if>
      <xsl:apply-templates
        select="document('content.xml')/office:document-content/office:automatic-styles/text:list-style"
        mode="num"/>
      <xsl:apply-templates
        select="document('styles.xml')/office:document-styles/office:styles/text:list-style"
        mode="num">
        <xsl:with-param name="offset" select="$automaticListStylesCount + $chapterListCount"/>
      </xsl:apply-templates>
      <xsl:apply-templates
        select="document('styles.xml')/office:document-styles/office:automatic-styles/text:list-style"
        mode="num">
        <xsl:with-param name="offset" select="$automaticListStylesCount + $chapterListCount"/>
      </xsl:apply-templates>
      <xsl:for-each select="document('content.xml')">
        <xsl:apply-templates select="key('restarting-lists','')" mode="num">
          <xsl:with-param name="offset"
            select="$automaticListStylesCount + $stylesListStyleCount + $chapterListCount"/>
        </xsl:apply-templates>
      </xsl:for-each>
    </w:numbering>
  </xsl:template>

  <xsl:template match="text:list-style" mode="numbering">
    <xsl:param name="offset" select="0"/>
    <w:abstractNum w:abstractNumId="{count(preceding-sibling::text:list-style)+2+$offset}">
      <xsl:apply-templates
        select="text:list-level-style-number|text:list-level-style-bullet|text:list-level-style-image"
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
  <xsl:template
    match="text:list-level-style-bullet|text:list-level-style-number|text:list-level-style-image"
    mode="numbering">
    <xsl:param name="style" select="'none'"/>

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

        <!-- image list -->
        <xsl:if test="name()='text:list-level-style-image' ">
          <w:start w:val="1"/>
          <w:numFmt w:val="bullet"/>
          <w:lvlText w:val=""/>
          <w:lvlPicBulletId>
            <xsl:attribute name="w:val">
              <xsl:call-template name="GetBulletId"/>
            </xsl:attribute>
          </w:lvlPicBulletId>
        </xsl:if>

        <!--numbered list-->
        <xsl:if test="name() = 'text:list-level-style-number' ">
          <xsl:call-template name="InsertNumberedListProperties">
            <xsl:with-param name="style" select="$style"/>
          </xsl:call-template>
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
            <xsl:call-template name="InsertListRunProperties"/>
          </xsl:otherwise>
        </xsl:choose>

      </w:lvl>
    </xsl:if>

  </xsl:template>

  <!-- Picture numbering symbol -->
  <xsl:template name="numPicBullet">
    <w:numPicBullet>
      <xsl:attribute name="w:numPicBulletId">
        <xsl:call-template name="GetBulletId"/>
      </xsl:attribute>
      <w:pict>
        <v:shape id="_x0000_i1032" type="#_x0000_t75" o:bullet="t">
          <xsl:attribute name="style">
            <xsl:if test="style:list-level-properties/@fo:width">
              <xsl:text>width:</xsl:text>
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:value-of select="style:list-level-properties/@fo:width"/>
                </xsl:with-param>
              </xsl:call-template>
              <xsl:text>pt;</xsl:text>
            </xsl:if>
            <xsl:if test="style:list-level-properties/@fo:height">
              <xsl:text>height:</xsl:text>
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:value-of select="style:list-level-properties/@fo:height"/>
                </xsl:with-param>
              </xsl:call-template>
              <xsl:text>pt</xsl:text>
            </xsl:if>
          </xsl:attribute>
          <v:imagedata>
            <xsl:attribute name="r:id">
              <xsl:value-of select="generate-id(.)"/>
            </xsl:attribute>
            <xsl:attribute name="o:title">
              <xsl:value-of select="@xlink:href"/>
            </xsl:attribute>
          </v:imagedata>
        </v:shape>
      </w:pict>
    </w:numPicBullet>
  </xsl:template>

  <!-- compute id of a bullet. Must be called from list-level-style-image context -->
  <xsl:template name="GetBulletId">
    <xsl:choose>
      <!-- first, if bullet is in styles -->
      <xsl:when test="ancestor::office:styles">
        <xsl:value-of
          select="1+count(preceding-sibling::text:list-level-style-image)+count(parent::node()/preceding-sibling::text:list-style/text:list-level-style-image)"
        />
      </xsl:when>
      <!-- if list is in automatic styles of styles.xml -->
      <xsl:when test="ancestor::office:automatic-styles and ancestor::office:document-styles">
        <xsl:variable name="bulletsCount">
          <xsl:value-of
            select="count(/office:document-styles/office:styles/text:list-style/text:list-level-style-image)"
          />
        </xsl:variable>
        <xsl:value-of
          select="1+$bulletsCount+count(preceding-sibling::text:list-level-style-image)+count(parent::node()/preceding-sibling::text:list-style/text:list-level-style-image)"
        />
      </xsl:when>
      <xsl:otherwise>
        <!-- list is in content.xml -->
        <xsl:variable name="bulletsCount">
          <xsl:for-each select="document('styles.xml')">
            <xsl:value-of select="count(key('bullets',''))"/>
          </xsl:for-each>
        </xsl:variable>
        <xsl:value-of
          select="1+$bulletsCount+count(preceding-sibling::text:list-level-style-image)+count(parent::node()/preceding-sibling::text:list-style/text:list-level-style-image)"
        />
      </xsl:otherwise>
    </xsl:choose>
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
    <!-- report lost min-label-distance -->
    <xsl:if test="$minLabelDistanceTwip != 0">
      <xsl:message terminate="no">feedback:Distance between numbering and text</xsl:message>
    </xsl:if>

    <!-- If @text:display-levels is defined and greater than 1, the tabs may not be well converted. -->
    <xsl:if
      test="not(@text:display-levels &gt; 1) or style:list-level-properties/@text:min-label-distance">
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
    </xsl:if>

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

  <!-- list run properties are created as in w:style -->
  <xsl:template name="InsertListRunProperties">
    <xsl:if test="@text:style-name">
      <xsl:variable name="styleName" select="@text:style-name"/>
      <xsl:for-each select="document('styles.xml')">
        <xsl:if test="key('styles',$styleName)/style:text-properties">
          <w:rPr>
            <xsl:apply-templates select="key('styles',$styleName)/style:text-properties" mode="rPr"
            />
          </w:rPr>
        </xsl:if>
      </xsl:for-each>
    </xsl:if>
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
      <xsl:when test="@text:bullet-char = '' "></xsl:when>
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
      <xsl:when test="$char  = ''  or $char ='' ">
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
    <xsl:param name="style" select="'none'"/>
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
    <!-- level style -->
    <xsl:choose>
      <xsl:when test="$style != 'none' ">
        <w:pStyle w:val="{$style}"/>
      </xsl:when>
      <xsl:when test="@text:style-name">
        <w:pStyle w:val="{@text:style-name}"/>
      </xsl:when>
      <xsl:otherwise/>
    </xsl:choose>
    <!-- content between numbered item and text -->
    <w:suff>
      <xsl:attribute name="w:val">
        <xsl:choose>
          <!-- to avoid problems when tab-stop cannot be evaluated precisely -->
          <xsl:when
            test="@text:display-levels &gt; 1 and not(style:list-level-properties/@text:min-label-distance)"
            >space</xsl:when>
          <!-- if no numbering defined -->
          <xsl:when test="@style:num-format = '' ">nothing</xsl:when>
          <!-- if numbering is too large -->
          <xsl:when
            test="string-length(@style:num-prefix) + string-length(@style:num-suffix) &gt; 2"
            >space</xsl:when>
          <!-- default value -->
          <xsl:otherwise>tab</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </w:suff>
    <!-- numbering format -->
    <w:lvlText>
      <xsl:attribute name="w:val">
        <xsl:if test="@style:num-format != '' ">
          <xsl:value-of select="@style:num-prefix"/>
          <xsl:call-template name="GetLevelText">
            <xsl:with-param name="displayLevels">
              <xsl:choose>
                <xsl:when test="@text:display-levels">
                  <xsl:value-of select="@text:display-levels"/>
                </xsl:when>
                <xsl:otherwise>1</xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
            <xsl:with-param name="level" select="@text:level"/>
            <xsl:with-param name="listStyleName">
              <xsl:choose>
                <xsl:when test="ancestor::text:outline-style">text:outline-style</xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="ancestor::text:list-style/@style:name"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
          </xsl:call-template>
          <xsl:value-of select="@style:num-suffix"/>
        </xsl:if>
      </xsl:attribute>
    </w:lvlText>
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
      <xsl:call-template name="InsertLevels"/>
    </w:abstractNum>
  </xsl:template>

  <!-- insert ten levels of list, even those not defined in original doc, in order to change default parameters -->
  <xsl:template name="InsertLevels">
    <xsl:param name="level" select="0"/>
    <xsl:param name="buildOutline" select="'true'"/>

    <xsl:if test="$level &lt; 9">
      <xsl:choose>
        <xsl:when test="text:outline-level-style/@text:level = $level+1">
          <xsl:for-each select="text:outline-level-style[@text:level = $level+1]">
            <w:lvl w:ilvl="{number(@text:level)-1}">
              <!--numbered list properties: num start, level, num type-->
              <xsl:call-template name="InsertNumberedListProperties"/>
              <!--justification-->
              <w:lvlJc w:val="left"/>
              <!--  list paragraph properties: indent -->
              <w:pPr>
                <xsl:call-template name="InsertListParagraphProperties"/>
              </w:pPr>
              <xsl:call-template name="InsertRunProperties"/>
            </w:lvl>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <!-- insert level with no numbering defined -->
          <w:lvl w:ilvl="{$level}">
            <w:numFmt w:val="none"/>
            <w:suff w:val="nothing"/>
            <w:lvlText w:val=""/>
            <w:pPr>
              <w:ind w:left="0"/>
            </w:pPr>
          </w:lvl>
        </xsl:otherwise>
      </xsl:choose>
      <xsl:if test="$buildOutline = 'true'">
        <xsl:call-template name="InsertLevels">
          <xsl:with-param name="level" select="$level+1"/>
        </xsl:call-template>
      </xsl:if>
    </xsl:if>
  </xsl:template>

  <!-- build a list for chapter numbering, called when ther is default outline level in a style -->
  <xsl:template name="BuildChapterList">
    <w:abstractNum w:abstractNumId="2">
      <xsl:call-template name="InsertChapterListElements"/>
    </w:abstractNum>
  </xsl:template>

  <xsl:template name="InsertChapterListElements">
    <xsl:param name="level" select="0"/>
    <xsl:if test="$level &lt; 9">
      <!-- find list name associated to level -->
      <xsl:variable name="listStyleName">
        <xsl:call-template name="GetListStyleName">
          <xsl:with-param name="level" select="$level"/>
        </xsl:call-template>
      </xsl:variable>
      <!-- insert element of given level -->
      <xsl:choose>
        <xsl:when test="$listStyleName != 'none'">
          <xsl:for-each select="document('content.xml')">
            <xsl:choose>
              <xsl:when
                test="key('list-style', $listStyleName)/text:list-level-style-number[@text:level = $level+1]">
                <xsl:apply-templates
                  select="key('list-style', $listStyleName)/text:list-level-style-number[@text:level = $level+1]"
                  mode="numbering">
                  <xsl:with-param name="style">
                    <xsl:call-template name="GetLevelStyleName">
                      <xsl:with-param name="level" select="$level"/>
                    </xsl:call-template>
                  </xsl:with-param>
                </xsl:apply-templates>
              </xsl:when>
              <xsl:otherwise>
                <xsl:for-each select="document('styles.xml')">
                  <xsl:if
                    test="key('list-style', $listStyleName)/text:list-level-style-number[@text:level = $level+1]">
                    <xsl:apply-templates
                      select="key('list-style', $listStyleName)/text:list-level-style-number[@text:level = $level+1]"
                      mode="numbering">
                      <xsl:with-param name="style">
                        <xsl:call-template name="GetLevelStyleName">
                          <xsl:with-param name="level" select="$level"/>
                        </xsl:call-template>
                      </xsl:with-param>
                    </xsl:apply-templates>
                  </xsl:if>
                </xsl:for-each>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <!-- insert level with outline numbering (default) -->
          <xsl:for-each
            select="document('styles.xml')/office:document-styles/office:styles/text:outline-style">
            <xsl:call-template name="InsertLevels">
              <xsl:with-param name="level" select="$level"/>
              <xsl:with-param name="buildOutline" select="'false'"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
      <!-- insert next level -->
      <xsl:call-template name="InsertChapterListElements">
        <xsl:with-param name="level" select="$level + 1"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

  <!-- find the style name associated to a chapter level -->
  <xsl:template name="GetLevelStyleName">
    <xsl:param name="level"/>
    <xsl:for-each select="document('content.xml')">
      <xsl:choose>
        <xsl:when
          test="key('outlined-styles','')[@style:default-outline-level = $level+1]/@style:name">
          <xsl:value-of
            select="key('outlined-styles','')[@style:default-outline-level = $level+1]/@style:name"
          />
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="document('styles.xml')">
            <xsl:choose>
              <xsl:when
                test="key('outlined-styles','')[@style:default-outline-level = $level+1]/@style:name">
                <xsl:value-of
                  select="key('outlined-styles','')[@style:default-outline-level = $level+1]/@style:name"
                />
              </xsl:when>
              <xsl:otherwise>none</xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <!-- find the list associated to a chapter level -->
  <xsl:template name="GetListStyleName">
    <xsl:param name="level"/>
    <xsl:for-each select="document('content.xml')">
      <xsl:choose>
        <xsl:when
          test="key('outlined-styles','')[@style:default-outline-level = $level+1]/@style:list-style-name">
          <xsl:value-of
            select="key('outlined-styles','')[@style:default-outline-level = $level+1]/@style:list-style-name"
          />
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="document('styles.xml')">
            <xsl:choose>
              <xsl:when
                test="key('outlined-styles','')[@style:default-outline-level = $level+1]/@style:list-style-name">
                <xsl:value-of
                  select="key('outlined-styles','')[@style:default-outline-level = $level+1]/@style:list-style-name"
                />
              </xsl:when>
              <xsl:otherwise>none</xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="BuildChapterListNum">
    <w:num w:numId="2">
      <w:abstractNumId w:val="2"/>
    </w:num>
  </xsl:template>

  <!-- Generate the format string for multiple levels -->
  <!--  level : current level -->
  <!-- displayLevels : number of preceding levels displayed at this level -->
  <xsl:template name="GetLevelText">
    <xsl:param name="displayLevels" select="1"/>
    <xsl:param name="level"/>
    <xsl:param name="listStyleName"/>

    <xsl:if test="$displayLevels &gt; 0">
      <xsl:call-template name="GetLevelText">
        <xsl:with-param name="displayLevels" select="$displayLevels - 1"/>
        <xsl:with-param name="level" select="$level - 1"/>
        <xsl:with-param name="listStyleName" select="$listStyleName"/>
      </xsl:call-template>
      <!-- levels with no formatting are not displayed -->
      <xsl:choose>
        <xsl:when test="$listStyleName='text:outline-style'">
          <xsl:if
            test="parent::text:outline-style/text:outline-level-style[@text:level = $level]/@style:num-format != '' ">
            <xsl:if test="$displayLevels &gt; 1">.</xsl:if>%<xsl:value-of select="$level"/>
          </xsl:if>
        </xsl:when>
        <xsl:otherwise>
          <xsl:if
            test="key('list-style', $listStyleName)/text:list-level-style-number[@text:level = $level]/@style:num-format != '' ">
            <xsl:if test="$displayLevels &gt; 1">.</xsl:if>%<xsl:value-of select="$level"/>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>
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
              select="text:list-level-style-number|text:list-level-style-bullet|text:list-level-style-image"
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
                    select="text:list-level-style-number|text:list-level-style-bullet|text:list-level-style-image"
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
          select="count(parent::text:list/preceding-sibling::text:list[text:list-item/@text:start-value])+2+$chapterListCount+$stylesListStyleCount + $automaticListStylesCount"
        />
      </xsl:when>

      <!-- look for this list style into content.xml -->
      <xsl:when
        test="key('list-style', $styleName) and not(ancestor::style:header) and not(ancestor::style:footer)">
        <xsl:value-of
          select="count(key('list-style', $styleName)/preceding-sibling::text:list-style)+2+$chapterListCount"
        />
      </xsl:when>

      <!-- otherwise, look into styles.xml (add the offset) -->
      <xsl:otherwise>
        <xsl:for-each select="document('styles.xml')">
          <xsl:value-of
            select="count(key('list-style', $styleName)/preceding-sibling::text:list-style)+2+$chapterListCount+$automaticListStylesCount"
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
              select="key('list-style', $listStyleName)[1]/*[@text:level = $level+1]/style:list-level-properties/@*[name()=$attribute]"
            />
          </xsl:when>
          <xsl:otherwise>
            <!-- look into styles.xml -->
            <xsl:for-each select="document('styles.xml')">
              <xsl:choose>
                <xsl:when
                  test="key('list-style', $listStyleName)[1]/*[@text:level = $level+1]/style:list-level-properties/@*[name()=$attribute]">
                  <xsl:value-of
                    select="key('list-style', $listStyleName)[1]/*[@text:level = $level+1]/style:list-level-properties/@*[name()=$attribute]"
                  />
                </xsl:when>
                <!-- default value -->
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

    <!-- A text:note-body does not depend on its parent list indentation anymore -->
    <xsl:if test="not(parent::text:note-body) and ancestor-or-self::text:list">
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

      <!-- numbered of levels displayed -->
      <xsl:variable name="displayedLevels">
        <xsl:choose>
          <!-- look into content.xml -->
          <xsl:when
            test="key('list-style', $listStyleName)[1]/*[@text:level = $level+1]/@text:display-levels">
            <xsl:value-of
              select="key('list-style', $listStyleName)[1]/*[@text:level = $level+1]/@text:display-levels"
            />
          </xsl:when>
          <xsl:otherwise>
            <!-- look into styles.xml -->
            <xsl:for-each select="document('styles.xml')">
              <xsl:choose>
                <xsl:when
                  test="key('list-style', $listStyleName)[1]/*[@text:level = $level+1]/@text:display-levels">
                  <xsl:value-of
                    select="key('list-style', $listStyleName)[1]/*[@text:level = $level+1]/@text:display-levels"
                  />
                </xsl:when>
                <!-- default value -->
                <xsl:otherwise>1</xsl:otherwise>
              </xsl:choose>
            </xsl:for-each>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <!-- report lost distance because of displayed levels -->
      <xsl:if test="$displayedLevels &gt; 1">
        <xsl:message terminate="no">feedback:Distance between numbering and text</xsl:message>
      </xsl:if>

      <!-- Minimum width of a number -->
      <xsl:variable name="minLabelWidthTwip">
        <xsl:call-template name="ComputeNumberingIndent">
          <xsl:with-param name="attribute" select="'text:min-label-width'"/>
          <xsl:with-param name="level" select="$level"/>
          <xsl:with-param name="listStyleName" select="$listStyleName"/>
        </xsl:call-template>
      </xsl:variable>

      <!-- Space before the number for all paragraphs at this level. -->
      <xsl:variable name="spaceBeforeTwip">
        <xsl:call-template name="ComputeNumberingIndent">
          <xsl:with-param name="attribute" select="'text:space-before'"/>
          <xsl:with-param name="level" select="$level"/>
          <xsl:with-param name="listStyleName" select="$listStyleName"/>
        </xsl:call-template>
      </xsl:variable>

      <!-- Minimum distance between the number and the text -->
      <xsl:variable name="minLabelDistanceTwip">
        <xsl:call-template name="ComputeNumberingIndent">
          <xsl:with-param name="attribute" select="'text:min-label-distance'"/>
          <xsl:with-param name="level" select="$level"/>
          <xsl:with-param name="listStyleName" select="$listStyleName"/>
        </xsl:call-template>
      </xsl:variable>
      <!-- report lost min-label-distance -->
      <xsl:if test="$minLabelDistanceTwip != 0">
        <xsl:message terminate="no">feedback:Distance between numbering and text</xsl:message>
      </xsl:if>

      <!-- Override tabs of numbered elements ('num' tabs) if additional value (left margin or first line indent) defined. -->
      <xsl:if
        test="not(ancestor-or-self::text:list-header) and (self::text:list-item or not(preceding-sibling::node())) and ($addLeftIndent != 0 or $firstLineIndent != 0)">
        <!--If @text:display-levels is defined and greater than 1, the tabs may not be well converted.-->
        <xsl:if test="not($displayedLevels &gt; 1) or $minLabelDistanceTwip &gt; 0">
          <w:tabs>
            <!-- clear tab calculated without additional value -->
            <w:tab w:val="clear">
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
            <!-- replace with same distance + additional value -->
            <w:tab>
              <xsl:choose>
                <xsl:when test="$firstLineIndent != 0">
                  <xsl:attribute name="w:val">num</xsl:attribute>
                  <xsl:attribute name="w:pos">
                    <xsl:value-of
                      select="$minLabelDistanceTwip + $addLeftIndent + $firstLineIndent + $minLabelWidthTwip"
                    />
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:attribute name="w:val">num</xsl:attribute>
                  <xsl:attribute name="w:pos">
                    <xsl:choose>
                      <xsl:when test="$spaceBeforeTwip &lt; 0">
                        <xsl:choose>
                          <xsl:when test="$minLabelWidthTwip &lt; $minLabelDistanceTwip">
                            <xsl:value-of
                              select="$addLeftIndent + $minLabelWidthTwip - $spaceBeforeTwip + $minLabelDistanceTwip"
                            />
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of
                              select="$addLeftIndent + $minLabelWidthTwip - $spaceBeforeTwip"/>
                          </xsl:otherwise>
                        </xsl:choose>
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
                </xsl:otherwise>
              </xsl:choose>
            </w:tab>
            <!-- add tabs of current paragraph so as not to lose them in post-processing -->
            <xsl:if test="key('automatic-styles',$styleName)//style:tab-stop">
              <xsl:for-each select="key('automatic-styles',$styleName)/style:paragraph-properties">
                <xsl:call-template name="ComputeParagraphTabs"/>
              </xsl:for-each>
            </xsl:if>
            <xsl:if test="key('styles',$styleName)//style:tab-stop">
              <xsl:for-each select="key('styles',$styleName)/style:paragraph-properties">
                <xsl:call-template name="ComputeParagraphTabs"/>
              </xsl:for-each>
            </xsl:if>
          </w:tabs>
        </xsl:if>
      </xsl:if>

      <!-- insert indent with paragraph and numbering properties
        if any value affects the numbering style defined, or if list element has no numbering -->
      <xsl:choose>
        <!-- List element with numbering -->
        <xsl:when
          test="not(ancestor-or-self::text:list-header) and (self::text:list-item or not(preceding-sibling::node()))">
          <xsl:if test="$addLeftIndent != 0 or $addRightIndent !=  0 or $firstLineIndent != 0">
            <w:ind w:left="{$addLeftIndent + $spaceBeforeTwip + $minLabelWidthTwip}"
              w:right="{$addRightIndent}">
              <!-- first line and hanging indent -->
              <xsl:choose>
                <xsl:when test="($firstLineIndent - $minLabelWidthTwip) &gt; 0">
                  <xsl:attribute name="w:firstLine">
                    <xsl:value-of select="$firstLineIndent - $minLabelWidthTwip"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:choose>
                    <xsl:when test="$minLabelWidthTwip &lt; 0">
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
        <!-- Other list element -->
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

  <!-- Inserts the number of a list item -->
  <xsl:template name="InsertNumbering">
    <xsl:param name="level"/>
    <xsl:variable name="listStyleName">
      <xsl:choose>
        <xsl:when test="key('automatic-styles', @text:style-name)/@style:list-style-name">
          <xsl:value-of select="key('automatic-styles', @text:style-name)/@style:list-style-name"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:variable name="stylename" select="@text:style-name"/>
          <xsl:for-each select="document('styles.xml')">
            <xsl:if test="key('styles', $stylename)/@style:list-style-name">
              <xsl:value-of select="key('styles', $stylename)/@style:list-style-name"/>
            </xsl:if>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <!-- list element -->
      <xsl:when
        test="(self::text:list-item and not(child::node()[1][self::text:h])) or self::text:list-header or $listStyleName != '' ">
        <!-- test : do not number list-headers, and do not number paragraph in the middle of a list-item or list-header -->
        <xsl:if
          test="not(self::text:list-header) and not(parent::text:list-header) and not(preceding-sibling::node()[name() != 'text:list-item' and name() != 'text:list-header'])">
          <w:numPr>
            <xsl:choose>
              <xsl:when test="self::text:list-item or parent::text:list-item">
                <w:ilvl w:val="{$level}"/>
              </xsl:when>
              <xsl:otherwise>
                <w:ilvl w:val="{@text:outline-level - 1}"/>
              </xsl:otherwise>
            </xsl:choose>
            <w:numId>
              <xsl:attribute name="w:val">
                <xsl:choose>
                  <xsl:when test="self::text:list-item">
                    <xsl:call-template name="GetNumberingId">
                      <xsl:with-param name="styleName" select="ancestor::text:list/@text:style-name"
                      />
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:otherwise>

                    <xsl:call-template name="GetNumberingId">
                      <xsl:with-param name="styleName" select="$listStyleName"/>
                    </xsl:call-template>

                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </w:numId>
          </w:numPr>
        </xsl:if>
      </xsl:when>
      <xsl:otherwise>
        <!-- if element is a header, or list element with header properties (outline) -->
        <xsl:if test="self::text:h or (self::text:list-item and child::node()[1][self::text:h])">
          <xsl:call-template name="InsertHeaderNumbering">
            <xsl:with-param name="style-name" select="@text:style-name"/>
            <xsl:with-param name="list-style-name" select="$listStyleName"/>
          </xsl:call-template>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- insert numbering properties of header element -->
  <xsl:template name="InsertHeaderNumbering">
    <xsl:param name="style-name"/>
    <xsl:param name="list-style-name"/>
    <xsl:variable name="styleName">
      <xsl:choose>
        <xsl:when test="$style-name">
          <xsl:value-of select="$style-name"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:if test="text:h[1]/@text:style-name">
            <xsl:value-of select="text:h[1]/@text:style-name"/>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="listStyleName">
      <xsl:choose>
        <xsl:when test="not($list-style-name = '' or count($list-style-name) = 0)">
          <xsl:value-of select="$list-style-name"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:choose>
            <xsl:when test="key('automatic-styles', $styleName)/@style:list-style-name">
              <xsl:value-of select="key('automatic-styles', $styleName)/@style:list-style-name"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:for-each select="document('styles.xml')">
                <xsl:if test="key('styles', $styleName)/@style:list-style-name">
                  <xsl:value-of select="key('styles', $styleName)/@style:list-style-name"/>
                </xsl:if>
              </xsl:for-each>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="key('automatic-styles', $styleName)/@style:default-outline-level">
        <w:numPr>
          <w:ilvl w:val="{key('automatic-styles', $styleName)/@style:default-outline-level - 1}"/>
          <xsl:choose>
            <xsl:when test="$listStyleName != '' ">
              <w:numId w:val="2"/>
            </xsl:when>
            <xsl:otherwise>
              <w:numId w:val="1"/>
            </xsl:otherwise>
          </xsl:choose>
        </w:numPr>
      </xsl:when>
      <xsl:otherwise>
        <xsl:for-each select="document('styles.xml')">
          <xsl:choose>
            <xsl:when test="key('styles', $styleName)/@style:default-outline-level">
              <w:numPr>
                <w:ilvl w:val="{key('styles', $styleName)/@style:default-outline-level - 1}"/>
                <xsl:choose>
                  <xsl:when test="$listStyleName != '' ">
                    <w:numId w:val="2"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <w:numId w:val="1"/>
                  </xsl:otherwise>
                </xsl:choose>
              </w:numPr>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if
                test="@text:outline-level &lt;= 9 and office:document-styles/office:styles/text:outline-style/text:outline-level-style/@style:num-format !='' ">
                <w:numPr>
                  <w:ilvl w:val="{@text:outline-level - 1}"/>
                  <w:numId w:val="1"/>
                </w:numPr>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>



  <!-- Inserts the outline level of a heading if needed -->
  <xsl:template name="InsertOutlineLevel">
    <xsl:param name="node"/>
    <!-- List item are first considered if exist than heading style-->
    <xsl:variable name="stylename" select="./@text:style-name"/>
    <xsl:variable name="list"
      select="//office:document-content/office:automatic-styles/style:style[@style:name = $stylename]"/>
    <xsl:if
      test="$node[self::text:h and not($list/@style:list-style-name) and (not(parent::text:list-item) or position() > 1) ]">
      <xsl:choose>
        <xsl:when test="not($node/@text:outline-level)">
          <w:outlineLvl w:val="0"/>
        </xsl:when>
        <xsl:when test="$node/@text:outline-level &lt;= 9">
          <w:outlineLvl w:val="{$node/@text:outline-level}"/>
        </xsl:when>
        <xsl:otherwise>
          <w:outlineLvl w:val="9"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <xsl:if test="$node[child::text:toc-mark-start]">
      <w:outlineLvl w:val="{$node/text:toc-mark-start/@text:outline-level}"/>
    </xsl:if>
  </xsl:template>


  <!-- suppress line numbering if required by configuration -->
  <xsl:template name="InsertSupressLineNumbers">
    <xsl:if
      test="document('styles.xml')/office:document-styles/office:styles/text:linenumbering-configuration/@text:count-empty-lines='false' and not(descendant::text())">
      <w:suppressLineNumbers w:val="true"/>
    </xsl:if>
  </xsl:template>


</xsl:stylesheet>

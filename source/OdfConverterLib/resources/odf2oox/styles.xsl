<?xml version="1.0" encoding="UTF-8" ?>
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
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/3/main"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  exclude-result-prefixes="office style fo text draw number svg">

  <xsl:variable name="asianLayoutId">1</xsl:variable>

  <!-- keys definition -->
  <xsl:key name="styles" match="office:styles/style:style" use="@style:name"/>

  <xsl:template name="styles">
    <w:styles>
      <w:docDefaults>
        <!-- Default text properties -->
        <xsl:variable name="paragraphDefaultStyle"
          select="document('styles.xml')/office:document-styles/office:styles/style:default-style[@style:family='paragraph']"/>
        <w:rPrDefault>
          <w:rPr>
            <xsl:apply-templates select="$paragraphDefaultStyle/style:text-properties" mode="rPr"/>
          </w:rPr>
        </w:rPrDefault>
        <!-- Default paragraph properties -->
        <w:pPrDefault>
          <w:pPr>
            <xsl:apply-templates select="$paragraphDefaultStyle/style:paragraph-properties"
              mode="pPr"/>
          </w:pPr>
        </w:pPrDefault>
      </w:docDefaults>
      <xsl:apply-templates select="document('styles.xml')/office:document-styles/office:styles"
        mode="styles"/>
      <xsl:apply-templates
        select="document('content.xml')/office:document-content/office:automatic-styles"
        mode="styles"/>
      <xsl:apply-templates
        select="document('styles.xml')/office:document-styles/office:automatic-styles" mode="styles"
      />
    </w:styles>
  </xsl:template>

  <!-- Computes the prefixed style name (for automatic styles) -->
  <xsl:template name="GetPrefixedStyleName">
    <xsl:param name="styleName"/>
    <xsl:choose>
      <xsl:when test="key('automatic-styles',$styleName)">
        <xsl:choose>
          <xsl:when test="ancestor::office:document-styles/office:automatic-styles">
            <xsl:value-of select="concat('STYLE_', $styleName)"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat('CONTENT_', $styleName)"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$styleName"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- styles -->
  <xsl:template match="style:style" mode="styles">

    <xsl:variable name="prefixedStyleName">
      <xsl:call-template name="GetPrefixedStyleName">
        <xsl:with-param name="styleName" select="@style:name"/>
      </xsl:call-template>
    </xsl:variable>

    <w:style>
      <xsl:attribute name="w:styleId">
        <xsl:value-of select="$prefixedStyleName"/>
      </xsl:attribute>
      <xsl:choose>
        <xsl:when test="@style:family = 'text' ">
          <xsl:attribute name="w:type">character</xsl:attribute>
        </xsl:when>
        <xsl:when test="@style:family = 'graphic' ">
          <xsl:attribute name="w:type">paragraph</xsl:attribute>
        </xsl:when>
        <xsl:when test="@style:family = 'section' ">
          <xsl:attribute name="w:type">paragraph</xsl:attribute>
        </xsl:when>
        <xsl:when test="@style:family = 'paragraph' ">
          <xsl:attribute name="w:type">paragraph</xsl:attribute>
        </xsl:when>
        <xsl:when test="@style:family = 'table' ">
          <xsl:attribute name="w:type">table</xsl:attribute>
        </xsl:when>
        <xsl:when test="@style:family = 'table-cell' ">
          <xsl:attribute name="w:type">table</xsl:attribute>
        </xsl:when>
        <xsl:when test="@style:family = 'table-row' ">
          <xsl:attribute name="w:type">table</xsl:attribute>
        </xsl:when>
        <xsl:when test="@style:family = 'table-column' ">
          <xsl:attribute name="w:type">table</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="w:type">character</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>

      <!-- Nested elements-->

      <xsl:choose>
        <xsl:when test="@style:display-name">
          <w:name w:val="{@style:display-name}"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:if test="@style:name">
            <w:name w:val="{$prefixedStyleName}"/>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:if test="@style:parent-style-name">
        <w:basedOn w:val="{@style:parent-style-name}"/>
      </xsl:if>
      <xsl:if test="@style:next-style-name">
        <w:next w:val="{@style:next-style-name}"/>
      </xsl:if>

      <!-- Automatic-styles are not displayed -->
      <xsl:if test="key('automatic-styles',@style:name)">
        <w:hidden/>
      </xsl:if>

      <!-- COMMENT: what is this for? -->
      <w:qFormat/>

      <w:pPr>
        <xsl:apply-templates mode="pPr"/>
      </w:pPr>

      <w:rPr>
        <xsl:apply-templates mode="rPr"/>
      </w:rPr>

    </w:style>
  </xsl:template>

  <!-- ignored (default styles already managed) -->
  <xsl:template match="style:default-style" mode="styles"/>

  <!-- paragraph properties -->
  <xsl:template match="style:paragraph-properties" mode="pPr">

    <xsl:if test="@fo:keep-with-next='always'">
      <w:keepNext/>
    </xsl:if>

    <xsl:if test="@fo:keep-together='always'">
      <w:keepLines/>
    </xsl:if>

    <xsl:if test="@fo:break-before='page'">
      <w:pageBreakBefore/>
    </xsl:if>
    <xsl:if test="parent::node()/@style:master-page-name != ''">
      <w:pageBreakBefore/>
    </xsl:if>

    <xsl:choose>
      <xsl:when test="@fo:widows != '0' or @fo:orphans != '0'">
        <w:widowControl w:val="on"/>
      </xsl:when>
      <xsl:otherwise>
        <w:widowControl w:val="off"/>
      </xsl:otherwise>
    </xsl:choose>

    <!-- border color + padding  -->
    <xsl:if test="@fo:border and @fo:border != 'none' ">
      <w:pBdr>
        <xsl:call-template name="InsertBorders">
          <xsl:with-param name="allSides">true</xsl:with-param>
        </xsl:call-template>
      </w:pBdr>
    </xsl:if>
    <xsl:if test="@fo:border-top or @fo:border-left or @fo:border-bottom or @fo:border-right">
      <w:pBdr>
        <xsl:call-template name="InsertBorders">
          <xsl:with-param name="allSides">false</xsl:with-param>
        </xsl:call-template>
      </w:pBdr>
    </xsl:if>

    <!-- background color -->
    <xsl:if test="@fo:background-color and (@fo:background-color != 'transparent')">
      <w:shd>
        <xsl:attribute name="w:val">clear</xsl:attribute>
        <xsl:attribute name="w:color">auto</xsl:attribute>
        <xsl:attribute name="w:fill">
          <xsl:value-of
            select="substring(@fo:background-color, 2, string-length(@fo:background-color) -1)"/>
        </xsl:attribute>
      </w:shd>
    </xsl:if>

    <!-- tabs -->
    <xsl:if test="style:tab-stops/style:tab-stop">
      <w:tabs>
        <xsl:for-each select="style:tab-stops/style:tab-stop">
          <xsl:call-template name="tabStop"/>
        </xsl:for-each>
        <xsl:variable name="styleName">
          <xsl:value-of select="parent::node()/@style:name"/>
        </xsl:variable>
        <xsl:variable name="parentstyleName">
          <xsl:value-of select="parent::node()/@style:parent-style-name"/>
        </xsl:variable>
        <xsl:for-each select="key('styles',$parentstyleName)//style:tab-stop">
          <xsl:if
            test="not(@style:position=key('automatic-styles',$styleName)//style:tab-stop/@style:position)">
            <xsl:call-template name="tabStop">
              <xsl:with-param name="styleType">clear</xsl:with-param>
            </xsl:call-template>
          </xsl:if>
        </xsl:for-each>
      </w:tabs>
    </xsl:if>

    <!-- Override hyphenation -->
    <xsl:choose>
      <xsl:when test="parent::node()/style:text-properties/@fo:hyphenate='true'">
        <w:suppressAutoHyphens w:val="false"/>
      </xsl:when>
      <xsl:otherwise>
        <w:suppressAutoHyphens w:val="true"/>
      </xsl:otherwise>
    </xsl:choose>

    <xsl:if test="@style:text-autospace">
      <w:autoSpaceDN>
        <xsl:choose>
          <xsl:when test="@style:text-autospace='none'">
            <xsl:attribute name="w:val">off</xsl:attribute>
          </xsl:when>
          <xsl:when test="@style:text-autospace='ideograph-alpha'">
            <xsl:attribute name="w:val">on</xsl:attribute>
          </xsl:when>
        </xsl:choose>
      </w:autoSpaceDN>
    </xsl:if>

    <xsl:if
      test="@style:line-height-at-least or @fo:line-height or @fo:margin-bottom or @fo:margin-top or @style:line-spacing">
      <w:spacing>
        <xsl:if test="@style:line-height-at-least">
          <xsl:attribute name="w:lineRule">atLeast</xsl:attribute>
          <xsl:attribute name="w:line">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length" select="@style:line-height-at-least"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="@style:line-spacing">
          <xsl:variable name="spacing">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length" select="@style:line-spacing"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:attribute name="w:lineRule">auto</xsl:attribute>
          <xsl:attribute name="w:line">
            <xsl:value-of select="240+$spacing"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="contains(@fo:line-height, '%')">
          <xsl:attribute name="w:lineRule">auto</xsl:attribute>
          <xsl:attribute name="w:line">
            <xsl:value-of
              select="round(number(substring-before(@fo:line-height, '%')) * 240 div 100)"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="contains(@fo:line-height, 'cm')">
          <xsl:attribute name="w:lineRule">exact</xsl:attribute>
          <xsl:attribute name="w:line">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length" select="@fo:line-height"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="@fo:margin-bottom">
          <xsl:attribute name="w:after">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length" select="@fo:margin-bottom"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="@fo:margin-top">
          <xsl:attribute name="w:before">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length" select="@fo:margin-top"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
      </w:spacing>
    </xsl:if>

    <xsl:if
      test="@fo:margin-left or @fo:margin-right or @fo:text-indent or @text:space-before or @fo:padding or @fo:padding-left or @fo:padding-right">
      <w:ind>
        <!-- If there is a border, the indent will have to take it into consideration. -->
        <xsl:variable name="leftBorderWidth">
          <xsl:choose>
            <xsl:when test="@fo:border">
              <xsl:call-template name="twips-measure">
                <xsl:with-param name="length" select="substring-before(@fo:border,' ')"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="@fo:border-left">
              <xsl:call-template name="twips-measure">
                <xsl:with-param name="length" select="substring-before(@fo:border-left,' ')"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="rightBorderWidth">
          <xsl:choose>
            <xsl:when test="@fo:border">
              <xsl:call-template name="twips-measure">
                <xsl:with-param name="length" select="substring-before(@fo:border,' ')"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="@fo:border-right">
              <xsl:call-template name="twips-measure">
                <xsl:with-param name="length" select="substring-before(@fo:border-right,' ')"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <!-- indent generated by the borders. Limited to 620 twip. -->
        <xsl:variable name="leftPadding">
          <xsl:choose>
            <xsl:when test="@fo:padding">
              <xsl:call-template name="indent-val">
                <xsl:with-param name="length" select="@fo:padding"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="@fo:padding-left">
              <xsl:call-template name="indent-val">
                <xsl:with-param name="length" select="@fo:padding-left"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="rightPadding">
          <xsl:choose>
            <xsl:when test="@fo:padding">
              <xsl:call-template name="indent-val">
                <xsl:with-param name="length" select="@fo:padding"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="@fo:padding-right">
              <xsl:call-template name="indent-val">
                <xsl:with-param name="length" select="@fo:padding-right"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <!-- paragraph margins -->
        <xsl:variable name="leftMargin">
          <xsl:call-template name="twips-measure">
            <xsl:with-param name="length" select="@fo:margin-left"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="rightMargin">
          <xsl:call-template name="twips-measure">
            <xsl:with-param name="length" select="@fo:margin-right"/>
          </xsl:call-template>
        </xsl:variable>

        <xsl:attribute name="w:left">
          <xsl:value-of select="$leftMargin + $leftPadding + $leftBorderWidth"/>
        </xsl:attribute>
        <xsl:attribute name="w:right">
          <xsl:value-of select="$rightMargin + $rightPadding + $rightBorderWidth"/>
        </xsl:attribute>

        <xsl:if test="@fo:text-indent and (@style:auto-text-indent!='true')">
          <xsl:attribute name="w:firstLine">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length" select="@fo:text-indent"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
        <!-- Trick to replace automatic indent for OOX -->
        <xsl:if test="@style:auto-text-indent='true'">
          <xsl:attribute name="w:firstLine">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length"
                select="parent::style:style/style:text-properties/@fo:font-size"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
      </w:ind>
    </xsl:if>

    <!-- TODO this should be modified when we will manage bidi properties -->
    <xsl:if test="@fo:text-align">
      <w:jc>
        <xsl:choose>
          <xsl:when test="@fo:text-align='center'">
            <xsl:attribute name="w:val">center</xsl:attribute>
          </xsl:when>
          <xsl:when test="@fo:text-align='start'">
            <xsl:attribute name="w:val">left</xsl:attribute>
          </xsl:when>
          <xsl:when test="@fo:text-align='end'">
            <xsl:attribute name="w:val">right</xsl:attribute>
          </xsl:when>
          <xsl:when test="@fo:text-align='justify'">
            <xsl:attribute name="w:val">both</xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="w:val">start</xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>

      </w:jc>
    </xsl:if>

    <xsl:if test="@style:vertical-align">
      <w:textAlignment>
        <xsl:choose>
          <xsl:when test="@style:vertical-align='bottom'">
            <xsl:attribute name="w:val">bottom</xsl:attribute>
          </xsl:when>
          <xsl:when test="@style:vertical-align='top'">
            <xsl:attribute name="w:val">top</xsl:attribute>
          </xsl:when>
          <xsl:when test="@style:vertical-align='middle'">
            <xsl:attribute name="w:val">center</xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="w:val">auto</xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </w:textAlignment>
    </xsl:if>
  </xsl:template>

  <!-- paragraph properties from text-properties -->
  <xsl:template match="style:text-properties" mode="pPr">
    <xsl:if test="not(parent::node()/style:paragraph-properties)">
      <xsl:choose>
        <xsl:when test="@fo:hyphenate='true'">
          <w:suppressAutoHyphens w:val="false"/>
        </xsl:when>
        <xsl:otherwise>
          <w:suppressAutoHyphens w:val="true"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
  </xsl:template>

  <!-- text-properties -->
  <xsl:template match="style:text-properties" mode="rPr">
    <xsl:param name="onlyToggle"/>

    <!-- toggle properties -->

    <xsl:if test="@fo:font-weight">
      <xsl:choose>
        <xsl:when test="@fo:font-weight != 'normal'">
          <w:b w:val="on"/>
        </xsl:when>
        <xsl:otherwise>
          <w:b w:val="off"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@fo:font-weight-complex">
      <xsl:choose>
        <xsl:when test="@style:font-weight-complex != 'normal'">
          <w:bCs w:val="on"/>
        </xsl:when>
        <xsl:otherwise>
          <w:bCs w:val="off"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@fo:font-style">
      <xsl:choose>
        <xsl:when test="@fo:font-style = 'italic' or @fo:font-style = 'oblique'">
          <w:i w:val="on"/>
        </xsl:when>
        <xsl:otherwise>
          <w:i w:val="off"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@fo:font-style-complex">
      <xsl:choose>
        <xsl:when test="@fo:font-style-complex = 'italic' or @fo:font-style-complex = 'oblique'">
          <w:iCs w:val="on"/>
        </xsl:when>
        <xsl:otherwise>
          <w:iCs w:val="off"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@fo:text-transform or @fo:font-variant">
      <xsl:choose>
        <xsl:when test="@fo:text-transform = 'uppercase'">
          <w:caps w:val="on"/>
        </xsl:when>
        <xsl:when test="@fo:text-transform = 'none' or @fo:font-variant = 'small-caps'">
          <w:caps w:val="off"/>
        </xsl:when>
        <!-- It could be also lowercase or capitalize in fo DTD, but it is not possible to set it via word 2007 interface -->
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@fo:font-variant">
      <xsl:choose>
        <xsl:when test="@fo:font-variant = 'small-caps'">
          <w:smallCaps w:val="on"/>
        </xsl:when>
        <xsl:when test="@fo:font-weight = 'normal'">
          <w:smallCaps w:val="off"/>
        </xsl:when>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@style:text-line-through-style != 'none'">
      <xsl:choose>
        <xsl:when test="@style:text-line-through-type = 'double'">
          <w:dstrike w:val="on"/>
        </xsl:when>
        <xsl:when test="@style:text-line-through-type = 'none'">
          <w:strike w:val="off"/>
        </xsl:when>
        <xsl:otherwise>
          <w:strike w:val="on"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@style:text-outline">
      <xsl:choose>
        <xsl:when test="@style:text-outline = 'true'">
          <w:outline w:val="on"/>
        </xsl:when>
        <xsl:when test="@style:text-outline = 'false'">
          <w:outline w:val="off"/>
        </xsl:when>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@fo:text-shadow">
      <xsl:choose>
        <xsl:when test="@fo:text-shadow = 'none'">
          <w:shadow w:val="off"/>
        </xsl:when>
        <xsl:otherwise>
          <w:shadow w:val="on"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@style:font-relief">
      <xsl:choose>
        <xsl:when test="@style:font-relief = 'embossed'">
          <w:emboss w:val="on"/>
        </xsl:when>
        <xsl:when test="@style:font-relief = 'engraved'">
          <w:imprint w:val="on"/>
        </xsl:when>
        <xsl:otherwise>
          <w:emboss w:val="off"/>
          <w:imprint w:val="off"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@text:display">
      <xsl:choose>
        <xsl:when test="@text:display = 'true'">
          <w:vanish w:val="on"/>
        </xsl:when>
        <xsl:when test="@text:display = 'none'">
          <w:vanish w:val="off"/>
        </xsl:when>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@style:text-underline-style">
      <xsl:choose>
        <xsl:when test="@style:text-underline-style != 'none' ">
          <w:u>
            <xsl:attribute name="w:val">
              <xsl:choose>
                <xsl:when test="@style:text-underline-style = 'dotted'">
                  <xsl:choose>
                    <xsl:when test="@style:text-underline-width = 'thick' ">dottedHeavy</xsl:when>
                    <xsl:when test="@style:text-underline-width = 'bold' ">dottedHeavy</xsl:when>
                    <xsl:otherwise>dotted</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:when test="@style:text-underline-style = 'dash' ">
                  <xsl:choose>
                    <xsl:when test="@style:text-underline-width = 'thick' ">dashedHeavy</xsl:when>
                    <xsl:when test="@style:text-underline-width = 'bold' ">dashedHeavy</xsl:when>
                    <xsl:otherwise>dash</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:when test="@style:text-underline-style = 'long-dash'">
                  <xsl:choose>
                    <xsl:when test="@style:text-underline-width = 'thick' ">dashLongHeavy</xsl:when>
                    <xsl:when test="@style:text-underline-width = 'bold' ">dashLongHeavy</xsl:when>
                    <xsl:otherwise>dashLong</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:when test="@style:text-underline-style = 'dot-dash' ">
                  <xsl:choose>
                    <xsl:when test="@style:text-underline-width = 'thick' ">dashDotHeavy</xsl:when>
                    <xsl:when test="@style:text-underline-width = 'bold' ">dashDotHeavy</xsl:when>
                    <xsl:otherwise>dotDash</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:when test="@style:text-underline-style = 'dot-dot-dash' ">
                  <xsl:choose>
                    <xsl:when test="@style:text-underline-width = 'thick' ">dashDotDotHeavy</xsl:when>
                    <xsl:when test="@style:text-underline-width = 'bold' ">dashDotDotHeavy</xsl:when>
                    <xsl:otherwise>dotDotDash</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:when test="@style:text-underline-style = 'wave' ">
                  <xsl:choose>
                    <xsl:when test="@style:text-underline-type = 'double' ">wavyDouble</xsl:when>
                    <xsl:when test="@style:text-underline-width = 'thick' ">wavyHeavy</xsl:when>
                    <xsl:when test="@style:text-underline-width = 'bold' ">wavyHeavy</xsl:when>
                    <xsl:otherwise>wave</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:choose>
                    <xsl:when test="@style:text-underline-type = 'double' ">double</xsl:when>
                    <xsl:when test="@style:text-underline-width = 'thick' ">thick</xsl:when>
                    <xsl:when test="@style:text-underline-width = 'bold' ">thick</xsl:when>
                    <xsl:when test="@style:text-underline-mode = 'skip-white-space' ">words</xsl:when>
                    <xsl:otherwise>single</xsl:otherwise>
                  </xsl:choose>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:if test="@style:text-underline-color">
              <xsl:attribute name="w:color">
                <xsl:choose>
                  <xsl:when test="@style:text-underline-color = 'font-color'">auto</xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of
                      select="substring(@style:text-underline-color, 2, string-length(@style:text-underline-color)-1)"
                    />
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </w:u>
        </xsl:when>
        <xsl:otherwise>
          <w:u w:val="none"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@fo:background-color">
      <w:highlight>
        <xsl:attribute name="w:val">
          <!-- Value must be a string (no RGB value authorized). -->
          <xsl:choose>
            <xsl:when test="@fo:background-color = 'transparent'">none</xsl:when>
            <xsl:when test="not(substring(@fo:background-color, 1,1)='#')">
              <xsl:value-of select="@fo:background-color"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="StringType-color">
                <xsl:with-param name="RGBcolor" select="@fo:background-color"/>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </w:highlight>
    </xsl:if>

    <xsl:if test="@style:text-emphasize">
      <w:em>
        <xsl:attribute name="w:val">
          <xsl:choose>
            <xsl:when test="@style:text-emphasize = 'accent above'">comma</xsl:when>
            <xsl:when test="@style:text-emphasize = 'dot above'">dot</xsl:when>
            <xsl:when test="@style:text-emphasize = 'circle above'">circle</xsl:when>
            <xsl:when test="@style:text-emphasize = 'dot below'">underDot</xsl:when>
            <xsl:otherwise>none</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </w:em>
    </xsl:if>

    <!-- end of toggle properties -->

    <!-- non-toggle properties -->

    <xsl:if test="$onlyToggle!='true'">

      <xsl:if test="@style:font-name">
        <w:rFonts>
          <xsl:if test="@style:font-name">
            <xsl:variable name="fontName">
              <xsl:call-template name="ComputeFontName">
                <xsl:with-param name="fontName" select="@style:font-name"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:attribute name="w:ascii">
              <xsl:value-of select="$fontName"/>
            </xsl:attribute>
            <xsl:attribute name="w:hAnsi">
              <xsl:value-of select="$fontName"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="@style:font-name-complex">
            <xsl:attribute name="w:cs">
              <xsl:call-template name="ComputeFontName">
                <xsl:with-param name="fontName" select="@style:font-name-complex"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="@style:font-name-asian">
            <xsl:attribute name="w:eastAsia">
              <xsl:call-template name="ComputeFontName">
                <xsl:with-param name="fontName" select="@style:font-name-asian"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
        </w:rFonts>
      </xsl:if>

      <xsl:if test="@fo:color">
        <w:color>
          <xsl:attribute name="w:val">
            <xsl:value-of select="substring(@fo:color, 2, string-length(@fo:color) -1)"/>
          </xsl:attribute>
        </w:color>
      </xsl:if>

      <xsl:if test="@fo:letter-spacing">
        <w:spacing>
          <xsl:attribute name="w:val">
            <xsl:choose>
              <xsl:when test="@fo:letter-spacing='normal'">
                <xsl:value-of select="0"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:call-template name="twips-measure">
                  <xsl:with-param name="length" select="@fo:letter-spacing"/>
                </xsl:call-template>
              </xsl:otherwise>
            </xsl:choose>

          </xsl:attribute>
        </w:spacing>
      </xsl:if>

      <xsl:if test="@style:text-scale">
        <w:w>
          <xsl:attribute name="w:val">
            <xsl:variable name="scale"
              select="substring(@style:text-scale, 1, string-length(@style:text-scale)-1)"/>
            <xsl:choose>
              <xsl:when test="number($scale) &lt; 600">
                <xsl:value-of select="$scale"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:message terminate="no">feedback: Text scale can't exceed 600%</xsl:message> 600
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </w:w>
      </xsl:if>

      <xsl:if test="@style:letter-kerning = 'true'">
        <w:kern w:val="16"/>
      </xsl:if>

      <xsl:if test="@fo:font-size">
        <xsl:variable name="sz">
          <xsl:call-template name="computeSize">
            <xsl:with-param name="node" select="current()"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:if test="number($sz)">
          <w:sz w:val="{$sz}"/>
        </xsl:if>
      </xsl:if>

      <xsl:if test="@fo:font-size-complex">
        <w:szCs>
          <xsl:attribute name="w:val">
            <xsl:value-of select="number(substring-before(@fo:font-size-complex, 'pt')) * 2"/>
          </xsl:attribute>
        </w:szCs>
      </xsl:if>

      <xsl:if test="@style:text-position">
        <xsl:choose>
          <xsl:when test="substring(@style:text-position, 1, 3) = 'sub'">
            <w:vertAlign w:val="subscript"/>
          </xsl:when>
          <xsl:when test="substring(@style:text-position, 1, 5) = 'super'">
            <w:vertAlign w:val="superscript"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:variable name="value">
              <xsl:value-of select="substring-before(@style:text-position, '%')"/>
            </xsl:variable>
            <xsl:if test="$value != 0">
              <xsl:choose>
                <xsl:when test="contains($value,'-')">
                  <w:vertAlign w:val="subscript"/>
                  <xsl:variable name="positionValue">
                    <xsl:value-of select="round(number(substring-after($value,'-')) div 10 * 2)"/>
                  </xsl:variable>
                  <w:position>
                    <xsl:attribute name="w:val">
                      <xsl:value-of select="concat('-',$positionValue)"/>
                    </xsl:attribute>
                  </w:position>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:variable name="positionValue">
                    <xsl:value-of select="round(number($value) div 10)"/>
                  </xsl:variable>
                  <w:vertAlign w:val="superscript"/>
                  <w:position>
                    <xsl:attribute name="w:val">
                      <xsl:value-of select="$positionValue"/>
                    </xsl:attribute>
                  </w:position>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:if>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>

      <xsl:if
        test="(@fo:language and @fo:country) or (@style:language-asian and @style:country-asian) or (@style:language-complex and @style:country-complex)">
        <w:lang>
          <xsl:if test="(@fo:language and @fo:country)">
            <xsl:attribute name="w:val">
              <xsl:value-of select="@fo:language"/>-<xsl:value-of select="@fo:country"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="(@style:language-asian and @style:country-asian)">
            <xsl:attribute name="w:eastAsia">
              <xsl:value-of select="@style:language-asian"/>-<xsl:value-of
                select="@style:country-asian"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="(@style:language-complex and @style:country-complex)">
            <xsl:attribute name="w:bidi">
              <xsl:value-of select="@style:language-complex"/>-<xsl:value-of
                select="@style:country-complex"/>
            </xsl:attribute>
          </xsl:if>
        </w:lang>
      </xsl:if>

      <xsl:if test="@style:text-combine">
        <xsl:choose>
          <xsl:when test="@style:text-combine = 'lines'">
            <w:eastAsianLayout>
              <xsl:attribute name="w:id">
                <xsl:number count="style:style"/>
              </xsl:attribute>
              <xsl:attribute name="w:combine">lines</xsl:attribute>
              <xsl:choose>
                <xsl:when
                  test="@style:text-combine-start-char = '&lt;' and @style:text-combine-end-char = '&gt;'">
                  <xsl:attribute name="w:combineBrackets">angle</xsl:attribute>
                </xsl:when>
                <xsl:when
                  test="@style:text-combine-start-char = '{' and @style:text-combine-end-char = '}'">
                  <xsl:attribute name="w:combineBrackets">curly</xsl:attribute>
                </xsl:when>
                <xsl:when
                  test="@style:text-combine-start-char = '(' and @style:text-combine-end-char = ')'">
                  <xsl:attribute name="w:combineBrackets">round</xsl:attribute>
                </xsl:when>
                <xsl:when
                  test="@style:text-combine-start-char = '[' and @style:text-combine-end-char = ']'">
                  <xsl:attribute name="w:combineBrackets">square</xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:attribute name="w:combineBrackets">none</xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
            </w:eastAsianLayout>
          </xsl:when>
        </xsl:choose>
      </xsl:if>

    </xsl:if>
    <!-- end of non-toggle properties -->
  </xsl:template>

  <!-- graphic properties -->
  <xsl:template match="style:graphic-properties[parent::style:style]" mode="pPr">
    <xsl:if test="@style:horizontal-pos">
      <w:jc>
        <xsl:attribute name="w:val">
          <xsl:choose>
            <xsl:when test="@style:horizontal-pos = 'center'">center</xsl:when>
            <xsl:when test="@style:horizontal-pos='left'">left</xsl:when>
            <xsl:when test="@style:horizontal-pos='right'">right</xsl:when>
            <xsl:otherwise>center</xsl:otherwise>
            <!--
							<value>from-left</value>
							<value>inside</value>
							<value>outside</value>
							<value>from-inside</value>
							-->
            <!-- TODO manage horizontal-rel -->
          </xsl:choose>
        </xsl:attribute>
      </w:jc>
    </xsl:if>
  </xsl:template>

  <!-- footnote and endnote reference and text styles -->
  <xsl:template
    match="text:notes-configuration[@text:note-class='footnote'] | text:notes-configuration[@text:note-class='endnote']"
    mode="styles">
    <w:style w:styleId="{concat(@text:note-class, 'Reference')}" w:type="character">
      <w:name w:val="note reference"/>
      <w:basedOn w:val="{@text:citation-body-style-name}"/>
      <w:semiHidden/>
      <w:unhideWhenUsed/>
    </w:style>
    <w:style w:styleId="{concat(@text:note-class, 'Text')}" w:type="paragraph">
      <w:name w:val="note text"/>
      <w:basedOn w:val="{@text:citation-style-name}"/>
      <!--w:link w:val="FootnoteTextChar"/-->
      <w:semiHidden/>
      <w:unhideWhenUsed/>
    </w:style>
  </xsl:template>

  <xsl:template name="computeSize">
    <xsl:param name="node"/>
    <xsl:if test="contains($node/@fo:font-size, 'pt')">
      <xsl:value-of select="round(number(substring-before($node/@fo:font-size, 'pt')) * 2)"/>
    </xsl:if>
    <xsl:if test="contains($node/@fo:font-size, '%')">
      <xsl:variable name="parentStyleName" select="$node/../@style:parent-style-name"/>
      <xsl:variable name="value">
        <xsl:call-template name="computeSize">
          <!-- should we look for @style:name in styles.xml, otherwise in content.xml ? -->
          <xsl:with-param name="node"
            select="/office:document-styles/office:styles/style:style[@style:name = $parentStyleName]/style:text-properties"
          />
        </xsl:call-template>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="number($value)">
          <xsl:value-of
            select="round(number(substring-before($node/@fo:font-size, '%')) div 100 * number($value))"
          />
        </xsl:when>
        <xsl:otherwise>
          <!-- fetch the default font size for this style family -->
          <xsl:variable name="defaultProps"
            select="/office:document-styles/office:styles/style:default-style[@style:family=$node/../@style:family]/style:text-properties"/>
          <xsl:variable name="defaultValue"
            select="number(substring-before($defaultProps/@fo:font-size, 'pt'))*2"/>
          <xsl:value-of
            select="round(number(substring-before($node/@fo:font-size, '%')) div 100 * number($defaultValue))"
          />
        </xsl:otherwise>
      </xsl:choose>

    </xsl:if>
  </xsl:template>

  <xsl:template name="tabStop">
    <xsl:param name="styleType"/>
    <w:tab>
      <xsl:attribute name="w:pos">
        <xsl:variable name="position">
          <xsl:choose>
            <xsl:when test="ancestor::office:automatic-styles">
              <xsl:variable name="name" select="ancestor::style:style/@style:parent-style-name"/>
              <xsl:call-template name="twips-measure">
                <xsl:with-param name="length"
                  select="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $name]/style:paragraph-properties/@fo:margin-left"
                />
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="twips-measure">
                <xsl:with-param name="length"
                  select="ancestor::style:paragraph-properties/@fo:margin-left"/>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="position2">
          <xsl:call-template name="twips-measure">
            <xsl:with-param name="length" select="@style:position"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:value-of select="$position+$position2"/>
      </xsl:attribute>
      <xsl:attribute name="w:val">
        <xsl:choose>
          <xsl:when test="@style:type">
            <xsl:value-of select="@style:type"/>
          </xsl:when>
          <xsl:when test="$styleType">
            <xsl:value-of select="$styleType"/>
          </xsl:when>
          <xsl:otherwise>left</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:attribute name="w:leader">
        <xsl:choose>
          <xsl:when test="@style:leader-text">
            <xsl:choose>
              <xsl:when test="@style:leader-text='.'">
                <xsl:value-of select="'dot'"/>
              </xsl:when>
              <xsl:when test="@style:leader-text='-'">
                <xsl:value-of select="'hyphen'"/>
              </xsl:when>
              <xsl:when test="@style:leader-text='_'">
                <xsl:value-of select="'heavy'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'none'"/>
              </xsl:otherwise>
            </xsl:choose>

          </xsl:when>
          <xsl:when test="@style:leader-style and not(@style:leader-text)">
            <xsl:choose>
              <xsl:when test="@style:leader-style='dotted'">
                <xsl:value-of select="'dot'"/>
              </xsl:when>
              <xsl:when test="@style:leader-style='solid'">
                <xsl:value-of select="'heavy'"/>
              </xsl:when>
              <xsl:when test="@style:leader-style='dash'">
                <xsl:value-of select="'hyphen'"/>
              </xsl:when>
              <xsl:when test="@style:leader-style='dotted'">
                <xsl:value-of select="'dot'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'none'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'none'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </w:tab>
  </xsl:template>

  <!-- Page Layout Properties -->
  <xsl:template match="style:page-layout-properties" mode="master-page">
    <xsl:choose>
      <xsl:when
        test="not(@style:print-orientation) and not(@fo:page-width) and not(@fo:page-height)">
        <xsl:apply-templates
          select="document('styles.xml')/office:document-styles/office:automatic-styles/style:page-layout[@style:name='pm1']/style:page-layout-properties"
          mode="properties"/>
      </xsl:when>
      <xsl:otherwise>
        <w:pgSz>
          <xsl:if test="@style:print-orientation != 'none' ">
            <xsl:attribute name="w:orient">
              <xsl:choose>
                <xsl:when test="@style:print-orientation = 'landscape' ">landscape</xsl:when>
                <xsl:otherwise>portrait</xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="@fo:page-width != 'none' ">
            <xsl:attribute name="w:w">
              <xsl:call-template name="twips-measure">
                <xsl:with-param name="length" select="@fo:page-width"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="@fo:page-height != 'none' ">
            <xsl:attribute name="w:h">
              <xsl:call-template name="twips-measure">
                <xsl:with-param name="length" select="@fo:page-height"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
        </w:pgSz>
      </xsl:otherwise>
    </xsl:choose>

    <xsl:choose>
      <xsl:when
        test="@fo:margin-top ='none' and @fo:margin-left='none' and @fo:margin-bottom='none'  and @fo:margin-right='none' ">
        <xsl:apply-templates
          select="document('styles.xml')/office:document-styles/office:automatic-styles/style:page-layout[@style:name='pm1']/style:page-layout-properties"
          mode="properties"/>
      </xsl:when>
      <xsl:otherwise>
        <w:pgMar>
          <xsl:if test="@fo:margin-top != 'none' ">
            <xsl:attribute name="w:top">
              <xsl:call-template name="twips-measure">
                <xsl:with-param name="length" select="@fo:margin-top"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="@fo:margin-left != 'none' ">
            <xsl:attribute name="w:left">
              <xsl:call-template name="twips-measure">
                <xsl:with-param name="length" select="@fo:margin-left"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="@fo:margin-bottom != 'none' ">
            <xsl:attribute name="w:bottom">
              <xsl:call-template name="twips-measure">
                <xsl:with-param name="length" select="@fo:margin-bottom"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="@fo:margin-right != 'none' ">
            <xsl:attribute name="w:right">
              <xsl:call-template name="twips-measure">
                <xsl:with-param name="length" select="@fo:margin-right"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
          <!-- not in odt : header and footer distance from page -->
          <xsl:if test="@fo:margin-top != 'none' ">
            <xsl:attribute name="w:header">
              <xsl:call-template name="twips-measure">
                <xsl:with-param name="length" select="@fo:margin-top"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="@fo:margin-bottom != 'none' ">
            <xsl:attribute name="w:footer">
              <xsl:call-template name="twips-measure">
                <xsl:with-param name="length" select="@fo:margin-bottom"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
          <!-- Don't exist in odt format -->
          <xsl:attribute name="w:gutter">0</xsl:attribute>
        </w:pgMar>
      </xsl:otherwise>
    </xsl:choose>

    <!-- border -->
    <xsl:if test="@fo:border and @fo:border != 'none' ">
      <w:pgBorders>
        <xsl:call-template name="InsertBorders">
          <xsl:with-param name="allSides">true</xsl:with-param>
        </xsl:call-template>
      </w:pgBorders>
    </xsl:if>
    <xsl:if test="@fo:border-top or @fo:border-left or @fo:border-bottom or @fo:border-right">
      <w:pgBorders>
        <xsl:call-template name="InsertBorders">
          <xsl:with-param name="allSides">false</xsl:with-param>
        </xsl:call-template>
      </w:pgBorders>
    </xsl:if>
    <xsl:apply-templates select="style:columns" mode="columns"/>
  </xsl:template>


  <xsl:template match="style:columns" mode="columns">
    <w:cols>
      <!-- nb columns -->
      <xsl:if test="@fo:column-count">
        <xsl:attribute name="w:num">
          <xsl:value-of select="@fo:column-count"/>
        </xsl:attribute>
      </xsl:if>
      <!-- separator -->
      <xsl:if test="style:column-sep">
        <xsl:attribute name="w:sep">
          <xsl:value-of select="1"/>
        </xsl:attribute>
      </xsl:if>

      <xsl:choose>
        <xsl:when test="@fo:column-gap">
          <xsl:attribute name="w:space">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length" select="@fo:column-gap"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="w:equalWidth">
            <xsl:value-of select="0"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>

      <!-- each column -->
      <xsl:for-each select="style:column">
        <w:col>
          <!-- the left and right spaces -->
          <xsl:variable name="start">
            <xsl:value-of select="number(substring-before(@fo:start-indent,'cm'))"/>
          </xsl:variable>
          <xsl:variable name="end">
            <xsl:value-of select="number(substring-before(@fo:end-indent,'cm'))"/>
          </xsl:variable>

          <!-- odt separate space between two columns ( col 1 : fo:end-indent and col 2 : fo:start-indent -->
          <xsl:variable name="spacesBetween">
            <xsl:choose>
              <xsl:when test="following-sibling::style:column/@fo:start-indent">
                <xsl:value-of
                  select="number(substring-before(following-sibling::style:column/@fo:start-indent,'cm')) + $end"
                />
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$end"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>

          <!-- Open xml space converted into twips -->
          <xsl:variable name="spaceTwips">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length" select="concat($spacesBetween,'cm') "/>
            </xsl:call-template>
          </xsl:variable>

          <!-- ODT spaces converted in twips-->
          <xsl:variable name="widthTwips">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length" select="concat($end+$start,'cm') "/>
            </xsl:call-template>
          </xsl:variable>

          <!-- space -->
          <xsl:attribute name="w:space">
            <xsl:value-of select="$spaceTwips"/>
          </xsl:attribute>

          <!-- width -->
          <xsl:attribute name="w:w">
            <xsl:value-of select="substring-before(@style:rel-width,'*') - $widthTwips"/>
          </xsl:attribute>
        </w:col>
      </xsl:for-each>

    </w:cols>
  </xsl:template>


  <xsl:template match="style:section-properties" mode="section">
    <xsl:apply-templates select="style:columns" mode="columns"/>
  </xsl:template>


  <!-- Insert borders, depending on allSides bool. If true, create all borders. -->
  <xsl:template name="InsertBorders">
    <xsl:param name="allSides" select="'false'"/>

    <xsl:if test="$allSides='true' or (@fo:border-top and (@fo:border-top != 'none'))">
      <w:top>
        <xsl:call-template name="border">
          <xsl:with-param name="side" select="'top'"/>
        </xsl:call-template>
      </w:top>
    </xsl:if>
    <xsl:if test="$allSides='true' or (@fo:border-left and (@fo:border-left != 'none'))">
      <w:left>
        <xsl:call-template name="border">
          <xsl:with-param name="side" select="'left'"/>
        </xsl:call-template>
      </w:left>
    </xsl:if>
    <xsl:if test="$allSides='true' or (@fo:border-bottom and (@fo:border-bottom != 'none'))">
      <w:bottom>
        <xsl:call-template name="border">
          <xsl:with-param name="side" select="'bottom'"/>
        </xsl:call-template>
      </w:bottom>
    </xsl:if>
    <xsl:if test="$allSides='true' or (@fo:border-right and (@fo:border-right != 'none'))">
      <w:right>
        <xsl:call-template name="border">
          <xsl:with-param name="side" select="'right'"/>
        </xsl:call-template>
      </w:right>
    </xsl:if>
    <xsl:if
      test="$allSides='true' and self::style:paragraph-properties and @style:join-border='false'">
      <w:between>
        <xsl:call-template name="border">
          <xsl:with-param name="side" select="'middle'"/>
        </xsl:call-template>
      </w:between>
    </xsl:if>
    <xsl:if
      test="$allSides='false' and @style:join-border='false' and (
      (@fo:border-bottom and (@fo:border-bottom != 'none'))
      or (@fo:border-top and (@fo:border-top != 'none')) )">
      <w:between>
        <xsl:call-template name="border">
          <xsl:with-param name="side" select="'middle'"/>
        </xsl:call-template>
      </w:between>
    </xsl:if>
  </xsl:template>

  <!-- Attributes of a border element. -->
  <xsl:template name="border">
    <xsl:param name="side"/>

    <xsl:variable name="borderStr">
      <xsl:choose>
        <xsl:when test="@fo:border">
          <xsl:value-of select="@fo:border"/>
        </xsl:when>
        <xsl:when test="$side='middle'">
          <xsl:choose>
            <xsl:when test="@fo:border-top != 'none'">
              <xsl:value-of select="@fo:border-top"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="@fo:border-bottom"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="attribute::node()[name()=concat('fo:border-',$side)]"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="padding">
      <xsl:choose>
        <xsl:when test="@fo:padding">
          <xsl:value-of select="@fo:padding"/>
        </xsl:when>
        <xsl:when test="$side='middle'">
          <xsl:choose>
            <xsl:when test="@fo:padding-top != 'none'">
              <xsl:value-of select="@fo:padding-top"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="@fo:padding-bottom"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="attribute::node()[name()=concat('fo:padding-',$side)]"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:attribute name="w:val">
      <xsl:call-template name="GetBorderStyle">
        <xsl:with-param name="side" select="$side"/>
        <xsl:with-param name="borderStr" select="$borderStr"/>
      </xsl:call-template>
    </xsl:attribute>

    <xsl:attribute name="w:sz">
      <xsl:call-template name="eightspoint-measure">
        <xsl:with-param name="length" select="substring-before($borderStr,  ' ')"/>
      </xsl:call-template>
    </xsl:attribute>

    <xsl:attribute name="w:space">
      <xsl:call-template name="padding-val">
        <xsl:with-param name="length" select="$padding"/>
      </xsl:call-template>
    </xsl:attribute>

    <xsl:attribute name="w:color">
      <xsl:value-of select="substring($borderStr, string-length($borderStr) -5, 6)"/>
    </xsl:attribute>

  </xsl:template>

  <!-- ignored -->
  <xsl:template match="text()" mode="styles"/>

</xsl:stylesheet>

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
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/3/main"
  exclude-result-prefixes="w">

  <xsl:import href="tables.xsl"/>

  <xsl:template name="content">
    <office:document-content>
      <office:scripts/>
      <office:font-face-decls/>
      <office:automatic-styles/>
      <office:body>
        <office:text>
          <xsl:apply-templates select="document('word/document.xml')/w:document/w:body"/>
        </office:text>
      </office:body>
    </office:document-content>
  </xsl:template>

  <!--  Check if the paragraf is heading -->
  <xsl:template name="GetOutlineLevel">
    <xsl:choose>

      <xsl:when test="../w:pPr/w:outlineLvl/@w:val">
        <xsl:value-of select="../w:pPr/w:outlineLvl/@w:val"/>
      </xsl:when>

      <xsl:when test="../w:pPr/w:pStyle/@w:val">

        <xsl:variable name="outline">
          <xsl:value-of select="../w:pPr/w:pStyle/@w:val"/>
        </xsl:variable>

        <xsl:value-of
          select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$outline]/w:pPr/w:outlineLvl/@w:val"
        />
      </xsl:when>

      <xsl:otherwise>

        <xsl:variable name="outline">
          <xsl:value-of select="w:rPr/w:rStyle/@w:val"/>
        </xsl:variable>

        <xsl:variable name="linkedStyleOutline">
          <xsl:value-of
            select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$outline]/w:link/@w:val"
          />
        </xsl:variable>

        <xsl:value-of
          select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$linkedStyleOutline]/w:pPr/w:outlineLvl/@w:val"
        />
      </xsl:otherwise>

    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertHeading">
    <xsl:param name="outlineLevel"/>
    <text:h>
      <xsl:attribute name="text:outline-level">
        <xsl:value-of select="$outlineLevel+1"/>
      </xsl:attribute>
      <xsl:value-of select="."/>
    </text:h>
  </xsl:template>

  <xsl:template name="w:p">
    <xsl:apply-templates/>
  </xsl:template>

  <xsl:template match="w:r">
    <xsl:variable name="outlineLevel">
      <xsl:call-template name="GetOutlineLevel"/>
    </xsl:variable>
    <xsl:choose>

      <!--  check if the paragraf is heading -->
      <xsl:when test="$outlineLevel != '' ">
        <xsl:call-template name="InsertHeading">
          <xsl:with-param name="outlineLevel" select="$outlineLevel"/>
        </xsl:call-template>
      </xsl:when>

      <!--  default scenario -->
      <xsl:otherwise>
        <xsl:apply-templates/>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>


  <xsl:template match="w:t">
    <xsl:message terminate="no">progress:w:t</xsl:message>
    <text:p text:style-name="Standard">
      <!-- todo without styles -->
      <xsl:value-of select="."/>
    </text:p>
  </xsl:template>



  <!-- conversion of paragraph properties -->
  <xsl:template match="w:pPr">
    <style:paragraph-properties>

      <xsl:if test="w:keep-next">
        <xsl:attribute name="fo:keep-with-next">
          <xsl:choose>
            <xsl:when
              test="w:keep-next/@w:val='off' or w:keep-next/@w:val='false' or w:keep-next/@w:val=0">
              <xsl:value-of select="'auto'"/>
            </xsl:when>
            <xsl:when
              test="w:keep-next/@w:val='on' or w:keep-next/@w:val='true' or w:keep-next/@w:val=1">
              <xsl:value-of select="'always'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'always'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>

      <xsl:if test="w:keepLines">
        <xsl:attribute name="fo:keep-together">
          <xsl:choose>
            <xsl:when
              test="w:keepLines/@w:val='off' or w:keepLines/@w:val='false' or w:keepLines/@w:val=0">
              <xsl:value-of select="'auto'"/>
            </xsl:when>
            <xsl:when
              test="w:keepLines/@w:val='on' or w:keepLines/@w:val='true' or w:keepLines/@w:val=1">
              <xsl:value-of select="'always'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'always'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>

      <!-- break before paragraph -->
      <xsl:if test="w:pageBreakBefore">
        <xsl:attribute name="fo:break-before">
          <xsl:choose>
            <xsl:when
              test="w:pageBreakBefore/@w:val='off' or w:pageBreakBefore/@w:val='false' or w:pageBreakBefore/@w:val=0">
              <xsl:value-of select="'auto'"/>
            </xsl:when>
            <xsl:when
              test="w:pageBreakBefore/@w:val='on' or w:pageBreakBefore/@w:val='true' or w:pageBreakBefore/@w:val=1">
              <xsl:value-of select="'page'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'page'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>

      <!-- widows and orphans -->
      <xsl:if test="w:widowControl">
        <xsl:choose>
          <xsl:when
            test="w:widowControl/@w:val='on' or w:widowControl/@w:val='true' or w:widowControl/@w:val=1 or not(w:widowControl/@w:val)">
            <xsl:attribute name="fo:widows">
              <xsl:value-of select="1"/>
            </xsl:attribute>
            <xsl:attribute name="fo:orphans">
              <xsl:value-of select="1"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise/>
        </xsl:choose>
      </xsl:if>

      <!-- borders. -->
      <xsl:if test="w:pBdr">
        <!-- TODO : if the four borders are equal, then create adequate attributes. -->
        <xsl:apply-templates select="w:pBdr/child::node()"/>
        <!-- add shadow -->
        <xsl:variable name="firstVal">
          <xsl:choose>
            <xsl:when
              test="w:pBdr/w:right/@w:shadow='true' or w:pBdr/w:right/@w:shadow=1 or w:pBdr/w:right/@w:shadow='on'">
              <xsl:call-template name="ConvertPoints">
                <xsl:with-param name="length">
                  <xsl:value-of select="w:pBdr/w:right/@w:space"/>
                </xsl:with-param>
                <xsl:with-param name="unit">cm</xsl:with-param>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when
                  test="w:pBdr/w:left/@w:shadow='true' or w:pBdr/w:left/@w:shadow=1 or w:pBdr/w:left/@w:shadow='on'">
                  <xsl:call-template name="ConvertPoints">
                    <xsl:with-param name="length">
                      <xsl:value-of select="w:pBdr/w:left/@w:space"/>
                    </xsl:with-param>
                    <xsl:with-param name="unit">cm</xsl:with-param>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="secondVal">
          <xsl:choose>
            <xsl:when
              test="w:pBdr/w:bottom/@w:shadow='true' or w:pBdr/w:bottom/@w:shadow=1 or w:pBdr/w:bottom/@w:shadow='on'">
              <xsl:call-template name="ConvertPoints">
                <xsl:with-param name="length">
                  <xsl:value-of select="w:pBdr/w:bottom/@w:space"/>
                </xsl:with-param>
                <xsl:with-param name="unit">cm</xsl:with-param>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when
                  test="w:pBdr/w:top/@w:shadow='true' or w:pBdr/w:top/@w:shadow=1 or w:pBdr/w:top/@w:shadow='on'">
                  <xsl:call-template name="ConvertPoints">
                    <xsl:with-param name="length">
                      <xsl:value-of select="w:pBdr/w:top/@w:space"/>
                    </xsl:with-param>
                    <xsl:with-param name="unit">cm</xsl:with-param>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:if test="$firstVal !=0 and $secondVal != 0">
          <xsl:attribute name="style:shadow">
            <xsl:value-of select="concat('#000000 ',$firstVal,' ',$secondVal)"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:if>

      <!-- bg color -->
      <xsl:if test="w:shd">
        <xsl:attribute name="fo:background-color">
          <xsl:choose>
            <xsl:when test="w:shd/@w:fill='auto' or not(w:shd/@w:fill)">
              <xsl:value-of select="'transparent'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="concat('#',w:shd/@w:fill)"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>

      <!-- text autospace -->
      <xsl:if test="w:autoSpaceDN or w:autoSpaceDE">
        <xsl:attribute name="style:text-autospace">
          <xsl:choose>
            <xsl:when
              test="w:autoSpaceDN/@w:val='off' or w:autoSpaceDN/@w:val='false' or w:autoSpaceDN/@w:val=0 or w:autoSpaceDE/@w:val='off' or w:autoSpaceDE/@w:val='false' or w:autoSpaceDE/@w:val=0">
              <xsl:value-of select="'none'"/>
            </xsl:when>
            <xsl:when
              test="w:autoSpaceDN/@w:val='on' or w:autoSpaceDN/@w:val='true' or w:autoSpaceDN/@w:val=1 or w:autoSpaceDE/@w:val='on' or w:autoSpaceDE/@w:val='true' or w:autoSpaceDE/@w:val=1">
              <xsl:value-of select="'ideograph-alpha'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'ideograph-alpha'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>

      <!-- space before/after -->
      <!-- TODO : Check how to use the w:spacing/w:afterLines or w:beforeLines elements -->
      <xsl:if test="w:spacing/@w:before">
        <xsl:attribute name="fo:margin-top">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:value-of select="w:spacing/@w:before"/>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="w:spacing/@w:after">
        <xsl:attribute name="fo:margin-bottom">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:value-of select="w:spacing/@w:after"/>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>

      <!-- line spacing -->
      <xsl:if test="w:spacing/@w:line">
        <xsl:choose>
          <xsl:when test="w:spacing/@w:lineRule='atLeast'">
            <xsl:attribute name="style:line-height-at-least">
              <!-- convert 20th pt to centimeter -->
              <xsl:value-of select="concat(w:spacing/@w:line * 2.54 div (1440 * 20),'cm')"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="w:spacing/@w:lineRule='exact'">
            <xsl:attribute name="style:line-height">
              <!-- convert 20th pt to centimeter -->
              <xsl:value-of select="concat(w:spacing/@w:line * 2.54 div (1440 * 20),'cm')"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <!-- value of lineRule is 'auto' -->
            <xsl:attribute name="style:line-height">
              <!-- convert 240th of line to percent -->
              <xsl:value-of select="concat(w:spacing/@w:line div 240,'%')"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>

      <!-- text:align -->
      <xsl:if test="w:jc">
        <xsl:attribute name="fo:text-align">
          <xsl:choose>
            <xsl:when test="w:jc/@w:val='center'">
              <xsl:value-of select="'center'"/>
            </xsl:when>
            <xsl:when test="w:jc/@w:val='left'">
              <xsl:value-of select="'start'"/>
            </xsl:when>
            <xsl:when test="w:jc/@w:val='right'">
              <xsl:value-of select="'end'"/>
            </xsl:when>
            <xsl:when test="w:jc/@w:val='both' or w:jc/@w:val='distribute'">
              <xsl:value-of select="'justify'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'start'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>

      <!-- vertical alignment of text -->
      <xsl:if test="w:textAlignment">
        <xsl:attribute name="style:vertical-align">
          <xsl:choose>
            <xsl:when test="w:textAlignment/@w:val='bottom'">
              <xsl:value-of select="'bottom'"/>
            </xsl:when>
            <xsl:when test="w:textAlignment/@w:val='top'">
              <xsl:value-of select="'top'"/>
            </xsl:when>
            <xsl:when test="w:textAlignment/@w:val='center'">
              <xsl:value-of select="'middle'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'auto'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>

      <!-- hyphenation -->
      <xsl:if test="w:suppressAutoHyphens">
        <xsl:attribute name="fo:hyphenate">
          <xsl:choose>
            <xsl:when
              test="w:suppressAutoHyphens/@w:val='off' or w:suppressAutoHyphens/@w:val='false' or w:suppressAutoHyphens/@w:val=0">
              <xsl:value-of select="'true'"/>
            </xsl:when>
            <xsl:when
              test="w:suppressAutoHyphens/@w:val='on' or w:suppressAutoHyphens/@w:val='true' or w:suppressAutoHyphens/@w:val=1">
              <xsl:value-of select="'false'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'false'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>

      <!-- tabs. TODO : write a template to handle tab stops. -->
      <xsl:if test="w:tabs/w:tab">
        <style:tab-stops>
          <xsl:apply-templates select="w:tabs/w:tab"/>
        </style:tab-stops>
      </xsl:if>

    </style:paragraph-properties>

  </xsl:template>

  <!-- compute attributes defining borders : color, style, width, padding... -->
  <xsl:template match="node()[parent::w:pBdr]">
    <xsl:choose>
      <xsl:when test="name()='w:between'">
        <xsl:attribute name="style:join-border">false</xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <!-- add padding -->
        <xsl:if test="@w:space">
          <xsl:attribute name="{concat('fo:padding-',substring-after(name(),'w:'))}">
            <xsl:call-template name="ConvertPoints">
              <xsl:with-param name="length">
                <xsl:value-of select="@w:space"/>
              </xsl:with-param>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
        <!-- add border with style, width and color -->
        <xsl:variable name="style">
          <xsl:call-template name="GetParagraphBorderStyle">
            <xsl:with-param name="style" select="@w:val"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:attribute name="{concat('fo:border-',substring-after(name(),'w:'))}">
          <xsl:variable name="size">
            <xsl:call-template name="ConvertEighthsPoints">
              <xsl:with-param name="length">
                <xsl:value-of select="@w:sz"/>
              </xsl:with-param>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="color">
            <xsl:choose>
              <xsl:when test="@w:color='auto'">
                <!-- TODO : check if 'transparent' is the right value -->
                <xsl:value-of select="'transparent'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="concat('#',@w:color)"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:value-of select="concat($size,' ',$style,' ',$color)"/>
        </xsl:attribute>
        <xsl:if test="$style='double' ">
          <xsl:attribute name="{concat('fo:border-line-width-',substring-after(name(),'w:'))}">
            <xsl:call-template name="ComputeDoubleBorderWidth">
              <xsl:with-param name="style" select="@w:val"/>
              <xsl:with-param name="borderWidth" select="@w:sz"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- convert the style of a border -->
  <xsl:template name="GetParagraphBorderStyle">
    <xsl:param name="style"/>
    <xsl:choose>
      <xsl:when test="$style='basicBlackDashes'">dashed</xsl:when>
      <xsl:when test="$style='basicBlackDots'">dotted</xsl:when>
      <xsl:when test="$style='basicThinLines'">double</xsl:when>
      <xsl:when test="$style='basicWideInLine'">double</xsl:when>
      <xsl:when test="$style='basicWideOutLine'">double</xsl:when>
      <xsl:when test="$style='dashed'">dashed</xsl:when>
      <xsl:when test="$style='dashSmallGap'">dashed</xsl:when>
      <xsl:when test="$style='dotted'">dotted</xsl:when>
      <xsl:when test="$style='double'">double</xsl:when>
      <xsl:when test="$style='inset'">inset</xsl:when>
      <xsl:when test="$style='nil'">hidden</xsl:when>
      <xsl:when test="$style='outset'">outset</xsl:when>
      <xsl:when test="$style='single'">solid</xsl:when>
      <xsl:when test="$style='thick'">solid</xsl:when>
      <xsl:when test="$style='thickThinSmallGap'">double</xsl:when>
      <xsl:when test="$style='thickThinMediumGap'">double</xsl:when>
      <xsl:when test="$style='thickThinLargeGap'">double</xsl:when>
      <xsl:when test="$style='thinThickSmallGap'">double</xsl:when>
      <xsl:when test="$style='thinThickMediumGap'">double</xsl:when>
      <xsl:when test="$style='thinThickLargeGap'">double</xsl:when>
      <xsl:otherwise>
        <!-- empty string -->
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- return a three arguments string defining the widths of a 'double' border -->
  <xsl:template name="ComputeDoubleBorderWidth">
    <xsl:param name="style"/>
    <xsl:param name="borderWidth"/>
    <!-- Approximate the width of each line : inner, middle (blank space), outer -->
    <xsl:variable name="inner">
      <xsl:call-template name="ConvertEighthsPoints">
        <xsl:with-param name="length">
          <xsl:choose>
            <xsl:when test="$style='basicWideInLine'">
              <xsl:value-of select="$borderWidth * 50 div 100"/>
            </xsl:when>
            <xsl:when test="$style='basicWideOutLine'">
              <xsl:value-of select="$borderWidth div 100"/>
            </xsl:when>
            <xsl:when test="$style='double' or $style='basicThinLines'">
              <xsl:value-of select="$borderWidth div 100"/>
            </xsl:when>
            <xsl:when test="$style='thickThinSmallGap'">
              <xsl:value-of select="$borderWidth * 67 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thickThinMediumGap'">
              <xsl:value-of select="$borderWidth * 47 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thickThinLargeGap'">
              <xsl:value-of select="$borderWidth * 27 div 100"/>
            </xsl:when>
            <xsl:when test="contains($style,'thinThick')">
              <xsl:value-of select="$borderWidth * 3 div 100"/>
            </xsl:when>
            <xsl:otherwise/>
          </xsl:choose>
        </xsl:with-param>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="middle">
      <xsl:call-template name="ConvertEighthsPoints">
        <xsl:with-param name="length">
          <xsl:choose>
            <xsl:when test="$style='basicWideInLine' or $style='basicWideOutLine'">
              <xsl:value-of select="$borderWidth * 49 div 100"/>
            </xsl:when>
            <xsl:when test="$style='double' or $style='basicThinLines'">
              <xsl:value-of select="$borderWidth * 98 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thickThinSmallGap'">
              <xsl:value-of select="$borderWidth * 30 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thickThinMediumGap'">
              <xsl:value-of select="$borderWidth * 50 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thickThinLargeGap'">
              <xsl:value-of select="$borderWidth * 70 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thinThickSmallGap'">
              <xsl:value-of select="$borderWidth * 30 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thinThickMediumGap'">
              <xsl:value-of select="$borderWidth * 50 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thinThickLargeGap'">
              <xsl:value-of select="$borderWidth * 70 div 100"/>
            </xsl:when>
            <xsl:otherwise/>
          </xsl:choose>
        </xsl:with-param>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="outer">
      <xsl:call-template name="ConvertEighthsPoints">
        <xsl:with-param name="length">
          <xsl:choose>
            <xsl:when test="$style='basicWideInLine'">
              <xsl:value-of select="$borderWidth div 100"/>
            </xsl:when>
            <xsl:when test="$style='basicWideOutLine'">
              <xsl:value-of select="$borderWidth * 50 div 100"/>
            </xsl:when>
            <xsl:when test="$style='double' or $style='basicThinLines'">
              <xsl:value-of select="$borderWidth div 100"/>
            </xsl:when>
            <xsl:when test="contains($style,'thickThin')">
              <xsl:value-of select="$borderWidth * 3 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thinThickSmallGap'">
              <xsl:value-of select="$borderWidth * 67 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thinThickMediumGap'">
              <xsl:value-of select="$borderWidth * 47 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thinThickLargeGap'">
              <xsl:value-of select="$borderWidth * 27 div 100"/>
            </xsl:when>
            <xsl:otherwise/>
          </xsl:choose>
        </xsl:with-param>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:value-of select="concat($inner,' ',$middle,' ',$outer)"/>
  </xsl:template>

</xsl:stylesheet>

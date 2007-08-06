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
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  exclude-result-prefixes="style svg fo">
  
  <!--
		Calculate a padding measure (limited to 31 pt)
	-->
  <xsl:template name="padding-val">
    <xsl:param name="length"/>
    <xsl:variable name="result">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length" select="$length"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$result > 31">
        <xsl:message terminate="no">translation.odf2oox.paddingShortened</xsl:message>
        <xsl:text>31</xsl:text>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$result"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="indent-val">
    <xsl:param name="length"/>
    <xsl:variable name="result">
      <xsl:call-template name="twips-measure">
        <xsl:with-param name="length" select="$length"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$result > 620">
        <xsl:message terminate="no">translation.odf2oox.paddingShortened</xsl:message>
        <xsl:text>620</xsl:text>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$result"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <!-- check if border is not too large -->
  <xsl:template name="CheckBorder">
    <xsl:param name="length"/>
    <xsl:param name="unit"/>
    <xsl:choose>
      <xsl:when test="$unit = 'twips' and $length &gt; 240">
        <xsl:message terminate="no">translation.odf2oox.borderShortened</xsl:message>
        <xsl:text>240</xsl:text>
      </xsl:when>
      <xsl:when test="$unit = 'eightspoint' and $length &gt; 96">
        <xsl:message terminate="no">translation.odf2oox.borderShortened</xsl:message>
        <xsl:text>96</xsl:text>
      </xsl:when>
      <xsl:when test="$unit = 'emu' and $length &gt; 152400">
        <xsl:message terminate="no">translation.odf2oox.borderShortened</xsl:message>
        <xsl:text>152400</xsl:text>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$length"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--
		Convert RGB code (#xxxxxx) to string-type color.
	-->
  <xsl:template name="StringType-color">
    <xsl:param name="RGBcolor"/>
    <xsl:variable name="code">
      <xsl:choose>
        <xsl:when test="substring($RGBcolor, 1,1) = '#'">
          <xsl:value-of
            select="translate(translate(substring($RGBcolor, 2, string-length($RGBcolor)-1),'f','F'),'c','C')"
          />
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="translate(translate($RGBcolor,'f','F'),'c','C')"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$code='000000'">black</xsl:when>
      <xsl:when test="$code='0000FF'">blue</xsl:when>
      <xsl:when test="$code='00FFFF'">cyan</xsl:when>
      <xsl:when test="$code='000080'">darkBlue</xsl:when>
      <xsl:when test="$code='008080'">darkCyan</xsl:when>
      <xsl:when test="$code='808080'">darkGray</xsl:when>
      <xsl:when test="$code='008000'">darkGreen</xsl:when>
      <xsl:when test="$code='800080'">darkMagenta</xsl:when>
      <xsl:when test="$code='800000'">darkRed</xsl:when>
      <xsl:when test="$code='808000'">darkYellow</xsl:when>
      <xsl:when test="$code='00FF00'">green</xsl:when>
      <xsl:when test="$code='C0C0C0'">lightGray</xsl:when>
      <xsl:when test="$code='FF00FF'">magenta</xsl:when>
      <xsl:when test="$code='FF0000'">red</xsl:when>
      <xsl:when test="$code='FFFFFF'">white</xsl:when>
      <xsl:when test="$code='FFFF00'">yellow</xsl:when>
      <xsl:otherwise>
        <xsl:message terminate="no">translation.odf2oox.textBgColor</xsl:message>
        <xsl:text>yellow</xsl:text>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- 
		Convert the style of a paragraph border.
	-->

  <xsl:template name="GetBorderStyle">
    <xsl:param name="side"/>
    <xsl:param name="borderStr"/>
    <xsl:param name="borderLineWidth"/>

    <xsl:choose>
      <xsl:when test="contains($borderStr, 'solid')">single</xsl:when>
      <xsl:when test="contains($borderStr, 'double')">
        <xsl:choose>
          <xsl:when test="not($borderLineWidth = '')">
            <xsl:call-template name="DoubleBorder">
              <xsl:with-param name="borderLineWidth" select="$borderLineWidth"/>
              <xsl:with-param name="side" select="$side"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>double</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="contains($borderStr, 'groove')">thinThickMediumGap</xsl:when>
      <xsl:when test="contains($borderStr, 'ridge')">thickThinMediumGap</xsl:when>
      <xsl:when test="contains($borderStr, 'dotted')">dotted</xsl:when>
      <xsl:when test="contains($borderStr, 'dashed')">dashed</xsl:when>
      <xsl:when test="contains($borderStr, 'inset')">inset</xsl:when>
      <xsl:when test="contains($borderStr, 'outset')">outset</xsl:when>
      <xsl:when test="contains($borderStr, 'hidden')">nil</xsl:when>
      <xsl:otherwise>none</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- additional template in case of 'double' style for border. -->
  <xsl:template name="DoubleBorder">
    <xsl:param name="borderLineWidth"/>
    <xsl:param name="side"/>

    <!-- Lengths converted to common measurement so as to be compared-->
    <xsl:variable name="inner">
      <xsl:call-template name="twips-measure">
        <xsl:with-param name="length" select="substring-before($borderLineWidth,' ')"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="between">
      <xsl:call-template name="twips-measure">
        <xsl:with-param name="length"
          select="substring-before(substring-after($borderLineWidth,' '),' ')"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="outer">
      <xsl:call-template name="twips-measure">
        <xsl:with-param name="length"
          select="substring-after(substring-after($borderLineWidth,' '),' ')"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$side='top' or $side='left'">
        <xsl:choose>
          <xsl:when test="$inner &gt; $outer">
            <xsl:choose>
              <xsl:when test="$inner &gt; $between">thickThinSmallGap</xsl:when>
              <xsl:when test="$inner &lt; $between">thickThinLargeGap</xsl:when>
              <xsl:otherwise>thickThinMediumGap</xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:when test="$inner &lt; $outer">
            <xsl:choose>
              <xsl:when test="$outer &gt; $between">thinThickSmallGap</xsl:when>
              <xsl:when test="$outer &lt; $between">thinThickLargeGap</xsl:when>
              <xsl:otherwise>thinThickMediumGap</xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>double</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="$side='bottom' or $side='right'">
        <xsl:choose>
          <xsl:when test="$inner &gt; $outer">
            <xsl:choose>
              <xsl:when test="$inner &gt; $between">thinThickSmallGap</xsl:when>
              <xsl:when test="$inner &lt; $between">thinThickLargeGap</xsl:when>
              <xsl:otherwise>thinThickMediumGap</xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:when test="$inner &lt; $outer">
            <xsl:choose>
              <xsl:when test="$outer &gt; $between">thickThinSmallGap</xsl:when>
              <xsl:when test="$outer &lt; $between">thickThinLargeGap</xsl:when>
              <xsl:otherwise>thickThinMediumGap</xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>double</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <!-- side='middle' line. Only thinThickThin border in OOX. -->
        <xsl:choose>
          <xsl:when test="$inner &lt; $outer">
            <xsl:choose>
              <xsl:when test="$outer &gt; $between">thinThickThinSmallGap</xsl:when>
              <xsl:when test="$outer &lt; $between">thinThickThinLargeGap</xsl:when>
              <xsl:otherwise>thinThickThinMediumGap</xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>triple</xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Get the number formatting switch -->
  <xsl:template name="GetNumberFormattingSwitch">
    <xsl:choose>
      <xsl:when test="@style:num-format = 'i' ">\* roman</xsl:when>
      <xsl:when test="@style:num-format = 'I' ">\* Roman</xsl:when>
      <xsl:when test="@style:num-format = 'a' ">\* alphabetic</xsl:when>
      <xsl:when test="@style:num-format = 'A' ">\* ALPHABETIC</xsl:when>
      <xsl:otherwise>\* arabic</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- 
		Generate a decimal identifier based on the position of the current 
		footenote/endnote among all the indexed footnotes/endnotes.
	-->
  <xsl:template name="GenerateId">
    <xsl:param name="node"/>
    <xsl:param name="nodetype"/>
    <xsl:variable name="positionInGroup">
      <xsl:for-each select="key(concat($nodetype,'s'), '')">
        <xsl:if test="generate-id($node) = generate-id(.)">
          <xsl:value-of select="position() + 1"/>
        </xsl:if>
      </xsl:for-each>
    </xsl:variable>
    <xsl:value-of select="$positionInGroup"/>
  </xsl:template>


</xsl:stylesheet>

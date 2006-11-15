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
		U n i t s
		
		1 pt = 20 twip
		1 in = 72 pt = 1440 twip
		1 cm = 1440 / 2.54 twip
		1 pica = 12 pt
		1 dpt (didot point) = 1/72 in (almost the same as 1 pt)
		1 px = 0.0264cm at 96dpi (Windows default)
		1 milimeter(mm) = 0.1cm
  -->

  <!-- Gets the unit of a length -->
  <xsl:template name="GetUnit">
    <xsl:param name="length"/>
    <xsl:choose>
      <xsl:when test="contains($length, 'cm')">cm</xsl:when>
      <xsl:when test="contains($length, 'mm')">mm</xsl:when>
      <xsl:when test="contains($length, 'in')">in</xsl:when>
      <xsl:when test="contains($length, 'pt')">pt</xsl:when>
      <xsl:when test="contains($length, 'twip')">twip</xsl:when>
      <xsl:when test="contains($length, 'pica')">pica </xsl:when>
      <xsl:when test="contains($length, 'dpt')">dpt</xsl:when>
      <xsl:when test="contains($length, 'px')">px</xsl:when>
    </xsl:choose>
  </xsl:template>


  <!-- Gets the unit-less value of a measured length -->
  <xsl:template name="GetValue">
    <xsl:param name="length"/>
    <xsl:choose>
      <xsl:when test="contains($length, 'cm')">
        <xsl:value-of select="substring-before($length,'cm')"/>
      </xsl:when>
      <xsl:when test="contains($length, 'mm')">
        <xsl:value-of select="substring-before($length,'mm')"/>
      </xsl:when>
      <xsl:when test="contains($length, 'in')">
        <xsl:value-of select="substring-before($length,'in')"/>
      </xsl:when>
      <xsl:when test="contains($length, 'pt')">
        <xsl:value-of select="substring-before($length,'pt')"/>
      </xsl:when>
      <xsl:when test="contains($length, 'twip')">
        <xsl:value-of select="substring-before($length,'twip')"/>
      </xsl:when>
      <xsl:when test="contains($length, 'pica')">
        <xsl:value-of select="substring-before($length,'pica')"/>
      </xsl:when>
      <xsl:when test="contains($length, 'dpt')">
        <xsl:value-of select="substring-before($length,'dpt')"/>
      </xsl:when>
      <xsl:when test="contains($length, 'px')">
        <xsl:value-of select="substring-before($length,'px')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$length"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Convert a measure with a given unit -->
  <xsl:template name="ConvertMeasure">
    <xsl:param name="length"/>
    <xsl:param name="unit"/>
    <xsl:choose>
      <xsl:when test="$unit = 'twips' ">
        <xsl:call-template name="twips-measure">
          <xsl:with-param name="length" select="$length"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$unit = 'milimeter' ">
        <xsl:call-template name="milimeter-measure">
          <xsl:with-param name="length" select="$length"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$unit = 'eightspoint' ">
        <xsl:call-template name="eightspoint-measure">
          <xsl:with-param name="length" select="$length"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$unit = 'emu' ">
        <xsl:call-template name="emu-measure">
          <xsl:with-param name="length" select="$length"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$unit = 'point' ">
        <xsl:call-template name="point-measure">
          <xsl:with-param name="length" select="$length"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>0</xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <!-- 
		Convert various length units to twips (twentieths of a point)
	-->
  <xsl:template name="twips-measure">
    <xsl:param name="length"/>
    <xsl:choose>
      <xsl:when test="contains($length, 'cm')">
        <xsl:value-of select="round(number(substring-before($length, 'cm')) * 1440 div 2.54)"/>
      </xsl:when>
      <xsl:when test="contains($length, 'mm')">
        <xsl:value-of select="round(number(substring-before($length, 'mm')) * 1440 div 25.4)"/>
      </xsl:when>
      <xsl:when test="contains($length, 'in')">
        <xsl:value-of select="round(number(substring-before($length, 'in')) * 1440)"/>
      </xsl:when>
      <xsl:when test="contains($length, 'pt')">
        <xsl:value-of select="round(number(substring-before($length, 'pt')) * 20)"/>
      </xsl:when>
      <xsl:when test="contains($length, 'twip')">
        <xsl:value-of select="substring-before($length, 'twip')"/>
      </xsl:when>
      <xsl:when test="contains($length, 'pica')">
        <xsl:value-of select="round(number(substring-before($length, 'pica')) * 240)"/>
      </xsl:when>
      <xsl:when test="contains($length, 'dpt')">
        <xsl:value-of select="round(number(substring-before($length, 'dpt')) * 20)"/>
      </xsl:when>
      <xsl:when test="contains($length, 'px')">
        <xsl:value-of select="round(number(substring-before($length, 'px')) * 1440 div 96.19)"/>
      </xsl:when>
      <xsl:when test="not($length) or $length='' ">0</xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$length"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!-- 
		Convert various length units to milimeters
	-->
  <xsl:template name="milimeter-measure">
    <xsl:param name="length"/>
    <xsl:choose>
      <xsl:when test="contains($length, 'cm')">
        <xsl:value-of select="number(substring-before($length, 'cm')) * 10"/>
      </xsl:when>
      <xsl:when test="contains($length, 'mm')">
        <xsl:value-of select="number(substring-before($length, 'mm'))"/>
      </xsl:when>
      <xsl:when test="contains($length, 'in')">
        <xsl:value-of select="number(substring-before($length, 'in')) * 25.4"/>
      </xsl:when>
      <xsl:when test="contains($length, 'pt')">
        <xsl:value-of select="number(substring-before($length, 'pt')) * 25.4 div 72"/>
      </xsl:when>
      <xsl:when test="contains($length, 'twip')">
        <xsl:value-of select="number(substring-before($length, 'twip')) * 25.4 div 1440"/>
      </xsl:when>
      <xsl:when test="contains($length, 'pica')">
        <xsl:value-of select="number(substring-before($length, 'pica')) * 25.4 div 6"/>
      </xsl:when>
      <xsl:when test="contains($length, 'dpt')">
        <xsl:value-of select="number(substring-before($length, 'pt')) * 25.4 div 72"/>
      </xsl:when>
      <xsl:when test="contains($length, 'px')">
        <xsl:value-of select="number(substring-before($length, 'px')) * 0.264"/>
      </xsl:when>
      <xsl:when test="not($length) or $length='' ">0</xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$length"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!-- 
		Convert  length units to eights of a point
	-->
  <xsl:template name="eightspoint-measure">
    <xsl:param name="length"/>
    <xsl:choose>
      <xsl:when test="contains($length, 'cm')">
        <xsl:value-of select="round(number(substring-before($length, 'cm')) * 576 div 2.54)"/>
      </xsl:when>
      <xsl:when test="contains($length, 'mm')">
        <xsl:value-of select="round(number(substring-before($length, 'mm')) * 576 div 25.4)"/>
      </xsl:when>
      <xsl:when test="contains($length, 'in')">
        <xsl:value-of select="round(number(substring-before($length, 'in')) * 576)"/>
      </xsl:when>
      <xsl:when test="contains($length, 'pt')">
        <xsl:value-of select="round(number(substring-before($length, 'pt')) * 8)"/>
      </xsl:when>
      <xsl:when test="contains($length, 'pica')">
        <xsl:value-of select="round(number(substring-before($length, 'pica')) * 96)"/>
      </xsl:when>
      <xsl:when test="contains($length, 'dpt')">
        <xsl:value-of select="round(number(substring-before($length, 'dpt')) * 8)"/>
      </xsl:when>
      <xsl:when test="contains($length, 'px')">
        <xsl:value-of select="round(number(substring-before($length, 'px')) * 576 div 96.19)"/>
      </xsl:when>
      <xsl:when test="not($length) or $length='' ">0</xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$length"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- 
		Convert  length units to point
	-->
  <xsl:template name="point-measure">
    <xsl:param name="length"/>
    <xsl:param name="round">true</xsl:param>
    <xsl:variable name="newlength">
      <xsl:choose>
        <xsl:when test="contains($length, 'cm')">
          <xsl:value-of select="number(substring-before($length, 'cm')) * 72 div 2.54"/>
        </xsl:when>
        <xsl:when test="contains($length, 'mm')">
          <xsl:value-of select="number(substring-before($length, 'mm')) * 72 div 25.4"/>
        </xsl:when>
        <xsl:when test="contains($length, 'in')">
          <xsl:value-of select="number(substring-before($length, 'in')) * 72"/>
        </xsl:when>
        <xsl:when test="contains($length, 'pt')">
          <xsl:value-of select="number(substring-before($length, 'pt'))"/>
        </xsl:when>
        <xsl:when test="contains($length, 'pica')">
          <xsl:value-of select="number(substring-before($length, 'pica')) * 12"/>
        </xsl:when>
        <xsl:when test="contains($length, 'dpt')">
          <xsl:value-of select="number(substring-before($length, 'dpt'))"/>
        </xsl:when>
        <xsl:when test="contains($length, 'px')">
          <xsl:value-of select="number(substring-before($length, 'px')) * 72 div 96.19"/>
        </xsl:when>
        <xsl:when test="not($length) or $length='' ">0</xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$length"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$round='true'">
        <xsl:value-of select="round($newlength)"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="(round($newlength * 100)) div 100"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- 
		Convert to emu
		1cm = 360000 emu
	-->
  <xsl:template name="emu-measure">
    <xsl:param name="length"/>
    <xsl:choose>
      <xsl:when test="contains($length, 'cm')">
        <xsl:value-of select="round(number(substring-before($length, 'cm')) * 360000)"/>
      </xsl:when>
      <xsl:when test="contains($length, 'mm')">
        <xsl:value-of select="round(number(substring-before($length, 'mm')) * 36000)"/>
      </xsl:when>
      <xsl:when test="contains($length, 'in')">
        <xsl:value-of select="round(number(substring-before($length, 'in')) * 360000 * 2.54)"/>
      </xsl:when>
      <xsl:when test="contains($length, 'pt')">
        <xsl:value-of select="round(number(substring-before($length, 'pt')) * 360000 * 2.54 div 72)"
        />
      </xsl:when>
      <xsl:when test="contains($length, 'px')">
        <xsl:value-of select="round(number(substring-before($length, 'px')) * 360000 div 37.87)"/>
      </xsl:when>
      <xsl:when test="not($length) or $length='' ">0</xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$length"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

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
      <xsl:when test="$result > 31"><xsl:message terminate="no">feedback:Padding
        shortened</xsl:message>31</xsl:when>
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
      <xsl:when test="$result > 620"><xsl:message terminate="no">feedback:Padding
        shortened</xsl:message>620</xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$result"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--
    Convert a radian angle to degrees
    'select' param indicates that only a few values are allowed (90,180)
  -->
  <xsl:template name="DegreesAngle">
    <xsl:param name="angle"/>
    <xsl:param name="select" select="'false'"/>
    <xsl:param name="revert" select="'false'"/>

    <xsl:variable name="val">
      <xsl:choose>
        <xsl:when test="not($angle)">0</xsl:when>
        <xsl:when test="not($select = 'true')">
          <xsl:value-of select="round(number($angle) * 90 div 1.5707963267946)"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:choose>
            <xsl:when test="$angle = 1.5707963267946">90</xsl:when>
            <xsl:when test="$angle = -1.5707963267946">-90</xsl:when>
            <xsl:when test="$angle = 3.1415926535892">180</xsl:when>
            <xsl:when test="$angle = -3.1415926535892">180</xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$revert = 'true' ">
        <xsl:value-of select=" - number($val)"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$val"/>
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
      <xsl:otherwise><xsl:message terminate="no">feedback:Text background
      color</xsl:message>yellow</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- 
		Convert the style of a paragraph border.
	-->

  <xsl:template name="GetBorderStyle">
    <xsl:param name="side"/>
    <xsl:param name="borderStr"/>

    <xsl:variable name="borderLineWidth">
      <xsl:choose>
        <xsl:when test="@style:border-line-width">
          <xsl:value-of select="@style:border-line-width"/>
        </xsl:when>
        <xsl:when test="$side='middle'">
          <xsl:choose>
            <xsl:when test="@style:border-line-width-top">
              <xsl:value-of select="@style:border-line-width-top"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="@fo:border-line-width-bottom"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="attribute::node()[name()=concat('style:border-line-width-',$side)]"
          />
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

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

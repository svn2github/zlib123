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
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/3/main"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" exclude-result-prefixes="w">



  <!--
		U n i t s
		
		1 pt = 20 twip
		1 in = 72 pt = 1440 twip
		1 cm = 1440 / 2.54 twip
		1 pica = 12 pt
		1 dpt (didot point) = 1/72 in (almost the same as 1 pt)
		1 px = 0.0264cm at 96dpi (Windows default)
		1 milimeter(mm) = 0.1cm
               1cm = 360000 emu
  -->

  <!-- Convert a measure in twips to a 'unit' measure -->
  <xsl:template name="ConvertTwips">
    <xsl:param name="length"/>
    <xsl:param name="unit"/>
    <xsl:choose>
      <xsl:when test="$length='0' or $length=''">
        <xsl:value-of select="concat(0, $unit)"/>
      </xsl:when>
      <xsl:when test="$unit = 'cm'">
        <xsl:value-of select="concat(format-number($length * 2.54 div 1440,'#.###'),'cm')"/>
      </xsl:when>
      <xsl:when test="$unit = 'mm'">
        <xsl:value-of select="concat(format-number($length * 25.4 div 1440,'#.###'),'mm')"/>
      </xsl:when>
      <xsl:when test="$unit= 'in'">
        <xsl:value-of select="concat(format-number($length div 1440,'#.###'),'in')"/>
      </xsl:when>
      <xsl:when test="$unit = 'pt'">
        <xsl:value-of select="concat(format-number($length div 20,'#.###'),'pt')"/>
      </xsl:when>
      <xsl:when test="$unit = 'twip'">
        <xsl:value-of select="concat($length,'twip')"/>
      </xsl:when>
      <xsl:when test="$unit = 'pica'">
        <xsl:value-of select="concat(format-number($length div 240,'#.###'),'pica')"/>
      </xsl:when>
      <xsl:when test="$unit = 'dpt'">
        <xsl:value-of select="concat(format-number($length div 20,'#.###'),'dpt')"/>
      </xsl:when>
      <xsl:when test="$unit = 'px'">
        <xsl:value-of select="concat(format-number($length * 96.19 div 1440,'#.###'),'px')"/>
      </xsl:when>
      <xsl:when test="$unit='pct'">
        <xsl:value-of select="concat(format-number($length div 50, '#.###'), '%')"/>
      </xsl:when>
      <xsl:when test="not($length)">
        <xsl:value-of select="concat(0,'cm')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$length"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Convert a measure in points to a 'unit' measure -->
  <xsl:template name="ConvertPoints">
    <xsl:param name="length"/>
    <xsl:param name="unit"/>
    <xsl:variable name="lengthVal">
      <xsl:choose>
        <xsl:when test="contains($length,'pt')">
          <xsl:value-of select="substring-before($length,'pt')"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$length"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$lengthVal='0' or $lengthVal=''">
        <xsl:value-of select="concat(0, $unit)"/>
      </xsl:when>
      <xsl:when test="$unit = 'cm'">
        <xsl:value-of select="concat(format-number($lengthVal * 2.54 div 72,'#.###'),'cm')"/>
      </xsl:when>
      <xsl:when test="$unit = 'mm'">
        <xsl:value-of select="concat(format-number($lengthVal * 25.4 div 72,'#.###'),'mm')"/>
      </xsl:when>
      <xsl:when test="$unit = 'in'">
        <xsl:value-of select="concat(format-number($lengthVal div 72,'#.###'),'in')"/>
      </xsl:when>
      <xsl:when test="$unit = 'pt'">
        <xsl:value-of select="concat($lengthVal,'pt')"/>
      </xsl:when>
      <xsl:when test="$unit = 'pica'">
        <xsl:value-of select="concat(format-number($lengthVal div 12,'#.###'),'pica')"/>
      </xsl:when>
      <xsl:when test="$unit = 'dpt'">
        <xsl:value-of select="concat($lengthVal,'dpt')"/>
      </xsl:when>
      <xsl:when test="$unit = 'px'">
        <xsl:value-of select="concat(format-number($lengthVal * 96.19 div 72,'#.###'),'px')"/>
      </xsl:when>
      <xsl:when test="not($lengthVal)">
        <xsl:value-of select="concat(0,'cm')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$lengthVal"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Convert a measure in half points to a 'unit' measure -->
  <xsl:template name="ConvertHalfPoints">
    <xsl:param name="length"/>
    <xsl:param name="unit"/>
    <xsl:choose>
      <xsl:when test="$length='0' or $length=''">
        <xsl:value-of select="concat(0, $unit)"/>
      </xsl:when>
      <xsl:when test="$unit = 'cm'">
        <xsl:value-of select="concat(format-number($length * 2.54 div 144,'#.###'),'cm')"/>
      </xsl:when>
      <xsl:when test="$unit = 'mm'">
        <xsl:value-of select="concat(format-number($length * 25.4 div 144,'#.###'),'mm')"/>
      </xsl:when>
      <xsl:when test="$unit = 'in'">
        <xsl:value-of select="concat(format-number($length div 144,'#.###'),'in')"/>
      </xsl:when>
      <xsl:when test="$unit = 'pt'">
        <xsl:value-of select="concat($length div 2,'pt')"/>
      </xsl:when>
      <xsl:when test="$unit = 'pica'">
        <xsl:value-of select="concat(format-number($length div 144,'#.###'),'pica')"/>
      </xsl:when>
      <xsl:when test="$unit = 'dpt'">
        <xsl:value-of select="concat($length div 2,'dpt')"/>
      </xsl:when>
      <xsl:when test="$unit = 'px'">
        <xsl:value-of select="concat(format-number($length * 96.19 div 144,'#.###'),'px')"/>
      </xsl:when>
      <xsl:when test="not($length)">
        <xsl:value-of select="concat(0,'cm')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$length"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Convert a measure in eigths of a point to a 'unit' measure -->
  <xsl:template name="ConvertEighthsPoints">
    <xsl:param name="length"/>
    <xsl:param name="unit"/>
    <xsl:choose>
      <xsl:when test="$length='0' or $length=''">
        <xsl:value-of select="concat(0, $unit)"/>
      </xsl:when>
      <xsl:when test="$unit = 'cm'">
        <xsl:value-of select="concat(format-number($length * 2.54 div 576,'#.###'),'cm')"/>
      </xsl:when>
      <xsl:when test="$unit = 'mm'">
        <xsl:value-of select="concat(format-number($length * 25.4 div 576,'#.###'),'mm')"/>
      </xsl:when>
      <xsl:when test="$unit = 'in'">
        <xsl:value-of select="concat(format-number($length div 576,'#.###'),'in')"/>
      </xsl:when>
      <xsl:when test="$unit = 'pt'">
        <xsl:value-of select="concat(format-number($length div 8,'#.###'),'pt')"/>
      </xsl:when>
      <xsl:when test="$unit = 'pica'">
        <xsl:value-of select="concat(format-number($length div 96,'#.###'),'pica')"/>
      </xsl:when>
      <xsl:when test="$unit = 'dpt'">
        <xsl:value-of select="concat(format-number($length div 8,'#.###'),'dpt')"/>
      </xsl:when>
      <xsl:when test="$unit = 'px'">
        <xsl:value-of select="concat(format-number($length * 96.19 div 576,'#.###'),'px')"/>
      </xsl:when>
      <xsl:when test="not($length)">
        <xsl:value-of select="concat(0,'cm')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$length"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--  converts emu to given unit-->
  <xsl:template name="ConvertEmu">
    <xsl:param name="length"/>
    <xsl:param name="unit"/>
    <xsl:choose>
      <xsl:when
        test="$length = '' or not($length) or $length = 0 or format-number($length div 360000, '#.##') = ''">
        <xsl:value-of select="concat(0,'cm')"/>
      </xsl:when>
      <xsl:when test="$unit = 'cm'">
        <xsl:value-of select="concat(format-number($length div 360000, '#.##'), 'cm')"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="ConvertEmu3">
    <xsl:param name="length"/>
    <xsl:param name="unit"/>
    <xsl:choose>
      <xsl:when
        test="$length = '' or not($length) or $length = 0 or format-number($length div 360000, '#.###') = ''">
        <xsl:value-of select="concat(0,'cm')"/>
      </xsl:when>
      <xsl:when test="$unit = 'cm'">
        <xsl:value-of select="concat(format-number($length div 360000, '#.###'), 'cm')"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

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

  <xsl:template name="ConvertMeasure">
    <xsl:param name="length"/>
    <xsl:param name="sourceUnit"/>
    <xsl:param name="destUnit"/>
    <xsl:param name="addUnit">true</xsl:param>
    <xsl:choose>
      <xsl:when
        test="$length='' or $length='0' or $length='0cm' or $length='0mm' or $length='0in' or $length='0pt' or $length='0twip' or $length='0pika' or $length='0dpt' or $length='0px'">
        <xsl:value-of select="'0'"/>
      </xsl:when>
      <!-- used when unit type is given in length string-->
      <xsl:when test="$sourceUnit = ''">
        <xsl:call-template name="ConvertToMeasure">
          <xsl:with-param name="length" select="$length"/>
          <xsl:with-param name="destUnit" select="$destUnit"/>
          <xsl:with-param name="addUnit" select="$addUnit"/>
        </xsl:call-template>
      </xsl:when>
      <!-- used when unit type is not given in length string-->
      <xsl:otherwise>
        <xsl:call-template name="ConvertFromMeasure">
          <xsl:with-param name="length" select="$length"/>
          <xsl:with-param name="sourceUnit" select="$sourceUnit"/>
          <xsl:with-param name="destUnit" select="$destUnit"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--  converts from given measure - for usage when unit type is not given in string-->
  <xsl:template name="ConvertFromMeasure">
    <xsl:param name="length"/>
    <xsl:param name="destUnit"/>
    <xsl:param name="sourceUnit"/>
    <xsl:choose>
      <xsl:when test="$sourceUnit = 'eighths-points' ">
        <xsl:call-template name="ConvertEighthsPoints">
          <xsl:with-param name="length" select="$length"/>
          <xsl:with-param name="unit" select="$destUnit"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$sourceUnit = 'half-points' ">
        <xsl:call-template name="ConvertHalfPoints">
          <xsl:with-param name="length" select="$length"/>
          <xsl:with-param name="unit" select="$destUnit"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$sourceUnit = 'pt' ">
        <xsl:call-template name="ConvertPoints">
          <xsl:with-param name="length" select="$length"/>
          <xsl:with-param name="unit" select="$destUnit"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$sourceUnit = 'twip' ">
        <xsl:call-template name="ConvertTwips">
          <xsl:with-param name="length" select="$length"/>
          <xsl:with-param name="unit" select="$destUnit"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$sourceUnit = 'emu' ">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="$length"/>
          <xsl:with-param name="unit" select="$destUnit"/>
        </xsl:call-template>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!--  converts to given measure - for usage when unit type is given in string-->
  <xsl:template name="ConvertToMeasure">
    <xsl:param name="length"/>
    <xsl:param name="destUnit"/>
    <xsl:param name="addUnit">true</xsl:param>
    <xsl:choose>
      <xsl:when test="contains($destUnit, 'cm')">
        <xsl:call-template name="ConvertToCentimeters">
          <xsl:with-param name="length" select="$length"/>
          <xsl:with-param name="addUnit" select="$addUnit"/>
        </xsl:call-template>
      </xsl:when>
      <!-- TODO other units-->
    </xsl:choose>
  </xsl:template>

  <!-- converts given unit to cm -->
  <xsl:template name="ConvertToCentimeters">
    <xsl:param name="length"/>
    <xsl:param name="round">false</xsl:param>
    <xsl:param name="addUnit">true</xsl:param>
    <xsl:variable name="newlength">
      <xsl:choose>
        <xsl:when test="contains($length, 'cm')">
          <xsl:value-of select="$length"/>
        </xsl:when>
        <xsl:when test="contains($length, 'mm')">
          <xsl:value-of select="format-number(substring-before($length, 'mm') div 10,'#.###')"/>
        </xsl:when>
        <xsl:when test="contains($length, 'in')">
          <xsl:value-of select="format-number(substring-before($length, 'in') * 2.54,'#.###')"/>
        </xsl:when>
        <xsl:when test="contains($length, 'pt')">
          <xsl:value-of
            select="format-number(substring-before($length, 'pt') * 2.54 div 72,'#.###')"/>
        </xsl:when>
        <xsl:when test="contains($length, 'twip')">
          <xsl:value-of
            select="format-number(substring-before($length, 'twip') * 2.54 div 1440,'#.###')"/>
        </xsl:when>
        <xsl:when test="contains($length, 'pica')">
          <xsl:value-of
            select="format-number(substring-before($length, 'pica') * 2.54 div 6,'#.###')"/>
        </xsl:when>
        <xsl:when test="contains($length, 'dpt')">
          <xsl:value-of
            select="format-number(substring-before($length, 'pt') * 2.54 div 72,'#.###')"/>
        </xsl:when>
        <xsl:when test="contains($length, 'px')">
          <xsl:value-of select="format-number(substring-before($length, 'px') * 0.0264,'#.###')"/>
        </xsl:when>
        <xsl:when test="not($length) or $length='' ">0</xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="format-number($length * 2.54 div 1440,'#.###')"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="roundLength">
      <xsl:choose>
        <xsl:when test="$round='true'">
          <xsl:value-of select="round($newlength)"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="(round($newlength * 1000)) div 1000"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$addUnit = 'true' ">
        <xsl:value-of select="concat($roundLength, 'cm')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$roundLength"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertColor">
    <xsl:param name="color"/>
    <xsl:choose>
      <xsl:when test="contains($color,'#')">
        <xsl:choose>
          <xsl:when test="contains($color,' ')">
            <xsl:value-of select="substring-before($color,' ')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$color"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <!--TODO standard colors mapping (there are 10 standard colors in Word)-->
        <xsl:choose>
          <xsl:when test="$color = 'aqua'">
            <xsl:text>#00ffff</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'black' or contains($color,'black')">
            <xsl:text>#000000</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'blue'">
            <xsl:text>#000080</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'fuchsia'">
            <xsl:text>#ff00ff</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'gray'">
            <xsl:text>#808080</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'green'">
            <xsl:text>#008000</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'lime'">
            <xsl:text>#00ff00</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'maroon'">
            <xsl:text>#800000</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'navy'">
            <xsl:text>#000080</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'olive'">
            <xsl:text>#808000</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'purple'">
            <xsl:text>#800080</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'red'">
            <xsl:text>#ff0000</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'silver'">
            <xsl:text>#c0c0c0</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'teal'">
            <xsl:text>#008080</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'white' or contains($color,'white')">
            <xsl:text>#ffffff</xsl:text>
          </xsl:when>
          <xsl:when test="$color='yellow'">
            <xsl:text>#ffff00</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>#000000</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--  hex to decimal -->
  <xsl:template name="HexToDec">
    <xsl:param name="number"/>
    <xsl:param name="step" select="0"/>
    <xsl:param name="value" select="0"/>
    <xsl:variable name="number1">
      <xsl:value-of select="translate($number,'ABCDEF','abcdef')"/>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="string-length($number1) &gt; 0">
        <xsl:variable name="one">
          <xsl:choose>
            <xsl:when test="substring($number1,string-length($number1) ) = 'a'">
              <xsl:text>10</xsl:text>
            </xsl:when>
            <xsl:when test="substring($number1,string-length($number1)) = 'b'">
              <xsl:text>11</xsl:text>
            </xsl:when>
            <xsl:when test="substring($number1,string-length($number1)) = 'c'">
              <xsl:text>12</xsl:text>
            </xsl:when>
            <xsl:when test="substring($number1,string-length($number1)) = 'd'">
              <xsl:text>13</xsl:text>
            </xsl:when>
            <xsl:when test="substring($number1,string-length($number1)) = 'e'">
              <xsl:text>14</xsl:text>
            </xsl:when>
            <xsl:when test="substring($number1,string-length($number1)) = 'f'">
              <xsl:text>15</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring($number1,string-length($number1))"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="power">
          <xsl:call-template name="Power">
            <xsl:with-param name="base">16</xsl:with-param>
            <xsl:with-param name="exponent">
              <xsl:value-of select="number($step)"/>
            </xsl:with-param>
            <xsl:with-param name="value1">16</xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="string-length($number1) = 1">
            <xsl:value-of select="($one * $power )+ number($value)"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="HexToDec">
              <xsl:with-param name="number">
                <xsl:value-of select="substring($number1,1,string-length($number1) - 1)"/>
              </xsl:with-param>
              <xsl:with-param name="step">
                <xsl:value-of select="number($step) + 1"/>
              </xsl:with-param>
              <xsl:with-param name="value">
                <xsl:value-of select="($one * $power) + number($value)"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="Power">
    <xsl:param name="base"/>
    <xsl:param name="exponent"/>
    <xsl:param name="value1"/>
    <xsl:choose>
      <xsl:when test="$exponent = 0">
        <xsl:text>1</xsl:text>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$exponent &gt; 1">
            <xsl:call-template name="Power">
              <xsl:with-param name="base">
                <xsl:value-of select="$base"/>
              </xsl:with-param>
              <xsl:with-param name="exponent">
                <xsl:value-of select="$exponent -1"/>
              </xsl:with-param>
              <xsl:with-param name="value1">
                <xsl:value-of select="$value1 * $base"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$value1"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

</xsl:stylesheet>

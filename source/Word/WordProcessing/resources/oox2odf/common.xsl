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
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">


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
          <xsl:when test="$color = 'aqua' or contains($color,'aqua')">
            <xsl:text>#00ffff</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'black' or contains($color,'black')">
            <xsl:text>#000000</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'blue' or contains($color,'blue')">
            <xsl:text>#0000ff</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'fuchsia' or contains($color,'fuchsia')">
            <xsl:text>#ff00ff</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'gray' or contains($color,'gray')">
            <xsl:text>#808080</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'green' or contains($color,'green')">
            <xsl:text>#008000</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'lime' or contains($color,'lime')">
            <xsl:text>#00ff00</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'maroon' or contains($color,'maroon')">
            <xsl:text>#800000</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'navy' or contains($color,'navy')">
            <xsl:text>#000080</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'olive' or contains($color,'olive')">
            <xsl:text>#808000</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'purple' or contains($color,'purple')">
            <xsl:text>#800080</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'red' or contains($color,'red')">
            <xsl:text>#ff0000</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'silver' or contains($color,'silver')">
            <xsl:text>#c0c0c0</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'teal' or contains($color,'teal')">
            <xsl:text>#008080</xsl:text>
          </xsl:when>
          <xsl:when test="$color = 'white' or contains($color,'white')">
            <xsl:text>#ffffff</xsl:text>
          </xsl:when>
          <xsl:when test="$color='yellow' or contains($color,'yellow')">
            <xsl:text>#ffff00</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>#000000</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>


</xsl:stylesheet>

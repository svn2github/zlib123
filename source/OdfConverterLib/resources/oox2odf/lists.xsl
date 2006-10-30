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
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  exclude-result-prefixes="w">
  
  <xsl:key name="numId" match="w:num" use="@w:numId"/>
  <xsl:key name="abstractNumId" match="w:abstractNum" use="@w:abstractNumId"/>
  
  <!--insert num template for each text-list style -->
  
  <xsl:template match="w:num">
    <xsl:variable name="id">
      <xsl:value-of select="@w:numId"/>
    </xsl:variable>
    
    <!-- apply abstractNum template with the same id -->
    <xsl:apply-templates select="key('abstractNumId',w:abstractNumId/@w:val)">
      <xsl:with-param name="id">
        <xsl:value-of  select="$id"/>
      </xsl:with-param>
    </xsl:apply-templates>
  </xsl:template>
  
  <!-- insert abstractNum template -->
  
  <xsl:template match="w:abstractNum">
    <xsl:param name="id"/>
    <text:list-style style:name="{concat('L',$id)}">
      <xsl:for-each select="w:lvl">
        <xsl:variable name="level" select="@w:ilvl"/>
        <xsl:choose>
          
          <!-- when numbering style is overriden, num template is used -->
          <xsl:when test="key('numId',$id)/w:lvlOverride[@w:ilvl = $level]">
            <xsl:apply-templates select="key('numId',$id)/w:lvlOverride[@w:ilvl = $level]/w:lvl[@w:ilvl = $level]"/>
          </xsl:when>
          
          <xsl:otherwise>
          <xsl:apply-templates select="."/>
          </xsl:otherwise>
        </xsl:choose>
        </xsl:for-each>
    </text:list-style>
  </xsl:template>
  
  <xsl:template match="w:lvl">
    <xsl:choose>
      
      <!--check if it's numbering or bullet -->
      <xsl:when test="w:numFmt[@w:val = 'bullet']">
        <text:list-level-style-bullet text:level="{number(@w:ilvl)+1}" text:style-name="Bullet_20_Symbols">
          <xsl:attribute name="text:bullet-char">
            <xsl:call-template name="TextChar"/>
          </xsl:attribute>
          <style:list-level-properties>
            <xsl:call-template name="ListLevelProperties"/>
          </style:list-level-properties>
          <style:text-properties style:font-name="StarSymbol"/>
        </text:list-level-style-bullet>
      </xsl:when>
      <xsl:otherwise>
        <text:list-level-style-number text:level="{number(@w:ilvl)+1}">
          <xsl:if test="not(number(substring(w:lvlText/@w:val,string-length(w:lvlText/@w:val))))">
            <xsl:attribute name="style:num-suffix">
              <xsl:value-of select="substring(w:lvlText/@w:val,string-length(w:lvlText/@w:val))"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:attribute name="style:num-format">
            <xsl:call-template name="NumFormat">
              <xsl:with-param name="format" select="w:numFmt/@w:val"/>
            </xsl:call-template>
          </xsl:attribute>
          <style:list-level-properties>
            <xsl:call-template name="ListLevelProperties"/>
          </style:list-level-properties>
        </text:list-level-style-number>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- numbering format -->
  
  <xsl:template name="NumFormat">
    <xsl:param name="format"/>
    <xsl:choose>
      <xsl:when test="$format= 'decimal' ">1</xsl:when>
      <xsl:when test="$format= 'lowerRoman' ">i</xsl:when>
      <xsl:when test="$format= 'upperRoman' ">I</xsl:when>
      <xsl:when test="$format= 'lowerLetter' ">a</xsl:when>
      <xsl:when test="$format= 'upperLetter' ">A</xsl:when>
      <xsl:otherwise>1</xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- properties for each list level -->
  
  <xsl:template name="ListLevelProperties">
    <xsl:variable name="Ind" select="w:pPr/w:ind"/> 
    <xsl:variable name="tab" select="w:pPr/w:tabs/w:tab"/>
    <xsl:choose>
      <xsl:when test="$Ind/@w:hanging">
        <xsl:attribute name="text:space-before">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
                  <xsl:value-of select="$Ind/@w:left - $Ind/@w:hanging"/>    
              </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
        <xsl:attribute name="text:min-label-width">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
                  <xsl:value-of select="number($Ind/@w:hanging)"/>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
        <xsl:attribute name="text:min-label-distance">
              <xsl:choose>
                <xsl:when test="number($Ind/@w:hanging) &lt; (number($tab/@pos) - number($Ind/@w:left)) ">
                  <xsl:value-of select="number($tab/@pos) - number($Ind/@w:left)"/>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="text:space-before">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
                  <xsl:value-of select="number($Ind/@w:left)-number($Ind/@w:firstLine)"/>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
        <xsl:attribute name="text:min-label-width">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:value-of select="number($Ind/@w:firstLine)"/>
              </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
        <xsl:attribute name="text:min-label-distance">
              <xsl:choose>
                <xsl:when test="(3 * number($Ind/@w:firstLine)) &lt; (number($tab/@pos) - number($Ind/@w:left)) ">
                  <xsl:value-of select="number($tab/@pos) - number($Ind/@w:left) - (2 * number($Ind/@w:firstLine))"/>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
         </xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- types of bullets -->
  
  <xsl:template name="TextChar">
    <xsl:choose>
      <xsl:when test="w:lvlText[@w:val = ''] "></xsl:when>
      <xsl:when test="w:lvlText[@w:val = '' ]"></xsl:when>
      <xsl:when test="w:lvlText[@w:val = '' ]">☑</xsl:when>
      <xsl:when test="w:lvlText[@w:val = '•' ]">•</xsl:when>
      <xsl:when test="w:lvlText[@w:val = '' ]">●</xsl:when>
      <xsl:when test="w:lvlText[@w:val = '' ]">➢</xsl:when>
      <xsl:when test="w:lvlText[@w:val = '' ]">✔</xsl:when>
      <xsl:when test="w:lvlText[@w:val = '' ]">■</xsl:when>
      <xsl:when test="w:lvlText[@w:val = 'o' ]">○</xsl:when>
      <xsl:when test="w:lvlText[@w:val = '' ]">➔</xsl:when>
      <xsl:when test="w:lvlText[@w:val = '' ]">✗</xsl:when>
      <xsl:when test="w:lvlText[@w:val = '-' ]">–</xsl:when>
      <xsl:otherwise>•</xsl:otherwise>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>

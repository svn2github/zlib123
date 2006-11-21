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
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" exclude-result-prefixes="w">

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
        <xsl:value-of select="$id"/>
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
            <xsl:apply-templates
              select="key('numId',$id)/w:lvlOverride[@w:ilvl = $level]/w:lvl[@w:ilvl = $level]"/>
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
        <text:list-level-style-bullet text:level="{number(@w:ilvl)+1}"
          text:style-name="Bullet_20_Symbols">
          <xsl:attribute name="text:bullet-char">
            <xsl:call-template name="TextChar"/>
          </xsl:attribute>
          <style:list-level-properties>
            <xsl:call-template name="InsertListLevelProperties"/>
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
          <xsl:if test="w:start and w:start/@w:val > 1">
            <xsl:attribute name="text:start-value">
              <xsl:value-of select="w:start/@w:val"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:variable name="display">
            <xsl:call-template name="CountDisplayListLevels">
              <xsl:with-param name="string">
                <xsl:value-of select="./w:lvlText/@w:val"/>
              </xsl:with-param>
              <xsl:with-param name="count">0</xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:if test="$display &gt; 1">
            <xsl:attribute name="text:display-levels">
              <xsl:value-of select="$display"/>
            </xsl:attribute>
          </xsl:if>
          <style:list-level-properties>
            <xsl:call-template name="InsertListLevelProperties"/>
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

  <xsl:template name="InsertListLevelProperties">
    <xsl:variable name="Ind" select="w:pPr/w:ind"/>
    <xsl:variable name="tab">
      <xsl:choose>
        <xsl:when test="w:pPr/w:tabs/w:tab">
          <xsl:value-of select="w:pPr/w:tabs/w:tab/@w:pos"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="document('word/settings.xml')//w:settings/w:defaultTabStop/@w:val"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

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
            <xsl:when test="number($Ind/@w:hanging) &lt; (number($tab) - number($Ind/@w:left)) ">
              <xsl:value-of select="number($tab) - number($Ind/@w:left)"/>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="text:space-before">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:value-of select="number($Ind/@w:left) + number($Ind/@w:firstLine)"/>
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
            <xsl:when
              test="(3 * number($Ind/@w:firstLine)) &lt; (number($tab) - number($Ind/@w:left)) ">
              <xsl:value-of
                select="number($tab) - number($Ind/@w:left) - (2 * number($Ind/@w:firstLine))"/>
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
      <xsl:when test="w:lvlText[@w:val = '–' ]">–</xsl:when>
      <xsl:otherwise>•</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- checks for list numPr properties (numid or level) in styles hierarchy -->
  <xsl:template name="GetListStyleProperty">
    <xsl:param name="style"/>
    <xsl:param name="property"/>

    <xsl:choose>
      <xsl:when test="$style/descendant::w:numPr">
        <xsl:choose>
          <xsl:when test="$property = 'w:ilvl' ">
            <xsl:choose>
              <xsl:when
                test="$style/descendant::w:numPr/w:numId and not($style/descendant::w:numPr/w:ilvl)">
                <xsl:text>0</xsl:text>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of
                  select="$style/descendant::w:numPr/child::node()[name() = $property]/@w:val"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of
              select="$style/descendant::w:numPr/child::node()[name() = $property]/@w:val"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="w:basedOn">
          <xsl:variable name="parentStyle"
            select="document('word/styles.xml')/w:styles/w:style[@w:styleId = w:basedOn/@w:val]"/>
          <xsl:call-template name="GetListStyleProperty">
            <xsl:with-param name="style" select="$parentStyle"/>
            <xsl:with-param name="property" select="$property"/>
          </xsl:call-template>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- heading numbering - gets heading num id which we need to generate text:outline-style-->
  <xsl:template name="GetOutlineListNumId">
    <xsl:variable name="headingElement" select="document('word/document.xml')/descendant::w:p[descendant::w:outlineLvl]"/>
    <!-- take heading style which is used in the document (it can have a numId) -->
    <xsl:variable name="headingStyle"
      select="document('word/styles.xml')/w:styles/w:style[descendant::w:outlineLvl 
      and document('word/document.xml')/descendant::w:p/descendant::w:pStyle/@w:val = @w:styleId]"/>

    <xsl:choose>
      <!-- when the outlineLvl is found in paragraph check for the numid in it and in it's style -->
      <xsl:when test="$headingElement">
        <xsl:variable name="headingNumId">
          <xsl:call-template name="GetListProperty">
            <xsl:with-param name="node" select="$headingElement"/>
            <xsl:with-param name="property" >w:numId</xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:if test="$headingNumId != ''">
          <xsl:value-of select="$headingNumId"/>
        </xsl:if>
      </xsl:when>
      
      <!-- when the outlineLvl is found in style check for the numid in it and in paragraph with this style -->
      <xsl:when test="$headingStyle">
         <xsl:variable name="styleNumId">
          <xsl:call-template name="GetListStyleProperty">
            <xsl:with-param name="style" select="$headingStyle"/>
            <xsl:with-param name="property">w:numId</xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:choose>
          <!--check for numId in style-->
          <xsl:when test="$styleNumId != '' ">
            <xsl:value-of select="$styleNumId"/>
          </xsl:when>
          <!--check for numId in paragraph which has the style -->
          <xsl:otherwise>
            <xsl:variable name="headingNumId">
              <xsl:call-template name="GetOutlineItemNumId">
                <xsl:with-param name="style" select="$headingStyle"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:if test="$headingNumId != ''">
              <xsl:value-of select="$headingNumId"/>
            </xsl:if>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- heading numbering - gets numid from element with given style and linked styles -->
  <xsl:template name="GetOutlineItemNumId">
    <xsl:param name="style"/>

    <xsl:variable name="styleId" select="$style/@w:styleId"/>
    <xsl:variable name="headingElement"
      select="document('word/document.xml')/descendant::w:p[descendant::w:pStyle/@w:val = $styleId and descendant::w:numId]"/>

    <xsl:choose>
      <xsl:when test="not($headingElement)">
        <xsl:variable name="linkedStyle" select="w:styles/w:style[@styleId = $style/w:link/@w:val]"/>
        <xsl:if test="$linkedStyle">
          <xsl:call-template name="GetOutlineItemNumId">
            <xsl:with-param name="style" select="$linkedStyle"/>
          </xsl:call-template>
        </xsl:if>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$headingElement/descendant::w:numId/@w:val"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- checks for list numPr properties (numid or level) for given element  -->
  <xsl:template name="GetListProperty">
    <xsl:param name="node"/>
    <xsl:param name="property" />
    
    <xsl:choose>
      <xsl:when test="$node/descendant::w:numPr">
        <xsl:value-of select="$node/descendant::w:numPr/child::node()[name() = $property]/@w:val"/>
      </xsl:when>
      
      <xsl:when test="$node/descendant::w:pStyle">
        <xsl:variable name="styleId" select="$node/descendant::w:pStyle/@w:val"/>
        
        <xsl:variable name="pStyle"
          select="document('word/styles.xml')/w:styles/w:style[@w:styleId = $styleId]"/>
        <xsl:variable name="propertyValue">
          <xsl:call-template name="GetListStyleProperty">
            <xsl:with-param name="style" select="$pStyle"/>
            <xsl:with-param name="property" select="$property"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:value-of select="$propertyValue"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <!-- heading list display levels  -->
    <xsl:template name="CountDisplayListLevels">
    <xsl:param name="string"/>
    <xsl:param name="count"/>
    <xsl:choose>
      <xsl:when test="string-length(substring-after($string,'%')) &gt; 0">
        <xsl:call-template name="CountDisplayListLevels">
          <xsl:with-param name="string">
            <xsl:value-of select="substring-after($string,'%')"/>
          </xsl:with-param>
          <xsl:with-param name="count">
            <xsl:value-of select="$count +1"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$count"/>
      </xsl:otherwise>
    </xsl:choose>
    </xsl:template>
  
  
  <!-- paragraph which is the first element of a list-->
  
  <xsl:template match="w:p" mode="list">
    <xsl:param name="numId"/>
    <xsl:param name="nestedLevel" select="w:pPr/w:numPr/w:ilvl/@w:val"/>
    <xsl:param name="level" select="w:pPr/w:numPr/w:ilvl/@w:val"/>
    <xsl:variable name="position" select="count(preceding-sibling::w:p)"/>
 
    <!-- if first element of a list -->
    <xsl:if
      test="not(preceding-sibling::node()[child::w:pPr/w:numPr/w:numId/@w:val = $numId and count(preceding-sibling::w:p)= $position -1])">
      <text:list text:style-name="{concat('L',$numId)}">
        <xsl:if
          test="preceding-sibling::w:p[child::w:pPr/w:numPr[w:numId/@w:val = $numId and w:ilvl/@w:val = $level]] or document('word\numbering.xml')//w:numbering/w:abstractNum[@w:abstractNumId = document('word\numbering.xml')//w:numbering/w:num[@w:numId = $numId]/w:abstractNumId/@w:val]/w:lvl[@w:ilvl = $level]/w:start">
          <xsl:attribute name="text:continue-numbering">true</xsl:attribute>
        </xsl:if>
        
        
        <!-- convert element as list item -->
        <xsl:apply-templates select="." mode="list-item">
          <xsl:with-param name="nestedLevel" select="$nestedLevel"/>
          <xsl:with-param name="level">
            <xsl:value-of select="w:pPr/w:numPr/w:ilvl/@w:val"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </text:list>
    </xsl:if>
  </xsl:template>
  
  <!-- pargraph which is a list-item -->
  
  <xsl:template match="w:p" mode="list-item">
    <xsl:param name="nestedLevel"/>
    <xsl:param name="level"/>
    <xsl:variable name="NumberingId" select="w:pPr/w:numPr/w:numId/@w:val"/>
    <xsl:variable name="position" select="count(preceding-sibling::w:p)"/>
    <xsl:variable name="notHigherLevelPosition"
      select="count(preceding-sibling::w:p[not(w:pPr/w:numPr/w:ilvl/@w:val &gt; $level - $nestedLevel)])"/>
    <xsl:choose>
      
      <!-- if there's a nested list we call the template recursively -->
      <xsl:when test="$nestedLevel &gt; 0">
        <text:list-item>
          <text:list text:style-name="{concat('L',$NumberingId)}">
            <xsl:if
              test="document('word\numbering.xml')//w:numbering/w:abstractNum[@w:abstractNumId = document('word\numbering.xml')//w:numbering/w:num[@w:numId = $NumberingId]/w:abstractNumId/@w:val]/w:lvl[@w:ilvl = $level]/w:start">
              <xsl:attribute name="text:continue-numbering">true</xsl:attribute>
            </xsl:if>
            <xsl:apply-templates select="." mode="list-item">
              <xsl:with-param name="nestedLevel" select="$nestedLevel -1"/>
              <xsl:with-param name="level" select="$level"/>
            </xsl:apply-templates>
          </text:list>
        </text:list-item>
        
        <!-- next paragraph on this level in same list -->
        <xsl:variable name="nextElement"
          select="following-sibling::w:p[child::w:pPr/w:numPr[w:numId/@w:val = $NumberingId and w:ilvl/@w:val = $level - $nestedLevel] and count(preceding-sibling::w:p[not(w:pPr/w:numPr/w:ilvl/@w:val &gt; $level - $nestedLevel)])= $notHigherLevelPosition]"/>
        <xsl:if test="$nextElement">
          <xsl:apply-templates select="$nextElement" mode="list-item">
            <xsl:with-param name="nestedLevel">
              <xsl:value-of select="$nextElement/w:pPr/w:numPr/w:ilvl/@w:val - $level+$nestedLevel"
              />
            </xsl:with-param>
            <xsl:with-param name="level">
              <xsl:value-of select="$nextElement/w:pPr/w:numPr/w:ilvl/@w:val"/>
            </xsl:with-param>
          </xsl:apply-templates>
        </xsl:if>
      </xsl:when>
      
      <xsl:otherwise>
        <xsl:variable name="outlineLevel">
          <xsl:call-template name="GetOutlineLevel">
            <xsl:with-param name="node" select="."/>
          </xsl:call-template>
        </xsl:variable>
        <text:list-item>
          <xsl:choose>
            <xsl:when test="$outlineLevel != ''">
                <xsl:apply-templates select="." mode="heading"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:apply-templates select="." mode="paragraph"/>
            </xsl:otherwise>
          </xsl:choose>
        </text:list-item>
        
        <!-- next paragraph  in same list -->
        <xsl:variable name="nextListItem"
          select="following-sibling::w:p[child::w:pPr/w:numPr[w:numId/@w:val = $NumberingId and not(w:ilvl/@w:val &lt; $level - $nestedLevel)] and count(preceding-sibling::w:p)= $position +1]"/>
        <xsl:if test="$nextListItem">
          <xsl:apply-templates select="$nextListItem" mode="list-item">
            <xsl:with-param name="nestedLevel">
              <xsl:value-of select="$nextListItem/w:pPr/w:numPr/w:ilvl/@w:val - $level+$nestedLevel"
              />
            </xsl:with-param>
            <xsl:with-param name="level">
              <xsl:value-of select="$nextListItem/w:pPr/w:numPr/w:ilvl/@w:val"/>
            </xsl:with-param>
          </xsl:apply-templates>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  </xsl:stylesheet>

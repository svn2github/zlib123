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
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
  exclude-result-prefixes="w text style">
  
  <xsl:key name="OutlineLevel" match="w:outlineLvl" use="''"/>

  <!-- paragraph which starts table of content -->
  <xsl:template match="w:p" mode="tocstart">
    <xsl:choose>
      <xsl:when test="w:r[contains(w:instrText,'TOC')]">
        <xsl:choose>
          <xsl:when test="w:pPr[contains(w:pStyle/@w:val,'TOC') or contains(w:pStyle/@w:val,'Contents_20')]">
        <text:table-of-content text:style-name="Sect1">
          <xsl:message terminate="no">feedback:TOC protection against manual changes</xsl:message>
          <xsl:attribute name="text:protected">false</xsl:attribute>
          <xsl:attribute name="text:name">Table of Contents1</xsl:attribute>
          <xsl:call-template name="InsertIndexProperties">
            <xsl:with-param name="type">TOC</xsl:with-param>
          </xsl:call-template>
          <text:index-body>
            <xsl:apply-templates select="." mode="index"/>
          </text:index-body>
        </text:table-of-content>
          </xsl:when>
          <xsl:otherwise>
            <text:table-index text:style-name="Sect1">
              <xsl:message terminate="no">feedback:TOC protection against manual changes</xsl:message>
              <xsl:attribute name="text:protected">false</xsl:attribute>
              <xsl:attribute name="text:name">Index of Tables1</xsl:attribute>
              <xsl:call-template name="InsertIndexProperties">
                <xsl:with-param name="type">INDEX</xsl:with-param>
              </xsl:call-template>
              <text:index-body>
                <xsl:apply-templates select="." mode="index"/>
              </text:index-body>
            </text:table-index>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="w:r[contains(w:instrText,'BIBLIOGRAPHY')]">
        <text:bibliography>
          <xsl:attribute name="text:name">
            <xsl:value-of select="generate-id(w:r[contains(w:instrText,'BIBLIOGRAPHY')])"/>
          </xsl:attribute>
          <text:bibliography-source>
            <text:index-title-template/>
            <xsl:call-template name="InsertBibliographyProperties"/>
          </text:bibliography-source>
          <text:index-body>            
            <xsl:apply-templates select="." mode="index"/>
          </text:index-body>
        </text:bibliography>        
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <!-- paragraph in index-->
  <xsl:template match="w:p" mode="index">
    <text:p text:style-name="{generate-id()}">
      <xsl:apply-templates mode="index"/>
    </text:p>
    <xsl:if test="following-sibling::w:p[1][count(preceding::w:fldChar[@w:fldCharType='begin']) &gt; count(preceding::w:fldChar[@w:fldCharType='end']) and descendant::text()]">
      <xsl:apply-templates select="following-sibling::w:p[1]" mode="index"/>
    </xsl:if>
  </xsl:template>
  
  <!--take content from multiple w:instrText elements -->
  <xsl:template match="w:instrText" mode="textContent">
    <xsl:param name="textContent"/>
    <xsl:variable name="text">
      <xsl:value-of select="."/>
    </xsl:variable>
    <xsl:choose>
   <xsl:when test="following-sibling::w:instrText">
      <xsl:apply-templates select="following-sibling::w:instrText[1]" mode="textContent">
        <xsl:with-param name="textContent">
          <xsl:value-of select="concat($textContent,$text)"/>
        </xsl:with-param>
      </xsl:apply-templates>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="concat($textContent,$text)"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="GetMaxLevelFromStyles">
    <xsl:param name="stylesWithLevels"/>
    <xsl:param name="defaultLevel"/>
    <xsl:choose>
      <xsl:when test="$stylesWithLevels!=''">
        <xsl:variable name="firstNum">
          <xsl:choose>
            <xsl:when test="contains(substring-after($stylesWithLevels,';'),';')">
              <xsl:value-of select="number(substring-before(substring-after($stylesWithLevels,';'),';'))"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="number(substring-after($stylesWithLevels,';'))"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="otherNum">
          <xsl:call-template name="GetMaxLevelFromStyles">
            <xsl:with-param name="stylesWithLevels">
              <xsl:value-of select="substring-after(substring-after($stylesWithLevels,';'),';')"/>
            </xsl:with-param>
            <xsl:with-param name="defaultLevel" select="$defaultLevel"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="num">
        <xsl:choose>
          <xsl:when test="number($firstNum) &gt; number($otherNum)">
            <xsl:value-of select ="$firstNum"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$otherNum"/>
          </xsl:otherwise>
        </xsl:choose>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="number($num) &gt; number($defaultLevel)">
            <xsl:value-of select="$num"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$defaultLevel"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>0</xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="GetMinLevelFromStyles">
    <xsl:param name="stylesWithLevels"/>
    <xsl:param name="defaultLevel"/>
    <xsl:choose>
      <xsl:when test="$stylesWithLevels!=''">
        <xsl:variable name="firstNum">
          <xsl:choose>
            <xsl:when test="contains(substring-after($stylesWithLevels,';'),';')">
              <xsl:value-of select="number(substring-before(substring-after($stylesWithLevels,';'),';'))"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="number(substring-after($stylesWithLevels,';'))"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="otherNum">
          <xsl:call-template name="GetMinLevelFromStyles">
            <xsl:with-param name="stylesWithLevels">
              <xsl:value-of select="substring-after(substring-after($stylesWithLevels,';'),';')"/>
            </xsl:with-param>
            <xsl:with-param name="defaultLevel" select="$defaultLevel"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="num">
        <xsl:choose>
          <xsl:when test="number($firstNum) &lt; number($otherNum)">
            <xsl:value-of select ="$firstNum"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$otherNum"/>
          </xsl:otherwise>
        </xsl:choose>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="number($num) &lt; number($defaultLevel)">
            <xsl:value-of select="$num"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$defaultLevel"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>10</xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="GetStyleForLevel">
    <xsl:param name="stylesWithLevels"/>
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test="contains(substring-after($stylesWithLevels,';'),';')">
        <xsl:choose>
          <xsl:when test="number(substring-before(substring-after($stylesWithLevels,';'),';'))=number($level)">
            <xsl:value-of select="number(substring(substring-before($stylesWithLevels,';'),string-length(substring-before($stylesWithLevels,';'))))"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="GetStyleForLevel">
              <xsl:with-param name="stylesWithLevels">
                <xsl:value-of select="substring-after(substring-after($stylesWithLevels,';'),';')"/>
              </xsl:with-param>
              <xsl:with-param name="level" select="$level"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="number(substring-after($stylesWithLevels,';'))=number($level)">
            <xsl:value-of select="number(substring(substring-before($stylesWithLevels,';'),string-length(substring-before($stylesWithLevels,';'))))"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$level"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
<xsl:template name="GetOutlineLevelMax">
    <xsl:param name="value"/>
    <xsl:param name="count"/>
      <xsl:variable name="getValue">
        <xsl:value-of select="key('OutlineLevel', '')[$count]/@w:val"/>
      </xsl:variable>
  <xsl:choose>
    <xsl:when test="$count > 0">      
      <xsl:call-template name="GetOutlineLevelMax">
        <xsl:with-param name="value">
          <xsl:choose>
            <xsl:when test="number($getValue) > number($value)">
              <xsl:value-of select="$getValue"/>
            </xsl:when>
            <xsl:otherwise><xsl:value-of select="$value"/></xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
        <xsl:with-param name="count">
          <xsl:value-of select="$count - 1"/>
        </xsl:with-param>
      </xsl:call-template>      
    </xsl:when>
    <xsl:otherwise>
      <xsl:value-of select="$value + 1"/>
    </xsl:otherwise>
  </xsl:choose>
</xsl:template>

  <!-- insert table-of content properties -->
  <xsl:template name="InsertIndexProperties">
    <xsl:param name="type"/>
    <xsl:variable name="instrTextContent">
      <xsl:apply-templates select="w:r/w:instrText[1]" mode="textContent">
        <xsl:with-param name="textContent"/>
      </xsl:apply-templates>
    </xsl:variable>
  
    <xsl:variable name="maxLevel">
      <xsl:choose>
        <xsl:when test="contains($instrTextContent,'\t')">
          <xsl:call-template name="GetMaxLevelFromStyles">
            <xsl:with-param name="stylesWithLevels">
              <xsl:value-of select="substring-before(substring-after(substring-after($instrTextContent,'\t'),'&quot;'),'&quot;')"/>
            </xsl:with-param>
            <xsl:with-param name="defaultLevel">
              <xsl:choose>
                <xsl:when test="contains($instrTextContent,'-')">
                  <xsl:value-of select="substring-before(substring-after($instrTextContent,'-'),'&quot;')"/>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="contains($instrTextContent,'-')">
          <xsl:value-of select="substring-before(substring-after($instrTextContent,'-'),'&quot;')"/>
        </xsl:when>
        
        <xsl:when test="contains($instrTextContent,'\o')"> 
        
          <xsl:variable name="CountOutlineLevel">
            <xsl:for-each select="document('word/document.xml')">
              <xsl:value-of select="count(key('OutlineLevel', '')/@w:val)"/>
            </xsl:for-each>  
          </xsl:variable>         
          <xsl:choose>
            <xsl:when test="$CountOutlineLevel > 0">
              <xsl:for-each select="document('word/document.xml')">
                <xsl:call-template name="GetOutlineLevelMax">
                  <xsl:with-param name="value">0</xsl:with-param>
                  <xsl:with-param name="count">                    
                    <xsl:value-of select="$CountOutlineLevel"/>
                  </xsl:with-param>                 
                </xsl:call-template>
              </xsl:for-each>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>         

        </xsl:when>        
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$type='TOC'">
    <text:table-of-content-source text:outline-level="{$maxLevel}">
      <xsl:call-template name="InsertContentOfIndexProperties">
        <xsl:with-param name="styleName">Contents_20_Heading</xsl:with-param>
        <xsl:with-param name="maxLevel" select="$maxLevel"/>
        <xsl:with-param name="instrTextContent" select="$instrTextContent"/>
        <xsl:with-param name="type" select="$type"/>
      </xsl:call-template>
    </text:table-of-content-source>
      </xsl:when>
      <xsl:otherwise>
        <text:table-index-source text:caption-sequence-name="{substring-before(substring-after($instrTextContent,'&quot;'),'&quot;')}" text:caption-sequence-format="text">
          <xsl:call-template name="InsertContentOfIndexProperties">
            <xsl:with-param name="styleName">Table_20_index_20_heading</xsl:with-param>
            <xsl:with-param name="maxLevel" select="$maxLevel"/>
            <xsl:with-param name="instrTextContent" select="$instrTextContent"/>
          </xsl:call-template>
          </text:table-index-source>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="InsertContentOfIndexProperties">
    <xsl:param name="styleName"/>
    <xsl:param name="maxLevel"/>
    <xsl:param name="instrTextContent"/>
    <xsl:param name="type"/>
    <xsl:if test="$maxLevel = 0"/>
    <text:index-title-template text:style-name="{$styleName}"/>
    <xsl:variable name="level">
      <xsl:choose>
        <xsl:when test="contains($instrTextContent,'\t')">
          <xsl:call-template name="GetMinLevelFromStyles">
            <xsl:with-param name="stylesWithLevels">
              <xsl:value-of select="substring-before(substring-after(substring-after($instrTextContent,'\t'),'&quot;'),'&quot;')"/>
            </xsl:with-param>
            <xsl:with-param name="defaultLevel">
              <xsl:choose>
                <xsl:when test="contains($instrTextContent,'-')">
                  <xsl:value-of select="substring-after(substring-before($instrTextContent,'-'),'&quot;')"/>
                </xsl:when>
                <xsl:otherwise>10</xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="contains($instrTextContent,'-')">
          <xsl:value-of select="substring-after(substring-before($instrTextContent,'-'),'&quot;')"/>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:call-template name="InsertIndexEntryProperties">
      <xsl:with-param name="level" select="$level"/>
      <xsl:with-param name="maxLevel" select="$maxLevel"/>
      <xsl:with-param name="instrTextContent" select="$instrTextContent"/>
      <xsl:with-param name="type" select="$type"/>
    </xsl:call-template>
  </xsl:template>
  
  <!-- insert entry properties -->
  <xsl:template name="InsertIndexEntryProperties">
    <xsl:param name="level"/>
    <xsl:param name="maxLevel"/>
    <xsl:param name="instrTextContent"/>
    <xsl:param name="type"/>
    <xsl:variable name="node" select="self::node()"/>
    <xsl:variable name="levelForStyle">
      <xsl:call-template name="GetStyleForLevel">
        <xsl:with-param name="stylesWithLevels">
          <xsl:value-of select="substring-before(substring-after(substring-after($instrTextContent,'\t'),'&quot;'),'&quot;')"/>
        </xsl:with-param>
        <xsl:with-param name="level" select="$level"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="not(number($level) &gt; number($maxLevel)) and $levelForStyle!='NaN'">
        <xsl:choose>
        <xsl:when test="$type='TOC'">
        <text:table-of-content-entry-template text:outline-level="{$level}">
          <xsl:call-template name="InsertContentOfIndexEntryProperties">
          <xsl:with-param name="level" select="$level"/>
            <xsl:with-param name="instrTextContent" select="$instrTextContent"/>
            <xsl:with-param name="levelForStyle" select="$levelForStyle"/>
            <xsl:with-param name="node" select="$node"/>
            <xsl:with-param name="type" select="$type"/>
          </xsl:call-template>
        </text:table-of-content-entry-template>
        </xsl:when>
          <xsl:otherwise>
            <text:table-index-entry-template>
              <xsl:call-template name="InsertContentOfIndexEntryProperties">
                <xsl:with-param name="level" select="$level"/>
                <xsl:with-param name="instrTextContent" select="$instrTextContent"/>
                <xsl:with-param name="levelForStyle" select="$levelForStyle"/>
                <xsl:with-param name="node" select="$node"/>
                <xsl:with-param name="type" select="$type"/>
              </xsl:call-template>
            </text:table-index-entry-template>
          </xsl:otherwise>
          </xsl:choose>
        <xsl:call-template name="InsertIndexEntryProperties">
          <xsl:with-param name="level" select="number($level)+1"/>
          <xsl:with-param name="maxLevel" select="$maxLevel"/>
          <xsl:with-param name="instrTextContent" select="$instrTextContent"/>
          <xsl:with-param name="type" select="$type"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise></xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="InsertContentOfIndexEntryProperties">
    <xsl:param name="level"/>
    <xsl:param name="instrTextContent"/>
    <xsl:param name="levelForStyle"/>
    <xsl:param name="node"/>
    <xsl:param name="type"/>
    <xsl:if test="$level = 0"/>
    <xsl:attribute name="text:style-name">
      <xsl:choose>
        <xsl:when test="$level=0">
          <xsl:call-template name="TocToContent">
            <xsl:with-param name="styleValue">
              <xsl:value-of select="w:pPr/w:pStyle/@w:val"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="contains($instrTextContent,'\t')">
          <xsl:value-of select="concat('Contents_20_',$levelForStyle)"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="concat('Contents_20_',$level)"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:call-template name="EntryIterator">
      <xsl:with-param name="fieldCharCount">0</xsl:with-param>
      <xsl:with-param name="level" select="$level"/>
      <xsl:with-param name="node" select="$node"/>
      <xsl:with-param name="instrTextContent" select="$instrTextContent"/>
      <xsl:with-param name="type" select="$type"/>
    </xsl:call-template>
  </xsl:template>
  
  <!--template which changes style 'TOC' to style 'Contents_20'-->
  <xsl:template name="TocToContent">
    <xsl:param name="styleValue"/>
    <xsl:choose>
      <xsl:when test="contains($styleValue,'TOC')">
        <xsl:value-of select="concat('Contents_20_',substring-after($styleValue,'TOC'))"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$styleValue"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- searches entry properties through all toc -->
  <xsl:template name="EntryIterator">
    <xsl:param name="fieldCharCount"/>
    <xsl:param name="level"/>
    <xsl:param name="node"/>
    <xsl:param name="instrTextContent"/>
    <xsl:param name="type"/>
    <xsl:variable name="styleLevel">
      <xsl:choose>
        <xsl:when test="contains($instrTextContent,'\t')">
          <xsl:variable name="levelForStyle">
            <xsl:call-template name="GetStyleForLevel">
              <xsl:with-param name="stylesWithLevels">
                <xsl:value-of select="substring-before(substring-after(substring-after($instrTextContent,'\t'),'&quot;'),'&quot;')"/>
              </xsl:with-param>
              <xsl:with-param name="level" select="$level"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test="$levelForStyle!='NaN'">
              <xsl:value-of select="$levelForStyle"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$level"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$level"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="count">
      <xsl:choose>
        <xsl:when test="descendant::w:fldChar/@w:fldCharType='begin'">
          <xsl:value-of select="number($fieldCharCount)+1"/>
        </xsl:when>
        <xsl:when test="descendant::w:fldChar/@w:fldCharType='end'">
          <xsl:value-of select="number($fieldCharCount)-1"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="number($fieldCharCount)"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:if test="$count &gt; 0">
      <xsl:choose>
        <xsl:when test="(contains(w:pPr/w:pStyle/@w:val,$styleLevel) and not(contains(preceding-sibling::w:p[(preceding-sibling::node()=$node or self::node()=$node)]/w:pPr/w:pStyle/@w:val,$styleLevel))) or $styleLevel = 0">
          <xsl:if test="$type='TOC'">
            <text:index-entry-chapter/>
          </xsl:if>
          <text:index-entry-text/>
          <xsl:apply-templates select="(descendant::w:r/w:tab)[number(last())]" mode="entry"/>
          <xsl:apply-templates select="(descendant::w:r/w:t)[number(last())]" mode="entry"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="following-sibling::w:p[1]">
            <xsl:call-template name="EntryIterator">
              <xsl:with-param name="fieldCharCount" select="$count"/>
              <xsl:with-param name="level" select="$level"/>
              <xsl:with-param name="node" select="$node"/>
              <xsl:with-param name="instrTextContent" select="$instrTextContent"/>
              <xsl:with-param name="type" select="$type"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
  </xsl:template> 

  <!--insert Bibliography properties -->
  
  <xsl:template name="InsertBibliographyProperties">
    
    <xsl:variable name="Style">
      <xsl:value-of select="document('customXml/item1.xml')/b:Sources/@StyleName"/>
    </xsl:variable> 
    
    <text:bibliography-entry-template text:bibliography-type="book">
      <xsl:attribute name="text:style-name">
        <xsl:value-of select="generate-id()"/>
      </xsl:attribute>
    <xsl:choose>
      <xsl:when test="$Style = 'APA'">
        <text:index-entry-bibliography text:bibliography-data-field="author"/><text:index-entry-span>.(</text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="year"/><text:index-entry-span>). </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="title"/><text:index-entry-span>. </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="address"/><text:index-entry-span>: </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="publisher"/>
      </xsl:when>
      <xsl:when test="$Style = 'Chicago' or $Style = 'ISO 690 - Numerical Reference' or $Style = 'MLA' or $Style = 'Turabian'">        
        <text:index-entry-bibliography text:bibliography-data-field="author"/>        
        <text:index-entry-span>. </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="title"/>
        <text:index-entry-span>. </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="address"/>
        <text:index-entry-span> : </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="publisher"/>
        <text:index-entry-span>, </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="year"/> 
        <text:index-entry-span>. </text:index-entry-span>
      </xsl:when>
      <xsl:when test="$Style = 'GB7714'">
        <text:index-entry-bibliography text:bibliography-data-field="author"/>
        <text:index-entry-span>.</text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="year"/>
        <text:index-entry-span>.</text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="title"/>
        <text:index-entry-span>. </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="address"/>
        <text:index-entry-span>: </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="publisher"/>
        <text:index-entry-span>, </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="year"/>
        <text:index-entry-span>. </text:index-entry-span>
      </xsl:when>
      <xsl:when test="$Style = 'GOST - Name Sort'">
        <text:index-entry-bibliography text:bibliography-data-field="author"/>
        <text:index-entry-span> </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="title"/>
        <text:index-entry-span>[</text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="bibliography-type"/>
        <text:index-entry-span>].-</text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="address"/>
        <text:index-entry-span> : </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="publisher"/>        
        <text:index-entry-span>, </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="year"/>
        <text:index-entry-span>. </text:index-entry-span>
      </xsl:when>
      <xsl:when test="$Style = 'GOST - Title Sort'">
        <text:index-entry-bibliography text:bibliography-data-field="title"/>
        <text:index-entry-span> </text:index-entry-span>
        <text:index-entry-span>[</text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="bibliography-type"/>
        <text:index-entry-span>]/ AUT. </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="author"/>
        <text:index-entry-span>. - </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="address"/>
        <text:index-entry-span> : </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="publisher"/>        
        <text:index-entry-span>, </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="year"/>
        <text:index-entry-span>. </text:index-entry-span>
      </xsl:when>
      <xsl:when test="$Style = 'ISO 690 - First Element and Date'">
       
        <text:index-entry-bibliography text:bibliography-data-field="author"/>
        <text:index-entry-span>. </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="year"/>
        <text:index-entry-span>. </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="title"/>
        <text:index-entry-span>. </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="address"/>
        <text:index-entry-span> : </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="publisher"/>
        <text:index-entry-span>, </text:index-entry-span>        
        <text:index-entry-bibliography text:bibliography-data-field="year"/>
        <text:index-entry-span>. </text:index-entry-span>
      </xsl:when>
      <xsl:when test="$Style = 'SIST02'">        
        <text:index-entry-bibliography text:bibliography-data-field="author"/>        
        <text:index-entry-span> </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="title"/>
        <text:index-entry-span>. </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="address"/>
        <text:index-entry-span>, </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="publisher"/>
        <text:index-entry-span>, </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="year"/> 
        <text:index-entry-span>. </text:index-entry-span>
      </xsl:when>
      <xsl:otherwise>
        <text:index-entry-bibliography text:bibliography-data-field="author"/>        
        <text:index-entry-span>. </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="title"/>
        <text:index-entry-span>. </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="address"/>
        <text:index-entry-span> : </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="publisher"/>
        <text:index-entry-span>, </text:index-entry-span>
        <text:index-entry-bibliography text:bibliography-data-field="year"/> 
        <text:index-entry-span>. </text:index-entry-span>
      </xsl:otherwise>
    </xsl:choose>
      </text:bibliography-entry-template>
 </xsl:template>
  
  <!-- insert page number -->
  <xsl:template match="w:t" mode="entry">
    <xsl:variable name="text">
      <xsl:value-of select="child::text()"/>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="number($text)">
        <text:index-entry-page-number/>
      </xsl:when>
      <xsl:otherwise/>
    </xsl:choose>
  </xsl:template>
  
  <!-- insert entry properties for tabs -->
  <xsl:template match="w:tab[not(parent::w:tabs)]" mode="entry">
    <xsl:variable name="tabCount">
      <xsl:value-of select="count(parent::w:r/preceding-sibling::w:r/w:tab)+1"/>
    </xsl:variable>
    <xsl:variable name="styleType">
      <xsl:call-template name="GetTabParams">
        <xsl:with-param name="param">w:val</xsl:with-param>
        <xsl:with-param name="tabCount" select="$tabCount"/>
        <xsl:with-param name="maxtabCount" select="$tabCount"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="leaderChar">
      <xsl:call-template name="GetTabParams">
        <xsl:with-param name="param">w:leader</xsl:with-param>
        <xsl:with-param name="tabCount" select="$tabCount"/>
        <xsl:with-param name="maxtabCount" select="$tabCount"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:if test="$styleType != '' and $styleType != 'clear'">
    <text:index-entry-tab-stop style:type="{$styleType}">
      <!--if style type is left, there must be style:position attribute -->
      
      <xsl:if test="$styleType = 'left'">
        <xsl:attribute name="style:position">
          <xsl:variable name="position">
            <xsl:call-template name="GetTabParams">
              <xsl:with-param name="param">w:pos</xsl:with-param>
              <xsl:with-param name="tabCount" select="$tabCount"/>
              <xsl:with-param name="maxtabCount" select="$tabCount"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length" select="$position"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="$leaderChar and $leaderChar!='' and $leaderChar!='heavy' and $leaderChar!='middleDot' and $leaderChar!='none'">
        <xsl:call-template name="InsertStyleLeaderChar">
          <xsl:with-param name="leaderChar" select="$leaderChar"/>
        </xsl:call-template>
      </xsl:if>
    </text:index-entry-tab-stop>
    </xsl:if>
  </xsl:template>
  
  <!-- insert tab-leader char -->
  <xsl:template name="InsertStyleLeaderChar">
    <xsl:param name="leaderChar"/>
    <xsl:attribute name="style:leader-char">
      <xsl:choose>
        <xsl:when test="$leaderChar='dot'">.</xsl:when>
        <xsl:when test="$leaderChar='hyphen'">-</xsl:when>
        <xsl:when test="$leaderChar='underscore'">_</xsl:when>
        <xsl:otherwise>none</xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>
  
  <!-- get properties of tabs -->
  <xsl:template name="GetTabParams">
    <xsl:param name="param"/>
    <xsl:param name="tabCount"/>
    <xsl:param name="maxtabCount"/>
    <xsl:choose>
      <xsl:when test="$tabCount > 0">
        <xsl:choose>
          <xsl:when test="preceding::w:tabs[1]/w:tab[number($tabCount)]">
            <xsl:choose>
              <xsl:when test="preceding::w:tabs[1]/w:tab[number($tabCount)]/attribute::node()[name()=$param]">
                <xsl:value-of select="preceding::w:tabs[1]/w:tab[number($tabCount)]/attribute::node()[name()=$param]"/>
              </xsl:when>
              <xsl:otherwise/>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="GetTabParams">
              <xsl:with-param name="param" select="$param"/>
              <xsl:with-param name="tabCount" select="number($tabCount)-1"/>
              <xsl:with-param name="maxtabCount" select="$maxtabCount"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="GetTabParamsFromStyles">
          <xsl:with-param name="tabStyle">
            <xsl:value-of select="ancestor::w:p/w:pPr/w:pStyle/@w:val"/>
          </xsl:with-param>
          <xsl:with-param name="attribute" select="$param"/>
          <xsl:with-param name="tabCount" select="$maxtabCount"/>
          <xsl:with-param name="maxtabCount" select="$maxtabCount"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!--get properties of tabs from styles.xml -->
  <xsl:template name="GetTabParamsFromStyles">
    <xsl:param name="tabStyle"/>
    <xsl:param name="attribute"/>
    <xsl:param name="tabCount"/>
    <xsl:param name="maxtabCount"/>
    
    <xsl:variable name="ancestor-style" select="ancestor::w:p/w:pPr/w:pStyle/@w:val"/>
    
    <xsl:for-each select="document('word/styles.xml')">
    <xsl:choose>
      <xsl:when test="$tabCount > 0">
        <xsl:choose>
          <xsl:when test="key('StyleId', $tabStyle)[1]/w:pPr/w:tabs/w:tab[number($tabCount)]">
            <xsl:choose>
              <xsl:when test="key('StyleId', $tabStyle)[1]/w:pPr/w:tabs/w:tab[number($tabCount)]/attribute::node()[name()=$attribute]">
                <xsl:value-of select="key('StyleId', $tabStyle)[1]/w:pPr/w:tabs/w:tab[number($tabCount)]/attribute::node()[name()=$attribute]"/>
              </xsl:when>
              <xsl:otherwise/>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="GetTabParamsFromStyles">
              <xsl:with-param name="tabStyle">
                <xsl:value-of select="$ancestor-style"/>
              </xsl:with-param>
              <xsl:with-param name="attribute" select="$attribute"/>
              <xsl:with-param name="tabCount" select="number($tabCount)-1"/>
              <xsl:with-param name="maxtabCount" select="$maxtabCount"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="key('StyleId', $tabStyle)[1]/w:basedOn/@w:val">
        <xsl:call-template name="GetTabParamsFromStyles">
          <xsl:with-param name="tabStyle">
            <xsl:value-of select="key('StyleId', $tabStyle)[1]/w:basedOn/@w:val"/>
          </xsl:with-param>
          <xsl:with-param name="attribute" select="$attribute"/>
          <xsl:with-param name="tabCount" select="$maxtabCount"/>
          <xsl:with-param name="maxtabCount" select="$maxtabCount"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise/>
    </xsl:choose>
    </xsl:for-each>
  </xsl:template>
  
  <!-- handle text in table-of content -->
  <xsl:template match="text()" mode="entry"/>
  
  <!-- handle runs -->
  <xsl:template match="w:r" mode="index">
    <xsl:choose>
      <!--  ignore text when we are in field-->
      <xsl:when test="w:fldChar or w:instrText"/>
      <!--attach style-->
      <xsl:when test="w:rPr">
        <text:span text:style-name="{generate-id(self::node())}">
          <xsl:apply-templates/>
        </text:span>
      </xsl:when>
      <!--default scenario-->
      <xsl:otherwise>
        <xsl:apply-templates />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>

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
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:oox="urn:oox" 
  exclude-result-prefixes="w oox">

  <xsl:key name="numId" match="w:num" use="@w:numId"/>
  <xsl:key name="abstractNumId" match="w:abstractNum" use="@w:abstractNumId"/>
  <xsl:key name="numPicBullet" match="w:numPicBullet " use="@w:numPicBulletId"/>

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
    <text:list-style>
      <xsl:attribute name="style:name">
        <xsl:choose>

          <!-- if there is w:lvlOverride, numbering properties can be taken from w:num, and list style must be referred to numId -->
          <xsl:when test="key('numId',$id)/w:lvlOverride">
            <xsl:value-of select="concat('LO',$id)"/>
          </xsl:when>
          
          <!-- math, dialogika: bugfix #1771273 list style is referred to NumId to avoid duplicate list styles BEGIN-->
          <xsl:otherwise>
            <!--<xsl:value-of select="concat('L',@w:abstractNumId)"/>-->
            <xsl:value-of select="concat('L',$id)"/>
          </xsl:otherwise>
          <!-- math, dialogika: bugfix #1771273 list style is referred to NumId to avoid duplicate list styles END-->
          
        </xsl:choose>
      </xsl:attribute>
      <xsl:for-each select="w:lvl">
        <xsl:variable name="level" select="@w:ilvl"/>
        <xsl:choose>

          <!-- when numbering style is overriden, num template is used -->
          <xsl:when test="key('numId',$id)/w:lvlOverride[@w:ilvl = $level]/w:lvl">

            <!--math, dialogika: changed for correct indentation calculation BEGIN -->
            <!--added parameter <xsl:with-param name="numId">-->

            <xsl:apply-templates
              select="key('numId',$id)/w:lvlOverride[@w:ilvl = $level]/w:lvl[@w:ilvl = $level]">
              <xsl:with-param name="numId">
                <xsl:value-of select="$id"/>
              </xsl:with-param>
            </xsl:apply-templates>

            <!--math, dialogika: changed for correct indentation calculation END -->
            
          </xsl:when>

          <xsl:otherwise>
            <xsl:apply-templates select=".">
              <xsl:with-param name="numId">
                <xsl:value-of select="$id"/>
              </xsl:with-param>
            </xsl:apply-templates>  
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
    </text:list-style>
  </xsl:template>

  <xsl:template match="w:lvl">
    <!--math, dialogika: changed for correct indentation calculation BEGIN -->
    <!--added parameter <xsl:param name="numId">-->
    <xsl:param name="numId"/>
    <!--math, dialogika: changed for correct indentation calculation END -->

    <xsl:variable name="lvl">
      <xsl:value-of select="number(@w:ilvl)+1"/>
    </xsl:variable>

    <xsl:choose>
<!-- if numFmt is none, nothing to do -->
      <xsl:when test="w:numFmt/@w:val = 'none' ">
        <text:list-level-style-number text:level="{$lvl}" style:num-format="">
          <xsl:attribute name="style:num-prefix">
            <xsl:call-template name="TextChar">
              <xsl:with-param name="isNumPrefix">true</xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </text:list-level-style-number>
      </xsl:when>

      <!--check if it's numbering, bullet or picture bullet -->
      <xsl:when test="w:numFmt[@w:val = 'bullet'] and w:lvlPicBulletId/@w:val != ''">

        <xsl:variable name="document">
          <xsl:call-template name="GetDocumentName"/>
        </xsl:variable>

        <xsl:variable name="lvlPicBulletId">
          <xsl:value-of select="w:lvlPicBulletId/@w:val"/>
        </xsl:variable>

        <xsl:variable name="rId">
          <xsl:value-of
            select="key('numPicBullet', $lvlPicBulletId)/w:pict/v:shape/v:imagedata/@r:id"/>
        </xsl:variable>

        <xsl:variable name="XlinkHref">
          <xsl:variable name="pzipsource">
            <xsl:value-of
              select="key('Part', concat('word/_rels/',$document,'.rels'))//node()[name() = 'Relationship'][@Id=$rId]/@Target"
            />
          </xsl:variable>
          <xsl:value-of select="concat('Pictures/', substring-after($pzipsource,'/'))"/>
        </xsl:variable>

        <xsl:call-template name="CopyPictures">
          <xsl:with-param name="document">
            <xsl:value-of select="$document"/>
          </xsl:with-param>
          <xsl:with-param name="rId">
            <xsl:value-of select="$rId"/>
          </xsl:with-param>
        </xsl:call-template>

        <text:list-level-style-image>
          <xsl:attribute name="text:level">
            <xsl:value-of select="$lvl"/>
          </xsl:attribute>
          <xsl:attribute name="xlink:href">
            <xsl:value-of select="$XlinkHref"/>
          </xsl:attribute>
          <style:list-level-properties>

            <!--math, dialogika: changed for correct indentation calculation BEGIN -->
            <!--added parameter <xsl:with-param name="numId">-->        
            
            <xsl:call-template name="InsertListLevelProperties">
              <xsl:with-param name="numId">
              <xsl:value-of select="$numId"/>
              </xsl:with-param>
            </xsl:call-template>

            <!--math, dialogika: changed for correct indentation calculation END -->          
            
            <style:text-properties fo:font-size="96"/>
          </style:list-level-properties>
        </text:list-level-style-image>
      </xsl:when>

      <xsl:when test="w:numFmt[@w:val = 'bullet']">
        <text:list-level-style-bullet>
          <xsl:attribute name="text:level">
            <xsl:value-of select="$lvl"/>
          </xsl:attribute>
          <xsl:choose>
            <xsl:when test="w:rPr/w:rStyle">
              <xsl:attribute name="text:style-name">
                <xsl:value-of select="w:rPr/w:rStyle"/>
              </xsl:attribute>
            </xsl:when>
            <!--clam: special style for bullets with symbol font (bug #1806059)-->
            <xsl:when test="w:rPr/w:rFonts/@w:ascii='Symbol'">
              <xsl:attribute name="text:style-name">BulletSymbol</xsl:attribute>
            </xsl:when>
            <xsl:otherwise />              
          </xsl:choose>
          <!--<xsl:if test="w:rPr/w:rStyle">
            <xsl:attribute name="text:style-name">
              <xsl:value-of select="w:rPr/w:rStyle"/>
            </xsl:attribute>
          </xsl:if>-->
          <xsl:attribute name="text:bullet-char">
            <xsl:call-template name="TextChar"/>
          </xsl:attribute>
          <style:list-level-properties>

            <!--math, dialogika: changed for correct indentation calculation BEGIN -->
            <!--added parameter <xsl:with-param name="numId">-->

            <xsl:call-template name="InsertListLevelProperties">
              <xsl:with-param name="numId">
                <xsl:value-of select="$numId"/>
              </xsl:with-param>
            </xsl:call-template>

            <!--math, dialogika: changed for correct indentation calculation END -->            

          </style:list-level-properties>
          <style:text-properties>
            <xsl:for-each select="w:rPr">
              <xsl:call-template name="InsertTextProperties"/>
            </xsl:for-each>
          </style:text-properties>
        </text:list-level-style-bullet>
      </xsl:when>
      <xsl:otherwise>
        <text:list-level-style-number>
          <xsl:attribute name="text:level">
            <xsl:value-of select="$lvl"/>
          </xsl:attribute>
          <xsl:if test="w:rPr/w:rStyle">
            <xsl:attribute name="text:style-name">
              <xsl:value-of select="w:rPr/w:rStyle"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="not(number(substring(w:lvlText/@w:val,string-length(w:lvlText/@w:val))))">
            <xsl:attribute name="style:num-suffix">
              <xsl:value-of select="substring(w:lvlText/@w:val,string-length(w:lvlText/@w:val))"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:call-template name="NumFormat">
            <xsl:with-param name="format" select="w:numFmt/@w:val"/>
            <xsl:with-param name="BeforeAfterNum">
              <xsl:value-of select="w:lvlText/@w:val"/>
            </xsl:with-param>
          </xsl:call-template>
          <xsl:choose>
            <!--dialogika, clam: bugfix #1831298-->
            <xsl:when test="starts-with(w:lvlText/@w:val,concat('%',$lvl,'%',$lvl))">
              <xsl:attribute name="style:num-letter-sync">true</xsl:attribute>
              <xsl:attribute name="text:start-value">27</xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
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
            </xsl:otherwise>
          </xsl:choose>
          
          
          <style:list-level-properties>

            <!--math, dialogika: changed for correct indentation calculation BEGIN -->
            <!--added parameter <xsl:with-param name="numId">-->

            <xsl:call-template name="InsertListLevelProperties">
              <xsl:with-param name="numId">
                <xsl:value-of select="$numId"/>
              </xsl:with-param>
            </xsl:call-template>

            <!--math, dialogika: changed for correct indentation calculation END -->

          </style:list-level-properties>
        </text:list-level-style-number>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- numbering format -->

  <xsl:template name="NumFormat">
    <xsl:param name="format"/>
    <xsl:param name="BeforeAfterNum"/>
    <xsl:attribute name="style:num-format">
      <xsl:if test="$BeforeAfterNum != ''">
        <xsl:choose>
          <xsl:when test="$format= 'decimal' ">1</xsl:when>
          <xsl:when test="$format= 'lowerRoman' ">i</xsl:when>
          <xsl:when test="$format= 'upperRoman' ">I</xsl:when>
          <xsl:when test="$format= 'lowerLetter' ">a</xsl:when>
          <xsl:when test="$format= 'upperLetter' ">A</xsl:when>
          <xsl:otherwise>1</xsl:otherwise>
        </xsl:choose>
      </xsl:if>
    </xsl:attribute>

    <xsl:if test="$BeforeAfterNum != ''">
      <xsl:variable name="NumPrefix">
        <xsl:value-of select="substring-before($BeforeAfterNum, '%')"/>
      </xsl:variable>
      <xsl:if test="$NumPrefix != ''">
        <xsl:attribute name="style:num-prefix">
          <xsl:value-of select="$NumPrefix"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:variable name="NumSuffix">
        <xsl:call-template name="AfterTextNumber">
          <!--math, dialogika: this is wrong:-->
          <!--!!! Parameter name used here is different as parameter name in template -> NumSuffix = NULL always !!!-->          
          <xsl:with-param name="string">
            <xsl:value-of select="$BeforeAfterNum"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:if test="$NumSuffix != ''">
        <xsl:attribute name="style:num-suffix">
          <xsl:value-of select="$NumSuffix"/>
        </xsl:attribute>
      </xsl:if>
    </xsl:if>
  </xsl:template>

  <!-- text after numbering in list -->
 
  <xsl:template name="AfterTextNumber">
    <xsl:param name="BeforeAfterNum"/>
    <xsl:choose>
      <xsl:when test="contains(substring-after($BeforeAfterNum,'%'), '%')">
        <xsl:call-template name="AfterTextNumber">
          <xsl:with-param name="BeforeAfterNum">
            <xsl:value-of select="substring-after($BeforeAfterNum,'%')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="substring(substring-after($BeforeAfterNum, '%'), 2) != ''">
            <xsl:value-of select="substring(substring-after($BeforeAfterNum, '%'), 2)"/>
          </xsl:when>
          <xsl:otherwise/>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <!--math, dialogika: changed for correct indentation calculation BEGIN -->  

  <!-- properties for each list level -->

  <xsl:template name="InsertListLevelProperties">
    <xsl:param name="numId"/>

    <xsl:variable name="Ind" select="w:pPr/w:ind"/>

    <xsl:variable name="ListStyleInd" select="w:pPr/w:ind"/>
    
    <xsl:variable name="tab">
      <xsl:value-of select="w:pPr/w:tabs/w:tab/@w:pos"/>
    </xsl:variable>
    
    <xsl:variable name="abstractNumId">
      <xsl:value-of select="parent::w:abstractNum/@w:abstractNumId"/>
    </xsl:variable>

    <xsl:variable name="ilvl">
      <xsl:value-of select="@w:ilvl"/>
    </xsl:variable>    
    
    <!--<xsl:variable name="numId">
      <xsl:value-of select="../following-sibling::w:num[w:abstractNumId/@w:val=$abstractNumId]/@w:numId"/>
    </xsl:variable>-->

    <xsl:variable name="StyleId">
      <xsl:value-of select="w:pStyle/@w:val"/>
    </xsl:variable>

    <xsl:variable name="paragraph_ref_this_list_level"
      select="key('Part', 'word/document.xml')/w:document/w:body/w:p[w:pPr/w:numPr/w:numId/@w:val=$numId and w:pPr/w:numPr/w:ilvl/@w:val=$ilvl][1]"/>

    <xsl:variable name="paragraph_ref_this_styleid"
      select="key('Part', 'word/document.xml')/w:document/w:body/w:p[w:pPr/w:pStyle/@w:val=$StyleId][1]"/>    
    
    <xsl:variable name="style"
      select="key('StyleId', $StyleId)[1]"/>

    <xsl:variable name="Hanging">
      <xsl:choose>        
        <xsl:when test="$paragraph_ref_this_list_level">
          <!--Paragraph directly referencing this numID exists-->
          <xsl:choose>            
            <xsl:when test="$paragraph_ref_this_list_level/w:pPr/w:ind/@w:hanging">
              <!--Paragraph has direct formatting -> take this value-->
              <xsl:value-of select="$paragraph_ref_this_list_level/w:pPr/w:ind/@w:hanging"/>
            </xsl:when>

            <!--FIRSTLINE: Paragraph has direct formatting firstLine -> firstLine counts-->
            <xsl:when test="$paragraph_ref_this_list_level/w:pPr/w:ind/@w:firstLine">0</xsl:when>
            
            <xsl:otherwise>
              <!--Paragraph has *NO* direct formatting-->
              <xsl:choose>
                <xsl:when test="$ListStyleInd/@w:hanging">
                  <!--take list style value-->
                  <xsl:value-of select="$ListStyleInd/@w:hanging"/>
                </xsl:when>
              <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>                      
          </xsl:choose>                    
        </xsl:when>

        <!--Paragraph directly referencing this numID does *NOT* exist-->

        <xsl:when test="$paragraph_ref_this_styleid/w:pPr/w:ind/@w:hanging">
          <!-- Use direct formatting of paragraph referencing this paragraph style-->
          <xsl:value-of select="$paragraph_ref_this_styleid/w:pPr/w:ind/@w:hanging"/>
        </xsl:when>

        <!--FIRSTLINE: direct formatting of paragraph referencing this paragraph style has firstLine -> firstLine counts-->
        <xsl:when test="$paragraph_ref_this_styleid/w:pPr/w:ind/@w:firstLine">0</xsl:when>
          
        <xsl:when test="$style/w:pPr/w:ind/@w:hanging">
          <!-- Use paragraph style value-->
          <xsl:value-of select="$style/w:pPr/w:ind/@w:hanging"/>
        </xsl:when>

        <!--FIRSTLINE: Paragraph style has firstLine -> firstLine counts-->        
        <xsl:when test="$style/w:pPr/w:ind/@w:firstLine">0</xsl:when>
        
        <xsl:when test="w:pPr/w:ind/@w:hanging">
          <!-- Use list style value-->
          <xsl:value-of select="w:pPr/w:ind/@w:hanging"/>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>


    <xsl:variable name="Left">
      <xsl:choose>
        <xsl:when test="$paragraph_ref_this_list_level">
          <!--Paragraph directly referencing this numID exists-->
          <xsl:choose>
            <xsl:when test="$paragraph_ref_this_list_level/w:pPr/w:ind/@w:left">
              <!--Paragraph has direct formatting -> take this value-->
              <xsl:value-of select="$paragraph_ref_this_list_level/w:pPr/w:ind/@w:left"/>
            </xsl:when>
            <xsl:otherwise>
              <!--Paragraph has *NO* direct formatting-->
              <xsl:choose>
                <xsl:when test="$ListStyleInd/@w:left">
                  <!--take list style value-->
                  <xsl:value-of select="$ListStyleInd/@w:left"/>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>

        <!--Paragraph directly referencing this numID does *NOT* exist-->

        <xsl:when test="$paragraph_ref_this_styleid/w:pPr/w:ind/@w:left">
          <!-- Use direct formatting of paragraph referencing this paragraph style-->
          <xsl:value-of select="$paragraph_ref_this_styleid/w:pPr/w:ind/@w:left"/>
        </xsl:when>        
        
        <xsl:when test="$style/w:pPr/w:ind/@w:left">
          <!-- Use paragraph style value-->
          <xsl:value-of select="$style/w:pPr/w:ind/@w:left"/>
        </xsl:when>
        <xsl:when test="w:pPr/w:ind/@w:left">
          <!-- Use list style value-->
          <xsl:value-of select="w:pPr/w:ind/@w:left"/>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>


    <xsl:variable name="FirstLine">
      <xsl:choose>
        <xsl:when test="$paragraph_ref_this_list_level">
          <!--Paragraph directly referencing this numID exists-->
          <xsl:choose>
            <xsl:when test="$paragraph_ref_this_list_level/w:pPr/w:ind/@w:firstLine">
              <!--Paragraph has direct formatting -> take this value-->
              <xsl:value-of select="$paragraph_ref_this_list_level/w:pPr/w:ind/@w:firstLine"/>
            </xsl:when>
            <xsl:otherwise>
              <!--Paragraph has *NO* direct formatting-->
              <xsl:choose>
                <xsl:when test="$ListStyleInd/@w:firstLine">
                  <!--take list style value-->
                  <xsl:value-of select="$ListStyleInd/@w:firstLine"/>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>

        <!--Paragraph directly referencing this numID does *NOT* exist-->

        <xsl:when test="$paragraph_ref_this_styleid/w:pPr/w:ind/@w:firstLine">
          <!-- Use direct formatting of paragraph referencing this paragraph style-->
          <xsl:value-of select="$paragraph_ref_this_styleid/w:pPr/w:ind/@w:firstLine"/>
        </xsl:when>

        <xsl:when test="$style/w:pPr/w:ind/@w:firstLine">
          <!-- Use paragraph style value-->
          <xsl:value-of select="$style/w:pPr/w:ind/@w:firstLine"/>
        </xsl:when>
        <xsl:when test="w:pPr/w:ind/@w:firstLine">
          <!-- Use list style value-->
          <xsl:value-of select="w:pPr/w:ind/@w:firstLine"/>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>    
    


    <!--<xsl:variable name="WLeft">
      <xsl:choose>
        <xsl:when test="$Ind/@w:left">
          <xsl:value-of select="$Ind/@w:left"/>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>-->    

    <!--<xsl:variable name="WLeft">
      <xsl:choose>
        <xsl:when test="$paragraph/w:pPr/w:ind/@w:left">
          <xsl:value-of select="$paragraph/w:pPr/w:ind/@w:left"/>
        </xsl:when>
        <xsl:when test="$style/w:pPr/w:ind/@w:left">
          <xsl:value-of select="$style/w:pPr/w:ind/@w:left"/>
        </xsl:when>
        <xsl:when test="w:pPr/w:ind/@w:left">
          <xsl:value-of select="w:pPr/w:ind/@w:left"/>
        </xsl:when>        
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>-->

    <!--<xsl:variable name="WHanging">
      <xsl:choose>
        <xsl:when test="$Ind/@w:hanging">
          <xsl:value-of select="$Ind/@w:hanging"/>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>-->

    <!--xsl:variable name="WFirstLine">
      <xsl:call-template name="FirstLine"/>
    </xsl:variable>-->
    

    <xsl:variable name="tabs"
      select="$paragraph_ref_this_list_level/w:pPr/w:tabs | $paragraph_ref_this_styleid/w:pPr/w:tabs | $style/w:pPr/w:tabs | w:pPr/w:tabs" />

    <xsl:variable name="SpaceToNextTab">

      <xsl:variable name="MinTabOffset">
        <xsl:value-of select ="350"/>
      </xsl:variable>      
      
      <xsl:choose>
        <xsl:when test="w:suff/@w:val='nothing'">0</xsl:when>
        <xsl:when test="w:suff/@w:val='space'">350</xsl:when>
        <xsl:otherwise>  
          
        <xsl:variable name="MinRelevantCustomTab">
          <xsl:choose>
            <xsl:when test="$Hanging != '0'">
              <!--hanging -> get min custom tab that is between start of first line and start of paragrah text-->
              <xsl:call-template name="GetMinVal">
                <xsl:with-param name="values" select="$tabs/w:tab[@w:pos &gt; ( $Left - $Hanging + $MinTabOffset) and @w:pos &lt; $Left ]/@w:pos" />
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>

              <!--no hanging -> get min custom tab that is bigger than the start of the first line-->
              <xsl:call-template name="GetMinVal">
                <xsl:with-param name="values" select="$tabs/w:tab[@w:pos &gt; ( $Left + $FirstLine + $MinTabOffset)]/@w:pos" />
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>      
        
        <xsl:choose>
          <xsl:when test="$Hanging != '0'">
            
            <!--hanging-->

            <!--<xsl:value-of select="$Hanging" />-->

            <xsl:choose>
              <xsl:when test="$MinRelevantCustomTab != 'NaN'">
                <!--take min relevant custom tab -->
                <xsl:value-of select="$Hanging - ($Left - $MinRelevantCustomTab)" />
              </xsl:when>
              <xsl:otherwise>
                <!--take hanging value-->
                <xsl:value-of select="$Hanging" />
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>

          <xsl:otherwise>

            <!--no hanging-->
            
            <xsl:variable name="DefaultTab">
              <xsl:value-of select="key('Part', 'word/settings.xml')/w:settings/w:defaultTabStop/@w:val"/>
            </xsl:variable>
            <xsl:variable name="NextDefaultTabPos">
              <xsl:value-of select="(floor(($Left + $FirstLine + $MinTabOffset) div $DefaultTab) + 1) * $DefaultTab"/>
            </xsl:variable>

            <xsl:choose>
              <xsl:when test="$MinRelevantCustomTab != 'NaN'">
                <!--take min relevant custom tab-->
                <xsl:value-of select="$MinRelevantCustomTab - ($Left + $FirstLine)"/>
              </xsl:when>

              <xsl:otherwise>
                <!--take next default tab-->
                <xsl:value-of select="$NextDefaultTabPos - ($Left + $FirstLine)"/>
              </xsl:otherwise>
            </xsl:choose>
            
          </xsl:otherwise>
        </xsl:choose>

        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
      
    
    
    <!--text:space-before-->

    <xsl:attribute name="text:space-before">
      <xsl:call-template name="ConvertTwips">
        <xsl:with-param name="length">
           <xsl:value-of select="$Left - $SpaceToNextTab"/>
        </xsl:with-param>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>

    <!--text:min-label-width-->

    <xsl:attribute name="text:min-label-width">
      <xsl:call-template name="ConvertTwips">
        <xsl:with-param name="length">
          <!--<xsl:choose>
            <xsl:when test="w:suff/@w:val='nothing'">0</xsl:when>
            <xsl:when test="w:suff/@w:val='space'">350</xsl:when>
            <xsl:otherwise>-->
              <xsl:value-of select="$SpaceToNextTab"/>
            <!--</xsl:otherwise>
          </xsl:choose>-->
        </xsl:with-param>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>

    <!--<xsl:choose>
      <xsl:when test="$Ind/@w:hanging">
        <xsl:attribute name="text:space-before">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:choose>
                <xsl:when test="$WFirstLine = 'NaN'">
                  <xsl:value-of select="$WLeft - $WHanging"/>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
        <xsl:attribute name="text:min-label-width">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:choose>
                <xsl:when test="$WFirstLine = 'NaN'">
                  <xsl:choose>
                    <xsl:when
                      test="w:suff/@w:val='nothing' or ($paragraph/w:pPr/w:ind/@w:left = 0 and $tab = '')"
                      >0</xsl:when>
                    <xsl:when test="w:suff/@w:val='space'">350</xsl:when>
                    <xsl:when test="$tab != '' and (number($tab) - number($WLeft) >= 0) and not(number($WLeft) = number($WHanging))">
                      <xsl:value-of select="number($tab) - number($WLeft) + number($WHanging)"/>
                    </xsl:when>
                    <xsl:when test="$paragraph/w:pPr/w:ind/@w:hanging">
                      <xsl:value-of select="$paragraph/w:pPr/w:ind/@w:hanging"/>
                    </xsl:when>
                    <xsl:when test="$style/w:pPr/w:ind/@w:hanging">
                      <xsl:value-of select="$style/w:pPr/w:ind/@w:hanging"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$Ind/@w:hanging"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
		<xsl:when test="$WFirstLine = '0'"><xsl:value-of
                  select="/oox:package/oox:part[@oox:name='word/settings.xml']/w:settings/w:defaultTabStop/@w:val"/></xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
        <xsl:attribute name="text:min-label-distance">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:choose>
                <xsl:when test="$WFirstLine != 'NaN'">
                  <xsl:choose>
                    <xsl:when test="$tab != '' and $tab - $WFirstLine > 0 ">
                      <xsl:value-of select="$tab - $WFirstLine"/>
                    </xsl:when>
                    <xsl:otherwise>0</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="text:space-before">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:value-of select="number($WLeft)"/>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
        <xsl:attribute name="text:min-label-width">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:choose>
                <xsl:when test="$WFirstLine = 'NaN'">
                  <xsl:choose>
                    <xsl:when
                      test="../w:multiLevelType/@w:val='multilevel' and number($tab) > number($WLeft)">
                      <xsl:value-of select="$tab - number($WLeft)"/>
                    </xsl:when>
                    <xsl:otherwise>0--><!--<xsl:value-of
                      select="document('word/settings.xml')/w:settings/w:defaultTabStop/@w:val"/> placed in comment to obtain right space between number and text and replaced by 0--><!--</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:when test="$WFirstLine = '0'">0--><!--<xsl:value-of
                  select="document('word/settings.xml')/w:settings/w:defaultTabStop/@w:val"/> placed in comment to obtain right space between number and text and replaced by 0--><!--</xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
        <xsl:attribute name="text:min-label-distance">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:choose>
                <xsl:when test="$tab != '' and $WFirstLine != 'NaN'">
                  <xsl:value-of select="$tab - $WFirstLine"/>
                </xsl:when>
                <xsl:when test="$WFirstLine != 'NaN'">0</xsl:when>
                <xsl:when test="(3 * number($WFirstLine)) &lt; (number($tab) - number($WLeft)) ">
                  <xsl:value-of select="number($tab) - number($WLeft) - (2 * number($WFirstLine))"/>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>-->

    <!--math, dialogika: changed for correct indentation calculation END -->    
    
    <!-- Picture Bullet Size -->
    <xsl:if test="w:lvlPicBulletId/@w:val != ''">
      <xsl:attribute name="fo:width">
        <xsl:call-template name="ConvertPoints">
          <xsl:with-param name="length">
            <xsl:choose>
              <xsl:when test="w:rPr/w:sz/@w:val">
                <xsl:value-of select="w:rPr/w:sz/@w:val div 2"/>
              </xsl:when>
              <xsl:otherwise>12</xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="fo:height">
        <xsl:call-template name="ConvertPoints">
          <xsl:with-param name="length">
            <xsl:choose>
              <xsl:when test="w:rPr/w:sz/@w:val">
                <xsl:value-of select="w:rPr/w:sz/@w:val div 2"/>
              </xsl:when>
              <xsl:otherwise>12</xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>

    <!-- Bullet Align -->
    <xsl:attribute name="fo:text-align">
      <xsl:choose>
        <xsl:when test="w:lvlJc/@w:val='center'">
          <xsl:text>center</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:text>left</xsl:text>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>

  <!--math, dialogika: Added template for getting min value of a node-set BEGIN -->  
  
  <xsl:template name="GetMinVal">
    <xsl:param name="values"/>
    <xsl:choose>
      <xsl:when test="$values">
        <xsl:for-each select="$values">
          <xsl:sort data-type="number" order="ascending"/>
          <xsl:if test="position()=1">
            <xsl:value-of select="."/>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:otherwise>NaN</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--math, dialogika: Added template for getting min value of a node-set END -->  
  
  <!-- ➢types of bullets -->

  <xsl:template name="TextChar">
    <xsl:param name="isNumPrefix"/>
    <xsl:choose>
      <xsl:when test="w:lvlText[@w:val = ''] "></xsl:when>
      <xsl:when test="w:lvlText[@w:val = '' ]"></xsl:when>
      <xsl:when test="w:lvlText[@w:val = '' ]">☑</xsl:when>
      <xsl:when test="w:lvlText[@w:val = '•' ]">•</xsl:when>
      <!--<xsl:when test="w:lvlText[@w:val = '' ]">●</xsl:when>-->
      <xsl:when test="w:lvlText[@w:val = '' ]">➢</xsl:when>
      <!--<xsl:when test="w:lvlText[@w:val = '➢' ]">➢</xsl:when>-->
      <xsl:when test="w:lvlText[@w:val = '' ]">✔</xsl:when>
      <xsl:when test="w:lvlText[@w:val = '' ]">■</xsl:when>
      <xsl:when test="w:lvlText[@w:val = 'o' ]">○</xsl:when>
      <xsl:when test="w:lvlText[@w:val = '']">➔</xsl:when>
      <xsl:when test="w:lvlText[@w:val = '' ]">✗</xsl:when>
      <xsl:when test="w:lvlText[@w:val = '-' ]">–</xsl:when>
      <xsl:when test="w:lvlText[@w:val = '–' ]">–</xsl:when>
      <xsl:when test="w:lvlText[@w:val = '' ]">–</xsl:when>
      <xsl:when test="w:lvlText[@w:val = '' ]"></xsl:when>
      <xsl:when test="$isNumPrefix and $isNumPrefix='true'">
        <xsl:value-of select="w:lvlText/@w:val"/>
      </xsl:when>
      <xsl:otherwise><xsl:value-of select="w:lvlText/@w:val"/></xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- checks for list numPr properties (numid or level) in styles hierarchy -->
  <xsl:template name="GetListStyleProperty">
    <xsl:param name="style"/>
    <xsl:param name="property"/>

    <xsl:choose>
      <xsl:when test="$style/w:pPr/w:numPr and $property = 'w:ilvl' ">
        <xsl:choose>
          <xsl:when test="$style/w:pPr/w:numPr/w:ilvl">
            <xsl:value-of select="$style/w:pPr/w:numPr/w:ilvl/@w:val"/>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="$style/w:pPr/w:numPr/*[name() = $property]/@w:val">
        <xsl:value-of select="$style/w:pPr/w:numPr/*[name() = $property]/@w:val"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="$style/w:basedOn">
          <xsl:variable name="parentStyle" select="$style/w:basedOn/@w:val"/>
          <xsl:for-each select="key('Part', 'word/styles.xml')">
            <xsl:call-template name="GetListStyleProperty">
              <xsl:with-param name="style" select="key('StyleId', $parentStyle)"/>
              <xsl:with-param name="property" select="$property"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- checks for list numPr properties (numid or level) for given element  -->
  <xsl:template name="GetListProperty">
    <xsl:param name="node"/>
    <xsl:param name="property"/>

    <!-- second condition in xsl:when checks if there isn't another w:p node before decendant -->
    <xsl:choose>
      <xsl:when
        test="$node/descendant::w:numPr[not(ancestor::w:pPrChange)] and $node/descendant::w:numPr/child::node()[name() = $property] and
        generate-id($node)=generate-id($node/descendant::w:numPr[not(ancestor::w:pPrChange)]/ancestor::w:p[1])">
        <xsl:value-of select="$node/descendant::w:numPr/child::node()[name() = $property]/@w:val"/>
      </xsl:when>

      <xsl:when
        test="$node/descendant::w:pStyle[not(ancestor::w:pPrChange)] and 
        generate-id($node)=generate-id($node/descendant::w:pStyle[not(ancestor::w:pPrChange)]/ancestor::w:p[1])">
        <xsl:variable name="styleId" select="$node/descendant::w:pStyle/@w:val"/>

        <xsl:variable name="pStyle"
          select="key('StyleId', $styleId)"/>
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


  <!-- paragraph which is the first element in list level-->
  <xsl:template match="w:p" mode="list">
    <xsl:param name="numId"/>
    <xsl:param name="level"/>
    <xsl:param name="listLevel"/>
    <xsl:param name="dummyOutlineLevel"/>
    <xsl:param name="isPrecedingHeading"/>
    <xsl:variable name="isFirstListItem">
      <xsl:call-template name="IsFirstListItem">
        <xsl:with-param name="node" select="."/>
        <xsl:with-param name="numId" select="$numId"/>
        <xsl:with-param name="listLevel" select="$listLevel"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="listNum">
      <xsl:choose>

        <!-- if there is w:lvlOverride, numbering properties can be taken from w:num, and list style must be referred to numId -->
        <xsl:when test="key('numId', $numId)/w:lvlOverride">
          <xsl:value-of select="concat('LO',$numId)"/>
        </xsl:when>

        <xsl:when
          test="key('abstractNumId', key('numId', $numId)/w:abstractNumId/@w:val)/w:numStyleLink">
          <xsl:variable name="linkedStyle">
            <xsl:value-of
              select="key('abstractNumId', key('numId', $numId)/w:abstractNumId/@w:val)/w:numStyleLink/@w:val"
            />
          </xsl:variable>
          <xsl:variable name="linkedNumId">
            <xsl:value-of select="key('StyleId', $linkedStyle)/w:pPr/w:numPr/w:numId/@w:val" />
          </xsl:variable>

          <!-- math, dialogika: bugfix #1771273 list style is referred to NumId to avoid duplicate list styles BEGIN-->          
          
          <!--<xsl:value-of select="concat('L',key('numId', $linkedNumId)/w:abstractNumId/@w:val)" />-->
          <xsl:value-of select="concat('L',$linkedNumId)" />-->
        </xsl:when>

        <xsl:otherwise>
          <!--<xsl:value-of select="concat('L',key('numId', $numId)/w:abstractNumId/@w:val)" />-->
          <xsl:value-of select="concat('L',$numId)" />
        </xsl:otherwise>
        <!-- math, dialogika: bugfix #1771273 list style is referred to NumId to avoid duplicate list styles END-->
        
      </xsl:choose>
    </xsl:variable>

    <!-- math, dialogika: bugfix #1841780 START-->
        
    <xsl:variable name="StyleId">
       <xsl:value-of select="w:pPr/w:pStyle/@w:val" />
    </xsl:variable>

    <xsl:variable name="IsDefaultHeading">
      <xsl:call-template name ="CheckDefaultHeading">
        <xsl:with-param name="Name">
          <xsl:value-of select="key('StyleId',$StyleId)/w:name/@w:val" />
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!--<xsl:if test="$isFirstListItem = 'true' or ($dummyOutlineLevel!='' and $isPrecedingHeading='true')">-->
    <xsl:if test="$isFirstListItem = 'true' or ($dummyOutlineLevel!='' and $isPrecedingHeading='true') or $IsDefaultHeading='true'">

    <!-- math, dialogika: bugfix #1841780 END-->
      
          <text:list text:style-name="{$listNum}">
            
            <!--  TODO - continue numbering-->
            <xsl:attribute name="text:continue-numbering">true</xsl:attribute>
            
            <xsl:variable name="currentListLevel">
              <xsl:call-template name="GetGeneratedListLevel">
                <xsl:with-param name="listLevel" select="$listLevel"/>
              </xsl:call-template>
            </xsl:variable>
            
            <xsl:call-template name="InsertListLevel">
              <xsl:with-param name="level" select="$level"/>
              <xsl:with-param name="numId" select="$numId"/>
              <xsl:with-param name="listLevel" select="$currentListLevel"/>
              <xsl:with-param name="dummyOutlineLevel" select="$dummyOutlineLevel"/>
              <!--math, dialogika: unused parameter-->
              <!--<xsl:with-param name="isPrecedingHeading" select="$isPrecedingHeading"/>-->
            </xsl:call-template>
          </text:list>
        
    </xsl:if>
  </xsl:template>

  <!--gets actually generated list level-->
  <xsl:template name="GetGeneratedListLevel">
    <xsl:param name="listLevel"/>
    <xsl:choose>
      <xsl:when test="$listLevel = ''">
        <xsl:text>0</xsl:text>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$listLevel"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- converts element as list item and insert nested level if there is any -->
  <xsl:template name="InsertListLevel">
    <xsl:param name="node" select="."/>
    <xsl:param name="level"/>
    <xsl:param name="numId"/>
    <xsl:param name="listLevel"/>
    <xsl:param name="dummyOutlineLevel"/>
    <!--math, dialogika: unused parameter-->
    <!--<xsl:param name="isPrecedingHeading"/>-->

    <xsl:variable name="isNestedList">
      <xsl:call-template name="IsNestedList">
        <xsl:with-param name="node" select="$node"/>
        <xsl:with-param name="level" select="$level"/>
        <xsl:with-param name="listLevel" select="$listLevel"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:apply-templates select="$node" mode="list-item">
      <xsl:with-param name="numId" select="$numId"/>
      <xsl:with-param name="level" select="$level"/>
      <xsl:with-param name="listLevel" select="$listLevel"/>
      <xsl:with-param name="isNestedList" select="$isNestedList"/>
    </xsl:apply-templates>
    
    <xsl:choose>
      <!-- don't put  numbered heading into list-->
      <xsl:when test="number($listLevel)=0 and number($level) &gt; 0 and $dummyOutlineLevel!=''"/>
      <xsl:when test="$isNestedList = 'true'">
        <xsl:call-template name="InsertCurrentLevel">
          <xsl:with-param name="node" select="$node"/>
          <xsl:with-param name="level" select="$listLevel"/>
          <xsl:with-param name="numId" select="$numId"/>
          <xsl:with-param name="listLevel" select="$listLevel"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="InsertFollowingLevel">
          <xsl:with-param name="node" select="$node"/>
          <xsl:with-param name="level" select="$listLevel"/>
          <xsl:with-param name="numId" select="$numId"/>
          <xsl:with-param name="listLevel" select="$listLevel"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- insert lists element content and nested list if there is following element is on differeynt level -->
  <xsl:template match="w:p" mode="list-item">
    <xsl:param name="numId"/>
    <xsl:param name="level"/>
    <xsl:param name="listLevel"/>
    <xsl:param name="isNestedList"/>

    <xsl:choose>
      <!--      if this paragraph is attached to preceding in track changes mode-->
      <xsl:when test="key('p', number(@oox:id)-1)/w:pPr/w:rPr/w:del"/>
           
      <xsl:otherwise>
        <text:list-item>
          <xsl:variable name="followingParagraph" select="self::node()[1]/following-sibling::w:p[1]"/>

          <xsl:choose>
            <xsl:when test="$listLevel = $level">

              <xsl:call-template name="InsertListItemContent">
                <xsl:with-param name="node" select="self::node()[1]"/>
              </xsl:call-template>

              <xsl:if test="$isNestedList = 'true' ">
                <xsl:variable name="followingLevel">
                  <xsl:call-template name="GetListProperty">
                    <xsl:with-param name="node" select="$followingParagraph"/>
                    <xsl:with-param name="property">w:ilvl</xsl:with-param>
                  </xsl:call-template>
                </xsl:variable>

                <xsl:apply-templates select="$followingParagraph" mode="list">
                  <xsl:with-param name="numId" select="$numId"/>
                  <xsl:with-param name="level" select="$followingLevel"/>
                  <xsl:with-param name="listLevel" select="$listLevel + 1"/>
                </xsl:apply-templates>
              </xsl:if>
            </xsl:when>

            <xsl:when test="$isNestedList = 'true' ">
              <xsl:apply-templates select="self::node()" mode="list">
                <xsl:with-param name="numId" select="$numId"/>
                <xsl:with-param name="level" select="$level"/>
                <xsl:with-param name="listLevel" select="$listLevel + 1"/>
              </xsl:apply-templates>
            </xsl:when>
          </xsl:choose>
        </text:list-item>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertListItemContent">
    <xsl:param name="node"/>

    <xsl:variable name="outlineLevel">
      <xsl:call-template name="GetOutlineLevel">
        <xsl:with-param name="node" select="$node"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="$outlineLevel != ''">
        <xsl:apply-templates select="$node" mode="heading"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates select="$node" mode="paragraph"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--inserts following list item if its level is higher-->
  <xsl:template name="InsertFollowingLevel">
    <xsl:param name="node"/>
    <xsl:param name="numId"/>
    <xsl:param name="level"/>
    <xsl:param name="listLevel"/>

    <xsl:variable name="followingParagraph" select="$node/following-sibling::w:p[1]"/>

    <xsl:variable name="followingNumId">
      <xsl:call-template name="GetListProperty">
        <xsl:with-param name="node" select="$followingParagraph"/>
        <xsl:with-param name="property">w:numId</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="followingLevel">
      <xsl:call-template name="GetListProperty">
        <xsl:with-param name="node" select="$followingParagraph"/>
        <xsl:with-param name="property">w:ilvl</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="isHigherLevel">
      <xsl:call-template name="isHigherLevelItem">
        <xsl:with-param name="numId" select="$numId"/>
        <xsl:with-param name="followingNumId" select="$followingNumId"/>
        <xsl:with-param name="level" select="$level"/>
        <xsl:with-param name="followingLevel" select="$followingLevel"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$isHigherLevel = 'true'">
      <xsl:call-template name="InsertListLevel">
        <xsl:with-param name="node" select="$followingParagraph"/>
        <xsl:with-param name="level" select="$followingLevel"/>
        <xsl:with-param name="numId" select="$followingNumId"/>
        <xsl:with-param name="listLevel" select="$listLevel"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

  <!--  inserts the remaining items on given level when following item is nested list element-->
  <xsl:template name="InsertCurrentLevel">
    <xsl:param name="node"/>
    <xsl:param name="numId"/>
    <xsl:param name="level"/>
    <xsl:param name="listLevel"/>

    <xsl:variable name="followingParagraphId">
      <xsl:call-template name="GetCurrentLevelItem">
        <xsl:with-param name="node" select="$node"/>
        <xsl:with-param name="level" select="$listLevel"/>
        <xsl:with-param name="numId" select="$numId"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$followingParagraphId != ''">

      <xsl:variable name="followingParagraphOnLevel"
        select="$node/following-sibling::w:p[generate-id() = $followingParagraphId]"/>

      <!--clam: bugfix #1788550
      the list has to stop if the following paragraph is a numbered heading
      -->
      <xsl:if test="not(contains($followingParagraphOnLevel/w:pPr/w:pStyle/@w:val,'Heading') and not(contains(./w:pPr/w:pStyle/@w:val,'Heading')))">
        <xsl:call-template name="InsertListLevel">
          <xsl:with-param name="node" select="$followingParagraphOnLevel"/>
          <xsl:with-param name="level" select="$listLevel"/>
          <xsl:with-param name="numId" select="$numId"/>
          <xsl:with-param name="listLevel" select="$listLevel"/>
        </xsl:call-template>
      </xsl:if>
      
    </xsl:if>
  </xsl:template>

  <!--checks list element level-->
  <xsl:template name="isHigherLevelItem">
    <xsl:param name="numId"/>
    <xsl:param name="followingNumId"/>
    <xsl:param name="level"/>
    <xsl:param name="followingLevel"/>

    <xsl:choose>
      <xsl:when test="$followingLevel &gt;= $level and $numId = $followingNumId">
        <xsl:text>true</xsl:text>
      </xsl:when>
      <xsl:otherwise>
        <xsl:text>false</xsl:text>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--  ckecks element is first on given list level-->
  <xsl:template name="IsFirstListItem">
    <xsl:param name="node"/>
    <xsl:param name="numId"/>
    <xsl:param name="listLevel"/>

    <xsl:choose>
      <xsl:when test="$listLevel != '' ">
        <xsl:text>true</xsl:text>
      </xsl:when>
      <xsl:when test="$node/preceding-sibling::node()[1]/w:pPr/w:numPr">
        <xsl:choose>
          <xsl:when test="$node/preceding-sibling::node()[1]/w:pPr/w:numPr/w:numId/@w:val = $numId">
            <xsl:text>false</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>true</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="pStyle"
          select="key('StyleId', $node/preceding-sibling::node()[1]/w:pPr/w:pStyle/@w:val)"/>
        <xsl:variable name="precedingNumId">
          <xsl:call-template name="GetListProperty">
            <xsl:with-param name="node" select="$node/preceding-sibling::node()[1]"/>
            <xsl:with-param name="property">w:numId</xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="$precedingNumId = $numId">
            <xsl:text>false</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>true</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- checks if given element should be generated as nested list)....-->
  <xsl:template name="IsNestedList">
    <xsl:param name="node"/>
    <xsl:param name="level"/>
    <xsl:param name="listLevel"/>

    <xsl:for-each select="$node">
      <xsl:if test="$level &gt; 0 and $level != $listLevel">
        <xsl:text>true</xsl:text>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

  <!-- searches for following list item on given level -->
  <xsl:template name="GetCurrentLevelItem">
    <xsl:param name="node"/>
    <xsl:param name="level"/>
    <xsl:param name="numId"/>

    <xsl:variable name="nodePosition">
      <xsl:for-each select="$node">
        <xsl:value-of select="position()"/>
      </xsl:for-each>
    </xsl:variable>

    <xsl:variable name="followingParagraph" select="$node/following-sibling::w:p[1]"/>

    <xsl:variable name="followingNumId">
      <xsl:call-template name="GetListProperty">
        <xsl:with-param name="node" select="$followingParagraph"/>
        <xsl:with-param name="property">w:numId</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$followingNumId = $numId">
      <xsl:variable name="followingLevel">
        <xsl:call-template name="GetListProperty">
          <xsl:with-param name="node" select="$followingParagraph"/>
          <xsl:with-param name="property">w:ilvl</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>

      <xsl:choose>
        <xsl:when test="$followingLevel = $level">
          <xsl:value-of select="generate-id($followingParagraph)"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:if test="$followingLevel &gt;= $level">
            <xsl:call-template name="GetCurrentLevelItem">
              <xsl:with-param name="node" select="$followingParagraph"/>
              <xsl:with-param name="level" select="$level"/>
              <xsl:with-param name="numId" select="$numId"/>
            </xsl:call-template>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
  </xsl:template>

  <!-- inserts automatic list styles with empty num format for elements which has non-existent w:num attached -->
  <xsl:template name="InsertDefaultListStyle">
    <xsl:param name="numId"/>

    <!-- list style with empty num format must be referred to numId, because there is no abstractNumId -->
    <xsl:variable name="listName" select="concat('LO',$numId)"/>
    <text:list-style style:name="{$listName}">
      <xsl:call-template name="InsertDefaultLevelProperties">
        <xsl:with-param name="levelNum" select="'1'"/>
      </xsl:call-template>
    </text:list-style>
  </xsl:template>

  <!-- inserts level with empty num format for elements which has non-existent w:num attached -->
  <xsl:template name="InsertDefaultLevelProperties">
    <xsl:param name="levelNum"/>
    <text:list-level-style-number text:level="{$levelNum}" style:num-format=""/>
    <xsl:if test="$levelNum &lt; 9">
      <xsl:call-template name="InsertDefaultLevelProperties">
        <xsl:with-param name="levelNum" select="$levelNum+1"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

  <!-- inserts automatic list styles with empty num format for elements which has non-existent w:num attached -->
  <xsl:template match="w:numId" mode="automaticstyles">
    <xsl:variable name="numId" select="@w:val"/>
    <xsl:for-each select="key('Part', 'word/numbering.xml')">
      <xsl:if test="key('numId',@w:val) = ''">
        <xsl:call-template name="InsertDefaultListStyle">
          <xsl:with-param name="numId" select="$numId"/>
        </xsl:call-template>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  
  
  <!--  heading numbering style -->
  <xsl:template name="InsertOutlineListStyle">
    <xsl:param name="numid"/>
    <text:outline-style>
      <!--create outline level style only for styles which have outlineLvl and numId what means that list level is linked to a style -->
      <xsl:for-each
        select="key('Part', 'word/styles.xml')/w:styles/w:style[child::w:pPr/w:outlineLvl and child::w:pPr/w:numPr and contains(@w:styleId,'Heading')]">
        <xsl:variable name="styleId">
          <xsl:value-of select="@w:styleId"/>
        </xsl:variable>

          <xsl:variable name="numId">
            <xsl:choose>
              <xsl:when test="w:pPr/w:numPr/w:numId/@w:val">
                <xsl:value-of select="w:pPr/w:numPr/w:numId/@w:val"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:call-template name="GetListStyleProperty">
                  <xsl:with-param name="style" select="."/>
                  <xsl:with-param name="property">w:numId</xsl:with-param>
                </xsl:call-template>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="levelId" select="w:pPr/w:numPr/w:ilvl/@w:val | w:pPr/w:outlineLvl/@w:val"/>
          <text:outline-level-style>
            <xsl:attribute name="text:level">
              <xsl:value-of select="./w:pPr/w:outlineLvl/@w:val +1"/>
            </xsl:attribute>
            <xsl:for-each select="key('Part', 'word/numbering.xml')">
              <xsl:variable name="abstractNum" select="key('abstractNumId',key('numId',$numId)/w:abstractNumId/@w:val)"/>
              <!--w:lvl shows which level defintion should be taken from abstract num-->
              <xsl:for-each select="$abstractNum/w:lvl[@w:ilvl = $levelId][1]">
                <xsl:if
                  test="not(number(substring(./w:lvlText/@w:val,string-length(./w:lvlText/@w:val)))) and ./w:lvlText/@w:val != 'nothing'">
                  <xsl:attribute name="style:num-suffix">
                    <xsl:variable name="suffix">
                      <xsl:value-of select="substring(./w:lvlText/@w:val,string-length(./w:lvlText/@w:val))"
                      />
                    </xsl:variable>
                    <xsl:choose>
                      <xsl:when test="w:suff/@w:val='space'">
                        <xsl:value-of select="concat($suffix,' ')"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <!--math, dialogika: test to improve appearance of TOC in ODT ???-->
                        <!--TODO: add a check whether space is already at the end of the suffix-->
                        <!--to avoid that roundtrip DOCX->ODT->DOCX->ODT adds another additional space-->
                        <!--<xsl:value-of select="concat($suffix,' ')"/>-->
                        <xsl:value-of select="$suffix"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
                </xsl:if>
                <!--math, dialogika: Bugfix #1948002 Insert style:num-format="" if numFmt = 'none' -->
                <!--<xsl:if test="./w:lvlText/@w:val != ''">-->
                  <xsl:call-template name="NumFormat">
                    <xsl:with-param name="format" select="./w:numFmt/@w:val"/>
                    <xsl:with-param name="BeforeAfterNum">                      
                      <xsl:if test="./w:numFmt/@w:val != 'none'">
                         <xsl:value-of select="./w:lvlText/@w:val" />
                      </xsl:if>                      
                    </xsl:with-param> 
                  </xsl:call-template>
                <!--</xsl:if>-->
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
                  <!--  <xsl:variable name="Ind" select="./w:pPr/w:ind"/>
                    <xsl:attribute name="text:space-before">
                    <xsl:value-of select="number($Ind/@w:left)-number($Ind/@w:firstLine)"/>
                    </xsl:attribute>
                    <xsl:attribute name="text:min-label-distance">
                    <xsl:value-of select="number(./w:pPr/w:tabs/w:tab/@w:pos)"/>
                    </xsl:attribute>-->

                  <!--math, dialogika: changed for correct indentation calculation BEGIN -->
                  <!--added parameter <xsl:with-param name="numId">-->

                  <!--math, dialogika: Bugfix #1948002 InsertListLevelProperties only if numFmt is != none -->
                  <xsl:if test="./w:numFmt/@w:val != 'none'">                  
                    <xsl:call-template name="InsertListLevelProperties">
                      <xsl:with-param name="numId">
                        <xsl:value-of select="$numid"/>
                      </xsl:with-param>
                    </xsl:call-template>
                  </xsl:if>

                  <!--math, dialogika: changed for correct indentation calculation END -->                  
                  
                </style:list-level-properties>
              </xsl:for-each>
            </xsl:for-each>
          </text:outline-level-style>
      </xsl:for-each>
    </text:outline-style>
  </xsl:template>
</xsl:stylesheet>

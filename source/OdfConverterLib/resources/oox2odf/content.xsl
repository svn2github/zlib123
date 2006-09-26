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
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/3/main"
  exclude-result-prefixes="w">

  <xsl:import href="tables.xsl"/>

  <xsl:strip-space elements="*"/>
  <xsl:preserve-space elements="w:p"/>
  <xsl:preserve-space elements="w:r"/>

  <xsl:key name="numId" match="w:num" use="w:abstractNumId/@w:val"/>

  <xsl:template name="content">
    <office:document-content>
      <office:scripts/>
      <office:font-face-decls/>
      <office:automatic-styles>
        <xsl:apply-templates select="document('word/document.xml')/w:document/w:body"
          mode="automaticstyles"/>
        <xsl:if test="document('word/document.xml')/w:document[descendant::w:numPr]">
          <xsl:apply-templates select="document('word/numbering.xml')/w:numbering"/>
        </xsl:if>
      </office:automatic-styles>
      <office:body>
        <office:text>
          <xsl:apply-templates select="document('word/document.xml')/w:document/w:body"/>
        </office:text>
      </office:body>
    </office:document-content>
  </xsl:template>

  <!-- create a style for each run. Do not take w:pPr/w:rPr into consideration. -->
  <xsl:template match="w:rPr[parent::w:r]" mode="automaticstyles">
    <xsl:message terminate="no">progress:w:pPr</xsl:message>
    <style:style style:name="{generate-id(parent::w:r)}" style:family="text">
      <style:text-properties>
        <xsl:call-template name="InsertTextProperties"/>
      </style:text-properties>
    </style:style>
  </xsl:template>

  <xsl:template match="w:t" mode="automaticstyles"/>

  <xsl:template match="w:t" mode="text">
    <xsl:apply-templates select="."/>
  </xsl:template>

  <xsl:template match="text()" mode="text">
    <xsl:message terminate="no">progress:text()</xsl:message>
    <xsl:value-of select="."/>
  </xsl:template>

  <!--  Check if the paragraf is heading -->

  <xsl:template name="CheckHeading">
    <xsl:param name="outline"/>
    <xsl:variable name="basedOn">
      <xsl:value-of
        select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$outline]/w:basedOn/@w:val"
      />
    </xsl:variable>

    <xsl:choose>
      <xsl:when
        test="document('word/styles.xml')/w:styles/w:style[@w:styleId=$basedOn]/w:pPr/w:outlineLvl/@w:val">
        <xsl:value-of
          select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$basedOn]/w:pPr/w:outlineLvl/@w:val"
        />
      </xsl:when>
      <xsl:when
        test="document('word/styles.xml')/w:styles/w:style[@w:styleId=$outline]/w:basedOn/@w:val">
        <xsl:call-template name="CheckHeading">
          <xsl:with-param name="outline">
            <xsl:value-of
              select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$outline]/w:basedOn/@w:val"
            />
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
    </xsl:choose>

  </xsl:template>


  <xsl:template name="GetOutlineLevel">
    <xsl:choose>
      <xsl:when test="w:pPr/w:outlineLvl/@w:val">
        <xsl:value-of select="w:pPr/w:outlineLvl/@w:val"/>
      </xsl:when>

      <xsl:when test="w:pPr/w:pStyle/@w:val">

        <xsl:variable name="outline">
          <xsl:value-of select="w:pPr/w:pStyle/@w:val"/>
        </xsl:variable>

        <xsl:choose>
          <xsl:when
            test="document('word/styles.xml')/w:styles/w:style[@w:styleId=$outline]/w:pPr/w:outlineLvl/@w:val">
            <xsl:value-of
              select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$outline]/w:pPr/w:outlineLvl/@w:val"
            />
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="CheckHeading">
              <xsl:with-param name="outline">
                <xsl:value-of select="w:pPr/w:pStyle/@w:val"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <xsl:otherwise>

        <xsl:variable name="outline">
          <xsl:value-of select="w:r/w:rPr/w:rStyle/@w:val"/>
        </xsl:variable>

        <xsl:variable name="linkedStyleOutline">
          <xsl:value-of
            select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$outline]/w:link/@w:val"
          />
        </xsl:variable>

        <xsl:choose>
          <xsl:when
            test="document('word/styles.xml')/w:styles/w:style[@w:styleId=$linkedStyleOutline]/w:pPr/w:outlineLvl/@w:val">
            <xsl:value-of
              select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$linkedStyleOutline]/w:pPr/w:outlineLvl/@w:val"
            />
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="CheckHeading">
              <xsl:with-param name="outline">
                <xsl:value-of
                  select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$outline]/w:link/@w:val"
                />
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>

    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertHeading">
    <xsl:param name="outlineLevel"/>
    <text:h>
      <xsl:attribute name="text:outline-level">
        <xsl:value-of select="$outlineLevel+1"/>
      </xsl:attribute>
      <xsl:for-each select="w:bookmarkStart">
        <text:bookmark-start>
          <xsl:attribute name="text:name">
            <xsl:value-of select="@w:name"/>
          </xsl:attribute>
        </text:bookmark-start>
      </xsl:for-each>
      <xsl:value-of select="."/>
      <xsl:for-each select="w:bookmarkEnd">
        <text:bookmark-end>
          <xsl:attribute name="text:name">
            <xsl:variable name="bookmarkId">
              <xsl:value-of select="@w:id"/>
            </xsl:variable>
            <xsl:value-of select="preceding::w:bookmarkStart[@w:id = $bookmarkId]/@w:name"/>
          </xsl:attribute>
        </text:bookmark-end>
      </xsl:for-each>
    </text:h>
  </xsl:template>

  <xsl:template match="w:p">
    <xsl:message terminate="no">progress:w:p</xsl:message>
    <xsl:choose>

      <!-- check if list starts -->
      <xsl:when test="w:pPr/w:numPr">
        <xsl:variable name="NumberingId" select="w:pPr/w:numPr/w:numId/@w:val"/>

        <xsl:variable name="position" select="count(preceding-sibling::w:p)"/>
        <xsl:if
          test="not(preceding-sibling::node()[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId and  count(preceding-sibling::w:p)= $position -1])">
          <xsl:apply-templates select="." mode="list"/>
        </xsl:if>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates select="." mode="paragraph"/>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <xsl:template match="w:p" mode="paragraph">
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
        <text:p text:style-name="Standard">
          <!-- todo without styles -->
          <xsl:apply-templates/>
        </text:p>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- paragraph has  numbering properties - is a list item -->
  <xsl:template match="w:p" mode="list">
    <xsl:variable name="NumberingId2" select="w:pPr/w:numPr/w:numId/@w:val"/>

    <xsl:variable name="position2" select="count(preceding-sibling::w:p)"/>
    <xsl:choose>
      <!-- Is first  element at list -->
      <xsl:when
        test="not(preceding-sibling::node()[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId2 and  count(preceding-sibling::w:p)= $position2 -1])">
        <text:list text:style-name="{concat('L',$NumberingId2)}">
          <xsl:if test="preceding-sibling::w:p[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId2]">
            <xsl:attribute name="text:continue-numbering">true</xsl:attribute>
          </xsl:if>
          <text:list-item>
            <xsl:apply-templates select="." mode="paragraph"/>
          </text:list-item>
          <!-- next paragraph in same list member -->
          <xsl:if
            test="following-sibling::w:p[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId2 and  count(preceding-sibling::w:p)= $position2 +1]">
            <xsl:apply-templates
              select="following-sibling::w:p[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId2 and  count(preceding-sibling::w:p)= $position2 +1]"
              mode="list"/>
          </xsl:if>
        </text:list>
      </xsl:when>
      <xsl:otherwise>
        <text:list-item>
          <xsl:apply-templates select="." mode="paragraph"/>
        </text:list-item>
        <xsl:if
          test="following-sibling::w:p[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId2 and  count(preceding-sibling::w:p)= $position2 +1]">
          <xsl:apply-templates
            select="following-sibling::w:p[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId2 and  count(preceding-sibling::w:p)= $position2 +1]"
            mode="list"/>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="w:r">
    <xsl:message terminate="no">progress:w:r</xsl:message>
    <xsl:choose>
      <xsl:when
        test="count(preceding::w:fldChar[@w:fldCharType = 'begin']) &gt; count(preceding::w:fldChar[@w:fldCharType = 'end']) and contains(preceding::w:instrText[last()],'HYPERLINK')">

        <xsl:call-template name="hyperlink"/>

      </xsl:when>
      <xsl:when test="w:instrText"> </xsl:when>
      <xsl:when test="w:rPr">
        <text:span text:style-name="{generate-id(self::node())}">
          <xsl:apply-templates/>
        </text:span>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates/>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <xsl:template match="w:hyperlink">
    <xsl:call-template name="hyperlink"/>
  </xsl:template>

  <!-- hyperlink template -->

  <xsl:template name="hyperlink">
    <text:a xlink:type="simple">
      <xsl:choose>
        <!-- simple hyperlink -->
        <xsl:when test="@w:anchor">

          <xsl:attribute name="xlink:href">

            <xsl:value-of select="concat('#',@w:anchor)"/>


          </xsl:attribute>
          <text:span text:style-name="Internet_20_link">
            <xsl:apply-templates select="w:r/w:t"/>
          </text:span>
        </xsl:when>
        <!-- fieldchar hyperlink -->
        <xsl:when test="self::w:r">
          <xsl:attribute name="xlink:href">
            <xsl:value-of
              select="concat('#_',substring-before(substring-after(preceding::w:instrText[last()],'_'),'&quot;'))"
            />
          </xsl:attribute>
          <text:span text:style-name="Internet_20_link">
            <xsl:apply-templates select="w:t"/>
          </text:span>
        </xsl:when>
        <xsl:otherwise> </xsl:otherwise>
      </xsl:choose>
      <!-- TODO internet hyperlinks - problem with resolving xml nodes from *.rels files -->
      <!-- <xsl:if test="@r:id">
        <xsl:variable name="relationshipId">
          <xsl:value-of select="@r:id"/>
        </xsl:variable>
        <xsl:attribute name="xlink:href">
            <xsl:value-of select="document('word/_rels/document.xml.rels')/rels/Relationships/@xmlns"/>
        </xsl:attribute>
        <xsl:apply-templates select="w:r/w:t"/>
      </xsl:if> -->

    </text:a>

  </xsl:template>

  <xsl:template match="w:t">
    <xsl:message terminate="no">progress:w:t</xsl:message>
    <xsl:value-of select="."/>
  </xsl:template>

  <!-- line breaks  (page-break todo)-->
  <xsl:template match="w:br">
    <text:line-break/>
  </xsl:template>

  <!-- following templates are from numbering.xml -->
  <xsl:template match="w:abstractNum">

    <text:list-style style:name="{concat('L',key('numId',@w:abstractNumId)/@w:numId)}">
      <xsl:apply-templates select="w:lvl"/>
    </text:list-style>
  </xsl:template>

  <xsl:template match="w:lvl">
    <xsl:choose>
      <xsl:when test="w:numFmt[@w:val = 'bullet']">
        <text:list-level-style-bullet text:level="{number(@w:ilvl)+1}"
          text:style-name="Bullet_20_Symbols">
          <xsl:attribute name="text:bullet-char">
            <xsl:call-template name="textChar"/>


          </xsl:attribute>
          <style:list-level-properties>
            <xsl:call-template name="listLevelProperties"/>
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
            <xsl:call-template name="num-format">
              <xsl:with-param name="format" select="w:numFmt/@w:val"/>
            </xsl:call-template>
          </xsl:attribute>
          <style:list-level-properties>
            <xsl:call-template name="listLevelProperties"/>
          </style:list-level-properties>
        </text:list-level-style-number>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- numbering format -->
  <xsl:template name="num-format">
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

  <xsl:template name="listLevelProperties">
    <xsl:variable name="Ind" select="w:pPr/w:ind"/>
    <xsl:attribute name="text:space-before">
      <xsl:value-of select="number($Ind/@w:left)-number($Ind/@w:hanging)"/>
    </xsl:attribute>
    <xsl:attribute name="text:min-label-width">
      <xsl:value-of select="number($Ind/@w:hanging)"/>
    </xsl:attribute>
  </xsl:template>

  <!-- types of bullets -->
  <xsl:template name="textChar">
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
      <xsl:when test="w:lvlText[@w:val = '–' ]">–</xsl:when>
      <xsl:otherwise>•</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

</xsl:stylesheet>

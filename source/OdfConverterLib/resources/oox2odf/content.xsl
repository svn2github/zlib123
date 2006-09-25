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
      <xsl:when test="w:pPr/w:outlineLvl/@w:val">
        <xsl:value-of select="w:pPr/w:outlineLvl/@w:val"/>
      </xsl:when>

      <xsl:when test="w:pPr/w:pStyle/@w:val">

        <xsl:variable name="outline">
          <xsl:value-of select="w:pPr/w:pStyle/@w:val"/>
        </xsl:variable>

        <xsl:value-of
          select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$outline]/w:pPr/w:outlineLvl/@w:val"/>
        <xsl:message> lvl val: <xsl:value-of select="$outline"/>
        </xsl:message>
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
    <xsl:choose>
      
      <!-- check if list starts -->
      <xsl:when test="w:pPr/w:numPr">
        <xsl:variable name="NumberingId" select="w:pPr/w:numPr/w:numId/@w:val"/>
        <xsl:variable name="position" select="count(preceding-sibling::w:p)"/>
        <xsl:if test="not(preceding-sibling::node()[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId and  count(preceding-sibling::w:p)= $position -1])">
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
    <xsl:variable name="NumberingId" select="w:pPr/w:numPr/w:numId/@w:val"/>
    <xsl:variable name="position" select="count(preceding-sibling::w:p)"/>
    <xsl:choose>
      <!-- Is first  element at list -->
      <xsl:when test="not(preceding-sibling::node()[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId and  count(preceding-sibling::w:p)= $position -1])">
        <text:list text:style-name="L1">
          <xsl:if test="preceding-sibling::w:p[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId]">
            <xsl:attribute name="text:continue-numbering">true</xsl:attribute>
          </xsl:if>
          <text:list-item>
            <xsl:apply-templates select="." mode="paragraph"/>
          </text:list-item>
          <!-- next paragraph in same list member -->
          <xsl:if test="following-sibling::w:p[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId and  count(preceding-sibling::w:p)= $position +1]">
            <xsl:apply-templates select="following-sibling::w:p[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId and  count(preceding-sibling::w:p)= $position +1]" mode="list"/>
          </xsl:if>
        </text:list>
      </xsl:when>
      <xsl:otherwise>
        <text:list-item>
          <xsl:apply-templates select="." mode="paragraph"/>
        </text:list-item>
        <xsl:if test="following-sibling::w:p[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId and  count(preceding-sibling::w:p)= $position +1]">
          <xsl:apply-templates select="following-sibling::w:p[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId and  count(preceding-sibling::w:p)= $position +1]" mode="list"/>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="w:r">

    <xsl:apply-templates/>

  </xsl:template>
  <xsl:template match="w:hyperlink">
  <xsl:call-template name="hyperlink"/>
  </xsl:template>
  <xsl:template name="hyperlink">
    <text:a xlink:type="simple"><xsl:if test="@w:anchor">
      
      <xsl:attribute name="xlink:href">
        
          <xsl:value-of select="concat('#',@w:anchor)"/>
      
      </xsl:attribute>
      <text:span text:style-name="Internet_20_link">
        <xsl:value-of select="w:r/w:t"/>
      </text:span>
    </xsl:if>
    </text:a>

  </xsl:template>
  <xsl:template match="w:t">
    <xsl:message terminate="no">progress:w:t</xsl:message>
    <xsl:value-of select="."/>
  </xsl:template>


  <!-- Piece of code used to call style properties templates.
  
  <xsl:template match="w:style">
    <style:style>
      <xsl:if test="w:pPr">
        <style:paragraph-properties>
          <xsl:for-each select="w:pPr">
            <xsl:call-template name="InsertParagraphProperties"/>
          </xsl:for-each>
        </style:paragraph-properties>
      </xsl:if>

      <xsl:if test="w:rPr">
        <style:text-properties>
          <xsl:for-each select="w:rPr">
            <xsl:call-template name="InsertTextProperties"/>
          </xsl:for-each>
          <xsl:for-each select="w:pPr">
            <xsl:call-template name="InsertpPrTextProperties"/>
          </xsl:for-each>
        </style:text-properties>
      </xsl:if>
    </style:style>
  </xsl:template>
  
-->
  <!-- conversion of paragraph properties -->
  <xsl:template name="InsertParagraphProperties">

    <xsl:if test="w:keepNext">
      <xsl:attribute name="fo:keep-with-next">
        <xsl:choose>
          <xsl:when
            test="w:keepNext/@w:val='off' or w:keepNext/@w:val='false' or w:keepNext/@w:val=0">
            <xsl:value-of select="'auto'"/>
          </xsl:when>
          <xsl:when test="w:keepNext/@w:val='on' or w:keepNext/@w:val='true' or w:keepNext/@w:val=1">
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
      <xsl:call-template name="InsertParagraphBorder"/>
      <!-- add shadow -->
      <xsl:call-template name="InsertParagraphShadow"/>
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
    <!-- w:afterLines and w:beforeLines elements are lost -->
    <xsl:if
      test="not(w:spacing/@w:beforeAutospacing='on' or w:spacing/@w:beforeAutospacing='1' or w:spacing/@w:beforeAutospacing='true') and w:spacing/@w:before">
      <xsl:attribute name="fo:margin-top">
        <xsl:call-template name="ConvertTwips">
          <xsl:with-param name="length">
            <xsl:value-of select="w:spacing/@w:before"/>
          </xsl:with-param>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:if
      test="not(w:spacing/@w:afterAutospacing='on' or w:spacing/@w:afterAutospacing='1' or w:spacing/@w:afterAutospacing='true') and w:spacing/@w:after">
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

    <!-- tab stops -->
    <xsl:if test="w:tabs/w:tab">
      <style:tab-stops>
        <xsl:apply-templates select="w:tabs/w:tab"/>
      </style:tab-stops>
    </xsl:if>

  </xsl:template>

  <xsl:template name="InsertParagraphBorder">
    <xsl:if test="w:pBdr/w:between">
      <xsl:attribute name="style:join-border">false</xsl:attribute>
    </xsl:if>

    <xsl:for-each select="w:pBdr">
      <xsl:choose>
        <!-- if the four borders are equal, then create adequate attributes. -->
        <xsl:when
          test="w:top/@color=w:bottom/@color and w:top/@space=w:bottom/@space and w:top/@sz=w:bottom/@sz and w:top/@val=w:bottom/@val
          and w:top/@color=w:left/@color and w:top/@space=w:left/@space and w:top/@sz=w:left/@sz and w:top/@val=w:left/@val
          and w:top/@color=w:right/@color and w:top/@space=w:right/@space and w:top/@sz=w:right/@sz and w:top/@val=w:right/@val">
          <xsl:call-template name="InsertBorderAttributes"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each
            select="child::node()[name() = 'w:top' or name() = 'w:left' or name() = 'w:bottom' or name() = 'w:right']">
            <xsl:call-template name="InsertBorderAttributes">
              <xsl:with-param name="side" select="substring-after(name(),'w:')"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <!-- compute attributes defining borders : color, style, width, padding... -->
  <xsl:template name="InsertBorderAttributes">
    <xsl:param name="side"/>

    <xsl:choose>
      <xsl:when test="$side">
        <!-- add padding -->
        <xsl:if test="@w:space">
          <xsl:attribute name="{concat('fo:padding-',$side)}">
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
        <xsl:attribute name="{concat('fo:border-',$side)}">
          <xsl:variable name="size">
            <xsl:call-template name="ConvertEighthsPoints">
              <xsl:with-param name="length">
                <xsl:value-of select="@w:sz"/>
              </xsl:with-param>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test="@w:color='auto'">
              <xsl:value-of select="concat($size,' ',$style)"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:variable name="color">
                <xsl:value-of select="concat('#',@w:color)"/>
              </xsl:variable>
              <xsl:value-of select="concat($size,' ',$style,' ',$color)"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:if test="$style='double' ">
          <xsl:attribute name="{concat('fo:border-line-width-',$side)}">
            <xsl:call-template name="ComputeDoubleBorderWidth">
              <xsl:with-param name="style" select="@w:val"/>
              <xsl:with-param name="borderWidth" select="@w:sz"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <xsl:otherwise>
        <!-- add padding -->
        <xsl:if test="w:top/@w:space">
          <xsl:attribute name="fo:padding">
            <xsl:call-template name="ConvertPoints">
              <xsl:with-param name="length">
                <xsl:value-of select="w:top/@w:space"/>
              </xsl:with-param>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
        <!-- add border with style, width and color -->
        <xsl:variable name="style">
          <xsl:call-template name="GetParagraphBorderStyle">
            <xsl:with-param name="style" select="w:top/@w:val"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:attribute name="fo:border">
          <xsl:variable name="size">
            <xsl:call-template name="ConvertEighthsPoints">
              <xsl:with-param name="length">
                <xsl:value-of select="w:top/@w:sz"/>
              </xsl:with-param>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test="@w:color='auto'">
              <xsl:value-of select="concat($size,' ',$style)"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:variable name="color">
                <xsl:value-of select="concat('#',w:top/@w:color)"/>
              </xsl:variable>
              <xsl:value-of select="concat($size,' ',$style,' ',$color)"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:if test="$style='double' ">
          <xsl:attribute name="fo:border-line-width">
            <xsl:call-template name="ComputeDoubleBorderWidth">
              <xsl:with-param name="style" select="w:top/@w:val"/>
              <xsl:with-param name="borderWidth" select="w:top/@w:sz"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <xsl:template name="InsertParagraphShadow">

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

  <!-- Handle tab stops -->
  <!-- TODO : check how to deal with tab stops inside a list -->
  <xsl:template match="w:tab">
    <style:tab-stop>
      <xsl:if test="@w:val != 'num' and @w:val != 'clear'">
        <!-- type -->
        <xsl:attribute name="style:type">
          <xsl:choose>
            <xsl:when test="@w:val='center'">center</xsl:when>
            <xsl:when test="@w:val='right'">right</xsl:when>
            <xsl:when test="@w:val='left'">left</xsl:when>
            <xsl:otherwise>left</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <!-- position -->
        <!-- TODO : what if @w:pos < 0 ? -->
        <xsl:if test="@w:pos >= 0">
          <xsl:attribute name="style:position">
            <xsl:call-template name="ConvertTwips">
              <xsl:with-param name="length" select="@w:pos"/>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
        <!-- leader char -->
        <xsl:if test="@w:leader">
          <xsl:attribute name="style:leader-style">
            <xsl:choose>
              <xsl:when test="@w:leader='dot'">dotted</xsl:when>
              <xsl:when test="@w:leader='heavy'">solid</xsl:when>
              <xsl:when test="@w:leader='hyphen'">dash</xsl:when>
              <xsl:when test="@w:leader='middleDot'">dotted</xsl:when>
              <xsl:when test="@w:leader='none'">none</xsl:when>
              <xsl:when test="@w:leader='underscore'">solid</xsl:when>
              <xsl:otherwise>none</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>
      </xsl:if>
    </style:tab-stop>
  </xsl:template>

  <xsl:template name="InsertpPrTextProperties">
    <!-- hyphenation -->
    <xsl:if test="w:suppressAutoHyphens">
      <xsl:choose>
        <xsl:when
          test="w:suppressAutoHyphens/@w:val='off' or w:suppressAutoHyphens/@w:val='false' or w:suppressAutoHyphens/@w:val=0">
          <xsl:attribute name="fo:hyphenate">true</xsl:attribute>
          <xsl:attribute name="fo:hyphenation-remain-char-count">2</xsl:attribute>
          <xsl:attribute name="fo:hyphenation-push-char-count">2</xsl:attribute>
        </xsl:when>
        <xsl:when
          test="w:suppressAutoHyphens/@w:val='on' or w:suppressAutoHyphens/@w:val='true' or w:suppressAutoHyphens/@w:val=1">
          <xsl:attribute name="fo:hyphenate">false</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="fo:hyphenate">false</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
  </xsl:template>

  <!-- conversion of text properties -->
  <xsl:template name="InsertTextProperties">

    <xsl:if test="w:b">
      <xsl:attribute name="fo:font-weight">
        <xsl:choose>
          <xsl:when test="w:b/@w:val='off' or w:b/@w:val='false' or w:b/@w:val=0">
            <xsl:value-of select="'normal'"/>
          </xsl:when>
          <xsl:when test="w:b/@w:val='on' or w:b/@w:val='true' or w:b/@w:val=1">
            <xsl:value-of select="'bold'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'bold'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:bCs">
      <xsl:attribute name="fo:font-weight-complex">
        <xsl:choose>
          <xsl:when test="w:bCs/@w:val='off' or w:bCs/@w:val='false' or w:bCs/@w:val=0">
            <xsl:value-of select="'normal'"/>
          </xsl:when>
          <xsl:when test="w:bCs/@w:val='on' or w:bCs/@w:val='true' or w:bCs/@w:val=1">
            <xsl:value-of select="'bold'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'bold'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:i">
      <xsl:attribute name="fo:font-style">
        <xsl:choose>
          <xsl:when test="w:i/@w:val='off' or w:i/@w:val='false' or w:i/@w:val=0">
            <xsl:value-of select="'normal'"/>
          </xsl:when>
          <xsl:when test="w:i/@w:val='on' or w:i/@w:val='true' or w:i/@w:val=1">
            <xsl:value-of select="'italic'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'italic'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:iCs">
      <xsl:attribute name="fo:font-style-complex">
        <xsl:choose>
          <xsl:when test="w:iCs/@w:val='off' or w:iCs/@w:val='false' or w:iCs/@w:val=0">
            <xsl:value-of select="'normal'"/>
          </xsl:when>
          <xsl:when test="w:iCs/@w:val='on' or w:iCs/@w:val='true' or w:iCs/@w:val=1">
            <xsl:value-of select="'italic'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'italic'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:caps">
      <xsl:attribute name="fo:text-transform">
        <xsl:choose>
          <xsl:when test="w:caps/@w:val='off' or w:caps/@w:val='false' or w:caps/@w:val=0">
            <xsl:value-of select="'none'"/>
          </xsl:when>
          <xsl:when test="w:caps/@w:val='on' or w:caps/@w:val='true' or w:caps/@w:val=1">
            <xsl:value-of select="'uppercase'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'uppercase'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:smallCaps">
      <xsl:attribute name="fo:font-variant">
        <xsl:choose>
          <xsl:when
            test="w:smallCaps/@w:val='off' or w:smallCaps/@w:val='false' or w:smallCaps/@w:val=0">
            <xsl:value-of select="'normal'"/>
          </xsl:when>
          <xsl:when
            test="w:smallCaps/@w:val='on' or w:smallCaps/@w:val='true' or w:smallCaps/@w:val=1">
            <xsl:value-of select="'small-caps'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'small-caps'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:dstrike or w:strike">
      <xsl:choose>
        <xsl:when test="w:strike/@w:val='on' or w:strike/@w:val='true' or w:strike/@w:val=1">
          <xsl:attribute name="style:text-line-through-type">single</xsl:attribute>
          <xsl:attribute name="style:text-line-through-style">solid</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:dstrike/@w:val='on' or w:dstrike/@w:val='true' or w:dstrike/@w:val=1">
          <xsl:attribute name="style:text-line-through-type">double</xsl:attribute>
          <xsl:attribute name="style:text-line-through-style">solid</xsl:attribute>
        </xsl:when>
        <xsl:otherwise/>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="w:outline">
      <xsl:attribute name="style:text-outline">
        <xsl:choose>
          <xsl:when test="w:outline/@w:val='off' or w:outline/@w:val='false' or w:outline/@w:val=0">
            <xsl:value-of select="'false'"/>
          </xsl:when>
          <xsl:when test="w:outline/@w:val='on' or w:outline/@w:val='true' or w:outline/@w:val=1">
            <xsl:value-of select="'true'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'true'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:shadow">
      <xsl:attribute name="fo:text-shadow">
        <xsl:choose>
          <xsl:when test="w:shadow/@w:val='off' or w:shadow/@w:val='false' or w:shadow/@w:val=0">
            <xsl:value-of select="'none'"/>
          </xsl:when>
          <xsl:when test="w:shadow/@w:val='on' or w:shadow/@w:val='true' or w:shadow/@w:val=1">
            <xsl:value-of select="'#000000 0.2em 0.2em'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'#000000 0.2em 0.2em'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:emboss or w:imprint">
      <xsl:attribute name="style:font-relief">
        <xsl:choose>
          <xsl:when test="w:imprint/@w:val='on' or w:imprint/@w:val='true' or w:imprint/@w:val=1">
            <xsl:value-of select="'engraved'"/>
          </xsl:when>
          <xsl:when test="w:emboss/@w:val='on' or w:emboss/@w:val='true' or w:emboss/@w:val=1">
            <xsl:value-of select="'embossed'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'none'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:vanish">
      <xsl:attribute name="text:display">
        <xsl:choose>
          <xsl:when test="w:vanish/@w:val='off' or w:vanish/@w:val='false' or w:vanish/@w:val=0">
            <xsl:value-of select="'true'"/>
          </xsl:when>
          <xsl:when test="w:vanish/@w:val='on' or w:vanish/@w:val='true' or w:vanish/@w:val=1">
            <xsl:value-of select="'none'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'true'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:em">
      <xsl:attribute name="style:text-emphasize">
        <xsl:choose>
          <xsl:when test="w:em/@w:val = 'circle'">circle above</xsl:when>
          <xsl:when test="w:em/@w:val = 'comma'">accent above</xsl:when>
          <xsl:when test="w:em/@w:val = 'dot'">dot above</xsl:when>
          <xsl:when test="w:em/@w:val = 'underDot'">dot below</xsl:when>
          <xsl:when test="w:em/@w:val = 'none'">none</xsl:when>
          <xsl:otherwise>none</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <!-- text bg color -->
    <xsl:if test="w:highlight">
      <xsl:attribute name="fo:background-color">
        <xsl:choose>
          <xsl:when test="w:highlight/@w:val = 'black'">#000000</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'blue'">#0000FF</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'cyan'">#00FFFF</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'darkBlue'">#000080</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'darkCyan'">#008080</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'darkGray'">#808080</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'darkGreen'">#008000</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'darkMagenta'">#800080</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'darkRed'">#800000</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'darkYellow'">#808000</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'green'">#00FF00</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'lightGray'">#C0C0C0</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'magenta'">#FF00FF</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'red'">#FF0000</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'white'">#FFFFFF</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'yellow'">#FFFF00</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'none'">transparent</xsl:when>
          <xsl:otherwise>transparent</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <!-- underline -->
    <xsl:if test="w:u">
      <xsl:call-template name="InsertUnderline"/>
    </xsl:if>

    <!-- fonts -->
    <xsl:if test="w:rFonts/@w:ascii">
      <xsl:attribute name="style:font-name">
        <xsl:call-template name="ComputeFontName">
          <xsl:with-param name="fontName" select="w:rFonts/@w:ascii"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="w:rFonts/@w:cs">
      <xsl:attribute name="style:font-name-complex">
        <xsl:call-template name="ComputeFontName">
          <xsl:with-param name="fontName" select="w:rFonts/@w:cs"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="w:rFonts/@w:eastAsia">
      <xsl:attribute name="style:font-name-asian">
        <xsl:call-template name="ComputeFontName">
          <xsl:with-param name="fontName" select="w:rFonts/@w:eastAsia"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>

    <!-- text color -->
    <xsl:if test="w:color/@w:val != 'auto'">
      <xsl:attribute name="fo:color">
        <xsl:value-of select="concat('#',w:color/@w:val)"/>
      </xsl:attribute>
    </xsl:if>

    <!-- letter spacing -->
    <xsl:if test="w:spacing">
      <xsl:attribute name="fo:letter-spacing">
        <xsl:choose>
          <xsl:when test="w:spacing/@w:val=0">normal</xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="ConvertTwips">
              <xsl:with-param name="length" select="w:spacing/@w:val"/>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:w/@w:val">
      <xsl:attribute name="style:text-scale">
        <xsl:value-of select="w:w/@w:val"/>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:kern">
      <xsl:attribute name="style:letter-kerning">true</xsl:attribute>
    </xsl:if>

    <xsl:if test="w:sz">
      <xsl:attribute name="fo:font-size">
        <xsl:call-template name="ConvertHalfPoints">
          <xsl:with-param name="length" select="w:sz/@w:val"/>
          <xsl:with-param name="unit">pt</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:szCs">
      <xsl:attribute name="fo:font-size-complex">
        <xsl:call-template name="ConvertHalfPoints">
          <xsl:with-param name="length" select="w:szCs/@w:val"/>
          <xsl:with-param name="unit">pt</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>

    <!-- text vertical positionning -->
    <xsl:if test="w:vertAlign or w:position">
      <xsl:call-template name="InsertTextPosition"/>
    </xsl:if>

    <!-- languages and countries -->
    <xsl:if test="w:lang">
      <xsl:if test="w:lang/@w:val">
        <xsl:attribute name="fo:language">
          <xsl:value-of select="substring-before(w:lang/@w:val,'-')"/>
        </xsl:attribute>
        <xsl:attribute name="fo:country">
          <xsl:value-of select="substring-after(w:lang/@w:val,'-')"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="w:lang/@w:bidi">
        <xsl:attribute name="style:language-complex">
          <xsl:value-of select="substring-before(w:lang/@w:bidi,'-')"/>
        </xsl:attribute>
        <xsl:attribute name="style:country-complex">
          <xsl:value-of select="substring-after(w:lang/@w:bidi,'-')"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="w:lang/@w:eastAsia">
        <xsl:attribute name="style:language-asian">
          <xsl:value-of select="substring-before(w:lang/@w:eastAsia,'-')"/>
        </xsl:attribute>
        <xsl:attribute name="style:country-asian">
          <xsl:value-of select="substring-after(w:lang/@w:eastAsia,'-')"/>
        </xsl:attribute>
      </xsl:if>
    </xsl:if>

    <xsl:if test="w:eastAsianLayout">
      <xsl:if test="w:eastAsianLayout/@w:combine = 'lines'">
        <xsl:attribute name="style:text-combine">lines</xsl:attribute>
        <xsl:choose>
          <xsl:when test="w:combineBrackets = 'angle'">
            <xsl:attribute name="style:text-combine-start-char">
              <xsl:value-of select="'&lt;'"/>
            </xsl:attribute>
            <xsl:attribute name="style:text-combine-end-char">
              <xsl:value-of select="'&gt;'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="w:combineBrackets = 'curly'">
            <xsl:attribute name="style:text-combine-start-char">{</xsl:attribute>
            <xsl:attribute name="style:text-combine-end-char">}</xsl:attribute>
          </xsl:when>
          <xsl:when test="w:combineBrackets = 'round'">
            <xsl:attribute name="style:text-combine-start-char">(</xsl:attribute>
            <xsl:attribute name="style:text-combine-end-char">)</xsl:attribute>
          </xsl:when>
          <xsl:when test="w:combineBrackets = 'square'">
            <xsl:attribute name="style:text-combine-start-char">[</xsl:attribute>
            <xsl:attribute name="style:text-combine-end-char">]</xsl:attribute>
          </xsl:when>
          <xsl:otherwise/>
        </xsl:choose>
      </xsl:if>
    </xsl:if>

  </xsl:template>

  <!-- insert underline attributes -->
  <xsl:template name="InsertUnderline">
    <xsl:if test="w:u/@w:val">
      <xsl:choose>
        <xsl:when test="w:u/@w:val = 'dash'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">dash</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'dashDotDotHeavy'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">dot-dot-dash</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'dashDotHeavy'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">dot-dash</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'dashedHeavy'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">dash</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'dashLong'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">long-dash</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'dashLongHeavy'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">long-dash</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'dotDash'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">dot-dash</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'dotDotDash'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">dot-dot-dash</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'dotted'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">dotted</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'dottedHeavy'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">dotted</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'double'">
          <xsl:attribute name="style:text-underline-type">double</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">solid</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'single'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">solid</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'thick'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">solid</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'wave'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">wave</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'wavyDouble'">
          <xsl:attribute name="style:text-underline-type">double</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">wave</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'wavyHeavy'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">wave</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'words'">
          <xsl:attribute name="style:text-underline-mode">skip-white-space</xsl:attribute>
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">solid</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'none'">
          <xsl:attribute name="style:text-underline-type">none</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">solid</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <xsl:if test="w:u/@w:color">
      <xsl:attribute name="style:text-underline-color">
        <xsl:choose>
          <xsl:when test="w:u/@w:color = 'auto'">font-color</xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat('#',w:u/@w:color)"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

  <!-- compute positionning of text -->
  <xsl:template name="InsertTextPosition">
    <xsl:variable name="percentValue">
      <xsl:choose>
        <xsl:when test="w:position/@w:val">
          <xsl:choose>
            <xsl:when test="w:sz/@w:val != 0">
              <xsl:value-of select="round(w:position/@w:val * 100 div w:sz/@w:val)"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="round(w:position/@w:val * 100 div 24)"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>none</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:choose>

      <xsl:when test="w:vertAlign = 'superscript'">
        <xsl:choose>
          <xsl:when test="$percentValue > 0">
            <xsl:attribute name="style:text-position">
              <xsl:value-of select="concat('super ',$percentValue)"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="style:text-position">super</xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <xsl:when test="w:vertAlign = 'subscript'">
        <xsl:choose>
          <xsl:when test="$percentValue &lt; 0">
            <xsl:attribute name="style:text-position">
              <xsl:value-of select="concat('sub ',number( - $percentValue))"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="style:text-position">sub</xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <xsl:when test="w:vertAlign = 'baseline'">
        <xsl:choose>
          <xsl:when test="$percentValue > 0">
            <xsl:attribute name="style:text-position">
              <xsl:value-of select="concat('super ',$percentValue)"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="$percentValue &lt; 0">
            <xsl:attribute name="style:text-position">
              <xsl:value-of select="concat('sub ',number( - $percentValue))"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise/>
        </xsl:choose>
      </xsl:when>

      <xsl:when test="$percentValue > 0">
        <xsl:attribute name="style:text-position">
          <xsl:value-of select="concat('super ',$percentValue)"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="$percentValue &lt; 0">
        <xsl:attribute name="style:text-position">
          <xsl:value-of select="concat('sub ',number( - $percentValue))"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:otherwise/>

    </xsl:choose>
  </xsl:template>

  <!-- return the name of the font -->
  <xsl:template name="ComputeFontName">
    <xsl:param name="fontName"/>
    <xsl:choose>
      <xsl:when test="$fontName = 'Symbol'">StarSymbol</xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$fontName"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

</xsl:stylesheet>

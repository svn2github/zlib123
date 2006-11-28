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
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  exclude-result-prefixes="text fo style office draw">


  <xsl:strip-space elements="*"/>
  <xsl:preserve-space elements="text:p"/>
  <xsl:preserve-space elements="text:span"/>

  <!-- key to find bookmarkt or reference-mark -->
  <xsl:key name="bookmark-reference-start"
    match="text:bookmark|text:bookmark-start|text:reference-mark-start|text:sequence"
    use="@text:name|@text:ref-name"/>

  <!--checks if element has style used to generate table of contents in document  -->
  <xsl:template name="IsTOCBookmark">
    <xsl:param name="styleName"/>
    <xsl:param name="tableOfContentsNum" select="count(key('toc',''))"/>

    <xsl:variable name="tableOfContent" select="key('toc', '')[$tableOfContentsNum]"/>
    <xsl:variable name="tocStyle">
      <xsl:call-template name="IsTOCStyleOrElement">
        <xsl:with-param name="sourceStyleNum"
          select="count($tableOfContent/text:table-of-content-source/text:index-source-styles)"/>
        <xsl:with-param name="styleName" select="$styleName"/>
        <xsl:with-param name="tableOfContent" select="$tableOfContent"/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="$tocStyle = 'true'">true</xsl:when>
      <xsl:otherwise>false</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertTOCBookmark">
    <xsl:param name="tableOfContentsNum" select="count(key('toc',''))"/>
    <xsl:param name="bookmarkType"/>
    <xsl:param name="styleName" select="@text:style-name"/>

    <xsl:choose>
      <xsl:when test="$tableOfContentsNum > 0">

        <xsl:variable name="isBookmarked">
          <xsl:call-template name="IsTOCBookmark">
            <xsl:with-param name="styleName" select="$styleName"/>
            <xsl:with-param name="tableOfContentsNum" select="$tableOfContentsNum"/>
          </xsl:call-template>
        </xsl:variable>

        <xsl:if test="$isBookmarked = 'true' ">
          <xsl:variable name="bookmarkId">
            <xsl:call-template name="CalculateBookmarkId">
              <xsl:with-param name="counter" select="1"/>
              <xsl:with-param name="tableOfContent" select="key('toc', '')[$tableOfContentsNum]"/>
            </xsl:call-template>
          </xsl:variable>

          <xsl:call-template name="InsertBookmarkStartTOC">
            <xsl:with-param name="tocId" select="$bookmarkId"/>
            <xsl:with-param name="tableOfContent" select="key('toc', '')[$tableOfContentsNum]"/>
          </xsl:call-template>

          <xsl:call-template name="InsertBookmarkEndTOC">
            <xsl:with-param name="tocId" select="$bookmarkId"/>
          </xsl:call-template>
        </xsl:if>

        <xsl:call-template name="InsertTOCBookmark">
          <xsl:with-param name="tableOfContentsNum" select="$tableOfContentsNum - 1"/>
          <xsl:with-param name="bookmarkType" select="$bookmarkType"/>
          <xsl:with-param name="styleName" select="$styleName"/>
        </xsl:call-template>
      </xsl:when>
    </xsl:choose>
  </xsl:template>



  <!-- Insert BookmarkStart Id or BookmarkEnd Id -->
  <xsl:template name="GenerateBookmarkId">
    <xsl:param name="TextName"/>
    <xsl:variable name="ReferenceMarkStart">
      <xsl:value-of select="count(key('bookmark-reference-start', $TextName)/preceding-sibling::text:reference-mark-start )+count(key('bookmark-reference-start', $TextName)/parent::text:p/preceding-sibling::node()/text:reference-mark-start )"/>
    </xsl:variable>
    <xsl:variable name="BookmarkStart">
      <xsl:value-of select="count(key('bookmark-reference-start', $TextName)/preceding-sibling::text:bookmark-start)+count(key('bookmark-reference-start', $TextName)/parent::text:p/preceding-sibling::node()/text:bookmark-start)"/>
    </xsl:variable>
    <xsl:variable name="Bookmark">
      <xsl:value-of select="count(key('bookmark-reference-start', $TextName)/preceding-sibling::text:bookmark)+count(key('bookmark-reference-start', $TextName)/parent::text:p/preceding-sibling::node()/text:bookmark)"/>
    </xsl:variable>
    <xsl:value-of
      select="$ReferenceMarkStart+$BookmarkStart+$Bookmark"/>
  </xsl:template>
  
  <!-- Insert BookmarkStart or ReferenceMarkStart-->
  <xsl:template match="text:bookmark-start|text:reference-mark-start|text:bookmark" mode="paragraph">    
    <w:bookmarkStart>
      <xsl:attribute name="w:id">
        <xsl:call-template name="GenerateBookmarkId">
          <xsl:with-param name="TextName">
            <xsl:value-of select="@text:name"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="w:name">
        <xsl:value-of select="@text:name"/>
      </xsl:attribute>
    </w:bookmarkStart>
    <xsl:if test="name()='text:bookmark'">
      <w:bookmarkEnd>
        <xsl:attribute name="w:id">
          <xsl:call-template name="GenerateBookmarkId">
            <xsl:with-param name="TextName">
              <xsl:value-of select="@text:name"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
        </w:bookmarkEnd>
    </xsl:if>
  </xsl:template>

  <!-- Insert BookmarkEnd-->
  <xsl:template match="text:bookmark-end|text:reference-mark-end" mode="paragraph">
    <w:bookmarkEnd>
      <xsl:attribute name="w:id">
        <xsl:call-template name="GenerateBookmarkId">
          <xsl:with-param name="TextName">
            <xsl:value-of select="@text:name"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </w:bookmarkEnd>
  </xsl:template>

  <!-- Insert Cross References (Bookmark) -->
  <xsl:template match="text:bookmark-ref|text:reference-ref|text:sequence-ref" mode="paragraph">
    <xsl:variable name="TextName">
      <xsl:value-of select="@text:ref-name"/>
    </xsl:variable>
    <xsl:variable name="masterPage"
      select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page/style:header/text:p"/>
    <xsl:if
     test="key('bookmark-reference-start', $TextName) or $masterPage/text:reference-mark-start[@text:name=$TextName] or $masterPage/text:bookmark-start[@text:name=$TextName]">
      <xsl:variable name="CrossReferences">
        <xsl:choose>
          <xsl:when
            test="@text:reference-format='page' or ../text:sequence-ref[@text:ref-name=$TextName]">
            <xsl:text>PAGEREF </xsl:text>
          </xsl:when>
          <xsl:otherwise>REF </xsl:otherwise>
        </xsl:choose>
        <xsl:choose>
          <xsl:when test="../text:sequence-ref[@text:ref-name=$TextName]">
            <xsl:value-of
              select="concat('_Toc', number(count(key('bookmark-reference-start', $TextName)/preceding::text:sequence))+1)"
            />
          </xsl:when>
          <xsl:otherwise>
            <xsl:choose>
              <xsl:when test="contains($TextName,' ' )">
                <xsl:value-of select="translate($TextName,' ','_')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$TextName"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
        <xsl:if test="@text:reference-format='direction'">
          <xsl:text>\p</xsl:text>
        </xsl:if>
        <xsl:text> \h</xsl:text>
      </xsl:variable>
      <w:r>
        <w:fldChar w:fldCharType="begin"/>
      </w:r>
      <w:r>
        <w:rPr>
          <w:lang/>
        </w:rPr>
        <w:instrText xml:space="preserve">
          <xsl:value-of select="$CrossReferences"/>
        </w:instrText>
      </w:r>
      <w:r>
        <w:rPr>
          <w:lang/>
        </w:rPr>
        <w:fldChar w:fldCharType="separate"/>
      </w:r>
      <w:r>
        <!--xsl:call-template name="InsertRunProperties"/-->
        <w:t>
          <xsl:value-of select="."/>
        </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:lang/>
        </w:rPr>
        <w:fldChar w:fldCharType="end"/>
      </w:r>
    </xsl:if>
  </xsl:template>


  <!-- bookmark start mark for elements contained in TOC -->
  <xsl:template name="InsertBookmarkStartTOC">
    <xsl:param name="tocId"/>
    <xsl:param name="tableOfContent"/>

    <w:bookmarkStart w:id="{$tocId}" w:name="{concat('_Toc',$tocId,generate-id($tableOfContent))}"/>
  </xsl:template>

  <!-- bookmark end mark for elements contained in TOC -->
  <xsl:template name="InsertBookmarkEndTOC">
    <xsl:param name="tocId"/>

    <w:bookmarkEnd w:id="{$tocId}"/>
  </xsl:template>

  <!-- checks if element has style or element used to generate TOC -->
  <xsl:template name="IsTOCStyleOrElement">
    <xsl:param name="sourceStyleNum"/>
    <xsl:param name="styleName"/>
    <xsl:param name="tableOfContent"/>

    <xsl:choose>

      <!-- empty elements are never bookmarked -->
      <xsl:when test="not(child::node())">false</xsl:when>

      <!--content of index body is never bookmarked-->
      <xsl:when test="ancestor::text:index-body">false</xsl:when>

      <!--checks if headings are used to generate TOC -->
      <xsl:when test="self::text:h">
        <xsl:choose>
          <!-- headings aren't used -->
          <xsl:when
            test="$tableOfContent/text:table-of-content-source/@text:use-outline-level = 'false' "
            >false</xsl:when>
          <!-- check is current heading level is used to generate TOC -->
          <xsl:otherwise>
            <xsl:choose>
              <xsl:when
                test="@text:outline-level &lt; ($tableOfContent/text:table-of-content-source/@text:outline-level+1)">
                <xsl:text>true</xsl:text>
              </xsl:when>
              <xsl:otherwise>false</xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <!-- checks if entries are to be included in TOC-->
      <xsl:when test="self::text:toc-mark-start">
        <xsl:choose>
          <xsl:when
            test="$tableOfContent/text:table-of-content-source/@text:use-index-marks = 'false' "
            >false</xsl:when>
          <xsl:otherwise>true</xsl:otherwise>
        </xsl:choose>
      </xsl:when>

      <!--  checks if style or parent style is used as a source style for TOC-->
      <xsl:when test="$sourceStyleNum > 0">
        <xsl:variable name="sourceStyleName"
          select="$tableOfContent/text:table-of-content-source/text:index-source-styles[$sourceStyleNum]/text:index-source-style/@text:style-name"/>

        <xsl:choose>
          <xsl:when
            test="(key('automatic-styles',$styleName) and key('automatic-styles',$styleName)/@style:parent-style-name = $sourceStyleName) 
            or $styleName = $sourceStyleName"
            >true</xsl:when>

          <!--  checks next source style-->
          <xsl:otherwise>
            <xsl:call-template name="IsTOCStyleOrElement">
              <xsl:with-param name="sourceStyleNum" select="$sourceStyleNum - 1"/>
              <xsl:with-param name="styleName" select="$styleName"/>
              <xsl:with-param name="tableOfContent" select="$tableOfContent"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>false</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--calculate bookmark id for element contained in TOC -->
  <xsl:template name="CalculateBookmarkId">
    <xsl:param name="counter"/>
    <xsl:param name="tableOfContent"/>
    <xsl:param name="sourceStyleNum"
      select="count($tableOfContent/text:table-of-content-source/text:index-source-styles)"/>

    <xsl:choose>

      <!--after counting source styles elements add headings number up to proper level defined in TOC and entry marks -->
      <xsl:when test="$sourceStyleNum = 0">
        <xsl:value-of
          select="$counter 
          + count(preceding::text:h[child::node() and not(ancestor::text:index-body) 
          and @text:outline-level &lt; ($tableOfContent/text:table-of-content-source/@text:outline-level+1)])
          + count(preceding::text:toc-mark-start[$tableOfContent/text:table-of-content-source/@text:use-index-marks != 'false' ])"
        />
      </xsl:when>

      <!--count element with source styles-->
      <xsl:when test="$sourceStyleNum > 0">
        <xsl:variable name="sourceStyleName"
          select="$tableOfContent/text:table-of-content-source/text:index-source-styles[$sourceStyleNum]/text:index-source-style/@text:style-name"/>
        <xsl:variable name="elementSum">
          <xsl:value-of
            select="$counter + count(preceding::text:p[@text:style-name = $sourceStyleName and child::node() and not(ancestor::text:index-body)]) +
            count(preceding::text:p[key('automatic-styles',@text:style-name)/@style:parent-style-name = $sourceStyleName and child::node() and not(ancestor::text:index-body)])"
          />
        </xsl:variable>

        <xsl:call-template name="CalculateBookmarkId">
          <xsl:with-param name="sourceStyleNum" select="$sourceStyleNum - 1"/>
          <xsl:with-param name="counter" select="$elementSum"/>
          <xsl:with-param name="tableOfContent" select="$tableOfContent"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$counter"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>



  <!-- index bookmark-->
  <xsl:template name="InsertIndexOfFiguresBookmark">

    <xsl:variable name="textName" select="@text:name"/>
    <xsl:variable name="id">
      <xsl:value-of select="number(count(preceding::text:sequence[@text:name = $textName]))+1"/>
    </xsl:variable>
    <xsl:variable name="indexOfObjects"
      select="key('indexes','')[child::*/@text:caption-sequence-name = $textName]"/>

    <w:bookmarkStart w:id="{$id}" w:name="{concat('_Toc',$id,generate-id($indexOfObjects))}"/>
    <w:r>
      <w:t>
        <xsl:value-of select="."/>
      </w:t>
    </w:r>
    <w:bookmarkEnd w:id="{$id}"/>
  </xsl:template>



</xsl:stylesheet>

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
  xmlns:xlink="http://www.w3.org/1999/xlink"
  exclude-result-prefixes="text fo style office draw">

	<!-- divo/20081008 xsl:strip-space must only be defined once in odf2oox.xls -->
	<!--<xsl:strip-space elements="*"/>
	<xsl:preserve-space elements="text:p"/>
	<xsl:preserve-space elements="text:span"/>-->

  <!-- key to find bookmarkt or reference-mark -->
  <xsl:key name="bookmark-reference-start"
    match="text:bookmark|text:bookmark-start|text:reference-mark-start|text:sequence"
    use="@text:name|@text:ref-name"/>

  <!-- 
  *************************************************************************
  MATCHING TEMPLATES
  *************************************************************************
  -->

  <!-- Insert BookmarkStart or ReferenceMarkStart-->
  <xsl:template match="text:bookmark-start | text:reference-mark-start | text:bookmark" mode="paragraph">
    <w:bookmarkStart>
      <xsl:attribute name="w:id">
        <xsl:call-template name="GenerateBookmarkId">
          <xsl:with-param name="TextName">
            <xsl:value-of select="@text:name"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="w:name">
        <xsl:call-template name="SuppressForbiddenChars">
          <xsl:with-param name="string" select="@text:name"/>
        </xsl:call-template>
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

  <!-- Insert BookmarkEnd -->
  <xsl:template match="text:bookmark-end | text:reference-mark-end" mode="paragraph">
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
  <xsl:template match="text:bookmark-ref | text:reference-ref | text:sequence-ref" mode="paragraph">
    <xsl:variable name="TextName" select="@text:ref-name"/>
    <xsl:variable name="masterPage"
      select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page/style:header/text:p"/>
    <xsl:if
      test="key('bookmark-reference-start', $TextName) or $masterPage/text:reference-mark-start[@text:name=$TextName] or $masterPage/text:bookmark-start[@text:name=$TextName]">
      <w:r>
        <w:fldChar w:fldCharType="begin"/>
      </w:r>
      <w:r>
        <xsl:call-template name="InsertRunProperties"/>
        <xsl:call-template name="InsertCrossReferences">
          <xsl:with-param name="TextName" select="$TextName"/>
        </xsl:call-template>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType="separate"/>
      </w:r>
      <w:r>
        <xsl:call-template name="InsertRunProperties"/>
        <xsl:apply-templates mode="text"/>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType="end"/>
      </w:r>
    </xsl:if>
  </xsl:template>

  <!-- 
  *************************************************************************
  CALLED TEMPLATES
  *************************************************************************
  -->

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


  <!--
  Summary:  Determines the outline level out of the given toc link
  Author:   makz (DIaLOGIKa)
  Params:   href: The href value of the toc link
  -->
  <xsl:template name="DetermineOutlineLevel">
    <xsl:param name="href" />
    <xsl:param name="outlineLevel" select="0" />

    <xsl:choose>
      <xsl:when test="starts-with($href, '#')">
        <xsl:call-template name="DetermineOutlineLevel">
          <xsl:with-param name="href" select="substring-after($href, '#')" />
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="substring-before($href, '.') != '' and number(substring-before($href, '.')) != 'NaN'">
        <xsl:call-template name="DetermineOutlineLevel">
          <xsl:with-param name="href" select="substring-after($href, '.')" />
          <xsl:with-param name="outlineLevel" select="$outlineLevel + 1" />
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$outlineLevel" />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <!--
  Summary:  Determines the referenced text out of the given toc link
  Author:   makz (DIaLOGIKa)
  Params:   href: The href value of the toc link
  -->
  <xsl:template name="DetermineReferencedText">
    <xsl:param name="href" />

    <xsl:choose>
      <xsl:when test="starts-with($href, '#')">
        <xsl:call-template name="DetermineReferencedText">
          <xsl:with-param name="href" select="substring-before(substring-after($href, '#'), '|')" />
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="substring-before($href, '.') != '' and number(substring-before($href, '.')) != 'NaN'">
        <xsl:call-template name="DetermineReferencedText">
          <xsl:with-param name="href" select="substring-after($href, '.')" />
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$href" />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <xsl:template name="InsertTOCBookmark">
    <xsl:param name="tableOfContentsNum" select="$tocCount"/>
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

          <!--
          makz: 
          The template CalculateBookmarkId caused heavy performance problems 
          and was replaced by the new toc/link mechanism below:
          
          <xsl:variable name="bookmarkId">
            <xsl:call-template name="CalculateBookmarkId">
              <xsl:with-param name="counter" select="1"/>
              <xsl:with-param name="tableOfContent" select="key('toc', '')[$tableOfContentsNum]"/>
            </xsl:call-template>
          </xsl:variable>
          -->
          
          <xsl:variable name="myToc" select="key('toc', '')[number($tableOfContentsNum)]" />
          <xsl:variable name="myTocId" select="generate-id($myToc)" />
          
          <!-- 
          Try to find the hyperlink that references this heading.
          We need the position of the paragraph that holds this hyperlink.
          -->
          <xsl:variable name="linkNr">

            <xsl:variable name="myOutlineLevel" select="@text:outline-level" />
            <xsl:variable name="myText" select="child::node()[1]" />
            
            <!--
            Check every paragraph in my TOC
            -->
            <xsl:for-each select="$myToc/text:index-body/text:p">

              <!--
              Determine the outline level out of the href
              -->
              <xsl:variable name="outlinelvl">
                <xsl:call-template name="DetermineOutlineLevel">
                  <xsl:with-param name="href"  select="text:a[1]/@xlink:href" />
                </xsl:call-template>
              </xsl:variable>

              <!--
              Determine the referenced text out of the href
              -->
              <xsl:variable name="refText">
                <xsl:call-template name="DetermineReferencedText">
                  <xsl:with-param name="href"  select="text:a[1]/@xlink:href" />
                </xsl:call-template>
              </xsl:variable>

              <!--
              If this headings outline level and text match, this is the link that references this heading
              -->
              <xsl:if test="$myOutlineLevel = $outlinelvl and $myText = $refText">
                <xsl:value-of select="position()"/>
              </xsl:if>

            </xsl:for-each>
          </xsl:variable>


          <xsl:call-template name="InsertBookmarkStartTOC">
            <xsl:with-param name="linkNr" select="$linkNr"/>
            <xsl:with-param name="tocId" select="$myTocId"/>
          </xsl:call-template>

          <xsl:call-template name="InsertBookmarkEndTOC">
            <xsl:with-param name="linkNr" select="$linkNr"/>
            <xsl:with-param name="tocId" select="$myTocId"/>
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
      <xsl:value-of
        select="count(key('bookmark-reference-start', $TextName)/preceding-sibling::text:reference-mark-start )+count(key('bookmark-reference-start', $TextName)/parent::text:p/preceding-sibling::node()/text:reference-mark-start )"
      />
    </xsl:variable>
    <xsl:variable name="BookmarkStart">
      <xsl:value-of
        select="count(key('bookmark-reference-start', $TextName)/preceding-sibling::text:bookmark-start)+count(key('bookmark-reference-start', $TextName)/parent::text:p/preceding-sibling::node()/text:bookmark-start)"
      />
    </xsl:variable>
    <xsl:variable name="Bookmark">
      <xsl:value-of
        select="count(key('bookmark-reference-start', $TextName)/preceding-sibling::text:bookmark)+count(key('bookmark-reference-start', $TextName)/parent::text:p/preceding-sibling::node()/text:bookmark)"
      />
    </xsl:variable>
    <xsl:value-of select="$ReferenceMarkStart+$BookmarkStart+$Bookmark"/>
  </xsl:template>

  <!-- insert reference field -->
  <xsl:template name="InsertCrossReferences">
    <xsl:param name="TextName"/>
    <!-- field type -->
    <xsl:choose>
      <xsl:when test="@text:reference-format='page' ">
        <w:instrText xml:space="preserve">PAGEREF \*MERGEFORMAT </w:instrText>
      </xsl:when>
      <xsl:otherwise>
        <w:instrText xml:space="preserve">REF \*MERGEFORMAT </w:instrText>
      </xsl:otherwise>
    </xsl:choose>
    <w:instrText>
      <xsl:choose>
        <xsl:when test="../text:sequence-ref[@text:ref-name=$TextName]">
          <xsl:variable name="indexOfObjects"
            select="generate-id(key('indexes','')[child::*/@text:caption-sequence-name = key('bookmark-reference-start', $TextName)/@text:name])"/>
          <xsl:value-of
            select="concat('_Toc', number(count(key('bookmark-reference-start', $TextName)/preceding::text:sequence))+1, $indexOfObjects)"
          />
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="SuppressForbiddenChars">
            <xsl:with-param name="string" select="$TextName"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </w:instrText>

    <xsl:choose>
      <xsl:when test="@text:reference-format='direction' ">
        <w:instrText xml:space="preserve"> \p</w:instrText>
      </xsl:when>
      <xsl:when test="@text:reference-format='chapter' ">
        <w:instrText xml:space="preserve"> \n</w:instrText>
      </xsl:when>
    </xsl:choose>

    <w:instrText xml:space="preserve"> \h </w:instrText>
  </xsl:template>

  <!-- bookmark start mark for elements contained in TOC -->
  <xsl:template name="InsertBookmarkStartTOC">
    <xsl:param name="linkNr"/>
    <xsl:param name="tocId"/>
    
    <w:bookmarkStart w:id="{concat('http://www.dialogika.de/replace/bookmarkid/', $tocId, $linkNr)}"
                     w:name="{concat('Toc_', $tocId, '_', $linkNr)}" />
  </xsl:template>

  <!-- bookmark end mark for elements contained in TOC -->
  <xsl:template name="InsertBookmarkEndTOC">
    <xsl:param name="linkNr"/>
    <xsl:param name="tocId"/>

    <w:bookmarkEnd w:id="{concat('http://www.dialogika.de/replace/bookmarkid/', $tocId, $linkNr)}" />
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
    <xsl:param name="sourceStyleNum" select="count($tableOfContent/text:table-of-content-source/text:index-source-styles)"/>

    <xsl:choose>

      <!--after counting source styles elements add headings number up to proper level defined in TOC and entry marks -->
      <xsl:when test="$sourceStyleNum = 0">
        <xsl:value-of
          select="$counter 
          + count(preceding::text:h[child::node() and not(ancestor::text:index-body) 
          and @text:outline-level &lt; ($tableOfContent/text:table-of-content-source/@text:outline-level+1)])
          + count(preceding::text:toc-mark-start[$tableOfContent/text:table-of-content-source/@text:use-index-marks != 'false' ])" />
      </xsl:when>

      <!--count element with source styles-->
      <xsl:when test="$sourceStyleNum > 0">
        <xsl:variable name="sourceStyleName"
          select="$tableOfContent/text:table-of-content-source/text:index-source-styles[$sourceStyleNum]/text:index-source-style/@text:style-name"/>
        <xsl:variable name="elementSum">
          <xsl:value-of
            select="$counter + count(preceding::text:p[@text:style-name = $sourceStyleName and child::node() and not(ancestor::text:index-body)]) +
            count(preceding::text:p[key('automatic-styles',@text:style-name)/@style:parent-style-name = $sourceStyleName and child::node() and not(ancestor::text:index-body)])" />
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
      select="generate-id(key('indexes','')[child::*/@text:caption-sequence-name = $textName])"/>

    <w:bookmarkStart w:id="{$id}" w:name="{concat('_Toc', $id, $indexOfObjects)}"/>
    <w:r>
      <w:t>
        <xsl:value-of select="."/>
      </w:t>
    </w:r>
    <w:bookmarkEnd w:id="{$id}"/>
  </xsl:template>

  <!-- compute a string to get an acceptable sting in OOX -->
  <xsl:template name="SuppressForbiddenChars">
    <xsl:param name="string"/>

    <xsl:variable name="newString">
      <xsl:choose>
        <xsl:when test="contains($string, ' ')">
          <xsl:value-of select="translate($string, ' ', '_')"/>
        </xsl:when>
        <xsl:when test="contains($string, ',')">
          <xsl:value-of select="translate($string, ',', '')"/>
        </xsl:when>
        <xsl:when test="contains($string, '?')">
          <xsl:value-of select="translate($string, '?', '')"/>
        </xsl:when>
        <xsl:when test="contains($string, ';')">
          <xsl:value-of select="translate($string, ';', '')"/>
        </xsl:when>
        <xsl:when test="contains($string, '.')">
          <xsl:value-of select="translate($string, '.', '')"/>
        </xsl:when>
        <xsl:when test="contains($string, ':')">
          <xsl:value-of select="translate($string, ':', '')"/>
        </xsl:when>
        <xsl:when test="contains($string, '/')">
          <xsl:value-of select="translate($string, '/', '')"/>
        </xsl:when>
        <xsl:when test="contains($string, '!')">
          <xsl:value-of select="translate($string, '!', '')"/>
        </xsl:when>
        <xsl:when test="contains($string, '§')">
          <xsl:value-of select="translate($string, '§', '')"/>
        </xsl:when>
        <xsl:when test="contains($string, '&quot;')">
          <xsl:value-of select="translate($string, '&quot;', '')"/>
        </xsl:when>
        <xsl:otherwise>clean</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="$newString = 'clean' ">
        <xsl:value-of select="$string"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="SuppressForbiddenChars">
          <xsl:with-param name="string" select="$newString"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

</xsl:stylesheet>

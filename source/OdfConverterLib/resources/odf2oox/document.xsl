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
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  exclude-result-prefixes="office text table fo style draw xlink v svg number"
  xmlns:w10="urn:schemas-microsoft-com:office:word">

  <xsl:import href="tables.xsl"/>

  <xsl:strip-space elements="*"/>
  <xsl:preserve-space elements="text:p"/>
  <xsl:preserve-space elements="text:span"/>

  <xsl:key name="annotations" match="office:annotation" use="''"/>
  <xsl:key name="automatic-styles" match="office:automatic-styles/style:style" use="@style:name"/>
  <xsl:key name="hyperlinks" match="text:a" use="''"/>
  <xsl:key name="headers" match="text:h" use="''"/>
  <xsl:key name="images"
    match="draw:frame[not(./draw:object-ole or ./draw:object) and ./draw:image/@xlink:href]"
    use="''"/>
  <xsl:key name="ole-objects" match="draw:frame[./draw:object-ole] " use="''"/>
  <xsl:key name="master-pages" match="style:master-page" use="@style:name"/>
  <xsl:key name="page-layouts" match="style:page-layout" use="@style:name"/>
  <xsl:key name="master-based-styles" match="style:style[@style:master-page-name]" use="@style:name"/>
  <xsl:key name="sections" match="style:style[@style:family='section']" use="@style:name"/>
  <xsl:key name="restarting-lists" match="text:list[text:list-item/@text:start-value]" use="''"/>
  <xsl:key name="toc" match="text:table-of-content" use="''"/>
  <xsl:key name="date-style" match="number:date-style" use="@style:name"/>
  <!-- key to find mailto hyperlinks. -->
  <xsl:key name="mailto-hyperlinks" match="text:a[contains(@xlink:href,'mailto')]"
    use="text:span/@text:style-name"/>


  <xsl:variable name="body" select="document('content.xml')/office:document-content/office:body"/>
  <!-- Set of text elements potentially tied to a master style -->
  <xsl:variable name="elts"
    select="$body/descendant::*[name()='text:p' or name() = 'text:h' or name() = 'table:table']"/>
  <!-- Text elements tied to a master style. 
    (check for empty @master-page-name values - happens with OpenOffice -->
  <xsl:variable name="master-elts"
    select="$elts[key('master-based-styles', @text:style-name|@table:style-name)[1]/@style:master-page-name != '' ]"/>
  <!-- Default master style -->
  <xsl:variable name="default-master-style"
    select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page[1]"/>
  <!-- The very first text element -->
  <xsl:variable name="first-elt" select="$elts[1]"/>
  <!-- table of content count -->
  <xsl:variable name="tocCount" select="$body//text:table-of-content"/>

  <!-- main document -->
  <xsl:template name="document">
    <w:document>
      <!-- Translate the default master page color as the document background color -->
      <xsl:for-each select="document('styles.xml')">
        <xsl:variable name="defaultBgColor"
          select="key('page-layouts', $default-master-style/@style:page-layout-name)[1]/style:page-layout-properties/@fo:background-color"/>
        <xsl:if test="$defaultBgColor != 'transparent'">
          <w:background w:color="{translate(substring-after($defaultBgColor,'#'),'f','F')}"/>
        </xsl:if>
      </xsl:for-each>
      <xsl:apply-templates select="document('content.xml')/office:document-content"/>
    </w:document>
  </xsl:template>

  <!-- document body -->
  <xsl:template match="office:body">
    <w:body>
      <xsl:apply-templates/>
      <xsl:call-template name="InsertDocumentFinalSectionProperties"/>
    </w:body>
  </xsl:template>

  <!-- Document final section properties -->
  <xsl:template name="InsertDocumentFinalSectionProperties">
    <w:sectPr>
      <!-- Last element tied to a master-style -->
      <xsl:variable name="last-elt" select="$master-elts[last()]"/>
      <!-- Its master page name -->
      <xsl:variable name="master-page-name"
        select="key('master-based-styles', $last-elt/@text:style-name | $last-elt/@table:style-name)[1]/@style:master-page-name"/>
      <!--  Continuous section. Looking up for a text:section 
        If there's no master-page used after the last text:section, then the sectPr is continuous. -->
      <xsl:variable name="last-section" select="descendant::text:section[last()]"/>
      <xsl:variable name="elts-after-section"
        select="$last-section/following::*[name() = 'text:p' or name() = 'text:h' or name='table:table']"/>
      <xsl:variable name="master-elt-after-section"
        select="$elts-after-section[key('master-based-styles', @text:style-name|@table:style-name)[1]/@style:master-page-name != '' ]"/>
      <xsl:variable name="continuous">
        <xsl:choose>
          <xsl:when test="$last-section and not($master-elt-after-section)">yes</xsl:when>
          <xsl:otherwise>no</xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:choose>
        <!-- if we use a master-page based style -->
        <xsl:when test="$master-page-name">
          <xsl:call-template name="InsertSectionProperties">
            <xsl:with-param name="continuous" select="$continuous"/>
            <xsl:with-param name="elt" select="$last-elt"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <!-- use default master page -->
          <xsl:call-template name="InsertSectionProperties">
            <xsl:with-param name="continuous" select="$continuous"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </w:sectPr>
  </xsl:template>

  <!-- OOX section properties (header/footer, footnotes/endnotes, page layout, etc) 
       param elt : specifies the element whose style references the master page to be transformed in OOX section properties
       param continuous : specifies if this OOX section should be continuous
       param section-ends : specifies if this OOX section corresponds to an ending ODF section
       param previous-section : specifies the ancestor ODF section 
   -->
  <xsl:template name="InsertSectionProperties">
    <xsl:param name="elt"/>
    <xsl:param name="continuous" select="'no'"/>
    <xsl:param name="notes-configuration"/>
    <xsl:param name="section-ends"/>
    <xsl:param name="previous-section"/>

    <xsl:choose>
      <xsl:when test="$elt">
        <xsl:variable name="eltstyle"
          select="key('master-based-styles', $elt/@text:style-name | $elt/@table:style-name)[1]"/>
        <xsl:if test="$continuous = 'no' ">
          <!-- header/footer -->
          <!-- Is it the first time we use this master style? In which case we have to reference the header/footer -->
          <xsl:choose>
            <!-- If the very first text elt is not explicitely tied to a master style, then default master style was used - do nothing -->
            <xsl:when
              test="$eltstyle/@style:master-page-name = $default-master-style/@style:name and not(boolean(key('master-based-styles', $first-elt/@text:style-name | $first-elt/@table:style-name)[1]))"/>
            <xsl:otherwise>
              <!-- check the same master style has not been used before -->
              <!-- elts tied to the same master style -->
              <xsl:variable name="elt-siblings"
                select="$master-elts[key('master-based-styles', @text:style-name|@table:style-name)[1]/@style:master-page-name = $eltstyle/@style:master-page-name]"/>
              <xsl:variable name="first-occurrence" select="$elt-siblings[1]"/>
              <!-- the first occurrence of a text elt with the same master-page name matches the current elt => this is the first use of this master page -->
              <xsl:if test="generate-id($elt) = generate-id($first-occurrence)">
                <xsl:for-each select="document('styles.xml')">
                  <xsl:call-template name="HeaderFooter">
                    <xsl:with-param name="master-page"
                      select="key('master-pages', $eltstyle/@style:master-page-name)"/>
                  </xsl:call-template>
                </xsl:for-each>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>

        <!-- notes configuration -->
        <xsl:choose>
          <xsl:when test="$notes-configuration">
            <xsl:apply-templates select="$notes-configuration" mode="note"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:apply-templates
              select="document('styles.xml')/office:document-styles/office:styles/text:notes-configuration"
              mode="note"/>
          </xsl:otherwise>
        </xsl:choose>

        <!-- continuous -->
        <xsl:if test="$continuous = 'yes' ">
          <w:type w:val="continuous"/>
        </xsl:if>

        <!-- page layout properties -->
        <xsl:for-each select="document('styles.xml')">
          <xsl:apply-templates
            select="key('page-layouts', key('master-pages', $eltstyle/@style:master-page-name)/@style:page-layout-name)/style:page-layout-properties"
            mode="master-page"/>
        </xsl:for-each>
        <xsl:if test="$section-ends = 'true' ">
          <xsl:apply-templates
            select="key('sections', $previous-section/@text:style-name)/style:section-properties"
            mode="section"/>
        </xsl:if>
        <xsl:for-each select="document('styles.xml')">
          <!-- Shall the header and footer be different on the first page -->
          <xsl:call-template name="TitlePg">
            <xsl:with-param name="master-page"
              select="key('master-pages', $eltstyle/@style:master-page-name)[1]"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
      <xsl:otherwise>
        <!-- if we fall here, it means no master page has been explicitely used so far -->
        <!-- header/footer -->
        <xsl:if test="$continuous = 'no' ">
          <xsl:for-each select="document('styles.xml')">
            <xsl:call-template name="HeaderFooter">
              <xsl:with-param name="master-page"
                select=" /office:document-styles/office:master-styles/style:master-page[1]"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:if>

        <!-- notes configuration -->
        <xsl:choose>
          <xsl:when test="$notes-configuration">
            <xsl:apply-templates select="$notes-configuration" mode="note"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:apply-templates
              select="document('styles.xml')/office:document-styles/office:styles/text:notes-configuration"
              mode="note"/>
          </xsl:otherwise>
        </xsl:choose>

        <!-- continuous -->
        <xsl:if test="$continuous = 'yes' ">
          <w:type w:val="continuous"/>
        </xsl:if>

        <!-- page layou properties -->
        <xsl:for-each select="document('styles.xml')">
          <xsl:apply-templates
            select="key('page-layouts', /office:document-styles/office:master-styles/style:master-page[1]/@style:page-layout-name)/style:page-layout-properties"
            mode="master-page"/>
        </xsl:for-each>
        <xsl:if test="$section-ends">
          <xsl:apply-templates
            select="key('sections', $previous-section/@text:style-name)/style:section-properties"
            mode="section"/>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--checks if element has style used to generate table of contents in document  TODO bookmarks for index marks -->
  <xsl:template name="IsTOCBookmark">
    <xsl:param name="styleName"/>
    <xsl:param name="tableOfContentsNum" select="count(key('toc',''))"/>
    <xsl:param name="isTocStyle"/>

    <xsl:choose>
      <xsl:when test="$tocCount &lt; 1">false</xsl:when>
      <xsl:otherwise>
        <xsl:variable name="tableOfContent" select="key('toc', '')[$tableOfContentsNum]"/>
        <xsl:choose>
          <xsl:when test="$tableOfContentsNum > 0 and $isTocStyle != 'true'">
            <xsl:variable name="tocStyle">
              <xsl:call-template name="IsTOCStyle">
                <xsl:with-param name="sourceStyleNum"
                  select="count($tableOfContent/text:table-of-content-source/text:index-source-styles)"/>
                <xsl:with-param name="styleName" select="$styleName"/>
                <xsl:with-param name="tableOfContent" select="$tableOfContent"/>
              </xsl:call-template>
            </xsl:variable>

            <xsl:call-template name="IsTOCBookmark">
              <xsl:with-param name="tableOfContentsNum" select="$tableOfContentsNum - 1"/>
              <xsl:with-param name="styleName" select="$styleName"/>
              <xsl:with-param name="isTocStyle" select="$tocStyle"/>
            </xsl:call-template>
          </xsl:when>

          <xsl:otherwise>
            <xsl:choose>
              <xsl:when test="$isTocStyle = 'true'">true</xsl:when>
              <xsl:otherwise>false</xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>



  <xsl:template name="InsertTOCBookmark">
    <xsl:param name="tableOfContentsNum" select="count(key('toc',''))"/>
    <xsl:param name="bookmarkType"/>


    <xsl:choose>
      <xsl:when test="$tableOfContentsNum > 0">

        <xsl:variable name="bookmarkId">
          <xsl:call-template name="CalculateBookmarkId">
            <xsl:with-param name="counter" select="1"/>
            <xsl:with-param name="tableOfContent" select="key('toc', '')[$tableOfContentsNum]"/>
          </xsl:call-template>
        </xsl:variable>

        <xsl:choose>
          <xsl:when test="$bookmarkType = 'start' ">
            <xsl:call-template name="InsertBookmarkStartTOC">
              <xsl:with-param name="tocId" select="$bookmarkId"/>
              <xsl:with-param name="tableOfContent" select="key('toc', '')[$tableOfContentsNum]"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="$bookmarkType = 'end' ">
            <xsl:call-template name="InsertBookmarkEndTOC">
              <xsl:with-param name="tocId" select="$bookmarkId"/>
            </xsl:call-template>
          </xsl:when>
        </xsl:choose>

        <xsl:call-template name="InsertTOCBookmark">
          <xsl:with-param name="tableOfContentsNum" select="$tableOfContentsNum - 1"/>
          <xsl:with-param name="bookmarkType" select="$bookmarkType"/>
        </xsl:call-template>

      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- paragraphs and headings -->
  <xsl:template match="text:p | text:h">
    <xsl:param name="level" select="0"/>
    <xsl:message terminate="no">progress:text:p</xsl:message>
    <w:p>

      <w:pPr>
        <xsl:call-template name="InsertParagraphProperties">
          <xsl:with-param name="level" select="$level"/>
        </xsl:call-template>
        <xsl:call-template name="InsertParagraphSectionProperties"/>
      </w:pPr>

      <!-- check if element is contained in TOC  -->
      <xsl:variable name="isBookmarked">
        <xsl:call-template name="IsTOCBookmark">
          <xsl:with-param name="styleName" select="@text:style-name"/>
        </xsl:call-template>
      </xsl:variable>

      <!--   insert bookmark start for element which is contained in TOC-->
      <xsl:if test="$isBookmarked = 'true' ">
        <xsl:variable name="bookmarkType">start</xsl:variable>
        <xsl:call-template name="InsertTOCBookmark">
          <xsl:with-param name="bookmarkType" select="$bookmarkType"/>
        </xsl:call-template>
      </xsl:if>

      <!-- footnotes or endnotes: insert the mark in the first paragraph -->
      <xsl:if test="parent::text:note-body and position() = 1">
        <xsl:apply-templates select="../../text:note-citation" mode="note"/>
      </xsl:if>

      <!-- paragraph content-->
      <xsl:call-template name="InsertParagraphContent"/>

      <!-- If there is a page-break-after in the paragraph style -->
      <xsl:call-template name="InsertPageBreakAfter"/>

      <!--   insert bookmark end for element which is contained in TOC-->
      <xsl:if test="$isBookmarked = 'true' ">
        <xsl:variable name="bookmarkType">end</xsl:variable>
        <xsl:call-template name="InsertTOCBookmark">
          <xsl:with-param name="bookmarkType" select="$bookmarkType"/>
        </xsl:call-template>
      </xsl:if>

    </w:p>
  </xsl:template>

  <!-- section detection and insertion for paragraph-->
  <xsl:template name="InsertParagraphSectionProperties">
    <!-- Section detection  : 3 cases -->
    <xsl:if
      test="not(ancestor::table:table) and not(ancestor::draw:frame) and not(ancestor::style:master-page)">
      <!-- Section detection  : 3 cases -->
      <!-- 1 - Following neighbour's (ie paragraph, heading or table) with non-empty reference to a master page  -->
      <xsl:variable name="followings"
        select="following::text:p[1] | following::text:h[1] | following::table:table[1]"/>
      <xsl:variable name="master-page-starts"
        select="boolean(key('master-based-styles', $followings[1]/@text:style-name | $followings[1]/@table:style-name)[1]/@style:master-page-name != '')"/>

      <!-- 2 - Section starts. The following paragraph is contained in the following section -->
      <xsl:variable name="following-section" select="following::text:section[1]"/>
      <!-- the following section is the same as the following neighbour's ancestor section -->
      <xsl:variable name="section-starts"
        select="$following-section and (generate-id($followings[1]/ancestor::text:section[1]) = generate-id($following-section))"/>

      <!-- 3 - Section ends. We are in a section and the following paragraph isn't -->
      <xsl:variable name="previous-section" select="ancestor::text:section[1]"/>
      <!-- the following neighbour's ancestor section and the current section are different -->
      <xsl:variable name="section-ends"
        select="$previous-section and not(generate-id($followings[1]/ancestor::text:section[1]) = generate-id($previous-section))"/>

      <!-- section creation -->
      <xsl:if
        test="($master-page-starts = 'true' or $section-starts = 'true' or $section-ends = 'true') and not(ancestor::text:note-body) and not(ancestor::table:table)">
        <w:sectPr>
          <!--  Continuous sections. Looking up for a text:section 
            If the first master style following the preceding section is the same as this paragraph's following master-style,
            then no other master style is used in-between. -->
          <xsl:variable name="ps" select="preceding::text:section[1]"/>
          <xsl:variable name="styles-after-section"
            select="$ps/following::text:p[key('master-based-styles', @text:style-name)[1]/@style:master-page-name != ''] | $ps/following::text:h[key('master-based-styles', @text:style-name)[1]/@style:master-page-name != ''] | $ps/following::text:table[key('master-based-styles', @table:style-name)[1]/@style:master-page-name != '']"/>
          <xsl:variable name="following-master-style"
            select="$followings[key('master-based-styles', @text:style-name|@table:style-name)]"/>
          <xsl:variable name="continuous">
            <xsl:choose>
              <xsl:when
                test="$section-ends or ($ps and (generate-id($styles-after-section[1]) = generate-id($following-master-style[1])))"
                >yes</xsl:when>
              <xsl:when test="ancestor::text:section[1]">yes</xsl:when>
              <xsl:otherwise>no</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>

          <!-- notes configuration -->
          <xsl:variable name="notes-configuration"
            select="key('sections', $previous-section/@text:style-name)[1]/style:section-properties/text:notes-configuration"/>

          <xsl:variable name="current-master-style"
            select="key('master-based-styles', @text:style-name)"/>

          <xsl:choose>
            <xsl:when test="boolean($current-master-style)">
              <!-- current element style is tied to a master page -->
              <xsl:call-template name="InsertSectionProperties">
                <xsl:with-param name="continuous" select="$continuous"/>
                <xsl:with-param name="elt" select="."/>
                <xsl:with-param name="notes-configuration" select="$notes-configuration"/>
                <xsl:with-param name="section-ends" select="$section-ends"/>
                <xsl:with-param name="previous-section" select="$previous-section"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <!-- current style is not tied to a master page (typically the case of an ODF section), find the preceding one -->
              <xsl:variable name="precedings"
                select="preceding::text:p[key('master-based-styles', @text:style-name)[1]/@style:master-page-name != ''] | preceding::text:h[key('master-based-styles', @text:style-name)[1]/@style:master-page-name != ''] | preceding::table:table[key('master-based-styles', @table:style-name)[1]/@style:master-page-name != '']"/>
              <xsl:variable name="preceding-master-style"
                select="key('master-based-styles', $precedings[last()]/@text:style-name | $precedings[last()]/@table:style-name)"/>
              <xsl:choose>
                <xsl:when test="boolean($preceding-master-style)">
                  <xsl:call-template name="InsertSectionProperties">
                    <xsl:with-param name="continuous" select="$continuous"/>
                    <xsl:with-param name="elt" select="$precedings[last()]"/>
                    <xsl:with-param name="notes-configuration" select="$notes-configuration"/>
                    <xsl:with-param name="section-ends" select="$section-ends"/>
                    <xsl:with-param name="previous-section" select="$previous-section"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <!-- otherwise, apply the default master style -->
                  <xsl:call-template name="InsertSectionProperties">
                    <xsl:with-param name="continuous" select="$continuous"/>
                    <xsl:with-param name="notes-configuration" select="$notes-configuration"/>
                    <xsl:with-param name="section-ends" select="$section-ends"/>
                    <xsl:with-param name="previous-section" select="$previous-section"/>
                  </xsl:call-template>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </w:sectPr>
      </xsl:if>
    </xsl:if>
  </xsl:template>

  <!-- conversion of paragraph content excluding not supported elements -->
  <xsl:template name="InsertParagraphContent">
    <xsl:choose>

      <!-- we are in table of contents -->
      <xsl:when test="parent::text:index-body">
        <xsl:call-template name="InsertIndexItem"/>
      </xsl:when>

      <!-- ignore draw:frame/draw:text-box if it's embedded in another draw:frame/draw:text-box becouse word doesn't support it -->
      <xsl:when test="self::node()[ancestor::draw:text-box and descendant::draw:text-box]">
        <xsl:message terminate="no">feedback:Nested frames</xsl:message>
      </xsl:when>
      
      <!-- drawing shapes -->
      <xsl:when test="child::draw:ellipse|child::draw:rect|child::draw:custom-shape">
        <xsl:apply-templates mode="shapes"/>
      </xsl:when>
      
      <xsl:otherwise>
        <xsl:apply-templates mode="paragraph"/>
      </xsl:otherwise>

    </xsl:choose>

  </xsl:template>

  <!-- inserts page-break-after if defined for paragraph  -->
  <xsl:template name="InsertPageBreakAfter">
    <xsl:if
      test="key('automatic-styles',@text:style-name)/style:paragraph-properties/@fo:break-after='page'">
      <w:r>
        <xsl:if test="ancestor::text:section/@text:display='none'">
          <w:rPr>
            <w:vanish/>
          </w:rPr>
        </xsl:if>
        <w:br w:type="page"/>
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

  <!-- checks if element has style used to generate TOC -->
  <xsl:template name="IsTOCStyle">
    <xsl:param name="sourceStyleNum"/>
    <xsl:param name="styleName"/>
    <xsl:param name="tableOfContent"/>

    <xsl:choose>

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

      <!-- empty elements are never bookmarked -->
      <xsl:when test="not(child::node())">false</xsl:when>

      <!--content of index body is never bookmarked-->
      <xsl:when test="ancestor::text:index-body">false</xsl:when>

      <xsl:when test="$sourceStyleNum > 0">
        <xsl:variable name="sourceStyleName"
          select="$tableOfContent/text:table-of-content-source/text:index-source-styles[$sourceStyleNum]/text:index-source-style/@text:style-name"/>

        <xsl:choose>
          <!--  checks if style or parent style is used as a source style for TOC-->
          <xsl:when
            test="(key('automatic-styles',$styleName) and key('automatic-styles',$styleName)/@style:parent-style-name = $sourceStyleName) 
            or $styleName = $sourceStyleName"
            >true</xsl:when>

          <!--  checks next source style-->
          <xsl:otherwise>
            <xsl:call-template name="IsTOCStyle">
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

      <!--after counting source styles elements add headings number up to proper level defined in TOC-->
      <xsl:when test="$sourceStyleNum = 0">
        <xsl:value-of
          select="$counter + count(preceding::text:h[child::node() and not(ancestor::text:index-body) and @text:outline-level &lt; ($tableOfContent/text:table-of-content-source/@text:outline-level+1)])"
        />
      </xsl:when>

      <!--count element with source styles-->
      <xsl:when test="$sourceStyleNum > 0">
        <xsl:variable name="sourceStyleName"
          select="//text:table-of-content/text:table-of-content-source/text:index-source-styles[$sourceStyleNum]/text:index-source-style/@text:style-name"/>
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

  <!-- Inserts the paragraph properties -->
  <xsl:template name="InsertParagraphProperties">
    <xsl:param name="level"/>

    <!-- insert paragraph style -->
    <xsl:call-template name="InsertParagraphStyle">
      <xsl:with-param name="styleName">
        <xsl:value-of select="@text:style-name"/>
      </xsl:with-param>
    </xsl:call-template>

    <!-- insert page break before table when required -->
    <xsl:call-template name="InsertPageBreakBefore"/>

    <!-- insert numbering properties -->
    <xsl:call-template name="InsertNumbering">
      <xsl:with-param name="level" select="."/>
    </xsl:call-template>

    <!-- override spacing before/after when required -->
    <xsl:call-template name="InsertParagraphSpacing"/>

    <!-- insert indentation if paragraph is in a list -->
    <xsl:call-template name="InsertIndent">
      <xsl:with-param name="level" select="$level"/>
    </xsl:call-template>

    <!-- insert heading outline level -->
    <xsl:call-template name="InsertOutlineLevel">
      <xsl:with-param name="node" select="."/>
    </xsl:call-template>


    <!-- if we are in an annotation, we may have to insert annotation reference -->
    <xsl:call-template name="InsertAnnotationReference"/>

    <!-- insert run properties -->
    <xsl:call-template name="InsertRunProperties"/>
  </xsl:template>

  <!-- Inserts the style of a paragraph -->
  <xsl:template name="InsertParagraphStyle">
    <xsl:param name="styleName"/>
    <xsl:variable name="prefixedStyleName">
      <xsl:call-template name="GetPrefixedStyleName">
        <xsl:with-param name="styleName" select="$styleName"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:if test="$prefixedStyleName != ''">
      <w:pStyle w:val="{$prefixedStyleName}"/>
    </xsl:if>
  </xsl:template>

  <!-- Inserts the outline level of a heading if needed -->
  <xsl:template name="InsertOutlineLevel">
    <xsl:param name="node"/>
    <!-- List item are first considered if exist than heading style-->
    <xsl:variable name="stylename" select="./@text:style-name"/>
    <xsl:variable name="list"
      select="//office:document-content/office:automatic-styles/style:style[@style:name = $stylename]"/>
    <xsl:if
      test="$node[self::text:h and not($list/@style:list-style-name) and (not(parent::text:list-item) or position() > 1) ]">
      <xsl:choose>
        <xsl:when test="not($node/@text:outline-level)">
          <w:outlineLvl w:val="0"/>
        </xsl:when>
        <xsl:when test="$node/@text:outline-level &lt;= 9">
          <w:outlineLvl w:val="{$node/@text:outline-level}"/>
        </xsl:when>
        <xsl:otherwise>
          <w:outlineLvl w:val="9"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <xsl:if test="$node[child::text:toc-mark-start]">
      <w:outlineLvl w:val="{$node/text:toc-mark-start/@text:outline-level}"/>
    </xsl:if>
  </xsl:template>



  <!-- Insert spacing in paragraph properties if table before/after w:p element has spacing after/before -->
  <xsl:template name="InsertParagraphSpacing">
    <xsl:if
      test="following-sibling::node()[1][name()='table:table'] or preceding-sibling::node()[1][name()='table:table']">
      <!-- Compute space after -->
      <xsl:variable name="spaceAfter">
        <xsl:call-template name="CompareSpacingValues">
          <xsl:with-param name="tableSide" select="'top'"/>
          <xsl:with-param name="paraSide" select="'bottom'"/>
        </xsl:call-template>
      </xsl:variable>
      <!-- Compute space before -->
      <xsl:variable name="spaceBefore">
        <xsl:call-template name="CompareSpacingValues">
          <xsl:with-param name="tableSide" select="'bottom'"/>
          <xsl:with-param name="paraSide" select="'top'"/>
        </xsl:call-template>
      </xsl:variable>
      <!-- override if needed -->
      <xsl:if test="$spaceBefore &gt; 0 or $spaceAfter &gt; 0">
        <w:spacing>
          <xsl:if test="$spaceBefore &gt; 0">
            <xsl:attribute name="w:before">
              <xsl:value-of select="$spaceBefore"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="$spaceAfter &gt; 0">
            <xsl:attribute name="w:after">
              <xsl:value-of select="$spaceAfter"/>
            </xsl:attribute>
          </xsl:if>
        </w:spacing>
      </xsl:if>
    </xsl:if>
  </xsl:template>

  <!-- compare values of spacing before/after and return 0 or the spacing value to override -->
  <xsl:template name="CompareSpacingValues">
    <xsl:param name="tableSide"/>
    <xsl:param name="paraSide"/>

    <!-- get spacing value of table properties -->
    <xsl:variable name="tableSpace">
      <xsl:choose>
        <xsl:when
          test="$tableSide='top' and key('automatic-styles',following-sibling::node()[1][name()='table:table']/@table:style-name)/style:table-properties/attribute::node()[name()=concat('fo:margin-',$tableSide)]">
          <xsl:call-template name="twips-measure">
            <xsl:with-param name="length"
              select="key('automatic-styles',following-sibling::node()[1][name()='table:table']/@table:style-name)/style:table-properties/attribute::node()[name()=concat('fo:margin-',$tableSide)]"
            />
          </xsl:call-template>
        </xsl:when>
        <xsl:when
          test="$tableSide='bottom' and key('automatic-styles',preceding-sibling::node()[1][name()='table:table']/@table:style-name)/style:table-properties/attribute::node()[name()=concat('fo:margin-',$tableSide)]">
          <xsl:call-template name="twips-measure">
            <xsl:with-param name="length"
              select="key('automatic-styles',preceding-sibling::node()[1][name()='table:table']/@table:style-name)/style:table-properties/attribute::node()[name()=concat('fo:margin-',$tableSide)]"
            />
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <!-- if spacing value of table is 0, do not override. -->
    <xsl:choose>
      <xsl:when test="$tableSpace != 0">
        <!-- get spacing value of paragraph style -->
        <xsl:variable name="paraSpace">
          <xsl:choose>
            <xsl:when
              test="key('automatic-styles',@text:style-name)/style:paragraph-properties/attribute::node()[name()=concat('fo:margin-',$paraSide)]">
              <xsl:call-template name="twips-measure">
                <xsl:with-param name="length"
                  select="key('automatic-styles',@text:style-name)/style:paragraph-properties/attribute::node()[name()=concat('fo:margin-',$paraSide)]"
                />
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:for-each select="document('styles.xml')">
                <xsl:choose>
                  <xsl:when
                    test="key('styles',@text:style-name)/style:paragraph-properties/attribute::node()[name()=concat('fo:margin-',$paraSide)]">
                    <xsl:call-template name="twips-measure">
                      <xsl:with-param name="length"
                        select="key('styles',@text:style-name)/style:paragraph-properties/attribute::node()[name()=concat('fo:margin-',$paraSide)]"
                      />
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:otherwise>0</xsl:otherwise>
                </xsl:choose>
              </xsl:for-each>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <!-- compare those two values and choose which one is best -->
        <xsl:choose>
          <xsl:when test="$tableSpace &gt; $paraSpace">
            <xsl:value-of select="$tableSpace"/>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>0</xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <!-- Computes the style name to be used be InsertIndent template -->
  <xsl:template name="GetStyleName">
    <xsl:choose>
      <xsl:when test="self::text:list-item">
        <xsl:value-of select="*[1][self::text:p]/@text:style-name"/>
      </xsl:when>
      <xsl:when
        test="parent::text:list-header|self::text:p|self::text:h|self::text:list-level-style-number">
        <xsl:value-of select="@text:style-name"/>
      </xsl:when>
      <xsl:when test="parent::text:p|parent::text:h">
        <xsl:value-of select="parent::node()/@text:style-name"/>
      </xsl:when>
      <xsl:when test="ancestor::text:span">
        <xsl:value-of select="ancestor::text:span[1]/@text:style-name"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="@text:style-name"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Inserts item for all types of index  -->
  <xsl:template name="InsertIndexItem">

    <xsl:variable name="indexElementPosition">
      <xsl:number/>
    </xsl:variable>

    <!-- inserts field code of index to first index element -->
    <xsl:if test="$indexElementPosition = 1">
      <xsl:call-template name="InsertIndexFieldCodeStart"/>
    </xsl:if>

    <xsl:choose>

      <!-- when hyperlink option is on in TOC -->
      <xsl:when test="text:a">
        <xsl:apply-templates select="text:a" mode="paragraph"/>
      </xsl:when>

      <!--default scenario-->
      <xsl:otherwise>
        <xsl:call-template name="InsertIndexItemContent"/>
      </xsl:otherwise>

    </xsl:choose>

    <!-- inserts field code end in last index element -->
    <xsl:if test="(count(following-sibling::text:p) = 0) and parent::text:index-body">
      <xsl:call-template name="InsertIndexFieldCodeEnd"/>
    </xsl:if>

  </xsl:template>

  <xsl:template name="InsertIndexFieldCodeEnd">
    <w:r>
      <xsl:if test="ancestor::text:section/@text:display='none'">
        <w:rPr>
          <w:vanish/>
        </w:rPr>
      </xsl:if>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
  </xsl:template>

  <xsl:template name="InsertIndexFieldCodeStart">
    <w:r>
      <xsl:if test="ancestor::text:section/@text:display='none'">
        <w:rPr>
          <w:vanish/>
        </w:rPr>
      </xsl:if>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
      <xsl:if test="ancestor::text:section/@text:display='none'">
        <w:rPr>
          <w:vanish/>
        </w:rPr>
      </xsl:if>
      <xsl:choose>
        <xsl:when test="ancestor::text:table-of-content">
          <xsl:call-template name="InsertTocPrefs"/>
        </xsl:when>
        <xsl:when test="ancestor::text:illustration-index">
          <xsl:call-template name="InsertIllustrationInPrefs"/>
        </xsl:when>
        <xsl:when test="ancestor::text:alphabetical-index">
          <xsl:call-template name="insertAlphabeticalPrefs"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="InsertIndexFiguresPrefs"/>
        </xsl:otherwise>
      </xsl:choose>
    </w:r>
    <w:r>
      <xsl:if test="ancestor::text:section/@text:display='none'">
        <w:rPr>
          <w:vanish/>
        </w:rPr>
      </xsl:if>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
  </xsl:template>

  <xsl:template name="InsertIndexFiguresPrefs">
    <w:instrText xml:space="preserve">
      <xsl:text>TOC  \c "</xsl:text>
      <xsl:value-of select="parent::text:index-body/preceding-sibling::text:table-index-source/@text:caption-sequence-name"/>
      <xsl:text> " </xsl:text> 
    </w:instrText>
  </xsl:template>

  <xsl:template name="insertAlphabeticalPrefs">
    <w:instrText xml:space="preserve">
      <xsl:text>INDEX \e "" \c "</xsl:text>
      <!--xsl:choose>
        <xsl:when test="key('automatic-styles',ancestor::text:alphabetical-index/@text:style-name)/style:section-properties/style:columns/@fo:column-count=0">
          <xsl:text> 1 </xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="key('automatic-styles',ancestor::text:alphabetical-index/@text:style-name)/style:section-properties/style:columns/@fo:column-count"/>
        </xsl:otherwise>
      </xsl:choose-->
      <xsl:text>x " \z "1045" </xsl:text> 
    </w:instrText>
  </xsl:template>

  <xsl:template name="InsertIllustrationInPrefs">
    <w:instrText xml:space="preserve"> 
      <xsl:text> TOC  \c "</xsl:text>
     <xsl:value-of select="parent::text:index-body/preceding-sibling::text:illustration-index-source/@text:caption-sequence-name"/>" 
    </w:instrText>
  </xsl:template>

  <xsl:template name="InsertTocPrefs">
    <xsl:variable name="tocSource"
      select="ancestor::text:table-of-content/text:table-of-content-source"/>

    <w:instrText>
      <xsl:text>TOC \o "1-</xsl:text>

      <!-- include elements with outline styles up to selected level  -->
      <xsl:choose>
        <xsl:when test="$tocSource/@text:outline-level=10">
          <xsl:text>9"</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$tocSource/@text:outline-level"/>
          <xsl:text>"</xsl:text>
        </xsl:otherwise>
      </xsl:choose>

      <!--include index marks-->
      <xsl:if test="not($tocSource[@text:use-index-marks = 'false'])">
        <xsl:text>\u  </xsl:text>
      </xsl:if>

      <!--use hyperlinks -->
      <xsl:if test="text:a">
        <xsl:text> \h </xsl:text>
      </xsl:if>

      <!-- include elements with additional styles-->
      <xsl:if test="$tocSource/text:index-source-styles">
        <xsl:text> \t "</xsl:text>
        <xsl:for-each select="$tocSource/text:index-source-styles">
          <xsl:variable name="additionalStyleName"
            select="./text:index-source-style/@text:style-name"/>
          <xsl:value-of select="$additionalStyleName"/>
          <xsl:text>; </xsl:text>
          <xsl:value-of select="@text:outline-level"/>
          <xsl:text>"</xsl:text>
        </xsl:for-each>
      </xsl:if>
    </w:instrText>
  </xsl:template>


  <!-- Inserts an annotation reference if needed -->
  <xsl:template name="InsertAnnotationReference">
    <xsl:if test="ancestor::office:annotation and position() = 1">
      <w:r>
        <xsl:if test="ancestor::text:section/@text:display='none'">
          <w:rPr>
            <w:vanish/>
          </w:rPr>
        </xsl:if>
        <w:annotationRef/>
      </w:r>
    </xsl:if>
  </xsl:template>

  <!-- note marks -->
  <xsl:template match="text:note-citation" mode="note">
    <w:r>
      <w:rPr>
        <w:rStyle w:val="{concat(../@text:note-class, 'Reference')}"/>
        <xsl:if test="ancestor::text:section/@text:display='none'">
          <w:vanish/>
        </xsl:if>
      </w:rPr>
      <xsl:choose>
        <xsl:when test="../text:note-citation/@text:label">
          <w:t>
            <xsl:value-of select="../text:note-citation"/>
          </w:t>
        </xsl:when>
        <xsl:otherwise>
          <xsl:choose>
            <xsl:when test="../@text:note-class = 'footnote'">
              <w:footnoteRef/>
            </xsl:when>
            <xsl:when test="../@text:note-class = 'endnote' ">
              <w:endnoteRef/>
            </xsl:when>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </w:r>
    <!-- add an extra tab -->
    <w:r>
      <xsl:if test="ancestor::text:section/@text:display='none'">
        <w:rPr>
          <w:vanish/>
        </w:rPr>
      </xsl:if>
      <w:tab/>
    </w:r>
  </xsl:template>

  <!-- annotations -->
  <xsl:template match="office:annotation" mode="paragraph">
    <xsl:choose>

      <!--annotation embedded in text-box is not supported in Word-->
      <xsl:when test="ancestor::draw:text-box">
        <xsl:message terminate="no">feedback:Annotations in text-box</xsl:message>
      </xsl:when>
      <xsl:when
        test="ancestor::style:header or ancestor::style:header-left or ancestor::style:footer or ancestor::style:footer-left">
        <xsl:message terminate="no">feedback:Annotations in header or footer</xsl:message>
      </xsl:when>
      <!--default scenario-->
      <xsl:otherwise>
        <xsl:variable name="id">
          <xsl:call-template name="GenerateId">
            <xsl:with-param name="node" select="."/>
            <xsl:with-param name="nodetype" select="'annotation'"/>
          </xsl:call-template>
        </xsl:variable>
        <w:r>
          <xsl:call-template name="InsertRunProperties"/>
          <w:commentReference>
            <xsl:attribute name="w:id">
              <xsl:value-of select="$id"/>
            </xsl:attribute>
          </w:commentReference>
        </w:r>
      </xsl:otherwise>

    </xsl:choose>
  </xsl:template>

  <!-- links -->
  <xsl:template match="text:a" mode="paragraph">
    <xsl:choose>
      <!-- TOC hyperlink -->
      <xsl:when test="ancestor::text:index-body and position() = 1">
        <xsl:variable name="tocId" select="count(../preceding-sibling::text:p)+1"/>
        <w:hyperlink w:history="1">
          <xsl:attribute name="w:anchor">
            <xsl:value-of
              select="concat('_Toc',$tocId,generate-id(ancestor::text:table-of-content))"/>
          </xsl:attribute>
          <xsl:call-template name="InsertIndexItemContent">
            <xsl:with-param name="tocId" select="$tocId"/>
          </xsl:call-template>
        </w:hyperlink>
      </xsl:when>

      <!--text body link-->
      <xsl:otherwise>
        <w:hyperlink r:id="{generate-id()}">
          <xsl:apply-templates mode="paragraph"/>
        </w:hyperlink>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- computes text-box margin -->
  <xsl:template name="ComputeMarginX">
    <xsl:param name="parent"/>
    <xsl:choose>
      <xsl:when test="$parent">
        <xsl:variable name="recursive_result">
          <xsl:call-template name="point-measure">
            <xsl:with-param name="length">
              <xsl:call-template name="ComputeMarginX">
                <xsl:with-param name="parent" select="$parent[position()>1]"/>
              </xsl:call-template>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="svgx">
          <xsl:choose>
            <xsl:when test="$parent[1]/@svg:x">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:value-of select="$parent[1]/@svg:x"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>

        </xsl:variable>
        <xsl:value-of select="$svgx+$recursive_result"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="0"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- computes text-box margin -->
  <xsl:template name="ComputeMarginY">
    <xsl:param name="parent"/>
    <xsl:choose>
      <xsl:when test="$parent">
        <xsl:variable name="recursive_result">
          <xsl:call-template name="point-measure">
            <xsl:with-param name="length">
              <xsl:call-template name="ComputeMarginY">
                <xsl:with-param name="parent" select="$parent[position()>1]"/>
              </xsl:call-template>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="svgy">
          <xsl:choose>
            <xsl:when test="$parent[1]/@svg:y">
              <xsl:call-template name="point-measure">
                <xsl:with-param name="length">
                  <xsl:value-of select="$parent[1]/@svg:y"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:value-of select="$svgy+$recursive_result"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="0"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- text boxes -->
  <xsl:template match="draw:text-box" mode="paragraph">
    <w:r>
      <w:rPr>
        <xsl:if test="ancestor::text:section/@text:display='none'">
          <w:rPr>
            <w:vanish/>
          </w:rPr>
        </xsl:if>
        <xsl:call-template name="InsertTextBoxStyle"/>
      </w:rPr>
      <w:pict>

        <!--this properties are needed to make z-index work properly-->
        <v:shapetype coordsize="21600,21600" path="m,l,21600r21600,l21600,xe"
          xmlns:o="urn:schemas-microsoft-com:office:office">
          <v:stroke joinstyle="miter"/>
          <v:path gradientshapeok="t" o:connecttype="rect"/>
        </v:shapetype>

        <v:shape type="#_x0000_t202">
          <xsl:if test="ancestor::draw:a/@xlink:href">
            <xsl:attribute name="href">
              <xsl:choose>
                <xsl:when test="substring-before(ancestor::draw:a/@xlink:href,'/')='..'">
                  <xsl:value-of select="substring-after(ancestor::draw:a/@xlink:href,'/')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="ancestor::draw:a/@xlink:href"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
          <xsl:variable name="styleName" select=" parent::draw:frame/@draw:style-name"/>
          <xsl:variable name="automaticStyle" select="key('automatic-styles', $styleName)"/>
          <xsl:variable name="officeStyle"
            select="document('styles.xml')//office:document-styles/office:styles/style:style[@style:name = $styleName]"/>

          <xsl:variable name="shapeStyle" select="$automaticStyle | $officeStyle"/>

          <xsl:call-template name="InsertShapeProperties">
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
          </xsl:call-template>

          <!--insert text-box-->
          <xsl:call-template name="InsertTextBox">
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
          </xsl:call-template>

        </v:shape>
      </w:pict>
    </w:r>
  </xsl:template>

  <xsl:template name="InsertShapePositionProperties">
    <xsl:param name="shapeStyle"/>

    <xsl:variable name="frameWrap">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:wrap</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!-- absolute positioning for text-box -->
    <xsl:if test="not(parent::draw:frame/@text:anchor-type = 'as-char') ">
      <xsl:value-of select="'position:absolute;'"/>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertShapeSize">

    <!--  text-box width and height-->
    <xsl:variable name="frameW">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length"
          select="parent::draw:frame/@svg:width|parent::draw:frame/@fo:min-width"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="frameH">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length" select="@fo:min-height|parent::draw:frame/@svg:height"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="relWidth" select="substring-before(parent::draw:frame/@style:rel-width,'%')"/>
    <xsl:variable name="relHeight"
      select="substring-before(parent::draw:frame/@style:rel-height,'%')"/>

    <xsl:value-of select="concat('width:',$frameW,'pt;')"/>
    <xsl:value-of select="concat('height:',$frameH,'pt;')"/>

    <!--relative width and hight-->
    <xsl:if test="$relWidth">
      <xsl:value-of select="concat('mso-width-percent:',$relWidth,'0;')"/>
    </xsl:if>
    <xsl:if test="$relHeight">
      <xsl:value-of select="concat('mso-height-percent:',$relHeight,'0;')"/>
    </xsl:if>

    <xsl:variable name="relativeTo">
      <xsl:choose>
        <xsl:when test="parent::draw:frame/@text:anchor-type = 'page'">
          <xsl:text>page</xsl:text>
        </xsl:when>
        <xsl:when test="parent::draw:frame/@text:anchor-type = 'paragraph'">
          <xsl:text>margin</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:text>margin</xsl:text>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:value-of select="concat('mso-width-relative:',$relativeTo,';')"/>
    <xsl:value-of select="concat('mso-height-relative:',$relativeTo,';')"/>

  </xsl:template>

  <xsl:template name="InsertShapeZindex">
    <xsl:param name="shapeStyle"/>

    <xsl:variable name="frameWrap">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:wrap</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="runThrought">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:run-through</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>


    <!--z-index that we need to convert properly openoffice wrap-throught property -->
    <xsl:variable name="zIndex">
      <xsl:choose>
        <xsl:when test="$frameWrap='run-through' 
          and $runThrought='background'"
          >-251658240</xsl:when>
        <xsl:when test="$frameWrap='run-through' 
          and not($runThrought)">251658240</xsl:when>
        <xsl:otherwise>
          <xsl:choose>
            <xsl:when test="parent::draw:frame/@draw:z-index">
              <xsl:value-of select="2 + parent::draw:frame/@draw:z-index"/>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:value-of select="concat('z-index:', $zIndex, ';')"/>
  </xsl:template>

  <xsl:template name="InsertShapeCoordinates">

    <!-- text-box coordinates -->
    <xsl:variable name="posL">
      <xsl:if test="parent::draw:frame/@svg:x">
        <xsl:variable name="leftM">
          <xsl:call-template name="ComputeMarginX">
            <xsl:with-param name="parent" select="ancestor::draw:frame"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:value-of select="$leftM"/>
      </xsl:if>
    </xsl:variable>
    <xsl:variable name="posT">
      <xsl:if test="parent::draw:frame/@svg:y">
        <xsl:variable name="topM">
          <xsl:call-template name="ComputeMarginY">
            <xsl:with-param name="parent" select="ancestor::draw:frame"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:value-of select="$topM"/>
      </xsl:if>
    </xsl:variable>

    <xsl:if test="parent::draw:frame/@svg:x">
      <xsl:value-of select="concat('margin-left:',$posL,'pt;')"/>
    </xsl:if>
    <xsl:if test="parent::draw:frame/@svg:y">
      <xsl:value-of select="concat('margin-top:',$posT,'pt;')"/>
    </xsl:if>

  </xsl:template>

  <xsl:template name="InsertShapePositionRelative">
    <xsl:param name="shapeStyle"/>

    <xsl:variable name="horizontalRel">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:horizontal-rel</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="verticalRel">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:vertical-rel</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!-- text-box relative position -->
    <xsl:choose>
      <xsl:when test="$horizontalRel = 'page-end-margin' ">mso-position-horizontal-relative:
        right-margin-area;</xsl:when>
      <xsl:when test="$horizontalRel = 'page-start-margin' ">mso-position-horizontal-relative:
        left-margin-area;</xsl:when>
      <xsl:when test="$horizontalRel = 'page' ">mso-position-horizontal-relative: page;</xsl:when>
      <xsl:when test="$horizontalRel = 'page-content' ">mso-position-horizontal-relative: column;</xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="parent::draw:frame/@text:anchor-type = 'page' "
            >mso-position-horizontal-relative: page;</xsl:when>
          <xsl:when test="parent::draw:frame/@text:anchor-type = 'paragraph' "
            >mso-position-horizontal-relative: column;</xsl:when>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>

    <xsl:choose>
      <xsl:when test="$verticalRel = 'page' ">mso-position-vertical-relative: page;</xsl:when>
      <xsl:when test="$verticalRel = 'page-content' ">mso-position-vertical-relative:
      margin;</xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertShapePosition">
    <xsl:param name="shapeStyle"/>

    <xsl:variable name="horizontalPos">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:horizontal-pos</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="verticalPos">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:vertical-pos</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!--horizontal position-->
    <!-- The same style defined in styles.xsl  TODO manage horizontal-rel-->
    <xsl:if test="$horizontalPos">
      <xsl:choose>
        <xsl:when test="$horizontalPos = 'center'">
          <xsl:value-of select="concat('mso-position-horizontal:', 'center',';')"/>
        </xsl:when>
        <xsl:when test="$horizontalPos='left'">
          <xsl:value-of select="concat('mso-position-horizontal:', 'left',';')"/>
        </xsl:when>
        <xsl:when test="$horizontalPos='right'">
          <xsl:value-of select="concat('mso-position-horizontal:', 'right',';')"/>
        </xsl:when>
        <!-- <xsl:otherwise><xsl:value-of select="concat('mso-position-horizontal:', 'center',';')"/></xsl:otherwise> -->
      </xsl:choose>
    </xsl:if>

    <!-- vertical position-->
    <xsl:if test="$verticalPos">
      <xsl:choose>
        <xsl:when test="$verticalPos = 'middle'">
          <xsl:value-of select="concat('mso-position-vertical:', 'center',';')"/>
        </xsl:when>
        <xsl:when test="$verticalPos='top'">
          <xsl:value-of select="concat('mso-position-vertical:', 'top',';')"/>
        </xsl:when>
        <xsl:when test="$verticalPos='bottom'">
          <xsl:value-of select="concat('mso-position-vertical:', 'bottom',';')"/>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertShapeMargin">
    <xsl:param name="shapeStyle"/>

    <!--text-box spacing/margins -->
    <xsl:variable name="marginL">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length">
          <xsl:call-template name="GetGraphicProperties">
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
            <xsl:with-param name="attribName">fo:margin-left</xsl:with-param>
          </xsl:call-template>
        </xsl:with-param>
        <xsl:with-param name="round">
        <xsl:text>false</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="marginT">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length">
          <xsl:call-template name="GetGraphicProperties">
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
            <xsl:with-param name="attribName">fo:margin-top</xsl:with-param>
          </xsl:call-template>
          </xsl:with-param>
        <xsl:with-param name="round">
          <xsl:text>false</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="marginR">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length">
          <xsl:call-template name="GetGraphicProperties">
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
            <xsl:with-param name="attribName">fo:margin-right</xsl:with-param>
          </xsl:call-template>
        </xsl:with-param>
        <xsl:with-param name="round">
          <xsl:text>false</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="marginB">
      <xsl:call-template name="point-measure">
        <xsl:with-param name="length">
          <xsl:call-template name="GetGraphicProperties">
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
            <xsl:with-param name="attribName">fo:margin-bottom</xsl:with-param>
          </xsl:call-template>
        </xsl:with-param>
        <xsl:with-param name="round">
          <xsl:text>false</xsl:text>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:if test="$marginL">
      <xsl:value-of select="concat('mso-wrap-distance-left:', $marginL,'pt;')"/>
    </xsl:if>
    <xsl:if test="$marginT">
      <xsl:value-of select="concat('mso-wrap-distance-top:', $marginT,'pt;')"/>
    </xsl:if>
    <xsl:if test="$marginR">
      <xsl:value-of select="concat('mso-wrap-distance-right:', $marginR,'pt;')"/>
    </xsl:if>
    <xsl:if test="$marginB">
      <xsl:value-of select="concat('mso-wrap-distance-bottom:', $marginB,'pt;')"/>
    </xsl:if>


  </xsl:template>

  <xsl:template name="InsertShapeFill">
    <xsl:param name="shapeStyle"/>

    <xsl:variable name="shapefillColor">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">fo:background-color</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!-- frame color is dependent on page background color in some cases -->
    <xsl:variable name="fillColor">
      <xsl:choose>
        <xsl:when test="$shapefillColor = 'transparent' or $shapefillColor = '' ">
          <!-- when no fill is set for frame it should take background color of page (as it is in ODF) -->
          <xsl:for-each select="document('styles.xml')">
            <xsl:variable name="defaultBgColor"
              select="key('page-layouts', $default-master-style/@style:page-layout-name)[1]/style:page-layout-properties/@fo:background-color"/>
            <xsl:value-of select="$defaultBgColor"/>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$shapefillColor"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <!--fill color-->
    <xsl:if test="$fillColor != '' ">
      <xsl:attribute name="fillcolor">
        <xsl:value-of select="$fillColor"/>
      </xsl:attribute>
    </xsl:if>

  </xsl:template>

  <xsl:template name="InsertShapeBorders">
    <xsl:param name="shapeStyle"/>

    <xsl:variable name="border">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">fo:border</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!--borders-->
    <xsl:choose>
      <!-- no border in current style -->
      <xsl:when test="$border='' or $border = 'none' ">
        <xsl:attribute name="stroked">f</xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="strokeColor" select="substring-after($border,'#')"/>
        <xsl:variable name="strokeWeight">
          <xsl:call-template name="point-measure">
            <xsl:with-param name="length" select="substring-before($border,' ')"/>
            <xsl:with-param name="round">
              <xsl:text>false</xsl:text>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="styleBorderLine">
          <xsl:call-template name="GetGraphicProperties">
            <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
            <xsl:with-param name="attribName">style:border-line-width</xsl:with-param>
          </xsl:call-template>
        </xsl:variable>

        <xsl:if test="$strokeColor != '' ">
          <xsl:attribute name="strokecolor">
            <xsl:value-of select="concat('#', $strokeColor)"/>
          </xsl:attribute>
        </xsl:if>

        <xsl:if test="$strokeWeight != '' ">
          <xsl:attribute name="strokeweight">
            <xsl:value-of select="concat($strokeWeight,'pt')"/>
          </xsl:attribute>
        </xsl:if>

        <!--  line styles -->
        <xsl:if test="substring-before(substring-after($border,' ' ),' ' ) != 'solid' ">
          <v:stroke>
            <xsl:attribute name="linestyle">
              <xsl:choose>
                <xsl:when test="$styleBorderLine">

                  <xsl:variable name="innerLineWidth">
                    <xsl:call-template name="point-measure">
                      <xsl:with-param name="length" select="substring-before($styleBorderLine,' ' )"
                      />
                    </xsl:call-template>
                  </xsl:variable>

                  <xsl:variable name="outerLineWidth">
                    <xsl:call-template name="point-measure">
                      <xsl:with-param name="length"
                        select="substring-after(substring-after($styleBorderLine,' ' ),' ' )"/>
                    </xsl:call-template>
                  </xsl:variable>

                  <xsl:if test="$innerLineWidth = $outerLineWidth">thinThin</xsl:if>
                  <xsl:if test="$innerLineWidth > $outerLineWidth">thinThick</xsl:if>
                  <xsl:if test="$outerLineWidth > $innerLineWidth  ">thickThin</xsl:if>

                </xsl:when>
              </xsl:choose>
            </xsl:attribute>
          </v:stroke>
        </xsl:if>
      </xsl:otherwise>

      <!--default scenario-->
      <!--     <xsl:otherwise>
        <xsl:attribute name="stroked">f</xsl:attribute>
        </xsl:otherwise>
       -->
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertShapeTransparency">
    <xsl:param name="shapeStyle"/>

    <xsl:variable name="backgroundColor">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">fo:background-color</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="backgroundTransparency">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:background-transparency</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!--fill  transparency-->
    <xsl:variable name="opacity" select="100 - substring-before($backgroundTransparency,'%')"/>

    <xsl:if test="$backgroundTransparency != '' and $backgroundColor != 'transparent' ">
      <v:fill>
        <xsl:attribute name="opacity">
          <xsl:value-of select="concat($opacity,'%')"/>
        </xsl:attribute>
      </v:fill>
    </xsl:if>
  </xsl:template>

  <!--converts oo frame style properties to shape properties for text-box-->
  <xsl:template name="InsertShapeProperties">
    <xsl:param name="shapeStyle"/>

    <xsl:attribute name="style">

      <xsl:call-template name="InsertShapePositionProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

      <xsl:call-template name="InsertShapeSize"/>

      <xsl:call-template name="InsertShapeZindex">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

      <xsl:call-template name="InsertShapeCoordinates"/>

      <xsl:call-template name="InsertShapePositionRelative">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

      <xsl:call-template name="InsertShapePosition">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

      <xsl:call-template name="InsertShapeMargin">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

    </xsl:attribute>

    <xsl:call-template name="InsertShapeFill">
      <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
    </xsl:call-template>

    <xsl:call-template name="InsertShapeBorders">
      <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
    </xsl:call-template>

    <xsl:call-template name="InsertShapeTransparency">
      <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
    </xsl:call-template>

  </xsl:template>

  <!-- finds shape graphic property recursively in style or parent style -->
  <xsl:template name="GetGraphicProperties">
    <xsl:param name="attribName"/>
    <xsl:param name="shapeStyle"/>

    <xsl:choose>
      <xsl:when test="$shapeStyle/style:graphic-properties/attribute::node()[name() = $attribName] ">
        <xsl:value-of
          select="$shapeStyle/style:graphic-properties/attribute::node()[name() = $attribName]"/>
      </xsl:when>

      <xsl:when test="$shapeStyle/@style:parent-style-name">
        <xsl:variable name="parentStyleName">
          <xsl:value-of select="$shapeStyle/@style:parent-style-name"/>
        </xsl:variable>

        <xsl:variable name="parentStyle"
          select="document('styles.xml')//office:document-styles/office:styles/style:style[@style:name = $parentStyleName]"/>

        <xsl:call-template name="GetGraphicProperties">
          <xsl:with-param name="shapeStyle" select="$parentStyle"/>
          <xsl:with-param name="attribName" select="$attribName"/>
        </xsl:call-template>
      </xsl:when>
      <!--    <xsl:otherwise></xsl:otherwise>-->
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertShapeWrap">
    <xsl:param name="shapeStyle"/>

    <xsl:variable name="frameWrap">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">style:wrap</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!--frame wrap-->
    <xsl:choose>
      <xsl:when test="parent::draw:frame/@text:anchor-type = 'as-char' ">
        <w10:wrap type="none"/>
        <w10:anchorlock/>
      </xsl:when>
      <xsl:when test="$frameWrap = 'left' ">
        <w10:wrap type="square" side="left"/>
      </xsl:when>
      <xsl:when test="$frameWrap = 'right' ">
        <w10:wrap type="square" side="right"/>
      </xsl:when>
      <xsl:when test="$frameWrap = 'none'">
        <w10:wrap type="topAndBottom"/>
      </xsl:when>
      <xsl:when test="$frameWrap = 'parallel' ">
        <w10:wrap type="square"/>
      </xsl:when>
      <xsl:when test="$frameWrap = 'dynamic' ">
        <w10:wrap type="square" side="largest"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!--inserts text-box into shape element -->
  <xsl:template name="InsertTextBox">
    <xsl:param name="shapeStyle"/>

    <v:textbox>
      <xsl:attribute name="style">
        <xsl:if test="@fo:min-height">
          <xsl:value-of select="'mso-fit-shape-to-text:t'"/>
        </xsl:if>
      </xsl:attribute>

      <xsl:call-template name="InsertTextBoxInset">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

      <w:txbxContent>
        <xsl:for-each select="child::node()">

          <xsl:choose>
            <!--   ignore embedded text-box becouse word doesn't support it-->
            <xsl:when test="self::node()[name(draw:text-box)]">
              <xsl:message terminate="no">feedback: Nested frames</xsl:message>
            </xsl:when>

            <!--default scenario-->
            <xsl:otherwise>
              <xsl:apply-templates select="."/>
            </xsl:otherwise>
          </xsl:choose>

        </xsl:for-each>
      </w:txbxContent>

      <xsl:call-template name="InsertShapeWrap">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
      </xsl:call-template>

    </v:textbox>
  </xsl:template>

  <!--converts oo frame padding into inset for text-box -->
  <xsl:template name="InsertTextBoxInset">
    <xsl:param name="shapeStyle"/>

    <xsl:variable name="padding">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">fo:padding</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="paddingTop">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">fo:padding-top</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="paddingRight">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">fo:padding-right</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="paddingLeft">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">fo:padding-left</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="paddingBottom">
      <xsl:call-template name="GetGraphicProperties">
        <xsl:with-param name="shapeStyle" select="$shapeStyle"/>
        <xsl:with-param name="attribName">fo:padding-bottom</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:attribute name="inset">
      <xsl:choose>
        <xsl:when test="$padding or $paddingTop">
          <xsl:call-template name="CalculateTextBoxPadding">
            <xsl:with-param name="padding" select="$padding"/>
            <xsl:with-param name="paddingTop" select="$paddingTop"/>
            <xsl:with-param name="paddingLeft" select="$paddingLeft"/>
            <xsl:with-param name="paddingRight" select="$paddingRight"/>
            <xsl:with-param name="paddingBottom" select="$paddingBottom"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="$padding or $paddingTop">
          <xsl:call-template name="CalculateTextBoxPadding">
            <xsl:with-param name="padding" select="$padding"/>
            <xsl:with-param name="paddingTop" select="$paddingTop"/>
            <xsl:with-param name="paddingLeft" select="$paddingLeft"/>
            <xsl:with-param name="paddingRight" select="$paddingRight"/>
            <xsl:with-param name="paddingBottom" select="$paddingBottom"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="CalculateTextBoxPadding"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>


  <!--calculates textbox inset attribute  -->
  <xsl:template name="CalculateTextBoxPadding">
    <xsl:param name="padding"/>
    <xsl:param name="paddingBottom"/>
    <xsl:param name="paddingLeft"/>
    <xsl:param name="paddingRight"/>
    <xsl:param name="paddingTop"/>


    <xsl:choose>
      <xsl:when test="not($padding)">0mm,0mm,0mm,0mm</xsl:when>
      <xsl:when test="$padding">
        <xsl:variable name="paddingVal">
          <xsl:call-template name="milimeter-measure">
            <xsl:with-param name="length" select="$padding"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:value-of
          select="concat($paddingVal,'mm,',$paddingVal,'mm,',$paddingVal,'mm,',$paddingVal,'mm')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="padding-top">
          <xsl:if test="$paddingTop">
            <xsl:call-template name="milimeter-measure">
              <xsl:with-param name="length" select="$paddingTop"/>
            </xsl:call-template>
          </xsl:if>
        </xsl:variable>
        <xsl:variable name="padding-right">
          <xsl:if test="$paddingRight">
            <xsl:call-template name="milimeter-measure">
              <xsl:with-param name="length" select="$paddingRight"/>
            </xsl:call-template>
          </xsl:if>
        </xsl:variable>
        <xsl:variable name="padding-bottom">
          <xsl:if test="$paddingBottom">
            <xsl:call-template name="milimeter-measure">
              <xsl:with-param name="length" select="$paddingBottom"/>
            </xsl:call-template>
          </xsl:if>
        </xsl:variable>
        <xsl:variable name="padding-left">
          <xsl:if test="$paddingLeft">
            <xsl:call-template name="milimeter-measure">
              <xsl:with-param name="length" select="$paddingLeft"/>
            </xsl:call-template>
          </xsl:if>
        </xsl:variable>
        <xsl:if test="$paddingTop">
          <xsl:value-of
            select="concat($padding-left,'mm,',$padding-top,'mm,',$padding-right,'mm,',$padding-bottom,'mm')"
          />
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Inserts the style of a text box -->
  <xsl:template name="InsertTextBoxStyle">
    <xsl:variable name="prefixedStyleName">
      <xsl:call-template name="GetPrefixedStyleName">
        <xsl:with-param name="styleName" select="parent::draw:frame/@draw:style-name"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:if test="$prefixedStyleName!=''">
      <w:rStyle w:val="{$prefixedStyleName}"/>
    </xsl:if>
  </xsl:template>

  <xsl:template match="draw:frame" mode="paragraph">
    <xsl:call-template name="InsertEmbeddedTextboxes"/>
  </xsl:template>

  <xsl:template match="draw:frame">
    <w:p>
      <xsl:call-template name="InsertEmbeddedTextboxes"/>
    </w:p>
  </xsl:template>

  <!-- inserts textboxes which are embedded in odf as one after another in word -->
  <xsl:template name="InsertEmbeddedTextboxes">
    <xsl:for-each select="descendant::draw:text-box">
      <xsl:apply-templates mode="paragraph" select="."/>
    </xsl:for-each>
  </xsl:template>

  <!-- lists -->
  <xsl:template match="text:list">
    <xsl:param name="level" select="-1"/>
    <xsl:apply-templates>
      <xsl:with-param name="level" select="$level+1"/>
    </xsl:apply-templates>
  </xsl:template>

  <!-- list items -->
  <xsl:template match="text:list-item|text:list-header">
    <xsl:param name="level"/>
    <xsl:choose>
      <xsl:when test="*[1][self::text:p or self::text:h]">
        <w:p>
          <w:pPr>

            <!-- insert style -->
            <xsl:call-template name="InsertParagraphStyle">
              <xsl:with-param name="styleName" select="*[1]/@text:style-name"/>
            </xsl:call-template>

            <!-- insert number -->
            <xsl:call-template name="InsertNumbering">
              <xsl:with-param name="level" select="$level"/>
            </xsl:call-template>

            <!-- override abstract num indent and tab if paragraph has margin defined -->
            <xsl:call-template name="InsertIndent">
              <xsl:with-param name="level" select="$level"/>
            </xsl:call-template>

            <!-- insert heading outline level -->
            <xsl:call-template name="InsertOutlineLevel">
              <xsl:with-param name="node" select="*[1]"/>
            </xsl:call-template>

            <!-- insert page break before table when required -->
            <xsl:call-template name="InsertPageBreakBefore"/>
          </w:pPr>

          <!--TOC-->
          <xsl:variable name="isBookmarked">
            <xsl:call-template name="IsTOCBookmark">
              <xsl:with-param name="styleName" select="child::*[1]/@text:style-name"/>
            </xsl:call-template>
          </xsl:variable>

          <xsl:if test="$isBookmarked = 'true' ">
            <xsl:variable name="bookmarkType">start</xsl:variable>
            <xsl:call-template name="InsertTOCBookmark">
              <xsl:with-param name="bookmarkType" select="$bookmarkType"/>
            </xsl:call-template>
          </xsl:if>

          <!-- if we are in an annotation, we may have to insert annotation reference -->
          <xsl:call-template name="InsertAnnotationReference"/>

          <!-- footnote or endnote - Include the mark to the first paragraph only when first child of 
          text:note-body is not paragraph -->
          <xsl:if
            test="ancestor::text:note and not(ancestor::text:note-body/child::*[1][self::text:p | self::text:h]) and position() = 1">
            <xsl:apply-templates select="ancestor::text:note/text:note-citation" mode="note"/>
          </xsl:if>

          <!-- first paragraph -->
          <xsl:apply-templates select="*[1]" mode="paragraph"/>

          <!--TOC end-->
          <xsl:if test="$isBookmarked = 'true' ">
            <xsl:variable name="bookmarkType">end</xsl:variable>
            <xsl:call-template name="InsertTOCBookmark">
              <xsl:with-param name="bookmarkType" select="$bookmarkType"/>
            </xsl:call-template>
          </xsl:if>

        </w:p>

        <!-- others (text:p or text:list) -->
        <xsl:apply-templates select="*[position() != 1]">
          <xsl:with-param name="level" select="$level"/>
        </xsl:apply-templates>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates>
          <xsl:with-param name="level" select="$level"/>
        </xsl:apply-templates>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Inserts the number of a list item -->
  <xsl:template name="InsertNumbering">
    <xsl:param name="level"/>
    <xsl:variable name="stylename" select="./@text:style-name"/>
    <xsl:variable name="list"
      select="//office:document-content/office:automatic-styles/style:style[@style:name = $stylename]"/>
    <xsl:choose>
      <xsl:when
        test="self::text:list-item or self::text:list-header or $list/@style:list-style-name">
        <xsl:if test="not(self::text:list-header)">
          <w:numPr>
            <xsl:choose>
              <xsl:when test="self::text:list-item">
                <w:ilvl w:val="{$level}"/>
              </xsl:when>
              <xsl:otherwise>
                <w:ilvl w:val="{$level/@text:outline-level - 1}"/>
              </xsl:otherwise>
            </xsl:choose>
            <w:numId>
              <xsl:attribute name="w:val">
                <xsl:choose>
                  <xsl:when test="self::text:list-item">
                    <xsl:call-template name="GetNumberingId">
                      <xsl:with-param name="styleName" select="ancestor::text:list/@text:style-name"
                      />
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:otherwise>

                    <xsl:call-template name="GetNumberingId">
                      <xsl:with-param name="styleName" select="$list/@style:list-style-name"/>
                    </xsl:call-template>

                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </w:numId>
          </w:numPr>
        </xsl:if>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if
          test="$level[self::text:h] and $level/@text:outline-level &lt;= 9 and document('styles.xml')//office:document-styles/office:styles/text:outline-style/text:outline-level-style/@style:num-format !=''">
          <w:numPr>
            <w:ilvl w:val="{$level/@text:outline-level - 1}"/>
            <w:numId w:val="1"/>
          </w:numPr>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--inserts index item content for all types of index-->
  <xsl:template name="InsertIndexItemContent">

    <!-- references to index bookmark id in text -->
    <xsl:param name="tocId" select="count(preceding-sibling::text:p)+1"/>

    <!-- insert text from the beginning of index item-->
    <w:r>
      <w:rPr>
        <xsl:if test="ancestor::text:section/@text:display='none'">
          <w:vanish/>
        </xsl:if>
        <w:noProof/>
      </w:rPr>
      <w:t>
        <xsl:choose>
          <xsl:when test="child::node()[position() &lt; last()]">
            <xsl:for-each select="child::node()[position() &lt; last()]">
              <xsl:value-of select="."/>
            </xsl:for-each>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="."/>
          </xsl:otherwise>
        </xsl:choose>
      </w:t>
    </w:r>

    <!-- insert tab in between if there is any -->
    <xsl:apply-templates select="text:tab|self::text:a/text:tab|text:span|parent::text:p/text:tab"
      mode="paragraph"/>

    <!-- alphabetical index doesn't support page reference link -->
    <xsl:if test="not(ancestor::text:alphabetical-index)">
      <xsl:call-template name="InsertIndexPageRefStart">
        <xsl:with-param name="tocId" select="$tocId"/>
      </xsl:call-template>
    </xsl:if>

    <!-- insert text at the end of index item -->
    <w:r>
      <w:rPr>
        <xsl:if test="ancestor::text:section/@text:display='none'">
          <w:vanish/>
        </xsl:if>
        <w:noProof/>
        <w:webHidden/>
      </w:rPr>
      <xsl:apply-templates select="descendant::text()[last()]" mode="text"/>
    </w:r>

    <!-- alphabetical index doesn't support page reference link -->
    <xsl:if test="not(ancestor::text:alphabetical-index)">
      <xsl:call-template name="InsertIndexPageRefEnd"/>
    </xsl:if>

  </xsl:template>

  <xsl:template name="InsertIndexPageRefEnd">
    <w:r>
      <w:rPr>
        <xsl:if test="ancestor::text:section/@text:display='none'">
          <w:vanish/>
        </xsl:if>
        <w:noProof/>
        <w:webHidden/>
      </w:rPr>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
  </xsl:template>

  <xsl:template name="InsertIndexPageRefStart">
    <xsl:param name="tocId"/>

    <w:r>
      <w:rPr>
        <xsl:if test="ancestor::text:section/@text:display='none'">
          <w:vanish/>
        </xsl:if>
        <w:noProof/>
        <w:webHidden/>
      </w:rPr>
      <w:fldChar w:fldCharType="begin">
        <w:fldData xml:space="preserve">CNDJ6nn5us4RjIIAqgBLqQsCAAAACAAAAA4AAABfAFQAbwBjADEANAAxADgAMwA5ADIANwA2AAAA</w:fldData>
      </w:fldChar>
    </w:r>
    <w:r>
      <w:rPr>
        <xsl:if test="ancestor::text:section/@text:display='none'">
          <w:vanish/>
        </xsl:if>
        <w:noProof/>
        <w:webHidden/>
      </w:rPr>
      <w:instrText xml:space="preserve"><xsl:value-of select="concat('PAGEREF _Toc', $tocId,generate-id(ancestor::text:table-of-content), ' \h')"/></w:instrText>
    </w:r>
    <w:r>
      <w:rPr>
        <xsl:if test="ancestor::text:section/@text:display='none'">
          <w:vanish/>
        </xsl:if>
        <w:noProof/>
        <w:webHidden/>
      </w:rPr>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
  </xsl:template>
  
  <!-- dealing with text next to shapes -->
  
  <xsl:template match="text()|text:s" mode="shapes">
    <xsl:apply-templates select="." mode="paragraph"/>
  </xsl:template>

  <!-- text and spaces -->
  
  <xsl:template match="text()|text:s" mode="paragraph">
    <w:r>
      <xsl:call-template name="InsertRunProperties"/>
      <xsl:apply-templates select="." mode="text"/>
    </w:r>
  </xsl:template>

  <!-- Inserts the Run properties -->
  <xsl:template name="InsertRunProperties">
    <!-- apply text properties if needed -->
    <xsl:choose>
      <xsl:when test="ancestor::text:span|ancestor::text:a|self::text:list-level-style-number">
        <w:rPr>
          <xsl:call-template name="InsertRunStyle">
            <xsl:with-param name="styleName">
              <xsl:call-template name="GetStyleName"/>
            </xsl:with-param>
          </xsl:call-template>
          <xsl:if test="ancestor::text:section/@text:display='none'">
            <w:vanish/>
          </xsl:if>
        </w:rPr>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="ancestor::text:section/@text:display='none'">
          <w:rPr>
            <w:vanish/>
          </w:rPr>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Inserts the style of a run -->
  <xsl:template name="InsertRunStyle">
    <xsl:param name="styleName"/>
    <xsl:variable name="prefixedStyleName">
      <xsl:call-template name="GetPrefixedStyleName">
        <xsl:with-param name="styleName" select="$styleName"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$prefixedStyleName!=''">
        <w:rStyle w:val="{$prefixedStyleName}"/>
      </xsl:when>
      <xsl:when test="ancestor::text:a">
        <w:rStyle w:val="Hyperlink"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- spaces (within a text flow) -->
  <xsl:template match="text:s" mode="text">
    <w:t xml:space="preserve"><xsl:call-template name="extra-spaces"><xsl:with-param name="spaces" select="@text:c"/></xsl:call-template></w:t>
  </xsl:template>

  <!-- simple text (within a text flow) -->
  <xsl:template match="text()" mode="text">
    <xsl:choose>
      <xsl:when test="ancestor::text:index-body">
        <xsl:choose>
          <xsl:when test="preceding-sibling::text:tab">
            <w:t xml:space="preserve"><xsl:value-of select="."/></w:t>
          </xsl:when>
          <xsl:when test="not(following-sibling::text:tab)">
            <xsl:choose>
              <xsl:when test="not(ancestor::text:a)">
                <w:t xml:space="preserve"><xsl:value-of select="."/></w:t>
              </xsl:when>
              <xsl:otherwise>
                <w:t>
                  <xsl:value-of select="ancestor::text:p/text()"/>
                </w:t>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise> </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <w:t xml:space="preserve"><xsl:value-of select="."/></w:t>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- tab stops -->
  <xsl:template match="text:tab-stop" mode="paragraph">
    <w:r>
      <xsl:if test="ancestor::text:section/@text:display='none'">
        <w:vanish/>
      </xsl:if>
      <w:tab/>
      <w:t/>
    </w:r>
  </xsl:template>

  <!-- tabs -->
  <xsl:template match="text:tab" mode="paragraph">
    <w:r>
      <w:rPr>
        <xsl:if test="ancestor::text:section/@text:display='none'">
          <w:vanish/>
        </xsl:if>
        <w:noProof/>
        <w:webHidden/>
      </w:rPr>
      <w:tab/>
    </w:r>
  </xsl:template>

  <!-- line breaks -->
  <xsl:template match="text:line-break" mode="paragraph">
    <w:r>
      <xsl:if test="ancestor::text:section/@text:display='none'">
        <w:rPr>
          <w:vanish/>
        </w:rPr>
      </xsl:if>
      <w:br/>
      <w:t/>
    </w:r>
  </xsl:template>

  <!-- line breaks (within the text flow) -->
  <xsl:template match="text:line-break" mode="text">
    <w:br/>
  </xsl:template>

  <!-- notes (footnotes or endnotes) -->
  <xsl:template match="text:note" mode="paragraph">
    <w:r>
      <w:rPr>
        <w:rStyle w:val="{concat(@text:note-class, 'Reference')}"/>
        <!-- COMMENT : why this color as a direct formatting property here ? -->
        <xsl:variable name="fo:color"
          select="substring-after(key('automatic-styles', parent::text:span/@text:style-name)/style:text-properties/@fo:color,'#')"/>
        <xsl:if test="$fo:color">
          <w:color w:val="{$fo:color}"/>
        </xsl:if>
        <xsl:if test="ancestor::text:section/@text:display='none'">
          <w:vanish/>
        </xsl:if>
      </w:rPr>
      <xsl:apply-templates select="." mode="text"/>
    </w:r>
  </xsl:template>

  <!-- empty alphabetical indexes creating mark entry -->
  <xsl:template match="text:alphabetical-index-mark" mode="paragraph">
    <w:r>
      <xsl:if test="ancestor::text:section/@text:display='none'">
        <w:rPr>
          <w:vanish/>
        </w:rPr>
      </xsl:if>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
      <xsl:if test="ancestor::text:section/@text:display='none'">
        <w:rPr>
          <w:vanish/>
        </w:rPr>
      </xsl:if>
      <w:instrText xml:space="preserve"> XE "</w:instrText>
    </w:r>
    <w:r>
      <xsl:if test="ancestor::text:section/@text:display='none'">
        <w:rPr>
          <w:vanish/>
        </w:rPr>
      </xsl:if>
      <w:instrText>
        <xsl:value-of select="./@text:string-value"/>
      </w:instrText>
    </w:r>
    <w:r>
      <xsl:if test="ancestor::text:section/@text:display='none'">
        <w:rPr>
          <w:vanish/>
        </w:rPr>
      </xsl:if>
      <w:instrText xml:space="preserve">" </w:instrText>
    </w:r>
    <w:r>
      <xsl:if test="ancestor::text:section/@text:display='none'">
        <w:rPr>
          <w:vanish/>
        </w:rPr>
      </xsl:if>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
   </xsl:template>

  <!-- alphabetical indexes creating mark entry -->
  <xsl:template match="text:alphabetical-index-mark-end" mode="paragraph">
    <w:r>
      <xsl:if test="ancestor::text:section/@text:display='none'">
        <w:rPr>
          <w:vanish/>
        </w:rPr>
      </xsl:if>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
      <xsl:if test="ancestor::text:section/@text:display='none'">
        <w:rPr>
          <w:vanish/>
        </w:rPr>
      </xsl:if>
      <w:instrText xml:space="preserve"> XE "</w:instrText>
    </w:r>
    <w:r>
      <xsl:if test="ancestor::text:section/@text:display='none'">
        <w:rPr>
          <w:vanish/>
        </w:rPr>
      </xsl:if>
      <w:instrText>
        <xsl:variable name="id" select="@text:id"/>
        <xsl:for-each select="preceding-sibling::node()">
          <xsl:if
            test="preceding-sibling::node()[name() = 'text:alphabetical-index-mark-start' and @text:id = $id]">
            <xsl:choose>
              <xsl:when test="self::text()">
                <xsl:value-of select="."/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:apply-templates select="."/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </w:instrText>
    </w:r>
    <w:r>
      <xsl:if test="ancestor::text:section/@text:display='none'">
        <w:rPr>
          <w:vanish/>
        </w:rPr>
      </xsl:if>
      <w:instrText xml:space="preserve">" </w:instrText>
    </w:r>
    <w:r>
      <xsl:if test="ancestor::text:section/@text:display='none'">
        <w:rPr>
          <w:vanish/>
        </w:rPr>
      </xsl:if>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
    <!-- <xsl:apply-templates select="text:s" mode="text"></xsl:apply-templates> -->
  </xsl:template>

  <!-- spaces -->
  <xsl:template match="text:s">
    <xsl:call-template name="extra-spaces">
      <xsl:with-param name="spaces" select="@text:c"/>
    </xsl:call-template>
  </xsl:template>

  <!-- sequences -->
  <xsl:template match="text:sequence" mode="paragraph">
    <xsl:variable name="id">
      <xsl:value-of select="number(count(preceding::text:sequence)+count( key('headers','')))+1"/>
      <!--<xsl:value-of select="count( key('headers',''))"/>-->
    </xsl:variable>
    <w:fldSimple>
      <xsl:variable name="numType">
        <xsl:choose>
          <xsl:when test="@style:num-format = 'i'">\* roman</xsl:when>
          <xsl:when test="@style:num-format = 'I'">\* Roman</xsl:when>
          <xsl:when test="@style:num-format = 'a'">\* alphabetic</xsl:when>
          <xsl:when test="@style:num-format = 'A'">\* ALPHABETIC</xsl:when>
          <xsl:otherwise>\* arabic</xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:attribute name="w:instr">
        <xsl:value-of select="concat('SEQ ', @text:name,' ', $numType)"/>
      </xsl:attribute>
      <w:bookmarkStart w:id="{$id}" w:name="{concat('_Toc',$id)}"/>
      <w:r>
        <xsl:if test="ancestor::text:section/@text:display='none'">
          <w:rPr>
            <w:vanish/>
          </w:rPr>
        </xsl:if>
        <w:t>
          <xsl:value-of select="."/>
        </w:t>
      </w:r>
      <w:bookmarkEnd w:id="{$id}"/>
    </w:fldSimple>
  </xsl:template>

  <!-- sections -->
  <xsl:template match="text:section">
    <xsl:message terminate="no">progress:text:section</xsl:message>
    <xsl:apply-templates/>
  </xsl:template>

  <!-- Extra spaces management, courtesy of J. David Eisenberg -->
  <xsl:variable name="spaces" xml:space="preserve">                                       </xsl:variable>

  <xsl:template name="extra-spaces">
    <xsl:param name="spaces"/>
    <xsl:choose>
      <xsl:when test="$spaces">
        <xsl:call-template name="insert-spaces">
          <xsl:with-param name="n" select="$spaces"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:text> </xsl:text>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="insert-spaces">
    <xsl:param name="n"/>
    <xsl:choose>
      <xsl:when test="$n &lt;= string-length($spaces)">
        <xsl:value-of select="substring($spaces, 1, $n)" xml:space="preserve"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$spaces"/>
        <xsl:call-template name="insert-spaces">
          <xsl:with-param name="n">
            <xsl:value-of select="$n - string-length($spaces)"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- ignored -->
  <xsl:template match="text()"/>
  <xsl:template match="text:tracked-changes"/>

</xsl:stylesheet>

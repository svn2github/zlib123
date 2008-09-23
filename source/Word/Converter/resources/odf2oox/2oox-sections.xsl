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
  xmlns:psect="urn:cleverage:xmlns:post-processings:sections"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  exclude-result-prefixes="office text table fo style v draw">


  <xsl:strip-space elements="*"/>


  <xsl:key name="master-pages" match="style:master-page" use="@style:name"/>
  <xsl:key name="page-layouts" match="style:page-layout" use="@style:name"/>
  <xsl:key name="master-based-styles" match="style:style[@style:master-page-name]" use="@style:name"/>
  <xsl:key name="sections" match="style:style[@style:family='section']" use="@style:name"/>
  <xsl:key name="automatic-styles" match="office:automatic-styles/style:style" use="@style:name"/>


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

    <xsl:template name="SoftPageBreaks">
      
    <!--math & clam, dialogika: bugfix #1838832 BEGIN--> 
    <!--w:pPr tag is now written in the OoxSectionsPostProcessor if w:sectPr is not skipped-->
    <!--<w:pPr>-->
      <w:sectPr>
        <xsl:attribute name="psect:next-soft-page-break">true</xsl:attribute>
      </w:sectPr>
    <!--</w:pPr>-->
    <!--math & clam, dialogika: bugfix #1838832 END-->
  </xsl:template>

  <!-- Document final section properties -->
  <xsl:template name="InsertDocumentFinalSectionProperties">
    <w:sectPr/>
  </xsl:template>

  <!-- Mark the text element if its style is tied to a master-page -->
  <xsl:template name="MarkMasterPage">
    <xsl:choose>
      <xsl:when test="self::text:p or self::text:h or self::table:table">
        <xsl:variable name="master-page-name">
          <xsl:call-template name="GetMasterPageNameFromHierarchy">
            <xsl:with-param name="style-name" select="@text:style-name|@table:style-name"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:if test="$master-page-name != '' ">
          <xsl:attribute name="psect:master-page-name">
            <xsl:value-of select="$master-page-name"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:variable name="pageNumber">
          <xsl:call-template name="GetPageStartNumber">
            <xsl:with-param name="style-name" select="@text:style-name|@table:style-name"/>
            <xsl:with-param name="context">
              <xsl:choose>
                <xsl:when test="ancestor::office:document-content">content.xml</xsl:when>
                <xsl:otherwise>styles.xml</xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:if test="number($pageNumber)">
          <xsl:attribute name="psect:page-number">
            <xsl:value-of select="$pageNumber"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <xsl:when test="self::text:list-item or self::text:list-header">
        <xsl:variable name="master-page-name">
          <xsl:call-template name="GetMasterPageNameFromHierarchy">
            <xsl:with-param name="style-name"
              select="*[1][self::text:p or self::text:h]/@text:style-name"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:if test="$master-page-name != '' ">
          <xsl:attribute name="psect:master-page-name">
            <xsl:value-of select="$master-page-name"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:variable name="pageNumber">
          <xsl:call-template name="GetPageStartNumber">
            <xsl:with-param name="style-name"
              select="*[1][self::text:p or self::text:h]/@text:style-name"/>
            <xsl:with-param name="context">
              <xsl:choose>
                <xsl:when test="ancestor::office:document-content">content.xml</xsl:when>
                <xsl:otherwise>styles.xml</xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:if test="number($pageNumber)">
          <xsl:attribute name="psect:page-number">
            <xsl:value-of select="$pageNumber"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- Look for a master-page-name into the style hierarchy 
    starting from 'style-name' and 'context' 
  -->
  <xsl:template name="GetMasterPageNameFromHierarchy">
    <xsl:param name="style-name"/>
    <xsl:param name="context" select="'content.xml'"/>
    <xsl:variable name="exists">
      <xsl:for-each select="document($context)">
        <xsl:value-of select="boolean(key('styles', $style-name))"/>
      </xsl:for-each>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$exists = 'true' ">
        <xsl:for-each select="document($context)">
          <xsl:variable name="style" select="key('styles', $style-name)[1]"/>
          <xsl:choose>
            <xsl:when test="$style/@style:master-page-name">
              <xsl:value-of select="$style/@style:master-page-name"/>
            </xsl:when>
            <xsl:when test="$style/@style:parent-style-name">
              <xsl:if test="$style/@style:parent-style-name != $style-name">
                <xsl:call-template name="GetMasterPageNameFromHierarchy">
                  <xsl:with-param name="style-name" select="$style/@style:parent-style-name"/>
                  <xsl:with-param name="context" select="$context"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:when>
          </xsl:choose>
        </xsl:for-each>
      </xsl:when>
      <!-- switch the context, let's look into styles.xml -->
      <xsl:when test="$context != 'styles.xml'">
        <xsl:call-template name="GetMasterPageNameFromHierarchy">
          <xsl:with-param name="style-name" select="$style-name"/>
          <xsl:with-param name="context" select="'styles.xml'"/>
        </xsl:call-template>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- find the number if it is set to restart -->
  <xsl:template name="GetPageStartNumber">
    <xsl:param name="style-name"/>
    <xsl:param name="context" select="'content.xml'"/>
    <xsl:variable name="exists">
      <xsl:for-each select="document($context)">
        <xsl:value-of select="boolean(key('styles', $style-name))"/>
      </xsl:for-each>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$exists = 'true' ">
        <xsl:for-each select="document($context)">
          <xsl:variable name="style" select="key('styles', $style-name)[1]"/>
          <xsl:choose>
            <xsl:when test="$style/style:paragraph-properties/@style:page-number &gt; 0">
              <xsl:value-of select="$style/style:paragraph-properties/@style:page-number"/>
            </xsl:when>
            <xsl:when test="$style/@style:parent-style-name">
              <xsl:if test="$style/@style:parent-style-name != $style-name">
                <xsl:call-template name="GetPageStartNumber">
                  <xsl:with-param name="style-name" select="$style/@style:parent-style-name"/>
                  <xsl:with-param name="context" select="$context"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:when>
          </xsl:choose>
        </xsl:for-each>
      </xsl:when>
      <!-- switch the context, let's look into styles.xml -->
      <xsl:when test="$context != 'styles.xml'">
        <xsl:call-template name="GetPageStartNumber">
          <xsl:with-param name="style-name" select="$style-name"/>
          <xsl:with-param name="context" select="'styles.xml'"/>
        </xsl:call-template>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- section detection and insertion for paragraph-->
  <xsl:template name="InsertParagraphSectionProperties">

    <xsl:if
      test="not(ancestor::table:table) and not(ancestor::draw:frame) and not(ancestor::draw:line) and not(ancestor::draw:rect) and not(ancestor::style:master-page)">
      <!-- Section detection  : 4 cases -->
      <!-- 1 - Following neighbour's (ie paragraph, heading or table) with non-empty reference to a master page  -->
      <xsl:variable name="followings"
        select="following::*[name()='text:p' or name()='text:h' or name()='table:table'][1]"/>

      <xsl:variable name="next-master-page">
        <xsl:choose>
          <xsl:when test="$followings[1]/@text:style-name">
            <xsl:call-template name="GetMasterPageNameFromHierarchy">
              <xsl:with-param name="style-name" select="$followings[1]/@text:style-name"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="$followings[1]/@table:style-name">
            <xsl:call-template name="GetMasterPageNameFromHierarchy">
              <xsl:with-param name="style-name" select="$followings[1]/@table:style-name"/>
            </xsl:call-template>
          </xsl:when>
        </xsl:choose>
      </xsl:variable>

      <!-- 2 - Section starts. The following paragraph is contained in the following section -->
      <xsl:variable name="following-section" select="following::text:section[1]"/>
      <!-- the following section is the same as the following neighbour's ancestor section -->
      <xsl:variable name="next-new-section"
        select="$following-section and (generate-id($followings[1]/ancestor::text:section[1]) = generate-id($following-section))"/>

      <!-- 3 - Section ends. We are in a section and the following paragraph isn't -->
      <xsl:variable name="previous-section" select="ancestor::text:section[1]"/>
      <!-- the following neighbour's ancestor section and the current section are different -->
      <xsl:variable name="next-end-section"
        select="$previous-section and not(generate-id($followings[1]/ancestor::text:section[1]) = generate-id($previous-section))"/>

      <!-- 4 - Detect a next-page-break -->
      <xsl:variable name="next-page-break">
        <xsl:call-template name="isPageBroken">
          <xsl:with-param name="following-elt" select="$followings[1]"/>
          <xsl:with-param name="self-style-name" select="@text:style-name"/>
        </xsl:call-template>
      </xsl:variable>


      <!-- section creation 
        A next-page-break or a master page start (not nested inside a text:section)
        or a section start or end not nested inside another section
        And there mustn't exist a text:note-body or table:table ancestor
      -->
      <xsl:if
        test="($next-page-break='true' or $next-master-page != '' or $next-new-section = 'true' or $next-end-section = 'true' ) 
        and not(ancestor::text:note-body or ancestor::table:table or ancestor::*[substring(name(), 1, 5) = 'draw:'])">
        <!--dialogika, clam: section breaks have their own paragraphs now (bug #1615686)-->
		<!-- 20080711/divo: Fix for #2014947: Don't add section breaks inside shapes, i.e. when there is an ancestor in the draw namespace -->
        <w:p>
          <w:pPr>
            <w:sectPr>
              <xsl:if test="$next-master-page != '' ">
                <xsl:attribute name="psect:next-master-page">true</xsl:attribute>
              </xsl:if>
              <xsl:if test="$next-page-break = 'true' ">
                <xsl:attribute name="psect:next-page-break">true</xsl:attribute>
              </xsl:if>
              <xsl:if test="$next-new-section = 'true' ">
                <xsl:attribute name="psect:next-new-section">true</xsl:attribute>
              </xsl:if>
              <xsl:if test="$next-end-section = 'true' ">
                <xsl:attribute name="psect:next-end-section">true</xsl:attribute>
                <xsl:apply-templates
                  select="key('sections', $previous-section/@text:style-name)[1]/style:section-properties/text:notes-configuration"
                  mode="note"/>
                <xsl:apply-templates
                  select="key('sections', $previous-section/@text:style-name)/style:section-properties"
                  mode="section"/>
              </xsl:if>
            </w:sectPr>
          </w:pPr>
        </w:p>
      </xsl:if>
    </xsl:if>
  </xsl:template>



  <!-- Manages sections within TABLES -->
  <xsl:template name="ManageSectionsInTable">

    <!-- Section detection  : 3 cases -->
    <xsl:if
      test="not(ancestor::table:table) and not(ancestor::draw:frame) and not(ancestor::draw:line) and not (ancestor::draw:rect) and not(ancestor::style:master-page)">
      <!-- 1 - Following neighbour's (ie paragraph, heading or table) master style  -->
      <xsl:variable name="followings"
        select="following::*[name()='text:p' or name()='text:h' or name()='table:table'][1]"/>

      <xsl:variable name="next-master-page">
        <xsl:choose>
          <xsl:when test="$followings[1]/@text:style-name">
            <xsl:call-template name="GetMasterPageNameFromHierarchy">
              <xsl:with-param name="style-name" select="$followings[1]/@text:style-name"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="$followings[1]/@table:style-name">
            <xsl:call-template name="GetMasterPageNameFromHierarchy">
              <xsl:with-param name="style-name" select="$followings[1]/@table:style-name"/>
            </xsl:call-template>
          </xsl:when>
        </xsl:choose>
      </xsl:variable>

      <!-- 2 - Section starts. The following paragraph is contained in the following section -->
      <xsl:variable name="following-section" select="following::text:section[1]"/>
      <!-- the following section is the same as the following neighbour's ancestor section -->
      <xsl:variable name="next-new-section"
        select="$following-section and (generate-id($followings[1]/ancestor::text:section[1]) = generate-id($following-section))"/>

      <!-- 3 - Section ends. We are in a section and the following paragraph isn't -->
      <xsl:variable name="previous-section" select="ancestor::text:section[1]"/>
      <!-- the following neighbour's ancestor section and the current section are different -->
      <xsl:variable name="next-end-section"
        select="$previous-section and not(generate-id($followings[1]/ancestor::text:section[1]) = generate-id($previous-section))"/>

      <xsl:variable name="next-page-break">
        <xsl:call-template name="isPageBroken">
          <xsl:with-param name="following-elt" select="$followings[1]"/>
          <xsl:with-param name="self-style-name" select="@table:style-name"/>
        </xsl:call-template>
      </xsl:variable>


      <xsl:if
        test="($next-master-page != '' or $next-new-section = 'true' or $next-end-section = 'true')
         and not(ancestor::text:note-body)">
        <w:p>
          <w:pPr>
            <w:sectPr>
              <xsl:if test="$next-master-page != '' ">
                <xsl:attribute name="psect:next-master-page">true</xsl:attribute>
              </xsl:if>
              <xsl:if test="$next-page-break = 'true' ">
                <xsl:attribute name="psect:next-page-break">true</xsl:attribute>
              </xsl:if>
              <xsl:if test="$next-new-section = 'true' ">
                <xsl:attribute name="psect:next-new-section">true</xsl:attribute>
              </xsl:if>
              <xsl:if test="$next-end-section = 'true' ">
                <xsl:attribute name="psect:next-end-section">true</xsl:attribute>
                <xsl:apply-templates
                  select="key('sections', $previous-section/@text:style-name)[1]/style:section-properties/text:notes-configuration"
                  mode="note"/>
                <xsl:apply-templates
                  select="key('sections', $previous-section/@text:style-name)/style:section-properties"
                  mode="section"/>
              </xsl:if>
            </w:sectPr>
          </w:pPr>
        </w:p>
      </xsl:if>
    </xsl:if>
  </xsl:template>



  <!-- a page break after the following table -->
  <xsl:template name="isPageBroken">
    <xsl:param name="following-elt"/>
    <xsl:param name="self-style-name"/>

    <xsl:for-each select="document('content.xml')">
      <xsl:variable name="myMasterPageName">
        <xsl:call-template name="GetMasterPageNameFromHierarchy">
          <xsl:with-param name="style-name" select="$self-style-name"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="nextPageName">
        <xsl:call-template name="GetMasterPageNameFromHierarchy">
          <xsl:with-param name="style-name" select="$following-elt/@text:style-name"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:choose>
        <xsl:when
          test="key('automatic-styles', $self-style-name)[1]/child::*/@fo:break-after = 'page' "
          >true</xsl:when>
        <xsl:when
          test="$following-elt and key('automatic-styles', $following-elt/@text:style-name|$following-elt/@table:style-name)[1]/child::*/@fo:break-before = 'page'"
          >true</xsl:when>
        <!--clam bugfix #1802267-->
        <xsl:when test="$following-elt and not($nextPageName = '') and not($nextPageName = $myMasterPageName)">
          <xsl:text>true</xsl:text>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="document('styles.xml')">
            <xsl:choose>
              <xsl:when test="key('styles', $self-style-name)[1]/child::*/@fo:break-after = 'page' "
                >true</xsl:when>
              <xsl:when
                test="$following-elt and key('styles', $following-elt/@text:style-name|$following-elt/@table:style-name)[1]/child::*/@fo:break-after = 'page'"
                > true </xsl:when>
              <xsl:otherwise>false</xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>

  </xsl:template>

  <!-- Preprocess oox sections properties -->
  <xsl:template name="sectionsPreProcessing">
    <psect:master-pages>
      <xsl:for-each select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page">
        <psect:master-page psect:name="{@style:name}">
          <xsl:if test="@style:next-style-name">
            <xsl:attribute name="psect:next-style">
              <xsl:value-of select="@style:next-style-name"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:call-template name="HeaderFooter">
            <xsl:with-param name="master-page" select="."/>
          </xsl:call-template>
          <xsl:apply-templates
            select="document('styles.xml')/office:document-styles/office:styles/text:notes-configuration"
            mode="note"/>

          <!--
          Context Switch For Each 
          -->
          <xsl:variable name="header" select="style:header" />
          <xsl:variable name="footer" select="style:footer" />
          <xsl:for-each select="key('page-layouts', @style:page-layout-name)[1]/style:page-layout-properties">
            <xsl:call-template name="InsertPageLayoutProperties" >
              <xsl:with-param name="hasHeader" select="$header" />
              <xsl:with-param name="hasFooter" select="$footer" />
            </xsl:call-template>
          </xsl:for-each>


          <!--clam, dialogika: bugfix 1947998-->
          <xsl:variable name="all-master-pages" select="document('styles.xml')//style:master-page"></xsl:variable>
          <xsl:if test ="//text:page-number[@text:select-page = 'previous']">
            <xsl:choose>
              <xsl:when test ="@style:next-style-name = 'Standard'">
                <w:pgNumType w:start="0"/>
              </xsl:when>
              <xsl:when test="count($all-master-pages) = 1">
                <w:pgNumType w:start="0"/>
              </xsl:when>
            </xsl:choose>           
          </xsl:if>
          
        </psect:master-page>
      </xsl:for-each>
    </psect:master-pages>
  </xsl:template>



  <!-- manage envelope -->
  <xsl:template name="InsertEnvelopeFrames">
    <xsl:if test="parent::office:text and not(preceding-sibling::*[self::text:p or self::text:h])">
      <xsl:variable name="master-page-name">
        <xsl:call-template name="GetMasterPageNameFromHierarchy">
          <xsl:with-param name="style-name" select="@text:style-name"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:if test="$master-page-name = 'Envelope' ">
        <xsl:for-each
          select="parent::office:text/draw:frame[not(@text:anchor-page-number) or @text:anchor-page-number='1']">
          <xsl:choose>
            <!-- sender field -->
            <xsl:when test="descendant::*[self::text:p or self::text:h]/@text:style-name='Sender' ">
              <xsl:for-each select="draw:text-box/child::node()">
                <xsl:choose>
                  <!--   ignore embedded text-box and draw:rect becouse word doesn't support it-->
                  <xsl:when test="self::node()[name(draw:text-box|draw:rect)]">
                    <xsl:message terminate="no">translation.odf2oox.nestedFrames</xsl:message>
                  </xsl:when>
                  <!-- warn loss of positioning for embedded drawn objects or pictures -->
                  <xsl:when test="contains(name(), 'draw:')">
                    <xsl:message terminate="no">translation.odf2oox.positionInsideTextbox</xsl:message>
                    <w:p>
                      <xsl:apply-templates select="." mode="paragraph"/>
                    </w:p>
                  </xsl:when>
                  <!--default scenario-->
                  <xsl:otherwise>
                    <xsl:apply-templates select="."/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:for-each>
            </xsl:when>
            <xsl:otherwise>
              <!-- adress field -->
              <xsl:if
                test="descendant::*[self::text:p or self::text:h]/@text:style-name='Addressee' ">
                <xsl:for-each select="draw:text-box/child::node()">
                  <xsl:choose>
                    <!--   ignore embedded text-box and draw:rect becouse word doesn't support it-->
                    <xsl:when test="self::node()[name(draw:text-box|draw:rect)]">
                      <xsl:message terminate="no">translation.odf2oox.nestedFrames</xsl:message>
                    </xsl:when>
                    <!-- warn loss of positioning for embedded drawn objects or pictures -->
                    <xsl:when test="contains(name(), 'draw:')">
                      <xsl:message terminate="no">translation.odf2oox.positionInsideTextbox</xsl:message>
                      <w:p>
                        <xsl:apply-templates select="." mode="paragraph"/>
                      </w:p>
                    </xsl:when>
                    <!--default scenario-->
                    <xsl:otherwise>
                      <xsl:apply-templates select="."/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:for-each>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
  </xsl:template>


  
</xsl:stylesheet>

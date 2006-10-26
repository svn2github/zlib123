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




  <!-- Document final section properties -->
  <xsl:template name="InsertDocumentFinalSectionProperties">
    <w:sectPr/>
  </xsl:template>
  
  <xsl:template name="MarkMasterPage">
    <xsl:choose>
      <xsl:when test="self::text:p or self::text:h or self::table:table">
        <xsl:if test="key('master-based-styles', @text:style-name|@table:style-name)[1]/@style:master-page-name != '' ">
          <xsl:attribute name="psect:master-page-name">
            <xsl:value-of select="key('master-based-styles', @text:style-name|@table:style-name)[1]/@style:master-page-name"/>
          </xsl:attribute>    
        </xsl:if>
      </xsl:when>
      <xsl:when test="self::text:list-item or self::text:list-header">
        <xsl:variable name="elt" select="*[1][self::text:p or self::text:h]"/>
        <xsl:if test="key('master-based-styles', $elt/@text:style-name)[1]/@style:master-page-name != '' ">
          <xsl:attribute name="psect:master-page-name">
            <xsl:value-of select="key('master-based-styles', $elt/@text:style-name)[1]/@style:master-page-name"/>
          </xsl:attribute>  
        </xsl:if>
      </xsl:when>
    </xsl:choose>
    
  </xsl:template>


  <!-- section detection and insertion for paragraph-->
  <xsl:template name="InsertParagraphSectionProperties">
    
    <xsl:if
      test="not(ancestor::table:table) and not(ancestor::draw:frame) and not(ancestor::draw:line) and not(ancestor::style:master-page)">
      <!-- Section detection  : 4 cases -->
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
      <xsl:variable name="ancestor-sections" select="ancestor::text:section"/>
      <xsl:variable name="previous-section" select="$ancestor-sections[1]"/>
      <!-- the following neighbour's ancestor section and the current section are different -->
      <xsl:variable name="section-ends"
        select="$previous-section and not(generate-id($followings[1]/ancestor::text:section[1]) = generate-id($previous-section))"/>

      <!-- 4 - Detect a page-break -->
      <xsl:variable name="page-break">
        <xsl:call-template name="isPageBroken">
          <xsl:with-param name="following-elt" select="$followings[1]"/>
          <xsl:with-param name="self-style-name" select="@text:style-name"/>
        </xsl:call-template>
      </xsl:variable>


      <!-- section creation 
        A page-break or a master page start (not nested inside a text:section)
        or a section start or end not nested inside another section
        And there mustn't exist a text:note-body or table:table ancestor
      -->
      <xsl:if
        test="((($page-break='true' or $master-page-starts = 'true') ) 
        or (($section-starts = 'true' or $section-ends = 'true') and count($ancestor-sections) &lt; 2)) 
        and not(ancestor::text:note-body or ancestor::table:table)">
        <w:sectPr>
          <xsl:if test="$page-break = 'true' ">
            <xsl:attribute name="psect:page-break">true</xsl:attribute>
          </xsl:if>
          <xsl:if test="$section-starts = 'true' ">
            <xsl:attribute name="psect:section-starts">true</xsl:attribute>
          </xsl:if>
          <xsl:if test="$section-ends = 'true' ">
            <xsl:attribute name="psect:section-ends">true</xsl:attribute>
            <xsl:apply-templates 
              select="key('sections', $previous-section/@text:style-name)[1]/style:section-properties/text:notes-configuration" 
              mode="note"/>
            <xsl:apply-templates
              select="key('sections', $previous-section/@text:style-name)/style:section-properties"
              mode="section"/>
          </xsl:if>
          </w:sectPr>
        
      </xsl:if>
    </xsl:if>
  </xsl:template>

 

  <!-- Manages sections within TABLES -->
  <xsl:template name="ManageSectionsInTable">
    
    <!-- Section detection  : 3 cases -->
    <xsl:if
      test="not(ancestor::table:table) and not(ancestor::draw:frame) and not(ancestor::draw:line) and not(ancestor::style:master-page)">
      <!-- 1 - Following neighbour's (ie paragraph, heading or table) master style  -->
      <xsl:variable name="followings"
        select="following::text:p[1] | following::text:h[1] | following::table:table[1]"/>
      <xsl:variable name="master-page-name"
        select="key('master-based-styles', $followings[1]/@text:style-name | $followings[1]/@table:style-name)[1]/@style:master-page-name"/>
      <xsl:variable name="masterPageStarts"
        select="boolean($master-page-name != '')"/>

      <!-- 2 - Section starts. The following paragraph is contained in the following section -->
      <xsl:variable name="followingSection" select="following::text:section[1]"/>
      <!-- the following section is the same as the following neighbour's ancestor section -->
      <xsl:variable name="section-starts"
        select="$followingSection and (generate-id($followings[1]/ancestor::text:section[1]) = generate-id($followingSection))"/>

      <!-- 3 - Section ends. We are in a section and the following paragraph isn't -->
      <xsl:variable name="ancestor-sections" select="ancestor::text:section"/>
      <xsl:variable name="previous-section" select="$ancestor-sections[1]"/>
      <!-- the following neighbour's ancestor section and the current section are different -->
      <xsl:variable name="section-ends"
        select="$previous-section and not(generate-id($followings[1]/ancestor::text:section[1]) = generate-id($previous-section))"/>

      <xsl:variable name="page-break">
        <xsl:call-template name="isPageBroken">
          <xsl:with-param name="following-elt" select="$followings[1]"/>
          <xsl:with-param name="self-style-name" select="@table:style-name"/>
        </xsl:call-template>
      </xsl:variable>

      
      <xsl:if
        test="(($masterPageStarts = 'true')
        or (($section-starts = 'true' or $section-ends = 'true') and count($ancestor-sections) &lt; 2))
         and not(ancestor::text:note-body)">
          <w:p><w:pPr>
            <w:sectPr>
              <xsl:if test="$page-break = 'true' ">
                <xsl:attribute name="psect:page-break">true</xsl:attribute>
              </xsl:if>
              <xsl:if test="$section-starts = 'true' ">
                <xsl:attribute name="psect:section-starts">true</xsl:attribute>
              </xsl:if>
              <xsl:if test="$section-ends = 'true' ">
                <xsl:attribute name="psect:section-ends">true</xsl:attribute>
                <xsl:apply-templates 
                  select="key('sections', $previous-section/@text:style-name)[1]/style:section-properties/text:notes-configuration" 
                  mode="note"/>
                <xsl:apply-templates
                  select="key('sections', $previous-section/@text:style-name)/style:section-properties"
                  mode="section"/>
              </xsl:if>
            </w:sectPr>
          </w:pPr></w:p>
      </xsl:if>
    </xsl:if>
  </xsl:template>



  <!-- a page break after the following table -->
  <xsl:template name="isPageBroken">
    <xsl:param name="following-elt"/>
    <xsl:param name="self-style-name"/>

    <xsl:for-each select="document('content.xml')">
      <xsl:choose>
        <xsl:when
          test="key('automatic-styles', $self-style-name)[1]/child::*/@fo:break-after = 'page' "
          >true</xsl:when>
        <xsl:when
          test="$following-elt and key('automatic-styles', $following-elt/@text:style-name|$following-elt/@table:style-name)[1]/child::*/@fo:break-before = 'page'"
          >true</xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="document('styles.xml')">
            <xsl:choose>
              <xsl:when
                test="key('styles', $self-style-name)[1]/child::*/@fo:break-after = 'page' "
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
      <xsl:for-each
        select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page">
        <psect:master-page psect:name="{@style:name}">
          <xsl:if  test="@style:next-style-name">  
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
          <xsl:apply-templates
            select="key('page-layouts', @style:page-layout-name)[1]/style:page-layout-properties" mode="master-page"/>
        </psect:master-page>
      </xsl:for-each>
    </psect:master-pages>
  </xsl:template>


</xsl:stylesheet>

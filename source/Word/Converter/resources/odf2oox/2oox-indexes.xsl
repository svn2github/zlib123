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
  exclude-result-prefixes="office text table fo style draw xlink v svg number">


  <xsl:strip-space elements="*"/>
  <xsl:preserve-space elements="text:p"/>
  <xsl:preserve-space elements="text:span"/>


  <xsl:key name="toc" match="text:table-of-content" use="''"/>
  <xsl:key name="indexes" match="text:illustration-index | text:table-index" use="''"/>
  <xsl:key name="user-indexes" match="text:user-index" use="''"/>
  <xsl:key name="user-index-by-name" match="text:user-index/text:user-index-source"
    use="@text:index-name"/>
  <xsl:key name="alphabetical-indexes" match="text:alphabetical-index" use="''"/>
  <xsl:key name="bibliography-entries" match="text:bibliography-mark" use="@text:identifier"/>
  <xsl:key name="index-styles" match="text:table-of-content-source/*" use="@text:style-name"/>

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
        <!-- apply templates to nodes except tabs who do not have preceding sibling other than tabs (converted into indent) -->
        <xsl:apply-templates
          select="child::node()[not(self::text:tab[not(preceding-sibling::node()[not(self::text:tab)])])]"
          mode="paragraph"/>
      </xsl:when>
      <!-- default scenario -->
      <xsl:otherwise>
        <xsl:apply-templates mode="paragraph"/>
      </xsl:otherwise>
    </xsl:choose>

    <!-- inserts field code end in last index element -->
    <xsl:if test="(count(following-sibling::text:p) = 0) and parent::text:index-body">
      <xsl:call-template name="InsertIndexFieldCodeEnd"/>
    </xsl:if>

  </xsl:template>

  <!-- end field -->
  <xsl:template name="InsertIndexFieldCodeEnd">
    <w:r>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
  </xsl:template>

  <!-- simple start field -->
  <xsl:template name="InsertIndexFieldCodeSimpleStart">
    <w:r>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
  </xsl:template>

  <!-- complex start field -->
  <xsl:template name="InsertIndexFieldCodeStart">
    <xsl:call-template name="InsertIndexFieldCodeSimpleStart"/>
    <w:r>
      <xsl:choose>
        <xsl:when test="ancestor::text:table-of-content">
          <xsl:call-template name="InsertTocPrefs"/>
        </xsl:when>
        <xsl:when test="ancestor::text:user-index">
          <xsl:call-template name="InsertUserIndexPrefs"/>
        </xsl:when>
        <xsl:when test="ancestor::text:illustration-index">
          <xsl:call-template name="InsertIllustrationInPrefs"/>
        </xsl:when>
        <xsl:when test="ancestor::text:alphabetical-index">
          <xsl:call-template name="insertAlphabeticalPrefs"/>
        </xsl:when>
        <xsl:when test="ancestor::text:bibliography">
          <xsl:call-template name="InsertBibliographyPrefs"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="InsertIndexFiguresPrefs"/>
        </xsl:otherwise>
      </xsl:choose>
    </w:r>
    <w:r>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
  </xsl:template>

  <xsl:template name="InsertIndexFiguresPrefs">
    <w:instrText xml:space="preserve"> TOC \c "</w:instrText>
    <w:instrText>
      <xsl:value-of
        select="parent::text:index-body/preceding-sibling::text:table-index-source/@text:caption-sequence-name"
      />
    </w:instrText>
    <w:instrText xml:space="preserve">" </w:instrText>
    <!-- no page numbering if not defined in index -->
    <xsl:if test="not(parent::text:index-body/preceding-sibling::*/*/text:index-entry-page-number)">
      <w:instrText xml:space="preserve">\n </w:instrText>
    </xsl:if>
    <xsl:if test="not(parent::text:index-body/preceding-sibling::*/*/text:index-entry-tab-stop[@style:type = 'right'])">
      <w:instrText xml:space="preserve">\p " " </w:instrText>
    </xsl:if>
    <!-- caption-format = 'text' is default. ='category-and-value' not handled -->
    <xsl:if test="parent::text:index-body/preceding-sibling::*/@text:caption-sequence-format = 'caption' ">
      <w:instrText xml:space="preserve">\a </w:instrText>
    </xsl:if>
    <xsl:if test="parent::text:index-body/preceding-sibling::*/@text:caption-sequence-format = 'category-and-value' ">
      <xsl:message terminate="no">"translation.odf2oox.TableIllustrationCaptionFormat</xsl:message>
    </xsl:if>
  </xsl:template>

  <!-- alphabetical index -->
  <xsl:template name="insertAlphabeticalPrefs">
    <w:instrText xml:space="preserve">INDEX </w:instrText>

    <!--Right Align Page Number-->
    <xsl:if
      test="key('styles', @text:style-name)/style:paragraph-properties/style:tab-stops/style:tab-stop/@style:type='right' ">
      <w:instrText xml:space="preserve">\e "</w:instrText>
      <w:tab/>
      <w:instrText xml:space="preserve">" </w:instrText>
    </xsl:if>

    <!-- column number -->
    <xsl:choose>
      <xsl:when
        test="key('styles', ancestor-or-self::text:alphabetical-index/@text:style-name)/style:section-properties/style:columns/@fo:column-count >4">
        <xsl:message terminate="no">translation.odf2oox.alphabeticalIndexColumnNumber</xsl:message>
        <w:instrText xml:space="preserve">\c "4" </w:instrText>
      </xsl:when>
      <xsl:when
        test="key('styles', ancestor-or-self::text:alphabetical-index/@text:style-name)/style:section-properties/style:columns/@fo:column-count >1">
        <w:instrText xml:space="preserve">\c "</w:instrText>
        <w:instrText>
          <xsl:value-of
            select="key('styles', ancestor-or-self::text:alphabetical-index/@text:style-name)/style:section-properties/style:columns/@fo:column-count"
          />
        </w:instrText>
        <w:instrText xml:space="preserve">" </w:instrText>
      </xsl:when>
      <xsl:otherwise>
        <w:instrText xml:space="preserve">\c "1" </w:instrText>
      </xsl:otherwise>
    </xsl:choose>

    <!-- language -->
    <xsl:if
      test="ancestor-or-self::text:alphabetical-index/text:alphabetical-index-source/@fo:language">
      <w:instrText xml:space="preserve">\z "</w:instrText>
      <w:instrText>
        <xsl:value-of
          select="ancestor-or-self::text:alphabetical-index/text:alphabetical-index-source/@fo:language"
        />
      </w:instrText>
      <w:instrText xml:space="preserve">"</w:instrText>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertIllustrationInPrefs">
    <w:instrText xml:space="preserve"> TOC  \c "</w:instrText>
    <w:instrText>
      <xsl:value-of
        select="parent::text:index-body/preceding-sibling::text:illustration-index-source/@text:caption-sequence-name"
      />
    </w:instrText>
    <w:instrText xml:space="preserve">" </w:instrText>
  </xsl:template>

  <!-- table of content -->
  <xsl:template name="InsertTocPrefs">
    <xsl:variable name="tocSource"
      select="ancestor::text:table-of-content/text:table-of-content-source"/>

    <w:instrText xml:space="preserve"> TOC </w:instrText>
    <!-- outline level -->
    <xsl:if test="$tocSource/@text:outline-level">
      <w:instrText xml:space="preserve">\o "1-</w:instrText>
      <w:instrText>
        <!-- include elements with outline styles up to selected level  -->
        <xsl:choose>
          <xsl:when test="$tocSource/@text:outline-level=10">9</xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$tocSource/@text:outline-level"/>
          </xsl:otherwise>
        </xsl:choose>
      </w:instrText>
      <w:instrText xml:space="preserve">" </w:instrText>
    </xsl:if>

    <!-- separator before page numbering. default is right align, null if no tab-stop defined -->
    <xsl:if test="$tocSource/text:table-of-content-entry-template and not($tocSource/text:table-of-content-entry-template/text:index-entry-tab-stop[@style:type = 'right'])">
      <w:instrText xml:space="preserve">\p " " </w:instrText>
    </xsl:if>

    <!--include index marks-->
    <xsl:if test="$tocSource[@text:use-index-marks] and not($tocSource[@text:use-index-marks = 'false'])">
      <w:instrText xml:space="preserve">\u </w:instrText>
    </xsl:if>

    <!--use hyperlinks -->
    <xsl:if test="text:a">
      <w:instrText xml:space="preserve">\h </w:instrText>
    </xsl:if>

    <!-- include elements with additional styles-->
    <xsl:if test="$tocSource/text:index-source-styles">
      <w:instrText xml:space="preserve">\t "</w:instrText>
      <w:instrText>
        <xsl:call-template name="InsertTOCLevelStyle">
          <xsl:with-param name="tocSource" select="$tocSource"/>
        </xsl:call-template>
      </w:instrText>
      <w:instrText xml:space="preserve">" </w:instrText>
    </xsl:if>
  </xsl:template>

  <!-- user defined index -->
  <xsl:template name="InsertUserIndexPrefs">
    <xsl:variable name="tocSource" select="ancestor::text:user-index/text:user-index-source"/>

    <w:instrText xml:space="preserve"> TOC </w:instrText>

    <!-- id to associate TC fields to this TOC -->
    <xsl:if test="$tocSource/@text:index-name">
      <w:instrText xml:space="preserve">\f "</w:instrText>
      <w:instrText>
        <xsl:value-of select="$tocSource/@text:index-name"/>
      </w:instrText>
      <w:instrText xml:space="preserve">" </w:instrText>
    </xsl:if>

    <!-- include elements with additional styles-->
    <xsl:if
      test="$tocSource/text:index-source-styles and $tocSource/@text:use-index-source-styles='true' ">
      <w:instrText xml:space="preserve">\t "</w:instrText>
      <w:instrText>
        <xsl:call-template name="InsertTOCLevelStyle">
          <xsl:with-param name="tocSource" select="$tocSource"/>
        </xsl:call-template>
      </w:instrText>
      <w:instrText xml:space="preserve">" </w:instrText>
    </xsl:if>

  </xsl:template>


  <xsl:template name="InsertTOCLevelStyle">
    <xsl:param name="level" select="1"/>
    <xsl:param name="tocSource"/>

    <xsl:if test="$level &lt; 10">
      <xsl:if
        test="$tocSource/text:index-source-styles[number(@text:outline-level) = number($level)]/text:index-source-style/@text:style-name">
        <xsl:for-each
          select="$tocSource/text:index-source-styles[number(@text:outline-level) = number($level)]/text:index-source-style">
          <xsl:variable name="levelStyleName" select="@text:style-name"/>
          <xsl:for-each select="document('styles.xml')">
            <xsl:for-each select="key('styles', $levelStyleName)">
              <xsl:choose>
                <xsl:when test="@style:display-name">
                  <xsl:value-of select="@style:display-name"/>
                  <xsl:text>;</xsl:text>
                  <xsl:value-of select="$level"/>
                  <xsl:text>;</xsl:text>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:if test="@style:name">
                    <xsl:value-of select="@style:name"/>
                    <xsl:text>;</xsl:text>
                    <xsl:value-of select="$level"/>
                    <xsl:text>;</xsl:text>
                  </xsl:if>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:for-each>
          </xsl:for-each>
        </xsl:for-each>
      </xsl:if>
      <!-- insert next level -->
      <xsl:call-template name="InsertTOCLevelStyle">
        <xsl:with-param name="level" select="$level + 1"/>
        <xsl:with-param name="tocSource" select="$tocSource"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>



  <xsl:template name="InsertIndexPageRefEnd">
    <w:r>
      <w:rPr>
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
        <w:noProof/>
        <w:webHidden/>
      </w:rPr>
      <w:fldChar w:fldCharType="begin">
        <w:fldData xml:space="preserve">CNDJ6nn5us4RjIIAqgBLqQsCAAAACAAAAA4AAABfAFQAbwBjADEANAAxADgAMwA5ADIANwA2AAAA</w:fldData>
      </w:fldChar>
    </w:r>
    <w:r>
      <w:rPr>
        <w:noProof/>
        <w:webHidden/>
      </w:rPr>
      <w:instrText xml:space="preserve"><xsl:value-of select="concat('PAGEREF _Toc', $tocId,generate-id(ancestor::node()[child::text:index-body]), ' \h')"/></w:instrText>
    </w:r>
    <w:r>
      <w:rPr>
        <w:noProof/>
        <w:webHidden/>
      </w:rPr>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
  </xsl:template>

  <!-- insert the bg color in paragraph properties -->
  <xsl:template name="InsertTOCBgColor">
    <xsl:if
      test="key('styles', ancestor::text:table-of-content/@text:style-name)/style:section-properties/@fo:background-color">
      <xsl:variable name="bgColor">
        <xsl:value-of
          select="key('styles', ancestor::text:table-of-content/@text:style-name)/style:section-properties/@fo:background-color"
        />
      </xsl:variable>
      <xsl:if test="$bgColor != 'transparent' ">
        <w:shd w:val="clear" w:color="auto"
          w:fill="{translate(translate(substring-after($bgColor, '#'),'f','F'),'c','C')}"/>
      </xsl:if>
    </xsl:if>
  </xsl:template>

  <!-- empty alphabetical indexes creating mark entry -->
  <xsl:template match="text:alphabetical-index-mark" mode="paragraph">
    <xsl:call-template name="InsertXEFieldInstructions">
      <xsl:with-param name="entryText" select="@text:string-value"/>
    </xsl:call-template>
  </xsl:template>

  <!-- alphabetical indexes creating mark entry -->
  <xsl:template match="text:alphabetical-index-mark-end" mode="paragraph">
    <xsl:variable name="id" select="@text:id"/>
    <xsl:variable name="entryText">
      <xsl:for-each
        select="preceding-sibling::node()[preceding-sibling::text:alphabetical-index-mark-start[@text:id = $id]]">
        <!-- ignore all ...mark-start/end and track-changes -->
        <xsl:if test="not(contains(name(), 'mark-') or contains(name(), 'change-'))">
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
    </xsl:variable>
    <xsl:for-each select="preceding-sibling::text:alphabetical-index-mark-start[@text:id = $id]">
      <xsl:call-template name="InsertXEFieldInstructions">
        <xsl:with-param name="entryText" select="$entryText"/>
      </xsl:call-template>
    </xsl:for-each>
  </xsl:template>

  <!-- insert field instruction for XE index entry -->
  <xsl:template name="InsertXEFieldInstructions">
    <xsl:param name="entryText"/>
    <xsl:call-template name="InsertIndexFieldCodeSimpleStart"/>
    <w:r>
      <w:instrText xml:space="preserve"> XE "</w:instrText>
      <w:instrText>
        <xsl:variable name="key1">
          <xsl:if test="@text:key1">
            <xsl:value-of select="concat(@text:key1, ':')"/>
          </xsl:if>
        </xsl:variable>
        <xsl:variable name="key2">
          <xsl:if test="@text:key2">
            <xsl:value-of select="concat(@text:key2, ':')"/>
          </xsl:if>
        </xsl:variable>
        <xsl:value-of select="concat($key1, $key2, $entryText)"/>
      </w:instrText>
      <w:instrText xml:space="preserve">" </w:instrText>
      <!-- find style associated to main entries. If more than one index, use style of first only. -->
      <xsl:if test="@text:main-entry='true' ">
        <xsl:variable name="MainStyleName">
          <xsl:value-of
            select="key('alphabetical-indexes', '')/text:alphabetical-index-source/@text:main-entry-style-name"
          />
        </xsl:variable>
        <xsl:for-each select="document('styles.xml')">
          <xsl:for-each select="key('styles', $MainStyleName)/style:text-properties">
            <xsl:if test="@fo:font-weight = 'bold' ">
              <w:instrText xml:space="preserve">\b </w:instrText>
            </xsl:if>
            <xsl:if test="@fo:font-style = 'italic' ">
              <w:instrText xml:space="preserve">\i </w:instrText>
            </xsl:if>
          </xsl:for-each>
        </xsl:for-each>
      </xsl:if>
    </w:r>
    <xsl:call-template name="InsertIndexFieldCodeEnd"/>
  </xsl:template>

  <xsl:template match="text()" mode="indexes">
    <xsl:choose>
      <xsl:when test="ancestor::text:index-title">
        <w:t xml:space="preserve"><xsl:value-of select="."/></w:t>
      </xsl:when>
      <xsl:when test="preceding-sibling::text:tab">
        <w:t xml:space="preserve"><xsl:value-of select="."/></w:t>
      </xsl:when>
      <xsl:when test="not(following-sibling::text:tab)">
        <xsl:choose>
          <xsl:when test="parent::text:a|parent::text:span">
            <w:t xml:space="preserve"><xsl:value-of select="."/></w:t>
          </xsl:when>
          <xsl:otherwise>
            <w:t>
              <xsl:value-of select="ancestor::text:p/text()"/>
            </w:t>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <w:t>
          <xsl:value-of select="."/>
        </w:t>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- empty user indexes creating mark entry -->
  <xsl:template match="text:user-index-mark" mode="paragraph">
    <xsl:if test="key('user-index-by-name', @text:index-name)/@text:use-index-marks = 'true' ">
      <xsl:call-template name="InsertUserFieldInstructions">
        <xsl:with-param name="entryText" select="@text:string-value"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

  <!-- alphabetical indexes creating mark entry -->
  <xsl:template match="text:user-index-mark-end" mode="paragraph">
    <xsl:variable name="id" select="@text:id"/>
    <xsl:variable name="entryText">
      <xsl:for-each
        select="preceding-sibling::node()[preceding-sibling::text:user-index-mark-start[@text:id = $id]]">
        <!-- ignore all ...mark-start/end and track-changes -->
        <xsl:if test="not(contains(name(), 'mark-') or contains(name(), 'change-'))">
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
    </xsl:variable>
    <xsl:for-each select="preceding-sibling::text:user-index-mark-start[@text:id = $id]">
      <xsl:if test="key('user-index-by-name', @text:index-name)/@text:use-index-marks = 'true' ">
        <xsl:call-template name="InsertUserFieldInstructions">
          <xsl:with-param name="entryText" select="$entryText"/>
        </xsl:call-template>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

  <!-- insert field instruction for user index entry -->
  <xsl:template name="InsertUserFieldInstructions">
    <xsl:param name="entryText"/>
    <xsl:param name="isIndexMark">true</xsl:param>
    <xsl:call-template name="InsertIndexFieldCodeSimpleStart"/>
    <w:r>
      <w:instrText xml:space="preserve"> TC "</w:instrText>
      <w:instrText>
        <xsl:value-of select="$entryText"/>
      </w:instrText>
      <w:instrText xml:space="preserve">" </w:instrText>
      <!-- index id -->
      <w:instrText xml:space="preserve">\f "</w:instrText>
      <w:instrText>
        <xsl:choose>
          <xsl:when test="$isIndexMark = 'true' ">
            <xsl:value-of select="@text:index-name"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="text:user-index-source/@text:index-name"/>
          </xsl:otherwise>
        </xsl:choose>
      </w:instrText>
      <w:instrText xml:space="preserve">" </w:instrText>
      <!-- outline level -->
      <xsl:if test="$isIndexMark = 'true' ">
        <w:instrText xml:space="preserve">\l "</w:instrText>
        <w:instrText>
          <xsl:value-of select="@text:outline-level"/>
        </w:instrText>
        <w:instrText xml:space="preserve">" </w:instrText>
      </xsl:if>
    </w:r>
    <xsl:call-template name="InsertIndexFieldCodeEnd"/>
  </xsl:template>


  <!-- insert a TC field for various uses of user-defined-TOC -->
  <xsl:template name="InsertTCField">
    <xsl:choose>

      <!-- first case : first paragraph of a table -->
      <xsl:when
        test="(self::text:p or self::text:h) and key('user-indexes', '')/text:user-index-source/@text:use-tables='true' ">
        <xsl:variable name="isFirstParagraphOfTable">
          <xsl:call-template name="IsFirstParagraphOfTable"/>
        </xsl:variable>
        <xsl:if test="$isFirstParagraphOfTable = 'true' ">
          <xsl:variable name="entryText" select="ancestor-or-self::table:table[last()]/@table:name"/>
          <!-- insert a TC field for every index that uses tables -->
          <xsl:for-each
            select="key('user-indexes', '')[text:user-index-source/@text:use-tables='true']">
            <xsl:call-template name="InsertUserFieldInstructions">
              <xsl:with-param name="entryText" select="$entryText"/>
              <xsl:with-param name="isIndexMark">false</xsl:with-param>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:if>
      </xsl:when>

      <!-- images (shapes not supported yet) -->
      <xsl:when test="self::draw:frame/descendant::draw:image and not(ancestor::draw:frame)">
        <xsl:for-each select="self::draw:frame/descendant::draw:image">
          <xsl:variable name="entryText" select="parent::draw:frame/@draw:name"/>
          <xsl:for-each
            select="key('user-indexes', '')[text:user-index-source/@text:use-graphics='true']">
            <xsl:call-template name="InsertUserFieldInstructions">
              <xsl:with-param name="entryText" select="$entryText"/>
              <xsl:with-param name="isIndexMark">false</xsl:with-param>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>

      <!-- text-boxes (floating-frames not supported yet) -->
      <xsl:when test="self::draw:frame/descendant::draw:text-box and not(ancestor::draw:frame)">
        <xsl:for-each select="self::draw:frame/descendant::draw:text-box">
          <xsl:variable name="entryText" select="parent::draw:frame/@draw:name"/>
          <xsl:for-each
            select="key('user-indexes', '')[text:user-index-source/@text:use-floating-frames='true']">
            <xsl:call-template name="InsertUserFieldInstructions">
              <xsl:with-param name="entryText" select="$entryText"/>
              <xsl:with-param name="isIndexMark">false</xsl:with-param>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>

      <!-- OLE object (other objects not supported yet) -->
      <xsl:when test="self::draw:frame/descendant::draw:object-ole and not(ancestor::draw:frame)">
        <xsl:for-each select="self::draw:frame/descendant::draw:object-ole">
          <xsl:variable name="entryText" select="parent::draw:frame/@draw:name"/>
          <xsl:for-each
            select="key('user-indexes', '')[text:user-index-source/@text:use-objects='true']">
            <xsl:call-template name="InsertUserFieldInstructions">
              <xsl:with-param name="entryText" select="$entryText"/>
              <xsl:with-param name="isIndexMark">false</xsl:with-param>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>

      <xsl:otherwise/>
    </xsl:choose>
  </xsl:template>


  <!-- bibliography : defined like a TOC -->
  <xsl:template name="InsertBibliographyPrefs">
    <xsl:variable name="tocSource" select="ancestor::text:bibliography/text:bibliography-source"/>

    <w:instrText xml:space="preserve"> TOC </w:instrText>

    <!-- id to associate TC fields to this TOC. It should be unique -->
    <w:instrText xml:space="preserve">\f "</w:instrText>
    <w:instrText>
      <xsl:value-of select="concat('Bibliography_', generate-id(ancestor::office:body))"/>
    </w:instrText>
    <w:instrText xml:space="preserve">" </w:instrText>

    <!-- no page number -->
    <w:instrText xml:space="preserve">\n </w:instrText>
  </xsl:template>

  <!-- bibliography entry -->
  <xsl:template match="text:bibliography-mark" mode="paragraph">
    <xsl:variable name="ref" select="generate-id(.)"/>
    <!-- create an entry only for the first occurence of the reference -->
    <xsl:if test="$ref = generate-id(key('bibliography-entries', @text:identifier)[1])">
      <xsl:call-template name="InsertIndexFieldCodeSimpleStart"/>
      <w:r>
        <w:instrText xml:space="preserve"> TC "</w:instrText>
        <!-- entry text -->
        <xsl:variable name="bibliographyType" select="@text:bibliography-type"/>
        <xsl:call-template name="InsertBibliographyEntryText">
          <xsl:with-param name="bibliographyConfiguration"
            select="document('styles.xml')/office:document-styles/office:styles/text:bibliography-configuration"/>
          <xsl:with-param name="entryTemplate"
            select="ancestor::office:text/text:bibliography/text:bibliography-source/text:bibliography-entry-template[@text:bibliography-type=$bibliographyType]/child::node()"
          />
        </xsl:call-template>
        <w:instrText xml:space="preserve">" </w:instrText>
        <!-- index id -->
        <w:instrText xml:space="preserve">\f "</w:instrText>
        <w:instrText>
          <xsl:value-of select="concat('Bibliography_', generate-id(ancestor::office:body))"/>
        </w:instrText>
        <w:instrText xml:space="preserve">" </w:instrText>
      </w:r>
      <xsl:call-template name="InsertIndexFieldCodeEnd"/>
    </xsl:if>
    <xsl:apply-templates mode="paragraph"/>
  </xsl:template>

  <!-- insert the entry text of a bibliography entry -->
  <xsl:template name="InsertBibliographyEntryText">
    <xsl:param name="entryTemplate"/>
    <xsl:param name="bibliographyConfiguration"/>
    <xsl:choose>
      <xsl:when test="$entryTemplate[self::text:index-entry-tab-stop]">
        <w:instrText>
          <xsl:call-template name="ComputeBibliographyEntry">
            <xsl:with-param name="entryTemplate"
              select="$entryTemplate[not(self::text:index-entry-tab-stop or preceding-sibling::text:index-entry-tab-stop)]"/>
            <xsl:with-param name="bibliographyConfiguration" select="$bibliographyConfiguration"/>
          </xsl:call-template>
        </w:instrText>
        <w:tab/>
        <xsl:call-template name="InsertBibliographyEntryText">
          <xsl:with-param name="bibliographyConfiguration" select="$bibliographyConfiguration"/>
          <xsl:with-param name="entryTemplate"
            select="$entryTemplate[preceding-sibling::text:index-entry-tab-stop]"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <w:instrText>
          <xsl:call-template name="ComputeBibliographyEntry">
            <xsl:with-param name="entryTemplate" select="$entryTemplate"/>
            <xsl:with-param name="bibliographyConfiguration" select="$bibliographyConfiguration"/>
          </xsl:call-template>
        </w:instrText>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- compute the text of a bibliography entry -->
  <xsl:template name="ComputeBibliographyEntry">
    <xsl:param name="bibliographyConfiguration"/>
    <xsl:param name="entryTemplate"/>

    <xsl:if test="count($entryTemplate) &gt; 0">
      <xsl:choose>
        <xsl:when
          test="$entryTemplate[1][self::text:index-entry-bibliography/@text:bibliography-data-field = 'identifier']">
          <xsl:choose>
            <xsl:when test="$bibliographyConfiguration/@text:numbered-entries = 'true' ">
              <xsl:value-of
                select="substring-before(substring-after(text(), $bibliographyConfiguration/@text:prefix), $bibliographyConfiguration/@text:suffix)"
              />
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of
                select="concat($bibliographyConfiguration/@text:prefix, @text:identifier, $bibliographyConfiguration/@text:suffix)"
              />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:when
          test="$entryTemplate[1][self::text:index-entry-bibliography/@text:bibliography-data-field != 'identifier']">
          <xsl:value-of
            select="@*[name() = concat('text:', $entryTemplate[1]/@text:bibliography-data-field)]"/>
        </xsl:when>
        <xsl:when test="$entryTemplate[1][self::text:index-entry-span]">
          <xsl:value-of select="$entryTemplate[1]/text()"/>
        </xsl:when>
        <xsl:otherwise/>
      </xsl:choose>

      <!-- write next element -->
      <xsl:call-template name="ComputeBibliographyEntry">
        <xsl:with-param name="entryTemplate" select="$entryTemplate[position() &gt; 1]"/>
        <xsl:with-param name="bibliographyConfiguration" select="$bibliographyConfiguration"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

  <!-- styles for indexes. They require a particular syntax -->
  <xsl:template name="InsertIndexStyles">
    <xsl:for-each select="document('content.xml')">
      <xsl:for-each select="key('toc', '')[1]">
        <xsl:call-template name="InsertIndexLevelStyle"/>
      </xsl:for-each>
    </xsl:for-each>
  </xsl:template>

  <!-- there can be only one style for the whole document (all TOCs) -->
  <xsl:template name="InsertIndexLevelStyle">
    <xsl:param name="level" select="1"/>

    <xsl:if test="$level &lt; 10">
      <xsl:variable name="levelStyleName">
        <xsl:value-of
          select="text:table-of-content-source/text:table-of-content-entry-template[@text:outline-level = $level]/@text:style-name"
        />
      </xsl:variable>
      <!-- if hyperlink -->
      <xsl:variable name="levelTextStyleName">
        <xsl:value-of
          select="text:table-of-content-source/text:table-of-content-entry-template[@text:outline-level = $level]/*[self::text:index-entry-link-start or self::text:index-entry-link-end]/@text:style-name"
        />
      </xsl:variable>
      <w:style w:styleId="{concat('TOC', $level)}" w:type="paragraph">
        <w:name w:val="{concat('toc ', $level)}"/>
        <w:basedOn w:val="{$levelStyleName}"/>
        <w:autoRedefine/>
        <w:semiHidden/>
        <w:pPr>
          <xsl:for-each
            select="text:table-of-content-source/text:table-of-content-entry-template[@text:outline-level = $level]">
            <xsl:call-template name="OverrideIndexParagraphTabs">
              <xsl:with-param name="levelStyleName" select="$levelStyleName"/>
              <xsl:with-param name="level" select="$level"/>
            </xsl:call-template>
          </xsl:for-each>
        </w:pPr>
        <xsl:if test="$levelTextStyleName != '' ">
          <!-- change context -->
          <xsl:for-each select="document('styles.xml')">
            <xsl:for-each select="key('styles', $levelStyleName)">
              <w:rPr>
                <w:rStyle w:val="{$levelTextStyleName}"/>
                <xsl:for-each select="document('styles.xml')">
                  <xsl:apply-templates select="key('styles', $levelTextStyleName)" mode="rPr"/>
                </xsl:for-each>
              </w:rPr>
            </xsl:for-each>
          </xsl:for-each>
        </xsl:if>
      </w:style>
      <!-- insert next level -->
      <xsl:call-template name="InsertIndexLevelStyle">
        <xsl:with-param name="level" select="$level + 1"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

  <!-- override tabs for index -->
  <xsl:template name="OverrideIndexParagraphTabs">
    <xsl:param name="levelStyleName"/>
    <xsl:param name="level"/>

    <xsl:variable name="leftTabStop">
      <xsl:if
        test="text:index-entry-text[1]/preceding-sibling::text:index-entry-tab-stop[@style:type!='right' and @style:position]">
        <xsl:call-template name="GetLargestTabStop">
          <xsl:with-param name="tabStops"
            select="text:index-entry-text[1]/preceding-sibling::text:index-entry-tab-stop[@style:type!='right' and @style:position]"
          />
        </xsl:call-template>
      </xsl:if>
    </xsl:variable>
    <xsl:variable name="numberingFormat">
      <xsl:call-template name="GetLevelNumberingFormat">
        <xsl:with-param name="level" select="$level - 1"/>
      </xsl:call-template>
    </xsl:variable>

    <!-- insert tabs if right tab defined or if a tab exist after/before entry-text or if parent style has tabs (to be overriden) -->
    <xsl:if
      test="($leftTabStop != '' and $numberingFormat != '' )
      or text:index-entry-text[1]/following-sibling::text:index-entry-tab-stop[@style:type='right' or @style:position]
      or document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$levelStyleName]//style:tab-stop">
      <w:tabs>
        <!-- clear all parent tabs -->
        <xsl:call-template name="ClearParentStyleTabs">
          <xsl:with-param name="parentstyleName" select="$levelStyleName"/>
        </xsl:call-template>

        <!-- declare 1 tab before text -->
        <xsl:if test="$leftTabStop != '' and $numberingFormat != '' ">
          <w:tab w:pos="{$leftTabStop}">
            <xsl:attribute name="w:val">
              <xsl:variable name="styleType">
                <xsl:value-of
                  select="text:index-entry-text[1]/preceding-sibling::text:index-entry-tab-stop[@style:type!='right' and @style:position]/@style:type"
                />
              </xsl:variable>
              <xsl:choose>
                <xsl:when test="$styleType">
                  <xsl:value-of select="$styleType"/>
                </xsl:when>
                <xsl:otherwise>left</xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name="w:leader">
              <xsl:call-template name="ComputeTabStopLeader">
                <xsl:with-param name="tabStop"
                  select="text:index-entry-text[1]/preceding-sibling::text:index-entry-tab-stop[@style:type!='right' and @style:position][1]"
                />
              </xsl:call-template>
            </xsl:attribute>
          </w:tab>
        </xsl:if>

        <!-- declare tabs after text -->
        <xsl:if
          test="($leftTabStop != '' and $numberingFormat != '' ) or text:index-entry-text[1]/following-sibling::text:index-entry-tab-stop[@style:type='right' or @style:position]">
          <!-- do not write tabs after text except right tab-stop -->
          <xsl:for-each
            select="text:index-entry-text[1]/following-sibling::text:index-entry-tab-stop[@style:type = 'right'][1]">
            <xsl:call-template name="tabStop"/>
          </xsl:for-each>
        </xsl:if>
      </w:tabs>
    </xsl:if>

    <!-- tabs before text are retained as indent if no numbering is defined -->
    <xsl:if
      test="text:index-entry-text[1]/preceding-sibling::text:index-entry-tab-stop[@style:type!='right' and @style:position]">
      <xsl:if test="$numberingFormat = '' ">
        <w:ind>
          <xsl:attribute name="w:left">
            <xsl:value-of select="$leftTabStop"/>
          </xsl:attribute>
        </w:ind>
      </xsl:if>
    </xsl:if>
  </xsl:template>

  <!-- transform a tab stop position into indent -->
  <xsl:template name="GetLargestTabStop">
    <xsl:param name="tabStops"/>
    <xsl:param name="result" select="0"/>
    <!-- get value of first tab-stop -->
    <xsl:variable name="toCompare">
      <xsl:choose>
        <xsl:when test="$tabStops[1]/@style:position != '' ">
          <xsl:call-template name="twips-measure">
            <xsl:with-param name="length" select="$tabStops[1]/@style:position"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <!-- add to other tab-stops -->
    <xsl:choose>
      <xsl:when test="count($tabStops) &gt; 1">
        <xsl:call-template name="GetLargestTabStop">
          <xsl:with-param name="tabStops" select="$tabStops[position() &gt; 1]"/>
          <xsl:with-param name="result">
            <xsl:choose>
              <xsl:when test="$result &gt; $toCompare">
                <xsl:value-of select="$result"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$toCompare"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$result &gt; $toCompare">
            <xsl:value-of select="$result"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$toCompare"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <!-- warn loss of index properties -->
  <xsl:template
    match="text:table-of-content|text:illustration-index|text:table-index|text:object-index|text:user-index|text:alphabetical-index|text:bibliography">
    <xsl:variable name="indexName">
      <xsl:choose>
        <xsl:when test="contains(name(), '-index')">
          <xsl:value-of select="substring-after(substring-before(name(), '-index'), 'text:')"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="substring-after(name(), 'text:')"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:if test="*/@text:index-scope = 'chapter' ">
      <xsl:message terminate="no">
        <xsl:text>translation.odf2oox.indexChapterScope%</xsl:text>
        <xsl:value-of select="$indexName"/>
      </xsl:message>
    </xsl:if>
    <xsl:if test="*/@text:relative-tab-stop-position = 'false' ">
      <xsl:message terminate="no">
        <xsl:text>translation.odf2oox.indexIndentProperty%</xsl:text>
        <xsl:value-of select="$indexName"/>
      </xsl:message>
    </xsl:if>
    <xsl:if test="*/@text:sort-algorithm">
      <xsl:message terminate="no">
        <xsl:text>translation.odf2oox.indexSortAlgorithm%</xsl:text>
        <xsl:value-of select="$indexName"/>
      </xsl:message>
    </xsl:if>
    <!-- report loss of toc protection -->
    <xsl:if test="@text:protected = 'true' ">
      <xsl:message terminate="no">
        <xsl:text>translation.odf2oox.indexProtection%</xsl:text>
        <xsl:value-of select="$indexName"/>
      </xsl:message>
    </xsl:if>
    <xsl:apply-templates/>
  </xsl:template>

  <!-- loss of concordance file -->
  <xsl:template match="text:alphabetical-index-auto-mark-file">
    <xsl:message terminate="no">translation.odf2oox.alphabeticalIndexConcordanceFile</xsl:message>
  </xsl:template>
  
  <xsl:template name="InsertIndexTabs">
    <xsl:variable name="styleName">
      <xsl:value-of select="@text:style-name"/>
    </xsl:variable>
    <xsl:if test="ancestor::office:document-content/office:automatic-styles/style:style[@style:name = $styleName]/style:paragraph-properties/style:tab-stops">
    <w:tabs>
      <xsl:variable name="tabInd">
        <xsl:call-template name="twips-measure">
          <xsl:with-param name="length">
            <xsl:value-of select="ancestor::office:document-content/office:automatic-styles/style:style[@style:name = $styleName]/style:paragraph-properties/@fo:margin-left"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
        <xsl:for-each select="ancestor::office:document-content/office:automatic-styles/style:style[@style:name = $styleName]/style:paragraph-properties/style:tab-stops/style:tab-stop">
        <w:tab>
          <xsl:attribute name="w:val">
            <xsl:choose>
              <xsl:when test="@style:type = 'left' ">left</xsl:when>
              <xsl:when test="@style:type = 'right' ">right</xsl:when>
              <xsl:when test="@style:type = 'center' ">center</xsl:when>
              <xsl:otherwise>left</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:attribute name="w:leader">
            <xsl:call-template name="ComputeTabStopLeader"/>
          </xsl:attribute>
          <xsl:variable name="pos">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length">
                <xsl:value-of select="@style:position"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:attribute name="w:pos">
                <xsl:value-of select="$pos+$tabInd"/>
          </xsl:attribute>
        </w:tab>
      </xsl:for-each>
    </w:tabs>
    </xsl:if>
  </xsl:template>

</xsl:stylesheet>

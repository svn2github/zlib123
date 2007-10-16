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
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  exclude-result-prefixes="office style fo text draw number svg dc">

  <xsl:variable name="asianLayoutId">1</xsl:variable>
  <!-- default language of the document -->
  <xsl:variable name="default-language"
    select="document('meta.xml')/office:document-meta/office:meta/dc:language"/>

  <!-- keys definition -->
  <xsl:key name="styles" match="style:style" use="@style:name"/>


  <xsl:template name="styles">
    <w:styles>
      <w:docDefaults>
        <!-- Default text properties -->
        <xsl:variable name="paragraphDefaultStyle"
          select="document('styles.xml')/office:document-styles/office:styles/style:default-style[@style:family='paragraph']"/>
        <w:rPrDefault>
          <w:rPr>
            <xsl:apply-templates select="$paragraphDefaultStyle/style:text-properties" mode="rPr"/>
          </w:rPr>
        </w:rPrDefault>
        <!-- Default paragraph properties -->
        <w:pPrDefault>
          <w:pPr>
            <xsl:apply-templates select="$paragraphDefaultStyle/style:paragraph-properties"
              mode="pPr"/>
          </w:pPr>
        </w:pPrDefault>
      </w:docDefaults>
      <xsl:apply-templates select="document('styles.xml')/office:document-styles/office:styles"
        mode="styles"/>
      <xsl:apply-templates
        select="document('content.xml')/office:document-content/office:automatic-styles"
        mode="styles"/>
      <xsl:apply-templates
        select="document('styles.xml')/office:document-styles/office:automatic-styles" mode="styles"/>
      <!-- styles to ensure that hyperlinks will be matched and followed. -->
      <xsl:call-template name="InsertHyperlinkStyles"/>
      <!-- styles to ensure that indexes keep the same style when updated -->
      <xsl:call-template name="InsertIndexStyles"/>
      <!-- warn if more than one master page style -->
      <xsl:if
        test="count(document('styles.xml')/office:document-styles/office:master-styles/style:master-page) &gt; 1">
        <xsl:message terminate="no">translation.odf2oox.pageLayout</xsl:message>
      </xsl:if>
    </w:styles>
  </xsl:template>

  <!-- Computes the prefixed style name (for automatic styles) -->
  <xsl:template name="GetPrefixedStyleName">
    <xsl:param name="styleName"/>
    <xsl:choose>

      <xsl:when test="key('automatic-styles',$styleName)">
        <xsl:choose>
          <xsl:when test="ancestor::office:document-styles/office:automatic-styles">
            <xsl:value-of select="concat('STYLE_', $styleName)"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat('CONTENT_', $styleName)"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$styleName='Sender' ">EnvelopeReturn</xsl:when>
          <xsl:when test="$styleName='Addressee' ">EnvelopeAddress</xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$styleName"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <xsl:template match="style:style" mode="styles">
    <xsl:variable name="prefixedStyleName">
      <xsl:call-template name="GetPrefixedStyleName">
        <xsl:with-param name="styleName" select="@style:name"/>
      </xsl:call-template>
    </xsl:variable>
    <w:style>
      <!-- Logical name -->
      <xsl:attribute name="w:styleId">
        <xsl:value-of select="$prefixedStyleName"/>
      </xsl:attribute>

      <!-- Style family -->
      <xsl:choose>
        <xsl:when test="@style:family = 'text' ">
          <xsl:attribute name="w:type">character</xsl:attribute>
        </xsl:when>
        <xsl:when test="@style:family = 'graphic' ">
          <xsl:attribute name="w:type">paragraph</xsl:attribute>
        </xsl:when>
        <xsl:when test="@style:family = 'section' ">
          <xsl:attribute name="w:type">paragraph</xsl:attribute>
        </xsl:when>
        <xsl:when test="@style:family = 'paragraph' ">
          <xsl:attribute name="w:type">paragraph</xsl:attribute>
        </xsl:when>
        <xsl:when test="@style:family = 'table' ">
          <xsl:attribute name="w:type">table</xsl:attribute>
        </xsl:when>
        <xsl:when test="@style:family = 'table-cell' ">
          <xsl:attribute name="w:type">table</xsl:attribute>
        </xsl:when>
        <xsl:when test="@style:family = 'table-row' ">
          <xsl:attribute name="w:type">table</xsl:attribute>
        </xsl:when>
        <xsl:when test="@style:family = 'table-column' ">
          <xsl:attribute name="w:type">table</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="w:type">character</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>

      <!-- Display name-->
      <xsl:choose>
        <xsl:when test="@style:name='Sender' ">
          <w:name w:val="envelope return"/>
        </xsl:when>
        <xsl:when test="@style:name='Addressee' ">
          <w:name w:val="envelope address"/>
        </xsl:when>
        <xsl:when test="@style:display-name">
          <w:name w:val="{@style:display-name}"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:if test="@style:name">
            <w:name w:val="{$prefixedStyleName}"/>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>

      <!-- Parent style name -->
      <xsl:choose>
        <xsl:when test="@style:parent-style-name = 'Addressee' ">
          <w:basedOn w:val="EnvelopeAddress"/>
        </xsl:when>
        <xsl:when test="@style:parent-style-name = 'Sender' ">
          <w:basedOn w:val="EnvelopeReturn"/>
        </xsl:when>
        <xsl:when test="@style:parent-style-name">
          <w:basedOn w:val="{@style:parent-style-name}"/>
        </xsl:when>
        <xsl:otherwise>
          <!-- if a hyperlink has current style and does not mention a particular parent style, set basedOn to 'Hyperlink' (built-in by converter, cf template 'InsertHyperlinkStyles') -->
          <xsl:if test="@style:family = 'text' ">
            <xsl:choose>
              <xsl:when test="ancestor::office:automatic-styles">
                <xsl:if test="key('style-modified-hyperlinks',@style:name)">
                  <w:basedOn w:val="Hyperlink"/>
                </xsl:if>
              </xsl:when>
              <xsl:otherwise>
                <xsl:variable name="styleName" select="@style:name"/>
                <xsl:for-each select="document('content.xml')">
                  <xsl:if test="key('style-modified-hyperlinks',$styleName)">
                    <w:basedOn w:val="Hyperlink"/>
                  </xsl:if>
                </xsl:for-each>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>

      <!-- Next style name -->
      <xsl:if test="@style:next-style-name">
        <w:next w:val="{@style:next-style-name}"/>
      </xsl:if>

      <!-- Hide automatic-styles -->
      <xsl:if test="key('automatic-styles', @style:name)">
        <w:hidden/>
      </xsl:if>

      <!-- QuickFormat style -->
      <w:qFormat/>

      <!-- style's paragraph properties -->
      <w:pPr>
        <xsl:apply-templates mode="pPr"/>
        <xsl:call-template name="InsertOutlineLevel"/>
        <xsl:if test="style:paragraph-properties and contains(@style:name,'NumPar')">
          <!-- outline level for numbering -->
          <xsl:choose>
            <xsl:when test="@style:default-outline-level">
              <w:outlineLvl w:val="{number(@style:default-outline-level) - 1}"/>
            </xsl:when>
            <xsl:otherwise>
              <w:outlineLvl w:val="9"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
      </w:pPr>

      <!-- style's text properties -->
      <w:rPr>
        <xsl:choose>
          <xsl:when test="style:text-properties">
            <xsl:apply-templates mode="rPr"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:if test="$default-language">
              <w:lang w:val="{$default-language}"/>
            </xsl:if>
          </xsl:otherwise>
        </xsl:choose>
      </w:rPr>

    </w:style>
  </xsl:template>



  <!-- Paragraph properties -->
  <xsl:template match="style:paragraph-properties" mode="pPr">
    <!-- report lost attributes -->
    <xsl:if test="@fo:text-align-last">
      <xsl:message terminate="no">translation.odf2oox.alignmentLastLine</xsl:message>
    </xsl:if>
    <xsl:if test="@style:background-image">
      <xsl:message terminate="no">translation.odf2oox.paragraphBgImage</xsl:message>
    </xsl:if>

    <!-- keep with next -->
    <xsl:if test="@fo:keep-with-next='always'">
      <w:keepNext/>
    </xsl:if>

    <!-- keep together -->
    <xsl:if test="@fo:keep-together='always'">
      <w:keepLines/>
    </xsl:if>

    <!-- page break before -->
    <xsl:choose>
      <xsl:when test="@fo:break-before='page'">
        <w:pageBreakBefore/>
      </xsl:when>
      <xsl:when test="@fo:break-before='auto'">
        <xsl:message terminate="no">translation.odf2oox.automaticPageBreak</xsl:message>
        <w:pageBreakBefore w:val="off"/>
      </xsl:when>
      <xsl:otherwise/>
    </xsl:choose>

    <!-- widow/orphan control -->
    <xsl:choose>
      <xsl:when test="@fo:widows != '0' or @fo:orphans != '0'">
        <xsl:if test="@fo:widows &gt; 2 or @fo:orphans &gt; 2">
          <xsl:message terminate="no">translation.odf2oox.numberWidowsOrphans</xsl:message>
        </xsl:if>
        <w:widowControl w:val="on"/>
      </xsl:when>
      <xsl:otherwise>
        <w:widowControl w:val="off"/>
      </xsl:otherwise>
    </xsl:choose>

    <!-- numbering properties -->
    <xsl:call-template name="InsertOutlineNumPr"/>

    <!-- line numbers -->
    <xsl:if
      test="not(document('styles.xml')/office:document-styles/office:styles/text:linenumbering-configuration/@text:number-lines='false' )">
      <xsl:choose>
        <xsl:when test="@text:number-lines='false' ">
          <w:suppressLineNumbers w:val="true"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:if test="@text:number-lines='true' ">
            <w:suppressLineNumbers w:val="false"/>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <!-- border color + padding + shadow -->
    <xsl:choose>
      <xsl:when test="@fo:border and @fo:border != 'none' ">
        <w:pBdr>
          <xsl:call-template name="InsertBorders">
            <xsl:with-param name="allSides">true</xsl:with-param>
          </xsl:call-template>
        </w:pBdr>
      </xsl:when>
      <xsl:when test="@fo:border = 'none' ">
        <w:pBdr>
          <xsl:call-template name="InsertEmptyBorders"/>
        </w:pBdr>
      </xsl:when>
      <xsl:when test="@fo:border-top or @fo:border-left or @fo:border-bottom or @fo:border-right">
        <w:pBdr>
          <xsl:call-template name="InsertBorders">
            <xsl:with-param name="allSides">false</xsl:with-param>
          </xsl:call-template>
        </w:pBdr>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="@style:shadow">
          <w:pBdr>
            <xsl:call-template name="InsertEmptyBorders"/>
          </w:pBdr>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>

    <!-- Background color -->
    <xsl:if test="@fo:background-color and (@fo:background-color != 'transparent')">
      <w:shd>
        <xsl:attribute name="w:val">clear</xsl:attribute>
        <xsl:attribute name="w:color">auto</xsl:attribute>
        <xsl:attribute name="w:fill">
          <xsl:value-of
            select="substring(@fo:background-color, 2, string-length(@fo:background-color) -1)"/>
        </xsl:attribute>
      </w:shd>
    </xsl:if>

    <!-- Tabs -->
    <xsl:if test="style:tab-stops/style:tab-stop">
      <w:tabs>
        <xsl:call-template name="ComputeParagraphTabs"/>
      </w:tabs>
    </xsl:if>

    <!-- Override hyphenation -->
    <xsl:choose>
      <xsl:when test="parent::node()/style:text-properties/@fo:hyphenate='true'">
        <w:suppressAutoHyphens w:val="false"/>
      </xsl:when>
      <xsl:otherwise>
        <w:suppressAutoHyphens w:val="true"/>
      </xsl:otherwise>
    </xsl:choose>

    <!-- Auto space -->
    <xsl:if test="@style:text-autospace">
      <w:autoSpaceDN>
        <xsl:choose>
          <xsl:when test="@style:text-autospace='none'">
            <xsl:attribute name="w:val">off</xsl:attribute>
          </xsl:when>
          <xsl:when test="@style:text-autospace='ideograph-alpha'">
            <xsl:attribute name="w:val">on</xsl:attribute>
          </xsl:when>
        </xsl:choose>
      </w:autoSpaceDN>
    </xsl:if>

    <xsl:if test="contains(@style:writing-mode, 'rl')">
      <w:bidi/>
    </xsl:if>
    
    <!-- Spacing -->
    <xsl:if
      test="@style:line-height-at-least or @fo:line-height or @fo:margin-bottom or @fo:margin-top or @style:line-spacing">
      <xsl:call-template name="ComputeParagraphSpacing"/>
    </xsl:if>

    <!-- indent -->
    <xsl:if
      test="@fo:margin-left or @fo:margin-right or @fo:text-indent or @text:space-before
      or @fo:padding or @fo:padding-left or @fo:padding-right
      or @fo:border or @fo:border-left or @fo:border-right">
      <xsl:call-template name="ComputeParagraphIndent"/>
    </xsl:if>

    <!-- Paragraph alignment -->
    <xsl:if test="@fo:text-align">
      <xsl:call-template name="InsertTextAlignment"/>
    </xsl:if>

    <!-- Vertical alignment of all text on each line of the paragraph -->
    <xsl:if test="@style:vertical-align">
      <w:textAlignment>
        <xsl:choose>
          <xsl:when test="@style:vertical-align='bottom'">
            <xsl:attribute name="w:val">bottom</xsl:attribute>
          </xsl:when>
          <xsl:when test="@style:vertical-align='top'">
            <xsl:attribute name="w:val">top</xsl:attribute>
          </xsl:when>
          <xsl:when test="@style:vertical-align='middle'">
            <xsl:attribute name="w:val">center</xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="w:val">auto</xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </w:textAlignment>
    </xsl:if>

    <!-- outline level for numbering -->
    <xsl:choose>
    <xsl:when test="parent::style:style/@style:default-outline-level">
      <w:outlineLvl w:val="{number(parent::style:style/@style:default-outline-level) - 1}"/>
    </xsl:when>
      <xsl:otherwise>
        <!-- avoid putting TOC heading into TOC (bug #1760863) -->
        <xsl:if test="parent::style:style[@style:name='Contents_20_Heading']">
          <w:outlineLvl w:val="10"/>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>


  <!--  Paragraph properties from attributes -->
  <xsl:template name="InsertOutlineLevel">
    <xsl:if test="not(style:paragraph-properties)">
      <xsl:if test="not(style:text-properties)">
        <xsl:call-template name="InsertOutlineNumPr"/>
      </xsl:if>
      <!-- outline level for numbering -->
      <xsl:choose>
        <xsl:when test="@style:default-outline-level">
        <w:outlineLvl w:val="{number(@style:default-outline-level) - 1}"/>
        </xsl:when>
        <xsl:otherwise>
          <w:outlineLvl w:val="9"/>
        </xsl:otherwise>
        </xsl:choose>
    </xsl:if>
  </xsl:template>

  <!-- insert a reference to list 1, or override this reference -->
  <xsl:template name="InsertOutlineNumPr">
    <!-- context can be style:style, style:para-props or style:text-props -->
    <xsl:choose>
      <xsl:when test="ancestor-or-self::style:style/@style:default-outline-level">
        <w:numPr>
          <w:numId w:val="1"/>
        </w:numPr>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if
          test="ancestor::office:styles and key('styles', ancestor-or-self::style:style/@style:parent-style-name)/@style:default-outline-level">
          <w:numPr>
            <w:numId w:val="0"/>
          </w:numPr>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <!-- Paragraph properties from text-properties -->
  <xsl:template match="style:text-properties" mode="pPr">
    <xsl:if test="not(parent::node()/style:paragraph-properties)">
      <xsl:call-template name="InsertOutlineNumPr"/>
      <xsl:choose>
        <xsl:when test="@fo:hyphenate='true'">
          <w:suppressAutoHyphens w:val="false"/>
        </xsl:when>
        <xsl:otherwise>
          <w:suppressAutoHyphens w:val="true"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
  </xsl:template>

  <!-- Text alignment property -->
  <!-- Also depends on the writing direction (righ to left, etc) -->
  <xsl:template name="InsertTextAlignment">
    <w:jc>
      <xsl:choose>
        <xsl:when test="@fo:text-align='center'">
          <xsl:attribute name="w:val">center</xsl:attribute>
        </xsl:when>
        <xsl:when test="@fo:text-align='start'">
          <xsl:if test="contains(@style:writing-mode, 'rl')">
            <xsl:attribute name="w:val">right</xsl:attribute>
          </xsl:if>
          <xsl:if test="not(contains(@style:writing-mode, 'rl'))">
            <xsl:attribute name="w:val">left</xsl:attribute>
          </xsl:if>
        </xsl:when>
        <xsl:when test="@fo:text-align='end'">
          <xsl:if test="contains(@style:writing-mode, 'rl')">
            <xsl:attribute name="w:val">left</xsl:attribute>
          </xsl:if>
          <xsl:if test="not(contains(@style:writing-mode, 'rl'))">
            <xsl:attribute name="w:val">right</xsl:attribute>
          </xsl:if>
        </xsl:when>
        <xsl:when test="@fo:text-align='justify'">
          <xsl:attribute name="w:val">both</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:if test="contains(@style:writing-mode, 'rl')">
            <xsl:attribute name="w:val">right</xsl:attribute>
          </xsl:if>
          <xsl:if test="not(contains(@style:writing-mode, 'rl'))">
            <xsl:attribute name="w:val">left</xsl:attribute>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>
    </w:jc>
  </xsl:template>

  <!-- Paragraph tabs property -->
  <xsl:template name="ComputeParagraphTabs">
    <xsl:for-each select="style:tab-stops/style:tab-stop">
      <xsl:call-template name="tabStop"/>
    </xsl:for-each>

    <!-- clear parent tabs if needed. -->
    <xsl:variable name="styleName">
      <xsl:value-of select="parent::node()/@style:name"/>
    </xsl:variable>
    <xsl:variable name="parentstyleName">
      <xsl:value-of select="parent::node()/@style:parent-style-name"/>
    </xsl:variable>
    <!-- remember original context -->
    <xsl:variable name="styleContext">
      <xsl:choose>
        <xsl:when test="ancestor::office:document-content">content</xsl:when>
        <xsl:when test="ancestor::office:document-styles">
          <xsl:choose>
            <xsl:when test="ancestor::office:automatic-styles">automaticStyles</xsl:when>
            <xsl:when test="ancestor::office:styles">styles</xsl:when>
          </xsl:choose>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>

   <!-- <xsl:call-template name="ClearParentStyleTabs">
      <xsl:with-param name="parentstyleName" select="$parentstyleName"/>
      <xsl:with-param name="styleContext" select="$styleContext"/>
      <xsl:with-param name="styleName" select="$styleName"/>
    </xsl:call-template>-->
  </xsl:template>

  <!-- clear tabs of parent style -->
  <xsl:template name="ClearParentStyleTabs">
    <xsl:param name="parentstyleName"/>
    <xsl:param name="styleContext"/>
    <xsl:param name="styleName"/>
    <!-- change context to see if it is necessary to clear parent tabs -->
    <xsl:for-each select="document('styles.xml')">
      <xsl:for-each select="key('styles', $parentstyleName)[1]//style:tab-stop">
        <xsl:variable name="parentPosition" select="@style:position"/>
        <xsl:variable name="clear">
          <!-- go back to original context, and check if parent tab is also present in unheriting style -->
          <xsl:choose>
            <xsl:when test="$styleContext='content' ">
              <xsl:for-each select="document('content.xml')">
                <xsl:if
                  test="not($parentPosition = key('automatic-styles', $styleName)[1]//style:tab-stop/@style:position)">
                  <xsl:value-of select="'true'"/>
                </xsl:if>
              </xsl:for-each>
            </xsl:when>
            <xsl:when test="$styleContext='automaticStyles' ">
              <xsl:if
                test="not($parentPosition = key('automatic-styles', $styleName)[1]//style:tab-stop/@style:position)">
                <xsl:value-of select="'true'"/>
              </xsl:if>
            </xsl:when>
            <xsl:when test="$styleContext='styles' ">
              <xsl:if
                test="not($parentPosition = key('styles', $styleName)[1]//style:tab-stop/@style:position)">
                <xsl:value-of select="'true'"/>
              </xsl:if>
            </xsl:when>
            <xsl:otherwise>true</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <!-- clear the tab, from the parent style context. -->
        <xsl:if test="$clear = 'true' ">
          <xsl:call-template name="tabStop">
            <xsl:with-param name="styleType">clear</xsl:with-param>
          </xsl:call-template>
        </xsl:if>
      </xsl:for-each>
    </xsl:for-each>
  </xsl:template>

  <!-- Paragraph spacing property -->
  <xsl:template name="ComputeParagraphSpacing">
    <w:spacing>
      <!-- line spacing -->
      <xsl:choose>
        <xsl:when test="@style:line-height-at-least">
          <xsl:attribute name="w:lineRule">atLeast</xsl:attribute>
          <xsl:attribute name="w:line">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length" select="@style:line-height-at-least"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="@style:line-spacing">
          <xsl:variable name="spacing">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length" select="@style:line-spacing"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:attribute name="w:lineRule">auto</xsl:attribute>
          <xsl:attribute name="w:line">
            <xsl:value-of select="number(240 + $spacing)"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="contains(@fo:line-height, '%')">
          <xsl:attribute name="w:lineRule">auto</xsl:attribute>
          <xsl:attribute name="w:line">
            <!-- w:line expressed in 240ths of a line height when w:lineRule='auto' -->
            <xsl:value-of
              select="round(number(substring-before(@fo:line-height, '%')) * 240 div 100)"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:if test="@fo:line-height">
            <xsl:attribute name="w:lineRule">exact</xsl:attribute>
            <xsl:attribute name="w:line">
              <xsl:choose>
                <xsl:when test="@fo:line-height = 'normal' ">240</xsl:when>
                <xsl:otherwise>
                  <xsl:call-template name="twips-measure">
                    <xsl:with-param name="length" select="@fo:line-height"/>
                  </xsl:call-template>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>
      <!-- top / bottom spacing -->
      <xsl:if test="@fo:margin-bottom">
        <xsl:attribute name="w:after">
          <xsl:call-template name="twips-measure">
            <xsl:with-param name="length" select="@fo:margin-bottom"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@fo:margin-top">
        <xsl:attribute name="w:before">
          <xsl:call-template name="twips-measure">
            <xsl:with-param name="length" select="@fo:margin-top"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
    </w:spacing>
  </xsl:template>



  <!-- Paragraph indent property -->
  <xsl:template name="ComputeParagraphIndent">
    <xsl:variable name="defaultOutlineLevel">
      <xsl:call-template name="GetDefaultOutlineLevel">
        <xsl:with-param name="styleName" select="parent::style:style/@style:name"/>
      </xsl:call-template>
    </xsl:variable>
    
    <!-- left indent -->
    <xsl:variable name="leftInd">
      <xsl:call-template name="ComputeAdditionalIndent">
        <xsl:with-param name="side" select="'left'"/>
        <xsl:with-param name="style" select="parent::style:style|parent::style:default-style" />
      </xsl:call-template>
    </xsl:variable>
    <!-- right indent -->
    <xsl:variable name="rightInd">
      <xsl:call-template name="ComputeAdditionalIndent">
        <xsl:with-param name="side" select="'right'"/>
        <xsl:with-param name="style" select="parent::style:style|parent::style:default-style" />
      </xsl:call-template>
    </xsl:variable>
    <!-- first line indent -->
    <xsl:variable name="firstLineIndent">
      <xsl:call-template name="GetFirstLineIndent">
        <xsl:with-param name="style" select="parent::style:style|parent::style:default-style" />
      </xsl:call-template>
    </xsl:variable>
    
    <xsl:choose>
      <xsl:when test="number($defaultOutlineLevel) or $defaultOutlineLevel = 0">
        <xsl:call-template name="OverrideNumberingProperty">
          <xsl:with-param name="level" select="$defaultOutlineLevel"/>
          <xsl:with-param name="property">indent</xsl:with-param>
        </xsl:call-template>
      </xsl:when>
     
      <xsl:otherwise>
        
        <w:ind>
         
          <xsl:attribute name="w:left">
            <xsl:value-of select="$leftInd"/>
          </xsl:attribute>
          
          <xsl:attribute name="w:right">
            <xsl:value-of select="$rightInd"/>
          </xsl:attribute>
          
          <xsl:choose>
            <xsl:when test="$firstLineIndent != 0">
              <xsl:choose>
                <xsl:when test="$firstLineIndent &lt; 0">
                  <xsl:attribute name="w:hanging">
                    <xsl:value-of select="-$firstLineIndent"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:attribute name="w:firstLine">
                    <xsl:value-of select="$firstLineIndent"/>
                  </xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
              <!-- default value is hanging=0 if nothing is specified -->
              <xsl:attribute name="w:hanging">0</xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </w:ind>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>



  <!-- Text properties -->
  <xsl:template match="style:text-properties" mode="rPr">

    <!-- report lost attributes -->
    <xsl:if test="@style:text-blinking">
      <xsl:message terminate="no">translation.odf2oox.blinkingText</xsl:message>
    </xsl:if>
    <xsl:if test="@style:text-rotation-angle or @style:text-rotation-scale">
      <xsl:message terminate="no">translation.odf2oox.rotatedText</xsl:message>
    </xsl:if>

    <!-- fonts -->
    <xsl:if
      test="@style:font-name or @style:font-name-complex or @style:font-name-asian
      or @fo:font-family or @style:font-family-asian or @style:font-family-complex">
      <w:rFonts>
        <xsl:variable name="fontName">
          <xsl:call-template name="ComputeFontName">
            <xsl:with-param name="fontName">
              <xsl:choose>
                <xsl:when test="@style:font-name">
                  <xsl:value-of select="@style:font-name"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:if test="@fo:font-family">
                    <xsl:value-of select="@fo:font-family"/>
                  </xsl:if>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:if test="$fontName != '' ">
          <xsl:attribute name="w:ascii">
            <xsl:value-of select="$fontName"/>
          </xsl:attribute>
          <xsl:attribute name="w:hAnsi">
            <xsl:value-of select="$fontName"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:choose>
          <xsl:when test="@style:font-name-complex">
            <xsl:attribute name="w:cs">
              <xsl:call-template name="ComputeFontName">
                <xsl:with-param name="fontName" select="@style:font-name-complex"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:if test="@fo:font-family-complex">
              <xsl:attribute name="w:cs">
                <xsl:call-template name="ComputeFontName">
                  <xsl:with-param name="fontName" select="@style:font-family-complex"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
          </xsl:otherwise>
        </xsl:choose>
        <xsl:choose>
          <xsl:when test="@style:font-name-asian">
            <xsl:attribute name="w:eastAsia">
              <xsl:call-template name="ComputeFontName">
                <xsl:with-param name="fontName" select="@style:font-name-asian"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:if test="@fo:font-family-asian">
              <xsl:attribute name="w:eastAsia">
                <xsl:call-template name="ComputeFontName">
                  <xsl:with-param name="fontName" select="@style:font-family-asian"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
          </xsl:otherwise>
        </xsl:choose>
      </w:rFonts>
    </xsl:if>

    <xsl:if test="@fo:font-weight">
      <xsl:choose>
        <xsl:when test="@fo:font-weight != 'normal'">
          <xsl:choose>
            <xsl:when test="@fo:font-weight != 'bold' and number(@fo:font-weight)">
              <xsl:message terminate="no">translation.odf2oox.fontWeight</xsl:message>
            </xsl:when>
            <xsl:otherwise>
              <w:b w:val="on"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <w:b w:val="off"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@fo:font-weight-complex">
      <xsl:choose>
        <xsl:when test="@style:font-weight-complex != 'normal'">
          <w:bCs w:val="on"/>
        </xsl:when>
        <xsl:otherwise>
          <w:bCs w:val="off"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@fo:font-style">
      <xsl:choose>
        <xsl:when test="@fo:font-style = 'italic' or @fo:font-style = 'oblique'">
          <w:i w:val="on"/>
        </xsl:when>
        <xsl:otherwise>
          <w:i w:val="off"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@fo:font-style-complex">
      <xsl:choose>
        <xsl:when test="@fo:font-style-complex = 'italic' or @fo:font-style-complex = 'oblique'">
          <w:iCs w:val="on"/>
        </xsl:when>
        <xsl:otherwise>
          <w:iCs w:val="off"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@fo:text-transform or @fo:font-variant">
      <xsl:choose>
        <xsl:when test="@fo:text-transform = 'uppercase'">
          <w:caps w:val="on"/>
        </xsl:when>
        <xsl:when test="@fo:text-transform = 'capitalize'">
          <!-- no equivalent of capitalize in OOX spec -->
          <xsl:message terminate="no">translation.odf2oox.capitalizedText</xsl:message>
          <w:caps w:val="off"/>
        </xsl:when>
        <xsl:when test="@fo:text-transform = 'lowercase'">
          <xsl:message terminate="no">translation.odf2oox.lowercaseText</xsl:message>
          <w:caps w:val="off"/>
        </xsl:when>
        <xsl:when test="@fo:text-transform = 'none' or @fo:font-variant = 'small-caps'">
          <w:caps w:val="off"/>
        </xsl:when>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@fo:font-variant">
      <xsl:choose>
        <xsl:when test="@fo:font-variant = 'small-caps'">
          <w:smallCaps w:val="on"/>
        </xsl:when>
        <xsl:when test="@fo:font-weight = 'normal'">
          <w:smallCaps w:val="off"/>
        </xsl:when>
      </xsl:choose>
    </xsl:if>

    <xsl:if
      test="@style:text-line-through-style != 'none' or @style:text-line-through-type != 'none' ">
      <xsl:choose>
        <xsl:when test="@style:text-line-through-type = 'double'">
          <w:dstrike w:val="on"/>
        </xsl:when>
        <xsl:when test="@style:text-line-through-type = 'none'">
          <w:strike w:val="off"/>
        </xsl:when>
        <xsl:otherwise>
          <w:strike w:val="on"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@style:text-line-through-width != 'none'">
      <!-- TODO : localize this -->
      <xsl:message terminate="no">translation.odf2oox.lineThroughWidth</xsl:message>
    </xsl:if>

    <xsl:if test="@style:text-line-through-color != 'none'">
      <!-- TODO : localize this -->
      <xsl:message terminate="no">translation.odf2oox.lineThroughColor</xsl:message>
    </xsl:if>

    <xsl:if test="@style:text-outline">
      <xsl:choose>
        <xsl:when test="@style:text-outline = 'true'">
          <w:outline w:val="on"/>
        </xsl:when>
        <xsl:when test="@style:text-outline = 'false'">
          <w:outline w:val="off"/>
        </xsl:when>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@fo:text-shadow">
      <xsl:choose>
        <xsl:when test="@fo:text-shadow = 'none'">
          <w:shadow w:val="off"/>
        </xsl:when>
        <xsl:otherwise>
          <w:shadow w:val="on"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@style:font-relief">
      <xsl:choose>
        <xsl:when test="@style:font-relief = 'embossed'">
          <w:emboss w:val="on"/>
        </xsl:when>
        <xsl:when test="@style:font-relief = 'engraved'">
          <w:imprint w:val="on"/>
        </xsl:when>
        <xsl:otherwise>
          <w:emboss w:val="off"/>
          <w:imprint w:val="off"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@text:display">
      <xsl:choose>
        <xsl:when test="@text:display = 'true'">
          <w:vanish w:val="on"/>
        </xsl:when>
        <xsl:when test="@text:display = 'none'">
          <w:vanish w:val="off"/>
        </xsl:when>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="@fo:color">
      <w:color>
        <xsl:attribute name="w:val">
          <xsl:value-of select="substring(@fo:color, 2, string-length(@fo:color) -1)"/>
        </xsl:attribute>
      </w:color>
    </xsl:if>

    <xsl:if test="@fo:letter-spacing">
      <w:spacing>
        <xsl:attribute name="w:val">
          <xsl:choose>
            <xsl:when test="@fo:letter-spacing='normal'">
              <xsl:value-of select="0"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="twips-measure">
                <xsl:with-param name="length" select="@fo:letter-spacing"/>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </w:spacing>
    </xsl:if>

    <xsl:if test="@style:text-scale">
      <w:w>
        <xsl:attribute name="w:val">
          <xsl:variable name="scale"
            select="substring(@style:text-scale, 1, string-length(@style:text-scale)-1)"/>
          <xsl:choose>
            <xsl:when test="number($scale) &lt; 600">
              <xsl:value-of select="$scale"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:message terminate="no">translation.odf2oox.textScale600pct</xsl:message>
              <xsl:text>600</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </w:w>
    </xsl:if>

    <xsl:if test="@style:letter-kerning = 'true' ">
      <w:kern w:val="16"/>
    </xsl:if>

    <!-- position : raise text with precise measurement -->
    <xsl:if test="@style:text-position">
      <xsl:if
        test="not(substring(@style:text-position, 1, 3) = 'sub') and not(substring(@style:text-position, 1, 5) = 'super') ">
        <!-- first percentage value specifies distance to baseline -->
        <xsl:variable name="textPosition">
          <xsl:value-of select="substring-before(@style:text-position, '%')"/>
        </xsl:variable>
        <xsl:if test="$textPosition != '' and $textPosition != 0">
          <xsl:variable name="fontHeight">
            <xsl:call-template name="computeSize"/>
          </xsl:variable>
          <xsl:if test="number($fontHeight)">
            <w:position>
              <xsl:attribute name="w:val">
                <xsl:choose>
                  <xsl:when test="contains($textPosition, '-')">
                    <xsl:value-of
                      select="concat('-', round(number(substring-after($textPosition, '-')) div 100 * $fontHeight))"
                    />
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="round(number($textPosition) div 100 * $fontHeight)"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </w:position>
          </xsl:if>
        </xsl:if>
      </xsl:if>
    </xsl:if>

    <!-- font size -->
    <xsl:choose>
      <xsl:when test="@style:text-position">
        <xsl:variable name="sz">
          <xsl:call-template name="ComputePositionFontSize"/>
        </xsl:variable>
        <xsl:if test="number($sz)">
          <w:sz w:val="{$sz}"/>
        </xsl:if>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="@fo:font-size">
          <xsl:variable name="sz">
            <xsl:call-template name="computeSize"/>
          </xsl:variable>
          <xsl:if test="number($sz)">
            <w:sz w:val="{$sz}"/>
          </xsl:if>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>

    <xsl:if test="@style:font-size-complex">
      <xsl:variable name="sz">
        <xsl:call-template name="computeSize">
          <xsl:with-param name="fontStyle">style:font-size-complex</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:if test="number($sz)">
        <w:szCs w:val="{$sz}"/>
      </xsl:if>
    </xsl:if>

    <xsl:if test="@fo:background-color">
      <w:highlight>
        <xsl:attribute name="w:val">
          <!-- Value must be a string (no RGB value authorized). -->
          <xsl:choose>
            <xsl:when test="@fo:background-color = 'transparent' ">none</xsl:when>
            <xsl:when test="not(substring(@fo:background-color, 1,1) = '#' )">
              <xsl:value-of select="@fo:background-color"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="StringType-color">
                <xsl:with-param name="RGBcolor" select="@fo:background-color"/>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </w:highlight>
    </xsl:if>

    <xsl:if test="@style:text-underline-style">
      <xsl:choose>
        <xsl:when test="@style:text-underline-style != 'none' ">
          <w:u>
            <xsl:attribute name="w:val">
              <xsl:choose>
                <xsl:when test="@style:text-underline-style = 'dotted'">
                  <xsl:choose>
                    <xsl:when test="@style:text-underline-width = 'thick' ">dottedHeavy</xsl:when>
                    <xsl:when test="@style:text-underline-width = 'bold' ">dottedHeavy</xsl:when>
                    <xsl:otherwise>dotted</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:when test="@style:text-underline-style = 'dash' ">
                  <xsl:choose>
                    <xsl:when test="@style:text-underline-width = 'thick' ">dashedHeavy</xsl:when>
                    <xsl:when test="@style:text-underline-width = 'bold' ">dashedHeavy</xsl:when>
                    <xsl:otherwise>dash</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:when test="@style:text-underline-style = 'long-dash'">
                  <xsl:choose>
                    <xsl:when test="@style:text-underline-width = 'thick' ">dashLongHeavy</xsl:when>
                    <xsl:when test="@style:text-underline-width = 'bold' ">dashLongHeavy</xsl:when>
                    <xsl:otherwise>dashLong</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:when test="@style:text-underline-style = 'dot-dash' ">
                  <xsl:choose>
                    <xsl:when test="@style:text-underline-width = 'thick' ">dashDotHeavy</xsl:when>
                    <xsl:when test="@style:text-underline-width = 'bold' ">dashDotHeavy</xsl:when>
                    <xsl:otherwise>dotDash</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:when test="@style:text-underline-style = 'dot-dot-dash' ">
                  <xsl:choose>
                    <xsl:when test="@style:text-underline-width = 'thick' ">dashDotDotHeavy</xsl:when>
                    <xsl:when test="@style:text-underline-width = 'bold' ">dashDotDotHeavy</xsl:when>
                    <xsl:otherwise>dotDotDash</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:when test="@style:text-underline-style = 'wave' ">
                  <xsl:choose>
                    <xsl:when test="@style:text-underline-type = 'double' ">wavyDouble</xsl:when>
                    <xsl:when test="@style:text-underline-width = 'thick' ">wavyHeavy</xsl:when>
                    <xsl:when test="@style:text-underline-width = 'bold' ">wavyHeavy</xsl:when>
                    <xsl:otherwise>wave</xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:choose>
                    <xsl:when test="@style:text-underline-type = 'double' ">double</xsl:when>
                    <xsl:when test="@style:text-underline-width = 'thick' ">thick</xsl:when>
                    <xsl:when test="@style:text-underline-width = 'bold' ">thick</xsl:when>
                    <xsl:when test="@style:text-underline-mode = 'skip-white-space' ">words</xsl:when>
                    <xsl:otherwise>single</xsl:otherwise>
                  </xsl:choose>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:if test="@style:text-underline-color">
              <xsl:attribute name="w:color">
                <xsl:choose>
                  <xsl:when test="@style:text-underline-color = 'font-color'">auto</xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of
                      select="substring(@style:text-underline-color, 2, string-length(@style:text-underline-color)-1)"
                    />
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </w:u>
        </xsl:when>
        <xsl:otherwise>
          <w:u w:val="none"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <!-- vertAlign : raise text with as super or sub script -->
    <xsl:if test="@style:text-position">
      <xsl:choose>
        <xsl:when test="substring(@style:text-position, 1, 3) = 'sub' ">
          <w:vertAlign w:val="subscript"/>
        </xsl:when>
        <xsl:when test="substring(@style:text-position, 1, 5) = 'super' ">
          <w:vertAlign w:val="superscript"/>
        </xsl:when>
        <xsl:otherwise>
          <!-- handled by position element -->
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    
    <xsl:if test="@style:text-emphasize">
      <w:em>
        <xsl:attribute name="w:val">
          <xsl:choose>
            <xsl:when test="@style:text-emphasize = 'accent above' ">comma</xsl:when>
            <xsl:when test="@style:text-emphasize = 'dot above' ">dot</xsl:when>
            <xsl:when test="@style:text-emphasize = 'circle above' ">circle</xsl:when>
            <xsl:when test="@style:text-emphasize = 'dot below' ">underDot</xsl:when>
            <xsl:otherwise>none</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </w:em>
    </xsl:if>

    <xsl:if
      test="$default-language or (@fo:language and @fo:country) or (@style:language-asian and @style:country-asian) or (@style:language-complex and @style:country-complex)">
      <w:lang>
        <xsl:choose>
          <xsl:when test="(@fo:language and @fo:country)">
            <xsl:attribute name="w:val">
              <xsl:value-of select="@fo:language"/>-<xsl:value-of select="@fo:country"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="$default-language">
            <xsl:attribute name="w:val">
              <xsl:value-of select="$default-language"/>
            </xsl:attribute>
          </xsl:when>
        </xsl:choose>
        <xsl:if test="(@style:language-asian and @style:country-asian)">
          <xsl:attribute name="w:eastAsia">
            <xsl:value-of select="@style:language-asian"/>-<xsl:value-of
              select="@style:country-asian"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="(@style:language-complex and @style:country-complex)">
          <xsl:attribute name="w:bidi">
            <xsl:value-of select="@style:language-complex"/>-<xsl:value-of
              select="@style:country-complex"/>
          </xsl:attribute>
        </xsl:if>
      </w:lang>
    </xsl:if>

    <xsl:if test="@style:text-combine">
      <xsl:choose>
        <xsl:when test="@style:text-combine = 'lines'">
          <w:eastAsianLayout>
            <xsl:attribute name="w:id">
              <xsl:number count="style:style"/>
            </xsl:attribute>
            <xsl:attribute name="w:combine">true</xsl:attribute>
            <xsl:choose>
              <xsl:when
                test="@style:text-combine-start-char = '&lt;' and @style:text-combine-end-char = '&gt;' ">
                <xsl:attribute name="w:combineBrackets">angle</xsl:attribute>
              </xsl:when>
              <xsl:when
                test="@style:text-combine-start-char = '{' and @style:text-combine-end-char = '}'">
                <xsl:attribute name="w:combineBrackets">curly</xsl:attribute>
              </xsl:when>
              <xsl:when
                test="@style:text-combine-start-char = '(' and @style:text-combine-end-char = ')'">
                <xsl:attribute name="w:combineBrackets">round</xsl:attribute>
              </xsl:when>
              <xsl:when
                test="@style:text-combine-start-char = '[' and @style:text-combine-end-char = ']'">
                <xsl:attribute name="w:combineBrackets">square</xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:attribute name="w:combineBrackets">none</xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>
          </w:eastAsianLayout>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
  </xsl:template>



  <!-- footnote and endnote reference and text styles -->
  <xsl:template
    match="text:notes-configuration[@text:note-class='footnote'] | text:notes-configuration[@text:note-class='endnote']"
    mode="styles">
    <w:style w:styleId="{concat(@text:note-class, 'Reference')}" w:type="character">
      <w:name w:val="note reference"/>
      <xsl:if test="@text:citation-body-style-name">
        <w:basedOn w:val="{@text:citation-body-style-name}"/>
      </xsl:if>
      <w:semiHidden/>
      <w:unhideWhenUsed/>
    </w:style>
    <w:style w:styleId="{concat(@text:note-class, 'Text')}" w:type="paragraph">
      <w:name w:val="note text"/>
      <xsl:if test="@text:citation-style-name">
        <w:basedOn w:val="{@text:citation-style-name}"/>
      </xsl:if>
      <!--w:link w:val="FootnoteTextChar"/-->
      <w:semiHidden/>
      <w:unhideWhenUsed/>
    </w:style>
  </xsl:template>



  <!-- Font size computation 
        Context must be document('styles.xml')
    -->
  <xsl:template name="computeSize">
    <xsl:param name="node" select="."/>
    <xsl:param name="fontStyle">fo:font-size</xsl:param>
    <xsl:choose>
      <!-- when there's no unit -->
      <xsl:when test="number($node/@*[name() = $fontStyle])">
        <xsl:value-of select="number($node/@*[name() = $fontStyle]) * 2"/>
      </xsl:when>
      <xsl:when test="contains($node/@*[name() = $fontStyle], 'pt')">
        <xsl:value-of
          select="round(number(substring-before($node/@*[name() = $fontStyle], 'pt')) * 2)"/>
      </xsl:when>
      <xsl:otherwise>
        <!-- look for font size in parent context -->
        <xsl:variable name="parentStyleName" select="$node/../@style:parent-style-name"/>
        <xsl:variable name="value">
          <xsl:choose>
            <xsl:when test="$parentStyleName != '' or count($parentStyleName) &gt; 0">
              <xsl:for-each select="document('styles.xml')">
                <xsl:call-template name="computeSize">
                  <xsl:with-param name="node"
                    select="key('styles', $parentStyleName)/style:text-properties"/>
                  <xsl:with-param name="fontStyle" select="$fontStyle"/>
                </xsl:call-template>
              </xsl:for-each>
            </xsl:when>
            <xsl:otherwise>none</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="number($value)">
            <xsl:choose>
              <xsl:when test="contains($node/@*[name() = $fontStyle], '%')">
                <xsl:value-of
                  select="round(number(substring-before($node/@*[name() = $fontStyle], '%')) div 100 * number($value))"
                />
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$value"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>
            <!-- fetch the default font size for this style family -->
            <xsl:variable name="defaultProps">
              <xsl:variable name="family">
                <xsl:value-of select="$node/../@style:family"/>
              </xsl:variable>
              <xsl:for-each select="document('styles.xml')">
                <xsl:choose>
                  <xsl:when
                    test="office:document-styles/office:styles/style:default-style[@style:family=$family]/style:text-properties/@*[name() = $fontStyle]">
                    <xsl:value-of
                      select="office:document-styles/office:styles/style:default-style[@style:family=$family]/style:text-properties/@*[name() = $fontStyle]"
                    />
                  </xsl:when>
                  <xsl:otherwise>
                    <!-- when there is no other possibility -->
                    <xsl:value-of
                      select="office:document-styles/office:styles/style:default-style[@style:family='paragraph']/style:text-properties/@*[name() = $fontStyle]"
                    />
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:for-each>
            </xsl:variable>
            <xsl:variable name="defaultValue"
              select="number(substring-before($defaultProps, 'pt'))*2"/>
            <xsl:choose>
              <xsl:when test="contains($node/@*[name() = $fontStyle], '%')">
                <xsl:value-of
                  select="round(number(substring-before($node/@*[name() = $fontStyle], '%')) div 100 * number($defaultValue))"
                />
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$defaultValue"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <!-- Compute font size when style has raised/lowered text -->
  <xsl:template name="ComputePositionFontSize">
    <!-- second percentage value specifies size -->
    <xsl:variable name="percentValue">
      <xsl:if
        test="not(substring(@style:text-position, 1, 3) = 'sub') and not(substring(@style:text-position, 1, 5) = 'super') ">
        <xsl:value-of select="substring-before(substring-after(@style:text-position, ' '), '%')"/>
      </xsl:if>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$percentValue != '' ">
        <xsl:variable name="fontHeight">
          <xsl:call-template name="computeSize"/>
        </xsl:variable>
        <xsl:value-of select="round(number($percentValue) div 100 * $fontHeight)"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="@fo:font-size">
          <xsl:call-template name="computeSize"/>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Single tab-stop processing -->
  <xsl:template name="tabStop">
    <xsl:param name="styleType"/>
    <w:tab>
      <!-- position of the tab stop with respect to the current page margin -->
      <xsl:attribute name="w:pos">
        <xsl:variable name="margin">
          <xsl:call-template name="twips-measure">
            <xsl:with-param name="length">
              <xsl:choose>
                <xsl:when test="ancestor::style:paragraph-properties/@fo:margin-left">
                  <xsl:value-of select="ancestor::style:paragraph-properties/@fo:margin-left"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:variable name="parentStyleName"
                    select="ancestor::style:style/@style:parent-style-name"/>
                  <xsl:for-each select="document('styles.xml')">
                    <xsl:choose>
                      <xsl:when
                        test="key('styles',$parentStyleName)/style:paragraph-properties/@fo:margin-left">
                        <xsl:value-of
                          select="key('styles',$parentStyleName)/style:paragraph-properties/@fo:margin-left"
                        />
                      </xsl:when>
                      <xsl:otherwise>0</xsl:otherwise>
                    </xsl:choose>
                  </xsl:for-each>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="position">
          <xsl:call-template name="twips-measure">
            <xsl:with-param name="length">
              <xsl:choose>
                <!-- particular case : right tab in indexes -->
                <xsl:when
                  test="self::text:index-entry-tab-stop[@style:type = 'right' and not(@style:position)]">
                  <xsl:for-each select="document('styles.xml')">
                    <xsl:variable name="pageW">
                      <xsl:call-template name="twips-measure">
                        <xsl:with-param name="length"
                          select="key('page-layouts', $default-master-style/@style:page-layout-name)[1]/style:page-layout-properties/@fo:page-width"
                        />
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:variable name="pageMarginL">
                      <xsl:call-template name="twips-measure">
                        <xsl:with-param name="length"
                          select="key('page-layouts', $default-master-style/@style:page-layout-name)[1]/style:page-layout-properties/@fo:margin-left"
                        />
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:variable name="pageMarginR">
                      <xsl:call-template name="twips-measure">
                        <xsl:with-param name="length"
                          select="key('page-layouts', $default-master-style/@style:page-layout-name)[1]/style:page-layout-properties/@fo:margin-right"
                        />
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:value-of select="$pageW - $pageMarginR - $pageMarginL"/>
                  </xsl:for-each>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="@style:position"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="(contains(ancestor::style:style/@style:name,'Contents_20')) and not (contains(ancestor::style:style/@style:name,'Contents_20_Heading'))">
            <xsl:value-of select="0"/>
          </xsl:when>
          <xsl:when test="ancestor::office:styles or ancestor::office:automatic-styles">
            <xsl:value-of select="$margin + $position"/>
          </xsl:when>
          <!-- particular case : indexes -->
          <xsl:when test="self::text:index-entry-tab-stop">
            <xsl:value-of select="$position"/>
          </xsl:when>
          <!-- paragraph may be embedded in columns -->
          <xsl:otherwise>
            <xsl:variable name="columnNumber">
              <xsl:choose>
                <!-- TODO : There can be many page-layouts in the document.
                  We need to know the master page in which our object is located. Then, we'll get the right page-layout.
                  pageLayoutName = key('master-pages', $masterPageName)[1]/@style:page-layout-name
                  key('page-layouts', $pageLayoutName)[1]/style:page-layout-properties/style:columns[@fo:column-count > 0]
                -->
                <xsl:when
                  test="document('styles.xml')/office:document-styles/office:automatic-styles/style:page-layout/style:page-layout-properties/style:columns[@fo:column-count > 0]">
                  <xsl:value-of
                    select="document('styles.xml')/office:document-styles/office:automatic-styles/style:page-layout/style:page-layout-properties/style:columns/@fo:column-count"
                  />
                </xsl:when>
                <xsl:otherwise>1</xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:variable name="columnGap">
              <xsl:choose>
                <xsl:when
                  test="document('styles.xml')/office:document-styles/office:automatic-styles/style:page-layout/style:page-layout-properties/style:columns/@fo:column-gap">
                  <xsl:call-template name="twips-measure">
                    <xsl:with-param name="length"
                      select="document('styles.xml')/office:document-styles/office:automatic-styles/style:page-layout/style:page-layout-properties/style:columns/@fo:column-gap"
                    />
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:value-of select="round(($margin + $position - $columnGap) div $columnNumber)"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>

      <!-- tab stop type -->
      <xsl:attribute name="w:val">
        <xsl:choose>
          <xsl:when test="$styleType">
            <xsl:value-of select="$styleType"/>
          </xsl:when>
          <xsl:when test="@style:type = 'left' ">left</xsl:when>
          <xsl:when test="@style:type = 'right' ">right</xsl:when>
          <xsl:when test="@style:type = 'center' ">center</xsl:when>
          <xsl:otherwise>left</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>

      <!-- tab leader character -->
      <xsl:attribute name="w:leader">
        <xsl:call-template name="ComputeTabStopLeader"/>
      </xsl:attribute>
    </w:tab>
  </xsl:template>

  <!-- compute the value of leader character for tab-stop -->
  <xsl:template name="ComputeTabStopLeader">
    <xsl:param name="tabStop" select="."/>

    <xsl:choose>
      <!-- leader character in paragraphs -->
      <xsl:when test="$tabStop/@style:leader-text">
        <xsl:choose>
          <xsl:when test="$tabStop/@style:leader-text = '.' ">
            <xsl:text>dot</xsl:text>
          </xsl:when>
          <xsl:when test="$tabStop/@style:leader-text = '-' ">
            <xsl:text>hyphen</xsl:text>
          </xsl:when>
          <xsl:when test="$tabStop/@style:leader-text = '_' ">
            <xsl:text>heavy</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <xsl:message terminate="no">translation.odf2oox.leaderText</xsl:message>
            <xsl:text>none</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="$tabStop/@style:leader-style and not($tabStop/@style:leader-text)">
        <xsl:choose>
          <xsl:when test="$tabStop/@style:leader-style = 'dotted' ">
            <xsl:text>dot</xsl:text>
          </xsl:when>
          <xsl:when test="$tabStop/@style:leader-style = 'solid' ">
            <xsl:text>heavy</xsl:text>
          </xsl:when>
          <xsl:when test="$tabStop/@style:leader-style = 'dash' ">
            <xsl:text>hyphen</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <xsl:if test="$tabStop/@style:leader-style != 'none' ">
              <xsl:message terminate="no">translation.odf2oox.leaderText</xsl:message>
            </xsl:if>
            <xsl:text>none</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <!-- leader character in indexes -->
      <xsl:when test="$tabStop/@style:leader-char">
        <xsl:choose>
          <xsl:when test="$tabStop/@style:leader-char = '.' ">
            <xsl:text>dot</xsl:text>
          </xsl:when>
          <xsl:when test="$tabStop/@style:leader-char = '-' ">
            <xsl:text>hyphen</xsl:text>
          </xsl:when>
          <xsl:when test="$tabStop/@style:leader-char = '_' ">
            <xsl:text>heavy</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <xsl:message terminate="no">translation.odf2oox.leaderText</xsl:message>
            <xsl:text>none</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>none</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- insert Hyperlink and Followed Hyperink styles -->
  <xsl:template name="InsertHyperlinkStyles">
    <xsl:for-each select="document('styles.xml')">
      <!-- hyperlink styles -->
      <w:style w:styleId="Hyperlink" w:type="character">
        <w:name w:val="Hyperlink"/>
        <xsl:choose>
          <xsl:when test="key('styles', 'Internet_20_link')">
            <xsl:if test="key('styles', 'Internet_20_link')/@style:parent-style-name">
              <w:basedOn w:val="{key('styles', 'Internet_20_link')/@style:parent-style-name}"/>
            </xsl:if>
            <!-- style's text properties -->
            <xsl:if test="key('styles', 'Internet_20_link')/style:text-properties">
              <w:rPr>
                <xsl:apply-templates
                  select="key('styles', 'Internet_20_link')/style:text-properties" mode="rPr"/>
              </w:rPr>
            </xsl:if>
          </xsl:when>
          <xsl:otherwise>
            <w:rPr>
              <w:color w:val="000080"/>
              <w:u w:val="single"/>
            </w:rPr>
          </xsl:otherwise>
        </xsl:choose>
      </w:style>
      <!-- visited hyperlink styles -->
      <w:style w:styleId="FollowedHyperlink" w:type="character">
        <w:name w:val="FollowedHyperlink"/>
        <xsl:choose>
          <xsl:when test="key('styles', 'Visited_20_Internet_20_Link')">
            <xsl:if test="key('styles', 'Visited_20_Internet_20_Link')/@style:parent-style-name">
              <w:basedOn
                w:val="{key('styles', 'Visited_20_Internet_20_Link')/@style:parent-style-name}"/>
            </xsl:if>
            <!-- style's text properties -->
            <xsl:if test="key('styles', 'Visited_20_Internet_20_Link')/style:text-properties">
              <w:rPr>
                <xsl:apply-templates
                  select="key('styles', 'Visited_20_Internet_20_Link')/style:text-properties"
                  mode="rPr"/>
              </w:rPr>
            </xsl:if>
          </xsl:when>
          <xsl:otherwise>
            <w:rPr>
              <w:color w:val="800080"/>
              <w:u w:val="single"/>
            </w:rPr>
          </xsl:otherwise>
        </xsl:choose>
      </w:style>
    </xsl:for-each>
  </xsl:template>


  <!-- Page Layout Properties -->
  <xsl:template match="style:page-layout-properties" mode="master-page">
    <!-- background color lost if it is not main layout -->
    <xsl:if
      test="@fo:background-color != 'transparent' and parent::style:page-layout/@style:name != $default-master-style/@style:page-layout-name">
      <xsl:message terminate="no">translation.odf2oox.pageBgColor</xsl:message>
    </xsl:if>
    <!-- page type -->
    <!-- TODO : review. This entails a page jump on an even or odd page with a potential unwanted blank page in between... -->
    <!--xsl:if test="parent::style:page-layout/@style:page-usage">
      <xsl:choose>
        <xsl:when test="parent::style:page-layout/@style:page-usage = 'left' ">
          <w:type w:val="evenPage"/>
        </xsl:when>
        <xsl:when test="parent::style:page-layout/@style:page-usage = 'right' ">
          <w:type w:val="oddPage"/>
        </xsl:when>
      </xsl:choose>
    </xsl:if-->
    <!-- page size -->
    <xsl:choose>
      <xsl:when
        test="not(@style:print-orientation) and not(@fo:page-width) and not(@fo:page-height)">
        <xsl:apply-templates
          select="key('page-layouts', $default-master-style/@style:page-layout-name)[1]/style:page-layout-properties"
          mode="master-page"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="ComputePageSize"/>
      </xsl:otherwise>
    </xsl:choose>
    <!-- page margins -->
    <xsl:choose>
      <xsl:when
        test="@fo:margin-top ='none' and @fo:margin-left='none' and @fo:margin-bottom='none'  and @fo:margin-right='none' ">
        <xsl:apply-templates
          select="key('page-layouts', $default-master-style/@style:page-layout-name)[1]/style:page-layout-properties"
          mode="master-page"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="ComputePageMargins"/>
      </xsl:otherwise>
    </xsl:choose>
    <!-- paper source : must be a decimal number in OOX. Cannot convert codes -->
    <xsl:choose>
      <xsl:when test="@style:paper-tray-name = 'default' ">
        <w:paperSrc w:first="1" w:other="1"/>
      </xsl:when>
      <xsl:when test="@style:paper-tray-name = '' ">
        <w:paperSrc w:first="1" w:other="1"/>
      </xsl:when>
      <xsl:when test="number(@style:paper-tray-name)">
        <w:paperSrc w:first="{@style:paper-tray-name}" w:other="{@style:paper-tray-name}"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="@style:paper-tray-name != '' ">
          <xsl:message terminate="no">translation.odf2oox.paperTray</xsl:message>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
    <!-- page borders -->
    <xsl:choose>
      <xsl:when test="@fo:border and @fo:border != 'none' ">
        <w:pgBorders>
          <xsl:call-template name="InsertBorders">
            <xsl:with-param name="allSides">true</xsl:with-param>
          </xsl:call-template>
        </w:pgBorders>
      </xsl:when>
      <xsl:when test="@fo:border = 'none'">
        <w:pgBorders>
          <xsl:call-template name="InsertEmptyBorders"/>
        </w:pgBorders>
      </xsl:when>
      <xsl:when test="@fo:border-top or @fo:border-left or @fo:border-bottom or @fo:border-right">
        <w:pgBorders>
          <xsl:call-template name="InsertBorders">
            <xsl:with-param name="allSides">false</xsl:with-param>
          </xsl:call-template>
        </w:pgBorders>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="@style:shadow">
          <w:pgBorders>
            <xsl:call-template name="InsertEmptyBorders"/>
          </w:pgBorders>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
    <!-- line numbering -->
    <xsl:for-each
      select="document('styles.xml')/office:document-styles/office:styles/text:linenumbering-configuration">
      <xsl:if test="not(@text:number-lines='false')">
        <w:lnNumType>
          <xsl:if
            test="@text:style-name or @style:num-format or @text:number-position or @text:count-text-boxes or @text:linenumbering-separator">
            <!-- TODO : Considering the above test, this message may not be really explicit... -->
            <xsl:message terminate="no">translation.odf2oox.lineNumbering</xsl:message>
          </xsl:if>
          <xsl:if test="@text:increment">
            <xsl:attribute name="w:countBy">
              <xsl:value-of select="@text:increment"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="@text:offset">
            <xsl:attribute name="w:distance">
              <xsl:call-template name="twips-measure">
                <xsl:with-param name="length" select="@text:offset"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
          <xsl:attribute name="w:restart">
            <xsl:choose>
              <xsl:when test="@text:restart-on-page">newPage</xsl:when>
              <xsl:otherwise>continuous</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </w:lnNumType>
      </xsl:if>
    </xsl:for-each>
    <!-- style of page number -->
    <xsl:choose>
      <xsl:when test="@style:num-format='i'">
        <w:pgNumType w:fmt="lowerRoman"/>
      </xsl:when>
      <xsl:when test="@style:num-format='I'">
        <w:pgNumType w:fmt="upperRoman"/>
      </xsl:when>
      <xsl:when test="@style:num-format='a'">
        <w:pgNumType w:fmt="lowerLetter"/>
      </xsl:when>
      <xsl:when test="@style:num-format='A'">
        <w:pgNumType w:fmt="upperLetter"/>
      </xsl:when>
    </xsl:choose>
    <!-- page columns -->
    <xsl:apply-templates select="style:columns" mode="columns"/>
  </xsl:template>

  <!-- Page size and orientation properties -->
  <xsl:template name="ComputePageSize">
    <w:pgSz>
      <xsl:if test="@style:print-orientation != 'none' ">
        <xsl:attribute name="w:orient">
          <xsl:choose>
            <xsl:when test="@style:print-orientation = 'landscape' ">landscape</xsl:when>
            <xsl:otherwise>portrait</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@fo:page-width != 'none' ">
        <xsl:attribute name="w:w">
          <xsl:call-template name="twips-measure">
            <xsl:with-param name="length" select="@fo:page-width"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@fo:page-height != 'none' ">
        <xsl:attribute name="w:h">
          <xsl:call-template name="twips-measure">
            <xsl:with-param name="length" select="@fo:page-height"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
    </w:pgSz>
  </xsl:template>

  <!-- Page margin properties -->
  <xsl:template name="ComputePageMargins">
    <!-- report loss of header/footer properties -->
    <xsl:variable name="header-properties"
      select="parent::style:page-layout/style:header-style/style:header-footer-properties"/>
    <xsl:variable name="footer-properties"
      select="parent::style:page-layout/style:footer-style/style:header-footer-properties"/>
    <!-- this attribute exists when the user choose to automatically adapt the header/footer height,
          otherwise svg:height is used and it gives the exact header/footer height -->
    <xsl:if test="$header-properties/@fo:min-height">
      <xsl:message terminate="no">translation.odf2oox.headerDistance</xsl:message>
    </xsl:if>
    <xsl:if test="$footer-properties/@fo:min-height">
      <xsl:message terminate="no">translation.odf2oox.footerDistance</xsl:message>
    </xsl:if>
    <w:pgMar>
      <xsl:if
        test="@fo:margin != 'none' or @fo:margin-top != 'none' or @fo:padding != 'none' or @fo:padding-top != 'none' ">
        <xsl:attribute name="w:top">
          <!-- distance from top edge of the page to document body -->
          <xsl:variable name="top">
            <xsl:call-template name="GetPageMargin">
              <xsl:with-param name="side">top</xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:choose>
            <!-- additional header height -->
            <xsl:when test="$header-properties">
              <xsl:variable name="min-header-height">
                <xsl:call-template name="twips-measure">
                  <xsl:with-param name="length">
                    <xsl:choose>
                      <xsl:when test="$header-properties/@fo:min-height">
                        <xsl:value-of select="$header-properties/@fo:min-height"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="$header-properties/@svg:height"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:variable>
              <xsl:value-of select="$top + $min-header-height"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$top"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
      <xsl:if
        test="@fo:margin != 'none' or @fo:margin-left != 'none' or @fo:padding != 'none' or @fo:padding-left != 'none' ">
        <xsl:attribute name="w:left">
          <xsl:call-template name="GetPageMargin">
            <xsl:with-param name="side">left</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if
        test="@fo:margin != 'none' or @fo:margin-bottom != 'none' or @fo:padding != 'none' or @fo:padding-bottom != 'none' ">
        <xsl:attribute name="w:bottom">
          <xsl:variable name="bottom">
            <xsl:call-template name="GetPageMargin">
              <xsl:with-param name="side">bottom</xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:choose>
            <!-- additional footer height -->
            <xsl:when test="$footer-properties">
              <xsl:variable name="min-footer-height">
                <xsl:call-template name="twips-measure">
                  <xsl:with-param name="length">
                    <xsl:choose>
                      <xsl:when test="$footer-properties/@fo:min-height">
                        <xsl:value-of select="$footer-properties/@fo:min-height"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="$footer-properties/@svg:height"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:variable>
              <xsl:value-of select="$bottom + $min-footer-height"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$bottom"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
      <xsl:if
        test="@fo:margin != 'none' or @fo:margin-right != 'none' or @fo:padding != 'none' or @fo:padding-right != 'none' ">
        <xsl:attribute name="w:right">
          <xsl:call-template name="GetPageMargin">
            <xsl:with-param name="side">right</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <!-- header and footer distance from page : get page margin -->
      <xsl:if
        test="@fo:margin != 'none' or @fo:margin-top != 'none' or @fo:padding != 'none' or @fo:padding-top != 'none' ">
        <xsl:attribute name="w:header">
          <xsl:call-template name="GetPageMargin">
            <xsl:with-param name="side">top</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if
        test="@fo:margin != 'none' or @fo:margin-bottom != 'none' or @fo:padding != 'none' or @fo:padding-bottom != 'none' ">
        <xsl:attribute name="w:footer">
          <xsl:call-template name="GetPageMargin">
            <xsl:with-param name="side">bottom</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <!-- Don't exist in odt format -->
      <xsl:attribute name="w:gutter">0</xsl:attribute>
    </w:pgMar>
  </xsl:template>

  <!-- Calculate one side of page -->
  <xsl:template name="GetPageMargin">
    <xsl:param name="side"/>

    <xsl:variable name="padding">
      <xsl:choose>
        <xsl:when test="@fo:padding != 'none' ">
          <xsl:call-template name="indent-val">
            <xsl:with-param name="length" select="@fo:padding"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="attribute::node()[name() = concat('fo:padding-',$side)] != 'none' ">
          <xsl:call-template name="indent-val">
            <xsl:with-param name="length"
              select="attribute::node()[name() = concat('fo:padding-',$side)]"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="margin">
      <xsl:choose>
        <xsl:when test="@fo:margin != 'none' ">
          <xsl:call-template name="twips-measure">
            <xsl:with-param name="length" select="@fo:margin"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="attribute::node()[name() = concat('fo:margin-',$side)] != 'none' ">
          <xsl:call-template name="twips-measure">
            <xsl:with-param name="length"
              select="attribute::node()[name() = concat('fo:margin-',$side)]"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:value-of select="$padding + $margin"/>
  </xsl:template>



  <xsl:template match="style:columns" mode="columns">
    <w:cols>
      <!-- nb columns -->
      <xsl:if test="@fo:column-count">
        <xsl:attribute name="w:num">
          <xsl:value-of select="@fo:column-count"/>
        </xsl:attribute>
      </xsl:if>
      <!-- separator -->
      <xsl:if test="style:column-sep">
        <!-- report lost attributes -->
        <xsl:message terminate="no">translation.odf2oox.columnSeparatorAttributes</xsl:message>
        <xsl:attribute name="w:sep">
          <xsl:value-of select="1"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:choose>
        <xsl:when test="@fo:column-gap">
          <xsl:attribute name="w:space">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length" select="@fo:column-gap"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="style:column">
          <xsl:attribute name="w:equalWidth">0</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="w:equalWidth">1</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:if test="style:column">
        <!-- for each column -->
        <xsl:for-each select="style:column">
          <w:col>
            <!-- the left and right spaces -->
            <xsl:variable name="start">
              <xsl:call-template name="twips-measure">
                <xsl:with-param name="length" select="@fo:start-indent"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="end">
              <xsl:call-template name="twips-measure">
                <xsl:with-param name="length" select="@fo:end-indent"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select="number($start + $end)"/>
            </xsl:variable>
            <!-- space -->
            <xsl:attribute name="w:space">
              <!-- odt separate space between two columns ( col 1 : fo:end-indent and col 2 : fo:start-indent ) -->
              <xsl:choose>
                <xsl:when test="following-sibling::style:column/@fo:start-indent">
                  <xsl:variable name="followingStart">
                    <xsl:call-template name="twips-measure">
                      <xsl:with-param name="length"
                        select="following-sibling::style:column/@fo:start-indent"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:value-of select="number($followingStart + $end)"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$end"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <!-- width -->
            <xsl:attribute name="w:w">
              <xsl:value-of select="substring-before(@style:rel-width,'*') - $width"/>
            </xsl:attribute>
          </w:col>
        </xsl:for-each>
      </xsl:if>
    </w:cols>
  </xsl:template>


  <!-- Section properties -->
  <xsl:template match="style:section-properties" mode="section">
    <!-- margins -->
    <xsl:if test="@fo:margin-left != '' or @fo:margin-right != '' ">
      <w:pgMar>
        <xsl:if test="@fo:margin-left != '' or @fo:margin-left != 'none' ">
          <xsl:attribute name="w:left">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length" select="@fo:margin-left"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="@fo:margin-right != '' or @fo:margin-right != 'none' ">
          <xsl:attribute name="w:right">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length" select="@fo:margin-right"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
      </w:pgMar>
    </xsl:if>
    <!-- columns -->
    <xsl:apply-templates select="style:columns" mode="columns"/>
  </xsl:template>


  <!-- Insert borders, depending on allSides bool. If true, create all borders. -->
  <xsl:template name="InsertBorders">
    <xsl:param name="allSides" select="'false'"/>
    <xsl:param name="node" select="."/>

    <xsl:if test="$allSides='true' or ($node/@fo:border-top and ($node/@fo:border-top != 'none'))">
      <w:top>
        <xsl:call-template name="border">
          <xsl:with-param name="side" select="'top'"/>
          <xsl:with-param name="node" select="$node"/>
        </xsl:call-template>
      </w:top>
    </xsl:if>
    <xsl:if test="$node/@fo:border-top = 'none'">
      <w:top>
        <xsl:call-template name="InsertEmptyBorder">
          <xsl:with-param name="side" select="'top'"/>
          <xsl:with-param name="node" select="$node"/>
        </xsl:call-template>
      </w:top>
    </xsl:if>
    <xsl:if test="$allSides='true' or ($node/@fo:border-left and ($node/@fo:border-left != 'none'))">
      <w:left>
        <xsl:call-template name="border">
          <xsl:with-param name="side" select="'left'"/>
          <xsl:with-param name="node" select="$node"/>
        </xsl:call-template>
      </w:left>
    </xsl:if>
    <xsl:if test="$node/@fo:border-left = 'none'">
      <w:left>
        <xsl:call-template name="InsertEmptyBorder">
          <xsl:with-param name="side" select="'left'"/>
          <xsl:with-param name="node" select="$node"/>
        </xsl:call-template>
      </w:left>
    </xsl:if>
    <xsl:if
      test="$allSides='true' or ($node/@fo:border-bottom and ($node/@fo:border-bottom != 'none'))">
      <w:bottom>
        <xsl:call-template name="border">
          <xsl:with-param name="side" select="'bottom'"/>
          <xsl:with-param name="node" select="$node"/>
        </xsl:call-template>
      </w:bottom>
    </xsl:if>
    <xsl:if test="$node/@fo:border-bottom = 'none'">
      <w:bottom>
        <xsl:call-template name="InsertEmptyBorder">
          <xsl:with-param name="side" select="'bottom'"/>
          <xsl:with-param name="node" select="$node"/>
        </xsl:call-template>
      </w:bottom>
    </xsl:if>
    <xsl:if
      test="$allSides='true' or ($node/@fo:border-right and ($node/@fo:border-right != 'none'))">
      <w:right>
        <xsl:call-template name="border">
          <xsl:with-param name="side" select="'right'"/>
          <xsl:with-param name="node" select="$node"/>
        </xsl:call-template>
      </w:right>
    </xsl:if>
    <xsl:if test="$node/@fo:border-right = 'none'">
      <w:right>
        <xsl:call-template name="InsertEmptyBorder">
          <xsl:with-param name="side" select="'right'"/>
          <xsl:with-param name="node" select="$node"/>
        </xsl:call-template>
      </w:right>
    </xsl:if>
    <xsl:if
      test="$allSides='true' and self::style:paragraph-properties and $node/@style:join-border='false'">
      <w:between>
        <xsl:call-template name="border">
          <xsl:with-param name="side" select="'middle'"/>
          <xsl:with-param name="node" select="$node"/>
        </xsl:call-template>
      </w:between>
    </xsl:if>
    <xsl:if
      test="$allSides='false' and $node/@style:join-border='false' and (
      ($node/@fo:border-bottom and ($node/@fo:border-bottom != 'none'))
      or ($node/@fo:border-top and ($node/@fo:border-top != 'none')) )">
      <w:between>
        <xsl:call-template name="border">
          <xsl:with-param name="side" select="'middle'"/>
          <xsl:with-param name="node" select="$node"/>
        </xsl:call-template>
      </w:between>
    </xsl:if>
  </xsl:template>

  <!-- insert borders with 'none' value (used to override parent properties) -->
  <xsl:template name="InsertEmptyBorders">
    <xsl:param name="node" select="."/>

    <xsl:choose>
      <xsl:when test="not($node/@style:shadow != 'none')">
        <w:top w:val="none"/>
        <w:left w:val="none"/>
        <w:bottom w:val="none"/>
        <w:right w:val="none"/>
      </xsl:when>
      <xsl:when test="$node/@style:shadow != 'none' ">
        <!-- border shadow may not be displayed properly -->
        <xsl:message terminate="no">translation.odf2oox.paragraphPageShadow</xsl:message>
        <xsl:variable name="firstShadow">
          <xsl:value-of select=" substring-before(substring-after($node/@style:shadow, ' '), ' ')"/>
        </xsl:variable>
        <xsl:variable name="secondShadow">
          <xsl:value-of select=" substring-after(substring-after($node/@style:shadow, ' '), ' ')"/>
        </xsl:variable>
        <w:top>
          <xsl:attribute name="w:val">none</xsl:attribute>
          <xsl:if test="contains($secondShadow,'-')">
            <xsl:attribute name="w:shadow">true</xsl:attribute>
          </xsl:if>
        </w:top>
        <w:left>
          <xsl:attribute name="w:val">none</xsl:attribute>
          <xsl:if test="contains($firstShadow,'-')">
            <xsl:attribute name="w:shadow">true</xsl:attribute>
          </xsl:if>
        </w:left>
        <w:bottom>
          <xsl:attribute name="w:val">none</xsl:attribute>
          <xsl:if test="not(contains($secondShadow,'-'))">
            <xsl:attribute name="w:shadow">true</xsl:attribute>
          </xsl:if>
        </w:bottom>
        <w:right>
          <xsl:attribute name="w:val">none</xsl:attribute>
          <xsl:if test="not(contains($firstShadow,'-'))">
            <xsl:attribute name="w:shadow">true</xsl:attribute>
          </xsl:if>
        </w:right>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- insert one empty border attributes -->
  <xsl:template name="InsertEmptyBorder">
    <xsl:param name="side"/>
    <xsl:param name="node" select="."/>

    <xsl:attribute name="w:val">none</xsl:attribute>

    <!-- insert shadow -->
    <xsl:if test="$node/@style:shadow != 'none' ">
      <!-- border shadow may not be displayed properly -->
      <xsl:message terminate="no">translation.odf2oox.paragraphPageShadow</xsl:message>
      <xsl:variable name="firstShadow">
        <xsl:value-of select=" substring-before(substring-after($node/@style:shadow, ' '), ' ')"/>
      </xsl:variable>
      <xsl:variable name="secondShadow">
        <xsl:value-of select=" substring-after(substring-after($node/@style:shadow, ' '), ' ')"/>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$side = 'top' and contains($secondShadow,'-')">
          <xsl:attribute name="w:shadow">true</xsl:attribute>
        </xsl:when>
        <xsl:when test="$side = 'left' and contains($firstShadow,'-')">
          <xsl:attribute name="w:shadow">true</xsl:attribute>
        </xsl:when>
        <xsl:when test="$side = 'bottom' and not(contains($secondShadow,'-'))">
          <xsl:attribute name="w:shadow">true</xsl:attribute>
        </xsl:when>
        <xsl:when test="$side = 'top' and not(contains($firstShadow,'-'))">
          <xsl:attribute name="w:shadow">true</xsl:attribute>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
  </xsl:template>

  <!-- Attributes of a border element. -->
  <xsl:template name="border">
    <xsl:param name="side"/>
    <xsl:param name="node" select="."/>

    <xsl:variable name="borderStr">
      <xsl:choose>
        <xsl:when test="$side = 'tl-br' or $side = 'bl-tr' ">
          <xsl:value-of select="$node/@*[name()=concat('style:diagonal-', $side)]"/>
        </xsl:when>
        <xsl:when test="$node/@fo:border">
          <xsl:value-of select="$node/@fo:border"/>
        </xsl:when>
        <xsl:when test="$side='middle'">
          <xsl:choose>
            <xsl:when test="$node/@fo:border-top != 'none'">
              <xsl:value-of select="$node/@fo:border-top"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$node/@fo:border-bottom"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$node/@*[name()=concat('fo:border-', $side)]"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <!-- padding = 0 if side border not defined ! -->
    <xsl:variable name="padding">
      <xsl:if test="not($node/self::style:table-cell-properties)">
        <xsl:choose>
          <xsl:when test="$node/@fo:border and $node/@fo:padding">
            <xsl:value-of select="$node/@fo:padding"/>
          </xsl:when>
          <xsl:when
            test="$node/@*[name()=concat('fo:border-', $side)] != 'none' and $node/@fo:padding">
            <xsl:value-of select="$node/@fo:padding"/>
          </xsl:when>
          <xsl:when test="$side = 'middle' ">
            <xsl:choose>
              <xsl:when test="$node/@fo:padding-top != 'none'">
                <xsl:value-of select="$node/@fo:padding-top"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$node/@fo:padding-bottom"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:when test="$node/@*[name()=concat('fo:padding-', $side)]">
            <xsl:value-of select="$node/@*[name()=concat('fo:padding-', $side)]"/>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:if>
    </xsl:variable>

    <!-- border line width -->
    <xsl:variable name="borderLineWidth">
      <xsl:call-template name="GetBorderLineWidth">
        <xsl:with-param name="side" select="$side"/>
        <xsl:with-param name="node" select="$node"/>
      </xsl:call-template>
    </xsl:variable>


    <xsl:attribute name="w:val">
      <xsl:call-template name="GetBorderStyle">
        <xsl:with-param name="side" select="$side"/>
        <xsl:with-param name="borderStr" select="$borderStr"/>
        <xsl:with-param name="borderLineWidth" select="$borderLineWidth"/>
      </xsl:call-template>
    </xsl:attribute>

    <xsl:attribute name="w:sz">
      <xsl:call-template name="CheckBorder">
        <xsl:with-param name="unit">eightspoint</xsl:with-param>
        <xsl:with-param name="length">
          <xsl:choose>
            <xsl:when test="$borderLineWidth != '' ">
              <xsl:call-template name="ComputeBorderLineWidth">
                <xsl:with-param name="borderLineWidth" select="$borderLineWidth"/>
                <xsl:with-param name="unit">eightspoint</xsl:with-param>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="eightspoint-measure">
                <xsl:with-param name="length">
                  <xsl:call-template name="ParseBorderWidth">
                    <xsl:with-param name="border" select="$borderStr"/>
                  </xsl:call-template>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>

    <xsl:if test="$padding != '' ">
      <xsl:attribute name="w:space">
        <xsl:call-template name="padding-val">
          <xsl:with-param name="length" select="$padding"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>

    <xsl:attribute name="w:color">
      <xsl:value-of select="substring($borderStr, string-length($borderStr) -5, 6)"/>
    </xsl:attribute>

    <xsl:if test="$node/@style:shadow != 'none' ">
      <xsl:variable name="firstShadow">
        <xsl:value-of select=" substring-before(substring-after($node/@style:shadow, ' '), ' ')"/>
      </xsl:variable>
      <xsl:variable name="secondShadow">
        <xsl:value-of select=" substring-after(substring-after($node/@style:shadow, ' '), ' ')"/>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$side = 'top' and contains($secondShadow,'-')">
          <!-- border shadow may not be displayed properly -->
          <xsl:message terminate="no">translation.odf2oox.paragraphPageShadow</xsl:message>
          <xsl:attribute name="w:shadow">true</xsl:attribute>
        </xsl:when>
        <xsl:when test="$side = 'left' and contains($firstShadow,'-')">
          <!-- border shadow may not be displayed properly -->
          <xsl:message terminate="no">translation.odf2oox.paragraphPageShadow</xsl:message>
          <xsl:attribute name="w:shadow">true</xsl:attribute>
        </xsl:when>
        <xsl:when test="$side = 'bottom' and not(contains($secondShadow,'-'))">
          <xsl:attribute name="w:shadow">true</xsl:attribute>
        </xsl:when>
        <xsl:when test="$side = 'right' and not(contains($firstShadow,'-'))">
          <xsl:attribute name="w:shadow">true</xsl:attribute>
        </xsl:when>
      </xsl:choose>
    </xsl:if>

  </xsl:template>

  <!-- find which of the three arguments in the string is a length value, and return it -->
  <xsl:template name="ParseBorderWidth">
    <xsl:param name="border"/>
    <xsl:param name="argNumber" select="1"/>
    <!-- find the value corresponding to the argument position -->
    <xsl:variable name="argument">
      <xsl:choose>
        <xsl:when test="$argNumber = 1">
          <xsl:value-of select="substring-before($border, ' ')"/>
        </xsl:when>
        <xsl:when test="$argNumber = 2">
          <xsl:value-of select="substring-before(substring-after($border, ' '), ' ')"/>
        </xsl:when>
        <xsl:when test="$argNumber = 3">
          <xsl:value-of select="substring-after(substring-after($border, ' '), ' ')"/>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="argVaue">
      <xsl:call-template name="GetValue">
        <xsl:with-param name="length" select="$argument"/>
      </xsl:call-template>
    </xsl:variable>
    <!-- if value is a number, return it. Otherwise, switch to next argument -->
    <xsl:choose>
      <xsl:when test="number($argVaue)">
        <xsl:value-of select="$argument"/>
      </xsl:when>
      <xsl:when test="$argNumber &lt; 3">
        <xsl:call-template name="ParseBorderWidth">
          <xsl:with-param name="border" select="$border"/>
          <xsl:with-param name="argNumber" select="$argNumber + 1"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>0</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- compute the width of a border using border-line-width attribute -->
  <xsl:template name="ComputeBorderLineWidth">
    <xsl:param name="borderLineWidth"/>
    <xsl:param name="unit"/>

    <xsl:variable name="inner">
      <xsl:call-template name="ConvertMeasure">
        <xsl:with-param name="length" select="substring-before($borderLineWidth,' ')"/>
        <xsl:with-param name="unit" select="$unit"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="between">
      <xsl:call-template name="ConvertMeasure">
        <xsl:with-param name="length"
          select="substring-before(substring-after($borderLineWidth,' '),' ')"/>
        <xsl:with-param name="unit" select="$unit"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="outer">
      <xsl:call-template name="ConvertMeasure">
        <xsl:with-param name="length"
          select="substring-after(substring-after($borderLineWidth,' '),' ')"/>
        <xsl:with-param name="unit" select="$unit"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select="$inner + $outer + $between"/>
  </xsl:template>

  <!-- get border line width attribute -->
  <xsl:template name="GetBorderLineWidth">
    <xsl:param name="side"/>
    <xsl:param name="node"/>
    <xsl:choose>
      <xsl:when test="$side = 'tl-br' or $side = 'bl-tr' ">
        <xsl:value-of select="$node/@*[name()=concat('style:diagonal-', $side, '-widths')]"/>
      </xsl:when>
      <xsl:when test="$node/@style:border-line-width">
        <xsl:value-of select="$node/@style:border-line-width"/>
      </xsl:when>
      <xsl:when test="$side='middle'">
        <xsl:choose>
          <xsl:when test="$node/@style:border-line-width-top">
            <xsl:value-of select="$node/@style:border-line-width-top"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$node/@fo:border-line-width-bottom"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="attribute::node()[name()=concat('style:border-line-width-',$side)]"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Compute values to be added for indent value. -->
  <xsl:template name="ComputeAdditionalIndent">
    <xsl:param name="side"/>
    <xsl:param name="style"/>
    <!-- If there is a border, the indent will have to take it into consideration. -->
    <xsl:variable name="BorderWidth">
      <xsl:call-template name="GetBorderWidth">
        <xsl:with-param name="side" select="$side"/>
        <xsl:with-param name="style" select="$style"/>
      </xsl:call-template>
    </xsl:variable>
    <!-- indent generated by the borders. Limited to 620 twip. -->
    <xsl:variable name="Padding">
      <xsl:call-template name="ComputePaddingValue">
        <xsl:with-param name="side" select="$side"/>
        <xsl:with-param name="style" select="$style"/>
      </xsl:call-template>
    </xsl:variable>
    <!-- paragraph margins -->
    <xsl:variable name="Margin">
      <xsl:call-template name="GetParagraphMargin">
        <xsl:with-param name="side" select="$side"/>
        <xsl:with-param name="style" select="$style"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select="$BorderWidth + $Padding + $Margin"/>
  </xsl:template>


  <!-- Compute the value of a border width -->
  <xsl:template name="GetBorderWidth">
    <xsl:param name="side"/>
    <xsl:param name="style"/>

    <xsl:variable name="styleName" select="$style/@style:parent-style-name"/>
    <xsl:variable name="borderWidth">
      <xsl:choose>
        <xsl:when
          test="$style/style:paragraph-properties/@fo:border and $style/style:paragraph-properties/@fo:border!='none'">
          <xsl:call-template name="GetConsistentBorderValue">
            <xsl:with-param name="side" select="$side"/>
            <xsl:with-param name="borderAttibute"
              select="$style/style:paragraph-properties/@fo:border"/>
            <xsl:with-param name="node" select="$style/style:paragraph-properties"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when
          test="$style/style:paragraph-properties/@*[name()=concat('fo:border-', $side)] and $style/style:paragraph-properties/@*[name()=concat('fo:border-', $side)] != 'none' ">
          <xsl:call-template name="GetConsistentBorderValue">
            <xsl:with-param name="side" select="$side"/>
            <xsl:with-param name="borderAttibute"
              select="$style/style:paragraph-properties/@*[name()=concat('fo:border-', $side)]"/>
            <xsl:with-param name="node" select="$style/style:paragraph-properties"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="$style/ancestor::office:automatic-styles">
          <xsl:for-each select="document('styles.xml')">
            <xsl:choose>
              <xsl:when
                test="key('styles',$styleName)[1]/style:paragraph-properties/@fo:border and key('styles',$styleName)[1]/style:paragraph-properties/@fo:border != 'none' ">
                <xsl:call-template name="GetConsistentBorderValue">
                  <xsl:with-param name="side" select="$side"/>
                  <xsl:with-param name="borderAttibute"
                    select="key('styles', $styleName)/style:paragraph-properties/@fo:border"/>
                  <xsl:with-param name="node"
                    select="key('styles', $styleName)/style:paragraph-properties"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when
                test="key('styles', $styleName)/style:paragraph-properties/@*[name()=concat('fo:border-',$side)] and key('styles', $styleName)/style:paragraph-properties/@*[name()=concat('fo:border-',$side)] != 'none'">
                <xsl:call-template name="GetConsistentBorderValue">
                  <xsl:with-param name="side" select="$side"/>
                  <xsl:with-param name="borderAttibute"
                    select="key('styles', $styleName)/style:paragraph-properties/@*[name()=concat('fo:border-', $side)]"/>
                  <xsl:with-param name="node"
                    select="key('styles', $styleName)/style:paragraph-properties"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:otherwise>0</xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <!-- if double border (should then contain ' '), compute width -->
    <xsl:call-template name="CheckBorder">
      <xsl:with-param name="unit">twips</xsl:with-param>
      <xsl:with-param name="length">
        <xsl:choose>
          <xsl:when test="contains($borderWidth, ' ')">
            <xsl:call-template name="ComputeBorderLineWidth">
              <xsl:with-param name="unit">twips</xsl:with-param>
              <xsl:with-param name="borderLineWidth" select="$borderWidth"> </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length">
                <xsl:value-of select="$borderWidth"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <!-- get the consistent value of border width using context -->
  <xsl:template name="GetConsistentBorderValue">
    <xsl:param name="side"/>
    <xsl:param name="node"/>
    <xsl:param name="borderAttibute"/>

    <xsl:choose>
      <xsl:when test="contains($borderAttibute, 'double')">
        <xsl:call-template name="GetBorderLineWidth">
          <xsl:with-param name="side" select="$side"/>
          <xsl:with-param name="node" select="$node"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="substring-before($borderAttibute,' ')"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Compute the value of padding -->
  <xsl:template name="ComputePaddingValue">
    <xsl:param name="side"/>
    <xsl:param name="style"/>
    <xsl:variable name="styleName" select="$style/@style:parent-style-name"/>

    <!-- padding = 0 if side border not defined ! -->
    <xsl:variable name="paddingVal">
      <xsl:choose>
        <xsl:when
          test="$style/style:paragraph-properties/@fo:border and $style/style:paragraph-properties/@fo:padding">
          <xsl:value-of select="$style/style:paragraph-properties/@fo:padding"/>
        </xsl:when>
        <xsl:when
          test="$style/style:paragraph-properties/@*[name()=concat('fo:border-', $side)] != 'none' and $style/style:paragraph-properties/@fo:padding">
          <xsl:value-of select="$style/style:paragraph-properties/@fo:padding"/>
        </xsl:when>
        <xsl:when
          test="$style/style:paragraph-properties/@*[name()=concat('fo:border-', $side)] != 'none' and $style/style:paragraph-properties/@*[name()=concat('fo:padding-',$side)]">
          <xsl:value-of
            select="$style/style:paragraph-properties/@*[name()=concat('fo:padding-',$side)]"/>
        </xsl:when>
        <xsl:when test="$style/ancestor::office:automatic-styles">
          <xsl:for-each select="document('styles.xml')">
            <xsl:choose>
              <xsl:when
                test="key('styles', $styleName)/style:paragraph-properties/@fo:border and key('styles', $styleName)/style:paragraph-properties/@fo:padding">
                <xsl:value-of
                  select="key('styles', $styleName)/style:paragraph-properties/@fo:padding"/>
              </xsl:when>
              <xsl:when
                test="key('styles', $styleName)/style:paragraph-properties/@*[name()=concat('fo:border-',$side)] != 'none' and key('styles', $styleName)/style:paragraph-properties/@fo:padding">
                <xsl:value-of
                  select="key('styles', $styleName)/style:paragraph-properties/@fo:padding"/>
              </xsl:when>
              <xsl:when
                test="key('styles', $styleName)/style:paragraph-properties/@*[name()=concat('fo:border-',$side)] != 'none' and key('styles', $styleName)/style:paragraph-properties/@*[name()=concat('fo:padding-',$side)]">
                <xsl:value-of
                  select="key('styles', $styleName)/style:paragraph-properties/@*[name()=concat('fo:padding-',$side)]"
                />
              </xsl:when>
              <xsl:otherwise>0</xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:call-template name="indent-val">
      <xsl:with-param name="length">
        <xsl:value-of select="$paddingVal"/>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>


  <!-- Compute a paragraph margin value -->
  <xsl:template name="GetParagraphMargin">
    <xsl:param name="side"/>
    <xsl:param name="style"/>
    <xsl:variable name="styleName" select="$style/@style:parent-style-name"/>
    <xsl:variable name="marginVal">
      <xsl:choose>
        <xsl:when test="$style/style:paragraph-properties/@*[name()=concat('fo:margin-',$side)]">
          <xsl:value-of
            select="$style/style:paragraph-properties/@*[name()=concat('fo:margin-',$side)]"/>
        </xsl:when>
        <xsl:when test="$style/ancestor::office:automatic-styles">
          <xsl:for-each select="document('styles.xml')">
            <xsl:choose>
              <xsl:when
                test="key('styles', $styleName)/style:paragraph-properties/@*[name()=concat('fo:margin-',$side)]">
                <xsl:value-of
                  select="key('styles', $styleName)/style:paragraph-properties/@*[name()=concat('fo:margin-',$side)]"
                />
              </xsl:when>
              <xsl:otherwise>0</xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:call-template name="twips-measure">
      <xsl:with-param name="length">
        <xsl:value-of select="$marginVal"/>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>


  <!-- calculate a first line indent -->
  <xsl:template name="GetFirstLineIndent">
    <xsl:param name="style"/>
    <xsl:variable name="styleName" select="$style/@style:parent-style-name"/>
    <xsl:call-template name="twips-measure">
      <xsl:with-param name="length">
        <xsl:choose>
          <xsl:when
            test="$style/style:paragraph-properties/@fo:text-indent and not($style/style:paragraph-properties/@style:auto-text-indent='true')">
            <xsl:value-of select="$style/style:paragraph-properties/@fo:text-indent"/>
          </xsl:when>
          <xsl:when
            test="$style/style:paragraph-properties/@style:auto-text-indent='true' and $style/style:text-properties/@fo:font-size">
            <xsl:variable name="fontSize">
              <xsl:call-template name="computeSize">
                <xsl:with-param name="node" select="$style/style:text-properties"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:value-of select="concat($fontSize div 2, 'pt')"/>
          </xsl:when>
          <xsl:when test="$style/ancestor::office:automatic-styles">
            <xsl:for-each select="document('styles.xml')">
              <xsl:choose>
                <xsl:when
                  test="key('styles', $styleName)/style:paragraph-properties/@fo:text-indent and not(key('styles', $styleName)/style:paragraph-properties/@style:auto-text-indent='true')">
                  <xsl:value-of
                    select="key('styles', $styleName)/style:paragraph-properties/@fo:text-indent"/>
                </xsl:when>
                <xsl:when
                  test="key('styles', $styleName)/style:paragraph-properties/@fo:auto-text-indent='true' and key('styles',$styleName)/style:text-properties/@fo:font-size">
                  <xsl:variable name="fontSize">
                    <xsl:call-template name="computeSize">
                      <xsl:with-param name="node"
                        select="key('styles', $styleName)[1]/style:text-properties"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:value-of select="concat($fontSize div 2, 'pt')"/>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:for-each>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>


  <!-- ignored -->
  <xsl:template match="text()" mode="styles"/>
  
  
  <!-- Climb style hierarchy for a property -->
  <xsl:template name="FindPropertyInStyleHierarchy">
    <xsl:param name="style-name"/>
    <xsl:param name="property-name"/>
    <xsl:param name="context">content.xml</xsl:param>
    
    <xsl:variable name="style-name-exists">
      <xsl:for-each select="document($context)">
        <xsl:value-of select="boolean(key('styles', $style-name))"/>
      </xsl:for-each>
    </xsl:variable>
    
    <xsl:choose>
      <xsl:when test="$style-name-exists = 'true' ">
        <xsl:for-each select="document($context)">
          <xsl:variable name="style" select="key('styles', $style-name)[1]"/>
          <xsl:choose>
            <xsl:when test="$style/style:paragraph-properties/@*[name() = $property-name]">
              <xsl:value-of select="$style/style:paragraph-properties/@*[name() = $property-name]"/>
            </xsl:when>
            <xsl:when test="$style/@style:parent-style-name">
              <xsl:if test="$style/@style:parent-style-name != $style-name">
                <xsl:call-template name="FindPropertyInStyleHierarchy">
                  <xsl:with-param name="style-name" select="$style/@style:parent-style-name"/>
                  <xsl:with-param name="property-name" select="$property-name"/>
                  <xsl:with-param name="context" select="$context"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:when>
            <xsl:otherwise></xsl:otherwise>
          </xsl:choose>
        </xsl:for-each>
      </xsl:when>
      <!-- switch the context, let's look into styles.xml -->
      <xsl:when test="$context != 'styles.xml'">
        <xsl:call-template name="FindPropertyInStyleHierarchy">
          <xsl:with-param name="style-name" select="$style-name"/>
          <xsl:with-param name="property-name" select="$property-name"/>
          <xsl:with-param name="context" select="'styles.xml'"/>
        </xsl:call-template>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  

</xsl:stylesheet>

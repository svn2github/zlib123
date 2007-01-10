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
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  exclude-result-prefixes="w r xlink number wp ">

  <xsl:import href="tables.xsl"/>
  <xsl:import href="lists.xsl"/>
  <xsl:import href="fonts.xsl"/>
  <xsl:import href="fields.xsl"/>
  <xsl:import href="footnotes.xsl"/>
  <xsl:import href="indexes.xsl"/>
  <xsl:import href="track.xsl"/>

  <xsl:import href="frames.xsl"/>
  <xsl:import href="sections.xsl"/>
  <xsl:import href="comments.xsl"/>

  <xsl:strip-space elements="*"/>
  <xsl:preserve-space elements="w:p"/>
  <xsl:preserve-space elements="w:r"/>

  <xsl:key name="InstrText" match="w:instrText" use="''"/>

  <!--main document-->
  <xsl:template name="content">
    <office:document-content>
      <office:scripts/>
      <office:font-face-decls>
        <xsl:apply-templates select="document('word/fontTable.xml')/w:fonts"/>
      </office:font-face-decls>
      <office:automatic-styles>
        <!-- automatic styles for document body -->
          <xsl:call-template name="InsertBodyStyles"/>
        <xsl:call-template name="InsertListStyles"/>
        <xsl:call-template name="InsertSectionsStyles"/>
          <xsl:call-template name="InsertFootnoteStyles"/>
      </office:automatic-styles>
      <office:body>
        <office:text>
          <xsl:call-template name="TrackChanges"/>
          <xsl:call-template name="InsertDocumentBody"/>
        </office:text>
      </office:body>
    </office:document-content>
  </xsl:template>

  <!--  generates automatic styles for sections-->
  <xsl:template name="InsertSectionsStyles">
    <xsl:if test="document('word/document.xml')/w:document/w:body/w:p/w:pPr/w:sectPr">
      <xsl:apply-templates
        select="document('word/document.xml')/w:document/w:body/w:p/w:pPr/w:sectPr"
        mode="automaticstyles"/>
    </xsl:if>
  </xsl:template>

  <!--  generates automatic styles for paragraphs  ho w does it exactly work ?? -->
  <xsl:template name="InsertBodyStyles">
    <xsl:apply-templates select="document('word/document.xml')/w:document/w:body"
      mode="automaticstyles"/>
  </xsl:template>

  <xsl:template name="InsertListStyles">
    <!-- document with lists-->
    <xsl:if
      test="document('word/document.xml')/w:document[descendant::w:numPr/w:numId] 
      or document('word/styles.xml')/w:styles/w:style[descendant::w:numPr/w:numId] ">
      <!-- automatic list styles with empty num format for elements which has non-existent w:num attached -->
      <xsl:apply-templates
        select="document('word/document.xml')/w:document/w:body/descendant::w:numId
        [not(document('word/numbering.xml')/w:numbering/w:num/@w:numId = @w:val)][1]"
        mode="automaticstyles"/>
      <!-- automatic list styles-->
      <xsl:apply-templates select="document('word/numbering.xml')/w:numbering/w:num"/>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertFootnoteStyles">
    <xsl:if
      test="document('word/footnotes.xml')/w:footnotes/w:footnote/w:p/w:r/w:rPr | 
      document('word/footnotes.xml')/w:footnotes/w:footnote/w:p/w:pPr">
      <xsl:apply-templates select="document('word/footnotes.xml')/w:footnotes/w:footnote/w:p"
        mode="automaticstyles"/>
    </xsl:if>
  </xsl:template>

  <!--  inserts document elements-->
  <xsl:template name="InsertDocumentBody">
    <xsl:choose>
      <xsl:when test="document('word/document.xml')/w:document/w:body/w:p/w:pPr/w:sectPr">
        <xsl:apply-templates select="document('word/document.xml')/w:document/w:body"
          mode="sections"/>
      </xsl:when>
    </xsl:choose>
    <xsl:apply-templates
      select="document('word/document.xml')/w:document/w:body/child::node()[not(following::w:p/w:pPr/w:sectPr) and not(descendant::w:sectPr)]"
    />
  </xsl:template>

  <!-- create a style for each paragraph. Do not take w:sectPr/w:rPr into consideration. -->
  <xsl:template
    match="w:pPr[parent::w:p]|w:r[parent::w:p[not(child::w:pPr)] and child::w:br[@w:type='page' or @w:type='column']]"
    mode="automaticstyles">
    <xsl:message terminate="no">progress:w:pPr</xsl:message>
    <style:style style:name="{generate-id(parent::w:p)}" style:family="paragraph">
      <xsl:call-template name="MasterPageName"/>
      <xsl:call-template name="InsertParagraphParentStyle"/>

      <style:paragraph-properties>
        <xsl:call-template name="InsertDefaultParagraphProperties"/>
        <xsl:call-template name="InsertParagraphProperties"/>
      </style:paragraph-properties>
    </style:style>
  </xsl:template>

  <xsl:template match="w:p[not (./w:pPr)]" mode="automaticstyles">
    <xsl:if test="document('word/styles.xml')/w:styles/w:docDefaults/w:pPrDefault/w:pPr">
      <style:style style:name="{generate-id(.)}" style:family="paragraph">
        <xsl:call-template name="MasterPageName"/>
        <style:paragraph-properties>
          <xsl:for-each
            select="document('word/styles.xml')/w:styles/w:docDefaults/w:pPrDefault/w:pPr">
            <xsl:call-template name="InsertParagraphProperties"/>
          </xsl:for-each>
        </style:paragraph-properties>
      </style:style>
    </xsl:if>
    <xsl:apply-templates mode="automaticstyles"/>
  </xsl:template>

  <!--  when paragraph has no parent style it should be set to Normal style which contains all default paragraph properties -->
  <xsl:template name="InsertParagraphParentStyle">
    <xsl:choose>
      <xsl:when test="w:pStyle">
        <xsl:attribute name="style:parent-style-name">
          <xsl:choose>
            <xsl:when test="contains(w:pStyle/@w:val,'TOC')">
              <xsl:value-of select="concat('Contents_20_',substring-after(w:pStyle/@w:val,'TOC'))"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="w:pStyle/@w:val"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="style:parent-style-name">
          <xsl:text>Normal</xsl:text>
        </xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- create a style for each run. Do not take w:pPr/w:rPr into consideration. Ignore runs with no properties. -->
  <xsl:template match="w:rPr[parent::w:r and not(count(child::node())=1 and child::w:noProof)]" mode="automaticstyles">
    <xsl:message terminate="no">progress:w:rPr</xsl:message>
    <style:style style:name="{generate-id(parent::w:r)}" style:family="text">
      <xsl:choose>
        <xsl:when test="w:rStyle">
          <xsl:attribute name="style:parent-style-name">
            <xsl:value-of select="w:rStyle/@w:val"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="style:parent-style-name">
            <xsl:text>Normal</xsl:text>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
      <style:text-properties>
        <xsl:call-template name="InsertTextProperties"/>
      </style:text-properties>
    </style:style>
  </xsl:template>

  <!-- ignore text in automatic styles mode. -->
  <xsl:template match="w:t" mode="automaticstyles">
    <!-- do nothing. -->
  </xsl:template>

  <!--  get outline level from styles hierarchy for headings -->
  <xsl:template name="GetStyleOutlineLevel">
    <xsl:param name="outline"/>
    <xsl:for-each select="document('word/styles.xml')">
      <xsl:variable name="basedOn">
        <xsl:value-of select="key('StyleId', $outline)[1]/w:basedOn/@w:val"/>
      </xsl:variable>
      <!--  Search outlineLvl in style based on style -->
      <xsl:choose>
        <xsl:when test="key('StyleId', $basedOn)[1]/w:pPr/w:outlineLvl/@w:val">
          <xsl:value-of select="key('StyleId', $basedOn)[1]/w:pPr/w:outlineLvl/@w:val"/>
        </xsl:when>
        <xsl:when test="key('StyleId', $outline)[1]/w:basedOn/@w:val">
          <xsl:call-template name="GetStyleOutlineLevel">
            <xsl:with-param name="outline">
              <xsl:value-of select="key('StyleId', $outline)[1]/w:basedOn/@w:val"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <!-- Get outlineLvl if the paragraf is heading -->
  <xsl:template name="GetOutlineLevel">
    <xsl:param name="node"/>
    <xsl:for-each select="document('word/styles.xml')">
      <xsl:choose>
        <xsl:when test="$node/w:pPr/w:outlineLvl/@w:val">
          <xsl:value-of select="$node/w:pPr/w:outlineLvl/@w:val"/>
        </xsl:when>
        <xsl:when test="$node/w:pPr/w:pStyle/@w:val">
          <xsl:variable name="outline">
            <xsl:value-of select="$node/w:pPr/w:pStyle/@w:val"/>
          </xsl:variable>
          <!--Search outlineLvl in styles.xml  -->
          <xsl:choose>
            <xsl:when test="key('StyleId', $outline)[1]/w:pPr/w:outlineLvl/@w:val">
              <xsl:value-of select="key('StyleId', $outline)[1]/w:pPr/w:outlineLvl/@w:val"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="GetStyleOutlineLevel">
                <xsl:with-param name="outline">
                  <xsl:value-of select="$node/w:pPr/w:pStyle/@w:val"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:variable name="outline">
            <xsl:value-of select="$node/w:r/w:rPr/w:rStyle/@w:val"/>
          </xsl:variable>
          <xsl:variable name="linkedStyleOutline">
            <xsl:value-of select="key('StyleId', $outline)[1]/w:link/@w:val"/>
          </xsl:variable>
          <!--if outlineLvl is not defined search in parent style by w:link-->
          <xsl:choose>
            <xsl:when test="key('StyleId', $linkedStyleOutline)[1]/w:pPr/w:outlineLvl/@w:val">
              <xsl:value-of
                select="key('StyleId', $linkedStyleOutline)[1]/w:pPr/w:outlineLvl/@w:val"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="GetStyleOutlineLevel">
                <xsl:with-param name="outline">
                  <xsl:value-of select="key('StyleId', $outline)[1]/w:link/@w:val"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <!--  paragraphs, lists, headings-->
  <xsl:template match="w:p">
    <xsl:message terminate="no">progress:w:p</xsl:message>

    <xsl:variable name="outlineLevel">
      <xsl:call-template name="GetOutlineLevel">
        <xsl:with-param name="node" select="."/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="numId">
      <xsl:call-template name="GetListProperty">
        <xsl:with-param name="node" select="."/>
        <xsl:with-param name="property">w:numId</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="level">
      <xsl:call-template name="GetListProperty">
        <xsl:with-param name="node" select="."/>
        <xsl:with-param name="property">w:ilvl</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:choose>

      <!--check if the paragraph starts a table-of content or Bibliography-->
      <xsl:when
        test="descendant::w:r[contains(w:instrText,'TOC') or contains(w:instrText,'BIBLIOGRAPHY')]">
        <xsl:apply-templates select="." mode="tocstart"/>
      </xsl:when>

      <!-- check if the pargraph is Citations -->

      <xsl:when test="descendant::w:r[contains(w:instrText,'CITATION')]">
        <xsl:variable name="InstrText">
          <xsl:value-of select="descendant::w:r/w:instrText"/>
        </xsl:variable>
        <text:p>
          <xsl:call-template name="TextBibliographyMark">
            <xsl:with-param name="TextIdentifier">
              <xsl:value-of
                select="substring-before(substring-after($InstrText, 'CITATION '), ' \')"/>
            </xsl:with-param>
          </xsl:call-template>
          <xsl:apply-templates select="following::w:p[1]/child::node()"/>
        </text:p>
      </xsl:when>

      <xsl:when test="preceding::w:p[1]/w:pPr/w:rPr/w:del"/>



      <!--  check if the paragraf is list element (it can be a heading also) -->
      <xsl:when test="$numId != '' and $level &lt; 10">
        <xsl:apply-templates select="." mode="list">
          <xsl:with-param name="numId" select="$numId"/>
          <xsl:with-param name="level" select="$level"/>
        </xsl:apply-templates>
      </xsl:when>

      <!--  check if the paragraf is heading -->
      <xsl:when test="$outlineLevel != ''">
        <xsl:apply-templates select="." mode="heading">
          <xsl:with-param name="outlineLevel">
            <xsl:value-of select="$outlineLevel"/>
          </xsl:with-param>
        </xsl:apply-templates>
      </xsl:when>

      <!--  default scenario - paragraph-->
      <xsl:otherwise>
        <xsl:apply-templates select="." mode="paragraph"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--paragraph with outline level is a heading-->
  <xsl:template match="w:p" mode="heading">
    <xsl:param name="outlineLevel"/>
    <xsl:variable name="numId">
      <xsl:call-template name="GetListProperty">
        <xsl:with-param name="node" select="."/>
        <xsl:with-param name="property">w:numId</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <text:h>
      <!--style name-->
      <xsl:if test="w:pPr or w:r/w:br[@w:type='page' or @w:type='column']">
        <xsl:attribute name="text:style-name">
          <xsl:value-of select="generate-id(self::node())"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="preceding::w:p[1]/w:pPr/w:rPr/w:ins and $numId!=''">
        <text:change-end>
          <xsl:attribute name="text:change-id">
            <xsl:value-of select="generate-id(preceding::w:p[1])"/>
          </xsl:attribute>
        </text:change-end>
      </xsl:if>
      <!--header outline level -->
      <xsl:call-template name="InsertHeadingOutlineLvl">
        <xsl:with-param name="outlineLevel" select="$outlineLevel"/>
      </xsl:call-template>
      <!-- unnumbered heading is list header  -->
      <xsl:call-template name="InsertHeadingAsListHeader"/>
      <xsl:apply-templates/>
      <xsl:if test="w:pPr/w:rPr/w:del">
        <!--      if this following paragraph is attached to this one in track changes mode-->
        <xsl:call-template name="InsertDeletedParagraph"/>
      </xsl:if>
      <xsl:if test="w:pPr/w:rPr/w:ins">
        <text:change-start>
          <xsl:attribute name="text:change-id">
            <xsl:value-of select="generate-id(self::node())"/>
          </xsl:attribute>
        </text:change-start>
      </xsl:if>
    </text:h>
    <xsl:if test="w:pPr/w:rPr/w:ins and $numId=''">
      <text:change-end>
        <xsl:attribute name="text:change-id">
          <xsl:value-of select="generate-id(self::node())"/>
        </xsl:attribute>
      </text:change-end>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertDeletedParagraph">
    <text:change>
      <xsl:attribute name="text:change-id">
        <xsl:value-of select="generate-id(w:pPr/w:rPr)"/>
      </xsl:attribute>
    </text:change>
    <xsl:apply-templates select="following::w:p[1]/child::node()"/>
    <xsl:if test="following::w:p[1]/w:pPr/w:rPr/w:del">
      <xsl:for-each select="following::w:p[1]">
        <xsl:call-template name="InsertDeletedParagraph"/>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertHeadingAsListHeader">
    <xsl:variable name="numId">
      <xsl:call-template name="GetListProperty">
        <xsl:with-param name="node" select="."/>
        <xsl:with-param name="property">w:numId</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:if test="$numId = ''">
      <xsl:attribute name="text:is-list-header">
        <xsl:text>true</xsl:text>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertHeadingOutlineLvl">
    <xsl:param name="outlineLevel"/>
    <xsl:attribute name="text:outline-level">
      <xsl:variable name="headingLvl">
        <xsl:call-template name="GetListProperty">
          <xsl:with-param name="node" select="."/>
          <xsl:with-param name="property">w:ilvl</xsl:with-param>
        </xsl:call-template>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$headingLvl != ''">
          <xsl:value-of select="$headingLvl+1"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$outlineLevel+1"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>

  <!--  paragraphs-->
  <xsl:template match="w:p" mode="paragraph">
    <xsl:choose>
      <!--avoid nested paragaraphs-->
      <xsl:when test="parent::w:p">
        <xsl:apply-templates select="child::node()"/>
      </xsl:when>
      <!--default scenario-->
      <xsl:otherwise>
        <xsl:variable name="numId">
          <xsl:call-template name="GetListProperty">
            <xsl:with-param name="node" select="."/>
            <xsl:with-param name="property">w:numId</xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <text:p>
          <xsl:if
            test="w:pPr or w:r/w:br[@w:type='page' or @w:type='column'] or document('word/styles.xml')/w:styles/w:docDefaults/w:pPrDefault/w:pPr">
            <xsl:attribute name="text:style-name">
              <xsl:value-of select="generate-id(self::node())"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="preceding::w:p[1]/w:pPr/w:rPr/w:ins and $numId!=''">
            <text:change-end>
              <xsl:attribute name="text:change-id">
                <xsl:value-of select="generate-id(preceding::w:p[1])"/>
              </xsl:attribute>
            </text:change-end>
          </xsl:if>
          <xsl:apply-templates/>
          <!--      if this following paragraph is attached to this one in track changes mode-->
          <xsl:if test="w:pPr/w:rPr/w:del">
            <xsl:call-template name="InsertDeletedParagraph"/>
          </xsl:if>
          <xsl:if test="w:pPr/w:rPr/w:ins">
            <text:change-start>
              <xsl:attribute name="text:change-id">
                <xsl:value-of select="generate-id(self::node())"/>
              </xsl:attribute>
            </text:change-start>
          </xsl:if>
        </text:p>
        <xsl:if test="w:pPr/w:rPr/w:ins and $numId=''">
          <text:change-end>
            <xsl:attribute name="text:change-id">
              <xsl:value-of select="generate-id(self::node())"/>
            </xsl:attribute>
          </text:change-end>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--tabs-->
  <xsl:template match="w:tab[not(parent::w:tabs)]">
    <text:tab/>
  </xsl:template>

  <!-- run -->
  <xsl:template match="w:r">
    <xsl:message terminate="no">progress:w:r</xsl:message>
    <xsl:choose>

      <!--fields-->
      <xsl:when test="preceding::w:fldChar[1][@w:fldCharType='begin' or @w:fldCharType='separate']">
        <xsl:call-template name="InsertField"/>
      </xsl:when>

      <!-- Comments -->
      <xsl:when test="w:commentReference/@w:id">
        <xsl:call-template name="InsertComment">
          <xsl:with-param name="Id">
            <xsl:value-of select="w:commentReference/@w:id"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>

      <!--  Track changes    -->
      <xsl:when test="parent::w:del">
        <xsl:call-template name="TrackChangesDeleteMade"/>
      </xsl:when>
      <xsl:when test="parent::w:ins">
        <xsl:call-template name="TrackChangesInsertMade"/>
      </xsl:when>
      <xsl:when test="descendant::w:rPrChange">
        <xsl:call-template name="TrackChangesChangesMade"/>
      </xsl:when>

      <!-- attach automatic style-->
      <xsl:when test="w:rPr[not(count(child::node())=1 and child::w:noProof)]">
        <text:span text:style-name="{generate-id(self::node())}">
          <xsl:apply-templates/>
        </text:span>
      </xsl:when>

      <!--default scenario-->
      <xsl:otherwise>
        <xsl:apply-templates/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- path for hyperlinks-->
  <xsl:template name="GetLinkPath">
    <xsl:param name="linkHref"/>
    <xsl:choose>
      <xsl:when
        test="contains($linkHref, 'file:///') or contains($linkHref, 'http://') or contains($linkHref, 'mailto:')">
        <xsl:value-of select="$linkHref"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="concat('../',$linkHref)"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- hyperlinks - w:hyperlink and fieldchar types-->
  <xsl:template name="InsertHyperlink">
    <text:a xlink:type="simple">
      <xsl:attribute name="xlink:href">
        <!-- document hyperlink -->
        <xsl:if test="@w:anchor">
          <xsl:value-of select="concat('#',@w:anchor)"/>
        </xsl:if>
        <!-- file or web page hyperlink with relationship id -->
        <xsl:if test="@r:id">
          <xsl:variable name="relationshipId">
            <xsl:value-of select="@r:id"/>
          </xsl:variable>
          <xsl:variable name="document">
            <xsl:call-template name="GetDocumentName">
              <xsl:with-param name="rootId">
                <xsl:value-of select="generate-id(/node())"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:for-each
            select="document(concat('word/_rels/',$document,'.rels'))//node()[name() = 'Relationship']">
            <xsl:if test="./@Id=$relationshipId">
              <xsl:call-template name="GetLinkPath">
                <xsl:with-param name="linkHref" select="@Target"/>
              </xsl:call-template>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
        <!-- file or web page hyperlink -  fieldchar type (can contain several paragraphs in Word) -->
        <xsl:if test="self::w:r">
          <xsl:call-template name="GetLinkPath">
            <xsl:with-param name="linkHref">
              <xsl:value-of
                select="substring-before(substring-after(preceding::w:instrText[1],'&quot;'),'&quot;')"
              />
            </xsl:with-param>
          </xsl:call-template>
        </xsl:if>
      </xsl:attribute>
      <!--hyperlink text-->
      <xsl:apply-templates/>
    </text:a>
  </xsl:template>

  <!-- hyperlinks -->
  <xsl:template match="w:hyperlink">
    <xsl:call-template name="InsertHyperlink"/>
  </xsl:template>
  <!--  text bookmark mark -->
  <xsl:template match="w:bookmarkStart">
    <text:bookmark-start>
      <xsl:attribute name="text:name">
        <xsl:value-of select="@w:name"/>
      </xsl:attribute>
    </text:bookmark-start>
    <text:bookmark-end>
      <xsl:attribute name="text:name">
        <xsl:value-of select="@w:name"/>
      </xsl:attribute>
    </text:bookmark-end>
  </xsl:template>

  <!-- simple text  -->
  <xsl:template match="w:t">
    <xsl:message terminate="no">progress:w:t</xsl:message>
    <xsl:choose>
      <!--check whether string contains  whitespace sequence-->
      <xsl:when test="not(contains(.,'  '))">
        <xsl:value-of select="."/>
      </xsl:when>
      <xsl:otherwise>
        <!--converts whitespaces sequence to text:s-->
        <xsl:call-template name="InsertWhiteSpaces"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--  convert multiple white spaces  -->
  <xsl:template name="InsertWhiteSpaces">
    <xsl:param name="string" select="."/>
    <xsl:param name="length" select="string-length(.)"/>
    <!-- string which doesn't contain whitespaces-->
    <xsl:choose>
      <xsl:when test="not(contains($string,' '))">
        <xsl:value-of select="$string"/>
      </xsl:when>
      <!-- convert white spaces  -->
      <xsl:otherwise>
        <xsl:variable name="before">
          <xsl:value-of select="substring-before($string,' ')"/>
        </xsl:variable>
        <xsl:variable name="after">
          <xsl:call-template name="CutStartSpaces">
            <xsl:with-param name="cuted">
              <xsl:value-of select="substring-after($string,' ')"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:if test="$before != '' ">
          <xsl:value-of select="concat($before,' ')"/>
        </xsl:if>
        <!--add remaining whitespaces as text:s if there are any-->
        <xsl:if test="string-length(concat($before,' ', $after)) &lt; $length ">
          <xsl:choose>
            <xsl:when test="($length - string-length(concat($before, $after))) = 1">
              <text:s/>
            </xsl:when>
            <xsl:otherwise>
              <text:s>
                <xsl:attribute name="text:c">
                  <xsl:choose>
                    <xsl:when test="$before = ''">
                      <xsl:value-of select="$length - string-length($after)"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$length - string-length(concat($before,' ', $after))"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
              </text:s>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
        <!--repeat it for substring which has whitespaces-->
        <xsl:if test="contains($string,' ') and $length &gt; 0">
          <xsl:call-template name="InsertWhiteSpaces">
            <xsl:with-param name="string">
              <xsl:value-of select="$after"/>
            </xsl:with-param>
            <xsl:with-param name="length">
              <xsl:value-of select="string-length($after)"/>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--  cut start spaces -->
  <xsl:template name="CutStartSpaces">
    <xsl:param name="cuted"/>
    <xsl:choose>
      <xsl:when test="starts-with($cuted,' ')">
        <xsl:call-template name="CutStartSpaces">
          <xsl:with-param name="cuted">
            <xsl:value-of select="substring-after($cuted,' ')"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$cuted"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- line breaks  -->
  <xsl:template match="w:br[not(@w:type) or @w:type!='page' and @w:type!='column']">
    <text:line-break/>
  </xsl:template>

  <!-- page or column break must have style defined in paragraph -->
  <xsl:template match="w:br[@w:type='page' or @w:type='column']" mode="automaticstyles">
    <xsl:if test="not(ancestor::w:p/w:pPr)">
      <style:style style:name="{generate-id(ancestor::w:p)}" style:family="paragraph">
        <xsl:call-template name="MasterPageName"/>
        <style:paragraph-properties>
          <xsl:call-template name="InsertParagraphProperties"/>
        </style:paragraph-properties>
      </style:style>
    </xsl:if>
  </xsl:template>

  <!--ignore text in automatic styles mode-->
  <xsl:template match="text()" mode="automaticstyles"/>

  <xsl:template name="MasterPageName">
    <xsl:if test="not(preceding::w:p)">
      <xsl:choose>
        <xsl:when test="following::w:p/descendant::w:sectPr">
          <xsl:choose>
            <xsl:when test="following::w:p/descendant::w:sectPr/w:titlePg">
              <xsl:attribute name="style:master-page-name">
                <xsl:value-of
                  select="concat('First_H_',generate-id(following::w:p/descendant::w:sectPr))"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="style:master-page-name">
                <xsl:value-of select="concat('H_',generate-id(following::w:p/descendant::w:sectPr))"
                />
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:choose>
            <xsl:when test="document('word/document.xml')/w:document/w:body/w:sectPr/w:titlePg">
              <xsl:attribute name="style:master-page-name">
                <xsl:text>First_Page</xsl:text>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="style:parent-style-name">
                <xsl:text>Standard</xsl:text>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <xsl:if test="preceding::w:p[parent::w:body][1]/descendant::w:sectPr">
      <xsl:if
        test="(preceding::w:sectPr[1]/w:pgSz/@w:w != following::w:sectPr[1]/w:pgSz/@w:w or preceding::w:sectPr[1]/w:pgSz/@w:h != following::w:sectPr[1]/w:pgSz/@w:h or preceding::w:sectPr[1]/w:pgSz/@w:orient != following::w:sectPr[1]/w:pgSz/@w:orient or document('word/document.xml')/w:document/w:body/w:sectPr/w:pgSz/@w:w != following::w:sectPr[1]/w:pgSz/@w:w or document('word/document.xml')/w:document/w:body/w:sectPr/w:pgSz/@w:h != ./w:pgSz/@w:h or document('word/document.xml')/w:document/w:body/w:sectPr/w:pgSz/@w:orient != following::w:sectPr[1]/w:pgSz/@w:orient) and not(following::w:sectPr[1]/w:headerReference) and not(following::w:sectPr[1]/w:footerReference)">
        <xsl:attribute name="style:master-page-name">
          <xsl:value-of select="concat('PAGE_',generate-id(.))"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:choose>
        <xsl:when test="following::w:p/descendant::w:sectPr">
          <xsl:choose>
            <xsl:when test="following::w:p/descendant::w:sectPr/w:titlePg">
              <xsl:attribute name="style:master-page-name">
                <xsl:value-of select="concat('First_H_',generate-id(following::w:sectPr))"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if
                test="(preceding::w:sectPr[1]/w:pgSz/@w:w != following::w:sectPr[1]/w:pgSz/@w:w
                or preceding::w:sectPr[1]/w:pgSz/@w:h != following::w:sectPr[1]/w:pgSz/@w:h
                or preceding::w:sectPr[1]/w:pgSz/@w:orient != following::w:sectPr[1]/w:pgSz/@w:orient 
                or document('word/document.xml')/w:document/w:body/w:sectPr/w:pgSz/@w:w != following::w:sectPr[1]/w:pgSz/@w:w
                or document('word/document.xml')/w:document/w:body/w:sectPr/w:pgSz/@w:h != ./w:pgSz/@w:h
                or document('word/document.xml')/w:document/w:body/w:sectPr/w:pgSz/@w:orient != following::w:sectPr[1]/w:pgSz/@w:orient) 
                or following::w:sectPr[1]/w:headerReference 
                or following::w:sectPr[1]/w:footerReference">
                <xsl:attribute name="style:master-page-name">
                  <xsl:value-of select="concat('H_',generate-id(following::w:sectPr))"/>
                </xsl:attribute>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:choose>
            <xsl:when test="following::w:sectPr/w:titlePg">
              <xsl:attribute name="style:master-page-name">
                <xsl:text>First_Page</xsl:text>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if
                test="(preceding::w:sectPr[1]/w:pgSz/@w:w != following::w:sectPr[1]/w:pgSz/@w:w 
                or preceding::w:sectPr[1]/w:pgSz/@w:h != following::w:sectPr[1]/w:pgSz/@w:h 
                or preceding::w:sectPr[1]/w:pgSz/@w:orient != following::w:sectPr[1]/w:pgSz/@w:orient 
                or following::w:sectPr[1]/w:headerReference 
                or following::w:sectPr[1]/w:footerReference) or
                following::w:sectPr[1]/w:cols">
                <xsl:attribute name="style:master-page-name">
                  <xsl:text>Standard</xsl:text>
                </xsl:attribute>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
  </xsl:template>
</xsl:stylesheet>

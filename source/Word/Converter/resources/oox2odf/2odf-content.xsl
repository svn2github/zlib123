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
  xmlns:pchar="urn:cleverage:xmlns:post-processings:characters"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
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
  xmlns:pcut="urn:cleverage:xmlns:post-processings:pcut"
  xmlns:v="urn:schemas-microsoft-com:vml" 
  xmlns:oox="urn:oox"
  exclude-result-prefixes="w r xlink number wp oox">

  <xsl:import href="2odf-tables.xsl"/>
  <xsl:import href="2odf-lists.xsl"/>
  <xsl:import href="2odf-fonts.xsl"/>
  <xsl:import href="2odf-fields.xsl"/>
  <xsl:import href="2odf-footnotes.xsl"/>
  <xsl:import href="2odf-indexes.xsl"/>
  <xsl:import href="2odf-track.xsl"/>
  <xsl:import href="2odf-frames.xsl"/>
  <xsl:import href="2odf-sections.xsl"/>
  <!--xsl:import href="2odf-comments.xsl"/-->

  <xsl:strip-space elements="*"/>
  <xsl:preserve-space elements="w:p"/>
  <xsl:preserve-space elements="w:r"/>

  <xsl:key name="InstrText" match="w:instrText" use="''"/>
  <xsl:key name="bookmarkStart" match="w:bookmarkStart" use="@w:id"/>
  <xsl:key name="pPr" match="w:pPr" use="''"/>
  <xsl:key name="sectPr" match="w:sectPr" use="''"/>
  
  <xsl:key name="p" match="w:p" use="@oox:id" />
  <xsl:key name="sectPr" match="w:sectPr" use="@oox:s" />

  <!--main document-->
  <xsl:template name="content">
    <office:document-content>
      <office:scripts/>
      <office:font-face-decls>
        <xsl:apply-templates select="key('Part', 'word/fontTable.xml')/w:fonts"/>
      </office:font-face-decls>
      <office:automatic-styles>
        <!-- automatic styles for document body -->
        <xsl:call-template name="InsertBodyStyles"/>
        <xsl:call-template name="InsertListStyles"/>
        <xsl:call-template name="InsertSectionsStyles"/>
        <xsl:call-template name="InsertFootnoteStyles"/>
        <xsl:call-template name="InsertEndnoteStyles"/>
        <xsl:call-template name="InsertFrameStyle"/>
      </office:automatic-styles>
      <office:body>
        <office:text>
          <xsl:call-template name="InsertUserFieldDecls" />
          <xsl:call-template name="TrackChanges"/>
          <xsl:call-template name="InsertDocumentBody"/>
        </office:text>
      </office:body>
    </office:document-content>
  </xsl:template>

  <!--  generates automatic styles for frames -->
  <xsl:template name="InsertFrameStyle">
    <!-- when w:pict is child of paragraph-->
    <xsl:if test="key('Part', 'word/document.xml')/w:document/w:body/w:p/w:r/w:pict">
      <xsl:apply-templates select="key('Part', 'word/document.xml')/w:document/w:body/w:p/w:r/w:pict"
        mode="automaticpict"/>
    </xsl:if>

    <!-- when w:pict is child of a cell-->
    <xsl:if test="key('Part', 'word/document.xml')/w:document//w:body/w:tbl/w:tr/w:tc/w:p/w:r/w:pict">
      <xsl:apply-templates select="key('Part', 'word/document.xml')/w:document//w:body/w:tbl/w:tr/w:tc/w:p/w:r/w:pict"
        mode="automaticpict"/>
    </xsl:if>
  </xsl:template>
  
  <!--
  Summary: Inserst the declarations of all SET fields
  Author: makz (DIaLOGIKa)
  Date: 2.11.2007
  -->
  <xsl:template name="InsertUserFieldDecls">
    <text:user-field-decls>
      <xsl:apply-templates select="key('Part', 'word/document.xml')/w:document/w:body//w:instrText" mode="UserFieldDecls" />
    </text:user-field-decls>
  </xsl:template>

  <!--  generates automatic styles for sections-->
  <xsl:template name="InsertSectionsStyles">
    <xsl:if test="key('Part', 'word/document.xml')/w:document/w:body/w:p/w:pPr/w:sectPr">
      <xsl:apply-templates
        select="key('Part', 'word/document.xml')/w:document/w:body/w:p/w:pPr/w:sectPr"
        mode="automaticstyles"/>
    </xsl:if>
  </xsl:template>

  <!--  generates automatic styles for paragraphs  ho w does it exactly work ?? -->
  <xsl:template name="InsertBodyStyles">
    <xsl:apply-templates select="key('Part', 'word/document.xml')/w:document/w:body"
      mode="automaticstyles"/>
  </xsl:template>

  <xsl:template name="InsertListStyles">
    <!-- document with lists-->
    <xsl:for-each select="key('Part', 'word/document.xml')">
      <xsl:choose>
        <xsl:when test="key('pPr', '')/w:numPr/w:numId">
          <!-- automatic list styles with empty num format for elements which has non-existent w:num attached -->
          <xsl:apply-templates
            select="key('pPr', '')/w:numPr/w:numId[not(key('Part', 'word/numbering.xml')/w:numbering/w:num/@w:numId = @w:val)][1]"
            mode="automaticstyles"/>
          <!-- automatic list styles-->
          <xsl:apply-templates select="key('Part', 'word/numbering.xml')/w:numbering/w:num"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="key('Part', 'word/styles.xml')">
            <xsl:if test="key('pPr', '')/w:numPr/w:numId">
              <!-- automatic list styles-->
              <xsl:apply-templates select="key('Part', 'word/numbering.xml')/w:numbering/w:num"/>
            </xsl:if>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="InsertFootnoteStyles">
    <xsl:if
      test="key('Part', 'word/footnotes.xml')/w:footnotes/w:footnote/w:p/w:r/w:rPr | 
      key('Part', 'word/footnotes.xml')/w:footnotes/w:footnote/w:p/w:pPr">
      <xsl:apply-templates select="key('Part', 'word/footnotes.xml')/w:footnotes/w:footnote/w:p"
        mode="automaticstyles"/>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertEndnoteStyles">
    <xsl:if
      test="key('Part', 'word/endnotes.xml')/w:endnotes/w:endnote/w:p/w:r/w:rPr | 
      key('Part', 'word/footnotes.xml')/w:footnotes/w:footnote/w:p/w:pPr">
      <!-- divo: TODO check for copy and paste error in above test. shouldn't it be endnotes.xml twice??? -->
      <xsl:apply-templates select="key('Part', 'word/endnotes.xml')/w:endnotes/w:endnote/w:p"
        mode="automaticstyles"/>
    </xsl:if>
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
            <xsl:when test="w:pStyle/@w:val='FootnoteText'">
              <xsl:text>Footnote_20_Symbol</xsl:text>
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

  <!--  inserts document elements-->
  <xsl:template name="InsertDocumentBody">
    <xsl:choose>
      <xsl:when test="key('Part', 'word/document.xml')/w:document/w:body/w:p/w:pPr/w:sectPr">
        <xsl:apply-templates select="key('Part', 'word/document.xml')/w:document/w:body"
          mode="sections"/>
      </xsl:when>
      <!-- divo: what is the else case here??? -->
    </xsl:choose>
    <xsl:apply-templates
      select="key('Part', 'word/document.xml')/w:document/w:body/child::node()[not(following::w:p/w:pPr/w:sectPr) and not(descendant::w:sectPr)]"
    />
  </xsl:template>

  <!-- create a style for each paragraph. Do not take w:sectPr/w:rPr into consideration. -->
  <xsl:template
    match="w:pPr[parent::w:p]|w:r[parent::w:p[not(child::w:pPr)] and (child::w:br[@w:type='page' or @w:type='column'] or contains(child::w:pict/v:shape/@style,'mso-position-horizontal-relative:char'))]"
    mode="automaticstyles">
    <xsl:message terminate="no">progress:w:pPr</xsl:message>
    <style:style style:name="{generate-id(parent::w:p)}" style:family="paragraph">
      <xsl:call-template name="InsertParagraphParentStyle"/>
      <xsl:call-template name="MasterPageName"/>

      <style:paragraph-properties>
        <xsl:call-template name="InsertDefaultTabStop"/>
        <xsl:call-template name="InsertParagraphProperties"/>
      </style:paragraph-properties>
      <!-- add text-properties to empty paragraphs. -->
      <xsl:if test="parent::w:p[count(child::node()) = 1]/w:pPr/w:rPr">
        <style:text-properties>
          <xsl:call-template name="InsertTextProperties"/>
        </style:text-properties>
      </xsl:if>
    </style:style>
  </xsl:template>

  <xsl:template
    match="w:p[not(./w:pPr) and not(w:r/w:br[@w:type='page' or @w:type='column']) and not(descendant::w:pict)]"
    mode="automaticstyles">
    <xsl:if test="key('Part', 'word/styles.xml')/w:styles/w:docDefaults/w:pPrDefault">
      <style:style style:name="{generate-id(.)}" style:family="paragraph">
        <xsl:call-template name="MasterPageName"/>
        <xsl:call-template name="InsertDefaultParagraphProperties"/>
      </style:style>
    </xsl:if>
    <xsl:apply-templates mode="automaticstyles"/>
  </xsl:template>

  <!-- create a style for each run. Do not take w:pPr/w:rPr into consideration. Ignore runs with no properties. -->
  <xsl:template match="w:rPr[parent::w:r and not(count(child::node())=1 and child::w:noProof)]"
    mode="automaticstyles">
    <xsl:message terminate="no">progress:w:rPr</xsl:message>
    <style:style style:name="{generate-id(parent::w:r)}" style:family="text">
      <xsl:if test="w:rStyle">
        <xsl:attribute name="style:parent-style-name">
          <!--clam bugfix #1806204-->
          <xsl:choose>
            <xsl:when test="../../preceding::w:r[contains(w:instrText,'TOC')] and w:rStyle/@w:val='Hyperlink'">X3AS7TOCHyperlink</xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="w:rStyle/@w:val"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
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
    <xsl:for-each select="key('Part', 'word/styles.xml')">
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
    <xsl:for-each select="key('Part', 'word/styles.xml')">
      <xsl:choose>
        <xsl:when test="$node/w:pPr/w:pStyle/@w:val">
          <xsl:variable name="outline">
            <xsl:value-of select="$node/w:pPr/w:pStyle/@w:val"/>
          </xsl:variable>
          <!--Search outlineLvl in styles.xml  -->
          <xsl:choose>
            <xsl:when test="key('StyleId', $outline)[1][w:pPr/w:outlineLvl/@w:val and w:basedOn/@w:val='Normal']">
              <xsl:value-of select="key('StyleId', $outline)[1]/w:pPr/w:outlineLvl/@w:val"/>
            </xsl:when>
            <xsl:otherwise/>
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

  <!-- get outlineLvl if the paragraph's ancestor is heading -->
  <xsl:template name="GetDummyOutlineLevel">
    <xsl:param name="node"/>
    <xsl:for-each select="key('Part', 'word/styles.xml')">
      <xsl:choose>
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


  <!--math, dialogika: added for bugfix #1802258 BEGIN -->  
  <!--Checks recursively whether given style is default heading (must start with Counter = 1)-->
  <xsl:template name ="CheckDefaultHeading">
    <xsl:param name="Name" />
    <xsl:param name="Counter" select="1"/>

    <xsl:choose>
      <xsl:when test="$Counter &gt; 9" >false</xsl:when>
      <xsl:when test="concat('heading ',$Counter) = $Name">true</xsl:when>
      <!--math, dialogika: added for bugfix #1792424 BEGIN -->
      <xsl:when test="concat('Heading ',$Counter) = $Name">true</xsl:when>
      <!--math, dialogika: added for bugfix #1792424 END -->
      <xsl:otherwise>
        <xsl:call-template name="CheckDefaultHeading">
          <xsl:with-param name="Name" select="$Name" />
          <xsl:with-param name="Counter">
            <xsl:value-of select="$Counter + 1" />
          </xsl:with-param>
        </xsl:call-template>        
      </xsl:otherwise>      
    </xsl:choose>
  </xsl:template>
  <!--math, dialogika: added for bugfix #1802258 END -->


  <!--  paragraphs, lists, headings-->
  <xsl:template match="w:p">
    <xsl:message terminate="no">progress:w:p</xsl:message>

    <xsl:variable name="outlineLevel">
      <xsl:call-template name="GetOutlineLevel">
        <xsl:with-param name="node" select="."/>
      </xsl:call-template>
    </xsl:variable>

    <xsl:variable name="dummyOutlineLevel">
      <xsl:call-template name="GetDummyOutlineLevel">
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

    <xsl:variable name="styleId" select="w:pPr/w:pStyle/@w:val"/>

    <xsl:variable name="ifNormal">
      <xsl:for-each select="key('Part', 'word/styles.xml')">
        <xsl:if test="key('StyleId', $styleId)/w:basedOn/@w:val='Normal'">true</xsl:if>
      </xsl:for-each>
    </xsl:variable>

    <!-- check if preceding paragraph is a heading in order not to lose numbered paragraphs inside headings-->
    <xsl:variable name="isPrecedingHeading">
      <xsl:if test="$numId!=''">
        <xsl:for-each select="key('p', number(@oox:id)-1)">
          <xsl:variable name="precedingOutlineLevel">
            <xsl:call-template name="GetOutlineLevel">
              <xsl:with-param name="node" select="."/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name="precedingStyleId" select="w:pPr/w:pStyle/@w:val"/>

          <xsl:variable name="ifPrecedingNormal">
            <xsl:for-each select="key('Part', 'word/styles.xml')">
              <xsl:if test="key('StyleId', $precedingStyleId)/w:basedOn/@w:val='Normal'">true</xsl:if>
            </xsl:for-each>
          </xsl:variable>
          <xsl:variable name="precedingNumId">
            <xsl:call-template name="GetListProperty">
              <xsl:with-param name="node" select="."/>
              <xsl:with-param name="property">w:numId</xsl:with-param>
            </xsl:call-template>
          </xsl:variable>

          <xsl:if test="$precedingOutlineLevel != '' and not(w:pPr/w:numPr) and $ifPrecedingNormal='true' and contains($precedingStyleId,'Heading') and $precedingNumId !=''">true</xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:variable>


    <!--math, dialogika: added for bugfix #1802258 BEGIN -->
    <xsl:variable name="IsDefaultHeading">
      <xsl:call-template name="CheckDefaultHeading">
        <xsl:with-param name="Name">
          <xsl:value-of select="key('Part', 'word/styles.xml')/w:styles/w:style[@w:styleId = $styleId]/w:name/@w:val" />
        </xsl:with-param>    
      </xsl:call-template>
    </xsl:variable>
    <!--math, dialogika: added for bugfix #1802258 END -->

    <xsl:choose>
      <!--check if the paragraph starts a table-of content or Bibliography or Alphabetical Index -->
      <xsl:when
        test="w:r[contains(w:instrText,'TOC') or contains(w:instrText,'BIBLIOGRAPHY') or contains(w:instrText, 'INDEX' )]">
        <xsl:apply-templates select="." mode="tocstart"/>
      </xsl:when>

      <!-- ignore paragraph if it's deleted in change tracking mode-->
      <!--xsl:when test="preceding::w:p[1]/w:pPr/w:rPr/w:del"/-->
      <xsl:when test="key('p', number(@oox:id)-1)/w:pPr/w:rPr/w:del"/>


      <!--  check if the paragraf is list element (it can be a heading but only if it's style is NOT linked to a list level 
        - for linked heading styles there's oultine list style created and they can't be in list (see bug  #1619448)) -->

      <xsl:when
        test="not($outlineLevel != '' and not(w:pPr/w:numPr) and $ifNormal='true' and contains($styleId,'Heading')) and $numId != '' and $level &lt; 10 and key('Part', 'word/numbering.xml')/w:numbering/w:num[@w:numId=$numId]/w:abstractNumId/@w:val != ''">
        <!--xsl:when
          test="$numId != '' and $level &lt; 10 and key('Part', 'word/numbering.xml')/w:numbering/w:num[@w:numId=$numId]/w:abstractNumId/@w:val != '' 
          and not(key('Part', 'word/styles.xml')/w:styles/w:style[@w:styleId = $styleId and child::w:pPr/w:outlineLvl and child::w:pPr/w:numPr/w:numId])"-->
        <xsl:apply-templates select="." mode="list">
          <xsl:with-param name="numId" select="$numId"/>
          <xsl:with-param name="level" select="$level"/>
          <xsl:with-param name="dummyOutlineLevel" select="$dummyOutlineLevel"/>
          <xsl:with-param name="isPrecedingHeading" select="$isPrecedingHeading"/>
        </xsl:apply-templates>
      </xsl:when>

      <!--math, dialogika: changed for bugfix #1802258 BEGIN -->
      <!--  check if the paragraf is heading -->
      <!--<xsl:when test="$outlineLevel != ''">-->
      <xsl:when test="$outlineLevel != '' and $IsDefaultHeading='true'">
      <!--math, dialogika: changed for bugfix #1802258 END -->        
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
      <!--header outline level -->
      <xsl:call-template name="InsertHeadingOutlineLvl">
        <xsl:with-param name="outlineLevel" select="$outlineLevel"/>
      </xsl:call-template>
      <!-- unnumbered heading is list header  -->
      <xsl:call-template name="InsertHeadingAsListHeader"/>
      <!--change track end-->
      <xsl:if test="key('p', number(@oox:id)-1)/w:pPr/w:rPr/w:ins and $numId!=''">
        <text:change-end>
          <xsl:attribute name="text:change-id">
            <xsl:value-of select="generate-id(key('p', number(@oox:id)-1))"/>
          </xsl:attribute>
        </text:change-end>
      </xsl:if>
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
    <xsl:apply-templates select="key('p', @oox:id+1)/child::node()"/>
    <xsl:if test="key('p', @oox:id+1)/w:pPr/w:rPr/w:del">
      <xsl:for-each select="key('p', @oox:id+1)">
        <xsl:call-template name="InsertDeletedParagraph"/>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>

  <!--  set heading as list header (needed when number was deleted manually)-->
  <xsl:template name="InsertHeadingAsListHeader">
    <xsl:variable name="numId">
      <xsl:call-template name="GetListProperty">
        <xsl:with-param name="node" select="."/>
        <xsl:with-param name="property">w:numId</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <!--check if there's a any numId in document-->
    <xsl:for-each select="key('Part', 'word/document.xml')">
      <xsl:choose>
        <xsl:when test="key('pPr','')/w:numPr/w:ins"/>
        <xsl:when test="key('pPr', '')/w:numPr/w:numId">
          <xsl:call-template name="InsertListHeader">
            <xsl:with-param name="numId" select="$numId"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <!--check if there's a any numId in styles-->
          <xsl:for-each select="key('Part', 'word/styles.xml')">
            <xsl:if test="key('pPr', '')/w:numPr/w:numId">
              <xsl:call-template name="InsertListHeader">
                <xsl:with-param name="numId" select="$numId"/>
              </xsl:call-template>
            </xsl:if>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <!--  set heading as list header (needed when number was deleted manually)-->
  <xsl:template name="InsertListHeader">
    <xsl:param name="numId"/>
    <xsl:for-each select="key('Part', 'word/numbering.xml')">
      <xsl:if test="not(key('numId', $numId))">
        <xsl:attribute name="text:is-list-header">
          <xsl:text>true</xsl:text>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
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
            test="w:pPr or w:r/w:br[@w:type='page' or @w:type='column'] or key('Part', 'word/styles.xml')/w:styles/w:docDefaults/w:pPrDefault">
            <xsl:attribute name="text:style-name">
              <xsl:choose>
                <xsl:when test="./w:r/w:ptab/@w:alignment = 'right' and ./w:pPr/w:pStyle/@w:val = 'Footer'">
                  <xsl:text>X3AS7TABSTYLE</xsl:text>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="generate-id(self::node())"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="key('p', number(@oox:id)-1)/w:pPr/w:rPr/w:ins and $numId!=''">
            <text:change-end>
              <xsl:attribute name="text:change-id">
                <xsl:value-of select="generate-id(key('p', number(@oox:id)-1))"/>
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

  <!-- Detect if there is text before a break page in a paragraph -->
  <!-- BUG 1583404 - Page breaks not correct converted - 17/07/2007-->
  <xsl:template match="w:r[w:br[@w:type='page']]">
    <xsl:if test="preceding-sibling::w:r/w:t[1]">
      <pcut:paragraph_to_cut/>
    </xsl:if>
    <xsl:apply-templates select="w:t"/>
  </xsl:template>

  <!--tabs-->
  <xsl:template match="w:tab[not(parent::w:tabs)]">
    <xsl:choose>
      <xsl:when test="ancestor::w:footnote or ancestor::w:endnote">
        <xsl:variable name="StyleId">
          <xsl:value-of select="ancestor::w:p/w:pPr/w:pStyle/@w:val"/>
        </xsl:variable>
        <xsl:choose>
          <xsl:when
            test="generate-id(.) = generate-id(ancestor::w:p/descendant::w:tab[1]) and (ancestor::w:p/w:pPr/w:ind/@w:hanging != '' or key('StyleId', $StyleId)/w:pPr/w:ind/@w:hanging != '')"/>
          <xsl:otherwise>
            <text:tab/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <text:tab/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--clam, dialogika: bugfix #1803097-->
  <!--<xsl:template match="w:ptab[@w:relativeTo='margin' and @w:alignment='right' and @w:leader='none']">-->
    <xsl:template match="w:ptab">
    <text:tab/>
  </xsl:template>

  <!-- run -->
  <xsl:template match="w:r">
    <xsl:message terminate="no">progress:w:r</xsl:message>
    <xsl:choose>

      <!--fields-->
      <!--xsl:when test="preceding::w:fldChar[1][@w:fldCharType='begin' or @w:fldCharType='separate']"-->
      <xsl:when test="@oox:f">
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
      <xsl:when test="contains($linkHref, 'file:///') or contains($linkHref, 'http://') or contains($linkHref, 'https://') or contains($linkHref, 'mailto:')">
        <xsl:value-of select="$linkHref"/>
      </xsl:when>
      <xsl:when test="contains($linkHref,'#')">
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
                <xsl:value-of select="generate-id(/node())" />
              </xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:for-each
            select="key('Part', concat('word/_rels/',$document,'.rels'))//node()[name() = 'Relationship']">
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
              <xsl:value-of select="substring-before(substring-after(preceding::w:instrText[1],'&quot;'),'&quot;')" />
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

  <!--  text bookmark-Start  -->
  <xsl:template match="w:bookmarkStart">
    <!--
    makz: check if the w:bookmarkStart doesn't belong to a user field.
    user fields are translated to user-field-decl
    -->
    <xsl:if test="ancestor::w:p and not(preceding-sibling::w:r[1]/w:fldChar/@w:fldCharType='seperate')">
      
      <xsl:variable name="NameBookmark">
        <xsl:value-of select="@w:name"/>
      </xsl:variable>
      <xsl:variable name="OutlineLvl">
        <xsl:value-of select="parent::w:p/w:pPr/w:outlineLvl/@w:val"/>
      </xsl:variable>
      <xsl:variable name="Id">
        <xsl:value-of select="@w:id"/>
      </xsl:variable>

      <xsl:choose>
        <!--math, dialogika: bugfix #1785483 BEGIN-->
        <!--<xsl:when test="contains($NameBookmark, '_Toc') and $OutlineLvl != ''">-->
        <xsl:when test="contains($NameBookmark, '_Toc') and $OutlineLvl != '' and $OutlineLvl !='9'">        
        <!--math, dialogika: bugfix #1785483 END-->
          <text:bookmark>
            <xsl:attribute name="text:name">
              <xsl:value-of select="$NameBookmark"/>
            </xsl:attribute>
          </text:bookmark>


          <text:toc-mark-start>
            <xsl:attribute name="text:id">
              <xsl:value-of select="@w:id"/>
            </xsl:attribute>
            <xsl:attribute name="text:outline-level">
              <xsl:value-of select="$OutlineLvl + 1"/>
            </xsl:attribute>
          </text:toc-mark-start>
          
        </xsl:when>
        <!-- a bookmark must begin and end in the same text flow -->
        <xsl:when test="ancestor::w:tc and not(ancestor::w:tc//w:bookmarkEnd[@w:id=$Id])">
          <text:bookmark>
            <xsl:attribute name="text:name">
              <xsl:value-of select="$NameBookmark"/>
            </xsl:attribute>
          </text:bookmark>
        </xsl:when>
        <xsl:otherwise>
          <text:bookmark-start>
            <xsl:attribute name="text:name">
              <xsl:value-of select="$NameBookmark"/>
            </xsl:attribute>
          </text:bookmark-start>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
  </xsl:template>
  
  <xsl:template match="w:bookmarkEnd">
    <xsl:variable name="Id">
      <xsl:value-of select="@w:id"/>
    </xsl:variable>
    
    <!--
    makz: check if the w:bookmarkStart doesn't belong to a user field.
    user fields are translated to user-field-decl
    -->
    <xsl:if test="ancestor::w:p and not(following-sibling::w:r[1]/w:fldChar/@w:fldCharType='end')">
      <xsl:variable name="NameBookmark">
        <xsl:value-of select="key('bookmarkStart', @w:id)/@w:name"/>
      </xsl:variable>
      <xsl:variable name="OutlineLvl">
        <xsl:value-of select="parent::w:p/w:pPr/w:outlineLvl/@w:val"/>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="contains($NameBookmark, '_Toc')  and  $OutlineLvl != ''">
          <text:toc-mark-end>
            <xsl:attribute name="text:id">
              <xsl:value-of select="@w:id"/>
            </xsl:attribute>
          </text:toc-mark-end>
        </xsl:when>
        <!-- a bookmark must begin and end in the same text flow -->
        <xsl:when test="not(ancestor::w:tc) or ancestor::w:tc//w:bookmarkStart[@w:id=$Id]">
          <text:bookmark-end>
            <xsl:attribute name="text:name">
              <xsl:value-of select="$NameBookmark"/>
            </xsl:attribute>
          </text:bookmark-end>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
  </xsl:template>

  <!-- simple text  -->
  <xsl:template match="w:t">
    <xsl:message terminate="no">progress:w:t</xsl:message>
    <xsl:call-template name="InsertDropCapText"/>
    <xsl:choose>
      <!--check whether string contains  whitespace sequence-->
      <xsl:when test="not(contains(., '  '))">
        <xsl:value-of select="."/>
      </xsl:when>
      <xsl:otherwise>
        <!--converts whitespaces sequence to text:s-->
        <xsl:call-template name="InsertWhiteSpaces"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <!-- drop cap text  (only on first w:t of first w:r of w:p) -->
  <xsl:template name="InsertDropCapText">
    <xsl:if test="not(preceding-sibling::w:t[1]) and not(parent::w:r/preceding-sibling::w:r[1])">
      <xsl:variable name="prev-paragraph"
        select="ancestor-or-self::w:p[1]/preceding-sibling::w:p[1]"/>
      <xsl:if test="$prev-paragraph/w:pPr/w:framePr[@w:dropCap]">
        <text:span text:style-name="{generate-id($prev-paragraph/w:r[1])}">
          <xsl:value-of select="$prev-paragraph/w:r[1]/w:t"/>
        </text:span>
      </xsl:if>
    </xsl:if>
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

  <xsl:template match="w:noBreakHyphen">–</xsl:template>


  <!-- 
  *************************************************************************
  Templates for creating Automatic Styles
  *************************************************************************
  -->
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

  <!-- symbols : text style -->
  <xsl:template match="w:sym" mode="automaticstyles">
    <style:style style:name="{generate-id(.)}" style:family="text">
      <xsl:if test="@w:font">
        <style:text-properties>
          <xsl:attribute name="style:font-name">
            <xsl:value-of select="@w:font"/>
          </xsl:attribute>
          <xsl:attribute name="style:font-name-complex">
            <xsl:value-of select="@w:font"/>
          </xsl:attribute>
        </style:text-properties>
      </xsl:if>
    </style:style>
  </xsl:template>

  <!-- symbols -->
  <xsl:template match="w:sym">
    <text:span text:style-name="{generate-id(.)}">
      <!-- character post-processing -->
      <pchar:symbol pchar:code="{@w:char}"/>
    </text:span>
  </xsl:template>

  <!--ignore text in automatic styles mode-->
  <xsl:template match="text()" mode="automaticstyles"/>

  <!-- insert a master page name when required -->
  <xsl:template name="MasterPageName">
    <xsl:if test="ancestor::w:body">
      <xsl:choose>
        <!-- particular case : if paragraph is last one of a section -->
        <xsl:when test="w:sectPr">
          <xsl:call-template name="ComputeMasterPageName">
            <xsl:with-param name="followingSectPr" select="w:sectPr"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="ComputeMasterPageName"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
  </xsl:template>

  <xsl:template name="ComputeMasterPageName">
    <!-- NB : precSectPr defines properties of preceding section,
        whereas followingSectPr defines properties of current section -->
    <xsl:param name="precSectPr" select="preceding::w:p[w:pPr/w:sectPr][1]/w:pPr/w:sectPr"/>
    <xsl:param name="followingSectPr" select="following::w:p[w:pPr/w:sectPr][1]/w:pPr/w:sectPr"/>
    <!--<xsl:param name="mainSectPr" select="key('Part', 'word/document.xml')/w:document/w:body/w:sectPr"/>-->
    
    <!--<xsl:param name="precSectPr" select="key('sectPr', number(ancestor::node()/@oox:s) - 1)"/>
    <xsl:param name="followingSectPr" select="key('sectPr', number(ancestor::node()/@oox:s))"/>-->
    <xsl:param name="mainSectPr" select="key('Part', 'word/document.xml')/w:document/w:body/w:sectPr"/>
    
    <xsl:choose>
      <!-- first case : current section is continuous with preceding section (test if precSectPr exist to avoid bugs) -->
      <xsl:when test="$precSectPr and $followingSectPr/w:type/@w:val = 'continuous' ">
        <!-- no new master page. Warn loss of page header/footer change (should not occure in OOX, but Word 2007 handles it) -->
        <xsl:if
          test="$followingSectPr/w:headerReference[@w:type='default']/@r:id != $precSectPr/w:headerReference[@w:type='default']/@r:id
          or $followingSectPr/w:headerReference[@w:type='even']/@r:id != $precSectPr/w:headerReference[@w:type='even']/@r:id 
          or $followingSectPr/w:headerReference[@w:type='first']/@r:id != $precSectPr/w:headerReference[@w:type='first']/@r:id 
          or $followingSectPr/w:footerReference[@w:type='default']/@r:id != $precSectPr/w:footerReference[@w:type='default']/@r:id
          or $followingSectPr/w:footerReference[@w:type='even']/@r:id != $precSectPr/w:footerReference[@w:type='even']/@r:id 
          or $followingSectPr/w:footerReference[@w:type='first']/@r:id != $precSectPr/w:footerReference[@w:type='first']/@r:id">
          <xsl:message terminate="no">
            <xsl:text>feedback:Header/footer change after continuous section break.</xsl:text>
          </xsl:message>
        </xsl:if>
      </xsl:when>
      <xsl:when test="not(preceding::w:p[ancestor::w:body])">
        <xsl:choose>
          <xsl:when test="$followingSectPr">
            <xsl:choose>
              <xsl:when
                test="$followingSectPr/w:titlePg">
                <xsl:attribute name="style:master-page-name">
                  <xsl:value-of select="concat('First_H_',generate-id($followingSectPr))"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:attribute name="style:master-page-name">
                  <xsl:value-of select="concat('H_',generate-id($followingSectPr))"/>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>
            <xsl:choose>
              <xsl:when
                test="$mainSectPr/w:titlePg">
                <xsl:attribute name="style:master-page-name">First_Page</xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:attribute name="style:master-page-name">Standard</xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <!--clam: bugfix #1800794-->
        <!--<xsl:if test="preceding::w:p[parent::w:body|parent::w:tbl/tr/tv][1]/w:pPr/w:sectPr">-->
        <xsl:if test="preceding::w:p[parent::w:body|parent::w:tc][1]/w:pPr/w:sectPr">
          <xsl:choose>
            <xsl:when
              test="$followingSectPr and not($followingSectPr/w:headerReference) and not($followingSectPr/w:footerReference)">
              <xsl:attribute name="style:master-page-name">
                <!-- jslaurent : hack to make it work in any situation. Does not make any sense though.
                master page names should be reviewed and unified : many names not consistent, many styles never used -->
                <xsl:value-of select="concat('H_',generate-id($followingSectPr))"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when test="$followingSectPr">
                  <xsl:choose>
                    <xsl:when test="$followingSectPr/w:titlePg">
                      <xsl:attribute name="style:master-page-name">
                        <xsl:value-of select="concat('First_H_',generate-id($followingSectPr))"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:if
                        test="$followingSectPr/w:headerReference or $followingSectPr/w:footerReference">
                        <xsl:attribute name="style:master-page-name">
                          <xsl:value-of select="concat('H_',generate-id($followingSectPr))"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:choose>
                    <xsl:when
                      test="$mainSectPr/w:titlePg">
                      <xsl:attribute name="style:master-page-name">First_Page</xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:choose>
                        <xsl:when test="following-sibling::w:r/w:lastRenderedPageBreak and $precSectPr/w:type/@w:val = 'continuous' and generate-id(..) = generate-id($precSectPr/following::w:p[1])">
                          <xsl:attribute name="style:master-page-name">Standard</xsl:attribute>                          
                        </xsl:when>
                        <xsl:when test="$precSectPr and $mainSectPr/w:type/@w:val = 'continuous' ">
                          <!-- no new master page. Warn loss of page header/footer change (should not occure in OOX, but Word 2007 handles it) -->
                          <xsl:if
                            test="$mainSectPr/w:headerReference[@w:type='default']/@r:id != $precSectPr/w:headerReference[@w:type='default']/@r:id
                              or $mainSectPr/w:headerReference[@w:type='even']/@r:id != $precSectPr/w:headerReference[@w:type='even']/@r:id 
                              or $mainSectPr/w:headerReference[@w:type='first']/@r:id != $precSectPr/w:headerReference[@w:type='first']/@r:id 
                              or $mainSectPr/w:footerReference[@w:type='default']/@r:id != $precSectPr/w:footerReference[@w:type='default']/@r:id
                              or $mainSectPr/w:footerReference[@w:type='even']/@r:id != $precSectPr/w:footerReference[@w:type='even']/@r:id 
                              or $mainSectPr/w:footerReference[@w:type='first']/@r:id != $precSectPr/w:footerReference[@w:type='first']/@r:id">
                            <xsl:message terminate="no">
                              <xsl:text>feedback:Header/footer change after continuous section break.</xsl:text>
                            </xsl:message>
                          </xsl:if>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:attribute name="style:master-page-name">Standard</xsl:attribute>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>

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
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns="http://schemas.openxmlformats.org/package/2006/relationships" exclude-result-prefixes="w"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">

  <xsl:import href="tables.xsl"/>
  <xsl:import href="lists.xsl"/>

  <xsl:strip-space elements="*"/>
  <xsl:preserve-space elements="w:p"/>
  <xsl:preserve-space elements="w:r"/>

  <xsl:key name="numId" match="w:num" use="w:abstractNumId/@w:val"/>

  <!--main document-->
  <xsl:template name="content">
    <office:document-content>
      <office:scripts/>
      <office:font-face-decls>
        <xsl:apply-templates select="document('word/document.xml')/w:document/w:body"
          mode="fontfacedecls"/>
      </office:font-face-decls>
      <office:automatic-styles>
        <xsl:apply-templates select="document('word/document.xml')/w:document/w:body"
          mode="automaticstyles"/>
        <xsl:if test="document('word/document.xml')/w:document[descendant::w:numPr]">
          <xsl:apply-templates select="document('word/numbering.xml')/w:numbering"/>
        </xsl:if>
      </office:automatic-styles>
      <office:body>
        <office:text>
          <xsl:apply-templates select="document('word/document.xml')/w:document/w:body"/>
        </office:text>
      </office:body>
    </office:document-content>
  </xsl:template>

 <xsl:template match="w:rPr[parent::w:r]|w:rPr[parent::w:pPr]" mode="fontfacedecls">
    <xsl:if test="w:rFonts">
      <xsl:variable name="ascii" select="w:rFonts/@w:ascii"/>
      <xsl:if test="not(preceding::w:rFonts[@w:ascii = $ascii])">
        <style:font-face>
          <xsl:attribute name="style:name">
            <xsl:value-of select="$ascii"/>
          </xsl:attribute>
          <xsl:attribute name="svg:font-family">
            <xsl:value-of select="w:rFonts/@w:hAnsi"/>
          </xsl:attribute>
        </style:font-face>
       </xsl:if>
    </xsl:if>
 </xsl:template>
  
  <xsl:template match="w:t" mode="fontfacedecls">
    <!--do nothing-->
  </xsl:template>
  
  <!-- create a style for each paragraph. Do not take w:sectPr/w:rPr into consideration. -->
  <xsl:template match="w:pPr[parent::w:p]" mode="automaticstyles">
    <xsl:message terminate="no">progress:w:pPr</xsl:message>
    <style:style style:name="{generate-id(parent::w:p)}" style:family="paragraph">
      <xsl:if test="w:pStyle">
        <xsl:attribute name="style:parent-style-name">
          <xsl:value-of select="w:pStyle/@w:val"/>
        </xsl:attribute>
      </xsl:if>
      <style:paragraph-properties>
        <xsl:call-template name="InsertParagraphProperties"/>
      </style:paragraph-properties>
      <xsl:if test="w:rPr">
        <xsl:for-each select="w:rPr">
          <style:text-properties>
            <xsl:call-template name="InsertTextProperties"/>
          </style:text-properties>
        </xsl:for-each>
      </xsl:if>
    </style:style>
  </xsl:template>

  <!-- create a style for each run. Do not take w:pPr/w:rPr into consideration. -->
  <xsl:template match="w:rPr[parent::w:r]" mode="automaticstyles">
    <xsl:message terminate="no">progress:w:rPr</xsl:message>
    <style:style style:name="{generate-id(parent::w:r)}" style:family="text">
      <style:text-properties>
        <xsl:call-template name="InsertTextProperties"/>
      </style:text-properties>
    </style:style>
  </xsl:template>

  <!-- ignore text in automatic styles mode. -->
  <xsl:template match="w:t" mode="automaticstyles">
    <!-- do nothing. -->
  </xsl:template>

  <!--  Check if the paragraf is heading -->
  <xsl:template name="CheckHeading">
    <xsl:param name="outline"/>
    <xsl:variable name="basedOn">
      <xsl:value-of
        select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$outline]/w:basedOn/@w:val"
      />
    </xsl:variable>

    <!--  Search outlineLvl in style based on style -->
    <xsl:choose>
      <xsl:when
        test="document('word/styles.xml')/w:styles/w:style[@w:styleId=$basedOn]/w:pPr/w:outlineLvl/@w:val">
        <xsl:value-of
          select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$basedOn]/w:pPr/w:outlineLvl/@w:val"
        />
      </xsl:when>
      <xsl:when
        test="document('word/styles.xml')/w:styles/w:style[@w:styleId=$outline]/w:basedOn/@w:val">
        <xsl:call-template name="CheckHeading">
          <xsl:with-param name="outline">
            <xsl:value-of
              select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$outline]/w:basedOn/@w:val"
            />
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!-- Get outlineLvl if the paragraf is heading -->
  <xsl:template name="GetOutlineLevel">
    <xsl:choose>
      <xsl:when test="w:pPr/w:outlineLvl/@w:val">
        <xsl:value-of select="w:pPr/w:outlineLvl/@w:val"/>
      </xsl:when>
      <xsl:when test="w:pPr/w:pStyle/@w:val">
        <xsl:variable name="outline">
          <xsl:value-of select="w:pPr/w:pStyle/@w:val"/>
        </xsl:variable>
        <!--Search outlineLvl in styles.xml  -->
        <xsl:choose>
          <xsl:when
            test="document('word/styles.xml')/w:styles/w:style[@w:styleId=$outline]/w:pPr/w:outlineLvl/@w:val">
            <xsl:value-of
              select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$outline]/w:pPr/w:outlineLvl/@w:val"
            />
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="CheckHeading">
              <xsl:with-param name="outline">
                <xsl:value-of select="w:pPr/w:pStyle/@w:val"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
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
        <!--if outlineLvl is not defined search in parent style by w:link-->
        <xsl:choose>
          <xsl:when
            test="document('word/styles.xml')/w:styles/w:style[@w:styleId=$linkedStyleOutline]/w:pPr/w:outlineLvl/@w:val">
            <xsl:value-of
              select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$linkedStyleOutline]/w:pPr/w:outlineLvl/@w:val"
            />
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="CheckHeading">
              <xsl:with-param name="outline">
                <xsl:value-of
                  select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$outline]/w:link/@w:val"
                />
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--  paragraphs, lists, headings-->
  <xsl:template match="w:p">
    <xsl:message terminate="no">progress:w:p</xsl:message>

    <xsl:variable name="outlineLevel">
      <xsl:call-template name="GetOutlineLevel"/>
    </xsl:variable>

    <xsl:choose>
      <!-- check if list starts -->
      <xsl:when test="w:pPr/w:numPr">
        <xsl:variable name="NumberingId" select="w:pPr/w:numPr/w:numId/@w:val"/>
        <xsl:variable name="position" select="count(preceding-sibling::w:p)"/>
        <xsl:if
          test="not(preceding-sibling::node()[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId and  count(preceding-sibling::w:p)= $position -1])">
          <xsl:apply-templates select="." mode="list"/>
        </xsl:if>
      </xsl:when>

      <!--  check if the paragraf is heading -->
      <xsl:when test="$outlineLevel != '' ">
        <xsl:apply-templates select="." mode="heading">
          <xsl:with-param name="outlineLevel" select="$outlineLevel"/>
        </xsl:apply-templates>
      </xsl:when>

      <xsl:otherwise>
        <!--  default scenario - paragraph-->
        <xsl:apply-templates select="." mode="paragraph"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--paragraph with ouline level is a heading-->
  <xsl:template match="w:p" mode="heading">
    <xsl:param name="outlineLevel"/>
    <text:h>
      <xsl:if test="w:pPr">
        <xsl:attribute name="text:style-name">
          <xsl:value-of select="generate-id(self::node())"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:attribute name="text:outline-level">
        <xsl:value-of select="$outlineLevel+1"/>
      </xsl:attribute>
      <xsl:apply-templates/>
    </text:h>
  </xsl:template>

  <!--  paragraphs-->
  <xsl:template match="w:p" mode="paragraph">
    <text:p>
      <xsl:if test="w:pPr">
        <xsl:attribute name="text:style-name">
          <xsl:value-of select="generate-id(self::node())"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:apply-templates/>
    </text:p>
  </xsl:template>

  <!-- paragraph which has  numbering properties is a list item -->
  <xsl:template match="w:p" mode="list">
    <xsl:variable name="NumberingId2" select="w:pPr/w:numPr/w:numId/@w:val"/>

    <xsl:variable name="position2" select="count(preceding-sibling::w:p)"/>
    <xsl:choose>
      <!-- Is first  element at list -->
      <xsl:when
        test="not(preceding-sibling::node()[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId2 and  count(preceding-sibling::w:p)= $position2 -1])">
        <text:list text:style-name="{concat('L',$NumberingId2)}">
          <xsl:if test="preceding-sibling::w:p[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId2]">
            <xsl:attribute name="text:continue-numbering">true</xsl:attribute>
          </xsl:if>
          <text:list-item>
            <xsl:apply-templates select="." mode="paragraph"/>
          </text:list-item>
          <!-- next paragraph in same list member -->
          <xsl:if
            test="following-sibling::w:p[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId2 and  count(preceding-sibling::w:p)= $position2 +1]">
            <xsl:apply-templates
              select="following-sibling::w:p[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId2 and  count(preceding-sibling::w:p)= $position2 +1]"
              mode="list"/>
          </xsl:if>
        </text:list>
      </xsl:when>
      <xsl:otherwise>
        <text:list-item>
          <xsl:apply-templates select="." mode="paragraph"/>
        </text:list-item>
        <xsl:if
          test="following-sibling::w:p[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId2 and  count(preceding-sibling::w:p)= $position2 +1]">
          <xsl:apply-templates
            select="following-sibling::w:p[child::w:pPr/w:numPr/w:numId/@w:val = $NumberingId2 and  count(preceding-sibling::w:p)= $position2 +1]"
            mode="list"/>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- run -->
  <xsl:template match="w:r">
    <xsl:message terminate="no">progress:w:r</xsl:message>

    <xsl:choose>
      <!--  fieldchar hyperlink -->
      <xsl:when test="contains(preceding::w:instrText[1],'HYPERLINK') and 
        count(preceding::w:fldChar[@w:fldCharType = 'begin']) &gt; count(preceding::w:fldChar[@w:fldCharType = 'end'])">
        <xsl:call-template name="InsertHyperlink"/>
      </xsl:when>
      <!-- attach automatic style-->
      <xsl:when test="w:rPr">
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

  <!-- ignore text inside a field code -->
  <xsl:template match="w:instrText"/>

 <!-- path for hyperlinks-->
  <xsl:template name="GetLinkPath">
    <xsl:param name="linkHref" />
      
    <xsl:choose>
      <xsl:when test="contains($linkHref, 'file:///') or contains($linkHref, 'http://') or contains($linkHref, 'mailto:')">
        <xsl:value-of select="$linkHref"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="concat('../',$linkHref)"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- hyperlinks -->
  <xsl:template name="InsertHyperlink">
    <text:a xlink:type="simple">
      <!-- document hyperlink -->
      <xsl:if test="@w:anchor">
        <xsl:attribute name="xlink:href">
          <xsl:value-of select="concat('#',@w:anchor)"/>
        </xsl:attribute>
      </xsl:if>

      <!-- file or web page hyperlink with relationship id -->
      <xsl:if test="@r:id">
        <xsl:variable name="relationshipId">
          <xsl:value-of select="@r:id"/>
        </xsl:variable>
        
        <xsl:for-each
          select="document('word/_rels/document.xml.rels')//node()[name() = 'Relationship']">
          <xsl:if test="./@Id=$relationshipId">
            <xsl:attribute name="xlink:href">
              <xsl:call-template name="GetLinkPath">
                <xsl:with-param name="linkHref" select="@Target" />
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
      
      <!-- file or web page hyperlink -  fieldchar type (can contain several paragraphs in Word) -->
      <xsl:if test="self::w:r">
        <xsl:attribute name="xlink:href">
            <xsl:call-template name="GetLinkPath">
              <xsl:with-param name="linkHref"><xsl:value-of select="substring-before(substring-after(preceding::w:instrText[1],'&quot;'),'&quot;')"/></xsl:with-param>
            </xsl:call-template>
         </xsl:attribute>
      </xsl:if>

      <xsl:choose>
        <!-- attach automatic style-->
        <xsl:when test="w:rPr">
          <text:span text:style-name="{generate-id(self::node())}">
            <xsl:apply-templates/>
          </text:span>
        </xsl:when>
        <xsl:otherwise>
          <xsl:apply-templates/>
        </xsl:otherwise>
      </xsl:choose>
    </text:a>
  </xsl:template>

  <!-- hyperlinks -->
  <xsl:template match="w:hyperlink">
    <xsl:call-template name="InsertHyperlink"/>
  </xsl:template>

  <!--  text bookmark mark start -->
  <xsl:template match="w:bookmarkStart">
    <text:bookmark-start>
      <xsl:attribute name="text:name">
        <xsl:value-of select="@w:name"/>
      </xsl:attribute>
    </text:bookmark-start>
  </xsl:template>

  <!--  text bookmark mark end-->
  <xsl:template match="w:bookmarkEnd">
    <text:bookmark-end>
      <xsl:variable name="IdBookmark">
        <xsl:value-of select="@w:id"/>
      </xsl:variable>
      <xsl:attribute name="text:name">
        <xsl:value-of select="ancestor::w:body/w:p/w:bookmarkStart[@w:id=$IdBookmark]/@w:name"/>
      </xsl:attribute>
      <xsl:message>
        <xsl:value-of select="ancestor::w:body/w:p/w:bookmarkStart[@w:id=$IdBookmark]/@w:name"/>
      </xsl:message>
    </text:bookmark-end>
  </xsl:template>

  <!-- footnotes -->
  <xsl:template match="w:r[w:footnoteReference]">
    <xsl:variable name="footnoteId" select="w:footnoteReference/@w:id"/>

    <!-- change context to get the footnote content -->
    <xsl:for-each select="document('word/footnotes.xml')/w:footnotes/w:footnote[@w:id=$footnoteId]">
      <text:span>
        <text:note text:id="{concat('ftn',count(preceding-sibling::w:footnote) - 1)}"
          text:note-class="footnote">
          <text:note-citation>
            <xsl:choose>
              <xsl:when
                test="@w:suppressRef = 1 or @w:suppressRef = 'true' or @w:suppressRef = 'On' ">
                <xsl:attribute name="text:label">
                  <xsl:apply-templates select="w:t"/>
                </xsl:attribute>
                <xsl:apply-templates select="w:t"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="count(preceding-sibling::w:footnote) - 1"/>
              </xsl:otherwise>
            </xsl:choose>
          </text:note-citation>
          <text:note-body>
            <!-- TODO : only take relevant content (ignore runs concerning footnoteproperties). -->
            <xsl:apply-templates select="."/>
          </text:note-body>
        </text:note>
      </text:span>
    </xsl:for-each>
  </xsl:template>

  <!-- endnote -->
  <xsl:template match="w:r[w:endnoteReference]">
    <xsl:variable name="endnoteId" select="w:endnoteReference/@w:id"/>

    <!-- change context to get the footnote content -->
    <xsl:for-each select="document('word/endnotes.xml')/w:endnotes/w:endnote[@w:id=$endnoteId]">
      <text:span>
        <text:note text:id="{concat('edn',count(preceding-sibling::w:endnote) - 1)}"
          text:note-class="endnote">
          <text:note-citation>
            <xsl:choose>
              <xsl:when
                test="@w:suppressRef = 1 or @w:suppressRef = 'true' or @w:suppressRef = 'On' ">
                <xsl:attribute name="text:label">
                  <xsl:apply-templates select="w:t"/>
                </xsl:attribute>
                <xsl:apply-templates select="w:t"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="count(preceding-sibling::w:endnote) - 1"/>
              </xsl:otherwise>
            </xsl:choose>
          </text:note-citation>
          <text:note-body>
            <!-- TODO : only take relevant content (ignore runs concerning footnoteproperties). -->
            <xsl:apply-templates select="."/>
          </text:note-body>
        </text:note>
      </text:span>
    </xsl:for-each>
  </xsl:template>

  <!-- simple text  -->
  <xsl:template match="w:t">
    <xsl:message terminate="no">progress:w:t</xsl:message>
    <xsl:value-of select="."/>
  </xsl:template>

  <!-- line breaks  (page-break todo)-->
  <xsl:template match="w:br">
    <text:line-break/>
  </xsl:template>

</xsl:stylesheet>

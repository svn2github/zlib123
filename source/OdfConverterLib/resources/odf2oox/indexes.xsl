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
      <xsl:text>TOC  \t "</xsl:text>
      <xsl:value-of select="parent::text:index-body/preceding-sibling::text:table-index-source/@text:caption-sequence-name"/>
      <xsl:text> " </xsl:text> 
    </w:instrText>
  </xsl:template>

  <xsl:template name="insertAlphabeticalPrefs">
    <xsl:choose>
      <xsl:when
        test="key('Index', @text:style-name)/style:paragraph-properties/style:tab-stops/style:tab-stop/@style:type='right'">
        <w:instrText xml:space="preserve">INDEX \e "</w:instrText>
        <w:tab/>
        <!--Right Align Page Number-->
        <xsl:choose>
          <xsl:when
            test="key('Index', ancestor-or-self::text:alphabetical-index/@text:style-name)/style:section-properties/style:columns/@fo:column-count>4"
            >4 <w:instrText xml:space="preserve">" \c "4" \z "1045"</w:instrText>
          </xsl:when>
          <xsl:when
            test="key('Index', ancestor-or-self::text:alphabetical-index/@text:style-name)/style:section-properties/style:columns/@fo:column-count>1">
            <w:instrText xml:space="preserve">" \c "
              <xsl:value-of select="key('Index', ancestor-or-self::text:alphabetical-index/@text:style-name)/style:section-properties/style:columns/@fo:column-count"/>
              " \z "1045"</w:instrText>
          </xsl:when>
          <xsl:otherwise>
            <w:instrText xml:space="preserve">" \c "1" \z "1045"</w:instrText>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when
            test="key('Index', ancestor-or-self::text:alphabetical-index/@text:style-name)/style:section-properties/style:columns/@fo:column-count>4">
            <w:instrText xml:space="preserve">INDEX \c "4"\z "1045" </w:instrText>
          </xsl:when>
          <xsl:when
            test="key('Index', ancestor-or-self::text:alphabetical-index/@text:style-name)/style:section-properties/style:columns/@fo:column-count>1">
            <w:instrText xml:space="preserve">INDEX \c "
              <xsl:value-of select="key('Index', ancestor-or-self::text:alphabetical-index/@text:style-name)/style:section-properties/style:columns/@fo:column-count"/>
              "\z "1045" </w:instrText>
          </xsl:when>
          <xsl:otherwise>
            <w:instrText xml:space="preserve">INDEX \c "1"\z "1045" </w:instrText>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertIllustrationInPrefs">
    <w:instrText xml:space="preserve"> 
      <xsl:text> TOC  \t "</xsl:text>
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



  <!--inserts index item content for all types of index-->
  <xsl:template name="InsertIndexItemContent">

    <!-- references to index bookmark id in text -->
    <xsl:param name="tocId" select="count(preceding-sibling::text:p)+1"/>

    <!-- alphabetical index doesn't support page reference link -->

    <!-- insert TOC -->
    <xsl:choose>
      <xsl:when test="self::text:a">
        <xsl:apply-templates mode="paragraph"/>
        <xsl:apply-templates select="parent::text:p/child::node()[not(self::text:a)]"
          mode="paragraph"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates mode="paragraph"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertIndexPageRefEnd">
    <w:r>
      <w:rPr>
        <w:noProof/>
        <xsl:if test="ancestor::text:section/@text:display='none'">
          <w:vanish/>
        </xsl:if>
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
        <xsl:if test="ancestor::text:section/@text:display='none'">
          <w:vanish/>
        </xsl:if>
        <w:webHidden/>
      </w:rPr>
      <w:fldChar w:fldCharType="begin">
        <w:fldData xml:space="preserve">CNDJ6nn5us4RjIIAqgBLqQsCAAAACAAAAA4AAABfAFQAbwBjADEANAAxADgAMwA5ADIANwA2AAAA</w:fldData>
      </w:fldChar>
    </w:r>
    <w:r>
      <w:rPr>
        <w:noProof/>
        <xsl:if test="ancestor::text:section/@text:display='none'">
          <w:vanish/>
        </xsl:if>
        <w:webHidden/>
      </w:rPr>
      <w:instrText xml:space="preserve"><xsl:value-of select="concat('PAGEREF _Toc', $tocId,generate-id(ancestor::node()[child::text:index-body]), ' \h')"/></w:instrText>
    </w:r>
    <w:r>
      <w:rPr>
        <w:noProof/>
        <xsl:if test="ancestor::text:section/@text:display='none'">
          <w:vanish/>
        </xsl:if>
        <w:webHidden/>
      </w:rPr>
      <w:fldChar w:fldCharType="separate"/>
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




</xsl:stylesheet>

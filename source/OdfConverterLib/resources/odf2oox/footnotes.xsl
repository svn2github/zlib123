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
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/3/main"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  exclude-result-prefixes="text style office xlink draw pzip">

  <!-- Group footnotes under the same key -->
  <xsl:key name="footnotes" match="text:note[@text:note-class='footnote']" use="''"/>

  <xsl:template name="footnotes">
    <w:footnotes>

      <!-- special footnotes -->
      <w:footnote w:type="separator" w:id="0">
        <w:p>
          <!-- TODO :  extract properties from document('styles.xml')/office:automatic-styles/style:page-layout/style:footnote-sep -->
          <w:pPr>
            <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
          </w:pPr>
          <w:r>
            <w:separator/>
          </w:r>
        </w:p>
      </w:footnote>
      <w:footnote w:type="continuationSeparator" w:id="1">
        <w:p>
          <!-- TODO :  extract properties from document('styles.xml')/office:automatic-styles/style:page-layout/ -->
          <w:pPr>
            <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
          </w:pPr>
          <w:r>
            <w:continuationSeparator/>
          </w:r>
        </w:p>
      </w:footnote>

      <!-- normal footnotes -->

      <!-- absurd hack for changing the context -->
      <xsl:for-each select="document('content.xml')">
        <!-- iterating over the footnotes -->
        <xsl:for-each select="key('footnotes', '')">
          <w:footnote w:type="normal" w:id="{position() + 1}">
            <!-- This a custom mark, don't increment the counter -->
            <xsl:if test="text:note-citation/@text:label">
              <xsl:attribute name="w:suppressRef">1</xsl:attribute>
            </xsl:if>
            <xsl:apply-templates select="text:note-body"/>
          </w:footnote>
        </xsl:for-each>
      </xsl:for-each>

    </w:footnotes>
  </xsl:template>



  <!-- footnotes configuration -->
  <xsl:template match="text:notes-configuration[@text:note-class='footnote']" mode="note">
    <xsl:param name="wide" select="no"/>
    <w:footnotePr>

      <xsl:choose>
        <xsl:when test="ancestor::style:style[@style:family='section']">
          <w:pos w:val="beneathText"/>
        </xsl:when>
        <xsl:when test="@text:footnotes-position">
          <w:pos>
            <xsl:attribute name="w:val">
              <xsl:choose>
                <xsl:when test="@text:footnotes-position = 'text' ">beneathText</xsl:when>
                <xsl:otherwise>pageBottom</xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </w:pos>
        </xsl:when>
      </xsl:choose>

      <xsl:if test="@style:num-format">
        <w:numFmt>
          <xsl:attribute name="w:val">
            <xsl:call-template name="num-format">
              <xsl:with-param name="format" select="@style:num-format"/>
            </xsl:call-template>
          </xsl:attribute>
        </w:numFmt>
      </xsl:if>

      <xsl:if test="@text:start-value">
        <w:numStart w:val="{number(@text:start-value)+1}"/>
      </xsl:if>

      <xsl:choose>
        <xsl:when test="ancestor::style:style[@style:family='section']">
          <w:numRestart w:val="eachSect"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:if test="@text:start-numbering-at">
            <w:numRestart>
              <xsl:attribute name="w:val">
                <xsl:choose>
                  <xsl:when test="@text:start-numbering-at = 'page' ">eachPage</xsl:when>
                  <xsl:when test="@text:start-numbering-at = 'chapter' ">eachSect</xsl:when>
                  <xsl:otherwise>continuous</xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </w:numRestart>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:if test="$wide = 'yes' ">
        <w:footnote w:id="0"/>
        <w:footnote w:id="1"/>
      </xsl:if>

    </w:footnotePr>
  </xsl:template>



  <!-- Footnotes internal relationships -->
  <xsl:template name="InsertFootnotesInternalRelationships">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <xsl:for-each select="document('content.xml')">

        <!-- hyperlinks relationships. Do not pick up hyperlinks other than those coming from footnotes.  -->
        <xsl:call-template name="InsertHyperlinksRelationships">
          <xsl:with-param name="hyperlinks"
            select="key('hyperlinks', '')[ancestor::text:note/@text:note-class = 'footnote' ]"/>
        </xsl:call-template>

        <xsl:call-template name="InsertImagesRelationships">
          <xsl:with-param name="images"
            select="key('images', '')[ancestor::text:note/@text:note-class = 'footnote' ]"/>
        </xsl:call-template>
        
      </xsl:for-each>
    </Relationships>
  </xsl:template>

</xsl:stylesheet>

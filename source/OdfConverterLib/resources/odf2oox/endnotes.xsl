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
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  exclude-result-prefixes="text style office xlink draw pzip">

  <xsl:key name="endnotes" match="text:note[@text:note-class='endnote']" use="''"/>

  <xsl:template name="endnotes">
    <w:endnotes>

      <!-- special endnotes -->
      <w:endnote w:type="separator" w:id="0">
        <w:p>
          <!-- there are no styles for endnotes separator in oo -->
          <w:pPr>
            <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
          </w:pPr>
          <w:r>
            <w:separator/>
          </w:r>
        </w:p>
      </w:endnote>
      <w:endnote w:type="continuationSeparator" w:id="1">
        <w:p>
          <w:pPr>
            <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
          </w:pPr>
          <w:r>
            <w:continuationSeparator/>
          </w:r>
        </w:p>
      </w:endnote>

      <!-- normal endnotes -->
      <xsl:for-each select="document('content.xml')">
        <!-- warn loss of page break before endnotes -->
        <xsl:if test="key('endnotes', '')">
          <xsl:message terminate="no">translation.odf2oox.pageBreakBeforeEndnotes</xsl:message>
        </xsl:if>
        <xsl:for-each select="key('endnotes', '')">
          <w:endnote w:type="normal" w:id="{position() + 1}">
            <!-- warn if list in note -->
            <xsl:if test="text:note-body/descendant-or-self::text:list">
              <xsl:message terminate="no">translation.odf2oox.noteInList</xsl:message>
            </xsl:if>
            <xsl:apply-templates select="text:note-body"/>
          </w:endnote>
        </xsl:for-each>
      </xsl:for-each>

    </w:endnotes>
  </xsl:template>

  <!-- endotes configuration -->
  <xsl:template match="text:notes-configuration[@text:note-class='endnote']" mode="note">
    <xsl:param name="wide">no</xsl:param>
    <w:endnotePr>

      <xsl:choose>
        <xsl:when test="ancestor::style:style[@style:family='section']">
          <w:pos w:val="sectEnd"/>
        </xsl:when>
        <xsl:otherwise>
          <w:pos w:val="docEnd"/>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:if test="@style:num-format">
        <w:numFmt>
          <xsl:attribute name="w:val">
            <xsl:call-template name="GetNumFormat">
              <xsl:with-param name="format" select="@style:num-format"/>
            </xsl:call-template>
          </xsl:attribute>
        </w:numFmt>
      </xsl:if>

      <xsl:if test="@text:start-value">
        <xsl:choose>
          <xsl:when test="ancestor::style:style[@style:family='section']">
            <w:numStart w:val="{@text:start-value}"/>
          </xsl:when>
          <xsl:otherwise>
            <w:numStart w:val="{number(@text:start-value)+1}"/>
          </xsl:otherwise>
        </xsl:choose>
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
        <w:endnote w:id="0"/>
        <w:endnote w:id="1"/>
      </xsl:if>

    </w:endnotePr>
  </xsl:template>

  
  <!-- Reference from the document -->
  <xsl:template match="text:note[@text:note-class='endnote']" mode="text">
    <w:endnoteReference>
      <xsl:attribute name="w:id">
        <xsl:call-template name="GenerateId">
          <xsl:with-param name="node" select="."/>
          <xsl:with-param name="nodetype" select="@text:note-class"/>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:if test="text:note-citation/@text:label">
        <xsl:attribute name="w:customMarkFollows">1</xsl:attribute>
      </xsl:if>
    </w:endnoteReference>
    <xsl:if test="text:note-citation/@text:label">
      <w:t>
        <xsl:value-of select="text:note-citation"/>
      </w:t>
    </xsl:if>
  </xsl:template>

  
  <xsl:template name="InsertEndnotesInternalRelationships">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <xsl:for-each select="document('content.xml')">

        <!-- hyperlinks relationships. Do not pick up hyperlinks other than those coming from footnotes.  -->
        <xsl:call-template name="InsertHyperlinksRelationships">
          <xsl:with-param name="hyperlinks"
            select="key('hyperlinks', '')[ancestor::text:note/@text:note-class = 'endnote' ]"/>
        </xsl:call-template>

        <xsl:call-template name="InsertImagesRelationships">
          <xsl:with-param name="images"
            select="key('images', '')[ancestor::text:note/@text:note-class = 'endnote' ]"/>
        </xsl:call-template>

      </xsl:for-each>
    </Relationships>
  </xsl:template>

</xsl:stylesheet>

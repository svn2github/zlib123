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
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">

  <xsl:template name="part_relationships">

    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">

      <!--
			<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>
			
			<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
			-->

      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
        Target="numbering.xml"/>
      <Relationship Id="rId2"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
        Target="styles.xml"/>
      <Relationship Id="rId3"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
        Target="fontTable.xml"/>
      <Relationship Id="rId4"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
        Target="settings.xml"/>
      <Relationship Id="rId5"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
        Target="footnotes.xml"/>
      <Relationship Id="rId6"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
        Target="endnotes.xml"/>

      <Relationship Id="rId7"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
        Target="comments.xml"/>
      
      <!-- Relationship for header and footer -->
      <xsl:for-each
        select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page">
        <xsl:if test="style:header">  
          <Relationship Id="{generate-id(style:header)}" 
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
            Target="header{position()}.xml"/>
        </xsl:if>
        <xsl:if test="style:footer">
          <Relationship  Id="{generate-id(style:footer)}" 
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
            Target="footer{position()}.xml"/>
        </xsl:if>
      </xsl:for-each>

      <!-- OLE Objects -->
      <xsl:for-each select="document('content.xml')">
        <xsl:for-each select="key('ole-objects', '')[not(ancestor::text:note)]">
          <pzip:copy pzip:source="{substring-after(draw:object-ole/@xlink:href,'./')}"
            pzip:target="word/embeddings/{translate(concat(substring-after(draw:object-ole/@xlink:href,'./'),'.bin'),' ','')}"/>
          <Relationship Id="{generate-id(draw:object-ole)}"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
            Target="embeddings/{translate(concat(substring-after(draw:object-ole/@xlink:href,'./'),'.bin'),' ','')}"/>
          <xsl:if test="draw:image">
            <pzip:copy pzip:source="{substring-after(draw:image/@xlink:href,'./')}"
              pzip:target="word/media/{translate(concat(substring-after(draw:image/@xlink:href,'ObjectReplacements/'),'.wmf'),' ','')}"/>
            <Relationship Id="{generate-id(draw:image)}"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
              Target="media/{translate(concat(substring-after(draw:image/@xlink:href,'ObjectReplacements/'),'.wmf'),' ','')}"
            />
          </xsl:if>
        </xsl:for-each>
      </xsl:for-each>

      <xsl:for-each select="document('content.xml')">
        <xsl:for-each select="key('images', '')[not(ancestor::text:note)]">
          <xsl:variable name="supported">
            <xsl:call-template name="image-support">
              <xsl:with-param name="name" select="draw:image/@xlink:href"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:if test="$supported = 'true' ">
            <xsl:choose>
              <!-- Internal image -->
              <xsl:when test="starts-with(draw:image/@xlink:href, 'Pictures/')">
                <!-- copy this image to the oox package -->
                <pzip:copy pzip:source="{draw:image/@xlink:href}"
                  pzip:target="word/media/{substring-after(draw:image/@xlink:href, 'Pictures/')}"/>
                <Relationship Id="{generate-id(draw:image)}"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                  Target="media/{substring-after(draw:image/@xlink:href, 'Pictures/')}"/>
                <Relationship Id="{generate-id(ancestor::draw:a)}"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                  TargetMode="External"
                  Target="{ancestor::draw:a/@xlink:href}"/>
              </xsl:when>
              <xsl:otherwise>
                <!-- External image : If relative path, image may not be converted. -->
                <Relationship>
                  <xsl:attribute name="Id">
                    <xsl:value-of select="generate-id(draw:image)"/>
                  </xsl:attribute>
                  <xsl:attribute name="Type"
                    >http://schemas.openxmlformats.org/officeDocument/2006/relationships/image</xsl:attribute>
                  <xsl:attribute name="Target">
                    <xsl:choose>
                      <xsl:when test="contains(draw:image/@xlink:href,'./')">
                        <xsl:value-of select="concat('../../',draw:image/@xlink:href)"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="draw:image/@xlink:href"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
                  <xsl:attribute name="TargetMode">External</xsl:attribute>
                </Relationship>
                <Relationship Id="{generate-id(ancestor::draw:a)}"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                  TargetMode="External"
                  Target="{ancestor::draw:a/@xlink:href}"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </xsl:for-each>
      <!-- hyperlinks relationships. 
				Do not pick up hyperlinks from notes (footnotes or endnotes).
				TODO : really needs a clean way to find the text:a back! 
			-->
      <xsl:for-each select="document('content.xml')">
        <xsl:for-each select="key('hyperlinks', '')[not(ancestor::text:note)]">
          <Relationship Id="{generate-id()}"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
            TargetMode="External">
            <xsl:attribute name="Target">
              <!-- having Target empty makes Word Beta 2007 crash -->
              <xsl:choose>
                <xsl:when test="string-length(@xlink:href) &gt; 0">
                  <xsl:value-of select="@xlink:href"/>
                </xsl:when>
                <xsl:otherwise>/</xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </Relationship>
        </xsl:for-each>
      </xsl:for-each>

    </Relationships>

  </xsl:template>

</xsl:stylesheet>

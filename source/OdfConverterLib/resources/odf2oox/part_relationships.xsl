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
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  exclude-result-prefixes="w r xlink office draw text style">

  <xsl:template name="InsertPartRelationships">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <!--  Static relationships -->
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

      <!-- Headers/Footers relationships -->
      <xsl:call-template name="InsertHeaderFooterRelationships"/>

      <!-- OLE objects relationships -->
      <xsl:for-each select="document('content.xml')">
        <xsl:call-template name="InsertOleObjectsRelationships">
          <xsl:with-param name="oleObjects"
            select="key('ole-objects', '')[not(ancestor::text:note)]"/>
        </xsl:call-template>
      </xsl:for-each>

      <!-- Images relationships -->
      <xsl:for-each select="document('content.xml')">
        <xsl:call-template name="InsertImagesRelationships">
          <xsl:with-param name="images" select="key('images', '')[not(ancestor::text:note)]"/>
        </xsl:call-template>
      </xsl:for-each>

      <!-- Hyperlinks relationships -->
      <xsl:for-each select="document('content.xml')">
        <xsl:call-template name="InsertHyperlinksRelationships">
          <xsl:with-param name="hyperlinks" select="key('hyperlinks', '')[not(ancestor::text:note)]"
          />
        </xsl:call-template>
      </xsl:for-each>
    </Relationships>
  </xsl:template>


  <!-- OLE Objects relationships -->
  <xsl:template name="InsertOleObjectsRelationships">
    <xsl:param name="oleObjects"/>

    <xsl:for-each select="$oleObjects">
      <xsl:variable name="oleName"
        select="translate(concat(substring-after(draw:object-ole/@xlink:href,'./'),'.bin'),' ','')"/>
      <pzip:copy pzip:source="{substring-after(draw:object-ole/@xlink:href,'./')}"
        pzip:target="word/embeddings/{$oleName}"/>
      <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
        Id="{generate-id(draw:object-ole)}"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
        Target="embeddings/{$oleName}"/>

      <xsl:if test="draw:image">
        <xsl:variable name="objectReplacement"
          select="translate(concat(substring-after(draw:image/@xlink:href,'ObjectReplacements/'),'.wmf'),' ','')"/>
        <pzip:copy pzip:source="{substring-after(draw:image/@xlink:href,'./')}"
          pzip:target="word/media/{$objectReplacement}"/>
        <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
          Id="{generate-id(draw:image)}"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
          Target="media/{$objectReplacement}"/>
        <xsl:if test="ancestor::draw:a">
          <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
            Id="{generate-id(ancestor::draw:a)}"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
            TargetMode="External" Target="{ancestor::draw:a/@xlink:href}"/>
        </xsl:if>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>


  <!-- Images relationships -->
  <xsl:template name="InsertImagesRelationships">
    <xsl:param name="images"/>

    <xsl:for-each select="$images">
      <xsl:variable name="supported">
        <xsl:call-template name="image-support">
          <xsl:with-param name="name" select="@xlink:href"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:if test="@xlink:href and $supported = 'true' ">
        <xsl:choose>
          <!-- Internal image -->
          <xsl:when test="starts-with(@xlink:href, 'Pictures/')">
            <!-- copy this image to the oox package -->
            <xsl:variable name="imageName"
              select="substring-after(@xlink:href, 'Pictures/')"/>
            <pzip:copy pzip:source="{@xlink:href}" pzip:target="word/media/{$imageName}"/>
            <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
              Id="{generate-id(.)}"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
              Target="media/{$imageName}"/>
            <xsl:if test="ancestor::draw:a">
              <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
                Id="{generate-id(ancestor::draw:a)}"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                TargetMode="External">
                <xsl:attribute name="Target">
                  <xsl:choose>

                    <!-- converting relative path -->
                    <xsl:when test="starts-with(ancestor::draw:a/@xlink:href, '../')">
                      <xsl:call-template name="HandlingSpaces">
                        <xsl:with-param name="path">
                          <xsl:value-of select="substring-after(ancestor::draw:a/@xlink:href,'../')"
                          />
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:when>
                    
                    <xsl:when test="starts-with(ancestor::draw:a/@xlink:href, '/')">
                      <xsl:call-template name="HandlingSpaces">
                        <xsl:with-param name="path">
                          <xsl:value-of select="substring-after(ancestor::draw:a/@xlink:href,'/')"
                          />
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:when>
                    
                    <xsl:otherwise>
                      <xsl:call-template name="HandlingSpaces">
                        <xsl:with-param name="path">
                          <xsl:value-of select="ancestor::draw:a/@xlink:href"/>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
              </Relationship>
            </xsl:if>
          </xsl:when>
          <xsl:otherwise>

            <!-- External image : If relative path, image may not be converted. -->
            <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
              Id="{generate-id(.)}"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image">
              <xsl:attribute name="Target">
                <xsl:choose>
                  <xsl:when test="contains(@xlink:href, './')">
                    <xsl:value-of select="concat('../../', @xlink:href)"/>
                  </xsl:when>
                  <xsl:when test="starts-with(@xlink:href, '/')">
                    <xsl:value-of select="substring-after(@xlink:href, '/')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="@xlink:href"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
              <xsl:attribute name="TargetMode">External</xsl:attribute>
            </Relationship>
            <xsl:if test="ancestor::draw:a">
              <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
                Id="{generate-id(ancestor::draw:a)}"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                TargetMode="External" Target="{ancestor::draw:a/@xlink:href}"/>
            </xsl:if>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

  <!--handling spaces in paths-->

  <xsl:template name="HandlingSpaces">
    <xsl:param name="path"/>
    <xsl:choose>
      <xsl:when test="contains($path,' ')">
        <xsl:variable name="subPath">
          <xsl:call-template name="HandlingSpaces">
            <xsl:with-param name="path" select="substring-after($path,' ')"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:value-of select="concat(substring-before($path,' '),'%20',$subPath)"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$path"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>



  <!-- Hyperlinks relationships -->
  <xsl:template name="InsertHyperlinksRelationships">
    <xsl:param name="hyperlinks"/>

    <xsl:for-each select="$hyperlinks">
      <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
        Id="{generate-id()}"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
        TargetMode="External">
        <xsl:attribute name="Target">
          <!-- having Target empty makes Word Beta 2007 crash -->
          <xsl:choose>
              <xsl:when test="contains(@xlink:href, './')">
                <xsl:variable name="substring" select="substring-after(@xlink:href, '../')"/>
                <xsl:choose>
                  <xsl:when test="string-length($substring) = 0">/</xsl:when>
                  <xsl:otherwise><xsl:value-of select="$substring"/></xsl:otherwise>
                </xsl:choose>
              </xsl:when>
              <xsl:when test="string-length(@xlink:href) &gt; 0">
              <xsl:value-of select="@xlink:href"/>
            </xsl:when>
            <xsl:otherwise>/</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </Relationship>
    </xsl:for-each>
  </xsl:template>

  
  <!-- Headers / footers relationships -->
  <xsl:template name="InsertHeaderFooterRelationships">
    <xsl:variable name="masterPages"
      select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page"/>
    
    <xsl:for-each select="$masterPages/style:header | $masterPages/style:header-left">
      <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
        Id="{generate-id()}"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
        Target="header{position()}.xml"/>
    </xsl:for-each>
    <xsl:for-each select="$masterPages/style:footer | $masterPages/style:footer-left">
      <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
        Id="{generate-id()}"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
        Target="footer{position()}.xml"/>
    </xsl:for-each>
  </xsl:template>
  

</xsl:stylesheet>

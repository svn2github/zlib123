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
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip" xmlns:xlink="http://www.w3.org/1999/xlink"
  exclude-result-prefixes="text style office pzip xlink draw">

  <xsl:template name="header">
    <xsl:param name="headerNode"/>
    <w:hdr>
      <xsl:apply-templates select="$headerNode"/>
    </w:hdr>
  </xsl:template>

  <xsl:template name="footer">
    <xsl:param name="footerNode"/>
    <w:ftr>
      <xsl:apply-templates select="$footerNode"/>
      <!--
      <xsl:for-each select="child::node()">
        <xsl:choose>
          <xsl:when test="./descendant::text:page-number">
            <w:sdt>
              <w:sdtContent>
                <w:p>
                  <w:pPr>
                    <w:jc>
                      <xsl:attribute name="w:val">
                        <xsl:variable name="style" select="descendant::draw:frame/@draw:style-name"/>
                        <xsl:variable name="alignment" select="//office:document-styles/office:automatic-styles/style:style[@style:name= $style]/style:graphic-properties/@style:horizontal-pos"/>
                       
                        <xsl:choose>
                          <xsl:when test="($alignment = 'center') or ($alignment = 'right') or ($alignment = 'left')">
                            <xsl:value-of select="$alignment"/>
                          </xsl:when>
                          <xsl:otherwise>center</xsl:otherwise>
                        </xsl:choose>
                      </xsl:attribute>
                    </w:jc>
                  </w:pPr>
                  <w:fldSimple w:instr=" PAGE   \* MERGEFORMAT "/>
                </w:p>
              </w:sdtContent>
            </w:sdt>
          </xsl:when>
          <xsl:otherwise>
            <xsl:apply-templates select="."/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
      -->
    </w:ftr>
  </xsl:template>

  <!-- Look for even and odd definitions -->
  <xsl:template name="isOddEven">
    <xsl:param name="pos">0</xsl:param>
    <xsl:param name="masterPage"
      select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page"/>
    <xsl:variable name="mp" select="$masterPage[$pos]"/>
    <xsl:choose>
      <xsl:when test="$mp/style:header-left or $mp/style:footer-left">
        <xsl:value-of select="$mp/@style:name"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="$mp/following-sibling::*">
          <xsl:call-template name="isOddEven">
            <xsl:with-param name="pos" select="$pos + 1"/>
          </xsl:call-template>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- Header/footer document references -->
  <xsl:template name="HeaderFooter">
    <xsl:param name="master-page"/>

    <xsl:choose>
      <xsl:when test="$master-page/@style:name = 'First_20_Page' ">
        <xsl:if test="$master-page/style:header">
          <w:headerReference w:type="first" r:id="{generate-id($master-page/style:header)}"/>
        </xsl:if>
        <xsl:if test="$master-page/style:footer">
          <w:footerReference w:type="first" r:id="{generate-id($master-page/style:footer)}"/>
        </xsl:if>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$master-page/style:header-left and $master-page/style:header">
            <w:headerReference w:type="default" r:id="{generate-id($master-page/style:header)}"/>
            <w:headerReference w:type="even" r:id="{generate-id($master-page/style:header-left)}"/>
          </xsl:when>
          <xsl:when test="$master-page/style:header">
            <w:headerReference w:type="default" r:id="{generate-id($master-page/style:header)}"/>
          </xsl:when>
        </xsl:choose>
        <xsl:choose>
          <xsl:when test="$master-page/style:footer and $master-page/style:footer-left">
            <w:footerReference w:type="default" r:id="{generate-id($master-page/style:footer)}"/>
            <w:footerReference w:type="even" r:id="{generate-id($master-page/style:footer-left)}"/>
          </xsl:when>
          <xsl:when test="$master-page/style:footer">
            <w:footerReference w:type="default" r:id="{generate-id($master-page/style:footer)}"/>
          </xsl:when>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <xsl:template name="InsertHeaderFooterSettings">
    <xsl:variable name="oddPage">
      <xsl:call-template name="isOddEven"/>
    </xsl:variable>
    <xsl:if test="string-length($oddPage)>0">
      <w:evenAndOddHeaders/>
    </xsl:if>
  </xsl:template>

  <xsl:template match="text:page-number" mode="paragraph">
    <w:fldSimple w:instr=" PAGE ">
      <xsl:apply-templates mode="paragraph"/>
    </w:fldSimple>
  </xsl:template>

  <xsl:template match="text:page-count" mode="paragraph">
    <w:fldSimple w:instr=" NUMPAGES ">
      <xsl:apply-templates mode="paragraph"/>
    </w:fldSimple>
  </xsl:template>

  <!-- Headers/Footers part relationships construction -->
  <!-- param node : style:header / style:footer -->
  <xsl:template name="InsertHeaderFooterInternalRelationships">
    <xsl:param name="node"/>
    <xsl:variable name="masterPageName" select="$node/ancestor::style:master-page[1]/@style:name"/>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <xsl:for-each select="document('styles.xml')">

        <!-- hyperlinks -->
        <xsl:call-template name="InsertHyperlinksRelationships">
          <xsl:with-param name="hyperlinks"
            select="key('hyperlinks', '')[ancestor::style:master-page[@style:name=$masterPageName] and ancestor::*[name()=name($node)]]"
          />
        </xsl:call-template>

        <!-- OLE -->
        <xsl:call-template name="InsertOleObjectsRelationships">
          <xsl:with-param name="oleObjects"
            select="key('ole-objects', '')[ancestor::style:master-page[@style:name=$masterPageName] and ancestor::*[name()=name($node)]]"
          />
        </xsl:call-template>

        <!-- Images -->
        <xsl:call-template name="InsertImagesRelationships">
          <xsl:with-param name="images"
            select="key('images', '')[ancestor::style:master-page[@style:name=$masterPageName] and ancestor::*[name()=name($node)]]"
          />
        </xsl:call-template>

      </xsl:for-each>
    </Relationships>
  </xsl:template>


  <!-- header and footer definition files / part relationships files -->
  <xsl:template name="InsertHeaderFooterParts">
    <xsl:variable name="masterPages"
      select="document('styles.xml')/office:document-styles/office:master-styles/style:master-page"/>

    <xsl:for-each select="$masterPages/style:header | $masterPages/style:header-left">
      <pzip:entry>
        <xsl:attribute name="pzip:target">
          <xsl:value-of select="concat('word/header', position(), '.xml')"/>
        </xsl:attribute>
        <xsl:call-template name="header">
          <xsl:with-param name="headerNode" select="."/>
        </xsl:call-template>
      </pzip:entry>
      <pzip:entry>
        <xsl:attribute name="pzip:target">
          <xsl:value-of select="concat('word/_rels/header', position(), '.xml.rels')"/>
        </xsl:attribute>
        <xsl:call-template name="InsertHeaderFooterInternalRelationships">
          <xsl:with-param name="node" select="."/>
        </xsl:call-template>
      </pzip:entry>
    </xsl:for-each>

    <xsl:for-each select="$masterPages/style:footer | $masterPages/style:footer-left">
      <pzip:entry>
        <xsl:attribute name="pzip:target">
          <xsl:value-of select="concat('word/footer', position(), '.xml')"/>
        </xsl:attribute>
        <xsl:call-template name="footer">
          <xsl:with-param name="footerNode" select="."/>
        </xsl:call-template>
      </pzip:entry>
      <pzip:entry>
        <xsl:attribute name="pzip:target">
          <xsl:value-of select="concat('word/_rels/footer', position(), '.xml.rels')"/>
        </xsl:attribute>
        <xsl:call-template name="InsertHeaderFooterInternalRelationships">
          <xsl:with-param name="node" select="."/>
        </xsl:call-template>
      </pzip:entry>
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

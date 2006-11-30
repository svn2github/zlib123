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
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  exclude-result-prefixes="w r draw number">
  
  <xsl:import href="footnotes.xsl"/>
  <xsl:key name="StyleId" match="w:style" use="@w:styleId"/>
  
  <xsl:template name="styles">
    <office:document-styles>
      <office:font-face-decls>
        <xsl:apply-templates select="document('word/fontTable.xml')/w:fonts"/>
      </office:font-face-decls>
      <!-- document styles -->
      <office:styles>
        <!-- document styles -->
        <xsl:apply-templates select="document('word/styles.xml')/w:styles"/>
        <xsl:call-template name="InsertNotesConfiguration"/>
      </office:styles>
      <!-- automatic styles -->
      <office:automatic-styles>
        <xsl:call-template name="HeaderFooterAutomaticStyle"/>
        <!-- TODO : create other automatic styles. This one handles only the default (last w:sectPr of document.xml). -->
        <xsl:if test="document('word/document.xml')/w:document/w:body/w:sectPr">
          <xsl:call-template name="InsertDefaultPageLayout"/>
        </xsl:if>
      </office:automatic-styles>
      <!-- master styles -->
      <office:master-styles>
        <!-- TODO : create other master-page styles. This one handles only the default (last w:sectPr of document.xml). -->
        <xsl:if test="document('word/document.xml')/w:document/w:body/w:sectPr">
          <xsl:call-template name="InsertDefaultMasterPage"/>
        </xsl:if>
      </office:master-styles>
    </office:document-styles>
  </xsl:template>


  <xsl:template name="HeaderFooterAutomaticStyle">
    <xsl:for-each select="document('word/document.xml')/w:document/w:body/w:sectPr">
      <xsl:for-each select="w:headerReference">
        <xsl:if test="./@w:type = 'default'">
          <xsl:variable name="headerId" select="./@r:id"/>
          <xsl:variable name="headerXmlDocument"
            select="concat('word/',document('word/_rels/document.xml.rels')/descendant::node()[@Id=$headerId]/@Target)"/>
          <!-- change context to get footer content -->
          <xsl:for-each select="document($headerXmlDocument)">
            <xsl:apply-templates mode="automaticstyles"/>
            <xsl:if test="document('word/document.xml')/w:document[descendant::w:numPr/w:numId] 
              or document('word/styles.xml')/w:styles/w:style[descendant::w:numPr/w:numId]">
              <xsl:apply-templates select="document('word/numbering.xml')/w:numbering/w:num"/>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
      </xsl:for-each>
      <xsl:for-each select="w:headerReference">
        <xsl:if test="./@w:type = 'even'">
          <xsl:variable name="headerId" select="./@r:id"/>
          <xsl:variable name="headerXmlDocument"
            select="concat('word/',document('word/_rels/document.xml.rels')/descendant::node()[@Id=$headerId]/@Target)"/>
          <!-- change context to get footer content -->
          <xsl:for-each select="document($headerXmlDocument)">
            <xsl:apply-templates mode="automaticstyles"/>
            <xsl:if test="document('word/document.xml')/w:document[descendant::w:numPr/w:numId] 
              or document('word/styles.xml')/w:styles/w:style[descendant::w:numPr/w:numId]">
              <xsl:apply-templates select="document('word/numbering.xml')/w:numbering/w:num"/>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
      </xsl:for-each>
      <xsl:for-each select="w:footerReference">
        <xsl:if test="./@w:type = 'default'">
          <xsl:variable name="footerId" select="./@r:id"/>
          <xsl:variable name="footerXmlDocument"
            select="concat('word/',document('word/_rels/document.xml.rels')/descendant::node()[@Id=$footerId]/@Target)"/>
          <!-- change context to get header content -->
          <xsl:for-each select="document($footerXmlDocument)">
            <xsl:apply-templates mode="automaticstyles"/>
            <xsl:if test="document('word/document.xml')/w:document[descendant::w:numPr/w:numId] 
              or document('word/styles.xml')/w:styles/w:style[descendant::w:numPr/w:numId]">
              <xsl:apply-templates select="document('word/numbering.xml')/w:numbering/w:num"/>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
      </xsl:for-each>
      <xsl:for-each select="w:footerReference">
        <xsl:if test="./@w:type = 'even'">
          <xsl:variable name="footerId" select="./@r:id"/>
          <xsl:variable name="footerXmlDocument"
            select="concat('word/',document('word/_rels/document.xml.rels')/descendant::node()[@Id=$footerId]/@Target)"/>
          <!-- change context to get header content -->
          <xsl:for-each select="document($footerXmlDocument)">
            <xsl:apply-templates mode="automaticstyles"/>
            <xsl:if test="document('word/document.xml')/w:document[descendant::w:numPr/w:numId] 
              or document('word/styles.xml')/w:styles/w:style[descendant::w:numPr/w:numId]">
              <xsl:apply-templates select="document('word/numbering.xml')/w:numbering/w:num"/>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
      </xsl:for-each>
    </xsl:for-each>
  </xsl:template>
  
  
  <!-- handle default master page style -->
  <xsl:template name="InsertDefaultMasterPage">
    <style:master-page style:name="Standard" style:page-layout-name="pm1">
      <xsl:for-each select="document('word/document.xml')/w:document/w:body/w:sectPr">
        <xsl:call-template name="HeaderFooter"/>
      </xsl:for-each>
    </style:master-page>
      <xsl:for-each select="document('word/document.xml')/w:document/w:body/w:p/w:pPr/w:sectPr">
        <xsl:if test="w:headerReference/@w:type = 'default' or w:headerReference/@w:type = 'even' or w:footerReference/@w:type = 'default' or w:footerReference/@w:type = 'even'">
          <xsl:choose>
            <xsl:when test="w:headerReference/@w:type = 'first' or w:footerReference/@w:type = 'first'">
              <style:master-page>
                <xsl:attribute name="style:name">
                  <xsl:value-of select="concat('First_H_',generate-id(.))"/>
                </xsl:attribute>
                <xsl:attribute name="style:next-style-name">
                  <xsl:value-of select="concat('H_',generate-id(.))"/>
                </xsl:attribute>
                <xsl:attribute name="style:page-layout-name">
                  <xsl:text>pm1</xsl:text>
                </xsl:attribute>
                <xsl:attribute name="style:display-name">
                  <xsl:value-of select="concat('H_',generate-id(.))"/>
                </xsl:attribute>
                <xsl:call-template name="HeaderFooterFirst"/>
              </style:master-page>
            </xsl:when>
            <xsl:when test="document('word/document.xml')/w:document/w:body/w:sectPr/w:headerReference/@w:type = 'first' or document('word/document.xml')/w:document/w:body/w:sectPr/w:footerReference/@w:type = 'first'">
              <style:master-page>
                <xsl:attribute name="style:name">
                  <xsl:value-of select="concat('First_H_',generate-id(.))"/>
                </xsl:attribute>
                <xsl:attribute name="style:next-style-name">
                  <xsl:value-of select="concat('H_',generate-id(.))"/>
                </xsl:attribute>
                <xsl:attribute name="style:page-layout-name">
                  <xsl:text>pm1</xsl:text>
                </xsl:attribute>
                <xsl:attribute name="style:display-name">
                  <xsl:value-of select="concat('First_H_',generate-id(.))"/>
                </xsl:attribute>
                <xsl:for-each select="document('word/document.xml')/w:document/w:body/w:sectPr">
                  <xsl:call-template name="HeaderFooterFirst"/>
                </xsl:for-each>
              </style:master-page>
              
            </xsl:when>
          </xsl:choose>
          
          <style:master-page>
            <xsl:attribute name="style:name">
              <xsl:value-of select="concat('H_',generate-id(.))"/>
            </xsl:attribute>
            <xsl:attribute name="style:page-layout-name">
              <xsl:text>pm1</xsl:text>
            </xsl:attribute>
            <xsl:attribute name="style:display-name">
              <xsl:value-of select="concat('H_',generate-id(.))"/>
            </xsl:attribute>
            <xsl:call-template name="HeaderFooter"/>
          </style:master-page>
        </xsl:if>
      </xsl:for-each>
    <xsl:if test="document('word/document.xml')/w:document/w:body/w:sectPr/w:headerReference/@w:type = 'first' or document('word/document.xml')/w:document/w:body/w:sectPr/w:footerReference/@w:type = 'first'">
    <style:master-page style:name="First_Page" style:page-layout-name="pm1" style:next-style-name="Standard" style:display-name="First Page">
      <xsl:for-each select="document('word/document.xml')/w:document/w:body/w:sectPr">
        <xsl:call-template name="HeaderFooterFirst"/>
      </xsl:for-each>
    </style:master-page>
    </xsl:if>
    <xsl:for-each select="document('word/document.xml')/w:document/w:body/w:p/w:pPr/w:sectPr">
      <xsl:if test="preceding::w:sectPr/w:pgSz/@w:w != ./w:pgSz/@w:w or preceding::w:sectPr/w:pgSz/@w:h != ./w:pgSz/@w:h or preceding::w:sectPr/w:pgSz/@w:orient != ./w:pgSz/@w:orient ">
      <style:master-page>
        <xsl:attribute name="style:name">
          <xsl:value-of select="concat('PAGE_',generate-id(.))"/>
        </xsl:attribute>
        <xsl:attribute name="style:master-page-name">
          <xsl:value-of select="First_Page"/>
        </xsl:attribute>
      </style:master-page>
        </xsl:if>
    </xsl:for-each>
    
  </xsl:template>

  
  <xsl:template name="HeaderFooter">
    <xsl:for-each select="w:headerReference">
      <xsl:if test="./@w:type = 'default'">
        <style:header>
          <xsl:variable name="headerId" select="./@r:id"/>
          <xsl:variable name="headerXmlDocument"
            select="concat('word/',document('word/_rels/document.xml.rels')/descendant::node()[@Id=$headerId]/@Target)"/>
          <!-- change context to get header content -->
          <xsl:for-each select="document($headerXmlDocument)">
            <xsl:apply-templates/>
          </xsl:for-each>
        </style:header>
      </xsl:if>
    </xsl:for-each>
    <xsl:for-each select="w:headerReference">
      <xsl:if test="./@w:type = 'even'">
        <style:header-left>
          <xsl:variable name="headerId" select="./@r:id"/>
          <xsl:variable name="headerXmlDocument"
            select="concat('word/',document('word/_rels/document.xml.rels')/descendant::node()[@Id=$headerId]/@Target)"/>
          <!-- change context to get header content -->
          <xsl:for-each select="document($headerXmlDocument)">
            <xsl:apply-templates/>
          </xsl:for-each>
        </style:header-left>
      </xsl:if>
      <!-- ./@w:type='default' TODO -->
    </xsl:for-each>
    <xsl:for-each select="w:footerReference">
      <xsl:if test="./@w:type = 'default'">
        <style:footer>
          <xsl:variable name="footerId" select="./@r:id"/>
          <xsl:variable name="footerXmlDocument"
            select="concat('word/',document('word/_rels/document.xml.rels')/descendant::node()[@Id=$footerId]/@Target)"/>
          <!-- change context to get footer content -->
          <xsl:for-each select="document($footerXmlDocument)">
            <xsl:apply-templates/>
          </xsl:for-each>
        </style:footer>
      </xsl:if>
    </xsl:for-each>
    <xsl:for-each select="w:footerReference">
      <xsl:if test="./@w:type = 'even'">
        <style:footer-left>
          <xsl:variable name="footerId" select="./@r:id"/>
          <xsl:variable name="footerXmlDocument"
            select="concat('word/',document('word/_rels/document.xml.rels')/descendant::node()[@Id=$footerId]/@Target)"/>
          <!-- change context to get footer content -->
          <xsl:for-each select="document($footerXmlDocument)">
            <xsl:apply-templates/>
          </xsl:for-each>
        </style:footer-left>
      </xsl:if>
      <!-- ./@w:type='default' TODO -->
    </xsl:for-each>
  </xsl:template>
  
  <xsl:template name="HeaderFooterFirst">
    <xsl:for-each select="w:headerReference">
      <xsl:if test="./@w:type = 'first'">
        <style:header>
          <xsl:variable name="headerId" select="./@r:id"/>
          <xsl:variable name="headerXmlDocument"
            select="concat('word/',document('word/_rels/document.xml.rels')/descendant::node()[@Id=$headerId]/@Target)"/>
          <!-- change context to get header content -->
          <xsl:for-each select="document($headerXmlDocument)">
            <xsl:apply-templates/>
          </xsl:for-each>
        </style:header>
      </xsl:if>
    </xsl:for-each>
    <xsl:for-each select="w:footerReference">
      <xsl:if test="./@w:type = 'first'">
        <style:footer>
          <xsl:variable name="footerId" select="./@r:id"/>
          <xsl:variable name="footerXmlDocument"
            select="concat('word/',document('word/_rels/document.xml.rels')/descendant::node()[@Id=$footerId]/@Target)"/>
          <!-- change context to get footer content -->
          <xsl:for-each select="document($footerXmlDocument)">
            <xsl:apply-templates/>
          </xsl:for-each>
        </style:footer>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  
  
  <!-- handle default page layout -->
  <xsl:template name="InsertDefaultPageLayout">
    <style:page-layout style:name="pm1">
      <xsl:for-each select="document('word/document.xml')/w:document/w:body/w:sectPr">
        <style:page-layout-properties>
          <xsl:call-template name="InsertPageLayoutProperties"/>
          <xsl:call-template name="InsertColumns"/>
        </style:page-layout-properties>
        <style:header-style>
          <style:header-footer-properties>
            <xsl:call-template name="InsertHeaderFooterProperties">
              <xsl:with-param name="object">header</xsl:with-param>
            </xsl:call-template>
          </style:header-footer-properties>
        </style:header-style>
        <style:footer-style>
          <style:header-footer-properties>
            <xsl:call-template name="InsertHeaderFooterProperties">
              <xsl:with-param name="object">footer</xsl:with-param>
            </xsl:call-template>
          </style:header-footer-properties>
        </style:footer-style>
      </xsl:for-each>
    </style:page-layout>
    <xsl:for-each select="document('word/document.xml')/w:document/w:body/w:p/w:pPr/w:sectPr">
      <xsl:if test="preceding::w:sectPr/w:pgSz/@w:w != ./w:pgSz/@w:w or preceding::w:sectPr/w:pgSz/@w:h != ./w:pgSz/@w:h or preceding::w:sectPr/w:pgSz/@w:orient != ./w:pgSz/@w:orient ">
      <style:page-layout>
        <xsl:attribute name="style:name">
          <xsl:value-of select="concat('PAGE',generate-id(.))"/>
        </xsl:attribute>
        <style:page-layout-properties>
          <xsl:call-template name="InsertPageLayoutProperties"/>
        </style:page-layout-properties>
      </style:page-layout>
        </xsl:if>
    </xsl:for-each>
  </xsl:template>
    
  <!-- conversion of page properties -->
  <!-- TODO : handle other properties -->
  <xsl:template name="InsertPageLayoutProperties">

    <!-- page size -->
    <xsl:if test="w:pgSz">
      <xsl:attribute name="fo:page-width">
        <xsl:call-template name="ConvertTwips">
          <xsl:with-param name="length" select="w:pgSz/@w:w"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="fo:page-height">
        <xsl:call-template name="ConvertTwips">
          <xsl:with-param name="length" select="w:pgSz/@w:h"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:if test="w:pgSz/@w:orient">
        <xsl:attribute name="style:print-orientation">
          <xsl:value-of select="w:pgSz/@w:orient"/>
        </xsl:attribute>
      </xsl:if>
    </xsl:if>

    <!-- page margins -->
    <xsl:if test="w:pgMar">
      <xsl:call-template name="ComputePageMargins"/>
    </xsl:if>

    <!--  page color  -->
    <xsl:if test="//w:document//w:background/@w:color">
      <xsl:attribute name="fo:background-color">
        <xsl:text>#</xsl:text>
        <xsl:value-of select="//w:document//w:background/@w:color"/>
      </xsl:attribute>
    </xsl:if>

    <!-- page numbering style. -->
    <!-- TODO : handle chapter numbering -->
    <xsl:if test="w:pgNumType">
      <xsl:call-template name="InsertPageNumbering"/>
    </xsl:if>

    
    
  </xsl:template>

  <!-- page margins -->
  <xsl:template name="ComputePageMargins">
    <xsl:attribute name="fo:margin-top">
      <xsl:call-template name="ConvertTwips">
        <xsl:with-param name="unit">cm</xsl:with-param>
        <xsl:with-param name="length">
          <xsl:choose>
            <xsl:when test="document('word/document.xml')//w:document/w:body/w:sectPr/w:headerReference[@w:type='default' or @w:type='even']">
              <xsl:choose>
                <xsl:when test="w:pgMar/@w:top &lt; 0">
                  <xsl:value-of select="w:pgMar/@w:top"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="w:pgMar/@w:header"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when test="document('word/settings.xml')//w:settings/w:gutterAtTop">
                  <xsl:choose>
                    <xsl:when test="w:pgMar/@w:top &lt; 0">
                      <xsl:value-of select="w:pgMar/@w:top + w:pgMar/@w:gutter"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:choose>
                        <xsl:when test="w:pgMar/@w:top + w:pgMar/@w:gutter &lt; w:pgMar/@w:header">
                          <xsl:value-of select="w:pgMar/@w:header"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="w:pgMar/@w:top + w:pgMar/@w:gutter"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:choose>
                    <xsl:when test="w:pgMar/@w:top &lt; 0">
                      <xsl:value-of select="w:pgMar/@w:top"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:choose>
                        <xsl:when test="w:pgMar/@w:top &lt; w:pgMar/@w:header">
                          <xsl:value-of select="w:pgMar/@w:header"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="w:pgMar/@w:top"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="fo:margin-left">
      <xsl:call-template name="ConvertTwips">
        <xsl:with-param name="length">
          <xsl:choose>
            <xsl:when test="document('word/settings.xml')//w:settings/w:gutterAtTop">
              <xsl:value-of select="w:pgMar/@w:left"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="w:pgMar/@w:left + w:pgMar/@w:gutter"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="fo:margin-bottom">
      <xsl:call-template name="ConvertTwips">
        <xsl:with-param name="unit">cm</xsl:with-param>
        <xsl:with-param name="length">
          <xsl:choose>
            <xsl:when test="document('word/document.xml')//w:document/w:body/w:sectPr/w:footerReference[@w:type='default' or @w:type='even']">
              <xsl:choose>
                <xsl:when test="w:pgMar/@w:bottom &lt; 0">
                  <xsl:value-of select="w:pgMar/@w:bottom"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="w:pgMar/@w:footer"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when test="w:pgMar/@w:bottom &lt; 0">
                  <xsl:value-of select="w:pgMar/@w:bottom"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:choose>
                    <xsl:when test="w:pgMar/@w:bottom &lt; w:pgMar/@w:footer">
                      <xsl:value-of select="w:pgMar/@w:footer"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="w:pgMar/@w:bottom"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="fo:margin-right">
      <xsl:call-template name="ConvertTwips">
        <xsl:with-param name="length" select="w:pgMar/@w:right"/>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>

  <!-- page numbers -->
  <xsl:template name="InsertPageNumbering">
    <xsl:if test="w:pgNumType/@w:fmt">
      <!-- Most number format are lost. -->
      <xsl:choose>
        <xsl:when
          test="contains(w:pgNumType/@w:fmt,'decimal') or w:pgNumType/@w:fmt = 'numberInDash' or w:pgNumType/@w:fmt = 'ordinal' ">
          <!-- prefix and suffix -->
          <xsl:choose>
            <xsl:when test="w:pgNumType/@w:fmt = 'numberInDash' ">
              <xsl:attribute name="style:num-prefix">- </xsl:attribute>
              <xsl:attribute name="style:num-suffix"> -</xsl:attribute>
            </xsl:when>
            <xsl:when test="w:pgNumType/@w:fmt = 'decimalEnclosedParen' ">
              <xsl:attribute name="style:num-prefix">(</xsl:attribute>
              <xsl:attribute name="style:num-suffix">)</xsl:attribute>
            </xsl:when>
            <xsl:when test="w:pgNumType/@w:fmt = 'decimalEnclosedFullstop' ">
              <xsl:attribute name="style:num-suffix">.</xsl:attribute>
            </xsl:when>
            <xsl:when test="contains(w:pgNumType/@w:fmt,'decimalEnclosedCircle')">
              <xsl:attribute name="style:num-prefix">(</xsl:attribute>
              <xsl:attribute name="style:num-suffix">)</xsl:attribute>
            </xsl:when>
            <xsl:otherwise/>
          </xsl:choose>
          <!-- start number -->
          <xsl:choose>
            <xsl:when test="w:pgNumType/@w:start">
              <xsl:attribute name="style:num-format">
                <xsl:value-of select="w:pgNumType/@w:start"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>1</xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="w:pgNumType/@w:fmt = 'lowerLetter' ">
          <xsl:attribute name="style:num-format">a</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:pgNumType/@w:fmt = 'upperLetter' ">
          <xsl:attribute name="style:num-format">A</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:pgNumType/@w:fmt = 'lowerRoman' ">
          <xsl:attribute name="style:num-format">i</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:pgNumType/@w:fmt = 'upperRoman' ">
          <xsl:attribute name="style:num-format">I</xsl:attribute>
        </xsl:when>
        <xsl:otherwise/>
      </xsl:choose>
    </xsl:if>
  </xsl:template>

  <!-- header / footer properties -->
  <!-- TODO : handle other properties -->
  <xsl:template name="InsertHeaderFooterProperties">
    <xsl:param name="object"/>

    <!-- dynamic spacing always false : no spacing defined in OOX -->
    <xsl:attribute name="style:dynamic-spacing">false</xsl:attribute>

    <!-- no spacing in OOX. -->
    <xsl:choose>
      <xsl:when test="$object = 'header' ">
        <xsl:attribute name="fo:margin-bottom">0cm</xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="fo:margin-top">0cm</xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:attribute name="fo:margin-left">0cm</xsl:attribute>
    <xsl:attribute name="fo:margin-right">0cm</xsl:attribute>

    <!-- min height calculated with page Margins properties. -->
    <xsl:attribute name="fo:min-height">
      <xsl:call-template name="ConvertTwips">
        <xsl:with-param name="unit">cm</xsl:with-param>
        <xsl:with-param name="length">
          <xsl:choose>
            <xsl:when test="$object = 'header' ">
              <xsl:choose>
                <xsl:when test="document('word/document.xml')//w:document/w:body/w:sectPr/w:headerReference[@w:type='default' or @w:type='even']">
                  <xsl:choose>
                    <xsl:when test="w:pgMar/@w:top &lt; 0">
                      <xsl:choose>
                        <xsl:when test="document('word/settings.xml')//w:settings/w:gutterAtTop">
                          <xsl:choose>
                            <xsl:when test=" - w:pgMar/@w:top + w:pgMar/@w:gutter &lt; w:pgMar/@w:header">0</xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select=" - w:pgMar/@w:top + w:pgMar/@w:gutter - w:pgMar/@w:header"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:choose>
                            <xsl:when test=" - w:pgMar/@w:top &lt; w:pgMar/@w:header">0</xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select=" - w:pgMar/@w:top - w:pgMar/@w:header"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:choose>
                        <xsl:when test="document('word/settings.xml')//w:settings/w:gutterAtTop">
                          <xsl:choose>
                            <xsl:when test="w:pgMar/@w:top + w:pgMar/@w:gutter &lt; w:pgMar/@w:header">0</xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="w:pgMar/@w:top + w:pgMar/@w:gutter - w:pgMar/@w:header"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:choose>
                            <xsl:when test="w:pgMar/@w:top &lt; w:pgMar/@w:header">0</xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="w:pgMar/@w:top - w:pgMar/@w:header"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>          
            </xsl:when>
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when test="document('word/document.xml')//w:document/w:body/w:sectPr/w:footerReference[@w:type='default' or @w:type='even']">
                  <xsl:choose>
                    <xsl:when test="w:pgMar/@w:bottom &lt; 0">
                      <xsl:choose>
                        <xsl:when test=" - w:pgMar/@w:bottom &lt; w:pgMar/@w:footer">0</xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select=" - w:pgMar/@w:bottom - w:pgMar/@w:footer"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:choose>
                        <xsl:when test="w:pgMar/@w:bottom &lt; w:pgMar/@w:footer">0</xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="w:pgMar/@w:bottom - w:pgMar/@w:footer"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>

  <!-- create styles -->
  <xsl:template match="w:style">
    <xsl:message terminate="no">progress:w:style</xsl:message>
    <xsl:variable name="currentStyleId">
      <xsl:value-of select="self::node()/@w:styleId"/>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="self::node()/@w:default and not($currentStyleId='Normal')">
        <xsl:if test="@w:type='table'">
          <style:default-style style:family="table">
            <xsl:call-template name="InsertTableProperties"/>
          </style:default-style>
        </xsl:if>
        <!--@TODO style:family: graphic, text etc-->
      </xsl:when>
      
      <!-- do not add anchor and symbol styles if there are there allready-->
      <xsl:when test="(($currentStyleId='FootnoteReference' or $currentStyleId='EndnoteReference') and document('word/styles.xml')/descendant::w:style[@w:styleId = concat(substring-before($currentStyleId,'Reference'),'_20_anchor')]) or (($currentStyleId='FootnoteText' or $currentStyleId='EndnoteText') and document('word/styles.xml')/descendant::w:style[@w:styleId = concat(substring-before($currentStyleId,'Text'),'_20_Symbol')])"></xsl:when>
      
      <xsl:otherwise>
        <style:style>
          <xsl:attribute name="style:name">
            <xsl:choose>
              <!-- change reference style to anchor style -->
            <xsl:when test="$currentStyleId='FootnoteReference' or $currentStyleId='EndnoteReference'">
              <xsl:value-of select="concat(substring-before($currentStyleId,'Reference'),'_20_anchor')"/>
            </xsl:when>
              <!--change Text style to Symbol style -->
              <xsl:when test="$currentStyleId='FootnoteText' or $currentStyleId='EndnoteText'">
                <xsl:value-of select="concat(substring-before($currentStyleId,'Text'),'_20_Symbol')"/>
              </xsl:when>
  
              <xsl:otherwise>
                <xsl:value-of select="$currentStyleId"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:attribute name="style:display-name">
            <xsl:value-of select="self::node()/w:name/@w:val"/>
          </xsl:attribute>
          <xsl:attribute name="style:family">
            <xsl:choose>
              <xsl:when test="self::node()/@w:type = 'character' ">
                <xsl:text>text</xsl:text>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="self::node()/@w:type"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>

          <xsl:if test="w:basedOn">
            <xsl:attribute name="style:parent-style-name">
              <xsl:value-of select="w:basedOn/@w:val"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:call-template name="InsertStyleProperties"/>
        </style:style>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- document defaults -->
  <xsl:template match="w:docDefaults">
    <style:default-style style:family="paragraph">
      <xsl:if test="w:rPrDefault">
        <style:text-properties>
          <xsl:if test="w:rPrDefault/w:rPr">
            <xsl:for-each select="w:rPrDefault/w:rPr">
              <xsl:call-template name="InsertTextProperties"/>
            </xsl:for-each>
            <xsl:for-each select="w:rPrDefault/w:pPr">
              <xsl:call-template name="InsertpPrTextProperties"/>
            </xsl:for-each>
          </xsl:if>
        </style:text-properties>
      </xsl:if>
    </style:default-style>
  </xsl:template>
  
  <!-- Compute style and text properties of context style. -->
  <xsl:template name="InsertStyleProperties">
      
    <xsl:if test="self::node()/@w:type = 'paragraph'">
      <style:paragraph-properties>
        <xsl:if test="w:pPr">
          <xsl:for-each select="w:pPr">
            <xsl:call-template name="InsertParagraphProperties"/>
          </xsl:for-each>
        </xsl:if>
        
       <!-- add default paragraph propeties to Normal or Default style (it can't be in style:default-style becouse of OO bug ? (see #1593148 image partly lost))-->
        <xsl:if test="self::node()/@w:styleId = 'Normal' or self::node()/@w:styleId = 'Default'">
            <xsl:for-each select="ancestor::w:styles/w:docDefaults">
              <xsl:if test="w:pPrDefault/w:pPr">
                <xsl:for-each select="w:pPrDefault/w:pPr">
                  <xsl:call-template name="InsertParagraphProperties"/>
                </xsl:for-each>
              </xsl:if>
            </xsl:for-each>
        </xsl:if>
      </style:paragraph-properties>
    </xsl:if>

        <xsl:if test="self::node()/@w:type = 'paragraph' or self::node()/@w:type = 'character'">
      <style:text-properties>
        <xsl:if test="w:rPr">
          <xsl:for-each select="w:rPr">
            <xsl:call-template name="InsertTextProperties"/>
          </xsl:for-each>
          <xsl:for-each select="w:pPr">
            <xsl:call-template name="InsertpPrTextProperties"/>
          </xsl:for-each>
        </xsl:if>
      </style:text-properties>
    </xsl:if>

  </xsl:template>

  <!--   Default paragraph properties from settings -->
  <xsl:template name="InsertDefaultParagraphProperties">
    <!--default tab-stop-->
    <xsl:variable name="tabStop" select="document('word/settings.xml')/w:settings/w:defaultTabStop/@w:val"/>
    <xsl:attribute name="style:tab-stop-distance">
      <xsl:call-template name="ConvertTwips">
        <xsl:with-param name="length" select="$tabStop"/>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:attribute>
   </xsl:template>

  <!-- conversion of paragraph properties -->
  <xsl:template name="InsertParagraphProperties">

    <xsl:if test="w:keepNext">
      <xsl:attribute name="fo:keep-with-next">
        <xsl:choose>
          <xsl:when
            test="w:keepNext/@w:val='off' or w:keepNext/@w:val='false' or w:keepNext/@w:val=0">
            <xsl:value-of select="'auto'"/>
          </xsl:when>
          <xsl:when test="w:keepNext/@w:val='on' or w:keepNext/@w:val='true' or w:keepNext/@w:val=1">
            <xsl:value-of select="'always'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'always'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:keepLines">
      <xsl:attribute name="fo:keep-together">
        <xsl:choose>
          <xsl:when
            test="w:keepLines/@w:val='off' or w:keepLines/@w:val='false' or w:keepLines/@w:val=0">
            <xsl:value-of select="'auto'"/>
          </xsl:when>
          <xsl:when
            test="w:keepLines/@w:val='on' or w:keepLines/@w:val='true' or w:keepLines/@w:val=1">
            <xsl:value-of select="'always'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'always'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <!-- break before paragraph -->
    <xsl:if test="w:pageBreakBefore">
      <xsl:attribute name="fo:break-before">
        <xsl:choose>
          <xsl:when
            test="w:pageBreakBefore/@w:val='off' or w:pageBreakBefore/@w:val='false' or w:pageBreakBefore/@w:val=0">
            <xsl:value-of select="'auto'"/>
          </xsl:when>
          <xsl:when
            test="w:pageBreakBefore/@w:val='on' or w:pageBreakBefore/@w:val='true' or w:pageBreakBefore/@w:val=1">
            <xsl:value-of select="'page'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'page'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <!-- page breaks and simple column breaks -->
    <xsl:if test="parent::w:p/w:r/w:br[@w:type='page' or @w:type='column']">
      <xsl:attribute name="fo:break-after">
        <xsl:value-of select="parent::w:p/w:r/w:br[@w:type='page' or @w:type='column']/@w:type"/>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:r/w:br[@w:type='page' or @w:type='column']">
      <xsl:attribute name="fo:break-after">
        <xsl:value-of select="w:br[@w:type='page' or @w:type='column']/@w:type"/>
      </xsl:attribute>
    </xsl:if>
    
    <!-- borders. -->
    <xsl:if test="w:pBdr">
      <xsl:call-template name="InsertParagraphBorder"/>
      <!-- add shadow -->
      <xsl:call-template name="InsertParagraphShadow"/>
    </xsl:if>

    <!-- bg color -->
    <xsl:if test="w:shd">
      <xsl:attribute name="fo:background-color">
        <xsl:choose>
          <xsl:when test="w:shd/@w:fill='auto' or not(w:shd/@w:fill)">
            <xsl:value-of select="'transparent'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat('#',w:shd/@w:fill)"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <!-- text autospace -->
    <xsl:if test="w:autoSpaceDN or w:autoSpaceDE">
      <xsl:attribute name="style:text-autospace">
        <xsl:choose>
          <xsl:when
            test="w:autoSpaceDN/@w:val='off' or w:autoSpaceDN/@w:val='false' or w:autoSpaceDN/@w:val=0 or w:autoSpaceDE/@w:val='off' or w:autoSpaceDE/@w:val='false' or w:autoSpaceDE/@w:val=0">
            <xsl:value-of select="'none'"/>
          </xsl:when>
          <xsl:when
            test="w:autoSpaceDN/@w:val='on' or w:autoSpaceDN/@w:val='true' or w:autoSpaceDN/@w:val=1 or w:autoSpaceDE/@w:val='on' or w:autoSpaceDE/@w:val='true' or w:autoSpaceDE/@w:val=1">
            <xsl:value-of select="'ideograph-alpha'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'ideograph-alpha'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <!-- space before/after -->

<xsl:variable name="StyleId">
  <xsl:value-of select="w:pStyle/@w:val"/>
</xsl:variable>
 <xsl:variable name="Ivl">
   <xsl:value-of select="w:numPr/w:ilvl/@w:val"/>
 </xsl:variable>

    <xsl:variable name="NumId">
      <xsl:choose>
        <xsl:when test="w:numPr/w:numId/@w:val">
          <xsl:value-of select="w:numPr/w:numId/@w:val"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="document('word/document.xml')/w:document/w:body/w:p/w:pPr/w:pStyle[@w:val=$StyleId]/following-sibling::w:numPr/w:numId/@w:val"/>  
        </xsl:otherwise>
      </xsl:choose>
</xsl:variable>    
    
 <xsl:variable name="AbstractNumId">
   <xsl:choose>
     <xsl:when test="document('word/numbering.xml')/w:numbering/w:num[@w:numId=$NumId]/w:abstractNumId/@w:val">
       <xsl:value-of select="document('word/numbering.xml')/w:numbering/w:num[@w:numId=$NumId]/w:abstractNumId/@w:val"/>
     </xsl:when>
     <xsl:otherwise>0</xsl:otherwise>
   </xsl:choose>
 </xsl:variable>
    
<xsl:variable name="LeftNumber">
  <xsl:choose>
    <xsl:when test="document('word/numbering.xml')/w:numbering/w:abstractNum[@w:abstractNumId = $AbstractNumId]/w:lvl[@w:ilvl=$Ivl]/w:pPr/w:ind/@w:left">
      <xsl:value-of select="document('word/numbering.xml')/w:numbering/w:abstractNum[@w:abstractNumId = $AbstractNumId]/w:lvl[@w:ilvl=$Ivl]/w:pPr/w:ind/@w:left"/>    
    </xsl:when>
    <xsl:otherwise>0</xsl:otherwise>
  </xsl:choose>
</xsl:variable>
    
 <xsl:variable name="HangingNumber">
   <xsl:choose>
     <xsl:when test="document('word/numbering.xml')/w:numbering/w:abstractNum[@w:abstractNumId = $AbstractNumId]/w:lvl[@w:ilvl=$Ivl]/w:pPr/w:ind/@w:hanging">
       <xsl:value-of select="document('word/numbering.xml')/w:numbering/w:abstractNum[@w:abstractNumId = $AbstractNumId]/w:lvl[@w:ilvl=$Ivl]/w:pPr/w:ind/@w:hanging"/>    
     </xsl:when>
     <xsl:otherwise>0</xsl:otherwise>
   </xsl:choose>
 </xsl:variable>
    
    <xsl:variable name="paragraphIndent">
      <xsl:choose>
        <xsl:when test="document('word/styles.xml')/w:styles/w:style[@w:styleId=$StyleId]/w:pPr/w:ind/@w:left">
          <xsl:value-of select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$StyleId]/w:pPr/w:ind/@w:left"/>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
      
    </xsl:variable>   
    <xsl:if test="w:ind">
      <xsl:if test="w:ind/@w:left">
        <xsl:attribute name="fo:margin-left">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:choose>
                <xsl:when test="w:ind/@w:left and w:numPr/w:numId/@w:val">
                  <xsl:value-of select="w:ind/@w:left  - $LeftNumber + $HangingNumber + $paragraphIndent"/>                  
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="w:ind/@w:left"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="w:ind/@w:right">
        <xsl:attribute name="fo:margin-right">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:value-of select="w:ind/@w:right"/>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="w:ind/@w:firstLine">
        <xsl:attribute name="fo:text-indent">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:value-of select="w:ind/@w:firstLine"/>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:choose>
        <xsl:when test="w:ind/@w:hanging">
          <xsl:attribute name="fo:text-indent">
            <xsl:call-template name="ConvertTwips">
              <xsl:with-param name="length">
                <xsl:value-of select="-w:ind/@w:hanging"/>
              </xsl:with-param>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="not(w:ind/@w:hanging) and w:numPr/w:numId/@w:val">
          <xsl:attribute name="fo:text-indent">
            <xsl:call-template name="ConvertTwips">
              <xsl:with-param name="length">
                <xsl:value-of select="-$HangingNumber"/>
              </xsl:with-param>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise/>        
      </xsl:choose>
      
      <xsl:if test="w:ind/@w:hanging">
        <xsl:attribute name="fo:text-indent">
          <xsl:call-template name="ConvertTwips">
            <xsl:with-param name="length">
              <xsl:value-of select="-w:ind/@w:hanging"/>
            </xsl:with-param>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
    </xsl:if>


    <!-- w:afterAutospacing and w:beforeAutospacing attributes are lost -->
    <!-- w:afterLines and w:beforeLines attributes are lost -->

    <xsl:if test="w:spacing/@w:before">
      <xsl:attribute name="fo:margin-top">
        <xsl:call-template name="ConvertTwips">
          <xsl:with-param name="length">
            <xsl:value-of select="w:spacing/@w:before"/>
          </xsl:with-param>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="w:spacing/@w:after">
      <xsl:attribute name="fo:margin-bottom">
        <xsl:call-template name="ConvertTwips">
          <xsl:with-param name="length">
            <xsl:value-of select="w:spacing/@w:after"/>
          </xsl:with-param>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>

    
    <!-- line spacing -->
    <xsl:if test="w:spacing/@w:line">
      <xsl:choose>
        <xsl:when test="w:spacing/@w:lineRule='atLeast'">
          <xsl:attribute name="style:line-height-at-least">
            <xsl:call-template name="ConvertTwips">
              <xsl:with-param name="length">
                <xsl:value-of select="w:spacing/@w:line"/>
              </xsl:with-param>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="w:spacing/@w:lineRule='exact'">
          <xsl:attribute name="fo:line-height">
            <xsl:call-template name="ConvertTwips">
              <xsl:with-param name="length">
                <xsl:value-of select="w:spacing/@w:line"/>
              </xsl:with-param>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <!-- value of lineRule is 'auto' -->
          <xsl:attribute name="fo:line-height">
            <!-- convert 240th of line to percent -->
            <xsl:value-of select="concat(w:spacing/@w:line div 240 * 100,'%')"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <!-- text:align -->
    <xsl:if test="w:jc">
      <xsl:attribute name="fo:text-align">
        <xsl:choose>
          <xsl:when test="w:jc/@w:val='center'">
            <xsl:value-of select="'center'"/>
          </xsl:when>
          <xsl:when test="w:jc/@w:val='left'">
            <xsl:value-of select="'start'"/>
          </xsl:when>
          <xsl:when test="w:jc/@w:val='right'">
            <xsl:value-of select="'end'"/>
          </xsl:when>
          <xsl:when test="w:jc/@w:val='both' or w:jc/@w:val='distribute'">
            <xsl:value-of select="'justify'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'start'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <!-- vertical alignment of text -->
    <xsl:if test="w:textAlignment">
      <xsl:attribute name="style:vertical-align">
        <xsl:choose>
          <xsl:when test="w:textAlignment/@w:val='bottom'">
            <xsl:value-of select="'bottom'"/>
          </xsl:when>
          <xsl:when test="w:textAlignment/@w:val='top'">
            <xsl:value-of select="'top'"/>
          </xsl:when>
          <xsl:when test="w:textAlignment/@w:val='center'">
            <xsl:value-of select="'middle'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'auto'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:snapToGrid">
      <xsl:attribute name="style:snap-to-layout-grid">
        <xsl:choose>
          <xsl:when
            test="w:snapToGrid/@w:val='off' or w:snapToGrid/@w:val='false' or w:snapToGrid/@w:val=0">
            <xsl:value-of select="'false'"/>
          </xsl:when>
          <xsl:when
            test="w:snapToGrid/@w:val='on' or w:snapToGrid/@w:val='true' or w:snapToGrid/@w:val=1">
            <xsl:value-of select="'true'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'true'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:suppressLineNumbers">
      <xsl:attribute name="text:number-lines">
        <xsl:choose>
          <xsl:when
            test="w:suppressLineNumbers/@w:val='off' or w:suppressLineNumbers/@w:val='false' or w:suppressLineNumbers/@w:val=0">
            <xsl:value-of select="'true'"/>
          </xsl:when>
          <xsl:when
            test="w:suppressLineNumbers/@w:val='on' or w:suppressLineNumbers/@w:val='true' or w:suppressLineNumbers/@w:val=1">
            <xsl:value-of select="'false'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'false'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:overflowPunct">
      <xsl:attribute name="style:punctuation-wrap">
        <xsl:choose>
          <xsl:when
            test="w:overflowPunct/@w:val='off' or w:overflowPunct/@w:val='false' or w:overflowPunct/@w:val=0">
            <xsl:value-of select="'hanging'"/>
          </xsl:when>
          <xsl:when
            test="w:overflowPunct/@w:val='on' or w:overflowPunct/@w:val='true' or w:overflowPunct/@w:val=1">
            <xsl:value-of select="'simple'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'simple'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <!-- text direction. btLr and tbLrV directions are lost. -->
    <xsl:if test="w:textDirection">
      <xsl:attribute name="style:writing-mode">
        <xsl:choose>
          <xsl:when test="w:textDirection/@w:val='lrTb'">lr-tb</xsl:when>
          <xsl:when test="w:textDirection/@w:val='lrTbV'">lr-tb</xsl:when>
          <xsl:when test="w:textDirection/@w:val='tbRl'">rl-tb</xsl:when>
          <xsl:when test="w:textDirection/@w:val='tbRlV'">tb-rl</xsl:when>
          <xsl:otherwise/>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

     <!-- widow and orphan-->
    <xsl:choose>
      <xsl:when test="w:widowControl/@w:val='0'">
        <xsl:attribute name="fo:widows">
          <xsl:value-of select="1"/>
        </xsl:attribute>
        <xsl:attribute name="fo:orphans">
          <xsl:value-of select="1"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="fo:widows">
          <xsl:value-of select="2"/>
        </xsl:attribute>
        <xsl:attribute name="fo:orphans">
          <xsl:value-of select="2"/>
        </xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>

    <!-- tab stops -->
    <xsl:if test="w:tabs">
      <style:tab-stops>
        <xsl:for-each select="w:tabs/w:tab">
          <xsl:call-template name="InsertTabs"/>
        </xsl:for-each>
      </style:tab-stops>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertParagraphBorder">
    <xsl:if test="w:pBdr/w:between">
      <xsl:attribute name="style:join-border">false</xsl:attribute>
    </xsl:if>

    <xsl:for-each select="w:pBdr">
      <xsl:choose>
        <!-- if the four borders are equal, then create adequate attributes. -->
        <xsl:when
          test="w:top/@color=w:bottom/@color and w:top/@space=w:bottom/@space and w:top/@sz=w:bottom/@sz and w:top/@val=w:bottom/@val
          and w:top/@color=w:left/@color and w:top/@space=w:left/@space and w:top/@sz=w:left/@sz and w:top/@val=w:left/@val
          and w:top/@color=w:right/@color and w:top/@space=w:right/@space and w:top/@sz=w:right/@sz and w:top/@val=w:right/@val">
          <xsl:call-template name="InsertBorderAttributes"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each
            select="child::node()[name() = 'w:top' or name() = 'w:left' or name() = 'w:bottom' or name() = 'w:right']">
            <xsl:if test="./@w:val != 'none'">
            <xsl:call-template name="InsertBorderAttributes">
              <xsl:with-param name="side" select="substring-after(name(),'w:')"/>
            </xsl:call-template>
            </xsl:if>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <!-- compute attributes defining borders : color, style, width, padding... -->
  <xsl:template name="InsertBorderAttributes">
    <xsl:param name="side"/>

    <xsl:choose>
      <xsl:when test="$side">
        <!-- add padding -->
        <xsl:if test="@w:space">
          <xsl:attribute name="{concat('fo:padding-',$side)}">
            <xsl:call-template name="ConvertPoints">
              <xsl:with-param name="length">
                <xsl:value-of select="@w:space"/>
              </xsl:with-param>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
        <!-- add border with style, width and color -->
        <xsl:variable name="style">
          <xsl:call-template name="GetParagraphBorderStyle">
            <xsl:with-param name="style" select="@w:val"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:attribute name="{concat('fo:border-',$side)}">
          <xsl:variable name="size">
            <xsl:call-template name="ConvertEighthsPoints">
              <xsl:with-param name="length">
                <xsl:value-of select="@w:sz"/>
              </xsl:with-param>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test="@w:color='auto'">
              <xsl:value-of select="concat($size,' ',$style,' #000000')"/>
            </xsl:when>
            <xsl:when test="@w:color=''">
              <xsl:value-of select="concat($size,' ',$style,' #000000')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:variable name="color">
                <xsl:value-of select="concat('#',@w:color)"/>
              </xsl:variable>
              <xsl:value-of select="concat($size,' ',$style,' ',$color)"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:if test="$style='double' ">
          <xsl:attribute name="{concat('fo:border-line-width-',$side)}">
            <xsl:call-template name="ComputeDoubleBorderWidth">
              <xsl:with-param name="style" select="@w:val"/>
              <xsl:with-param name="borderWidth" select="@w:sz"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <xsl:otherwise>
        <!-- add padding -->
        <xsl:if test="w:top/@w:space">
          <xsl:attribute name="fo:padding">
            <xsl:call-template name="ConvertPoints">
              <xsl:with-param name="length">
                <xsl:value-of select="w:top/@w:space"/>
              </xsl:with-param>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
        <!-- add border with style, width and color -->
        <xsl:variable name="style">
          <xsl:call-template name="GetParagraphBorderStyle">
            <xsl:with-param name="style" select="w:top/@w:val"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:attribute name="fo:border">
          <xsl:variable name="size">
            <xsl:call-template name="ConvertEighthsPoints">
              <xsl:with-param name="length">
                <xsl:value-of select="w:top/@w:sz"/>
              </xsl:with-param>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test="@w:color='auto'">
              <xsl:value-of select="concat($size,' ',$style)"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:variable name="color">
                <xsl:value-of select="concat('#',w:top/@w:color)"/>
              </xsl:variable>
              <xsl:value-of select="concat($size,' ',$style,' ',$color)"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:if test="$style='double' ">
          <xsl:attribute name="fo:border-line-width">
            <xsl:call-template name="ComputeDoubleBorderWidth">
              <xsl:with-param name="style" select="w:top/@w:val"/>
              <xsl:with-param name="borderWidth" select="w:top/@w:sz"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <xsl:template name="InsertParagraphShadow">

    <xsl:variable name="firstVal">
      <xsl:choose>
        <xsl:when
          test="w:pBdr/w:right/@w:shadow='true' or w:pBdr/w:right/@w:shadow=1 or w:pBdr/w:right/@w:shadow='on'">0.0701in</xsl:when>
        <xsl:otherwise>
          <xsl:choose>
            <xsl:when
              test="w:pBdr/w:left/@w:shadow='true' or w:pBdr/w:left/@w:shadow=1 or w:pBdr/w:left/@w:shadow='on'">0.0701in</xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:variable name="secondVal">
      <xsl:choose>
        <xsl:when
          test="w:pBdr/w:bottom/@w:shadow='true' or w:pBdr/w:bottom/@w:shadow=1 or w:pBdr/w:bottom/@w:shadow='on'">0.0701in</xsl:when>
        <xsl:otherwise>
          <xsl:choose>
            <xsl:when
              test="w:pBdr/w:top/@w:shadow='true' or w:pBdr/w:top/@w:shadow=1 or w:pBdr/w:top/@w:shadow='on'">0.0701in</xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:if test="$firstVal !=0 and $secondVal != 0">
      <xsl:attribute name="style:shadow">
        <xsl:value-of select="concat('#000000 ',$firstVal,' ',$secondVal)"/>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

  <!-- convert the style of a border -->
  <xsl:template name="GetParagraphBorderStyle">
    <xsl:param name="style"/>
    <xsl:choose>
      <xsl:when test="$style='basicBlackDashes'">dashed</xsl:when>
      <xsl:when test="$style='basicBlackDots'">dotted</xsl:when>
      <xsl:when test="$style='basicThinLines'">double</xsl:when>
      <xsl:when test="$style='basicWideInLine'">double</xsl:when>
      <xsl:when test="$style='basicWideOutLine'">double</xsl:when>
      <xsl:when test="$style='dashed'">dashed</xsl:when>
      <xsl:when test="$style='dashSmallGap'">dashed</xsl:when>
      <xsl:when test="$style='dotted'">dotted</xsl:when>
      <xsl:when test="$style='double'">double</xsl:when>
      <xsl:when test="$style='inset'">inset</xsl:when>
      <xsl:when test="$style='nil'">hidden</xsl:when>
      <xsl:when test="$style='outset'">outset</xsl:when>
      <xsl:when test="$style='single'">solid</xsl:when>
      <xsl:when test="$style='thick'">solid</xsl:when>
      <xsl:when test="$style='thickThinSmallGap'">double</xsl:when>
      <xsl:when test="$style='thickThinMediumGap'">double</xsl:when>
      <xsl:when test="$style='thickThinLargeGap'">double</xsl:when>
      <xsl:when test="$style='thinThickSmallGap'">double</xsl:when>
      <xsl:when test="$style='thinThickMediumGap'">double</xsl:when>
      <xsl:when test="$style='thinThickLargeGap'">double</xsl:when>
      <xsl:otherwise>
        <!-- empty string -->
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- return a three arguments string defining the widths of a 'double' border -->
  <xsl:template name="ComputeDoubleBorderWidth">
    <xsl:param name="style"/>
    <xsl:param name="borderWidth"/>
    <!-- Approximate the width of each line : inner, middle (blank space), outer -->
    <xsl:variable name="inner">
      <xsl:call-template name="ConvertEighthsPoints">
        <xsl:with-param name="length">
          <xsl:choose>
            <xsl:when test="$style='basicWideInLine'">
              <xsl:value-of select="$borderWidth * 50 div 100"/>
            </xsl:when>
            <xsl:when test="$style='basicWideOutLine'">
              <xsl:value-of select="$borderWidth div 100"/>
            </xsl:when>
            <xsl:when test="$style='double' or $style='basicThinLines'">
              <xsl:value-of select="$borderWidth div 100"/>
            </xsl:when>
            <xsl:when test="$style='thickThinSmallGap'">
              <xsl:value-of select="$borderWidth * 67 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thickThinMediumGap'">
              <xsl:value-of select="$borderWidth * 47 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thickThinLargeGap'">
              <xsl:value-of select="$borderWidth * 27 div 100"/>
            </xsl:when>
            <xsl:when test="contains($style,'thinThick')">
              <xsl:value-of select="$borderWidth * 3 div 100"/>
            </xsl:when>
            <xsl:otherwise/>
          </xsl:choose>
        </xsl:with-param>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="middle">
      <xsl:call-template name="ConvertEighthsPoints">
        <xsl:with-param name="length">
          <xsl:choose>
            <xsl:when test="$style='basicWideInLine' or $style='basicWideOutLine'">
              <xsl:value-of select="$borderWidth * 49 div 100"/>
            </xsl:when>
            <xsl:when test="$style='double' or $style='basicThinLines'">
              <xsl:value-of select="$borderWidth * 98 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thickThinSmallGap'">
              <xsl:value-of select="$borderWidth * 30 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thickThinMediumGap'">
              <xsl:value-of select="$borderWidth * 50 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thickThinLargeGap'">
              <xsl:value-of select="$borderWidth * 70 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thinThickSmallGap'">
              <xsl:value-of select="$borderWidth * 30 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thinThickMediumGap'">
              <xsl:value-of select="$borderWidth * 50 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thinThickLargeGap'">
              <xsl:value-of select="$borderWidth * 70 div 100"/>
            </xsl:when>
            <xsl:otherwise/>
          </xsl:choose>
        </xsl:with-param>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="outer">
      <xsl:call-template name="ConvertEighthsPoints">
        <xsl:with-param name="length">
          <xsl:choose>
            <xsl:when test="$style='basicWideInLine'">
              <xsl:value-of select="$borderWidth div 100"/>
            </xsl:when>
            <xsl:when test="$style='basicWideOutLine'">
              <xsl:value-of select="$borderWidth * 50 div 100"/>
            </xsl:when>
            <xsl:when test="$style='double' or $style='basicThinLines'">
              <xsl:value-of select="$borderWidth div 100"/>
            </xsl:when>
            <xsl:when test="contains($style,'thickThin')">
              <xsl:value-of select="$borderWidth * 3 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thinThickSmallGap'">
              <xsl:value-of select="$borderWidth * 67 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thinThickMediumGap'">
              <xsl:value-of select="$borderWidth * 47 div 100"/>
            </xsl:when>
            <xsl:when test="$style='thinThickLargeGap'">
              <xsl:value-of select="$borderWidth * 27 div 100"/>
            </xsl:when>
            <xsl:otherwise/>
          </xsl:choose>
        </xsl:with-param>
        <xsl:with-param name="unit">cm</xsl:with-param>
      </xsl:call-template>
    </xsl:variable>

    <xsl:value-of select="concat($inner,' ',$middle,' ',$outer)"/>
  </xsl:template>

  <!-- Handle tab stops -->
  <!-- TODO : check how to deal with tab stops inside a list -->
  <xsl:template name="InsertTabs">
    <style:tab-stop>
      <xsl:if test="./@w:val != 'num' and ./@w:val != 'clear'">
        <!--   type -->
        <xsl:attribute name="style:type">
          <xsl:choose>
            <xsl:when test="./@w:val='center'">center</xsl:when>
            <xsl:when test="./@w:val='right'">right</xsl:when>
            <xsl:when test="./@w:val='left'">left</xsl:when>
            <xsl:when test="./@w:val='decimal'">char</xsl:when>
            <xsl:otherwise>left</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:if test="./@w:val='decimal'">
          <xsl:attribute name="style:char">
            <xsl:value-of select="','"/>
          </xsl:attribute>
        </xsl:if>
        <!-- position 
         TODO : what if @w:pos < 0 ? -->
        <xsl:if test="./@w:pos >= 0">
          <xsl:attribute name="style:position">
            <xsl:call-template name="ConvertTwips">
              <xsl:with-param name="length" select="./@w:pos"/>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
        <!-- leader char -->
        <xsl:if test="./@w:leader">
          <xsl:attribute name="style:leader-style">
            <xsl:choose>
              <xsl:when test="./@w:leader='dot'">dotted</xsl:when>
              <xsl:when test="./@w:leader='heavy'">solid</xsl:when>
              <xsl:when test="./@w:leader='hyphen'">solid</xsl:when>
              <xsl:when test="./@w:leader='middleDot'">dotted</xsl:when>
              <xsl:when test="./@w:leader='none'">none</xsl:when>
              <xsl:when test="./@w:leader='underscore'">solid</xsl:when>
              <xsl:otherwise>none</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>
        <!--        leader text  -->
        <xsl:if test="./@w:leader">
          <xsl:attribute name="style:leader-text">
            <xsl:choose>
              <xsl:when test="./@w:leader='dot'">.</xsl:when>
              <xsl:when test="./@w:leader='heavy'"/>
              <xsl:when test="./@w:leader='hyphen'">-</xsl:when>
              <xsl:when test="./@w:leader='middleDot'"/>
              <xsl:when test="./@w:leader='none'"/>
              <xsl:when test="./@w:leader='underscore'">_</xsl:when>
              <xsl:otherwise>none</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>
      </xsl:if>
    </style:tab-stop>
  </xsl:template>

  <!-- ODF Text properties contained in OOX pPr element -->
  <xsl:template name="InsertpPrTextProperties">
    <!-- hyphenation -->
    <xsl:if test="w:suppressAutoHyphens">
      <xsl:choose>
        <xsl:when
          test="w:suppressAutoHyphens/@w:val='off' or w:suppressAutoHyphens/@w:val='false' or w:suppressAutoHyphens/@w:val=0">
          <xsl:attribute name="fo:hyphenate">true</xsl:attribute>
          <xsl:attribute name="fo:hyphenation-remain-char-count">2</xsl:attribute>
          <xsl:attribute name="fo:hyphenation-push-char-count">2</xsl:attribute>
        </xsl:when>
        <xsl:when
          test="w:suppressAutoHyphens/@w:val='on' or w:suppressAutoHyphens/@w:val='true' or w:suppressAutoHyphens/@w:val=1">
          <xsl:attribute name="fo:hyphenate">false</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="fo:hyphenate">false</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
  </xsl:template>

  <!-- conversion of text properties -->
  <xsl:template name="InsertTextProperties">

    <xsl:if test="w:b">
      <xsl:attribute name="fo:font-weight">
        <xsl:choose>
          <xsl:when test="w:b/@w:val='off' or w:b/@w:val='false' or w:b/@w:val=0">
            <xsl:value-of select="'normal'"/>
          </xsl:when>
          <xsl:when test="w:b/@w:val='on' or w:b/@w:val='true' or w:b/@w:val=1">
            <xsl:value-of select="'bold'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'bold'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:bCs">
      <xsl:attribute name="style:font-weight-complex">
        <xsl:choose>
          <xsl:when test="w:bCs/@w:val='off' or w:bCs/@w:val='false' or w:bCs/@w:val=0">
            <xsl:value-of select="'normal'"/>
          </xsl:when>
          <xsl:when test="w:bCs/@w:val='on' or w:bCs/@w:val='true' or w:bCs/@w:val=1">
            <xsl:value-of select="'bold'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'bold'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:i">
      <xsl:attribute name="fo:font-style">
        <xsl:choose>
          <xsl:when test="w:i/@w:val='off' or w:i/@w:val='false' or w:i/@w:val=0">
            <xsl:value-of select="'normal'"/>
          </xsl:when>
          <xsl:when test="w:i/@w:val='on' or w:i/@w:val='true' or w:i/@w:val=1">
            <xsl:value-of select="'italic'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'italic'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:iCs">
      <xsl:attribute name="style:font-style-complex">
        <xsl:choose>
          <xsl:when test="w:iCs/@w:val='off' or w:iCs/@w:val='false' or w:iCs/@w:val=0">
            <xsl:value-of select="'normal'"/>
          </xsl:when>
          <xsl:when test="w:iCs/@w:val='on' or w:iCs/@w:val='true' or w:iCs/@w:val=1">
            <xsl:value-of select="'italic'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'italic'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:caps">
      <xsl:attribute name="fo:text-transform">
        <xsl:choose>
          <xsl:when test="w:caps/@w:val='off' or w:caps/@w:val='false' or w:caps/@w:val=0">
            <xsl:value-of select="'none'"/>
          </xsl:when>
          <xsl:when test="w:caps/@w:val='on' or w:caps/@w:val='true' or w:caps/@w:val=1">
            <xsl:value-of select="'uppercase'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'uppercase'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:smallCaps">
      <xsl:attribute name="fo:font-variant">
        <xsl:choose>
          <xsl:when
            test="w:smallCaps/@w:val='off' or w:smallCaps/@w:val='false' or w:smallCaps/@w:val=0">
            <xsl:value-of select="'normal'"/>
          </xsl:when>
          <xsl:when
            test="w:smallCaps/@w:val='on' or w:smallCaps/@w:val='true' or w:smallCaps/@w:val=1">
            <xsl:value-of select="'small-caps'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'small-caps'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <!-- line through text -->
    <xsl:if test="w:dstrike or w:strike">
      <xsl:choose>
        <xsl:when
          test="w:strike/@w:val='on' or w:strike/@w:val='true' or w:strike/@w:val=1 or w:strike/@w:val=1 or (w:strike and (not(w:strike/@w:val) or w:strike/@w:val = ''))">
          <xsl:attribute name="style:text-line-through-type">single</xsl:attribute>
          <xsl:attribute name="style:text-line-through-style">solid</xsl:attribute>
        </xsl:when>
        <xsl:when
          test="w:dstrike/@w:val='on' or w:dstrike/@w:val='true' or w:dstrike/@w:val=1 or w:dstrike/@w:val=1 or (w:dstrike and (not(w:dstrike/@w:val) or w:dstrike/@w:val = ''))">
          <xsl:attribute name="style:text-line-through-type">double</xsl:attribute>
          <xsl:attribute name="style:text-line-through-style">solid</xsl:attribute>
        </xsl:when>
        <xsl:otherwise/>
      </xsl:choose>
    </xsl:if>

    <xsl:if test="w:outline">
      <xsl:attribute name="style:text-outline">
        <xsl:choose>
          <xsl:when test="w:outline/@w:val='off' or w:outline/@w:val='false' or w:outline/@w:val=0">
            <xsl:value-of select="'false'"/>
          </xsl:when>
          <xsl:when test="w:outline/@w:val='on' or w:outline/@w:val='true' or w:outline/@w:val=1">
            <xsl:value-of select="'true'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'true'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:shadow">
      <xsl:attribute name="fo:text-shadow">
        <xsl:choose>
          <xsl:when test="w:shadow/@w:val='off' or w:shadow/@w:val='false' or w:shadow/@w:val=0">
            <xsl:value-of select="'none'"/>
          </xsl:when>
          <xsl:when test="w:shadow/@w:val='on' or w:shadow/@w:val='true' or w:shadow/@w:val=1">
            <xsl:value-of select="'#000000 0.2em 0.2em'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'#000000 0.2em 0.2em'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:imprint">
      <xsl:attribute name="style:font-relief">
        <xsl:value-of select="'engraved'"/>
      </xsl:attribute>
    </xsl:if>
    
    <xsl:if test="w:emboss">
      <xsl:attribute name="style:font-relief">
        <xsl:value-of select="'embossed'"/>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:vanish">
      <xsl:attribute name="text:display">
         <xsl:value-of select="'true'"/>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:em">
      <xsl:attribute name="style:text-emphasize">
        <xsl:choose>
          <xsl:when test="w:em/@w:val = 'circle'">circle above</xsl:when>
          <xsl:when test="w:em/@w:val = 'comma'">accent above</xsl:when>
          <xsl:when test="w:em/@w:val = 'dot'">dot above</xsl:when>
          <xsl:when test="w:em/@w:val = 'underDot'">dot below</xsl:when>
          <xsl:when test="w:em/@w:val = 'none'">none</xsl:when>
          <xsl:otherwise>none</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <!-- text bg color -->
    <xsl:if test="w:highlight">
      <xsl:attribute name="fo:background-color">
        <xsl:choose>
          <xsl:when test="w:highlight/@w:val = 'black'">#000000</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'blue'">#0000FF</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'cyan'">#00FFFF</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'darkBlue'">#000080</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'darkCyan'">#008080</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'darkGray'">#808080</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'darkGreen'">#008000</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'darkMagenta'">#800080</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'darkRed'">#800000</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'darkYellow'">#808000</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'green'">#00FF00</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'lightGray'">#C0C0C0</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'magenta'">#FF00FF</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'red'">#FF0000</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'white'">#FFFFFF</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'yellow'">#FFFF00</xsl:when>
          <xsl:when test="w:highlight/@w:val = 'none'">transparent</xsl:when>
          <xsl:otherwise>transparent</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <!-- underline -->
    <xsl:if test="w:u">
      <xsl:call-template name="InsertUnderline"/>
    </xsl:if>

    <!-- fonts -->
    <xsl:choose>
      <xsl:when test="ancestor::node()/w:style[@w:type='paragraph' and @w:default='1']">
        <xsl:attribute name="style:font-name">
           <xsl:value-of select="ancestor::node()/w:style[@w:type='paragraph' and @w:default='1']/w:rPr/w:rFonts/@w:ascii"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
    
    <xsl:if test="w:rFonts/@w:asciiTheme">
      <xsl:attribute name="style:font-name">
        <xsl:call-template name="ComputeThemeFontName">
          <xsl:with-param name="fontTheme" select="w:rFonts/@w:asciiTheme"/>
          <xsl:with-param name="fontType">a:latin</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:if test="w:rFonts/@w:ascii">
      <xsl:attribute name="style:font-name">
        <xsl:value-of select="w:rFonts/@w:ascii"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="w:rFonts/@w:cstheme">
      <xsl:attribute name="style:font-name-complex">
        <xsl:call-template name="ComputeThemeFontName">
          <xsl:with-param name="fontTheme" select="w:rFonts/@w:cstheme"/>
          <xsl:with-param name="fontType">a:cs</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="w:rFonts/@w:cs">
      <xsl:attribute name="style:font-name-complex">
        <xsl:value-of select="w:rFonts/@w:cs"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="w:rFonts/@w:eastAsiaTheme">
      <xsl:attribute name="style:font-name-asian">
        <xsl:call-template name="ComputeThemeFontName">
          <xsl:with-param name="fontTheme" select="w:rFonts/@w:eastAsiaTheme"/>
          <xsl:with-param name="fontType">a:ea</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="w:rFonts/@w:eastAsia">
      <xsl:attribute name="style:font-name-asian">
        <xsl:value-of select="w:rFonts/@w:eastAsia"/>
      </xsl:attribute>
    </xsl:if>

    <!-- text color -->
    <xsl:if test="w:color/@w:val != 'auto'">
      <xsl:attribute name="fo:color">
        <xsl:value-of select="concat('#',w:color/@w:val)"/>
      </xsl:attribute>
    </xsl:if>

    <!-- letter spacing -->
    <xsl:if test="w:spacing">
      <xsl:attribute name="fo:letter-spacing">
        <xsl:choose>
          <xsl:when test="w:spacing/@w:val=0">normal</xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="ConvertTwips">
              <xsl:with-param name="length" select="w:spacing/@w:val"/>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:w/@w:val">
      <xsl:attribute name="style:text-scale">
        <xsl:value-of select="w:w/@w:val"/>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:kern">
      <xsl:attribute name="style:letter-kerning">true</xsl:attribute>
    </xsl:if>

    <xsl:if test="w:sz">
      <xsl:attribute name="fo:font-size">
        <xsl:call-template name="ConvertHalfPoints">
          <xsl:with-param name="length" select="w:sz/@w:val"/>
          <xsl:with-param name="unit">pt</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="w:szCs">
      <xsl:attribute name="style:font-size-complex">
        <xsl:call-template name="ConvertHalfPoints">
          <xsl:with-param name="length" select="w:szCs/@w:val"/>
          <xsl:with-param name="unit">pt</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>

    <!-- text vertical positionning -->
    <xsl:if test="w:vertAlign or w:position">
      <xsl:call-template name="InsertTextPosition"/>
    </xsl:if>

    <!-- languages and countries -->
    <xsl:if test="w:lang">
      <xsl:if test="w:lang/@w:val">
        <xsl:attribute name="fo:language">
          <xsl:value-of select="substring-before(w:lang/@w:val,'-')"/>
        </xsl:attribute>
        <xsl:attribute name="fo:country">
          <xsl:value-of select="substring-after(w:lang/@w:val,'-')"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="w:lang/@w:bidi">
        <xsl:attribute name="style:language-complex">
          <xsl:value-of select="substring-before(w:lang/@w:bidi,'-')"/>
        </xsl:attribute>
        <xsl:attribute name="style:country-complex">
          <xsl:value-of select="substring-after(w:lang/@w:bidi,'-')"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="w:lang/@w:eastAsia">
        <xsl:attribute name="style:language-asian">
          <xsl:value-of select="substring-before(w:lang/@w:eastAsia,'-')"/>
        </xsl:attribute>
        <xsl:attribute name="style:country-asian">
          <xsl:value-of select="substring-after(w:lang/@w:eastAsia,'-')"/>
        </xsl:attribute>
      </xsl:if>
    </xsl:if>

    <xsl:if test="w:eastAsianLayout">
      <xsl:if test="w:eastAsianLayout/@w:combine = 'lines'">
        <xsl:attribute name="style:text-combine">lines</xsl:attribute>
        <xsl:choose>
          <xsl:when test="w:combineBrackets = 'angle'">
            <xsl:attribute name="style:text-combine-start-char">
              <xsl:value-of select="'&lt;'"/>
            </xsl:attribute>
            <xsl:attribute name="style:text-combine-end-char">
              <xsl:value-of select="'&gt;'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="w:combineBrackets = 'curly'">
            <xsl:attribute name="style:text-combine-start-char">{</xsl:attribute>
            <xsl:attribute name="style:text-combine-end-char">}</xsl:attribute>
          </xsl:when>
          <xsl:when test="w:combineBrackets = 'round'">
            <xsl:attribute name="style:text-combine-start-char">(</xsl:attribute>
            <xsl:attribute name="style:text-combine-end-char">)</xsl:attribute>
          </xsl:when>
          <xsl:when test="w:combineBrackets = 'square'">
            <xsl:attribute name="style:text-combine-start-char">[</xsl:attribute>
            <xsl:attribute name="style:text-combine-end-char">]</xsl:attribute>
          </xsl:when>
          <xsl:otherwise/>
        </xsl:choose>
      </xsl:if>
    </xsl:if>

    <!-- script type -->
    <xsl:if
      test="w:cs/@w:val='on' or w:cs/@w:val='true' or w:cs/@w:val=1 or (w:cs and (not(w:cs/@w:val) or w:cs/@w:val = ''))">
      <xsl:attribute name="style:script-type">complex</xsl:attribute>
    </xsl:if>

    <!-- text effect. Mostly lost. -->
    <xsl:if test="w:effect/@w:val='blinkBackground'">
      <xsl:attribute name="style:text-blinking">true</xsl:attribute>
    </xsl:if>
  </xsl:template>

  <!-- insert underline attributes -->
  <xsl:template name="InsertUnderline">
    <xsl:if test="w:u/@w:val">
      <xsl:choose>
        <xsl:when test="w:u/@w:val = 'dash'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">dash</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'dashDotDotHeavy'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">dot-dot-dash</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'dashDotHeavy'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">dot-dash</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'dashedHeavy'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">dash</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'dashLong'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">long-dash</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'dashLongHeavy'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">long-dash</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'dotDash'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">dot-dash</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'dotDotDash'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">dot-dot-dash</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'dotted'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">dotted</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'dottedHeavy'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">dotted</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'double'">
          <xsl:attribute name="style:text-underline-type">double</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">solid</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'single'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">solid</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'thick'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">solid</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'wave'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">wave</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'wavyDouble'">
          <xsl:attribute name="style:text-underline-type">double</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">wave</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'wavyHeavy'">
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">wave</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">thick</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'words'">
          <xsl:attribute name="style:text-underline-mode">skip-white-space</xsl:attribute>
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">solid</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:when>
        <xsl:when test="w:u/@w:val = 'none'">
          <xsl:attribute name="style:text-underline-type">none</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="style:text-underline-type">single</xsl:attribute>
          <xsl:attribute name="style:text-underline-style">solid</xsl:attribute>
          <xsl:attribute name="style:text-underline-width">normal</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <xsl:if test="w:u/@w:color">
      <xsl:attribute name="style:text-underline-color">
        <xsl:choose>
          <xsl:when test="w:u/@w:color = 'auto'">font-color</xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat('#',w:u/@w:color)"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

  <!-- compute positionning of text -->
  <xsl:template name="InsertTextPosition">
    <xsl:variable name="percentValue">
      <xsl:choose>
        <xsl:when test="w:position/@w:val">
          <xsl:choose>
            <xsl:when test="w:sz/@w:val != 0">
              <xsl:value-of select="round(w:position/@w:val * 100 div w:sz/@w:val)"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:variable name="defaultFontSize">
                <xsl:value-of
                  select="document('word/styles.xml')/w:styles/w:docDefaults/w:rPrDefault/w:rPr/w:sz/@w:val"
                />
              </xsl:variable>
              <xsl:value-of select="round(w:position/@w:val * 100 div number($defaultFontSize))"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:choose>

      <!-- positioning of superscript -->
      <xsl:when test="w:vertAlign/@w:val = 'superscript'">
        <xsl:attribute name="style:text-position">
          <xsl:value-of select="concat('super ',number(58 + $percentValue))"/>
        </xsl:attribute>
      </xsl:when>

      <!-- positioning of subscript -->
      <xsl:when test="w:vertAlign/@w:val = 'subscript'">
        <xsl:attribute name="style:text-position">
          <xsl:value-of select="concat('sub ',number(58 - $percentValue))"/>
        </xsl:attribute>
      </xsl:when>

      <!-- positioning of normal text -->
      <xsl:when test="w:vertAlign = 'baseline'">
        <xsl:attribute name="style:text-position">
          <xsl:value-of select="concat(number($percentValue),' 100')"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="style:text-position">
          <xsl:value-of select="concat(number($percentValue),' 100')"/>
        </xsl:attribute>
      </xsl:otherwise>

    </xsl:choose>
  </xsl:template>

  <!-- get font name from theme -->
  <xsl:template name="ComputeThemeFontName"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <xsl:param name="fontTheme"/>
    <xsl:param name="fontType"/>
    <xsl:variable name="fontScheme"
      select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme"/>

    <xsl:variable name="fontName">
      <xsl:choose>
        <xsl:when test="contains($fontTheme,'minor')">
          <xsl:value-of select="$fontScheme/a:minorFont/child::node()[name() = $fontType]/@typeface"
          />
        </xsl:when>
        <xsl:when test="contains($fontTheme,'major')">
          <xsl:value-of select="$fontScheme/a:majorFont/child::node()[name() = $fontType]/@typeface"
          />
        </xsl:when>
      </xsl:choose>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="$fontName = ''">
        <xsl:text>none</xsl:text>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$fontName"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
 </xsl:stylesheet>

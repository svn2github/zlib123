<?xml version="1.0" encoding="utf-8" ?>
<!--
Copyright (c) 2007, Sonata Software Limited
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
*     * Neither the name of Sonata Software Limited nor the names of its contributors
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
* SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  exclude-result-prefixes="a fo r">
  <xsl:template name ="presProps">
    <p:presentationPr>
      <p:showPr>
        <xsl:for-each select ="document('content.xml')//office:body/office:presentation">
          <xsl:variable name ="noOfSlides" select ="count(./draw:page)"/>
          <xsl:if test ="./presentation:settings">
            <xsl:for-each select ="./presentation:settings">
              <!-- Animations-->
              <xsl:if test ="./@presentation:animations='disabled'">
                <xsl:attribute name ="showAnimation">
                  <xsl:value-of select ="'0'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test ="not(./@presentation:animations) or ./@presentation:animations='enabled'">
                <xsl:attribute name ="showAnimation">
                  <xsl:value-of select ="'1'"/>
                </xsl:attribute>
              </xsl:if>
              <!-- Loop until ESC-->
              <xsl:if test ="./@presentation:endless and ./@presentation:pause">
                <xsl:attribute name ="loop">
                  <xsl:value-of select ="'1'"/>
                </xsl:attribute>
              </xsl:if>
              <!-- Change slides manually-->
              <xsl:if test ="./@presentation:force-manual='true' or ./@presentation:transition-on-click='disabled' ">
                <xsl:attribute name ="useTimings">
                  <xsl:value-of select ="'0'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test ="not(./@presentation:full-screen or ./@presentation:full-screen='true')">
                <p:present/>
              </xsl:if>
              <xsl:if test ="./@presentation:full-screen='false'">
                <p:browse/>
              </xsl:if>
              <xsl:if test ="./@presentation:start-page">
                <xsl:variable name ="pageName" >
                  <xsl:value-of select ="./@presentation:start-page"/>
                </xsl:variable>
                <xsl:variable name ="startPage">
                  <xsl:for-each select ="./parent::node()/draw:page">
                    <xsl:if test ="./@draw:name=$pageName">
                      <xsl:value-of select ="position()"/>
                    </xsl:if>
                  </xsl:for-each>
                </xsl:variable>
                <p:sldRg>
                  <xsl:attribute name ="st">
                    <xsl:value-of select ="$startPage"/>
                  </xsl:attribute>
                  <xsl:attribute name ="end">
                    <xsl:value-of select ="$noOfSlides"/>
                  </xsl:attribute>
                </p:sldRg>
              </xsl:if>
              <xsl:if test ="not(./@presentation:start-page or ./@presentation:show)">
                <p:sldAll />
              </xsl:if>
              <xsl:if test ="not(./@presentation:full-screen or ./@presentation:full-screen='true')">
                <p:penClr>
                  <a:srgbClr val="FF0000" />
                </p:penClr>
              </xsl:if>
              <xsl:if test ="./@presentation:show">
                <xsl:variable name ="custPresentationName">
                  <xsl:value-of select ="./@presentation:show"/>
                </xsl:variable>
                <xsl:variable name ="custPresentationId" >
                  <xsl:for-each select ="presentation:show">
                    <xsl:if test ="@presentation:name=$custPresentationName">
                      <xsl:value-of select ="position()"/>
                    </xsl:if>
                  </xsl:for-each>
                </xsl:variable>
                <p:custShow>
                  <xsl:attribute name ="id">
                    <xsl:value-of select ="$custPresentationId"/>
                  </xsl:attribute>
                </p:custShow>
              </xsl:if>
            </xsl:for-each>
          </xsl:if>
        </xsl:for-each>
      </p:showPr>
    </p:presentationPr>
  </xsl:template>
</xsl:stylesheet>

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
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  exclude-result-prefixes="office fo style">

  <xsl:template name="settings">
    <w:settings>
      <w:view w:val="print"/>
      <xsl:if
        test="document('styles.xml')//style:page-layout[@style:name=//office:master-styles/style:master-page/@style:page-layout-name]/style:page-layout-properties/@fo:background-color">
        <w:displayBackgroundShape/>
      </xsl:if>

      <!-- overwritten in each paragraph if necessary -->
      <w:autoHyphenation w:val="true"/>
      <w:consecutiveHyphenLimit w:val="0"/>
      <w:doNotHyphenateCaps w:val="false"/>

      <!-- Header and Footer even and odd propertie -->
      <xsl:call-template name="EvenAndOddConfiguration"/>

      <!-- Footnotes document wide properties -->
      <xsl:call-template name="footnotes-configuration">
        <xsl:with-param name="wide">yes</xsl:with-param>
      </xsl:call-template>

      <!-- Endnotes document wide properties -->
      <xsl:call-template name="endnotes-configuration">
        <xsl:with-param name="wide">yes</xsl:with-param>
      </xsl:call-template>

    </w:settings>
  </xsl:template>



</xsl:stylesheet>

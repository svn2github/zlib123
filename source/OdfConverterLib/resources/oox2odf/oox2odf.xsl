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
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:oox="urn:oox"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip" exclude-result-prefixes="oox">

  <xsl:import href="common.xsl"/>
  <xsl:import href="content.xsl"/>
  <xsl:import href="pictures.xsl"/>

  <xsl:param name="outputFile"/>
  <xsl:output method="xml" encoding="UTF-8"/>

  <!-- App version number -->
  <!-- WARNING: it has to be of type xx.yy -->
  <!-- (otherwise Word cannot open the doc) -->
  <xsl:variable name="app-version">0.3</xsl:variable>

  <xsl:template match="/oox:source">

    <pzip:archive pzip:target="{$outputFile}">

      <!-- Manifest -->
      <pzip:entry pzip:target="META-INF/manifest.xml">
        <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
          <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.text"
            manifest:full-path="/"/>
          <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
        </manifest:manifest>
      </pzip:entry>

      <!-- main content -->
      <pzip:entry pzip:target="content.xml">
        <xsl:call-template name="content"/>
      </pzip:entry>

      <!-- styles -->
      <pzip:entry pzip:target="styles.xml">

        <office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
          xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
          xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
          xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">
          <office:font-face-decls>
            <style:font-face style:name="Tahoma1" svg:font-family="Tahoma"/>
            <style:font-face style:name="Arial Unicode MS" svg:font-family="'Arial Unicode MS'"
              style:font-pitch="variable"/>
            <style:font-face style:name="MS Mincho" svg:font-family="'MS Mincho'"
              style:font-pitch="variable"/>
            <style:font-face style:name="Tahoma" svg:font-family="Tahoma"
              style:font-pitch="variable"/>
            <style:font-face style:name="Times New Roman" svg:font-family="'Times New Roman'"
              style:font-family-generic="roman" style:font-pitch="variable"/>
            <style:font-face style:name="Arial" svg:font-family="Arial"
              style:font-family-generic="swiss" style:font-pitch="variable"/>
          </office:font-face-decls>
          <office:styles>
            <style:default-style style:family="paragraph">
              <style:paragraph-properties fo:hyphenation-ladder-count="no-limit"
                style:text-autospace="ideograph-alpha" style:punctuation-wrap="hanging"
                style:line-break="strict" style:tab-stop-distance="1.251cm"
                style:writing-mode="page"/>
              <style:text-properties style:use-window-font-color="true"
                style:font-name="Times New Roman" fo:font-size="12pt" fo:language="fr"
                fo:country="FR" style:font-name-asian="Arial Unicode MS"
                style:font-size-asian="12pt" style:language-asian="none" style:country-asian="none"
                style:font-name-complex="Tahoma" style:font-size-complex="12pt"
                style:language-complex="none" style:country-complex="none" fo:hyphenate="false"
                fo:hyphenation-remain-char-count="2" fo:hyphenation-push-char-count="2"/>
            </style:default-style>
            <style:style style:name="Standard" style:family="paragraph" style:class="text"/>
          </office:styles>
        </office:document-styles>

      </pzip:entry>

    </pzip:archive>
  </xsl:template>


</xsl:stylesheet>

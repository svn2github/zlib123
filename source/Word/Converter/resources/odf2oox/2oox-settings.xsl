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
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"
  xmlns:ooo="http://openoffice.org/2004/office"
  exclude-result-prefixes="office fo style config ooo text">

  
  <xsl:variable name="configuration-settings"
    select="document('settings.xml')/office:document-settings/office:settings/config:config-item-set[@config:name='ooo:configuration-settings']"/>
  <!-- read only configuration setting -->
  <xsl:variable name="load-readonly"
    select="$configuration-settings/config:config-item[@config:name='LoadReadonly' and @config:type='boolean'] = 'true'"/>
  
  
  <xsl:template name="InsertSettings">
    <w:settings>

      <!--clam, dialogika: bugfix 1945545-->
      <w:stylePaneFormatFilter w:val="1021" />

      <!-- view layout -->
      <xsl:variable name="view-settings"
        select="document('settings.xml')/office:document-settings/office:settings/config:config-item-set[@config:name='ooo:view-settings']"/>
      <w:view>
        <xsl:attribute name="w:val">
          <xsl:choose>
            <xsl:when
              test="$view-settings/config:config-item[@config:name='InBrowseMode' and @config:type='boolean'] = 'true' "
              >web</xsl:when>
            <xsl:otherwise>print</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </w:view>

      <xsl:for-each select="document('styles.xml')">
        <xsl:if
          test="key('page-layouts', $default-master-style/@style:page-layout-name)/style:page-layout-properties/@fo:background-color">
          <w:displayBackgroundShape/>
        </xsl:if>
      </xsl:for-each>

      <!-- track changes -->
      <xsl:if
        test="document('content.xml')/office:document-content/office:body/office:text/text:tracked-changes/@text:track-changes">
        <w:trackRevisions
          w:val="{document('content.xml')/office:document-content/office:body/office:text/text:tracked-changes/@text:track-changes}"
        />
      </xsl:if>
      
      <!-- document protection -->
      <xsl:if test="$protected-sections[1] or boolean($load-readonly)">
        <w:documentProtection w:edit="readOnly" w:enforcement="1"/>
      </xsl:if>

      <xsl:if
        test="document('styles.xml')/office:document-styles/office:styles/style:default-style[@style:family='paragraph']/style:paragraph-properties/@style:tab-stop-distance">
        <w:defaultTabStop>
          <xsl:attribute name="w:val">
            <xsl:call-template name="twips-measure">
              <xsl:with-param name="length"
                select="document('styles.xml')/office:document-styles/office:styles/style:default-style[@style:family='paragraph']/style:paragraph-properties/@style:tab-stop-distance"
              />
            </xsl:call-template>
          </xsl:attribute>
        </w:defaultTabStop>
      </xsl:if>

      <!-- overwritten in each paragraph if necessary -->
      <w:autoHyphenation w:val="true"/>
      <w:consecutiveHyphenLimit w:val="0"/>
      <w:doNotHyphenateCaps w:val="false"/>

      <!-- Header and Footer settings -->
      <xsl:call-template name="InsertHeaderFooterSettings"/>

      <!-- Automatically update fields -->
      <!--xsl:if
        test="document('settings.xml')/office:document-settings/office:settings/config:config-item-set[@config:name='ooo:configuration-settings']/config:config-item[@config:name='FieldAutoUpdate']/text()='true' ">
        <w:updateFields w:val="true"/>
      </xsl:if-->

      <!-- Footnotes document wide properties -->
      <xsl:apply-templates
        select="document('styles.xml')/office:document-styles/office:styles/text:notes-configuration[@text:note-class='footnote']"
        mode="note">
        <xsl:with-param name="wide">yes</xsl:with-param>
      </xsl:apply-templates>

      <!-- Endnotes document wide properties -->
      <xsl:apply-templates
        select="document('styles.xml')/office:document-styles/office:styles/text:notes-configuration[@text:note-class='endnote']"
        mode="note">
        <xsl:with-param name="wide">yes</xsl:with-param>
      </xsl:apply-templates>

      <!-- Compatibility settings -->
      <w:compat>
        <!-- Keep space before at top of page. -->
        <w:suppressTopSpacing>
          <xsl:attribute name="w:val">
            <xsl:choose>
              <xsl:when
                test="$configuration-settings/config:config-item[@config:name='AddParaTableSpacingAtStart']/text()='false'">
                <xsl:message terminate="no">translation.odf2oox.spacingTopPageAndTable</xsl:message>
                <xsl:value-of select="'true'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'false'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </w:suppressTopSpacing>
        <w:doNotUseHTMLParagraphAutoSpacing>
          <xsl:attribute name="w:val">
            <xsl:choose>
              <xsl:when
                test="$configuration-settings/config:config-item[@config:name='AddParaTableSpacing']/text()='false'">
                <xsl:value-of select="'false'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'true'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </w:doNotUseHTMLParagraphAutoSpacing>

        <!--divo, dialogika: retain Use Printer Metrics compatibility setting BEGIN -->
        <xsl:if test="$configuration-settings/config:config-item[@config:name='PrinterIndependentLayout']/text()='disabled'">
          <w:usePrinterMetrics/>
        </xsl:if>
        <!--divo, dialogika: retain Use Printer Metrics compatibility setting END -->

      </w:compat>
    </w:settings>
  </xsl:template>



</xsl:stylesheet>

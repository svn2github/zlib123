<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:rels="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ooo="http://openoffice.org/2004/office" xmlns:oox="urn:oox" 
  office:version="1.0"
  exclude-result-prefixes="oox">

  <xsl:template name="settings">
    <office:document-settings>
      <office:settings>
        <config:config-item-set config:name="ooo:view-settings">
          <xsl:choose>
            <xsl:when
              test="key('Part', 'word/document.xml')/w:document/w:body/descendant::w:ins or 
        key('Part', 'word/document.xml')/w:document/w:body/descendant::w:del or
        key('Part', 'word/document.xml')/w:document/w:body/descendant::w:pPrChange or
        key('Part', 'word/document.xml')/w:document/w:body/descendant::w:rPrChange">
              <config:config-item config:name="ShowRedlineChanges" config:type="boolean"
              >true</config:config-item>
            </xsl:when>
            <xsl:otherwise>
              <config:config-item config:name="ShowRedlineChanges" config:type="boolean"
              >false</config:config-item>
            </xsl:otherwise>
          </xsl:choose>
        </config:config-item-set>

        <!-- Uncommenting AddParaTableSpacing will introduce a side effect on line length in OpenOffice 2.x.!!  -->
        <!--
        <config:config-item-set config:name="ooo:configuration-settings">                   
          <config:config-item config:name="AddParaTableSpacing" config:type="boolean">
            <xsl:choose>
              <xsl:when test="key('Part', 'word/settings.xml')/w:settings/w:compat/w:doNotUseHTMLParagraphAutoSpacing">true</xsl:when>
              <xsl:otherwise>false</xsl:otherwise>
            </xsl:choose>                        
          </config:config-item>
          <config:config-item config:name="UseOldNumbering" config:type="boolean">false</config:config-item>
        </config:config-item-set>
          -->
        
        <config:config-item-set config:name="ooo:configuration-settings">

          <xsl:variable name="CompatibilitySettingsFile">
            <xsl:for-each select="key('Part', 'word/_rels/document.xml.rels')/rels:Relationships/rels:Relationship[@Type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml']">
              <xsl:if test="key('Part',substring-after(@Target,'/'))/CompatibilitySettings">               
                <xsl:value-of select="substring-after(@Target,'/')"/>
              </xsl:if>
            </xsl:for-each>
          </xsl:variable>

          <xsl:choose>
            <xsl:when test="$CompatibilitySettingsFile">
              <xsl:for-each select="key('Part', $CompatibilitySettingsFile)/CompatibilitySettings/CompatibilitySetting">
                <config:config-item>
                  <xsl:attribute name="config:name">
                    <xsl:value-of select="@NAME"/>
                  </xsl:attribute>
                  <xsl:attribute name="config:type">
                    <xsl:value-of select="@TYPE"/>
                  </xsl:attribute>
                  <xsl:choose>
                    <xsl:when test="@TYPE = 'boolean' and @VALUE = '1'">
                      <xsl:value-of select="'true'"/>
                    </xsl:when>
                    <xsl:when test="@TYPE = 'boolean' and @VALUE = '0'">
                      <xsl:value-of select="'false'"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="@VALUE"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </config:config-item>
              </xsl:for-each>
            </xsl:when>
            <xsl:otherwise>
              <!--clam, dialogika: re-inserted "true" part (bug #1787090)-->
              <config:config-item config:name="AddParaTableSpacing" config:type="boolean">
                <xsl:choose>
                  <xsl:when test="key('Part', 'word/settings.xml')/w:settings/w:compat/w:doNotUseHTMLParagraphAutoSpacing">true</xsl:when>
                  <xsl:otherwise>false</xsl:otherwise>
                </xsl:choose>
              </config:config-item>

              <!--clam, dialogika: bugfix 1948059-->
              <config:config-item config:name="AddParaSpacingToTableCells" config:type="boolean">true</config:config-item>
              <config:config-item config:name="AddParaTableSpacingAtStart" config:type="boolean">true</config:config-item>

              <!--math, dialogika: Added for correct indentation calculation BEGIN -->

              <config:config-item config:name="IgnoreFirstLineIndentInNumbering" config:type="boolean">
                <xsl:value-of select="'false'"/>
              </config:config-item>

              <!--math, dialogika: Added for correct indentation calculation END -->

              <!--divo, dialogika: retain Use Printer Metrics compatibility setting BEGIN -->
              <config:config-item config:name="PrinterIndependentLayout" config:type="string">
                <xsl:choose>
                  <xsl:when test="key('Part', 'word/settings.xml')/w:settings/w:compat/w:usePrinterMetrics">
                    <xsl:value-of select="'disabled'"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'high-resolution'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </config:config-item>
              <!--divo, dialogika: retain Use Printer Metrics compatibility setting END -->


              <!--
          makz (DIaLOGIKa): 
          Use new text wrapping to emulate Word text wrapping
          -->
              <config:config-item config:name="UseFormerTextWrapping" config:type="boolean">
                <xsl:value-of select="'false'"/>
              </config:config-item>
            </xsl:otherwise>
          </xsl:choose>

         
        </config:config-item-set >

        </office:settings>
    </office:document-settings>
  </xsl:template>

  <xsl:template match="w:compat"> </xsl:template>

</xsl:stylesheet>

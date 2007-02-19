<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ooo="http://openoffice.org/2004/office" office:version="1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main" exclude-result-prefixes="e r">

  <xsl:template name="InsertSettings">
    <office:document-settings>
      <office:settings>
        <config:config-item-set config:name="ooo:view-settings">
          <config:config-item-map-indexed config:name="Views">
            <config:config-item-map-entry>

              <!-- Set default Active Table (Sheet) -->
              <config:config-item config:name="ActiveTable" config:type="string">
                <xsl:for-each select="document('xl/workbook.xml')">
                  <xsl:variable name="ActiveTabNumber">
                    <xsl:choose>
                      <xsl:when test="e:workbook/e:bookViews/e:workbookView/@activeTab">
                        <xsl:value-of select="e:workbook/e:bookViews/e:workbookView/@activeTab"/>
                      </xsl:when>
                      <xsl:otherwise>0</xsl:otherwise>
                    </xsl:choose>
                  </xsl:variable>
                  <xsl:for-each
                    select="e:workbook/e:sheets/e:sheet[position() = $ActiveTabNumber + 1]/@name">
                    <xsl:value-of select="."/>
                  </xsl:for-each>
                </xsl:for-each>
              </config:config-item>

            </config:config-item-map-entry>
          </config:config-item-map-indexed>
        </config:config-item-set>
      </office:settings>
    </office:document-settings>
  </xsl:template>
</xsl:stylesheet>
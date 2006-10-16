<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"
  xmlns:ooo="http://openoffice.org/2004/office" office:version="1.0">
  <xsl:template name="settings">
    <office:document-settings>
      <office:settings>
        <config:config-item-set config:name="ooo:view-settings"></config:config-item-set>
        <config:config-item-set config:name="ooo:configuration-settings"></config:config-item-set>
      </office:settings>
    </office:document-settings>
  </xsl:template>
</xsl:stylesheet>

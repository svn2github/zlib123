<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"
  xmlns:p="http://schemas.openxmlformats.org/presentation/2006/main"
  xmlns:ooo="http://openoffice.org/2004/office" office:version="1.0">

  <xsl:template name="settings">
    <office:document-settings>
      <office:settings>
        <config:config-item-set config:name="ooo:view-settings">
          <xsl:choose>
            <xsl:when 
            test="document('ppt/presentation.xml')/p:document/p:body/descendant::p:ins or 
        document('word/document.xml')/p:document/p:body/descendant::p:del or
        document('word/document.xml')/p:document/p:body/descendant::p:pPrChange or
        document('word/document.xml')/p:document/p:body/descendant::p:rPrChange">
            <config:config-item config:name="ShowRedlineChanges" config:type="boolean" >true</config:config-item>
            </xsl:when>
            <xsl:otherwise>
              <config:config-item config:name="ShowRedlineChanges" config:type="boolean">false</config:config-item>
            </xsl:otherwise>
            </xsl:choose>
        </config:config-item-set>
      </office:settings>
    </office:document-settings>
  </xsl:template>
</xsl:stylesheet>

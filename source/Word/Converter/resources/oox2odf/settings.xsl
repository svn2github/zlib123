<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:ooo="http://openoffice.org/2004/office" office:version="1.0">

  <xsl:template name="settings">
    <office:document-settings>
      <office:settings>
        <config:config-item-set config:name="ooo:view-settings">
          <xsl:choose>
            <xsl:when
              test="document('word/document.xml')/w:document/w:body/descendant::w:ins or 
        document('word/document.xml')/w:document/w:body/descendant::w:del or
        document('word/document.xml')/w:document/w:body/descendant::w:pPrChange or
        document('word/document.xml')/w:document/w:body/descendant::w:rPrChange">
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
              <xsl:when test="document('word/settings.xml')/w:settings/w:compat/w:doNotUseHTMLParagraphAutoSpacing">true</xsl:when>
              <xsl:otherwise>false</xsl:otherwise>
            </xsl:choose>                        
          </config:config-item>
          <config:config-item config:name="UseOldNumbering" config:type="boolean">false</config:config-item>
        </config:config-item-set>
          -->
        
        <config:config-item-set config:name="ooo:configuration-settings">
          <config:config-item config:name="AddParaTableSpacing" config:type="boolean">
            <xsl:value-of select="'false'"/>
          </config:config-item>
        </config:config-item-set >

        </office:settings>
    </office:document-settings>
  </xsl:template>

  <xsl:template match="w:compat"> </xsl:template>

</xsl:stylesheet>

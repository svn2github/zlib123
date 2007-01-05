<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  exclude-result-prefixes="w r xlink number wp ">
  
  <xsl:template match="w:r" mode="trackchanges">
    <xsl:choose>
      <xsl:when test="parent::w:del">
        <text:changed-region>
          <xsl:attribute name="text:id">
            <xsl:value-of select="generate-id(.)"/>
          </xsl:attribute>
          <text:deletion>
            <office:change-info>
              <dc:creator>
                <xsl:value-of select="parent::w:del/@w:author"/>
              </dc:creator>
              <dc:date>
                <xsl:choose>
                  <xsl:when test="contains(parent::w:del/@w:date,'Z')">
                    <xsl:value-of select="substring-before(parent::w:del/@w:date,'Z')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="parent::w:del/@w:date"/>
                  </xsl:otherwise>
                </xsl:choose>
              </dc:date>
            </office:change-info>
            <text:p>
              <xsl:attribute name="text:style-name">
                <xsl:value-of select="generate-id(ancestor::w:p)"/>
              </xsl:attribute>
              <xsl:value-of select="w:delText"/>
            </text:p>
          </text:deletion>
        </text:changed-region>
      </xsl:when>
      <xsl:when test="parent::w:ins">
        <text:changed-region>
          <xsl:attribute name="text:id">
            <xsl:value-of select="generate-id(.)"/>
          </xsl:attribute>
          <text:insertion>
            <office:change-info>
              <dc:creator>
                <xsl:value-of select="parent::w:ins/@w:author"/>
              </dc:creator>
              <dc:date>
                <xsl:choose>
                  <xsl:when test="contains(parent::w:ins/@w:date,'Z')">
                    <xsl:value-of select="substring-before(parent::w:ins/@w:date,'Z')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="parent::w:ins/@w:date"/>
                  </xsl:otherwise>
                </xsl:choose>
              </dc:date>
            </office:change-info>
          </text:insertion>
        </text:changed-region>
      </xsl:when>
      <xsl:when test="descendant::w:rPrChange">
        <text:changed-region>
          <xsl:attribute name="text:id">
            <xsl:value-of select="generate-id(.)"/>
          </xsl:attribute>
          <text:format-change>
            <office:change-info>
              <dc:creator>
                <xsl:value-of select="descendant::w:rPrChange/@w:author"/>
              </dc:creator>
              <dc:date>
                <xsl:choose>
                  <xsl:when test="contains(descendant::w:rPrChange/@w:date,'Z')">
                    <xsl:value-of select="substring-before(descendant::w:rPrChange/@w:date,'Z')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="descendant::w:rPrChange/@w:date"/>
                  </xsl:otherwise>
                </xsl:choose>
              </dc:date>
            </office:change-info>
          </text:format-change>
        </text:changed-region>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
</xsl:stylesheet>

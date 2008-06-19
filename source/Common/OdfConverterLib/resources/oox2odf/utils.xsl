<?xml version="1.0" encoding="UTF-8"?>
<?xml-stylesheet href="utils.xsl" type="text/xsl" media="screen"?>

<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
  
  <xsl:template name="substring-after-last">
    <xsl:param name="string" />
    <xsl:param name="occurrence" />

    <xsl:choose>
      <xsl:when test="contains($string, $occurrence)">
        <xsl:call-template name="substring-after-last">
          <xsl:with-param name="string" select="substring-after($string, $occurrence)" />
          <xsl:with-param name="occurrence" select="$occurrence"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$string"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
</xsl:stylesheet>
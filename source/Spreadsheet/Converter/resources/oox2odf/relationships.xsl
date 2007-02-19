<?xml version="1.0" encoding="UTF-8"?>
<!--<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:uri="http://schemas.openxmlformats.org/drawingml/2006/picture"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" version="1.0">-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:e="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  exclude-result-prefixes="e a pic r w">

  <xsl:template name="GetTarget">
    <xsl:param name="document"/>
    <xsl:param name="file" select="$document"/>
    <xsl:param name="id"/>

    <xsl:choose>
      <xsl:when test="contains($file,'/')">
        <xsl:call-template name="GetTarget">
          <xsl:with-param name="file">
            <xsl:value-of select="substring-after($file,'/')"/>
          </xsl:with-param>
          <xsl:with-param name="id" select="$id"/>
          <xsl:with-param name="document" select="$document"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="document(concat(substring-before($document,$file),'_rels/',$file,'.rels'))">
          <xsl:for-each
            select="document(concat(substring-before($document,$file),'_rels/',$file,'.rels'))//node()[name()='Relationship']">
            <xsl:if test="./@Id=$id">
              <xsl:value-of select="./@Target"/>
              </xsl:if>
          </xsl:for-each>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>

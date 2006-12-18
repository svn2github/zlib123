<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">

  <xsl:template match="w:sectPr[parent::w:pPr]" mode="sections">
    <xsl:variable name="id">
      <xsl:value-of select="generate-id(preceding::w:p/w:pPr/w:sectPr)"/>
    </xsl:variable>
    <xsl:variable name="id2">
      <xsl:value-of select="generate-id(.)"/>
    </xsl:variable>
    <text:section>
      <xsl:attribute name="text:style-name">
        <xsl:value-of select="$id2"/>
      </xsl:attribute>
      <xsl:attribute name="text:name">
        <xsl:value-of select="concat('S_',$id2)"/>
      </xsl:attribute>
      <xsl:choose>
        <xsl:when 
          test="./w:titlePg">
        <text:p>
          <xsl:attribute name="text:style-name">
            <xsl:value-of select="concat('F_P_',$id2)"/>
          </xsl:attribute>
        </text:p>
      </xsl:when>
      <xsl:when
        test="preceding::w:sectPr/w:pgSz/@w:w != ./w:pgSz/@w:w
        or preceding::w:sectPr/w:pgSz/@w:h != ./w:pgSz/@w:h
        or preceding::w:sectPr/w:pgSz/@w:orient != ./w:pgSz/@w:orient
        or preceding::w:sectPr/w:pgMar/@w:top != ./w:pgMar/@w:top
        or preceding::w:sectPr/w:pgMar/@w:left != ./w:pgMar/@w:left
        or preceding::w:sectPr/w:pgMar/@w:right != ./w:pgMar/@w:right
        or preceding::w:sectPr/w:pgMar/@w:bottom != ./w:pgMar/@w:bottom
        or preceding::w:sectPr/w:pgMar/@w:header != ./w:pgMar/@w:header
        or preceding::w:sectPr/w:pgMar/@w:footer != ./w:pgMar/@w:footer
        or ./w:headerReference
        or ./w:footerReference
        or document('word/document.xml')/w:document/w:body/w:sectPr/w:pgSz/@w:w != ./w:pgSz/@w:w 
        or document('word/document.xml')/w:document/w:body/w:sectPr/w:pgSz/@w:h != ./w:pgSz/@w:h
        or document('word/document.xml')/w:document/w:body/w:sectPr/w:pgSz/@w:orient != ./w:pgSz/@w:orient">
        <text:p>
          <xsl:attribute name="text:style-name">
            <xsl:value-of select="concat('P_',$id2)"/>
          </xsl:attribute>
        </text:p>
      </xsl:when>
        </xsl:choose>
      <xsl:apply-templates
        select="document('word/document.xml')/w:document/w:body/child::node()[(generate-id(following::w:sectPr) = $id2 and generate-id(.) != $id2 and generate-id(.) != $id and not(descendant::w:sectPr)) or generate-id(descendant::w:sectPr) = $id2]"/>
      <xsl:if
        test="(preceding::w:sectPr/w:pgSz/@w:w != ./w:pgSz/@w:w
        or preceding::w:sectPr/w:pgSz/@w:h != ./w:pgSz/@w:h
        or preceding::w:sectPr/w:pgSz/@w:orient != ./w:pgSz/@w:orient
        or preceding::w:sectPr/w:pgMar/@w:top != ./w:pgMar/@w:top
        or preceding::w:sectPr/w:pgMar/@w:left != ./w:pgMar/@w:left
        or preceding::w:sectPr/w:pgMar/@w:right != ./w:pgMar/@w:right
        or preceding::w:sectPr/w:pgMar/@w:bottom != ./w:pgMar/@w:bottom
        or preceding::w:sectPr/w:pgMar/@w:header != ./w:pgMar/@w:header
        or preceding::w:sectPr/w:pgMar/@w:footer != ./w:pgMar/@w:footer
        or document('word/document.xml')/w:document/w:body/w:sectPr/w:pgSz/@w:w != ./w:pgSz/@w:w
        or document('word/document.xml')/w:document/w:body/w:sectPr/w:pgSz/@w:h != ./w:pgSz/@w:h
        or document('word/document.xml')/w:document/w:body/w:sectPr/w:pgSz/@w:orient != ./w:pgSz/@w:orient
        or following::w:sectPr[1]/w:headerReference
        or following::w:sectPr[1]/w:footerReference)
        and (not(following::w:p/w:pPr/w:sectPr) and not(document('word/document.xml')/w:document/w:body/w:sectPr/w:titlePg))">
        <text:p>
          <xsl:attribute name="text:style-name">
            <xsl:text>P_Standard</xsl:text>
          </xsl:attribute>
        </text:p>
      </xsl:if>
    </text:section>
  </xsl:template>

  <xsl:template name="ParagraphFromSectionsStyles">
    <xsl:for-each select="document('word/document.xml')/w:document/w:body/w:p/w:pPr/w:sectPr">
      <xsl:if test="(preceding::w:sectPr/w:pgSz/@w:w != ./w:pgSz/@w:w or preceding::w:sectPr/w:pgSz/@w:h != ./w:pgSz/@w:h or preceding::w:sectPr/w:pgSz/@w:orient != ./w:pgSz/@w:orient or document('word/document.xml')/w:document/w:body/w:sectPr/w:pgSz/@w:w != ./w:pgSz/@w:w or document('word/document.xml')/w:document/w:body/w:sectPr/w:pgSz/@w:h != ./w:pgSz/@w:h or document('word/document.xml')/w:document/w:body/w:sectPr/w:pgSz/@w:orient != ./w:pgSz/@w:orient) and not(./w:headerReference) and not(./w:footerReference)">
        <style:style>
          <xsl:attribute name="style:name">
            <xsl:value-of select="concat('P_',generate-id(.))"/>
          </xsl:attribute>
          <xsl:attribute name="style:family">
            <xsl:text>paragraph</xsl:text>
          </xsl:attribute>
          <xsl:attribute name="style:master-page-name">
            <xsl:value-of select="concat('PAGE_',generate-id(.))"/>
          </xsl:attribute>
          <style:text-properties fo:language="en" fo:country="US" text:display="true"/>
        </style:style>
      </xsl:if>
      <style:style>
        <xsl:attribute name="style:name">
          <xsl:value-of select="concat('F_P_',generate-id(.))"/>
        </xsl:attribute>
        <xsl:attribute name="style:family">
          <xsl:text>paragraph</xsl:text>
        </xsl:attribute>
        <xsl:attribute name="style:master-page-name">
          <xsl:value-of select="concat('First_H_',generate-id(.))"/>
        </xsl:attribute>
        <style:text-properties fo:language="en" fo:country="US" text:display="true"/>
      </style:style>
      <style:style>
        <xsl:attribute name="style:name">
          <xsl:value-of select="concat('P_',generate-id(.))"/>
        </xsl:attribute>
        <xsl:attribute name="style:family">
          <xsl:text>paragraph</xsl:text>
        </xsl:attribute>
        <xsl:attribute name="style:master-page-name">
          <xsl:value-of select="concat('H_',generate-id(.))"/>
        </xsl:attribute>
        <style:text-properties fo:language="en" fo:country="US" text:display="true"/>
      </style:style>
    </xsl:for-each>
    <xsl:if
      test="document('word/document.xml')/w:document/w:body/w:sectPr/w:titlePg">
      <style:style>
        <xsl:attribute name="style:name">
          <xsl:text>P_F</xsl:text>
        </xsl:attribute>
        <xsl:attribute name="style:family">
          <xsl:text>paragraph</xsl:text>
        </xsl:attribute>
        <xsl:attribute name="style:parent-style-name">
          <xsl:text>Standard</xsl:text>
        </xsl:attribute>
        <xsl:attribute name="style:master-page-name">
          <xsl:text>First_Page</xsl:text>
        </xsl:attribute>
        <style:text-properties fo:language="en" fo:country="US" text:display="true"/>
      </style:style>
    </xsl:if>
    <style:style>
      <xsl:attribute name="style:name">
        <xsl:text>P_Standard</xsl:text>
      </xsl:attribute>
      <xsl:attribute name="style:family">
        <xsl:text>paragraph</xsl:text>
      </xsl:attribute>
      <xsl:attribute name="style:master-page-name">
        <xsl:text>Standard</xsl:text>
      </xsl:attribute>
      <style:text-properties fo:language="en" fo:country="US" text:display="true"/>
    </style:style>
  </xsl:template>
  
  <xsl:template name="InsertColumns">
    
    <xsl:choose>
      <xsl:when test="w:cols/@w:num">
        <style:columns>
          <xsl:attribute name="fo:column-count">
            <xsl:value-of select="w:cols/@w:num"/>
          </xsl:attribute>
          <xsl:choose>
            <xsl:when test="w:cols/@w:equalWidth = '0'">
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="fo:column-gap">
                <xsl:call-template name="ConvertTwips">
                  <xsl:with-param name="length">
                    <xsl:value-of select="w:cols/@w:space"/>
                  </xsl:with-param>
                  <xsl:with-param name="unit">cm</xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
          <xsl:if test="w:cols/@w:sep = '1'">
            <style:column-sep style:width="0.002cm" style:color="#000000" style:height="100%"/>
          </xsl:if>
          <xsl:for-each select="w:cols/w:col">
            <style:column>
              <xsl:attribute name="style:rel-width">
                <xsl:variable name="width">
                  <xsl:value-of select="./@w:w"/>
                </xsl:variable>
                <xsl:choose>
                  <xsl:when test="preceding-sibling::w:col/@w:space">
                    <xsl:variable name="space">
                      <xsl:value-of select="format-number(preceding-sibling::w:col/@w:space div 2,'#')"/>
                    </xsl:variable>
                    <xsl:value-of select="concat($width + $space,'*')"/>
                  </xsl:when>
                  <xsl:when test="./@w:space">
                    <xsl:variable name="space">
                      <xsl:value-of select="format-number(./@w:space div 2,'#')"/>
                    </xsl:variable>
                    <xsl:value-of select="concat($width + $space,'*')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="$width"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
              <xsl:attribute name="fo:start-indent">
                <xsl:choose>
                  <xsl:when test="preceding-sibling::w:col/@w:space">
                    <xsl:variable name="width">
                      <xsl:call-template name="ConvertTwips">
                        <xsl:with-param name="length">
                          <xsl:value-of select="format-number(preceding-sibling::w:col/@w:space div 2, '#.###')"/>
                        </xsl:with-param>
                        <xsl:with-param name="unit">cm</xsl:with-param>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:value-of select="$width"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:text>0</xsl:text>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
              <xsl:attribute name="fo:end-indent">
                <xsl:choose>
                  <xsl:when test="./@w:space">
                    <xsl:variable name="width">
                      <xsl:call-template name="ConvertTwips">
                        <xsl:with-param name="length">
                          <xsl:value-of select="format-number(./@w:space div 2, '#.###')"/>
                        </xsl:with-param>
                        <xsl:with-param name="unit">cm</xsl:with-param>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:value-of select="$width"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:text>0</xsl:text>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </style:column>
          </xsl:for-each>
        </style:columns>
      </xsl:when>
      <xsl:otherwise>
        <style:columns>
          <xsl:attribute name="fo:column-count">
            <xsl:text>1</xsl:text>
          </xsl:attribute>
        </style:columns>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template match="w:sectPr[parent::w:pPr]" mode="automaticstyles">
    <style:style style:name="{generate-id(.)}" style:family="section">
      <style:section-properties>
        <xsl:call-template name="InsertColumns"/>
      </style:section-properties>      
    </style:style>
  </xsl:template>
  
  <xsl:template match="text()" mode="sections"/>
</xsl:stylesheet>

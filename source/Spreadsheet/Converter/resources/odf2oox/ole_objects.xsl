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
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" exclude-result-prefixes="table">

  <xsl:template name="InsertOLE_Object">
    <xsl:if test="table:shapes/draw:frame/draw:object">
      <oleObjects>
        <xsl:apply-templates select="table:shapes/draw:frame[1]" mode="OLEobject"/>
      </oleObjects>
    </xsl:if>
  </xsl:template>

  
  <xsl:template match="draw:frame" mode="OLEobject">
    <xsl:param name="LinkId" select="1"/>

    <xsl:variable name="apos">
      <xsl:text>&apos;&apos;&apos;&apos;</xsl:text>
    </xsl:variable>
    <oleObject progId="opendocument.WriterDocument.1" oleUpdate="OLEUPDATE_ALWAYS">
      <xsl:attribute name="link">
        <xsl:value-of select="concat('[', $LinkId, ']!', $apos)"/>
      </xsl:attribute>

      <xsl:attribute name="shapeId">
        <xsl:value-of select="1025 + count(preceding-sibling::draw:frame) + 1"/>
      </xsl:attribute>

      <xsl:attribute name="progId">
        <xsl:value-of select="generate-id()"/>
      </xsl:attribute>
    </oleObject>
    <!--xsl:value-of select="$LinkId"/-->

    <xsl:if test="following-sibling::draw:frame">
      <xsl:apply-templates select="following-sibling::draw:frame[1]" mode="OLEobject">
        <xsl:with-param name="LinkId">
          <xsl:value-of select="$LinkId + 1"/>
        </xsl:with-param>
      </xsl:apply-templates>
    </xsl:if>

  </xsl:template>

  <!--xsl:template name="InsertVMLDrawing">
    <v:shape id="_x0000_s1025" type="#_x0000_t75"
      style="position:absolute;
        margin-left:336pt;margin-top:150pt;width:34.5pt;height:38.25pt;z-index:1"
      filled="t" fillcolor="window [65]" stroked="t" strokecolor="windowText [64]"
      o:insetmode="auto">
      <v:fill color2="window [65]"/>
      <v:imagedata o:relid="rId1" o:title=""/>
      <x:ClientData ObjectType="Pict">
        <x:SizeWithCells/>
        <x:Anchor> 7, 0, 10, 0, 7, 46, 12, 11</x:Anchor>
        <x:CF>Pict</x:CF>
        <x:AutoPict/>
        <x:DDE/>
      </x:ClientData>
    </v:shape>
  </xsl:template-->

  <xsl:template name="InsertOLE_rels">
    
      <pzip:entry pzip:target="{concat('xl/drawings/_rels/vmlDrawing',position(),'.vml.rels')}">
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            Target="../media/image1.emf"/>
        </Relationships>
      </pzip:entry>
    
  </xsl:template>

  <xsl:template name="InsertOLEexternalLinks">
    <xsl:for-each select="table:shapes/draw:frame">
      <pzip:entry pzip:target="{concat('xl/externalLinks/externalLink',position(),'.xml')}">
        <externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <oleLink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            r:id="rId1">
            <xsl:attribute name="progId">
              <xsl:value-of select="generate-id()"/>
            </xsl:attribute>
            <oleItems>
              <oleItem name="'" advise="1" preferPic="1"/>
            </oleItems>
          </oleLink>
        </externalLink>
      </pzip:entry>
    </xsl:for-each>
  </xsl:template>


  <xsl:template name="OLEexternalLinks_rels">
    <xsl:for-each select="table:shapes/draw:frame/draw:object">
      <pzip:entry
        pzip:target="{concat('xl/externalLinks/_rels/externalLink',position(),'.xml.rels')}">
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
            TargetMode="External">

            <xsl:attribute name="Target">
              <xsl:if test="@xlink:href">

                <xsl:choose>
                  <!-- when hyperlink to a site or mailto-->
                  <xsl:when test="contains(@xlink:href,':') and not(starts-with(@xlink:href,'/'))">
                    <xsl:call-template name="TranslateIllegalChars">
                      <xsl:with-param name="string" select="@xlink:href"/>
                    </xsl:call-template>
                  </xsl:when>
                  <!-- when hyperlink to an document-->
                  <xsl:otherwise>

                    <xsl:variable name="translatedTarget">
                      <xsl:call-template name="SpaceTo20Percent">
                        <xsl:with-param name="string" select="@xlink:href"/>
                      </xsl:call-template>
                    </xsl:variable>

                    <xsl:choose>
                      <!-- when starts with up folder sign -->
                      <xsl:when test="starts-with($translatedTarget,'../' )">
                        <xsl:value-of
                          select="translate(substring-after($translatedTarget,'../'),'/','\')"/>
                      </xsl:when>
                      <!-- when file is in another disk -->
                      <xsl:when test="starts-with($translatedTarget,'/')">
                        <xsl:value-of
                          select="concat('file:///',translate(substring-after($translatedTarget,'/'),'/','\'))"
                        />
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="translate($translatedTarget,'/','\')"/>
                      </xsl:otherwise>
                    </xsl:choose>

                  </xsl:otherwise>
                </xsl:choose>

              </xsl:if>
            </xsl:attribute>

          </Relationship>
        </Relationships>
      </pzip:entry>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="ExternalReference">
    <xsl:for-each select="table:shapes">
      <xsl:if test="draw:frame">
      <externalReferences>
        <xsl:for-each select="draw:frame">
          <externalReference>
            <xsl:attribute name="r:id">
              <xsl:value-of select="generate-id()"/>
            </xsl:attribute>
          </externalReference>
        </xsl:for-each>
      </externalReferences>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

</xsl:stylesheet>

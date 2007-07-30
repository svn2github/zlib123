<?xml version="1.0" encoding="utf-8"?>
<!--
Copyright (c) 2007, Sonata Software Limited
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
*     * Neither the name of Sonata Software Limited nor the names of its contributors
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
* SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE
-->
<xsl:stylesheet version="1.0"
xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
xmlns:xlink="http://www.w3.org/1999/xlink"
xmlns:dom="http://www.w3.org/2001/xml-events"
xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
xmlns:rels="http://schemas.openxmlformats.org/package/2006/relationships"
xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
xmlns:dc="http://purl.org/dc/elements/1.1/"
xmlns:dcterms="http://purl.org/dc/terms/"
exclude-result-prefixes="p a r xlink ">
  <xsl:import href="common.xsl"/>
  <xsl:template name="InsertPicture">
    <xsl:param name ="slideRel"/>
    <xsl:variable name ="imageId">
      <xsl:value-of select ="./p:blipFill/a:blip/@r:embed"/>
    </xsl:variable>
    <xsl:variable name ="sourceFiile">
      <xsl:for-each select ="document($slideRel)//node()[@Id = $imageId]">        
          <xsl:value-of select ="@Target"/>        
      </xsl:for-each>
    </xsl:variable >
    <xsl:variable name ="targetFile">
      <xsl:value-of select ="substring-after(substring-after($sourceFiile,'/'),'/')"/>
    </xsl:variable>
    <pzip:copy pzip:source="{concat('ppt',substring-after($sourceFiile,'..'))}"
				   pzip:target="{concat('Pictures/',$targetFile)}"/>
    <draw:frame draw:layer="layout">
      <!--Edited by vipul to get cordinates from Layout-->
      <xsl:choose>
        <xsl:when test ="p:spPr/a:xfrm/a:off">
      <xsl:call-template name="tmpWriteCordinates"/>
        </xsl:when>
        <xsl:when test ="not(p:spPr/a:xfrm/a:off)">
          <xsl:variable name ="LayoutFileNo">
            <xsl:for-each select ="document($slideRel)//node()/@Target[contains(.,'slideLayouts')]">
              <xsl:value-of select ="concat('ppt',substring(.,3))"/>
            </xsl:for-each>
          </xsl:variable>
          <xsl:for-each select ="document($LayoutFileNo)/p:sldLayout/p:cSld/p:spTree/p:sp">
            <xsl:if test="p:nvSpPr/p:nvPr/p:ph[@type='pic']">
              <xsl:call-template name="tmpWriteCordinates"/>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
      <draw:image xlink:href="Pictures/100000000000032000000258B0234CE5.jpg" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad">
        <xsl:attribute name ="xlink:href"  >
          <xsl:value-of select ="concat('Pictures/',$targetFile)"/>
        </xsl:attribute>
        <text:p />
      </draw:image>
    </draw:frame>
  </xsl:template>
  
</xsl:stylesheet>
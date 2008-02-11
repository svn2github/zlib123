<?xml version="1.0" encoding="utf-8" ?>
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
<xsl:stylesheet version="2.0" 
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink" 
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
  xmlns:rels="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"
  exclude-result-prefixes="p a r dc xlink draw rels">
    <!-- Shape constants-->
	<!-- Arrow size -->
	<xsl:variable name="sm-sm">
		<xsl:value-of select ="'0.14'"/>
	</xsl:variable>
	<xsl:variable name="sm-med">
		<xsl:value-of select ="'0.245'"/>
	</xsl:variable>
	<xsl:variable name="sm-lg">
		<xsl:value-of select ="'0.2'"/>
	</xsl:variable>
	<xsl:variable name="med-sm">
		<xsl:value-of select ="'0.234'" />
	</xsl:variable>
	<xsl:variable name="med-med">
		<xsl:value-of select ="'0.351'"/>
	</xsl:variable>
	<xsl:variable name="med-lg">
		<xsl:value-of select ="'0.3'" />
	</xsl:variable>
	<xsl:variable name="lg-sm">
		<xsl:value-of select ="'0.31'" />
	</xsl:variable>
	<xsl:variable name="lg-med">
		<xsl:value-of select ="'0.35'" />
	</xsl:variable>
	<xsl:variable name="lg-lg">
		<xsl:value-of select ="'0.4'" />
	</xsl:variable>

  <xsl:template name ="drawShapes">
    <xsl:param name ="SlideID" />
    <xsl:param name ="SlideRelationId" />
    <xsl:param name ="grID" />
    <xsl:param name ="prID" />

    <xsl:call-template name ="shapes">
      <!-- Extra parameter "slideId" added by lohith,requierd for template AddTextHyperlinks -->
      <xsl:with-param name="slideId" select ="$SlideID"/>
      <xsl:with-param name="GraphicId" select ="concat($SlideID,$grID,position())"/>
      <xsl:with-param name ="ParaId" select ="concat($SlideID,$prID,position())" />
      <xsl:with-param name ="SlideRelationId" select="$SlideRelationId" />
    </xsl:call-template>
  </xsl:template>

  <!-- Template for Shapes in reverse conversion -->
  <xsl:template  name="shapes">
    <xsl:param name="GraphicId" />
    <xsl:param name="ParaId" />
    <xsl:param name="SlideRelationId" />
    <xsl:param name="TypeId" />
    <xsl:param name="var_pos" />
    <xsl:param name="grpBln" />
    <xsl:param name ="grpCordinates"/>
    <!-- Extra parameter "slideId" added by lohith,requierd for template AddTextHyperlinks -->
    <xsl:param name="slideId" />
    <xsl:variable name="varHyperLinksForShapes">
      <!-- Added by lohith.ar - Start - Mouse click hyperlinks -->
      <office:event-listeners>
        <xsl:for-each select ="p:nvSpPr/p:cNvPr">
            <xsl:choose>
            <!-- Start => Go to previous slide-->
            <xsl:when test="a:hlinkClick/@action[contains(.,'jump=previousslide')]">
              <presentation:event-listener>
                <xsl:attribute name ="script:event-name">
                  <xsl:value-of select ="'dom:click'"/>
                </xsl:attribute>
                <xsl:attribute name ="presentation:action">
                  <xsl:value-of select ="'previous-page'"/>
                </xsl:attribute>
              </presentation:event-listener>
            </xsl:when>
            <!-- End => Go to previous slide-->
            <!-- Start => Go to Next slide -->
            <xsl:when test="a:hlinkClick/@action[contains(.,'jump=nextslide')]">
              <presentation:event-listener>
                <xsl:attribute name ="script:event-name">
                  <xsl:value-of select ="'dom:click'"/>
                </xsl:attribute>
                <xsl:attribute name ="presentation:action">
                  <xsl:value-of select ="'next-page'"/>
                </xsl:attribute>
              </presentation:event-listener>
            </xsl:when>
            <!-- End => Go to Next slide-->
            <!-- Start => Go to First slide -->
            <xsl:when test="a:hlinkClick/@action[contains(.,'jump=firstslide')]">
              <presentation:event-listener>
                <xsl:attribute name ="script:event-name">
                  <xsl:value-of select ="'dom:click'"/>
                </xsl:attribute>
                <xsl:attribute name ="presentation:action">
                  <xsl:value-of select ="'first-page'"/>
                </xsl:attribute>
              </presentation:event-listener>
            </xsl:when>
            <!-- End => Go to First slide -->
            <!-- Start => Go to Last slide -->
            <xsl:when test="a:hlinkClick/@action[contains(.,'jump=lastslide')]">
              <presentation:event-listener>
                <xsl:attribute name ="script:event-name">
                  <xsl:value-of select ="'dom:click'"/>
                </xsl:attribute>
                <xsl:attribute name ="presentation:action">
                  <xsl:value-of select ="'last-page'"/>
                </xsl:attribute>
              </presentation:event-listener>
            </xsl:when>
            <!-- End => Go to Last slide -->
            <!-- Start => EndShow -->
            <xsl:when test="a:hlinkClick/@action[contains(.,'jump=endshow')]">
              <presentation:event-listener>
                <xsl:attribute name ="script:event-name">
                  <xsl:value-of select ="'dom:click'"/>
                </xsl:attribute>
                <xsl:attribute name ="presentation:action">
                  <xsl:value-of select ="'stop'"/>
                </xsl:attribute>
              </presentation:event-listener>
            </xsl:when>
            <!-- End => End show -->
            <!-- Start => Go to Page or Object && Go to Other document && Run program  -->
            <xsl:when test="a:hlinkClick/@action[contains(.,'ppaction://hlinksldjump') or contains(.,'ppaction://hlinkfile') or contains(.,'ppaction://program') ]">
              <presentation:event-listener>
                <xsl:attribute name ="script:event-name">
                  <xsl:value-of select ="'dom:click'"/>
                </xsl:attribute>
                <xsl:variable name="RelationId">
                  <xsl:value-of select="a:hlinkClick/@r:id"/>
                </xsl:variable>
                <xsl:variable name="SlideVal">
                  <xsl:value-of select="document($SlideRelationId)/rels:Relationships/rels:Relationship[@Id=$RelationId]/@Target"/>
                </xsl:variable>
                <!-- Condn Go to Other page/slide-->
                <xsl:if test="a:hlinkClick/@action[contains(.,'ppaction://hlinksldjump')]">
                  <xsl:attribute name ="presentation:action">
                    <xsl:value-of select ="'show'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="xlink:href">
                    <xsl:value-of select ="concat('#page',substring-before(substring-after($SlideVal,'slide'),'.xml'))"/>
                  </xsl:attribute>
                </xsl:if>
                <!-- Condn Go to Other document-->
                <xsl:if test="a:hlinkClick/@action[contains(.,'ppaction://hlinkfile')]">
                  <xsl:attribute name ="presentation:action">
                    <xsl:value-of select ="'show'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="xlink:href">
                    <xsl:if test="string-length(substring-after($SlideVal,'file:///')) > 0">
                      <xsl:value-of select ="concat('/',translate(substring-after($SlideVal,'file:///'),'\','/'))"/>
                    </xsl:if>
                    <xsl:if test="string-length(substring-after($SlideVal,'file:///')) = 0 ">
                      <xsl:value-of select ="concat('../',$SlideVal)"/>
                    </xsl:if>
                  </xsl:attribute>
                </xsl:if>
                <!-- Condn Go to Run program-->
                <xsl:if test="a:hlinkClick/@action[contains(.,'ppaction://program') ]">
                  <xsl:attribute name ="presentation:action">
                    <xsl:value-of select ="'execute'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="xlink:href">
                    <xsl:if test="string-length(substring-after($SlideVal,'file:///')) > 0">
                      <xsl:value-of select ="concat('/',translate(substring-after($SlideVal,'file:///'),'\','/'))"/>
                    </xsl:if>
                    <xsl:if test="string-length(substring-after($SlideVal,'file:///')) = 0">
                      <xsl:value-of select ="concat('../',$SlideVal)"/>
                    </xsl:if>
                  </xsl:attribute>
                </xsl:if>
                <xsl:attribute name ="xlink:type">
                  <xsl:value-of select ="'simple'"/>
                </xsl:attribute>
                <xsl:attribute name ="xlink:show">
                  <xsl:value-of select ="'new'"/>
                </xsl:attribute>
                <xsl:attribute name ="xlink:actuate">
                  <xsl:value-of select ="'onRequest'"/>
                </xsl:attribute>
              </presentation:event-listener>
            </xsl:when>
            <!-- End => Go to Page or Object && Go to Other document && Run program -->
            <!-- Start => Play sound  -->
            <xsl:when test="a:hlinkClick/a:snd">
              <presentation:event-listener>
                <xsl:attribute name ="script:event-name">
                  <xsl:value-of select ="'dom:click'"/>
                </xsl:attribute>
                <xsl:attribute name ="presentation:action">
                  <xsl:value-of select ="'sound'"/>
                </xsl:attribute>
                <presentation:sound>
                  <xsl:variable name="varMediaFileRelId">
                    <xsl:value-of select="a:hlinkClick/a:snd/@r:embed"/>
                  </xsl:variable>
                  <xsl:variable name="varMediaFileTargetPath">
                    <xsl:value-of select="document($SlideRelationId)/rels:Relationships/rels:Relationship[@Id=$varMediaFileRelId]/@Target"/>
                  </xsl:variable>
                  <xsl:variable name="varPptMediaFileTargetPath">
                    <xsl:value-of select="concat('ppt/',substring-after($varMediaFileTargetPath,'/'))"/>
                  </xsl:variable>
                  <!--<xsl:variable name="varDocumentModifiedTime">
                    <xsl:value-of select="document('docProps/core.xml')/cp:coreProperties/dcterms:modified"/>
                  </xsl:variable>-->
                  <!--<xsl:variable name="varDestMediaFileTargetPath">
                    <xsl:value-of select="concat(translate($varDocumentModifiedTime,':','-'),'|',$varMediaFileRelId,'.wav')"/>
                  </xsl:variable>-->
                  <!--<xsl:variable name="varMediaFilePathForOdp">
                    <xsl:value-of select="concat('../_MediaFilesForOdp_',translate($varDocumentModifiedTime,':','-'),'/',$varMediaFileRelId,'.wav')"/>
                  </xsl:variable>-->
          
                  <xsl:variable name="FolderNameGUID">
                    <xsl:call-template name="GenerateGUIDForFolderName">
                      <xsl:with-param name="RootNode" select="." />
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:variable name="varDestMediaFileTargetPath">
                    <xsl:value-of select="concat($FolderNameGUID,'|',$varMediaFileRelId,'.wav')"/>
                  </xsl:variable>
                  <xsl:variable name="varMediaFilePathForOdp">
                    <xsl:value-of select="concat('../',$FolderNameGUID,'/',$varMediaFileRelId,'.wav')"/>
                  </xsl:variable>
                  <xsl:attribute name ="xlink:href">
                    <xsl:value-of select ="$varMediaFilePathForOdp"/>
                  </xsl:attribute>
                  <xsl:attribute name ="xlink:type">
                    <xsl:value-of select ="'simple'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="xlink:show">
                    <xsl:value-of select ="'new'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="xlink:actuate">
                    <xsl:value-of select ="'onRequest'"/>
                  </xsl:attribute>
                  <pzip:extract pzip:source="{$varPptMediaFileTargetPath}" pzip:target="{$varDestMediaFileTargetPath}" />
                </presentation:sound>
              </presentation:event-listener>
            </xsl:when>
              <xsl:otherwise>
                <xsl:if test="a:hlinkClick/@r:id">
                <presentation:event-listener script:event-name="dom:click" presentation:action="show"
                                      xlink:type="simple" xlink:show="new" xlink:actuate="onRequest">
                  <xsl:variable name="RelationId">
                    <xsl:value-of select="a:hlinkClick/@r:id"/>
                  </xsl:variable>
                  <xsl:variable name="Target">
                    <xsl:value-of select="document($SlideRelationId)/rels:Relationships/rels:Relationship[@Id=$RelationId]/@Target"/>
                  </xsl:variable>
                  <xsl:variable name="type">
                    <xsl:value-of select="document($SlideRelationId)/rels:Relationships/rels:Relationship[@Id=$RelationId]/@Type"/>
                  </xsl:variable>
                  <xsl:choose>
                    <xsl:when test="contains($Target,'mailto:') or contains($Target,'http:') or contains($Target,'https:')">
                      <xsl:attribute name="xlink:href">
                        <xsl:value-of select="$Target"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test="contains($Target,'slide')">
                      <xsl:attribute name="xlink:href">
                        <xsl:value-of select="concat('#Slide ',substring-before(substring-after($Target,'slide'),'.xml'))"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test="contains($Target,'file:///')">
                      <xsl:attribute name="xlink:href">
                        <xsl:value-of select="concat('/',translate(substring-after($Target,'file:///'),'\','/'))"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test="contains($type,'hyperlink') and not(contains($Target,'http:')) and not(contains($Target,'https:'))">
                      <!-- warn if hyperlink Path  -->
                      <xsl:message terminate="no">translation.oox2odf.hyperlinkTypeRelativePath</xsl:message>
                      <xsl:attribute name="xlink:href">
                        <!--Link Absolute Path-->
                        <xsl:value-of select ="concat('hyperlink-path:',$Target)"/>
                        <!--End-->
                      </xsl:attribute>
                    </xsl:when>
                  </xsl:choose>
                </presentation:event-listener>
                </xsl:if>
              </xsl:otherwise>
            <!-- End => Play sound  -->
          </xsl:choose>
        </xsl:for-each>
      </office:event-listeners>
      <!-- End - Mouse click hyperlinks-->
    </xsl:variable>
    <xsl:variable name="varHyperLinksForConnectors">
      <!-- Added by lohith.ar - Start - Mouse click hyperlinks -->
      <office:event-listeners>
        <xsl:for-each select ="p:nvCxnSpPr/p:cNvPr">
          <xsl:choose>
            <!-- Start => Go to previous slide-->
            <xsl:when test="a:hlinkClick/@action[contains(.,'jump=previousslide')]">
              <presentation:event-listener>
                <xsl:attribute name ="script:event-name">
                  <xsl:value-of select ="'dom:click'"/>
                </xsl:attribute>
                <xsl:attribute name ="presentation:action">
                  <xsl:value-of select ="'previous-page'"/>
                </xsl:attribute>
              </presentation:event-listener>
            </xsl:when>
            <!-- End => Go to previous slide-->
            <!-- Start => Go to Next slide -->
            <xsl:when test="a:hlinkClick/@action[contains(.,'jump=nextslide')]">
              <presentation:event-listener>
                <xsl:attribute name ="script:event-name">
                  <xsl:value-of select ="'dom:click'"/>
                </xsl:attribute>
                <xsl:attribute name ="presentation:action">
                  <xsl:value-of select ="'next-page'"/>
                </xsl:attribute>
              </presentation:event-listener>
            </xsl:when>
            <!-- End => Go to Next slide-->
            <!-- Start => Go to First slide -->
            <xsl:when test="a:hlinkClick/@action[contains(.,'jump=firstslide')]">
              <presentation:event-listener>
                <xsl:attribute name ="script:event-name">
                  <xsl:value-of select ="'dom:click'"/>
                </xsl:attribute>
                <xsl:attribute name ="presentation:action">
                  <xsl:value-of select ="'first-page'"/>
                </xsl:attribute>
              </presentation:event-listener>
            </xsl:when>
            <!-- End => Go to First slide -->
            <!-- Start => Go to Last slide -->
            <xsl:when test="a:hlinkClick/@action[contains(.,'jump=lastslide')]">
              <presentation:event-listener>
                <xsl:attribute name ="script:event-name">
                  <xsl:value-of select ="'dom:click'"/>
                </xsl:attribute>
                <xsl:attribute name ="presentation:action">
                  <xsl:value-of select ="'last-page'"/>
                </xsl:attribute>
              </presentation:event-listener>
            </xsl:when>
            <!-- End => Go to Last slide -->
            <!-- Start => EndShow -->
            <xsl:when test="a:hlinkClick/@action[contains(.,'jump=endshow')]">
              <presentation:event-listener>
                <xsl:attribute name ="script:event-name">
                  <xsl:value-of select ="'dom:click'"/>
                </xsl:attribute>
                <xsl:attribute name ="presentation:action">
                  <xsl:value-of select ="'stop'"/>
                </xsl:attribute>
              </presentation:event-listener>
            </xsl:when>
            <!-- End => End show -->
            <!-- Start => Go to Page or Object && Go to Other document && Run program  -->
            <xsl:when test="a:hlinkClick/@action[contains(.,'ppaction://hlinksldjump') or contains(.,'ppaction://hlinkfile') or contains(.,'ppaction://program') ]">
              <presentation:event-listener>
                <xsl:attribute name ="script:event-name">
                  <xsl:value-of select ="'dom:click'"/>
                </xsl:attribute>
                <xsl:variable name="RelationId">
                  <xsl:value-of select="a:hlinkClick/@r:id"/>
                </xsl:variable>
                <xsl:variable name="SlideVal">
                  <xsl:value-of select="document($SlideRelationId)/rels:Relationships/rels:Relationship[@Id=$RelationId]/@Target"/>
                </xsl:variable>
                <!-- Condn Go to Other page/slide-->
                <xsl:if test="a:hlinkClick/@action[contains(.,'ppaction://hlinksldjump')]">
                  <xsl:attribute name ="presentation:action">
                    <xsl:value-of select ="'show'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="xlink:href">
                    <xsl:value-of select ="concat('#page',substring-before(substring-after($SlideVal,'slide'),'.xml'))"/>
                  </xsl:attribute>
                </xsl:if>
                <!-- Condn Go to Other document-->
                <xsl:if test="a:hlinkClick/@action[contains(.,'ppaction://hlinkfile')]">
                  <xsl:attribute name ="presentation:action">
                    <xsl:value-of select ="'show'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="xlink:href">
                    <xsl:if test="string-length(substring-after($SlideVal,'file:///')) > 0">
                      <xsl:value-of select ="concat('/',translate(substring-after($SlideVal,'file:///'),'\','/'))"/>
                    </xsl:if>
                    <xsl:if test="string-length(substring-after($SlideVal,'file:///')) = 0 ">
                      <xsl:value-of select ="concat('../',$SlideVal)"/>
                    </xsl:if>
                  </xsl:attribute>
                </xsl:if>
                <!-- Condn Go to Run program-->
                <xsl:if test="a:hlinkClick/@action[contains(.,'ppaction://program') ]">
                  <xsl:attribute name ="presentation:action">
                    <xsl:value-of select ="'execute'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="xlink:href">
                    <xsl:if test="string-length(substring-after($SlideVal,'file:///')) > 0">
                      <xsl:value-of select ="concat('/',translate(substring-after($SlideVal,'file:///'),'\','/'))"/>
                    </xsl:if>
                    <xsl:if test="string-length(substring-after($SlideVal,'file:///')) = 0">
                      <xsl:value-of select ="concat('../',$SlideVal)"/>
                    </xsl:if>
                  </xsl:attribute>
                </xsl:if>
                <xsl:attribute name ="xlink:type">
                  <xsl:value-of select ="'simple'"/>
                </xsl:attribute>
                <xsl:attribute name ="xlink:show">
                  <xsl:value-of select ="'new'"/>
                </xsl:attribute>
                <xsl:attribute name ="xlink:actuate">
                  <xsl:value-of select ="'onRequest'"/>
                </xsl:attribute>
              </presentation:event-listener>
            </xsl:when>
            <!-- End => Go to Page or Object && Go to Other document && Run program -->
            <!-- Start => Play sound  -->
            <xsl:when test="a:hlinkClick/a:snd">
              <presentation:event-listener>
                <xsl:attribute name ="script:event-name">
                  <xsl:value-of select ="'dom:click'"/>
                </xsl:attribute>
                <xsl:attribute name ="presentation:action">
                  <xsl:value-of select ="'sound'"/>
                </xsl:attribute>
                <presentation:sound>
                  <xsl:variable name="varMediaFileRelId">
                    <xsl:value-of select="a:hlinkClick/a:snd/@r:embed"/>
                  </xsl:variable>
                  <xsl:variable name="varMediaFileTargetPath">
                    <xsl:value-of select="document($SlideRelationId)/rels:Relationships/rels:Relationship[@Id=$varMediaFileRelId]/@Target"/>
                  </xsl:variable>
                  <xsl:variable name="varPptMediaFileTargetPath">
                    <xsl:value-of select="concat('ppt/',substring-after($varMediaFileTargetPath,'/'))"/>
                  </xsl:variable>
                  <!--<xsl:variable name="varDocumentModifiedTime">
                    <xsl:value-of select="document('docProps/core.xml')/cp:coreProperties/dcterms:modified"/>
                  </xsl:variable>-->
                  <!--<xsl:variable name="varDestMediaFileTargetPath">
                    <xsl:value-of select="concat(translate($varDocumentModifiedTime,':','-'),'|',$varMediaFileRelId,'.wav')"/>
                  </xsl:variable>
                  <xsl:variable name="varMediaFilePathForOdp">
                    <xsl:value-of select="concat('../_MediaFilesForOdp_',translate($varDocumentModifiedTime,':','-'),'/',$varMediaFileRelId,'.wav')"/>
                  </xsl:variable>-->

                  <xsl:variable name="FolderNameGUID">
                    <xsl:call-template name="GenerateGUIDForFolderName">
                      <xsl:with-param name="RootNode" select="." />
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:variable name="varDestMediaFileTargetPath">
                    <xsl:value-of select="concat($FolderNameGUID,'|',$varMediaFileRelId,'.wav')"/>
                  </xsl:variable>
                  <xsl:variable name="varMediaFilePathForOdp">
                    <xsl:value-of select="concat('../',$FolderNameGUID,'/',$varMediaFileRelId,'.wav')"/>
                  </xsl:variable>
                  <xsl:attribute name ="xlink:href">
                    <xsl:value-of select ="$varMediaFilePathForOdp"/>
                  </xsl:attribute>
                  <xsl:attribute name ="xlink:type">
                    <xsl:value-of select ="'simple'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="xlink:show">
                    <xsl:value-of select ="'new'"/>
                  </xsl:attribute>
                  <xsl:attribute name ="xlink:actuate">
                    <xsl:value-of select ="'onRequest'"/>
                  </xsl:attribute>
                  <pzip:extract pzip:source="{$varPptMediaFileTargetPath}" pzip:target="{$varDestMediaFileTargetPath}" />
                </presentation:sound>
              </presentation:event-listener>
            </xsl:when>

            <!-- End => Play sound  -->
          </xsl:choose>
        </xsl:for-each>
      </office:event-listeners>
      <!-- End - Mouse click hyperlinks-->
    </xsl:variable>

    <xsl:choose>

      <!-- Basic shapes start-->

      <!--Custom shape - Rectangle -->
		  <xsl:when test ="p:spPr/a:prstGeom/@prst='rect'">
      <xsl:choose>
        <xsl:when test="p:style">
          <draw:custom-shape draw:layer="layout" >
            <xsl:call-template name ="CreateShape">
              <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
              <xsl:with-param name="sldId" select="$slideId" />
              <xsl:with-param name="grID" select ="$GraphicId"/>
              <xsl:with-param name ="prID" select="$ParaId" />
              <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
              <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
              <xsl:with-param name="TypeId" select ="$TypeId" />
              <xsl:with-param name="grpBln" select ="$grpBln" />
              <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
              <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            </xsl:call-template>
            <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
                       draw:type="rectangle" 
                       draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N">
              <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
                <xsl:attribute name ="draw:mirror-horizontal">
                  <xsl:value-of select="'true'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
                <xsl:attribute name ="draw:mirror-vertical">
                  <xsl:value-of select="'true'"/>
                </xsl:attribute>
              </xsl:if>
            </draw:enhanced-geometry>
            <xsl:copy-of select="$varHyperLinksForShapes" />
          </draw:custom-shape>
        </xsl:when>
        <xsl:otherwise>
          <draw:frame draw:layer="layout">

            <xsl:call-template name ="CreateShape">

              <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
              <xsl:with-param name="flagTextBox" select="'true'" />
              <xsl:with-param name="sldId" select="$slideId" />
              <xsl:with-param name="grID" select ="$GraphicId"/>
              <xsl:with-param name ="prID" select="$ParaId" />
              <xsl:with-param name="TypeId" select ="$TypeId" />
              <xsl:with-param name="grpBln" select ="$grpBln" />
              <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
              <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
              <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
              <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            </xsl:call-template>
            <xsl:copy-of select="$varHyperLinksForShapes" />
          </draw:frame>
        </xsl:otherwise>
      </xsl:choose>
      
      </xsl:when>		
      <!-- Oval(Custom shape) -->
      <xsl:when test ="(p:nvSpPr/p:cNvPr/@name[contains(., 'Oval')]) and (p:spPr/a:prstGeom/@prst='ellipse')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
		     	<xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
          <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
											draw:text-areas="3200 3200 18400 18400" 
											draw:type="ellipse"
											draw:glue-points="10800 0 3160 3160 0 10800 3160 18440 10800 21600 18440 18440 21600 10800 18440 3160">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
	  </xsl:when>

      <!--Added by Mathi for bug Fix on 17th Sep 2007-->
      <xsl:when test ="(not(p:nvSpPr/p:cNvPr/@name[contains(., 'Ellipse ')]) or (p:nvSpPr/p:cNvPr/@name[contains(., 'Ellipse ')])) and (p:spPr/a:prstGeom/@prst='ellipse')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
											draw:text-areas="3200 3200 18400 18400" 
											draw:type="ellipse"
											draw:glue-points="10800 0 3160 3160 0 10800 3160 18440 10800 21600 18440 18440 21600 10800 18440 3160">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!--End of Code-->
      
      <!--Right Arrow (Added by A.Mathi as on 2/07/2007) -->
      <xsl:when test = "(p:spPr/a:prstGeom/@prst='rightArrow')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
			<xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
					  draw:text-areas="0 ?f0 ?f5 ?f2" 
					  draw:type="right-arrow" draw:modifiers="16200 5400" 
					  draw:enhanced-path="M 0 ?f0 L ?f1 ?f0 ?f1 0 21600 10800 ?f1 21600 ?f1 ?f2 0 ?f2 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="$1 "/>
            <draw:equation draw:name="f1" draw:formula="$0 "/>
            <draw:equation draw:name="f2" draw:formula="21600-$1 "/>
            <draw:equation draw:name="f3" draw:formula="21600-?f1 "/>
            <draw:equation draw:name="f4" draw:formula="?f3 *?f0 /10800"/>
            <draw:equation draw:name="f5" draw:formula="?f1 +?f4 "/>
            <draw:equation draw:name="f6" draw:formula="?f1 *?f0 /10800"/>
            <draw:equation draw:name="f7" draw:formula="?f1 -?f6 "/>
            <draw:handle draw:handle-position="$0 $1" 
						  draw:handle-range-x-minimum="0" 
						  draw:handle-range-x-maximum="21600" 
						  draw:handle-range-y-minimum="0" 
						  draw:handle-range-y-maximum="10800"/>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!--Up Arrow (Added by A.Mathi as on 2/07/2007) -->
      <xsl:when test = "(p:spPr/a:prstGeom/@prst='upArrow')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
			<xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name ="grID" select="$GraphicId" />
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name ="SlideRelationId" select="$SlideRelationId" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
					  draw:text-areas="?f0 ?f7 ?f2 21600" 
					  draw:type="up-arrow" draw:modifiers="5400 5400" 
					  draw:enhanced-path="M ?f0 21600 L ?f0 ?f1 0 ?f1 10800 0 21600 ?f1 ?f2 ?f1 ?f2 21600 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="$1 "/>
            <draw:equation draw:name="f1" draw:formula="$0 "/>
            <draw:equation draw:name="f2" draw:formula="21600-$1 "/>
            <draw:equation draw:name="f3" draw:formula="21600-?f1 "/>
            <draw:equation draw:name="f4" draw:formula="?f3 *?f0 /10800"/>
            <draw:equation draw:name="f5" draw:formula="?f1 +?f4 "/>
            <draw:equation draw:name="f6" draw:formula="?f1 *?f0 /10800"/>
            <draw:equation draw:name="f7" draw:formula="?f1 -?f6 "/>
            <draw:handle draw:handle-position="$1 $0" 
						  draw:handle-range-x-minimum="0" 
						  draw:handle-range-x-maximum="10800" 
						  draw:handle-range-y-minimum="0" 
						  draw:handle-range-y-maximum="21600"/>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes"/>
        </draw:custom-shape>
      </xsl:when>
      <!--Left Arrow (Added by A.Mathi as on 2/07/2007) -->
      <xsl:when test = "(p:spPr/a:prstGeom/@prst='leftArrow')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
			<xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name ="grID" select="$GraphicId" />
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name ="SlideRelationId" select="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
					  draw:text-areas="?f7 ?f0 21600 ?f2" 
					  draw:type="left-arrow" 
					  draw:modifiers="5400 5400" 
					  draw:enhanced-path="M 21600 ?f0 L ?f1 ?f0 ?f1 0 0 10800 ?f1 21600 ?f1 ?f2 21600 ?f2 Z N">
             <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="$1 "/>
            <draw:equation draw:name="f1" draw:formula="$0 "/>
            <draw:equation draw:name="f2" draw:formula="21600-$1 "/>
            <draw:equation draw:name="f3" draw:formula="21600-?f1 "/>
            <draw:equation draw:name="f4" draw:formula="?f3 *?f0 /10800"/>
            <draw:equation draw:name="f5" draw:formula="?f1 +?f4 "/>
            <draw:equation draw:name="f6" draw:formula="?f1 *?f0 /10800"/>
            <draw:equation draw:name="f7" draw:formula="?f1 -?f6 "/>
            <draw:handle draw:handle-position="$0 $1" 
						  draw:handle-range-x-minimum="0" 
						  draw:handle-range-x-maximum="21600" 
						  draw:handle-range-y-minimum="0" 
						  draw:handle-range-y-maximum="10800"/>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes"/>
        </draw:custom-shape>
      </xsl:when>
      <!--Down Arrow (Added by A.Mathi as on 2/07/2007) -->
      <xsl:when test = "(p:spPr/a:prstGeom/@prst='downArrow')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
			<xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name ="grID" select ="$GraphicId" />
            <xsl:with-param name ="prID" select ="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name ="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
					  draw:text-areas="?f0 0 ?f2 ?f5" 
					  draw:type="down-arrow" 
					  draw:modifiers="16200 5400" 
					  draw:enhanced-path="M ?f0 0 L ?f0 ?f1 0 ?f1 10800 21600 21600 ?f1 ?f2 ?f1 ?f2 0 Z N">
              <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="$1 "/>
            <draw:equation draw:name="f1" draw:formula="$0 "/>
            <draw:equation draw:name="f2" draw:formula="21600-$1 "/>
            <draw:equation draw:name="f3" draw:formula="21600-?f1 "/>
            <draw:equation draw:name="f4" draw:formula="?f3 *?f0 /10800"/>
            <draw:equation draw:name="f5" draw:formula="?f1 +?f4 "/>
            <draw:equation draw:name="f6" draw:formula="?f1 *?f0 /10800"/>
            <draw:equation draw:name="f7" draw:formula="?f1 -?f6 "/>
            <draw:handle draw:handle-position="$1 $0" 
						  draw:handle-range-x-minimum="0" 
						  draw:handle-range-x-maximum="10800" 
						  draw:handle-range-y-minimum="0" 
						  draw:handle-range-y-maximum="21600"/>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes"/>
        </draw:custom-shape>
      </xsl:when>
      <!--LeftRight Arrow (Added by A.Mathi as on 3/07/2007) -->
      <xsl:when test = "(p:spPr/a:prstGeom/@prst='leftRightArrow')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
			  <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name ="grID" select ="$GraphicId" />
            <xsl:with-param name ="prID" select ="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name ="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
					  draw:text-areas="?f5 ?f1 ?f6 ?f3" 
					  draw:type="left-right-arrow" 
					  draw:modifiers="4300 5400" 
					  draw:enhanced-path="M 0 10800 L ?f0 0 ?f0 ?f1 ?f2 ?f1 ?f2 0 21600 10800 ?f2 21600 ?f2 ?f3 ?f0 ?f3 ?f0 21600 Z N">
              <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="$0 "/>
            <draw:equation draw:name="f1" draw:formula="$1 "/>
            <draw:equation draw:name="f2" draw:formula="21600-$0 "/>
            <draw:equation draw:name="f3" draw:formula="21600-$1 "/>
            <draw:equation draw:name="f4" draw:formula="10800-$1 "/>
            <draw:equation draw:name="f5" draw:formula="$0 *?f4 /10800"/>
            <draw:equation draw:name="f6" draw:formula="21600-?f5 "/>
            <draw:equation draw:name="f7" draw:formula="10800-$0 "/>
            <draw:equation draw:name="f8" draw:formula="$1 *?f7 /10800"/>
            <draw:equation draw:name="f9" draw:formula="21600-?f8 "/>
            <draw:handle draw:handle-position="$0 $1" 
						  draw:handle-range-x-minimum="0" 
						  draw:handle-range-x-maximum="10800" 
						  draw:handle-range-y-minimum="0" 
						  draw:handle-range-y-maximum="10800"/>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes"/>
        </draw:custom-shape>
      </xsl:when>
      <!-- UpDown Arrow (Added by A.Mathi as on 4/07/2007) -->
      <xsl:when test = "(p:spPr/a:prstGeom/@prst='upDownArrow')" >
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
			<xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name ="grID" select ="$GraphicId" />
            <xsl:with-param name ="prID" select ="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name ="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
					  draw:text-areas="?f0 ?f8 ?f2 ?f9" 
					  draw:type="up-down-arrow" 
					  draw:modifiers="5400 4300" 
					  draw:enhanced-path="M 0 ?f1 L 10800 0 21600 ?f1 ?f2 ?f1 ?f2 ?f3 21600 ?f3 10800 21600 0 ?f3 ?f0 ?f3 ?f0 ?f1 Z N">
              <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="$0 "/>
            <draw:equation draw:name="f1" draw:formula="$1 "/>
            <draw:equation draw:name="f2" draw:formula="21600-$0 "/>
            <draw:equation draw:name="f3" draw:formula="21600-$1 "/>
            <draw:equation draw:name="f4" draw:formula="10800-$1 "/>
            <draw:equation draw:name="f5" draw:formula="$0 *?f4 /10800"/>
            <draw:equation draw:name="f6" draw:formula="21600-?f5 "/>
            <draw:equation draw:name="f7" draw:formula="10800-$0 "/>
            <draw:equation draw:name="f8" draw:formula="$1 *?f7 /10800"/>
            <draw:equation draw:name="f9" draw:formula="21600-?f8 "/>
            <draw:handle draw:handle-position="$0 $1" 
						  draw:handle-range-x-minimum="0" 
						  draw:handle-range-x-maximum="10800" 
						  draw:handle-range-y-minimum="0" 
						  draw:handle-range-y-maximum="10800"/>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes"/>
        </draw:custom-shape>
      </xsl:when>
      <!-- Isosceles Triangle -->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='triangle')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
						draw:glue-points="10800 0 ?f1 10800 0 21600 10800 21600 21600 21600 ?f7 10800" 
						draw:text-areas="?f1 10800 ?f2 18000 ?f3 7200 ?f4 21600" 
						draw:type="isosceles-triangle" draw:modifiers="10800" 
						draw:enhanced-path="M ?f0 0 L 21600 21600 0 21600 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="$0 "/>
            <draw:equation draw:name="f1" draw:formula="$0 /2"/>
            <draw:equation draw:name="f2" draw:formula="?f1 +10800"/>
            <draw:equation draw:name="f3" draw:formula="$0 *2/3"/>
            <draw:equation draw:name="f4" draw:formula="?f3 +7200"/>
            <draw:equation draw:name="f5" draw:formula="21600-?f0 "/>
            <draw:equation draw:name="f6" draw:formula="?f5 /2"/>
            <draw:equation draw:name="f7" draw:formula="21600-?f6 "/>
            <draw:handle draw:handle-position="$0 top" 
									 draw:handle-range-x-minimum="0" 
									 draw:handle-range-x-maximum="21600"/>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- Right Triangle -->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='rtTriangle')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
											draw:glue-points="10800 0 5400 10800 0 21600 10800 21600 21600 21600 16200 10800" 
											draw:text-areas="1900 12700 12700 19700" 
											draw:type="right-triangle" 
											draw:enhanced-path="M 0 0 L 21600 21600 0 21600 0 0 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- Parallelogram -->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='parallelogram')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
						draw:glue-points="?f6 0 10800 ?f8 ?f11 10800 ?f9 21600 10800 ?f10 ?f5 10800" 
						draw:text-areas="?f3 ?f3 ?f4 ?f4" draw:type="parallelogram" 
						draw:modifiers="5400" 
            draw:enhanced-path="M ?f0 0 L 21600 0 ?f1 21600 0 21600 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="$0 "/>
            <draw:equation draw:name="f1" draw:formula="21600-$0 "/>
            <draw:equation draw:name="f2" draw:formula="$0 *10/24"/>
            <draw:equation draw:name="f3" draw:formula="?f2 +1750"/>
            <draw:equation draw:name="f4" draw:formula="21600-?f3 "/>
            <draw:equation draw:name="f5" draw:formula="?f0 /2"/>
            <draw:equation draw:name="f6" draw:formula="10800+?f5 "/>
            <draw:equation draw:name="f7" draw:formula="?f0 -10800"/>
            <draw:equation draw:name="f8" draw:formula="if(?f7 ,?f13 ,0)"/>
            <draw:equation draw:name="f9" draw:formula="10800-?f5 "/>
            <draw:equation draw:name="f10" draw:formula="if(?f7 ,?f12 ,21600)"/>
            <draw:equation draw:name="f11" draw:formula="21600-?f5 "/>
            <draw:equation draw:name="f12" draw:formula="21600*10800/?f0 "/>
            <draw:equation draw:name="f13" draw:formula="21600-?f12 "/>
            <draw:handle draw:handle-position="$0 top" 
									 draw:handle-range-x-minimum="0" 
									 draw:handle-range-x-maximum="21600"/>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- Trapezoid (Added by A.Mathi as on 24/07/2007) -->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='trapezoid')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
			<xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 914400 1216152" 
          draw:extrusion-allowed="true" 
          draw:text-areas="152400 202692 762000 1216152" 
          draw:glue-points="457200 0 114300 608076 457200 1216152 800100 608076" 
          draw:type="mso-spt100" 
          draw:enhanced-path="M 0 1216152 L 228600 0 L 685800 0 L 914400 1216152 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- Diamond   -->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='diamond')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
											draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" 
											draw:text-areas="5400 5400 16200 16200" 
											draw:type="diamond" 
											draw:enhanced-path="M 10800 0 L 21600 10800 10800 21600 0 10800 10800 0 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>

      </xsl:when>
      <!-- Regular Pentagon -->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='pentagon')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
											draw:glue-points="10800 0 0 8260 4230 21600 10800 21600 17370 21600 21600 8260" 
											draw:text-areas="4230 5080 17370 21600" 
											draw:type="pentagon" 
											draw:enhanced-path="M 10800 0 L 0 8260 4230 21600 17370 21600 21600 8260 10800 0 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- Hexagon -->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='hexagon')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
											draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" 
											draw:text-areas="?f3 ?f3 ?f4 ?f4" draw:type="hexagon" 
											draw:modifiers="5400" 
											draw:enhanced-path="M ?f0 0 L ?f1 0 21600 10800 ?f1 21600 ?f0 21600 0 10800 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="$0 "/>
            <draw:equation draw:name="f1" draw:formula="21600-$0 "/>
            <draw:equation draw:name="f2" draw:formula="$0 *100/234"/>
            <draw:equation draw:name="f3" draw:formula="?f2 +1700"/>
            <draw:equation draw:name="f4" draw:formula="21600-?f3 "/>
            <draw:handle draw:handle-position="$0 top" 
                   draw:handle-range-x-minimum="0" 
                   draw:handle-range-x-maximum="10800"/>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- Octagon -->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='octagon')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
						draw:glue-points="10800 0 0 10800 10800 21600 21600 10800"
						draw:path-stretchpoint-x="10800" draw:path-stretchpoint-y="10800"
						draw:text-areas="?f5 ?f6 ?f7 ?f8" draw:type="octagon" draw:modifiers="5000"
						draw:enhanced-path="M ?f0 0 L ?f2 0 21600 ?f1 21600 ?f3 ?f2 21600 ?f0 21600 0 ?f3 0 ?f1 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="left+$0 "/>
            <draw:equation draw:name="f1" draw:formula="top+$0 "/>
            <draw:equation draw:name="f2" draw:formula="right-$0 "/>
            <draw:equation draw:name="f3" draw:formula="bottom-$0 "/>
            <draw:equation draw:name="f4" draw:formula="$0 /2"/>
            <draw:equation draw:name="f5" draw:formula="left+?f4 "/>
            <draw:equation draw:name="f6" draw:formula="top+?f4 "/>
            <draw:equation draw:name="f7" draw:formula="right-?f4 "/>
            <draw:equation draw:name="f8" draw:formula="bottom-?f4 "/>
            <draw:handle draw:handle-position="$0 top"
									 draw:handle-range-x-minimum="0"
									 draw:handle-range-x-maximum="10800"/>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>

      <!-- Circular Arrow or CurvedLeftArrow or curvedRightArrow or  CurvedDownArrow or CurvedUpArrow -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='circularArrow' or p:spPr/a:prstGeom/@prst='curvedRightArrow' or p:spPr/a:prstGeom/@prst='curvedLeftArrow' or p:spPr/a:prstGeom/@prst='curvedDownArrow' or p:spPr/a:prstGeom/@prst='curvedUpArrow'">
        <!-- warn if CurvedLeftArrow or curvedRightArrow or  CurvedDownArrow or CurvedUpArrow   -->
        <xsl:message terminate="no">translation.oox2odf.shapesTypeCurvedLeftRightUpDownArrow</xsl:message>
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <xsl:if test="p:spPr/a:prstGeom/@prst='circularArrow' or p:spPr/a:prstGeom/@prst='curvedDownArrow' or p:spPr/a:prstGeom/@prst='curvedUpArrow'">
            <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:text-areas="0 0 21600 21600" 
                                  draw:modifiers="180 0 5500" draw:enhanced-path="B ?f3 ?f3 ?f20 ?f20 ?f19 ?f18 ?f17 ?f16 W 0 0 21600 21600 ?f9 ?f8 ?f11 ?f10 L ?f24 ?f23 ?f36 ?f35 ?f29 ?f28 Z N">
              <xsl:if test="p:spPr/a:prstGeom/@prst='curvedUpArrow'">
                <xsl:attribute name ="draw:type">
                  <xsl:value-of select="'curvedUpArrow'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="p:spPr/a:prstGeom/@prst='curvedDownArrow'">
                <xsl:attribute name ="draw:type">
                  <xsl:value-of select="'curvedDownArrow'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="p:spPr/a:prstGeom/@prst='circularArrow'">
                <xsl:attribute name ="draw:type">
                  <xsl:value-of select="'circular-arrow'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="p:spPr/a:prstGeom/@prst='curvedUpArrow'">
                <xsl:attribute name ="draw:mirror-vertical">
                  <xsl:value-of select="'true'"/>
                </xsl:attribute>
              </xsl:if>
              <draw:equation draw:name="f0" draw:formula="$0 "/>
              <draw:equation draw:name="f1" draw:formula="$1 "/>
              <draw:equation draw:name="f2" draw:formula="$2 "/>
              <draw:equation draw:name="f3" draw:formula="10800+$2 "/>
              <draw:equation draw:name="f4" draw:formula="10800*sin($0 *(pi/180))"/>
              <draw:equation draw:name="f5" draw:formula="10800*cos($0 *(pi/180))"/>
              <draw:equation draw:name="f6" draw:formula="10800*sin($1 *(pi/180))"/>
              <draw:equation draw:name="f7" draw:formula="10800*cos($1 *(pi/180))"/>
              <draw:equation draw:name="f8" draw:formula="?f4 +10800"/>
              <draw:equation draw:name="f9" draw:formula="?f5 +10800"/>
              <draw:equation draw:name="f10" draw:formula="?f6 +10800"/>
              <draw:equation draw:name="f11" draw:formula="?f7 +10800"/>
              <draw:equation draw:name="f12" draw:formula="?f3 *sin($0 *(pi/180))"/>
              <draw:equation draw:name="f13" draw:formula="?f3 *cos($0 *(pi/180))"/>
              <draw:equation draw:name="f14" draw:formula="?f3 *sin($1 *(pi/180))"/>
              <draw:equation draw:name="f15" draw:formula="?f3 *cos($1 *(pi/180))"/>
              <draw:equation draw:name="f16" draw:formula="?f12 +10800"/>
              <draw:equation draw:name="f17" draw:formula="?f13 +10800"/>
              <draw:equation draw:name="f18" draw:formula="?f14 +10800"/>
              <draw:equation draw:name="f19" draw:formula="?f15 +10800"/>
              <draw:equation draw:name="f20" draw:formula="21600-?f3 "/>
              <draw:equation draw:name="f21" draw:formula="13500*sin($1 *(pi/180))"/>
              <draw:equation draw:name="f22" draw:formula="13500*cos($1 *(pi/180))"/>
              <draw:equation draw:name="f23" draw:formula="?f21 +10800"/>
              <draw:equation draw:name="f24" draw:formula="?f22 +10800"/>
              <draw:equation draw:name="f25" draw:formula="$2 -2700"/>
              <draw:equation draw:name="f26" draw:formula="?f25 *sin($1 *(pi/180))"/>
              <draw:equation draw:name="f27" draw:formula="?f25 *cos($1 *(pi/180))"/>
              <draw:equation draw:name="f28" draw:formula="?f26 +10800"/>
              <draw:equation draw:name="f29" draw:formula="?f27 +10800"/>
              <draw:equation draw:name="f30" draw:formula="($1+45)*pi/180"/>
              <draw:equation draw:name="f31" draw:formula="sqrt(((?f29-?f24)*(?f29-?f24))+((?f28-?f23)*(?f28-?f23)))"/>
              <draw:equation draw:name="f32" draw:formula="sqrt(2)/2*?f31"/>
              <draw:equation draw:name="f33" draw:formula="?f32*sin(?f30)"/>
              <draw:equation draw:name="f34" draw:formula="?f32*cos(?f30)"/>
              <draw:equation draw:name="f35" draw:formula="?f28+?f33"/>
              <draw:equation draw:name="f36" draw:formula="?f29+?f34"/>
              <draw:handle draw:handle-position="10800 $0" draw:handle-polar="10800 10800" draw:handle-radius-range-minimum="10800" draw:handle-radius-range-maximum="10800"/>
              <draw:handle draw:handle-position="$2 $1" draw:handle-polar="10800 10800" draw:handle-radius-range-minimum="0" draw:handle-radius-range-maximum="10800"/>
            </draw:enhanced-geometry>
          </xsl:if>
          <xsl:if test="p:spPr/a:prstGeom/@prst='curvedRightArrow' or p:spPr/a:prstGeom/@prst='curvedLeftArrow'">
            <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:text-areas="0 0 21600 21600"                                                                                  
                                      draw:mirror-vertical="true" draw:type="circular-arrow" draw:modifiers="84.1499444519454 -80.8262021942408 7331.32459905244" draw:enhanced-path="B ?f3 ?f3 ?f20 ?f20 ?f19 ?f18 ?f17 ?f16 W 0 0 21600 21600 ?f9 ?f8 ?f11 ?f10 L ?f24 ?f23 ?f36 ?f35 ?f29 ?f28 Z N">
              <xsl:if test="p:spPr/a:prstGeom/@prst='curvedRightArrow'">
                <xsl:attribute name ="draw:type">
                  <xsl:value-of select="'curvedRightArrow'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="p:spPr/a:prstGeom/@prst='curvedLeftArrow'">
                <xsl:attribute name ="draw:type">
                  <xsl:value-of select="'curvedLeftArrow'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="p:spPr/a:prstGeom/@prst='curvedRightArrow'">
                <xsl:attribute name ="draw:mirror-horizontal">
                  <xsl:value-of select="'false'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="p:spPr/a:prstGeom/@prst='curvedLeftArrow'">
                <xsl:attribute name ="draw:mirror-horizontal">
                  <xsl:value-of select="'true'"/>
                </xsl:attribute>
              </xsl:if>
              <draw:equation draw:name="f0" draw:formula="$0 "/>
              <draw:equation draw:name="f1" draw:formula="$1 "/>
              <draw:equation draw:name="f2" draw:formula="$2 "/>
              <draw:equation draw:name="f3" draw:formula="10800+$2 "/>
              <draw:equation draw:name="f4" draw:formula="10800*sin($0 *(pi/180))"/>
              <draw:equation draw:name="f5" draw:formula="10800*cos($0 *(pi/180))"/>
              <draw:equation draw:name="f6" draw:formula="10800*sin($1 *(pi/180))"/>
              <draw:equation draw:name="f7" draw:formula="10800*cos($1 *(pi/180))"/>
              <draw:equation draw:name="f8" draw:formula="?f4 +10800"/>
              <draw:equation draw:name="f9" draw:formula="?f5 +10800"/>
              <draw:equation draw:name="f10" draw:formula="?f6 +10800"/>
              <draw:equation draw:name="f11" draw:formula="?f7 +10800"/>
              <draw:equation draw:name="f12" draw:formula="?f3 *sin($0 *(pi/180))"/>
              <draw:equation draw:name="f13" draw:formula="?f3 *cos($0 *(pi/180))"/>
              <draw:equation draw:name="f14" draw:formula="?f3 *sin($1 *(pi/180))"/>
              <draw:equation draw:name="f15" draw:formula="?f3 *cos($1 *(pi/180))"/>
              <draw:equation draw:name="f16" draw:formula="?f12 +10800"/>
              <draw:equation draw:name="f17" draw:formula="?f13 +10800"/>
              <draw:equation draw:name="f18" draw:formula="?f14 +10800"/>
              <draw:equation draw:name="f19" draw:formula="?f15 +10800"/>
              <draw:equation draw:name="f20" draw:formula="21600-?f3 "/>
              <draw:equation draw:name="f21" draw:formula="13500*sin($1 *(pi/180))"/>
              <draw:equation draw:name="f22" draw:formula="13500*cos($1 *(pi/180))"/>
              <draw:equation draw:name="f23" draw:formula="?f21 +10800"/>
              <draw:equation draw:name="f24" draw:formula="?f22 +10800"/>
              <draw:equation draw:name="f25" draw:formula="$2 -2700"/>
              <draw:equation draw:name="f26" draw:formula="?f25 *sin($1 *(pi/180))"/>
              <draw:equation draw:name="f27" draw:formula="?f25 *cos($1 *(pi/180))"/>
              <draw:equation draw:name="f28" draw:formula="?f26 +10800"/>
              <draw:equation draw:name="f29" draw:formula="?f27 +10800"/>
              <draw:equation draw:name="f30" draw:formula="($1+45)*pi/180"/>
              <draw:equation draw:name="f31" draw:formula="sqrt(((?f29-?f24)*(?f29-?f24))+((?f28-?f23)*(?f28-?f23)))"/>
              <draw:equation draw:name="f32" draw:formula="sqrt(2)/2*?f31"/>
              <draw:equation draw:name="f33" draw:formula="?f32*sin(?f30)"/>
              <draw:equation draw:name="f34" draw:formula="?f32*cos(?f30)"/>
              <draw:equation draw:name="f35" draw:formula="?f28+?f33"/>
              <draw:equation draw:name="f36" draw:formula="?f29+?f34"/>
              <draw:handle draw:handle-position="10800 $0" draw:handle-polar="10800 10800" draw:handle-radius-range-minimum="10800" draw:handle-radius-range-maximum="10800"/>
              <draw:handle draw:handle-position="$2 $1" draw:handle-polar="10800 10800" draw:handle-radius-range-minimum="0" draw:handle-radius-range-maximum="10800"/>
            </draw:enhanced-geometry>
          </xsl:if>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      
      <!-- Left-Up Arrow -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='leftUpArrow'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
			<xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:mirror-horizontal="false" 
            draw:text-areas="?f2 ?f7 ?f1 ?f1 ?f7 ?f2 ?f1 ?f1" draw:type="mso-spt89" 
            draw:modifiers="10062 18378 6098" 
            draw:enhanced-path="M 0 ?f5 L ?f2 ?f0 ?f2 ?f7 ?f7 ?f7 ?f7 ?f2 ?f0 ?f2 ?f5 0 21600 ?f2 ?f1 ?f2 ?f1 ?f1 ?f2 ?f1 ?f2 21600 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="$0 "/>
            <draw:equation draw:name="f1" draw:formula="$1 "/>
            <draw:equation draw:name="f2" draw:formula="$2 "/>
            <draw:equation draw:name="f3" draw:formula="21600-$0 "/>
            <draw:equation draw:name="f4" draw:formula="?f3 /2"/>
            <draw:equation draw:name="f5" draw:formula="$0 +?f4 "/>
            <draw:equation draw:name="f6" draw:formula="21600-$1 "/>
            <draw:equation draw:name="f7" draw:formula="$0 +?f6 "/>
            <draw:equation draw:name="f8" draw:formula="21600-?f6 "/>
            <draw:equation draw:name="f9" draw:formula="?f8 -?f6 "/>
            <draw:handle draw:handle-position="$1 $2" draw:handle-range-x-minimum="?f5" draw:handle-range-x-maximum="21600" draw:handle-range-y-minimum="0" draw:handle-range-y-maximum="$0"/>
            <draw:handle draw:handle-position="$0 top" draw:handle-range-x-minimum="$2" draw:handle-range-x-maximum="?f9"/>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      
      <!-- Bent-Up Arrow -->
      <!--Bug Fix for Shape Corner-Right Arrow from ODP to PPtx-->
      <xsl:when test="(p:nvSpPr/p:cNvPr/@name[contains(., 'bentUpArrow ')]) and not(p:nvSpPr/p:cNvPr/@name[contains(., 'Bent-Up Arrow ')]) and (p:spPr/a:prstGeom/@prst='bentUpArrow')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name ="grID" select ="$GraphicId" />
            <xsl:with-param name ="prID" select ="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!-- End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering -->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 841 854" 
            draw:mirror-horizontal="false" 
            draw:mirror-vertical="false" 
            draw:type="non-primitive" 
            draw:enhanced-path="M 517 247 L 517 415 264 415 264 0 0 0 0 680 517 680 517 854 841 547 517 247 Z N">
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!--End of bug fix code-->
      
      <xsl:when test ="p:spPr/a:prstGeom/@prst='bentUpArrow'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
			 <xsl:with-param name="sldId" select="$slideId" /> 
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
			<xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 3276600 1905000" draw:extrusion-allowed="true" 
              draw:text-areas="0 1428750 3038475 1905000" 
              draw:glue-points="2800350 0 2324100 476250 0 1666874 1519238 1905000 3038475 1190624 3276600 476250" 
              draw:type="mso-spt100" 
              draw:enhanced-path="M 0 1428750 L 2562225 1428750 L 2562225 476250 L 2324100 476250 L 2800350 0 L 3276600 476250 L 3038475 476250 L 3038475 1905000 L 0 1905000 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      
      <!-- Cube -->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='cube')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
						draw:glue-points="?f7 0 ?f6 ?f1 0 ?f10 ?f6 21600 ?f4 ?f10 21600 ?f9" 
						draw:path-stretchpoint-x="10800" draw:path-stretchpoint-y="10800" 
						draw:text-areas="0 ?f1 ?f4 ?f12" draw:type="cube" draw:modifiers="5400" 
						draw:enhanced-path="M 0 ?f12 L 0 ?f1 ?f2 0 ?f11 0 ?f11 ?f3 ?f4 ?f12 Z N M 0 ?f1 L ?f2 0 ?f11 0 ?f4 ?f1 Z N M ?f4 ?f12 L ?f4 ?f1 ?f11 0 ?f11 ?f3 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="$0 "/>
            <draw:equation draw:name="f1" draw:formula="top+?f0 "/>
            <draw:equation draw:name="f2" draw:formula="left+?f0 "/>
            <draw:equation draw:name="f3" draw:formula="bottom-?f0 "/>
            <draw:equation draw:name="f4" draw:formula="right-?f0 "/>
            <draw:equation draw:name="f5" draw:formula="right-?f2 "/>
            <draw:equation draw:name="f6" draw:formula="?f5 /2"/>
            <draw:equation draw:name="f7" draw:formula="?f2 +?f6 "/>
            <draw:equation draw:name="f8" draw:formula="bottom-?f1 "/>
            <draw:equation draw:name="f9" draw:formula="?f8 /2"/>
            <draw:equation draw:name="f10" draw:formula="?f1 +?f9 "/>
            <draw:equation draw:name="f11" draw:formula="right"/>
            <draw:equation draw:name="f12" draw:formula="bottom"/>
            <draw:handle draw:handle-position="left $0" 
									 draw:handle-switched="true" 
									 draw:handle-range-y-minimum="0" 
									 draw:handle-range-y-maximum="21600"/>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- Can -->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='can')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 88 21600"
						draw:glue-points="44 ?f6 44 0 0 10800 44 21600 88 10800"
						draw:text-areas="0 ?f6 88 ?f3" draw:type="can" draw:modifiers="5400"
						draw:enhanced-path="M 44 0 C 20 0 0 ?f2 0 ?f0 L 0 ?f3 C 0 ?f4 20 21600 44 21600 68 21600 88 ?f4 88 ?f3 L 88 ?f0 C 88 ?f2 68 0 44 0 Z N M 44 0 C 20 0 0 ?f2 0 ?f0 0 ?f5 20 ?f6 44 ?f6 68 ?f6 88 ?f5 88 ?f0 88 ?f2 68 0 44 0 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="$0 *2/4" />
            <draw:equation draw:name="f1" draw:formula="?f0 *6/11" />
            <draw:equation draw:name="f2" draw:formula="?f0 -?f1" />
            <draw:equation draw:name="f3" draw:formula="21600-?f0" />
            <draw:equation draw:name="f4" draw:formula="?f3 +?f1" />
            <draw:equation draw:name="f5" draw:formula="?f0 +?f1" />
            <draw:equation draw:name="f6" draw:formula="$0 *2/2" />
            <draw:equation draw:name="f7" draw:formula="44" />
            <draw:handle draw:handle-position="?f7 $0" draw:handle-range-y-minimum="0" draw:handle-range-y-maximum="10800" />
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- Cross (Added by A.Mathi as on 19/07/2007)-->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='plus')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
			<xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name ="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select ="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
            draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" 
            draw:path-stretchpoint-x="10800" 
            draw:path-stretchpoint-y="10800" 
            draw:text-areas="?f1 ?f1 ?f2 ?f3" 
            draw:type="cross" 
            draw:modifiers="5400" 
            draw:enhanced-path="M ?f1 0 L ?f2 0 ?f2 ?f1 21600 ?f1 21600 ?f3 ?f2 ?f3 ?f2 21600 ?f1 21600 ?f1 ?f3 0 ?f3 0 ?f1 ?f1 ?f1 ?f1 0 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="$0 *10799/10800"/>
            <draw:equation draw:name="f1" draw:formula="?f0 "/>
            <draw:equation draw:name="f2" draw:formula="right-?f0 "/>
            <draw:equation draw:name="f3" draw:formula="bottom-?f0 "/>
            <draw:handle draw:handle-position="$0 top" 
              draw:handle-switched="true" 
              draw:handle-range-x-minimum="0" 
              draw:handle-range-x-maximum="10800"/>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- "No" Symbol (Added by A.Mathi as on 19/07/2007)-->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='noSmoking')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
			<xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name ="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select ="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
            draw:glue-points="10800 0 3160 3160 0 10800 3160 18440 10800 21600 18440 18440 21600 10800 18440 3160" 
            draw:text-areas="3200 3200 18400 18400" 
            draw:type="forbidden" 
            draw:modifiers="2700" 
            draw:enhanced-path="U 10800 10800 10800 10800 0 23592960 Z B ?f0 ?f0 ?f1 ?f1 ?f9 ?f10 ?f11 ?f12 Z B ?f0 ?f0 ?f1 ?f1 ?f13 ?f14 ?f15 ?f16 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="$0 "/>
            <draw:equation draw:name="f1" draw:formula="21600-$0 "/>
            <draw:equation draw:name="f2" draw:formula="10800-$0 "/>
            <draw:equation draw:name="f3" draw:formula="$0 /2"/>
            <draw:equation draw:name="f4" draw:formula="sqrt(?f2 *?f2 -?f3 *?f3 )"/>
            <draw:equation draw:name="f5" draw:formula="10800-?f3 "/>
            <draw:equation draw:name="f6" draw:formula="10800+?f3 "/>
            <draw:equation draw:name="f7" draw:formula="10800-?f4 "/>
            <draw:equation draw:name="f8" draw:formula="10800+?f4 "/>
            <draw:equation draw:name="f9" draw:formula="(cos(45*(pi/180))*(?f5 -10800)+sin(45*(pi/180))*(?f7 -10800))+10800"/>
            <draw:equation draw:name="f10" draw:formula="-(sin(45*(pi/180))*(?f5 -10800)-cos(45*(pi/180))*(?f7 -10800))+10800"/>
            <draw:equation draw:name="f11" draw:formula="(cos(45*(pi/180))*(?f5 -10800)+sin(45*(pi/180))*(?f8 -10800))+10800"/>
            <draw:equation draw:name="f12" draw:formula="-(sin(45*(pi/180))*(?f5 -10800)-cos(45*(pi/180))*(?f8 -10800))+10800"/>
            <draw:equation draw:name="f13" draw:formula="(cos(45*(pi/180))*(?f6 -10800)+sin(45*(pi/180))*(?f8 -10800))+10800"/>
            <draw:equation draw:name="f14" draw:formula="-(sin(45*(pi/180))*(?f6 -10800)-cos(45*(pi/180))*(?f8 -10800))+10800"/>
            <draw:equation draw:name="f15" draw:formula="(cos(45*(pi/180))*(?f6 -10800)+sin(45*(pi/180))*(?f7 -10800))+10800"/>
            <draw:equation draw:name="f16" draw:formula="-(sin(45*(pi/180))*(?f6 -10800)-cos(45*(pi/180))*(?f7 -10800))+10800"/>
            <draw:handle draw:handle-position="$0 10800" 
              draw:handle-range-x-minimum="0" 
              draw:handle-range-x-maximum="7200"/>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!--  Folded Corner (Added by A.Mathi as on 19/07/2007) -->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='foldedCorner')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name="CreateShape">
			<xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId" />
            <xsl:with-param name ="prID" select ="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
            draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" 
            draw:text-areas="0 0 21600 ?f11" 
            draw:type="paper" 
            draw:modifiers="18900" 
            draw:enhanced-path="M 0 0 L 21600 0 21600 ?f0 ?f0 21600 0 21600 Z N M ?f0 21600 L ?f3 ?f0 C ?f8 ?f9 ?f10 ?f11 21600 ?f0 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="$0 "/>
            <draw:equation draw:name="f1" draw:formula="21600-?f0 "/>
            <draw:equation draw:name="f2" draw:formula="?f1 *8000/10800"/>
            <draw:equation draw:name="f3" draw:formula="21600-?f2 "/>
            <draw:equation draw:name="f4" draw:formula="?f1 /2"/>
            <draw:equation draw:name="f5" draw:formula="?f1 /4"/>
            <draw:equation draw:name="f6" draw:formula="?f1 /7"/>
            <draw:equation draw:name="f7" draw:formula="?f1 /16"/>
            <draw:equation draw:name="f8" draw:formula="?f3 +?f5 "/>
            <draw:equation draw:name="f9" draw:formula="?f0 +?f6 "/>
            <draw:equation draw:name="f10" draw:formula="21600-?f4 "/>
            <draw:equation draw:name="f11" draw:formula="?f0 +?f7 "/>
            <draw:handle draw:handle-position="$0 bottom" 
              draw:handle-range-x-minimum="10800" 
              draw:handle-range-x-maximum="21600"/>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!--  Lightning Bolt (Added by A.Mathi as on 20/07/2007) -->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='lightningBolt')">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
			<xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name ="grID" select ="$GraphicId" />
            <xsl:with-param name ="prID" select ="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 640 861" 
            draw:text-areas="257 295 414 566" 
            draw:type="non-primitive" 
            draw:enhanced-path="M 640 233 L 221 293 506 12 367 0 29 406 431 347 145 645 99 520 0 861 326 765 209 711 640 233 640 233 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
		<!--  Explosion 1 (Modified by A.Mathi) -->
		<xsl:when test ="(p:spPr/a:prstGeom/@prst='irregularSeal1')">
			<draw:custom-shape draw:layer="layout">
				<xsl:call-template name ="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
					<!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
										draw:glue-points="14623 106 106 8718 8590 21600 21600 13393"
										draw:text-areas="4680 6570 16140 13280"
										draw:type="mso-spt71"
										draw:enhanced-path="M 10901 5905 L 8458 2399 7417 6425 476 2399 4732 7722 106 8718 3828 11880 243 14689 5772 14041 4868 17719 7819 15730 8590 21600 10637 15038 13349 19840 14125 14561 18248 18195 16938 13044 21600 13393 17710 10579 21198 8242 16806 7417 18482 4560 14257 5429 14623 106 10901 5905 Z N">
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
      <!-- Chord (Added by A.Mathi as on 20/07/2007)-->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='chord')">
        <!-- warn if chord -->
        <xsl:message terminate="no">translation.oox2odf.shapesTypeChord</xsl:message>
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name ="grID" select ="$GraphicId" />
            <xsl:with-param name ="prID" select ="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 914400 914400" 
            draw:extrusion-allowed="true" 
            draw:text-areas="133911 133911 780489 780489" 
            draw:glue-points="780489 780489 457201 0 618845 390244" 
            draw:type="mso-spt100" draw:enhanced-path="M 780489 780489 W 0 0 914400 914400 780489 780489 457200 0 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- Left Bracket (Added by A.Mathi as on 20/07/2007) -->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='leftBracket')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
			<xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name ="grID" select ="$GraphicId" />
            <xsl:with-param name ="prID" select ="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="21600 0 0 10800 21600 21600" draw:text-areas="6350 ?f3 21600 ?f4" draw:type="left-bracket" draw:modifiers="1800" 
            draw:enhanced-path="M 21600 0 C 10800 0 0 ?f3 0 ?f1 L 0 ?f2 C 0 ?f4 10800 21600 21600 21600 N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="$0 /2"/>
            <draw:equation draw:name="f1" draw:formula="top+$0 "/>
            <draw:equation draw:name="f2" draw:formula="bottom-$0 "/>
            <draw:equation draw:name="f3" draw:formula="top+?f0 "/>
            <draw:equation draw:name="f4" draw:formula="bottom-?f0 "/>
            <draw:handle draw:handle-position="left $0" draw:handle-range-y-minimum="0" draw:handle-range-y-maximum="10800"/>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- Right Bracket (Added by A.Mathi as on 20/07/2007) -->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='rightBracket')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
			<xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name ="grID" select ="$GraphicId" />
            <xsl:with-param name ="prID" select ="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="0 0 0 21600 21600 10800" draw:text-areas="0 ?f3 15150 ?f4" draw:type="right-bracket" draw:modifiers="1800" 
            draw:enhanced-path="M 0 0 C 10800 0 21600 ?f3 21600 ?f1 L 21600 ?f2 C 21600 ?f4 10800 21600 0 21600 N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="$0 /2"/>
            <draw:equation draw:name="f1" draw:formula="top+$0 "/>
            <draw:equation draw:name="f2" draw:formula="bottom-$0 "/>
            <draw:equation draw:name="f3" draw:formula="top+?f0 "/>
            <draw:equation draw:name="f4" draw:formula="bottom-?f0 "/>
            <draw:handle draw:handle-position="right $0" draw:handle-range-y-minimum="0" draw:handle-range-y-maximum="10800"/>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- Left Brace (Added by A.Mathi as on 20/07/2007) -->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='leftBrace')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
			  <xsl:with-param name="sldId" select="$slideId" />
			  <xsl:with-param name ="grID" select ="$GraphicId" />
			  <xsl:with-param name ="prID" select ="$ParaId" />
			  <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
			  <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
			  <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
			  <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
		  </xsl:call-template>
		  <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="21600 0 0 10800 21600 21600" draw:text-areas="13800 ?f9 21600 ?f10" draw:type="left-brace" draw:modifiers="1800 10800" 
        draw:enhanced-path="M 21600 0 C 16200 0 10800 ?f0 10800 ?f1 L 10800 ?f2 C 10800 ?f3 5400 ?f4 0 ?f4 5400 ?f4 10800 ?f5 10800 ?f6 L 10800 ?f7 C 10800 ?f8 16200 21600 21600 21600 N">
        <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
          <xsl:attribute name ="draw:mirror-horizontal">
            <xsl:value-of select="'true'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
          <xsl:attribute name ="draw:mirror-vertical">
            <xsl:value-of select="'true'"/>
          </xsl:attribute>
        </xsl:if>
			  <draw:equation draw:name="f0" draw:formula="$0 /2"/>
			  <draw:equation draw:name="f1" draw:formula="$0 "/>
			  <draw:equation draw:name="f2" draw:formula="?f4 -$0 "/>
			  <draw:equation draw:name="f3" draw:formula="?f4 -?f0 "/>
			  <draw:equation draw:name="f4" draw:formula="$1 "/>
			  <draw:equation draw:name="f5" draw:formula="?f4 +?f0 "/>
			  <draw:equation draw:name="f6" draw:formula="?f4 +$0 "/>
			  <draw:equation draw:name="f7" draw:formula="21600-$0 "/>
			  <draw:equation draw:name="f8" draw:formula="21600-?f0 "/>
			  <draw:equation draw:name="f9" draw:formula="$0 *10000/31953"/>
			  <draw:equation draw:name="f10" draw:formula="21600-?f9 "/>
			  <draw:handle draw:handle-position="10800 $0" draw:handle-range-y-minimum="0" draw:handle-range-y-maximum="5400"/>
			  <draw:handle draw:handle-position="left $1" draw:handle-range-y-minimum="0" draw:handle-range-y-maximum="21600"/>
		  </draw:enhanced-geometry>
		  <xsl:copy-of select="$varHyperLinksForShapes" />
	  </draw:custom-shape>
  </xsl:when>
      <!-- Right Brace (Added by A.Mathi as on 23/07/2007) -->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='rightBrace')">
	  <draw:custom-shape draw:layer="layout" >
		  <xsl:call-template name ="CreateShape">
			<xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name ="grID" select ="$GraphicId" />
            <xsl:with-param name ="prID" select ="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
            draw:glue-points="0 0 0 21600 21600 10800" 
            draw:text-areas="0 ?f9 7800 ?f10" 
            draw:type="right-brace" 
            draw:modifiers="1800 10800" 
            draw:enhanced-path="M 0 0 C 5400 0 10800 ?f0 10800 ?f1 L 10800 ?f2 C 10800 ?f3 16200 ?f4 21600 ?f4 16200 ?f4 10800 ?f5 10800 ?f6 L 10800 ?f7 C 10800 ?f8 5400 21600 0 21600 N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="$0 /2"/>
            <draw:equation draw:name="f1" draw:formula="$0 "/>
            <draw:equation draw:name="f2" draw:formula="?f4 -$0 "/>
            <draw:equation draw:name="f3" draw:formula="?f4 -?f0 "/>
            <draw:equation draw:name="f4" draw:formula="$1 "/>
            <draw:equation draw:name="f5" draw:formula="?f4 +?f0 "/>
            <draw:equation draw:name="f6" draw:formula="?f4 +$0 "/>
            <draw:equation draw:name="f7" draw:formula="21600-$0 "/>
            <draw:equation draw:name="f8" draw:formula="21600-?f0 "/>
            <draw:equation draw:name="f9" draw:formula="$0 *10000/31953"/>
            <draw:equation draw:name="f10" draw:formula="21600-?f9 "/>
            <draw:handle draw:handle-position="10800 $0" draw:handle-range-y-minimum="0" draw:handle-range-y-maximum="5400"/>
            <draw:handle draw:handle-position="right $1" draw:handle-range-y-minimum="0" draw:handle-range-y-maximum="21600"/>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
		<!-- Rectangular Callout (modified by A.Mathi) -->
		<xsl:when test ="(p:spPr/a:prstGeom/@prst='wedgeRectCallout')">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
					<!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
				  draw:glue-points="10800 0 0 10800 10800 21600 21600 10800 ?f40 ?f41" 
				  draw:text-areas="0 0 21600 21600" 
				  draw:type="rectangular-callout"
				  draw:enhanced-path="M 0 0 L 0 3590 ?f2 ?f3 0 8970 0 12630 ?f4 ?f5 0 18010 0 21600 3590 21600 ?f6 ?f7 8970 21600 12630 21600 ?f8 ?f9 18010 21600 21600 21600 21600 18010 ?f10 ?f11 21600 12630 21600 8970 ?f12 ?f13 21600 3590 21600 0 18010 0 ?f14 ?f15 12630 0 8970 0 ?f16 ?f17 3590 0 0 0 Z N">
					<xsl:call-template name="tmpCalloutAdj">
						<xsl:with-param name="defaultVal" select="'6300 24300'"/>
					</xsl:call-template>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="$0 -10800"/>
					<draw:equation draw:name="f1" draw:formula="$1 -10800"/>
					<draw:equation draw:name="f2" draw:formula="if(?f18 ,$0 ,0)"/>
					<draw:equation draw:name="f3" draw:formula="if(?f18 ,$1 ,6280)"/>
					<draw:equation draw:name="f4" draw:formula="if(?f23 ,$0 ,0)"/>
					<draw:equation draw:name="f5" draw:formula="if(?f23 ,$1 ,15320)"/>
					<draw:equation draw:name="f6" draw:formula="if(?f26 ,$0 ,6280)"/>
					<draw:equation draw:name="f7" draw:formula="if(?f26 ,$1 ,21600)"/>
					<draw:equation draw:name="f8" draw:formula="if(?f29 ,$0 ,15320)"/>
					<draw:equation draw:name="f9" draw:formula="if(?f29 ,$1 ,21600)"/>
					<draw:equation draw:name="f10" draw:formula="if(?f32 ,$0 ,21600)"/>
					<draw:equation draw:name="f11" draw:formula="if(?f32 ,$1 ,15320)"/>
					<draw:equation draw:name="f12" draw:formula="if(?f34 ,$0 ,21600)"/>
					<draw:equation draw:name="f13" draw:formula="if(?f34 ,$1 ,6280)"/>
					<draw:equation draw:name="f14" draw:formula="if(?f36 ,$0 ,15320)"/>
					<draw:equation draw:name="f15" draw:formula="if(?f36 ,$1 ,0)"/>
					<draw:equation draw:name="f16" draw:formula="if(?f38 ,$0 ,6280)"/>
					<draw:equation draw:name="f17" draw:formula="if(?f38 ,$1 ,0)"/>
					<draw:equation draw:name="f18" draw:formula="if($0 ,-1,?f19 )"/>
					<draw:equation draw:name="f19" draw:formula="if(?f1 ,-1,?f22 )"/>
					<draw:equation draw:name="f20" draw:formula="abs(?f0 )"/>
					<draw:equation draw:name="f21" draw:formula="abs(?f1 )"/>
					<draw:equation draw:name="f22" draw:formula="?f20 -?f21 "/>
					<draw:equation draw:name="f23" draw:formula="if($0 ,-1,?f24 )"/>
					<draw:equation draw:name="f24" draw:formula="if(?f1 ,?f22 ,-1)"/>
					<draw:equation draw:name="f25" draw:formula="$1 -21600"/>
					<draw:equation draw:name="f26" draw:formula="if(?f25 ,?f27 ,-1)"/>
					<draw:equation draw:name="f27" draw:formula="if(?f0 ,-1,?f28 )"/>
					<draw:equation draw:name="f28" draw:formula="?f21 -?f20 "/>
					<draw:equation draw:name="f29" draw:formula="if(?f25 ,?f30 ,-1)"/>
					<draw:equation draw:name="f30" draw:formula="if(?f0 ,?f28 ,-1)"/>
					<draw:equation draw:name="f31" draw:formula="$0 -21600"/>
					<draw:equation draw:name="f32" draw:formula="if(?f31 ,?f33 ,-1)"/>
					<draw:equation draw:name="f33" draw:formula="if(?f1 ,?f22 ,-1)"/>
					<draw:equation draw:name="f34" draw:formula="if(?f31 ,?f35 ,-1)"/>
					<draw:equation draw:name="f35" draw:formula="if(?f1 ,-1,?f22 )"/>
					<draw:equation draw:name="f36" draw:formula="if($1 ,-1,?f37 )"/>
					<draw:equation draw:name="f37" draw:formula="if(?f0 ,?f28 ,-1)"/>
					<draw:equation draw:name="f38" draw:formula="if($1 ,-1,?f39 )"/>
					<draw:equation draw:name="f39" draw:formula="if(?f0 ,-1,?f28 )"/>
					<draw:equation draw:name="f40" draw:formula="$0 "/>
					<draw:equation draw:name="f41" draw:formula="$1 "/>
					<draw:handle draw:handle-position="$0 $1"/>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Rounded Rectangular Callout (modified by A.Mathi) -->
		<xsl:when test ="(p:spPr/a:prstGeom/@prst='wedgeRoundRectCallout')">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
					<!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
				  draw:text-areas="800 800 20800 20800" 
				  draw:type="round-rectangular-callout"
				  draw:enhanced-path="M 3590 0 X 0 3590 L ?f2 ?f3 0 8970 0 12630 ?f4 ?f5 0 18010 Y 3590 21600 L ?f6 ?f7 8970 21600 12630 21600 ?f8 ?f9 18010 21600 X 21600 18010 L ?f10 ?f11 21600 12630 21600 8970 ?f12 ?f13 21600 3590 Y 18010 0 L ?f14 ?f15 12630 0 8970 0 ?f16 ?f17 Z N">
					<xsl:call-template name="tmpCalloutAdj">
						<xsl:with-param name="defaultVal" select="'6300 24300'"/>
					</xsl:call-template>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="$0 -10800"/>
					<draw:equation draw:name="f1" draw:formula="$1 -10800"/>
					<draw:equation draw:name="f2" draw:formula="if(?f18 ,$0 ,0)"/>
					<draw:equation draw:name="f3" draw:formula="if(?f18 ,$1 ,6280)"/>
					<draw:equation draw:name="f4" draw:formula="if(?f23 ,$0 ,0)"/>
					<draw:equation draw:name="f5" draw:formula="if(?f23 ,$1 ,15320)"/>
					<draw:equation draw:name="f6" draw:formula="if(?f26 ,$0 ,6280)"/>
					<draw:equation draw:name="f7" draw:formula="if(?f26 ,$1 ,21600)"/>
					<draw:equation draw:name="f8" draw:formula="if(?f29 ,$0 ,15320)"/>
					<draw:equation draw:name="f9" draw:formula="if(?f29 ,$1 ,21600)"/>
					<draw:equation draw:name="f10" draw:formula="if(?f32 ,$0 ,21600)"/>
					<draw:equation draw:name="f11" draw:formula="if(?f32 ,$1 ,15320)"/>
					<draw:equation draw:name="f12" draw:formula="if(?f34 ,$0 ,21600)"/>
					<draw:equation draw:name="f13" draw:formula="if(?f34 ,$1 ,6280)"/>
					<draw:equation draw:name="f14" draw:formula="if(?f36 ,$0 ,15320)"/>
					<draw:equation draw:name="f15" draw:formula="if(?f36 ,$1 ,0)"/>
					<draw:equation draw:name="f16" draw:formula="if(?f38 ,$0 ,6280)"/>
					<draw:equation draw:name="f17" draw:formula="if(?f38 ,$1 ,0)"/>
					<draw:equation draw:name="f18" draw:formula="if($0 ,-1,?f19 )"/>
					<draw:equation draw:name="f19" draw:formula="if(?f1 ,-1,?f22 )"/>
					<draw:equation draw:name="f20" draw:formula="abs(?f0 )"/>
					<draw:equation draw:name="f21" draw:formula="abs(?f1 )"/>
					<draw:equation draw:name="f22" draw:formula="?f20 -?f21 "/>
					<draw:equation draw:name="f23" draw:formula="if($0 ,-1,?f24 )"/>
					<draw:equation draw:name="f24" draw:formula="if(?f1 ,?f22 ,-1)"/>
					<draw:equation draw:name="f25" draw:formula="$1 -21600"/>
					<draw:equation draw:name="f26" draw:formula="if(?f25 ,?f27 ,-1)"/>
					<draw:equation draw:name="f27" draw:formula="if(?f0 ,-1,?f28 )"/>
					<draw:equation draw:name="f28" draw:formula="?f21 -?f20 "/>
					<draw:equation draw:name="f29" draw:formula="if(?f25 ,?f30 ,-1)"/>
					<draw:equation draw:name="f30" draw:formula="if(?f0 ,?f28 ,-1)"/>
					<draw:equation draw:name="f31" draw:formula="$0 -21600"/>
					<draw:equation draw:name="f32" draw:formula="if(?f31 ,?f33 ,-1)"/>
					<draw:equation draw:name="f33" draw:formula="if(?f1 ,?f22 ,-1)"/>
					<draw:equation draw:name="f34" draw:formula="if(?f31 ,?f35 ,-1)"/>
					<draw:equation draw:name="f35" draw:formula="if(?f1 ,-1,?f22 )"/>
					<draw:equation draw:name="f36" draw:formula="if($1 ,-1,?f37 )"/>
					<draw:equation draw:name="f37" draw:formula="if(?f0 ,?f28 ,-1)"/>
					<draw:equation draw:name="f38" draw:formula="if($1 ,-1,?f39 )"/>
					<draw:equation draw:name="f39" draw:formula="if(?f0 ,-1,?f28 )"/>
					<draw:equation draw:name="f40" draw:formula="$0 "/>
					<draw:equation draw:name="f41" draw:formula="$1 "/>
					<draw:handle draw:handle-position="$0 $1"/>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Oval Callout (modified by A.Mathi) -->
		<xsl:when test ="(p:spPr/a:prstGeom/@prst='wedgeEllipseCallout')">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
					<!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
				  draw:glue-points="10800 0 3160 3160 0 10800 3160 18440 10800 21600 18440 18440 21600 10800 18440 3160 ?f14 ?f15" 
				  draw:text-areas="3200 3200 18400 18400" 
				  draw:type="round-callout"
       				  draw:enhanced-path="W 0 0 21600 21600 ?f22 ?f23 ?f18 ?f19 L ?f14 ?f15 Z N">
					<xsl:call-template name="tmpCalloutAdj">
						<xsl:with-param name="defaultVal" select="'6300 24300'"/>
					</xsl:call-template>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="$0 -10800"/>
					<draw:equation draw:name="f1" draw:formula="$1 -10800"/>
					<draw:equation draw:name="f2" draw:formula="?f0 *?f0 "/>
					<draw:equation draw:name="f3" draw:formula="?f1 *?f1 "/>
					<draw:equation draw:name="f4" draw:formula="?f2 +?f3 "/>
					<draw:equation draw:name="f5" draw:formula="sqrt(?f4 )"/>
					<draw:equation draw:name="f6" draw:formula="?f5 -10800"/>
					<draw:equation draw:name="f7" draw:formula="atan2(?f1 ,?f0 )/(pi/180)"/>
					<draw:equation draw:name="f8" draw:formula="?f7 -10"/>
					<draw:equation draw:name="f9" draw:formula="?f7 +10"/>
					<draw:equation draw:name="f10" draw:formula="10800*cos(?f7 *(pi/180))"/>
					<draw:equation draw:name="f11" draw:formula="10800*sin(?f7 *(pi/180))"/>
					<draw:equation draw:name="f12" draw:formula="?f10 +10800"/>
					<draw:equation draw:name="f13" draw:formula="?f11 +10800"/>
					<draw:equation draw:name="f14" draw:formula="if(?f6 ,$0 ,?f12 )"/>
					<draw:equation draw:name="f15" draw:formula="if(?f6 ,$1 ,?f13 )"/>
					<draw:equation draw:name="f16" draw:formula="10800*cos(?f8 *(pi/180))"/>
					<draw:equation draw:name="f17" draw:formula="10800*sin(?f8 *(pi/180))"/>
					<draw:equation draw:name="f18" draw:formula="?f16 +10800"/>
					<draw:equation draw:name="f19" draw:formula="?f17 +10800"/>
					<draw:equation draw:name="f20" draw:formula="10800*cos(?f9 *(pi/180))"/>
					<draw:equation draw:name="f21" draw:formula="10800*sin(?f9 *(pi/180))"/>
					<draw:equation draw:name="f22" draw:formula="?f20 +10800"/>
					<draw:equation draw:name="f23" draw:formula="?f21 +10800"/>
					<draw:handle draw:handle-position="$0 $1"/>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Cloud Callout (modified by A.Mathi) -->
		<xsl:when test ="(p:spPr/a:prstGeom/@prst='cloudCallout')">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
					<!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
				  draw:text-areas="3000 3320 17110 17330" 
				  draw:type="cloud-callout"
				  draw:enhanced-path="M 1930 7160 C 1530 4490 3400 1970 5270 1970 5860 1950 6470 2210 6970 2600 7450 1390 8340 650 9340 650 10004 690 10710 1050 11210 1700 11570 630 12330 0 13150 0 13840 0 14470 460 14870 1160 15330 440 16020 0 16740 0 17910 0 18900 1130 19110 2710 20240 3150 21060 4580 21060 6220 21060 6720 21000 7200 20830 7660 21310 8460 21600 9450 21600 10460 21600 12750 20310 14680 18650 15010 18650 17200 17370 18920 15770 18920 15220 18920 14700 18710 14240 18310 13820 20240 12490 21600 11000 21600 9890 21600 8840 20790 8210 19510 7620 20000 7930 20290 6240 20290 4850 20290 3570 19280 2900 17640 1300 17600 480 16300 480 14660 480 13900 690 13210 1070 12640 380 12160 0 11210 0 10120 0 8590 840 7330 1930 7160 Z N M 1930 7160 C 1950 7410 2040 7690 2090 7920 F N M 6970 2600 C 7200 2790 7480 3050 7670 3310 F N M 11210 1700 C 11130 1910 11080 2160 11030 2400 F N M 14870 1160 C 14720 1400 14640 1720 14540 2010 F N M 19110 2710 C 19130 2890 19230 3290 19190 3380 F N M 20830 7660 C 20660 8170 20430 8620 20110 8990 F N M 18660 15010 C 18740 14200 18280 12200 17000 11450 F N M 14240 18310 C 14320 17980 14350 17680 14370 17360 F N M 8220 19510 C 8060 19250 7960 18950 7860 18640 F N M 2900 17640 C 3090 17600 3280 17540 3460 17450 F N M 1070 12640 C 1400 12900 1780 13130 2330 13040 F N U ?f17 ?f18 1800 1800 0 23592960 Z N U ?f19 ?f20 1200 1200 0 23592960 Z N U ?f13 ?f14 700 700 0 23592960 Z N">
					<xsl:call-template name="tmpCalloutAdj">
						<xsl:with-param name="defaultVal" select="'6300 24300'"/>
					</xsl:call-template>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="$0 -10800"/>
					<draw:equation draw:name="f1" draw:formula="$1 -10800"/>
					<draw:equation draw:name="f2" draw:formula="atan2(?f1 ,?f0 )/(pi/180)"/>
					<draw:equation draw:name="f3" draw:formula="10800*cos(?f2 *(pi/180))"/>
					<draw:equation draw:name="f4" draw:formula="10800*sin(?f2 *(pi/180))"/>
					<draw:equation draw:name="f5" draw:formula="?f3 +10800"/>
					<draw:equation draw:name="f6" draw:formula="?f4 +10800"/>
					<draw:equation draw:name="f7" draw:formula="$0 -?f5 "/>
					<draw:equation draw:name="f8" draw:formula="$1 -?f6 "/>
					<draw:equation draw:name="f9" draw:formula="?f7 /3"/>
					<draw:equation draw:name="f10" draw:formula="?f8 /3"/>
					<draw:equation draw:name="f11" draw:formula="?f7 *2/3"/>
					<draw:equation draw:name="f12" draw:formula="?f8 *2/3"/>
					<draw:equation draw:name="f13" draw:formula="$0 "/>
					<draw:equation draw:name="f14" draw:formula="$1 "/>
					<draw:equation draw:name="f15" draw:formula="?f3 /12"/>
					<draw:equation draw:name="f16" draw:formula="?f4 /12"/>
					<draw:equation draw:name="f17" draw:formula="?f9 +?f5 -?f15 "/>
					<draw:equation draw:name="f18" draw:formula="?f10 +?f6 -?f16 "/>
					<draw:equation draw:name="f19" draw:formula="?f11 +?f5 "/>
					<draw:equation draw:name="f20" draw:formula="?f12 +?f6 "/>
					<draw:handle draw:handle-position="$0 $1"/>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Line Callout 1) -->
		<xsl:when test ="(p:spPr/a:prstGeom/@prst='borderCallout1')">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
					<!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
					draw:mirror-horizontal="false"
					draw:type="line-callout-1"
					draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 Z N M ?f0 ?f1 L ?f2 ?f3 N">
					<xsl:call-template name="tmpCalloutLineAdj">
						<xsl:with-param name="defaultVal" select="'-8300 24500 -1800 4000'"/>
					</xsl:call-template>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="$0 "/>
					<draw:equation draw:name="f1" draw:formula="$1 "/>
					<draw:equation draw:name="f2" draw:formula="$2 "/>
					<draw:equation draw:name="f3" draw:formula="$3 "/>
					<draw:equation draw:name="f4" draw:formula="$4 "/>
					<draw:equation draw:name="f5" draw:formula="$5 "/>
					<draw:equation draw:name="f6" draw:formula="$6 "/>
					<draw:equation draw:name="f7" draw:formula="$7 "/>
					<draw:handle draw:handle-position="$0 $1"/>
					<draw:handle draw:handle-position="$2 $3"/>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Line Callout 2) -->
		<xsl:when test ="(p:spPr/a:prstGeom/@prst='borderCallout2')">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
					<!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
					draw:mirror-horizontal="false"
					draw:type="line-callout-2"
					draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 Z N M ?f0 ?f1 L ?f2 ?f3 N M ?f2 ?f3 L ?f4 ?f5 N">
					<xsl:call-template name="tmpCalloutLineAdj">
						<xsl:with-param name="defaultVal" select="'-10000 24500 -3600 4000 -1800 4000'"/>
					</xsl:call-template>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="$0 "/>
					<draw:equation draw:name="f1" draw:formula="$1 "/>
					<draw:equation draw:name="f2" draw:formula="$2 "/>
					<draw:equation draw:name="f3" draw:formula="$3 "/>
					<draw:equation draw:name="f4" draw:formula="$4 "/>
					<draw:equation draw:name="f5" draw:formula="$5 "/>
					<draw:equation draw:name="f6" draw:formula="$6 "/>
					<draw:equation draw:name="f7" draw:formula="$7 "/>
					<draw:handle draw:handle-position="$0 $1"/>
					<draw:handle draw:handle-position="$2 $3"/>
					<draw:handle draw:handle-position="$4 $5"/>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Line Callout 3) -->
		<xsl:when test ="(p:spPr/a:prstGeom/@prst='borderCallout3')">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
					<!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
					draw:type="mso-spt49"
					draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 Z N M ?f6 ?f7 F L ?f4 ?f5 ?f2 ?f3 ?f0 ?f1 N">
					<xsl:call-template name="tmpCalloutLineAdj">
						<xsl:with-param name="defaultVal" select="'-1800 0 -3600 0 -3600 0 -1800 4000'"/>
					</xsl:call-template>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="$0 "/>
					<draw:equation draw:name="f1" draw:formula="$1 "/>
					<draw:equation draw:name="f2" draw:formula="$2 "/>
					<draw:equation draw:name="f3" draw:formula="$3 "/>
					<draw:equation draw:name="f4" draw:formula="$4 "/>
					<draw:equation draw:name="f5" draw:formula="$5 "/>
					<draw:equation draw:name="f6" draw:formula="$6 "/>
					<draw:equation draw:name="f7" draw:formula="$7 "/>
					<draw:handle draw:handle-position="$0 $1"/>
					<draw:handle draw:handle-position="$2 $3"/>
					<draw:handle draw:handle-position="$4 $5"/>
					<draw:handle draw:handle-position="$6 $7"/>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		
      <!-- Bent Arrow (Added by A.Mathi as on 23/07/2007) -->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='bentArrow')">
        <!-- warn if bent Arrow -->
        <xsl:message terminate="no">translation.oox2odf.shapesTypeBentArrow</xsl:message>
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
			<xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name ="grID" select ="$GraphicId" />
            <xsl:with-param name ="prID" select ="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!-- End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering -->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 813816 868680" draw:extrusion-allowed="true" 
            draw:text-areas="0 0 813816 868680" draw:glue-points="610362 0 610362 406908 101727 868680 813816 203454" 
            draw:type="mso-spt100" 
            draw:enhanced-path="M 0 868680 L 0 457772 W 0 101727 712090 813817 0 457772 356046 101727 L 610362 101727 L 610362 0 L 813816 203454 L 610362 406908 L 610362 305181 L 356045 305181 A 203454 305181 508636 610363 356045 305181 203454 457772 L 203454 868680 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- U-Turn Arrow (Added by A.Mathi as on 23/07/2007) -->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='uturnArrow')">
        <!-- warn if U-Turn Arrow -->
        <xsl:message terminate="no">translation.oox2odf.shapesTypeUTurnArrow</xsl:message>
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
			<xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name ="grID" select ="$GraphicId" />
            <xsl:with-param name ="prID" select ="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!-- End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering -->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 886968 877824" 
            draw:extrusion-allowed="true" draw:text-areas="0 0 886968 877824" 
            draw:glue-points="448056 438912 667512 658368 886968 438912 388620 0 109728 877824" 
            draw:type="mso-spt100" draw:enhanced-path="M 0 877824 L 0 384048 W 0 0 768096 768096 0 384048 384049 0 L 393192 0 W 9144 0 777240 768096 393192 0 777240 384049 L 777240 438912 L 886968 438912 L 667512 658368 L 448056 438912 L 557784 438912 L 557784 384048 A 228600 219456 557784 548640 557784 384048 393192 219456 L 384048 219456 A 219456 219456 548640 548640 384048 219456 219456 384048 L 219456 877824 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
		<!-- Quad Arrow -->
		<xsl:when test ="(p:spPr/a:prstGeom/@prst='quadArrow')">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering -->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
					<!-- End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering -->
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
										draw:text-areas="0 0 21600 21600"
										draw:type="quad-arrow"
										draw:modifiers="6500 8600 4300"
										draw:enhanced-path="M 0 10800 L ?f0 ?f1 ?f0 ?f2 ?f2 ?f2 ?f2 ?f0 ?f1 ?f0 10800 0 ?f3 ?f0 ?f4 ?f0 ?f4 ?f2 ?f5 ?f2 ?f5 ?f1 21600 10800 ?f5 ?f3 ?f5 ?f4 ?f4 ?f4 ?f4 ?f5 ?f3 ?f5 10800 21600 ?f1 ?f5 ?f2 ?f5 ?f2 ?f4 ?f0 ?f4 ?f0 ?f3 Z N">
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="$2 "/>
					<draw:equation draw:name="f1" draw:formula="$0 "/>
					<draw:equation draw:name="f2" draw:formula="$1 "/>
					<draw:equation draw:name="f3" draw:formula="21600-$0 "/>
					<draw:equation draw:name="f4" draw:formula="21600-$1 "/>
					<draw:equation draw:name="f5" draw:formula="21600-$2 "/>
					<draw:handle draw:handle-position="$1 $2" draw:handle-range-x-minimum="$0" draw:handle-range-x-maximum="10800" draw:handle-range-y-minimum="0" draw:handle-range-y-maximum="$0"/>
					<draw:handle draw:handle-position="$0 top" draw:handle-range-x-minimum="$2" draw:handle-range-x-maximum="$1"/>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Block Arc -->
		<xsl:when test ="(p:spPr/a:prstGeom/@prst='blockArc')">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering -->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
					<!-- End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering -->
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
										draw:type="block-arc"
										draw:modifiers="180 5400"
										draw:enhanced-path="B 0 0 21600 21600 ?f4 ?f3 ?f2 ?f3 W ?f5 ?f5 ?f6 ?f6 ?f2 ?f3 ?f4 ?f3 Z N">
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="10800*cos($0 *(pi/180))"/>
					<draw:equation draw:name="f1" draw:formula="10800*sin($0 *(pi/180))"/>
					<draw:equation draw:name="f2" draw:formula="?f0 +10800"/>
					<draw:equation draw:name="f3" draw:formula="?f1 +10800"/>
					<draw:equation draw:name="f4" draw:formula="21600-?f2 "/>
					<draw:equation draw:name="f5" draw:formula="10800-$1 "/>
					<draw:equation draw:name="f6" draw:formula="10800+$1 "/>
					<draw:equation draw:name="f7" draw:formula="?f5 *cos($0 *(pi/180))"/>
					<draw:equation draw:name="f8" draw:formula="?f5 *sin($0 *(pi/180))"/>
					<draw:handle draw:handle-position="$1 $0" draw:handle-polar="10800 10800" draw:handle-radius-range-minimum="0" draw:handle-radius-range-maximum="10800"/>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Notched Right Arrow -->
		<xsl:when test ="(p:spPr/a:prstGeom/@prst='notchedRightArrow')">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering -->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
					<!-- End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering -->
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
										draw:text-areas="0 0 21600 21600"
										draw:type="notched-right-arrow"
										draw:modifiers="16200 5400"
	                                    draw:enhanced-path="M 0 ?f1 L ?f0 ?f1 ?f0 0 21600 10800 ?f0 21600 ?f0 ?f2 0 ?f2 ?f5 10800 0 ?f1 Z N">
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="$0 "/>
					<draw:equation draw:name="f1" draw:formula="$1 "/>
					<draw:equation draw:name="f2" draw:formula="21600-$1 "/>
					<draw:equation draw:name="f3" draw:formula="21600-$0 "/>
					<draw:equation draw:name="f4" draw:formula="10800-$1 "/>
					<draw:equation draw:name="f5" draw:formula="?f3 *?f4 /10800"/>
					<draw:handle draw:handle-position="$0 $1" draw:handle-range-x-minimum="0" draw:handle-range-x-maximum="21600" draw:handle-range-y-minimum="0" draw:handle-range-y-maximum="10800"/>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Pentagon -->
		<xsl:when test ="(p:spPr/a:prstGeom/@prst='homePlate')">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering -->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
					<!-- End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering -->
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
										draw:text-areas="0 0 21600 21600"
										draw:type="pentagon-right"
										draw:modifiers="16200"
										draw:enhanced-path="M 0 0 L ?f0 0 21600 10800 ?f0 21600 0 21600 Z N">
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="$0 "/>
					<draw:handle draw:handle-position="$0 top" draw:handle-range-x-minimum="0" draw:handle-range-x-maximum="21600"/>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Chevron -->
		<xsl:when test ="(p:spPr/a:prstGeom/@prst='chevron')">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering -->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
					<!-- End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering -->
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
										draw:text-areas="0 0 21600 21600"
										draw:type="chevron"
										draw:modifiers="16200"
										draw:enhanced-path="M 0 0 L ?f0 0 21600 10800 ?f0 21600 0 21600 ?f1 10800 Z N">
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="$0 "/>
					<draw:equation draw:name="f1" draw:formula="21600-?f0 "/>
					<draw:handle draw:handle-position="$0 top" draw:handle-range-x-minimum="0" draw:handle-range-x-maximum="21600"/>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>

		<!--Equation Shapes-->
		<!--Equal-->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='mathEqual'">
			<draw:custom-shape draw:layer="layout">
				<xsl:call-template name="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering -->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 1295400 990600" draw:extrusion-allowed="true"
										draw:text-areas="171705 204064 1123695 786536"
										draw:glue-points="1123695 320558 1123695 670042 647700 786536 171705 320558 171705 670042 647700 204064"
										draw:enhanced-path="M 171705 204064 L 1123695 204064 L 1123695 437053 L 171705 437053 Z M 171705 553547 L 1123695 553547 L 1123695 786536 L 171705 786536 Z N">
					<xsl:if test="p:spPr/a:prstGeom/@prst='mathEqual'">
						<xsl:attribute name ="draw:type">
							<xsl:value-of select="'mathequal'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!--Not Equal-->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='mathNotEqual'">
			<draw:custom-shape draw:layer="layout">
				<xsl:call-template name="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering -->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 1371600 1066800" draw:extrusion-allowed="true"
										draw:text-areas="181806 219761 1189794 847039"
										draw:glue-points="1189794 345216 1189794 721583 507275 1023891 181806 345216 181806 721583 864324 42908"
										draw:enhanced-path="M 181806 219761 L 666448 219761 L 746435 0 L 982214 85817 L 933462 219761 L 1189794 219761 L 1189794 470672 L 842138 470672 L 796476 596128 L 1189794 596128 L 1189794 847039 L 705152 847039 L 625165 1066800 L 389386 980983 L 438138 847039 L 181806 847039 L 181806 596128 L 529462 596128 L 575124 470672 L 181806 470672 Z N">
					<xsl:if test="p:spPr/a:prstGeom/@prst='mathNotEqual'">
						<xsl:attribute name ="draw:type">
							<xsl:value-of select="'mathnotequal'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!--Plus-->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='mathPlus'">
			<draw:custom-shape draw:layer="layout">
				<xsl:call-template name="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering -->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 1524000 1371600" draw:extrusion-allowed="true"
										draw:text-areas="202006 524500 1321994 847100"
										draw:glue-points="1321994 685800 762000 1189794 202006 685800 762000 181806"
										draw:enhanced-path="M 202006 524500 L 600700 524500 L 600700 181806 L 923300 181806 L 923300 524500 L 1321994 524500 L 1321994 847100 L 923300 847100 L 923300 1189794 L 600700 1189794 L 600700 847100 L 202006 847100 Z N">
					<xsl:if test="p:spPr/a:prstGeom/@prst='mathPlus'">
						<xsl:attribute name ="draw:type">
							<xsl:value-of select="'mathplus'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!--Minus-->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='mathMinus'">
			<draw:custom-shape draw:layer="layout">
				<xsl:call-template name="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering -->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 1371600 1219200" draw:extrusion-allowed="true"
										draw:text-areas="181806 466222 1189794 752978"
										draw:glue-points="1189794 609600 685800 752978 181806 609600 685800 466222"
										draw:enhanced-path="M 181806 466222 L 1189794 466222 L 1189794 752978 L 181806 752978 Z N">
					<xsl:if test="p:spPr/a:prstGeom/@prst='mathMinus'">
						<xsl:attribute name ="draw:type">
							<xsl:value-of select="'mathminus'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!--Multiply-->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='mathMultiply'">
			<draw:custom-shape draw:layer="layout">
				<xsl:call-template name="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering -->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 1676400 1600200"
										draw:extrusion-allowed="true"
										draw:text-areas="272693 248204 1403707 1351996"
										draw:glue-points="402629 384328 1273771 384328 1273771 1215872 402629 1215872"
										draw:enhanced-path="M 272693 520452 L 532566 248204 L 838200 539946 L 1143834 248204 L 1403707 520452 L 1110742 800100 L 1403707 1079748 L 1143834 1351996 L 838200 1060254 L 532566 1351996 L 272693 1079748 L 565658 800100 Z N">
					<xsl:if test="p:spPr/a:prstGeom/@prst='mathMultiply'">
						<xsl:attribute name ="draw:type">
							<xsl:value-of select="'mathmultiply'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!--Division-->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='mathDivide'">
			<draw:custom-shape draw:layer="layout">
				<xsl:call-template name="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering -->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 3200400 1524000"
										draw:extrusion-allowed="true"
										draw:text-areas="0 0 3200400 1524000"
										draw:glue-points="1600200 179679 1779422 358901 1600200 538123 1420978 358901 1600200 179679 1600200 1344321 1420978 1165099 1600200 985877 1779422 1165099 1600200 1344321 424213 582778 2776186 582778 2776186 941222 424213 941222 424213 582778"
										draw:enhanced-path="M 1600200 179679 C 1699182 179679 1779422 259920 1779422 358901 C 1779422 457883 1699181 538123 1600200 538123 C 1501218 538123 1420978 457882 1420978 358901 C 1420978 259919 1501219 179679 1600200 179679 Z M 1600200 1344321 C 1501218 1344321 1420978 1264080 1420978 1165099 C 1420978 1066117 1501219 985877 1600200 985877 C 1699182 985877 1779422 1066118 1779422 1165099 C 1779422 1264081 1699181 1344321 1600200 1344321 Z M 424213 582778 L 2776187 582778 L 2776187 941222 L 424213 941222 L 424213 582778 Z N">
					<xsl:if test="p:spPr/a:prstGeom/@prst='mathDivide'">
						<xsl:attribute name ="draw:type">
							<xsl:value-of select="'mathdivide'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
     
      <!--End-->
      <!-- Connectors -->
      <!-- Line -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst = 'line'">
        <draw:line draw:layer="layout">
          <xsl:call-template name ="DrawLine">
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
          </xsl:call-template>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:line>
      </xsl:when>
      <!-- Straight Connector-->
      <xsl:when test ="p:spPr/a:prstGeom/@prst[contains(., 'straightConnector')]">
        <draw:line draw:layer="layout">
          <xsl:call-template name ="DrawLine">
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
          </xsl:call-template>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:line >
      </xsl:when>
      <!-- Elbow Connector-->
      <xsl:when test ="p:spPr/a:prstGeom/@prst[contains(., 'bentConnector')]">
        <draw:connector draw:layer="layout">
          <xsl:call-template name ="DrawLine">
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
          </xsl:call-template>
          <xsl:copy-of select="$varHyperLinksForConnectors" />
        </draw:connector >
      </xsl:when>
      <!--Curved Connector-->
      <xsl:when test ="p:spPr/a:prstGeom/@prst[contains(., 'curvedConnector')]">
        <draw:connector draw:layer="layout" draw:type="curve">
          <xsl:call-template name ="DrawLine">
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
          </xsl:call-template>
          <xsl:copy-of select="$varHyperLinksForConnectors" />
        </draw:connector >
      </xsl:when>
      <!-- Custom shapes: -->
      <!-- Rounded  Rectangle -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='roundRect'">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter inserted by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
            draw:path-stretchpoint-x="10800" draw:path-stretchpoint-y="10800"
            draw:text-areas="?f3 ?f4 ?f5 ?f6" draw:type="round-rectangle" draw:modifiers="3600"
            draw:enhanced-path="M ?f7 0 X 0 ?f8 L 0 ?f9 Y ?f7 21600 L ?f10 21600 X 21600 ?f9 L 21600 ?f8 Y ?f10 0 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="45"/>
            <draw:equation draw:name="f1" draw:formula="$0 *sin(?f0 *(pi/180))"/>
            <draw:equation draw:name="f2" draw:formula="?f1 *3163/7636"/>
            <draw:equation draw:name="f3" draw:formula="left+?f2 "/>
            <draw:equation draw:name="f4" draw:formula="top+?f2 "/>
            <draw:equation draw:name="f5" draw:formula="right-?f2 "/>
            <draw:equation draw:name="f6" draw:formula="bottom-?f2 "/>
            <draw:equation draw:name="f7" draw:formula="left+$0 "/>
            <draw:equation draw:name="f8" draw:formula="top+$0 "/>
            <draw:equation draw:name="f9" draw:formula="bottom-$0 "/>
            <draw:equation draw:name="f10" draw:formula="right-$0 "/>
            <draw:handle draw:handle-position="$0 top"
                   draw:handle-switched="true"
                   draw:handle-range-x-minimum="0"
                   draw:handle-range-x-maximum="10800"/>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- Snip Single Corner Rectangle -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='snip1Rect'">
        <!-- warn if Snip Single Corner Rectangle-->
        <xsl:message terminate="no">translation.oox2odf.shapesTypeSnipSingleCornerRectangle</xsl:message>
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter inserted by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
						draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" 
						draw:text-areas="0 4300 21600 21600" 
						draw:mirror-horizontal="true" draw:type="flowchart-card" 
						draw:enhanced-path="M 4300 0 L 21600 0 21600 21600 0 21600 0 4300 4300 0 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape> 
      </xsl:when>

		<!-- Snip Same Side Corner Rectangle -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='snip2SameRect'">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering -->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 2362200 1066800"
										draw:extrusion-allowed="true"
										draw:text-areas="88902 88902 2273298 1066799"
										draw:glue-points="2362200 533400 1181100 1066800 0 533400 1181100 0"
										draw:enhanced-path="M 177804 0 L 2184396 0 L 2362200 177804 L 2362200 1066800 L 2362200 1066800 L 0 1066800 L 0 1066800 L 0 177804 Z N">
					<xsl:if test="p:spPr/a:prstGeom/@prst='snip2SameRect'">
						<xsl:attribute name ="draw:type">
							<xsl:value-of select="'snip2samerect'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Snip Diagonal Corner Rectangle -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='snip2DiagRect'">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering -->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 2590800 1371600" draw:extrusion-allowed="true"
										draw:text-areas="114302 114302 2476498 1257298"
										draw:glue-points="2590800 685800 1295400 1371600 0 685800 1295400 0"
										draw:enhanced-path="M 0 0 L 2362195 0 L 2590800 228605 L 2590800 1371600 L 2590800 1371600 L 228605 1371600 L 0 1142995 L 0 0 Z N">
					<xsl:if test="p:spPr/a:prstGeom/@prst='snip2DiagRect'">
						<xsl:attribute name ="draw:type">
							<xsl:value-of select="'snip2diagrect'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Snip and Round Single Corner Rectangle -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='snipRoundRect'">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name="CreateShape">
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name ="grID" select ="$GraphicId" />
					<xsl:with-param name ="prID" select ="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering -->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 2514600 1295400" draw:extrusion-allowed="true"
										draw:text-areas="63236 63236 2406647 1295400"
										draw:glue-points="2514600 647700 1257300 1295400 0 647700 1257300 0"
										draw:enhanced-path="M 215904 0 L 2298696 0 L 2514600 215904 L 2514600 1295400 L 0 1295400 L 0 215904 W 0 0 431808 431808 0 215904 215904 0 Z N">
					<xsl:if test="p:spPr/a:prstGeom/@prst='snipRoundRect'">
						<xsl:attribute name ="draw:type">
							<xsl:value-of select="'sniproundrect'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Round Single Corner Rectangle -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='round1Rect'">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
					<!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 2438400 1066800" draw:extrusion-allowed="true"
										draw:text-areas="0 0 2386323 1066800"
										draw:glue-points="1219200 0 0 533400 1219200 1066800 2438400 533400" 
										draw:enhanced-path="M 0 0 L 2260596 0 W 2082792 0 2438400 355608 2260596 0 2438400 177804 L 2438400 1066800 L 0 1066800 Z N">
					<xsl:if test="p:spPr/a:prstGeom/@prst='round1Rect'">
						<xsl:attribute name="draw:type">
							<xsl:value-of select="'round1rect'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Round Same Side Corner Rectangle -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='round2SameRect'">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
					<!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 2362200 1371600" draw:extrusion-allowed="true"
										draw:text-areas="66956 66956 2295244 1371600"
										draw:glue-points="2362200 685800 1181100 1371600 0 685800 1181100 0"
										draw:enhanced-path="M 228605 0 L 2133595 0 W 1904990 0 2362200 457210 2133595 0 2362200 228605 L 2362200 1371600 L 0 1371600 L 0 228605 W 0 0 457210 457210 0 228605 228605 0 Z N">
					<xsl:if test="p:spPr/a:prstGeom/@prst='round2SameRect'">
						<xsl:attribute name="draw:type">
							<xsl:value-of select="'round2samerect'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Round Diagonal Corner Rectangle -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='round2DiagRect'">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
					<!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 2286000 1371600" draw:extrusion-allowed="true"
										draw:text-areas="66956 66956 2219044 1304644"
										draw:glue-points="2286000 685800 1143000 1371600 0 685800 1143000 0"
										draw:enhanced-path="M 228605 0 L 2286000 0 L 2286000 1142995 W 1828790 914390 2286000 1371600 2286000 1142995 2057395 1371600 L 0 1371600 L 0 228605 W 0 0 457210 457210 0 228605 228605 0 Z N">
					<xsl:if test="p:spPr/a:prstGeom/@prst='round2DiagRect'">
						<xsl:attribute name="draw:type">
							<xsl:value-of select="'round2diagrect'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Explosion 2 -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='irregularSeal2'">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
										draw:glue-points="9722 1887 0 12875 11614 18844 21600 6646"
                                        draw:text-areas="5400 6570 14160 15290"
										draw:type="bang"
                                        draw:enhanced-path="M 11464 4340 L 9722 1887 8548 6383 4503 3626 5373 7816 1174 8270 3934 11592 0 12875 3329 15372 1283 17824 4804 18239 4918 21600 7525 18125 8698 19712 9871 17371 11614 18844 12178 15937 14943 17371 14640 14348 18878 15632 16382 12311 18270 11292 16986 9404 21600 6646 16382 6533 18005 3172 14524 5778 14789 0 11464 4340 Z N">
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>

		<!-- Heptagon -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='heptagon'">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 1828800 1600200"
										draw:extrusion-allowed="true"
										draw:text-areas="181107 316941 1647693 1283259"
                                        draw:glue-points="1647693 316941 1828805 1029099 1321344 1600208 507456 1600208 -5 1029099 181107 316941 914400 0"
									    draw:enhanced-path="M -5 1029099 L 181107 316941 L 914400 0 L 1647693 316941 L 1828805 1029099 L 1321344 1600208 L 507456 1600208 Z N">
					<xsl:if test="p:spPr/a:prstGeom/@prst='heptagon'">
						<xsl:attribute name="draw:type">
							<xsl:value-of select="'HEPTAGON'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Decagon -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='decagon'">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 2057400 1752600" draw:extrusion-allowed="true"
										draw:text-areas="196464 334718 1860936 1417882"
                                        draw:glue-points="1860936 334718 2057400 876300 1860936 1417882 1346585 1752598 710815 1752598 196464 1417882 0 876300 196464 334718 710815 2 1346585 2"
                                        draw:enhanced-path="M 0 876300 L 196464 334718 L 710815 2 L 1346585 2 L 1860936 334718 L 2057400 876300 L 1860936 1417882 L 1346585 1752598 L 710815 1752598 L 196464 1417882 Z N">
					<xsl:if test="p:spPr/a:prstGeom/@prst='decagon'">
						<xsl:attribute name="draw:type">
							<xsl:value-of select="'DECAGON'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Dodecagon -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='dodecagon'">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 1905000 1828800" draw:extrusion-allowed="true"
										draw:text-areas="255235 245025 1649765 1583775"
									    draw:glue-points="1649765 245025 1905000 669375 1905000 1159425 1649765 1583775 1207735 1828800 697265 1828800 255235 1583775 0 1159425 0 669375 255235 245025 697265 0 1207735 0"
                                        draw:enhanced-path="M 0 669375 L 255235 245025 L 697265 0 L 1207735 0 L 1649765 245025 L 1905000 669375 L 1905000 1159425 L 1649765 1583775 L 1207735 1828800 L 697265 1828800 L 255235 1583775 L 0 1159425 Z N">
					<xsl:if test="p:spPr/a:prstGeom/@prst='dodecagon'">
						<xsl:attribute name="draw:type">
							<xsl:value-of select="'DODECAGON'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Pie -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='pie'">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry draw:glue-point-type="segments"
                                  draw:type="mso-spt100"
                                  draw:modifiers="-90 0"
                                  draw:enhanced-path="V 0 0 21600 21600 ?f5 ?f7 ?f1 ?f3 L 10800 10800 Z N">
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="10800*cos($0 *(pi/180))"/>
					<draw:equation draw:name="f1" draw:formula="?f0 +10800"/>
					<draw:equation draw:name="f2" draw:formula="10800*sin($0 *(pi/180))"/>
					<draw:equation draw:name="f3" draw:formula="?f2 +10800"/>
					<draw:equation draw:name="f4" draw:formula="10800*cos($1 *(pi/180))"/>
					<draw:equation draw:name="f5" draw:formula="?f4 +10800"/>
					<draw:equation draw:name="f6" draw:formula="10800*sin($1 *(pi/180))"/>
					<draw:equation draw:name="f7" draw:formula="?f6 +10800"/>
					<draw:handle draw:handle-position="10800 $0" draw:handle-polar="10800 10800" draw:handle-radius-range-minimum="10800" draw:handle-radius-range-maximum="10800"/>
					<draw:handle draw:handle-position="10800 $1" draw:handle-polar="10800 10800" draw:handle-radius-range-minimum="10800" draw:handle-radius-range-maximum="10800"/>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Frame -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='frame'">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<!--Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks-->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!--Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
										draw:extrusion-origin="0 0"
										draw:extrusion-viewpoint="(0cm 0cm 25cm)"
                                        draw:extrusion-skew="0 0"
										dr3d:projection="perspective"
										draw:extrusion-depth="2.5cm 0"
										draw:text-areas="?f4 ?f6 ?f5 ?f7"
										draw:glue-point-type="segments"
										draw:path-stretchpoint-x="10800"
										draw:path-stretchpoint-y="10800"
										draw:type="frame"
										draw:modifiers="2000"
										draw:enhanced-path="M ?f0 ?f2 L ?f1 ?f2 ?f1 ?f3 ?f0 ?f3 Z M ?f4 ?f6 L ?f5 ?f6 ?f5 ?f7 ?f4 ?f7 Z N">
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="left"/>
					<draw:equation draw:name="f1" draw:formula="right"/>
					<draw:equation draw:name="f2" draw:formula="top"/>
					<draw:equation draw:name="f3" draw:formula="bottom"/>
					<draw:equation draw:name="f4" draw:formula="left+$0"/>
					<draw:equation draw:name="f5" draw:formula="right-$0"/>
					<draw:equation draw:name="f6" draw:formula="top+$0"/>
					<draw:equation draw:name="f7" draw:formula="bottom-$0"/>
					<draw:handle draw:handle-position="left $0" draw:handle-switched="true" draw:handle-range-y-minimum="0" draw:handle-range-y-maximum="10800"/>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Half Frame -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='halfFrame'">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<!--Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks-->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!--Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 1371600 1371600"
										draw:extrusion-allowed="true"
										draw:text-areas="0 0 1371600 1371600"
										draw:glue-points="1143002 228598 228598 1143002 0 685800 685800 0"
										draw:enhanced-path="M 0 0 L 1371600 0 L 914405 457195 L 457195 457195 L 457195 914405 L 0 1371600 Z N">
					<xsl:if test="p:spPr/a:prstGeom/@prst='halfFrame'">
						<xsl:attribute name="draw:type">
							<xsl:value-of select="'HALFFRAME'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- L-Shape -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='corner'">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<!--Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks-->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!--Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 914400 914400"
							 draw:extrusion-allowed="true"
							 draw:text-areas="0 0 457200 914400"
							 draw:glue-points="914400 685800 457200 914400 0 457200 228600 0"
						     draw:enhanced-path="M 0 0 L 457200 0 L 457200 457200 L 914400 457200 L 914400 914400 L 0 914400 Z N">
					<xsl:if test="p:spPr/a:prstGeom/@prst='corner'">
						<xsl:attribute name="draw:type">
							<xsl:value-of select="'CORNER'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Diagonal Stripe -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='diagStripe'">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<!--Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks-->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!--Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 1524000 1219200"
										draw:extrusion-allowed="true"
                                        draw:text-areas="0 0 1142999 914400"
										draw:glue-points="762000 609600 0 914400 381000 304800 1142999 0"
                                        draw:enhanced-path="M 0 609600 L 762000 0 L 1524000 0 L 0 1219200 Z N">
					<xsl:if test="p:spPr/a:prstGeom/@prst='diagStripe'">
						<xsl:attribute name="draw:type">
							<xsl:value-of select="'diagstripe'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Plaque -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='plaque'">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<!--Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks-->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!--Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
										draw:glue-points="10800 0 0 10800 10800 21600 21600 10800"
                                        draw:path-stretchpoint-x="10800" draw:path-stretchpoint-y="10800"
										draw:text-areas="?f12 ?f13 ?f14 ?f15"
                                        draw:type="mso-spt21"
										draw:modifiers="3600"
                                        draw:enhanced-path="M ?f0 0 Y 0 ?f1 L 0 ?f2 X ?f0 21600 L ?f3 21600 Y 21600 ?f2 L 21600 ?f1 X ?f3 0 Z N">
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="left+$0 "/>
					<draw:equation draw:name="f1" draw:formula="top+$0 "/>
					<draw:equation draw:name="f2" draw:formula="bottom-$0 "/>
					<draw:equation draw:name="f3" draw:formula="right-$0 "/>
					<draw:equation draw:name="f4" draw:formula="-(sin(45*(pi/180))*($0 -10800)-cos(45*(pi/180))*(0-10800))+10800"/>
					<draw:equation draw:name="f5" draw:formula="?f4 -10800"/>
					<draw:equation draw:name="f6" draw:formula="-$0 "/>
					<draw:equation draw:name="f7" draw:formula="?f6 -?f5 "/>
					<draw:equation draw:name="f8" draw:formula="left-?f7 "/>
					<draw:equation draw:name="f9" draw:formula="top-?f7 "/>
					<draw:equation draw:name="f10" draw:formula="right+?f7 "/>
					<draw:equation draw:name="f11" draw:formula="bottom+?f7 "/>
					<draw:equation draw:name="f12" draw:formula="left-?f5 "/>
					<draw:equation draw:name="f13" draw:formula="top-?f5 "/>
					<draw:equation draw:name="f14" draw:formula="right+?f5 "/>
					<draw:equation draw:name="f15" draw:formula="bottom+?f5 "/>
					<draw:handle draw:handle-position="$0 top" draw:handle-switched="true" draw:handle-range-x-minimum="0" draw:handle-range-x-maximum="10800"/>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Bevel -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='bevel'">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<!--Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks-->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!--Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
										draw:path-stretchpoint-x="10800"
										draw:path-stretchpoint-y="10800"
										draw:text-areas="?f0 ?f0 ?f1 ?f2"
										draw:type="quad-bevel"
										draw:modifiers="2700"
										draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 Z N M 0 0 L 21600 0 ?f1 ?f0 ?f0 ?f0 Z N M 21600 0 L 21600 21600 ?f1 ?f2 ?f1 ?f0 Z N M 21600 21600 L 0 21600 ?f0 ?f2 ?f1 ?f2 Z N M 0 21600 L 0 0 ?f0 ?f0 ?f0 ?f2 Z N">
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="$0 *21599/21600"/>
					<draw:equation draw:name="f1" draw:formula="right-?f0 "/>
					<draw:equation draw:name="f2" draw:formula="bottom-?f0 "/>
					<draw:handle draw:handle-position="$0 top" draw:handle-switched="true" draw:handle-range-x-minimum="0" draw:handle-range-x-maximum="10800"/>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Donut -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='donut'">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<!--Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks-->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!--Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
										draw:glue-points="10800 0 3160 3160 0 10800 3160 18440 10800 21600 18440 18440 21600 10800 18440 3160"
										draw:text-areas="3200 3200 18400 18400"
										draw:type="ring"
										draw:modifiers="5400"
										draw:enhanced-path="U 10800 10800 10800 10800 0 23592960 Z U 10800 10800 ?f1 ?f1 0 23592960 N">
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="$0 "/>
					<draw:equation draw:name="f1" draw:formula="10800-$0 "/>
					<draw:handle draw:handle-position="$0 10800" draw:handle-range-x-minimum="0" draw:handle-range-x-maximum="10800"/>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Teardrop -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='teardrop'">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<!--Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks-->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!--Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 2743200 1600200"
										draw:extrusion-allowed="true"
										draw:text-areas="0 0 2743200 1600200"
										draw:glue-points="0 685800 567077 93422 1143001 2 2285998 0 2286000 685800 1718924 1278180 1142999 1371600 567074 1278179 -1 685798 0 685800"
										draw:enhanced-path="M 0 800100 C 1 515531 259112 252377 680492 108992 C 890259 37613 1128753 2 1371601 2 L 2743199 0 C 2743200 266700 2743200 533400 2743200 800100 C 2743200 1084669 2484089 1347823 2062709 1491210 C 1852942 1562589 1614448 1600200 1371599 1600200 C 1128751 1600200 890256 1562588 680489 1491209 C 259109 1347822 -2 1084668 -1 800098 C -1 800099 0 800099 0 800100 Z N">
					<xsl:if test="p:spPr/a:prstGeom/@prst='teardrop'">
						<xsl:attribute name="draw:type">
							<xsl:value-of select="'TEARDROP'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>

      <!-- Flow chart shapes -->

      <!-- Flowchart: Process -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartProcess'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" draw:type="flowchart-process" draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- Flowchart: Alternate Process -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartAlternateProcess'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:path-stretchpoint-x="10800" draw:path-stretchpoint-y="10800" draw:text-areas="?f3 ?f4 ?f5 ?f6" draw:type="round-rectangle" draw:modifiers="3600" draw:enhanced-path="M ?f7 0 X 0 ?f8 L 0 ?f9 Y ?f7 21600 L ?f10 21600 X 21600 ?f9 L 21600 ?f8 Y ?f10 0 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <draw:equation draw:name="f0" draw:formula="45"/>
            <draw:equation draw:name="f1" draw:formula="$0 *sin(?f0 *(pi/180))"/>
            <draw:equation draw:name="f2" draw:formula="?f1 *3163/7636"/>
            <draw:equation draw:name="f3" draw:formula="left+?f2 "/>
            <draw:equation draw:name="f4" draw:formula="top+?f2 "/>
            <draw:equation draw:name="f5" draw:formula="right-?f2 "/>
            <draw:equation draw:name="f6" draw:formula="bottom-?f2 "/>
            <draw:equation draw:name="f7" draw:formula="left+$0 "/>
            <draw:equation draw:name="f8" draw:formula="top+$0 "/>
            <draw:equation draw:name="f9" draw:formula="bottom-$0 "/>
            <draw:equation draw:name="f10" draw:formula="right-$0 "/>
            <draw:handle draw:handle-position="$0 top" draw:handle-switched="true" draw:handle-range-x-minimum="0" draw:handle-range-x-maximum="10800"/>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Decision -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartDecision'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" draw:text-areas="5400 5400 16200 16200" draw:type="flowchart-decision" 
            draw:enhanced-path="M 0 10800 L 10800 0 21600 10800 10800 21600 0 10800 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Data -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartInputOutput'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="12960 0 10800 0 2160 10800 8600 21600 10800 21600 19400 10800" draw:text-areas="4230 0 17370 21600" draw:type="flowchart-data" draw:enhanced-path="M 4230 0 L 21600 0 17370 21600 0 21600 4230 0 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Predefined Process-->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartPredefinedProcess'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:text-areas="2540 0 19060 21600" draw:type="flowchart-predefined-process" draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 Z N M 2540 0 L 2540 21600 N M 19060 0 L 19060 21600 N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Internal Storage -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartInternalStorage'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:text-areas="4230 4230 21600 21600" draw:type="flowchart-internal-storage" draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 Z N M 4230 0 L 4230 21600 N M 0 4230 L 21600 4230 N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Document -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartDocument'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 20320 21600 10800" draw:text-areas="0 0 21600 17360" draw:type="flowchart-document" draw:enhanced-path="M 0 0 L 21600 0 21600 17360 C 13050 17220 13340 20770 5620 21600 2860 21100 1850 20700 0 20120 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Multi document -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartMultidocument'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 19890 21600 10800" draw:text-areas="0 3600 18600 18009" draw:type="flowchart-multidocument" draw:enhanced-path="M 0 3600 L 1500 3600 1500 1800 3000 1800 3000 0 21600 0 21600 14409 20100 14409 20100 16209 18600 16209 18600 18009 C 11610 17893 11472 20839 4833 21528 2450 21113 1591 20781 0 20300 Z N M 1500 3600 F L 18600 3600 18600 16209 N M 3000 1800 F L 20100 1800 20100 14409 N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Terminator -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartTerminator'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" draw:text-areas="1060 3180 20540 18420" draw:type="flowchart-terminator" draw:enhanced-path="M 3470 21600 X 0 10800 3470 0 L 18130 0 X 21600 10800 18130 21600 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Preparation -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartPreparation'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" draw:text-areas="4350 0 17250 21600" draw:type="flowchart-preparation" draw:enhanced-path="M 4350 0 L 17250 0 21600 10800 17250 21600 4350 21600 0 10800 4350 0 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Manual Input -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartManualInput'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 2150 0 10800 10800 19890 21600 10800" draw:text-areas="0 4300 21600 21600" draw:type="flowchart-manual-input" draw:enhanced-path="M 0 4300 L 21600 0 21600 21600 0 21600 0 4300 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Manual Operation -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartManualOperation'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 2160 10800 10800 21600 19440 10800" draw:text-areas="4350 0 17250 21600" draw:type="flowchart-manual-operation" draw:enhanced-path="M 0 0 L 21600 0 17250 21600 4350 21600 0 0 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Connector -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartConnector'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 3160 3160 0 10800 3160 18440 10800 21600 18440 18440 21600 10800 18440 3160" draw:text-areas="3180 3180 18420 18420" draw:type="flowchart-connector" draw:enhanced-path="U 10800 10800 10800 10800 0 23592960 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Off-page Connector -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartOffpageConnector'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" draw:text-areas="0 0 21600 17150" draw:type="flowchart-off-page-connector" draw:enhanced-path="M 0 0 L 21600 0 21600 17150 10800 21600 0 17150 0 0 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Card -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartPunchedCard'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" draw:text-areas="0 4300 21600 21600" draw:type="flowchart-card" draw:enhanced-path="M 4300 0 L 21600 0 21600 21600 0 21600 0 4300 4300 0 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Punched Tape -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartPunchedTape'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 2020 0 10800 10800 19320 21600 10800" draw:text-areas="0 4360 21600 17240" draw:type="flowchart-punched-tape" draw:enhanced-path="M 0 2230 C 820 3990 3410 3980 5370 4360 7430 4030 10110 3890 10690 2270 11440 300 14200 160 16150 0 18670 170 20690 390 21600 2230 L 21600 19420 C 20640 17510 18320 17490 16140 17240 14710 17370 11310 17510 10770 19430 10150 21150 7380 21290 5290 21600 3220 21250 610 21130 0 19420 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Summing Junction -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartSummingJunction'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 3160 3160 0 10800 3160 18440 10800 21600 18440 18440 21600 10800 18440 3160" draw:text-areas="3100 3100 18500 18500" draw:type="flowchart-summing-junction" draw:enhanced-path="U 10800 10800 10800 10800 0 23592960 Z N M 3100 3100 L 18500 18500 N M 3100 18500 L 18500 3100 N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Or -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartOr'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 3160 3160 0 10800 3160 18440 10800 21600 18440 18440 21600 10800 18440 3160" draw:text-areas="3100 3100 18500 18500" draw:type="flowchart-or" draw:enhanced-path="U 10800 10800 10800 10800 0 23592960 Z N M 0 10800 L 21600 10800 N M 10800 0 L 10800 21600 N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Collate -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartCollate'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 10800 10800 10800 21600" draw:text-areas="5400 5400 16200 16200" draw:type="flowchart-collate" draw:enhanced-path="M 0 0 L 21600 21600 0 21600 21600 0 0 0 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Sort -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartSort'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:text-areas="5400 5400 16200 16200" draw:type="flowchart-sort" draw:enhanced-path="M 0 10800 L 10800 0 21600 10800 10800 21600 Z N M 0 10800 L 21600 10800 N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Extract -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartExtract'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 5400 10800 10800 21600 16200 10800" draw:text-areas="5400 10800 16200 21600" draw:type="flowchart-extract" draw:enhanced-path="M 10800 0 L 21600 21600 0 21600 10800 0 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Merge-->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartMerge'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 5400 10800 10800 21600 16200 10800" draw:text-areas="5400 0 16200 10800" draw:type="flowchart-merge" draw:enhanced-path="M 0 0 L 21600 0 10800 21600 0 0 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Stored Data -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartOnlineStorage'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 18000 10800" draw:text-areas="3600 0 18000 21600" draw:type="flowchart-stored-data" draw:enhanced-path="M 3600 21600 X 0 10800 3600 0 L 21600 0 X 18000 10800 21600 21600 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Delay -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartDelay'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" draw:text-areas="0 3100 18500 18500" draw:type="flowchart-delay" draw:enhanced-path="M 10800 0 X 21600 10800 10800 21600 L 0 21600 0 0 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Sequential Access Storage -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartMagneticTape'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" draw:text-areas="3100 3100 18500 18500" draw:type="flowchart-sequential-access" draw:enhanced-path="M 20980 18150 L 20980 21600 10670 21600 C 4770 21540 0 16720 0 10800 0 4840 4840 0 10800 0 16740 0 21600 4840 21600 10800 21600 13520 20550 16160 18670 18170 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Direct Access Storage -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartMagneticDrum'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 14800 10800 21600 10800" draw:text-areas="3400 0 14800 21600" draw:type="flowchart-direct-access-storage" draw:enhanced-path="M 18200 0 X 21600 10800 18200 21600 L 3400 21600 X 0 10800 3400 0 Z N M 18200 0 X 14800 10800 18200 21600 N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Magnetic Disk-->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartMagneticDisk'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 6800 10800 0 0 10800 10800 21600 21600 10800" draw:text-areas="0 6800 21600 18200" draw:type="flowchart-magnetic-disk" draw:enhanced-path="M 0 3400 Y 10800 0 21600 3400 L 21600 18200 Y 10800 21600 0 18200 Z N M 0 3400 Y 10800 6800 21600 3400 N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>
      <!-- FlowChart: Display-->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartDisplay'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="sldId" select="$slideId" />
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
            <xsl:with-param name="grpBln" select ="$grpBln" />
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
            <!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" draw:text-areas="3600 0 17800 21600" draw:type="flowchart-display" draw:enhanced-path="M 3600 0 L 17800 0 X 21600 10800 17800 21600 L 3600 21600 0 10800 Z N">
            <xsl:if test="p:spPr/a:xfrm/@flipH='1'">
              <xsl:attribute name ="draw:mirror-horizontal">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="p:spPr/a:xfrm/@flipV='1'">
              <xsl:attribute name ="draw:mirror-vertical">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
      </xsl:when>

		<!-- Action Buttons -->
		<!-- Action Buttons Back or Previous -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='actionButtonBackPrevious'">
			<draw:custom-shape draw:layer="layout" >
				<xsl:call-template name ="CreateShape">
					<!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
										draw:path-stretchpoint-x="10800"
										draw:path-stretchpoint-y="10800"
										draw:text-areas="?f1 ?f2 ?f3 ?f4"
										draw:type="mso-spt194"
										draw:modifiers="1400"
										draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 Z N M 0 0 L 21600 0 ?f3 ?f2 ?f1 ?f2 Z N M 21600 0 L 21600 21600 ?f3 ?f4 ?f3 ?f2 Z N M 21600 21600 L 0 21600 ?f1 ?f4 ?f3 ?f4 Z N M 0 21600 L 0 0 ?f1 ?f2 ?f1 ?f4 Z N M ?f10 ?f8 L ?f14 ?f12 ?f14 ?f16 Z N">
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="$0" />
					<draw:equation draw:name="f1" draw:formula="left+$0" />
					<draw:equation draw:name="f2" draw:formula="top+$0" />
					<draw:equation draw:name="f3" draw:formula="right-$0" />
					<draw:equation draw:name="f4" draw:formula="bottom-$0" />
					<draw:equation draw:name="f5" draw:formula="10800-$0" />
					<draw:equation draw:name="f6" draw:formula="?f5 /10800" />
					<draw:equation draw:name="f7" draw:formula="right/2" />
					<draw:equation draw:name="f8" draw:formula="bottom/2" />
					<draw:equation draw:name="f9" draw:formula="-8050*?f6" />
					<draw:equation draw:name="f10" draw:formula="?f9 +?f7" />
					<draw:equation draw:name="f11" draw:formula="-8050*?f6" />
					<draw:equation draw:name="f12" draw:formula="?f11 +?f8" />
					<draw:equation draw:name="f13" draw:formula="8050*?f6" />
					<draw:equation draw:name="f14" draw:formula="?f13 +?f7" />
					<draw:equation draw:name="f15" draw:formula="8050*?f6" />
					<draw:equation draw:name="f16" draw:formula="?f15 +?f8" />
					<draw:handle draw:handle-position="$0 top" draw:handle-switched="true" draw:handle-range-x-minimum="0" draw:handle-range-x-maximum="5400" />
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape >
		</xsl:when>
		<!-- Action Buttons Forward or Next -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='actionButtonForwardNext'">
			<draw:custom-shape draw:layer="layout">
				<xsl:call-template name ="CreateShape">
					<!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
										draw:path-stretchpoint-x="10800"
										draw:path-stretchpoint-y="10800"
										draw:text-areas="?f1 ?f2 ?f3 ?f4"
										draw:type="mso-spt193"
										draw:modifiers="1400"
										draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 Z N M 0 0 L 21600 0 ?f3 ?f2 ?f1 ?f2 Z N M 21600 0 L 21600 21600 ?f3 ?f4 ?f3 ?f2 Z N M 21600 21600 L 0 21600 ?f1 ?f4 ?f3 ?f4 Z N M 0 21600 L 0 0 ?f1 ?f2 ?f1 ?f4 Z N M ?f10 ?f12 L ?f14 ?f8 ?f10 ?f16 Z N">
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="$0" />
					<draw:equation draw:name="f1" draw:formula="left+$0" />
					<draw:equation draw:name="f2" draw:formula="top+$0" />
					<draw:equation draw:name="f3" draw:formula="right-$0" />
					<draw:equation draw:name="f4" draw:formula="bottom-$0" />
					<draw:equation draw:name="f5" draw:formula="10800-$0" />
					<draw:equation draw:name="f6" draw:formula="?f5 /10800" />
					<draw:equation draw:name="f7" draw:formula="right/2" />
					<draw:equation draw:name="f8" draw:formula="bottom/2" />
					<draw:equation draw:name="f9" draw:formula="-8050*?f6" />
					<draw:equation draw:name="f10" draw:formula="?f9 +?f7" />
					<draw:equation draw:name="f11" draw:formula="-8050*?f6" />
					<draw:equation draw:name="f12" draw:formula="?f11 +?f8" />
					<draw:equation draw:name="f13" draw:formula="8050*?f6" />
					<draw:equation draw:name="f14" draw:formula="?f13 +?f7" />
					<draw:equation draw:name="f15" draw:formula="8050*?f6" />
					<draw:equation draw:name="f16" draw:formula="?f15 +?f8" />
					<draw:handle draw:handle-position="$0 top" draw:handle-switched="true" draw:handle-range-x-minimum="0" draw:handle-range-x-maximum="5400" />
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape >
		</xsl:when>
		<!-- Action Buttons Beginning -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='actionButtonBeginning'">
			<draw:custom-shape draw:layer="layout">
				<xsl:call-template name ="CreateShape">
					<!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
										draw:path-stretchpoint-x="10800"
										draw:path-stretchpoint-y="10800"
										draw:text-areas="?f1 ?f2 ?f3 ?f4"
										draw:type="mso-spt196"
										draw:modifiers="1400"
										draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 Z N M 0 0 L 21600 0 ?f3 ?f2 ?f1 ?f2 Z N M 21600 0 L 21600 21600 ?f3 ?f4 ?f3 ?f2 Z N M 21600 21600 L 0 21600 ?f1 ?f4 ?f3 ?f4 Z N M 0 21600 L 0 0 ?f1 ?f2 ?f1 ?f4 Z N M ?f10 ?f8 L ?f14 ?f12 ?f14 ?f16 Z N M ?f18 ?f12 L ?f20 ?f12 ?f20 ?f16 ?f18 ?f16 Z N">
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="$0" />
					<draw:equation draw:name="f1" draw:formula="left+$0" />
					<draw:equation draw:name="f2" draw:formula="top+$0" />
					<draw:equation draw:name="f3" draw:formula="right-$0" />
					<draw:equation draw:name="f4" draw:formula="bottom-$0" />
					<draw:equation draw:name="f5" draw:formula="10800-$0" />
					<draw:equation draw:name="f6" draw:formula="?f5 /10800" />
					<draw:equation draw:name="f7" draw:formula="right/2" />
					<draw:equation draw:name="f8" draw:formula="bottom/2" />
					<draw:equation draw:name="f9" draw:formula="-4020*?f6" />
					<draw:equation draw:name="f10" draw:formula="?f9 +?f7" />
					<draw:equation draw:name="f11" draw:formula="-8050*?f6" />
					<draw:equation draw:name="f12" draw:formula="?f11 +?f8" />
					<draw:equation draw:name="f13" draw:formula="8050*?f6" />
					<draw:equation draw:name="f14" draw:formula="?f13 +?f7" />
					<draw:equation draw:name="f15" draw:formula="8050*?f6" />
					<draw:equation draw:name="f16" draw:formula="?f15 +?f8" />
					<draw:equation draw:name="f17" draw:formula="-8050*?f6" />
					<draw:equation draw:name="f18" draw:formula="?f17 +?f7" />
					<draw:equation draw:name="f19" draw:formula="-6140*?f6" />
					<draw:equation draw:name="f20" draw:formula="?f19 +?f7" />
					<draw:equation draw:name="f21" draw:formula="4020*?f6" />
					<draw:equation draw:name="f22" draw:formula="?f21 +?f7" />
					<draw:equation draw:name="f23" draw:formula="6140*?f6" />
					<draw:equation draw:name="f24" draw:formula="?f23 +?f7" />
					<draw:handle draw:handle-position="$0 top" draw:handle-switched="true" draw:handle-range-x-minimum="0" draw:handle-range-x-maximum="5400" />
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape >
		</xsl:when>
		<!-- Action Buttons end -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='actionButtonEnd'">
			<draw:custom-shape draw:layer="layout">
				<xsl:call-template name ="CreateShape">
					<!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
										draw:path-stretchpoint-x="10800"
										draw:path-stretchpoint-y="10800"
										draw:text-areas="?f1 ?f2 ?f3 ?f4"
										draw:type="mso-spt195"
										draw:modifiers="1400"
										draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 Z N M 0 0 L 21600 0 ?f3 ?f2 ?f1 ?f2 Z N M 21600 0 L 21600 21600 ?f3 ?f4 ?f3 ?f2 Z N M 21600 21600 L 0 21600 ?f1 ?f4 ?f3 ?f4 Z N M 0 21600 L 0 0 ?f1 ?f2 ?f1 ?f4 Z N M ?f22 ?f8 L ?f18 ?f16 ?f18 ?f12 Z N M ?f24 ?f12 L ?f24 ?f16 ?f14 ?f16 ?f14 ?f12 Z N">
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="$0" />
					<draw:equation draw:name="f1" draw:formula="left+$0" />
					<draw:equation draw:name="f2" draw:formula="top+$0" />
					<draw:equation draw:name="f3" draw:formula="right-$0" />
					<draw:equation draw:name="f4" draw:formula="bottom-$0" />
					<draw:equation draw:name="f5" draw:formula="10800-$0" />
					<draw:equation draw:name="f6" draw:formula="?f5 /10800" />
					<draw:equation draw:name="f7" draw:formula="right/2" />
					<draw:equation draw:name="f8" draw:formula="bottom/2" />
					<draw:equation draw:name="f9" draw:formula="-4020*?f6" />
					<draw:equation draw:name="f10" draw:formula="?f9 +?f7" />
					<draw:equation draw:name="f11" draw:formula="-8050*?f6" />
					<draw:equation draw:name="f12" draw:formula="?f11 +?f8" />
					<draw:equation draw:name="f13" draw:formula="8050*?f6" />
					<draw:equation draw:name="f14" draw:formula="?f13 +?f7" />
					<draw:equation draw:name="f15" draw:formula="8050*?f6" />
					<draw:equation draw:name="f16" draw:formula="?f15 +?f8" />
					<draw:equation draw:name="f17" draw:formula="-8050*?f6" />
					<draw:equation draw:name="f18" draw:formula="?f17 +?f7" />
					<draw:equation draw:name="f19" draw:formula="-6140*?f6" />
					<draw:equation draw:name="f20" draw:formula="?f19 +?f7" />
					<draw:equation draw:name="f21" draw:formula="4020*?f6" />
					<draw:equation draw:name="f22" draw:formula="?f21 +?f7" />
					<draw:equation draw:name="f23" draw:formula="6140*?f6" />
					<draw:equation draw:name="f24" draw:formula="?f23 +?f7" />
					<draw:handle draw:handle-position="$0 top" draw:handle-switched="true" draw:handle-range-x-minimum="0" draw:handle-range-x-maximum="5400" />
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape >
		</xsl:when>
		<!-- Action Buttons Home -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='actionButtonHome'">
			<draw:custom-shape draw:layer="layout">
				<xsl:call-template name ="CreateShape">
					<!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!-- Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks -->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
										draw:path-stretchpoint-x="10800"
										draw:path-stretchpoint-y="10800"
										draw:text-areas="?f1 ?f2 ?f3 ?f4"
										draw:type="mso-spt190"
										draw:modifiers="1400"
										draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 Z N M 0 0 L 21600 0 ?f3 ?f2 ?f1 ?f2 Z N M 21600 0 L 21600 21600 ?f3 ?f4 ?f3 ?f2 Z N M 21600 21600 L 0 21600 ?f1 ?f4 ?f3 ?f4 Z N M 0 21600 L 0 0 ?f1 ?f2 ?f1 ?f4 Z N M ?f7 ?f10 L ?f12 ?f14 ?f12 ?f16 ?f18 ?f16 ?f18 ?f20 ?f22 ?f8 ?f24 ?f8 ?f24 ?f26 ?f28 ?f26 ?f28 ?f8 ?f30 ?f8 Z N M ?f12 ?f14 L ?f12 ?f16 ?f18 ?f16 ?f18 ?f20 Z N M ?f32 ?f36 L ?f34 ?f36 ?f34 ?f26 ?f24 ?f26 ?f24 ?f8 ?f28 ?f8 ?f28 ?f26 ?f32 ?f26 Z N">
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="$0" />
					<draw:equation draw:name="f1" draw:formula="left+$0" />
					<draw:equation draw:name="f2" draw:formula="top+$0" />
					<draw:equation draw:name="f3" draw:formula="right-$0" />
					<draw:equation draw:name="f4" draw:formula="bottom-$0" />
					<draw:equation draw:name="f5" draw:formula="10800-$0" />
					<draw:equation draw:name="f6" draw:formula="?f5 /10800" />
					<draw:equation draw:name="f7" draw:formula="right/2" />
					<draw:equation draw:name="f8" draw:formula="bottom/2" />
					<draw:equation draw:name="f9" draw:formula="-8000*?f6" />
					<draw:equation draw:name="f10" draw:formula="?f9 +?f8" />
					<draw:equation draw:name="f11" draw:formula="2960*?f6" />
					<draw:equation draw:name="f12" draw:formula="?f11 +?f7" />
					<draw:equation draw:name="f13" draw:formula="-5000*?f6" />
					<draw:equation draw:name="f14" draw:formula="?f13 +?f8" />
					<draw:equation draw:name="f15" draw:formula="-7000*?f6" />
					<draw:equation draw:name="f16" draw:formula="?f15 +?f8" />
					<draw:equation draw:name="f17" draw:formula="5000*?f6" />
					<draw:equation draw:name="f18" draw:formula="?f17 +?f7" />
					<draw:equation draw:name="f19" draw:formula="-2960*?f6" />
					<draw:equation draw:name="f20" draw:formula="?f19 +?f8" />
					<draw:equation draw:name="f21" draw:formula="8000*?f6" />
					<draw:equation draw:name="f22" draw:formula="?f21 +?f7" />
					<draw:equation draw:name="f23" draw:formula="6100*?f6" />
					<draw:equation draw:name="f24" draw:formula="?f23 +?f7" />
					<draw:equation draw:name="f25" draw:formula="8260*?f6" />
					<draw:equation draw:name="f26" draw:formula="?f25 +?f8" />
					<draw:equation draw:name="f27" draw:formula="-6100*?f6" />
					<draw:equation draw:name="f28" draw:formula="?f27 +?f7" />
					<draw:equation draw:name="f29" draw:formula="-8000*?f6" />
					<draw:equation draw:name="f30" draw:formula="?f29 +?f7" />
					<draw:equation draw:name="f31" draw:formula="-1060*?f6" />
					<draw:equation draw:name="f32" draw:formula="?f31 +?f7" />
					<draw:equation draw:name="f33" draw:formula="1060*?f6" />
					<draw:equation draw:name="f34" draw:formula="?f33 +?f7" />
					<draw:equation draw:name="f35" draw:formula="4020*?f6" />
					<draw:equation draw:name="f36" draw:formula="?f35 +?f8" />
					<draw:handle draw:handle-position="$0 top" draw:handle-switched="true" draw:handle-range-x-minimum="0" draw:handle-range-x-maximum="5400" />
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape >
		</xsl:when>
		<!-- Action Buttons Information -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='actionButtonInformation'">
			<draw:custom-shape draw:layer="layout">
				<xsl:call-template name ="CreateShape">
					<!--Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks-->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!--Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
											draw:path-stretchpoint-x="10800"
											draw:path-stretchpoint-y="10800"
											draw:text-areas="?f1 ?f2 ?f3 ?f4"
											draw:type="mso-spt192"
											draw:modifiers="1400"
											draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 Z N M 0 0 L 21600 0 ?f3 ?f2 ?f1 ?f2 Z N M 21600 0 L 21600 21600 ?f3 ?f4 ?f3 ?f2 Z N M 21600 21600 L 0 21600 ?f1 ?f4 ?f3 ?f4 Z N M 0 21600 L 0 0 ?f1 ?f2 ?f1 ?f4 Z N M ?f7 ?f12 X ?f10 ?f8 ?f7 ?f16 ?f14 ?f8 ?f7 ?f12 Z N M ?f7 ?f20 X ?f18 ?f42 ?f7 ?f24 ?f22 ?f42 ?f7 ?f20 Z N M ?f26 ?f28 L ?f30 ?f28 ?f30 ?f32 ?f34 ?f32 ?f34 ?f36 ?f26 ?f36 ?f26 ?f32 ?f38 ?f32 ?f38 ?f40 ?f26 ?f40 Z N">
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="$0 "/>
					<draw:equation draw:name="f1" draw:formula="left+$0 "/>
					<draw:equation draw:name="f2" draw:formula="top+$0 "/>
					<draw:equation draw:name="f3" draw:formula="right-$0 "/>
					<draw:equation draw:name="f4" draw:formula="bottom-$0 "/>
					<draw:equation draw:name="f5" draw:formula="10800-$0 "/>
					<draw:equation draw:name="f6" draw:formula="?f5 /10800"/>
					<draw:equation draw:name="f7" draw:formula="right/2"/>
					<draw:equation draw:name="f8" draw:formula="bottom/2"/>
					<draw:equation draw:name="f9" draw:formula="-8050*?f6 "/>
					<draw:equation draw:name="f10" draw:formula="?f9 +?f7 "/>
					<draw:equation draw:name="f11" draw:formula="-8050*?f6 "/>
					<draw:equation draw:name="f12" draw:formula="?f11 +?f8 "/>
					<draw:equation draw:name="f13" draw:formula="8050*?f6 "/>
					<draw:equation draw:name="f14" draw:formula="?f13 +?f7 "/>
					<draw:equation draw:name="f15" draw:formula="8050*?f6 "/>
					<draw:equation draw:name="f16" draw:formula="?f15 +?f8 "/>
					<draw:equation draw:name="f17" draw:formula="-2060*?f6 "/>
					<draw:equation draw:name="f18" draw:formula="?f17 +?f7 "/>
					<draw:equation draw:name="f19" draw:formula="-7620*?f6 "/>
					<draw:equation draw:name="f20" draw:formula="?f19 +?f8 "/>
					<draw:equation draw:name="f21" draw:formula="2060*?f6 "/>
					<draw:equation draw:name="f22" draw:formula="?f21 +?f7 "/>
					<draw:equation draw:name="f23" draw:formula="-3500*?f6 "/>
					<draw:equation draw:name="f24" draw:formula="?f23 +?f8 "/>
					<draw:equation draw:name="f25" draw:formula="-2960*?f6 "/>
					<draw:equation draw:name="f26" draw:formula="?f25 +?f7 "/>
					<draw:equation draw:name="f27" draw:formula="-2960*?f6 "/>
					<draw:equation draw:name="f28" draw:formula="?f27 +?f8 "/>
					<draw:equation draw:name="f29" draw:formula="1480*?f6 "/>
					<draw:equation draw:name="f30" draw:formula="?f29 +?f7 "/>
					<draw:equation draw:name="f31" draw:formula="5080*?f6 "/>
					<draw:equation draw:name="f32" draw:formula="?f31 +?f8 "/>
					<draw:equation draw:name="f33" draw:formula="2960*?f6 "/>
					<draw:equation draw:name="f34" draw:formula="?f33 +?f7 "/>
					<draw:equation draw:name="f35" draw:formula="6140*?f6 "/>
					<draw:equation draw:name="f36" draw:formula="?f35 +?f8 "/>
					<draw:equation draw:name="f37" draw:formula="-1480*?f6 "/>
					<draw:equation draw:name="f38" draw:formula="?f37 +?f7 "/>
					<draw:equation draw:name="f39" draw:formula="-1920*?f6 "/>
					<draw:equation draw:name="f40" draw:formula="?f39 +?f8 "/>
					<draw:equation draw:name="f41" draw:formula="-5560*?f6 "/>
					<draw:equation draw:name="f42" draw:formula="?f41 +?f8 "/>
					<draw:handle draw:handle-position="$0 top" draw:handle-switched="true" draw:handle-range-x-minimum="0" draw:handle-range-x-maximum="5400"/>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>
		<!-- Action Buttons Return -->
		<xsl:when test ="p:spPr/a:prstGeom/@prst='actionButtonReturn'">
			<draw:custom-shape draw:layer="layout">
				<xsl:call-template name ="CreateShape">
					<!--Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks-->
					<xsl:with-param name="sldId" select="$slideId" />
					<xsl:with-param name="grID" select ="$GraphicId"/>
					<xsl:with-param name ="prID" select="$ParaId" />
					<xsl:with-param name="TypeId" select ="$TypeId" />
					<xsl:with-param name="grpBln" select ="$grpBln" />
					<xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
					<!--Extra parameter "SlideRelationId" added by lohith,requierd for template AddTextHyperlinks-->
					<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
										draw:path-stretchpoint-x="10800"
										draw:path-stretchpoint-y="10800"
										draw:text-areas="?f1 ?f2 ?f3 ?f4"
										draw:type="mso-spt197"
										draw:modifiers="1400"
										draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 Z N M 0 0 L 21600 0 ?f3 ?f2 ?f1 ?f2 Z N M 21600 0 L 21600 21600 ?f3 ?f4 ?f3 ?f2 Z N M 21600 21600 L 0 21600 ?f1 ?f4 ?f3 ?f4 Z N M 0 21600 L 0 0 ?f1 ?f2 ?f1 ?f4 Z N M ?f10 ?f12 L ?f14 ?f12 ?f14 ?f16 C ?f14 ?f18 ?f20 ?f22 ?f24 ?f22 L ?f7 ?f22 C ?f26 ?f22 ?f28 ?f18 ?f28 ?f16 L ?f28 ?f12 ?f7 ?f12 ?f30 ?f32 ?f34 ?f12 ?f36 ?f12 ?f36 ?f16 C ?f36 ?f38 ?f40 ?f42 ?f7 ?f42 L ?f24 ?f42 C ?f44 ?f42 ?f10 ?f38 ?f10 ?f16 Z N">
					<xsl:if test="p:spPr/a:xfrm/@flipH='1'">
						<xsl:attribute name ="draw:mirror-horizontal">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="p:spPr/a:xfrm/@flipV='1'">
						<xsl:attribute name ="draw:mirror-vertical">
							<xsl:value-of select="'true'"/>
						</xsl:attribute>
					</xsl:if>
					<draw:equation draw:name="f0" draw:formula="$0 "/>
					<draw:equation draw:name="f1" draw:formula="left+$0 "/>
					<draw:equation draw:name="f2" draw:formula="top+$0 "/>
					<draw:equation draw:name="f3" draw:formula="right-$0 "/>
					<draw:equation draw:name="f4" draw:formula="bottom-$0 "/>
					<draw:equation draw:name="f5" draw:formula="10800-$0 "/>
					<draw:equation draw:name="f6" draw:formula="?f5 /10800"/>
					<draw:equation draw:name="f7" draw:formula="right/2"/>
					<draw:equation draw:name="f8" draw:formula="bottom/2"/>
					<draw:equation draw:name="f9" draw:formula="-8050*?f6 "/>
					<draw:equation draw:name="f10" draw:formula="?f9 +?f7 "/>
					<draw:equation draw:name="f11" draw:formula="-3800*?f6 "/>
					<draw:equation draw:name="f12" draw:formula="?f11 +?f8 "/>
					<draw:equation draw:name="f13" draw:formula="-4020*?f6 "/>
					<draw:equation draw:name="f14" draw:formula="?f13 +?f7 "/>
					<draw:equation draw:name="f15" draw:formula="2330*?f6 "/>
					<draw:equation draw:name="f16" draw:formula="?f15 +?f8 "/>
					<draw:equation draw:name="f17" draw:formula="3390*?f6 "/>
					<draw:equation draw:name="f18" draw:formula="?f17 +?f8 "/>
					<draw:equation draw:name="f19" draw:formula="-3100*?f6 "/>
					<draw:equation draw:name="f20" draw:formula="?f19 +?f7 "/>
					<draw:equation draw:name="f21" draw:formula="4230*?f6 "/>
					<draw:equation draw:name="f22" draw:formula="?f21 +?f8 "/>
					<draw:equation draw:name="f23" draw:formula="-1910*?f6 "/>
					<draw:equation draw:name="f24" draw:formula="?f23 +?f7 "/>
					<draw:equation draw:name="f25" draw:formula="1190*?f6 "/>
					<draw:equation draw:name="f26" draw:formula="?f25 +?f7 "/>
					<draw:equation draw:name="f27" draw:formula="2110*?f6 "/>
					<draw:equation draw:name="f28" draw:formula="?f27 +?f7 "/>
					<draw:equation draw:name="f29" draw:formula="4030*?f6 "/>
					<draw:equation draw:name="f30" draw:formula="?f29 +?f7 "/>
					<draw:equation draw:name="f31" draw:formula="-7830*?f6 "/>
					<draw:equation draw:name="f32" draw:formula="?f31 +?f8 "/>
					<draw:equation draw:name="f33" draw:formula="8250*?f6 "/>
					<draw:equation draw:name="f34" draw:formula="?f33 +?f7 "/>
					<draw:equation draw:name="f35" draw:formula="6140*?f6 "/>
					<draw:equation draw:name="f36" draw:formula="?f35 +?f7 "/>
					<draw:equation draw:name="f37" draw:formula="5510*?f6 "/>
					<draw:equation draw:name="f38" draw:formula="?f37 +?f8 "/>
					<draw:equation draw:name="f39" draw:formula="3180*?f6 "/>
					<draw:equation draw:name="f40" draw:formula="?f39 +?f7 "/>
					<draw:equation draw:name="f41" draw:formula="8450*?f6 "/>
					<draw:equation draw:name="f42" draw:formula="?f41 +?f8 "/>
					<draw:equation draw:name="f43" draw:formula="-5090*?f6 "/>
					<draw:equation draw:name="f44" draw:formula="?f43 +?f7 "/>
					<draw:handle draw:handle-position="$0 top" draw:handle-switched="true" draw:handle-range-x-minimum="0" draw:handle-range-x-maximum="5400"/>
				</draw:enhanced-geometry>
				<xsl:copy-of select="$varHyperLinksForShapes" />
			</draw:custom-shape>
		</xsl:when>

    </xsl:choose>
  </xsl:template>
  <!-- Draw Shape reading values from pptx p:spPr-->
  <xsl:template name ="CreateShape">
    <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
    <xsl:param name ="sldId" />
    <xsl:param name ="grID" />
    <xsl:param name ="prID" />
    <xsl:param name ="TypeId" />
    <xsl:param name ="grpBln" />
    <xsl:param name ="flagTextBox" />
    <xsl:param name ="grpCordinates"/>
    <!-- Addition of a parameter,by Vijayets ,for bullets and numbering in shapes-->
    <xsl:param name="SlideRelationId"/>

    <xsl:attribute name ="draw:style-name">
      <xsl:value-of select ="$grID"/>
    </xsl:attribute>
    <xsl:attribute name ="draw:text-style-name">
      <xsl:value-of select ="$prID"/>
    </xsl:attribute>
	  <!-- animation Id-->
	  <xsl:attribute name ="draw:id">
		  <xsl:value-of select ="concat('sldraw',$sldId,'an',p:nvSpPr/p:cNvPr/@id)"/>
	  </xsl:attribute>
    <!-- For the Grouping of Shapes Bug Fixing -->
    <xsl:choose>
      <xsl:when test="$grpBln ='true'">
        <xsl:call-template name="tmpGropingWriteCordinates">
          <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="tmpWriteCordinates"/>
      </xsl:otherwise>
      
    </xsl:choose>

    <xsl:choose>
      <xsl:when test ="$flagTextBox='true'">
        <draw:text-box>
          <xsl:call-template name ="AddShapeText">
            <xsl:with-param name ="prID" select ="$prID" />
            <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
            <xsl:with-param name ="sldID" select ="$sldId" />
            <!-- Addition of a parameter,by vijayeta,for bullets and numbering in shapes-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <xsl:with-param name="TypeId" select ="$TypeId" />
          </xsl:call-template>
        </draw:text-box>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name ="AddShapeText">
          <xsl:with-param name ="prID" select ="$prID" />
          <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
          <xsl:with-param name ="sldID" select ="$sldId" />
          <!-- Addition of a parameter,by vijayeta,for bullets and numbering in shapes-->
          <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
          <xsl:with-param name="TypeId" select ="$TypeId" />
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!-- Draw line -->
  <xsl:template name ="DrawLine">
    <xsl:param name ="grID" />
    <xsl:param name ="grpBln" />
    <xsl:param name="sldId" />
    <xsl:param name ="grpCordinates" />
    <xsl:attribute name ="draw:style-name">
      <xsl:value-of select ="$grID"/>
    </xsl:attribute>
	  <xsl:attribute name ="draw:id">
		  <xsl:value-of select ="concat('sldraw',$sldId,'an',p:nvCxnSpPr/p:cNvPr/@id)"/>
	  </xsl:attribute>
    <xsl:choose>
      <xsl:when test="$grpBln ='true'">
        <xsl:call-template name="tmpGropingWriteCordinates">
            <xsl:with-param name ="grpCordinates" select ="$grpCordinates" />
          <xsl:with-param name ="ShapeType" select ="'Line'" />
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
    <xsl:for-each select ="p:spPr/a:xfrm">
      <xsl:variable name ="xCenter">
        <xsl:value-of select ="a:off/@x + (a:ext/@cx div 2)"/>
      </xsl:variable>
      <xsl:variable name ="yCenter">
        <xsl:value-of select ="a:off/@y + (a:ext/@cy div 2)"/>
      </xsl:variable>
      <xsl:variable name ="angle">
        <xsl:if test ="not(@rot)">
          <xsl:value-of select="0" />
        </xsl:if>
        <xsl:if test ="@rot">
          <xsl:value-of select ="(@rot div 60000) * ((22 div 7) div 180)"/>
        </xsl:if>
      </xsl:variable>
      <xsl:variable name ="cxBy2">
        <xsl:if test ="(@flipH = 1)">
          <xsl:value-of select ="(-1 * a:ext/@cx) div 2"/>
        </xsl:if>
        <xsl:if test ="not(@flipH) or (@flipH != 1) ">
          <xsl:value-of select ="a:ext/@cx div 2"/>
        </xsl:if>
      </xsl:variable>
      <xsl:variable name ="cyBy2">
        <xsl:if test ="(@flipV = 1)">
          <xsl:value-of select ="(-1 * a:ext/@cy) div 2"/>
        </xsl:if>
        <xsl:if test ="not(@flipV) or (@flipV != 1)">
          <xsl:value-of select ="a:ext/@cy div 2"/>
        </xsl:if>
      </xsl:variable>
      <xsl:attribute name ="svg:x1">
        <xsl:value-of select ="concat('svg-x1:',$xCenter, ':', $cxBy2, ':', $cyBy2, ':', $angle)"/>
      </xsl:attribute>
      <xsl:attribute name ="svg:y1">
        <xsl:value-of select ="concat('svg-y1:',$yCenter, ':', $cxBy2, ':', $cyBy2, ':', $angle)"/>
      </xsl:attribute>
      <xsl:attribute name ="svg:x2">
        <xsl:value-of select ="concat('svg-x2:',$xCenter, ':', $cxBy2, ':', $cyBy2, ':', $angle)"/>
      </xsl:attribute>
      <xsl:attribute name ="svg:y2">
        <xsl:value-of select ="concat('svg-y2:',$yCenter, ':', $cxBy2, ':', $cyBy2, ':', $angle)"/>
      </xsl:attribute>
         </xsl:for-each>
      </xsl:otherwise>

    </xsl:choose>
 
	  <xsl:if test="(p:nvCxnSpPr/p:cNvCxnSpPr/a:stCxn)">

		  <xsl:variable name="pNvprId">
			  <xsl:value-of select="p:nvCxnSpPr/p:cNvCxnSpPr/a:stCxn/@id" />
		  </xsl:variable>
		  <xsl:attribute name ="draw:start-shape">
			  <xsl:value-of select ="concat('sl',$sldId,'an', $pNvprId)"/>
		  </xsl:attribute>
		  <xsl:variable name="stGluePoint">
			  <xsl:value-of select="p:nvCxnSpPr/p:cNvCxnSpPr/a:stCxn/@idx"/>
		  </xsl:variable>

		  <xsl:attribute name ="draw:start-glue-point">
			  	  <xsl:choose>
					  <xsl:when test="$stGluePoint = 0">
						  <xsl:value-of select="'0'"/>
					  </xsl:when>
					  <xsl:when test="$stGluePoint = 1">
						  <xsl:value-of select="'3'"/>
					  </xsl:when>
					  <xsl:when test="$stGluePoint = 2">
						  <xsl:value-of select="'2'"/>
					  </xsl:when>
					  <xsl:when test="$stGluePoint = 3">
						  <xsl:value-of select="'1'"/>
					  </xsl:when>
					  <xsl:otherwise>
						  <xsl:value-of select="'0'"/>
					  </xsl:otherwise>
				  </xsl:choose>
		  </xsl:attribute>
		  
		  </xsl:if>
		<xsl:if test="(p:nvCxnSpPr/p:cNvCxnSpPr/a:endCxn)">
		  <xsl:variable name="endGluePoint">
			  <xsl:value-of select="p:nvCxnSpPr/p:cNvCxnSpPr/a:endCxn/@idx"/>
		  </xsl:variable>
		  <xsl:attribute name ="draw:end-glue-point">

			  <xsl:choose>
				  <xsl:when test="parent::node()/p:sp/p:spPr/a:prstGeom/@prst = 'snip1Rect'">
					  <xsl:choose>
						  <xsl:when test="$endGluePoint = 0">
							  <xsl:value-of select="'1'"/>
						  </xsl:when>
						  <xsl:when test="$endGluePoint = 1">
							  <xsl:value-of select="'3'"/>
						  </xsl:when>
						  <xsl:when test="$endGluePoint = 2">
							  <xsl:value-of select="'2'"/>
						  </xsl:when>
						  <xsl:when test="$endGluePoint = 3">
							  <xsl:value-of select="'0'"/>
						  </xsl:when>
						  <xsl:otherwise>
							  <xsl:value-of select="'0'"/>
						  </xsl:otherwise>
					  </xsl:choose>

				  </xsl:when>
				  <xsl:otherwise>
					  <xsl:choose>
						  <xsl:when test="$endGluePoint = 0">
							  <xsl:value-of select="'0'"/>
						  </xsl:when>
						  <xsl:when test="$endGluePoint = 1">
							  <xsl:value-of select="'3'"/>
						  </xsl:when>
						  <xsl:when test="$endGluePoint = 2">
							  <xsl:value-of select="'2'"/>
						  </xsl:when>
						  <xsl:when test="$endGluePoint = 3">
							  <xsl:value-of select="'1'"/>
						  </xsl:when>
						  <xsl:otherwise>
							  <xsl:value-of select="'0'"/>
						  </xsl:otherwise>
					  </xsl:choose>

				  </xsl:otherwise>
			  </xsl:choose>
		  </xsl:attribute>
		  
		  <xsl:attribute name ="draw:end-shape">
			  <xsl:value-of select ="concat('sl',$sldId,'an', p:nvCxnSpPr/p:cNvCxnSpPr/a:endCxn/@id)"/>
		  </xsl:attribute>
	  </xsl:if>
	  
  </xsl:template>
  <!-- Add text to the shape -->
  <xsl:template name ="AddShapeText">
    <!-- Extra parameter "sldId" added by lohith,requierd for template AddTextHyperlinks -->
    <xsl:param name ="sldID" />
    <xsl:param name ="prID" />
    <xsl:param name="SlideRelationId"/>
    <xsl:param name="TypeId"/>
    <xsl:variable name ="SlideNumber" select ="substring-before(substring-after($SlideRelationId,'ppt/slides/_rels/'),'.xml.rels')"/>
    <!--concat($SlideNumber,'textboxshape_List',position()-->
    <xsl:for-each select ="p:txBody">
      <xsl:variable name ="textNumber" select ="./parent::node()/p:nvSpPr/p:cNvPr/@id"/>
      <!--code inserted by Vijayeta,InsertStyle For Bullets and Numbering-->
      <xsl:variable name ="listStyleName">
        <xsl:value-of select ="concat($SlideNumber,'textboxshape_List',$textNumber)"/>
      </xsl:variable>
      <xsl:for-each select ="a:p">
        <!--Code Inserted by Vijayeta for Bullets And Numbering
                  check for levels and then depending on the condition,insert bullets,Layout or Master properties-->
        <xsl:if test ="a:pPr/a:buChar or a:pPr/a:buAutoNum or a:pPr/a:buBlip">
          <xsl:call-template name ="insertBulletsNumbersoox2odf">
            <xsl:with-param name ="listStyleName" select ="concat($listStyleName,position())"/>
            <xsl:with-param name ="ParaId" select ="$prID"/>
            <!-- parameters 'slideRelationId' and 'slideId' added by lohith - required to set Hyperlinks for bulleted text -->
            <xsl:with-param name="slideRelationId" select="$SlideRelationId" />
            <xsl:with-param name="slideId" select="$sldID" />
            <xsl:with-param name="TypeId" select="$TypeId" />
          </xsl:call-template>
        </xsl:if>
        <xsl:if test ="not(a:pPr/a:buChar) and not(a:pPr/a:buAutoNum)and not(a:pPr/a:buBlip)">
          <text:p >
			<xsl:attribute name ="text:id">
			   <xsl:value-of select ="concat('slText',$sldID,'an',./parent::node()/parent::node()/p:nvSpPr/p:cNvPr/@id, position())"/>
			</xsl:attribute>
            <xsl:attribute name ="text:style-name">
              <xsl:value-of select ="concat($prID,position())"/>
            </xsl:attribute>
            <xsl:for-each select ="node()">
              <xsl:if test ="name()='a:r'">
                <text:span>
                  <xsl:attribute name="text:style-name">
                    <xsl:value-of select="concat($TypeId,generate-id())"/>
                  </xsl:attribute>
                  <!-- varibale 'nodeTextSpan' added by lohith.ar - need to have the text inside <text:a> tag if assigned with hyperlinks -->
                  <xsl:variable name="nodeTextSpan">
                    <!--<xsl:value-of select ="a:t"/>-->
                    <!--converts whitespaces sequence to text:s-->
                    <!-- 1699083 bug fix  -->
                    <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
                    <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
                    <xsl:choose >
                      <xsl:when test ="a:rPr[@cap='all']">
                        <xsl:choose >
                          <xsl:when test =".=''">
                            <text:s/>
                          </xsl:when>
                          <xsl:when test ="not(contains(.,'  '))">
                            <xsl:value-of select ="translate(.,$lcletters,$ucletters)"/>
                          </xsl:when>
                          <xsl:when test =". =' '">
                            <text:s/>
                          </xsl:when>
                          <xsl:otherwise >
                            <xsl:call-template name ="InsertWhiteSpaces">
                              <xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
                            </xsl:call-template>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:when>
                      <xsl:when test ="a:rPr[@cap='small']">
                        <xsl:choose >
                          <xsl:when test =".=''">
                            <text:s/>
                          </xsl:when>
                          <xsl:when test ="not(contains(.,'  '))">
                            <xsl:value-of select ="translate(.,$ucletters,$lcletters)"/>
                          </xsl:when>
                          <xsl:when test =".= ' '">
                            <text:s/>
                          </xsl:when>
                          <xsl:otherwise >
                            <xsl:call-template name ="InsertWhiteSpaces">
                              <xsl:with-param name ="string" select ="translate(.,$lcletters,$ucletters)"/>
                            </xsl:call-template>
                          </xsl:otherwise>
                        </xsl:choose >
                      </xsl:when>
                      <xsl:otherwise >
                        <xsl:choose >
                          <xsl:when test =".=''">
                            <text:s/>
                          </xsl:when>
                          <xsl:when test ="not(contains(.,'  '))">
                            <xsl:value-of select ="."/>
                          </xsl:when>
                          <xsl:otherwise >
                            <xsl:call-template name ="InsertWhiteSpaces">
                              <xsl:with-param name ="string" select ="."/>
                            </xsl:call-template>
                          </xsl:otherwise >
                        </xsl:choose>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:variable>
                  <!-- Added by lohith.ar - Code for text Hyperlinks -->
                  <xsl:if test="node()/a:hlinkClick and not(node()/a:hlinkClick/a:snd)">
                    <text:a>
                      <xsl:call-template name="AddTextHyperlinks">
                        <xsl:with-param name="nodeAColonR" select="node()" />
                        <xsl:with-param name="slideRelationId" select="$SlideRelationId" />
                        <xsl:with-param name="slideId" select="$sldID" />
                      </xsl:call-template>
                      <xsl:copy-of select="$nodeTextSpan"/>
                    </text:a>
                  </xsl:if>
                  <xsl:if test="not(node()/a:hlinkClick and not(node()/a:hlinkClick/a:snd))">
                    <xsl:copy-of select="$nodeTextSpan"/>
                  </xsl:if>
                  <!-- End - Code for text Hyperlinks -->
                </text:span>
              </xsl:if >
              <xsl:if test ="name()='a:br'">
                <text:line-break/>
              </xsl:if>
              <xsl:if test="name()='a:fld'">
                <xsl:variable name="pos" select="position()"/>
                <xsl:if test="@type='datetime'">
                  <text:span>
                    <xsl:attribute name="text:style-name">
                      <xsl:value-of select="concat($TypeId,generate-id())"/>
                    </xsl:attribute>
                    <presentation:date-time/>
                  </text:span>
                </xsl:if>
                <xsl:if test="@type='slidenum'">
                  <text:span>
                    <xsl:attribute name="text:style-name">
                      <xsl:value-of select="concat($TypeId,generate-id())"/>
                    </xsl:attribute>
                  <text:page-number/>
                
                  </text:span>
                </xsl:if>
              </xsl:if>

            
              <!-- Added by lohith.ar for fix 1731885-->
              <xsl:if test="name()='a:endParaRPr' and not(a:endParaRPr/a:hlinkClick)">
                <text:span>
                  <xsl:attribute name="text:style-name">
                    <xsl:value-of select="concat($TypeId,generate-id())"/>
                  </xsl:attribute>
                </text:span>
              </xsl:if>
            </xsl:for-each>
          </text:p>
          <!--Code inserted by vijayeta,for Bullets and Numbering If Bullets are present-->
        </xsl:if >
      </xsl:for-each>
      <!--If no bullets are present or default bullets-->

    </xsl:for-each>

  </xsl:template>
  <xsl:template name ="InsertStylesForGraphicProperties"  >
    <xsl:for-each select ="document('ppt/presentation.xml')/p:presentation/p:sldIdLst/p:sldId">
      <!-- Added by vipul-->
      <!--Start-->
      <xsl:variable name="SlidePos" select="position()"/>
      <!--End-->
      <xsl:variable name ="SlideId">
        <xsl:value-of  select ="concat(concat('slide',position()),'.xml')" />
      </xsl:variable>
      <xsl:for-each select ="document(concat('ppt/slides/',$SlideId))/p:sld/p:cSld/p:spTree">
        <xsl:call-template name ="getGraphicProperties">
          <xsl:with-param name ="SlideId" select="$SlideId" />
          <xsl:with-param name ="grID" select ="'gr'" />
        </xsl:call-template>
      </xsl:for-each>
      <!--Added by Vipul to insert style for Layout shapes-->
      <!--Start-->
      <xsl:variable name ="slideRel">
        <xsl:value-of select ="concat(concat('ppt/slides/_rels/',$SlideId),'.rels')"/>
      </xsl:variable>
      <xsl:variable name ="LayoutFileNo">
        <xsl:for-each select ="document($slideRel)//node()/@Target[contains(.,'slideLayouts')]">
          <xsl:value-of select ="concat('ppt',substring(.,3))"/>
        </xsl:for-each>
      </xsl:variable>
      <xsl:for-each select ="document($LayoutFileNo)/p:sldLayout/p:cSld/p:spTree/p:sp">
        <xsl:if test="not(p:nvSpPr/p:nvPr/p:ph)">
          <xsl:variable  name ="GraphicId">
            <xsl:value-of select ="concat('SL',$SlidePos,'LYT','gr',position())"/>
          </xsl:variable>
          <xsl:variable name ="ParaId">
            <xsl:value-of select ="concat('SL',$SlidePos,'LYT','PARA',position())"/>
          </xsl:variable>
          <xsl:variable name="flagTextBox">
            <xsl:if test="p:nvSpPr/p:cNvSpPr/@txBox='1'">
              <xsl:value-of select ="'True'"/>
            </xsl:if>
          </xsl:variable>
          <style:style style:family="graphic" style:parent-style-name="standard">
            <xsl:attribute name ="style:name">
              <xsl:value-of select ="$GraphicId"/>
            </xsl:attribute >
            <style:graphic-properties>
              <!--FILL-->
              <xsl:call-template name ="Fill" />
              <!--LINE COLOR-->
              <xsl:call-template name ="LineColor" />
              <!--LINE STYLE-->
              <xsl:call-template name ="LineStyle"/>
              <!--TEXT ALIGNMENT-->
              <xsl:call-template name ="TextLayout" />
            </style:graphic-properties >
          </style:style>
          <xsl:call-template name="tmpShapeTextProcess">
            <xsl:with-param name="ParaId" select="$ParaId"/>
            <xsl:with-param name="flagTextBox" select="$flagTextBox"/>
          </xsl:call-template>
        </xsl:if>
      </xsl:for-each>
      <!--End-->
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="getGraphicProperties">
    <xsl:param name ="SlideId" />
    <xsl:param name ="grID" />
    <!-- Graphic properties for shapes with p:sp nodes-->
    <xsl:for-each select ="p:sp">
      <!-- Generate graphic properties ID-->
      <xsl:variable  name ="GraphicId">
        <xsl:value-of select ="concat('slide',substring($SlideId,6,string-length($SlideId)-9) ,concat($grID,position()))"/>
      </xsl:variable>
      <style:style style:family="graphic" style:parent-style-name="standard">
        <xsl:attribute name ="style:name">
          <xsl:value-of select ="$GraphicId"/>
        </xsl:attribute >
        <style:graphic-properties>

          <!-- FILL -->
          <xsl:call-template name ="Fill"/>

           <!-- LINE COLOR -->
          <xsl:call-template name ="LineColor" />

          <!-- LINE STYLE -->
          <xsl:call-template name ="LineStyle"/>

          <!-- TEXT ALIGNMENT -->
          <xsl:call-template name ="TextLayout" />

		  <!-- SHADOW IMPLEMENTATION -->
		  <xsl:if test="p:spPr/a:effectLst/a:outerShdw ">
			<xsl:call-template name ="ShapesShadow"/>
		  </xsl:if>
          <xsl:if test="p:spPr/a:effectLst/a:innerShdw ">
            <!-- warn if inner Shadow -->
            <xsl:message terminate="no">translation.oox2odf.shapesTypeInnerShadow</xsl:message>
          </xsl:if>
          <xsl:message terminate="no">translation.oox2odf.shapesTypeOuterShadow</xsl:message>
        </style:graphic-properties >
        <xsl:if test ="p:txBody/a:bodyPr/@vert">
          <style:paragraph-properties>
            <xsl:attribute name ="style:writing-mode">
              <xsl:call-template name ="getTextDirection">
                <xsl:with-param name ="vert" select ="p:txBody/a:bodyPr/@vert" />
              </xsl:call-template>
            </xsl:attribute>
          </style:paragraph-properties>
        </xsl:if>
      </style:style>
      <!--</xsl:if >-->
    </xsl:for-each>
    <!-- Graphic properties for shapes with p:cxnSp nodes-->
    <xsl:for-each select ="p:cxnSp">
      <!-- Generate graphic properties ID-->
      <xsl:variable  name ="GraphicId">
        <xsl:value-of select ="concat('slide',substring($SlideId,6,string-length($SlideId)-9) ,concat($grID,'Line',position()))"/>
      </xsl:variable>

      <style:style style:family="graphic" style:parent-style-name="standard">
        <xsl:attribute name ="style:name">
          <xsl:value-of select ="$GraphicId"/>
        </xsl:attribute >
        <style:graphic-properties>

          <!-- LINE COLOR -->
          <xsl:call-template name ="LineColor" />

          <!-- LINE STYLE -->
          <xsl:call-template name ="LineStyle"/>

        </style:graphic-properties >
      </style:style>

    </xsl:for-each>
    <!-- Graphic properties for grouped shapes with p:grpSp nodes-->
    <xsl:for-each select ="p:grpSp">
      <xsl:call-template name ="getGraphicProperties">
        <xsl:with-param name ="SlideId" select="$SlideId" />
        <xsl:with-param name ="grID" select ="concat('grp',generate-id())" />
      </xsl:call-template>
    </xsl:for-each>
  </xsl:template>
  <!-- Get fill color for shape-->
  <xsl:template name="Fill">
    <xsl:param name="var_pos"/>
    <xsl:param name="FileType"/>
    <xsl:param name="flagGroup"/>
    <xsl:param name="SMName"/>
    <xsl:choose>
      <!-- No fill -->
      <xsl:when test ="p:spPr/a:noFill">
        <xsl:attribute name ="draw:fill">
          <xsl:value-of select="'none'" />
        </xsl:attribute>
        <xsl:attribute name ="draw:fill-color">
          <xsl:value-of select="'#ffffff'"/>
        </xsl:attribute>
      </xsl:when>
      <!-- Solid fill-->
      <xsl:when test ="p:spPr/a:solidFill">
        <xsl:if test ="p:spPr/a:solidFill/a:srgbClr/@val">
        <xsl:attribute name ="draw:fill">
          <xsl:value-of select="'solid'" />
        </xsl:attribute>
          <xsl:attribute name ="draw:fill-color">
            <xsl:value-of select="concat('#',p:spPr/a:solidFill/a:srgbClr/@val)"/>
          </xsl:attribute>
          <!-- Transparency percentage-->
          <xsl:if test="p:spPr/a:solidFill/a:srgbClr/a:alpha/@val">
            <xsl:variable name ="alpha">
              <xsl:value-of select ="p:spPr/a:solidFill/a:srgbClr/a:alpha/@val"/>
            </xsl:variable>
            <xsl:if test="($alpha != '') or ($alpha != 0)">
              <xsl:attribute name ="draw:opacity">
                <xsl:value-of select="concat(($alpha div 1000), '%')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
        </xsl:if>
        <!--Theme color-->
        <xsl:if test ="p:spPr/a:solidFill/a:schemeClr/@val">
          <xsl:variable name="var_schemeClr" select="p:spPr/a:solidFill/a:schemeClr/@val"/>
          <xsl:attribute name ="draw:fill">
            <xsl:value-of select="'solid'" />
          </xsl:attribute>
          <xsl:attribute name ="draw:fill-color">
            <xsl:call-template name ="getColorCode">
              <xsl:with-param name ="color">
                <xsl:choose>
                  <xsl:when test="$SMName!=''">
                    <xsl:for-each select="document(concat('ppt/slideMasters/',$SMName))//p:clrMap">
                      <xsl:call-template name="tmpThemeClr">
                        <xsl:with-param name="ClrMap" select="$var_schemeClr"/>
                      </xsl:call-template>
                    </xsl:for-each>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="$var_schemeClr"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:with-param>
              <xsl:with-param name ="lumMod">
                <xsl:value-of select="p:spPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumOff">
                <xsl:value-of select="p:spPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="shade">
                <xsl:for-each select="p:spPr/a:solidFill/a:schemeClr/a:shade/@val">
                  <xsl:value-of select=". div 1000"/>
                </xsl:for-each>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
          <!-- Transparency percentage-->
          <xsl:if test="p:spPr/a:solidFill/a:schemeClr/a:alpha/@val">
            <xsl:variable name ="alpha">
              <xsl:value-of select ="p:spPr/a:solidFill/a:schemeClr/a:alpha/@val"/>
            </xsl:variable>
            <xsl:if test="($alpha != '') or ($alpha != 0)">
              <xsl:attribute name ="draw:opacity">
                <xsl:value-of select="concat(($alpha div 1000), '%')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
        </xsl:if>
      </xsl:when>
      <xsl:when test="p:spPr/a:blipFill and p:spPr/a:blipFill/a:blip/@r:embed ">
        <xsl:attribute name="draw:fill-color">
          <xsl:value-of select="'#FFFFFF'"/>
        </xsl:attribute>
        <xsl:attribute name="draw:fill">
          <xsl:value-of select="'bitmap'"/>
        </xsl:attribute>
        <xsl:choose>
          <xsl:when test="p:spPr/a:blipFill/a:stretch/a:fillRect">
            <xsl:attribute name="style:repeat">
              <xsl:value-of select="'stretch'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="p:spPr/a:blipFill/a:tile">
            <xsl:for-each select="./p:spPr/a:blipFill/a:tile">
              <xsl:if test="@sx">
                <xsl:attribute name="draw:fill-image-ref-point-x">
                  <xsl:value-of select ="concat(format-number(@sx div 1000, '#'), '%')"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="@sy">
                <xsl:attribute name="draw:fill-image-ref-point-y">
                  <xsl:value-of select ="concat(format-number(@sy div 1000, '#'), '%')"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="@algn">
                <xsl:attribute name="draw:fill-image-ref-point">
                  <xsl:choose>
                    <xsl:when test="@algn='tl'">
                      <xsl:value-of select ="'top-left'"/>
                    </xsl:when>
                    <xsl:when test="@algn='t'">
                      <xsl:value-of select ="'top'"/>
                    </xsl:when>
                    <xsl:when test="@algn='tr'">
                      <xsl:value-of select ="'top-right'"/>
                    </xsl:when>
                    <xsl:when test="@algn='r'">
                      <xsl:value-of select ="'right'"/>
                    </xsl:when>
                    <xsl:when test="@algn='bl'">
                      <xsl:value-of select ="'bottom-left'"/>
                    </xsl:when>
                    <xsl:when test="@algn='br'">
                      <xsl:value-of select ="'bottom-right'"/>
                    </xsl:when>
                    <xsl:when test="@algn='b'">
                      <xsl:value-of select ="'bottom'"/>
                    </xsl:when>
                    <xsl:when test="@algn='ctr'">
                      <xsl:value-of select ="'center'"/>
                    </xsl:when>
                  </xsl:choose>
                </xsl:attribute>
              </xsl:if>
              <!--<xsl:attribute name="draw:tile-repeat-offset">
                      <xsl:value-of select ="'100% horizontal'"/>
                    </xsl:attribute>-->
            </xsl:for-each>
          </xsl:when>
        </xsl:choose>
        <xsl:choose>
          <xsl:when test="$flagGroup='True'">
            <xsl:attribute name="draw:fill-image-name">
              <xsl:value-of select="concat($FileType,'-grpshapeImg',$var_pos)"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="draw:fill-image-name">
              <xsl:value-of select="concat($FileType,'-shapeImg',$var_pos)"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>

      </xsl:when>
      <!--added by vipul for gradient fill-->
      <xsl:when test="p:spPr/a:gradFill">
        <xsl:call-template name="tmpGradFillColor">
          <xsl:with-param name="var_pos" select="$var_pos"/>
          <xsl:with-param name="FileType" select="$FileType"/>
          <xsl:with-param name="flagGroup" select="$flagGroup"/>
          <xsl:with-param name="SMName" select="$SMName"/>
        </xsl:call-template>
        </xsl:when>
      <xsl:when test ="p:style/a:fillRef">
        <!--Fill refernce-->
        <xsl:choose>
          <xsl:when test ="p:style/a:fillRef/@idx = 0">
            <xsl:attribute name ="draw:fill">
              <xsl:value-of select="'none'" />
            </xsl:attribute>
            
          </xsl:when>
          <xsl:when test ="p:style/a:fillRef/@idx > 0">
            <xsl:attribute name ="draw:fill">
              <xsl:value-of select="'solid'" />
            </xsl:attribute>

          </xsl:when>
        </xsl:choose>
          <!-- Standard color-->
          <xsl:if test ="p:style/a:fillRef/a:srgbClr/@val">
            <xsl:attribute name ="draw:fill-color">
              <xsl:value-of select="concat('#',p:style/a:fillRef/a:srgbClr/@val)"/>
            </xsl:attribute>
            <!-- Shade percentage-->
            <!--<xsl:if test="p:style/a:fillRef/a:srgbClr/a:shade/@val">
						<xsl:variable name ="shade">
							<xsl:value-of select ="a:solidFill/a:srgbClr/a:shade/@val"/>
						</xsl:variable>
						<xsl:if test="($shade != '') or ($shade != 0)">
							<xsl:attribute name ="draw:shadow-opacity">
								<xsl:value-of select="concat(($shade div 1000), '%')"/>
							</xsl:attribute>
						</xsl:if>
					</xsl:if>-->
          </xsl:if>

          <!--Theme color-->
          <xsl:if test ="p:style/a:fillRef/a:schemeClr/@val">
            <xsl:attribute name ="draw:fill-color">
              <xsl:call-template name ="getColorCode">
                <xsl:with-param name ="color">
                  <xsl:value-of select="p:style/a:fillRef/a:schemeClr/@val"/>
                </xsl:with-param>
                <xsl:with-param name ="lumMod">
                  <xsl:value-of select="p:style/a:fillRef/a:schemeClr/a:lumMod/@val"/>
                </xsl:with-param>
                <xsl:with-param name ="lumOff">
                  <xsl:value-of select="p:style/a:fillRef/a:schemeClr/a:lumOff/@val"/>
                </xsl:with-param>
                <xsl:with-param name ="shade">
                  <xsl:for-each select="p:style/a:fillRef/a:schemeClr/a:shade/@val">
                    <xsl:value-of select=". div 1000"/>
                  </xsl:for-each>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
            <!-- Shade percentage-->
            <!--<xsl:if test="a:solidFill/a:schemeClr/a:shade/@val">
						<xsl:variable name ="shade">
							<xsl:value-of select ="a:solidFill/a:schemeClr/a:shade/@val"/>
						</xsl:variable>
						<xsl:if test="($shade != '') or ($shade != 0)">
							<xsl:attribute name ="draw:shadow-opacity">
								<xsl:value-of select="concat(($shade div 1000), '%')"/>
							</xsl:attribute>
						</xsl:if>
					</xsl:if>-->
          </xsl:if>
        

      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name ="draw:fill">
          <xsl:value-of select="'none'" />
        </xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!-- Get border color for shape -->
  <xsl:template name ="LineColor">
    <xsl:param name="SMName"/>
    <xsl:choose>
      <!-- No line-->
      <xsl:when test ="p:spPr/a:ln/a:noFill">
        <xsl:attribute name ="draw:stroke">
          <xsl:value-of select="'none'" />
        </xsl:attribute>
      </xsl:when>

      <!-- Solid line color-->
      <xsl:when test ="p:spPr/a:ln/a:solidFill">
        <xsl:attribute name ="draw:stroke">
          <xsl:value-of select="'solid'" />
        </xsl:attribute>
        <!-- Standard color for border-->
        <xsl:if test ="p:spPr/a:ln/a:solidFill/a:srgbClr/@val">
          <xsl:attribute name ="svg:stroke-color">
            <xsl:value-of select="concat('#',p:spPr/a:ln/a:solidFill/a:srgbClr/@val)"/>
          </xsl:attribute>
          <!-- Transparency percentage-->
          <xsl:if test="p:spPr/a:ln/a:solidFill/a:srgbClr/a:alpha/@val">
            <xsl:variable name ="alpha">
              <xsl:value-of select ="p:spPr/a:ln/a:solidFill/a:srgbClr/a:alpha/@val"/>
            </xsl:variable>
            <xsl:if test="($alpha != '') or ($alpha != 0)">
              <xsl:attribute name ="svg:stroke-opacity">
                <xsl:value-of select="concat(($alpha div 1000), '%')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
        </xsl:if>
        <!-- Theme color for border-->
        <xsl:if test ="p:spPr/a:ln/a:solidFill/a:schemeClr/@val">
          <xsl:variable name="var_schemeClr" select="p:spPr/a:ln/a:solidFill/a:schemeClr/@val"/>
          <xsl:attribute name ="svg:stroke-color">
            <xsl:call-template name ="getColorCode">
              <xsl:with-param name ="color">
                <xsl:choose>
                  <xsl:when test="$SMName!=''">
                    <xsl:for-each select="document(concat('ppt/slideMasters/',$SMName))//p:clrMap">
                      <xsl:call-template name="tmpThemeClr">
                        <xsl:with-param name="ClrMap" select="$var_schemeClr"/>
                      </xsl:call-template>
                    </xsl:for-each>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="$var_schemeClr"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:with-param>
              <xsl:with-param name ="lumMod">
                <xsl:value-of select="p:spPr/a:ln/a:solidFill/a:schemeClr/a:lumMod/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumOff">
                <xsl:value-of select="p:spPr/a:ln/a:solidFill/a:schemeClr/a:lumOff/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="shade">
                <xsl:for-each select="p:spPr/a:ln/a:solidFill/a:schemeClr/a:shade/@val">
                  <xsl:value-of select=". div 1000"/>
                </xsl:for-each>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
          <!-- Transparency percentage-->
          <xsl:if test="p:spPr/a:ln/a:solidFill/a:schemeClr/a:alpha/@val">
            <xsl:variable name ="alpha">
              <xsl:value-of select ="p:spPr/a:ln/a:solidFill/a:schemeClr/a:alpha/@val"/>
            </xsl:variable>
            <xsl:if test="($alpha != '') or ($alpha != 0)">
              <xsl:attribute name ="svg:stroke-opacity">
                <xsl:value-of select="concat(($alpha div 1000), '%')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
        </xsl:if>
      </xsl:when>
		    	  <xsl:when test ="p:style/a:lnRef">
            <xsl:attribute name ="draw:stroke">
              <xsl:value-of select="'solid'" />
            </xsl:attribute>
            <!--Standard color for border-->
            <xsl:if test ="p:style/a:lnRef/a:srgbClr/@val">
              <xsl:attribute name ="svg:stroke-color">
                <xsl:value-of select="concat('#',p:style/a:lnRef/a:srgbClr/@val)"/>
              </xsl:attribute>
            </xsl:if>
            <!--Theme color for border-->
            <xsl:if test ="p:style/a:lnRef/a:schemeClr/@val">
              <xsl:attribute name ="svg:stroke-color">
                <xsl:call-template name ="getColorCode">
                  <xsl:with-param name ="color">
                    <xsl:value-of select="p:style/a:lnRef/a:schemeClr/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumMod">
                    <xsl:value-of select="p:style/a:lnRef/a:schemeClr/a:lumMod/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="lumOff">
                    <xsl:value-of select="p:style/a:lnRef/a:schemeClr/a:lumOff/@val"/>
                  </xsl:with-param>
                  <xsl:with-param name ="shade">
                    <xsl:for-each select="p:style/a:lnRef/a:schemeClr/a:shade/@val">
                      <xsl:value-of select=". div 1000"/>
                    </xsl:for-each>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
                 </xsl:if>
          </xsl:when>
      <xsl:when test ="not(p:spPr/a:ln)">
        <xsl:attribute name ="draw:stroke">
          <xsl:value-of select="'none'" />
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!-- Get line styles for shape -->
  <xsl:template name ="LineStyle">
    <xsl:param name="ThemeName"/>
    <!-- Line width-->
    <xsl:for-each select ="p:spPr">
      <xsl:if test ="a:ln/@w">
        <xsl:attribute name ="svg:stroke-width">
          <xsl:call-template name="ConvertEmu">
            <xsl:with-param name="length" select="a:ln/@w"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if> 
      <xsl:if test ="not(a:ln/@w) and (parent::node()/p:nvSpPr/p:cNvPr/@name[contains(., 'Text')])">
        <xsl:attribute name ="draw:stroke">
          <xsl:value-of select ="'none'"/>
        </xsl:attribute>
      </xsl:if>
    
      <!--Bug fix for default BlueBorder in textbox by Mathi on 26thAug 2007-->
      <xsl:if test="(not(a:ln) and not(parent::node()/p:nvSpPr/p:cNvSpPr/@txBox) and not(parent::node()/p:nvSpPr/p:nvPr/p:ph) and not(parent::node()/p:style)) or 
		              (not(a:ln/@w) and (parent::node()/p:nvSpPr/p:cNvPr/@name[contains(., 'Content Placeholder')]) and not(parent::node()/p:nvSpPr/p:nvPr/p:ph) and not(parent::node()/p:style)) or 
					  ((parent::node()/p:nvSpPr/p:cNvPr/@name[contains(., 'Title ')]) and (parent::node()/p:nvSpPr/p:cNvSpPr/@txBox) and not(parent::node()/p:nvSpPr/p:nvPr/p:ph) and not(parent::node()/p:style))">
			<xsl:attribute name ="draw:stroke">
				<xsl:value-of select ="'none'"/>
			</xsl:attribute>
		</xsl:if>
		<!--End of Code-->
    </xsl:for-each>
    <xsl:choose>
      <xsl:when test ="p:style/a:lnRef/@idx and not(p:spPr/a:ln/@w)">
        <xsl:variable name="idx" select="p:style/a:lnRef/@idx"/>
        <xsl:for-each select ="document($ThemeName)//a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln">
          <xsl:if test="position()=$idx">
            <xsl:attribute name ="svg:stroke-width">
              <xsl:value-of select="concat(format-number(./@w div 360000, '#.#####'), 'cm')"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="not(p:spPr/a:ln/@w) and p:spPr/a:ln">
        <xsl:for-each select ="document($ThemeName)//a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln">
          <xsl:if test="position()=1">
            <xsl:attribute name ="svg:stroke-width">
              <xsl:value-of select="concat(format-number(./@w div 360000, '#.#####'), 'cm')"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
    </xsl:choose>
  
    <!-- Line Dash property-->
    <xsl:for-each select ="p:spPr/a:ln">
      <xsl:if test ="not(a:noFill)">
        <xsl:choose>
			<!--Bug fix for shape outline by Mathi-->
			<xsl:when test ="(a:prstDash/@val='solid') or not(a:prstDash/@val)">
				<xsl:attribute name ="draw:stroke">
					<xsl:value-of select ="'solid'"/>
				</xsl:attribute>
			</xsl:when>
			<xsl:otherwise>
				<xsl:attribute name ="draw:stroke">
					<xsl:value-of select ="'dash'"/>
				</xsl:attribute>
				<xsl:attribute name ="draw:stroke-dash">
              <xsl:choose>
                <xsl:when test="(a:prstDash/@val='sysDot') and (@cap='rnd')">
                  <xsl:value-of select ="'sysDotRound'"/>
                </xsl:when>
                <xsl:when test="a:prstDash/@val='sysDot'">
                  <xsl:value-of select ="'sysDot'"/>
                </xsl:when>
                <xsl:when test ="(a:prstDash/@val='sysDash') and (@cap='rnd')">
                  <xsl:value-of select ="'sysDashRound'"/>
                </xsl:when>
                <xsl:when test ="a:prstDash/@val='sysDash'">
                  <xsl:value-of select ="'sysDash'"/>
                </xsl:when>
                <xsl:when test ="(a:prstDash/@val='dash') and (@cap='rnd')">
                  <xsl:value-of select ="'dashRound'"/>
                </xsl:when>
                <xsl:when test ="a:prstDash/@val='dash'">
                  <xsl:value-of select ="'dash'"/>
                </xsl:when>
                <xsl:when test ="(a:prstDash/@val='dashDot') and (@cap='rnd')">
                  <xsl:value-of select ="'dashDotRound'"/>
                </xsl:when>
                <xsl:when test ="a:prstDash/@val='dashDot'">
                  <xsl:value-of select ="'dashDot'"/>
                </xsl:when>
                <xsl:when test ="(a:prstDash/@val='lgDash') and (@cap='rnd')">
                  <xsl:value-of select ="'lgDashRound'"/>
                </xsl:when>
                <xsl:when test ="a:prstDash/@val='lgDash'">
                  <xsl:value-of select ="'lgDash'"/>
                </xsl:when>
                <xsl:when test ="(a:prstDash/@val='lgDashDot') and (@cap='rnd')">
                  <xsl:value-of select ="'lgDashDotRound'"/>
                </xsl:when>
                <xsl:when test ="a:prstDash/@val='lgDashDot'">
                  <xsl:value-of select ="'lgDashDot'"/>
                </xsl:when>
                <xsl:when test ="(a:prstDash/@val='lgDashDotDot') and (@cap='rnd')">
                  <xsl:value-of select ="'lgDashDotDotRound'"/>
                </xsl:when>
                <xsl:when test ="a:prstDash/@val='lgDashDotDot'">
                  <xsl:value-of select ="'lgDashDotDot'"/>
                </xsl:when>
              </xsl:choose>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
    </xsl:for-each >

    <!-- Line join property -->
    <xsl:choose>
      <xsl:when test ="p:spPr/a:ln/a:miter">
        <xsl:attribute name ="draw:stroke-linejoin">
          <xsl:value-of select ="'miter'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="p:spPr/a:ln/a:bevel">
        <xsl:attribute name ="draw:stroke-linejoin">
          <xsl:value-of select ="'bevel'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="p:spPr/a:ln/a:round">
        <xsl:attribute name ="draw:stroke-linejoin">
          <xsl:value-of select ="'round'"/>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>

	  <!-- Line Arrow -->
	  <!-- Head End-->
	  <xsl:for-each select ="p:spPr/a:ln/a:headEnd">
		  <xsl:if test ="@type">
			  <xsl:attribute name ="draw:marker-start">
				  <xsl:value-of select ="@type"/>
			  </xsl:attribute>
			  <xsl:variable name="lnw">
				  <xsl:call-template name="ConvertEmu">
					  <xsl:with-param name="length" select="parent::node()/parent::node()/a:ln/@w"/>
					  <xsl:with-param name="unit">cm</xsl:with-param>
				  </xsl:call-template>
			  </xsl:variable>
			  <xsl:attribute name ="draw:marker-start-width">
				  <xsl:call-template name ="getArrowSize">
					  <xsl:with-param name ="pptlw" select ="parent::node()/parent::node()/a:ln/@w" />
					  <xsl:with-param name ="lw" select ="substring-before($lnw,'cm')" />
					  <xsl:with-param name ="type" select ="@type" />
					  <xsl:with-param name ="w" select ="@w" />
					  <xsl:with-param name ="len" select ="@len" />
				  </xsl:call-template>
			  </xsl:attribute>
		  </xsl:if>
	  </xsl:for-each>


	  <!-- Tail End-->
	  <xsl:for-each select ="p:spPr/a:ln/a:tailEnd">
		  <xsl:if test ="@type">
			  <xsl:attribute name ="draw:marker-end">
				  <xsl:value-of select ="@type"/>
			  </xsl:attribute>
			  <xsl:variable name="lnw">
				  <xsl:call-template name="ConvertEmu">
					  <xsl:with-param name="length" select="parent::node()/parent::node()/a:ln/@w"/>
					  <xsl:with-param name="unit">cm</xsl:with-param>
				  </xsl:call-template>
			  </xsl:variable>
			  <xsl:attribute name ="draw:marker-end-width">
				  <xsl:call-template name ="getArrowSize">
					  <xsl:with-param name ="pptlw" select ="parent::node()/parent::node()/a:ln/@w" />
					  <xsl:with-param name ="lw" select ="substring-before($lnw,'cm')" />
					  <xsl:with-param name ="type" select ="@type" />
					  <xsl:with-param name ="w" select ="@w" />
					  <xsl:with-param name ="len" select ="@len" />
				  </xsl:call-template>
			  </xsl:attribute>
		  </xsl:if>
	  </xsl:for-each>
  </xsl:template>
	
  <!-- Get arrow size -->
	<xsl:template name ="getArrowSize">
		<xsl:param name ="pptlw" />
		<xsl:param name ="lw" />
		<xsl:param name ="type" />
		<xsl:param name ="w" />
		<xsl:param name ="len" />
    <!-- warn, arrow head and tail size -->
    <xsl:message terminate="no">translation.oox2odf.shapesTypeArrowHeadTailSize</xsl:message>
		<xsl:choose>
			<!-- selection for (top row arrow type) and (arrow size as 'arrow') -->
			<xsl:when test ="(($pptlw &gt; '24765') and ($type = 'arrow') and ($w='sm') and ($len='med')) or (($pptlw &gt; '24765') and ($type = 'arrow') and ($w='sm') and ($len='sm')) or (($pptlw &gt; '24765') and ($type = 'arrow') and ($w='sm') and ($len='lg'))">
				<xsl:value-of select ="concat(($lw) * (3.5),'cm')"/>
			</xsl:when>
			<xsl:when test ="(($pptlw &lt; '24766')  and ($type = 'arrow') and ($w='sm') and ($len='med')) or (($pptlw &lt; '24766')  and ($type = 'arrow') and ($w='sm') and ($len='sm')) or (($pptlw &lt; '24766')  and ($type = 'arrow') and ($w='sm') and ($len='lg'))">
				<xsl:value-of select ="concat($sm-med,'cm')"/>
			</xsl:when>
			<!-- selection for (top row arrow type) and non-selection of (arrow size as 'arrow') -->
			<xsl:when test ="(($pptlw &gt; '25400') and ($w='sm') and ($len='med')) or (($pptlw &gt; '25400') and ($w='sm') and ($len='sm')) or (($pptlw &gt; '25400') and ($w='sm') and ($len='lg'))">
				<xsl:value-of select ="concat(($lw) * (2),'cm')"/>
			</xsl:when>
			<xsl:when test ="(($pptlw &lt; '25401')  and ($w='sm') and ($len='med')) or (($pptlw &lt; '25401')  and ($w='sm') and ($len='sm')) or (($pptlw &lt; '25401')  and ($w='sm') and ($len='lg'))">
				<xsl:value-of select ="concat($sm-sm,'cm')"/>
			</xsl:when>

			<!-- selection for (middle row arrow type) and (arrow size as 'arrow') -->
			<xsl:when test ="(($pptlw &gt; '28575') and ($type = 'arrow')) or (($pptlw &gt; '28575') and ($type = 'arrow') and ($w='med') and ($len='sm')) or (($pptlw &gt; '28575') and ($type = 'arrow') and ($w='med') and ($len='med')) or (($pptlw &gt; '28575') and ($type = 'arrow') and ($w='med') and ($len='lg'))">
				<xsl:value-of select ="concat(($lw) * (4.5),'cm')"/>
			</xsl:when>
			<xsl:when test ="(($pptlw &lt; '28576') and ($type = 'arrow')) or (($pptlw &lt; '28576')  and ($type = 'arrow') and ($w='med') and ($len='sm')) or (($pptlw &lt; '28576')  and ($type = 'arrow') and ($w='med') and ($len='med')) or (($pptlw &lt; '28576')  and ($type = 'arrow') and ($w='med') and ($len='lg'))">
				<xsl:value-of select ="concat($med-med,'cm')"/>
			</xsl:when>
			<!-- selection for (middle row arrow type) and non-selection of (arrow size as 'arrow') -->
			<xsl:when test ="(($pptlw &gt; '28575')) or (($pptlw &gt; '28575') and ($w='med') and ($len='sm')) or (($pptlw &gt; '28575') and ($w='med') and ($len='med')) or (($pptlw &gt; '28575') and ($w='med') and ($len='lg'))">
				<xsl:value-of select ="concat(($lw) * (3),'cm')"/>
			</xsl:when>
			<xsl:when test ="(($pptlw &lt; '28576')) or (($pptlw &lt; '28576') and ($w='med') and ($len='sm')) or (($pptlw &lt; '28576') and ($w='med') and ($len='med')) or (($pptlw &lt; '28576') and ($w='med') and ($len='lg'))">
				<xsl:value-of select ="concat($med-sm,'cm')"/>
			</xsl:when>

			<!-- selection for (bottom row arrow type) and (arrow size as 'arrow') -->
			<xsl:when test ="(($pptlw &gt; '24765') and ($type = 'arrow') and ($w='lg') and ($len='med')) or  (($pptlw &gt; '24765') and ($type = 'arrow') and ($w='lg') and ($len='sm')) or (($pptlw &gt; '24765') and ($type = 'arrow') and ($w='lg') and ($len='lg'))">
				<xsl:value-of select ="concat(($lw) * (6),'cm')"/>
			</xsl:when>
			<xsl:when test ="(($pptlw &lt; '24766')  and ($type = 'arrow') and ($w='lg') and ($len='med')) or (($pptlw &lt; '24766') and ($type = 'arrow') and ($w='lg') and ($len='sm')) or (($pptlw &lt; '24766')  and ($type = 'arrow') and ($w='lg') and ($len='lg'))">
				<xsl:value-of select ="concat($lg-med,'cm')"/>
			</xsl:when>
			<!-- selection for (bottom row arrow type) and non-selection of (arrow size as 'arrow') -->
			<xsl:when test ="(($pptlw &gt; '25400') and ($w='lg') and ($len='med')) or (($pptlw &gt; '25400') and ($w='lg') and ($len='sm')) or (($pptlw &gt; '25400') and ($w='lg') and ($len='lg'))">
				<xsl:value-of select ="concat(($lw) * (5),'cm')"/>
			</xsl:when>
			<xsl:when test ="(($pptlw &lt; '25401')  and ($w='lg') and ($len='med')) or (($pptlw &lt; '25401')  and ($w='lg') and ($len='sm')) or (($pptlw &lt; '25401')  and ($w='lg') and ($len='lg'))">
				<xsl:value-of select ="concat($lg-sm,'cm')"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select ="concat($med-med,'cm')"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	
  <!-- Get text layout for shapes -->
  <xsl:template name ="TextLayout">
    <xsl:for-each select="p:txBody">
      <xsl:if test="a:bodyPr/@vert='vert'">
        <xsl:choose>
          <xsl:when test="(a:bodyPr/@anchor='b' and a:bodyPr/@anchorCtr='0') or (a:bodyPr/@anchor='b' and not(a:bodyPr/@anchorCtr))">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'left'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'top'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="(a:bodyPr/@anchor='t' and a:bodyPr/@anchorCtr='0') or (a:bodyPr/@anchor='t' and not(a:bodyPr/@anchorCtr))">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'right'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'top'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="(a:bodyPr/@anchor='ctr' and a:bodyPr/@anchorCtr='0') or (a:bodyPr/@anchor='ctr' and not(a:bodyPr/@anchorCtr))">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'center'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'top'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="a:bodyPr/@anchor='t' and a:bodyPr/@anchorCtr='1'">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'right'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'middle'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="a:bodyPr/@anchor='ctr' and a:bodyPr/@anchorCtr='1'">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'center'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'middle'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="a:bodyPr/@anchor='b' and a:bodyPr/@anchorCtr='1'">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'left'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'middle'"/>
            </xsl:attribute>
          </xsl:when>
        </xsl:choose>
      </xsl:if>
      <xsl:if test="a:bodyPr/@vert='horz' or not(a:bodyPr/@vert) ">
        <xsl:choose>
          <xsl:when test="(a:bodyPr/@anchor='t' and a:bodyPr/@anchorCtr='0') or (a:bodyPr/@anchor='t' and not(a:bodyPr/@anchorCtr))">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'justify'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'top'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="(a:bodyPr/@anchor='ctr' and a:bodyPr/@anchorCtr='0') or (a:bodyPr/@anchor='ctr' and not(a:bodyPr/@anchorCtr))  ">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'justify'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'middle'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="a:bodyPr/@anchor='t' and a:bodyPr/@anchorCtr='1'">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'center'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'top'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="(a:bodyPr/@anchor='b' and a:bodyPr/@anchorCtr='0') or (a:bodyPr/@anchor='b' and not(a:bodyPr/@anchorCtr))">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'justify'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'bottom'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="a:bodyPr/@anchor='ctr' and a:bodyPr/@anchorCtr='1'">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'center'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'middle'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="a:bodyPr/@anchor='b' and a:bodyPr/@anchorCtr='1'">
            <xsl:attribute name ="draw:textarea-horizontal-align">
              <xsl:value-of select="'center'"/>
            </xsl:attribute>
            <xsl:attribute name ="draw:textarea-vertical-align">
              <xsl:value-of select="'bottom'"/>
            </xsl:attribute>
          </xsl:when>
        </xsl:choose>
      </xsl:if>

      <xsl:if test ="a:bodyPr/@lIns">
        <xsl:attribute name ="fo:padding-left">
          <xsl:call-template name="ConvertEmu">
            <xsl:with-param name="length" select="a:bodyPr/@lIns"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test ="not(a:bodyPr/@lIns)">
        <xsl:attribute name ="fo:padding-left">
          <xsl:call-template name="ConvertEmu">
            <xsl:with-param name="length" select="'91440'"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>

      <xsl:if test ="a:bodyPr/@tIns">
        <xsl:attribute name ="fo:padding-top">
          <xsl:call-template name="ConvertEmu">
            <xsl:with-param name="length" select="a:bodyPr/@tIns"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test ="not(a:bodyPr/@tIns)">
        <xsl:attribute name ="fo:padding-top">
          <xsl:call-template name="ConvertEmu">
            <xsl:with-param name="length" select="'45720'"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>

      <xsl:if test ="a:bodyPr/@rIns">
        <xsl:attribute name ="fo:padding-right">
          <xsl:call-template name="ConvertEmu">
            <xsl:with-param name="length" select="a:bodyPr/@rIns"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test ="not(a:bodyPr/@rIns)">
        <xsl:attribute name ="fo:padding-right">
          <xsl:call-template name="ConvertEmu">
            <xsl:with-param name="length" select="'91440'"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>

      <xsl:if test ="a:bodyPr/@bIns">
        <xsl:attribute name ="fo:padding-bottom">
          <xsl:call-template name="ConvertEmu">
            <xsl:with-param name="length" select="a:bodyPr/@bIns"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test ="not(a:bodyPr/@bIns)">
        <xsl:attribute name ="fo:padding-bottom">
          <xsl:call-template name="ConvertEmu">
            <xsl:with-param name="length" select="'45720'"/>
            <xsl:with-param name="unit">cm</xsl:with-param>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>

      <!--Wrap text in shape -->
      <xsl:choose>
        <xsl:when test="(a:bodyPr/@wrap='none')">
          <xsl:attribute name ="fo:wrap-option">
            <xsl:value-of select ="'wrap'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="(a:bodyPr/@wrap='square')  or (a:p/a:pPr/@fontAlgn='auto')">
          <xsl:attribute name ="fo:wrap-option">
            <xsl:value-of select ="'no-wrap'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name ="fo:wrap-option">
            <xsl:value-of select ="'no-wrap'"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>

		<xsl:choose>
			<xsl:when test ="(( (a:bodyPr/a:spAutoFit) or (a:bodyPr/@wrap='square') ) and ((parent::node()/p:spPr/a:prstGeom/@prst='rect') or (parent::node()/p:spPr/a:prstGeom/@prst='ellipse')) )">
				<xsl:attribute name="draw:auto-grow-height">
					<xsl:value-of select ="'false'"/>
				</xsl:attribute>
			</xsl:when>

			<!--Code Added for internal bug fix for explosion2 by Mathi-->
			<xsl:when test="( ((a:bodyPr/a:spAutoFit) and not(a:bodyPr/@wrap)) and (parent::node()/p:spPr/a:prstGeom/@prst='irregularSeal2') )">
				<xsl:attribute name="draw:auto-grow-height">
					<xsl:value-of select ="'false'"/>
				</xsl:attribute>
			</xsl:when>
			<!--End-->
			
			<!--Code Added for bug fix of text wrap inside the custom shapes by Mathi on 31st Aug2007-->
			<xsl:when test ="( (a:bodyPr/a:spAutoFit) )">   <!--or (a:bodyPr/@wrap='square') )">-->
				<xsl:attribute name="draw:auto-grow-height">
					<xsl:value-of select ="'true'"/>
				</xsl:attribute>
			</xsl:when>

			<xsl:when test ="(not(a:bodyPr/a:spAutoFit) and not(a:bodyPr/@wrap='square')) or (a:bodyPr/@wrap='none')">  <!--<xsl:when test ="(not(a:bodyPr/a:spAutoFit)">-->
				<xsl:attribute name="draw:auto-grow-height">
					<xsl:value-of select ="'false'"/>
				</xsl:attribute>
				<xsl:attribute name="draw:auto-grow-width">
					<xsl:value-of select ="'false'"/>
				</xsl:attribute>
			</xsl:when>

			<xsl:when test="(not(a:bodyPr/a:spAutoFit) and (a:bodyPr/@wrap='square'))">
				<xsl:attribute name="draw:auto-grow-height">
					<xsl:value-of select ="'false'"/>
				</xsl:attribute>
			</xsl:when>
			<!--end of bug fix code-->
		</xsl:choose>
    </xsl:for-each>
  </xsl:template>
  <!-- Get text padding-->
  <xsl:template name ="getPadding">
    <xsl:param name ="length" />
    <xsl:choose>
      <xsl:when test ="($length != '')">
        <xsl:call-template name="ConvertEmu">
          <xsl:with-param name="length" select="$length"/>
          <xsl:with-param name="unit">cm</xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test ="(a:bodyPr/@wrap='square') or (a:p/a:pPr/@fontAlgn='auto')">
        <xsl:value-of select ="'0cm'"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!--Text direction-->
  <xsl:template name ="getTextDirection">
    <xsl:param name ="vert" />
    <xsl:choose>
      <xsl:when test ="$vert = 'vert'">
        <xsl:value-of select ="'tb-rl'"/>
      </xsl:when>
      <!--<xsl:when test ="$vert = 'vert270'">
					<xsl:value-of select ="'tb-lr'"/>
				</xsl:when>
				<xsl:when test ="$vert = 'wordArtVert'">
					<xsl:value-of select ="'tb'"/>
				</xsl:when>
				<xsl:when test ="$vert = 'eaVert'">
					<xsl:value-of select ="'lr'" />
				</xsl:when>
				<xsl:when test ="$vert = 'mongolianVert'">
					<xsl:value-of select ="'lr'" />
				</xsl:when>
				<xsl:when test ="$vert = 'wordArtVertRtl'">
					<xsl:value-of select ="'lr'" />
				</xsl:when>-->
    </xsl:choose>
  </xsl:template>

	<!-- Shadow Implementation-->
	<xsl:template name ="ShapesShadow">
		<xsl:variable name ="distVal" >
			<xsl:choose>
				<xsl:when test="(p:spPr/a:effectLst/a:outerShdw/@dist != '') or (p:spPr/a:effectLst/a:outerShdw/@dist != 0)">
					<xsl:value-of select ="p:spPr/a:effectLst/a:outerShdw/@dist"/>
				</xsl:when>
				<!-- get center selection value -->
				<xsl:when test="(p:spPr/a:effectLst/a:outerShdw/@sx != '') or (p:spPr/a:effectLst/a:outerShdw/@sx != 0)">
					<xsl:value-of select ="p:spPr/a:effectLst/a:outerShdw/@sx"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select ="'0'"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>

	    <xsl:variable name ="dirVal" >
		    <xsl:choose>
				<xsl:when test="(p:spPr/a:effectLst/a:outerShdw/@dir != '') or (p:spPr/a:effectLst/a:outerShdw/@dir != 0)">
					<xsl:value-of select ="p:spPr/a:effectLst/a:outerShdw/@dir"/>
				</xsl:when>
				<!-- get center selection value -->
				<xsl:when test="(p:spPr/a:effectLst/a:outerShdw/@sy != '') or (p:spPr/a:effectLst/a:outerShdw/@sy != 0)">
					<xsl:value-of select ="p:spPr/a:effectLst/a:outerShdw/@sy"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select ="'0'"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>

		<xsl:attribute name ="draw:shadow">
			<xsl:value-of select="'visible'"/>
		</xsl:attribute>

		<xsl:choose>
			<xsl:when test="((p:spPr/a:effectLst/a:outerShdw/@sy != '') or (p:spPr/a:effectLst/a:outerShdw/@sx != ''))">
				<xsl:attribute name ="draw:shadow-offset-x">
					<xsl:value-of select ="'0cm'"/>
				</xsl:attribute>
				<xsl:attribute name ="draw:shadow-offset-y">
					<xsl:value-of select ="'0cm'"/>
				</xsl:attribute>
			</xsl:when>
			<xsl:otherwise>
				<xsl:attribute name ="draw:shadow-offset-x">
					<xsl:value-of select ="concat('shadow-offset-y:',$distVal, ':', $dirVal)"/>
				</xsl:attribute>
				<xsl:attribute name ="draw:shadow-offset-y">
					<xsl:value-of select ="concat('shadow-offset-x:',$distVal, ':', $dirVal)"/>
				</xsl:attribute>
			</xsl:otherwise>
		</xsl:choose>

		<!-- Theme Color -->
		<xsl:if test ="p:spPr/a:effectLst/a:outerShdw/a:schemeClr/@val">
			<xsl:attribute name ="draw:shadow-color">
				<xsl:call-template name ="getColorCode">
					<xsl:with-param name ="color">
						<xsl:value-of select="p:spPr/a:effectLst/a:outerShdw/a:schemeClr/@val"/>
					</xsl:with-param>
					<xsl:with-param name ="lumMod">
						<xsl:value-of select="p:spPr/a:effectLst/a:outerShdw/a:schemeClr/a:lumMod/@val"/>
					</xsl:with-param>
					<xsl:with-param name ="lumOff">
						<xsl:value-of select="p:spPr/a:effectLst/a:outerShdw/a:schemeClr/a:lumOff/@val"/>
					</xsl:with-param>
				</xsl:call-template>
			</xsl:attribute>
			<!-- Transparency percentage-->
			<xsl:if test="p:spPr/a:effectLst/a:outerShdw/a:schemeClr/a:alpha/@val">
				<xsl:variable name ="alpha">
					<xsl:value-of select ="p:spPr/a:effectLst/a:outerShdw/a:schemeClr/a:alpha/@val"/>
				</xsl:variable>
				<xsl:if test="($alpha != '') or ($alpha != 0)">
					<xsl:attribute name ="draw:shadow-opacity">
						<xsl:value-of select="concat(($alpha div 1000), '%')"/>
					</xsl:attribute>
				</xsl:if>
			</xsl:if>
		</xsl:if>

		<xsl:if test ="p:spPr/a:effectLst/a:outerShdw/a:srgbClr/@val">
			<xsl:attribute name ="draw:shadow-color">
				<xsl:value-of select="concat('#',p:spPr/a:effectLst/a:outerShdw/a:srgbClr/@val)"/>
			</xsl:attribute>
			<!-- Transparency percentage-->
			<xsl:if test="p:spPr/a:effectLst/a:outerShdw/a:srgbClr/a:alpha/@val">
				<xsl:variable name ="alpha">
					<xsl:value-of select ="p:spPr/a:effectLst/a:outerShdw/a:srgbClr/a:alpha/@val"/>
				</xsl:variable>
				<xsl:if test="($alpha != '') or ($alpha != 0)">
					<xsl:attribute name ="draw:shadow-opacity">
						<xsl:value-of select="concat(($alpha div 1000), '%')"/>
					</xsl:attribute>
				</xsl:if>
			</xsl:if>
		</xsl:if>

		<xsl:if test ="p:spPr/a:effectLst/a:outerShdw/a:prstClr/@val">
			<xsl:attribute name ="draw:shadow-color">
				<xsl:value-of select="'#000000'"/>
			</xsl:attribute>
			<!-- Transparency percentage-->
			<xsl:if test="p:spPr/a:effectLst/a:outerShdw/a:prstClr/a:alpha/@val">
				<xsl:variable name ="alpha">
					<xsl:value-of select ="p:spPr/a:effectLst/a:outerShdw/a:prstClr/a:alpha/@val"/>
				</xsl:variable>
				<xsl:if test="($alpha != '') or ($alpha != 0)">
					<xsl:attribute name ="draw:shadow-opacity">
						<xsl:value-of select="concat(($alpha div 1000), '%')"/>
					</xsl:attribute>
				</xsl:if>
			</xsl:if>
		</xsl:if>
	</xsl:template>


  <xsl:template name ="PictureBorderColor">
    <xsl:for-each select ="document('ppt/presentation.xml')/p:presentation/p:photoAlbum[@bw=1]">
      <xsl:if test="position()=1">
      <xsl:attribute name="draw:color-mode">
        <xsl:value-of select="'greyscale'"/>
      </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
    <xsl:choose>
      <!-- No line-->
      <xsl:when test ="p:spPr/a:ln/a:noFill">
        <xsl:attribute name ="draw:stroke">
          <xsl:value-of select="'none'" />
        </xsl:attribute>
      </xsl:when>
      <!-- Solid line color-->
      <xsl:when test ="p:spPr/a:ln/a:solidFill">
        <!-- Standard color for border-->
        <xsl:if test ="p:spPr/a:ln/a:solidFill/a:srgbClr/@val">
          <xsl:attribute name ="svg:stroke-color">
            <xsl:value-of select="concat('#',p:spPr/a:ln/a:solidFill/a:srgbClr/@val)"/>
          </xsl:attribute>
          <!-- Transparency percentage-->
          <xsl:if test="p:spPr/a:ln/a:solidFill/a:srgbClr/a:alpha/@val">
            <xsl:variable name ="alpha">
              <xsl:value-of select ="p:spPr/a:ln/a:solidFill/a:srgbClr/a:alpha/@val"/>
            </xsl:variable>
            <xsl:if test="($alpha != '') or ($alpha != 0)">
              <xsl:attribute name ="svg:stroke-opacity">
                <xsl:value-of select="concat(($alpha div 1000), '%')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
        </xsl:if>
        <!-- Theme color for border-->
        <xsl:if test ="p:spPr/a:ln/a:solidFill/a:schemeClr/@val">
          <xsl:attribute name ="svg:stroke-color">
            <xsl:call-template name ="getColorCode">
              <xsl:with-param name ="color">
                <xsl:value-of select="p:spPr/a:ln/a:solidFill/a:schemeClr/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumMod">
                <xsl:value-of select="p:spPr/a:ln/a:solidFill/a:schemeClr/a:lumMod/@val"/>
              </xsl:with-param>
              <xsl:with-param name ="lumOff">
                <xsl:value-of select="p:spPr/a:ln/a:solidFill/a:schemeClr/a:lumOff/@val"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
          <!-- Transparency percentage-->
          <xsl:if test="p:spPr/a:ln/a:solidFill/a:schemeClr/a:alpha/@val">
            <xsl:variable name ="alpha">
              <xsl:value-of select ="p:spPr/a:ln/a:solidFill/a:schemeClr/a:alpha/@val"/>
            </xsl:variable>
            <xsl:if test="($alpha != '') or ($alpha != 0)">
              <xsl:attribute name ="svg:stroke-opacity">
                <xsl:value-of select="concat(($alpha div 1000), '%')"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
        </xsl:if>
      </xsl:when>
      
      <xsl:otherwise>
        <!--Line reference-->
        <xsl:if test ="p:style/a:lnRef">
          <xsl:attribute name ="draw:stroke">
            <xsl:value-of select="'solid'" />
          </xsl:attribute>
          <!--Standard color for border-->
          <xsl:if test ="p:style/a:lnRef/a:srgbClr/@val">
            <xsl:attribute name ="svg:stroke-color">
              <xsl:value-of select="concat('#',p:style/a:lnRef/a:srgbClr/@val)"/>
            </xsl:attribute>
          </xsl:if>
          <!--Theme color for border-->
          <xsl:if test ="p:style/a:lnRef/a:schemeClr/@val">
            <xsl:attribute name ="svg:stroke-color">
              <xsl:call-template name ="getColorCode">
                <xsl:with-param name ="color">
                  <xsl:value-of select="p:style/a:lnRef/a:schemeClr/@val"/>
                </xsl:with-param>
                <xsl:with-param name ="lumMod">
                  <xsl:value-of select="p:style/a:lnRef/a:schemeClr/a:lumMod/@val"/>
                </xsl:with-param>
                <xsl:with-param name ="lumOff">
                  <xsl:value-of select="p:style/a:lnRef/a:schemeClr/a:lumOff/@val"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet >
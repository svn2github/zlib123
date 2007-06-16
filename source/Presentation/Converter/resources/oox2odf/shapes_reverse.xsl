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
  exclude-result-prefixes="p a dc xlink draw rels">
  <!-- Import Bullets and numbering-->
  <xsl:import href ="BulletsNumberingoox2odf.xsl"/>
	<xsl:import href="common.xsl"/>
	<!--<xsl:import href ="trignm.xsl"/>-->
    <!-- Shape constants-->
	<!-- Arrow size -->
	<xsl:variable name="sm-sm">
		<xsl:value-of select ="'0.15'"/>
	</xsl:variable>
	<xsl:variable name="sm-med">
		<xsl:value-of select ="'0.18'"/>
	</xsl:variable>
	<xsl:variable name="sm-lg">
		<xsl:value-of select ="'0.2'"/>
	</xsl:variable>
	<xsl:variable name="med-sm">
		<xsl:value-of select ="'0.21'" />
	</xsl:variable>
	<xsl:variable name="med-med">
		<xsl:value-of select ="'0.25'"/>
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
	
	<!-- Template for Shapes in reverse conversion -->
	<xsl:template  name="shapes">
		<xsl:param name="GraphicId" />
		<xsl:param name="ParaId" />
		<xsl:param name="SlideRelationId" />
		 
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
                    <xsl:value-of select ="concat('#Page',substring-before(substring-after($SlideVal,'slide'),'.xml'))"/>
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
                  <xsl:variable name="varDocumentModifiedTime">
                    <xsl:value-of select="document('docProps/core.xml')/cp:coreProperties/dcterms:modified"/>
                  </xsl:variable>
                  <xsl:variable name="varDestMediaFileTargetPath">
                    <xsl:value-of select="concat(translate($varDocumentModifiedTime,':','-'),'|',$varMediaFileRelId,'.wav')"/>
                  </xsl:variable>
                  <xsl:variable name="varMediaFilePathForOdp">
                    <xsl:value-of select="concat('../_MediaFilesForOdp_',translate($varDocumentModifiedTime,':','-'),'/',$varMediaFileRelId,'.wav')"/>
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
                    <xsl:value-of select ="concat('#Page',substring-before(substring-after($SlideVal,'slide'),'.xml'))"/>
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
                  <xsl:variable name="varDocumentModifiedTime">
                    <xsl:value-of select="document('docProps/core.xml')/cp:coreProperties/dcterms:modified"/>
                  </xsl:variable>
                  <xsl:variable name="varDestMediaFileTargetPath">
                    <xsl:value-of select="concat(translate($varDocumentModifiedTime,':','-'),'|',$varMediaFileRelId,'.wav')"/>
                  </xsl:variable>
                  <xsl:variable name="varMediaFilePathForOdp">
                    <xsl:value-of select="concat('../_MediaFilesForOdp_',translate($varDocumentModifiedTime,':','-'),'/',$varMediaFileRelId,'.wav')"/>
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
			<xsl:when test ="(p:nvSpPr/p:cNvPr/@name[contains(., 'Rectangle Custom')]) and (p:spPr/a:prstGeom/@prst='rect')">
				<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
										 draw:type="rectangle" 
										 draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- Rectangle -->
			<xsl:when test ="(p:nvSpPr/p:cNvPr/@name[contains(., 'Rectangle')]) and (p:spPr/a:prstGeom/@prst='rect')">
				<draw:rect draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					</xsl:call-template>
          <xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:rect>
			</xsl:when>
			<!-- Oval(Custom shape) -->
			<xsl:when test ="(p:nvSpPr/p:cNvPr/@name[contains(., 'Oval Custom')]) and (p:spPr/a:prstGeom/@prst='ellipse')">
				<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
											draw:text-areas="3200 3200 18400 18400" 
											draw:type="ellipse"
											draw:glue-points="10800 0 3160 3160 0 10800 3160 18440 10800 21600 18440 18440 21600 10800 18440 3160"/>
          <xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- Ellipse(Basic shape) -->
			<xsl:when test ="(p:nvSpPr/p:cNvPr/@name[contains(., 'Oval')]) and (p:spPr/a:prstGeom/@prst='ellipse')">
				<draw:ellipse draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
						<!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
						<xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
						<!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					</xsl:call-template>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:ellipse>
			</xsl:when>
			<!-- Isosceles Triangle -->
			<xsl:when test ="(p:spPr/a:prstGeom/@prst='triangle')">
				<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
						draw:glue-points="10800 0 ?f1 10800 0 21600 10800 21600 21600 21600 ?f7 10800" 
						draw:text-areas="?f1 10800 ?f2 18000 ?f3 7200 ?f4 21600" 
						draw:type="isosceles-triangle" draw:modifiers="10800" 
						draw:enhanced-path="M ?f0 0 L 21600 21600 0 21600 Z N">
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
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
											draw:glue-points="10800 0 5400 10800 0 21600 10800 21600 21600 21600 16200 10800" 
											draw:text-areas="1900 12700 12700 19700" 
											draw:type="right-triangle" 
											draw:enhanced-path="M 0 0 L 21600 21600 0 21600 0 0 Z N"/>
          <xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- Parallelogram -->
			<xsl:when test ="(p:spPr/a:prstGeom/@prst='parallelogram')">
				<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
						draw:glue-points="?f6 0 10800 ?f8 ?f11 10800 ?f9 21600 10800 ?f10 ?f5 10800" 
						draw:text-areas="?f3 ?f3 ?f4 ?f4" draw:type="parallelogram" 
						draw:modifiers="5400" draw:enhanced-path="M ?f0 0 L 21600 0 ?f1 21600 0 21600 Z N">
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
			<!-- Trapezoid -->
			<xsl:when test ="(p:spPr/a:prstGeom/@prst='trapezoid')">
				<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 914400 914400" 
											draw:extrusion-allowed="true" 
											draw:text-areas="152400 152400 762000 914400" 
											draw:glue-points="457200 0 114300 457200 457200 914400 800100 457200" 
											draw:type="mso-spt100" 
											draw:enhanced-path="M 0 914400 L 228600 0 L 685800 0 L 914400 914400 Z N">
						<draw:handle/>
					</draw:enhanced-geometry>
          <xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- Diamond   -->
			<xsl:when test ="(p:spPr/a:prstGeom/@prst='diamond')">
				<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
											draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" 
											draw:text-areas="5400 5400 16200 16200" 
											draw:type="diamond" 
											draw:enhanced-path="M 10800 0 L 21600 10800 10800 21600 0 10800 10800 0 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
				
			</xsl:when>
			<!-- Regular Pentagon -->
			<xsl:when test ="(p:spPr/a:prstGeom/@prst='pentagon')">
				<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
											draw:glue-points="10800 0 0 8260 4230 21600 10800 21600 17370 21600 21600 8260" 
											draw:text-areas="4230 5080 17370 21600" 
											draw:type="pentagon" 
											draw:enhanced-path="M 10800 0 L 0 8260 4230 21600 17370 21600 21600 8260 10800 0 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- Hexagon -->
			<xsl:when test ="(p:spPr/a:prstGeom/@prst='hexagon')">
				<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
											draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" 
											draw:text-areas="?f3 ?f3 ?f4 ?f4" draw:type="hexagon" 
											draw:modifiers="5400" 
											draw:enhanced-path="M ?f0 0 L ?f1 0 21600 10800 ?f1 21600 ?f0 21600 0 10800 Z N">
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
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
						draw:glue-points="10800 0 0 10800 10800 21600 21600 10800"
						draw:path-stretchpoint-x="10800" draw:path-stretchpoint-y="10800"
						draw:text-areas="?f5 ?f6 ?f7 ?f8" draw:type="octagon" draw:modifiers="5000"
						draw:enhanced-path="M ?f0 0 L ?f2 0 21600 ?f1 21600 ?f3 ?f2 21600 ?f0 21600 0 ?f3 0 ?f1 Z N">
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
      <!-- Cube -->
      <xsl:when test ="(p:spPr/a:prstGeom/@prst='cube')">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
						draw:glue-points="?f7 0 ?f6 ?f1 0 ?f10 ?f6 21600 ?f4 ?f10 21600 ?f9" 
						draw:path-stretchpoint-x="10800" draw:path-stretchpoint-y="10800" 
						draw:text-areas="0 ?f1 ?f4 ?f12" draw:type="cube" draw:modifiers="5400" 
						draw:enhanced-path="M 0 ?f12 L 0 ?f1 ?f2 0 ?f11 0 ?f11 ?f3 ?f4 ?f12 Z N M 0 ?f1 L ?f2 0 ?f11 0 ?f4 ?f1 Z N M ?f4 ?f12 L ?f4 ?f1 ?f11 0 ?f11 ?f3 Z N">
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
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 88 21600"
						draw:glue-points="44 ?f6 44 0 0 10800 44 21600 88 10800"
						draw:text-areas="0 ?f6 88 ?f3" draw:type="can" draw:modifiers="5400"
						draw:enhanced-path="M 44 0 C 20 0 0 ?f2 0 ?f0 L 0 ?f3 C 0 ?f4 20 21600 44 21600 68 21600 88 ?f4 88 ?f3 L 88 ?f0 C 88 ?f2 68 0 44 0 Z N M 44 0 C 20 0 0 ?f2 0 ?f0 0 ?f5 20 ?f6 44 ?f6 68 ?f6 88 ?f5 88 ?f0 88 ?f2 68 0 44 0 Z N">
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
      <!--Text Box-->
      <xsl:when test ="p:nvSpPr/p:cNvPr/@name[contains(., 'TextBox')] or p:nvSpPr/p:cNvPr/@name[contains(., 'Text Box')]">
        <xsl:choose>
          <xsl:when test ="p:nvSpPr/p:cNvPr/@name[contains(., 'TextBox Custom')]">
            <draw:custom-shape draw:layer="layout">
              <xsl:call-template name ="CreateShape">
                <xsl:with-param name="grID" select ="$GraphicId"/>
                <xsl:with-param name ="prID" select="$ParaId" />
                <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
                <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
                <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
							</xsl:call-template>
							<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
													draw:type="mso-spt202" 
													draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N"/>
							<xsl:copy-of select="$varHyperLinksForShapes" />
						</draw:custom-shape>
					</xsl:when>
					<xsl:otherwise>
						<draw:frame draw:layer="layout">
							<xsl:call-template name ="CreateShape">
								<xsl:with-param name="grID" select ="$GraphicId"/>
								<xsl:with-param name ="prID" select="$ParaId" />
							</xsl:call-template>
							<xsl:copy-of select="$varHyperLinksForShapes" />
						</draw:frame>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			
			<!-- Basic shapes end-->

			<!-- Connectors -->
			<!-- Line -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst = 'line'">
				<draw:line draw:layer="layout">
					<xsl:call-template name ="DrawLine">
						<xsl:with-param name="grID" select ="$GraphicId"/>
					</xsl:call-template>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:line>
			</xsl:when>
			<!-- Straight Connector-->
			<xsl:when test ="p:spPr/a:prstGeom/@prst[contains(., 'straightConnector')]">
				<draw:line draw:layer="layout">
					<xsl:call-template name ="DrawLine">
						<xsl:with-param name="grID" select ="$GraphicId"/>
					</xsl:call-template>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:line >
			</xsl:when>
			<!-- Elbow Connector-->
			<xsl:when test ="p:spPr/a:prstGeom/@prst[contains(., 'bentConnector')]">
				<draw:connector draw:layer="layout">
					<xsl:call-template name ="DrawLine">
						<xsl:with-param name="grID" select ="$GraphicId"/>
					</xsl:call-template>
					<xsl:copy-of select="$varHyperLinksForConnectors" />
				</draw:connector >
			</xsl:when>
			<!--Curved Connector-->
			<xsl:when test ="p:spPr/a:prstGeom/@prst[contains(., 'curvedConnector')]">
				<draw:connector draw:layer="layout" draw:type="curve">
					<xsl:call-template name ="DrawLine">
						<xsl:with-param name="grID" select ="$GraphicId"/>
					</xsl:call-template>
					<xsl:copy-of select="$varHyperLinksForConnectors" />
				</draw:connector >
			</xsl:when>
			
			<!-- Custom shapes: -->
		 			 
			<!--Rectangle -->
			<xsl:when test ="p:nvSpPr/p:cNvPr/@name[contains(., 'Rectangle Custom')]">
				<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
										 draw:type="rectangle" 
										 draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- Rounded  Rectangle -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='roundRect'">
				<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
            draw:path-stretchpoint-x="10800" draw:path-stretchpoint-y="10800"
            draw:text-areas="?f3 ?f4 ?f5 ?f6" draw:type="round-rectangle" draw:modifiers="3600"
            draw:enhanced-path="M ?f7 0 X 0 ?f8 L 0 ?f9 Y ?f7 21600 L ?f10 21600 X 21600 ?f9 L 21600 ?f8 Y ?f10 0 Z N">
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
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
						draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" 
						draw:text-areas="0 4300 21600 21600" 
						draw:mirror-horizontal="true" draw:type="flowchart-card" 
						draw:enhanced-path="M 4300 0 L 21600 0 21600 21600 0 21600 0 4300 4300 0 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- Snip Same Side Corner Rectangle -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='snip2SameRect'">
				<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
										 draw:type="rectangle" 
										 draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
				<!--<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 914400 304800" 
											draw:extrusion-allowed="true" 
											draw:text-areas="25401 25401 888999 304800" 
											draw:glue-points="914400 152400 457200 304800 0 152400 457200 0"
											draw:type="mso-spt100" 
											draw:enhanced-path="M 50801 0 L 863599 0 L 914400 50801 L 914400 304800 L 914400 304800 L 0 304800 L 0 304800 L 0 50801 Z N">
						<draw:handle/>
					</draw:enhanced-geometry>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>-->
			</xsl:when>
			<!-- Snip Diagonal Corner Rectangle -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='snip2DiagRect'">
				<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
										 draw:type="rectangle" 
										 draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
				<!--<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 914400 228600" 
											draw:extrusion-allowed="true" 
											draw:text-areas="19050 19050 895350 209550" 
											draw:glue-points="914400 114300 457200 228600 0 114300 457200 0" 
											draw:type="mso-spt100" 
											draw:enhanced-path="M 0 0 L 876299 0 L 914400 38101 L 914400 228600 L 914400 228600 L 38101 228600 L 0 190499 L 0 0 Z N">
						<draw:handle/>
					</draw:enhanced-geometry>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>-->
			</xsl:when>
			<!-- Snip and Round Single Corner Rectangle -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='snipRoundRect'">
				<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
										 draw:type="rectangle"
										 draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N"/>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
        <!--<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 914400 228600" 
											draw:extrusion-allowed="true" 
											draw:text-areas="11159 11159 895349 228600" 
											draw:glue-points="914400 114300 457200 228600 0 114300 457200 0" 
											draw:type="mso-spt100" 
											draw:enhanced-path="M 38101 0 L 876299 0 L 914400 38101 L 914400 228600 L 0 228600 L 0 38101 W 0 0 76202 76202 0 38101 38101 0 Z N">
						<draw:handle/>
					</draw:enhanced-geometry>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>-->
      </xsl:when>
      <!-- Round Single Corner Rectangle -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='round1Rect'">
        <draw:custom-shape draw:layer="layout" >
          <xsl:call-template name ="CreateShape">
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
										 draw:type="rectangle" 
										 draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
				<!--<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 914400 228600" 
											draw:extrusion-allowed="true" 
											draw:text-areas="0 0 903241 228600" 
											draw:glue-points="457200 0 0 114300 457200 228600 914400 114300" 
											draw:type="mso-spt100" 
											draw:enhanced-path="M 0 0 L 876299 0 W 838198 0 914400 76202 876299 0 914400 38101 L 914400 228600 L 0 228600 Z N">
						<draw:handle/>
					</draw:enhanced-geometry>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>-->
			</xsl:when>
			<!-- Round Same Side Corner Rectangle -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='round2SameRect'">
				<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" 
										 draw:type="rectangle" 
										 draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
				<!--<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 914400 914400" 
											draw:extrusion-allowed="true" 
											draw:text-areas="44637 44637 869763 914400" 
											draw:glue-points="914400 457200 457200 914400 0 457200 457200 0" 
											draw:type="mso-spt100" 
											draw:enhanced-path="M 152403 0 L 761997 0 W 609594 0 914400 304806 761997 0 914400 152403 L 914400 914400 L 0 914400 L 0 152403 W 0 0 304806 304806 0 152403 152403 0 Z N">
						<draw:handle/>
					</draw:enhanced-geometry>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>-->
			</xsl:when>
			<!-- Round Diagonal Corner Rectangle -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='round2DiagRect'">
				<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
          </xsl:call-template>
          <draw:enhanced-geometry svg:viewBox="0 0 21600 21600"
										 draw:type="rectangle"
										 draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N"/>
          <xsl:copy-of select="$varHyperLinksForShapes" />
        </draw:custom-shape>
        <!--<draw:custom-shape draw:layer="layout" >
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 2286000 1524000" 
						draw:extrusion-allowed="true" 
						draw:text-areas="74396 74396 2211604 1449604" 
						draw:glue-points="2286000 762000 1143000 1524000 0 762000 1143000 0" 
						draw:type="mso-spt100" 
						draw:enhanced-path="M 254005 0 L 2286000 0 L 2286000 1269995 W 1777990 1015990 2286000 1524000 2286000 1269995 2031995 1524000 L 0 1524000 L 0 254005 W 0 0 508010 508010 0 254005 254005 0 Z N">
						<draw:handle/>
					</draw:enhanced-geometry>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>-->
      </xsl:when>

      <!-- Flow chart shapes -->

      <!-- Flowchart: Process -->
      <xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartProcess'">
        <draw:custom-shape draw:layer="layout">
          <xsl:call-template name ="CreateShape">
            <xsl:with-param name="grID" select ="$GraphicId"/>
            <xsl:with-param name ="prID" select="$ParaId" />
            <!-- Extra parameter inserted by Vijayeta,For Bullets and numbering-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
            <!--End of definition of Extra parameter inserted by Vijayeta,For Bullets and numbering-->
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" draw:type="flowchart-process" draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- Flowchart: Alternate Process -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartAlternateProcess'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:path-stretchpoint-x="10800" draw:path-stretchpoint-y="10800" draw:text-areas="?f3 ?f4 ?f5 ?f6" draw:type="round-rectangle" draw:modifiers="3600" draw:enhanced-path="M ?f7 0 X 0 ?f8 L 0 ?f9 Y ?f7 21600 L ?f10 21600 X 21600 ?f9 L 21600 ?f8 Y ?f10 0 Z N">
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
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" draw:text-areas="5400 5400 16200 16200" draw:type="flowchart-decision" draw:enhanced-path="M 0 10800 L 10800 0 21600 10800 10800 21600 0 10800 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Data -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartInputOutput'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="12960 0 10800 0 2160 10800 8600 21600 10800 21600 19400 10800" draw:text-areas="4230 0 17370 21600" draw:type="flowchart-data" draw:enhanced-path="M 4230 0 L 21600 0 17370 21600 0 21600 4230 0 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Predefined Process-->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartPredefinedProcess'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:text-areas="2540 0 19060 21600" draw:type="flowchart-predefined-process" draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 Z N M 2540 0 L 2540 21600 N M 19060 0 L 19060 21600 N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Internal Storage -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartInternalStorage'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:text-areas="4230 4230 21600 21600" draw:type="flowchart-internal-storage" draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 Z N M 4230 0 L 4230 21600 N M 0 4230 L 21600 4230 N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Document -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartDocument'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 20320 21600 10800" draw:text-areas="0 0 21600 17360" draw:type="flowchart-document" draw:enhanced-path="M 0 0 L 21600 0 21600 17360 C 13050 17220 13340 20770 5620 21600 2860 21100 1850 20700 0 20120 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Multi document -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartMultidocument'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 19890 21600 10800" draw:text-areas="0 3600 18600 18009" draw:type="flowchart-multidocument" draw:enhanced-path="M 0 3600 L 1500 3600 1500 1800 3000 1800 3000 0 21600 0 21600 14409 20100 14409 20100 16209 18600 16209 18600 18009 C 11610 17893 11472 20839 4833 21528 2450 21113 1591 20781 0 20300 Z N M 1500 3600 F L 18600 3600 18600 16209 N M 3000 1800 F L 20100 1800 20100 14409 N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Terminator -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartTerminator'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" draw:text-areas="1060 3180 20540 18420" draw:type="flowchart-terminator" draw:enhanced-path="M 3470 21600 X 0 10800 3470 0 L 18130 0 X 21600 10800 18130 21600 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Preparation -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartPreparation'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" draw:text-areas="4350 0 17250 21600" draw:type="flowchart-preparation" draw:enhanced-path="M 4350 0 L 17250 0 21600 10800 17250 21600 4350 21600 0 10800 4350 0 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Manual Input -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartManualInput'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 2150 0 10800 10800 19890 21600 10800" draw:text-areas="0 4300 21600 21600" draw:type="flowchart-manual-input" draw:enhanced-path="M 0 4300 L 21600 0 21600 21600 0 21600 0 4300 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Manual Operation -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartManualOperation'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 2160 10800 10800 21600 19440 10800" draw:text-areas="4350 0 17250 21600" draw:type="flowchart-manual-operation" draw:enhanced-path="M 0 0 L 21600 0 17250 21600 4350 21600 0 0 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Connector -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartConnector'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 3160 3160 0 10800 3160 18440 10800 21600 18440 18440 21600 10800 18440 3160" draw:text-areas="3180 3180 18420 18420" draw:type="flowchart-connector" draw:enhanced-path="U 10800 10800 10800 10800 0 23592960 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Off-page Connector -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartOffpageConnector'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" draw:text-areas="0 0 21600 17150" draw:type="flowchart-off-page-connector" draw:enhanced-path="M 0 0 L 21600 0 21600 17150 10800 21600 0 17150 0 0 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Card -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartPunchedCard'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" draw:text-areas="0 4300 21600 21600" draw:type="flowchart-card" draw:enhanced-path="M 4300 0 L 21600 0 21600 21600 0 21600 0 4300 4300 0 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Punched Tape -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartPunchedTape'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 2020 0 10800 10800 19320 21600 10800" draw:text-areas="0 4360 21600 17240" draw:type="flowchart-punched-tape" draw:enhanced-path="M 0 2230 C 820 3990 3410 3980 5370 4360 7430 4030 10110 3890 10690 2270 11440 300 14200 160 16150 0 18670 170 20690 390 21600 2230 L 21600 19420 C 20640 17510 18320 17490 16140 17240 14710 17370 11310 17510 10770 19430 10150 21150 7380 21290 5290 21600 3220 21250 610 21130 0 19420 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Summing Junction -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartSummingJunction'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 3160 3160 0 10800 3160 18440 10800 21600 18440 18440 21600 10800 18440 3160" draw:text-areas="3100 3100 18500 18500" draw:type="flowchart-summing-junction" draw:enhanced-path="U 10800 10800 10800 10800 0 23592960 Z N M 3100 3100 L 18500 18500 N M 3100 18500 L 18500 3100 N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Or -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartOr'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 3160 3160 0 10800 3160 18440 10800 21600 18440 18440 21600 10800 18440 3160" draw:text-areas="3100 3100 18500 18500" draw:type="flowchart-or" draw:enhanced-path="U 10800 10800 10800 10800 0 23592960 Z N M 0 10800 L 21600 10800 N M 10800 0 L 10800 21600 N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Collate -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartCollate'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 10800 10800 10800 21600" draw:text-areas="5400 5400 16200 16200" draw:type="flowchart-collate" draw:enhanced-path="M 0 0 L 21600 21600 0 21600 21600 0 0 0 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Sort -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartSort'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
					<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:text-areas="5400 5400 16200 16200" draw:type="flowchart-sort" draw:enhanced-path="M 0 10800 L 10800 0 21600 10800 10800 21600 Z N M 0 10800 L 21600 10800 N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Extract -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartExtract'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 5400 10800 10800 21600 16200 10800" draw:text-areas="5400 10800 16200 21600" draw:type="flowchart-extract" draw:enhanced-path="M 10800 0 L 21600 21600 0 21600 10800 0 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Merge-->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartMerge'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 5400 10800 10800 21600 16200 10800" draw:text-areas="5400 0 16200 10800" draw:type="flowchart-merge" draw:enhanced-path="M 0 0 L 21600 0 10800 21600 0 0 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Stored Data -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartOnlineStorage'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 18000 10800" draw:text-areas="3600 0 18000 21600" draw:type="flowchart-stored-data" draw:enhanced-path="M 3600 21600 X 0 10800 3600 0 L 21600 0 X 18000 10800 21600 21600 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Delay -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartDelay'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" draw:text-areas="0 3100 18500 18500" draw:type="flowchart-delay" draw:enhanced-path="M 10800 0 X 21600 10800 10800 21600 L 0 21600 0 0 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Sequential Access Storage -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartMagneticTape'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" draw:text-areas="3100 3100 18500 18500" draw:type="flowchart-sequential-access" draw:enhanced-path="M 20980 18150 L 20980 21600 10670 21600 C 4770 21540 0 16720 0 10800 0 4840 4840 0 10800 0 16740 0 21600 4840 21600 10800 21600 13520 20550 16160 18670 18170 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Direct Access Storage -->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartMagneticDrum'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 14800 10800 21600 10800" draw:text-areas="3400 0 14800 21600" draw:type="flowchart-direct-access-storage" draw:enhanced-path="M 18200 0 X 21600 10800 18200 21600 L 3400 21600 X 0 10800 3400 0 Z N M 18200 0 X 14800 10800 18200 21600 N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Magnetic Disk-->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartMagneticDisk'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 6800 10800 0 0 10800 10800 21600 21600 10800" draw:text-areas="0 6800 21600 18200" draw:type="flowchart-magnetic-disk" draw:enhanced-path="M 0 3400 Y 10800 0 21600 3400 L 21600 18200 Y 10800 21600 0 18200 Z N M 0 3400 Y 10800 6800 21600 3400 N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
			<!-- FlowChart: Display-->
			<xsl:when test ="p:spPr/a:prstGeom/@prst='flowChartDisplay'">
				<draw:custom-shape draw:layer="layout">
					<xsl:call-template name ="CreateShape">
						<xsl:with-param name="grID" select ="$GraphicId"/>
						<xsl:with-param name ="prID" select="$ParaId" />
					</xsl:call-template>
				<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:glue-points="10800 0 0 10800 10800 21600 21600 10800" draw:text-areas="3600 0 17800 21600" draw:type="flowchart-display" draw:enhanced-path="M 3600 0 L 17800 0 X 21600 10800 17800 21600 L 3600 21600 0 10800 Z N"/>
					<xsl:copy-of select="$varHyperLinksForShapes" />
				</draw:custom-shape>
			</xsl:when>
		 				 
		</xsl:choose>
	</xsl:template>
	<!-- Draw Shape reading values from pptx p:spPr-->
	<xsl:template name ="CreateShape">
		<xsl:param name ="grID" />
		<xsl:param name ="prID" />
    <!-- Addition of a parameter,by Vijayets ,for bullets and numbering in shapes-->
    <xsl:param name="SlideRelationId"/>

		<xsl:attribute name ="draw:style-name">
			<xsl:value-of select ="$grID"/>
		</xsl:attribute>
		<xsl:attribute name ="draw:text-style-name">
			<xsl:value-of select ="$prID"/>
		</xsl:attribute>

		<xsl:for-each select ="p:spPr/a:xfrm">
			<xsl:attribute name ="svg:width">
				<xsl:call-template name="ConvertEmu">
					<xsl:with-param name="length" select="a:ext/@cx"/>
					<xsl:with-param name="unit">cm</xsl:with-param>
				</xsl:call-template>
			</xsl:attribute>

			<xsl:attribute name ="svg:height">
				<xsl:call-template name="ConvertEmu">
					<xsl:with-param name="length" select="a:ext/@cy"/>
					<xsl:with-param name="unit">cm</xsl:with-param>
				</xsl:call-template>
			</xsl:attribute>

			<xsl:attribute name ="svg:x">
				<xsl:call-template name="ConvertEmu">
					<xsl:with-param name="length" select="a:off/@x"/>
					<xsl:with-param name="unit">cm</xsl:with-param>
				</xsl:call-template>
			</xsl:attribute>

			<xsl:attribute name ="svg:y">
				<xsl:call-template name="ConvertEmu">
					<xsl:with-param name="length" select="a:off/@y"/>
					<xsl:with-param name="unit">cm</xsl:with-param>
				</xsl:call-template>
			</xsl:attribute>
		</xsl:for-each>
		<xsl:choose>
			<xsl:when test ="(p:nvSpPr/p:cNvPr/@name[contains(., 'TextBox')] or p:nvSpPr/p:cNvPr/@name[contains(., 'Text Box')]) and not(p:nvSpPr/p:cNvPr/@name[contains(., 'TextBox Custom')])">
				<draw:text-box>
					<xsl:call-template name ="AddShapeText">
						<xsl:with-param name ="prID" select ="$prID" />
            <!-- Addition of a parameter,by vijayeta,for bullets and numbering in shapes-->
            <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
					</xsl:call-template> 
				</draw:text-box>
			</xsl:when>
			<xsl:otherwise>
				<xsl:call-template name ="AddShapeText">
					<xsl:with-param name ="prID" select ="$prID" />
          <!-- Addition of a parameter,by vijayeta,for bullets and numbering in shapes-->
          <xsl:with-param name="SlideRelationId" select ="$SlideRelationId" />
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!-- Draw line -->
	<xsl:template name ="DrawLine">
		<xsl:param name ="grID" />

		<xsl:attribute name ="draw:style-name">
			<xsl:value-of select ="$grID"/>
		</xsl:attribute>

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
			<!--<xsl:attribute name ="svg:x1">
				<xsl:value-of select ="$xCenter - m:cos($angle) * $cxBy2 + m:sin($angle) * $cyBy2"/>
			</xsl:attribute>
			<xsl:attribute name ="svg:y1">
				<xsl:value-of select ="$yCenter - m:sin($angle) * $cxBy2 - m:cos($angle) * $cyBy2"/>
			</xsl:attribute>
			<xsl:attribute name ="svg:x2">
				<xsl:value-of select ="$xCenter + m:cos($angle) * $cxBy2 - m:sin($angle) * $cyBy2"/>
			</xsl:attribute>
			<xsl:attribute name ="svg:y2">
				<xsl:value-of select ="$yCenter + m:sin($angle) * $cxBy2 + m:cos($angle) * $cyBy2"/>
			</xsl:attribute>-->

		</xsl:for-each>

	</xsl:template>
	<!-- Add text to the shape -->
	<xsl:template name ="AddShapeText">
		<xsl:param name ="prID" />
    <xsl:param name="SlideRelationId"/>
    <xsl:variable name ="SlideNumber" select ="substring-before(substring-after($SlideRelationId,'ppt/slides/_rels/'),'.xml.rels')"/>
    <!--concat($SlideNumber,'textboxshape_List',position()-->
    <xsl:for-each select ="p:txBody">
      <!--code inserted by Vijayeta,InsertStyle For Bullets and Numbering-->
      <xsl:variable name ="listStyleName">
        <xsl:value-of select ="concat($SlideNumber,'textboxshape_List',position())"/>
      </xsl:variable>
      <xsl:for-each select ="a:p">
        <!--Code Inserted by Vijayeta for Bullets And Numbering
                  check for levels and then depending on the condition,insert bullets,Layout or Master properties-->
        <xsl:if test ="a:pPr/a:buChar or a:pPr/a:buAutoNum or a:pPr/a:buBlip">
          <xsl:call-template name ="insertBulletsNumbersoox2odf">
            <xsl:with-param name ="listStyleName" select ="concat($listStyleName,position())"/>
            <xsl:with-param name ="ParaId" select ="$prID"/>
          </xsl:call-template>
        </xsl:if>
        <xsl:if test ="not(a:pPr/a:buChar) and not(a:pPr/a:buAutoNum)and not(a:pPr/a:buBlip)">
			<text:p >
				<xsl:attribute name ="text:style-name">
					<xsl:value-of select ="concat($prID,position())"/>
				</xsl:attribute>
				<xsl:for-each select ="node()">
					<xsl:if test ="name()='a:r'">
						<text:span text:style-name="{generate-id()}">
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
						</text:span>
					</xsl:if >
					<xsl:if test ="name()='a:br'">
						<text:line-break/>
					</xsl:if>
				</xsl:for-each>
			</text:p>
          <!--Code inserted by vijayeta,for Bullets and Numbering If Bullets are present-->
        </xsl:if >
      </xsl:for-each>
       <!--If no bullets are present or default bullets-->
           
		</xsl:for-each>

	</xsl:template>
	<!-- Generate autometic styles in contet.xsl for graphic properties-->
	<xsl:template name ="InsertStylesForGraphicProperties"  >
		<xsl:for-each select ="document('ppt/presentation.xml')/p:presentation/p:sldIdLst/p:sldId">
			<xsl:variable name ="SlideId">
				<xsl:value-of  select ="concat(concat('slide',position()),'.xml')" />
			</xsl:variable>
			<!-- Graphic properties for shapes with p:sp nodes-->
			<xsl:for-each select ="document(concat('ppt/slides/',$SlideId))/p:sld/p:cSld/p:spTree/p:sp">
				<!--Check for shape or texbox -->
				<!--<xsl:if test = "p:style or p:nvSpPr/p:cNvPr/@name[contains(., 'TextBox')] or 
										   p:style or p:nvSpPr/p:cNvPr/@name[contains(., 'Text Box')] or 
										   p:nvSpPr/p:cNvPr/@name[contains(., 'Rectangle')] or 
										   p:nvSpPr/p:cNvPr/@name[contains(., 'Line')]">-->
					<!-- Generate graphic properties ID-->
					<xsl:variable  name ="GraphicId">
						<xsl:value-of select ="concat('s',substring($SlideId,6,string-length($SlideId)-9) ,concat('gr',position()))"/>
					</xsl:variable>

					<style:style style:family="graphic" style:parent-style-name="standard">
						<xsl:attribute name ="style:name">
							<xsl:value-of select ="$GraphicId"/>
						</xsl:attribute >
						<style:graphic-properties>

							<!-- FILL -->
							<xsl:call-template name ="Fill" />

							<!-- LINE COLOR -->
							<xsl:call-template name ="LineColor" />

							<!-- LINE STYLE -->
							<xsl:call-template name ="LineStyle"/>

							<!-- TEXT ALIGNMENT -->
							<xsl:call-template name ="TextLayout" />

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
			<xsl:for-each select ="document(concat('ppt/slides/',$SlideId))/p:sld/p:cSld/p:spTree/p:cxnSp">
				<!-- Generate graphic properties ID-->
				<xsl:variable  name ="GraphicId">
					<xsl:value-of select ="concat('s',substring($SlideId,6,string-length($SlideId)-9) ,concat('grLine',position()))"/>
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
		</xsl:for-each>
	</xsl:template>
	<!-- Get fill color for shape-->
	<xsl:template name="Fill">
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
				<xsl:attribute name ="draw:fill">
					<xsl:value-of select="'solid'" />
				</xsl:attribute>
				<!-- Standard color-->
				<xsl:if test ="p:spPr/a:solidFill/a:srgbClr/@val">
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
					<xsl:attribute name ="draw:fill-color">
						<xsl:call-template name ="getColorCode">
							<xsl:with-param name ="color">
								<xsl:value-of select="p:spPr/a:solidFill/a:schemeClr/@val"/>
							</xsl:with-param>
							<xsl:with-param name ="lumMod">
								<xsl:value-of select="p:spPr/a:solidFill/a:schemeClr/a:lumMod/@val"/>
							</xsl:with-param>
							<xsl:with-param name ="lumOff">
								<xsl:value-of select="p:spPr/a:solidFill/a:schemeClr/a:lumOff/@val"/>
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
			
			<xsl:otherwise>
			<!--Fill refernce-->
			<xsl:if test ="p:style/a:fillRef">
				<xsl:attribute name ="draw:fill">
					<xsl:value-of select="'solid'" />
				</xsl:attribute>
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
				<xsl:if test ="p:style/a:fillRef//a:schemeClr/@val">
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
			</xsl:if>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!-- Get border color for shape -->
	<xsl:template name ="LineColor">

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
				<xsl:if test ="not( (p:spPr/a:prstGeom/@prst='flowChartInternalStorage') or
									(p:spPr/a:prstGeom/@prst='flowChartPredefinedProcess') or
									(p:spPr/a:prstGeom/@prst='flowChartSummingJunction') or
									(p:spPr/a:prstGeom/@prst='flowChartOr') )">
				<xsl:if test ="p:style/a:lnRef">
					<xsl:attribute name ="draw:stroke">
						<xsl:value-of select="'solid'" />
					</xsl:attribute>
					<!--Standard color for border-->
					<xsl:if test ="p:style/a:lnRef/a:srgbClr/@val">
						<xsl:attribute name ="svg:stroke-color">
							<xsl:value-of select="concat('#',p:style/a:lnRef/a:srgbClr/@val)"/>
						</xsl:attribute>

						<!--Shade percentage-->
						<!--
					<xsl:if test="p:style/a:lnRef/a:srgbClr/a:shade/@val">
						<xsl:variable name ="shade">
							<xsl:value-of select ="p:style/a:lnRef/a:srgbClr/a:shade/@val"/>
						</xsl:variable>
						-->
						<!--<xsl:if test="($shade != '') or ($shade != 0)">
							<xsl:attribute name ="svg:stroke-opacity">
								<xsl:value-of select="concat(($shade div 1000), '%')"/>
							</xsl:attribute>
						</xsl:if>-->
						<!--
					</xsl:if>-->
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
						<!--Shade percentage -->
						<!--<xsl:if test="p:style/a:lnRef/a:schemeClr/a:shade/@val">
						<xsl:variable name ="shade">
							<xsl:value-of select ="p:style/a:lnRef/a:schemeClr/a:shade/@val"/>
						</xsl:variable>
						-->
						<!--<xsl:if test="($shade != '') or ($shade != 0)">
							<xsl:attribute name ="svg:stroke-opacity">
								<xsl:value-of select="concat(($shade div 1000), '%')"/>
							</xsl:attribute>
						</xsl:if>-->
						<!--
					</xsl:if>-->
					</xsl:if>
				</xsl:if>
				</xsl:if>
			</xsl:otherwise> 
		</xsl:choose>
	</xsl:template>
	<!-- Get line styles for shape -->
	<xsl:template name ="LineStyle">
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
		</xsl:for-each>

		<!-- Line Dash property-->
		<xsl:for-each select ="p:spPr/a:ln">
			<xsl:if test ="not(a:noFill)">
				<xsl:choose>
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
				<xsl:attribute name ="draw:marker-start-width">
					<xsl:call-template name ="getArrowSize">
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
				<xsl:attribute name ="draw:marker-end-width">
					<xsl:call-template name ="getArrowSize">
						<xsl:with-param name ="w" select ="@w" />
						<xsl:with-param name ="len" select ="@len" />
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>
		</xsl:for-each>
	</xsl:template>
	<!-- Get arrow size -->
	<xsl:template name ="getArrowSize">
		<xsl:param name ="w" />
		<xsl:param name ="len" />

		<xsl:choose>
			<xsl:when test ="($w = 'sm') and ($len = 'sm')">
				<xsl:value-of select ="concat($sm-sm,'cm')"/>
			</xsl:when>
			<xsl:when test ="($w = 'sm') and ($len = 'med')">
				<xsl:value-of select ="concat($sm-med,'cm')"/>
			</xsl:when>
			<xsl:when test ="($w = 'sm') and ($len = 'lg')">
				<xsl:value-of select ="concat($sm-lg,'cm')"/>
			</xsl:when>
			<xsl:when test ="($w = 'med') and ($len = 'sm')">
				<xsl:value-of select ="concat($med-sm,'cm')" />
			</xsl:when>
			<xsl:when test ="($w = 'med') and ($len = 'lg')">
				<xsl:value-of select ="concat($med-lg,'cm')" />
			</xsl:when>
			<xsl:when test ="($w = 'lg') and ($len = 'sm')">
				<xsl:value-of select ="concat($lg-sm,'cm')" />
			</xsl:when>
			<xsl:when test ="($w = 'lg') and ($len = 'med')">
				<xsl:value-of select ="concat($lg-med,'cm')" />
			</xsl:when>
			<xsl:when test ="($w = 'lg') and ($len = 'lg')">
				<xsl:value-of select ="concat($lg-lg,'cm')" />
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select ="concat($med-med,'cm')"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!-- Get text layout for shapes -->
	<xsl:template name ="TextLayout">
		<xsl:for-each select="p:txBody">
			<xsl:attribute name ="draw:textarea-horizontal-align">
				<xsl:choose>
					<!--BugFixed-1709909-->
					<xsl:when test ="p:txBody/a:bodyPr/@anchorCtr= 1">
						<xsl:value-of select ="'center'"/>
					</xsl:when>
					<xsl:when test ="p:txBody/a:bodyPr/@anchorCtr= 0">
						<xsl:value-of select ="'justify'"/>
					</xsl:when>
					<!--Justify alignment-->
					<xsl:otherwise>
						<xsl:value-of select ="'justify'"/>
					</xsl:otherwise>
				</xsl:choose>
				<!-- Center alignment
					<xsl:when test ="(a:p/a:pPr/@algn ='ctr') or (a:bodyPr/@anchorCtr = '1')">
						<xsl:value-of select ="'center'"/>
					</xsl:when>-->
				<!-- Right alignment
					<xsl:when test ="a:p/a:pPr/@algn ='r'">
						<xsl:value-of select ="'right'"/>
					</xsl:when>-->
				<!--Justify alignment
					<xsl:when test ="(a:p/a:pPr/@algn ='just') or (a:bodyPr/@anchorCtr = '0')">
						<xsl:value-of select ="'justify'"/>
					</xsl:when>-->
				<!-- Left alignment
					<xsl:otherwise>
						<xsl:value-of select ="'left'"/>
					</xsl:otherwise>-->

			</xsl:attribute>
			<xsl:attribute name ="draw:textarea-vertical-align">
				<xsl:choose>
					<!-- Middle alignment-->
					<xsl:when test ="a:bodyPr/@anchor ='ctr'">
						<xsl:value-of select ="'middle'"/>
					</xsl:when>
					<!-- Bottom alignment-->
					<xsl:when test ="a:bodyPr/@anchor ='b'">
						<xsl:value-of select ="'bottom'"/>
					</xsl:when>
					<!-- Top alignment -->
					<xsl:otherwise>
						<xsl:value-of select ="'top'"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:attribute>

			<!--fo:padding-->
			<xsl:if test ="a:bodyPr/@lIns or (a:bodyPr/@wrap='square') or (a:p/a:pPr/@fontAlgn='auto')">
				<xsl:attribute name ="fo:padding-left">
					<xsl:call-template name="getPadding">
						<xsl:with-param name="length" select="a:bodyPr/@lIns"/>
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>

			<xsl:if test ="a:bodyPr/@tIns or (a:bodyPr/@wrap='square') or (a:p/a:pPr/@fontAlgn='auto')">
				<xsl:attribute name ="fo:padding-top">
					<xsl:call-template name="getPadding">
						<xsl:with-param name="length" select="a:bodyPr/@tIns"/>
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>

			<xsl:if test ="a:bodyPr/@rIns or (a:bodyPr/@wrap='square') or (a:p/a:pPr/@fontAlgn='auto')">
				<xsl:attribute name ="fo:padding-right">
					<xsl:call-template name="getPadding">
						<xsl:with-param name="length" select="a:bodyPr/@rIns"/>
					</xsl:call-template>
				</xsl:attribute>
			</xsl:if>

			<xsl:if test ="a:bodyPr/@bIns or (a:bodyPr/@wrap='square') or (a:p/a:pPr/@fontAlgn='auto')">
				<xsl:attribute name ="fo:padding-bottom">
					<xsl:call-template name="getPadding">
						<xsl:with-param name="length" select="a:bodyPr/@bIns"/>
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
			</xsl:choose>

			<xsl:if test ="( (a:bodyPr/a:spAutoFit) or (a:bodyPr/@wrap='square') )">
				<xsl:attribute name="draw:auto-grow-height">
					<xsl:value-of select ="'true'"/>
				</xsl:attribute>
			</xsl:if>
			<xsl:if test ="not(a:bodyPr/a:spAutoFit)">
				<xsl:attribute name="draw:auto-grow-height">
					<xsl:value-of select ="'false'"/>
				</xsl:attribute>
				<xsl:attribute name="draw:auto-grow-width">
					<xsl:value-of select ="'false'"/>
				</xsl:attribute>
			</xsl:if>
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
		
</xsl:stylesheet >
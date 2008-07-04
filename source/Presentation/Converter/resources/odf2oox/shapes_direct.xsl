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
<xsl:stylesheet version="1.0" 
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"   
  xmlns:odf="urn:odf"
  xmlns:pzip="urn:cleverage:xmlns:post-processings:zip"  
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" 
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:page="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" 
  xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  exclude-result-prefixes="odf style text number draw page r presentation fo script xlink svg">
  

  <!-- Shape constants-->
  <xsl:variable name ="dot">
    <xsl:value-of select ="'0.07'"/>
  </xsl:variable>
  <xsl:variable name ="dash">
    <xsl:value-of select ="'0.282'"/>
  </xsl:variable>
  <xsl:variable name ="longDash">
    <xsl:value-of select ="'0.564'"/>
  </xsl:variable>
  <xsl:variable name ="distance">
    <xsl:value-of select ="'0.211'"/>
  </xsl:variable>
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

   <!-- Generate nvPrId-->
   <xsl:template name ="shapeGetnvPrId">
		<xsl:param name ="spId"/>
		<xsl:variable name ="varSpid">
			<xsl:for-each select ="parent::node()">
				<xsl:for-each select ="node()">
					<xsl:variable name ="nvPrId">
						<xsl:value-of select ="position()"/>
					</xsl:variable>
					<xsl:variable name ="drawId">
						<xsl:if test ="$spId =@draw:id">
              <xsl:call-template name="getShapePosTemp">
                <xsl:with-param name="var_pos" select="position()"/>
              </xsl:call-template>
						</xsl:if >
					</xsl:variable >
					<xsl:variable name ="paraId">
						<xsl:for-each select =".//text:p">
							<xsl:if test ="$spId =@text:id">
								<xsl:value-of select ="position()"/>
							</xsl:if>
						</xsl:for-each>
					</xsl:variable>
					<xsl:choose>
						<xsl:when test ="$paraId!=''">
							<xsl:value-of select ="$nvPrId +1 "/>
						</xsl:when>
						<xsl:when test ="$paraId='' and $drawId !=''">
							<xsl:value-of select ="$nvPrId +1 "/>
						</xsl:when>
					</xsl:choose>
				</xsl:for-each>
			</xsl:for-each >
		</xsl:variable>
     <xsl:if test="$varSpid !=''">
		<xsl:value-of select ="$varSpid"/>
     </xsl:if>
     <xsl:if test="$varSpid =''">
       <xsl:value-of select ="position()"/>
     </xsl:if>
	</xsl:template>
  <!-- Template for Shapes in direct conversion -->
  <xsl:template name ="shapes">
    <xsl:param name ="fileName" />
    <xsl:param name ="styleName" />
    <xsl:param name ="var_pos" />
    <xsl:param name ="NvPrId" />
    <xsl:param name ="grpFlag" />
    <xsl:param name ="UniqueId" />
         <xsl:choose>
        <xsl:when test ="name()='draw:rect'">
          <xsl:variable name="shapeCount" select="position()"/>
          <xsl:for-each select=".">
            <xsl:call-template name ="CreateShape">
              <xsl:with-param name ="fileName" select ="$fileName" />
              <xsl:with-param name ="shapeName" select="'Rectangle '" />
              <xsl:with-param name ="shapeCount" select="$var_pos" />
              <xsl:with-param name ="grpFlag" select="$grpFlag" />
              <xsl:with-param name ="UniqueId" select="generate-id()" />
            </xsl:call-template>
          </xsl:for-each>
         </xsl:when>
        <xsl:when test ="name()='draw:ellipse' or name()='draw:circle'">
          <xsl:if test="not(@draw:kind)">
            <xsl:variable name="shapeCount" select="position()"/>
            <xsl:for-each select=".">
              <xsl:call-template name="CreateShape">
				<xsl:with-param name ="fileName" select ="$fileName" />
                <xsl:with-param name="shapeName" select="'Oval '" />
                <xsl:with-param name ="shapeCount" select="$var_pos" />
                <xsl:with-param name ="grpFlag" select="$grpFlag" />
                <xsl:with-param name ="UniqueId" select="generate-id()" />
                <!--<xsl:with-param name ="NvPrId" select="$NvPrId" />-->
              </xsl:call-template>
            </xsl:for-each>
          </xsl:if>
        </xsl:when>
        <xsl:when test ="name()='draw:custom-shape'">

          <xsl:variable name ="shapeName">
            <xsl:call-template name="tmpgetCustShapeType"/>
            </xsl:variable>
          <!-- code for the bug 1746360 -->
          <!--<xsl:for-each select=".">-->
            <xsl:if test="$shapeName != ''">
              <xsl:variable name="shapeCount" select="position()"/>
              <xsl:call-template name ="CreateShape">
                <xsl:with-param name ="fileName" select ="$fileName" />
                <xsl:with-param name ="shapeName" select="$shapeName" />
                <xsl:with-param name ="customShape" select ="'true'" />
                <xsl:with-param name ="shapeCount" select="$var_pos" />
                <xsl:with-param name ="grpFlag" select="$grpFlag" />
                <xsl:with-param name ="UniqueId" select="generate-id()" />
                <!--<xsl:with-param name ="NvPrId" select="$NvPrId" />-->
              </xsl:call-template>
            </xsl:if>
			<xsl:if test="draw:enhanced-geometry/@draw:type='fontwork-plain-text' or
						  draw:enhanced-geometry/@draw:type='fontwork-wave' or
						  draw:enhanced-geometry/@draw:type='fontwork-fade-up-and-right' or 
						  draw:enhanced-geometry/@draw:type='fontwork-inflate' or
						  draw:enhanced-geometry/@draw:type='fontwork-curve-up' or
						  draw:enhanced-geometry/@draw:type='fontwork-slant-up' or
						  draw:enhanced-geometry/@draw:type='fontwork-fade-up' or
						  draw:enhanced-geometry/@draw:type='fontwork-chevron-up' or
						  draw:enhanced-geometry/@draw:type='fontwork-triangle-down' or
						  draw:enhanced-geometry/@draw:type='fontwork-curve-down' or
						  draw:enhanced-geometry/@draw:type='fontwork-arch-down-pour' or 
						  draw:enhanced-geometry/@draw:type='fontwork-arch-up-curve' or
						  draw:enhanced-geometry/@draw:type='fontwork-fade-down' or
						  draw:enhanced-geometry/@draw:type='fontwork-triangle-up' or
						  draw:enhanced-geometry/@draw:type='fontwork-fade-right' or 
						  draw:enhanced-geometry/@draw:type='fontwork-stop' or
						  draw:enhanced-geometry/@draw:type='fontwork-chevron-down' or
						  draw:enhanced-geometry/@draw:type='fontwork-arch-up-pour' or
						  draw:enhanced-geometry/@draw:type='fontwork-circle-pour' or
						  draw:enhanced-geometry/@draw:enhanced-path='V 0 0 21600 21600 ?f2 ?f3 ?f2 ?f4 N'">
				<xsl:call-template name ="CreateShape">
					<xsl:with-param name ="fileName" select ="'content.xml'"/>
					<xsl:with-param name ="shapeName" select="'TextBox '" />
					<xsl:with-param name ="shapeCount" select="$var_pos" />
				</xsl:call-template>
				<!--</xsl:for-each>-->
			</xsl:if>
        </xsl:when>
        <xsl:when test ="name()='draw:line'">
        <!--<xsl:for-each select=".">-->
          <xsl:call-template name ="drawLine">
            <xsl:with-param name ="fileName" select ="$fileName" />
            <xsl:with-param name ="connectorType" select ="'Straight Connector '" />
            <xsl:with-param name ="shapeCount" select="$var_pos" />
            <xsl:with-param name ="grpFlag" select="$grpFlag" />
            <!--<xsl:with-param name ="NvPrId" select="$NvPrId" />-->
          </xsl:call-template>
        <!--</xsl:for-each>-->
        </xsl:when>
        <xsl:when test ="name()='draw:connector'">
          <xsl:variable name ="type">
            <xsl:choose>
              <xsl:when test ="@draw:type='line'">
                <xsl:value-of select ="'Straight Arrow Connector '"/>
              </xsl:when>
		<!-- Bug Fix for StraightLine Connector by Mathi on 6thSep 2007-->
			  <xsl:when test ="@draw:type='lines' and not(@draw:line-skew)">
				  <xsl:value-of select="'Straight Arrow Connector '"/>		
			  </xsl:when>
			  <!--End of Bug Fix Code-->
	       <xsl:when test ="@draw:type='curve'">
                <xsl:value-of select ="'Curved Connector '"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select ="'Elbow Connector '"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <!--<xsl:for-each select=".">-->
            <xsl:call-template name ="drawLine">
              <xsl:with-param name ="fileName" select ="$fileName" />
              <xsl:with-param name ="connectorType" select ="$type" />
            <xsl:with-param name ="shapeCount" select="$var_pos" />
            <xsl:with-param name ="grpFlag" select="$grpFlag" />
              <!--<xsl:with-param name ="NvPrId" select="$NvPrId" />-->
            </xsl:call-template>
          <!--</xsl:for-each>-->
        </xsl:when>
        <!--<xsl:when test ="draw:g">
          <xsl:for-each select="node()">
            <xsl:call-template name="shapes">
              <xsl:with-param name ="fileName" select ="$fileName" />
            </xsl:call-template >
          </xsl:for-each>
        </xsl:when>-->
        </xsl:choose>
  </xsl:template>
  <!-- Create p:sp node for shape -->
  <xsl:template name ="CreateShape">
    <xsl:param name ="fileName" />
    <xsl:param name ="shapeName" />
    <xsl:param name ="customShape" />
    <xsl:param name ="shapeCount" />
    <xsl:param name ="NvPrId" />
    <xsl:param name ="grpFlag" />
    <xsl:param name ="UniqueId" />
    
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr>
          <xsl:attribute name ="id">
            <xsl:value-of select ="$shapeCount+1"/>
          </xsl:attribute>
          <xsl:attribute name ="name">
            <xsl:value-of select ="concat($shapeName, position())"/>
          </xsl:attribute>
          <!-- ADDED BY Sateesh - HYPER LINKS FOR SHAPES-->
          <xsl:variable name="PostionCount">
            <xsl:value-of select="$shapeCount"/>
          </xsl:variable>
          <xsl:variable name="ShapeType">
            <xsl:choose>
              <xsl:when test="$shapeName ='Rectangle '">
                <xsl:value-of select="'RectAtachFileId'"/>
              </xsl:when>
              <xsl:when test="$shapeName = 'TextBox '">
                <xsl:value-of select="'TxtBoxAtchFileId'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'ShapeFileId'"/>
              </xsl:otherwise>
            </xsl:choose>
           </xsl:variable>
          <xsl:if test="$grpFlag !='true'">
          <xsl:if test="count(office:event-listeners/presentation:event-listener) = 1">
            <xsl:if test="office:event-listeners">
              <xsl:for-each select ="office:event-listeners/presentation:event-listener">
                <xsl:if test="@script:event-name[contains(.,'dom:click')]">
                 
                    <xsl:choose>
                      <!--Go to previous slide-->
                      <xsl:when test="@presentation:action[ contains(.,'previous-page')]">
                        <a:hlinkClick>
                        <xsl:attribute name="action">
                          <xsl:value-of select="'ppaction://hlinkshowjump?jump=previousslide'"/>
                        </xsl:attribute>
                          <xsl:call-template name="tmpSetRid">
                            <xsl:with-param name="ShapeType" select="$ShapeType"/>
                            <xsl:with-param name="PostionCount" select="$PostionCount"/>
                          </xsl:call-template>
                        </a:hlinkClick>
                      </xsl:when>
                      <!--Go to Next slide-->
                      <xsl:when test="@presentation:action[ contains(.,'next-page')]">
                        <a:hlinkClick>
                          <xsl:attribute name="action">
                            <xsl:value-of select="'ppaction://hlinkshowjump?jump=nextslide'"/>
                          </xsl:attribute>
                          <xsl:call-template name="tmpSetRid">
                            <xsl:with-param name="ShapeType" select="$ShapeType"/>
                            <xsl:with-param name="PostionCount" select="$PostionCount"/>
                          </xsl:call-template>
                        </a:hlinkClick>
                        </xsl:when>
                      <!--Go to First slide-->
                      <xsl:when test="@presentation:action[ contains(.,'first-page')]">
                        <a:hlinkClick>
                          <xsl:attribute name="action">
                            <xsl:value-of select="'ppaction://hlinkshowjump?jump=firstslide'"/>
                          </xsl:attribute>
                          <xsl:call-template name="tmpSetRid">
                            <xsl:with-param name="ShapeType" select="$ShapeType"/>
                            <xsl:with-param name="PostionCount" select="$PostionCount"/>
                          </xsl:call-template>
                          </a:hlinkClick>
                        </xsl:when>
                      <!--Go to Last slide-->
                      <xsl:when test="@presentation:action[ contains(.,'last-page')]">
                        <a:hlinkClick>
                        <xsl:attribute name="action">
                          <xsl:value-of select="'ppaction://hlinkshowjump?jump=lastslide'"/>
                        </xsl:attribute>
                        <xsl:call-template name="tmpSetRid">
                          <xsl:with-param name="ShapeType" select="$ShapeType"/>
                          <xsl:with-param name="PostionCount" select="$PostionCount"/>
                        </xsl:call-template>
                        </a:hlinkClick>
                      </xsl:when>
                      <!--End Presentation-->
                      <xsl:when test="@presentation:action[ contains(.,'stop')]">
                        <a:hlinkClick>
                          <xsl:attribute name="action">
                            <xsl:value-of select="'ppaction://hlinkshowjump?jump=endshow'"/>
                          </xsl:attribute>
                          <xsl:call-template name="tmpSetRid">
                            <xsl:with-param name="ShapeType" select="$ShapeType"/>
                            <xsl:with-param name="PostionCount" select="$PostionCount"/>
                          </xsl:call-template>
                          </a:hlinkClick>
                      </xsl:when>
                      <!--Run program-->
                      <xsl:when test="@xlink:href and @presentation:action[ contains(.,'execute')]">
                        <a:hlinkClick>
                        <xsl:attribute name="action">
                          <xsl:value-of select="'ppaction://program'"/>
                        </xsl:attribute>
                          <xsl:call-template name="tmpSetRid">
                            <xsl:with-param name="ShapeType" select="$ShapeType"/>
                            <xsl:with-param name="PostionCount" select="$PostionCount"/>
                          </xsl:call-template>
                        </a:hlinkClick>
                      </xsl:when>
                      <!--Go to slide-->
                      <!--<xsl:when test="@xlink:href[ contains(.,'#page')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinksldjump'"/>
                      </xsl:attribute>
                    </xsl:when>-->
                      <xsl:when test="@xlink:href[ contains(.,'#')] and string-length(substring-before(@xlink:href,'#')) = 0 ">
                        <xsl:variable name="pageID">
                          <xsl:call-template name="getThePageId">
                            <xsl:with-param name="PageName" select="substring-after(@xlink:href,'#')"/>
                          </xsl:call-template>
                        </xsl:variable>
                        <xsl:if test="$pageID > 0">
                          <a:hlinkClick>
                            <xsl:attribute name="action">
                              <xsl:value-of select="'ppaction://hlinksldjump'"/>
                            </xsl:attribute>
                            <xsl:attribute name="r:id">
                              <xsl:value-of select="concat($ShapeType,$PostionCount)"/>
                            </xsl:attribute>
                            </a:hlinkClick>
                        </xsl:if>
                      </xsl:when>
                      <!--Go to Http-->
                      <xsl:when test="@xlink:href[ contains(.,'http')] and @presentation:action[ contains(.,'show')] ">
                        <a:hlinkClick>
                          <xsl:call-template name="tmpSetRid">
                            <xsl:with-param name="ShapeType" select="$ShapeType"/>
                            <xsl:with-param name="PostionCount" select="$PostionCount"/>
                          </xsl:call-template>
                        </a:hlinkClick>
                      </xsl:when>
                      <!--Go to document-->
                      <xsl:when test="@xlink:href and @presentation:action[ contains(.,'show')] ">
                        <a:hlinkClick>
                        <xsl:if test="not(@xlink:href[ contains(.,'#page')])">
                          <xsl:attribute name="action">
                            <xsl:value-of select="'ppaction://hlinkfile'"/>
                          </xsl:attribute>
                        </xsl:if>
                          <xsl:call-template name="tmpSetRid">
                            <xsl:with-param name="ShapeType" select="$ShapeType"/>
                            <xsl:with-param name="PostionCount" select="$PostionCount"/>
                          </xsl:call-template>
                          </a:hlinkClick>
                      </xsl:when>
                    
                    </xsl:choose>
                    <!--set value for attribute r:id-->
                  
                    <!-- Play Sound -->
                    <xsl:if test="@presentation:action[ contains(.,'sound')]">
                      <a:hlinkClick>
                        <a:snd>
                          <xsl:variable name="varMediaFilePath">
                            <xsl:if test="presentation:sound/@xlink:href [ contains(.,'../')]">
                              <xsl:value-of select="presentation:sound/@xlink:href" />
                            </xsl:if>
                            <xsl:if test="not(presentation:sound/@xlink:href [ contains(.,'../')])">
                              <xsl:value-of select="substring-after(presentation:sound/@xlink:href,'/')" />
                            </xsl:if>
                          </xsl:variable>
                          <xsl:variable name="varFileRelId">
                            <xsl:value-of select="translate(translate(translate(translate(translate($varMediaFilePath,'/','_'),'..','_'),'.','_'),':','_'),'%20D','_')"/>
                          </xsl:variable>
                          <xsl:attribute name="r:embed">
                            <xsl:value-of select="$varFileRelId"/>
                          </xsl:attribute>
                          <xsl:attribute name="name">
                            <xsl:value-of select="concat('SoundFileForShape',$PostionCount)"/>
                          </xsl:attribute>
                          <xsl:attribute name="builtIn">
                            <xsl:value-of select='1'/>
                          </xsl:attribute>
                          <pzip:import pzip:source="{$varMediaFilePath}" pzip:target="{concat('ppt/media/',$varFileRelId,'.wav')}" />
                        </a:snd>
                      </a:hlinkClick>
                    </xsl:if>

                  </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </xsl:if>
          </xsl:if>        
         </p:cNvPr >
        <p:cNvSpPr />
        <p:nvPr />
      </p:nvSpPr>
      <p:spPr>

        <xsl:choose>
          <xsl:when test="$grpFlag='true'">
            <a:xfrm>
            <xsl:call-template name ="tmpGroupdrawCordinates"/>
            </a:xfrm>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name ="tmpdrawCordinates"/>
          </xsl:otherwise>
        </xsl:choose>

        <xsl:call-template name ="getPresetGeom">
          <xsl:with-param name ="prstGeom" select ="$shapeName" />
        </xsl:call-template>

        <xsl:call-template name ="getGraphicProperties">
          <xsl:with-param name ="fileName" select ="$fileName" />
          <xsl:with-param name ="gr" >
            <xsl:choose>
              <xsl:when test="@draw:style-name">
              <xsl:value-of select ="@draw:style-name"/>
              </xsl:when>
              <xsl:when test="@presentation:style-name">
              <xsl:value-of select ="@presentation:style-name"/>
              </xsl:when>
            </xsl:choose>
          </xsl:with-param >
          <xsl:with-param name ="shapeCount" select ="$shapeCount" />
          <xsl:with-param name ="grpFlag" select ="$grpFlag" />
          <xsl:with-param name ="UniqueId" select ="$UniqueId" />
          
        </xsl:call-template>
      </p:spPr>
      <p:style>
        <a:lnRef idx="0">
          <a:schemeClr val="dk1" />
        </a:lnRef>
        <a:fillRef idx="0">
          <a:srgbClr val="99ccff">
            <xsl:call-template name ="getShade">
              <xsl:with-param name ="fileName" select ="$fileName" />
              <xsl:with-param name ="gr" >
                <xsl:if test ="@draw:style-name">
                  <xsl:value-of select ="@draw:style-name"/>
                </xsl:if>
                <xsl:if test ="@presentation:style-name">
                  <xsl:value-of select ="@presentation:style-name"/>
                </xsl:if >
              </xsl:with-param >
            </xsl:call-template>
          </a:srgbClr>
        </a:fillRef>
        <a:effectRef idx="0">
          <a:schemeClr val="accent1"/>
        </a:effectRef>
        <a:fontRef idx="minor">
          <a:schemeClr val="dk1"/>
        </a:fontRef>
      </p:style>
      <p:txBody>
        <xsl:call-template name ="getParagraphProperties">
          <xsl:with-param name ="fileName" select ="$fileName" />
          <xsl:with-param name ="gr" >
            <xsl:if test ="@draw:style-name">
              <xsl:value-of select ="@draw:style-name"/>
            </xsl:if>
            <xsl:if test ="@presentation:style-name">
              <xsl:value-of select ="@presentation:style-name"/>
            </xsl:if >
          </xsl:with-param >
          <xsl:with-param name ="customShape" select="$customShape" />
        </xsl:call-template>

        <xsl:choose>
          <xsl:when test ="draw:text-box/text:p or draw:text-box/text:list">
            <xsl:for-each select ="draw:text-box">
              <!-- Paremeter added by vijayeta,get master page name, dated:11-7-07-->
              <xsl:variable name ="masterPageName" select ="./parent::node()/parent::node()/@draw:master-page-name"/>
              <xsl:call-template name ="processShapeText">
                <xsl:with-param name ="fileName" select ="$fileName" />
                <xsl:with-param name="shapeType" select="name()" />
                <xsl:with-param name ="shapeCount" select="$shapeCount" />
                <xsl:with-param name ="grpFlag" select="$grpFlag" />
                <xsl:with-param name ="UniqueId" select="$UniqueId" />
                <xsl:with-param name ="gr" >
                  <xsl:if test ="parent::node()/@draw:style-name">
                    <xsl:value-of select ="parent::node()/@draw:style-name"/>
                  </xsl:if>
                  <xsl:if test ="parent::node()/@presentation:style-name">
                    <xsl:value-of select ="parent::node()/@presentation:style-name"/>
                  </xsl:if >
                </xsl:with-param >
                <xsl:with-param name ="masterPageName" select ="$masterPageName" />
              </xsl:call-template>
            </xsl:for-each>
          </xsl:when>
          <xsl:when test ="not(draw:text-box/text:p) and not(draw:text-box/text:list) and not(text:p or text:list)">
              <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
            </xsl:when>
          <xsl:otherwise>

            <xsl:call-template name ="processShapeText">
              <xsl:with-param name ="fileName" select ="$fileName" />
              <xsl:with-param name="shapeType" select="name()" />
              <xsl:with-param name ="shapeCount" select="$shapeCount" />
              <xsl:with-param name ="UniqueId" select="$UniqueId" />
              <xsl:with-param name ="grpFlag" select="$grpFlag" />
              <xsl:with-param name ="gr" >
                <xsl:if test ="@draw:style-name">
                  <xsl:value-of select ="@draw:style-name"/>
                </xsl:if>
                <xsl:if test ="@presentation:style-name">
                  <xsl:value-of select ="@presentation:style-name"/>
                </xsl:if >
              </xsl:with-param >
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </p:txBody>
    </p:sp>
  </xsl:template>
  <xsl:template name="tmpSetRid">
    <xsl:param name="ShapeType"/>
    <xsl:param name="PostionCount"/>
    <xsl:choose>
      <!--For jump to next,previous,first,last-->
      <xsl:when test="@presentation:action[ contains(.,'page') or contains(.,'stop')]">
        <xsl:attribute name="r:id">
          <xsl:value-of select="''"/>
        </xsl:attribute>
      </xsl:when>
      <!--For Run program & got to slide-->
      <!--
                    <xsl:when test="@xlink:href[contains(.,'#page')]">
                      <xsl:attribute name="r:id">
                        <xsl:value-of select="concat($ShapeType,$PostionCount)"/>
                      </xsl:attribute>-->
      <!--
                    </xsl:when>-->
      <!--For Go to document-->
      <xsl:when test="( @xlink:href[ contains(.,'http')] or @xlink:href[ contains(.,'mailto:')]) and @presentation:action[ contains(.,'show')] ">
        <xsl:attribute name="r:id">
          <xsl:value-of select="concat($ShapeType,$PostionCount)"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="@xlink:href and @presentation:action[ contains(.,'show')] ">
        <xsl:if test="not(@xlink:href[ contains(.,'#page')])">
          <xsl:attribute name="r:id">
            <xsl:value-of select="concat($ShapeType,$PostionCount)"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:when>
      <!--For Go to Run Programs-->
      <xsl:when test="@xlink:href and @presentation:action[ contains(.,'execute')]">
        <xsl:attribute name="r:id">
          <xsl:value-of select="concat($ShapeType,$PostionCount)"/>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpShapeType">
    <xsl:param name="nvprId"/>
    <xsl:choose>
      <xsl:when test="parent::node()/draw:custom-shape[@draw:id=$nvprId][position()=1] | parent::node()/draw:g/draw:custom-shape[@draw:id=$nvprId][position()=1]">
        <xsl:for-each select="parent::node()/draw:custom-shape[@draw:id=$nvprId][position()=1] | parent::node()/draw:g/draw:custom-shape[@draw:id=$nvprId][position()=1]">
        <xsl:choose>
          <xsl:when test="draw:enhanced-geometry[@draw:enhanced-path='M ?f0 0 L 21600 21600 0 21600 Z N' or @draw:type='isosceles-triangle']">
            <xsl:value-of select="'triangle'"/>
        </xsl:when>
          <xsl:when test="draw:enhanced-geometry[@draw:enhanced-path='M 0 0 L 21600 21600 0 21600 0 0 Z N' or @draw:type='rtTriangle']">
            <xsl:value-of select="'triangle'"/>
          </xsl:when>
          <xsl:when test="draw:enhanced-geometry[@draw:enhanced-path='M ?f0 0 L 21600 0 ?f1 21600 0 21600 Z N' or @draw:type='parallelogram']">
            <xsl:value-of select="'parallelogram'"/>
          </xsl:when>
          <xsl:when test="draw:enhanced-geometry[@draw:enhanced-path='M 0 1216152 L 228600 0 L 685800 0 L 914400 1216152 Z N' or @draw:type='mso-spt100']">
            <xsl:value-of select="'trapezoid'"/>
          </xsl:when>
          <xsl:when test="draw:enhanced-geometry[@draw:enhanced-path='M 10800 0 L 21600 10800 10800 21600 0 10800 10800 0 Z N' or @draw:type='diamond']">
            <xsl:value-of select="'diamond'"/>
          </xsl:when>
          <xsl:when test="draw:enhanced-geometry[@draw:enhanced-path='M 10800 0 L 0 8260 4230 21600 17370 21600 21600 8260 10800 0 Z N' or @draw:type='pentagon']">
            <xsl:value-of select="'pentagon'"/>
          </xsl:when>
          <xsl:when test="draw:enhanced-geometry[@draw:enhanced-path='M ?f0 0 L ?f1 0 21600 10800 ?f1 21600 ?f0 21600 0 10800 Z N' or @draw:type='hexagon']">
            <xsl:value-of select="'hexagon'"/>
          </xsl:when>
          <xsl:when test="draw:enhanced-geometry[@draw:type='ellipse']">
            <xsl:value-of select="'ellipse'"/>
          </xsl:when>
        </xsl:choose>
        </xsl:for-each>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="'rect'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
   <!-- Draw line-->
  <xsl:template name ="drawLine">
    <xsl:param name ="fileName" />
    <xsl:param name ="shapeCount" />
    <xsl:param name ="connectorType" />
    <xsl:param name ="NvPrId" />
    <xsl:param name ="grpFlag" />
    
    <xsl:variable name="x1" >
      <xsl:call-template name="convertUnitsToCm">
        <xsl:with-param name="length" select="@svg:x1"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="x2" >
      <xsl:call-template name="convertUnitsToCm">
        <xsl:with-param name="length" select="@svg:x2"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="y1" >
      <xsl:call-template name="convertUnitsToCm">
        <xsl:with-param name="length" select="@svg:y1"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="y2" >
      <xsl:call-template name="convertUnitsToCm">
        <xsl:with-param name="length" select="@svg:y2"/>
      </xsl:call-template>
    </xsl:variable>
    <p:cxnSp>
      <p:nvCxnSpPr>
        <p:cNvPr>
          <xsl:attribute name ="id">
				<xsl:value-of select ="$shapeCount+1"/>
          </xsl:attribute>
          <xsl:attribute name ="name">
            <xsl:value-of select ="concat($connectorType, position())" />
          </xsl:attribute>
          <!-- ADDED BY LOHITH - HYPER LINKS FOR SHAPES-->
          <xsl:variable name="PostionCount">
            <xsl:value-of select="position()"/>
          </xsl:variable>
          <xsl:if test="$grpFlag != 'true'">
          <xsl:if test="office:event-listeners">
            <xsl:for-each select ="office:event-listeners/presentation:event-listener">
              <xsl:if test="@script:event-name[contains(.,'dom:click')]">
                <a:hlinkClick>
                  <xsl:choose>
                    <!-- Go to previous slide-->
                    <xsl:when test="@presentation:action[ contains(.,'previous-page')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinkshowjump?jump=previousslide'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- Go to Next slide-->
                    <xsl:when test="@presentation:action[ contains(.,'next-page')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinkshowjump?jump=nextslide'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- Go to First slide-->
                    <xsl:when test="@presentation:action[ contains(.,'first-page')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinkshowjump?jump=firstslide'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- Go to Last slide-->
                    <xsl:when test="@presentation:action[ contains(.,'last-page')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinkshowjump?jump=lastslide'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- End Presentation -->
                    <xsl:when test="@presentation:action[ contains(.,'stop')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinkshowjump?jump=endshow'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- Run program-->
                    <xsl:when test="@xlink:href and @presentation:action[ contains(.,'execute')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://program'"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- Go to slide -->
                    <!--<xsl:when test="@xlink:href[ contains(.,'#page')]">
                      <xsl:attribute name="action">
                        <xsl:value-of select="'ppaction://hlinksldjump'"/>
                      </xsl:attribute>
                    </xsl:when>-->
                    <xsl:when test="@xlink:href[ contains(.,'#')] and string-length(substring-before(@xlink:href,'#')) = 0 ">
                      <xsl:variable name="pageID">
                        <xsl:call-template name="getThePageId">
                          <xsl:with-param name="PageName" select="substring-after(@xlink:href,'#')"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:if test="$pageID > 0">
                        <xsl:attribute name="action">
                          <xsl:value-of select="'ppaction://hlinksldjump'"/>
                        </xsl:attribute>
                        <xsl:attribute name="r:id">
                          <xsl:value-of select="concat('LineFileId',$PostionCount)"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                    <!-- Go to document -->
                    <xsl:when test="@xlink:href and @presentation:action[ contains(.,'show')] ">
                      <xsl:if test="not(@xlink:href[ contains(.,'#page')])">
                        <xsl:attribute name="action">
                          <xsl:value-of select="'ppaction://hlinkfile'"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                  </xsl:choose>
                  <!-- set value for attribute r:id-->
                  <xsl:choose>
                    <!-- For jump to next,previous,first,last-->
                    <xsl:when test="@presentation:action[ contains(.,'page') or contains(.,'stop')]">
                      <xsl:attribute name="r:id">
                        <xsl:value-of select="''"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!-- For Run program & got to slide-->
                    <!--<xsl:when test="@xlink:href[contains(.,'#page')]">
                      <xsl:attribute name="r:id">
                        <xsl:value-of select="concat('LineFileId',$PostionCount)"/>
                      </xsl:attribute>
                    </xsl:when>-->
                    <!-- For Go to document -->
                    <xsl:when test="@xlink:href and @presentation:action[ contains(.,'show')] ">
                      <xsl:if test="not(@xlink:href[ contains(.,'#page')])">
                        <xsl:attribute name="r:id">
                          <xsl:value-of select="concat('LineFileId',$PostionCount)"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:when>
                    <!-- For Go to Run Programs -->
                    <xsl:when test="@xlink:href and @presentation:action[ contains(.,'execute')]">
                      <xsl:attribute name="r:id">
                        <xsl:value-of select="concat('LineFileId',$PostionCount)"/>
                      </xsl:attribute>
                    </xsl:when>
                  </xsl:choose>
                  <!-- Play Sound -->
                  <xsl:if test="@presentation:action[ contains(.,'sound')]">
                    <a:snd>
                      <xsl:variable name="varMediaFilePath">
                        <xsl:if test="presentation:sound/@xlink:href [ contains(.,'../')]">
                          <xsl:value-of select="presentation:sound/@xlink:href" />
                        </xsl:if>
                        <xsl:if test="not(presentation:sound/@xlink:href [ contains(.,'../')])">
                          <xsl:value-of select="substring-after(presentation:sound/@xlink:href,'/')" />
                        </xsl:if>
                      </xsl:variable>
                      <xsl:variable name="varFileRelId">
                        <xsl:value-of select="translate(translate(translate(translate(translate($varMediaFilePath,'/','_'),'..','_'),'.','_'),':','_'),'%20D','_')"/>
                      </xsl:variable>
                      <xsl:attribute name="r:embed">
                        <xsl:value-of select="$varFileRelId"/>
                      </xsl:attribute>
                      <xsl:attribute name="name">
                        <xsl:value-of select="concat('SoundFileForShapeLine',$PostionCount)"/>
                      </xsl:attribute>
                      <xsl:attribute name="builtIn">
                        <xsl:value-of select='1'/>
                      </xsl:attribute>
                      <pzip:import pzip:source="{$varMediaFilePath}" pzip:target="{concat('ppt/media/',$varFileRelId,'.wav')}" />
                    </a:snd>
                  </xsl:if>
                </a:hlinkClick>
              </xsl:if>
            </xsl:for-each>
          </xsl:if>
          <!-- ADDED BY LOHITH - HYPER LINKS FOR SHAPES-->
          </xsl:if>
        </p:cNvPr>
		  <p:cNvCxnSpPr>

        <xsl:variable name="startShapepresetGm">
          <xsl:call-template name="tmpShapeType">
            <xsl:with-param name="nvprId" select="@draw:start-shape"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="endShapepresetGm">
          <xsl:call-template name="tmpShapeType">
            <xsl:with-param name="nvprId" select="@draw:end-shape"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="maxstartSpgluePointCount">
          <xsl:choose>
            <xsl:when test="$startShapepresetGm='ellipse' or $startShapepresetGm='octagon' or $startShapepresetGm='noSmoking'">
              <xsl:value-of select="'8'"/>
            </xsl:when>
            <xsl:when test="$startShapepresetGm='pentagon' or $startShapepresetGm='hexagon'
                         or $startShapepresetGm='cube' or $startShapepresetGm='parallelogram'">
              <xsl:value-of select="'6'"/>
            </xsl:when>
            <xsl:when test="$startShapepresetGm='triangle' or $startShapepresetGm='rtTriangle' or $startShapepresetGm='can'">
              <xsl:value-of select="'5'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'4'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="maxendSpgluePointCount">
          <xsl:choose>
            <xsl:when test="$endShapepresetGm='ellipse' or $endShapepresetGm='octagon' or $endShapepresetGm='noSmoking'">
              <xsl:value-of select="'8'"/>
            </xsl:when>
            <xsl:when test="$endShapepresetGm='pentagon' or $endShapepresetGm='hexagon'
                         or $endShapepresetGm='cube' or $endShapepresetGm='parallelogram'">
              <xsl:value-of select="'6'"/>
            </xsl:when>
            <xsl:when test="$endShapepresetGm='triangle' or $endShapepresetGm='rtTriangle' or $endShapepresetGm='can'">
              <xsl:value-of select="'5'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'4'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="stGluePoint">
          <xsl:value-of select="@draw:start-glue-point"/>
        </xsl:variable>
        <xsl:variable name="endGluePoint">
          <xsl:value-of select="@draw:end-glue-point"/>
        </xsl:variable>
			  <xsl:if test="@draw:start-glue-point">
				  <xsl:variable name="startShape">
					  <xsl:call-template name ="shapeGetnvPrId">
						  <xsl:with-param name ="spId">
							  <xsl:value-of select="@draw:start-shape"/>
						  </xsl:with-param>
					  </xsl:call-template>
				  </xsl:variable>
          <xsl:if test="$startShape !=''">
            <a:stCxn>
              <xsl:attribute name="id">
						  <xsl:value-of  select="$startShape" />
					  </xsl:attribute>
         
					  <xsl:attribute name="idx">
						  <xsl:choose>
                  <xsl:when test="$maxstartSpgluePointCount &gt;4 ">
                    <xsl:choose>
                      <xsl:when test="$stGluePoint = 0">
								  <xsl:value-of select="'0'"/>
							  </xsl:when>
                      <xsl:when test="number($maxstartSpgluePointCount) - number($stGluePoint) &gt;= 4">
                        <xsl:variable name="glueptval">
                          <xsl:value-of select="number($maxstartSpgluePointCount) + ( ( number($maxstartSpgluePointCount) - number($stGluePoint) + 4))"/>
                        </xsl:variable>
                        <xsl:choose>
                          <xsl:when test="$glueptval &lt; 0">
                            <xsl:value-of select="$glueptval * -1"/>
							  </xsl:when>
                          <xsl:when test="contains($glueptval,'NaN')">
                            <xsl:value-of select="'0'"/>
							  </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="$glueptval"/>
                        </xsl:otherwise>
                        </xsl:choose>
							  </xsl:when>
                      <xsl:when test="number($maxstartSpgluePointCount) - number($stGluePoint) &lt; 4">
                        <xsl:variable name="glueptval">
                          <xsl:value-of select="number($maxstartSpgluePointCount) - ( 4 +  ( number($maxstartSpgluePointCount) - number($stGluePoint)))"/>
                        </xsl:variable>
                        <xsl:choose>
                          <xsl:when test="$glueptval &lt; 0">
                            <xsl:value-of select="$glueptval * -1"/>
							  </xsl:when>
                          <xsl:when test="contains($glueptval,'NaN')">
                            <xsl:value-of select="'0'"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="$glueptval"/>
                          </xsl:otherwise>
                        </xsl:choose>
							  </xsl:when>
							  <xsl:otherwise>
								  <xsl:value-of select="'0'"/>
							  </xsl:otherwise>
                    </xsl:choose>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:choose>
                      <xsl:when test="$stGluePoint = 0">
                        <xsl:value-of select="'0'"/>
                      </xsl:when>
                      <xsl:when test="number($maxstartSpgluePointCount) - number($stGluePoint) &gt; 0">
                        <xsl:variable name="glueptval">
                          <xsl:value-of select="number($maxstartSpgluePointCount) - number($stGluePoint)"/>
                        </xsl:variable>
                        <xsl:choose>
                          <xsl:when test="$glueptval &lt; 0">
                            <xsl:value-of select="$glueptval * -1"/>
                          </xsl:when>
                          <xsl:when test="contains($glueptval,'NaN')">
                            <xsl:value-of select="'0'"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="$glueptval"/>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:when>
                      <xsl:when test="number($maxstartSpgluePointCount) - number($stGluePoint) &lt; 0">
                        <xsl:variable name="glueptval">
                          <xsl:value-of select="number($maxstartSpgluePointCount) - number($stGluePoint)"/>
                        </xsl:variable>
                        <xsl:choose>
                          <xsl:when test="$glueptval &lt; 0">
                            <xsl:value-of select="$glueptval * -1"/>
                          </xsl:when>
                          <xsl:when test="contains($glueptval,'NaN')">
                            <xsl:value-of select="'0'"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="$glueptval"/>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:when>
                      <xsl:when test="number($maxstartSpgluePointCount) - number($stGluePoint) = 0">
                        <xsl:value-of select="number($stGluePoint)"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="'0'"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:otherwise>
						  </xsl:choose>
					  </xsl:attribute>
				  </a:stCxn>
			  </xsl:if>
        </xsl:if>
       			  <xsl:if test="@draw:end-glue-point">
                      <xsl:variable name="endShape">
						  <xsl:call-template name ="shapeGetnvPrId">
							  <xsl:with-param name ="spId">
								  <xsl:value-of select="@draw:end-shape"/>
							  </xsl:with-param>
						  </xsl:call-template>
					  </xsl:variable>
          <xsl:if test="$endShape !=''">
				  <a:endCxn>
			  <xsl:attribute name="id">
						  <xsl:value-of  select="$endShape" />
					  </xsl:attribute>
					  <xsl:attribute name="idx">
						  <xsl:choose>
                <xsl:when test="$maxendSpgluePointCount &gt;4 ">
                  <xsl:choose>
                    <xsl:when test="$endGluePoint = 0">
								  <xsl:value-of select="'0'"/>
							  </xsl:when>
                    <xsl:when test="number($maxendSpgluePointCount) - number($endGluePoint) &gt;= 4">
                      <xsl:variable name="glueptval">
                        <xsl:value-of select="number($maxendSpgluePointCount) + ( ( number($maxendSpgluePointCount) - number($endGluePoint) + 4))"/>
                      </xsl:variable>
                      <xsl:choose>
                        <xsl:when test="$glueptval &lt; 0">
                          <xsl:value-of select="$glueptval * -1"/>
                        </xsl:when>
                        <xsl:when test="contains($glueptval,'NaN')">
                          <xsl:value-of select="'0'"/>
							  </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$glueptval"/>
                        </xsl:otherwise>
                      </xsl:choose>
							  </xsl:when>
                    <xsl:when test="number($maxendSpgluePointCount) - number($endGluePoint) &lt; 4">
                      <xsl:variable name="glueptval">
                        <xsl:value-of select="number($maxendSpgluePointCount) - ( 4 +  ( number($maxendSpgluePointCount) - number($endGluePoint)))"/>
                      </xsl:variable>
                      <xsl:choose>
                        <xsl:when test="$glueptval &lt; 0">
                          <xsl:value-of select="$glueptval * -1"/>
							  </xsl:when>
                        <xsl:when test="contains($glueptval,'NaN')">
                          <xsl:value-of select="'0'"/>
							  </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$glueptval"/>
                        </xsl:otherwise>
                      </xsl:choose>
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
                    <xsl:when test="number($maxendSpgluePointCount) - number($endGluePoint) &gt; 0">
                      <xsl:variable name="glueptval">
                        <xsl:value-of select="number($maxendSpgluePointCount) - number($endGluePoint)"/>
                      </xsl:variable>
                      <xsl:choose>
                        <xsl:when test="$glueptval &lt; 0">
                          <xsl:value-of select="$glueptval * -1"/>
                        </xsl:when>
                        <xsl:when test="contains($glueptval,'NaN')">
                          <xsl:value-of select="'0'"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$glueptval"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>
                    <xsl:when test="number($maxendSpgluePointCount) - number($stGluePoint) &lt; 0">
                      <xsl:variable name="glueptval">
                        <xsl:value-of select="number($maxendSpgluePointCount) - number($endGluePoint)"/>
                      </xsl:variable>
                      <xsl:choose>
                        <xsl:when test="$glueptval &lt; 0">
                          <xsl:value-of select="$glueptval * -1"/>
                        </xsl:when>
                        <xsl:when test="contains($glueptval,'NaN')">
                          <xsl:value-of select="'0'"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$glueptval"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>
                    <xsl:when test="number($maxendSpgluePointCount) - number($endGluePoint) = 0">
                      <xsl:value-of select="number($endGluePoint)"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="'0'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:otherwise>
						  </xsl:choose>
					  </xsl:attribute>
				  </a:endCxn>
        </xsl:if>
          </xsl:if>
		</p:cNvCxnSpPr>
        <p:nvPr/>
      </p:nvCxnSpPr>
      <p:spPr>
        <a:xfrm>
          <xsl:attribute name ="rot">
            <xsl:value-of select ="concat('cxnSp:rot:',$x1, ':',
																   $x2, ':', 
																   $y1, ':', 
																   $y2)"/>
          </xsl:attribute>
          <xsl:attribute name ="flipH">
            <xsl:value-of select ="concat('cxnSp:flipH:',$x1, ':',
																   $x2, ':', 
																   $y1, ':', 
																   $y2)"/>
          </xsl:attribute>
          <xsl:attribute name ="flipV">
            <xsl:value-of select ="concat('cxnSp:flipV:',$x1, ':',
																   $x2, ':', 
																   $y1, ':', 
																   $y2)"/>
          </xsl:attribute>
          <a:off >
            <xsl:attribute name ="x">
              <xsl:value-of select ="concat('cxnSp:x:',$x1, ':',
																   $x2, ':', 
																   $y1, ':', 
																   $y2,':',$grpFlag)"/>
            </xsl:attribute>
            <xsl:attribute name ="y">
              <xsl:value-of select ="concat('cxnSp:y:',$x1, ':',
																   $x2, ':', 
																   $y1, ':', 
																  $y2,':',$grpFlag)"/>
            </xsl:attribute>
          </a:off>
          <a:ext>
            <xsl:attribute name ="cx">
              <xsl:value-of select ="concat('cxnSp:cx:',$x1, ':',
																   $x2, ':', 
																   $y1, ':', 
																  $y2,':',$grpFlag)"/>
            </xsl:attribute>
            <xsl:attribute name ="cy">
              <xsl:value-of select ="concat('cxnSp:cy:',$x1, ':',
																   $x2, ':', 
																   $y1, ':', 
																  $y2,':',$grpFlag)"/>
            </xsl:attribute>

          </a:ext>
        </a:xfrm>

        <xsl:call-template name ="getPresetGeom">
          <xsl:with-param name ="prstGeom" select ="$connectorType" />
        </xsl:call-template>

        <xsl:call-template name ="getGraphicProperties">
          <xsl:with-param name ="fileName" select ="$fileName" />
          <xsl:with-param name ="gr" >
            <xsl:if test ="@draw:style-name">
              <xsl:value-of select ="@draw:style-name"/>
            </xsl:if>
            <xsl:if test ="@presentation:style-name">
              <xsl:value-of select ="@presentation:style-name"/>
            </xsl:if >
          </xsl:with-param >
          <xsl:with-param name ="grpFlag" select ="$grpFlag" />
        </xsl:call-template>
      </p:spPr>
      <p:style>
        <a:lnRef idx="2">
          <a:schemeClr val="dk1" />
        </a:lnRef>
        <a:fillRef idx="1">
          <a:srgbClr val="99ccff" />
        </a:fillRef>
        <a:effectRef idx="0">
          <a:schemeClr val="accent1"/>
        </a:effectRef>
        <a:fontRef idx="minor">
          <a:schemeClr val="dk1"/>
        </a:fontRef>
      </p:style>

    </p:cxnSp>
  </xsl:template>
  <!-- Get graphic properties for shape in context.xml-->
  <xsl:template name ="getGraphicProperties">
    <xsl:param name ="fileName" />
    <xsl:param name ="gr" />
    <xsl:param name ="shapeCount"  />
    <xsl:param name ="grpFlag"/>
    <xsl:param name ="UniqueId"  />
    <xsl:for-each select ="document($fileName)//office:automatic-styles/style:style[@style:name=$gr]/style:graphic-properties">
      <xsl:if test="position()=1">
      <!-- Parent style name-->
      <!--FILL-->
      <xsl:variable name="parentStyle" select="./parent::node()/@style:parent-style-name"/>
      <xsl:call-template name ="tmpshapefillColor">
        <xsl:with-param name ="parentStyle" select="$parentStyle" />
        <xsl:with-param name ="gr" select="$gr" />
        <xsl:with-param name ="shapeCount" select ="$shapeCount" />
        <xsl:with-param name ="grpFlag" select ="$grpFlag" />
        <xsl:with-param name ="UniqueId" select ="$UniqueId" />
      </xsl:call-template >

      <!--LINE COLOR AND STYLE-->
      <xsl:call-template name ="getLineStyle">
        <xsl:with-param name ="parentStyle" select="$parentStyle" />
      </xsl:call-template >

		<!-- SHADOW IMPLEMENTATION -->
      <xsl:choose>
        <xsl:when test="@draw:shadow='visible'">
          <xsl:call-template name ="tmpShapeShadow">
            <xsl:with-param name ="parentStyle" select="$parentStyle" />
          </xsl:call-template >
        </xsl:when>
        <xsl:when test="not(@draw:shadow) and $parentStyle !=''">
          <xsl:if test="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$parentStyle]/style:graphic-properties[@draw:shadow='visible']">
          <xsl:call-template name ="tmpShapeShadow">
            <xsl:with-param name ="parentStyle" select="$parentStyle" />
          </xsl:call-template >
            </xsl:if>
        </xsl:when>
      </xsl:choose>
      </xsl:if>
    </xsl:for-each>

  </xsl:template>
  <xsl:template name="tmpShapeShadow">
    <xsl:param name="parentStyle"/>
    <xsl:param name="filename"/>

			<xsl:variable name="shadowOffsetX">
      
    
				<xsl:if test="@draw:shadow-offset-x">
          <xsl:variable name="convertUnit">
            <xsl:call-template name="getConvertUnit">
              <xsl:with-param name="length" select="@draw:shadow-offset-x"/>
            </xsl:call-template>
          </xsl:variable>
					<xsl:value-of select="substring-before(@draw:shadow-offset-x,$convertUnit)"/>
				</xsl:if>
				<xsl:if test="not(@draw:shadow-offset-x)">
					<xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$parentStyle]/style:graphic-properties">
						<xsl:choose>
							<xsl:when test="(@draw:shadow-offset-x != 0)">
                <xsl:variable name="convertUnit">
                  <xsl:call-template name="getConvertUnit">
                    <xsl:with-param name="length" select="@draw:shadow-offset-x"/>
                  </xsl:call-template>
                </xsl:variable>
						<xsl:value-of select="substring-before(@draw:shadow-offset-x,$convertUnit)"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="'0'"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:for-each>
				</xsl:if>
			</xsl:variable>

			<xsl:variable name="shadowOffsetY">
				<xsl:if test="@draw:shadow-offset-y">
          <xsl:variable name="convertUnit">
            <xsl:call-template name="getConvertUnit">
              <xsl:with-param name="length" select="@draw:shadow-offset-y"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:value-of select="substring-before(@draw:shadow-offset-y,$convertUnit)"/>
				</xsl:if>
				<xsl:if test="not(@draw:shadow-offset-y)">
					<xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$parentStyle]/style:graphic-properties">
						<xsl:choose>
							<xsl:when test="(@draw:shadow-offset-y != 0)">
                <xsl:variable name="convertUnit">
                  <xsl:call-template name="getConvertUnit">
                    <xsl:with-param name="length" select="@draw:shadow-offset-y"/>
                  </xsl:call-template>
                </xsl:variable>
						<xsl:value-of select="substring-before(@draw:shadow-offset-y,$convertUnit)"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="'0'"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:for-each>
				</xsl:if>
			</xsl:variable>

			<xsl:variable name="shadowColor">
        <xsl:choose>
				<xsl:when test="@draw:shadow-color">
					<xsl:value-of select="substring-after(@draw:shadow-color,'#')"/>
				</xsl:when>
				<xsl:when test="not(@draw:shadow-color)">
					<xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$parentStyle]/style:graphic-properties">
						<xsl:value-of select="substring-after(@draw:shadow-color,'#')"/>
					</xsl:for-each>
				</xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'000000'"/>
        </xsl:otherwise>
        </xsl:choose>
			</xsl:variable>

			<!-- Transparency -->
			<xsl:variable name="shadowOpacity">
				<xsl:if test ="@draw:shadow-opacity != ''">
					<xsl:variable name ="alpha" select ="substring-before(@draw:shadow-opacity,'%')" />
					<xsl:if test ="$alpha = 0">
						<xsl:value-of select ="100000"/>
					</xsl:if>
					<xsl:if test ="$alpha != 0">
						<xsl:value-of select ="$alpha * 1000"/>
					</xsl:if>
				</xsl:if>
				<xsl:if test="not(@draw:shadow-opacity)">
					<xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name=$parentStyle]/style:graphic-properties">
						<xsl:variable name ="alpha" select ="substring-before(@draw:shadow-opacity,'%')" />
						<xsl:if test ="$alpha = ''">
							<xsl:value-of select ="100000"/>
						</xsl:if>
						<xsl:if test ="$alpha != ''">
							<xsl:value-of select ="$alpha * 1000"/>
						</xsl:if>
					</xsl:for-each>
				</xsl:if>
			</xsl:variable>

			<xsl:call-template name ="getShadowEffect" >
				<xsl:with-param name ="parentStyle" select="$parentStyle" />
				<xsl:with-param name ="shadowOffsetX" select="$shadowOffsetX"/>
				<xsl:with-param name ="shadowOffsetY" select="$shadowOffsetY"/>
				<xsl:with-param name ="shadowColor" select="$shadowColor"/>
				<xsl:with-param name ="shadowOpacity" select="$shadowOpacity"/>
			</xsl:call-template>
  </xsl:template>
  <!-- Get preset geometry-->
  <xsl:template name ="getPresetGeom">
    <xsl:param name ="prstGeom" />
    <xsl:choose>

      <!-- Oval -->
      <xsl:when test ="contains($prstGeom, 'Oval')">
        <a:prstGeom prst="ellipse">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Isosceles Triangle -->
      <xsl:when test ="contains($prstGeom, 'Isosceles Triangle')">
        <a:prstGeom prst="triangle">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Left-Right Arrow (Added by Mathi as on 19/7/2007)-->
      <xsl:when test ="contains($prstGeom, 'Left-Right Arrow')">
        <a:prstGeom prst="leftRightArrow">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Right Arrow (Added by Mathi as on 4/7/2007) -->
      <xsl:when test ="contains($prstGeom, 'Right Arrow')">
        <a:prstGeom prst="rightArrow">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Left Arrow (Added by Mathi as on 5/7/2007)-->
      <xsl:when test ="contains($prstGeom, 'Left Arrow')">
        <a:prstGeom prst="leftArrow">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!--Up Arrow (Added by Mathi as on 19/7/2007)-->
      <xsl:when test ="contains($prstGeom, 'Up Arrow')">
        <a:prstGeom prst="upArrow">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Up-Down Arrow (Added by Mathi as on 19/7/2007)-->
      <xsl:when test ="contains($prstGeom, 'Up-Down Arrow')">
        <a:prstGeom prst="upDownArrow">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Down Arrow (Added by Mathi as on 19/7/2007)-->
      <xsl:when test ="contains($prstGeom, 'Down Arrow')">
        <a:prstGeom prst="downArrow">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Right Triangle -->
      <xsl:when test ="contains($prstGeom, 'Right Triangle')">
        <a:prstGeom prst="rtTriangle">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Parallelogram -->
      <xsl:when test ="contains($prstGeom, 'Parallelogram')">
        <a:prstGeom prst="parallelogram">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Trapezoid (Added by A.Mathi as on 24/07/2007) -->
      <xsl:when test ="contains($prstGeom, 'Trapezoid')">
        <a:prstGeom prst="trapezoid">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Diamond -->
      <xsl:when test ="contains($prstGeom, 'Diamond')">
        <a:prstGeom prst="diamond">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Regular Pentagon -->
      <xsl:when test ="contains($prstGeom, 'Regular Pentagon')">
        <a:prstGeom prst="pentagon">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Hexagon -->
      <xsl:when test ="contains($prstGeom, 'Hexagon')">
        <a:prstGeom prst="hexagon">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Octagon -->
      <xsl:when test ="contains($prstGeom, 'Octagon')">
        <a:prstGeom prst="octagon">
          <a:avLst>
            <a:gd name="adj" fmla="val 32414"/>
          </a:avLst>
        </a:prstGeom>
      </xsl:when>
      <!-- Cube -->
      <xsl:when test ="contains($prstGeom, 'Cube')">
        <a:prstGeom prst="cube">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Can -->
      <xsl:when test ="contains($prstGeom, 'Can')">
        <a:prstGeom prst="can">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Chord (Added by Mathi on 17/10/2007)-->
      <xsl:when test ="contains($prstGeom, 'Chord')">
        <a:prstGeom prst="chord">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Cross (Added by Mathi on 19/7/2007)-->
      <xsl:when test ="contains($prstGeom, 'plus')">
        <a:prstGeom prst="plus">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- "No" Symbol (Added by A.Mathi as on 19/07/2007)-->
      <xsl:when test ="contains($prstGeom, 'noSmoking')">
        <a:prstGeom prst="noSmoking">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!--  Folded Corner (Added by A.Mathi as on 19/07/2007) -->
      <xsl:when test ="contains($prstGeom, 'foldedCorner')">
        <a:prstGeom prst="foldedCorner">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!--  Lightning Bolt (Added by A.Mathi as on 20/07/2007) -->
      <xsl:when test ="contains($prstGeom, 'lightningBolt')">
        <a:prstGeom prst="lightningBolt">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!--  Explosion 1 (Added by A.Mathi as on 20/07/2007) -->
      <xsl:when test ="contains($prstGeom, 'irregularSeal1')">
        <a:prstGeom prst="irregularSeal1">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Left Bracket (Added by A.Mathi as on 20/07/2007) -->
      <xsl:when test ="contains($prstGeom, 'Left Bracket')">
        <a:prstGeom prst="leftBracket">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Right Bracket (Added by A.Mathi as on 20/07/2007) -->
      <xsl:when test ="contains($prstGeom, 'Right Bracket')">
        <a:prstGeom prst="rightBracket">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Left Brace (Added by A.Mathi as on 20/07/2007) -->
      <xsl:when test ="contains($prstGeom, 'Left Brace')">
        <a:prstGeom prst="leftBrace">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Right Brace (Added by A.Mathi as on 23/07/2007) -->
      <xsl:when test ="contains($prstGeom, 'Right Brace')">
        <a:prstGeom prst="rightBrace">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
		<!-- Rectangular Callout (modified by A.Mathi) -->
		<xsl:when test ="contains($prstGeom, 'Rectangular Callout')">
			<xsl:choose>
				<xsl:when test="not(draw:enhanced-geometry/@draw:modifiers='6300 24300')">
			                <xsl:call-template name="tmpCalloutAdjustment">
						<xsl:with-param name="prst" select="'wedgeRectCallout'"/>
					</xsl:call-template>
				</xsl:when>
				<xsl:otherwise>
					<a:prstGeom prst="wedgeRectCallout">
						<a:avLst/>
					</a:prstGeom>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:when>
		<!-- Rounded Rectangular Callout (modified by A.Mathi) -->
		<xsl:when test ="contains($prstGeom, 'wedgeRoundRectCallout')">
			<xsl:choose>
				<xsl:when test="not(draw:enhanced-geometry/@draw:modifiers='6300 24300')">
			                <xsl:call-template name="tmpCalloutAdjustment">
						<xsl:with-param name="prst" select="'wedgeRoundRectCallout'"/>
					</xsl:call-template>
		                </xsl:when>
				<xsl:otherwise>
					<a:prstGeom prst="wedgeRoundRectCallout">
						<a:avLst/>
					</a:prstGeom>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:when>
		<!-- Oval Callout (modified by A.Mathi) -->
		<xsl:when test ="contains($prstGeom, 'wedgeEllipseCallout')">
			<xsl:choose>
				<xsl:when test="not(draw:enhanced-geometry/@draw:modifiers='6300 24300')">
			                <xsl:call-template name="tmpCalloutAdjustment">
						<xsl:with-param name="prst" select="'wedgeEllipseCallout'"/>
					</xsl:call-template>
				</xsl:when>
				<xsl:otherwise>
					<a:prstGeom prst="wedgeEllipseCallout">
						<a:avLst/>
					</a:prstGeom>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:when>
		<!-- Cloud Callout (modified by A.Mathi) -->
		<xsl:when test ="contains($prstGeom, 'Cloud Callout')">
			<xsl:choose>
				<xsl:when test="not(draw:enhanced-geometry/@draw:modifiers='6300 24300')">
			                <xsl:call-template name="tmpCalloutAdjustment">
						<xsl:with-param name="prst" select="'cloudCallout'"/>
					</xsl:call-template>
				</xsl:when>
				<xsl:otherwise>
					<a:prstGeom prst="cloudCallout">
						<a:avLst/>
					</a:prstGeom>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:when>
		<!-- Line Callout 1 -->
		<xsl:when test ="contains($prstGeom, 'borderCallout1')">
			<xsl:choose>
				<xsl:when test="not(draw:enhanced-geometry/@draw:modifiers='-8300 24500 -1800 4000')">
					<xsl:call-template name="tmpCalloutAdjustment" >
						<xsl:with-param name="prst" select="'borderCallout1'"/>
					</xsl:call-template>
				</xsl:when>
				<xsl:otherwise>
					<a:prstGeom prst="borderCallout1">
						<a:avLst/>
					</a:prstGeom>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:when>
		<!-- Line Callout 2 -->
		<xsl:when test ="contains($prstGeom, 'borderCallout2')">
			<xsl:choose>
				<xsl:when test="not(draw:enhanced-geometry/@draw:modifiers='-10000 24500 -3600 4000 -1800 4000')">
					<xsl:call-template name="tmpCalloutAdjustment" >
						<xsl:with-param name="prst" select="'borderCallout2'"/>
					</xsl:call-template>
				</xsl:when>
				<xsl:otherwise>
					<a:prstGeom prst="borderCallout2">
						<a:avLst/>
					</a:prstGeom>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:when>
		<!-- Line Callout 3 -->
		<xsl:when test ="contains($prstGeom, 'borderCallout3')">
			<xsl:choose>
				<xsl:when test="not(draw:enhanced-geometry/@draw:modifiers='-1800 0 -3600 0 -3600 0 -1800 4000')">
					<xsl:call-template name="tmpCalloutAdjustment" >
						<xsl:with-param name="prst" select="'borderCallout3'"/>
					</xsl:call-template>
				</xsl:when>
				<xsl:otherwise>
					<a:prstGeom prst="borderCallout3">
						<a:avLst/>
					</a:prstGeom>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:when>
		
      <!-- Bent Arrow (Added by A.Mathi as on 23/07/2007) -->
      <xsl:when test ="contains($prstGeom, 'bentArrow')">
        <a:prstGeom prst="bentArrow">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- U-Turn Arrow (Added by A.Mathi as on 23/07/2007) -->
      <xsl:when test ="contains($prstGeom, 'U-Turn Arrow')">
        <a:prstGeom prst="uturnArrow">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
		<!-- Quad Arrow -->
		<xsl:when test ="contains($prstGeom, 'Quad Arrow')">
			<a:prstGeom prst="quadArrow">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Block Arc -->
		<xsl:when test ="contains($prstGeom, 'Block Arc')">
			<a:prstGeom prst="blockArc">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Notched Right Arrow -->
		<xsl:when test ="contains($prstGeom, 'notchedRightArrow')">
			<a:prstGeom prst="notchedRightArrow">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Pentagon -->
		<xsl:when test ="contains($prstGeom, 'Pentagon')">
			<a:prstGeom prst="homePlate">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Chevron -->
		<xsl:when test ="contains($prstGeom, 'Chevron')">
			<a:prstGeom prst="chevron">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>

		<!--Equation Shapes-->
		<!--Not Equal-->
		<xsl:when test ="contains($prstGeom, 'Not Equal ')">
			<a:prstGeom prst="mathNotEqual">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!--Equal-->
		<xsl:when test ="contains($prstGeom, 'Equal ')">
			<a:prstGeom prst="mathEqual">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!--Plus-->
		<xsl:when test ="contains($prstGeom, 'Plus ')">
			<a:prstGeom prst="mathPlus">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!--Minus-->
		<xsl:when test ="contains($prstGeom, 'Minus ')">
			<a:prstGeom prst="mathMinus">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!--Multiply-->
		<xsl:when test ="contains($prstGeom, 'Multiply ')">
			<a:prstGeom prst="mathMultiply">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!--Division-->
		<xsl:when test ="contains($prstGeom, 'Division ')">
			<a:prstGeom prst="mathDivide">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		
      <!-- Rounded Rectangle -->
      <xsl:when test ="contains($prstGeom, 'Rounded Rectangle')">
        <a:prstGeom prst="roundRect">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Snip Single Corner Rectangle -->
      <xsl:when test ="contains($prstGeom, 'Snip Single Corner Rectangle')">
        <a:prstGeom prst="snip1Rect">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
		<!-- Snip Same Side Corner Rectangle -->
		<xsl:when test ="contains($prstGeom, 'Snip Same Side Corner Rectangle')">
			<a:prstGeom prst="snip2SameRect">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Snip Diagonal Corner Rectangle -->
		<xsl:when test ="contains($prstGeom, 'Snip Diagonal Corner Rectangle')">
			<a:prstGeom prst="snip2DiagRect">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Snip and Round Single Corner Rectangle -->
		<xsl:when test ="contains($prstGeom, 'Snip and Round Single Corner Rectangle')">
			<a:prstGeom prst="snipRoundRect">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
      
      <!-- Circular Arrow -->
      <xsl:when test ="contains($prstGeom, 'circular-arrow')">
        <a:prstGeom prst="circularArrow">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <xsl:when test ="contains($prstGeom, 'curvedUpArrow')">
        <a:prstGeom prst="curvedUpArrow">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <xsl:when test ="contains($prstGeom, 'curvedDownArrow')">
        <a:prstGeom prst="curvedDownArrow">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <xsl:when test ="contains($prstGeom, 'curvedLeftArrow')">
        <a:prstGeom prst="curvedLeftArrow">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <xsl:when test ="contains($prstGeom, 'curvedRightArrow')">
        <a:prstGeom prst="curvedRightArrow">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- End of Circular Arrow Code-->
      
      <!-- leftUp Arrow -->
      <xsl:when test ="contains($prstGeom, 'leftUpArrow')">
        <a:prstGeom prst="leftUpArrow">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- bentUp Arrow -->
      <xsl:when test ="contains($prstGeom, 'bentUpArrow')">
        <a:prstGeom prst="bentUpArrow">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      
		<!-- Round Single Corner Rectangle -->
		<xsl:when test ="contains($prstGeom, 'Round Single Corner Rectangle')">
			<a:prstGeom prst="round1Rect">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Round Same Side Corner Rectangle -->
		<xsl:when test ="contains($prstGeom, 'Round Same Side Corner Rectangle')">
			<a:prstGeom prst="round2SameRect">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Round Diagonal Corner Rectangle -->
		<xsl:when test ="contains($prstGeom, 'Round Diagonal Corner Rectangle')">
			<a:prstGeom prst="round2DiagRect">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Explosion 2 -->
		<xsl:when test ="contains($prstGeom, 'Explosion 2')">
			<a:prstGeom prst="irregularSeal2">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Heptagon -->
		<xsl:when test ="contains($prstGeom, 'Heptagon')">
			<a:prstGeom prst="heptagon">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Decagon -->
		<xsl:when test ="contains($prstGeom, 'Decagon')">
			<a:prstGeom prst="decagon">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Dodecagon -->
		<xsl:when test ="contains($prstGeom, 'Dodecagon')">
			<a:prstGeom prst="dodecagon">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Half Frame -->
		<xsl:when test="contains($prstGeom, 'Half Frame')">
			<a:prstGeom prst="halfFrame">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Pie -->
		<xsl:when test ="contains($prstGeom, 'Pie')">
			<a:prstGeom prst="pie">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Frame -->
		<xsl:when test="contains($prstGeom, 'Frame')">
			<a:prstGeom prst="frame">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- L-Shape -->
		<xsl:when test="contains($prstGeom, 'L-Shape')">
			<a:prstGeom prst="corner">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Diagonal Stripe -->
		<xsl:when test="contains($prstGeom, 'Diagonal Stripe')">
			<a:prstGeom prst="diagStripe">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Plaque -->
		<xsl:when test="contains($prstGeom, 'Plaque')">
			<a:prstGeom prst="plaque">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Bevel -->
		<xsl:when test="contains($prstGeom, 'Bevel')">
			<a:prstGeom prst="bevel">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Donut -->
		<xsl:when test="contains($prstGeom, 'Donut')">
			<a:prstGeom prst="donut">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- TearDrop -->
		<xsl:when test="contains($prstGeom, 'Teardrop')">
			<a:prstGeom prst="teardrop">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		
      <!-- Rectangles -->
      <xsl:when test ="(contains($prstGeom, 'Rectangle')) or (contains($prstGeom, 'TextBox'))">
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>

      <!-- Connectors-->
      <!-- Straight line-->
      <xsl:when test ="contains($prstGeom, 'Straight Connector')">
        <a:prstGeom prst="line">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Straight Arrow Connector -->
      <xsl:when test ="contains($prstGeom, 'Straight Arrow Connector')">
        <a:prstGeom prst="straightConnector1">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Elbow Connector -->
      <xsl:when test ="contains($prstGeom, 'Elbow Connector')">
        <a:prstGeom prst="bentConnector3">
          <a:avLst>
            <a:gd name="adj1" fmla="val 50000"/>
          </a:avLst>
        </a:prstGeom>

      </xsl:when>
      <!-- Curved Connector -->
      <xsl:when test ="contains($prstGeom, 'Curved Connector')">
        <a:prstGeom prst="curvedConnector3">
          <a:avLst>
            <a:gd name="adj1" fmla="val 60417"/>
          </a:avLst>
        </a:prstGeom>
      </xsl:when>

      <!-- Flow Chart: Process -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Process')">
        <a:prstGeom prst="flowChartProcess">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Alternate Process -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Alternate Process')">
        <a:prstGeom prst="flowChartAlternateProcess">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Decision -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Decision')">
        <a:prstGeom prst="flowChartDecision">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Data -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Data')">
        <a:prstGeom prst="flowChartInputOutput">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Predefined-process -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Predefined Process')">
        <a:prstGeom prst="flowChartPredefinedProcess">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Internal-storage -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Internal Storage')">
        <a:prstGeom prst="flowChartInternalStorage">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Document -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Document')">
        <a:prstGeom prst="flowChartDocument">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Multidocument -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Multi document')">
        <a:prstGeom prst="flowChartMultidocument">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Terminator -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Terminator')">
        <a:prstGeom prst="flowChartTerminator">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Preparation -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Preparation')">
        <a:prstGeom prst="flowChartPreparation">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Manual-input -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Manual Input')">
        <a:prstGeom prst="flowChartManualInput">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Manual-operation -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Manual Operation')">
        <a:prstGeom prst="flowChartManualOperation">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
		<xsl:when test ="contains($prstGeom, 'flowchart-manual-operation')">
			<a:prstGeom prst="flowChartManualOperation">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
      <!-- Flow Chart: Connector -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Connector')">
        <a:prstGeom prst="flowChartConnector">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Off-page-connector -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Off-page Connector')">
        <a:prstGeom prst="flowChartOffpageConnector">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Card -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Card')">
        <a:prstGeom prst="flowChartPunchedCard">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Punched-tape -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Punched Tape')">
        <a:prstGeom prst="flowChartPunchedTape">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Summing-junction -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Summing Junction')">
        <a:prstGeom prst="flowChartSummingJunction">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Or -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Or')">
        <a:prstGeom prst="flowChartOr">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Collate -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Collate')">
        <a:prstGeom prst="flowChartCollate">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Sort -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Sort')">
        <a:prstGeom prst="flowChartSort">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Extract -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Extract')">
        <a:prstGeom prst="flowChartExtract">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Merge -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Merge')">
        <a:prstGeom prst="flowChartMerge">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Stored-data -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Stored Data')">
        <a:prstGeom prst="flowChartOnlineStorage">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Delay-->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Delay')">
        <a:prstGeom prst="flowChartDelay">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Sequential-access -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Sequential Access Storage')">
        <a:prstGeom prst="flowChartMagneticTape">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Direct-access-storage -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Direct Access Storage')">
        <a:prstGeom prst="flowChartMagneticDrum">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Magnetic-disk -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Magnetic Disk')">
        <a:prstGeom prst="flowChartMagneticDisk">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
      <!-- Flow Chart: Display -->
      <xsl:when test ="contains($prstGeom, 'Flowchart: Display')">
        <a:prstGeom prst="flowChartDisplay">
          <a:avLst/>
        </a:prstGeom>
      </xsl:when>
		<!-- Action Buttons Back or Previous -->
		<xsl:when test ="contains($prstGeom, 'actionButtonBackPrevious')">
			<a:prstGeom prst="actionButtonBackPrevious">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Action Buttons Forward or Next -->
		<xsl:when test ="contains($prstGeom, 'actionButtonForwardNext')">
			<a:prstGeom prst="actionButtonForwardNext">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Action Buttons Beginning -->
		<xsl:when test ="contains($prstGeom, 'actionButtonBeginning')">
			<a:prstGeom prst="actionButtonBeginning">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Action Buttons end -->
		<xsl:when test ="contains($prstGeom, 'actionButtonEnd')">
			<a:prstGeom prst="actionButtonEnd">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Action Buttons Home -->
		<xsl:when test ="contains($prstGeom, 'actionButtonHome')">
			<a:prstGeom prst="actionButtonHome">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Action Buttons Information -->
		<xsl:when test ="contains($prstGeom, 'actionButtonInformation')">
			<a:prstGeom prst="actionButtonInformation">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		<!-- Action Buttons Return -->
		<xsl:when test ="contains($prstGeom, 'actionButtonReturn')">
			<a:prstGeom prst="actionButtonReturn">
				<a:avLst/>
			</a:prstGeom>
		</xsl:when>
		
    </xsl:choose>
  </xsl:template>
  <!-- Get fill details-->
  <xsl:template name ="tmpshapefillColor">
    <xsl:param name ="parentStyle" />
    <xsl:param name ="gr" />
    <xsl:param name ="shapeCount" />
    <xsl:param name ="grpFlag" />
    <xsl:param name ="UniqueId"  />
    <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
    <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
    <xsl:variable name="tranparencyinContent" select="substring-before(@draw:opacity,'%')"/>
    <xsl:choose>
      <xsl:when test="@draw:fill or @draw:fill-color ">
        <xsl:call-template name ="tmpFill">
          <xsl:with-param name="tranparencyinContent" select="$tranparencyinContent"/>
          <xsl:with-param name="parentStyle" select="$parentStyle"/>
          <xsl:with-param name ="shapeCount" select ="$shapeCount" />
          <xsl:with-param name ="grpFlag" select ="$grpFlag" />
          <xsl:with-param name ="UniqueId" select ="$UniqueId" />
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$parentStyle !='' and not(@draw:fill or @draw:fill-color)">
        <xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $parentStyle]/style:graphic-properties">
          <xsl:call-template name ="tmpFill">
            <xsl:with-param name="tranparencyinStyle" select="substring-before(@draw:opacity,'%')"/>
            <xsl:with-param name="tranparencyinContent" select="$tranparencyinContent"/>
            <xsl:with-param name ="shapeCount" select ="$shapeCount" />
            <xsl:with-param name ="grpFlag" select ="$grpFlag" />
            <xsl:with-param name ="UniqueId" select ="$UniqueId" />
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
    </xsl:choose>
      
  </xsl:template>
  <xsl:template name="tmpFill">
    <xsl:param name="tranparencyinContent"/>
    <xsl:param name="tranparencyinStyle"/>
    <xsl:param name="parentStyle"/>
    <xsl:param name ="shapeCount" />
    <xsl:param name ="grpFlag"/>
    <xsl:param name ="UniqueId" />
    <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
    <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
    <xsl:choose>
      <xsl:when test ="@draw:fill='solid'">
        <a:solidFill>
          <a:srgbClr  >
            <xsl:attribute name ="val">
              <xsl:value-of select ="translate(substring-after(@draw:fill-color,'#'),$lcletters,$ucletters)"/>
            </xsl:attribute>
            <xsl:call-template name="tmpFillTransperancy">
              <xsl:with-param name="tranparencyinStyle" select="$tranparencyinStyle"/>
              <xsl:with-param name="tranparencyinContent" select="$tranparencyinContent"/>
              <xsl:with-param name="parentStyle" select="$parentStyle"/>
            </xsl:call-template>
          </a:srgbClr >
        </a:solidFill>
      </xsl:when>
      <xsl:when test ="@draw:fill='none'">
        <a:noFill/>
      </xsl:when>
      <xsl:when test ="@draw:fill='gradient'">
        <xsl:call-template name="tmpGradientFill">
          <xsl:with-param name="gradStyleName" select="@draw:fill-gradient-name"/>
          <xsl:with-param  name="opacity" select="substring-before(@draw:opacity,'%')"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test ="@draw:fill='bitmap' and $grpFlag!='true'">
        <xsl:call-template name="tmpBitmapFill">
          <xsl:with-param name="FileName" select="concat('bitmap',$shapeCount)" />
          <xsl:with-param name="var_imageName" select="@draw:fill-image-name" />
        </xsl:call-template>
      </xsl:when>

      <xsl:when test ="@draw:fill='bitmap' and $grpFlag='true'">
        <xsl:call-template name="tmpBitmapFill">
          <xsl:with-param name="FileName" select="concat('grpbitmap',$UniqueId)" />
          <xsl:with-param name="var_imageName" select="@draw:fill-image-name" />
          <xsl:with-param name ="UniqueId" select ="$UniqueId" />
        </xsl:call-template>
      </xsl:when>
      
      <xsl:when test ="@draw:fill-color">
        <a:solidFill>
          <a:srgbClr  >
            <xsl:attribute name ="val">
              <xsl:value-of select ="translate(substring-after(@draw:fill-color,'#'),$lcletters,$ucletters)"/>
            </xsl:attribute>
            <xsl:call-template name="tmpFillTransperancy">
              <xsl:with-param name="tranparencyinStyle" select="$tranparencyinStyle"/>
              <xsl:with-param name="tranparencyinContent" select="$tranparencyinContent"/>
              <xsl:with-param name="parentStyle" select="$parentStyle"/>
            </xsl:call-template>
          </a:srgbClr >
        </a:solidFill>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpFillTransperancy">
    <xsl:param name="tranparencyinContent"/>
    <xsl:param name="tranparencyinStyle"/>
    <xsl:param name="parentStyle"/>
              <xsl:choose>
      <xsl:when test ="$tranparencyinContent ='0' and $tranparencyinStyle =''">
        <a:alpha val="0"/>
      </xsl:when>
      <xsl:when test ="$tranparencyinContent !='' and $tranparencyinStyle =''">
                  <a:alpha>
                    <xsl:attribute name="val">
            <xsl:value-of select="format-number($tranparencyinContent * 1000,'#')" />
                    </xsl:attribute>
                  </a:alpha>
                </xsl:when>
      <xsl:when test ="$tranparencyinContent !='' and $tranparencyinStyle !=''">
        <a:alpha>
          <xsl:attribute name="val">
            <xsl:value-of select="format-number($tranparencyinContent * 1000,'#')" />
          </xsl:attribute>
        </a:alpha>
      </xsl:when>
      <xsl:when test ="$tranparencyinContent ='' and $tranparencyinStyle ='' ">
        <xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $parentStyle]/style:graphic-properties">
          <xsl:if test="@draw:opacity">
            <xsl:choose>
              <xsl:when test="substring-before(@draw:opacity,'%') =0">
                <a:alpha val="0"/>
              </xsl:when>
              <xsl:otherwise>
          <a:alpha>
            <xsl:attribute name="val">
              <xsl:value-of select="format-number( substring-before(@draw:opacity,'%')* 1000,'#')" />
            </xsl:attribute>
          </a:alpha>
              </xsl:otherwise>
            </xsl:choose>
            </xsl:if>
        </xsl:for-each>
              </xsl:when>
      <xsl:when test ="$tranparencyinContent ='' and $tranparencyinStyle ='0'">
        <a:alpha val="0"/>
      </xsl:when>
      <xsl:when test ="$tranparencyinContent ='' and $tranparencyinStyle !=''">
        <a:alpha>
          <xsl:attribute name="val">
            <xsl:value-of select="format-number($tranparencyinStyle * 1000,'#')" />
          </xsl:attribute>
        </a:alpha>
      </xsl:when>

    </xsl:choose>
  </xsl:template>
  <xsl:template name="tmpLineTransperancy">
    <xsl:param name="tranparencyinContent"/>
    <xsl:param name="tranparencyinStyle"/>
    <xsl:param name="parentStyle"/>
    <xsl:choose>
      <xsl:when test ="$tranparencyinContent ='0' and $tranparencyinStyle =''">
        <a:alpha val="0"/>
      </xsl:when>
      <xsl:when test ="$tranparencyinContent !='' and $tranparencyinStyle =''">
        <a:alpha>
          <xsl:attribute name="val">
            <xsl:value-of select="format-number($tranparencyinContent * 1000,'#')" />
          </xsl:attribute>
        </a:alpha>
      </xsl:when>
      <xsl:when test ="$tranparencyinContent !='' and $tranparencyinStyle !=''">
        <a:alpha>
          <xsl:attribute name="val">
            <xsl:value-of select="format-number($tranparencyinContent * 1000,'#')" />
          </xsl:attribute>
        </a:alpha>
      </xsl:when>
      <xsl:when test ="$tranparencyinContent ='' and $tranparencyinStyle ='' ">
        <xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $parentStyle]/style:graphic-properties">
          <xsl:if test="@svg:stroke-opacity">
            <xsl:choose>
              <xsl:when test="substring-before(@svg:stroke-opacity,'%') =0">
                <a:alpha val="0"/>
              </xsl:when>
              <xsl:otherwise>
            <a:alpha>
              <xsl:attribute name="val">
                <xsl:value-of select="format-number( substring-before(@svg:stroke-opacity,'%')* 1000,'#')" />
              </xsl:attribute>
            </a:alpha>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="$tranparencyinContent ='' and $tranparencyinStyle ='0'">
        <a:alpha val="0"/>
      </xsl:when>
      <xsl:when test ="$tranparencyinContent ='' and $tranparencyinStyle !=''">
        <a:alpha>
          <xsl:attribute name="val">
            <xsl:value-of select="format-number($tranparencyinStyle * 1000,'#')" />
          </xsl:attribute>
        </a:alpha>
      </xsl:when>

    </xsl:choose>
  </xsl:template>
  <xsl:template name ="tmpLinefillColor">
    <xsl:param name ="parentStyle" />
    <xsl:param name ="gr" />
    <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
    <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
    <xsl:variable name="tranparencyinContent" select="substring-before(@svg:stroke-opacity,'%')"/>
    <xsl:choose>
      <xsl:when test="@draw:stroke or @svg:stroke-color ">
        <xsl:call-template name ="tmpLineFill">
          <xsl:with-param name="tranparencyinContent" select="$tranparencyinContent"/>
          <xsl:with-param name="parentStyle" select="$parentStyle"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$parentStyle !='' and not(@draw:stroke or @svg:stroke-color)">
        <xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $parentStyle]/style:graphic-properties">
          <xsl:call-template name ="tmpLineFill">
            <xsl:with-param name="tranparencyinStyle" select="substring-before(@svg:stroke-opacity,'%')"/>
            <xsl:with-param name="tranparencyinContent" select="$tranparencyinContent"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:when>
    </xsl:choose>

  </xsl:template>
  <xsl:template name="tmpLineFill">
    <xsl:param name="tranparencyinContent"/>
    <xsl:param name="tranparencyinStyle"/>
    <xsl:param name="parentStyle"/>
    <xsl:variable name="lcletters">abcdefghijklmnopqrstuvwxyz</xsl:variable>
    <xsl:variable name="ucletters">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
    <xsl:choose>
      <xsl:when test ="@draw:stroke='solid' or @draw:stroke='dash'">
        <a:solidFill>
          <a:srgbClr  >
            <xsl:attribute name ="val">
              <xsl:if test="@svg:stroke-color">
              <xsl:value-of select ="translate(substring-after(@svg:stroke-color,'#'),$lcletters,$ucletters)"/>
              </xsl:if>
              <xsl:if test="not(@svg:stroke-color)">
                <xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $parentStyle]/style:graphic-properties">
                  <xsl:value-of select ="translate(substring-after(@svg:stroke-color,'#'),$lcletters,$ucletters)"/>
                </xsl:for-each>
              </xsl:if>
            </xsl:attribute>
            <xsl:call-template name="tmpLineTransperancy">
              <xsl:with-param name="tranparencyinStyle" select="$tranparencyinStyle"/>
              <xsl:with-param name="tranparencyinContent" select="$tranparencyinContent"/>
              <xsl:with-param name="parentStyle" select="$parentStyle"/>
            </xsl:call-template>
          </a:srgbClr >
        </a:solidFill>
      </xsl:when>
      <xsl:when test ="@draw:stroke='none'">
        <a:noFill/>
      </xsl:when>
      <xsl:when test ="@draw:stroke='gradient'">
        <xsl:call-template name="tmpGradientFill">
          <xsl:with-param name="gradStyleName" select="@draw:fill-gradient-name"/>
          <xsl:with-param  name="opacity" select="substring-before(@svg:stroke-opacity,'%')"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test ="@svg:stroke-color">
        <a:solidFill>
          <a:srgbClr  >
            <xsl:attribute name ="val">
              <xsl:value-of select ="translate(substring-after(@svg:stroke-color,'#'),$lcletters,$ucletters)"/>
            </xsl:attribute>
            <xsl:call-template name="tmpLineTransperancy">
              <xsl:with-param name="tranparencyinStyle" select="$tranparencyinStyle"/>
              <xsl:with-param name="tranparencyinContent" select="$tranparencyinContent"/>
              <xsl:with-param name="parentStyle" select="$parentStyle"/>
            </xsl:call-template>
          </a:srgbClr >
        </a:solidFill>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
   <xsl:template name ="getFillDetails">
    <xsl:param name ="parentStyle" />
    <xsl:variable name ="opacity" >
      <xsl:value-of select ="@draw:opacity"/>
    </xsl:variable>
    <xsl:choose>
      <!-- No fill-->
      <xsl:when test ="@draw:fill='none'">
        <a:noFill/>
      </xsl:when>
      <!-- Solid fill-->
      <xsl:when test="@draw:fill-color">
        <xsl:call-template name ="getFillColor">
          <xsl:with-param name ="fill-color" select ="@draw:fill-color" />
          <xsl:with-param name ="opacity" select ="@draw:opacity" />
        </xsl:call-template>
      </xsl:when>
      <!-- Default fill color from styles.xml-->
      <xsl:when test ="($parentStyle != '')">
        <xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $parentStyle]/style:graphic-properties">
          <xsl:choose>
            <xsl:when test ="@draw:fill='none'">
              <a:noFill/>
            </xsl:when>
            <xsl:when test="@draw:fill-color">
              <xsl:call-template name ="getFillColor">
                <xsl:with-param name ="fill-color" select ="@draw:fill-color" />
                <xsl:with-param name ="opacity" select ="$opacity" />
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:for-each>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!-- Get line color and style-->
  <xsl:template name ="getLineStyle">
    <xsl:param name ="parentStyle" />
    <a:ln>
      <!-- Border width -->
      <xsl:choose>
        <xsl:when test ="@svg:stroke-width">
          <xsl:attribute name ="w">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length">
                <xsl:call-template name="convertUnitsToCm">
                  <xsl:with-param name="length" select="@svg:stroke-width"/>
                </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:when>
        <!-- Default border width from styles.xml-->
        <xsl:when test ="($parentStyle != '')">
          <xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $parentStyle]/style:graphic-properties">
            <xsl:if test ="@svg:stroke-width">
              <xsl:attribute name ="w">
                <xsl:call-template name ="convertToPoints">
                   <xsl:with-param name ="unit" select ="'cm'"/>
                  <xsl:with-param name ="length">
                    <xsl:call-template name="convertUnitsToCm">
                      <xsl:with-param name="length"  select ="@svg:stroke-width"/>
                    </xsl:call-template>
                  </xsl:with-param>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if >
            <!-- Code for the Bug 1746356 -->
            <xsl:if test ="not(@svg:stroke-width)">
              <xsl:attribute name ="w">
                <xsl:value-of select ="'0'"/>
              </xsl:attribute>
            </xsl:if >
            <!-- Code for the Bug 1746356 -->
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
      <!-- Cap type-->
      <xsl:if test ="@draw:stroke-dash">
        <xsl:variable name ="dash" select ="@draw:stroke-dash" />
        <xsl:if test ="document('styles.xml')/office:document-styles/office:styles/draw:stroke-dash[@draw:name=$dash]">
          <xsl:if test ="document('styles.xml')/office:document-styles/office:styles/draw:stroke-dash[@draw:name=$dash]/@draw:style='round'">
            <xsl:attribute name ="cap">
              <xsl:value-of select ="'rnd'"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>
      </xsl:if>
      <xsl:call-template name ="tmpLinefillColor">
        <xsl:with-param name ="parentStyle" select="$parentStyle" />
                </xsl:call-template>
             <!-- Dash type-->
      <xsl:choose>
        <xsl:when test ="(@draw:stroke='dash')">
          <a:prstDash>
            <xsl:attribute name ="val">
              <xsl:call-template name ="getDashType">
                <xsl:with-param name ="stroke-dash" select ="@draw:stroke-dash" />
              </xsl:call-template>
            </xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test ="not(@draw:stroke) and $parentStyle!=''">
          <xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $parentStyle]/style:graphic-properties">
            <xsl:if test="@draw:stroke='dash'">
        <a:prstDash>
          <xsl:attribute name ="val">
            <xsl:call-template name ="getDashType">
              <xsl:with-param name ="stroke-dash" select ="@draw:stroke-dash" />
            </xsl:call-template>
          </xsl:attribute>
        </a:prstDash>
      </xsl:if>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
      <!-- Line join type-->
      <xsl:if test ="@draw:stroke-linejoin">
        <xsl:call-template name ="getJoinType">
          <xsl:with-param name ="stroke-linejoin" select ="@draw:stroke-linejoin" />
        </xsl:call-template>
      </xsl:if>
      <!-- added by yeswanth , Fix for join type , bug# 1811327 -->
      <xsl:if test="not(./@draw:stroke-linejoin)">
        <xsl:if test="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name='standard']/style:graphic-properties/@draw:stroke-linejoin">
          <xsl:call-template name ="getJoinType">
            <xsl:with-param name ="stroke-linejoin" select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name='standard']/style:graphic-properties/@draw:stroke-linejoin"/>
          </xsl:call-template>
        </xsl:if>
      </xsl:if>
      <!-- End -->

      <!--Arrow type-->
      <xsl:if test="(@draw:marker-start) and (@draw:marker-start != '')">
        <a:headEnd>
          <xsl:attribute name ="type">
            <xsl:call-template name ="getArrowType">
              <xsl:with-param name ="ArrowType" select ="@draw:marker-start" />
            </xsl:call-template>
          </xsl:attribute>
          <xsl:if test ="@draw:marker-start-width">
            <xsl:variable name="Unit">
              <xsl:call-template name="getConvertUnit">
                <xsl:with-param name="length" select="@draw:marker-start-width"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:call-template name ="setArrowSize">
              <xsl:with-param name ="size" select ="substring-before(@draw:marker-start-width,$Unit)" />
            </xsl:call-template >
          </xsl:if>
        </a:headEnd>
      </xsl:if>
      <xsl:if test="not(@draw:marker-start) and not(@draw:marker-start != '') and $parentStyle !=''">
        <xsl:for-each select ="document('styles.xml')//style:style[@style:name = $parentStyle]/style:graphic-properties">
          <xsl:if test="@draw:marker-start">
            <a:headEnd>
              <xsl:attribute name ="type">
                <xsl:call-template name ="getArrowType">
                  <xsl:with-param name ="ArrowType" select ="@draw:marker-start" />
                </xsl:call-template>
              </xsl:attribute>
              <xsl:if test ="@draw:marker-start-width">
                <xsl:variable name="Unit">
                  <xsl:call-template name="getConvertUnit">
                    <xsl:with-param name="length" select="@draw:marker-start-width"/>
                  </xsl:call-template>
                </xsl:variable>
                <xsl:call-template name ="setArrowSize">
                  <xsl:with-param name ="size" select ="substring-before(@draw:marker-start-width,$Unit)" />
                </xsl:call-template >
              </xsl:if>
            </a:headEnd>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
      <xsl:if test="(@draw:marker-end) and (@draw:marker-end != '')">
        <a:tailEnd>
          <xsl:attribute name ="type">
            <xsl:call-template name ="getArrowType">
              <xsl:with-param name ="ArrowType" select ="@draw:marker-end" />
            </xsl:call-template>
          </xsl:attribute>

          <xsl:if test ="@draw:marker-end-width">
            <xsl:variable name="Unit">
              <xsl:call-template name="getConvertUnit">
                <xsl:with-param name="length" select="@draw:marker-end-width"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:call-template name ="setArrowSize">
              <xsl:with-param name ="size" select ="substring-before(@draw:marker-end-width,$Unit)" />
            </xsl:call-template >
          </xsl:if>
        </a:tailEnd>
      </xsl:if>
      <xsl:if test="not(@draw:marker-end) and not(@draw:marker-end != '') and $parentStyle !=''">
        <xsl:for-each select ="document('styles.xml')//style:style[@style:name = $parentStyle]/style:graphic-properties">
          <xsl:if test="@draw:marker-end">
            <a:tailEnd>
              <xsl:attribute name ="type">
                <xsl:call-template name ="getArrowType">
                  <xsl:with-param name ="ArrowType" select ="@draw:marker-end" />
                </xsl:call-template>
              </xsl:attribute>
              <xsl:variable name="Unit">
                <xsl:call-template name="getConvertUnit">
                  <xsl:with-param name="length" select="@draw:marker-end-width"/>
                </xsl:call-template>
              </xsl:variable>
              <xsl:if test ="@draw:marker-end-width">
                <xsl:call-template name ="setArrowSize">
                  <xsl:with-param name ="size" select ="substring-before(@draw:marker-end-width,$Unit)" />
                </xsl:call-template >
              </xsl:if>
            </a:tailEnd>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </a:ln>
  </xsl:template>

	<!-- Get Shadow Effect -->
	<xsl:template name="getShadowEffect">
		<xsl:param name ="parentStyle" />
		<xsl:param name="shadowOffsetX" />
		<xsl:param name="shadowOffsetY" />
		<xsl:param name="shadowColor" />
		<xsl:param name="shadowOpacity" />

		<a:effectLst>
			<a:outerShdw>
				<xsl:attribute name="blurRad">
					<xsl:value-of select="50800"/>
				</xsl:attribute>

				<!-- the if condition is to check for center selections -->
				<xsl:if test="not($shadowOffsetX =  0 and $shadowOffsetY = 0) and ($shadowOffsetX !='' and $shadowOffsetY !='')">
					<xsl:variable name ="distVal" >
						<xsl:value-of select ="$shadowOffsetX"/>
					</xsl:variable>
					<xsl:variable name ="dirVal" >
						<xsl:value-of select ="$shadowOffsetY"/>
					</xsl:variable>

					<xsl:attribute name ="dist">
						<xsl:value-of select ="concat('a-outerShdw-dist:',$distVal, ':', $dirVal)"/>
					</xsl:attribute>

					<xsl:attribute name ="dir">
						<xsl:value-of select ="concat('a-outerShdw-dir:',$distVal, ':', $dirVal)"/>
					</xsl:attribute>
				</xsl:if>

				<!-- center-->
				<xsl:if test="($shadowOffsetX =  0 and $shadowOffsetY = 0)">

					<xsl:attribute name ="sx">
						<xsl:value-of select ="'102000'"/>
					</xsl:attribute>
					<xsl:attribute name ="sy">
						<xsl:value-of select ="'102000'"/>
					</xsl:attribute>
				</xsl:if>

				<xsl:if test="not($shadowOffsetX =  0 and $shadowOffsetY &lt; 0)">
					<xsl:attribute name ="algn">

						<xsl:choose>
							<xsl:when test="$shadowOffsetX &lt;  0 and $shadowOffsetY &lt; 0">
								<xsl:value-of select ="'br'"/>
							</xsl:when>
							<xsl:when test="$shadowOffsetX &gt;  0 and $shadowOffsetY &lt; 0">
								<xsl:value-of select ="'bl'"/>
							</xsl:when>
							<xsl:when test="$shadowOffsetX &lt;  0 and $shadowOffsetY &gt; 0">
								<xsl:value-of select ="'tr'"/>
							</xsl:when>
							<xsl:when test="$shadowOffsetX &gt;  0 and $shadowOffsetY &gt; 0">
								<xsl:value-of select ="'tl'"/>
							</xsl:when>
							<xsl:when test="$shadowOffsetX &lt;  0">
								<xsl:value-of select ="'r'"/>
							</xsl:when>
							<xsl:when test="$shadowOffsetX &gt;  0">
								<xsl:value-of select ="'l'"/>
							</xsl:when>

							<xsl:when test="$shadowOffsetY &gt; 0">
								<xsl:value-of select ="'t'"/>
							</xsl:when>

							<xsl:otherwise>
								<xsl:value-of select ="'ctr'"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
				</xsl:if>
				<xsl:attribute name ="rotWithShape">
					<xsl:value-of select ="0"/>
				</xsl:attribute>

				<a:srgbClr>
					<xsl:attribute name="val">
						<xsl:value-of select="$shadowColor"/>
					</xsl:attribute>

					<a:alpha>
						<xsl:attribute name="val">
              <!--<xsl:value-of select="$shadowOpacity"/>-->
              <xsl:choose>
                <xsl:when test = "$shadowOpacity != ''">
							<xsl:value-of select="$shadowOpacity"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'0'"/>
                </xsl:otherwise>
               </xsl:choose>
						</xsl:attribute>
					</a:alpha>
				</a:srgbClr>
			</a:outerShdw>
		</a:effectLst>
	</xsl:template>
	
  <!-- Get color code -->
  <xsl:template name ="getFillColor">
    <xsl:param name ="fill-color" />
    <xsl:param name ="opacity" />
    <xsl:if test ="$fill-color != ''">
      <a:solidFill>
        <a:srgbClr>
          <xsl:attribute name ="val">
            <xsl:value-of select ="substring-after($fill-color,'#')"/>
          </xsl:attribute>
          <xsl:if test ="$opacity != ''">
            <a:alpha>
              <xsl:variable name ="alpha" select ="substring-before($opacity,'%')" />
              <xsl:attribute name ="val">
                <xsl:if test ="$alpha = 0">
                  <xsl:value-of select ="0000"/>
                </xsl:if>
                <xsl:if test ="$alpha != 0">
                  <xsl:value-of select ="$alpha * 1000"/>
                </xsl:if>
              </xsl:attribute>
            </a:alpha>
          </xsl:if>
        </a:srgbClr>
      </a:solidFill>
    </xsl:if>
  </xsl:template>
  <!-- Get shade for fill reference-->
  <xsl:template name ="getShade">
    <xsl:param name ="fileName" />
    <xsl:param name ="gr" />
    <xsl:if test ="document($fileName)//office:automatic-styles/style:style[@style:name=$gr]/style:graphic-properties/@draw:shadow-opacity">
      <xsl:variable name ="shade">
        <xsl:value-of select ="document($fileName)//office:automatic-styles/style:style[@style:name=$gr]/style:graphic-properties/@draw:shadow-opacity"/>
      </xsl:variable>
      <a:shade>
        <xsl:variable name ="alpha" select ="substring-before($shade,'%')" />
        <xsl:attribute name ="val">
          <xsl:if test ="$alpha = 0">
            <xsl:value-of select ="0000"/>
          </xsl:if>
          <xsl:if test ="$alpha != 0">
            <xsl:value-of select ="$alpha * 1000"/>
          </xsl:if>
        </xsl:attribute>
      </a:shade>
    </xsl:if>
  </xsl:template>
  <!-- Get arrow type-->
  <xsl:template name ="setArrowSize">
    <xsl:param name ="size" />
    <xsl:choose>
      <xsl:when test ="($size &lt;= $sm-sm)">
        <xsl:attribute name ="w">
          <xsl:value-of select ="'sm'"/>
        </xsl:attribute>
        <xsl:attribute name ="len">
          <xsl:value-of select ="'sm'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="($size &gt; $sm-sm) and ($size &lt;= $sm-med)">
        <xsl:attribute name ="w">
          <xsl:value-of select ="'sm'"/>
        </xsl:attribute>
        <xsl:attribute name ="len">
          <xsl:value-of select ="'med'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="($size &gt; $sm-med) and ($size &lt;= $sm-lg)">
        <xsl:attribute name ="w">
          <xsl:value-of select ="'sm'"/>
        </xsl:attribute>
        <xsl:attribute name ="len">
          <xsl:value-of select ="'lg'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="($size &gt; $sm-lg) and ($size &lt;= $med-sm)">
        <xsl:attribute name ="w">
          <xsl:value-of select ="'med'"/>
        </xsl:attribute>
        <xsl:attribute name ="len">
          <xsl:value-of select ="'sm'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="($size &gt; $med-med) and ($size &lt;= $med-lg)">
        <xsl:attribute name ="w">
          <xsl:value-of select ="'med'"/>
        </xsl:attribute>
        <xsl:attribute name ="len">
          <xsl:value-of select ="'lg'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="($size &gt; $med-lg) and ($size &lt;= $lg-sm)">
        <xsl:attribute name ="w">
          <xsl:value-of select ="'lg'"/>
        </xsl:attribute>
        <xsl:attribute name ="len">
          <xsl:value-of select ="'sm'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="($size &gt; $lg-sm) and ($size &lt;= $lg-med)">
        <xsl:attribute name ="w">
          <xsl:value-of select ="'lg'"/>
        </xsl:attribute>
        <xsl:attribute name ="len">
          <xsl:value-of select ="'med'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test ="($size &gt; $lg-med) and ($size &lt;= $lg-lg)">
        <xsl:attribute name ="w">
          <xsl:value-of select ="'lg'"/>
        </xsl:attribute>
        <xsl:attribute name ="len">
          <xsl:value-of select ="'lg'"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name ="w">
          <xsl:value-of select ="'med'"/>
        </xsl:attribute>
        <xsl:attribute name ="len">
          <xsl:value-of select ="'med'"/>
        </xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>
  <!-- Get line join type-->
  <xsl:template name ="getJoinType">
    <xsl:param name ="stroke-linejoin" />
    <xsl:choose>
      <xsl:when test ="$stroke-linejoin='miter'">
        <a:miter lim="800000" />
      </xsl:when>
      <xsl:when test ="$stroke-linejoin='bevel'">
        <a:bevel/>
      </xsl:when>
      <xsl:when test ="$stroke-linejoin='round'">
        <a:round/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!-- Get dash type-->
  <xsl:template name ="getDashType">
    <xsl:param name ="stroke-dash" />
    <xsl:if test ="document('styles.xml')/office:document-styles/office:styles/draw:stroke-dash[@draw:name=$stroke-dash]">
      <xsl:for-each select="document('styles.xml')/office:document-styles/office:styles/draw:stroke-dash[@draw:name=$stroke-dash]">
        <xsl:if test="position()=1">
        <xsl:variable name="Unit1">
          <xsl:call-template name="getConvertUnit">
            <xsl:with-param name="length" select="@draw:dots1-length"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="Unit2">
          <xsl:call-template name="getConvertUnit">
            <xsl:with-param name="length" select="@draw:dots2-length"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="Unit3">
          <xsl:call-template name="getConvertUnit">
            <xsl:with-param name="length" select="@draw:distance"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name ="dots1" select="@draw:dots1"/>
        <xsl:variable name ="dots1-length" select ="substring-before(@draw:dots1-length, $Unit1)" />
        <xsl:variable name ="dots2" select="@draw:dots2"/>
        <xsl:variable name ="dots2-length" select ="substring-before(@draw:dots2-length, $Unit2)"/>
        <xsl:variable name ="distance" select ="substring-before(@draw:distance, $Unit3)" />



      <xsl:choose>
        <xsl:when test ="($dots1=1) and ($dots2=1)">
          <xsl:choose>
            <xsl:when test ="($dots1-length &lt;= $dot) and ($dots2-length &lt;= $dot)">
              <xsl:value-of select ="'sysDot'" />
            </xsl:when>
            <xsl:when test ="(($dots1-length &lt;= $dot) and ($dots2-length &lt;= $dash)) or
										 (($dots1-length &lt;= $dash) and ($dots2-length &lt;= $dot)) ">
              <xsl:value-of select ="'dashDot'" />
            </xsl:when>
            <xsl:when test ="($dots1-length &lt;= $dash) and ($dots2-length &lt;= $dash)">
              <xsl:value-of select ="'dash'" />
            </xsl:when>
            <xsl:when test ="(($dots1-length &lt;= $dot) and (($dots2-length &gt;= $dash) and ($dots2-length &lt;= $longDash))) or
										 (($dots2-length &lt;= $dot) and (($dots1-length &gt;= $dash) and ($dots1-length &lt;= $longDash))) ">
              <xsl:value-of select ="'lgDashDot'" />
            </xsl:when>
            <xsl:when test ="(($dots1-length &gt;= $dash) and ($dots1-length &lt;= $longDash)) or 
										 (($dots2-length &gt;= $dash) and ($dots2-length &lt;= $longDash))">
              <xsl:value-of select ="'lgDash'" />
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select ="'sysDash'" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:when test ="(($dots1=2) and ($dots2=1)) or (($dots1=1) and ($dots2=2))">
          <!--<xsl:if test ="(($dots1-length &lt;= $dot) and ($dots2-length &gt;= $dash) and ($dots2-length &lt;= $longDash))">-->
          <xsl:value-of select ="'lgDashDotDot'" />
          <!--</xsl:if>-->
        </xsl:when>
        <!--<xsl:when test ="($dots1=1) and ($dots2=2)">
					-->
        <!--<xsl:if test ="(($dots2-length &lt;= $dot) and ($dots1-length &gt;= $dash) and ($dots1-length &lt;= $longDash))">-->
        <!--
						<xsl:value-of select ="'lgDashDotDot'" />
					-->
        <!--</xsl:if>-->
        <!--
				</xsl:when>-->
        <xsl:when test ="(($dots1 &gt;= 1) and not($dots2))">
          <xsl:choose>
            <xsl:when test ="($dots1-length &lt;= $dot)">
              <xsl:value-of select ="'sysDash'" />
            </xsl:when>
            <xsl:when test ="($dots1-length &gt;= $dot) and ($dots1-length &lt;= $dash)">
              <xsl:value-of select ="'dash'" />
            </xsl:when>
            <xsl:when test ="($dots1-length &gt;= $dash) and ($dots1-length &lt;= $longDash)">
              <xsl:value-of select ="'lgDash'" />
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select ="'sysDash'" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:when test ="(($dots2 &gt;= 1) and not($dots1))">
          <xsl:choose>
            <xsl:when test ="($dots2-length &lt;= $dot)">
              <xsl:value-of select ="'sysDash'" />
            </xsl:when>
            <xsl:when test ="($dots2-length &gt;= $dot) and ($dots2-length &lt;= $dash)">
              <xsl:value-of select ="'dash'" />
            </xsl:when>
            <xsl:when test ="($dots2-length &gt;= $dash) and ($dots2-length &lt;= $longDash)">
              <xsl:value-of select ="'lgDash'" />
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select ="'sysDash'" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select ="'sysDash'" />
        </xsl:otherwise>
      </xsl:choose>
        </xsl:if>
      </xsl:for-each>
    </xsl:if >
      <xsl:if test ="not(document('styles.xml')/office:document-styles/office:styles/draw:stroke-dash[@draw:name=$stroke-dash])">
        <xsl:value-of select ="'sysDash'" />
      </xsl:if>
   </xsl:template>
  <!-- Get text layout for shape-->
  <xsl:template name ="getParagraphProperties">
    <xsl:param name ="fileName" />
    <xsl:param name ="gr" />
    <xsl:param name ="customShape" />

    <a:bodyPr>

		<xsl:if test="./a:normAutofit">
			<xsl:message terminate="no">translation.odf2oox.shapeTopBottomWrapping</xsl:message>
		</xsl:if>
      <!-- Text direction-->
      <xsl:for-each select ="document($fileName)//office:automatic-styles/style:style[@style:name=$gr]/style:paragraph-properties">
        <xsl:if test ="@style:writing-mode">
          <xsl:attribute name ="vert">
            <xsl:choose>
              <xsl:when test ="@style:writing-mode='tb-rl'">
                <xsl:value-of select ="'vert'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select ="'horz'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>
      </xsl:for-each>

		<!--Added by Mathi for TextAlign in Textbox-->
      <xsl:for-each select ="document($fileName)/child::node()[1]/office:automatic-styles/style:style[@style:name=$gr]/style:graphic-properties">
        <!-- Vertical alignment -->
        <xsl:if test="./parent::node()/style:paragraph-properties/@style:writing-mode='tb-rl'">
          <xsl:choose>
            <xsl:when test ="(@draw:textarea-vertical-align='top' and @draw:textarea-horizontal-align='left')">
              <xsl:attribute name ="anchor">
                <xsl:value-of select ="'b'"/>
              </xsl:attribute>
              <xsl:attribute name ="anchorCtr">
                <xsl:value-of select ="'0'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test ="(@draw:textarea-vertical-align='top' and @draw:textarea-horizontal-align='right')">
              <xsl:attribute name ="anchor">
                <xsl:value-of select ="'t'"/>
              </xsl:attribute>
              <xsl:attribute name ="anchorCtr">
                <xsl:value-of select ="'0'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test ="(@draw:textarea-vertical-align='top' and @draw:textarea-horizontal-align='center')">
              <xsl:attribute name ="anchor">
                <xsl:value-of select ="'ctr'"/>
              </xsl:attribute>
              <xsl:attribute name ="anchorCtr">
                <xsl:value-of select ="'0'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test ="(@draw:textarea-vertical-align='middle' and @draw:textarea-horizontal-align='right')">
              <xsl:attribute name ="anchor">
                <xsl:value-of select ="'t'"/>
              </xsl:attribute>
              <xsl:attribute name ="anchorCtr">
                <xsl:value-of select ="'1'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test ="(@draw:textarea-vertical-align='middle' and @draw:textarea-horizontal-align='center')">
              <xsl:attribute name ="anchor">
                <xsl:value-of select ="'ctr'"/>
              </xsl:attribute>
              <xsl:attribute name ="anchorCtr">
                <xsl:value-of select ="'1'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test ="(@draw:textarea-vertical-align='middle' and @draw:textarea-horizontal-align='left')">
              <xsl:attribute name ="anchor">
                <xsl:value-of select ="'b'"/>
              </xsl:attribute>
              <xsl:attribute name ="anchorCtr">
                <xsl:value-of select ="'1'"/>
              </xsl:attribute>
            </xsl:when>
          </xsl:choose>
        </xsl:if>
        <!-- Horizontal alignment -->
        <xsl:if test="not(./parent::node()/style:paragraph-properties/@style:writing-mode='tb-rl')">
          <xsl:choose>
            <xsl:when test ="(@draw:textarea-vertical-align='middle' and @draw:textarea-horizontal-align='center')">
				  <xsl:attribute name ="anchor">
					  <xsl:value-of select ="'ctr'"/>
				  </xsl:attribute>
				  <xsl:attribute name ="anchorCtr">
                <xsl:value-of select ="'1'"/>
				  </xsl:attribute>
			  </xsl:when>
            <xsl:when test ="(@draw:textarea-vertical-align='bottom' and @draw:textarea-horizontal-align='center')">
				  <xsl:attribute name ="anchor">
					  <xsl:value-of select ="'b'"/>
				  </xsl:attribute>
				  <xsl:attribute name ="anchorCtr">
                <xsl:value-of select ="'1'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test ="(@draw:textarea-vertical-align='top' and @draw:textarea-horizontal-align='center')">
              <xsl:attribute name ="anchor">
                <xsl:value-of select ="'t'"/>
              </xsl:attribute>
              <xsl:attribute name ="anchorCtr">
                <xsl:value-of select ="'1'"/>
              </xsl:attribute>
            </xsl:when>
			      <xsl:when test ="@draw:textarea-vertical-align='middle'">  
              <xsl:attribute name ="anchor">
                <xsl:value-of select ="'ctr'"/>
              </xsl:attribute>
              <xsl:attribute name ="anchorCtr">
					  <xsl:value-of select ="'0'"/>
              </xsl:attribute>
            </xsl:when>
			      <xsl:when test ="@draw:textarea-vertical-align='bottom'"> 
              <xsl:attribute name ="anchor">
                <xsl:value-of select ="'b'"/>
              </xsl:attribute>
              <xsl:attribute name ="anchorCtr">
                <xsl:value-of select="'0'"/>
				  </xsl:attribute>
			  </xsl:when>
            <xsl:when test ="@draw:textarea-vertical-align='top'">
				  <xsl:attribute name ="anchor">
					  <xsl:value-of select="'t'"/>
			      </xsl:attribute>
				  <xsl:attribute name ="anchorCtr">
                <xsl:value-of select ="'0'"/>
				  </xsl:attribute>
			  </xsl:when>
          </xsl:choose>
        </xsl:if>


        <!-- Default text style-->
        <xsl:variable name ="parentStyle">
          <xsl:value-of select ="parent::node()/@style:parent-style-name"/>
        </xsl:variable>

      

      
       
        <xsl:variable name ="default-wrap-option">
          <xsl:call-template name ="getDefaultStyle">
            <xsl:with-param name ="parentStyle" select ="$parentStyle" />
            <xsl:with-param name ="attributeName" select ="'wrap-option'" />
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name ="default-autogrow-height">
          <xsl:call-template name ="getDefaultStyle">
            <xsl:with-param name ="parentStyle" select ="$parentStyle" />
            <xsl:with-param name ="attributeName" select ="'auto-grow-height'" />
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name ="default-autogrow-width">
          <xsl:call-template name ="getDefaultStyle">
            <xsl:with-param name ="parentStyle" select ="$parentStyle" />
            <xsl:with-param name ="attributeName" select ="'auto-grow-width'" />
          </xsl:call-template>
        </xsl:variable>

        <!--Text margins in shape-->
        <xsl:if test ="@fo:padding-left">


          <xsl:attribute name ="lIns">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length">
                <xsl:call-template name="convertUnitsToCm">
              <xsl:with-param name ="length" select ="@fo:padding-left"/>
            </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="not(@fo:padding-left)">
          <xsl:variable name ="default-padding-left">
            <xsl:call-template name ="getDefaultStyle">
              <xsl:with-param name ="parentStyle" select ="$parentStyle" />
              <xsl:with-param name ="attributeName" select ="'padding-left'" />
            </xsl:call-template>
          </xsl:variable>
          <xsl:attribute name ="lIns">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length">
                <xsl:call-template name="convertUnitsToCm">
              <xsl:with-param name ="length" select ="$default-padding-left"/>
            </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>


        <xsl:if test ="@fo:padding-top">
          <xsl:attribute name ="tIns">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length">
                <xsl:call-template name="convertUnitsToCm">
              <xsl:with-param name ="length" select ="@fo:padding-top"/>
            </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="not(@fo:padding-top)">
          <xsl:variable name ="default-padding-top">
            <xsl:call-template name ="getDefaultStyle">
              <xsl:with-param name ="parentStyle" select ="$parentStyle" />
              <xsl:with-param name ="attributeName" select ="'padding-top'" />
            </xsl:call-template>
          </xsl:variable>
          <xsl:attribute name ="tIns">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length">
                <xsl:call-template name="convertUnitsToCm">
              <xsl:with-param name ="length" select ="$default-padding-top"/>
            </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>


        <xsl:if test ="@fo:padding-right">
          <xsl:attribute name ="rIns">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length">
                <xsl:call-template name="convertUnitsToCm">
              <xsl:with-param name ="length" select ="@fo:padding-right"/>
            </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="not(@fo:padding-right)">
          <xsl:variable name ="default-padding-right">
            <xsl:call-template name ="getDefaultStyle">
              <xsl:with-param name ="parentStyle" select ="$parentStyle" />
              <xsl:with-param name ="attributeName" select ="'padding-right'" />
            </xsl:call-template>
          </xsl:variable>
          <xsl:attribute name ="rIns">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length">
                <xsl:call-template name="convertUnitsToCm">
              <xsl:with-param name ="length" select ="$default-padding-right"/>
            </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>


        <xsl:if test ="@fo:padding-bottom">
          <xsl:attribute name ="bIns">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length">
                <xsl:call-template name="convertUnitsToCm">
              <xsl:with-param name ="length" select ="@fo:padding-bottom"/>
            </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="not(@fo:padding-bottom)">
          <xsl:variable name ="default-padding-bottom">
            <xsl:call-template name ="getDefaultStyle">
              <xsl:with-param name ="parentStyle" select ="$parentStyle" />
              <xsl:with-param name ="attributeName" select ="'padding-bottom'" />
            </xsl:call-template>
          </xsl:variable>
          <xsl:attribute name ="bIns">
            <xsl:call-template name ="convertToPoints">
              <xsl:with-param name ="unit" select ="'cm'"/>
              <xsl:with-param name ="length">
                <xsl:call-template name="convertUnitsToCm">
              <xsl:with-param name ="length" select ="$default-padding-bottom"/>
            </xsl:call-template>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
        <!-- Text wrapping in shape-->
        <xsl:choose>
					<xsl:when test ="@fo:wrap-option='no-wrap'">
						<xsl:attribute name ="wrap">
							<xsl:value-of select ="'square'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test ="@fo:wrap-option='wrap'">
						<xsl:attribute name ="wrap">
							<xsl:value-of select ="'none'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test ="$default-wrap-option='no-wrap'">
						<xsl:attribute name ="wrap">
							<xsl:value-of select ="'square'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test ="$default-wrap-option='wrap'">
						<xsl:attribute name ="wrap">
							<xsl:value-of select ="'none'"/>
						</xsl:attribute>
					</xsl:when>
      						<xsl:otherwise>
						<xsl:attribute name ="wrap">
              <xsl:value-of select ="'square'"/>
						</xsl:attribute>
					</xsl:otherwise>  
				</xsl:choose>
        <xsl:choose>
          <xsl:when test ="@draw:auto-grow-height = 'true' or @draw:auto-grow-width = 'true'">
					<a:spAutoFit/>
          </xsl:when>
          <xsl:when test ="$default-autogrow-height = 'true' or $default-autogrow-width = 'true'">
            <a:spAutoFit/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </a:bodyPr>
    </xsl:template>
  <xsl:template name="tmpWrap">
        <xsl:choose>
          <xsl:when test ="@fo:wrap-option='no-wrap'">
            <xsl:attribute name ="wrap">
              <xsl:value-of select ="'square'"/>
            </xsl:attribute>
          </xsl:when>
      <xsl:when test ="@fo:wrap-option='wrap'">
            <xsl:attribute name ="wrap">
              <xsl:value-of select ="'none'"/>
            </xsl:attribute>
          </xsl:when>
               <xsl:otherwise>
            <xsl:attribute name ="wrap">
          <xsl:value-of select ="'square'"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      <xsl:if test ="@draw:auto-grow-height = 'true' or @draw:auto-grow-width = 'true'">
          <a:spAutoFit/>
        </xsl:if>
  
  </xsl:template>
  <!-- Get arrow type-->
    <xsl:template name ="getArrowType">
    <xsl:param name ="ArrowType" />

    <xsl:if test ="document('styles.xml')/office:document-styles/office:styles/draw:marker[@draw:name=$ArrowType]">
      <xsl:variable name ="marker" select ="document('styles.xml')/office:document-styles/office:styles/draw:marker[@draw:name=$ArrowType]/@svg:d"/>

      <xsl:choose>
        <!-- Arrow-->
        <xsl:when test ="$marker = 'm0 2108v17 17l12 42 30 34 38 21 43 4 29-8 30-21 25-26 13-34 343-1532 339 1520 13 42 29 34 39 21 42 4 42-12 34-30 21-42v-39-12l-4 4-440-1998-9-42-25-39-38-25-43-8-42 8-38 25-26 39-8 42z'">
          <xsl:value-of select ="'arrow'" />
        </xsl:when>
        <!-- Stealth-->
        <xsl:when test ="$marker = 'm1013 1491 118 89-567-1580-564 1580 114-85 136-68 148-46 161-17 161 13 153 46z'">
          <xsl:value-of select ="'stealth'" />
        </xsl:when>
        <!-- Oval-->
        <xsl:when test ="$marker = 'm462 1118-102-29-102-51-93-72-72-93-51-102-29-102-13-105 13-102 29-106 51-102 72-89 93-72 102-50 102-34 106-9 101 9 106 34 98 50 93 72 72 89 51 102 29 106 13 102-13 105-29 102-51 102-72 93-93 72-98 51-106 29-101 13z'">
          <xsl:value-of select ="'oval'"/>
        </xsl:when>
        <!-- Diamond-->
        <xsl:when test ="$marker = 'm0 564 564 567 567-567-567-564z'">
          <xsl:value-of select ="'diamond'"/>
        </xsl:when>
        <!-- Coded by pradeep for the bug 1742753 -->
        <!-- square to Diamond -->
        <xsl:when test ="$marker = 'm0 0h10v10h-10z'">
          <xsl:value-of select ="'diamond'"/>
        </xsl:when>
        <!-- End of Code by pradeep for the bug 1742753 -->
        <xsl:otherwise>
          <xsl:value-of select ="'triangle'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>

    <xsl:if test ="not(document('styles.xml')/office:document-styles/office:styles/draw:marker[@draw:name=$ArrowType])">
      <xsl:value-of  select ="'none'"/>
    </xsl:if >
  </xsl:template>
  <!-- Text formatting-->
  <xsl:template name ="processShapeText" >
    <xsl:param name ="fileName" />
    <xsl:param name ="shapeType" />
    <xsl:param name ="shapeCount" />
    <xsl:param name ="gr" />
    <xsl:param name ="shapeName" />
    <xsl:param name ="grpFlag" />
    <xsl:param name ="UniqueId" />
    <!--<xsl:param name ="ShapeType" />-->
    <!-- Paremeter added by vijayeta,get master page name, dated:11-7-07-->
    <xsl:param name ="masterPageName"/>
    <xsl:variable name ="finalFileName">
      <xsl:if test ="$fileName!=''">
        <xsl:value-of select ="$fileName"/>
      </xsl:if>
      <xsl:if test ="$fileName=''">
        <xsl:value-of select ="'content.xml'"/>
      </xsl:if>
    </xsl:variable>
    <xsl:variable name="shapeId">
      <xsl:value-of select="concat(substring-after($shapeType,'draw:'),$shapeCount)"/>
    </xsl:variable>
    <xsl:variable name ="prClsName">
      <xsl:value-of select ="document($finalFileName)//office:automatic-styles/style:style[@style:name=$gr]/@style:parent-style-name"/>
    </xsl:variable>
    <xsl:variable name ="defFontSize">
      <xsl:value-of  select ="office:document-styles/office:styles/style:style/style:text-properties/@fo:font-size"/>
    </xsl:variable>
    <xsl:for-each select ="node()">
      <xsl:choose>
        <xsl:when test ="name()='text:p'" >
        <xsl:variable name="aColonhlinkClick">
          <xsl:if test="$grpFlag !='true'">
          <xsl:if test="text:a">
            <a:hlinkClick>
              <xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'Link',position())"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'Link',position())"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'Link',position())"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>

          <xsl:if test="text:span/text:a">
            <a:hlinkClick>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'Link',position())"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'Link',position())"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'Link',position())"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
          </xsl:if>
        </xsl:variable>
        <xsl:if test ="child::node()">
          <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <xsl:call-template name ="paraProperties" >
              <xsl:with-param name ="paraId"  select="@text:style-name"/>
              <xsl:with-param name ="isBulleted" select ="'false'"/>
              <xsl:with-param name ="level" select ="'0'"/>
              <xsl:with-param name ="grpFlag" select ="$grpFlag"/>
              <xsl:with-param name="framePresentaionStyleId" select="parent::node()/parent::node()/./@presentation:style-name" />
              <xsl:with-param name ="isNumberingEnabled" select ="'false'"/>
              <xsl:with-param name ="slideMaster" select ="$fileName"/>
              <xsl:with-param name ="masterPageName" select ="$masterPageName"/>              
              <xsl:with-param name ="flagPresentationClass" select ="'No'"/>
              <xsl:with-param name ="parentStyleName" select ="$prClsName"/>
            </xsl:call-template >
            <xsl:for-each select ="child::node()[position()]">
              <xsl:choose >
                <xsl:when test ="name()='text:span'">
                  <xsl:choose>
                    <xsl:when test ="text:page-number">
                      <a:fld >
                        <xsl:attribute name ="id">
                          <xsl:value-of select ="'{763D1470-AB83-4C4C-B3B3-7F0C9DC8E8D6}'"/>
                        </xsl:attribute>
                        <xsl:attribute name ="type">
                          <xsl:value-of select ="'slidenum'"/>
                        </xsl:attribute>
                        <a:rPr lang="en-US" dirty="0" smtClean="0">
                          <!--Font Size -->
                          <xsl:variable name ="textId">
                            <xsl:value-of select ="@text:style-name"/>
                          </xsl:variable>
                          <xsl:if test ="not($textId ='')">
                            <xsl:call-template name ="fontStyles">
                              <xsl:with-param name ="Tid" select ="$textId" />
                              <xsl:with-param name ="prClassName" select ="$prClsName"/>
                              <xsl:with-param name ="slideMaster" select ="$fileName"/>
                              <xsl:with-param name ="fileName" select ="$fileName"/>
                              <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                              <xsl:with-param name ="flagPresentationClass" select ="'No'"/>
                              <xsl:with-param name ="parentStyleName" select ="$prClsName"/>
                            </xsl:call-template>
                          </xsl:if>
                          <xsl:if test ="$textId =''">
                            <xsl:variable name ="VarSz">
                              <xsl:call-template name ="getDefaultFontSize">
                                <xsl:with-param name ="className" select ="$prClsName"/>
                              </xsl:call-template >
                            </xsl:variable>
                            <xsl:if test ="$VarSz!=''">
                              <xsl:attribute name ="sz">
                                <xsl:value-of select ="substring-before($VarSz,'pt')*100"/>
                              </xsl:attribute>
                            </xsl:if>
                          </xsl:if>
                        </a:rPr>
                        <a:t>
                          <xsl:value-of select="."/>
                        </a:t>
                      </a:fld>
                    </xsl:when>
                    <xsl:when test ="text:date or presentation:date-time" >
                      <a:fld >
                        <xsl:attribute name ="id">
                          <xsl:value-of select ="'{86419996-E19B-43D7-A4AA-D671C2F15715}'"/>
                        </xsl:attribute>
                        <xsl:attribute name ="type">
                          <xsl:choose >
                            <xsl:when test ="@style:data-style-name ='D3'">
                              <xsl:value-of select ="'datetime1'"/>
                            </xsl:when>
                            <xsl:when test ="@style:data-style-name ='D8'">
                              <xsl:value-of select ="'datetime2'"/>
                            </xsl:when>
                            <xsl:when test ="@style:data-style-name ='D6'">
                              <xsl:value-of select ="'datetime4'"/>
                            </xsl:when>
                            <xsl:when test ="@style:data-style-name ='D5'">
                              <xsl:value-of select ="'datetime4'"/>
                            </xsl:when>
                            <xsl:when test ="@style:data-style-name ='D3T2'">
                              <xsl:value-of select ="'datetime8'"/>
                            </xsl:when>
                            <xsl:when test ="@style:data-style-name ='D3T5'">
                              <xsl:value-of select ="'datetime8'"/>
                            </xsl:when>
                            <xsl:when test ="@style:data-style-name ='T2'">
                              <xsl:value-of select ="'datetime10'"/>
                            </xsl:when>
                            <xsl:when test ="@style:data-style-name ='T3'">
                              <xsl:value-of select ="'datetime11'"/>
                            </xsl:when>
                            <xsl:when test ="@style:data-style-name ='T5'">
                              <xsl:value-of select ="'datetime12'"/>
                            </xsl:when>
                            <xsl:when test ="@style:data-style-name ='T6'">
                              <xsl:value-of select ="'datetime13'"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select ="'datetime1'"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
                        <a:rPr lang="en-US" dirty="0" smtClean="0">
                          <!--Font Size -->
                          <xsl:variable name ="textId">
                            <xsl:value-of select ="@text:style-name"/>
                          </xsl:variable>
                          <xsl:if test ="not($textId ='')">
                            <xsl:call-template name ="fontStyles">
                              <xsl:with-param name ="Tid" select ="$textId" />
                              <xsl:with-param name ="prClassName" select ="$prClsName"/>
                              <xsl:with-param name ="slideMaster" select ="$fileName"/>
                              <xsl:with-param name ="fileName" select ="$fileName"/>
                              <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                              <xsl:with-param name ="flagPresentationClass" select ="'No'"/>
                              <xsl:with-param name ="parentStyleName" select ="$prClsName"/>
                            </xsl:call-template>
                          </xsl:if>
                          <xsl:if test ="$textId =''">
                            <xsl:variable name ="VarSz">
                              <xsl:call-template name ="getDefaultFontSize">
                                <xsl:with-param name ="className" select ="$prClsName"/>
                              </xsl:call-template >
                            </xsl:variable>
                            <xsl:if test ="$VarSz!=''">
                              <xsl:attribute name ="sz">
                                <xsl:value-of select ="substring-before($VarSz,'pt')*100"/>
                              </xsl:attribute>
                            </xsl:if>
                          </xsl:if>
                        </a:rPr>
                        <a:t>
                          <xsl:value-of select ="."/>
                        </a:t>
                      </a:fld>
                    </xsl:when>
                    <xsl:otherwise>
                  <xsl:if test="node()">
                    <a:r>
                      <a:rPr lang="en-US" smtClean="0">
                        <!--Font Size -->
                        <xsl:variable name ="textId">
                          <xsl:value-of select ="@text:style-name"/>
                        </xsl:variable>
                        <xsl:if test ="not($textId ='')">
                          <xsl:call-template name ="fontStyles">
                            <xsl:with-param name ="Tid" select ="$textId" />
                            <xsl:with-param name ="prClassName" select ="$prClsName"/>
                            <xsl:with-param name ="slideMaster" select ="$fileName"/>
                            <xsl:with-param name ="fileName" select ="$fileName"/>
                            <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                            <xsl:with-param name ="flagPresentationClass" select ="'No'"/>
                            <xsl:with-param name ="parentStyleName" select ="$prClsName"/>
                          </xsl:call-template>
                        </xsl:if>
						<xsl:if test ="$textId =''">
							<xsl:variable name ="VarSz">
								<xsl:call-template name ="getDefaultFontSize">
									<xsl:with-param name ="className" select ="$prClsName"/>
								</xsl:call-template >
							</xsl:variable>
							<xsl:if test ="$VarSz!=''">
								<xsl:attribute name ="sz">
									<xsl:value-of select ="substring-before($VarSz,'pt')*100"/>
								</xsl:attribute>
							</xsl:if>
						</xsl:if>
                        <xsl:if test="./text:a">
                          <xsl:if test="not(parent::node()/parent::node()/office:event-listeners/presentation:event-listener)">
                        <xsl:copy-of select="$aColonhlinkClick"/>
                        </xsl:if>
                        </xsl:if>
                     </a:rPr >
                      <a:t>
                        <xsl:call-template name ="insertTab" />
                      </a:t>
                    </a:r>
                    <xsl:if test ="text:line-break">
                      <xsl:call-template name ="processBR">
                        <xsl:with-param name ="T" select ="@text:style-name" />
                        <xsl:with-param name ="prClassName" select ="$prClsName"/>
                      </xsl:call-template>
                    </xsl:if>
                  </xsl:if>
                  <!-- Added by lohith - bug fix 1731885 -->
                  <xsl:if test="not(node())">
                    <a:endParaRPr lang="en-US" smtClean="0">
                      <!--Font Size -->
                      <xsl:variable name ="textId">
                        <xsl:value-of select ="@text:style-name"/>
                      </xsl:variable>
						<xsl:if test ="$textId =''">
							<xsl:variable name ="VarSz">
								<xsl:call-template name ="getDefaultFontSize">
									<xsl:with-param name ="className" select ="$prClsName"/>
								</xsl:call-template >
							</xsl:variable>
							<xsl:if test ="$VarSz!=''">
								<xsl:attribute name ="sz">
									<xsl:value-of select ="substring-before($VarSz,'pt')*100"/>
								</xsl:attribute>
							</xsl:if>
						</xsl:if>
                      <xsl:if test ="not($textId ='')">
                        <xsl:call-template name ="fontStyles">
                          <xsl:with-param name ="Tid" select ="$textId" />
                          <xsl:with-param name ="prClassName" select ="$prClsName"/>
                          <xsl:with-param name ="slideMaster" select ="$fileName"/>
                          <xsl:with-param name ="flagPresentationClass" select ="'No'"/>
                          <xsl:with-param name ="parentStyleName" select ="$prClsName"/>
                        </xsl:call-template>
                      </xsl:if>
                    </a:endParaRPr>
                  </xsl:if>
                  </xsl:otherwise>
                  </xsl:choose>
                </xsl:when >
                <xsl:when test ="name()='text:line-break'">
                  <xsl:call-template name ="processBR">
                    <xsl:with-param name ="T" select ="@text:style-name" />
                    <xsl:with-param name ="prClassName" select ="$prClsName"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:when test ="name()='text:tab'">
                  <a:r>
                    <a:rPr lang="en-US" smtClean="0">
                      <!--Font Size -->
                      <xsl:variable name ="textId">
                        <xsl:value-of select ="./parent::node()/@text:style-name"/>
                      </xsl:variable>
                      <xsl:if test ="not($textId ='')">
                        <xsl:call-template name ="fontStyles">
                          <xsl:with-param name ="Tid" select ="$textId" />
                          <xsl:with-param name ="prClassName" select ="$prClsName"/>
                          <xsl:with-param name ="slideMaster" select ="$fileName"/>
                          <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                          <xsl:with-param name ="flagPresentationClass" select ="'No'"/>
                          <xsl:with-param name ="parentStyleName" select ="$prClsName"/>
                        </xsl:call-template>
                      </xsl:if>
                      <xsl:if test ="$textId =''">
                        <xsl:variable name ="VarSz">
                          <xsl:call-template name ="getDefaultFontSize">
                            <xsl:with-param name ="className" select ="$prClsName"/>
                          </xsl:call-template >
                        </xsl:variable>
                        <xsl:if test ="$VarSz!=''">
                          <xsl:attribute name ="sz">
                            <xsl:value-of select ="substring-before($VarSz,'pt')*100"/>
                          </xsl:attribute>
                        </xsl:if>
                      </xsl:if>
                      <xsl:if test ="$textId =''">
                        <a:latin charset="0"  >
                          <xsl:attribute name ="typeface">
                            <xsl:call-template name ="getDefaultFonaName">
                              <xsl:with-param name ="className" select ="$prClsName"/>
                              <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                            </xsl:call-template>
                          </xsl:attribute>
                        </a:latin >
                      </xsl:if >
                    </a:rPr >
                    <a:t>
                      <xsl:value-of select ="'&#09;'"/>
                    </a:t>
                  </a:r>
                 
                </xsl:when >
                <xsl:when test ="not(name()='text:span')">
                  <a:r>
                    <a:rPr lang="en-US" smtClean="0">
                      <!--Font Size -->
                      <xsl:variable name ="textId">
                        <xsl:value-of select ="./parent::node()/@text:style-name"/>
                      </xsl:variable>
                      <xsl:if test ="not($textId ='')">
                        <xsl:call-template name ="fontStyles">
                          <xsl:with-param name ="Tid" select ="$textId" />
                          <xsl:with-param name ="prClassName" select ="$prClsName"/>
                          <xsl:with-param name ="slideMaster" select ="$fileName"/>
                          <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                          <xsl:with-param name ="flagPresentationClass" select ="'No'"/>
                          <xsl:with-param name ="parentStyleName" select ="$prClsName"/>
                        </xsl:call-template>						  
                      </xsl:if>
						<xsl:if test ="$textId =''">
							<xsl:variable name ="VarSz">
								<xsl:call-template name ="getDefaultFontSize">
									<xsl:with-param name ="className" select ="$prClsName"/>
								</xsl:call-template >
							</xsl:variable>
							<xsl:if test ="$VarSz!=''">
								<xsl:attribute name ="sz">
									<xsl:value-of select ="substring-before($VarSz,'pt')*100"/>
								</xsl:attribute>
							</xsl:if>
						</xsl:if>
                      <!--<xsl:copy-of select="$varTextHyperLinks"/>-->
                      <xsl:if test ="$textId =''">
                        <a:latin charset="0"  >
                          <xsl:attribute name ="typeface">
                            <xsl:call-template name ="getDefaultFonaName">
                              <xsl:with-param name ="className" select ="$prClsName"/>
                              <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                            </xsl:call-template>
                          </xsl:attribute>
                        </a:latin >
                      </xsl:if >
                      <xsl:if test="not(parent::node()/parent::node()/office:event-listeners/presentation:event-listener)">
                        <xsl:copy-of select="$aColonhlinkClick"/>
                      </xsl:if>
                      <!--<xsl:copy-of select="$varTextHyperLinks"/>-->
                    
                    </a:rPr >
                    <a:t>
                      <xsl:call-template name ="insertTab" />
                    </a:t>
                  </a:r>
                </xsl:when >
              </xsl:choose>
            </xsl:for-each>
            <!--<xsl:copy-of select="$varFrameHyperLinks"/>-->
          </a:p>
        </xsl:if>
        <!-- Change made by vijayeta
        Bug fix	1731885
        Dated 20th July
        Added a condition to test if text:p has child nodes or not, if not, set a default font to the empty line-->
        <xsl:if test ="not(child::node())">
          <a:p>
            <a:endParaRPr lang="en-US" dirty="0" smtClean="0">
				<xsl:variable name ="VarSz">
					<xsl:call-template name ="getDefaultFontSize">
						<xsl:with-param name ="className" select ="$prClsName"/>
					</xsl:call-template >
				</xsl:variable>
				<xsl:if test ="$VarSz!=''">
					<xsl:attribute name ="sz">
						<xsl:value-of select ="substring-before($VarSz,'pt')*100"/>
					</xsl:attribute>
				</xsl:if>				
              <a:latin charset="0"  >
                <xsl:attribute name ="typeface">
                  <xsl:call-template name ="getDefaultFonaName">
                    <xsl:with-param name ="className" select ="$prClsName"/>
                    <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                  </xsl:call-template>
                </xsl:attribute>
              </a:latin >
            </a:endParaRPr>
          </a:p>
        </xsl:if>
        <!-- Change made by vijayeta, Bug fix	1731885-->
        </xsl:when>
        <xsl:when test ="name()='text:list'">
        <xsl:variable name="forCount" select="position()" />
        <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <!--Code inserted by Vijayets for Bullets and numbering-->
          <!--Check if Levels are prese nt-->
          <xsl:variable name ="lvl">
            <xsl:if test ="./text:list-item/text:list">
              <xsl:call-template name ="getListLevel">
                <xsl:with-param name ="levelCount"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="not(./text:list-item/text:list)">
              <xsl:value-of select ="'0'"/>
            </xsl:if>
          </xsl:variable >
          <xsl:variable name="paragraphId" >
            <xsl:call-template name ="getParaStyleNameForTextBox">
              <xsl:with-param name ="lvl" select ="$lvl"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:variable name ="isNumberingEnabled">
            <xsl:if test ="document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering">
              <xsl:value-of select ="document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering"/>
            </xsl:if>
            <xsl:if test ="not(document('content.xml')//style:style[@style:name=$paragraphId]/style:paragraph-properties/@text:enable-numbering)">
              <xsl:variable name ="styleNameFromStyles" >
                <xsl:choose >
                  <xsl:when test ="$prClsName='subtitle' or $prClsName='title'">
                    <xsl:value-of select ="concat($prClsName,'-',$masterPageName)"/>
                  </xsl:when>
                  <xsl:when test ="$prClsName='outline'">
                    <xsl:value-of select ="concat($masterPageName,'-outline',$lvl+1)"/>
                  </xsl:when>
                  <xsl:otherwise >
                    <xsl:value-of select ="concat($prClsName,'-',$masterPageName)"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>
              <xsl:if test ="document('styles.xml')//office:styles/style:style[@style:name=$styleNameFromStyles]/style:paragraph-properties/@text:enable-numbering">
                <xsl:value-of select ="document('styles.xml')//office:styles/style:style[@style:name=$styleNameFromStyles]/style:paragraph-properties/@text:enable-numbering"/>
              </xsl:if>
              <xsl:if test ="not(document('styles.xml')//office:styles/style:style[@style:name=$styleNameFromStyles]/style:paragraph-properties/@text:enable-numbering)">
                <xsl:value-of select ="'true'"/>
              </xsl:if>
            </xsl:if>
          </xsl:variable>
          <xsl:call-template name ="paraProperties" >
            <xsl:with-param name ="paraId" >
              <xsl:value-of select ="$paragraphId"/>
            </xsl:with-param >
            <!-- list property also included-->
            <xsl:with-param name ="listId"  select="@text:style-name"/>
            <!-- Parameters added by vijayeta,Set bulleting as true/false,and set level -->
            <xsl:with-param name ="isBulleted" select ="'true'"/>
            <xsl:with-param name ="level" select ="$lvl"/>
            <xsl:with-param name ="grpFlag" select ="$grpFlag"/>
            <xsl:with-param name ="isNumberingEnabled" select ="$isNumberingEnabled"/>
            <xsl:with-param name ="BuImgRel">
              <xsl:choose>
                <xsl:when test="$grpFlag='true'">
                  <xsl:value-of select ="concat($shapeCount,$forCount,$UniqueId)"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="concat($shapeCount,$forCount)"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:with-param> 
            <!-- Paremeter added by vijayeta,get master page name, dated:11-7-07-->
            <xsl:with-param name ="slideMaster" select ="$fileName"/>
            <xsl:with-param name ="pos" select ="$forCount"/>
            <xsl:with-param name ="shapeCount" select ="$shapeCount"/>
            <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
            <xsl:with-param name ="flagPresentationClass" select ="'No'"/>
            <xsl:with-param name ="parentStyleName" select ="$prClsName"/>
          </xsl:call-template >
          <!--End of Code inserted by Vijayets for Bullets and numbering-->
          <xsl:for-each select ="child::node()[position()]">
            <xsl:choose >
              <xsl:when test ="name()='text:list-item'">
                <xsl:variable name="textHyperlinksForBullets">
                  <xsl:call-template name="getHyperlinksForBTShape">
                    <xsl:with-param name="bulletLevel" select="$lvl" />
                    <xsl:with-param name="listItemCount" select="$forCount"/>
                    <xsl:with-param name="rootNode" select="."/>
                    <xsl:with-param name="shapeId" select="$shapeId"/>
                  </xsl:call-template>
                </xsl:variable>
                <xsl:variable name ="currentNodeStyle">
                  <xsl:call-template name ="getTextNodeForFontStyle">
                    <xsl:with-param name ="prClassName" select ="$prClsName"/>
                    <xsl:with-param name ="lvl" select ="$lvl"/>
                    <xsl:with-param name ="HyperlinksForBullets" select ="$textHyperlinksForBullets"/>
                    <!-- Paremeter added by vijayeta,get master page name, dated:11-7-07-->
                    <xsl:with-param name ="masterPageName" select ="$masterPageName"/>
                    <xsl:with-param name ="fileName" select ="$fileName"/>
                    <xsl:with-param name ="flagPresentationClass" select ="'No'"/>
                  </xsl:call-template>
                </xsl:variable>
                <xsl:copy-of select ="$currentNodeStyle"/>
              </xsl:when >
            </xsl:choose>
          </xsl:for-each>
          <!--<xsl:copy-of select="$varFrameHyperLinks"/>-->
        </a:p >
        </xsl:when>
      
      </xsl:choose>

    </xsl:for-each>
  </xsl:template>
  <!-- Blank lines in text-->
  <xsl:template name ="processBR">
    <xsl:param name ="T" />
    <a:br>
      <a:rPr lang="en-US" smtClean="0">
        <!--Font Size -->
        <xsl:if test ="not($T ='')">
          <xsl:call-template name ="fontStyles">
            <xsl:with-param name ="Tid" select ="$T" />
          </xsl:call-template>
        </xsl:if>
      </a:rPr >

    </a:br>
  </xsl:template>
  <!-- Get graphic properties from styles.xml -->
  <xsl:template name ="getDefaultStyle">
    <xsl:param name ="parentStyle" />
    <xsl:param name ="attributeName" />
    <xsl:for-each select ="document('styles.xml')/office:document-styles/office:styles/style:style[@style:name = $parentStyle]/style:graphic-properties">
      <xsl:choose>
        <xsl:when test ="$attributeName = 'fill'">
          <xsl:value-of select ="@draw:fill"/>
        </xsl:when>
        <xsl:when test ="$attributeName = 'padding-left'">
          <xsl:if test ="@fo:padding-left">
            <xsl:value-of select ="@fo:padding-left"/>
          </xsl:if>
          <xsl:if test ="not(@fo:padding-left)">
            <xsl:value-of select ="'0'"/>
          </xsl:if>
        </xsl:when>
        <xsl:when test ="$attributeName = 'padding-top'">
          <xsl:if test ="@fo:padding-top">
            <xsl:value-of select ="@fo:padding-top"/>
          </xsl:if>
          <xsl:if test ="not(@fo:padding-top)">
            <xsl:value-of select ="'0'"/>
          </xsl:if>
        </xsl:when>
        <xsl:when test ="$attributeName = 'padding-right'">
          <xsl:if test ="@fo:padding-right">
            <xsl:value-of select ="@fo:padding-right"/>
          </xsl:if>
          <xsl:if test ="not(@fo:padding-right)">
            <xsl:value-of select ="'0'"/>
          </xsl:if>
        </xsl:when>
        <xsl:when test ="$attributeName = 'padding-bottom'">
          <xsl:if test ="@fo:padding-bottom">
            <xsl:value-of select ="@fo:padding-bottom"/>
          </xsl:if>
          <xsl:if test ="not(@fo:padding-bottom)">
            <xsl:value-of select ="'0'"/>
          </xsl:if>
        </xsl:when>
        <xsl:when test ="$attributeName = 'wrap-option'">
          <xsl:if test ="@fo:wrap-option">
            <xsl:value-of select ="@fo:wrap-option"/>
          </xsl:if>
          <xsl:if test ="not(@fo:wrap-option)">
            <xsl:value-of select ="'square'"/>
          </xsl:if>
        </xsl:when>
        <xsl:when test ="$attributeName = 'auto-grow-height'">
          <xsl:value-of select ="@draw:auto-grow-height"/>
        </xsl:when>
        <xsl:when test ="$attributeName = 'auto-grow-width'">
          <xsl:value-of select ="@draw:auto-grow-width"/>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <!-- Vijayeta,for bullets and numbering,get para style name depending on levels-->
  <xsl:template name ="getParaStyleNameForTextBox">
    <xsl:param name ="lvl"/>
    <xsl:choose>
      <xsl:when test ="$lvl='0'">
        <xsl:value-of select ="./text:list-item/text:p/@text:style-name"/>
      </xsl:when >
      <xsl:when test ="$lvl='1'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='2'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='3'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='4'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='5'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='6'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='7'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='8'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
      <xsl:when test ="$lvl='9'">
        <xsl:value-of select ="./text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p/@text:style-name"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!-- Inner levels,styles-->
  
  <!-- End of code by Vijayeta,for bullets and numbering,get para style name depending on levels-->

  <xsl:template name="getHyperlinksForBTShape">
    <xsl:param name ="bulletLevel"/>
    <xsl:param name ="listItemCount" />
    <xsl:param name ="rootNode" />
    <xsl:param name ="shapeId" />
    <xsl:choose>
      <xsl:when test ="./text:p">
        <xsl:for-each select="./text:p">
          <xsl:if test="text:a">
            <a:hlinkClick>
              <xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
          <xsl:if test="text:span/text:a">
            <a:hlinkClick>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:p">
        <xsl:for-each select="./text:list/text:list-item/text:p">
          <xsl:if test="text:a">
            <a:hlinkClick>
              <xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
          <xsl:if test="text:span/text:a">
            <a:hlinkClick>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:if test="text:a">
            <a:hlinkClick>
              <xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
          <xsl:if test="text:span/text:a">
            <a:hlinkClick>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:if test="text:a">
            <a:hlinkClick>
              <xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
          <xsl:if test="text:span/text:a">
            <a:hlinkClick>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:if test="text:a">
            <a:hlinkClick>
              <xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
          <xsl:if test="text:span/text:a">
            <a:hlinkClick>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:if test="text:a">
            <a:hlinkClick>
              <xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
          <xsl:if test="text:span/text:a">
            <a:hlinkClick>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:if test="text:a">
            <a:hlinkClick>
              <xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
          <xsl:if test="text:span/text:a">
            <a:hlinkClick>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:if test="text:a">
            <a:hlinkClick>
              <xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
          <xsl:if test="text:span/text:a">
            <a:hlinkClick>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test ="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
        <xsl:for-each select="./text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:list/text:list-item/text:p">
          <xsl:if test="text:a">
            <a:hlinkClick>
              <xsl:if test="text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
          <xsl:if test="text:span/text:a">
            <a:hlinkClick>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'#Slide')]">
                <xsl:attribute name="action">
                  <xsl:value-of select="'ppaction://hlinksldjump'"/>
                </xsl:attribute>
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(text:span/text:a/@xlink:href[ contains(.,'http://') or contains(.,'mailto:')]) and text:span/text:a/@xlink:href[ contains(.,':') or contains(.,'../')]">
                <xsl:attribute name="r:id">
                  <xsl:value-of select="concat($shapeId,'BLVL',$bulletLevel,'Link',$listItemCount)"/>
                </xsl:attribute>
              </xsl:if>
            </a:hlinkClick>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet >
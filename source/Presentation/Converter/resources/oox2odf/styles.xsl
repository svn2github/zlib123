<?xml version="1.0" encoding="UTF-8" ?>
<!--
   Pradeep Nemadi
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
				xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
				xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
				xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:v="urn:schemas-microsoft-com:vml" exclude-result-prefixes="w r draw number wp xlink">
  
  
  <xsl:key name="StyleId" match="w:style" use="@w:styleId"/>
  <xsl:key name="default-styles"
    match="w:style[@w:default = 1 or @w:default = 'true' or @w:default = 'on']" use="@w:type"/>

  <xsl:template name="styles">
    <office:document-styles>
      <office:font-face-decls>
        <xsl:apply-templates select="document('ppt/fontTable.xml')/w:fonts"/>
      </office:font-face-decls>
      <!-- document styles -->
      <office:styles>
        <xsl:call-template name="InsertDefaultStyles"/>
      </office:styles>
      <!-- automatic styles -->
      <office:automatic-styles>
		  <!--  <xsl:call-template name="InsertSlideFormat"/>-->    
		 <xsl:call-template name="InsertSlideSize"/>
		 <xsl:call-template name="InsertNotesSize"/>
	  </office:automatic-styles>
      <!-- master styles -->
      <office:master-styles>
		  <xsl:call-template name="InsertMasterStylesDefinition"/>
	  </office:master-styles>
    </office:document-styles>
  </xsl:template>
  <xsl:template name="InsertDefaultStyles">
    <xsl:for-each select="document('ppt/styles.xml')">
    </xsl:for-each>
  </xsl:template>
  <!-- Changes made by Vijayeta-->	
	<!-- Slide Size-->
  <xsl:template name="InsertSlideSize">
    <xsl:for-each select ="document('ppt/presentation.xml')//p:presentation/p:sldSz">		
      <style:page-layout style:name="PM1">
        <style:page-layout-properties >
          <xsl:attribute name ="fo:page-width">
            <xsl:call-template name="ConvertEmu">
              <xsl:with-param name="length" select="@cx"/>
              <xsl:with-param name="unit">cm</xsl:with-param>
            </xsl:call-template>
          </xsl:attribute>
		  <xsl:attribute name ="fo:page-height">
				<xsl:call-template name="ConvertEmu">
					<xsl:with-param name="length" select="@cy"/>
					<xsl:with-param name="unit">cm</xsl:with-param>
				</xsl:call-template>
			</xsl:attribute>
          <xsl:attribute name ="style:print-orientation">            
            <xsl:call-template name="CheckOrientation">
              <xsl:with-param name="cx" select="@cx"/>
              <xsl:with-param name="cy" select="@cy"/>
            </xsl:call-template>
          </xsl:attribute>
        </style:page-layout-properties>
      </style:page-layout>
    </xsl:for-each>
  </xsl:template>
  <!-- Notes Size-->  
  <xsl:template name="InsertNotesSize">
    <!-- Check if notesSlide is present in the package-->
    <xsl:if test ="document('ppt/notesSlides/notesSlide1.xml')">
      <xsl:variable name ="Flag">
        <!--Check if size defined in notesSlide -->
        <xsl:for-each select ="document('ppt/notesSlides/notesSlide1.xml')//p:notes/p:cSld/p:spTree/p:sp">
          <xsl:if test ="p:nvSpPr/p:cNvPr/@name[contains(.,'Notes Placeholder 2')]">
            <xsl:if test ="p:spPr/a:xfrm/a:ext[@cx]">
              <xsl:value-of select ="'true'"/>
            </xsl:if>
            <xsl:if test ="not(p:spPr/a:xfrm/a:ext[@cx])">
              <xsl:value-of select ="'false'"/>
            </xsl:if>
          </xsl:if>
        </xsl:for-each>
      </xsl:variable>
      <xsl:choose >
        <xsl:when test ="$Flag='true'">
        <!-- notesSlide has Size Definition(user defined)-->
          <xsl:for-each select ="document('ppt/notesSlides/notesSlide1.xml')//p:notes/p:cSld/p:spTree/p:sp">
            <xsl:if test ="p:nvSpPr/p:cNvPr/@name[contains(.,'Notes Placeholder 2')]">
              <style:page-layout style:name="PM2">
                <style:page-layout-properties>
                  <xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cx]">
                    <xsl:attribute name ="fo:page-width">
                      <xsl:call-template name="ConvertEmu">
                        <xsl:with-param name="length" select="@cx"/>
                        <xsl:with-param name="unit">cm</xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:for-each>
                  <xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cy]">
                    <xsl:attribute name ="fo:page-height">
                      <xsl:call-template name="ConvertEmu">
                        <xsl:with-param name="length" select="@cy"/>
                        <xsl:with-param name="unit">cm</xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:for-each>
                  <xsl:attribute name ="style:print-orientation">
                    <xsl:call-template name="CheckOrientation">
                      <xsl:with-param name="cx" select="@cx"/>
                      <xsl:with-param name="cy" select="@cy"/>
                    </xsl:call-template>
                  </xsl:attribute>
                </style:page-layout-properties>
              </style:page-layout>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <!-- pre-defined size definition in notesMaster-->
          <xsl:for-each select ="document('ppt/notesMasters/notesMaster1.xml')//p:notesMaster/p:cSld/p:spTree/p:sp">
            <xsl:if test ="p:nvSpPr/p:cNvPr/@name[contains(.,'Notes Placeholder 4')]">
              <style:page-layout style:name="PM2">
                <style:page-layout-properties>
                  <xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cx]">
                    <xsl:attribute name ="fo:page-width">
                      <xsl:call-template name="ConvertEmu">
                        <xsl:with-param name="length" select="@cx"/>
                        <xsl:with-param name="unit">cm</xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:for-each>
                  <xsl:for-each select ="p:spPr/a:xfrm/a:ext[@cy]">
                    <xsl:attribute name ="fo:page-height">
                      <xsl:call-template name="ConvertEmu">
                        <xsl:with-param name="length" select="@cy"/>
                        <xsl:with-param name="unit">cm</xsl:with-param>
                      </xsl:call-template>
                    </xsl:attribute>
                  </xsl:for-each>
                  <xsl:attribute name ="style:print-orientation">
                    <xsl:call-template name="CheckOrientation">
                      <xsl:with-param name="cx" select="@cx"/>
                      <xsl:with-param name="cy" select="@cy"/>
                    </xsl:call-template>
                  </xsl:attribute>
                </style:page-layout-properties>
              </style:page-layout>
            </xsl:if>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>     
    </xsl:if>
    <!--Default Size in presentation.xml-->
    <xsl:if test ="not(document('ppt/notesSlides/notesSlide1.xml'))">
      <xsl:for-each select ="document('ppt/presentation.xml')//p:presentation/p:notesSz">
        <style:page-layout style:name="PM2">
          <style:page-layout-properties>
            <xsl:attribute name ="fo:page-width">
              <xsl:call-template name="ConvertEmu">
                <xsl:with-param name="length" select="@cx"/>
                <xsl:with-param name="unit">cm</xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
            <xsl:attribute name ="fo:page-height">
              <xsl:call-template name="ConvertEmu">
                <xsl:with-param name="length" select="@cy"/>
                <xsl:with-param name="unit">cm</xsl:with-param>
              </xsl:call-template>
            </xsl:attribute>
            <xsl:attribute name ="style:print-orientation">
              <xsl:call-template name="CheckOrientation">
                <xsl:with-param name="cx" select="@cx"/>
                <xsl:with-param name="cy" select="@cy"/>
              </xsl:call-template>
            </xsl:attribute>
          </style:page-layout-properties>
        </style:page-layout>
      </xsl:for-each>
    </xsl:if>          
  </xsl:template>  
	<!-- Slide Number,The equvivalent Not present in ODP-->
	<xsl:template name="InsertSlideNumber">
		<xsl:for-each select ="document('ppt/presentation.xml')//p:presentation">
			<draw:frame presentation:class="page-number">
				<draw:text-box>
					<text:p>
						<xsl:if test ="@firstSlideNum">
							<text:page-number>
								<xsl:value-of select ="@firstSlideNum"/>
							</text:page-number>						
					    </xsl:if>
					</text:p>
				</draw:text-box>
			</draw:frame>
		</xsl:for-each>
	</xsl:template>   
	<!-- Add MasterStyles Definition-->
	<xsl:template name="InsertMasterStylesDefinition">
    <style:handout-master style:page-layout-name="PM0"/>    
    <style:master-page style:name="Default" style:page-layout-name="PM1" >
			<presentation:notes style:page-layout-name="PM2">			
			<xsl:call-template name="InsertSlideNumber" />
      </presentation:notes>
		</style:master-page>		
	</xsl:template>
  <!--checks cx/cy ratio,for orientation-->
  <xsl:template name ="CheckOrientation">
    <xsl:param name="cx"/>
    <xsl:param name="cy"/>
    <xsl:variable name="orientation"/>
    <xsl:choose>
      <xsl:when test ="$cx > $cy">        
         <xsl:value-of select="'landscape'" />        
      </xsl:when>
      <xsl:otherwise >        
        <xsl:value-of select ="'portrait'"/>       
      </xsl:otherwise>
    </xsl:choose>
    <xsl:value-of select="$orientation"/>
  </xsl:template>
  <!--  converts emu to given unit-->
  <xsl:template name="ConvertEmu">
    <xsl:param name="length"/>
    <xsl:param name="unit"/>
    <xsl:choose>
      <xsl:when
			  test="$length = '' or not($length) or $length = 0 or format-number($length div 360000, '#.##') = ''">
        <xsl:value-of select="concat(0,'cm')"/>
      </xsl:when>
      <xsl:when test="$unit = 'cm'">
        <xsl:value-of select="concat(format-number($length div 360000, '#.##'), 'cm')"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!-- Changes made by Vijayeta-->
</xsl:stylesheet>
 

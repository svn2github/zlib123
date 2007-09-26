<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
	xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
	xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
	xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
	xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" 
	xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
	xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
	xmlns:xlink="http://www.w3.org/1999/xlink"
	xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
	xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" 
	xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
	xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
	xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
	xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
	xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
	exclude-result-prefixes="xlink number">
 
  <xsl:template name ="presentation">
    <p:presentation>
      <p:sldMasterIdLst>
        
        <xsl:for-each select ="document('styles.xml')//office:master-styles/style:master-page ">
          <p:sldMasterId >
            <xsl:attribute name ="id">
              <xsl:value-of select="2147483648 + (12 * (position() - 1 ))" />
            </xsl:attribute>
            <xsl:attribute name ="r:id">
              <xsl:value-of select ="concat('smId',position())"/>
            </xsl:attribute>
          </p:sldMasterId >
        </xsl:for-each>
        
      </p:sldMasterIdLst>
      <p:notesMasterIdLst>
        <p:notesMasterId r:id="noteMasterRel"/>
      </p:notesMasterIdLst>
<!-- Inserted by vijayeta
           Add handoutMasterIdLst in presentation.xml 
           Date: 30th July '07-->
      <xsl:for-each select ="document('styles.xml')//office:master-styles/style:handout-master">
        <p:handoutMasterIdLst>
          <p:handoutMasterId >
            <xsl:attribute name ="r:id">
              <xsl:value-of select ="concat('hoId',position())"/>
            </xsl:attribute>
          </p:handoutMasterId>
        </p:handoutMasterIdLst>
      </xsl:for-each>
      <!-- End of Snippet Inserted by vijayeta
           Add handoutMasterIdLst in presentation.xml 
           Date: 30th July '07-->
      <p:sldIdLst>
        <xsl:for-each select ="document('content.xml')
					/office:document-content/office:body/office:presentation/draw:page">
          <p:sldId>
            <xsl:attribute name ="id">
              <xsl:value-of select ="255 +position()"/>
            </xsl:attribute>
            <xsl:attribute name ="r:id">
              <xsl:value-of select ="concat('sId',position())"/>
            </xsl:attribute>
          </p:sldId>
        </xsl:for-each>
      </p:sldIdLst>
      <p:sldSz >
        <!-- Page width and height-->
        <xsl:for-each select="document('styles.xml')//office:master-styles/style:master-page">
          <xsl:if test="position()=1">
            <xsl:variable name="PMName" select="@style:page-layout-name"/>
            <xsl:for-each select ="document('styles.xml')//style:page-layout[@style:name=$PMName]">
          <xsl:attribute name ="cx" >
            <xsl:if test="style:page-layout-properties/@fo:page-width">
              <xsl:call-template name ="convertToPoints">
                <xsl:with-param name ="unit" select ="'cm'"/>
                <xsl:with-param name ="length" select ="style:page-layout-properties/@fo:page-width"/>
              </xsl:call-template>
            </xsl:if>
            <!-- Modified by lohith - Always cx should be greater than 0-->
            <xsl:if test="not(style:page-layout-properties/@fo:page-width)">
              <xsl:value-of select ='10080000'/>
            </xsl:if>
          </xsl:attribute>
          <xsl:attribute name ="cy" >
            <xsl:if test="style:page-layout-properties/@fo:page-height">
              <xsl:call-template name ="convertToPoints">
                <xsl:with-param name ="unit" select ="'cm'"/>
                <xsl:with-param name ="length" select ="style:page-layout-properties/@fo:page-height"/>
              </xsl:call-template>
            </xsl:if>
            <!-- Modified by lohith - Always cy should be greater than 0-->
            <xsl:if test="not(style:page-layout-properties/@fo:page-height)">
              <xsl:value-of select ='7560000'/>
            </xsl:if>
          </xsl:attribute>
        </xsl:for-each >
          </xsl:if>
        </xsl:for-each>
      </p:sldSz >
      <!--NotesMaster-->
         <xsl:for-each select="document('styles.xml')//office:master-styles/style:master-page">
        <xsl:if test="position()=1">
          <xsl:variable name="PMName" select="./presentation:notes/@style:page-layout-name"/>
          <p:notesSz>
            <xsl:if test ="document('styles.xml')//style:page-layout/@style:name = $PMName">
              <xsl:for-each select ="document('styles.xml')//style:page-layout[@style:name=$PMName]">
                <xsl:attribute name ="cx" >
                  <xsl:if test="style:page-layout-properties/@fo:page-width">
                    <xsl:call-template name ="convertToPoints">
                      <xsl:with-param name ="unit" select ="'cm'"/>
                      <xsl:with-param name ="length" select ="style:page-layout-properties/@fo:page-width"/>
                    </xsl:call-template>
                  </xsl:if>
                  <!-- Modified by lohith - Always cx should be greater than 0-->
                  <xsl:if test="not(style:page-layout-properties/@fo:page-width)">
                    <xsl:value-of select ='7772400'/>
                  </xsl:if>
                </xsl:attribute>
                <xsl:attribute name ="cy" >
                  <xsl:if test="style:page-layout-properties/@fo:page-height">
                    <xsl:call-template name ="convertToPoints">
                      <xsl:with-param name ="unit" select ="'cm'"/>
                      <xsl:with-param name ="length" select ="style:page-layout-properties/@fo:page-height"/>
                    </xsl:call-template>
                  </xsl:if>
                  <!-- Modified by lohith - Always cy should be greater than 0-->
                  <xsl:if test="not(style:page-layout-properties/@fo:page-height)">
                    <xsl:value-of select ='10058400'/>
                  </xsl:if>
                </xsl:attribute>
              </xsl:for-each>
            </xsl:if>
            <xsl:if test ="not(document('styles.xml')//style:page-layout/@style:name = $PMName)">
              <xsl:attribute name ="cx" >
                <xsl:value-of select ='7772400'/>
              </xsl:attribute>
              <xsl:attribute name ="cy" >
                <xsl:value-of select ='10058400'/>
              </xsl:attribute>
            </xsl:if>
          </p:notesSz>
        </xsl:if>
      </xsl:for-each>
      <!--END-->
      
      <!-- Changes made by vijayeta-->
      <!-- Custome Slide show - added by lohith.ar-->
      <xsl:call-template name="TemplateCustomSlideShow"/>
      <!-- Default Style to pptx-->
      <p:defaultTextStyle>
        <a:defPPr>
          <a:defRPr lang="en-US"/>
        </a:defPPr>
        <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1800" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl1pPr>
        <a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1800" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl2pPr>
        <a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1800" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl3pPr>
        <a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1800" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl4pPr>
        <a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1800" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl5pPr>
        <a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1800" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl6pPr>
        <a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1800" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl7pPr>
        <a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1800" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl8pPr>
        <a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1800" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl9pPr>
      </p:defaultTextStyle>
    </p:presentation>
  </xsl:template>
  <xsl:template name ="presentationRel">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <!--<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>-->
      <!--<Relationship Id="sId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>-->
      <!--<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>-->
      <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>
      <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
      <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>
      <!--<Relationship Id="sId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>-->
      <xsl:for-each select ="document('content.xml')
					/office:document-content/office:body/office:presentation/draw:page">
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide">
          <xsl:attribute name ="Id">
            <xsl:value-of select ="concat('sId',position())"/>
          </xsl:attribute>
          <xsl:attribute name ="Target">
            <xsl:value-of select ="concat(concat('slides/slide',position()),'.xml')"/>
          </xsl:attribute>
        </Relationship >
      </xsl:for-each >     
      <!-- @@Slide master code begins Pradeep Nemadi-->
      
      
      <xsl:for-each select ="document('styles.xml')//office:master-styles/style:master-page ">
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster">
          <xsl:attribute name ="Id">
            <xsl:value-of select ="concat('smId',position())"/>
          </xsl:attribute>
          <xsl:attribute name ="Target">
            <xsl:value-of select ="concat('slideMasters/slideMaster',position(),'.xml')"/>
          </xsl:attribute>
        </Relationship>
      </xsl:for-each>

      <Relationship Id="noteMasterRel" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"
                    Target="notesMasters/notesMaster1.xml"/>
      <!-- @@Slide master code ends Pradeep Nemadi-->		 

      <!-- Inserted by vijayeta
       Add handoutmaster relatuionship in presentation.xml.rels 
       Date: 30th July '07-->
      <xsl:for-each select ="document('styles.xml')//office:master-styles/style:handout-master">
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster">
          <xsl:attribute name ="Id">
            <xsl:value-of select ="concat('hoId',position())"/>
          </xsl:attribute>
          <xsl:attribute name ="Target">
            <xsl:value-of select ="concat('handoutMasters/handoutMaster',position(),'.xml')"/>
          </xsl:attribute>
        </Relationship>
      </xsl:for-each>
      <!-- End of Snippet Inserted by vijayeta
       Add handoutmaster relatuionship in presentation.xml.rels 
       Date: 30th July '07-->
      <!-- Inserted by vijayeta
       Add PresProps relationship in presentation.xml.rels 
       Date: 4th Aug '07-->
      <xsl:for-each select ="document('content.xml')//office:body/office:presentation/presentation:settings">
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps">
          <xsl:attribute name ="Id">
            <xsl:value-of select ="concat('presProps',position())"/>
          </xsl:attribute>
          <xsl:attribute name ="Target">
            <xsl:value-of select ="'presProps.xml'"/>
          </xsl:attribute>
        </Relationship>
      </xsl:for-each>
      <!--End of Snippet to add PresProps relationship in presentation.xml.rels -->
	</Relationships>
  </xsl:template >

  <!-- Template for Custome Slide show = Added by lohith.ar-->
  <xsl:template name="TemplateCustomSlideShow">
    <p:custShowLst>
      <xsl:for-each select ="document('content.xml')
					/office:document-content/office:body/office:presentation/presentation:settings/presentation:show">
        <p:custShow>
          <xsl:attribute name ="name">
            <xsl:value-of select ="@presentation:name"/>
          </xsl:attribute>
          <xsl:attribute name ="id">
            <xsl:value-of select ="position()"/>
          </xsl:attribute>
          <p:sldLst>
            <xsl:call-template name="GetSlideList">
              <xsl:with-param name="pages" select="@presentation:pages"/>
            </xsl:call-template>
          </p:sldLst>
        </p:custShow>
      </xsl:for-each>
    </p:custShowLst>
  </xsl:template>
  <xsl:template name="GetSlideList">
    <xsl:param name="pages"/>
    <xsl:variable name="varshowList" select="$pages"/>
    <xsl:variable name="first" select='substring-before($pages,",")'/>
    <xsl:variable name='rest' select='substring-after($pages,",")'/>
    <xsl:if test='$first'>
      <xsl:for-each select="document('content.xml')/office:document-content/office:body/office:presentation/draw:page">
        <xsl:if test="@draw:name = $first">
          <p:sld>
            <xsl:attribute name ="r:id">
              <xsl:value-of select ="concat('sId',position())"/>
            </xsl:attribute>
          </p:sld>
        </xsl:if>
      </xsl:for-each>
    </xsl:if>
    <xsl:if test='$rest'>
      <xsl:call-template name='GetSlideList'>
        <xsl:with-param name='pages' select='$rest'/>
      </xsl:call-template>
    </xsl:if>
    <xsl:if test='not($rest)'>
      <xsl:for-each select="document('content.xml')/office:document-content/office:body/office:presentation/draw:page">
        <xsl:if test="@draw:name = $pages">
          <p:sld>
            <xsl:attribute name ="r:id">
              <xsl:value-of select ="concat('sId',position())"/>
            </xsl:attribute>
          </p:sld>
        </xsl:if>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>
</xsl:stylesheet>